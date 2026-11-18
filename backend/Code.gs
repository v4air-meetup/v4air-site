/* V4AIR forms backend: Google Sheet + Apps Script web app.
   Bound to the "V4AIR forms" spreadsheet (created by clasp create --type sheets).
   Tabs: Applications (form A), Nominations and Nominees (form B), Speakers (form S).
   Spec: forms/survey-spec.md sections 5.2 (save and resume), 5.11 (export), 5.17 (IDs). */

var ID_ALPHABET = '23456789ABCDEFGHJKLMNPQRSTUVWXYZ';
var META = ['response_id', 'token_hash', 'status', 'started_at', 'last_saved_at', 'submitted_at', 'screen', 'answers_json'];
var NOMINEE_COLS = ['nomination_id', 'application_id', 'nominator_name', 'nominator_email', 'n', 'type', 'name', 'contact', 'affiliation', 'why', 'saved_at'];
var FORMS = {
  A: { tab: 'Applications', prefix: 'V4A', keyRe: /^A[1-8]\.[A-Za-z0-9_.]+$/ },
  B: { tab: 'Nominations', prefix: 'V4N', keyRe: /^B\.[A-Za-z0-9_.]+$/ },
  /* invited speakers (public/invited-speaker/): shared questions keep the A variable IDs, speaker-only ones are S.* */
  S: { tab: 'Speakers', prefix: 'V4S', keyRe: /^(A[1-8]|S)\.[A-Za-z0-9_.]+$/ }
};
var MAX_BODY = 15000000, MAX_VALUE = 12000, MAX_JSON = 49000, MAX_KEYS = 800, MAX_CARDS = 9; /* a Sheets cell holds 50,000 characters */
var EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
/* A7 upload (form A) and speaker photo (form S): one file per record, kept under its original name in "V4AIR uploads/<application id>/" */
var UPLOAD_ROOT = 'V4AIR uploads', MAX_FILE = 10 * 1024 * 1024;
var FILE_TYPES = { 'application/pdf': 'pdf', 'application/vnd.openxmlformats-officedocument.presentationml.presentation': 'pptx', 'image/png': 'png', 'image/jpeg': 'jpg' };

/* ---------- entry points ---------- */

function doPost(e) {
  try {
    var body = (e && e.postData && e.postData.contents) || '';
    if (body.length > MAX_BODY) return reply({ ok: false, error: 'too_large' });
    var req = JSON.parse(body);
    if (req.hp) return reply({ ok: true, id: 'V4X-0000-0000', token: '' }); /* honeypot: pretend success, store nothing */
    if (req.action === 'load') return reply(load(req));
    if (req.action === 'save') return reply(save(req));
    return reply({ ok: false, error: 'bad_action' });
  } catch (err) {
    var msg = String(err && err.message || err);
    if (/^file_(type|size)$/.test(msg)) return reply({ ok: false, error: msg });
    console.error(err && err.stack || err);
    return reply({ ok: false, error: 'server' });
  }
}

function doGet() {
  return reply({ ok: true, service: 'v4air-forms' });
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* Run once from the editor as petkout1: creates the tabs and triggers the OAuth consent. */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ['A', 'B'].forEach(function (f) { sheetFor(f); });
  var nom = ss.getSheetByName('Nominees') || ss.insertSheet('Nominees');
  if (nom.getLastColumn() === 0) { nom.getRange(1, 1, 1, NOMINEE_COLS.length).setValues([NOMINEE_COLS]); nom.setFrozenRows(1); }
  var first = ss.getSheets()[0];
  if (first.getName() === 'Sheet1' && first.getLastRow() === 0) ss.deleteSheet(first);
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('EMAIL_ENABLED')) props.setProperty('EMAIL_ENABLED', 'false');
  if (!props.getProperty('SITE_URL')) props.setProperty('SITE_URL', '');
  MailApp.getRemainingDailyQuota(); /* touches the mail scope so consent covers it */
  UrlFetchApp.getRequest('https://api.resend.com/emails'); /* same for external requests (Resend); sends nothing */
  if (!props.getProperty('RESEND_API_KEY') && GmailApp.getAliases().indexOf(MAIL_FROM) === -1) throw new Error('Gmail cannot send as ' + MAIL_FROM);
  uploadRoot(); /* same for Drive */
  DocumentApp.openById(props.getProperty('EMAILS_DOC_ID') || EMAILS_DOC).getName(); /* same for reading the email texts */
  if (!ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'sendMail'; })) ScriptApp.newTrigger('sendMail').timeBased().everyMinutes(5).create();
}

/* ---------- actions ---------- */

function load(req) {
  var f = FORMS[req.form];
  if (!f || !req.id || !req.token) return { ok: false, error: 'bad_request' };
  var sh = sheetFor(req.form);
  var row = findRow(sh, req.id);
  if (!row) return { ok: false, error: 'not_found' };
  var head = header(sh);
  var vals = sh.getRange(row, 1, 1, head.length).getValues()[0];
  if (!tokenOk(vals[head.indexOf('token_hash')], req.token)) return { ok: false, error: 'not_found' };
  return { ok: true, id: req.id, status: vals[head.indexOf('status')], answers: JSON.parse(vals[head.indexOf('answers_json')] || '{}') };
}

function save(req) {
  var f = FORMS[req.form];
  if (!f || !req.answers || typeof req.answers !== 'object') return { ok: false, error: 'bad_request' };
  var answers = clean(req.answers, f.keyRe);
  if (JSON.stringify(answers.json).length > MAX_JSON) return { ok: false, error: 'too_large' };
  var email = String(req.answers[req.form === 'B' ? 'B.email' : 'A1.email'] || '').trim();
  if (!EMAIL_RE.test(email)) return { ok: false, error: 'email' };

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  var created = false, id = req.id, token = req.token, now = new Date();
  try {
    /* one header read and one row read per save; every Sheets call costs 0.2 to 1 s inside Apps Script */
    var sh = sheetFor(req.form);
    var head = header(sh);
    var row = id ? findRow(sh, id) : 0, rec = {};
    if (row) {
      var cur = sh.getRange(row, 1, 1, head.length).getValues()[0];
      head.forEach(function (h, i) { rec[h] = cur[i]; });
      if (!tokenOk(rec.token_hash, token)) return { ok: false, error: 'not_found' };
    } else {
      /* no record yet, or a stale id the sheet doesn't know: start a new record.
         32^8 ids, so no lookup for collisions */
      id = newId(f.prefix);
      token = newToken();
      row = sh.getLastRow() + 1;
      created = true;
      rec.response_id = id; rec.token_hash = hash(token); rec.started_at = now; rec.status = 'in progress';
    }
    rec.last_saved_at = now;
    rec.screen = typeof req.answers.__screen === 'number' ? req.answers.__screen : '';
    /* a nomination sent again: nominees who declined or unsubscribed are wiped before anything is written, and the others
       keep what the sheet already knows about them (decision, mailer columns, saved_at); the Worker does the same (nomineePlan) */
    var plan = req.form === 'B' && (req.final || rec.status === 'submitted') ? nomineePlan(created ? '' : id, answers.json.__cards || []) : null;
    rec.answers_json = JSON.stringify(answers.json);
    /* GA4 client id (the _ga cookie value); joins a row to its traffic source in GA */
    if (typeof req.ga_cid === 'string' && /^[0-9.]{1,40}$/.test(req.ga_cid)) rec.ga_client_id = req.ga_cid;
    if (req.final) { rec.status = 'submitted'; if (!rec.submitted_at) rec.submitted_at = now; }
    Object.keys(answers.flat).forEach(function (k) { rec[k] = answers.flat[k]; });
    if ((req.form === 'A' || req.form === 'S') && (req.file || req.removeFile)) {
      var stored = storeFile(id, req.file), col = req.form === 'S' ? 'S.photo' : 'A7.file';
      answers.json.__file = stored;
      rec.answers_json = JSON.stringify(answers.json);
      rec[col + '_name'] = stored ? stored.name : '';
      rec[col + '_url'] = stored ? stored.url : '';
    }
    head = addColumns(sh, head, Object.keys(rec));
    /* the whole row is rewritten, so text read from the sheet needs the same protection as new answers */
    sh.getRange(row, 1, 1, head.length).setValues([head.map(function (h) { return !(h in rec) ? '' : typeof rec[h] === 'string' ? cellText(rec[h]) : rec[h]; })]);
    if (req.form === 'B' && req.final) writeNominees(id, answers.json, now, plan);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  var out = { ok: true, id: id, token: token, created: created };
  if ('__file' in answers.json) out.file = answers.json.__file;
  return out;
}

/* ---------- uploads ---------- */

function uploadRoot() {
  var it = DriveApp.getFoldersByName(UPLOAD_ROOT);
  return it.hasNext() ? it.next() : DriveApp.createFolder(UPLOAD_ROOT);
}

/* Replaces whatever is in the applicant's folder with the new file; null file = remove only. Returns {name, size, url} or null. */
function storeFile(id, file) {
  var root = uploadRoot();
  var it = root.getFoldersByName(id);
  var folder = it.hasNext() ? it.next() : null;
  if (folder) { var old = folder.getFiles(); while (old.hasNext()) old.next().setTrashed(true); }
  if (!file) return null;
  if (!file.data || typeof file.data !== 'string' || !FILE_TYPES[file.type]) throw new Error('file_type');
  var bytes = Utilities.base64Decode(file.data);
  if (bytes.length > MAX_FILE) throw new Error('file_size');
  var name = String(file.name || 'upload').replace(/[\\\/:*?"<>|\u0000-\u001f]/g, '_').slice(0, 200);
  if (!new RegExp('\\.' + FILE_TYPES[file.type] + '$', 'i').test(name) && !(file.type === 'image/jpeg' && /\.jpeg$/i.test(name))) name += '.' + FILE_TYPES[file.type];
  if (!folder) folder = root.createFolder(id);
  var f = folder.createFile(Utilities.newBlob(bytes, file.type, name));
  return { name: f.getName(), size: bytes.length, url: f.getUrl() };
}

/* ---------- sheet helpers ---------- */

function sheetFor(form) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = FORMS[form].tab;
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, META.length).setValues([META]);
    sh.setFrozenRows(1);
    sh.setFrozenColumns(1);
  }
  return sh;
}

function header(sh) {
  var n = sh.getLastColumn();
  return n ? sh.getRange(1, 1, 1, n).getValues()[0].map(String) : [];
}

/* Appends a column for every answer key not yet in the header, in the order the form sends them. */
function ensureColumns(sh, keys) {
  var head = header(sh);
  var missing = keys.filter(function (k) { return head.indexOf(k) === -1; });
  if (missing.length) {
    if (sh.getMaxColumns() < head.length + missing.length) sh.insertColumnsAfter(sh.getMaxColumns(), head.length + missing.length - sh.getMaxColumns());
    sh.getRange(1, head.length + 1, 1, missing.length).setValues([missing]);
    head = head.concat(missing);
  }
  return head;
}

/* Same as ensureColumns, with the header already in hand (saves a read). */
function addColumns(sh, head, keys) {
  var missing = keys.filter(function (k) { return head.indexOf(k) === -1; });
  if (!missing.length) return head;
  if (sh.getMaxColumns() < head.length + missing.length) sh.insertColumnsAfter(sh.getMaxColumns(), head.length + missing.length - sh.getMaxColumns());
  sh.getRange(1, head.length + 1, 1, missing.length).setValues([missing]);
  return head.concat(missing);
}

function newId(prefix) {
  var out = '';
  for (var i = 0; i < 8; i++) { out += ID_ALPHABET.charAt(Math.floor(Math.random() * ID_ALPHABET.length)); if (i === 3) out += '-'; }
  return prefix + '-' + out;
}

function findRow(sh, id) {
  if (!/^V4[ANCS]-[2-9A-HJ-NP-Z]{4}-[2-9A-HJ-NP-Z]{4}$/.test(String(id))) return 0;
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var hit = sh.getRange(2, 1, last - 1, 1).createTextFinder(id).matchEntireCell(true).findNext();
  return hit ? hit.getRow() : 0;
}

function rowObject(sh, row, head) {
  var vals = sh.getRange(row, 1, 1, head.length).getValues()[0], o = {};
  head.forEach(function (h, i) { o[h] = vals[i]; });
  return o;
}

function uniqueId(sh, prefix) {
  for (var tries = 0; tries < 20; tries++) {
    var out = '';
    for (var i = 0; i < 8; i++) { out += ID_ALPHABET.charAt(Math.floor(Math.random() * ID_ALPHABET.length)); if (i === 3) out += '-'; }
    var id = prefix + '-' + out;
    if (!findRow(sh, id)) return id;
  }
  throw new Error('could not generate a unique id');
}

/* Before a nomination is (re)written: for each card, the nominee's old Nominees row (same nomination, same address) and
   whether they declined (address in Unsubscribed). Declined cards are wiped in place, so their details reach neither
   answers_json nor Nominees again. Same as the Worker's nomineePlan (backend-worker/src/index.js). */
var NOM_DATA = ['nomination_id', 'application_id', 'nominator_name', 'nominator_email', 'n', 'type', 'name', 'contact', 'affiliation', 'why'];
function nomineePlan(nominationId, cards) {
  if (!cards.length) return [];
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Nominees'), head = sh ? header(sh) : [], mine = [];
  if (sh && nominationId && sh.getLastRow() >= 2) {
    sh.getRange(2, 1, sh.getLastRow() - 1, head.length).getValues().forEach(function (r) {
      if (String(r[0]) !== nominationId) return;
      var o = {};
      head.forEach(function (h, k) { o[h] = r[k]; });
      mine.push(o);
    });
  }
  var unsub = unsubscribed();
  return cards.map(function (c) {
    var email = lower(c.contact), old = null;
    for (var i = 0; email && i < mine.length && !old; i++) if (lower(mine[i].contact) === email) old = mine[i];
    var declined = !!email && !!unsub[email];
    if (declined) { c.name = ''; c.contact = ''; c.affiliation = ''; c.why = ''; }
    return { old: old, declined: declined };
  });
}

/* Rewrites a nomination's Nominees rows. Columns outside NOM_DATA (decision, decided_at, decision_notice, the mailer's
   *_sent_at, saved_at) are carried over from the nominee's old row, so consent, sent emails and the 30-day clock survive. */
function writeNominees(nominationId, json, now, plan) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Nominees');
  var last = sh.getLastRow();
  /* resubmission replaces this nomination's rows */
  if (last >= 2) {
    var ids = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var r = ids.length - 1; r >= 0; r--) if (ids[r][0] === nominationId) sh.deleteRow(r + 2);
  }
  var cards = json.__cards || [];
  if (!cards.length) return;
  var head = header(sh);
  if ((plan || []).some(function (p) { return p.declined; })) head = addColumns(sh, head, ['decision', 'decided_at', 'decision_notice']);
  var rows = cards.map(function (c, i) {
    var p = (plan || [])[i] || {}, rec = {};
    if (p.old) head.forEach(function (h) { if (NOM_DATA.indexOf(h) === -1 && h in p.old) rec[h] = p.old[h]; });
    rec.nomination_id = nominationId; rec.application_id = json['B.application_id'] || ''; rec.nominator_name = json['B.name'] || '';
    rec.nominator_email = json['B.email'] || ''; rec.n = i + 1; rec.type = c.type; rec.name = c.name; rec.contact = c.contact;
    rec.affiliation = c.affiliation; rec.why = c.why;
    if (!rec.saved_at) rec.saved_at = now;
    if (p.declined) { rec.decision = 'declined'; rec.decided_at = rec.decided_at || now; }
    return head.map(function (h) { var x = rec[h]; return x === undefined || x === null ? '' : NOM_DATA.indexOf(h) !== -1 && typeof x === 'string' ? cellText(x) : x; });
  });
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, head.length).setValues(rows);
}

/* ---------- input hygiene ---------- */

/* Keeps only known key shapes and caps sizes.
   flat = one sheet column per variable; json = what resume restores (includes __keywords, __cards). */
function clean(answers, keyRe) {
  var flat = {}, json = {}, n = 0;
  Object.keys(answers).forEach(function (k) {
    if (n >= MAX_KEYS) return;
    var v = answers[k];
    if (/^(A7\.file|S\.photo)_(name|url)$/.test(k)) return; /* server-owned columns */
    if (keyRe.test(k)) {
      if (typeof v === 'boolean') { flat[k] = v; json[k] = v; n++; }
      else if (typeof v === 'string' || typeof v === 'number') { var s = String(v).slice(0, MAX_VALUE); flat[k] = s; json[k] = s; n++; }
    } else if (k === '__keywords' && Array.isArray(answers[k])) {
      json[k] = v.slice(0, 5).map(function (x) { return String(x).slice(0, 200); });
    } else if (k === '__cards' && Array.isArray(v)) {
      json[k] = v.slice(0, MAX_CARDS).map(function (c) {
        c = c || {};
        return { type: String(c.type || '').slice(0, 20), name: String(c.name || '').slice(0, 300), contact: String(c.contact || '').slice(0, 300), affiliation: String(c.affiliation || '').slice(0, 300), why: String(c.why || '').slice(0, 1000) };
      });
    } else if (k === '__screen' || k === '__prefilled') {
      json[k] = v;
    } else if (k === '__layout' && typeof v === 'number') {
      json[k] = v; /* screen order the __screen index refers to; see LAYOUT in public/apply/index.html */
    } else if (k === '__file' && v && typeof v === 'object') {
      /* what the client remembers about an earlier upload; the server copy is rewritten whenever a file is sent */
      json[k] = { name: String(v.name || '').slice(0, 200), size: +v.size || 0, url: /^https:\/\/(drive|docs)\.google\.com\//.test(v.url) ? v.url : '' };
    }
  });
  return { flat: flat, json: json };
}

/* A leading apostrophe makes Sheets keep the cell as typed, so =1+1, 0123, 1/2, 50% and TRUE are not read as formulas, numbers or dates.
   Plain whole numbers stay numbers. */
function cellText(s) { return s === '' || /^(0|[1-9][0-9]{0,14})$/.test(s) ? s : "'" + s; }

/* token_hash holds the browser's token plus one more per email sent, separated by spaces. Plain tokens are never stored. */
function tokenOk(stored, token) { return !!token && String(stored).split(' ').indexOf(hash(token)) !== -1; }
function newToken() { return Utilities.getUuid() + Utilities.getUuid().slice(0, 8); }

function hash(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s), Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join('');
}

/* ---------- email ---------- */

/* Gmail path: Peter's verified Gmail alias. A consumer Gmail account sends to about 100 recipients a day.
   Gmail drops "Send mail as" for third-party addresses in January 2027, so mail goes through Resend once the
   script property RESEND_API_KEY is set: from RESEND_FROM (default info@v4air.eu), DKIM-signed for v4air.eu. Free plan: 100 a day. */
var MAIL_FROM = 'peter@v4air.eu', MAIL_NAME = 'V4 AI Researchers Meetup', MAIL_REPLY = 'info@v4air.eu';
var RESEND_FROM = 'info@v4air.eu', RESEND_DAILY = 100;
/* Secrets and mail settings reach the script without the editor: a JSON file with this name, e.g.
   {"RESEND_API_KEY": "re_...", "UNSUB_SECRET": "...", "MAIL_MODE": "redirect"}, is read into script properties on the next run.
   Two places are read: petkout1's My Drive root (the file is then trashed), and the folder in script property
   SETTINGS_FOLDER_ID, Petr's own folder shared with petkout1, so Petr can change the mail settings himself (that file
   is read once per change, as petkout1 cannot trash it; Petr deletes it after the import). Only files in these two places
   count, never a file someone else shares, and only the keys listed here are accepted; an empty value removes the property. */
var SECRETS_FILE = 'v4air-script-secrets.json';
var SECRET_KEYS = ['RESEND_API_KEY', 'UNSUB_SECRET', 'RESEND_FROM', 'EMAIL_ENABLED', 'MAIL_MODE', 'MAIL_KINDS', 'MAIL_REDIRECT', 'MAIL_START',
  'EMAILS_DOC_ID', 'SETTINGS_FOLDER_ID', 'NOMINEE_PURGE', 'MAIL_ONLY', 'MAIL_TEST_DAY_MINUTES'];
var MAIL_BATCH = 5; /* the script lock is held during a run, and saves through Apps Script wait for it at most 20 s */
var HOUR = 60 * 60 * 1000, DAY = 24 * HOUR, LATE_MS = 3 * DAY, MAX_TRIES = 3, NOMINEE_KEEP_MS = 30 * DAY;
var DEADLINE = new Date('2027-01-16T00:00:00+01:00'); /* applications close on 15 January 2027: after it only confirmations go out */
var DEADLINE_WEEK = new Date('2027-01-08T09:00:00+01:00'), DEADLINE_LAST = new Date('2027-01-14T09:00:00+01:00');
var NO_RESUME2_FROM = new Date('2027-01-05T09:00:00+01:00'); /* "skipped if a deadline reminder is due within 3 days" */
/* example.com, .net and .org are reserved and never receive mail (RFC 2606); the form tests use them, so their rows are marked, not mailed.
   Other test addresses: a row in the tab "Unsubscribed" with source "test" stops every email to that address, confirmations included. */
var TEST_ADDRESS = /@example\.(com|net|org)$/i;
var PREVIEW_TAB = 'Mail preview', UNSUB_TAB = 'Unsubscribed', LOG_TAB = 'Email log', NOMINEE_TAB = 'Nominees';
var LOG_HEAD = ['sent_at', 'mode', 'form', 'kind', 'response_id', 'to', 'sent_to', 'subject', 'status', 'detail'];
var MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

/* Every email the mailer sends, in priority order: a person gets at most one email a day, the first one due here goes and
   the rest wait for the next day. form A = Applications, S = Speakers, B = Nominations (the nominator), N = Nominees.
   tab = the tab of the emails Doc. Timing: after = days after base ('saved' = last save, 'sent' = form sent,
   'prev' = when the previous email of the chain went); prev = that email must have gone first; from/until = a date window.
   late = a "<tab> (late)" text is used when the email goes more than 3 days after the form was sent or last saved. */
var FLOWS = [
  { form: 'A', kind: 'confirm', tab: 'Application confirmation', late: true },
  { form: 'S', kind: 'confirm', tab: 'Speaker confirmation', late: true },
  { form: 'B', kind: 'confirm', tab: 'Nomination confirmation' },
  { form: 'A', kind: 'deadline1', tab: 'Deadline reminder: 1 week', from: DEADLINE_WEEK, until: DEADLINE_LAST },
  { form: 'A', kind: 'deadline2', tab: 'Deadline reminder: last day', from: DEADLINE_LAST, until: DEADLINE },
  { form: 'A', kind: 'resume', tab: 'Application resume', after: 1, base: 'saved', late: true },
  { form: 'A', kind: 'resume2', tab: 'Application resume 2', after: 7, base: 'saved', prev: 'resume', until: NO_RESUME2_FROM },
  { form: 'S', kind: 'resume', tab: 'Speaker resume', after: 1, base: 'saved', late: true },
  { form: 'S', kind: 'resume2', tab: 'Speaker resume 2', after: 4, base: 'saved', prev: 'resume' },
  { form: 'S', kind: 'resume3', tab: 'Speaker resume 3', after: 10, base: 'saved', prev: 'resume2' },
  { form: 'N', kind: 'invite', tab: 'Nominee invitation' },
  { form: 'N', kind: 'remind', tab: 'Nominee reminder', after: 7, base: 'prev', prev: 'invite' },
  { form: 'A', kind: 'nominate1', tab: 'Nominate reminder 1', after: 3, base: 'sent', prev: 'confirm' },
  { form: 'A', kind: 'nominate2', tab: 'Nominate reminder 2', after: 14, base: 'sent', prev: 'nominate1' },
  { form: 'A', kind: 'nominate3', tab: 'Nominate reminder 3', after: 28, base: 'sent', prev: 'nominate2' }
];
var WANTS_STATUS = { confirm: 'submitted', resume: 'in progress', resume2: 'in progress', resume3: 'in progress', deadline1: 'in progress',
  deadline2: 'in progress', nominate1: 'submitted', nominate2: 'submitted', nominate3: 'submitted', invite: 'nominee', remind: 'nominee' };

/* Runs every 5 minutes (setup() makes the trigger). This is the only sender, because the Worker cannot send mail.
   Memory: a column <kind>_sent_at on each tab, and the tab "Email log" (every attempt), which survives a nomination being
   sent again (that rewrites its Nominees rows).
   Script property MAIL_MODE:
     'dry' (default when unset): every run rewrites the tab "Mail preview" with each email that is due now and every one that is
       skipped, and why, plus the nominee details the daily check would delete. Nothing is sent, changed or deleted.
     'redirect': due emails go to the addresses in MAIL_REDIRECT (comma separated) instead, with the real recipient in the subject.
       Rows are not marked and no token is made (the links are placeholders); each email goes once per row, recorded in "Email log".
     'live': real mail. The first live run stores MAIL_START.
   Redirect and live also need EMAIL_ENABLED=true and UNSUB_SECRET. MAIL_KINDS (comma separated kinds from FLOWS, e.g. 'confirm'
   or 'confirm, resume'; unset = all) switches kinds on one at a time. Backlog: a row saved before MAIL_START counts as saved
   at MAIL_START, so earlier applicants get their confirmation on day 0 and reminders count from go-live.
   Nominee details: deleted once a day by purgeNominees in live mode, or when NOMINEE_PURGE=true. Texts: see loadEmails.
   Test mode: MAIL_ONLY (comma separated addresses, * matches any text, e.g. "chlupp+*@natur.cuni.cz") limits every email
   and the nominee deletion to those addresses; all other rows stay as they are, MAIL_START is not stored and "Mail preview"
   still shows what the others would get. MAIL_TEST_DAY_MINUTES shortens a "day" for those addresses only (every timing:
   reminders, late texts, one email a day, the 30-day deletion), so a person can walk through every email in an afternoon. */
function sendMail() {
  var props = PropertiesService.getScriptProperties();
  importSecrets(props);
  var cfg = mailConfig(props);
  if (!cfg) return;
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return;
  try {
    if (cfg.mode === 'dry') {
      writePreview(cfg);
    } else {
      if (cfg.mode === 'live' && !cfg.only && !props.getProperty('MAIL_START')) props.setProperty('MAIL_START', cfg.start.toISOString());
      ['A', 'S', 'B', 'N'].forEach(function (form) {
        var sh = mailSheet(form);
        if (sh) addColumns(sh, header(sh), FLOWS.filter(function (f) { return f.form === form; }).map(function (f) { return f.kind + '_sent_at'; }));
      });
      var log = emailLog(true), limit = cfg.resend ? RESEND_DAILY - log.sentToday : MailApp.getRemainingDailyQuota();
      var items = dueMail(cfg, log).filter(function (it) { return !it.skip; }).slice(0, Math.max(0, Math.min(MAIL_BATCH, limit)));
      for (var i = 0; i < items.length; i++) if (!sendOne(cfg, log, items[i])) break;
      if (cfg.only) writePreview(cfg);
    }
    var today = Utilities.formatDate(cfg.now, 'Europe/Prague', 'yyyy-MM-dd');
    if (cfg.mode === 'live' && cfg.only) purgeNominees(cfg, true); /* test mode: only the test nominees, on every run */
    else if ((cfg.mode === 'live' || props.getProperty('NOMINEE_PURGE') === 'true') && props.getProperty('PURGE_DAY') !== today) {
      purgeNominees(cfg, true);
      props.setProperty('PURGE_DAY', today);
    }
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
}

/* null when the mailer must not run at all */
function mailConfig(props) {
  var p = function (k) { return String(props.getProperty(k) || '').trim(); };
  var site = (p('SITE_URL') || 'https://v4air.eu/').replace(/\/?$/, '/');
  var mode = p('MAIL_MODE') === 'live' || p('MAIL_MODE') === 'redirect' ? p('MAIL_MODE') : 'dry';
  var list = function (k) { return p(k).toLowerCase().split(/[\s,;]+/).filter(String); };
  var kinds = list('MAIL_KINDS'), redirect = list('MAIL_REDIRECT'), only = list('MAIL_ONLY');
  /* test mode: no backlog shift for the test rows, which are new anyway */
  var start = p('MAIL_START') ? new Date(p('MAIL_START')) : only.length ? new Date(0) : new Date();
  if (!/^https:\/\//.test(site) || isNaN(start)) return null;
  if (mode !== 'dry') {
    if (p('EMAIL_ENABLED') !== 'true') return null;
    if (!p('UNSUB_SECRET')) { console.error('MAIL_MODE=' + mode + ' needs UNSUB_SECRET: nothing sent'); return null; }
    if (mode === 'redirect' && !redirect.length) { console.error('MAIL_MODE=redirect needs MAIL_REDIRECT: nothing sent'); return null; }
  }
  var on = {};
  kinds.forEach(function (k) { on[k] = true; });
  var res = only.map(function (x) { return new RegExp('^' + x.replace(/[.+?^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*') + '$'); });
  var minutes = +p('MAIL_TEST_DAY_MINUTES');
  var cfg = { site: site, mode: mode, kinds: kinds.length ? on : null, redirect: redirect, start: start, now: new Date(), resend: !!p('RESEND_API_KEY'),
    only: res.length ? function (to) { to = lower(to); return res.some(function (r) { return r.test(to); }); } : null,
    testDay: res.length && minutes > 0 ? minutes * 60 * 1000 : 0,
    secret: p('UNSUB_SECRET'), emails: {}, textError: '' };
  try { cfg.emails = loadEmails(p('EMAILS_DOC_ID') || EMAILS_DOC); } catch (e) { cfg.textError = 'cannot read the emails Doc: ' + String(e).slice(0, 120); }
  return cfg;
}

/* how long a "day" is for this address: shorter only for test addresses in test mode */
function dayMs(cfg, to) { return cfg.only && cfg.testDay && cfg.only(to) ? cfg.testDay : DAY; }

function mailSheet(form) {
  if (form === 'N') return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOMINEE_TAB);
  return sheetFor(form);
}

/* The rows of one tab as records the flows can read. Nominees: one record per nominee, id "<nomination id>.<n>". */
function mailRecords(form) {
  var sh = mailSheet(form);
  if (!sh || sh.getLastRow() < 2) return [];
  var head = header(sh), col = {};
  head.forEach(function (h, i) { col[h] = i; });
  var g = function (r, k) { return col[k] === undefined ? '' : r[col[k]]; };
  return sh.getRange(2, 1, sh.getLastRow() - 1, head.length).getValues().map(function (r, i) {
    var rec = { form: form, row: i + 2, col: col, marks: {}, tokenHash: g(r, 'token_hash') };
    FLOWS.forEach(function (f) { if (f.form === form) rec.marks[f.kind] = g(r, f.kind + '_sent_at'); });
    if (form === 'N') {
      var contact = String(g(r, 'contact') || '').trim();
      rec.id = g(r, 'nomination_id') + '.' + g(r, 'n'); rec.nominationId = String(g(r, 'nomination_id')); rec.n = +g(r, 'n');
      rec.to = EMAIL_RE.test(contact) ? contact : ''; rec.contact = contact; rec.name = String(g(r, 'name') || '').trim();
      rec.wiped = !rec.name && !contact; rec.decision = String(g(r, 'decision') || '');
      rec.status = rec.wiped || rec.decision === 'declined' ? 'gone' : 'nominee';
      rec.saved = asDate(g(r, 'saved_at')); rec.sent = rec.saved; rec.nominator = String(g(r, 'nominator_name') || '').trim();
      return rec;
    }
    var p = form === 'B' ? 'B.' : 'A1.';
    rec.id = String(g(r, 'response_id')); rec.status = g(r, 'status');
    rec.to = String(g(r, p + 'email') || '').trim(); rec.name = String(g(r, p + 'name') || '').trim();
    rec.saved = asDate(g(r, 'last_saved_at')); rec.sent = asDate(g(r, 'submitted_at')); rec.started = asDate(g(r, 'started_at'));
    rec.applicationId = form === 'B' ? String(g(r, 'B.application_id') || '').trim().toUpperCase() : '';
    return rec;
  });
}

/* Everything due now, in sending order, without changing anything. Each item: {flow, rec, late, skip ('' or why)}.
   Emails that already went, that wait for an earlier email of their chain, or whose date window has not opened are not listed. */
function dueMail(cfg, log) {
  var unsub = unsubscribed(), recs = {}, items = [];
  ['A', 'S', 'B', 'N'].forEach(function (f) { recs[f] = mailRecords(f); });
  var submitted = { A: {}, S: {} }, nominated = {};
  ['A', 'S'].forEach(function (f) { recs[f].forEach(function (r) { if (r.status === 'submitted') submitted[f][lower(r.to)] = true; }); });
  recs.B.forEach(function (r) { if (r.status === 'submitted') { nominated[lower(r.to)] = true; if (r.applicationId) nominated[r.applicationId] = true; } });
  FLOWS.forEach(function (flow) {
    recs[flow.form].forEach(function (rec) {
      if (rec.status !== WANTS_STATUS[flow.kind]) return;
      if (sentAt(cfg, log, flow, flow.kind, rec)) return;
      if (flow.from && cfg.now < flow.from) return;
      if (flow.until && cfg.now >= flow.until) return;
      var prevAt = flow.prev ? sentAt(cfg, log, flow, flow.prev, rec) : null;
      if (flow.prev && !prevAt) return;
      var to = lower(rec.to), it = { flow: flow, rec: rec, skip: '' }, base = null, D = dayMs(cfg, to);
      if (flow.base === 'saved') base = rec.saved;
      else if (flow.base === 'sent') base = rec.sent;
      else if (flow.base === 'prev') base = isDate(prevAt) ? prevAt : null;
      if (base && flow.base !== 'prev') base = new Date(Math.max(base, cfg.start));
      var when = flow.kind === 'confirm' ? rec.sent : rec.saved;
      it.late = !!flow.late && isDate(when) && cfg.now - when > LATE_MS / DAY * D && !!cfg.emails[flow.tab + ' (late)'];
      var problem = textProblem(cfg, it);
      if (!to) it.skip = flow.form === 'N' ? 'no email address (contact: ' + (rec.contact || 'empty') + ')' : 'no email';
      else if (cfg.only && !cfg.only(to)) it.skip = 'test mode: not on MAIL_ONLY';
      else if (TEST_ADDRESS.test(to)) it.skip = 'test address';
      else if (unsub[to] === 'test') it.skip = 'test address (tab Unsubscribed, source test)';
      else if (flow.kind !== 'confirm' && unsub[to]) it.skip = 'unsubscribed';
      else if (flow.kind !== 'confirm' && cfg.now >= DEADLINE) it.skip = 'after 15 January 2027: only confirmations';
      else if ((flow.form === 'A' || flow.form === 'S') && WANTS_STATUS[flow.kind] === 'in progress' && submitted[flow.form][to]) it.skip = 'same email already sent this form';
      else if (/^nominate/.test(flow.kind) && (nominated[to] || nominated[String(rec.id).toUpperCase()])) it.skip = 'already nominated someone';
      else if (flow.form === 'N' && submitted.A[to]) it.skip = 'already applied';
      else if (flow.after && !base) it.skip = 'no date to count from';
      else if (flow.after && cfg.now - base < flow.after * D) it.skip = 'not yet: due ' + Utilities.formatDate(new Date(+base + flow.after * D), 'Europe/Prague', 'dd.MM.yyyy HH:mm');
      else if (problem) it.skip = problem;
      else if (cfg.kinds && !cfg.kinds[flow.kind]) it.skip = 'kind not switched on (MAIL_KINDS)';
      items.push(it);
    });
  });
  /* each reminder once per address, at most one email per person per day: the first one in FLOWS order goes */
  var today = cfg.mode === 'redirect' ? log.redirectToday : log.liveToday, last = cfg.mode === 'redirect' ? log.redirectLast : log.liveLast, seen = {}, once = {};
  items.forEach(function (it) {
    if (it.skip) return;
    var to = lower(it.rec.to), key = it.flow.form + '|' + it.flow.kind + '|' + to, D = dayMs(cfg, to);
    /* a calendar day in Prague; a test address in test mode: a short rolling "day" */
    var busy = seen[to] || (D === DAY ? !!today[to] : isDate(last[to]) && cfg.now - last[to] < D);
    if (once[key] && it.flow.kind !== 'confirm') it.skip = 'goes to this address from another row';
    else if (busy) it.skip = 'one email a day: waits for tomorrow';
    once[key] = true; seen[to] = true;
  });
  return items;
}

/* When (or whether) an email of this kind went to this record: its *_sent_at cell, or the Email log (by id, then by address,
   since each reminder goes once per person; a confirmation goes for every form sent). A mark like "skipped: test address" or "failed: ..." counts as done.
   Redirect mode also counts what went in redirect mode, so a rehearsal walks through the chains too. */
function sentAt(cfg, log, flow, kind, rec) {
  var m = rec.marks[kind];
  if (m !== '' && m !== undefined) return isDate(m) ? m : true;
  var books = cfg.mode === 'redirect' ? [log.live, log.redirect] : [log.live];
  for (var i = 0; i < books.length; i++) {
    var hit = books[i][flow.form + '|' + kind + '|' + rec.id] || (kind !== 'confirm' && rec.to && books[i][flow.form + '|' + kind + '|@' + lower(rec.to)]);
    if (hit) return hit;
  }
  return null;
}

/* Sends one item; false stops the run (Resend refuses more today). */
function sendOne(cfg, log, it) {
  var live = cfg.mode === 'live', rec = it.rec, flow = it.flow, when = new Date(), sh = mailSheet(flow.form), link = '<personal link: made only for the real email>';
  if (live && (flow.form === 'A' || flow.form === 'S')) {
    var token = newToken();
    sh.getRange(rec.row, rec.col.token_hash + 1).setValue(String(rec.tokenHash) + ' ' + hash(token));
    link = mailLink(cfg.site, flow.form, rec.id, token);
  }
  var m = emailText(cfg, it, link, live ? unsubscribeLink(rec.to) : '<unsubscribe link of ' + rec.to + '>', live ? acceptLink(cfg, rec.id) : '<accept link of ' + rec.id + '>');
  var targets = live ? [rec.to] : cfg.redirect, subject = live ? m.subject : '[to ' + rec.to + '] ' + m.subject;
  var key = flow.form + '|' + flow.kind + '|' + rec.id, ok = true, detail = '';
  try {
    detail = deliver(targets, subject, m.body, live ? rec.to : targets[0]);
  } catch (err) {
    ok = false; detail = String(err).slice(0, 300);
    console.error('mail failed: ' + key + ': ' + detail);
  }
  logMail([when, cfg.mode, flow.form, flow.kind + (it.late ? ' (late)' : ''), rec.id, rec.to, targets.join(', '), subject, ok ? 'sent' : 'failed', detail]);
  var book = live ? log.live : log.redirect;
  if (ok) { book[key] = when; book[flow.form + '|' + flow.kind + '|@' + lower(rec.to)] = when; (live ? log.liveLast : log.redirectLast)[lower(rec.to)] = when; }
  if (live) {
    var tries = (log.failures[key] || 0) + (ok ? 0 : 1), col = rec.col[flow.kind + '_sent_at'];
    /* marked only on success, or after MAX_TRIES failures, so a passing error is retried on later runs and one bad address cannot block the queue */
    if (col !== undefined && ok) sh.getRange(rec.row, col + 1).setValue(when);
    else if (col !== undefined && tries >= MAX_TRIES) sh.getRange(rec.row, col + 1).setValue('failed: ' + detail.slice(0, 80));
    log.failures[key] = tries;
  }
  return ok || !/Resend 429/.test(detail);
}

function mailLink(site, form, id, token) {
  return site + (form === 'S' ? 'invited-speaker/' : 'apply/') + '?id=' + encodeURIComponent(id) + '&t=' + encodeURIComponent(token);
}

/* https://v4air.eu/accept-nomination/?id=<nomination id>.<n>&t=<b64url HMAC-SHA256(UNSUB_SECRET, 'nomination:' + id)>,
   the link the Worker's nomination() checks (site PR #41). */
function acceptLink(cfg, id) {
  if (!cfg.secret) return '';
  var b64 = function (x) { return Utilities.base64EncodeWebSafe(x).replace(/=+$/, ''); };
  return cfg.site + 'accept-nomination/?id=' + encodeURIComponent(id) + '&t=' + b64(Utilities.computeHmacSha256Signature('nomination:' + id, cfg.secret));
}

/* The email texts live in a Google Doc, never in this code: one tab per email, named as in FLOWS (a "(late)" tab is
   optional), each with a line "Subject: ..." and the body after it; lines above the subject (the Doc's "Sent:" rule) are ignored.
   Script property EMAILS_DOC_ID picks the Doc (default: "V4AIR form emails"; Typecase can export into another Doc of the same shape).
   The Doc is read once per run, so an edit there goes out from the next run on. An email whose text is missing or has an
   unknown placeholder, or lacks its link ({LINK}, or {ACCEPT_LINK} for nominees) or {UNSUBSCRIBE_LINK}, is not sent; the preview says why. */
var EMAILS_DOC = '1EkQNlVa_DGdI5lYP727lhxpAr0XhKTinMEud09Y_g9g';
var PLACEHOLDERS = ['NAME', 'ID', 'LINK', 'NOMINATE_LINK', 'UNSUBSCRIBE_LINK', 'SENT_DATE', 'STARTED_DATE', 'NOMINATOR', 'ACCEPT_LINK'];

function loadEmails(id) {
  var out = {};
  var walk = function (tabs) {
    tabs.forEach(function (t) {
      var e = parseEmail(t.asDocumentTab().getBody().getText());
      if (e) out[t.getTitle().trim()] = e;
      walk(t.getChildTabs());
    });
  };
  walk(DocumentApp.openById(id).getTabs());
  return out;
}

/* {subject, body} from one tab's text, or null when the tab has no "Subject:" line (such as "How it works") */
function parseEmail(text) {
  var lines = String(text).replace(/\r\n?|\u000b/g, '\n').split('\n'), i = -1;
  for (var j = 0; j < lines.length && i < 0; j++) if (/^Subject:/.test(lines[j])) i = j;
  if (i < 0) return null;
  return { subject: lines[i].replace(/^Subject:\s*/, '').trim(), body: lines.slice(i + 1).join('\n').replace(/^\n+|\s+$/g, '') + '\n' };
}

function emailFor(cfg, it) {
  var tab = it.flow.tab;
  return { tab: tab, t: (it.late && cfg.emails[tab + ' (late)']) || cfg.emails[tab] };
}

/* '' when the item's text is usable, otherwise why not */
function textProblem(cfg, it) {
  if (cfg.textError) return cfg.textError;
  var e = emailFor(cfg, it), all = e.t ? e.t.subject + '\n' + e.t.body : '', need = it.flow.form === 'N' ? '{ACCEPT_LINK}' : '{LINK}';
  if (!e.t) return 'no tab "' + e.tab + '" with a Subject line in the emails Doc';
  if (!e.t.subject) return 'empty subject in tab "' + e.tab + '"';
  var bad = (all.match(/\{[A-Z_]+\}/g) || []).filter(function (x) { return PLACEHOLDERS.indexOf(x.slice(1, -1)) === -1; });
  if (bad.length) return 'unknown placeholder ' + bad[0] + ' in tab "' + e.tab + '"';
  if (e.t.body.indexOf(need) === -1) return 'no ' + need + ' in tab "' + e.tab + '"';
  if (e.t.body.indexOf('{UNSUBSCRIBE_LINK}') === -1) return 'no {UNSUBSCRIBE_LINK} in tab "' + e.tab + '"';
  return '';
}

/* The item's text from the emails Doc with the placeholders filled in.
   {LINK}: the personal form link (applications, speakers), the nominate page (nominations, nominate reminders) or the apply page (nominees). */
function emailText(cfg, it, link, unsubLink, acceptLinkText) {
  var e = emailFor(cfg, it), tab = e.tab, t = e.t, rec = it.rec, form = it.flow.form;
  if (!t) throw new Error('no email text for ' + tab);
  var utm = '?utm_source=forms-mail&utm_medium=email&utm_campaign=' + encodeURIComponent(tab.toLowerCase().replace(/\W+/g, '-'));
  var nominate = cfg.site + 'nominate/' + utm;
  var fill = {
    ID: form === 'N' ? rec.nominationId : rec.id, UNSUBSCRIBE_LINK: unsubLink, SENT_DATE: dayText(rec.sent), STARTED_DATE: dayText(rec.started || rec.sent),
    NOMINATE_LINK: nominate, NOMINATOR: rec.nominator || '', ACCEPT_LINK: acceptLinkText || '',
    LINK: form === 'A' && /^nominate/.test(it.flow.kind) ? nominate : form === 'A' || form === 'S' ? link : form === 'B' ? nominate : cfg.site + 'apply/' + utm
  };
  var put = function (s) {
    s = s.replace(/ \{NAME\}/g, rec.name ? ' ' + rec.name : '').replace(/\{NAME\}/g, rec.name);
    return s.replace(/\{([A-Z_]+)\}/g, function (all, k) { return k in fill ? fill[k] : all; });
  };
  return { subject: put(t.subject), body: put(t.body) };
}

/* "12 October 2026" in Prague time, without depending on the script's locale */
function dayText(d) {
  if (!(isDate(d))) return '';
  var p = Utilities.formatDate(d, 'Europe/Prague', 'd M yyyy').split(' ');
  return p[0] + ' ' + MONTHS[+p[1] - 1] + ' ' + p[2];
}

/* Object.prototype.toString, not instanceof: the tests hand in dates from another realm */
function isDate(x) { return Object.prototype.toString.call(x) === '[object Date]' && !isNaN(x); }

/* a Sheets date, or text like "2026-11-12 10:05:00" (Prague time) as some rows store it */
function asDate(x) {
  if (isDate(x)) return x;
  var m = /^(\d{4})-(\d{2})-(\d{2})[ T](\d{1,2}):(\d{2})(?::(\d{2}))?/.exec(String(x || ''));
  return m ? new Date(m[1] + '-' + m[2] + '-' + m[3] + 'T' + ('0' + m[4]).slice(-2) + ':' + m[5] + ':' + (m[6] || '00') + '+01:00') : null;
}

function lower(s) { return String(s || '').trim().toLowerCase(); }

/* {email: source} from the tab "Unsubscribed" */
function unsubscribed() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(UNSUB_TAB), map = {};
  if (!sh || sh.getLastRow() < 2) return map;
  sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues().forEach(function (r) { var e = lower(r[0]); if (e) map[e] = lower(r[2]) || 'unsubscribed'; });
  return map;
}

/* What the tab "Email log" says: what went (live and redirect, keyed "form|kind|id" and "form|kind|@address", value = when),
   today's count and addresses (Prague), and failed live attempts per email. create: make the tab if missing. */
function emailLog(create) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(LOG_TAB);
  var out = { sentToday: 0, liveToday: {}, redirectToday: {}, liveLast: {}, redirectLast: {}, live: {}, redirect: {}, failures: {} };
  if (!sh && create) { sh = ss.insertSheet(LOG_TAB); sh.getRange(1, 1, 1, LOG_HEAD.length).setValues([LOG_HEAD]); sh.setFrozenRows(1); }
  if (!sh || sh.getLastRow() < 2) return out;
  var day = function (d) { return isDate(d) ? Utilities.formatDate(d, 'Europe/Prague', 'yyyy-MM-dd') : ''; }, today = day(new Date());
  sh.getRange(2, 1, sh.getLastRow() - 1, LOG_HEAD.length).getValues().forEach(function (r) {
    var mode = r[1], kind = String(r[3]).replace(' (late)', ''), key = r[2] + '|' + kind + '|' + r[4], sent = r[8] === 'sent';
    var book = mode === 'redirect' ? out.redirect : out.live, at = isDate(r[0]) ? r[0] : true;
    if (sent) {
      book[key] = at; book[r[2] + '|' + kind + '|@' + lower(r[5])] = at;
      var lastBook = mode === 'redirect' ? out.redirectLast : out.liveLast, a = lower(r[5]);
      if (isDate(at) && !(lastBook[a] >= at)) lastBook[a] = at;
    }
    if (sent && day(r[0]) === today) { out.sentToday += String(r[6]).split(',').length; out[mode === 'redirect' ? 'redirectToday' : 'liveToday'][lower(r[5])] = true; }
    if (mode === 'live' && !sent) out.failures[key] = (out.failures[key] || 0) + 1;
  });
  return out;
}

function logMail(row) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_TAB);
  sh.getRange(sh.getLastRow() + 1, 1, 1, row.length).setValues([row.map(function (x) { return isDate(x) ? x : cellText(String(x)); })]);
}

/* Nominee details are kept 30 days at most (privacy notice, site PR #39). Once a day this empties name, contact, affiliation
   and why in Nominees, and the nominee's card in the nomination's answers_json, exactly like the Worker's decline(), for:
   - every nominee nominated more than 30 days ago who has no submitted application with the same address; decision
     becomes 'expired' if it was empty, so an 'accepted' row keeps its consent record;
   - every nominee whose address is in tab Unsubscribed (a "no" by email that a person put there, or a nomination sent again
     after a decline); decision becomes 'declined' if it was empty.
   apply=false only lists them (dry-run preview). Returns [{rec, reason}]. */
function purgeNominees(cfg, apply) {
  var recs = mailRecords('N').filter(function (r) { return !r.wiped; });
  if (!recs.length) return [];
  var unsub = unsubscribed(), applied = {}, out = [];
  mailRecords('A').forEach(function (r) { if (r.status === 'submitted') applied[lower(r.to)] = true; });
  recs.forEach(function (r) {
    var to = lower(r.contact);
    if (cfg.only && !cfg.only(to)) return;
    if (unsub[to]) out.push({ rec: r, reason: 'address in tab Unsubscribed', decision: 'declined' });
    else if (isDate(r.saved) && cfg.now - r.saved > NOMINEE_KEEP_MS / DAY * dayMs(cfg, to) && !applied[to]) out.push({ rec: r, reason: 'nominated more than 30 days ago, not applied', decision: 'expired' });
  });
  if (!apply || !out.length) return out;
  var sh = mailSheet('N');
  addColumns(sh, header(sh), ['decision', 'decided_at']);
  var head = header(sh), col = {};
  head.forEach(function (h, i) { col[h] = i; });
  var noms = sheetFor('B'), nhead = header(noms), nId = nhead.indexOf('response_id'), nJson = nhead.indexOf('answers_json');
  var nrows = noms.getLastRow() >= 2 ? noms.getRange(2, 1, noms.getLastRow() - 1, nhead.length).getValues() : [];
  out.forEach(function (x) {
    ['name', 'contact', 'affiliation', 'why'].forEach(function (k) { if (col[k] !== undefined) sh.getRange(x.rec.row, col[k] + 1).setValue(''); });
    if (!x.rec.decision) { sh.getRange(x.rec.row, col.decision + 1).setValue(x.decision); sh.getRange(x.rec.row, col.decided_at + 1).setValue(new Date()); }
    for (var j = 0; j < nrows.length; j++) {
      if (String(nrows[j][nId]) !== x.rec.nominationId) continue;
      var a = {};
      try { a = JSON.parse(nrows[j][nJson] || '{}'); } catch (e) { a = {}; }
      var c = (a.__cards || [])[x.rec.n - 1];
      if (c) { c.name = ''; c.contact = ''; c.affiliation = ''; c.why = ''; nrows[j][nJson] = JSON.stringify(a); noms.getRange(j + 2, nJson + 1).setValue(cellText(nrows[j][nJson])); }
    }
    console.log('nominee details deleted: ' + x.rec.id + ' (' + x.reason + ')');
  });
  return out;
}

/* Dry run output: one row per due or skipped email, and one per nominee whose details the daily check would delete. */
function writePreview(cfg) {
  var props = PropertiesService.getScriptProperties();
  var from = cfg.resend ? MAIL_NAME + ' <' + (props.getProperty('RESEND_FROM') || RESEND_FROM) + '> via Resend' : MAIL_NAME + ' <' + MAIL_FROM + '> via Gmail';
  var stamp = Utilities.formatDate(cfg.now, 'Europe/Prague', 'dd.MM.yyyy HH:mm');
  var rows = [['generated_at', 'would_send', 'skip_reason', 'form', 'kind', 'email', 'late', 'response_id', 'to', 'from', 'subject', 'body']];
  dueMail(cfg, emailLog(false)).forEach(function (it) {
    var rec = it.rec;
    var m = emailFor(cfg, it).t ? emailText(cfg, it, '<personal link, made at send time>', unsubscribeLink(rec.to) || '<unsubscribe link: UNSUB_SECRET is not set yet>',
      acceptLink(cfg, rec.id) || '<accept link: UNSUB_SECRET is not set yet>') : { subject: '', body: '' };
    rows.push([stamp, it.skip ? 'no' : 'yes', it.skip, it.flow.form, it.flow.kind, it.flow.tab, it.late ? 'yes' : '', rec.id, rec.to, from, m.subject, m.body]);
  });
  purgeNominees(cfg, false).forEach(function (x) {
    rows.push([stamp, 'delete details', x.reason, 'N', 'purge', '', '', x.rec.id, x.rec.contact, '', '', '']);
  });
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(PREVIEW_TAB) || ss.insertSheet(PREVIEW_TAB);
  sh.clearContents();
  sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows.map(function (r) { return r.map(function (x) { return cellText(String(x)); }); }));
  sh.setFrozenRows(1);
}

/* One email to one or more addresses (several only in redirect mode). Resend when RESEND_API_KEY is set, otherwise Gmail.
   unsubFor: the address whose unsubscribe link goes into the List-Unsubscribe header. Returns the Resend id; throws on failure. */
function deliver(to, subject, body, unsubFor) {
  var props = PropertiesService.getScriptProperties(), key = props.getProperty('RESEND_API_KEY');
  if (!key) {
    GmailApp.sendEmail(to.join(','), subject, body, { from: MAIL_FROM, name: MAIL_NAME, replyTo: MAIL_REPLY });
    return 'gmail';
  }
  var msg = { from: MAIL_NAME + ' <' + (props.getProperty('RESEND_FROM') || RESEND_FROM) + '>', to: to, reply_to: MAIL_REPLY, subject: subject, text: body };
  var unsub = unsubscribeLink(unsubFor);
  /* RFC 8058 one-click: mail apps show an Unsubscribe button and POST to the link; the Worker accepts that POST */
  if (unsub) msg.headers = { 'List-Unsubscribe': '<' + unsub + '>', 'List-Unsubscribe-Post': 'List-Unsubscribe=One-Click' };
  var res = UrlFetchApp.fetch('https://api.resend.com/emails', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + key }, payload: JSON.stringify(msg)
  });
  var code = res.getResponseCode();
  if (code < 200 || code >= 300) throw new Error('Resend ' + code + ': ' + res.getContentText().slice(0, 300));
  try { return JSON.parse(res.getContentText()).id || ''; } catch (e) { return ''; }
}

/* https://v4air.eu/unsubscribe/?e=<base64url email>&t=<base64url HMAC-SHA256(UNSUB_SECRET, email)>, the format the Worker checks.
   Empty while UNSUB_SECRET is not set. */
function unsubscribeLink(email) {
  var secret = PropertiesService.getScriptProperties().getProperty('UNSUB_SECRET');
  if (!secret) return '';
  var e = String(email).trim().toLowerCase(), b64 = function (x) { return Utilities.base64EncodeWebSafe(x).replace(/=+$/, ''); };
  return 'https://v4air.eu/unsubscribe/?e=' + b64(Utilities.newBlob(e).getBytes()) + '&t=' + b64(Utilities.computeHmacSha256Signature(e, secret));
}

function importSecrets(props) {
  var take = function (f, where) {
    var data;
    try { data = JSON.parse(f.getBlob().getDataAsString()); } catch (e) { console.error(SECRETS_FILE + ' in ' + where + ' is not valid JSON'); data = {}; }
    var got = SECRET_KEYS.filter(function (k) { return typeof data[k] === 'string'; });
    got.forEach(function (k) { if (data[k]) props.setProperty(k, data[k]); else props.deleteProperty(k); });
    console.log('settings imported from ' + where + ': ' + (got.join(', ') || 'none'));
  };
  var files = DriveApp.getRootFolder().getFilesByName(SECRETS_FILE);
  while (files.hasNext()) { var f = files.next(); take(f, 'My Drive'); f.setTrashed(true); }
  var folder = props.getProperty('SETTINGS_FOLDER_ID');
  if (!folder) return;
  try {
    files = DriveApp.getFolderById(folder).getFilesByName(SECRETS_FILE);
  } catch (e) { console.error('SETTINGS_FOLDER_ID: ' + e); return; }
  while (files.hasNext()) {
    var g = files.next(), stamp = g.getId() + ' ' + g.getLastUpdated().getTime();
    if (props.getProperty('SETTINGS_SEEN') === stamp) continue;
    take(g, 'the settings folder');
    props.setProperty('SETTINGS_SEEN', stamp);
  }
}
