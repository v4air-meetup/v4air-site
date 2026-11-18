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
    if (req.form === 'B' && req.final) writeNominees(id, answers.json, now);
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

function writeNominees(nominationId, json, now) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Nominees');
  var last = sh.getLastRow();
  /* resubmission replaces this nomination's rows */
  if (last >= 2) {
    var ids = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var r = ids.length - 1; r >= 0; r--) if (ids[r][0] === nominationId) sh.deleteRow(r + 2);
  }
  var rows = (json.__cards || []).map(function (c, i) {
    return [nominationId, json['B.application_id'] || '', json['B.name'] || '', json['B.email'] || '', i + 1, c.type, c.name, c.contact, c.affiliation, c.why]
      .map(function (v) { return typeof v === 'string' ? cellText(v) : v; }).concat([now]);
  });
  if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, NOMINEE_COLS.length).setValues(rows);
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
var SECRET_KEYS = ['RESEND_API_KEY', 'UNSUB_SECRET', 'RESEND_FROM', 'EMAIL_ENABLED', 'MAIL_MODE', 'MAIL_KINDS', 'MAIL_REDIRECT', 'MAIL_START', 'EMAILS_DOC_ID', 'SETTINGS_FOLDER_ID'];
var MAIL_BATCH = 5; /* the script lock is held during a run, and saves through Apps Script wait for it at most 20 s */
var HOUR = 60 * 60 * 1000, RESUME_AFTER_MS = 24 * HOUR, LATE_MS = 3 * 24 * HOUR, MAX_TRIES = 3;
var DEADLINE = new Date('2027-01-16T00:00:00+01:00'); /* applications close on 15 January 2027: no application reminders after it */
/* example.com, .net and .org are reserved and never receive mail (RFC 2606); the form tests use them, so their rows are marked, not mailed.
   Other test addresses: a row in the tab "Unsubscribed" with source "test" stops every email to that address, confirmations included. */
var TEST_ADDRESS = /@example\.(com|net|org)$/i;
var PREVIEW_TAB = 'Mail preview', UNSUB_TAB = 'Unsubscribed', LOG_TAB = 'Email log';
var LOG_HEAD = ['sent_at', 'mode', 'form', 'kind', 'response_id', 'to', 'sent_to', 'subject', 'status', 'detail'];
var MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

/* Runs every 5 minutes (setup() makes the trigger). This is the only sender, because the Worker cannot send mail.
   The sheet is the queue: confirm_sent_at and resume_sent_at say what went out to the real person; the tab "Email log" lists every attempt.
   Script property MAIL_MODE:
     'dry' (default when unset): every run rewrites the tab "Mail preview" with each email that is due now and every one that is
       skipped, and why. Nothing is sent, nothing is marked, no token is made.
     'redirect': due emails go to the addresses in MAIL_REDIRECT (comma separated) instead, with the real recipient in the subject.
       Rows are not marked and no token is made (the links are placeholders); each email goes once per row, recorded in "Email log".
     'live': real mail. The first live run stores MAIL_START.
   Redirect and live also need EMAIL_ENABLED=true and UNSUB_SECRET. MAIL_KINDS ('confirm', 'resume' or both, comma separated;
   unset = all) switches kinds on one at a time. Backlog: a row saved before MAIL_START counts as saved at MAIL_START, so earlier
   applicants get their confirmation on day 0 and reminders count from go-live; a confirmation or resume more than 3 days after
   the form was sent or last saved uses the "(late)" text, which says why it comes now. Texts: see loadEmails. Other script properties: SITE_URL, RESEND_*. */
function sendMail() {
  var props = PropertiesService.getScriptProperties();
  importSecrets(props);
  var cfg = mailConfig(props);
  if (!cfg) return;
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return;
  try {
    if (cfg.mode === 'dry') { writePreview(cfg); return; }
    if (cfg.mode === 'live' && !props.getProperty('MAIL_START')) props.setProperty('MAIL_START', cfg.start.toISOString());
    ['A', 'S'].forEach(function (form) { var sh = sheetFor(form); addColumns(sh, header(sh), ['confirm_sent_at', 'resume_sent_at']); });
    var log = emailLog(true), limit = cfg.resend ? RESEND_DAILY - log.sentToday : MailApp.getRemainingDailyQuota();
    var items = dueMail(cfg, log).filter(function (it) { return !it.skip; }).slice(0, Math.max(0, Math.min(MAIL_BATCH, limit)));
    for (var i = 0; i < items.length; i++) if (!sendOne(cfg, log, items[i])) break;
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
  var kinds = list('MAIL_KINDS'), redirect = list('MAIL_REDIRECT'), start = p('MAIL_START') ? new Date(p('MAIL_START')) : new Date();
  if (!/^https:\/\//.test(site) || isNaN(start)) return null;
  if (mode !== 'dry') {
    if (p('EMAIL_ENABLED') !== 'true') return null;
    if (!p('UNSUB_SECRET')) { console.error('MAIL_MODE=' + mode + ' needs UNSUB_SECRET: nothing sent'); return null; }
    if (mode === 'redirect' && !redirect.length) { console.error('MAIL_MODE=redirect needs MAIL_REDIRECT: nothing sent'); return null; }
  }
  var on = {};
  kinds.forEach(function (k) { on[k] = true; });
  var cfg = { site: site, mode: mode, kinds: kinds.length ? on : null, redirect: redirect, start: start, now: new Date(), resend: !!p('RESEND_API_KEY'), emails: {}, textError: '' };
  try { cfg.emails = loadEmails(p('EMAILS_DOC_ID') || EMAILS_DOC); } catch (e) { cfg.textError = 'cannot read the emails Doc: ' + String(e).slice(0, 120); }
  return cfg;
}

/* Everything due now, in sending order (confirmations first), without changing anything.
   Each item: {form, kind, late, row, mark (0-based column of *_sent_at or undefined), to, id, name, started, sent, skip ('' or why)}. */
function dueMail(cfg, log) {
  var unsub = unsubscribed(), items = [];
  ['confirm', 'resume'].forEach(function (kind) {
    ['A', 'S'].forEach(function (form) {
      var sh = sheetFor(form), head = header(sh), last = sh.getLastRow(), col = {}, done = {};
      if (last < 2) return;
      head.forEach(function (h, i) { col[h] = i; });
      var rows = sh.getRange(2, 1, last - 1, head.length).getValues(), mark = col[kind + '_sent_at'];
      /* addresses that already sent this form: their unfinished drafts get no reminder (the "V4AIR form emails" Doc rule) */
      rows.forEach(function (r) { if (r[col.status] === 'submitted') done[lower(r[col['A1.email']])] = true; });
      rows.forEach(function (r, i) {
        if (mark !== undefined && r[mark] !== '') return; /* sent, marked as test or given up after MAX_TRIES */
        if (r[col.status] !== (kind === 'confirm' ? 'submitted' : 'in progress')) return;
        var to = String(r[col['A1.email']] || '').trim(), id = r[col.response_id], saved = r[col.last_saved_at], sent = r[col.submitted_at];
        var it = { form: form, kind: kind, row: i + 2, mark: mark, to: to, id: id, name: String(r[col['A1.name']] || '').trim(),
          started: r[col.started_at], sent: sent, tokenCol: col.token_hash, tokenHash: r[col.token_hash], skip: '' };
        var base = isDate(saved) ? new Date(Math.max(saved, cfg.start)) : null, key = id + '|' + kind;
        it.late = kind === 'confirm' ? isDate(sent) && cfg.now - sent > LATE_MS : isDate(saved) && cfg.now - saved > LATE_MS;
        if (!to) it.skip = 'no email';
        else if (TEST_ADDRESS.test(to)) it.skip = 'test address';
        else if (unsub[lower(to)] === 'test') it.skip = 'test address (tab Unsubscribed, source test)';
        else if (kind !== 'confirm' && unsub[lower(to)]) it.skip = 'unsubscribed';
        else if (!base) it.skip = 'no last_saved_at';
        else if (kind === 'resume' && done[lower(to)]) it.skip = 'same email already submitted this form';
        else if (kind === 'resume' && form === 'A' && cfg.now >= DEADLINE) it.skip = 'after the application deadline';
        else if (kind === 'resume' && cfg.now - base < RESUME_AFTER_MS) it.skip = 'not yet: resume goes 24 h after the last save (or after go-live)';
        else if (textProblem(cfg, it)) it.skip = textProblem(cfg, it);
        else if (cfg.kinds && !cfg.kinds[kind]) it.skip = 'kind not switched on (MAIL_KINDS)';
        else if (cfg.mode === 'redirect' && log.redirected[key]) it.skip = 'already sent in redirect mode';
        items.push(it);
      });
    });
  });
  /* at most one email per person per day: the first one in sending order goes, the rest wait for tomorrow */
  var today = cfg.mode === 'redirect' ? log.redirectToday : log.liveToday, seen = {};
  Object.keys(today).forEach(function (k) { seen[k] = true; });
  items.forEach(function (it) {
    if (it.skip) return;
    if (seen[lower(it.to)]) it.skip = 'one email a day: waits for tomorrow';
    seen[lower(it.to)] = true;
  });
  return items;
}

/* Sends one item; false stops the run (Resend refuses more today). */
function sendOne(cfg, log, it) {
  var live = cfg.mode === 'live', sh = sheetFor(it.form), when = new Date(), link = '<personal link: made only for the real email>';
  if (live) {
    var token = newToken();
    sh.getRange(it.row, it.tokenCol + 1).setValue(String(it.tokenHash) + ' ' + hash(token));
    link = mailLink(cfg.site, it.form, it.id, token);
  }
  var m = emailText(cfg, it, link, live ? unsubscribeLink(it.to) : '<unsubscribe link of ' + it.to + '>');
  var targets = live ? [it.to] : cfg.redirect, subject = live ? m.subject : '[to ' + it.to + '] ' + m.subject;
  var key = it.id + '|' + it.kind, ok = true, detail = '';
  try {
    detail = deliver(targets, subject, m.body, live ? it.to : targets[0]);
  } catch (err) {
    ok = false; detail = String(err).slice(0, 300);
    console.error('mail failed: ' + it.form + ' ' + it.kind + ' ' + it.id + ': ' + detail);
  }
  logMail([when, cfg.mode, it.form, it.kind + (it.late ? ' (late)' : ''), it.id, it.to, targets.join(', '), subject, ok ? 'sent' : 'failed', detail]);
  if (live) {
    var tries = (log.failures[key] || 0) + (ok ? 0 : 1);
    /* marked only on success, or after MAX_TRIES failures, so a passing error is retried on later runs and one bad address cannot block the queue */
    if (ok) sh.getRange(it.row, it.mark + 1).setValue(when);
    else if (tries >= MAX_TRIES) sh.getRange(it.row, it.mark + 1).setValue('failed: ' + detail.slice(0, 80));
    log.failures[key] = tries;
  }
  return ok || !/Resend 429/.test(detail);
}

function mailLink(site, form, id, token) {
  return site + (form === 'S' ? 'invited-speaker/' : 'apply/') + '?id=' + encodeURIComponent(id) + '&t=' + encodeURIComponent(token);
}

/* The email texts live in a Google Doc, never in this code: one tab per email, named as in EMAIL_TABS (a "(late)" tab is
   optional), each with a line "Subject: ..." and the body after it; lines above the subject (the Doc's "Sent:" rule) are ignored.
   Script property EMAILS_DOC_ID picks the Doc (default: "V4AIR form emails"; Typecase can export into another Doc of the same shape).
   The Doc is read once per run, so an edit there goes out from the next run on. An email whose text is missing or has an
   unknown placeholder, no {LINK} or no {UNSUBSCRIBE_LINK} is not sent; the preview says why. */
var EMAILS_DOC = '1EkQNlVa_DGdI5lYP727lhxpAr0XhKTinMEud09Y_g9g';
var EMAIL_TABS = { A: { confirm: 'Application confirmation', resume: 'Application resume' }, S: { confirm: 'Speaker confirmation', resume: 'Speaker resume' } };
var PLACEHOLDERS = ['NAME', 'ID', 'LINK', 'NOMINATE_LINK', 'UNSUBSCRIBE_LINK', 'SENT_DATE', 'STARTED_DATE'];

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
  var tab = EMAIL_TABS[it.form][it.kind];
  return { tab: tab, t: (it.late && cfg.emails[tab + ' (late)']) || cfg.emails[tab] };
}

/* '' when the item's text is usable, otherwise why not */
function textProblem(cfg, it) {
  if (cfg.textError) return cfg.textError;
  var e = emailFor(cfg, it), all = e.t ? e.t.subject + '\n' + e.t.body : '';
  if (!e.t) return 'no tab "' + e.tab + '" with a Subject line in the emails Doc';
  if (!e.t.subject) return 'empty subject in tab "' + e.tab + '"';
  var bad = (all.match(/\{[A-Z_]+\}/g) || []).filter(function (x) { return PLACEHOLDERS.indexOf(x.slice(1, -1)) === -1; });
  if (bad.length) return 'unknown placeholder ' + bad[0] + ' in tab "' + e.tab + '"';
  if (e.t.body.indexOf('{LINK}') === -1) return 'no {LINK} in tab "' + e.tab + '"';
  if (e.t.body.indexOf('{UNSUBSCRIBE_LINK}') === -1) return 'no {UNSUBSCRIBE_LINK} in tab "' + e.tab + '"';
  return '';
}

/* The item's text from the emails Doc with the placeholders filled in */
function emailText(cfg, it, link, unsubLink) {
  var e = emailFor(cfg, it), tab = e.tab, t = e.t;
  if (!t) throw new Error('no email text for ' + tab);
  var fill = {
    ID: it.id, LINK: link, UNSUBSCRIBE_LINK: unsubLink, SENT_DATE: dayText(it.sent), STARTED_DATE: dayText(it.started || it.sent),
    NOMINATE_LINK: cfg.site + 'nominate/?utm_source=forms-mail&utm_medium=email&utm_campaign=' + encodeURIComponent(tab.toLowerCase().replace(/\W+/g, '-'))
  };
  var put = function (s) {
    s = s.replace(/ \{NAME\}/g, it.name ? ' ' + it.name : '').replace(/\{NAME\}/g, it.name);
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

function lower(s) { return String(s || '').trim().toLowerCase(); }

/* {email: source} from the tab "Unsubscribed" */
function unsubscribed() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(UNSUB_TAB), map = {};
  if (!sh || sh.getLastRow() < 2) return map;
  sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues().forEach(function (r) { var e = lower(r[0]); if (e) map[e] = lower(r[2]) || 'unsubscribed'; });
  return map;
}

/* What the tab "Email log" says about today (Prague) and about earlier attempts. create: make the tab if missing. */
function emailLog(create) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(LOG_TAB);
  var out = { sentToday: 0, liveToday: {}, redirectToday: {}, redirected: {}, failures: {} };
  if (!sh && create) { sh = ss.insertSheet(LOG_TAB); sh.getRange(1, 1, 1, LOG_HEAD.length).setValues([LOG_HEAD]); sh.setFrozenRows(1); }
  if (!sh || sh.getLastRow() < 2) return out;
  var day = function (d) { return isDate(d) ? Utilities.formatDate(d, 'Europe/Prague', 'yyyy-MM-dd') : ''; }, today = day(new Date());
  sh.getRange(2, 1, sh.getLastRow() - 1, LOG_HEAD.length).getValues().forEach(function (r) {
    var mode = r[1], key = r[4] + '|' + String(r[3]).replace(' (late)', ''), sent = r[8] === 'sent';
    if (sent && day(r[0]) === today) { out.sentToday += String(r[6]).split(',').length; out[mode === 'redirect' ? 'redirectToday' : 'liveToday'][lower(r[5])] = true; }
    if (mode === 'redirect' && sent) out.redirected[key] = true;
    if (mode === 'live' && !sent) out.failures[key] = (out.failures[key] || 0) + 1;
  });
  return out;
}

function logMail(row) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_TAB);
  sh.getRange(sh.getLastRow() + 1, 1, 1, row.length).setValues([row.map(function (x) { return isDate(x) ? x : cellText(String(x)); })]);
}

/* Dry run output: one row per due or skipped email, rewritten on every run. */
function writePreview(cfg) {
  var props = PropertiesService.getScriptProperties();
  var from = cfg.resend ? MAIL_NAME + ' <' + (props.getProperty('RESEND_FROM') || RESEND_FROM) + '> via Resend' : MAIL_NAME + ' <' + MAIL_FROM + '> via Gmail';
  var stamp = Utilities.formatDate(cfg.now, 'Europe/Prague', 'dd.MM.yyyy HH:mm');
  var rows = [['generated_at', 'would_send', 'skip_reason', 'form', 'kind', 'late', 'response_id', 'to', 'from', 'subject', 'body']];
  dueMail(cfg, emailLog(false)).forEach(function (it) {
    var m = emailFor(cfg, it).t ? emailText(cfg, it, '<personal link, made at send time>', unsubscribeLink(it.to) || '<unsubscribe link: UNSUB_SECRET is not set yet>') : { subject: '', body: '' };
    rows.push([stamp, it.skip ? 'no' : 'yes', it.skip, it.form, it.kind, it.late ? 'yes' : '', it.id, it.to, from, m.subject, m.body]);
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
