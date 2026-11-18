/* The forms mailer (sendMail in backend/Code.gs) on imitated Apps Script services. Its texts come from a copy of the Doc
   "V4AIR form emails" (fixtures/emails-doc.json, refreshed by fixture-from-doc.py).
   Nothing reaches Google, Gmail or Resend: deliveries are recorded. Run: node mail/run.mjs (from tests/). */
import { readFileSync } from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import { createHmac, randomUUID, createHash } from 'node:crypto';
import { Book, appsScript, serial } from '../forms/fake-google.mjs';

const ROOT = path.resolve(path.dirname(new URL(import.meta.url).pathname), '../..');
const HEAD = ['response_id', 'token_hash', 'status', 'started_at', 'last_saved_at', 'submitted_at', 'screen', 'answers_json', 'A1.name', 'A1.email'];
const DAY = 864e5, now = Date.now(), ago = d => serial(new Date(now - d * DAY));
let fails = 0;
const check = (name, ok, got) => { console.log((ok ? 'PASS ' : 'FAIL ') + name + (ok ? '' : '  got: ' + JSON.stringify(got))); if (!ok) fails++; };

function prague(d, pattern) {
  const p = Object.fromEntries(new Intl.DateTimeFormat('en-GB', { timeZone: 'Europe/Prague', year: 'numeric', month: 'numeric', day: 'numeric', hour: '2-digit', minute: '2-digit', hourCycle: 'h23' })
    .formatToParts(new Date(d)).map(x => [x.type, x.value]));
  const two = s => String(s).padStart(2, '0');
  return pattern.replace(/yyyy|MM|M|dd|d|HH|mm/g, t => ({ yyyy: p.year, MM: two(p.month), M: String(+p.month), dd: two(p.day), d: String(+p.day), HH: two(p.hour), mm: two(p.minute) }[t]));
}

const DOC = JSON.parse(readFileSync(new URL('./fixtures/emails-doc.json', import.meta.url)));
const docApp = (tabs, broken) => ({ openById: () => { if (broken) throw new Error('no access'); const wrap = t => ({ getTitle: () => t.title, asDocumentTab: () => ({ getBody: () => ({ getText: () => t.text }) }), getChildTabs: () => t.children.map(wrap) }); return { getTabs: () => tabs.map(wrap), getName: () => 'emails' }; } });

/* a Date the mailer sees: new Date() without arguments is clock.now */
const clockDate = clock => class extends Date { constructor(...a) { if (a.length) super(...a); else super(clock.now); } static now() { return clock.now; } };

function setup(props, { resend = 200, unsub = [], doc = DOC, brokenDoc = false, folderFiles = [], seed, clock } = {}) {
  const book = new Book();
  if (seed) seed(book); else book.seed('Applications', HEAD, [
    ['V4A-OLD1-AAAA', 'h1', 'submitted', ago(31), ago(30), ago(30), '', '{}', 'Olga', 'old@uni.cz'],
    ['V4A-NEW1-AAAA', 'h2', 'submitted', ago(0.1), ago(0.05), ago(0.05), '', '{}', 'Nina', 'new@uni.cz'],
    ['V4A-DRF1-AAAA', 'h3', 'in progress', ago(12), ago(10), '', '', '{}', 'Dan', 'draft@uni.cz'],
    ['V4A-TST1-AAAA', 'h4', 'submitted', ago(1), ago(1), ago(1), '', '{}', 'Test', 'x@example.com'],
    ['V4A-TST2-AAAA', 'h5', 'submitted', ago(1), ago(1), ago(1), '', '{}', 'John Doe', 'a@a.aa'],
    ['V4A-DUP1-AAAA', 'h6', 'in progress', ago(5), ago(5), '', '', '{}', 'Nina', 'NEW@uni.cz'],
    ['V4A-TWC1-AAAA', 'h7', 'submitted', ago(0.2), ago(0.2), ago(0.2), '', '{}', 'Tom', 'twice@uni.cz'],
    ['V4A-TWC2-AAAA', 'h8', 'submitted', ago(0.1), ago(0.1), ago(0.1), '', '{}', 'Tom', 'twice@uni.cz']
  ]);
  if (!seed) book.seed('Speakers', HEAD, []);
  book.seed('Unsubscribed', ['email', 'unsubscribed_at', 'source'], [['a@a.aa', '', 'test'], ...unsub.map(e => [e, '', 'page'])]);
  const sent = [], sha = (alg, s) => [...createHash(alg).update(s, 'utf8').digest()].map(b => (b > 127 ? b - 256 : b));
  const ctx = vm.createContext(Object.assign(appsScript(book, { props }), clock ? { Date: clockDate(clock) } : {}, {
    console: { log() {}, error() {} },
    Utilities: {
      DigestAlgorithm: { SHA_256: 'sha256' }, Charset: { UTF_8: 'utf8' }, getUuid: () => randomUUID(), computeDigest: (a, s) => sha(a, s),
      newBlob: s => ({ getBytes: () => [...Buffer.from(s)] }),
      base64EncodeWebSafe: x => Buffer.from(Array.isArray(x) ? x.map(b => b & 255) : x).toString('base64').replace(/\+/g, '-').replace(/\//g, '_'),
      computeHmacSha256Signature: (v, k) => [...createHmac('sha256', k).update(v).digest()].map(b => (b > 127 ? b - 256 : b)),
      formatDate: (d, tz, pattern) => prague(d, pattern)
    },
    DocumentApp: docApp(doc, brokenDoc),
    DriveApp: {
      getRootFolder: () => ({ getFilesByName: () => ({ hasNext: () => false }) }),
      getFolderById: id => { if (id !== 'petr-folder') throw new Error('no folder'); let i = 0; return { getFilesByName: () => ({ hasNext: () => i < folderFiles.length, next: () => folderFiles[i++] }) }; }
    },
    MailApp: { getRemainingDailyQuota: () => 100 },
    GmailApp: { sendEmail: (to, subject, body) => sent.push({ via: 'gmail', to: to.split(','), subject, body }) },
    UrlFetchApp: {
      fetch: (url, o) => { const m = JSON.parse(o.payload); sent.push({ via: 'resend', ...m, body: m.text });
        const code = typeof resend === 'function' ? resend() : resend; return { getResponseCode: () => code, getContentText: () => (code === 200 ? '{"id":"re_1"}' : 'error') }; }
    }
  }));
  vm.runInContext(readFileSync(path.join(ROOT, 'backend/Code.gs'), 'utf8'), ctx);
  const live = ctx.PropertiesService.getScriptProperties();
  return { book, sent, run: () => ctx.sendMail(), prop: k => live.getProperty(k) };
}
const rows = (book, tab) => book.table(tab).rows;
const by = (rs, id) => rs.filter(r => r.response_id === id);

/* 1. dry run (no MAIL_MODE): preview only */
{
  const t = setup({});
  t.run();
  const p = rows(t.book, 'Mail preview'), yes = p.filter(r => r.would_send === 'yes').map(r => r.response_id + ' ' + r.kind).sort();
  check('dry: nothing sent, no log, no marks', t.sent.length === 0 && !t.book.tab('Email log') && !t.book.table('Applications').head.includes('confirm_sent_at'), t.sent);
  check('dry: backlog confirmations due on day 0, one a day for the same address', JSON.stringify(yes) === JSON.stringify(['V4A-NEW1-AAAA confirm', 'V4A-OLD1-AAAA confirm', 'V4A-TWC1-AAAA confirm']), yes);
  const old = by(p, 'V4A-OLD1-AAAA')[0], fresh = by(p, 'V4A-NEW1-AAAA')[0];
  check('dry: a 30-day-old application gets the late text with its date', old.late === 'yes' && /which you sent on \d{1,2} [A-Z][a-z]+ 2026\./.test(old.body) && /not set up yet/.test(old.body), old.body.slice(0, 200));
  check('dry: a fresh one gets the normal text with the nominate link', fresh.late === '' && /Thank you\. We received/.test(fresh.body) && /nominate\/\?utm_source=forms-mail/.test(fresh.body), fresh.body.slice(0, 120));
  const why = id => (by(p, id)[0] || {}).skip_reason;
  check('dry: test rows skipped', /test address/.test(why('V4A-TST1-AAAA')) && /source test/.test(why('V4A-TST2-AAAA')), [why('V4A-TST1-AAAA'), why('V4A-TST2-AAAA')]);
  check('dry: old draft waits 24 h after go-live; duplicate draft skipped', /not yet/.test(why('V4A-DRF1-AAAA')) && /already sent this form/.test(why("V4A-DUP1-AAAA")), [why('V4A-DRF1-AAAA'), why('V4A-DUP1-AAAA')]);
  check('dry: second twice@ row waits a day', /one email a day/.test(why('V4A-TWC2-AAAA')), why('V4A-TWC2-AAAA'));
}

/* 2. redirect: real mail, but only to the redirect address, rows untouched, each email once */
{
  const t = setup({ EMAIL_ENABLED: 'true', MAIL_MODE: 'redirect', MAIL_REDIRECT: 'petr@test.cz', UNSUB_SECRET: 's' });
  t.run(); const first = t.sent.length; t.run();
  check('redirect: only to the redirect address, recipient in the subject', t.sent.length === 3 && first === 3 && t.sent.every(m => m.to.join() === 'petr@test.cz' && /^\[to [^\]]+\] /.test(m.subject)), t.sent.map(m => [m.to, m.subject]));
  const a = rows(t.book, 'Applications');
  check('redirect: no row marked, no token added', a.every(r => !r.confirm_sent_at) && a.every(r => !String(r.token_hash).includes(' ')), a.map(r => [r.confirm_sent_at, r.token_hash]));
  check('redirect: no working personal or unsubscribe link in the body', t.sent.every(m => !/apply\/\?id=/.test(m.body) && !/unsubscribe\/\?e=/.test(m.body)), t.sent[0].body);
  check('redirect: logged', rows(t.book, 'Email log').filter(r => r.mode === 'redirect' && r.status === 'sent').length === 3, rows(t.book, 'Email log'));
}

/* 3. live, confirmations only, through Resend */
{
  const t = setup({ EMAIL_ENABLED: 'true', MAIL_MODE: 'live', MAIL_KINDS: 'confirm', UNSUB_SECRET: 's', RESEND_API_KEY: 'k' });
  t.run();
  const a = rows(t.book, 'Applications'), marked = a.filter(r => Object.prototype.toString.call(r.confirm_sent_at) === '[object Date]' || typeof r.confirm_sent_at === 'number').map(r => r.response_id).sort();
  check('live: MAIL_START stored on the first live run', !!t.prop('MAIL_START'), t.prop('MAIL_START'));
  check('live: three confirmations sent and marked', JSON.stringify(marked) === JSON.stringify(['V4A-NEW1-AAAA', 'V4A-OLD1-AAAA', 'V4A-TWC1-AAAA']), marked);
  check('live: no resume yet (MAIL_KINDS=confirm)', !a.some(r => r.resume_sent_at), a.map(r => r.resume_sent_at));
  const m = t.sent.find(x => x.to[0] === 'old@uni.cz');
  check('live: real link, unsubscribe link and header', /apply\/\?id=V4A-OLD1-AAAA&t=/.test(m.body) && /Unsubscribe: https:\/\/v4air\.eu\/unsubscribe\/\?e=/.test(m.body) && m.headers['List-Unsubscribe-Post'] === 'List-Unsubscribe=One-Click', m);
  check('live: token hash appended', String(by(a, 'V4A-OLD1-AAAA')[0].token_hash).split(' ').length === 2, by(a, 'V4A-OLD1-AAAA')[0].token_hash);
  t.run();
  check('live: second run the same day sends no more confirmations to the same people', t.sent.length === 3, t.sent.map(x => x.to));
}

/* 4. failures are retried, then given up; a 429 stops the run */
{
  const t = setup({ EMAIL_ENABLED: 'true', MAIL_MODE: 'live', MAIL_KINDS: 'confirm', UNSUB_SECRET: 's', RESEND_API_KEY: 'k' }, { resend: 500 });
  t.run();
  const after1 = rows(t.book, 'Applications').filter(r => r.confirm_sent_at !== '').length;
  t.run(); t.run();
  const a = rows(t.book, 'Applications'), failed = a.filter(r => /^failed/.test(r.confirm_sent_at)).length;
  check('fail: not marked after one failure, marked failed after three', after1 === 0 && failed === 3, { after1, failed });
  const u = setup({ EMAIL_ENABLED: 'true', MAIL_MODE: 'live', UNSUB_SECRET: 's', RESEND_API_KEY: 'k' }, { resend: 429 });
  u.run();
  check('fail: Resend 429 stops the run after one try', u.sent.length === 1, u.sent.length);
}

/* 5. late resume after go-live, unsubscribed, safety switches */
{
  const t = setup({ EMAIL_ENABLED: 'true', MAIL_MODE: 'live', MAIL_START: new Date(now - 2 * DAY).toISOString(), UNSUB_SECRET: 's' }, { unsub: ['old@uni.cz'] });
  t.run();
  const r = t.sent.find(m => m.to[0] === 'draft@uni.cz');
  check('resume: due a day after go-live, late text with the start date', !!r && /^On \d{1,2} [A-Z][a-z]+ 2026 you started/m.test(r.body), r && r.body.slice(0, 160));
  check('unsubscribed address still gets its confirmation', t.sent.some(m => m.to[0] === 'old@uni.cz'), t.sent.map(m => m.to));
  for (const [name, props] of [['live without EMAIL_ENABLED', { MAIL_MODE: 'live', UNSUB_SECRET: 's' }], ['live without UNSUB_SECRET', { EMAIL_ENABLED: 'true', MAIL_MODE: 'live' }], ['redirect without MAIL_REDIRECT', { EMAIL_ENABLED: 'true', MAIL_MODE: 'redirect', UNSUB_SECRET: 's' }]]) {
    const x = setup(props); x.run();
    check(name + ': sends nothing, writes nothing', x.sent.length === 0 && !x.book.tab('Email log') && !x.book.tab('Mail preview'), x.sent.length);
  }
}
/* 6. texts come from the Doc at run time; a broken text or an unreadable Doc stops that email, not the run */
{
  const edited = JSON.parse(JSON.stringify(DOC));
  edited.find(t => t.title === 'Application confirmation').text = edited.find(t => t.title === 'Application confirmation').text.replace('Thank you. We received', 'EDITED IN THE DOC. We received');
  edited.find(t => t.title === 'Application resume (late)').text += '\n{OOPS}';
  const t = setup({ MAIL_START: new Date(now - 2 * DAY).toISOString() }, { doc: edited });
  t.run();
  const p = rows(t.book, 'Mail preview'), fresh = by(p, 'V4A-NEW1-AAAA')[0], drf = by(p, 'V4A-DRF1-AAAA')[0];
  check('doc: an edit in the Doc is used on the next run', /EDITED IN THE DOC/.test(fresh.body), fresh.body.slice(0, 80));
  check('doc: an unknown placeholder stops that email, with the reason', drf.would_send === 'no' && /unknown placeholder \{OOPS\}/.test(drf.skip_reason), drf.skip_reason);
  const u = setup({ EMAIL_ENABLED: 'true', MAIL_MODE: 'live', UNSUB_SECRET: 's' }, { brokenDoc: true });
  u.run();
  check('doc: unreadable Doc sends nothing', u.sent.length === 0, u.sent.length);
}
/* 7. Petr's settings folder: read once per change, never trashed */
{
  const file = { getId: () => 'f1', getLastUpdated: () => new Date(1000), getBlob: () => ({ getDataAsString: () => JSON.stringify({ RESEND_API_KEY: 'from-petr', MAIL_MODE: 'redirect', NOT_ALLOWED: 'x' }) }), setTrashed() { throw new Error('not the owner'); } };
  const t = setup({ SETTINGS_FOLDER_ID: 'petr-folder' }, { folderFiles: [file] });
  t.run();
  check('settings folder: allowed keys imported, others ignored', t.prop('RESEND_API_KEY') === 'from-petr' && t.prop('MAIL_MODE') === 'redirect' && t.prop('NOT_ALLOWED') === null, [t.prop('RESEND_API_KEY'), t.prop('MAIL_MODE')]);
  const u = setup({ SETTINGS_FOLDER_ID: 'petr-folder', SETTINGS_SEEN: 'f1 1000', MAIL_MODE: 'dry' }, { folderFiles: [file] });
  u.run();
  check('settings folder: an unchanged file is not imported again', u.prop('MAIL_MODE') === 'dry', u.prop('MAIL_MODE'));
}
/* 8. every flow over 31 simulated days: order, chains, nominees, one a day, the 30-day deletion */
{
  const T0 = Date.UTC(2026, 10, 20, 9, 0), clock = { now: T0 }, at = d => serial(new Date(T0 + d * DAY));
  const BHEAD = ['response_id', 'token_hash', 'status', 'started_at', 'last_saved_at', 'submitted_at', 'screen', 'answers_json', 'B.name', 'B.email', 'B.application_id'];
  const NHEAD = ['nomination_id', 'application_id', 'nominator_name', 'nominator_email', 'n', 'type', 'name', 'contact', 'affiliation', 'why', 'saved_at'];
  const cards = [{ name: 'Nick', contact: 'nick@uni.cz', affiliation: 'X', why: 'good' }, { name: 'Lin', contact: 'linkedin.com/in/lin', affiliation: 'Y', why: 'w' }, { name: 'Pat', contact: 'applied@uni.cz', affiliation: 'Z', why: 'w' }, { name: 'Dee', contact: 'dee@uni.cz', affiliation: 'Q', why: 'w' }];
  const seed = book => {
    book.seed('Applications', HEAD, [
      ['V4A-APP1-AAAA', 'h1', 'submitted', at(-0.1), at(-0.05), at(-0.05), '', '{}', 'Ana', 'app@uni.cz'],
      ['V4A-DRF2-AAAA', 'h2', 'in progress', at(-0.1), at(-0.05), '', '', '{}', 'Dora', 'draft2@uni.cz'],
      ['V4A-APP3-AAAA', 'h3', 'submitted', at(-0.1), at(-0.05), at(-0.05), '', '{}', 'Pat', 'applied@uni.cz']
    ]);
    book.seed('Speakers', HEAD, [['V4S-SPK1-AAAA', 'h4', 'in progress', at(-0.1), at(-0.05), '', '', '{}', 'Sam', 'spk@uni.cz']]);
    book.seed('Nominations', BHEAD, [['V4N-NOM1-AAAA', 'h5', 'submitted', at(-0.1), at(-0.05), at(-0.05), '', JSON.stringify({ __cards: cards }), 'Ana', 'app@uni.cz', 'V4A-APP1-AAAA']]);
    book.seed('Nominees', NHEAD.concat(['decision']), cards.map((c, i) => ['V4N-NOM1-AAAA', 'V4A-APP1-AAAA', 'Ana', 'app@uni.cz', i + 1, 'researcher', c.name, c.contact, c.affiliation, c.why, at(-0.05), i === 3 ? 'declined' : '']));
  };
  const t = setup({ EMAIL_ENABLED: 'true', MAIL_MODE: 'live', UNSUB_SECRET: 's', RESEND_API_KEY: 'k', MAIL_START: new Date(T0 - 2 * 3600e3).toISOString() }, { seed, clock });
  const day = d => { clock.now = T0 + d * DAY + 60e3; for (let i = 0; i < 4; i++) t.run(); };
  const got = () => t.sent.map(m => m.to[0] + ' ' + m.subject);
  const since = n => t.sent.slice(n).map(m => m.to[0] + ' | ' + m.subject).sort();
  day(0);
  check('day 0: confirmations and the nominee invitation; the nominator\'s confirmation waits (one a day)', JSON.stringify(since(0)) === JSON.stringify([
    'app@uni.cz | We received your V4AIR application V4A-APP1-AAAA', 'applied@uni.cz | We received your V4AIR application V4A-APP3-AAAA', 'nick@uni.cz | You were nominated for V4AIR']), since(0));
  const inv = t.sent.find(m => m.to[0] === 'nick@uni.cz'), tok = createHmac('sha256', 's').update('nomination:V4N-NOM1-AAAA.1').digest('base64url');
  check('invitation names the nominator and has the right accept link', /^Ana nominated you/m.test(inv.body) && inv.body.includes('https://v4air.eu/accept-nomination/?id=V4N-NOM1-AAAA.1&t=' + tok) && /Unsubscribe: https:/.test(inv.body), inv.body.slice(0, 400));
  check('no invitation to a contact without email, to someone who applied, or to a declined nominee', !t.sent.some(m => /linkedin|dee@/.test(m.to[0])) && t.sent.filter(m => m.to[0] === 'applied@uni.cz').length === 1, got());
  const n0 = t.sent.length;
  // the nominator sends the nomination again: the Nominees rows are rewritten without their marks
  t.book.tab('Nominees').rows.forEach((r, i) => { if (i) r.length = 12; });
  day(1);
  check('day 1: nominator\'s confirmation, application and speaker resume; no second invitation after the rows were rewritten', JSON.stringify(since(n0)) === JSON.stringify([
    'app@uni.cz | We received your V4AIR nominations V4N-NOM1-AAAA', 'draft2@uni.cz | Your V4AIR application is saved', 'spk@uni.cz | Your V4AIR speaker form is saved']), since(n0));
  const conf = t.sent.find(m => /nominations/.test(m.subject));
  check('nomination confirmation goes to the nominator and links the nominate page', conf.to[0] === 'app@uni.cz' && /nominate\/\?utm_source=forms-mail/.test(conf.body), conf.body.slice(0, 300));
  let n = t.sent.length; day(3);
  check('day 3: nominate reminder only to the applicant who nominated nobody', JSON.stringify(since(n)) === JSON.stringify(['applied@uni.cz | One request about V4AIR']), since(n));
  n = t.sent.length; day(4);
  check('day 4: speaker resume 2', JSON.stringify(since(n)) === JSON.stringify(['spk@uni.cz | Your V4AIR speaker form is not finished']), since(n));
  t.book.tab('Nominees').rows[1][11] = 'accepted';
  n = t.sent.length; day(7);
  check('day 7: application resume 2 and the nominee reminder (also after accepting)', JSON.stringify(since(n)) === JSON.stringify(['draft2@uni.cz | Your V4AIR application is not sent yet', 'nick@uni.cz | Your V4AIR nomination']), since(n));
  n = t.sent.length; day(10); day(14); day(28);
  check('days 10, 14, 28: speaker resume 3, nominate reminders 2 and 3', JSON.stringify(since(n)) === JSON.stringify(['applied@uni.cz | One last request about V4AIR', 'applied@uni.cz | Who else should be at V4AIR?', 'spk@uni.cz | Last reminder: your V4AIR speaker form']), since(n));
  n = t.sent.length; day(29); day(30);
  check('nothing more after the chains end', t.sent.length === n, since(n));
  day(31);
  const nom = rows(t.book, 'Nominees'), card = JSON.parse(rows(t.book, 'Nominations')[0].answers_json).__cards;
  check('day 31: nominee details deleted, accepted keeps its decision, the applicant\'s nominee row stays', !nom[0].name && !nom[0].contact && nom[0].decision === 'accepted' && !nom[1].contact && nom[1].decision === 'expired' && nom[2].contact === 'applied@uni.cz' && !card[0].contact && !card[1].contact && card[2].contact === 'applied@uni.cz', nom.map(r => [r.name, r.contact, r.decision]));
  check('every email went at most once per form, kind and address', new Set(rows(t.book, 'Email log').map(r => r.form + r.kind + r.to)).size === rows(t.book, 'Email log').length, rows(t.book, 'Email log').map(r => r.kind + ' ' + r.to));
}

/* 9. the application deadline */
{
  const seed = book => {
    book.seed('Applications', HEAD, [
      ['V4A-DRF3-AAAA', 'h1', 'in progress', serial(new Date('2026-12-20T10:00:00Z')), serial(new Date('2026-12-20T10:00:00Z')), '', '', '{}', 'Eva', 'late@uni.cz']
    ]);
    book.seed('Speakers', HEAD, []);
  };
  const clock = { now: Date.parse('2027-01-06T10:00:00Z') };
  const t = setup({ EMAIL_ENABLED: 'true', MAIL_MODE: 'live', UNSUB_SECRET: 's', MAIL_START: '2026-12-01T00:00:00Z' }, { seed, clock });
  t.run(); t.run();
  check('5 January on: no application resume 2, resume 1 went', t.sent.length === 1 && /is saved/.test(t.sent[0].subject), t.sent.map(m => m.subject));
  clock.now = Date.parse('2027-01-08T09:30:00Z'); t.run();
  check('8 January: deadline reminder, one week', t.sent.length === 2 && /close on 15 January/.test(t.sent[1].subject), t.sent.map(m => m.subject));
  clock.now = Date.parse('2027-01-14T09:30:00Z'); t.run();
  check('14 January: deadline reminder, last day', t.sent.length === 3 && /close tomorrow/.test(t.sent[2].subject), t.sent.map(m => m.subject));
  clock.now = Date.parse('2027-01-16T08:00:00Z');
  t.book.tab('Applications').rows.push(['V4A-SUB3-AAAA', 'h2', 'submitted', serial(new Date('2027-01-15T20:00:00Z')), serial(new Date('2027-01-15T20:00:00Z')), serial(new Date('2027-01-15T22:30:00Z')), '', '{}', 'Ivo', 'last@uni.cz']);
  t.run(); t.run();
  check('after 15 January: only the confirmation of a form sent at the last minute', t.sent.length === 4 && /We received your V4AIR application V4A-SUB3/.test(t.sent[3].subject), t.sent.map(m => m.subject));
}

/* 10. dry run lists the nominee details that would be deleted, and deletes nothing */
{
  const T0 = Date.UTC(2026, 10, 20, 9, 0), at = d => serial(new Date(T0 + d * DAY));
  const seed = book => {
    book.seed('Applications', HEAD, []); book.seed('Speakers', HEAD, []);
    book.seed('Nominations', ['response_id', 'status', 'answers_json'], [['V4N-OLD1-AAAA', 'submitted', JSON.stringify({ __cards: [{ name: 'Old', contact: 'old-nominee@uni.cz' }] })]]);
    book.seed('Nominees', ['nomination_id', 'n', 'name', 'contact', 'nominator_name', 'saved_at'], [['V4N-OLD1-AAAA', 1, 'Old', 'old-nominee@uni.cz', 'Ana', at(-40)]]);
  };
  const t = setup({}, { seed, clock: { now: T0 } });
  t.run();
  const p = rows(t.book, 'Mail preview').filter(r => r.would_send === 'delete details');
  check('dry: the 30-day deletion is previewed, nothing deleted', p.length === 1 && rows(t.book, 'Nominees')[0].contact === 'old-nominee@uni.cz', p);
}
console.log(fails ? fails + ' failed' : 'all passed');
process.exit(fails ? 1 : 0);
