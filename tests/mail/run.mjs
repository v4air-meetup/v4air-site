/* The forms mailer (sendMail in backend/Code.gs, texts in backend/Emails.gs) on imitated Apps Script services.
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

function setup(props, { resend = 200, unsub = [] } = {}) {
  const book = new Book();
  book.seed('Applications', HEAD, [
    ['V4A-OLD1-AAAA', 'h1', 'submitted', ago(31), ago(30), ago(30), '', '{}', 'Olga', 'old@uni.cz'],
    ['V4A-NEW1-AAAA', 'h2', 'submitted', ago(0.1), ago(0.05), ago(0.05), '', '{}', 'Nina', 'new@uni.cz'],
    ['V4A-DRF1-AAAA', 'h3', 'in progress', ago(12), ago(10), '', '', '{}', 'Dan', 'draft@uni.cz'],
    ['V4A-TST1-AAAA', 'h4', 'submitted', ago(1), ago(1), ago(1), '', '{}', 'Test', 'x@example.com'],
    ['V4A-TST2-AAAA', 'h5', 'submitted', ago(1), ago(1), ago(1), '', '{}', 'John Doe', 'a@a.aa'],
    ['V4A-DUP1-AAAA', 'h6', 'in progress', ago(5), ago(5), '', '', '{}', 'Nina', 'NEW@uni.cz'],
    ['V4A-TWC1-AAAA', 'h7', 'submitted', ago(0.2), ago(0.2), ago(0.2), '', '{}', 'Tom', 'twice@uni.cz'],
    ['V4A-TWC2-AAAA', 'h8', 'submitted', ago(0.1), ago(0.1), ago(0.1), '', '{}', 'Tom', 'twice@uni.cz']
  ]);
  book.seed('Speakers', HEAD, []);
  book.seed('Unsubscribed', ['email', 'unsubscribed_at', 'source'], [['a@a.aa', '', 'test'], ...unsub.map(e => [e, '', 'page'])]);
  const sent = [], sha = (alg, s) => [...createHash(alg).update(s, 'utf8').digest()].map(b => (b > 127 ? b - 256 : b));
  const ctx = vm.createContext(Object.assign(appsScript(book, { props }), {
    console: { log() {}, error() {} },
    Utilities: {
      DigestAlgorithm: { SHA_256: 'sha256' }, Charset: { UTF_8: 'utf8' }, getUuid: () => randomUUID(), computeDigest: (a, s) => sha(a, s),
      newBlob: s => ({ getBytes: () => [...Buffer.from(s)] }),
      base64EncodeWebSafe: x => Buffer.from(Array.isArray(x) ? x.map(b => b & 255) : x).toString('base64').replace(/\+/g, '-').replace(/\//g, '_'),
      computeHmacSha256Signature: (v, k) => [...createHmac('sha256', k).update(v).digest()].map(b => (b > 127 ? b - 256 : b)),
      formatDate: (d, tz, pattern) => prague(d, pattern)
    },
    DriveApp: { getRootFolder: () => ({ getFilesByName: () => ({ hasNext: () => false }) }) },
    MailApp: { getRemainingDailyQuota: () => 100 },
    GmailApp: { sendEmail: (to, subject, body) => sent.push({ via: 'gmail', to: to.split(','), subject, body }) },
    UrlFetchApp: {
      fetch: (url, o) => { const m = JSON.parse(o.payload); sent.push({ via: 'resend', ...m, body: m.text });
        const code = typeof resend === 'function' ? resend() : resend; return { getResponseCode: () => code, getContentText: () => (code === 200 ? '{"id":"re_1"}' : 'error') }; }
    }
  }));
  vm.runInContext(readFileSync(path.join(ROOT, 'backend/Emails.gs'), 'utf8'), ctx);
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
  check('dry: old draft waits 24 h after go-live; duplicate draft skipped', /not yet/.test(why('V4A-DRF1-AAAA')) && /already submitted/.test(why('V4A-DUP1-AAAA')), [why('V4A-DRF1-AAAA'), why('V4A-DUP1-AAAA')]);
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
console.log(fails ? fails + ' failed' : 'all passed');
process.exit(fails ? 1 : 0);
