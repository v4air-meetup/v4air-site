/* Every form email once, filled with made-up data, for a person to read and sign off before the mailer goes live.
   The texts come through the real backend/Code.gs (emailText) from tests/mail/fixtures/emails-doc.json, so refresh that
   from the Doc first. No Sheet is read and nothing is sent unless --send is given.

     python3 tests/mail/fixture-from-doc.py                 # today's texts from the Doc "V4AIR form emails"
     node tests/mail/samples.mjs                             # list the 17 emails
     node tests/mail/samples.mjs --out samples.txt           # write all of them, in full, to a file
     node tests/mail/samples.mjs --send petr@example.org     # send each one there through Resend, from info@v4air.eu

   --send needs a Resend sending key in RESEND_API_KEY or in ~/.config/v4air-kit/resend_api_key. Subjects start with
   "[Sample n/17]". The personal, accept and unsubscribe links are made with a made-up secret and ids, so they look real
   but lead to "This link is not valid" or an empty form: they test the layout, not the pages. */
import { readFileSync, writeFileSync, existsSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import vm from 'node:vm';
import { createHmac } from 'node:crypto';

const ROOT = path.resolve(path.dirname(new URL(import.meta.url).pathname), '../..');
const args = process.argv.slice(2), opt = k => { const i = args.indexOf(k); return i < 0 ? null : args[i + 1]; };
const DOC = JSON.parse(readFileSync(new URL('./fixtures/emails-doc.json', import.meta.url)));
/* a tab body the way Apps Script hands it out: one child per paragraph (the Docs API text has one paragraph per line) */
const body = text => { const ps = String(text).split('\n'); if (ps[ps.length - 1] === '') ps.pop(); return { getNumChildren: () => ps.length, getChild: i => ({ getType: () => 'PARAGRAPH', asText: () => ({ getText: () => ps[i] }) }) }; };
const wrap = t => ({ getTitle: () => t.title, asDocumentTab: () => ({ getBody: () => body(t.text) }), getChildTabs: () => t.children.map(wrap) });
const props = { UNSUB_SECRET: 'sample-secret-not-the-real-one' };
const ctx = vm.createContext({
  console,
  PropertiesService: { getScriptProperties: () => ({ getProperty: k => props[k] ?? null }) },
  DocumentApp: { openById: () => ({ getTabs: () => DOC.map(wrap) }) },
  Utilities: {
    formatDate: (d, tz, p) => {
      const x = Object.fromEntries(new Intl.DateTimeFormat('en-GB', { timeZone: 'Europe/Prague', year: 'numeric', month: 'numeric', day: 'numeric' }).formatToParts(new Date(d)).map(q => [q.type, q.value]));
      return p.replace('yyyy', x.year).replace('M', String(+x.month)).replace('d', String(+x.day));
    },
    newBlob: s => ({ getBytes: () => [...Buffer.from(s)] }),
    base64EncodeWebSafe: x => Buffer.from(Array.isArray(x) ? x.map(b => b & 255) : x).toString('base64').replace(/\+/g, '-').replace(/\//g, '_'),
    computeHmacSha256Signature: (v, k) => [...createHmac('sha256', k).update(v).digest()].map(b => (b > 127 ? b - 256 : b))
  }
});
vm.runInContext(readFileSync(path.join(ROOT, 'backend/Code.gs'), 'utf8'), ctx);
const run = src => vm.runInContext(src, ctx);
const cfg = { site: 'https://v4air.eu/', secret: props.UNSUB_SECRET, emails: run("loadEmails('sample')"), textError: '' };
const FLOWS = run('FLOWS');

const day = s => new Date(s + 'T10:00:00+01:00');
const people = {
  A: { id: 'V4A-SAMP-LE01', name: 'Jana Nováková', to: 'jana.sample@example.org', sent: day('2026-11-18'), saved: day('2026-11-18'), started: day('2026-11-18') },
  S: { id: 'V4S-SAMP-LE02', name: 'Tomasz Kowalski', to: 'tomasz.sample@example.org', sent: day('2026-11-18'), saved: day('2026-11-18'), started: day('2026-11-18') },
  B: { id: 'V4N-SAMP-LE03', name: 'Jana Nováková', to: 'jana.sample@example.org', sent: day('2026-11-18'), saved: day('2026-11-18') },
  N: { id: 'V4N-SAMP-LE03.1', nominationId: 'V4N-SAMP-LE03', name: 'Eszter Nagy', to: 'eszter.sample@example.org', nominator: 'Jana Nováková', sent: day('2026-11-18'), saved: day('2026-11-18') }
};
const items = FLOWS.map(f => ({ flow: f, rec: people[f.form], late: false }));
/* the backlog versions: an application sent and a draft started weeks before the emails were switched on */
const old = { ...people.A, sent: day('2026-10-21'), saved: day('2026-10-21'), started: day('2026-10-20') };
items.splice(1, 0, { flow: FLOWS[0], rec: old, late: true });
items.splice(items.findIndex(it => it.flow.kind === 'resume' && it.flow.form === 'A') + 1, 0, { flow: FLOWS.find(f => f.form === 'A' && f.kind === 'resume'), rec: old, late: true });

ctx.__items = items; ctx.__cfg = cfg;
const out = items.map((it, i) => {
  ctx.__it = it;
  const problem = run('textProblem(__cfg, __it)');
  const link = run("mailLink(__cfg.site, __it.flow.form, __it.rec.id, 'sample-token')");
  const m = run('emailText(__cfg, __it, ' + JSON.stringify(link) + ', unsubscribeLink(__it.rec.to), acceptLink(__cfg, __it.rec.id))');
  const name = it.flow.tab + (it.late ? ' (late)' : '');
  return { n: i + 1, name, to: it.rec.to, subject: m.subject, body: m.body, problem };
});
const total = out.length;
for (const e of out) console.log(`${String(e.n).padStart(2)}/${total}  ${e.name.padEnd(34)} ${e.problem ? 'BLOCKED: ' + e.problem : e.subject}`);
if (opt('--out')) {
  writeFileSync(opt('--out'), out.map(e => `===== ${e.n}/${total} ${e.name}\nWould go to: ${e.to}\nSubject: ${e.subject}\n\n${e.body}`).join('\n\n'));
  console.log('written to ' + opt('--out'));
}
const to = opt('--send');
if (to) {
  const keyFile = path.join(os.homedir(), '.config/v4air-kit/resend_api_key');
  const key = process.env.RESEND_API_KEY || (existsSync(keyFile) ? readFileSync(keyFile, 'utf8').trim() : '');
  if (!key) { console.error('no Resend key'); process.exit(1); }
  if (out.some(e => e.problem)) { console.error('some texts are blocked; fix the Doc first'); process.exit(1); }
  for (const e of out) {
    const res = await fetch('https://api.resend.com/emails', {
      method: 'POST', headers: { Authorization: 'Bearer ' + key, 'Content-Type': 'application/json', 'User-Agent': 'v4air-mail-samples/1.0' },
      body: JSON.stringify({ from: 'V4 AI Researchers Meetup <info@v4air.eu>', to: [to], reply_to: 'info@v4air.eu', subject: `[Sample ${e.n}/${total}] ${e.subject}`, text: e.body })
    });
    console.log(`${e.n}/${total} ${res.status} ${(await res.text()).slice(0, 80)}`);
    await new Promise(r => setTimeout(r, 700)); /* Resend allows 2 requests a second */
  }
}
