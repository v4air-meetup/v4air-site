/* Accept-nomination test: the real Worker (backend-worker) on the in-memory Sheet (forms/fake-google.mjs), then the page
   /accept-nomination/ in Chrome with its Worker calls answered by the same Worker. Nothing real is touched.
   Checks: a bad or foreign token is refused; "yes" needs the privacy notice and records consent; "no" wipes the nominee
   in Nominees and in the nomination's answers_json and adds the address to Unsubscribed; unsubscribing a nominee's address
   does the same; a wiped nominee reads as declined; an unknown nominee reads as gone; the page sends Yes only after the tick
   and then opens the application form.
   Run: cd tests && npm run nominee   (CHROME=path uses a local Chrome) */
import { createHmac, generateKeyPairSync } from 'node:crypto';
import { readFileSync, existsSync, statSync } from 'node:fs';
import http from 'node:http';
import path from 'node:path';
import { chromium } from 'playwright';
import { Book, restHandler } from '../forms/fake-google.mjs';

const ROOT = path.resolve(new URL('../..', import.meta.url).pathname), PUB = path.join(ROOT, 'public');
const HEAD = JSON.parse(readFileSync(new URL('../forms/fixtures/sheet-headers.json', import.meta.url)));
const SECRET = 'test-secret', NID = 'V4N-TEST-2345';
const b64url = b => Buffer.from(b).toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
const tok = msg => b64url(createHmac('sha256', SECRET).update(msg).digest());
const link = n => ({ id: `${NID}.${n}`, t: tok(`nomination:${NID}.${n}`) });

const book = new Book();
const cards = [1, 2, 3].map(n => ({ type: 'researcher', name: `Nominee ${n}`, contact: `nominee${n}@example.org`, affiliation: `Uni ${n}`, why: `Reason ${n}` }));
const nomRow = HEAD.Nominations.map(h => ({ response_id: NID, status: 'submitted', answers_json: JSON.stringify({ __cards: cards }), 'B.name': 'Alice Nominator', 'B.email': 'alice@example.org' })[h] ?? '');
const nomineeRows = cards.map((c, i) => HEAD.Nominees.map(h => ({ nomination_id: NID, nominator_name: 'Alice Nominator', nominator_email: 'alice@example.org', n: i + 1, type: c.type, name: c.name, contact: c.contact, affiliation: c.affiliation, why: c.why, saved_at: 46340 })[h] ?? ''));
book.seed('Nominations', HEAD.Nominations, [nomRow]);
book.seed('Nominees', HEAD.Nominees, nomineeRows);
book.seed('Applications', HEAD.Applications, []);

const google = restHandler(book), realFetch = globalThis.fetch;
globalThis.fetch = (u, i) => (/^https:\/\/(sheets|oauth2)\.googleapis\.com\//.test(String(u)) ? google(String(u), i) : realFetch(u, i));
const { privateKey } = generateKeyPairSync('rsa', { modulusLength: 2048 });
const env = { SHEET_ID: 'test', ALLOWED_ORIGINS: 'http://127.0.0.1:4320', UNSUB_SECRET: SECRET, GOOGLE_SA_KEY: JSON.stringify({ client_email: 'ci@test', private_key: privateKey.export({ type: 'pkcs8', format: 'pem' }) }) };
const worker = (await import(path.join(ROOT, 'backend-worker/src/index.js'))).default;
const post = async body => (await worker.fetch(new Request('https://w.test/', { method: 'POST', body: JSON.stringify(body) }), env)).json();

let failed = 0;
const check = (name, ok, extra) => { console.log((ok ? 'PASS ' : 'FAIL ') + name + (ok || extra === undefined ? '' : '  ' + JSON.stringify(extra))); if (!ok) failed++; };
const tab = t => { const x = book.tab(t); return x.rows.slice(1).map(r => Object.fromEntries(x.rows[0].map((h, i) => [h, r[i] ?? '']))); };
const nominee = n => tab('Nominees').find(r => r.nomination_id === NID && String(r.n) === String(n));
const card = n => JSON.parse(tab('Nominations')[0].answers_json).__cards[n - 1];
const unsub = () => tab('Unsubscribed').map(r => r.email + '/' + r.source);

/* ---------- Worker ---------- */
let r = await post({ action: 'nomination', id: `${NID}.1`, t: tok(`nomination:${NID}.2`) });
check('token of another nominee is refused', r.error === 'bad_link', r);
r = await post({ action: 'nomination', id: `${NID}.1`, t: 'x'.repeat(43) });
check('made-up token is refused', r.error === 'bad_link', r);
r = await post({ action: 'nomination', ...link(1) });
check('open link shows the nominator and state open', r.ok && r.state === 'open' && r.nominator === 'Alice Nominator' && r.name === 'Nominee 1', r);
r = await post({ action: 'nomination', ...link(1), decision: 'yes' });
check('yes without the privacy notice is refused', r.error === 'notice' && !nominee(1).decision, r);
r = await post({ action: 'nomination', ...link(1), decision: 'yes', notice: 'privacy-v1' });
const n1 = nominee(1);
check('yes records consent (decision, date, notice version) and keeps the details', r.state === 'accepted' && n1.decision === 'accepted' && n1.decision_notice === 'privacy-v1' && typeof n1.decided_at === 'number' && n1.name === 'Nominee 1', n1);
r = await post({ action: 'nomination', ...link(1) });
check('an accepted nomination reads as accepted', r.state === 'accepted', r);
r = await post({ action: 'nomination', ...link(2), decision: 'no' });
const n2 = nominee(2);
check('no wipes name, contact, affiliation and reason in Nominees', r.state === 'declined' && n2.decision === 'declined' && !n2.name && !n2.contact && !n2.affiliation && !n2.why, n2);
check('no wipes the card in the nomination, other cards untouched', !card(2).name && !card(2).contact && !card(2).why && card(1).name === 'Nominee 1' && card(3).name === 'Nominee 3', [card(1), card(2), card(3)]);
check('no puts the address on the do-not-write list', unsub().includes('nominee2@example.org/nominee-declined'), unsub());
r = await post({ action: 'nomination', ...link(2), decision: 'yes', notice: 'privacy-v1' });
check('a declined nominee cannot be revived from the link', r.state === 'declined' && nominee(2).decision === 'declined', r);
const e3 = b64url(Buffer.from('nominee3@example.org')), t3 = tok('nominee3@example.org');
r = await post({ action: 'unsubscribe', e: e3, t: t3 });
check('unsubscribing a nominee wipes them like a no', r.ok && nominee(3).decision === 'declined' && !nominee(3).name && !card(3).name, nominee(3));
r = await post({ action: 'nomination', id: `${NID}.4`, t: tok(`nomination:${NID}.4`) });
check('an unknown nominee reads as gone', r.ok && r.state === 'gone', r);
check('other columns of the nominee rows are unchanged', nominee(2).nominator_name === 'Alice Nominator' && nominee(2).saved_at === 46340 && nominee(1).affiliation === 'Uni 1');

/* ---------- page in Chrome ---------- */
book.seed('Nominees', book.tab('Nominees').rows[0], nomineeRows.slice(0, 1).map(r => r.slice())); /* a fresh, undecided nominee 1 */
const types = { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css', '.woff2': 'font/woff2', '.svg': 'image/svg+xml' };
http.createServer((req, res) => {
  let f = path.join(PUB, decodeURIComponent(new URL(req.url, 'http://x').pathname));
  if (existsSync(f) && statSync(f).isDirectory()) f = path.join(f, 'index.html');
  if (!existsSync(f)) { res.writeHead(404); res.end(); return; }
  res.writeHead(200, { 'Content-Type': types[path.extname(f)] || 'application/octet-stream' }); res.end(readFileSync(f));
}).listen(4320, '127.0.0.1');
const browser = await chromium.launch(process.env.CHROME ? { executablePath: process.env.CHROME } : {});
const page = await browser.newPage({ viewport: { width: 390, height: 800 } });
const errors = []; page.on('pageerror', e => errors.push(String(e)));
await page.route('https://v4air-forms.v4air.workers.dev/**', async route => {
  const w = await worker.fetch(new Request('https://w.test/', { method: 'POST', body: route.request().postData() }), env);
  route.fulfill({ status: w.status, contentType: 'application/json', body: await w.text() });
});
const l = link(1);
await page.goto(`http://127.0.0.1:4320/accept-nomination/?id=${encodeURIComponent(l.id)}&t=${l.t}`);
await page.waitForSelector('#open.on');
check('page shows who nominated them', (await page.textContent('#nominator')) === 'Alice Nominator');
check('Yes is disabled until the privacy notice is ticked', await page.isDisabled('#yes'));
await page.check('#notice');
await Promise.all([page.waitForURL(/\/apply\/\?utm_source=nomination/), page.click('#yes')]);
check('Yes records consent and opens the application form', nominee(1).decision === 'accepted' && nominee(1).decision_notice === 'privacy-v1', nominee(1));
await page.goto('http://127.0.0.1:4320/accept-nomination/');
check('a link without id and token shows "not valid"', await page.isVisible('#bad'));
check('no JavaScript errors on the page', !errors.filter(e => !/apply/.test(e)).length, errors);
await browser.close();
console.log(failed ? `\n${failed} check(s) failed` : '\nall checks passed');
process.exit(failed ? 1 : 0);
