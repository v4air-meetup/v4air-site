"""Take the email texts from the Doc "V4AIR form emails" into the forms mailer, on purpose.

The mailer sends from its own copy of the texts. An export from Typecase only changes the Doc; this script is the step
that puts the Doc's texts into use:
  1. reads the Doc the way the mailer does and shows what would change against the texts in use (Sheet tab "Email texts");
  2. after you confirm, writes {"EMAILS_UPTAKE": "<time>"} into the mailer's settings folder (Petr's Drive);
  3. waits for the mailer's next run (at most about 6 minutes) and prints its answer: taken in, or refused with the reasons
     (then the texts in use stay as they were).

  python3 scripts/emails-uptake.py           # show the changes, ask, take in
  python3 scripts/emails-uptake.py --diff    # only show the changes
  python3 scripts/emails-uptake.py --yes     # no question

Needs a Google token with Docs, Drive and Sheets access in ~/.claude_google_token.json (or GOOGLE_TOKEN_FILE).
Rules of the emails: Doc "V4AIR form emails: how sending works"; the Doc layout: docs/form-emails.md."""
import datetime, difflib, io, json, os, re, sys, time
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

DOC_ID = os.environ.get('EMAILS_DOC_ID', '1EkQNlVa_DGdI5lYP727lhxpAr0XhKTinMEud09Y_g9g')
SHEET_ID = '1dDRwBAy8YIQUQIZDiYjlbabDyJ3ZXbUOfl6sznkQYdc'
SETTINGS_FOLDER = os.environ.get('MAIL_SETTINGS_FOLDER', '1-RhuGZtM8uVPpK8qD-kH3xjNX8czK2Mq')
SETTINGS_FILE = 'v4air-script-secrets.json'

creds = Credentials.from_authorized_user_info(json.load(open(os.environ.get('GOOGLE_TOKEN_FILE', os.path.expanduser('~/.claude_google_token.json')))))
docs = build('docs', 'v1', credentials=creds, cache_discovery=False)
drive = build('drive', 'v3', credentials=creds, cache_discovery=False)
sheets = build('sheets', 'v4', credentials=creds, cache_discovery=False).spreadsheets()

def key(s):
    return re.sub(r'[^a-z0-9]+', '', s.lower())

def read_doc():
    """{tab key: (tab, subject, body)}, read like backend/Code.gs parseEmail"""
    d = docs.documents().get(documentId=DOC_ID, includeTabsContent=True).execute()
    out = {}
    def walk(ts):
        for t in ts:
            paras = [re.sub(r'[ \t]+$', '', ''.join(r.get('textRun', {}).get('content', '') for r in el['paragraph']['elements']).rstrip('\n').replace('\x0b', '\n'), flags=re.M)
                     for el in t['documentTab']['body']['content'] if 'paragraph' in el]
            i = next((n for n, p in enumerate(paras) if p.startswith('Subject:')), None)
            if i is not None:
                first = paras[i].split('\n'); rest = '\n'.join(first[1:])
                after = ([rest] if rest.strip() else []) + paras[i + 1:]
                while after and not after[0].strip(): after.pop(0)
                while after and not after[-1].strip(): after.pop()
                body = '\n'.join(after) if any(not p.strip() for p in after) else '\n\n'.join(p.strip('\n') for p in after)
                title = t['tabProperties']['title']
                out[key(title)] = (title, re.sub(r'^Subject:\s*', '', first[0]).strip(), body + '\n', any(p.strip() for p in paras[:i]))
            walk(t.get('childTabs', []))
    walk(d['tabs'])
    return out

def in_use():
    try:
        v = sheets.values().get(spreadsheetId=SHEET_ID, range="'Email texts'").execute().get('values', [])
    except Exception:
        return {}, None
    rows = [r + [''] * (6 - len(r)) for r in v[1:]]
    return {key(r[3]): (r[3], r[4], r[5]) for r in rows[1:] if r[1] == 'in use'}, (rows[0] if rows else None)

def show_changes(doc, live):
    changed = 0
    for k in sorted(set(doc) | set(live)):
        if k not in live:
            print(f'NEW      {doc[k][0]}'); changed += 1
        elif k not in doc:
            print(f'GONE     {live[k][0]} (no tab in the Doc any more)'); changed += 1
        elif (doc[k][1], doc[k][2].rstrip('\n')) != (live[k][1], live[k][2].rstrip('\n')):
            changed += 1
            print(f'CHANGED  {doc[k][0]}')
            a = ['Subject: ' + live[k][1]] + live[k][2].rstrip('\n').split('\n')
            b = ['Subject: ' + doc[k][1]] + doc[k][2].rstrip('\n').split('\n')
            for line in list(difflib.unified_diff(a, b, lineterm='', n=0))[2:]:
                print('         ' + line)
        if k in doc and doc[k][3]:
            print(f'         ! text above the Subject line in "{doc[k][0]}": the mailer will refuse the uptake')
    print(f'{changed} of {len(set(doc) | set(live))} emails change' if live else f'{len(doc)} emails in the Doc; none in use yet')
    return changed

def ask_mailer(stamp):
    files = drive.files().list(q=f"'{SETTINGS_FOLDER}' in parents and name = '{SETTINGS_FILE}' and trashed = false", fields='files(id)').execute()['files']
    media = MediaIoBaseUpload(io.BytesIO(json.dumps({'EMAILS_UPTAKE': stamp}).encode()), mimetype='application/json')
    if files:
        drive.files().update(fileId=files[0]['id'], media_body=media).execute()
    else:
        drive.files().create(body={'name': SETTINGS_FILE, 'parents': [SETTINGS_FOLDER]}, media_body=media).execute()

if __name__ == '__main__':
    doc, (live, last) = read_doc(), in_use()
    changed = show_changes(doc, live)
    if '--diff' in sys.argv:
        sys.exit(0)
    if '--yes' not in sys.argv and input('Take these texts into use? [y/N] ').strip().lower() != 'y':
        sys.exit('nothing changed')
    stamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    ask_mailer(stamp)
    print(f'asked the mailer (request {stamp}); waiting for its next run ...')
    for _ in range(30):
        time.sleep(20)
        _, last = in_use()
        if last and last[0] == stamp:
            print(last[1])
            sys.exit(0 if last[1].startswith('taken in') else 1)
    sys.exit('no answer within 10 minutes: check the Sheet tab "Email texts" later')
