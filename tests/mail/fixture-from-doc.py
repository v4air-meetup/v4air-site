"""Refresh tests/mail/fixtures/emails-doc.json from the Google Doc "V4AIR form emails" (the texts the mailer reads at run time).
The mail test runs on this copy, so it needs no Google access. Run it when the Doc's layout changes:
  python3 tests/mail/fixture-from-doc.py        (needs a Docs token in ~/.claude_google_token.json or GOOGLE_TOKEN_FILE)"""
import json, os
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

DOC_ID = '1EkQNlVa_DGdI5lYP727lhxpAr0XhKTinMEud09Y_g9g'
tok = json.load(open(os.environ.get('GOOGLE_TOKEN_FILE', os.path.expanduser('~/.claude_google_token.json'))))
doc = build('docs', 'v1', credentials=Credentials.from_authorized_user_info(tok), cache_discovery=False).documents().get(documentId=DOC_ID, includeTabsContent=True).execute()
text = lambda c: ''.join(r.get('textRun', {}).get('content', '') for el in c if 'paragraph' in el for r in el['paragraph']['elements'])
def tabs(ts):
    return [{'title': t['tabProperties']['title'], 'text': text(t['documentTab']['body']['content']), 'children': tabs(t.get('childTabs', []))} for t in ts]
out = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'fixtures', 'emails-doc.json')
json.dump(tabs(doc['tabs']), open(out, 'w', encoding='utf-8'), indent=1, ensure_ascii=False)
print('wrote', out)
