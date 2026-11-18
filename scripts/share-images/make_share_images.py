"""Personal share images for V4AIR applicants.

Each submitted applicant gets the two invitation visuals (portrait for posts, wide banner for
newsletters) with their own QR code. The QR code points to v4air.eu with UTM parameters whose
utm_content is the applicant's code, so visits and applications from a shared image can be traced
back to the person who shared it (Google Analytics, utm_content).

    code = application ID without "V4A-", lower case   (V4A-A7HL-4P34 -> a7hl-4p34)
    out/<code>.jpg          portrait, QR v4air.eu/?utm_source=wom&utm_medium=share&utm_content=<code>
    out/<code>-banner.jpg   banner,   QR v4air.eu/?utm_source=wom&utm_medium=forward&utm_content=<code>

The applicant list comes from the Twenty CRM REST API (env TWENTY_API_KEY). The deploy workflow
runs this after `npm run build` with --out dist/share, so the images exist only in the deployed
site and the repository never holds the list of codes. Without a key it does nothing.

Run locally:  python scripts/share-images/make_share_images.py --out out/share
              python scripts/share-images/make_share_images.py --codes a7hl-4p34 --out out/share
"""
import argparse, json, os, re, sys, urllib.parse, urllib.request
from pathlib import Path

import qrcode
from PIL import Image

HERE = Path(__file__).parent
API = os.environ.get('TWENTY_URL', 'https://chlup-crm.dyn.cloud.e-infra.cz').rstrip('/')
CODE = re.compile(r'^[a-z0-9]{4}-[a-z0-9]{4}$')

# White QR box in each base image, measured once (left, top, right, bottom in pixels).
TEMPLATES = {
    '': {'file': 'portrait.jpg', 'box': (1118, 1307, 1266, 1454), 'medium': 'share', 'campaign': 'email3'},
    '-banner': {'file': 'banner.jpg', 'box': (2296, 747, 2452, 903), 'medium': 'forward', 'campaign': 'email6'},
}
CORNER = 6  # keep the box's rounded corners: clear only inside this inset


def link(code, medium, campaign):
    """Address encoded in the QR code. utm_campaign is left out on purpose: utm_medium already
    tells the two images apart, and the shorter address keeps the code at 33 x 33 modules, which
    scans at full and half size (tested with ZXing on all applicants, 18.11.2026)."""
    return f'https://v4air.eu/?utm_source=wom&utm_medium={medium}&utm_content={code}'


def qr_matrix(data):
    qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_L, border=0)
    qr.add_data(data)
    qr.make(fit=True)
    return qr.get_matrix()


def render(code, out_dir):
    for suffix, t in TEMPLATES.items():
        base = Image.open(HERE / 'base' / t['file']).convert('RGB')
        l, top, r, b = t['box']
        base.paste((255, 255, 255), (l + CORNER, top + CORNER, r - CORNER, b - CORNER))  # clear the old code
        m = qr_matrix(link(code, t['medium'], t['campaign']))
        n = len(m)
        # whole pixels per module, leaving at least 2 modules of white quiet zone inside the box
        px = min(r - l, b - top) // (n + 4)
        side = n * px
        x0 = l + (r - l - side) // 2
        y0 = top + (b - top - side) // 2
        qr = Image.new('RGB', (side, side), 'white')
        for i, row in enumerate(m):
            for j, dark in enumerate(row):
                if dark:
                    qr.paste((0, 0, 0), (j * px, i * px, (j + 1) * px, (i + 1) * px))
        base.paste(qr, (x0, y0))
        base.save(out_dir / f'{code}{suffix}.jpg', quality=92, subsampling=0, optimize=True)


def fetch_codes(key):
    codes, cursor = [], None
    while True:
        params = {'filter': 'status[eq]:SUBMITTED', 'limit': 60}
        if cursor:
            params['starting_after'] = cursor
        req = urllib.request.Request(f'{API}/rest/applications?' + urllib.parse.urlencode(params),
                                     headers={'Authorization': 'Bearer ' + key, 'Accept': 'application/json'})
        with urllib.request.urlopen(req, timeout=60) as resp:
            d = json.load(resp)
        for a in d['data']['applications']:
            code = str(a.get('responseId') or '').strip().lower().removeprefix('v4a-')
            if CODE.match(code):
                codes.append(code)
        info = d.get('pageInfo') or {}
        if not info.get('hasNextPage'):
            return sorted(set(codes))
        cursor = info['endCursor']


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--out', default='out/share')
    ap.add_argument('--codes', nargs='*', help='render these codes instead of reading the CRM')
    a = ap.parse_args()
    if a.codes:
        codes = [c.lower() for c in a.codes if CODE.match(c.lower())]
    else:
        key = os.environ.get('TWENTY_API_KEY', '').strip()
        if not key:
            print('share images: TWENTY_API_KEY not set, skipping')
            return 0
        codes = fetch_codes(key)
    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    for c in codes:
        render(c, out)
    print(f'share images: {len(codes)} applicants, {len(codes) * len(TEMPLATES)} images in {out}')
    return 0


if __name__ == '__main__':
    sys.exit(main())
