#!/usr/bin/env python3
"""
Download MyGuide-progress.xlsx from a SharePoint share link and regenerate
progress.json for the dashboard.

Reads the same two sheets the dashboard's "Import from Excel" button reads,
using the same coercion rules, so the automated output is byte-identical to
doing it by hand.

  Progress sheet  60 rows: Segment, Audience, Language, Chapter,
                           Progress %, Feedback 1, Feedback 2, Validated
  Slides sheet    20 rows: Segment, Audience, Chapter, Slides

Environment:
  XLSX_URL        SharePoint "Anyone with the link" URL   (required)
  OUT             output path       (default MyGuide-webflow/progress.json)
  ALLOW_PARTIAL   set to 1 to publish even when rows are missing
"""
import os, sys, json, re, io, datetime, urllib.request, urllib.error

OUT = os.environ.get('OUT', 'MyGuide-webflow/progress.json')
URL = os.environ.get('XLSX_URL', '').strip()
ALLOW_PARTIAL = os.environ.get('ALLOW_PARTIAL', '') == '1'


def die(msg, code=1):
    """Fail loudly and visibly. The ::error:: annotation puts the reason on the
    run's summary page, so nobody has to dig through step logs to find it."""
    print('ERROR: ' + msg, file=sys.stderr)
    one_line = ' '.join(msg.split())
    print('::error::' + one_line)
    summary = os.environ.get('GITHUB_STEP_SUMMARY')
    if summary:
        try:
            with open(summary, 'a', encoding='utf-8') as f:
                f.write('\n### Sync failed\n\n' + msg + '\n')
        except OSError:
            pass
    sys.exit(code)


def note(msg):
    print(msg)
    summary = os.environ.get('GITHUB_STEP_SUMMARY')
    if summary:
        try:
            with open(summary, 'a', encoding='utf-8') as f:
                f.write(msg + '\n\n')
        except OSError:
            pass


def download(url):
    """SharePoint share links need download=1 to return the file itself
    rather than the Office web viewer wrapped in HTML."""
    if 'download=1' not in url:
        url += ('&' if '?' in url else '?') + 'download=1'
    req = urllib.request.Request(url, headers={
        'User-Agent': 'Mozilla/5.0 (compatible; MyGuide-dashboard-sync)'})
    try:
        with urllib.request.urlopen(req, timeout=120) as r:
            data = r.read()
    except urllib.error.HTTPError as e:
        die('SharePoint returned HTTP %s. The share link may have expired, or it '
            'may not be an "Anyone with the link" link.' % e.code)
    except Exception as e:
        die('Could not download the workbook: %s' % e)
    if data[:2] != b'PK':
        head = data[:400].decode('utf-8', 'replace')
        die('The URL did not return an .xlsx file. It usually means the link asks '
            'for a sign-in, so SharePoint sent an HTML login page instead.\n'
            'Recreate the link with Share -> Anyone with the link -> Can view.\n'
            'First bytes received:\n' + head)
    return data


def norm(v):
    return re.sub(r'[^a-z0-9]', '', str('' if v is None else v).lower())


def find_col(headers, *wanted):
    m = {norm(h): h for h in headers if h is not None}
    for w in wanted:
        n = norm(w)
        if n in m:
            return m[n]
        for k in m:
            if k.startswith(n):
                return m[k]
    return None


def to_pct(v):
    if v is None or v == '':
        return None
    if isinstance(v, bool):
        return None
    try:
        n = float(v) if not isinstance(v, str) else float(v.replace('%', '').replace(',', '.').strip())
    except ValueError:
        return None
    if 0 < n <= 1:
        n *= 100                      # a cell formatted as 80% arrives as 0.8
    n = round(n / 10.0) * 10
    return int(max(0, min(100, n)))


def to_fb(v):
    if v is None or v == '':
        return 0
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return int(v) if int(v) in (1, 2) else 0
    n = norm(v)
    if 'process' in n or n == 'done' or n == '2' or 'verwerkt' in n:
        return 2
    if 'review' in n or n == '1' or 'nazicht' in n or 'lopend' in n:
        return 1
    return 0


def to_bool(v):
    if v is True:
        return True
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return int(v) == 1
    return norm(v) in ('yes', 'y', 'true', 'x', '1', 'ja', 'oui')


def sheet_rows(ws):
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return [], []
    headers = [('' if c is None else str(c).strip()) for c in rows[0]]
    out = []
    for r in rows[1:]:
        d = {}
        any_val = False
        for h, val in zip(headers, r):
            if not h:
                continue
            d[h] = '' if val is None else val
            if val not in (None, ''):
                any_val = True
        if any_val:
            out.append(d)
    return headers, out


def pick_sheet(wb, name, needed):
    for s in wb.sheetnames:
        if norm(s) == norm(name):
            return wb[s]
    for s in wb.sheetnames:
        headers, _ = sheet_rows(wb[s])
        if headers and all(find_col(headers, c) for c in needed):
            return wb[s]
    return None


def main():
    if not URL:
        die('XLSX_URL is not set. Add the SharePoint share link as a repository '
            'secret named XLSX_URL (Settings -> Secrets and variables -> Actions).')

    try:
        from openpyxl import load_workbook
    except ImportError:
        die('openpyxl is not installed.')

    # Existing progress.json is the template: it carries chapter titles, the
    # weighting settings and the deck structure, none of which come from Excel.
    if not os.path.exists(OUT):
        die('%s not found. It is the template for structure and settings; commit '
            'it once before enabling this workflow.' % OUT)
    with open(OUT, encoding='utf-8') as f:
        base = json.load(f)

    SEG   = base['segments']
    AUD   = base['audiences']
    LANG  = base['languages']
    NCH   = len(base['chapters'])
    expected_cells  = len(SEG) * len(AUD) * len(LANG) * NCH
    expected_slides = len(SEG) * len(AUD) * NCH

    wb = load_workbook(io.BytesIO(download(URL)), data_only=True, read_only=True)

    pws = pick_sheet(wb, 'Progress', ['segment', 'audience', 'language', 'chapter'])
    if pws is None:
        die('No Progress sheet found. Sheets in the workbook: %s' % ', '.join(wb.sheetnames))

    def match(v, options):
        n = norm(v)
        for o in options:
            if norm(o) == n:
                return o
        return None

    headers, rows = sheet_rows(pws)
    cSeg  = find_col(headers, 'segment')
    cAud  = find_col(headers, 'audience')
    cLang = find_col(headers, 'language')
    cCh   = find_col(headers, 'chapter')
    cP    = find_col(headers, 'progress', 'progress%', 'percent')
    cF1   = find_col(headers, 'feedback1', 'fb1')
    cF2   = find_col(headers, 'feedback2', 'fb2')
    cV    = find_col(headers, 'validated', 'validation')

    cells = {}
    unmatched = []
    for ix, row in enumerate(rows):
        seg  = match(row.get(cSeg), SEG)
        aud  = match(row.get(cAud), AUD)
        lang = match(row.get(cLang), LANG)
        try:
            ch = int(row.get(cCh))
        except (TypeError, ValueError):
            ch = 0
        if not (seg and aud and lang and 1 <= ch <= NCH):
            looks_like_data = str(row.get(cSeg, '')).strip() != '' and (
                str(row.get(cLang, '')).strip() != '' or str(row.get(cCh, '')).strip() != '')
            if looks_like_data and len(unmatched) < 10:
                unmatched.append('row %d: %s / %s / %s / %s' % (
                    ix + 2, row.get(cSeg), row.get(cAud), row.get(cLang), row.get(cCh)))
            continue
        key = '%s|%s|%s' % (aud, lang, seg)
        cells.setdefault(key, [None] * NCH)
        p = to_pct(row.get(cP)) if cP else None
        cells[key][ch - 1] = {
            'p':  0 if p is None else p,
            'f1': to_fb(row.get(cF1)) if cF1 else 0,
            'f2': to_fb(row.get(cF2)) if cF2 else 0,
            'v':  to_bool(row.get(cV)) if cV else False,
        }

    n_cells = sum(1 for v in cells.values() for c in v if c is not None)

    slides = {}
    n_slides = 0
    sws = pick_sheet(wb, 'Slides', ['segment', 'audience', 'chapter', 'slides'])
    if sws is not None:
        sheaders, srows = sheet_rows(sws)
        sSeg = find_col(sheaders, 'segment')
        sAud = find_col(sheaders, 'audience')
        sCh  = find_col(sheaders, 'chapter')
        sSl  = find_col(sheaders, 'slides', 'slide')
        for row in srows:
            seg = match(row.get(sSeg), SEG)
            aud = match(row.get(sAud), AUD)
            try:
                ch = int(row.get(sCh))
            except (TypeError, ValueError):
                ch = 0
            if not (seg and aud and 1 <= ch <= NCH):
                continue
            try:
                n = int(float(row.get(sSl) or 0))
            except (TypeError, ValueError):
                n = 0
            slides.setdefault('%s|%s' % (aud, seg), [0] * NCH)[ch - 1] = max(0, min(999, n))
            n_slides += 1

    note('Matched %d/%d progress cells and %d/%d slide values.'
         % (n_cells, expected_cells, n_slides, expected_slides))
    for u in unmatched:
        print('  unmatched %s' % u)

    if n_cells == 0:
        die('No row matched a known deck. Check that Segment reads %s, Audience reads %s '
            'and Language reads %s.' % ('/'.join(SEG), '/'.join(AUD), '/'.join(LANG)))
    if n_cells < expected_cells and not ALLOW_PARTIAL:
        die('Only %d of %d progress cells matched. Refusing to publish, because the '
            'missing chapters would be written as 0%%. Fix the workbook, or set '
            'ALLOW_PARTIAL=1 if this is deliberate.' % (n_cells, expected_cells))

    # fill gaps and write
    for seg in SEG:
        for aud in AUD:
            for lang in LANG:
                key = '%s|%s|%s' % (aud, lang, seg)
                cells.setdefault(key, [None] * NCH)
                for i in range(NCH):
                    if cells[key][i] is None:
                        cells[key][i] = {'p': 0, 'f1': 0, 'f2': 0, 'v': False}
            if n_slides:
                slides.setdefault('%s|%s' % (aud, seg), [0] * NCH)

    base['cells'] = cells
    if n_slides:
        base['slides'] = slides
    base['locked'] = True
    base['updated'] = datetime.date.today().isoformat()

    new = json.dumps(base, indent=2, ensure_ascii=False) + '\n'
    old = open(OUT, encoding='utf-8').read() if os.path.exists(OUT) else ''

    # Ignore a date-only difference so an unchanged workbook does not produce a
    # commit every fifteen minutes.
    def strip_date(s):
        return re.sub(r'"updated":\s*"[^"]*"', '"updated":""', s)

    if strip_date(new) == strip_date(old):
        note('No change in the figures — leaving progress.json alone.')
        return

    with open(OUT, 'w', encoding='utf-8') as f:
        f.write(new)
    note('Wrote %s' % OUT)


if __name__ == '__main__':
    main()
