# -*- coding: utf-8 -*-
"""
GST RATE MASTER — a dated, sourced list of GST rates on goods, kept as one Excel workbook.

Standalone: this file needs only openpyxl, so any notebook can use it:

    import gpe_rates
    M = gpe_rates.load_master('/content/drive/MyDrive/GST_Prime_Investigation/_RATE_MASTER/GST_RATE_MASTER_v3_2026-10-12.xlsx')
    gpe_rates.lookup(M['rows'], '7204', 'Oct-2025')

Every row carries the notification, schedule and serial number it comes from, the dates it was in force,
and the official file it was taken from (name and MD5). A rate is never inferred: an HSN with no row in
force for the month returns 'NO ENTRY IN MASTER'.
"""
import os, re, datetime, hashlib, calendar
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

RATES_VERSION = "1.2"

COLUMNS = ['Row ID', 'HSN', 'Level', 'Description of goods (verbatim)', 'IGST rate %', 'CGST rate %',
           'SGST/UTGST rate %', 'Compensation cess', 'Valid from', 'Valid to', 'Notification', 'Schedule',
           'S.No.', 'Excluding HSN', 'Scope', 'Source file', 'Source page', 'Source MD5', 'Status', 'Remarks']
REQUIRED = ['HSN', 'Description of goods (verbatim)', 'IGST rate %', 'Valid from', 'Notification']
STATUSES = ('OFFICIAL FILE', 'COMPILATION', 'TO VERIFY')
SLABS = (0, 0.1, 0.25, 1, 1.5, 3, 5, 6, 7.5, 12, 18, 28, 40)   # IGST-equivalent rates that exist in the law
TOL = 0.10                                                      # percentage points: rounding, not a rate difference

def md5_file(path):
    h = hashlib.md5()
    with open(path, 'rb') as f:
        for b in iter(lambda: f.read(1 << 20), b''):
            h.update(b)
    return h.hexdigest()

def norm_hsn(v):
    """'7204' / 7204 / '72.04' / '7204 10 00' -> '7204' ... Leading zero restored where a numeric export
    dropped it (401 -> 0401). '*' (or 'ANY', 'RESIDUAL') is the residual entry."""
    if v is None: return None
    s = str(v).strip().upper()
    if s in ('*', 'RESIDUAL'): return '*'
    if s in ('ANY', 'ANY CHAPTER'): return 'ANY'
    if isinstance(v, float) and v.is_integer(): s = str(int(v))
    s = re.sub(r'[\s.\-]', '', s)
    if not s.isdigit(): return None
    if len(s) % 2 == 1: s = '0' + s
    return s

def level_of(h):
    return {'*': 'Residual', 'ANY': 'Any Chapter (by description)', 2: 'Chapter', 4: 'Heading', 6: 'Sub-heading',
            8: 'Tariff item'}.get(h if h in ('*', 'ANY') else len(h), f'{len(h)} digits')

def to_date(v):
    if v is None or v == '': return None
    if isinstance(v, datetime.datetime): return v.date()
    if isinstance(v, datetime.date): return v
    s = str(v).strip()
    for f in ('%d-%m-%Y', '%d.%m.%Y', '%d/%m/%Y', '%Y-%m-%d', '%d-%b-%Y', '%Y-%m-%d %H:%M:%S'):
        try: return datetime.datetime.strptime(s, f).date()
        except ValueError: pass
    raise ValueError(f'date not understood: {v!r} (use DD-MM-YYYY)')

def to_rate(v):
    if v is None or v == '': return None
    if isinstance(v, (int, float)): return float(v)
    s = str(v).strip().replace('%', '')
    return float(s) if s else None

def month_bounds(period):
    """'Sep-2025' -> (2025-09-01, 2025-09-30)"""
    m = re.fullmatch(r'([A-Za-z]{3})[-/ ](\d{4})', str(period).strip())
    if not m: raise ValueError(f'period not understood: {period!r}')
    mon = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'].index(m.group(1).title()) + 1
    y = int(m.group(2))
    return datetime.date(y, mon, 1), datetime.date(y, mon, calendar.monthrange(y, mon)[1])

# ------------------------------------------------------------------ read / validate / write
def validate(rows):
    """Returns (clean_rows, errors). Nothing is guessed: a row with a problem is reported, not repaired,
    except that a missing Row ID is numbered and CGST/SGST are derived as IGST/2 when both are blank."""
    errs, out, seen = [], [], set()
    nxt = 1 + max([int(r['Row ID']) for r in rows if str(r.get('Row ID') or '').isdigit()] or [0])
    for i, r in enumerate(rows, 2):
        if not any(v not in (None, '') for v in r.values()): continue
        e = []
        for c in REQUIRED:
            if r.get(c) in (None, ''): e.append(f'{c} is blank')
        h = norm_hsn(r.get('HSN'))
        if r.get('HSN') not in (None, '') and h is None: e.append(f'HSN {r.get("HSN")!r} is not digits or *')
        try:
            ig = to_rate(r.get('IGST rate %')); cg = to_rate(r.get('CGST rate %')); sg = to_rate(r.get('SGST/UTGST rate %'))
            if ig is not None and not (0 <= ig <= 40): e.append(f'IGST rate {ig} outside 0-40')
            if ig is not None and cg is None and sg is None: cg = sg = ig / 2
            if None not in (ig, cg, sg) and abs(cg + sg - ig) > 0.001:
                e.append(f'CGST {cg} + SGST {sg} does not equal IGST {ig}')
        except ValueError:
            e.append('a rate is not a number'); ig = cg = sg = None
        try:
            vf, vt = to_date(r.get('Valid from')), to_date(r.get('Valid to'))
            if vf and vt and vt < vf: e.append('Valid to is before Valid from')
        except ValueError as x:
            e.append(str(x)); vf = vt = None
        st = (r.get('Status') or 'TO VERIFY').strip().upper()
        if st not in STATUSES: e.append(f'Status {st!r} must be one of {STATUSES}')
        rid = r.get('Row ID')
        rid = int(rid) if str(rid or '').strip().isdigit() else None
        if rid is None: rid = nxt; nxt += 1
        if rid in seen: e.append(f'Row ID {rid} repeated')
        seen.add(rid)
        if e:
            errs.append(f'row {i}: ' + '; '.join(e)); continue
        ex = [norm_hsn(x) for x in re.split(r'[,;]', str(r.get('Excluding HSN') or '')) if x.strip()]
        sc = str(r.get('Scope') or '').strip().upper()
        sc = 'PART OF CODE' if sc.startswith('PART') else ('WHOLE CODE' if sc.startswith('WHOLE') else None)
        if h == 'ANY': sc = 'PART OF CODE'
        out.append({**{c: r.get(c) for c in COLUMNS}, 'Row ID': rid, 'HSN': h, 'Level': level_of(h),
                    'Excluding HSN': ', '.join(x for x in ex if x) or None, 'Scope': sc,
                    'IGST rate %': ig, 'CGST rate %': cg, 'SGST/UTGST rate %': sg,
                    'Valid from': vf, 'Valid to': vt, 'Status': st})
    return out, errs

def _sheet_rows(ws):
    it = ws.iter_rows(values_only=True)
    hdr = [str(x).strip() if x is not None else '' for x in next(it, [])]
    return hdr, [dict(zip(hdr, r)) for r in it]

def load_master(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    if 'RATES' not in wb.sheetnames:
        raise ValueError('Not a rate master: no RATES sheet.')
    hdr, raw = _sheet_rows(wb['RATES'])
    miss = [c for c in REQUIRED if c not in hdr]
    if miss: raise ValueError(f'RATES sheet lacks column(s): {miss}')
    rows, errs = validate(raw)
    extra = {}
    for s in ('SOURCES', 'CHANGE_LOG'):
        extra[s] = _sheet_rows(wb[s])[1] if s in wb.sheetnames else []
    m = re.search(r'_v(\d+)_', os.path.basename(path))
    return dict(path=path, version=int(m.group(1)) if m else 0, rows=rows, errors=errs,
                sources=extra['SOURCES'], changelog=extra['CHANGE_LOG'], md5=md5_file(path))

def diff(old, new):
    """Row-level difference keyed on Row ID: added, removed, changed (field by field)."""
    o = {r['Row ID']: r for r in old}; n = {r['Row ID']: r for r in new}
    added = [n[k] for k in n if k not in o]
    removed = [o[k] for k in o if k not in n]
    changed = []
    for k in n:
        if k in o:
            d = [(c, o[k].get(c), n[k].get(c)) for c in COLUMNS if c != 'Row ID' and o[k].get(c) != n[k].get(c)]
            if d: changed.append((k, d))
    return added, removed, changed

README = [
    'GST RATE MASTER',
    '',
    'One row per rate entry of a rate notification, as in force between Valid from and Valid to (blank = still in force).',
    'HSN is written as in the notification: 2 digits = Chapter, 4 = Heading, 6 = Sub-heading, 8 = Tariff item; * = the residual',
    'entry (goods not elsewhere specified). Keep the description VERBATIM: the rate turns on the description, and the HSN is',
    'only the route to it (Explanation to the rate notifications: Chapter, heading, sub-heading and tariff item mean those of',
    'the First Schedule to the Customs Tariff Act, 1975, read with its rules for interpretation and notes).',
    '',
    'Status: OFFICIAL FILE = taken from the Gazette / CBIC copy of the notification named; file name and MD5 recorded.',
    '        COMPILATION   = taken from a published compilation of the notification (for example the ICAI print).',
    '        TO VERIFY     = entered by hand without a source file.',
    '        Only OFFICIAL FILE entries are counted in the findings; the others are shown as indicative. To upgrade,',
    '        read the Gazette / CBIC copy through the notebook: it replaces the entries of the same notification.',
    '"Excluding HSN" lists codes the notification carves out of an entry ([other than ...], (Except ...)).',
    'HSN "ANY" = an entry for goods of any Chapter, decided by description alone; it is shown, never matched by code.',
    'Scope "PART OF CODE" = the entry covers only some goods of its code (a Chapter-level entry with a specific description,',
    '        or a heading entry such as 2309 "Fish soluble paste"). Where only such entries match, the residual entry is',
    '        offered as the rate for the other goods; a shorter-code PART entry at a lower rate is shown as a concession.',
    '',
    'How the check reads it: for each HSN line of a month, every row in force on any day of that month whose HSN is a',
    'prefix of the reported HSN is a candidate; the longest prefix wins. If the reported HSN is shorter than entries in the',
    'master (reported 7204, master has 72041000 and 72042100 at different rates), the result is DEPENDS ON DESCRIPTION.',
    'If the rate changed during the month, both rates are shown and the month is marked RATE CHANGED DURING THE MONTH.',
    'If nothing matches and a residual (*) row is in force, the result is RESIDUAL ENTRY.',
    '',
    'To update: download, edit or add rows in RATES (keep Row ID for existing rows; leave it blank for new ones), fill in',
    'Notification, Schedule, S.No., Valid from/to and Source, and upload through the Rate master cell. Each upload is saved',
    'as a new numbered version; earlier versions are never overwritten, and CHANGE_LOG records what changed.',
]

def save_master(rows, path, sources=None, changelog=None):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'RATES'
    B = Font(bold=True); T = Side(style='thin'); BD = Border(left=T, right=T, top=T, bottom=T)
    ws.append(COLUMNS)
    for c in ws[1]:
        c.font = B; c.border = BD; c.alignment = Alignment(wrap_text=True, vertical='top')
    for r in sorted(rows, key=lambda r: (r['HSN'] == '*', r['HSN'], r['Valid from'] or datetime.date.min, r['Row ID'])):
        ws.append([r.get(c) for c in COLUMNS])
    for i, w in enumerate([8, 11, 11, 60, 9, 9, 9, 14, 12, 12, 30, 10, 7, 16, 14, 34, 8, 34, 14, 40], 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.border = BD; c.alignment = Alignment(wrap_text=True, vertical='top')
        row[1].number_format = '@'
        for k in (8, 9): row[k].number_format = 'DD-MM-YYYY'
        if row[18].value != 'OFFICIAL FILE': row[18].fill = PatternFill('solid', fgColor='FFD9D9')
    ws.freeze_panes = 'A2'
    s = wb.create_sheet('SOURCES')
    s.append(['File', 'MD5', 'Notification', 'Dated', 'Effective from', 'Where downloaded', 'Loaded on', 'Rows'])
    for r in sources or []: s.append([r.get(k) for k in ['File', 'MD5', 'Notification', 'Dated', 'Effective from',
                                                         'Where downloaded', 'Loaded on', 'Rows']])
    l = wb.create_sheet('CHANGE_LOG')
    l.append(['Version', 'Date', 'Rows added', 'Rows changed', 'Rows removed', 'What changed'])
    for r in changelog or []: l.append([r.get(k) for k in ['Version', 'Date', 'Rows added', 'Rows changed',
                                                           'Rows removed', 'What changed']])
    for sh in (s, l):
        for c in sh[1]: c.font = B
    rd = wb.create_sheet('README')
    for t in README: rd.append([t])
    rd.column_dimensions['A'].width = 120
    wb.save(path)
    return path

# ------------------------------------------------------------------ lookup
def lookup(rows, hsn, period):
    """All rows that can govern `hsn` in `period`, and what they say together."""
    h = norm_hsn(hsn)
    first, last = month_bounds(period)
    live = [r for r in rows if r['Valid from'] and r['Valid from'] <= last
            and (r['Valid to'] is None or r['Valid to'] >= first)]
    if h is None:
        return dict(status='HSN NOT READABLE', rows=[], rates=[], finer=[])
    def _excl(r):
        return any(h.startswith(x) for x in str(r.get('Excluding HSN') or '').replace(' ', '').split(',') if x)
    anych = [r for r in live if r['HSN'] == 'ANY']
    live = [r for r in live if r['HSN'] != 'ANY' and not _excl(r)]
    pre = [r for r in live if r['HSN'] != '*' and h.startswith(r['HSN'])]
    finer = [r for r in live if r['HSN'] != '*' and r['HSN'].startswith(h) and len(r['HSN']) > len(h)]
    basis = 'longest prefix'
    if pre:
        L = max(len(r['HSN']) for r in pre)
        chosen = [r for r in pre if len(r['HSN']) == L]
    elif finer:
        chosen, basis = [], 'finer entries only'
    else:
        # a heading number beyond the last heading the rate notifications list for that Chapter is most likely not a
        # heading of the Customs Tariff at all (9090, 9099: Chapter 90 ends at 9033) — reported, not matched to the residual
        heads = [r['HSN'][:4] for r in rows if r['HSN'] not in ('*', 'ANY') and len(r['HSN']) >= 4 and r['HSN'][:2] == h[:2]]
        if len(h) >= 4 and h[:2] != '99' and heads and h[:4] > max(heads):
            return dict(status='HSN CODE NOT RECOGNISED', rows=[], rates=[], finer=[], basis='none', any_chapter=anych,
                        concessions=[], covered_from=None, all_official=False,
                        note=f'heading {h[:4]} is beyond the last heading listed for Chapter {h[:2]} ({max(heads)})')
        chosen = [r for r in live if r['HSN'] == '*']
        basis = 'residual entry' if chosen else 'none'
    narrow = bool(chosen) and basis == 'longest prefix' and all(r.get('Scope') == 'PART OF CODE' for r in chosen)
    if narrow:
        # the only entries for this code cover some goods of it; the rest fall to the residual entry
        chosen = chosen + [r for r in live if r['HSN'] == '*']
        basis = 'narrow entry, residual entry for other goods'
    cand = chosen + [r for r in finer if r not in chosen]
    main = min((r['IGST rate %'] for r in chosen), default=None)
    conc = [r for r in pre if r not in chosen and r.get('Scope') == 'PART OF CODE'
            and main is not None and r['IGST rate %'] < main]
    rates = sorted({r['IGST rate %'] for r in cand})
    # a rate change inside the month: some candidate entry starts or ends part-way through it, and the rates differ
    split = len(rates) > 1 and any((first < r['Valid from'] <= last) or (r['Valid to'] and first <= r['Valid to'] < last)
                                   for r in cand)
    covered = min((r['Valid from'] for r in cand), default=None)
    if not cand:
        status = 'NO ENTRY IN MASTER'
    elif covered > first:
        status = 'MASTER COVERS ONLY PART OF THE MONTH'
    elif split:
        status = 'RATE CHANGED DURING THE MONTH'
    elif len(rates) > 1:
        status = 'DEPENDS ON DESCRIPTION'
    elif basis == 'residual entry':
        status = 'RESIDUAL ENTRY'
    else:
        status = 'SINGLE RATE'
    return dict(status=status, rows=cand, rates=rates, finer=finer, basis=basis, any_chapter=anych, concessions=conc,
                covered_from=covered,
                all_official=bool(cand) and all(r['Status'] == 'OFFICIAL FILE' for r in cand))

def nearest_slab(rate):
    return min(SLABS, key=lambda s: abs(s - rate))

def judge(rows, hsn, period, value, tax):
    """Compare the rate actually charged on one HSN line with the rate(s) notified for that month."""
    out = dict(hsn=norm_hsn(hsn) or str(hsn), period=period, value=value, tax=tax, eff=None, slab=None,
               off_slab=False, notified=None, result='', short=0.0, countable=False, look=None)
    if value is None or value <= 0:
        out['result'] = 'NOT TESTED — taxable value nil or negative (credit notes are tested on 19B)'
        return out
    eff = 100.0 * (tax or 0.0) / value
    out['eff'] = eff; out['slab'] = nearest_slab(eff); out['off_slab'] = abs(eff - out['slab']) > TOL
    lk = lookup(rows, hsn, period) if rows else dict(status='NO ENTRY IN MASTER', rows=[], rates=[], finer=[],
                                                     all_official=False, basis='none')
    out['look'] = lk
    rates = lk['rates']
    if not rates:
        out['result'] = lk['status'] + (f' — {lk["note"]}; the true rate cannot be found from this code' if lk.get('note') else '')
        return out
    lo = min(rates)
    out['notified'] = rates[0] if len(rates) == 1 else rates
    if any(abs(eff - r) <= TOL for r in rates):
        out['result'] = ('AT THE NOTIFIED RATE' if len(rates) == 1 else
                         'AT ONE OF THE POSSIBLE RATES — confirm the description')
    elif eff < lo - TOL:
        out['short'] = value * (lo - eff) / 100.0
        out['result'] = ('BELOW THE NOTIFIED RATE' if len(rates) == 1 else
                         'BELOW EVERY POSSIBLE RATE — shortfall shown against the lowest')
        # a residual match is never counted: exempt goods sit in a separate exemption notification
        out['countable'] = lk['status'] == 'SINGLE RATE' and lk['all_official']
        conc = lk.get('concessions') or []
        if conc and eff >= min(r['IGST rate %'] for r in conc) - TOL:
            cr = sorted({r['IGST rate %'] for r in conc})
            out['result'] += (f' — but at or above the concessional rate of {"/".join(f"{x:g}" for x in cr)}% '
                              f'for specified goods (' + '; '.join(f'{r["Notification"].split(" dated")[0]} Sch {r.get("Schedule") or "-"} '
                                                                   f'S.No. {r.get("S.No.") or "-"}' for r in conc[:3])
                              + '): a mix, or goods that qualify — check the invoices')
            out['countable'] = False
    elif eff > max(rates) + TOL:
        out['result'] = 'ABOVE THE NOTIFIED RATE (tax over-charged; no revenue loss)'
    else:
        out['result'] = 'BETWEEN THE POSSIBLE RATES — mixed supplies or wrong rate on part; check invoices'
    if lk['status'] == 'MASTER COVERS ONLY PART OF THE MONTH':
        # compared, but the days of the month before the first entry in force have no entry in the master
        out['result'] += f' — the master covers this month only from {lk["covered_from"]:%d-%m-%Y}'
        out['countable'] = False
    return out

# ------------------------------------------------------------------ reading a rate notification PDF
_ROMAN = {'I': 1, 'II': 2, 'III': 3, 'IV': 4, 'V': 5, 'VI': 6, 'VII': 7, 'VIII': 8, 'IX': 9, 'X': 10}
_HSN_TOK = re.compile(r'^\(?(?:\d{2,8}|Except|Any|Chapter|to|or|and|&)\)?[,;]?$', re.I)
_MONTHS = {m: i for i, m in enumerate(['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august',
                                       'september', 'october', 'november', 'december'], 1)}

def _lines(page):
    ws = page.extract_words(x_tolerance=1.5, y_tolerance=3, keep_blank_chars=False)
    rows = {}
    for w in ws:
        rows.setdefault(round(w['top'] / 4), []).append(w)
    return [sorted(rows[k], key=lambda w: w['x0']) for k in sorted(rows)]

def _code_tokens_ok(t):
    return bool(_HSN_TOK.match(t))

def parse_hsn_cell(text):
    """'0202, 0203, 0210' / '7208 to 7212' / '0507 (Except 050790)' / '7310 or 7326' / '0101 21 00' / 'Any Chapter'
    -> (codes, excluded codes). Ranges are expanded heading by heading."""
    t = ' '.join(text.split())
    t = re.sub(r'\bor(?=\d)', 'or ', t)
    t = re.sub(r'\[\s*Except\b', '(Except', t, flags=re.I)
    t = re.sub(r'\(\s*other\s+than\b', '[other than', t, flags=re.I)
    t = re.sub(r'(\[\s*other\s+than[^\])]*)[)\]]?\s*$', r'\1]', t, flags=re.I)   # unclosed bracket at the end
    t = re.sub(r'(\(\s*Except[^\])]*)\]', r'\1)', t, flags=re.I)
    excl = []
    for m in re.finditer(r'\(\s*Except\s+([^)]*)\)', t, re.I):
        excl += [re.sub(r'\s', '', x) for x in re.split(r',|\band\b|\bor\b', m.group(1)) if re.sub(r'\s', '', x)]
    t = re.sub(r'\(\s*Except\s+[^)]*\)', ' ', t, flags=re.I)
    for m in re.finditer(r'\[\s*other\s+than\s+([^\]]*)\]', t, re.I):
        excl += [re.sub(r'\s', '', x) for x in re.split(r',|\band\b|\bor\b', m.group(1)) if re.sub(r'\s', '', x)]
    t = re.sub(r'\[\s*other\s+than\s+[^\]]*\]', ' ', t, flags=re.I)
    codes = []
    if re.search(r'any\s+(other\s+)?chapter', t, re.I):
        codes.append('ANY')
        t = re.sub(r'(\bor\s+)?any\s+(other\s+)?chapter', ' ', t, flags=re.I)
    for part in re.split(r',|\bor\b|\band\b|&', t):
        part = part.strip()
        if not part: continue
        m = re.fullmatch(r'([\d ]+?)\s+to\s+([\d ]+)', part)
        if m:
            a, b = re.sub(r'\s', '', m.group(1)), re.sub(r'\s', '', m.group(2))
            if len(a) == len(b) and a.isdigit() and b.isdigit() and int(b) >= int(a) and int(b) - int(a) < 200:
                codes += [str(x).zfill(len(a)) for x in range(int(a), int(b) + 1)]
                continue
        c = re.sub(r'\s', '', part)
        if c.isdigit(): codes.append(c)
        else: return None, excl            # not understood: the caller reports it
    return codes, excl

def parse_notification_pdf(path):
    """Read the Schedules of a GST goods-rate notification (the layout of 1/2017- and 9/2025-Integrated Tax (Rate):
    S.No. | Chapter/Heading/Sub-heading/Tariff item | Description). Returns (meta, entries, problems).
    Nothing is assumed: the rate of each Schedule is read from the notification's own opening paragraph and
    checked against the Schedule's heading; serial numbers are checked for gaps."""
    import pdfplumber
    pdf = pdfplumber.open(path)
    full = '\n'.join((p.extract_text() or '') for p in pdf.pages[:3])
    meta = dict(file=os.path.basename(path), md5=md5_file(path), pages=len(pdf.pages))
    m = re.search(r'NOTIFICATION\s+NO\.?\s*(\d+/\d{4})\s*-\s*(INTEGRATED|CENTRAL|STATE|UNION TERRITORY)\s+TAX\s*\(RATE\)'
                  r'\s*DATED\s*([\d\-./]+)', full, re.I)
    if m:
        meta['notification'] = f'{m.group(1)}-{m.group(2).title()} Tax (Rate)'; meta['dated'] = m.group(3)
        meta['tax'] = m.group(2).upper()
    if re.search(r'corrigendum', full, re.I):
        c = re.search(r'corrected by corrigendum,?\s*dated\s*([\d\-./]+)', full, re.I)
        meta['corrigendum'] = c.group(1) if c else 'yes'
    meta['compilation'] = 'idtc.icai.org' if 'idtc.icai.org' in full else None
    sup = re.search(r'in supersession of the notification .*?No\.\s*([\w/]+\s*-\s*[A-Za-z ]+\(Rate\))', full.replace('\n', ' '), re.I)
    meta['supersedes'] = ' '.join(sup.group(1).split()) if sup else None
    rates = {}
    for r_, s_ in re.findall(r'(\d+(?:\.\d+)?)\s*per\s*cent\.?\s*in\s+respect\s+of\s+goods\s+specified\s+in\s+Schedule\s+([IVX]+)',
                             full.replace('\n', ' '), re.I):
        rates[s_.upper()] = float(r_)
    meta['schedule_rates'] = rates
    tail = '\n'.join((p.extract_text() or '') for p in pdf.pages[-2:]).replace('\n', ' ')
    e = re.search(r'come into force with effect from the (\d+)\w*\s+day of (\w+),?\s*(\d{4})', tail, re.I)
    if e and e.group(2).lower() in _MONTHS:
        meta['effective_from'] = datetime.date(int(e.group(3)), _MONTHS[e.group(2).lower()], int(e.group(1)))
    entries, problems = [], []
    sched, cur, done, in_list = None, None, False, False
    page_dx = None
    for pno, page in enumerate(pdf.pages, 1):
        lines = _lines(page)
        has_fn = any('Corrected vide' in ' '.join(w['text'] for w in L) for L in lines)
        # where the description column starts on this page: the most common position of the first word after
        # the codes on numbered entry lines (a word that cannot be part of a code)
        starts, sch_ = {}, sched
        for L in lines:
            t0 = ' '.join(w['text'] for w in L)
            m0 = re.match(r'^Schedule\s+([IVX]+)\s*[–\-—]\s*[\d.]+\s*%', t0)
            if m0: sch_ = m0.group(1); continue
            if re.fullmatch(r'\d+\.', L[0]['text']) and L[0]['x0'] < 110 and len(L) > 2:
                inb = False
                for w in L[1:]:
                    t_ = w['text']
                    if t_.startswith('['): inb = True
                    if inb or _HSN_TOK.match(t_):
                        if t_.endswith(']'): inb = False
                        continue
                    starts.setdefault(sch_, []).append(round(w['x0'])); break
        page_dxs = {k: max(set(v), key=v.count) for k, v in starts.items()}
        if sched in page_dxs: page_dx = page_dxs[sched]
        for L in lines:
            text = ' '.join(w['text'] for w in L)
            if done: break
            if text.startswith('Printed from') or re.match(r'^NOTIFICATION NO\.\s*\d+/\d{4}', text): continue
            if re.match(r'^List\s+\d+\s*\[See', text):
                in_list = True; cur = None; continue
            if in_list and not re.match(r'^Schedule\s+[IVX]+\s*[–\-—]', text): continue
            if 'Corrected vide' in text or (has_fn and re.fullmatch(r'\d{1,2}', text)): continue
            sh = re.match(r'^Schedule\s+([IVX]+)\s*[–\-—]\s*([\d.]+)\s*%', text)
            if sh:
                sched = sh.group(1); in_list = False
                if sched in page_dxs: page_dx = page_dxs[sched]
                if sched in rates and abs(rates[sched] - float(sh.group(2))) > 1e-9:
                    problems.append(f'Schedule {sched}: heading says {sh.group(2)}% but the opening paragraph says {rates[sched]}%')
                rates.setdefault(sched, float(sh.group(2)))
                cur = None; continue
            if sched is None: continue
            if re.match(r'^Explanation\.?\s*[—\-–]\s*For the purposes of this notification', text) or \
               re.match(r'^\d\.\s*This notification shall come into force', text):
                done = True; break
            if re.fullmatch(r'\(1\)\s*\(2\)\s*\(3\)', text) or re.match(r'^S\.?\s*(No\.)?\s*Chapter', text) or \
               re.fullmatch(r'(No\.|S\.|item|Tariff|/|Sub-?heading|heading|Heading|Subheading|Chapter|-|\s)+', text) or \
               text in ('No.', 'item'):
                continue
            first = L[0]
            if re.fullmatch(r'\d+\.', first['text']) and first['x0'] < 110:
                hs, ds, dx = [], [], None
                for w in L[1:]:
                    code_col = (w['x0'] < page_dx - 4) if page_dx else bool(_HSN_TOK.match(w['text']))
                    if dx is None and code_col:
                        hs.append(w['text'])
                    else:
                        if dx is None: dx = w['x0']
                        ds.append(w['text'])
                cur = dict(schedule=sched, sno=int(first['text'][:-1]), hsn=hs, desc=ds, page=pno, dx=dx or page_dx)
                entries.append(cur)
                continue
            if cur is None: continue
            lim = (page_dx if (page_dx and cur['page'] != pno) else cur['dx']) or 250
            for w in L:
                (cur['hsn'] if w['x0'] < lim - 4 else cur['desc']).append(w['text'])
    out = []
    for e_ in entries:
        codes, excl = parse_hsn_cell(' '.join(e_['hsn']))
        desc = ' '.join(e_['desc']).strip()
        if codes == ['ANY'] and re.match(r'Goods which are not specified in', desc, re.I):
            codes = ['*']                               # the residual entry
        if not codes:
            problems.append(f'Schedule {e_["schedule"]} S.No. {e_["sno"]} (page {e_["page"]}): code column not understood: '
                            f'{" ".join(e_["hsn"])!r}')
            continue
        for c in codes:
            out.append(dict(schedule=e_['schedule'], sno=e_['sno'], hsn=c, excluded=excl, desc=desc, page=e_['page'],
                            hsn_text=' '.join(e_['hsn']), rate=rates.get(e_['schedule'])))
    # serial numbers must run 1..N in every Schedule
    for s in sorted({e_['schedule'] for e_ in entries}, key=lambda x: _ROMAN.get(x, 99)):
        sn = [e_['sno'] for e_ in entries if e_['schedule'] == s]
        exp = list(range(1, max(sn) + 1))
        miss = sorted(set(exp) - set(sn)); dup = sorted({x for x in sn if sn.count(x) > 1})
        if miss: problems.append(f'Schedule {s}: serial number(s) not found: {miss[:30]}')
        if dup: problems.append(f'Schedule {s}: serial number(s) read twice: {dup[:30]}')
    meta['entries'] = {s: len({e_['sno'] for e_ in entries if e_['schedule'] == s}) for s in rates}
    return meta, out, problems

def rows_from_notification(meta, parsed, status, valid_from=None, valid_to=None, start_id=1):
    """Turn parsed entries into master rows. `status` is OFFICIAL FILE only for the Gazette / CBIC copy."""
    vf = valid_from or meta.get('effective_from')
    note = ''
    if meta.get('corrigendum'): note += f'As corrected by corrigendum dated {meta["corrigendum"]}. '
    if meta.get('compilation'): note += f'Read from the compilation printed from {meta["compilation"]}. '
    rows = []
    for i, e in enumerate(parsed, start_id):
        ig = e['rate']
        rows.append({'Row ID': i, 'HSN': e['hsn'], 'Description of goods (verbatim)': e['desc'],
                     'IGST rate %': ig, 'CGST rate %': ig / 2 if ig is not None else None,
                     'SGST/UTGST rate %': ig / 2 if ig is not None else None, 'Compensation cess': None,
                     'Valid from': vf, 'Valid to': valid_to,
                     'Notification': f'{meta.get("notification")} dated {meta.get("dated")}',
                     'Schedule': e['schedule'], 'S.No.': e['sno'], 'Source file': meta['file'], 'Source page': e['page'],
                     'Source MD5': meta['md5'], 'Status': status, 'Excluding HSN': ', '.join(e['excluded']) or None,
                     'Scope': ('PART OF CODE' if (e['hsn'] == 'ANY' or (len(e['hsn']) == 2 and
                               not e['desc'].lower().startswith('all'))) else None),
                     'Remarks': (note + (f'Code column reads "{e["hsn_text"]}".' if e['hsn_text'] and
                                         re.sub(r'\D', '', e['hsn_text']) != e['hsn'] else '')).strip() or None})
    return rows

def merge_notification(old_rows, meta, parsed, status):
    """Add the entries of one rate notification to the master.
    * Entries already held from the SAME notification are replaced (re-reading a corrected copy is safe).
    * If the notification says it supersedes another (9/2025-IT(R) supersedes 1/2017-IT(R)), every open entry of
      the superseded notification is closed the day before the new one takes effect.
    Returns (rows, summary)."""
    key = meta.get('notification') or ''
    num = re.match(r'0*(\d+/\d{4})', key)
    kind = 'Integrated' if 'Integrated' in key else ('Central' if 'Central' in key else '')
    def same(r, n):
        t = str(r.get('Notification') or '')
        return bool(re.search(rf'(^|\D)0*{re.escape(n)}\b', t)) and kind in t
    keep = [dict(r) for r in old_rows if not (num and same(r, num.group(1)))]
    replaced = len(old_rows) - len(keep)
    closed = 0
    eff = meta.get('effective_from')
    if meta.get('supersedes') and eff:
        sm = re.match(r'0*(\d+/\d{4})', meta['supersedes'])
        if sm:
            for r in keep:
                if same(r, sm.group(1)) and (r['Valid to'] is None or r['Valid to'] >= eff):
                    r['Valid to'] = eff - datetime.timedelta(days=1); closed += 1
    start = 1 + max([int(r['Row ID']) for r in keep] or [0])
    new = rows_from_notification(meta, parsed, status, start_id=start)
    src = {'File': meta['file'], 'MD5': meta['md5'], 'Notification': key, 'Dated': meta.get('dated'),
           'Effective from': eff, 'Where downloaded': meta.get('compilation') or 'official copy',
           'Loaded on': datetime.date.today(), 'Rows': len(new)}
    return keep + new, dict(replaced=replaced, closed=closed, added=len(new), source=src)
