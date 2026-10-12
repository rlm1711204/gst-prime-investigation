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

RATES_VERSION = "1.0"

COLUMNS = ['Row ID', 'HSN', 'Level', 'Description of goods (verbatim)', 'IGST rate %', 'CGST rate %',
           'SGST/UTGST rate %', 'Compensation cess', 'Valid from', 'Valid to', 'Notification', 'Schedule',
           'S.No.', 'Source file', 'Source page', 'Source MD5', 'Status', 'Remarks']
REQUIRED = ['HSN', 'Description of goods (verbatim)', 'IGST rate %', 'Valid from', 'Notification']
STATUSES = ('OFFICIAL FILE', 'TO VERIFY')
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
    if s in ('*', 'ANY', 'RESIDUAL', 'ANY CHAPTER'): return '*'
    if isinstance(v, float) and v.is_integer(): s = str(int(v))
    s = re.sub(r'[\s.\-]', '', s)
    if not s.isdigit(): return None
    if len(s) % 2 == 1: s = '0' + s
    return s

def level_of(h):
    return {'*': 'Residual', 2: 'Chapter', 4: 'Heading', 6: 'Sub-heading', 8: 'Tariff item'}.get(
        h if h == '*' else len(h), f'{len(h)} digits')

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
        out.append({**{c: r.get(c) for c in COLUMNS}, 'Row ID': rid, 'HSN': h, 'Level': level_of(h),
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
    'Status: OFFICIAL FILE = taken from the official notification named, whose file name and MD5 are recorded;',
    '        TO VERIFY   = entered without an official file. Such a row is shown in the rate check but never counted.',
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
    for i, w in enumerate([8, 11, 11, 60, 9, 9, 9, 14, 12, 12, 30, 10, 7, 34, 8, 34, 14, 40], 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.border = BD; c.alignment = Alignment(wrap_text=True, vertical='top')
        row[1].number_format = '@'
        for k in (8, 9): row[k].number_format = 'DD-MM-YYYY'
        if row[16].value == 'TO VERIFY': row[16].fill = PatternFill('solid', fgColor='FFD9D9')
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
    pre = [r for r in live if r['HSN'] != '*' and h.startswith(r['HSN'])]
    finer = [r for r in live if r['HSN'] != '*' and r['HSN'].startswith(h) and len(r['HSN']) > len(h)]
    basis = 'longest prefix'
    if pre:
        L = max(len(r['HSN']) for r in pre)
        chosen = [r for r in pre if len(r['HSN']) == L]
    elif finer:
        chosen, basis = [], 'finer entries only'
    else:
        chosen = [r for r in live if r['HSN'] == '*']
        basis = 'residual entry' if chosen else 'none'
    cand = chosen + [r for r in finer if r not in chosen]
    rates = sorted({r['IGST rate %'] for r in cand})
    windows = sorted({(r['Valid from'], r['Valid to']) for r in cand}, key=lambda w: w[0])
    split = len(windows) > 1 and len({tuple(sorted({x['IGST rate %'] for x in cand if (x['Valid from'], x['Valid to']) == w}))
                                      for w in windows}) > 1
    if not cand:
        status = 'NO ENTRY IN MASTER'
    elif split:
        status = 'RATE CHANGED DURING THE MONTH'
    elif len(rates) > 1:
        status = 'DEPENDS ON DESCRIPTION'
    elif basis == 'residual entry':
        status = 'RESIDUAL ENTRY'
    else:
        status = 'SINGLE RATE'
    return dict(status=status, rows=cand, rates=rates, finer=finer, basis=basis,
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
        out['result'] = lk['status']
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
        out['countable'] = lk['status'] == 'SINGLE RATE' and lk['all_official']
    elif eff > max(rates) + TOL:
        out['result'] = 'ABOVE THE NOTIFIED RATE (tax over-charged; no revenue loss)'
    else:
        out['result'] = 'BETWEEN THE POSSIBLE RATES — mixed supplies or wrong rate on part; check invoices'
    return out
