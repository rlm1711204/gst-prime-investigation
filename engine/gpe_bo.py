# -*- coding: utf-8 -*-
"""
GST-PRIME PRELIMINARY INVESTIGATION ENGINE  —  BO comparison report
Reads the 'Goods and Services Tax - Tax liabilities and ITC Comparison' workbook downloaded from the
BO portal (one file per financial year), compares it with the figures the notebook built from the
GST-Prime exports, and supplies the final position on which the findings rest.

The BO report is taken as the authoritative figure wherever it covers a tax period:
  * it is the department's own comparison, the one Rule 88C (DRC-01B) and Rule 88D (DRC-01C) run on;
  * it reads GSTR-1/IFF/1A as filed during the month, amendments included, and GSTR-2B, which since
    1 January 2022 governs entitlement to credit (section 16(2)(aa), Rule 36(4));
  * it carries the time it was generated, so a later download can be told from an earlier one.
A month is NOT taken from the BO report when the report was generated before that month's GSTR-3B was
filed or fell due (the report would show a return that did not yet exist). The GST-Prime figure is then
kept for that month and the workbook says so.
"""
import datetime, re, warnings
import openpyxl
warnings.filterwarnings('ignore', message='DrawingML support is incomplete')
from openpyxl.utils import column_index_from_string, get_column_letter
from gpe_core import parse_period, period_key, fy_of, n0, inr

BO_VERSION = "1.2"
BO_CODE = 'BO_CMP'
TOL = 5.0          # rupees: Prime reports whole rupees per head, BO reports paise; up to Rs.5 is rounding
H4 = ('IGST', 'CGST', 'SGST', 'CESS')
H3 = ('IGST', 'CGST', 'SGST')
H2 = ('IGST', 'CESS')

# (section, sheet, [(side, first column, heads, keyword the block header must contain)])
LAYOUT = [
    ('Outward',       'Tax liability',          [('GSTR-3B', 'B', H4, 'GSTR-3B'), ('GSTR-1', 'F', H4, 'GSTR-1')]),
    ('Export-SEZ',    'Export and SEZ',         [('GSTR-3B', 'B', H2, 'GSTR-3B'), ('GSTR-1', 'D', H2, 'GSTR-1')]),
    ('RCM liability', 'Reverse charge',         [('GSTR-3B', 'B', H4, 'GSTR-3B'), ('GSTR-2B', 'F', H4, 'GSTR-2B')]),
    ('ITC other',     'ITC (Other than IMPG)',  [('GSTR-3B gross', 'B', H4, 'excluding ITC Reversal'),
                                                 ('GSTR-2B', 'F', H4, 'GSTR-2B'),
                                                 ('GSTR-3B net', 'V', H4, 'considering ITC Reversal')]),
    ('ITC import',    'ITC (IMPG)',             [('GSTR-3B', 'B', H2, 'GSTR-3B'), ('GSTR-2B', 'D', H2, 'GSTR-2B')]),
    ('ITC on RCM',    'RCM_LIABILITY_ITC',      [('GSTR-3B', 'B', H4, 'ITC claimed')]),
    ('DRC-03',        'Tax Liability Summary',  [('paid', 'Q', H4, 'DRC-03')]),
]
# the BO report's own "Shortfall (-) / Excess (+)" block: (sheet, first column, heads, minuend side, subtrahend side)
SELF_CHECK = [('Tax liability', 'J', H4, ('Outward', 'GSTR-3B'), ('Outward', 'GSTR-1')),
              ('Reverse charge', 'J', H4, ('RCM liability', 'GSTR-3B'), ('RCM liability', 'GSTR-2B')),
              ('ITC (Other than IMPG)', 'J', H4, ('ITC other', 'GSTR-3B gross'), ('ITC other', 'GSTR-2B')),
              ('ITC (Other than IMPG)', 'Z', H4, ('ITC other', 'GSTR-3B net'), ('ITC other', 'GSTR-2B')),
              ('Export and SEZ', 'F', H2, ('Export-SEZ', 'GSTR-3B'), ('Export-SEZ', 'GSTR-1')),
              ('ITC (IMPG)', 'F', H2, ('ITC import', 'GSTR-3B'), ('ITC import', 'GSTR-2B'))]

TABLES = {   # what each block is, in the words of the report, for the SRC sheet and the derivation sheet
    ('Outward', 'GSTR-3B'): 'GSTR-3B Table 3.1(a) + 3.1.1(i)',
    ('Outward', 'GSTR-1'): 'GSTR-1/IFF/1A filed during the month — Tables 4A, 4C, 5, 6C, 7, 9A, 9B, 9C, 10, 11, 15, 15A (other than SEZ)',
    ('Export-SEZ', 'GSTR-3B'): 'GSTR-3B Table 3.1(b)',
    ('Export-SEZ', 'GSTR-1'): 'GSTR-1/IFF Tables 6A, 6B, 9A, 9B, 9C, 15, 15A (SEZ only)',
    ('RCM liability', 'GSTR-3B'): 'GSTR-3B Table 3.1(d)',
    ('RCM liability', 'GSTR-2B'): 'GSTR-2B reverse-charge B2B, B2BA, CDNR, CDNRA',
    ('ITC other', 'GSTR-3B gross'): 'GSTR-3B 4A(4)+4A(5)-4D(1) (before reversal)',
    ('ITC other', 'GSTR-2B'): 'GSTR-2B B2B, B2BA, CDNR, CDNRA, ECO, ECOA, ISD, ISDA (excluding RCM)',
    ('ITC other', 'GSTR-3B net'): 'GSTR-3B 4A(4)+4A(5)-4D(1)-4B(1)-4B(2) (after reversal)',
    ('ITC import', 'GSTR-3B'): 'GSTR-3B Table 4A(1)',
    ('ITC import', 'GSTR-2B'): 'GSTR-2B IMPG, IMPG (SEZ)',
    ('ITC on RCM', 'GSTR-3B'): 'GSTR-3B Table 4A(2)+4A(3) (ITC on inward supplies liable to reverse charge)',
    ('DRC-03', 'paid'): 'Payment made through DRC-03 (Tax Liability Summary)',
}

def col(section, side, head):
    return f'{section} | {side} | {head}'

FIELDS = [col(s, side, h) for s, _, blocks in LAYOUT for side, _, heads, _ in blocks for h in heads]

_MONTH_RE = re.compile(r'^([A-Za-z]{3})-(\d{2})$')

def _period(v):
    """'Apr-23' -> 'Apr-2023'."""
    if v is None: return None
    m = _MONTH_RE.match(str(v).strip())
    if not m: return None
    t = parse_period(f'{m.group(1)}-20{m.group(2)}')
    return t[0] if t else None

def _num(v):
    if v is None: return None
    if isinstance(v, (int, float)): return float(v)
    s = str(v).strip().replace(',', '')
    if s in ('', '-'): return None
    try: return float(s)
    except ValueError: return None

def _head(v):
    s = str(v or '').strip().upper()
    return 'SGST' if s.startswith('SGST') else s

# ------------------------------------------------------------------ reading
def is_bo_file(path):
    try:
        wb = openpyxl.load_workbook(path, read_only=True)
        names = set(wb.sheetnames)
        ok = {'Comparison Summary', 'ITC (Other than IMPG)', 'Tax liability'} <= names
        if ok:
            ws = wb[wb.sheetnames[0]]
            a1 = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))[0]
            ok = 'tax liabilities and itc comparison' in str(a1 or '').lower()
        wb.close()
        return ok
    except Exception:
        return False

def _meta(wb):
    m = dict(gstin=None, legal=None, trade=None, fy=None, generated=None)
    for ws in wb.worksheets[:2]:
        for row in ws.iter_rows(min_row=1, max_row=6, values_only=True):
            for c in row:
                s = str(c or '')
                if not s: continue
                if m['gstin'] is None and (g := re.search(r'GSTIN:\s*([0-9A-Z]{15})', s)): m['gstin'] = g.group(1)
                if m['legal'] is None and (g := re.search(r'Legal name:\s*(.+)', s)): m['legal'] = g.group(1).strip()
                if m['trade'] is None and (g := re.search(r'Trade name:\s*(.+)', s)): m['trade'] = g.group(1).strip()
                if m['fy'] is None and (g := re.search(r'Financial Year:\s*(\d{4})-(\d{2})', s)):
                    m['fy'] = f'{g.group(1)}-{g.group(2)}'
                if m['generated'] is None and (g := re.search(r'generated at:\s*(\d{2}/\d{2}/\d{4}\s+\d{1,2}:\d{2}\s*[AP]M)', s, re.I)):
                    try: m['generated'] = datetime.datetime.strptime(re.sub(r'\s+', ' ', g.group(1)).upper(), '%d/%m/%Y %I:%M %p')
                    except ValueError: pass
    return m

def _grid(ws):
    return [list(r) for r in ws.iter_rows(values_only=True)]

def _locate(grid):
    """(row index of 'Tax Period', first data row, last data row, total row) — 0-based."""
    hdr = next((i for i, r in enumerate(grid) if r and str(r[0] or '').strip().lower() == 'tax period'), None)
    if hdr is None: return None
    data = [i for i in range(hdr + 1, len(grid)) if grid[i] and _period(grid[i][0])]
    tot = next((i for i in range(hdr + 1, len(grid)) if grid[i] and str(grid[i][0] or '').strip().lower() == 'total'), None)
    return hdr, data, tot

def read_bo(path):
    """-> (meta, rows, notes). rows: one dict per tax period, keyed by FIELDS. Raises ValueError if the
    file is a BO comparison report whose layout is not the one this reader knows."""
    wb = openpyxl.load_workbook(path, data_only=True)
    meta = _meta(wb)
    notes, rows = [], {}
    # Reports generated up to about 2022 carry ONE GSTR-3B block on the ITC sheet, already net of the Table 4B reversal
    # ("4A(4)+4A(5)-4B(1)-4B(2)"), and no before-reversal block. Later reports carry both.
    old_itc = False
    if 'ITC (Other than IMPG)' in wb.sheetnames:
        g = _grid(wb['ITC (Other than IMPG)']); loc = _locate(g)
        if loc and loc[1]:
            t = ' '.join(str(g[r][1] or '') for r in range(loc[0], loc[1][0] - 1)).lower()
            old_itc = ('excluding itc reversal' not in t) and ('4b(1)' in t)
    meta['old_itc'] = old_itc
    layout = [(s_, sh, ([('GSTR-3B net', 'B', H4, '4B(1)'), ('GSTR-2B', 'F', H4, 'GSTR-2B')] if s_ == 'ITC other' and old_itc else bl))
              for s_, sh, bl in LAYOUT]
    for section, sheet, blocks in layout:
        if sheet not in wb.sheetnames:
            raise ValueError(f'sheet "{sheet}" is missing')
        g = _grid(wb[sheet]); loc = _locate(g)
        if not loc or not loc[1]:
            raise ValueError(f'no tax-period rows found on "{sheet}"')
        hdr, data, tot = loc
        headrow = data[0] - 1
        for side, c0, heads, kw in blocks:
            ci = column_index_from_string(c0) - 1
            text = ' '.join(str(g[r][ci] or '') for r in range(hdr, headrow))
            if kw.lower() not in text.lower():
                raise ValueError(f'"{sheet}" column {c0}: expected a block about {kw}, found "{text[:80]}"')
            got = [_head(g[headrow][ci + k]) for k in range(len(heads))]
            if got != list(heads):
                raise ValueError(f'"{sheet}" columns {c0}..: expected heads {heads}, found {got}')
            for r in data:
                p = _period(g[r][0])
                d = rows.setdefault(p, {})
                for k, h in enumerate(heads):
                    v = _num(g[r][ci + k])
                    if v is None and section == 'DRC-03':
                        v = 0.0          # BO leaves a month with no DRC-03 blank and totals it as 0
                    d[col(section, side, h)] = v
            if tot is not None:
                for k, h in enumerate(heads):
                    t = _num(g[tot][ci + k])
                    s = sum(n0(rows[_period(g[r][0])].get(col(section, side, h))) for r in data)
                    if t is not None and abs(t - s) > 1:
                        notes.append(f'{sheet}: the Total row for {side} {h} is {t:,.2f} but the months add to {s:,.2f}.')
    # the report's own difference columns must equal 3B minus the other side
    for sheet, c0, heads, a, b in SELF_CHECK:
        if old_itc and sheet == 'ITC (Other than IMPG)':
            if c0 == 'Z': continue                       # no after-reversal block in the old layout
            a = ('ITC other', 'GSTR-3B net')             # its single block is already net of reversal
        g = _grid(wb[sheet]); hdr, data, _ = _locate(g)
        ci = column_index_from_string(c0) - 1
        for r in data:
            p = _period(g[r][0])
            for k, h in enumerate(heads):
                v = _num(g[r][ci + k])
                if v is None: continue
                exp = n0(rows[p].get(col(a[0], a[1], h))) - n0(rows[p].get(col(b[0], b[1], h)))
                if abs(v - exp) > 0.05:
                    notes.append(f'{sheet} {p} {h}: the report\'s own difference is {v:,.2f}, but {a[1]} minus {b[1]} is {exp:,.2f}.')
    wb.close()
    if len(rows) != 12:
        notes.append(f'{len(rows)} tax periods found; a full financial year has 12.')
    if meta['fy']:
        bad = [p for p in rows if fy_of(p) != meta['fy']]
        if bad: notes.append(f'Periods outside FY {meta["fy"]}: {", ".join(bad)}.')
    out = []
    for p in sorted(rows, key=period_key):
        d = {'Return Period': p, 'Financial Year': fy_of(p),
             'Report generated at': meta['generated'].strftime('%d-%m-%Y %H:%M') if meta['generated'] else None}
        for f in FIELDS: d[f] = rows[p].get(f)
        d['_period'], d['_fy'] = p, fy_of(p)
        d['_generated'] = meta['generated']
        d['_old_itc'] = old_itc
        d['_blank_cols'] = [f for f in FIELDS if d[f] is None]
        d['_row'] = None
        out.append(d)
    return meta, out, notes

# ------------------------------------------------------------------ ingest
def add_bo(ing, path, rec, batch):
    """Called by Ingest.add for a BO comparison report. A period already held is replaced only by a report
    GENERATED LATER — the upload order does not matter, because the report carries its own timestamp."""
    from gpe_core import REPORT_SPECS, diff_rows
    name = rec['file']
    try:
        meta, parsed, notes = read_bo(path)
    except Exception as e:
        rec.update(code=BO_CODE, title=REPORT_SPECS[BO_CODE]['title'], status='NOT RECOGNISED',
                   note=f'BO comparison report, but its layout could not be read: {e}. Nothing taken from it.')
        ing.register.append(rec); return rec
    rec.update(code=BO_CODE, title=REPORT_SPECS[BO_CODE]['title'], rows=len(parsed), declared=len(parsed),
               subject_gstin=meta['gstin'], detect_note='BO comparison report (title and sheet names)')
    rec['bo_meta'] = meta
    gen = meta['generated']
    head = (f'BO report for FY {meta["fy"]}, generated {gen:%d-%m-%Y %H:%M}. ' if gen else
            f'BO report for FY {meta["fy"]} — NO GENERATION TIME FOUND in it. ')
    if meta.get('old_itc'):
        head += ('Older layout: its ITC sheet gives GSTR-3B credit only net of the Table 4B reversal (one block), so the '
                 'before-reversal columns are blank. ')
    if ing.case_gstin and meta['gstin'] and meta['gstin'] != ing.case_gstin:
        rec.update(status='NOT USED - OTHER GSTIN',
                   note=head + f'WRONG TAXPAYER: the report is for GSTIN {meta["gstin"]}, not {ing.case_gstin}. '
                               'Nothing has been taken from it.')
        ing.register.append(rec); return rec
    ing._seen_md5[rec['md5']] = name
    bucket = ing.data.setdefault(BO_CODE, [])
    held = {r['_period']: i for i, r in enumerate(bucket)}
    added = replaced = kept_old = rec_changed = vals_changed = 0
    for x in parsed:
        row = dict(x, _file=name, _batch=batch)
        p = x['_period']
        if p not in held:
            held[p] = len(bucket); bucket.append(row); added += 1; continue
        prev = bucket[held[p]]
        g0, g1 = prev.get('_generated'), gen
        newer = (g1 is not None and (g0 is None or g1 > g0)) or (g1 == g0 and batch != prev.get('_batch'))
        if not newer:
            kept_old += 1
            continue
        ch = [c for c in diff_rows(prev, row) if c[0] != 'Report generated at']
        replaced += 1
        row['_supersedes'] = prev.get('_file')
        row['_history'] = dict(supersedes=prev.get('_file'), changes=ch, uncertain=(g1 == g0))
        if ch:
            rec_changed += 1; vals_changed += len(ch)
            for f, a0, b0 in ch:
                three_b = 'GSTR-3B' in f
                ing.changes.append(dict(
                    code=BO_CODE, title=REPORT_SPECS[BO_CODE]['title'], record=p, field=f, old=a0, new=b0,
                    old_file=prev.get('_file'), new_file=name, nature='CHANGED',
                    level='REVIEW' if three_b else 'NOTE',
                    reading=(f'A BO report generated {g1:%d-%m-%Y} replaced one generated {g0:%d-%m-%Y}. '
                             if g0 and g1 else 'Two BO reports for the same month. ')
                            + ('GSTR-3B cannot be revised, so a GSTR-3B figure that moved between two BO reports is unusual: '
                               'check the return on the portal.' if three_b else
                               'GSTR-1 amendments filed later, or GSTR-2B entries of a supplier, can move this figure. '
                               'The later report is used.'),
                    order_uncertain=(g1 == g0)))
        bucket[held[p]] = row
    rec.update(rows_new=added, rows_replaced=replaced, records_changed=rec_changed, values_changed=vals_changed)
    tail = []
    if replaced:
        tail.append(f'{replaced} month(s) were already held from an older BO report and have been REPLACED by this later one'
                    + (f'; {vals_changed} value(s) in {rec_changed} month(s) differ — see 09A_RESTATEMENTS.' if vals_changed else
                       '; no value differs.'))
    if kept_old:
        tail.append(f'{kept_old} month(s) are already held from a BO report generated LATER than this one, so this older '
                    f'report is NOT used for them.')
    if notes:
        tail.append('INTEGRITY: ' + ' '.join(notes))
    rec['detect_note'] = 'BO comparison report (recognised by its title and sheet names). ' + head
    rec['note'] = ' '.join(tail)
    rec['status'] = 'COUNTED' if (added or replaced) else 'SUPERSEDED - OLDER BO REPORT'
    ing.register.append(rec)
    return rec

# ------------------------------------------------------------------ figures per month
def _g(d, section, side, h):
    return n0(d.get(col(section, side, h)))

def bo_month(d):
    """Head-wise figures for one month of the BO report. Sign: positive = against the taxpayer."""
    f = {}
    for h in H4:
        exp3 = _g(d, 'Export-SEZ', 'GSTR-3B', h) if h in H2 else 0.0
        exp1 = _g(d, 'Export-SEZ', 'GSTR-1', h) if h in H2 else 0.0
        f[('out3b', h)] = _g(d, 'Outward', 'GSTR-3B', h) + exp3
        f[('outr1', h)] = _g(d, 'Outward', 'GSTR-1', h) + exp1
        f[('out_short', h)] = f[('outr1', h)] - f[('out3b', h)]
        f[('rcm3b', h)] = _g(d, 'RCM liability', 'GSTR-3B', h)
        f[('rcm2b', h)] = _g(d, 'RCM liability', 'GSTR-2B', h)
        f[('rcm_short', h)] = f[('rcm2b', h)] - f[('rcm3b', h)]
        f[('itc3b_net', h)] = _g(d, 'ITC other', 'GSTR-3B net', h)
        # old layout: no before-reversal figure was reported, so 'gross' is taken as the net figure (reversal shown as nil)
        f[('itc3b_gross', h)] = (_g(d, 'ITC other', 'GSTR-3B gross', h) if d.get(col('ITC other', 'GSTR-3B gross', h)) is not None
                                 else f[('itc3b_net', h)])
        f[('itc2b', h)] = _g(d, 'ITC other', 'GSTR-2B', h)
        f[('itc_excess', h)] = f[('itc3b_net', h)] - f[('itc2b', h)]
        f[('impg3b', h)] = _g(d, 'ITC import', 'GSTR-3B', h) if h in H2 else 0.0
        f[('impg2b', h)] = _g(d, 'ITC import', 'GSTR-2B', h) if h in H2 else 0.0
        f[('impg_excess', h)] = f[('impg3b', h)] - f[('impg2b', h)]
        f[('rcmitc', h)] = _g(d, 'ITC on RCM', 'GSTR-3B', h)
        f[('rcmitc_excess', h)] = f[('rcmitc', h)] - f[('rcm3b', h)]
        f[('drc03', h)] = _g(d, 'DRC-03', 'paid', h)
    s3 = lambda k: sum(f[(k, h)] for h in H3)
    f['out3b'] = s3('out3b'); f['outr1'] = s3('outr1'); f['short'] = s3('out_short')
    f['rcm3b'] = s3('rcm3b'); f['rcm2b'] = s3('rcm2b')
    f['itc_net'] = s3('itc3b_net') + f[('impg3b', 'IGST')] + s3('rcmitc')
    f['itc_gross'] = s3('itc3b_gross') + f[('impg3b', 'IGST')] + s3('rcmitc')
    f['reversal'] = s3('itc3b_gross') - s3('itc3b_net')
    f['excess'] = s3('itc_excess') + f[('impg_excess', 'IGST')]
    f['exp3b'] = f[('out3b', 'IGST')] - _g(d, 'Outward', 'GSTR-3B', 'IGST')
    f['expr1'] = f[('outr1', 'IGST')] - _g(d, 'Outward', 'GSTR-1', 'IGST')
    f['any3b'] = any(abs(_g(d, s, side, h)) > 0.5 for s, side in
                     (('Outward', 'GSTR-3B'), ('Export-SEZ', 'GSTR-3B'), ('RCM liability', 'GSTR-3B'),
                      ('ITC other', 'GSTR-3B gross'), ('ITC import', 'GSTR-3B'), ('ITC on RCM', 'GSTR-3B'))
                     for h in H4 if col(s, side, h) in d)
    f['anyr1'] = any(abs(f[('outr1', h)]) > 0.5 for h in H4)
    return f

def _due(p):
    _, y, m = parse_period(p)
    m += 1
    if m > 12: m, y = 1, y + 1
    return datetime.date(y, m, 20)

def _rs(v): return f'Rs.{inr(v)}'

def _explain(diff, cands):
    """A sentence when the difference is exactly one (or two) identifiable components, else None."""
    c = [(n, v) for n, v in cands if abs(v) > TOL]
    for n, v in c:
        if abs(abs(diff) - abs(v)) <= TOL:
            return f'The difference is exactly the {n} ({_rs(abs(v))}): one source counts it here and the other does not.'
    for i in range(len(c)):
        for j in range(i + 1, len(c)):
            if abs(abs(diff) - abs(c[i][1] + c[j][1])) <= TOL:
                return (f'The difference is exactly the {c[i][0]} plus the {c[j][0]} '
                        f'({_rs(abs(c[i][1] + c[j][1]))}): the two sources draw the line differently.')
    return None

MEASURES = [
    ('out3b',  'GSTR-3B outward tax, other than reverse charge (IGST+CGST+SGST)',
     'R3B TAX PAYABLE DETAILS: Outward Total GST', 'Outward GSTR-3B (3.1(a)+3.1.1(i)) + Export-SEZ GSTR-3B (3.1(b)) IGST'),
    ('rcm3b',  'GSTR-3B reverse-charge tax, Table 3.1(d) (IGST+CGST+SGST)',
     'TAX PAID DETAILS: GST OP Total minus R3B TAX PAYABLE: Outward Total GST', 'RCM liability GSTR-3B'),
    ('r1',     'GSTR-1 tax (IGST+CGST+SGST)',
     'R1 FILINGS IN THE YEAR: Total GST', 'Outward GSTR-1 + Export-SEZ GSTR-1 IGST'),
    ('itc',    'ITC availed in GSTR-3B (IGST+CGST+SGST)',
     'TAX PAID DETAILS: GST ITC Total', 'ITC other GSTR-3B net + ITC import GSTR-3B IGST + ITC on RCM'),
    ('itc_s',  'SGST ITC availed in GSTR-3B',
     'EXCESS ITC CLAIMED: SGSTITC Claimed In R3B', 'ITC other GSTR-3B net SGST + ITC on RCM SGST'),
    ('itc_i',  'IGST ITC availed in GSTR-3B',
     'EXCESS ITC CLAIMED: IGSTITC Claimed In R3B', 'ITC other GSTR-3B net IGST + ITC import IGST + ITC on RCM IGST'),
    ('av_s',   'SGST ITC available: Prime GSTR-2A against BO GSTR-2B',
     'EXCESS ITC CLAIMED: SGSTITC Available In R2A', 'ITC other GSTR-2B SGST'),
    ('av_i',   'IGST ITC available: Prime GSTR-2A against BO GSTR-2B',
     'EXCESS ITC CLAIMED: IGSTITC Available In R2A', 'ITC other GSTR-2B IGST + ITC import GSTR-2B IGST'),
    ('liab_m', 'GSTR-3B total tax liability, all heads incl. cess (outward + export + reverse charge)',
     'R3B FILINGS IN THE YEAR: Total Tax Liability', 'Outward + Export-SEZ + RCM liability, GSTR-3B, all heads'),
    ('itc_m',  'GSTR-3B net ITC claimed (IGST+CGST+SGST)',
     'R3B FILINGS IN THE YEAR: Net ITC Claimed', 'ITC other GSTR-3B net + ITC import GSTR-3B IGST + ITC on RCM'),
    ('short',  'RESULT — output tax short-paid: GSTR-1 minus GSTR-3B (IGST+CGST+SGST)',
     '17_R1_vs_R3B: GSTR-1 tax minus GSTR-3B outward tax', 'GSTR-1 minus GSTR-3B, outward and export-SEZ'),
    ('excess', 'RESULT — excess ITC: availed minus available (IGST+CGST+SGST)',
     '13_EXCESS_ITC: total excess over GSTR-2A (CGST taken = SGST)', 'ITC other GSTR-3B net minus GSTR-2B, plus import'),
]
MEASURE = {m[0]: m for m in MEASURES}

READING = {
    'out3b': ('GSTR-3B cannot be revised, so both sources should give the same figure. A difference means Prime and BO '
              'group the tables differently, or one of them holds an older copy of the return.'),
    'rcm3b': ('The notebook derives reverse-charge tax as GST OP Total minus outward tax. If OP Total carries anything else '
              '(cess, interest, late fee), the derived figure moves with it.'),
    'r1':    ('BO counts GSTR-1/IFF/1A FILED DURING the month, amendments (Tables 9A, 9C, 10, 11) included. The Prime report '
              'gives the return for the period. A difference usually means an amendment or an advance adjusted in a later '
              'return, or a GSTR-1 filed late.'),
    'itc':   ('Prime\'s ITC total may be before or after the Table 4B reversal; BO gives both. Where Prime agrees with the '
              'figure before reversal, the reversal is the whole difference.'),
    'itc_s': 'As for ITC availed, head by head.',
    'itc_i': 'As for ITC availed, head by head. Import IGST (Table 4A(1)) and IGST on reverse charge are included on the BO side.',
    'av_s':  ('DIFFERENT STATEMENTS: GSTR-2A is dynamic and moves whenever a supplier files late or amends; GSTR-2B is fixed '
              'for the month. From 1 January 2022 credit is restricted to what GSTR-2B shows (section 16(2)(aa), Rule 36(4)), '
              'so the BO figure is the one that counts.'),
    'av_i':  'As for SGST: GSTR-2A against GSTR-2B. GSTR-2B governs from 1 January 2022.',
    'liab_m': ('A second GST-Prime route to the GSTR-3B liability. Prime rounds each head; a larger difference usually means '
               'interest, late fee or cess counted on one side only, or a different copy of the return.'),
    'itc_m':  'A second GST-Prime route to the ITC availed. As for ITC availed above.',
    'short': ('Result of the two GSTR-1 and GSTR-3B lines above; any difference there carries through.'),
    'excess':('The notebook measures excess credit against GSTR-2A (and takes CGST equal to SGST); BO measures it head by head '
              'against GSTR-2B. GSTR-2B governs from 1 January 2022, so the BO figure is the one that counts.'),
}

def chk_bo(ing, master, r1res=None, exres=None):
    if not ing.have(BO_CODE):
        return dict(available=False, missing=[BO_CODE])
    bo = ing.by_period(BO_CODE)
    mrow = {m['period']: m for m in master}
    tp, tx, r1, ex, fl, fm = (ing.by_period(c) for c in ('TAX_PAID', 'R3B_TAX_PAYABLE', 'R1_MONTHLY', 'EXCESS_ITC', 'R3B_FILING',
                                                        'R3B_MONTHLY'))
    r1map = {r['period']: r for r in (r1res or {}).get('rows', [])} if (r1res or {}).get('available') else {}
    exmap = {r['period']: r for r in (exres or {}).get('rows', [])} if (exres or {}).get('available') else {}
    months, cmp_rows = [], []
    for p in sorted(bo, key=period_key):
        d = bo[p]; f = bo_month(d); gen = d.get('_generated')
        m = mrow.get(p)
        filed = (fl.get(p) or {}).get('Date Of Filing') or (fm.get(p) or {}).get('Date of Filing')
        prime_3b = (bool(m) and (abs(m['turnover']) > 0.5 or abs(m['out_total']) > 0.5 or abs(m['itc']) > 0.5)) or \
                   any(abs(n0((fm.get(p) or {}).get(k))) > 0.5 for k in ('Total Tax Liability', 'Net ITC Claimed'))
        covered, why = True, ''
        if gen is not None and isinstance(filed, (datetime.date, datetime.datetime)) and \
                (filed.date() if isinstance(filed, datetime.datetime) else filed) > gen.date():
            covered, why = False, (f'BO report generated {gen:%d-%m-%Y}, before GSTR-3B was filed on '
                                   f'{filed:%d-%m-%Y}. Download a fresh BO report for FY {d["_fy"]}.')
        elif gen is not None and gen.date() < _due(p) and not f['any3b']:
            covered, why = False, (f'BO report generated {gen:%d-%m-%Y}, before GSTR-3B for {p} fell due '
                                   f'({_due(p):%d-%m-%Y}). Download a fresh BO report for FY {d["_fy"]}.')
        elif not f['any3b'] and prime_3b:
            covered, why = False, ('BO shows no GSTR-3B for the month while Prime shows a return. The return was most likely filed '
                                   'after the BO report was generated. Download a fresh BO report, then compare again.')
        months.append(dict(period=p, fy=d['_fy'], generated=gen, file=d.get('_file'), covered=covered, why=why,
                           f=f, in_prime=bool(m), no_3b=not f['any3b']))
        # ---- comparison, measure by measure
        P = {}
        if p in tx and (tx[p].get('Outward Total GST') is not None):
            P['out3b'] = n0(tx[p].get('Outward Total GST'))
            heads_p = {h: n0(tx[p].get(f'Outward {h}')) for h in H3}
        if p in tx and p in tp and tp[p].get('GST OP Total') is not None and tx[p].get('Outward Total GST') is not None:
            P['rcm3b'] = n0(tp[p].get('GST OP Total')) - n0(tx[p].get('Outward Total GST'))
        if p in r1 and r1[p].get('Total GST') is not None:
            P['r1'] = n0(r1[p].get('Total GST'))
        if p in tp and tp[p].get('GST ITC Total') is not None:
            P['itc'] = n0(tp[p].get('GST ITC Total'))
        if p in ex:
            e = ex[p]
            if e.get('SGSTITC Claimed In R3B') is not None: P['itc_s'] = n0(e.get('SGSTITC Claimed In R3B'))
            if e.get('IGSTITC Claimed In R3B') is not None: P['itc_i'] = n0(e.get('IGSTITC Claimed In R3B'))
            if e.get('SGSTITC Available In R2A') is not None: P['av_s'] = n0(e.get('SGSTITC Available In R2A'))
            if e.get('IGSTITC Available In R2A') is not None: P['av_i'] = n0(e.get('IGSTITC Available In R2A'))
        rm = ing.by_period('R3B_MONTHLY').get(p) if ing.have('R3B_MONTHLY') else None
        if rm:
            if rm.get('Total Tax Liability') is not None: P['liab_m'] = n0(rm.get('Total Tax Liability'))
            if rm.get('Net ITC Claimed') is not None: P['itc_m'] = n0(rm.get('Net ITC Claimed'))
        if p in r1map and r1map[p].get('diff') is not None and 'out3b' in P:
            P['short'] = r1map[p]['diff']
        if p in exmap and p in ex:
            P['excess'] = exmap[p]['total']
        B = {'out3b': f['out3b'], 'rcm3b': f['rcm3b'], 'r1': f['outr1'], 'itc': f['itc_net'],
             'itc_s': f[('itc3b_net', 'SGST')] + f[('rcmitc', 'SGST')],
             'itc_i': f[('itc3b_net', 'IGST')] + f[('impg3b', 'IGST')] + f[('rcmitc', 'IGST')],
             'av_s': f[('itc2b', 'SGST')], 'av_i': f[('itc2b', 'IGST')] + f[('impg2b', 'IGST')],
             'short': f['short'], 'excess': f['excess'],
             'liab_m': sum(f[('out3b', h)] + f[('rcm3b', h)] for h in H4), 'itc_m': f['itc_net']}
        ALT = {} if d.get('_old_itc') else {'itc': f['itc_gross'], 'itc_m': f['itc_gross'],
               'itc_s': f[('itc3b_gross', 'SGST')] + f[('rcmitc', 'SGST')],
               'itc_i': f[('itc3b_gross', 'IGST')] + f[('impg3b', 'IGST')] + f[('rcmitc', 'IGST')]}
        CAND = {'out3b': [('export / SEZ tax of Table 3.1(b)', f['exp3b']), ('reverse-charge tax of Table 3.1(d)', f['rcm3b'])],
                'rcm3b': [('export / SEZ tax of Table 3.1(b)', f['exp3b'])],
                'r1': [('export / SEZ tax in GSTR-1', f['expr1'])],
                'itc': [('Table 4B reversal', f['reversal']), ('import IGST of Table 4A(1)', f[('impg3b', 'IGST')]),
                        ('ITC on reverse charge of Table 4A(2)+4A(3)', sum(f[('rcmitc', h)] for h in H3))],
                'itc_s': [('Table 4B reversal (SGST)', f[('itc3b_gross', 'SGST')] - f[('itc3b_net', 'SGST')]),
                          ('SGST ITC on reverse charge', f[('rcmitc', 'SGST')])],
                'itc_i': [('Table 4B reversal (IGST)', f[('itc3b_gross', 'IGST')] - f[('itc3b_net', 'IGST')]),
                          ('import IGST of Table 4A(1)', f[('impg3b', 'IGST')]), ('IGST ITC on reverse charge', f[('rcmitc', 'IGST')])],
                'av_i': [('import IGST in GSTR-2B (IMPG)', f[('impg2b', 'IGST')])],
                'liab_m': [('cess', sum(f[(k, 'CESS')] for k in ('out3b', 'rcm3b'))), ('reverse-charge tax of Table 3.1(d)', f['rcm3b']),
                           ('export / SEZ tax of Table 3.1(b)', f['exp3b'])],
                'itc_m': [('Table 4B reversal', f['reversal']), ('import IGST of Table 4A(1)', f[('impg3b', 'IGST')]),
                          ('ITC on reverse charge of Table 4A(2)+4A(3)', sum(f[('rcmitc', h)] for h in H3))],
                'short': [('export / SEZ tax', f['exp3b'] - f['expr1'])],
                'excess': [('CGST, which the notebook takes as equal to SGST', f[('itc_excess', 'CGST')] - f[('itc_excess', 'SGST')]),
                           ('Table 4B reversal, which BO deducts from the credit availed', f['reversal'])]}
        for mid, label, psrc, bsrc in MEASURES:
            pv, bv = P.get(mid), B[mid]
            alt = ALT.get(mid)
            diff = None if pv is None else bv - pv
            if not covered:
                status, taken = 'BO NOT USED', 'NOTEBOOK'
            elif pv is None:
                status, taken = 'BO ONLY', 'BO'
            elif abs(diff) <= TOL:
                status, taken = 'MATCH', 'BO'
            elif alt is not None and abs(alt - pv) <= TOL:
                status, taken = 'MATCH (before reversal)', 'BO'
            elif mid in ('av_s', 'av_i'):
                status, taken = 'DIFFERENT STATEMENT (2A vs 2B)', 'BO'
            else:
                status, taken = 'DIFFERS', 'BO'
            reading = ''
            if not covered:
                reading = why
            elif pv is None:
                reading = 'The Prime report for this figure is not held for the month. BO figure taken.'
            elif status == 'MATCH (before reversal)':
                reading = (f'Prime agrees with the BO figure BEFORE the Table 4B reversal; the reversal of '
                           f'{_rs(alt - bv)} is the whole difference. BO\'s figure after reversal is taken.')
            elif status != 'MATCH':
                ex_ = _explain(diff, CAND.get(mid, []))
                reading = (ex_ + ' ' if ex_ else '') + READING[mid] + ' BO figure taken.'
                if d.get('_old_itc') and mid in ('itc', 'itc_m', 'itc_s', 'itc_i', 'excess'):
                    reading += (' This BO report (older layout) gives GSTR-3B credit only NET of the Table 4B reversal, '
                                'reversals for reverse charge and imports included, so a figure before reversal cannot be '
                                'compared.')
            elif mid == 'out3b' and 'out3b' in P:
                bad = [h for h in H3 if abs(heads_p[h] - f[('out3b', h)]) > TOL]
                if bad:
                    status = 'DIFFERS (HEADS)'
                    reading = ('The total agrees but the split between heads does not: '
                               + '; '.join(f'{h} Prime {_rs(heads_p[h])}, BO {_rs(f[("out3b", h)])}' for h in bad)
                               + '. Tax paid under the wrong head is not a payment of the right one (sections 77 IGST / 19 '
                                 'IGST Act). BO figure taken.')
            cmp_rows.append(dict(period=p, fy=d['_fy'], mid=mid, measure=label, prime=pv, prime_src=psrc,
                                 bo=bv, bo_src=bsrc, alt=alt, diff=diff, status=status, taken=taken, reading=reading))
    covered_p = {x['period'] for x in months if x['covered']}
    # ---- financial-year final position, covered months only
    fyres = {}
    for x in months:
        if not x['covered']: continue
        a = fyres.setdefault(x['fy'], dict(months=[], generated=set(), files=set()))
        a['months'].append(x['period']); a['files'].add(x['file'])
        if x['generated']: a['generated'].add(x['generated'])
        f = x['f']
        for k in ('out_short', 'rcm_short', 'drc03', 'itc_excess', 'impg_excess', 'rcmitc_excess',
                  'out3b', 'outr1', 'rcm3b', 'rcm2b', 'itc3b_net', 'itc2b', 'impg3b', 'impg2b', 'rcmitc'):
            for h in H4:
                a[(k, h)] = a.get((k, h), 0.0) + f[(k, h)]
                if k in ('out_short', 'rcm_short', 'itc_excess', 'impg_excess', 'rcmitc_excess'):
                    a[(k + '_gross', h)] = a.get((k + '_gross', h), 0.0) + max(0.0, f[(k, h)])
        a.setdefault('short_months', [])
        if any(f[('out_short', h)] > 1 for h in H4): a['short_months'].append((x['period'], 'outward',
                                                                               sum(max(0, f[('out_short', h)]) for h in H4)))
        if any(f[('rcm_short', h)] > 1 for h in H4): a['short_months'].append((x['period'], 'reverse charge',
                                                                               sum(max(0, f[('rcm_short', h)]) for h in H4)))
        a.setdefault('neg_r1', [])
        if sum(f[('outr1', h)] for h in H4) < -TOL:
            a['neg_r1'].append((x['period'], sum(f[('outr1', h)] for h in H4), sum(f[('out3b', h)] for h in H4)))
        a.setdefault('excess_months', [])
        if any(f[('itc_excess', h)] > 1 for h in H4) or any(f[('impg_excess', h)] > 1 for h in H2) \
                or any(f[('rcmitc_excess', h)] > 1 for h in H4):
            a['excess_months'].append((x['period'], sum(max(0, f[(k, h)]) for k in ('itc_excess', 'impg_excess', 'rcmitc_excess')
                                                       for h in H4)))
    for fy, a in fyres.items():
        a['liab_due'] = {h: max(0.0, a[('out_short', h)] + a[('rcm_short', h)] - a[('drc03', h)]) for h in H4}
        a['liab_due_total'] = sum(a['liab_due'].values())
        a['out_net'] = sum(a[('out_short', h)] for h in H4)
        a['rcm_net'] = sum(a[('rcm_short', h)] for h in H4)
        a['drc03_total'] = sum(a[('drc03', h)] for h in H4)
        a['liab_gross'] = sum(a[('out_short_gross', h)] + a[('rcm_short_gross', h)] for h in H4)
        a['itc_due'] = {h: max(0.0, a[('itc_excess', h)]) + max(0.0, a[('impg_excess', h)]) + max(0.0, a[('rcmitc_excess', h)])
                        for h in H4}
        a['itc_due_total'] = sum(a['itc_due'].values())
        a['itc_gross'] = sum(a[(k + '_gross', h)] for k in ('itc_excess', 'impg_excess', 'rcmitc_excess') for h in H4)
        a['rcmitc_due'] = sum(max(0.0, a[('rcmitc_excess', h)]) for h in H4)
        a['impg_due'] = sum(max(0.0, a[('impg_excess', h)]) for h in H4)
        a['generated'] = sorted(a['generated']); a['files'] = sorted(a['files'])
        a['pre_2022'] = [p for p in a['months'] if period_key(p) < (2022, 1)]
        # the notebook's own figures for the same months, for the change report
        a['prime_short'] = sum(r1map[p]['short'] for p in a['months'] if p in r1map)
        a['prime_short_months'] = sum(1 for p in a['months'] if p in r1map)
        a['prime_excess'] = sum(exmap[p]['total'] for p in a['months'] if p in exmap)
        a['prime_excess_months'] = sum(1 for p in a['months'] if p in exmap)
    st = {}
    for c in cmp_rows: st.setdefault(c['mid'], {}).setdefault(c['status'], 0); st[c['mid']][c['status']] += 1
    differs = [c for c in cmp_rows if c['status'] in ('DIFFERS', 'DIFFERS (HEADS)')]
    files = []
    for r in ing.register:
        if r.get('code') == BO_CODE and r.get('bo_meta'):
            mm = r['bo_meta']
            files.append(dict(file=r['file'], status=r['status'], fy=mm['fy'], generated=mm['generated'],
                              gstin=mm['gstin'], legal=mm['legal'], trade=mm['trade'], note=r.get('note', '')))
    prime_periods = [m['period'] for m in master]
    return dict(available=True, months=months, cmp=cmp_rows, fy=fyres, files=files, stats=st, differs=differs,
                covered=covered_p, not_covered=[x for x in months if not x['covered']],
                prime_without_bo=[p for p in prime_periods if p not in bo],
                liab_due=sum(a['liab_due_total'] for a in fyres.values()),
                itc_due=sum(a['itc_due_total'] for a in fyres.values()),
                fys=sorted(fyres, key=lambda f: int(f.split('-')[0])))

# ------------------------------------------------------------------ workbook sheets
def _t(section, side, heads, sign=1):
    return [(sign, col(section, side, h)) for h in heads]

BO_TERMS = {
    'out3b': _t('Outward', 'GSTR-3B', H3) + _t('Export-SEZ', 'GSTR-3B', ('IGST',)),
    'rcm3b': _t('RCM liability', 'GSTR-3B', H3),
    'r1': _t('Outward', 'GSTR-1', H3) + _t('Export-SEZ', 'GSTR-1', ('IGST',)),
    'itc': _t('ITC other', 'GSTR-3B net', H3) + _t('ITC import', 'GSTR-3B', ('IGST',)) + _t('ITC on RCM', 'GSTR-3B', H3),
    'itc_s': _t('ITC other', 'GSTR-3B net', ('SGST',)) + _t('ITC on RCM', 'GSTR-3B', ('SGST',)),
    'itc_i': _t('ITC other', 'GSTR-3B net', ('IGST',)) + _t('ITC import', 'GSTR-3B', ('IGST',)) + _t('ITC on RCM', 'GSTR-3B', ('IGST',)),
    'av_s': _t('ITC other', 'GSTR-2B', ('SGST',)),
    'av_i': _t('ITC other', 'GSTR-2B', ('IGST',)) + _t('ITC import', 'GSTR-2B', ('IGST',)),
    'excess': _t('ITC other', 'GSTR-3B net', H3) + _t('ITC other', 'GSTR-2B', H3, -1)
              + _t('ITC import', 'GSTR-3B', ('IGST',)) + _t('ITC import', 'GSTR-2B', ('IGST',), -1),
}
BO_TERMS['short'] = BO_TERMS['r1'] + [(-s, f) for s, f in BO_TERMS['out3b']]
BO_TERMS['liab_m'] = _t('Outward', 'GSTR-3B', H4) + _t('Export-SEZ', 'GSTR-3B', H2) + _t('RCM liability', 'GSTR-3B', H4)
BO_TERMS['itc_m'] = BO_TERMS['itc']
ALT_TERMS = {
    'itc_m': _t('ITC other', 'GSTR-3B gross', H3) + _t('ITC import', 'GSTR-3B', ('IGST',)) + _t('ITC on RCM', 'GSTR-3B', H3),
    'itc': _t('ITC other', 'GSTR-3B gross', H3) + _t('ITC import', 'GSTR-3B', ('IGST',)) + _t('ITC on RCM', 'GSTR-3B', H3),
    'itc_s': _t('ITC other', 'GSTR-3B gross', ('SGST',)) + _t('ITC on RCM', 'GSTR-3B', ('SGST',)),
    'itc_i': _t('ITC other', 'GSTR-3B gross', ('IGST',)) + _t('ITC import', 'GSTR-3B', ('IGST',)) + _t('ITC on RCM', 'GSTR-3B', ('IGST',)),
}
SHEETS = ('19E_BO_vs_NOTEBOOK', '19F_BO_FINAL_POSITION', '19G_CHANGE_REPORT')

def _src_ref(b):
    """period -> row on SRC_BO_CMP (same order build_sources writes), and field -> column letter."""
    rows = sorted(b.ing.data.get(BO_CODE, []), key=lambda r: period_key(r['_period']))
    prow = {r['_period']: i for i, r in enumerate(rows, 2)}
    cl = {f: b.src_col(BO_CODE, f) for f in FIELDS}
    return prow, cl

def _sum_formula(terms, row, cl, sheet="'SRC_BO_CMP'"):
    s = ''.join(f"{'+' if sg > 0 else '-'}{sheet}!{cl[f]}{row}" for sg, f in terms)
    return '=' + (s[1:] if s.startswith('+') else s)

def _prime_formula(b, mid, ref):
    def L(code, name):
        v = b.lookup(code, name, 'Return Period', ref)
        return v[1:] if isinstance(v, str) else '0'
    if mid == 'out3b':  return '=' + L('R3B_TAX_PAYABLE', 'Outward Total GST')
    if mid == 'rcm3b':  return f"={L('TAX_PAID', 'GST OP Total')}-{L('R3B_TAX_PAYABLE', 'Outward Total GST')}"
    if mid == 'r1':     return '=' + L('R1_MONTHLY', 'Total GST')
    if mid == 'itc':    return '=' + L('TAX_PAID', 'GST ITC Total')
    if mid == 'itc_s':  return '=' + L('EXCESS_ITC', 'SGSTITC Claimed In R3B')
    if mid == 'itc_i':  return '=' + L('EXCESS_ITC', 'IGSTITC Claimed In R3B')
    if mid == 'av_s':   return '=' + L('EXCESS_ITC', 'SGSTITC Available In R2A')
    if mid == 'av_i':   return '=' + L('EXCESS_ITC', 'IGSTITC Available In R2A')
    if mid == 'liab_m': return '=' + L('R3B_MONTHLY', 'Total Tax Liability')
    if mid == 'itc_m':  return '=' + L('R3B_MONTHLY', 'Net ITC Claimed')
    if mid == 'short':  return f"={L('R1_MONTHLY', 'Total GST')}-{L('R3B_TAX_PAYABLE', 'Outward Total GST')}"
    if mid == 'excess': return f"=2*{L('EXCESS_ITC', 'SGSTExcess Claimed')}+{L('EXCESS_ITC', 'IGSTExcess Claimed')}"
    return ''

STATUSES = ['MATCH', 'MATCH (before reversal)', 'DIFFERS', 'DIFFERS (HEADS)', 'DIFFERENT STATEMENT (2A vs 2B)',
            'BO ONLY', 'BO NOT USED']

def build_compare(b):
    from gpe_build import BOLD, BD, NUM, MISSING_FILL
    from openpyxl.styles import Alignment
    bo = b.R.get('bo') or {}
    ws = b.sheet(SHEETS[0])
    if not bo.get('available'):
        return b.unavailable(ws, bo, 'GST-PRIME FIGURES COMPARED WITH THE BO COMPARISON REPORT')
    prow, cl = _src_ref(b)
    ws['A1'] = 'THE NOTEBOOK\'S FIGURES (GST-PRIME) COMPARED WITH THE BO COMPARISON REPORT — month by month'; ws['A1'].font = BOLD
    ws['A2'] = ('The BO report is the department\'s own comparison and is taken as final wherever it covers a month. '
                'Every difference is listed with the reason where the figures show it. Differences up to Rs.5 are rounding '
                '(Prime reports whole rupees, BO paise).')
    ws['A2'].alignment = Alignment(wrap_text=True, vertical='top')
    r = 4
    ws.cell(row=r, column=1, value='BO REPORTS HELD').font = BOLD; r += 1
    H = ['File', 'FY', 'GSTIN in the report', 'Legal / trade name', 'Generated at', 'Months used', 'Months NOT used', 'Register note']
    b.head(ws, r, H); r += 1
    used = {}
    for x in bo['months']:
        used.setdefault(x['file'], [0, []])
        if x['covered']: used[x['file']][0] += 1
        else: used[x['file']][1].append(x['period'])
    for fr in bo['files']:
        u = used.get(fr['file'], [0, []])
        vals = [fr['file'], fr['fy'], fr['gstin'], f"{fr['legal'] or ''} / {fr['trade'] or ''}",
                fr['generated'].strftime('%d-%m-%Y %H:%M') if fr['generated'] else 'not stated', u[0],
                ', '.join(u[1]) or '—', fr['status'] + ('. ' + fr['note'] if fr['note'] else '')]
        for j, v in enumerate(vals, 1):
            c = ws.cell(row=r, column=j, value=v); c.border = BD; c.alignment = Alignment(wrap_text=True, vertical='top')
        if u[1] and b.mark: ws.cell(row=r, column=7).fill = MISSING_FILL
        r += 1
    if bo['prime_without_bo']:
        r += 1
        r = b.note(ws, r, 'Months in the GST-Prime data with NO BO report: ' + ', '.join(bo['prime_without_bo'])
                   + '. Their figures rest on GST-Prime alone. Download the BO report for those financial years.', bold=True)
        if b.mark: ws.cell(row=r - 1, column=1).fill = MISSING_FILL
    # ---- summary by measure (counts are formulas over the detail table below)
    r += 1
    ws.cell(row=r, column=1, value='SUMMARY BY MEASURE').font = BOLD; r += 1
    sh = r
    SH = ['Measure', 'Months listed'] + STATUSES + ['Sum of differences (BO − notebook)', 'Largest difference (+)', 'Largest difference (−)']
    b.head(ws, sh, SH); r += 1
    s0 = r; r += len(MEASURES)
    r += 1
    ws.cell(row=r, column=1, value='EVERY COMPARISON').font = BOLD; r += 1
    hd = r
    DH = ['No.', 'Return Period', 'FY', 'Measure', 'Notebook (GST-Prime)', 'Notebook source', 'BO', 'BO before reversal',
          'BO source', 'Difference (BO − notebook)', 'Status', 'Figure taken', 'Reading']
    b.head(ws, hd, DH, [8, 13, 9, 44, 17, 34, 17, 15, 40, 17, 22, 11, 90])
    d0 = hd + 1
    d1 = d0 + len(bo['cmp']) - 1
    for i, c in enumerate(bo['cmp']):
        rr = d0 + i; p = c['period']; ref = f'$B{rr}'
        ws.cell(row=rr, column=1, value=i + 1)
        ws.cell(row=rr, column=2, value=p); ws.cell(row=rr, column=3, value=c['fy'])
        ws.cell(row=rr, column=4, value=c['measure'])
        ws.cell(row=rr, column=5, value=_prime_formula(b, c['mid'], ref) if c['prime'] is not None else '')
        ws.cell(row=rr, column=6, value=c['prime_src'] if c['prime'] is not None else 'not held for this month')
        ws.cell(row=rr, column=7, value=_sum_formula(BO_TERMS[c['mid']], prow[p], cl))
        ws.cell(row=rr, column=8, value=_sum_formula(ALT_TERMS[c['mid']], prow[p], cl) if c['alt'] is not None else '')
        ws.cell(row=rr, column=9, value=c['bo_src'])
        ws.cell(row=rr, column=10, value=f'=IF(E{rr}="","",G{rr}-E{rr})')
        if c['status'] in ('BO NOT USED', 'DIFFERS (HEADS)'):
            st = c['status']
        else:
            other = '"DIFFERENT STATEMENT (2A vs 2B)"' if c['mid'] in ('av_s', 'av_i') else '"DIFFERS"'
            st = (f'=IF(E{rr}="","BO ONLY",IF(ABS(J{rr})<={TOL:g},"MATCH",'
                  f'IF(AND(H{rr}<>"",ABS(N(H{rr})-E{rr})<={TOL:g}),"MATCH (before reversal)",{other})))')
        ws.cell(row=rr, column=11, value=st)
        ws.cell(row=rr, column=12, value=c['taken'])
        ws.cell(row=rr, column=13, value=c['reading'])
        if b.mark and c['status'] in ('DIFFERS', 'DIFFERS (HEADS)', 'BO NOT USED'):
            ws.cell(row=rr, column=11).fill = MISSING_FILL
    b.box(ws, d0, max(d0, d1), 1, 13, numfrom=5)
    for rr in range(d0, d1 + 1):
        for cc in (2, 3, 4, 6, 9, 11, 12, 13):
            ws.cell(row=rr, column=cc).number_format = 'General'
        for cc in (4, 6, 9, 13):
            ws.cell(row=rr, column=cc).alignment = Alignment(wrap_text=True, vertical='top')
    D, K, J = f'$D${d0}:$D${d1}', f'$K${d0}:$K${d1}', f'$J${d0}:$J${d1}'
    for k, (mid, label, _, _) in enumerate(MEASURES):
        rr = s0 + k
        ws.cell(row=rr, column=1, value=label)
        ws.cell(row=rr, column=2, value=f'=COUNTIF({D},$A{rr})')
        for j, stt in enumerate(STATUSES, 3):
            ws.cell(row=rr, column=j, value=f'=COUNTIFS({D},$A{rr},{K},"{stt}")')
        n = 3 + len(STATUSES)
        ws.cell(row=rr, column=n, value=f'=SUMIFS({J},{D},$A{rr})')
        ws.cell(row=rr, column=n + 1, value=f'=MAX(0,_xlfn.MAXIFS({J},{D},$A{rr}))')
        ws.cell(row=rr, column=n + 2, value=f'=MIN(0,_xlfn.MINIFS({J},{D},$A{rr}))')
    b.box(ws, s0, s0 + len(MEASURES) - 1, 1, len(SH), numfrom=2)
    for rr in range(s0, s0 + len(MEASURES)):
        ws.cell(row=rr, column=1).alignment = Alignment(wrap_text=True, vertical='top')
    b.bo_cmp_range = (d0, d1)
    n = d1 + 2
    n = b.note(ws, n, 'HOW TO READ THIS SHEET', bold=True)
    for t in ('Notebook value: a live lookup into the GST-Prime SRC_ sheet named in "Notebook source". BO value: the sum of '
              'the SRC_BO_CMP cells named in "BO source" — click the cell to see them.',
              'MATCH — within Rs.5. MATCH (before reversal) — Prime gives ITC before the Table 4B reversal; BO\'s figure after '
              'reversal is the one taken. DIFFERS — a real difference, explained in "Reading" where the figures show the cause. '
              'DIFFERS (HEADS) — the total agrees but the split between IGST, CGST and SGST does not. DIFFERENT STATEMENT — '
              'GSTR-2A against GSTR-2B, which are not meant to agree. BO ONLY — no Prime figure held. BO NOT USED — the BO '
              'report was generated before that month\'s return was filed or fell due, so the notebook\'s figure is kept.',
              'Wherever the BO report covers a month, the BO figure is the one used in 19F, in 20_FINDINGS and in the report. '
              'The notebook\'s own figure stays on its sheet (13, 17 ...) for comparison and is not used for any amount.'):
        n = b.note(ws, n, t)
    ws.column_dimensions['A'].width = 44

# ---- 19F month table: (header, terms) ; terms None for the derived positive-sum columns
def _month_cols():
    C = [('GSTR-1 tax, all heads (outward + export-SEZ)', _t('Outward', 'GSTR-1', H4) + _t('Export-SEZ', 'GSTR-1', H2)),
         ('GSTR-3B tax, all heads (outward + export-SEZ)', _t('Outward', 'GSTR-3B', H4) + _t('Export-SEZ', 'GSTR-3B', H2))]
    for h in H4:
        C.append((f'Outward shortfall {h} (GSTR-1 − GSTR-3B)',
                  _t('Outward', 'GSTR-1', (h,)) + (_t('Export-SEZ', 'GSTR-1', (h,)) if h in H2 else [])
                  + _t('Outward', 'GSTR-3B', (h,), -1) + (_t('Export-SEZ', 'GSTR-3B', (h,), -1) if h in H2 else [])))
    C += [('Reverse charge in GSTR-2B, all heads', _t('RCM liability', 'GSTR-2B', H4)),
          ('Reverse charge declared in GSTR-3B 3.1(d), all heads', _t('RCM liability', 'GSTR-3B', H4))]
    for h in H4:
        C.append((f'RCM shortfall {h} (GSTR-2B − GSTR-3B)', _t('RCM liability', 'GSTR-2B', (h,)) + _t('RCM liability', 'GSTR-3B', (h,), -1)))
    for h in H4:
        C.append((f'DRC-03 paid {h}', _t('DRC-03', 'paid', (h,))))
    C += [('ITC availed, GSTR-3B net of reversal (other than import & RCM), all heads', _t('ITC other', 'GSTR-3B net', H4)),
          ('ITC in GSTR-2B (other than import & RCM), all heads', _t('ITC other', 'GSTR-2B', H4))]
    for h in H4:
        C.append((f'ITC excess {h} (GSTR-3B net − GSTR-2B)', _t('ITC other', 'GSTR-3B net', (h,)) + _t('ITC other', 'GSTR-2B', (h,), -1)))
    for h in H2:
        C.append((f'Import ITC excess {h} (GSTR-3B 4A(1) − GSTR-2B IMPG)', _t('ITC import', 'GSTR-3B', (h,)) + _t('ITC import', 'GSTR-2B', (h,), -1)))
    for h in H4:
        C.append((f'ITC on reverse charge over RCM tax declared {h}', _t('ITC on RCM', 'GSTR-3B', (h,)) + _t('RCM liability', 'GSTR-3B', (h,), -1)))
    C += [('Month\'s liability shortfall (positive heads only)', None), ('Month\'s ITC excess (positive heads only)', None)]
    return C

def build_position(b):
    from gpe_build import BOLD, BD, NUM, MISSING_FILL
    from openpyxl.styles import Alignment, PatternFill
    bo = b.R.get('bo') or {}
    ws = b.sheet(SHEETS[1])
    if not bo.get('available'):
        return b.unavailable(ws, bo, 'FINAL POSITION AS PER THE BO COMPARISON REPORT')
    prow, cl = _src_ref(b)
    ws['A1'] = 'FINAL POSITION AS PER THE BO COMPARISON REPORT — the figures the findings rest on'; ws['A1'].font = BOLD
    ws['A2'] = ('Every figure is a formula over SRC_BO_CMP. Heads are kept apart: a shortfall in one head is not met by an excess '
                'in another. Within a year a month\'s shortfall is set off against a later month\'s excess payment (that is only '
                'timing — interest under section 50 still runs), and DRC-03 payments are set off as the BO report does.')
    ws['A2'].alignment = Alignment(wrap_text=True, vertical='top')
    C = _month_cols()
    H = ['Return Period', 'FY', 'Used for the final figures?', 'Why not used'] + [c[0] for c in C]
    hr = 4
    b.head(ws, hr, H, [13, 9, 12, 40] + [15] * len(C))
    ws.row_dimensions[hr].height = 75
    months = bo['months']
    m0 = hr + 1; m1 = m0 + len(months) - 1
    L = {c[0]: get_column_letter(5 + k) for k, c in enumerate(C)}
    out_cols = [L[f'Outward shortfall {h} (GSTR-1 − GSTR-3B)'] for h in H4]
    rcm_cols = [L[f'RCM shortfall {h} (GSTR-2B − GSTR-3B)'] for h in H4]
    exc_cols = ([L[f'ITC excess {h} (GSTR-3B net − GSTR-2B)'] for h in H4]
                + [L[f'Import ITC excess {h} (GSTR-3B 4A(1) − GSTR-2B IMPG)'] for h in H2]
                + [L[f'ITC on reverse charge over RCM tax declared {h}'] for h in H4])
    for i, x in enumerate(months):
        rr = m0 + i
        ws.cell(row=rr, column=1, value=x['period']); ws.cell(row=rr, column=2, value=x['fy'])
        c = ws.cell(row=rr, column=3, value='YES' if x['covered'] else 'NO')
        if not x['covered'] and b.mark: c.fill = MISSING_FILL
        ws.cell(row=rr, column=4, value=x['why']).alignment = Alignment(wrap_text=True, vertical='top')
        for k, (hname, terms) in enumerate(C):
            if terms is None:
                cols = out_cols + rcm_cols if 'liability' in hname else exc_cols
                v = '=' + '+'.join(f'MAX(0,{cc}{rr})' for cc in cols)
            else:
                v = _sum_formula(terms, prow[x['period']], cl)
            ws.cell(row=rr, column=5 + k, value=v)
    b.box(ws, m0, m1, 1, len(H), numfrom=5)
    # ---- financial-year position
    r = m1 + 3
    ws.cell(row=r, column=1, value='FINANCIAL-YEAR POSITION (months marked YES only)').font = BOLD; r += 1
    FYC = ['Item', 'IGST', 'CGST', 'SGST', 'CESS', 'TOTAL']
    rng = lambda letter: f'${letter}${m0}:${letter}${m1}'
    crit = lambda fy: f'{rng("B")},"{fy}",{rng("C")},"YES"'
    b.bo_fy_cells = {}
    grand_h, grand_n = [], []
    for fy in bo['fys']:
        a = bo['fy'][fy]
        gen = ', '.join(g.strftime('%d-%m-%Y') for g in a['generated']) or 'not stated'
        c = ws.cell(row=r, column=1, value=f'FY {fy} — BO report generated {gen} — {len(a["months"])} month(s) used')
        c.font = BOLD; r += 1
        b.head(ws, r, FYC); ws.freeze_panes = None; r += 1
        cells = {}
        def line(key, label, per_head=None, total=None, bold=False):
            nonlocal r
            ws.cell(row=r, column=1, value=label)
            if per_head:
                for j, h in enumerate(H4, 2):
                    v = per_head(h, j)
                    ws.cell(row=r, column=j, value=v if v is not None else 0)
                ws.cell(row=r, column=6, value=f'=SUM(B{r}:E{r})')
            elif total:
                ws.cell(row=r, column=6, value=total)
            for j in range(1, 7):
                x = ws.cell(row=r, column=j); x.border = BD
                if j > 1: x.number_format = NUM
                if bold: x.font = BOLD
            ws.cell(row=r, column=1).alignment = Alignment(wrap_text=True, vertical='top')
            cells[key] = r; r += 1
        S = lambda letter: f'=SUMIFS({rng(letter)},{crit(fy)})'
        line('a', 'a. Tax as per GSTR-1 (outward + export-SEZ)', total=S(L[C[0][0]]))
        line('b', 'b. Tax paid through GSTR-3B (outward + export-SEZ)', total=S(L[C[1][0]]))
        line('c', 'c. Outward shortfall, net for the year (a − b)',
             per_head=lambda h, j: S(L[f'Outward shortfall {h} (GSTR-1 − GSTR-3B)']))
        line('d', 'd. Reverse charge in GSTR-2B', total=S(L['Reverse charge in GSTR-2B, all heads']))
        line('e', 'e. Reverse charge declared in GSTR-3B 3.1(d)', total=S(L['Reverse charge declared in GSTR-3B 3.1(d), all heads']))
        line('f', 'f. Reverse-charge shortfall, net for the year (d − e)',
             per_head=lambda h, j: S(L[f'RCM shortfall {h} (GSTR-2B − GSTR-3B)']))
        line('g', 'g. Paid through DRC-03', per_head=lambda h, j: S(L[f'DRC-03 paid {h}']))
        cl_ = lambda j: get_column_letter(j)
        line('h', 'h. TAX STILL DUE = c + f − g, head by head, not below nil',
             per_head=lambda h, j: f'=MAX(0,{cl_(j)}{cells["c"]}+{cl_(j)}{cells["f"]}-{cl_(j)}{cells["g"]})', bold=True)
        line('i', 'i. ITC availed, GSTR-3B net of reversal (other than import and reverse charge)',
             total=S(L['ITC availed, GSTR-3B net of reversal (other than import & RCM), all heads']))
        line('j', 'j. ITC in GSTR-2B (other than import and reverse charge)',
             total=S(L['ITC in GSTR-2B (other than import & RCM), all heads']))
        line('k', 'k. ITC excess over GSTR-2B, net for the year (i − j), not below nil',
             per_head=lambda h, j: f'=MAX(0,{S(L[f"ITC excess {h} (GSTR-3B net − GSTR-2B)"])[1:]})')
        line('l', 'l. Import IGST credit over GSTR-2B (IMPG), not below nil',
             per_head=lambda h, j: (f'=MAX(0,{S(L[f"Import ITC excess {h} (GSTR-3B 4A(1) − GSTR-2B IMPG)"])[1:]})' if h in H2 else 0))
        line('m', 'm. ITC on reverse charge above the reverse-charge tax declared, not below nil',
             per_head=lambda h, j: f'=MAX(0,{S(L[f"ITC on reverse charge over RCM tax declared {h}"])[1:]})')
        line('n', 'n. EXCESS ITC TO BE REVERSED = k + l + m',
             per_head=lambda h, j: f'={cl_(j)}{cells["k"]}+{cl_(j)}{cells["l"]}+{cl_(j)}{cells["m"]}', bold=True)
        line('o', 'o. For reference — sum of every month\'s shortfall (gross; interest under section 50 runs on these)',
             total=S(L["Month's liability shortfall (positive heads only)"]))
        line('p', 'p. For reference — sum of every month\'s ITC excess (gross)',
             total=S(L["Month's ITC excess (positive heads only)"]))
        b.bo_fy_cells[fy] = {k: f"'{SHEETS[1]}'!F{v}" for k, v in cells.items()}
        grand_h.append(f'F{cells["h"]}'); grand_n.append(f'F{cells["n"]}')
        r += 1
    ws.cell(row=r, column=1, value='ALL YEARS — TAX STILL DUE (sum of line h)').font = BOLD
    ws.cell(row=r, column=6, value='=' + '+'.join(grand_h)).font = BOLD
    ws.cell(row=r + 1, column=1, value='ALL YEARS — EXCESS ITC TO BE REVERSED (sum of line n)').font = BOLD
    ws.cell(row=r + 1, column=6, value='=' + '+'.join(grand_n)).font = BOLD
    ws.cell(row=r + 2, column=1, value='ALL YEARS — TOTAL ON THE BO FIGURES').font = BOLD
    ws.cell(row=r + 2, column=6, value=f'=F{r}+F{r+1}').font = BOLD
    for rr in (r, r + 1, r + 2):
        ws.cell(row=rr, column=6).number_format = NUM
        for j in range(1, 7): ws.cell(row=rr, column=j).border = BD
    b.bo_total_cells = (f"'{SHEETS[1]}'!F{r}", f"'{SHEETS[1]}'!F{r+1}")
    n = r + 4
    n = b.note(ws, n, 'Net for the year, head by head, is the defensible demand: a month short-paid and made good later is timing. '
                      'Lines o and p are the gross figures — the outer limit, and the base for interest under section 50(1) on tax '
                      'paid late and section 50(3) on credit wrongly availed and utilised.', bold=True)
    n = b.note(ws, n, 'A month whose GSTR-1 is NEGATIVE (credit notes) is set off against the other months\' shortfall in the '
                      'yearly net, as the BO report does. That is right only if the credit notes are valid under section 34; '
                      'where they are doubtful, line o is the exposure.')
    n = b.note(ws, n, 'Line h is the position Rule 88C / FORM DRC-01B is built on (GSTR-1 and GSTR-2B reverse charge against GSTR-3B). '
                      'Line n is the position Rule 88D / FORM DRC-01C is built on (GSTR-3B credit against GSTR-2B).')
    n = b.note(ws, n, 'Reverse-charge tax must be paid in cash (section 49(4)), and credit on it is available only once it is paid '
                      '(section 16(2)(c)) — line m catches credit taken on reverse charge beyond the tax declared.')
    if any(bo['fy'][f]['pre_2022'] for f in bo['fys']):
        b.note(ws, n, 'Months before January 2022: Rule 36(4) as it then stood allowed provisional credit of up to 5% (10% / 20% '
                      'earlier) over what suppliers had reported. The excess above is computed without that allowance; reduce it '
                      'by the allowance for those months before any demand.')
    ws.column_dimensions['A'].width = 58

def build_changes(b, F_after, F_before, findings_total_cell):
    from gpe_build import BOLD, BD, NUM, MISSING_FILL
    from openpyxl.styles import Alignment
    bo = b.R.get('bo') or {}
    ws = b.sheet(SHEETS[2])
    if not bo.get('available'):
        return b.unavailable(ws, bo, 'CHANGE REPORT — WHAT THE BO COMPARISON REPORT CHANGES')
    ws['A1'] = 'CHANGE REPORT — what taking the BO comparison report as final changes in the conclusions'; ws['A1'].font = BOLD
    ws['A2'] = ('For each financial year the BO report covers: the figure the notebook reached from GST-Prime alone, the figure on the '
                'BO report, and which one 20_FINDINGS uses. Months the BO report does not cover keep the GST-Prime figure.')
    ws['A2'].alignment = Alignment(wrap_text=True, vertical='top')
    r = 4
    ws.cell(row=r, column=1, value='A. EACH CONCLUSION, BEFORE AND AFTER').font = BOLD; r += 1
    H = ['FY', 'Months taken from BO', 'Conclusion', 'Notebook figure (GST-Prime)', 'Notebook basis', 'BO figure', 'BO basis',
         'Change (BO − notebook)', 'Used in 20_FINDINGS', 'Reading']
    b.head(ws, r, H, [10, 30, 40, 18, 40, 18, 40, 18, 16, 70]); r += 1
    a0 = r
    r1 = b.R.get('r1_vs_r3b') or {}; ex = b.R.get('excess_itc') or {}; rc = b.R.get('rcm') or {}
    s17 = '17_R1_vs_R3B' in b.wb.sheetnames and r1.get('available')
    s13 = '13_EXCESS_ITC' in b.wb.sheetnames and ex.get('available')
    n17 = 3 + len(r1.get('rows', [])) if s17 else 0
    n13 = 4 + len(ex.get('rows', [])) if s13 else 0
    for fy in bo['fys']:
        a = bo['fy'][fy]; cells = b.bo_fy_cells[fy]
        arr = '{' + ','.join(f'"{p}"' for p in a['months']) + '}'
        mtxt = f'{len(a["months"])}: {a["months"][0]} to {a["months"][-1]}'
        rcm_nb = (sum(v * 0.05 for p, v in (rc.get('value_without_tax') or []) if p in a['months'])
                  if rc.get('available') and any(p in a['months'] for p in (b.periods or [])) else None)
        rows = [
            ('Output tax short-paid (GSTR-1 and GSTR-2B reverse charge against GSTR-3B)',
             (f"=SUMPRODUCT(SUMIFS('17_R1_vs_R3B'!$L$4:$L${n17},'17_R1_vs_R3B'!$A$4:$A${n17},{arr}))"
              if s17 and a['prime_short_months'] else ''),
             '17_R1_vs_R3B col L: months in which GSTR-1 tax exceeded GSTR-3B outward tax, added up (gross)',
             f'={cells["h"]}', '19F line h: net for the year head by head, reverse charge included, less DRC-03',
             'BO',
             'The notebook adds the months of shortfall; BO nets the year head by head and sets off DRC-03 — the basis of '
             f'DRC-01B under Rule 88C. BO\'s gross figure is 19F line o (Rs.{inr(a["liab_gross"])}).'),
            ('Excess ITC (credit availed against credit in the auto-drafted statement)',
             (f"=SUMPRODUCT(SUMIFS('13_EXCESS_ITC'!$J$5:$J${n13},'13_EXCESS_ITC'!$A$5:$A${n13},{arr}))"
              if s13 and a['prime_excess_months'] else ''),
             '13_EXCESS_ITC col J: against GSTR-2A, net, CGST taken equal to SGST',
             f'={cells["n"]}', '19F line n: against GSTR-2B, head by head, import and reverse-charge credit included',
             'BO',
             'GSTR-2B, not 2A, governs credit from 1 January 2022 (section 16(2)(aa), Rule 36(4)); BO also has CGST as filed '
             f'rather than assumed. BO\'s gross figure is 19F line p (Rs.{inr(a["itc_gross"])}).'),
            ('Reverse charge not declared',
             rcm_nb if rcm_nb is not None else '',
             '16_RCM_CHECK: indicative 5% on reverse-charge value declared with no tax',
             f'=MAX(0,{cells["f"]})', '19F line f: GSTR-2B reverse charge against GSTR-3B 3.1(d), net (already inside line h)',
             'BO (in line h)',
             'BO measures reverse charge against what suppliers reported in GSTR-2B. Reverse charge on supplies from unregistered '
             'persons or on notified services not reported in 2B is NOT caught by either figure.'),
        ]
        for label, nb, nbb, bov, bob, used, rd in rows:
            vals = [fy, mtxt, label, nb if nb != '' else 'not held', nbb, bov, bob, None, used, rd]
            for j, v in enumerate(vals, 1):
                c = ws.cell(row=r, column=j, value=v); c.border = BD; c.alignment = Alignment(wrap_text=True, vertical='top')
            ws.cell(row=r, column=8, value=f'=IF(ISNUMBER(D{r}),F{r}-D{r},F{r})')
            for j in (4, 6, 8): ws.cell(row=r, column=j).number_format = NUM
            r += 1
    r += 1
    ws.cell(row=r, column=1, value='B. REVENUE IMPLICATION IN 20_FINDINGS').font = BOLD; r += 1
    before = sum(f['a'] for f in F_before); after = sum(f['a'] for f in F_after)
    for label, v in (('On the GST-Prime figures alone (as the notebook would conclude without the BO report)', round(before)),
                     ('With the BO figures taken as final (the total now in 20_FINDINGS)', f'={findings_total_cell}')):
        ws.cell(row=r, column=3, value=label).alignment = Alignment(wrap_text=True, vertical='top')
        c = ws.cell(row=r, column=6, value=v); c.number_format = NUM; c.font = BOLD
        r += 1
    ws.cell(row=r, column=3, value='Change').font = BOLD
    c = ws.cell(row=r, column=6, value=f'=F{r-1}-F{r-2}'); c.number_format = NUM; c.font = BOLD
    r += 2
    ws.cell(row=r, column=1, value='C. FINDINGS WHOSE AMOUNT CHANGED').font = BOLD; r += 1
    b.head(ws, r, ['', '', 'Finding', 'Amount without the BO report', '', 'Amount with the BO report', '', 'Change'])
    ws.freeze_panes = None; r += 1
    fb = {f['t']: f['a'] for f in F_before}; fa = {f['t']: f['a'] for f in F_after}
    keys = [t for t in fa if abs(fa[t] - fb.get(t, 0)) > 0.5] + [t for t in fb if t not in fa and abs(fb[t]) > 0.5]
    if not keys:
        ws.cell(row=r, column=3, value='No amount changed.'); r += 1
    for t in keys:
        ws.cell(row=r, column=3, value=t).alignment = Alignment(wrap_text=True, vertical='top')
        ws.cell(row=r, column=4, value=round(fb.get(t, 0))); ws.cell(row=r, column=6, value=round(fa.get(t, 0)))
        ws.cell(row=r, column=8, value=f'=F{r}-D{r}')
        for j in (3, 4, 6, 8):
            ws.cell(row=r, column=j).border = BD
            if j > 3: ws.cell(row=r, column=j).number_format = NUM
        r += 1
    r += 1
    ws.cell(row=r, column=1, value='D. WHERE GST-PRIME AND BO DIFFER (from 19E_BO_vs_NOTEBOOK)').font = BOLD; r += 1
    b.head(ws, r, ['', '', 'Measure', 'Months that DIFFER', '', 'Sum of differences (BO − notebook)'])
    ws.freeze_panes = None; r += 1
    d0, d1 = b.bo_cmp_range
    E = f"'{SHEETS[0]}'"
    for mid, label, _, _ in MEASURES:
        ws.cell(row=r, column=3, value=label).alignment = Alignment(wrap_text=True, vertical='top')
        ws.cell(row=r, column=4, value=f'=COUNTIFS({E}!$D${d0}:$D${d1},C{r},{E}!$K${d0}:$K${d1},"DIFFERS*")')
        ws.cell(row=r, column=6, value=f'=SUMIFS({E}!$J${d0}:$J${d1},{E}!$D${d0}:$D${d1},C{r},{E}!$K${d0}:$K${d1},"DIFFERS*")')
        for j in (3, 4, 6):
            ws.cell(row=r, column=j).border = BD
            if j > 3: ws.cell(row=r, column=j).number_format = NUM
        r += 1
    r += 1
    b.note(ws, r, 'The BO report is the department\'s own comparison: Rule 88C and Rule 88D intimations are generated on it, it reads '
                  'GSTR-1 as amended and GSTR-2B as fixed for the month, and it carries the time it was generated. Where it covers a '
                  'month its figure is final in this workbook. The GST-Prime figure remains on its own sheet so a senior can see both.',
           bold=True)

# ------------------------------------------------------------------ findings, actions, report text
def _heads_txt(dct):
    return ', '.join(f'{h} {_rs(v)}' for h, v in dct.items() if v > 0.5) or 'nil'

def bo_findings(R):
    bo = R.get('bo') or {}
    if not bo.get('available'): return []
    F = []
    fys = bo['fys']
    gen = lambda a: ', '.join(g.strftime('%d-%m-%Y') for g in a['generated']) or 'date not stated'
    due = [(f, bo['fy'][f]) for f in fys if bo['fy'][f]['liab_due_total'] > 1]
    if due:
        F.append(dict(
            t='Output tax short-paid against GSTR-1 and GSTR-2B (reverse charge) — as per the BO comparison report',
            w=('; '.join(f'FY {f} (BO report generated {gen(a)}, {len(a["months"])} month(s)): GSTR-1 tax '
                         f'{_rs(sum(a[("outr1", h)] for h in H4))} against {_rs(sum(a[("out3b", h)] for h in H4))} paid in GSTR-3B'
                         + (f'; reverse charge in GSTR-2B {_rs(sum(a[("rcm2b", h)] for h in H4))} against '
                            f'{_rs(sum(a[("rcm3b", h)] for h in H4))} declared' if abs(a['rcm_net']) > 1 or
                            abs(sum(a[("rcm2b", h)] for h in H4)) > 1 else '')
                         + (f'; DRC-03 {_rs(a["drc03_total"])}' if a['drc03_total'] > 0.5 else '')
                         + f' — STILL DUE {_rs(a["liab_due_total"])} ({_heads_txt(a["liab_due"])})' for f, a in due)
               + '. Net for each year, head by head; the months of shortfall are on 19F.'),
            a=bo['liab_due'], s='19F_BO_FINAL_POSITION (line h); SRC_BO_CMP',
            p=('Section 75(12) with Explanation — tax declared in GSTR-1 and not paid through GSTR-3B is self-assessed tax, '
               'recoverable under section 79 without a show cause notice; Rule 88C / FORM DRC-01B is generated on exactly this '
               'comparison. Reverse charge: sections 9(3), 9(4), 49(4). Section 50 interest.'),
            r='HIGH — directly recoverable'))
    exc = [(f, bo['fy'][f]) for f in fys if bo['fy'][f]['itc_due_total'] > 1]
    if exc:
        pre = [p for f, a in exc for p in a['pre_2022']]
        F.append(dict(
            t='Excess ITC availed over GSTR-2B — as per the BO comparison report',
            w=('; '.join(f'FY {f} (BO report generated {gen(a)}): {_rs(a["itc_due_total"])} ({_heads_txt(a["itc_due"])})'
                         + (f', of which import IGST {_rs(a["impg_due"])}' if a['impg_due'] > 0.5 else '')
                         + (f', of which credit on reverse charge above the tax declared {_rs(a["rcmitc_due"])}'
                            if a['rcmitc_due'] > 0.5 else '') for f, a in exc)
               + '. Net for each year, head by head, after the Table 4B reversal.'
               + (f' {len(pre)} month(s) fall before January 2022, when Rule 36(4) allowed a provisional margin: reduce for it.'
                  if pre else '')),
            a=bo['itc_due'], s='19F_BO_FINAL_POSITION (line n); SRC_BO_CMP',
            p=('Section 16(2)(aa) and Rule 36(4); Rule 88D / FORM DRC-01C; section 16(2)(c) with section 49(4) for credit on '
               'reverse charge; sections 73 / 74 / 74A; section 50(3); Rule 86A.'),
            r='VERY HIGH'))
    for f in fys:
        a = bo['fy'][f]
        if a['neg_r1'] and a['liab_gross'] > a['liab_due_total'] + 1:
            neg = sum(-v for _, v, _ in a['neg_r1'])
            at_risk = min(a['liab_gross'] - a['liab_due_total'], neg)
            F.append(dict(
                t=f'FY {f}: shortfall of {_rs(a["liab_gross"])} cancelled in the year-end net by a NEGATIVE GSTR-1 (credit notes)',
                w=(f'The months of shortfall ({", ".join(f"{p} {_rs(v)}" for p, k, v in a["short_months"])}) add up to '
                   f'{_rs(a["liab_gross"])}, but the year nets to {_rs(a["liab_due_total"])} only because GSTR-1 for '
                   + ', '.join(f'{p} is MINUS {_rs(-v)} (GSTR-3B for the month {_rs(t3)})' for p, v, t3 in a['neg_r1'])
                   + '. The BO report sets the negative month off against the others, as it would a genuine credit note. '
                     f'That holds only if the credit notes are valid; if they are not, up to {_rs(at_risk)} more is due. '
                     'Not counted in the total.'),
                a=0.0, s='19F_BO_FINAL_POSITION (month table, line o); 19B_CREDIT_NOTES',
                p=('Section 34(1) and (2) — a credit note must relate to an identified supply and be declared by 30 November '
                   'following the year; the proviso bars the reduction where the incidence of tax has been passed on and the '
                   'recipient has not reversed the credit. Section 75(12); sections 74 / 74A where suppression is shown.'),
                r='VERY HIGH — lead'))
    tim = []
    for f in fys:
        a = bo['fy'][f]
        if a['liab_gross'] > a['liab_due_total'] + 1 and a['short_months'] and not a['neg_r1']:
            tim.append(f'FY {f}: tax short in ' + ', '.join(f'{p} ({k}) {_rs(v)}' for p, k, v in a['short_months'])
                       + f', against {_rs(a["liab_due_total"])} still due for the year')
        if a['itc_gross'] > a['itc_due_total'] + 1 and a['excess_months']:
            tim.append(f'FY {f}: credit in excess of GSTR-2B in ' + ', '.join(f'{p} {_rs(v)}' for p, v in a['excess_months'])
                       + f', against {_rs(a["itc_due_total"])} for the year')
    if tim:
        F.append(dict(
            t='Monthly shortfalls offset by other months of the same year — interest',
            w='; '.join(tim) + '. Within each year other months offset these, in part or in full; only the yearly net is '
                               'demanded (findings above). Interest runs on a month\'s shortfall until it was made good.',
            a=0.0, s='19F_BO_FINAL_POSITION (lines o and p; month table)',
            p='Section 50(1) on tax paid late; section 50(3) on credit wrongly availed and utilised. Rule 88B.',
            r='MEDIUM — interest'))
    rc = []
    for x in bo['months']:
        if not x['covered']: continue
        f = x['f']
        r2 = sum(f[('rcm2b', h)] for h in H4); r3 = sum(f[('rcm3b', h)] for h in H4); ri = sum(f[('rcmitc', h)] for h in H4)
        if (abs(r2) > 1 or abs(ri) > 1) and abs(r3) <= 1:
            rc.append((x['period'], r2, ri))
    if rc:
        F.append(dict(
            t='Reverse-charge supplies in GSTR-2B, or credit taken on reverse charge, with no reverse-charge tax declared',
            w=('; '.join(f'{p}: reverse charge in GSTR-2B {_rs(r2) if r2 >= 0 else "MINUS " + _rs(-r2)}, credit claimed on reverse '
                         f'charge (Table 4A(2)/(3)) {_rs(ri) if ri >= 0 else "MINUS " + _rs(-ri)}, reverse-charge tax in Table 3.1(d) nil'
                         for p, r2, ri in rc)
               + '. Credit on reverse charge is available only after the tax is paid, and that tax is payable only in cash. '
                 'Whatever the yearly net, these entries show that the taxpayer receives supplies on reverse charge; the supplier '
                 'and the nature of the supply are in the supplier-wise GSTR-2B for those months.'),
            a=0.0, s='19F_BO_FINAL_POSITION; SRC_BO_CMP',
            p='Sections 9(3), 9(4) with Notifications 13/2017-CT(R) and 4/2017-CT(R); section 49(4); section 16(2)(c); section 31(3)(f).',
            r='HIGH — lead'))
    d = bo['differs']
    if d:
        by = {}
        for c in d: by.setdefault(c['measure'], []).append(c)
        big = sorted(d, key=lambda c: -abs(c['diff'] or 0))[:6]
        F.append(dict(
            t=f'GST-Prime figures differ from the BO comparison report in {len(d)} place(s) — BO figures taken',
            w=('; '.join(f'{m}: {len(v)} month(s)' for m, v in by.items()) + '. Largest: '
               + '; '.join(f'{c["period"]} {c["measure"].split(" (")[0]} Prime {_rs(c["prime"])}, BO {_rs(c["bo"])}' for c in big)
               + '. Every amount in this workbook for those months is taken from the BO report; the reason for each difference, '
                 'where the figures show it, is on 19E.'),
            a=0.0, s='19E_BO_vs_NOTEBOOK; 19G_CHANGE_REPORT', p='—', r='CHECK — BO figures taken'))
    if bo['not_covered']:
        F.append(dict(
            t=f'BO report older than the return for {len(bo["not_covered"])} month(s) — not used for them',
            w='; '.join(f'{x["period"]}: {x["why"]}' for x in bo['not_covered'][:12]),
            a=0.0, s='19E_BO_vs_NOTEBOOK; 19F_BO_FINAL_POSITION', p='—', r='PENDING DATA'))
    if bo['prime_without_bo']:
        fy_miss = sorted({fy_of(p) for p in bo['prime_without_bo']}, key=lambda f: int(f.split('-')[0]))
        F.append(dict(
            t='No BO comparison report for FY ' + ', '.join(fy_miss),
            w=(f'{len(bo["prime_without_bo"])} month(s) in the GST-Prime data ({bo["prime_without_bo"][0]} to '
               f'{bo["prime_without_bo"][-1]}) rest on GST-Prime alone. Download the BO comparison report for these years and '
               f'run again; the conclusions for them will then be taken from it.'),
            a=0.0, s='23_DATA_GAPS', p='—', r='PENDING DATA'))
    return F

def bo_actions(R):
    bo = R.get('bo') or {}
    if not bo.get('available'): return []
    A = []
    if bo['liab_due'] > 1:
        A.append(['HIGH', f'Issue the Rule 88C intimation in FORM DRC-01B for the tax of {_rs(bo["liab_due"])} still due on the BO '
                          f'comparison (19F line h), year by year and head by head.',
                  'Self-assessed tax under the Explanation to section 75(12), recoverable under section 79 without a show cause '
                  'notice. The BO figures are the ones the intimation itself is built on.',
                  'Rule 88C, FORM DRC-01B, section 75(12), section 79, section 50'])
    if bo['itc_due'] > 1:
        A.append(['HIGH', f'Issue the Rule 88D intimation in FORM DRC-01C for the credit of {_rs(bo["itc_due"])} availed over '
                          f'GSTR-2B (19F line n), and draw the supplier-wise GSTR-2B for the months carrying it.',
                  'Credit over GSTR-2B is barred by section 16(2)(aa) from 1 January 2022. The supplier-wise 2B shows which '
                  'supplier\'s invoices are missing.', 'Rule 88D, FORM DRC-01C, section 16(2)(aa), Rule 36(4), Rule 86A'])
    rcm = [x['period'] for x in bo['months'] if x['covered'] and
           (abs(sum(x['f'][('rcm2b', h)] for h in H4)) > 1 or abs(sum(x['f'][('rcmitc', h)] for h in H4)) > 1)
           and abs(sum(x['f'][('rcm3b', h)] for h in H4)) <= 1]
    if rcm:
        A.append(['HIGH', f'Draw the supplier-wise GSTR-2B (reverse-charge section) for {", ".join(rcm)}; identify the supplier, the '
                          f'service and the invoice, and quantify the reverse-charge tax for the whole period.',
                  'Reverse-charge supplies are being received but not declared in Table 3.1(d); the same supplier is likely to '
                  'recur in other months.', 'Sections 9(3), 9(4), 49(4), 16(2)(c), section 70'])
    if bo['not_covered']:
        fys = sorted({x['fy'] for x in bo['not_covered']})
        A.append(['HIGH', 'Download a fresh BO comparison report for FY ' + ', '.join(fys) + ' and run again.',
                  'The report held was generated before some returns were filed, so those months still rest on GST-Prime.', '—'])
    if bo['prime_without_bo']:
        fys = sorted({fy_of(p) for p in bo['prime_without_bo']}, key=lambda f: int(f.split('-')[0]))
        A.append(['MEDIUM', 'Download the BO comparison report (Tax liabilities and ITC comparison) for FY ' + ', '.join(fys)
                            + ' and add it to the case.',
                  'For those years the conclusions rest on GST-Prime alone.', '—'])
    if bo['differs']:
        A.append(['MEDIUM', f'Look at the {len(bo["differs"])} difference(s) between GST-Prime and BO on 19E before relying on any '
                            f'GST-Prime figure for the same months in correspondence.',
                  'The BO figure has been used for every amount; the Prime figure should not be quoted where they differ.', '—'])
    return A

def bo_report_lines(R):
    bo = R.get('bo') or {}
    if not bo.get('available'): return []
    st = {}
    for c in bo['cmp']:
        st[c['status']] = st.get(c['status'], 0) + 1
    files = '; '.join(f'FY {x["fy"]} generated {x["generated"]:%d-%m-%Y}' if x['generated'] else f'FY {x["fy"]}'
                      for x in sorted(bo['files'], key=lambda x: x['fy'] or '') if x['status'] == 'COUNTED')
    used = sum(1 for x in bo['months'] if x['covered'])
    t = (f'2.7  The BO portal\'s comparison report (Tax liabilities and ITC comparison) is held for {files}, covering '
         f'{len(bo["months"])} month(s), of which {used} are taken as final. Every figure the notebook built from GST-Prime for '
         f'those months has been compared with it (sheet 19E): '
         + ', '.join(f'{v} {k}' for k, v in sorted(st.items(), key=lambda kv: -kv[1])) + '. '
         f'On the BO figures the tax still due is {_rs(bo["liab_due"])} and the credit availed over GSTR-2B is '
         f'{_rs(bo["itc_due"])} (sheet 19F). The amounts in the findings below are taken from the BO report wherever it covers '
         f'a month; sheet 19G sets out what that changed.')
    if bo['prime_without_bo']:
        t += f' {len(bo["prime_without_bo"])} month(s) of GST-Prime data have no BO report and rest on GST-Prime alone.'
    return [t, '']

DRIVES_TEXT = ('Final figures: GSTR-1 and GSTR-2B against GSTR-3B, head by head, month by month, as the BO portal computes them. '
               'Taken as final wherever it covers a month; compared with every GST-Prime figure (19E, 19F, 19G).')
SHEET_NOTES = {
    'SRC_BO_CMP': 'The BO comparison report(s), one row per month, every block of every sheet of the report.',
    '19E_BO_vs_NOTEBOOK': 'Every GST-Prime figure compared with the BO report, month by month, with the reason for each difference.',
    '19F_BO_FINAL_POSITION': 'The final position on the BO figures: tax still due and excess credit, year by year, head by head.',
    '19G_CHANGE_REPORT': 'What taking the BO report as final changed in the conclusions and in the revenue figure.',
}
DERIV_ROWS = [
    ['SRC_BO_CMP', 'Every column', 'Read block by block from each sheet of the BO report; column name = section | side | head',
     'BO comparison report', 'The block headers and head labels are checked before any value is read, the months are checked '
     'against the report\'s own Total row, and the report\'s own difference columns are re-computed. A month with no DRC-03 is '
     'blank in the report and totalled there as nil, so it is held as 0.'],
    ['SRC_BO_CMP', 'ITC other | GSTR-3B gross (blank)', 'Left blank for a BO report in the older layout (generated up to about '
     '2022), whose ITC sheet has one GSTR-3B block already net of the Table 4B reversal', 'BO comparison report',
     'That block is read as the after-reversal figure, which is the one every computation uses; a before-reversal figure '
     'was not reported and is not invented.'],
    ['SRC_BO_CMP', 'Which BO report is used for a month', 'The one GENERATED LAST (time printed on the report), whatever the upload order',
     'Report header', 'A later download reflects GSTR-1 amendments and late returns filed in between.'],
    ['19E_BO_vs_NOTEBOOK', 'Status', 'MATCH within Rs.5; MATCH (before reversal) where Prime agrees with BO before Table 4B; '
     'DIFFERENT STATEMENT for 2A against 2B; BO NOT USED where the report pre-dates the return', 'This sheet',
     'Prime reports whole rupees per head, BO paise. 2A and 2B are different statements and are not expected to agree.'],
    ['19E_BO_vs_NOTEBOOK', 'Reading', 'Where the difference equals one identifiable component (export tax, reverse charge, the 4B '
     'reversal, import IGST, CGST assumed) the sheet names it', 'SRC_BO_CMP', 'Shows the cause, not only the size.'],
    ['19F_BO_FINAL_POSITION', 'Used for the final figures?', 'NO where the BO report was generated before the month\'s GSTR-3B was '
     'filed (date from R3B PAYMENT SUMMARY) or fell due, or shows no GSTR-3B while Prime does', 'SRC_BO_CMP; SRC_R3B_FILING',
     'A report generated before the return existed would show it as nil. The GST-Prime figure is kept for that month.'],
    ['19F_BO_FINAL_POSITION', 'h. Tax still due', 'MAX(0, outward shortfall + reverse-charge shortfall − DRC-03), per head, '
     'net for the year', 'Month table on the same sheet', 'Heads are not interchangeable; within a year a later payment makes '
     'good an earlier shortfall (interest apart). Same basis as the BO report and Rule 88C.'],
    ['19F_BO_FINAL_POSITION', 'n. Excess ITC', 'MAX(0, GSTR-3B net of reversal − GSTR-2B) + import IGST over IMPG + credit on '
     'reverse charge over the tax declared, per head, net for the year', 'Month table on the same sheet',
     'Section 16(2)(aa) and Rule 36(4); section 16(2)(c) and 49(4) for reverse charge. Same basis as Rule 88D.'],
    ['19G_CHANGE_REPORT', 'Notebook figure', 'SUMPRODUCT(SUMIFS) over 17_R1_vs_R3B / 13_EXCESS_ITC for exactly the months the BO '
     'report covers', '17_R1_vs_R3B; 13_EXCESS_ITC', 'Like for like: the same months on both sides.'],
]
