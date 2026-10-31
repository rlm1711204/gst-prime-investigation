# -*- coding: utf-8 -*-
"""
GST-PRIME PRELIMINARY INVESTIGATION ENGINE  —  e-way bills (EWB document-wise summary)

Reads GST-Prime's "Taxpayer Profile > E-Way Bill > DOCUMENT-WISE SUMMARY" page saved to Excel, in whatever shape the
save produced:
  * a normal table (SNo | Period | Document Type | Category | No of EWBs | Assessable Value | SGST | CGST | IGST | CESS | GST),
    starting in any column, possibly twice on one sheet;
  * the whole page pasted into ONE cell as text (tab-separated rows, as Ctrl+A / Ctrl+C from the browser gives).
One workbook may hold several pages (one per sheet), each a 12-13 month window, and the windows overlap
(Aug-2024..Aug-2025 and Aug-2025..Sep-2026 both carry Aug-2025).

Every row read is kept, with the sheet, block and position it came from. One record is then taken per
(Return Period, Document Type, Category): copies that agree are counted once; copies that differ are listed
as a CONFLICT and the copy from the most recent window is used. Nothing is added up twice.

The document-wise summary does not say whether the GSTIN is the consignor or the consignee of these e-way
bills. Every reading below says "if these are outward" where it matters.
"""
import os, re, datetime
from collections import OrderedDict

EWB_VERSION = "1.0"
EWB_CODE = 'EWB_DOC'
SHEETS = ['19H_EWB_COPIES', '19I_EWB_RECORDS', '19J_EWB_MONTHLY', '19K_EWB_FY', '19L_EWB_FINDINGS', '19M_EWB_HSN_RATE']

HEADER = ['SNo', 'Period', 'Document Type', 'Category', 'No of EWBs', 'Assessable Value',
          'SGST', 'CGST', 'IGST', 'CESS', 'GST']
FIG = ['No of EWBs', 'Assessable Value', 'SGST', 'CGST', 'IGST', 'CESS', 'GST']
GSTIN_RE = re.compile(r'\b(\d{2}[A-Z]{5}\d{4}[A-Z][1-9A-Z]Z[0-9A-Z])\b')
MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
RATE_CUT = datetime.date(2025, 9, 22)          # Notification 9/2025-Integrated Tax (Rate): rates from 22-09-2025
SLABS = (0, 0.25, 1.5, 3, 5, 12, 18, 28, 40)    # rates on GOODS (6% and 7.5% are services / special cases)
LOW_RATE = 5.0           # % : tax below this on ...
LOW_RATE_VALUE = 1_000_000   # ... a value of Rs.10 lakh or more is a lead
SPIKE = 3.0              # a month above 3 x the median month


# ------------------------------------------------------------------ small helpers (no gpe_core import at module load)
def _num(v):
    from gpe_core import to_num
    return to_num(v)

def _period(v):
    if isinstance(v, (datetime.datetime, datetime.date)):
        return f'{MON[v.month - 1]}-{v.year}'
    from gpe_core import parse_period
    t = parse_period(v)
    return t[0] if t else None

def _pkey(p):
    from gpe_core import period_key
    return period_key(p)

def _fy(p):
    from gpe_core import fy_of
    return fy_of(p)

def _pdate(p):
    m, y = p.split('-')
    return datetime.date(int(y), MON.index(m) + 1, 1)

def _month_range(a, b):
    out, d = [], _pdate(a)
    end = _pdate(b)
    while d <= end:
        out.append(f'{MON[d.month - 1]}-{d.year}')
        d = datetime.date(d.year + (d.month == 12), d.month % 12 + 1, 1)
    return out

def _inr(x):
    from gpe_core import inr
    return inr(x)

def _lakh(x):
    return f'Rs.{x / 1e5:,.2f} lakh'

def nearest_slab(rate):
    return min(SLABS, key=lambda s: abs(s - rate))


# ------------------------------------------------------------------ recognise
def _strings(ws, limit=400):
    n = 0
    for row in ws.iter_rows(values_only=True):
        for v in row:
            if isinstance(v, str):
                yield v
        n += 1
        if n >= limit: break

def is_ewb_file(path):
    """A GST-Prime EWB document-wise summary: 'DOCUMENT-WISE SUMMARY' on the page and a 'No of EWBs' column."""
    if not str(path).lower().endswith(('.xlsx', '.xlsm')):
        return False
    try:
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        for ws in wb.worksheets:
            txt = ' '.join(_strings(ws, 200))
            if 'DOCUMENT-WISE SUMMARY' in txt.upper() and 'No of EWBs' in txt:
                return True
        return False
    except Exception:
        return False


# ------------------------------------------------------------------ read
def _row_from(vals, where):
    """vals: the 11 cells of one data row in HEADER order. Returns a dict or None (not a data row)."""
    if not vals or vals[0] is None: return None
    s0 = str(vals[0]).strip()
    if not re.fullmatch(r'\d+(\.0)?', s0): return None
    p = _period(vals[1])
    if p is None: return None
    d = OrderedDict()
    d['SNo'] = int(float(s0)); d['Return Period'] = p
    d['Document Type'] = str(vals[2] or '').strip()
    d['Category'] = str(vals[3] or '').strip().upper()
    for k, v in zip(FIG, vals[4:11]):
        d[k] = _num(v)
    d['_where'] = where
    return d

def _meta_from_text(txt):
    m = {}
    g = GSTIN_RE.findall(txt)
    if g: m['gstin'] = g[0]
    x = re.search(r'Legal Name[\t ]*\n?[\t ]*(.+?)[\t ]+Jurisdiction', txt)
    if x: m['legal'] = x.group(1).strip()
    x = re.search(r'FROM\s+PrevNext\s+([A-Z][a-z]{2})\s+(\d{4})\s+TO\s+PrevNext\s+([A-Z][a-z]{2})\s+(\d{4})', txt)
    if x: m['window'] = (f'{x.group(1)}-{x.group(2)}', f'{x.group(3)}-{x.group(4)}')
    x = re.search(r'Total Records:\s*(\d+)', txt)
    if x: m['declared'] = int(x.group(1))
    x = re.search(r'Office[\t ]+(.+?)[\t ]+Tax Group', txt)
    if x: m['office'] = x.group(1).strip()
    return m

def _blocks_from_text(txt, sheet):
    """The whole page in one cell. Rows are tab-separated lines after the 'SNo<TAB>Period' header line."""
    lines = txt.replace('\r', '').split('\n')
    blocks, cur = [], None
    for i, ln in enumerate(lines):
        cells = [c.strip() for c in ln.split('\t')]
        if len(cells) >= 11 and cells[0] == 'SNo' and cells[1] == 'Period':
            cur = dict(sheet=sheet, layout='text in one cell', start=f'line {i + 1}', rows=[])
            blocks.append(cur); continue
        if cur is None: continue
        if ln.startswith('Total Records') or ln.startswith('©'):
            cur = None; continue
        r = _row_from(cells[:11], f'{sheet} · text line {i + 1}')
        if r: cur['rows'].append(r)
    return blocks

def _blocks_from_table(rows, sheet):
    from openpyxl.utils import get_column_letter
    blocks = []
    for i, row in enumerate(rows):
        for j in range(len(row) - 1):
            if row[j] == 'SNo' and row[j + 1] == 'Period':
                b = dict(sheet=sheet, layout='table', start=f'{get_column_letter(j + 1)}{i + 1}', rows=[])
                for k in range(i + 1, len(rows)):
                    rr = rows[k]
                    c0 = rr[j] if j < len(rr) else None
                    if isinstance(c0, str) and (c0.startswith('Total Records') or c0.startswith('©')): break
                    vals = list(rr[j:j + 11]) + [None] * 11
                    r = _row_from(vals[:11], f'{sheet}!{get_column_letter(j + 1)}{k + 1}')
                    if r: b['rows'].append(r)
                    elif c0 is not None and not isinstance(c0, (int, float)): break
                blocks.append(b)
    return blocks

def read_ewb(path):
    """Returns (pages, problems). One page per sheet: its meta and its block(s) of rows."""
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)
    pages, problems = [], []
    for ws in wb.worksheets:
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        txt_cells = [v for r in rows for v in r if isinstance(v, str)]
        joined = '\n'.join(txt_cells)
        if 'DOCUMENT-WISE SUMMARY' not in joined.upper():
            continue
        meta = _meta_from_text(joined)
        if 'legal' not in meta:
            for r in rows:
                for j, v in enumerate(r[:-1]):
                    if v == 'Legal Name' and r[j + 1]:
                        meta['legal'] = str(r[j + 1]).strip(); break
                if 'legal' in meta: break
        blocks = []
        for v in txt_cells:
            if '\n' in v and 'SNo\tPeriod' in v:
                blocks += _blocks_from_text(v, ws.title)
        blocks += _blocks_from_table(rows, ws.title)
        if not blocks:
            problems.append(f'Sheet "{ws.title}": the page title is there but no table of e-way bill rows was found.')
            continue
        for k, b in enumerate(blocks, 1):
            b['block'] = k
            ps = [r['Return Period'] for r in b['rows']]
            b['first'] = min(ps, key=_pkey) if ps else None
            b['last'] = max(ps, key=_pkey) if ps else None
        pages.append(dict(sheet=ws.title, meta=meta, blocks=blocks))
    return pages, problems


def add_ewb(ing, path, rec, batch):
    """Called by Ingest.add. Every row of every block is kept as a COPY; one record per
    (period, document type, category) is chosen later, in chk_ewb."""
    from gpe_core import REPORT_SPECS
    name = rec['file']
    title = REPORT_SPECS[EWB_CODE]['title']
    try:
        pages, problems = read_ewb(path)
    except Exception as e:
        rec.update(code=EWB_CODE, title=title, status='UNREADABLE', note=f'EWB summary, but it could not be opened: {e}')
        ing.register.append(rec); return rec
    if not pages:
        rec.update(code=EWB_CODE, title=title, status='NOT RECOGNISED',
                   note='EWB document-wise summary, but no rows could be read. ' + ' '.join(problems))
        ing.register.append(rec); return rec
    gstins = {p['meta'].get('gstin') for p in pages} - {None}
    file_gstin = sorted(gstins)[0] if len(gstins) == 1 else None
    named = {p['meta'].get('legal') for p in pages if p['meta'].get('gstin')} - {None}
    legal_ref = sorted(named)[0] if named else None
    notes, info, used_pages, refused = [], [], [], []
    if len(gstins) > 1:
        rec.update(code=EWB_CODE, title=title, status='NOT USED - OTHER GSTIN',
                   note=f'One file carries pages for more than one GSTIN ({", ".join(sorted(gstins))}). Save each '
                        f'taxpayer in its own file. Nothing has been taken from it.')
        ing.register.append(rec); return rec
    if ing.case_gstin and file_gstin and file_gstin != ing.case_gstin:
        rec.update(code=EWB_CODE, title=title, status='NOT USED - OTHER GSTIN', subject_gstin=file_gstin,
                   note=f'WRONG TAXPAYER: the e-way bill summary is for GSTIN {file_gstin}, not {ing.case_gstin}. '
                        'Nothing has been taken from it.')
        ing.register.append(rec); return rec
    if legal_ref is None:
        names = {p['meta'].get('legal') for p in pages} - {None}
        legal_ref = sorted(names)[0] if len(names) == 1 else None
    for p in pages:
        lg = p['meta'].get('legal')
        if legal_ref and lg and lg.strip().lower() != legal_ref.strip().lower():
            refused.append(f'sheet "{p["sheet"]}" (legal name "{lg}", not "{legal_ref}")')
            continue
        used_pages.append(p)
    gstin = file_gstin or ing.case_gstin
    if file_gstin is None:
        notes.append(f'The GSTIN is not printed on any page of this file; it is filed under the case GSTIN {gstin} on the '
                     f'strength of the legal name "{legal_ref}". Check that the name is the taxpayer\'s.')
    else:
        no_g = [p['sheet'] for p in used_pages if not p['meta'].get('gstin')]
        if no_g:
            info.append(f'GSTIN {file_gstin} is printed only on sheet(s) '
                         f'{", ".join(p["sheet"] for p in used_pages if p["meta"].get("gstin"))}; sheet(s) {", ".join(no_g)} '
                         f'carry the same legal name ("{legal_ref}") and are taken as the same taxpayer.')
    if refused:
        notes.append('NOT USED — other taxpayer: ' + '; '.join(refused) + '.')
    if problems:
        notes.append(' '.join(problems))
    bucket = ing.data.setdefault(EWB_CODE, [])
    blocks = ing.__dict__.setdefault('ewb_blocks', [])
    n = 0
    for p in used_pages:
        for b in p['blocks']:
            bid = len(blocks) + 1
            blocks.append(dict(id=bid, file=name, batch=batch, sheet=p['sheet'], block=b['block'], layout=b['layout'],
                               start=b['start'], rows=len(b['rows']), first=b['first'], last=b['last'],
                               window=p['meta'].get('window'), declared=p['meta'].get('declared'),
                               gstin=p['meta'].get('gstin'), legal=p['meta'].get('legal'), office=p['meta'].get('office')))
            for r in b['rows']:
                x = dict(r)
                x.update(_period=r['Return Period'], _fy=_fy(r['Return Period']), _file=name, _batch=batch, _block=bid)
                bucket.append(x); n += 1
    ing._seen_md5[rec['md5']] = name
    span = sorted({r['_period'] for r in bucket if r['_file'] == name}, key=_pkey)
    lay = sorted({b['layout'] for b in blocks if b['file'] == name})
    rec.update(code=EWB_CODE, title=title, rows=n, declared=None, subject_gstin=gstin,
               status='COUNTED' if n else 'NOT RECOGNISED',
               detect_note=f'EWB document-wise summary (page title and "No of EWBs" column); '
                           f'{len([b for b in blocks if b["file"] == name])} block(s) on {len(used_pages)} sheet(s), '
                           f'layout: {", ".join(lay)}; {span[0] if span else "-"} to {span[-1] if span else "-"}; '
                           f'{n} row(s) read, overlapping windows included — one record per month, document type and '
                           f'category is chosen in 19I_EWB_RECORDS. ' + ' '.join(info),
               note=' '.join(notes).strip())
    ing.register.append(rec)
    return rec


# ------------------------------------------------------------------ one record per key
FIGS6 = ['No of EWBs', 'Assessable Value', 'SGST', 'CGST', 'IGST', 'CESS']

def _same(a, b, tol=1.0):
    for k in FIG:
        x, y = a.get(k), b.get(k)
        if (x is None) != (y is None): return False
        if x is not None and abs(x - y) > tol: return False
    return True

def dedupe(copies, blocks):
    """copies: every row read. Returns (records, copy_status) where copy_status[i] says what happened to copies[i]."""
    bmap = {b['id']: b for b in blocks}
    def recency(i):
        b = bmap[copies[i]['_block']]
        return (_pkey(b['last']) if b['last'] else (0, 0), b['batch'], b['id'])
    groups = OrderedDict()
    for i, c in enumerate(copies):
        groups.setdefault((c['_period'], c['Document Type'], c['Category']), []).append(i)
    records, status = [], {}
    for key, idx in groups.items():
        keep = max(idx, key=recency)
        agree = [i for i in idx if _same(copies[i], copies[keep])]
        differ = [i for i in idx if i not in agree]
        for i in idx:
            status[i] = ('USED' if i == keep else
                         'REPEAT — SAME FIGURES, NOT COUNTED AGAIN' if i in agree else
                         'REPEAT — FIGURES DIFFER, NOT USED (see conflict)')
        r = dict(copies[keep])
        r.update(_key=key, _copies=len(idx), _agree=len(agree), _differ=len(differ), _keep=keep,
                 _differ_detail=[(bmap[copies[i]['_block']], {k: copies[i].get(k) for k in FIG}) for i in differ])
        records.append(r)
    records.sort(key=lambda r: (_pkey(r['_period']), r['Document Type'], r['Category']))
    return records, status


# ------------------------------------------------------------------ analysis
def _rate(r):
    v = r.get('Assessable Value') or 0
    t = sum((r.get(k) or 0) for k in ('SGST', 'CGST', 'IGST'))
    return (100.0 * t / v) if v else None, t

def chk_ewb(ing, master=None, rate_master=None):
    if not ing.have(EWB_CODE):
        return dict(available=False, missing=[EWB_CODE])
    copies = ing.data[EWB_CODE]
    blocks = getattr(ing, 'ewb_blocks', [])
    recs, cstat = dedupe(copies, blocks)
    flags = []          # (record index or None, code, level, text)
    for i, r in enumerate(recs):
        rate, tax = _rate(r)
        r['_tax'] = tax
        r['_rate'] = rate
        v = r.get('Assessable Value') or 0
        n = r.get('No of EWBs') or 0
        r['_per_ewb'] = v / n if n else None
        r['_slab'] = nearest_slab(rate) if rate is not None else None
        r['_cess_rate'] = (100.0 * (r.get('CESS') or 0) / v) if v else None
        fl = []
        gst_sum = sum((r.get(k) or 0) for k in ('SGST', 'CGST', 'IGST', 'CESS'))
        if r.get('GST') is not None and abs(gst_sum - r['GST']) > 1:
            fl.append(('TIE', 'NOTE', f'GST column {_inr(r["GST"])} is not SGST+CGST+IGST+CESS ({_inr(gst_sum)}).'))
        if r['Category'] == 'INTRA' and (r.get('IGST') or 0) > 0:
            fl.append(('HEAD', 'HIGH', 'IGST charged on an intra-State movement.'))
        if r['Category'] == 'INTER' and ((r.get('SGST') or 0) + (r.get('CGST') or 0)) > 0:
            fl.append(('HEAD', 'HIGH', 'CGST/SGST charged on an inter-State movement.'))
        if r['Category'] == 'INTRA' and abs((r.get('SGST') or 0) - (r.get('CGST') or 0)) > 1:
            fl.append(('HEAD', 'NOTE', 'SGST and CGST are not equal.'))
        if r['Document Type'].lower().startswith('bill of supply') and tax > 0:
            fl.append(('BOS_TAX', 'HIGH',
                       f'Bill of Supply carrying tax of {_inr(tax)} ({rate:.2f}% of value). A bill of supply is issued '
                       f'for exempt supplies or by a composition taxpayer and carries no tax (s.31(3)(c), Rule 49).'))
        if v == 0 and n:
            fl.append(('ZERO', 'NOTE', f'{int(n)} e-way bill(s) with NIL value.'))
        if rate is not None and v >= LOW_RATE_VALUE and rate < LOW_RATE and r['Document Type'].lower().startswith('tax'):
            big = (r['_per_ewb'] or 0) >= 500000
            fl.append(('LOW_RATE', 'HIGH',
                       f'Tax only {rate:.2f}% of a value of {_lakh(v)} ({int(n)} EWB, {_lakh(r["_per_ewb"] or 0)} each). '
                       + ('Large consignments with almost no tax.' if big else
                          'Most of the value carries no tax.')))
        elif rate is not None and v and abs(rate - r['_slab']) > 0.25 and tax > 0:
            fl.append(('MIXED', 'NOTE', f'Effective rate {rate:.2f}% is not a GST slab: goods at different rates, or part '
                                        f'of the value untaxed.'))
        if (r.get('CESS') or 0) > 0:
            post = _pdate(r['_period']) >= datetime.date(2025, 10, 1)
            fl.append(('CESS', 'HIGH' if post else 'MEDIUM',
                       f'Compensation cess {_inr(r["CESS"])} = {r["_cess_rate"]:.2f}% of value.'
                       + (' Cess on motor vehicles and most goods is NIL from 22-09-2025.' if post else '')))
        r['_flags'] = fl
    # months
    span = _month_range(recs[0]['_period'], recs[-1]['_period'])
    months = []
    vals = []
    for p in span:
        rs = [r for r in recs if r['_period'] == p]
        v = sum(r.get('Assessable Value') or 0 for r in rs)
        months.append(dict(period=p, fy=_fy(p), n=sum(r.get('No of EWBs') or 0 for r in rs), value=v,
                           tax=sum(r['_tax'] for r in rs), cess=sum(r.get('CESS') or 0 for r in rs), rows=len(rs)))
        if rs: vals.append(v)
    med = sorted(vals)[len(vals) // 2] if vals else 0
    for m in months:
        m['spike'] = bool(med and m['value'] > SPIKE * med)
        m['gap'] = (m['rows'] == 0)
    # rate regime (tax invoices)
    def wr(sel):
        v = sum(r.get('Assessable Value') or 0 for r in sel); t = sum(r['_tax'] for r in sel)
        return (100 * t / v if v else None), v, t
    ti = [r for r in recs if r['Document Type'].lower().startswith('tax')]
    clean = [r for r in ti if not any(f[0] == 'LOW_RATE' for f in r['_flags'])]
    pre = [r for r in clean if _pdate(r['_period']) < datetime.date(2025, 9, 1)]
    post = [r for r in clean if _pdate(r['_period']) >= datetime.date(2025, 10, 1)]
    regime = dict(pre=wr(pre), post=wr(post), sep=wr([r for r in clean if r['_period'] == 'Sep-2025']))
    # conflicts
    conflicts = [r for r in recs if r['_differ']]
    # declared counts per page
    dec = []
    for b in blocks:
        if b.get('declared') is not None and b['layout'] == 'text in one cell':
            dec.append((b, b['declared'], b['rows']))
    # comparison with returns (when held)
    cmp_rows = _compare_returns(ing, months)
    hsn = _hsn_compare(ing, recs, rate_master)
    res = dict(available=True, copies=copies, copy_status=cstat, blocks=blocks, records=recs, months=months,
               median=med, regime=regime, conflicts=conflicts, declared=dec, compare=cmp_rows, hsn=hsn,
               first=span[0], last=span[-1],
               gstin=next((b['gstin'] for b in blocks if b.get('gstin')), None),
               legal=next((b['legal'] for b in blocks if b.get('legal')), None))
    order = {'HIGH': 0, 'MEDIUM': 1, 'NOTE': 2, 'INFO': 3}
    res['findings'] = sorted(ewb_findings(res), key=lambda f: (order.get(f['level'], 9), -(f.get('amount') or 0)))
    return res


def _compare_returns(ing, months):
    """EWB value of the month against GSTR-1 (R1_MONTHLY) and GSTR-3B (TAX_PAID taxable turnover), when held."""
    r1 = ing.by_period('R1_MONTHLY') if ing.have('R1_MONTHLY') else {}
    tp = ing.by_period('TAX_PAID') if ing.have('TAX_PAID') else {}
    if not r1 and not tp:
        return dict(available=False, missing=['R1_MONTHLY', 'TAX_PAID'], any_of=True)
    out = []
    for m in months:
        a, b = r1.get(m['period']), tp.get(m['period'])
        rv = _num(a.get('Taxable Value')) if a else None
        rt = _num(a.get('Total GST')) if a else None
        tv = _num(b.get('Taxable TO')) if b else None
        out.append(dict(period=m['period'], ewb=m['value'], ewb_tax=m['tax'], r1=rv, r1_tax=rt, r3b=tv))
    return dict(available=True, rows=out, have_r1=bool(r1), have_tp=bool(tp))


def _hsn_compare(ing, recs, rate_master):
    """Rate charged on the e-way bills against the rates of the HSNs declared in GSTR-1 Table 12 for the month,
    and against the rate master."""
    out = dict(have_hsn=ing.have('HSN'), months=[], candidates=[])
    if isinstance(rate_master, dict):
        rate_master = rate_master.get('rows')
    if rate_master:
        cand = {}
        for row in rate_master:
            h = str(row.get('HSN') or '')
            if not h[:2] in ('87', '40', '84', '85'): continue
            vf, vt, rt = row.get('Valid from'), row.get('Valid to'), row.get('IGST rate %')
            cand.setdefault(h, []).append((vf, vt, rt, row.get('Notification'), row.get('S.No.'),
                                           str(row.get('Description of goods (verbatim)') or '')[:160]))
        for h, lst in cand.items():
            before = [x for x in lst if x[0] and x[0] < RATE_CUT and (x[1] is None or x[1] >= datetime.date(2025, 9, 21))]
            after = [x for x in lst if x[0] and x[0] >= RATE_CUT]
            if any(x[2] == 28 for x in before) and any(x[2] == 18 for x in after) and h[:2] in ('87', '40'):
                b = next(x for x in before if x[2] == 28); a = next(x for x in after if x[2] == 18)
                out['candidates'].append(dict(hsn=h, before=b, after=a))
        out['candidates'].sort(key=lambda d: d['hsn'])
    if not out['have_hsn']:
        return out
    by = {}
    for h in ing.data['HSN']:
        p = h.get('_period'); v = _num(h.get('Taxable Value')) or 0; t = _num(h.get('Total GST')) or 0
        c = _num(h.get('CESS')) or 0
        by.setdefault(p, []).append((str(h.get('HSN Code')), v, t - c, (100 * (t - c) / v) if v else None))
    for p in sorted({r['_period'] for r in recs}, key=_pkey):
        rs = [r for r in recs if r['_period'] == p and r['Document Type'].lower().startswith('tax')]
        v = sum(r.get('Assessable Value') or 0 for r in rs); t = sum(r['_tax'] for r in rs)
        ew = 100 * t / v if v else None
        lines = by.get(p, [])
        goods = [x for x in lines if not x[0].startswith('99')]
        rates = [x[3] for x in lines if x[3] is not None and x[1] > 0]
        reading = ''
        if not lines:
            reading = 'No HSN line in GSTR-1 for this month.' if v else ''
        elif v and not goods:
            reading = ('GSTR-1 declares only services (SAC 99..) this month, but e-way bills moved goods worth '
                       f'{_lakh(v)} at {ew:.2f}%. If these e-way bills are outward, the goods are not declared under a goods HSN.')
        elif ew is not None and rates and ew > max(rates) + 0.5:
            reading = (f'E-way bills carry {ew:.2f}%, above every rate declared in GSTR-1 HSN lines ({", ".join(f"{x:.1f}%" for x in sorted(set(round(r, 1) for r in rates)))}). '
                       'Possible misclassification to a lower-rated HSN, or supplies missing from GSTR-1.')
        out['months'].append(dict(period=p, ewb_value=v, ewb_rate=ew, lines=lines, reading=reading))
    return out


# ------------------------------------------------------------------ findings (words)
def ewb_findings(res):
    R = res['records']; F = []
    def add(level, head, rows, text, basis, action, amount=None):
        F.append(dict(level=level, head=head, periods=', '.join(dict.fromkeys(r['_period'] for r in rows)) if rows else '',
                      text=text, basis=basis, action=action, amount=amount, keys=[r['_key'] for r in rows]))
    bos = [r for r in R if any(f[0] == 'BOS_TAX' for f in r['_flags'])]
    if bos:
        v = sum(r['Assessable Value'] or 0 for r in bos); t = sum(r['_tax'] for r in bos)
        n = sum(r['No of EWBs'] or 0 for r in bos)
        add('HIGH', 'Bills of supply carrying GST', bos,
            f'{int(n)} e-way bill(s) were raised on a BILL OF SUPPLY yet show tax of Rs.{_inr(t)} on a value of Rs.{_inr(v)} '
            f'(about {100 * t / v:.1f}%). The taxpayer is a regular taxpayer. A bill of supply carries no tax, so either the '
            f'supply was taxable and should have been on a tax invoice (and reported in GSTR-1 as such), or it was exempt '
            f'and the tax shown is wrong. Values are round and repeat (e.g. Rs.{_inr((bos[-1]["_per_ewb"] or 0))} per EWB in '
            f'{bos[-1]["_period"]}), which points to a recurring charge rather than a sale of stock.',
            'Section 31(3)(c) CGST Act and Rule 49 (bill of supply); s.31(1) and Rule 46 (tax invoice); Rule 138 (e-way bill).',
            'Get the e-way bill detail (Part A: HSN, value, consignee) for these months and match each with GSTR-1 (B2B/B2C '
            'and Table 8 exempt). Ask the taxpayer what moved on a bill of supply.', amount=t)
    low = [r for r in R if any(f[0] == 'LOW_RATE' for f in r['_flags'])]
    if low:
        v = sum(r['Assessable Value'] or 0 for r in low); t = sum(r['_tax'] for r in low)
        at_pre = res['regime']['pre'][0]
        gap = (v * (at_pre or 0) / 100 - t) if at_pre else None
        add('HIGH', 'High-value movements taxed at a fraction of the normal rate', low,
            f'{len(low)} month/document line(s) moved goods worth Rs.{_inr(v)} on tax invoices with tax of only '
            f'Rs.{_inr(t)} ({100 * t / v:.2f}%): '
            + '; '.join(f'{r["_period"]} {r["Category"]}: {int(r["No of EWBs"])} EWB, {_lakh(r["_per_ewb"] or 0)} each, '
                        f'tax {r["_rate"]:.2f}%' for r in low)
            + '. In its other months the same taxpayer\'s tax invoices carry about '
            + (f'{at_pre:.1f}%' if at_pre else 'the normal rate')
            + '. Large values with almost no tax fit (a) sale of used vehicles or other capital goods taxed only on the '
              'margin (Rule 32(5); Notification 8/2018-Central Tax (Rate): sale value less depreciated value), '
              '(b) goods moved on a tax invoice where most of the value is exempt or not a supply, or (c) under-declared tax.'
            + (f' At the taxpayer\'s usual rate of {at_pre:.1f}% the tax on this value would be Rs.{_inr(gap + t)}, '
               f'Rs.{_inr(gap)} more than shown; at 18% it would be Rs.{_inr(v * 0.18 - t)} more. These are outer limits, '
               f'not a demand — the margin rule, if it applies, can bring it close to nil.' if gap else ''),
            'Rule 32(5) CGST Rules; Notification 8/2018-Central Tax (Rate); s.18(6) CGST Act (supply of capital goods on '
            'which ITC was taken: higher of ITC reduced by 5% per quarter and tax on transaction value); s.15 (value).',
            'Pull the e-way bills for these months: HSN, description, consignee, vehicle. If used buses were sold, get the '
            'purchase invoices, ITC taken on them (GSTR-2A/3B of the purchase year) and depreciation schedule; test s.18(6).',
            amount=gap)
    bigs = [r for r in R if (r.get('_per_ewb') or 0) >= 1_000_000 and r['Document Type'].lower().startswith('tax')]
    big_low = [r for r in bigs if r['_rate'] is not None and r['_rate'] < 2]
    big_norm = [r for r in bigs if r['_rate'] is not None and r['_rate'] >= 5]
    if big_low and big_norm:
        add('HIGH', 'Consignments of the same size taxed very differently', big_low + big_norm,
            'Lines whose average e-way bill is Rs.10 lakh or more — the size of a vehicle or a vehicle body — are taxed at '
            'the full rate in some months and at almost nothing in others. Full rate: '
            + '; '.join(f'{r["_period"]} {r["Category"]} {int(r["No of EWBs"])} x {_lakh(r["_per_ewb"])} at {r["_rate"]:.2f}%'
                        for r in big_norm)
            + '. Almost nil: '
            + '; '.join(f'{r["_period"]} {r["Category"]} {int(r["No of EWBs"])} x {_lakh(r["_per_ewb"])} at {r["_rate"]:.2f}%'
                        for r in big_low)
            + '. If the goods are alike, the low-tax months need an explanation (margin scheme for used goods, s.18(6), '
              'or tax short-declared).',
            'Section 15 CGST Act (value); Rule 32(5); s.18(6).',
            'Compare HSN, description and buyer of the e-way bills in the two groups.')
    cess = [r for r in R if (r.get('CESS') or 0) > 0]
    if cess:
        c = sum(r['CESS'] for r in cess)
        add('MEDIUM', 'Compensation cess on e-way bills', cess,
            'Compensation cess of Rs.' + _inr(c) + ' appears in ' + ', '.join(
                f'{r["_period"]} ({r["_cess_rate"]:.2f}% of Rs.{_inr(r["Assessable Value"])}, GST {r["_rate"]:.2f}%)' for r in cess)
            + ('. It is the only cess in the whole period' if len(cess) == 1 else '')
            + '. In the cess schedule 12% is the rate for aerated waters (2202 10); no motor-vehicle entry is 12% (8702 buses '
              '15%, 8703 cars nil/1%/3%/15% and more by later amendments, 8711 motorcycles over 350 cc 3%; goods vehicles, '
              'bodies and parts carry none). Either different goods moved that month, or the cess field was filled wrongly. '
              'If cess was charged, check it was paid in GSTR-3B and that cess credit was not used against CGST/SGST/IGST.',
            'Section 8 GST (Compensation to States) Act, 2017; Notification 1/2017-Compensation Cess (Rate); cess made nil on '
            'motor vehicles and most goods from 22-09-2025 (Notification 2/2025-Compensation Cess (Rate)).',
            'Pull the e-way bills of that month and check HSN and cess against GSTR-1 and GSTR-3B (cess column).')
    rg = res['regime']
    if rg['pre'][0] and rg['post'][0]:
        add('INFO', 'Rate profile: before and after 22-09-2025', [],
            f'On tax invoices (high-value low-rate lines left out), the value-weighted rate was {rg["pre"][0]:.2f}% up to '
            f'Aug-2025 (a mix of 28% and lower-rated goods) and {rg["post"][0]:.2f}% from Oct-2025' + (f' (Sep-2025, the month of the change: {rg["sep"][0]:.2f}%)' if rg['sep'][0] else '')
            + '. A fall from about 28% to about 18% on 22-09-2025 fits goods moved from 28% to 18% by Notification '
              '9/2025-Integrated Tax (Rate) — in a transport business: buses (8702), goods vehicles (8704), bodies (8707), '
              'parts (8708) or tyres (4011). See 19M_EWB_HSN_RATE for the rate-master rows. '
            + ('Compare these with the HSN codes the taxpayer declares in GSTR-1.' if not res['hsn']['have_hsn'] else ''),
            'Notification 1/2017 and 9/2025-Integrated Tax (Rate).', 'Use the HSN-wise e-way bill detail to fix the HSN.')
    late28 = [r for r in R if _pdate(r['_period']) >= datetime.date(2025, 10, 1) and r['_rate'] and r['_rate'] > 20
              and (r['Assessable Value'] or 0) > 0]
    if late28:
        add('MEDIUM', 'Rate above 18% after the 22-09-2025 cut', late28,
            'After Sep-2025 these lines still carry more than 20%: ' + '; '.join(
                f'{r["_period"]} {r["Document Type"]} {r["Category"]} {r["_rate"]:.2f}%' for r in late28)
            + '. Either goods that stayed at 28% / moved to 40%, or the old rate kept on the invoice.',
            'Notification 9/2025-Integrated Tax (Rate).', 'Check HSN in the e-way bill detail.')
    mixed_inter = [r for r in R if any(f[0] == 'MIXED' for f in r['_flags']) and r['Category'] == 'INTER'
                   and r['_rate'] is not None and r['_rate'] < 12 and (r['Assessable Value'] or 0) > 200000]
    if mixed_inter:
        add('MEDIUM', 'Inter-State movements at a low blended rate', mixed_inter,
            'Inter-State lines with an effective IGST rate between slabs and well below the usual rate: ' + '; '.join(
                f'{r["_period"]} {r["_rate"]:.2f}% on Rs.{_inr(r["Assessable Value"])}' for r in mixed_inter)
            + '. Part of the value may be untaxed (exempt goods, or goods sent for repair / job work / own use on a tax '
              'invoice) or carry a lower rate.', 'Section 5 IGST Act; s.15 CGST Act.',
            'Pull the e-way bills and check HSN-wise value and tax.')
    spikes = [m for m in res['months'] if m['spike']]
    if spikes:
        add('MEDIUM', 'Months far above the usual level', [],
            f'Median month: Rs.{_inr(res["median"])} of e-way bills. These months are more than {SPIKE:.0f} times that: '
            + '; '.join(f'{m["period"]} Rs.{_inr(m["value"])}' for m in spikes) + '.',
            '—', 'Match these months with GSTR-1 and GSTR-3B turnover first; they decide most of the value.')
    gaps = [m['period'] for m in res['months'] if m['gap']]
    if gaps:
        add('INFO', 'Months with no e-way bill', [],
            'No e-way bill in ' + ', '.join(gaps) + ' (inside the period covered). Compare with GSTR-1 for those months: '
            'supplies without e-way bills over Rs.50,000 a consignment are a lead.',
            'Rule 138(1) CGST Rules.', 'Check GSTR-1 of these months.')
    zero = [r for r in R if any(f[0] == 'ZERO' for f in r['_flags'])]
    if zero:
        add('INFO', 'E-way bills with nil value', zero,
            '; '.join(f'{r["_period"]} {r["Category"]}: {int(r["No of EWBs"])} EWB with value 0' for r in zero)
            + '. A nil value on a tax invoice is not possible for a supply; check the document.', 'Rule 138.', '—')
    if res['conflicts']:
        add('NOTE', 'The same month read twice with different figures', res['conflicts'],
            'Overlapping pages gave different figures for: ' + '; '.join(
                f'{r["_period"]} {r["Document Type"]} {r["Category"]}' for r in res['conflicts'])
            + '. The page with the latest window is used. Download that window again to settle it.', '—', '—')
    cmpr = res['compare']
    if cmpr.get('available'):
        over = [x for x in cmpr['rows'] if x['ewb'] and x['r1'] is not None and x['ewb'] > x['r1'] + max(100000, 0.1 * x['r1'])]
        nor1 = [x for x in cmpr['rows'] if x['ewb'] and cmpr['have_r1'] and x['r1'] is None]
        if over or nor1:
            add('HIGH', 'E-way bill value above GSTR-1 (if outward)', [],
                ('Months where e-way bills exceed the GSTR-1 taxable value by over Rs.1 lakh and 10%: '
                 + '; '.join(f'{x["period"]}: EWB Rs.{_inr(x["ewb"])} vs GSTR-1 Rs.{_inr(x["r1"])}' for x in over) + '. ' if over else '')
                + ('Months with e-way bills but no GSTR-1 row: ' + ', '.join(x['period'] for x in nor1) + '.' if nor1 else ''),
                'Sections 37 and 39 CGST Act; Rule 138.', 'Confirm the direction of the e-way bills, then reconcile invoice-wise.',
                amount=sum(x['ewb'] - x['r1'] for x in over))
    hm = [m for m in res['hsn']['months'] if m['reading'] and ('misclassification' in m['reading']
                                                                  or 'not declared under a goods HSN' in m['reading'])]
    if hm:
        add('HIGH', 'HSN declared in GSTR-1 does not fit the e-way bills', [],
            f'{len(hm)} month(s): ' + ' | '.join(f'{m["period"]}: {m["reading"]}' for m in hm[:12])
            + (f' | ... and {len(hm) - 12} more (sheet 19M).' if len(hm) > 12 else ''),
            'Section 37 and Rule 59 (HSN in GSTR-1); the notified rate of the correct HSN.',
            'Get the HSN on each e-way bill and compare with GSTR-1 Table 12.')
        F[-1]['periods'] = ', '.join(m['period'] for m in hm)
    return F


def ewb_report_lines(R):
    """Paragraph for 21_REPORT. E-way bill leads are not added to the revenue total: they need the EWB detail first."""
    E = R.get('ewb') or {}
    if not E.get('available'):
        return []
    v = sum(r['Assessable Value'] or 0 for r in E['records'])
    T = ['2E. E-WAY BILLS', '',
         (f'2E.1  The GST-Prime e-way bill document-wise summary for {E["first"]} to {E["last"]} was read from '
          f'{len(E["blocks"])} page block(s): {len(E["copies"])} rows, of which {len(E["copies"]) - len(E["records"])} were '
          f'repeats from overlapping windows and are counted once. {len(E["records"])} records remain: '
          f'{int(sum(r["No of EWBs"] or 0 for r in E["records"]))} e-way bills, value Rs.{_inr(v)} (sheets 19H to 19K). '
          'The summary does not state whether the taxpayer is consignor or consignee.'), '']
    for i, f in enumerate([f for f in E['findings'] if f['level'] in ('HIGH', 'MEDIUM')], 2):
        T += [f'2E.{i}  {f["head"]} ({f["periods"] or "see 19L"}). {f["text"]}', '']
    T += ['2E.x  These are leads for the e-way bill detail and are not added to the revenue figure in paragraph 4 '
          '(sheet 19L).', '']
    return T


# ------------------------------------------------------------------ workbook sheets
def _styles():
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    t = Side(style='thin')
    return dict(B=Font(bold=True), BD=Border(left=t, right=t, top=t, bottom=t),
                W=Alignment(wrap_text=True, vertical='top'),
                RED=PatternFill('solid', fgColor='FFD9D9'), AMB=PatternFill('solid', fgColor='FFF2CC'),
                ORG=PatternFill('solid', fgColor='F8CBAD'), GRN=PatternFill('solid', fgColor='E2EFDA'),
                GRY=PatternFill('solid', fgColor='EDEDED'))

def _head(ws, r, hdr, widths=None, S=None):
    from openpyxl.utils import get_column_letter
    from openpyxl.styles import Alignment
    for j, h in enumerate(hdr, 1):
        c = ws.cell(row=r, column=j, value=h); c.font = S['B']; c.border = S['BD']; c.fill = S['GRY']
        c.alignment = Alignment(wrap_text=True, vertical='top', horizontal='center')
    if widths:
        for i, w in enumerate(widths, 1): ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = ws.cell(row=r + 1, column=1)

def _note(ws, r, text, S, bold=False, cols=None):
    c = ws.cell(row=r, column=1, value=text); c.alignment = S['W']
    if bold: c.font = S['B']
    if cols:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=cols)
        ws.row_dimensions[r].height = max(15, 15 * (1 + len(text) // (cols * 14)))
    return r + 1

def build_ewb_sheets(wb, E, mark=True, r1_lookup=None):
    """Writes 19H..19M into workbook wb from E = chk_ewb(...). r1_lookup(period_cell, field) -> formula or None."""
    from openpyxl.utils import get_column_letter as L
    S = _styles()
    if not E.get('available'):
        ws = wb.create_sheet(SHEETS[0])
        ws['A1'] = 'E-WAY BILLS'; ws['A1'].font = S['B']
        ws['A3'] = ('THIS TEST COULD NOT BE RUN. The GST-Prime E-Way Bill DOCUMENT-WISE SUMMARY has not been uploaded. '
                    'Save each page of it (Excel or copy-paste into a sheet) and upload it.')
        if mark: ws['A3'].fill = S['RED']
        ws.column_dimensions['A'].width = 130
        return
    NUMF = '#,##0'
    # ---------------------------------------------------------------- 19H every copy
    ws = wb.create_sheet(SHEETS[0])
    ws['A1'] = 'E-WAY BILLS — EVERY ROW READ, FROM EVERY PAGE (overlapping windows included)'; ws['A1'].font = S['B']
    r = _note(ws, 2, 'Each row is exactly as the page showed it. The same month can appear on two pages because the '
                     'windows overlap; column N marks which copy is used. A repeat with the same figures is NOT counted '
                     'again; a repeat with different figures is listed in 19L as a conflict.', S, cols=16)
    r += 1
    ws.cell(row=r, column=1, value='PAGES (BLOCKS) READ').font = S['B']; r += 1
    BH = ['Block', 'File', 'Sheet', 'Layout', 'Starts at', 'Window printed on page', 'First month', 'Last month',
          'Rows read', 'Records declared by page', 'GSTIN on page', 'Legal name on page', 'Office']
    _head(ws, r, BH, S=S); bh = r; r += 1
    for b in E['blocks']:
        vals = [b['id'], b['file'], b['sheet'], b['layout'], b['start'],
                (f'{b["window"][0]} to {b["window"][1]}' if b.get('window') else 'not printed (month picker only)'),
                b['first'], b['last'], b['rows'], b.get('declared'), b.get('gstin') or 'not printed', b.get('legal'),
                b.get('office')]
        for j, v in enumerate(vals, 1):
            c = ws.cell(row=r, column=j, value=v); c.border = S['BD']; c.alignment = S['W']
        if b.get('declared') is not None and b['layout'] == 'text in one cell' and b['declared'] != b['rows'] and mark:
            ws.cell(row=r, column=10).fill = S['ORG']
        r += 1
    r += 1
    ws.cell(row=r, column=1, value='ROWS').font = S['B']; r += 1
    H = ['Copy no.', 'Block', 'Read from (sheet!cell)', 'SNo on page', 'Return Period', 'Document Type', 'Category',
         'No of EWBs', 'Assessable Value', 'SGST', 'CGST', 'IGST', 'CESS', 'GST', 'Key (period|document|category)',
         'Times this key read so far', 'Copy used?', 'GST = SGST+CGST+IGST+CESS?']
    _head(ws, r, H, [8, 7, 26, 7, 11, 15, 9, 9, 15, 13, 13, 13, 12, 14, 36, 10, 40, 14], S=S)
    a0 = r + 1
    for i, c in enumerate(E['copies']):
        rr = a0 + i
        vals = [i + 1, c['_block'], c['_where'], c['SNo'], c['_period'], c['Document Type'], c['Category']] + \
               [c.get(k) for k in FIG]
        for j, v in enumerate(vals, 1):
            x = ws.cell(row=rr, column=j, value=v); x.border = S['BD']
            if j >= 8: x.number_format = NUMF
            if v is None and j >= 8 and mark: x.fill = S['RED']
        ws.cell(row=rr, column=15, value=f'=E{rr}&"|"&F{rr}&"|"&G{rr}').border = S['BD']
        ws.cell(row=rr, column=16, value=f'=COUNTIF($O${a0}:O{rr},O{rr})').border = S['BD']
        st = E['copy_status'][i]
        x = ws.cell(row=rr, column=17, value=st); x.border = S['BD']
        if mark and st != 'USED': x.fill = S['AMB'] if 'SAME' in st else S['ORG']
        ws.cell(row=rr, column=18, value=f'=IF(ABS(J{rr}+K{rr}+L{rr}+M{rr}-N{rr})<=1,"YES","NO")').border = S['BD']
    a1 = a0 + len(E['copies']) - 1
    ws.auto_filter.ref = f'A{a0 - 1}:R{a1}'
    ws.freeze_panes = f'A{a0}'
    E['_a'] = (a0, a1)
    A = f"'{SHEETS[0]}'"
    # ---------------------------------------------------------------- 19I records
    ws = wb.create_sheet(SHEETS[1])
    ws['A1'] = 'E-WAY BILLS — ONE RECORD PER MONTH, DOCUMENT TYPE AND CATEGORY (no row counted twice)'; ws['A1'].font = S['B']
    _note(ws, 2, 'Columns H-N are pulled live from the copy marked USED in 19H (column G gives its copy number). '
                 'Column E counts how often the key was read; column F how many of those copies carry the same figures. '
                 'E = F means every repeat agreed.', S, cols=16)
    H = ['No.', 'Return Period', 'FY', 'Document Type', 'Category', 'Times read', 'Copies with same figures',
         'Copy used (19H Copy no.)', 'No of EWBs', 'Assessable Value', 'SGST', 'CGST', 'IGST', 'CESS',
         'Tax (SGST+CGST+IGST)', 'Effective rate %', 'Cess % of value', 'Value per EWB', 'Nearest GST slab %', 'Flags']
    _head(ws, 4, H, [6, 11, 9, 15, 9, 8, 10, 10, 9, 15, 13, 13, 13, 12, 14, 10, 9, 13, 9, 90], S=S)
    b0 = 5
    for i, rec in enumerate(E['records']):
        rr = b0 + i
        ws.cell(row=rr, column=1, value=i + 1)
        ws.cell(row=rr, column=2, value=rec['_period']); ws.cell(row=rr, column=3, value=_fy(rec['_period']))
        ws.cell(row=rr, column=4, value=rec['Document Type']); ws.cell(row=rr, column=5, value=rec['Category'])
        key = f'B{rr}&"|"&D{rr}&"|"&E{rr}'
        ws.cell(row=rr, column=6, value=f'=COUNTIF({A}!$O${a0}:$O${a1},{key})')
        used_row = a0 + rec['_keep']
        ws.cell(row=rr, column=7, value=f'=COUNTIFS({A}!$O${a0}:$O${a1},{key},{A}!$H${a0}:$H${a1},I{rr},'
                                         f'{A}!$I${a0}:$I${a1},J{rr},{A}!$N${a0}:$N${a1},{A}!N{used_row})')
        ws.cell(row=rr, column=8, value=rec['_keep'] + 1)
        for j, col in enumerate('HIJKLM', 9):
            ws.cell(row=rr, column=j, value=f'=INDEX({A}!${col}${a0}:${col}${a1},$H{rr})')
        ws.cell(row=rr, column=15, value=f'=K{rr}+L{rr}+M{rr}')
        ws.cell(row=rr, column=16, value=f'=IF(J{rr}=0,"",ROUND(O{rr}/J{rr}*100,2))')
        ws.cell(row=rr, column=17, value=f'=IF(J{rr}=0,"",ROUND(N{rr}/J{rr}*100,2))')
        ws.cell(row=rr, column=18, value=f'=IF(I{rr}=0,"",J{rr}/I{rr})')
        ws.cell(row=rr, column=19, value=rec['_slab'])
        ftxt = ' | '.join(f'[{lv}] {t}' for _, lv, t in rec['_flags'])
        ws.cell(row=rr, column=20, value=ftxt or '—')
        for j in range(1, 21):
            x = ws.cell(row=rr, column=j); x.border = S['BD']
            if j in (9, 10, 11, 12, 13, 14, 15, 18): x.number_format = NUMF
            if j == 20: x.alignment = S['W']
        if mark and any(f[1] == 'HIGH' for f in rec['_flags']):
            for j in (2, 4, 5, 20): ws.cell(row=rr, column=j).fill = S['ORG']
        elif mark and any(f[1] == 'MEDIUM' for f in rec['_flags']):
            for j in (2, 4, 5, 20): ws.cell(row=rr, column=j).fill = S['AMB']
        if mark and rec['_differ']:
            ws.cell(row=rr, column=7).fill = S['RED']
    b1 = b0 + len(E['records']) - 1
    rr = b1 + 1
    ws.cell(row=rr, column=4, value='TOTAL').font = S['B']
    for j in (9, 10, 11, 12, 13, 14, 15):
        c = ws.cell(row=rr, column=j, value=f'=SUM({L(j)}{b0}:{L(j)}{b1})'); c.font = S['B']; c.number_format = NUMF
    ws.cell(row=rr, column=16, value=f'=ROUND(O{rr}/J{rr}*100,2)').font = S['B']
    ws.cell(row=rr + 1, column=4, value='Rows read in 19H'); ws.cell(row=rr + 1, column=6, value=f'=COUNTA({A}!E{a0}:E{a1})')
    ws.cell(row=rr + 2, column=4, value='Repeats not counted'); ws.cell(row=rr + 2, column=6, value=f'=F{rr + 1}-COUNTA(B{b0}:B{b1})')
    ws.auto_filter.ref = f'A4:T{b1}'
    E['_b'] = (b0, b1)
    B = f"'{SHEETS[1]}'"
    # ---------------------------------------------------------------- 19J monthly
    ws = wb.create_sheet(SHEETS[2])
    ws['A1'] = 'E-WAY BILLS — MONTH BY MONTH (every month in the period, including months with none)'; ws['A1'].font = S['B']
    have_cmp = E['compare'].get('available') and r1_lookup is not None
    H = ['Return Period', 'FY', 'No of EWBs', 'Assessable Value', 'of which Tax Invoice', 'of which Bill of Supply',
         'of which other documents', 'Inter-State value', 'Tax (SGST+CGST+IGST)', 'Tax on Bills of Supply', 'CESS',
         'Effective rate %', 'Times the median month', 'Reading']
    if have_cmp:
        H += ['GSTR-1 taxable value', 'EWB − GSTR-1', 'GSTR-1 tax', 'GSTR-3B taxable turnover', 'EWB − GSTR-3B']
    _head(ws, 4, H, [11, 9, 9, 15, 15, 14, 12, 14, 14, 13, 12, 9, 9, 60] + [15] * 5, S=S)
    _note(ws, 2, 'Every figure is a SUMIFS over 19I_EWB_RECORDS, so each one traces to the copy used and the page it '
                 'came from.' + (' GSTR-1 / GSTR-3B columns read the source sheets of those returns.' if have_cmp else
                                 ' Upload GSTR-1 (R1 FILINGS IN THE YEAR) and TAX PAID DETAILS to compare with returns.'),
          S, cols=14)
    c0 = 5
    med_cell = None
    for i, m in enumerate(E['months']):
        rr = c0 + i; p = f'$A{rr}'
        ws.cell(row=rr, column=1, value=m['period']); ws.cell(row=rr, column=2, value=m['fy'])
        def sf(col, extra=''):
            return f'=SUMIFS({B}!${col}${b0}:${col}${b1},{B}!$B${b0}:$B${b1},{p}{extra})'
        ws.cell(row=rr, column=3, value=sf('I'))
        ws.cell(row=rr, column=4, value=sf('J'))
        ws.cell(row=rr, column=5, value=sf('J', f',{B}!$D${b0}:$D${b1},"Tax Invoice"'))
        ws.cell(row=rr, column=6, value=sf('J', f',{B}!$D${b0}:$D${b1},"Bill Of Supply"'))
        ws.cell(row=rr, column=7, value=f'=D{rr}-E{rr}-F{rr}')
        ws.cell(row=rr, column=8, value=sf('J', f',{B}!$E${b0}:$E${b1},"INTER"'))
        ws.cell(row=rr, column=9, value=sf('O'))
        ws.cell(row=rr, column=10, value=sf('O', f',{B}!$D${b0}:$D${b1},"Bill Of Supply"'))
        ws.cell(row=rr, column=11, value=sf('N'))
        ws.cell(row=rr, column=12, value=f'=IF(D{rr}=0,"",ROUND(I{rr}/D{rr}*100,2))')
        ws.cell(row=rr, column=13, value=f'=IF(OR(D{rr}=0,$D$MED=0),"",ROUND(D{rr}/$D$MED,1))')
        rd = []
        if m['gap']: rd.append('NO E-WAY BILL this month')
        if m['spike']: rd.append(f'more than {SPIKE:.0f}x the median month')
        for rec in E['records']:
            if rec['_period'] == m['period']:
                for code, lv, t in rec['_flags']:
                    if lv in ('HIGH', 'MEDIUM'):
                        rd.append(f'{rec["Document Type"]} {rec["Category"]}: {code.replace("_", " ").lower()}')
        ws.cell(row=rr, column=14, value='; '.join(rd) or '—')
        if have_cmp:
            ws.cell(row=rr, column=15, value=r1_lookup(p, 'R1_MONTHLY', 'Taxable Value'))
            ws.cell(row=rr, column=16, value=f'=IF(O{rr}="","",D{rr}-O{rr})')
            ws.cell(row=rr, column=17, value=r1_lookup(p, 'R1_MONTHLY', 'Total GST'))
            ws.cell(row=rr, column=18, value=r1_lookup(p, 'TAX_PAID', 'Taxable TO'))
            ws.cell(row=rr, column=19, value=f'=IF(R{rr}="","",D{rr}-R{rr})')
        ncol = 19 if have_cmp else 14
        for j in range(1, ncol + 1):
            x = ws.cell(row=rr, column=j); x.border = S['BD']
            if j in (3, 4, 5, 6, 7, 8, 9, 10, 11, 15, 16, 17, 18, 19): x.number_format = NUMF
            if j == 14: x.alignment = S['W']
        if mark and m['gap']:
            for j in range(1, 15): ws.cell(row=rr, column=j).fill = S['RED']
        elif mark and (m['spike'] or rd):
            ws.cell(row=rr, column=14).fill = S['AMB']
    c1 = c0 + len(E['months']) - 1
    tr = c1 + 1
    ws.cell(row=tr, column=1, value='TOTAL').font = S['B']
    for j in (3, 4, 5, 6, 7, 8, 9, 10, 11):
        x = ws.cell(row=tr, column=j, value=f'=SUM({L(j)}{c0}:{L(j)}{c1})'); x.font = S['B']; x.number_format = NUMF
    ws.cell(row=tr, column=12, value=f'=ROUND(I{tr}/D{tr}*100,2)').font = S['B']
    mr = tr + 1
    ws.cell(row=mr, column=1, value='Median month (months with e-way bills)').font = S['B']
    from openpyxl.worksheet.formula import ArrayFormula
    ws.cell(row=mr, column=4).value = ArrayFormula(f'D{mr}', f'=MEDIAN(IF(D{c0}:D{c1}>0,D{c0}:D{c1}))')
    ws.cell(row=mr, column=4).number_format = NUMF
    for i in range(len(E['months'])):
        rr = c0 + i
        ws.cell(row=rr, column=13, value=ws.cell(row=rr, column=13).value.replace('$D$MED', f'$D${mr}'))
    ws.cell(row=mr + 1, column=1, value='Check: total value here = total in 19I')
    ws.cell(row=mr + 1, column=4, value=f'=IF(ABS(D{tr}-{B}!J{b1 + 1})<1,"AGREES","DIFFERS")')
    ws.freeze_panes = 'C5'
    E['_c'] = (c0, c1)
    C = f"'{SHEETS[2]}'"
    # ---------------------------------------------------------------- 19K FY
    ws = wb.create_sheet(SHEETS[3])
    ws['A1'] = 'E-WAY BILLS — FINANCIAL YEAR TOTALS'; ws['A1'].font = S['B']
    _note(ws, 2, 'SUMIFS over 19J. A financial year at either end may be part-covered: see the months column.', S, cols=10)
    H = ['FY', 'Months in data', 'Months with no EWB', 'No of EWBs', 'Assessable Value', 'Tax Invoice value',
         'Bill of Supply value', 'Inter-State value', 'Tax', 'Tax on Bills of Supply', 'CESS', 'Effective rate %']
    _head(ws, 4, H, [10, 9, 9, 10, 16, 16, 14, 14, 14, 13, 12, 10], S=S)
    fys = list(dict.fromkeys(m['fy'] for m in E['months']))
    for i, f in enumerate(fys):
        rr = 5 + i
        ws.cell(row=rr, column=1, value=f)
        ws.cell(row=rr, column=2, value=f'=COUNTIF({C}!$B${c0}:$B${c1},A{rr})')
        ws.cell(row=rr, column=3, value=f'=COUNTIFS({C}!$B${c0}:$B${c1},A{rr},{C}!$D${c0}:$D${c1},0)')
        for j, col in zip((4, 5, 6, 7, 8, 9, 10, 11), 'CDEFHIJK'):
            ws.cell(row=rr, column=j, value=f'=SUMIFS({C}!${col}${c0}:${col}${c1},{C}!$B${c0}:$B${c1},$A{rr})')
        ws.cell(row=rr, column=12, value=f'=IF(E{rr}=0,"",ROUND(I{rr}/E{rr}*100,2))')
        for j in range(1, 13):
            x = ws.cell(row=rr, column=j); x.border = S['BD']
            if 4 <= j <= 11: x.number_format = NUMF
    d1 = 4 + len(fys); tr = d1 + 1
    ws.cell(row=tr, column=1, value='TOTAL').font = S['B']
    for j in range(2, 12):
        x = ws.cell(row=tr, column=j, value=f'=SUM({L(j)}5:{L(j)}{d1})'); x.font = S['B']; x.number_format = NUMF
    ws.cell(row=tr, column=12, value=f'=ROUND(I{tr}/E{tr}*100,2)').font = S['B']
    r = tr + 2
    ws.cell(row=r, column=1, value='RATE BEFORE AND AFTER 22-09-2025 (tax invoices; high-value low-rate lines left out)').font = S['B']
    r += 1
    _head(ws, r, ['Period', 'Value', 'Tax', 'Rate %'], S=S); ws.freeze_panes = 'A5'; r += 1
    for lab, k in (('Up to Aug-2025', 'pre'), ('Sep-2025 (month of change)', 'sep'), ('From Oct-2025', 'post')):
        rt, v, t = E['regime'][k]
        ws.cell(row=r, column=1, value=lab); ws.cell(row=r, column=2, value=v); ws.cell(row=r, column=3, value=t)
        ws.cell(row=r, column=4, value=f'=IF(B{r}=0,"",ROUND(C{r}/B{r}*100,2))')
        for j in range(1, 5):
            ws.cell(row=r, column=j).border = S['BD']
            if j in (2, 3): ws.cell(row=r, column=j).number_format = NUMF
        r += 1
    _note(ws, r, 'These three lines are totals computed by the engine from 19I rows whose Document Type is Tax Invoice '
                 'and which carry no "low rate" flag; filter 19I the same way to check them.', S, cols=10)
    # ---------------------------------------------------------------- 19L findings
    ws = wb.create_sheet(SHEETS[4])
    ws['A1'] = 'E-WAY BILLS — FINDINGS AND LEADS'; ws['A1'].font = S['B']
    _note(ws, 2, 'The document-wise summary does not say whether this GSTIN generated these e-way bills as supplier '
                 '(outward) or received them (inward). Read every finding with that in mind, and confirm it from the '
                 'e-way bill detail before acting. Amounts are leads, not demands.', S, bold=True, cols=7)
    H = ['No.', 'Level', 'Finding', 'Months', 'What the figures show', 'Law', 'Next step', 'Amount the lead turns on (Rs.) — not a demand']
    _head(ws, 4, H, [5, 9, 30, 22, 90, 40, 50, 16], S=S)
    for i, f in enumerate(E['findings']):
        rr = 5 + i
        vals = [i + 1, f['level'], f['head'], f['periods'], f['text'], f['basis'], f['action'],
                round(f['amount']) if f.get('amount') else None]
        for j, v in enumerate(vals, 1):
            x = ws.cell(row=rr, column=j, value=v); x.border = S['BD']; x.alignment = S['W']
            if j == 8: x.number_format = NUMF
        if mark and f['level'] == 'HIGH':
            ws.cell(row=rr, column=2).fill = S['ORG']
        elif mark and f['level'] == 'MEDIUM':
            ws.cell(row=rr, column=2).fill = S['AMB']
    # ---------------------------------------------------------------- 19M HSN and rates
    ws = wb.create_sheet(SHEETS[5])
    ws['A1'] = 'E-WAY BILLS — HSN AND RATE'; ws['A1'].font = S['B']
    hs = E['hsn']; r = 3
    if hs['have_hsn']:
        r = _note(ws, r, 'Rate on the e-way bills (tax invoices) against the HSN lines declared in GSTR-1 Table 12 for the '
                         'same month.', S, bold=True, cols=6)
        _head(ws, r, ['Return Period', 'EWB value (tax invoices)', 'EWB rate %', 'HSN lines declared in GSTR-1 '
                      '(HSN: value @ rate)', 'Reading'], [12, 16, 10, 70, 80], S=S); r += 1
        for m in hs['months']:
            vals = [m['period'], m['ewb_value'], round(m['ewb_rate'], 2) if m['ewb_rate'] is not None else None,
                    '; '.join(f'{h}: {_inr(v)} @ {rt:.2f}%' if rt is not None else f'{h}: {_inr(v)}' for h, v, t, rt in m['lines']) or '—',
                    m['reading'] or '—']
            for j, v in enumerate(vals, 1):
                x = ws.cell(row=r, column=j, value=v); x.border = S['BD']; x.alignment = S['W']
                if j == 2: x.number_format = NUMF
            if mark and m['reading'] and m['reading'] != '—': ws.cell(row=r, column=5).fill = S['AMB']
            r += 1
        r += 1
    else:
        x = ws.cell(row=r, column=1, value='GSTR-1 HSN SUMMARY (HSN_LIST) NOT UPLOADED for this taxpayer — the HSN declared '
                                           'cannot be compared with the e-way bills. Upload HSN_LIST to get the month-wise test.')
        x.font = S['B']
        if mark: x.fill = S['RED']
        r += 2
    r = _note(ws, r, 'RATE FINGERPRINT', S, bold=True)
    rg = E['regime']
    r = _note(ws, r, f'Tax invoices: {rg["pre"][0]:.2f}% up to Aug-2025, {rg["post"][0]:.2f}% from Oct-2025.'
              if rg['pre'][0] and rg['post'][0] else 'Not enough tax-invoice months either side of 22-09-2025.', S, cols=6)
    r += 1
    if hs['candidates']:
        r = _note(ws, r, 'HSNs in the GST rate master (chapters 40 and 87) that were 28% up to 21-09-2025 and 18% from '
                         '22-09-2025 — the goods a 28%-to-18% fall can come from. Rate master rows are quoted as held.',
                  S, bold=True, cols=6)
        _head(ws, r, ['HSN', 'Rate to 21-09-2025', 'Notification / S.No.', 'Rate from 22-09-2025', 'Notification / S.No.',
                      'Description (rate master, from 22-09-2025)'], [12, 12, 40, 12, 40, 90], S=S); r += 1
        for c in hs['candidates']:
            b, a = c['before'], c['after']
            vals = [c['hsn'], b[2], f'{b[3]} / {b[4]}', a[2], f'{a[3]} / {a[4]}', a[5]]
            for j, v in enumerate(vals, 1):
                x = ws.cell(row=r, column=j, value=v); x.border = S['BD']; x.alignment = S['W']
            r += 1
        r += 1
    r = _note(ws, r, 'REFERENCE — rules used in this test (compiled; check the notification text before relying on it)',
              S, bold=True)
    _head(ws, r, ['Item', 'Rule', 'Source'], [30, 90, 70], S=S); r += 1
    for it in REFERENCE:
        for j, v in enumerate(it, 1):
            x = ws.cell(row=r, column=j, value=v); x.border = S['BD']; x.alignment = S['W']
        r += 1
    ws.freeze_panes = None


REFERENCE = [
    ('E-way bill threshold', 'Required for movement of goods of consignment value above Rs.50,000 (Rule 138(1)); '
     'document types: tax invoice, bill of supply, delivery challan, bill of entry, others.',
     'Rule 138, CGST Rules 2017'),
    ('Bill of supply', 'Issued instead of a tax invoice for exempt supplies or by a composition taxpayer; it carries no tax.',
     'Section 31(3)(c) CGST Act; Rule 49'),
    ('Rate cut of 22-09-2025', 'Buses (8702), goods vehicles (8704), bodies (8707), parts (8708) and many other goods '
     'moved from 28% to 18% from 22-09-2025.', 'Notification 9/2025-Integrated Tax (Rate) dt 17-09-2025 (in the rate master)'),
    ('Compensation cess — earlier rates', 'Aerated waters 12%. Motor vehicles: 8702 for 10-13 persons 15%; cars 1%/3% '
     '(small), 15%-22% (mid, large, SUV); motorcycles over 350 cc 3%; goods vehicles and buses for more than 13 persons nil.',
     'Notification 1/2017-Compensation Cess (Rate) as amended — summary at taxguru.in/?p=612328'),
    ('Compensation cess from 22-09-2025', 'Cess substituted with NIL for the listed S.Nos., which include motor vehicles.',
     'Notification 2/2025-Compensation Cess (Rate) dt 17-09-2025 — caclubindia.com notice'),
    ('Used motor vehicles', 'Concessional rate (12%, or 18% for larger vehicles) on the margin: sale value less purchase '
     'value (Rule 32(5)), or less depreciated value where depreciation is claimed under s.32 Income-tax Act; negative margin '
     'ignored. Supply of capital goods on which ITC was taken: s.18(6).',
     'Notification 8/2018-Central Tax (Rate); Rule 32(5); s.18(6) — summary at taxguru.in (sale of old vehicles)'),
]


# ------------------------------------------------------------------ stand-alone workbook
def build_ewb_workbook(ing, E, gstin, name, out_path, mark=True, run_dt=None, context=None):
    """E-way bills only: summary page + 19H..19M. Returns a dict of headline figures."""
    from openpyxl import Workbook
    S = _styles()
    wb = Workbook(); ws = wb.active; ws.title = '00_EWB_SUMMARY'
    build_ewb_sheets(wb, E, mark)
    run_dt = run_dt or datetime.datetime.now()
    ws['A1'] = 'E-WAY BILL ANALYSIS'; ws['A1'].font = S['B']
    rows = [('GSTIN', gstin), ('Legal name (as on the EWB pages)', E.get('legal') or name), ('Period covered', f'{E["first"]} to {E["last"]}'),
            ('Workbook prepared', run_dt.strftime('%d-%m-%Y %H:%M'))]
    r = 3
    for k, v in rows:
        ws.cell(row=r, column=1, value=k).font = S['B']; ws.cell(row=r, column=2, value=v); r += 1
    r += 1
    B = f"'{SHEETS[1]}'"; b0, b1 = E['_b']; a0, a1 = E['_a']; c0, c1 = E['_c']
    C = f"'{SHEETS[2]}'"
    ws.cell(row=r, column=1, value='HEADLINE FIGURES (formulas)').font = S['B']; r += 1
    for k, f in (('Rows read from all pages', f"=COUNTA('{SHEETS[0]}'!E{a0}:E{a1})"),
                 ('Records after removing repeats', f'=COUNTA({B}!B{b0}:B{b1})'),
                 ('Months covered', f'=COUNTA({C}!A{c0}:A{c1})'),
                 ('Months with no e-way bill', f'=COUNTIF({C}!D{c0}:D{c1},0)'),
                 ('Number of e-way bills', f'=SUM({B}!I{b0}:I{b1})'),
                 ('Assessable value (Rs.)', f'=SUM({B}!J{b0}:J{b1})'),
                 ('Tax SGST+CGST+IGST (Rs.)', f'=SUM({B}!O{b0}:O{b1})'),
                 ('Cess (Rs.)', f'=SUM({B}!N{b0}:N{b1})'),
                 ('Value on bills of supply (Rs.)', f'=SUMIFS({B}!J{b0}:J{b1},{B}!D{b0}:D{b1},"Bill Of Supply")'),
                 ('Tax shown on bills of supply (Rs.)', f'=SUMIFS({B}!O{b0}:O{b1},{B}!D{b0}:D{b1},"Bill Of Supply")'),
                 ('Inter-State value (Rs.)', f'=SUMIFS({B}!J{b0}:J{b1},{B}!E{b0}:E{b1},"INTER")')):
        ws.cell(row=r, column=1, value=k); x = ws.cell(row=r, column=2, value=f); x.number_format = '#,##0'
        x.border = S['BD']; r += 1
    r += 1
    ws.cell(row=r, column=1, value='FINDINGS (detail in 19L)').font = S['B']; r += 1
    for i, f in enumerate(E['findings'], 1):
        ws.cell(row=r, column=1, value=f'{i}. [{f["level"]}] {f["head"]}').alignment = S['W']
        ws.cell(row=r, column=2, value=f['periods'] or '—').alignment = S['W']
        r += 1
    if context:
        r += 1
        ws.cell(row=r, column=1, value='BACKGROUND FROM PUBLIC SOURCES (not GST records — verify before use)').font = S['B']; r += 1
        for k, v in context:
            ws.cell(row=r, column=1, value=k).alignment = S['W']; ws.cell(row=r, column=2, value=v).alignment = S['W']; r += 1
    r += 1
    _note(ws, r, 'Direction not stated: the GST-Prime document-wise summary does not say whether this GSTIN is the '
                 'supplier or the recipient on these e-way bills.', S, bold=True, cols=2)
    r += 2
    ws.cell(row=r, column=1, value='SHEETS').font = S['B']; r += 1
    for s, d in SHEET_NOTES.items():
        ws.cell(row=r, column=1, value=s); ws.cell(row=r, column=2, value=d).alignment = S['W']; r += 1
    ws.column_dimensions['A'].width = 66; ws.column_dimensions['B'].width = 105
    for s in wb.worksheets:
        s.page_setup.orientation = 'landscape'; s.page_setup.fitToWidth = 1
        s.sheet_properties.pageSetUpPr.fitToPage = True
    wb.save(out_path)
    return dict(records=len(E['records']), copies=len(E['copies']), months=len(E['months']),
                value=sum(r['Assessable Value'] or 0 for r in E['records']), findings=len(E['findings']))


SHEET_NOTES = OrderedDict([
    ('19H_EWB_COPIES', 'Every row read from every page and block, with where it was read from; which copy is used.'),
    ('19I_EWB_RECORDS', 'One record per month, document type and category; values pulled from the copy used; rate, '
                        'cess %, value per EWB and flags.'),
    ('19J_EWB_MONTHLY', 'Every month in the period (none skipped), SUMIFS over 19I; with GSTR-1/3B when held.'),
    ('19K_EWB_FY', 'Financial-year totals; rate before and after 22-09-2025.'),
    ('19L_EWB_FINDINGS', 'Findings and leads, the law, and the next step.'),
    ('19M_EWB_HSN_RATE', 'HSN declared in GSTR-1 against the e-way bill rate; rate-master rows for the 28%-to-18% '
                         'goods; reference rules used.'),
])
