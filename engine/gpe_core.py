# -*- coding: utf-8 -*-
"""
GST-PRIME PRELIMINARY INVESTIGATION ENGINE  —  core
Ingest, detect, parse and validate GST-Prime Analytica exports.
Every parsed value keeps its provenance: file, sheet, row, column.
"""
import os, re, json, hashlib, datetime, shutil
from collections import OrderedDict
import openpyxl

ENGINE_VERSION = "1.7"

# ----------------------------------------------------------------- helpers
MONTHS = {m: i for i, m in enumerate(
    ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'], 1)}

def md5(path, buf=1 << 20):
    h = hashlib.md5()
    with open(path, 'rb') as f:
        while True:
            b = f.read(buf)
            if not b: break
            h.update(b)
    return h.hexdigest()

def parse_period(v):
    """'Aug-2026' -> ('Aug-2026', 2026, 8). Returns None if not a tax period."""
    if v is None: return None
    s = str(v).strip()
    m = re.fullmatch(r'([A-Za-z]{3})[-/ ](\d{4})', s)
    if not m: return None
    mon = m.group(1).title()
    if mon not in MONTHS: return None
    return (f"{mon}-{m.group(2)}", int(m.group(2)), MONTHS[mon])

def period_key(p):
    t = parse_period(p)
    return (t[1], t[2]) if t else (9999, 99)

def fy_of(p):
    t = parse_period(p)
    if not t: return None
    _, y, m = t
    return f"{y-1}-{str(y)[2:]}" if m <= 3 else f"{y}-{str(y+1)[2:]}"

def fy_long(fy):           # '2021-22' -> '2021-2022'
    a, b = fy.split('-');  return f"{a}-{a[:2]}{b}"

def norm_fy(v):
    """Accept '2021-2022', '2021-22', 2021 -> '2021-22'."""
    if v is None: return None
    s = str(v).strip()
    m = re.fullmatch(r'(\d{4})\s*-\s*(\d{2,4})', s)
    if not m: return None
    a, b = m.group(1), m.group(2)
    return f"{a}-{b[-2:]}"

_NUM_RE = re.compile(r'^-?[\d,]*\.?\d+$')

def to_num(v):
    """Indian-grouped numerics -> float. BLANK STAYS None. '0' stays 0.0."""
    if v is None: return None
    if isinstance(v, bool): return None
    if isinstance(v, (int, float)): return float(v)
    s = str(v).strip()
    if s == '' or s.lower() in ('nan', 'none', '-', 'na', 'n/a'): return None
    s2 = s.replace(',', '').replace('₹', '').replace('Rs.', '').strip()
    if s2.startswith('(') and s2.endswith(')'): s2 = '-' + s2[1:-1]
    if _NUM_RE.fullmatch(s2):
        try: return float(s2)
        except ValueError: return None
    return None

def to_date(v):
    if v is None: return None
    if isinstance(v, datetime.datetime): return v.date()
    if isinstance(v, datetime.date): return v
    s = str(v).strip()
    for f in ('%d-%m-%Y', '%Y-%m-%d', '%d/%m/%Y', '%Y-%m-%d %H:%M:%S', '%d-%b-%Y'):
        try: return datetime.datetime.strptime(s, f).date()
        except ValueError: pass
    return None

def inr(x, dec=0):
    """Indian digit grouping: 10186350 -> '1,01,86,350'. Used in every narrative line."""
    if x is None: return ''
    neg = x < 0; x = abs(float(x))
    whole = int(round(x)) if dec == 0 else int(x)
    frac = '' if dec == 0 else ('%.*f' % (dec, x - whole))[1:]
    s = str(whole)
    if len(s) > 3:
        last3, rest = s[-3:], s[:-3]
        parts = []
        while len(rest) > 2:
            parts.insert(0, rest[-2:]); rest = rest[:-2]
        if rest: parts.insert(0, rest)
        s = ','.join(parts + [last3])
    return ('-' if neg else '') + s + frac


def n0(v):
    """Blank-safe zero for arithmetic. Use ONLY where blank truly means nil."""
    return 0.0 if v is None else float(v)

# ----------------------------------------------------------------- specs
# Each spec: the columns that must be present for a confident match.
# 'key' is the column that identifies a row (tax period / FY / challan).
REPORT_SPECS = OrderedDict([
 ('R3B_TURNOVER', dict(
    must={'Return Period','OutwardTaxable Supplies(Other Than Zero)','Inward Supplies(Reverse Charge)','Total Turnover'},
    key='Return Period', grain='period', fname=r'TURNOVER_OUTWARD_TAX_SUPPLIES',
    title='GSTR-3B turnover and outward taxable supplies')),
 ('R3B_TAX_PAYABLE', dict(
    must={'Return Period','Outward SGST','Outward CGST','Outward IGST','Outward Total GST','Inward Total GST'},
    key='Return Period', grain='period', fname=r'R3B_TAX_PAYABLE_DETAILS',
    title='GSTR-3B tax payable, outward and inward')),
 ('TAX_PAID', dict(
    must={'Return Period','Filing Status','Total TO','Taxable TO','GST OP Total','GST ITC Total'},
    key='Return Period', grain='period', fname=r'TAX_PAID_DETAILS',
    title='GSTR-3B tax paid summary')),
 ('EXCESS_ITC', dict(
    must={'Return Period','SGSTITC Claimed In R3B','SGSTITC Available In R2A','IGSTITC Claimed In R3B','IGSTITC Available In R2A'},
    key='Return Period', grain='period', fname=r'EXCESS_ITC_CLAIMED',
    title='ITC claimed in GSTR-3B against ITC available in GSTR-2A')),
 ('SGST_SETTLEMENT', dict(
    must={'Return Period','SGST Cash Paid','IGST Paid through SGST ITC','SGST Paid through IGST ITC','SGST Settlement'},
    key='Return Period', grain='period', fname=r'SGST_SETTLEMENT_DETAILS',
    title='SGST settlement details')),
 ('HSN', dict(
    must={'Return Period','HSN Code','Taxable Value','Total GST'},
    key='Return Period', grain='period_multi', fname=r'HSN_LIST',
    title='HSN summary (GSTR-1 Table 12)')),
 ('GSTR9', dict(
    must={'Return Period','Date of Filing'},
    key='Return Period', grain='fy', fname=r'GSTR9_FILING_DETAILS',
    title='GSTR-9 annual return filing details')),
 ('R1_MONTHLY', dict(
    must={'Return Period','No Of Sellers','No Of Invoices','Taxable Value','Total GST'},
    key='Return Period', grain='period', fname=r'R1_FILINGS_IN_THE_YEAR',
    title='GSTR-1 filings, month-wise')),
 ('R2A_MONTHLY', dict(
    must={'Return Period','Seller Count','Invoices Count','Taxable Value','Total GST'},
    key='Return Period', grain='period', fname=r'R2A_FILINGS_IN_THE_YEAR',
    title='GSTR-2A filings, month-wise')),
 ('R3B_FILING', dict(
    must={'Return Period','Date Of Filing','Total Turnover','GST Total','GST Input Tax Credit'},
    key='Return Period', grain='period', fname=r'R3B_PAYMENT_SUMMARY_IN_THE_YEAR',
    title='GSTR-3B payment summary with date of filing')),
 ('PAYMENTS', dict(
    must={'Payment Month','Payment Date','CIN','TOTAL'},
    key='CIN', grain='challan', fname=r'PAYMENT_DETAILS',
    title='Cash payment challans')),
 ('R3B_FY_SUMMARY', dict(
    must={'Financial Year','Payment Count','Total Turnover','GST Total','GST Input Tax Credit'},
    key='Financial Year', grain='fy', fname=r'PAYMENTS_GSTR3B',
    title='GSTR-3B payment summary, FY level')),
 ('R2A_FY_SUMMARY', dict(
    must={'Financial Year','GST R2A Count','Seller Count','Invoices Count','Total GST'},
    key='Financial Year', grain='fy', fname=r'GST-R2A_FILING_SUMMARY',
    title='GSTR-2A filing summary, FY level')),
 ('R7_TDS', dict(
    must={'Return Period','No Of Deductors','SGST','CGST','IGST','Total GST'},
    key='Return Period', grain='period', fname=r'R7_FILINGS_IN_THE_YEAR', master=False,
    title='GSTR-7 — tax deducted at source from this taxpayer under section 51, month-wise')),
 ('R7_DETAIL', dict(
    must={'Deductee GSTIN','Trade Name','Taxable Value','SGST','CGST','IGST','Total GST'},
    key='Return Period', grain='period_multi', fname=r'R7_Filing_Details_For_The_Month', master=False,
    period_from_filename=True,
    title='GSTR-7 filing details for one month — party by party, with taxable value')),
 ('R1_FY_SUMMARY', dict(
    must={'Financial Year','Filing Count','No Of Sellers','No Of Invoices','Total GST'},
    key='Financial Year', grain='fy', fname=r'GST-R1_FILING_SUMMARY',
    title='GSTR-1 filing summary, FY level')),
])

DATE_COLS = {'Date of Filing','Date Of Filing','Payment Date'}
TEXT_COLS = {'Deductee GSTIN','Trade Name','Return Period','Financial Year','Filing Status','CIN','Payment Month','HSN Code','Trend',
             'Trend For Other Than Zero Rated','Trend For Zero Rated'}

# ----------------------------------------------------------------- reader
def _find_header(rows):
    """GST-Prime sheets carry a title block then a header row starting 'SNo'."""
    for i, r in enumerate(rows):
        if r and r[0] is not None and str(r[0]).strip().lower() == 'sno':
            return i
    return None

def _declared_records(rows, hdr_i):
    """'Total Records: 12 (All Amounts Are In INR)' -> 12"""
    for r in rows[:hdr_i]:
        for c in r or ():
            if c is None: continue
            m = re.search(r'Total\s+Records\s*:\s*(\d+)', str(c), re.I)
            if m: return int(m.group(1))
    return None

def read_sheet(path):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = [tuple(r) for r in ws.iter_rows(values_only=True)]
    wb.close()
    hdr_i = _find_header(rows)
    if hdr_i is None:
        return None, None, None, rows
    header = [(str(x).strip() if x is not None else '') for x in rows[hdr_i]]
    while header and header[-1] == '': header.pop()
    return hdr_i, header, _declared_records(rows, hdr_i), rows

def _norm_name(filename):
    """Prime saves 'EXCESS ITC CLAIMED FOR SGST-IGST.xlsx' when downloaded, but the same report
    arrives as 'EXCESS_ITC_CLAIMED_FOR_SGST-IGST.xlsx' when re-uploaded. Compare on a form
    where every run of non-alphanumerics is one underscore, so spaces, brackets and dashes
    can never make a correct filename look wrong."""
    return re.sub(r'[^A-Za-z0-9]+', '_', filename).strip('_')

def detect(header, filename):
    """Return (code, score, why). Header signature decides; filename corroborates."""
    filename = _norm_name(filename)
    hs = set(h for h in header if h)
    best, best_score, why = None, 0.0, ''
    for code, sp in REPORT_SPECS.items():
        must = sp['must']
        if not must <= hs:
            continue
        score = len(must) / max(len(hs), 1)
        fn_ok = bool(re.search(sp['fname'].replace('-', '_'), filename, re.I))
        score += 0.5 if fn_ok else 0.0
        if score > best_score:
            best, best_score = code, score
            why = ('header signature matched' + (' and filename agrees' if fn_ok
                   else ' (FILENAME DOES NOT AGREE — header was trusted)'))
    if best is None:
        for code, sp in REPORT_SPECS.items():
            if re.search(sp['fname'].replace('-', '_'), filename, re.I):
                return None, 0.0, (f'filename looks like {code} but the header does not carry '
                                   f'the expected columns — NOT PARSED, needs a look')
    return best, best_score, why

GSTIN_IN_TEXT = re.compile(r'\b(\d{2}[A-Z0-9]{10}\d[A-Z0-9][A-Z0-9])\b')

def subject_gstin(rows, hdr_i):
    """The GSTIN a Prime export names in its title block ('R7 FILING DETAILS OF: 33...'), if any."""
    for r in rows[:hdr_i]:
        for c in r or ():
            if c is None: continue
            m = GSTIN_IN_TEXT.search(str(c).upper())
            if m: return m.group(1)
    return None

def period_from_filename(filename):
    """'R7_Filing_Details_For_The_Month_Mar-2026.xlsx' -> 'Mar-2026'. Used only for a report that
    carries no period inside it; the register says so."""
    m = re.search(r'([A-Za-z]{3})[-_ ](\d{4})(?!.*[A-Za-z]{3}[-_ ]\d{4})', filename)
    if not m: return None
    t = parse_period(f'{m.group(1)}-{m.group(2)}')
    return t[0] if t else None

def parse_rows(header, rows, hdr_i, code, default_period=None):
    """Rows as dicts. BLANK stays None so 'not filed' never reads as 'nil'."""
    sp = REPORT_SPECS[code]
    out = []
    for ri, r in enumerate(rows[hdr_i + 1:], start=hdr_i + 2):
        if not r or r[0] is None: continue
        if str(r[0]).strip().lower() in ('sub total', 'total', 'grand total'): continue
        d, blanks = {}, []
        for ci, col in enumerate(header):
            if not col: continue
            v = r[ci] if ci < len(r) else None
            if v is None or (isinstance(v, str) and v.strip() == ''):
                d[col] = None; blanks.append(col); continue
            if col in DATE_COLS:      d[col] = to_date(v)
            elif col in TEXT_COLS:
                d[col] = (str(int(v)) if isinstance(v, float) and v.is_integer() else str(v).strip())
            else:                     d[col] = to_num(v)
        d['_row'] = ri
        d['_blank_cols'] = blanks
        if default_period and not d.get('Return Period'):
            d['Return Period'] = default_period
        if sp['grain'] in ('period', 'period_multi'):
            t = parse_period(d.get('Return Period'))
            if not t: continue
            d['_period'], d['_fy'] = t[0], fy_of(t[0])
        elif sp['grain'] == 'fy':
            f = norm_fy(d.get('Financial Year') or d.get('Return Period'))
            if f is None: continue
            d['_fy'] = f
        out.append(d)
    return out

# ----------------------------------------------------------------- restatements
_SKIP_FIELD = re.compile(r'^SNo$|Growth|Trend', re.I)   # position / derived columns, not data

def diff_rows(old, new):
    """Fields whose VALUE differs between two versions of the same record.
    SNo is only the row position in that export, and Growth / Trend are derived, so neither counts.
    A difference under Re.1 is rounding, not a restatement."""
    out = []
    for k, b in new.items():
        if k.startswith('_') or _SKIP_FIELD.search(k):
            continue
        a = old.get(k)
        if a is None and b is None:
            continue
        if isinstance(a, (int, float)) and isinstance(b, (int, float)):
            if abs(a - b) < 1.0:
                continue
        elif a == b:
            continue
        out.append((k, a, b))
    return out

def is_dynamic(code, field):
    """GSTR-2A moves on its own: it changes whenever a supplier files late or amends. A change
    in a 2A figure is therefore expected and says nothing about the taxpayer's own returns."""
    if code in ('R2A_MONTHLY', 'R2A_FY_SUMMARY'):
        return True
    if code == 'EXCESS_ITC' and ('R2A' in field or 'Excess' in field):
        return True
    return False

def classify_change(code, field, old, new):
    """(nature, level, reading). Wording is deliberately cautious: the export alone cannot say
    WHY a value moved, only that it did."""
    if old is None and new is not None:
        nature = 'POPULATED'
    elif old is not None and new is None:
        nature = 'BLANKED'
    else:
        nature = 'CHANGED'
    if is_dynamic(code, field):
        return (nature, 'INFO',
                'GSTR-2A is a dynamic statement and moves when suppliers file late or amend. '
                'Expected. Not a restatement by the taxpayer.')
    if nature == 'POPULATED':
        return (nature, 'NOTE',
                'Blank or absent in the earlier export and present now. The return has most likely '
                'been filed, or the data populated, since the earlier export was taken.')
    if nature == 'BLANKED':
        return (nature, 'REVIEW',
                'Held a value in the earlier export and is blank now. Unusual. Check the later '
                'export was not truncated, then check the portal.')
    return (nature, 'REVIEW',
            'A value already held has changed. GSTR-3B and payment challans cannot be revised, so a '
            'change here is unusual. Establish whether the earlier export was taken before the period '
            'was final, whether the taxpayer amended (GSTR-1 amendments are made through later '
            'returns), or whether Prime refreshed its own data. Verify on the portal before relying '
            'on either figure.')

def key_label(code, row):
    """A human-readable name for a record: the tax period, FY, or challan number."""
    g = REPORT_SPECS[code]['grain']
    if g == 'period':       return str(row.get('_period'))
    if g == 'period_multi':
        if row.get('Deductee GSTIN'): return f"{row.get('_period')} / {row.get('Deductee GSTIN')} / Rs.{row.get('Taxable Value')}"
        return f"{row.get('_period')} / HSN {row.get('HSN Code')}"
    if g == 'fy':           return f"FY {row.get('_fy')}"
    if g == 'challan':      return f"{row.get('Payment Month')} / {row.get('CIN')}"
    return str(row.get('SNo'))

# ----------------------------------------------------------------- ingest
def row_key(code, r):
    """The natural identity of a row, so the same record supplied twice is stored once.

    A byte-identical file is rejected before it is ever parsed. This is the second line of
    defence, for the case that actually corrupts figures: a LATER EXPORT of the same report
    carrying the same records plus a few new ones. Without a row key those overlapping records
    would be appended a second time and every total built on them would roughly double."""
    sp = REPORT_SPECS[code]
    g = sp['grain']
    if g == 'period':
        return ('p', r.get('_period'))
    if g == 'fy':
        return ('f', r.get('_fy'))
    if g == 'period_multi':
        # one period may legitimately carry several HSN lines (or several lines for one party),
        # so the whole line is the key
        return ('m', r.get('_period'), str(r.get('HSN Code') or r.get('Deductee GSTIN')),
                r.get('Taxable Value'), r.get('Total GST'))
    if g == 'challan':
        cin = r.get('CIN')
        if cin:
            return ('c', str(cin).strip())
        return ('c', r.get('Payment Month'), str(r.get('Payment Date')), r.get('TOTAL'))
    return ('x', r.get('SNo'))


class Ingest:
    """Holds every file seen, what it was, and whether it was counted."""
    def __init__(self, case_gstin=None):
        self.case_gstin = (case_gstin or '').strip().upper() or None
        self.register = []     # one row per upload
        self.data = {}         # code -> list of row dicts (deduped, merged)
        self._seen_md5 = {}
        self._index = {}       # code -> {row_key: position in self.data[code]}
        self.changes = []      # every value that a later export changed from an earlier one

    def add(self, path, batch=1):
        name = os.path.basename(path)
        h = md5(path)
        rec = dict(file=name, batch=batch, md5=h, size=os.path.getsize(path),
                   code=None, title=None, rows=0, declared=None,
                   status='', note='', detect_note='')
        if h in self._seen_md5:
            rec.update(status='DUPLICATE - NOT COUNTED',
                       note=f'Byte-identical to {self._seen_md5[h]}. Adds nothing; excluded from every count.')
            self.register.append(rec); return rec
        try:
            hdr_i, header, declared, rows = read_sheet(path)
        except Exception as e:
            rec.update(status='UNREADABLE', note=f'Could not open: {e}')
            self.register.append(rec); return rec
        if hdr_i is None:
            rec.update(status='NOT RECOGNISED', note='No header row beginning "SNo" was found.')
            self.register.append(rec); return rec
        code, score, why = detect(header, name)
        rec['declared'] = declared
        rec['detect_note'] = why
        if code is None:
            rec.update(status='NOT RECOGNISED',
                       note=why or 'Header does not match any known GST-Prime report.')
            self.register.append(rec); return rec
        rec['subject_gstin'] = subject_gstin(rows, hdr_i)
        notes = []
        dperiod = None
        if REPORT_SPECS[code].get('period_from_filename'):
            dperiod = period_from_filename(name)
            if dperiod is None:
                rec.update(code=code, title=REPORT_SPECS[code]['title'], status='NOT RECOGNISED',
                           note='This report carries no tax period inside it, and none could be read from the file name '
                                '(expected e.g. ..._Month_Mar-2026.xlsx). Rename the file with its month and upload again.')
                self.register.append(rec); return rec
            notes.append(f'Tax period {dperiod} taken from the FILE NAME — the export carries no period inside it. '
                         f'Keep the Prime file name unchanged.')
        parsed = parse_rows(header, rows, hdr_i, code, default_period=dperiod)
        rec.update(code=code, title=REPORT_SPECS[code]['title'], rows=len(parsed), status='COUNTED')
        if self.case_gstin and rec['subject_gstin'] and rec['subject_gstin'] != self.case_gstin:
            notes.append(f'WRONG TAXPAYER? This export is titled for GSTIN {rec["subject_gstin"]}, not {self.case_gstin}.')
        if declared is not None and declared != len(parsed):
            notes.append(f'Report header declares {declared} records; {len(parsed)} data rows were '
                         f'found. The parsed rows are used. Verify the export on the portal.')
        self._seen_md5[h] = name

        # ---- merge rows by their natural key: a record already held is REPLACED, never added.
        # A key that repeats INSIDE one file (two identical HSN lines, say) is two real rows, so
        # each occurrence is numbered and matches only the same occurrence in a later export.
        bucket = self.data.setdefault(code, [])
        index = self._index.setdefault(code, {})
        added = replaced = rec_changed = vals_changed = 0
        occ = {}
        if REPORT_SPECS[code]['grain'] == 'period_multi' and bucket:
            # A period carries SEVERAL lines (HSN lines, or one line per deduction). A later export of
            # that period is the whole of it: matching line by line would keep a restated line twice
            # (the old value and the new), and keep a line the later export dropped. So every period
            # this file covers is REPLACED as a block, and the difference is recorded line by line.
            blk = self._replace_period_block(code, name, batch, parsed)
            if blk:
                bucket = self.data[code]; index = self._index[code]
                replaced, added, rec_changed, vals_changed = blk
                notes.append(f'Covers {len({x["_period"] for x in parsed})} period(s) already held from an earlier export; '
                             f'those periods were REPLACED AS A WHOLE by this file ({replaced} line(s) unchanged, '
                             f'{rec_changed} line(s) added or dropped). Nothing is counted twice.'
                             + (' See 09A_RESTATEMENTS.' if rec_changed else ''))
                for x in parsed:
                    row = dict(x, _file=name, _batch=batch)
                    k0 = row_key(code, row); n = occ.get(k0, 0); occ[k0] = n + 1
                    index[k0 + (n,)] = len(bucket); bucket.append(row)
                rec['rows_new'] = added; rec['rows_replaced'] = replaced
                rec['records_changed'] = rec_changed; rec['values_changed'] = vals_changed
                rec['note'] = ' '.join(notes)
                self.register.append(rec); return rec
        for x in parsed:
            row = dict(x, _file=name, _batch=batch)
            k0 = row_key(code, row)
            n = occ.get(k0, 0); occ[k0] = n + 1
            k = k0 + (n,)
            if k in index:
                prev = bucket[index[k]]
                ch = diff_rows(prev, row)
                replaced += 1
                # Two files that arrived in the SAME upload batch give no evidence of which export is
                # newer: they share a timestamp, so the order between them is only filename order.
                uncertain = (prev.get('_batch') is not None and prev.get('_batch') == batch)
                row['_supersedes'] = prev.get('_file')
                row['_history'] = dict(supersedes=prev.get('_file'), changes=ch, uncertain=uncertain)
                if ch:
                    rec_changed += 1
                    vals_changed += len(ch)
                    for f, a0, b0 in ch:
                        nature, level, reading = classify_change(code, f, a0, b0)
                        if uncertain:
                            level = 'REVIEW'
                            reading = ('ORDER UNCERTAIN. Both files arrived in the same upload, so nothing shows which '
                                       'export is newer; the value called "later" is only the one that sorted last by '
                                       'filename. Establish which export was taken more recently, delete the other from '
                                       '01_RAW_UPLOADS and run Step 7 again. ') + reading
                        self.changes.append(dict(
                            code=code, title=REPORT_SPECS[code]['title'], record=key_label(code, row),
                            field=f, old=a0, new=b0, old_file=prev.get('_file'), new_file=name,
                            nature=nature, level=level, reading=reading, order_uncertain=uncertain))
                bucket[index[k]] = row
            else:
                index[k] = len(bucket)
                bucket.append(row)
                added += 1
        rec['rows_new'] = added
        rec['rows_replaced'] = replaced
        rec['records_changed'] = rec_changed
        rec['values_changed'] = vals_changed
        if replaced:
            notes.append(f'{replaced} record(s) in this file were already held from an earlier export '
                         f'and have been REPLACED, not added again. {added} record(s) are new.')
            if vals_changed:
                notes.append(f'{vals_changed} VALUE(S) in {rec_changed} of those records DIFFER from the earlier '
                             f'export — see 09A_RESTATEMENTS.')
        elif added == 0 and parsed:
            notes.append('Every record in this file was already held from an earlier export. '
                         'Nothing was added.')
        rec['note'] = ' '.join(notes)
        self.register.append(rec); return rec

    def _replace_period_block(self, code, name, batch, parsed):
        """Drop every line held for the periods this file covers (from other files) and report the
        line-level difference. Returns None when none of its periods is already held."""
        from collections import Counter
        periods = {x['_period'] for x in parsed}
        bucket = self.data[code]
        old = [r for r in bucket if r.get('_period') in periods]
        if not old:
            return None
        sig = lambda r: row_key(code, r)
        c_old, c_new = Counter(sig(r) for r in old), Counter(sig(x) for x in parsed)
        same = sum((c_old & c_new).values())
        gone, came = c_old - c_new, c_new - c_old
        first = {}
        for r in old: first.setdefault(sig(r), r)
        uncertain = any(r.get('_batch') == batch for r in old)
        n_chg = 0
        for k, n in list(gone.items()) + list(came.items()):
            is_new = k in came
            r = (next(x for x in parsed if sig(x) == k)) if is_new else first[k]
            label = key_label(code, dict(r, _period=r['_period']))
            nature = 'ADDED' if is_new else 'DROPPED'
            level = 'REVIEW' if (not is_new or uncertain) else 'NOTE'
            reading = ('A line present in the later export of this period but not in the earlier one — a return filed or '
                       'amended since, or a corrected line (look for a DROPPED line of the same party in the same period).'
                       if is_new else
                       'A line in the earlier export of this period that the later export no longer carries — amended or '
                       'withdrawn, or a corrected line (look for an ADDED line of the same party in the same period). '
                       'The later export is used.')
            if uncertain:
                reading = ('ORDER UNCERTAIN. Both exports of this period arrived in the same upload, so nothing shows which '
                           'is newer. Delete the one that is not current from 01_RAW_UPLOADS and run Step 7 again. ') + reading
            for _ in range(n):
                n_chg += 1
                self.changes.append(dict(code=code, title=REPORT_SPECS[code]['title'], record=label, field='whole line',
                                         old=None if is_new else 'present', new='present' if is_new else None,
                                         old_file=', '.join(sorted({x['_file'] for x in old})), new_file=name,
                                         nature=nature, level=level, reading=reading, order_uncertain=uncertain))
        self.data[code] = [r for r in bucket if r.get('_period') not in periods]
        idx, occ = {}, {}
        for i, r in enumerate(self.data[code]):
            k0 = row_key(code, r); n = occ.get(k0, 0); occ[k0] = n + 1
            idx[k0 + (n,)] = i
        self._index[code] = idx
        n_new = sum(came.values())
        return same, n_new, n_chg, n_chg

    # ---- access -------------------------------------------------------
    def have(self, code): return code in self.data and len(self.data[code]) > 0

    def by_period(self, code):
        """period -> row. Later files win only if the earlier row is emptier."""
        out = {}
        for r in self.data.get(code, []):
            p = r.get('_period')
            if p is None: continue
            if p not in out or len(r.get('_blank_cols', [])) < len(out[p].get('_blank_cols', [])):
                out[p] = r
        return out

    def by_fy(self, code):
        out = {}
        for r in self.data.get(code, []):
            f = r.get('_fy')
            if f: out[f] = r
        return out

    def periods(self):
        """Every tax period seen in any period-grained report, chronological."""
        s = set()
        for code, sp in REPORT_SPECS.items():
            if sp['grain'] in ('period',) and code in self.data and sp.get('master', True):
                s |= {r['_period'] for r in self.data[code] if r.get('_period')}
        return sorted(s, key=period_key)

    def all_periods(self):
        """Every tax period in any report, including those (GSTR-7) that do not open a row in the
        master. Used for limitation, which runs on every year the evidence touches."""
        s = set(self.periods())
        for code, sp in REPORT_SPECS.items():
            if sp['grain'] in ('period', 'period_multi') and code in self.data:
                s |= {r['_period'] for r in self.data[code] if r.get('_period')}
        return sorted(s, key=period_key)

    def counted(self):  return [r for r in self.register if r['status'] == 'COUNTED']
    def duplicates(self): return [r for r in self.register if r['status'].startswith('DUPLICATE')]
    def problems(self): return [r for r in self.register
                                if r['status'] in ('NOT RECOGNISED', 'UNREADABLE')]
