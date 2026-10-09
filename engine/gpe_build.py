# -*- coding: utf-8 -*-
"""
GST-PRIME PRELIMINARY INVESTIGATION ENGINE  —  workbook builder
Source sheets are written verbatim first. Every computed sheet then reads them
with live Excel formulas, so any value can be traced back to a source cell.
"""
import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter
from gpe_core import REPORT_SPECS, period_key, fy_of, n0, inr
from gpe_checks import due_date

BUILD_VERSION = "1.6"
BOLD = Font(bold=True)
THIN = Side(style='thin')
BD = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
MISSING_FILL = PatternFill('solid', fgColor='FFD9D9')   # light red — data not available
SUPERSEDED_FILL = PatternFill('solid', fgColor='FFF2CC')  # light amber — record replaced by a later export
CHANGED_FILL = PatternFill('solid', fgColor='F8CBAD')     # orange — a value that differs from the earlier export
NUM = '#,##0'

def _q(s):  return f"'{s}'"

def _frank(f):
    """Findings order: revenue first; among equal amounts, by risk."""
    r = f.get('r', '')
    k = 0 if r.startswith('VERY HIGH') else 1 if r.startswith('HIGH') else 2 if r.startswith('MEDIUM') else 3
    return (-f['a'], k)

class Builder:
    def __init__(self, ing, R, gstin, name, mark_missing_light_red=True,
                 extra_profile=None):
        self.ing, self.R = ing, R
        self.gstin, self.name = gstin, name
        self.mark = mark_missing_light_red
        self.profile = extra_profile or {}
        self.wb = Workbook(); self.wb.remove(self.wb.active)
        self.src_rows = {}          # code -> last data row
        self.periods = [m['period'] for m in R['master']]

    # ---------------------------------------------------------- utilities
    def sheet(self, title):  return self.wb.create_sheet(title)

    def head(self, ws, r, hdr, widths=None):
        for j, h in enumerate(hdr, 1):
            c = ws.cell(row=r, column=j, value=h); c.font = BOLD; c.border = BD
            c.alignment = Alignment(wrap_text=True, vertical='top', horizontal='center')
        if widths:
            for i, w in enumerate(widths, 1): ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = ws.cell(row=r + 1, column=1)

    def box(self, ws, r0, r1, c0, c1, numfrom=None):
        for r in range(r0, r1 + 1):
            for c in range(c0, c1 + 1):
                x = ws.cell(row=r, column=c); x.border = BD
                if numfrom and c >= numfrom and not isinstance(x.value, (datetime.date, datetime.datetime)):
                    x.number_format = NUM

    @staticmethod
    def _safe(v):
        """Stop a descriptive string that starts with '=' being read as a formula."""
        return ('Formula: ' + v[1:].strip()) if isinstance(v, str) and v.startswith('=') and ' ' in v[:14] else v

    def note(self, ws, r, text, bold=False, col=1):
        c = ws.cell(row=r, column=col, value=text)
        c.alignment = Alignment(wrap_text=True, vertical='top')
        if bold: c.font = BOLD
        return r + 1

    def na(self, ws, r, c, text='NOT YET AVAILABLE'):
        x = ws.cell(row=r, column=c, value=text)
        if self.mark: x.fill = MISSING_FILL
        x.alignment = Alignment(horizontal='center')
        return x

    def unavailable(self, ws, res, what):
        ws['A1'] = what; ws['A1'].font = BOLD
        r = self.note(ws, 3, 'THIS TEST COULD NOT BE RUN.', bold=True)
        r = self.note(ws, r, 'The following GST-Prime report(s) have not been supplied: '
                             + ', '.join(REPORT_SPECS[m]['title'] + f'  [{m}]'
                                         for m in res.get('missing', [])))
        r = self.note(ws, r, 'Upload the report and run the notebook again. Nothing has been '
                             'estimated or assumed in its place.')
        if self.mark: ws.cell(row=3, column=1).fill = MISSING_FILL
        ws.column_dimensions['A'].width = 130

    # ---------------------------------------------------- 09A restatements
    def build_restatements(self):
        ch = self.ing.changes
        counted = self.ing.counted()
        n_repl = sum(r.get('rows_replaced', 0) for r in counted)
        n_recchg = sum(r.get('records_changed', 0) for r in counted)
        ws = self.sheet('09A_RESTATEMENTS')
        ws['A1'] = 'WHAT A LATER EXPORT CHANGED FROM AN EARLIER ONE'; ws['A1'].font = BOLD
        ws['A2'] = ('When a later Prime export carries a record already held, the later record replaces the earlier one. '
                    'Where the two differ, the difference is listed here with both values, so a restatement is never '
                    'absorbed silently. The later value is the one used everywhere in this workbook.')
        r = 4
        ws.cell(row=r, column=1, value='Records REPLACED by a later export').font = BOLD
        ws.cell(row=r, column=4, value=n_repl); r += 1
        ws.cell(row=r, column=1, value='   of which the values were IDENTICAL').font = BOLD
        ws.cell(row=r, column=4, value=n_repl - n_recchg); r += 1
        ws.cell(row=r, column=1, value='   of which one or more values DIFFER').font = BOLD
        ws.cell(row=r, column=4, value=n_recchg); r += 2
        if not ch:
            self.note(ws, r, 'No value in any later export differs from the earlier export it replaced.' if n_repl
                             else 'No record was supplied more than once. Nothing was replaced.', bold=True)
            ws.column_dimensions['A'].width = 60
            return
        order = {'REVIEW': 0, 'NOTE': 1, 'INFO': 2}
        rows = sorted(ch, key=lambda c: (order[c['level']], c['code'], str(c['record']), c['field']))
        ws.cell(row=r, column=1, value='SUMMARY').font = BOLD; r += 1
        H = ['Level','What it means','Values','Records']
        for j, h in enumerate(H, 1):
            x = ws.cell(row=r, column=j, value=h); x.font = BOLD; x.border = BD
        meaning = {'REVIEW': 'A value the taxpayer reported, or a challan, differs from the earlier export. Verify on the portal before relying on either.',
                   'NOTE': 'Blank in the earlier export and populated now — most likely a return filed since.',
                   'INFO': 'GSTR-2A moved, as it does whenever a supplier files late or amends. Expected.'}
        r += 1; sum0 = r
        det0 = sum0 + 3 + 3          # detail header row is computed below; formulas use whole-column ranges instead
        for lv in ('REVIEW', 'NOTE', 'INFO'):
            ws.cell(row=r, column=1, value=lv)
            ws.cell(row=r, column=2, value=meaning[lv]).alignment = Alignment(wrap_text=True, vertical='top')
            ws.cell(row=r, column=3, value=sum(1 for c in ch if c['level'] == lv))
            ws.cell(row=r, column=4, value=len({(c['code'], c['record']) for c in ch if c['level'] == lv}))
            for j in range(1, 5): ws.cell(row=r, column=j).border = BD
            if lv == 'REVIEW' and self.mark and any(c['level'] == 'REVIEW' for c in ch):
                ws.cell(row=r, column=1).fill = MISSING_FILL
            r += 1
        r += 1
        ws.cell(row=r, column=1, value='EVERY DIFFERENCE').font = BOLD; r += 1
        H = ['Level','Report','Record','Field','Value in the EARLIER export','Value in the LATER export',
             'Difference (later minus earlier)','Nature','Which is newer?','File labelled earlier',
             'File labelled later','Reading']
        self.head(ws, r, H, [10,40,26,26,22,22,22,12,26,42,42,88])
        hdr = r; r += 1
        for c in rows:
            vals = [c['level'], c['title'], c['record'], c['field'], c['old'], c['new'], None, c['nature'],
                    ('UNCERTAIN — same upload' if c.get('order_uncertain') else 'Later upload batch'),
                    c['old_file'], c['new_file'], c['reading']]
            for j, v in enumerate(vals, 1):
                x = ws.cell(row=r, column=j, value=v); x.border = BD
                x.alignment = Alignment(wrap_text=True, vertical='top')
            ws.cell(row=r, column=7, value=f'=IF(AND(ISNUMBER(E{r}),ISNUMBER(F{r})),F{r}-E{r},"")')
            for j in (5, 6, 7): ws.cell(row=r, column=j).number_format = NUM
            if c['level'] == 'REVIEW' and self.mark: ws.cell(row=r, column=1).fill = MISSING_FILL
            if c.get('order_uncertain') and self.mark: ws.cell(row=r, column=9).fill = MISSING_FILL
            r += 1
        r += 1
        n_unc = sum(1 for c in ch if c.get('order_uncertain'))
        if n_unc:
            r = self.note(ws, r, f'ORDER UNCERTAIN for {n_unc} value(s). The two files arrived in the same upload, so the '
                                 f'upload gives no evidence of which is newer, and Prime exports carry no date of their own. '
                                 f'The value called "later" is only the one that sorted last by filename. Establish which '
                                 f'export was taken more recently, delete the other from 01_RAW_UPLOADS and run again.', bold=True)
        r = self.note(ws, r, 'WHY THIS MATTERS. A later export that changes a value already held may mean the taxpayer amended '
                             'something, or that Prime refreshed its own data, or that the earlier export was taken before the '
                             'period was final. The export cannot say which. GSTR-3B and payment challans cannot be revised, so a '
                             'REVIEW difference in either is unusual and should be checked on the portal before either figure is '
                             'quoted in a notice.', bold=True)
        self.note(ws, r, 'The later value is the one used in every computed sheet. If the earlier value is the correct one, '
                         'remove the later file from 01_RAW_UPLOADS and run again.')

    # ------------------------------------------------------ source sheets
    def build_sources(self):
        order = [c for c in REPORT_SPECS if self.ing.have(c)]
        for code in order:
            sp = REPORT_SPECS[code]
            rows = self.ing.data[code]
            if sp['grain'] in ('period', 'period_multi'):
                rows = sorted(rows, key=lambda r: (period_key(r['_period']), r.get('SNo') or 0))
            elif sp['grain'] == 'fy':
                rows = sorted(rows, key=lambda r: r.get('_fy') or '')
            cols = [k for k in rows[0].keys() if not k.startswith('_')]
            for r in rows:
                for k in r:
                    if not k.startswith('_') and k not in cols: cols.append(k)
            ws = self.sheet(f'SRC_{code}'[:31])
            ws.cell(row=1, column=1, value='Source file')
            for j, c in enumerate(cols, 2): ws.cell(row=1, column=j, value=c)
            HIST = len(cols) + 2                       # last column: record history
            ws.cell(row=1, column=HIST, value='Record history')
            col_of = {c: j for j, c in enumerate(cols, 2)}
            n_super = n_chg = 0
            for i, r in enumerate(rows, 2):
                ws.cell(row=i, column=1, value=r.get('_file'))
                for j, c in enumerate(cols, 2):
                    v = r.get(c)
                    cell = ws.cell(row=i, column=j, value=v)
                    if isinstance(v, (datetime.date, datetime.datetime)):
                        cell.number_format = 'DD-MM-YYYY'
                    elif v is None and self.mark and c not in ('CESS',):
                        cell.fill = MISSING_FILL
                h = r.get('_history')
                if h:
                    n_super += 1
                    ch = h.get('changes') or []
                    txt = ((f'ORDER UNCERTAIN (same upload). ' if h.get('uncertain') and ch else '')
                           + f'REPLACES the record held from {h.get("supersedes")}. '
                           + (f'{len(ch)} value(s) differ: ' + '; '.join(f'{f}: {a0} -> {b0}' for f, a0, b0 in ch)
                              if ch else 'No value differs.'))
                    hc = ws.cell(row=i, column=HIST, value=txt)
                    hc.alignment = Alignment(wrap_text=True, vertical='top')
                    if ch:
                        n_chg += 1
                        if self.mark:
                            for j in range(1, HIST + 1):
                                if ws.cell(row=i, column=j).fill != MISSING_FILL:
                                    ws.cell(row=i, column=j).fill = SUPERSEDED_FILL
                    for f, a0, b0 in ch:
                        j = col_of.get(f)
                        if j:
                            cc = ws.cell(row=i, column=j)
                            if self.mark: cc.fill = CHANGED_FILL
                            cc.comment = Comment(f'Earlier export ({h.get("supersedes")}) held: {a0}\n'
                                                 f'This value comes from {r.get("_file")}.', 'Engine')
            self.head(ws, 1, ['Source file'] + cols + ['Record history'],
                      [34] + [16] * len(cols) + [58])
            self.box(ws, 2, 1 + len(rows), 1, HIST, numfrom=3)
            self.src_rows[code] = 1 + len(rows)
            self.col_index = None
            r = 3 + len(rows)
            self.note(ws, r, f'{sp["title"]}. Reproduced exactly as received — no value altered, '
                             f'no blank filled in. A cell shaded light red was BLANK in the source, '
                             f'which is not the same as nil.', bold=True)
            if n_super:
                self.note(ws, r + 1, f'{n_super} row(s) come from a LATER export and REPLACED a record held from an earlier one '
                                     f'(see the Record history column). {n_chg} of them differ from the record they replaced and '
                                     f'are shaded amber; an orange cell is the value that differs — hover over it for the earlier '
                                     f'value. The other {n_super - n_chg} were replaced by identical values and are not shaded. '
                                     f'Every difference is listed in 09A_RESTATEMENTS.', bold=True)
        return order

    def src_col(self, code, name):
        """Column letter of a named column on the SRC sheet for `code`, or None if that
        report was never supplied and so has no source sheet."""
        title = f'SRC_{code}'[:31]
        if title not in self.wb.sheetnames:
            return None
        ws = self.wb[title]
        for c in range(1, ws.max_column + 1):
            if ws.cell(row=1, column=c).value == name: return get_column_letter(c)
        return None

    def lookup(self, code, colname, keycol, key_ref):
        """Live INDEX/MATCH into a source sheet; 0 when the report was not supplied at all,
        so a missing report leaves an honest zero rather than breaking the workbook."""
        if code not in self.src_rows:
            return 0
        s = _q(f'SRC_{code}'[:31]); last = self.src_rows[code]
        col = self.src_col(code, colname); kc = self.src_col(code, keycol)
        if col is None or kc is None: return 0
        return (f'=IFERROR(INDEX({s}!${col}$2:${col}${last},'
                f'MATCH({key_ref},{s}!${kc}$2:${kc}${last},0)),0)')

    # ------------------------------------------------------------ master
    def build_master(self):
        ws = self.sheet('10_MONTHLY_MASTER')
        if not self.R['master']:
            self.master_last = 3
            return self.unavailable(ws, dict(missing=['TAX_PAID', 'R3B_TAX_PAYABLE', 'R3B_TURNOVER']),
                                    'MONTHLY MASTER — no GSTR-3B report supplied, so there is no tax period to build it on')
        ws['A1'] = ('MONTHLY MASTER — one row per tax period, chronological. Every figure below '
                    'row 3 is a live formula reading the SRC_ sheets.')
        ws['A1'].font = BOLD
        H = ['Return Period','FY','Filing status','Total turnover (3B)','Taxable turnover (3B)',
             'Outward SGST','Outward CGST','Outward IGST','Outward total tax (3.1(a))',
             'GST OP Total (incl. RCM)','Implied RCM / other liability','RCM inward supply value',
             'ITC availed','SGST ITC claimed 3B','SGST available 2A','SGST excess',
             'IGST ITC claimed 3B','IGST available 2A','IGST excess','SGST cash paid',
             'SGST settlement','Gross SGST collected','NIL turnover','ITC availed in NIL period',
             'GSTR-2A data availability']
        self.head(ws, 3, H, [13,9,13]+[15]*19+[11,16,20])
        P = lambda r: f'$A{r}'
        for i, m in enumerate(self.R['master'], 4):
            ws.cell(row=i, column=1, value=m['period'])
            ws.cell(row=i, column=2, value=m['fy'])
            ws.cell(row=i, column=3, value=self.lookup('TAX_PAID','Filing Status','Return Period',P(i)))
            ws.cell(row=i, column=4, value=self.lookup('TAX_PAID','Total TO','Return Period',P(i)))
            ws.cell(row=i, column=5, value=self.lookup('TAX_PAID','Taxable TO','Return Period',P(i)))
            for col, nm in zip(range(6,10), ['Outward SGST','Outward CGST','Outward IGST','Outward Total GST']):
                ws.cell(row=i, column=col, value=self.lookup('R3B_TAX_PAYABLE', nm, 'Return Period', P(i)))
            ws.cell(row=i, column=10, value=self.lookup('TAX_PAID','GST OP Total','Return Period',P(i)))
            ws.cell(row=i, column=11, value=f'=J{i}-I{i}')
            ws.cell(row=i, column=12, value=self.lookup('R3B_TURNOVER','Inward Supplies(Reverse Charge)','Return Period',P(i)))
            ws.cell(row=i, column=13, value=self.lookup('TAX_PAID','GST ITC Total','Return Period',P(i)))
            for col, nm in zip(range(14,20), ['SGSTITC Claimed In R3B','SGSTITC Available In R2A','SGSTExcess Claimed',
                                              'IGSTITC Claimed In R3B','IGSTITC Available In R2A','IGSTExcess Claimed']):
                ws.cell(row=i, column=col, value=self.lookup('EXCESS_ITC', nm, 'Return Period', P(i)))
            for col, nm in zip(range(20,23), ['SGST Cash Paid','SGST Settlement','Gross SGST Collected']):
                ws.cell(row=i, column=col, value=self.lookup('TAX_PAID', nm, 'Return Period', P(i)))
            ws.cell(row=i, column=23, value=f'=IF(D{i}=0,"NIL","")')
            ws.cell(row=i, column=24, value=f'=IF(W{i}="NIL",M{i},0)')
            c = ws.cell(row=i, column=25, value='2A NOT AVAILABLE' if m['r2a_not_available'] else '')
            if m['r2a_not_available'] and self.mark: c.fill = MISSING_FILL
        last = 3 + len(self.R['master'])
        ws.cell(row=last+1, column=1, value='TOTAL').font = BOLD
        for c in list(range(4,23)) + [24]:
            L = get_column_letter(c)
            ws.cell(row=last+1, column=c, value=f'=SUM({L}4:{L}{last})').font = BOLD
        self.box(ws, 4, last+1, 1, 25, numfrom=4)
        for r in range(4, last+2):
            for c in (3, 23, 25): ws.cell(row=r, column=c).number_format = 'General'
        self.master_last = last
        return last

    def m(self, col, row):  return f"'10_MONTHLY_MASTER'!{col}{row}"

    # --------------------------------------------------------- FY summary
    def build_fy(self):
        ws = self.sheet('11_FY_SUMMARY')
        if not self.R['master']:
            self.fy_last = 4
            return self.unavailable(ws, dict(missing=['TAX_PAID', 'R3B_TAX_PAYABLE']),
                                    'FINANCIAL-YEAR SUMMARY OF GSTR-3B')
        ws['A1'] = 'FINANCIAL-YEAR SUMMARY — SUMIFS over 10_MONTHLY_MASTER'; ws['A1'].font = BOLD
        H = ['FY','Tax periods','NIL periods','Turnover (3B)','Output tax (incl. RCM)','ITC availed',
             'ITC as % of output tax','SGST cash paid','Net SGST to exchequer','SGST excess over 2A',
             'IGST excess over 2A','ITC availed in NIL periods']
        self.head(ws, 3, H, [10,11,11,18,18,16,13,15,17,16,16,18])
        M = _q('10_MONTHLY_MASTER'); L = self.master_last
        fys = sorted(self.R['fy'], key=lambda x: int(x.split('-')[0]))
        for i, f in enumerate(fys, 4):
            ws.cell(row=i, column=1, value=f)
            ws.cell(row=i, column=2, value=f'=COUNTIF({M}!$B$4:$B${L},$A{i})')
            ws.cell(row=i, column=3, value=f'=COUNTIFS({M}!$B$4:$B${L},$A{i},{M}!$W$4:$W${L},"NIL")')
            for col, src in zip([4,5,6,8,9,10,11,12], ['D','J','M','T','V','P','S','X']):
                ws.cell(row=i, column=col,
                        value=f'=SUMIF({M}!$B$4:$B${L},$A{i},{M}!${src}$4:${src}${L})')
            ws.cell(row=i, column=7, value=f'=IFERROR(F{i}/E{i},"")')
        r = 4 + len(fys)
        ws.cell(row=r, column=1, value='TOTAL').font = BOLD
        for c in [2,3,4,5,6,8,9,10,11,12]:
            LL = get_column_letter(c); ws.cell(row=r, column=c, value=f'=SUM({LL}4:{LL}{r-1})').font = BOLD
        ws.cell(row=r, column=7, value=f'=IFERROR(F{r}/E{r},"")').font = BOLD
        self.box(ws, 4, r, 1, 12, numfrom=2)
        for i in range(4, r+1): ws.cell(row=i, column=7).number_format = '0.0%'
        self.fy_last = r
        self.note(ws, r+2, 'Net SGST to exchequer is Gross SGST Collected as computed by GST-Prime '
                           '(cash plus settlement). A negative figure means the State bore a net outflow that year.')
        return r

    # --------------------------------------------- 13 excess ITC (demand)
    def build_excess(self):
        res = self.R['excess_itc']; ws = self.sheet('13_EXCESS_ITC')
        if not res.get('available'): return self.unavailable(ws, res, 'ITC AVAILED AGAINST ITC AVAILABLE IN GSTR-2A')
        ws['A1'] = 'ITC AVAILED IN GSTR-3B AGAINST ITC AVAILABLE IN GSTR-2A — head by head, month by month'
        ws['A1'].font = BOLD
        ws['A2'] = 'THIS IS THE DEMAND FIGURE. Heads are compared separately because credit in one head is not set off against a shortfall in another.'
        H = ['Return Period','FY','SGST claimed 3B','SGST available 2A','SGST excess',
             'CGST excess (mirror of SGST)','IGST claimed 3B','IGST available 2A','IGST excess',
             'Total excess (S+C+I)','Positive excess only','2A availability']
        self.head(ws, 4, H, [13,9]+[17]*9+[18])
        L = self.master_last
        for i, m in enumerate(self.R['master'], 5):
            mr = i - 1
            ws.cell(row=i, column=1, value=f'={self.m("A", mr)}')
            ws.cell(row=i, column=2, value=f'={self.m("B", mr)}')
            for col, src in zip([3,4,5,7,8,9], ['N','O','P','Q','R','S']):
                ws.cell(row=i, column=col, value=f'={self.m(src, mr)}')
            ws.cell(row=i, column=6, value=f'=E{i}')
            ws.cell(row=i, column=10, value=f'=E{i}+F{i}+I{i}')
            ws.cell(row=i, column=11, value=f'=MAX(0,E{i})+MAX(0,F{i})+MAX(0,I{i})')
            c = ws.cell(row=i, column=12, value='2A NOT AVAILABLE' if m['r2a_not_available'] else '')
            if m['r2a_not_available'] and self.mark: c.fill = MISSING_FILL
        r = 4 + len(self.R['master']) + 1
        ws.cell(row=r, column=1, value='TOTAL').font = BOLD
        for c in range(3, 12):
            LL = get_column_letter(c); ws.cell(row=r, column=c, value=f'=SUM({LL}5:{LL}{r-1})').font = BOLD
        self.box(ws, 5, r, 1, 12, numfrom=3)
        for i in range(5, r+1): ws.cell(row=i, column=12).number_format = 'General'
        self.excess_net_cell = f"'13_EXCESS_ITC'!J{r}"
        self.excess_gross_cell = f"'13_EXCESS_ITC'!K{r}"
        s = r + 2
        ws.cell(row=s, column=1, value='FINANCIAL-YEAR ROLL-UP').font = BOLD
        H2 = ['FY','SGST excess (net)','CGST excess (net)','IGST excess (net)','Total excess (net)','Total excess (gross)']
        self.head(ws, s+1, H2)
        fys = sorted(self.R['fy'], key=lambda x: int(x.split('-')[0]))
        for i, f in enumerate(fys, s+2):
            ws.cell(row=i, column=1, value=f)
            ws.cell(row=i, column=2, value=f'=SUMIF($B$5:$B${r-1},$A{i},$E$5:$E${r-1})')
            ws.cell(row=i, column=3, value=f'=B{i}')
            ws.cell(row=i, column=4, value=f'=SUMIF($B$5:$B${r-1},$A{i},$I$5:$I${r-1})')
            ws.cell(row=i, column=5, value=f'=B{i}+C{i}+D{i}')
            ws.cell(row=i, column=6, value=f'=SUMIF($B$5:$B${r-1},$A{i},$K$5:$K${r-1})')
        r2 = s + 2 + len(fys)
        ws.cell(row=r2, column=1, value='TOTAL').font = BOLD
        for c in range(2, 7):
            LL = get_column_letter(c); ws.cell(row=r2, column=c, value=f'=SUM({LL}{s+2}:{LL}{r2-1})').font = BOLD
        self.box(ws, s+2, r2, 1, 6, numfrom=2)
        n = r2 + 2
        n = self.note(ws, n, 'NET is the sum of all periods, positive and negative. GROSS counts only the periods of over-claim. '
                             'A short claim in one period followed by a larger claim in the next is ordinarily only timing, '
                             'so NET is the defensible exposure and GROSS the outer limit.', bold=True)
        n = self.note(ws, n, 'CGST is not reported by the source statement. It is taken as equal to SGST, which holds for '
                             'intra-State supply. Confirm from the electronic credit ledger.')
        self.note(ws, n, 'Legal handles: section 16(2)(aa) and Rule 36(4); Rule 88D / FORM DRC-01C; Rule 86A where the credit is found ineligible.')

    # ----------------------------------------------- 17 GSTR-1 vs GSTR-3B
    def build_r1(self):
        res = self.R['r1_vs_r3b']; ws = self.sheet('17_R1_vs_R3B')
        if not res.get('available'): return self.unavailable(ws, res, 'GSTR-1 COMPARED WITH GSTR-3B')
        ws['A1'] = 'GSTR-1 COMPARED WITH GSTR-3B — every tax period'; ws['A1'].font = BOLD
        H = ['Return Period','FY','GSTR-1 filing status','Recipients','Documents','GSTR-1 taxable value',
             'GSTR-3B turnover','Difference in value','GSTR-1 tax','GSTR-3B outward tax (3.1(a))',
             'Difference in tax','Short declared in 3B','Remark']
        self.head(ws, 3, H, [13,9,27,11,11,18,18,17,17,19,17,18,72])
        for i, row in enumerate(res['rows'], 4):
            mr = i
            ws.cell(row=i, column=1, value=row['period']); ws.cell(row=i, column=2, value=row['fy'])
            c = ws.cell(row=i, column=3, value=row['status'])
            if row['status'].startswith('NOT FILED') and self.mark: c.fill = MISSING_FILL
            for col, k in zip([4,5], ['sellers','invoices']):
                ws.cell(row=i, column=col, value=row[k] if row[k] is not None else '')
            ws.cell(row=i, column=6, value=row['r1_value'] if row['r1_value'] is not None else '')
            ws.cell(row=i, column=7, value=f'={self.m("D", mr)}')
            ws.cell(row=i, column=8, value=f'=IF(F{i}="","",F{i}-G{i})')
            ws.cell(row=i, column=9, value=row['r1_tax'] if row['r1_tax'] is not None else '')
            ws.cell(row=i, column=10, value=f'={self.m("I", mr)}')
            ws.cell(row=i, column=11, value=f'=IF(I{i}="","",I{i}-J{i})')
            ws.cell(row=i, column=12, value=f'=IF(K{i}="",0,MAX(0,K{i}))')
            rm = ''
            if row['status'] == 'NO RETURN OF ANY KIND':
                rm = 'No GSTR-3B either. Period of registration or a period with no return of any kind.'
            elif row['status'].startswith('NOT FILED'):
                rm = (f'GSTR-1 NOT ON RECORD. GSTR-3B nonetheless declares turnover of '
                      f'Rs.{inr(row["r3b_value"])} and outward tax of Rs.{inr(row["r3b_tax"])}. '
                      f'No invoice-level data exists on the portal for this period.'
                      if row['r3b_value'] or row['r3b_tax'] else
                      'GSTR-1 NOT ON RECORD. GSTR-3B for the period is nil.')
            elif row['diff'] is not None and row['diff'] > 1:
                rm = (f'Tax of Rs.{inr(row["diff"])} declared in GSTR-1 and not carried into GSTR-3B. '
                      + ('GSTR-3B shows NIL turnover against a GSTR-1 supply.' if row['r3b_value'] == 0
                         else 'Turnover carried into GSTR-3B but the tax short-declared.'))
            elif row['r1_tax'] is not None and row['r1_tax'] < -1:
                rm = ('GSTR-1 carries a NEGATIVE tax value — credit notes. '
                      + ('TAXABLE VALUE IS NOT NEGATIVE: the return is internally inconsistent on its face.'
                         if (row['r1_value'] or 0) >= 0 else 'Not carried into GSTR-3B.'))
            elif row['diff'] is not None and row['diff'] < -1:
                rm = 'GSTR-3B tax exceeds GSTR-1 — usually the reverse-charge liability of Table 3.1(d). Reconcile.'
            c = ws.cell(row=i, column=13, value=rm); c.alignment = Alignment(wrap_text=True, vertical='top')
        r = 3 + len(res['rows']) + 1
        ws.cell(row=r, column=1, value='TOTAL').font = BOLD
        for c in [6,7,8,9,10,11,12]:
            LL = get_column_letter(c); ws.cell(row=r, column=c, value=f'=SUM({LL}4:{LL}{r-1})').font = BOLD
        self.box(ws, 4, r, 1, 13, numfrom=4)
        for i in range(4, r+1): ws.cell(row=i, column=3).number_format = 'General'
        self.short_cell = f"'17_R1_vs_R3B'!L{r}"
        n = r + 1
        ws.cell(row=n, column=1, value='Periods with NO GSTR-1 on record').font = BOLD
        ws.cell(row=n, column=3, value=f'=COUNTIF(C4:C{r-1},"NOT FILED*")').font = BOLD
        ws.cell(row=n+1, column=1, value='Turnover declared in GSTR-3B for those periods').font = BOLD
        ws.cell(row=n+1, column=7, value=f'=SUMIF(C4:C{r-1},"NOT FILED*",G4:G{r-1})').font = BOLD
        ws.cell(row=n+2, column=1, value='Outward tax declared in GSTR-3B for those periods').font = BOLD
        ws.cell(row=n+2, column=10, value=f'=SUMIF(C4:C{r-1},"NOT FILED*",J4:J{r-1})').font = BOLD
        for rr in (n, n+1, n+2):
            for cc in (3,7,10): ws.cell(row=rr, column=cc).number_format = NUM
        self.nf_count_cell = f"'17_R1_vs_R3B'!C{n}"
        self.nf_to_cell = f"'17_R1_vs_R3B'!G{n+1}"
        self.nf_tax_cell = f"'17_R1_vs_R3B'!J{n+2}"
        self.note(ws, n+4, 'Compared against GSTR-3B Table 3.1(a) and not against GST OP Total, because OP Total '
                           'includes the reverse-charge liability of Table 3.1(d), which has no counterpart in GSTR-1.')
        self.note(ws, n+5, 'A positive difference is self-assessed tax within the Explanation to section 75(12), recoverable '
                           'under section 79 without a show cause notice. Intimation route: Rule 88C / FORM DRC-01B.')

    # ------------------------------------- 19B credit notes without supply
    def build_credit_notes(self):
        res = self.R['credit_notes']; ws = self.sheet('19B_CREDIT_NOTES')
        if not res.get('available'): return self.unavailable(ws, res, 'CREDIT NOTES TESTED AGAINST THE ORIGINAL SUPPLY')
        ws['A1'] = 'CREDIT NOTES TESTED AGAINST EVERY POSITIVE SUPPLY EVER DECLARED UNDER THE SAME HSN'
        ws['A1'].font = BOLD
        ws['A2'] = ('Section 34(1) requires a credit note to be referable to an identified original tax invoice. '
                    'Section 34(2) allows the reduction in liability only up to 30 November following the end of the '
                    'financial year of that supply.')
        H = ['Return Period','FY','Negative turnover declared','Negative tax declared','HSN code(s)',
             'Positive supply ever declared under the same HSN','As % of ALL positive turnover ever declared','Assessment']
        self.head(ws, 4, H, [13,9,22,20,14,46,22,80])
        r = 5
        if not res['events']:
            self.note(ws, r, 'No tax period declares a negative outward supply. Nothing to test.', bold=True); r += 2
        for e in res['events']:
            ws.cell(row=r, column=1, value=e['period']); ws.cell(row=r, column=2, value=e['fy'])
            ws.cell(row=r, column=3, value=e['turnover']); ws.cell(row=r, column=4, value=e['tax'])
            ws.cell(row=r, column=5, value=', '.join(e['hsn']) if e['hsn'] else 'not reported')
            found = []
            for code, lst in e['prior_positive_same_hsn'].items():
                found.append(f'{code}: NONE — no positive supply under this HSN in any period'
                             if not lst else
                             f'{code}: ' + '; '.join(f'{p} Rs.{inr(v)}' for p, v, _ in lst))
            cell = ws.cell(row=r, column=6, value='\n'.join(found) if found else 'HSN not reported for this period')
            cell.alignment = Alignment(wrap_text=True, vertical='top')
            if any('NONE' in x for x in found) and self.mark: cell.fill = MISSING_FILL
            ws.cell(row=r, column=7, value=e['pct_of_all_positive_turnover'] or 0).number_format = '0.0%'
            none_found = any('NONE' in x for x in found)
            a = ('NO ORIGINAL SUPPLY FOUND. The taxpayer has never declared a positive outward supply under this HSN '
                 'in any tax period on record. A credit note under section 34(1) must be referable to an identified '
                 'original tax invoice. Call for the invoice, the credit note and the contract.'
                 if none_found else
                 'Positive supply under the same HSN does exist. Check the value, the tax head and the section 34(2) '
                 'time limit against the original invoice.')
            c = ws.cell(row=r, column=8, value=a); c.alignment = Alignment(wrap_text=True, vertical='top')
            r += 1
        self.box(ws, 5, max(r-1, 5), 1, 8, numfrom=3)
        r += 1
        r = self.note(ws, r, 'NEGATIVE TAX DECLARED IN GSTR-1 BUT NOT CARRIED INTO GSTR-3B', bold=True)
        H2 = ['Return Period','GSTR-1 tax','GSTR-1 taxable value','GSTR-3B outward tax','Documents','Assessment']
        self.head(ws, r, H2, [13,18,20,19,12,96]); hdr = r; r += 1
        if not res['r1_only']:
            self.note(ws, r, 'None.'); r += 1
        for x in res['r1_only']:
            ws.cell(row=r, column=1, value=x['period']); ws.cell(row=r, column=2, value=x['r1_tax'])
            ws.cell(row=r, column=3, value=x['r1_value']); ws.cell(row=r, column=4, value=x['r3b_tax'])
            ws.cell(row=r, column=5, value=x['docs'])
            if x['inconsistent']:
                imp = abs(x['r1_tax']) / 0.18
                a = (f'INTERNALLY INCONSISTENT ON ITS FACE. The tax column is negative while the taxable value column '
                     f'is not. At 18% the implied credit-note value is about Rs.{inr(imp)}, none of which appears in the '
                     f'taxable value column. The return cannot be read from a summary — draw the full GSTR-1 table by '
                     f'table (B2B, CDNR, Table 12) from the portal.')
            else:
                a = ('Credit notes declared in GSTR-1 and not carried into GSTR-3B. The liability was therefore not '
                     'reduced in GSTR-3B. Establish what the credit notes relate to.')
            c = ws.cell(row=r, column=6, value=a); c.alignment = Alignment(wrap_text=True, vertical='top')
            r += 1
        self.box(ws, hdr+1, max(r-1, hdr+1), 1, 6, numfrom=2)
        self.note(ws, r+1, 'The percentage in column G is measured against every rupee of positive turnover the taxpayer '
                           'has declared across the whole record, which is the test of whether a credit note is proportionate '
                           'to the business that was actually done.')

    # ---------------------------------------- 14 / 15 / 16 / 18 / 19 / 19A
    def build_dormancy(self):
        res = self.R['dormancy']; ws = self.sheet('14_DORMANCY_ITC')
        if not res.get('available'): return self.unavailable(ws, res, 'ITC AVAILED IN NIL-TURNOVER PERIODS')
        ws['A1'] = 'ITC AVAILED IN TAX PERIODS DECLARING NIL OUTWARD SUPPLY'; ws['A1'].font = BOLD
        self.head(ws, 3, ['Return Period','FY','Turnover (3B)','Output tax (3B)','ITC availed','NIL flag',
                          'ITC availed in NIL period','Consecutive NIL count'], [13,9,18,16,16,10,20,18])
        for i, m in enumerate(self.R['master'], 4):
            mr = i
            for col, src in zip([1,2,3,4,5,6,7], ['A','B','D','J','M','W','X']):
                ws.cell(row=i, column=col, value=f'={self.m(src, mr)}')
            ws.cell(row=i, column=8, value='=IF(F4="NIL",1,0)' if i == 4 else f'=IF(F{i}="NIL",H{i-1}+1,0)')
        r = 3 + len(self.R['master']) + 1
        ws.cell(row=r, column=1, value='TOTAL').font = BOLD
        for c in [3,4,5,7]:
            LL = get_column_letter(c); ws.cell(row=r, column=c, value=f'=SUM({LL}4:{LL}{r-1})').font = BOLD
        ws.cell(row=r, column=8, value=f'=MAX(H4:H{r-1})').font = BOLD
        self.box(ws, 4, r, 1, 8, numfrom=3)
        for i in range(4, r+1): ws.cell(row=i, column=6).number_format = 'General'
        self.dorm_cell = None
        n = r + 2
        if res['longest']:
            a, b, c = res['longest']
            ws.cell(row=n, column=1, value=f'Longest unbroken NIL run: {c} tax periods, {a} to {b}.').font = BOLD
            ws.cell(row=n+1, column=1, value=f'ITC availed in NIL periods from {a} onwards').font = BOLD
            ws.cell(row=n+1, column=7, value=res['itc_in_nil_from_longest']).number_format = NUM
            self.dorm_cell = f"'14_DORMANCY_ITC'!G{n+1}"
            n += 2
        n = self.note(ws, n+1, 'Section 16(1) allows credit only on inputs and input services used or intended to be used '
                               'in the course or furtherance of business. Also section 17(5) for blocked credits, and Rule 86A.')
        self.note(ws, n, 'Credit accumulated through a dormancy is what later discharges liability without cash. '
                         'Compare with sheet 19.')

    def build_86b(self):
        res = self.R['rule86b']; ws = self.sheet('15_RULE_86B')
        if not res.get('available'): return self.unavailable(ws, res, 'RULE 86B TEST')
        ws['A1'] = 'RULE 86B — AT LEAST 1 PER CENT OF OUTPUT TAX TO BE PAID IN CASH'; ws['A1'].font = BOLD
        self.head(ws, 3, ['Return Period','Taxable turnover','Output tax','SGST cash paid','Cash proxy (SGST x 2)',
                          'Cash as % of output tax','Rule applicable? (taxable supply > Rs.50 lakh)','Prima facie result'],
                  [13,18,16,15,17,16,22,24])
        for i, m in enumerate(self.R['master'], 4):
            mr = i
            ws.cell(row=i, column=1, value=f'={self.m("A", mr)}')
            ws.cell(row=i, column=2, value=f'={self.m("E", mr)}')
            ws.cell(row=i, column=3, value=f'={self.m("J", mr)}')
            ws.cell(row=i, column=4, value=f'={self.m("T", mr)}')
            ws.cell(row=i, column=5, value=f'=D{i}*2')
            ws.cell(row=i, column=6, value=f'=IFERROR(IF(C{i}<=0,"",E{i}/C{i}),"")')
            ws.cell(row=i, column=7, value=f'=IF(B{i}>5000000,"YES","No")')
            ws.cell(row=i, column=8, value=f'=IF(G{i}="YES",IF(C{i}>0,IF(E{i}/C{i}<0.01,"BELOW 1% - EXAMINE",""),""),"")')
        r = 3 + len(self.R['master'])
        self.box(ws, 4, r, 1, 8, numfrom=2)
        for i in range(4, r+1):
            ws.cell(row=i, column=6).number_format = '0.00%'
            for c in (7, 8): ws.cell(row=i, column=c).number_format = 'General'
        n = self.note(ws, r+2, 'Periods flagged: ' + (', '.join(res['flagged']) if res['flagged'] else 'none'), bold=True)
        n = self.note(ws, n, 'ASSUMPTION: the source statement gives cash paid under the SGST head only. CGST cash is taken '
                             'as equal and IGST cash is not captured at all. Re-test on the electronic cash ledger, and '
                             'cross-check against the challans in sheet 19.')
        self.note(ws, n, 'Exceptions in the proviso to be ruled out first: income tax of more than Rs.1 lakh paid in each of '
                         'the two preceding financial years; refund of more than Rs.1 lakh on zero-rated or inverted duty; '
                         'cumulative discharge of 1% in cash from the start of the financial year; and the status of the person.')

    def build_rcm(self):
        res = self.R['rcm']; ws = self.sheet('16_RCM_CHECK')
        if not res.get('available'): return self.unavailable(ws, res, 'REVERSE CHARGE')
        ws['A1'] = 'REVERSE CHARGE — INWARD SUPPLIES DECLARED AND TAX IMPLIED BY THE RETURNS'; ws['A1'].font = BOLD
        self.head(ws, 3, ['Return Period','FY','RCM inward value (3.1(d))','GST OP Total','Outward tax (3.1(a))',
                          'Implied RCM liability','Implied rate','Remark'], [13,9,22,16,18,20,12,62])
        for i, m in enumerate(self.R['master'], 4):
            mr = i
            for col, src in zip([1,2,3,4,5,6], ['A','B','L','J','I','K']):
                ws.cell(row=i, column=col, value=f'={self.m(src, mr)}')
            ws.cell(row=i, column=7, value=f'=IFERROR(IF(C{i}=0,"",F{i}/C{i}),"")')
            rm = ''
            if m['rcm_value'] > 0 and abs(m['rcm_tax']) < 1:
                rm = (f'RCM inward value of Rs.{inr(m["rcm_value"])} declared but NO RCM tax implied by the return. '
                      f'Tax of about Rs.{inr(m["rcm_value"]*0.05)} at 5% appears not to have been discharged.')
            elif m['rcm_value'] > 0 and abs(m['rcm_tax'] / m['rcm_value'] - 0.05) > 0.002:
                rm = f'Implied rate {m["rcm_tax"]/m["rcm_value"]*100:.2f}% against 5.00% elsewhere. Reconcile.'
            elif m['period'] == res['last_period']:
                rm = 'Last period in which any reverse-charge inward supply was declared.'
            c = ws.cell(row=i, column=8, value=rm); c.alignment = Alignment(wrap_text=True, vertical='top')
        r = 3 + len(self.R['master']) + 1
        ws.cell(row=r, column=1, value='TOTAL').font = BOLD
        for c in [3,4,5,6]:
            LL = get_column_letter(c); ws.cell(row=r, column=c, value=f'=SUM({LL}4:{LL}{r-1})').font = BOLD
        self.box(ws, 4, r, 1, 8, numfrom=3)
        for i in range(4, r+1): ws.cell(row=i, column=7).number_format = '0.00%'
        n = r + 2
        if res['silent_periods']:
            n = self.note(ws, n, f'NO reverse-charge inward supply has been declared in any of the {res["silent_periods"]} '
                                 f'tax periods after {res["last_period"]}, although ITC was availed in most of them.', bold=True)
        self.note(ws, n, 'Sections 9(3) and 9(4) with Notification 13/2017-CT(R). Section 31(3)(f) self-invoicing. '
                         'For a working unit, nil RCM on goods transport, legal services, security services, director '
                         'remuneration and renting of immovable property over a long stretch requires explanation.')

    def build_filing(self):
        res = self.R['filing']; ws = self.sheet('18_RETURN_FILING')
        if not res.get('available'): return self.unavailable(ws, res, 'RETURN FILING RECORD')
        ws['A1'] = 'RETURN FILING RECORD — DATE OF FILING AGAINST THE STATUTORY DUE DATE'; ws['A1'].font = BOLD
        self.head(ws, 3, ['Return Period','FY','GSTR-3B date of filing','Statutory due date','Delay in days',
                          'GSTR-1 status','GSTR-2A report present?','3B turnover','3B output tax'],
                  [13,9,20,20,12,27,18,16,16])
        for i, row in enumerate(res['rows'], 4):
            mr = i
            ws.cell(row=i, column=1, value=row['period']); ws.cell(row=i, column=2, value=row['fy'])
            c = ws.cell(row=i, column=3, value=row['filed']); c.number_format = 'DD-MM-YYYY'
            if row['filed'] is None and self.mark: c.fill = MISSING_FILL
            c = ws.cell(row=i, column=4, value=row['due']); c.number_format = 'DD-MM-YYYY'
            ws.cell(row=i, column=5, value=f'=IF(C{i}="","",MAX(0,C{i}-D{i}))')
            c = ws.cell(row=i, column=6, value=f"='17_R1_vs_R3B'!C{i}")
            c2 = ws.cell(row=i, column=7, value='yes' if row['r2a_present'] else 'NOT SUPPLIED')
            if not row['r2a_present'] and self.mark: c2.fill = MISSING_FILL
            ws.cell(row=i, column=8, value=f'={self.m("D", mr)}')
            ws.cell(row=i, column=9, value=f'={self.m("J", mr)}')
        r = 3 + len(res['rows']) + 1
        ws.cell(row=r, column=1, value='TOTAL').font = BOLD
        for c in [5,8,9]:
            LL = get_column_letter(c); ws.cell(row=r, column=c, value=f'=SUM({LL}4:{LL}{r-1})').font = BOLD
        ws.cell(row=r+1, column=1, value='Tax periods filed after the due date').font = BOLD
        ws.cell(row=r+1, column=5, value=f'=COUNTIF(E4:E{r-1},">0")').font = BOLD
        self.box(ws, 4, r+1, 1, 9, numfrom=5)
        for i in range(4, r+2):
            for c in (3,4): ws.cell(row=i, column=c).number_format = 'DD-MM-YYYY'
            for c in (6,7): ws.cell(row=i, column=c).number_format = 'General'
        self.delay_cell = f"'18_RETURN_FILING'!E{r}"
        n = self.note(ws, r+3, 'The due date is taken as the 20th of the month following the tax period, the monthly filer '
                               'date. If the taxpayer was on QRMP in any period, that period must be re-set before late fee '
                               'is computed. Section 47(1) late fee; section 50 interest where tax was paid late.', bold=True)
        self.note(ws, n, 'Where GSTR-1 is not filed, the recipients of those supplies could not have taken credit through '
                         'GSTR-2B, so the identity of the recipients has to come from the invoices themselves. Section 37 read with Rule 59.')

    def build_cash(self):
        res = self.R['cash']; ws = self.sheet('19_CASH_RECON')
        if not res.get('available'): return self.unavailable(ws, res, 'CASH ACTUALLY PAID TO THE EXCHEQUER')
        ws['A1'] = 'CASH ACTUALLY PAID TO THE EXCHEQUER — every challan since registration'; ws['A1'].font = BOLD
        self.head(ws, 3, ['Payment month','Payment date','CIN','SGST','CGST','IGST','TOTAL','Running total'],
                  [15,14,26,15,15,15,16,16])
        last = self.src_rows.get('PAYMENTS')
        s = _q('SRC_PAYMENTS')
        for i, p in enumerate(res['challans'], 4):
            ws.cell(row=i, column=1, value=p.get('Payment Month'))
            c = ws.cell(row=i, column=2, value=p.get('Payment Date')); c.number_format = 'DD-MM-YYYY'
            ws.cell(row=i, column=3, value=p.get('CIN'))
            for col, k in zip([4,5,6,7], ['SGST','CGST','IGST','TOTAL']):
                ws.cell(row=i, column=col, value=n0(p.get(k)))
            ws.cell(row=i, column=8, value=f'=SUM($G$4:$G{i})')
        r = 3 + len(res['challans']) + 1
        ws.cell(row=r, column=3, value='TOTAL CASH EVER PAID').font = BOLD
        for c in range(4, 8):
            LL = get_column_letter(c); ws.cell(row=r, column=c, value=f'=SUM({LL}4:{LL}{r-1})').font = BOLD
        self.box(ws, 4, r, 1, 8, numfrom=4)
        for i in range(4, r+1):
            ws.cell(row=i, column=2).number_format = 'DD-MM-YYYY'
            for c in (1, 3): ws.cell(row=i, column=c).number_format = 'General'
        self.cash_cell = f"'19_CASH_RECON'!G{r}"
        s0 = r + 2
        ws.cell(row=s0, column=1, value='CASH AGAINST DECLARED LIABILITY').font = BOLD
        F = self.fy_last
        pairs = [('Number of challans since registration', f'=COUNTA(C4:C{r-1})'),
                 ('Total output tax declared in GSTR-3B', f"='11_FY_SUMMARY'!E{F}"),
                 ('Total ITC availed in GSTR-3B', f"='11_FY_SUMMARY'!F{F}"),
                 ('Total cash actually paid, all heads', f'=G{r}'),
                 ('Cash as a proportion of output tax declared', f'=IFERROR(B{s0+4}/B{s0+2},"")'),
                 ('Liability discharged otherwise than in cash', f'=B{s0+2}-B{s0+4}'),
                 ('ITC availed as a proportion of output tax', f'=IFERROR(B{s0+3}/B{s0+2},"")')]
        for i, (a, b) in enumerate(pairs, s0+1):
            ws.cell(row=i, column=1, value=a); ws.cell(row=i, column=2, value=b).number_format = NUM
        ws.cell(row=s0+5, column=2).number_format = '0.0%'
        ws.cell(row=s0+7, column=2).number_format = '0.0%'
        self.box(ws, s0+1, s0+len(pairs), 1, 2)
        t = s0 + len(pairs) + 2
        ws.cell(row=t, column=1, value='CASH BY FINANCIAL YEAR OF THE CHALLAN').font = BOLD
        self.head(ws, t+1, ['FY of challan','Cash paid'], [22, 18])
        for i, (f, v) in enumerate(sorted(res['by_fy'].items()), t+2):
            ws.cell(row=i, column=1, value=f); ws.cell(row=i, column=2, value=v).number_format = NUM
        self.box(ws, t+2, t+1+len(res['by_fy']), 1, 2, numfrom=2)
        self.note(ws, t+3+len(res['by_fy']), 'Payment Month is the month of the challan, which discharges the preceding tax '
                  'period. A run of very small challans is ordinarily late fee, not tax — check against the delays in sheet 18.',
                  bold=True)

    def build_itc_2a(self):
        res = self.R['itc_vs_2a']; ws = self.sheet('19A_ITC_vs_2A')
        if not res.get('available'): return self.unavailable(ws, res, 'ITC AGAINST GSTR-2A — INDEPENDENT ROUTE')
        ws['A1'] = 'INDEPENDENT CORROBORATION — ITC AVAILED AGAINST GSTR-2A, EVERY TAX PERIOD'; ws['A1'].font = BOLD
        ws['A2'] = ('Built from the GSTR-2A reports and the GSTR-3B payment summaries, neither of which feeds sheet 13. '
                    'Two independent routes to the same conclusion. Sheet 13 remains the demand figure because it compares '
                    'head by head; this sheet nets the heads within a period.')
        self.head(ws, 4, ['Return Period','FY','2A report present?','2A sellers','2A invoices','2A taxable value',
                          '2A credit available','ITC availed in GSTR-3B','GAP','Positive gap only','Remark'],
                  [13,9,16,12,13,18,20,19,18,16,66])
        s = _q('SRC_R2A_MONTHLY'); last = self.src_rows.get('R2A_MONTHLY')
        for i, row in enumerate(res['rows'], 5):
            mr = i - 1
            ws.cell(row=i, column=1, value=row['period']); ws.cell(row=i, column=2, value=row['fy'])
            c = ws.cell(row=i, column=3, value='yes' if row['present'] else 'NOT IN DATA')
            if not row['present'] and self.mark: c.fill = MISSING_FILL
            for col, k in zip([4,5,6,7], ['sellers','invoices','taxable','avail']):
                ws.cell(row=i, column=col, value=row[k] if row[k] is not None else '')
            ws.cell(row=i, column=8, value=f'={self.m("M", mr)}')
            ws.cell(row=i, column=9, value=f'=IF($C{i}="NOT IN DATA","",H{i}-G{i})')
            ws.cell(row=i, column=10, value=f'=IF(I{i}="",0,MAX(0,I{i}))')
            rm = ''
            if not row['present']:
                rm = ('No GSTR-2A report in the data for this period. ' +
                      ('ITC availed is nil, so nothing turns on it.' if row['itc'] == 0
                       else f'Rs.{inr(row["itc"])} of credit availed cannot be tested.'))
            elif row['gap'] and row['gap'] > 100000:
                rm = (f'Only {inr(row["sellers"])} suppliers and {inr(row["invoices"])} invoices in 2A carrying '
                      f'Rs.{inr(row["avail"])}, against Rs.{inr(row["itc"])} availed.')
            elif row['gap'] and row['gap'] < -100000:
                rm = 'Credit available but not taken, or credit reversed in this period. Ordinarily timing.'
            c = ws.cell(row=i, column=11, value=rm); c.alignment = Alignment(wrap_text=True, vertical='top')
        r = 4 + len(res['rows']) + 1
        ws.cell(row=r, column=1, value='TOTAL').font = BOLD
        for c in range(4, 11):
            LL = get_column_letter(c); ws.cell(row=r, column=c, value=f'=SUM({LL}5:{LL}{r-1})').font = BOLD
        self.box(ws, 5, r, 1, 11, numfrom=4)
        for i in range(5, r+1): ws.cell(row=i, column=3).number_format = 'General'
        t = r + 2
        ws.cell(row=t, column=1, value='FINANCIAL-YEAR ROLL-UP, AND THE SAME TEST RUN AGAIN OFF THE ANNUAL SUMMARIES').font = BOLD
        self.head(ws, t+1, ['FY','2A credit available (monthly reports)','ITC availed in periods that HAVE a 2A report',
                            'GAP (monthly route)','ITC availed in periods with NO 2A report — not testable',
                            '2A credit available (annual summary)','ITC availed (annual summary)','GAP (annual route)',
                            'Do the two routes agree?'],
                  [13,22,24,18,24,22,22,18,18])
        fys = sorted(self.R['fy'], key=lambda x: int(x.split('-')[0]))
        ann = res.get('annual', {})
        for i, f in enumerate(fys, t+2):
            ws.cell(row=i, column=1, value=f)
            ws.cell(row=i, column=2, value=f'=SUMIFS($G$5:$G${r-1},$B$5:$B${r-1},$A{i},$C$5:$C${r-1},"yes")')
            ws.cell(row=i, column=3, value=f'=SUMIFS($H$5:$H${r-1},$B$5:$B${r-1},$A{i},$C$5:$C${r-1},"yes")')
            ws.cell(row=i, column=4, value=f'=C{i}-B{i}')
            ws.cell(row=i, column=5, value=f'=SUMIFS($H$5:$H${r-1},$B$5:$B${r-1},$A{i},$C$5:$C${r-1},"NOT IN DATA")')
            if f in ann:
                ws.cell(row=i, column=6, value=ann[f]['avail']); ws.cell(row=i, column=7, value=ann[f]['itc'])
                ws.cell(row=i, column=8, value=f'=G{i}-F{i}')
                ws.cell(row=i, column=9, value=f'=IF(ABS((D{i}+E{i})-H{i})<100,"yes","CHECK")')
            else:
                for c in (6,7,8): self.na(ws, i, c, '-')
                ws.cell(row=i, column=9, value='annual summary not supplied')
        r2 = t + 2 + len(fys)
        ws.cell(row=r2, column=1, value='TOTAL').font = BOLD
        for c in range(2, 9):
            LL = get_column_letter(c); ws.cell(row=r2, column=c, value=f'=SUM({LL}{t+2}:{LL}{r2-1})').font = BOLD
        self.box(ws, t+2, r2, 1, 9, numfrom=2)
        for i in range(t+2, r2+1): ws.cell(row=i, column=9).number_format = 'General'
        self.note(ws, r2+2, 'Column E isolates credit availed in periods for which no GSTR-2A report exists. The monthly route '
                            'cannot test it; the annual summary silently includes it. Keeping it in its own column is what '
                            'makes the two routes reconcile.', bold=True)

    def build_hsn(self):
        res = self.R['hsn']; ws = self.sheet('12_HSN_PROFILE')
        if not res.get('available'): return self.unavailable(ws, res, 'HSN PROFILE')
        ws['A1'] = 'HSN PROFILE — WHAT THE TAXPAYER SAYS IT SUPPLIES, IN THE ORDER IT SAID IT'; ws['A1'].font = BOLD
        self.head(ws, 3, ['Return Period','FY','HSN code','Taxable value','Total GST','Note'], [13,9,14,18,18,70])
        for i, (p, code, v, t) in enumerate(res['rows'], 4):
            ws.cell(row=i, column=1, value=p); ws.cell(row=i, column=2, value=fy_of(p))
            ws.cell(row=i, column=3, value=code); ws.cell(row=i, column=4, value=v); ws.cell(row=i, column=5, value=t)
            n = 'NEGATIVE — credit note. See 19B_CREDIT_NOTES.' if v < 0 else ''
            c = ws.cell(row=i, column=6, value=n); c.alignment = Alignment(wrap_text=True, vertical='top')
        r = 3 + len(res['rows']) + 1
        ws.cell(row=r, column=1, value='TOTAL').font = BOLD
        for c in (4, 5):
            LL = get_column_letter(c); ws.cell(row=r, column=c, value=f'=SUM({LL}4:{LL}{r-1})').font = BOLD
        self.box(ws, 4, r, 1, 6, numfrom=4)
        for i in range(4, r+1): ws.cell(row=i, column=3).number_format = 'General'
        n = self.note(ws, r+2, 'HSN codes in the order they first appear: ' + ' -> '.join(res['codes']), bold=True)
        self.note(ws, n, 'A drift away from the registered nature of business, and in particular the appearance of goods '
                         'HSNs in a services taxpayer, points to disposal of capital goods. Section 18(6) with Rule 44(6): '
                         'tax is the higher of the credit taken less the prescribed reduction, or the tax on transaction value.')

    def build_limitation(self):
        ws = self.sheet('24_LIMITATION')
        ws['A1'] = 'LIMITATION — INDICATIVE OUTER DATES BY FINANCIAL YEAR'; ws['A1'].font = BOLD
        ws['A2'] = ('COMPUTED FROM THE STATUTE ALONE. Extension notifications are NOT applied and several have been the '
                    'subject of litigation. Verify every date against the notifications in force before acting on it.')
        self.head(ws, 4, ['FY','Annual return due','Section 73 — SCN by','Section 73 — order by','Section 73 still open?',
                          'Section 74 — SCN by','Section 74 — order by','Section 74 still open?','Regime',
                          'Section 74A — SCN by (42 months)','Section 74A — order by (12 months, +6 extendable)',
                          'Section 74A still open?'],
                  [10,18,18,18,18,18,18,18,40,20,22,16])
        for i, l in enumerate(self.R['limitation'], 5):
            ws.cell(row=i, column=1, value=l['fy'])
            c = ws.cell(row=i, column=2, value=l['annual_return_due']); c.number_format = 'DD-MM-YYYY'
            if l.get('is74a'):
                for col in (3, 4, 5, 6, 7, 8):
                    ws.cell(row=i, column=col, value='not applicable')
                for col, k in zip([10, 11], ['s74a_scn_by', 's74a_order_by']):
                    c = ws.cell(row=i, column=col, value=l[k]); c.number_format = 'DD-MM-YYYY'
                ws.cell(row=i, column=12, value='OPEN' if l['s74a_open'] else 'TIME-BARRED')
            else:
                for col, k in zip([3,4,6,7], ['s73_scn_by','s73_order_by','s74_scn_by','s74_order_by']):
                    c = ws.cell(row=i, column=col, value=l[k]); c.number_format = 'DD-MM-YYYY'
                c = ws.cell(row=i, column=5, value='OPEN' if l['s73_open'] else 'TIME-BARRED')
                if not l['s73_open'] and self.mark: c.fill = MISSING_FILL
                ws.cell(row=i, column=8, value='OPEN' if l['s74_open'] else 'TIME-BARRED')
                for col in (10, 11, 12): ws.cell(row=i, column=col, value='not applicable')
            ws.cell(row=i, column=9, value=l['regime'])
        r = 4 + len(self.R['limitation'])
        self.box(ws, 5, r, 1, 12)
        for i in range(5, r+1):
            for c in (3, 4, 5, 6, 7, 8, 9, 10, 11, 12):
                if not isinstance(ws.cell(row=i, column=c).value, datetime.date):
                    ws.cell(row=i, column=c).number_format = 'General'
        n = self.note(ws, r+2, 'Where section 73 is time-barred, a demand for that year can be raised only under section 74, '
                               'which requires fraud, wilful misstatement or suppression of facts to be alleged and made out '
                               'on evidence. That decision has to be taken before the evidence is gathered, not after.', bold=True)
        self.note(ws, n, 'Section 74A applies to demands for FY 2024-25 onwards (sections 73(12) and 74(12) confine 73 and 74 to '
                         'FY 2023-24 and earlier): notice within 42 months of the due date of the annual return (s.74A(2)), order '
                         'within 12 months of the notice, extendable by up to 6 months (s.74A(10)). One limitation for both the '
                         'non-fraud and the fraud case; the difference lies in the penalty.')


    # ------------------------------------------------- 19C GSTR-7 TDS (v1.5)
    def _src_range(self, code, fy, colname):
        """Contiguous rows of an SRC_ sheet that belong to one FY (SRC sheets are sorted by period),
        as an A1 range, so an FY total stays a live SUM over the source cells."""
        if code not in self.src_rows: return None
        rows = sorted(self.ing.data[code], key=lambda r: (period_key(r['_period']), r.get('SNo') or 0))
        idx = [i for i, r in enumerate(rows, 2) if r.get('_fy') == fy]
        col = self.src_col(code, colname)
        if not idx or col is None: return None
        return f"{_q(f'SRC_{code}'[:31])}!{col}{idx[0]}:{col}{idx[-1]}"

    def _master_range(self, fy, col, upto=None):
        rows = [i for i, m in enumerate(self.R['master'], 4)
                if m['fy'] == fy and (upto is None or period_key(m['period']) <= period_key(upto))]
        if not rows: return None
        return f"'10_MONTHLY_MASTER'!{col}{rows[0]}:{col}{rows[-1]}"

    def build_tds(self):
        res = self.R['tds']; ws = self.sheet('19C_TDS_GSTR7')
        if not res.get('available'):
            return self.unavailable(ws, res, 'TAX DEDUCTED AT SOURCE BY GOVERNMENT (GSTR-7)')
        ws['A1'] = ('TAX DEDUCTED AT SOURCE FROM THIS TAXPAYER BY GOVERNMENT DEDUCTORS (GSTR-7) — '
                    'what the deductors paid for, against what the taxpayer declared'); ws['A1'].font = BOLD
        ws['A2'] = ('Section 51 CGST Act with Notification 50/2018-CT: a Government department, local authority or Government '
                    'agency paying a supplier under a contract above Rs.2.5 lakh deducts 1% CGST + 1% SGST (or 2% IGST) of the '
                    'value of supply, excluding the tax on the invoice (s.51(2)), files GSTR-7 (Rule 66), and the amount is '
                    'credited to the supplier\'s electronic cash ledger (Rule 87(9)). TDS / 2% is therefore the value of '
                    'supplies the deductors have PAID for — a floor on turnover, not the whole of it.')
        ws['A2'].alignment = Alignment(wrap_text=True, vertical='top')
        ws.merge_cells('A2:T2'); ws.row_dimensions[2].height = 48
        ws['A3'] = 'Rate of tax ASSUMED for the indicative figure only — confirm from the contracts and invoices'
        ws['A3'].font = BOLD
        rc = ws['H3']; rc.value = 0.18; rc.number_format = '0%'; rc.font = BOLD; rc.border = BD
        self.tds_rate_cell = "'19C_TDS_GSTR7'!$H$3"
        H = ['Return Period','FY','No. of deductors','SGST deducted','CGST deducted','IGST deducted','Cess',
             'Total GST (as reported)','Rounding check (heads + cess - total)','VALUE PAID FOR by the deductors (=(total-cess)/2%)',
             'Cumulative value paid for, in the FY','GSTR-3B status for the month','GSTR-3B taxable turnover, the month',
             'Cumulative GSTR-3B taxable turnover, FY to this month','CUMULATIVE SHORTFALL (value paid for less declared)',
             'GSTR-3B outward CGST, the month','CGST deducted exceeds CGST declared?','GSTR-1 taxable value, the month',
             'Above Rs.50 lakh (Rule 86B)?','Observation']
        self.head(ws, 5, H, [12,9,10,13,13,12,8,14,12,18,18,20,17,18,18,15,13,16,12,70])
        P = lambda r: f'$A{r}'
        mrow = {m['period']: i for i, m in enumerate(self.R['master'], 4)}
        ML = getattr(self, 'master_last', 3)
        M = _q('10_MONTHLY_MASTER')
        first = 6
        for i, x in enumerate(res['rows'], first):
            ws.cell(row=i, column=1, value=x['period'])
            ws.cell(row=i, column=2, value=x['fy'])
            for col, nm in zip(range(3, 9), ['No Of Deductors','SGST','CGST','IGST','CESS','Total GST']):
                ws.cell(row=i, column=col, value=self.lookup('R7_TDS', nm, 'Return Period', P(i)))
            ws.cell(row=i, column=9, value=f'=D{i}+E{i}+F{i}+G{i}-H{i}')
            ws.cell(row=i, column=10, value=f'=(H{i}-G{i})/0.02')
            ws.cell(row=i, column=11, value=f'=SUMIF($B${first}:B{i},B{i},$J${first}:J{i})')
            if res['have_3b']:
                c = ws.cell(row=i, column=12, value=x['r3b_status'])
                if x['r3b_status'] not in ('FILED', 'NOT YET DUE') and self.mark: c.fill = MISSING_FILL
                if x['period'] in mrow:
                    ws.cell(row=i, column=13, value=f'=INDEX({M}!$E$4:$E${ML},MATCH($A{i},{M}!$A$4:$A${ML},0))')
                    ws.cell(row=i, column=16, value=f'=INDEX({M}!$G$4:$G${ML},MATCH($A{i},{M}!$A$4:$A${ML},0))')
                    ws.cell(row=i, column=17, value=f'=IF(E{i}>P{i}+1,"YES","")')
                else:
                    self.na(ws, i, 13, 'NO 3B'); self.na(ws, i, 16, 'NO 3B')
                rng = self._master_range(x['fy'], 'E', upto=x['period'])
                ws.cell(row=i, column=14, value=f'=SUM({rng})' if rng else 0)
                if x['r3b_status'] == 'NOT YET DUE':
                    self.na(ws, i, 15, 'NOT YET DUE')
                else:
                    ws.cell(row=i, column=15, value=f'=MAX(0,K{i}-N{i})')
            else:
                for col in (12, 13, 14, 15, 16, 17): self.na(ws, i, col, '3B NOT SUPPLIED')
            if res['have_r1']:
                if x['r1_status'] in ('NOT FILED', 'NOT YET DUE'):
                    self.na(ws, i, 18, x['r1_status'])
                else:
                    ws.cell(row=i, column=18, value=self.lookup('R1_MONTHLY', 'Taxable Value', 'Return Period', P(i)))
            else:
                self.na(ws, i, 18, 'R1 NOT SUPPLIED')
            ws.cell(row=i, column=19, value=f'=IF(J{i}>5000000,"YES","")')
            ws.cell(row=i, column=20, value='; '.join(x['obs']))
        last = first + len(res['rows']) - 1
        t = last + 1
        ws.cell(row=t, column=1, value='TOTAL').font = BOLD
        for c in (4, 5, 6, 7, 8, 10):
            L = get_column_letter(c); ws.cell(row=t, column=c, value=f'=SUM({L}{first}:{L}{last})').font = BOLD
        self.box(ws, first, t, 1, 20, numfrom=3)
        for r in range(first, t + 1):
            for c in (1, 2, 12, 17, 19, 20):
                ws.cell(row=r, column=c).number_format = 'General'
                ws.cell(row=r, column=c).alignment = Alignment(wrap_text=True, vertical='top')
        self.tds_first, self.tds_last = first, last

        # ---- financial-year roll-up
        r0 = t + 3
        ws.cell(row=r0 - 1, column=1, value='FINANCIAL-YEAR POSITION').font = BOLD
        FH = ['FY','Months with TDS','TDS deducted','VALUE PAID FOR by the deductors','GSTR-3B months on record',
              'GSTR-3B taxable turnover, months on record','SHORTFALL against GSTR-3B (months not yet due left out)','Indicative tax on the shortfall at the assumed rate',
              'GSTR-1 taxable value','SHORTFALL against GSTR-1 (months not yet due left out)','SGST deducted (credited to cash ledger)',
              'SGST paid in cash in GSTR-3B','SGST TDS NOT ABSORBED (balance building up)','ITC availed in GSTR-3B',
              'GSTR-9 required on receipts subject to TDS alone (> Rs.2 crore)?','GSTR-9C required (> Rs.5 crore)?']
        for j, h in enumerate(FH, 1):
            x = ws.cell(row=r0, column=j, value=h); x.font = BOLD; x.border = BD
            x.alignment = Alignment(wrap_text=True, vertical='top', horizontal='center')
        ws.row_dimensions[r0].height = 75
        B = f'$B${first}:$B${last}'
        fys = sorted(res['fy'], key=lambda f: int(f.split('-')[0]))
        for i, f in enumerate(fys, r0 + 1):
            a = res['fy'][f]
            ws.cell(row=i, column=1, value=f)
            ws.cell(row=i, column=2, value=f'=COUNTIF({B},$A{i})')
            ws.cell(row=i, column=3, value=f'=SUMIF({B},$A{i},$H${first}:$H${last})-SUMIF({B},$A{i},$G${first}:$G${last})')
            ws.cell(row=i, column=4, value=f'=SUMIF({B},$A{i},$J${first}:$J${last})')
            if res['have_3b']:
                ws.cell(row=i, column=5, value=f'=COUNTIF({M}!$B$4:$B${ML},$A{i})')
                ws.cell(row=i, column=6, value=f'=SUMIF({M}!$B$4:$B${ML},$A{i},{M}!$E$4:$E${ML})')
                ws.cell(row=i, column=7, value=f'=MAX(0,SUMIFS($J${first}:$J${last},{B},$A{i},$L${first}:$L${last},"<>NOT YET DUE")-F{i})')
                ws.cell(row=i, column=8, value=f'=G{i}*$H$3')
                ws.cell(row=i, column=14, value=f'=SUMIF({M}!$B$4:$B${ML},$A{i},{M}!$M$4:$M${ML})')
            else:
                for c in (5, 6, 7, 8, 14): self.na(ws, i, c, '3B NOT SUPPLIED')
            if res['have_r1']:
                rng = self._src_range('R1_MONTHLY', f, 'Taxable Value')
                ws.cell(row=i, column=9, value=f'=SUM({rng})' if rng else 0)
                ws.cell(row=i, column=10, value=f'=MAX(0,SUMIFS($J${first}:$J${last},{B},$A{i},$R${first}:$R${last},"<>NOT YET DUE")-I{i})')
            else:
                for c in (9, 10): self.na(ws, i, c, 'R1 NOT SUPPLIED')
            ws.cell(row=i, column=11, value=f'=SUMIF({B},$A{i},$D${first}:$D${last})')
            if res['have_tax_paid']:
                ws.cell(row=i, column=12, value=f'=SUMIF({M}!$B$4:$B${ML},$A{i},{M}!$T$4:$T${ML})')
                ws.cell(row=i, column=13, value=f'=MAX(0,K{i}-L{i})')
            else:
                for c in (12, 13): self.na(ws, i, c, '3B NOT SUPPLIED')
            ws.cell(row=i, column=15, value=f'=IF(D{i}>20000000,"YES","not on these receipts alone")')
            ws.cell(row=i, column=16, value=f'=IF(D{i}>50000000,"YES","not on these receipts alone")')
        fl = r0 + len(fys)
        self.box(ws, r0 + 1, fl, 1, 16, numfrom=2)
        for r in range(r0 + 1, fl + 1):
            for c in (1, 15, 16): ws.cell(row=r, column=c).number_format = 'General'
        self.tds_fy_rows = {f: i for i, f in enumerate(fys, r0 + 1)}

        # ---- the tests, their legal basis, and whether each could be run
        r = fl + 3
        ws.cell(row=r - 1, column=1, value='THE CROSS-CHECKS — legal basis, and whether each could be run on the data held').font = BOLD
        TH = ['No.','Test','Provision','Reports it needs','Status','Result']
        for j, h in enumerate(TH, 1):
            x = ws.cell(row=r, column=j, value=h); x.font = BOLD; x.border = BD
        ws.merge_cells(start_row=r, start_column=6, end_row=r, end_column=20)
        r += 1
        for k, (test, prov, needs, ok, result) in enumerate(res['tests'], 1):
            vals = [k, test, prov, needs, 'RUN' if ok else 'NOT RUN — upload the reports named', result]
            for j, v in enumerate(vals, 1):
                c = ws.cell(row=r, column=j, value=v); c.border = BD
                c.alignment = Alignment(wrap_text=True, vertical='top')
            if not ok and self.mark: ws.cell(row=r, column=5).fill = MISSING_FILL
            ws.merge_cells(start_row=r, start_column=6, end_row=r, end_column=20)
            ws.row_dimensions[r].height = 62
            r += 1
        r += 1
        r = self.note(ws, r, 'HOW TO READ THE SHORTFALL. Payment month is not invoice month, so a single month can show a '
                             'shortfall that is only timing. The firm test is the CUMULATIVE column: for services (works contract '
                             'is a service) the time of supply is the earlier of invoice and payment (s.13(2)), and for goods the '
                             'invoice is issued on removal (s.31(1)), so a supply the deductors have paid for by a month must have been '
                             'declared by that month. Exception: an advance for goods is not taxed on receipt (Notification '
                             '66/2017-CT) — rule it out on the contract before relying on a shortfall.', bold=True)
        r = self.note(ws, r, 'The value paid for is a FLOOR. Supplies to buyers who do not deduct, and contracts of Rs.2.5 lakh or '
                             'less, carry no TDS and are not in it. A GSTR-3B turnover above it proves nothing; one below it is a '
                             'shortfall on business with the deductors alone.')
        self.note(ws, r, 'The rate in H3 is an assumption used only for the indicative tax in column H of the FY block. It is not '
                         'carried into 20_FINDINGS. Works contracts for Government are ordinarily 18% under entry 3 of Notification '
                         '11/2017-CT(Rate) as amended, but the rate turns on the contract — confirm it before any figure is quoted.')
        ws.freeze_panes = 'C6'


    # ------------------------------------------ 19D GSTR-7 party-wise (v1.7)
    def build_tds_parties(self):
        res = self.R.get('tds_detail') or {}
        ws = self.sheet('19D_TDS_PARTIES')
        if not res.get('available'):
            return self.unavailable(ws, res, 'GSTR-7 PARTY-WISE DETAIL — who deducted, on what value')
        ws['A1'] = 'GSTR-7 PARTY-WISE DETAIL — who the parties are, whether section 51 lets them deduct, and what was deducted'
        ws['A1'].font = BOLD
        ws['A2'] = ('Export titled for GSTIN: ' + (', '.join(res['subject']) or 'not stated')
                    + ('   — DOES NOT MATCH THE CASE GSTIN' if res['subject_mismatch'] else ''))
        ws['A2'].font = BOLD
        if res['subject_mismatch'] and self.mark: ws['A2'].fill = MISSING_FILL
        ws['A3'] = 'WHO IS WHO. ' + res['role']
        ws['A3'].alignment = Alignment(wrap_text=True, vertical='top'); ws.merge_cells('A3:K3')
        ws.row_dimensions[3].height = 62
        ws['A4'] = ('WHO MAY DEDUCT. Section 51(1) read with Notification 50/2018-CT: (a) a department or establishment of '
                    'the Central or a State Government; (b) a local authority; (c) a Governmental agency; (d) an authority, '
                    'board or body set up by Parliament, a State Legislature or a Government with 51% or more Government '
                    'equity or control; a society established by a Government or local authority; a public sector '
                    'undertaking. A private proprietor, firm or company outside these is not a deductor, and a tax-deductor '
                    'registration granted to one is itself irregular.')
        ws['A4'].alignment = Alignment(wrap_text=True, vertical='top'); ws.merge_cells('A4:K4')
        ws.row_dimensions[4].height = 62
        S = _q('SRC_R7_DETAIL'); last = self.src_rows['R7_DETAIL']
        cg = self.src_col('R7_DETAIL', 'Deductee GSTIN'); cv = self.src_col('R7_DETAIL', 'Taxable Value')
        ct = self.src_col('R7_DETAIL', 'Total GST'); cp = self.src_col('R7_DETAIL', 'Return Period')
        rg = lambda c: f'{S}!${c}$2:${c}${last}'
        r = 6
        H = ['GSTIN listed','Trade name','Registration type (GSTIN position 14)','PAN / TAN holder type (PAN position 4)',
             'Can it deduct under section 51?','Months listed','Lines','Taxable value','TDS (Total GST)','Share of value',
             'Note']
        self.head(ws, r, H, [18,26,24,24,58,22,7,15,13,9,40]); ws.freeze_panes = None
        p0 = r + 1
        for i, p in enumerate(res['parties'], p0):
            vals = [p['gstin'], ', '.join(p['names']), p['reg'], p['holder'], p['eligibility'], ', '.join(p['months'])]
            for j, v in enumerate(vals, 1): ws.cell(row=i, column=j, value=v)
            ws.cell(row=i, column=7, value=f'=COUNTIF({rg(cg)},$A{i})')
            ws.cell(row=i, column=8, value=f'=SUMIF({rg(cg)},$A{i},{rg(cv)})')
            ws.cell(row=i, column=9, value=f'=SUMIF({rg(cg)},$A{i},{rg(ct)})')
            if p['eligibility'].startswith('NOT A PERMITTED') and self.mark:
                ws.cell(row=i, column=5).fill = MISSING_FILL
        pl = p0 + len(res['parties']) - 1
        for i in range(p0, pl + 1):
            ws.cell(row=i, column=10, value=f'=IFERROR(H{i}/H${pl+1},"")').number_format = '0.0%'
        ws.cell(row=pl + 1, column=1, value='TOTAL').font = BOLD
        for c in (7, 8, 9):
            L = get_column_letter(c); ws.cell(row=pl + 1, column=c, value=f'=SUM({L}{p0}:{L}{pl})').font = BOLD
        self.box(ws, p0, pl + 1, 1, 11, numfrom=7)
        for i in range(p0, pl + 2):
            for c in range(1, 12): ws.cell(row=i, column=c).alignment = Alignment(wrap_text=True, vertical='top')
            ws.cell(row=i, column=10).number_format = '0.0%'
        # ---- month reconciliation with the summary report
        r = pl + 4
        ws.cell(row=r - 1, column=1, value='MONTH BY MONTH — the detail against the summary report (R7 FILINGS IN THE YEAR)').font = BOLD
        H = ['Return Period','Lines','Parties','Taxable value (detail)','TDS (detail)','Parties per summary',
             'TDS per summary','TDS difference','Value implied by the summary (TDS / 2%)','Agrees?']
        for j, h in enumerate(H, 1):
            x = ws.cell(row=r, column=j, value=h); x.font = BOLD; x.border = BD
            x.alignment = Alignment(wrap_text=True, vertical='top', horizontal='center')
        m0 = r + 1
        for i, m in enumerate(res['months'], m0):
            ws.cell(row=i, column=1, value=m['period'])
            ws.cell(row=i, column=2, value=f'=COUNTIF({rg(cp)},$A{i})')
            ws.cell(row=i, column=3, value=m['parties'])
            ws.cell(row=i, column=4, value=f'=SUMIF({rg(cp)},$A{i},{rg(cv)})')
            ws.cell(row=i, column=5, value=f'=SUMIF({rg(cp)},$A{i},{rg(ct)})')
            if 'R7_TDS' in self.src_rows:
                ws.cell(row=i, column=6, value=self.lookup('R7_TDS', 'No Of Deductors', 'Return Period', f'$A{i}'))
                ws.cell(row=i, column=7, value=self.lookup('R7_TDS', 'Total GST', 'Return Period', f'$A{i}'))
                ws.cell(row=i, column=8, value=f'=E{i}-G{i}')
                ws.cell(row=i, column=9, value=f'=G{i}/0.02')
                ws.cell(row=i, column=10, value=f'=IF(AND(ABS(H{i})<=1+0.5*B{i},C{i}=F{i}),"YES","NO — CHECK")')
            else:
                for c in (6, 7, 8, 9, 10): self.na(ws, i, c, 'SUMMARY NOT SUPPLIED')
        ml = m0 + len(res['months']) - 1
        self.box(ws, m0, ml, 1, 10, numfrom=2)
        r = ml + 2
        if res['missing_months']:
            r = self.note(ws, r, 'DETAIL NOT YET UPLOADED for ' + ', '.join(res['missing_months']) + ' — months the summary '
                          'shows TDS for. Export "R7 Filing Details For The Month" for each and upload under the same GSTIN.',
                          bold=True)
            if self.mark: ws.cell(row=r - 1, column=1).fill = MISSING_FILL
        if res['near_threshold']:
            r = self.note(ws, r + 1, f'{len(res["near_threshold"])} line(s) carry a taxable value just above the Rs.2.5 lakh '
                          f'contract value at which section 51 deduction begins: '
                          + '; '.join(f'{p} {g} Rs.{inr(v)}' for p, g, v in res['near_threshold'])
                          + '. The threshold applies to the value of the contract, not to each invoice — establish what contract '
                            'each line was deducted under.')
        if res['off_rate']:
            r = self.note(ws, r + 1, 'Lines where tax deducted is not 1% + 1% (or 2% IGST) of the taxable value: '
                          + '; '.join(f'{p} {g} value Rs.{inr(v)} CGST Rs.{inr(c)}' for p, g, v, c in res['off_rate']), bold=True)
        else:
            r = self.note(ws, r + 1, 'Every line was deducted at 1% CGST + 1% SGST of its taxable value (to the rupee).')
        ws.column_dimensions['A'].width = 18

    # ------------------------------------------------------- 20 FINDINGS
    def findings(self):
        R = self.R; F = []
        cn = R['credit_notes']
        if cn.get('available'):
            for e in cn['events']:
                none_found = any(not v for v in e['prior_positive_same_hsn'].values()) or not e['hsn']
                F.append(dict(
                    t=f'Credit note of Rs.{inr(abs(e["turnover"]))} in {e["period"]}'
                      + (' with no traceable original supply' if none_found else ''),
                    w=(f'GSTR-3B for {e["period"]} declares taxable turnover of MINUS Rs.{inr(abs(e["turnover"]))} and '
                       f'outward tax of MINUS Rs.{inr(abs(e["tax"]))}'
                       + (f', under HSN {", ".join(e["hsn"])}' if e['hsn'] else '') + '. '
                       + ('The taxpayer has NEVER declared a positive outward supply under that HSN in any tax period '
                          'on record. ' if none_found else '')
                       + f'It equals {e["pct_of_all_positive_turnover"]*100:.1f}% of every rupee of positive turnover '
                         f'declared since registration.'),
                    a=abs(e['tax']), s='19B_CREDIT_NOTES; 12_HSN_PROFILE',
                    p=('Section 34(1) and 34(2) — a credit note must be referable to an identified original tax invoice '
                       'and the reduction is available only up to 30 November following the end of the financial year of '
                       'that supply. Section 15.'),
                    r='VERY HIGH'))
            for x in cn['r1_only']:
                F.append(dict(
                    t=f'{x["period"]}: credit note of Rs.{inr(abs(x["r1_tax"]))} in GSTR-1, not carried into GSTR-3B',
                    w=(f'GSTR-1 declares tax of MINUS Rs.{inr(abs(x["r1_tax"]))} against a taxable value of '
                       f'{"PLUS" if x["r1_value"]>=0 else "MINUS"} Rs.{inr(abs(x["r1_value"]))}'
                       + (f' on {inr(x["docs"])} documents' if x['docs'] else '') + '. '
                       + ('The two figures cannot both be right: the return is internally inconsistent on its face. '
                          f'At 18% the implied credit-note value is about Rs.{inr(abs(x["r1_tax"])/0.18)}, none of which '
                          'appears in the taxable value column. ' if x['inconsistent'] else '')
                       + 'Nothing was carried into GSTR-3B, so no amount is demanded here. The entry is listed because '
                         'of what it points to.'),
                    a=0.0, s='19B_CREDIT_NOTES; 17_R1_vs_R3B',
                    p='Section 34, section 37, section 70. Also establish whether any audit, scrutiny or intelligence '
                      'proceeding was on foot in that period.',
                    r='VERY HIGH — lead'))
        ex = R['excess_itc']; i2 = R['itc_vs_2a']
        if ex.get('available'):
            corr = ''
            if i2.get('available') and i2.get('annual'):
                worst = max(i2['annual'].items(), key=lambda kv: kv[1]['gap'], default=None)
                if worst and worst[1]['gap'] > 0:
                    corr = (f' INDEPENDENTLY CORROBORATED: for FY {worst[0]} the GSTR-2A summary shows credit available '
                            f'of Rs.{inr(worst[1]["avail"])} against ITC availed of Rs.{inr(worst[1]["itc"])} — a gap of '
                            f'Rs.{inr(worst[1]["gap"])}, reached from reports that do not feed this computation.')
            F.append(dict(
                t='Excess input tax credit availed over GSTR-2A',
                w=(f'ITC availed in GSTR-3B exceeds credit available in GSTR-2A, in at least one head, in '
                   f'{ex["months_excess"]} of the {len(ex["rows"])} tax periods (SGST in {ex["months_excess_sgst"]}, '
                   f'IGST in {ex["months_excess_igst"]}). Net excess after setting off the periods of short claim is '
                   f'Rs.{inr(ex["net"])}; gross, counting only periods of over-claim, Rs.{inr(ex["gross"])}.' + corr),
                a=ex['net'], s='13_EXCESS_ITC; 19A_ITC_vs_2A',
                p='Section 16(2)(aa) and Rule 36(4). Rule 88D / FORM DRC-01C. Rule 86A for blocking.',
                r='VERY HIGH'))
        r1 = R['r1_vs_r3b']
        if r1.get('available'):
            det = [x for x in r1['rows'] if x['short'] > 1]
            if det:
                F.append(dict(
                    t='Output tax declared in GSTR-1 but not discharged in GSTR-3B',
                    w=('; '.join(f'{x["period"]} Rs.{inr(x["short"])}' for x in det)
                       + f'. Rs.{inr(r1["short_total"])} in all.'),
                    a=r1['short_total'], s='17_R1_vs_R3B',
                    p=('Section 75(12) with Explanation — self-assessed tax includes tax payable on outward supplies '
                       'declared in GSTR-1 but not included in GSTR-3B; recoverable under section 79 without a show '
                       'cause notice. Rule 88C / FORM DRC-01B.'),
                    r='HIGH — directly recoverable'))
            if r1['not_filed']:
                F.append(dict(
                    t=f'GSTR-1 not filed for {len(r1["not_filed"])} tax periods',
                    w=(f'No GSTR-1 is on record for {", ".join(r1["not_filed"])}. GSTR-3B was filed for every one of them. '
                       f'Rs.{inr(r1["not_filed_turnover"])} of turnover and Rs.{inr(r1["not_filed_tax"])} of outward tax '
                       f'therefore sit in GSTR-3B with NO invoice-level record on the portal, and no way for any recipient '
                       f'to have taken credit through GSTR-2B.'),
                    a=0.0, s='17_R1_vs_R3B; 18_RETURN_FILING',
                    p='Section 37 read with Rule 59. Section 47(1) late fee.',
                    r='VERY HIGH — evidentiary' if r1['not_filed_turnover'] > 0 else 'MEDIUM'))
        dm = R['dormancy']
        if dm.get('available') and dm['longest'] and dm['longest'][2] >= 6:
            a, b, c = dm['longest']
            F.append(dict(
                t=f'Credit availed through a {c}-period dormancy',
                w=(f'NIL outward supply in {c} consecutive tax periods, {a} to {b}. ITC continued to be availed — '
                   f'Rs.{inr(dm["itc_in_nil_from_longest"])} in NIL periods from {a} onwards. Across the whole record '
                   f'{dm["nil_months"]} periods are NIL.'),
                a=dm['itc_in_nil_from_longest'], s='14_DORMANCY_ITC',
                p='Section 16(1) — credit admissible only on inputs used or intended to be used in the course or '
                  'furtherance of business. Section 17(5). Rule 86A.',
                r='HIGH'))
        rc = R['rcm']
        if rc.get('available') and (rc['silent_periods'] >= 12 or rc['value_without_tax']):
            amt = sum(v * 0.05 for _, v in rc['value_without_tax'])
            F.append(dict(
                t='Reverse charge liability not declared',
                w=(f'Reverse-charge inward supplies totalling Rs.{inr(rc["declared"])} were declared up to '
                   f'{rc["last_period"]}, with tax of Rs.{inr(rc["tax"])} implied by the returns. '
                   f'No reverse-charge inward supply has been declared in any of the {rc["silent_periods"]} tax periods '
                   f'since, although ITC was availed in most of them.'
                   + (''.join(f' {p} declares RCM inward value of Rs.{inr(v)} with no RCM tax implied at all.'
                              for p, v in rc['value_without_tax']))),
                a=amt, s='16_RCM_CHECK',
                p='Sections 9(3) and 9(4) with Notification 13/2017-CT(R). Section 31(3)(f).',
                r='MEDIUM — quantum to be established'))
        ca = R['cash']
        if ca.get('available'):
            F.append(dict(
                t='Cash actually contributed to the exchequer',
                w=(f'{ca["count"]} challans in all, Rs.{inr(ca["totals"]["TOTAL"])} (SGST Rs.{inr(ca["totals"]["SGST"])}, '
                   f'CGST Rs.{inr(ca["totals"]["CGST"])}, IGST Rs.{inr(ca["totals"]["IGST"])}) against output tax declared '
                   f'of Rs.{inr(ca["output_tax"])}'
                   + (f' — {ca["cash_pct"]*100:.1f}%.' if ca['cash_pct'] is not None else '.')
                   + f' ITC availed over the same period is Rs.{inr(ca["itc"])}'
                   + (f', an availment ratio of {ca["itc"]/ca["output_tax"]*100:.1f}%.' if ca['output_tax'] else '.')),
                a=0.0, s='19_CASH_RECON; 11_FY_SUMMARY', p='Rule 86B. Section 49.', r='HIGH — context'))
        rb = R['rule86b']
        if rb.get('available') and rb['flagged']:
            F.append(dict(
                t='Rule 86B — output tax discharged almost wholly from credit',
                w=('Periods below the 1% cash floor on taxable supply above Rs.50 lakh: '
                   + ', '.join(f'{x["period"]} ({x["pct"]*100:.2f}% on taxable supply of Rs.{inr(x["taxable"])})'
                               for x in rb['rows'] if x['flagged']) + '.'),
                a=0.0, s='15_RULE_86B; 19_CASH_RECON',
                p='Rule 86B, subject to the exceptions in the proviso.',
                r='MEDIUM — prima facie, cash ledger to be verified'))
        fl = R['filing']
        if fl.get('available') and fl['late']:
            F.append(dict(
                t=f'GSTR-3B filed late in {len(fl["late"])} tax periods',
                w=('Total delay ' + str(fl['total_delay']) + ' days: '
                   + ', '.join(f'{p} ({d} days)' for p, _, d in fl['late']) + '.'),
                a=0.0, s='18_RETURN_FILING', p='Section 47(1) and 47(2). Section 50 where tax was paid late.', r='LOW'))
        an = R['annual']
        need = [x for x in an['rows'] if x['gstr9_required'] and not x['on_record']]
        if need:
            F.append(dict(
                t='Annual return and reconciliation statement not on record',
                w=('; '.join(f'FY {x["fy"]} turnover Rs.{inr(x["turnover"])} — GSTR-9 required'
                             + (' and GSTR-9C required' if x['gstr9c_required'] else '') + ', not evidenced'
                             for x in need) + '.'),
                a=0.0, s='SRC_GSTR9; 11_FY_SUMMARY', p='Section 44 read with Rule 80. Section 47(2).',
                r='MEDIUM — compliance'))
        hs = R['hsn']
        if hs.get('available') and len(hs['codes']) >= 3:
            F.append(dict(
                t='Line of business as declared through HSN codes',
                w=('HSN reporting moves ' + ' -> '.join(hs['codes']) + ' over the period. A drift of this kind, and in '
                   'particular goods HSNs appearing in a services taxpayer, points to disposal of capital goods.'),
                a=0.0, s='12_HSN_PROFILE',
                p='Section 18(6) with Rule 44(6). Section 25 and Rule 19 for amendment of registration particulars.',
                r='MEDIUM'))
        F += self._tds_findings()
        F += self._tds_detail_findings()
        rev = [c for c in self.ing.changes if c['level'] == 'REVIEW']
        if rev:
            reps = sorted({c['title'] for c in rev})
            recs = sorted({(c['code'], c['record']) for c in rev})
            F.append(dict(
                t='A later export changed values already held (possible restatement)',
                w=(f'{len(rev)} value(s) across {len(recs)} record(s) in {len(reps)} report(s) differ between an earlier and a '
                   f'later Prime export: ' + '; '.join(f'{c["record"]} {c["field"]} {c["old"]} -> {c["new"]}' for c in rev[:6])
                   + (f'; and {len(rev)-6} more' if len(rev) > 6 else '') + '. The later value is the one used. '
                   + (f'{sum(1 for c in rev if c.get("order_uncertain"))} of these come from files uploaded together, so which is '
                      f'newer cannot be told. ' if any(c.get('order_uncertain') for c in rev) else '')
                   + 'GSTR-3B and payment challans cannot be revised, so a change in either is unusual. It may mean the '
                   'taxpayer amended something, that Prime refreshed its data, or that the earlier export was taken before '
                   'the period was final.'),
                a=0.0, s='09A_RESTATEMENTS',
                p='Verify on the portal before either figure is quoted in a notice. Section 70 / 71 to call for the returns and records.',
                r='MEDIUM — verify'))
        return F

    def _tds_findings(self):
        t = self.R.get('tds') or {}
        if not t.get('available'): return []
        F = []; fys = sorted(t['fy'], key=lambda f: int(f.split('-')[0]))
        rate = 0.18
        F.append(dict(
            t=f'Tax deducted under section 51 on supplies by this taxpayer of at least Rs.{inr(t["total_value"])} ({t["first"]} to {t["last"]})',
            w=(f'GSTR-7 shows tax deducted at source of Rs.{inr(t["total_tds"])} by up to {int(t["deductor_max"])} '
               f'deductors a month. At 2% that is payment for taxable supplies of: '
               + '; '.join(f'FY {f} Rs.{inr(t["fy"][f]["value"])} over {t["fy"][f]["months"]} month(s) '
                           f'({t["fy"][f]["first"]} to {t["fy"][f]["last"]})' for f in fys)
               + '. ' + ('Only CGST and SGST were deducted, so every deductor is in the same State. ' if not t['interstate'] else '')
               + 'This is a floor on turnover: supplies to buyers who do not deduct, and contracts of Rs.2.5 lakh or less, '
                 'carry no TDS.'),
            a=0.0, s='19C_TDS_GSTR7; SRC_R7_TDS',
            p='Section 51 with Notification 50/2018-CT; Rule 66 (GSTR-7); Rule 87(9) (credit to the cash ledger).',
            r='BASELINE — the floor every return is tested against'))
        if not t['have_3b'] or not t['have_r1']:
            miss = ([] if t['have_3b'] else ['GSTR-3B (TAX PAID DETAILS, R3B TAX PAYABLE DETAILS, TURNOVER OUTWARD TAX SUPPLIES, '
                                              'R3B PAYMENT SUMMARY)'])
            miss += [] if t['have_r1'] else ['GSTR-1 (R1 FILINGS IN THE YEAR, HSN LIST)']
            F.append(dict(
                t='The GSTR-7 figures could not yet be tested against the taxpayer\'s own returns',
                w=('The following are not in the case file, so the tests that compare what the deductors paid for with what '
                   'the taxpayer declared were NOT RUN: ' + '; '.join(miss) + '. Upload them under the same GSTIN and run '
                   'again. Sheet 19C lists each test, its provision and the report it needs.'),
                a=0.0, s='19C_TDS_GSTR7; 23_DATA_GAPS', p='—', r='PENDING DATA'))
        if t['no_3b']:
            v = sum(r['value'] for r in t['rows'] if r['period'] in t['no_3b'])
            F.append(dict(
                t=f'No GSTR-3B for {len(t["no_3b"])} month(s) in which the deductors paid for supplies',
                w=(f'TDS was deducted in {", ".join(t["no_3b"])}, proving payment for supplies of Rs.{inr(v)}, but no GSTR-3B '
                   f'is on record for those months.'),
                a=0.0, s='19C_TDS_GSTR7',
                p='Section 39. Section 46 notice in FORM GSTR-3A; section 62 best-judgment assessment on the GSTR-7 value. '
                  'Section 47 late fee, section 50 interest.',
                r='VERY HIGH — non-filer'))
        if t['nil_3b']:
            v = sum(r['value'] for r in t['rows'] if r['period'] in t['nil_3b'])
            F.append(dict(
                t=f'GSTR-3B filed NIL in {len(t["nil_3b"])} month(s) in which the deductors paid for supplies',
                w=(f'{", ".join(t["nil_3b"])}: GSTR-3B declares no outward supply while the deductors paid for supplies of '
                   f'Rs.{inr(v)} and deducted tax on them.'),
                a=0.0, s='19C_TDS_GSTR7; 10_MONTHLY_MASTER',
                p='Sections 39, 73 / 74 / 74A (FY 2024-25 onwards), 50, 122.', r='VERY HIGH'))
        sh = [(f, t['fy'][f]) for f in fys if (t['fy'][f]['short_3b'] or 0) > 1 or (t['fy'][f]['max_cum_short'] or 0) > 1]
        if sh:
            F.append(dict(
                t='Turnover declared in GSTR-3B is below what the deductors alone paid for',
                w=('; '.join(f'FY {f}: the deductors paid for Rs.{inr(a["value"])}, GSTR-3B declares taxable turnover of '
                             f'Rs.{inr(a["r3b_taxable"])} over {a["r3b_months"]} month(s) on record — shortfall Rs.{inr(a["short_3b"])}'
                             + (f', peak cumulative shortfall within the year Rs.{inr(a["max_cum_short"])}'
                                if (a['max_cum_short'] or 0) > (a['short_3b'] or 0) + 1 else '')
                             + f' (tax at 18% would be about Rs.{inr((a["short_3b"] or 0) * rate)} — indicative only, rate to be '
                               f'confirmed from the contracts)' for f, a in sh)
                   + '. Not counted in the total below, because the rate is assumed.'),
                a=0.0, s='19C_TDS_GSTR7 (FY position)',
                p='Sections 12, 13(2), 31 and 39 (time and declaration of supply); 74A for FY 2024-25 onwards; section 50 interest.',
                r='VERY HIGH — suppression on supplies to the deductors'))
        if t['cgst_over']:
            F.append(dict(
                t=f'CGST deducted at 1% exceeds the whole CGST declared, in {len(t["cgst_over"])} month(s)',
                w=f'{", ".join(t["cgst_over"])}. At any rate of 2% or more the tax declared cannot be below the tax deducted at 1%.',
                a=0.0, s='19C_TDS_GSTR7', p='Sections 39 and 51.', r='HIGH'))
        if t['no_r1']:
            F.append(dict(
                t=f'No GSTR-1 for {len(t["no_r1"])} month(s) with TDS payments',
                w=(f'{", ".join(t["no_r1"])}: the B2B invoices to the deductor GSTINs are not on record, although the deductors '
                   f'paid and deducted tax.'),
                a=0.0, s='19C_TDS_GSTR7', p='Section 37, Rule 59; Rule 59(6); section 47.', r='HIGH'))
        r1s = [(f, t['fy'][f]) for f in fys if (t['fy'][f]['short_r1'] or 0) > 1]
        if r1s:
            F.append(dict(
                t='GSTR-1 taxable value below what the deductors paid for',
                w='; '.join(f'FY {f}: Rs.{inr(a["value"])} paid for, GSTR-1 Rs.{inr(a["r1_value"])}, short Rs.{inr(a["short_r1"])}'
                            for f, a in r1s) + '.',
                a=0.0, s='19C_TDS_GSTR7 (FY position)', p='Section 37, Rule 59.', r='HIGH'))
        un = [(f, t['fy'][f]['tds_sgst_unabsorbed']) for f in fys if (t['fy'][f]['tds_sgst_unabsorbed'] or 0) > 1]
        if un:
            F.append(dict(
                t='TDS credit not used to pay tax — a cash-ledger balance building up',
                w=('; '.join(f'FY {f}: SGST deducted exceeds SGST paid in cash by Rs.{inr(v)}' for f, v in un)
                   + '. Output tax is being discharged from ITC while the TDS sits in the cash ledger, where it can be claimed '
                     'as refund.'),
                a=0.0, s='19C_TDS_GSTR7 (FY position)',
                p='Rule 87(9), section 49, section 54 and Rule 89 (refund of excess cash-ledger balance). Rule 86B.',
                r='MEDIUM — refund watch'))
        comp = []
        for f in fys:
            a = t['fy'][f]
            if a['gstr9c']: comp.append(f'FY {f}: Rs.{inr(a["value"])} from the deductors alone — GSTR-9 and self-certified GSTR-9C '
                                        f'required (s.44, Rule 80), e-invoicing from the next FY (Rule 48(4)), QRMP not available')
            elif a['gstr9']: comp.append(f'FY {f}: Rs.{inr(a["value"])} from the deductors alone ({a["first"]} to {a["last"]}) — GSTR-9 '
                                         f'required' + ('; within reach of Rs.5 crore, when GSTR-9C, e-invoicing from the next '
                                                        'FY and the end of QRMP follow' if a['value'] > 0.75 * 50_000_000 else ''))
        over = [p for f in fys for p in t['fy'][f]['over50']]
        if over: comp.append('Rule 86B threshold of Rs.50 lakh crossed on receipts subject to TDS alone in ' + ', '.join(over)
                             + ' — TDS actually used counts as cash, so test on the cash ledger before invoking it')
        if comp:
            F.append(dict(t='Compliance thresholds reached on receipts subject to TDS alone', w='; '.join(comp) + '.',
                          a=0.0, s='19C_TDS_GSTR7', p='Section 44, Rule 80, Rule 48(4), Rule 61A, Rule 86B.',
                          r='MEDIUM — compliance'))
        return F

    def _tds_detail_findings(self):
        d = self.R.get('tds_detail') or {}
        if not d.get('available'): return []
        F = []
        if d['subject_mismatch']:
            F.append(dict(t='GSTR-7 detail export titled for a different GSTIN',
                          w=f'The export names {", ".join(d["subject_mismatch"])}, not the case GSTIN. It may belong to another case.',
                          a=0.0, s='19D_TDS_PARTIES; 09_FILE_REGISTER', p='—', r='CHECK BEFORE RELYING ON IT'))
        if d['ineligible']:
            b = d['ineligible']
            F.append(dict(
                t=f'TDS deducted by {len(b)} part{"y" if len(b)==1 else "ies"} that section 51 does not allow to deduct',
                w=('In the months held (' + ', '.join(m['period'] for m in d['months']) + '), GSTR-7 records tax deducted of '
                   f'Rs.{inr(d["ineligible_tds"])} on supplies of Rs.{inr(d["ineligible_value"])} by: '
                   + '; '.join(f'{", ".join(p["names"])} ({p["gstin"]}) — {p["lines"]} line(s), Rs.{inr(p["value"])}'
                               for p in b)
                   + '. Every one holds a tax-deductor registration (D in position 14) on the PAN of an individual / '
                     'proprietor (P in position 4 of the PAN). Trade names indicate private steel and metal traders. None is a '
                     'Government department, local authority, Governmental agency, Government-controlled body or PSU, so none '
                     'is a deductor under section 51(1) or Notification 50/2018-CT, and the REG-07 registrations are themselves '
                     'irregular. The deductions nevertheless put cash-ledger credit in this taxpayer’s name. '
                   + (f'{", ".join(b[0]["names"])} alone accounts for {d["top_share"]*100:.0f}% of the value. ' if d['top_share'] else '')
                   + (f'{len(d["near_threshold"])} line(s) sit just above the Rs.2.5 lakh deduction threshold. ' if d['near_threshold'] else '')
                   + 'The reason for the arrangement is not on record and must be established.'),
                a=0.0, s='19D_TDS_PARTIES; SRC_R7_DETAIL',
                p=('Section 51(1) with Notification 50/2018-CT (who may deduct); section 24(vi), section 25 and Rule 12 '
                   '(deductor registration) — cancellation under section 29 / Rule 12(3); Rule 87(9) and section 49 (credit '
                   'to the cash ledger); section 54 and Rule 89 (refund of a cash-ledger balance); section 70 (summons).'),
                r='VERY HIGH — lead'))
        bad_m = [m for m in d['months'] if m['agrees'] is False]
        if bad_m:
            F.append(dict(t='GSTR-7 detail does not agree with the summary report',
                          w='; '.join(f'{m["period"]}: detail Rs.{inr(m["tds"])} / {m["parties"]} parties, summary '
                                      f'Rs.{inr(n0(m["s_tds"]))} / {int(n0(m["s_parties"]))}' for m in bad_m) + '.',
                          a=0.0, s='19D_TDS_PARTIES', p='—', r='CHECK THE EXPORTS'))
        return F

    def build_findings(self):
        F = self.findings(); ws = self.sheet('20_FINDINGS')
        ws['A1'] = f'CONSOLIDATED FINDINGS — {self.name}'; ws['A1'].font = BOLD
        ws['A2'] = f'GSTIN {self.gstin}. Ranked by revenue implication. Every figure carries the sheet it is computed on.'
        self.head(ws, 4, ['No.','Finding','What the returns show','Amount involved (Rs.)','Computed on sheet',
                          'Provision attracted','Risk'], [5,46,100,20,28,54,24])
        F = sorted(F, key=_frank)
        for i, f in enumerate(F, 5):
            ws.cell(row=i, column=1, value=i-4); ws.cell(row=i, column=2, value=f['t'])
            ws.cell(row=i, column=3, value=f['w']); ws.cell(row=i, column=4, value=round(f['a']))
            ws.cell(row=i, column=5, value=f['s']); ws.cell(row=i, column=6, value=f['p'])
            ws.cell(row=i, column=7, value=f['r'])
        r = 4 + len(F) + 1
        ws.cell(row=r, column=2, value='TOTAL REVENUE IMPLICATION (excluding interest under section 50 and penalty)').font = BOLD
        ws.cell(row=r, column=4, value=f'=SUM(D5:D{r-1})').font = BOLD
        for rr in range(5, r+1):
            for c in range(1, 8):
                x = ws.cell(row=rr, column=c); x.border = BD
                x.alignment = Alignment(wrap_text=True, vertical='top')
                if c == 4: x.number_format = NUM
        self.total_findings = sum(f['a'] for f in F)
        n = self.note(ws, r+2, f'Rs. {inr(self.total_findings)} before interest and penalty.', bold=True)
        n = self.note(ws, n, 'OVERLAP CAUTION: where the same credit is caught by more than one finding — for example '
                             'credit availed in a NIL period that is also in excess of GSTR-2A — the amount must not be '
                             'demanded twice. Reconcile before the notice is drafted.')
        self.note(ws, n, 'This is a preliminary investigation on returns data. It is not a finding of fact. Every figure is '
                         'to be re-verified against the common portal, the electronic ledgers, the returns themselves and '
                         'the books of account before any notice is issued.')
        return F

    # ------------------------------------- 21 REPORT / 22 ACTIONS / 23 GAPS
    def build_report(self, F):
        R = self.R; ws = self.sheet('21_REPORT')
        fy = R['fy']; fys = sorted(fy, key=lambda x: int(x.split('-')[0]))
        tot_to = sum(v['turnover'] for v in fy.values())
        tot_op = sum(v['op'] for v in fy.values())
        tot_itc = sum(v['itc'] for v in fy.values())
        cash = R['cash']['totals']['TOTAL'] if R['cash'].get('available') else None
        r1 = R['r1_vs_r3b']; fl = R['filing']
        peak = max(fys, key=lambda f: fy[f]['turnover']) if fys else None
        ap = self.ing.all_periods()
        span = (f'{len(R["master"])} tax periods, {R["master"][0]["period"]} to {R["master"][-1]["period"]}' if R['master']
                else f'{ap[0]} to {ap[-1]}' if ap else 'no tax period')
        T = [f'PRELIMINARY INVESTIGATION REPORT — ANALYSIS OF GST RETURNS DATA',
             f'{self.name.upper()}', f'GSTIN {self.gstin}', '',
             '1. SUBJECT', '',
             (f'Analysis of {len(self.ing.counted())} unique GST-Prime Analytica statements covering '
              f'{span}, to identify '
              f'indicators of evasion and to quantify the revenue at stake. '
              + (f'{len(self.ing.register)} files were received; {len(self.ing.duplicates())} were exact duplicates of '
                 f'files already held and have been excluded from every count and computation. '
                 if self.ing.duplicates() else '')
              + 'The register is at sheet 09_FILE_REGISTER.'), '',
             '2. WHAT THE RETURNS SHOW IN OUTLINE', '']
        if not fys:
            T += [('2.1  No GSTR-3B report has been supplied, so the taxpayer\'s own declarations are not yet before this '
                   'analysis. What follows rests on the third-party statements held.'), '']
        else:
          T += [(f'2.1  Across the whole period the unit declared turnover of Rs.{inr(tot_to)}, output tax of '
               f'Rs.{inr(tot_op)} and availed input tax credit of Rs.{inr(tot_itc)}'
               + (f' — an availment ratio of {tot_itc/tot_op*100:.1f} per cent' if tot_op else '')
               + (f'. Against that it has paid Rs.{inr(cash)} in cash, in {R["cash"]["count"]} challans.' if cash is not None else '.')), '']
          T += [(f'2.2  The peak year is FY {peak}, with turnover of Rs.{inr(fy[peak]["turnover"])} and output tax of '
               f'Rs.{inr(fy[peak]["op"])} over {fy[peak]["periods"]} tax periods. '
               + '; '.join(f'FY {f} Rs.{inr(fy[f]["turnover"])}' for f in fys) + '.'), '']
        if R['dormancy'].get('available') and R['dormancy']['longest']:
            a, b, c = R['dormancy']['longest']
            T += [(f'2.3  The longest stretch of NIL outward supply runs {c} consecutive tax periods, {a} to {b}, during '
                   f'which credit of Rs.{inr(R["dormancy"]["itc_in_nil_from_longest"])} was nonetheless availed.'), '']
        if R.get('unfiled'):
            T += [(f'2.4  GSTR-3B is shown as NOT FILED for {", ".join(R["unfiled"])}. That is a period with no return of '
                   f'any kind, and it has been kept out of the tax-period count and out of the NIL-turnover and dormancy '
                   f'tests, so that it is not silently read as a period of nil supply.'), '']
        td = R.get('tds') or {}
        if td.get('available'):
            T += [(f'2.5  GSTR-7 shows that the deductors deducted tax at source of Rs.{inr(td["total_tds"])} from this '
                   f'taxpayer between {td["first"]} and {td["last"]}, which at 2% proves payment for taxable supplies of at least '
                   f'Rs.{inr(td["total_value"])} ('
                   + '; '.join(f'FY {f} Rs.{inr(td["fy"][f]["value"])}' for f in sorted(td['fy'])) + '). '
                   + ('These figures have been tested against GSTR-3B and GSTR-1 at sheet 19C_TDS_GSTR7.'
                      if td['have_3b'] and td['have_r1'] else
                      'GSTR-3B and/or GSTR-1 have not yet been supplied, so the comparison at sheet 19C_TDS_GSTR7 is pending; '
                      'the tests it will run, and the provision behind each, are listed there.')), '']
        dd = R.get('tds_detail') or {}
        if dd.get('available'):
            T += [('2.6  The party-wise GSTR-7 detail is held for ' + ', '.join(m['period'] for m in dd['months'])
                   + f' ({len(dd["parties"])} parties, Rs.{inr(dd["total_value"])} of supplies, TDS Rs.{inr(dd["total_tds"])}). '
                   + dd['role'] + ' '
                   + (f'{len(dd["ineligible"])} of the {len(dd["parties"])} parties are, on the face of their GSTIN and PAN, '
                      f'private proprietors holding tax-deductor registrations — not persons who may deduct under section 51. '
                      if dd['ineligible'] else '')
                   + (f'Detail for {", ".join(dd["missing_months"])} is still to be uploaded.' if dd['missing_months'] else '')), '']
        T += ['3. PRINCIPAL FINDINGS', '']
        for i, f in enumerate(sorted(F, key=_frank), 1):
            T += [f'3.{i}  {f["t"]}.  {f["w"]}'
                  + (f'  This accounts for Rs.{inr(f["a"])} of the revenue at stake.' if f['a'] else ''), '']
        T += ['4. THE REVENUE AT STAKE', '',
              (f'4.1  On the face of the returns the revenue implication is Rs.{inr(self.total_findings)}, before interest '
               f'under section 50 and before penalty. The composition is set out in sheet 20_FINDINGS.'), '',
              '5. LIMITATION', '',
              ('5.1  Sheet 24_LIMITATION sets out the indicative outer dates by financial year, computed from the statute '
               'alone. Extension notifications are not applied there and several have been the subject of litigation, so '
               'every date must be verified against the notifications in force.'), '',
              ('5.2  Where section 73 has run out for a year, a demand for that year can be raised only under section 74, '
               'which requires fraud, wilful misstatement or suppression of facts to be alleged and made out on evidence. '
               'That decision has to be taken before the evidence is gathered, not after.'), '',
              '6. CAVEATS', '',
              ('6.1  This report is built only from the GST-Prime Analytica statements listed at sheet 09_FILE_REGISTER. '
               'It is a preliminary investigation on returns data and is not a finding of fact.'), '',
              ('6.2  Nothing has been estimated, interpolated or assumed. Where a report was not supplied, the test that '
               'needed it says so and was not run. Where a cell was blank in the source it has been kept blank, because a '
               'blank is not a nil.'), '',
              ('6.3  Wherever CGST has been taken as equal to SGST, the assumption is recorded on the sheet concerned and '
               'must be confirmed from the electronic credit ledger.'), '',
              *((('6.4  A later Prime export changed values that an earlier export had already supplied '
                  f'({sum(1 for c in self.ing.changes if c["level"]=="REVIEW")} value(s) needing review). The later value is used '
                  'throughout. Sheet 09A_RESTATEMENTS lists each difference with both values and the files they came from. '
                  'Verify on the portal before relying on either.'), '')
                if any(c['level'] == 'REVIEW' for c in self.ing.changes) else ()),
              ('6.5  Every figure is to be re-verified against the common portal, the electronic credit and cash ledgers, '
               'GSTR-1, GSTR-2A and 2B, GSTR-9 and 9C and the books of account before any notice is issued.')]
        for i, t in enumerate(T, 1):
            c = ws.cell(row=i, column=1, value=t); c.alignment = Alignment(wrap_text=True, vertical='top')
            if t and (t.isupper() or (t[0].isdigit() and '. ' in t[:5]) or t.startswith('GSTIN')): c.font = BOLD
        ws.column_dimensions['A'].width = 145

    def build_actions(self, F):
        ws = self.sheet('22_ACTION_POINTS')
        ws['A1'] = 'PROPOSED COURSE OF ACTION'; ws['A1'].font = BOLD
        A = []
        R = self.R
        if R['credit_notes'].get('available') and R['credit_notes']['events']:
            e = R['credit_notes']['events'][0]
            A.append(['IMMEDIATE', f'Call for the tax invoice, the credit note, the contract and the transport documents '
                                   f'behind the {e["period"]} entry of Rs.{inr(abs(e["turnover"]))}'
                                   + (f' under HSN {", ".join(e["hsn"])}' if e['hsn'] else '') + '.',
                      'Usually the largest single item and the freshest, so the evidence is most likely to still exist and '
                      'the recipient can still be cross-verified.', 'Section 34, section 70'])
            A.append(['IMMEDIATE', 'Examine whether ITC in the electronic credit ledger should be blocked pending verification.',
                      'A credit note that drives output tax negative can create or inflate a credit balance which may be '
                      'utilised or claimed as refund while the enquiry runs.', 'Rule 86A'])
        if R['credit_notes'].get('available') and R['credit_notes']['r1_only']:
            ps = ', '.join(x['period'] for x in R['credit_notes']['r1_only'])
            A.append(['IMMEDIATE', f'Draw the full GSTR-1 for {ps} from the portal, table by table (B2B, B2CS, CDNR, Table 12).',
                      'These returns carry negative tax that does not reconcile with their own taxable value column. They '
                      'cannot be read from a summary.', 'Section 37, section 70, section 151'])
            A.append(['IMMEDIATE', 'Establish whether any ASMT-10, DRC-01A, DRC-03, audit or intelligence proceeding was in '
                                   'progress against this taxpayer in those periods, and obtain the record.',
                      'Large credit reversals and credit notes in one window usually follow departmental contact. If a '
                      'proceeding exists it changes both the evidence available and the limitation position.',
                      'Section 6(2)(b) — no parallel proceeding on the same subject matter'])
        r1 = R['r1_vs_r3b']
        if r1.get('available') and r1['not_filed'] and r1['not_filed_turnover'] > 0:
            A.append(['HIGH', f'Obtain the invoice-level outward record for {", ".join(r1["not_filed"])} from the taxpayer, '
                              f'no GSTR-1 existing on the portal for those periods.',
                      f'Rs.{inr(r1["not_filed_turnover"])} of turnover sits in GSTR-3B with no invoice-level record. '
                      'Identify the recipients and confirm the supplies took place.', 'Section 37, section 70, section 71'])
        if r1.get('available') and r1['short_total'] > 0:
            A.append(['HIGH', f'Issue the Rule 88C intimation in FORM DRC-01B for the GSTR-1 versus GSTR-3B difference of '
                              f'Rs.{inr(r1["short_total"])}.',
                      'Self-assessed tax under the Explanation to section 75(12), recoverable under section 79 without a '
                      'show cause notice and needing no investigation to establish. The fastest realisation in the case.',
                      'Rule 88C, FORM DRC-01B, section 75(12), section 79'])
        if R['excess_itc'].get('available') and R['excess_itc']['net'] > 0:
            A.append(['HIGH', 'Issue the Rule 88D intimation in FORM DRC-01C and obtain the supplier-wise break-up for the '
                              'years carrying the excess.',
                      'The supplier trail has to be established on documents — whether the suppliers are traceable, whether '
                      'they filed, and whether their registrations have since been cancelled.',
                      'Rule 88D, FORM DRC-01C, section 16(2)(aa), Rule 36(4)'])
        if R['dormancy'].get('available') and R['dormancy']['itc_in_nil_from_longest'] > 0:
            A.append(['HIGH', f'Call for the invoices behind the ITC of Rs.{inr(R["dormancy"]["itc_in_nil_from_longest"])} '
                              f'availed in NIL-turnover periods.',
                      'Credit taken against no outward supply must be shown to relate to inputs used or intended to be used '
                      'in business.', 'Section 16(1), section 17(5), section 70'])
        A += self._tds_actions()
        A += [['HIGH', 'Draw the electronic credit and cash ledgers from registration to date, and GSTR-2B for every period.',
               'GSTR-2B and not 2A governs entitlement after 1 January 2022, and only the ledgers settle the CGST '
               'assumptions and confirm the true cash position.', 'Section 151, Rule 36(4), Rule 60'],
              ['HIGH', 'Conduct a physical verification of the declared place of business.',
               'Where a registration has been kept alive through long dormancy, the existence and functioning of the unit '
               'is a question of fact.', 'Section 67, section 71, Rule 25']]
        if R['rcm'].get('available') and R['rcm']['silent_periods'] >= 12:
            A.append(['MEDIUM', 'Call for the goods transport, legal, security and rent expenditure from the books for the '
                                'periods with no reverse-charge declaration, and quantify the liability.',
                      'A long stretch of nil RCM while credit was being availed is not consistent with a working unit.',
                      'Sections 9(3), 9(4), Notification 13/2017-CT(R), section 31(3)(f)'])
        if R['hsn'].get('available'):
            A.append(['MEDIUM', 'Verify the section 18(6) liability on any supply of capital goods, including plant and '
                                'machinery on which credit was taken.',
                      'Tax is the higher of the credit taken less the prescribed reduction, or the tax on transaction value.',
                      'Section 18(6), Rule 44(6)'])
        need = [x for x in R['annual']['rows'] if x['gstr9_required'] and not x['on_record']]
        if need:
            A.append(['MEDIUM', 'Verify the filing of GSTR-9 and GSTR-9C for '
                                + ', '.join(f'FY {x["fy"]}' for x in need) + ', and initiate late fee where not filed.',
                      'GSTR-9C is the reconciliation statement that ties the returns to the audited accounts for the years '
                      'carrying the exposure.', 'Section 44, Rule 80, section 47(2)'])
        if R['rule86b'].get('available') and R['rule86b']['flagged']:
            A.append(['MEDIUM', 'Re-run the Rule 86B test on the electronic cash ledger for '
                                + ', '.join(R['rule86b']['flagged']) + ', and rule out the exceptions in the proviso.',
                      'The present result rests on an SGST-only proxy. The exceptions must be eliminated before the rule '
                      'can be invoked.', 'Rule 86B'])
        A += [['MEDIUM', 'Obtain the audited financial statements, income tax returns, Form 26AS, bank statements and the '
                         'e-way bill data for the whole period.',
               'Turnover per the books against turnover per the returns is the cleanest independent test. E-way bills are '
               'the test of whether the goods moved.', 'Section 71, section 151, Rule 138'],
              ['AS ADVISED', 'Fix the limitation position for each financial year against the notifications in force, and '
                             'settle the section under which each year is to be taken up.',
               'Where section 73 has run out, section 74 requires suppression to be alleged and proved, so the route must '
               'be decided before the evidence is gathered.', 'Sections 73, 74, 74A']]
        PR = {'IMMEDIATE': 0, 'HIGH': 1, 'MEDIUM': 2}
        A = sorted(A, key=lambda a: PR.get(a[0], 3))          # stable: order within a priority is kept
        self.head(ws, 3, ['No.','Priority','Action','Why it comes in this order','Provision / form'], [5,12,78,78,46])
        for i, a in enumerate(A, 4):
            ws.cell(row=i, column=1, value=i-3)
            for j, v in enumerate(a, 2): ws.cell(row=i, column=j, value=v)
        for rr in range(4, 4+len(A)):
            for c in range(1, 6):
                x = ws.cell(row=rr, column=c); x.border = BD
                x.alignment = Alignment(wrap_text=True, vertical='top')

    def _tds_actions(self):
        t = self.R.get('tds') or {}
        if not t.get('available'): return []
        A = []
        if not t['have_3b'] or not t['have_r1']:
            A.append(['IMMEDIATE', 'Export from GST-Prime, for the same GSTIN, the GSTR-3B reports (TAX PAID DETAILS, R3B TAX PAYABLE '
                                   'DETAILS, TURNOVER OUTWARD TAX SUPPLIES, R3B PAYMENT SUMMARY), GSTR-1 (R1 FILINGS, HSN LIST), '
                                   'GSTR-2A (R2A FILINGS, EXCESS ITC CLAIMED) and PAYMENT DETAILS, upload them and run again.',
                      'Every test of the GSTR-7 figures against the taxpayer’s own returns needs them. Until then only the '
                      'value the deductors paid for is established.', 'Section 151, section 70'])
        d = self.R.get('tds_detail') or {}
        if d.get('available') and d['ineligible']:
            names = '; '.join(f'{", ".join(p["names"])} ({p["gstin"]})' for p in d['ineligible'])
            A.append(['IMMEDIATE', f'Confirm on the portal the registration type, constitution, status and jurisdiction of: {names}. '
                                   'Find the ordinary (Z) registration held under the same PAN.',
                      'One search per GSTIN settles that these are tax-deductor registrations held by private proprietors — the '
                      'basis of the finding.', 'Section 25, Rule 12'])
            A.append(['IMMEDIATE', 'Refer the irregular tax-deductor registrations to the officers having jurisdiction over them, '
                                   'for cancellation and for examination of every GSTR-7 they have filed and every deductee they '
                                   'have credited.',
                      'The arrangement is unlikely to be confined to this taxpayer.', 'Section 29(2), Rule 12(3), section 51'])
            A.append(['IMMEDIATE', 'Check whether this taxpayer accepted the TDS credit, how it was used, and whether any refund of '
                                   'the cash-ledger balance has been claimed or sanctioned; hold any pending refund for verification.',
                      'Credit created by deductions that section 51 does not permit is the money at risk.',
                      'Rule 87(9), section 49, section 54, Rule 89'])
            A.append(['HIGH', 'Record statements of the proprietor of this taxpayer and of each listed party: the goods or services '
                              'supplied, the invoices, the payments through bank, and why tax was deducted.',
                      'Nothing in the returns explains why private traders would deduct and deposit 2% in cash.', 'Section 70'])
        if d.get('available') and d['missing_months']:
            A.append(['HIGH', 'Export the "R7 Filing Details For The Month" report for ' + ', '.join(d['missing_months'])
                              + ' and upload it.', 'The party-wise test has so far run on '
                              + str(len(d['months'])) + ' of ' + str(len(d['months']) + len(d['missing_months'])) + ' months.',
                      'Section 151'])
        A.append(['IMMEDIATE', 'Draw the deductor-wise GSTR-7 detail and the GSTR-7A certificates: deductor GSTIN, contract or '
                               'work-order reference, amount paid and date, TDS deducted, month by month.',
                  'The statement used here gives only counts and totals. The deductor-wise detail names each deductor, so the '
                  'supplies can be matched invoice by invoice to GSTR-1.', 'Section 51, Rule 66, Rule 87(9)'])
        if not (d.get('available') and d['parties'] and len(d['ineligible']) == len(d['parties'])):
          A.append(['HIGH', 'Write to the drawing and disbursing officer of each deductor for the work orders, running and final '
                          'bills, measurement books and payment vouchers for the period.',
                  'Independent third-party evidence of the gross contract value, the bill dates (time of supply) and the GST '
                  'charged on each bill — the rate question is settled here.', 'Section 70, section 151'])
        if t['no_3b']:
            A.append(['IMMEDIATE', f'Issue the section 46 notice in FORM GSTR-3A for {", ".join(t["no_3b"])}; if the returns are '
                                   f'not filed within 15 days, assess under section 62 on the GSTR-7 value.',
                      'TDS payments prove supply in months with no return.', 'Sections 46, 62, 47, 50'])
        if any((t['fy'][f]['short_3b'] or 0) > 1 or (t['fy'][f]['max_cum_short'] or 0) > 1 for f in t['fy']) or t['nil_3b']:
            A.append(['HIGH', 'Quantify the shortfall bill by bill from the deductor records, fix the rate from the contract, and '
                              'proceed by DRC-01A intimation and show cause notice.',
                      'The GSTR-7 value is a floor; the bills give the true value and date of each supply.',
                      'Rule 142(1A), sections 73 / 74 / 74A, section 50'])
        A.append(['HIGH', 'Watch for, and verify before sanctioning, any refund claim of the electronic cash-ledger balance.',
                  'TDS credit lands in the cash ledger. If output tax is paid from ITC, the TDS accumulates and can be claimed '
                  'back as refund — before the shortfall is established.', 'Section 54, Rule 89, Rule 87(9)'])
        A.append(['MEDIUM', 'Obtain Form 26AS for the same years and compare the section 194C income-tax TDS by the same '
                            'deductors.',
                  'The same payments are reported to the Income Tax Department. An independent check on the gross amounts paid.',
                  'Section 151'])
        if any(t['fy'][f]['gstr9'] for f in t['fy']):
            A.append(['MEDIUM', 'Calendar the GSTR-9 / GSTR-9C and e-invoicing position for '
                                + ', '.join(f'FY {f}' for f in sorted(t['fy']) if t['fy'][f]['gstr9'])
                                + ', and check whether the taxpayer is on QRMP.',
                      'The thresholds are reached on receipts subject to TDS alone.', 'Section 44, Rule 80, Rule 48(4), Rule 61A'])
        return A

    def build_gaps(self):
        ws = self.sheet('23_DATA_GAPS')
        ws['A1'] = 'DATA POSITION — what is held, and what is still required'; ws['A1'].font = BOLD
        self.head(ws, 3, ['Status','Report','What it drives','Position'], [22, 48, 62, 44])
        r = 4
        for code, sp in REPORT_SPECS.items():
            have = self.ing.have(code)
            ws.cell(row=r, column=1, value='HELD' if have else 'NOT SUPPLIED')
            if not have and self.mark: ws.cell(row=r, column=1).fill = MISSING_FILL
            ws.cell(row=r, column=2, value=f'{sp["title"]}  [{code}]')
            ws.cell(row=r, column=3, value=DRIVES.get(code, ''))
            ws.cell(row=r, column=4, value=(f'{len(self.ing.data[code])} rows from '
                                            f'{len({x["_file"] for x in self.ing.data[code]})} file(s)')
                    if have else 'The tests that need it were not run.')
            r += 1
        for x in EXTERNAL:
            ws.cell(row=r, column=1, value='NOT IN GST-PRIME')
            if self.mark: ws.cell(row=r, column=1).fill = MISSING_FILL
            ws.cell(row=r, column=2, value=x[0]); ws.cell(row=r, column=3, value=x[1])
            ws.cell(row=r, column=4, value=x[2]); r += 1
        for rr in range(4, r):
            for c in range(1, 5):
                y = ws.cell(row=rr, column=c); y.border = BD
                y.alignment = Alignment(wrap_text=True, vertical='top')
        bad = [x for x in self.ing.counted() if x['note']]
        if bad:
            r += 1
            r = self.note(ws, r, 'INTEGRITY NOTES ON THE EXPORTS SUPPLIED', bold=True)
            for x in bad:
                r = self.note(ws, r, f'{x["file"]}: {x["note"]}')

DRIVES = {
 'R3B_TURNOVER':'Turnover, and the reverse-charge inward value used in the RCM test.',
 'R3B_TAX_PAYABLE':'Outward tax at Table 3.1(a), the basis of the GSTR-1 comparison.',
 'TAX_PAID':'Turnover, output tax, ITC availed and cash — the spine of every sheet.',
 'EXCESS_ITC':'The head-wise excess credit computation. THE DEMAND FIGURE.',
 'SGST_SETTLEMENT':'Settlement and gross SGST collected.',
 'HSN':'The HSN profile and the credit-note test against the original supply.',
 'GSTR9':'Whether the annual return and reconciliation statement were filed.',
 'R1_MONTHLY':'GSTR-1 against GSTR-3B, and which periods have no GSTR-1 at all.',
 'R2A_MONTHLY':'The independent credit test, with supplier and invoice counts.',
 'R3B_FILING':'Date of filing, delay in days, late fee and interest exposure.',
 'PAYMENTS':'Cash actually paid, challan by challan.',
 'R3B_FY_SUMMARY':'FY-level cross-check of the ITC figures.',
 'R2A_FY_SUMMARY':'FY-level cross-check of the credit available.',
 'R1_FY_SUMMARY':'FY-level cross-check of the GSTR-1 totals.',
 'R7_DETAIL':'Party-wise GSTR-7 detail for a month: who deducted, on what taxable value; tested against section 51.',
 'R7_TDS':'Tax deducted under section 51: the value of supplies the deductors paid for, tested against GSTR-3B and GSTR-1.',
}
EXTERNAL = [
 ('Deductor-wise GSTR-7 detail and GSTR-7A certificates',
  'Names each deductor, the contract and the payment, so the TDS can be matched to GSTR-1 invoice by invoice.',
  'Draw from the portal / back-office; the Prime statement gives counts and totals only.'),
 ('Supplier-wise (GSTIN-wise) GSTR-2A / 2B detail',
  'Who the suppliers are, whether they exist, whether they filed, whether they have been cancelled. '
  'The section 74 case usually turns on this.', 'Draw from the portal or from GST-Prime supplier reports.'),
 ('GSTR-2B for every period',
  'Entitlement after 1 January 2022 is governed by 2B, not 2A. The DRC-01C intimation must be built on 2B.',
  'Not available in the statements used here.'),
 ('Electronic credit ledger and cash ledger',
  'Confirms the CGST assumptions, the true cash position for Rule 86B, and what a credit note did to the balance.',
  'Draw from the portal.'),
 ('Full GSTR-1, table by table (B2B, B2CS, CDNR, Table 12)',
  'Needed wherever a return-level total does not reconcile with its own taxable value column.', 'Draw from the portal.'),
 ('E-way bill data',
  'Whether the goods actually moved, for high-value supplies and for the goods behind any credit note.',
  'GST-Prime carries an E-Way Bill tab; export separately.'),
 ('Audited accounts, income tax returns, Form 26AS, bank statements',
  'Independent test of turnover, and needed to rule out the Rule 86B exceptions.', 'Call for from the taxpayer.'),
 ('Departmental record — ASMT-10, DRC-01A, DRC-03, audit, intelligence',
  'Explains any unwinding of credit, and determines whether section 6(2)(b) bars a fresh proceeding.',
  'Obtain from the jurisdictional record.'),
]

# ------------------------------------------------------------- index etc.
def _register(b):
    ws = b.sheet('09_FILE_REGISTER')
    ing = b.ing
    ws['A1'] = 'REGISTER OF FILES RECEIVED'; ws['A1'].font = BOLD
    ws['A2'] = ('Duplicates are identified by MD5 checksum of the file itself, not by file name, so a file renamed on '
                're-upload is still caught. A duplicate is excluded from every count and every computation.')
    b.head(ws, 4, ['No.','File as uploaded','Batch','Checksum (MD5)','Report identified','Rows parsed',
                   'Records NEW to the case','Records REPLACED (already held)','Values CHANGED in those records',
                   'Records declared by the report','Status','Note'],
           [5,46,8,36,40,12,15,17,17,16,26,74])
    for i, r in enumerate(ing.register, 5):
        ws.cell(row=i, column=1, value=i-4); ws.cell(row=i, column=2, value=r['file'])
        ws.cell(row=i, column=3, value=r['batch']); ws.cell(row=i, column=4, value=r['md5'])
        ws.cell(row=i, column=5, value=r['title'] or '')
        ws.cell(row=i, column=6, value=r['rows'] or 0)
        ws.cell(row=i, column=7, value=r.get('rows_new', '') if r['status'] == 'COUNTED' else '')
        cc = ws.cell(row=i, column=8, value=r.get('rows_replaced', '') if r['status'] == 'COUNTED' else '')
        if r.get('rows_replaced') and b.mark: cc.fill = MISSING_FILL
        vc = ws.cell(row=i, column=9, value=r.get('values_changed', '') if r['status'] == 'COUNTED' else '')
        if r.get('values_changed') and b.mark: vc.fill = MISSING_FILL
        ws.cell(row=i, column=10, value=r['declared'] if r['declared'] is not None else '')
        c = ws.cell(row=i, column=11, value=r['status'])
        if r['status'] != 'COUNTED' and b.mark: c.fill = MISSING_FILL
        ws.cell(row=i, column=12, value=r['note'] or r['detect_note'] or '')
    last = 4 + len(ing.register)
    for rr in range(5, last+1):
        for c in range(1, 13):
            x = ws.cell(row=rr, column=c); x.border = BD
            x.alignment = Alignment(wrap_text=True, vertical='top')
    ws.cell(row=last+2, column=2, value='Files received').font = BOLD
    ws.cell(row=last+2, column=6, value=len(ing.register)).font = BOLD
    ws.cell(row=last+3, column=2, value='Exact duplicates excluded (byte-identical)').font = BOLD
    ws.cell(row=last+3, column=6, value=f'=COUNTIF(K5:K{last},"DUPLICATE*")').font = BOLD
    ws.cell(row=last+4, column=2, value='Not recognised or unreadable').font = BOLD
    ws.cell(row=last+4, column=6, value=f'=COUNTIF(K5:K{last},"NOT RECOGNISED")+COUNTIF(K5:K{last},"UNREADABLE")').font = BOLD
    ws.cell(row=last+5, column=2, value='Unique reports relied on').font = BOLD
    ws.cell(row=last+5, column=6, value=f'=F{last+2}-F{last+3}-F{last+4}').font = BOLD
    ws.cell(row=last+6, column=2, value='Records REPLACED by a later export (not added twice)').font = BOLD
    ws.cell(row=last+6, column=6, value=f'=SUM(H5:H{last})').font = BOLD
    ws.cell(row=last+7, column=2, value='Values that DIFFER from the earlier export — see 09A_RESTATEMENTS').font = BOLD
    ws.cell(row=last+7, column=6, value=f'=SUM(I5:I{last})').font = BOLD
    ws.freeze_panes = 'A5'
    n = last + 9
    b.note(ws, n, 'A file is rejected as a DUPLICATE only when it is byte-identical to one already held. '
                  'A later export of the same report that carries the same records plus new ones is NOT a '
                  'duplicate — it is accepted, and the records already held are REPLACED rather than added '
                  'a second time. Column H counts those. Without that, a re-export would roughly double '
                  'every total built on the overlapping records.', bold=True)

def _derivation(b):
    ws = b.sheet('08_DERIVATION_LOGIC')
    ws['A1'] = 'DERIVATION LOGIC — every derived column, its formula, its source and the reason for deriving it'
    ws['A1'].font = BOLD
    b.head(ws, 3, ['Sheet','Derived column','Formula logic','Source read','Reason / legal basis'], [24,28,46,28,86])
    for i, row in enumerate(DERIV, 4):
        for j, v in enumerate(row, 1):
            c = ws.cell(row=i, column=j, value=Builder._safe(v)); c.border = BD
            c.alignment = Alignment(wrap_text=True, vertical='top')

DERIV = [
 ['10_MONTHLY_MASTER','Every figure','INDEX/MATCH on Return Period into the SRC_ sheet named in column D',
  'SRC_TAX_PAID, SRC_R3B_TAX_PAYABLE, SRC_R3B_TURNOVER, SRC_EXCESS_ITC',
  'Nothing is retyped. Each cell is a live lookup, so a value can be traced to the exact source row it came from.'],
 ['10_MONTHLY_MASTER','Implied RCM / other liability','[GST OP Total] minus [Outward total tax 3.1(a)]',
  'SRC_TAX_PAID; SRC_R3B_TAX_PAYABLE',
  'GST OP Total includes the Table 3.1(d) reverse-charge liability while Table 3.1(a) does not. The difference isolates '
  'RCM so it can be tested on its own under section 9(3) and 9(4).'],
 ['10_MONTHLY_MASTER','NIL turnover / ITC availed in NIL period','IF(turnover = 0, "NIL", "") and IF(NIL, ITC, 0)',
  'Same sheet','Isolates credit availed with no corresponding business activity. Section 16(1).'],
 ['10_MONTHLY_MASTER','GSTR-2A data availability','Set where the 2A column was BLANK, not zero, in the source',
  'SRC_EXCESS_ITC','A blank is not a nil. Treating an unpopulated 2A as zero credit would overstate the excess.'],
 ['11_FY_SUMMARY','All columns','SUMIFS / COUNTIFS over 10_MONTHLY_MASTER by financial year','10_MONTHLY_MASTER',
  'Demands, annual returns and limitation all run financial-year wise.'],
 ['13_EXCESS_ITC','Excess claimed','Taken as reported by GST-Prime, head by head; CGST mirrored from SGST',
  'SRC_EXCESS_ITC','Section 16(2)(aa) and Rule 36(4) restrict credit to what the supplier reported. Heads are kept '
  'separate because an excess in one head is not set off against a shortfall in another for the purpose of a demand.'],
 ['13_EXCESS_ITC','Net versus gross','Net sums every period; gross counts only periods of over-claim','Same sheet',
  'A short claim followed by a larger claim is ordinarily only timing. Net is the defensible exposure.'],
 ['17_R1_vs_R3B','GSTR-1 filing status','Three-way test: period absent from the report / row present but blank / filed',
  'SRC_R1_MONTHLY','A period that does not appear, or appears blank, has no GSTR-1 on record. Distinguishing absent from '
  'filed-nil is what turns a blank into a section 37 contravention.'],
 ['17_R1_vs_R3B','Difference in tax','[GSTR-1 tax] minus [GSTR-3B Table 3.1(a) tax]','SRC_R1_MONTHLY; 10_MONTHLY_MASTER',
  'Compared against 3.1(a) and not against OP Total, because OP Total carries the RCM liability which has no counterpart '
  'in GSTR-1. A positive difference is self-assessed tax under the Explanation to section 75(12).'],
 ['19B_CREDIT_NOTES','Positive supply ever declared under the same HSN',
  'Every negative outward declaration is matched against every positive HSN row in the whole record',
  'SRC_HSN; 10_MONTHLY_MASTER',
  'Section 34(1) requires a credit note to be referable to an identified original tax invoice. If the taxpayer has never '
  'declared a positive supply under that HSN, there is nothing for the credit note to be referable to.'],
 ['19B_CREDIT_NOTES','Internally inconsistent',
  'Flagged where GSTR-1 tax is negative while GSTR-1 taxable value is not','SRC_R1_MONTHLY',
  'The two columns of the same return contradict each other. The return cannot be read from a summary and must be drawn '
  'table by table from the portal.'],
 ['18_RETURN_FILING','Statutory due date','20th of the month following the tax period','Constant, not from the source data',
  'The monthly filer due date. If the taxpayer was on QRMP in any period, that period must be re-set before late fee is computed.'],
 ['18_RETURN_FILING','Delay in days','MAX(0, [date of filing] minus [due date])','SRC_R3B_FILING',
  'Late fee under section 47(1) and interest under section 50 both run on days of delay.'],
 ['19_CASH_RECON','Cash as a proportion of output tax','[total cash paid] divided by [total output tax declared]',
  '19_CASH_RECON; 11_FY_SUMMARY','The single ratio that describes the case: what was declared against what was parted with.'],
 ['19A_ITC_vs_2A','GAP','[ITC availed] minus [2A credit available], suppressed where no 2A report exists',
  'SRC_R2A_MONTHLY; 10_MONTHLY_MASTER',
  'Built from reports that do not feed sheet 13, so the excess credit finding rests on two independent sources.'],
 ['19A_ITC_vs_2A','Do the two routes agree?','IF(ABS((monthly gap + untestable ITC) - annual gap) < 100, "yes", "CHECK")',
  'Same sheet','A live consistency test between the month-wise files and the annual summary. If a later upload breaks the '
  'agreement the cell says so.'],
 ['15_RULE_86B','Cash proxy','[SGST cash paid] x 2','SRC_TAX_PAID',
  'Only the SGST head of cash is reported. CGST cash is taken as equal; IGST cash is not captured. Prima facie only — '
  'cross-check against the challans in sheet 19.'],
 ['19C_TDS_GSTR7','Value paid for by the deductors','([Total GST] minus [Cess]) / 2%','SRC_R7_TDS',
  'Section 51: 1% CGST + 1% SGST, or 2% IGST, of the value of supply excluding the tax on the invoice. Total GST is used '
  'because Prime rounds each head but totals the unrounded heads. Cess is not deducted under section 51.'],
 ['19C_TDS_GSTR7','Cumulative value / cumulative GSTR-3B turnover','SUMIF of the rows so far in the same FY; SUM of the '
  'master rows from the first period of the FY to this period','This sheet; 10_MONTHLY_MASTER',
  'Payment month is not invoice month. By s.13(2) (services) and s.31(1) (goods) a supply paid for by a month must have been '
  'declared by that month, so the cumulative position is the firm test; a single month can be timing.'],
 ['19C_TDS_GSTR7','SGST TDS not absorbed','MAX(0, SGST deducted - SGST paid in cash in GSTR-3B), by FY',
  'This sheet; 10_MONTHLY_MASTER','TDS is credited to the cash ledger (Rule 87(9)). If it is not used, a balance builds up '
  'that can be claimed as refund under section 54.'],
 ['19C_TDS_GSTR7','Indicative tax','Shortfall x the rate typed in H3','This sheet',
  'The only assumed figure in the workbook, kept on this sheet and out of 20_FINDINGS.'],
 ['19D_TDS_PARTIES','Registration type; PAN holder type','Position 14 of the GSTIN (D = tax-deductor registration, Z = '
  'ordinary); position 4 of the PAN (P individual, F firm, C company, G Government, L local authority ...)','SRC_R7_DETAIL',
  'Read off the GSTIN itself. Section 51(1) and Notification 50/2018-CT limit deductors to Government, local authorities, '
  'Governmental agencies, Government-controlled bodies and PSUs. To be confirmed on the portal.'],
 ['19D_TDS_PARTIES','Lines / taxable value / TDS','COUNTIF / SUMIF of SRC_R7_DETAIL by GSTIN and by month','SRC_R7_DETAIL',
  'Live totals over the source rows.'],
 ['19D_TDS_PARTIES','Agrees?','TDS difference within Re.1 plus half a rupee per line, and the party count equal to the summary\'s '
  '"No Of Deductors"','SRC_R7_DETAIL; SRC_R7_TDS','The detail heads are rounded per line, so a few rupees of rounding is expected.'],
 ['SRC_R7_DETAIL','Return Period','Taken from the FILE NAME (..._Month_Mar-2026.xlsx)','File name',
  'The export carries no period inside it. The register says so against every such file.'],
 ['24_LIMITATION','All dates','Annual return due 31 December following the FY; section 73 order +3 years, notice 3 months '
  'earlier; section 74 order +5 years, notice 6 months earlier','Computed from the statute',
  'Indicative only. Extension notifications are not applied and several have been litigated.'],
 ['SRC_ sheets','Record history column; amber row; orange cell',
  'History text on every record replaced by a later export; amber row and orange cell only where a value differs',
  'The engine, at ingest','A trail inside the workbook, not only in the register. Hover over an orange cell for the value it replaced.'],
 ['09A_RESTATEMENTS','Difference (later minus earlier)','IF both values are numbers, later minus earlier',
  'Same sheet','Shows the size of the movement. SNo, Growth and Trend columns are ignored: SNo is only the row position in an '
  'export and the other two are derived. A difference under Re.1 is treated as rounding.'],
 ['09A_RESTATEMENTS','Level (REVIEW / NOTE / INFO)',
  'INFO for GSTR-2A, which moves whenever a supplier files or amends; NOTE for blank-to-populated; REVIEW otherwise',
  'Report type and field','GSTR-2A changing is expected and is not a restatement by the taxpayer. GSTR-3B and challans cannot be '
  'revised, so a change there is the one to look at.'],
 ['09_FILE_REGISTER','Unique reports relied on','[files received] minus duplicates minus unrecognised','Same sheet',
  'Duplicates are identified by MD5 of the file, so a renamed re-upload is still caught.'],
]

def _index(b, version, run_dt):
    ws = b.sheet('00_INDEX')
    rows = [['GST-PRIME PRELIMINARY INVESTIGATION', '', ''],
            ['Taxpayer', b.name, ''],
            ['GSTIN', b.gstin, '']]
    for k, v in (b.profile or {}).items():
        rows.append([k, v, ''])
    m = b.R['master']
    ap = b.ing.all_periods()
    cover = (f'{m[0]["period"]} to {m[-1]["period"]} ({len(m)} tax periods with GSTR-3B)' if m else
             (f'{ap[0]} to {ap[-1]} — NO GSTR-3B REPORT SUPPLIED; only ' +
              ', '.join(REPORT_SPECS[c]['title'] for c in REPORT_SPECS if b.ing.have(c)) if ap else 'no periods'))
    rows += [['Period covered', cover, ''],
             ['Periods with no GSTR-3B', ', '.join(b.R.get('unfiled') or []) or 'none'
                 + ('  (kept out of the period count and the NIL tests)' if b.R.get('unfiled') else ''), ''],
             ['Reports relied on', f'{len(b.ing.counted())} unique GST-Prime statements '
                                   f'({len(b.ing.register)} files received, {len(b.ing.duplicates())} exact duplicates '
                                   f'excluded) — see 09_FILE_REGISTER', ''],
             ['Revenue implication', f'Rs. {inr(b.total_findings)} before interest and penalty — see 20_FINDINGS', ''],
             ['', '', ''],
             ['HOW THIS WORKBOOK IS PUT TOGETHER', '', ''],
             ['SRC_ sheets', 'Every GST-Prime report, reproduced exactly as received. No value altered. A cell shaded '
                             'light red was BLANK in the source, which is not the same as nil.', ''],
             ['08_DERIVATION_LOGIC', 'For every derived column: the formula, the source it reads, and the reason for '
                                     'deriving it.', ''],
             ['09_FILE_REGISTER', 'Every file received, with checksum, what it was identified as, and whether it was counted.', ''],
             ['10 to 19D', 'The computation sheets. Every figure is a live Excel formula reading an SRC_ sheet or the '
                           'monthly master — nothing is retyped.', ''],
             ['20 to 24', 'Findings, report, action points, data position and limitation.', ''],
             ['', '', ''],
             ['SHEET INDEX', '', '']]
    order = [s for s in b.wb.sheetnames]
    for s in order:
        rows.append([s, SHEET_NOTE.get(s, 'Source data, reproduced as received.' if s.startswith('SRC_') else ''), ''])
    rows += [['', '', ''],
             ['VERSION', '', ''],
             ['Engine version', version, ''],
             ['Run at', run_dt.strftime('%d-%m-%Y %H:%M'), ''],
             ['', '', ''],
             ['SCOPE', '', ''],
             ['This is a PRELIMINARY INVESTIGATION built only from GST-Prime Analytica statements. It is an analysis of '
              'returns data and is not a finding of fact.', '', ''],
             ['Nothing has been estimated, interpolated or assumed. Where a report was not supplied, the test that needed '
              'it says so and was not run.', '', ''],
             ['Every figure is to be re-verified against the GST common portal, the electronic ledgers, the returns '
              'themselves and the books of account before any notice is issued.', '', '']]
    for i, r in enumerate(rows, 1):
        for j, v in enumerate(r, 1):
            if v:
                c = ws.cell(row=i, column=j, value=v); c.alignment = Alignment(wrap_text=True, vertical='top')
    for i, r in enumerate(rows, 1):
        if r[0] and r[0].isupper() and not r[1]: ws.cell(row=i, column=1).font = BOLD
    ws.cell(row=1, column=1).font = BOLD
    for i, w in enumerate([32, 92, 20], 1): ws.column_dimensions[get_column_letter(i)].width = w

SHEET_NOTE = {
 '00_INDEX':'This sheet.',
 '08_DERIVATION_LOGIC':'Every derived column explained.',
 '09_FILE_REGISTER':'Files received, checksums, duplicates.',
 '09A_RESTATEMENTS':'What a later export changed from an earlier one.',
 '10_MONTHLY_MASTER':'One row per tax period; every cell a live lookup.',
 '11_FY_SUMMARY':'Financial-year roll-up.',
 '12_HSN_PROFILE':'What the taxpayer says it supplies, in order.',
 '13_EXCESS_ITC':'ITC against GSTR-2A, head by head. THE DEMAND FIGURE.',
 '14_DORMANCY_ITC':'Credit availed in NIL-turnover periods.',
 '15_RULE_86B':'Cash component of the output tax discharge.',
 '16_RCM_CHECK':'Reverse charge declared and implied.',
 '17_R1_vs_R3B':'GSTR-1 against GSTR-3B, every period, with filing status.',
 '18_RETURN_FILING':'Filing dates against due dates; delay in days.',
 '19_CASH_RECON':'Every challan, and cash against declared liability.',
 '19A_ITC_vs_2A':'The credit test re-run from independent reports.',
 '19B_CREDIT_NOTES':'Credit notes tested against the original supply.',
 '19D_TDS_PARTIES':'GSTR-7 party by party: who deducted, whether section 51 allows them to, and the month reconciliation.',
 '19C_TDS_GSTR7':'What the deductors paid for (GSTR-7 TDS / 2%) against GSTR-3B and GSTR-1; the legal cross-checks.',
 '20_FINDINGS':'Consolidated findings with quantification.',
 '21_REPORT':'The report in narrative form.',
 '22_ACTION_POINTS':'What to do, in the order to do it.',
 '23_DATA_GAPS':'What is held and what is still required.',
 '24_LIMITATION':'Indicative outer dates by financial year.',
}

ORDER_HINT = ['00_INDEX'] + [f'SRC_{c}'[:31] for c in REPORT_SPECS] + \
             ['08_DERIVATION_LOGIC','09_FILE_REGISTER','09A_RESTATEMENTS','10_MONTHLY_MASTER','11_FY_SUMMARY','12_HSN_PROFILE',
              '13_EXCESS_ITC','14_DORMANCY_ITC','15_RULE_86B','16_RCM_CHECK','17_R1_vs_R3B','18_RETURN_FILING',
              '19_CASH_RECON','19A_ITC_vs_2A','19B_CREDIT_NOTES','19C_TDS_GSTR7','19D_TDS_PARTIES','20_FINDINGS','21_REPORT','22_ACTION_POINTS',
              '23_DATA_GAPS','24_LIMITATION']

def build_workbook(ing, R, gstin, name, out_path, mark_missing_light_red=True, profile=None):
    b = Builder(ing, R, gstin, name, mark_missing_light_red, profile)
    b.build_sources()
    b.build_master(); b.build_fy()
    b.build_hsn(); b.build_excess(); b.build_dormancy(); b.build_86b(); b.build_rcm()
    b.build_r1(); b.build_filing(); b.build_cash(); b.build_itc_2a(); b.build_credit_notes(); b.build_tds(); b.build_tds_parties()
    F = b.build_findings(); b.build_report(F); b.build_actions(F); b.build_gaps(); b.build_limitation()
    b.build_restatements(); _derivation(b); _register(b)
    _index(b, BUILD_VERSION, datetime.datetime.now())
    names = [s for s in ORDER_HINT if s in b.wb.sheetnames] + \
            [s for s in b.wb.sheetnames if s not in ORDER_HINT]
    b.wb._sheets = [b.wb[s] for s in names]
    for s in b.wb.worksheets:
        s.page_setup.orientation = 'landscape'; s.page_setup.fitToWidth = 1
        s.sheet_properties.pageSetUpPr.fitToPage = True
    b.wb.active = 0
    b.wb.save(out_path)
    return out_path, b.total_findings
