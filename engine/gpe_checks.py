# -*- coding: utf-8 -*-
"""
GST-PRIME PRELIMINARY INVESTIGATION ENGINE  —  checks
Each check states which reports it needs. If any is absent the check returns
available=False and names what is missing, instead of inventing a result.
"""
import datetime, re
from gpe_core import (REPORT_SPECS, parse_period, period_key, fy_of, fy_long,
                      n0, to_date, inr)

CHECKS_VERSION = "1.2"
FY_ORDER = ['2021-22','2022-23','2023-24','2024-25','2025-26','2026-27',
            '2027-28','2028-29','2029-30']

def _fys(periods):
    seen = []
    for p in periods:
        f = fy_of(p)
        if f and f not in seen: seen.append(f)
    return sorted(seen, key=lambda f: int(f.split('-')[0]))

def _need(ing, *codes):
    miss = [c for c in codes if not ing.have(c)]
    return miss

def due_date(p):
    _, y, m = parse_period(p)
    m += 1
    if m > 12: m, y = 1, y + 1
    return datetime.date(y, m, 20)

# ------------------------------------------------------------------ master
def unfiled_periods(ing):
    """Periods for which GSTR-3B itself is shown as not filed (typically the month of
    registration). They are not tax periods with a return, so they are kept OUT of the
    master, out of every period count, and out of the NIL / dormancy tests. Otherwise a
    period with no return at all is silently counted as a nil-turnover period."""
    tp = ing.by_period('TAX_PAID')
    return sorted([p for p, r in tp.items()
                   if (r.get('Filing Status') or '').strip().lower().startswith('not')],
                  key=period_key)

def build_master(ing):
    """One row per tax period, every figure traced to its source report."""
    skip = set(unfiled_periods(ing))
    per = [p for p in ing.periods() if p not in skip]
    tp, tx, to, it, se = (ing.by_period(c) for c in
        ('TAX_PAID','R3B_TAX_PAYABLE','R3B_TURNOVER','EXCESS_ITC','SGST_SETTLEMENT'))
    rows = []
    for p in per:
        a, b, c, d, e = tp.get(p,{}), tx.get(p,{}), to.get(p,{}), it.get(p,{}), se.get(p,{})
        r2a_blank = ('SGSTITC Available In R2A' in d.get('_blank_cols', []) or
                     'IGSTITC Available In R2A' in d.get('_blank_cols', []))
        rows.append(dict(
            period=p, fy=fy_of(p),
            filing_status=a.get('Filing Status'),
            turnover=n0(a.get('Total TO')), taxable=n0(a.get('Taxable TO')),
            out_sgst=n0(b.get('Outward SGST')), out_cgst=n0(b.get('Outward CGST')),
            out_igst=n0(b.get('Outward IGST')), out_total=n0(b.get('Outward Total GST')),
            op_total=n0(a.get('GST OP Total')),
            rcm_value=n0(c.get('Inward Supplies(Reverse Charge)')),
            itc=n0(a.get('GST ITC Total')),
            sgst_claim=n0(d.get('SGSTITC Claimed In R3B')), sgst_avail=n0(d.get('SGSTITC Available In R2A')),
            sgst_excess=n0(d.get('SGSTExcess Claimed')),
            igst_claim=n0(d.get('IGSTITC Claimed In R3B')), igst_avail=n0(d.get('IGSTITC Available In R2A')),
            igst_excess=n0(d.get('IGSTExcess Claimed')),
            cash_sgst=n0(a.get('SGST Cash Paid')), settlement=n0(a.get('SGST Settlement')),
            gross_sgst=n0(a.get('Gross SGST Collected')),
            r2a_not_available=r2a_blank,
        ))
    for r in rows:
        r['rcm_tax'] = r['op_total'] - r['out_total']
        r['nil'] = (r['turnover'] == 0)
        r['itc_in_nil'] = r['itc'] if r['nil'] else 0.0
    return rows

# ------------------------------------------------------- 1. GSTR-1 vs 3B
def chk_r1_vs_r3b(ing, master):
    miss = _need(ing, 'R1_MONTHLY', 'R3B_TAX_PAYABLE', 'TAX_PAID')
    if miss: return dict(available=False, missing=miss)
    r1 = ing.by_period('R1_MONTHLY')
    rows, short, notfiled, noreturn = [], 0.0, [], []
    nf_to = nf_tax = 0.0
    for m in master:
        p = m['period']; row = r1.get(p)
        r3b_absent = (m['filing_status'] or '').strip().lower().startswith('not')
        if r3b_absent:
            # No GSTR-3B either: this is a period with no return of any kind,
            # not a GSTR-1 gap in a month that was otherwise reported.
            st, v, t = 'NO RETURN OF ANY KIND', None, None
            noreturn.append(p)
        elif row is None:
            st, v, t = 'NOT FILED - no return on record', None, None
        elif row.get('Total GST') is None:
            st, v, t = 'NOT FILED - row blank', None, None
        elif n0(row.get('Total GST')) == 0 and n0(row.get('Taxable Value')) == 0:
            st, v, t = 'FILED - NIL', 0.0, 0.0
        else:
            st, v, t = 'FILED - WITH DATA', n0(row.get('Taxable Value')), n0(row.get('Total GST'))
        if st.startswith('NOT FILED'):
            notfiled.append(p); nf_to += m['turnover']; nf_tax += m['out_total']
        d = None if t is None else t - m['out_total']
        if d and d > 1: short += d
        rows.append(dict(period=p, fy=m['fy'], status=st, r1_value=v, r3b_value=m['turnover'],
                         r1_tax=t, r3b_tax=m['out_total'], diff=d,
                         short=max(0.0, d) if d else 0.0,
                         sellers=None if row is None else row.get('No Of Sellers'),
                         invoices=None if row is None else row.get('No Of Invoices')))
    return dict(available=True, rows=rows, short_total=short, not_filed=notfiled,
                not_filed_turnover=nf_to, not_filed_tax=nf_tax, no_return=noreturn)

# ------------------------------------------------- 2. excess ITC (demand)
def chk_excess_itc(ing, master):
    miss = _need(ing, 'EXCESS_ITC')
    if miss: return dict(available=False, missing=miss)
    rows, net, gross = [], 0.0, 0.0
    for m in master:
        s, i = m['sgst_excess'], m['igst_excess']
        tot = s * 2 + i                      # CGST mirrors SGST on intra-State supply
        pos = max(0.0, s) * 2 + max(0.0, i)
        net += tot; gross += pos
        rows.append(dict(period=m['period'], fy=m['fy'],
                         sgst_claim=m['sgst_claim'], sgst_avail=m['sgst_avail'], sgst_excess=s,
                         cgst_excess=s, igst_claim=m['igst_claim'], igst_avail=m['igst_avail'],
                         igst_excess=i, total=tot, positive=pos,
                         r2a_not_available=m['r2a_not_available']))
    return dict(available=True, rows=rows, net=net, gross=gross,
                months_excess=sum(1 for r in rows if r['sgst_excess'] > 0.5 or r['igst_excess'] > 0.5),
                months_excess_sgst=sum(1 for r in rows if r['sgst_excess'] > 0.5),
                months_excess_igst=sum(1 for r in rows if r['igst_excess'] > 0.5))

# --------------------------------------- 3. ITC vs 2A, independent route
def chk_itc_vs_2a(ing, master):
    miss = _need(ing, 'R2A_MONTHLY', 'TAX_PAID')
    if miss: return dict(available=False, missing=miss)
    r2 = ing.by_period('R2A_MONTHLY')
    rows, gap, untested = [], 0.0, 0.0
    for m in master:
        a = r2.get(m['period'])
        if a is None:
            rows.append(dict(period=m['period'], fy=m['fy'], present=False, sellers=None,
                             invoices=None, taxable=None, avail=None, itc=m['itc'], gap=None))
            untested += m['itc']; continue
        av = n0(a.get('Total GST')); g = m['itc'] - av; gap += g
        rows.append(dict(period=m['period'], fy=m['fy'], present=True,
                         sellers=a.get('Seller Count'), invoices=a.get('Invoices Count'),
                         taxable=a.get('Taxable Value'), avail=av, itc=m['itc'], gap=g))
    ann = {}
    if ing.have('R2A_FY_SUMMARY') and ing.have('R3B_FY_SUMMARY'):
        f2, f3 = ing.by_fy('R2A_FY_SUMMARY'), ing.by_fy('R3B_FY_SUMMARY')
        for f in set(f2) | set(f3):
            ann[f] = dict(avail=n0(f2.get(f, {}).get('Total GST')),
                          itc=n0(f3.get(f, {}).get('GST Input Tax Credit')))
            ann[f]['gap'] = ann[f]['itc'] - ann[f]['avail']
    return dict(available=True, rows=rows, gap=gap, untested_itc=untested, annual=ann)

# ------------------------------------------------------------ 4. dormancy
def chk_dormancy(ing, master):
    if not ing.have('TAX_PAID'): return dict(available=False, missing=['TAX_PAID'])
    run = 0; runs = []; cur = []
    for m in master:
        if m['nil']:
            cur.append(m['period']); run += 1
        else:
            if len(cur) >= 3: runs.append((cur[0], cur[-1], len(cur)))
            cur, run = [], 0
    if len(cur) >= 3: runs.append((cur[0], cur[-1], len(cur)))
    longest = max(runs, key=lambda x: x[2]) if runs else None
    itc_nil = sum(m['itc_in_nil'] for m in master)
    after = 0.0
    if longest:
        k = period_key(longest[0])
        after = sum(m['itc_in_nil'] for m in master if period_key(m['period']) >= k)
    return dict(available=True, runs=runs, longest=longest,
                itc_in_nil_all=itc_nil, itc_in_nil_from_longest=after,
                nil_months=sum(1 for m in master if m['nil']))

# ----------------------------------------------------------- 5. Rule 86B
def chk_rule86b(ing, master, threshold=5_000_000, floor=0.01):
    if not ing.have('TAX_PAID'): return dict(available=False, missing=['TAX_PAID'])
    rows, flagged = [], []
    for m in master:
        cash = m['cash_sgst'] * 2            # CGST cash taken as equal; IGST cash not captured
        pct = (cash / m['op_total']) if m['op_total'] > 0 else None
        hit = (m['taxable'] > threshold and m['op_total'] > 0 and pct is not None and pct < floor)
        if hit: flagged.append(m['period'])
        rows.append(dict(period=m['period'], taxable=m['taxable'], op=m['op_total'],
                         cash_sgst=m['cash_sgst'], cash_proxy=cash, pct=pct,
                         applicable=m['taxable'] > threshold, flagged=hit))
    return dict(available=True, rows=rows, flagged=flagged)

# ----------------------------------------------------------------- 6. RCM
def chk_rcm(ing, master):
    miss = _need(ing, 'R3B_TURNOVER', 'R3B_TAX_PAYABLE', 'TAX_PAID')
    if miss: return dict(available=False, missing=miss)
    rows = []; declared = 0.0; tax = 0.0; last = None; gap_periods = []
    for m in master:
        rate = (m['rcm_tax'] / m['rcm_value']) if m['rcm_value'] else None
        declared += m['rcm_value']; tax += m['rcm_tax']
        if m['rcm_value'] > 0: last = m['period']
        if m['rcm_value'] > 0 and abs(m['rcm_tax']) < 1:
            gap_periods.append((m['period'], m['rcm_value']))
        rows.append(dict(period=m['period'], value=m['rcm_value'], op=m['op_total'],
                         out=m['out_total'], tax=m['rcm_tax'], rate=rate))
    since = 0
    if last:
        k = period_key(last)
        since = sum(1 for m in master if period_key(m['period']) > k)
    return dict(available=True, rows=rows, declared=declared, tax=tax,
                last_period=last, silent_periods=since, value_without_tax=gap_periods)

# ------------------------- 7. credit notes with no original supply (NEW)
def chk_credit_notes(ing, master):
    """A negative outward declaration is tested against every positive supply
    ever declared under the same HSN, and against total positive turnover."""
    if not ing.have('R3B_TAX_PAYABLE'): return dict(available=False, missing=['R3B_TAX_PAYABLE'])
    pos_turnover = sum(m['turnover'] for m in master if m['turnover'] > 0)
    hsn_pos, hsn_rows = {}, ing.data.get('HSN', [])
    for h in hsn_rows:
        code = str(h.get('HSN Code'))
        v = n0(h.get('Taxable Value'))
        if v > 0: hsn_pos.setdefault(code, []).append((h['_period'], v, n0(h.get('Total GST'))))
    events = []
    for m in master:
        if m['turnover'] >= 0 and m['out_total'] >= 0: continue
        codes = [str(h.get('HSN Code')) for h in hsn_rows
                 if h.get('_period') == m['period'] and n0(h.get('Taxable Value')) < 0]
        matched = {c: hsn_pos.get(c, []) for c in codes}
        events.append(dict(period=m['period'], fy=m['fy'], turnover=m['turnover'],
                           tax=m['out_total'], op_total=m['op_total'], hsn=codes,
                           prior_positive_same_hsn=matched,
                           pct_of_all_positive_turnover=(abs(m['turnover'])/pos_turnover
                                                         if pos_turnover else None)))
    # negative declared in GSTR-1 but not carried into 3B
    r1only = []
    if ing.have('R1_MONTHLY'):
        r1 = ing.by_period('R1_MONTHLY')
        for m in master:
            row = r1.get(m['period'])
            if row is None or row.get('Total GST') is None: continue
            t = n0(row.get('Total GST'))
            if t < -1 and m['out_total'] >= -1:
                r1only.append(dict(period=m['period'], r1_tax=t,
                                   r1_value=n0(row.get('Taxable Value')),
                                   r3b_tax=m['out_total'],
                                   inconsistent=(n0(row.get('Taxable Value')) >= 0),
                                   docs=row.get('No Of Invoices')))
    return dict(available=True, events=events, r1_only=r1only,
                positive_turnover_total=pos_turnover)

# -------------------------------------------------------- 8. return filing
def chk_return_filing(ing, master, r1res):
    if not ing.have('R3B_FILING'): return dict(available=False, missing=['R3B_FILING'])
    fl = ing.by_period('R3B_FILING')
    rows, late, delay = [], [], 0
    r1map = {r['period']: r['status'] for r in r1res['rows']} if r1res.get('available') else {}
    r2 = ing.by_period('R2A_MONTHLY') if ing.have('R2A_MONTHLY') else {}
    for m in master:
        f = fl.get(m['period'], {})
        d = f.get('Date Of Filing')
        dd = (d - due_date(m['period'])).days if d else None
        if dd and dd > 0: late.append((m['period'], d, dd)); delay += dd
        rows.append(dict(period=m['period'], fy=m['fy'], filed=d, due=due_date(m['period']),
                         delay=max(0, dd) if dd is not None else None,
                         r1_status=r1map.get(m['period'], 'not tested'),
                         r2a_present=m['period'] in r2,
                         turnover=m['turnover'], op=m['op_total']))
    return dict(available=True, rows=rows, late=late, total_delay=delay)

# ----------------------------------------------------------------- 9. cash
def chk_cash(ing, master):
    if not ing.have('PAYMENTS'): return dict(available=False, missing=['PAYMENTS'])
    pay = sorted(ing.data['PAYMENTS'], key=lambda r: (r.get('Payment Date') or datetime.date(1900,1,1)))
    tot = dict(SGST=0.0, CGST=0.0, IGST=0.0, TOTAL=0.0)
    by_fy = {}
    for r in pay:
        for k in tot: tot[k] += n0(r.get(k))
        d = r.get('Payment Date')
        if d:
            f = f"{d.year}-{str(d.year+1)[2:]}" if d.month > 3 else f"{d.year-1}-{str(d.year)[2:]}"
            by_fy[f] = by_fy.get(f, 0.0) + n0(r.get('TOTAL'))
    op = sum(m['op_total'] for m in master)
    itc = sum(m['itc'] for m in master)
    return dict(available=True, challans=pay, totals=tot, by_fy=by_fy,
                output_tax=op, itc=itc,
                cash_pct=(tot['TOTAL']/op if op else None),
                count=len(pay))

# ---------------------------------------------------------- 10. HSN drift
def chk_hsn(ing):
    if not ing.have('HSN'): return dict(available=False, missing=['HSN'])
    rows = sorted(ing.data['HSN'], key=lambda r: period_key(r['_period']))
    seq = []
    for r in rows:
        seq.append((r['_period'], str(r.get('HSN Code')), n0(r.get('Taxable Value')), n0(r.get('Total GST'))))
    codes = []
    for _, c, _, _ in seq:
        if c not in codes: codes.append(c)
    return dict(available=True, rows=seq, codes=codes)

# ------------------------------------------------------ 11. annual return
def chk_annual(ing, master):
    filed = {r['_fy']: r.get('Date of Filing') for r in ing.data.get('GSTR9', [])}
    by_fy = {}
    for m in master:
        by_fy[m['fy']] = by_fy.get(m['fy'], 0.0) + m['turnover']
    out = []
    for f in sorted(by_fy, key=lambda x: int(x.split('-')[0])):
        t = by_fy[f]
        out.append(dict(fy=f, turnover=t, gstr9_required=t > 20_000_000,
                        gstr9c_required=t > 50_000_000,
                        gstr9_filed=filed.get(f), on_record=f in filed))
    return dict(available=ing.have('GSTR9'), rows=out,
                missing=[] if ing.have('GSTR9') else ['GSTR9'])

# ----------------------------------------------------------- 12. FY summary
def fy_summary(master):
    out = {}
    for m in master:
        a = out.setdefault(m['fy'], dict(periods=0, nil=0, turnover=0.0, op=0.0, itc=0.0,
                                         cash=0.0, gross_sgst=0.0, sgst_exc=0.0,
                                         igst_exc=0.0, itc_nil=0.0))
        a['periods'] += 1; a['nil'] += 1 if m['nil'] else 0
        a['turnover'] += m['turnover']; a['op'] += m['op_total']; a['itc'] += m['itc']
        a['cash'] += m['cash_sgst']; a['gross_sgst'] += m['gross_sgst']
        a['sgst_exc'] += m['sgst_excess']; a['igst_exc'] += m['igst_excess']
        a['itc_nil'] += m['itc_in_nil']
    return out

# ------------------------------------------------------- 13. limitation
def limitation(fys, today=None):
    """Indicative only. Extension notifications must be checked separately."""
    today = today or datetime.date.today()
    out = []
    for f in sorted(fys, key=lambda x: int(x.split('-')[0])):
        end = int(f.split('-')[0]) + 1
        ar_due = datetime.date(end, 12, 31)                 # GSTR-9 due date
        s73_order = ar_due.replace(year=ar_due.year + 3)
        s73_scn = s73_order - datetime.timedelta(days=92)
        s74_order = ar_due.replace(year=ar_due.year + 5)
        s74_scn = s74_order - datetime.timedelta(days=183)
        is74a = end >= 2025                                  # FY 2024-25 onwards: s.73(12)/74(12) stop, s.74A runs
        m42 = ar_due.month + 42
        s74a_scn = datetime.date(ar_due.year + (m42 - 1) // 12, (m42 - 1) % 12 + 1, 30 if ((m42 - 1) % 12 + 1) in (4, 6, 9, 11) else 31) \
                   if ((m42 - 1) % 12 + 1) != 2 else datetime.date(ar_due.year + (m42 - 1) // 12, 2, 28)
        s74a_order = s74a_scn.replace(year=s74a_scn.year + 1)
        out.append(dict(fy=f, annual_return_due=ar_due, is74a=is74a,
                        s74a_scn_by=s74a_scn if is74a else None, s74a_order_by=s74a_order if is74a else None,
                        s74a_open=(today <= s74a_scn) if is74a else None,
                        s73_scn_by=s73_scn, s73_order_by=s73_order,
                        s73_open=today <= s73_scn,
                        s74_scn_by=s74_scn, s74_order_by=s74_order,
                        s74_open=today <= s74_scn,
                        regime='Section 74A applies (FY 2024-25 onwards)' if end >= 2025
                               else 'Sections 73 / 74 apply'))
    return out

# ------------------------------------------------- 14. GSTR-7 TDS (NEW v1.5)
R3B_CODES = ('TAX_PAID', 'R3B_TAX_PAYABLE', 'R3B_TURNOVER', 'R3B_FILING')
TDS_RATE = 0.02                 # s.51(1): 1% CGST + 1% SGST, or 2% IGST, of the value of supply
RULE86B_LIMIT = 5_000_000       # Rule 86B: taxable supply in a month above Rs.50 lakh
GSTR9_LIMIT, GSTR9C_LIMIT = 20_000_000, 50_000_000

def chk_tds(ing, master, today=None):
    """GSTR-7: tax deducted at source from THIS taxpayer by Government departments, local
    authorities and Government agencies (s.51, Notification 50/2018-CT). The deductor pays the
    supplier the contract value less 2%, and the 2% is credited to the supplier's electronic cash
    ledger (Rule 87(9)). So every rupee of TDS proves a payment received for a taxable supply of
    50 times that amount, excluding the tax on the invoice (s.51(2)).

    What that proves, and what it does not:
      * It is a FLOOR on the taxpayer's turnover with Government. Supplies to anyone else, and
        Government contracts below Rs.2.5 lakh, carry no TDS and are not in it.
      * Payment month is not invoice month. For services (works contract is a service) the
        time of supply is the earlier of invoice and payment (s.13(2)), and for goods the invoice
        is issued on removal (s.31(1)), so a supply paid for by month M must have been declared
        by month M at the latest. The firm test is therefore CUMULATIVE within the financial year,
        not month against month. The one exception is an advance received for goods, which is
        not taxable on receipt (Notification 66/2017-CT).
    """
    if not ing.have('R7_TDS'):
        return dict(available=False, missing=['R7_TDS'])
    r7 = ing.by_period('R7_TDS')
    have_3b = any(ing.have(c) for c in R3B_CODES)
    have_r1 = ing.have('R1_MONTHLY')
    mm = {m['period']: m for m in master}
    unfiled = set(unfiled_periods(ing))
    r1 = ing.by_period('R1_MONTHLY') if have_r1 else {}
    rows = []
    cum_tds_val, cum_3b = {}, {}
    # 3B cumulated from the first period of each FY, so a supply invoiced before the first TDS month
    # is still counted in the taxpayer's favour.
    master_by_fy = {}
    for m in master:
        master_by_fy.setdefault(m['fy'], []).append(m)
    for p in sorted(r7, key=period_key):
        x = r7[p]; f = fy_of(p)
        sg, cg, ig, ce = (x.get(k) for k in ('SGST', 'CGST', 'IGST', 'CESS'))
        tot = x.get('Total GST')
        heads = n0(sg) + n0(cg) + n0(ig)
        # Total GST is used where present: Prime rounds each head to the rupee but totals the
        # unrounded heads, so the total is the more exact figure. Cess is not deducted under s.51.
        value = ((n0(tot) - n0(ce)) if tot is not None else heads) / TDS_RATE
        cum_tds_val[f] = cum_tds_val.get(f, 0.0) + value
        # 3B position for this period
        today_ = today or datetime.date.today()
        if not have_3b:
            st3 = None
        elif p not in mm and due_date(p) >= today_:
            st3 = 'NOT YET DUE'          # GSTR-3B for this month is due on the 20th of the next month
        elif p in unfiled:
            st3 = 'NOT FILED'
        elif p in mm:
            st3 = 'FILED - NIL' if mm[p]['turnover'] == 0 else 'FILED'
        else:
            st3 = 'NO GSTR-3B ON RECORD'
        m = mm.get(p)
        cum = None
        if have_3b:
            cum = sum(z['taxable'] for z in master_by_fy.get(f, []) if period_key(z['period']) <= period_key(p))
        # GSTR-1 position
        if not have_r1:
            st1, v1 = None, None
        else:
            rr = r1.get(p)
            _, y_, m_ = parse_period(p)
            r1_due = datetime.date(y_ + (m_ == 12), (m_ % 12) + 1, 11)   # monthly GSTR-1: 11th of the next month
            if (rr is None or rr.get('Total GST') is None) and r1_due >= today_:
                st1, v1 = 'NOT YET DUE', None
            elif rr is None or rr.get('Total GST') is None:
                st1, v1 = 'NOT FILED', None
            else:
                v1 = n0(rr.get('Taxable Value'))
                st1 = 'FILED - NIL' if v1 == 0 and n0(rr.get('Total GST')) == 0 else 'FILED'
        obs = []
        if tot is not None and abs(heads + n0(ce) - n0(tot)) > 1.0:
            obs.append(f'Heads add to {heads + n0(ce):,.0f} but Total GST reads {n0(tot):,.0f}')
        if n0(sg) != n0(cg):
            obs.append('CGST and SGST deducted differ — check the deductor-wise detail')
        if st3 in ('NOT FILED', 'NO GSTR-3B ON RECORD'):
            obs.append('TDS deducted but no GSTR-3B for the month')
        elif st3 == 'FILED - NIL':
            obs.append('GSTR-3B filed NIL in a month in which the deductors paid for supplies')
        if m is not None and n0(cg) > m['out_cgst'] + 1:
            obs.append('CGST deducted at 1% exceeds the entire CGST declared in GSTR-3B')
        if st1 in ('NOT FILED',):
            obs.append('No GSTR-1 for a month with TDS payments — B2B invoices to the deductors not on record')
        if value > RULE86B_LIMIT:
            obs.append('Value paid for exceeds Rs.50 lakh: Rule 86B threshold crossed on supplies to the deductors alone')
        rows.append(dict(period=p, fy=f, deductors=x.get('No Of Deductors'), sgst=sg, cgst=cg, igst=ig, cess=ce,
                         total=tot, value=value, cum_value=cum_tds_val[f],
                         r3b_status=st3, r3b_taxable=(m['taxable'] if m else None), cum_3b=cum,
                         cum_short=(max(0.0, cum_tds_val[f] - cum) if cum is not None and st3 != 'NOT YET DUE' else None),
                         r3b_out_cgst=(m['out_cgst'] if m else None),
                         r1_status=st1, r1_value=v1, obs=obs))
    # financial-year roll-up
    fys = {}
    for r in rows:
        a = fys.setdefault(r['fy'], dict(months=0, tds=0.0, sgst=0.0, value=0.0, first=r['period'], last=r['period'],
                                         over50=[]))
        a['months'] += 1; a['tds'] += r['value'] * TDS_RATE
        a['sgst'] += n0(r['sgst']); a['value'] += r['value']; a['last'] = r['period']
        if r['value'] > RULE86B_LIMIT: a['over50'].append(r['period'])
    for f, a in fys.items():
        ms = master_by_fy.get(f, [])
        a['r3b_taxable'] = sum(z['taxable'] for z in ms) if have_3b else None
        a['r3b_months'] = len(ms)
        a['itc'] = sum(z['itc'] for z in ms) if have_3b else None
        a['cash_sgst'] = sum(z['cash_sgst'] for z in ms) if ing.have('TAX_PAID') else None
        a['r1_value'] = (sum(n0(r1[p].get('Taxable Value')) for p in r1 if fy_of(p) == f) if have_r1 else None)
        # a month whose return is not yet due is left out of the comparison, not counted as a shortfall
        a['value_due'] = sum(r['value'] for r in rows if r['fy'] == f and r['r3b_status'] != 'NOT YET DUE')
        a['value_due_r1'] = sum(r['value'] for r in rows if r['fy'] == f and r['r1_status'] != 'NOT YET DUE')
        a['short_3b'] = (max(0.0, a['value_due'] - a['r3b_taxable']) if a['r3b_taxable'] is not None else None)
        a['short_r1'] = (max(0.0, a['value_due_r1'] - a['r1_value']) if a['r1_value'] is not None else None)
        a['max_cum_short'] = max([r['cum_short'] or 0.0 for r in rows if r['fy'] == f], default=0.0) if have_3b else None
        a['gstr9'] = a['value'] > GSTR9_LIMIT
        a['gstr9c'] = a['value'] > GSTR9C_LIMIT
        # TDS sits in the cash ledger; SGST cash used in 3B below the SGST TDS credited means the
        # TDS was not absorbed — a balance building up, and a refund claim to watch for.
        a['tds_sgst_unabsorbed'] = (max(0.0, a['sgst'] - a['cash_sgst']) if a['cash_sgst'] is not None else None)
    no3b = [r['period'] for r in rows if r['r3b_status'] in ('NOT FILED', 'NO GSTR-3B ON RECORD')]
    nil3b = [r['period'] for r in rows if r['r3b_status'] == 'FILED - NIL']
    cgst_over = [r['period'] for r in rows if r['r3b_out_cgst'] is not None and n0(r['cgst']) > r['r3b_out_cgst'] + 1]
    no_r1 = [r['period'] for r in rows if r['r1_status'] == 'NOT FILED']
    total_mismatch = [r['period'] for r in rows if any(o.startswith('Heads add') for o in r['obs'])]
    deductor_max = max([n0(r['deductors']) for r in rows], default=0)

    # ---- the tests, each with its legal basis and whether the data held lets it run
    def rs(v): return f"Rs.{inr(v)}"
    fyl = sorted(fys, key=lambda f: int(f.split('-')[0]))
    T = []
    short_fy = [(f, fys[f]['short_3b']) for f in fyl if (fys[f]['short_3b'] or 0) > 1]
    cum_fy = [(f, fys[f]['max_cum_short']) for f in fyl if (fys[f]['max_cum_short'] or 0) > 1]
    T.append(('Value the deductors paid for against taxable turnover declared in GSTR-3B — month, cumulative and FY',
              'Sections 12, 13(2), 31 and 39; Rule 61. A supply paid for must have been declared.',
              'GSTR-3B: TAX PAID DETAILS (and R3B TAX PAYABLE)', have_3b,
              (('SHORTFALL: ' + '; '.join(f'FY {f} {rs(v)}' for f, v in short_fy) + '. ' if short_fy else
                'No FY shortfall: GSTR-3B taxable turnover covers the value the deductors paid for. ')
               + ('Peak cumulative shortfall within the year: ' + '; '.join(f'FY {f} {rs(v)}' for f, v in cum_fy) + '.'
                  if cum_fy else '')) if have_3b else 'Not run.'))
    T.append(('Months in which TDS was deducted but GSTR-3B was not filed or was filed NIL',
              'Section 39; section 46 (FORM GSTR-3A); section 62 best-judgment assessment of a non-filer.',
              'GSTR-3B: TAX PAID DETAILS', have_3b,
              ((f'No GSTR-3B: {", ".join(no3b)}. ' if no3b else '') + (f'GSTR-3B NIL: {", ".join(nil3b)}. ' if nil3b else '')
               + ('None — every month with TDS has a GSTR-3B with turnover.' if not (no3b or nil3b) else ''))
              if have_3b else 'Not run.'))
    T.append(('CGST deducted at 1% against the whole CGST declared in GSTR-3B for the month',
              'Sections 39 and 51. At any rate of 2% or more, tax declared cannot be below tax deducted at 1%.',
              'R3B TAX PAYABLE DETAILS', ing.have('R3B_TAX_PAYABLE'),
              ((f'CGST deducted EXCEEDS CGST declared in: {", ".join(cgst_over)}.' if cgst_over else 'No month where it does.')
               if ing.have('R3B_TAX_PAYABLE') else 'Not run.')))
    T.append(('GSTR-1 filed for every month with TDS, and its taxable value against the value paid for (FY)',
              'Section 37 and Rule 59 (B2B invoices to each deductor GSTIN); Rule 59(6). Section 47 late fee.',
              'R1 FILINGS IN THE YEAR', have_r1,
              (((f'No GSTR-1: {", ".join(no_r1)}. ' if no_r1 else 'GSTR-1 on record for every TDS month. ')
                + '; '.join(f'FY {f} shortfall against GSTR-1 {rs(fys[f]["short_r1"])}'
                            for f in fyl if (fys[f]['short_r1'] or 0) > 1)) if have_r1 else 'Not run.')))
    T.append(('TDS credited to the cash ledger against cash actually used in GSTR-3B (SGST head)',
              'Rule 87(9); section 49. Unused TDS builds a cash balance that can be claimed as refund under section 54 '
              '(excess balance in the electronic cash ledger).',
              'TAX PAID DETAILS', ing.have('TAX_PAID'),
              ('; '.join(f'FY {f}: SGST deducted {rs(fys[f]["sgst"])}, SGST cash used {rs(fys[f]["cash_sgst"])}, '
                         f'not absorbed {rs(fys[f]["tds_sgst_unabsorbed"])}' for f in fyl) + '.'
               if ing.have('TAX_PAID') else 'Not run.')))
    T.append(('ITC availed against GSTR-2A in the months the deductors paid for supplies',
              'Section 16(2)(aa), Rule 36(4); Rule 88D. A supplier paid by deductors who discharges output tax '
              'wholly from ITC while TDS sits unused is the pattern behind cash-ledger refund claims.',
              'EXCESS ITC CLAIMED, R2A FILINGS', ing.have('EXCESS_ITC') or ing.have('R2A_MONTHLY'),
              ('See 13_EXCESS_ITC and 19A_ITC_vs_2A for the same months.'
               if (ing.have('EXCESS_ITC') or ing.have('R2A_MONTHLY')) else 'Not run.')))
    T.append(('Annual return thresholds on receipts subject to TDS alone',
              'Section 44 and Rule 80: GSTR-9 above Rs.2 crore aggregate turnover (below it, optional where the year\'s notification exempts); self-certified GSTR-9C above Rs.5 crore.',
              'GSTR-7 alone', True,
              '; '.join(f'FY {f} ({fys[f]["first"]} to {fys[f]["last"]}): {rs(fys[f]["value"])} — '
                        + ('GSTR-9C required' if fys[f]['gstr9c'] else 'GSTR-9 required' if fys[f]['gstr9']
                           else 'below Rs.2 crore on these receipts alone') for f in fyl) + '.'))
    T.append(('E-invoicing and QRMP eligibility',
              'Rule 48(4) with Notification 13/2020-CT as amended (e-invoice once aggregate turnover in any preceding FY '
              'exceeds Rs.5 crore); Rule 61A (QRMP only up to Rs.5 crore).',
              'GSTR-7 alone', True,
              '; '.join(f'FY {f}: {rs(fys[f]["value"])} from the deductors alone'
                        + (' — above Rs.5 crore: e-invoicing from the next FY, QRMP not available'
                           if fys[f]['value'] > GSTR9C_LIMIT else
                           (' in a part year — close to Rs.5 crore; watch the rest of the year'
                            if fys[f]['value'] > 0.75 * GSTR9C_LIMIT else '')) for f in fyl) + '.'))
    over = [p for f in fyl for p in fys[f]['over50']]
    T.append(('Rule 86B — months in which the deductors alone paid for more than Rs.50 lakh of supplies',
              'Rule 86B (not more than 99% of output tax from ITC, subject to the proviso). TDS credit sits in the cash '
              'ledger, so TDS actually USED counts as cash.',
              'GSTR-7 alone; TAX PAID DETAILS to test', True,
              (f'Threshold crossed on receipts subject to TDS alone in {", ".join(over)}. Whether the 1% was met depends on the '
               f'cash and TDS actually used — test on the electronic cash ledger.' if over else 'No month above Rs.50 lakh.')))
    T.append(('Place of supply — intra-State or inter-State deduction',
              'Section 51(1) and its proviso: IGST is deducted on an inter-State supply; nothing is deducted where the '
              "supplier's location and the place of supply are both in a State other than the recipient's registration.", 'GSTR-7 alone', True,
              ('IGST was deducted in some months — inter-State supplies; check GSTR-1 shows them as inter-State.'
               if any(n0(r['igst']) > 0 for r in rows) else
               'Only CGST and SGST deducted: every deductor is in the same State as the supply. GSTR-1 should show these '
               'as intra-State B2B supplies.')))

    return dict(available=True, rows=rows, fy=fys, have_3b=have_3b, have_r1=have_r1, tests=T,
                have_tax_paid=ing.have('TAX_PAID'), have_2a=ing.have('R2A_MONTHLY') or ing.have('EXCESS_ITC'),
                no_3b=no3b, nil_3b=nil3b, cgst_over=cgst_over, no_r1=no_r1, total_mismatch=total_mismatch,
                interstate=any(n0(r['igst']) > 0 for r in rows), deductor_max=deductor_max,
                total_value=sum(r['value'] for r in rows), total_tds=sum(a['tds'] for a in fys.values()),
                first=rows[0]['period'] if rows else None, last=rows[-1]['period'] if rows else None)

# ------------------------------------- 15. GSTR-7 party-wise detail (NEW v1.7)
PAN_HOLDER = {'P': 'Individual / proprietor', 'F': 'Firm / LLP', 'H': 'HUF', 'C': 'Company',
              'A': 'Association of persons', 'T': 'Trust', 'B': 'Body of individuals',
              'L': 'Local authority', 'J': 'Artificial juridical person', 'G': 'Government', 'K': 'Krish (unused)'}
# who may deduct under s.51(1)(a)-(d) read with Notification 50/2018-CT
ELIGIBLE = {'G': 'ELIGIBLE TYPE — Government', 'L': 'ELIGIBLE TYPE — local authority'}
VERIFY = {'C': 'VERIFY — a company can deduct only if it is a public sector undertaking or a Government-controlled body',
          'J': 'VERIFY — eligible only if an authority / board set up by statute or Government',
          'A': 'VERIFY — eligible only if a society established by Government or a local authority',
          'T': 'VERIFY — eligible only if a society / body established by Government',
          'B': 'VERIFY — eligible only if a body established by Government'}
TAN_BODY = re.compile(r'^[A-Z]{4}\d{5}[A-Z]$')
PAN_BODY = re.compile(r'^[A-Z]{5}\d{4}[A-Z]$')

def gstin_profile(g):
    """What a GSTIN says about itself, on its face. Position 14 is 'Z' for an ordinary registration;
    a tax-deductor registration (REG-07, section 24(vi)) carries 'D' there. Positions 3-12 are the
    PAN (or, for most Government deductors, the TAN); the 4th character of a PAN is the holder type."""
    g = (g or '').strip().upper()
    if len(g) != 15: return dict(gstin=g, reg='INVALID LENGTH', body='', holder='', eligibility='VERIFY — malformed GSTIN')
    body, c14 = g[2:12], g[13]
    reg = {'Z': 'Ordinary registration', 'D': 'Tax-deductor registration (REG-07)', 'C': 'Tax-collector registration'}.get(
        c14, f'Unusual 14th character "{c14}"')
    if TAN_BODY.match(body):
        return dict(gstin=g, reg=reg, body='TAN', holder='TAN-based (typical of a Government DDO)',
                    eligibility='ELIGIBLE TYPE — TAN-based deductor (confirm it is a notified deductor)')
    if PAN_BODY.match(body):
        h = body[3]
        el = ELIGIBLE.get(h) or VERIFY.get(h) or ('NOT A PERMITTED DEDUCTOR ON THE FACE OF THE PAN — holder type "' +
                                                    PAN_HOLDER.get(h, f'PAN type {h}') +
                                                    '" is not within section 51(1) or Notification 50/2018-CT')
        return dict(gstin=g, reg=reg, body='PAN', holder=PAN_HOLDER.get(h, f'PAN type {h}'), eligibility=el)
    return dict(gstin=g, reg=reg, body='?', holder='Neither PAN nor TAN pattern', eligibility='VERIFY — malformed body')

def chk_tds_detail(ing, case_gstin=None):
    """The party-wise GSTR-7 detail, one Prime export per month. Tested for: who the listed parties are
    (registration type and PAN holder type, against who section 51 allows to deduct); agreement with the
    month-wise summary; rows just above the Rs.2.5 lakh deduction threshold; and which months are missing."""
    if not ing.have('R7_DETAIL'):
        return dict(available=False, missing=['R7_DETAIL'])
    rows = sorted(ing.data['R7_DETAIL'], key=lambda r: (period_key(r['_period']), r.get('SNo') or 0))
    subj = sorted({x.get('subject_gstin') for x in ing.register if x.get('code') == 'R7_DETAIL' and x.get('subject_gstin')})
    parties = {}
    for r in rows:
        g = (r.get('Deductee GSTIN') or '').strip().upper()
        p = parties.setdefault(g, dict(names=[], months=[], lines=0, value=0.0, tds=0.0, **gstin_profile(g)))
        nm = r.get('Trade Name')
        if nm and nm not in p['names']: p['names'].append(nm)
        if r['_period'] not in p['months']: p['months'].append(r['_period'])
        p['lines'] += 1; p['value'] += n0(r.get('Taxable Value')); p['tds'] += n0(r.get('Total GST'))
    # month reconciliation with the summary
    summ = ing.by_period('R7_TDS') if ing.have('R7_TDS') else {}
    months = []
    for m in sorted({r['_period'] for r in rows}, key=period_key):
        mr = [r for r in rows if r['_period'] == m]
        s = summ.get(m)
        tds = sum(n0(r.get('Total GST')) for r in mr)
        months.append(dict(period=m, lines=len(mr), parties=len({r.get('Deductee GSTIN') for r in mr}),
                           value=sum(n0(r.get('Taxable Value')) for r in mr), tds=tds,
                           s_parties=(s or {}).get('No Of Deductors'), s_tds=(s or {}).get('Total GST'),
                           agrees=(None if s is None else
                                   abs(tds - n0(s.get('Total GST'))) <= 1.0 + 0.5 * len(mr)
                                   and int(n0(s.get('No Of Deductors'))) == len({r.get('Deductee GSTIN') for r in mr}))))
    missing = sorted(set(summ) - {m['period'] for m in months}, key=period_key)
    # rate check line by line: each head should be 1% of the taxable value (2% IGST)
    off_rate = []
    near = []
    for r in rows:
        v = n0(r.get('Taxable Value'))
        if v and abs(n0(r.get('CGST')) - v * 0.01) > 1.0 and abs(n0(r.get('IGST')) - v * 0.02) > 1.0:
            off_rate.append((r['_period'], r.get('Deductee GSTIN'), v, n0(r.get('CGST'))))
        if 250_000 < v <= 260_000:
            near.append((r['_period'], r.get('Deductee GSTIN'), v))
    lst = list(parties.values())
    all_d = all(p['reg'].startswith('Tax-deductor') for p in lst)
    subj_z = bool(subj) and all(x[13] == 'Z' for x in subj)
    if all_d and subj_z:
        role = ('The listed parties hold TAX-DEDUCTOR registrations (14th character D) and the GSTIN the export is titled '
                'for holds an ordinary registration (14th character Z). A GSTR-7 can be filed only from a deductor '
                'registration, so the LISTED PARTIES ARE THE DEDUCTORS and the case taxpayer is the DEDUCTEE. Prime heads '
                'the column "Deductee GSTIN"; the summary report counts the same parties as "No Of Deductors". Confirm '
                'one GSTIN on the portal (Search Taxpayer shows the registration type).')
        role_code = 'DEDUCTORS'
    else:
        role = ('The export does not settle on its face whether the listed parties are the deductors or the deductees. '
                'Check each GSTIN on the portal before relying on the eligibility column.')
        role_code = 'UNSETTLED'
    bad = [p for p in lst if p['eligibility'].startswith('NOT A PERMITTED')]
    total_v = sum(p['value'] for p in lst)
    return dict(available=True, rows=rows, parties=sorted(lst, key=lambda p: -p['value']), months=months,
                missing_months=missing, subject=subj,
                subject_mismatch=[x for x in subj if case_gstin and x != case_gstin.upper()],
                role=role, role_code=role_code, ineligible=bad,
                ineligible_value=sum(p['value'] for p in bad), ineligible_tds=sum(p['tds'] for p in bad),
                total_value=total_v, total_tds=sum(p['tds'] for p in lst),
                top_share=(max((p['value'] for p in lst), default=0) / total_v if total_v else None),
                off_rate=off_rate, near_threshold=near)

def run_all(ing):
    m = build_master(ing)
    r1 = chk_r1_vs_r3b(ing, m)
    res = dict(
        master=m, fy=fy_summary(m),
        r1_vs_r3b=r1,
        excess_itc=chk_excess_itc(ing, m),
        itc_vs_2a=chk_itc_vs_2a(ing, m),
        dormancy=chk_dormancy(ing, m),
        rule86b=chk_rule86b(ing, m),
        rcm=chk_rcm(ing, m),
        credit_notes=chk_credit_notes(ing, m),
        filing=chk_return_filing(ing, m, r1),
        cash=chk_cash(ing, m),
        hsn=chk_hsn(ing),
        annual=chk_annual(ing, m),
        tds=chk_tds(ing, m),
        tds_detail=chk_tds_detail(ing, getattr(ing, 'case_gstin', None)),
        limitation=limitation(_fys(ing.all_periods())),
        unfiled=unfiled_periods(ing),
    )
    return res
