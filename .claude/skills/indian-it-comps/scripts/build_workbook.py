"""Phase 5b: build output/trading_comps.xlsx (live formulas) + Python shadow calculation.

Inputs (as printed, blue) -> Components (SUMIFS-able line items) -> Standardised (INR Cr)
-> EV Bridge -> Comps (multiples + Mean/Median/P25/P75). Audit Log cites every input.
Shadow results (same maths in Python) -> work/comps_shadow.json ; cell map -> work/cellmap.json.

Usage:  python build_workbook.py [co_id ...]     (default: every peer with extracted.json + market.json OK)
"""
from __future__ import annotations

import re
import statistics
import sys

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from common import OUTPUT_DIR, ROOT, WORK_DIR, co_work_dir, load_yaml, read_json, write_json

UNIT_TO_CR = {"million": 0.1, "crore": 1.0, "lakh": 0.01, "thousand": 0.0001}
BLUE, BOLD = Font(color="0000CC"), Font(bold=True)
HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(bold=True, color="FFFFFF")
SUB_FILL = PatternFill("solid", fgColor="D9E1F2")
THIN = Border(top=Side(style="thin"))
NUM, NUM0, MULT, PCT = '#,##0.00;(#,##0.00);"-"', '#,##0;(#,##0);"-"', '0.0"x"', '0.0%'
FIRST_COL = 3          # companies start in column C
NM_CAP = 100           # multiples above this are NM


def v(c):
    return None if c is None else c["value"]


# ------------------------------------------------------------------ data
def load_companies(co_ids):
    peers = {p["co_id"]: p for p in load_yaml("peers.yaml")["peers"]}
    out = []
    for co in co_ids:
        wd = co_work_dir(co)
        ex, mk = read_json(wd / "extracted.json"), read_json(wd / "market.json")
        if mk.get("status") != "OK":
            raise SystemExit(f"{co}: market.json status {mk.get('status')} - run market_data.py first")
        out.append({"co": co, "peer": peers[co], "ex": ex, "mk": mk, "manifest": read_json(wd / "manifest.json")})
    return out


def _sub(m, k):
    subs = m.get("listed_subsidiaries") or []
    return subs[0][k] if subs else None


INPUT_ROWS = [  # key, label, unit, getter(ex, mk) -> cite dict | scalar | None
    ("hdr", "Market data", None, None),
    ("price", "Share price (last close)", "INR", lambda e, m: m["price_inr"]),
    ("price_date", "Price date", "date", lambda e, m: m["price_date"]),
    ("basic_shares", "Basic shares (AR year-end +/- share bridge, or override)", "count", lambda e, m: m["basic_shares"]),
    ("low_52w", "52-week low (daily close)", "INR", lambda e, m: m.get("low_52w")),
    ("high_52w", "52-week high (daily close)", "INR", lambda e, m: m.get("high_52w")),
    ("hdr", "Listed subsidiary (NCI at market value)", None, None),
    ("sub_price", "Listed subsidiary: share price", "INR", lambda e, m: _sub(m, "price_inr")),
    ("sub_shares", "Listed subsidiary: shares outstanding", "count", lambda e, m: _sub(m, "shares")),
    ("sub_nci_pct", "Listed subsidiary: NCI share", "%", lambda e, m: _sub(m, "nci_pct")),
    ("sub_nci_book", "Listed subsidiary: book NCI (replaced by market value)", "rep", lambda e, m: _sub(m, "nci_book")),
    ("hdr", "Reporting basis", None, None),
    ("unit_factor", "Reported unit -> INR Cr factor", "x", lambda e, m: UNIT_TO_CR[e["unit_as_printed"]]),
    ("hdr", "P&L as printed (reported unit)", None, None),
    ("revenue", "Revenue from operations", "rep", lambda e, m: e["pl"]["revenue_from_operations"]),
    ("other_income", "Other income", "rep", lambda e, m: e["pl"]["other_income"]),
    ("fx_in_oi", "  of which FX gain/(loss) inside Other income", "rep", lambda e, m: e["pl"]["fx_in_other_income"]),
    ("share_eq", "Share of profit/(loss) of associates & JVs (excluded from EBITDA)", "rep",
     lambda e, m: e["pl"].get("share_of_equity_investees")),
    ("total_expenses", "Total expenses", "rep", lambda e, m: e["pl"]["total_expenses"]),
    ("finance_costs", "Finance costs", "rep", lambda e, m: e["pl"]["finance_costs"]),
    ("dep_amort", "Depreciation & amortisation", "rep", lambda e, m: e["pl"]["depreciation_amortisation"]),
    ("fx_line", "Separate FX line on face (memo, inside total expenses)", "rep", lambda e, m: e["pl"]["fx_line_charge"]),
    ("exceptional", "Exceptional items - charge/(gain)", "rep", lambda e, m: e["pl"]["exceptional_charge"]),
    ("exc_tax", "Tax effect on exceptional (if disclosed)", "rep", lambda e, m: e["pl"]["exceptional_tax_effect"]),
    ("pbt", "Profit before tax (continuing)", "rep", lambda e, m: e["pl"]["profit_before_tax"]),
    ("tax", "Total tax expense", "rep", lambda e, m: e["pl"]["total_tax_expense"]),
    ("pat_cont", "Profit - continuing operations", "rep", lambda e, m: e["pl"]["profit_continuing_ops"]),
    ("pat_disc", "Profit - discontinued operations", "rep", lambda e, m: e["pl"]["profit_discontinued_ops"]),
    ("disc_owners", "Discontinued profit attributable to owners", "rep", lambda e, m: e["pl"]["discontinued_attrib_owners"]),
    ("profit_year", "Profit for the year", "rep", lambda e, m: e["pl"]["profit_for_year"]),
    ("pat_owners", "Profit attributable to owners", "rep", lambda e, m: e["pl"]["pat_owners"]),
    ("pat_nci", "Profit attributable to NCI", "rep", lambda e, m: e["pl"]["pat_nci"]),
    ("eps_dil", "Diluted EPS as reported (INR, memo)", "INR", lambda e, m: e["pl"]["eps_diluted"]),
    ("hdr", "Balance sheet as printed (reported unit)", None, None),
    ("borr_nc", "Borrowings - non-current", "rep", lambda e, m: e["bs"]["borrowings_non_current"]),
    ("borr_c", "Borrowings - current", "rep", lambda e, m: e["bs"]["borrowings_current"]),
    ("lease_nc", "Lease liabilities - non-current", "rep", lambda e, m: e["bs"]["lease_liabilities_non_current"]),
    ("lease_c", "Lease liabilities - current", "rep", lambda e, m: e["bs"]["lease_liabilities_current"]),
    ("nci_book", "Non-controlling interests (book)", "rep", lambda e, m: e["bs"]["nci"]),
    ("eq_method", "Investments in associates / JVs", "rep", lambda e, m: e["bs"]["equity_method_investments"]),
    ("cash_face", "Cash & cash equivalents (face, memo)", "rep", lambda e, m: e["bs"]["cash_and_cash_equivalents"]),
    ("obb_face", "Other bank balances (face, memo)", "rep", lambda e, m: e["bs"]["other_bank_balances"]),
    ("ci_face", "Current investments (face, memo)", "rep", lambda e, m: e["bs"]["current_investments"]),
    ("nci_face", "Non-current investments (face, memo)", "rep", lambda e, m: e["bs"]["non_current_investments"]),
]


def components(c, policy):
    """Rows: (category, label, value, include, strike, cite)."""
    ex, pol = c["ex"], policy["enterprise_value"]["cash"]
    rows = []
    for x in ex["cash_components"]:
        rows.append(("cash", f"[{x['bucket']}] {x['label']}", x["value"], not x["restricted"], None, x))
    for x in ex.get("restricted_carveouts", []):
        rows.append(("carveout", x["label"], x["value"], True, None, x))
    invit_rx = re.compile(r"\binvits?\b|\breits?\b|infrastructure\s+investment\s+trust|real\s+estate\s+investment\s+trust", re.I)
    for x in ex["investment_components"]:
        treasury = x["classification"] == "treasury" or (
            pol.get("invits_reits_as_treasury") and x["classification"] == "other" and invit_rx.search(x["label"]))
        inc = treasury and (x["bucket"] == "current" and pol["current_investments_liquid"]
                                                     or x["bucket"] == "non_current" and pol["non_current_treasury_investments"])
        rows.append(("investment", f"[{x['bucket']}/{x['classification']}] {x['label']}", x["value"], inc, None, x))
    for x in ex["pl"].get("fx_in_other_income_extra") or []:
        rows.append(("fx_oi", f"[extra FX line in Other income] {x['label']}", x["value"], True, None, x))
    for x in ex["debt_like_items"]:
        rows.append(("debt_like", f"[{x['type']}] {x['label']}", x["value"], True, None, x))
    for p in ex["esop_plans"]:
        rows.append(("esop", f"[{p.get('instrument')}] {p['plan']}", p["outstanding_year_end"], True,
                     p["waep_inr"] or 0.0, {**p, "label": p["plan"], "note": None}))
    return rows


# ------------------------------------------------------------------ shadow (pure Python)
def shadow(c, policy):
    ex, mk = c["ex"], c["mk"]
    pl, bs = ex["pl"], ex["bs"]
    f = UNIT_TO_CR[ex["unit_as_printed"]]
    n = lambda x: v(x) or 0.0
    t = policy["pe"]["exceptional_tax_rate_default"]
    comp = components(c, policy)
    s = {}
    s["revenue"] = v(pl["revenue_from_operations"]) * f
    s["ebitda"] = (v(pl["profit_before_tax"]) + v(pl["finance_costs"]) + v(pl["depreciation_amortisation"])
                   + n(pl["exceptional_charge"]) - v(pl["other_income"]) + n(pl["fx_in_other_income"])
                   + sum(x["value"] for x in pl.get("fx_in_other_income_extra") or [])
                   - n(pl.get("share_of_equity_investees"))) * f
    exc = n(pl["exceptional_charge"])
    exc_tax = v(pl["exceptional_tax_effect"]) if pl["exceptional_tax_effect"] is not None else exc * t
    disc_own = v(pl["discontinued_attrib_owners"]) if pl["discontinued_attrib_owners"] is not None else \
        n(pl["profit_discontinued_ops"]) * (1 - policy["pe"]["discontinued_nci_share_default"])
    s["adj_pat"] = (v(pl["pat_owners"]) - disc_own + exc - exc_tax) * f
    s["rep_pat"] = v(pl["pat_owners"]) * f
    sumc = lambda cat: sum(r[2] for r in comp if r[0] == cat and r[3])
    s["cash"] = (sumc("cash") - sumc("carveout") + sumc("investment")) * f
    s["borrowings"] = (n(bs["borrowings_non_current"]) + n(bs["borrowings_current"])) * f
    s["leases"] = (n(bs["lease_liabilities_non_current"]) + n(bs["lease_liabilities_current"])) * f
    s["debt_like"] = sumc("debt_like") * f
    sub = (mk.get("listed_subsidiaries") or [None])[0]
    s["nci"] = (n(bs["nci"]) - (sub["nci_book"] if sub else 0)) * f + (
        sub["price_inr"] * sub["shares"] * sub["nci_pct"] / 1e7 if sub else 0)
    s["eq_method"] = n(bs["equity_method_investments"]) * f
    price = mk["price_inr"]
    tsm = sum(r[2] * (1 - r[4] / price) for r in comp if r[0] == "esop" and r[4] < price)
    s["diluted_shares_mn"] = (mk["basic_shares"] + tsm) / 1e6
    s["equity_value"] = price * s["diluted_shares_mn"] / 10
    s["ev"] = s["equity_value"] + s["borrowings"] + s["leases"] + s["debt_like"] + s["nci"] - s["cash"] - s["eq_method"]
    mult = lambda a, b: "NM" if b <= 0 or a / b > NM_CAP or a / b < 0 else a / b
    s["ev_rev"], s["ev_ebitda"] = mult(s["ev"], s["revenue"]), mult(s["ev"], s["ebitda"])
    s["pe_adj"], s["pe_rep"] = mult(s["equity_value"], s["adj_pat"]), mult(s["equity_value"], s["rep_pat"])
    # football field: EV/EBITDA re-priced at the 52-week closing low / high (diluted shares and FY26 net debt held)
    ev_at = lambda p: s["ev"] + (p - price) * s["diluted_shares_mn"] / 10
    s["ff_low"] = ev_at(mk["low_52w"]) / s["ebitda"] if s["ebitda"] > 0 else 0
    s["ff_high"] = ev_at(mk["high_52w"]) / s["ebitda"] if s["ebitda"] > 0 else 0
    return s


def pctl(xs, q):     # Excel PERCENTILE / PERCENTILE.INC (linear interpolation)
    xs = sorted(xs)
    if len(xs) == 1:
        return xs[0]
    k = (len(xs) - 1) * q
    lo = int(k)
    return xs[lo] + (xs[min(lo + 1, len(xs) - 1)] - xs[lo]) * (k - lo)


# ------------------------------------------------------------------ workbook helpers
def hdr_row(ws, r, labels, start=1):
    for i, t in enumerate(labels):
        cell = ws.cell(r, start + i, t)
        cell.fill, cell.font = HDR_FILL, HDR_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def section(ws, r, text, ncols):
    for col in range(1, ncols + 1):
        ws.cell(r, col).fill = SUB_FILL
    ws.cell(r, 1, text).font = BOLD


def cite_comment(cite, pdf):
    if not isinstance(cite, dict) or "uid" not in cite:
        return None
    return Comment(f"{pdf} | PDF p.{int(cite['uid'][1:4])} ({cite['uid'][-1]}) | printed p.{cite.get('printed_page')}\n"
                   f"\"{cite.get('label', '')[:120]}\"", "comps-engine")


def _d(iso):     # 2026-11-17 -> 17 Nov 2026
    import datetime as dt
    return dt.date.fromisoformat(iso[:10]).strftime("%d %b %Y").lstrip("0")


def data_dates(cos):
    peers_cfg = load_yaml("peers.yaml")
    px = sorted({c["mk"]["price_date"] for c in cos})
    fetched = max(c["mk"]["fetched_at"] for c in cos)
    price_txt = _d(px[0]) if len(px) == 1 else f"{_d(px[0])} to {_d(px[-1])}"
    return (f"Financials: {peers_cfg['fiscal_year']} ({_d(str(peers_cfg['balance_sheet_date']))}).  "
            f"Prices: last NSE close {price_txt} (fetched {_d(fetched)}).")


def date_banner(ws, text, ncols, row=2):
    for col in range(1, ncols + 1):
        ws.cell(row, col).fill = PatternFill("solid", fgColor="FFF2CC")
    c = ws.cell(row, 1, text)
    c.font = Font(bold=True, size=11, color="7F6000")


def known_gaps(cos):
    """Company-specific limitations, derived from extraction flags, manifests and market data."""
    out = []
    deal_rx = re.compile(r"subsequent|after (the )?(year|balance sheet|reporting)|post[- ](year|balance)", re.I)
    for c in cos:
        nm, ex, mk = c["peer"]["name"], c["ex"], c["mk"]
        sbp = c["manifest"]["targets"].get("share_based_payments", [{}])
        trunc = any(e.get("status") == "truncated" for e in sbp) or any(
            re.search(r"truncat|cut off", f, re.I) and re.search(r"esop|share[- ]based|option", f, re.I) for f in ex["flags"])
        if not ex["esop_plans"]:
            out.append(("ESOP dilution", nm, "No ESOP/RSU outstanding table found in the sliced notes - diluted shares "
                                            "exclude options (typically 0.5-1.5% of equity value)."))
        elif trunc:
            out.append(("ESOP dilution", nm, "ESOP note truncated in the slice - later plans/series may be missing; "
                                            "dilution may be understated."))
        for f in ex["flags"]:
            if deal_rx.search(f):
                out.append(("Post-March event", nm, f[:230]))
        if mk.get("shares_override"):
            out.append(("Share count", nm, f"Basic shares taken from yfinance by instruction ({mk['basic_shares']:,}); "
                                          f"differs from AR year-end by {mk['yf_unexplained_residual']:+.1%}, cause "
                                          f"undocumented - perimeter may not match FY26 financials."))
        if c["peer"].get("exclude_from_stats"):
            out.append(("Excluded from stats", nm, c["peer"].get("stats_note", "")))
    return out


def add_methodology(wb, cos, policy, banner):
    ws = wb.create_sheet("Methodology & Limitations", 1)
    ws["A1"], ws["A1"].font = "Methodology & limitations", Font(bold=True, size=14)
    date_banner(ws, banner, 3)
    subs = [f"{s['name']} ({c['peer']['name']}): NCI {s['nci_pct']:.2%} x market cap at INR {s['price_inr']:,.2f} "
            f"({_d(s['price_date'])}); replaces book NCI INR {s['nci_book']:,} mn [{s['source']}]"
            for c in cos for s in c["mk"].get("listed_subsidiaries") or []]
    sections = [
        ("Scope", ["Indian mid-tier IT services peers; consolidated financial statements only (never standalone). "
                   "All figures INR Crore unless stated; reported INR million converted in Python/Excel (x0.1)."]),
        ("Enterprise value", ["EV = diluted equity value + borrowings + lease liabilities + debt-like items + "
                              "non-controlling interests - unrestricted cash & treasury investments - associates/JVs.",
                              "Balance-sheet items at 31 Mar 2026; equity value at the latest close."]),
        ("EBITDA & earnings", ["Adjusted EBITDA = revenue from operations - operating expenses (continuing operations): "
                               "excludes Other income (except FX), exceptional items and share of associates/JVs. "
                               "FX gains/losses are operating wherever presented.",
                               "Adjusted PAT = PAT to owners - discontinued operations + exceptional items x (1 - tax; "
                               f"{policy['pe']['exceptional_tax_rate_default']:.2%} if not disclosed). Reported P/E shown as memo."]),
        ("Cash", ["Included: cash & cash equivalents, other bank balances, current and non-current treasury investments "
                  "(mutual funds, bonds, term deposits, InvIT/REIT units).",
                  "Excluded: restricted balances (unclaimed dividend, margin money, deposits under lien or against "
                  "guarantees, escrow), strategic equity stakes, finance-lease receivables."]),
        ("Leases & debt-like items", ["Ind AS 116 lease liabilities (current + non-current) are included in EV for every "
                                      "peer; EBITDA is therefore pre-rent (post-Ind AS 116), consistent across peers.",
                                      "Debt-like: contingent/deferred acquisition consideration (earn-outs) and put-option "
                                      "liabilities over NCI, at carrying value."]),
        ("Non-controlling interests", ["Unlisted subsidiaries at book value."] + (subs or ["No listed subsidiaries."])),
        ("Share count", ["Basic = AR year-end shares (net of ESOP-trust shares) + documented issuances that belong to "
                         "the FY26 perimeter (config/share_bridge.yaml); yfinance count used only as a reconciliation "
                         "check (max 0.5% unexplained) unless overridden by instruction (see gaps).",
                         "Diluted = basic + treasury-stock-method dilution per ESOP/RSU plan at the current price."]),
        ("Peer statistics & chart", [f"Mean / median / 25th / 75th percentile over companies with 'In peer stats? = Yes'; "
                                     f"multiples <= 0 or > {NM_CAP}x are NM and excluded.",
                                     "Football field: EV/EBITDA re-priced at the 52-week closing low and high, holding "
                                     "diluted shares and FY26 net debt constant (an approximation)."]),
        ("Sources", ["Financial statements: FY26 Annual Reports (page-cited in the Audit Log). Prices, 52-week range, "
                     "listed-subsidiary market data and cross-check share counts: Yahoo Finance via yfinance."]),
    ]
    r = 4
    for head, items in sections:
        ws.cell(r, 1, head).font = BOLD
        for t in items:
            ws.cell(r, 2, t).alignment = Alignment(wrap_text=True, vertical="top")
            r += 1
        r += 0
    r += 1
    ws.cell(r, 1, "Known gaps & limitations").font = Font(bold=True, size=12, color="C00000")
    r += 1
    for cat, nm, txt in known_gaps(cos):
        ws.cell(r, 1, cat).font = Font(bold=True)
        ws.cell(r, 2, f"{nm}: {txt}").alignment = Alignment(wrap_text=True, vertical="top")
        r += 1
    ws.cell(r + 1, 1, "Not covered in v1: LTM / forward multiples, pro-forma adjustments for acquisitions, "
                      "calendarisation (all peers are March year-end).").font = Font(italic=True, color="808080")
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 120
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_area = f"A1:B{r + 1}"
    return ws


def add_football_field(wb, cos, IR, ER, SR, L, cell_map, banner):
    from openpyxl.chart import BarChart, LineChart, Reference
    from openpyxl.chart.marker import Marker
    ws = wb.create_sheet("Football Field", 1)
    ws["A1"], ws["A1"].font = "EV / EBITDA football field - 52-week trading range (FY26A EBITDA)", Font(bold=True, size=14)
    date_banner(ws, banner, 8)
    hdr_row(ws, 4, ["Company", "EV/EBITDA at 52w low", "EV/EBITDA at 52w high", "Current EV/EBITDA",
                    "Range (high - low)", "Peer median (current)", "In peer stats?"])
    med = cell_map["stats.Median.ev_ebitda"]
    r0 = 5
    for i, c in enumerate(cos):
        r, cl = r0 + i, L(i)
        ebitda, ev, dil, px = (f"Standardised!{cl}{SR['ebitda']}", f"'EV Bridge'!{cl}{ER['ev']}",
                               f"'EV Bridge'!{cl}{ER['diluted']}", f"Inputs!{cl}{IR['price']}")
        at = lambda p: f"=IF({ebitda}<=0,0,({ev}+({p}-{px})*{dil}/10)/{ebitda})"
        ws.cell(r, 1, f'=Inputs!{cl}$3&IF(Comps!O{r}="No"," (excl.)","")')
        ws.cell(r, 2, at(f"Inputs!{cl}{IR['low_52w']}")).number_format = MULT
        ws.cell(r, 3, at(f"Inputs!{cl}{IR['high_52w']}")).number_format = MULT
        ws.cell(r, 4, f"=IF({ebitda}<=0,0,{ev}/{ebitda})").number_format = MULT
        ws.cell(r, 5, f"=C{r}-B{r}").number_format = MULT
        ws.cell(r, 6, f'=IF(ISNUMBER({med}),{med},0)').number_format = MULT
        ws.cell(r, 7, f"=Comps!O{r}").alignment = Alignment(horizontal="center")
        cell_map[f"ff.low.{c['co']}"] = f"'Football Field'!B{r}"
        cell_map[f"ff.high.{c['co']}"] = f"'Football Field'!C{r}"
    rl = r0 + len(cos) - 1
    for col, w in zip("ABCDEFG", (30, 14, 14, 14, 14, 14, 10)):
        ws.column_dimensions[col].width = w

    bar = BarChart()
    bar.type, bar.grouping, bar.overlap, bar.gapWidth = "col", "stacked", 100, 60
    bar.add_data(Reference(ws, min_col=2, min_row=4, max_row=rl), titles_from_data=True)
    bar.add_data(Reference(ws, min_col=5, min_row=4, max_row=rl), titles_from_data=True)
    bar.set_categories(Reference(ws, min_col=1, min_row=r0, max_row=rl))
    bar.series[0].graphicalProperties.noFill = True          # invisible base -> floating range bars
    bar.series[0].graphicalProperties.line.noFill = True
    bar.series[1].graphicalProperties.solidFill = "8EA9DB"
    bar.series[1].tx.strRef = None
    from openpyxl.chart.series import SeriesLabel
    bar.series[1].tx = SeriesLabel(v="52-week EV/EBITDA range")
    bar.series[0].tx = SeriesLabel(v=" ")
    line = LineChart()
    line.add_data(Reference(ws, min_col=4, min_row=4, max_row=rl), titles_from_data=True)
    line.add_data(Reference(ws, min_col=6, min_row=4, max_row=rl), titles_from_data=True)
    cur, mdn = line.series
    cur.marker = Marker(symbol="diamond", size=9)
    cur.marker.graphicalProperties.solidFill = "1F3864"
    cur.graphicalProperties.line.noFill = True
    mdn.graphicalProperties.line.solidFill = "C00000"
    mdn.graphicalProperties.line.dashStyle = "dash"
    mdn.graphicalProperties.line.width = 22000
    mdn.marker = Marker(symbol="none")
    bar += line
    bar.title = "EV / EBITDA (FY26A): 52-week range, current and peer median"
    bar.y_axis.title = "EV / EBITDA (x)"
    bar.y_axis.numFmt = '0.0"x"'
    bar.y_axis.majorGridlines = None
    bar.x_axis.delete = False
    bar.y_axis.delete = False
    bar.legend.position = "b"
    bar.height, bar.width = 11, 24
    ws.add_chart(bar, f"A{rl + 3}")
    ws.cell(rl + 2, 1, "Bars: EV/EBITDA at the 52-week closing low/high (diluted shares and FY26 net debt held constant). "
                       "Diamond: current. Dashed line: peer median (excludes companies marked excl.).").font = \
        Font(italic=True, color="808080")
    return ws


def build(cos, policy):
    wb = Workbook()
    banner = data_dates(cos)
    ncol = FIRST_COL + len(cos) - 1
    L = lambda i: get_column_letter(FIRST_COL + i)          # company column letter
    cell_map, audit = {}, []

    # ---------------- Policy & Flags (first so other sheets can reference it)
    wp = wb.active
    wp.title = "Policy & Flags"
    wp["A1"], wp["A1"].font = "Policy assumptions & flags", Font(bold=True, size=14)
    wp["A3"], wp["B3"] = "Tax rate applied to exceptional items when not disclosed", policy["pe"]["exceptional_tax_rate_default"]
    wp["B3"].font, wp["B3"].number_format = BLUE, "0.000%"
    wp["A4"], wp["B4"] = "NCI share of discontinued profit when not disclosed", policy["pe"]["discontinued_nci_share_default"]
    wp["B4"].font, wp["B4"].number_format = BLUE, "0.0%"
    wp["A5"], wp["B5"] = "Multiples above this are shown as NM", NM_CAP
    wp["B5"].font = BLUE
    TAXR, DNCI, CAP = "'Policy & Flags'!$B$3", "'Policy & Flags'!$B$4", "'Policy & Flags'!$B$5"
    r = 7
    wp.cell(r, 1, "Methodology, data dates and known limitations: see the 'Methodology & Limitations' tab.").font = BOLD
    r += 2
    wp.cell(r, 1, "Flags (from extraction, market data and share bridge)").font = BOLD
    for c in cos:
        for adj in c["mk"].get("share_adjustments", []):
            r += 1
            state = "INCLUDED" if adj["include"] else "EXCLUDED (OPEN decision)"
            wp.cell(r, 1, f"{c['peer']['name']}: share bridge {state}: {adj['label']} "
                          f"{adj['shares']:,} shares - {adj['source']}")
        r += 1
        if c["mk"].get("shares_override"):
            cell = wp.cell(r, 1, f"{c['peer']['name']}: BASIC SHARES = {c['mk']['shares_override']}: "
                                 f"{c['mk']['basic_shares']:,} - {c['mk']['shares_override_reason']}")
            cell.font = Font(bold=True, color="C00000")
        else:
            wp.cell(r, 1, f"{c['peer']['name']}: yfinance share count {c['mk']['shares_yf']:,} reconciles to AR + bridge "
                          f"within {c['mk']['yf_unexplained_residual']:+.2%}")
        for fl in c["ex"]["flags"]:
            r += 1
            wp.cell(r, 1, f"{c['peer']['name']}: {fl}").alignment = Alignment(wrap_text=True, vertical="top")
    wp.column_dimensions["A"].width = 150
    wp.column_dimensions["B"].width = 12

    # ---------------- Inputs
    wi = wb.create_sheet("Inputs")
    wi["A1"], wi["A1"].font = "Inputs - as printed in the Annual Report (blue = hard-coded input)", Font(bold=True, size=13)
    date_banner(wi, banner, ncol)
    hdr_row(wi, 3, ["Line item", "Unit"] + [c["peer"]["name"] for c in cos])
    IR = {}
    r = 4
    for key, label, unit, getter in INPUT_ROWS:
        if key == "hdr":
            section(wi, r, label, ncol)
            r += 1
            continue
        IR[key] = r
        wi.cell(r, 1, label)
        wi.cell(r, 2, {"rep": "as printed"}.get(unit, unit))
        for i, c in enumerate(cos):
            raw = getter(c["ex"], c["mk"])
            val = v(raw) if isinstance(raw, dict) else raw
            cell = wi.cell(r, FIRST_COL + i, val)
            cell.font = BLUE
            cell.number_format = NUM0 if key in ("basic_shares", "sub_shares") else (
                "0.00%" if key == "sub_nci_pct" else (NUM if key != "price_date" else "@"))
            if key == "unit_factor":
                cell.number_format = "0.0000"
            cm = cite_comment(raw, c["ex"].get("_pdf", c["manifest"]["pdf"]))
            if cm:
                cell.comment = cm
            src = raw if isinstance(raw, dict) else None
            audit.append((c["peer"]["name"], "Inputs", label, val, c["ex"]["unit_as_printed"] if unit == "rep" else unit,
                          c["manifest"]["pdf"], src, f"Inputs!{L(i)}{r}"))
        r += 1
    wi.column_dimensions["A"].width = 52
    wi.column_dimensions["B"].width = 11
    for i in range(len(cos)):
        wi.column_dimensions[L(i)].width = 18
    wi.freeze_panes = "C4"

    # ---------------- Components
    wc = wb.create_sheet("Components")
    wc["A1"], wc["A1"].font = "Components - every cash / investment / debt-like / ESOP line (values as printed)", Font(bold=True, size=13)
    hdr_row(wc, 3, ["Company", "Category", "Line item (as printed)", "Value / count", "Include?",
                    "Exercise price (INR)", "TSM dilutive shares", "Printed page", "Slice"])
    r = 4
    price_row = IR["price"]
    first_comp = r
    for c in cos:
        for cat, label, val, inc, strike, cite in components(c, policy):
            wc.cell(r, 1, c["peer"]["name"])
            wc.cell(r, 2, cat)
            wc.cell(r, 3, label)
            vc = wc.cell(r, 4, val)
            vc.font, vc.number_format = BLUE, NUM0 if cat == "esop" else NUM
            ic = wc.cell(r, 5, bool(inc))
            ic.font = BLUE
            if cat == "esop":
                sc = wc.cell(r, 6, strike)
                sc.font, sc.number_format = BLUE, NUM
                price = f"INDEX(Inputs!$C${price_row}:${get_column_letter(ncol)}${price_row},MATCH(A{r},Inputs!$C$3:${get_column_letter(ncol)}$3,0))"
                wc.cell(r, 7, f"=IF(AND(E{r},F{r}<{price}),D{r}*(1-F{r}/{price}),0)").number_format = NUM0
            wc.cell(r, 8, cite.get("printed_page"))
            wc.cell(r, 9, cite.get("uid"))
            audit.append((c["peer"]["name"], f"Components/{cat}", label, val,
                          "count" if cat == "esop" else c["ex"]["unit_as_printed"], c["manifest"]["pdf"], cite, f"Components!D{r}"))
            r += 1
    last_comp = max(r - 1, first_comp)
    for col, w in zip("ABCDEFGHI", (20, 12, 70, 16, 9, 14, 16, 11, 9)):
        wc.column_dimensions[col].width = w
    wc.freeze_panes = "A4"
    RNG = lambda col: f"Components!${col}${first_comp}:${col}${last_comp}"

    def sumifs(cat, co_cell, col="D"):
        return f'SUMIFS({RNG(col)},{RNG("A")},{co_cell},{RNG("B")},"{cat}",{RNG("E")},TRUE)'

    # ---------------- Standardised (INR Cr)
    wsd = wb.create_sheet("Standardised")
    wsd["A1"], wsd["A1"].font = "Standardised financials - INR Crore (all formulas)", Font(bold=True, size=13)
    date_banner(wsd, banner, ncol)
    hdr_row(wsd, 3, ["Line item", "Unit"] + [c["peer"]["name"] for c in cos])
    SR = {}
    std_rows = [
        ("hdr", "Operating performance (continuing operations)"),
        ("revenue", "Revenue from operations", lambda c, q: f"=Inputs!{c}{IR['revenue']}*{q}"),
        ("fx_total", "  FX gain/(loss) inside Other income - all lines (reported unit)",
         lambda c, q: f"=N(Inputs!{c}{IR['fx_in_oi']})+{sumifs('fx_oi', c + '$3')}"),
        ("oi_nonop", "Non-operating Other income (stripped)", lambda c, q: f"=(Inputs!{c}{IR['other_income']}-{c}{{fx_total}})*{q}"),
        ("ebitda", "Adjusted EBITDA", lambda c, q: (f"=(Inputs!{c}{IR['pbt']}+Inputs!{c}{IR['finance_costs']}+Inputs!{c}{IR['dep_amort']}"
                                                   f"+N(Inputs!{c}{IR['exceptional']})-Inputs!{c}{IR['other_income']}+{c}{{fx_total}}"
                                                   f"-N(Inputs!{c}{IR['share_eq']}))*{q}")),
        ("ebitda_chk", "  check: Revenue - operating expenses + FX in OI - EBITDA (should be 0)",
         lambda c, q: (f"=(Inputs!{c}{IR['revenue']}-(Inputs!{c}{IR['total_expenses']}-Inputs!{c}{IR['finance_costs']}"
                       f"-Inputs!{c}{IR['dep_amort']})+{c}{{fx_total}})*{q}-{c}{{ebitda}}")),
        ("margin", "Adjusted EBITDA margin", lambda c, q: f"=IF({c}{{revenue}}=0,0,{c}{{ebitda}}/{c}{{revenue}})"),
        ("hdr", "Earnings"),
        ("rep_pat", "Reported PAT attributable to owners", lambda c, q: f"=Inputs!{c}{IR['pat_owners']}*{q}"),
        ("disc_own", "Less: discontinued ops attributable to owners",
         lambda c, q: f"=IF(Inputs!{c}{IR['disc_owners']}=\"\",N(Inputs!{c}{IR['pat_disc']})*(1-{DNCI}),Inputs!{c}{IR['disc_owners']})*{q}"),
        ("exc_post_tax", "Add: exceptional charge, post-tax",
         lambda c, q: (f"=(N(Inputs!{c}{IR['exceptional']})-IF(Inputs!{c}{IR['exc_tax']}=\"\","
                       f"N(Inputs!{c}{IR['exceptional']})*{TAXR},Inputs!{c}{IR['exc_tax']}))*{q}")),
        ("adj_pat", "Adjusted PAT (continuing, ex-exceptional, owners)", lambda c, q: f"={c}{{rep_pat}}-{c}{{disc_own}}+{c}{{exc_post_tax}}"),
        ("hdr", "Cash & liquid investments"),
        ("cash_unres", "Unrestricted cash & bank balances", lambda c, q: f"={sumifs('cash', c + '$3')}*{q}"),
        ("carve", "Less: restricted carve-outs (narrative)", lambda c, q: f"={sumifs('carveout', c + '$3')}*{q}"),
        ("treasury", "Treasury investments (current + non-current)", lambda c, q: f"={sumifs('investment', c + '$3')}*{q}"),
        ("cash_total", "Total cash & liquid investments", lambda c, q: f"={c}{{cash_unres}}-{c}{{carve}}+{c}{{treasury}}"),
        ("cash_chk", "  check: cash components - face cash & bank (reported unit, should be 0)",
         lambda c, q: (f'=SUMIFS({RNG("D")},{RNG("A")},{c}$3,{RNG("B")},"cash")-N(Inputs!{c}{IR["cash_face"]})-N(Inputs!{c}{IR["obb_face"]})')),
        ("inv_chk", "  check: investment components - face investments (reported unit, should be 0)",
         lambda c, q: (f'=SUMIFS({RNG("D")},{RNG("A")},{c}$3,{RNG("B")},"investment")-N(Inputs!{c}{IR["ci_face"]})-N(Inputs!{c}{IR["nci_face"]})')),
        ("hdr", "Debt & debt-like"),
        ("borrowings", "Borrowings (current + non-current)", lambda c, q: f"=(N(Inputs!{c}{IR['borr_nc']})+N(Inputs!{c}{IR['borr_c']}))*{q}"),
        ("leases", "Lease liabilities - Ind AS 116 (current + non-current)", lambda c, q: f"=(N(Inputs!{c}{IR['lease_nc']})+N(Inputs!{c}{IR['lease_c']}))*{q}"),
        ("debt_like", "Debt-like items (earn-outs, deferred consideration, put options)", lambda c, q: f"={sumifs('debt_like', c + '$3')}*{q}"),
        ("nci", "Non-controlling interests (book; listed subsidiaries at market value)",
         lambda c, q: (f"=(N(Inputs!{c}{IR['nci_book']})-N(Inputs!{c}{IR['sub_nci_book']}))*{q}"
                       f"+N(Inputs!{c}{IR['sub_price']})*N(Inputs!{c}{IR['sub_shares']})*N(Inputs!{c}{IR['sub_nci_pct']})/10^7")),
        ("eq_method", "Investments in associates / JVs", lambda c, q: f"=N(Inputs!{c}{IR['eq_method']})*{q}"),
    ]
    r = 4
    pending = []
    for row in std_rows:
        if row[0] == "hdr":
            section(wsd, r, row[1], ncol)
            r += 1
            continue
        key, label, fn = row
        SR[key] = r
        wsd.cell(r, 1, label)
        wsd.cell(r, 2, "%" if key == "margin" else ("as printed" if key.endswith("_chk") and key != "ebitda_chk" else "INR Cr"))
        pending.append((r, key, fn))
        r += 1
    for r_, key, fn in pending:
        for i in range(len(cos)):
            c = L(i)
            q = f"Inputs!{c}{IR['unit_factor']}"
            f = fn(c, q)
            for k, rr in SR.items():
                f = f.replace(f"{{{k}}}", str(rr))
            cell = wsd.cell(r_, FIRST_COL + i, f)
            cell.number_format = PCT if key == "margin" else NUM
            if key in ("ebitda", "adj_pat", "cash_total"):
                cell.font = BOLD
            if key.endswith("_chk"):
                cell.font = Font(italic=True, color="808080")
            cell_map[f"std.{key}.{cos[i]['co']}"] = f"Standardised!{c}{r_}"
    wsd.column_dimensions["A"].width = 66
    wsd.column_dimensions["B"].width = 11
    for i in range(len(cos)):
        wsd.column_dimensions[L(i)].width = 18
    wsd.freeze_panes = "C4"

    # ---------------- EV Bridge
    we = wb.create_sheet("EV Bridge")
    we["A1"], we["A1"].font = "Enterprise value bridge - INR Crore (all formulas)", Font(bold=True, size=13)
    date_banner(we, banner, ncol)
    hdr_row(we, 3, ["Bridge item", "Unit"] + [c["peer"]["name"] for c in cos])
    ER = {}
    ev_rows = [
        ("price", "Share price", "INR", lambda c: f"=Inputs!{c}{IR['price']}"),
        ("basic", "Basic shares", "mn", lambda c: f"=Inputs!{c}{IR['basic_shares']}/10^6"),
        ("tsm", "Add: ESOP/RSU dilution (treasury stock method)", "mn",
         lambda c: f'=SUMIFS({RNG("G")},{RNG("A")},{c}$3,{RNG("B")},"esop")/10^6'),
        ("diluted", "Diluted shares", "mn", lambda c: f"={c}{{basic}}+{c}{{tsm}}"),
        ("eqv", "Diluted equity value (market cap)", "INR Cr", lambda c: f"={c}{{price}}*{c}{{diluted}}/10"),
        ("borr", "Add: borrowings", "INR Cr", lambda c: f"=Standardised!{c}{SR['borrowings']}"),
        ("lease", "Add: lease liabilities (Ind AS 116)", "INR Cr", lambda c: f"=Standardised!{c}{SR['leases']}"),
        ("dlike", "Add: debt-like items", "INR Cr", lambda c: f"=Standardised!{c}{SR['debt_like']}"),
        ("nci", "Add: non-controlling interests", "INR Cr", lambda c: f"=Standardised!{c}{SR['nci']}"),
        ("cash", "Less: unrestricted cash & treasury investments", "INR Cr", lambda c: f"=-Standardised!{c}{SR['cash_total']}"),
        ("eqm", "Less: associates / JVs", "INR Cr", lambda c: f"=-Standardised!{c}{SR['eq_method']}"),
        ("ev", "Enterprise value", "INR Cr", lambda c: f"=SUM({c}{{eqv}}:{c}{{eqm}})"),
    ]
    r = 4
    for key, label, unit, _ in ev_rows:
        ER[key] = r
        we.cell(r, 1, label).font = BOLD if key in ("eqv", "ev") else Font()
        we.cell(r, 2, unit)
        r += 1
    for key, label, unit, fn in ev_rows:
        for i in range(len(cos)):
            c = L(i)
            f = fn(c)
            for k, rr in ER.items():
                f = f.replace(f"{{{k}}}", str(rr))
            cell = we.cell(ER[key], FIRST_COL + i, f)
            cell.number_format = NUM
            if key in ("eqv", "ev"):
                cell.font = BOLD
                cell.border = THIN
            cell_map[f"ev.{key}.{cos[i]['co']}"] = f"'EV Bridge'!{c}{ER[key]}"
    we.column_dimensions["A"].width = 48
    for i in range(len(cos)):
        we.column_dimensions[L(i)].width = 18
    we.freeze_panes = "C4"

    # ---------------- Comps
    wk = wb.create_sheet("Comps", 0)
    wk["A1"], wk["A1"].font = "Indian Mid-Tier IT Services - Trading Comparables (FY26A)", Font(bold=True, size=14)
    date_banner(wk, banner, 20)
    wk["A3"] = "INR Crore unless stated. All figures are live formulas from the Inputs, Components, Standardised and EV Bridge tabs."
    wk["A3"].font = Font(italic=True, color="808080")
    cols = ["Company", "Ticker", "Price (INR)", "Price date", "Diluted equity value", "Enterprise value", "Revenue",
            "Adj. EBITDA", "EBITDA margin", "Adj. PAT", "EV / Revenue", "EV / EBITDA", "P / E (adjusted)", "P / E (reported, memo)",
            "In peer stats?", "Status / note",
            "stats: EV/Rev", "stats: EV/EBITDA", "stats: P/E adj", "stats: P/E rep"]
    hdr_row(wk, 4, cols)
    r0 = 5
    STAT_COLS = {11: 17, 12: 18, 13: 19, 14: 20}       # multiple column -> helper column feeding the statistics
    for i, c in enumerate(cos):
        r, cl = r0 + i, L(i)
        incl = not c["peer"].get("exclude_from_stats", False)
        ic = wk.cell(r, 15, "Yes" if incl else "No")
        ic.font, ic.alignment = BLUE, Alignment(horizontal="center")
        note = wk.cell(r, 16, c["peer"].get("stats_note", ""))
        note.alignment = Alignment(wrap_text=False)
        if not incl:
            for col in range(1, 15):
                wk.cell(r, col).fill = PatternFill("solid", fgColor="FFF2CC")
        for mcol, hcol in STAT_COLS.items():
            ml = get_column_letter(mcol)
            hc = wk.cell(r, hcol, f'=IF($O{r}="Yes",{ml}{r},"excl.")')
            hc.number_format, hc.font = MULT, Font(color="808080")
        wk.cell(r, 1, f"=Inputs!{cl}$3")
        wk.cell(r, 2, c["mk"]["ticker"])
        wk.cell(r, 3, f"='EV Bridge'!{cl}{ER['price']}").number_format = NUM
        wk.cell(r, 4, f"=Inputs!{cl}{IR['price_date']}")
        wk.cell(r, 5, f"='EV Bridge'!{cl}{ER['eqv']}").number_format = NUM0
        wk.cell(r, 6, f"='EV Bridge'!{cl}{ER['ev']}").number_format = NUM0
        wk.cell(r, 7, f"=Standardised!{cl}{SR['revenue']}").number_format = NUM0
        wk.cell(r, 8, f"=Standardised!{cl}{SR['ebitda']}").number_format = NUM0
        wk.cell(r, 9, f"=Standardised!{cl}{SR['margin']}").number_format = PCT
        wk.cell(r, 10, f"=Standardised!{cl}{SR['adj_pat']}").number_format = NUM0
        rep = f"Standardised!{cl}{SR['rep_pat']}"
        for col, num, den in ((11, f"F{r}", f"G{r}"), (12, f"F{r}", f"H{r}"), (13, f"E{r}", f"J{r}"), (14, f"E{r}", rep)):
            cell = wk.cell(r, col, f'=IFERROR(IF(OR({den}<=0,{num}/{den}<0,{num}/{den}>{CAP}),"NM",{num}/{den}),"NM")')
            cell.number_format = MULT
            cell.alignment = Alignment(horizontal="right")
        for k, col in (("ev_rev", 11), ("ev_ebitda", 12), ("pe_adj", 13), ("pe_rep", 14)):
            cell_map[f"comps.{k}.{c['co']}"] = f"Comps!{get_column_letter(col)}{r}"
        cell_map[f"comps.ev.{c['co']}"] = f"Comps!F{r}"
        cell_map[f"comps.eqv.{c['co']}"] = f"Comps!E{r}"
    rl = r0 + len(cos) - 1
    rs = rl + 2
    for j, (name, fn) in enumerate((("Mean", "AVERAGE({rng})"), ("Median", "MEDIAN({rng})"),
                                     ("25th percentile", "PERCENTILE({rng},0.25)"), ("75th percentile", "PERCENTILE({rng},0.75)"))):
        r = rs + j
        wk.cell(r, 1, name).font = BOLD
        for col, k in ((11, "ev_rev"), (12, "ev_ebitda"), (13, "pe_adj"), (14, "pe_rep")):
            cl = get_column_letter(col)
            hl = get_column_letter(STAT_COLS[col])
            rng = f"{hl}{r0}:{hl}{rl}"        # helper column: excluded / NM rows are text and are ignored
            cell = wk.cell(r, col, f'=IF(COUNT({rng})=0,"n/a",{fn.format(rng=rng)})')
            cell.number_format, cell.font = MULT, BOLD
            if j == 0:
                cell.border = THIN
            cell_map[f"stats.{name}.{k}"] = f"Comps!{cl}{r}"
        wk.cell(r, 1).border = THIN if j == 0 else Border()
    n_in = sum(1 for c in cos if not c["peer"].get("exclude_from_stats", False))
    wk.cell(r, 15, "n =").font = BOLD
    wk.cell(r, 16, f'=COUNTIF(O{r0}:O{rl},"Yes")').font = BOLD
    wk.cell(rs + 5, 1, f"Benchmark statistics use only rows with 'In peer stats?' = Yes ({n_in} of {len(cos)}) and exclude NM "
                       f"values (negative/zero denominators or multiples > {NM_CAP}x). Highlighted rows are shown for "
                       f"reference only - see Status / note.").font = Font(italic=True, color="808080")
    if n_in < 5:
        wk.cell(rs + 6, 1, f"WARNING: only {n_in} compan{'y' if n_in == 1 else 'ies'} in the peer statistics - "
                           f"not a meaningful benchmark.").font = Font(bold=True, color="C00000")
    for col, w in zip(range(1, 21), (24, 14, 11, 11, 14, 14, 12, 12, 10, 12, 11, 11, 12, 13, 9, 60, 10, 10, 10, 10)):
        wk.column_dimensions[get_column_letter(col)].width = w
    wk.row_dimensions[4].height = 32
    wk.freeze_panes = "B5"

    # ---------------- Football field + Methodology (placed right after Comps)
    add_methodology(wb, cos, policy, banner)
    add_football_field(wb, cos, IR, ER, SR, L, cell_map, banner)

    # ---------------- Audit Log
    wa = wb.create_sheet("Audit Log")
    wa["A1"], wa["A1"].font = "Audit log - source of every hard-coded input", Font(bold=True, size=13)
    hdr_row(wa, 3, ["Company", "Sheet / category", "Line item", "Value", "Unit", "PDF", "PDF page", "Half",
                    "Printed page", "Label as printed", "Note", "Slice file", "Workbook cell"])
    r = 4
    for co_name, where, label, val, unit, pdf, src, addr in audit:
        uid = src.get("uid") if src else None
        row = [co_name, where, label, val, unit, pdf if uid else "-", int(uid[1:4]) if uid else None,
               {"L": "left", "R": "right", "F": "full"}.get(uid[-1]) if uid else None,
               src.get("printed_page") if src else None, src.get("label") if src else
               ("market data / config - see Policy & Flags" if val is not None else "not presented in report"),
               src.get("note") if src else None, f"work/{{co}}/pages/{uid}.txt" if uid else None, addr]
        for j, x in enumerate(row):
            wa.cell(r, j + 1, x)
        wa.cell(r, 4).number_format = NUM
        r += 1
    for c in cos:   # market data provenance
        mk = c["mk"]
        for label, val in (("Share price", f"{mk['price_inr']} on {mk['price_date']} - {mk['price_source']} ({mk['ticker']})"),
                           ("AR year-end shares", f"{mk['shares_ar_year_end']:,} - {mk['shares_ar_cite']}"),
                           ("yfinance share count (check only)", f"{mk['shares_yf']:,}; unexplained residual {mk['yf_unexplained_residual']:+.2%}"),
                           ("Market data fetched", mk["fetched_at"])):
            wa.cell(r, 1, c["peer"]["name"]); wa.cell(r, 2, "Market data"); wa.cell(r, 3, label); wa.cell(r, 10, val)
            r += 1
    for col, w in zip("ABCDEFGHIJKLM", (20, 20, 52, 16, 10, 18, 9, 7, 11, 70, 8, 30, 16)):
        wa.column_dimensions[col].width = w
    wa.freeze_panes = "A4"
    wa.auto_filter.ref = f"A3:M{r - 1}"
    # fix slice paths with company ids
    for row in wa.iter_rows(min_row=4, max_row=r - 1):
        if row[11].value and "{co}" in row[11].value:
            co_id = next(c["co"] for c in cos if c["peer"]["name"] == row[0].value)
            row[11].value = row[11].value.replace("{co}", co_id)
    return wb, cell_map


def main(argv):
    policy = load_yaml("policy.yaml")
    peers = [p["co_id"] for p in load_yaml("peers.yaml")["peers"]]
    co_ids = argv or [p for p in peers if (WORK_DIR / p / "extracted.json").exists() and (WORK_DIR / p / "market.json").exists()]
    cos = load_companies(co_ids)
    wb, cell_map = build(cos, policy)
    OUTPUT_DIR.mkdir(exist_ok=True)
    path = OUTPUT_DIR / "trading_comps.xlsx"
    try:
        wb.save(path)
    except PermissionError:          # file open in Excel -> write alongside, never lose the build
        import datetime as dt
        path = OUTPUT_DIR / f"trading_comps_{dt.datetime.now():%Y%m%d_%H%M}.xlsx"
        wb.save(path)
        print(f"NOTE: output/trading_comps.xlsx is open/locked - wrote {path.name} instead")

    sh = {c["co"]: shadow(c, policy) for c in cos}
    stats = {}
    for k in ("ev_rev", "ev_ebitda", "pe_adj", "pe_rep"):
        xs = [sh[c["co"]][k] for c in cos if sh[c["co"]][k] != "NM" and not c["peer"].get("exclude_from_stats", False)]
        stats[k] = {"Mean": statistics.fmean(xs) if xs else "n/a", "Median": statistics.median(xs) if xs else "n/a",
                    "25th percentile": pctl(xs, 0.25) if xs else "n/a", "75th percentile": pctl(xs, 0.75) if xs else "n/a"}
    cells = {}
    for co, x in sh.items():
        cells[f"ff.low.{co}"], cells[f"ff.high.{co}"] = x["ff_low"], x["ff_high"]
        cells[f"std.nci.{co}"] = x["nci"]
    write_json(WORK_DIR / "comps_shadow.json", {"companies": sh, "stats": stats, "cells": cells,
                                              "workbook": path.relative_to(ROOT).as_posix()})  # root-relative: portable
    write_json(WORK_DIR / "cellmap.json", cell_map)
    print(f"wrote {path} ({len(cos)} companies)")
    for co, s in sh.items():
        fm = lambda x: x if x == "NM" else f"{x:.1f}x"
        print(f"  {co:10s} EqV {s['equity_value']:>10,.0f}  EV {s['ev']:>10,.0f}  Rev {s['revenue']:>9,.0f}  "
              f"EBITDA {s['ebitda']:>8,.0f} ({s['ebitda'] / s['revenue']:.1%})  AdjPAT {s['adj_pat']:>7,.0f}  "
              f"EV/Rev {fm(s['ev_rev'])}  EV/EBITDA {fm(s['ev_ebitda'])}  P/E {fm(s['pe_adj'])} (rep {fm(s['pe_rep'])})")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
