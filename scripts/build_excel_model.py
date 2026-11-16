"""Build model/UPS_vs_FedEx_Valuation_Model.xlsx from the CSVs in /data.

Every number in the workbook is either a sourced input (blue) or a live Excel formula
(black, green when it links to another sheet). Run:  python scripts/build_excel_model.py
"""
from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as CL

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from src.valuation import DRIVER_SHOCKS, STRATEGIES  # noqa: E402

DATA = ROOT / "data"
OUT = ROOT / "model" / "UPS_vs_FedEx_Valuation_Model.xlsx"

# ------------------------------------------------------------------ styling
FONT = "Arial"
BLUE, BLACK, GREEN, WHITE, GREY, NAVY = "0000FF", "000000", "008000", "FFFFFF", "7F7F7F", "1F3864"
UPS_BROWN, FDX_PURPLE = "5C3D2E", "4D148C"
YELLOW = PatternFill("solid", fgColor="FFFF00")
HDR = PatternFill("solid", fgColor=NAVY)
SEC = PatternFill("solid", fgColor="D9E1F2")
LIGHT = PatternFill("solid", fgColor="F2F2F2")
thin = Side(style="thin", color="BFBFBF")
BOX = Border(top=thin, bottom=thin, left=thin, right=thin)
TOP = Border(top=Side(style="thin", color="000000"))
DTOP = Border(top=Side(style="thin", color="000000"), bottom=Side(style="double", color="000000"))

USD = '$#,##0;($#,##0);"-"'
PS = '$#,##0.00;($#,##0.00);"-"'
PCT = '0.0%;(0.0%);"-"'
PCT2 = '0.00%;(0.00%);"-"'
MULT = '0.00"x";(0.00"x");"-"'
DAYS = '0.0;(0.0);"-"'
NUM1 = '#,##0.0;(#,##0.0);"-"'
IDX = '0.00'


def style(c, color=BLACK, bold=False, fmt=None, fill=None, size=10, italic=False, align=None, wrap=False):
    c.font = Font(name=FONT, color=color, bold=bold, size=size, italic=italic)
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    if align or wrap:
        c.alignment = Alignment(horizontal=align, vertical="top" if wrap else "center", wrap_text=wrap)
    return c


def put(ws, ref, val, color=BLACK, **kw):
    c = ws[ref]
    c.value = val
    return style(c, color, **kw)


def title(ws, text, sub):
    put(ws, "A1", text, NAVY, bold=True, size=14)
    put(ws, "A2", sub, GREY, italic=True, size=9)
    ws.sheet_view.showGridLines = False


def header(ws, row, col, labels, height=None):
    for i, lab in enumerate(labels):
        c = ws.cell(row=row, column=col + i, value=lab)
        style(c, WHITE, bold=True, fill=HDR, align="center", wrap=True)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    if height:
        ws.row_dimensions[row].height = height


def section(ws, row, text, span=10, col=1):
    for c in range(col, col + span):
        ws.cell(row=row, column=c).fill = SEC
    style(ws.cell(row=row, column=col, value=text), NAVY, bold=True)


def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w


def note(ws, ref, text):
    ws[ref].comment = Comment(text, "Model")


def tidy_chart(ch, ttl, fmt=None):
    ch.title = ttl
    ch.roundedCorners = False
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.y_axis.majorGridlines = None
    if fmt:
        ch.y_axis.numFmt = fmt
    if ch.legend is not None:
        ch.legend.position = "b"
        ch.legend.overlay = False
    ch.height, ch.width = 7.5, 15


# ------------------------------------------------------------------- data
CO = {
    "UPS": {"sheet": "Fin_UPS", "csv": "ups_financials.csv", "name": "UPS", "col": "B", "color": UPS_BROWN, "first_fy": 2026},
    "FDX": {"sheet": "Fin_FDX", "csv": "fedex_financials.csv", "name": "FedEx", "col": "C", "color": FDX_PURPLE, "first_fy": 2027},
}
assump = pd.read_csv(DATA / "valuation_assumptions.csv")
tsr = pd.read_csv(DATA / "tsr_index.csv")

wb = Workbook()
ws_cover = wb.active
ws_cover.title = "Cover"
names = ["Summary", "Assumptions", "Fin_UPS", "Fin_FDX", "Ex01_Returns", "Ex02_Ratios", "Ex03_DuPont", "Ex04_FCF",
         "Ex05_Forecast", "Ex06_WACC", "Ex07_Valuation", "Ex08_Value_Drivers", "Ex09_Strategy", "DCF_Engine"]
S = {n: wb.create_sheet(n) for n in names}

# ------------------------------------------------------- financial sheets
FROW = {}   # FROW[ticker][line_item] -> row
YCOL = {"UPS": {}, "FDX": {}}
for t, meta in CO.items():
    ws = S[meta["sheet"]]
    df = pd.read_csv(DATA / meta["csv"])
    years = [c for c in df.columns if c.startswith("FY")]
    fy_note = "Fiscal years end December 31" if t == "UPS" else "Fiscal years end May 31"
    title(ws, f"{meta['name']} - Financial statements from Form 10-K ($ millions unless noted)", f"{fy_note} | Blue = reported data | Source column cites the filing")
    header(ws, 4, 1, ["Statement", "Line item"] + years + ["Source"], 30)
    widths(ws, {"A": 10, "B": 42, "C": 12, "D": 12, "E": 12, "F": 80})
    FROW[t] = {}
    labels = {"IS": "Income", "BS": "Balance", "CF": "Cash flow", "MKT": "Market"}
    for i, rec in df.iterrows():
        r = 5 + i
        FROW[t][rec["line_item"]] = r
        put(ws, f"A{r}", labels[rec["statement"]], GREY, size=9)
        is_total = rec["line_item"].startswith("Total") or rec["line_item"] in ("Operating income (EBIT)", "Net income", "Revenue")
        put(ws, f"B{r}", rec["line_item"], bold=is_total)
        for j, y in enumerate(years):
            v = rec[y]
            fmt = PS if "($)" in rec["line_item"] else USD
            if pd.notna(v):
                put(ws, f"{CL(3 + j)}{r}", float(v), BLUE, fmt=fmt, bold=is_total)
        put(ws, f"F{r}", rec["source"], GREY, size=8)
        if rec["statement"] != df.loc[i - 1, "statement"] if i > 0 else False:
            for c in range(1, 7):
                ws.cell(row=r, column=c).border = TOP
    YCOL[t] = {y: CL(3 + j) for j, y in enumerate(years)}
    # balance-sheet check
    r = 5 + len(df) + 1
    put(ws, f"B{r}", "Check: assets = liabilities + equity", GREY, italic=True)
    for y, col in YCOL[t].items():
        liab = "+".join(f"{col}{FROW[t][k]}" for k in ["Total current liabilities", "Long-term debt", "Non-current operating lease liabilities",
                                                       "Pension and postretirement obligations", "Deferred income tax liabilities",
                                                       "Other non-current liabilities", "Total shareholders equity"])
        put(ws, f"{col}{r}", f'=IF(ABS({col}{FROW[t]["Total assets"]}-({liab}))<1,"OK","CHECK")', GREY, bold=True, align="center")
    ws.freeze_panes = "C5"


def F(t, item, year_idx=-1, sheet_prefix=True):
    """Reference to a financial-sheet cell. year_idx: 0,1,2 or -1 (latest)."""
    cols = list(YCOL[t].values())
    ref = f"${cols[year_idx]}${FROW[t][item]}"
    return f"{CO[t]['sheet']}!{ref}" if sheet_prefix else ref


# ------------------------------------------------------------ Assumptions
a = S["Assumptions"]
title(a, "Assumptions - market inputs, forecast drivers and scenario shocks", "Blue = input you can change | Yellow = key judgment calls | Every formula in the model reads from here")
widths(a, {"A": 44, "B": 13, "C": 13, "D": 8, "E": 100})
header(a, 4, 1, ["Parameter", "UPS", "FedEx", "Unit", "Source / note"], 28)
AROW = {}
key_params = {"Operating margin", "Terminal growth rate", "Revenue growth Y2", "Revenue growth Y3", "Revenue growth Y4", "Revenue growth Y5", "Capex % of revenue Y2-Y5"}
for i, rec in assump.iterrows():
    r = 5 + i
    AROW[rec["parameter"]] = r
    put(a, f"A{r}", rec["parameter"])
    for col, t in (("B", "UPS"), ("C", "FDX")):
        raw = rec[t]
        unit = rec["unit"]
        if unit == "date":
            put(a, f"{col}{r}", raw, BLUE, align="center")
            continue
        v = float(raw)
        fmt = {"%": PCT2, "$": PS, "x": IDX, "millions": NUM1}[unit]
        put(a, f"{col}{r}", v, BLUE, fmt=fmt, align="right", fill=YELLOW if rec["parameter"] in key_params else None)
    put(a, f"D{r}", rec["unit"], GREY, size=9, align="center")
    put(a, f"E{r}", rec["note"], GREY, size=9)

r0 = 5 + len(assump) + 1
section(a, r0, "Value-driver shocks: each driver is moved by one percentage point, one at a time (Exhibit 08)", 5)
header(a, r0 + 1, 1, ["Driver shock", "Change", "", "", "How it is applied"])
SHOCK = {}  # name -> (row, {param: cell})
param_label = {"growth": "every year's revenue growth", "op_margin": "operating margin", "tax_rate": "tax rate",
               "capex_y1": "capex % (year 1)", "capex_y2_5": "capex % (years 2-5)", "nwc_pct": "NWC % of revenue change",
               "wacc": "WACC", "terminal_g": "terminal growth"}
r = r0 + 2
for name, d in DRIVER_SHOCKS:
    put(a, f"A{r}", name)
    val = list(d.values())[0]
    put(a, f"B{r}", val, BLUE, fmt=PCT2, align="right")
    put(a, f"E{r}", "Applied to " + " and ".join(param_label[k] for k in d), GREY, size=9)
    SHOCK[name] = (r, {k: f"Assumptions!$B${r}" for k in d})
    r += 1

r += 1
section(a, r, "Strategy scenarios: each layer adds to the one before it (Exhibit 09)", 5)
header(a, r + 1, 1, ["Scenario", "Change", "Company", "", "Driver changed"])
r += 2
STRAT = {"UPS": [], "FDX": []}
strat_notes = {
    "S1: Network reconfiguration & automation": "UPS closed 93 buildings and cut ~48,000 operational positions in 2025, with ~$3B more savings planned for 2026 (4Q25 release). +0.5pt margin on top of guidance.",
    "S2: + Healthcare & SMB growth": "Higher-yield healthcare logistics (Andlauer acquisition in 2025) and small-business volume replacing Amazon volume. +0.5pt revenue growth a year.",
    "S3: + Capex discipline": "Automation lets UPS hold reinvestment at ~4.25% of revenue. -0.25pt capex in years 2-5.",
    "S1: Network 2.0 & one FedEx savings": "FedEx targets $1.0B of structural cost reduction by end of calendar 2026 (FY26 10-K outlook). ~$1B on ~$98B revenue = +1.0pt margin.",
    "S2: + Capex held at ~4% of revenue": "FY26 capex was 4% of revenue, the lowest in eleven years (FY26 shareholder letter). Hold it there: -0.6pt capex in years 2-5.",
    "S3: + International & B2B growth": "Faster growth in higher-margin B2B and international export lanes. +0.5pt revenue growth a year.",
}
for t in ("UPS", "FDX"):
    for name, d in STRATEGIES[t]:
        put(a, f"A{r}", name)
        k, v = list(d.items())[0]
        put(a, f"B{r}", v, BLUE, fmt=PCT2, align="right", fill=YELLOW)
        put(a, f"C{r}", CO[t]["name"], GREY, align="center")
        put(a, f"E{r}", f"{param_label[k].capitalize()}. {strat_notes[name]}", GREY, size=9)
        STRAT[t].append((name, k, f"Assumptions!$B${r}"))
        r += 1

r += 1
section(a, r, "Sensitivity grid settings (Exhibit 07)", 5)
GRID_STEP = f"Assumptions!$B${r + 1}"
put(a, f"A{r + 1}", "WACC step between grid columns")
put(a, f"B{r + 1}", 0.005, BLUE, fmt=PCT2, align="right")
put(a, f"A{r + 2}", "Terminal growth values to test")
G_CELLS = []
for k, g in enumerate([0.01, 0.015, 0.02, 0.025, 0.03]):
    put(a, f"{CL(2 + k)}{r + 2}", g, BLUE, fmt=PCT, align="right")
    G_CELLS.append(f"Assumptions!${CL(2 + k)}${r + 2}")
a.column_dimensions["F"].width = 10
a.freeze_panes = "B5"


def A(t, param):
    return f"Assumptions!${CO[t]['col']}${AROW[param]}"


# --------------------------------------------------------------- Ex06 WACC
w = S["Ex06_WACC"]
title(w, "Exhibit 06 - Weighted average cost of capital (WACC)", "CAPM cost of equity, cost of debt from the 10-K, market-value weights at each valuation date")
widths(w, {"A": 42, "B": 15, "C": 15, "D": 70})
header(w, 4, 1, ["Component", "UPS", "FedEx", "How it's calculated"], 24)
WR = {}
wrows = [
    ("Valuation date", lambda t: f"={A(t, 'Valuation date')}", None, "Latest fiscal year end"),
    ("Risk-free rate (10-yr Treasury)", lambda t: f"={A(t, 'Risk-free rate')}", PCT2, "Assumptions"),
    ("Equity beta", lambda t: f"={A(t, 'Equity beta')}", IDX, "Assumptions (5-year monthly beta vs S&P 500)"),
    ("Equity risk premium", lambda t: f"={A(t, 'Equity risk premium')}", PCT2, "Assumptions"),
    ("Cost of equity (CAPM)", None, PCT2, "Risk-free + beta x equity risk premium"),
    ("Interest expense", lambda t: f"={F(t, 'Interest expense')}", USD, "Latest fiscal year (10-K)"),
    ("Total debt - prior year", lambda t: f"={F(t, 'Current debt', 1)}+{F(t, 'Long-term debt', 1)}", USD, "Current + long-term debt incl. finance leases (operating leases excluded)"),
    ("Total debt - latest year", lambda t: f"={F(t, 'Current debt')}+{F(t, 'Long-term debt')}", USD, "Current + long-term debt incl. finance leases"),
    ("Pre-tax cost of debt", None, PCT2, "Interest expense / average debt"),
    ("Tax rate", lambda t: f"={A(t, 'Tax rate (forecast)')}", PCT2, "Assumptions"),
    ("After-tax cost of debt", None, PCT2, "Pre-tax cost of debt x (1 - tax rate)"),
    ("Share price", lambda t: f"={A(t, 'Share price at valuation date')}", PS, "Assumptions"),
    ("Diluted shares (millions)", lambda t: f"={A(t, 'Diluted shares (millions)')}", NUM1, "Assumptions"),
    ("Market value of equity", None, USD, "Price x diluted shares"),
    ("Debt", None, USD, "Latest total debt (book value as a proxy for market value)"),
    ("Weight of equity", None, PCT, "E / (D + E)"),
    ("Weight of debt", None, PCT, "D / (D + E)"),
    ("WACC", None, PCT2, "We x Ke + Wd x Kd after tax"),
]
for i, (lab, fn, fmt, how) in enumerate(wrows):
    WR[lab] = 5 + i
for i, (lab, fn, fmt, how) in enumerate(wrows):
    r = 5 + i
    bold = lab in ("WACC", "Cost of equity (CAPM)")
    put(w, f"A{r}", lab, bold=bold)
    put(w, f"D{r}", how, GREY, size=9)
    for col, t in (("B", "UPS"), ("C", "FDX")):
        R = lambda k: f"{col}{WR[k]}"  # noqa: E731
        if fn:
            f = fn(t)
            put(w, f"{col}{r}", f, GREEN, fmt=fmt, align="right")
        else:
            f = {"Cost of equity (CAPM)": f"={R('Risk-free rate (10-yr Treasury)')}+{R('Equity beta')}*{R('Equity risk premium')}",
                 "Pre-tax cost of debt": f"=IF(({R('Total debt - prior year')}+{R('Total debt - latest year')})=0,0,{R('Interest expense')}/AVERAGE({R('Total debt - prior year')},{R('Total debt - latest year')}))",
                 "After-tax cost of debt": f"={R('Pre-tax cost of debt')}*(1-{R('Tax rate')})",
                 "Market value of equity": f"={R('Share price')}*{R('Diluted shares (millions)')}",
                 "Debt": f"={R('Total debt - latest year')}",
                 "Weight of equity": f"={R('Market value of equity')}/({R('Market value of equity')}+{R('Debt')})",
                 "Weight of debt": f"=1-{R('Weight of equity')}",
                 "WACC": f"={R('Weight of equity')}*{R('Cost of equity (CAPM)')}+{R('Weight of debt')}*{R('After-tax cost of debt')}"}[lab]
            c = put(w, f"{col}{r}", f, fmt=fmt, bold=bold, align="right")
            if lab == "WACC":
                c.border = DTOP
                c.fill = LIGHT
WACC_CELL = {"UPS": f"Ex06_WACC!$B${WR['WACC']}", "FDX": f"Ex06_WACC!$C${WR['WACC']}"}

# ------------------------------------------------------------- DCF Engine
e = S["DCF_Engine"]
title(e, "DCF engine - every scenario is one row (base case, 7 value-driver shocks, 3 strategy layers per company)",
      "Revenue_t = Revenue_t-1 x (1+g_t) | FCF_t = Revenue_t x (margin x (1-tax) + D&A% - capex%) - NWC% x change in revenue | TV = FCF_5 x (1+g) / (WACC-g)")
ecols = ["Company", "Scenario", "Base revenue", "g Y1", "g Y2", "g Y3", "g Y4", "g Y5", "Op margin", "Tax rate", "D&A %",
         "Capex % Y1", "Capex % Y2-5", "NWC % of chg rev", "WACC", "Terminal g", "Cash", "Debt", "Pension (after tax)",
         "Diluted shares", "Share price", "Revenue Y1", "Revenue Y2", "Revenue Y3", "Revenue Y4", "Revenue Y5",
         "FCF Y1", "FCF Y2", "FCF Y3", "FCF Y4", "FCF Y5", "PV of FCF", "Terminal value", "PV of terminal value",
         "Enterprise value", "Equity value", "Value per share", "vs share price", "Change vs base"]
header(e, 4, 1, ecols, 42)
for i in range(len(ecols)):
    e.column_dimensions[CL(i + 1)].width = 11
e.column_dimensions["B"].width = 38
PCOL = {"base_revenue": "C", "growth": ["D", "E", "F", "G", "H"], "op_margin": "I", "tax_rate": "J", "da_pct": "K",
        "capex_y1": "L", "capex_y2_5": "M", "nwc_pct": "N", "wacc": "O", "terminal_g": "P", "cash": "Q", "debt": "R",
        "pension_after_tax": "S", "shares": "T", "price": "U"}
PFMT = {"C": USD, "I": PCT2, "J": PCT2, "K": PCT2, "L": PCT2, "M": PCT2, "N": PCT2, "O": PCT2, "P": PCT2,
        "Q": USD, "R": USD, "S": USD, "T": NUM1, "U": PS}
for c in "DEFGH":
    PFMT[c] = PCT2
EROW = {"UPS": {}, "FDX": {}}


def engine_calc(r):
    fm = lambda c, f, fmt=USD, bold=False: put(e, f"{c}{r}", f, fmt=fmt, bold=bold, align="right")  # noqa: E731
    revs = ["V", "W", "X", "Y", "Z"]
    fcfs = ["AA", "AB", "AC", "AD", "AE"]
    gcols = PCOL["growth"]
    for k in range(5):
        prev = "C" if k == 0 else revs[k - 1]
        fm(revs[k], f"={prev}{r}*(1+{gcols[k]}{r})")
        capex = "L" if k == 0 else "M"
        fm(fcfs[k], f"={revs[k]}{r}*(I{r}*(1-J{r})+K{r}-{capex}{r})-N{r}*({revs[k]}{r}-{prev}{r})")
    fm("AF", "=" + "+".join(f"{fcfs[k]}{r}/(1+O{r})^{k + 1}" for k in range(5)))
    fm("AG", f"=AE{r}*(1+P{r})/(O{r}-P{r})")
    fm("AH", f"=AG{r}/(1+O{r})^5")
    fm("AI", f"=AF{r}+AH{r}")
    fm("AJ", f"=AI{r}+Q{r}-R{r}-S{r}")
    fm("AK", f"=AJ{r}/T{r}", PS, True)
    fm("AL", f"=AK{r}/U{r}-1", PCT)


r = 5
for t in ("UPS", "FDX"):
    base = r
    EROW[t]["Base case"] = r
    put(e, f"A{r}", CO[t]["name"], bold=True)
    put(e, f"B{r}", "Base case", bold=True)
    links = {
        "C": f"={F(t, 'Revenue')}",
        **{g: f"={A(t, f'Revenue growth Y{k + 1}')}" for k, g in enumerate(PCOL["growth"])},
        "I": f"={A(t, 'Operating margin')}", "J": f"={A(t, 'Tax rate (forecast)')}", "K": f"={A(t, 'D&A % of revenue')}",
        "L": f"={A(t, 'Capex % of revenue Y1')}", "M": f"={A(t, 'Capex % of revenue Y2-Y5')}",
        "N": f"={A(t, 'Net working capital % of revenue change')}", "O": f"={WACC_CELL[t]}", "P": f"={A(t, 'Terminal growth rate')}",
        "Q": f"={F(t, 'Cash and cash equivalents')}+{F(t, 'Marketable securities')}",
        "R": f"={F(t, 'Current debt')}+{F(t, 'Long-term debt')}",
        "S": f"={F(t, 'Pension and postretirement obligations')}*(1-J{r})",
        "T": f"={A(t, 'Diluted shares (millions)')}", "U": f"={A(t, 'Share price at valuation date')}",
    }
    for c, f in links.items():
        put(e, f"{c}{r}", f, GREEN, fmt=PFMT[c], align="right")
    engine_calc(r)
    put(e, f"AM{r}", 0, fmt=PS, align="right")
    for c in range(1, len(ecols) + 1):
        e.cell(row=r, column=c).fill = LIGHT
    r += 1

    def shocked_row(r, label, src_row, deltas):
        """deltas: {param: cell_ref}; other params link to src_row."""
        put(e, f"A{r}", CO[t]["name"])
        put(e, f"B{r}", label)
        for p, col in PCOL.items():
            cols = col if isinstance(col, list) else [col]
            for c in cols:
                f = f"={c}{src_row}"
                if p in deltas:
                    f += f"+{deltas[p]}"
                put(e, f"{c}{r}", f, fmt=PFMT[c], align="right", color=BLUE if False else BLACK)
        engine_calc(r)
        put(e, f"AM{r}", f"=AK{r}-AK{base}", fmt=PS, align="right")

    for name, d in DRIVER_SHOCKS:
        shocked_row(r, name, base, SHOCK[name][1])
        EROW[t][name] = r
        r += 1
    prev = base
    for name, k, ref in STRAT[t]:
        shocked_row(r, name, prev, {k: ref})
        EROW[t][name] = r
        prev = r
        r += 1
    for c in range(1, len(ecols) + 1):
        e.cell(row=base, column=c).border = TOP
e.freeze_panes = "C5"
note(e, "S4", "Unfunded pension and postretirement obligations are treated as debt-like, net of tax.")
note(e, "Q4", "All cash and marketable securities at the valuation date are added back.")


def E(t, scen, col):
    return f"DCF_Engine!${col}${EROW[t][scen]}"


# ---------------------------------------------------------- Ex05 Forecast
fc = S["Ex05_Forecast"]
title(fc, "Exhibit 05 - Five-year free cash flow forecast (base case)", "Built from the drivers on the Assumptions sheet; the check row ties every year to the DCF engine")
widths(fc, {"A": 36, "B": 13, "C": 13, "D": 13, "E": 13, "F": 13, "G": 13, "H": 50})
r = 4
for t in ("UPS", "FDX"):
    fy0 = CO[t]["first_fy"] - 1
    section(fc, r, f"{CO[t]['name']} ($ millions)", 8)
    header(fc, r + 1, 1, ["Line item", f"FY{fy0}A"] + [f"FY{fy0 + k}F" for k in range(1, 6)] + ["Driver"])
    R = {k: r + 2 + i for i, k in enumerate(["Revenue", "Revenue growth", "Operating margin", "EBIT", "Tax rate", "NOPAT",
                                             "D&A", "Capex", "Change in NWC", "Free cash flow", "Check vs DCF engine"])}
    put(fc, f"A{R['Revenue']}", "Revenue", bold=True)
    put(fc, f"B{R['Revenue']}", f"={F(t, 'Revenue')}", GREEN, fmt=USD)
    put(fc, f"B{R['EBIT']}", f"={F(t, 'Operating income (EBIT)')}", GREEN, fmt=USD)
    put(fc, f"B{R['Operating margin']}", f"=B{R['EBIT']}/B{R['Revenue']}", fmt=PCT)
    put(fc, f"B{R['D&A']}", f"={F(t, 'Depreciation and amortization')}", GREEN, fmt=USD)
    put(fc, f"B{R['Capex']}", f"={F(t, 'Capital expenditures')}", GREEN, fmt=USD)
    drivers = {"Revenue growth": "Assumptions: revenue growth by year", "Operating margin": "Assumptions: operating margin",
               "Tax rate": "Assumptions: forecast tax rate", "D&A": "Revenue x D&A %", "Capex": "Revenue x capex % (Y1 differs)",
               "Change in NWC": "NWC % x change in revenue", "Free cash flow": "NOPAT + D&A - capex - change in NWC",
               "Check vs DCF engine": "Each year must equal the DCF_Engine base-case FCF"}
    for lab, row in R.items():
        if lab != "Revenue":
            put(fc, f"A{row}", lab, bold=lab in ("Free cash flow", "EBIT", "NOPAT"))
        if lab in drivers:
            put(fc, f"H{row}", drivers[lab], GREY, size=9)
    for k in range(5):
        c, p = CL(3 + k), CL(2 + k)
        put(fc, f"{c}{R['Revenue growth']}", f"={A(t, f'Revenue growth Y{k + 1}')}", GREEN, fmt=PCT)
        put(fc, f"{c}{R['Revenue']}", f"={p}{R['Revenue']}*(1+{c}{R['Revenue growth']})", fmt=USD, bold=True)
        put(fc, f"{c}{R['Operating margin']}", f"={A(t, 'Operating margin')}", GREEN, fmt=PCT)
        put(fc, f"{c}{R['EBIT']}", f"={c}{R['Revenue']}*{c}{R['Operating margin']}", fmt=USD, bold=True)
        put(fc, f"{c}{R['Tax rate']}", f"={A(t, 'Tax rate (forecast)')}", GREEN, fmt=PCT)
        put(fc, f"{c}{R['NOPAT']}", f"={c}{R['EBIT']}*(1-{c}{R['Tax rate']})", fmt=USD, bold=True)
        put(fc, f"{c}{R['D&A']}", f"={c}{R['Revenue']}*{A(t, 'D&A % of revenue')}", fmt=USD)
        cap = "Capex % of revenue Y1" if k == 0 else "Capex % of revenue Y2-Y5"
        put(fc, f"{c}{R['Capex']}", f"={c}{R['Revenue']}*{A(t, cap)}", fmt=USD)
        put(fc, f"{c}{R['Change in NWC']}", f"=({c}{R['Revenue']}-{p}{R['Revenue']})*{A(t, 'Net working capital % of revenue change')}", fmt=USD)
        cf = put(fc, f"{c}{R['Free cash flow']}", f"={c}{R['NOPAT']}+{c}{R['D&A']}-{c}{R['Capex']}-{c}{R['Change in NWC']}", fmt=USD, bold=True)
        cf.border = DTOP
        put(fc, f"{c}{R['Check vs DCF engine']}", f'=IF(ABS({c}{R["Free cash flow"]}-{E(t, "Base case", ["AA", "AB", "AC", "AD", "AE"][k])})<0.5,"OK","CHECK")', GREY, bold=True, align="center")
    r = R["Check vs DCF engine"] + 2

# --------------------------------------------------------- Ex07 Valuation
v = S["Ex07_Valuation"]
title(v, "Exhibit 07 - DCF valuation, sensitivity and market-implied expectations", "Base case | Values in $ millions except per share")
widths(v, {"A": 36, "B": 14, "C": 14, "D": 14, "E": 4, "F": 36, "G": 14, "H": 14, "I": 14})
VR = {}
for t, c0 in (("UPS", 1), ("FDX", 6)):
    L = lambda k: CL(c0 + k)  # noqa: E731
    section(v, 4, f"{CO[t]['name']} - valuation at {('Dec 31, 2025' if t == 'UPS' else 'May 31, 2026')}", 4, c0)
    rows = [("WACC", E(t, "Base case", "O"), PCT2), ("Terminal growth rate", E(t, "Base case", "P"), PCT2)]
    for i, (lab, ref, fmt) in enumerate(rows):
        put(v, f"{L(0)}{5 + i}", lab)
        put(v, f"{L(1)}{5 + i}", f"={ref}", GREEN, fmt=fmt, align="right")
    header(v, 8, c0, ["Free cash flow", "Cash flow", "PV", "Cumulative PV"])
    fy1 = CO[t]["first_fy"]
    for k in range(5):
        rr = 9 + k
        put(v, f"{L(0)}{rr}", f"FY{fy1 + k}F", align="center")
        put(v, f"{L(1)}{rr}", f"={E(t, 'Base case', ['AA', 'AB', 'AC', 'AD', 'AE'][k])}", GREEN, fmt=USD)
        put(v, f"{L(2)}{rr}", f"={L(1)}{rr}/(1+{L(1)}$5)^{k + 1}", fmt=USD)
        put(v, f"{L(3)}{rr}", f"=SUM({L(2)}$9:{L(2)}{rr})", fmt=USD)
    put(v, f"{L(0)}14", "Terminal value", bold=True)
    put(v, f"{L(1)}14", f"={E(t, 'Base case', 'AG')}", GREEN, fmt=USD)
    put(v, f"{L(2)}14", f"={L(1)}14/(1+{L(1)}$5)^5", fmt=USD)
    put(v, f"{L(3)}14", f"={L(3)}13+{L(2)}14", fmt=USD, bold=True)
    items = [
        (16, "Enterprise value", f"={L(3)}14", USD, True),
        (17, "+ Cash and marketable securities", f"={E(t, 'Base case', 'Q')}", USD, False),
        (18, "- Debt", f"={E(t, 'Base case', 'R')}", USD, False),
        (19, "- Pension obligations (after tax)", f"={E(t, 'Base case', 'S')}", USD, False),
        (20, "Estimated value of equity", f"={L(1)}16+{L(1)}17-{L(1)}18-{L(1)}19", USD, True),
        (21, "Diluted shares (millions)", f"={E(t, 'Base case', 'T')}", NUM1, False),
        (22, "Estimated value per share", f"={L(1)}20/{L(1)}21", PS, True),
        (23, "Actual share price", f"={E(t, 'Base case', 'U')}", PS, False),
        (24, "Difference per share", f"={L(1)}22-{L(1)}23", PS, False),
        (25, "Upside / (downside)", f"={L(1)}22/{L(1)}23-1", PCT, True),
        (26, "Terminal value as % of EV", f"={L(2)}14/{L(1)}16", PCT, False),
        (27, "Check: equals DCF engine", f'=IF(ABS({L(1)}22-{E(t, "Base case", "AK")})<0.01,"OK","CHECK")', None, False),
    ]
    for rr, lab, f, fmt, bold in items:
        put(v, f"{L(0)}{rr}", lab, bold=bold)
        color = GREEN if "DCF_Engine" in f and "IF(" not in f else BLACK
        c = put(v, f"{L(1)}{rr}", f, GREY if lab.startswith("Check") else color, fmt=fmt, bold=bold, align="right")
        if lab in ("Estimated value per share",):
            c.fill = LIGHT
            c.border = DTOP
    VR[t] = {"vps": f"Ex07_Valuation!${L(1)}$22", "upside": f"Ex07_Valuation!${L(1)}$25", "ev": f"Ex07_Valuation!${L(1)}$16",
             "tvpct": f"Ex07_Valuation!${L(1)}$26"}

    # market multiples cross-check
    section(v, 29, "Multiples cross-check (latest fiscal year)", 4, c0)
    mult = [
        (30, "EBITDA (EBIT + D&A)", f"={F(t, 'Operating income (EBIT)')}+{F(t, 'Depreciation and amortization')}", USD),
        (31, "Market enterprise value", f"={E(t, 'Base case', 'U')}*{E(t, 'Base case', 'T')}+{E(t, 'Base case', 'R')}+{E(t, 'Base case', 'S')}-{E(t, 'Base case', 'Q')}", USD),
        (32, "Market EV / EBITDA", f"={L(1)}31/{L(1)}30", MULT),
        (33, "DCF EV / EBITDA", f"={L(1)}16/{L(1)}30", MULT),
        (34, "Price / earnings (diluted EPS)", f"={E(t, 'Base case', 'U')}/{F(t, 'Diluted EPS ($)')}", MULT),
        (35, "Dividend yield", f"={F(t, 'Dividends per share ($)')}/{E(t, 'Base case', 'U')}", PCT),
    ]
    for rr, lab, f, fmt in mult:
        put(v, f"{L(0)}{rr}", lab)
        put(v, f"{L(1)}{rr}", f, GREEN if rr in (30, 31) else BLACK, fmt=fmt, align="right")
    VR[t].update({"mkt_ev_ebitda": f"Ex07_Valuation!${L(1)}$33".replace("33", "32"), "pe": f"Ex07_Valuation!${L(1)}$34",
                  "yield": f"Ex07_Valuation!${L(1)}$35", "dcf_ev_ebitda": f"Ex07_Valuation!${L(1)}$33"})

    # reverse DCF
    section(v, 37, "Market-implied expectations (reverse DCF)", 4, c0)
    put(v, f"{L(0)}38", "Base-case operating margin")
    put(v, f"{L(1)}38", f"={E(t, 'Base case', 'I')}", GREEN, fmt=PCT2, align="right")
    put(v, f"{L(0)}39", "Value per share at margin +1pt")
    put(v, f"{L(1)}39", f"={E(t, 'Operating margin +1pt', 'AK')}", GREEN, fmt=PS, align="right")
    put(v, f"{L(0)}40", "Operating margin implied by share price", bold=True)
    c = put(v, f"{L(1)}40", f"={L(1)}38+({L(1)}23-{L(1)}22)/({L(1)}39-{L(1)}22)*0.01", fmt=PCT2, bold=True, align="right")
    c.fill = LIGHT
    put(v, f"{L(0)}41", "Gap vs base-case margin (pts)")
    put(v, f"{L(1)}41", f"={L(1)}40-{L(1)}38", fmt=PCT2, align="right")
    note(v, f"{L(1)}40", "Value per share is linear in operating margin, so the margin that makes value = price is found exactly by interpolation.")
    VR[t].update({"implied_margin": f"Ex07_Valuation!${L(1)}$40", "base_margin": f"Ex07_Valuation!${L(1)}$38"})

# sensitivity grids, to the right of the valuation blocks (columns K-P)
put(v, "R3", "t (helper)", GREY, size=8, align="center")
for k in range(5):
    put(v, f"{CL(19 + k)}3", k + 1, GREY, size=8, align="center")
trange = "$S$3:$W$3"
for idx, t in enumerate(("UPS", "FDX")):
    top = 4 + idx * 9
    section(v, top, f"{CO[t]['name']} - value per share: WACC (rows) x terminal growth (columns)", 6, 11)
    fcf_range = f"DCF_Engine!$AA${EROW[t]['Base case']}:$AE${EROW[t]['Base case']}"
    fcf5 = E(t, "Base case", "AE")
    q, d_, s_, sh = (E(t, "Base case", x) for x in ("Q", "R", "S", "T"))
    grid_cols = ["L", "M", "N", "O", "P"]
    put(v, f"K{top + 1}", "WACC \\ growth", bold=True, align="center", fill=SEC)
    for k, gc in enumerate(grid_cols):
        put(v, f"{gc}{top + 1}", f"={G_CELLS[k]}", GREEN, fmt=PCT, bold=True, align="center", fill=SEC)
    for j, step in enumerate((-2, -1, 0, 1, 2)):
        rr = top + 2 + j
        put(v, f"K{rr}", f"={E(t, 'Base case', 'O')}+({step})*{GRID_STEP}", GREEN, fmt=PCT2, bold=True, align="center", fill=SEC)
        for gc in grid_cols:
            wcell, gcell = f"$K{rr}", f"{gc}${top + 1}"
            f = (f"=(SUMPRODUCT({fcf_range}/(1+{wcell})^{trange})+{fcf5}*(1+{gcell})/({wcell}-{gcell})/(1+{wcell})^5"
                 f"+{q}-{d_}-{s_})/{sh}")
            c = put(v, f"{gc}{rr}", f, fmt=PS, align="right")
            if step == 0 and gc == "N":
                c.fill = YELLOW
    VR[t]["grid_center"] = f"Ex07_Valuation!$N${top + 4}"
    note(v, f"N{top + 4}", "Yellow cell = base case. It equals the estimated value per share in the DCF block.")
put(v, "K22", "Each cell re-runs the DCF with that WACC and terminal growth, using the base-case cash flows.", GREY, italic=True, size=9)
v.column_dimensions["J"].width = 3
v.column_dimensions["K"].width = 16
for col in "LMNOP":
    v.column_dimensions[col].width = 11
for col in "RSTUVW":
    v.column_dimensions[col].width = 4

# ------------------------------------------------------ Ex08 Value drivers
vd = S["Ex08_Value_Drivers"]
title(vd, "Exhibit 08 - Key value drivers: impact of a one-point change on value per share", "Each driver is moved by one percentage point on its own; every other assumption stays at the base case")
widths(vd, {"A": 34, "B": 12, "C": 14, "D": 14, "E": 14, "F": 4, "G": 34, "H": 12, "I": 14, "J": 14, "K": 14})
for t, c0 in (("UPS", 1), ("FDX", 7)):
    L = lambda k: CL(c0 + k)  # noqa: E731
    section(vd, 4, CO[t]["name"], 5, c0)
    header(vd, 5, c0, ["Value driver", "Change", "Base value", "New value", "Impact per share"], 30)
    for i, (name, d) in enumerate(DRIVER_SHOCKS):
        rr = 6 + i
        put(vd, f"{L(0)}{rr}", name)
        put(vd, f"{L(1)}{rr}", f"=Assumptions!$B${SHOCK[name][0]}", GREEN, fmt=PCT2, align="right")
        put(vd, f"{L(2)}{rr}", f"={E(t, 'Base case', 'AK')}", GREEN, fmt=PS)
        put(vd, f"{L(3)}{rr}", f"={E(t, name, 'AK')}", GREEN, fmt=PS)
        put(vd, f"{L(4)}{rr}", f"={L(3)}{rr}-{L(2)}{rr}", fmt=PS, bold=True)
    put(vd, f"{L(0)}14", "Impact = value per share after the change minus base-case value per share.", GREY, italic=True, size=9)
    ch = BarChart()
    ch.type = "bar"
    ch.add_data(Reference(vd, min_col=c0 + 4, min_row=5, max_row=12), titles_from_data=True)
    ch.set_categories(Reference(vd, min_col=c0, min_row=6, max_row=12))
    ch.series[0].graphicalProperties.solidFill = CO[t]["color"]
    ch.series[0].graphicalProperties.line.solidFill = CO[t]["color"]
    tidy_chart(ch, f"{CO[t]['name']}: $ per share from a 1pt change", '$0')
    ch.legend = None
    ch.x_axis.scaling.orientation = "maxMin"
    vd.add_chart(ch, f"{L(0)}16")

# ---------------------------------------------------------- Ex09 Strategy
st = S["Ex09_Strategy"]
title(st, "Exhibit 09 - Strategy for increasing shareholder value (cumulative scenarios)", "Each layer keeps the changes before it, like Exhibits 9-11 of a classic value-driver analysis")
widths(st, {"A": 44, "B": 14, "C": 14, "D": 14, "E": 4, "F": 44, "G": 14, "H": 14, "I": 14})
for t, c0 in (("UPS", 1), ("FDX", 6)):
    L = lambda k: CL(c0 + k)  # noqa: E731
    section(st, 4, CO[t]["name"], 4, c0)
    header(st, 5, c0, ["Scenario", "Value per share", "Added by this step", "vs share price"], 30)
    scen = ["Base case"] + [s[0] for s in STRAT[t]]
    for i, s in enumerate(scen):
        rr = 6 + i
        put(st, f"{L(0)}{rr}", s, bold=(i == 0))
        put(st, f"{L(1)}{rr}", f"={E(t, s, 'AK')}", GREEN, fmt=PS, bold=True)
        put(st, f"{L(2)}{rr}", "-" if i == 0 else f"={L(1)}{rr}-{L(1)}{rr - 1}", fmt=PS, align="right")
        put(st, f"{L(3)}{rr}", f"={L(1)}{rr}/{E(t, 'Base case', 'U')}-1", fmt=PCT)
    put(st, f"{L(0)}10", "Actual share price", bold=True)
    put(st, f"{L(1)}10", f"={E(t, 'Base case', 'U')}", GREEN, fmt=PS, bold=True)
    ch = BarChart()
    ch.type = "col"
    ch.add_data(Reference(st, min_col=c0 + 1, min_row=5, max_row=10), titles_from_data=True)
    ch.set_categories(Reference(st, min_col=c0, min_row=6, max_row=10))
    ch.series[0].graphicalProperties.solidFill = CO[t]["color"]
    ch.series[0].graphicalProperties.line.solidFill = CO[t]["color"]
    tidy_chart(ch, f"{CO[t]['name']}: value per share by scenario vs share price", '$0')
    ch.legend = None
    st.add_chart(ch, f"{L(0)}12")

# ------------------------------------------------------------ Ex01 Returns
rt = S["Ex01_Returns"]
title(rt, "Exhibit 01 - Total shareholder returns vs the market", "Cumulative value of $100 invested with dividends reinvested, from each company's 10-K performance graph | Blue = reported index values")
widths(rt, {"A": 34, **{CL(c): 12 for c in range(2, 9)}})
r = 4
RET = {}
for t in ("UPS", "FDX"):
    sub = tsr[tsr.company == t]
    dates = list(dict.fromkeys(sub.period_end))
    series = list(dict.fromkeys(sub.series))
    section(rt, r, f"{CO[t]['name']} ({'calendar years' if t == 'UPS' else 'fiscal years ending May 31'})", 8)
    header(rt, r + 1, 1, ["Index value ($100 invested)"] + [d[:7] for d in dates])
    for i, s in enumerate(series):
        rr = r + 2 + i
        put(rt, f"A{rr}", s, bold=(i == 0))
        for j, d in enumerate(dates):
            val = float(sub[(sub.series == s) & (sub.period_end == d)].index_value.iloc[0])
            put(rt, f"{CL(2 + j)}{rr}", val, BLUE, fmt=IDX)
    r2 = r + 2 + len(series) + 1
    header(rt, r2, 1, ["Annual total return"] + [d[:7] for d in dates[1:]])
    for i, s in enumerate(series):
        rr = r2 + 1 + i
        put(rt, f"A{rr}", s, bold=(i == 0))
        for j in range(1, len(dates)):
            src = r + 2 + i
            put(rt, f"{CL(1 + j)}{rr}", f"={CL(2 + j)}{src}/{CL(1 + j)}{src}-1", fmt=PCT)
    r3 = r2 + len(series) + 2
    header(rt, r3, 1, ["$1,000 invested", "Last 3 years", "Last 5 years", "5-yr CAGR"])
    for i, s in enumerate(series):
        rr = r3 + 1 + i
        src = r + 2 + i
        last, three, five = CL(1 + len(dates)), CL(1 + len(dates) - 3), "B"
        put(rt, f"A{rr}", s, bold=(i == 0))
        put(rt, f"B{rr}", f"=1000*{last}{src}/{three}{src}", fmt=USD)
        put(rt, f"C{rr}", f"=1000*{last}{src}/{five}{src}", fmt=USD)
        put(rt, f"D{rr}", f"=({last}{src}/{five}{src})^(1/5)-1", fmt=PCT)
        if i == 0:
            RET[t] = {"3y": f"Ex01_Returns!$B${rr}", "5y_cagr": f"Ex01_Returns!$D${rr}", "spx_5y_cagr": f"Ex01_Returns!$D${rr + 1}"}
    r = r3 + len(series) + 3
put(rt, f"A{r}", "Return = Index(t) / Index(t-1) - 1. Periods differ because FedEx's fiscal year ended May 31 until 2026; the Python notebook compares both on identical dates.", GREY, italic=True, size=9)

# ------------------------------------------------------------- Ex02 Ratios
ra = S["Ex02_Ratios"]
title(ra, "Exhibit 02 - Ratio analysis: UPS compared to FedEx", "Year-end balances | Time-series = latest vs earliest year | Cross-section = which company is stronger on the latest year")
widths(ra, {"A": 36, **{CL(c): 11 for c in range(2, 8)}, "H": 10, "I": 10, "J": 12})
yu, yf = list(YCOL["UPS"]), list(YCOL["FDX"])
header(ra, 4, 1, ["Ratio"] + [f"UPS {y}" for y in yu] + [f"FedEx {y}" for y in yf] + ["UPS trend", "FedEx trend", "Stronger"], 32)
ratio_defs = [
    ("Liquidity", None, None, None),
    ("Current ratio", "{TCA}/{TCL}", MULT, True),
    ("Quick ratio", "({CASH}+{MS}+{AR})/{TCL}", MULT, True),
    ("Activity", None, None, None),
    ("Receivables collection period (days)", "{AR}/{REV}*365", DAYS, False),
    ("Fixed asset turnover", "{REV}/{PPE}", MULT, True),
    ("Total asset turnover", "{REV}/{TA}", MULT, True),
    ("Debt", None, None, None),
    ("Debt / equity", "({CD}+{LTD})/{EQ}", PCT, False),
    ("Debt / EBITDA", "({CD}+{LTD})/({EBIT}+{DA})", MULT, False),
    ("Interest coverage (EBIT / interest)", "{EBIT}/{INT}", MULT, True),
    ("Profitability", None, None, None),
    ("Operating margin", "{EBIT}/{REV}", PCT, True),
    ("Net profit margin", "{NI}/{REV}", PCT, True),
    ("Return on assets", "{NI}/{TA}", PCT, True),
    ("Return on equity", "{NI}/{EQ}", PCT, True),
    ("Return on invested capital", "{EBIT}*(1-{TAX}/{EBT})/({CD}+{LTD}+{EQ}-{CASH}-{MS})", PCT, True),
]
keys = {"TCA": "Total current assets", "TCL": "Total current liabilities", "CASH": "Cash and cash equivalents", "MS": "Marketable securities",
        "AR": "Accounts receivable net", "REV": "Revenue", "PPE": "Property plant and equipment net", "TA": "Total assets",
        "CD": "Current debt", "LTD": "Long-term debt", "EQ": "Total shareholders equity", "EBIT": "Operating income (EBIT)",
        "DA": "Depreciation and amortization", "INT": "Interest expense", "NI": "Net income", "TAX": "Income tax expense",
        "EBT": "Income before income taxes"}
RATIO_ROW = {}
for i, (lab, f, fmt, hib) in enumerate(ratio_defs):
    rr = 5 + i
    if f is None:
        put(ra, f"A{rr}", lab, NAVY, bold=True)
        continue
    RATIO_ROW[lab] = rr
    put(ra, f"A{rr}", lab)
    col = 2
    for t in ("UPS", "FDX"):
        for yi in range(3):
            refs = {k: F(t, item, yi) for k, item in keys.items()}
            put(ra, f"{CL(col)}{rr}", "=" + f.format(**refs), GREEN, fmt=fmt, align="right")
            col += 1
    for tcol, (c_first, c_last) in (("H", ("B", "D")), ("I", ("E", "G"))):
        put(ra, f"{tcol}{rr}", f'=IF({c_last}{rr}>{c_first}{rr}*1.02,"Up",IF({c_last}{rr}<{c_first}{rr}*0.98,"Down","Steady"))', align="center")
    cmp = ">" if hib else "<"
    put(ra, f"J{rr}", f'=IF(D{rr}{cmp}G{rr},"UPS","FedEx")', bold=True, align="center")
put(ra, "A24", "Stronger = higher value, except collection period, debt/equity and debt/EBITDA where lower is better. UPS 'Marketable securities' is zero in FY2025 because the 10-K folds it into other current assets.", GREY, italic=True, size=9)
ra.freeze_panes = "B5"

# -------------------------------------------------------------- Ex03 DuPont
dp = S["Ex03_DuPont"]
title(dp, "Exhibit 03 - Return on equity disaggregation (5-factor DuPont)", "ROE = Tax burden x Interest burden x Operating margin x Asset turnover x Financial leverage")
widths(dp, {"A": 16, "B": 11, "C": 3, "D": 12, "E": 3, "F": 12, "G": 3, "H": 12, "I": 3, "J": 12, "K": 3, "L": 12, "M": 12})
header(dp, 4, 1, ["", "ROE =", "", "Tax burden", "", "Interest burden", "", "Operating margin", "", "Asset turnover", "", "Financial leverage", "Check"], 30)
header(dp, 5, 1, ["", "NI / Equity", "", "NI / EBT", "", "EBT / EBIT", "", "EBIT / Sales", "", "Sales / Assets", "", "Assets / Equity", "Product"])
r = 6
for t in ("UPS", "FDX"):
    for yi, y in enumerate(YCOL[t]):
        g = lambda item: F(t, item, yi)  # noqa: E731
        put(dp, f"A{r}", f"{CO[t]['name']} {y}", bold=True)
        put(dp, f"B{r}", f"={g('Net income')}/{g('Total shareholders equity')}", GREEN, fmt=PCT, bold=True)
        put(dp, f"C{r}", "=", align="center")
        put(dp, f"D{r}", f"={g('Net income')}/{g('Income before income taxes')}", GREEN, fmt=PCT)
        put(dp, f"E{r}", "x", align="center")
        put(dp, f"F{r}", f"={g('Income before income taxes')}/{g('Operating income (EBIT)')}", GREEN, fmt=PCT)
        put(dp, f"G{r}", "x", align="center")
        put(dp, f"H{r}", f"={g('Operating income (EBIT)')}/{g('Revenue')}", GREEN, fmt=PCT)
        put(dp, f"I{r}", "x", align="center")
        put(dp, f"J{r}", f"={g('Revenue')}/{g('Total assets')}", GREEN, fmt=PCT)
        put(dp, f"K{r}", "x", align="center")
        put(dp, f"L{r}", f"={g('Total assets')}/{g('Total shareholders equity')}", GREEN, fmt=PCT)
        put(dp, f"M{r}", f"=D{r}*F{r}*H{r}*J{r}*L{r}", fmt=PCT, italic=True)
        r += 1
    r += 1
put(dp, f"A{r}", "Interest burden above 100% (FedEx) means non-operating income, mainly pension income and interest income, exceeded interest expense.", GREY, italic=True, size=9)

# ---------------------------------------------------------------- Ex04 FCF
fx = S["Ex04_FCF"]
title(fx, "Exhibit 04 - Historical free cash flow", "NOPAT-based free cash flow plus two cash-flow-statement cross-checks | $ millions")
widths(fx, {"A": 40, "B": 12, "C": 12, "D": 12, "E": 3, "F": 40, "G": 12, "H": 12, "I": 12})
for t, c0 in (("UPS", 1), ("FDX", 6)):
    L = lambda k: CL(c0 + k)  # noqa: E731
    yrs = list(YCOL[t])
    header(fx, 4, c0, [CO[t]["name"]] + yrs)
    lines = ["EBIT", "Effective tax rate", "NOPAT", "+ Depreciation & amortization", "- Capital expenditures", "+ Proceeds from disposals",
             "Operating NWC (AR + inventory + other CA - AP - accruals)", "- Increase in operating NWC", "Free cash flow (NOPAT basis)",
             "", "Memo: Unlevered FCF = CFO + after-tax interest - net capex", "Memo: Levered FCF = CFO - capex",
             "Dividends + share repurchases", "Shareholder payouts / levered FCF"]
    R = {lab: 5 + i for i, lab in enumerate(lines)}
    for lab, rr in R.items():
        if lab:
            put(fx, f"{L(0)}{rr}", lab, bold=lab in ("NOPAT", "Free cash flow (NOPAT basis)"))
    for yi in range(3):
        c = L(1 + yi)
        g = lambda item: F(t, item, yi)  # noqa: E731
        put(fx, f"{c}{R['EBIT']}", f"={g('Operating income (EBIT)')}", GREEN, fmt=USD)
        put(fx, f"{c}{R['Effective tax rate']}", f"={g('Income tax expense')}/{g('Income before income taxes')}", GREEN, fmt=PCT)
        put(fx, f"{c}{R['NOPAT']}", f"={c}{R['EBIT']}*(1-{c}{R['Effective tax rate']})", fmt=USD, bold=True)
        put(fx, f"{c}{R['+ Depreciation & amortization']}", f"={g('Depreciation and amortization')}", GREEN, fmt=USD)
        put(fx, f"{c}{R['- Capital expenditures']}", f"={g('Capital expenditures')}", GREEN, fmt=USD)
        put(fx, f"{c}{R['+ Proceeds from disposals']}", f"={g('Proceeds from disposals')}", GREEN, fmt=USD)
        put(fx, f"{c}{R['Operating NWC (AR + inventory + other CA - AP - accruals)']}",
            f"={g('Accounts receivable net')}+{g('Inventories and supplies')}+{g('Other current assets')}-{g('Accounts payable')}-{g('Accrued wages and benefits')}-{g('Other current liabilities')}", GREEN, fmt=USD)
        nwc = R["Operating NWC (AR + inventory + other CA - AP - accruals)"]
        if yi == 0:
            put(fx, f"{c}{R['- Increase in operating NWC']}", "n/a", GREY, align="right")
            put(fx, f"{c}{R['Free cash flow (NOPAT basis)']}", "n/a", GREY, align="right")
        else:
            p = L(yi)
            put(fx, f"{c}{R['- Increase in operating NWC']}", f"={c}{nwc}-{p}{nwc}", fmt=USD)
            cc = put(fx, f"{c}{R['Free cash flow (NOPAT basis)']}",
                     f"={c}{R['NOPAT']}+{c}{R['+ Depreciation & amortization']}-{c}{R['- Capital expenditures']}+{c}{R['+ Proceeds from disposals']}-{c}{R['- Increase in operating NWC']}", fmt=USD, bold=True)
            cc.border = DTOP
        put(fx, f"{c}{R['Memo: Unlevered FCF = CFO + after-tax interest - net capex']}",
            f"={g('Cash from operating activities')}+{g('Interest expense')}*(1-{c}{R['Effective tax rate']})-{g('Capital expenditures')}+{g('Proceeds from disposals')}", GREEN, fmt=USD)
        put(fx, f"{c}{R['Memo: Levered FCF = CFO - capex']}", f"={g('Cash from operating activities')}-{g('Capital expenditures')}", GREEN, fmt=USD)
        put(fx, f"{c}{R['Dividends + share repurchases']}", f"={g('Dividends paid')}+{g('Share repurchases')}", GREEN, fmt=USD)
        put(fx, f"{c}{R['Shareholder payouts / levered FCF']}",
            f"=IF({c}{R['Memo: Levered FCF = CFO - capex']}=0,0,{c}{R['Dividends + share repurchases']}/{c}{R['Memo: Levered FCF = CFO - capex']})", fmt=PCT)
put(fx, "A20", "Payouts above 100% of levered FCF mean dividends and buybacks were partly funded with cash on hand or new debt.", GREY, italic=True, size=9)

# ------------------------------------------------------------------ Summary
sm = S["Summary"]
title(sm, "UPS vs FedEx - valuation summary", "Each company valued at its latest fiscal year end with a 5-year DCF | All values link to the exhibits")
widths(sm, {"A": 40, "B": 16, "C": 16, "D": 60})
header(sm, 4, 1, ["Metric", "UPS", "FedEx", "Where it comes from"], 24)
summary_rows = [
    ("Valuation date", lambda t: f"={A(t, 'Valuation date')}", None, "Assumptions"),
    ("Share price at valuation date", lambda t: f"={E(t, 'Base case', 'U')}", PS, "Assumptions"),
    ("DCF value per share (base case)", lambda t: f"={VR[t]['vps']}", PS, "Exhibit 07"),
    ("Upside / (downside)", lambda t: f"={VR[t]['upside']}", PCT, "Exhibit 07"),
    ("Verdict", None, None, "Undervalued if upside > 15%, overvalued if < -15%"),
    ("WACC", lambda t: f"={WACC_CELL[t]}", PCT2, "Exhibit 06"),
    ("Base-case operating margin", lambda t: f"={VR[t]['base_margin']}", PCT2, "Assumptions"),
    ("Operating margin implied by the share price", lambda t: f"={VR[t]['implied_margin']}", PCT2, "Exhibit 07 reverse DCF"),
    ("Terminal value as % of EV", lambda t: f"={VR[t]['tvpct']}", PCT, "Exhibit 07"),
    ("Market EV / EBITDA", lambda t: f"={VR[t]['mkt_ev_ebitda']}", MULT, "Exhibit 07"),
    ("Price / earnings", lambda t: f"={VR[t]['pe']}", MULT, "Exhibit 07"),
    ("Dividend yield", lambda t: f"={VR[t]['yield']}", PCT, "Exhibit 07"),
    ("Return on invested capital (latest year)", lambda t: f"=Ex02_Ratios!${'D' if t == 'UPS' else 'G'}${RATIO_ROW['Return on invested capital']}", PCT, "Exhibit 02"),
    ("5-year total shareholder return (CAGR)", lambda t: f"={RET[t]['5y_cagr']}", PCT, "Exhibit 01 (10-K performance graph)"),
    ("S&P 500 over the same 5 years (CAGR)", lambda t: f"={RET[t]['spx_5y_cagr']}", PCT, "Exhibit 01"),
    ("Value with all strategy layers", lambda t: f"={E(t, STRAT[t][-1][0], 'AK')}", PS, "Exhibit 09"),
]
SR = {}
for i, (lab, fn, fmt, src) in enumerate(summary_rows):
    rr = 5 + i
    SR[lab] = rr
    put(sm, f"A{rr}", lab, bold=lab in ("DCF value per share (base case)", "Upside / (downside)", "Verdict"))
    put(sm, f"D{rr}", src, GREY, size=9)
    for col, t in (("B", "UPS"), ("C", "FDX")):
        if fn:
            put(sm, f"{col}{rr}", fn(t), GREEN, fmt=fmt, bold=lab.startswith(("DCF", "Upside")), align="right")
        else:
            up = f"{col}{SR['Upside / (downside)']}"
            put(sm, f"{col}{rr}", f'=IF({up}>0.15,"Undervalued",IF({up}<-0.15,"Overvalued","Fairly valued"))', bold=True, align="right")
section(sm, 23, "Model checks", 4)
checks = [
    ("Balance sheets balance (all years)", '=IF(AND(COUNTIF(Fin_UPS!C{r1}:E{r1},"OK")=3,COUNTIF(Fin_FDX!C{r2}:E{r2},"OK")=3),"OK","CHECK")'.format(
        r1=max(FROW["UPS"].values()) + 2, r2=max(FROW["FDX"].values()) + 2)),
    ("Forecast ties to DCF engine", '=IF(COUNTIF(Ex05_Forecast!C1:G40,"CHECK")=0,"OK","CHECK")'),
    ("Valuation exhibit ties to DCF engine", '=IF(AND(Ex07_Valuation!B27="OK",Ex07_Valuation!G27="OK"),"OK","CHECK")'),
    ("Sensitivity grid centre = base value", f'=IF(AND(ABS({VR["UPS"]["grid_center"]}-{VR["UPS"]["vps"]})<0.01,ABS({VR["FDX"]["grid_center"]}-{VR["FDX"]["vps"]})<0.01),"OK","CHECK")'),
    ("DuPont product = ROE", '=IF(SUMPRODUCT(ABS(Ex03_DuPont!B6:B8-Ex03_DuPont!M6:M8))+SUMPRODUCT(ABS(Ex03_DuPont!B10:B12-Ex03_DuPont!M10:M12))<0.0001,"OK","CHECK")'),
]
for i, (lab, f) in enumerate(checks):
    rr = 24 + i
    put(sm, f"A{rr}", lab)
    put(sm, f"B{rr}", f, GREEN, bold=True, align="center")
ch = BarChart()
ch.type = "col"
for rr, colr in ((SR["Share price at valuation date"], "A5A5A5"), (SR["DCF value per share (base case)"], NAVY), (SR["Value with all strategy layers"], "70AD47")):
    ch.add_data(Reference(sm, min_col=1, max_col=3, min_row=rr), from_rows=True, titles_from_data=True)
    ch.series[-1].graphicalProperties.solidFill = colr
    ch.series[-1].graphicalProperties.line.solidFill = colr
ch.set_categories(Reference(sm, min_col=2, max_col=3, min_row=4))
tidy_chart(ch, "Share price vs DCF value per share", '$0')
sm.add_chart(ch, "A31")

# ------------------------------------------------------------------- Cover
cv = ws_cover
title(cv, "UPS vs FedEx: Valuation, Strategy and Forecasting Model", "Portfolio project by Sai Pritam Gattoju | Built from the companies' Form 10-K filings | Educational analysis, not investment advice")
widths(cv, {"A": 24, "B": 70, "C": 30})
put(cv, "A4", "What this model answers", NAVY, bold=True, size=11)
cv.merge_cells("A5:C5")
put(cv, "A5", "Are UPS and FedEx shares priced above or below what their cash flows justify, which assumptions drive that answer, and which "
    "strategic moves would add the most value per share? The model compares five years of shareholder returns, runs ratio and DuPont "
    "analysis, rebuilds free cash flow from the 10-Ks, forecasts five years of cash flow, estimates WACC with CAPM, values each company "
    "with a DCF, tests every value driver and strategy scenario, and backs out the operating margin each share price implies.", wrap=True)
cv.row_dimensions[5].height = 80
put(cv, "A7", "How to use it", NAVY, bold=True, size=11)
for i, txt in enumerate(["1. Change any blue or yellow cell on Assumptions (growth, margin, capex, beta, terminal growth...). Everything recalculates.",
                         "2. Read the answer on Summary, then the detail in Exhibits 01-09.",
                         "3. DCF_Engine holds one row per scenario; Exhibits 07-09 read from it, so every scenario uses identical math."]):
    put(cv, f"A{8 + i}", txt)
put(cv, "A12", "Color legend", NAVY, bold=True, size=11)
for i, (lab, colr, fill, desc) in enumerate([("Blue text", BLUE, None, "Reported data or an input you can change"),
                                             ("Black text", BLACK, None, "Formula"),
                                             ("Green text", GREEN, None, "Formula linking to another sheet"),
                                             ("Yellow fill", BLUE, YELLOW, "Key judgment call")]):
    put(cv, f"A{13 + i}", lab, colr, bold=True, fill=fill)
    put(cv, f"B{13 + i}", desc)
put(cv, "A18", "Sheet guide", NAVY, bold=True, size=11)
guide = [("Summary", "Value vs price, verdict, implied margin, key multiples and model checks"),
         ("Assumptions", "Market inputs, forecast drivers, value-driver shocks, strategy scenarios"),
         ("Fin_UPS / Fin_FDX", "Income statement, balance sheet and cash flow lines from the 10-Ks, with sources"),
         ("Ex01_Returns", "Total shareholder return vs S&P 500 and Dow Jones Transports"),
         ("Ex02_Ratios", "Liquidity, activity, leverage and profitability ratios with trend and peer verdicts"),
         ("Ex03_DuPont", "Five-factor ROE disaggregation"),
         ("Ex04_FCF", "Historical NOPAT, reinvestment and free cash flow"),
         ("Ex05_Forecast", "Five-year free cash flow forecast"),
         ("Ex06_WACC", "Cost of equity, cost of debt and WACC"),
         ("Ex07_Valuation", "DCF, multiples cross-check, reverse DCF and WACC x growth sensitivity"),
         ("Ex08_Value_Drivers", "Value per share impact of a one-point change in each driver"),
         ("Ex09_Strategy", "Cumulative strategy scenarios"),
         ("DCF_Engine", "One row per scenario: the calculation core")]
header(cv, 19, 1, ["Sheet", "What it shows"])
for i, (s, d) in enumerate(guide):
    put(cv, f"A{20 + i}", s, bold=True)
    put(cv, f"B{20 + i}", d)
put(cv, "A34", "Sources", NAVY, bold=True, size=11)
for i, s in enumerate(["UPS Form 10-K for fiscal 2025 and 2024 (SEC / investors.ups.com); UPS 4Q 2025 earnings release",
                       "FedEx FY26 Annual Report / Form 10-K and FY25 Form 10-K (SEC / investors.fedex.com)",
                       "Share prices: Yahoo Finance history; 10-year Treasury yields; Damodaran implied equity risk premium (Jan 2026)"]):
    put(cv, f"A{35 + i}", s, GREY, size=9)
put(cv, "A39", "Educational analysis using public data. Not investment advice.", GREY, italic=True, size=9)

# -------------------------------------------------------------- finishing
tab = {"Cover": NAVY, "Summary": NAVY, "Assumptions": "FFC000", "DCF_Engine": "A5A5A5"}
for ws in wb.worksheets:
    ws.sheet_properties.tabColor = tab.get(ws.title, "8EA9DB")
OUT.parent.mkdir(exist_ok=True)
wb.save(OUT)
print(f"saved {OUT}")
