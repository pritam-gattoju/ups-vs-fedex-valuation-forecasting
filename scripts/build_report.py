"""Build report/UPS_vs_FedEx_Valuation_Report.pdf and the README charts in images/.

Every number in the text is computed from the same engine as the Excel model, so the report
cannot drift from the model. Run:  python scripts/build_report.py
"""
from __future__ import annotations

import sys
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import numpy as np  # noqa: E402
import pandas as pd  # noqa: E402
from reportlab.lib import colors  # noqa: E402
from reportlab.lib.enums import TA_JUSTIFY  # noqa: E402
from reportlab.lib.pagesizes import letter  # noqa: E402
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet  # noqa: E402
from reportlab.lib.units import inch  # noqa: E402
from reportlab.platypus import (Image, KeepTogether, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table,  # noqa: E402
                                TableStyle)

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from src.data_loader import fiscal_years, load_financials  # noqa: E402
from src.ratios import dupont, free_cash_flow, ratios  # noqa: E402
from src.valuation import (DRIVER_SHOCKS, STRATEGIES, base_inputs, forecast_table, implied_operating_margin,  # noqa: E402
                           monte_carlo, run_dcf, sensitivity_grid, strategy_layers, value_drivers, wacc_build)

IMG = ROOT / "images"
IMG.mkdir(exist_ok=True)
OUT = ROOT / "report" / "UPS_vs_FedEx_Valuation_Report.pdf"
UPS_C, FDX_C, NAVY, GREY = "#5C3D2E", "#4D148C", "#1F3864", "#9AA0A6"
NAME = {"UPS": "UPS", "FDX": "FedEx"}

# ------------------------------------------------------------------ numbers
X = {t: base_inputs(t) for t in ("UPS", "FDX")}
R = {t: run_dcf(X[t]) for t in X}
W = {t: wacc_build(t) for t in X}
IM = {t: implied_operating_margin(X[t]) for t in X}
VD = {t: value_drivers(X[t]) for t in X}
SL = {t: strategy_layers(t, X[t]) for t in X}
RA = {t: ratios(t) for t in X}
DP = {t: dupont(t) for t in X}
FCF = {t: free_cash_flow(t) for t in X}
FIN = {t: load_financials(t) for t in X}
YRS = {t: fiscal_years(t) for t in X}
MC = {t: monte_carlo(X[t]) for t in X}
tsr = pd.read_csv(ROOT / "data" / "tsr_index.csv")


def idx(t, series):
    return tsr[(tsr.company == t) & (tsr.series == series)].index_value.tolist()


TS = {"UPS": (idx("UPS", "United Parcel Service"), idx("UPS", "S&P 500 Index"), idx("UPS", "Dow Jones Transportation Average")),
      "FDX": (idx("FDX", "FedEx Corporation"), idx("FDX", "S&P 500 Index"), idx("FDX", "Dow Jones Transportation Average"))}
TS_DATES = {t: list(dict.fromkeys(tsr[tsr.company == t].period_end)) for t in X}


def m(v, d=0):
    return f"(${abs(v):,.{d}f})" if v < 0 else f"${v:,.{d}f}"


def p(v, d=1):
    return f"{v:.{d}%}"


def bn(v):
    return f"${v / 1000:,.1f} billion"


def f_(t, item, y=-1):
    return float(FIN[t].loc[item, YRS[t][y]])


# ------------------------------------------------------------------- charts
plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 10, "axes.spines.top": False, "axes.spines.right": False})


def save(fig, name):
    fig.savefig(IMG / name, dpi=170, bbox_inches="tight")
    plt.close(fig)


fig, ax = plt.subplots(figsize=(8, 3.6))
labels = ["Share price", "DCF value (base)", "DCF value (all strategies)"]
vals = {t: [X[t].price, R[t]["value_per_share"], SL[t]["Value per share"].iloc[-1]] for t in X}
xpos = np.arange(2)
for i, (lab, col) in enumerate(zip(labels, [GREY, NAVY, "#2E8B57"])):
    bars = ax.bar(xpos + (i - 1) * 0.26, [vals["UPS"][i], vals["FDX"][i]], 0.26, label=lab, color=col)
    ax.bar_label(bars, labels=[m(v) for v in [vals["UPS"][i], vals["FDX"][i]]], fontsize=8, padding=2)
ax.set_xticks(xpos, ["UPS (Dec 31, 2025)", "FedEx (May 31, 2026, pre-spin-off)"])
ax.set_ylabel("$ per share")
ax.set_title("Share price vs intrinsic value per share", loc="left", fontsize=11)
ax.legend(frameon=False, ncol=3, fontsize=8, loc="upper left")
ax.set_ylim(0, max(max(v) for v in vals.values()) * 1.2)
save(fig, "valuation_summary.png")

fig, axes = plt.subplots(1, 2, figsize=(10, 3.4))
for ax, t, col in zip(axes, ("UPS", "FDX"), (UPS_C, FDX_C)):
    d = VD[t].sort_values("Impact per share")
    ax.barh(d["Value driver"], d["Impact per share"], color=col)
    for y, v in enumerate(d["Impact per share"]):
        ax.text(v, y, f" +${v:,.2f}", va="center", fontsize=8)
    ax.set_title(f"{NAME[t]}: $ per share from a 1-point change", loc="left", fontsize=10)
    ax.set_xlim(0, d["Impact per share"].max() * 1.3)
fig.tight_layout(w_pad=3)
save(fig, "value_drivers.png")

fig, axes = plt.subplots(1, 2, figsize=(9, 3.4))
for ax, t, col in zip(axes, ("UPS", "FDX"), (UPS_C, FDX_C)):
    d = SL[t]
    short = ["Base"] + [s.split(":")[0] for s in d["Scenario"].iloc[1:]]
    bars = ax.bar(short, d["Value per share"], color=col)
    ax.bar_label(bars, labels=[m(v) for v in d["Value per share"]], fontsize=8)
    ax.axhline(X[t].price, color="#C0504D", ls="--", lw=1.2, label=f"Share price {m(X[t].price, 2)}")
    ax.legend(frameon=False, fontsize=8, loc="upper left")
    ax.set_title(f"{NAME[t]}: cumulative strategy scenarios", loc="left", fontsize=10)
    ax.set_ylim(0, max(d["Value per share"].max(), X[t].price) * 1.3)
save(fig, "strategy_scenarios.png")

fig, axes = plt.subplots(1, 2, figsize=(9, 3.2))
for ax, t, col in zip(axes, ("UPS", "FDX"), (UPS_C, FDX_C)):
    co, spx, djt = TS[t]
    dates = [d[:7] for d in TS_DATES[t]]
    ax.plot(dates, co, color=col, lw=2.5, marker="o", label=NAME[t])
    ax.plot(dates, spx, color=GREY, lw=1.8, ls="--", label="S&P 500")
    ax.plot(dates, djt, color="#E8A33D", lw=1.5, ls=":", label="Dow Jones Transports")
    ax.set_title(f"{NAME[t]}: $100 invested, dividends reinvested", loc="left", fontsize=10)
    ax.tick_params(axis="x", labelsize=7, rotation=30)
    ax.legend(frameon=False, fontsize=7)
save(fig, "shareholder_returns.png")

fig, ax = plt.subplots(figsize=(8, 3.2))
for t, col in (("UPS", UPS_C), ("FDX", FDX_C)):
    ax.hist(MC[t] / X[t].price - 1, bins=80, alpha=0.6, color=col, label=f"{NAME[t]}: {p((MC[t] > X[t].price).mean(), 0)} of simulations above price")
ax.axvline(0, color="black", lw=1)
ax.xaxis.set_major_formatter(matplotlib.ticker.PercentFormatter(1.0))
ax.set_xlabel("Simulated value per share vs share price")
ax.set_title("Monte Carlo: 10,000 DCFs per company with varied growth, margin, capex, WACC and terminal growth", loc="left", fontsize=10)
ax.legend(frameon=False, fontsize=8)
save(fig, "monte_carlo.png")

# ------------------------------------------------------------------ styles
ss = getSampleStyleSheet()
BODY = ParagraphStyle("body", parent=ss["Normal"], fontName="Helvetica", fontSize=9.6, leading=13.2, alignment=TA_JUSTIFY, spaceAfter=5)
H1 = ParagraphStyle("h1", parent=ss["Heading1"], fontName="Helvetica-Bold", fontSize=11.5, textColor=colors.HexColor(NAVY), spaceBefore=8, spaceAfter=4)
H2 = ParagraphStyle("h2", parent=ss["Heading2"], fontName="Helvetica-Bold", fontSize=9.8, spaceBefore=4, spaceAfter=2)
TITLE = ParagraphStyle("title", parent=ss["Title"], fontName="Helvetica-Bold", fontSize=15, textColor=colors.HexColor(NAVY), spaceAfter=2)
SUB = ParagraphStyle("sub", parent=BODY, fontSize=8.5, textColor=colors.HexColor("#666666"), alignment=1)
BUL = ParagraphStyle("bul", parent=BODY, leftIndent=12, bulletIndent=2, spaceAfter=3)
CAP = ParagraphStyle("cap", parent=BODY, fontSize=8, textColor=colors.HexColor("#666666"), alignment=0)


def para(text):
    return Paragraph(text, BODY)


def bullet(text):
    return Paragraph(text, BUL, bulletText="•")


def table(data, col_widths, title=None, bold_rows=(), header_rows=1):
    t = Table(data, colWidths=col_widths, repeatRows=header_rows)
    style = [("FONT", (0, 0), (-1, -1), "Helvetica", 7.8), ("FONT", (0, 0), (-1, header_rows - 1), "Helvetica-Bold", 7.8),
             ("BACKGROUND", (0, 0), (-1, header_rows - 1), colors.HexColor(NAVY)), ("TEXTCOLOR", (0, 0), (-1, header_rows - 1), colors.white),
             ("ALIGN", (1, 0), (-1, -1), "RIGHT"), ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#BFBFBF")),
             ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("TOPPADDING", (0, 0), (-1, -1), 2), ("BOTTOMPADDING", (0, 0), (-1, -1), 2)]
    for r in bold_rows:
        style += [("FONT", (0, r), (-1, r), "Helvetica-Bold", 7.8), ("BACKGROUND", (0, r), (-1, r), colors.HexColor("#F2F2F2"))]
    t.setStyle(TableStyle(style))
    items = [Paragraph(f"<b>{title}</b>", H2)] if title else []
    return KeepTogether(items + [t, Spacer(1, 8)])


# --------------------------------------------------------------- narrative
u, fx = "UPS", "FDX"
ups_rev_chg = f_(u, "Revenue") / f_(u, "Revenue", 0) - 1
fdx_rev_chg = f_(fx, "Revenue") / f_(fx, "Revenue", 1) - 1
pt_chg = f_(u, "Purchased transportation") / f_(u, "Purchased transportation", 1) - 1
co_u, spx_u, _ = TS[u]
co_f, spx_f, _ = TS[fx]
ups_cagr, fdx_cagr = (co_u[-1] / co_u[0]) ** 0.2 - 1, (co_f[-1] / co_f[0]) ** 0.2 - 1
spx_u_cagr, spx_f_cagr = (spx_u[-1] / spx_u[0]) ** 0.2 - 1, (spx_f[-1] / spx_f[0]) ** 0.2 - 1
ups_payout = {y: FCF[u].loc["Dividends + buybacks", y] / FCF[u].loc["Levered FCF (CFO - capex)", y] for y in YRS[u]}
fdx_adj_margin = (f_(fx, "Operating income (EBIT)") + f_(fx, "One-time transformation and other charges")) / f_(fx, "Revenue")
vd_rank = {t: VD[t].sort_values("Impact per share", ascending=False).reset_index(drop=True) for t in X}


PHRASE = {"WACC -1pt": "a one-point lower WACC", "Terminal growth +1pt": "one point more terminal growth",
          "Capex % of revenue -1pt": "spending one point less of revenue on capex", "Operating margin +1pt": "one point more operating margin",
          "Revenue growth +1pt": "one point more revenue growth", "Tax rate -1pt": "a one-point lower tax rate",
          "NWC % of revenue change -1pt": "one point less working capital"}


def ratio(t, name, y=-1):
    return RA[t].loc[name].iloc[y]


story = []
story += [Paragraph("STRATEGY FOR MAXIMIZING LONG-RUN VALUE: UPS VS FEDEX", TITLE),
          Paragraph("Equity valuation, value-driver analysis and machine-learning forecasting | Sai Pritam Gattoju | "
                    "Built from Form 10-K filings | Educational analysis, not investment advice", SUB), Spacer(1, 8)]

story += [Paragraph("INTRODUCTION", H1), para(
    "This report evaluates United Parcel Service (UPS) and FedEx Corporation (FedEx), the two largest U.S. parcel carriers, to judge whether "
    "their share prices reflect the cash flows their businesses can generate. It compares five years of total shareholder returns with the "
    "S&amp;P 500, analyzes liquidity, efficiency, leverage and profitability ratios, disaggregates return on equity into five drivers, "
    "rebuilds historical free cash flow from the 10-K filings, forecasts five years of free cash flow, estimates each company's weighted "
    "average cost of capital (WACC), and values both companies with a discounted cash flow (DCF) model. A reverse DCF then backs out the "
    "operating margin each share price implies, and a value-driver analysis shows which levers matter most. The report closes with "
    "strategies to raise shareholder value, quantified in the model. A companion Python project adds risk analytics, four price-forecasting "
    "models and a 10,000-path Monte Carlo valuation.")]
story.append(para(
    f"Each company is valued at its latest fiscal year end: <b>UPS at December 31, 2025</b> ({m(X[u].price, 2)} per share) and <b>FedEx at "
    f"May 31, 2026</b> ({m(X[fx].price, 2)} per share), the last trading day before FedEx completed the spin-off of FedEx Freight. "
    f"Valuing FedEx on its consolidated pre-spin results keeps the cash flows and the share price on the same basis."))

story += [Paragraph("STRATEGIC OVERVIEW", H1), para(
    f"<b>UPS</b> is in the middle of a deliberate downsizing. It is shrinking its largest customer, Amazon, through a planned volume "
    f"\"glide-down,\" and reconfiguring its network around fewer, more automated buildings. In 2025 it closed 93 buildings to daily "
    f"operations, cut about 48,000 operational positions and reported roughly $3.5 billion of cost savings, with about $3 billion more "
    f"targeted for 2026. Revenue fell {p(abs(ups_rev_chg))} between 2023 and 2025 to {bn(f_(u, 'Revenue'))}, and purchased transportation "
    f"dropped {p(abs(pt_chg))} in 2025 alone as UPS relied less on outside carriers. UPS is redirecting capital toward "
    f"higher-yield segments such as healthcare logistics, where 2025 acquisitions lifted goodwill from {bn(f_(u, 'Goodwill', 1))} to "
    f"{bn(f_(u, 'Goodwill'))}. Its 2026 guidance calls for about $89.7 billion of revenue, a 9.6% adjusted operating margin and "
    f"about $3.0 billion of capital expenditure.")]
story.append(para(
    f"<b>FedEx</b> is restructuring in a different way. Its DRIVE cost program and Network 2.0 are merging the formerly separate Express and "
    f"Ground networks into one FedEx, and management targets $1.0 billion of structural cost reduction by the end of calendar 2026. In fiscal "
    f"2026 revenue rose {p(fdx_rev_chg)} to {bn(f_(fx, 'Revenue'))}, and capital expenditure fell to about 4% of revenue, its lowest level "
    f"in eleven years. On June 1, 2026 FedEx spun off FedEx Freight, its less-than-truckload business, distributing one Freight share for "
    f"every two FedEx shares and keeping a 19.9% stake. It has also moved its fiscal year end from May 31 to December 31."))

story += [Paragraph("FINANCIAL ANALYSIS", H1), Paragraph("Shareholder returns", H2), para(
    f"Over the five years in each company's 10-K performance graph, $100 invested in UPS fell to {m(co_u[-1], 2)} (a {p(ups_cagr)} annual "
    f"return) while the S&amp;P 500 grew to {m(spx_u[-1], 2)} ({p(spx_u_cagr)} a year). UPS lost value in each of the last four calendar "
    f"years as volume and margins declined (Exhibit 01). FedEx was far more volatile: its index fell to {m(co_f[2], 2)} by May 2023, then "
    f"nearly doubled in fiscal 2026 ({p(co_f[-1] / co_f[-2] - 1)}) as investors priced in cost savings and the Freight separation, ending "
    f"at {m(co_f[-1], 2)} ({p(fdx_cagr)} a year) against {m(spx_f[-1], 2)} for the S&amp;P 500 ({p(spx_f_cagr)} a year). Neither stock kept "
    f"pace with the broad market."), Paragraph("Ratio analysis", H2)]
story.append(bullet(
    f"<b>Profitability:</b> UPS's operating margin slid from {p(ratio(u, 'Operating margin', 0))} in 2023 to {p(ratio(u, 'Operating margin'))} "
    f"in 2025, but it still earns more than FedEx ({p(ratio(fx, 'Operating margin'))} in fiscal 2026, or {p(fdx_adj_margin)} before "
    f"one-time separation and restructuring charges). UPS's return on invested capital of {p(ratio(u, 'Return on invested capital'))} is "
    f"almost twice FedEx's {p(ratio(fx, 'Return on invested capital'))} (Exhibit 02)."))
story.append(bullet(
    f"<b>Efficiency:</b> UPS turns its assets over {ratio(u, 'Total asset turnover'):.2f} times a year versus {ratio(fx, 'Total asset turnover'):.2f} "
    f"for FedEx, whose aircraft fleet and {bn(f_(fx, 'Operating lease right-of-use assets'))} of leased facilities make it more "
    f"asset-heavy. Collection periods are similar at about {ratio(u, 'Receivables collection period (days)'):.0f} and "
    f"{ratio(fx, 'Receivables collection period (days)'):.0f} days."))
story.append(bullet(
    f"<b>Leverage and liquidity:</b> UPS is the more leveraged company. Debt to equity rose to {p(ratio(u, 'Debt / equity'), 0)} and interest "
    f"coverage fell from {ratio(u, 'Interest coverage (EBIT / interest)', 0):.1f}x to {ratio(u, 'Interest coverage (EBIT / interest)'):.1f}x "
    f"as UPS borrowed while paying out more than its free cash flow: dividends and buybacks were {p(ups_payout[YRS[u][0]], 0)} of levered free "
    f"cash flow in 2023 and {p(ups_payout[YRS[u][-1]], 0)} in 2025. FedEx's debt to equity is {p(ratio(fx, 'Debt / equity'), 0)}, and its "
    f"current ratio of {ratio(fx, 'Current ratio'):.2f} is lifted by an unusually large cash balance of {bn(f_(fx, 'Cash and cash equivalents'))} after $5.3 billion of new debt in fiscal 2026."))
story += [Paragraph("Return on equity disaggregation", H2), para(
    f"UPS's return on equity of {p(DP[u].loc['ROE'].iloc[-1])} is more than twice FedEx's {p(DP[fx].loc['ROE'].iloc[-1])}. The five-factor "
    f"DuPont breakdown shows UPS ahead on three factors (Exhibit 03): a higher operating margin ({p(DP[u].loc['Operating margin (EBIT/Sales)'].iloc[-1])} "
    f"vs {p(DP[fx].loc['Operating margin (EBIT/Sales)'].iloc[-1])}, the largest gap), faster asset turnover, and more financial leverage "
    f"({DP[u].loc['Financial leverage (Assets/Equity)'].iloc[-1]:.1f}x assets to equity vs {DP[fx].loc['Financial leverage (Assets/Equity)'].iloc[-1]:.1f}x), "
    f"the result of years of payouts shrinking its equity base. Part of UPS's advantage is therefore borrowed rather than earned. "
    f"FedEx's interest burden is above 100% because pension and interest income exceed its interest expense. UPS's ROE has fallen "
    f"since 2023 as its margin declined, even though rising leverage propped it up.")]

story += [Paragraph("VALUATION", H1), para(
    f"Free cash flow is forecast for five years from a small set of drivers (Exhibit 05). For UPS, year-one revenue and margin follow 2026 "
    f"guidance ({p(X[u].growth[0])} growth, {p(X[u].op_margin)} adjusted operating margin), with growth of 2.5-3.0% afterwards. For FedEx, "
    f"revenue grows {p(X[fx].growth[0])} in year one, slowing to {p(X[fx].growth[-1], 2)}, at a {p(X[fx].op_margin)} operating margin, its "
    f"fiscal 2026 margin before one-time charges. Capital expenditure starts at each company's guided level and normalizes slightly above "
    f"depreciation. Terminal growth is {p(X[u].terminal_g)} for both, below nominal GDP growth for mature carriers.")]
story.append(para(
    f"The WACC is {p(W[u]['WACC'], 2)} for UPS and {p(W[fx]['WACC'], 2)} for FedEx (Exhibit 06). CAPM gives costs of equity of "
    f"{p(W[u]['Cost of equity'], 2)} and {p(W[fx]['Cost of equity'], 2)}, using 5-year betas of {W[u]['Equity beta']:.2f} and "
    f"{W[fx]['Equity beta']:.2f}, 10-year Treasury yields of {p(W[u]['Risk-free rate'], 2)} and {p(W[fx]['Risk-free rate'], 2)}, and a "
    f"{p(W[u]['Equity risk premium'], 2)} equity risk premium. After-tax costs of debt are {p(W[u]['After-tax cost of debt'], 2)} and "
    f"{p(W[fx]['After-tax cost of debt'], 2)}, based on interest expense over average debt."))
story.append(para(
    f"The DCF values UPS at <b>{m(R[u]['value_per_share'], 2)} per share</b>, <b>{p(R[u]['upside'])} above</b> its {m(X[u].price, 2)} price, and "
    f"FedEx at <b>{m(R[fx]['value_per_share'], 2)}</b>, <b>{p(abs(R[fx]['upside']))} below</b> its {m(X[fx].price, 2)} price (Exhibit 07). "
    f"Unlike a model calibrated to the market, this result says the two stocks were priced on opposite assumptions. A reverse DCF makes "
    f"this concrete. UPS's price implied an operating margin of only <b>{p(IM[u], 2)}</b>, {abs(IM[u] - X[u].op_margin) * 100:.1f} points "
    f"below its own guidance, so the market was discounting the Amazon glide-down and more. FedEx's price implied <b>{p(IM[fx], 2)}</b>, "
    f"{(IM[fx] - X[fx].op_margin) * 100:.1f} points above its current margin, so investors had already paid for Network 2.0 and the "
    f"Freight separation to succeed. Across 10,000 Monte Carlo simulations, UPS's value exceeded its price in "
    f"{p((MC[u] > X[u].price).mean(), 0)} of cases and FedEx's in {p((MC[fx] > X[fx].price).mean(), 0)}. UPS shares traded near $111 in "
    f"November 2026, about 12% above the valuation-date price."))

story += [Paragraph("ANALYSIS OF THE COMPANIES' VALUE DRIVERS", H1), para(
    "Each driver was moved by one percentage point while every other assumption stayed at the base case (Exhibit 08). "
    + " ".join(f"For {NAME[t]}, the largest impacts come from {PHRASE[vd_rank[t].loc[0, 'Value driver']]} "
               f"({m(vd_rank[t].loc[0, 'Impact per share'], 2)} per share) and {PHRASE[vd_rank[t].loc[1, 'Value driver']]} "
               f"({m(vd_rank[t].loc[1, 'Impact per share'], 2)})." for t in X)
    + f" Capital intensity matters more for FedEx because it reinvests about {p(X[fx].capex_y2_5)} of revenue. A one-point margin gain is "
      f"worth {m(VD[fx].set_index('Value driver').loc['Operating margin +1pt', 'Impact per share'], 2)} per FedEx share against "
      f"{m(VD[u].set_index('Value driver').loc['Operating margin +1pt', 'Impact per share'], 2)} for UPS, because FedEx has far fewer shares. "
      f"Revenue growth adds comparatively little value at these returns, so both companies create more value by improving margins and "
      f"capital efficiency than by chasing volume. Terminal value is {p(R[u]['tv_share_of_ev'], 0)} of UPS's enterprise value and "
      f"{p(R[fx]['tv_share_of_ev'], 0)} of FedEx's, which is why the WACC and terminal growth matter so much.")]

story += [Paragraph("STRATEGY FOR INCREASING SHAREHOLDER VALUE", H1)]
ul = SL[u]["Value per share"].tolist()
fl = SL[fx]["Value per share"].tolist()
story.append(bullet(
    f"<b>UPS - finish the network reconfiguration and automate:</b> the building closures and position cuts already underway support "
    f"guidance. Converting them into an extra half point of margin through automation would lift value from {m(ul[0], 2)} to "
    f"{m(ul[1], 2)} per share (Exhibit 09)."))
story.append(bullet(
    f"<b>UPS - replace Amazon volume with higher-yield growth:</b> healthcare logistics and small and mid-sized business customers carry better "
    f"margins than the volume being shed. Adding half a point of annual growth raises value to {m(ul[2], 2)}, and holding reinvestment "
    f"at about 4.25% of revenue adds another {m(ul[3] - ul[2], 2)}, for {m(ul[3], 2)} in total."))
story.append(bullet(
    f"<b>UPS - right-size the payout:</b> dividends of about $5.4 billion a year exceeded levered free cash flow in 2023 and 2025 and pushed "
    f"leverage higher. Holding the dividend flat and pausing buybacks until free cash flow covers payouts would rebuild balance-sheet "
    f"capacity without changing the operating value above."))
story.append(bullet(
    f"<b>FedEx - deliver Network 2.0 savings:</b> the $1.0 billion structural cost target is worth about one point of margin, raising value "
    f"from {m(fl[0], 2)} to {m(fl[1], 2)} per share. Keeping capex near its fiscal 2026 level of 4% of revenue adds "
    f"{m(fl[2] - fl[1], 2)}, and faster B2B and international growth adds {m(fl[3] - fl[2], 2)}, for {m(fl[3], 2)}."))
story.append(bullet(
    f"<b>FedEx - let the spin-off unlock value:</b> even with all three strategies the DCF stays {p(abs(fl[3] / X[fx].price - 1))} below the "
    f"pre-spin price. The gap suggests the market valued FedEx Freight at a premium less-than-truckload multiple that a consolidated DCF "
    f"cannot capture, and expected margins near {p(IM[fx], 1)}. FedEx's priority should be proving that margin level as a pure parcel "
    f"company while monetizing its remaining 19.9% Freight stake."))

story += [Paragraph("DATA SCIENCE EXTENSION", H1), para(
    "A companion Python project downloads about 16 years of daily prices for UPS, FedEx, the S&amp;P 500 and the transportation sector, "
    "measures volatility, drawdowns, Sharpe ratios and rolling betas, and engineers 31 daily features. Four models (a naive random walk, "
    "ARIMA, gradient boosting and an LSTM neural network) are tested on the final 252 trading days with RMSE, MAE, MAPE and directional "
    "accuracy. Daily prices behave close to a random walk, so the naive benchmark is hard to beat. That result supports anchoring "
    "valuation in cash-flow fundamentals rather than price patterns. A Streamlit dashboard lets users change any assumption and see the "
    "valuation, sensitivity table, value drivers and Monte Carlo distribution update instantly.")]

story += [Paragraph("CONCLUSION", H1), para(
    f"UPS and FedEx were priced on opposite stories. At the end of 2025 the market treated UPS as a business in structural decline, pricing "
    f"in a {p(IM[u], 1)} margin even though management guided to {p(X[u].op_margin)}. Our base case put its value {p(R[u]['upside'], 0)} above "
    f"the share price, provided the network reconfiguration delivers and payouts come back in line with free cash flow. FedEx, after a "
    f"{p(co_f[-1] / co_f[-2] - 1, 0)} total return in fiscal 2026, was priced for success: its share price required margins of about "
    f"{p(IM[fx], 1)} that FedEx has not yet earned. For both companies the value drivers point the same way. Margin and capital "
    f"discipline create far more value than volume growth, and the cost of capital and long-run growth assumptions dominate the answer. "
    f"Investors should therefore watch operating margin delivery above all else.")]

# ------------------------------------------------------------------ exhibits
story += [PageBreak(), Paragraph("EXHIBITS", H1)]
story += [Image(str(IMG / "shareholder_returns.png"), width=7.2 * inch, height=2.56 * inch), Paragraph("Exhibit 01 - Total shareholder return (10-K performance graphs)", CAP), Spacer(1, 6)]
rows = [["Annual total return"] + [d[:7] for d in TS_DATES[u][1:]]]
for name, s in zip(["UPS", "S&P 500", "Dow Jones Transports"], TS[u]):
    rows.append([name] + [p(s[i] / s[i - 1] - 1) for i in range(1, 6)])
rows.append(["Annual total return"] + [d[:7] for d in TS_DATES[fx][1:]])
for name, s in zip(["FedEx", "S&P 500", "Dow Jones Transports"], TS[fx]):
    rows.append([name] + [p(s[i] / s[i - 1] - 1) for i in range(1, 6)])
story.append(table(rows, [1.9 * inch] + [0.95 * inch] * 5, bold_rows=(4,)))

hdr = ["Ratio"] + [f"UPS {y}" for y in YRS[u]] + [f"FedEx {y}" for y in YRS[fx]]
rows = [hdr]
for name in RA[u].index:
    pct_like = name in ("Debt / equity", "Operating margin", "Net profit margin", "Return on assets", "Return on equity", "Return on invested capital")
    fmt = (lambda v: p(v)) if pct_like else ((lambda v: f"{v:.1f}") if "days" in name else (lambda v: f"{v:.2f}x"))
    rows.append([name] + [fmt(v) for v in RA[u].loc[name]] + [fmt(v) for v in RA[fx].loc[name]])
story.append(table(rows, [2.3 * inch] + [0.82 * inch] * 6, "Exhibit 02 - Ratio analysis"))

rows = [["", "ROE", "Tax burden", "Interest burden", "Operating margin", "Asset turnover", "Fin. leverage"]]
for t in X:
    for y in YRS[t]:
        c = DP[t][y]
        rows.append([f"{NAME[t]} {y}", p(c["ROE"]), p(c["Tax burden (NI/EBT)"]), p(c["Interest burden (EBT/EBIT)"]),
                     p(c["Operating margin (EBIT/Sales)"]), f"{c['Asset turnover (Sales/Assets)']:.2f}x", f"{c['Financial leverage (Assets/Equity)']:.2f}x"])
story.append(table(rows, [1.2 * inch] + [0.98 * inch] * 6, "Exhibit 03 - Return on equity disaggregation (ROE = product of the five factors)"))

rows = [["$ millions"] + [f"UPS {y}" for y in YRS[u]] + [f"FedEx {y}" for y in YRS[fx]]]
for line in ["NOPAT", "+ Depreciation & amortization", "- Net capital expenditures", "- Increase in operating NWC", "Free cash flow (NOPAT basis)",
             "Unlevered FCF (cash-flow statement)", "Levered FCF (CFO - capex)", "Dividends + buybacks"]:
    rows.append([line] + ["n/a" if pd.isna(v) else m(v) for v in list(FCF[u].loc[line]) + list(FCF[fx].loc[line])])
story.append(table(rows, [2.3 * inch] + [0.82 * inch] * 6, "Exhibit 04 - Historical free cash flow", bold_rows=(5,)))

for t in X:
    ft = forecast_table(X[t], 2026 if t == u else 2027)
    rows = [[f"{NAME[t]} ($ millions)"] + list(ft.index)]
    for line in ["Revenue", "Growth", "EBIT", "NOPAT", "D&A", "Capex", "Change in NWC", "Free cash flow"]:
        rows.append([line] + [p(v) if line == "Growth" else m(v) for v in ft[line]])
    story.append(table(rows, [1.9 * inch] + [1.0 * inch] * 5, f"Exhibit 05{'a' if t == u else 'b'} - {NAME[t]} five-year forecast (base case)", bold_rows=(8,)))

rows = [["Component", "UPS", "FedEx"]]
for k, fmt in [("Risk-free rate", "p2"), ("Equity beta", "x"), ("Equity risk premium", "p2"), ("Cost of equity", "p2"), ("Pre-tax cost of debt", "p2"),
               ("Tax rate", "p2"), ("After-tax cost of debt", "p2"), ("Market value of equity", "m"), ("Debt", "m"), ("Weight of equity", "p1"), ("WACC", "p2")]:
    f = {"p2": lambda v: p(v, 2), "p1": lambda v: p(v), "x": lambda v: f"{v:.2f}", "m": lambda v: m(v)}[fmt]
    rows.append([k, f(W[u][k]), f(W[fx][k])])
story.append(table(rows, [2.6 * inch, 1.3 * inch, 1.3 * inch], "Exhibit 06 - Weighted average cost of capital", bold_rows=(11,)))

rows = [["$ millions except per share", "UPS", "FedEx"],
        ["PV of free cash flow (5 years)", m(R[u]["pv_fcf"]), m(R[fx]["pv_fcf"])],
        ["Terminal value", m(R[u]["terminal_value"]), m(R[fx]["terminal_value"])],
        ["PV of terminal value", m(R[u]["pv_terminal"]), m(R[fx]["pv_terminal"])],
        ["Enterprise value", m(R[u]["enterprise_value"]), m(R[fx]["enterprise_value"])],
        ["+ Cash and marketable securities", m(X[u].cash), m(X[fx].cash)],
        ["- Debt", m(X[u].debt), m(X[fx].debt)],
        ["- Pension obligations (after tax)", m(X[u].pension_after_tax), m(X[fx].pension_after_tax)],
        ["Estimated value of equity", m(R[u]["equity_value"]), m(R[fx]["equity_value"])],
        ["Diluted shares (millions)", f"{X[u].shares:,.1f}", f"{X[fx].shares:,.1f}"],
        ["Estimated value per share", m(R[u]["value_per_share"], 2), m(R[fx]["value_per_share"], 2)],
        ["Actual share price", m(X[u].price, 2), m(X[fx].price, 2)],
        ["Upside / (downside)", p(R[u]["upside"]), p(R[fx]["upside"])],
        ["Operating margin implied by the share price", p(IM[u], 2), p(IM[fx], 2)]]
story.append(table(rows, [2.9 * inch, 1.3 * inch, 1.3 * inch], "Exhibit 07 - DCF valuation and reverse DCF", bold_rows=(10, 12)))
for t in X:
    g = sensitivity_grid(X[t])
    rows = [[f"{NAME[t]}: WACC \\ terminal growth"] + list(g.columns)] + [[i] + [m(v, 2) for v in g.loc[i]] for i in g.index]
    story.append(table(rows, [2.0 * inch] + [0.95 * inch] * 5, f"Exhibit 07{'a' if t == u else 'b'} - {NAME[t]} value per share sensitivity"))

story += [Image(str(IMG / "value_drivers.png"), width=7.2 * inch, height=2.72 * inch), Paragraph("Exhibit 08 - Key value drivers: $ per share from a one-point change", CAP), Spacer(1, 6)]
rows = [["Value driver", "UPS impact", "UPS new value", "FedEx impact", "FedEx new value"]]
for (name, _), ru, rf in zip(DRIVER_SHOCKS, VD[u].itertuples(), VD[fx].itertuples()):
    rows.append([name, m(ru[3], 2), m(ru[2], 2), m(rf[3], 2), m(rf[2], 2)])
story.append(table(rows, [2.3 * inch] + [1.15 * inch] * 4))
story += [Image(str(IMG / "strategy_scenarios.png"), width=7.2 * inch, height=2.72 * inch), Paragraph("Exhibit 09 - Strategy scenarios (each layer adds to the one before)", CAP), Spacer(1, 4)]
rows = [["Scenario", "Value per share", "vs share price"]]
for t in X:
    for rr in SL[t].itertuples():
        rows.append([f"{NAME[t]} - {rr.Scenario}", m(rr[2], 2), p(rr[3])])
story.append(table(rows, [3.8 * inch, 1.3 * inch, 1.3 * inch]))
story += [Image(str(IMG / "monte_carlo.png"), width=7.0 * inch, height=2.8 * inch), Paragraph("Exhibit 10 - Monte Carlo distribution of value per share", CAP)]
story += [Spacer(1, 8), Paragraph("Sources: UPS Form 10-K for fiscal 2025 and 2024 and 4Q 2025 earnings release; FedEx FY26 Annual Report / Form 10-K and FY25 "
                                  "Form 10-K; Yahoo Finance share prices and betas; 10-year Treasury yields; Damodaran implied equity risk premium (January 2026). "
                                  "Full model: model/UPS_vs_FedEx_Valuation_Model.xlsx.", CAP)]


def footer(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica", 7.5)
    canvas.setFillColor(colors.HexColor("#888888"))
    canvas.drawString(0.65 * inch, 0.45 * inch, "UPS vs FedEx: Strategy for Maximizing Long-Run Value | Sai Pritam Gattoju")
    canvas.drawRightString(letter[0] - 0.65 * inch, 0.45 * inch, f"Page {doc.page}")
    canvas.restoreState()


OUT.parent.mkdir(exist_ok=True)
doc = SimpleDocTemplate(str(OUT), pagesize=letter, leftMargin=0.65 * inch, rightMargin=0.65 * inch, topMargin=0.6 * inch, bottomMargin=0.7 * inch,
                        title="UPS vs FedEx: Strategy for Maximizing Long-Run Value", author="Sai Pritam Gattoju")
doc.build(story, onFirstPage=footer, onLaterPages=footer)
print(f"saved {OUT}")
