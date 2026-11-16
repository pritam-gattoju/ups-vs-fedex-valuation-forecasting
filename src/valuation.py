"""Discounted cash flow engine, WACC, value drivers, strategy scenarios and Monte Carlo.

Every formula here is mirrored cell for cell in model/UPS_vs_FedEx_Valuation_Model.xlsx
(sheet DCF_Engine). tests/test_valuation_parity.py checks that both give the same answer.
"""
from __future__ import annotations

from dataclasses import dataclass, replace

import numpy as np
import pandas as pd

from .data_loader import assumption, fiscal_years, load_financials

YEARS = 5


@dataclass(frozen=True)
class DCFInputs:
    base_revenue: float
    growth: tuple  # revenue growth for years 1-5
    op_margin: float
    tax_rate: float
    da_pct: float
    capex_y1: float
    capex_y2_5: float
    nwc_pct: float  # net working capital as % of the change in revenue
    wacc: float
    terminal_g: float
    cash: float
    debt: float
    pension_after_tax: float
    shares: float
    price: float


# --------------------------------------------------------------------- WACC
def balance(ticker: str, item: str, year: str) -> float:
    return float(load_financials(ticker).loc[item, year])


def total_debt(ticker: str, year: str) -> float:
    return balance(ticker, "Current debt", year) + balance(ticker, "Long-term debt", year)


def wacc_build(ticker: str) -> dict:
    fin = load_financials(ticker)
    yrs = fiscal_years(ticker)
    last, prev = yrs[-1], yrs[-2]
    rf = assumption(ticker, "Risk-free rate")
    beta = assumption(ticker, "Equity beta")
    erp = assumption(ticker, "Equity risk premium")
    tax = assumption(ticker, "Tax rate (forecast)")
    price = assumption(ticker, "Share price at valuation date")
    shares = assumption(ticker, "Diluted shares (millions)")
    debt, debt_prev = total_debt(ticker, last), total_debt(ticker, prev)
    kd_pre = fin.loc["Interest expense", last] / ((debt + debt_prev) / 2)
    ke = rf + beta * erp
    kd_after = kd_pre * (1 - tax)
    equity_mv = price * shares
    w_e = equity_mv / (equity_mv + debt)
    wacc = w_e * ke + (1 - w_e) * kd_after
    return {
        "Risk-free rate": rf, "Equity beta": beta, "Equity risk premium": erp, "Cost of equity": ke,
        "Pre-tax cost of debt": kd_pre, "Tax rate": tax, "After-tax cost of debt": kd_after,
        "Market value of equity": equity_mv, "Debt": debt, "Weight of equity": w_e,
        "Weight of debt": 1 - w_e, "WACC": wacc,
    }


def base_inputs(ticker: str) -> DCFInputs:
    fin = load_financials(ticker)
    last = fiscal_years(ticker)[-1]
    a = lambda n: assumption(ticker, n)  # noqa: E731
    tax = a("Tax rate (forecast)")
    return DCFInputs(
        base_revenue=float(fin.loc["Revenue", last]),
        growth=tuple(a(f"Revenue growth Y{i}") for i in range(1, YEARS + 1)),
        op_margin=a("Operating margin"),
        tax_rate=tax,
        da_pct=a("D&A % of revenue"),
        capex_y1=a("Capex % of revenue Y1"),
        capex_y2_5=a("Capex % of revenue Y2-Y5"),
        nwc_pct=a("Net working capital % of revenue change"),
        wacc=wacc_build(ticker)["WACC"],
        terminal_g=a("Terminal growth rate"),
        cash=float(fin.loc["Cash and cash equivalents", last] + fin.loc["Marketable securities", last]),
        debt=total_debt(ticker, last),
        pension_after_tax=float(fin.loc["Pension and postretirement obligations", last]) * (1 - tax),
        shares=a("Diluted shares (millions)"),
        price=a("Share price at valuation date"),
    )


# ---------------------------------------------------------------------- DCF
def run_dcf(x: DCFInputs) -> dict:
    rev, fcf = [], []
    prev = x.base_revenue
    for t in range(YEARS):
        r = prev * (1 + x.growth[t])
        capex = x.capex_y1 if t == 0 else x.capex_y2_5
        f = r * (x.op_margin * (1 - x.tax_rate) + x.da_pct - capex) - x.nwc_pct * (r - prev)
        rev.append(r)
        fcf.append(f)
        prev = r
    disc = [(1 + x.wacc) ** (t + 1) for t in range(YEARS)]
    pv_fcf = sum(f / d for f, d in zip(fcf, disc))
    tv = fcf[-1] * (1 + x.terminal_g) / (x.wacc - x.terminal_g)
    pv_tv = tv / disc[-1]
    ev = pv_fcf + pv_tv
    equity = ev + x.cash - x.debt - x.pension_after_tax
    vps = equity / x.shares
    return {"revenue": rev, "fcf": fcf, "pv_fcf": pv_fcf, "terminal_value": tv, "pv_terminal": pv_tv,
            "enterprise_value": ev, "equity_value": equity, "value_per_share": vps,
            "upside": vps / x.price - 1, "tv_share_of_ev": pv_tv / ev}


def forecast_table(x: DCFInputs, first_year_label: int) -> pd.DataFrame:
    """Line-by-line 5-year forecast (Exhibit 05)."""
    rows = []
    prev = x.base_revenue
    for t in range(YEARS):
        r = prev * (1 + x.growth[t])
        ebit = r * x.op_margin
        nopat = ebit * (1 - x.tax_rate)
        da = r * x.da_pct
        capex = r * (x.capex_y1 if t == 0 else x.capex_y2_5)
        dnwc = x.nwc_pct * (r - prev)
        rows.append({"Year": f"{first_year_label + t}F", "Revenue": r, "Growth": x.growth[t], "EBIT": ebit,
                     "NOPAT": nopat, "D&A": da, "Capex": capex, "Change in NWC": dnwc,
                     "Free cash flow": nopat + da - capex - dnwc})
        prev = r
    return pd.DataFrame(rows).set_index("Year")


def shock(x: DCFInputs, **deltas) -> DCFInputs:
    """Return inputs with additive changes (growth is applied to every year)."""
    kw = {}
    for k, d in deltas.items():
        if k == "growth":
            kw[k] = tuple(g + d for g in x.growth)
        else:
            kw[k] = getattr(x, k) + d
    return replace(x, **kw)


# Shocks are mirrored on the Assumptions sheet of the Excel model.
DRIVER_SHOCKS = [
    ("Revenue growth +1pt", {"growth": 0.01}),
    ("Operating margin +1pt", {"op_margin": 0.01}),
    ("Tax rate -1pt", {"tax_rate": -0.01}),
    ("Capex % of revenue -1pt", {"capex_y1": -0.01, "capex_y2_5": -0.01}),
    ("NWC % of revenue change -1pt", {"nwc_pct": -0.01}),
    ("WACC -1pt", {"wacc": -0.01}),
    ("Terminal growth +1pt", {"terminal_g": 0.01}),
]

STRATEGIES = {
    "UPS": [
        ("S1: Network reconfiguration & automation", {"op_margin": 0.005}),
        ("S2: + Healthcare & SMB growth", {"growth": 0.005}),
        ("S3: + Capex discipline", {"capex_y2_5": -0.0025}),
    ],
    "FDX": [
        ("S1: Network 2.0 & one FedEx savings", {"op_margin": 0.010}),
        ("S2: + Capex held at ~4% of revenue", {"capex_y2_5": -0.006}),
        ("S3: + International & B2B growth", {"growth": 0.005}),
    ],
}


def value_drivers(x: DCFInputs) -> pd.DataFrame:
    base = run_dcf(x)["value_per_share"]
    out = []
    for name, d in DRIVER_SHOCKS:
        v = run_dcf(shock(x, **d))["value_per_share"]
        out.append({"Value driver": name, "Value per share": v, "Impact per share": v - base})
    return pd.DataFrame(out)


def strategy_layers(ticker: str, x: DCFInputs) -> pd.DataFrame:
    rows = [{"Scenario": "Base case", "Value per share": run_dcf(x)["value_per_share"]}]
    cum = x
    for name, d in STRATEGIES[ticker]:
        cum = shock(cum, **d)
        rows.append({"Scenario": name, "Value per share": run_dcf(cum)["value_per_share"]})
    df = pd.DataFrame(rows)
    df["vs share price"] = df["Value per share"] / x.price - 1
    return df


def implied_operating_margin(x: DCFInputs) -> float:
    """Margin the market price implies (value per share is linear in margin)."""
    v0 = run_dcf(x)["value_per_share"]
    v1 = run_dcf(shock(x, op_margin=0.01))["value_per_share"]
    return x.op_margin + (x.price - v0) / (v1 - v0) * 0.01


def sensitivity_grid(x: DCFInputs, wacc_step=0.005, g_values=(0.01, 0.015, 0.02, 0.025, 0.03)) -> pd.DataFrame:
    waccs = [x.wacc + wacc_step * k for k in (-2, -1, 0, 1, 2)]
    grid = {f"{g:.1%}": [run_dcf(replace(x, wacc=w, terminal_g=g))["value_per_share"] for w in waccs]
            for g in g_values}
    return pd.DataFrame(grid, index=[f"{w:.2%}" for w in waccs]).rename_axis("WACC \\ g")


# ------------------------------------------------------------- Monte Carlo
def monte_carlo(x: DCFInputs, n=10_000, seed=42, sd_growth=0.01, sd_margin=0.0075, sd_wacc=0.005,
                sd_g=0.004, sd_capex=0.004) -> np.ndarray:
    """Vectorised DCF over n random draws of the key drivers. Returns value per share."""
    rng = np.random.default_rng(seed)
    g_shift = rng.normal(0, sd_growth, n)
    margin = x.op_margin + rng.normal(0, sd_margin, n)
    wacc = x.wacc + rng.normal(0, sd_wacc, n)
    tg = np.clip(x.terminal_g + rng.normal(0, sd_g, n), None, wacc - 0.01)
    capex_shift = rng.normal(0, sd_capex, n)
    prev = np.full(n, x.base_revenue)
    pv = np.zeros(n)
    f = np.zeros(n)
    for t in range(YEARS):
        r = prev * (1 + x.growth[t] + g_shift)
        capex = (x.capex_y1 if t == 0 else x.capex_y2_5) + capex_shift
        f = r * (margin * (1 - x.tax_rate) + x.da_pct - capex) - x.nwc_pct * (r - prev)
        pv += f / (1 + wacc) ** (t + 1)
        prev = r
    tv = f * (1 + tg) / (wacc - tg)
    ev = pv + tv / (1 + wacc) ** YEARS
    return (ev + x.cash - x.debt - x.pension_after_tax) / x.shares
