"""Ratio analysis, 5-factor DuPont and historical free cash flow (Exhibits 02-04)."""
from __future__ import annotations

import pandas as pd

from .data_loader import fiscal_years, load_financials

# name: (higher_is_better, format)
RATIO_SPECS = {
    "Current ratio": (True, "x"),
    "Quick ratio": (True, "x"),
    "Receivables collection period (days)": (False, "days"),
    "Fixed asset turnover": (True, "x"),
    "Total asset turnover": (True, "x"),
    "Debt / equity": (False, "%"),
    "Debt / EBITDA": (False, "x"),
    "Interest coverage (EBIT / interest)": (True, "x"),
    "Operating margin": (True, "%"),
    "Net profit margin": (True, "%"),
    "Return on assets": (True, "%"),
    "Return on equity": (True, "%"),
    "Return on invested capital": (True, "%"),
}


def _row(f: pd.DataFrame, item: str, y: str) -> float:
    return float(f.loc[item, y])


def ratios(ticker: str) -> pd.DataFrame:
    f = load_financials(ticker)
    out = {}
    for y in fiscal_years(ticker):
        g = lambda i: _row(f, i, y)  # noqa: E731
        rev, ebit, ni, ta, eq = g("Revenue"), g("Operating income (EBIT)"), g("Net income"), g("Total assets"), g("Total shareholders equity")
        debt = g("Current debt") + g("Long-term debt")
        cash = g("Cash and cash equivalents") + g("Marketable securities")
        tax_rate = g("Income tax expense") / g("Income before income taxes")
        out[y] = {
            "Current ratio": g("Total current assets") / g("Total current liabilities"),
            "Quick ratio": (cash + g("Accounts receivable net")) / g("Total current liabilities"),
            "Receivables collection period (days)": g("Accounts receivable net") / rev * 365,
            "Fixed asset turnover": rev / g("Property plant and equipment net"),
            "Total asset turnover": rev / ta,
            "Debt / equity": debt / eq,
            "Debt / EBITDA": debt / (ebit + g("Depreciation and amortization")),
            "Interest coverage (EBIT / interest)": ebit / g("Interest expense"),
            "Operating margin": ebit / rev,
            "Net profit margin": ni / rev,
            "Return on assets": ni / ta,
            "Return on equity": ni / eq,
            "Return on invested capital": ebit * (1 - tax_rate) / (debt + eq - cash),
        }
    return pd.DataFrame(out)


def dupont(ticker: str) -> pd.DataFrame:
    f = load_financials(ticker)
    out = {}
    for y in fiscal_years(ticker):
        g = lambda i: _row(f, i, y)  # noqa: E731
        ni, ebt, ebit, rev = g("Net income"), g("Income before income taxes"), g("Operating income (EBIT)"), g("Revenue")
        ta, eq = g("Total assets"), g("Total shareholders equity")
        out[y] = {"Tax burden (NI/EBT)": ni / ebt, "Interest burden (EBT/EBIT)": ebt / ebit,
                  "Operating margin (EBIT/Sales)": ebit / rev, "Asset turnover (Sales/Assets)": rev / ta,
                  "Financial leverage (Assets/Equity)": ta / eq, "ROE": ni / eq}
    return pd.DataFrame(out)


def operating_nwc(f: pd.DataFrame, y: str) -> float:
    ca = f.loc["Accounts receivable net", y] + f.loc["Inventories and supplies", y] + f.loc["Other current assets", y]
    cl = f.loc["Accounts payable", y] + f.loc["Accrued wages and benefits", y] + f.loc["Other current liabilities", y]
    return float(ca - cl)


def free_cash_flow(ticker: str) -> pd.DataFrame:
    f = load_financials(ticker)
    yrs = fiscal_years(ticker)
    out = {}
    for i, y in enumerate(yrs):
        g = lambda it: _row(f, it, y)  # noqa: E731
        t = g("Income tax expense") / g("Income before income taxes")
        nopat = g("Operating income (EBIT)") * (1 - t)
        net_capex = g("Capital expenditures") - g("Proceeds from disposals")
        dnwc = operating_nwc(f, y) - operating_nwc(f, yrs[i - 1]) if i > 0 else float("nan")
        out[y] = {
            "EBIT": g("Operating income (EBIT)"), "Effective tax rate": t, "NOPAT": nopat,
            "+ Depreciation & amortization": g("Depreciation and amortization"),
            "- Net capital expenditures": net_capex, "- Increase in operating NWC": dnwc,
            "Free cash flow (NOPAT basis)": nopat + g("Depreciation and amortization") - net_capex - dnwc,
            "Unlevered FCF (cash-flow statement)": g("Cash from operating activities") + g("Interest expense") * (1 - t) - net_capex,
            "Levered FCF (CFO - capex)": g("Cash from operating activities") - g("Capital expenditures"),
            "Dividends + buybacks": g("Dividends paid") + g("Share repurchases"),
        }
    return pd.DataFrame(out)
