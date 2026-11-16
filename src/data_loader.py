"""Load the 10-K financial data and valuation assumptions shipped in /data."""
from __future__ import annotations

from functools import lru_cache
from pathlib import Path

import pandas as pd

DATA_DIR = Path(__file__).resolve().parents[1] / "data"
OUTPUT_DIR = Path(__file__).resolve().parents[1] / "outputs"

COMPANIES = {
    "UPS": {"name": "United Parcel Service", "file": "ups_financials.csv", "color": "#5C3D2E", "fy_label": "Dec FY"},
    "FDX": {"name": "FedEx", "file": "fedex_financials.csv", "color": "#4D148C", "fy_label": "May FY"},
}


@lru_cache(maxsize=None)
def load_financials(ticker: str) -> pd.DataFrame:
    """Return line items as rows and fiscal years as float columns ($ millions)."""
    df = pd.read_csv(DATA_DIR / COMPANIES[ticker]["file"])
    years = [c for c in df.columns if c.startswith("FY")]
    df[years] = df[years].apply(pd.to_numeric, errors="coerce")
    return df.set_index("line_item")


def fiscal_years(ticker: str) -> list[str]:
    return [c for c in load_financials(ticker).columns if c.startswith("FY")]


@lru_cache(maxsize=None)
def load_assumptions() -> pd.DataFrame:
    df = pd.read_csv(DATA_DIR / "valuation_assumptions.csv").set_index("parameter")
    return df


def assumption(ticker: str, name: str) -> float:
    val = load_assumptions().loc[name, ticker]
    try:
        return float(val)
    except ValueError:
        return val
