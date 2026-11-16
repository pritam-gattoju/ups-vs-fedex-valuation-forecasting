"""Check that the Excel model and the Python engine give identical answers.

Run after recalculating the workbook (Excel saves values automatically; in a script use LibreOffice):
    python -m pytest tests/ -q
"""
from pathlib import Path

import pytest
from openpyxl import load_workbook

from src.ratios import dupont, free_cash_flow, ratios
from src.valuation import DRIVER_SHOCKS, STRATEGIES, base_inputs, implied_operating_margin, run_dcf, shock, wacc_build

XLSX = Path(__file__).resolve().parents[1] / "model" / "UPS_vs_FedEx_Valuation_Model.xlsx"


@pytest.fixture(scope="module")
def wb():
    book = load_workbook(XLSX, data_only=True)
    if book["DCF_Engine"]["AK5"].value is None:
        pytest.skip("Workbook has no cached values yet: open and save it in Excel, or recalculate with LibreOffice")
    return book


def engine_rows(wb):
    ws = wb["DCF_Engine"]
    rows = {}
    for r in range(5, ws.max_row + 1):
        company, scen = ws[f"A{r}"].value, ws[f"B{r}"].value
        if company:
            rows[("UPS" if company == "UPS" else "FDX", scen)] = r
    return rows


@pytest.mark.parametrize("ticker", ["UPS", "FDX"])
def test_wacc(wb, ticker):
    col = "B" if ticker == "UPS" else "C"
    ws = wb["Ex06_WACC"]
    xl = {ws[f"A{r}"].value: ws[f"{col}{r}"].value for r in range(5, 23)}
    assert xl["WACC"] == pytest.approx(wacc_build(ticker)["WACC"], abs=1e-9)


@pytest.mark.parametrize("ticker", ["UPS", "FDX"])
def test_every_scenario_matches(wb, ticker):
    ws = wb["DCF_Engine"]
    rows = engine_rows(wb)
    x = base_inputs(ticker)
    expected = {"Base case": run_dcf(x)["value_per_share"]}
    for name, d in DRIVER_SHOCKS:
        expected[name] = run_dcf(shock(x, **d))["value_per_share"]
    cum = x
    for name, d in STRATEGIES[ticker]:
        cum = shock(cum, **d)
        expected[name] = run_dcf(cum)["value_per_share"]
    for scen, val in expected.items():
        assert ws[f"AK{rows[(ticker, scen)]}"].value == pytest.approx(val, rel=1e-9), scen


@pytest.mark.parametrize("ticker,col", [("UPS", "B"), ("FDX", "G")])
def test_implied_margin(wb, ticker, col):
    assert wb["Ex07_Valuation"][f"{col}40"].value == pytest.approx(implied_operating_margin(base_inputs(ticker)), abs=1e-9)


def test_ratios_dupont_fcf(wb):
    ra = wb["Ex02_Ratios"]
    xl_roe_ups = [ra[f"{c}20"].value for c in "BCD"]
    assert xl_roe_ups == pytest.approx(list(ratios("UPS").loc["Return on equity"]), abs=1e-9)
    assert wb["Ex03_DuPont"]["B10"].value == pytest.approx(dupont("FDX").loc["ROE"].iloc[0], abs=1e-9)
    assert wb["Ex04_FCF"]["D13"].value == pytest.approx(free_cash_flow("UPS").loc["Free cash flow (NOPAT basis)"].iloc[-1], abs=1e-6)


def test_model_checks_pass(wb):
    sm = wb["Summary"]
    assert [sm[f"B{r}"].value for r in range(24, 29)] == ["OK"] * 5
