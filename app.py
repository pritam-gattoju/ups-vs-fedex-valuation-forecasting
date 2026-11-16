"""UPS vs FedEx - Valuation, Strategy & Forecasting dashboard.

Run locally:   streamlit run app.py
"""
from __future__ import annotations

import json
from dataclasses import replace

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from src.data_loader import COMPANIES, DATA_DIR, OUTPUT_DIR, fiscal_years, load_financials
from src.ratios import RATIO_SPECS, dupont, free_cash_flow, ratios
from src.valuation import (DRIVER_SHOCKS, STRATEGIES, DCFInputs, base_inputs, forecast_table, implied_operating_margin,
                           monte_carlo, run_dcf, shock, wacc_build)

st.set_page_config(page_title="UPS vs FedEx | Valuation & Forecasting", page_icon="📦", layout="wide")

COLOR = {t: m["color"] for t, m in COMPANIES.items()}
NAME = {t: m["name"] for t, m in COMPANIES.items()}
SHORT = {"UPS": "UPS", "FDX": "FedEx"}
VAL_DATE = {"UPS": "Dec 31, 2025", "FDX": "May 31, 2026 (before the FedEx Freight spin-off)"}
FIRST_FY = {"UPS": 2026, "FDX": 2027}
GREY = "#9AA0A6"


def money(x, d=0):
    return f"${x:,.{d}f}"


def pct(x, d=1):
    return f"{x:.{d}%}"


def style_fig(fig, height=380):
    fig.update_layout(height=height, margin=dict(l=10, r=10, t=50, b=10), legend=dict(orientation="h", y=-0.15),
                      plot_bgcolor="white", font=dict(size=12))
    fig.update_xaxes(showgrid=False)
    fig.update_yaxes(gridcolor="#ECEFF3")
    return fig


@st.cache_data
def base(t: str) -> DCFInputs:
    return base_inputs(t)


@st.cache_data
def tsr() -> pd.DataFrame:
    return pd.read_csv(DATA_DIR / "tsr_index.csv", parse_dates=["period_end"])


@st.cache_data
def forecast_outputs():
    need = ["model_metrics.csv", "actual_vs_predicted.csv", "forecast_30d.csv", "risk_stats.csv"]
    if not all((OUTPUT_DIR / f).exists() for f in need):
        return None
    out = {
        "metrics": pd.read_csv(OUTPUT_DIR / "model_metrics.csv"),
        "avp": pd.read_csv(OUTPUT_DIR / "actual_vs_predicted.csv", parse_dates=["date"]),
        "fc": pd.read_csv(OUTPUT_DIR / "forecast_30d.csv", parse_dates=["date"]),
        "risk": pd.read_csv(OUTPUT_DIR / "risk_stats.csv"),
    }
    for f, key in (("annual_returns.csv", "annual"), ("rolling_beta.csv", "beta"), ("prices.csv", "prices")):
        p = OUTPUT_DIR / f
        out[key] = pd.read_csv(p, index_col=0, parse_dates=f != "annual_returns.csv") if p.exists() else None
    meta = OUTPUT_DIR / "run_metadata.json"
    out["meta"] = json.loads(meta.read_text()) if meta.exists() else {}
    return out


# ------------------------------------------------------------------ header
st.title("📦 UPS vs FedEx: Valuation, Strategy & Forecasting")
st.caption("Built from the companies' Form 10-K filings, a formula-driven Excel model and four machine-learning forecasters. "
           "Portfolio project by Sai Pritam Gattoju · Educational analysis, not investment advice.")

tabs = st.tabs(["Overview", "Financial analysis", "DCF valuation lab", "Monte Carlo", "Price forecasting (ML)", "Methodology"])

# ================================================================= OVERVIEW
with tabs[0]:
    cols = st.columns(2)
    summary = {}
    for col, t in zip(cols, ("UPS", "FDX")):
        x = base(t)
        r = run_dcf(x)
        im = implied_operating_margin(x)
        summary[t] = (x, r, im)
        with col:
            st.subheader(NAME[t])
            st.caption(f"Valued at {VAL_DATE[t]}")
            a, b, c = st.columns(3)
            a.metric("Share price", money(x.price, 2))
            b.metric("DCF value per share", money(r["value_per_share"], 2), pct(r["upside"]), delta_color="normal")
            c.metric("WACC", pct(x.wacc, 2))
            a, b, c = st.columns(3)
            a.metric("Base-case op. margin", pct(x.op_margin))
            b.metric("Margin the price implies", pct(im), f"{(im - x.op_margin) * 100:+.1f} pts", delta_color="off")
            c.metric("Terminal value % of EV", pct(r["tv_share_of_ev"], 0))

    ups_x, ups_r, ups_im = summary["UPS"]
    fdx_x, fdx_r, fdx_im = summary["FDX"]
    st.info(
        f"**The story in one line:** UPS's price implied an operating margin of **{pct(ups_im)}**, below its own "
        f"{pct(ups_x.op_margin)} guidance, so the market was pricing in the Amazon volume loss and more. FedEx's price implied "
        f"**{pct(fdx_im)}**, well above today's {pct(fdx_x.op_margin)}, so investors were already paying for Network 2.0 "
        f"savings and the Freight spin-off to work."
    )

    left, right = st.columns([1, 1])
    with left:
        rows = []
        for t in ("UPS", "FDX"):
            x, r, _ = summary[t]
            cum = x
            for _, d in STRATEGIES[t]:
                cum = shock(cum, **d)
            rows += [{"Company": SHORT[t], "Measure": "Share price", "Value": x.price},
                     {"Company": SHORT[t], "Measure": "DCF value (base)", "Value": r["value_per_share"]},
                     {"Company": SHORT[t], "Measure": "DCF value (all strategies)", "Value": run_dcf(cum)["value_per_share"]}]
        fig = px.bar(pd.DataFrame(rows), x="Company", y="Value", color="Measure", barmode="group", text_auto="$.0f",
                     color_discrete_sequence=[GREY, "#1F3864", "#2E8B57"], title="Share price vs intrinsic value per share")
        st.plotly_chart(style_fig(fig), width="stretch")
    with right:
        d = tsr()
        fig = go.Figure()
        for t, dash in (("UPS", None), ("FDX", None)):
            s = d[(d.company == t) & (d.series != "Dow Jones Transportation Average")]
            for series, g in s.groupby("series", sort=False):
                is_co = not series.startswith("S&P")
                fig.add_trace(go.Scatter(x=g.period_end, y=g.index_value, mode="lines+markers",
                                         name=f"{series}" if is_co else f"S&P 500 ({SHORT[t]} periods)",
                                         line=dict(color=COLOR[t] if is_co else GREY, dash=None if is_co else "dot", width=3 if is_co else 2)))
        fig.update_layout(title="Total shareholder return: $100 invested 5 years ago (10-K data)", yaxis_title="$")
        st.plotly_chart(style_fig(fig), width="stretch")

# ======================================================= FINANCIAL ANALYSIS
with tabs[1]:
    st.subheader("Ratio analysis")
    st.caption("Year-end balances from the 10-Ks. UPS fiscal years end Dec 31; FedEx fiscal years end May 31.")
    tables = []
    for t in ("UPS", "FDX"):
        rt = ratios(t)
        rt.columns = pd.MultiIndex.from_product([[SHORT[t]], rt.columns])
        tables.append(rt)
    rt = pd.concat(tables, axis=1)
    stronger = []
    for name in rt.index:
        hib = RATIO_SPECS[name][0]
        u, f = rt.loc[name, ("UPS", fiscal_years("UPS")[-1])], rt.loc[name, ("FedEx", fiscal_years("FDX")[-1])]
        stronger.append("UPS" if (u > f) == hib else "FedEx")
    fmt = rt.copy().astype(object)
    for name in rt.index:
        kind = RATIO_SPECS[name][1]
        fmt.loc[name] = [f"{v:.1%}" if kind == "%" else (f"{v:.1f}" if kind == "days" else f"{v:.2f}x") for v in rt.loc[name]]
    fmt[("", "Stronger (latest)")] = stronger
    st.dataframe(fmt, width="stretch")

    c1, c2 = st.columns(2)
    with c1:
        rows = []
        for t in ("UPS", "FDX"):
            f = load_financials(t)
            for y in fiscal_years(t):
                rows.append({"Company": SHORT[t], "Fiscal year": f"{SHORT[t]} {y}", "Revenue ($B)": f.loc["Revenue", y] / 1000,
                             "Operating margin": f.loc["Operating income (EBIT)", y] / f.loc["Revenue", y]})
        df = pd.DataFrame(rows)
        fig = go.Figure()
        fig.add_bar(x=df["Fiscal year"], y=df["Revenue ($B)"], marker_color=[COLOR["UPS"]] * 3 + [COLOR["FDX"]] * 3, name="Revenue ($B)")
        fig.add_scatter(x=df["Fiscal year"], y=df["Operating margin"] * 100, yaxis="y2", mode="lines+markers", name="Operating margin (%)",
                        line=dict(color="#E8A33D", width=3))
        fig.update_layout(title="Revenue and GAAP operating margin", yaxis=dict(title="$ billions"),
                          yaxis2=dict(title="%", overlaying="y", side="right", showgrid=False))
        st.plotly_chart(style_fig(fig), width="stretch")
    with c2:
        rows = []
        for t in ("UPS", "FDX"):
            dp = dupont(t)
            for y in dp.columns:
                rows.append({"Fiscal year": f"{SHORT[t]} {y}", "ROE": dp.loc["ROE", y], "Leverage": dp.loc["Financial leverage (Assets/Equity)", y],
                             "Margin": dp.loc["Operating margin (EBIT/Sales)", y], "Turnover": dp.loc["Asset turnover (Sales/Assets)", y]})
        df = pd.DataFrame(rows)
        fig = px.bar(df, x="Fiscal year", y="ROE", text_auto=".1%", title="Return on equity (5-factor DuPont below)",
                     color=["UPS"] * 3 + ["FedEx"] * 3, color_discrete_map={"UPS": COLOR["UPS"], "FedEx": COLOR["FDX"]})
        fig.update_yaxes(tickformat=".0%")
        st.plotly_chart(style_fig(fig), width="stretch")
    st.caption("UPS's ROE beats FedEx's on three DuPont factors: a higher operating margin (the biggest gap), faster asset turnover, "
               "and more leverage (assets are ~4.5x equity after years of buybacks and dividends).")
    dps = pd.concat({SHORT[t]: dupont(t) for t in ("UPS", "FDX")}, axis=1)
    st.dataframe(dps.style.format("{:.1%}"), width="stretch")

    st.subheader("Free cash flow and shareholder payouts")
    rows = []
    for t in ("UPS", "FDX"):
        fcf = free_cash_flow(t)
        for y in fcf.columns:
            rows += [{"Year": f"{SHORT[t]} {y}", "Measure": "Levered FCF (CFO - capex)", "$M": fcf.loc["Levered FCF (CFO - capex)", y]},
                     {"Year": f"{SHORT[t]} {y}", "Measure": "Dividends + buybacks", "$M": fcf.loc["Dividends + buybacks", y]}]
    fig = px.bar(pd.DataFrame(rows), x="Year", y="$M", color="Measure", barmode="group", color_discrete_sequence=["#1F3864", "#E8A33D"],
                 title="Did free cash flow cover dividends and buybacks?")
    st.plotly_chart(style_fig(fig), width="stretch")
    st.caption("UPS's dividends and buybacks exceeded its levered free cash flow in 2023 and 2025 (and used 95% of it in 2024), "
               "which is why its dividend policy is a central strategic question.")

# ========================================================= DCF VALUATION LAB
with tabs[2]:
    t = st.radio("Company", ["UPS", "FDX"], format_func=lambda k: NAME[k], horizontal=True, key="lab_co")
    b = base(t)
    w = wacc_build(t)
    st.caption(f"Valued at {VAL_DATE[t]}. Start from the base case and move any driver; everything below recalculates.")
    if st.button("Reset to base case", key=f"reset_{t}"):
        for k in list(st.session_state):
            if k.startswith(f"{t}_"):
                del st.session_state[k]

    with st.expander("Assumptions", expanded=True):
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown("**Growth and margins**")
            growth = tuple(st.slider(f"Revenue growth Y{i + 1} (%)", -5.0, 10.0, round(b.growth[i] * 100, 2), 0.25, key=f"{t}_g{i}") / 100 for i in range(5))
            margin = st.slider("Operating margin (%)", 3.0, 15.0, b.op_margin * 100, 0.1, key=f"{t}_m") / 100
            tax = st.slider("Tax rate (%)", 15.0, 30.0, b.tax_rate * 100, 0.5, key=f"{t}_tax") / 100
        with c2:
            st.markdown("**Reinvestment**")
            da = st.slider("D&A (% of revenue)", 2.0, 7.0, b.da_pct * 100, 0.1, key=f"{t}_da") / 100
            cap1 = st.slider("Capex Y1 (% of revenue)", 2.0, 8.0, b.capex_y1 * 100, 0.05, key=f"{t}_c1") / 100
            cap2 = st.slider("Capex Y2-Y5 (% of revenue)", 2.0, 8.0, b.capex_y2_5 * 100, 0.05, key=f"{t}_c2") / 100
            nwc = st.slider("NWC (% of revenue change)", -5.0, 5.0, b.nwc_pct * 100, 0.1, key=f"{t}_nwc") / 100
        with c3:
            st.markdown("**Discount rate**")
            rf = st.slider("Risk-free rate (%)", 2.0, 6.0, w["Risk-free rate"] * 100, 0.05, key=f"{t}_rf") / 100
            beta = st.slider("Equity beta", 0.5, 2.0, float(w["Equity beta"]), 0.01, key=f"{t}_beta")
            erp = st.slider("Equity risk premium (%)", 3.0, 7.0, w["Equity risk premium"] * 100, 0.05, key=f"{t}_erp") / 100
            tg = st.slider("Terminal growth (%)", 0.0, 4.0, b.terminal_g * 100, 0.25, key=f"{t}_tg") / 100
    ke = rf + beta * erp
    wacc = w["Weight of equity"] * ke + w["Weight of debt"] * w["Pre-tax cost of debt"] * (1 - tax)
    if wacc - tg < 0.01:
        st.error("WACC must be at least 1 point above terminal growth. Lower terminal growth or raise the discount rate.")
        st.stop()
    x = replace(b, growth=growth, op_margin=margin, tax_rate=tax, da_pct=da, capex_y1=cap1, capex_y2_5=cap2, nwc_pct=nwc,
                wacc=wacc, terminal_g=tg)
    r = run_dcf(x)
    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Value per share", money(r["value_per_share"], 2), pct(r["upside"]))
    m2.metric("Share price", money(x.price, 2))
    m3.metric("WACC", pct(wacc, 2), f"Ke {pct(ke, 2)}", delta_color="off")
    m4.metric("Enterprise value", f"${r['enterprise_value'] / 1000:,.1f}B")
    m5.metric("Margin the price implies", pct(implied_operating_margin(x)))

    c1, c2 = st.columns([1.1, 1])
    with c1:
        ft = forecast_table(x, FIRST_FY[t])
        show = ft.T.copy().astype(object)
        for row in show.index:
            show.loc[row] = [f"{v:.1%}" if row == "Growth" else f"${v:,.0f}" for v in ft[row]]
        st.markdown("**Five-year forecast ($ millions)**")
        st.dataframe(show, width="stretch")
    with c2:
        fig = go.Figure(go.Waterfall(
            x=["PV of FCF", "PV of terminal value", "+ Cash", "- Debt", "- Pension (after tax)", "Equity value"],
            measure=["relative", "relative", "relative", "relative", "relative", "total"],
            y=[r["pv_fcf"], r["pv_terminal"], x.cash, -x.debt, -x.pension_after_tax, 0],
            text=[f"${v / 1000:,.1f}B" for v in [r["pv_fcf"], r["pv_terminal"], x.cash, -x.debt, -x.pension_after_tax, r["equity_value"]]],
            connector=dict(line=dict(color="#C9CED6")), increasing=dict(marker_color="#1F3864"),
            decreasing=dict(marker_color="#C0504D"), totals=dict(marker_color=COLOR[t])))
        fig.update_layout(title="From cash flows to equity value ($ millions)", showlegend=False)
        st.plotly_chart(style_fig(fig, 360), width="stretch")

    c1, c2 = st.columns(2)
    with c1:
        waccs = [wacc + k * 0.005 for k in (-2, -1, 0, 1, 2)]
        gs = [tg + k * 0.005 for k in (-2, -1, 0, 1, 2)]
        z = [[run_dcf(replace(x, wacc=wv, terminal_g=gv))["value_per_share"] if wv - gv > 0.005 else np.nan for gv in gs] for wv in waccs]
        fig = px.imshow(np.array(z), x=[pct(g) for g in gs], y=[pct(wv, 2) for wv in waccs], text_auto="$.0f", aspect="auto",
                        color_continuous_scale="RdYlGn", labels=dict(x="Terminal growth", y="WACC", color="$/share"))
        fig.update_layout(title=f"Value per share: WACC x terminal growth (price {money(x.price)})")
        st.plotly_chart(style_fig(fig, 360), width="stretch")
    with c2:
        v0 = r["value_per_share"]
        rows = [{"Driver": n, "Impact": run_dcf(shock(x, **d))["value_per_share"] - v0} for n, d in DRIVER_SHOCKS]
        df = pd.DataFrame(rows).sort_values("Impact")
        fig = px.bar(df, x="Impact", y="Driver", orientation="h", text_auto="$.2f", color_discrete_sequence=[COLOR[t]],
                     title="Value drivers: $ per share from a 1-point change")
        st.plotly_chart(style_fig(fig, 360), width="stretch")

    rows = [{"Scenario": "Base case (your inputs)", "Value per share": v0}]
    cum = x
    for name, d in STRATEGIES[t]:
        cum = shock(cum, **d)
        rows.append({"Scenario": name, "Value per share": run_dcf(cum)["value_per_share"]})
    df = pd.DataFrame(rows)
    fig = px.bar(df, x="Scenario", y="Value per share", text_auto="$.2f", color_discrete_sequence=[COLOR[t]],
                 title="Strategy scenarios (each step adds to the one before)")
    fig.add_hline(y=x.price, line_dash="dash", line_color=GREY, annotation_text=f"Share price {money(x.price, 2)}")
    st.plotly_chart(style_fig(fig, 360), width="stretch")

# ============================================================== MONTE CARLO
with tabs[3]:
    st.caption("Instead of one answer, run thousands of DCFs with randomly varied drivers to see the range of fair values.")
    c1, c2 = st.columns([1, 2.2])
    with c1:
        t = st.radio("Company", ["UPS", "FDX"], format_func=lambda k: NAME[k], horizontal=True, key="mc_co")
        n = st.slider("Simulations", 1_000, 50_000, 10_000, 1_000)
        sd_g = st.slider("Revenue growth uncertainty (+/- pts, 1 sd)", 0.0, 3.0, 1.0, 0.25) / 100
        sd_m = st.slider("Operating margin uncertainty (pts)", 0.0, 2.0, 0.75, 0.25) / 100
        sd_w = st.slider("WACC uncertainty (pts)", 0.0, 1.5, 0.5, 0.25) / 100
        sd_tg = st.slider("Terminal growth uncertainty (pts)", 0.0, 1.0, 0.4, 0.1) / 100
        seed = st.number_input("Random seed", 0, 10_000, 42)
    x = base(t)
    sims = monte_carlo(x, n=n, seed=int(seed), sd_growth=sd_g, sd_margin=sd_m, sd_wacc=sd_w, sd_g=sd_tg)
    sims = sims[np.isfinite(sims)]
    with c2:
        p = np.percentile(sims, [5, 25, 50, 75, 95])
        fig = px.histogram(sims, nbins=80, color_discrete_sequence=[COLOR[t]], title=f"{n:,} simulated values per share for {NAME[t]}")
        fig.add_vline(x=x.price, line_dash="dash", line_color="#C0504D", annotation_text=f"Price {money(x.price)}")
        fig.add_vline(x=p[2], line_color="#1F3864", annotation_text=f"Median {money(p[2])}", annotation_position="bottom right")
        fig.update_layout(showlegend=False, xaxis_title="Value per share ($)", yaxis_title="Simulations")
        st.plotly_chart(style_fig(fig, 420), width="stretch")
    a, b2, c, d, e = st.columns(5)
    a.metric("5th percentile", money(p[0]))
    b2.metric("25th percentile", money(p[1]))
    c.metric("Median", money(p[2]))
    d.metric("95th percentile", money(p[4]))
    e.metric("Chance value > price", pct((sims > x.price).mean(), 0))

# ======================================================== PRICE FORECASTING
with tabs[4]:
    fo = forecast_outputs()
    if fo is None:
        st.warning("No forecast results yet. Run `notebooks/01_market_data_and_ml_forecasting.ipynb` in Google Colab "
                   "(Runtime > Run all), unzip `forecast_outputs.zip` into the `outputs/` folder of this repo, then commit it.")
    else:
        meta = fo["meta"]
        if meta.get("demo_mode"):
            st.error("These results come from simulated test data, not real prices. Re-run the notebook with internet access.")
        st.caption(f"Data {meta.get('data_start', '?')} to {meta.get('data_end', '?')} · models tested on the last "
                   f"{meta.get('test_days', 252)} trading days · notebook run {meta.get('run_at_utc', '?')} UTC")
        t = st.radio("Company", ["UPS", "FDX"], format_func=lambda k: NAME[k], horizontal=True, key="fc_co")
        models = ["Naive", "ARIMA", "Gradient Boosting", "LSTM"]
        chosen = st.multiselect("Models to plot", models, default=["ARIMA", "Gradient Boosting", "LSTM"])
        last_n = st.slider("Days to show", 20, 252, 90, 10)
        d = fo["avp"][fo["avp"].ticker == t].tail(last_n)
        fig = go.Figure()
        fig.add_scatter(x=d.date, y=d.actual, name="Actual", line=dict(color="black", width=3))
        palette = {"Naive": GREY, "ARIMA": "#2E86AB", "Gradient Boosting": "#E8A33D", "LSTM": "#C0504D"}
        for mname in chosen:
            fig.add_scatter(x=d.date, y=d[mname], name=mname, line=dict(color=palette[mname], width=1.6, dash="dot"))
        fig.update_layout(title=f"{NAME[t]}: actual vs predicted next-day price (out-of-sample test)", yaxis_title="Adjusted price ($)")
        st.plotly_chart(style_fig(fig, 420), width="stretch")

        c1, c2 = st.columns([1.2, 1])
        with c1:
            mt = fo["metrics"][fo["metrics"].ticker == t].drop(columns="ticker").set_index("model")
            st.markdown("**Test-set accuracy** (lower RMSE / MAE / MAPE is better)")
            st.dataframe(mt.style.format({"RMSE": "${:.2f}", "MAE": "${:.2f}", "MAPE": "{:.2%}", "Directional accuracy": "{:.1%}",
                                          "RMSE vs naive": "{:+.2%}"}, na_rep="n/a"), width="stretch")
            best = mt["RMSE"].idxmin()
            st.caption(f"Lowest RMSE: **{best}**. The naive 'tomorrow = today' forecast is the benchmark every model has to beat; "
                       "daily prices are close to a random walk, which is why the valuation relies on fundamentals.")
        with c2:
            f30 = fo["fc"][fo["fc"].ticker == t]
            fig = go.Figure()
            fig.add_scatter(x=f30.date, y=f30.p95, line=dict(width=0), showlegend=False)
            fig.add_scatter(x=f30.date, y=f30.p5, fill="tonexty", fillcolor="rgba(31,56,100,0.15)", line=dict(width=0), name="5th-95th percentile")
            fig.add_scatter(x=f30.date, y=f30.p75, line=dict(width=0), showlegend=False)
            fig.add_scatter(x=f30.date, y=f30.p25, fill="tonexty", fillcolor="rgba(31,56,100,0.30)", line=dict(width=0), name="25th-75th percentile")
            fig.add_scatter(x=f30.date, y=f30.p50, line=dict(color="#1F3864", width=2), name="Median")
            fig.update_layout(title="30-day range from 10,000 Monte Carlo paths", yaxis_title="$")
            st.plotly_chart(style_fig(fig, 360), width="stretch")

        st.markdown("**Risk and return (last 5 years)**")
        rk = fo["risk"].set_index("Ticker")
        st.dataframe(rk.style.format({c: ("{:.2f}" if ("Beta" in c or "Sharpe" in c or "Correlation" in c) else "{:.1%}") for c in rk.columns}),
                     width="stretch")
        c1, c2 = st.columns(2)
        if fo.get("annual") is not None:
            with c1:
                ar = fo["annual"].tail(10)
                fig = px.imshow(ar.T.values, x=[str(i) for i in ar.index], y=list(ar.columns), text_auto=".0%",
                                color_continuous_scale="RdYlGn", zmin=-0.4, zmax=0.4, aspect="auto",
                                title="Calendar-year total returns (latest year to date)")
                st.plotly_chart(style_fig(fig, 320), width="stretch")
        if fo.get("beta") is not None:
            with c2:
                rb = fo["beta"]
                fig = go.Figure()
                for tk in rb.columns:
                    fig.add_scatter(x=rb.index, y=rb[tk], name=SHORT.get(tk, tk), line=dict(color=COLOR.get(tk, GREY), width=2.5))
                fig.update_layout(title="Rolling 5-year beta vs S&P 500 (WACC input)")
                st.plotly_chart(style_fig(fig, 320), width="stretch")

# ============================================================== METHODOLOGY
with tabs[5]:
    st.markdown("""
### Data
* **Financial statements:** UPS Form 10-K for fiscal 2025 and 2024; FedEx FY26 Annual Report / 10-K and FY25 10-K. Every line is in
  `data/*.csv` with its source.
* **Market inputs:** share prices at each fiscal year end, 10-year Treasury yields, Damodaran's implied equity risk premium (Jan 2026),
  and 5-year betas. **Shareholder returns** come from the performance graph in each 10-K.
* **Daily prices** for the forecasting models are downloaded with `yfinance` in the notebook (adjusted for dividends and the FedEx Freight spin-off).

### Valuation
* **Free cash flow** = revenue x operating margin x (1 - tax) + D&A - capex - change in working capital.
* **WACC** = CAPM cost of equity and after-tax cost of debt (interest / average debt), weighted at market value.
* **Terminal value** = Gordon growth on year-5 free cash flow. Equity = EV + cash - debt - after-tax pension deficit.
* **Reverse DCF:** value is linear in operating margin, so the margin that makes value equal the share price is solved exactly.
* The same engine runs in the Excel model (`model/UPS_vs_FedEx_Valuation_Model.xlsx`); `tests/` checks both give identical answers.

### Forecasting
* 31 engineered features per day (lagged returns, volatility, momentum, RSI, moving-average gaps, market and peer returns, volume).
* Models: Naive random walk, ARIMA (order picked by AIC), Gradient Boosting, and an LSTM neural network (60-day window).
* Walk-forward test on the final 252 trading days with RMSE, MAE, MAPE and directional accuracy.

### Caveats
FedEx is valued on consolidated results before the June 1, 2026 FedEx Freight spin-off, matching the pre-spin share price.
Forecast assumptions are analyst judgment and are listed on the Assumptions sheet. This is an educational project, not investment advice.
""")
