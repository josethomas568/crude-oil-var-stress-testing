"""
Crude Oil Market Risk Report
============================
Hypothetical position: long WTI barrels, partially hedged with short Brent.

Computes, for the position held constant through history:
  1. Daily mark-to-market (MTM) value and P&L
  2. 1-day VaR at 95% / 99% -- historical simulation, parametric (variance-covariance),
     Monte Carlo (normal and Student-t), plus Expected Shortfall
  3. Rolling VaR backtest with exception counts, Kupiec POF test,
     Christoffersen independence test and Basel traffic light
  4. Historical and hypothetical stress scenarios

Data: EIA daily spot prices (RWTC = WTI Cushing, RBRTE = Brent Europe), $/bbl.

Usage:
  python crude_var.py                 # uses CSVs in ./data
  python crude_var.py --source eia    # downloads fresh data from eia.gov

Design note: VaR is built on *dollar price changes* (dP), not % returns, because
WTI settled at -$36.98 on 2020-04-20 and percentage/log returns are undefined
across a negative price. P&L for a linear position is exactly qty * dP, so this
also keeps MTM, VaR and backtest on one consistent P&L definition.
"""
from __future__ import annotations

import argparse
import json
from dataclasses import dataclass, field
from pathlib import Path

import numpy as np
import pandas as pd
from scipy import stats

HERE = Path(__file__).resolve().parent
DATA = HERE / "data"
OUT = HERE / "output"

EIA_URLS = {
    "WTI": "https://www.eia.gov/dnav/pet/hist_xls/RWTCd.xls",
    "Brent": "https://www.eia.gov/dnav/pet/hist_xls/RBRTEd.xls",
}


# --------------------------------------------------------------------------- config
@dataclass
class Config:
    positions: dict = field(default_factory=lambda: {"WTI": 100_000, "Brent": -50_000})  # barrels
    start: str = "2000-01-01"          # first date kept in the dataset
    var_lookback: int = 500            # days used for today's VaR (~2 years)
    bt_window: int = 250               # rolling window for the backtest (~1 year)
    confidence: tuple = (0.95, 0.99)
    n_sims: int = 10_000
    t_df: int = 4                      # degrees of freedom for the fat-tailed MC
    seed: int = 42


# --------------------------------------------------------------------------- data
def load_prices(source: str = "local", start: str = "2000-01-01") -> pd.DataFrame:
    """Return a DataFrame indexed by date with columns WTI, Brent (common dates only)."""
    series = {}
    for name, url in EIA_URLS.items():
        if source == "eia":
            raw = pd.read_excel(url, sheet_name="Data 1", skiprows=2)
            raw.columns = ["Date", "Price"]
        else:
            raw = pd.read_csv(DATA / f"{name.lower()}_daily.csv")
        raw["Date"] = pd.to_datetime(raw["Date"])
        series[name] = raw.set_index("Date")["Price"].astype(float)
    px = pd.concat(series, axis=1, sort=True).dropna()
    return px.loc[start:]


# --------------------------------------------------------------------------- MTM
def mark_to_market(px: pd.DataFrame, qty: pd.Series) -> pd.DataFrame:
    out = px.copy()
    dP = px.diff()
    for c in px.columns:
        out[f"d{c}"] = dP[c]
        out[f"PnL_{c}"] = qty[c] * dP[c]
    out["MTM"] = (px * qty).sum(axis=1)
    out["PnL"] = out[[f"PnL_{c}" for c in px.columns]].sum(axis=1)
    out.loc[out.index[0], ["PnL"] + [f"PnL_{c}" for c in px.columns]] = np.nan
    out["CumPnL"] = out["PnL"].fillna(0).cumsum()
    return out


# --------------------------------------------------------------------------- VaR
def _q(x, a):
    # numpy 'linear' == Excel PERCENTILE.INC, so the workbook reconciles exactly
    return np.quantile(x, a, method="linear")


def var_historical(dP: np.ndarray, qty: np.ndarray, alpha: float):
    pnl = dP @ qty
    var = -_q(pnl, 1 - alpha)
    es = -pnl[pnl <= -var].mean()
    return var, es


def var_parametric(dP: np.ndarray, qty: np.ndarray, alpha: float):
    cov = np.cov(dP, rowvar=False, ddof=1)
    sigma = float(np.sqrt(qty @ cov @ qty))
    z = stats.norm.ppf(alpha)
    var = z * sigma                                   # zero-mean convention for 1-day VaR
    es = sigma * stats.norm.pdf(z) / (1 - alpha)
    # component VaR: how much each leg contributes (sums to total VaR)
    comp = qty * (cov @ qty) / sigma * z
    return var, es, sigma, comp


def mc_draws(n: int, seed: int, df: int):
    rng = np.random.default_rng(seed)
    Z = rng.standard_normal((n, 2))
    W = rng.chisquare(df, n)
    return Z, W


def var_monte_carlo(dP, qty, alpha, Z, W=None, df=None):
    """Simulate correlated dP via Cholesky of the sample covariance.
    If W/df are given, scale to a multivariate Student-t with the same covariance."""
    cov = np.cov(dP, rowvar=False, ddof=1)
    L = np.linalg.cholesky(cov)
    shocks = Z @ L.T
    if W is not None:
        shocks = shocks * np.sqrt((df - 2) / W)[:, None]
    pnl = shocks @ qty
    var = -_q(pnl, 1 - alpha)
    es = -pnl[pnl <= -var].mean()
    return var, es, pnl


# --------------------------------------------------------------------------- backtest
def kupiec_pof(x: int, T: int, p: float):
    """Proportion-of-failures likelihood ratio test (H0: exception rate == p)."""
    phat = x / T
    ll0 = (T - x) * np.log(1 - p) + x * np.log(p)
    ll1 = (T - x) * np.log(1 - phat) + (x * np.log(phat) if x > 0 else 0)
    lr = -2 * (ll0 - ll1)
    return lr, 1 - stats.chi2.cdf(lr, 1)


def christoffersen_ind(hits: np.ndarray):
    """Independence test: does an exception today make one tomorrow more likely?"""
    h = hits.astype(int)
    a, b = h[:-1], h[1:]
    n00 = np.sum((a == 0) & (b == 0)); n01 = np.sum((a == 0) & (b == 1))
    n10 = np.sum((a == 1) & (b == 0)); n11 = np.sum((a == 1) & (b == 1))
    p01 = n01 / max(n00 + n01, 1); p11 = n11 / max(n10 + n11, 1)
    p = (n01 + n11) / max(n00 + n01 + n10 + n11, 1)

    def ll(k, pr):
        return k * np.log(pr) if pr > 0 and k > 0 else 0.0

    l0 = ll(n00 + n10, 1 - p) + ll(n01 + n11, p)
    l1 = ll(n00, 1 - p01) + ll(n01, p01) + ll(n10, 1 - p11) + ll(n11, p11)
    lr = -2 * (l0 - l1)
    return lr, 1 - stats.chi2.cdf(lr, 1), {"n00": n00, "n01": n01, "n10": n10, "n11": n11}


def basel_zone(x: int) -> str:
    return "Green" if x <= 4 else ("Yellow" if x <= 9 else "Red")


def rolling_backtest(mtm: pd.DataFrame, qty: np.ndarray, cfg: Config, Z) -> pd.DataFrame:
    dP = mtm[["dWTI", "dBrent"]].to_numpy()
    pnl = mtm["PnL"].to_numpy()
    w = cfg.bt_window
    zs = {a: stats.norm.ppf(a) for a in cfg.confidence}
    rows = []
    for t in range(w + 1, len(mtm)):              # row 0 has no dP
        win = dP[t - w:t]
        sp = win @ qty
        cov = np.cov(win, rowvar=False, ddof=1)
        sig = np.sqrt(qty @ cov @ qty)
        L = np.linalg.cholesky(cov)
        mc = Z @ (L.T @ qty)                       # simulated P&L, same draws every day
        r = {"Date": mtm.index[t], "PnL": pnl[t]}
        for a in cfg.confidence:
            k = int(a * 100)
            r[f"Hist{k}"] = -_q(sp, 1 - a)
            r[f"Param{k}"] = zs[a] * sig
            r[f"MC{k}"] = -_q(mc, 1 - a)
        rows.append(r)
    bt = pd.DataFrame(rows).set_index("Date")
    for col in [c for c in bt.columns if c != "PnL"]:
        bt[f"X_{col}"] = (-bt["PnL"] > bt[col]).astype(int)
    return bt


def backtest_summary(bt: pd.DataFrame, cfg: Config) -> pd.DataFrame:
    rows = []
    T = len(bt)
    last = bt.iloc[-250:]
    for method in ["Hist", "Param", "MC"]:
        for a in cfg.confidence:
            k = int(a * 100)
            hits = bt[f"X_{method}{k}"].to_numpy()
            x = int(hits.sum())
            lr_pof, p_pof = kupiec_pof(x, T, 1 - a)
            lr_ind, p_ind, _ = christoffersen_ind(hits)
            x250 = int(last[f"X_{method}{k}"].sum())
            rows.append({
                "Method": method, "Confidence": a, "Days": T,
                "Expected": round(T * (1 - a), 1), "Exceptions": x,
                "Rate": x / T, "Kupiec_LR": lr_pof, "Kupiec_p": p_pof,
                "Kupiec": "Pass" if p_pof > 0.05 else "Reject",
                "Christoffersen_LR": lr_ind, "Christoffersen_p": p_ind,
                "Last250_Exceptions": x250,
                "Basel_Zone_Last250": basel_zone(x250) if k == 99 else "",
            })
    return pd.DataFrame(rows)


# --------------------------------------------------------------------------- stress
def stress_scenarios(px: pd.DataFrame) -> pd.DataFrame:
    """Each scenario defines a shock to WTI and Brent, either in % or in $/bbl.
    Historical scenarios are measured from the data; hypothetical ones are set by hand."""
    def pct(a, b, c):
        return px.loc[b, c] / px.loc[a, c] - 1

    def usd(a, b, c):
        return px.loc[b, c] - px.loc[a, c]

    s = []
    # Negative WTI: % change is meaningless through zero, so apply the $ move
    s.append(("Negative WTI (20 Apr 2020, 1 day)", "Historical", "$",
              usd("2020-04-17", "2020-04-20", "WTI"), usd("2020-04-17", "2020-04-20", "Brent"),
              "2020-04-17 to 2020-04-20"))
    s.append(("COVID / OPEC+ price war (9 Mar 2020, 1 day)", "Historical", "%",
              pct("2020-03-06", "2020-03-09", "WTI"), pct("2020-03-06", "2020-03-09", "Brent"),
              "2020-03-06 to 2020-03-09"))
    s.append(("2022 Russia spike (24 Feb - 8 Mar 2022)", "Historical", "%",
              pct("2022-02-23", "2022-03-08", "WTI"), pct("2022-02-23", "2022-03-08", "Brent"),
              "2022-02-23 to 2022-03-08"))
    s.append(("2022 spike reversal (9 Mar 2022, 1 day)", "Historical", "%",
              pct("2022-03-08", "2022-03-09", "WTI"), pct("2022-03-08", "2022-03-09", "Brent"),
              "2022-03-08 to 2022-03-09"))
    s.append(("2008 GFC collapse (Jul - Dec 2008)", "Historical", "%",
              pct("2008-07-03", "2008-12-19", "WTI"), pct("2008-07-03", "2008-12-19", "Brent"),
              "2008-07-03 to 2008-12-19"))
    s.append(("2014-15 OPEC supply glut", "Historical", "%",
              pct("2014-06-20", "2015-01-13", "WTI"), pct("2014-06-20", "2015-01-13", "Brent"),
              "2014-06-20 to 2015-01-13"))
    s.append(("Parallel crash -30%", "Hypothetical", "%", -0.30, -0.30, ""))
    s.append(("Parallel spike +30%", "Hypothetical", "%", 0.30, 0.30, ""))
    s.append(("Spread blowout: WTI -$5, Brent +$10", "Hypothetical", "$", -5.0, 10.0,
              "Cushing glut while seaborne crude tightens"))
    s.append(("WTI to -$40 at expiry, Brent -20%", "Hypothetical", "mixed", None, -0.20,
              "WTI set to -$40/bbl"))
    return pd.DataFrame(s, columns=["Scenario", "Type", "Shock_Unit", "WTI_Shock", "Brent_Shock", "Note"])


def apply_stress(sc: pd.DataFrame, p0: pd.Series, qty: pd.Series) -> pd.DataFrame:
    out = sc.copy()
    newW, newB = [], []
    for _, r in sc.iterrows():
        if r.Shock_Unit == "%":
            w, b = p0.WTI * (1 + r.WTI_Shock), p0.Brent * (1 + r.Brent_Shock)
        elif r.Shock_Unit == "$":
            w, b = p0.WTI + r.WTI_Shock, p0.Brent + r.Brent_Shock
        else:  # mixed: WTI to a fixed level, Brent in %
            w, b = -40.0, p0.Brent * (1 + r.Brent_Shock)
        newW.append(w); newB.append(b)
    out["WTI_Stressed"] = newW
    out["Brent_Stressed"] = newB
    out["PnL_WTI"] = qty.WTI * (out.WTI_Stressed - p0.WTI)
    out["PnL_Brent"] = qty.Brent * (out.Brent_Stressed - p0.Brent)
    out["PnL_Total"] = out.PnL_WTI + out.PnL_Brent
    return out


# --------------------------------------------------------------------------- charts
def make_charts(px, mtm, hist_pnl, var_tbl, bt, stress, cfg):
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import matplotlib.ticker as mt

    plt.rcParams.update({"font.size": 10, "axes.spines.top": False, "axes.spines.right": False,
                         "axes.grid": True, "grid.alpha": 0.25, "text.parse_math": False})
    BLUE, ORANGE, RED, GREY = "#2563eb", "#ea580c", "#dc2626", "#6b7280"
    musd = mt.FuncFormatter(lambda x, _: f"${x/1e6:,.1f}M")
    kusd = mt.FuncFormatter(lambda x, _: f"${x/1e3:,.0f}k")

    # 1 prices
    fig, ax = plt.subplots(figsize=(11, 4.2))
    ax.plot(px.index, px.WTI, color=BLUE, lw=0.8, label="WTI Cushing")
    ax.plot(px.index, px.Brent, color=ORANGE, lw=0.8, label="Brent")
    ax.axhline(0, color=GREY, lw=0.6)
    ax.annotate("WTI -$36.98\n20 Apr 2020", xy=(pd.Timestamp("2020-04-20"), -36.98),
                xytext=(pd.Timestamp("2014-01-01"), -30), arrowprops=dict(arrowstyle="->", color=GREY))
    ax.set_title("EIA daily spot prices ($/bbl)", loc="left", fontweight="bold")
    ax.legend(frameon=False, loc="upper left")
    fig.tight_layout(); fig.savefig(OUT / "01_prices.png", dpi=150); plt.close(fig)

    # 2 P&L distribution with VaR lines
    fig, ax = plt.subplots(figsize=(11, 4.2))
    ax.hist(hist_pnl, bins=80, color=BLUE, alpha=0.75)
    styles = {"Historical": "-", "Parametric": "--", "Monte Carlo (normal)": ":"}
    for m, ls in styles.items():
        v = var_tbl.loc[m, "VaR_99"]
        ax.axvline(-v, color=RED, ls=ls, lw=1.4, label=f"99% VaR {m}: ${v/1e3:,.0f}k")
    ax.xaxis.set_major_formatter(kusd)
    ax.set_title(f"Historical 1-day P&L scenarios, last {cfg.var_lookback} days", loc="left", fontweight="bold")
    ax.legend(frameon=False, fontsize=9)
    fig.tight_layout(); fig.savefig(OUT / "02_pnl_distribution.png", dpi=150); plt.close(fig)

    # 3 backtest
    fig, ax = plt.subplots(figsize=(11, 4.6))
    ax.bar(bt.index, bt.PnL, width=1.5, color=GREY, alpha=0.6, label="Actual daily P&L")
    ax.plot(bt.index, -bt.Hist99, color=RED, lw=0.8, label="-99% Historical VaR")
    ax.plot(bt.index, -bt.Param99, color=BLUE, lw=0.8, label="-99% Parametric VaR")
    x = bt[bt.X_Hist99 == 1]
    ax.scatter(x.index, x.PnL, s=10, color=RED, zorder=3, label=f"Exceptions (hist 99%): {len(x)}")
    ax.yaxis.set_major_formatter(musd)
    ax.set_ylim(bt.PnL.quantile(0.0005) * 1.3, bt.PnL.quantile(0.9995) * 1.3)
    lo = ax.get_ylim()[0]
    worst = bt.PnL.idxmin()
    if bt.PnL.min() < lo:
        ax.annotate(f"{worst:%d %b %Y}: ${bt.PnL.min()/1e6:,.1f}M (off scale)", xy=(worst, lo * 0.97),
                    xytext=(pd.Timestamp("2012-01-01"), lo * 0.9), fontsize=9, color=RED,
                    arrowprops=dict(arrowstyle="->", color=RED))
    ax.set_title(f"VaR backtest, rolling {cfg.bt_window}-day window", loc="left", fontweight="bold")
    ax.legend(frameon=False, fontsize=9, loc="lower left")
    fig.tight_layout(); fig.savefig(OUT / "03_backtest.png", dpi=150); plt.close(fig)

    # 4 stress
    s = stress.sort_values("PnL_Total")
    fig, ax = plt.subplots(figsize=(11, 4.6))
    cols = [RED if v < 0 else "#16a34a" for v in s.PnL_Total]
    ax.barh(s.Scenario, s.PnL_Total, color=cols)
    v99 = var_tbl.loc["Historical", "VaR_99"]
    ax.axvline(-v99, color=GREY, ls="--", lw=1, label=f"-99% hist VaR (${v99/1e3:,.0f}k)")
    ax.xaxis.set_major_formatter(musd)
    ax.set_title("Stress scenario P&L on today's position", loc="left", fontweight="bold")
    ax.legend(frameon=False, fontsize=9)
    fig.tight_layout(); fig.savefig(OUT / "04_stress.png", dpi=150); plt.close(fig)


# --------------------------------------------------------------------------- main
def run(cfg: Config, source: str = "local", charts: bool = True) -> dict:
    OUT.mkdir(exist_ok=True)
    px = load_prices(source, cfg.start)
    qty = pd.Series(cfg.positions)[px.columns]
    q = qty.to_numpy(float)

    mtm = mark_to_market(px, qty)
    asof = mtm.index[-1]
    p0 = px.iloc[-1]

    # today's VaR on the last N price changes
    dP_win = mtm[["dWTI", "dBrent"]].dropna().to_numpy()[-cfg.var_lookback:]
    Z, W = mc_draws(cfg.n_sims, cfg.seed, cfg.t_df)
    var_rows = {}
    comp = {}
    for a in cfg.confidence:
        k = int(a * 100)
        hv, hes = var_historical(dP_win, q, a)
        pv, pes, sigma, c = var_parametric(dP_win, q, a)
        mv, mes, mc_pnl = var_monte_carlo(dP_win, q, a, Z)
        tv, tes, _ = var_monte_carlo(dP_win, q, a, Z, W, cfg.t_df)
        comp[k] = dict(zip(px.columns, c))
        for name, v, e in [("Historical", hv, hes), ("Parametric", pv, pes),
                           ("Monte Carlo (normal)", mv, mes), (f"Monte Carlo (t, df={cfg.t_df})", tv, tes)]:
            var_rows.setdefault(name, {})[f"VaR_{k}"] = v
            var_rows[name][f"ES_{k}"] = e
    var_tbl = pd.DataFrame(var_rows).T

    cov = np.cov(dP_win, rowvar=False, ddof=1)
    stats_tbl = {
        "sd_dWTI": float(np.sqrt(cov[0, 0])), "sd_dBrent": float(np.sqrt(cov[1, 1])),
        "corr": float(cov[0, 1] / np.sqrt(cov[0, 0] * cov[1, 1])),
        "sigma_portfolio": float(np.sqrt(q @ cov @ q)),
    }

    bt = rolling_backtest(mtm, q, cfg, Z)
    bt_sum = backtest_summary(bt, cfg)
    bt_year = bt.filter(like="X_").groupby(bt.index.year).sum()
    bt_year.insert(0, "Days", bt.groupby(bt.index.year).size())

    stress = apply_stress(stress_scenarios(px), p0, qty)
    stress["x_VaR99_Hist"] = stress.PnL_Total / var_tbl.loc["Historical", "VaR_99"]

    # save
    mtm.to_csv(OUT / "mtm_pnl.csv")
    var_tbl.to_csv(OUT / "var_summary.csv")
    bt.to_csv(OUT / "backtest_daily.csv")
    bt_sum.to_csv(OUT / "backtest_summary.csv", index=False)
    bt_year.to_csv(OUT / "backtest_by_year.csv")
    stress.to_csv(OUT / "stress_results.csv", index=False)
    if charts:
        make_charts(px, mtm, dP_win @ q, var_tbl, bt, stress, cfg)

    summary = {
        "as_of": str(asof.date()), "prices": p0.to_dict(), "positions": cfg.positions,
        "mtm_value": float(mtm.MTM.iloc[-1]), "last_day_pnl": float(mtm.PnL.iloc[-1]),
        "var": var_tbl.round(2).to_dict(orient="index"), "component_var": comp,
        "window_stats": stats_tbl,
    }
    (OUT / "summary.json").write_text(json.dumps(summary, indent=2, default=float))
    return {"px": px, "mtm": mtm, "var": var_tbl, "bt": bt, "bt_sum": bt_sum,
            "bt_year": bt_year, "stress": stress, "summary": summary, "Z": Z, "W": W}


def print_report(r: dict):
    s = r["summary"]
    pd.set_option("display.width", 160, "display.float_format", lambda v: f"{v:,.2f}")
    print(f"\nCRUDE OIL MARKET RISK REPORT  --  as of {s['as_of']}")
    print(f"Position: {s['positions']}   Prices: WTI ${s['prices']['WTI']:.2f}  Brent ${s['prices']['Brent']:.2f}")
    print(f"MTM value: ${s['mtm_value']:,.0f}   Last-day P&L: ${s['last_day_pnl']:,.0f}\n")
    print("1-day VaR / Expected Shortfall ($):"); print(r["var"].round(0), "\n")
    print("Backtest:"); print(r["bt_sum"][["Method", "Confidence", "Expected", "Exceptions", "Rate",
                                          "Kupiec", "Kupiec_p", "Christoffersen_p", "Last250_Exceptions",
                                          "Basel_Zone_Last250"]].to_string(index=False), "\n")
    print("Stress:"); print(r["stress"][["Scenario", "WTI_Stressed", "Brent_Stressed", "PnL_Total",
                                         "x_VaR99_Hist"]].to_string(index=False))


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--source", choices=["local", "eia"], default="local")
    ap.add_argument("--no-charts", action="store_true")
    args = ap.parse_args()
    res = run(Config(), args.source, charts=not args.no_charts)
    print_report(res)
