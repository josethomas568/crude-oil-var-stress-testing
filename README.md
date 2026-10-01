# Crude Oil Market Risk Report — WTI / Brent VaR, Backtest & Stress Testing

A market risk report for a hypothetical crude oil position, built twice: a formula-driven
**Excel workbook** and a **Python** model. The two reconcile to the cent (36/36 checks in
`reconcile.py`).

**Position:** long 100,000 bbl WTI Cushing, short 50,000 bbl Brent (a storage trader long
physical barrels, partially hedged with Brent, which leaves WTI–Brent basis risk).
**Data:** EIA daily spot prices, WTI (RWTC) and Brent (RBRTE), 2000 to 2026-09-29.

## What it does

| Component | Method |
|---|---|
| Mark-to-market | Daily MTM value and P&L = qty × daily $ price change, per leg and total |
| Historical VaR | 1-day 95% / 99% from the empirical distribution of the last 500 daily P&L outcomes |
| Parametric VaR | z × portfolio σ, where σ = √(q′Σq) from the WTI/Brent covariance matrix; component VaR by leg |
| Monte Carlo VaR | 10,000 correlated draws via Cholesky decomposition; normal and Student-t (df = 4) variants |
| Expected Shortfall | Reported next to every VaR (Basel FRTB moved to 97.5% ES) |
| Backtest | Rolling 250-day VaR vs next-day actual P&L, 2001–2026; Kupiec POF, Christoffersen independence, Basel traffic light, exceptions by year |
| Stress tests | Negative WTI (Apr 2020), COVID/price war, 2022 Russia spike and reversal, 2008 GFC, 2014–15 glut, plus hypothetical parallel, spread and -$40 scenarios |

## Results (as of 2026-09-29)

MTM value **$3.92M** (WTI $96.16, Brent $113.96).

| 1-day, $ loss | VaR 95% | ES 95% | VaR 99% | ES 99% |
|---|---|---|---|---|
| Historical | 195,600 | 418,200 | 474,615 | 802,800 |
| Parametric | 265,690 | 333,187 | 375,771 | 430,508 |
| Monte Carlo (normal) | 263,629 | 328,459 | 365,439 | 417,356 |
| Monte Carlo (t, df=4) | 249,044 | 380,152 | 456,786 | 618,666 |

**Backtest (5,784 days):** 95% parametric and MC pass Kupiec; every 99% model is rejected
(97–123 exceptions vs 58 expected). Last 250 days: Yellow zone for historical and parametric,
Red for MC-normal.

**Worst stress:** WTI to -$40 with Brent -20% loses $12.5M (26x 99% VaR); a replay of
20 April 2020 loses $5.4M (11x).

## Findings worth talking about

1. **Fat tails.** Historical 99% VaR is 1.26x parametric on the same data, and historical ES 99% is
   1.9x parametric ES. Normal-distribution VaR understates tail risk in crude, which is why every
   99% model fails backtesting while 95% models mostly pass. The Student-t Monte Carlo closes
   most of the gap.
2. **Exceptions cluster.** At 99% they bunch in 2008, 2014 and 2026 (see `backtest_by_year.csv`). A
   rolling 250-day window reacts slowly to regime changes.
3. **One outlier distorts parametric VaR for a year.** The -$55 WTI move on 20 April 2020 more than doubled
   rolling parametric VaR for exactly 250 days, then it vanished overnight when the day rolled out of
   the window (visible as the blue "cliff" in `03_backtest.png`). Historical VaR at 99% barely moved,
   because one point is just one observation out of 250.
4. **Why $ changes, not returns.** WTI settled at -$36.98, so % and log returns are undefined across
   that day. Using dollar price changes keeps MTM, VaR and backtest on one P&L definition and
   handles negative prices. The trade-off: $ moves from a $40 market are applied to a $96 market,
   which understates risk when prices have risen (and vice versa).
5. **Hedges aren't free of risk.** The short Brent leg reduces parametric VaR (negative component
   VaR) but loses in the 2022 spike and the spread-blowout scenario. VaR on a hedged book hides
   basis risk that stress tests reveal.
6. **VaR is not a worst case.** Multi-month scenarios (2008, 2014–15) lose 4–7x the 1-day 99% VaR;
   the liquidity horizon matters as much as the confidence level.

## Files

```
crude_var.py        Python risk engine (run this first)
build_excel.py      Builds the Excel workbook with live formulas
reconcile.py        Compares the recalculated workbook against Python output
data/               EIA price history (CSV)
output/
  Crude_Market_Risk_Report.xlsx
  01_prices.png  02_pnl_distribution.png  03_backtest.png  04_stress.png
  var_summary.csv  backtest_summary.csv  backtest_by_year.csv  backtest_daily.csv
  stress_results.csv  mtm_pnl.csv  summary.json  report.txt
```

### Excel workbook

| Sheet | Contents |
|---|---|
| Summary | Snapshot, VaR/ES table, backtest results, stress table, key findings, P&L vs VaR chart |
| Inputs | Positions, lookback, confidence levels, t df (yellow cells — change these) |
| Prices_MTM | Prices, daily $ changes, MTM, P&L by leg, cumulative P&L |
| VaR | Window selection, σ, correlation, covariance, VaR/ES by method, component VaR |
| MC_Sim | 10,000 seeded draws, Cholesky-correlated shocks, simulated P&L |
| Backtest | Rolling VaR per day, exception flags, Kupiec/Basel summary, exceptions by year |
| Stress | Scenario definitions (dates or hand-set shocks), stressed prices, P&L, multiple of VaR |

Blue text = hardcoded inputs, black = formulas, green = links to other sheets. Change the
position on Inputs (try flipping Brent to +50,000) and everything recalculates.

## Run it

```bash
pip install -r requirements.txt
python crude_var.py               # uses data/*.csv
python crude_var.py --source eia  # pulls fresh data straight from eia.gov
python build_excel.py             # rebuild the workbook
# open the workbook in Excel, save, then:
python reconcile.py
```

## Data note

The bundled CSVs come from the DataHub `oil-prices` mirror of the EIA series. That mirror is missing
some days in 2000–2006 and 2011, so a few early price changes span more than one trading day. The
current VaR window (2024–2026) is complete. Running with `--source eia` uses the complete EIA files.

## Possible extensions

- Filtered historical simulation (scale returns by EWMA or GARCH volatility) to fix the slow
  reaction to regime changes
- Futures curve instead of spot: roll yield, calendar spreads, contango/backwardation
- 10-day VaR and liquidity-horizon scaling
- Incremental VaR for a proposed trade
