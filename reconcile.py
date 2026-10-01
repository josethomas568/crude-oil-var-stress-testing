"""Reconcile the recalculated Excel workbook against the Python model.

Run after opening/saving the workbook in Excel (or LibreOffice) so cached values exist:
    python reconcile.py
"""
import json
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

HERE = Path(__file__).resolve().parent
OUT = HERE / "output"

wb = load_workbook(OUT / "Crude_Market_Risk_Report.xlsx", data_only=True)
v, bt, st = wb["VaR"], wb["Backtest"], wb["Stress"]
var_py = pd.read_csv(OUT / "var_summary.csv", index_col=0)
bts = pd.read_csv(OUT / "backtest_summary.csv")
stress = pd.read_csv(OUT / "stress_results.csv")
summ = json.loads((OUT / "summary.json").read_text())

checks = []


def chk(name, xl, py, tol=0.01):
    ok = abs(float(xl) - float(py)) <= tol * max(1.0, abs(float(py))) / 100 + 1e-6
    checks.append((name, xl, py, "OK" if ok else "MISMATCH"))


chk("MTM value", v["B15"].value, summ["mtm_value"])
methods = ["Historical", "Parametric", "Monte Carlo (normal)", "Monte Carlo (t, df=4)"]
for i, m in enumerate(methods):
    r = 29 + i
    for col, key in zip("BCDE", ["VaR_95", "ES_95", "VaR_99", "ES_99"]):
        chk(f"{m} {key}", v[f"{col}{r}"].value, var_py.loc[m, key])
chk("Portfolio sigma = STDEV(P&L)", v["B21"].value, v["B22"].value)

xl_cols = {("Hist", 0.95): "B", ("Param", 0.95): "C", ("Hist", 0.99): "D", ("Param", 0.99): "E"}
for (m, a), col in xl_cols.items():
    row = bts[(bts.Method == m) & (bts.Confidence == a)].iloc[0]
    chk(f"Backtest {m} {a:.0%} exceptions", bt[f"{col}8"].value, row.Exceptions, tol=0)
    chk(f"Backtest {m} {a:.0%} Kupiec LR", bt[f"{col}10"].value, row.Kupiec_LR)
for i, r in stress.iterrows():
    chk(f"Stress: {r.Scenario}", st[f"L{10 + i}"].value, r.PnL_Total)

df = pd.DataFrame(checks, columns=["Check", "Excel", "Python", "Result"])
pd.set_option("display.width", 160, "display.float_format", lambda x: f"{x:,.2f}")
print(df.to_string(index=False))
print(f"\n{(df.Result == 'OK').sum()}/{len(df)} checks match")
