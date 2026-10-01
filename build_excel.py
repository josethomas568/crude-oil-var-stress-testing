"""Build the Excel version of the crude oil market risk report.

Every calculation in the workbook is a live Excel formula driven by the Inputs sheet;
only raw prices and the seeded Monte Carlo random draws are stored as values (the draws
are the same ones the Python model uses, so the two versions reconcile exactly).
"""
from pathlib import Path

import numpy as np
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

import crude_var as cv

HERE = Path(__file__).resolve().parent
OUTFILE = HERE / "output" / "Crude_Market_Risk_Report.xlsx"

cfg = cv.Config()
px = cv.load_prices("local", cfg.start)
Z, W = cv.mc_draws(cfg.n_sims, cfg.seed, cfg.t_df)
N = len(px)
FIRST, LAST = 2, N + 1          # data rows on Prices_MTM

# ------------------------------------------------------------------ styles
F = "Arial"
f_norm = Font(name=F, size=10)
f_bold = Font(name=F, size=10, bold=True)
f_title = Font(name=F, size=14, bold=True, color="1F2937")
f_sub = Font(name=F, size=10, italic=True, color="6B7280")
f_input = Font(name=F, size=10, color="0000FF")
f_link = Font(name=F, size=10, color="008000")
f_head = Font(name=F, size=10, bold=True, color="FFFFFF")
fill_head = PatternFill("solid", fgColor="1F3A5F")
fill_key = PatternFill("solid", fgColor="FFFF00")
fill_sec = PatternFill("solid", fgColor="E5E7EB")
thin = Side(style="thin", color="D1D5DB")
box = Border(top=thin, bottom=thin, left=thin, right=thin)
USD = '$#,##0;($#,##0);"-"'
USD2 = '$#,##0.00;($#,##0.00);"-"'
PCT = '0.0%;(0.0%);"-"'
PCT2 = '0.00%'
NUM = '#,##0;(#,##0);"-"'
DATE = 'yyyy-mm-dd'

wb = Workbook()


def sheet(name, widths):
    ws = wb.create_sheet(name)
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[L(i)].width = w
    ws.sheet_view.showGridLines = False
    return ws


def put(ws, ref, value, font=f_norm, fmt=None, fill=None, align=None, border=False):
    c = ws[ref]
    c.value = value
    c.font = font
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    if align:
        c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
    if border:
        c.border = box
    return c


def header_row(ws, row, labels, col=1):
    for i, lab in enumerate(labels):
        put(ws, f"{L(col + i)}{row}", lab, f_head, fill=fill_head, align="center", border=True)


def title(ws, text, sub):
    put(ws, "A1", text, f_title)
    put(ws, "A2", sub, f_sub)


# ================================================================== Inputs
wi = sheet("Inputs", [38, 16, 70])
title(wi, "Inputs", "Blue = hardcoded input. Yellow = key assumption you can change; every other sheet recalculates.")
header_row(wi, 4, ["Parameter", "Value", "Note"])
inputs = [
    ("Position: WTI (bbl, + = long)", cfg.positions["WTI"], NUM, "Hypothetical long Cushing barrels (e.g. storage inventory)."),
    ("Position: Brent (bbl, + = long)", cfg.positions["Brent"], NUM, "Short Brent as a partial hedge; leaves basis (WTI-Brent spread) risk."),
    ("VaR lookback window (days)", cfg.var_lookback, NUM, "Number of most recent daily price changes used for today's VaR."),
    ("Confidence level 1", 0.95, PCT, "1-day VaR confidence."),
    ("Confidence level 2", 0.99, PCT, "1-day VaR confidence (Basel standard)."),
    ("Student-t degrees of freedom (MC)", cfg.t_df, NUM, "Fat-tailed Monte Carlo variant. Lower = fatter tails; must be > 2."),
    ("Backtest window (days)", cfg.bt_window, NUM, "Fixed at 250 in the Backtest sheet formulas (non-volatile ranges for speed). Python version is configurable."),
]
for i, (lab, val, fmt, note) in enumerate(inputs):
    r = 5 + i
    put(wi, f"A{r}", lab, border=True)
    put(wi, f"B{r}", val, f_input, fmt, fill_key if i < 6 else None, border=True)
    put(wi, f"C{r}", note, f_sub, border=True)
QW, QB, LB, C1, C2, DF = "Inputs!$B$5", "Inputs!$B$6", "Inputs!$B$7", "Inputs!$B$8", "Inputs!$B$9", "Inputs!$B$10"

put(wi, "A14", "Data source", f_bold)
put(wi, "A15", "EIA spot prices: WTI Cushing (RWTC) and Brent Europe (RBRTE), $/bbl, daily. "
               "https://www.eia.gov/dnav/pet/pet_pri_spt_s1_d.htm", f_sub)
put(wi, "A16", f"Common trading dates from {px.index[0].date()} to {px.index[-1].date()} ({N:,} rows).", f_sub)
put(wi, "A18", "Method notes", f_bold)
notes = [
    "P&L is computed from dollar price changes (dP), not % returns: WTI settled at -$36.98 on 2020-04-20, where % and log returns break.",
    "Daily P&L = WTI qty x dWTI + Brent qty x dBrent, for the position held constant through history (hypothetical P&L).",
    "Historical VaR = -PERCENTILE of the last N daily P&L outcomes. Parametric = z x portfolio sigma (zero mean). "
    "Monte Carlo = 10,000 correlated draws via Cholesky of the same covariance.",
    "Expected Shortfall (ES) = average loss beyond VaR, shown alongside because Basel FRTB replaced 99% VaR with 97.5% ES.",
]
for i, n in enumerate(notes):
    put(wi, f"A{19 + i}", f"{i + 1}. {n}", f_norm)

# ================================================================== Prices_MTM
wp = sheet("Prices_MTM", [12, 10, 10, 10, 10, 15, 14, 14, 14, 15])
labels = ["Date", "WTI ($/bbl)", "Brent ($/bbl)", "dWTI ($)", "dBrent ($)", "MTM value ($)",
          "P&L WTI ($)", "P&L Brent ($)", "Daily P&L ($)", "Cumulative P&L ($)"]
header_row(wp, 1, labels)
wp.freeze_panes = "B2"
for i, (d, row) in enumerate(px.iterrows()):
    r = FIRST + i
    wp.cell(r, 1, d.to_pydatetime()).number_format = DATE
    wp.cell(r, 2, float(row.WTI)).font = f_input
    wp.cell(r, 3, float(row.Brent)).font = f_input
    wp.cell(r, 6, f"={QW}*B{r}+{QB}*C{r}").number_format = USD
    if r > FIRST:
        wp.cell(r, 4, f"=B{r}-B{r-1}")
        wp.cell(r, 5, f"=C{r}-C{r-1}")
        wp.cell(r, 7, f"={QW}*D{r}")
        wp.cell(r, 8, f"={QB}*E{r}")
        wp.cell(r, 9, f"=G{r}+H{r}")
        wp.cell(r, 10, f"=J{r-1}+I{r}")
    else:
        wp.cell(r, 10, 0)
    for c in (4, 5):
        wp.cell(r, c).number_format = USD2
    for c in (7, 8, 9, 10):
        wp.cell(r, c).number_format = USD
wp["B1"].comment = Comment("Source: EIA RWTC daily spot, via datasets/oil-prices mirror of EIA.", "model")
for row in wp.iter_rows(min_row=FIRST, max_row=LAST, min_col=1, max_col=10):
    for c in row:
        if c.font != f_input:
            c.font = f_norm

# ================================================================== VaR
wv = sheet("VaR", [40, 16, 16, 16, 16, 4, 34, 16])
title(wv, "1-Day Value at Risk", "Today's VaR on the last N daily price changes (N set on Inputs).")

# window definition
put(wv, "A4", "Window", f_bold, fill=fill_sec)
put(wv, "A5", "Last data row on Prices_MTM"); put(wv, "B5", "=COUNTA(Prices_MTM!A:A)", fmt=NUM)
put(wv, "A6", "First row of window"); put(wv, "B6", f"=B5-{LB}+1", fmt=NUM)
put(wv, "A7", "Window start date"); put(wv, "B7", "=INDEX(Prices_MTM!A:A,B6)", fmt=DATE)
put(wv, "A8", "As-of date"); put(wv, "B8", "=INDEX(Prices_MTM!A:A,B5)", fmt=DATE)
RW = lambda col: f"INDEX(Prices_MTM!{col}:{col},$B$6):INDEX(Prices_MTM!{col}:{col},$B$5)"
PNL, DW, DB = RW("I"), RW("D"), RW("E")

# position & market stats
put(wv, "A10", "Position and window statistics", f_bold, fill=fill_sec)
rows = [
    ("WTI price ($/bbl)", "=INDEX(Prices_MTM!B:B,B5)", USD2),
    ("Brent price ($/bbl)", "=INDEX(Prices_MTM!C:C,B5)", USD2),
    ("WTI position (bbl)", f"={QW}", NUM),
    ("Brent position (bbl)", f"={QB}", NUM),
    ("MTM value ($)", "=INDEX(Prices_MTM!F:F,B5)", USD),
    ("Gross notional ($)", "=ABS(B13*B11)+ABS(B14*B12)", USD),
    ("Std dev dWTI ($/bbl/day)", f"=STDEV({DW})", USD2),
    ("Std dev dBrent ($/bbl/day)", f"=STDEV({DB})", USD2),
    ("Correlation dWTI, dBrent", f"=CORREL({DW},{DB})", "0.000"),
    ("Covariance (sample)", "=B19*B17*B18", "0.0000"),
    ("Portfolio sigma ($/day)", "=SQRT(B13^2*B17^2+B14^2*B18^2+2*B13*B14*B20)", USD),
    ("Check: STDEV of window P&L", f"=STDEV({PNL})", USD),
]
for i, (lab, f, fmt) in enumerate(rows):
    r = 11 + i
    put(wv, f"A{r}", lab); put(wv, f"B{r}", f, fmt=fmt)
put(wv, "C22", "Must equal B21: sigma of portfolio P&L = sqrt(q'Σq)", f_sub)

# VaR table
put(wv, "A25", "1-day VaR and Expected Shortfall ($ loss)", f_bold, fill=fill_sec)
header_row(wv, 26, ["Method", "VaR @ CL1", "ES @ CL1", "VaR @ CL2", "ES @ CL2"])
put(wv, "A27", "Confidence"); put(wv, "B27", f"={C1}", fmt=PCT); put(wv, "C27", f"={C1}", fmt=PCT)
put(wv, "D27", f"={C2}", fmt=PCT); put(wv, "E27", f"={C2}", fmt=PCT)
put(wv, "A28", "z-score"); put(wv, "B28", "=NORMSINV(B27)", fmt="0.000"); put(wv, "D28", "=NORMSINV(D27)", fmt="0.000")


def var_rows():
    out = []
    # Historical
    h = ["Historical simulation"]
    for c in ("B", "D"):
        h.append(f"=-PERCENTILE({PNL},1-{c}$27)")
        h.append(f'=-AVERAGEIF({PNL},"<="&-{c}29)')
    out.append(h)
    # Parametric
    p = ["Parametric (variance-covariance)"]
    for c in ("B", "D"):
        p.append(f"={c}$28*$B$21")
        p.append(f"=$B$21*NORMDIST({c}$28,0,1,FALSE)/(1-{c}$27)")
    out.append(p)
    # MC normal / t
    m = ["Monte Carlo (normal)"]
    t = ["Monte Carlo (Student-t)"]
    for c in ("B", "D"):
        m.append(f"=-PERCENTILE(MC_Sim!$F$9:$F${8 + cfg.n_sims},1-{c}$27)")
        m.append(f'=-AVERAGEIF(MC_Sim!$F$9:$F${8 + cfg.n_sims},"<="&-{c}31)')
        t.append(f"=-PERCENTILE(MC_Sim!$G$9:$G${8 + cfg.n_sims},1-{c}$27)")
        t.append(f'=-AVERAGEIF(MC_Sim!$G$9:$G${8 + cfg.n_sims},"<="&-{c}32)')
    out += [m, t]
    return out


for i, row in enumerate(var_rows()):
    r = 29 + i
    put(wv, f"A{r}", row[0], border=True)
    for j, f in enumerate(row[1:]):
        col = L(2 + j)
        # fix ES self-references to the VaR cell on the same row
        put(wv, f"{col}{r}", f.replace("B29", f"B{r}").replace("D29", f"D{r}")
            .replace("B31", f"B{r}").replace("D31", f"D{r}")
            .replace("B32", f"B{r}").replace("D32", f"D{r}"), fmt=USD, border=True)
put(wv, "A33", "VaR as % of gross notional (hist, CL2)"); put(wv, "D33", "=D29/B16", fmt=PCT2)
put(wv, "A34", "Ratio hist / parametric (CL2): >1 = fat tails"); put(wv, "D34", "=D29/D30", fmt="0.00x")

# component VaR
put(wv, "A37", "Parametric component VaR @ CL2 (sums to total)", f_bold, fill=fill_sec)
header_row(wv, 38, ["Leg", "Component VaR ($)", "% of total"])
put(wv, "A39", "WTI"); put(wv, "B39", "=B13*(B13*B17^2+B14*B20)/B21*D28", fmt=USD); put(wv, "C39", "=B39/B41", fmt=PCT)
put(wv, "A40", "Brent"); put(wv, "B40", "=B14*(B14*B18^2+B13*B20)/B21*D28", fmt=USD); put(wv, "C40", "=B40/B41", fmt=PCT)
put(wv, "A41", "Total", f_bold); put(wv, "B41", "=B39+B40", f_bold, USD); put(wv, "C41", "=C39+C40", f_bold, PCT)
put(wv, "A42", "A negative component means that leg is hedging (reducing) total risk.", f_sub)

# ================================================================== MC_Sim
wm = sheet("MC_Sim", [10, 12, 12, 12, 14, 16, 16])
title(wm, "Monte Carlo simulation (10,000 draws)",
      "Z1, Z2 ~ N(0,1) and W ~ chi-square(df) are seeded draws (seed 42), identical to the Python model. "
      "Replace with =NORMSINV(RAND()) to re-randomise.")
put(wm, "A4", "Cholesky of window covariance", f_bold)
put(wm, "A5", "L11"); put(wm, "B5", "=VaR!B17", f_link, "0.0000")
put(wm, "C5", "L21"); put(wm, "D5", "=VaR!B19*VaR!B18", f_norm, "0.0000")
put(wm, "E5", "L22"); put(wm, "F5", "=VaR!B18*SQRT(1-VaR!B19^2)", f_norm, "0.0000")
put(wm, "A6", "df"); put(wm, "B6", f"={DF}", f_link)
header_row(wm, 8, ["Z1", "Z2", "W (chi2)", "sim dWTI", "sim dBrent", "P&L normal ($)", "P&L Student-t ($)"])
for i in range(cfg.n_sims):
    r = 9 + i
    wm.cell(r, 1, round(float(Z[i, 0]), 10)).font = f_input
    wm.cell(r, 2, round(float(Z[i, 1]), 10)).font = f_input
    wm.cell(r, 3, round(float(W[i]), 10)).font = f_input
    wm.cell(r, 4, f"=$B$5*A{r}")
    wm.cell(r, 5, f"=$D$5*A{r}+$F$5*B{r}")
    wm.cell(r, 6, f"={QW}*D{r}+{QB}*E{r}").number_format = USD
    wm.cell(r, 7, f"=F{r}*SQRT(($B$6-2)/C{r})").number_format = USD
    for c in (1, 2, 3, 4, 5):
        wm.cell(r, c).number_format = "0.0000"
wm.freeze_panes = "A9"

# ================================================================== Backtest
BT0 = 20                      # first data row on Backtest sheet
first_src = FIRST + cfg.bt_window + 1   # Prices_MTM row with a full 250-day window before it
nbt = LAST - first_src + 1
BTL = BT0 + nbt - 1
wb_ = sheet("Backtest", [12, 14, 14, 14, 14, 14, 10, 10, 10, 10])
title(wb_, "VaR backtest (rolling 250-day window)",
      "Each day's VaR uses only the 250 days before it, then is compared with that day's actual P&L. "
      "Exception = loss larger than VaR. Monte Carlo backtest is in the Python version.")

put(wb_, "A4", "Summary", f_bold, fill=fill_sec)
header_row(wb_, 5, ["Test", "Hist CL1", "Param CL1", "Hist CL2", "Param CL2"])
flagcols = {"B": "G", "C": "H", "D": "I", "E": "J"}
conf = {"B": C1, "C": C1, "D": C2, "E": C2}
srows = [
    ("Days tested", lambda c, f: f"=COUNT($B${BT0}:$B${BTL})", NUM),
    ("Expected exceptions", lambda c, f: f"={c}6*(1-{conf[c]})", "#,##0.0"),
    ("Actual exceptions", lambda c, f: f"=SUM({f}{BT0}:{f}{BTL})", NUM),
    ("Exception rate", lambda c, f: f"={c}8/{c}6", PCT2),
    ("Kupiec POF LR", lambda c, f: f"=-2*(({c}6-{c}8)*LN({conf[c]})+{c}8*LN(1-{conf[c]})-(({c}6-{c}8)*LN(1-{c}9)+IF({c}8>0,{c}8*LN({c}9),0)))", "0.00"),
    ("Kupiec p-value", lambda c, f: f"=CHIDIST({c}10,1)", "0.0000"),
    ("Kupiec result (5%)", lambda c, f: f'=IF({c}11>0.05,"Pass","Reject")', None),
    ("Exceptions, last 250 days", lambda c, f: f"=SUM({f}{BTL-249}:{f}{BTL})", NUM),
    ("Basel traffic light (CL2 only)", lambda c, f: (f'=IF({c}13<=4,"Green",IF({c}13<=9,"Yellow","Red"))' if c in "DE" else '="n/a"'), None),
]
for i, (lab, fn, fmt) in enumerate(srows):
    r = 6 + i
    put(wb_, f"A{r}", lab, border=True)
    for c, fcol in flagcols.items():
        put(wb_, f"{c}{r}", fn(c, fcol), fmt=fmt, border=True, align="center")
put(wb_, "A15", "Basel zones for 250 days at 99%: Green 0-4, Yellow 5-9, Red 10+. Kupiec tests whether the exception "
                "rate over the full history equals 1 - confidence.", f_sub)
for c in "BCDE":
    for v, color in (("Pass", "C6EFCE"), ("Green", "C6EFCE"), ("Reject", "FFC7CE"), ("Red", "FFC7CE"), ("Yellow", "FFEB9C")):
        wb_.conditional_formatting.add(f"{c}12:{c}14", CellIsRule(operator="equal", formula=[f'"{v}"'],
                                       fill=PatternFill("solid", fgColor=color)))

header_row(wb_, BT0 - 1, ["Date", "Actual P&L ($)", "Hist VaR CL1", "Param VaR CL1", "Hist VaR CL2",
                          "Param VaR CL2", "Exc H1", "Exc P1", "Exc H2", "Exc P2"])
for k in range(nbt):
    r, s = BT0 + k, first_src + k
    win = f"Prices_MTM!I{s - cfg.bt_window}:I{s - 1}"
    wb_.cell(r, 1, f"=Prices_MTM!A{s}").number_format = DATE
    wb_.cell(r, 2, f"=Prices_MTM!I{s}").number_format = USD
    wb_.cell(r, 3, f"=-PERCENTILE({win},1-{C1})").number_format = USD
    wb_.cell(r, 4, f"=NORMSINV({C1})*STDEV({win})").number_format = USD
    wb_.cell(r, 5, f"=-PERCENTILE({win},1-{C2})").number_format = USD
    wb_.cell(r, 6, f"=NORMSINV({C2})*STDEV({win})").number_format = USD
    for j, vc in enumerate("CDEF"):
        wb_.cell(r, 7 + j, f"=IF(-B{r}>{vc}{r},1,0)")
wb_.freeze_panes = f"B{BT0}"

# exceptions by year
years = sorted(set(px.index[first_src - FIRST:].year))
YC = 12
put(wb_, f"{L(YC)}4", "Exceptions by year (clusters in 2008, 2014, 2026)", f_bold)
header_row(wb_, 5, ["Year", "Days", "Hist CL1", "Param CL1", "Hist CL2", "Param CL2"], col=YC)
for c in range(YC, YC + 6):
    wb_.column_dimensions[L(c)].width = 11
for i, y in enumerate(years):
    r = 6 + i
    put(wb_, f"{L(YC)}{r}", str(y), border=True, align="center")
    yr = f'(TEXT($A${BT0}:$A${BTL},"yyyy")=${L(YC)}{r})'
    put(wb_, f"{L(YC+1)}{r}", f"=SUMPRODUCT(--{yr})", fmt=NUM, border=True)
    for j, fcol in enumerate("GHIJ"):
        put(wb_, f"{L(YC+2+j)}{r}", f"=SUMPRODUCT({yr}*${fcol}${BT0}:${fcol}${BTL})", fmt=NUM, border=True)
YLAST = 6 + len(years) - 1

# ================================================================== Stress
ws = sheet("Stress", [44, 13, 12, 12, 12, 12, 12, 12, 14, 14, 15, 12, 40])
title(ws, "Stress scenarios on today's position",
      "Historical shocks are measured from the price history between the two dates; hypothetical shocks are set by hand (blue).")
put(ws, "A4", "Today's WTI ($/bbl)"); put(ws, "B4", "=VaR!B11", f_link, USD2)
put(ws, "A5", "Today's Brent ($/bbl)"); put(ws, "B5", "=VaR!B12", f_link, USD2)
put(ws, "A6", "99% historical VaR ($)"); put(ws, "B6", "=VaR!D29", f_link, USD)
put(ws, "A7", "Units: % = relative move, $ = absolute $/bbl move, level = price set to the value", f_sub)
header_row(ws, 9, ["Scenario", "Start date", "End date", "WTI unit", "WTI shock", "Brent unit", "Brent shock",
                   "WTI stressed", "Brent stressed", "P&L WTI ($)", "P&L Brent ($)", "P&L total ($)", "x 99% VaR"])
ws.column_dimensions["L"].width = 15
ws.column_dimensions["M"].width = 11
ws.column_dimensions["N"].width = 44
put(ws, "N9", "Note", f_head, fill=fill_head, align="center", border=True)

scen = [
    ("Negative WTI (20 Apr 2020, 1 day)", "2020-04-17", "2020-04-20", "$", "$", "Applied in $ since % change through zero is meaningless."),
    ("COVID / OPEC+ price war (9 Mar 2020, 1 day)", "2020-03-06", "2020-03-09", "%", "%", "Saudi-Russia price war + COVID demand shock."),
    ("2022 Russia spike (24 Feb - 8 Mar 2022)", "2022-02-23", "2022-03-08", "%", "%", "Spike: the short Brent leg loses, long WTI gains."),
    ("2022 spike reversal (9 Mar 2022, 1 day)", "2022-03-08", "2022-03-09", "%", "%", "Largest 1-day drop after the 2022 peak."),
    ("2008 GFC collapse (Jul - Dec 2008)", "2008-07-03", "2008-12-19", "%", "%", "Multi-month: illustrates liquidity-horizon risk."),
    ("2014-15 OPEC supply glut", "2014-06-20", "2015-01-13", "%", "%", "Multi-month shale/OPEC price war."),
    ("Parallel crash -30%", None, None, "%", "%", "Hypothetical."),
    ("Parallel spike +30%", None, None, "%", "%", "Hypothetical."),
    ("Spread blowout: WTI -$5, Brent +$10", None, None, "$", "$", "Hypothetical Cushing glut while seaborne crude tightens."),
    ("WTI to -$40 at expiry, Brent -20%", None, None, "level", "%", "Hypothetical repeat of April 2020 storage squeeze."),
]
hyp = {"Parallel crash -30%": (-0.30, -0.30), "Parallel spike +30%": (0.30, 0.30),
       "Spread blowout: WTI -$5, Brent +$10": (-5, 10), "WTI to -$40 at expiry, Brent -20%": (-40, -0.20)}
import datetime as dt
SR0 = 10
for i, (name, d0, d1, uw, ub, note) in enumerate(scen):
    r = SR0 + i
    put(ws, f"A{r}", name, border=True)
    put(ws, f"D{r}", uw, f_input, border=True, align="center")
    put(ws, f"F{r}", ub, f_input, border=True, align="center")
    if d0:
        put(ws, f"B{r}", dt.datetime.fromisoformat(d0), f_input, DATE, border=True)
        put(ws, f"C{r}", dt.datetime.fromisoformat(d1), f_input, DATE, border=True)
        for col, src, unit in (("E", "B", "D"), ("G", "C", "F")):
            p0 = f"INDEX(Prices_MTM!{src}:{src},MATCH($B{r},Prices_MTM!$A:$A,0))"
            p1 = f"INDEX(Prices_MTM!{src}:{src},MATCH($C{r},Prices_MTM!$A:$A,0))"
            put(ws, f"{col}{r}", f'=IF({unit}{r}="%",{p1}/{p0}-1,{p1}-{p0})', border=True)
    else:
        put(ws, f"B{r}", "", border=True); put(ws, f"C{r}", "", border=True)
        put(ws, f"E{r}", hyp[name][0], f_input, border=True)
        put(ws, f"G{r}", hyp[name][1], f_input, border=True)
    for col, unit in (("E", "D"), ("G", "F")):
        ws[f"{col}{r}"].number_format = PCT if (uw if col == "E" else ub) == "%" else USD2
    put(ws, f"H{r}", f'=IF(D{r}="%",$B$4*(1+E{r}),IF(D{r}="$",$B$4+E{r},E{r}))', fmt=USD2, border=True)
    put(ws, f"I{r}", f'=IF(F{r}="%",$B$5*(1+G{r}),IF(F{r}="$",$B$5+G{r},G{r}))', fmt=USD2, border=True)
    put(ws, f"J{r}", f"={QW}*(H{r}-$B$4)", fmt=USD, border=True)
    put(ws, f"K{r}", f"={QB}*(I{r}-$B$5)", fmt=USD, border=True)
    put(ws, f"L{r}", f"=J{r}+K{r}", f_bold, USD, border=True)
    put(ws, f"M{r}", f"=L{r}/$B$6", fmt='0.0"x";-0.0"x"', border=True)
    put(ws, f"N{r}", note, f_sub, border=True)
SRL = SR0 + len(scen) - 1
ws.conditional_formatting.add(f"L{SR0}:L{SRL}", CellIsRule(operator="lessThan", formula=["0"],
                              font=Font(name=F, bold=True, color="C00000")))
put(ws, f"A{SRL+2}", "Worst scenario loss ($)", f_bold); put(ws, f"L{SRL+2}", f"=MIN(L{SR0}:L{SRL})", f_bold, USD)
put(ws, f"A{SRL+3}", "Worst scenario", f_bold)
put(ws, f"L{SRL+3}", f"=INDEX(A{SR0}:A{SRL},MATCH(L{SRL+2},L{SR0}:L{SRL},0))", f_bold)

ch = BarChart(); ch.type = "bar"; ch.style = 10
ch.title = "Stress P&L on today's position ($)"; ch.y_axis.numFmt = '$#,##0'
ch.add_data(Reference(ws, min_col=12, min_row=9, max_row=SRL), titles_from_data=True)
ch.set_categories(Reference(ws, min_col=1, min_row=SR0, max_row=SRL))
ch.legend = None; ch.height = 9; ch.width = 22
ws.add_chart(ch, f"A{SRL+6}")

# ================================================================== Summary (first sheet)
su = wb.active; su.title = "Summary"
for i, w in enumerate([40, 17, 17, 17, 17, 4, 17], 1):
    su.column_dimensions[L(i)].width = w
su.sheet_view.showGridLines = False
title(su, "Crude Oil Market Risk Report",
      "Hypothetical long WTI / short Brent position. All figures recalculate from the Inputs sheet.")
put(su, "A4", "Position snapshot", f_bold, fill=fill_sec)
snap = [("As-of date", "=VaR!B8", DATE), ("WTI price ($/bbl)", "=VaR!B11", USD2),
        ("Brent price ($/bbl)", "=VaR!B12", USD2), ("WTI position (bbl)", f"={QW}", NUM),
        ("Brent position (bbl)", f"={QB}", NUM), ("MTM value ($)", "=VaR!B15", USD),
        ("Last-day P&L ($)", "=INDEX(Prices_MTM!I:I,VaR!B5)", USD),
        ("P&L, trailing 250 days ($)", "=INDEX(Prices_MTM!J:J,VaR!B5)-INDEX(Prices_MTM!J:J,VaR!B5-250)", USD),
        ("VaR window", '=TEXT(VaR!B7,"yyyy-mm-dd")&" to "&TEXT(VaR!B8,"yyyy-mm-dd")', None)]
for i, (lab, f, fmt) in enumerate(snap):
    put(su, f"A{5+i}", lab, border=True); put(su, f"B{5+i}", f, f_link, fmt, border=True, align="right")

put(su, "A15", "1-day VaR and Expected Shortfall ($ loss)", f_bold, fill=fill_sec)
header_row(su, 16, ["Method", "VaR 95%", "ES 95%", "VaR 99%", "ES 99%"])
for i in range(4):
    r = 17 + i
    put(su, f"A{r}", f"=VaR!A{29+i}", f_link, border=True)
    for c in "BCDE":
        put(su, f"{c}{r}", f"=VaR!{c}{29+i}", f_link, USD, border=True)
put(su, "B16", '="VaR "&TEXT(Inputs!B8,"0%")', f_head, fill=fill_head, align="center")
put(su, "C16", '="ES "&TEXT(Inputs!B8,"0%")', f_head, fill=fill_head, align="center")
put(su, "D16", '="VaR "&TEXT(Inputs!B9,"0%")', f_head, fill=fill_head, align="center")
put(su, "E16", '="ES "&TEXT(Inputs!B9,"0%")', f_head, fill=fill_head, align="center")

put(su, "A23", "Backtest (rolling 250-day VaR vs actual P&L)", f_bold, fill=fill_sec)
header_row(su, 24, ["Test", "Hist CL1", "Param CL1", "Hist CL2", "Param CL2"])
for i, src in enumerate([6, 7, 8, 9, 11, 12, 13, 14]):
    r = 25 + i
    put(su, f"A{r}", f"=Backtest!A{src}", f_link, border=True)
    fmt = {6: NUM, 7: "#,##0.0", 8: NUM, 9: PCT2, 11: "0.0000", 13: NUM}.get(src)
    for c in "BCDE":
        put(su, f"{c}{r}", f"=Backtest!{c}{src}", f_link, fmt, border=True, align="center")
for c in "BCDE":
    for v, color in (("Pass", "C6EFCE"), ("Green", "C6EFCE"), ("Reject", "FFC7CE"), ("Red", "FFC7CE"), ("Yellow", "FFEB9C")):
        su.conditional_formatting.add(f"{c}30:{c}32", CellIsRule(operator="equal", formula=[f'"{v}"'],
                                      fill=PatternFill("solid", fgColor=color)))

put(su, "A35", "Stress scenarios", f_bold, fill=fill_sec)
header_row(su, 36, ["Scenario", "WTI stressed", "Brent stressed", "P&L ($)", "x 99% VaR"])
for i in range(len(scen)):
    r, s = 37 + i, SR0 + i
    put(su, f"A{r}", f"=Stress!A{s}", f_link, border=True)
    put(su, f"B{r}", f"=Stress!H{s}", f_link, USD2, border=True)
    put(su, f"C{r}", f"=Stress!I{s}", f_link, USD2, border=True)
    put(su, f"D{r}", f"=Stress!L{s}", f_link, USD, border=True)
    put(su, f"E{r}", f"=Stress!M{s}", f_link, '0.0"x";-0.0"x"', border=True)
su.conditional_formatting.add(f"D37:D{36+len(scen)}", CellIsRule(operator="lessThan", formula=["0"],
                              font=Font(name=F, color="C00000", bold=True)))

kr = 37 + len(scen) + 2
put(su, f"A{kr}", "Key findings", f_bold, fill=fill_sec)
findings = [
    '="Fat tails: 99% historical VaR is "&TEXT(VaR!D34,"0.00")&"x the parametric (normal) VaR on the same window."',
    '="99% VaR exceptions over "&TEXT(Backtest!D6,"#,##0")&" days: "&Backtest!D8&" historical vs "&TEXT(Backtest!D7,"0")&" expected (Kupiec: "&Backtest!D12&")."',
    '="Worst stress: "&Stress!L' + str(SRL + 3) + '&", a loss of "&TEXT(-Stress!L' + str(SRL + 2) + ',"$#,##0")&" ("&TEXT(-Stress!L' + str(SRL + 2) + '/VaR!D29,"0.0")&"x 99% VaR)."',
    '="Short Brent hedge: component VaR of Brent leg = "&TEXT(VaR!B40,"$#,##0;($#,##0)")&" vs WTI "&TEXT(VaR!B39,"$#,##0")&"."',
]
for i, f in enumerate(findings):
    put(su, f"A{kr+1+i}", f)

# backtest chart on Summary
lc = LineChart(); lc.title = "Daily P&L vs 99% VaR (rolling 250d)"; lc.height = 9; lc.width = 24
lc.y_axis.numFmt = '$#,##0'; lc.y_axis.title = None
lc.add_data(Reference(wb_, min_col=2, min_row=BT0 - 1, max_row=BTL), titles_from_data=True)
lc.add_data(Reference(wb_, min_col=5, min_row=BT0 - 1, max_row=BTL), titles_from_data=True)
lc.set_categories(Reference(wb_, min_col=1, min_row=BT0, max_row=BTL))
for s_ in lc.series:
    s_.smooth = False
    s_.graphicalProperties.line.width = 6000
lc.x_axis.number_format = "yyyy"; lc.x_axis.tickLblSkip = 500
su.add_chart(lc, "G4")

for wsx in wb.worksheets:
    wsx.sheet_properties.tabColor = {"Summary": "1F3A5F", "Inputs": "FFFF00"}.get(wsx.title)

OUTFILE.parent.mkdir(exist_ok=True)
wb.save(OUTFILE)
print("saved", OUTFILE, "| backtest rows", nbt, "| years", years[0], "-", years[-1])
