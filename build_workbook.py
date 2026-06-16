#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Gold Technical Analysis workbook generator.

Builds a fully self-updating .xlsx for gold (XAU) technical analysis.
The ONLY input sheet is "Data" (Date | Open | High | Low | Close).
Every other sheet recalculates automatically from Excel formulas.

Design
------
Hidden calculation engine:
  Calc_Daily    - mirrors Data, computes every indicator per bar, and tags
                  week-end / month-end bars for resampling.
  Calc_Weekly   - weekly OHLC resampled from Calc_Daily, then all indicators.
  Calc_Monthly  - monthly OHLC resampled from Calc_Daily, then all indicators.
  ChartData     - the last N bars pulled into a fixed block so charts always
                  show the most recent window and still auto-update.
  Metrics       - the single source of truth: latest value of every indicator
                  on every timeframe + classification scores + horizon scores.

Visible sheets (Read Me, Dashboard, Trend, Momentum, Volatility,
Support & Resistance, Matrix, Outlook) only read from Metrics.

No VBA. Workbook is written with fullCalcOnLoad so Excel recomputes on open.
"""

import math
import os
import random
import datetime as dt

import xlsxwriter
from xlsxwriter.utility import xl_col_to_name, xl_range

OUT = os.environ.get("GX_OUT", "Gold_Technical_Analysis.xlsx")

# ----------------------------------------------------------------------------
# Capacity of the engine (how many bars it can hold). Generous but light.
# Overridable via env vars to allow building a tiny workbook for testing.
# ----------------------------------------------------------------------------
DAILY_MAX = int(os.environ.get("GX_DAILY", 2000))     # ~8 years of trading days
WEEKLY_MAX = int(os.environ.get("GX_WEEKLY", 440))    # ~8.5 years of weeks
MONTHLY_MAX = int(os.environ.get("GX_MONTHLY", 120))  # 10 years of months
SAMPLE_N = int(os.environ.get("GX_SAMPLE", 340))      # rows of sample data
CHART_N = 180             # bars shown in the recent-window charts

DROW = DAILY_MAX + 1      # last worksheet row of daily data (header = row 1)
WROW = WEEKLY_MAX + 1
MROW = MONTHLY_MAX + 1
FR = 2                    # first data row on every engine sheet

# Lookbacks
RSI_N = 14
STOCH_N = 14
ATR_N = 14
ADX_N = 14
BB_N = 20
DON_N = 20
HV_N = 20
EMA_FAST, EMA_SLOW, EMA_SIG = 12, 26, 9
FIB_LB = 100              # swing lookback for Fibonacci / regression
REG_N = 100
SWING_K = 4              # fractal half-width for swing high/low detection

# ----------------------------------------------------------------------------
# Bloomberg / Refinitiv-style dark palette
# ----------------------------------------------------------------------------
BG        = "#0B0F14"   # page background (near black)
PANEL     = "#141A22"   # panel fill
PANEL2    = "#1B2531"   # alt panel fill
CARD      = "#10161D"
LINE      = "#2A3340"   # borders / gridlines
TXT       = "#E6EDF3"   # primary text
MUTED     = "#8A97A6"   # secondary text
GOLD      = "#E8B923"   # accent (gold)
GOLD_DK   = "#3a2f0c"
BULL      = "#2EBD85"   # green
BULL_DK   = "#10362b"
BEAR      = "#F6465D"   # red
BEAR_DK   = "#3a141c"
NEUTRAL   = "#9AA7B4"
NEUTRAL_DK= "#222b35"
SLBULL    = "#7FC8A9"
SLBEAR    = "#E58FA0"
HEADER_BG = "#0D1117"

# ============================================================================
# Engine column layout (identical prefix on every engine sheet)
# ============================================================================
CALC_FIELDS = [
    "date", "open", "high", "low", "close",            # A-E
    "idx", "ret",
    "ema20", "ema50", "ema100", "ema200",
    "gain", "loss", "rsi",
    "ema12", "ema26", "macd", "macd_sig", "macd_hist",
    "hh14", "ll14", "stoch_k", "stoch_d",
    "rsi_min", "rsi_max", "srsi_raw", "srsi_k", "srsi_d",
    "tr", "atr", "plus_dm", "minus_dm", "plus_di", "minus_di", "dx", "adx",
    "bb_mid", "bb_std", "bb_up", "bb_low", "bb_pctb", "bb_bw",
    "don_up", "don_low", "don_mid",
    "hv",
    "tenkan", "kijun", "senkouA", "senkouB",
    "fib_hh", "fib_ll",
    "reg_slope", "reg_int", "reg_mid", "reg_steyx",
    "swing_hi", "swing_lo",
]
# Daily-only resample helper columns (appended after the shared prefix)
DAILY_EXTRA = [
    "wkstart", "wk_is_start", "wk_is_end", "wk_open", "wk_high", "wk_low", "wk_endpos",
    "mkstart", "mk_is_start", "mk_is_end", "mk_open", "mk_high", "mk_low", "mk_endpos",
]

CL = {f: xl_col_to_name(i) for i, f in enumerate(CALC_FIELDS)}
BASE_LEN = len(CALC_FIELDS)
EX = {f: xl_col_to_name(BASE_LEN + i) for i, f in enumerate(DAILY_EXTRA)}
POS_COL = xl_col_to_name(BASE_LEN)   # weekly/monthly resample position helper

# short aliases for OHLC columns
O, H, L, C = CL["open"], CL["high"], CL["low"], CL["close"]
A = CL["date"]


def rng(letter, r, n):
    """Absolute range covering the last `n` rows ending at row r (clamped)."""
    start = max(FR, r - n + 1)
    return f"${letter}${start}:${letter}${r}"


def cnt(r):
    """COUNT of close values from first data row through r."""
    return f"COUNT($E$2:$E{r})"


# ----------------------------------------------------------------------------
# Per-row indicator formula generator (shared by all three timeframes)
# ----------------------------------------------------------------------------
def ind_formula(field, r, ann):
    p = r - 1
    first = (r == FR)

    if field == "idx":
        return f'=IF($E{r}="","",N({CL["idx"]}{p})+1)'
    if field == "ret":
        if first:
            return '=""'
        return f'=IF(OR($E{r}="",$E{p}=""),"",LN($E{r}/$E{p}))'

    if field in ("ema20", "ema50", "ema100", "ema200", "ema12", "ema26"):
        period = {"ema20": 20, "ema50": 50, "ema100": 100, "ema200": 200,
                  "ema12": EMA_FAST, "ema26": EMA_SLOW}[field]
        k = 2.0 / (period + 1)
        col = CL[field]
        if first:
            return f'=IF($E{FR}="","",$E{FR})'
        return (f'=IF($E{r}="","",IF($E{p}="",$E{r},'
                f'$E{r}*{k:.10f}+{col}{p}*{1 - k:.10f}))')

    if field == "gain":
        if first:
            return '=""'
        return f'=IF(OR($E{r}="",$E{p}=""),"",MAX($E{r}-$E{p},0))'
    if field == "loss":
        if first:
            return '=""'
        return f'=IF(OR($E{r}="",$E{p}=""),"",MAX($E{p}-$E{r},0))'
    if field == "rsi":
        g = rng(CL["gain"], r, RSI_N)
        l = rng(CL["loss"], r, RSI_N)
        return (f'=IF({cnt(r)}<{RSI_N + 1},"",'
                f'IF(AVERAGE({l})=0,100,100-100/(1+AVERAGE({g})/AVERAGE({l}))))')

    if field == "macd":
        return f'=IF(OR({CL["ema12"]}{r}="",{CL["ema26"]}{r}=""),"",{CL["ema12"]}{r}-{CL["ema26"]}{r})'
    if field == "macd_sig":
        col = CL["macd_sig"]
        k = 2.0 / (EMA_SIG + 1)
        if first:
            return f'=IF({CL["macd"]}{FR}="","",{CL["macd"]}{FR})'
        return (f'=IF({CL["macd"]}{r}="","",IF({col}{p}="",{CL["macd"]}{r},'
                f'{CL["macd"]}{r}*{k:.10f}+{col}{p}*{1 - k:.10f}))')
    if field == "macd_hist":
        return f'=IF(OR({CL["macd"]}{r}="",{CL["macd_sig"]}{r}=""),"",{CL["macd"]}{r}-{CL["macd_sig"]}{r})'

    if field == "hh14":
        return f'=IF({cnt(r)}<{STOCH_N},"",MAX({rng(H, r, STOCH_N)}))'
    if field == "ll14":
        return f'=IF({cnt(r)}<{STOCH_N},"",MIN({rng(L, r, STOCH_N)}))'
    if field == "stoch_k":
        hh, ll = CL["hh14"] + str(r), CL["ll14"] + str(r)
        return (f'=IF(OR({hh}="",{ll}="",{hh}={ll}),"",'
                f'100*($E{r}-{ll})/({hh}-{ll}))')
    if field == "stoch_d":
        k3 = rng(CL["stoch_k"], r, 3)
        return f'=IF(COUNT({k3})<3,"",AVERAGE({k3}))'

    if field == "rsi_min":
        rr = rng(CL["rsi"], r, RSI_N)
        return f'=IF(COUNT({rr})<{RSI_N},"",MIN({rr}))'
    if field == "rsi_max":
        rr = rng(CL["rsi"], r, RSI_N)
        return f'=IF(COUNT({rr})<{RSI_N},"",MAX({rr}))'
    if field == "srsi_raw":
        rsi, mn, mx = CL["rsi"] + str(r), CL["rsi_min"] + str(r), CL["rsi_max"] + str(r)
        return (f'=IF(OR({rsi}="",{mn}="",{mx}="",{mx}={mn}),"",'
                f'100*({rsi}-{mn})/({mx}-{mn}))')
    if field == "srsi_k":
        s3 = rng(CL["srsi_raw"], r, 3)
        return f'=IF(COUNT({s3})<3,"",AVERAGE({s3}))'
    if field == "srsi_d":
        s3 = rng(CL["srsi_k"], r, 3)
        return f'=IF(COUNT({s3})<3,"",AVERAGE({s3}))'

    if field == "tr":
        if first:
            return f'=IF($E{FR}="","",$C{FR}-$D{FR})'
        return (f'=IF($E{r}="","",MAX($C{r}-$D{r},'
                f'ABS($C{r}-$E{p}),ABS($D{r}-$E{p})))')
    if field == "atr":
        return f'=IF({cnt(r)}<{ATR_N + 1},"",AVERAGE({rng(CL["tr"], r, ATR_N)}))'
    if field == "plus_dm":
        if first:
            return f'=IF($E{FR}="","",0)'
        return (f'=IF(OR($E{r}="",$E{p}=""),"",'
                f'IF(AND(($C{r}-$C{p})>($D{p}-$D{r}),($C{r}-$C{p})>0),$C{r}-$C{p},0))')
    if field == "minus_dm":
        if first:
            return f'=IF($E{FR}="","",0)'
        return (f'=IF(OR($E{r}="",$E{p}=""),"",'
                f'IF(AND(($D{p}-$D{r})>($C{r}-$C{p}),($D{p}-$D{r})>0),$D{p}-$D{r},0))')
    if field == "plus_di":
        s_tr = f'SUM({rng(CL["tr"], r, ADX_N)})'
        s_dm = f'SUM({rng(CL["plus_dm"], r, ADX_N)})'
        return f'=IF({cnt(r)}<{ADX_N + 1},"",IF({s_tr}=0,0,100*{s_dm}/{s_tr}))'
    if field == "minus_di":
        s_tr = f'SUM({rng(CL["tr"], r, ADX_N)})'
        s_dm = f'SUM({rng(CL["minus_dm"], r, ADX_N)})'
        return f'=IF({cnt(r)}<{ADX_N + 1},"",IF({s_tr}=0,0,100*{s_dm}/{s_tr}))'
    if field == "dx":
        pd, md = CL["plus_di"] + str(r), CL["minus_di"] + str(r)
        return (f'=IF(OR({pd}="",{md}=""),"",'
                f'IF(({pd}+{md})=0,0,100*ABS({pd}-{md})/({pd}+{md})))')
    if field == "adx":
        d = rng(CL["dx"], r, ADX_N)
        return f'=IF(COUNT({d})<{ADX_N},"",AVERAGE({d}))'

    if field == "bb_mid":
        return f'=IF({cnt(r)}<{BB_N},"",AVERAGE({rng(C, r, BB_N)}))'
    if field == "bb_std":
        return f'=IF({cnt(r)}<{BB_N},"",STDEVP({rng(C, r, BB_N)}))'
    if field == "bb_up":
        m, s = CL["bb_mid"] + str(r), CL["bb_std"] + str(r)
        return f'=IF({m}="","",{m}+2*{s})'
    if field == "bb_low":
        m, s = CL["bb_mid"] + str(r), CL["bb_std"] + str(r)
        return f'=IF({m}="","",{m}-2*{s})'
    if field == "bb_pctb":
        u, lo = CL["bb_up"] + str(r), CL["bb_low"] + str(r)
        return f'=IF(OR({u}="",{lo}="",{u}={lo}),"",($E{r}-{lo})/({u}-{lo}))'
    if field == "bb_bw":
        u, lo, m = CL["bb_up"] + str(r), CL["bb_low"] + str(r), CL["bb_mid"] + str(r)
        return f'=IF(OR({u}="",{m}="",{m}=0),"",({u}-{lo})/{m})'

    if field == "don_up":
        return f'=IF({cnt(r)}<{DON_N},"",MAX({rng(H, r, DON_N)}))'
    if field == "don_low":
        return f'=IF({cnt(r)}<{DON_N},"",MIN({rng(L, r, DON_N)}))'
    if field == "don_mid":
        u, lo = CL["don_up"] + str(r), CL["don_low"] + str(r)
        return f'=IF(OR({u}="",{lo}=""),"",({u}+{lo})/2)'

    if field == "hv":
        rr = rng(CL["ret"], r, HV_N)
        return f'=IF(COUNT({rr})<{HV_N},"",STDEV({rr})*SQRT({ann}))'

    if field == "tenkan":
        return f'=IF({cnt(r)}<9,"",(MAX({rng(H, r, 9)})+MIN({rng(L, r, 9)}))/2)'
    if field == "kijun":
        return f'=IF({cnt(r)}<26,"",(MAX({rng(H, r, 26)})+MIN({rng(L, r, 26)}))/2)'
    if field == "senkouA":
        t, k = CL["tenkan"] + str(r), CL["kijun"] + str(r)
        return f'=IF(OR({t}="",{k}=""),"",({t}+{k})/2)'
    if field == "senkouB":
        return f'=IF({cnt(r)}<52,"",(MAX({rng(H, r, 52)})+MIN({rng(L, r, 52)}))/2)'

    if field == "fib_hh":
        return f'=IF({cnt(r)}<20,"",MAX({rng(H, r, FIB_LB)}))'
    if field == "fib_ll":
        return f'=IF({cnt(r)}<20,"",MIN({rng(L, r, FIB_LB)}))'

    if field == "reg_slope":
        y, x = rng(C, r, REG_N), rng(CL["idx"], r, REG_N)
        return f'=IF({cnt(r)}<20,"",SLOPE({y},{x}))'
    if field == "reg_int":
        y, x = rng(C, r, REG_N), rng(CL["idx"], r, REG_N)
        return f'=IF({cnt(r)}<20,"",INTERCEPT({y},{x}))'
    if field == "reg_mid":
        s, i = CL["reg_slope"] + str(r), CL["reg_int"] + str(r)
        return f'=IF(OR({s}="",{i}=""),"",{i}+{s}*{CL["idx"]}{r})'
    if field == "reg_steyx":
        y, x = rng(C, r, REG_N), rng(CL["idx"], r, REG_N)
        return f'=IF({cnt(r)}<20,"",STEYX({y},{x}))'

    if field == "swing_hi":
        # confirmed fractal high: needs SWING_K bars either side
        if r < FR + SWING_K or r > DROW - SWING_K:
            return '=""'
        win = f'${H}${r - SWING_K}:${H}${r + SWING_K}'
        fwd = f'COUNT($E${r + 1}:$E${r + SWING_K})'
        return (f'=IF($E{r}="","",IF(AND({fwd}>={SWING_K},'
                f'$C{r}=MAX({win})),$C{r},""))')
    if field == "swing_lo":
        if r < FR + SWING_K or r > DROW - SWING_K:
            return '=""'
        win = f'${L}${r - SWING_K}:${L}${r + SWING_K}'
        fwd = f'COUNT($E${r + 1}:$E${r + SWING_K})'
        return (f'=IF($E{r}="","",IF(AND({fwd}>={SWING_K},'
                f'$D{r}=MIN({win})),$D{r},""))')

    return '=""'


# ----------------------------------------------------------------------------
# Daily-only resample helper formulas
# ----------------------------------------------------------------------------
def daily_extra_formula(field, r):
    p = r - 1
    n = r + 1
    first = (r == FR)
    if field == "wkstart":
        return f'=IF($A{r}="","",$A{r}-WEEKDAY($A{r},2)+1)'
    if field == "wk_is_start":
        if first:
            return f'=IF($A{FR}="","",1)'
        ws = EX["wkstart"]
        return f'=IF($A{r}="","",IF($A{p}="",1,IF({ws}{r}<>{ws}{p},1,0)))'
    if field == "wk_is_end":
        ws = EX["wkstart"]
        return f'=IF($A{r}="","",IF($A{n}="",1,IF({ws}{n}<>{ws}{r},1,0)))'
    if field == "wk_open":
        if first:
            return f'=IF($A{FR}="","",$B{FR})'
        return f'=IF($A{r}="","",IF({EX["wk_is_start"]}{r}=1,$B{r},{EX["wk_open"]}{p}))'
    if field == "wk_high":
        if first:
            return f'=IF($A{FR}="","",$C{FR})'
        return f'=IF($A{r}="","",IF({EX["wk_is_start"]}{r}=1,$C{r},MAX($C{r},{EX["wk_high"]}{p})))'
    if field == "wk_low":
        if first:
            return f'=IF($A{FR}="","",$D{FR})'
        return f'=IF($A{r}="","",IF({EX["wk_is_start"]}{r}=1,$D{r},MIN($D{r},{EX["wk_low"]}{p})))'
    if field == "wk_endpos":
        return f'=IF($A{r}="","",IF({EX["wk_is_end"]}{r}=1,{CL["idx"]}{r},""))'

    if field == "mkstart":
        return f'=IF($A{r}="","",DATE(YEAR($A{r}),MONTH($A{r}),1))'
    if field == "mk_is_start":
        if first:
            return f'=IF($A{FR}="","",1)'
        ms = EX["mkstart"]
        return f'=IF($A{r}="","",IF($A{p}="",1,IF({ms}{r}<>{ms}{p},1,0)))'
    if field == "mk_is_end":
        ms = EX["mkstart"]
        return f'=IF($A{r}="","",IF($A{n}="",1,IF({ms}{n}<>{ms}{r},1,0)))'
    if field == "mk_open":
        if first:
            return f'=IF($A{FR}="","",$B{FR})'
        return f'=IF($A{r}="","",IF({EX["mk_is_start"]}{r}=1,$B{r},{EX["mk_open"]}{p}))'
    if field == "mk_high":
        if first:
            return f'=IF($A{FR}="","",$C{FR})'
        return f'=IF($A{r}="","",IF({EX["mk_is_start"]}{r}=1,$C{r},MAX($C{r},{EX["mk_high"]}{p})))'
    if field == "mk_low":
        if first:
            return f'=IF($A{FR}="","",$D{FR})'
        return f'=IF($A{r}="","",IF({EX["mk_is_start"]}{r}=1,$D{r},MIN($D{r},{EX["mk_low"]}{p})))'
    if field == "mk_endpos":
        return f'=IF($A{r}="","",IF({EX["mk_is_end"]}{r}=1,{CL["idx"]}{r},""))'
    return '=""'


# ============================================================================
# Build the workbook
# ============================================================================
wb = xlsxwriter.Workbook(OUT, {"nan_inf_to_errors": True})

# ---- formats ---------------------------------------------------------------
fmt_cache = {}


def F(**kw):
    key = tuple(sorted(kw.items()))
    if key not in fmt_cache:
        fmt_cache[key] = wb.add_format(kw)
    return fmt_cache[key]


page = F(bg_color=BG, font_color=TXT, font_name="Segoe UI")
num_plain = F(num_format="#,##0.00")
date_plain = F(num_format="yyyy-mm-dd")

print("formats ready")

# ============================================================================
# DATA (input)
# ============================================================================
ws_data = wb.add_worksheet("Data")
ws_data.hide_gridlines(2)
ws_data.set_tab_color(GOLD)
ws_data.set_column("A:CA", None, page)
ws_data.set_column("A:A", 14)
ws_data.set_column("B:E", 12)
ws_data.set_column("G:G", 64)

title_data = F(bg_color=BG, font_color=GOLD, bold=True, font_size=14, font_name="Segoe UI")
note = F(bg_color=BG, font_color=MUTED, italic=True, text_wrap=True, valign="top", font_name="Segoe UI")

# --- sample data (clearly labelled; user replaces with their own) -----------
random.seed(7)
sample = []
d = dt.date(2026, 6, 12)
# walk backwards to collect business days, then reverse
days = []
while len(days) < SAMPLE_N:
    if d.weekday() < 5:
        days.append(d)
    d -= dt.timedelta(days=1)
days.reverse()
price = 1850.0
for day in days:
    drift = 0.0004
    shock = random.gauss(0, 1) * 0.0095
    new_close = price * math.exp(drift + shock)
    op = price * math.exp(random.gauss(0, 1) * 0.003)
    hi = max(op, new_close) * (1 + abs(random.gauss(0, 1)) * 0.004)
    lo = min(op, new_close) * (1 - abs(random.gauss(0, 1)) * 0.004)
    sample.append((day, round(op, 2), round(hi, 2), round(lo, 2), round(new_close, 2)))
    price = new_close

n_sample = len(sample)
headers = ["Date", "Open", "High", "Low", "Close"]
for ci, htxt in enumerate(headers):
    ws_data.write(0, ci, htxt)
for ri, (day, op, hi, lo, cl) in enumerate(sample, start=1):
    ws_data.write_datetime(ri, 0, dt.datetime(day.year, day.month, day.day), date_plain)
    ws_data.write_number(ri, 1, op, num_plain)
    ws_data.write_number(ri, 2, hi, num_plain)
    ws_data.write_number(ri, 3, lo, num_plain)
    ws_data.write_number(ri, 4, cl, num_plain)

ws_data.add_table(0, 0, n_sample, 4, {
    "name": "tblData",
    "style": "Table Style Dark 9",
    "columns": [{"header": h} for h in headers],
})
ws_data.write("G1", "INPUT SHEET — paste your gold OHLC history here.", title_data)
ws_data.write("G2",
    "Columns must stay in this order: Date | Open | High | Low | Close. "
    "Dates oldest→newest (top to bottom). To use your own data: select the "
    "sample rows below the header and paste over them; the table and every "
    "other sheet update automatically. The engine supports up to "
    f"{DAILY_MAX:,} daily rows. The data shipped here is SAMPLE data — "
    "replace it. See the 'Read Me' sheet for details.", note)
ws_data.set_row(0, 18)
ws_data.freeze_panes(1, 0)
print("Data sheet done")

# ============================================================================
# ENGINE SHEETS
# ============================================================================
def new_engine_sheet(name):
    ws = wb.add_worksheet(name)
    ws.hide()
    return ws


def write_indicators(ws, last_row, ann):
    """Write the shared indicator block (idx .. swing_lo) for rows FR..last_row."""
    for f in CALC_FIELDS[5:]:           # skip date/open/high/low/close
        c = CL[f]
        ci = ord("A")  # not used; we use xl by name
        col_idx = CALC_FIELDS.index(f)
        for r in range(FR, last_row + 1):
            ws.write_formula(r - 1, col_idx, ind_formula(f, r, ann))


cd = new_engine_sheet("Calc_Daily")
cw = new_engine_sheet("Calc_Weekly")
cm = new_engine_sheet("Calc_Monthly")

# --- Calc_Daily base = mirror of Data --------------------------------------
for r in range(FR, DROW + 1):
    cd.write_formula(r - 1, 0, f'=IF(Data!$A{r}="","",Data!$A{r})', date_plain)
    cd.write_formula(r - 1, 1, f'=IF(Data!$B{r}="","",Data!$B{r})')
    cd.write_formula(r - 1, 2, f'=IF(Data!$C{r}="","",Data!$C{r})')
    cd.write_formula(r - 1, 3, f'=IF(Data!$D{r}="","",Data!$D{r})')
    cd.write_formula(r - 1, 4, f'=IF(Data!$E{r}="","",Data!$E{r})')
write_indicators(cd, DROW, 252)
# daily resample helpers
for f in DAILY_EXTRA:
    col_idx = BASE_LEN + DAILY_EXTRA.index(f)
    for r in range(FR, DROW + 1):
        cd.write_formula(r - 1, col_idx, daily_extra_formula(f, r))
print("Calc_Daily done")

# --- Calc_Weekly / Calc_Monthly base = resampled from Calc_Daily ------------
def write_resampled_base(ws, last_row, endpos_col, open_col, high_col, low_col):
    pos_idx = BASE_LEN  # POS_COL
    we = f"Calc_Daily!${endpos_col}$2:${endpos_col}${DROW}"
    d_date = f"Calc_Daily!$A$2:$A${DROW}"
    d_close = f"Calc_Daily!$E$2:$E${DROW}"
    d_open = f"Calc_Daily!${open_col}$2:${open_col}${DROW}"
    d_high = f"Calc_Daily!${high_col}$2:${high_col}${DROW}"
    d_low = f"Calc_Daily!${low_col}$2:${low_col}${DROW}"
    for r in range(FR, last_row + 1):
        k = r - 1  # 1-based bucket number
        ws.write_formula(r - 1, pos_idx, f'=IFERROR(SMALL({we},{k}),"")')
        pos = f"${POS_COL}{r}"
        ws.write_formula(r - 1, 0, f'=IF({pos}="","",INDEX({d_date},{pos}))', date_plain)
        ws.write_formula(r - 1, 1, f'=IF({pos}="","",INDEX({d_open},{pos}))')
        ws.write_formula(r - 1, 2, f'=IF({pos}="","",INDEX({d_high},{pos}))')
        ws.write_formula(r - 1, 3, f'=IF({pos}="","",INDEX({d_low},{pos}))')
        ws.write_formula(r - 1, 4, f'=IF({pos}="","",INDEX({d_close},{pos}))')


write_resampled_base(cw, WROW, EX["wk_endpos"], EX["wk_open"], EX["wk_high"], EX["wk_low"])
write_indicators(cw, WROW, 52)
print("Calc_Weekly done")

write_resampled_base(cm, MROW, EX["mk_endpos"], EX["mk_open"], EX["mk_high"], EX["mk_low"])
write_indicators(cm, MROW, 12)
print("Calc_Monthly done")

# ============================================================================
# METRICS  (single source of truth for every visible sheet)
# ============================================================================
ms = wb.add_worksheet("Metrics")
ms.hide()
ms.write_formula("H1", f'=COUNT(Calc_Daily!$E$2:$E${DROW})')
ms.write_formula("H2", f'=COUNT(Calc_Weekly!$E$2:$E${WROW})')
ms.write_formula("H3", f'=COUNT(Calc_Monthly!$E$2:$E${MROW})')

TFD = {"D": ("Calc_Daily", DROW, "$H$1"),
       "W": ("Calc_Weekly", WROW, "$H$2"),
       "M": ("Calc_Monthly", MROW, "$H$3")}
TFCOL = {"D": "C", "W": "D", "M": "E"}
TFIDX = {"C": 2, "D": 3, "E": 4}

mrow = {}
row = 4

# --- latest value of each field per timeframe ------------------------------
latest_specs = [
    ("close", "E", 0), ("open", O, 0), ("high", H, 0), ("low", L, 0),
    ("close1", "E", 1), ("close5", "E", 5), ("close21", "E", 21),
    ("ema20", CL["ema20"], 0), ("ema50", CL["ema50"], 0),
    ("ema100", CL["ema100"], 0), ("ema200", CL["ema200"], 0),
    ("ema20_5", CL["ema20"], 5), ("ema50_5", CL["ema50"], 5),
    ("ema100_5", CL["ema100"], 5), ("ema200_5", CL["ema200"], 5),
    ("rsi", CL["rsi"], 0),
    ("macd", CL["macd"], 0), ("macd_sig", CL["macd_sig"], 0), ("macd_hist", CL["macd_hist"], 0),
    ("stoch_k", CL["stoch_k"], 0), ("stoch_d", CL["stoch_d"], 0),
    ("srsi_k", CL["srsi_k"], 0), ("srsi_d", CL["srsi_d"], 0),
    ("adx", CL["adx"], 0), ("plus_di", CL["plus_di"], 0), ("minus_di", CL["minus_di"], 0),
    ("atr", CL["atr"], 0),
    ("bb_up", CL["bb_up"], 0), ("bb_mid", CL["bb_mid"], 0), ("bb_low", CL["bb_low"], 0),
    ("bb_pctb", CL["bb_pctb"], 0), ("bb_bw", CL["bb_bw"], 0),
    ("don_up", CL["don_up"], 0), ("don_low", CL["don_low"], 0), ("don_mid", CL["don_mid"], 0),
    ("hv", CL["hv"], 0),
    ("tenkan", CL["tenkan"], 0), ("kijun", CL["kijun"], 0),
    ("senkouA_now", CL["senkouA"], 26), ("senkouB_now", CL["senkouB"], 26),
    ("fib_hh", CL["fib_hh"], 0), ("fib_ll", CL["fib_ll"], 0),
    ("reg_slope", CL["reg_slope"], 0), ("reg_mid", CL["reg_mid"], 0), ("reg_steyx", CL["reg_steyx"], 0),
]
for name, cc, off in latest_specs:
    mrow[name] = row
    ms.write(row - 1, 1, name)
    for tf in "DWM":
        sheet, maxr, ncell = TFD[tf]
        rngs = f"{sheet}!${cc}$2:${cc}${maxr}"
        pos = f"({ncell}-{off})" if off else ncell
        ms.write_formula(row - 1, TFIDX[TFCOL[tf]], f'=IFERROR(INDEX({rngs},{pos}),"")')
    row += 1

# hv 100-bar average (volatility-regime baseline)
mrow["hv_avg"] = row
ms.write(row - 1, 1, "hv_avg")
for tf in "DWM":
    sheet, maxr, ncell = TFD[tf]
    rngs = f"{sheet}!${CL['hv']}$2:${CL['hv']}${maxr}"
    ms.write_formula(row - 1, TFIDX[TFCOL[tf]],
                     f'=IFERROR(AVERAGE(INDEX({rngs},MAX(1,{ncell}-99)):INDEX({rngs},{ncell})),"")')
row += 1


def addrow(name, formula_C):
    """Daily-only single-cell metric in column C."""
    global row
    mrow[name] = row
    ms.write(row - 1, 1, name)
    ms.write_formula(row - 1, 2, formula_C)
    row += 1


def v(name, tf="D"):
    return f"${TFCOL[tf]}${mrow[name]}"


def MV(name, tf):
    return f"Metrics!${TFCOL[tf]}${mrow[name]}"


def MVd(name):
    return f"Metrics!$C${mrow[name]}"


# --- daily-only context cells ----------------------------------------------
dateR = f"Calc_Daily!$A$2:$A${DROW}"
closeR = f"Calc_Daily!$E$2:$E${DROW}"
addrow("lastdate", f'=IFERROR(INDEX({dateR},$H$1),"")')
addrow("ytdbase", f'=IFERROR(LOOKUP(DATE(YEAR(INDEX({dateR},$H$1)),1,1)-1,{dateR},{closeR}),"")')

# returns
addrow("ret_d", f'=IFERROR({v("close","D")}/{v("close1","D")}-1,"")')
addrow("ret_w", f'=IFERROR({v("close","D")}/{v("close5","D")}-1,"")')
addrow("ret_m", f'=IFERROR({v("close","D")}/{v("close21","D")}-1,"")')
addrow("ret_ytd", f'=IFERROR({v("close","D")}/{v("ytdbase")}-1,"")')

# classic pivot levels (from latest daily bar)
addrow("pp", f'=IFERROR(({v("high","D")}+{v("low","D")}+{v("close","D")})/3,"")')
addrow("piv_r1", f'=IFERROR(2*{v("pp")}-{v("low","D")},"")')
addrow("piv_s1", f'=IFERROR(2*{v("pp")}-{v("high","D")},"")')
addrow("piv_r2", f'=IFERROR({v("pp")}+({v("high","D")}-{v("low","D")}),"")')
addrow("piv_s2", f'=IFERROR({v("pp")}-({v("high","D")}-{v("low","D")}),"")')
addrow("piv_r3", f'=IFERROR({v("high","D")}+2*({v("pp")}-{v("low","D")}),"")')
addrow("piv_s3", f'=IFERROR({v("low","D")}-2*({v("high","D")}-{v("pp")}),"")')

# swing-based support / resistance (AGGREGATE = k-th nearest, ignoring errors)
swhi = f"Calc_Daily!${CL['swing_hi']}$2:${CL['swing_hi']}${DROW}"
swlo = f"Calc_Daily!${CL['swing_lo']}$2:${CL['swing_lo']}${DROW}"
prc = v("close", "D")
addrow("res1", f'=IFERROR(_xlfn.AGGREGATE(15,6,{swhi}/({swhi}>{prc}),1),"")')
addrow("res2", f'=IFERROR(_xlfn.AGGREGATE(15,6,{swhi}/(({swhi}>{prc})*({swhi}>{v("res1")})),1),"")')
addrow("res3", f'=IFERROR(_xlfn.AGGREGATE(15,6,{swhi}/(({swhi}>{prc})*({swhi}>{v("res2")})),1),"")')
addrow("sup1", f'=IFERROR(_xlfn.AGGREGATE(14,6,{swlo}/({swlo}<{prc}),1),"")')
addrow("sup2", f'=IFERROR(_xlfn.AGGREGATE(14,6,{swlo}/(({swlo}<{prc})*({swlo}<{v("sup1")})),1),"")')
addrow("sup3", f'=IFERROR(_xlfn.AGGREGATE(14,6,{swlo}/(({swlo}<{prc})*({swlo}<{v("sup2")})),1),"")')

# Fibonacci retracement of the most recent major swing (lookback FIB_LB)
highR = f"Calc_Daily!${H}$2:${H}${DROW}"
lowR = f"Calc_Daily!${L}$2:${L}${DROW}"
hwin = f"INDEX({highR},MAX(1,$H$1-{FIB_LB - 1})):INDEX({highR},$H$1)"
lwin = f"INDEX({lowR},MAX(1,$H$1-{FIB_LB - 1})):INDEX({lowR},$H$1)"
addrow("fib_dir", f'=IFERROR(IF(MATCH({v("fib_hh","D")},{hwin},0)>=MATCH({v("fib_ll","D")},{lwin},0),1,-1),1)')
FIB_RATIOS = [0.0, 0.236, 0.382, 0.5, 0.618, 0.786, 1.0]
for rt in FIB_RATIOS:
    key = f"fib_{int(rt * 1000):04d}"
    addrow(key, f'=IFERROR(IF({v("fib_dir")}=1,'
                f'{v("fib_hh","D")}-({v("fib_hh","D")}-{v("fib_ll","D")})*{rt},'
                f'{v("fib_ll","D")}+({v("fib_hh","D")}-{v("fib_ll","D")})*{rt}),"")')

# --- status texts (daily) ---------------------------------------------------
addrow("trend_status",
       f'=IF({v("close","D")}="","-",'
       f'IF(AND({v("close","D")}>{v("ema50","D")},{v("ema50","D")}>{v("ema200","D")}),'
       f'IF({v("adx","D")}>=25,"STRONG UPTREND","UPTREND"),'
       f'IF(AND({v("close","D")}<{v("ema50","D")},{v("ema50","D")}<{v("ema200","D")}),'
       f'IF({v("adx","D")}>=25,"STRONG DOWNTREND","DOWNTREND"),"SIDEWAYS")))')
addrow("mom_status",
       f'=IF({v("rsi","D")}="","-",'
       f'IF(AND({v("rsi","D")}>=55,{v("macd_hist","D")}>0),"BULLISH",'
       f'IF(AND({v("rsi","D")}<=45,{v("macd_hist","D")}<0),"BEARISH","NEUTRAL")))')
addrow("vol_status",
       f'=IF(OR({v("hv","D")}="",{v("hv_avg","D")}="",{v("hv_avg","D")}=0),"-",'
       f'IF({v("hv","D")}/{v("hv_avg","D")}>=1.2,"HIGH",'
       f'IF({v("hv","D")}/{v("hv_avg","D")}<=0.8,"LOW","NORMAL")))')

# ============================================================================
# Classification scores (-2..+2) per indicator per timeframe
# ============================================================================
IND = ["RSI", "MACD", "Stochastic", "Stoch RSI", "ADX/DMI", "EMA20", "EMA50",
       "EMA100", "EMA200", "Ichimoku", "Bollinger Bands", "Donchian Channel",
       "Fibonacci", "Regression Channel"]
NIND = len(IND)


def score_formula(ind, tf):
    def x(f):
        return v(f, tf)
    if ind == "RSI":
        r = x("rsi")
        return f'=IF({r}="",0,IF({r}>=65,2,IF({r}>=55,1,IF({r}>=45,0,IF({r}>=35,-1,-2)))))'
    if ind == "MACD":
        return f'=IF({x("macd")}="",0,MEDIAN(-2,SIGN({x("macd")}-{x("macd_sig")})+SIGN({x("macd")}),2))'
    if ind == "Stochastic":
        return f'=IF({x("stoch_k")}="",0,MEDIAN(-2,IF({x("stoch_k")}>{x("stoch_d")},1,-1)+IF({x("stoch_k")}>50,1,-1),2))'
    if ind == "Stoch RSI":
        return f'=IF({x("srsi_k")}="",0,MEDIAN(-2,IF({x("srsi_k")}>{x("srsi_d")},1,-1)+IF({x("srsi_k")}>50,1,-1),2))'
    if ind == "ADX/DMI":
        return f'=IF({x("adx")}="",0,IF({x("adx")}<20,0,SIGN({x("plus_di")}-{x("minus_di")})*IF({x("adx")}>=25,2,1)))'
    if ind in ("EMA20", "EMA50", "EMA100", "EMA200"):
        e = x(ind.lower()); e5 = x(ind.lower() + "_5"); pr = x("close")
        return f'=IF(OR({e}="",{pr}=""),0,IF({pr}>{e},1,-1)+IF({e}>{e5},1,-1))'
    if ind == "Ichimoku":
        pr, t, k = x("close"), x("tenkan"), x("kijun")
        sa, sb = x("senkouA_now"), x("senkouB_now")
        cloud = f'IF(OR({sa}="",{sb}=""),0,IF({pr}>MAX({sa},{sb}),1,IF({pr}<MIN({sa},{sb}),-1,0)))'
        return f'=IF(OR({pr}="",{t}="",{k}=""),0,MEDIAN(-2,{cloud}+SIGN({t}-{k}),2))'
    if ind == "Bollinger Bands":
        b = x("bb_pctb")
        return f'=IF({b}="",0,IF({b}>=0.8,2,IF({b}>=0.55,1,IF({b}>0.45,0,IF({b}>0.2,-1,-2)))))'
    if ind == "Donchian Channel":
        du, dl, pr = x("don_up"), x("don_low"), x("close")
        pos = f'(({pr}-{dl})/({du}-{dl}))'
        return f'=IF(OR({du}="",{dl}="",{du}={dl}),0,IF({pos}>=0.8,2,IF({pos}>=0.55,1,IF({pos}>0.45,0,IF({pos}>0.2,-1,-2)))))'
    if ind == "Fibonacci":
        hh, ll, pr = x("fib_hh"), x("fib_ll"), x("close")
        pos = f'(({pr}-{ll})/({hh}-{ll}))'
        return f'=IF(OR({hh}="",{ll}="",{hh}={ll}),0,IF({pos}>=0.8,2,IF({pos}>=0.618,1,IF({pos}>0.382,0,IF({pos}>0.2,-1,-2)))))'
    if ind == "Regression Channel":
        sl, mid, pr = x("reg_slope"), x("reg_mid"), x("close")
        return f'=IF(OR({sl}="",{mid}="",{pr}=""),0,MEDIAN(-2,SIGN({sl})+IF({pr}>{mid},1,-1),2))'
    return "=0"


row += 1  # gap
srow = {}
score_first = row
for ind in IND:
    srow[ind] = row
    ms.write(row - 1, 1, ind)
    for tf in "DWM":
        ms.write_formula(row - 1, TFIDX[TFCOL[tf]], score_formula(ind, tf))
    row += 1
score_last = row - 1


def sc(ind, tf):
    return f"Metrics!${TFCOL[tf]}${srow[ind]}"


# --- horizon aggregation ----------------------------------------------------
row += 1
hrow = {}


def haddrow(name, fmt_by_tf):
    global row
    hrow[name] = row
    ms.write(row - 1, 1, name)
    for tf in "DWM":
        ms.write_formula(row - 1, TFIDX[TFCOL[tf]], fmt_by_tf(tf))
    row += 1


def srange(tf):
    return f"{TFCOL[tf]}{score_first}:{TFCOL[tf]}{score_last}"


def hv_(name, tf):
    return f"${TFCOL[tf]}${hrow[name]}"


haddrow("score_avg", lambda tf: f"=AVERAGE({srange(tf)})")
haddrow("score100", lambda tf: f"=ROUND(({hv_('score_avg', tf)}+2)/4*100,0)")
haddrow("bias", lambda tf: f'=IF({hv_("score100", tf)}>=60,"BULLISH",IF({hv_("score100", tf)}<=40,"BEARISH","NEUTRAL"))')
haddrow("n_bull", lambda tf: f'=COUNTIF({srange(tf)},">0")')
haddrow("n_neu", lambda tf: f'=COUNTIF({srange(tf)},0)')
haddrow("n_bear", lambda tf: f'=COUNTIF({srange(tf)},"<0")')
haddrow("pct_bull", lambda tf: f"={hv_('n_bull', tf)}/{NIND}")
haddrow("pct_neu", lambda tf: f"={hv_('n_neu', tf)}/{NIND}")
haddrow("pct_bear", lambda tf: f"={hv_('n_bear', tf)}/{NIND}")


def MH(name, tf):
    return f"Metrics!${TFCOL[tf]}${hrow[name]}"


# --- outlook (1W->Daily, 1M->Weekly, 3M->Monthly) --------------------------
row += 1
HZ = [("1W", "D", 5), ("1M", "W", 21), ("3M", "M", 63)]
HZCOL = {"1W": "C", "1M": "D", "3M": "E"}
orow = {}


def oaddrow(name, fmt):
    global row
    orow[name] = row
    ms.write(row - 1, 1, name)
    for hz, tf, days in HZ:
        ms.write_formula(row - 1, TFIDX[HZCOL[hz]], fmt(hz, tf, days))
    row += 1


def ov(name, hz):
    return f"${HZCOL[hz]}${orow[name]}"


hvref = v("hv", "D")          # daily annualised HV
prref = v("close", "D")
oaddrow("expvol", lambda hz, tf, d: f'=IFERROR({hvref}*SQRT({d}/252),"")')
oaddrow("biasstr", lambda hz, tf, d: f'=({hv_("score100", tf)}-50)/50')
oaddrow("target", lambda hz, tf, d: f'=IFERROR({prref}*(1+{ov("biasstr", hz)}*{ov("expvol", hz)}),"")')
oaddrow("low", lambda hz, tf, d: f'=IFERROR({prref}*(1-{ov("expvol", hz)}),"")')
oaddrow("high", lambda hz, tf, d: f'=IFERROR({prref}*(1+{ov("expvol", hz)}),"")')
oaddrow("conf", lambda hz, tf, d: f'=MIN(95,ROUND(ABS({hv_("score100", tf)}-50)*1.9,0))')


def MO(name, hz):
    return f"Metrics!${HZCOL[hz]}${orow[name]}"


def MObias(hz):
    # bias text of the timeframe mapped to this horizon
    tf = dict((h, t) for h, t, _ in HZ)[hz]
    return MH("bias", tf)


print("Metrics done")

# ============================================================================
# Defined names (clean references for the visible sheets)
# ============================================================================
wb.define_name("CurrentPrice", f"=Metrics!$C${mrow['close']}")
wb.define_name("AsOfDate", f"=Metrics!$C${mrow['lastdate']}")
wb.define_name("RetDay", f"=Metrics!$C${mrow['ret_d']}")
wb.define_name("RetWeek", f"=Metrics!$C${mrow['ret_w']}")
wb.define_name("RetMonth", f"=Metrics!$C${mrow['ret_m']}")
wb.define_name("RetYTD", f"=Metrics!$C${mrow['ret_ytd']}")
wb.define_name("TrendStatus", f"=Metrics!$C${mrow['trend_status']}")
wb.define_name("MomentumStatus", f"=Metrics!$C${mrow['mom_status']}")
wb.define_name("VolatilityStatus", f"=Metrics!$C${mrow['vol_status']}")

# ============================================================================
# Presentation formats
# ============================================================================
PRICE, PRICE0, PCT2, PCT0, NUM2, INT = "#,##0.00", "#,##0", "0.00%", "0%", "0.00", "0"

fmt_title = F(bg_color=BG, font_color=GOLD, bold=True, font_size=20, font_name="Segoe UI")
fmt_sub = F(bg_color=BG, font_color=MUTED, font_size=10, italic=True, font_name="Segoe UI")
fmt_section = F(bg_color=PANEL2, font_color=GOLD, bold=True, font_size=11, align="left",
                valign="vcenter", border=1, border_color=LINE, indent=1, font_name="Segoe UI")
fmt_klabel = F(bg_color=PANEL, font_color=MUTED, font_size=9, align="left", valign="vcenter",
               indent=1, border=1, border_color=LINE)
fmt_kval = F(bg_color=PANEL, font_color=TXT, bold=True, font_size=16, align="left",
             valign="vcenter", indent=1, border=1, border_color=LINE)
fmt_kprice = F(bg_color=PANEL, font_color=GOLD, bold=True, font_size=22, align="left",
               valign="vcenter", indent=1, border=1, border_color=LINE, num_format=PRICE)
fmt_kpct = F(bg_color=PANEL, font_color=TXT, bold=True, font_size=16, align="left",
             valign="vcenter", indent=1, border=1, border_color=LINE, num_format=PCT2)
fmt_chip = F(bg_color=PANEL, font_color=TXT, bold=True, font_size=14, align="center",
             valign="vcenter", border=1, border_color=LINE)
fmt_th = F(bg_color=PANEL2, font_color=GOLD, bold=True, font_size=10, align="center",
           valign="vcenter", border=1, border_color=LINE)
fmt_thl = F(bg_color=PANEL2, font_color=GOLD, bold=True, font_size=10, align="left",
            valign="vcenter", border=1, border_color=LINE, indent=1)
fmt_td = F(bg_color=PANEL, font_color=TXT, font_size=10, align="center", valign="vcenter",
           border=1, border_color=LINE)
fmt_tdl = F(bg_color=PANEL, font_color=TXT, font_size=10, align="left", valign="vcenter",
            border=1, border_color=LINE, indent=1)
fmt_tdl_m = F(bg_color=PANEL, font_color=MUTED, font_size=10, align="left", valign="vcenter",
              border=1, border_color=LINE, indent=1)
fmt_tdp = F(bg_color=PANEL, font_color=TXT, font_size=10, align="center", valign="vcenter",
            border=1, border_color=LINE, num_format=PRICE)
fmt_tdpct = F(bg_color=PANEL, font_color=TXT, font_size=10, align="center", valign="vcenter",
              border=1, border_color=LINE, num_format=PCT2)
fmt_tdnum = F(bg_color=PANEL, font_color=TXT, font_size=10, align="center", valign="vcenter",
              border=1, border_color=LINE, num_format=NUM2)
fmt_matrix = F(bg_color=PANEL, font_color=TXT, bold=True, font_size=9, align="center",
               valign="vcenter", border=1, border_color=LINE)
fmt_nav = F(bg_color=GOLD_DK, font_color=GOLD, bold=True, font_size=9, align="center",
            valign="vcenter", border=1, border_color="#5a4a12")
fmt_note = F(bg_color=BG, font_color=MUTED, font_size=9, italic=True, text_wrap=True, valign="top")
fmt_body = F(bg_color=BG, font_color=TXT, font_size=10, text_wrap=True, valign="top")
fmt_bodyb = F(bg_color=BG, font_color=GOLD, font_size=11, bold=True, valign="top")

# conditional-format target formats
cf_bull = F(bg_color=BULL_DK, font_color="#7EE2B8", bold=True, border=1, border_color=LINE, align="center", valign="vcenter")
cf_slbull = F(bg_color="#13241d", font_color=SLBULL, border=1, border_color=LINE, align="center", valign="vcenter")
cf_neu = F(bg_color=NEUTRAL_DK, font_color=NEUTRAL, border=1, border_color=LINE, align="center", valign="vcenter")
cf_slbear = F(bg_color="#2a1820", font_color=SLBEAR, border=1, border_color=LINE, align="center", valign="vcenter")
cf_bear = F(bg_color=BEAR_DK, font_color="#FF93A4", bold=True, border=1, border_color=LINE, align="center", valign="vcenter")
cf_warn = F(bg_color=GOLD_DK, font_color=GOLD, bold=True, border=1, border_color=LINE, align="center", valign="vcenter")
cf_greenf = F(font_color="#3DDc97", bold=True)
cf_redf = F(font_color="#FF6B7E", bold=True)

NAV = [("Dashboard", "Dashboard"), ("Trend", "Trend"), ("Momentum", "Momentum"),
       ("Volatility", "Volatility"), ("S & R", "Support & Resistance"),
       ("Matrix", "Matrix"), ("Outlook", "Outlook"), ("Data", "Data"),
       ("Read Me", "Read Me")]


def xlref(r1, c1, r2, c2):
    return xl_range(r1, c1, r2, c2)


def mform(ws, r1, c1, r2, c2, formula, fmt):
    if (r1, c1) == (r2, c2):
        ws.write_formula(r1, c1, formula, fmt)
    else:
        ws.merge_range(r1, c1, r2, c2, "", fmt)
        ws.write_formula(r1, c1, formula, fmt)


def mtext(ws, r1, c1, r2, c2, text, fmt):
    if (r1, c1) == (r2, c2):
        ws.write(r1, c1, text, fmt)
    else:
        ws.merge_range(r1, c1, r2, c2, text, fmt)


def new_sheet(name, tab=GOLD, last_col=11, colw=13):
    ws = wb.add_worksheet(name)
    ws.hide_gridlines(2)
    ws.set_tab_color(tab)
    ws.set_column("A:BZ", None, page)
    ws.set_column(0, 0, 2)
    ws.set_column(1, last_col, colw)
    return ws


def draw_nav(ws, row):
    for i, (label, sheet) in enumerate(NAV):
        ws.write_url(row, 1 + i, f"internal:'{sheet}'!A1", fmt_nav, label)
    ws.set_row(row, 18)


def section(ws, row, c1, c2, text):
    mtext(ws, row, c1, row, c2, text, fmt_section)
    ws.set_row(row, 22)


def style_chart(chart, title, legend=True):
    chart.set_title({"name": title, "name_font": {"color": TXT, "size": 11, "bold": True, "name": "Segoe UI"}})
    chart.set_chartarea({"fill": {"color": CARD}, "border": {"color": LINE}})
    chart.set_plotarea({"fill": {"color": PANEL}, "border": {"none": True}})
    if legend:
        chart.set_legend({"font": {"color": MUTED, "size": 8}, "position": "bottom"})
    else:
        chart.set_legend({"none": True})
    chart.show_hidden_data()
    chart.set_size({"width": 470, "height": 250})


def style_axes(chart):
    chart.set_x_axis({"num_font": {"color": MUTED, "size": 8}, "line": {"color": LINE},
                      "major_gridlines": {"visible": False}})
    chart.set_y_axis({"num_font": {"color": MUTED, "size": 8}, "line": {"color": LINE},
                      "major_gridlines": {"visible": True, "line": {"color": LINE, "dash_type": "dash"}}})


def cf_updown(ws, rng):
    ws.conditional_format(rng, {"type": "cell", "criteria": ">", "value": 0, "format": cf_greenf})
    ws.conditional_format(rng, {"type": "cell", "criteria": "<", "value": 0, "format": cf_redf})


def cf_status(ws, rng, kind):
    if kind == "trend":
        pairs = [("UP", cf_bull), ("DOWN", cf_bear), ("SIDE", cf_neu)]
    elif kind == "mom":
        pairs = [("BULL", cf_bull), ("BEAR", cf_bear), ("NEUTRAL", cf_neu)]
    else:  # vol
        pairs = [("HIGH", cf_warn), ("LOW", cf_bull), ("NORMAL", cf_neu)]
    for sub, f in pairs:
        ws.conditional_format(rng, {"type": "text", "criteria": "containing", "value": sub, "format": f})


def cf_matrix(ws, rng):
    tl = rng.split(":")[0]
    rules = [("Bullish", cf_bull), ("Slightly Bullish", cf_slbull), ("Neutral", cf_neu),
             ("Slightly Bearish", cf_slbear), ("Bearish", cf_bear)]
    for txt, f in rules:
        ws.conditional_format(rng, {"type": "formula", "criteria": f'=EXACT({tl},"{txt}")', "format": f})


def matrix_text(ind, tf):
    return (f'=CHOOSE({sc(ind, tf)}+3,"Bearish","Slightly Bearish",'
            f'"Neutral","Slightly Bullish","Bullish")')


# ============================================================================
# ChartData (hidden) — last CHART_N daily bars, auto-updating
# ============================================================================
cdata = wb.add_worksheet("ChartData")
cdata.hide()
CD_COLS = [
    ("Date", A, "f"), ("Open", O, "f"), ("High", H, "f"), ("Low", L, "f"), ("Close", "E", "f"),
    ("EMA20", CL["ema20"], "f"), ("EMA50", CL["ema50"], "f"), ("EMA100", CL["ema100"], "f"),
    ("EMA200", CL["ema200"], "f"), ("RSI", CL["rsi"], "f"), ("RSI70", 70, "c"), ("RSI30", 30, "c"),
    ("MACD", CL["macd"], "f"), ("Signal", CL["macd_sig"], "f"), ("Hist", CL["macd_hist"], "f"),
    ("BBup", CL["bb_up"], "f"), ("BBmid", CL["bb_mid"], "f"), ("BBlow", CL["bb_low"], "f"),
    ("ATR", CL["atr"], "f"), ("HV", CL["hv"], "f"), ("DonUp", CL["don_up"], "f"), ("DonLow", CL["don_low"], "f"),
]
CDI = {name: i for i, (name, _, _) in enumerate(CD_COLS)}
for ci, (name, _, _) in enumerate(CD_COLS):
    cdata.write(0, ci, name)
for i in range(1, CHART_N + 1):
    pe = f"(Metrics!$H$1-{CHART_N}+{i})"
    for ci, (name, src, kind) in enumerate(CD_COLS):
        if kind == "c":
            cdata.write_formula(i, ci, f'=IF({pe}<1,"",{src})')
        else:
            f = (f'=IF({pe}<1,"",INDEX(Calc_Daily!${src}$2:${src}${DROW},{pe}))')
            cdata.write_formula(i, ci, f, date_plain if name == "Date" else None)
CR0, CR1 = 1, CHART_N    # chart data row span (0-based)
cat = ["ChartData", CR0, CDI["Date"], CR1, CDI["Date"]]


def cd_series(name, color, width=1.5, dash=None, y2=False):
    s = {"categories": cat,
         "values": ["ChartData", CR0, CDI[name], CR1, CDI[name]],
         "name": name, "line": {"color": color, "width": width}}
    if dash:
        s["line"]["dash_type"] = dash
    if y2:
        s["y2_axis"] = True
    return s


print("ChartData done")

# ============================================================================
# READ ME
# ============================================================================
rm = new_sheet("Read Me", tab=GOLD, last_col=10, colw=11)
rm.set_column(1, 1, 22)
rm.set_column(2, 10, 12)
rm.merge_range(0, 1, 0, 9, "GOLD TECHNICAL ANALYSIS  —  WORKBOOK GUIDE", fmt_title)
rm.merge_range(1, 1, 1, 9, "Institutional-style, fully self-updating technical model for gold (XAU).", fmt_sub)
draw_nav(rm, 2)
rr = 4
section(rm, rr, 1, 9, "HOW TO USE"); rr += 1
how = [
    "1.  Open the 'Data' sheet — it is the ONLY place you enter anything.",
    "2.  Paste your gold price history with columns in this exact order:  Date | Open | High | Low | Close.",
    "3.  Keep dates sorted oldest → newest (top to bottom). Daily bars are expected.",
    "4.  Replace the shipped SAMPLE data: select the sample rows under the header and paste over them.",
    "5.  Everything else (Dashboard, Trend, Momentum, Volatility, Support & Resistance, Matrix, Outlook)",
    "      recalculates automatically. No buttons, no macros, no manual steps.",
]
for t in how:
    rm.merge_range(rr, 1, rr, 9, t, fmt_body); rr += 1
rr += 1
section(rm, rr, 1, 9, "WHAT IS CALCULATED"); rr += 1
meth = [
    ("Trend", "EMA 20 / 50 / 100 / 200, price-vs-MA position, EMA slope, composite trend score."),
    ("Momentum", "RSI(14), MACD(12,26,9), Stochastic %K/%D (14,3), Stochastic-RSI(14)."),
    ("Volatility", "ATR(14), Bollinger Bands(20,2), 20-bar historical volatility (annualised), volatility regime."),
    ("Support / Resistance", "Fractal swing highs/lows, classic pivots, Fibonacci retracements, Donchian(20)."),
    ("Multi-Timeframe Matrix", "14 indicators classified Bearish→Bullish on Daily, Weekly and Monthly resampled data."),
    ("Outlook", "1-week / 1-month / 3-month bias, target, expected range and confidence."),
]
for h, t in meth:
    rm.merge_range(rr, 1, rr, 2, h, fmt_bodyb)
    rm.merge_range(rr, 3, rr, 9, t, fmt_body); rm.set_row(rr, 28); rr += 1
rr += 1
section(rm, rr, 1, 9, "METHOD NOTES & ASSUMPTIONS"); rr += 1
notes = [
    "• Weekly and monthly series are resampled from your daily data inside the hidden engine (week = Mon–Fri grouping).",
    "• RSI / ATR / ADX use simple-moving-average smoothing (Cutler's method) for clean, stable spreadsheet formulas.",
    "• EMAs are seeded from the first available price and converge as history grows; supply ample history for EMA200.",
    "• Ichimoku cloud uses the standard 9/26/52 settings; the cloud is read 26 bars back for the price-vs-cloud signal.",
    "• Classification maps each indicator to a -2..+2 score; a horizon score (0–100) is the rescaled average of 14 indicators.",
    "• Outlook targets/ranges are derived from horizon bias and historical volatility (±1σ scaled by horizon).",
    f"• Engine capacity: up to {DAILY_MAX:,} daily / {WEEKLY_MAX:,} weekly / {MONTHLY_MAX:,} monthly bars. Hidden engine sheets",
    "   (Calc_Daily/Weekly/Monthly, ChartData, Metrics) can be unhidden to inspect every formula.",
    "• Hidden sheets are right-clickable → Unhide. Do not delete them; the visible sheets read from them.",
]
for t in notes:
    rm.merge_range(rr, 1, rr, 9, t, fmt_note); rm.set_row(rr, 14); rr += 1
rr += 1
rm.merge_range(rr, 1, rr, 9,
    "DISCLAIMER — For research and educational purposes only. Technical analysis is not investment advice. "
    "Markets involve risk; validate all figures before relying on them.", fmt_note)
rm.set_row(rr, 28)
print("Read Me done")

# ============================================================================
# DASHBOARD
# ============================================================================
dsh = new_sheet("Dashboard", tab=GOLD, last_col=11, colw=12)
dsh.set_column(1, 11, 12.5)
dsh.merge_range(0, 1, 0, 10, "GOLD  —  TECHNICAL ANALYSIS DASHBOARD", fmt_title)
mform(dsh, 1, 1, 1, 10,
      '="XAU  ·  As of "&TEXT(AsOfDate,"dd mmm yyyy")&"  ·  Auto-updating workbook"', fmt_sub)
draw_nav(dsh, 2)
dsh.freeze_panes(4, 0)

section(dsh, 4, 1, 10, "MARKET SNAPSHOT")
# KPI cards row 5 (label) / 6-7 (value)
dsh.set_row(6, 22); dsh.set_row(7, 14)
kpis = [
    ("CURRENT PRICE", "=CurrentPrice", fmt_kprice, False),
    ("DAILY", "=RetDay", fmt_kpct, True),
    ("WEEKLY", "=RetWeek", fmt_kpct, True),
    ("MONTHLY", "=RetMonth", fmt_kpct, True),
    ("YTD", "=RetYTD", fmt_kpct, True),
]
for i, (lab, fm, vfmt, updown) in enumerate(kpis):
    c = 1 + i * 2
    mtext(dsh, 5, c, 5, c + 1, lab, fmt_klabel)
    mform(dsh, 6, c, 7, c + 1, fm, vfmt)
    if updown:
        cf_updown(dsh, xlref(6, c, 7, c + 1))

section(dsh, 9, 1, 10, "MARKET REGIME")
dsh.set_row(11, 24)
regime = [("TREND", "=TrendStatus", "trend"), ("MOMENTUM", "=MomentumStatus", "mom"),
          ("VOLATILITY", "=VolatilityStatus", "vol")]
rc = [(1, 3), (4, 6), (7, 10)]
for (lab, fm, kind), (c1, c2) in zip(regime, rc):
    mtext(dsh, 10, c1, 10, c2, lab, fmt_klabel)
    mform(dsh, 11, c1, 11, c2, fm, fmt_chip)
    cf_status(dsh, xlref(11, c1, 11, c2), kind)

# Key levels (resistance / support)
section(dsh, 13, 1, 5, "KEY LEVELS")
section(dsh, 13, 6, 10, "OUTLOOK")
lv = [("Resistance 3", "res3"), ("Resistance 2", "res2"), ("Resistance 1", "res1"),
      ("Support 1", "sup1"), ("Support 2", "sup2"), ("Support 3", "sup3")]
for i, (lab, key) in enumerate(lv):
    r = 14 + i
    fmt_lab = F(bg_color=PANEL, font_color=(BEAR if "Resist" in lab else BULL), font_size=10,
                align="left", valign="vcenter", indent=1, border=1, border_color=LINE)
    mtext(dsh, r, 1, r, 2, lab, fmt_lab)
    mform(dsh, r, 3, r, 5, f'=IFERROR({MVd(key)},"-")', fmt_tdp)

# Outlook mini-table (cols 6-10): Horizon | Bias | Target
mtext(dsh, 14, 6, 14, 6, "Horizon", fmt_th)
mtext(dsh, 14, 7, 14, 8, "Bias", fmt_th)
mtext(dsh, 14, 9, 14, 10, "Target", fmt_th)
for i, (hz, tf, _d) in enumerate(HZ):
    r = 15 + i
    mtext(dsh, r, 6, r, 6, hz, fmt_td)
    mform(dsh, r, 7, r, 8, f'={MObias(hz)}', fmt_td)
    cf_status(dsh, xlref(r, 7, r, 8), "mom")
    mform(dsh, r, 9, r, 10, f'=IFERROR({MO("target", hz)},"-")', fmt_tdp)

# price chart on the dashboard
ch = wb.add_chart({"type": "line"})
for nm, col in [("Close", TXT), ("EMA20", GOLD), ("EMA50", "#4FA3FF"), ("EMA200", "#FF7CA8")]:
    ch.add_series(cd_series(nm, col, width=1.75 if nm == "Close" else 1.25))
style_chart(ch, "Price & Moving Averages (recent)")
style_axes(ch)
dsh.insert_chart(19, 1, ch, {"x_offset": 2, "y_offset": 6})
print("Dashboard done")

# ============================================================================
# TREND
# ============================================================================
tr = new_sheet("Trend", tab="#4FA3FF", last_col=11, colw=13)
tr.merge_range(0, 1, 0, 10, "TREND", fmt_title)
tr.merge_range(1, 1, 1, 10, "Moving-average structure and trend strength (daily).", fmt_sub)
draw_nav(tr, 2); tr.freeze_panes(4, 0)
section(tr, 4, 1, 6, "MOVING AVERAGES (DAILY)")
mtext(tr, 5, 1, 5, 1, "Moving Average", fmt_thl)
mtext(tr, 5, 2, 5, 2, "Value", fmt_th)
mtext(tr, 5, 3, 5, 3, "Price vs MA", fmt_th)
mtext(tr, 5, 4, 5, 6, "Position", fmt_th)
for i, ema in enumerate(["ema20", "ema50", "ema100", "ema200"]):
    r = 6 + i
    mtext(tr, r, 1, r, 1, ema.upper(), fmt_tdl)
    mform(tr, r, 2, r, 2, f'=IFERROR({MV(ema, "D")},"-")', fmt_tdp)
    mform(tr, r, 3, r, 3, f'=IFERROR({MV("close", "D")}/{MV(ema, "D")}-1,"-")', fmt_tdpct)
    cf_updown(tr, xlref(r, 3, r, 3))
    mform(tr, r, 4, r, 6, f'=IF({MV("close", "D")}>{MV(ema, "D")},"Above","Below")', fmt_td)
    cf_status(tr, xlref(r, 4, r, 6), "trend")
# trend score card
section(tr, 11, 1, 6, "TREND SCORE")
mtext(tr, 12, 1, 12, 3, "Composite (-100 bearish … +100 bullish)", fmt_klabel)
trend_score = (f'=ROUND(AVERAGE(Metrics!$C${srow["EMA20"]}:$C${srow["EMA200"]})/2*100,0)')
mform(tr, 12, 4, 12, 6, trend_score, F(bg_color=PANEL, font_color=GOLD, bold=True, font_size=18,
      align="center", valign="vcenter", border=1, border_color=LINE, num_format=INT))
mtext(tr, 13, 1, 13, 3, "Trend status", fmt_klabel)
mform(tr, 13, 4, 13, 6, "=TrendStatus", fmt_chip)
cf_status(tr, xlref(13, 4, 13, 6), "trend")

# charts: candlestick + price vs MAs
stock = wb.add_chart({"type": "stock"})
for nm in ["Open", "High", "Low", "Close"]:
    stock.add_series({"categories": cat,
                      "values": ["ChartData", CR0, CDI[nm], CR1, CDI[nm]],
                      "line": {"none": True}, "marker": {"type": "none"}})
stock.set_up_down_bars({"up": {"fill": {"color": BULL}, "border": {"color": BULL}},
                        "down": {"fill": {"color": BEAR}, "border": {"color": BEAR}}})
stock.set_high_low_lines({"line": {"color": MUTED}})
style_chart(stock, "Gold — Candlestick (recent)", legend=False)
style_axes(stock)
tr.insert_chart(5, 7, stock)

linec = wb.add_chart({"type": "line"})
for nm, color in [("Close", TXT), ("EMA20", GOLD), ("EMA50", "#4FA3FF"),
                  ("EMA100", "#B07CFF"), ("EMA200", "#FF7CA8")]:
    linec.add_series(cd_series(nm, color, width=1.75 if nm == "Close" else 1.1))
style_chart(linec, "Price vs Moving Averages")
style_axes(linec)
tr.insert_chart(20, 7, linec)
print("Trend done")

# ============================================================================
# MOMENTUM
# ============================================================================
mo = new_sheet("Momentum", tab="#B07CFF", last_col=11, colw=13)
mo.merge_range(0, 1, 0, 10, "MOMENTUM", fmt_title)
mo.merge_range(1, 1, 1, 10, "Oscillators across timeframes (RSI, MACD, Stochastic, Stoch-RSI).", fmt_sub)
draw_nav(mo, 2); mo.freeze_panes(4, 0)
section(mo, 4, 1, 6, "DAILY READINGS")
read = [
    ("RSI (14)", f'=IFERROR({MV("rsi", "D")},"-")', fmt_tdnum,
     f'=IF({MV("rsi", "D")}="","-",IF({MV("rsi", "D")}>=70,"Overbought",IF({MV("rsi", "D")}<=30,"Oversold","Neutral")))'),
    ("MACD", f'=IFERROR({MV("macd", "D")},"-")', fmt_tdnum,
     f'=IF({MV("macd_hist", "D")}="","-",IF({MV("macd_hist", "D")}>0,"Bullish","Bearish"))'),
    ("MACD signal", f'=IFERROR({MV("macd_sig", "D")},"-")', fmt_tdnum, '=""'),
    ("Stochastic %K", f'=IFERROR({MV("stoch_k", "D")},"-")', fmt_tdnum,
     f'=IF({MV("stoch_k", "D")}="","-",IF({MV("stoch_k", "D")}>=80,"Overbought",IF({MV("stoch_k", "D")}<=20,"Oversold","Neutral")))'),
    ("Stoch-RSI %K", f'=IFERROR({MV("srsi_k", "D")},"-")', fmt_tdnum,
     f'=IF({MV("srsi_k", "D")}="","-",IF({MV("srsi_k", "D")}>=80,"Overbought",IF({MV("srsi_k", "D")}<=20,"Oversold","Neutral")))'),
]
mtext(mo, 5, 1, 5, 2, "Indicator", fmt_thl)
mtext(mo, 5, 3, 5, 4, "Value", fmt_th)
mtext(mo, 5, 5, 5, 6, "State", fmt_th)
for i, (lab, fm, vfmt, state) in enumerate(read):
    r = 6 + i
    mtext(mo, r, 1, r, 2, lab, fmt_tdl)
    mform(mo, r, 3, r, 4, fm, vfmt)
    mform(mo, r, 5, r, 6, state, fmt_td)
    cf_status(mo, xlref(r, 5, r, 6), "mom")

section(mo, 12, 1, 6, "MULTI-TIMEFRAME SIGNAL")
mtext(mo, 13, 1, 13, 2, "Indicator", fmt_thl)
mtext(mo, 13, 3, 13, 3, "Daily", fmt_th)
mtext(mo, 13, 4, 13, 4, "Weekly", fmt_th)
mtext(mo, 13, 5, 13, 6, "Monthly", fmt_th)
mind = ["RSI", "MACD", "Stochastic", "Stoch RSI"]
for i, ind in enumerate(mind):
    r = 14 + i
    mtext(mo, r, 1, r, 2, ind, fmt_tdl)
    mform(mo, r, 3, r, 3, matrix_text(ind, "D"), fmt_matrix)
    mform(mo, r, 4, r, 4, matrix_text(ind, "W"), fmt_matrix)
    mform(mo, r, 5, r, 6, matrix_text(ind, "M"), fmt_matrix)
    cf_matrix(mo, xlref(r, 3, r, 6))

rsi_c = wb.add_chart({"type": "line"})
rsi_c.add_series(cd_series("RSI", GOLD, 1.6))
rsi_c.add_series(cd_series("RSI70", BEAR, 1, "dash"))
rsi_c.add_series(cd_series("RSI30", BULL, 1, "dash"))
style_chart(rsi_c, "RSI (14)")
style_axes(rsi_c)
rsi_c.set_y_axis({"min": 0, "max": 100, "num_font": {"color": MUTED, "size": 8},
                  "line": {"color": LINE}, "major_gridlines": {"visible": True, "line": {"color": LINE, "dash_type": "dash"}}})
mo.insert_chart(5, 7, rsi_c)

macd_col = wb.add_chart({"type": "column"})
macd_col.add_series({"categories": cat, "values": ["ChartData", CR0, CDI["Hist"], CR1, CDI["Hist"]],
                     "name": "Hist", "fill": {"color": "#5b6675"}, "border": {"none": True}})
macd_line = wb.add_chart({"type": "line"})
macd_line.add_series(cd_series("MACD", GOLD, 1.5))
macd_line.add_series(cd_series("Signal", "#4FA3FF", 1.3))
macd_col.combine(macd_line)
style_chart(macd_col, "MACD (12,26,9)")
style_axes(macd_col)
mo.insert_chart(20, 7, macd_col)
print("Momentum done")

# ============================================================================
# VOLATILITY
# ============================================================================
vo = new_sheet("Volatility", tab="#2EBD85", last_col=11, colw=13)
vo.merge_range(0, 1, 0, 10, "VOLATILITY", fmt_title)
vo.merge_range(1, 1, 1, 10, "ATR, Bollinger Bands, historical volatility and regime.", fmt_sub)
draw_nav(vo, 2); vo.freeze_panes(4, 0)
section(vo, 4, 1, 6, "DAILY READINGS")
vrows = [
    ("ATR (14)", f'=IFERROR({MV("atr", "D")},"-")', fmt_tdnum),
    ("ATR % of price", f'=IFERROR({MV("atr", "D")}/{MV("close", "D")},"-")', fmt_tdpct),
    ("Bollinger bandwidth", f'=IFERROR({MV("bb_bw", "D")},"-")', fmt_tdpct),
    ("Bollinger %B", f'=IFERROR({MV("bb_pctb", "D")},"-")', fmt_tdpct),
    ("Hist. vol (ann.)", f'=IFERROR({MV("hv", "D")},"-")', fmt_tdpct),
]
mtext(vo, 5, 1, 5, 3, "Metric", fmt_thl)
mtext(vo, 5, 4, 5, 6, "Value", fmt_th)
for i, (lab, fm, vfmt) in enumerate(vrows):
    r = 6 + i
    mtext(vo, r, 1, r, 3, lab, fmt_tdl)
    mform(vo, r, 4, r, 6, fm, vfmt)
section(vo, 12, 1, 6, "VOLATILITY REGIME")
mtext(vo, 13, 1, 13, 3, "Current regime", fmt_klabel)
mform(vo, 13, 4, 13, 6, "=VolatilityStatus", fmt_chip)
cf_status(vo, xlref(13, 4, 13, 6), "vol")
mtext(vo, 14, 1, 14, 3, "HV vs 100-bar average", fmt_klabel)
mform(vo, 14, 4, 14, 6, f'=IFERROR({MV("hv", "D")}/{MV("hv_avg", "D")}-1,"-")',
      F(bg_color=PANEL, font_color=TXT, bold=True, font_size=14, align="center",
        valign="vcenter", border=1, border_color=LINE, num_format=PCT0))
cf_updown(vo, xlref(14, 4, 14, 6))

bb_c = wb.add_chart({"type": "line"})
bb_c.add_series(cd_series("BBup", "#6f7b8a", 1, "dash"))
bb_c.add_series(cd_series("Close", TXT, 1.6))
bb_c.add_series(cd_series("BBmid", GOLD, 1.1, "dash"))
bb_c.add_series(cd_series("BBlow", "#6f7b8a", 1, "dash"))
style_chart(bb_c, "Bollinger Bands (20, 2)")
style_axes(bb_c)
vo.insert_chart(5, 7, bb_c)

atr_c = wb.add_chart({"type": "line"})
atr_c.add_series(cd_series("ATR", GOLD, 1.5))
atr_c.add_series(cd_series("HV", "#4FA3FF", 1.3, y2=True))
style_chart(atr_c, "ATR (14) & Historical Volatility")
style_axes(atr_c)
atr_c.set_y2_axis({"num_font": {"color": MUTED, "size": 8}})
vo.insert_chart(20, 7, atr_c)
print("Volatility done")

# ============================================================================
# SUPPORT & RESISTANCE
# ============================================================================
sr = new_sheet("Support & Resistance", tab="#F6465D", last_col=11, colw=13)
sr.merge_range(0, 1, 0, 10, "SUPPORT & RESISTANCE", fmt_title)
sr.merge_range(1, 1, 1, 10, "Swing levels, pivots, Fibonacci retracements and Donchian channel.", fmt_sub)
draw_nav(sr, 2); sr.freeze_panes(4, 0)

section(sr, 4, 1, 3, "SWING LEVELS")
mtext(sr, 5, 1, 5, 2, "Level", fmt_thl); mtext(sr, 5, 3, 5, 3, "Price", fmt_th)
swing = [("Resistance 3", "res3", BEAR), ("Resistance 2", "res2", BEAR), ("Resistance 1", "res1", BEAR),
         ("Current price", "close", GOLD), ("Support 1", "sup1", BULL),
         ("Support 2", "sup2", BULL), ("Support 3", "sup3", BULL)]
for i, (lab, key, col) in enumerate(swing):
    r = 6 + i
    mtext(sr, r, 1, r, 2, lab, F(bg_color=PANEL, font_color=col, font_size=10, align="left",
          valign="vcenter", indent=1, border=1, border_color=LINE, bold=(key == "close")))
    src = MV(key, "D") if key == "close" else MVd(key)
    mform(sr, r, 3, r, 3, f'=IFERROR({src},"-")', fmt_tdp)

section(sr, 4, 5, 7, "CLASSIC PIVOTS")
mtext(sr, 5, 5, 5, 6, "Level", fmt_thl); mtext(sr, 5, 7, 5, 7, "Price", fmt_th)
piv = [("R3", "piv_r3", BEAR), ("R2", "piv_r2", BEAR), ("R1", "piv_r1", BEAR),
       ("Pivot (PP)", "pp", GOLD), ("S1", "piv_s1", BULL), ("S2", "piv_s2", BULL), ("S3", "piv_s3", BULL)]
for i, (lab, key, col) in enumerate(piv):
    r = 6 + i
    mtext(sr, r, 5, r, 6, lab, F(bg_color=PANEL, font_color=col, font_size=10, align="left",
          valign="vcenter", indent=1, border=1, border_color=LINE, bold=(key == "pp")))
    mform(sr, r, 7, r, 7, f'=IFERROR({MVd(key)},"-")', fmt_tdp)

section(sr, 4, 9, 10, "DONCHIAN (20)")
don = [("Upper", "don_up"), ("Mid", "don_mid"), ("Lower", "don_low")]
for i, (lab, key) in enumerate(don):
    r = 6 + i
    mtext(sr, r, 9, r, 9, lab, fmt_tdl)
    mform(sr, r, 10, r, 10, f'=IFERROR({MV(key, "D")},"-")', fmt_tdp)

section(sr, 14, 1, 4, "FIBONACCI RETRACEMENT (recent major swing)")
mtext(sr, 15, 1, 15, 2, "Ratio", fmt_thl); mtext(sr, 15, 3, 15, 4, "Price", fmt_th)
for i, rt in enumerate(FIB_RATIOS):
    r = 16 + i
    key = f"fib_{int(rt * 1000):04d}"
    mtext(sr, r, 1, r, 2, f"{rt * 100:.1f}%", fmt_tdl)
    mform(sr, r, 3, r, 4, f'=IFERROR({MVd(key)},"-")', fmt_tdp)
sr.merge_range(24, 1, 26, 4,
    "Swing highs/lows use a confirmed 4-bar fractal. Resistance/Support pick the nearest confirmed "
    "swing levels above/below the current price. Fibonacci spans the most recent major swing within the "
    f"last {FIB_LB} bars; direction is detected automatically.", fmt_note)

don_c = wb.add_chart({"type": "line"})
don_c.add_series(cd_series("DonUp", BEAR, 1.1, "dash"))
don_c.add_series(cd_series("Close", TXT, 1.6))
don_c.add_series(cd_series("DonLow", BULL, 1.1, "dash"))
style_chart(don_c, "Price & Donchian(20)")
style_axes(don_c)
sr.insert_chart(5, 11, don_c)
print("S&R done")

# ============================================================================
# MATRIX
# ============================================================================
mx = new_sheet("Matrix", tab=GOLD, last_col=12, colw=13)
mx.set_column(1, 1, 22)
mx.merge_range(0, 1, 0, 8, "MULTI-TIMEFRAME INDICATOR MATRIX", fmt_title)
mx.merge_range(1, 1, 1, 8, "14 indicators classified Bearish → Bullish on each timeframe.", fmt_sub)
draw_nav(mx, 2); mx.freeze_panes(5, 0)
section(mx, 4, 1, 4, "INDICATOR MATRIX")
mtext(mx, 5, 1, 5, 1, "Indicator", fmt_thl)
mtext(mx, 5, 2, 5, 2, "Daily", fmt_th)
mtext(mx, 5, 3, 5, 3, "Weekly", fmt_th)
mtext(mx, 5, 4, 5, 4, "Monthly", fmt_th)
m_first = 6
for i, ind in enumerate(IND):
    r = m_first + i
    mtext(mx, r, 1, r, 1, ind, fmt_tdl)
    mform(mx, r, 2, r, 2, matrix_text(ind, "D"), fmt_matrix)
    mform(mx, r, 3, r, 3, matrix_text(ind, "W"), fmt_matrix)
    mform(mx, r, 4, r, 4, matrix_text(ind, "M"), fmt_matrix)
cf_matrix(mx, xlref(m_first, 2, m_first + NIND - 1, 4))

# horizon score table
section(mx, 4, 6, 9, "HORIZON SCORE")
mtext(mx, 5, 6, 5, 6, "Horizon", fmt_th)
mtext(mx, 5, 7, 5, 7, "Score", fmt_th)
mtext(mx, 5, 8, 5, 9, "Bias", fmt_th)
for i, (lab, tf) in enumerate([("Daily", "D"), ("Weekly", "W"), ("Monthly", "M")]):
    r = 6 + i
    mtext(mx, r, 6, r, 6, lab, fmt_td)
    mform(mx, r, 7, r, 7, f'={MH("score100", tf)}&"/100"', fmt_td)
    mform(mx, r, 8, r, 9, f'={MH("bias", tf)}', fmt_td)
    cf_status(mx, xlref(r, 8, r, 9), "mom")

# percentage table (feeds chart)
section(mx, 11, 6, 9, "SIGNAL BREAKDOWN")
pct_hdr_row = 12
mtext(mx, pct_hdr_row, 6, pct_hdr_row, 6, "", fmt_th)
mtext(mx, pct_hdr_row, 7, pct_hdr_row, 7, "Bullish", fmt_th)
mtext(mx, pct_hdr_row, 8, pct_hdr_row, 8, "Neutral", fmt_th)
mtext(mx, pct_hdr_row, 9, pct_hdr_row, 9, "Bearish", fmt_th)
for i, (lab, tf) in enumerate([("Daily", "D"), ("Weekly", "W"), ("Monthly", "M")]):
    r = pct_hdr_row + 1 + i
    mtext(mx, r, 6, r, 6, lab, fmt_tdl)
    mform(mx, r, 7, r, 7, f'={MH("pct_bull", tf)}', fmt_tdpct)
    mform(mx, r, 8, r, 8, f'={MH("pct_neu", tf)}', fmt_tdpct)
    mform(mx, r, 9, r, 9, f'={MH("pct_bear", tf)}', fmt_tdpct)
pr0 = pct_hdr_row + 1
pct_c = wb.add_chart({"type": "column"})
for nm, ccol, color in [("Bullish", 7, BULL), ("Neutral", 8, NEUTRAL), ("Bearish", 9, BEAR)]:
    pct_c.add_series({"name": ["Matrix", pct_hdr_row, ccol],
                      "categories": ["Matrix", pr0, 6, pr0 + 2, 6],
                      "values": ["Matrix", pr0, ccol, pr0 + 2, ccol],
                      "fill": {"color": color}, "border": {"none": True}})
style_chart(pct_c, "Bullish / Neutral / Bearish share")
style_axes(pct_c)
pct_c.set_y_axis({"min": 0, "max": 1, "num_format": "0%", "num_font": {"color": MUTED, "size": 8},
                  "line": {"color": LINE}, "major_gridlines": {"visible": True, "line": {"color": LINE, "dash_type": "dash"}}})
mx.insert_chart(17, 6, pct_c)
print("Matrix done")

# ============================================================================
# OUTLOOK
# ============================================================================
ol = new_sheet("Outlook", tab=GOLD, last_col=11, colw=15)
ol.set_column(1, 1, 12)
ol.merge_range(0, 1, 0, 9, "OUTLOOK", fmt_title)
ol.merge_range(1, 1, 1, 9, "Forward view derived from the indicator matrix and volatility.", fmt_sub)
draw_nav(ol, 2); ol.freeze_panes(4, 0)
section(ol, 4, 1, 9, "DIRECTIONAL OUTLOOK")
hdrs = ["Horizon", "Bias", "Target Price", "Expected Range", "Confidence"]
spans = [(1, 1), (2, 3), (4, 5), (6, 7), (8, 9)]
for h, (c1, c2) in zip(hdrs, spans):
    mtext(ol, 5, c1, 5, c2, h, fmt_th)
for i, (hz, tf, _d) in enumerate(HZ):
    r = 6 + i
    mtext(ol, r, 1, r, 1, {"1W": "1 Week", "1M": "1 Month", "3M": "3 Months"}[hz], fmt_tdl)
    mform(ol, r, 2, r, 3, f'={MObias(hz)}', fmt_td)
    cf_status(ol, xlref(r, 2, r, 3), "mom")
    mform(ol, r, 4, r, 5, f'=IFERROR({MO("target", hz)},"-")', fmt_tdp)
    mform(ol, r, 6, r, 7,
          f'=IFERROR(TEXT({MO("low", hz)},"#,##0")&"  -  "&TEXT({MO("high", hz)},"#,##0"),"-")', fmt_td)
    mform(ol, r, 8, r, 9,
          f'=IFERROR(TEXT({MO("conf", hz)}/100,"0%")&"  ("&IF({MO("conf", hz)}>=66,"High",'
          f'IF({MO("conf", hz)}>=40,"Medium","Low"))&")","-")', fmt_td)
ol.set_row(6, 22); ol.set_row(7, 22); ol.set_row(8, 22)
section(ol, 11, 1, 9, "METHOD")
ol.merge_range(12, 1, 15, 9,
    "Bias for each horizon comes from the matrix horizon score (1 Week ← Daily, 1 Month ← Weekly, "
    "3 Months ← Monthly). The target applies a volatility-scaled drift in the direction of bias; the "
    "expected range is a ±1-sigma band using annualised historical volatility scaled to the horizon "
    "(5 / 21 / 63 trading days). Confidence reflects how far the horizon score sits from neutral (50). "
    "All figures recompute automatically from the Data sheet.", fmt_note)

out_c = wb.add_chart({"type": "line"})
# expected range visual: target line + bands per horizon using a small helper table
orow_tbl = 17
mtext(ol, orow_tbl, 1, orow_tbl, 1, "", fmt_th)
for j, (hz, _t, _d) in enumerate(HZ):
    mtext(ol, orow_tbl, 2 + j, orow_tbl, 2 + j, {"1W": "1 Week", "1M": "1 Month", "3M": "3 Months"}[hz], fmt_th)
for k, (lab, key, color) in enumerate([("High", "high", BEAR), ("Target", "target", GOLD), ("Low", "low", BULL)]):
    r = orow_tbl + 1 + k
    mtext(ol, r, 1, r, 1, lab, fmt_tdl)
    for j, (hz, _t, _d) in enumerate(HZ):
        mform(ol, r, 2 + j, r, 2 + j, f'=IFERROR({MO(key, hz)},"-")', fmt_tdp)
for k, (lab, color) in enumerate([("High", BEAR), ("Target", GOLD), ("Low", BULL)]):
    r = orow_tbl + 1 + k
    out_c.add_series({"name": ["Outlook", r, 1],
                      "categories": ["Outlook", orow_tbl, 2, orow_tbl, 4],
                      "values": ["Outlook", r, 2, r, 4],
                      "line": {"color": color, "width": 1.75},
                      "marker": {"type": "circle", "size": 6, "fill": {"color": color}}})
style_chart(out_c, "Projected target & expected range")
style_axes(out_c)
ol.insert_chart(5, 11, out_c)
print("Outlook done")

# ============================================================================
# Tab order + activation, then write
# ============================================================================
desired = ["Dashboard", "Trend", "Momentum", "Volatility", "Support & Resistance",
           "Matrix", "Outlook", "Data", "Read Me",
           "ChartData", "Metrics", "Calc_Daily", "Calc_Weekly", "Calc_Monthly"]
wb.worksheets_objs.sort(key=lambda w: desired.index(w.name))
dsh.activate()
dsh.set_first_sheet()

wb.close()
print("WROTE", OUT)
