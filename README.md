# Gold Technical Analysis Workbook

A professional, **fully self-updating** Excel workbook for gold (XAU) technical
analysis. You enter data in **one** place — the `Data` sheet — and every other
sheet (Dashboard, Trend, Momentum, Volatility, Support & Resistance, Matrix,
Outlook) recalculates automatically with native Excel formulas, charts and
conditional formatting. **No macros / VBA.**

## Deliverable

- **`Gold_Technical_Analysis.xlsx`** — the workbook. This is the product.
- `build_workbook.py` — the generator that produces the `.xlsx` (kept so the
  workbook can be regenerated or extended). Requires `XlsxWriter`.

## How to use

1. Open `Gold_Technical_Analysis.xlsx` and go to the **`Data`** sheet.
2. Paste your gold history in the exact column order **Date | Open | High | Low | Close**,
   sorted oldest → newest. (The shipped rows are *sample* data — paste over them.)
3. Everything else updates automatically. See the in-workbook **`Read Me`** sheet
   for methodology and assumptions.

## Sheets

| Sheet | Contents |
|-------|----------|
| Dashboard | Current price, D/W/M/YTD returns, trend/momentum/volatility regime, key levels, outlook |
| Trend | EMA 20/50/100/200, price-vs-MA, trend score, candlestick + MA charts |
| Momentum | RSI, MACD, Stochastic, Stoch-RSI, multi-timeframe signals + charts |
| Volatility | ATR, Bollinger Bands, historical volatility, regime + charts |
| Support & Resistance | Swing levels, classic pivots, Fibonacci retracements, Donchian |
| Matrix | 14 indicators × Daily/Weekly/Monthly classification, horizon scores, breakdown chart |
| Outlook | 1-week / 1-month / 3-month bias, target, expected range, confidence |
| Data | **The only input sheet** |
| Read Me | Usage + methodology |

Hidden engine sheets (`Calc_Daily`, `Calc_Weekly`, `Calc_Monthly`, `ChartData`,
`Metrics`) hold the calculations; unhide them to inspect every formula.

## Regenerate

```bash
pip install XlsxWriter
python3 build_workbook.py
```

## Notes

- Weekly/monthly series are resampled from your daily data inside the engine.
- RSI/ATR/ADX use SMA-based (Cutler) smoothing for stable spreadsheet formulas.
- Engine capacity: up to 2,000 daily / 440 weekly / 120 monthly bars.
- For research/education only; not investment advice.
