# Google Sheets backtesting template for MES, M2K, MYM, and MCL

This repository contains a Google Apps Script that instantly builds a Google Sheets workbook for backtesting the MES, M2K, MYM, and MCL futures contracts. The template supports one to two trades per day (or no trades), color-coded drop-downs for outcomes, automatic R-multiple calculations, and a summary dashboard with the key metrics you asked for.

## What the template provides
- **Trades sheet**
  - Columns for Date, Ticker, Trade #, Points, Risk, Result (Green/Red/Breakeven/No Trade), R Multiple, and Notes.
  - Drop-down lists for tickers and results with conditional formatting that colors each outcome (Green, Red, Breakeven in gray/black text, and No Trade in amber).
  - R Multiple column is pre-filled with a formula so each row auto-calculates risk/reward when Points and Risk are entered.
- **Summary sheet**
  - Total points, total risk, average risk, overall risk/reward (total points divided by total risk), average R Multiple, win/loss ratio, win rate, trade counts, no-trade days, and ticker-level trade counts and points.

## How to use the template in Google Sheets
1. Create a new Google Sheet.
2. Open **Extensions → Apps Script**.
3. Replace the default code with the contents of [`apps_script/backtesting_template.gs`](apps_script/backtesting_template.gs).
4. Save, then run **Run → buildTemplate** once (Google will prompt for permissions).
5. Back in the sheet, use the **Backtest Template** menu → **Build/Reset Template** if you ever want to rebuild the layout.
6. Start logging trades on the **Trades** tab:
   - Set **Result** to **Green**, **Red**, **Breakeven**, or **No Trade** (color-coded automatically).
   - Enter **Points** and **Risk**; **R Multiple** will compute automatically per row.
7. Check the **Summary** tab for totals, averages, win/loss ratio, win rate, per-ticker counts, and no-trade days.

### Customization tips
- Increase `maxRows` in `setupTradesSheet_` if you want more pre-filled rows.
- Add more tickers by editing the ticker list in both `setupTradesSheet_` and `setupSummarySheet_`.
- Formulas are all on the Summary sheet, so you can tweak metrics without touching the Trades tab.
