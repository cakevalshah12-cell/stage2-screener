# Stage 2 + fundamentals screener (NSE)

## Runs by itself

GitHub runs the scan every **Saturday at about 8 AM IST** (`.github/workflows/weekly-scan.yml`)
on all NSE stocks, then:

- updates the dashboard website (`docs/`, published with GitHub Pages),
- saves the week's Excel in `results/` and the data in `docs/data/history/`.

To run it now: **Actions** tab → *Weekly Stage 2 scan* → **Run workflow**.
GitHub may start scheduled runs some minutes late.

## Install (once)

```powershell
python -m pip install --user yfinance pandas requests openpyxl
```

## Run

```powershell
python stage2_screener.py                                  # Nifty 500, Upstox prices (no login needed)
python stage2_screener.py --symbols MAHABANK PARAGMILK     # specific stocks
python stage2_screener.py --tickers-file my_list.txt       # your own list
python stage2_screener.py --screener-file screener.xlsx    # Screener.in export (needs an 'NSE Code' column)
python stage2_screener.py --loose                          # relaxed rules, more results
python stage2_screener.py --source dhan                    # Dhan prices (needs keys, see below)
python stage2_screener.py --source yahoo                   # Yahoo prices (split/bonus adjusted)
```

Results print to the screen and save to `stage2_results_<date>.xlsx`
(sheets: Stage2 candidates, All scanned, Rules).

Weekly bars use completed weeks only until Friday 4 pm; add `--include-current-week`
to use the still-forming bar.

## Data sources

| Source | Login | Notes |
|---|---|---|
| `upstox` (default) | None for historical candles | Official exchange data. Set `UPSTOX_ACCESS_TOKEN` only if Upstox starts requiring it. |
| `dhan` | `DHAN_CLIENT_ID` + `DHAN_ACCESS_TOKEN` | Needs Dhan's Data API subscription. Tokens expire, so regenerate when it says rejected. |
| `yahoo` | None | Prices adjusted for splits and bonuses. |

Fundamentals (growth, ROE, debt) always come from Yahoo Finance; brokers don't provide them.

Promoter holding, its 1-year change, and **promoter pledge** come from NSE's quarterly
shareholding-pattern filings (only for the Stage 2 candidates, so ~1 minute).
A stock fails if more than 5% of promoter shares are pledged; a promoter stake fall of
3+ points in a year is flagged in `fund_notes`. These NSE pages are not an official API:
if NSE blocks the request, the scan continues and says so. Skip it with `--no-nse`.

### Setting keys (PowerShell, your own machine only)

For the current window only:

```powershell
$env:DHAN_CLIENT_ID = "your-client-id"
$env:DHAN_ACCESS_TOKEN = "your-token"
```

Keep keys out of the script and never share them in chat or screenshots.

## Caveats

- Broker candles may not be adjusted for splits or bonuses. If a stock had one in the last
  3 years and its result looks odd, re-check it with `--source yahoo`.
- Yahoo fundamentals have gaps (base effects, lenders fail the debt test).
  Treat "weak/unknown funda" as "check by hand".
- No valuation filter: look at the `pe` column yourself.
- Screening only, not investment advice.
