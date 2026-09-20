# TEDPIX Historical Market Report

## Overview

This is a practical Python reporting tool for the Tehran Stock Exchange **Total Index** (شاخص کل / TEDPIX, TSETMC instrument `32097828799138957`). It collects daily index history for the last five Gregorian calendar years and writes a formatted Excel workbook. The report includes daily prices, first-and-last close comparisons by year, and charts of the five-year trend and each year's path. Dates are stored in both Gregorian and Jalali calendars. For the current year, the last value is **today's last price**, not a year-end close.

## What It Produces

The script writes `output/TEDPIX_Total_Index_5Y.xlsx` with these sheets:

| Sheet | Contents |
| --- | --- |
| **Charts** | 5-year trend chart, year-overlay chart, and a first/last close snapshot |
| **Year first-last** | Gregorian and Jalali yearly comparison tables |
| **Daily** | Date, Jalali date, OHLC, and year markers |
| **Year overlay data** | Each year's close aligned by day of year |
| **Notes** | Data-source notes and how to re-run |

In more detail:

* **Historical TEDPIX data** — daily Total Index values for the last five Gregorian calendar years, ending today.
* **OHLC information** — open, high, low, and close on the Daily sheet. Open can be missing when history comes from the Hugging Face TSETMC archive rather than the live TSETMC API.
* **Yearly comparisons** — first trading-day close, last trading-day close, absolute change, percent change, and trading-day count. The Gregorian table also includes year high and year low of the close.
* **Charts** — a 5-year daily trend with green diamonds on each year's first close and red diamonds on each year's last close, plus an overlay of each year's path by calendar day of year.
* **Excel output** — a styled `.xlsx` workbook generated with openpyxl, including filters, freeze panes, and number formats.
* **Jalali / Gregorian dates** — every daily row has both calendars. Summaries are grouped by Gregorian year and again by Jalali (Iranian) year via `jdatetime`.

## Data Pipeline

```
Data Source
→ Collection
→ Processing
→ Analysis
→ Excel Report
```

1. **Data source** — primary: TSETMC `GetIndexB2History` for [IndexInfo/32097828799138957](https://www.tsetmc.com/IndexInfo/32097828799138957). If that API is unreachable, history comes from the TSETMC daily archive on Hugging Face (`Farmaanaa/iran_tedpix_stock_index_daily`), later sessions from TGJU, and today's last price from the TGJU live ticker.
2. **Collection** — HTTP requests assemble daily rows (`date`, `jalali`, `open`, `high`, `low`, `close`).
3. **Processing** — rows are concatenated, de-duplicated by date, sorted, and clipped to the five-year window. Today's live last price replaces any same-day history row.
4. **Analysis** — first/last close, change, and trading days are computed per Gregorian year and per Jalali year. Closes are also aligned by day of year for the overlay chart.
5. **Excel report** — `build_tedpix_excel.py` writes the workbook above.

## Example Output

A generated sample workbook is in the repository:

[`output/TEDPIX_Total_Index_5Y.xlsx`](output/TEDPIX_Total_Index_5Y.xlsx)

Re-running the script overwrites that file with a fresh report ending on the current session.

## Technical Stack

* Python 3
* pandas
* requests
* openpyxl
* jdatetime
* pyarrow (Parquet cache for the Hugging Face fallback)

## Project Structure

```
build_tedpix_excel.py          data collection, analysis, and Excel generation
requirements.txt               Python dependencies
output/TEDPIX_Total_Index_5Y.xlsx   generated report (sample included)
data/                          local Parquet cache (not committed)
```

One script does the full pipeline. Cached TSETMC archive files under `data/` are ignored by git.

## How to Run

Network access is required. A virtual environment is recommended.

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe build_tedpix_excel.py
```

macOS / Linux:

```bash
python -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python build_tedpix_excel.py
```

Output: `output/TEDPIX_Total_Index_5Y.xlsx`

## Skills Demonstrated

* Python
* Financial Data Collection
* Data Processing
* Excel Reporting
* Time-Series Analysis
* Persian/Jalali Date Handling
