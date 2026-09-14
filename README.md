# TEDPIX 5-year Total Index report

Builds an Excel workbook of the Tehran Stock Exchange **Total Index** (شاخص کل / TEDPIX, TSETMC code `32097828799138957`) for the last five calendar years.

The workbook has:

- A 5-year daily trend chart with the **first close** and **last close** of each year labeled
- An overlay chart of each year's path by day of year
- A first/last table (Gregorian and Jalali years)
- Daily OHLC data

For the current year, the last value is **today's last price**, not year-end.

## Setup

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

## Run

```powershell
.\.venv\Scripts\python.exe build_tedpix_excel.py
```

Output: `output/TEDPIX_Total_Index_5Y.xlsx`

## Data source

The intended source is [TSETMC IndexInfo](https://www.tsetmc.com/IndexInfo/32097828799138957). If that API is blocked (common with VPN), the script uses:

1. The TSETMC daily archive on Hugging Face (`Farmaanaa/iran_tedpix_stock_index_daily`)
2. TGJU for sessions after that archive
3. Today's last price from the TGJU live ticker
