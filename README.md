# Daily Data Excel UDFs (SQLite backend)

## Overview
This project provides Excel UDFs to query a local SQLite prices database:
- `get_daily_data(accord_code, field, date)` — single value
- `get_series(accord_code, field, start_date, end_date)` — date/value series
- `get_daily_matrix(date, field)` — table of all companies on a date (spill)
- `get_all_prices(accord_code, field)` — all dates for a company (spill)

Implemented in `daily_data_udf.py`. Uses SQLite, parameterized queries, LRU caching and logging.

---

## Files
- `daily_data_udf.py` — Python UDF server file
- `prices.db` — sample SQLite DB (place in project root)
- `config.ini` — configuration file
- `schema.sql` — SQL index creation script
- `query_log.txt` — generated log file (created automatically)
- `README.md` — this file

---

## Setup

1. Clone/copy project to: `C:\Users\<you>\Desktop\itus`
2. Create and activate virtual environment (Windows):

```powershell
cd C:\Users\<you>\Desktop\itus
python -m venv venv
venv\Scripts\activate

## Dependencies
pip install --upgrade pip
pip install xlwings==0.30.12 pywin32
