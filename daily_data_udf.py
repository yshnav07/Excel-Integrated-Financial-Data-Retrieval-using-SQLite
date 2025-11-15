import sqlite3
import xlwings as xw
import time
import logging
from logging.handlers import RotatingFileHandler
from functools import lru_cache
import configparser

# Read config
cfg = configparser.ConfigParser()
cfg.read("config.ini")

DB = cfg.get("database", "sqlite_path", fallback=r"C:\Users\vaish\Downloads\prices.db")
TABLE = cfg.get("database", "table_name", fallback="prices")
LOG_FILE = cfg.get("logging", "log_file", fallback="query_log.txt")

# Logging setup
logger = logging.getLogger("excel_udf")
if not logger.handlers:
    handler = RotatingFileHandler(LOG_FILE, maxBytes=1048576, backupCount=2)
    formatter = logging.Formatter("%(asctime)s | %(message)s")
    handler.setFormatter(formatter)
    logger.addHandler(handler)
    logger.setLevel(logging.INFO)

def connection():
    conn = sqlite3.connect(DB)
    conn.row_factory = sqlite3.Row
    return conn

def run_query(query, params):
    conn = connection()
    cur = conn.cursor()
    cur.execute(query, params)
    rows = cur.fetchall()
    conn.close()
    return rows

def log_call(func_name, params, start_time, error=None):
    duration = (time.perf_counter() - start_time) * 1000
    status = "Success" if not error else f"Failed: {error}"
    logger.info(f"{func_name} | params={params} | time_ms={duration:.2f} | {status}")

# -------------------------
#       UDF FUNCTIONS
# -------------------------



@xw.func
def get_daily_data(accord_code, field, date):
    start = time.perf_counter()
    try:
        query = f"SELECT {field} FROM {TABLE} WHERE accord_code=? AND date=?"
        rows = run_query(query, (accord_code, date))
        log_call("get_daily_data", (accord_code, field, date), start)
        return rows[0][field] if rows else "No data"
    except Exception as e:
        log_call("get_daily_data", (accord_code, field, date), start, e)
        return f"Error: {e}"

@xw.func
def get_series(accord_code, field, start_date, end_date):
    start = time.perf_counter()
    try:
        query = f"SELECT date, {field} FROM {TABLE} WHERE accord_code=? AND date BETWEEN ? AND ? ORDER BY date"
        rows = run_query(query, (accord_code, start_date, end_date))
        log_call("get_series", (accord_code, field, start_date, end_date), start)
        if not rows:
            return "No data"
        data = [["Date", field]]
        for r in rows:
            data.append([r["date"], r[field]])
        return data
    except Exception as e:
        log_call("get_series", (accord_code, field, start_date, end_date), start, e)
        return f"Error: {e}"

@xw.func
def get_daily_matrix(date, field):
    start = time.perf_counter()
    try:
        query = f"SELECT accord_code, company_name, sector, mcap_category, {field} FROM {TABLE} WHERE date=?"
        rows = run_query(query, (date,))
        log_call("get_daily_matrix", (date, field), start)
        if not rows:
            return "No data"
        result = [["accord_code", "company_name", "sector", "mcap_category", field]]
        for r in rows:
            result.append([
                r["accord_code"], r["company_name"],
                r["sector"], r["mcap_category"], r[field]
            ])
        return result
    except Exception as e:
        log_call("get_daily_matrix", (date, field), start, e)
        return f"Error: {e}"

@xw.func
def get_all_prices(accord_code, field):
    start = time.perf_counter()
    try:
        query = f"SELECT date, {field} FROM {TABLE} WHERE accord_code=? ORDER BY date"
        rows = run_query(query, (accord_code,))
        log_call("get_all_prices", (accord_code, field), start)
        if not rows:
            return "No data"
        result = [["Date", field]]
        for r in rows:
            result.append([r["date"], r[field]])
        return result
    except Exception as e:
        log_call("get_all_prices", (accord_code, field), start, e)
        return f"Error: {e}"
