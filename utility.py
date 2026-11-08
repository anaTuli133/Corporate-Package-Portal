import os
import time
import threading
import oracledb

DB_USER = os.getenv("DB_USER_203", "dwh_user")
DB_PASS = os.getenv("DB_PASS_203", "dwh_user_123")
DB_DSN = os.getenv("DB_SVC_203", "192.168.61.203:1521/dwhdb02")

_pool = None
_pool_lock = threading.Lock()


def get_pool():
    """Connection pool: created once, so a new connection is not opened on every search (faster)."""
    global _pool
    if _pool is None:
        with _pool_lock:
            if _pool is None:
                _pool = oracledb.create_pool(
                    user=DB_USER, password=DB_PASS, dsn=DB_DSN,
                    min=2, max=8, increment=1,
                )
    return _pool


# ---------- Package list (cached 10 min) ----------
_pkg_cache = {"ts": 0, "data": []}
PKG_TTL = 600


def fetch_packages(force=False):
    if not force and _pkg_cache["data"] and time.time() - _pkg_cache["ts"] < PKG_TTL:
        return _pkg_cache["data"]

    sql = """
        SELECT DISTINCT OFFERING_NAME
        FROM offer_dim
        WHERE UPPER(OFFERING_NAME) LIKE '%CORPORATE%'
        ORDER BY OFFERING_NAME
    """
    with get_pool().acquire() as conn:
        with conn.cursor() as cur:
            cur.execute(sql)
            names = [r[0] for r in cur.fetchall() if r[0]]

    _pkg_cache.update(ts=time.time(), data=names)
    return names


# ---------- Subscriber + revenue ----------
MAX_ROWS = int(os.getenv("MAX_ROWS", 1_000_000))   # safety limit for one search result


def fetch_subscriber_revenue(package, start_date, end_date):
    """
    package   : 'ALL' or exact OFFERING_NAME (must contain CORPORATE)
    start_date: datetime.date (inclusive)
    end_date  : datetime.date (inclusive)

    Returns (rows, total_revenue)
    rows = [[package, 'YYYY-MM-DD', msisdn_with_880, revenue_of_that_day], ...]
    One row per package + date + subscriber (day-wise revenue).
    """
    sql = """
        SELECT
            o.OFFERING_NAME,
            TO_CHAR(TRUNC(d.DATE_VALUE), 'YYYY-MM-DD') AS DAY,
            CASE
                WHEN s.MSISDN LIKE '880%' THEN s.MSISDN
                WHEN s.MSISDN LIKE '01%'  THEN '88'  || s.MSISDN
                ELSE                           '880' || s.MSISDN
            END AS MSISDN_WITH_880,
            SUM(NVL(s.VOICE_REVENUE, 0) + NVL(s.DATA_REVENUE, 0)
              + NVL(s.SMS_REVENUE, 0)   + NVL(s.COMBO_BUNDLE_REVENUE, 0)) AS TOTAL_REVENUE
        FROM subscriber_usage_analysis s
        JOIN offer_dim o ON s.PRODUCT_ID = o.OFFERING_ID
        JOIN date_dim  d ON s.DATE_KEY   = d.DATE_KEY
        WHERE UPPER(o.OFFERING_NAME) LIKE '%CORPORATE%'
          {pkg_filter}
          AND d.DATE_VALUE >= :sd
          AND d.DATE_VALUE <  :ed + 1
        GROUP BY o.OFFERING_NAME, TRUNC(d.DATE_VALUE), s.MSISDN
        ORDER BY o.OFFERING_NAME, TRUNC(d.DATE_VALUE), MSISDN_WITH_880
    """
    params = {"sd": start_date, "ed": end_date}
    if package and package != "ALL":
        sql = sql.format(pkg_filter="AND o.OFFERING_NAME = :pkg")
        params["pkg"] = package
    else:
        sql = sql.format(pkg_filter="")

    rows, total = [], 0.0
    with get_pool().acquire() as conn:
        conn.call_timeout = 300_000  # 5 min
        with conn.cursor() as cur:
            cur.arraysize = 5000     # bulk fetch -> much faster
            cur.prefetchrows = 5001
            cur.execute(sql, params)
            while True:
                chunk = cur.fetchmany()
                if not chunk:
                    break
                for name, day, msisdn, rev in chunk:
                    rev = round(float(rev or 0), 2)
                    total += rev
                    rows.append([name, day, msisdn, rev])
                if len(rows) > MAX_ROWS:
                    raise ValueError(
                        f"Result is too large (more than {MAX_ROWS:,} records). "
                        "Please select a shorter date range or a specific package."
                    )
    return rows, round(total, 2)
