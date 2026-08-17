"""
cama_extract.py — Pull CAMA sales data from Oracle iasWorld
Reads the SQL script, injects the calculated date range, executes the query.
"""

import re
import sys
import oracledb
import pandas as pd
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
from datetime import date, timedelta

from credentials import ORACLE_USER, ORACLE_PASSWORD, ORACLE_DSN

# Path to the SQL file (same folder as this script)
SQL_FILE = Path(__file__).parent / (
    "PROD - IASW.sqlCAMAvsMLSwCITYwNOPARwADDPARwADDKEYwTOTBASEFINv4.sql"
)


def get_week_dates(run_date=None):
    """
    Calculate CAMA query date range for a Monday run.
    Rule: previous Saturday - 7 days (inclusive) → previous Saturday (exclusive).
    Example: run Monday April 14 → range April 5 ≤ SALEDT < April 12.

    Returns (start_date, end_date) as date objects.
    """
    today = run_date or date.today()
    days_since_saturday = (today.weekday() - 5) % 7
    last_saturday = today - timedelta(days=days_since_saturday)
    start_date    = last_saturday - timedelta(days=7)
    end_date      = last_saturday
    return start_date, end_date


def extract_cama(output_path: Path, run_date=None, start_date=None, end_date=None) -> pd.DataFrame:
    """
    Run the CAMA SQL query and save results to CSV.
    Returns a DataFrame of the results.

    Pass explicit start_date / end_date (date objects) to override the
    automatic weekly range calculated from run_date.
    The SQL uses  SALEDT >= start_date  AND  SALEDT < end_date.
    """
    if start_date is None or end_date is None:
        start_date, end_date = get_week_dates(run_date)
    print(f"  Sale date range: {start_date} to {end_date} (exclusive)")

    # ── Read SQL and inject dates ──────────────────────────────────────────────
    if not SQL_FILE.exists():
        print(f"  ✗ SQL file not found: {SQL_FILE}")
        sys.exit(1)

    sql = SQL_FILE.read_text(encoding="utf-8")

    # Replace the two hardcoded dates in the SQL file
    sql = re.sub(
        r"S\.SALEDT >= TO_DATE\('[^']+',\s*'YYYY-MM-DD'\)",
        f"S.SALEDT >= TO_DATE('{start_date}','YYYY-MM-DD')",
        sql,
    )
    sql = re.sub(
        r"S\.SALEDT <\s+TO_DATE\('[^']+',\s*'YYYY-MM-DD'\)",
        f"S.SALEDT <  TO_DATE('{end_date}','YYYY-MM-DD')",
        sql,
    )

    # Update TAXYR to current year (changes once per year during reappraisal)
    current_year = (run_date or date.today()).year
    sql = re.sub(r"P\.TAXYR\s*=\s*\d{4}", f"P.TAXYR = {current_year}", sql)
    sql = re.sub(r"D\.TAXYR\s*=\s*P\.TAXYR", "D.TAXYR = P.TAXYR", sql)  # leave joins as-is

    # Strip trailing semicolon — oracledb cursor.execute() does not allow it
    sql = sql.rstrip().rstrip(';')

    # ── Connect and execute ────────────────────────────────────────────────────
    print(f"  Connecting to Oracle ({ORACLE_DSN})...")
    try:
        conn = oracledb.connect(user=ORACLE_USER, password=ORACLE_PASSWORD, dsn=ORACLE_DSN)
    except oracledb.Error as e:
        print(f"  ✗ Oracle connection failed: {e}")
        sys.exit(1)

    try:
        with conn.cursor() as cur:
            cur.execute(sql)
            cols = [col[0].upper() for col in cur.description]
            rows = cur.fetchall()
    except oracledb.Error as e:
        print(f"  ✗ Query failed: {e}")
        conn.close()
        sys.exit(1)
    finally:
        conn.close()

    df = pd.DataFrame(rows, columns=cols)
    print(f"  ✓ {len(df)} rows returned from Oracle")

    df.to_csv(output_path, index=False)
    print(f"  ✓ Saved: {output_path.name}")
    return df
