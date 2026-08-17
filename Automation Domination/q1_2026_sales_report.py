"""
q1_2026_sales_report.py — Q1 2026 residential sales extracts for boss report
  Dataset 1: R-class sales, price <= $100,000
  Dataset 2: R-class sales, price <= $800,000, NOPAR > 1
"""

import sys
import oracledb
import pandas as pd
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

sys.path.insert(0, str(Path(__file__).parent))
import credentials

OUTPUT_DIR = Path(__file__).parent.parent / "MLSvsCAMA" / "4-20-26"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

BASE_SQL = """
SELECT
  P.JUR,
  P.PARID,
  P.TAXYR,
  S.NOPAR,
  S.PRICE,
  S.SALEDT
FROM PARDAT P
JOIN SALES S
  ON S.PARID = P.PARID
WHERE P.TAXYR  = 2026
  AND P.CLASS  = 'R'
  AND S.SALEDT BETWEEN TO_DATE('1-JAN-2026', 'DD-MON-YYYY')
                   AND TO_DATE('31-MAR-2026', 'DD-MON-YYYY')
  AND {price_filter}
  {nopar_filter}
ORDER BY S.SALEDT, P.PARID
"""

QUERIES = [
    {
        "label":        "Q1 2026 Residential Sales - $100k and Under",
        "filename":     "Q1_2026_Residential_100k_and_Under.xlsx",
        "price_filter": "S.PRICE <= 100000",
        "nopar_filter": "",
    },
    {
        "label":        "Q1 2026 Residential Sales - $800k and Under, Multi-Parcel",
        "filename":     "Q1_2026_Residential_800k_Under_MultiParcel.xlsx",
        "price_filter": "S.PRICE < 800000",
        "nopar_filter": "AND S.NOPAR > 1",
    },
]

print("Connecting to Oracle...")
conn = oracledb.connect(
    user=credentials.ORACLE_USER,
    password=credentials.ORACLE_PASSWORD,
    dsn=credentials.ORACLE_DSN,
)
print("  Connected")

for q in QUERIES:
    sql = BASE_SQL.format(
        price_filter=q["price_filter"],
        nopar_filter=q["nopar_filter"],
    ).rstrip()
    sql = sql.rstrip(";")

    print(f"\nRunning: {q['label']}...")
    df = pd.read_sql(sql, conn)

    # Format sale date as date only (no time)
    if "SALEDT" in df.columns:
        df["SALEDT"] = pd.to_datetime(df["SALEDT"]).dt.date

    out_path = OUTPUT_DIR / q["filename"]
    df.to_excel(out_path, index=False)
    print(f"  {len(df)} rows -> {out_path.name}")

conn.close()
print("\nDone.")
