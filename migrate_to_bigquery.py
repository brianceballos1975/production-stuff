"""
migrate_to_bigquery.py
Migrates all rows from master_products.db (SQLite) into BigQuery.
  Project : maindb-79403
  Dataset : products
  Table   : master_products
All columns are loaded as STRING to match the SQLite TEXT schema.
"""

import re
import sqlite3
import time
from pathlib import Path

import pandas as pd
from google.cloud import bigquery

# ── config ────────────────────────────────────────────────────────────────────
PROJECT   = "maindb-79403"
DATASET   = "products"
TABLE     = "master_products"
DB_PATH   = Path(__file__).parent / "master_products.db"
CHUNK     = 25_000          # rows per BigQuery load job
# ─────────────────────────────────────────────────────────────────────────────


def sanitize_col(name: str) -> str:
    """Make a column name safe for BigQuery (letters, digits, underscores; no leading digit)."""
    name = re.sub(r"[^a-zA-Z0-9_]", "_", name)
    if name and name[0].isdigit():
        name = "_" + name
    return name or "_col"


def main():
    client = bigquery.Client(project=PROJECT)

    # ── 1. ensure dataset exists ───────────────────────────────────────────
    ds_ref = bigquery.DatasetReference(PROJECT, DATASET)
    try:
        client.get_dataset(ds_ref)
        print(f"Dataset {PROJECT}.{DATASET} already exists.")
    except Exception:
        ds = bigquery.Dataset(ds_ref)
        ds.location = "US"
        client.create_dataset(ds)
        print(f"Created dataset {PROJECT}.{DATASET}")

    # ── 2. read column names from SQLite ──────────────────────────────────
    con = sqlite3.connect(f"file:{DB_PATH}?mode=ro", uri=True)
    cur = con.cursor()
    cur.execute("SELECT * FROM products LIMIT 0")
    raw_cols   = [d[0] for d in cur.description]
    clean_cols = [sanitize_col(c) for c in raw_cols]

    # warn about any renames
    renames = [(r, c) for r, c in zip(raw_cols, clean_cols) if r != c]
    if renames:
        print("Column renames:")
        for r, c in renames:
            print(f"  {r!r} → {c!r}")

    # ── 3. build BigQuery schema (all STRING) ─────────────────────────────
    bq_schema = [bigquery.SchemaField(c, "STRING") for c in clean_cols]

    # ── 4. (re)create table ───────────────────────────────────────────────
    table_ref = f"{PROJECT}.{DATASET}.{TABLE}"
    table_obj = bigquery.Table(table_ref, schema=bq_schema)
    client.delete_table(table_ref, not_found_ok=True)
    client.create_table(table_obj)
    print(f"Created table {table_ref} with {len(bq_schema)} columns.")

    # ── 5. stream chunks ──────────────────────────────────────────────────
    cur.execute("SELECT COUNT(*) FROM products")
    total = cur.fetchone()[0]
    print(f"Total rows to migrate: {total:,}")

    job_config = bigquery.LoadJobConfig(
        schema=bq_schema,
        write_disposition=bigquery.WriteDisposition.WRITE_APPEND,
    )

    uploaded = 0
    offset   = 0
    t0       = time.time()

    while True:
        df = pd.read_sql_query(
            f"SELECT * FROM products LIMIT {CHUNK} OFFSET {offset}",
            con,
        )
        if df.empty:
            break

        # rename columns to sanitized names
        df.columns = clean_cols

        # ensure all values are strings (or None)
        df = df.where(pd.notnull(df), None).astype(object)
        for col in df.columns:
            df[col] = df[col].apply(lambda v: str(v) if v is not None else None)

        job = client.load_table_from_dataframe(df, table_ref, job_config=job_config)
        job.result()  # wait for job to complete

        uploaded += len(df)
        offset   += CHUNK
        elapsed   = time.time() - t0
        pct       = uploaded / total * 100
        rate      = uploaded / elapsed if elapsed > 0 else 0
        eta       = (total - uploaded) / rate if rate > 0 else 0
        print(
            f"  {uploaded:>7,} / {total:,}  ({pct:5.1f}%)  "
            f"{rate:,.0f} rows/s  ETA {eta:.0f}s"
        )

    con.close()
    print(f"\nDone! {uploaded:,} rows loaded in {time.time()-t0:.1f}s")
    print(f"Table: https://console.cloud.google.com/bigquery?project={PROJECT}"
          f"&ws=!1m5!1m4!4m3!1s{PROJECT}!2s{DATASET}!3s{TABLE}")


if __name__ == "__main__":
    main()
