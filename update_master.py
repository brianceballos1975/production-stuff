"""
update_master.py — Merge a new Amazon listing file into master_products.csv + master_products.db

Usage:
    python update_master.py <new_file>

Supported formats: .xlsm, .xlsx, .csv

Behavior:
  - New rows whose SKU does not exist in master → ADDED
  - Rows whose SKU already exists in master    → UPDATED (new data wins)
  - master_products.csv is updated in place
  - master_products.db (SQLite) is kept in sync
  - A timestamped backup of the CSV is created before every update
"""

import sys
import os
import re
import shutil
import sqlite3
import pandas as pd
from datetime import datetime

BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
MASTER_CSV = os.path.join(BASE_DIR, "master_products.csv")
MASTER_DB  = os.path.join(BASE_DIR, "master_products.db")
BACKUP_DIR = os.path.join(BASE_DIR, "backups")

SKU_COL    = "SKU"
SOURCE_COL = "source_file"


# ---------- helpers ----------

def clean_col(c):
    c = str(c).strip().lower()
    c = re.sub(r"[^a-z0-9_]", "_", c)
    c = re.sub(r"_+", "_", c).strip("_")
    return c or "col"


def make_unique_cols(raw_cols):
    used = set()
    result = []
    for c in [clean_col(x) for x in raw_cols]:
        candidate, i = c, 1
        while candidate in used:
            candidate = f"{c}_dup{i}"
            i += 1
        used.add(candidate)
        result.append(candidate)
    return result


def load_new_file(path: str) -> pd.DataFrame:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        df = pd.read_csv(path, dtype=str)
    elif ext in (".xlsx", ".xlsm"):
        try:
            df = pd.read_excel(path, sheet_name="Template", header=3, skiprows=[4, 5], dtype=str)
        except Exception:
            df = pd.read_excel(path, dtype=str)
    else:
        raise ValueError(f"Unsupported file type: {ext}")

    df = df.dropna(how="all")

    if SKU_COL in df.columns:
        pass
    else:
        df = df.rename(columns={df.columns[2]: SKU_COL})

    df = df[df[SKU_COL].notna() & (df[SKU_COL].str.strip() != "")]
    df[SOURCE_COL] = os.path.basename(path)
    return df


def backup_master():
    os.makedirs(BACKUP_DIR, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    dst = os.path.join(BACKUP_DIR, f"master_products_{ts}.csv")
    shutil.copy2(MASTER_CSV, dst)
    return dst


def merge(master: pd.DataFrame, new: pd.DataFrame) -> tuple:
    all_cols = list(dict.fromkeys(list(master.columns) + list(new.columns)))
    master   = master.reindex(columns=all_cols)
    new      = new.reindex(columns=all_cols)

    existing_skus = set(master[SKU_COL].str.strip())
    new_skus      = new[SKU_COL].str.strip()

    updates   = new[new_skus.isin(existing_skus)]
    additions = new[~new_skus.isin(existing_skus)]

    if not updates.empty:
        master = master[~master[SKU_COL].str.strip().isin(set(updates[SKU_COL].str.strip()))]
        master = pd.concat([master, updates], ignore_index=True, sort=False)

    if not additions.empty:
        master = pd.concat([master, additions], ignore_index=True, sort=False)

    return master, len(updates), len(additions)


# ---------- SQLite sync ----------

def sync_db(merged: pd.DataFrame):
    """Rebuild SQLite DB from the merged DataFrame."""
    print("Syncing SQLite database...")

    db_cols = make_unique_cols(merged.columns)
    df_db   = merged.copy()
    df_db.columns = db_cols

    if os.path.exists(MASTER_DB):
        os.remove(MASTER_DB)

    conn = sqlite3.connect(MASTER_DB)
    chunk = 10000
    first = True
    for start in range(0, len(df_db), chunk):
        df_db.iloc[start:start + chunk].to_sql(
            "products", conn,
            if_exists="replace" if first else "append",
            index=False
        )
        first = False

    # Indexes
    for name, col in [
        ("idx_sku",          "sku"),
        ("idx_status",       "status"),
        ("idx_product_type", "product_type"),
        ("idx_brand",        "brand_name"),
        ("idx_asin",         "asin"),
        ("idx_parent_sku",   "parent_sku"),
        ("idx_source",       "source_file"),
    ]:
        try:
            conn.execute(f"CREATE INDEX IF NOT EXISTS {name} ON products({col})")
        except sqlite3.OperationalError:
            pass  # column may not exist for all file types

    conn.commit()
    conn.close()
    print(f"  DB updated: {os.path.getsize(MASTER_DB)/1024/1024:.1f} MB")


# ---------- main ----------

def main():
    if len(sys.argv) < 2:
        print("Usage: python update_master.py <new_file.xlsm|xlsx|csv>")
        sys.exit(1)

    new_path = sys.argv[1]
    if not os.path.exists(new_path):
        print(f"File not found: {new_path}")
        sys.exit(1)
    if not os.path.exists(MASTER_CSV):
        print(f"Master CSV not found: {MASTER_CSV}")
        sys.exit(1)

    print(f"Loading master CSV...")
    import io
    with open(MASTER_CSV, "rb") as f:
        master = pd.read_csv(
            io.TextIOWrapper(f, encoding="utf-8-sig", errors="replace"),
            dtype=str, low_memory=False
        )
    print(f"  {len(master):,} existing products")

    print(f"Loading new file: {new_path}")
    new_df = load_new_file(new_path)
    print(f"  {len(new_df):,} rows in new file")

    backup_path = backup_master()
    print(f"Backup saved: {backup_path}")

    merged, n_updated, n_added = merge(master, new_df)

    print("Saving master CSV...")
    merged.to_csv(MASTER_CSV, index=False, encoding="utf-8-sig")

    sync_db(merged)

    print(f"\nDone.")
    print(f"  Updated : {n_updated:,} existing listings")
    print(f"  Added   : {n_added:,} new listings")
    print(f"  Total   : {len(merged):,} products in master")


if __name__ == "__main__":
    main()
