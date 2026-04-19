"""
Rebuild master_products.db keeping only columns that have at least one non-empty value.
Drops 1,297 empty columns, retains 229 populated columns.
"""
import sqlite3, os, shutil, sys
from datetime import datetime

sys.stdout.reconfigure(encoding='utf-8')

DB      = r'C:\Users\breez\Downloads\Category+Listings+Report_04-10-2026\master_products.db'
DB_NEW  = r'C:\Users\breez\Downloads\Category+Listings+Report_04-10-2026\master_products_slim.db'
DB_BAK  = r'C:\Users\breez\Downloads\Category+Listings+Report_04-10-2026\backups\master_products_pre_slim_{}.db'.format(
              datetime.now().strftime('%Y%m%d_%H%M%S'))

def log(msg): print(msg, flush=True)

# ── 1. Identify columns with data ──────────────────────────────────────────
log('Scanning columns for data...')
conn = sqlite3.connect(DB)
conn.execute('PRAGMA journal_mode=WAL')
all_cols = [r[1] for r in conn.execute('PRAGMA table_info(products)').fetchall()]

keep = []
for i, col in enumerate(all_cols):
    n = conn.execute(
        f'SELECT COUNT(*) FROM products WHERE "{col}" IS NOT NULL AND TRIM(CAST("{col}" AS TEXT)) != ""'
    ).fetchone()[0]
    if n > 0:
        keep.append(col)
    if (i + 1) % 200 == 0:
        log(f'  {i+1}/{len(all_cols)} checked — keeping {len(keep)} so far')

conn.close()
log(f'Columns to keep : {len(keep)}')
log(f'Columns to drop : {len(all_cols) - len(keep)}')

# ── 2. Build slim DB via CREATE TABLE AS SELECT ─────────────────────────────
if os.path.exists(DB_NEW):
    os.remove(DB_NEW)

log(f'\nBuilding slim DB: {DB_NEW}')
col_list = ', '.join(f'"{c}"' for c in keep)

src  = sqlite3.connect(DB)
dest = sqlite3.connect(DB_NEW)
dest.execute('PRAGMA journal_mode=WAL')
dest.execute('PRAGMA synchronous=NORMAL')
dest.execute('PRAGMA cache_size=20000')

# Create table schema
src.row_factory = sqlite3.Row
sample = src.execute(f'SELECT {col_list} FROM products LIMIT 0').description
col_defs = ', '.join(f'"{d[0]}" TEXT' for d in sample)
dest.execute(f'CREATE TABLE products ({col_defs})')

# Copy data in chunks
log('Copying rows...')
CHUNK = 5000
offset = 0
total  = 0
while True:
    rows = src.execute(f'SELECT {col_list} FROM products LIMIT {CHUNK} OFFSET {offset}').fetchall()
    if not rows:
        break
    dest.executemany(f'INSERT INTO products VALUES ({",".join(["?"]*len(keep))})', rows)
    dest.commit()
    total  += len(rows)
    offset += CHUNK
    if total % 50000 == 0:
        log(f'  {total:,} rows copied...')

src.close()
log(f'  {total:,} rows copied total')

# ── 3. Rebuild indexes ───────────────────────────────────────────────────────
log('Creating indexes...')
index_map = {
    'idx_sku':          'sku',
    'idx_status':       'status',
    'idx_product_type': 'product_type',
    'idx_brand':        'brand_name',
    'idx_asin':         'asin',
    'idx_parent_sku':   'parent_sku',
    'idx_source':       'source_file',
    'idx_cdr_file':     'cdr_file',
}
for name, col in index_map.items():
    if col in keep:
        dest.execute(f'CREATE INDEX IF NOT EXISTS {name} ON products({col})')

dest.execute('PRAGMA optimize')
dest.commit()
dest.close()

# ── 4. Swap files ────────────────────────────────────────────────────────────
log('\nBacking up original DB...')
os.makedirs(os.path.dirname(DB_BAK), exist_ok=True)
shutil.move(DB, DB_BAK)
log(f'  Backup: {DB_BAK}')

shutil.move(DB_NEW, DB)
log(f'  Slim DB moved to: {DB}')

# ── 5. Report ────────────────────────────────────────────────────────────────
orig_mb = os.path.getsize(DB_BAK) / 1024 / 1024
new_mb  = os.path.getsize(DB)     / 1024 / 1024
log(f'\nOriginal DB : {orig_mb:,.0f} MB')
log(f'Slim DB     : {new_mb:,.0f} MB  ({100*(1-new_mb/orig_mb):.0f}% smaller)')
log(f'Columns     : {len(all_cols)} → {len(keep)}  ({len(all_cols)-len(keep)} removed)')
log(f'Rows        : {total:,}')

# Quick verify
conn = sqlite3.connect(DB)
r = conn.execute("SELECT COUNT(*), COUNT(sku), COUNT(cdr_file) FROM products").fetchone()
conn.close()
log(f'\nVerification — total rows: {r[0]:,} | with SKU: {r[1]:,} | with CDR: {r[2]:,}')
log('\nDone.')
