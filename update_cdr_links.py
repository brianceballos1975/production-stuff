"""
Extract CDR manufacturing file links from ScanInventory xlsx
and update master_products.db — uses SAX streaming to handle 49MB sharedStrings.xml
"""
import sys, sqlite3, zipfile, re
from xml.etree.ElementTree import iterparse
from io import BytesIO

sys.stdout.reconfigure(encoding='utf-8')
sys.stdout.flush()

XLSX = r'C:\Users\breez\Downloads\ScanInventory (1).xlsx'
DB   = r'C:\Users\breez\Downloads\Category+Listings+Report_04-10-2026\master_products.db'
NS   = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

def log(msg):
    print(msg, flush=True)

# ── Step 1: stream shared strings ───────────────────────────────────────────
log('Streaming sharedStrings.xml ...')
shared = []
with zipfile.ZipFile(XLSX) as zf:
    with zf.open('xl/sharedStrings.xml') as f:
        buf = ''
        for event, elem in iterparse(f, events=('end',)):
            if elem.tag == f'{{{NS}}}si':
                text = ''.join(t for t in elem.itertext())
                shared.append(text)
                elem.clear()
log(f'Shared strings: {len(shared):,}')

# ── Step 2: parse workbook rels to map rId → sheet file path ────────────────
log('Parsing workbook structure ...')
with zipfile.ZipFile(XLSX) as zf:
    wb_data = zf.read('xl/workbook.xml')
    rels_data = zf.read('xl/_rels/workbook.xml.rels')

    wb_tree_root = None
    for event, elem in iterparse(BytesIO(wb_data), events=('end',)):
        if elem.tag.endswith('}sheets'):
            wb_tree_root = elem
            break

    sheet_info = []  # [(name, rId)]
    if wb_tree_root is not None:
        for s in wb_tree_root:
            name = s.attrib.get('name', '')
            rid  = s.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id', '')
            sheet_info.append((name, rid))

    rid_to_path = {}
    for event, elem in iterparse(BytesIO(rels_data), events=('end',)):
        if elem.tag.endswith('}Relationship'):
            rid_to_path[elem.attrib['Id']] = 'xl/' + elem.attrib['Target'].lstrip('/')

log(f'Sheets: {len(sheet_info)}')

# ── Step 3: for each sheet, read the CDR column using iterparse ──────────────
def col_index(col_str):
    """Convert Excel column letter(s) to 0-based index. e.g. A→0, D→3"""
    idx = 0
    for ch in col_str.upper():
        idx = idx * 26 + (ord(ch) - 64)
    return idx - 1

def parse_cell_ref(ref):
    """Split 'D5' → ('D', 5)"""
    m = re.match(r'([A-Z]+)(\d+)', ref)
    return (m.group(1), int(m.group(2))) if m else (None, None)

records = {}  # sku → (cdr_url, score)

with zipfile.ZipFile(XLSX) as zf:
    for sheet_name, rid in sheet_info:
        if sheet_name == 'Welcome':
            continue
        path = rid_to_path.get(rid)
        if not path or path not in zf.namelist():
            continue

        # First pass: read header row to find SKU (col 0) and CDR col index
        header = {}  # col_idx → header text
        sku_idx = 0
        cdr_idx = None
        in_row1 = False
        row1_done = False

        with zf.open(path) as f:
            for event, elem in iterparse(f, events=('start', 'end')):
                tag = elem.tag.split('}')[-1]
                if event == 'start' and tag == 'row':
                    rnum = int(elem.attrib.get('r', 0))
                    if rnum == 1:
                        in_row1 = True
                    elif rnum > 1:
                        row1_done = True
                        break
                elif event == 'end' and tag == 'c' and in_row1 and not row1_done:
                    ref = elem.attrib.get('r', '')
                    col_str, _ = parse_cell_ref(ref)
                    if col_str is None:
                        continue
                    cidx = col_index(col_str)
                    t = elem.attrib.get('t', '')
                    v_el = elem.find(f'{{{NS}}}v')
                    if v_el is not None and v_el.text:
                        val = shared[int(v_el.text)] if t == 's' else v_el.text
                    else:
                        val = ''
                    header[cidx] = str(val).strip()
                    if 'CDR' in str(val).upper():
                        cdr_idx = cidx

        if cdr_idx is None:
            continue

        # Second pass: read data rows
        count = 0
        with zf.open(path) as f:
            current_row = {}
            cur_rnum = 0
            for event, elem in iterparse(f, events=('start', 'end')):
                tag = elem.tag.split('}')[-1]
                if event == 'start' and tag == 'row':
                    cur_rnum = int(elem.attrib.get('r', 0))
                    current_row = {}
                elif event == 'end' and tag == 'c':
                    if cur_rnum <= 1:
                        continue
                    ref = elem.attrib.get('r', '')
                    col_str, _ = parse_cell_ref(ref)
                    if col_str is None:
                        continue
                    cidx = col_index(col_str)
                    if cidx not in (sku_idx, cdr_idx):
                        continue
                    t = elem.attrib.get('t', '')
                    v_el = elem.find(f'{{{NS}}}v')
                    if v_el is not None and v_el.text:
                        val = shared[int(v_el.text)] if t == 's' else v_el.text
                    else:
                        val = ''
                    current_row[cidx] = str(val).strip()
                    elem.clear()
                elif event == 'end' and tag == 'row':
                    if cur_rnum <= 1:
                        continue
                    sku = current_row.get(sku_idx, '')
                    cdr = current_row.get(cdr_idx, '')
                    if not sku or sku == 'None':
                        continue
                    score = 2 if cdr.lower().endswith('.cdr') else (1 if cdr.startswith('http') else 0)
                    if sku not in records or score > records[sku][1]:
                        records[sku] = (cdr or None, score)
                    count += 1

        log(f'  {sheet_name}: {count:,} rows, cdr_col={cdr_idx}')

log(f'\nTotal unique SKUs: {len(records):,}')
full_cdr = sum(1 for v in records.values() if v[0] and v[0].lower().endswith('.cdr'))
log(f'  Full .cdr URLs:   {full_cdr:,}')
log(f'  Base/folder URLs: {len(records) - full_cdr:,}')

# ── Step 4: update SQLite ────────────────────────────────────────────────────
log('\nUpdating database ...')
conn = sqlite3.connect(DB, timeout=120)
conn.execute('PRAGMA journal_mode=WAL')
conn.execute('PRAGMA synchronous=NORMAL')
conn.execute('PRAGMA cache_size=10000')

cols = [r[1] for r in conn.execute('PRAGMA table_info(products)').fetchall()]
if 'cdr_file' not in cols:
    conn.execute('ALTER TABLE products ADD COLUMN cdr_file TEXT')
    log('Added cdr_file column')

conn.execute('CREATE TABLE IF NOT EXISTS _cdr_temp (sku TEXT PRIMARY KEY, cdr_file TEXT)')
conn.execute('DELETE FROM _cdr_temp')
conn.executemany('INSERT OR REPLACE INTO _cdr_temp VALUES (?,?)',
                 [(sku, v[0]) for sku, v in records.items()])
conn.execute('''
    UPDATE products
    SET cdr_file = (SELECT cdr_file FROM _cdr_temp WHERE _cdr_temp.sku = products.sku)
    WHERE sku IN (SELECT sku FROM _cdr_temp)
''')
updated = conn.execute('SELECT changes()').fetchone()[0]
conn.execute('DROP TABLE _cdr_temp')
conn.execute('CREATE INDEX IF NOT EXISTS idx_cdr_file ON products(cdr_file)')
conn.commit()
conn.close()

log(f'DB rows updated: {updated:,}')

# ── Step 5: verify ───────────────────────────────────────────────────────────
conn = sqlite3.connect(DB)
total   = conn.execute("SELECT COUNT(*) FROM products WHERE cdr_file IS NOT NULL AND cdr_file != ''").fetchone()[0]
full_db = conn.execute("SELECT COUNT(*) FROM products WHERE cdr_file LIKE '%.cdr'").fetchone()[0]
sample  = conn.execute("SELECT sku, cdr_file FROM products WHERE cdr_file LIKE '%.cdr' LIMIT 4").fetchall()
conn.close()

log(f'\nProducts with CDR link in DB : {total:,}')
log(f'  Full .cdr file URLs        : {full_db:,}')
log(f'  Base/folder URLs           : {total - full_db:,}')
log('\nSample:')
for r in sample:
    log(f'  SKU : {r[0]}')
    log(f'  URL : {r[1]}')
log('\nDone.')
