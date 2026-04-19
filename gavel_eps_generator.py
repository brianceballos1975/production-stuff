"""
gavel_eps_generator.py — Generate SVG production layouts for custom gavel band orders.

Pulls pending (not-yet-shipped) gavel orders from ShipStation, downloads Amazon
customization data, and writes a single print-ready SVG layout (24" × 12" pages,
3 columns × 10 rows, 0.25" spacing) with every design on one file.

Usage:
    python gavel_eps_generator.py [--days N] [--output DIR]

Options:
    --days N      Look back N days for pending orders (default: 1)
    --output DIR  Output folder for layout file (default: ./gavel_eps)
"""

import argparse
import csv
import io
import json
import os
import struct
import sys
import urllib.request
import zipfile
from datetime import datetime, timedelta, timezone
from pathlib import Path


def _load_env(path: str = ".env") -> None:
    """Load key=value pairs from a .env file into os.environ (no dependencies)."""
    env_file = Path(path)
    if not env_file.exists():
        # Try relative to this script's directory
        env_file = Path(__file__).parent / path
    if not env_file.exists():
        return
    for line in env_file.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, val = line.partition("=")
        os.environ.setdefault(key.strip(), val.strip())

_load_env()

# Fix Windows cp1252 stdout choking on Unicode customer names/text
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

# ── config ────────────────────────────────────────────────────────────────────

API_KEY       = os.environ.get("SHIPSTATION_API_KEY", "")
BASE_URL      = "https://api.shipstation.com/v2"
PAGE_SIZE     = 100
TEMPLATE_PATH = r"C:\Users\breez\Downloads\gavelband_template.eps"

# ── Trello ────────────────────────────────────────────────────────────────────
TRELLO_API_KEY     = os.environ.get("TRELLO_API_KEY", "")
TRELLO_TOKEN       = os.environ.get("TRELLO_TOKEN", "")
TRELLO_BOARD_NAME  = "customs"
TRELLO_LIST_NAME   = "Test List"

# ── coordinate system: 1 SVG user unit = 1 point (1/72 inch) ─────────────────

PT = 72.0   # points per inch

# Band dimensions in points (from template BoundingBox)
BAND_W  = 495.25
BAND_H  = 68.65
BAND_CX = BAND_W / 2   # 247.625 pt  horizontal center
BAND_CY = BAND_H / 2   # 34.325  pt  vertical center

# Layout grid (all in points)
PAGE_W = 24 * PT        # 1728 pt
PAGE_H = 12 * PT        # 864  pt
COLS   = 3
ROWS   = 10
GAP    = 0.25 * PT      # 18   pt

# Font size (points) by number of non-empty text lines
FONT_SIZE_PT = {1: 15.0, 2: 13.0, 3: 11.0, 4: 10.0}
LEADING = 1.25

# CMYK(0, 0.993347, 1, 0) → RGB #FF0200
BAND_STROKE_COLOR = "#FF0200"
BAND_STROKE_WIDTH = 0.5   # points (fallback)

# SKU substrings that identify gavel band products
GAVEL_SKU_PATTERNS = ["CGVL", "GVLSB", "GFCSTM-GVL", "GF-VLU"]

# Google Fonts families (lowercase)
_GOOGLE_FONTS = {
    "lato", "open sans", "roboto", "montserrat", "oswald", "raleway",
    "source sans pro", "ubuntu", "nunito", "poppins", "merriweather",
    "playfair display", "old standard tt", "homemade apple", "dancing script",
    "great vibes", "pacifico", "lobster", "abril fatface", "noto serif",
    "noto sans", "pt sans", "pt serif", "libre baskerville",
    "cormorant garamond",
}


def is_google_font(family: str) -> bool:
    return family.strip().lower() in _GOOGLE_FONTS


_UNICODE_SUBS = str.maketrans({
    '\u2018': "'", '\u2019': "'",
    '\u201c': '"', '\u201d': '"',
    '\u2013': '-', '\u2014': '--',
    '\u2026': '...', '\u00ae': '(R)', '\u00a9': '(C)',
    '\u00b0': 'deg', '\u2122': '(TM)',
    '\u00e9': 'e', '\u00e8': 'e',
    '\u00e0': 'a', '\u00e2': 'a',
    '\u00f1': 'n', '\u00fc': 'u', '\u00e4': 'a',
})

def sanitize_text(text: str) -> str:
    return text.translate(_UNICODE_SUBS)


def xml_escape_attr(text: str) -> str:
    """Escape text for use inside an XML attribute value (quoted with \")."""
    return (
        text.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;")
    )

def xml_escape(text: str) -> str:
    """Escape text for use as XML element content (between tags).
    Double quotes do NOT need escaping here; CorelDRAW renders &quot; literally."""
    return (
        text.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
    )


# ── Band path extraction ──────────────────────────────────────────────────────

def _load_band_commands() -> tuple[list, float]:
    """
    Parse the band outline from the Illustrator template into a list of
    (svg_op, [float_args]) tuples and the stroke width in points.

    Template path coords are already in points, Y-down — no conversion needed.
    """
    with open(TEMPLATE_PATH, "rb") as f:
        raw = f.read()
    if raw[:4] == b"\xc5\xd0\xd3\xc6":
        ps_offset = struct.unpack_from("<I", raw, 4)[0]
        ps_length = struct.unpack_from("<I", raw, 8)[0]
        ps = raw[ps_offset : ps_offset + ps_length].decode("latin-1")
    else:
        ps = raw.decode("latin-1")

    page_start = ps.find("%%Page:")
    idx_sop    = ps.find("false sop", page_start)
    chunk      = ps[page_start:idx_sop]
    idx_clp    = chunk.rfind("clp\r\n")
    drawing    = chunk[idx_clp + 5:].strip()

    commands = []
    stroke_w  = BAND_STROKE_WIDTH

    for raw_line in drawing.splitlines():
        line = raw_line.strip()
        if not line or line.startswith("%"):
            continue
        tokens = line.split()
        if not tokens:
            continue
        op   = tokens[-1]
        args = tokens[:-1]

        if op == "lw":
            stroke_w = float(args[0])
        elif op == "mo":
            commands.append(("M", [float(args[0]), float(args[1])]))
        elif op == "li":
            commands.append(("L", [float(args[0]), float(args[1])]))
        elif op == "cv":
            commands.append(("C", [float(a) for a in args]))
        elif op == "cp":
            commands.append(("Z", []))
        # lc, lj, ml, dsh, sadj → not needed in SVG

    return commands, stroke_w


def _band_path_d(commands: list, ox: float, oy: float) -> str:
    """Render band path offset by (ox, oy) as an SVG path d-string."""
    parts = []
    for op, args in commands:
        if op == "M":
            parts.append(f"M {args[0]+ox:.3f} {args[1]+oy:.3f}")
        elif op == "L":
            parts.append(f"L {args[0]+ox:.3f} {args[1]+oy:.3f}")
        elif op == "C":
            parts.append(
                f"C {args[0]+ox:.3f} {args[1]+oy:.3f} "
                f"{args[2]+ox:.3f} {args[3]+oy:.3f} "
                f"{args[4]+ox:.3f} {args[5]+oy:.3f}"
            )
        elif op == "Z":
            parts.append("Z")
    return " ".join(parts)


# ── Individual SVG writer ─────────────────────────────────────────────────────

def write_individual_svg(output_path: str, text_lines: list[str], font_name: str) -> None:
    """
    Write a single-band SVG file for one gavel order item.
    Physical size matches the band: 6.878" × 0.954" (495.25 × 68.65 pt).
    Uses the same absolute-coordinate, one-<text>-per-line approach as the
    bulk layout so it opens correctly in both Illustrator and CorelDRAW.
    """
    band_cmds, stroke_w = _load_band_commands()

    n          = len(text_lines)
    fs         = FONT_SIZE_PT.get(n, 10.0)
    line_h     = fs * LEADING
    block_h    = (n - 1) * line_h + fs
    baseline_y = BAND_CY - block_h / 2 + fs * 0.75   # first baseline (relative to band)
    cx         = BAND_CX                               # horizontal center

    google_families = [font_name] if is_google_font(font_name) else []

    out = []
    out.append('<?xml version="1.0" encoding="UTF-8"?>')
    out.append('<svg xmlns="http://www.w3.org/2000/svg"')
    out.append('     xmlns:xlink="http://www.w3.org/1999/xlink"')
    out.append(f'     width="{BAND_W/PT:.6f}in" height="{BAND_H/PT:.6f}in"')
    out.append(f'     viewBox="0 0 {BAND_W:.3f} {BAND_H:.3f}">')

    if google_families:
        fam_param = google_families[0].replace(" ", "+")
        out.append('  <defs><style><![CDATA[')
        out.append(f"    @import url('https://fonts.googleapis.com/css2?family={fam_param}&display=swap');")
        out.append('  ]]></style></defs>')

    # Band outline at origin (0, 0)
    d = _band_path_d(band_cmds, 0.0, 0.0)
    out.append(
        f'  <path d="{d}"'
        f' fill="none" stroke="{BAND_STROKE_COLOR}"'
        f' stroke-width="{stroke_w:.3f}"/>'
    )

    # Text lines
    ff = xml_escape_attr(font_name) + ", Helvetica, Arial, sans-serif"
    for i, tline in enumerate(text_lines):
        safe  = xml_escape(sanitize_text(tline))
        y_abs = baseline_y + i * line_h
        out.append(
            f'  <text'
            f' x="{cx:.3f}" y="{y_abs:.3f}"'
            f' text-anchor="middle"'
            f' font-family="{ff}"'
            f' font-size="{fs:.3f}"'
            f' fill="#000000">{safe}</text>'
        )

    out.append('</svg>')
    Path(output_path).write_text("\n".join(out), encoding="utf-8")


# ── Trello helpers ────────────────────────────────────────────────────────────

def _trello_get(path: str, **params) -> object:
    base_params = {"key": TRELLO_API_KEY, "token": TRELLO_TOKEN}
    base_params.update(params)
    url = "https://api.trello.com/1" + path + "?" + "&".join(f"{k}={urllib.request.quote(str(v))}" for k, v in base_params.items())
    req = urllib.request.Request(url, headers={"Accept": "application/json"})
    with urllib.request.urlopen(req, timeout=15) as r:
        return json.loads(r.read())


def _trello_post(path: str, data: dict) -> dict:
    base = {"key": TRELLO_API_KEY, "token": TRELLO_TOKEN}
    base.update(data)
    body = "&".join(f"{k}={urllib.request.quote(str(v), safe='')}" for k, v in base.items())
    url  = "https://api.trello.com/1" + path
    req  = urllib.request.Request(
        url,
        data=body.encode("utf-8"),
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=15) as r:
        return json.loads(r.read())


def trello_create_gavel_card(order_numbers: list[str]) -> tuple[str, str]:
    """
    Create a Trello card on the 'customs' board under 'Test List'.
    Title  : Auto Gavels {DATE}
    Desc   : order numbers, one per line
    Returns: (card_url, card_id)
    """
    if not TRELLO_API_KEY or TRELLO_API_KEY == "YOUR_TRELLO_API_KEY":
        print("  ⚠  Trello credentials not set — skipping card creation.")
        return "", ""

    date_str  = datetime.now().strftime("%m/%d/%Y")
    card_name = f"Auto Gavels {date_str}"
    desc      = "\n".join(order_numbers)

    # Find board
    boards = _trello_get("/members/me/boards", fields="name")
    board  = next((b for b in boards if b["name"].lower() == TRELLO_BOARD_NAME.lower()), None)
    if not board:
        raise ValueError(f"Trello board '{TRELLO_BOARD_NAME}' not found. "
                         f"Available: {[b['name'] for b in boards]}")

    # Find list on board
    lists       = _trello_get(f"/boards/{board['id']}/lists", fields="name")
    trello_list = next((l for l in lists if l["name"].lower() == TRELLO_LIST_NAME.lower()), None)
    if not trello_list:
        raise ValueError(f"List '{TRELLO_LIST_NAME}' not found on board '{TRELLO_BOARD_NAME}'. "
                         f"Available: {[l['name'] for l in lists]}")

    # Create card
    card = _trello_post("/cards", {
        "idList": trello_list["id"],
        "name":   card_name,
        "desc":   desc,
    })
    return card.get("url", ""), card.get("id", "")


def trello_attach_svg(card_id: str, file_path: str) -> None:
    """Upload an SVG file as an attachment to a Trello card."""
    filename = Path(file_path).name
    boundary = b"TrelloBoundaryXx7zA9qPmN"

    with open(file_path, "rb") as f:
        file_data = f.read()

    body = b""
    for field_name, value in [("key", TRELLO_API_KEY), ("token", TRELLO_TOKEN), ("name", filename)]:
        body += b"--" + boundary + b"\r\n"
        body += f'Content-Disposition: form-data; name="{field_name}"\r\n\r\n'.encode()
        body += value.encode() + b"\r\n"

    body += b"--" + boundary + b"\r\n"
    body += f'Content-Disposition: form-data; name="file"; filename="{filename}"\r\n'.encode()
    body += b"Content-Type: image/svg+xml\r\n\r\n"
    body += file_data + b"\r\n"
    body += b"--" + boundary + b"--\r\n"

    url = f"https://api.trello.com/1/cards/{card_id}/attachments"
    req = urllib.request.Request(
        url, data=body,
        headers={"Content-Type": f"multipart/form-data; boundary={boundary.decode()}"},
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=120) as r:
        r.read()


# ── SVG layout builder ────────────────────────────────────────────────────────

def build_layout_svg(items: list[dict]) -> str:
    """
    Arrange all gavel items on 24" × 12" pages (3 cols × 10 rows, 0.25" gap).

    All coordinates are absolute points (no transforms, no <use>, no <tspan>).
    Each line of text gets its own <text> element so CorelDRAW's SVG importer
    cannot misinterpret inherited text-anchor or tspan dy values.
    """
    band_cmds, stroke_w = _load_band_commands()

    items_per_page = COLS * ROWS
    num_pages      = max(1, (len(items) + items_per_page - 1) // items_per_page)
    total_h        = num_pages * PAGE_H

    # Center the grid on each page
    grid_w   = COLS * BAND_W + (COLS - 1) * GAP
    grid_h   = ROWS * BAND_H + (ROWS - 1) * GAP
    margin_x = (PAGE_W - grid_w) / 2
    margin_y = (PAGE_H - grid_h) / 2

    # Google Fonts @import
    google_families = sorted({
        item["font"] for item in items if is_google_font(item["font"])
    })

    out = []
    out.append('<?xml version="1.0" encoding="UTF-8"?>')
    out.append('<svg xmlns="http://www.w3.org/2000/svg"')
    out.append('     xmlns:xlink="http://www.w3.org/1999/xlink"')
    out.append(f'     width="{24}in" height="{num_pages * 12}in"')
    out.append(f'     viewBox="0 0 {PAGE_W:.3f} {total_h:.3f}">')

    if google_families:
        fam_param = "&family=".join(f.replace(" ", "+") for f in google_families)
        out.append('  <defs><style><![CDATA[')
        out.append(
            f"    @import url('https://fonts.googleapis.com/css2?family={fam_param}&display=swap');"
        )
        out.append('  ]]></style></defs>')

    # Page-break guide lines
    for p in range(1, num_pages):
        y = p * PAGE_H
        out.append(
            f'  <line x1="0" y1="{y:.3f}" x2="{PAGE_W:.3f}" y2="{y:.3f}"'
            f' stroke="#aaaaaa" stroke-width="1" stroke-dasharray="9 4.5"/>'
        )

    # Place each gavel item using fully absolute coordinates — no transforms.
    # One <path> per band outline, one <text> per text line.
    # text-anchor="middle" lives directly on each <text> with a matching x=cx.
    for idx, item in enumerate(items):
        page = idx // items_per_page
        slot = idx %  items_per_page
        col  = slot %  COLS
        row  = slot // COLS

        # Top-left of this band in absolute page coordinates (points)
        bx = margin_x + col * (BAND_W + GAP)
        by = page * PAGE_H + margin_y + row * (BAND_H + GAP)

        # Horizontal center (absolute)
        cx = bx + BAND_CX

        # Font metrics
        text_lines = item["lines"]
        font_name  = item["font"]
        n          = len(text_lines)
        fs         = FONT_SIZE_PT.get(n, 10.0)
        line_h     = fs * LEADING
        block_h    = (n - 1) * line_h + fs
        # Baseline of first line — centers text block vertically on band
        baseline_y = by + BAND_CY - block_h / 2 + fs * 0.75

        # ── band outline (inlined, absolute coordinates) ──────────────────
        d = _band_path_d(band_cmds, bx, by)
        out.append(
            f'  <path d="{d}"'
            f' fill="none" stroke="{BAND_STROKE_COLOR}"'
            f' stroke-width="{stroke_w:.3f}"/>'
        )

        # ── text lines (one <text> per line for CorelDRAW compatibility) ──
        ff = xml_escape_attr(font_name) + ", Helvetica, Arial, sans-serif"
        for i, tline in enumerate(text_lines):
            safe  = xml_escape(sanitize_text(tline))
            y_abs = baseline_y + i * line_h
            out.append(
                f'  <text'
                f' x="{cx:.3f}" y="{y_abs:.3f}"'
                f' text-anchor="middle"'
                f' font-family="{ff}"'
                f' font-size="{fs:.3f}"'
                f' fill="#000000">{safe}</text>'
            )

    out.append('</svg>')
    return "\n".join(out)


# ── ShipStation helpers ───────────────────────────────────────────────────────

def ss_get(path: str, params: dict) -> dict:
    url = f"{BASE_URL}{path}?" + "&".join(f"{k}={v}" for k, v in params.items())
    req = urllib.request.Request(url, headers={"api-key": API_KEY})
    with urllib.request.urlopen(req) as r:
        return json.loads(r.read())


def is_gavel(sku: str) -> bool:
    return any(p in sku for p in GAVEL_SKU_PATTERNS)


def fetch_gavel_shipments(days: int) -> list[dict]:
    cutoff = (datetime.now(timezone.utc) - timedelta(days=days)).strftime("%Y-%m-%dT%H:%M:%SZ")
    params = {
        "page_size": PAGE_SIZE,
        "sort_dir": "desc",
        "sort_by": "modified_at",
        "modified_at_start": cutoff,
        "shipment_status": "pending",
    }
    page, total_pages, results = 1, 1, []
    while page <= total_pages:
        params["page"] = page
        data = ss_get("/shipments", params)
        for s in data.get("shipments", []):
            for item in s.get("items", []):
                if is_gavel(item.get("sku", "")) and any(
                    o.get("name") == "CustomizedURL" for o in item.get("options", [])
                ):
                    results.append(s)
                    break
        total = data.get("total", 0)
        total_pages = (total + PAGE_SIZE - 1) // PAGE_SIZE
        print(f"  Page {page}/{total_pages} — {len(results)} gavel orders so far")
        page += 1
    return results


# ── Amazon customization helpers ──────────────────────────────────────────────

def fetch_customization(url: str) -> dict:
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=20) as r:
        content = r.read()
    z = zipfile.ZipFile(io.BytesIO(content))
    for name in z.namelist():
        if name.endswith(".json"):
            return json.loads(z.read(name))
    return {}


def extract_gavel_text(cust_json: dict) -> tuple[list[str], str]:
    """Return (non_empty_lines, font_name)."""
    lines, font = [], "Arial"
    for surf in (
        cust_json.get("version3.0", {})
        .get("customizationInfo", {})
        .get("surfaces", [])
    ):
        for area in surf.get("areas", []):
            text = (area.get("text") or "").strip()
            if area.get("fontFamily") and font == "Arial":
                font = area["fontFamily"]
            if text:
                lines.append(text)

    if not lines:
        def walk(node):
            nonlocal font
            if isinstance(node, dict):
                if node.get("type") == "TextCustomization":
                    v = (node.get("inputValue") or "").strip()
                    if v:
                        lines.append(v)
                if node.get("type") == "FontCustomization":
                    fam = node.get("fontSelection", {}).get("family")
                    if fam:
                        font = fam
                for v in node.values():
                    walk(v)
            elif isinstance(node, list):
                for item in node:
                    walk(item)
        walk(cust_json.get("customizationData", {}))

    return lines, font


# ── main ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Generate SVG gavel band layouts")
    parser.add_argument("--days",   type=int, default=1,  help="Days to look back (default 1)")
    parser.add_argument("--output", default="gavel_eps",  help="Output directory")
    args = parser.parse_args()

    out_dir = Path(args.output)
    out_dir.mkdir(parents=True, exist_ok=True)

    print(f"\n{'='*60}")
    print("Gavel Band SVG Layout Generator")
    print(f"Look-back: {args.days} day(s)  |  Output: {out_dir.resolve()}")
    print(f"{'='*60}\n")

    print("Fetching gavel orders from ShipStation...")
    shipments = fetch_gavel_shipments(days=args.days)
    print(f"\nFound {len(shipments)} gavel shipment(s)\n")

    if not shipments:
        print("Nothing to do.")
        return

    all_items    = []
    summary      = []
    order_nums   = []   # unique order numbers for Trello description
    ok           = 0
    errors       = 0

    for idx, ship in enumerate(shipments, 1):
        order_num = ship.get("shipment_number", f"ship_{idx}")
        customer  = (ship.get("ship_to") or {}).get("name", "")
        print(f"[{idx}/{len(shipments)}] {order_num} — {customer}")

        gavel_items = [
            item for item in ship.get("items", [])
            if is_gavel(item.get("sku", "")) and any(
                o.get("name") == "CustomizedURL" for o in item.get("options", [])
            )
        ]

        order_had_success = False

        for item_idx, item in enumerate(gavel_items, 1):
            sku = item.get("sku", "NOSKU")
            qty = item.get("quantity", 1)
            url = next(
                o["value"] for o in item.get("options", [])
                if o.get("name") == "CustomizedURL"
            )

            safe_order = order_num.replace("/", "-").replace("\\", "-")
            safe_sku   = sku.replace("/", "-").replace("\\", "-")
            svg_name   = f"{safe_order}_{item_idx}_{safe_sku}.svg"
            svg_path_i = str(out_dir / svg_name)

            print(f"   Item {item_idx}: {sku}  (qty {qty})")
            print(f"     Fetching customization... ", end="", flush=True)

            try:
                cust_json        = fetch_customization(url)
                text_lines, font = extract_gavel_text(cust_json)

                if not text_lines:
                    raise ValueError("No text found in customization data")

                print(f"OK  ({len(text_lines)} lines, font={font})")
                for i, ln in enumerate(text_lines, 1):
                    print(f"       L{i}: {ln}")

                # Individual SVG file
                write_individual_svg(svg_path_i, text_lines, font)
                print(f"     Individual SVG saved → {svg_name}")
                ok += 1
                order_had_success = True

                all_items.append({
                    "order_number": order_num,
                    "customer":     customer,
                    "sku":          sku,
                    "qty":          qty,
                    "font":         font,
                    "lines":        text_lines,
                })
                summary.append({
                    "order_number": order_num,
                    "customer":     customer,
                    "sku":          sku,
                    "qty":          qty,
                    "font":         font,
                    "lines":        " | ".join(text_lines),
                    "svg_file":     svg_name,
                    "status":       "ok",
                })

            except Exception as e:
                print(f"ERROR: {e}")
                errors += 1
                summary.append({
                    "order_number": order_num,
                    "customer":     customer,
                    "sku":          sku,
                    "qty":          item.get("quantity", 1),
                    "font":         "",
                    "lines":        "",
                    "svg_file":     "",
                    "status":       f"error: {e}",
                })

        if order_had_success and order_num not in order_nums:
            order_nums.append(order_num)

    # Build and write the bulk SVG layout file
    items_per_page = COLS * ROWS
    num_pages = max(1, (len(all_items) + items_per_page - 1) // items_per_page)
    print(f"\nBuilding bulk SVG layout: {len(all_items)} items across {num_pages} page(s)...")
    layout_path = out_dir / "gavel_layout.svg"
    layout_path.write_text(build_layout_svg(all_items), encoding="utf-8")
    print(f"Bulk layout saved → {layout_path.resolve()}")

    # Write summary CSV
    csv_path = out_dir / "summary.csv"
    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=["order_number", "customer", "sku", "qty", "font", "lines", "svg_file", "status"],
        )
        writer.writeheader()
        writer.writerows(summary)

    # Post Trello card and attach all SVGs
    print(f"\nPosting Trello card ({len(order_nums)} order numbers)...")
    try:
        card_url, card_id = trello_create_gavel_card(order_nums)
        if card_url:
            print(f"  Card created → {card_url}")
        if card_id:
            # Collect all SVGs: bulk layout first, then individual files
            svgs_to_upload = [layout_path] + sorted(out_dir.glob("*.svg"))
            # Remove duplicates (layout already in glob) while preserving order
            seen = set()
            upload_list = []
            for p in svgs_to_upload:
                if p not in seen:
                    seen.add(p)
                    upload_list.append(p)

            for i, svg in enumerate(upload_list, 1):
                print(f"  [{i}/{len(upload_list)}] Uploading {svg.name}...", end="", flush=True)
                trello_attach_svg(card_id, str(svg))
                print(" ✓")
    except Exception as e:
        print(f"  Trello error: {e}")

    print(f"\n{'='*60}")
    print(f"Done.  Individual SVGs: {ok}  Errors: {errors}")
    print(f"Bulk layout : {layout_path.resolve()}")
    print(f"Summary CSV : {csv_path.resolve()}")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()
