"""
gavel_cdr_generator.py — Generate CorelDRAW CDR layouts for custom gavel band orders.

Usage:
    python gavel_cdr_generator.py [--days N] [--output DIR] [--visible]

Options:
    --days N      Look back N days for pending orders (default: 1)
    --output DIR  Output folder for CDR files (default: ./gavel_cdrs)
    --visible     Show CorelDRAW window while processing (default: hidden)

Requirements:
    CorelDRAW 2025 (v27) must be installed. Run: pip install pywin32 requests

Output:
    {order_number}_{item_index}_{sku}.cdr   — one CDR per custom gavel line item
    gavel_cdrs/summary.csv                  — run log
"""

import argparse
import csv
import io
import json
import os
import sys
import urllib.request
import zipfile
from datetime import datetime, timedelta, timezone
from pathlib import Path

import win32com.client

# ── config ───────────────────────────────────────────────────────────────────

API_KEY        = "cfaAl+ea1wj7jmWab9bbh2VSS+WBRC+Np+62K/0dT5g"
BASE_URL       = "https://api.shipstation.com/v2"
PAGE_SIZE      = 100
COREL_PROGID   = "CorelDRAW.Application.27"
TEMPLATE_PATH  = r"C:\Users\breez\Downloads\gavelband_template.cdr"

# SKU substrings that identify gavel band products
GAVEL_SKU_PATTERNS = ["CGVL", "GVLSB", "GFCSTM-GVL", "GF-VLU"]

# Font size by non-empty line count
FONT_SIZE_MAP = {1: 15.0, 2: 13.0, 3: 11.0, 4: 10.0}

# CorelDRAW alignment constants
CDR_CENTER_ALIGN = 2

# ── ShipStation helpers ───────────────────────────────────────────────────────

def ss_get(path: str, params: dict = None) -> dict:
    url = f"{BASE_URL}{path}"
    if params:
        url += "?" + "&".join(f"{k}={v}" for k, v in params.items())
    req = urllib.request.Request(url, headers={"api-key": API_KEY})
    with urllib.request.urlopen(req) as r:
        return json.loads(r.read())


def is_gavel_sku(sku: str) -> bool:
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
    page, total_pages = 1, 1
    results = []

    while page <= total_pages:
        params["page"] = page
        data = ss_get("/shipments", params)
        for s in data.get("shipments", []):
            for item in s.get("items", []):
                if is_gavel_sku(item.get("sku", "")) and any(
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
    """Return (non_empty_lines, font_name) from customization JSON."""
    lines = []
    font = "Times New Roman"

    surfaces = (
        cust_json.get("version3.0", {})
        .get("customizationInfo", {})
        .get("surfaces", [])
    )
    for surf in surfaces:
        for area in surf.get("areas", []):
            text = (area.get("text") or "").strip()
            if area.get("fontFamily") and font == "Times New Roman":
                font = area["fontFamily"]
            if text:
                lines.append(text)

    # Fallback: walk customizationData tree
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


# ── CDR generation ────────────────────────────────────────────────────────────

_corel_app = None

def get_corel_app(visible: bool) -> object:
    global _corel_app
    if _corel_app is None:
        _corel_app = win32com.client.Dispatch(COREL_PROGID)
        _corel_app.Visible = visible
    return _corel_app


def generate_cdr(
    output_path: str,
    text_lines: list[str],
    font_name: str,
    visible: bool = False,
) -> None:
    """Open the gavel band template, add centered text, save as output_path."""
    app = get_corel_app(visible)

    doc = app.OpenDocument(TEMPLATE_PATH)
    try:
        page = doc.Pages.First
        pw = page.SizeWidth    # 6.875"
        ph = page.SizeHeight   # 0.95"
        layer = page.ActiveLayer

        n = len(text_lines)
        font_size = FONT_SIZE_MAP.get(n, 10.0)

        # Horizontal margin keeps text off the edges of the band
        margin_x = 0.08
        margin_y = 0.02

        # CreateParagraphText(x1, y1, x2, y2)  — diagonal corners
        # x1,y1 = top-left;  x2,y2 = bottom-right (doc coords: y=0 is bottom)
        x1 = margin_x
        y1 = ph - margin_y
        x2 = pw - margin_x
        y2 = margin_y

        shape = layer.CreateParagraphText(x1, y1, x2, y2)
        story = shape.Text.Story

        story.WideText = "\r".join(text_lines)
        story.Font      = font_name
        story.Size      = font_size
        story.Alignment = CDR_CENTER_ALIGN

        # Vertical centering: move shape so text block is centered on page
        shape.PositionX = pw / 2
        shape.PositionY = ph / 2

        doc.SaveAs(output_path, 0)   # 0 = CDR format
    finally:
        doc.Close()


# ── main ─────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Generate CDR layouts for gavel band orders")
    parser.add_argument("--days",    type=int, default=1,       help="Days to look back (default 1)")
    parser.add_argument("--output",  default="gavel_cdrs",      help="Output directory")
    parser.add_argument("--visible", action="store_true",        help="Show CorelDRAW window")
    args = parser.parse_args()

    out_dir = Path(args.output)
    out_dir.mkdir(parents=True, exist_ok=True)

    print(f"\n{'='*60}")
    print("Gavel Band CDR Generator")
    print(f"Look-back: {args.days} day(s)  |  Output: {out_dir.resolve()}")
    print(f"{'='*60}\n")

    print("Fetching gavel orders from ShipStation...")
    shipments = fetch_gavel_shipments(days=args.days)
    print(f"\nFound {len(shipments)} gavel shipment(s)\n")

    if not shipments:
        print("Nothing to do.")
        return

    summary = []
    ok = errors = 0

    for idx, ship in enumerate(shipments, 1):
        order_num = ship.get("shipment_number", f"ship_{idx}")
        customer  = ship.get("ship_to", {}).get("name", "")
        print(f"[{idx}/{len(shipments)}] {order_num} — {customer}")

        gavel_items = [
            item for item in ship.get("items", [])
            if is_gavel_sku(item.get("sku", "")) and any(
                o.get("name") == "CustomizedURL" for o in item.get("options", [])
            )
        ]

        for item_idx, item in enumerate(gavel_items, 1):
            sku = item.get("sku", "NOSKU")
            url = next(
                o["value"] for o in item.get("options", [])
                if o.get("name") == "CustomizedURL"
            )

            safe_order = order_num.replace("/", "-").replace("\\", "-")
            safe_sku   = sku.replace("/", "-").replace("\\", "-")
            cdr_name   = f"{safe_order}_{item_idx}_{safe_sku}.cdr"
            cdr_path   = str(out_dir / cdr_name)

            print(f"   Item {item_idx}: {sku}")
            print(f"     Fetching customization... ", end="", flush=True)

            try:
                cust_json         = fetch_customization(url)
                text_lines, font  = extract_gavel_text(cust_json)

                if not text_lines:
                    raise ValueError("No text found in customization data")

                print(f"OK  ({len(text_lines)} lines, font={font})")
                for i, ln in enumerate(text_lines, 1):
                    print(f"       L{i}: {ln}")

                print(f"     Generating CDR... ", end="", flush=True)
                generate_cdr(cdr_path, text_lines, font, visible=args.visible)
                print(f"saved → {cdr_name}")
                ok += 1

                summary.append({
                    "order_number": order_num,
                    "customer":     customer,
                    "sku":          sku,
                    "font":         font,
                    "lines":        " | ".join(text_lines),
                    "cdr_file":     cdr_name,
                    "status":       "ok",
                })

            except Exception as e:
                print(f"ERROR: {e}")
                errors += 1
                summary.append({
                    "order_number": order_num,
                    "customer":     customer,
                    "sku":          sku,
                    "font":         "",
                    "lines":        "",
                    "cdr_file":     "",
                    "status":       f"error: {e}",
                })

    # ── summary CSV ──
    csv_path = out_dir / "summary.csv"
    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["order_number","customer","sku","font","lines","cdr_file","status"])
        writer.writeheader()
        writer.writerows(summary)

    print(f"\n{'='*60}")
    print(f"Done.  Generated: {ok}  Errors: {errors}")
    print(f"Summary: {csv_path.resolve()}")
    print(f"{'='*60}\n")

    # Quit CorelDRAW if we opened it
    global _corel_app
    if _corel_app:
        try:
            _corel_app.Quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()
