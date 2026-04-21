"""
shipstation_layouts.py — Pull custom orders from ShipStation and generate print-ready PDF layouts.

Usage:
    python shipstation_layouts.py [--days N] [--output DIR] [--status STATUS]

Options:
    --days N       Look back N days for orders (default: 1 = today)
    --output DIR   Output directory for PDFs (default: ./layouts)
    --status       Order status to pull: awaiting_shipment | awaiting_pickup | all (default: awaiting_shipment)

Output:
    One PDF per custom line item, named:  {OrderNumber}_{ItemIndex}_{SKU}.pdf
    A summary log file:                   layouts/summary.csv
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


def _load_env(path: str = ".env") -> None:
    env_file = Path(path)
    if not env_file.exists():
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

from PIL import Image
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas

# ── config ──────────────────────────────────────────────────────────────────

API_KEY    = os.environ.get("SHIPSTATION_API_KEY", "")
BASE_URL   = "https://api.shipstation.com/v2"
PAGE_SIZE  = 100

# ── ShipStation API ──────────────────────────────────────────────────────────

def ss_get(path: str, params: dict = None) -> dict:
    url = f"{BASE_URL}{path}"
    if params:
        qs = "&".join(f"{k}={v}" for k, v in params.items())
        url = f"{url}?{qs}"
    req = urllib.request.Request(url, headers={"api-key": API_KEY, "Content-Type": "application/json"})
    with urllib.request.urlopen(req) as r:
        return json.loads(r.read())


def fetch_custom_shipments(days: int, status: str) -> list[dict]:
    """Return all shipments with at least one item carrying a CustomizedURL."""
    cutoff = (datetime.now(timezone.utc) - timedelta(days=days)).strftime("%Y-%m-%dT%H:%M:%SZ")
    params = {
        "page_size": PAGE_SIZE,
        "sort_dir": "desc",
        "sort_by": "modified_at",
        "modified_at_start": cutoff,
    }
    if status != "all":
        params["shipment_status"] = status  # v2 values: pending, label_purchased

    page, total_pages = 1, 1
    custom = []

    while page <= total_pages:
        params["page"] = page
        data = ss_get("/shipments", params)
        shipments = data.get("shipments", [])

        for s in shipments:
            for item in s.get("items", []):
                if any(o.get("name") == "CustomizedURL" for o in item.get("options", [])):
                    custom.append(s)
                    break

        total = data.get("total", 0)
        total_pages = (total + PAGE_SIZE - 1) // PAGE_SIZE
        print(f"  Page {page}/{total_pages} — {len(shipments)} shipments scanned, {len(custom)} custom so far")
        page += 1

    return custom


# ── Amazon customization ZIP ─────────────────────────────────────────────────

def fetch_customization(url: str) -> dict:
    """Download the Amazon customization ZIP and return parsed data."""
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=20) as r:
        content = r.read()

    z = zipfile.ZipFile(io.BytesIO(content))
    result = {"preview_jpg": None, "svg": None, "json": None}

    for name in z.namelist():
        if name.lower().endswith(".jpg") or name.lower().endswith(".jpeg"):
            result["preview_jpg"] = z.read(name)
        elif name.lower().endswith(".svg"):
            result["svg"] = z.read(name).decode("utf-8", errors="ignore")
        elif name.lower().endswith(".json"):
            try:
                result["json"] = json.loads(z.read(name))
            except Exception:
                pass

    return result


def extract_text_lines(cust_json: dict) -> list[dict]:
    """Extract text fields from the Amazon v3 customization JSON."""
    lines = []
    try:
        surfaces = cust_json.get("version3.0", {}).get("customizationInfo", {}).get("surfaces", [])
        for surface in surfaces:
            for area in surface.get("areas", []):
                if area.get("text"):
                    lines.append({
                        "label":  area.get("label") or area.get("name", "Text"),
                        "text":   area["text"],
                        "font":   area.get("fontFamily", "Helvetica"),
                        "color":  area.get("fill", "#000000"),
                    })
    except Exception:
        pass

    # Fallback: walk customizationData tree
    if not lines:
        def walk(node):
            if isinstance(node, dict):
                if node.get("type") == "TextCustomization" and node.get("inputValue"):
                    lines.append({
                        "label": node.get("label", "Text"),
                        "text":  node["inputValue"],
                        "font":  "Helvetica",
                        "color": "#000000",
                    })
                for v in node.values():
                    walk(v)
            elif isinstance(node, list):
                for item in node:
                    walk(item)
        walk(cust_json.get("customizationData", {}))

    return lines


# ── PDF layout ───────────────────────────────────────────────────────────────

BRAND_COLOR  = colors.HexColor("#1A2E4A")   # navy
ACCENT_COLOR = colors.HexColor("#E8701A")   # orange
LIGHT_GRAY   = colors.HexColor("#F5F5F5")
MED_GRAY     = colors.HexColor("#CCCCCC")


def draw_pdf(output_path: str, shipment: dict, item: dict, cust: dict):
    """Render a single-page print-ready production layout PDF."""
    W, H = letter  # 8.5 x 11 inches
    c = canvas.Canvas(output_path, pagesize=letter)

    order_num   = shipment.get("shipment_number", "")
    ship_to     = shipment.get("ship_to", {})
    customer    = ship_to.get("name", "")
    ship_date   = (shipment.get("ship_by_date") or shipment.get("ship_date") or "")[:10]
    sku         = item.get("sku", "")
    product     = item.get("name", "")
    qty         = item.get("quantity", 1)
    text_lines  = extract_text_lines(cust["json"]) if cust.get("json") else []
    preview_jpg = cust.get("preview_jpg")

    # ── header bar ──
    c.setFillColor(BRAND_COLOR)
    c.rect(0, H - 1.1 * inch, W, 1.1 * inch, fill=1, stroke=0)

    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 18)
    c.drawString(0.35 * inch, H - 0.55 * inch, "Pacific Sign and Stamp")
    c.setFont("Helvetica", 10)
    c.drawRightString(W - 0.35 * inch, H - 0.45 * inch, f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}")

    # ── order info block ──
    y = H - 1.5 * inch
    c.setFillColor(LIGHT_GRAY)
    c.rect(0.35 * inch, y - 0.85 * inch, W - 0.7 * inch, 0.9 * inch, fill=1, stroke=0)

    c.setFillColor(BRAND_COLOR)
    c.setFont("Helvetica-Bold", 11)
    c.drawString(0.5 * inch, y - 0.25 * inch, f"Order #:  {order_num}")
    c.drawString(0.5 * inch, y - 0.5 * inch,  f"Customer: {customer}")

    c.setFont("Helvetica", 11)
    c.drawString(3.5 * inch, y - 0.25 * inch, f"Ship By: {ship_date or '—'}")
    c.drawString(3.5 * inch, y - 0.5 * inch,  f"Qty:     {qty}")

    # ── accent divider ──
    c.setStrokeColor(ACCENT_COLOR)
    c.setLineWidth(2)
    c.line(0.35 * inch, y - 0.95 * inch, W - 0.35 * inch, y - 0.95 * inch)

    # ── product section ──
    y -= 1.15 * inch
    c.setFillColor(BRAND_COLOR)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(0.35 * inch, y, "PRODUCT")
    c.setFont("Helvetica", 10)
    c.setFillColor(colors.black)
    # Wrap long product name
    max_w = W - 0.7 * inch
    words, line, lines_out = product.split(), "", []
    for w in words:
        test = (line + " " + w).strip()
        if c.stringWidth(test, "Helvetica", 10) < max_w:
            line = test
        else:
            lines_out.append(line)
            line = w
    lines_out.append(line)
    for i, ln in enumerate(lines_out[:3]):
        c.drawString(0.35 * inch, y - 0.18 * inch * (i + 1), ln)

    c.setFillColor(MED_GRAY)
    c.setFont("Helvetica", 9)
    c.drawString(0.35 * inch, y - 0.55 * inch, f"SKU: {sku}")

    # ── preview image (right column) ──
    img_x = W - 3.0 * inch
    img_y = y - 2.4 * inch
    img_size = 2.6 * inch

    if preview_jpg:
        try:
            pil_img = Image.open(io.BytesIO(preview_jpg)).convert("RGB")
            img_buf = io.BytesIO()
            pil_img.save(img_buf, format="JPEG")
            img_buf.seek(0)
            ir = ImageReader(img_buf)
            c.drawImage(ir, img_x, img_y, width=img_size, height=img_size, preserveAspectRatio=True)
            # border
            c.setStrokeColor(MED_GRAY)
            c.setLineWidth(1)
            c.rect(img_x, img_y, img_size, img_size, fill=0, stroke=1)
            c.setFillColor(BRAND_COLOR)
            c.setFont("Helvetica-Bold", 8)
            c.drawCentredString(img_x + img_size / 2, img_y - 0.15 * inch, "CUSTOMER PREVIEW")
        except Exception as e:
            c.setFillColor(colors.red)
            c.setFont("Helvetica", 8)
            c.drawString(img_x, img_y + img_size / 2, f"[Preview error: {e}]")

    # ── customization text lines ──
    y -= 0.7 * inch
    c.setFillColor(BRAND_COLOR)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(0.35 * inch, y, "CUSTOMIZATION")

    c.setStrokeColor(BRAND_COLOR)
    c.setLineWidth(0.5)
    c.line(0.35 * inch, y - 0.05 * inch, img_x - 0.2 * inch, y - 0.05 * inch)

    text_right_edge = img_x - 0.25 * inch
    ty = y - 0.25 * inch

    if text_lines:
        for tl in text_lines:
            if ty < 0.8 * inch:
                break
            label = tl["label"]
            text  = tl["text"]
            font  = tl.get("font", "Helvetica")

            c.setFillColor(colors.HexColor("#666666"))
            c.setFont("Helvetica", 8)
            c.drawString(0.35 * inch, ty, label.upper())

            c.setFillColor(colors.black)
            c.setFont("Helvetica-Bold", 13)
            # truncate if too long
            max_text_w = text_right_edge - 0.35 * inch
            display = text
            while c.stringWidth(display, "Helvetica-Bold", 13) > max_text_w and len(display) > 4:
                display = display[:-2] + "…"
            c.drawString(0.35 * inch, ty - 0.2 * inch, display)
            ty -= 0.52 * inch
    else:
        c.setFillColor(colors.HexColor("#999999"))
        c.setFont("Helvetica-Oblique", 10)
        c.drawString(0.35 * inch, ty, "No customization text extracted.")

    # ── buyer notes ──
    notes = (shipment.get("notes_from_buyer") or "").strip()
    if notes and ty > 1.5 * inch:
        c.setStrokeColor(ACCENT_COLOR)
        c.setLineWidth(2)
        c.rect(0.35 * inch, ty - 0.6 * inch, text_right_edge - 0.35 * inch, 0.65 * inch, fill=0, stroke=1)
        c.setFillColor(ACCENT_COLOR)
        c.setFont("Helvetica-Bold", 8)
        c.drawString(0.5 * inch, ty - 0.15 * inch, "BUYER NOTE:")
        c.setFillColor(colors.black)
        c.setFont("Helvetica", 9)
        c.drawString(0.5 * inch, ty - 0.35 * inch, notes[:120])
        ty -= 0.75 * inch

    # ── footer ──
    c.setFillColor(BRAND_COLOR)
    c.rect(0, 0, W, 0.35 * inch, fill=1, stroke=0)
    c.setFillColor(colors.white)
    c.setFont("Helvetica", 8)
    c.drawCentredString(W / 2, 0.12 * inch, f"Order {order_num}  •  SKU {sku}  •  Qty {qty}")

    c.save()


# ── main ─────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Generate print-ready layouts from ShipStation custom orders")
    parser.add_argument("--days",   type=int, default=1,  help="Days to look back (default 1)")
    parser.add_argument("--output", default="layouts",    help="Output directory (default: layouts)")
    parser.add_argument("--status", default="pending",
                        choices=["pending", "label_purchased", "all"],
                        help="Shipment status filter: pending=awaiting shipment (default: pending)")
    args = parser.parse_args()

    out_dir = Path(args.output)
    out_dir.mkdir(parents=True, exist_ok=True)

    print(f"\n{'='*60}")
    print(f"ShipStation Layout Generator")
    status_label = {"pending": "awaiting shipment", "label_purchased": "label purchased", "all": "all"}.get(args.status, args.status)
    print(f"Looking back {args.days} day(s)  |  Status: {status_label}")
    print(f"Output: {out_dir.resolve()}")
    print(f"{'='*60}\n")

    print("Fetching custom orders from ShipStation...")
    shipments = fetch_custom_shipments(days=args.days, status=args.status)
    print(f"\nFound {len(shipments)} shipment(s) with custom items\n")

    if not shipments:
        print("Nothing to do.")
        return

    summary_rows = []
    generated = 0
    errors    = 0

    for ship_idx, shipment in enumerate(shipments, 1):
        order_num = shipment.get("shipment_number", f"ship_{ship_idx}")
        customer  = shipment.get("ship_to", {}).get("name", "")
        print(f"[{ship_idx}/{len(shipments)}] Order {order_num} — {customer}")

        custom_items = [
            item for item in shipment.get("items", [])
            if any(o.get("name") == "CustomizedURL" for o in item.get("options", []))
        ]

        for item_idx, item in enumerate(custom_items, 1):
            sku = item.get("sku", "NOSKU")
            url = next(
                o["value"] for o in item.get("options", [])
                if o.get("name") == "CustomizedURL"
            )

            safe_order = order_num.replace("/", "-").replace("\\", "-")
            safe_sku   = sku.replace("/", "-").replace("\\", "-")
            pdf_name   = f"{safe_order}_{item_idx}_{safe_sku}.pdf"
            pdf_path   = out_dir / pdf_name

            print(f"   Item {item_idx}: {sku[:40]}")
            print(f"     Fetching customization... ", end="", flush=True)

            try:
                cust = fetch_customization(url)
                text_lines = extract_text_lines(cust["json"]) if cust.get("json") else []
                print(f"OK ({len(text_lines)} text field(s))")

                draw_pdf(str(pdf_path), shipment, item, cust)
                print(f"     PDF saved: {pdf_name}")
                generated += 1

                summary_rows.append({
                    "order_number": order_num,
                    "customer":     customer,
                    "sku":          sku,
                    "product":      item.get("name", "")[:80],
                    "qty":          item.get("quantity", 1),
                    "text_lines":   " | ".join(f"{t['label']}: {t['text']}" for t in text_lines),
                    "pdf_file":     pdf_name,
                    "status":       "ok",
                })

            except Exception as e:
                print(f"ERROR: {e}")
                errors += 1
                summary_rows.append({
                    "order_number": order_num,
                    "customer":     customer,
                    "sku":          sku,
                    "product":      item.get("name", "")[:80],
                    "qty":          item.get("quantity", 1),
                    "text_lines":   "",
                    "pdf_file":     "",
                    "status":       f"error: {e}",
                })

    # ── write summary CSV ──
    summary_path = out_dir / "summary.csv"
    with open(summary_path, "w", newline="", encoding="utf-8") as f:
        fieldnames = ["order_number", "customer", "sku", "product", "qty", "text_lines", "pdf_file", "status"]
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(summary_rows)

    print(f"\n{'='*60}")
    print(f"Done.  Generated: {generated}  Errors: {errors}")
    print(f"Summary: {summary_path.resolve()}")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()
