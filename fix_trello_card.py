"""
fix_trello_card.py — Remove all attachments from the Auto Gavels card and
re-upload only the correct files: gavel_layout.svg + the individual SVGs
listed in summary.csv.
"""

import csv
import json
import os
import sys
import urllib.request
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

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

TRELLO_API_KEY  = os.environ.get("TRELLO_API_KEY", "")
TRELLO_TOKEN    = os.environ.get("TRELLO_TOKEN", "")
TRELLO_BASE     = "https://api.trello.com/1"
BOARD_NAME      = "customs"
CARD_TITLE      = "Auto Gavels 04/21/2026"    # exact card title
GAVEL_EPS_DIR   = Path(__file__).parent / "gavel_eps"
SUMMARY_CSV     = GAVEL_EPS_DIR / "summary.csv"


def trello_get(path, **params):
    p = {"key": TRELLO_API_KEY, "token": TRELLO_TOKEN}
    p.update(params)
    url = TRELLO_BASE + path + "?" + "&".join(
        f"{k}={urllib.request.quote(str(v))}" for k, v in p.items()
    )
    req = urllib.request.Request(url, headers={"Accept": "application/json"})
    with urllib.request.urlopen(req, timeout=20) as r:
        return json.loads(r.read())


def trello_delete(path):
    p = {"key": TRELLO_API_KEY, "token": TRELLO_TOKEN}
    url = TRELLO_BASE + path + "?" + "&".join(
        f"{k}={urllib.request.quote(str(v))}" for k, v in p.items()
    )
    req = urllib.request.Request(url, method="DELETE")
    with urllib.request.urlopen(req, timeout=20) as r:
        return r.read()


def trello_put(path, data):
    """Update a Trello resource via PUT (form-encoded body)."""
    base = {"key": TRELLO_API_KEY, "token": TRELLO_TOKEN}
    base.update(data)
    body = "&".join(
        f"{k}={urllib.request.quote(str(v), safe='')}" for k, v in base.items()
    )
    url = TRELLO_BASE + path
    req = urllib.request.Request(
        url,
        data=body.encode("utf-8"),
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        method="PUT",
    )
    with urllib.request.urlopen(req, timeout=15) as r:
        return json.loads(r.read())


def trello_attach(card_id, file_path):
    filename = Path(file_path).name
    boundary = b"FixBoundaryXx7zA9qPmN"
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
    url = f"{TRELLO_BASE}/cards/{card_id}/attachments"
    req = urllib.request.Request(
        url, data=body,
        headers={"Content-Type": f"multipart/form-data; boundary={boundary.decode()}"},
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=120) as r:
        r.read()


def find_card(board_id, title_lower):
    """Return card dict whose name matches title_lower (case-insensitive)."""
    cards = trello_get(f"/boards/{board_id}/cards", fields="name")
    for c in cards:
        if c.get("name", "").lower() == title_lower:
            return c
    return None


def main():
    print("=" * 60)
    print("Trello Card Attachment Fixer")
    print("=" * 60)

    # 1. Find board
    boards = trello_get("/members/me/boards", fields="name")
    board  = next((b for b in boards if b["name"].lower() == BOARD_NAME.lower()), None)
    if not board:
        raise SystemExit(f"Board '{BOARD_NAME}' not found. Available: {[b['name'] for b in boards]}")
    print(f"Board : {board['name']}  ({board['id']})")

    # 2. Find card
    card = find_card(board["id"], CARD_TITLE.lower())
    if not card:
        # Try partial match
        all_cards = trello_get(f"/boards/{board['id']}/cards", fields="name")
        gavel_cards = [c for c in all_cards if "auto gavels" in c.get("name", "").lower()]
        if len(gavel_cards) == 1:
            card = gavel_cards[0]
        elif len(gavel_cards) > 1:
            print("Multiple Auto Gavels cards found:")
            for c in gavel_cards:
                print(f"  [{c['id']}] {c['name']}")
            raise SystemExit("Please set CARD_TITLE to the exact card name above.")
        else:
            raise SystemExit(f"No card matching '{CARD_TITLE}' found on board '{BOARD_NAME}'.")
    print(f"Card  : {card['name']}  ({card['id']})")
    card_id = card["id"]

    # 3. List current attachments
    attachments = trello_get(f"/cards/{card_id}/attachments")
    print(f"\nCurrent attachments: {len(attachments)}")

    # 4. Delete all existing attachments
    if attachments:
        print("Deleting all existing attachments...")
        for i, att in enumerate(attachments, 1):
            print(f"  [{i}/{len(attachments)}] Deleting {att.get('name', att['id'])}...", end="", flush=True)
            trello_delete(f"/cards/{card_id}/attachments/{att['id']}")
            print(" ✓")
        print("All attachments removed.")
    else:
        print("No attachments to remove.")

    # 5. Read summary.csv — build upload list and collect order numbers
    layout_svg  = GAVEL_EPS_DIR / "gavel_layout.svg"
    upload_list = []
    order_nums  = []   # unique order numbers from this run

    if layout_svg.exists():
        upload_list.append(layout_svg)
    else:
        print(f"WARNING: layout file not found: {layout_svg}")

    if SUMMARY_CSV.exists():
        with open(SUMMARY_CSV, newline="", encoding="utf-8") as f:
            for row in csv.DictReader(f):
                if row.get("status") == "ok" and row.get("svg_file"):
                    p = GAVEL_EPS_DIR / row["svg_file"]
                    if p.exists():
                        upload_list.append(p)
                    else:
                        print(f"  WARNING: SVG not found on disk — {row['svg_file']}")
                    onum = row.get("order_number", "").strip()
                    if onum and onum not in order_nums:
                        order_nums.append(onum)
    else:
        print(f"WARNING: summary.csv not found at {SUMMARY_CSV}")

    # Deduplicate upload list while preserving order
    seen = set()
    deduped = []
    for p in upload_list:
        if p not in seen:
            seen.add(p)
            deduped.append(p)
    upload_list = deduped

    # 6. Update card description to match summary.csv order numbers
    if order_nums:
        new_desc = "\n".join(order_nums)
        print(f"\nUpdating card description ({len(order_nums)} order number(s))...")
        trello_put(f"/cards/{card_id}", {"desc": new_desc})
        print("  Description updated ✓")
    else:
        print("\nWARNING: No order numbers found in summary.csv — description not updated.")

    # 7. Upload corrected attachments
    print(f"\nUploading {len(upload_list)} file(s)...")
    for i, svg in enumerate(upload_list, 1):
        print(f"  [{i}/{len(upload_list)}] {svg.name}...", end="", flush=True)
        trello_attach(card_id, str(svg))
        print(" ✓")

    print(f"\nDone. Card now has {len(upload_list)} attachment(s).")
    print(f"Card URL: https://trello.com/c/{card_id}")


if __name__ == "__main__":
    main()
