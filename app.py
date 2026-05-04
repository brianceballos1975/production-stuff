import os
import re
import sys
import subprocess
import threading
import uuid
from datetime import datetime, timezone
from pathlib import Path

import io, tempfile
from datetime import timedelta

from flask import Flask, jsonify, render_template, request
from flask_cors import CORS
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger
from google.cloud import firestore

GAVEL_ORDERS_COL = "gavel_orders"

app = Flask(__name__, template_folder="web_templates")
CORS(app, origins=[
    "https://sku-scanner-41063282006.us-central1.run.app",
    "https://file.allquality.com",
    "http://localhost:3000",   # local dev
])

# ── Firestore ──────────────────────────────────────────────────────────────────
_db_client = None

def db():
    global _db_client
    if _db_client is None:
        _db_client = firestore.Client()
    return _db_client

RUNS_COL  = "gavel_runs"
CFG_COL   = "gavel_config"
SCHED_DOC = "schedule"

# ── Scheduler ─────────────────────────────────────────────────────────────────
scheduler = BackgroundScheduler(timezone="UTC")
scheduler.start()

# ── Generator ─────────────────────────────────────────────────────────────────
SCRIPT = str(Path(__file__).parent / "gavel_eps_generator.py")

_run_lock     = threading.Lock()
_active_run_id: str | None = None


def _run_generator(run_id: str, extra_args: list[str]):
    global _active_run_id
    ref = db().collection(RUNS_COL).document(run_id)
    cmd = [sys.executable, SCRIPT] + extra_args
    log_lines: list[str] = []

    try:
        proc = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            bufsize=1,
            cwd=str(Path(__file__).parent),
        )
        for raw in proc.stdout:
            line = raw.rstrip("\n")
            log_lines.append(line)
            if len(log_lines) % 5 == 0:
                ref.update({"log": "\n".join(log_lines)})
        proc.wait()
        rc = proc.returncode
    except Exception as exc:
        log_lines.append(f"EXCEPTION: {exc}")
        rc = 1

    full_log = "\n".join(log_lines)

    svgs = errors = 0
    m = re.search(r"Individual SVGs:\s*(\d+)", full_log)
    if m:
        svgs = int(m.group(1))
    m = re.search(r"Errors:\s*(\d+)", full_log)
    if m:
        errors = int(m.group(1))

    # Extract unique order numbers processed this run
    order_numbers = list(dict.fromkeys(re.findall(r'\b\d{3}-\d{7}-\d{7}\b', full_log)))

    if rc != 0:
        status = "error"
    elif errors > 0:
        status = "warning"
    else:
        status = "success"

    ref.update({
        "log":           full_log,
        "status":        status,
        "svgs":          svgs,
        "errors":        errors,
        "order_numbers": order_numbers,
        "finished_at":   datetime.now(timezone.utc),
    })

    with _run_lock:
        if _active_run_id == run_id:
            _active_run_id = None


def _create_run_doc(trigger: str) -> str:
    run_id = str(uuid.uuid4())
    db().collection(RUNS_COL).document(run_id).set({
        "id":          run_id,
        "started_at":  datetime.now(timezone.utc),
        "finished_at": None,
        "status":      "running",
        "trigger":     trigger,
        "log":           "",
        "svgs":          0,
        "errors":        0,
        "order_numbers": [],
        "notes":         "",
    })
    return run_id


def _launch(run_id: str, extra_args: list[str] = []):
    global _active_run_id
    with _run_lock:
        _active_run_id = run_id
    t = threading.Thread(target=_run_generator, args=(run_id, extra_args), daemon=True)
    t.start()


# ── Schedule helpers ───────────────────────────────────────────────────────────
def _load_sched() -> dict:
    doc = db().collection(CFG_COL).document(SCHED_DOC).get()
    if doc.exists:
        return doc.to_dict()
    return {"enabled": False, "cron": "0 8 * * 1-5", "timezone": "America/New_York"}


def _apply_sched(cfg: dict):
    if scheduler.get_job("auto"):
        scheduler.remove_job("auto")
    if cfg.get("enabled") and cfg.get("cron"):
        tz = cfg.get("timezone", "America/New_York")
        try:
            scheduler.add_job(
                lambda: _launch(_create_run_doc("scheduled")),
                CronTrigger.from_crontab(cfg["cron"], timezone=tz),
                id="auto",
                replace_existing=True,
            )
            app.logger.info(f"Schedule active: {cfg['cron']} ({tz})")
        except Exception as exc:
            app.logger.warning(f"Schedule apply failed: {exc}")


_apply_sched(_load_sched())


# ── Gavel order sync ──────────────────────────────────────────────────────────
def _sync_gavel_orders(days=7):
    from gavel_eps_generator import (
        ss_get, is_gavel, fetch_customization, parse_customization,
        _is_silver_band_order, _sku_has_soundblock, PAGE_SIZE,
    )

    # ── Open a run-history entry so the dashboard shows this sync ──
    run_id    = str(uuid.uuid4())
    run_ref   = db().collection(RUNS_COL).document(run_id)
    run_start = datetime.now(timezone.utc)
    run_ref.set({
        "id": run_id, "started_at": run_start, "finished_at": None,
        "status": "running", "trigger": "sync", "log": "",
        "svgs": 0, "errors": 0, "order_numbers": [], "notes": "",
    })

    cutoff     = run_start - timedelta(days=days)
    cutoff_str = cutoff.strftime("%Y-%m-%d %H:%M:%S")
    log_lines  = [f"ShipStation sync — last {days} days (from {cutoff_str} UTC)"]

    params = {
        "orderStatus": "awaiting_shipment",
        "createDateStart": cutoff_str,
        "pageSize": PAGE_SIZE,
        "page": 1,
    }

    synced_count         = 0
    error_count          = 0
    cancelled_count      = 0
    synced_order_numbers = set()   # order numbers for which we created new docs
    seen_order_numbers   = set()   # all order numbers returned by ShipStation
    col = db().collection(GAVEL_ORDERS_COL)

    while True:
        data = ss_get("/orders", params)
        orders = data.get("orders", [])
        if not orders:
            break

        for order in orders:
            order_number = order.get("orderNumber", "")
            seen_order_numbers.add(order_number)
            ship = order
            for item_idx, item in enumerate(order.get("items", [])):
                if not is_gavel(item):
                    continue
                opts = {o["name"]: o["value"] for o in item.get("options", []) if "name" in o}
                url = opts.get("CustomizedURL")
                if not url:
                    continue

                doc_id  = f"{order_number}_{item_idx}"
                if col.document(doc_id).get().exists:
                    continue

                sku          = item.get("sku", "")
                item_name    = item.get("name", "")
                ship_to      = order.get("shipTo", {})
                ship_by_raw  = order.get("shipByDate") or ""
                order_date_raw = order.get("orderDate") or ""

                try:
                    cust_json = fetch_customization(url)
                    parsed    = parse_customization(cust_json)
                    want_sb   = parsed["sb_option"] == "custom_engraved" and bool(parsed["sb_lines"])
                    flags = {
                        "gavel_only":         parsed["sb_option"] in (None, "no_engraving", "unknown") and not _sku_has_soundblock(sku),
                        "sound_block_no":     (parsed["sb_option"] in ("no_engraving", "unknown") or parsed["sb_option"] is None) and _sku_has_soundblock(sku) and not want_sb,
                        "sound_block_custom": want_sb,
                        "gift_bag_gavel":     parsed["wants_suede_gavel"],
                        "gift_bag_sb":        parsed["wants_suede_sb"],
                    }
                    doc = {
                        "order_number": order_number,
                        "item_idx":     item_idx,
                        "sku":          sku,
                        "item_name":    item_name,
                        "qty":          item.get("quantity", 1),
                        "customer":     ship_to.get("name", ""),
                        "order_date":   order_date_raw[:10] if order_date_raw else "",
                        "ship_by_date": ship_by_raw[:10]    if ship_by_raw    else "",
                        "font":         parsed.get("font", "Arial"),
                        "text_lines":   parsed.get("band_lines", []),
                        "sb_option":    parsed.get("sb_option"),
                        "sb_lines":     parsed.get("sb_lines", []),
                        "sb_font":      parsed.get("sb_font", "Arial"),
                        "want_sb":      want_sb,
                        "wants_suede_gavel": parsed.get("wants_suede_gavel", False),
                        "wants_suede_sb":    parsed.get("wants_suede_sb",    False),
                        "is_silver":    _is_silver_band_order(ship),
                        "band_template": "7inch" if any(kw in item_name.lower() for kw in ["walnut", "black"]) else "standard",
                        "flags":        flags,
                        "ship_to":      ship_to,
                        "synced_at":    datetime.now(timezone.utc),
                        "status":       "ready",
                        "error":        None,
                    }
                    col.document(doc_id).set(doc)
                    synced_count += 1
                    synced_order_numbers.add(order_number)
                    log_lines.append(f"  + {order_number} item {item_idx}  ({ship_to.get('name', '')})")
                except Exception as e:
                    error_count += 1
                    col.document(doc_id).set({
                        "order_number": order_number,
                        "item_idx":     item_idx,
                        "sku":          sku,
                        "item_name":    item_name,
                        "qty":          item.get("quantity", 1),
                        "customer":     ship_to.get("name", ""),
                        "order_date":   order_date_raw[:10] if order_date_raw else "",
                        "ship_by_date": ship_by_raw[:10]    if ship_by_raw    else "",
                        "ship_to":      ship_to,
                        "synced_at":    datetime.now(timezone.utc),
                        "status":       "error",
                        "error":        str(e),
                    })
                    log_lines.append(f"  ! {order_number} item {item_idx} — {e}")

        total = data.get("total", 0)
        page  = params["page"]
        if page * PAGE_SIZE >= total:
            break
        params["page"] = page + 1

    # ── Mark orders no longer in awaiting_shipment as cancelled ──
    if seen_order_numbers:
        try:
            existing_in_window = col.where("synced_at", ">=", cutoff).stream()
            for ex_doc in existing_in_window:
                ed = ex_doc.to_dict()
                on = ed.get("order_number", "")
                if on and on not in seen_order_numbers and ed.get("status") != "cancelled":
                    col.document(ex_doc.id).update({"status": "cancelled"})
                    cancelled_count += 1
                    log_lines.append(f"  ~ {on} marked cancelled")
        except Exception as ce:
            app.logger.warning(f"Cancellation check failed: {ce}")
            log_lines.append(f"  ! Cancellation check failed: {ce}")

    log_lines.append(
        f"Done — {synced_count} new item(s), {error_count} error(s), {cancelled_count} cancelled"
    )
    final_status = "error" if error_count > 0 else "success"
    run_ref.update({
        "finished_at":   datetime.now(timezone.utc),
        "status":        final_status,
        "errors":        error_count,
        "order_numbers": list(synced_order_numbers),
        "notes":         f"{synced_count} new, {cancelled_count} cancelled",
        "log":           "\n".join(log_lines),
    })

    return synced_count, error_count


scheduler.add_job(
    lambda: _sync_gavel_orders(),
    "interval", hours=1, id="gavel_sync", replace_existing=True,
    next_run_time=datetime.now(timezone.utc),
)


# ── Routes ────────────────────────────────────────────────────────────────────
@app.route("/")
def dashboard():
    return render_template("dashboard.html")


@app.route("/api/run", methods=["POST"])
def trigger_run():
    global _active_run_id
    with _run_lock:
        if _active_run_id:
            return jsonify({"error": "run_in_progress", "run_id": _active_run_id}), 409

    data       = request.json or {}
    orders     = [o.strip() for o in data.get("orders", []) if o.strip()]
    extra_args = ["--orders"] + orders if orders else []
    days       = data.get("days")
    if days:
        extra_args += ["--days", str(int(days))]
    extra_args.append("--no-trello")

    run_id = _create_run_doc("manual")
    _launch(run_id, extra_args)
    return jsonify({"run_id": run_id})


@app.route("/api/run/<run_id>")
def get_run(run_id):
    doc = db().collection(RUNS_COL).document(run_id).get()
    if not doc.exists:
        return jsonify({"error": "not found"}), 404
    return jsonify(_serial(doc.to_dict()))


@app.route("/api/runs")
def list_runs():
    docs = (
        db().collection(RUNS_COL)
        .order_by("started_at", direction=firestore.Query.DESCENDING)
        .limit(100)
        .stream()
    )
    return jsonify([_serial(d.to_dict()) for d in docs])


@app.route("/api/notes/<run_id>", methods=["POST"])
def set_notes(run_id):
    notes = (request.json or {}).get("notes", "")
    db().collection(RUNS_COL).document(run_id).update({"notes": notes})
    return jsonify({"ok": True})


@app.route("/api/schedule", methods=["GET"])
def get_schedule():
    cfg = _load_sched()
    job = scheduler.get_job("auto")
    cfg["next_run"] = job.next_run_time.isoformat() if job and job.next_run_time else None
    return jsonify(cfg)


@app.route("/api/schedule", methods=["POST"])
def set_schedule():
    body    = request.json or {}
    cron    = body.get("cron", "").strip()
    enabled = bool(body.get("enabled", False))
    tz      = body.get("timezone", "America/New_York").strip()

    if enabled:
        try:
            CronTrigger.from_crontab(cron, timezone=tz)
        except Exception as exc:
            return jsonify({"error": str(exc)}), 400

    cfg = {"enabled": enabled, "cron": cron, "timezone": tz}
    db().collection(CFG_COL).document(SCHED_DOC).set(cfg)
    _apply_sched(cfg)

    job = scheduler.get_job("auto")
    cfg["next_run"] = job.next_run_time.isoformat() if job and job.next_run_time else None
    return jsonify(cfg)


@app.route("/api/status")
def api_status():
    with _run_lock:
        return jsonify({"active_run_id": _active_run_id})


# Called by Cloud Scheduler (optional — APScheduler handles it in-process too)
@app.route("/cron", methods=["POST"])
def cron_endpoint():
    token = request.headers.get("X-Cron-Token", "")
    if token != os.environ.get("CRON_TOKEN", ""):
        return jsonify({"error": "unauthorized"}), 401
    run_id = _create_run_doc("scheduled")
    _launch(run_id, ["--no-trello"])
    return jsonify({"run_id": run_id})


@app.route("/gavels")
def gavels_page():
    return render_template("gavels.html")


@app.route("/api/gavel-orders")
def list_gavel_orders():
    from collections import defaultdict
    days = int(request.args.get("days", 14))
    show_cancelled = request.args.get("cancelled", "false").lower() == "true"
    cutoff = datetime.now(timezone.utc) - timedelta(days=days)
    docs = (
        db().collection(GAVEL_ORDERS_COL)
        .where("synced_at", ">=", cutoff)
        .order_by("synced_at", direction=firestore.Query.DESCENDING)
        .limit(500)
        .stream()
    )

    groups = defaultdict(list)
    for d in docs:
        data = {**_serial(d.to_dict()), "doc_id": d.id}
        groups[data.get("order_number", d.id)].append(data)

    orders = []
    for order_number, items in groups.items():
        items.sort(key=lambda x: x.get("item_idx", 0))
        all_cancelled = all(it.get("status") == "cancelled" for it in items)
        if not show_cancelled and all_cancelled:
            continue

        first = items[0]
        merged_flags = {}
        for it in items:
            for k, v in (it.get("flags") or {}).items():
                if v:
                    merged_flags[k] = True

        total_qty = sum(it.get("qty", 1) for it in items)
        statuses = {it.get("status") for it in items}
        if all_cancelled:
            status = "cancelled"
        elif "error" in statuses:
            status = "error"
        else:
            status = "ready"

        orders.append({
            "order_number": order_number,
            "customer": first.get("customer", ""),
            "order_date": first.get("order_date", ""),
            "ship_by_date": first.get("ship_by_date", ""),
            "total_qty": total_qty,
            "item_count": len(items),
            "flags": merged_flags,
            "status": status,
            "ship_to": first.get("ship_to", {}),
            "items": items,
        })

    orders.sort(key=lambda o: o.get("order_date") or "", reverse=True)
    return jsonify(orders)


@app.route("/api/gavel-orders/sync", methods=["POST"])
def sync_gavel_orders():
    days = int((request.json or {}).get("days", 7))
    try:
        synced, errors = _sync_gavel_orders(days=days)
        return jsonify({"synced": synced, "errors": errors})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/gavel-orders/<order_number>/svg")
def download_order_svg(order_number):
    from gavel_eps_generator import write_individual_svg, build_layout_svg, resolve_font_path, TEMPLATE_PATH, TEMPLATE_PATH_7
    from flask import Response
    docs_stream = db().collection(GAVEL_ORDERS_COL).where("order_number", "==", order_number).stream()
    items_data = [d.to_dict() for d in docs_stream]
    items_data = [d for d in items_data if d.get("status") != "cancelled"]
    if not items_data:
        return jsonify({"error": "not found"}), 404
    items_data.sort(key=lambda x: x.get("item_idx", 0))

    layout_items = []
    for d in items_data:
        if not d.get("text_lines"):
            continue
        fp, ef = resolve_font_path(d.get("font", "Arial"))
        if not fp:
            continue
        bt = TEMPLATE_PATH_7 if d.get("band_template") == "7inch" else TEMPLATE_PATH
        for _ in range(d.get("qty", 1)):
            layout_items.append({
                "order_number": d["order_number"],
                "customer": d.get("customer", ""),
                "sku": d.get("sku", ""),
                "font": ef, "font_path": fp,
                "lines": d.get("text_lines", []),
                "template_path": bt,
                "is_silver": d.get("is_silver", False),
            })
    if not layout_items:
        return jsonify({"error": "no valid items"}), 400

    if len(layout_items) == 1:
        it = layout_items[0]
        with tempfile.NamedTemporaryFile(suffix=".svg", delete=False) as tmp:
            tmp_path = tmp.name
        try:
            write_individual_svg(tmp_path, it["lines"], it["font"], it["font_path"], it["template_path"])
            svg_content = open(tmp_path, "r", encoding="utf-8").read()
        finally:
            try: os.unlink(tmp_path)
            except: pass
    else:
        svg_content = build_layout_svg(layout_items)

    safe = order_number.replace("-", "")
    return Response(svg_content, mimetype="image/svg+xml",
        headers={"Content-Disposition": f'attachment; filename="{safe}_band.svg"'})


@app.route("/api/gavel-orders/<order_number>/pdf")
def download_order_pdf(order_number):
    from gavel_eps_generator import write_packing_slip, html_to_pdf
    from flask import Response
    docs_stream = db().collection(GAVEL_ORDERS_COL).where("order_number", "==", order_number).stream()
    items_data = [d.to_dict() for d in docs_stream]
    items_data = [d for d in items_data if d.get("status") != "cancelled"]
    if not items_data:
        return jsonify({"error": "not found"}), 404
    items_data.sort(key=lambda x: x.get("item_idx", 0))

    first = items_data[0]
    ship = {
        "orderNumber": order_number,
        "orderDate": first.get("order_date", ""),
        "requestedShippingService": "",
        "shipTo": first.get("ship_to", {}),
        "items": [],
    }
    slip_items = []
    for d in items_data:
        slip_items.append({
            "sku": d.get("sku", ""), "item_name": d.get("item_name", ""),
            "qty": d.get("qty", 1), "font": d.get("font", "Arial"),
            "text_lines": d.get("text_lines", []),
            "sb_option": d.get("sb_option"), "sb_lines": d.get("sb_lines", []),
            "sb_font": d.get("sb_font", "Arial"), "want_sb": d.get("want_sb", False),
            "wants_suede_gavel": d.get("wants_suede_gavel", False),
            "wants_suede_sb": d.get("wants_suede_sb", False), "sb_font_error": None,
        })
    with tempfile.NamedTemporaryFile(suffix=".html", delete=False) as h:
        html_path = h.name
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as p:
        pdf_path = p.name
    try:
        write_packing_slip(html_path, order_number, first.get("customer", ""), ship, slip_items)
        ok = html_to_pdf(html_path, pdf_path)
        if not ok:
            return jsonify({"error": "PDF generation failed"}), 500
        pdf_bytes = open(pdf_path, "rb").read()
    finally:
        for f in [html_path, pdf_path]:
            try: os.unlink(f)
            except: pass
    safe = order_number.replace("-", "")
    return Response(pdf_bytes, mimetype="application/pdf",
        headers={"Content-Disposition": f'attachment; filename="{safe}_workorder.pdf"'})


@app.route("/api/gavel-orders/combined", methods=["POST"])
def combined_gavel_layout():
    from gavel_eps_generator import build_layout_svg, resolve_font_path, TEMPLATE_PATH, TEMPLATE_PATH_7
    from flask import Response
    body = request.json or {}
    order_numbers = body.get("order_numbers", [])
    if not order_numbers:
        return jsonify({"error": "no order_numbers"}), 400
    items = []
    col = db().collection(GAVEL_ORDERS_COL)
    for order_number in order_numbers:
        docs_stream = col.where("order_number", "==", order_number).stream()
        order_docs = sorted(
            [d.to_dict() for d in docs_stream],
            key=lambda x: x.get("item_idx", 0)
        )
        for d in order_docs:
            if d.get("status") == "cancelled":
                continue
            if not d.get("text_lines"):
                continue
            fp, ef = resolve_font_path(d.get("font", "Arial"))
            if not fp:
                continue
            bt = TEMPLATE_PATH_7 if d.get("band_template") == "7inch" else TEMPLATE_PATH
            for _ in range(d.get("qty", 1)):
                items.append({
                    "order_number": d["order_number"], "customer": d.get("customer", ""),
                    "sku": d.get("sku", ""), "font": ef, "font_path": fp,
                    "lines": d.get("text_lines", []), "template_path": bt,
                    "is_silver": d.get("is_silver", False),
                })
    if not items:
        return jsonify({"error": "no valid items"}), 400
    svg = build_layout_svg(items)
    return Response(svg, mimetype="image/svg+xml",
        headers={"Content-Disposition": 'attachment; filename="combined_layout.svg"'})


def _serial(d: dict) -> dict:
    return {
        k: (v.isoformat() if hasattr(v, "isoformat") else v)
        for k, v in d.items()
    }


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port, debug=False)
