import os
import re
import sys
import subprocess
import threading
import uuid
from datetime import datetime, timezone
from pathlib import Path

from flask import Flask, jsonify, render_template, request
from flask_cors import CORS
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger
from google.cloud import firestore

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
    _launch(run_id)
    return jsonify({"run_id": run_id})


def _serial(d: dict) -> dict:
    return {
        k: (v.isoformat() if hasattr(v, "isoformat") else v)
        for k, v in d.items()
    }


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port, debug=False)
