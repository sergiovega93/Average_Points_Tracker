"""
Flask routes for Zapier → single-loan queue. Load from your create_app() via importlib
(see README). Ensures this repo root is on sys.path before importing config.
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

from flask import jsonify, request


def _ensure_repo_root(repo_root: Path | None) -> Path:
    root = repo_root or Path(__file__).resolve().parent
    s = str(root)
    if s not in sys.path:
        sys.path.insert(0, s)
    return root


def _auth_ok() -> bool:
    import config as cfg

    if not cfg.LOAN_AUTOMATOR_WEBHOOK_SECRET:
        return os.getenv("LOAN_AUTOMATOR_WEBHOOK_ALLOW_INSECURE") == "1"
    sec = cfg.LOAN_AUTOMATOR_WEBHOOK_SECRET
    if request.headers.get("X-Loan-Automator-Secret") == sec:
        return True
    if request.headers.get("X-Webhook-Secret") == sec:
        return True
    auth = request.headers.get("Authorization", "") or ""
    if auth.lower().startswith("bearer ") and auth[7:].strip() == sec:
        return True
    if request.args.get("token") == sec:
        return True
    body = request.get_json(silent=True) or {}
    if isinstance(body, dict) and body.get("secret") == sec:
        return True
    return False


def _extract_loan_id() -> int | None:
    body = request.get_json(silent=True)
    candidates: list = []

    def scan(d: dict) -> None:
        for k in ("loan_id", "loanId", "api_loan_id", "API Loan ID"):
            if k in d and d[k] not in (None, ""):
                candidates.append(d[k])

    if isinstance(body, dict):
        scan(body)
        inner = body.get("data")
        if isinstance(inner, dict):
            scan(inner)

    for raw in candidates:
        try:
            return int(str(raw).strip().replace(",", ""))
        except ValueError:
            continue

    for k in ("loan_id", "loanId"):
        v = request.form.get(k)
        if v:
            try:
                return int(str(v).strip())
            except ValueError:
                pass

    v = request.args.get("loan_id")
    if v:
        try:
            return int(str(v).strip())
        except ValueError:
            pass
    return None


def register_routes(app, *, repo_root: str | Path | None = None) -> None:
    """
    Register POST /webhook/loan-automator (queue) and GET (health).

    repo_root: absolute path to the loan-automator clone on disk (defaults to this file's dir).
    """
    _ensure_repo_root(Path(repo_root) if repo_root else None)

    import config as cfg

    from loan_automator_queue import enqueue, ensure_worker_started

    @app.route("/webhook/loan-automator", methods=["GET", "POST"])
    def loan_automator_webhook():
        if request.method == "GET":
            return jsonify(
                {
                    "service": "loan-automator",
                    "webhook_secret_configured": bool(cfg.LOAN_AUTOMATOR_WEBHOOK_SECRET),
                    "queue_db": cfg.QUEUE_DB_PATH,
                }
            )

        if not _auth_ok():
            return jsonify({"error": "unauthorized"}), 401

        lid = _extract_loan_id()
        if lid is None:
            return jsonify({"error": "loan_id required (JSON, form, or query)"}), 400

        raw = request.get_data(as_text=True)
        job_id = enqueue(lid, raw[:8000] if raw else None)
        ensure_worker_started()
        return jsonify({"queued": True, "job_id": job_id, "loan_id": lid}), 202

    app.logger.warning(
        "Loan Automator webhook registered: POST/GET /webhook/loan-automator"
    )
