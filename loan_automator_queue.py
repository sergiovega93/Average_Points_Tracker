"""
SQLite-backed job queue for single-loan webhook processing (one worker thread).
"""
from __future__ import annotations

import json
import sqlite3
import threading
import time
from datetime import datetime, timezone
from pathlib import Path

import config

_schema = """
CREATE TABLE IF NOT EXISTS jobs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    loan_id INTEGER NOT NULL,
    payload_json TEXT,
    status TEXT NOT NULL DEFAULT 'pending',
    error TEXT,
    created_at TEXT NOT NULL,
    started_at TEXT,
    finished_at TEXT
);
CREATE INDEX IF NOT EXISTS idx_jobs_status ON jobs(status, id);
"""


def _conn() -> sqlite3.Connection:
    path = Path(config.QUEUE_DB_PATH)
    path.parent.mkdir(parents=True, exist_ok=True)
    c = sqlite3.connect(str(path), timeout=30, isolation_level=None)
    c.row_factory = sqlite3.Row
    return c


def init_db() -> None:
    with _conn() as c:
        c.executescript(_schema)


def enqueue(loan_id: int, payload_json: str | None = None) -> int:
    init_db()
    now = datetime.now(timezone.utc).isoformat()
    with _conn() as c:
        cur = c.execute(
            "INSERT INTO jobs (loan_id, payload_json, status, created_at) VALUES (?,?,?,?)",
            (int(loan_id), payload_json or "", "pending", now),
        )
        return int(cur.lastrowid)


def claim_next() -> tuple[int, int] | None:
    """Return (job_id, loan_id) or None. Uses BEGIN IMMEDIATE for single-writer lock."""
    init_db()
    c = _conn()
    try:
        c.execute("BEGIN IMMEDIATE")
        row = c.execute(
            "SELECT id, loan_id FROM jobs WHERE status = 'pending' ORDER BY id LIMIT 1"
        ).fetchone()
        if not row:
            c.rollback()
            return None
        jid, lid = int(row["id"]), int(row["loan_id"])
        now = datetime.now(timezone.utc).isoformat()
        c.execute(
            "UPDATE jobs SET status = 'processing', started_at = ? WHERE id = ?",
            (now, jid),
        )
        c.commit()
        return jid, lid
    finally:
        c.close()


def complete_job(job_id: int, *, error: str | None = None) -> None:
    now = datetime.now(timezone.utc).isoformat()
    status = "failed" if error else "done"
    with _conn() as c:
        c.execute(
            "UPDATE jobs SET status = ?, error = ?, finished_at = ? WHERE id = ?",
            (status, error or "", now, int(job_id)),
        )


_worker_started = False
_worker_lock = threading.Lock()


def _worker_loop() -> None:
    import logging

    log = logging.getLogger("loan_automator.worker")
    while True:
        jid: int | None = None
        try:
            claimed = claim_next()
            if not claimed:
                time.sleep(1.0)
                continue
            jid, lid = claimed
            log.info("Processing job id=%s loan_id=%s", jid, lid)
            from steps.single_loan import run_pipeline_for_loan_id

            run_pipeline_for_loan_id(lid, dry_run=False)
            complete_job(jid, error=None)
            log.info("Job id=%s done", jid)
        except Exception as exc:
            log.exception("Job failed: %s", exc)
            if jid is not None:
                try:
                    complete_job(jid, error=str(exc)[:2000])
                except Exception:
                    pass


def ensure_worker_started() -> None:
    global _worker_started
    with _worker_lock:
        if _worker_started:
            return
        init_db()
        t = threading.Thread(target=_worker_loop, name="loan-automator-queue", daemon=True)
        t.start()
        _worker_started = True
