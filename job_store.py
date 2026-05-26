"""
job_store.py
============
SQLite-backed persistent store for RPVE job metadata.

Thread-safety:
  - Each public function opens its own connection so that worker threads
    don't share a single connection object.
  - WAL mode is enabled on DB creation, allowing concurrent readers while
    a writer is active.
  - Writes use BEGIN IMMEDIATE to serialise concurrent writes.

No ORM — plain sqlite3 from the stdlib.
"""

import json
import sqlite3
import threading
from datetime import datetime
from pathlib import Path
from typing import Optional

from job_models import JobMeta

# Absolute path to the SQLite database file, next to this module.
_DB_PATH: Path = Path(__file__).parent / "rpve_jobs.db"
_INIT_LOCK = threading.Lock()
_initialised = False


# ─────────────────────────────────────────────────────────────────────────────
# Internal helpers
# ─────────────────────────────────────────────────────────────────────────────

def _connect() -> sqlite3.Connection:
    """Open a new SQLite connection with WAL mode and row_factory."""
    conn = sqlite3.connect(str(_DB_PATH), check_same_thread=False, timeout=10)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=NORMAL")
    return conn


def _ensure_schema() -> None:
    """Create the jobs table if it doesn't exist yet (idempotent)."""
    global _initialised
    with _INIT_LOCK:
        if _initialised:
            return
        conn = _connect()
        try:
            conn.execute("""
                CREATE TABLE IF NOT EXISTS jobs (
                    job_id          TEXT PRIMARY KEY,
                    status          TEXT NOT NULL DEFAULT 'queued',
                    template_type   TEXT,
                    created_at      TEXT NOT NULL,
                    updated_at      TEXT NOT NULL,
                    error           TEXT,
                    phase1_output   TEXT,
                    phase2_output   TEXT,
                    phase3_output   TEXT,
                    phase4_output   TEXT,
                    final_output    TEXT,
                    result_json     TEXT
                )
            """)
            conn.commit()
        finally:
            conn.close()
        _initialised = True


def _row_to_meta(row: sqlite3.Row) -> JobMeta:
    d = dict(row)
    return JobMeta(**d)


# ─────────────────────────────────────────────────────────────────────────────
# Public API
# ─────────────────────────────────────────────────────────────────────────────

def create_job(job_id: str) -> JobMeta:
    """Insert a new job record with status=queued."""
    _ensure_schema()
    now = datetime.utcnow().isoformat()
    meta = JobMeta(job_id=job_id, status="queued", created_at=now, updated_at=now)
    conn = _connect()
    try:
        with conn:
            conn.execute(
                """
                INSERT INTO jobs
                    (job_id, status, created_at, updated_at)
                VALUES (?, ?, ?, ?)
                """,
                (meta.job_id, meta.status, meta.created_at, meta.updated_at),
            )
    finally:
        conn.close()
    return meta


def update_status(
    job_id: str,
    status: str,
    error: Optional[str] = None,
    template_type: Optional[str] = None,
) -> None:
    """Transition a job to a new lifecycle status."""
    _ensure_schema()
    now = datetime.utcnow().isoformat()
    conn = _connect()
    try:
        with conn:
            conn.execute(
                """
                UPDATE jobs
                SET    status        = ?,
                       updated_at    = ?,
                       error         = COALESCE(?, error),
                       template_type = COALESCE(?, template_type)
                WHERE  job_id = ?
                """,
                (status, now, error, template_type, job_id),
            )
    finally:
        conn.close()


def set_phase_output(job_id: str, phase: int, path: str) -> None:
    """Record the output file path for a completed phase (1–4)."""
    _ensure_schema()
    col = f"phase{phase}_output"
    now = datetime.utcnow().isoformat()
    conn = _connect()
    try:
        with conn:
            conn.execute(
                f"UPDATE jobs SET {col} = ?, updated_at = ? WHERE job_id = ?",
                (path, now, job_id),
            )
    finally:
        conn.close()


def set_final_output(job_id: str, final_path: str, result_dict: dict) -> None:
    """Record the final download path and the rich result JSON blob."""
    _ensure_schema()
    now = datetime.utcnow().isoformat()
    conn = _connect()
    try:
        with conn:
            conn.execute(
                """
                UPDATE jobs
                SET    final_output = ?,
                       result_json  = ?,
                       updated_at   = ?
                WHERE  job_id = ?
                """,
                (final_path, json.dumps(result_dict, ensure_ascii=False), now, job_id),
            )
    finally:
        conn.close()


def get_job(job_id: str) -> Optional[JobMeta]:
    """Fetch a single job by ID. Returns None if not found."""
    _ensure_schema()
    conn = _connect()
    try:
        row = conn.execute(
            "SELECT * FROM jobs WHERE job_id = ?", (job_id,)
        ).fetchone()
        return _row_to_meta(row) if row else None
    finally:
        conn.close()


def list_jobs(limit: int = 50) -> list[JobMeta]:
    """Return the most recent jobs, newest first."""
    _ensure_schema()
    conn = _connect()
    try:
        rows = conn.execute(
            "SELECT * FROM jobs ORDER BY created_at DESC LIMIT ?", (limit,)
        ).fetchall()
        return [_row_to_meta(r) for r in rows]
    finally:
        conn.close()


def recover_stale_jobs() -> int:
    """
    Mark any job that was left in a non-terminal state (e.g. after a crash)
    as failed. Called once at server startup.

    Returns the number of jobs recovered.
    """
    _ensure_schema()
    now = datetime.utcnow().isoformat()
    conn = _connect()
    try:
        cur = conn.execute(
            """
            UPDATE jobs
            SET    status     = 'failed',
                   error      = 'Server restarted while job was running',
                   updated_at = ?
            WHERE  status NOT IN ('completed', 'failed')
            """,
            (now,),
        )
        conn.commit()
        return cur.rowcount
    finally:
        conn.close()
