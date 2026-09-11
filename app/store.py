"""SQLite persistence for report history and last-known-hash state."""
from __future__ import annotations

import sqlite3
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from app.config import DB_PATH

SCHEMA = """
CREATE TABLE IF NOT EXISTS reports (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    filename TEXT NOT NULL,
    generated_at TEXT NOT NULL,
    trigger TEXT NOT NULL CHECK (trigger IN ('manual', 'scheduled', 'forced')),
    source_hash TEXT NOT NULL,
    notes TEXT
);

CREATE TABLE IF NOT EXISTS state (
    key TEXT PRIMARY KEY,
    value TEXT
);

-- Deliberately separate from `reports` and never pruned by reconcile_reports:
-- that table's rows get deleted once their file is gone from disk (so the UI
-- stops showing phantom entries), but a filename must stay "used" forever
-- once assigned, or deleting a report and regenerating the same day would
-- silently reuse its name for different content.
CREATE TABLE IF NOT EXISTS used_filenames (
    filename TEXT PRIMARY KEY,
    recorded_at TEXT NOT NULL
);
"""


@contextmanager
def _connect(db_path: Path | None = None):
    # Resolved from the module global at call time rather than bound as a
    # default parameter value — a default is frozen at import time, so
    # tests monkeypatching store.DB_PATH would otherwise silently have no
    # effect and keep writing to the real database.
    if db_path is None:
        db_path = DB_PATH
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        conn.executescript(SCHEMA)
        existing_cols = {r["name"] for r in conn.execute("PRAGMA table_info(reports)").fetchall()}
        if "kpi_json" not in existing_cols:
            conn.execute("ALTER TABLE reports ADD COLUMN kpi_json TEXT")
        # Backfill for databases that predate the used_filenames table (or
        # any row inserted before it existed) — cheap and idempotent, so it
        # just runs on every connect rather than needing migration tracking.
        conn.execute(
            "INSERT OR IGNORE INTO used_filenames (filename, recorded_at) "
            "SELECT filename, generated_at FROM reports"
        )
        yield conn
        conn.commit()
    finally:
        conn.close()


@dataclass
class ReportRecord:
    id: int
    filename: str
    generated_at: str
    trigger: str
    source_hash: str
    notes: str | None
    kpi_json: str | None = None


def record_report(filename: str, trigger: str, source_hash: str, notes: str | None = None, kpi_json: str | None = None) -> ReportRecord:
    with _connect() as conn:
        now = datetime.now().isoformat(timespec="seconds")
        cur = conn.execute(
            "INSERT INTO reports (filename, generated_at, trigger, source_hash, notes, kpi_json) VALUES (?, ?, ?, ?, ?, ?)",
            (filename, now, trigger, source_hash, notes, kpi_json),
        )
        conn.execute(
            "INSERT OR IGNORE INTO used_filenames (filename, recorded_at) VALUES (?, ?)",
            (filename, now),
        )
        conn.execute(
            "INSERT INTO state (key, value) VALUES ('last_source_hash', ?) "
            "ON CONFLICT(key) DO UPDATE SET value = excluded.value",
            (source_hash,),
        )
        row = conn.execute("SELECT * FROM reports WHERE id = ?", (cur.lastrowid,)).fetchone()
        return ReportRecord(**dict(row))


def list_reports() -> list[ReportRecord]:
    with _connect() as conn:
        rows = conn.execute("SELECT * FROM reports ORDER BY id DESC").fetchall()
        return [ReportRecord(**dict(r)) for r in rows]


def reconcile_reports(reports_dir: Path) -> list[ReportRecord]:
    """Drops any history entry whose file no longer exists on disk (e.g. the
    user deleted it manually outside the app) and keeps last_source_hash in
    sync with what's actually still downloadable — otherwise a deleted report
    both lingers in the UI and falsely reports 'up to date' with nothing left
    to show for it. Returns the reconciled list, newest first."""
    with _connect() as conn:
        rows = conn.execute("SELECT * FROM reports ORDER BY id DESC").fetchall()
        stale_ids = [r["id"] for r in rows if not (reports_dir / r["filename"]).exists()]
        if stale_ids:
            conn.executemany("DELETE FROM reports WHERE id = ?", [(i,) for i in stale_ids])

        remaining = conn.execute("SELECT * FROM reports ORDER BY id DESC").fetchall()
        newest_hash = remaining[0]["source_hash"] if remaining else None
        if newest_hash is not None:
            conn.execute(
                "INSERT INTO state (key, value) VALUES ('last_source_hash', ?) "
                "ON CONFLICT(key) DO UPDATE SET value = excluded.value",
                (newest_hash,),
            )
        else:
            conn.execute("DELETE FROM state WHERE key = 'last_source_hash'")

        return [ReportRecord(**dict(r)) for r in remaining]


def get_report(report_id: int) -> ReportRecord | None:
    with _connect() as conn:
        row = conn.execute("SELECT * FROM reports WHERE id = ?", (report_id,)).fetchone()
        return ReportRecord(**dict(row)) if row else None


def all_filenames() -> set[str]:
    """Every filename ever assigned, permanently — including ones whose
    `reports` row has since been pruned by reconcile_reports because the
    file was deleted from disk. Read from `used_filenames`, not `reports`:
    the latter is a *visible history* that's allowed to shrink, this is a
    *permanent ledger* that must not, or a deleted report's name could be
    handed out again for different content."""
    with _connect() as conn:
        return {r["filename"] for r in conn.execute("SELECT filename FROM used_filenames").fetchall()}


def get_last_source_hash() -> str | None:
    with _connect() as conn:
        row = conn.execute("SELECT value FROM state WHERE key = 'last_source_hash'").fetchone()
        return row["value"] if row else None


def set_last_checked(when: str) -> None:
    with _connect() as conn:
        conn.execute(
            "INSERT INTO state (key, value) VALUES ('last_checked_at', ?) "
            "ON CONFLICT(key) DO UPDATE SET value = excluded.value",
            (when,),
        )
