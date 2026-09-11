"""Covers app.store — report history, the reconcile-against-disk self-heal,
and filename uniqueness across the DB's full history (not just current disk
state). Uses a tmp_path SQLite file via monkeypatch, never the real state.db.
"""
import pytest

from app import store


@pytest.fixture(autouse=True)
def _isolated_db(tmp_path, monkeypatch):
    monkeypatch.setattr(store, "DB_PATH", tmp_path / "test_state.db")


def test_record_and_list_reports():
    store.record_report("a.pptx", "manual", "hash1")
    store.record_report("b.pptx", "scheduled", "hash2")
    reports = store.list_reports()
    assert [r.filename for r in reports] == ["b.pptx", "a.pptx"]  # newest first


def test_record_report_updates_last_source_hash():
    store.record_report("a.pptx", "manual", "hash1")
    assert store.get_last_source_hash() == "hash1"
    store.record_report("b.pptx", "manual", "hash2")
    assert store.get_last_source_hash() == "hash2"


def test_reconcile_drops_rows_whose_file_is_missing(tmp_path):
    reports_dir = tmp_path / "reports"
    reports_dir.mkdir()
    (reports_dir / "kept.pptx").write_bytes(b"fake pptx bytes")
    # "deleted.pptx" is recorded but never written to disk.
    store.record_report("deleted.pptx", "manual", "hash1")
    store.record_report("kept.pptx", "manual", "hash2")

    remaining = store.reconcile_reports(reports_dir)

    assert [r.filename for r in remaining] == ["kept.pptx"]
    assert [r.filename for r in store.list_reports()] == ["kept.pptx"]


def test_reconcile_resyncs_last_source_hash_to_newest_survivor(tmp_path):
    reports_dir = tmp_path / "reports"
    reports_dir.mkdir()
    (reports_dir / "kept.pptx").write_bytes(b"x")
    store.record_report("kept.pptx", "manual", "hash-old")
    store.record_report("deleted.pptx", "manual", "hash-new")  # never written to disk

    store.reconcile_reports(reports_dir)

    # The surviving (older) report's hash is now what "last_source_hash"
    # should reflect — not the hash of the report that got dropped.
    assert store.get_last_source_hash() == "hash-old"


def test_reconcile_clears_last_source_hash_when_everything_is_gone(tmp_path):
    reports_dir = tmp_path / "reports"
    reports_dir.mkdir()
    store.record_report("gone.pptx", "manual", "hash1")

    store.reconcile_reports(reports_dir)

    assert store.get_last_source_hash() is None
    assert store.list_reports() == []


def test_all_filenames_includes_rows_even_after_their_file_is_deleted(tmp_path):
    """Regression: filenames must stay unique against the DB's full
    history, not just current disk state — otherwise deleting a report and
    regenerating the same day silently reuses its name."""
    reports_dir = tmp_path / "reports"
    reports_dir.mkdir()
    store.record_report("SC-FC-PMER-2024-01-01.pptx", "manual", "hash1")
    # Note: never reconciled/deleted from the DB, simulating a file that
    # was removed from disk without going through reconcile_reports.
    assert "SC-FC-PMER-2024-01-01.pptx" in store.all_filenames()


def test_all_filenames_survives_reconcile_pruning_the_visible_row(tmp_path):
    """Regression: reconcile_reports() deletes the `reports` row for a
    missing file (so the UI stops showing it) — that must NOT also erase
    the filename from all_filenames()'s permanent ledger, or the two
    features (hide deleted reports / never reuse a filename) undo each
    other, exactly as they did before used_filenames was split out."""
    reports_dir = tmp_path / "reports"
    reports_dir.mkdir()
    store.record_report("SC-FC-PMER-2024-01-01.pptx", "manual", "hash1")
    # File never actually written to reports_dir, so reconcile treats it as deleted.

    remaining = store.reconcile_reports(reports_dir)

    assert remaining == []  # gone from the visible history, as expected
    assert "SC-FC-PMER-2024-01-01.pptx" in store.all_filenames()  # but the name stays retired forever


def test_used_filenames_backfills_from_pre_existing_reports_rows(tmp_path, monkeypatch):
    """Regression: a database created before the used_filenames table
    existed must have its existing report filenames backfilled into it on
    the next connect — otherwise upgrading leaves old filenames
    unprotected against same-day reuse after their file is later deleted."""
    db_path = tmp_path / "legacy.db"
    monkeypatch.setattr(store, "DB_PATH", db_path)

    # Simulate a pre-migration database: create only the original schema,
    # bypassing store.record_report (which would also populate the new table).
    import sqlite3
    conn = sqlite3.connect(db_path)
    conn.execute(
        "CREATE TABLE reports (id INTEGER PRIMARY KEY AUTOINCREMENT, filename TEXT NOT NULL, "
        "generated_at TEXT NOT NULL, trigger TEXT NOT NULL, source_hash TEXT NOT NULL, notes TEXT)"
    )
    conn.execute(
        "INSERT INTO reports (filename, generated_at, trigger, source_hash) VALUES (?, ?, ?, ?)",
        ("legacy-report.pptx", "2024-01-01T00:00:00", "manual", "hash1"),
    )
    conn.commit()
    conn.close()

    assert "legacy-report.pptx" in store.all_filenames()


def test_get_report_returns_none_for_unknown_id():
    assert store.get_report(9999) is None


def test_kpi_json_round_trips():
    store.record_report("a.pptx", "manual", "hash1", kpi_json='{"status": "GREEN"}')
    record = store.list_reports()[0]
    assert record.kpi_json == '{"status": "GREEN"}'
