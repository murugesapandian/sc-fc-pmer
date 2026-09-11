"""Covers app.generate — the orchestration layer end to end, using a fully
isolated environment (see isolated_app_env in conftest.py) so these tests
never touch the real workbook, database, or reports folder. Uses the real
PPTX template (read-only, safe) to exercise the actual build_presentation
path rather than mocking it away.
"""
import json

import openpyxl

from app.generate import run_generation
from app import store


def test_first_generation_succeeds_and_writes_a_dated_filename(isolated_app_env):
    result = run_generation(trigger="manual")
    assert result.generated is True
    assert result.report.filename.startswith("SC-FC-PMER-")
    assert (isolated_app_env["reports_dir"] / result.report.filename).exists()


def test_second_generation_with_no_change_is_skipped(isolated_app_env):
    run_generation(trigger="manual")
    result = run_generation(trigger="manual")
    assert result.generated is False
    assert "No changes" in result.reason
    # Only one file should exist — the skip must not have written anything.
    assert len(list(isolated_app_env["reports_dir"].glob("*.pptx"))) == 1


def test_generation_triggers_after_a_real_workbook_change(isolated_app_env):
    run_generation(trigger="manual")

    wb = openpyxl.load_workbook(isolated_app_env["workbook_path"])
    wb["Master Plan"].cell(row=7, column=14, value="Complete")
    wb.save(isolated_app_env["workbook_path"])

    result = run_generation(trigger="manual")
    assert result.generated is True
    assert len(list(isolated_app_env["reports_dir"].glob("*.pptx"))) == 2


def test_force_with_no_real_change_refuses_to_create_duplicate(isolated_app_env):
    """Regression: 'Regenerate anyway' must still refuse to write a
    near-duplicate file when the rendered content hasn't actually changed,
    even though it bypasses the fast hash check."""
    first = run_generation(trigger="manual", force=True)
    assert first.generated is True

    second = run_generation(trigger="manual", force=True)
    assert second.generated is False
    assert "same content" in second.reason
    assert len(list(isolated_app_env["reports_dir"].glob("*.pptx"))) == 1
    # No leftover scratch file from the comparison.
    assert not list(isolated_app_env["reports_dir"].glob(".scratch-*"))


def test_force_with_a_genuine_change_still_creates_a_new_report(isolated_app_env):
    run_generation(trigger="manual", force=True)

    manual_path = isolated_app_env["manual_inputs_path"]
    data = json.loads(manual_path.read_text())
    data["decisions_needed"][0]["recommendation"] = "A genuinely different recommendation text."
    manual_path.write_text(json.dumps(data))

    result = run_generation(trigger="manual", force=True)
    assert result.generated is True
    assert len(list(isolated_app_env["reports_dir"].glob("*.pptx"))) == 2


def test_force_on_first_ever_generation_does_not_need_a_previous_report(isolated_app_env):
    """force=True with empty history must just generate normally, not try
    to compare against a report that doesn't exist yet."""
    result = run_generation(trigger="manual", force=True)
    assert result.generated is True


def test_scheduled_trigger_is_recorded_as_scheduled(isolated_app_env):
    result = run_generation(trigger="scheduled")
    assert result.report.trigger == "scheduled"


def test_forced_trigger_is_recorded_as_forced_even_when_caller_says_manual(isolated_app_env):
    result = run_generation(trigger="manual", force=True)
    assert result.report.trigger == "forced"


def test_filenames_stay_unique_after_deleting_and_regenerating_same_day(isolated_app_env):
    """Regression: deleting a report and regenerating on the same calendar
    day must not silently reuse the deleted report's filename."""
    first = run_generation(trigger="manual", force=True)
    (isolated_app_env["reports_dir"] / first.report.filename).unlink()

    manual_path = isolated_app_env["manual_inputs_path"]
    data = json.loads(manual_path.read_text())
    data["decisions_needed"][0]["recommendation"] = "Different, to force a real change."
    manual_path.write_text(json.dumps(data))

    second = run_generation(trigger="manual", force=True)
    assert second.generated is True
    assert second.report.filename != first.report.filename


def test_kpi_snapshot_is_stored_and_retrievable(isolated_app_env):
    result = run_generation(trigger="manual")
    assert result.report.kpi_json is not None
    kpi = json.loads(result.report.kpi_json)
    assert "status" in kpi
    assert "days_to_go_live" in kpi
