"""Covers the parsing/derivation logic in app.excel_reader, including
regression tests for the two real bugs found and fixed in this project:
- current_phase must be the earliest INCOMPLETE phase (data-driven), not a
  calendar-date lookup that breaks when 'today' is outside the plan's dates.
- phase/workstream/KPI rollups must be derived from Master Plan's raw Status
  column, not from Dashboard's own formula cells (which were found to have
  been silently overwritten with stale literal values).
"""
from datetime import date

from app.excel_reader import load_snapshot


def test_meta_and_target_go_live(test_workbook_path):
    s = load_snapshot(test_workbook_path)
    assert s.meta.name == "Test Project — Widget Rollout"
    assert s.meta.start == date(2024, 1, 1)
    # Derived as max(task.due), not read from a Dashboard formula cell.
    assert s.meta.target_go_live == date(2024, 1, 25)


def test_phase_rollups_derived_from_tasks_not_dashboard_formulas(test_workbook_path):
    s = load_snapshot(test_workbook_path)
    alpha, beta = s.phases
    assert alpha.name == "1. Alpha"
    assert alpha.tasks == 2 and alpha.complete == 2 and alpha.pct_complete == 1.0
    assert beta.tasks == 2 and beta.complete == 0
    assert beta.in_progress == 1
    # Sums across phases must exactly equal the top-line KPI totals — this
    # is the specific consistency the chart vs. KPI-tile bug violated.
    assert sum(p.complete for p in s.phases) == s.kpis.complete
    assert sum(p.tasks for p in s.phases) == s.kpis.total_tasks


def test_current_phase_is_earliest_incomplete_not_date_based(test_workbook_path):
    """Regression test: the fixture's dates are all in January 2024, far in
    the past relative to whenever this test runs — a date-based
    'start <= today <= due' lookup would find nothing and fall through to
    an arbitrary default. The fix must still correctly identify Beta (the
    earliest phase that isn't 100% complete) purely from completion data."""
    s = load_snapshot(test_workbook_path)
    assert s.current_phase.name == "2. Beta"
    assert s.next_phase is None  # Beta is the last phase


def test_gate_status_derived_from_raw_items_not_summary_cell(test_workbook_path):
    s = load_snapshot(test_workbook_path)
    gate1, gate2 = s.gates
    assert gate1.number == 1
    assert gate1.status_text == "PASSED"
    assert gate1.done_fraction == "2 / 2 done"
    assert gate1.is_open is False

    assert gate2.status_text == "OPEN — 2 item(s)"
    assert gate2.done_fraction == "0 / 2 done"
    assert gate2.is_open is True

    # current_gate = first open gate = Gate 2, independent of current_phase.
    assert s.current_gate.number == 2
    # Each gate's due date is attached from its corresponding phase.
    assert gate1.due == s.phases[0].due
    assert gate2.due == s.phases[1].due


def test_phase_gate_status_matches_the_gate_object(test_workbook_path):
    """Slide 2 and slide 4 both describe 'the gate tied to the current
    phase' — this is the field that must agree between them."""
    s = load_snapshot(test_workbook_path)
    assert s.phases[0].gate_status == "PASSED"
    assert s.phases[1].gate_status == "OPEN — 2 item(s)"


def test_kpi_totals(test_workbook_path):
    s = load_snapshot(test_workbook_path)
    assert s.kpis.total_tasks == 4
    assert s.kpis.complete == 2
    assert s.kpis.pct_complete == 0.5
    assert s.kpis.in_progress == 1
    assert s.kpis.overdue_tasks == 2  # T-003, T-004: not complete, due date long past
    assert s.kpis.open_raid == 2
    assert s.kpis.high_severity_raid == 1
    assert s.kpis.open_actions == 1


def test_workstream_rollup(test_workbook_path):
    s = load_snapshot(test_workbook_path)
    assert len(s.workstreams) == 1
    ws = s.workstreams[0]
    assert ws.name == "Engineering"
    assert ws.tasks == 4
    assert ws.complete == 2


def test_raid_and_action_open_filters(test_workbook_path):
    s = load_snapshot(test_workbook_path)
    assert len(s.open_raid_items) == 2
    assert len(s.high_severity_open_raid) == 1
    assert s.high_severity_open_raid[0].id == "R-001"
    assert len(s.open_actions) == 1


def test_stakeholder_name_lookup(test_workbook_path):
    s = load_snapshot(test_workbook_path)
    assert s.stakeholder_name("SH-01") == "Project Manager"
    assert s.stakeholder_name(None) == "Unassigned"
    assert s.stakeholder_name("SH-99") == "SH-99"  # unknown id falls back to the raw id


def test_days_to_go_live_and_reporting_month_are_computed_live(test_workbook_path):
    s = load_snapshot(test_workbook_path)
    assert s.days_to_go_live == (s.meta.target_go_live - date.today()).days
    assert s.reporting_month_number >= 1
