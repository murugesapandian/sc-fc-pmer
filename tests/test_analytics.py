"""Covers app.analytics — the schedule forecast, budget pace, trend deltas,
and the two rule-based risk flags. These are plain arithmetic over real
fields, so tests assert the actual formula results, not just 'it runs'."""
from datetime import date, timedelta

from app import analytics
from app.excel_reader import load_snapshot
from app.manual_inputs import DEFAULT_MANUAL_INPUTS


def _manual(**overrides):
    from app.manual_inputs import ManualInputs
    data = {**DEFAULT_MANUAL_INPUTS, **overrides}
    return ManualInputs(
        program_status_override=data.get("program_status_override"),
        budget=data.get("budget", {}),
        decisions_needed=data.get("decisions_needed", []),
        objective_overrides=data.get("objective_overrides", {}),
        workstream_focus_overrides=data.get("workstream_focus_overrides", {}),
        milestones=data.get("milestones", []),
    )


def test_schedule_forecast_direction_behind_when_completions_ran_late(test_workbook_path):
    """T-001 was due 05-Jan-24, last_updated 06-Jan-24 -> +1 day late.
    T-002 was due 12-Jan-24, last_updated 13-Jan-24 -> +1 day late.
    Average variance = +1 day; plus 2 currently-overdue open tasks add
    concrete slippage on top -> direction must be 'behind', never 'ahead'."""
    s = load_snapshot(test_workbook_path)
    forecast = analytics.schedule_forecast(s)
    assert forecast.direction == "behind"
    assert forecast.shift_days > 0
    assert forecast.projected_go_live > s.meta.target_go_live
    assert forecast.sample_size == 2  # T-001, T-002 completed with both dates present


def test_schedule_forecast_low_confidence_with_thin_sample(test_workbook_path):
    s = load_snapshot(test_workbook_path)
    forecast = analytics.schedule_forecast(s)
    assert forecast.confidence == "low"  # only 2 completed tasks


def test_budget_pace_remaining_and_run_rate(test_workbook_path):
    s = load_snapshot(test_workbook_path)
    manual = _manual(budget={
        "forecast_at_completion": 120,
        "by_category": [{"name": "Only", "approved": 100, "committed": 50, "spent": 20}],
    })
    pace = analytics.budget_pace(s, manual)
    assert pace.remaining == 80  # 100 approved - 20 spent
    # The project's own start (2024-01-01) is far in the past relative to
    # "today" for this test, so elapsed_pct should clamp at 1.0, not exceed it.
    assert 0.0 <= pace.elapsed_pct <= 1.0


def test_budget_pace_skips_run_rate_with_no_elapsed_time():
    """If today == project start, elapsed_pct is ~0 and a run-rate
    projection (spend / elapsed_pct) would divide by ~zero — must be
    skipped rather than returning a misleading number."""
    from app.excel_reader import ProjectSnapshot, ProjectMeta, KpiTotals

    snapshot = ProjectSnapshot(
        meta=ProjectMeta(name="X", start=date.today(), target_go_live=date.today() + timedelta(days=100)),
        kpis=KpiTotals(0, 0, 0, 0, 0, 0, 0, 0, 0),
        phases=[], workstreams=[], stakeholder_loads=[], stakeholders={},
        objectives=[], gates=[], raid_items=[], actions=[], tasks=[],
    )
    manual = _manual(budget={"forecast_at_completion": 100, "by_category": [{"name": "A", "approved": 100, "committed": 10, "spent": 5}]})
    pace = analytics.budget_pace(snapshot, manual)
    assert pace.has_enough_data is False
    assert pace.run_rate_projection is None


def test_trend_deltas_no_previous_report_returns_none():
    deltas = analytics.trend_deltas({"open_raid": 5}, None)
    assert deltas.open_raid_delta is None


def test_trend_deltas_computes_movement_both_directions():
    current = {"pct_complete": 0.10, "open_raid": 7, "high_severity_raid": 2, "days_to_go_live": 380}
    previous = {"pct_complete": 0.05, "open_raid": 5, "high_severity_raid": 1, "days_to_go_live": 390}
    deltas = analytics.trend_deltas(current, previous)
    assert deltas.open_raid_delta == 2
    assert deltas.days_to_go_live_delta == -10  # moved 10 days earlier


def test_workload_imbalance_none_with_too_few_stakeholders(test_workbook_path):
    s = load_snapshot(test_workbook_path)
    # Fixture has only one stakeholder with open work — below the
    # 3-stakeholder minimum sample size for a meaningful outlier check.
    assert analytics.workload_imbalance(s) is None


def test_gate_slip_risk_flags_gate_due_soon_with_open_items():
    from app.excel_reader import ProjectSnapshot, ProjectMeta, KpiTotals, Gate, GateItem

    gate = Gate(number=1, name="Gate 1", status_text="OPEN — 2 item(s)", done_fraction="0 / 2 done",
                items=[GateItem("G1", "a", None, "Open"), GateItem("G1", "b", None, "Open")],
                due=date.today() + timedelta(days=5))
    snapshot = ProjectSnapshot(
        meta=ProjectMeta(name="X", start=date.today(), target_go_live=date.today() + timedelta(days=100)),
        kpis=KpiTotals(0, 0, 0, 0, 0, 0, 0, 0, 0),
        phases=[], workstreams=[], stakeholder_loads=[], stakeholders={},
        objectives=[], gates=[gate], raid_items=[], actions=[], tasks=[],
    )
    flag = analytics.gate_slip_risk(snapshot)
    assert flag is not None
    assert "Gate 1" in flag


def test_gate_slip_risk_none_when_gate_far_out():
    from app.excel_reader import ProjectSnapshot, ProjectMeta, KpiTotals, Gate, GateItem

    gate = Gate(number=1, name="Gate 1", status_text="OPEN — 2 item(s)", done_fraction="0 / 2 done",
                items=[GateItem("G1", "a", None, "Open")],
                due=date.today() + timedelta(days=60))
    snapshot = ProjectSnapshot(
        meta=ProjectMeta(name="X", start=date.today(), target_go_live=date.today() + timedelta(days=100)),
        kpis=KpiTotals(0, 0, 0, 0, 0, 0, 0, 0, 0),
        phases=[], workstreams=[], stakeholder_loads=[], stakeholders={},
        objectives=[], gates=[gate], raid_items=[], actions=[], tasks=[],
    )
    assert analytics.gate_slip_risk(snapshot) is None
