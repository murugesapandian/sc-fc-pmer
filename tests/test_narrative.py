"""Covers app.narrative — the rule-based text generation, including the
workstream_focus regression (comparing task.start to real 'today' breaks for
a plan dated entirely in the past or future; the fix reasons about phase
order instead)."""
from datetime import date

from app import narrative as nar
from app.excel_reader import load_snapshot


def test_program_status_red_when_overdue_tasks(test_workbook_path):
    from app.manual_inputs import ManualInputs
    s = load_snapshot(test_workbook_path)
    manual = ManualInputs(None, {}, [], {}, {}, [])
    # The fixture has 2 overdue tasks (due dates in the past, not complete).
    assert nar.program_status(s, manual) == "RED"


def test_program_status_override_wins(test_workbook_path):
    from app.manual_inputs import ManualInputs
    s = load_snapshot(test_workbook_path)
    manual = ManualInputs("green", {}, [], {}, {}, [])
    assert nar.program_status(s, manual) == "GREEN"


def test_workstream_health_flags_overdue_as_red(test_workbook_path):
    s = load_snapshot(test_workbook_path)
    ws = s.workstreams[0]
    assert nar.workstream_health(ws, s) == "Red"  # has overdue tasks


def test_workstream_focus_override_takes_precedence(test_workbook_path):
    s = load_snapshot(test_workbook_path)
    ws = s.workstreams[0]
    assert nar.workstream_focus(ws, s, "Custom focus line") == "Custom focus line"


def test_workstream_focus_reasons_about_phase_order_not_calendar_dates(test_workbook_path):
    """Regression: the fixture's tasks are dated in Jan 2024, far from
    whatever 'today' is when the test runs. workstream_focus must not
    collapse to 'Not yet active' just because task.start < today is false
    for a plan dated in the past (or > today for one dated in the future) —
    it must reason relative to the plan's own current phase."""
    s = load_snapshot(test_workbook_path)
    ws = s.workstreams[0]  # Engineering — has open work in the current phase (Beta)
    focus = nar.workstream_focus(ws, s, None)
    assert focus != "Not yet active"


def test_top_risks_excludes_low_severity(test_workbook_path):
    s = load_snapshot(test_workbook_path)
    risks = nar.top_risks(s)
    ids = [r.id for r in risks]
    assert "R-001" in ids  # High
    assert "R-002" not in ids  # Low — excluded, called out separately as "not shown"


def test_severity_distribution_counts_open_items_only(test_workbook_path):
    s = load_snapshot(test_workbook_path)
    dist = nar.severity_distribution(s)
    assert dist == {"High": 1, "Medium": 0, "Low": 1}


def test_phase_commentary_reflects_current_phase_data(test_workbook_path):
    s = load_snapshot(test_workbook_path)
    lines = nar.phase_commentary(s)
    joined = " ".join(lines)
    assert "Beta" in joined  # current phase, not frozen on phase 1


def test_exec_summary_sentence_uses_gate_tied_to_current_phase(test_workbook_path):
    """Regression: must reference the gate number that corresponds to the
    current PHASE (phase index + 1), not snapshot.current_gate (the first
    open gate overall) — those can legitimately diverge and pairing them in
    one sentence read as self-contradictory."""
    s = load_snapshot(test_workbook_path)
    sentence = nar.exec_summary_sentence(s, "AMBER")
    assert "Beta" in sentence
    assert "Gate 2" in sentence  # Beta is phase index 1 -> Gate 2, matching phase.gate_status


def test_path_forward_boxes_reference_current_and_next_phase(test_workbook_path):
    s = load_snapshot(test_workbook_path)
    boxes = nar.path_forward_boxes(s)
    titles = [title for title, _ in boxes]
    assert any("Beta" in t for t in titles)


def test_fmt_date_handles_none():
    assert nar.fmt_date(None) == "TBD"
    assert nar.fmt_date(date(2024, 1, 1)) == "01-Jan-24"


def test_short_helper_truncates_at_word_boundary():
    text = "This is a moderately long sentence that should be trimmed"
    short = nar._short(text, 20)
    assert len(short) <= 21  # 20 + ellipsis char
    assert not short.endswith(" ")
