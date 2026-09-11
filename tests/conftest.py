"""Shared fixtures: a minimal but structurally faithful synthetic workbook
(same sheet/column layout excel_reader.py expects, tiny data set) so tests
run fast and don't depend on the real, constantly-changing sample workbook.
"""
from __future__ import annotations

from datetime import datetime

import openpyxl
import pytest


def _write_headers(ws, row: int, headers: list[str]):
    for i, h in enumerate(headers, start=1):
        ws.cell(row=row, column=i, value=h)


def build_test_workbook(path):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # --- Dashboard ---
    ws = wb.create_sheet("Dashboard")
    ws["B4"] = "Test Project — Widget Rollout"
    ws["B5"] = datetime(2024, 1, 1)
    phase_names = ["1. Alpha", "2. Beta"]
    for i, name in enumerate(phase_names):
        ws.cell(row=15 + i, column=1, value=name)
    ws.cell(row=17, column=1, value="Total")
    _write_headers(ws, 46, ["Stakeholder", "Responsible for", "Accountable for", "Open (R)", "Overdue (R)", "Open RAID owned", "Open actions owned"])
    ws.cell(row=47, column=1, value="Project Manager")
    for col, val in enumerate([3, 1, 1, 0, 0, 0], start=2):
        ws.cell(row=47, column=col, value=val)

    # --- Stakeholders ---
    ws = wb.create_sheet("Stakeholders")
    _write_headers(ws, 4, ["Stakeholder ID", "Role (used in dropdowns)", "Name / Title", "Organisation", "Workstream", "Responsibilities", "Email", "Escalation Contact"])
    ws.append([])
    ws.cell(row=5, column=1, value="SH-01")
    ws.cell(row=5, column=2, value="Project Manager")
    ws.cell(row=5, column=3, value="Jane Doe")
    ws.cell(row=6, column=1, value="SH-02")
    ws.cell(row=6, column=2, value="Engineering Lead")
    ws.cell(row=6, column=3, value="Sam Lee")

    # --- Objectives ---
    ws = wb.create_sheet("Objectives")
    _write_headers(ws, 4, ["Objective ID", "Objective", "Success Criteria / Measure", "Accountable (ID)", "Tasks Linked", "Tasks Complete", "% Complete", "Status"])
    ws.append(["OBJ-01", "Ship the widget", "Widget ships on time", "SH-01", 4, 2, 0.5, "On Track"])

    # --- Master Plan ---
    ws = wb.create_sheet("Master Plan")
    _write_headers(ws, 4, ["Task ID", "Phase", "Workstream", "Objective ID", "Deliverable", "Task / Activity",
                           "Responsible (ID)", "Accountable (ID)", "Consulted (ID)", "Informed (ID)", "Start", "Due",
                           "Duration (days)", "Status", "% Complete", "Predecessors", "AI Copilot Moment", "Gate",
                           "Evidence / Link", "Days Overdue", "Health", "Last Updated", "Notes"])
    tasks = [
        ("T-001", "1. Alpha", "Engineering", "OBJ-01", "Spec", "Write spec", "SH-01", "SH-01", None, None,
         datetime(2024, 1, 1), datetime(2024, 1, 5), 5, "Complete", 1, None, None, "G1", None, 0, "Done", datetime(2024, 1, 6), None),
        ("T-002", "1. Alpha", "Engineering", "OBJ-01", "Prototype", "Build prototype", "SH-02", "SH-01", None, None,
         datetime(2024, 1, 6), datetime(2024, 1, 12), 6, "Complete", 1, "T-001", None, "G1", None, 0, "Done", datetime(2024, 1, 13), None),
        ("T-003", "2. Beta", "Engineering", "OBJ-01", "Beta build", "Ship beta build", "SH-02", "SH-01", None, None,
         datetime(2024, 1, 13), datetime(2024, 1, 20), 7, "In Progress", 0.5, "T-002", None, "G2", None, 0, "On Track", datetime(2024, 1, 15), None),
        ("T-004", "2. Beta", "Engineering", "OBJ-01", "Beta test", "Run beta test", "SH-02", "SH-01", None, None,
         datetime(2024, 1, 21), datetime(2024, 1, 25), 4, "Not Started", 0, "T-003", None, "G2", None, 0, None, None, None),
    ]
    for t in tasks:
        ws.append(list(t))

    # --- Gate Checklists ---
    ws = wb.create_sheet("Gate Checklists")
    _write_headers(ws, 4, ["Gate", "Checklist Item", "Owner (ID)", "Linked Task(s)", "Status", "Evidence / Link", "Reviewed By", "Review Date", "Comments"])
    ws.cell(row=5, column=1, value="Gate 1 — Alpha complete")
    ws.cell(row=5, column=5, value="Gate status:")
    ws.cell(row=6, column=1, value="G1")
    ws.cell(row=6, column=2, value="Spec approved")
    ws.cell(row=6, column=3, value="SH-01")
    ws.cell(row=6, column=5, value="Complete")
    ws.cell(row=7, column=1, value="G1")
    ws.cell(row=7, column=2, value="Prototype approved")
    ws.cell(row=7, column=3, value="SH-01")
    ws.cell(row=7, column=5, value="Complete")
    ws.cell(row=9, column=1, value="Gate 2 — Beta complete")
    ws.cell(row=9, column=5, value="Gate status:")
    ws.cell(row=10, column=1, value="G2")
    ws.cell(row=10, column=2, value="Beta build shipped")
    ws.cell(row=10, column=3, value="SH-02")
    ws.cell(row=10, column=5, value="Open")
    ws.cell(row=11, column=1, value="G2")
    ws.cell(row=11, column=2, value="Beta test passed")
    ws.cell(row=11, column=3, value="SH-02")
    ws.cell(row=11, column=5, value="Open")

    # --- RAID Log ---
    ws = wb.create_sheet("RAID Log")
    _write_headers(ws, 4, ["RAID ID", "Type", "Description", "Linked Task", "Workstream", "Owner (ID)", "Raised On",
                           "Probability (1-5)", "Impact (1-5)", "Score", "Severity", "Mitigation / Response",
                           "Due Date", "Status", "Days Overdue", "Source", "Last Updated"])
    ws.append(["R-001", "Risk", "Vendor may be late", "T-003", "Engineering", "SH-02", datetime(2024, 1, 2),
               4, 4, 16, "High", "Weekly check-in", datetime(2024, 1, 20), "Open", 0, "Standup", datetime(2024, 1, 10)])
    ws.append(["R-002", "Risk", "Minor tooling gap", "T-004", "Engineering", "SH-02", datetime(2024, 1, 2),
               2, 2, 4, "Low", "Backlog item", datetime(2024, 1, 25), "Open", 0, "Standup", datetime(2024, 1, 10)])

    # --- Action Log ---
    ws = wb.create_sheet("Action Log")
    _write_headers(ws, 4, ["Action ID", "Action", "Source Meeting / Date", "Owner (ID)", "Linked Task", "Due Date",
                           "Status", "Days Overdue", "Closed On", "Notes"])
    ws.append(["A-001", "Confirm vendor SLA", "Standup 02-Jan", "SH-02", "T-003", datetime(2024, 1, 15), "Open", 0, None, None])

    wb.save(path)
    return path


@pytest.fixture
def test_workbook_path(tmp_path):
    path = tmp_path / "test_workbook.xlsx"
    return build_test_workbook(path)


@pytest.fixture
def isolated_app_env(tmp_path, monkeypatch, test_workbook_path):
    """Redirects every path the generation pipeline touches to tmp_path, so
    tests never read or write the real workbook, database, or reports
    folder. Each module that does `from app.config import X` holds its own
    bound copy of X in its namespace, so each one needs patching separately
    — patching app.config.X alone would not reach them."""
    from app.config import TEMPLATE_PPTX_PATH
    from app import generate, hashing, manual_inputs, store

    reports_dir = tmp_path / "reports"
    reports_dir.mkdir()
    manual_inputs_path = tmp_path / "manual_inputs.json"

    monkeypatch.setattr(generate, "EXCEL_PATH", test_workbook_path)
    monkeypatch.setattr(generate, "REPORTS_DIR", reports_dir)
    monkeypatch.setattr(hashing, "MANUAL_INPUTS_PATH", manual_inputs_path)
    monkeypatch.setattr(manual_inputs, "MANUAL_INPUTS_PATH", manual_inputs_path)
    monkeypatch.setattr(store, "DB_PATH", tmp_path / "state.db")

    return {
        "workbook_path": test_workbook_path,
        "reports_dir": reports_dir,
        "manual_inputs_path": manual_inputs_path,
        "template_path": TEMPLATE_PPTX_PATH,
    }
