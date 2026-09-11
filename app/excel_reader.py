"""Reads Greenfield_SCM_GoLive_Tracking_Playbook.xlsx into structured, typed data.

Only ever opens the workbook read-only with cached (data_only) formula values —
never writes back to it.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.worksheet.worksheet import Worksheet


def _to_date(value: Any) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return None


def _to_number(value: Any, default: float = 0.0) -> float:
    if isinstance(value, (int, float)):
        return value
    return default


def _rows_as_dicts(ws: Worksheet, header_row: int, start_row: int, end_row: int) -> list[dict]:
    headers = [c.value for c in ws[header_row]]
    out = []
    for r in range(start_row, end_row + 1):
        values = [c.value for c in ws[r]]
        if all(v is None for v in values):
            continue
        row = {headers[i]: values[i] for i in range(len(headers)) if headers[i] is not None}
        out.append(row)
    return out


@dataclass
class ProjectMeta:
    name: str
    start: date
    target_go_live: date


@dataclass
class KpiTotals:
    total_tasks: int
    complete: int
    pct_complete: float
    in_progress: int
    at_risk: int
    overdue_tasks: int
    open_raid: int
    high_severity_raid: int
    open_actions: int


@dataclass
class PhaseRow:
    name: str
    tasks: int
    complete: int
    pct_complete: float
    in_progress: int
    at_risk: int
    overdue: int
    start: date | None
    due: date | None
    gate_status: str


@dataclass
class WorkstreamRow:
    name: str
    tasks: int
    complete: int
    pct_complete: float
    in_progress: int
    at_risk: int
    overdue: int


@dataclass
class StakeholderLoad:
    role: str
    responsible_for: int
    accountable_for: int
    open_r: int
    overdue_r: int
    open_raid_owned: int
    open_actions_owned: int


@dataclass
class Stakeholder:
    id: str
    role: str
    name: str
    organisation: str
    workstream: str
    responsibilities: str
    email: str | None
    escalation_contact: str | None


@dataclass
class Objective:
    id: str
    text: str
    criteria: str
    accountable_id: str
    tasks_linked: int
    tasks_complete: int
    pct_complete: float
    status: str
    override_note: str | None = None


@dataclass
class GateItem:
    gate: str
    text: str
    owner_id: str | None
    status: str


@dataclass
class Gate:
    number: int
    name: str
    status_text: str
    done_fraction: str
    items: list[GateItem] = field(default_factory=list)
    due: date | None = None

    @property
    def is_open(self) -> bool:
        return self.status_text.strip().upper().startswith("OPEN")


@dataclass
class RaidItem:
    id: str
    type: str
    description: str
    linked_task: str | None
    workstream: str
    owner_id: str
    raised_on: date | None
    probability: int
    impact: int
    score: int
    severity: str
    mitigation: str
    due_date: date | None
    status: str
    days_overdue: int
    source: str | None
    last_updated: date | None

    @property
    def is_open(self) -> bool:
        return self.status.strip().lower() != "closed"


@dataclass
class ActionItem:
    id: str
    action: str
    source_meeting: str | None
    owner_id: str
    linked_task: str | None
    due_date: date | None
    status: str
    days_overdue: int
    closed_on: date | None
    notes: str | None

    @property
    def is_open(self) -> bool:
        return self.status.strip().lower() not in ("complete", "closed")


@dataclass
class Task:
    id: str
    phase: str
    workstream: str
    objective_id: str | None
    deliverable: str
    activity: str
    responsible_id: str | None
    accountable_id: str | None
    start: date | None
    due: date | None
    duration: int
    status: str
    pct_complete: float
    gate: str | None
    days_overdue: int
    health: str | None
    last_updated: date | None
    notes: str | None


@dataclass
class ProjectSnapshot:
    meta: ProjectMeta
    kpis: KpiTotals
    phases: list[PhaseRow]
    workstreams: list[WorkstreamRow]
    stakeholder_loads: list[StakeholderLoad]
    stakeholders: dict[str, Stakeholder]
    objectives: list[Objective]
    gates: list[Gate]
    raid_items: list[RaidItem]
    actions: list[ActionItem]
    tasks: list[Task]

    def stakeholder_name(self, sh_id: str | None) -> str:
        if not sh_id:
            return "Unassigned"
        sh = self.stakeholders.get(sh_id)
        return sh.role if sh else sh_id

    @property
    def open_raid_items(self) -> list[RaidItem]:
        return [r for r in self.raid_items if r.is_open]

    @property
    def high_severity_open_raid(self) -> list[RaidItem]:
        return [r for r in self.open_raid_items if r.severity.strip().lower() == "high"]

    @property
    def open_actions(self) -> list[ActionItem]:
        return [a for a in self.actions if a.is_open]

    @property
    def current_phase(self) -> PhaseRow | None:
        """The earliest phase that isn't yet 100% complete — mirrors
        `current_gate`'s "first open gate" logic and matches the plan's own
        gated methodology (a phase can't really be 'next' while an earlier
        one still has open work). Deliberately data-driven rather than
        date-based: a calendar comparison against `start <= today <= due`
        silently breaks whenever 'today' doesn't fall inside the plan's own
        date window (e.g. a plan dated entirely in the future), always
        collapsing to phase 1 regardless of how much has actually been
        completed — which is exactly what was making the phase-progress
        commentary look frozen while the chart kept moving."""
        for p in self.phases:
            if p.pct_complete < 1.0:
                return p
        return self.phases[-1] if self.phases else None

    @property
    def next_phase(self) -> PhaseRow | None:
        cur = self.current_phase
        if not cur or cur not in self.phases:
            return None
        idx = self.phases.index(cur)
        return self.phases[idx + 1] if idx + 1 < len(self.phases) else None

    @property
    def current_gate(self) -> Gate | None:
        for g in self.gates:
            if g.is_open:
                return g
        return self.gates[-1] if self.gates else None

    @property
    def days_to_go_live(self) -> int:
        return (self.meta.target_go_live - date.today()).days

    @property
    def reporting_month_number(self) -> int:
        today = date.today()
        start = self.meta.start
        months = (today.year - start.year) * 12 + (today.month - start.month) + 1
        return max(1, months)


def _parse_dashboard_meta(ws: Worksheet) -> tuple[ProjectMeta, list[str], list[StakeholderLoad]]:
    """Only the genuinely raw, PM-authored inputs: project name/start, the
    ordered list of phase names, and the stakeholder-load table. Everything
    else Dashboard *computes* (phase/workstream rollups, the KPI summary
    row) is deliberately NOT read here — see the module docstring note below
    `_derive_phase_rows` for why."""
    name = ws["B4"].value
    start = _to_date(ws["B5"].value)
    meta_partial = ProjectMeta(name=name, start=start, target_go_live=None)  # target filled in from tasks later

    phase_names = []
    r = 15
    while True:
        val = ws.cell(row=r, column=1).value
        if not val or val == "Total":
            break
        phase_names.append(val)
        r += 1

    stakeholder_loads = []
    r = 47
    while True:
        row = [c.value for c in ws[r]]
        if not row[0]:
            break
        stakeholder_loads.append(
            StakeholderLoad(
                role=row[0],
                responsible_for=int(_to_number(row[1])),
                accountable_for=int(_to_number(row[2])),
                open_r=int(_to_number(row[3])),
                overdue_r=int(_to_number(row[4])),
                open_raid_owned=int(_to_number(row[5])),
                open_actions_owned=int(_to_number(row[6])),
            )
        )
        r += 1

    return meta_partial, phase_names, stakeholder_loads


def _task_status(t: "Task") -> str:
    return t.status.strip().lower()


def _is_task_overdue(t: "Task", today: date) -> bool:
    return _task_status(t) != "complete" and bool(t.due) and t.due < today


def _derive_phase_rows(phase_names: list[str], tasks: list["Task"]) -> list[PhaseRow]:
    """Computes every phase-level number directly from Master Plan task rows
    instead of trusting Dashboard's own per-phase formulas.

    Found live in this workbook: several of Dashboard's phase-table cells
    (the 'Complete' count, the 'Gate Status' text) had been overwritten with
    plain literal values at some point — e.g. a literal `7` and the literal
    string `'PASSED'` sitting where a `=COUNTIF(...)` / `='Gate Checklists'!F5`
    formula used to be. A formula cell can be silently clobbered like that at
    any time (a stray paste, a manual edit while testing) and there's no way
    to detect it from the cached value alone — it just quietly stops tracking
    Master Plan. Deriving these numbers ourselves from the one place that's
    genuinely raw input (the Status column PMs actually type into) makes the
    chart, the KPI tiles and this commentary agree by construction, not by
    hoping every formula in Dashboard stayed intact."""
    by_phase: dict[str, list[Task]] = {}
    for t in tasks:
        by_phase.setdefault(t.phase, []).append(t)

    today = date.today()
    rows = []
    for name in phase_names:
        group = by_phase.get(name, [])
        total = len(group)
        complete = sum(1 for t in group if _task_status(t) == "complete")
        in_progress = sum(1 for t in group if _task_status(t) == "in progress")
        at_risk = sum(1 for t in group if _task_status(t) in ("at risk", "blocked"))
        overdue = sum(1 for t in group if _is_task_overdue(t, today))
        starts = [t.start for t in group if t.start]
        dues = [t.due for t in group if t.due]
        rows.append(
            PhaseRow(
                name=name,
                tasks=total,
                complete=complete,
                pct_complete=(complete / total) if total else 0.0,
                in_progress=in_progress,
                at_risk=at_risk,
                overdue=overdue,
                start=min(starts) if starts else None,
                due=max(dues) if dues else None,
                gate_status="",  # filled in from the matching Gate once gates are parsed
            )
        )
    return rows


def _derive_workstream_rows(tasks: list["Task"]) -> list[WorkstreamRow]:
    """Same reasoning as `_derive_phase_rows`, grouped by workstream instead
    of phase — feeds the Workstream Health slide's table and doughnut."""
    by_ws: dict[str, list[Task]] = {}
    for t in tasks:
        if t.workstream:
            by_ws.setdefault(t.workstream, []).append(t)

    today = date.today()
    rows = []
    for name, group in by_ws.items():
        total = len(group)
        if total == 0:
            continue
        complete = sum(1 for t in group if _task_status(t) == "complete")
        in_progress = sum(1 for t in group if _task_status(t) == "in progress")
        at_risk = sum(1 for t in group if _task_status(t) in ("at risk", "blocked"))
        overdue = sum(1 for t in group if _is_task_overdue(t, today))
        rows.append(
            WorkstreamRow(
                name=name,
                tasks=total,
                complete=complete,
                pct_complete=(complete / total) if total else 0.0,
                in_progress=in_progress,
                at_risk=at_risk,
                overdue=overdue,
            )
        )
    return rows


def _derive_kpis(tasks: list["Task"], raid_items: list["RaidItem"], actions: list["ActionItem"]) -> KpiTotals:
    """The headline numbers on the KPI tiles, derived the same way as the
    phase/workstream rollups above and for the same reason — so they always
    agree with the phase chart instead of depending on a separate Dashboard
    summary row that can drift from it."""
    today = date.today()
    total = len(tasks)
    complete = sum(1 for t in tasks if _task_status(t) == "complete")
    in_progress = sum(1 for t in tasks if _task_status(t) == "in progress")
    at_risk = sum(1 for t in tasks if _task_status(t) in ("at risk", "blocked"))
    overdue = sum(1 for t in tasks if _is_task_overdue(t, today))
    open_raid = sum(1 for r in raid_items if r.is_open)
    high_severity_raid = sum(1 for r in raid_items if r.is_open and r.severity.strip().lower() == "high")
    open_actions = sum(1 for a in actions if a.is_open)
    return KpiTotals(
        total_tasks=total,
        complete=complete,
        pct_complete=(complete / total) if total else 0.0,
        in_progress=in_progress,
        at_risk=at_risk,
        overdue_tasks=overdue,
        open_raid=open_raid,
        high_severity_raid=high_severity_raid,
        open_actions=open_actions,
    )


def _parse_stakeholders(ws: Worksheet) -> dict[str, Stakeholder]:
    rows = _rows_as_dicts(ws, header_row=4, start_row=5, end_row=ws.max_row)
    out = {}
    for row in rows:
        sid = row.get("Stakeholder ID")
        if not sid:
            continue
        out[sid] = Stakeholder(
            id=sid,
            role=row.get("Role (used in dropdowns)") or "",
            name=row.get("Name / Title") or "",
            organisation=row.get("Organisation") or "",
            workstream=row.get("Workstream") or "",
            responsibilities=row.get("Responsibilities") or "",
            email=row.get("Email"),
            escalation_contact=row.get("Escalation Contact"),
        )
    return out


def _parse_objectives(ws: Worksheet) -> list[Objective]:
    rows = _rows_as_dicts(ws, header_row=4, start_row=5, end_row=ws.max_row)
    out = []
    for row in rows:
        oid = row.get("Objective ID")
        if not oid:
            continue
        out.append(
            Objective(
                id=oid,
                text=row.get("Objective") or "",
                criteria=row.get("Success Criteria / Measure") or "",
                accountable_id=row.get("Accountable (ID)") or "",
                tasks_linked=int(_to_number(row.get("Tasks Linked"))),
                tasks_complete=int(_to_number(row.get("Tasks Complete"))),
                pct_complete=_to_number(row.get("% Complete")),
                status=row.get("Status") or "",
            )
        )
    return out


_GATE_ITEM_RE = re.compile(r"^G\d+$")


def _parse_gates(ws: Worksheet) -> list[Gate]:
    gates: list[Gate] = []
    current: Gate | None = None
    for r in range(5, ws.max_row + 1):
        row = [c.value for c in ws[r]]
        col_a, col_b, col_e = row[0], row[1], row[4]
        if col_a and str(col_a).startswith("Gate ") and col_e == "Gate status:":
            m = re.match(r"Gate (\d+)", str(col_a))
            number = int(m.group(1)) if m else len(gates) + 1
            current = Gate(number=number, name=str(col_a), status_text="", done_fraction="")
            gates.append(current)
        elif col_a and _GATE_ITEM_RE.match(str(col_a)) and current is not None:
            current.items.append(GateItem(gate=col_a, text=col_b or "", owner_id=row[2], status=row[4] or ""))

    # Derived from the raw per-item statuses just parsed above, mirroring
    # the sheet's own formula (`=IF(COUNTIFS(...Open...In Progress...)=0,
    # "PASSED", "OPEN — N item(s)")`) — computed by us instead of trusting
    # that cached cell, for the same reason as the phase/workstream rollups.
    for g in gates:
        total = len(g.items)
        done = sum(1 for i in g.items if i.status.strip().lower() in ("complete", "n/a"))
        pending = sum(1 for i in g.items if i.status.strip().lower() not in ("complete", "n/a"))
        g.done_fraction = f"{done} / {total} done"
        g.status_text = "PASSED" if pending == 0 else f"OPEN — {pending} item(s)"

    return gates


def _parse_raid(ws: Worksheet) -> list[RaidItem]:
    rows = _rows_as_dicts(ws, header_row=4, start_row=5, end_row=ws.max_row)
    out = []
    for row in rows:
        rid = row.get("RAID ID")
        if not rid:
            continue
        out.append(
            RaidItem(
                id=rid,
                type=row.get("Type") or "",
                description=row.get("Description") or "",
                linked_task=row.get("Linked Task"),
                workstream=row.get("Workstream") or "",
                owner_id=row.get("Owner (ID)") or "",
                raised_on=_to_date(row.get("Raised On")),
                probability=int(_to_number(row.get("Probability (1-5)"))),
                impact=int(_to_number(row.get("Impact (1-5)"))),
                score=int(_to_number(row.get("Score"))),
                severity=row.get("Severity") or "",
                mitigation=row.get("Mitigation / Response") or "",
                due_date=_to_date(row.get("Due Date")),
                status=row.get("Status") or "",
                days_overdue=int(_to_number(row.get("Days Overdue"))),
                source=row.get("Source"),
                last_updated=_to_date(row.get("Last Updated")),
            )
        )
    return out


def _parse_actions(ws: Worksheet) -> list[ActionItem]:
    rows = _rows_as_dicts(ws, header_row=4, start_row=5, end_row=ws.max_row)
    out = []
    for row in rows:
        aid = row.get("Action ID")
        if not aid:
            continue
        out.append(
            ActionItem(
                id=aid,
                action=row.get("Action") or "",
                source_meeting=row.get("Source Meeting / Date"),
                owner_id=row.get("Owner (ID)") or "",
                linked_task=row.get("Linked Task"),
                due_date=_to_date(row.get("Due Date")),
                status=row.get("Status") or "",
                days_overdue=int(_to_number(row.get("Days Overdue"))),
                closed_on=_to_date(row.get("Closed On")),
                notes=row.get("Notes"),
            )
        )
    return out


def _parse_master_plan(ws: Worksheet) -> list[Task]:
    rows = _rows_as_dicts(ws, header_row=4, start_row=5, end_row=ws.max_row)
    out = []
    for row in rows:
        tid = row.get("Task ID")
        if not tid:
            continue
        out.append(
            Task(
                id=tid,
                phase=row.get("Phase") or "",
                workstream=row.get("Workstream") or "",
                objective_id=row.get("Objective ID"),
                deliverable=row.get("Deliverable") or "",
                activity=row.get("Task / Activity") or "",
                responsible_id=row.get("Responsible (ID)"),
                accountable_id=row.get("Accountable (ID)"),
                start=_to_date(row.get("Start")),
                due=_to_date(row.get("Due")),
                duration=int(_to_number(row.get("Duration (days)"))),
                status=row.get("Status") or "",
                pct_complete=_to_number(row.get("% Complete")),
                gate=row.get("Gate"),
                days_overdue=int(_to_number(row.get("Days Overdue"))),
                health=row.get("Health"),
                last_updated=_to_date(row.get("Last Updated")),
                notes=row.get("Notes"),
            )
        )
    return out


def load_snapshot(excel_path: Path) -> ProjectSnapshot:
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    try:
        meta_partial, phase_names, stakeholder_loads = _parse_dashboard_meta(wb["Dashboard"])
        stakeholders = _parse_stakeholders(wb["Stakeholders"])
        objectives = _parse_objectives(wb["Objectives"])
        gates = _parse_gates(wb["Gate Checklists"])
        raid_items = _parse_raid(wb["RAID Log"])
        actions = _parse_actions(wb["Action Log"])
        tasks = _parse_master_plan(wb["Master Plan"])
    finally:
        wb.close()

    # Everything below is derived from the raw sheets above (tasks, gate
    # items, RAID/action rows) rather than from Dashboard's own rollup
    # formulas — see `_derive_phase_rows` for why.
    phases = _derive_phase_rows(phase_names, tasks)
    workstreams = _derive_workstream_rows(tasks)
    kpis = _derive_kpis(tasks, raid_items, actions)

    gate_by_number = {g.number: g for g in gates}
    for i, phase in enumerate(phases):
        gate = gate_by_number.get(i + 1)
        if gate:
            phase.gate_status = gate.status_text
            gate.due = phase.due

    due_dates = [t.due for t in tasks if t.due]
    target_go_live = max(due_dates) if due_dates else None
    meta = ProjectMeta(name=meta_partial.name, start=meta_partial.start, target_go_live=target_go_live)

    return ProjectSnapshot(
        meta=meta,
        kpis=kpis,
        phases=phases,
        workstreams=workstreams,
        stakeholder_loads=stakeholder_loads,
        stakeholders=stakeholders,
        objectives=objectives,
        gates=gates,
        raid_items=raid_items,
        actions=actions,
        tasks=tasks,
    )
