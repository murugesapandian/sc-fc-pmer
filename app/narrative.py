"""Deterministic, rule-based narrative derived from the workbook snapshot.

No free-text generation: every string here is built from a template plus
real field values, so the output is auditable and reproducible. Judgment
calls that cannot be derived (budget commentary, decisions, objective
overrides) come from manual_inputs.py and are passed through untouched.
"""
from __future__ import annotations

import re
from datetime import date, timedelta

from app.excel_reader import ProjectSnapshot, PhaseRow, WorkstreamRow, RaidItem, Task
from app.manual_inputs import ManualInputs
from app.analytics import ScheduleForecast, BudgetPace, TrendDeltas

_DATE_FMT = "%d-%b-%y"
_PHASE_LABEL_RE = re.compile(r"^\d+\.\s*")


def short_phase_name(name: str) -> str:
    return _PHASE_LABEL_RE.sub("", name)


def fmt_date(d: date | None) -> str:
    return d.strftime(_DATE_FMT) if d else "TBD"


def _short(text: str, max_len: int) -> str:
    """Executive-brief phrasing: a clause, not a paragraph. Cuts at the last
    word boundary within the limit rather than mid-word, so bullets read as
    intentionally short rather than clipped."""
    text = text.strip()
    if len(text) <= max_len:
        return text
    cut = text[:max_len].rsplit(" ", 1)[0]
    return (cut or text[:max_len]).rstrip(",;:.-") + "…"


def program_status(snapshot: ProjectSnapshot, manual: ManualInputs) -> str:
    if manual.program_status_override:
        return manual.program_status_override.upper()

    overdue_high_raid = any(r.days_overdue > 0 for r in snapshot.high_severity_open_raid)
    if snapshot.kpis.overdue_tasks > 0 or overdue_high_raid:
        return "RED"

    gate = snapshot.current_gate
    gate_due_soon = bool(gate and gate.due and 0 <= (gate.due - date.today()).days <= 14 and gate.is_open)
    if snapshot.kpis.at_risk > 0 or snapshot.kpis.high_severity_raid >= 1 or gate_due_soon:
        return "AMBER"

    return "GREEN"


def workstream_health(ws: WorkstreamRow, snapshot: ProjectSnapshot) -> str:
    if ws.overdue > 0:
        return "Red"
    high_risk_here = any(
        r.is_open and r.severity.strip().lower() == "high" and r.workstream == ws.name
        for r in snapshot.raid_items
    )
    if ws.at_risk > 0 or high_risk_here:
        return "Amber"
    return "Green"


def workstream_focus(ws: WorkstreamRow, snapshot: ProjectSnapshot, override: str | None) -> str:
    """Reasons about the workstream's position in the *project's own* phase
    order, not the wall-clock date — the sample plan's dates run into the
    future, so comparing task.start to date.today() would call every
    workstream 'not yet active' before the project's start date arrives."""
    if override:
        return override

    ws_tasks = [t for t in snapshot.tasks if t.workstream == ws.name]
    if not ws_tasks:
        return "No tasks in this workstream"

    phase_order = {p.name: i for i, p in enumerate(snapshot.phases)}
    cur_phase = snapshot.current_phase
    cur_idx = phase_order.get(cur_phase.name, 0) if cur_phase else 0
    earliest_idx = min(phase_order.get(t.phase, 0) for t in ws_tasks)

    if earliest_idx - cur_idx > 1:
        return "Not yet active"

    open_tasks = [t for t in ws_tasks if t.status.strip().lower() != "complete"]
    open_tasks.sort(key=lambda t: (phase_order.get(t.phase, 0), t.due or date.max))
    if open_tasks:
        return _short(open_tasks[0].deliverable, 34)
    return "All tasks complete"


def exec_summary_sentence(snapshot: ProjectSnapshot, status: str) -> str:
    phase = snapshot.current_phase
    phase_bit = f"{phase.name} {phase.pct_complete:.0%} complete" if phase else "phase status unavailable"

    # Deliberately the gate tied to the *current phase* (phase index + 1),
    # not snapshot.current_gate (the first open gate overall) — a gate's
    # checklist can legitimately finish ahead of that phase's full task list,
    # and pairing "Initiation 71% complete" with an unrelated "Gate 2" in the
    # same sentence read as contradictory rather than as two distinct,
    # correctly-independent trackers.
    gate_bit = ""
    if phase and phase in snapshot.phases and phase.gate_status:
        gate_num = snapshot.phases.index(phase) + 1
        gate_bit = f"Gate {gate_num} {phase.gate_status}"

    return (
        f"Overall program status: {status}. {phase_bit}"
        f"{'; ' + gate_bit if gate_bit else ''}. "
        f"{snapshot.days_to_go_live} days to target go-live ({fmt_date(snapshot.meta.target_go_live)})."
    )


def schedule_forecast_line(forecast: ScheduleForecast) -> str:
    if forecast.direction == "on track":
        verb = "tracking to"
    elif forecast.direction == "ahead":
        verb = f"trending {abs(forecast.shift_days)}d ahead of"
    else:
        verb = f"trending {forecast.shift_days}d behind"
    confidence_note = "" if forecast.confidence == "high" else f" (low sample — {forecast.sample_size} tasks closed)" if forecast.confidence == "low" else ""
    return f"AI schedule forecast: {verb} target — projected go-live {fmt_date(forecast.projected_go_live)}{confidence_note}"


def budget_pace_line(pace: BudgetPace, approved: float) -> str | None:
    if not pace.has_enough_data:
        return None
    div_pct = (pace.divergence_vs_pm_forecast / approved) if approved else 0
    if abs(div_pct) < 0.02:
        return f"AI run-rate check: current spend pace lines up with the ${approved/1_000_000:.1f}M forecast"
    direction = "above" if pace.divergence_vs_pm_forecast > 0 else "below"
    return (
        f"AI run-rate check: at the current spend pace, full-project cost projects "
        f"${pace.run_rate_projection/1_000_000:.1f}M — {abs(div_pct):.0%} {direction} the PM forecast"
    )


def trend_line(deltas: TrendDeltas) -> str | None:
    if deltas.open_raid_delta is None:
        return None
    parts = []
    if deltas.open_raid_delta != 0:
        arrow = "up" if deltas.open_raid_delta > 0 else "down"
        parts.append(f"open RAID {arrow} {abs(deltas.open_raid_delta)} since last report")
    if deltas.days_to_go_live_delta is not None and deltas.days_to_go_live_delta != 0:
        moved = "later" if deltas.days_to_go_live_delta > 0 else "earlier"
        parts.append(f"schedule pace pulled go-live {abs(deltas.days_to_go_live_delta)}d {moved}")
    return "Since last report: " + "; ".join(parts) if parts else None


def accomplished_this_period(snapshot: ProjectSnapshot, since: date) -> list[str]:
    """Deliverable names only, not the full task sentence — the master plan
    has the detail; the exec summary needs the headline. Capped to 4: this is
    a 'what changed' pointer, not a status report of every task."""
    done = [
        t for t in snapshot.tasks
        if t.status.strip().lower() == "complete" and t.last_updated and t.last_updated >= since
    ]
    done.sort(key=lambda t: t.last_updated, reverse=True)
    seen, bullets = set(), []
    for t in done:
        if t.deliverable in seen:
            continue
        seen.add(t.deliverable)
        bullets.append(_short(t.deliverable, 50))
        if len(bullets) == 4:
            break
    if len(done) > len(bullets):
        bullets[-1] = f"{bullets[-1]} (+{len(done) - len(bullets) + 1} more this period)"
    return bullets or ["No tasks completed in this reporting window."]


def watch_items(snapshot: ProjectSnapshot, extra_flags: list[str] | None = None) -> list[str]:
    """Pointer-style, not the full RAID description — the Top Risks slide
    already carries the full detail table; this is just 'what to ask about'.
    `extra_flags` folds in deterministic analytics call-outs (gate pace,
    workload concentration) alongside the RAID-sourced items."""
    items = sorted(snapshot.high_severity_open_raid, key=lambda r: r.score, reverse=True)[:2]
    out = [f"{r.id} — {_short(r.description, 42)} (due {fmt_date(r.due_date)})" for r in items]
    out.extend(extra_flags or [])
    return out or ["No open High-severity RAID items."]


def phase_commentary(snapshot: ProjectSnapshot, forecast: ScheduleForecast | None = None) -> list[str]:
    """Tied to the exact same numbers driving the stacked-bar chart
    (per-phase complete/at-risk/total counts) rather than calendar dates —
    a date comparison silently goes stale whenever 'today' isn't inside the
    plan's own date window, which is what made this section look frozen
    while the chart kept updating. Also flags overdue open tasks directly,
    since a phase whose due date has already passed with work still open is
    worth calling out regardless of which phase is nominally 'current'."""
    cur = snapshot.current_phase
    cur_idx = snapshot.phases.index(cur) if cur and cur in snapshot.phases else -1
    today = date.today()
    out = []

    for i, p in enumerate(snapshot.phases):
        if i < cur_idx and p.pct_complete < 1.0:
            out.append(f"{short_phase_name(p.name)}: still open behind the current phase ({p.pct_complete:.0%} complete)")
        elif p.due and p.due < today and p.pct_complete < 1.0:
            out.append(f"{short_phase_name(p.name)}: past its {fmt_date(p.due)} due date at {p.pct_complete:.0%} complete")

    for p in snapshot.phases:
        if p.at_risk > 0:
            out.append(f"{short_phase_name(p.name)}: {p.at_risk} task(s) at risk / blocked")

    if cur:
        out.append(f"Current — {short_phase_name(cur.name)}: {cur.complete}/{cur.tasks} tasks, Gate {cur_idx + 1} {cur.gate_status}")

    if forecast:
        out.append(schedule_forecast_line(forecast))

    return out[:4] or ["All phases tracking to plan."]


def top_risks(snapshot: ProjectSnapshot) -> list[RaidItem]:
    """High/Medium severity open items, ranked by score — matches the deck's
    'Top risks and issues' table (Low-severity items are called out as not
    shown, in the slide's footnote)."""
    shown = [r for r in snapshot.open_raid_items if r.severity.strip().lower() != "low"]
    return sorted(shown, key=lambda r: r.score, reverse=True)


def severity_distribution(snapshot: ProjectSnapshot) -> dict[str, int]:
    counts = {"High": 0, "Medium": 0, "Low": 0}
    for r in snapshot.open_raid_items:
        sev = r.severity.strip().capitalize()
        if sev in counts:
            counts[sev] += 1
    return counts


def path_forward_boxes(snapshot: ProjectSnapshot) -> list[tuple[str, list[str]]]:
    """Deliverable names, not task-activity sentences — three short bullets
    per box is a scan; four-plus long sentences is a memo."""
    boxes: list[tuple[str, list[str]]] = []

    cur = snapshot.current_phase
    if cur:
        remaining = [t for t in snapshot.tasks if t.phase == cur.name and t.status.strip().lower() != "complete"]
        remaining.sort(key=lambda t: (t.due or date.max))
        bullets = [_short(t.deliverable, 46) for t in remaining[:3]] or ["No remaining open tasks in this phase."]
        boxes.append((f"Close out {short_phase_name(cur.name)}", bullets))

    nxt = snapshot.next_phase
    if nxt:
        upcoming = [t for t in snapshot.tasks if t.phase == nxt.name]
        upcoming.sort(key=lambda t: (t.start or date.max))
        bullets = [_short(t.deliverable, 46) for t in upcoming[:3]] or ["No tasks scheduled yet."]
        boxes.append((f"Launch {short_phase_name(nxt.name)}", bullets))

    high_raid = sorted(snapshot.high_severity_open_raid, key=lambda r: r.score, reverse=True)
    bullets = [f"{r.id}: {_short(r.description, 40)}" for r in high_raid[:3]] or ["No open High-severity items to retire."]
    boxes.append(("Retire High-severity RAID items", bullets))

    boxes.append((
        "Continue automated status reporting",
        ["Weekly status drafted from the tracker", "Gate checklists audited automatically",
         "Meeting actions logged for PM confirmation"],
    ))
    return boxes


def next_report_out_line(snapshot: ProjectSnapshot, trend: TrendDeltas | None = None) -> str:
    nxt = snapshot.next_phase or snapshot.current_phase
    high_count = len(snapshot.high_severity_open_raid)
    base = (
        f"Next report-out (Month {snapshot.reporting_month_number + 1}): "
        f"{nxt.name if nxt else 'next phase'} progressing, "
        f"{high_count} open High-severity RAID item(s) tracked to closure, budget forecast reviewed."
    )
    if trend:
        line = trend_line(trend)
        if line:
            return f"{base} {line}."
    return base
