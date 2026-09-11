"""Forward-looking analytics derived from the workbook — not a black-box ML
model, but transparent, explainable arithmetic over real task/RAID/budget
data, so every number in the deck can be traced back to a formula and a
source cell. Each function documents exactly what it computes and flags low
confidence when the sample size is thin, rather than overclaiming precision.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date, timedelta

from app.excel_reader import ProjectSnapshot
from app.manual_inputs import ManualInputs, budget_summary


@dataclass
class ScheduleForecast:
    projected_go_live: date
    shift_days: int
    avg_completion_variance_days: float
    overdue_slippage_days: int
    sample_size: int
    confidence: str  # "low" | "medium" | "high"
    direction: str  # "ahead" | "behind" | "on track"


def schedule_forecast(snapshot: ProjectSnapshot) -> ScheduleForecast:
    """Projects a go-live date from actual execution pace instead of the
    static target: tasks finishing before/after their planned due date shift
    the projection earlier/later, and any task already overdue (still open
    past its due date) adds concrete, already-realized slippage on top —
    exactly the 'finished early -> pulls the date in, delay -> pushes it out'
    behavior requested. Not a committed schedule — a directional estimate
    that gets more reliable as more tasks close (see `confidence`)."""
    completed = [
        t for t in snapshot.tasks
        if t.status.strip().lower() == "complete" and t.due and t.last_updated
    ]
    variances = [(t.last_updated - t.due).days for t in completed]
    avg_variance = (sum(variances) / len(variances)) if variances else 0.0

    today = date.today()
    overdue_open = [
        t for t in snapshot.tasks
        if t.status.strip().lower() != "complete" and t.due and t.due < today
    ]
    overdue_days = max((today - min(t.due for t in overdue_open)).days, 0) if overdue_open else 0

    shift_days = round(avg_variance) + overdue_days
    projected = snapshot.meta.target_go_live + timedelta(days=shift_days)

    if len(completed) < 8:
        confidence = "low"
    elif len(completed) < 25:
        confidence = "medium"
    else:
        confidence = "high"

    direction = "ahead" if shift_days < 0 else ("behind" if shift_days > 0 else "on track")

    return ScheduleForecast(
        projected_go_live=projected,
        shift_days=shift_days,
        avg_completion_variance_days=avg_variance,
        overdue_slippage_days=overdue_days,
        sample_size=len(completed),
        confidence=confidence,
        direction=direction,
    )


@dataclass
class BudgetPace:
    remaining: float
    elapsed_pct: float
    run_rate_projection: float | None
    divergence_vs_pm_forecast: float | None
    has_enough_data: bool


def budget_pace(snapshot: ProjectSnapshot, manual: ManualInputs) -> BudgetPace:
    """Cross-checks the PM's manual forecast-at-completion against a simple
    run-rate projection (spend so far, extrapolated across the full project
    timeline at the same rate) — the same 'are we pacing consistently with
    what we've committed to' check a portfolio analyst would do by hand.
    Skips the run-rate figure entirely once there's too little elapsed
    project time for a rate to mean anything, rather than showing a
    misleadingly precise number off a near-zero denominator."""
    bsum = budget_summary(manual.budget)
    remaining = bsum["approved"] - bsum["spent"]

    total_days = (snapshot.meta.target_go_live - snapshot.meta.start).days
    elapsed_days = (date.today() - snapshot.meta.start).days
    elapsed_pct = max(0.0, min(elapsed_days / total_days, 1.0)) if total_days > 0 else 0.0

    has_enough_data = elapsed_pct >= 0.03  # need at least ~3% of the timeline elapsed for a rate to be meaningful
    if has_enough_data:
        run_rate_projection = bsum["spent"] / elapsed_pct
        divergence = run_rate_projection - bsum["forecast"]
    else:
        run_rate_projection = None
        divergence = None

    return BudgetPace(
        remaining=remaining,
        elapsed_pct=elapsed_pct,
        run_rate_projection=run_rate_projection,
        divergence_vs_pm_forecast=divergence,
        has_enough_data=has_enough_data,
    )


@dataclass
class TrendDeltas:
    pct_complete_delta: float | None
    open_raid_delta: int | None
    high_severity_raid_delta: int | None
    days_to_go_live_delta: int | None


def trend_deltas(current_kpi: dict, previous_kpi: dict | None) -> TrendDeltas:
    """Report-over-report movement, made possible because every generation
    stores a small KPI snapshot — the same way a person would flip back to
    last month's report and compare notes, just automated."""
    if not previous_kpi:
        return TrendDeltas(None, None, None, None)
    return TrendDeltas(
        pct_complete_delta=current_kpi.get("pct_complete", 0) - previous_kpi.get("pct_complete", 0),
        open_raid_delta=current_kpi.get("open_raid", 0) - previous_kpi.get("open_raid", 0),
        high_severity_raid_delta=current_kpi.get("high_severity_raid", 0) - previous_kpi.get("high_severity_raid", 0),
        days_to_go_live_delta=current_kpi.get("days_to_go_live", 0) - previous_kpi.get("days_to_go_live", 0),
    )


def workload_imbalance(snapshot: ProjectSnapshot) -> str | None:
    """Flags a stakeholder carrying a disproportionate share of open work —
    a simple outlier check (more than double the average open-task load
    across everyone with any assignment), the kind of thing a PM would
    eventually notice manually but easy to miss week to week."""
    loads = [s for s in snapshot.stakeholder_loads if s.open_r > 0]
    if len(loads) < 3:
        return None
    avg = sum(s.open_r for s in loads) / len(loads)
    if avg <= 0:
        return None
    top = max(loads, key=lambda s: s.open_r)
    if top.open_r >= avg * 2 and top.open_r >= 5:
        return f"{top.role} holds {top.open_r} open tasks — more than {top.open_r / avg:.1f}x the average load"
    return None


def gate_slip_risk(snapshot: ProjectSnapshot) -> str | None:
    """Flags the current gate if it's due soon but still has multiple open
    checklist items — a simple pace check (days remaining vs. items
    remaining), not a full critical-path simulation."""
    gate = snapshot.current_gate
    if not gate or not gate.due:
        return None
    days_left = (gate.due - date.today()).days
    open_items = [i for i in gate.items if i.status.strip().lower() != "complete"]
    if 0 <= days_left <= 14 and len(open_items) >= 2:
        return f"Gate {gate.number} due in {days_left}d with {len(open_items)} checklist items still open"
    return None
