"""Populates a fresh copy of the Monthly Leadership Report-Out template from a
ProjectSnapshot + ManualInputs. Never opens the template for writing in place —
always Presentation(template_path) then .save(new_output_path).

Generic helpers (set_bullets / set_table_rows / set_chart_data) do most of the
work since the same "N lines of text" / "N rows of data" / "replace chart
series" patterns repeat across the deck. Slide-specific code below maps exact
shape names (confirmed by inspecting the template) to snapshot/manual fields.
"""
from __future__ import annotations

import re
from copy import deepcopy
from datetime import date
from pathlib import Path

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.util import Pt

from app.excel_reader import ProjectSnapshot
from app.manual_inputs import ManualInputs, budget_summary
from app import narrative as nar
from app import layout_guard as lg
from app import analytics

GREEN = RGBColor(0x1E, 0x8E, 0x3E)
AMBER = RGBColor(0xC8, 0x7A, 0x00)
RED = RGBColor(0xC5, 0x22, 0x1F)
BLUE = RGBColor(0x00, 0x53, 0xE2)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
YELLOW = RGBColor(0xFF, 0xC2, 0x20)

_STATUS_COLOR = {
    "green": GREEN, "on track": GREEN, "complete": GREEN, "closed": GREEN,
    "amber": AMBER, "at risk": AMBER, "in progress": AMBER, "open": AMBER, "medium": AMBER,
    "red": RED, "off track": RED, "overdue": RED, "blocked": RED, "high": RED,
}
_STATUS_BG = {
    "green": RGBColor(0xE6, 0xF4, 0xEA), "on track": RGBColor(0xE6, 0xF4, 0xEA),
    "complete": RGBColor(0xE6, 0xF4, 0xEA), "closed": RGBColor(0xE6, 0xF4, 0xEA),
    "amber": RGBColor(0xFD, 0xF1, 0xDE), "at risk": RGBColor(0xFD, 0xF1, 0xDE),
    "in progress": RGBColor(0xFD, 0xF1, 0xDE), "open": RGBColor(0xFD, 0xF1, 0xDE), "medium": RGBColor(0xFD, 0xF1, 0xDE),
    "red": RGBColor(0xFB, 0xE7, 0xE6), "off track": RGBColor(0xFB, 0xE7, 0xE6),
    "overdue": RGBColor(0xFB, 0xE7, 0xE6), "blocked": RGBColor(0xFB, 0xE7, 0xE6), "high": RGBColor(0xFB, 0xE7, 0xE6),
}


def _status_key(text: str) -> str | None:
    """Matches a status word even when it's embedded in a larger cell value
    (e.g. '15 High' for a score+severity cell), not just an exact match —
    otherwise a combined cell like that silently gets no color at all."""
    lower = text.strip().lower()
    if lower in _STATUS_COLOR:
        return lower
    for key in sorted(_STATUS_COLOR, key=len, reverse=True):
        if re.search(rf"\b{re.escape(key)}\b", lower):
            return key
    return None


def status_color(text: str) -> RGBColor | None:
    key = _status_key(text)
    return _STATUS_COLOR.get(key) if key else None


def status_bg(text: str) -> RGBColor | None:
    key = _status_key(text)
    return _STATUS_BG.get(key) if key else None


# ---------------------------------------------------------------- generic helpers

def shape_by_name(slide, name: str):
    for shape in slide.shapes:
        if shape.name == name:
            return shape
    raise KeyError(f"Shape {name!r} not found on slide")


def _set_run_text(paragraph, text: str, color: RGBColor | None = None):
    if not paragraph.runs:
        run = paragraph.add_run()
    else:
        run = paragraph.runs[0]
        for extra in paragraph.runs[1:]:
            extra.text = ""
    run.text = text
    if color is not None:
        run.font.color.rgb = color
    return run


def _enable_shrink_to_fit(text_frame):
    """Secondary safety net: if content still runs long, PowerPoint shrinks
    the font to stay inside the box rather than spilling out of it. Primary
    protection is the geometry-based capping below, since autofit only
    recomputes when a user edits the text box, not on every open."""
    text_frame.word_wrap = True
    try:
        text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    except Exception:
        pass


def set_text(shape, text: str, color: RGBColor | None = None, max_chars: int | None = None):
    if max_chars:
        text = lg.truncate_chars(text, max_chars)
    tf = shape.text_frame
    paragraphs = list(tf.paragraphs)
    _set_run_text(paragraphs[0], text, color)
    for p in paragraphs[1:]:
        for r in p.runs:
            r.text = ""
    _enable_shrink_to_fit(tf)


def set_bullets(slide, shape, lines: list[str]):
    """Grows or shrinks the shape's paragraphs to match `lines`, cloning the
    last paragraph's XML for new lines so bullet formatting survives. Caps
    the number of lines to what actually fits above the nearest shape below
    it on the slide, appending a '+N more' line instead of overflowing into
    that neighbor (footer, next box, image, etc.)."""
    if not lines:
        lines = [""]

    available = lg.available_height_below(slide, shape)
    lines = lg.cap_lines_to_fit(shape, lines, available)

    tf = shape.text_frame
    paragraphs = list(tf.paragraphs)
    if len(lines) > len(paragraphs):
        template_p = paragraphs[-1]._p
        for _ in range(len(lines) - len(paragraphs)):
            new_p = deepcopy(template_p)
            template_p.addnext(new_p)
            template_p = new_p
        paragraphs = list(tf.paragraphs)
    elif len(lines) < len(paragraphs):
        for p in paragraphs[len(lines):]:
            p._p.getparent().remove(p._p)
        paragraphs = list(tf.paragraphs)
    for p, line in zip(paragraphs, lines):
        _set_run_text(p, line)
    _enable_shrink_to_fit(tf)


def set_cell_text(cell, text: str, color: RGBColor | None = None, bg: RGBColor | None = None):
    tf = cell.text_frame
    paragraphs = list(tf.paragraphs)
    for extra in paragraphs[1:]:
        extra._p.getparent().remove(extra._p)
    _set_run_text(tf.paragraphs[0], text, color)
    if bg is not None:
        cell.fill.solid()
        cell.fill.fore_color.rgb = bg


def set_table_rows(slide, table_shape, rows_data: list[list], header_rows: int = 1, color_cols: set[int] | None = None):
    """Rebuilds the table's data rows (below `header_rows`) from `rows_data`,
    cloning the template's first data row's XML for every row so column
    widths/fonts/borders survive. Caps the row count to what fits above the
    nearest shape below the table on the slide (a caption, footer, or the
    next box) — table rows auto-expand to fit their content in PowerPoint,
    so an uncapped row count is exactly what pushes a table down into
    whatever sits below it. Truncated rows collapse into a single
    '+N more' summary row instead of growing the table further."""
    color_cols = color_cols or set()
    table = table_shape.table
    tbl = table._tbl
    existing = list(table.rows)
    if len(existing) <= header_rows:
        return

    template_tr = existing[header_rows]._tr
    header_height = sum(r.height for r in existing[:header_rows])
    row_height = existing[header_rows].height
    n_cols = len(existing[header_rows].cells)

    available = lg.available_height_below(slide, table_shape)
    max_rows = lg.max_rows_for(available, header_height, row_height)

    def _note_row(hidden_count: int) -> list[str]:
        row = [""] * n_cols
        row[0] = f"+{hidden_count} more — see the workbook for the full list"
        return row

    rows_data = lg.cap_rows(rows_data, max_rows, _note_row)

    for row in existing[header_rows:]:
        tbl.remove(row._tr)
    for _ in rows_data:
        tbl.append(deepcopy(template_tr))
    rows = list(table.rows)
    for i, row_values in enumerate(rows_data):
        row = rows[header_rows + i]
        for j, val in enumerate(row_values):
            if j >= len(row.cells):
                continue
            color = status_color(str(val)) if j in color_cols else None
            bg = status_bg(str(val)) if j in color_cols else None
            set_cell_text(row.cells[j], str(val), color, bg)


def clone_shape(slide, shape):
    """Appends a deep copy of `shape` to the slide and returns the new shape
    object — used to add a second timeline marker in the same visual style
    as an existing one (the 'Today' line) rather than hand-building new
    shape XML from scratch, which would risk not matching the template."""
    new_el = deepcopy(shape._element)
    slide.shapes._spTree.append(new_el)
    return slide.shapes[-1]


def interpolate_x(anchors: list[tuple[date, int]], target: date) -> int:
    """Linearly interpolates (or extrapolates past the ends of) an EMU
    x-coordinate for `target` across (date, x) anchor points — how a
    schedule-forecast date gets placed on the gate roadmap's timeline."""
    pairs = sorted(anchors, key=lambda p: p[0])
    dates = [d for d, _ in pairs]
    xs = [x for _, x in pairs]
    if target <= dates[0]:
        i0, i1 = 0, 1
    elif target >= dates[-1]:
        i0, i1 = len(dates) - 2, len(dates) - 1
    else:
        i0, i1 = 0, 1
        for i in range(len(dates) - 1):
            if dates[i] <= target <= dates[i + 1]:
                i0, i1 = i, i + 1
                break
    span = (dates[i1] - dates[i0]).days
    if span <= 0:
        return xs[i0]
    frac = (target - dates[i0]).days / span
    return int(xs[i0] + frac * (xs[i1] - xs[i0]))


def set_shape_fill(shape, color: RGBColor):
    """Recolors a solid-fill accent shape (a KPI tile's icon badge, a gate
    roadmap marker) to a computed status color — the template already uses
    these small colored squares as at-a-glance status indicators; this makes
    them reflect real data instead of the template's static example colors."""
    shape.fill.solid()
    shape.fill.fore_color.rgb = color


def set_chart_data(chart, categories: list[str], series: dict[str, list[float]]):
    data = CategoryChartData()
    data.categories = categories
    for name, values in series.items():
        data.add_series(name, values)
    chart.replace_data(data)


def money_m(value: float) -> str:
    return f"${value / 1_000_000:.1f}M"


# ---------------------------------------------------------------- slide builders

def _build_slide1_title(slide, snapshot: ProjectSnapshot, status: str):
    phase = snapshot.current_phase
    phase_label = nar.short_phase_name(phase.name) if phase else "n/a"
    set_text(
        shape_by_name(slide, "Text 5"),
        f"Reporting period: Month {snapshot.reporting_month_number} — {phase_label}  |  "
        f"Program status: {status}  |  Prepared for: Fulfillment Leadership / SVP Review",
    )
    set_text(
        shape_by_name(slide, "Text 6"),
        f"Prepared by: Program Management Office  |  Generated {date.today():%d-%b-%Y}  |  Confidential — internal use only",
    )


def _build_slide2_exec_summary(slide, snapshot: ProjectSnapshot, manual: ManualInputs, status: str, since: date):
    set_text(shape_by_name(slide, "Text 4"), status, color=WHITE)
    set_shape_fill(shape_by_name(slide, "Shape 3"), status_color(status) or GREEN)
    set_text(shape_by_name(slide, "Text 5"), nar.exec_summary_sentence(snapshot, status))
    set_bullets(slide, shape_by_name(slide, "Text 9"), nar.accomplished_this_period(snapshot, since))

    analytic_flags = [f for f in (analytics.gate_slip_risk(snapshot), analytics.workload_imbalance(snapshot)) if f]
    set_bullets(slide, shape_by_name(slide, "Text 13"), nar.watch_items(snapshot, extra_flags=analytic_flags))

    decision_lines = [f"{d['title']} — {d['recommendation']}" for d in manual.decisions_needed[:4]] or ["No decisions pending leadership input this period."]
    set_bullets(slide, shape_by_name(slide, "Text 17"), decision_lines)


def _build_slide3_kpis(slide, snapshot: ProjectSnapshot, manual: ManualInputs):
    kpis = snapshot.kpis
    bsum = budget_summary(manual.budget)
    closed_gates = sum(1 for g in snapshot.gates if not g.is_open)

    set_text(shape_by_name(slide, "Text 4"), str(snapshot.days_to_go_live))
    set_text(shape_by_name(slide, "Text 6"), f"Target {nar.fmt_date(snapshot.meta.target_go_live)}")

    set_text(shape_by_name(slide, "Text 9"), f"{kpis.pct_complete:.0%}")
    set_text(shape_by_name(slide, "Text 11"), f"{kpis.complete} of {kpis.total_tasks} tasks; {kpis.in_progress} in progress")

    set_text(shape_by_name(slide, "Text 14"), str(kpis.overdue_tasks))
    set_text(shape_by_name(slide, "Text 16"), f"{kpis.at_risk} at risk / blocked")

    set_text(shape_by_name(slide, "Text 19"), str(kpis.open_raid))
    set_text(shape_by_name(slide, "Text 21"), f"{kpis.high_severity_raid} High severity — see Risks")

    set_text(shape_by_name(slide, "Text 24"), f"{bsum['variance_pct']:+.1%}")
    set_text(shape_by_name(slide, "Text 26"), f"{money_m(bsum['forecast'])} FAC vs. {money_m(bsum['approved'])} approved")

    cur_gate = snapshot.current_gate
    set_text(shape_by_name(slide, "Text 29"), f"{closed_gates} / {len(snapshot.gates)}")
    if cur_gate:
        set_text(shape_by_name(slide, "Text 31"), f"Gate {cur_gate.number}: {cur_gate.done_fraction}")

    # Each tile's icon square is a real status light, not decoration —
    # recolored from the template's static example colors to what the data
    # actually says, so a leader can read the slide by color before reading
    # a single number.
    overall = nar.program_status(snapshot, manual)
    tasks_color = status_color(overall) or GREEN
    overdue_color = RED if kpis.overdue_tasks > 0 else GREEN
    gate_overdue = bool(cur_gate and cur_gate.due and cur_gate.due < date.today())
    raid_color = RED if gate_overdue or kpis.high_severity_raid >= 2 else (AMBER if kpis.high_severity_raid >= 1 else GREEN)
    variance_abs = abs(bsum["variance_pct"])
    budget_color = GREEN if variance_abs <= 0.05 else (AMBER if variance_abs <= 0.10 else RED)
    gates_color = RED if gate_overdue else (AMBER if cur_gate and cur_gate.is_open else GREEN)

    set_shape_fill(shape_by_name(slide, "Shape 8"), tasks_color)
    set_shape_fill(shape_by_name(slide, "Shape 13"), overdue_color)
    set_shape_fill(shape_by_name(slide, "Shape 18"), raid_color)
    set_shape_fill(shape_by_name(slide, "Shape 23"), budget_color)
    set_shape_fill(shape_by_name(slide, "Shape 28"), gates_color)


def _build_slide4_progress_by_phase(slide, snapshot: ProjectSnapshot):
    set_text(shape_by_name(slide, "Text 1"), f"Task completion against the {snapshot.kpis.total_tasks}-task master plan — stacked by status")
    for shape in slide.shapes:
        if shape.has_chart:
            categories = [p.name for p in snapshot.phases]
            complete = [p.complete for p in snapshot.phases]
            in_progress = [p.in_progress + p.at_risk for p in snapshot.phases]
            not_started = [max(p.tasks - p.complete - p.in_progress - p.at_risk, 0) for p in snapshot.phases]
            set_chart_data(shape.chart, categories, {
                "Complete": complete, "In Progress": in_progress, "Not Started": not_started,
            })
    forecast = analytics.schedule_forecast(snapshot)
    set_bullets(slide, shape_by_name(slide, "Text 4"), nar.phase_commentary(snapshot, forecast))


def _build_slide5_gate_roadmap(slide, snapshot: ProjectSnapshot, manual: ManualInputs, slide_width: int):
    date_shapes = ["Text 7", "Text 11", "Text 15", "Text 19", "Text 23", "Text 27", "Text 31", "Text 35"]
    for shape_name, gate in zip(date_shapes, snapshot.gates):
        set_text(shape_by_name(slide, shape_name), nar.fmt_date(gate.due))

    set_text(shape_by_name(slide, "Text 37"), f"Today (Month {snapshot.reporting_month_number})")

    milestone_date_shapes = ["Text 42", "Text 45", "Text 48", "Text 51", "Text 54", "Text 57"]
    for shape_name, milestone in zip(milestone_date_shapes, manual.milestones):
        target = milestone.get("target")
        if not target:
            gate_num = milestone.get("fallback_gate")
            gate = next((g for g in snapshot.gates if g.number == gate_num), None)
            target_display = nar.fmt_date(gate.due) if gate else "TBD"
        else:
            target_display = target
        set_text(shape_by_name(slide, shape_name), target_display)

    # Gate badges as a real status trail: green = passed, yellow = the gate
    # in play right now, red = open past its due date, white = not started —
    # extending the template's own yellow/white convention with live data.
    gate_configs = [
        ("Shape 4", "Text 5"), ("Shape 8", "Text 9"), ("Shape 12", "Text 13"), ("Shape 16", "Text 17"),
        ("Shape 20", "Text 21"), ("Shape 24", "Text 25"), ("Shape 28", "Text 29"), ("Shape 32", "Text 33"),
    ]
    cur_gate = snapshot.current_gate
    today = date.today()
    DARK_BLUE = RGBColor(0x00, 0x1E, 0x60)
    for (shape_name, label_name), gate in zip(gate_configs, snapshot.gates):
        if not gate.is_open:
            fill, text_color = GREEN, WHITE
        elif gate.due and gate.due < today:
            fill, text_color = RED, WHITE
        elif cur_gate and gate.number == cur_gate.number:
            fill, text_color = YELLOW, DARK_BLUE
        else:
            fill, text_color = WHITE, BLUE
        set_shape_fill(shape_by_name(slide, shape_name), fill)
        set_text(shape_by_name(slide, label_name), f"G{gate.number}", color=text_color)

    # AI schedule forecast, plotted on the same timeline as "Today" instead
    # of just stated in text — a marker that visibly sits left of Today when
    # tasks are finishing early, or right of it when the program is slipping.
    forecast = analytics.schedule_forecast(snapshot)
    gate_shape_names = [cfg[0] for cfg in gate_configs]
    anchors = [
        (gate.due, shape_by_name(slide, name).left + shape_by_name(slide, name).width // 2)
        for name, gate in zip(gate_shape_names, snapshot.gates)
        if gate.due
    ]
    if len(anchors) >= 2:
        forecast_x = interpolate_x(anchors, forecast.projected_go_live)
        forecast_color = GREEN if forecast.direction == "ahead" else (RED if forecast.direction == "behind" else BLUE)

        today_line = shape_by_name(slide, "Shape 36")
        today_label = shape_by_name(slide, "Text 37")
        margin = Pt(6)

        forecast_line = clone_shape(slide, today_line)
        line_x = max(margin, min(forecast_x - forecast_line.width // 2, slide_width - forecast_line.width - margin))
        forecast_line.left = line_x
        forecast_line.line.color.rgb = forecast_color

        forecast_label = clone_shape(slide, today_label)
        label_x = max(margin, min(forecast_x - forecast_label.width // 2, slide_width - forecast_label.width - margin))
        forecast_label.left = label_x
        sign = "+" if forecast.shift_days > 0 else ""
        set_text(forecast_label, f"AI forecast {sign}{forecast.shift_days}d", color=forecast_color)


def _build_slide6_okr(slide, snapshot: ProjectSnapshot, manual: ManualInputs):
    table_shape = shape_by_name(slide, "Table 0")
    rows = []
    override_notes = []
    for obj in snapshot.objectives:
        override = manual.objective_overrides.get(obj.id)
        status = override["status"] if override else obj.status
        title = obj.text if len(obj.text) <= 48 else obj.text[:45] + "..."
        rows.append([f"{obj.id}  {title}", lg.truncate_chars(obj.criteria, 110), obj.tasks_linked, f"{obj.pct_complete:.0%}", status])
        if override and override.get("note"):
            override_notes.append(f"{obj.id} {override.get('note')}")
    set_table_rows(slide, table_shape, rows, header_rows=1, color_cols={4})

    note = " | ".join(override_notes) if override_notes else "All objectives tracking per the tracker; no PMO overrides applied this period."
    set_text(shape_by_name(slide, "Text 3"), note, max_chars=260)


def _build_slide7_workstream_health(slide, snapshot: ProjectSnapshot, manual: ManualInputs):
    workstreams = sorted(snapshot.workstreams, key=lambda w: w.tasks, reverse=True)
    set_text(shape_by_name(slide, "Text 1"), f"Leads, load and status across the {len(workstreams)} active workstreams")
    rows = []
    health_counts = {"Green": 0, "Amber": 0, "Red": 0}
    amber_lines = []
    for ws in workstreams:
        health = nar.workstream_health(ws, snapshot)
        health_counts[health] += 1
        focus = nar.workstream_focus(ws, snapshot, manual.workstream_focus_overrides.get(ws.name))
        rows.append([ws.name, ws.tasks, f"{ws.complete} / {ws.in_progress}", health, focus])
        if health == "Amber":
            amber_lines.append(f"{ws.name} — {focus}")

    table_shape = shape_by_name(slide, "Table 0")
    set_table_rows(slide, table_shape, rows, header_rows=1, color_cols={3})

    for shape in slide.shapes:
        if shape.has_chart:
            set_chart_data(shape.chart, ["Green", "Amber", "Red"], {
                "Health": [health_counts["Green"], health_counts["Amber"], health_counts["Red"]],
            })

    set_bullets(slide, shape_by_name(slide, "Text 4"), amber_lines or ["No Amber or Red workstreams this period."])


def _build_slide8_top_risks(slide, snapshot: ProjectSnapshot):
    open_count = len(snapshot.open_raid_items)
    set_text(shape_by_name(slide, "Text 1"), f"From the RAID log — {open_count} open item{'s' if open_count != 1 else ''}, ranked by probability × impact")
    risks = nar.top_risks(snapshot)
    rows = [
        [r.id, r.type, lg.truncate_chars(r.description, 100), snapshot.stakeholder_name(r.owner_id),
         f"{r.score} {r.severity}", lg.truncate_chars(r.mitigation, 110), nar.fmt_date(r.due_date)]
        for r in risks
    ]
    table_shape = shape_by_name(slide, "Table 0")
    set_table_rows(slide, table_shape, rows, header_rows=1, color_cols={4})

    for shape in slide.shapes:
        if shape.has_chart:
            dist = nar.severity_distribution(snapshot)
            set_chart_data(shape.chart, ["High", "Medium", "Low"], {
                "Open RAID": [dist["High"], dist["Medium"], dist["Low"]],
            })

    not_shown = [r for r in snapshot.open_raid_items if r.severity.strip().lower() == "low"]
    if not_shown:
        text = "Not shown: " + "; ".join(f"{r.id} {r.description} (Low)" for r in not_shown) + f". All {len(snapshot.open_raid_items)} open items have named owners and due dates."
    else:
        text = f"All {len(snapshot.open_raid_items)} open RAID items shown above have named owners and due dates."
    set_text(shape_by_name(slide, "Text 3"), text, max_chars=260)


def _build_slide9_budget(slide, snapshot: ProjectSnapshot, manual: ManualInputs):
    bsum = budget_summary(manual.budget)
    categories = manual.budget.get("by_category", [])
    set_text(
        shape_by_name(slide, "Text 1"),
        f"Approved {money_m(bsum['approved'])} capex baseline — illustrative figures pending Finance month-end close",
    )

    for shape in slide.shapes:
        if shape.has_chart:
            set_chart_data(
                shape.chart,
                [c["name"] for c in categories],
                {
                    "Approved": [c["approved"] / 1_000_000 for c in categories],
                    "Committed": [c["committed"] / 1_000_000 for c in categories],
                    "Spent to date": [c["spent"] / 1_000_000 for c in categories],
                },
            )

    set_text(shape_by_name(slide, "Text 3"), money_m(bsum["approved"]))
    set_text(shape_by_name(slide, "Text 6"), money_m(bsum["committed"]))
    set_text(shape_by_name(slide, "Text 7"), f"Committed ({bsum['committed_pct']:.0%})")
    set_text(shape_by_name(slide, "Text 9"), money_m(bsum["spent"]))
    set_text(shape_by_name(slide, "Text 10"), f"Spent to date ({bsum['spent_pct']:.0%})")
    set_text(shape_by_name(slide, "Text 12"), money_m(bsum["forecast"]))
    sign = "+" if bsum["variance"] >= 0 else ""
    set_text(shape_by_name(slide, "Text 15"), f"{sign}{money_m(bsum['variance'])}")
    set_text(shape_by_name(slide, "Text 16"), f"Variance ({bsum['variance_pct']:+.1%})")

    drivers = manual.budget.get("variance_drivers", [])
    rows = [[d["driver"], f"{d['amount_m']:+.1f}", d["status"]] for d in drivers]
    table_shape = shape_by_name(slide, "Table 0")
    set_table_rows(slide, table_shape, rows, header_rows=1, color_cols={2})

    top_committed = sorted(categories, key=lambda c: c["committed"], reverse=True)[:3]
    breakdown = ", ".join(f"{c['name']} {money_m(c['committed'])}" for c in top_committed)
    pace = analytics.budget_pace(snapshot, manual)
    pace_line = nar.budget_pace_line(pace, bsum["approved"])
    remaining_line = f" {money_m(pace.remaining)} remaining of {money_m(bsum['approved'])} approved."
    set_text(
        shape_by_name(slide, "Text 18"),
        f"Ask: release {money_m(bsum['committed'])} in vendor/contract commitments this month to hold manufacturing and "
        f"install slots ({breakdown}).{remaining_line}{' ' + pace_line + '.' if pace_line else ''}",
        max_chars=320,
    )


def _build_slide10_decisions(slide, manual: ManualInputs):
    field_shapes = [
        ("Text 5", "Text 6", "Text 7", "Text 8"),
        ("Text 12", "Text 13", "Text 14", "Text 15"),
        ("Text 19", "Text 20", "Text 21", "Text 22"),
        ("Text 26", "Text 27", "Text 28", "Text 29"),
    ]
    decisions = manual.decisions_needed[:4]
    for i, shapes in enumerate(field_shapes):
        title_s, rec_s, rationale_s, owner_s = shapes
        if i < len(decisions):
            d = decisions[i]
            set_text(shape_by_name(slide, title_s), d["title"])
            set_text(shape_by_name(slide, rec_s), f"Recommend: {d['recommendation']}")
            set_text(shape_by_name(slide, rationale_s), d["rationale"])
            set_text(shape_by_name(slide, owner_s), d["owner"])
        else:
            set_text(shape_by_name(slide, title_s), "No decision pending")
            set_text(shape_by_name(slide, rec_s), "")
            set_text(shape_by_name(slide, rationale_s), "")
            set_text(shape_by_name(slide, owner_s), "")


def _build_slide11_path_forward(slide, snapshot: ProjectSnapshot, trend: analytics.TrendDeltas):
    boxes = nar.path_forward_boxes(snapshot)
    title_bullet_shapes = [("Text 4", "Text 5"), ("Text 8", "Text 9"), ("Text 12", "Text 13"), ("Text 16", "Text 17")]
    for (title_s, bullets_s), (title, bullets) in zip(title_bullet_shapes, boxes):
        set_text(shape_by_name(slide, title_s), title)
        set_bullets(slide, shape_by_name(slide, bullets_s), bullets)
    set_text(shape_by_name(slide, "Text 19"), nar.next_report_out_line(snapshot, trend), max_chars=280)


def _build_slide12_appendix(slide, snapshot: ProjectSnapshot):
    gate = snapshot.current_gate
    if gate:
        set_text(shape_by_name(slide, "Text 0"), f"Appendix — Gate {gate.number} checklist and open actions")
        gate_label = gate.name.split("—")[-1].strip() if "—" in gate.name else gate.name
        set_text(shape_by_name(slide, "Text 1"), f"Detail behind the {gate_label} status")
        rows = [[item.text, item.status] for item in gate.items]
        table_shape = shape_by_name(slide, "Table 0")
        set_table_rows(slide, table_shape, rows, header_rows=1, color_cols={1})
        set_text(
            shape_by_name(slide, "Text 2"),
            f"Gate status: {gate.status_text} ({gate.done_fraction}). Expected to pass {nar.fmt_date(gate.due)}.",
        )

    actions = sorted(snapshot.actions, key=lambda a: (a.due_date or date.max))
    action_rows = [[lg.truncate_chars(a.action, 90), snapshot.stakeholder_name(a.owner_id), nar.fmt_date(a.due_date), a.status] for a in actions]
    table1_shape = shape_by_name(slide, "Table 1")
    set_table_rows(slide, table1_shape, action_rows, header_rows=1, color_cols={3})

    set_bullets(slide, shape_by_name(slide, "Text 5"), [
        "Single source: the Greenfield Tracking Playbook workbook (Dashboard, Master Plan, RACI, Gate Checklists, RAID, Action Log)",
        "SC-FC-PMER regenerates this deck on demand or weekly whenever the workbook or manual budget/decision inputs change",
        "Budget figures and leadership decisions are PM-maintained inputs — always verify before sending externally",
    ])


def extract_text_signature(pptx_path: Path) -> str:
    """Concatenates every visible text run and table cell across all slides.
    Used to compare a freshly-built deck against the most recent one on
    disk: two presentations with this same signature are, for an executive
    reader, the same deck — even if e.g. a chart's underlying series order
    differed at the XML level. Kept deliberately simple (text only, no
    formatting/position) so the comparison is about content, not encoding."""
    prs = Presentation(str(pptx_path))
    parts: list[str] = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame and shape.text_frame.text:
                parts.append(shape.text_frame.text)
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        if cell.text:
                            parts.append(cell.text)
    return "\n".join(parts)


def build_presentation(
    snapshot: ProjectSnapshot,
    manual: ManualInputs,
    template_path: Path,
    output_path: Path,
    since: date,
    previous_kpi: dict | None = None,
) -> None:
    prs = Presentation(str(template_path))
    slides = list(prs.slides)
    status = nar.program_status(snapshot, manual)
    trend = analytics.trend_deltas(_current_kpi_dict(snapshot, manual), previous_kpi)

    _build_slide1_title(slides[0], snapshot, status)
    _build_slide2_exec_summary(slides[1], snapshot, manual, status, since)
    _build_slide3_kpis(slides[2], snapshot, manual)
    _build_slide4_progress_by_phase(slides[3], snapshot)
    _build_slide5_gate_roadmap(slides[4], snapshot, manual, prs.slide_width)
    _build_slide6_okr(slides[5], snapshot, manual)
    _build_slide7_workstream_health(slides[6], snapshot, manual)
    _build_slide8_top_risks(slides[7], snapshot)
    _build_slide9_budget(slides[8], snapshot, manual)
    _build_slide10_decisions(slides[9], manual)
    _build_slide11_path_forward(slides[10], snapshot, trend)
    _build_slide12_appendix(slides[11], snapshot)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output_path))


def _current_kpi_dict(snapshot: ProjectSnapshot, manual: ManualInputs) -> dict:
    """Same shape as generate.py's stored kpi_json, computed fresh here so
    trend deltas can be derived without a round-trip through the DB."""
    return {
        "pct_complete": snapshot.kpis.pct_complete,
        "open_raid": snapshot.kpis.open_raid,
        "high_severity_raid": snapshot.kpis.high_severity_raid,
        "days_to_go_live": snapshot.days_to_go_live,
    }
