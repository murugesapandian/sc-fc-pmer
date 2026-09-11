"""Orchestrates one report generation run. Shared by the FastAPI endpoint and
the standalone CLI used by the weekly launchd job, so both paths behave
identically (same skip-if-unchanged rule, same filename convention)."""
from __future__ import annotations

import json
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path

from app.config import EXCEL_PATH, REPORT_NAME_PREFIX, REPORTS_DIR, TEMPLATE_PPTX_PATH
from app.excel_reader import ProjectSnapshot, load_snapshot
from app.hashing import compute_source_hash
from app.manual_inputs import ManualInputs, load_manual_inputs
from app.pptx_builder import build_presentation, extract_text_signature
from app import narrative as nar
from app.manual_inputs import budget_summary
from app import store


def _kpi_snapshot(snapshot: ProjectSnapshot, manual: ManualInputs) -> str:
    """A small headline-numbers snapshot captured at generation time, so the
    web UI can show a live 'program at a glance' strip without re-opening
    the generated .pptx just to read a few numbers back out of it."""
    status = nar.program_status(snapshot, manual)
    phase = snapshot.current_phase
    bsum = budget_summary(manual.budget)
    return json.dumps({
        "status": status,
        "phase": nar.short_phase_name(phase.name) if phase else None,
        "pct_complete": round(snapshot.kpis.pct_complete, 4),
        "days_to_go_live": snapshot.days_to_go_live,
        "open_raid": snapshot.kpis.open_raid,
        "high_severity_raid": snapshot.kpis.high_severity_raid,
        "overdue_tasks": snapshot.kpis.overdue_tasks,
        "budget_variance_pct": round(bsum["variance_pct"], 4),
        "total_tasks": snapshot.kpis.total_tasks,
        "complete_tasks": snapshot.kpis.complete,
    })


@dataclass
class GenerateResult:
    generated: bool
    reason: str
    report: store.ReportRecord | None = None


def _next_available_path(base_name: str) -> Path:
    """Unique against the DB's full filename history, not just current disk
    state — otherwise deleting a report and regenerating the same day would
    reuse its name, leaving an old history row silently aliased to new
    content instead of correctly showing as gone."""
    used = store.all_filenames()

    def _is_free(name: str) -> bool:
        return name not in used and not (REPORTS_DIR / name).exists()

    candidate_name = f"{base_name}.pptx"
    if _is_free(candidate_name):
        return REPORTS_DIR / candidate_name
    n = 2
    while not _is_free(f"{base_name}-v{n}.pptx"):
        n += 1
    return REPORTS_DIR / f"{base_name}-v{n}.pptx"


def _since_date(reports: list[store.ReportRecord]) -> date:
    if reports:
        try:
            return datetime.fromisoformat(reports[0].generated_at).date()
        except ValueError:
            pass
    return date.today() - timedelta(days=30)


def run_generation(trigger: str, force: bool = False) -> GenerateResult:
    """trigger: 'manual' | 'scheduled' | 'forced'."""
    store.set_last_checked(datetime.now().isoformat(timespec="seconds"))

    # Self-heal against reports deleted outside the app: a stale DB entry for
    # a missing file would otherwise both clutter history and make
    # last_source_hash lie about there being nothing left to regenerate.
    reports = store.reconcile_reports(REPORTS_DIR)

    current_hash = compute_source_hash(EXCEL_PATH)
    last_hash = store.get_last_source_hash()

    if not force and last_hash == current_hash:
        return GenerateResult(generated=False, reason="No changes to the workbook or manual inputs since the last report.")

    snapshot = load_snapshot(EXCEL_PATH)
    manual = load_manual_inputs()
    since = _since_date(reports)

    previous_kpi = None
    if reports and reports[0].kpi_json:
        try:
            previous_kpi = json.loads(reports[0].kpi_json)
        except ValueError:
            previous_kpi = None

    base_name = f"{REPORT_NAME_PREFIX}-{date.today():%Y-%m-%d}"
    output_path = _next_available_path(base_name)

    # "Regenerate anyway" bypasses the source-hash check above, but it should
    # still refuse to pile up a near-duplicate file when the workbook hasn't
    # actually moved the needle — e.g. someone clicking it out of habit, or a
    # manual-input edit that turned out not to change any visible content.
    # Build to a scratch file first and compare its actual rendered content
    # against the most recent report before committing to a new one.
    if force and reports:
        prev_path = REPORTS_DIR / reports[0].filename
        scratch_path = REPORTS_DIR / f".scratch-{output_path.name}"
        try:
            build_presentation(snapshot, manual, TEMPLATE_PPTX_PATH, scratch_path, since=since, previous_kpi=previous_kpi)
            if prev_path.exists() and extract_text_signature(scratch_path) == extract_text_signature(prev_path):
                return GenerateResult(
                    generated=False,
                    reason=f"The workbook still evaluates to the same content as {reports[0].filename} — no new report created.",
                )
            scratch_path.rename(output_path)
        finally:
            scratch_path.unlink(missing_ok=True)
    else:
        build_presentation(snapshot, manual, TEMPLATE_PPTX_PATH, output_path, since=since, previous_kpi=previous_kpi)

    effective_trigger = "forced" if force else trigger
    record = store.record_report(
        filename=output_path.name,
        trigger=effective_trigger,
        source_hash=current_hash,
        kpi_json=_kpi_snapshot(snapshot, manual),
    )
    return GenerateResult(generated=True, reason="Report generated.", report=record)
