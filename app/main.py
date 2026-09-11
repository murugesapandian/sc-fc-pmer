"""FastAPI app: serves the responsive web UI and the report-generation API."""
from __future__ import annotations

import asyncio
import json
import logging
from contextlib import asynccontextmanager

from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

from app.config import AUTO_CHECK_INTERVAL_SECONDS, EXCEL_PATH, REPORTS_DIR, WEB_DIR
from app.generate import run_generation
from app.hashing import compute_source_hash
from app import store

logger = logging.getLogger("sc-fc-pmer")


async def _auto_check_loop() -> None:
    """Runs for the lifetime of the server: checks the workbook every
    AUTO_CHECK_INTERVAL_SECONDS and only writes a new report when something
    actually changed — run_generation's own hash check does that skip, this
    loop just calls it on a timer. Runs in a worker thread since it's a
    blocking, moderately expensive operation (opens the workbook, and on a
    real change, builds the whole deck) and must not stall the event loop
    that's serving the UI in the meantime."""
    while True:
        try:
            result = await asyncio.to_thread(run_generation, trigger="scheduled")
            if result.generated:
                logger.info("Auto-check generated %s", result.report.filename)
        except Exception:
            logger.exception("Auto-check run failed")
        await asyncio.sleep(AUTO_CHECK_INTERVAL_SECONDS)


@asynccontextmanager
async def lifespan(app: FastAPI):
    task = asyncio.create_task(_auto_check_loop())
    try:
        yield
    finally:
        task.cancel()


app = FastAPI(title="SC-FC-PMER — Project Management Executive Report", lifespan=lifespan)


@app.get("/api/status")
def get_status():
    if not EXCEL_PATH.exists():
        raise HTTPException(status_code=404, detail=f"Workbook not found at {EXCEL_PATH}")

    reports = store.reconcile_reports(REPORTS_DIR)
    current_hash = compute_source_hash(EXCEL_PATH)
    last_hash = store.get_last_source_hash()
    last_report = reports[0] if reports else None

    return {
        "up_to_date": last_hash is not None and last_hash == current_hash,
        "last_report": _report_dict(last_report) if last_report else None,
        "report_count": len(reports),
    }


@app.post("/api/generate")
def post_generate(force: bool = False):
    result = run_generation(trigger="manual", force=force)
    return {
        "generated": result.generated,
        "reason": result.reason,
        "report": _report_dict(result.report) if result.report else None,
    }


@app.get("/api/reports")
def get_reports():
    return [_report_dict(r) for r in store.reconcile_reports(REPORTS_DIR)]


@app.get("/api/reports/{report_id}/download")
def download_report(report_id: int):
    record = store.get_report(report_id)
    if not record:
        raise HTTPException(status_code=404, detail="Report not found")
    path = REPORTS_DIR / record.filename
    if not path.exists():
        raise HTTPException(status_code=404, detail="Report file missing on disk")
    return FileResponse(
        path,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename=record.filename,
    )


def _report_dict(record: store.ReportRecord) -> dict:
    kpi = None
    if record.kpi_json:
        try:
            kpi = json.loads(record.kpi_json)
        except ValueError:
            kpi = None
    return {
        "id": record.id,
        "filename": record.filename,
        "generated_at": record.generated_at,
        "trigger": record.trigger,
        "source_hash": record.source_hash[:12],
        "kpi": kpi,
    }


app.mount("/", StaticFiles(directory=str(WEB_DIR), html=True), name="web")
