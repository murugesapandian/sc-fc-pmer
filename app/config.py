"""Paths and constants shared across the app."""
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parent.parent

EXCEL_PATH = ROOT_DIR / "Greenfield_SCM_GoLive_Tracking_Playbook.xlsx"
TEMPLATE_PPTX_PATH = ROOT_DIR / "Greenfield_DC_Monthly_Leadership_ReportOut.pptx"

REPORTS_DIR = ROOT_DIR / "reports"
DATA_DIR = ROOT_DIR / "data"
WEB_DIR = ROOT_DIR / "web"

DB_PATH = DATA_DIR / "state.db"
MANUAL_INPUTS_PATH = DATA_DIR / "manual_inputs.json"
LAUNCHD_LOG_PATH = DATA_DIR / "launchd.log"

REPORT_NAME_PREFIX = "SC-FC-PMER"

# How often the running web app checks the workbook for changes on its own,
# independent of anyone clicking Generate. Separate from the weekly launchd
# job (which covers the case where the app isn't running at all) — this is
# the fast, in-app path for while the server is up.
AUTO_CHECK_INTERVAL_SECONDS = 120

REPORTS_DIR.mkdir(parents=True, exist_ok=True)
DATA_DIR.mkdir(parents=True, exist_ok=True)
