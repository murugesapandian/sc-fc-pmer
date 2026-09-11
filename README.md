# SC-FC-PMER — Project Management Executive Report Generator

Turns the live `Greenfield_SCM_GoLive_Tracking_Playbook.xlsx` workbook into the
executive deck (`Greenfield_DC_Monthly_Leadership_ReportOut.pptx` template),
on demand from a webpage or automatically once a week — **only when the
underlying data actually changed**.

## How it works

- **Data source**: every tab of the workbook except "Read Me" (Dashboard,
  Stakeholders, Objectives, Master Plan, Gate Checklists, RAID Log, Action Log).
- **Manual inputs**: a few fields have no source in the workbook — budget
  figures, the "decisions needed" slide, and objective status overrides. These
  live in `data/manual_inputs.json` (auto-created on first run with sensible
  defaults) and are edited by the PMO the same way the workbook is.
- **Change detection**: a hash of every sheet's cell values plus
  `manual_inputs.json` is compared to the last generated report. If nothing
  changed, clicking Generate (or the weekly job) does nothing.
- **Output**: `reports/SC-FC-PMER-YYYY-MM-DD.pptx`, listed with a download
  link in the web UI, oldest to newest, kept indefinitely.
- **Review step**: nothing is sent anywhere automatically — a report only
  ever lands in the `reports/` folder for a human to open, check, and share.
  Budget and decision content especially should be verified before it goes to
  leadership, since those fields are PM-maintained, not workbook-derived.

## Setup

```bash
cd /Users/murugesapandianthangaraj/SC-FC-PMER
python3 -m venv venv
venv/bin/pip install -r requirements.txt
```

## Run the web app

```bash
venv/bin/uvicorn app.main:app --reload --port 8000
```

Open http://localhost:8000 — status card, Generate button, and report
history.

## Generate from the command line (no server needed)

```bash
venv/bin/python scripts/check_and_generate.py
```

Useful for testing the generation engine directly, or for the weekly job below.

## Weekly automatic generation (macOS)

Runs independently of whether the web app is open, via a `launchd`
LaunchAgent that calls the same skip-if-unchanged logic as the button:

```bash
scripts/install_launchd.sh            # uses venv/bin/python3 by default
```

Default schedule: every **Monday 06:00** local time. To change it, edit the
`StartCalendarInterval` block in `scripts/com.scfcpmer.weeklyreport.plist`
and re-run `scripts/install_launchd.sh`. Logs go to `data/launchd.log`.

To test the job immediately: `launchctl start com.scfcpmer.weeklyreport`.
To remove it: `launchctl unload ~/Library/LaunchAgents/com.scfcpmer.weeklyreport.plist && rm ~/Library/LaunchAgents/com.scfcpmer.weeklyreport.plist`.

## Editing the manual inputs

`data/manual_inputs.json` — created automatically on first run. Fields:

- `program_status_override`: `"GREEN"` / `"AMBER"` / `"RED"` / `null` (auto-derived from the workbook when null)
- `budget.by_category`: list of `{name, approved, committed, spent}` — drives the budget chart and KPI tiles
- `budget.forecast_at_completion`: single PM judgement figure
- `budget.variance_drivers`: list of `{driver, amount_m, status}` for the budget slide's table
- `decisions_needed`: up to 4 `{title, recommendation, rationale, owner}` — the template's Decisions slide has 4 fixed card slots
- `objective_overrides`: `{"OBJ-03": {"status": "At Risk", "note": "..."}}` — overrides the tracker's computed status where PM judgement differs
- `workstream_focus_overrides`: `{"Workstream Name": "custom focus line"}` — otherwise auto-derived from the next open task in that workstream
- `milestones`: gate-roadmap milestone labels; `target` (a date string) if fixed, otherwise the date is pulled from `fallback_gate`'s due date

Editing this file counts as a change for the weekly/on-demand generation check,
same as editing the workbook.

## Project layout

```
app/                 generation engine + FastAPI app
web/                 responsive frontend (no build step)
reports/             generated .pptx files
data/                state.db (SQLite history), manual_inputs.json, launchd.log
scripts/             CLI entrypoint + launchd install
```
