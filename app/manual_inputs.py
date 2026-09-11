"""Loads the small human-maintained inputs that have no source in the workbook
(budget figures, leadership decisions, objective judgement overrides, milestone
targets). Seeded on first run with the template deck's own example values so
the first generation succeeds unattended; the PMO edits this file thereafter."""
from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path

from app.config import MANUAL_INPUTS_PATH

DEFAULT_MANUAL_INPUTS = {
    "program_status_override": None,
    "budget": {
        "forecast_at_completion": 46_400_000,
        "by_category": [
            {"name": "Building & site fit-out", "approved": 12_000_000, "committed": 4_200_000, "spent": 900_000},
            {"name": "MHE automation", "approved": 18_500_000, "committed": 14_800_000, "spent": 400_000},
            {"name": "WMS / WCS software & services", "approved": 6_000_000, "committed": 4_600_000, "spent": 500_000},
            {"name": "IT infrastructure", "approved": 2_500_000, "committed": 600_000, "spent": 100_000},
            {"name": "Integration & data", "approved": 1_500_000, "committed": 300_000, "spent": 50_000},
            {"name": "Staffing, training & change", "approved": 2_000_000, "committed": 100_000, "spent": 50_000},
            {"name": "PMO & contingency", "approved": 3_000_000, "committed": 400_000, "spent": 100_000},
        ],
        "variance_drivers": [
            {"driver": "MHE — expedite fee to protect manufacturing slot (R-001)", "amount_m": 0.6, "status": "Provisional"},
            {"driver": "IT infrastructure — additional Wi-Fi access points (I-005)", "amount_m": 0.3, "status": "Decision requested"},
        ],
    },
    "workstream_focus_overrides": {},
    "decisions_needed": [
        {
            "title": "Release MHE and WMS vendor contracts ($24.5M)",
            "recommendation": "APPROVE now.",
            "rationale": "Locks manufacturing slot and install window; mitigates the MHE lead-time risk. Each month of delay moves go-live by ~1 month.",
            "owner": "Sponsor / Finance",
        },
        {
            "title": "Approve Wi-Fi design change (+$0.3M from contingency)",
            "recommendation": "APPROVE.",
            "rationale": "Closes the mezzanine Wi-Fi coverage issue before network install; contingency remains healthy after all known variances.",
            "owner": "Sponsor / IT",
        },
        {
            "title": "Start go-live recruiting 4 weeks early",
            "recommendation": "APPROVE.",
            "rationale": "Adds modest opex; removes the largest people risk ahead of training and dress rehearsal.",
            "owner": "Operations / HR",
        },
        {
            "title": "Endorse AI Copilot for status, RAID and gate audits program-wide",
            "recommendation": "APPROVE (pilot already in use by PMO).",
            "rationale": "Cuts reporting effort 40-50%; every gate audited against checklist so nothing is missed. No incremental cost.",
            "owner": "SVP Fulfillment",
        },
    ],
    "objective_overrides": {
        "OBJ-03": {
            "status": "At Risk",
            "note": "flagged At Risk by PMO judgement (tracker shows On Track): the High-severity MHE lead-time risk (R-001) sits directly on the throughput objective.",
        }
    },
    "milestones": [
        {"label": "Vendor contracts executed", "target": None, "fallback_gate": 2},
        {"label": "MHE equipment on site", "target": None, "fallback_gate": 3},
        {"label": "Certificate of occupancy", "target": None, "fallback_gate": 4},
        {"label": "Throughput acceptance ≥ 95%", "target": None, "fallback_gate": 5},
        {"label": "Go / No-Go decision", "target": None, "fallback_gate": 7},
        {"label": "Production go-live", "target": None, "fallback_gate": 7},
    ],
}


@dataclass
class ManualInputs:
    program_status_override: str | None
    budget: dict
    decisions_needed: list[dict]
    objective_overrides: dict
    workstream_focus_overrides: dict
    milestones: list[dict]
    raw: dict = field(default_factory=dict)


def _seed_if_missing(path: Path) -> None:
    if path.exists():
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(DEFAULT_MANUAL_INPUTS, indent=2))


def load_manual_inputs(path: Path | None = None) -> ManualInputs:
    # See app.store._connect for why this is resolved at call time rather
    # than bound as a default parameter value.
    if path is None:
        path = MANUAL_INPUTS_PATH
    _seed_if_missing(path)
    data = json.loads(path.read_text())
    merged = {**DEFAULT_MANUAL_INPUTS, **data}
    return ManualInputs(
        program_status_override=merged.get("program_status_override"),
        budget=merged.get("budget", {}),
        decisions_needed=merged.get("decisions_needed", []),
        objective_overrides=merged.get("objective_overrides", {}),
        workstream_focus_overrides=merged.get("workstream_focus_overrides", {}),
        milestones=merged.get("milestones", []),
        raw=merged,
    )


def raw_bytes(path: Path | None = None) -> bytes:
    if path is None:
        path = MANUAL_INPUTS_PATH
    _seed_if_missing(path)
    return path.read_bytes()


def budget_summary(budget: dict) -> dict:
    """Sums the by-category breakdown into the headline figures shown as KPI
    tiles; forecast-at-completion stays a standalone PM judgement figure."""
    categories = budget.get("by_category", [])
    approved = sum(c.get("approved", 0) for c in categories)
    committed = sum(c.get("committed", 0) for c in categories)
    spent = sum(c.get("spent", 0) for c in categories)
    forecast = budget.get("forecast_at_completion", approved)
    variance = forecast - approved
    return {
        "approved": approved,
        "committed": committed,
        "spent": spent,
        "forecast": forecast,
        "variance": variance,
        "variance_pct": (variance / approved) if approved else 0.0,
        "committed_pct": (committed / approved) if approved else 0.0,
        "spent_pct": (spent / approved) if approved else 0.0,
    }
