"""Content-based change detection.

Hashes only cell *values* (not formatting) across every tracking sheet, plus
the manual_inputs.json bytes, so a cosmetic Excel save doesn't trigger a
regeneration but any real data or manual-input edit does.
"""
from __future__ import annotations

import hashlib
from pathlib import Path

import openpyxl

from app.config import EXCEL_PATH, MANUAL_INPUTS_PATH
from app.manual_inputs import raw_bytes

IGNORED_SHEETS = {"Read Me"}


def _canonical_workbook_bytes(excel_path: Path) -> bytes:
    wb = openpyxl.load_workbook(excel_path, data_only=True, read_only=True)
    try:
        parts = []
        for ws in wb.worksheets:
            if ws.title in IGNORED_SHEETS:
                continue
            parts.append(f"##SHEET:{ws.title}")
            for row in ws.iter_rows():
                for cell in row:
                    if cell.value is not None:
                        parts.append(f"{cell.coordinate}={cell.value!r}")
        return "\n".join(parts).encode("utf-8")
    finally:
        wb.close()


def compute_source_hash(excel_path: Path | None = None, manual_inputs_path: Path | None = None) -> str:
    if excel_path is None:
        excel_path = EXCEL_PATH
    if manual_inputs_path is None:
        manual_inputs_path = MANUAL_INPUTS_PATH
    h = hashlib.sha256()
    h.update(_canonical_workbook_bytes(excel_path))
    h.update(b"\n##MANUAL_INPUTS\n")
    h.update(raw_bytes(manual_inputs_path))
    return h.hexdigest()
