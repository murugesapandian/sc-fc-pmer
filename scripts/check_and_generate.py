#!/usr/bin/env python3
"""Weekly launchd entrypoint (also runnable by hand for testing).

Regenerates the executive report only if the workbook or manual inputs
changed since the last report — same skip-if-unchanged rule the web UI's
Generate button uses, just with trigger='scheduled' for the history log.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from app.generate import run_generation  # noqa: E402


def main() -> None:
    result = run_generation(trigger="scheduled")
    if result.generated:
        print(f"[SC-FC-PMER] Generated {result.report.filename}")
    else:
        print(f"[SC-FC-PMER] Skipped: {result.reason}")


if __name__ == "__main__":
    main()
