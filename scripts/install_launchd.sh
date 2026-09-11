#!/usr/bin/env bash
# Installs the weekly report-check as a macOS LaunchAgent, independent of
# whether the web app is running. Re-run after editing the .plist's schedule.
set -euo pipefail

PROJECT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
PYTHON_BIN="${1:-$PROJECT_DIR/venv/bin/python3}"

if [ ! -x "$PYTHON_BIN" ]; then
  echo "Python interpreter not found or not executable at: $PYTHON_BIN"
  echo "Usage: $0 [path-to-python-in-venv]"
  echo "(Create the venv first: python3 -m venv venv && venv/bin/pip install -r requirements.txt)"
  exit 1
fi

PLIST_LABEL="com.scfcpmer.weeklyreport"
TARGET="$HOME/Library/LaunchAgents/${PLIST_LABEL}.plist"

mkdir -p "$HOME/Library/LaunchAgents"
mkdir -p "$PROJECT_DIR/data"

sed \
  -e "s|__PYTHON_BIN__|$PYTHON_BIN|g" \
  -e "s|__PROJECT_DIR__|$PROJECT_DIR|g" \
  "$PROJECT_DIR/scripts/com.scfcpmer.weeklyreport.plist" > "$TARGET"

launchctl unload "$TARGET" 2>/dev/null || true
launchctl load "$TARGET"

echo "Installed and loaded: $TARGET"
echo "Runs every Monday at 06:00 local time, whether or not the web app is open."
echo "To change the schedule: edit the StartCalendarInterval block in"
echo "  $PROJECT_DIR/scripts/com.scfcpmer.weeklyreport.plist"
echo "then re-run this script."
echo
echo "To test immediately:   launchctl start $PLIST_LABEL"
echo "To uninstall:          launchctl unload $TARGET && rm $TARGET"
