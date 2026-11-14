#!/usr/bin/env bash
# Runs the follow-up agent every weekday at 9:00 on this Mac, with a notification,
# and opens the drafts page when someone is due.
#   bash scripts/schedule-mac.sh          turn the schedule on
#   bash scripts/schedule-mac.sh --off    turn it off
set -euo pipefail
REPO="$(cd "$(dirname "$0")/.." && pwd)"
LABEL="com.kinmo.followup-agent"
PLIST="$HOME/Library/LaunchAgents/$LABEL.plist"
PY="$(command -v python3)"

if [ "${1:-}" = "--off" ]; then
  launchctl bootout "gui/$(id -u)/$LABEL" 2>/dev/null || true
  rm -f "$PLIST"
  echo "Schedule turned off."
  exit 0
fi

mkdir -p "$HOME/Library/LaunchAgents" "$REPO/logs"
DAYS=""
for d in 1 2 3 4 5; do
  DAYS="$DAYS<dict><key>Weekday</key><integer>$d</integer><key>Hour</key><integer>9</integer><key>Minute</key><integer>0</integer></dict>"
done
cat > "$PLIST" <<PLIST
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0"><dict>
  <key>Label</key><string>$LABEL</string>
  <key>ProgramArguments</key><array><string>$PY</string><string>$REPO/followup.py</string><string>--open</string><string>--notify</string></array>
  <key>WorkingDirectory</key><string>$REPO</string>
  <key>StartCalendarInterval</key><array>$DAYS</array>
  <key>StandardOutPath</key><string>$REPO/logs/followup.log</string>
  <key>StandardErrorPath</key><string>$REPO/logs/followup.log</string>
</dict></plist>
PLIST
launchctl bootout "gui/$(id -u)/$LABEL" 2>/dev/null || true
launchctl bootstrap "gui/$(id -u)" "$PLIST"
echo "Scheduled: every weekday at 9:00. Log: $REPO/logs/followup.log"
