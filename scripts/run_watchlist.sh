#!/bin/bash
# Wrapper to run AppleScript and pipe to Python processor
# Make executable: chmod +x run_watchlist.sh

set -u

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
APPLE_SCRIPT="$SCRIPT_DIR/extract_mail.scpt"
PY_SCRIPT="$SCRIPT_DIR/process_watchlist.py"
LOG_DIR="$SCRIPT_DIR/logs"
mkdir -p "$LOG_DIR"
LOGFILE="$LOG_DIR/watchlist_$(date +%Y%m%d).log"

# Use full path to osascript and prefer running the Python inside the 'futures' conda env
OSASCRIPT_BIN="/usr/bin/osascript"

# Try to locate conda
if command -v conda >/dev/null 2>&1; then
	CONDA_CMD="conda"
elif [ -x "$HOME/miniconda3/bin/conda" ]; then
	CONDA_CMD="$HOME/miniconda3/bin/conda"
elif [ -x "$HOME/opt/miniconda3/bin/conda" ]; then
	CONDA_CMD="$HOME/opt/miniconda3/bin/conda"
else
	CONDA_CMD=""
fi

if [ -n "$CONDA_CMD" ]; then
	PY_RUNNER=("$CONDA_CMD" run -n futures --no-capture-output python "$PY_SCRIPT")
else
	# Fallback to system python
	PYTHON_BIN="/usr/bin/python3"
	PY_RUNNER=("$PYTHON_BIN" "$PY_SCRIPT")
fi

OUT_DIR="$(cd "$SCRIPT_DIR/.." && pwd)/outputs"
mkdir -p "$OUT_DIR"

LAST_RUN_HAS_NEW_DATA=0
LAST_RUN_STATUS=0

run_watchlist_once() {
	local run_label="$1"
	local run_tmp
	run_tmp=$(mktemp)

	echo "[$(date '+%Y-%m-%d %H:%M:%S %Z')] Starting ${run_label} run" >> "$LOGFILE"
	"$OSASCRIPT_BIN" "$APPLE_SCRIPT" | "${PY_RUNNER[@]}" --out "$OUT_DIR/watchlist_summary.xlsx" > "$run_tmp" 2>&1
	LAST_RUN_STATUS=$?

	cat "$run_tmp" >> "$LOGFILE"

	if [ "$LAST_RUN_STATUS" -eq 0 ] && grep -Eq 'added [1-9][0-9]* new date\(s\)' "$run_tmp"; then
		LAST_RUN_HAS_NEW_DATA=1
		echo "[$(date '+%Y-%m-%d %H:%M:%S %Z')] New data detected (${run_label})" >> "$LOGFILE"
	else
		LAST_RUN_HAS_NEW_DATA=0
		echo "[$(date '+%Y-%m-%d %H:%M:%S %Z')] No new data detected (${run_label})" >> "$LOGFILE"
	fi

	rm -f "$run_tmp"
}

run_watchlist_once "initial"

if [ "$LAST_RUN_STATUS" -eq 0 ] && [ "$LAST_RUN_HAS_NEW_DATA" -eq 0 ]; then
	echo "[$(date '+%Y-%m-%d %H:%M:%S %Z')] Waiting 30 minutes before retry" >> "$LOGFILE"
	sleep 1800
	run_watchlist_once "retry-30m"
fi

if [ "$LAST_RUN_STATUS" -eq 0 ] && [ "$LAST_RUN_HAS_NEW_DATA" -eq 0 ]; then
	current_h=$((10#$(date +%H)))
	current_m=$((10#$(date +%M)))
	current_s=$((10#$(date +%S)))
	current_total=$((current_h * 3600 + current_m * 60 + current_s))
	target_total=$((8 * 3600))

	if [ "$current_total" -lt "$target_total" ]; then
		sleep_seconds=$((target_total - current_total))
		echo "[$(date '+%Y-%m-%d %H:%M:%S %Z')] No new data after 30m retry; waiting until 08:00 (${sleep_seconds}s)" >> "$LOGFILE"
		sleep "$sleep_seconds"
	else
		echo "[$(date '+%Y-%m-%d %H:%M:%S %Z')] No new data after 30m retry; 08:00 already passed, running fallback now" >> "$LOGFILE"
	fi

	run_watchlist_once "fallback-8am"
fi

# Auto-commit and push to GitHub
REPO_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"
DAY_NAME=$(date +%A)
cd "$REPO_DIR"
git add -A >> "$LOGFILE" 2>&1

if git diff --cached --quiet; then
	echo "[$(date '+%Y-%m-%d %H:%M:%S %Z')] No git changes to commit" >> "$LOGFILE"
else
	git commit -m "updated ${DAY_NAME} data" >> "$LOGFILE" 2>&1
	git push origin >> "$LOGFILE" 2>&1
fi
