#!/bin/bash
# Wrapper to run AppleScript and pipe to Python processor
# Make executable: chmod +x run_watchlist.sh

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
	# Use conda run to execute the Python script within the 'futures' env
	# --no-capture-output keeps stdout/stderr behavior normal
	"$OSASCRIPT_BIN" "$APPLE_SCRIPT" | "$CONDA_CMD" run -n futures --no-capture-output python "$PY_SCRIPT" --out "$SCRIPT_DIR/watchlist_summary.xlsx" >> "$LOGFILE" 2>&1
else
	# Fallback to system python
	PYTHON_BIN="/usr/bin/python3"
	"$OSASCRIPT_BIN" "$APPLE_SCRIPT" | "$PYTHON_BIN" "$PY_SCRIPT" --out "$SCRIPT_DIR/watchlist_summary.xlsx" >> "$LOGFILE" 2>&1
fi

# Auto-commit and push to GitHub
REPO_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"
DAY_NAME=$(date +%A)
cd "$REPO_DIR"
git add -A >> "$LOGFILE" 2>&1
git commit -m "updated ${DAY_NAME} data" >> "$LOGFILE" 2>&1
git push origin >> "$LOGFILE" 2>&1
