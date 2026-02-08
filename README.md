# Watchlist Summary extractor

Files created:

- `scripts/extract_mail.scpt` — AppleScript that finds messages in Mail with subject "Watchlist Summary (futures)" from the last 24 hours and prints bodies separated by `---MSG---`.
- `scripts/process_watchlist.py` — Python script that reads the concatenated messages from stdin, attempts to parse tabular data, and writes each message into a timestamped Excel sheet.
- `scripts/run_watchlist.sh` — Shell wrapper to run the AppleScript and pipe output to the Python script; logs to `scripts/logs/`.
- `requirements.txt` — Python dependencies.

Setup

1. Install Python deps:

```bash
python3 -m pip install -r requirements.txt
```

Conda (recommended)

1. Create the `futures` environment from `environment.yml`:

```bash
conda env create -f environment.yml
```

2. Activate the environment:

```bash
conda activate futures
```

3. If you prefer to use `pip` inside the conda env instead of the `environment.yml` pip section:

```bash
conda create -n futures python=3.11 pandas openpyxl -c conda-forge
conda activate futures
python3 -m pip install -r requirements.txt
```

Make the wrapper executable:

```bash
chmod +x scripts/run_watchlist.sh
```

Cron (weekdays at 2:00 AM)

Add this line to your crontab (`crontab -e`) to run weekdays (Mon-Fri) at 2:00 AM:

```
0 2 * * 1-5 /Users/bett/practice/daily-futures/scripts/run_watchlist.sh
```

Notes

- The AppleScript searches the `Inbox` for messages with the exact subject and received within the last 24 hours. If you need a different mailbox or time window, edit `scripts/extract_mail.scpt`.
- The Python parser tries common delimiters (comma, tab, multiple spaces). If your watchlist has a specific format, tell me and I can adapt parsing to produce nicely structured Excel sheets.
- Cron jobs run with a minimal environment — modify `run_watchlist.sh` to point to a different `python3` if you use a virtualenv.
