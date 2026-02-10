# Daily Futures Watchlist Scraper

Automated system to extract daily futures watchlist emails from Mail.app, parse price data, and organize into Excel files by month/week. Runs nightly via cron and auto-commits to GitHub.

## Project Structure

```
daily-futures/
├── scripts/
│   ├── extract_mail.scpt      # AppleScript to extract emails from Mail.app
│   ├── process_watchlist.py   # Python parser & Excel writer
│   ├── run_watchlist.sh       # Shell wrapper with auto-commit/push
│   └── logs/                  # Daily execution logs
├── outputs/                   # Generated Excel files (gitignored)
│   ├── november.xlsx
│   ├── december.xlsx
│   ├── january.xlsx
│   └── february.xlsx
├── environment.yml
├── requirements.txt
└── README.md
```

## Features

- **Extracts emails** with subject "Watchlist Summary (futures)" from all Mail.app mailboxes
- **Parses futures data** (Symbol, Latest, Change, %Change, Open, High, Low, Volume, Time)
- **Organizes by month & week** - Creates one Excel file per month with sheets for Week1, Week2, etc.
- **Incremental updates** - Merges new data without overwriting existing dates
- **Auto-commits to GitHub** - Commits and pushes after each run with message "updated [Day] data"
- **Runs unattended** - Scheduled cron job with auto-wake at 2am EAT

## Setup

### 1. Install Dependencies

Create the conda environment:

```bash
conda env create -f environment.yml
conda activate futures
```

Or install via pip:

```bash
python3 -m pip install pandas openpyxl
```

### 2. Configure Git Remote

Ensure your GitHub remote is configured with authentication:

```bash
git remote -v
```

If using HTTPS, set up a personal access token in the remote URL or configure credential storage.

### 3. Grant macOS Permissions

The script needs permission for `osascript` to control Mail.app:

1. Run a test extraction:
   ```bash
   osascript scripts/extract_mail.scpt
   ```

2. If prompted, approve in **System Settings > Privacy & Security > Automation**

3. Ensure Terminal/cron has **Full Disk Access** (for file operations)

### 4. Set Up Cron Job

The cron runs Tuesday-Saturday at 2am EAT (to process Monday-Friday recaps):

```bash
# Install cron schedule
(echo "# Daily futures watchlist - runs at 2am EAT, Tue-Sat (for Mon-Fri recaps)"; \
 echo "0 2 * * 2,3,4,5,6 /bin/bash /Users/bett/practice/daily-futures/scripts/run_watchlist.sh") | crontab -

# Verify installation
crontab -l
```

### 5. Configure Auto-Wake (Optional)

To ensure your Mac wakes at 2am for cron execution:

```bash
sudo pmset repeat wakeorpoweron MTWRFS 01:55:00
```

This schedules wake on Mon-Sat at 1:55am EAT. Check with:

```bash
pmset -g sched
```

**Note:** Your Mac must be logged in (screen can be locked) for cron jobs to run.

## Manual Execution

Run the pipeline manually:

```bash
bash scripts/run_watchlist.sh
```

Output goes to `outputs/*.xlsx` and logs to `scripts/logs/watchlist_YYYYMMDD.log`.

## How It Works

1. **Extract** - AppleScript searches Mail.app for "Watchlist Summary (futures)" emails
2. **Parse** - Python extracts price data and date from each email body
3. **Organize** - Groups data by month (e.g., February) and calendar week (Week1: days 1-7, Week2: 8-14, etc.)
4. **Merge** - Appends new dates to existing Excel files without duplicates
5. **Commit** - Runs `git add -A && git commit -m "updated Tuesday data" && git push origin`

## Output Format

Each monthly Excel file contains:
- **One sheet per week** (Week1, Week2, Week3, Week4, Week5)
- **Horizontal date blocks** - Each date's data in adjacent columns
- **Formatted numbers** - Percentages, decimals, and volume properly formatted
- **Incremental updates** - New dates append without overwriting existing data

Example: `outputs/february.xlsx`
- Week1 sheet: Feb 2, 3, 4, 5, 6
- Week2 sheet: Feb 9, 10, 11...

## Troubleshooting

### Cron doesn't run
- Check cron is active: `crontab -l`
- Verify Mac woke up: `pmset -g log | grep -i wake`
- Check logs: `cat scripts/logs/watchlist_$(date +%Y%m%d).log`

### Mail.app permission denied
- Re-grant access in **System Settings > Privacy & Security > Automation**
- Test manually: `osascript scripts/extract_mail.scpt`

### Git push fails
- Check credentials: `git config --list | grep credential`
- Verify remote: `git remote -v`
- Test manually: `git push origin`

### Wrong timezone
- System should be EAT (Africa/Nairobi)
- Check: `date +%Z` should show "EAT"
