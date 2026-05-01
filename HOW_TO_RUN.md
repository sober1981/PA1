# PA1 - How to Run the Reports

## What PA1 produces

Every PA1 run generates **two attachments**:
- **PDF**: full performance report (Categories 1, 2, 3, and 4 on Friday).
  Category 1A includes three KPI tables: **Summary**, **Detailed**, and
  **Longest Run** — with motor-type fills, a green-yellow-red gradient
  on the Summary numeric columns, red/green diff fonts, and a per-row
  **OP w/ More Runs** column showing the operator with the most runs.
- **Excel**: standalone Weekly KPI Summary with the same three tables
  for QC and side-by-side comparison.

Both files are emailed to `jsoberanes@scoutdownhole.com` and saved to
the project folder.

## Prerequisites

- The master Excel file (`MASTER_MCS_MERGE_*.xlsx`) must already be in
  SharePoint / Teams sync or accessible locally.
- Outlook must be open on your machine (so the email can be sent).

---

## The three `.bat` files at a glance

PA1 has three batch wrappers. Pick the one that matches your situation:

| `.bat` file | When to use it | How it's launched | Prompts? | Email? | Logs to file? |
|---|---|---|---|---|---|
| **`run_manual.bat`** | Manual Wednesday run (you double-click it) | Double-click in Explorer | Yes — file + week | Yes | No (visible in console) |
| **`run_wednesday.bat`** | Wednesday auto safety net at 14:00 if you forgot | Windows Task Scheduler | No (auto-detect) | Yes | Yes — `logs\wednesday_YYYYMMDD.log` |
| **`run_friday.bat`** | Official Friday report at 10:00 | Windows Task Scheduler | No (uses Wed-saved state) | Yes | Yes — `logs\friday_YYYYMMDD.log` |

All three call the same entry point — `run_agent.py` — but with different flags and run modes.

### Detailed differences

| Behavior | `run_manual.bat` | `run_wednesday.bat` | `run_friday.bat` |
|---|---|---|---|
| Python flag passed | `--report wednesday` | `--report wednesday` | `--report friday` |
| TTY (terminal) | Yes (interactive) | No (scheduled, headless) | No (scheduled, headless) |
| Window stays open | Yes (`pause` at end) | N/A — runs hidden | N/A — runs hidden |
| Output redirected to log | No (you see it on screen) | Yes — `logs\wednesday_*.log` | Yes — `logs\friday_*.log` |
| Failure-email fallback | No (you'd see the error) | Yes — emails the last 2,000 chars of the log | Yes — emails the last 2,000 chars of the log |
| File source | Interactive picker (lists local masters) | Auto-detect latest local master | Reads `state\last_run.json` and finds the same Wednesday filename in Teams sync |
| Week | Interactive picker (10 most recent) | Auto-detect latest week in data | Auto-detect latest week in data |
| "Skip if already run today" guard | No — you can always re-run | Yes — if Wednesday state is already saved for today, the scheduled task exits silently | No |
| Pre-QC snapshot saved | Yes — `state\wednesday_snapshot.xlsx` | Yes — `state\wednesday_snapshot.xlsx` | No (uses the existing snapshot) |
| State file `state\last_run.json` written | Yes — used by Friday | Yes — used by Friday | No (only reads it) |
| Categories rendered | 1, 2, 3 | 1, 2, 3 | 1, 2, 3, **4 (QC Audit)** |
| POOH column used | `REASON_POOH` (raw) | `REASON_POOH` (raw) | `REASON_POOH_QC` (post-QC) |

### Why three files instead of one?

- **`run_manual.bat`** is *interactive*: it does NOT redirect output to a
  log file (so you can see prompts and answer them) and it pauses at the
  end so you can read the result before the window closes.
- **`run_wednesday.bat`** and **`run_friday.bat`** are *unattended*: they
  redirect everything to a log file, run hidden (no console), and email
  a failure notification if anything goes wrong. They have no prompts to
  answer because there's no user there.

A single bat file can't cleanly do both modes — interactive prompts can't
be answered when stdin is redirected to a log, and a paused window
doesn't make sense for a scheduled task that has to run at 2 AM.

---

## Manual Wednesday Report — the easy way

**Double-click `run_manual.bat`** in File Explorer:
```
C:\Users\jsoberanes\Projects\scorecard-pa\run_manual.bat
```

A console window opens. You'll be prompted twice:

### Prompt 1 — pick the master file
```
  Found N master file(s):

    [1] MASTER_MCS_MERGE_20260429_085609_Apr 29.xlsx
        Path: C:\Users\jsoberanes\OneDrive ...
        Modified: 2026-04-29 08:56
    [2] ...
  Enter number [1-N] to select, or paste a full file path:
  >
```
Type the number (usually `1`) and press Enter.

### Prompt 2 — pick the week
```
  Available weeks (most recent first):
    [1] 26-W17  (2026-04-20 to 2026-04-26)  - 57 runs
    [2] 26-W16  (2026-04-13 to 2026-04-19)  - 51 runs
    [3] 26-W15  (2026-04-06 to 2026-04-12)  - 62 runs
    ...
    [Enter] = latest (option 1)
    Or type a week ID (e.g. 26-W12)
  >
```
Press Enter for the latest, or type a number / week ID.

PA1 then processes all 8 steps and emails the PDF + KPI Excel.

---

## Manual Wednesday Report — from a terminal

If you prefer the command line (skip the .bat wrapper):

```
cd C:\Users\jsoberanes\Projects\scorecard-pa
"C:\Users\jsoberanes\AppData\Local\Programs\Python\Python313\python.exe" run_agent.py --report wednesday
```

You get the same two prompts (file + week).

> Note: plain `python` won't work — it's not on PATH. Use the full path above.

---

## Skip prompts (specify everything via flags)

```
"C:\...\python.exe" run_agent.py --report wednesday --week 26-W17 --file "path\to\master.xlsx"
```
Or just one flag — whichever is given skips its prompt; the others still ask.

Skip the email too:
```
"C:\...\python.exe" run_agent.py --report wednesday --no-email
```

---

## Automatic Wednesday Safety Net

If you forget to run the Wednesday report, it auto-runs at **2:00 PM on
Wednesdays** via Windows Task Scheduler (`run_wednesday.bat`). It:

- Skips if you already ran it manually earlier today (state-checked).
- Auto-picks the most recent master file (no prompts).
- Auto-uses the latest week.
- Sends PDF + KPI Excel as usual.

The "skip if already run today" guard only applies to the **scheduled
non-interactive run**. You can always re-run manually any time.

---

## Friday Report (automatic)

The Friday report runs automatically every **Friday at 10:00 AM** via
Task Scheduler (`run_friday.bat`). No action needed. It:

- Reads `state\last_run.json` to find the same Wednesday filename.
- Finds the **QC'd version** of that file in the Teams sync folder.
- Compares it against the Wednesday snapshot (Category 4: QC Audit).
- Uses `REASON_POOH_QC` (the cleaned column) instead of `REASON_POOH`.
- Emails PDF + KPI Excel with all 4 categories.

To run Friday manually if needed:
```
"C:\...\python.exe" run_agent.py --report friday
```

---

## Standalone KPI Excel (ad-hoc week analysis)

If you only want the **Weekly KPI Summary Excel** (no PDF, no email),
use the standalone tool:

```
cd C:\Users\jsoberanes\Projects\scorecard-pa
"C:\Users\jsoberanes\AppData\Local\Programs\Python\Python313\python.exe" run_kpi_excel.py
```

Prompts:
1. **Source picker** — `[1] Latest LOCAL  [2] SharePoint  [3] Browse local files`
2. **File confirmation** — proceeds with the detected file (Y/n)
3. **Week picker** — same list as in `run_agent.py`

Output: `PA1 - Weekly KPI Summary - Week 26-WNN.xlsx` in the project folder.

Skip prompts with flags:
```
run_kpi_excel.py --source local --week 26-W17 --yes
run_kpi_excel.py --source sharepoint --week 26-W16 --yes
run_kpi_excel.py --file "C:\path\to\master.xlsx" --week 26-W17 --yes
```

---

## What gets saved where

| File | Location |
|---|---|
| PDF report | `PA1 - {Wed/Fri} Report {dates} - Week {N} - {master}.xlsx.pdf` |
| KPI Excel (from PA1) | `PA1 - {Wed/Fri} Weekly KPI - Week {N} - {master}.xlsx` |
| KPI Excel (standalone) | `PA1 - Weekly KPI Summary - Week 26-W{NN}.xlsx` |
| Wednesday snapshot | `state\wednesday_snapshot.xlsx` (for Friday QC audit) |
| Run state | `state\last_run.json` (so Friday knows which file Wednesday used) |
| Logs (scheduled runs) | `logs\wednesday_YYYYMMDD.log`, `logs\friday_YYYYMMDD.log` |

---

## Quick Reference

| Action | Command |
|---|---|
| Manual Wed (easy)              | Double-click `run_manual.bat` |
| Manual Wed (terminal)          | `run_agent.py --report wednesday` |
| Manual Wed for a specific week | `run_agent.py --report wednesday --week 26-W17` |
| Manual Wed without email       | `run_agent.py --report wednesday --no-email` |
| Manual Fri                     | `run_agent.py --report friday` |
| Standalone KPI Excel only      | `run_kpi_excel.py` (or `--source local --week ... --yes`) |

> Prepend each Python command with `"C:\Users\jsoberanes\AppData\Local\Programs\Python\Python313\python.exe"` and `cd` first to the project folder.

---

## If a manual run fails

The console window stays open (`pause` at end). Read the error message:

- **`Permission denied: ...xlsx`** — the output file is open in Excel. Close it and re-run.
- **`SharePoint connection failed (MFA)`** — pick option `[1] Latest LOCAL` instead, or use the OneDrive-synced file.
- **`Cannot find ... in Teams sync` (Friday)** — the QC'd version isn't synced yet. Wait for the Teams sync to complete, or run with `--file` pointing to the right path.
- **Anything else** — paste the last 10–20 lines of the console into a chat with Claude and I'll debug.

---

## If a scheduled run fails

You'll receive a `PA1 - REPORT FAILED (timestamp)` email automatically,
containing the last 2,000 characters of the corresponding log file. Open
the full log here:

```
C:\Users\jsoberanes\Projects\scorecard-pa\logs\wednesday_YYYYMMDD.log
C:\Users\jsoberanes\Projects\scorecard-pa\logs\friday_YYYYMMDD.log
```

To re-run manually after a failure: use `run_manual.bat` (Wednesday) or
the `run_agent.py --report friday` command (Friday).
