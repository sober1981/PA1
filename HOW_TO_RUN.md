# PA1 - How to Run the Reports

## What PA1 produces

Every PA1 run generates **two attachments**:
- **PDF**: full performance report (Categories 1, 2, 3, and 4 on Friday).
  Category 1A now includes three KPI tables: **Summary**, **Detailed**,
  and **Longest Run** (per hole size).
- **Excel**: Weekly KPI Summary (the same three tables in a standalone
  workbook for QC and side-by-side comparison).

Both files are emailed to `jsoberanes@scoutdownhole.com` and saved to
the project folder.

## Prerequisites

- The master Excel file (`MASTER_MCS_MERGE_*.xlsx`) must already be in
  SharePoint / Teams sync or accessible locally.
- Outlook must be open on your machine (so the email can be sent).

---

## Manual Wednesday Report — the fast way

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

## Automatic Safety Net (Wednesday)

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

- Picks the QC'd version of Wednesday's master file from Teams sync.
- Compares it against the Wednesday snapshot (Category 4: QC Audit).
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
