# PA1 - How to Generate the Wednesday Report

## Prerequisites

- The master Excel file (`MASTER_MCS_MERGE_*.xlsx`) must already be uploaded to SharePoint/Teams or available locally
- Outlook must be open on your machine (for the email to send)

## Step-by-Step Instructions

### Step 1: Open Command Prompt

Press `Win + R`, type `cmd`, press Enter.

### Step 2: Navigate to the project folder

```
cd C:\Users\jsoberanes\Projects\scorecard-pa
```

### Step 3: Run the Wednesday report

```
"C:\Users\jsoberanes\AppData\Local\Programs\Python\Python313\python.exe" run_agent.py --report wednesday
```

> Note: `python` alone won't work because it's not in your PATH. You must use the full path above.

### Step 4: Select the master file

PA1 will list available `MASTER_MCS_MERGE_*.xlsx` files sorted by date (newest first):

```
  Found 5 master file(s):

    [1] MASTER_MCS_MERGE_20260318_091234_Mar 18.xlsx
        Modified: 2026-03-18 09:12

    [2] MASTER_MCS_MERGE_20260311_123229_Mar 11.xlsx
        Modified: 2026-03-12 08:29

  Enter number [1-5] to select, or paste a full file path:
  >
```

Type the number of the file you want (usually `1` for the latest) and press Enter.

### Step 5: Wait for the report to finish

PA1 will process through 8 steps:
1. Load configuration
2. Locate master file
3. Load and clean data
4. Filter runs
5. Load comparison data
6. Run analysis (Categories 1-3)
7. Generate console report
8. Generate PDF + send email

### What happens after

- **PDF saved to**: `C:\Users\jsoberanes\Projects\scorecard-pa\PA1 - Wed Report [dates] - Week [##] - [filename].pdf`
- **Email sent to**: jsoberanes@scoutdownhole.com (with PDF attached)
- **Snapshot saved to**: `state\wednesday_snapshot.xlsx` (frozen copy for Friday QC audit)
- **State saved to**: `state\last_run.json` (so Friday knows which file to look for)

## Optional: Skip the email

If you just want the PDF without sending an email:

```
"C:\Users\jsoberanes\AppData\Local\Programs\Python\Python313\python.exe" run_agent.py --report wednesday --no-email
```

## Automatic Safety Net

If you forget to run the Wednesday report, it will auto-run at **2:00 PM on Wednesdays** via Windows Task Scheduler. The auto-run:
- Checks if you already ran it today (skips if so)
- Auto-picks the most recent master file (no interactive selection)
- Sends the report + saves snapshot as usual

## Friday Report

The Friday report runs automatically every **Friday at 10:00 AM** via Task Scheduler. No action needed from you. It:
- Picks up the QC'd version of Wednesday's file from Teams sync
- Compares it against the Wednesday snapshot (Category 4: QC Audit)
- Emails the full report with all 4 categories

## Quick Reference

| Action | Command |
|--------|---------|
| Wednesday report | `"C:\Users\jsoberanes\AppData\Local\Programs\Python\Python313\python.exe" run_agent.py --report wednesday` |
| Wednesday (no email) | `"C:\Users\jsoberanes\AppData\Local\Programs\Python\Python313\python.exe" run_agent.py --report wednesday --no-email` |
| Friday report (manual) | `"C:\Users\jsoberanes\AppData\Local\Programs\Python\Python313\python.exe" run_agent.py --report friday` |
| Specific week | `"C:\Users\jsoberanes\AppData\Local\Programs\Python\Python313\python.exe" run_agent.py --week 26-W11` |
