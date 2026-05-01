# PA1 - Performance Analysis 1

Weekly drilling run analysis tool for Scout Downhole. Reads the master Excel file, identifies the latest week's runs, and generates automated performance reports organized into 4 analysis categories.

## How It Works

Every Wednesday, new weekly drilling runs (Monday-Sunday) are uploaded to a SharePoint/Teams master Excel file (`MASTER_MCS_MERGE_*.xlsx`). PA1 analyzes the new runs and generates a structured report comparing current performance against previous periods and historical baselines.

```
run_manual.bat                                         # Manual Wed run via double-click (interactive)
python run_agent.py --report wednesday                 # Wednesday report (interactive file + week pick if TTY)
python run_agent.py --report friday                    # Friday report (auto-picks QC'd Teams file)
python run_agent.py --week 26-W07                      # Specific week (ad-hoc, skips week prompt)
python run_agent.py --date-range 2026-02-09 2026-02-15 # Date range (ad-hoc)
python run_agent.py --file path/to/master.xlsx         # Custom file (skips file prompt)
python run_kpi_excel.py                                # Standalone Weekly KPI Excel (no PDF/email)
```

Add `--no-email` to any `run_agent.py` command to skip emailing.

Each report run produces **two attachments** that are emailed and saved to the project folder:
- **PDF**: full performance report with all 4 categories
- **Weekly KPI Excel**: per-hole-size Summary / Detailed / Longest Run tables (also embedded in Cat 1A of the PDF)

For step-by-step instructions, see [HOW_TO_RUN.md](HOW_TO_RUN.md).

## Three .bat wrappers — at a glance

| `.bat` file | When to use it | Launch | Prompts? | Email? |
|---|---|---|---|---|
| **`run_manual.bat`** | Manual Wednesday run (you double-click it) | Double-click in Explorer | Yes — file + week | Yes |
| **`run_wednesday.bat`** | Wednesday auto safety net at 14:00 if you forgot | Task Scheduler | No — auto-detect | Yes |
| **`run_friday.bat`** | Official Friday report at 10:00 | Task Scheduler | No — uses Wed-saved state | Yes |

All three call `run_agent.py` with different `--report` flags and run modes.
See [HOW_TO_RUN.md](HOW_TO_RUN.md) for the full per-bat behavior matrix.

## Report Structure (v2.5)

The report is organized into 4 categories:

### Category 1: Week vs Previous Week
| Section | Description |
|---------|-------------|
| **A - Weekly Summary (KPI tables)** | Three per-hole-size tables: **Summary** (totals + diff vs prev week, with green/yellow/red gradients on Total Runs/Hrs/Drill, reversed gradient on Total Incidents, red/green diff fonts, and an `OP w/ More Runs` column in `EXXON (4)` format), **Detailed** (per MOTOR_TYPE2 / JOB_TYPE / SERIES 20 with motor-type fills, per-hole-size winner highlight, and per-row `OP w/ More Runs`), **Longest Run** (sorted by Total Drill descending, with Operator / Job Number / Phase / Bend as separate columns). Same tables exported as a standalone Excel. |
| **B - Curves Analysis** | 1-run curves (best outcome) vs multi-run curves. Flags operators. SOURCE = Motor_KPI only. |
| **C - Reason to POOH** | Breakdown by classification (TD, ROP, Bit, Motor, MWD, BHA, Pressure, Other). Motor detail shows Operator, Hole Size, and SN. Wednesday uses REASON_POOH, Friday uses REASON_POOH_QC. |

### Category 2: Monthly Highlights
| Section | Description |
|---------|-------------|
| **A - Longest Runs** | Top 5 by TOTAL_DRILL, current + previous month, showing MOTOR_TYPE2 and BEND_HSG. |
| **B - Monthly Summary** | Same grouping as Cat1-A but monthly, current vs previous month. |
| **C - Fastest Sections** | Highest AVG_ROP by operator within same HOLE_SIZE (current month). |
| **D - Operator Success Rate** | (TD runs / total runs) per operator. |
| **E - Motor Failures** | (Motor failure runs / total runs) per operator. |
| **F - Curve Success Rate** | (1-run curves / total curves) per operator. |

### Category 3: Historical Analysis (2025+ Baseline)
| Section | Description |
|---------|-------------|
| **A - AVG ROP** | Rate of penetration vs baseline with multi-level fallback comparison. |
| **B - Longest Runs** | Top 5 by TOTAL_DRILL vs baseline. |
| **C - Sliding %** | SLIDE_DRILLED / TOTAL_DRILL for LAT phase only. |
| **D - Pattern Highlights** | Cross-variable-group highlights and lowlights with operator attribution. |

### Category 4: QC Audit (Friday Only)
Compares Wednesday (pre-QC) and Friday (post-QC) files cell by cell.

| Section | Description |
|---------|-------------|
| **A - Column Change Summary** | Which columns were corrected most often during QC. Identifies data source issues. |
| **B - QC Reviewer Workload** | RHC vs YGG: rows assigned, rows changed, total cell edits, average edits per row. |
| **C - Operator QC Trends** | Per operator: how many rows changed, cell edit count, and top corrected columns. |
| **D - Auto-Detected Patterns** | Systematic corrections (same operator+column repeated), broken columns (>50% corrected), high-effort QC rows, new/removed rows. |

## Wednesday / Friday Workflow

PA1 generates two reports per week, emailed to jsoberanes@scoutdownhole.com via Outlook:

### Wednesday - "First Look"
- **Trigger**: Manual (`run_manual.bat` or `run_agent.py --report wednesday`)
- **File selection**: Interactive (TTY) -- PA1 lists available master files, you pick the one you just uploaded
- **Week selection**: Interactive (TTY) -- PA1 lists the 10 most recent weeks in the data, defaults to latest
- **Data**: Raw, not yet QC'd
- **Categories**: 1-3 only (no QC Audit)
- **State**: Saves the selected filename + path to `state/last_run.json` for Friday reuse
- **Auto safety net**: Scheduled task at Wed 14:00 (`run_wednesday.bat`) -- skips if you already ran it manually today, otherwise auto-picks latest file + latest week, no prompts

### Friday - "QC'd Report"
- **Trigger**: Automatic -- Windows Task Scheduler, every Friday at 10:00 AM
- **File selection**: Auto -- reads Wednesday's filename from state, finds the same file in the Teams sync folder (QC'd version)
- **Data**: Same file as Wednesday, now QC'd by the team
- **Categories**: All 4 (includes QC Audit comparing Wed vs Fri)
- **Failure notification**: If the report fails for any reason, a failure email is sent with the error details instead. You always get an email on Friday (requires Outlook open + active session).

### Key Concept
Same filename, different location. Wednesday = local copy (pre-QC). Friday = Teams sync copy (post-QC).

### Weekly Timeline

| Day | Event | PA1 Action |
|-----|-------|------------|
| Mon-Sun | Drilling runs occur | -- |
| Wednesday | New week uploaded to SharePoint | User triggers Wed report (manual, interactive file pick) |
| Wed-Fri | Team QCs the data in Teams | -- |
| Friday 10 AM | QC complete | Fri report auto-generated from Teams sync + emailed (with QC Audit). On failure, error email sent instead. |

## Data Rules

| Scope | Date Filter | Purpose |
|-------|-------------|---------|
| **Comparison baseline** | 2025-01-01 to present (excl. target week) | Performance benchmarking, KPI flagging |
| **Full historical** | All available data (2017+) | Cumulative/aggregate queries |

2025+ data has significantly better quality than older data. All benchmarking comparisons use 2025+ only.

## Variable Groups (Category 3)

Comparisons use 5 variable groups with a **multi-level fallback chain** -- starting from the most specific grouping and falling back to broader groups until finding enough baseline runs (minimum 10).

| Group | Columns |
|-------|---------|
| Motor | MOTOR_MODEL, LOBE/STAGE, MOTOR_TYPE2, BEND_HSG |
| Bit | BIT_MODEL |
| Location | BASIN, COUNTY, FORMATION |
| Section | HOLE_SIZE, Phase_CALC |
| Depth | DEPTH_IN, DEPTH_OUT, TOTAL_DRILL |

## POOH Classifications (Category 1)

| Classification | REASON_POOH Values |
|---------------|-------------------|
| **TD** | TD, Section TD, Well TD |
| **ROP** | ROP, Build Rates |
| **Bit** | BIT |
| **Motor** | Motor, Motor Failure, Motor chunked |
| **MWD** | MWD, MWD Failure |
| **BHA** | BHA |
| **Pressure** | Pressure |
| **Other** | No Info, Hole problems, Stuck Pipe, etc. |

## Project Structure

```
scorecard-pa/
├── run_agent.py              # CLI entry point (--report wednesday/friday/--week/--date-range)
├── run_kpi_excel.py          # Standalone Weekly KPI Summary Excel generator (ad-hoc week analysis)
├── run_manual.bat            # Double-click manual Wednesday run (interactive prompts)
├── run_wednesday.bat         # Task Scheduler: Wed 14:00 auto-run (with failure email fallback)
├── run_friday.bat            # Task Scheduler: Fri 10:00 auto-run (with failure email fallback)
├── HOW_TO_RUN.md             # Step-by-step user guide
├── config/
│   └── settings.yaml         # Configuration (paths, thresholds, variable groups, POOH classes)
├── src/
│   ├── __init__.py
│   ├── data_loader.py        # Read Excel, clean data, filter by week/month, interactive file picker
│   ├── cat1_weekly.py        # Category 1: Week vs Previous Week (3 sections, includes KPI compute hook)
│   ├── cat2_monthly.py       # Category 2: Monthly Highlights (6 sections)
│   ├── kpi_engine.py         # Category 3: Historical KPIs + pattern highlights
│   ├── qc_audit.py           # Category 4: QC Audit — Wed vs Fri diff engine (Friday only)
│   ├── weekly_kpi.py         # Weekly KPI Summary: per-hole-size Summary / Detailed / Longest Run computation
│   ├── weekly_kpi_excel.py   # Weekly KPI Summary Excel writer (3 tables, gradients, motor-type fills)
│   ├── comparator.py         # Multi-level fallback baseline comparison
│   ├── state.py              # Wednesday/Friday state management (last_run.json)
│   ├── report.py             # Console output (all 4 categories)
│   ├── pdf_report.py         # PDF report (all 4 categories, fpdf2). Cat 1A renders KPI tables.
│   └── emailer.py            # Outlook COM auto-send with attachments (PDF + KPI Excel)
├── state/                    # Runtime state (auto-generated, git-ignored)
│   └── last_run.json
├── logs/                     # Scheduled-run logs (git-ignored)
├── tests/
│   └── __init__.py
├── requirements.txt          # pandas, openpyxl, pyyaml, fpdf2, pywin32, etc.
├── .env                      # SharePoint credentials (git-ignored)
└── .gitignore
```

## Pipeline

```
[1] Load config (settings.yaml)
[2] Locate master file:
    - Wednesday TTY: interactive file picker
    - Wednesday non-TTY (scheduled): auto-detect latest local
    - Friday: auto-pick QC'd version from Teams sync (same filename as Wed)
    - Ad-hoc: auto-detect latest or use --file
[3] Load & clean data
[4] Filter current week:
    - TTY (manual, not Friday): interactive week picker
    - Non-TTY or --week given: use given/latest
[5] Derive comparison time periods:
    - prev_week                   → Category 1
    - current_month, prev_month   → Category 2
    - baseline (2025+)            → Category 3
[6] Run analysis:
    - run_category1(...)  ← also computes Weekly KPI structure (per-hole-size tables)
    - run_category2(...)
    - run_category3(...)
    - run_qc_audit(...)   ← Friday only
[7] Generate reports:
    - Console output (all 4 categories)
    - PDF (Cat 1A renders the 3 KPI tables; rest of categories follow)
    - Weekly KPI Excel (standalone copy of Summary / Detailed / Longest Run)
[8] Email via Outlook (PDF + KPI Excel attached) unless --no-email
[9] Save state (Wednesday only)
```

## Flagging Logic (Category 3)

- **Underperforming**: Run is > 1.0 std dev below the baseline mean
- **Top performer**: Run is > 1.5 std dev above the baseline mean
- **Sliding % (LAT)**: Flagged if above the historical 75th percentile
- Minimum 3 runs in a group for flagging to activate

## Data Cleaning (on load)

- **BASIN**: Alias normalization ("TX-LA-MS Salt" -> "TX-LA-MS Sal")
- **HOLE_SIZE**: Outliers > 26" set to NaN (data entry errors like 222.0)
- **LOBE/STAGE**: Separator standardized to ":" (e.g., "6/7-7.6" -> "6/7:7.6")
- **FORMATION / COUNTY**: Uppercase normalized
- **BEND_HSG**: Converted to numeric (mixed string/numeric in source)
- **Phase_CALC**: Uppercase normalized, NaN preserved
- **ROTATE_DRILL**: Converted to numeric (object type in source)
- **Total Hrs (C+D)**: Converted to numeric
- **RUNS PER CUR**: Converted to numeric

## Setup

1. Install dependencies: `pip install -r requirements.txt`
2. Copy `.env.example` to `.env` and fill in SharePoint credentials (optional, for future use)
3. Sync the Teams document library to OneDrive (for Friday auto-pick)
4. Run: `python run_agent.py --report wednesday --no-email` to test

## Roadmap

- [x] Failure notification emails (v2.2 -- always get an email on Friday)
- [x] QC Audit snapshot fix + Wednesday scheduled task + HOW_TO_RUN guide (v2.3)
- [x] Weekly KPI Summary tables -- Summary / Detailed / Longest Run, in PDF Cat 1A and as standalone Excel; interactive week picker for manual runs (v2.4)
- [x] OP w/ More Runs column (Summary + Detailed); Longest Run sorted by Total Drill with Operator / Job Number / Phase / Bend split out of the comment (v2.5)
- [ ] Force Friday to use same week as Wednesday (prevent week drift from extra rows)
- [ ] Friday executive summary PDF (concise, management-ready)
- [ ] QC Audit historical trends (track QC workload week over week)
- [ ] More KPIs: rotate ROP, drilling hours efficiency, depth interval analysis
- [ ] Motor yield reference file integration
- [ ] Claude API integration for natural language insights
- [ ] RAG knowledge base with motor yields, formation characteristics
- [ ] Additional email recipients (team distribution list)
- [ ] PA2 - Reliability Agent: motor failures, yields vs expected, unplanned trips
