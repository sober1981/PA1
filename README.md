# PA1 - Performance Analysis 1

Weekly drilling run analysis tool for Scout Downhole. Reads the master Excel file, identifies the latest week's runs, and generates automated performance reports organized into 4 analysis categories.

## How It Works

Every Wednesday, new weekly drilling runs (Monday-Sunday) are uploaded to a SharePoint/Teams master Excel file (`MASTER_MCS_MERGE_*.xlsx`). PA1 analyzes the new runs and generates a structured report comparing current performance against previous periods and historical baselines.

```
python run_agent.py --report wednesday                 # Wednesday report (interactive file pick)
python run_agent.py --report friday                    # Friday report (auto-picks QC'd Teams file)
python run_agent.py --week 26-W07                      # Specific week (ad-hoc)
python run_agent.py --date-range 2026-02-09 2026-02-15 # Date range (ad-hoc)
python run_agent.py --file path/to/master.xlsx         # Custom file
```

Add `--no-email` to any command to skip emailing and just generate console + PDF output.

## Report Structure (v2.1)

The report is organized into 4 categories:

### Category 1: Week vs Previous Week
| Section | Description |
|---------|-------------|
| **A - Weekly Summary** | Run count, total footage, total hours by JOB_TYPE and MOTOR_TYPE2. Current vs previous week with deltas. |
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
- **Trigger**: Manual
- **File selection**: Interactive -- PA1 lists available master files, you pick the one you just uploaded (local folder)
- **Data**: Raw, not yet QC'd
- **Categories**: 1-3 only (no QC Audit)
- **State**: Saves the selected filename + path to `state/last_run.json` for Friday reuse

### Friday - "QC'd Report"
- **Trigger**: Automatic -- Windows Task Scheduler, every Friday at 10:00 AM
- **File selection**: Auto -- reads Wednesday's filename from state, finds the same file in the Teams sync folder (QC'd version)
- **Data**: Same file as Wednesday, now QC'd by the team
- **Categories**: All 4 (includes QC Audit comparing Wed vs Fri)

### Key Concept
Same filename, different location. Wednesday = local copy (pre-QC). Friday = Teams sync copy (post-QC).

### Weekly Timeline

| Day | Event | PA1 Action |
|-----|-------|------------|
| Mon-Sun | Drilling runs occur | -- |
| Wednesday | New week uploaded to SharePoint | User triggers Wed report (manual, interactive file pick) |
| Wed-Fri | Team QCs the data in Teams | -- |
| Friday 10 AM | QC complete | Fri report auto-generated from Teams sync + emailed (with QC Audit) |

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
├── run_friday.bat            # Batch script for Windows Task Scheduler (Fri 10 AM)
├── config/
│   └── settings.yaml         # Configuration (paths, thresholds, variable groups, POOH classes)
├── src/
│   ├── __init__.py
│   ├── data_loader.py        # Read Excel, clean data, filter by week/month, interactive file picker
│   ├── cat1_weekly.py        # Category 1: Week vs Previous Week (3 sections)
│   ├── cat2_monthly.py       # Category 2: Monthly Highlights (6 sections)
│   ├── kpi_engine.py         # Category 3: Historical KPIs + pattern highlights
│   ├── qc_audit.py           # Category 4: QC Audit — Wed vs Fri diff engine (Friday only)
│   ├── comparator.py         # Multi-level fallback baseline comparison
│   ├── state.py              # Wednesday/Friday state management (last_run.json)
│   ├── report.py             # Console output (all 4 categories)
│   ├── pdf_report.py         # PDF report (all 4 categories, fpdf2)
│   └── emailer.py            # Outlook COM auto-send with PDF attachment
├── state/                    # Runtime state (auto-generated, git-ignored)
│   └── last_run.json
├── logs/                     # Friday auto-run logs (git-ignored)
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
    - Wednesday: interactive picker (user selects from numbered list)
    - Friday: auto-pick from Teams sync (same filename as Wednesday)
    - Ad-hoc: auto-detect latest or use --file
[3] Load & clean data
[4] Derive time periods:
    - current_week, prev_week     → Category 1
    - current_month, prev_month   → Category 2
    - baseline (2025+)            → Category 3
[5] Run analysis:
    - run_category1(current_week, prev_week, config, report_type)
    - run_category2(current_month, prev_month, config, report_type)
    - run_category3(current_week, baseline, config)
    - run_qc_audit(wed_df, fri_df, config, week_start, week_end)  ← Friday only
[6] Generate reports (console + PDF)
[7] Email via Outlook (unless --no-email)
[8] Save state (Wednesday only)
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

- [ ] Friday executive summary PDF (concise, management-ready)
- [ ] QC Audit historical trends (track QC workload week over week)
- [ ] More KPIs: rotate ROP, drilling hours efficiency, depth interval analysis
- [ ] Motor yield reference file integration
- [ ] Claude API integration for natural language insights
- [ ] RAG knowledge base with motor yields, formation characteristics
- [ ] Additional email recipients (team distribution list)
- [ ] PA2 - Reliability Agent: motor failures, yields vs expected, unplanned trips
