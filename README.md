# PA1 - Performance Agent 1

Weekly drilling run analysis agent for Scout Downhole. Reads the master Excel file, identifies the latest week's runs, and generates automated performance highlights.

## How It Works

Every Wednesday, new weekly drilling runs (Monday-Sunday) are uploaded to a SharePoint master Excel file (`MASTER_MCS_MERGE_*.xlsx`). The file syncs to local OneDrive. PA1 analyzes the new runs and generates a report with KPI comparisons against historical baselines.

```
python run_agent.py                                    # Auto-detect latest week
python run_agent.py --week 26-W07                      # Specific week
python run_agent.py --date-range 2026-02-09 2026-02-15 # Date range
python run_agent.py --file path/to/master.xlsx         # Custom file
python run_agent.py --week 26-W06 --pdf                # Generate PDF report
```

## Data Rules

| Scope | Date Filter | Purpose |
|-------|-------------|---------|
| **Comparison baseline** | 2025-01-01 to present (excl. target week) | Performance benchmarking, KPI flagging |
| **Full historical** | All available data (2017+) | Cumulative/aggregate queries |

2025+ data has significantly better quality than older data. All benchmarking comparisons use 2025+ only.

## Current KPIs (v1)

| KPI | Description | Phase Filter |
|-----|-------------|--------------|
| **AVG ROP** | Rate of penetration vs. baseline | All phases |
| **Longest Runs** | Top 5 by TOTAL_DRILL vs. baseline | All phases |
| **Sliding %** | SLIDE_DRILLED / TOTAL_DRILL | LAT only |

## Variable Groups

Comparisons use 5 variable groups to find the most relevant baseline. The agent uses a **multi-level fallback chain** — it starts with the most specific grouping and falls back to broader groups until it finds enough baseline runs (minimum 10).

### Motor Group
| Column | Fill Rate | Description |
|--------|-----------|-------------|
| MOTOR_MODEL | 99.8% | Motor model number (650, 712, 962, etc.) |
| LOBE/STAGE | 95.4% | Combined lobe/stage config (e.g., "6/7:7.6") |
| MOTOR_TYPE2 | 100% | Category: CAM RENTAL, TDI CONV, CAM DD, 3RD PARTY |
| BEND_HSG | 100% | Housing bend angle (numeric) |

### Bit Group
| Column | Fill Rate | Description |
|--------|-----------|-------------|
| BIT_MODEL | 77.8% | Specific bit model (997 unique values) |

### Location Group
| Column | Fill Rate | Description |
|--------|-----------|-------------|
| BASIN | 99.9% | Geographic basin (Delaware, Midland, TX-LA-MS Sal, etc.) |
| COUNTY | 99.4% | County name |
| FORMATION | 61.0% | Geological formation (uppercase normalized) |

Note: `FORM_FAM` column exists but is only populated for Midland and Delaware basins.

### Section Group
| Column | Fill Rate | Description |
|--------|-----------|-------------|
| HOLE_SIZE | 100% | Hole diameter in inches (bigger hole = shallower = faster ROP) |
| Phase_CALC | 79.7% | Well phase: LAT, CUR, INT, VER, SUR, etc. |

### Depth Group
| Column | Fill Rate | Description |
|--------|-----------|-------------|
| DEPTH_IN | 79.8% | Starting measured depth of the run |
| DEPTH_OUT | 79.8% | Ending measured depth of the run |
| TOTAL_DRILL | 99.8% | Total footage drilled in the run |

## Comparison Fallback Chain (AVG ROP Example)

```
Level 1: HOLE_SIZE + Phase_CALC + BASIN + COUNTY + FORMATION + MOTOR_MODEL
Level 2: HOLE_SIZE + Phase_CALC + BASIN + COUNTY + MOTOR_MODEL
Level 3: HOLE_SIZE + Phase_CALC + BASIN + MOTOR_MODEL
Level 4: HOLE_SIZE + Phase_CALC + BASIN
Level 5: HOLE_SIZE + BASIN
Level 6: BASIN + Phase_CALC
Level 7: BASIN
Final:   Overall average (all 2025+ data)
```

Each level requires a minimum of 10 baseline runs. The report shows which level was used for each comparison (the "match level").

## Data Cleaning (on load)

- **BASIN**: Alias normalization ("TX-LA-MS Salt" → "TX-LA-MS Sal")
- **HOLE_SIZE**: Outliers > 26" set to NaN (data entry errors like 222.0)
- **LOBE/STAGE**: Separator standardized to ":" (e.g., "6/7-7.6" → "6/7:7.6")
- **FORMATION**: Uppercase normalized for consistent grouping
- **COUNTY**: Uppercase normalized
- **BEND_HSG**: Converted to numeric (mixed string/numeric in source)
- **Phase_CALC**: Uppercase normalized, NaN preserved

## Project Structure

```
scorecard-agent/
├── run_agent.py              # CLI entry point
├── config/
│   └── settings.yaml         # Configuration (paths, thresholds, variable groups)
├── src/
│   ├── __init__.py
│   ├── data_loader.py        # Read Excel, clean data, filter by week
│   ├── kpi_engine.py         # KPI calculations (modular, one function per KPI)
│   ├── comparator.py         # Multi-level fallback baseline comparison
│   ├── report.py             # Console output
│   └── pdf_report.py         # PDF report generation (fpdf2)
├── tests/
│   └── __init__.py
├── requirements.txt          # pandas, openpyxl, pyyaml, fpdf2
└── .gitignore
```

## Pipeline

```
[1] Load config (settings.yaml)
[2] Find master file (auto-detect or --file)
[3] Load & clean data (Sheet1 only)
[4] Filter new runs (by week or date range)
[5] Load 2025+ comparison baseline
[6] Run KPIs (AVG ROP, Longest Runs, Sliding %)
[7] Generate report (console + optional PDF)
```

## Flagging Logic

- **Underperforming**: Run is > 1.0 std dev below the baseline mean
- **Top performer**: Run is > 1.5 std dev above the baseline mean
- **Sliding % (LAT)**: Flagged if above the historical 75th percentile
- Minimum 3 runs in a group for flagging to activate

## Future Roadmap

- [ ] More KPIs: rotate ROP, drilling hours efficiency, depth interval analysis
- [ ] Claude API integration for natural language insights
- [ ] RAG knowledge base with motor yield baselines, formation characteristics
- [ ] Reliability Agent (PA2): motor failures, yields vs expected, unplanned trips
- [ ] Excel report export for team sharing
- [ ] Auto-detect file changes for scheduled runs
- [ ] Motor yield reference file integration
