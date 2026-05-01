# How to Run PCR — Performance Calculator and Ranking

Quick guide for running PCR and reading the output. For the methodology
reference, see [pcr.md](pcr.md). For the high-level overview, see
[PCR_README.md](PCR_README.md).

---

## Prerequisites

- Python dependencies installed (`pip install -r requirements.txt`).
- A master Excel file accessible via:
  - **Shared (recommended)**: `Scout Downhole\DB Runs Update - Documents\General\...` (Teams sync, Friday-QC'd).
  - Or a Local OneDrive copy (pre-QC; not recommended for ranking).
- No Outlook required — PCR doesn't email.

---

## Quick start

### Easiest — interactive

```
cd C:\Users\jsoberanes\Projects\scorecard-pa
"C:\Users\jsoberanes\AppData\Local\Programs\Python\Python313\python.exe" run_pcr.py
```

Prompts:
1. **Source picker** — defaults to **Shared (QC'd)** when you press Enter.
2. **File confirmation** — shows the detected file's name, path, mod time, size. Press Enter or `y` to proceed, `n` to abort.
3. PCR computes (≈ 30 seconds for ~3 000 scoreable runs).
4. Excel saved as `PCR Ranking - Week {YY-W##}.xlsx` in the project folder.

### Skip prompts (good for repeat runs)

```
python run_pcr.py --source shared --week 26-W17 --yes
```

| Flag | Effect |
|---|---|
| `--source shared` | Skip source picker; use Shared/QC'd master. |
| `--source local` | Use Local pre-QC file (not recommended). |
| `--source browse` | Show numbered list of all candidates. |
| `--week 26-W17` | Use this week for the "This Week" sheet. |
| `--file path/to/file.xlsx` | Skip auto-pick; use this file. |
| `--output my.xlsx` | Custom output file name. |
| `--yes` / `-y` | Skip the file confirmation prompt. |

If `--week` is omitted, PCR auto-detects the latest week present in the data.

---

## Output Excel

Three sheets:

### 1. `This Week`

Every scored run for the chosen week, sorted by PCR descending.

| Column group | Columns |
|---|---|
| Identification | `DATE_OUT`, `Week #`, `OPERATOR`, `WELL`, `JOB_NUM` |
| Drilling conditions | `HOLE_SIZE`, `Phase_CALC`, `FOOTAGE_BUCKET`, `COUNTY`, `FORMATION`, `BASIN`, `MOTOR_TYPE2`, `MOTOR_MODEL`, `SERIES 20` |
| Raw metrics | `TOTAL_DRILL`, `AVG_ROP`, `ROTATE_ROP`, `SLIDE_ROP` |
| **PCR result** | **`pcr_score`**, **`pcr_rank`**, **`pcr_band`**, `pcr_peer_level`, `pcr_peer_n` |
| Per-metric audit | `pcr_total_drill_pct`, `pcr_avg_rop_pct`, `pcr_rotate_rop_pct`, `pcr_slide_rop_pct` |
| Rolling reference | `peer_avg_pcr_365d`, `delta_vs_365d` |

**Visual cues**:
- **`pcr_score`** — red→yellow→green color scale on values 0–100.
- **`pcr_band`** — solid fill: Excellent green, Good light green, Average yellow, Poor red.
- **`delta_vs_365d`** — red below avg, white at 0, green above.

### 2. `All Scored`

Same column shape, but all 2025+ runs in the master (typically 2 500 – 3 500 rows). Useful for filtering / pivoting in Excel.

### 3. `Weekly Aggregate`

One row per week.

| Column | What it tells you |
|---|---|
| `Week #` | ISO week (e.g., `26-W17`) |
| `n_runs` | scored runs that week |
| `avg_pcr` | mean PCR — should hover near 50 (since PCR is percentile-based) |
| `median_pcr`, `min_pcr`, `max_pcr` | distribution shape |
| `avg_peer_365d` | mean of each run's 365-day rolling peer average |
| `avg_delta` | week's avg PCR minus its 1-year peer avg — **positive = better-than-typical week** |

**Conditional fill** on `avg_pcr` so you can spot strong vs weak weeks at a glance.

---

## Interpreting individual scores

| `pcr_score` | Band | Meaning |
|---|---|---|
| ≥ 85 | **Excellent** | Top 15% in the dataset. Reference run for the peer group. |
| 65 – 85 | **Good** | Top third. Solid run. |
| 35 – 65 | **Average** | Middle. Within peer-group expectations. |
| < 35 | **Poor** | Bottom third. Worth investigating — bit, motor, formation surprise, etc. |

Score = **percentile within the peer group**, blended across the 4 metrics by weight. 50 = exactly median peer performance.

### Reading `pcr_peer_level`

| Level | Keys used | Comparison precision |
|---|---|---|
| `L1` | hole + phase + county + formation + motor + bucket | Tightest — same conditions in every dimension |
| `L2`–`L3` | drop one or two keys at a time | Still tight |
| `L4`–`L5` | drop bucket + formation, fall back to county | Moderate |
| `L6` | hole + phase + basin | Broad — only basin-level geo match |
| `L7` | hole + phase | Broadest — only physical section |

A score at L1 with 30+ peers is a **strong signal**. At L7 with only 12 peers, treat as **directional** — formation/motor mix isn't matched.

### Reading `delta_vs_365d`

- **+15** → run scored 15 PCR points above the peer group's 1-year rolling average. Strong outperformance.
- **0** → exactly typical for that peer group over the past year.
- **−20** → 20 PCR points below typical. Worth digging into.

---

## Tuning behavior — edit the YAML, not the code

Open [config/pcr.yaml](../config/pcr.yaml) to change:

| To do this | Edit this section |
|---|---|
| Change metric weights | `variables: <metric>: weight:` |
| Add/remove a metric | `variables:` (add or remove a block) |
| Change baseline cutoff date | `absolute_baseline.start_date` |
| Change minimum peer count | `peer_group.min_peer_runs` |
| Adjust footage bucket cutoffs | `footage_bucket.by_phase` |
| Change rolling window | `rolling_baseline.window_days` |
| Change band thresholds | `bands.{excellent,good,average}` |
| Add a skip rule for a metric | `variables: <metric>: skip_when:` |

After editing, re-run `python run_pcr.py` — no code changes needed.

If you make a structural change (e.g. add a metric that requires a new derived column), also update [docs/pcr.md](pcr.md) so the methodology spec stays canonical.

---

## Troubleshooting

| Symptom | Likely cause | Fix |
|---|---|---|
| `ERROR: 'PCR Ranking - Week ...xlsx' is open` | The output file is open in Excel | Close it and re-run, or pass `--output other_name.xlsx` |
| Auto-pick selects an old master | A user touched a stale file's mtime | Pass `--file <path>` explicitly to override, or use `--source shared` |
| Empty `This Week` sheet | The `--week` you chose has no scored runs in this master | Check the master actually contains that week (`--week` defaults to latest if omitted) |
| Many runs scored at L7 | Sparse peer groups in your basin/county | Either accept the broader comparison or lower `min_peer_runs` in the YAML |
| `pcr_score` all NaN | Filter is excluding everything (likely null AVG_ROP / TOTAL_DRILL) | Check the master for missing values; filters live under `filters.exclude_when` in the YAML |
| Long runtime (> 1 min) | Master has >10k runs | Acceptable; or filter to a specific year before running |
| Scores look skewed (mostly 50s) | Peer groups are tiny so percentile resolution is coarse | Increase `min_peer_runs` (less specific peer groups) or accept that small groups give blunt signal |

---

## Quick reference

| Action | Command |
|---|---|
| Default run (Shared file, latest week) | `python run_pcr.py` |
| Specific week, skip prompts | `python run_pcr.py --source shared --week 26-W17 --yes` |
| Run on a specific file | `python run_pcr.py --file "C:\path\to\master.xlsx" --yes` |
| Custom output name | `python run_pcr.py --output "weekly_review.xlsx"` |
| Show top 10 in Python | `import yaml, pandas as pd; from src.pcr import rank_runs; ranked = rank_runs(pd.read_excel(p), yaml.safe_load(open("config/pcr.yaml"))); print(ranked.nlargest(10, "pcr_score"))` |

> Prepend each Python command with `"C:\Users\jsoberanes\AppData\Local\Programs\Python\Python313\python.exe"` and `cd` first to the project folder.
