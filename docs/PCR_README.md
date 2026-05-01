# PCR — Performance Calculator and Ranking

Multi-variable, peer-group-based scoring system for drilling runs.
**Standalone, project-independent module** — designed to be reusable
in any future project that needs to rank drilling runs.

---

## What it does

Given a DataFrame of drilling runs (one row per run), PCR computes:

- A **0–100 composite performance score** per run.
- A **rank** (1 = best) across the dataset.
- A **performance band** (Excellent / Good / Average / Poor).
- The **peer group's 1-year rolling average PCR**, so you can see how
  the run compares to recent typical performance in similar conditions.
- An **audit trail** — which fallback level was used, how many peers,
  and the per-metric percentile contributions.

The method is fully driven by config — no code change to tune weights,
thresholds, or which variables count.

---

## How it differs from PA1

| | PA1 reports | PCR module |
|---|---|---|
| Purpose | Weekly executive summary (PDF + email) | Per-run analytics (Excel) |
| Output | PDF + KPI Excel + email | Excel only (3 sheets) |
| Typical use | Auto-scheduled Wed/Fri | Ad-hoc, on demand |
| Coupling | PA1-specific (Outlook, master-file paths, snapshot/state) | **Zero PA1 dependencies** — pure analytics |
| Reusability | Tied to this drilling workflow | Can be lifted into any future project |

PA1 might import PCR in the future to surface rankings inside the report,
but PCR does not depend on PA1.

---

## Files

Three artifacts make up the canonical method, all named "pcr":

| File | Purpose |
|---|---|
| [`docs/pcr.md`](pcr.md) | **Methodology spec** — variables, formula, baseline scope, peer-group logic, bands. The technical reference. |
| [`config/pcr.yaml`](../config/pcr.yaml) | **Tunable parameters** — weights, thresholds, fallback chain, output column names. Edit here to change behavior without touching code. |
| [`src/pcr.py`](../src/pcr.py) | **Implementation** — `rank_runs(df, config) → DataFrame`. Pure module, no PA1 imports. |

Plus the runner:

| File | Purpose |
|---|---|
| [`run_pcr.py`](../run_pcr.py) | CLI runner. Loads a master file, calls `rank_runs`, writes Excel. |

And the user-facing guide:

| File | Purpose |
|---|---|
| [`docs/HOW_TO_PCR.md`](HOW_TO_PCR.md) | Step-by-step: how to run, interpret output, tune weights, troubleshoot. |

---

## Method at a glance

1. **Pool**: all runs from `2025-01-01` onward (the QC'd Shared/Teams master is the recommended source).
2. **Peer group** for each run is found via a **multi-level fallback** of drilling-condition keys:
   `HOLE_SIZE → Phase_CALC → COUNTY → FORMATION → MOTOR_TYPE2 → FOOTAGE_BUCKET`. Walk down until at least 10 peers exist.
3. **Per metric** (Total Drill, Avg ROP, Rotate ROP, Slide ROP), compute the run's percentile rank within its peer group.
4. **Composite score** = weighted sum of the per-metric percentiles. Default weights:

   | Metric | Weight |
   |---|---|
   | Total Drill | 0.50 |
   | Avg ROP | 0.35 |
   | Rotate ROP | 0.10 |
   | Slide ROP | 0.05 |

   `Series 20 = Y` runs skip Slide ROP (Series-20 motors are rotate-only) — its weight is redistributed proportionally.

5. **Band**: Excellent ≥ p85, Good ≥ p65, Average ≥ p35, Poor < p35 (of the score distribution).
6. **Rolling reference**: per peer group, the 365-day rolling average PCR is attached so the run is compared against recent typical performance, not just absolute baseline.

The full spec is in [`docs/pcr.md`](pcr.md).

---

## Why each design choice

- **Peer-group baselining**: a 6.75" lateral run ≠ an 8.75" intermediate run, so we don't compare them. Each metric is normalized inside the run's own conditions.
- **Multi-level fallback**: if specific peers are sparse (e.g. unfamiliar basin), we relax the keys to broaden the comparison until ≥ 10 peers exist. Output records which level was used so the user can audit precision.
- **FOOTAGE_BUCKET** (short / medium / long, phase-specific): guards against "new bit" effect — a 200 ft run with high ROP isn't "great" if compared only to other short-footage peers.
- **Series-20-Y rule**: those motors rotate only by design; scoring them on Slide ROP is a category error. Skipped + weight redistributed.
- **Shared (QC'd) data preferred**: pre-QC Local files contain raw entry errors in ROP/drill that move a run several percentile points. Always score against the Friday-QC'd file.

---

## Status

- v0 design **locked**: peer fallback, footage buckets, weights, Series-20-Y skip.
- v0 **implementation**: `src/pcr.py` working, runner script in `run_pcr.py`.
- Sample output: `PCR Ranking - Week 26-W17.xlsx`.
- Future: add Motor Yield (MY) and Sliding % as additional metrics if needed; integrate into PA1 reports if useful.

---

## Reusing PCR in another project

`src/pcr.py` is dependency-free of PA1 internals. To use it elsewhere:

1. Copy `src/pcr.py` and `config/pcr.yaml` into the new project.
2. Adjust `config/pcr.yaml` for the new project's columns/priorities.
3. Call `rank_runs(df, config)` against any DataFrame:

```python
import yaml, pandas as pd
from pcr import rank_runs

config = yaml.safe_load(open("pcr.yaml"))
df = pd.read_excel("data.xlsx")
ranked = rank_runs(df, config)
print(ranked.nlargest(10, "pcr_score"))
```

If the analytics surface keeps growing, lift these modules into a small shared library (e.g. `scout-analytics` package) and pip-install across projects.
