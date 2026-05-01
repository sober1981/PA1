# PCR — Performance Calculator and Ranking

A reusable, project-independent method to score and rank drilling runs
across multiple performance variables. Designed to be portable: PA1 uses it
today; future projects can drop the same module in.

> **Status:** v0 in progress. Peer definition + baseline locked.
> Metric weights and output deltas still to be chosen.
> This document is the canonical reference — update it alongside any
> change to `src/pcr.py` or `config/pcr.yaml`.

## What it does

Given a DataFrame of drilling runs (one row per run), PCR computes a
**composite performance score** per run by comparing each metric against
the run's **peer group** (other runs drilled in similar conditions),
then aggregates the per-metric scores into a single 0–100 score.

The method is fully driven by [`config/pcr.yaml`](../config/pcr.yaml) — no
code change needed to tune weights, thresholds, or which variables are
included.

---

## 1. Inputs / outputs

**Inputs**
- `df` — pandas DataFrame of drilling runs (must contain columns for the
  metrics, peer-group keys, and `DATE_OUT`).
- `config` — Python dict loaded from `config/pcr.yaml`.

**Outputs** — same DataFrame with these columns appended:

| Column | Description |
|---|---|
| `pcr_score` | Composite score, 0–100 (higher = better). |
| `pcr_rank` | Rank within the dataset (1 = best). |
| `pcr_band` | Excellent / Good / Average / Poor. |
| `pcr_peer_level` | Which fallback level the peer group came from (transparency). |
| `pcr_peer_n` | Number of peers used in the comparison. |
| `pcr_<metric>_pct` | Per-metric percentile rank within the peer group (audit trail). |
| `peer_avg_pcr_365d` | The peer group's rolling 365-day average PCR. |
| `delta_vs_365d` | This run's PCR minus the peer group's 365-day rolling avg. |

---

## 2. Baseline scope

**Absolute baseline = all runs from 2025-01-01 onward** (matches PA1's
existing Cat 3 baseline; pre-2025 data is excluded due to lower data
quality). Each run's `pcr_score` is computed against peer runs from this
2025+ pool.

### Data source — always use Friday-QC'd masters

PCR rankings are sensitive to data-entry errors in `AVG_ROP`,
`TOTAL_DRILL`, `ROTATE_ROP`, and `SLIDE_ROP` because those are the
scoring inputs. Before QC, raw entries can contain typos that move a
run several percentile points.

**Always run PCR against the Shared / Teams-synced master file** —
that's the version updated each Friday after the team's QC pass. The
Local OneDrive copies are pre-QC and may contain raw errors that skew
the rankings.

`run_pcr.py` defaults to `--source shared`. To override for an ad-hoc
analysis, pass `--source local` or `--file`.

**Rolling baseline = 365 days** (1-year). Used for the
`peer_avg_pcr_365d` column — for each peer group, average the PCR scores
of peer runs whose `DATE_OUT` is within 365 days before the target run's
`DATE_OUT`.

---

## 3. Peer group definition

A **peer** is another run that matches the target's drilling conditions.
PCR uses a multi-level fallback chain — most-specific first, falling back
to broader groups until the target has at least `min_peer_runs` peers
(default 10).

### Peer-group keys (in priority order)

1. `HOLE_SIZE`
2. `Phase_CALC`
3. `COUNTY` (preferred over BASIN — gives finer geography)
4. `FORMATION` (skipped automatically when null on the target run)
5. `MOTOR_TYPE2`
6. `FOOTAGE_BUCKET` (derived — see §4 below)

### Multi-level fallback chain

PCR walks down the list and stops at the first level with ≥ `min_peer_runs`:

```
L1: HOLE_SIZE + Phase_CALC + COUNTY + FORMATION + MOTOR_TYPE2 + FOOTAGE_BUCKET
L2: HOLE_SIZE + Phase_CALC + COUNTY + FORMATION + FOOTAGE_BUCKET
L3: HOLE_SIZE + Phase_CALC + COUNTY + FOOTAGE_BUCKET
L4: HOLE_SIZE + Phase_CALC + COUNTY + FORMATION
L5: HOLE_SIZE + Phase_CALC + COUNTY
L6: HOLE_SIZE + Phase_CALC + BASIN
L7: HOLE_SIZE + Phase_CALC
```

If the target has any null in a level's keys, that level is skipped and
the chain continues to the next.

### Worked example

**Target:** `LOGOS — Ignacio 33-7 29P #3H` (8.5" INT, La Plata, MANCOS,
CAM DD, 10,518 ft = bucket "long").

| Level | Peers in 2025+ pool | Notes |
|---|---|---|
| L1–L3 | 0 | bucket too sparse in La Plata |
| L4 | 1 | too few |
| L5 | 1 | too few |
| L6 (basin Other) | 6 | still < 10, fall down |
| **L7 (hole+phase)** | **19** | ✓ used |

So this run is scored against the 19 8.5" INT peers in 2025+ (Comstock,
MITSUI, etc.). The output records `pcr_peer_level = L7` and `pcr_peer_n = 19`
so the user knows the comparison was broad.

For a Permian 6.75" LAT run in Karnes / Eagle Ford / TDI CONV, the chain
typically stops at L1 or L2 with dozens of tightly-matched peers.

---

## 4. FOOTAGE_BUCKET (derived)

Derived per run at preprocessing time. Buckets are phase-specific because
"short" means different things on a surface section vs a lateral.

| Phase_CALC | Short | Medium | Long |
|---|---|---|---|
| SUR | < 500 | 500 – 1500 | > 1500 |
| INT | < 2000 | 2000 – 6000 | > 6000 |
| CUR | < 500 | 500 – 1500 | > 1500 |
| LAT | < 3000 | 3000 – 7000 | > 7000 |
| Multi-phase (e.g. INT-CUR-LAT) | (null — bucket-aware levels skip these) | | |

This guards against "new bit" effect where a short, hot run inflates ROP.
A 2,000 ft Granite run is now compared only to other **medium-footage**
Granite runs, not to 18,000 ft Wolfcamp runs.

---

## 5. Variables / metrics (v0)

The composite PCR uses **4 metrics**:

| Metric | Column | Direction | Weight | Skip when |
|---|---|---|---|---|
| Total Drill | `TOTAL_DRILL` | higher is better | **0.50** | — |
| Avg ROP | `AVG_ROP` | higher is better | **0.35** | — |
| Rotate ROP | `ROTATE_ROP` | higher is better | **0.10** | — |
| Slide ROP | `SLIDE_ROP` | higher is better | **0.05** | **`SERIES 20` = `Y`** |
| **Total** | | | **1.00** | |

### Skip rule — Series 20 motors

Series-20 motors are rotate-only by design, so Slide ROP is meaningless on
them. When `SERIES 20 = "Y"` on a run:
- Slide ROP is **skipped** (not used for that run's score).
- Its 0.05 weight is **redistributed proportionally** to the remaining
  three metrics, keeping the weights summed to 1.0:

| Metric | Normal weight | Series-20-Y weight |
|---|---|---|
| Total Drill | 0.50 | 0.526 |
| Avg ROP | 0.35 | 0.368 |
| Rotate ROP | 0.10 | 0.105 |
| Slide ROP | 0.05 | — (skipped) |

### Algorithm per metric

1. Find the run's peer group (per §3).
2. Compute the run's percentile rank for the metric within the peers
   (0 = worst, 100 = best for `higher_is_better`).
3. Record the per-metric score as `pcr_<metric>_pct`.

### Not in v0 (intentionally dropped)

- **Motor Yield (MY)** — may be added later if motor-side analysis is needed.
- **Sliding %** — phase-conditional; deferred for now.

These can be added by appending entries to `variables:` in the YAML and
rebalancing weights — no code change needed.

---

## 6. Composite formula

```
pcr_score = Σ ( weight_i × score_i )
```

where `score_i` is the per-metric percentile rank (after applying
direction). Weights are configured in `config/pcr.yaml` and **must sum
to 1.0** (PCR will renormalize if they don't).

If a metric is "skipped" for this run (e.g. Slide % on a multi-phase run),
its weight is redistributed proportionally to the remaining metrics so
the total stays 1.0.

---

## 7. Bands

By default, the score is bucketed by percentile of the score distribution
across all scored runs:

| Band | Cutoff |
|---|---|
| Excellent | score percentile ≥ 0.85 |
| Good | 0.65 – 0.85 |
| Average | 0.35 – 0.65 |
| Poor | < 0.35 |

Adjustable in `config/pcr.yaml`.

---

## 8. Filters

Runs are excluded from scoring (not just from peer pools) when:

- `AVG_ROP` or `TOTAL_DRILL` is null (insufficient data to score).
- Any other null-column or source filter listed under
  `filters.exclude_when` in `config/pcr.yaml`.

---

## 9. Rolling 365-day average per peer group

After every run in the 2025+ pool has a `pcr_score`, PCR computes for each
target run:

- The set of peers (using the same fallback chain).
- The peers whose `DATE_OUT` falls within 365 days **before** the target's
  `DATE_OUT` (rolling, not centered).
- The average PCR of those peers → `peer_avg_pcr_365d`.
- `delta_vs_365d = pcr_score − peer_avg_pcr_365d`.

This lets weekly / monthly / yearly aggregate reports compare new runs
against the **moving 1-year peer average**, so trends are visible
without redefining the absolute baseline every period.

---

## 10. Usage (once `src/pcr.py` exists)

```python
import yaml
import pandas as pd
from src.pcr import rank_runs

config = yaml.safe_load(open("config/pcr.yaml"))
df = pd.read_excel("path/to/master.xlsx")
ranked = rank_runs(df, config)

# Top 10 runs across the dataset
print(ranked.nsmallest(10, "pcr_rank"))

# This week's runs vs their 1-year peer average
this_week = ranked[ranked["Week #"] == "26-W17"]
print(this_week[["WELL", "pcr_score", "peer_avg_pcr_365d", "delta_vs_365d"]])
```

---

## 11. Reusing PCR in another project

Module is designed with **zero dependencies on PA1 internals** (no PDF
code, no Outlook, no master-file path assumptions). To reuse:

1. Copy `src/pcr.py` and `config/pcr.yaml` into the new project.
2. Adjust `config/pcr.yaml` for the new project's columns / scoring
   priorities.
3. Call `rank_runs(df, config)` against any DataFrame.

---

## Change log

- *2026-05-02 — v0 design locked.* Peer fallback (L1–L7), footage
  buckets, absolute baseline = 2025+, rolling = 365d, 4 metrics
  (Total Drill 0.50 / Avg ROP 0.35 / Rotate ROP 0.10 / Slide ROP 0.05),
  Series-20-Y skip rule for Slide ROP with proportional weight
  redistribution. **MY and Sliding % deliberately excluded from v0.**
