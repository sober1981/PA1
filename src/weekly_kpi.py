"""
Weekly KPI Summary — per-hole-size breakdown of weekly runs.

For each hole size in the current week, groups runs by (MOTOR_TYPE2, JOB_TYPE, SERIES 20)
and computes per-row metrics, plus per-hole-size and grand totals with prev-week comparison.
"""

import pandas as pd

MOTOR_TYPE_ORDER = {"TDI CONV": 1, "CAM RENTAL": 2, "CAM DD": 3, "3RD PARTY": 4}

COL_TOTAL_HRS = "Total Hrs (C+D)"
COL_TOTAL_DRILL = "TOTAL_DRILL"
COL_DRILLING_HRS = "DRILLING_HOURS"
COL_SLIDE_DRILLED = "SLIDE_DRILLED"
COL_MY = "MY"
COL_INCIDENT = "INCIDENT_NUM"
COL_SERIES_20 = "SERIES 20"


def _safe_div(a, b):
    return float(a) / float(b) if b else 0.0


def _compute_metrics(df_group):
    """Compute aggregate metrics for a set of runs."""
    runs = len(df_group)
    total_hrs = df_group[COL_TOTAL_HRS].sum() if COL_TOTAL_HRS in df_group else 0.0
    total_drill = df_group[COL_TOTAL_DRILL].sum() if COL_TOTAL_DRILL in df_group else 0.0
    drilling_hrs = df_group[COL_DRILLING_HRS].sum() if COL_DRILLING_HRS in df_group else 0.0
    slide_drilled = df_group[COL_SLIDE_DRILLED].sum() if COL_SLIDE_DRILLED in df_group else 0.0

    my_vals = df_group[COL_MY].dropna() if COL_MY in df_group else pd.Series(dtype=float)
    my_avg = float(my_vals.mean()) if len(my_vals) > 0 else None

    incident_count = int(df_group[COL_INCIDENT].notna().sum()) if COL_INCIDENT in df_group else 0

    return {
        "runs": int(runs),
        "total_hrs": round(float(total_hrs), 2),
        "total_drill": int(round(float(total_drill))),
        "drilling_hrs": round(float(drilling_hrs), 2),
        "slide_drilled": float(slide_drilled),
        "avg_rop": round(_safe_div(total_drill, drilling_hrs), 2),
        "avg_slide_pct": round(_safe_div(slide_drilled, total_drill), 4),
        "avg_run_length": round(_safe_div(total_drill, runs), 1),
        "my_avg": round(my_avg, 2) if my_avg is not None else None,
        "incident_count": incident_count,
    }


def _empty_metrics():
    return {
        "runs": 0, "total_hrs": 0.0, "total_drill": 0, "drilling_hrs": 0.0,
        "slide_drilled": 0.0, "avg_rop": 0.0, "avg_slide_pct": 0.0,
        "avg_run_length": 0.0, "my_avg": None, "incident_count": 0,
    }


def compute_weekly_kpi(current_df, prev_df, week=None):
    """
    Build the Weekly KPI Summary structure.

    Returns:
      {
        "week": str,
        "grand_total_hrs": float,    # current week, all hole sizes
        "grand_total_drill": float,  # current week, all hole sizes
        "blocks": [
          {
            "hole_size": float,
            "rows": [ {motor_type, job_type, series_20, runs, total_hrs,
                       w_pct_hrs, g_pct_hrs, total_drill, w_pct_drill, g_pct_drill,
                       drilling_hrs, avg_rop, avg_slide_pct, avg_run_length,
                       my_avg, incident_count}, ... ],
            "curr_total": {...metrics, g_pct_hrs, g_pct_drill},
            "prev_total": {...metrics},
            "diff": {runs, total_hrs, total_hrs_pct, total_drill, total_drill_pct},
          }, ...
        ]
      }
    """
    cur = current_df[current_df["HOLE_SIZE"].notna()].copy() if "HOLE_SIZE" in current_df.columns else current_df.copy()
    prev = prev_df[prev_df["HOLE_SIZE"].notna()].copy() if "HOLE_SIZE" in prev_df.columns else prev_df.copy()

    # Grand totals (current week, all hole sizes)
    g_total_hrs = float(cur[COL_TOTAL_HRS].sum()) if COL_TOTAL_HRS in cur else 0.0
    g_total_drill = float(cur[COL_TOTAL_DRILL].sum()) if COL_TOTAL_DRILL in cur else 0.0

    # Grand totals (previous week, all hole sizes — including any that aren't in current)
    g_prev_hrs = float(prev[COL_TOTAL_HRS].sum()) if COL_TOTAL_HRS in prev else 0.0
    g_prev_drill = float(prev[COL_TOTAL_DRILL].sum()) if COL_TOTAL_DRILL in prev else 0.0
    g_prev_runs = int(len(prev))
    g_prev_incidents = int(prev[COL_INCIDENT].notna().sum()) if COL_INCIDENT in prev else 0

    blocks = []
    for hs in sorted(cur["HOLE_SIZE"].unique()):
        sub = cur[cur["HOLE_SIZE"] == hs]
        prev_sub = prev[prev["HOLE_SIZE"] == hs] if "HOLE_SIZE" in prev.columns else prev.iloc[0:0]

        w_total_hrs = float(sub[COL_TOTAL_HRS].sum()) if COL_TOTAL_HRS in sub else 0.0
        w_total_drill = float(sub[COL_TOTAL_DRILL].sum()) if COL_TOTAL_DRILL in sub else 0.0

        rows = []
        group_cols = ["MOTOR_TYPE2", "JOB_TYPE", COL_SERIES_20]
        for keys, grp in sub.groupby(group_cols, dropna=False):
            mt, jt, s20 = keys
            metrics = _compute_metrics(grp)
            rows.append({
                "motor_type": "" if pd.isna(mt) else str(mt),
                "job_type": "" if pd.isna(jt) else str(jt),
                "series_20": "" if pd.isna(s20) else str(s20),
                **metrics,
                "w_pct_hrs": _safe_div(metrics["total_hrs"], w_total_hrs),
                "g_pct_hrs": _safe_div(metrics["total_hrs"], g_total_hrs),
                "w_pct_drill": _safe_div(metrics["total_drill"], w_total_drill),
                "g_pct_drill": _safe_div(metrics["total_drill"], g_total_drill),
            })
        rows.sort(key=lambda r: (
            MOTOR_TYPE_ORDER.get(r["motor_type"].upper(), 99),
            r["job_type"],
            r["series_20"],
        ))

        curr_total = _compute_metrics(sub)
        curr_total["g_pct_hrs"] = _safe_div(curr_total["total_hrs"], g_total_hrs)
        curr_total["g_pct_drill"] = _safe_div(curr_total["total_drill"], g_total_drill)

        # Longest single run in this hole size (by TOTAL_DRILL)
        longest_run = None
        if len(sub) and COL_TOTAL_DRILL in sub.columns:
            valid = sub[sub[COL_TOTAL_DRILL].notna()]
            if len(valid):
                idx = valid[COL_TOTAL_DRILL].idxmax()
                run = sub.loc[idx]
                run_metrics = _compute_metrics(sub.loc[[idx]])

                def _fmt(col, as_int=False):
                    v = run.get(col)
                    if pd.isna(v):
                        return ""
                    if as_int:
                        try:
                            return str(int(v))
                        except (ValueError, TypeError):
                            return str(v)
                    return str(v)

                comment = " | ".join([
                    _fmt("JOB_NUM", as_int=True),
                    _fmt("OPERATOR"),
                    _fmt("Phase_CALC"),
                    _fmt("BEND_HSG"),
                ])

                longest_run = {
                    "motor_type": _fmt("MOTOR_TYPE2"),
                    "job_type": _fmt("JOB_TYPE"),
                    "series_20": _fmt(COL_SERIES_20),
                    **run_metrics,
                    "w_pct_hrs": _safe_div(run_metrics["total_hrs"], w_total_hrs),
                    "g_pct_hrs": _safe_div(run_metrics["total_hrs"], g_total_hrs),
                    "w_pct_drill": _safe_div(run_metrics["total_drill"], w_total_drill),
                    "g_pct_drill": _safe_div(run_metrics["total_drill"], g_total_drill),
                    "comment": comment,
                }

        prev_total = _compute_metrics(prev_sub) if len(prev_sub) else _empty_metrics()

        diff = {
            "runs": curr_total["runs"] - prev_total["runs"],
            "total_hrs": round(curr_total["total_hrs"] - prev_total["total_hrs"], 2),
            "total_hrs_pct": _safe_div(curr_total["total_hrs"] - prev_total["total_hrs"], prev_total["total_hrs"]),
            "total_drill": curr_total["total_drill"] - prev_total["total_drill"],
            "total_drill_pct": _safe_div(curr_total["total_drill"] - prev_total["total_drill"], prev_total["total_drill"]),
        }

        blocks.append({
            "hole_size": float(hs),
            "rows": rows,
            "curr_total": curr_total,
            "longest_run": longest_run,
            "prev_total": prev_total,
            "diff": diff,
        })

    return {
        "week": week,
        "grand_total_hrs": g_total_hrs,
        "grand_total_drill": g_total_drill,
        "grand_prev_total_hrs": g_prev_hrs,
        "grand_prev_total_drill": g_prev_drill,
        "grand_prev_runs": g_prev_runs,
        "grand_prev_incidents": g_prev_incidents,
        "blocks": blocks,
    }
