"""
KPI Engine Module
Calculates KPIs for new runs and compares against historical baselines.
Each KPI is a separate function - easy to add new ones.

IMPORTANT: All comparisons use 2025+ data only (comparison baseline).
Historical/aggregate queries use all available data.
"""

import pandas as pd
import numpy as np
from .comparator import calculate_baseline, compare_run_to_baseline, find_baseline_for_run


def _safe_val(run, col, default="N/A"):
    """Get a value from a run, returning default if NaN/None."""
    val = run.get(col)
    if val is None or (isinstance(val, (float, np.floating)) and pd.isna(val)):
        return default
    return val


def kpi_avg_rop(new_runs, baseline_df, config):
    """
    KPI: Average ROP
    Compare each new run's AVG_ROP against 2025+ baseline
    using multi-level fallback: HOLE_SIZE+Phase+BASIN+COUNTY+FORMATION+MOTOR → BASIN.
    """
    kpi_config = config["kpis"]["avg_rop"]
    if not kpi_config["enabled"]:
        return None

    comparison_levels = kpi_config["comparison_levels"]
    min_sample = config["comparison"]["min_sample_size"]

    # Analyze each new run with valid ROP
    results = []
    valid_runs = new_runs.dropna(subset=["AVG_ROP"])

    for idx, run in valid_runs.iterrows():
        baseline_row, match_level = find_baseline_for_run(
            run, baseline_df, comparison_levels, "AVG_ROP", min_sample
        )
        comparison = compare_run_to_baseline(run["AVG_ROP"], baseline_row, kpi_config)

        results.append({
            "index": idx,
            "operator": _safe_val(run, "OPERATOR", "Unknown"),
            "well": _safe_val(run, "WELL", "Unknown"),
            "rig": _safe_val(run, "RIG", "Unknown"),
            # Location group
            "basin": _safe_val(run, "BASIN"),
            "county": _safe_val(run, "COUNTY"),
            "formation": _safe_val(run, "FORMATION"),
            # Section group
            "hole_size": _safe_val(run, "HOLE_SIZE"),
            "phase": _safe_val(run, "Phase_CALC"),
            # Motor group
            "motor_model": _safe_val(run, "MOTOR_MODEL"),
            "lobe_stage": _safe_val(run, "LOBE/STAGE"),
            "motor_type2": _safe_val(run, "MOTOR_TYPE2"),
            "bend_hsg": _safe_val(run, "BEND_HSG"),
            # Bit group
            "bit_model": _safe_val(run, "BIT_MODEL"),
            # Depth group
            "depth_in": _safe_val(run, "DEPTH_IN"),
            "depth_out": _safe_val(run, "DEPTH_OUT"),
            "total_drill": _safe_val(run, "TOTAL_DRILL"),
            # KPI value
            "value": round(run["AVG_ROP"], 1),
            "match_level": match_level,
            **comparison,
        })

    results_df = pd.DataFrame(results)

    # Summary stats
    summary = {
        "kpi_name": "Average ROP",
        "unit": "ft/hr",
        "total_runs": len(valid_runs),
        "week_avg": round(valid_runs["AVG_ROP"].mean(), 1) if len(valid_runs) > 0 else 0,
        "week_median": round(valid_runs["AVG_ROP"].median(), 1) if len(valid_runs) > 0 else 0,
        "flagged_below": len(results_df[results_df["flag"] == "below"]) if len(results_df) > 0 else 0,
        "flagged_above": len(results_df[results_df["flag"] == "above"]) if len(results_df) > 0 else 0,
        "results": results_df,
    }

    # Basin breakdown
    if len(valid_runs) > 0:
        basin_avg = valid_runs.groupby("BASIN")["AVG_ROP"].agg(["count", "mean"]).round(1)
        basin_avg = basin_avg.sort_values("count", ascending=False)
        summary["basin_breakdown"] = basin_avg

    return summary


def kpi_longest_runs(new_runs, baseline_df, config):
    """
    KPI: Longest Runs (TOTAL_DRILL)
    Rank new runs by footage drilled, compare against 2025+ baseline.
    """
    kpi_config = config["kpis"]["longest_runs"]
    if not kpi_config["enabled"]:
        return None

    comparison_levels = kpi_config["comparison_levels"]
    top_n = kpi_config["top_n"]
    min_sample = config["comparison"]["min_sample_size"]

    # Get valid runs sorted by footage
    valid_runs = new_runs.dropna(subset=["TOTAL_DRILL"])
    valid_runs = valid_runs[valid_runs["TOTAL_DRILL"] > 0]
    sorted_runs = valid_runs.sort_values("TOTAL_DRILL", ascending=False)

    # Analyze top N runs
    results = []
    for idx, run in sorted_runs.head(top_n).iterrows():
        baseline_row, match_level = find_baseline_for_run(
            run, baseline_df, comparison_levels, "TOTAL_DRILL", min_sample
        )
        comparison = compare_run_to_baseline(run["TOTAL_DRILL"], baseline_row, kpi_config)

        results.append({
            "rank": len(results) + 1,
            "index": idx,
            "operator": _safe_val(run, "OPERATOR", "Unknown"),
            "well": _safe_val(run, "WELL", "Unknown"),
            "rig": _safe_val(run, "RIG", "Unknown"),
            # Location group
            "basin": _safe_val(run, "BASIN"),
            "county": _safe_val(run, "COUNTY"),
            "formation": _safe_val(run, "FORMATION"),
            # Section group
            "hole_size": _safe_val(run, "HOLE_SIZE"),
            "phase": _safe_val(run, "Phase_CALC"),
            # Motor group
            "motor_model": _safe_val(run, "MOTOR_MODEL"),
            "lobe_stage": _safe_val(run, "LOBE/STAGE"),
            "motor_type2": _safe_val(run, "MOTOR_TYPE2"),
            # KPI value
            "value": round(run["TOTAL_DRILL"], 0),
            "drilling_hours": round(run.get("DRILLING_HOURS", 0), 1) if pd.notna(run.get("DRILLING_HOURS")) else None,
            "avg_rop": round(run.get("AVG_ROP", 0), 1) if pd.notna(run.get("AVG_ROP")) else None,
            "match_level": match_level,
            **comparison,
        })

    results_df = pd.DataFrame(results)

    summary = {
        "kpi_name": "Longest Runs (Total Drill)",
        "unit": "ft",
        "total_runs": len(valid_runs),
        "top_n": top_n,
        "week_avg": round(valid_runs["TOTAL_DRILL"].mean(), 0) if len(valid_runs) > 0 else 0,
        "week_max": round(valid_runs["TOTAL_DRILL"].max(), 0) if len(valid_runs) > 0 else 0,
        "week_min": round(valid_runs["TOTAL_DRILL"].min(), 0) if len(valid_runs) > 0 else 0,
        "results": results_df,
    }

    return summary


def kpi_sliding_pct(new_runs, baseline_df, config):
    """
    KPI: Sliding Percentage (LAT phase only)
    Calculate SLIDE_DRILLED / TOTAL_DRILL * 100 for lateral runs.
    Compare against 2025+ baseline.
    """
    kpi_config = config["kpis"]["sliding_pct"]
    if not kpi_config["enabled"]:
        return None

    comparison_levels = kpi_config["comparison_levels"]
    phase_filter = [p.upper() for p in kpi_config["phase_filter"]]
    flag_percentile = kpi_config["flag_above_percentile"]
    min_sample = config["comparison"]["min_sample_size"]

    # Filter for LAT phase only
    lat_runs = new_runs[new_runs["Phase_CALC"].isin(phase_filter)].copy()

    # Calculate sliding % where data is available
    lat_runs = lat_runs.dropna(subset=["SLIDE_DRILLED", "TOTAL_DRILL"])
    lat_runs = lat_runs[lat_runs["TOTAL_DRILL"] > 0]
    lat_runs["slide_pct"] = (lat_runs["SLIDE_DRILLED"] / lat_runs["TOTAL_DRILL"]) * 100

    # Calculate baseline sliding % for LAT runs
    lat_baseline = baseline_df[baseline_df["Phase_CALC"].isin(phase_filter)].copy()
    lat_baseline = lat_baseline.dropna(subset=["SLIDE_DRILLED", "TOTAL_DRILL"])
    lat_baseline = lat_baseline[lat_baseline["TOTAL_DRILL"] > 0]
    lat_baseline["slide_pct"] = (lat_baseline["SLIDE_DRILLED"] / lat_baseline["TOTAL_DRILL"]) * 100

    # Historical percentile threshold
    hist_threshold = lat_baseline["slide_pct"].quantile(flag_percentile / 100) if len(lat_baseline) > 0 else None

    # Analyze each LAT run
    results = []
    for idx, run in lat_runs.iterrows():
        baseline_row, match_level = find_baseline_for_run(
            run, lat_baseline, comparison_levels, "slide_pct", min_sample
        )

        comparison = compare_run_to_baseline(run["slide_pct"], baseline_row, kpi_config)

        # Override flag: for sliding, ABOVE average is the concern
        if hist_threshold and run["slide_pct"] > hist_threshold:
            comparison["flag"] = "above"

        results.append({
            "index": idx,
            "operator": _safe_val(run, "OPERATOR", "Unknown"),
            "well": _safe_val(run, "WELL", "Unknown"),
            "rig": _safe_val(run, "RIG", "Unknown"),
            # Location group
            "basin": _safe_val(run, "BASIN"),
            "county": _safe_val(run, "COUNTY"),
            "formation": _safe_val(run, "FORMATION"),
            # Section group
            "hole_size": _safe_val(run, "HOLE_SIZE"),
            # Motor group
            "motor_model": _safe_val(run, "MOTOR_MODEL"),
            "lobe_stage": _safe_val(run, "LOBE/STAGE"),
            "motor_type2": _safe_val(run, "MOTOR_TYPE2"),
            # KPI values
            "total_drill": round(run["TOTAL_DRILL"], 0),
            "slide_drilled": round(run["SLIDE_DRILLED"], 0),
            "value": round(run["slide_pct"], 1),
            "match_level": match_level,
            **comparison,
        })

    results_df = pd.DataFrame(results)

    summary = {
        "kpi_name": "Sliding % (LAT Phase)",
        "unit": "%",
        "total_lat_runs": len(lat_runs),
        "total_new_runs": len(new_runs),
        "week_avg": round(lat_runs["slide_pct"].mean(), 1) if len(lat_runs) > 0 else 0,
        "week_median": round(lat_runs["slide_pct"].median(), 1) if len(lat_runs) > 0 else 0,
        "week_min": round(lat_runs["slide_pct"].min(), 1) if len(lat_runs) > 0 else 0,
        "week_max": round(lat_runs["slide_pct"].max(), 1) if len(lat_runs) > 0 else 0,
        "hist_threshold_pct": round(hist_threshold, 1) if hist_threshold else None,
        "flagged_above": len(results_df[results_df["flag"] == "above"]) if len(results_df) > 0 else 0,
        "results": results_df,
    }

    return summary


def run_all_kpis(new_runs, baseline_df, config):
    """Run all enabled KPIs and return results."""
    results = {}

    results["avg_rop"] = kpi_avg_rop(new_runs, baseline_df, config)
    results["longest_runs"] = kpi_longest_runs(new_runs, baseline_df, config)
    results["sliding_pct"] = kpi_sliding_pct(new_runs, baseline_df, config)

    return results
