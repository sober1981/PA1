"""
Comparator Module
Calculates historical baselines and compares new runs against them.
Uses multi-level fallback: most specific grouping → least specific.
"""

import pandas as pd
import numpy as np


def calculate_baseline(baseline_df, group_by_cols, metric_col):
    """
    Calculate statistical baselines for a metric, grouped by specified dimensions.

    Returns a DataFrame with columns:
    - group_by columns
    - count, mean, median, std, p25, p75
    """
    # Filter to rows with valid metric values
    valid = baseline_df.dropna(subset=[metric_col])

    # Also need valid grouping columns
    for col in group_by_cols:
        valid = valid[valid[col].notna()]

    if len(valid) == 0:
        return pd.DataFrame()

    grouped = valid.groupby(group_by_cols)[metric_col].agg(
        count="count",
        mean="mean",
        median="median",
        std="std",
        p25=lambda x: x.quantile(0.25),
        p75=lambda x: x.quantile(0.75),
    ).reset_index()

    # Fill NaN std with 0 (single-value groups)
    grouped["std"] = grouped["std"].fillna(0)

    return grouped


def compare_run_to_baseline(run_value, baseline_row, config_kpi):
    """
    Compare a single run's value against its baseline group.

    Returns a dict with:
    - diff_pct: percentage difference from mean
    - flag: 'below', 'above', or None
    - baseline_mean, baseline_median, baseline_count
    """
    if baseline_row is None or (isinstance(baseline_row, pd.DataFrame) and len(baseline_row) == 0):
        return {
            "diff_pct": None,
            "flag": None,
            "baseline_mean": None,
            "baseline_median": None,
            "baseline_count": 0,
            "baseline_std": None,
            "match_level": None,
            "message": "No baseline data available",
        }

    b_mean = baseline_row["mean"]
    b_std = baseline_row["std"]
    b_count = baseline_row["count"]

    # Calculate percentage difference
    if b_mean > 0:
        diff_pct = ((run_value - b_mean) / b_mean) * 100
    else:
        diff_pct = 0

    # Determine flag
    flag = None
    flag_below = config_kpi.get("flag_below_std", 1.0)
    flag_above = config_kpi.get("flag_above_std", 1.5)

    if b_std > 0 and b_count >= 3:
        if run_value < (b_mean - flag_below * b_std):
            flag = "below"
        elif run_value > (b_mean + flag_above * b_std):
            flag = "above"

    return {
        "diff_pct": round(diff_pct, 1),
        "flag": flag,
        "baseline_mean": round(b_mean, 1),
        "baseline_median": round(baseline_row["median"], 1),
        "baseline_count": int(b_count),
        "baseline_std": round(b_std, 1),
    }


def find_baseline_for_run(run, baseline_df, comparison_levels, metric_col, min_sample_size=10):
    """
    Find the best matching baseline for a run using multi-level fallback.

    Tries each comparison level in order (most specific first).
    The first level that has enough baseline runs (>= min_sample_size) is used.

    Args:
        run: Series - the run to compare
        baseline_df: DataFrame - the 2025+ comparison baseline data
        comparison_levels: list of lists - each inner list is a set of grouping columns
        metric_col: str - the metric column to aggregate
        min_sample_size: int - minimum runs needed for a valid comparison

    Returns:
        (baseline_row_dict, match_level_str) or (None, "no_data")
    """
    if len(baseline_df) == 0:
        return None, "no_data"

    # Only consider rows with valid metric values
    valid_baseline = baseline_df.dropna(subset=[metric_col])
    if len(valid_baseline) == 0:
        return None, "no_data"

    for level_cols in comparison_levels:
        # Check that the run has values for all columns in this level
        run_vals = {}
        skip_level = False
        for col in level_cols:
            val = run.get(col)
            if pd.isna(val) if isinstance(val, (float, np.floating)) else (val is None):
                skip_level = True
                break
            run_vals[col] = val

        if skip_level:
            continue

        # Filter baseline to matching rows
        mask = pd.Series([True] * len(valid_baseline), index=valid_baseline.index)
        for col, val in run_vals.items():
            if col in valid_baseline.columns:
                mask = mask & (valid_baseline[col] == val)
            else:
                skip_level = True
                break

        if skip_level:
            continue

        matched = valid_baseline.loc[mask, metric_col]
        if len(matched) >= min_sample_size:
            level_label = "+".join(level_cols)
            return {
                "count": len(matched),
                "mean": matched.mean(),
                "median": matched.median(),
                "std": matched.std() if len(matched) > 1 else 0,
                "p25": matched.quantile(0.25),
                "p75": matched.quantile(0.75),
            }, level_label

    # Final fallback: overall average (no grouping)
    all_vals = valid_baseline[metric_col]
    if len(all_vals) > 0:
        return {
            "count": len(all_vals),
            "mean": all_vals.mean(),
            "median": all_vals.median(),
            "std": all_vals.std() if len(all_vals) > 1 else 0,
            "p25": all_vals.quantile(0.25),
            "p75": all_vals.quantile(0.75),
        }, "overall"

    return None, "no_data"
