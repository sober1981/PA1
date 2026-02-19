"""
Category 1: Week vs Previous Week
Compares current week performance against the immediately previous week.

Section A: Weekly Summary (runs, footage, hours by JOB_TYPE -> MOTOR_TYPE2)
Section B: Curves Analysis (Motor_KPI source, RUNS PER CUR)
Section C: Reason to POOH breakdown
"""

import pandas as pd
import numpy as np


def _group_summary(df, config):
    """
    Group runs by JOB_TYPE -> MOTOR_TYPE2 and aggregate metrics.
    Returns DataFrame with columns: JOB_TYPE, MOTOR_TYPE2, runs, total_drill, total_hrs
    """
    cfg = config["category1"]["weekly_summary"]
    primary = cfg["group_by_primary"]
    secondary = cfg["group_by_secondary"]

    if len(df) == 0:
        return pd.DataFrame(columns=[primary, secondary, "runs", "total_drill", "total_hrs"])

    grouped = df.groupby([primary, secondary], dropna=False).agg(
        runs=(primary, "size"),
        total_drill=("TOTAL_DRILL", "sum"),
        total_hrs=("Total Hrs (C+D)", "sum"),
    ).reset_index()

    grouped["total_drill"] = grouped["total_drill"].round(0)
    grouped["total_hrs"] = grouped["total_hrs"].round(1)

    # Sort by JOB_TYPE then runs descending
    grouped = grouped.sort_values([primary, "runs"], ascending=[True, False])
    return grouped


def section_a_weekly_summary(current_week, prev_week, config):
    """
    Section A: Weekly Summary
    Total runs, TOTAL_DRILL, Total Hrs (C+D) grouped by JOB_TYPE then MOTOR_TYPE2.
    Current week vs previous week side by side with delta.
    """
    current = _group_summary(current_week, config)
    previous = _group_summary(prev_week, config)

    cfg = config["category1"]["weekly_summary"]
    primary = cfg["group_by_primary"]
    secondary = cfg["group_by_secondary"]

    # Grand totals
    def _totals(df_raw):
        if len(df_raw) == 0:
            return {"runs": 0, "total_drill": 0, "total_hrs": 0}
        return {
            "runs": len(df_raw),
            "total_drill": round(df_raw["TOTAL_DRILL"].sum(), 0) if "TOTAL_DRILL" in df_raw.columns else 0,
            "total_hrs": round(df_raw["Total Hrs (C+D)"].sum(), 1) if "Total Hrs (C+D)" in df_raw.columns else 0,
        }

    current_totals = _totals(current_week)
    previous_totals = _totals(prev_week)
    delta_totals = {
        "runs": current_totals["runs"] - previous_totals["runs"],
        "total_drill": current_totals["total_drill"] - previous_totals["total_drill"],
        "total_hrs": round(current_totals["total_hrs"] - previous_totals["total_hrs"], 1),
    }

    return {
        "current": current,
        "previous": previous,
        "current_totals": current_totals,
        "previous_totals": previous_totals,
        "delta_totals": delta_totals,
    }


def section_b_curves(current_week, prev_week, config):
    """
    Section B: Curves Analysis (Motor_KPI source only)
    Counts 1-run curves, multi-run curves, QC-needed curves.
    """
    cfg = config["category1"]["curves"]
    source_filter = cfg["source_filter"]
    rpc_col = cfg["runs_per_cur_col"]
    phase_contains = cfg["phase_contains"]

    def _analyze_curves(df):
        if len(df) == 0:
            return {
                "total_motor_kpi": 0,
                "one_run_count": 0,
                "multi_run_count": 0,
                "qc_needed_count": 0,
                "total_with_rpc": 0,
                "operator_multi_run": pd.DataFrame(columns=["OPERATOR", "multi_run_count"]),
            }

        # Filter to Motor_KPI source
        motor_kpi = df[df["SOURCE"] == source_filter].copy() if "SOURCE" in df.columns else pd.DataFrame()

        if len(motor_kpi) == 0:
            return {
                "total_motor_kpi": 0,
                "one_run_count": 0,
                "multi_run_count": 0,
                "qc_needed_count": 0,
                "total_with_rpc": 0,
                "operator_multi_run": pd.DataFrame(columns=["OPERATOR", "multi_run_count"]),
            }

        # Runs with RUNS PER CUR data
        has_rpc = motor_kpi[motor_kpi[rpc_col].notna()] if rpc_col in motor_kpi.columns else pd.DataFrame()

        one_run = has_rpc[has_rpc[rpc_col] == 1] if len(has_rpc) > 0 else pd.DataFrame()
        multi_run = has_rpc[has_rpc[rpc_col] > 1] if len(has_rpc) > 0 else pd.DataFrame()

        # QC needed: Phase_CALC contains CUR AND SOURCE=Motor_KPI AND RUNS PER CUR is blank
        qc_mask = (
            motor_kpi["Phase_CALC"].str.contains(phase_contains, na=False) &
            motor_kpi[rpc_col].isna()
        ) if rpc_col in motor_kpi.columns else pd.Series(False, index=motor_kpi.index)
        qc_needed = motor_kpi[qc_mask]

        # Operator breakdown for multi-run curves
        if len(multi_run) > 0:
            op_multi = (multi_run.groupby("OPERATOR").size()
                       .reset_index(name="multi_run_count")
                       .sort_values("multi_run_count", ascending=False))
        else:
            op_multi = pd.DataFrame(columns=["OPERATOR", "multi_run_count"])

        return {
            "total_motor_kpi": len(motor_kpi),
            "one_run_count": len(one_run),
            "multi_run_count": len(multi_run),
            "qc_needed_count": len(qc_needed),
            "total_with_rpc": len(has_rpc),
            "operator_multi_run": op_multi,
        }

    current = _analyze_curves(current_week)
    previous = _analyze_curves(prev_week)

    delta = {
        "one_run_count": current["one_run_count"] - previous["one_run_count"],
        "multi_run_count": current["multi_run_count"] - previous["multi_run_count"],
        "qc_needed_count": current["qc_needed_count"] - previous["qc_needed_count"],
    }

    return {
        "current": current,
        "previous": previous,
        "delta": delta,
    }


def _classify_pooh(reason, classifications):
    """Classify a REASON_POOH value into a category."""
    if reason is None or (isinstance(reason, float) and pd.isna(reason)):
        return "unknown"

    reason_upper = str(reason).strip().upper()

    for category, values in classifications.items():
        for v in values:
            if v.upper() in reason_upper:
                return category
    return "other"


def section_c_reason_pooh(current_week, prev_week, config, report_type="wednesday"):
    """
    Section C: Reason to POOH breakdown
    Filter: SOURCE in [Motor_KPI, CAM_Run_Tracker]
    Wednesday uses REASON_POOH, Friday uses REASON_POOH_QC
    """
    cfg = config["category1"]["reason_pooh"]
    source_filter = cfg["source_filter"]
    reason_col = cfg["friday_col"] if report_type == "friday" else cfg["wednesday_col"]
    classifications = cfg["classifications"]

    def _analyze_pooh(df):
        if len(df) == 0 or "SOURCE" not in df.columns:
            return {
                "total_filtered": 0,
                "breakdown": pd.DataFrame(columns=["category", "count", "pct"]),
                "raw_breakdown": pd.DataFrame(columns=["reason", "count"]),
                "motor_by_operator": pd.DataFrame(columns=["OPERATOR", "motor_count", "total_count", "motor_pct"]),
            }

        # Filter to relevant sources
        filtered = df[df["SOURCE"].isin(source_filter)].copy()
        if len(filtered) == 0:
            return {
                "total_filtered": 0,
                "breakdown": pd.DataFrame(columns=["category", "count", "pct"]),
                "raw_breakdown": pd.DataFrame(columns=["reason", "count"]),
                "motor_by_operator": pd.DataFrame(columns=["OPERATOR", "motor_count", "total_count", "motor_pct"]),
            }

        # Classify each reason
        if reason_col in filtered.columns:
            filtered["pooh_category"] = filtered[reason_col].apply(
                lambda x: _classify_pooh(x, classifications)
            )
        else:
            filtered["pooh_category"] = "unknown"

        total = len(filtered)

        # Category breakdown
        cat_counts = filtered["pooh_category"].value_counts().reset_index()
        cat_counts.columns = ["category", "count"]
        cat_counts["pct"] = (cat_counts["count"] / total * 100).round(1)

        # Raw reason breakdown (top 15)
        if reason_col in filtered.columns:
            raw_counts = filtered[reason_col].value_counts().head(15).reset_index()
            raw_counts.columns = ["reason", "count"]
        else:
            raw_counts = pd.DataFrame(columns=["reason", "count"])

        # Motor issues by operator
        motor_runs = filtered[filtered["pooh_category"] == "motor_issues"]
        if len(motor_runs) > 0:
            op_totals = filtered.groupby("OPERATOR").size().reset_index(name="total_count")
            op_motor = motor_runs.groupby("OPERATOR").size().reset_index(name="motor_count")
            motor_by_op = op_totals.merge(op_motor, on="OPERATOR", how="inner")
            motor_by_op["motor_pct"] = (motor_by_op["motor_count"] / motor_by_op["total_count"] * 100).round(1)
            motor_by_op = motor_by_op.sort_values("motor_count", ascending=False)
        else:
            motor_by_op = pd.DataFrame(columns=["OPERATOR", "motor_count", "total_count", "motor_pct"])

        return {
            "total_filtered": total,
            "breakdown": cat_counts,
            "raw_breakdown": raw_counts,
            "motor_by_operator": motor_by_op,
        }

    current = _analyze_pooh(current_week)
    previous = _analyze_pooh(prev_week)

    return {
        "reason_col_used": reason_col,
        "current": current,
        "previous": previous,
    }


def run_category1(current_week, prev_week, config, report_type="wednesday"):
    """Run all Category 1 analyses."""
    if not config.get("category1", {}).get("enabled", True):
        return None

    return {
        "category": "Week vs Previous Week",
        "sections": {
            "A_weekly_summary": section_a_weekly_summary(current_week, prev_week, config),
            "B_curves": section_b_curves(current_week, prev_week, config),
            "C_reason_pooh": section_c_reason_pooh(current_week, prev_week, config, report_type),
        }
    }
