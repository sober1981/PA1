"""
Category 2: Monthly Highlights
Analyzes performance at the monthly level across operators.

Section A: Longest Runs by MOTOR_TYPE2
Section B: Monthly Summary by JOB_TYPE -> MOTOR_TYPE2
Section C: Fastest Sections by operator within HOLE_SIZE
Section D: Operator Success Rate (TD runs / total)
Section E: Motor Failures by Operator
Section F: Curve Success Rate per Operator
"""

import pandas as pd
import numpy as np


def _get_reason_col(config, report_type):
    """Get the correct REASON_POOH column based on report type."""
    if report_type == "friday":
        return config["category1"]["reason_pooh"]["friday_col"]
    return config["category1"]["reason_pooh"]["wednesday_col"]


def section_a_longest_runs(current_month, prev_month, config):
    """
    Section A: Top runs by TOTAL_DRILL for current and previous month.
    """
    top_n = config["category2"]["longest_runs"]["top_n"]

    def _top_runs(df):
        if len(df) == 0:
            return pd.DataFrame()
        valid = df.dropna(subset=["TOTAL_DRILL"])
        valid = valid[valid["TOTAL_DRILL"] > 0]
        top = valid.nlargest(top_n, "TOTAL_DRILL")

        result = []
        for rank, (_, run) in enumerate(top.iterrows(), 1):
            result.append({
                "rank": rank,
                "operator": run.get("OPERATOR", "Unknown"),
                "well": run.get("WELL", "Unknown"),
                "motor_type2": run.get("MOTOR_TYPE2", "N/A"),
                "total_drill": round(run["TOTAL_DRILL"], 0),
                "avg_rop": round(run["AVG_ROP"], 1) if pd.notna(run.get("AVG_ROP")) else None,
                "drilling_hours": round(run["DRILLING_HOURS"], 1) if pd.notna(run.get("DRILLING_HOURS")) else None,
                "basin": run.get("BASIN", "N/A"),
                "hole_size": run.get("HOLE_SIZE", "N/A"),
                "phase": run.get("Phase_CALC", "N/A"),
                "bend_hsg": run.get("BEND_HSG", "N/A"),
            })
        return pd.DataFrame(result)

    return {
        "current_month": _top_runs(current_month),
        "previous_month": _top_runs(prev_month),
    }


def section_b_monthly_summary(current_month, prev_month, config):
    """
    Section B: Monthly summary by JOB_TYPE -> MOTOR_TYPE2.
    Same grouping as Cat1 Section A but at monthly level.
    """
    primary = config["category1"]["weekly_summary"]["group_by_primary"]
    secondary = config["category1"]["weekly_summary"]["group_by_secondary"]

    def _group(df):
        if len(df) == 0:
            return pd.DataFrame(columns=[primary, secondary, "runs", "total_drill", "total_hrs"])

        grouped = df.groupby([primary, secondary], dropna=False).agg(
            runs=(primary, "size"),
            total_drill=("TOTAL_DRILL", "sum"),
            total_hrs=("Total Hrs (C+D)", "sum"),
        ).reset_index()
        grouped["total_drill"] = grouped["total_drill"].round(0)
        grouped["total_hrs"] = grouped["total_hrs"].round(1)
        return grouped.sort_values([primary, "runs"], ascending=[True, False])

    def _totals(df):
        if len(df) == 0:
            return {"runs": 0, "total_drill": 0, "total_hrs": 0}
        return {
            "runs": len(df),
            "total_drill": round(df["TOTAL_DRILL"].sum(), 0) if "TOTAL_DRILL" in df.columns else 0,
            "total_hrs": round(df["Total Hrs (C+D)"].sum(), 1) if "Total Hrs (C+D)" in df.columns else 0,
        }

    current = _group(current_month)
    previous = _group(prev_month)
    ct = _totals(current_month)
    pt = _totals(prev_month)

    return {
        "current": current,
        "previous": previous,
        "current_totals": ct,
        "previous_totals": pt,
        "delta_totals": {
            "runs": ct["runs"] - pt["runs"],
            "total_drill": ct["total_drill"] - pt["total_drill"],
            "total_hrs": round(ct["total_hrs"] - pt["total_hrs"], 1),
        },
    }


def section_c_fastest_sections(current_month, config):
    """
    Section C: Fastest sections (highest AVG_ROP) by operator within same HOLE_SIZE.
    """
    top_n = config["category2"]["fastest_sections"]["top_n"]

    if len(current_month) == 0:
        return {"by_hole_size": {}}

    valid = current_month.dropna(subset=["AVG_ROP", "HOLE_SIZE"])
    valid = valid[valid["AVG_ROP"] > 0]

    if len(valid) == 0:
        return {"by_hole_size": {}}

    results = {}
    for hole_size, group in valid.groupby("HOLE_SIZE"):
        if len(group) < 3:
            continue

        hole_avg = round(group["AVG_ROP"].mean(), 1)
        top = group.nlargest(top_n, "AVG_ROP")

        entries = []
        for _, run in top.iterrows():
            entries.append({
                "operator": run.get("OPERATOR", "Unknown"),
                "well": run.get("WELL", "Unknown"),
                "avg_rop": round(run["AVG_ROP"], 1),
                "total_drill": round(run["TOTAL_DRILL"], 0) if pd.notna(run.get("TOTAL_DRILL")) else None,
                "basin": run.get("BASIN", "N/A"),
                "county": run.get("COUNTY", "N/A"),
                "phase": run.get("Phase_CALC", "N/A"),
            })

        results[hole_size] = {
            "hole_avg_rop": hole_avg,
            "run_count": len(group),
            "top_operators": pd.DataFrame(entries),
        }

    return {"by_hole_size": results}


def section_d_operator_success(current_month, config, report_type="wednesday"):
    """
    Section D: Operator success rate = (TD runs / total runs) per operator.
    """
    reason_col = _get_reason_col(config, report_type)
    source_filter = config["category2"]["operator_analysis"]["source_filter"]
    success_values = [v.upper() for v in config["category2"]["operator_analysis"]["success_values"]]

    if len(current_month) == 0 or "SOURCE" not in current_month.columns:
        return pd.DataFrame(columns=["operator", "total_runs", "td_runs", "success_pct"])

    filtered = current_month[current_month["SOURCE"].isin(source_filter)].copy()
    if len(filtered) == 0 or reason_col not in filtered.columns:
        return pd.DataFrame(columns=["operator", "total_runs", "td_runs", "success_pct"])

    # Check success: reason value matches any success value
    filtered["is_td"] = filtered[reason_col].apply(
        lambda x: str(x).strip().upper() in success_values if pd.notna(x) else False
    )

    op_stats = filtered.groupby("OPERATOR").agg(
        total_runs=("OPERATOR", "size"),
        td_runs=("is_td", "sum"),
    ).reset_index()

    op_stats["success_pct"] = (op_stats["td_runs"] / op_stats["total_runs"] * 100).round(1)
    op_stats = op_stats.sort_values("success_pct", ascending=False)
    op_stats.columns = ["operator", "total_runs", "td_runs", "success_pct"]

    return op_stats


def section_e_motor_failures(current_month, config, report_type="wednesday"):
    """
    Section E: Motor failures by operator (failure runs / total runs as %).
    """
    reason_col = _get_reason_col(config, report_type)
    source_filter = config["category2"]["operator_analysis"]["source_filter"]
    failure_values = [v.upper() for v in config["category2"]["operator_analysis"]["motor_failure_values"]]

    if len(current_month) == 0 or "SOURCE" not in current_month.columns:
        return pd.DataFrame(columns=["operator", "total_runs", "failure_runs", "failure_pct"])

    filtered = current_month[current_month["SOURCE"].isin(source_filter)].copy()
    if len(filtered) == 0 or reason_col not in filtered.columns:
        return pd.DataFrame(columns=["operator", "total_runs", "failure_runs", "failure_pct"])

    # Check failure using contains matching (e.g., "MOTOR FAILURE" matches "MOTOR FAILURE - STATOR")
    filtered["is_failure"] = filtered[reason_col].apply(
        lambda x: any(fv in str(x).strip().upper() for fv in failure_values) if pd.notna(x) else False
    )

    op_stats = filtered.groupby("OPERATOR").agg(
        total_runs=("OPERATOR", "size"),
        failure_runs=("is_failure", "sum"),
    ).reset_index()

    op_stats["failure_pct"] = (op_stats["failure_runs"] / op_stats["total_runs"] * 100).round(1)
    # Only show operators with at least 1 failure
    op_stats = op_stats[op_stats["failure_runs"] > 0].sort_values("failure_pct", ascending=False)
    op_stats.columns = ["operator", "total_runs", "failure_runs", "failure_pct"]

    return op_stats


def section_f_curve_success(current_month, config):
    """
    Section F: Curve success rate per operator.
    Success = (RUNS PER CUR=1 / total non-null RUNS PER CUR) per operator.
    Only Motor_KPI source.
    """
    curve_source = config["category2"]["operator_analysis"]["curve_source_filter"]
    rpc_col = "RUNS PER CUR"

    if len(current_month) == 0 or "SOURCE" not in current_month.columns:
        return {"data": pd.DataFrame(columns=["operator", "total_curves", "one_run", "success_pct"]),
                "total_with_data": 0, "total_motor_kpi": 0}

    motor_kpi = current_month[current_month["SOURCE"] == curve_source].copy()
    if len(motor_kpi) == 0 or rpc_col not in motor_kpi.columns:
        return {"data": pd.DataFrame(columns=["operator", "total_curves", "one_run", "success_pct"]),
                "total_with_data": 0, "total_motor_kpi": len(motor_kpi)}

    # Only rows with RUNS PER CUR data
    has_data = motor_kpi[motor_kpi[rpc_col].notna()]
    if len(has_data) == 0:
        return {"data": pd.DataFrame(columns=["operator", "total_curves", "one_run", "success_pct"]),
                "total_with_data": 0, "total_motor_kpi": len(motor_kpi)}

    has_data = has_data.copy()
    has_data["is_one_run"] = has_data[rpc_col] == 1

    op_stats = has_data.groupby("OPERATOR").agg(
        total_curves=("OPERATOR", "size"),
        one_run=("is_one_run", "sum"),
    ).reset_index()

    op_stats["success_pct"] = (op_stats["one_run"] / op_stats["total_curves"] * 100).round(1)
    op_stats = op_stats.sort_values("success_pct", ascending=False)
    op_stats.columns = ["operator", "total_curves", "one_run", "success_pct"]

    return {
        "data": op_stats,
        "total_with_data": len(has_data),
        "total_motor_kpi": len(motor_kpi),
    }


def run_category2(current_month, prev_month, config, report_type="wednesday"):
    """Run all Category 2 analyses."""
    if not config.get("category2", {}).get("enabled", True):
        return None

    # Derive month labels from data
    cm_label = "Current Month"
    pm_label = "Previous Month"
    if len(current_month) > 0 and "DATE_OUT" in current_month.columns:
        dates = current_month["DATE_OUT"].dropna()
        if len(dates) > 0:
            cm_label = dates.iloc[0].strftime("%B %Y")
    if len(prev_month) > 0 and "DATE_OUT" in prev_month.columns:
        dates = prev_month["DATE_OUT"].dropna()
        if len(dates) > 0:
            pm_label = dates.iloc[0].strftime("%B %Y")

    return {
        "category": "Monthly Highlights",
        "current_month_label": cm_label,
        "previous_month_label": pm_label,
        "sections": {
            "A_longest_runs": section_a_longest_runs(current_month, prev_month, config),
            "B_monthly_summary": section_b_monthly_summary(current_month, prev_month, config),
            "C_fastest_sections": section_c_fastest_sections(current_month, config),
            "D_operator_success": section_d_operator_success(current_month, config, report_type),
            "E_motor_failures": section_e_motor_failures(current_month, config, report_type),
            "F_curve_success": section_f_curve_success(current_month, config),
        }
    }
