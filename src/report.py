"""
Report Generator Module
Formats analysis results into readable console output.
Organized into 3 categories:
  Category 1: Week vs Previous Week
  Category 2: Monthly Highlights
  Category 3: Historical Analysis (2025+ Baseline)
"""

import pandas as pd


def _safe_str(val, default="N/A"):
    """Convert a value to string, handling NaN/None gracefully."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return default
    s = str(val)
    if s.lower() in ("nan", "none", ""):
        return default
    return s


def _fmt_hole(val):
    """Format hole size for display."""
    s = _safe_str(val)
    if s == "N/A":
        return s
    try:
        return f'{float(s):g}"'
    except (ValueError, TypeError):
        return s


def _delta_str(val, unit="", fmt=".0f"):
    """Format a delta value with +/- sign."""
    if val > 0:
        return f"+{val:{fmt}}{unit}"
    elif val < 0:
        return f"{val:{fmt}}{unit}"
    return f"0{unit}"


# =========================================================================
# Header / Footer
# =========================================================================

def print_header(meta):
    print()
    print("=" * 80)
    print(f"  PA1 - WEEKLY PERFORMANCE REPORT")
    print(f"  Week {meta['week']}  |  {meta['week_start'].date()} to {meta['week_end'].date()}")
    print(f"  Master: {meta['master_filename']}")
    print("=" * 80)
    print(f"  Total New Runs: {meta['total_new_runs']}")
    print(f"  Report Type: {meta['report_type'].title()}")
    print()


def print_footer():
    print()
    print("=" * 80)
    print("  End of Report")
    print("=" * 80)
    print()


def print_category_header(number, title):
    print()
    print("*" * 80)
    print(f"  CATEGORY {number}: {title.upper()}")
    print("*" * 80)


# =========================================================================
# Category 1: Week vs Previous Week
# =========================================================================

def print_cat1_section_a(section):
    print()
    print("-" * 80)
    print("  1A. WEEKLY SUMMARY (JOB_TYPE -> MOTOR_TYPE2)")
    print("-" * 80)

    ct = section["current_totals"]
    pt = section["previous_totals"]
    dt = section["delta_totals"]

    print(f"\n  Grand Totals:")
    print(f"  {'Metric':<18s} {'Current':>10s} {'Previous':>10s} {'Delta':>10s}")
    print(f"  {'-'*18} {'-'*10} {'-'*10} {'-'*10}")
    print(f"  {'Runs':<18s} {ct['runs']:>10d} {pt['runs']:>10d} {_delta_str(dt['runs']):>10s}")
    print(f"  {'Total Footage':<18s} {ct['total_drill']:>10,.0f} {pt['total_drill']:>10,.0f} {_delta_str(dt['total_drill']):>10s}")
    print(f"  {'Total Hours':<18s} {ct['total_hrs']:>10,.1f} {pt['total_hrs']:>10,.1f} {_delta_str(dt['total_hrs'], fmt='.1f'):>10s}")

    current = section["current"]
    if len(current) > 0:
        print(f"\n  Current Week Breakdown:")
        print(f"  {'JOB_TYPE':<15s} {'MOTOR_TYPE2':<15s} {'Runs':>6s} {'Footage':>10s} {'Hours':>8s}")
        print(f"  {'-'*15} {'-'*15} {'-'*6} {'-'*10} {'-'*8}")
        for _, row in current.iterrows():
            jt = _safe_str(row.get("JOB_TYPE"))
            mt = _safe_str(row.get("MOTOR_TYPE2"))
            print(f"  {jt:<15s} {mt:<15s} {row['runs']:>6d} {row['total_drill']:>10,.0f} {row['total_hrs']:>8,.1f}")


def print_cat1_section_b(section):
    print()
    print("-" * 80)
    print("  1B. CURVES ANALYSIS (Motor_KPI Source)")
    print("-" * 80)

    cur = section["current"]
    prv = section["previous"]
    delta = section["delta"]

    print(f"\n  {'Metric':<25s} {'Current':>10s} {'Previous':>10s} {'Delta':>10s}")
    print(f"  {'-'*25} {'-'*10} {'-'*10} {'-'*10}")
    print(f"  {'Motor_KPI Runs':<25s} {cur['total_motor_kpi']:>10d} {prv['total_motor_kpi']:>10d}")
    print(f"  {'With RUNS PER CUR':<25s} {cur['total_with_rpc']:>10d} {prv['total_with_rpc']:>10d}")
    print(f"  {'1-Run Curves (best)':<25s} {cur['one_run_count']:>10d} {prv['one_run_count']:>10d} {_delta_str(delta['one_run_count']):>10s}")
    print(f"  {'Multi-Run Curves':<25s} {cur['multi_run_count']:>10d} {prv['multi_run_count']:>10d} {_delta_str(delta['multi_run_count']):>10s}")
    print(f"  {'QC Needed':<25s} {cur['qc_needed_count']:>10d} {prv['qc_needed_count']:>10d} {_delta_str(delta['qc_needed_count']):>10s}")

    op_multi = cur.get("operator_multi_run")
    if op_multi is not None and len(op_multi) > 0:
        print(f"\n  Multi-Run Operators (Current Week):")
        for _, row in op_multi.iterrows():
            print(f"    {row['OPERATOR']:<30s}: {int(row['multi_run_count'])} multi-run curves")


def print_cat1_section_c(section):
    print()
    print("-" * 80)
    print(f"  1C. REASON TO POOH ({section['reason_col_used']})")
    print("-" * 80)

    cur = section["current"]
    prv = section["previous"]

    print(f"\n  Filtered runs: {cur['total_filtered']} (current) | {prv['total_filtered']} (previous)")

    if len(cur["breakdown"]) > 0:
        print(f"\n  Category Breakdown (Current Week):")
        print(f"  {'Category':<18s} {'Count':>8s} {'%':>8s}")
        print(f"  {'-'*18} {'-'*8} {'-'*8}")
        for _, row in cur["breakdown"].iterrows():
            print(f"  {str(row['category']):<18s} {row['count']:>8d} {row['pct']:>7.1f}%")

    if len(cur["motor_by_operator"]) > 0:
        print(f"\n  Motor Issues by Operator (Current Week):")
        print(f"  {'Operator':<30s} {'Motor':>8s} {'Total':>8s} {'%':>8s}")
        print(f"  {'-'*30} {'-'*8} {'-'*8} {'-'*8}")
        for _, row in cur["motor_by_operator"].iterrows():
            print(f"  {str(row['OPERATOR'])[:30]:<30s} {int(row['motor_count']):>8d} {int(row['total_count']):>8d} {row['motor_pct']:>7.1f}%")


def print_category1(cat1):
    if cat1 is None:
        return
    print_category_header(1, cat1["category"])
    sections = cat1["sections"]
    print_cat1_section_a(sections["A_weekly_summary"])
    print_cat1_section_b(sections["B_curves"])
    print_cat1_section_c(sections["C_reason_pooh"])


# =========================================================================
# Category 2: Monthly Highlights
# =========================================================================

def print_cat2_section_a(section, cm_label, pm_label):
    print()
    print("-" * 80)
    print("  2A. LONGEST RUNS (Top 5 by Total Footage)")
    print("-" * 80)

    for label, key in [(cm_label, "current_month"), (pm_label, "previous_month")]:
        df = section[key]
        if len(df) == 0:
            print(f"\n  {label}: No data")
            continue
        print(f"\n  {label}:")
        print(f"  {'#':>3s} {'Operator':<25s} {'Well':<30s} {'Footage':>10s} {'ROP':>8s} {'Hrs':>7s} {'Type':<8s} {'Hole':>6s}")
        print(f"  {'-'*3} {'-'*25} {'-'*30} {'-'*10} {'-'*8} {'-'*7} {'-'*8} {'-'*6}")
        for _, run in df.iterrows():
            rop = f"{run['avg_rop']:.1f}" if pd.notna(run.get('avg_rop')) else "N/A"
            hrs = f"{run['drilling_hours']:.1f}" if pd.notna(run.get('drilling_hours')) else "N/A"
            print(f"  {run['rank']:>3d} {str(run['operator'])[:25]:<25s} {str(run['well'])[:30]:<30s} "
                  f"{run['total_drill']:>10,.0f} {rop:>8s} {hrs:>7s} {_safe_str(run['motor_type2']):<8s} {_fmt_hole(run['hole_size']):>6s}")


def print_cat2_section_b(section, cm_label, pm_label):
    print()
    print("-" * 80)
    print("  2B. MONTHLY SUMMARY (JOB_TYPE -> MOTOR_TYPE2)")
    print("-" * 80)

    ct = section["current_totals"]
    pt = section["previous_totals"]
    dt = section["delta_totals"]

    print(f"\n  Grand Totals: {cm_label} vs {pm_label}")
    print(f"  {'Metric':<18s} {'Current':>10s} {'Previous':>10s} {'Delta':>10s}")
    print(f"  {'-'*18} {'-'*10} {'-'*10} {'-'*10}")
    print(f"  {'Runs':<18s} {ct['runs']:>10d} {pt['runs']:>10d} {_delta_str(dt['runs']):>10s}")
    print(f"  {'Total Footage':<18s} {ct['total_drill']:>10,.0f} {pt['total_drill']:>10,.0f} {_delta_str(dt['total_drill']):>10s}")
    print(f"  {'Total Hours':<18s} {ct['total_hrs']:>10,.1f} {pt['total_hrs']:>10,.1f} {_delta_str(dt['total_hrs'], fmt='.1f'):>10s}")


def print_cat2_section_c(section):
    print()
    print("-" * 80)
    print("  2C. FASTEST SECTIONS (Highest AVG_ROP by HOLE_SIZE)")
    print("-" * 80)

    by_hole = section["by_hole_size"]
    if not by_hole:
        print("\n  No data with sufficient runs.")
        return

    for hole_size in sorted(by_hole.keys()):
        info = by_hole[hole_size]
        print(f"\n  Hole Size: {_fmt_hole(hole_size)} | Avg ROP: {info['hole_avg_rop']:.1f} ft/hr | Runs: {info['run_count']}")
        top = info["top_operators"]
        if len(top) > 0:
            for _, row in top.iterrows():
                print(f"    {str(row['operator'])[:25]:<25s} {row['avg_rop']:>7.1f} ft/hr | {_safe_str(row.get('basin')):<15s} {_safe_str(row.get('phase'))}")


def print_cat2_section_d(section):
    print()
    print("-" * 80)
    print("  2D. OPERATOR SUCCESS RATE (TD Runs / Total)")
    print("-" * 80)

    if isinstance(section, pd.DataFrame) and len(section) > 0:
        print(f"\n  {'Operator':<30s} {'TD':>6s} {'Total':>6s} {'Success%':>10s}")
        print(f"  {'-'*30} {'-'*6} {'-'*6} {'-'*10}")
        for _, row in section.iterrows():
            print(f"  {str(row['operator'])[:30]:<30s} {int(row['td_runs']):>6d} {int(row['total_runs']):>6d} {row['success_pct']:>9.1f}%")
    else:
        print("\n  No data available.")


def print_cat2_section_e(section):
    print()
    print("-" * 80)
    print("  2E. MOTOR FAILURES BY OPERATOR")
    print("-" * 80)

    if isinstance(section, pd.DataFrame) and len(section) > 0:
        print(f"\n  {'Operator':<30s} {'Failures':>10s} {'Total':>6s} {'Failure%':>10s}")
        print(f"  {'-'*30} {'-'*10} {'-'*6} {'-'*10}")
        for _, row in section.iterrows():
            print(f"  {str(row['operator'])[:30]:<30s} {int(row['failure_runs']):>10d} {int(row['total_runs']):>6d} {row['failure_pct']:>9.1f}%")
    else:
        print("\n  No motor failures this month.")


def print_cat2_section_f(section):
    print()
    print("-" * 80)
    print("  2F. CURVE SUCCESS RATE (RUNS PER CUR = 1)")
    print("-" * 80)

    data = section.get("data") if isinstance(section, dict) else section
    if isinstance(data, pd.DataFrame) and len(data) > 0:
        total_with = section.get("total_with_data", 0) if isinstance(section, dict) else 0
        total_mkpi = section.get("total_motor_kpi", 0) if isinstance(section, dict) else 0
        print(f"\n  Motor_KPI with data: {total_with} of {total_mkpi}")
        print(f"\n  {'Operator':<30s} {'1-Run':>6s} {'Total':>6s} {'Success%':>10s}")
        print(f"  {'-'*30} {'-'*6} {'-'*6} {'-'*10}")
        for _, row in data.iterrows():
            print(f"  {str(row['operator'])[:30]:<30s} {int(row['one_run']):>6d} {int(row['total_curves']):>6d} {row['success_pct']:>9.1f}%")
    else:
        print("\n  No curve data available.")


def print_category2(cat2):
    if cat2 is None:
        return
    print_category_header(2, cat2["category"])

    cm_label = cat2.get("current_month_label", "Current Month")
    pm_label = cat2.get("previous_month_label", "Previous Month")
    sections = cat2["sections"]

    print_cat2_section_a(sections["A_longest_runs"], cm_label, pm_label)
    print_cat2_section_b(sections["B_monthly_summary"], cm_label, pm_label)
    print_cat2_section_c(sections["C_fastest_sections"])
    print_cat2_section_d(sections["D_operator_success"])
    print_cat2_section_e(sections["E_motor_failures"])
    print_cat2_section_f(sections["F_curve_success"])


# =========================================================================
# Category 3: Historical Analysis (2025+ Baseline)
# =========================================================================

def print_avg_rop(summary):
    if summary is None:
        return
    print()
    print("-" * 80)
    print(f"  3A. AVG ROP ANALYSIS")
    print(f"  Runs with ROP data: {summary['total_runs']}")
    print(f"  Week Average: {summary['week_avg']} ft/hr | Median: {summary['week_median']} ft/hr")
    print("-" * 80)

    results = summary["results"]
    if len(results) == 0:
        print("  No runs with ROP data this week.")
        return

    if "basin_breakdown" in summary:
        print("\n  By Basin:")
        for basin, row in summary["basin_breakdown"].iterrows():
            print(f"    {basin:<20s}: {row['mean']:>7.1f} ft/hr  (n={int(row['count'])})")

    flagged_below = results[results["flag"] == "below"].sort_values("diff_pct")
    if len(flagged_below) > 0:
        print(f"\n  [!] UNDERPERFORMING RUNS ({len(flagged_below)} flagged):")
        for _, run in flagged_below.head(10).iterrows():
            print(f"      {run['operator']:<30s} | {run['well'][:35]:<35s}")
            print(f"        ROP: {run['value']:>7.1f} ft/hr | Baseline avg: {run['baseline_mean']} ft/hr | {run['diff_pct']:+.1f}%")
            print(f"        Hole: {_fmt_hole(run['hole_size'])} | Basin: {_safe_str(run['basin'])} | Phase: {_safe_str(run['phase'])}")
            print(f"        Motor: {_safe_str(run['motor_model'])} | Compared at: {_safe_str(run.get('match_level'))} (n={run['baseline_count']})")
            print()

    flagged_above = results[results["flag"] == "above"].sort_values("diff_pct", ascending=False)
    if len(flagged_above) > 0:
        print(f"\n  [*] TOP PERFORMERS ({len(flagged_above)} highlighted):")
        for _, run in flagged_above.head(5).iterrows():
            print(f"      {run['operator']:<30s} | {run['well'][:35]:<35s}")
            print(f"        ROP: {run['value']:>7.1f} ft/hr | Baseline avg: {run['baseline_mean']} ft/hr | {run['diff_pct']:+.1f}%")
            print(f"        Hole: {_fmt_hole(run['hole_size'])} | Basin: {_safe_str(run['basin'])} | Motor: {_safe_str(run['motor_model'])}")
            print()


def print_longest_runs(summary):
    if summary is None:
        return
    print()
    print("-" * 80)
    print(f"  3B. LONGEST RUNS (Top {summary['top_n']})")
    print(f"  Runs with drill data: {summary['total_runs']}")
    print(f"  Week Avg: {summary['week_avg']:.0f} ft | Max: {summary['week_max']:.0f} ft | Min: {summary['week_min']:.0f} ft")
    print("-" * 80)

    results = summary["results"]
    if len(results) == 0:
        print("  No runs with drilling data this week.")
        return

    print()
    for _, run in results.iterrows():
        rank = int(run["rank"])
        diff_str = f"{run['diff_pct']:+.1f}% vs avg" if run["diff_pct"] is not None else "no baseline"
        hrs_str = f"{run['drilling_hours']} hrs" if run["drilling_hours"] and not pd.isna(run["drilling_hours"]) else "N/A"
        rop_str = f"{run['avg_rop']} ft/hr" if run["avg_rop"] and not pd.isna(run["avg_rop"]) else "N/A"

        print(f"  #{rank}. {run['operator']:<30s} | {run['well'][:35]:<35s}")
        print(f"      Footage: {run['value']:>8,.0f} ft | {diff_str}")
        print(f"      Hours: {hrs_str} | ROP: {rop_str}")
        print(f"      Hole: {_fmt_hole(run['hole_size'])} | Basin: {_safe_str(run['basin'])} | Phase: {_safe_str(run['phase'])}")
        if run["baseline_mean"]:
            ml = _safe_str(run.get('match_level'))
            print(f"      Baseline: {run['baseline_mean']:,.0f} ft (n={run['baseline_count']}) [{ml}]")
        print()


def print_sliding_pct(summary):
    if summary is None:
        return
    print()
    print("-" * 80)
    print(f"  3C. SLIDING % (LAT Phase Only)")
    print(f"  LAT runs with slide data: {summary['total_lat_runs']} of {summary['total_new_runs']} total runs")
    if summary["total_lat_runs"] > 0:
        print(f"  Week Avg: {summary['week_avg']}% | Median: {summary['week_median']}%")
        if summary["hist_threshold_pct"]:
            print(f"  Historical 75th percentile threshold: {summary['hist_threshold_pct']}%")
    print("-" * 80)

    results = summary["results"]
    if len(results) == 0:
        print("  No LAT runs with sliding data this week.")
        return

    sorted_results = results.sort_values("value", ascending=False)

    flagged = sorted_results[sorted_results["flag"] == "above"]
    if len(flagged) > 0:
        print(f"\n  [!] HIGH SLIDING % ({len(flagged)} flagged):")
        for _, run in flagged.iterrows():
            print(f"      {run['operator']:<30s} | {run['well'][:35]:<35s}")
            print(f"        Sliding: {run['value']:>5.1f}% | Slide: {run['slide_drilled']:,.0f} ft of {run['total_drill']:,.0f} ft")
            if run["baseline_mean"]:
                ml = _safe_str(run.get('match_level'))
                print(f"        Baseline avg: {run['baseline_mean']}% | {run['diff_pct']:+.1f}% vs avg [{ml}]")
            print()

    print(f"\n  All LAT Runs:")
    print(f"  {'Operator':<22s} {'Well':<28s} {'Hole':>6s} {'Slide%':>7s} {'Slide ft':>9s} {'Total ft':>9s} {'Basin':<15s}")
    print(f"  {'-'*22} {'-'*28} {'-'*6} {'-'*7} {'-'*9} {'-'*9} {'-'*15}")
    for _, run in sorted_results.iterrows():
        flag_marker = " [!]" if run["flag"] == "above" else "    "
        print(f"  {run['operator'][:22]:<22s} {run['well'][:28]:<28s} {_fmt_hole(run['hole_size']):>6s} "
              f"{run['value']:>6.1f}% {run['slide_drilled']:>8,.0f} {run['total_drill']:>8,.0f} "
              f"{_safe_str(run['basin']):<15s}{flag_marker}")


def print_pattern_highlights(patterns):
    if patterns is None:
        return
    print()
    print("-" * 80)
    print("  3D. PATTERN HIGHLIGHTS")
    print("-" * 80)

    highlights = patterns.get("highlights", [])
    lowlights = patterns.get("lowlights", [])

    if highlights:
        print(f"\n  [*] ABOVE BASELINE ({len(highlights)}):")
        for h in highlights:
            group_desc = " + ".join(f"{k}={v}" for k, v in h["grouping_values"].items())
            ops_str = ", ".join(
                f"{o['operator']} ({o['avg']}, n={o['count']})" for o in h["top_operators"]
            )
            print(f"    {group_desc}")
            print(f"      {h['metric']}: {h['week_avg']} (week) vs {h['baseline_avg']} (baseline) | {h['diff_pct']:+.1f}%")
            print(f"      Top operators: {ops_str}")
            print()

    if lowlights:
        print(f"\n  [!] BELOW BASELINE ({len(lowlights)}):")
        for item in lowlights:
            group_desc = " + ".join(f"{k}={v}" for k, v in item["grouping_values"].items())
            ops_str = ", ".join(
                f"{o['operator']} ({o['avg']}, n={o['count']})" for o in item["top_operators"]
            )
            print(f"    {group_desc}")
            print(f"      {item['metric']}: {item['week_avg']} (week) vs {item['baseline_avg']} (baseline) | {item['diff_pct']:+.1f}%")
            print(f"      Bottom operators: {ops_str}")
            print()

    if not highlights and not lowlights:
        print("\n  No significant deviations found.")


def print_category3(cat3):
    if cat3 is None:
        return
    print_category_header(3, cat3["category"])
    sections = cat3["sections"]

    print_avg_rop(sections.get("A_avg_rop"))
    print_longest_runs(sections.get("B_longest_runs"))
    print_sliding_pct(sections.get("C_sliding_pct"))
    print_pattern_highlights(sections.get("D_pattern_highlights"))


# =========================================================================
# Main entry point
# =========================================================================

def generate_report(all_results):
    """Generate the full console report with all 3 categories."""
    meta = all_results["meta"]
    print_header(meta)

    print_category1(all_results.get("category1"))
    print_category2(all_results.get("category2"))
    print_category3(all_results.get("category3"))

    print_footer()
