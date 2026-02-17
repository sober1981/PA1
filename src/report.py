"""
Report Generator Module
Formats KPI results into readable console output.
Shows all 5 variable groups: Location, Section, Motor, Bit, Depth.
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


def print_header(week, date_start, date_end, total_runs):
    """Print the report header."""
    print()
    print("=" * 80)
    print(f"  WEEKLY PERFORMANCE REPORT - Week {week}")
    print(f"  Period: {date_start.date()} to {date_end.date()}")
    print("=" * 80)
    print(f"  Total New Runs: {total_runs}")
    print()


def print_avg_rop(summary):
    """Print AVG ROP section."""
    if summary is None:
        return

    print("-" * 80)
    print(f"  AVG ROP ANALYSIS")
    print(f"  Runs with ROP data: {summary['total_runs']}")
    print(f"  Week Average: {summary['week_avg']} ft/hr | Median: {summary['week_median']} ft/hr")
    print("-" * 80)

    results = summary["results"]
    if len(results) == 0:
        print("  No runs with ROP data this week.")
        return

    # Basin breakdown
    if "basin_breakdown" in summary:
        print("\n  By Basin:")
        for basin, row in summary["basin_breakdown"].iterrows():
            print(f"    {basin:<20s}: {row['mean']:>7.1f} ft/hr  (n={int(row['count'])})")

    # Flagged runs - below average
    flagged_below = results[results["flag"] == "below"].sort_values("diff_pct")
    if len(flagged_below) > 0:
        print(f"\n  [!] UNDERPERFORMING RUNS ({len(flagged_below)} flagged):")
        for _, run in flagged_below.head(10).iterrows():
            basin = _safe_str(run['basin'])
            county = _safe_str(run['county'])
            phase = _safe_str(run['phase'])
            motor = _safe_str(run['motor_model'])
            hole = _fmt_hole(run['hole_size'])
            formation = _safe_str(run['formation'])
            lobe_stage = _safe_str(run['lobe_stage'])
            match_lvl = _safe_str(run.get('match_level'), 'N/A')
            print(f"      {run['operator']:<30s} | {run['well'][:35]:<35s}")
            print(f"        ROP: {run['value']:>7.1f} ft/hr | Baseline avg: {run['baseline_mean']} ft/hr | {run['diff_pct']:+.1f}%")
            print(f"        Hole: {hole} | Basin: {basin} | County: {county} | Phase: {phase}")
            print(f"        Motor: {motor} | L/S: {lobe_stage} | Type: {run['motor_type2']} | Formation: {formation}")
            print(f"        Compared at: {match_lvl} (n={run['baseline_count']})")
            print()

    # Flagged runs - above average (highlights)
    flagged_above = results[results["flag"] == "above"].sort_values("diff_pct", ascending=False)
    if len(flagged_above) > 0:
        print(f"\n  [*] TOP PERFORMERS ({len(flagged_above)} highlighted):")
        for _, run in flagged_above.head(5).iterrows():
            basin = _safe_str(run['basin'])
            county = _safe_str(run['county'])
            phase = _safe_str(run['phase'])
            motor = _safe_str(run['motor_model'])
            hole = _fmt_hole(run['hole_size'])
            lobe_stage = _safe_str(run['lobe_stage'])
            match_lvl = _safe_str(run.get('match_level'), 'N/A')
            print(f"      {run['operator']:<30s} | {run['well'][:35]:<35s}")
            print(f"        ROP: {run['value']:>7.1f} ft/hr | Baseline avg: {run['baseline_mean']} ft/hr | {run['diff_pct']:+.1f}%")
            print(f"        Hole: {hole} | Basin: {basin} | County: {county} | Phase: {phase}")
            print(f"        Motor: {motor} | L/S: {lobe_stage} | Compared at: {match_lvl}")
            print()


def print_longest_runs(summary):
    """Print Longest Runs section."""
    if summary is None:
        return

    print("-" * 80)
    print(f"  LONGEST RUNS (Top {summary['top_n']})")
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

        basin = _safe_str(run['basin'])
        county = _safe_str(run['county'])
        phase = _safe_str(run['phase'])
        motor = _safe_str(run['motor_model'])
        hole = _fmt_hole(run['hole_size'])
        formation = _safe_str(run['formation'])
        match_lvl = _safe_str(run.get('match_level'), 'N/A')

        print(f"  #{rank}. {run['operator']:<30s} | {run['well'][:35]:<35s}")
        print(f"      Footage: {run['value']:>8,.0f} ft | {diff_str}")
        print(f"      Hours: {hrs_str} | ROP: {rop_str}")
        print(f"      Hole: {hole} | Basin: {basin} | County: {county} | Phase: {phase}")
        print(f"      Motor: {motor} ({run['motor_type2']}) | Formation: {formation}")
        if run["baseline_mean"]:
            print(f"      Baseline: {run['baseline_mean']:,.0f} ft (n={run['baseline_count']}) [{match_lvl}]")
        print()


def print_sliding_pct(summary):
    """Print Sliding % section."""
    if summary is None:
        return

    print("-" * 80)
    print(f"  SLIDING % (LAT Phase Only)")
    print(f"  LAT runs with slide data: {summary['total_lat_runs']} of {summary['total_new_runs']} total runs")
    if summary["total_lat_runs"] > 0:
        print(f"  Week Avg: {summary['week_avg']}% | Median: {summary['week_median']}%")
        print(f"  Range: {summary['week_min']}% to {summary['week_max']}%")
        if summary["hist_threshold_pct"]:
            print(f"  Historical 75th percentile threshold: {summary['hist_threshold_pct']}%")
    print("-" * 80)

    results = summary["results"]
    if len(results) == 0:
        print("  No LAT runs with sliding data this week.")
        return

    sorted_results = results.sort_values("value", ascending=False) if len(results) > 0 else results

    # Flagged runs (high sliding %)
    flagged = sorted_results[sorted_results["flag"] == "above"]
    if len(flagged) > 0:
        print(f"\n  [!] HIGH SLIDING % ({len(flagged)} flagged):")
        for _, run in flagged.iterrows():
            hole = _fmt_hole(run['hole_size'])
            county = _safe_str(run['county'])
            match_lvl = _safe_str(run.get('match_level'), 'N/A')
            print(f"      {run['operator']:<30s} | {run['well'][:35]:<35s}")
            print(f"        Sliding: {run['value']:>5.1f}% | Slide: {run['slide_drilled']:,.0f} ft of {run['total_drill']:,.0f} ft")
            if run["baseline_mean"]:
                print(f"        Baseline avg: {run['baseline_mean']}% | {run['diff_pct']:+.1f}% vs avg [{match_lvl}]")
            print(f"        Hole: {hole} | Basin: {run['basin']} | County: {county} | Motor: {run['motor_model']}")
            print()

    # Summary table of all LAT runs
    print(f"\n  All LAT Runs:")
    print(f"  {'Operator':<22s} {'Well':<28s} {'Hole':>6s} {'Slide%':>7s} {'Slide ft':>9s} {'Total ft':>9s} {'Basin':<15s} {'County':<12s}")
    print(f"  {'-'*22} {'-'*28} {'-'*6} {'-'*7} {'-'*9} {'-'*9} {'-'*15} {'-'*12}")
    for _, run in sorted_results.iterrows():
        flag_marker = " [!]" if run["flag"] == "above" else "    "
        hole = _fmt_hole(run['hole_size'])
        county = _safe_str(run['county'])
        print(f"  {run['operator'][:22]:<22s} {run['well'][:28]:<28s} {hole:>6s} {run['value']:>6.1f}% {run['slide_drilled']:>8,.0f} {run['total_drill']:>8,.0f} {run['basin']:<15s} {county:<12s}{flag_marker}")


def print_footer():
    """Print report footer."""
    print()
    print("=" * 80)
    print("  End of Report")
    print("=" * 80)
    print()


def generate_report(week, date_start, date_end, new_runs_count, kpi_results):
    """Generate the full console report."""
    print_header(week, date_start, date_end, new_runs_count)

    if "avg_rop" in kpi_results:
        print_avg_rop(kpi_results["avg_rop"])

    if "longest_runs" in kpi_results:
        print_longest_runs(kpi_results["longest_runs"])

    if "sliding_pct" in kpi_results:
        print_sliding_pct(kpi_results["sliding_pct"])

    print_footer()
