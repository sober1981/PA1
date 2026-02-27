"""
PA1 - Performance Agent CLI

Usage:
  python run_agent.py --report wednesday                 # Wednesday "First Look" report
  python run_agent.py --report friday                    # Friday executive summary + full report
  python run_agent.py --week 26-W07                      # Specific week (ad-hoc)
  python run_agent.py --date-range 2026-02-09 2026-02-15 # Date range (ad-hoc)
  python run_agent.py --file path/to/master.xlsx         # Custom file
"""

import argparse
import sys
import os
import platform
import traceback

# Add project root to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.data_loader import (
    load_config, find_master_file, find_master_file_interactive,
    find_file_by_name, _copy_to_temp, load_and_clean,
    filter_new_runs, filter_previous_week,
    filter_current_month, filter_previous_month,
    load_comparison_data,
)
from src.cat1_weekly import run_category1
from src.cat2_monthly import run_category2
from src.kpi_engine import run_category3
from src.qc_audit import run_qc_audit
from src.report import generate_report
from src.pdf_report import generate_pdf
from src.emailer import send_report_email
from src.state import save_wednesday_state, load_wednesday_state


def _build_report_title(report_type, week_start, week_end, week, master_filename):
    """
    Build report title in the format:
      PA1 - Wed Report Feb 9 to 15 / Week 07 / MASTER_MCS_MERGE_20260218_092739_Feb 18.xlsx
    """
    prefix = "Wed" if report_type == "wednesday" else "Fri"

    # Format dates without leading zeros
    if platform.system() == "Windows":
        mon_str = week_start.strftime("%b %#d")
        sun_str = week_end.strftime("%#d")
    else:
        mon_str = week_start.strftime("%b %-d")
        sun_str = week_end.strftime("%-d")

    week_num = week.split("-W")[-1] if "-W" in str(week) else str(week)
    basename = os.path.basename(master_filename)
    return f"PA1 - {prefix} Report {mon_str} to {sun_str} / Week {week_num} / {basename}"


def main():
    parser = argparse.ArgumentParser(description="PA1 - Weekly Drilling Performance Analysis")
    parser.add_argument("--report", type=str, choices=["wednesday", "friday"], default=None,
                        help="Report type: wednesday (First Look) or friday (Executive Summary + Full)")
    parser.add_argument("--week", type=str, default=None, help="Week to analyze (e.g., 26-W07)")
    parser.add_argument("--date-range", nargs=2, type=str, default=None, metavar=("START", "END"),
                        help="Date range (e.g., 2026-02-09 2026-02-15)")
    parser.add_argument("--file", type=str, default=None, help="Path to master Excel file")
    parser.add_argument("--config", type=str, default=None, help="Path to config file")
    parser.add_argument("--pdf", action="store_true", default=False, help="Generate PDF report")
    parser.add_argument("--no-email", action="store_true", default=False, help="Skip sending email")

    args = parser.parse_args()
    if args.report:
        args.pdf = True

    report_type = args.report or "wednesday"

    # =========================================================================
    # [1/8] Load configuration
    # =========================================================================
    print("\n[1/8] Loading configuration...")
    config = load_config(args.config)

    # =========================================================================
    # [2/8] Locate master file
    # =========================================================================
    print("\n[2/8] Locating master file...")
    wednesday_filepath = None  # For QC audit state tracking
    wed_pre_qc_path = None     # Wednesday pre-QC file for Friday QC audit
    if args.file:
        # Explicit file path provided
        if not os.path.exists(args.file):
            print(f"  ERROR: File not found: {args.file}")
            sys.exit(1)
        print(f"  Using specified file: {args.file}")
        wednesday_filepath = os.path.abspath(args.file)
        read_path = _copy_to_temp(args.file)
        original_filename = os.path.basename(args.file)

    elif report_type == "wednesday":
        # Wednesday: interactive file selection
        print("  Wednesday report -- interactive file selection")
        local_path, original_filename = find_master_file_interactive(config)
        wednesday_filepath = local_path  # Save original path for Friday QC audit
        read_path = _copy_to_temp(local_path)

    elif report_type == "friday":
        # Friday: auto-find the QC'd version of Wednesday's file in Teams sync
        wed_state = load_wednesday_state()
        wed_pre_qc_path = None  # For QC audit
        if wed_state:
            target = wed_state["wednesday_filename"]
            wed_pre_qc_path = wed_state.get("wednesday_filepath")
            print(f"  Friday report -- looking for QC'd version of: {target}")
            teams_path = find_file_by_name(config, target)
            if teams_path:
                from datetime import datetime as dt
                mtime = dt.fromtimestamp(os.path.getmtime(teams_path)).strftime("%Y-%m-%d %H:%M")
                print(f"  Found in Teams sync: {teams_path}")
                print(f"  Modified: {mtime} (QC'd version)")
                read_path = _copy_to_temp(teams_path)
                original_filename = os.path.basename(teams_path)
            else:
                print(f"  WARNING: Could not find '{target}' in Teams sync.")
                print(f"  Falling back to auto-detect...")
                read_path, original_filename = find_master_file(config)
        else:
            print("  WARNING: No Wednesday state found. Run Wednesday report first.")
            print("  Falling back to auto-detect...")
            read_path, original_filename = find_master_file(config)

    else:
        # Ad-hoc run: auto-detect
        read_path, original_filename = find_master_file(config)

    # =========================================================================
    # [3/8] Load and clean data
    # =========================================================================
    print("\n[3/8] Loading and cleaning data...")
    df = load_and_clean(read_path, config)

    # =========================================================================
    # [4/8] Filter current week runs
    # =========================================================================
    print("\n[4/8] Filtering runs...")
    date_start = args.date_range[0] if args.date_range else None
    date_end = args.date_range[1] if args.date_range else None

    new_runs, week, week_start, week_end = filter_new_runs(
        df, config, week=args.week, date_start=date_start, date_end=date_end
    )

    if len(new_runs) == 0:
        print("\n  No new runs found for the specified period.")
        sys.exit(0)

    # Save state for Wednesday runs (so Friday can find the same file)
    if report_type == "wednesday" and args.report:
        save_wednesday_state(original_filename, week, week_start, week_end, filepath=wednesday_filepath)

    # =========================================================================
    # [5/8] Load time-period data slices
    # =========================================================================
    print("\n[5/8] Loading comparison data...")

    # Previous week (Category 1)
    prev_week, prev_week_start, prev_week_end = filter_previous_week(df, config, week_start)

    # Current month (Category 2)
    current_month, cm_start, cm_end, cm_label = filter_current_month(df, config, week_end)

    # Previous month (Category 2)
    prev_month, pm_start, pm_end, pm_label = filter_previous_month(df, config, week_end)

    # Comparison baseline 2025+ (Category 3)
    baseline_df = load_comparison_data(df, config, week_start)

    # =========================================================================
    # [6/8] Run all 3 categories
    # =========================================================================
    print("\n[6/8] Running analysis...")

    print("  Category 1: Week vs Previous Week...")
    cat1_results = run_category1(new_runs, prev_week, config, report_type)

    print("  Category 2: Monthly Highlights...")
    cat2_results = run_category2(current_month, prev_month, config, report_type)

    print("  Category 3: Historical Analysis...")
    cat3_results = run_category3(new_runs, baseline_df, config)

    # Category 4: QC Audit (Friday only — compare Wed pre-QC vs Fri post-QC)
    cat4_results = None
    if report_type == "friday" and wed_pre_qc_path and os.path.exists(wed_pre_qc_path):
        print("  Category 4: QC Audit (comparing Wed vs Fri)...")
        try:
            wed_raw_path = _copy_to_temp(wed_pre_qc_path)
            import pandas as pd
            wed_raw_df = pd.read_excel(wed_raw_path, sheet_name=config.get("data", {}).get("sheet_name", "Sheet1"))
            # Friday df is already loaded as 'df' (cleaned)
            # Load Wednesday raw too and apply same basic cleaning for comparable dtypes
            wed_clean = load_and_clean(wed_raw_path, config)
            cat4_results = run_qc_audit(wed_clean, df, config, week_start, week_end)
            if cat4_results:
                meta = cat4_results["meta"]
                print(f"    Matched rows: {meta['matched']} | Changes: {meta['total_changes']} | New: {meta['new_in_fri']} | Removed: {meta['removed_from_wed']}")
        except Exception as e:
            print(f"  WARNING: QC Audit failed: {e}")
            cat4_results = None
    elif report_type == "friday":
        print("  Category 4: QC Audit skipped (no Wednesday file path in state)")

    all_results = {
        "category1": cat1_results,
        "category2": cat2_results,
        "category3": cat3_results,
        "category4": cat4_results,
        "meta": {
            "week": week,
            "week_start": week_start,
            "week_end": week_end,
            "prev_week_start": prev_week_start,
            "prev_week_end": prev_week_end,
            "current_month_label": cm_label,
            "previous_month_label": pm_label,
            "report_type": report_type,
            "total_new_runs": len(new_runs),
            "master_filename": original_filename,
        }
    }

    # =========================================================================
    # [7/8] Generate console report
    # =========================================================================
    print("\n[7/8] Generating console report...")
    generate_report(all_results)

    # =========================================================================
    # [8/8] Generate PDF + email
    # =========================================================================
    print("\n[8/8] Generating output...")

    if args.pdf or args.report:
        title = _build_report_title(report_type, week_start, week_end, week, original_filename)
        print(f"\n  Report title: {title}")

        pdf_path = generate_pdf(all_results, report_title=title)
        print(f"  PDF saved: {pdf_path}")
        pdf_paths = [pdf_path]

        if args.report == "friday":
            print("  Executive summary: [coming soon]")

        # Send email
        if args.report and not args.no_email:
            recipient = config.get("email", {}).get("recipient", "jsoberanes@scoutdownhole.com")
            body = (
                f"PA1 - {title}\n\n"
                f"Total new runs: {len(new_runs)}\n"
                f"Week: {week} ({week_start.date()} to {week_end.date()})\n"
                f"Master file: {original_filename}\n\n"
                f"See attached PDF(s) for full analysis.\n\n"
                f"-- PA1"
            )
            print("\n  Sending email via Outlook...")
            send_report_email(subject=title, body_text=body, pdf_paths=pdf_paths, recipient=recipient)
        elif args.no_email:
            print("\n  Email skipped (--no-email flag)")
    else:
        print("  Console report only. Use --pdf or --report to generate PDF.")

    # Cleanup temp file
    project_dir = os.path.dirname(os.path.abspath(__file__))
    temp_path = os.path.join(project_dir, "master_temp.xlsx")
    if os.path.exists(temp_path):
        try:
            os.remove(temp_path)
        except OSError:
            pass


def _send_failure_email(error_msg):
    """Send a failure notification email when the report cannot be generated."""
    try:
        from src.emailer import send_report_email
        from datetime import datetime
        now = datetime.now().strftime("%Y-%m-%d %H:%M")
        subject = f"PA1 - REPORT FAILED ({now})"
        body = (
            f"PA1 failed to generate the scheduled report.\n\n"
            f"Time: {now}\n"
            f"Error:\n{error_msg}\n\n"
            f"Action needed: Check the issue and re-run manually if necessary.\n"
            f"  python run_agent.py --report friday --no-email\n\n"
            f"-- PA1"
        )
        send_report_email(
            subject=subject,
            body_text=body,
            pdf_paths=[],
            recipient="jsoberanes@scoutdownhole.com",
        )
    except Exception as email_err:
        print(f"  CRITICAL: Could not send failure email: {email_err}")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        error_msg = traceback.format_exc()
        print(f"\n  FATAL ERROR:\n{error_msg}")
        # Only send failure email for scheduled reports (--report flag)
        if "--report" in sys.argv:
            _send_failure_email(error_msg)
        sys.exit(1)
