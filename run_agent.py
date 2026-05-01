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
    load_config, find_master_file_interactive,
    find_master_file_local,
    find_file_by_name, _copy_to_temp, load_and_clean,
    filter_new_runs, filter_previous_week,
    filter_current_month, filter_previous_month,
    load_comparison_data, get_week_date_range,
)
from src.cat1_weekly import run_category1
from src.cat2_monthly import run_category2
from src.kpi_engine import run_category3
from src.qc_audit import run_qc_audit
from src.report import generate_report
from src.pdf_report import generate_pdf
from src.emailer import send_report_email
from src.state import save_wednesday_state, load_wednesday_state
from src.weekly_kpi import compute_weekly_kpi
from src.weekly_kpi_excel import write_kpi_excel


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


def _choose_scope_interactive(default="local"):
    """Ask Local or Shared. Returns 'local' or 'shared'."""
    print("\n  Local or Shared file source?")
    print("    [1] Local  -- your OneDrive copy (pre-QC, won't be touched by team)")
    print("    [2] Shared -- Teams sync (post-QC, the file the team edits)")
    default_label = "1" if default == "local" else "2"
    print(f"    [Enter] = {default} (option {default_label})")
    choice = input("  > ").strip().lower()
    if choice == "":
        return default
    if choice == "1" or choice == "local":
        return "local"
    if choice == "2" or choice == "shared":
        return "shared"
    print(f"  Unrecognized choice '{choice}', defaulting to {default}.")
    return default


def _confirm_file(original, display_path):
    from datetime import datetime as _dt
    print("\n  ============================================================")
    print(f"  Detected file: {original}")
    print(f"  Path:          {display_path}")
    if os.path.exists(display_path):
        mtime = _dt.fromtimestamp(os.path.getmtime(display_path)).strftime("%Y-%m-%d %H:%M")
        size_mb = os.path.getsize(display_path) / 1024 / 1024
        print(f"  Modified:      {mtime}  ({size_mb:.1f} MB)")
    print("  ============================================================")
    ans = input("  Proceed with this file? [Y/n]: ").strip().lower()
    return ans in ("", "y", "yes")


def _choose_week_interactive(df, config):
    """List recent weeks in the data and let the user pick one. Returns a week ID or None."""
    week_col = config["filtering"]["week_column"]
    if week_col not in df.columns:
        return None

    week_counts = (
        df[df[week_col].notna()]
        .groupby(week_col).size().reset_index(name="runs")
        .sort_values(week_col, ascending=False)
    )
    if week_counts.empty:
        return None

    print("\n  Available weeks (most recent first):")
    shown = []
    for i, (_, row) in enumerate(week_counts.head(10).iterrows(), 1):
        wk = row[week_col]
        runs = int(row["runs"])
        try:
            ws_dt, we_dt = get_week_date_range(wk)
            date_str = f"{ws_dt.date()} to {we_dt.date()}"
        except Exception:
            date_str = ""
        shown.append(wk)
        print(f"    [{i}] {wk}  ({date_str})  - {runs} runs")
    print("    [Enter] = latest (option 1)")
    print("    Or type a week ID (e.g. 26-W12)")

    choice = input("  > ").strip()
    if choice == "":
        return shown[0]
    if choice.isdigit():
        n = int(choice)
        if 1 <= n <= len(shown):
            return shown[n - 1]
        for wk in shown:
            if isinstance(wk, str) and wk.endswith(f"-W{n:02d}"):
                return wk
    up = choice.upper().replace(" ", "")
    if up.startswith("W") and up[1:].isdigit():
        n = int(up[1:])
        for wk in shown:
            if isinstance(wk, str) and wk.endswith(f"-W{n:02d}"):
                return wk
    if "-W" in up:
        return choice
    print(f"  Unrecognized choice '{choice}', using latest ({shown[0]}).")
    return shown[0]


def _build_kpi_excel_filename(report_type, week, master_filename):
    """Filename for the Weekly KPI Summary Excel attached to the report."""
    prefix = "Wed" if report_type == "wednesday" else "Fri"
    week_num = week.split("-W")[-1] if "-W" in str(week) else str(week)
    master_base = os.path.splitext(os.path.basename(master_filename))[0]
    return f"PA1 - {prefix} Weekly KPI - Week {week_num} - {master_base}.xlsx"


def main():
    parser = argparse.ArgumentParser(description="PA1 - Weekly Drilling Performance Analysis")
    parser.add_argument("--report", type=str, choices=["wednesday", "friday"], default=None,
                        help="Report type: wednesday (First Look) or friday (Executive Summary + Full)")
    parser.add_argument("--week", type=str, default=None, help="Week to analyze (e.g., 26-W07)")
    parser.add_argument("--date-range", nargs=2, type=str, default=None, metavar=("START", "END"),
                        help="Date range (e.g., 2026-02-09 2026-02-15)")
    parser.add_argument("--file", type=str, default=None, help="Path to master Excel file")
    parser.add_argument("--source", choices=["local", "shared"], default=None,
                        help="Which paths to search: 'local' (OneDrive) or 'shared' (Teams). "
                             "Defaults: wednesday=local, friday=shared. Skips the Local/Shared prompt.")
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
        # Skip-if-already-generated-today guard: only for scheduled (non-TTY) runs.
        if args.report and not sys.stdin.isatty():
            wed_state = load_wednesday_state()
            if wed_state and wed_state.get("saved_at", ""):
                from datetime import date
                saved_date = wed_state["saved_at"][:10]  # "YYYY-MM-DD"
                if saved_date == date.today().isoformat():
                    print("  Wednesday report already generated today. Skipping.")
                    sys.exit(0)

        # Determine scope: --source flag wins, else TTY prompt (default local), else local.
        if args.source:
            scope = args.source
            print(f"  Wednesday report -- source: {scope} (from --source flag)")
        elif sys.stdin.isatty():
            scope = _choose_scope_interactive(default="local")
            if scope == "shared":
                print("\n  WARNING: You picked Shared for a Wednesday run.")
                print("  The snapshot will be the post-QC version, so Friday's Cat 4 audit")
                print("  will show no changes against this snapshot.")
                confirm = input("  Continue anyway? [y/N]: ").strip().lower()
                if confirm not in ("y", "yes"):
                    print("  Aborted by user.")
                    sys.exit(0)
        else:
            scope = "local"
            print(f"  Wednesday report -- source: local (auto)")

        if sys.stdin.isatty():
            print(f"  Wednesday report -- interactive file selection ({scope})")
            local_path, original_filename = find_master_file_interactive(config, scope=scope)
        else:
            print(f"  Wednesday report -- auto-detecting latest {scope} master file")
            local_path = find_master_file_local(config, scope=scope)
            original_filename = os.path.basename(local_path)

        wednesday_filepath = local_path  # Save original path for Friday QC audit
        read_path = _copy_to_temp(local_path)

    elif report_type == "friday":
        # Friday: find the QC'd version of Wednesday's file in the Shared (Teams) folder.
        scope = args.source or "shared"
        wed_state = load_wednesday_state()
        wed_pre_qc_path = None  # For QC audit
        if wed_state:
            target = wed_state["wednesday_filename"]
            wed_pre_qc_path = wed_state.get("wednesday_snapshot") or wed_state.get("wednesday_filepath")
            print(f"  Friday report -- looking for '{target}' in {scope} paths")
            teams_path = find_file_by_name(config, target, scope=scope)
            if teams_path:
                from datetime import datetime as dt
                mtime = dt.fromtimestamp(os.path.getmtime(teams_path)).strftime("%Y-%m-%d %H:%M")
                print(f"  Found: {teams_path}")
                print(f"  Modified: {mtime}")
                read_path = _copy_to_temp(teams_path)
                original_filename = os.path.basename(teams_path)
            else:
                print(f"  WARNING: '{target}' not found in {scope} paths.")
                # Fallback to local — warn + confirm if TTY
                fallback = "local" if scope == "shared" else "shared"
                if sys.stdin.isatty():
                    confirm = input(f"  Fall back to {fallback}? [Y/n]: ").strip().lower()
                    if confirm not in ("", "y", "yes"):
                        print("  Aborted by user.")
                        sys.exit(0)
                else:
                    print(f"  Falling back to {fallback} (non-interactive).")
                fallback_path = find_file_by_name(config, target, scope=fallback)
                if fallback_path:
                    read_path = _copy_to_temp(fallback_path)
                    original_filename = os.path.basename(fallback_path)
                else:
                    print(f"  Not found in {fallback} either. Auto-detecting latest in any scope...")
                    local_path_any = find_master_file_local(config, scope="all")
                    read_path = _copy_to_temp(local_path_any)
                    original_filename = os.path.basename(local_path_any)
        else:
            print("  WARNING: No Wednesday state found. Run Wednesday report first.")
            print(f"  Falling back to auto-detect latest from {scope} paths...")
            local_path_fb = find_master_file_local(config, scope=scope)
            read_path = _copy_to_temp(local_path_fb)
            original_filename = os.path.basename(local_path_fb)

    else:
        # Ad-hoc run: auto-detect across all paths
        local_path_any = find_master_file_local(config, scope="all")
        read_path = _copy_to_temp(local_path_any)
        original_filename = os.path.basename(local_path_any)

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

    # Interactive week picker for manual runs only.
    # Skipped when: --week or --date-range provided, or Friday auto, or non-TTY.
    if (sys.stdin.isatty() and not args.week and not args.date_range
            and args.report != "friday"):
        chosen_week = _choose_week_interactive(df, config)
        if chosen_week:
            args.week = chosen_week

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
    cat1_results = run_category1(new_runs, prev_week, config, report_type, week=week, full_df=df)

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
        attachments = [pdf_path]

        # Weekly KPI Summary Excel (per-hole-size table)
        try:
            print("\n  Computing Weekly KPI Summary...")
            kpi_data = compute_weekly_kpi(new_runs, prev_week, week=week)
            kpi_filename = _build_kpi_excel_filename(report_type, week, original_filename)
            kpi_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), kpi_filename)
            write_kpi_excel(kpi_data, kpi_path)
            print(f"  KPI Excel saved: {kpi_path}")
            attachments.append(kpi_path)
        except Exception as e:
            print(f"  WARNING: Weekly KPI Excel failed: {e}")

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
                f"Attached:\n"
                f"  - Full analysis PDF\n"
                f"  - Weekly KPI Summary (per-hole-size table)\n\n"
                f"-- PA1"
            )
            print("\n  Sending email via Outlook...")
            send_report_email(subject=title, body_text=body, pdf_paths=attachments, recipient=recipient)
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
