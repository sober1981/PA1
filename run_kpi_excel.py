"""
PA1 - Weekly KPI Summary (Excel) — standalone QC tool.

Generates the per-hole-size KPI table as Excel for review before being
integrated into the PA1 Wed/Fri PDF reports.

Usage:
  python run_kpi_excel.py                              # Interactive source picker
  python run_kpi_excel.py --source local               # Latest local file (skip prompt)
  python run_kpi_excel.py --source sharepoint          # Latest SharePoint file
  python run_kpi_excel.py --source browse              # Pick from numbered list of locals
  python run_kpi_excel.py --week 26-W15                # Specific week
  python run_kpi_excel.py --date-range 2026-04-06 2026-04-12
  python run_kpi_excel.py --file path/to/master.xlsx   # Explicit file
  python run_kpi_excel.py --output custom_name.xlsx
"""

import argparse
import sys
import os
import shutil
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.data_loader import (
    load_config,
    find_master_file_local,
    find_master_file_sharepoint,
    find_master_file_interactive,
    _copy_to_temp,
    load_and_clean,
    filter_new_runs,
    filter_previous_week,
    get_week_date_range,
)
from src.weekly_kpi import compute_weekly_kpi
from src.weekly_kpi_excel import write_kpi_excel


def _choose_source_interactive():
    print("\n  Select master file source:")
    print("    [1] Latest LOCAL file (OneDrive sync, fastest)")
    print("    [2] SharePoint - latest update (downloads via auth)")
    print("    [3] Browse local files (pick from list)")
    choice = input("  > ").strip()
    if choice in ("", "1"):
        return "local"
    if choice == "2":
        return "sharepoint"
    if choice == "3":
        return "browse"
    print(f"  Unrecognized choice '{choice}', defaulting to local.")
    return "local"


def _resolve_master_file(source, config):
    """Return (read_path, original_filename, display_path) for the chosen source."""
    if source == "local":
        local_path = find_master_file_local(config)
        return _copy_to_temp(local_path), os.path.basename(local_path), local_path

    if source == "sharepoint":
        result = find_master_file_sharepoint(config)
        if result is None:
            print("  SharePoint failed. Falling back to latest local file.")
            local_path = find_master_file_local(config)
            return _copy_to_temp(local_path), os.path.basename(local_path), local_path
        read_path, original = result
        return read_path, original, f"(SharePoint) {original}"

    if source == "browse":
        local_path, original = find_master_file_interactive(config)
        return _copy_to_temp(local_path), original, local_path

    raise ValueError(f"Unknown source: {source}")


def _choose_week_interactive(df, config):
    """List recent weeks in the data and let the user pick one. Returns a week ID or None for latest."""
    week_col = config["filtering"]["week_column"]
    date_col = config["filtering"]["date_column"]

    if week_col not in df.columns:
        print(f"  WARNING: Week column '{week_col}' not in data; using auto-detect.")
        return None

    week_counts = (
        df[df[week_col].notna()]
        .groupby(week_col)
        .size()
        .reset_index(name="runs")
        .sort_values(week_col, ascending=False)
    )
    if week_counts.empty:
        print("  No weeks found in data; using auto-detect.")
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
    # Numeric: try as list-position first, then as a week number to match in shown
    if choice.isdigit():
        n = int(choice)
        if 1 <= n <= len(shown):
            return shown[n - 1]
        # Could be a week number like "16" — match in shown by suffix
        for wk in shown:
            if isinstance(wk, str) and wk.endswith(f"-W{n:02d}"):
                return wk
        # No match — list other weeks in the data
        print(f"  Week {n:02d} not in the top {len(shown)} listed weeks.")
        return None
    # Strings like "W16", "26-W16", "w16"
    up = choice.upper().replace(" ", "")
    if up.startswith("W") and up[1:].isdigit():
        n = int(up[1:])
        for wk in shown:
            if isinstance(wk, str) and wk.endswith(f"-W{n:02d}"):
                return wk
        return up if False else choice  # unlikely
    if "-W" in up:
        return choice
    print(f"  Unrecognized choice '{choice}', using latest ({shown[0]}).")
    return shown[0]


def _confirm_file(original, display_path):
    print("\n  ============================================================")
    print(f"  Detected file: {original}")
    print(f"  Path:          {display_path}")
    if os.path.exists(display_path):
        mtime = datetime.fromtimestamp(os.path.getmtime(display_path)).strftime("%Y-%m-%d %H:%M")
        size_mb = os.path.getsize(display_path) / 1024 / 1024
        print(f"  Modified:      {mtime}  ({size_mb:.1f} MB)")
    print("  ============================================================")
    ans = input("  Proceed with this file? [Y/n]: ").strip().lower()
    return ans in ("", "y", "yes")


def main():
    parser = argparse.ArgumentParser(description="PA1 - Weekly KPI Summary Excel")
    parser.add_argument("--source", choices=["local", "sharepoint", "browse"], default=None,
                        help="Master file source. Omit for interactive prompt.")
    parser.add_argument("--week", type=str, default=None, help="Week (e.g., 26-W15)")
    parser.add_argument("--date-range", nargs=2, type=str, default=None,
                        metavar=("START", "END"))
    parser.add_argument("--file", type=str, default=None, help="Explicit master Excel file path")
    parser.add_argument("--config", type=str, default=None)
    parser.add_argument("--output", type=str, default=None, help="Output Excel path")
    parser.add_argument("--yes", "-y", action="store_true", help="Skip the confirm prompt after file detection")
    args = parser.parse_args()

    print("[1/5] Loading config...")
    config = load_config(args.config)

    print("[2/5] Locating master file...")
    if args.file:
        if not os.path.exists(args.file):
            print(f"  ERROR: File not found: {args.file}")
            sys.exit(1)
        read_path = _copy_to_temp(args.file)
        original = os.path.basename(args.file)
        display_path = args.file
    else:
        source = args.source or _choose_source_interactive()
        print(f"  Source: {source}")
        read_path, original, display_path = _resolve_master_file(source, config)

    if not args.yes:
        if not _confirm_file(original, display_path):
            print("  Aborted by user.")
            sys.exit(0)

    print(f"  Master: {original}")

    print("[3/5] Loading data...")
    df = load_and_clean(read_path, config)

    print("[4/5] Filtering current + previous week...")
    date_start = args.date_range[0] if args.date_range else None
    date_end = args.date_range[1] if args.date_range else None

    week_arg = args.week
    if not week_arg and not args.date_range and not args.yes:
        week_arg = _choose_week_interactive(df, config)

    cur, week, ws_dt, we_dt = filter_new_runs(
        df, config, week=week_arg, date_start=date_start, date_end=date_end
    )
    print(f"  Current week {week}: {ws_dt.date()} to {we_dt.date()} | {len(cur)} runs")

    if len(cur) == 0:
        print("  No runs found.")
        sys.exit(0)

    prev, ps_dt, pe_dt = filter_previous_week(df, config, ws_dt)
    print(f"  Previous week:        {ps_dt.date()} to {pe_dt.date()} | {len(prev)} runs")

    print("[5/5] Computing KPIs and writing Excel...")
    kpi = compute_weekly_kpi(cur, prev, week=week)

    out = args.output or f"PA1 - Weekly KPI Summary - Week {week}.xlsx"
    try:
        write_kpi_excel(kpi, out)
    except PermissionError as e:
        print(f"\n  ERROR: Cannot write '{out}'.")
        print(f"  Looks like the file is open in Excel. Close it and re-run.")
        print(f"  ({e})")
        sys.exit(1)
    print(f"  Saved: {out}")
    print(f"  Hole sizes: {len(kpi['blocks'])} | Grand total hrs: {kpi['grand_total_hrs']:.2f} | Grand total drill: {kpi['grand_total_drill']:,.0f}")


if __name__ == "__main__":
    main()
