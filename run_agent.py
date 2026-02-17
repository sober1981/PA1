"""
Scorecard Performance Agent - CLI Entry Point

Usage:
  python run_agent.py                                      # Auto-detect latest week
  python run_agent.py --week 26-W07                        # Specific week
  python run_agent.py --date-range 2026-02-09 2026-02-15   # Date range
  python run_agent.py --file path/to/master.xlsx           # Custom file
  python run_agent.py --week 26-W06 --pdf                  # Generate PDF report
"""

import argparse
import sys
import os

# Add project root to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.data_loader import load_config, find_master_file, load_and_clean, filter_new_runs, load_comparison_data
from src.kpi_engine import run_all_kpis
from src.report import generate_report
from src.pdf_report import generate_pdf


def main():
    parser = argparse.ArgumentParser(
        description="Scorecard Performance Agent - Weekly Drilling Run Analysis"
    )
    parser.add_argument(
        "--week", type=str, default=None,
        help="Week to analyze (e.g., 26-W07)"
    )
    parser.add_argument(
        "--date-range", nargs=2, type=str, default=None,
        metavar=("START", "END"),
        help="Date range to analyze (e.g., 2026-02-09 2026-02-15)"
    )
    parser.add_argument(
        "--file", type=str, default=None,
        help="Path to master Excel file (overrides auto-detection)"
    )
    parser.add_argument(
        "--config", type=str, default=None,
        help="Path to config file (default: config/settings.yaml)"
    )
    parser.add_argument(
        "--pdf", action="store_true", default=False,
        help="Generate a PDF report in addition to console output"
    )

    args = parser.parse_args()

    # Load configuration
    print("\n[1/5] Loading configuration...")
    config = load_config(args.config)

    # Find or use specified file
    print("\n[2/5] Locating master file...")
    if args.file:
        file_path = args.file
        if not os.path.exists(file_path):
            print(f"  ERROR: File not found: {file_path}")
            sys.exit(1)
        print(f"  Using specified file: {file_path}")
    else:
        file_path = find_master_file(config)

    # Load and clean data
    print("\n[3/5] Loading and cleaning data...")
    df = load_and_clean(file_path, config)

    # Filter new runs
    print("\n[4/5] Filtering new runs...")
    date_start = args.date_range[0] if args.date_range else None
    date_end = args.date_range[1] if args.date_range else None

    new_runs, week, week_start, week_end = filter_new_runs(
        df, config,
        week=args.week,
        date_start=date_start,
        date_end=date_end
    )

    if len(new_runs) == 0:
        print("\n  No new runs found for the specified period.")
        print("  Try a different week or date range.")
        sys.exit(0)

    # Load comparison baseline (2025+ data)
    baseline_df = load_comparison_data(df, config, week_start)

    # Run KPIs
    print("\n[5/5] Running KPI analysis...")
    kpi_results = run_all_kpis(new_runs, baseline_df, config)

    # Generate console report
    generate_report(week, week_start, week_end, len(new_runs), kpi_results)

    # Generate PDF if requested
    if args.pdf:
        print("\nGenerating PDF report...")
        pdf_path = generate_pdf(week, week_start, week_end, len(new_runs), kpi_results)
        print(f"  PDF saved to: {pdf_path}")


if __name__ == "__main__":
    main()
