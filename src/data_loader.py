"""
Data Loader Module
Reads the master Excel file, cleans data, and provides filtered views.
"""

import pandas as pd
import numpy as np
import glob
import os
from datetime import datetime, timedelta
import yaml


def load_config(config_path=None):
    """Load configuration from settings.yaml"""
    if config_path is None:
        config_path = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            "config", "settings.yaml"
        )
    with open(config_path, "r") as f:
        return yaml.safe_load(f)


def find_master_file(config):
    """Find the most recent MASTER_MCS_MERGE file across search paths."""
    pattern = config["data"]["file_pattern"]
    candidates = []

    for search_path in config["data"]["search_paths"]:
        full_pattern = os.path.join(search_path, pattern)
        matches = glob.glob(full_pattern)
        candidates.extend(matches)

    if not candidates:
        raise FileNotFoundError(
            f"No files matching '{pattern}' found in search paths: "
            f"{config['data']['search_paths']}"
        )

    # Sort by modification time, most recent first
    candidates.sort(key=os.path.getmtime, reverse=True)
    selected = candidates[0]
    print(f"  Auto-detected master file: {os.path.basename(selected)}")
    print(f"  Path: {selected}")
    print(f"  Modified: {datetime.fromtimestamp(os.path.getmtime(selected)).strftime('%Y-%m-%d %H:%M')}")
    return selected


def load_and_clean(file_path, config):
    """Load the master Excel file and apply data cleaning."""
    print(f"\n  Loading data from Sheet1...")
    df = pd.read_excel(file_path, sheet_name=config["data"]["sheet_name"])
    print(f"  Loaded {len(df)} rows, {len(df.columns)} columns")

    # Strip whitespace from text columns (preserve actual NaN values)
    for col in config["cleaning"]["strip_whitespace"]:
        if col in df.columns:
            # Only strip non-null values to avoid converting NaN to 'nan' string
            mask = df[col].notna()
            df.loc[mask, col] = df.loc[mask, col].astype(str).str.strip()
            # Clean up any 'nan', 'None', or empty strings that slipped through
            df.loc[df[col].isin(["nan", "None", ""]), col] = np.nan

    # Convert numeric columns
    for col in config["cleaning"]["to_numeric"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # Convert datetime columns
    for col in config["cleaning"]["to_datetime"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    # Normalize Phase_CALC (uppercase, strip - preserve NaN)
    if "Phase_CALC" in df.columns:
        mask = df["Phase_CALC"].notna()
        df.loc[mask, "Phase_CALC"] = df.loc[mask, "Phase_CALC"].astype(str).str.strip().str.upper()
        df.loc[df["Phase_CALC"].isin(["NAN", "NONE", ""]), "Phase_CALC"] = np.nan

    # --- Variable group cleaning ---

    # Basin name normalization (fix duplicates like "TX-LA-MS Salt" → "TX-LA-MS Sal")
    if "basin_aliases" in config["cleaning"] and "BASIN" in df.columns:
        for wrong, correct in config["cleaning"]["basin_aliases"].items():
            df.loc[df["BASIN"] == wrong, "BASIN"] = correct

    # HOLE_SIZE: remove outliers (e.g., 222.0 is a data entry error)
    if "HOLE_SIZE" in df.columns and "hole_size_max" in config["cleaning"]:
        max_hole = config["cleaning"]["hole_size_max"]
        outliers = df["HOLE_SIZE"] > max_hole
        if outliers.any():
            print(f"  Cleaned {outliers.sum()} HOLE_SIZE outlier(s) > {max_hole}\"")
            df.loc[outliers, "HOLE_SIZE"] = np.nan

    # LOBE/STAGE: standardize separator to ":"
    if "LOBE/STAGE" in df.columns:
        mask = df["LOBE/STAGE"].notna()
        sep = config["cleaning"].get("lobe_stage_separator", ":")
        df.loc[mask, "LOBE/STAGE"] = (
            df.loc[mask, "LOBE/STAGE"].astype(str)
            .str.strip()
            .str.replace("-", sep, regex=False)
        )
        df.loc[df["LOBE/STAGE"].isin(["nan", "None", ""]), "LOBE/STAGE"] = np.nan

    # FORMATION: normalize uppercase for consistent grouping
    if "FORMATION" in df.columns:
        mask = df["FORMATION"].notna()
        df.loc[mask, "FORMATION"] = df.loc[mask, "FORMATION"].astype(str).str.strip().str.upper()
        df.loc[df["FORMATION"].isin(["NAN", "NONE", ""]), "FORMATION"] = np.nan

    # COUNTY: normalize uppercase
    if "COUNTY" in df.columns:
        mask = df["COUNTY"].notna()
        df.loc[mask, "COUNTY"] = df.loc[mask, "COUNTY"].astype(str).str.strip().str.upper()
        df.loc[df["COUNTY"].isin(["NAN", "NONE", ""]), "COUNTY"] = np.nan

    print(f"  Data cleaning complete")
    return df


def get_week_date_range(week_str):
    """
    Convert a week string like '26-W07' to a Monday-Sunday date range.
    Returns (monday_date, sunday_date).
    """
    parts = week_str.split("-W")
    year = 2000 + int(parts[0])
    week_num = int(parts[1])

    # ISO week: Monday is day 1
    monday = datetime.strptime(f"{year}-W{week_num:02d}-1", "%Y-W%W-%w")
    # Adjust: Python's %W is Monday-based but starts week 0
    # Use ISO calendar instead for accuracy
    from datetime import date
    jan4 = date(year, 1, 4)  # Jan 4 is always in ISO week 1
    start_of_week1 = jan4 - timedelta(days=jan4.isoweekday() - 1)
    monday = start_of_week1 + timedelta(weeks=week_num - 1)
    sunday = monday + timedelta(days=6)

    return pd.Timestamp(monday), pd.Timestamp(sunday)


def filter_new_runs(df, config, week=None, date_start=None, date_end=None):
    """
    Filter for new runs based on DATE_OUT.

    Priority:
    1. If week is provided (e.g., '26-W07'), convert to date range
    2. If date_start/date_end provided, use those directly
    3. If nothing provided, auto-detect the latest week in the data
    """
    date_col = config["filtering"]["date_column"]
    week_col = config["filtering"]["week_column"]

    if week:
        # Convert week string to date range
        date_start, date_end = get_week_date_range(week)
        print(f"\n  Filtering for week {week}: {date_start.date()} to {date_end.date()}")
    elif date_start and date_end:
        date_start = pd.Timestamp(date_start)
        date_end = pd.Timestamp(date_end)
        print(f"\n  Filtering for date range: {date_start.date()} to {date_end.date()}")
    else:
        # Auto-detect: find the latest week in the data
        latest_week = df[week_col].dropna().sort_values().iloc[-1]
        date_start, date_end = get_week_date_range(latest_week)
        week = latest_week
        print(f"\n  Auto-detected latest week: {week} ({date_start.date()} to {date_end.date()})")

    # Primary filter: DATE_OUT within range
    mask = (df[date_col] >= date_start) & (df[date_col] <= date_end + timedelta(hours=23, minutes=59))

    # Also include runs where Week # matches (catches late-reported runs)
    if week and week_col in df.columns:
        mask = mask | (df[week_col] == week)

    new_runs = df[mask].copy()

    print(f"  Found {len(new_runs)} new runs")
    if len(new_runs) > 0:
        print(f"  DATE_OUT range: {new_runs[date_col].min()} to {new_runs[date_col].max()}")

    return new_runs, week, date_start, date_end


def load_comparison_data(df, config, target_week_start):
    """
    Load comparison baseline: 2025+ data, excluding the target week.
    Used for performance benchmarking.
    """
    comparison_start = pd.Timestamp(config["filtering"]["comparison_start_date"])
    date_col = config["filtering"]["date_column"]

    mask = (df[date_col] >= comparison_start) & (df[date_col] < target_week_start)
    baseline = df[mask].copy()

    print(f"\n  Comparison baseline: {comparison_start.date()} to {target_week_start.date()}")
    print(f"  Baseline runs: {len(baseline)}")
    return baseline


def load_full_historical(df):
    """
    Load all available data for cumulative/aggregate queries.
    """
    print(f"\n  Full historical data: {len(df)} runs")
    return df.copy()
