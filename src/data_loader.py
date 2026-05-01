"""
Data Loader Module
Reads the master Excel file from SharePoint (or local fallback), cleans data,
and provides filtered views for different time periods.
"""

import pandas as pd
import numpy as np
import glob
import os
import shutil
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


def _resolve_paths(config, scope="all"):
    """Return list of search paths based on scope.
    scope: 'local' (OneDrive personal copy), 'shared' (Teams sync), or 'all' (union)."""
    data_cfg = config.get("data", {})
    if scope == "local":
        paths = data_cfg.get("local_paths") or []
        if not paths:
            paths = data_cfg.get("search_paths", [])
        return paths
    if scope == "shared":
        paths = data_cfg.get("shared_paths") or []
        if not paths:
            paths = data_cfg.get("search_paths", [])
        return paths
    # "all" - union of local and shared (preserves order: local first, then shared)
    paths = list(data_cfg.get("local_paths") or [])
    for p in (data_cfg.get("shared_paths") or []):
        if p not in paths:
            paths.append(p)
    if not paths:
        paths = data_cfg.get("search_paths", [])
    return paths


def find_master_file_sharepoint(config):
    """
    Find and download the latest MASTER_MCS_MERGE file from SharePoint.
    Replicates Power BI's connection:
      SharePoint.Files("https://scoutdownholecom.sharepoint.com/sites/DBRunsUpdate")
    """
    from dotenv import load_dotenv
    load_dotenv(os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), ".env"))

    site_url = os.getenv("SP_SITE_URL", config["sharepoint"]["site_url"])
    username = os.getenv("SP_USERNAME")
    password = os.getenv("SP_PASSWORD")

    if not username or not password or password == "REPLACE_WITH_YOUR_PASSWORD":
        print("  SharePoint credentials not configured in .env file.")
        print("  Falling back to local file search...")
        return None

    try:
        from office365.sharepoint.client_context import ClientContext
        from office365.runtime.auth.user_credential import UserCredential

        print(f"  Connecting to SharePoint: {site_url}")
        ctx = ClientContext(site_url).with_credentials(UserCredential(username, password))

        # List all files in the document library
        doc_lib = ctx.web.default_document_library()
        files = doc_lib.root_folder.get_files(True).execute_query()

        # Filter for MASTER_MCS_MERGE files
        pattern = config["sharepoint"]["file_pattern"]
        exclude = config["sharepoint"].get("exclude_patterns", [])

        candidates = []
        for f in files:
            name = f.name
            if not name.startswith(pattern):
                continue
            if not name.endswith(".xlsx"):
                continue
            # Exclude downloaded/test/backup copies
            skip = False
            for exc in exclude:
                if exc.lower() in name.lower():
                    skip = True
                    break
            if skip:
                continue
            candidates.append({
                "name": name,
                "url": f.serverRelativeUrl,
                "modified": f.time_last_modified,
                "file_obj": f,
            })

        if not candidates:
            print("  No MASTER_MCS_MERGE files found on SharePoint.")
            return None

        # Sort by modified date, newest first
        candidates.sort(key=lambda x: x["modified"], reverse=True)
        selected = candidates[0]

        print(f"  Found {len(candidates)} master files on SharePoint")
        print(f"  Latest: {selected['name']}")
        print(f"  Modified: {selected['modified']}")

        # Download to temp
        project_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        temp_path = os.path.join(project_dir, "master_temp.xlsx")

        print(f"  Downloading to temp...")
        with open(temp_path, "wb") as local_file:
            selected["file_obj"].download(local_file).execute_query()

        print(f"  Download complete: {os.path.basename(temp_path)}")
        return temp_path, selected["name"]

    except Exception as e:
        print(f"  SharePoint connection failed: {e}")
        print("  Falling back to local file search...")
        return None


def find_master_file_local(config, scope="all"):
    """
    Find the most recent MASTER_MCS_MERGE file by searching the configured paths.
    scope: 'local' (OneDrive only), 'shared' (Teams sync only), or 'all' (both).
    """
    pattern = config["data"]["file_pattern"]
    exclude_patterns = config["sharepoint"].get("exclude_patterns", [])
    candidates = []

    for search_path in _resolve_paths(config, scope):
        if config["data"].get("recursive", False):
            # Recursive search
            full_pattern = os.path.join(search_path, "**", pattern)
            matches = glob.glob(full_pattern, recursive=True)
        else:
            full_pattern = os.path.join(search_path, pattern)
            matches = glob.glob(full_pattern)

        for m in matches:
            basename = os.path.basename(m)
            # Exclude downloaded/test/backup files
            skip = False
            for exc in exclude_patterns:
                if exc.lower() in basename.lower():
                    skip = True
                    break
            if not skip:
                candidates.append(m)

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


def find_file_by_name(config, target_filename, scope="all"):
    """
    Search configured paths for a specific filename.
    scope: 'local', 'shared', or 'all'.
    Used by Friday runs to find the QC'd Teams version of the Wednesday file.
    Returns the full path if found, None otherwise.
    """
    exclude_patterns = config["sharepoint"].get("exclude_patterns", [])

    for search_path in _resolve_paths(config, scope):
        if config["data"].get("recursive", False):
            full_pattern = os.path.join(search_path, "**", target_filename)
            matches = glob.glob(full_pattern, recursive=True)
        else:
            full_pattern = os.path.join(search_path, target_filename)
            matches = glob.glob(full_pattern)

        for m in matches:
            basename = os.path.basename(m)
            skip = False
            for exc in exclude_patterns:
                if exc.lower() in basename.lower():
                    skip = True
                    break
            if not skip:
                return m

    return None


def find_master_file_interactive(config, scope="all"):
    """
    Interactive file selection for Wednesday runs.
    Lists available MASTER_MCS_MERGE files in the chosen scope and lets the
    user pick one. scope: 'local', 'shared', or 'all'.
    Returns (file_path, original_filename) tuple.
    """
    pattern = config["data"]["file_pattern"]
    exclude_patterns = config["sharepoint"].get("exclude_patterns", [])
    candidates = []

    for search_path in _resolve_paths(config, scope):
        if config["data"].get("recursive", False):
            full_pattern = os.path.join(search_path, "**", pattern)
            matches = glob.glob(full_pattern, recursive=True)
        else:
            full_pattern = os.path.join(search_path, pattern)
            matches = glob.glob(full_pattern)

        for m in matches:
            basename = os.path.basename(m)
            skip = False
            for exc in exclude_patterns:
                if exc.lower() in basename.lower():
                    skip = True
                    break
            if not skip:
                candidates.append(m)

    if not candidates:
        raise FileNotFoundError(
            f"No files matching '{pattern}' found in search paths: "
            f"{config['data']['search_paths']}"
        )

    # Sort by modification time, most recent first
    candidates.sort(key=os.path.getmtime, reverse=True)

    # Show candidates
    print(f"\n  Found {len(candidates)} master file(s):\n")
    for i, path in enumerate(candidates[:10], 1):
        basename = os.path.basename(path)
        mtime = datetime.fromtimestamp(os.path.getmtime(path)).strftime("%Y-%m-%d %H:%M")
        print(f"    [{i}] {basename}")
        print(f"        Path: {path}")
        print(f"        Modified: {mtime}")
        print()

    # Ask user to pick
    print(f"  Enter number [1-{min(len(candidates), 10)}] to select, or paste a full file path:")
    choice = input("  > ").strip()

    if choice.lower() in ("", "1"):
        selected = candidates[0]
    elif choice.isdigit() and 1 <= int(choice) <= min(len(candidates), 10):
        selected = candidates[int(choice) - 1]
    elif os.path.exists(choice):
        selected = choice
    else:
        print(f"  ERROR: Invalid selection: {choice}")
        raise SystemExit(1)

    original_name = os.path.basename(selected)
    print(f"\n  Selected: {original_name}")
    print(f"  Path: {selected}")
    print(f"  Modified: {datetime.fromtimestamp(os.path.getmtime(selected)).strftime('%Y-%m-%d %H:%M')}")

    return selected, original_name


def _copy_to_temp(local_path):
    """Copy a file to temp to avoid lock issues. Returns read path."""
    project_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    temp_path = os.path.join(project_dir, "master_temp.xlsx")
    try:
        shutil.copy2(local_path, temp_path)
        print(f"  Copied to temp: {os.path.basename(temp_path)}")
        return temp_path
    except PermissionError:
        print(f"  WARNING: Could not copy file (locked?). Reading directly...")
        return local_path


def find_master_file(config):
    """
    Find the master file. Tries SharePoint first, falls back to local.
    Returns (file_path, original_filename) tuple.
    """
    # Try SharePoint first
    result = find_master_file_sharepoint(config)
    if result:
        return result  # (temp_path, original_filename)

    # Fall back to local
    local_path = find_master_file_local(config)
    original_name = os.path.basename(local_path)
    read_path = _copy_to_temp(local_path)
    return read_path, original_name


def load_and_clean(file_path, config):
    """Load the master Excel file and apply data cleaning."""
    print(f"\n  Loading data from Sheet1...")
    df = pd.read_excel(file_path, sheet_name=config["data"]["sheet_name"])
    print(f"  Loaded {len(df)} rows, {len(df.columns)} columns")

    # Strip whitespace from text columns (preserve actual NaN values)
    for col in config["cleaning"]["strip_whitespace"]:
        if col in df.columns:
            mask = df[col].notna()
            df.loc[mask, col] = df.loc[mask, col].astype(str).str.strip()
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

    # Uppercase columns (REASON_POOH, REASON_POOH_QC)
    for col in config["cleaning"].get("uppercase", []):
        if col in df.columns:
            mask = df[col].notna()
            df.loc[mask, col] = df.loc[mask, col].astype(str).str.strip().str.upper()
            df.loc[df[col].isin(["NAN", "NONE", ""]), col] = np.nan

    # --- Variable group cleaning ---

    # Basin name normalization
    if "basin_aliases" in config["cleaning"] and "BASIN" in df.columns:
        for wrong, correct in config["cleaning"]["basin_aliases"].items():
            df.loc[df["BASIN"] == wrong, "BASIN"] = correct

    # HOLE_SIZE: remove outliers
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

    # Also include runs where Week # matches
    if week and week_col in df.columns:
        mask = mask | (df[week_col] == week)

    new_runs = df[mask].copy()

    print(f"  Found {len(new_runs)} new runs")
    if len(new_runs) > 0:
        print(f"  DATE_OUT range: {new_runs[date_col].min()} to {new_runs[date_col].max()}")

    return new_runs, week, date_start, date_end


def filter_previous_week(df, config, current_week_start):
    """
    Filter for the week immediately before the current target week.
    Returns (prev_week_runs, prev_week_start, prev_week_end).
    """
    date_col = config["filtering"]["date_column"]
    prev_week_end = current_week_start - timedelta(days=1)  # Previous Sunday
    prev_week_start = prev_week_end - timedelta(days=6)     # Previous Monday

    mask = (df[date_col] >= prev_week_start) & (df[date_col] <= prev_week_end + timedelta(hours=23, minutes=59))
    prev_runs = df[mask].copy()

    print(f"\n  Previous week: {prev_week_start.date()} to {prev_week_end.date()}")
    print(f"  Previous week runs: {len(prev_runs)}")

    return prev_runs, prev_week_start, prev_week_end


def filter_current_month(df, config, reference_date):
    """
    Filter for runs in the month containing reference_date.
    Returns (month_runs, month_start, month_end, month_label).
    """
    date_col = config["filtering"]["date_column"]
    year = reference_date.year
    month = reference_date.month

    month_start = pd.Timestamp(year, month, 1)
    # Last day of month
    if month == 12:
        month_end = pd.Timestamp(year + 1, 1, 1) - timedelta(days=1)
    else:
        month_end = pd.Timestamp(year, month + 1, 1) - timedelta(days=1)

    mask = (df[date_col] >= month_start) & (df[date_col] <= month_end + timedelta(hours=23, minutes=59))
    month_runs = df[mask].copy()

    month_label = reference_date.strftime("%B %Y")
    print(f"\n  Current month: {month_label} ({month_start.date()} to {month_end.date()})")
    print(f"  Current month runs: {len(month_runs)}")

    return month_runs, month_start, month_end, month_label


def filter_previous_month(df, config, reference_date):
    """
    Filter for runs in the month before the month containing reference_date.
    Returns (month_runs, month_start, month_end, month_label).
    """
    date_col = config["filtering"]["date_column"]
    year = reference_date.year
    month = reference_date.month

    # Previous month
    if month == 1:
        prev_year = year - 1
        prev_month = 12
    else:
        prev_year = year
        prev_month = month - 1

    month_start = pd.Timestamp(prev_year, prev_month, 1)
    month_end = pd.Timestamp(year, month, 1) - timedelta(days=1)

    mask = (df[date_col] >= month_start) & (df[date_col] <= month_end + timedelta(hours=23, minutes=59))
    month_runs = df[mask].copy()

    month_label = month_start.strftime("%B %Y")
    print(f"\n  Previous month: {month_label} ({month_start.date()} to {month_end.date()})")
    print(f"  Previous month runs: {len(month_runs)}")

    return month_runs, month_start, month_end, month_label


def load_comparison_data(df, config, target_week_start):
    """
    Load comparison baseline: 2025+ data, excluding the target week.
    Used for performance benchmarking (Category 3).
    """
    comparison_start = pd.Timestamp(config["filtering"]["comparison_start_date"])
    date_col = config["filtering"]["date_column"]

    mask = (df[date_col] >= comparison_start) & (df[date_col] < target_week_start)
    baseline = df[mask].copy()

    print(f"\n  Comparison baseline: {comparison_start.date()} to {target_week_start.date()}")
    print(f"  Baseline runs: {len(baseline)}")
    return baseline


def get_reason_pooh_col(config, report_type):
    """Return the correct REASON_POOH column name based on report type."""
    if report_type == "friday":
        return config["category1"]["reason_pooh"]["friday_col"]
    return config["category1"]["reason_pooh"]["wednesday_col"]
