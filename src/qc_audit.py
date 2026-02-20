"""
Category 4: QC Audit
Compares Wednesday (pre-QC) and Friday (post-QC) files to identify
what changed during the QC process.
Only runs in Friday reports.
"""

import pandas as pd
import numpy as np


def _build_key(df, key_cols):
    """Build a composite string key from multiple columns."""
    parts = []
    for col in key_cols:
        if col in df.columns:
            parts.append(df[col].astype(str).str.strip())
        else:
            parts.append(pd.Series([""] * len(df), index=df.index))
    return parts[0].str.cat(parts[1:], sep="|")


def _values_differ(val_wed, val_fri, float_tol=0.01):
    """Check if two values are meaningfully different."""
    # Both NaN/None → no change
    wed_na = pd.isna(val_wed)
    fri_na = pd.isna(val_fri)
    if wed_na and fri_na:
        return False
    # One NaN, other not → change
    if wed_na != fri_na:
        return True
    # Both numeric → compare with tolerance
    try:
        w = float(val_wed)
        f = float(val_fri)
        return abs(w - f) > float_tol
    except (ValueError, TypeError):
        pass
    # String comparison (strip + case-insensitive for text)
    return str(val_wed).strip() != str(val_fri).strip()


def run_qc_audit(wed_df, fri_df, config, week_start, week_end):
    """
    Main QC audit: compare Wednesday and Friday DataFrames.
    Returns dict with sections 4A-4D results.
    """
    cfg = config.get("category4", {})
    if not cfg.get("enabled", True):
        return None

    key_cols = cfg.get("match_keys", ["OPERATOR", "WELL", "SN", "DEPTH_IN"])
    exclude = set(cfg.get("exclude_from_diff", []))
    qc_by_col = cfg.get("qc_by_col", "QC BY")
    float_tol = cfg.get("float_tolerance", 0.01)

    # Filter to current week only
    for df in [wed_df, fri_df]:
        if "DATE_OUT" in df.columns:
            df["DATE_OUT"] = pd.to_datetime(df["DATE_OUT"], errors="coerce")

    week_start_ts = pd.Timestamp(week_start)
    week_end_ts = pd.Timestamp(week_end) + pd.Timedelta(days=1)  # inclusive

    wed_week = wed_df[
        (wed_df["DATE_OUT"] >= week_start_ts) & (wed_df["DATE_OUT"] < week_end_ts)
    ].copy()
    fri_week = fri_df[
        (fri_df["DATE_OUT"] >= week_start_ts) & (fri_df["DATE_OUT"] < week_end_ts)
    ].copy()

    if len(wed_week) == 0 or len(fri_week) == 0:
        return _empty_results(len(wed_week), len(fri_week))

    # Build composite keys
    wed_week["_key"] = _build_key(wed_week, key_cols)
    fri_week["_key"] = _build_key(fri_week, key_cols)

    # Find matched, new, and removed rows
    wed_keys = set(wed_week["_key"])
    fri_keys = set(fri_week["_key"])
    matched_keys = wed_keys & fri_keys
    new_in_fri = fri_keys - wed_keys
    removed_from_wed = wed_keys - fri_keys

    # Columns to compare (present in both, not excluded)
    all_cols = [c for c in fri_week.columns if c in wed_week.columns and c not in exclude and c != "_key"]

    # Cell-by-cell comparison for matched rows
    wed_indexed = wed_week.set_index("_key")
    fri_indexed = fri_week.set_index("_key")

    # Handle duplicate keys — keep first occurrence
    wed_indexed = wed_indexed[~wed_indexed.index.duplicated(keep="first")]
    fri_indexed = fri_indexed[~fri_indexed.index.duplicated(keep="first")]

    changes = []  # list of {key, column, wed_value, fri_value, operator, qc_by}
    row_change_counts = {}  # key → number of changed cells

    for key in matched_keys:
        wed_row = wed_indexed.loc[key]
        fri_row = fri_indexed.loc[key]
        operator = fri_row.get("OPERATOR", "Unknown")
        qc_by = fri_row.get(qc_by_col, "Unknown")
        row_changes = 0

        for col in all_cols:
            if _values_differ(wed_row.get(col), fri_row.get(col), float_tol):
                changes.append({
                    "key": key,
                    "column": col,
                    "wed_value": wed_row.get(col),
                    "fri_value": fri_row.get(col),
                    "operator": operator,
                    "qc_by": str(qc_by).strip() if pd.notna(qc_by) else "Unknown",
                })
                row_changes += 1

        row_change_counts[key] = {
            "changes": row_changes,
            "operator": operator,
            "qc_by": str(qc_by).strip() if pd.notna(qc_by) else "Unknown",
        }

    changes_df = pd.DataFrame(changes) if changes else pd.DataFrame(
        columns=["key", "column", "wed_value", "fri_value", "operator", "qc_by"]
    )

    total_matched = len(matched_keys)

    # === Section 4A: Column Change Summary ===
    section_a = _section_a_column_summary(changes_df, total_matched)

    # === Section 4B: QC Reviewer Workload ===
    section_b = _section_b_reviewer_workload(changes_df, row_change_counts, total_matched)

    # === Section 4C: Operator QC Trends ===
    section_c = _section_c_operator_trends(changes_df, row_change_counts, total_matched)

    # === Section 4D: Auto-Detected Patterns ===
    section_d = _section_d_patterns(changes_df, row_change_counts, total_matched, new_in_fri, removed_from_wed)

    return {
        "meta": {
            "wed_rows": len(wed_week),
            "fri_rows": len(fri_week),
            "matched": total_matched,
            "new_in_fri": len(new_in_fri),
            "removed_from_wed": len(removed_from_wed),
            "total_changes": len(changes_df),
            "columns_compared": len(all_cols),
        },
        "sections": {
            "A_column_summary": section_a,
            "B_reviewer_workload": section_b,
            "C_operator_trends": section_c,
            "D_patterns": section_d,
        },
    }


def _empty_results(wed_count, fri_count):
    """Return empty result structure when no data to compare."""
    return {
        "meta": {
            "wed_rows": wed_count,
            "fri_rows": fri_count,
            "matched": 0,
            "new_in_fri": 0,
            "removed_from_wed": 0,
            "total_changes": 0,
            "columns_compared": 0,
        },
        "sections": {
            "A_column_summary": pd.DataFrame(columns=["column", "changes", "pct"]),
            "B_reviewer_workload": pd.DataFrame(columns=["reviewer", "rows_assigned", "rows_changed", "cell_changes", "avg_per_row"]),
            "C_operator_trends": pd.DataFrame(columns=["operator", "rows", "changed_rows", "cell_changes", "top_columns"]),
            "D_patterns": {"systematic": [], "broken_columns": [], "high_effort_rows": [], "new_rows": 0, "removed_rows": 0},
        },
    }


def _section_a_column_summary(changes_df, total_matched):
    """4A: Which columns changed most during QC."""
    if len(changes_df) == 0:
        return pd.DataFrame(columns=["column", "changes", "pct"])

    col_counts = changes_df.groupby("column").size().reset_index(name="changes")
    col_counts["pct"] = (col_counts["changes"] / total_matched * 100).round(1)
    col_counts = col_counts.sort_values("changes", ascending=False).reset_index(drop=True)
    return col_counts


def _section_b_reviewer_workload(changes_df, row_change_counts, total_matched):
    """4B: QC reviewer workload comparison."""
    if not row_change_counts:
        return pd.DataFrame(columns=["reviewer", "rows_assigned", "rows_changed", "cell_changes", "avg_per_row"])

    # Build per-reviewer stats from row_change_counts
    reviewer_stats = {}
    for key, info in row_change_counts.items():
        reviewer = info["qc_by"]
        if reviewer not in reviewer_stats:
            reviewer_stats[reviewer] = {"rows_assigned": 0, "rows_changed": 0, "cell_changes": 0}
        reviewer_stats[reviewer]["rows_assigned"] += 1
        if info["changes"] > 0:
            reviewer_stats[reviewer]["rows_changed"] += 1
        reviewer_stats[reviewer]["cell_changes"] += info["changes"]

    rows = []
    for reviewer, stats in reviewer_stats.items():
        avg = round(stats["cell_changes"] / stats["rows_assigned"], 1) if stats["rows_assigned"] > 0 else 0
        rows.append({
            "reviewer": reviewer,
            "rows_assigned": stats["rows_assigned"],
            "rows_changed": stats["rows_changed"],
            "cell_changes": stats["cell_changes"],
            "avg_per_row": avg,
        })

    result = pd.DataFrame(rows)
    return result.sort_values("cell_changes", ascending=False).reset_index(drop=True)


def _section_c_operator_trends(changes_df, row_change_counts, total_matched):
    """4C: Per-operator QC correction trends."""
    if not row_change_counts:
        return pd.DataFrame(columns=["operator", "rows", "changed_rows", "cell_changes", "top_columns"])

    # Operator-level aggregation
    op_stats = {}
    for key, info in row_change_counts.items():
        op = info["operator"]
        if op not in op_stats:
            op_stats[op] = {"rows": 0, "changed_rows": 0, "cell_changes": 0}
        op_stats[op]["rows"] += 1
        if info["changes"] > 0:
            op_stats[op]["changed_rows"] += 1
        op_stats[op]["cell_changes"] += info["changes"]

    # Top changed columns per operator
    op_top_cols = {}
    if len(changes_df) > 0:
        for op, grp in changes_df.groupby("operator"):
            top = grp["column"].value_counts().head(3).index.tolist()
            op_top_cols[op] = ", ".join(top)

    rows = []
    for op, stats in op_stats.items():
        rows.append({
            "operator": op,
            "rows": stats["rows"],
            "changed_rows": stats["changed_rows"],
            "cell_changes": stats["cell_changes"],
            "top_columns": op_top_cols.get(op, "—"),
        })

    result = pd.DataFrame(rows)
    return result.sort_values("cell_changes", ascending=False).reset_index(drop=True)


def _section_d_patterns(changes_df, row_change_counts, total_matched, new_in_fri, removed_from_wed):
    """4D: Auto-detected patterns."""
    systematic = []  # operator+column combos that repeat
    broken_columns = []  # columns where >50% of rows changed
    high_effort_rows = []  # rows with the most changes

    if len(changes_df) > 0 and total_matched > 0:
        # Systematic: operator+column combos with 3+ occurrences
        op_col = changes_df.groupby(["operator", "column"]).size().reset_index(name="count")
        op_col = op_col[op_col["count"] >= 3].sort_values("count", ascending=False)
        for _, row in op_col.head(10).iterrows():
            systematic.append({
                "operator": row["operator"],
                "column": row["column"],
                "count": int(row["count"]),
            })

        # Broken columns: >50% of matched rows changed
        col_counts = changes_df.groupby("column").agg(
            rows_changed=("key", "nunique")
        ).reset_index()
        col_counts["pct"] = (col_counts["rows_changed"] / total_matched * 100).round(1)
        broken = col_counts[col_counts["pct"] >= 50].sort_values("pct", ascending=False)
        for _, row in broken.iterrows():
            broken_columns.append({
                "column": row["column"],
                "rows_changed": int(row["rows_changed"]),
                "pct": row["pct"],
            })

    # High-effort rows: top 5 with most changes
    if row_change_counts:
        sorted_rows = sorted(row_change_counts.items(), key=lambda x: x[1]["changes"], reverse=True)
        for key, info in sorted_rows[:5]:
            if info["changes"] > 0:
                parts = key.split("|")
                high_effort_rows.append({
                    "operator": info["operator"],
                    "well": parts[1] if len(parts) > 1 else "Unknown",
                    "sn": parts[2] if len(parts) > 2 else "Unknown",
                    "changes": info["changes"],
                    "qc_by": info["qc_by"],
                })

    return {
        "systematic": systematic,
        "broken_columns": broken_columns,
        "high_effort_rows": high_effort_rows,
        "new_rows": len(new_in_fri),
        "removed_rows": len(removed_from_wed),
    }
