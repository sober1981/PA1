"""
Trends — multi-week historical computations and chart rendering for the
PDF report's Category 1D / 1E / 1F sections.

Public entry points:
    compute_trends(df, current_week, num_weeks=12, config=...)
    render_trend_chart(trend_data, metric, chart_type, output_path=None)
"""

import os
import tempfile
import colorsys
from datetime import timedelta

import pandas as pd

# Force non-interactive backend so matplotlib works inside scheduled tasks.
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.colors as mcolors

from src.data_loader import get_week_date_range


COL_WEEK_DEFAULT = "Week #"
COL_DRILL = "TOTAL_DRILL"
COL_HOLE = "HOLE_SIZE"
COL_SOURCE = "SOURCE"

MAX_HOLE_SIZE = 12.25       # Hole sizes greater than this are excluded.
MIN_WEEKS_PRESENT = 2       # Sizes present in fewer weeks than this are excluded.


def _previous_n_weeks(current_week, n):
    """Return chronologically ordered list of n ISO-week strings ending at
    current_week (inclusive). Format: 'YY-Wnn'."""
    cur_start, _ = get_week_date_range(current_week)
    weeks = [current_week]
    cur_dt = cur_start
    for _ in range(n - 1):
        cur_dt = cur_dt - timedelta(days=7)
        iso = cur_dt.isocalendar()
        weeks.append(f"{iso.year % 100:02d}-W{iso.week:02d}")
    weeks.reverse()
    return weeks


def compute_trends(df, current_week, num_weeks=12, config=None,
                   max_hole_size=MAX_HOLE_SIZE, min_weeks_present=MIN_WEEKS_PRESENT):
    """Compute Total Drill and # of Runs per (Week #, HOLE_SIZE) over the last
    `num_weeks` ISO weeks ending at `current_week` (inclusive).

    Filters (set thresholds high/low to disable):
      1. Hole sizes > max_hole_size excluded (default 12.25").
      2. Hole sizes present in fewer than min_weeks_present weeks excluded.
    """
    week_col = (config or {}).get("filtering", {}).get("week_column", COL_WEEK_DEFAULT)
    weeks = _previous_n_weeks(current_week, num_weeks)

    if week_col not in df.columns or COL_HOLE not in df.columns or COL_DRILL not in df.columns:
        return None

    sub = df[df[week_col].isin(weeks) & df[COL_HOLE].notna()].copy()
    if len(sub) == 0:
        return None

    grouped = sub.groupby([week_col, COL_HOLE]).agg(
        drill=(COL_DRILL, "sum"),
        runs=(COL_DRILL, "count"),
    ).reset_index()

    all_hole_sizes = sorted(grouped[COL_HOLE].unique().tolist())
    drill_data = {}
    runs_data = {}
    for hs in all_hole_sizes:
        sub_hs = grouped[grouped[COL_HOLE] == hs].set_index(week_col)
        d_dict = sub_hs["drill"].to_dict()
        r_dict = sub_hs["runs"].to_dict()
        drill_data[hs] = [d_dict.get(wk, None) for wk in weeks]
        runs_data[hs] = [r_dict.get(wk, None) for wk in weeks]

    # Apply filters
    kept = []
    for hs in all_hole_sizes:
        if max_hole_size is not None and hs > max_hole_size:
            continue
        non_null = sum(1 for v in drill_data[hs] if v is not None)
        if min_weeks_present is not None and non_null < min_weeks_present:
            continue
        kept.append(hs)
    drill_data = {hs: drill_data[hs] for hs in kept}
    runs_data = {hs: runs_data[hs] for hs in kept}

    notes = []
    if max_hole_size is not None and max_hole_size < float("inf"):
        notes.append(f'Hole sizes > {max_hole_size:g}" excluded')
    if min_weeks_present is not None and min_weeks_present > 1:
        notes.append(
            f"Hole sizes appearing in fewer than {min_weeks_present} of {num_weeks} weeks excluded"
        )
    if not notes:
        notes.append("All hole sizes included (no filters)")

    return {
        "current_week": current_week,
        "weeks": weeks,
        "hole_sizes": kept,
        "drill_data": drill_data,
        "runs_data": runs_data,
        "filter_notes": notes,
    }


# Base color per integer-group (e.g., "6" group, "7" group, ...).
# Picks from a curated qualitative palette. Distinct between groups, with
# enough headroom to derive shades within a group.
_GROUP_BASE_COLORS = [
    "#1f77b4",  # blue       (6)
    "#ff7f0e",  # orange     (7)
    "#7f7f7f",  # gray       (8)
    "#d62728",  # red        (9)
    "#9467bd",  # purple     (10)
    "#2ca02c",  # green      (11)
    "#8c564b",  # brown      (12)
    "#e377c2",  # pink       (13)
    "#bcbd22",  # olive      (14)
    "#17becf",  # cyan       (15)
]


def _color_palette_grouped(hole_sizes):
    """Map hole sizes to RGB tuples. Sizes sharing an integer part get the
    same color family (different lightness); each integer-group gets a
    distinct base color."""
    groups = {}
    for hs in hole_sizes:
        groups.setdefault(int(hs), []).append(hs)

    sorted_keys = sorted(groups.keys())
    color_for = {}
    for i, key in enumerate(sorted_keys):
        base = _GROUP_BASE_COLORS[i % len(_GROUP_BASE_COLORS)]
        base_rgb = mcolors.to_rgb(base)
        h, l, s = colorsys.rgb_to_hls(*base_rgb)
        sizes_in_group = sorted(groups[key])
        n = len(sizes_in_group)
        if n == 1:
            color_for[sizes_in_group[0]] = base_rgb
        else:
            # Smaller hole = lighter; larger hole within group = darker.
            l_min = max(0.25, l * 0.55)
            l_max = min(0.80, max(l + (1 - l) * 0.45, l_min + 0.15))
            for j, hs in enumerate(sizes_in_group):
                t = (n - 1 - j) / (n - 1)  # smallest size → t=1 (lightest)
                new_l = l_min + (l_max - l_min) * t
                # Preserve original saturation — important for neutrals (gray) so they
                # don't become tinted when shaded.
                color_for[hs] = colorsys.hls_to_rgb(h, new_l, s)
    return color_for


# =====================================================================
# Reason-to-POOH 12-week stacked-bar trend
# =====================================================================

POOH_CATEGORY_ORDER = ["TD", "ROP", "Bit", "Motor", "MWD", "BHA", "Pressure", "Other"]
_POOH_CATEGORY_KEY_MAP = {
    "td": "TD", "rop": "ROP", "bit": "Bit", "motor": "Motor",
    "mwd": "MWD", "bha": "BHA", "pressure": "Pressure",
    "other": "Other", "unknown": "Other",
}
_POOH_CATEGORY_COLORS = {
    "TD":       "#2ca02c",  # green  — desirable outcome
    "ROP":      "#ff7f0e",  # orange
    "Bit":      "#8c564b",  # brown
    "Motor":    "#d62728",  # red
    "MWD":      "#9467bd",  # purple
    "BHA":      "#1f77b4",  # blue
    "Pressure": "#e377c2",  # pink
    "Other":    "#7f7f7f",  # gray
}


def compute_pooh_trend(df, current_week, num_weeks=12, config=None,
                       report_type="wednesday", apply_source_filter=False):
    """Reason-to-POOH counts per (Week #, category) over the last `num_weeks`
    ISO weeks ending at `current_week`. Wednesday uses REASON_POOH; Friday
    uses REASON_POOH_QC.

    apply_source_filter:
      - False (default): include ALL runs regardless of SOURCE column
      - True: filter to SOURCE in config's reason_pooh.source_filter (matches Cat 1C)
    """
    week_col = (config or {}).get("filtering", {}).get("week_column", COL_WEEK_DEFAULT)
    pooh_cfg = (config or {}).get("category1", {}).get("reason_pooh", {})
    if not pooh_cfg:
        return None
    source_filter = pooh_cfg.get("source_filter", []) if apply_source_filter else []
    reason_col = pooh_cfg["friday_col"] if report_type == "friday" else pooh_cfg["wednesday_col"]
    classifications = pooh_cfg.get("classifications", {})

    weeks = _previous_n_weeks(current_week, num_weeks)

    if reason_col not in df.columns or week_col not in df.columns:
        return None

    sub = df[df[week_col].isin(weeks)].copy()
    if apply_source_filter and source_filter and COL_SOURCE in sub.columns:
        sub = sub[sub[COL_SOURCE].isin(source_filter)]
    if len(sub) == 0:
        return None

    # Reuse classifier
    from src.cat1_weekly import _classify_pooh
    sub["_cat_key"] = sub[reason_col].apply(lambda v: _classify_pooh(v, classifications))
    sub["_cat"] = sub["_cat_key"].map(_POOH_CATEGORY_KEY_MAP).fillna("Other")

    grouped = sub.groupby([week_col, "_cat"]).size().reset_index(name="count")

    data = {}
    for cat in POOH_CATEGORY_ORDER:
        cat_sub = grouped[grouped["_cat"] == cat].set_index(week_col)["count"].to_dict()
        data[cat] = [int(cat_sub.get(wk, 0)) for wk in weeks]

    return {
        "current_week": current_week,
        "weeks": weeks,
        "categories": POOH_CATEGORY_ORDER,
        "data": data,
        "reason_col": reason_col,
        "report_type": report_type,
        "source_filter": source_filter,
        "apply_source_filter": apply_source_filter,
    }


def _contrast_text_color(color):
    """Return 'white' or 'black' based on the luminance of the given color
    (so labels remain readable on any background)."""
    rgb = mcolors.to_rgb(color)
    luminance = 0.299 * rgb[0] + 0.587 * rgb[1] + 0.114 * rgb[2]
    return "white" if luminance < 0.55 else "black"


def render_pooh_trend_chart(trend_data, output_path=None, dpi=150,
                            label_min_pct=4.0, compact=False):
    """Render a 100% stacked-bar chart of POOH categories per week with a
    total-runs overlay line on a secondary Y axis.

    label_min_pct: percentage threshold below which segment labels are hidden.
    compact: True = shorter aspect so chart + a follow-up table fit on one page."""
    if not trend_data:
        return None
    weeks = trend_data["weeks"]
    categories = trend_data["categories"]
    data = trend_data["data"]
    current_week = trend_data["current_week"]
    reason_col = trend_data["reason_col"]
    source_filter = trend_data.get("source_filter", [])

    # Per-week totals (used for percentage normalization and the overlay line)
    totals = [sum(data[cat][i] for cat in categories) for i in range(len(weeks))]

    fig, ax = plt.subplots(figsize=(11, 5.0), dpi=dpi)

    # 100% stacked bars (each bar normalized to category share)
    bottom = [0.0] * len(weeks)
    bar_handles = []
    for cat in categories:
        values = data[cat]
        if all(v == 0 for v in values):
            continue  # skip empty categories so they don't crowd the legend
        pct_vals = [
            (v / t * 100.0) if t > 0 else 0.0
            for v, t in zip(values, totals)
        ]
        color = _POOH_CATEGORY_COLORS.get(cat, "#999999")
        bars = ax.bar(weeks, pct_vals, bottom=bottom, label=cat, color=color,
                      edgecolor="white", linewidth=0.5, width=0.7)
        bar_handles.append(bars)
        # Inside-bar percentage labels (skip too-small segments)
        text_color = _contrast_text_color(color)
        for i, pct in enumerate(pct_vals):
            if pct < label_min_pct:
                continue
            y_center = bottom[i] + pct / 2.0
            ax.text(
                i, y_center, f"{int(round(pct))}%",
                ha="center", va="center",
                fontsize=7, fontweight="bold",
                color=text_color,
            )
        bottom = [b + p for b, p in zip(bottom, pct_vals)]

    ax.set_xlabel("Week", fontsize=10)
    ax.set_ylabel("% of Runs", fontsize=10)
    ax.set_ylim(0, 100)
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda y, _p: f"{int(y)}%"))
    ax.set_title(
        f"Reason to POOH ({reason_col}) — Last {len(weeks)} Weeks (current: {current_week})",
        fontsize=11, fontweight="bold",
    )
    ax.grid(True, axis="y", linestyle="--", alpha=0.3)

    plt.xticks(rotation=45, ha="right", fontsize=9)
    ax.tick_params(axis="y", labelsize=9)

    # Secondary Y axis: total runs overlay line
    ax2 = ax.twinx()
    line, = ax2.plot(weeks, totals, color="#C00000", linewidth=2.0,
                     marker="o", markersize=7,
                     markerfacecolor="#C00000", markeredgecolor="white",
                     markeredgewidth=1.0,
                     label="Total Runs")
    ax2.set_ylabel("Total Runs", fontsize=10, color="#C00000")
    ax2.tick_params(axis="y", colors="#C00000", labelsize=9)
    # Headroom above the highest bar/total so labels don't overlap
    ax2.set_ylim(0, max(totals) * 1.20 if totals else 1)

    # Combined legend (categories + total runs line) below the chart
    handles_bars, labels_bars = ax.get_legend_handles_labels()
    ax.legend(
        handles_bars + [line],
        labels_bars + ["Total Runs"],
        title="POOH Category / Overlay",
        loc="upper center",
        bbox_to_anchor=(0.5, -0.20),
        ncol=min(len(handles_bars) + 1, 9),
        fontsize=9,
        title_fontsize=9,
        frameon=False,
    )

    if trend_data.get("apply_source_filter") and source_filter:
        source_label = " + ".join(source_filter)
    else:
        source_label = "all sources (no filter)"
    note = (
        f"Notes: Bars = % of weekly runs | Red line = Total Runs (right axis) | "
        f"Source = {source_label} | Column = {reason_col}"
    )
    fig.text(0.5, 0.005, note,
             ha="center", va="bottom",
             fontsize=7, color="#555555", style="italic")

    plt.tight_layout()
    plt.subplots_adjust(bottom=0.30)

    if output_path is None:
        tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
        output_path = tmp.name
        tmp.close()
    plt.savefig(output_path, dpi=dpi, bbox_inches="tight")
    plt.close(fig)
    return output_path


# =====================================================================
# Hole-size trend metric metadata (used by render_trend_chart below)
# =====================================================================

METRIC_INFO = {
    "drill": {
        "data_key": "drill_data",
        "y_label": "Total Drill (ft)",
        "base_title": "Total Drill by Hole Size",
        "thousands_format": True,
    },
    "runs": {
        "data_key": "runs_data",
        "y_label": "# of Runs",
        "base_title": "Number of Runs by Hole Size",
        "thousands_format": False,
    },
}


def render_trend_chart(trend_data, metric="drill", chart_type="line",
                       output_path=None, dpi=150, compact=False):
    """Render a trend chart to PNG.

    metric: 'drill' or 'runs'.
    chart_type: 'line' (per-hole-size line plot) or 'stacked' (stacked area).
    compact: legacy flag kept for API compatibility — no longer changes
             matplotlib figsize. Page-fit is achieved by sizing the
             pdf.image() embed instead.

    Returns the output PNG path (caller is responsible for cleanup).
    """
    if not trend_data or not trend_data.get("hole_sizes"):
        return None
    if metric not in METRIC_INFO:
        metric = "drill"
    info = METRIC_INFO[metric]
    data = trend_data[info["data_key"]]

    weeks = trend_data["weeks"]
    hole_sizes = trend_data["hole_sizes"]
    current_week = trend_data["current_week"]
    last_idx = len(weeks) - 1
    color_for = _color_palette_grouped(hole_sizes)

    fig, ax = plt.subplots(figsize=(11, 5.0), dpi=dpi)

    if chart_type == "stacked":
        # Stack requires non-None values; treat None as 0.
        matrix = []
        labels = []
        colors = []
        for hs in hole_sizes:
            matrix.append([0.0 if v is None else float(v) for v in data[hs]])
            labels.append(f'{hs:g}"')
            colors.append(color_for[hs])
        ax.stackplot(weeks, matrix, labels=labels, colors=colors, alpha=0.90,
                     edgecolor="white", linewidth=0.5)
        title_suffix = f"({len(weeks)}-Week Stacked Area, current: {current_week})"
    else:
        chart_type = "line"
        for hs in hole_sizes:
            values = data[hs]
            plot_vals = [float(v) if v is not None else float("nan") for v in values]
            ax.plot(
                weeks, plot_vals,
                marker="o", markersize=5,
                color=color_for[hs],
                linewidth=1.5,
                label=f'{hs:g}"',
            )
            cur_val = values[last_idx]
            if cur_val is not None:
                ax.plot(
                    [weeks[last_idx]], [cur_val],
                    marker="o", markersize=12,
                    markerfacecolor=color_for[hs],
                    markeredgecolor="black", markeredgewidth=1.5,
                    linestyle="None",
                )
        title_suffix = f"({len(weeks)}-Week Line Trend, current: {current_week})"

    ax.set_xlabel("Week", fontsize=10)
    ax.set_ylabel(info["y_label"], fontsize=10)
    ax.set_title(f"{info['base_title']} {title_suffix}", fontsize=11, fontweight="bold")
    ax.grid(True, linestyle="--", alpha=0.3)
    if info["thousands_format"]:
        ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _p: f"{int(x):,}"))

    plt.xticks(rotation=45, ha="right", fontsize=9)
    plt.yticks(fontsize=9)

    ncol = min(len(hole_sizes), 8)
    ax.legend(
        title="Hole Size",
        loc="upper center",
        bbox_to_anchor=(0.5, -0.20),
        ncol=ncol,
        fontsize=9,
        title_fontsize=9,
        frameon=False,
    )

    notes = trend_data.get("filter_notes") or []
    if notes:
        note_text = "Notes: " + " | ".join(notes)
        fig.text(
            0.5, 0.005, note_text,
            ha="center", va="bottom",
            fontsize=7, color="#555555", style="italic",
        )

    plt.tight_layout()
    plt.subplots_adjust(bottom=0.30)

    if output_path is None:
        tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
        output_path = tmp.name
        tmp.close()
    plt.savefig(output_path, dpi=dpi, bbox_inches="tight")
    plt.close(fig)
    return output_path
