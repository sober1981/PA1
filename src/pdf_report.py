"""
PDF Report Generator Module
Generates a professional PDF report organized into 3 categories:
  Category 1: Week vs Previous Week
  Category 2: Monthly Highlights
  Category 3: Historical Analysis (2025+ Baseline)
"""

import os
import platform
from datetime import datetime
from fpdf import FPDF
import pandas as pd
import numpy as np


class PerformanceReport(FPDF):
    """Custom PDF class with header/footer."""

    def __init__(self, meta):
        super().__init__(orientation="L", format="letter")
        self.meta = meta
        self.set_auto_page_break(auto=True, margin=20)

    def header(self):
        self.set_font("Helvetica", "B", 10)
        self.set_text_color(100, 100, 100)
        week = self.meta.get("week", "")
        self.cell(0, 6, f"Scout Downhole - PA1 Performance Report | Week {week}", align="L")
        self.ln(4)
        self.set_draw_color(200, 60, 30)
        self.set_line_width(0.8)
        self.line(10, self.get_y(), self.w - 10, self.get_y())
        self.ln(6)

    def footer(self):
        self.set_y(-15)
        self.set_font("Helvetica", "I", 8)
        self.set_text_color(150, 150, 150)
        self.cell(0, 10, f"Generated {datetime.now().strftime('%Y-%m-%d %H:%M')} | Page {self.page_no()}/{{nb}}", align="C")

    def category_header(self, number, title):
        self.set_font("Helvetica", "B", 16)
        self.set_text_color(200, 60, 30)
        self.cell(0, 12, f"Category {number}: {title}", new_x="LMARGIN", new_y="NEXT")
        self.set_draw_color(200, 60, 30)
        self.set_line_width(0.6)
        self.line(10, self.get_y(), self.w - 10, self.get_y())
        self.ln(6)

    def section_title(self, title):
        self.set_font("Helvetica", "B", 12)
        self.set_text_color(200, 60, 30)
        self.cell(0, 9, title, new_x="LMARGIN", new_y="NEXT")
        self.set_draw_color(200, 60, 30)
        self.set_line_width(0.3)
        self.line(10, self.get_y(), self.w - 10, self.get_y())
        self.ln(3)

    def sub_title(self, title):
        self.set_font("Helvetica", "B", 10)
        self.set_text_color(50, 50, 50)
        self.cell(0, 7, title, new_x="LMARGIN", new_y="NEXT")
        self.ln(1)

    def body_text(self, text):
        self.set_font("Helvetica", "", 9)
        self.set_text_color(60, 60, 60)
        self.cell(0, 5, _latin1(text), new_x="LMARGIN", new_y="NEXT")

    def stat_line(self, label, value, color=None):
        self.set_font("Helvetica", "", 9)
        self.set_text_color(80, 80, 80)
        self.cell(55, 5, f"  {label}:", align="R")
        self.set_font("Helvetica", "B", 9)
        if color:
            self.set_text_color(*color)
        else:
            self.set_text_color(30, 30, 30)
        self.cell(0, 5, f"  {_latin1(str(value))}", new_x="LMARGIN", new_y="NEXT")

    def kpi_card(self, label, value, unit="", color=(30, 30, 30)):
        x = self.get_x()
        y = self.get_y()
        w = 60
        h = 18
        self.set_fill_color(245, 245, 248)
        self.rect(x, y, w, h, "F")
        self.set_xy(x, y + 2)
        self.set_font("Helvetica", "B", 14)
        self.set_text_color(*color)
        self.cell(w, 8, f"{value} {unit}", align="C")
        self.set_xy(x, y + 10)
        self.set_font("Helvetica", "", 7)
        self.set_text_color(120, 120, 120)
        self.cell(w, 6, label, align="C")
        self.set_xy(x + w + 5, y)

    def check_space(self, needed=30):
        """Add a new page if not enough space."""
        if self.get_y() + needed > self.h - 25:
            self.add_page()


def _latin1(text):
    """Sanitize text for Helvetica (Latin-1 encoding)."""
    if not isinstance(text, str):
        text = str(text)
    return text.encode("latin-1", errors="replace").decode("latin-1")


def _safe(val, default="N/A"):
    if val is None or (isinstance(val, (float, np.floating)) and pd.isna(val)):
        return default
    s = str(val)
    return default if s.lower() in ("nan", "none", "") else s


def _fmt_hole(val):
    s = _safe(val)
    if s == "N/A":
        return s
    try:
        return f'{float(s):g}"'
    except (ValueError, TypeError):
        return s


def _delta_str(val, fmt=".0f"):
    if val > 0:
        return f"+{val:{fmt}}"
    elif val < 0:
        return f"{val:{fmt}}"
    return "0"


def _table_header(pdf, col_widths, headers, bg_color=(50, 50, 60)):
    pdf.set_font("Helvetica", "B", 7)
    pdf.set_fill_color(*bg_color)
    pdf.set_text_color(255, 255, 255)
    for i, h in enumerate(headers):
        pdf.cell(col_widths[i], 6, h, align="C", fill=True)
    pdf.ln()


def _table_row_color(pdf, alt, flagged=False, flag_color=(255, 240, 238)):
    if flagged:
        pdf.set_fill_color(*flag_color)
    elif alt:
        pdf.set_fill_color(245, 245, 248)
    else:
        pdf.set_fill_color(255, 255, 255)
    pdf.set_text_color(60, 60, 60)
    pdf.set_font("Helvetica", "", 7)


# =========================================================================
# Category 1 rendering
# =========================================================================

def _render_cat1_summary(pdf, section):
    """1A: Weekly Summary table."""
    pdf.section_title("1A. Weekly Summary (JOB_TYPE / MOTOR_TYPE2)")

    ct = section["current_totals"]
    pt = section["previous_totals"]
    dt = section["delta_totals"]

    # Grand totals row
    cols = [55, 30, 30, 30]
    _table_header(pdf, cols, ["Metric", "Current Week", "Previous Week", "Delta"])
    for label, ck, fmt in [("Runs", "runs", "d"), ("Total Footage (ft)", "total_drill", ",.0f"), ("Total Hours", "total_hrs", ",.1f")]:
        _table_row_color(pdf, label == "Total Footage (ft)")
        pdf.cell(cols[0], 5, f"  {label}", fill=True)
        pdf.cell(cols[1], 5, f"{ct[ck]:{fmt}}", align="C", fill=True)
        pdf.cell(cols[2], 5, f"{pt[ck]:{fmt}}", align="C", fill=True)
        pdf.set_font("Helvetica", "B", 7)
        pdf.cell(cols[3], 5, _delta_str(dt[ck], fmt=".0f" if "d" in fmt or ",.0f" in fmt else ".1f"), align="C", fill=True)
        pdf.set_font("Helvetica", "", 7)
        pdf.ln()
    pdf.ln(3)

    # Breakdown table
    current = section["current"]
    if len(current) > 0:
        pdf.sub_title("Current Week Breakdown")
        bcols = [45, 45, 20, 30, 25]
        _table_header(pdf, bcols, ["JOB_TYPE", "MOTOR_TYPE2", "Runs", "Footage (ft)", "Hours"])
        alt = False
        for _, row in current.iterrows():
            _table_row_color(pdf, alt)
            pdf.cell(bcols[0], 5, f"  {_safe(row.get('JOB_TYPE'))[:25]}", fill=True)
            pdf.cell(bcols[1], 5, f"  {_safe(row.get('MOTOR_TYPE2'))[:25]}", fill=True)
            pdf.cell(bcols[2], 5, f"{row['runs']}", align="C", fill=True)
            pdf.cell(bcols[3], 5, f"{row['total_drill']:,.0f}", align="C", fill=True)
            pdf.cell(bcols[4], 5, f"{row['total_hrs']:,.1f}", align="C", fill=True)
            pdf.ln()
            alt = not alt
    pdf.ln(4)


def _render_cat1_curves(pdf, section):
    """1B: Curves Analysis."""
    pdf.check_space(40)
    pdf.section_title("1B. Curves Analysis (Motor_KPI Source)")

    cur = section["current"]
    prv = section["previous"]
    delta = section["delta"]

    cols = [55, 30, 30, 30]
    _table_header(pdf, cols, ["Metric", "Current Week", "Previous Week", "Delta"])
    metrics = [
        ("Motor_KPI Runs", "total_motor_kpi", None),
        ("With RUNS PER CUR", "total_with_rpc", None),
        ("1-Run Curves (best)", "one_run_count", "one_run_count"),
        ("Multi-Run Curves", "multi_run_count", "multi_run_count"),
        ("QC Needed", "qc_needed_count", "qc_needed_count"),
    ]
    alt = False
    for label, key, delta_key in metrics:
        _table_row_color(pdf, alt)
        pdf.cell(cols[0], 5, f"  {label}", fill=True)
        pdf.cell(cols[1], 5, f"{cur[key]}", align="C", fill=True)
        pdf.cell(cols[2], 5, f"{prv[key]}", align="C", fill=True)
        if delta_key:
            pdf.set_font("Helvetica", "B", 7)
            pdf.cell(cols[3], 5, _delta_str(delta[delta_key]), align="C", fill=True)
            pdf.set_font("Helvetica", "", 7)
        else:
            pdf.cell(cols[3], 5, "", fill=True)
        pdf.ln()
        alt = not alt

    # Multi-run operators
    op_multi = cur.get("operator_multi_run")
    if op_multi is not None and len(op_multi) > 0:
        pdf.ln(2)
        pdf.sub_title("Multi-Run Operators (Current Week)")
        ocols = [60, 25]
        _table_header(pdf, ocols, ["Operator", "Multi-Run Count"])
        alt = False
        for _, row in op_multi.head(10).iterrows():
            _table_row_color(pdf, alt)
            pdf.cell(ocols[0], 5, f"  {_latin1(str(row['OPERATOR']))[:35]}", fill=True)
            pdf.cell(ocols[1], 5, f"{int(row['multi_run_count'])}", align="C", fill=True)
            pdf.ln()
            alt = not alt
    pdf.ln(4)


def _render_cat1_pooh(pdf, section):
    """1C: Reason to POOH."""
    pdf.check_space(40)
    pdf.section_title(f"1C. Reason to POOH ({section['reason_col_used']})")

    cur = section["current"]
    prv = section["previous"]
    pdf.body_text(f"Filtered runs: {cur['total_filtered']} (current) | {prv['total_filtered']} (previous)")
    pdf.ln(2)

    # Category breakdown
    if len(cur["breakdown"]) > 0:
        cols = [50, 25, 25]
        _table_header(pdf, cols, ["Category", "Count", "%"])
        alt = False
        for _, row in cur["breakdown"].iterrows():
            _table_row_color(pdf, alt)
            pdf.cell(cols[0], 5, f"  {_latin1(str(row['category']))}", fill=True)
            pdf.cell(cols[1], 5, f"{row['count']}", align="C", fill=True)
            pdf.cell(cols[2], 5, f"{row['pct']:.1f}%", align="C", fill=True)
            pdf.ln()
            alt = not alt
        pdf.ln(3)

    # Motor issues by operator
    motor_ops = cur["motor_by_operator"]
    if len(motor_ops) > 0:
        pdf.sub_title("Motor Issues by Operator (Current Week)")
        mcols = [60, 25, 25, 25]
        _table_header(pdf, mcols, ["Operator", "Motor Issues", "Total Runs", "Motor %"], bg_color=(180, 50, 30))
        alt = False
        for _, row in motor_ops.head(10).iterrows():
            _table_row_color(pdf, alt, flagged=True, flag_color=(255, 245, 243))
            pdf.cell(mcols[0], 5, f"  {_latin1(str(row['OPERATOR']))[:35]}", fill=True)
            pdf.set_text_color(180, 40, 20)
            pdf.cell(mcols[1], 5, f"{int(row['motor_count'])}", align="C", fill=True)
            pdf.set_text_color(60, 60, 60)
            pdf.cell(mcols[2], 5, f"{int(row['total_count'])}", align="C", fill=True)
            pdf.cell(mcols[3], 5, f"{row['motor_pct']:.1f}%", align="C", fill=True)
            pdf.ln()
            alt = not alt
    pdf.ln(4)


# =========================================================================
# Category 2 rendering
# =========================================================================

def _render_cat2_longest_runs(pdf, section, cm_label, pm_label):
    """2A: Longest Runs."""
    pdf.section_title("2A. Longest Runs (Top 5 by Total Footage)")

    for label, key in [(cm_label, "current_month"), (pm_label, "previous_month")]:
        df = section[key]
        if len(df) == 0:
            pdf.body_text(f"{label}: No data")
            continue

        pdf.sub_title(label)
        cols = [8, 42, 50, 26, 20, 18, 22, 16, 24]
        headers = ["#", "Operator", "Well", "Footage", "ROP", "Hrs", "Type", "Hole", "Basin"]
        _table_header(pdf, cols, headers)
        alt = False
        for _, run in df.iterrows():
            pdf.check_space(8)
            _table_row_color(pdf, alt)
            pdf.cell(cols[0], 5, f"{int(run['rank'])}", align="C", fill=True)
            pdf.cell(cols[1], 5, f"  {_latin1(str(run['operator']))[:25]}", fill=True)
            pdf.cell(cols[2], 5, f"  {_latin1(str(run['well']))[:30]}", fill=True)
            pdf.set_font("Helvetica", "B", 7)
            pdf.cell(cols[3], 5, f"{run['total_drill']:,.0f}", align="C", fill=True)
            pdf.set_font("Helvetica", "", 7)
            rop = f"{run['avg_rop']:.1f}" if run['avg_rop'] else "N/A"
            hrs = f"{run['drilling_hours']:.1f}" if run['drilling_hours'] else "N/A"
            pdf.cell(cols[4], 5, rop, align="C", fill=True)
            pdf.cell(cols[5], 5, hrs, align="C", fill=True)
            pdf.cell(cols[6], 5, _safe(run.get('motor_type2'))[:12], align="C", fill=True)
            pdf.cell(cols[7], 5, _fmt_hole(run.get('hole_size')), align="C", fill=True)
            pdf.cell(cols[8], 5, _safe(run.get('basin'))[:14], align="C", fill=True)
            pdf.ln()
            alt = not alt
        pdf.ln(3)
    pdf.ln(2)


def _render_cat2_monthly_summary(pdf, section, cm_label, pm_label):
    """2B: Monthly Summary."""
    pdf.check_space(35)
    pdf.section_title(f"2B. Monthly Summary ({cm_label} vs {pm_label})")

    ct = section["current_totals"]
    pt = section["previous_totals"]
    dt = section["delta_totals"]

    cols = [55, 30, 30, 30]
    _table_header(pdf, cols, ["Metric", cm_label[:18], pm_label[:18], "Delta"])
    for label, ck, fmt in [("Runs", "runs", "d"), ("Total Footage (ft)", "total_drill", ",.0f"), ("Total Hours", "total_hrs", ",.1f")]:
        _table_row_color(pdf, label == "Total Footage (ft)")
        pdf.cell(cols[0], 5, f"  {label}", fill=True)
        pdf.cell(cols[1], 5, f"{ct[ck]:{fmt}}", align="C", fill=True)
        pdf.cell(cols[2], 5, f"{pt[ck]:{fmt}}", align="C", fill=True)
        pdf.set_font("Helvetica", "B", 7)
        d_fmt = ".0f" if "d" in fmt or ",.0f" in fmt else ".1f"
        pdf.cell(cols[3], 5, _delta_str(dt[ck], fmt=d_fmt), align="C", fill=True)
        pdf.set_font("Helvetica", "", 7)
        pdf.ln()
    pdf.ln(4)


def _render_cat2_fastest(pdf, section):
    """2C: Fastest Sections."""
    pdf.check_space(35)
    pdf.section_title("2C. Fastest Sections (Highest AVG_ROP by Hole Size)")

    by_hole = section.get("by_hole_size", {})
    if not by_hole:
        pdf.body_text("No data with sufficient runs.")
        return

    for hole_size in sorted(by_hole.keys()):
        info = by_hole[hole_size]
        pdf.check_space(20)
        pdf.sub_title(f"Hole Size: {_fmt_hole(hole_size)} | Avg ROP: {info['hole_avg_rop']:.1f} ft/hr | Runs: {info['run_count']}")

        top = info["top_operators"]
        if len(top) > 0:
            cols = [50, 25, 30, 25, 25]
            _table_header(pdf, cols, ["Operator", "AVG ROP", "Footage", "Basin", "Phase"], bg_color=(30, 140, 60))
            alt = False
            for _, row in top.iterrows():
                _table_row_color(pdf, alt, flagged=True, flag_color=(235, 255, 240))
                pdf.cell(cols[0], 5, f"  {_latin1(str(row['operator']))[:30]}", fill=True)
                pdf.set_text_color(20, 120, 40)
                pdf.set_font("Helvetica", "B", 7)
                pdf.cell(cols[1], 5, f"{row['avg_rop']:.1f}", align="C", fill=True)
                pdf.set_font("Helvetica", "", 7)
                pdf.set_text_color(60, 60, 60)
                td = f"{row['total_drill']:,.0f}" if row.get('total_drill') else "N/A"
                pdf.cell(cols[2], 5, td, align="C", fill=True)
                pdf.cell(cols[3], 5, _safe(row.get('basin'))[:14], align="C", fill=True)
                pdf.cell(cols[4], 5, _safe(row.get('phase')), align="C", fill=True)
                pdf.ln()
                alt = not alt
            pdf.ln(2)
    pdf.ln(2)


def _render_cat2_operator_success(pdf, section):
    """2D: Operator Success Rate."""
    pdf.check_space(35)
    pdf.section_title("2D. Operator Success Rate (TD Runs / Total)")

    if not isinstance(section, pd.DataFrame) or len(section) == 0:
        pdf.body_text("No data available.")
        return

    cols = [60, 25, 25, 30]
    _table_header(pdf, cols, ["Operator", "TD Runs", "Total Runs", "Success %"])
    alt = False
    for _, row in section.head(15).iterrows():
        pdf.check_space(7)
        _table_row_color(pdf, alt)
        pdf.cell(cols[0], 5, f"  {_latin1(str(row['operator']))[:35]}", fill=True)
        pdf.cell(cols[1], 5, f"{int(row['td_runs'])}", align="C", fill=True)
        pdf.cell(cols[2], 5, f"{int(row['total_runs'])}", align="C", fill=True)
        # Color code success %
        pct = row['success_pct']
        if pct >= 70:
            pdf.set_text_color(20, 120, 40)
        elif pct < 40:
            pdf.set_text_color(200, 40, 20)
        pdf.set_font("Helvetica", "B", 7)
        pdf.cell(cols[3], 5, f"{pct:.1f}%", align="C", fill=True)
        pdf.set_font("Helvetica", "", 7)
        pdf.set_text_color(60, 60, 60)
        pdf.ln()
        alt = not alt
    pdf.ln(4)


def _render_cat2_motor_failures(pdf, section):
    """2E: Motor Failures by Operator."""
    pdf.check_space(30)
    pdf.section_title("2E. Motor Failures by Operator")

    if not isinstance(section, pd.DataFrame) or len(section) == 0:
        pdf.body_text("No motor failures this month.")
        return

    cols = [60, 25, 25, 30]
    _table_header(pdf, cols, ["Operator", "Failures", "Total Runs", "Failure %"], bg_color=(180, 50, 30))
    alt = False
    for _, row in section.head(15).iterrows():
        pdf.check_space(7)
        _table_row_color(pdf, alt, flagged=True, flag_color=(255, 245, 243))
        pdf.cell(cols[0], 5, f"  {_latin1(str(row['operator']))[:35]}", fill=True)
        pdf.set_text_color(180, 40, 20)
        pdf.cell(cols[1], 5, f"{int(row['failure_runs'])}", align="C", fill=True)
        pdf.set_text_color(60, 60, 60)
        pdf.cell(cols[2], 5, f"{int(row['total_runs'])}", align="C", fill=True)
        pdf.cell(cols[3], 5, f"{row['failure_pct']:.1f}%", align="C", fill=True)
        pdf.ln()
        alt = not alt
    pdf.ln(4)


def _render_cat2_curve_success(pdf, section):
    """2F: Curve Success Rate."""
    pdf.check_space(30)
    pdf.section_title("2F. Curve Success Rate (RUNS PER CUR = 1)")

    data = section.get("data") if isinstance(section, dict) else section
    if not isinstance(data, pd.DataFrame) or len(data) == 0:
        pdf.body_text("No curve data available.")
        return

    total_with = section.get("total_with_data", 0) if isinstance(section, dict) else 0
    total_mkpi = section.get("total_motor_kpi", 0) if isinstance(section, dict) else 0
    pdf.body_text(f"Motor_KPI runs with RUNS PER CUR data: {total_with} of {total_mkpi}")
    pdf.ln(2)

    cols = [60, 25, 25, 30]
    _table_header(pdf, cols, ["Operator", "1-Run", "Total Curves", "Success %"])
    alt = False
    for _, row in data.head(15).iterrows():
        pdf.check_space(7)
        _table_row_color(pdf, alt)
        pdf.cell(cols[0], 5, f"  {_latin1(str(row['operator']))[:35]}", fill=True)
        pdf.cell(cols[1], 5, f"{int(row['one_run'])}", align="C", fill=True)
        pdf.cell(cols[2], 5, f"{int(row['total_curves'])}", align="C", fill=True)
        pct = row['success_pct']
        if pct >= 70:
            pdf.set_text_color(20, 120, 40)
        elif pct < 40:
            pdf.set_text_color(200, 40, 20)
        pdf.set_font("Helvetica", "B", 7)
        pdf.cell(cols[3], 5, f"{pct:.1f}%", align="C", fill=True)
        pdf.set_font("Helvetica", "", 7)
        pdf.set_text_color(60, 60, 60)
        pdf.ln()
        alt = not alt
    pdf.ln(4)


# =========================================================================
# Category 3 rendering (existing KPIs, adapted)
# =========================================================================

def _render_cat3_avg_rop(pdf, summary):
    """3A: AVG ROP Analysis."""
    if summary is None:
        return

    pdf.section_title("3A. AVG ROP Analysis")
    pdf.stat_line("Runs with ROP data", str(summary['total_runs']))
    pdf.stat_line("Week Average", f"{summary['week_avg']} ft/hr")
    pdf.stat_line("Week Median", f"{summary['week_median']} ft/hr")
    pdf.ln(3)

    # Basin breakdown
    if "basin_breakdown" in summary:
        pdf.sub_title("ROP by Basin")
        bcols = [60, 35, 25]
        _table_header(pdf, bcols, ["Basin", "Avg ROP (ft/hr)", "Runs"])
        alt = False
        for basin, row in summary["basin_breakdown"].iterrows():
            _table_row_color(pdf, alt)
            pdf.cell(bcols[0], 5, f"  {_latin1(str(basin))[:35]}", fill=True)
            pdf.cell(bcols[1], 5, f"{row['mean']:.1f}", align="C", fill=True)
            pdf.cell(bcols[2], 5, f"{int(row['count'])}", align="C", fill=True)
            pdf.ln()
            alt = not alt
        pdf.ln(3)

    results = summary["results"]

    # Underperforming runs
    flagged_below = results[results["flag"] == "below"].sort_values("diff_pct")
    if len(flagged_below) > 0:
        pdf.check_space(20)
        pdf.sub_title(f"Underperforming Runs ({len(flagged_below)} flagged)")
        col_widths = [38, 48, 16, 22, 18, 24, 22, 20, 18, 24, 20]
        headers = ["Operator", "Well", "Hole", "ROP", "Diff%", "Basin", "County", "Motor", "L/S", "Formation", "Match"]
        _table_header(pdf, col_widths, headers, bg_color=(200, 60, 30))
        alt = False
        for _, run in flagged_below.head(12).iterrows():
            pdf.check_space(7)
            _table_row_color(pdf, alt, flagged=True)
            pdf.cell(col_widths[0], 5, f"  {_latin1(str(run['operator']))[:22]}", fill=True)
            pdf.cell(col_widths[1], 5, f"  {_latin1(str(run['well']))[:28]}", fill=True)
            pdf.cell(col_widths[2], 5, _fmt_hole(run['hole_size']), align="C", fill=True)
            pdf.set_text_color(200, 40, 20)
            pdf.cell(col_widths[3], 5, f"{run['value']:.1f}", align="C", fill=True)
            pdf.cell(col_widths[4], 5, f"{run['diff_pct']:+.0f}%", align="C", fill=True)
            pdf.set_text_color(60, 60, 60)
            pdf.cell(col_widths[5], 5, _safe(run['basin'])[:14], align="C", fill=True)
            pdf.cell(col_widths[6], 5, _safe(run['county'])[:12], align="C", fill=True)
            pdf.cell(col_widths[7], 5, _safe(run['motor_model'])[:12], align="C", fill=True)
            pdf.cell(col_widths[8], 5, _safe(run['lobe_stage'])[:10], align="C", fill=True)
            pdf.cell(col_widths[9], 5, _safe(run['formation'])[:14], align="C", fill=True)
            pdf.set_font("Helvetica", "I", 5)
            pdf.cell(col_widths[10], 5, _safe(run.get('match_level'))[:12], align="C", fill=True)
            pdf.set_font("Helvetica", "", 7)
            pdf.ln()
            alt = not alt
        pdf.ln(3)

    # Top performers
    flagged_above = results[results["flag"] == "above"].sort_values("diff_pct", ascending=False)
    if len(flagged_above) > 0:
        pdf.check_space(20)
        pdf.sub_title(f"Top Performers ({len(flagged_above)} highlighted)")
        col_widths = [38, 48, 16, 22, 18, 24, 22, 20, 18, 24, 20]
        headers = ["Operator", "Well", "Hole", "ROP", "Diff%", "Basin", "County", "Motor", "L/S", "Formation", "Match"]
        _table_header(pdf, col_widths, headers, bg_color=(30, 140, 60))
        alt = False
        for _, run in flagged_above.head(8).iterrows():
            pdf.check_space(7)
            _table_row_color(pdf, alt, flagged=True, flag_color=(235, 255, 240))
            pdf.cell(col_widths[0], 5, f"  {_latin1(str(run['operator']))[:22]}", fill=True)
            pdf.cell(col_widths[1], 5, f"  {_latin1(str(run['well']))[:28]}", fill=True)
            pdf.cell(col_widths[2], 5, _fmt_hole(run['hole_size']), align="C", fill=True)
            pdf.set_text_color(20, 120, 40)
            pdf.cell(col_widths[3], 5, f"{run['value']:.1f}", align="C", fill=True)
            pdf.cell(col_widths[4], 5, f"{run['diff_pct']:+.0f}%", align="C", fill=True)
            pdf.set_text_color(60, 60, 60)
            pdf.cell(col_widths[5], 5, _safe(run['basin'])[:14], align="C", fill=True)
            pdf.cell(col_widths[6], 5, _safe(run['county'])[:12], align="C", fill=True)
            pdf.cell(col_widths[7], 5, _safe(run['motor_model'])[:12], align="C", fill=True)
            pdf.cell(col_widths[8], 5, _safe(run['lobe_stage'])[:10], align="C", fill=True)
            pdf.cell(col_widths[9], 5, _safe(run['formation'])[:14], align="C", fill=True)
            pdf.set_font("Helvetica", "I", 5)
            pdf.cell(col_widths[10], 5, _safe(run.get('match_level'))[:12], align="C", fill=True)
            pdf.set_font("Helvetica", "", 7)
            pdf.ln()
            alt = not alt
        pdf.ln(3)


def _render_cat3_longest_runs(pdf, summary):
    """3B: Longest Runs."""
    if summary is None:
        return

    pdf.add_page()
    pdf.section_title(f"3B. Longest Runs - Top {summary['top_n']}")
    pdf.stat_line("Runs with drill data", str(summary['total_runs']))
    pdf.stat_line("Week Average", f"{summary['week_avg']:,.0f} ft")
    pdf.stat_line("Week Max", f"{summary['week_max']:,.0f} ft")
    pdf.ln(3)

    results = summary["results"]
    if len(results) == 0:
        return

    dcols = [8, 38, 48, 26, 20, 20, 16, 24, 22, 24, 24]
    dheaders = ["#", "Operator", "Well", "Footage", "Hours", "ROP", "Hole", "Basin", "County", "Phase", "Motor"]
    _table_header(pdf, dcols, dheaders)
    alt = False
    for _, run in results.iterrows():
        pdf.check_space(12)
        _table_row_color(pdf, alt)
        pdf.cell(dcols[0], 5, f"{int(run['rank'])}", align="C", fill=True)
        pdf.cell(dcols[1], 5, f"  {_latin1(str(run['operator']))[:22]}", fill=True)
        pdf.cell(dcols[2], 5, f"  {_latin1(str(run['well']))[:28]}", fill=True)
        pdf.set_font("Helvetica", "B", 7)
        pdf.cell(dcols[3], 5, f"{run['value']:,.0f}", align="C", fill=True)
        pdf.set_font("Helvetica", "", 7)
        hrs = f"{run['drilling_hours']:.1f}" if run['drilling_hours'] and not pd.isna(run['drilling_hours']) else "N/A"
        rop = f"{run['avg_rop']:.1f}" if run['avg_rop'] and not pd.isna(run['avg_rop']) else "N/A"
        pdf.cell(dcols[4], 5, hrs, align="C", fill=True)
        pdf.cell(dcols[5], 5, rop, align="C", fill=True)
        pdf.cell(dcols[6], 5, _fmt_hole(run['hole_size']), align="C", fill=True)
        pdf.cell(dcols[7], 5, _safe(run['basin'])[:14], align="C", fill=True)
        pdf.cell(dcols[8], 5, _safe(run['county'])[:12], align="C", fill=True)
        pdf.cell(dcols[9], 5, _safe(run['phase']), align="C", fill=True)
        pdf.cell(dcols[10], 5, _safe(run['motor_model'])[:14], align="C", fill=True)
        pdf.ln()

        # Historical comparison line
        if run["baseline_mean"]:
            pdf.set_font("Helvetica", "I", 5)
            pdf.set_text_color(120, 120, 120)
            match_lvl = _safe(run.get('match_level'))
            pdf.cell(dcols[0], 4, "", fill=False)
            pdf.cell(sum(dcols[1:]), 4,
                     f"    Baseline: {run['baseline_mean']:,.0f} ft (n={int(run['baseline_count'])}) | {run['diff_pct']:+.1f}% vs avg | Match: {match_lvl}",
                     fill=False)
            pdf.ln()
            pdf.set_font("Helvetica", "", 7)

        alt = not alt
    pdf.ln(4)


def _render_cat3_sliding_pct(pdf, summary):
    """3C: Sliding %."""
    if summary is None or summary["total_lat_runs"] == 0:
        return

    pdf.check_space(40)
    pdf.section_title("3C. Sliding % (LAT Phase Only)")
    pdf.stat_line("LAT Runs", f"{summary['total_lat_runs']} of {summary['total_new_runs']} total")
    pdf.stat_line("Week Average", f"{summary['week_avg']}%")
    pdf.stat_line("Week Median", f"{summary['week_median']}%")
    if summary["hist_threshold_pct"]:
        pdf.stat_line("75th Pct Threshold", f"{summary['hist_threshold_pct']}%")
    pdf.ln(3)

    results = summary["results"]
    if len(results) == 0:
        return

    sorted_slides = results.sort_values("value", ascending=False)
    scols = [38, 50, 16, 22, 26, 26, 24, 22, 22, 24]
    sheaders = ["Operator", "Well", "Hole", "Slide%", "Slide(ft)", "Total(ft)", "Basin", "County", "Motor", "Match"]
    _table_header(pdf, scols, sheaders)
    alt = False
    for _, run in sorted_slides.iterrows():
        pdf.check_space(7)
        is_flagged = run["flag"] == "above"
        _table_row_color(pdf, alt, flagged=is_flagged)
        pdf.cell(scols[0], 5, f"  {_latin1(str(run['operator']))[:22]}", fill=True)
        pdf.cell(scols[1], 5, f"  {_latin1(str(run['well']))[:28]}", fill=True)
        pdf.cell(scols[2], 5, _fmt_hole(run['hole_size']), align="C", fill=True)
        if is_flagged:
            pdf.set_text_color(200, 40, 20)
        pdf.set_font("Helvetica", "B", 7)
        pdf.cell(scols[3], 5, f"{run['value']:.1f}%", align="C", fill=True)
        pdf.set_font("Helvetica", "", 7)
        pdf.set_text_color(60, 60, 60)
        pdf.cell(scols[4], 5, f"{run['slide_drilled']:,.0f}", align="C", fill=True)
        pdf.cell(scols[5], 5, f"{run['total_drill']:,.0f}", align="C", fill=True)
        pdf.cell(scols[6], 5, _safe(run['basin'])[:14], align="C", fill=True)
        pdf.cell(scols[7], 5, _safe(run['county'])[:12], align="C", fill=True)
        pdf.cell(scols[8], 5, _safe(run['motor_model'])[:12], align="C", fill=True)
        pdf.set_font("Helvetica", "I", 5)
        pdf.cell(scols[9], 5, _safe(run.get('match_level'))[:14], align="C", fill=True)
        pdf.set_font("Helvetica", "", 7)
        pdf.ln()
        alt = not alt
    pdf.ln(4)


def _render_cat3_patterns(pdf, patterns):
    """3D: Pattern Highlights."""
    if patterns is None:
        return

    highlights = patterns.get("highlights", [])
    lowlights = patterns.get("lowlights", [])

    if not highlights and not lowlights:
        return

    pdf.check_space(30)
    pdf.section_title("3D. Pattern Highlights")

    if highlights:
        pdf.sub_title(f"Above Baseline ({len(highlights)} patterns)")
        for h in highlights[:5]:
            pdf.check_space(15)
            group_desc = " + ".join(f"{k}={v}" for k, v in h["grouping_values"].items())
            ops_str = ", ".join(f"{o['operator']} ({o['avg']})" for o in h["top_operators"])
            pdf.set_font("Helvetica", "B", 8)
            pdf.set_text_color(20, 120, 40)
            pdf.cell(0, 5, f"  {_latin1(group_desc)} | {h['metric']}: {h['diff_pct']:+.1f}%", new_x="LMARGIN", new_y="NEXT")
            pdf.set_font("Helvetica", "", 7)
            pdf.set_text_color(60, 60, 60)
            pdf.cell(0, 4, f"    Week: {h['week_avg']} vs Baseline: {h['baseline_avg']} (n={h['baseline_count']})", new_x="LMARGIN", new_y="NEXT")
            pdf.cell(0, 4, f"    Top operators: {_latin1(ops_str)}", new_x="LMARGIN", new_y="NEXT")
            pdf.ln(1)

    if lowlights:
        pdf.ln(2)
        pdf.sub_title(f"Below Baseline ({len(lowlights)} patterns)")
        for item in lowlights[:5]:
            pdf.check_space(15)
            group_desc = " + ".join(f"{k}={v}" for k, v in item["grouping_values"].items())
            ops_str = ", ".join(f"{o['operator']} ({o['avg']})" for o in item["top_operators"])
            pdf.set_font("Helvetica", "B", 8)
            pdf.set_text_color(200, 40, 20)
            pdf.cell(0, 5, f"  {_latin1(group_desc)} | {item['metric']}: {item['diff_pct']:+.1f}%", new_x="LMARGIN", new_y="NEXT")
            pdf.set_font("Helvetica", "", 7)
            pdf.set_text_color(60, 60, 60)
            pdf.cell(0, 4, f"    Week: {item['week_avg']} vs Baseline: {item['baseline_avg']} (n={item['baseline_count']})", new_x="LMARGIN", new_y="NEXT")
            pdf.cell(0, 4, f"    Bottom operators: {_latin1(ops_str)}", new_x="LMARGIN", new_y="NEXT")
            pdf.ln(1)
    pdf.ln(4)


# =========================================================================
# Main entry point
# =========================================================================

def generate_pdf(all_results, output_dir=None, report_title=None):
    """Generate the full PDF performance report with all 3 categories."""

    meta = all_results["meta"]
    pdf = PerformanceReport(meta)
    pdf.alias_nb_pages()
    pdf.add_page()

    # =====================================================================
    # TITLE PAGE
    # =====================================================================
    if report_title:
        pdf.set_font("Helvetica", "B", 16)
        pdf.set_text_color(30, 30, 30)
        pdf.cell(0, 14, _latin1(report_title), new_x="LMARGIN", new_y="NEXT")
    else:
        pdf.set_font("Helvetica", "B", 22)
        pdf.set_text_color(30, 30, 30)
        pdf.cell(0, 14, "PA1 - Weekly Performance Report", new_x="LMARGIN", new_y="NEXT")

    week = meta["week"]
    ws = meta["week_start"]
    we = meta["week_end"]
    pdf.set_font("Helvetica", "", 12)
    pdf.set_text_color(100, 100, 100)
    pdf.cell(0, 8, f"Week {week}  |  {ws.date()} to {we.date()}", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(4)

    # KPI summary cards
    cat3 = all_results.get("category3")
    rop_summary = cat3["sections"].get("A_avg_rop") if cat3 else None
    drill_summary = cat3["sections"].get("B_longest_runs") if cat3 else None
    slide_summary = cat3["sections"].get("C_sliding_pct") if cat3 else None

    pdf.kpi_card("Total New Runs", str(meta["total_new_runs"]))
    if rop_summary:
        pdf.kpi_card("Week Avg ROP", f"{rop_summary['week_avg']}", "ft/hr")
        pdf.kpi_card("Underperforming", str(rop_summary['flagged_below']), "runs", (200, 60, 30))
        pdf.kpi_card("Top Performers", str(rop_summary['flagged_above']), "runs", (30, 140, 60))
    pdf.ln(24)

    if drill_summary:
        pdf.kpi_card("Week Avg Footage", f"{drill_summary['week_avg']:,.0f}", "ft")
        pdf.kpi_card("Max Footage", f"{drill_summary['week_max']:,.0f}", "ft")
    if slide_summary and slide_summary["total_lat_runs"] > 0:
        pdf.kpi_card("LAT Runs", str(slide_summary['total_lat_runs']))
        pdf.kpi_card("Avg Sliding %", f"{slide_summary['week_avg']}", "%")
    pdf.ln(24)

    # =====================================================================
    # CATEGORY 1: Week vs Previous Week
    # =====================================================================
    cat1 = all_results.get("category1")
    if cat1:
        pdf.add_page()
        pdf.category_header(1, cat1["category"])
        sections = cat1["sections"]
        _render_cat1_summary(pdf, sections["A_weekly_summary"])
        _render_cat1_curves(pdf, sections["B_curves"])
        _render_cat1_pooh(pdf, sections["C_reason_pooh"])

    # =====================================================================
    # CATEGORY 2: Monthly Highlights
    # =====================================================================
    cat2 = all_results.get("category2")
    if cat2:
        pdf.add_page()
        pdf.category_header(2, cat2["category"])
        cm_label = cat2.get("current_month_label", "Current Month")
        pm_label = cat2.get("previous_month_label", "Previous Month")
        sections = cat2["sections"]
        _render_cat2_longest_runs(pdf, sections["A_longest_runs"], cm_label, pm_label)
        _render_cat2_monthly_summary(pdf, sections["B_monthly_summary"], cm_label, pm_label)
        _render_cat2_fastest(pdf, sections["C_fastest_sections"])
        _render_cat2_operator_success(pdf, sections["D_operator_success"])
        _render_cat2_motor_failures(pdf, sections["E_motor_failures"])
        _render_cat2_curve_success(pdf, sections["F_curve_success"])

    # =====================================================================
    # CATEGORY 3: Historical Analysis
    # =====================================================================
    if cat3:
        pdf.add_page()
        pdf.category_header(3, cat3["category"])
        sections = cat3["sections"]
        _render_cat3_avg_rop(pdf, sections.get("A_avg_rop"))
        _render_cat3_longest_runs(pdf, sections.get("B_longest_runs"))
        _render_cat3_sliding_pct(pdf, sections.get("C_sliding_pct"))
        _render_cat3_patterns(pdf, sections.get("D_pattern_highlights"))

    # =====================================================================
    # SAVE PDF
    # =====================================================================
    if output_dir is None:
        output_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    if report_title:
        safe_name = report_title.replace("/", "-").replace("\\", "-").replace(":", "-")
        safe_name = safe_name.replace("*", "").replace("?", "").replace('"', "")
        safe_name = safe_name.replace("<", "").replace(">", "").replace("|", "")
        filename = f"{safe_name}.pdf"
    else:
        filename = f"PA1_Report_{week}_{datetime.now().strftime('%Y%m%d')}.pdf"

    filepath = os.path.join(output_dir, filename)
    pdf.output(filepath)

    return filepath
