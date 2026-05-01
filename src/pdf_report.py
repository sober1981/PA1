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
        self.set_auto_page_break(auto=True, margin=15)

    def header(self):
        self.set_font("Helvetica", "B", 9)
        self.set_text_color(100, 100, 100)
        week = self.meta.get("week", "")
        self.cell(0, 5, f"Scout Downhole - PA1 Performance Report | Week {week}", align="L")
        self.ln(2)
        self.set_draw_color(200, 60, 30)
        self.set_line_width(0.6)
        self.line(10, self.get_y(), self.w - 10, self.get_y())
        self.ln(3)

    def footer(self):
        self.set_y(-12)
        self.set_font("Helvetica", "I", 7)
        self.set_text_color(150, 150, 150)
        self.cell(0, 8, f"Generated {datetime.now().strftime('%Y-%m-%d %H:%M')} | Page {self.page_no()}/{{nb}}", align="C")

    def category_header(self, number, title):
        self.set_font("Helvetica", "B", 12)
        self.set_text_color(200, 60, 30)
        self.cell(0, 8, f"Category {number}: {title}", new_x="LMARGIN", new_y="NEXT")
        self.set_draw_color(200, 60, 30)
        self.set_line_width(0.5)
        self.line(10, self.get_y(), self.w - 10, self.get_y())
        self.ln(3)

    def section_title(self, title):
        self.set_font("Helvetica", "B", 10)
        self.set_text_color(200, 60, 30)
        self.cell(0, 6, title, new_x="LMARGIN", new_y="NEXT")
        self.set_draw_color(200, 60, 30)
        self.set_line_width(0.3)
        self.line(10, self.get_y(), self.w - 10, self.get_y())
        self.ln(2)

    def sub_title(self, title):
        self.set_font("Helvetica", "B", 9)
        self.set_text_color(50, 50, 50)
        self.cell(0, 5, title, new_x="LMARGIN", new_y="NEXT")
        self.ln(0.5)

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

    def cell(self, w=None, h=None, text="", *args, **kwargs):
        """Override to auto-sanitize text for latin-1 fonts."""
        return super().cell(w, h, _latin1(text), *args, **kwargs)

    def multi_cell(self, w=None, h=None, text="", *args, **kwargs):
        """Override to auto-sanitize text for latin-1 fonts."""
        return super().multi_cell(w, h, _latin1(text), *args, **kwargs)

    def check_space(self, needed=30):
        """Add a new page if not enough space."""
        if self.get_y() + needed > self.h - 25:
            self.add_page()


_UNICODE_REPLACEMENTS = {
    "\u2014": "--",   # em dash
    "\u2013": "-",    # en dash
    "\u2018": "'",    # left single quote
    "\u2019": "'",    # right single quote
    "\u201c": '"',    # left double quote
    "\u201d": '"',    # right double quote
    "\u2026": "...",  # ellipsis
    "\u2022": "*",    # bullet
    "\u00b7": "*",    # middle dot
    "\u2032": "'",    # prime
    "\u2033": '"',    # double prime
    "\u2010": "-",    # hyphen
    "\u2011": "-",    # non-breaking hyphen
    "\u2012": "-",    # figure dash
    "\u2015": "--",   # horizontal bar
    "\u2212": "-",    # minus sign
}


def _latin1(text):
    """Sanitize text for Helvetica (Latin-1 encoding)."""
    if not isinstance(text, str):
        text = str(text)
    for char, replacement in _UNICODE_REPLACEMENTS.items():
        text = text.replace(char, replacement)
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

_MOTOR_RGB = {
    "TDI CONV":   (213, 232, 212),
    "CAM RENTAL": (231, 230, 230),
    "CAM DD":     (244, 204, 204),
    "3RD PARTY":  (207, 226, 243),
}
_DIFF_NEG_RGB = (192, 0, 0)
_DIFF_POS_RGB = (0, 100, 0)
_GRAD_LOW_RGB = (248, 105, 107)   # red
_GRAD_MID_RGB = (255, 235, 132)   # yellow
_GRAD_HIGH_RGB = (99, 190, 123)   # green


def _motor_key(s):
    if not s:
        return None
    k = str(s).upper().strip()
    return k if k in _MOTOR_RGB else None


def _interp(c1, c2, t):
    return tuple(int(a + (b - a) * t) for a, b in zip(c1, c2))


def _gradient_color(value, vmin, vmax, more_is_better=True):
    """Return RGB tuple along red→yellow→green (or reversed for more-is-worse)."""
    if vmax <= vmin:
        return (255, 255, 255)
    t = (value - vmin) / (vmax - vmin)
    t = max(0.0, min(1.0, t))
    if not more_is_better:
        t = 1.0 - t
    if t <= 0.5:
        return _interp(_GRAD_LOW_RGB, _GRAD_MID_RGB, t * 2)
    return _interp(_GRAD_MID_RGB, _GRAD_HIGH_RGB, (t - 0.5) * 2)


def _diff_text(value, fmt=",.0f"):
    if value is None:
        return ""
    if value > 0:
        return f"+{value:{fmt}}"
    return f"{value:{fmt}}"


def _kpi_summary_table(pdf, kpi):
    """Summary KPI table — per hole size + Grand Total."""
    cols = [10, 18, 18, 22, 24, 22, 24, 22, 22, 22, 22]  # 226mm
    headers = ["Week", "Hole Size", "Total Runs", "Total Hrs", "Hrs Diff vs Prev",
               "% G (Hrs)", "Total Drill", "Ftg vs Prev", "% G (Drill)",
               "Total Incidents", "OP w/ More Runs"]
    pdf.sub_title("Summary Table (per Hole Size)")
    pdf.set_x((pdf.w - sum(cols)) / 2)  # center
    _table_header(pdf, cols, headers, bg_color=(31, 78, 120))

    week_str = str(kpi.get("week", ""))
    week_num = week_str.split("-W")[-1] if "-W" in week_str else week_str
    blocks = kpi.get("blocks", [])

    # Min/max across data rows for gradient
    runs_vals = [b["curr_total"]["runs"] for b in blocks]
    hrs_vals = [b["curr_total"]["total_hrs"] for b in blocks]
    drill_vals = [b["curr_total"]["total_drill"] for b in blocks]
    inc_vals = [b["curr_total"]["incident_count"] for b in blocks]

    pdf.set_font("Helvetica", "", 7)
    for b in blocks:
        ct = b["curr_total"]
        diff = b["diff"]
        pdf.set_x((pdf.w - sum(cols)) / 2)
        # Cells with gradient fills
        # Week
        pdf.set_fill_color(255, 255, 255)
        pdf.set_text_color(60, 60, 60)
        pdf.cell(cols[0], 5, str(week_num), align="C", fill=True)
        pdf.cell(cols[1], 5, f'{b["hole_size"]:g}"', align="C", fill=True)
        # Total Runs (gradient: more is better)
        r, g, bl = _gradient_color(ct["runs"], min(runs_vals), max(runs_vals), True)
        pdf.set_fill_color(r, g, bl)
        pdf.cell(cols[2], 5, f"{ct['runs']}", align="C", fill=True)
        # Total Hrs (gradient: more is better)
        r, g, bl = _gradient_color(ct["total_hrs"], min(hrs_vals), max(hrs_vals), True)
        pdf.set_fill_color(r, g, bl)
        pdf.cell(cols[3], 5, f"{ct['total_hrs']:,.1f}", align="C", fill=True)
        # Hrs Diff (red/green font)
        pdf.set_fill_color(255, 255, 255)
        if diff["total_hrs"] < 0:
            pdf.set_text_color(*_DIFF_NEG_RGB)
        elif diff["total_hrs"] > 0:
            pdf.set_text_color(*_DIFF_POS_RGB)
        pdf.cell(cols[4], 5, _diff_text(diff["total_hrs"], ",.1f"), align="C", fill=True)
        pdf.set_text_color(60, 60, 60)
        # % G Hrs
        pdf.cell(cols[5], 5, f"{ct['g_pct_hrs']*100:.1f}%", align="C", fill=True)
        # Total Drill (gradient)
        r, g, bl = _gradient_color(ct["total_drill"], min(drill_vals), max(drill_vals), True)
        pdf.set_fill_color(r, g, bl)
        pdf.cell(cols[6], 5, f"{ct['total_drill']:,.0f}", align="C", fill=True)
        # Ftg Diff
        pdf.set_fill_color(255, 255, 255)
        if diff["total_drill"] < 0:
            pdf.set_text_color(*_DIFF_NEG_RGB)
        elif diff["total_drill"] > 0:
            pdf.set_text_color(*_DIFF_POS_RGB)
        pdf.cell(cols[7], 5, _diff_text(diff["total_drill"], ",.0f"), align="C", fill=True)
        pdf.set_text_color(60, 60, 60)
        # % G Drill
        pdf.cell(cols[8], 5, f"{ct['g_pct_drill']*100:.1f}%", align="C", fill=True)
        # Incidents (gradient: more is worse)
        if max(inc_vals) > 0:
            r, g, bl = _gradient_color(ct["incident_count"], min(inc_vals), max(inc_vals), False)
            pdf.set_fill_color(r, g, bl)
        else:
            pdf.set_fill_color(255, 255, 255)
        pdf.cell(cols[9], 5, f"{ct['incident_count']}", align="C", fill=True)
        # OP w/ More Runs
        pdf.set_fill_color(255, 255, 255)
        pdf.cell(cols[10], 5, _latin1(ct.get("top_operator", ""))[:18], align="C", fill=True)
        pdf.ln()

    # Grand Total row (no fill, bold, true diffs)
    g_runs = sum(runs_vals)
    g_hrs = round(sum(hrs_vals), 2)
    g_drill = sum(drill_vals)
    g_inc = sum(inc_vals)
    g_prev_hrs = kpi.get("grand_prev_total_hrs", 0.0)
    g_prev_drill = kpi.get("grand_prev_total_drill", 0)
    g_hrs_diff = round(g_hrs - g_prev_hrs, 2)
    g_drill_diff = int(round(g_drill - g_prev_drill))
    g_pct_hrs_sum = sum(b["curr_total"]["g_pct_hrs"] for b in blocks)
    g_pct_drill_sum = sum(b["curr_total"]["g_pct_drill"] for b in blocks)

    pdf.set_x((pdf.w - sum(cols)) / 2)
    pdf.set_fill_color(255, 255, 255)
    pdf.set_text_color(30, 30, 30)
    pdf.set_font("Helvetica", "B", 7)
    pdf.cell(cols[0] + cols[1], 5, "Grand Total", align="C", fill=True)
    pdf.cell(cols[2], 5, f"{g_runs}", align="C", fill=True)
    pdf.cell(cols[3], 5, f"{g_hrs:,.1f}", align="C", fill=True)
    if g_hrs_diff < 0:
        pdf.set_text_color(*_DIFF_NEG_RGB)
    elif g_hrs_diff > 0:
        pdf.set_text_color(*_DIFF_POS_RGB)
    pdf.cell(cols[4], 5, _diff_text(g_hrs_diff, ",.1f"), align="C", fill=True)
    pdf.set_text_color(30, 30, 30)
    pdf.cell(cols[5], 5, f"{g_pct_hrs_sum*100:.1f}%", align="C", fill=True)
    pdf.cell(cols[6], 5, f"{g_drill:,.0f}", align="C", fill=True)
    if g_drill_diff < 0:
        pdf.set_text_color(*_DIFF_NEG_RGB)
    elif g_drill_diff > 0:
        pdf.set_text_color(*_DIFF_POS_RGB)
    pdf.cell(cols[7], 5, _diff_text(g_drill_diff, ",.0f"), align="C", fill=True)
    pdf.set_text_color(30, 30, 30)
    pdf.cell(cols[8], 5, f"{g_pct_drill_sum*100:.1f}%", align="C", fill=True)
    pdf.cell(cols[9], 5, f"{g_inc}", align="C", fill=True)
    pdf.cell(cols[10], 5, _latin1(kpi.get("grand_top_operator", ""))[:18], align="C", fill=True)
    pdf.ln()
    pdf.set_font("Helvetica", "", 7)


def _kpi_detailed_table(pdf, kpi):
    """Detailed KPI table — per Motor Type / Job Type / Series 20 grouped by hole size."""
    cols = [8, 12, 18, 16, 12, 10, 14, 12, 12, 16, 12, 12, 14, 12, 12, 14, 10, 12, 18]  # 246mm
    headers = ["Wk", "Hole", "Motor Type", "Job Type", "S20",
               "Runs", "Hrs", "% W Hrs", "% G Hrs",
               "Total Drill", "% W Drl", "% G Drl",
               "Drill Hrs", "Avg ROP", "Slide %", "Avg Run Len",
               "MY Avg", "Incid", "OP w/ More Runs"]
    pdf.check_space(40)
    pdf.sub_title("Detailed Table (per Motor Type / Job Type / Series 20)")
    left_x = (pdf.w - sum(cols)) / 2

    pdf.set_x(left_x)
    _table_header(pdf, cols, headers, bg_color=(31, 78, 120))

    week_str = str(kpi.get("week", ""))
    week_num = week_str.split("-W")[-1] if "-W" in week_str else week_str

    pdf.set_font("Helvetica", "", 6)
    for b in kpi.get("blocks", []):
        # Per-hole-size winner (group with most TOTAL_DRILL)
        winner_idx = -1
        if b["rows"]:
            winner_idx = max(range(len(b["rows"])),
                             key=lambda i: b["rows"][i]["total_drill"])
        for ri, r in enumerate(b["rows"]):
            mkey = _motor_key(r.get("motor_type"))
            fill = _MOTOR_RGB.get(mkey, (255, 255, 255))
            pdf.set_fill_color(*fill)
            pdf.set_text_color(60, 60, 60)
            pdf.set_x(left_x)
            pdf.cell(cols[0], 4, str(week_num), align="C", fill=True)
            pdf.cell(cols[1], 4, f'{b["hole_size"]:g}"', align="C", fill=True)
            pdf.cell(cols[2], 4, _latin1(str(r.get("motor_type", "")))[:12], align="C", fill=True)
            pdf.cell(cols[3], 4, _latin1(str(r.get("job_type", "")))[:11], align="C", fill=True)
            pdf.cell(cols[4], 4, str(r.get("series_20", "")), align="C", fill=True)
            pdf.cell(cols[5], 4, f"{r['runs']}", align="C", fill=True)
            pdf.cell(cols[6], 4, f"{r['total_hrs']:,.1f}", align="C", fill=True)
            pdf.cell(cols[7], 4, f"{r['w_pct_hrs']*100:.1f}%", align="C", fill=True)
            pdf.cell(cols[8], 4, f"{r['g_pct_hrs']*100:.1f}%", align="C", fill=True)
            # Total Drill — darker fill if this is the per-hole-size winner
            if ri == winner_idx and mkey:
                # 70% darker shade of motor-type fill
                dark = tuple(max(0, int(c * 0.55)) for c in fill)
                pdf.set_fill_color(*dark)
                pdf.set_text_color(255, 255, 255)
                pdf.set_font("Helvetica", "B", 6)
                pdf.cell(cols[9], 4, f"{r['total_drill']:,.0f}", align="C", fill=True)
                pdf.set_font("Helvetica", "", 6)
                pdf.set_fill_color(*fill)
                pdf.set_text_color(60, 60, 60)
            else:
                pdf.cell(cols[9], 4, f"{r['total_drill']:,.0f}", align="C", fill=True)
            pdf.cell(cols[10], 4, f"{r['w_pct_drill']*100:.1f}%", align="C", fill=True)
            pdf.cell(cols[11], 4, f"{r['g_pct_drill']*100:.1f}%", align="C", fill=True)
            pdf.cell(cols[12], 4, f"{r['drilling_hrs']:,.1f}", align="C", fill=True)
            pdf.cell(cols[13], 4, f"{r['avg_rop']:,.1f}", align="C", fill=True)
            pdf.cell(cols[14], 4, f"{r['avg_slide_pct']*100:.1f}%", align="C", fill=True)
            pdf.cell(cols[15], 4, f"{r['avg_run_length']:,.0f}", align="C", fill=True)
            my = r.get("my_avg")
            pdf.cell(cols[16], 4, f"{my:.1f}" if my is not None else "", align="C", fill=True)
            inc = r.get("incident_count", 0)
            pdf.cell(cols[17], 4, f"{inc}" if inc else "", align="C", fill=True)
            pdf.cell(cols[18], 4, _latin1(r.get("top_operator", ""))[:14], align="C", fill=True)
            pdf.ln()
        # Slim spacer line between hole sizes
        pdf.ln(0.5)
    pdf.set_font("Helvetica", "", 7)


def _kpi_longest_run_table(pdf, kpi):
    """Longest Run table — one row per hole size, sorted by Total Drill descending."""
    cols = [8, 12, 18, 16, 10, 14, 18, 14, 12, 12, 10, 10, 22, 16, 22, 12]  # 226mm
    headers = ["Wk", "Hole", "Motor Type", "Job Type", "S20",
               "Total Hrs", "Total Drill", "Drill Hrs",
               "Avg ROP", "Slide %", "MY Avg", "Incid",
               "Operator", "Job Number", "Phase", "Bend"]
    pdf.check_space(35)
    pdf.sub_title("Longest Run Table (sorted by Total Drill)")
    left_x = (pdf.w - sum(cols)) / 2

    pdf.set_x(left_x)
    _table_header(pdf, cols, headers, bg_color=(31, 78, 120))

    week_str = str(kpi.get("week", ""))
    week_num = week_str.split("-W")[-1] if "-W" in week_str else week_str

    # Collect entries from all blocks, sort by Total Drill descending
    lr_entries = []
    for b in kpi.get("blocks", []):
        lr = b.get("longest_run")
        if lr:
            lr_entries.append((b["hole_size"], lr))
    lr_entries.sort(key=lambda x: -x[1].get("total_drill", 0))

    pdf.set_font("Helvetica", "", 6)
    for hs, lr in lr_entries:
        mkey = _motor_key(lr.get("motor_type"))
        fill = _MOTOR_RGB.get(mkey, (255, 255, 255))
        pdf.set_fill_color(*fill)
        pdf.set_text_color(60, 60, 60)
        pdf.set_x(left_x)
        pdf.cell(cols[0], 4, str(week_num), align="C", fill=True)
        pdf.cell(cols[1], 4, f'{hs:g}"', align="C", fill=True)
        pdf.cell(cols[2], 4, _latin1(str(lr.get("motor_type", "")))[:12], align="C", fill=True)
        pdf.cell(cols[3], 4, _latin1(str(lr.get("job_type", "")))[:11], align="C", fill=True)
        pdf.cell(cols[4], 4, str(lr.get("series_20", "")), align="C", fill=True)
        pdf.cell(cols[5], 4, f"{lr['total_hrs']:,.1f}", align="C", fill=True)
        pdf.cell(cols[6], 4, f"{lr['total_drill']:,.0f}", align="C", fill=True)
        pdf.cell(cols[7], 4, f"{lr['drilling_hrs']:,.1f}", align="C", fill=True)
        pdf.cell(cols[8], 4, f"{lr['avg_rop']:,.1f}", align="C", fill=True)
        pdf.cell(cols[9], 4, f"{lr['avg_slide_pct']*100:.1f}%", align="C", fill=True)
        my = lr.get("my_avg")
        pdf.cell(cols[10], 4, f"{my:.1f}" if my is not None else "", align="C", fill=True)
        inc = lr.get("incident_count", 0)
        pdf.cell(cols[11], 4, f"{inc}" if inc else "", align="C", fill=True)
        pdf.cell(cols[12], 4, _latin1(str(lr.get("operator", "")))[:14], align="C", fill=True)
        pdf.cell(cols[13], 4, _latin1(str(lr.get("job_num", ""))), align="C", fill=True)
        pdf.cell(cols[14], 4, _latin1(str(lr.get("phase", "")))[:14], align="C", fill=True)
        pdf.cell(cols[15], 4, _latin1(str(lr.get("bend", ""))), align="C", fill=True)
        pdf.ln()
    pdf.set_font("Helvetica", "", 7)


_CHART_ASPECT = 0.462  # rendered image height/width ratio for matplotlib (11x5) figures


def _render_cat1_trend_chart(pdf, section, metric, chart_type, label, compact=False):
    """Embed one trend chart in the PDF.
    compact=True uses a smaller image (~180mm wide) so two charts fit on one page."""
    if not section or not section.get("hole_sizes"):
        return
    pdf.section_title(label)

    try:
        from src.trends import render_trend_chart
        png_path = render_trend_chart(section, metric=metric, chart_type=chart_type)
    except Exception as e:
        pdf.body_text(f"Chart rendering failed: {e}")
        return

    if not png_path or not os.path.exists(png_path):
        pdf.body_text("Chart rendering produced no output.")
        return

    page_w = pdf.w
    img_w = 180 if compact else (page_w - 20)
    x = (page_w - img_w) / 2  # center horizontally
    pdf.image(png_path, x=x, y=pdf.get_y(), w=img_w)
    pdf.set_y(pdf.get_y() + img_w * _CHART_ASPECT)
    pdf.ln(2)

    try:
        os.remove(png_path)
    except OSError:
        pass


def _render_cat1_trends(pdf, section):
    """1D / 1E: 12-week stacked-area trends — Total Drill and Number of Runs."""
    if not section or not section.get("hole_sizes"):
        return
    n_weeks = len(section["weeks"])
    _render_cat1_trend_chart(
        pdf, section, metric="drill", chart_type="stacked",
        label=f"1D. Trends -- Total Drill by Hole Size ({n_weeks}-Week Stacked Area)",
    )
    _render_cat1_trend_chart(
        pdf, section, metric="runs", chart_type="stacked",
        label=f"1E. Trends -- Number of Runs by Hole Size ({n_weeks}-Week Stacked Area)",
    )


def _render_cat1_summary(pdf, section, trends_section=None):
    """1A: Weekly Summary — three KPI tables, plus 12-week trend charts immediately after."""
    pdf.section_title("1A. Weekly Summary")
    kpi = section.get("kpi")
    if not kpi or not kpi.get("blocks"):
        # Fallback: legacy mini-summary if KPI data is missing
        ct = section["current_totals"]
        pt = section["previous_totals"]
        dt = section["delta_totals"]
        cols = [55, 30, 30, 30]
        _table_header(pdf, cols, ["Metric", "Current Week", "Previous Week", "Delta"])
        for label, ck, fmt in [("Runs", "runs", "d"),
                               ("Total Footage (ft)", "total_drill", ",.0f"),
                               ("Total Hours", "total_hrs", ",.1f")]:
            _table_row_color(pdf, label == "Total Footage (ft)")
            pdf.cell(cols[0], 5, f"  {label}", fill=True)
            pdf.cell(cols[1], 5, f"{ct[ck]:{fmt}}", align="C", fill=True)
            pdf.cell(cols[2], 5, f"{pt[ck]:{fmt}}", align="C", fill=True)
            pdf.set_font("Helvetica", "B", 7)
            pdf.cell(cols[3], 5, _delta_str(dt[ck], fmt=".0f" if "d" in fmt or ",.0f" in fmt else ".1f"),
                     align="C", fill=True)
            pdf.set_font("Helvetica", "", 7)
            pdf.ln()
        pdf.ln(4)
        return

    # Page 1 of Cat 1: Summary table + Longest Run table.
    # (Detailed table is rendered later, on its own page — see _render_cat1_detailed.)
    _kpi_summary_table(pdf, kpi)
    pdf.ln(4)
    _kpi_longest_run_table(pdf, kpi)
    pdf.ln(4)

    # Page 2 of Cat 1: trend charts (forced new page so both fit together).
    if trends_section and trends_section.get("hole_sizes"):
        pdf.add_page()
        n_weeks = len(trends_section["weeks"])
        _render_cat1_trend_chart(
            pdf, trends_section, metric="drill", chart_type="stacked",
            label=f"Trends -- Total Drill by Hole Size ({n_weeks}-Week Stacked Area)",
            compact=True,
        )
        _render_cat1_trend_chart(
            pdf, trends_section, metric="runs", chart_type="stacked",
            label=f"Trends -- Number of Runs by Hole Size ({n_weeks}-Week Stacked Area)",
            compact=True,
        )


def _render_cat1_detailed(pdf, section):
    """Page 3 of Cat 1: the Detailed Table on a fresh page.
    Curves Analysis (1B) flows below if there is room."""
    pdf.add_page()
    kpi = section.get("kpi")
    if not kpi or not kpi.get("blocks"):
        return
    _kpi_detailed_table(pdf, kpi)
    pdf.ln(4)


def _render_cat1_curves(pdf, section):
    """1B: Curves Analysis. Flows on the same page as the Detailed table when possible."""
    pdf.check_space(20)
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
    """1C: Reason to POOH — 12-week stacked-bar trend + Motor Issues table.
    Forces a fresh page so the chart and table fit together."""
    pdf.add_page()
    pdf.section_title(f"1C. Reason to POOH ({section['reason_col_used']})")

    cur = section["current"]
    prv = section["previous"]
    pdf.body_text(f"Filtered runs: {cur['total_filtered']} (current) | {prv['total_filtered']} (previous)")
    pdf.ln(2)

    # 12-week POOH 100% stacked-bar chart (smaller embed so the Motor Issues table fits below).
    pooh_trend = section.get("pooh_trend")
    if pooh_trend and pooh_trend.get("weeks"):
        try:
            from src.trends import render_pooh_trend_chart
            png_path = render_pooh_trend_chart(pooh_trend)
            if png_path and os.path.exists(png_path):
                page_w = pdf.w
                img_w = 200  # narrower so the Motor Issues table fits on the same page
                x = (page_w - img_w) / 2
                pdf.image(png_path, x=x, y=pdf.get_y(), w=img_w)
                pdf.set_y(pdf.get_y() + img_w * _CHART_ASPECT)
                pdf.ln(2)
                try:
                    os.remove(png_path)
                except OSError:
                    pass
        except Exception as e:
            pdf.body_text(f"POOH chart rendering failed: {e}")

    # Motor issues detail
    motor_detail = cur["motor_detail"]
    if len(motor_detail) > 0:
        pdf.sub_title("Motor Issues Detail (Current Week)")
        mcols = [55, 20, 30, 45]
        _table_header(pdf, mcols, ["Operator", "Hole", "SN", "Reason"], bg_color=(180, 50, 30))
        alt = False
        for _, row in motor_detail.iterrows():
            pdf.check_space(8)
            _table_row_color(pdf, alt, flagged=True, flag_color=(255, 245, 243))
            pdf.cell(mcols[0], 5, f"  {_latin1(str(row['operator']))[:30]}", fill=True)
            pdf.cell(mcols[1], 5, _fmt_hole(row.get('hole_size')), align="C", fill=True)
            pdf.cell(mcols[2], 5, f"{_safe(row.get('sn', 'N/A'))[:18]}", align="C", fill=True)
            pdf.set_text_color(180, 40, 20)
            pdf.cell(mcols[3], 5, f"  {_latin1(str(row.get('reason', 'N/A')))[:28]}", fill=True)
            pdf.set_text_color(60, 60, 60)
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
        cols = [8, 40, 46, 24, 18, 16, 20, 14, 16, 22]
        headers = ["#", "Operator", "Well", "Footage", "ROP", "Hrs", "Type", "Hole", "Bend", "Basin"]
        _table_header(pdf, cols, headers)
        alt = False
        for _, run in df.iterrows():
            pdf.check_space(8)
            _table_row_color(pdf, alt)
            pdf.cell(cols[0], 5, f"{int(run['rank'])}", align="C", fill=True)
            pdf.cell(cols[1], 5, f"  {_latin1(str(run['operator']))[:23]}", fill=True)
            pdf.cell(cols[2], 5, f"  {_latin1(str(run['well']))[:27]}", fill=True)
            pdf.set_font("Helvetica", "B", 7)
            pdf.cell(cols[3], 5, f"{run['total_drill']:,.0f}", align="C", fill=True)
            pdf.set_font("Helvetica", "", 7)
            rop = f"{run['avg_rop']:.1f}" if run['avg_rop'] else "N/A"
            hrs = f"{run['drilling_hours']:.1f}" if run['drilling_hours'] else "N/A"
            pdf.cell(cols[4], 5, rop, align="C", fill=True)
            pdf.cell(cols[5], 5, hrs, align="C", fill=True)
            pdf.cell(cols[6], 5, _safe(run.get('motor_type2'))[:12], align="C", fill=True)
            pdf.cell(cols[7], 5, _fmt_hole(run.get('hole_size')), align="C", fill=True)
            pdf.cell(cols[8], 5, _safe(run.get('bend_hsg', 'N/A'))[:10], align="C", fill=True)
            pdf.cell(cols[9], 5, _safe(run.get('basin'))[:12], align="C", fill=True)
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
# Category 4 rendering (QC Audit — Friday only)
# =========================================================================

def _render_cat4_column_summary(pdf, section_a):
    """4A: Column Change Summary."""
    pdf.section_title("4A. Column Change Summary (Most Corrected Columns)")
    if len(section_a) == 0:
        pdf.body_text("No column changes detected.")
        return

    cols = [80, 30, 30]
    _table_header(pdf, cols, ["Column", "Changes", "% of Rows"])
    alt = False
    for _, row in section_a.head(20).iterrows():
        pdf.check_space(7)
        _table_row_color(pdf, alt)
        pdf.cell(cols[0], 5, f"  {_latin1(str(row['column']))[:45]}", fill=True)
        pdf.cell(cols[1], 5, f"{int(row['changes'])}", align="C", fill=True)
        pdf.cell(cols[2], 5, f"{row['pct']:.1f}%", align="C", fill=True)
        pdf.ln()
        alt = not alt
    pdf.ln(4)


def _render_cat4_reviewer_workload(pdf, section_b):
    """4B: QC Reviewer Workload."""
    pdf.check_space(30)
    pdf.section_title("4B. QC Reviewer Workload (RHC vs YGG)")
    if len(section_b) == 0:
        pdf.body_text("No reviewer data available.")
        return

    cols = [35, 25, 25, 30, 25]
    _table_header(pdf, cols, ["Reviewer", "Rows", "Changed", "Cell Edits", "Avg/Row"])
    alt = False
    for _, row in section_b.iterrows():
        _table_row_color(pdf, alt)
        pdf.cell(cols[0], 5, f"  {_latin1(str(row['reviewer']))}", fill=True)
        pdf.cell(cols[1], 5, f"{int(row['rows_assigned'])}", align="C", fill=True)
        pdf.cell(cols[2], 5, f"{int(row['rows_changed'])}", align="C", fill=True)
        pdf.set_font("Helvetica", "B", 7)
        pdf.cell(cols[3], 5, f"{int(row['cell_changes'])}", align="C", fill=True)
        pdf.set_font("Helvetica", "", 7)
        pdf.cell(cols[4], 5, f"{row['avg_per_row']:.1f}", align="C", fill=True)
        pdf.ln()
        alt = not alt
    pdf.ln(4)


def _render_cat4_operator_trends(pdf, section_c):
    """4C: Operator QC Trends."""
    pdf.check_space(30)
    pdf.section_title("4C. Operator QC Trends")

    # Only show operators with changes
    has_changes = section_c[section_c["cell_changes"] > 0] if len(section_c) > 0 else section_c
    if len(has_changes) == 0:
        pdf.body_text("No operator corrections detected.")
        return

    cols = [50, 18, 22, 22, 60]
    _table_header(pdf, cols, ["Operator", "Rows", "Changed", "Edits", "Top Columns"])
    alt = False
    for _, row in has_changes.head(15).iterrows():
        pdf.check_space(7)
        _table_row_color(pdf, alt)
        pdf.cell(cols[0], 5, f"  {_latin1(str(row['operator']))[:28]}", fill=True)
        pdf.cell(cols[1], 5, f"{int(row['rows'])}", align="C", fill=True)
        pdf.cell(cols[2], 5, f"{int(row['changed_rows'])}", align="C", fill=True)
        pdf.cell(cols[3], 5, f"{int(row['cell_changes'])}", align="C", fill=True)
        pdf.set_font("Helvetica", "I", 6)
        pdf.cell(cols[4], 5, f"  {_latin1(str(row['top_columns']))[:38]}", fill=True)
        pdf.set_font("Helvetica", "", 7)
        pdf.ln()
        alt = not alt
    pdf.ln(4)


def _render_cat4_patterns(pdf, section_d):
    """4D: Auto-Detected Patterns."""
    pdf.check_space(30)
    pdf.section_title("4D. Auto-Detected Patterns")

    new_rows = section_d.get("new_rows", 0)
    removed_rows = section_d.get("removed_rows", 0)
    if new_rows > 0 or removed_rows > 0:
        pdf.body_text(f"Row changes: {new_rows} new rows added, {removed_rows} rows removed during QC")
        pdf.ln(2)

    broken = section_d.get("broken_columns", [])
    if broken:
        pdf.sub_title("Columns Needing Source Fix (>50% corrected)")
        for item in broken:
            pdf.set_font("Helvetica", "B", 8)
            pdf.set_text_color(200, 40, 20)
            pdf.cell(0, 5, f"  {_latin1(item['column'])}  —  {item['rows_changed']} rows ({item['pct']:.1f}%)", new_x="LMARGIN", new_y="NEXT")
            pdf.set_font("Helvetica", "", 7)
            pdf.set_text_color(60, 60, 60)
        pdf.ln(3)

    systematic = section_d.get("systematic", [])
    if systematic:
        pdf.sub_title("Systematic Corrections (same operator+column, 3+ times)")
        scols = [60, 50, 25]
        _table_header(pdf, scols, ["Operator", "Column", "Count"], bg_color=(180, 50, 30))
        alt = False
        for item in systematic[:10]:
            _table_row_color(pdf, alt, flagged=True, flag_color=(255, 245, 243))
            pdf.cell(scols[0], 5, f"  {_latin1(str(item['operator']))[:35]}", fill=True)
            pdf.cell(scols[1], 5, f"  {_latin1(str(item['column']))[:30]}", fill=True)
            pdf.set_text_color(180, 40, 20)
            pdf.cell(scols[2], 5, f"{item['count']}", align="C", fill=True)
            pdf.set_text_color(60, 60, 60)
            pdf.ln()
            alt = not alt
        pdf.ln(3)

    high_effort = section_d.get("high_effort_rows", [])
    if high_effort:
        pdf.sub_title("High-Effort QC Rows (most cell changes)")
        hcols = [55, 55, 25, 20]
        _table_header(pdf, hcols, ["Operator", "Well", "Changes", "QC By"])
        alt = False
        for item in high_effort:
            _table_row_color(pdf, alt)
            pdf.cell(hcols[0], 5, f"  {_latin1(str(item['operator']))[:30]}", fill=True)
            pdf.cell(hcols[1], 5, f"  {_latin1(str(item['well']))[:30]}", fill=True)
            pdf.set_font("Helvetica", "B", 7)
            pdf.cell(hcols[2], 5, f"{item['changes']}", align="C", fill=True)
            pdf.set_font("Helvetica", "", 7)
            pdf.cell(hcols[3], 5, f"{_latin1(str(item['qc_by']))}", align="C", fill=True)
            pdf.ln()
            alt = not alt
        pdf.ln(3)

    if not broken and not systematic and not high_effort and new_rows == 0 and removed_rows == 0:
        pdf.body_text("No significant patterns detected.")
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

    # =====================================================================
    # CATEGORY 1: Week vs Previous Week
    # =====================================================================
    cat1 = all_results.get("category1")
    if cat1:
        pdf.category_header(1, cat1["category"])
        sections = cat1["sections"]
        # Page 1: Summary + Longest Run | Page 2: Trends
        _render_cat1_summary(pdf, sections["A_weekly_summary"],
                             trends_section=sections.get("D_trends"))
        # Page 3: Detailed table + 1B Curves Analysis
        _render_cat1_detailed(pdf, sections["A_weekly_summary"])
        _render_cat1_curves(pdf, sections["B_curves"])
        # Page 4: 1C Reason to POOH + Motor Issues (forces page break internally)
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
    cat3 = all_results.get("category3")
    if cat3:
        pdf.add_page()
        pdf.category_header(3, cat3["category"])
        sections = cat3["sections"]
        _render_cat3_avg_rop(pdf, sections.get("A_avg_rop"))
        _render_cat3_longest_runs(pdf, sections.get("B_longest_runs"))
        _render_cat3_sliding_pct(pdf, sections.get("C_sliding_pct"))
        _render_cat3_patterns(pdf, sections.get("D_pattern_highlights"))

    # =====================================================================
    # CATEGORY 4: QC Audit (Friday only)
    # =====================================================================
    cat4 = all_results.get("category4")
    if cat4:
        pdf.add_page()
        pdf.category_header(4, "QC AUDIT (Wed vs Fri Comparison)")
        meta4 = cat4["meta"]
        pdf.body_text(f"Wednesday rows: {meta4['wed_rows']} | Friday rows: {meta4['fri_rows']} | Matched: {meta4['matched']}")
        pdf.body_text(f"Total cell changes: {meta4['total_changes']} across {meta4['columns_compared']} columns compared")
        pdf.ln(3)
        sections = cat4["sections"]
        _render_cat4_column_summary(pdf, sections["A_column_summary"])
        _render_cat4_reviewer_workload(pdf, sections["B_reviewer_workload"])
        _render_cat4_operator_trends(pdf, sections["C_operator_trends"])
        _render_cat4_patterns(pdf, sections["D_patterns"])

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
