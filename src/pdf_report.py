"""
PDF Report Generator Module
Generates a professional PDF report with weekly performance highlights.
Shows all 5 variable groups: Location, Section, Motor, Bit, Depth.
"""

import os
from datetime import datetime
from fpdf import FPDF
import pandas as pd


class PerformanceReport(FPDF):
    """Custom PDF class with header/footer."""

    def __init__(self, week, date_start, date_end):
        super().__init__(orientation="L", format="letter")
        self.week = week
        self.date_start = date_start
        self.date_end = date_end
        self.set_auto_page_break(auto=True, margin=20)

    def header(self):
        self.set_font("Helvetica", "B", 10)
        self.set_text_color(100, 100, 100)
        self.cell(0, 6, f"Scout Downhole - Weekly Performance Report | Week {self.week}", align="L")
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

    def section_title(self, title):
        self.set_font("Helvetica", "B", 13)
        self.set_text_color(200, 60, 30)
        self.cell(0, 10, title, new_x="LMARGIN", new_y="NEXT")
        self.set_draw_color(200, 60, 30)
        self.set_line_width(0.4)
        self.line(10, self.get_y(), self.w - 10, self.get_y())
        self.ln(4)

    def sub_title(self, title):
        self.set_font("Helvetica", "B", 10)
        self.set_text_color(50, 50, 50)
        self.cell(0, 7, title, new_x="LMARGIN", new_y="NEXT")
        self.ln(1)

    def body_text(self, text):
        self.set_font("Helvetica", "", 9)
        self.set_text_color(60, 60, 60)
        self.cell(0, 5, text, new_x="LMARGIN", new_y="NEXT")

    def stat_line(self, label, value, color=None):
        self.set_font("Helvetica", "", 9)
        self.set_text_color(80, 80, 80)
        self.cell(55, 5, f"  {label}:", align="R")
        self.set_font("Helvetica", "B", 9)
        if color:
            self.set_text_color(*color)
        else:
            self.set_text_color(30, 30, 30)
        self.cell(0, 5, f"  {value}", new_x="LMARGIN", new_y="NEXT")

    def kpi_card(self, label, value, unit="", color=(30, 30, 30)):
        """Draw a KPI summary card."""
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


def _safe(val, default="N/A"):
    """Safe string conversion for NaN values."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return default
    s = str(val)
    return default if s.lower() in ("nan", "none", "") else s


def _fmt_hole(val):
    """Format hole size for display."""
    s = _safe(val)
    if s == "N/A":
        return s
    try:
        return f'{float(s):g}"'
    except (ValueError, TypeError):
        return s


def generate_pdf(week, date_start, date_end, new_runs_count, kpi_results, output_dir=None):
    """Generate the PDF performance report."""

    pdf = PerformanceReport(week, date_start, date_end)
    pdf.alias_nb_pages()
    pdf.add_page()

    # =====================================================================
    # TITLE PAGE / SUMMARY
    # =====================================================================
    pdf.set_font("Helvetica", "B", 22)
    pdf.set_text_color(30, 30, 30)
    pdf.cell(0, 14, "Weekly Performance Report", new_x="LMARGIN", new_y="NEXT")

    pdf.set_font("Helvetica", "", 12)
    pdf.set_text_color(100, 100, 100)
    pdf.cell(0, 8, f"Week {week}  |  {date_start.date()} to {date_end.date()}", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(6)

    # KPI Cards row
    rop_summary = kpi_results.get("avg_rop")
    drill_summary = kpi_results.get("longest_runs")
    slide_summary = kpi_results.get("sliding_pct")

    pdf.kpi_card("Total New Runs", str(new_runs_count))
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
    # AVG ROP SECTION
    # =====================================================================
    if rop_summary:
        pdf.section_title("AVG ROP Analysis")

        pdf.stat_line("Runs with ROP data", str(rop_summary['total_runs']))
        pdf.stat_line("Week Average", f"{rop_summary['week_avg']} ft/hr")
        pdf.stat_line("Week Median", f"{rop_summary['week_median']} ft/hr")
        pdf.ln(4)

        # Basin breakdown table
        if "basin_breakdown" in rop_summary:
            pdf.sub_title("ROP by Basin")
            pdf.set_font("Helvetica", "B", 8)
            pdf.set_fill_color(50, 50, 60)
            pdf.set_text_color(255, 255, 255)
            pdf.cell(60, 6, "  Basin", fill=True)
            pdf.cell(35, 6, "Avg ROP (ft/hr)", align="C", fill=True)
            pdf.cell(25, 6, "Runs", align="C", fill=True)
            pdf.ln()

            pdf.set_font("Helvetica", "", 8)
            alt = False
            for basin, row in rop_summary["basin_breakdown"].iterrows():
                if alt:
                    pdf.set_fill_color(245, 245, 248)
                else:
                    pdf.set_fill_color(255, 255, 255)
                pdf.set_text_color(60, 60, 60)
                pdf.cell(60, 5, f"  {basin}", fill=True)
                pdf.cell(35, 5, f"{row['mean']:.1f}", align="C", fill=True)
                pdf.cell(25, 5, f"{int(row['count'])}", align="C", fill=True)
                pdf.ln()
                alt = not alt
            pdf.ln(4)

        # Underperforming runs table (with new columns)
        results = rop_summary["results"]
        flagged_below = results[results["flag"] == "below"].sort_values("diff_pct")
        if len(flagged_below) > 0:
            pdf.sub_title(f"Underperforming Runs ({len(flagged_below)} flagged)")

            pdf.set_font("Helvetica", "B", 6)
            pdf.set_fill_color(200, 60, 30)
            pdf.set_text_color(255, 255, 255)
            col_widths = [38, 48, 16, 22, 18, 24, 22, 20, 18, 24, 20]
            headers = ["Operator", "Well", "Hole", "ROP", "Diff%", "Basin", "County", "Motor", "L/S", "Formation", "Match"]
            for i, h in enumerate(headers):
                pdf.cell(col_widths[i], 6, h, align="C", fill=True)
            pdf.ln()

            pdf.set_font("Helvetica", "", 6)
            alt = False
            for _, run in flagged_below.head(12).iterrows():
                if alt:
                    pdf.set_fill_color(255, 240, 238)
                else:
                    pdf.set_fill_color(255, 255, 255)
                pdf.set_text_color(60, 60, 60)
                pdf.cell(col_widths[0], 5, f"  {str(run['operator'])[:22]}", fill=True)
                pdf.cell(col_widths[1], 5, f"  {str(run['well'])[:28]}", fill=True)
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
                pdf.set_font("Helvetica", "", 6)
                pdf.ln()
                alt = not alt
            pdf.ln(4)

        # Top performers table
        flagged_above = results[results["flag"] == "above"].sort_values("diff_pct", ascending=False)
        if len(flagged_above) > 0:
            pdf.sub_title(f"Top Performers ({len(flagged_above)} highlighted)")

            pdf.set_font("Helvetica", "B", 6)
            pdf.set_fill_color(30, 140, 60)
            pdf.set_text_color(255, 255, 255)
            for i, h in enumerate(headers):
                pdf.cell(col_widths[i], 6, h, align="C", fill=True)
            pdf.ln()

            pdf.set_font("Helvetica", "", 6)
            alt = False
            for _, run in flagged_above.head(8).iterrows():
                if alt:
                    pdf.set_fill_color(235, 255, 240)
                else:
                    pdf.set_fill_color(255, 255, 255)
                pdf.set_text_color(60, 60, 60)
                pdf.cell(col_widths[0], 5, f"  {str(run['operator'])[:22]}", fill=True)
                pdf.cell(col_widths[1], 5, f"  {str(run['well'])[:28]}", fill=True)
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
                pdf.set_font("Helvetica", "", 6)
                pdf.ln()
                alt = not alt
            pdf.ln(4)

    # =====================================================================
    # LONGEST RUNS SECTION
    # =====================================================================
    if drill_summary:
        pdf.add_page()
        pdf.section_title(f"Longest Runs - Top {drill_summary['top_n']}")

        pdf.stat_line("Runs with drill data", str(drill_summary['total_runs']))
        pdf.stat_line("Week Average", f"{drill_summary['week_avg']:,.0f} ft")
        pdf.stat_line("Week Max", f"{drill_summary['week_max']:,.0f} ft")
        pdf.ln(4)

        drill_results = drill_summary["results"]
        if len(drill_results) > 0:
            pdf.set_font("Helvetica", "B", 6)
            pdf.set_fill_color(50, 50, 60)
            pdf.set_text_color(255, 255, 255)
            dcols = [8, 38, 48, 26, 20, 20, 16, 24, 22, 24, 24]
            dheaders = ["#", "Operator", "Well", "Footage", "Hours", "ROP", "Hole", "Basin", "County", "Phase", "Motor"]
            for i, h in enumerate(dheaders):
                pdf.cell(dcols[i], 6, h, align="C", fill=True)
            pdf.ln()

            pdf.set_font("Helvetica", "", 6)
            alt = False
            for _, run in drill_results.iterrows():
                if alt:
                    pdf.set_fill_color(245, 245, 248)
                else:
                    pdf.set_fill_color(255, 255, 255)
                pdf.set_text_color(60, 60, 60)
                pdf.cell(dcols[0], 5, f"{int(run['rank'])}", align="C", fill=True)
                pdf.cell(dcols[1], 5, f"  {str(run['operator'])[:22]}", fill=True)
                pdf.cell(dcols[2], 5, f"  {str(run['well'])[:28]}", fill=True)
                pdf.set_font("Helvetica", "B", 6)
                pdf.cell(dcols[3], 5, f"{run['value']:,.0f}", align="C", fill=True)
                pdf.set_font("Helvetica", "", 6)
                hrs = f"{run['drilling_hours']:.1f}" if run['drilling_hours'] and not pd.isna(run['drilling_hours']) else "N/A"
                rop = f"{run['avg_rop']:.1f}" if run['avg_rop'] and not pd.isna(run['avg_rop']) else "N/A"
                pdf.cell(dcols[4], 5, hrs, align="C", fill=True)
                pdf.cell(dcols[5], 5, rop, align="C", fill=True)
                pdf.cell(dcols[6], 5, _fmt_hole(run['hole_size']), align="C", fill=True)
                pdf.cell(dcols[7], 5, _safe(run['basin'])[:14], align="C", fill=True)
                pdf.cell(dcols[8], 5, _safe(run['county'])[:12], align="C", fill=True)
                pdf.cell(dcols[9], 5, _safe(run['phase']), align="C", fill=True)
                pdf.cell(dcols[10], 5, _safe(run['motor_model']), align="C", fill=True)
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
                    pdf.set_font("Helvetica", "", 6)

                alt = not alt
            pdf.ln(4)

    # =====================================================================
    # SLIDING % SECTION
    # =====================================================================
    if slide_summary and slide_summary["total_lat_runs"] > 0:
        pdf.section_title("Sliding % (LAT Phase Only)")

        pdf.stat_line("LAT Runs", f"{slide_summary['total_lat_runs']} of {slide_summary['total_new_runs']} total")
        pdf.stat_line("Week Average", f"{slide_summary['week_avg']}%")
        pdf.stat_line("Week Median", f"{slide_summary['week_median']}%")
        pdf.stat_line("Range", f"{slide_summary['week_min']}% to {slide_summary['week_max']}%")
        if slide_summary["hist_threshold_pct"]:
            pdf.stat_line("75th Pct Threshold", f"{slide_summary['hist_threshold_pct']}%")
        pdf.ln(4)

        slide_results = slide_summary["results"]
        if len(slide_results) > 0:
            sorted_slides = slide_results.sort_values("value", ascending=False)

            pdf.set_font("Helvetica", "B", 6)
            pdf.set_fill_color(50, 50, 60)
            pdf.set_text_color(255, 255, 255)
            scols = [38, 50, 16, 22, 26, 26, 24, 22, 22, 24]
            sheaders = ["Operator", "Well", "Hole", "Slide%", "Slide(ft)", "Total(ft)", "Basin", "County", "Motor", "Match"]
            for i, h in enumerate(sheaders):
                pdf.cell(scols[i], 6, h, align="C", fill=True)
            pdf.ln()

            pdf.set_font("Helvetica", "", 6)
            alt = False
            for _, run in sorted_slides.iterrows():
                is_flagged = run["flag"] == "above"
                if is_flagged:
                    pdf.set_fill_color(255, 240, 238)
                elif alt:
                    pdf.set_fill_color(245, 245, 248)
                else:
                    pdf.set_fill_color(255, 255, 255)
                pdf.set_text_color(60, 60, 60)
                pdf.cell(scols[0], 5, f"  {str(run['operator'])[:22]}", fill=True)
                pdf.cell(scols[1], 5, f"  {str(run['well'])[:28]}", fill=True)
                pdf.cell(scols[2], 5, _fmt_hole(run['hole_size']), align="C", fill=True)
                if is_flagged:
                    pdf.set_text_color(200, 40, 20)
                pdf.set_font("Helvetica", "B", 6)
                pdf.cell(scols[3], 5, f"{run['value']:.1f}%", align="C", fill=True)
                pdf.set_font("Helvetica", "", 6)
                pdf.set_text_color(60, 60, 60)
                pdf.cell(scols[4], 5, f"{run['slide_drilled']:,.0f}", align="C", fill=True)
                pdf.cell(scols[5], 5, f"{run['total_drill']:,.0f}", align="C", fill=True)
                pdf.cell(scols[6], 5, _safe(run['basin'])[:14], align="C", fill=True)
                pdf.cell(scols[7], 5, _safe(run['county'])[:12], align="C", fill=True)
                pdf.cell(scols[8], 5, _safe(run['motor_model'])[:12], align="C", fill=True)
                pdf.set_font("Helvetica", "I", 5)
                pdf.cell(scols[9], 5, _safe(run.get('match_level'))[:14], align="C", fill=True)
                pdf.set_font("Helvetica", "", 6)
                pdf.ln()
                alt = not alt

    # =====================================================================
    # SAVE PDF
    # =====================================================================
    if output_dir is None:
        output_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    filename = f"Performance_Report_{week}_{datetime.now().strftime('%Y%m%d')}.pdf"
    filepath = os.path.join(output_dir, filename)
    pdf.output(filepath)

    return filepath
