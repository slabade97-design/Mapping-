import sys
import subprocess
import threading
import time
import re

# Auto-install required packages into the active Python interpreter if missing
for pkg_name, import_name in [
    ("openpyxl", "openpyxl"),
    ("numpy", "numpy"),
    ("matplotlib", "matplotlib"),
    ("python-docx", "docx"),
    ("customtkinter", "customtkinter")
]:
    try:
        __import__(import_name)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg_name])

import customtkinter as ctk
from tkinter import messagebox, filedialog
import numpy as np
import datetime
import os
import io

# Excel imports
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.styles.colors import Color
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.label import DataLabelList
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.colors import ColorChoice, SchemeColor
from openpyxl.drawing.text import (
    Paragraph as XLParagraph, ParagraphProperties, CharacterProperties,
    Font as DrawingFont, RichTextProperties, ListStyle, RegularTextRun
)
from openpyxl.chart.text import RichText, Text
from openpyxl.chart.title import Title
from openpyxl.chart.shapes import GraphicalProperties

# Matplotlib for rendering Excel-style trend charts into Word
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch

# Word (.docx) imports
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# ==============================================================================
# CORE LOGIC & HELPERS
# ==============================================================================
def calc_mkt(temps):
    temps_k = np.array(temps) + 273.15
    mkt_k = 10000 / (-np.log(np.sum(np.exp(-10000 / temps_k)) / len(temps_k)))
    return mkt_k - 273.15

def parse_flexible_date(date_str):
    date_str = date_str.strip()
    for fmt in ("%d/%m/%y", "%d/%m/%Y", "%d-%m-%y", "%d-%m-%Y"):
        try:
            return datetime.datetime.strptime(date_str, fmt)
        except ValueError:
            continue
    raise ValueError(f"Invalid date format: '{date_str}'. Use DD/MM/YY or DD/MM/YYYY.")

def sanitize_filename(name):
    return re.sub(r'[\\/*?:"<>|]', "_", name).strip()

def make_chart_title(text_str, size=1400, rot=0):
    rpr = CharacterProperties(
        sz=size, b=False, i=False, u="none", strike="noStrike", kern=1200,
        spc=0 if size == 1400 else None, baseline=0,
        solidFill=ColorChoice(schemeClr=SchemeColor(val="tx1", lumMod=65000, lumOff=35000)),
        latin=DrawingFont(typeface="+mn-lt"),
        ea=DrawingFont(typeface="+mn-ea"),
        cs=DrawingFont(typeface="+mn-cs")
    )
    p = XLParagraph(
        pPr=ParagraphProperties(defRPr=rpr),
        r=[RegularTextRun(rPr=CharacterProperties(lang="en-IN"), t=text_str)]
    )
    rt = RichText(
        bodyPr=RichTextProperties(
            rot=rot, spcFirstLastPara=True, vertOverflow="ellipsis",
            vert="horz", wrap="square", anchor="ctr", anchorCtr=True
        ),
        lstStyle=ListStyle(),
        p=[p]
    )
    return Title(
        tx=Text(rich=rt),
        overlay=True,
        spPr=GraphicalProperties(noFill=True, ln=LineProperties(noFill=True))
    )

def make_axis_txpr(rot=-60000000, size=900):
    rpr = CharacterProperties(
        sz=size, b=False, i=False, u="none", strike="noStrike", kern=1200, baseline=0,
        solidFill=ColorChoice(schemeClr=SchemeColor(val="tx1", lumMod=65000, lumOff=35000)),
        latin=DrawingFont(typeface="+mn-lt"),
        ea=DrawingFont(typeface="+mn-ea"),
        cs=DrawingFont(typeface="+mn-cs")
    )
    p = XLParagraph(pPr=ParagraphProperties(defRPr=rpr), endParaRPr=CharacterProperties(lang="en-US"))
    return RichText(
        bodyPr=RichTextProperties(
            rot=rot, spcFirstLastPara=True, vertOverflow="ellipsis",
            vert="horz", wrap="square", anchor="ctr", anchorCtr=True
        ),
        lstStyle=ListStyle(),
        p=[p]
    )

def style_line_chart(chart, title_text, y_title_text, is_rh=False, y_min=None, y_max=None):
    chart.style = 2
    chart.roundedCorners = True
    chart.grouping = "standard"
    chart.varyColors = True
    chart.smooth = False
    chart.plot_vis_only = True
    chart.display_blanks = "gap"

    chart.title = make_chart_title(title_text, size=1400, rot=0)
    chart.auto_title_deleted = False

    if not is_rh:
        chart.layout = Layout(manualLayout=ManualLayout(
            layoutTarget="inner", xMode="edge", yMode="edge",
            x=0.08993685771766159, y=0.11641666666666667,
            w=0.8926455274756951, h=0.4616136111111111
        ))
    else:
        chart.layout = Layout(manualLayout=ManualLayout(
            layoutTarget="inner", xMode="edge", yMode="edge",
            x=0.07902069444444444, y=0.14816666666666667,
            w=0.9015765277777778, h=0.38753027777777777
        ))
    chart.plot_area.graphicalProperties = GraphicalProperties(noFill=True, ln=LineProperties(noFill=True))

    chart.dataLabels = DataLabelList(
        showLegendKey=False, showVal=False, showCatName=False,
        showSerName=False, showPercent=False, showBubbleSize=False
    )

    chart.x_axis.delete = False
    chart.x_axis.axPos = "b"
    chart.x_axis.title = make_chart_title("Time", size=1000, rot=0)
    chart.x_axis.number_format = "General"
    chart.x_axis.sourceLinked = True
    chart.x_axis.majorTickMark = "none"
    chart.x_axis.minorTickMark = "none"
    chart.x_axis.tickLblPos = "nextTo"
    chart.x_axis.crosses = "autoZero"
    chart.x_axis.auto = True
    chart.x_axis.lblAlgn = "ctr"
    chart.x_axis.lblOffset = 100
    chart.x_axis.noMultiLvlLbl = True
    chart.x_axis.txPr = make_axis_txpr(rot=-60000000, size=900)
    chart.x_axis.spPr = GraphicalProperties(
        noFill=True,
        ln=LineProperties(w=9525, cap="flat", cmpd="sng", algn="ctr", solidFill=ColorChoice(schemeClr=SchemeColor(val="tx1", lumMod=15000, lumOff=85000)), round=True)
    )

    chart.y_axis.delete = False
    chart.y_axis.axPos = "l"
    if y_max is not None:
        chart.y_axis.scaling.max = y_max
    if y_min is not None:
        chart.y_axis.scaling.min = y_min
    chart.y_axis.title = make_chart_title(y_title_text, size=1000, rot=-5400000)
    chart.y_axis.number_format = "0.0"
    chart.y_axis.sourceLinked = True
    chart.y_axis.majorTickMark = "none"
    chart.y_axis.minorTickMark = "none"
    chart.y_axis.tickLblPos = "nextTo"
    chart.y_axis.crosses = "autoZero"
    chart.y_axis.crossBetween = "between"
    chart.y_axis.txPr = make_axis_txpr(rot=-60000000, size=900)
    chart.y_axis.spPr = GraphicalProperties(noFill=True, ln=LineProperties(noFill=True))
    chart.y_axis.majorGridlines = ChartLines(
        spPr=GraphicalProperties(ln=LineProperties(w=9525, cap="flat", cmpd="sng", algn="ctr", solidFill=ColorChoice(schemeClr=SchemeColor(val="tx1", lumMod=15000, lumOff=85000)), round=True))
    )

    chart.legend.position = "b"
    chart.legend.overlay = False
    chart.legend.spPr = GraphicalProperties(noFill=True, ln=LineProperties(noFill=True))
    chart.legend.txPr = make_axis_txpr(rot=0, size=900)
    if is_rh:
        chart.legend.layout = Layout(manualLayout=ManualLayout(xMode="edge", yMode="edge", x=0.065875, y=0.8910794444444444, w=0.9, h=0.06093791051498412))

    chart.graphical_properties = GraphicalProperties(
        solidFill=ColorChoice(schemeClr=SchemeColor(val="bg1")),
        ln=LineProperties(w=9525, cap="flat", cmpd="sng", algn="ctr", solidFill=ColorChoice(schemeClr=SchemeColor(val="tx1", lumMod=15000, lumOff=85000)), round=True)
    )

    accent_names = ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"]
    for idx, s in enumerate(chart.series):
        s.smooth = True
        s.marker.symbol = "none"
        acc = accent_names[idx % 6]
        cycle = idx // 6
        if cycle == 0:
            sc = SchemeColor(val=acc)
        elif cycle == 1:
            sc = SchemeColor(val=acc, lumMod=60000)
        else:
            sc = SchemeColor(val=acc, lumMod=80000, lumOff=20000)
        s.graphicalProperties.line = LineProperties(w=28575, cap="rnd", solidFill=ColorChoice(schemeClr=sc), round=True)


def render_excel_chart_png(dates_and_times, series_list, ch_labels, title_text, y_label):
    excel_palette = ["#5B9BD5", "#ED7D31", "#A5A5A5", "#FFC000", "#4472C4", "#70AD47", "#255E91", "#9E480E", "#636363", "#997300", "#264478", "#43682B"]
    fig, ax = plt.subplots(figsize=(8.2, 4.35), dpi=220)
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")

    x = np.arange(len(dates_and_times))
    for idx, arr in enumerate(series_list):
        color = excel_palette[idx % len(excel_palette)]
        ax.plot(x, arr, label=ch_labels[idx], color=color, linewidth=1.6, alpha=0.95)

    ax.set_title(title_text, fontsize=12.5, color="#595959", pad=14, fontweight="normal")
    ax.set_ylabel(y_label, fontsize=9.5, color="#595959")
    ax.set_xlabel("Time", fontsize=9.5, color="#595959", labelpad=8)

    step = max(1, len(dates_and_times) // 42)
    tick_indices = list(range(0, len(dates_and_times), step))
    ax.set_xticks(tick_indices)
    ax.set_xticklabels([dates_and_times[i] for i in tick_indices], rotation=90, fontsize=6.8, color="#595959")

    all_vals = [v for arr in series_list for v in arr]
    vmin, vmax = min(all_vals), max(all_vals)
    span = max(0.5, vmax - vmin)
    y_bottom = np.floor((vmin - span * 0.25) * 2) / 2.0
    y_top = np.ceil((vmax + span * 0.15) * 2) / 2.0
    ax.set_ylim(y_bottom, y_top)
    ax.yaxis.set_major_formatter(matplotlib.ticker.FormatStrFormatter("%.1f"))
    ax.tick_params(axis="y", labelsize=8, colors="#595959", length=0)
    ax.tick_params(axis="x", length=0)

    ax.grid(axis="y", color="#D9D9D9", linestyle="-", linewidth=0.7)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_visible(False)
    ax.spines["bottom"].set_color("#D9D9D9")
    ax.spines["bottom"].set_linewidth(0.8)

    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.44), ncol=min(8, len(ch_labels)), frameon=False, fontsize=7.8, handlelength=1.8, columnspacing=1.2, labelcolor="#595959")
    plt.subplots_adjust(left=0.09, right=0.97, top=0.89, bottom=0.35)

    rect = FancyBboxPatch((0.01, 0.02), 0.98, 0.96, boxstyle="round,pad=0.01,rounding_size=0.02", transform=fig.transFigure, fill=False, edgecolor="#D9D9D9", linewidth=1.0, clip_on=False)
    fig.patches.append(rect)

    buf = io.BytesIO()
    plt.savefig(buf, format="png", dpi=220)
    plt.close(fig)
    buf.seek(0)
    return buf

# Word Table Formatting Helpers
def set_cell_shading(cell, color_hex):
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading)

def set_cell_width_dxa(cell, width_inches):
    tcPr = cell._tc.get_or_add_tcPr()
    w_dxa = int(round(width_inches * 1440))
    tcW = parse_xml(f'<w:tcW {nsdecls("w")} w:w="{w_dxa}" w:type="dxa"/>')
    tcPr.append(tcW)

def set_cell_nowrap(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.append(parse_xml(f'<w:noWrap {nsdecls("w")}/>'))

def set_table_fixed_layout(table, total_width_inches=9.12):
    tblPr = table._tbl.tblPr
    w_dxa = int(round(total_width_inches * 1440))
    tblPr.append(parse_xml(f'<w:tblW {nsdecls("w")} w:w="{w_dxa}" w:type="dxa"/>'))
    tblPr.append(parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>'))

def set_table_borders(table, sz="4", color="000000"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:left w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:bottom w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:right w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:insideH w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:insideV w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)

def set_table_horizontal_borders_only(table, sz="4", color="000000"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:left w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:bottom w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:right w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:insideH w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:insideV w:val="none"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)

def set_cell_right_border(cell, sz="4", color="000000"):
    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(f'<w:tcBorders {nsdecls("w")}><w:right w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/></w:tcBorders>')
    tcPr.append(borders)

def set_row_height(row, height_pt, exact=True):
    trPr = row._tr.get_or_add_trPr()
    val = int(height_pt * 20)
    rule = "exact" if exact else "atLeast"
    trHeight = parse_xml(f'<w:trHeight {nsdecls("w")} w:val="{val}" w:hRule="{rule}"/>')
    trPr.append(trHeight)

def set_row_header(row):
    trPr = row._tr.get_or_add_trPr()
    trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))

def prevent_row_split(row):
    trPr = row._tr.get_or_add_trPr()
    trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

def set_cell_margins(table, top=14, bottom=14, left=20, right=20):
    tblPr = table._tbl.tblPr
    cellMar = parse_xml(f'<w:tblCellMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tblCellMar>')
    tblPr.append(cellMar)

def format_cell(cell, text, bold=False, size=10, align=WD_ALIGN_PARAGRAPH.CENTER, valign=WD_ALIGN_VERTICAL.CENTER, highlight_gray=False, nowrap=False):
    cell.vertical_alignment = valign
    if nowrap:
        set_cell_nowrap(cell)
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.0
    for r in list(p.runs):
        p._p.remove(r._r)
    lines = str(text).split("\n")
    for idx, line in enumerate(lines):
        if idx > 0:
            p = cell.add_paragraph()
            p.alignment = align
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.0
        run = p.add_run(line)
        run.bold = bold
        run.font.name = "Calibri"
        run.font.size = Pt(size)
        if highlight_gray and line.strip():
            run.font.highlight_color = WD_COLOR_INDEX.GRAY_25

def add_page_number_fields(para, font_size=10):
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after = Pt(0)
    def add_txt(t, b=False):
        r = para.add_run(t)
        r.bold = b
        r.font.name = "Calibri"
        r.font.size = Pt(font_size)
    def add_fld(fld_name, b=True):
        r = para.add_run()
        r.bold = b
        r.font.name = "Calibri"
        r.font.size = Pt(font_size)
        fldChar1 = parse_xml(r'<w:fldChar %s w:fldCharType="begin"/>' % nsdecls("w"))
        instrText = parse_xml(r'<w:instrText %s xml:space="preserve"> %s </w:instrText>' % (nsdecls("w"), fld_name))
        fldChar2 = parse_xml(r'<w:fldChar %s w:fldCharType="separate"/>' % nsdecls("w"))
        fldChar3 = parse_xml(r'<w:fldChar %s w:fldCharType="end"/>' % nsdecls("w"))
        r._r.append(fldChar1)
        r._r.append(instrText)
        r._r.append(fldChar2)
        r._r.append(fldChar3)
    add_txt("Page ", False)
    add_fld("PAGE", True)
    add_txt(" of ", False)
    add_fld("NUMPAGES", True)

# ==============================================================================
# DOCX GENERATORS
# ==============================================================================
def generate_raw_data_docx(
    docx_path, client_name, client_address, target_type, target_name, equip_id,
    start_date, next_due, dl_count, start_ht_dl,
    sr_nos, dates_and_times, all_temps, all_rhs, mkt_vals, is_temp_only=False
):
    doc = Document()
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = Inches(11.0)
    section.page_height = Inches(8.5)
    section.top_margin = Inches(0.60)
    section.bottom_margin = Inches(0.55)
    section.left_margin = Inches(0.85)
    section.right_margin = Inches(0.85)
    section.footer_distance = Inches(0.35)

    footer_p = section.footer.paragraphs[0]
    footer_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    add_page_number_fields(footer_p, font_size=10)

    if target_type == "Equipment":
        target_display = f"Equipment Name : {target_name.strip()} ({equip_id.strip()})"
    else:
        target_display = f"Area Name : {target_name.strip()}"

    top_lines = [
        client_name.strip(),
        client_address.strip(),
        target_display,
        f"Temp.Mapping Date : {start_date.strftime('%d/%m/%y')}",
        f"Next Due Date : {next_due.strftime('%d/%m/%y')}"
    ]
    for line in top_lines:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(1.5)
        p.paragraph_format.line_spacing = 1.05
        r = p.add_run(line)
        r.bold = True
        r.font.name = "Calibri"
        r.font.size = Pt(8.5)

    cols_per_dl = 1 if is_temp_only else 2
    total_cols = 2 + dl_count * cols_per_dl
    num_readings = len(sr_nos)
    total_w = 9.30
    sr_w = 0.42
    dt_w = 1.42 
    ch_w = (total_w - sr_w - dt_w) / (dl_count * cols_per_dl)
    col_widths_in = [sr_w, dt_w] + [ch_w] * (dl_count * cols_per_dl)

    tbl = doc.add_table(rows=3 + num_readings, cols=total_cols)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    set_table_fixed_layout(tbl, total_w)
    set_table_borders(tbl, sz="4", color="000000")
    set_cell_margins(tbl, top=10, bottom=10, left=12, right=12)

    for c_i, w_val in enumerate(col_widths_in):
        tbl.columns[c_i].width = Inches(w_val)

    f_sz = 8.0 if dl_count <= 9 else 7.0

    for r_i in range(3):
        set_row_header(tbl.rows[r_i])
        prevent_row_split(tbl.rows[r_i])
        set_row_height(tbl.rows[r_i], 11.5, exact=False)
        for c_i in range(total_cols):
            set_cell_shading(tbl.cell(r_i, c_i), "CAEDFB")
            set_cell_width_dxa(tbl.cell(r_i, c_i), col_widths_in[c_i])

    tbl.cell(0, 0).merge(tbl.cell(1, 1))
    set_cell_width_dxa(tbl.cell(0, 0), sr_w + dt_w)
    format_cell(tbl.cell(0, 0), "ID Number", bold=True, size=f_sz, nowrap=True)
    format_cell(tbl.cell(2, 0), "Sr.No", bold=True, size=f_sz, nowrap=True)
    format_cell(tbl.cell(2, 1), "Date/time", bold=True, size=f_sz, nowrap=True)

    for d in range(dl_count):
        c_start = 2 + d * cols_per_dl
        if is_temp_only:
            set_cell_width_dxa(tbl.cell(0, c_start), ch_w)
            format_cell(tbl.cell(0, c_start), f"DL-{d+1:02d}", bold=True, size=f_sz, nowrap=True)

            set_cell_width_dxa(tbl.cell(1, c_start), ch_w)
            format_cell(tbl.cell(1, c_start), f"HT-DL-{start_ht_dl + d}", bold=True, size=f_sz, nowrap=True)

            format_cell(tbl.cell(2, c_start), f"CH {d+1:02d}", bold=True, size=f_sz, nowrap=True)
        else:
            tbl.cell(0, c_start).merge(tbl.cell(0, c_start + 1))
            set_cell_width_dxa(tbl.cell(0, c_start), ch_w * 2)
            format_cell(tbl.cell(0, c_start), f"DL-{d+1:02d}", bold=True, size=f_sz, nowrap=True)

            tbl.cell(1, c_start).merge(tbl.cell(1, c_start + 1))
            set_cell_width_dxa(tbl.cell(1, c_start), ch_w * 2)
            format_cell(tbl.cell(1, c_start), f"HT-DL-{start_ht_dl + d}", bold=True, size=f_sz, nowrap=True)

            format_cell(tbl.cell(2, c_start), f"CH {d*2+1:02d}", bold=True, size=f_sz, nowrap=True)
            format_cell(tbl.cell(2, c_start + 1), f"CH {d*2+2:02d}", bold=True, size=f_sz, nowrap=True)

    for idx in range(num_readings):
        row_obj = tbl.rows[3 + idx]
        prevent_row_split(row_obj)
        set_row_height(row_obj, 11.0, exact=False)

        set_cell_width_dxa(row_obj.cells[0], sr_w)
        format_cell(row_obj.cells[0], str(sr_nos[idx]), bold=True, size=f_sz, nowrap=True)

        dt_non_breaking = dates_and_times[idx].replace(" ", "\u00A0")
        set_cell_width_dxa(row_obj.cells[1], dt_w)
        format_cell(row_obj.cells[1], dt_non_breaking, bold=False, size=f_sz, nowrap=True)

        for d in range(dl_count):
            c_start = 2 + d * cols_per_dl
            set_cell_width_dxa(row_obj.cells[c_start], ch_w)
            format_cell(row_obj.cells[c_start], f"{all_temps[d][idx]:.1f}", bold=False, size=f_sz, nowrap=True)
            if not is_temp_only:
                set_cell_width_dxa(row_obj.cells[c_start + 1], ch_w)
                format_cell(row_obj.cells[c_start + 1], f"{all_rhs[d][idx]:.1f}", bold=False, size=f_sz, nowrap=True)

    sp = doc.add_paragraph()
    sp.paragraph_format.space_before = Pt(4)
    sp.paragraph_format.space_after = Pt(4)

    sum_labels = ["Average", "Min.", "Max.", "Fluctuation", "MKT"]
    sum_tbl = doc.add_table(rows=5, cols=total_cols)
    sum_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    sum_tbl.autofit = False
    set_table_fixed_layout(sum_tbl, total_w)
    set_table_borders(sum_tbl, sz="4", color="000000")
    set_cell_margins(sum_tbl, top=10, bottom=10, left=12, right=12)

    for c_i, w_val in enumerate(col_widths_in):
        sum_tbl.columns[c_i].width = Inches(w_val)

    for s_idx, s_lbl in enumerate(sum_labels):
        r_obj = sum_tbl.rows[s_idx]
        prevent_row_split(r_obj)
        set_row_height(r_obj, 11.5, exact=False)
        set_cell_shading(r_obj.cells[0], "CAEDFB")
        set_cell_shading(r_obj.cells[1], "CAEDFB")
        r_obj.cells[0].merge(r_obj.cells[1])
        set_cell_width_dxa(r_obj.cells[0], sr_w + dt_w)
        format_cell(r_obj.cells[0], s_lbl, bold=True, size=f_sz, nowrap=True)

        for d in range(dl_count):
            c_start = 2 + d * cols_per_dl
            t_arr = all_temps[d]
            
            if s_idx == 0: 
                v1 = f"{np.mean(t_arr):.1f}"
                v2 = f"{np.mean(all_rhs[d]):.1f}" if not is_temp_only else ""
            elif s_idx == 1: 
                v1 = f"{np.min(t_arr):.1f}"
                v2 = f"{np.min(all_rhs[d]):.1f}" if not is_temp_only else ""
            elif s_idx == 2: 
                v1 = f"{np.max(t_arr):.1f}"
                v2 = f"{np.max(all_rhs[d]):.1f}" if not is_temp_only else ""
            elif s_idx == 3: 
                v1 = f"{np.max(t_arr) - np.min(t_arr):.1f}"
                v2 = f"{np.max(all_rhs[d]) - np.min(all_rhs[d]):.1f}" if not is_temp_only else ""
            else: 
                v1 = f"{mkt_vals[d]:.1f}"
                v2 = "NA"
            
            set_cell_width_dxa(r_obj.cells[c_start], ch_w)
            format_cell(r_obj.cells[c_start], v1, bold=False, size=f_sz, nowrap=True)
            if not is_temp_only:
                set_cell_width_dxa(r_obj.cells[c_start + 1], ch_w)
                format_cell(r_obj.cells[c_start + 1], v2, bold=False, size=f_sz, nowrap=True)

    doc.save(docx_path)


def generate_mapping_report_docx(
    docx_path, client_name, client_address, target_type, target_name, equip_id, location_name,
    report_no, validated_by, start_date, end_date, next_due,
    duration_hrs, freq_mins, dl_count, start_ht_dl,
    temp_criteria, rh_criteria,
    dl_cal_date, dl_due_date, dl_cert_no,
    master_id, master_cert, master_cal_date, master_due_date,
    all_temps, all_rhs, mkt_vals,
    chart_t_buf, chart_rh_buf, is_temp_only=False
):
    doc = Document()

    t_mins = [round(float(np.min(arr)), 1) for arr in all_temps]
    t_maxs = [round(float(np.max(arr)), 1) for arr in all_temps]
    t_flucs = [round(t_maxs[d] - t_mins[d], 1) for d in range(dl_count)]

    min_mkt_val = min(mkt_vals)
    min_mkt_dl = f"DL{int(np.argmin(mkt_vals)) + 1:02d}"
    max_mkt_val = max(mkt_vals)
    max_mkt_dl = f"DL{int(np.argmax(mkt_vals)) + 1:02d}"
    max_t_fluc_val = max(t_flucs)
    max_t_fluc_dl = f"DL{int(np.argmax(t_flucs)) + 1:02d}"

    if not is_temp_only:
        rh_mins = [round(float(np.min(arr)), 1) for arr in all_rhs]
        rh_maxs = [round(float(np.max(arr)), 1) for arr in all_rhs]
        rh_flucs = [round(rh_maxs[d] - rh_mins[d], 1) for d in range(dl_count)]
        rh_avgs = [round(float(np.mean(arr)), 1) for arr in all_rhs]

        min_rh_avg_val = min(rh_avgs)
        min_rh_avg_dl = f"DL{int(np.argmin(rh_avgs)) + 1:02d}"
        max_rh_avg_val = max(rh_avgs)
        max_rh_avg_dl = f"DL{int(np.argmax(rh_avgs)) + 1:02d}"
        max_rh_fluc_val = max(rh_flucs)
        max_rh_fluc_dl = f"DL{int(np.argmax(rh_flucs)) + 1:02d}"

    sec1 = doc.sections[0]
    sec1.orientation = WD_ORIENT.LANDSCAPE
    sec1.page_width = Inches(11.0)
    sec1.page_height = Inches(8.5)
    sec1.top_margin = Inches(1.5)
    sec1.bottom_margin = Inches(1.0)
    sec1.left_margin = Inches(0.94)
    sec1.right_margin = Inches(0.94)

    # PAGE 1: COVER
    cover_box = doc.add_table(rows=1, cols=1)
    cover_box.alignment = WD_TABLE_ALIGNMENT.CENTER
    cover_box.autofit = False
    cover_box.columns[0].width = Inches(9.12)
    set_table_borders(cover_box, sz="6", color="000000")
    set_row_height(cover_box.rows[0], 345.0, exact=True)
    c_cell = cover_box.cell(0, 0)
    c_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    def add_cover_line(cell, text, size, bold=True, space_after=8, first=False):
        p = cell.paragraphs[0] if first else cell.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(space_after)
        r = p.add_run(text)
        r.bold = bold
        r.font.name = "Calibri"
        r.font.size = Pt(size)

    report_title = "Temperature Mapping Report of" if is_temp_only else "Temperature & Humidity Mapping Report of"
    add_cover_line(c_cell, report_title, 24, True, 18, first=True)
    if target_type == "Equipment":
        add_cover_line(c_cell, f"{target_name.strip().title()} ID No. {equip_id.strip()}", 24, True, 16)
    else:
        add_cover_line(c_cell, target_name.strip().title(), 24, True, 16)
        
    add_cover_line(c_cell, f"Report No.: {report_no}", 13, True, 12)
    add_cover_line(c_cell, f"Start Date: {start_date.strftime('%d/%m/%Y')}                End Date: {end_date.strftime('%d/%m/%Y')}", 13, True, 12)
    add_cover_line(c_cell, f"Next Due Date: {next_due.strftime('%d/%m/%Y')}", 13, True, 22)
    add_cover_line(c_cell, "Customer Name & Address:", 18.5, True, 10)
    add_cover_line(c_cell, client_name.strip(), 18.5, True, 8)

    addr_parts = client_address.strip().split("Tel-")
    if len(addr_parts) == 2:
        add_cover_line(c_cell, addr_parts[0].strip(), 11, True, 4)
        add_cover_line(c_cell, "Tel-" + addr_parts[1].strip(), 11, True, 0)
    else:
        add_cover_line(c_cell, client_address.strip(), 11, True, 0)

    sec2 = doc.add_section()
    sec2.orientation = WD_ORIENT.LANDSCAPE
    sec2.page_width = Inches(11.0)
    sec2.page_height = Inches(8.5)
    sec2.top_margin = Inches(0.75)
    sec2.bottom_margin = Inches(0.65)
    sec2.left_margin = Inches(0.94)
    sec2.right_margin = Inches(0.94)
    sec2.header_distance = Inches(0.45)
    sec2.footer_distance = Inches(0.45)
    sec2.header.is_linked_to_previous = False
    sec2.footer.is_linked_to_previous = False

    hdr_p0 = sec2.header.paragraphs[0]
    hdr_p0.paragraph_format.space_before = Pt(0)
    hdr_p0.paragraph_format.space_after = Pt(0)

    hdr_tbl = sec2.header.add_table(rows=2, cols=6, width=Inches(9.12))
    hdr_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr_tbl.autofit = False
    set_table_horizontal_borders_only(hdr_tbl, sz="4", color="000000")
    set_cell_margins(hdr_tbl, top=20, bottom=20, left=40, right=40)

    hdr_widths = [1.14, 0.18, 3.26, 1.15, 0.18, 3.21]
    for c_i, w_v in enumerate(hdr_widths): hdr_tbl.columns[c_i].width = Inches(w_v)

    set_row_height(hdr_tbl.rows[0], 24.0, exact=True)
    set_row_height(hdr_tbl.rows[1], 16.0, exact=True)
    set_cell_right_border(hdr_tbl.cell(0, 2), sz="4", color="000000")
    set_cell_right_border(hdr_tbl.cell(1, 2), sz="4", color="000000")

    format_cell(hdr_tbl.cell(0, 0), "Customer Name", bold=True, size=9.5, align=WD_ALIGN_PARAGRAPH.LEFT)
    format_cell(hdr_tbl.cell(0, 1), ":", bold=True, size=9.5, align=WD_ALIGN_PARAGRAPH.CENTER)
    format_cell(hdr_tbl.cell(0, 2), client_name.strip().upper(), bold=True, size=10.5, align=WD_ALIGN_PARAGRAPH.LEFT)
    
    lbl_str = "Equipment Name" if target_type == "Equipment" else "Area Name"
    format_cell(hdr_tbl.cell(0, 3), lbl_str, bold=True, size=11, align=WD_ALIGN_PARAGRAPH.LEFT)
    format_cell(hdr_tbl.cell(0, 4), ":", bold=True, size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
    format_cell(hdr_tbl.cell(0, 5), target_name.strip().title(), bold=True, size=11, align=WD_ALIGN_PARAGRAPH.LEFT)

    format_cell(hdr_tbl.cell(1, 0), "Report No.", bold=True, size=9.5, align=WD_ALIGN_PARAGRAPH.LEFT)
    format_cell(hdr_tbl.cell(1, 1), ":", bold=True, size=9.5, align=WD_ALIGN_PARAGRAPH.CENTER)
    format_cell(hdr_tbl.cell(1, 2), report_no, bold=True, size=10.5, align=WD_ALIGN_PARAGRAPH.LEFT)
    format_cell(hdr_tbl.cell(1, 3), "Validated By", bold=True, size=11, align=WD_ALIGN_PARAGRAPH.LEFT)
    format_cell(hdr_tbl.cell(1, 4), ":", bold=True, size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
    format_cell(hdr_tbl.cell(1, 5), validated_by, bold=True, size=11, align=WD_ALIGN_PARAGRAPH.LEFT)

    ftr_p = sec2.footer.paragraphs[0]
    ftr_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_page_number_fields(ftr_p, font_size=11)

    def add_framed_divider_page(lines_with_hl, page_break_after=True):
        sp = doc.add_paragraph()
        sp.paragraph_format.space_before = Pt(4)
        sp.paragraph_format.space_after = Pt(4)
        box = doc.add_table(rows=1, cols=1)
        box.alignment = WD_TABLE_ALIGNMENT.CENTER
        box.autofit = False
        box.columns[0].width = Inches(9.12)
        set_table_borders(box, sz="6", color="000000")
        set_row_height(box.rows[0], 355.0, exact=True)
        cell = box.cell(0, 0)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        for idx, (txt, sz, hl) in enumerate(lines_with_hl):
            p = cell.paragraphs[0] if idx == 0 else cell.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            r = p.add_run(txt)
            r.bold = True
            r.font.name = "Calibri"
            r.font.size = Pt(sz)
            if hl: r.font.highlight_color = WD_COLOR_INDEX.GRAY_25

        if page_break_after: doc.add_page_break()

    # PAGE 2: TOC
    sp2 = doc.add_paragraph()
    sp2.paragraph_format.space_before = Pt(6)
    sp2.paragraph_format.space_after = Pt(6)

    if not is_temp_only:
        toc_rows = [
            ("Sr. No", "Content", "Page No"),
            ("1.", "Mapping Details and Details of Data Logger Used", "3-4"),
            ("2.", "Diagram Indicating Data Logger Locations", "5-6"),
            ("3.", "Summary Report (Min. Max & Maximum Fluctuation And MKT)", "7-10"),
            ("4.", "Trend (Temperature)", "11-12"),
            ("5.", "Trend (RH)", "13-14"),
            ("6.", "Observation & Conclusion", "15-16"),
            ("7.", "Diagram Indicating Minimum, Maximum & Maximum Fluctuation Locations", "17-18"),
            ("8.", "Attachments – Raw Data & Traceability", "19"),
            ("", "", "")
        ]
    else:
        toc_rows = [
            ("Sr. No", "Content", "Page No"),
            ("1.", "Mapping Details and Details of Data Logger Used", "3-4"),
            ("2.", "Diagram Indicating Data Logger Locations", "5-6"),
            ("3.", "Summary Report (Min. Max & Maximum Fluctuation And MKT)", "7-8"),
            ("4.", "Trend (Temperature)", "9-10"),
            ("5.", "Observation & Conclusion", "11-12"),
            ("6.", "Diagram Indicating Minimum, Maximum & Maximum Fluctuation Locations", "13-14"),
            ("7.", "Attachments – Raw Data & Traceability", "15"),
            ("", "", "")
        ]

    toc_tbl = doc.add_table(rows=len(toc_rows), cols=3)
    toc_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    toc_tbl.autofit = False
    set_table_borders(toc_tbl, sz="4", color="000000")
    set_cell_margins(toc_tbl, top=25, bottom=25, left=60, right=60)
    toc_widths = [0.80, 7.29, 1.03]
    for c_i, w_v in enumerate(toc_widths): toc_tbl.columns[c_i].width = Inches(w_v)

    last_r_idx = len(toc_rows) - 1
    for r_i, (c1, c2, c3) in enumerate(toc_rows):
        set_row_height(toc_tbl.rows[r_i], 22.0 if r_i in (0, last_r_idx) else 27.0, exact=True)
        if r_i == 0:
            format_cell(toc_tbl.cell(r_i, 0), c1, bold=True, size=11, align=WD_ALIGN_PARAGRAPH.LEFT)
            format_cell(toc_tbl.cell(r_i, 1), c2, bold=True, size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
            format_cell(toc_tbl.cell(r_i, 2), c3, bold=True, size=11, align=WD_ALIGN_PARAGRAPH.LEFT)
        elif r_i == last_r_idx:
            toc_tbl.cell(last_r_idx, 0).merge(toc_tbl.cell(last_r_idx, 2))
        else:
            format_cell(toc_tbl.cell(r_i, 0), c1, bold=False, size=10.5, align=WD_ALIGN_PARAGRAPH.CENTER)
            format_cell(toc_tbl.cell(r_i, 1), c2, bold=False, size=10.5, align=WD_ALIGN_PARAGRAPH.LEFT)
            format_cell(toc_tbl.cell(r_i, 2), c3, bold=False, size=10.5, align=WD_ALIGN_PARAGRAPH.LEFT)

    doc.add_page_break()

    # PAGE 3: SECTION 2.0 DIVIDER
    add_framed_divider_page([("2.0 - MAPPING DETAILS & DETAILS OF DATA LOGGERS USED:", 16.5, True)])

    # PAGE 4: 2.0 MAPPING DETAILS TABLES
    p_h2 = doc.add_paragraph()
    p_h2.paragraph_format.space_before = Pt(2)
    p_h2.paragraph_format.space_after = Pt(4)
    r_h2 = p_h2.add_run("2.0 MAPPING DETAILS")
    r_h2.bold = True
    r_h2.font.name = "Calibri"
    r_h2.font.size = Pt(11)

    t1 = doc.add_table(rows=4, cols=8)
    t1.alignment = WD_TABLE_ALIGNMENT.CENTER
    t1.autofit = False
    set_table_borders(t1, sz="4", color="000000")
    set_cell_margins(t1, top=20, bottom=20, left=45, right=45)

    t1_widths = [1.46, 1.69, 1.04, 0.72, 0.36, 0.83, 1.82, 1.20]
    for c_i, w_v in enumerate(t1_widths): t1.columns[c_i].width = Inches(w_v)
    for r_i in range(4): set_row_height(t1.rows[r_i], 26.5, exact=True)

    t1.cell(0, 2).merge(t1.cell(0, 5))
    t1.cell(2, 2).merge(t1.cell(2, 3))
    t1.cell(2, 4).merge(t1.cell(2, 5))
    t1.cell(3, 2).merge(t1.cell(3, 3))
    t1.cell(3, 4).merge(t1.cell(3, 5))

    if target_type == "Equipment":
        format_cell(t1.cell(0, 0), "Name of the Equipment", bold=True, size=10.5, align=WD_ALIGN_PARAGRAPH.LEFT)
        format_cell(t1.cell(0, 1), target_name.strip().title(), bold=False, size=10.5, align=WD_ALIGN_PARAGRAPH.LEFT)
        format_cell(t1.cell(1, 0), "ID No.", bold=True, size=10.5, align=WD_ALIGN_PARAGRAPH.LEFT)
        format_cell(t1.cell(1, 1), equip_id.strip(), bold=False, size=10.5, align=WD_ALIGN_PARAGRAPH.LEFT)
    else:
        t1.cell(0, 0).merge(t1.cell(1, 0))
        t1.cell(0, 1).merge(t1.cell(1, 1))
        format_cell(t1.cell(0, 0), "Name of Area", bold=True, size=10.5, align=WD_ALIGN_PARAGRAPH.LEFT)
        format_cell(t1.cell(0, 1), target_name.strip().title(), bold=False, size=10.5, align=WD_ALIGN_PARAGRAPH.LEFT)

    format_cell(t1.cell(0, 2), "Acceptance Criteria", bold=True, size=10.5, align=WD_ALIGN_PARAGRAPH.CENTER)
    format_cell(t1.cell(0, 6), "Frequency of Reading", bold=True, size=10.5, align=WD_ALIGN_PARAGRAPH.LEFT)
    format_cell(t1.cell(0, 7), f"{freq_mins} minutes", bold=False, size=10.5, align=WD_ALIGN_PARAGRAPH.LEFT)

    format_cell(t1.cell(1, 2), "Temperature", bold=True, size=10.5, align=WD_ALIGN_PARAGRAPH.LEFT)
    format_cell(t1.cell(1, 3), temp_criteria, bold=False, size=10, align=WD_ALIGN_PARAGRAPH.CENTER)
    format_cell(t1.cell(1, 4), "RH", bold=True, size=10.5, align=WD_ALIGN_PARAGRAPH.CENTER)
    format_cell(t1.cell(1, 5), "N/A" if is_temp_only else rh_criteria, bold=False, size=10, align=WD_ALIGN_PARAGRAPH.CENTER)
    format_cell(t1.cell(1, 6), "Mapping Duration", bold=True, size=10.5, align=WD_ALIGN_PARAGRAPH.LEFT)
    format_cell(t1.cell(1, 7), f"{duration_hrs} Hour", bold=False, size=10.5, align=WD_ALIGN_PARAGRAPH.LEFT)

    format_cell(t1.cell(2, 0), "Location", bold=True, size=10.5, align=WD_ALIGN_PARAGRAPH.LEFT)
    format_cell(t1.cell(2, 1), location_name, bold=False, size=10.5, align=WD_ALIGN_PARAGRAPH.LEFT)
    format_cell(t1.cell(2, 2), "Mapping Date (Start)", bold=True, size=10.5, align=WD_ALIGN_PARAGRAPH.LEFT)
    format_cell(t1.cell(2, 4), start_date.strftime("%d/%m/%Y"), bold=False, size=10.5, align=WD_ALIGN_PARAGRAPH.LEFT)
    format_cell(t1.cell(2, 6), "Start Time", bold=True, size=10.5, align=WD_ALIGN_PARAGRAPH.LEFT)
    format_cell(t1.cell(2, 7), start_date.strftime("%H:%M:%S"), bold=False, size=10.5, align=WD_ALIGN_PARAGRAPH.LEFT)

    format_cell(t1.cell(3, 0), "No. of Data Loggers", bold=True, size=10.5, align=WD_ALIGN_PARAGRAPH.LEFT)
    format_cell(t1.cell(3, 1), f"{dl_count:02d}", bold=False, size=10.5, align=WD_ALIGN_PARAGRAPH.LEFT)
    format_cell(t1.cell(3, 2), "Mapping Date (End)", bold=True, size=10.5, align=WD_ALIGN_PARAGRAPH.LEFT)
    format_cell(t1.cell(3, 4), end_date.strftime("%d/%m/%Y"), bold=False, size=10.5, align=WD_ALIGN_PARAGRAPH.LEFT)
    format_cell(t1.cell(3, 6), "End Time", bold=True, size=10.5, align=WD_ALIGN_PARAGRAPH.LEFT)
    format_cell(t1.cell(3, 7), end_date.strftime("%H:%M:%S"), bold=False, size=10.5, align=WD_ALIGN_PARAGRAPH.LEFT)

    p_dl = doc.add_paragraph()
    p_dl.paragraph_format.space_before = Pt(8)
    p_dl.paragraph_format.space_after = Pt(2)
    r_dl = p_dl.add_run("DETAILS OF DATA LOGGERS USED:")
    r_dl.bold = True
    r_dl.font.name = "Calibri"
    r_dl.font.size = Pt(11)

    t2 = doc.add_table(rows=2, cols=8)
    t2.alignment = WD_TABLE_ALIGNMENT.CENTER
    t2.autofit = False
    set_table_borders(t2, sz="4", color="000000")
    set_cell_margins(t2, top=20, bottom=20, left=45, right=45)

    t2_widths = [1.15, 1.14, 1.07, 1.10, 1.03, 1.24, 1.24, 1.15]
    for c_i, w_v in enumerate(t2_widths): t2.columns[c_i].width = Inches(w_v)
    set_row_height(t2.rows[0], 19.0, exact=True)
    set_row_height(t2.rows[1], 48.0, exact=True)

    t2_headers = ["Name & make", "ID No.", "Make", "Range", "Least Count", "Certificate No.", "Calibration Date", "Next Due Date"]
    t2_vals = [
        "Temperature PDF Data\nLogger" if is_temp_only else "Temperature &\nRH PDF Data\nLogger",
        f"HT-DL-{start_ht_dl} to\nHT-DL-{start_ht_dl + dl_count - 1}",
        "Testo",
        "- 20 to 70°C" if is_temp_only else "- 20 to 70°C /\n0 to 100% RH",
        "0.1 ºC" if is_temp_only else "0.1 ºC / 0.1 %\nRH",
        dl_cert_no.replace(" & ", "\n&\n"),
        dl_cal_date,
        dl_due_date
    ]
    for c_i in range(8):
        format_cell(t2.cell(0, c_i), t2_headers[c_i], bold=True, size=10.5, align=WD_ALIGN_PARAGRAPH.LEFT)
        format_cell(t2.cell(1, c_i), t2_vals[c_i], bold=False, size=10.5, align=WD_ALIGN_PARAGRAPH.LEFT)

    p_mc = doc.add_paragraph()
    p_mc.paragraph_format.space_before = Pt(10)
    p_mc.paragraph_format.space_after = Pt(2)
    r_mc = p_mc.add_run("MASTER CALIBRATION DETAILS:")
    r_mc.bold = True
    r_mc.font.name = "Calibri"
    r_mc.font.size = Pt(11)

    t3 = doc.add_table(rows=2, cols=6)
    t3.alignment = WD_TABLE_ALIGNMENT.CENTER
    t3.autofit = False
    set_table_borders(t3, sz="4", color="000000")
    set_cell_margins(t3, top=20, bottom=20, left=45, right=45)

    t3_widths = [1.75, 0.99, 1.65, 1.38, 1.54, 1.81]
    for c_i, w_v in enumerate(t3_widths): t3.columns[c_i].width = Inches(w_v)
    set_row_height(t3.rows[0], 19.0, exact=True)
    set_row_height(t3.rows[1], 50.0, exact=True)

    t3_headers = ["Description", "ID NO", "Certificate No.", "Calibration Date", "Calibration Due", "Calibration By"]
    t3_vals = [
        "Digital Temperature Indicator\nwith Sensor" if is_temp_only else "Digital Temperature &\nhumidity Indicator with\nSensor",
        master_id,
        master_cert,
        master_cal_date,
        master_due_date,
        "RK Quality Measurement\nServices Pvt. Ltd."
    ]
    for c_i in range(6):
        format_cell(t3.cell(0, c_i), t3_headers[c_i], bold=True, size=10.5, align=WD_ALIGN_PARAGRAPH.LEFT)
        format_cell(t3.cell(1, c_i), t3_vals[c_i], bold=False, size=10.5, align=WD_ALIGN_PARAGRAPH.LEFT)

    doc.add_page_break()

    # PAGE 5: SECTION 3.0 DIVIDER
    add_framed_divider_page([("3.0 - DIAGRAM INDICATING DATA LOGGERS LOCATIONS", 16.5, True)])

    # PAGE 6: 3.0 DIAGRAM BOX
    sp6 = doc.add_paragraph()
    sp6.paragraph_format.space_before = Pt(4)
    sp6.paragraph_format.space_after = Pt(4)
    box6 = doc.add_table(rows=1, cols=1)
    box6.alignment = WD_TABLE_ALIGNMENT.CENTER
    box6.autofit = False
    box6.columns[0].width = Inches(9.12)
    set_table_borders(box6, sz="6", color="000000")
    set_row_height(box6.rows[0], 345.0, exact=False)
    c6 = box6.cell(0, 0)
    c6.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    p6 = c6.paragraphs[0]
    p6.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p6.paragraph_format.space_before = Pt(6)
    p6.paragraph_format.space_after = Pt(12)
    r6 = p6.add_run("3.0 - DIAGRAM INDICATING DATA LOGGERS LOCATIONS")
    r6.bold = True
    r6.font.name = "Calibri"
    r6.font.size = Pt(16.5)
    c6.add_paragraph().alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_page_break()

    # SECTION 4.0 DIVIDER
    add_framed_divider_page([("4.0 - Summary Report", 16.5, True), ("(Min. Max & Maximum Fluctuation and MKT):", 16.5, True)])

    # SUMMARY REPORT (TEMPERATURE)
    sp8 = doc.add_paragraph()
    sp8.paragraph_format.space_before = Pt(2)
    sp8.paragraph_format.space_after = Pt(2)

    total_t_rows = 4 + dl_count + 5
    t8 = doc.add_table(rows=total_t_rows, cols=5)
    t8.alignment = WD_TABLE_ALIGNMENT.CENTER
    t8.autofit = False
    set_table_borders(t8, sz="4", color="000000")
    set_cell_margins(t8, top=14, bottom=14, left=30, right=30)
    t8_widths = [1.10, 1.76, 1.58, 2.06, 2.62]
    for c_i, w_v in enumerate(t8_widths): t8.columns[c_i].width = Inches(w_v)

    dl_row_h = min(25.5, 205.0 / max(1, dl_count))
    set_row_height(t8.rows[0], 25.0, exact=True)
    for i in range(1, 4): set_row_height(t8.rows[i], 14.0, exact=True)

    for c_i in range(5): set_cell_shading(t8.cell(0, c_i), "F2F2F2")
    t8.cell(0, 0).merge(t8.cell(0, 4))
    format_cell(t8.cell(0, 0), f"4.0 - Summary Report - (Acceptance Criteria – Temperature {temp_criteria} )", bold=True, size=11)
    t8.cell(1, 0).merge(t8.cell(1, 4))
    format_cell(t8.cell(1, 0), "Minimum, Maximum, Maximum Fluctuation of Temperature and MKT During Monitoring Period:", bold=True, size=10.5)
    t8.cell(2, 0).merge(t8.cell(2, 4))
    format_cell(t8.cell(2, 0), "Temperature", bold=True, size=10.5)

    t8_headers = ["Data Logger", "Min. Temp (ºC)", "Max. Temp (ºC)", "Fluctuation (ºC)", "MKT"]
    for c_i, h_str in enumerate(t8_headers): format_cell(t8.cell(3, c_i), h_str, bold=True, size=10.5)

    f_dl_sz = 10.5 if dl_count <= 9 else 9.0
    for d in range(dl_count):
        r_idx = 4 + d
        set_row_height(t8.rows[r_idx], dl_row_h, exact=True)
        format_cell(t8.cell(r_idx, 0), f"DL{d+1:02d}", bold=False, size=f_dl_sz)
        format_cell(t8.cell(r_idx, 1), f"{t_mins[d]:.1f}", bold=False, size=f_dl_sz)
        format_cell(t8.cell(r_idx, 2), f"{t_maxs[d]:.1f}", bold=False, size=f_dl_sz)
        format_cell(t8.cell(r_idx, 3), f"{t_flucs[d]:.1f}", bold=False, size=f_dl_sz)
        format_cell(t8.cell(r_idx, 4), f"{mkt_vals[d]:.1f}", bold=False, size=f_dl_sz)

    b_start = 4 + dl_count
    for offset in range(5): set_row_height(t8.rows[b_start + offset], 14.0, exact=True)

    t8.cell(b_start, 0).merge(t8.cell(b_start, 4))
    format_cell(t8.cell(b_start, 0), "Minimum, Maximum, Maximum Fluctuation Temperature Locations", bold=True, size=10.5)

    t8.cell(b_start + 1, 0).merge(t8.cell(b_start + 4, 0))
    format_cell(t8.cell(b_start + 1, 0), "Observation & Locations", bold=True, size=10.5)

    t8.cell(b_start + 1, 1).merge(t8.cell(b_start + 1, 2))
    format_cell(t8.cell(b_start + 1, 1), "Min. MKT Temp (ºC)", bold=True, size=10.5)
    t8.cell(b_start + 1, 3).merge(t8.cell(b_start + 1, 4))
    format_cell(t8.cell(b_start + 1, 3), "Max. MKT Temp (ºC)", bold=True, size=10.5)

    t8.cell(b_start + 2, 1).merge(t8.cell(b_start + 2, 2))
    format_cell(t8.cell(b_start + 2, 1), f"{min_mkt_val:.1f} at Locations {min_mkt_dl}", bold=False, size=10.5)
    t8.cell(b_start + 2, 3).merge(t8.cell(b_start + 2, 4))
    format_cell(t8.cell(b_start + 2, 3), f"{max_mkt_val:.1f} at Locations {max_mkt_dl}", bold=False, size=10.5)

    t8.cell(b_start + 3, 1).merge(t8.cell(b_start + 3, 4))
    format_cell(t8.cell(b_start + 3, 1), "Maximum Fluctuation", bold=True, size=10.5)
    t8.cell(b_start + 4, 1).merge(t8.cell(b_start + 4, 4))
    format_cell(t8.cell(b_start + 4, 1), f"{max_t_fluc_val:.1f} at Location {max_t_fluc_dl}", bold=False, size=10.5)

    doc.add_page_break()

    # SUMMARY REPORT (RH) - ONLY IF NOT TEMP ONLY
    if not is_temp_only:
        sp9 = doc.add_paragraph()
        sp9.paragraph_format.space_before = Pt(2)
        sp9.paragraph_format.space_after = Pt(2)

        t9 = doc.add_table(rows=total_t_rows, cols=5)
        t9.alignment = WD_TABLE_ALIGNMENT.CENTER
        t9.autofit = False
        set_table_borders(t9, sz="4", color="000000")
        set_cell_margins(t9, top=14, bottom=14, left=30, right=30)
        for c_i, w_v in enumerate(t8_widths): t9.columns[c_i].width = Inches(w_v)

        set_row_height(t9.rows[0], 25.0, exact=True)
        for i in range(1, 4): set_row_height(t9.rows[i], 14.0, exact=True)

        for c_i in range(5): set_cell_shading(t9.cell(0, c_i), "F2F2F2")
        t9.cell(0, 0).merge(t9.cell(0, 4))
        format_cell(t9.cell(0, 0), f"4.0 - Summary Report - (Acceptance Criteria – RH {rh_criteria})", bold=True, size=11)
        t9.cell(1, 0).merge(t9.cell(1, 4))
        format_cell(t9.cell(1, 0), "Minimum, Maximum, Fluctuation of RH and average During Monitoring Period:", bold=True, size=10.5)
        t9.cell(2, 0).merge(t9.cell(2, 4))
        format_cell(t9.cell(2, 0), "RH", bold=True, size=10.5)

        t9_headers = ["Data Logger", "Min. RH (%)", "Max. RH (%)", "Fluctuation", "Average"]
        for c_i, h_str in enumerate(t9_headers): format_cell(t9.cell(3, c_i), h_str, bold=True, size=10.5)

        for d in range(dl_count):
            r_idx = 4 + d
            set_row_height(t9.rows[r_idx], dl_row_h, exact=True)
            format_cell(t9.cell(r_idx, 0), f"DL{d+1:02d}", bold=False, size=f_dl_sz)
            format_cell(t9.cell(r_idx, 1), f"{rh_mins[d]:.1f}", bold=False, size=f_dl_sz)
            format_cell(t9.cell(r_idx, 2), f"{rh_maxs[d]:.1f}", bold=False, size=f_dl_sz)
            format_cell(t9.cell(r_idx, 3), f"{rh_flucs[d]:.1f}", bold=False, size=f_dl_sz)
            format_cell(t9.cell(r_idx, 4), f"{rh_avgs[d]:.1f}", bold=False, size=f_dl_sz)

        for offset in range(5): set_row_height(t9.rows[b_start + offset], 14.0, exact=True)

        t9.cell(b_start, 0).merge(t9.cell(b_start, 4))
        format_cell(t9.cell(b_start, 0), "Minimum, Maximum, Maximum Fluctuation RH Locations", bold=True, size=10.5)

        t9.cell(b_start + 1, 0).merge(t9.cell(b_start + 4, 0))
        format_cell(t9.cell(b_start + 1, 0), "Observation & Locations", bold=True, size=10.5)

        t9.cell(b_start + 1, 1).merge(t9.cell(b_start + 1, 2))
        format_cell(t9.cell(b_start + 1, 1), "Min. Avg RH (%)", bold=True, size=10.5)
        t9.cell(b_start + 1, 3).merge(t9.cell(b_start + 1, 4))
        format_cell(t9.cell(b_start + 1, 3), "Max. Avg RH (%)", bold=True, size=10.5)

        t9.cell(b_start + 2, 1).merge(t9.cell(b_start + 2, 2))
        format_cell(t9.cell(b_start + 2, 1), f"{min_rh_avg_val:.1f} % at Locations {min_rh_avg_dl}", bold=False, size=10.5)
        t9.cell(b_start + 2, 3).merge(t9.cell(b_start + 2, 4))
        format_cell(t9.cell(b_start + 2, 3), f"{max_rh_avg_val:.1f} at Locations {max_rh_avg_dl}", bold=False, size=10.5)

        t9.cell(b_start + 3, 1).merge(t9.cell(b_start + 3, 4))
        format_cell(t9.cell(b_start + 3, 1), "Maximum Fluctuation", bold=True, size=10.5)
        t9.cell(b_start + 4, 1).merge(t9.cell(b_start + 4, 4))
        format_cell(t9.cell(b_start + 4, 1), f"{max_rh_fluc_val:.1f} at Location {max_rh_fluc_dl}", bold=False, size=10.5)

        doc.add_page_break()

    # MKT EXPLANATION PAGE
    p10_title = doc.add_paragraph()
    p10_title.paragraph_format.space_before = Pt(14)
    p10_title.paragraph_format.space_after = Pt(12)
    p10_title.add_run("MEAN KINETIC TEMPERATURE:").bold = True
    p10_title.runs[0].font.name = "Calibri"
    p10_title.runs[0].font.size = Pt(13)

    p10_desc = doc.add_paragraph()
    p10_desc.paragraph_format.space_before = Pt(0)
    p10_desc.paragraph_format.space_after = Pt(18)
    r10_q = p10_desc.add_run("“")
    r10_q.font.name, r10_q.font.size = "Calibri", Pt(13)
    r10_d = p10_desc.add_run("A Single derived temperature that, if maintained over a defined period of time, affords the same thermal challenge to a drug substance or drug product as\nwould be experienced over a range of both higher and lower temperature for an equivalent defined period.")
    r10_d.font.name, r10_d.font.size = "Calibri", Pt(10.5)

    p10_eq = doc.add_paragraph()
    p10_eq.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p10_eq.paragraph_format.space_before = Pt(6)
    p10_eq.paragraph_format.space_after = Pt(20)
    r_eq = p10_eq.add_run("                  -ΔH / R\nTK  =  ─────────────────────────────────────────\n        - ln ( (e^(-ΔH/RT₁) + e^(-ΔH/RT₂) + ... + e^(-ΔH/RTₙ)) / n )")
    r_eq.font.name, r_eq.font.size = "Calibri", Pt(11)

    where_lines = [
        "Where:",
        "ΔH = the heat of activation for the degradation reaction; assumed to be 83.144 kJ per mole",
        "R = 8.3144 × 10⁻³ kJ per degree per mole (the universal gas constant)",
        "T1 = the average temperature, in degrees Kelvin, during the first time period",
        "T2 = the average temperature, in degrees Kelvin, during the second time period",
        "Tn = the average temperature, in degrees Kelvin, during the nth time period",
        "n = the total number of Temperature recorded."
    ]
    for wl in where_lines:
        pw = doc.add_paragraph()
        pw.paragraph_format.space_before = Pt(0)
        pw.paragraph_format.space_after = Pt(2)
        rw = pw.add_run(wl)
        rw.font.name, rw.font.size = "Calibri", Pt(10.5)

    doc.add_page_break()

    # TEMPERATURE TREND
    add_framed_divider_page([("5.0 – TREND", 16.5, True), ("Temperature", 11.0, False)])

    sp12 = doc.add_paragraph()
    sp12.paragraph_format.space_before = Pt(4)
    sp12.paragraph_format.space_after = Pt(4)
    box12 = doc.add_table(rows=1, cols=1)
    box12.alignment = WD_TABLE_ALIGNMENT.CENTER
    box12.autofit = False
    box12.columns[0].width = Inches(9.12)
    set_table_borders(box12, sz="6", color="000000")
    set_row_height(box12.rows[0], 355.0, exact=True)
    c12 = box12.cell(0, 0)
    c12.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p12 = c12.paragraphs[0]
    p12.alignment = WD_ALIGN_PARAGRAPH.CENTER
    chart_t_buf.seek(0)
    p12.add_run().add_picture(chart_t_buf, width=Inches(7.9))

    doc.add_page_break()

    # RH TREND - ONLY IF NOT TEMP ONLY
    if not is_temp_only:
        add_framed_divider_page([("6.0 – TREND", 16.5, True), ("RH", 11.0, False)])

        sp14 = doc.add_paragraph()
        sp14.paragraph_format.space_before = Pt(4)
        sp14.paragraph_format.space_after = Pt(4)
        box14 = doc.add_table(rows=1, cols=1)
        box14.alignment = WD_TABLE_ALIGNMENT.CENTER
        box14.autofit = False
        box14.columns[0].width = Inches(9.12)
        set_table_borders(box14, sz="6", color="000000")
        set_row_height(box14.rows[0], 355.0, exact=True)
        c14 = box14.cell(0, 0)
        c14.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p14 = c14.paragraphs[0]
        p14.alignment = WD_ALIGN_PARAGRAPH.CENTER
        chart_rh_buf.seek(0)
        p14.add_run().add_picture(chart_rh_buf, width=Inches(7.9))

        doc.add_page_break()

    # OBSERVATION & CONCLUSION
    sec_num_obs = "6.0" if is_temp_only else "7.0"
    add_framed_divider_page([(f"{sec_num_obs} – OBSERVATION & FINAL CONCLUSION", 16.5, True)])

    p16_obs = doc.add_paragraph()
    p16_obs.paragraph_format.space_before = Pt(8)
    p16_obs.paragraph_format.space_after = Pt(14)
    r16_o = p16_obs.add_run("OBSERVATION")
    r16_o.bold = True
    r16_o.font.name, r16_o.font.size = "Calibri", Pt(16.5)

    p16_b1 = doc.add_paragraph()
    p16_b1.paragraph_format.left_indent = Inches(0.25)
    p16_b1.paragraph_format.space_after = Pt(36)
    if is_temp_only:
        obs_text = "•   During Mapping of Temperature, it was found to have been distributed uniformly within the specified limit."
    else:
        obs_text = "•   During Mapping of Temperature & RH, both these parameters were found to have been distributed uniformly within the specified limit."
    r16_b1 = p16_b1.add_run(obs_text)
    r16_b1.font.name, r16_b1.font.size = "Calibri", Pt(11)

    p16_fc = doc.add_paragraph()
    p16_fc.paragraph_format.space_before = Pt(0)
    p16_fc.paragraph_format.space_after = Pt(14)
    r16_fc = p16_fc.add_run("FINAL CONCLUSION")
    r16_fc.bold = True
    r16_fc.font.name, r16_fc.font.size = "Calibri", Pt(16.5)

    if is_temp_only:
        concl_items = [
            "•   Based upon the above Study It is Concluded That Temperature is within the acceptance criteria.",
            f"•   Minimum Temperature at Location {min_mkt_dl} & Maximum Temperature at Location {max_mkt_dl}.",
            f"•   Maximum Temperature Fluctuation at Location {max_t_fluc_dl}."
        ]
    else:
        concl_items = [
            "•   Based upon the above Study It is Concluded That Temperature & RH is within the acceptance criteria.",
            f"•   Minimum Temperature at Location {min_mkt_dl} & Maximum Temperature at Location {max_mkt_dl}.",
            f"•   Maximum Temperature Fluctuation at Location {max_t_fluc_dl}.",
            f"•   Minimum RH at Location {min_rh_avg_dl} & Maximum RH at Location {max_rh_avg_dl}.",
            f"•   Maximum RH Fluctuation at Location {max_rh_fluc_dl}."
        ]

    for c_item in concl_items:
        pc = doc.add_paragraph()
        pc.paragraph_format.left_indent = Inches(0.25)
        pc.paragraph_format.space_before = Pt(0)
        pc.paragraph_format.space_after = Pt(22)
        rc = pc.add_run(c_item)
        rc.font.name, rc.font.size = "Calibri", Pt(11)

    doc.add_page_break()

    # DIAGRAM LOCATIONS
    sec_num_diag = "7.0" if is_temp_only else "8.0"
    add_framed_divider_page([
        (f"{sec_num_diag} - DIAGRAM INDICATING MAXIMUM & MINIMUM AND MAXIMUM", 16.5, True),
        ("FLUCTUATION TEMPERATURE LOCATIONS:", 16.5, True)
    ])

    sp18 = doc.add_paragraph()
    sp18.paragraph_format.space_before = Pt(4)
    sp18.paragraph_format.space_after = Pt(4)
    box18 = doc.add_table(rows=1, cols=1)
    box18.alignment = WD_TABLE_ALIGNMENT.CENTER
    box18.autofit = False
    box18.columns[0].width = Inches(9.12)
    set_table_borders(box18, sz="6", color="000000")
    set_row_height(box18.rows[0], 355.0, exact=False)
    c18 = box18.cell(0, 0)
    c18.vertical_alignment = WD_ALIGN_VERTICAL.BOTTOM

    p18_top = c18.paragraphs[0]
    p18_top.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p18_top.paragraph_format.space_before = Pt(6)
    p18_top.paragraph_format.space_after = Pt(240)

    p18_leg = c18.add_paragraph()
    p18_leg.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p18_leg.paragraph_format.space_before = Pt(6)
    p18_leg.paragraph_format.space_after = Pt(12)

    r_g_dot = p18_leg.add_run("● ")
    r_g_dot.font.name, r_g_dot.font.size, r_g_dot.font.color.rgb = "Calibri", Pt(18), RGBColor(0x00, 0xB0, 0x50)
    r_g_txt = p18_leg.add_run("Minimum Temperature          ")
    r_g_txt.font.name, r_g_txt.font.size, r_g_txt.font.color.rgb = "Calibri", Pt(16), RGBColor(0x00, 0xB0, 0x50)

    r_r_dot = p18_leg.add_run("● ")
    r_r_dot.font.name, r_r_dot.font.size, r_r_dot.font.color.rgb = "Calibri", Pt(18), RGBColor(0xFF, 0x00, 0x00)
    r_r_txt = p18_leg.add_run("Maximum Temperature          ")
    r_r_txt.font.name, r_r_txt.font.size, r_r_txt.font.color.rgb = "Calibri", Pt(16), RGBColor(0xFF, 0x00, 0x00)

    r_p_dot = p18_leg.add_run("● ")
    r_p_dot.font.name, r_p_dot.font.size, r_p_dot.font.color.rgb = "Calibri", Pt(18), RGBColor(0x70, 0x30, 0xA0)
    r_p_txt = p18_leg.add_run("Max. Fluctuation")
    r_p_txt.font.name, r_p_txt.font.size, r_p_txt.font.color.rgb = "Calibri", Pt(16), RGBColor(0x70, 0x30, 0xA0)

    doc.add_page_break()

    # ATTACHMENTS & SIGN-OFF
    sec_num_att = "8.0" if is_temp_only else "9.0"
    sp19 = doc.add_paragraph()
    sp19.paragraph_format.space_before = Pt(4)
    sp19.paragraph_format.space_after = Pt(4)
    box19 = doc.add_table(rows=1, cols=1)
    box19.alignment = WD_TABLE_ALIGNMENT.CENTER
    box19.autofit = False
    box19.columns[0].width = Inches(9.12)
    set_table_borders(box19, sz="6", color="000000")
    set_row_height(box19.rows[0], 205.0, exact=True)
    c19 = box19.cell(0, 0)
    c19.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    raw_att_title = "1. Raw Data (Temperature)" if is_temp_only else "1. Raw Data (Temperature & RH)"
    att_lines = [f"{sec_num_att} – Attachments–", raw_att_title, "2. Calibration Certificates", "With Traceability"]
    for idx, al_str in enumerate(att_lines):
        p = c19.paragraphs[0] if idx == 0 else c19.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(al_str)
        r.bold = True
        r.font.name, r.font.size = "Calibri", Pt(16.5)
        r.font.highlight_color = WD_COLOR_INDEX.GRAY_25

    sp19_mid = doc.add_paragraph()
    sp19_mid.paragraph_format.space_before = Pt(6)
    sp19_mid.paragraph_format.space_after = Pt(6)

    sig_tbl = doc.add_table(rows=4, cols=5)
    sig_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    sig_tbl.autofit = False
    set_table_borders(sig_tbl, sz="4", color="000000")
    set_cell_margins(sig_tbl, top=20, bottom=20, left=45, right=45)

    sig_w = [1.82, 1.82, 1.82, 1.82, 1.84]
    for c_i, w_v in enumerate(sig_w): sig_tbl.columns[c_i].width = Inches(w_v)
    for r_i in range(4): set_row_height(sig_tbl.rows[r_i], 32.0, exact=True)

    sig_tbl.cell(3, 1).merge(sig_tbl.cell(3, 2))
    sig_tbl.cell(3, 3).merge(sig_tbl.cell(3, 4))

    format_cell(sig_tbl.cell(0, 0), "", bold=True, size=10.5, align=WD_ALIGN_PARAGRAPH.LEFT)
    format_cell(sig_tbl.cell(0, 1), "DONE BY", bold=True, size=10.5, align=WD_ALIGN_PARAGRAPH.LEFT)
    format_cell(sig_tbl.cell(0, 2), "CHECKED BY", bold=True, size=10.5, align=WD_ALIGN_PARAGRAPH.LEFT)
    format_cell(sig_tbl.cell(0, 3), "REVIEWED BY", bold=True, size=10.5, align=WD_ALIGN_PARAGRAPH.LEFT)
    format_cell(sig_tbl.cell(0, 4), "APPROVED BY", bold=True, size=10.5, align=WD_ALIGN_PARAGRAPH.LEFT)

    format_cell(sig_tbl.cell(1, 0), "SIGNATURE & DATE", bold=True, size=10.5, align=WD_ALIGN_PARAGRAPH.LEFT)
    format_cell(sig_tbl.cell(2, 0), "NAME", bold=True, size=10.5, align=WD_ALIGN_PARAGRAPH.LEFT)
    format_cell(sig_tbl.cell(3, 0), "ORGANIZATION", bold=True, size=10.5, align=WD_ALIGN_PARAGRAPH.LEFT)
    format_cell(sig_tbl.cell(3, 1), f"FOR {validated_by.strip().upper()}", bold=True, size=10.5, align=WD_ALIGN_PARAGRAPH.LEFT)
    format_cell(sig_tbl.cell(3, 3), f"FOR. {client_name.strip().upper()}", bold=True, size=10.5, align=WD_ALIGN_PARAGRAPH.LEFT)

    doc.save(docx_path)


# ==============================================================================
# GUI CLASS (CUSTOMTKINTER)
# ==============================================================================
class MappingApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Mapping Report Generator - Professional Edition")
        self.geometry("750x670")
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Tabview
        self.tabview = ctk.CTkTabview(self)
        self.tabview.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        
        self.tab_target = self.tabview.add("1. Target Details")
        self.tab_params = self.tabview.add("2. Mapping Parameters")
        self.tab_calib = self.tabview.add("3. Calibration & Output")

        self.build_target_tab()
        self.build_params_tab()
        self.build_calib_tab()

        # Bottom Frame (Progress & Action)
        self.bottom_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.bottom_frame.grid(row=1, column=0, padx=20, pady=(0, 20), sticky="ew")
        self.bottom_frame.grid_columnconfigure(0, weight=1)

        self.status_label = ctk.CTkLabel(self.bottom_frame, text="Ready", font=("Segoe UI", 12))
        self.status_label.grid(row=0, column=0, sticky="w")

        self.progress = ctk.CTkProgressBar(self.bottom_frame, mode="determinate")
        self.progress.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(5, 15))
        self.progress.set(0)

        self.btn_generate = ctk.CTkButton(self.bottom_frame, text="Generate Reports", height=40, font=("Segoe UI", 14, "bold"), command=self.start_generation_thread)
        self.btn_generate.grid(row=2, column=0, columnspan=2, sticky="ew")

    def build_target_tab(self):
        self.tab_target.grid_columnconfigure(1, weight=1)
        
        def create_row(parent, row, label_text, default_val, widget_type="entry", values=None):
            ctk.CTkLabel(parent, text=label_text).grid(row=row, column=0, padx=10, pady=10, sticky="w")
            if widget_type == "entry":
                widget = ctk.CTkEntry(parent)
                widget.insert(0, default_val)
            elif widget_type == "combo":
                widget = ctk.CTkComboBox(parent, values=values)
                widget.set(default_val)
            widget.grid(row=row, column=1, padx=10, pady=10, sticky="ew")
            return widget

        self.entry_client = create_row(self.tab_target, 0, "Client Name:", "M/S LIFESTAR PHARMACEUTICALS PVT. LTD.")
        self.entry_address = create_row(self.tab_target, 1, "Client Address:", "BIRGUNJ-32, PARSA, NEPAL.")
        self.combo_target = create_row(self.tab_target, 2, "Mapping Target:", "Equipment", "combo", ["Area", "Equipment"])
        self.entry_target_name = create_row(self.tab_target, 3, "Area / Equip Name:", "Stability Chamber")
        self.entry_equip_id = create_row(self.tab_target, 4, "Equip ID No. (If Equip):", "QAD/A/001")
        self.entry_location = create_row(self.tab_target, 5, "Location / Dept:", "Stability Room")
        self.entry_rep_no = create_row(self.tab_target, 6, "Report No.:", "HT/STC/2026/07/03")
        self.entry_val_by = create_row(self.tab_target, 7, "Validated By:", "Horizon Technology")

    def toggle_rh_fields(self):
        is_temp_only = self.checkbox_temp_only.get() == 1
        state = "disabled" if is_temp_only else "normal"
        self.entry_rmin.configure(state=state)
        self.entry_rmax.configure(state=state)
        self.entry_rcrit.configure(state=state)

    def build_params_tab(self):
        self.tab_params.grid_columnconfigure(1, weight=1)
        
        frame_dl = ctk.CTkFrame(self.tab_params, fg_color="transparent")
        frame_dl.grid(row=0, column=1, sticky="ew", padx=10, pady=10)
        self.entry_dl_count = ctk.CTkEntry(frame_dl, width=80); self.entry_dl_count.insert(0, "8"); self.entry_dl_count.pack(side="left", padx=(0, 10))
        self.entry_htdl = ctk.CTkEntry(frame_dl, width=80); self.entry_htdl.insert(0, "67"); self.entry_htdl.pack(side="left")
        ctk.CTkLabel(self.tab_params, text="Loggers Count / Start HT-DL:").grid(row=0, column=0, sticky="w", padx=10)

        frame_dt = ctk.CTkFrame(self.tab_params, fg_color="transparent")
        frame_dt.grid(row=1, column=1, sticky="ew", padx=10, pady=10)
        self.entry_date = ctk.CTkEntry(frame_dt, width=100); self.entry_date.insert(0, "24/07/26"); self.entry_date.pack(side="left", padx=(0, 10))
        self.entry_time = ctk.CTkEntry(frame_dt, width=100); self.entry_time.insert(0, "10:00:00"); self.entry_time.pack(side="left")
        ctk.CTkLabel(self.tab_params, text="Start Date & Time:").grid(row=1, column=0, sticky="w", padx=10)

        ctk.CTkLabel(self.tab_params, text="Next Due Date:").grid(row=2, column=0, sticky="w", padx=10)
        self.entry_due_date = ctk.CTkEntry(self.tab_params); self.entry_due_date.insert(0, "28/07/27"); self.entry_due_date.grid(row=2, column=1, sticky="ew", padx=10, pady=10)

        frame_dur = ctk.CTkFrame(self.tab_params, fg_color="transparent")
        frame_dur.grid(row=3, column=1, sticky="ew", padx=10, pady=10)
        self.entry_duration = ctk.CTkEntry(frame_dur, width=80); self.entry_duration.insert(0, "24"); self.entry_duration.pack(side="left", padx=(0, 10))
        self.entry_freq = ctk.CTkEntry(frame_dur, width=80); self.entry_freq.insert(0, "10"); self.entry_freq.pack(side="left")
        ctk.CTkLabel(self.tab_params, text="Duration (Hrs) / Freq (Mins):").grid(row=3, column=0, sticky="w", padx=10)

        frame_t = ctk.CTkFrame(self.tab_params, fg_color="transparent")
        frame_t.grid(row=4, column=1, sticky="ew", padx=10, pady=10)
        self.entry_tmin = ctk.CTkEntry(frame_t, width=80); self.entry_tmin.insert(0, "29.2"); self.entry_tmin.pack(side="left", padx=(0, 10))
        self.entry_tmax = ctk.CTkEntry(frame_t, width=80); self.entry_tmax.insert(0, "30.9"); self.entry_tmax.pack(side="left")
        ctk.CTkLabel(self.tab_params, text="Temp Range (Min / Max):").grid(row=4, column=0, sticky="w", padx=10)

        frame_r = ctk.CTkFrame(self.tab_params, fg_color="transparent")
        frame_r.grid(row=5, column=1, sticky="ew", padx=10, pady=10)
        self.entry_rmin = ctk.CTkEntry(frame_r, width=80); self.entry_rmin.insert(0, "73.9"); self.entry_rmin.pack(side="left", padx=(0, 10))
        self.entry_rmax = ctk.CTkEntry(frame_r, width=80); self.entry_rmax.insert(0, "76.1"); self.entry_rmax.pack(side="left")
        ctk.CTkLabel(self.tab_params, text="RH Range (Min / Max):").grid(row=5, column=0, sticky="w", padx=10)

        frame_crit = ctk.CTkFrame(self.tab_params, fg_color="transparent")
        frame_crit.grid(row=6, column=1, sticky="ew", padx=10, pady=10)
        self.entry_tcrit = ctk.CTkEntry(frame_crit, width=120); self.entry_tcrit.insert(0, "30°C ± 2 °C"); self.entry_tcrit.pack(side="left", padx=(0, 10))
        self.entry_rcrit = ctk.CTkEntry(frame_crit, width=120); self.entry_rcrit.insert(0, "75 % ± 5 %"); self.entry_rcrit.pack(side="left")
        ctk.CTkLabel(self.tab_params, text="Acceptance (Temp / RH):").grid(row=6, column=0, sticky="w", padx=10)
        
        self.entry_hotspot = ctk.CTkEntry(self.tab_params)
        self.entry_hotspot.insert(0, "8")
        self.entry_hotspot.grid(row=7, column=1, sticky="ew", padx=10, pady=10)
        ctk.CTkLabel(self.tab_params, text="Hotspot Logger (Optional):").grid(row=7, column=0, sticky="w", padx=10)

        self.combo_pattern = ctk.CTkComboBox(self.tab_params, values=[
            "Standard Cycle (2 PM Peak)", 
            "Shifted Cycle (Random Peak)", 
            "High Noise Cycle", 
            "Low Amplitude Cycle"
        ])
        self.combo_pattern.set("Standard Cycle (2 PM Peak)")
        self.combo_pattern.grid(row=8, column=1, sticky="ew", padx=10, pady=10)
        ctk.CTkLabel(self.tab_params, text="Wave Profile:").grid(row=8, column=0, sticky="w", padx=10)

        self.checkbox_temp_only = ctk.CTkCheckBox(
            self.tab_params, 
            text="Temperature Mapping Only (Exclude Relative Humidity)",
            command=self.toggle_rh_fields
        )
        self.checkbox_temp_only.grid(row=9, column=1, sticky="w", padx=10, pady=10)

    def build_calib_tab(self):
        self.tab_calib.grid_columnconfigure(1, weight=1)

        frame_dldate = ctk.CTkFrame(self.tab_calib, fg_color="transparent")
        frame_dldate.grid(row=0, column=1, sticky="ew", padx=10, pady=10)
        self.entry_dl_cal = ctk.CTkEntry(frame_dldate, width=120); self.entry_dl_cal.insert(0, "10/01/2026"); self.entry_dl_cal.pack(side="left", padx=(0, 10))
        self.entry_dl_due = ctk.CTkEntry(frame_dldate, width=120); self.entry_dl_due.insert(0, "09/01/2027"); self.entry_dl_due.pack(side="left")
        ctk.CTkLabel(self.tab_calib, text="Logger Cal / Due Date:").grid(row=0, column=0, sticky="w", padx=10)

        self.entry_dl_cert = ctk.CTkEntry(self.tab_calib)
        self.entry_dl_cert.insert(0, "RK230125/N002 & RK230125/N003")
        self.entry_dl_cert.grid(row=1, column=1, sticky="ew", padx=10, pady=10)
        ctk.CTkLabel(self.tab_calib, text="Logger Certificate No.:").grid(row=1, column=0, sticky="w", padx=10)

        frame_mst1 = ctk.CTkFrame(self.tab_calib, fg_color="transparent")
        frame_mst1.grid(row=2, column=1, sticky="ew", padx=10, pady=10)
        self.entry_mst_id = ctk.CTkEntry(frame_mst1, width=100); self.entry_mst_id.insert(0, "RK-STD-103"); self.entry_mst_id.pack(side="left", padx=(0, 10))
        self.entry_mst_cert = ctk.CTkEntry(frame_mst1, width=150); self.entry_mst_cert.insert(0, "SAN/25/1439/TH/1"); self.entry_mst_cert.pack(side="left")
        ctk.CTkLabel(self.tab_calib, text="Master ID / Cert No.:").grid(row=2, column=0, sticky="w", padx=10)

        frame_mst2 = ctk.CTkFrame(self.tab_calib, fg_color="transparent")
        frame_mst2.grid(row=3, column=1, sticky="ew", padx=10, pady=10)
        self.entry_mst_cal = ctk.CTkEntry(frame_mst2, width=120); self.entry_mst_cal.insert(0, "14 November 2025"); self.entry_mst_cal.pack(side="left", padx=(0, 10))
        self.entry_mst_due = ctk.CTkEntry(frame_mst2, width=120); self.entry_mst_due.insert(0, "13 November 2026"); self.entry_mst_due.pack(side="left")
        ctk.CTkLabel(self.tab_calib, text="Master Cal / Due Date:").grid(row=3, column=0, sticky="w", padx=10)

        frame_out = ctk.CTkFrame(self.tab_calib, fg_color="transparent")
        frame_out.grid(row=4, column=1, sticky="ew", padx=10, pady=10)
        self.entry_folder = ctk.CTkEntry(frame_out)
        self.entry_folder.insert(0, r"D:\New folder")
        self.entry_folder.pack(side="left", fill="x", expand=True, padx=(0, 10))
        btn_browse = ctk.CTkButton(frame_out, text="Browse", width=60, command=self.browse_folder)
        btn_browse.pack(side="right")
        ctk.CTkLabel(self.tab_calib, text="Output Folder:").grid(row=4, column=0, sticky="w", padx=10)

    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.entry_folder.delete(0, "end")
            self.entry_folder.insert(0, folder)

    def set_progress(self, val, text):
        self.progress.set(val)
        self.status_label.configure(text=text)
        self.update_idletasks()

    def start_generation_thread(self):
        self.btn_generate.configure(state="disabled")
        self.set_progress(0, "Initializing data...")
        threading.Thread(target=self.generate_report_logic, daemon=True).start()

    def generate_report_logic(self):
        try:
            # 1. Fetch values
            is_temp_only = self.checkbox_temp_only.get() == 1

            client_name = self.entry_client.get()
            client_address = self.entry_address.get()
            target_type = self.combo_target.get()
            target_name = self.entry_target_name.get()
            equip_id = self.entry_equip_id.get()
            location_name = self.entry_location.get()
            report_no = self.entry_rep_no.get()
            validated_by = self.entry_val_by.get()

            dl_count = int(self.entry_dl_count.get())
            start_ht_dl = int(self.entry_htdl.get())

            start_date_str = self.entry_date.get().strip()
            start_time_str = self.entry_time.get().strip()
            next_due_str = self.entry_due_date.get().strip()

            duration_hrs = int(self.entry_duration.get())
            freq_mins = int(self.entry_freq.get())

            temp_min = float(self.entry_tmin.get())
            temp_max = float(self.entry_tmax.get())
            
            if not is_temp_only:
                rh_min = float(self.entry_rmin.get())
                rh_max = float(self.entry_rmax.get())
                rh_criteria = self.entry_rcrit.get().strip()
            else:
                rh_min, rh_max, rh_criteria = 0.0, 0.0, "N/A"

            temp_criteria = self.entry_tcrit.get().strip()

            dl_cal_date = self.entry_dl_cal.get().strip()
            dl_due_date = self.entry_dl_due.get().strip()
            dl_cert_no = self.entry_dl_cert.get().strip()

            master_id = self.entry_mst_id.get().strip()
            master_cert = self.entry_mst_cert.get().strip()
            master_cal_date = self.entry_mst_cal.get().strip()
            master_due_date = self.entry_mst_due.get().strip()

            hotspot_str = self.entry_hotspot.get().strip()
            data_pattern = self.combo_pattern.get()
            output_folder = self.entry_folder.get().strip()
            
            safe_name = sanitize_filename(target_name)
            os.makedirs(output_folder, exist_ok=True)
            
            self.set_progress(0.1, "Simulating Mapping Data...")
            time.sleep(0.2)

            parsed_start_d = parse_flexible_date(start_date_str)
            parsed_time = datetime.datetime.strptime(start_time_str, "%H:%M:%S").time()
            start_date = datetime.datetime.combine(parsed_start_d.date(), parsed_time)
            next_due = parse_flexible_date(next_due_str) if next_due_str else start_date.replace(year=start_date.year + 1)

            num_readings = int((duration_hrs * 60) / freq_mins) + 1
            timestamps = [start_date + datetime.timedelta(minutes=freq_mins * i) for i in range(num_readings)]
            end_date = timestamps[-1]
            dates_and_times = [t.strftime("%d/%m/%y %H:%M:%S") for t in timestamps]
            sr_nos = list(range(1, num_readings + 1))

            if hotspot_str:
                hotspot_idx = int(hotspot_str) - 1
                if hotspot_idx < 0 or hotspot_idx >= dl_count: raise ValueError(f"Hotspot Logger must be between 1 and {dl_count}")
            else:
                hotspot_idx = np.random.randint(0, dl_count)

            # Excel Workbook Setup
            wb = Workbook()
            ws = wb.active
            ws.title = "Mapping Data"
            font_reg = Font(name="Calibri", size=11, bold=False)
            font_bold = Font(name="Calibri", size=11, bold=True)
            align_left = Alignment(horizontal="left")
            align_center = Alignment(horizontal="center")
            align_center_vcenter = Alignment(horizontal="center", vertical="center")
            thin = Side(style="thin")
            border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
            border_l_tb = Border(left=thin, top=thin, bottom=thin)
            border_tb = Border(top=thin, bottom=thin)
            border_r_tb = Border(right=thin, top=thin, bottom=thin)
            fill_summary = PatternFill(fill_type="solid", fgColor=Color(theme=8, tint=0.7999816888943144))

            ws.column_dimensions["A"].width = 8.88671875
            ws.column_dimensions["B"].width = 21.44140625
            ws.column_dimensions["C"].width = 8.88671875
            
            cols_per_dl = 1 if is_temp_only else 2
            total_cols = 2 + (dl_count * cols_per_dl)

            target_display = f"Equipment Name : {target_name.strip()} ({equip_id.strip()})" if target_type == "Equipment" else f"Area Name : {target_name.strip()}"
            top_headers = [client_name, client_address, target_display, f"Temp.Mapping Date : {start_date.strftime('%d/%m/%y')}", f"Next Due Date : {next_due.strftime('%d/%m/%y')}"]
            for r_idx, text_val in enumerate(top_headers, start=1):
                c = ws.cell(row=r_idx, column=1, value=text_val)
                c.font, c.alignment = font_bold, align_left

            ws.cell(row=6, column=1, value="ID Number")
            ws.merge_cells(start_row=6, start_column=1, end_row=7, end_column=2)
            for r in (6, 7):
                for c_idx in (1, 2):
                    cell = ws.cell(row=r, column=c_idx)
                    cell.font = font_bold if (r == 6 and c_idx == 1) else font_reg
                    cell.border = border_all
                    if r == 6 and c_idx == 1: cell.alignment = align_center_vcenter

            c_sr = ws.cell(row=8, column=1, value="Sr.No")
            c_sr.font, c_sr.alignment, c_sr.border = font_bold, align_center, border_all
            c_dt = ws.cell(row=8, column=2, value="Date/time")
            c_dt.font, c_dt.border = font_bold, border_all

            current_ht_dl = start_ht_dl
            for i in range(dl_count):
                col_start = 3 + (i * cols_per_dl)
                col_end = col_start + cols_per_dl - 1

                c6 = ws.cell(row=6, column=col_start, value=f"DL-{i+1:02d}")
                c6.font, c6.alignment = font_bold, align_center
                if not is_temp_only:
                    ws.merge_cells(start_row=6, start_column=col_start, end_row=6, end_column=col_end)
                    ws.cell(row=6, column=col_start).border, ws.cell(row=6, column=col_end).border = border_all, border_all
                    ws.cell(row=6, column=col_end).font = font_reg
                else:
                    c6.border = border_all

                c7 = ws.cell(row=7, column=col_start, value=f"HT-DL-{current_ht_dl}")
                c7.font, c7.alignment = font_bold, align_center
                if not is_temp_only:
                    ws.merge_cells(start_row=7, start_column=col_start, end_row=7, end_column=col_end)
                    ws.cell(row=7, column=col_start).border, ws.cell(row=7, column=col_end).border = border_all, border_all
                    ws.cell(row=7, column=col_end).font = font_reg
                else:
                    c7.border = border_all
                current_ht_dl += 1

                ch_title = f"CH {i+1:02d}" if is_temp_only else f"CH {i*2 + 1:02d}"
                c8_t = ws.cell(row=8, column=col_start, value=ch_title)
                c8_t.font, c8_t.alignment, c8_t.border = font_bold, align_center, border_all
                
                if not is_temp_only:
                    c8_rh = ws.cell(row=8, column=col_end, value=f"CH {i*2 + 2:02d}")
                    c8_rh.font, c8_rh.alignment, c8_rh.border = font_bold, align_center, border_all

            all_temps, all_rhs = [[] for _ in range(dl_count)], [[] for _ in range(dl_count)]
            temp_span = temp_max - temp_min
            rh_span = rh_max - rh_min if not is_temp_only else 0
            
            base_temps, base_rhs = [], []
            for d in range(dl_count):
                if d == hotspot_idx: 
                    base_temps.append(temp_max - (temp_span * 0.15) - np.random.uniform(0, temp_span * 0.05))
                else: 
                    base_temps.append(temp_min + (temp_span * 0.3) + np.random.uniform(-temp_span * 0.1, temp_span * 0.15))
                if not is_temp_only:
                    base_rhs.append((rh_min + (rh_span / 2)) + np.random.uniform(-rh_span * 0.15, rh_span * 0.15))

            # Set wave parameters based on selection
            if data_pattern == "Shifted Cycle (Random Peak)":
                peak_hour = np.random.uniform(0, 24)
                amp_mult, noise_mult = 0.15, 0.05
            elif data_pattern == "High Noise Cycle":
                peak_hour = 14
                amp_mult, noise_mult = 0.15, 0.12
            elif data_pattern == "Low Amplitude Cycle":
                peak_hour = 14
                amp_mult, noise_mult = 0.08, 0.03
            else: # Standard Cycle
                peak_hour = 14
                amp_mult, noise_mult = 0.15, 0.05

            current_row = 9
            
            for i in range(num_readings):
                if i % 100 == 0: self.set_progress(0.1 + (0.3 * (i/num_readings)), f"Writing row {i} of {num_readings}...")
                
                c_s = ws.cell(row=current_row, column=1, value=sr_nos[i])
                c_s.font, c_s.alignment, c_s.border = font_reg, align_center, border_all
                c_d = ws.cell(row=current_row, column=2, value=dates_and_times[i])
                c_d.font, c_d.border = font_reg, border_all

                current_time = timestamps[i]
                hour_decimal = current_time.hour + current_time.minute / 60.0
                
                time_based_fluctuation_t = np.cos((hour_decimal - peak_hour) * 2 * np.pi / 24) * (temp_span * amp_mult)
                if not is_temp_only:
                    time_based_fluctuation_rh = -np.cos((hour_decimal - peak_hour) * 2 * np.pi / 24) * (rh_span * amp_mult)

                for d in range(dl_count):
                    raw_temp = np.clip(base_temps[d] + time_based_fluctuation_t + np.random.normal(0, temp_span * noise_mult), temp_min, temp_max)
                    if not is_temp_only:
                        raw_rh = np.clip(base_rhs[d] + time_based_fluctuation_rh + np.random.normal(0, rh_span * noise_mult), rh_min, rh_max)
                        
                    temp = round(float(raw_temp), 1)
                    all_temps[d].append(temp)

                    col_start = 3 + (d * cols_per_dl)
                    ct = ws.cell(row=current_row, column=col_start, value=temp)
                    ct.font, ct.alignment, ct.border, ct.number_format = font_reg, align_center, border_all, "0.0"
                    
                    if not is_temp_only:
                        rh = round(float(raw_rh), 1)
                        all_rhs[d].append(rh)
                        crh = ws.cell(row=current_row, column=col_start + 1, value=rh)
                        crh.font, crh.alignment, crh.border, crh.number_format = font_reg, align_center, border_all, "0.0"

                current_row += 1

            last_data_row = current_row - 1

            # Excel Footer Calcs
            self.set_progress(0.4, "Calculating MKT & Summaries...")
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=total_cols)
            for c_idx in range(1, total_cols + 1):
                cell = ws.cell(row=current_row, column=c_idx)
                cell.font, cell.alignment = font_reg, align_center
                if c_idx == 1: cell.border = border_l_tb
                elif c_idx == total_cols: cell.border = border_r_tb
                else: cell.border = border_tb
            current_row += 1

            calc_labels = ["Average", "Min.", "Max.", "Fluctuation ", "MKT"]
            for offset, label in enumerate(calc_labels):
                r = current_row + offset
                ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
                c1 = ws.cell(row=r, column=1, value=label)
                c1.font, c1.alignment, c1.fill, c1.border = font_bold, align_center, fill_summary, border_l_tb
                c2 = ws.cell(row=r, column=2)
                c2.font, c2.alignment, c2.fill, c2.border = font_bold, align_center, fill_summary, border_r_tb

            mkt_vals = []
            for d in range(dl_count):
                t_arr = all_temps[d]
                col_start = 3 + (d * cols_per_dl)
                mkt_v = round(float(calc_mkt(t_arr)), 1)
                mkt_vals.append(mkt_v)
                t_vals = [round(float(np.mean(t_arr)), 1), round(float(np.min(t_arr)), 1), round(float(np.max(t_arr)), 1), round(float(np.max(t_arr) - np.min(t_arr)), 1), mkt_v]
                
                if not is_temp_only:
                    rh_arr = all_rhs[d]
                    rh_vals = [round(float(np.mean(rh_arr)), 1), round(float(np.min(rh_arr)), 1), round(float(np.max(rh_arr)), 1), round(float(np.max(rh_arr) - np.min(rh_arr)), 1), "NA"]

                for offset in range(5):
                    r = current_row + offset
                    ct = ws.cell(row=r, column=col_start, value=t_vals[offset])
                    ct.font, ct.alignment, ct.border, ct.number_format = font_reg, align_center, border_all, "0.0"
                    if not is_temp_only:
                        crh = ws.cell(row=r, column=col_start + 1, value=rh_vals[offset])
                        crh.font, crh.alignment, crh.border, crh.number_format = font_reg, align_center, border_all, "0.0"

            current_row += 7

            for idx, h_text in enumerate(calc_labels):
                c = ws.cell(row=current_row, column=2 + idx, value=h_text)
                c.font, c.alignment, c.border = font_reg, align_left, border_all
            current_row += 1

            for d in range(dl_count):
                t_arr = all_temps[d]
                t_vals = [round(float(np.mean(t_arr)), 1), round(float(np.min(t_arr)), 1), round(float(np.max(t_arr)), 1), round(float(np.max(t_arr) - np.min(t_arr)), 1), mkt_vals[d]]
                
                for idx, val in enumerate(t_vals):
                    c = ws.cell(row=current_row, column=2 + idx, value=val)
                    c.font, c.alignment, c.border, c.number_format = font_reg, align_center, border_all, "0.0"
                current_row += 1

                if not is_temp_only:
                    rh_arr = all_rhs[d]
                    rh_vals = [round(float(np.mean(rh_arr)), 1), round(float(np.min(rh_arr)), 1), round(float(np.max(rh_arr)), 1), round(float(np.max(rh_arr) - np.min(rh_arr)), 1), "NA"]
                    for idx, val in enumerate(rh_vals):
                        c = ws.cell(row=current_row, column=2 + idx, value=val)
                        c.font, c.alignment, c.border, c.number_format = font_reg, align_center, border_all, "0.0"
                    current_row += 1

            self.set_progress(0.45, "Rendering Excel Charts...")
            chart_t = LineChart()
            dates_ref = Reference(ws, min_col=2, min_row=9, max_row=last_data_row)
            for d in range(dl_count):
                chart_t.add_data(Reference(ws, min_col=3 + (d * cols_per_dl), min_row=8, max_row=last_data_row), titles_from_data=True)
            chart_t.set_categories(dates_ref)
            style_line_chart(chart_t, f"{target_name.strip()} - Temperature Trend", "Temperature (°C)", is_rh=False, y_max=25.0 if temp_max <= 25.0 else None)
            ws.add_chart(chart_t)
            chart_t.anchor = OneCellAnchor(_from=AnchorMarker(col=max(26, total_cols + 2), colOff=317778, row=max(0, last_data_row - 7), rowOff=124768), ext=XDRPositiveSize2D(cx=8118231, cy=3731884))

            if not is_temp_only:
                chart_rh = LineChart()
                for d in range(dl_count):
                    chart_rh.add_data(Reference(ws, min_col=4 + (d * cols_per_dl), min_row=8, max_row=last_data_row), titles_from_data=True)
                chart_rh.set_categories(dates_ref)
                style_line_chart(chart_rh, f"{target_name.strip()} - Relative Humidity Trend", "Relative Humidity (%)", is_rh=True, y_min=49.0 if rh_min >= 49.0 else None)
                ws.add_chart(chart_rh)
                chart_rh.anchor = OneCellAnchor(_from=AnchorMarker(col=max(26, total_cols + 2) + 1, colOff=278423, row=max(0, last_data_row - 7) + 32, rowOff=102576), ext=XDRPositiveSize2D(cx=8005960, cy=3516924))
            
            excel_filename = f"{safe_name}_Mapping_Data.xlsx"
            wb.save(os.path.join(output_folder, excel_filename))
            
            # Matplotlib Charts for Word
            self.set_progress(0.5, "Rendering Image Charts for Word...")
            if is_temp_only:
                t_ch_labels = [f"CH {d+1:02d}" for d in range(dl_count)]
                chart_rh_buf = None
            else:
                t_ch_labels = [f"CH {d*2+1:02d}" for d in range(dl_count)]
                rh_ch_labels = [f"CH {d*2+2:02d}" for d in range(dl_count)]
                chart_rh_buf = render_excel_chart_png(dates_and_times, all_rhs, rh_ch_labels, f"{target_name.strip()} - Relative Humidity Trend", "Relative Humidity (%)")

            chart_t_buf = render_excel_chart_png(dates_and_times, all_temps, t_ch_labels, f"{target_name.strip()} - Temperature Trend", "Temperature (°C)")

            # Word Doc 1
            self.set_progress(0.65, "Building Raw Data Document...")
            raw_docx_filename = f"{safe_name}_RAW_DATA.docx"
            generate_raw_data_docx(
                os.path.join(output_folder, raw_docx_filename), client_name, client_address, target_type, target_name, equip_id,
                start_date, next_due, dl_count, start_ht_dl, sr_nos, dates_and_times, all_temps, all_rhs, mkt_vals, is_temp_only
            )
            
            # Word Doc 2
            self.set_progress(0.85, "Building Final Report Document...")
            report_docx_filename = f"{safe_name}_REPORT.docx"
            generate_mapping_report_docx(
                os.path.join(output_folder, report_docx_filename), client_name, client_address, target_type, target_name, equip_id, location_name,
                report_no, validated_by, start_date, end_date, next_due, duration_hrs, freq_mins, dl_count, start_ht_dl,
                temp_criteria, rh_criteria, dl_cal_date, dl_due_date, dl_cert_no, master_id, master_cert, master_cal_date, master_due_date,
                all_temps, all_rhs, mkt_vals, chart_t_buf, chart_rh_buf, is_temp_only
            )

            self.set_progress(1.0, "Complete!")
            self.after(100, lambda: messagebox.showinfo("Success", f"All files successfully generated in:\n{output_folder}"))
            
        except Exception as e:
            self.after(100, lambda: messagebox.showerror("Error", str(e)))
            self.set_progress(0, "Failed.")
        finally:
            self.btn_generate.configure(state="normal")

if __name__ == "__main__":
    app = MappingApp()
    app.mainloop()
