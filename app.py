import re
import io
import zipfile
import datetime
import numpy as np
import streamlit as st

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
# DOCX GENERATORS (SAVES TO STREAM / BUFFER OR PATH)
# ==============================================================================
def generate_raw_data_docx(
    docx_target, client_name, client_address, target_type, target_name, equip_id,
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
            set_cell_width_dxa(tbl.cell(1, c_start), ch_w * 
