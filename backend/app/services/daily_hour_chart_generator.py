import io

import matplotlib.pyplot as plt
import pandas as pd
from docx import Document
from docx.enum.table import (
    WD_CELL_VERTICAL_ALIGNMENT,
    WD_ROW_HEIGHT_RULE,
    WD_TABLE_ALIGNMENT,
)
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt
from app.services.report_layout import (
    FONT_NAME,
    SECTION_TITLE_SIZE,
    SUBHEADING_SIZE,
    TABLE_HEADER_SIZE,
    TABLE_BODY_SIZE,
)

TABLE_WIDTHS = [1316, 1153, 789, 1260, 1248]
COMPACT_TABLE_WIDTHS = TABLE_WIDTHS


def set_row_height(row, height):
    row.height = Inches(height)
    row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST


def set_row_height_points(row, points):
    row.height = Pt(points)
    row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST


def set_cell_width(cell, width):
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_w = tc_pr.find(qn("w:tcW"))

    if tc_w is None:
        tc_w = OxmlElement("w:tcW")
        tc_pr.append(tc_w)

    tc_w.set(qn("w:w"), str(width))
    tc_w.set(qn("w:type"), "dxa")


def set_fixed_table_layout(table):
    tbl_pr = table._tbl.tblPr
    tbl_layout = tbl_pr.find(qn("w:tblLayout"))

    if tbl_layout is None:
        tbl_layout = OxmlElement("w:tblLayout")
        tbl_pr.append(tbl_layout)

    tbl_layout.set(qn("w:type"), "fixed")


def set_table_grid(table, widths):
    tbl = table._tbl
    existing_grid = tbl.find(qn("w:tblGrid"))

    if existing_grid is not None:
        tbl.remove(existing_grid)

    tbl_grid = OxmlElement("w:tblGrid")

    for width in widths:
        grid_col = OxmlElement("w:gridCol")
        grid_col.set(qn("w:w"), str(width))
        tbl_grid.append(grid_col)

    tbl.insert(0, tbl_grid)


def set_table_borders(table, size="8"):
    tbl_pr = table._tbl.tblPr
    borders = tbl_pr.find(qn("w:tblBorders"))

    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)

    for border_name in ["top", "left", "bottom", "right", "insideH", "insideV"]:
        border = borders.find(qn(f"w:{border_name}"))

        if border is None:
            border = OxmlElement(f"w:{border_name}")
            borders.append(border)

        border.set(qn("w:val"), "single")
        border.set(qn("w:sz"), size)
        border.set(qn("w:space"), "0")
        border.set(qn("w:color"), "000000")


def set_cell_borders(cell, value="single", size="6"):
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.find(qn("w:tcBorders"))

    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)

    for border_name in ["top", "left", "bottom", "right", "insideH", "insideV"]:
        border = borders.find(qn(f"w:{border_name}"))

        if border is None:
            border = OxmlElement(f"w:{border_name}")
            borders.append(border)

        border.set(qn("w:val"), value)
        if value != "nil":
            border.set(qn("w:sz"), size)
            border.set(qn("w:space"), "0")
            border.set(qn("w:color"), "000000")


def remove_table_borders(table):
    tbl_pr = table._tbl.tblPr
    borders = tbl_pr.find(qn("w:tblBorders"))

    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)

    for border_name in ["top", "left", "bottom", "right", "insideH", "insideV"]:
        border = borders.find(qn(f"w:{border_name}"))

        if border is None:
            border = OxmlElement(f"w:{border_name}")
            borders.append(border)

        border.set(qn("w:val"), "nil")


def set_table_indent(table, width):
    tbl_pr = table._tbl.tblPr
    tbl_ind = tbl_pr.find(qn("w:tblInd"))

    if tbl_ind is None:
        tbl_ind = OxmlElement("w:tblInd")
        tbl_pr.append(tbl_ind)

    tbl_ind.set(qn("w:w"), str(width))
    tbl_ind.set(qn("w:type"), "dxa")


def apply_widths_to_all_cells(table, widths):
    for row in table.rows:
        for i, cell in enumerate(row.cells):
            set_cell_width(cell, widths[i])


def remove_empty_cell_paragraphs(cell):
    for paragraph in cell.paragraphs:
        if paragraph.text:
            continue

        paragraph._element.getparent().remove(paragraph._element)


def set_cell_text(cell, text):
    cell.text = str(text)


def format_count(value):
    if pd.isna(value):
        return ""

    return f"{int(value):,}"


def build_daily_hour_chart_data(daily_df: pd.DataFrame) -> pd.DataFrame:
    """
    Build data used for Daily Hourly Data table and chart.

    N = D + S
    M = M
    Q = H - C
    X = N + M
    """

    df = daily_df[daily_df["DATE"].astype(str).str.lower() != "totals"].copy()

    chart_df = pd.DataFrame()

    chart_df["TIME"] = df["TIME"]
    chart_df["N"] = df["D"] + df["S"]
    chart_df["M"] = df["M"]
    chart_df["Q"] = df["Q"]
    chart_df["X"] = chart_df["N"] + chart_df["M"]

    return chart_df


def create_daily_hour_chart_image(chart_df):
    buffer = io.BytesIO()

    fig = plt.figure(figsize=(6.4, 4.18), dpi=150)
    ax = fig.add_axes([0.08, 0.29, 0.89, 0.58])
    fig.patch.set_facecolor("white")
    fig.patch.set_edgecolor("#D9D9D9")
    fig.patch.set_linewidth(1.0)

    ax.set_ylim(-12, 320)
    ax.set_yticks(range(0, 301, 50))
    ax.plot(chart_df["TIME"], chart_df["N"], label="N=(D+S)", color="#4472C4", linewidth=2.2)
    ax.plot(chart_df["TIME"], chart_df["M"], label="(M)", color="#ED7D31", linewidth=2.2)
    ax.plot(chart_df["TIME"], chart_df["Q"], label="Q= H-C", color="#A5A5A5", linewidth=2.6)
    ax.plot(chart_df["TIME"], chart_df["X"], label="X= (D+S+M)", color="#FFC000", linewidth=2.6)

    ax.set_title(
        "Graph on Trucks Weighed per Hour",
        fontsize=14,
        fontweight="bold",
        color="#595959",
        pad=12,
    )
    ax.tick_params(axis="x", labelrotation=48, labelsize=7, colors="#595959")
    ax.tick_params(axis="y", labelsize=8, colors="#595959")
    ax.grid(axis="y", color="#D9D9D9", linewidth=0.9)
    ax.grid(axis="x", visible=False)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color("#BFBFBF")
    ax.spines["bottom"].set_color("#BFBFBF")
    ax.legend(
        fontsize=7,
        loc="center",
        bbox_to_anchor=(0.5, 0.08),
        bbox_transform=fig.transFigure,
        ncol=4,
        frameon=False,
        handlelength=2.4,
        handletextpad=0.3,
        columnspacing=1.1,
    )

    fig.savefig(buffer, format="png", dpi=150)
    plt.close()

    buffer.seek(0)
    return buffer


def style_table_cell(
    cell,
    font_size=10,
    bold=False,
    align=WD_ALIGN_PARAGRAPH.CENTER,
):
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    for paragraph in cell.paragraphs:
        paragraph.alignment = align
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)

        for run in paragraph.runs:
            run.font.name = "Arial"
            run.font.size = Pt(font_size)
            run.bold = bold

            r_pr = run._element.get_or_add_rPr()
            r_fonts = r_pr.rFonts
            if r_fonts is None:
                r_fonts = OxmlElement("w:rFonts")
                r_pr.append(r_fonts)

            r_fonts.set(qn("w:ascii"), "Arial")
            r_fonts.set(qn("w:hAnsi"), "Arial")


def add_vertical_spacer(doc, points):
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(points)

    run = paragraph.add_run("")
    run.font.size = Pt(1)


def add_daily_hour_chart_section(doc: Document, daily_df):
    chart_df = build_daily_hour_chart_data(daily_df)

    heading = doc.add_paragraph()
    heading.paragraph_format.space_before = Pt(8)
    heading.paragraph_format.space_after = Pt(0)
    run = heading.add_run("2. DAILY HOURLY DATA")
    run.bold = True
    run.font.name = "Arial"
    run.font.size = Pt(11)

    small_table = doc.add_table(rows=28, cols=5)
    small_table.autofit = False
    small_table.allow_autofit = False
    small_table.alignment = WD_TABLE_ALIGNMENT.LEFT

    set_fixed_table_layout(small_table)
    set_table_grid(small_table, COMPACT_TABLE_WIDTHS)
    set_table_indent(small_table, 0)
    set_table_borders(small_table, size="6")

    header_row = small_table.rows[0]
    formula_row = small_table.rows[1]
    duplicate_formula_row = small_table.rows[2]

    set_row_height_points(header_row, 33.7)
    set_row_height_points(formula_row, 21.55)
    set_row_height_points(duplicate_formula_row, 13.8)

    time_cell = header_row.cells[0].merge(duplicate_formula_row.cells[0])

    headers = [
        "Time",
        "Multideck\nweighed",
        "Manually ",
        "HSWIM\nCLEARED",
        "Total\nweighed",
    ]

    for i, header in enumerate(headers):
        cell = time_cell if i == 0 else header_row.cells[i]
        set_cell_text(cell, header)
        set_cell_width(cell, COMPACT_TABLE_WIDTHS[i])
        style_table_cell(cell, font_size=10, bold=True)
        set_cell_borders(cell, size="6")

    set_cell_borders(header_row.cells[0], size="6")
    set_cell_borders(formula_row.cells[0], size="6")
    set_cell_borders(duplicate_formula_row.cells[0], size="6")

    formulas = ["N=(D+S)", "(M)", "Q = H-C", "X= (N+M)"]

    for target_row in [formula_row, duplicate_formula_row]:
        for i, formula in enumerate(formulas, start=1):
            cell = target_row.cells[i]
            set_cell_text(cell, formula)
            set_cell_width(cell, COMPACT_TABLE_WIDTHS[i])
            style_table_cell(cell, font_size=10, bold=True)
            set_cell_borders(cell, size="6")

    for row_index, (_, row) in enumerate(chart_df.iterrows(), start=3):
        row_obj = small_table.rows[row_index]
        set_row_height_points(row_obj, 16.1)

        values = [row["TIME"], row["N"], row["M"], row["Q"], row["X"]]

        for i, value in enumerate(values):
            cell = row_obj.cells[i]
            set_cell_text(cell, value if i == 0 else format_count(value))
            set_cell_width(cell, COMPACT_TABLE_WIDTHS[i])

            style_table_cell(
                cell,
                font_size=10 if i in {0, 1, 3, 4} else 11,
                align=WD_ALIGN_PARAGRAPH.LEFT,
            )
            set_cell_borders(cell, size="6")

    totals = small_table.rows[-1]
    set_row_height_points(totals, 16.1)

    total_values = [
        "Total",
        chart_df["N"].sum(),
        chart_df["M"].sum(),
        chart_df["Q"].sum(),
        chart_df["X"].sum(),
    ]

    for i, value in enumerate(total_values):
        cell = totals.cells[i]
        set_cell_text(cell, value if i == 0 else format_count(value))
        set_cell_width(cell, COMPACT_TABLE_WIDTHS[i])
        style_table_cell(
            cell,
            font_size=10 if i == 0 else 11,
            bold=True,
            align=WD_ALIGN_PARAGRAPH.CENTER if i == 0 else WD_ALIGN_PARAGRAPH.LEFT,
        )
        set_cell_borders(cell, size="6")

    chart_image = create_daily_hour_chart_image(chart_df)

    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run()
    run.add_picture(chart_image, width=Inches(6.05), height=Inches(5.84))
