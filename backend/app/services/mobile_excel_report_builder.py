import io
from copy import copy
from datetime import datetime
from typing import Any

import pandas as pd
from openpyxl import Workbook
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.text import RichText
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor
from openpyxl.drawing.text import CharacterProperties, Paragraph, ParagraphProperties, RichTextProperties
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.worksheet import Worksheet

from app.services.daily_hour_processor import HOURS
from app.services.mobile_report_processor import summarize_mobile_report


FONT_NAME = "Arial"
DETAIL_FONT_NAME = "Calibri"
DATE_FORMAT = "dd-mmm-yy"

THIN_SIDE = Side(style="thin", color="000000")
MEDIUM_SIDE = Side(style="medium", color="000000")
THIN_BORDER = Border(
    left=THIN_SIDE,
    right=THIN_SIDE,
    top=THIN_SIDE,
    bottom=THIN_SIDE,
)
MEDIUM_BORDER = Border(
    left=MEDIUM_SIDE,
    right=MEDIUM_SIDE,
    top=MEDIUM_SIDE,
    bottom=MEDIUM_SIDE,
)

NO_FILL = PatternFill(fill_type=None)
RED_FILL = PatternFill("solid", fgColor="FFFF0000")
DARK_BLUE_LINE = "1F4E79"
MAROON_LINE = "800000"


SUMMARY_COLUMN_WIDTHS = {
    "B": 5.43,
    "C": 12.29,
    "D": 12.0,
    "E": 10.86,
    "F": 10.14,
    "G": 11.29,
    "H": 12.0,
    "I": 11.86,
    "J": 11.43,
    "K": 12.86,
    "L": 13.29,
    "M": 12.43,
    "N": 13.14,
    "O": 13.14,
    "P": 13.14,
    "Q": 13.14,
    "R": 13.14,
    "S": 13.14,
    "T": 13.14,
    "U": 13.14,
    "V": 13.14,
}

DETAIL_COLUMN_WIDTHS = {
    "A": 2.0,
    "B": 11.43,
    "C": 16.43,
    "D": 43.29,
    "E": 12.14,
    "F": 27.29,
    "G": 19.43,
    "H": 21.71,
    "I": 13.57,
    "J": 31.29,
    "K": 65.71,
    "L": 42.71,
    "M": 23.14,
    "N": 132.86,
}


def _font(
    size: float = 11,
    bold: bool = False,
    *,
    name: str = FONT_NAME,
) -> Font:
    return Font(name=name, size=size, bold=bold)


def _alignment(
    horizontal: str | None = "center",
    vertical: str | None = "center",
    wrap_text: bool | None = True,
) -> Alignment:
    return Alignment(horizontal=horizontal, vertical=vertical, wrap_text=wrap_text)


def _set_cell(
    ws: Worksheet,
    coordinate: str,
    value: Any = None,
    *,
    size: float = 11,
    bold: bool = False,
    font_name: str = FONT_NAME,
    fill: PatternFill | None = NO_FILL,
    border: Border | None = THIN_BORDER,
    horizontal: str | None = "center",
    vertical: str | None = "center",
    wrap_text: bool | None = True,
    number_format: str = "General",
) -> None:
    cell = ws[coordinate]
    cell.value = value
    cell.font = _font(size=size, bold=bold, name=font_name)
    cell.alignment = _alignment(horizontal, vertical, wrap_text)
    if fill is not None:
        cell.fill = copy(fill)
    if border is not None:
        cell.border = copy(border)
    cell.number_format = number_format


def _style_range(
    ws: Worksheet,
    cell_range: str,
    *,
    size: float = 11,
    bold: bool = False,
    font_name: str = FONT_NAME,
    fill: PatternFill | None = NO_FILL,
    border: Border | None = THIN_BORDER,
    horizontal: str | None = "center",
    vertical: str | None = "center",
    wrap_text: bool | None = True,
    number_format: str = "General",
) -> None:
    for row in ws[cell_range]:
        for cell in row:
            cell.font = _font(size=size, bold=bold, name=font_name)
            cell.alignment = _alignment(horizontal, vertical, wrap_text)
            if fill is not None:
                cell.fill = copy(fill)
            if border is not None:
                cell.border = copy(border)
            cell.number_format = number_format


def _merge_and_set(
    ws: Worksheet,
    cell_range: str,
    value: Any = None,
    **style_kwargs,
) -> None:
    _style_range(ws, cell_range, **style_kwargs)
    ws.merge_cells(cell_range)
    _set_cell(ws, cell_range.split(":", 1)[0], value, **style_kwargs)


def _set_detail_text_cell(
    ws: Worksheet,
    coordinate: str,
    value: Any = None,
    *,
    border: Border | None = THIN_BORDER,
    vertical: str | None = "center",
    number_format: str = "General",
) -> None:
    _set_cell(
        ws,
        coordinate,
        value,
        font_name=DETAIL_FONT_NAME,
        border=border,
        horizontal="left",
        vertical=vertical,
        wrap_text=False,
        number_format=number_format,
    )


def _set_detail_center_cell(
    ws: Worksheet,
    coordinate: str,
    value: Any = None,
    *,
    border: Border | None = THIN_BORDER,
    number_format: str = "General",
    size: float = 11,
) -> None:
    _set_cell(
        ws,
        coordinate,
        value,
        size=size,
        font_name=DETAIL_FONT_NAME,
        border=border,
        horizontal="center",
        vertical="center",
        wrap_text=False,
        number_format=number_format,
    )


def _chart_axis_text_properties() -> RichText:
    return RichText(
        bodyPr=RichTextProperties(
            rot=-60000000,
            spcFirstLastPara=True,
            vertOverflow="ellipsis",
            vert="horz",
            wrap="square",
            anchor="ctr",
            anchorCtr=True,
        ),
        p=[
            Paragraph(
                pPr=ParagraphProperties(
                    defRPr=CharacterProperties(
                        sz=900,
                        b=False,
                        i=False,
                        u="none",
                        strike="noStrike",
                        kern=1200,
                        baseline=0,
                    )
                )
            )
        ],
    )


def _note_key_value(key: str, value: str) -> CellRichText:
    return CellRichText(
        TextBlock(InlineFont(b=True), f"{key} ="),
        f" {value}",
    )


def _merge_and_set_note_row(
    ws: Worksheet,
    cell_range: str,
    key: str,
    value: str,
    *,
    fill: PatternFill | None = NO_FILL,
) -> None:
    _style_range(
        ws,
        cell_range,
        fill=fill,
        border=MEDIUM_BORDER,
        horizontal="left",
        wrap_text=False,
    )
    ws.merge_cells(cell_range)
    cell = ws[cell_range.split(":", 1)[0]]
    cell.value = _note_key_value(key, value)


def _upper(value: Any) -> str:
    if value is None:
        return ""

    if isinstance(value, (list, tuple)):
        return " / ".join(_upper(item) for item in value if _upper(item))

    text = str(value).strip()
    if not text:
        return ""
    return text.upper()


def _manual_source(session) -> dict[str, Any]:
    mobile = session.manual_inputs.get("mobile_report")
    if isinstance(mobile, dict):
        return mobile
    return {}


def _manual_value(session, *keys: str, default: Any = "") -> Any:
    mobile = _manual_source(session)
    for source in (mobile, session.manual_inputs):
        for key in keys:
            value = source.get(key)
            if value not in (None, ""):
                return value
    return default


def _manual_int(session, *keys: str, default: int = 0) -> int:
    value = _manual_value(session, *keys, default=default)
    number = pd.to_numeric(value, errors="coerce")
    if pd.isna(number):
        return default
    return int(number)


def _transgressions_count(session) -> int:
    value = _manual_value(session, "transgressions_count", "transgressions", default=0)
    if isinstance(value, dict):
        return len(value.get("daily_transgressions") or [])
    if isinstance(value, list):
        return len(value)
    return _manual_int(session, "transgressions_count", default=0)


def _report_date(session, records: pd.DataFrame) -> datetime | str:
    if not records.empty and "date_time" in records:
        dates = records["date_time"].dt.date.dropna().unique().tolist()
        if len(dates) == 1:
            return datetime.combine(dates[0], datetime.min.time())

    try:
        return datetime.strptime(session.report_date, "%Y-%m-%d")
    except ValueError:
        return session.report_date


def _setup_summary_sheet(ws: Worksheet) -> None:
    ws.title = "Weigh & Hourly Summary"
    for column, width in SUMMARY_COLUMN_WIDTHS.items():
        ws.column_dimensions[column].width = width

    ws.row_dimensions[2].height = 15.75
    ws.row_dimensions[3].height = 21.75
    ws.row_dimensions[4].height = 24.75
    for row in range(5, 30):
        ws.row_dimensions[row].height = 15.75 if row in {5, 6, 7, 29} else 15.0

    ws.row_dimensions[30].height = 15.75
    ws.row_dimensions[31].height = 15.75
    ws.row_dimensions[32].height = 55.5
    ws.row_dimensions[33].height = 18.75
    ws.row_dimensions[34].height = 18.75
    ws.row_dimensions[35].height = 19.5
    ws.row_dimensions[36].height = 15.0
    ws.row_dimensions[37].height = 23.25

    _merge_and_set(ws, "C3:C4", "Date", bold=True, border=MEDIUM_BORDER)
    _merge_and_set(ws, "D3:D4", "Time", bold=True, border=MEDIUM_BORDER)
    _merge_and_set(ws, "E3:E4", "Trucks Weighed", bold=True, border=MEDIUM_BORDER)
    _merge_and_set(ws, "F3:F4", "Warned Trucks", bold=True, border=MEDIUM_BORDER)
    _merge_and_set(ws, "G3:G4", "Charged Trucks", bold=True, border=MEDIUM_BORDER)
    _set_cell(ws, "H3", "Excess GVW ", bold=True, border=MEDIUM_BORDER, vertical="top")
    _set_cell(ws, "H4", "Weight(KGS)", bold=True, border=MEDIUM_BORDER, vertical="top")
    _merge_and_set(ws, "J3:L3", "hourly summary", bold=True, border=MEDIUM_BORDER)
    _set_cell(ws, "J4", "Date", bold=True, border=MEDIUM_BORDER)
    _set_cell(ws, "K4", "WEIGHED", bold=True, border=MEDIUM_BORDER)
    _set_cell(ws, "L4", "CHARGED", bold=True, border=MEDIUM_BORDER)
    _merge_and_set(ws, "N3:V30", None, border=MEDIUM_BORDER, wrap_text=False)

    summary_headers = {
        "D32": "Total Weighed (X)",
        "E32": "Warned & Charged (Y)",
        "F32": "Charged & Prohibited (Z)",
        "G32": "Warned (A)",
        "H32": "Cases Cleared in Court (B)",
        "I32": "  Transgressions (L)",
        "J32": "Exempted permit (E)",
        "K32": "Manually Weighed (M)",
    }
    for coordinate, label in summary_headers.items():
        _set_cell(ws, coordinate, label, bold=True, border=MEDIUM_BORDER)

    _merge_and_set(ws, "O32:T32", "Notes", size=28, bold=True, border=MEDIUM_BORDER, wrap_text=False)
    _merge_and_set_note_row(
        ws,
        "O33:T33",
        "Red",
        "formulae, do not edit",
        fill=RED_FILL,
    )
    _merge_and_set_note_row(
        ws,
        "O34:T34",
        "No fill",
        "manual entries fill in from data collection forms",
    )
    _merge_and_set_note_row(
        ws,
        "O35:T35",
        "Data",
        "in the table is for illustration purposes ONLY",
    )


def _setup_detail_sheet(ws: Worksheet, title: str) -> None:
    ws.title = "Mobile Daily Report"
    for column, width in DETAIL_COLUMN_WIDTHS.items():
        ws.column_dimensions[column].width = width

    ws.row_dimensions[1].height = 15.75
    ws.row_dimensions[2].height = 25.5
    ws.row_dimensions[3].height = 14.25
    ws.row_dimensions[4].height = 18.75
    ws.row_dimensions[5].height = 31.5
    ws.row_dimensions[6].height = 18.0
    ws.row_dimensions[7].height = 15.75
    ws.row_dimensions[8].height = 15.75
    ws.row_dimensions[10].height = 30.0

    _merge_and_set(ws, "B2:N2", title, size=14, bold=True, border=None)
    _merge_and_set(ws, "K4:M4", "MILEAGE", bold=True, border=MEDIUM_BORDER)

    for coordinate, value in {
        "B5": "DATE",
        "E5": "DANKA STAFF",
        "G5": "POLICE OFFICERS",
        "I5": "TRUCKS WEIGHED",
        "J5": "TRUCKS CHARGED",
        "K5": "MILEAGE START",
        "L5": "MILEAGE END",
        "M5": "KMS",
        "N5": "MOBILE VEHICLE",
    }.items():
        _set_cell(ws, coordinate, value, bold=True, border=MEDIUM_BORDER)

    _merge_and_set(ws, "C5:D5", "", border=MEDIUM_BORDER)
    _merge_and_set(ws, "E5:F5", "DANKA STAFF", bold=True, border=MEDIUM_BORDER)
    _merge_and_set(ws, "G5:H5", "POLICE OFFICERS", bold=True, border=MEDIUM_BORDER)
    _merge_and_set(ws, "C6:D6", "", border=THIN_BORDER)
    _merge_and_set(ws, "C7:D7", "", border=THIN_BORDER)

    _set_cell(ws, "G8", "TOTAL", bold=True, border=None)
    _set_cell(ws, "J8", "=J7+J6", bold=True, fill=RED_FILL)
    _set_cell(ws, "M8", "=M6", bold=True, fill=RED_FILL)

    _merge_and_set(ws, "B9:E9", "", border=MEDIUM_BORDER)
    _merge_and_set(ws, "F9:G9", "EXCESS   WEIGHT", bold=True, border=MEDIUM_BORDER)
    _merge_and_set(ws, "H9:N9", "", border=MEDIUM_BORDER)

    headers = {
        "B10": "DATE  WEIGHED",
        "C10": "VEHICLE REG",
        "D10": "TRANSPORTER",
        "E10": "CONFIG.",
        "F10": "GVW",
        "G10": "AXLE",
        "H10": "ORIGIN",
        "I10": "DESTIN.",
        "J10": "CARGO",
        "K10": "COMPUTER OPERATER \n(DANKA STAFF)",
        "L10": "OFFICERS",
        "M10": "REMARKS",
        "N10": "ROUTE",
    }
    for coordinate, value in headers.items():
        _set_cell(
            ws,
            coordinate,
            value,
            bold=True,
            border=MEDIUM_BORDER,
            horizontal="center",
        )


def _write_summary_rows(
    ws: Worksheet,
    records: pd.DataFrame,
    report_date: datetime | str,
    summary: dict[str, Any],
) -> None:
    for offset, hour in enumerate(HOURS):
        row = 5 + offset
        hour_records = records.loc[records["hour_band"].eq(hour)]
        warned = hour_records["remarks"].str.strip().str.upper().eq("WARNED")
        charged = hour_records["is_gvw_axle_charge"] | hour_records["is_dimension_charge"]
        excess_gvw = hour_records["gvw_difference_kg"].clip(lower=0).sum()

        _set_cell(
            ws,
            f"C{row}",
            report_date if offset == 0 else None,
            number_format=DATE_FORMAT,
        )
        _set_cell(ws, f"D{row}", hour)
        _set_cell(ws, f"E{row}", int(hour_records["is_weighed"].sum()))
        _set_cell(ws, f"F{row}", int(warned.sum()))
        _set_cell(ws, f"G{row}", int(charged.sum()))
        _set_cell(ws, f"H{row}", int(excess_gvw), number_format="#,##0")
        _set_cell(ws, f"J{row}", hour)
        _set_cell(ws, f"K{row}", f"=E{row}", fill=RED_FILL)
        _set_cell(ws, f"L{row}", f"=G{row}", fill=RED_FILL)

    for coordinate, formula in {
        "D29": "Total",
        "E29": "=SUM(E5:E28)",
        "F29": "=SUM(F5:F28)",
        "G29": "=SUM(G5:G28)",
        "H29": "=SUM(H5:H28)",
        "J29": "Total",
        "K29": "=SUM(K5:K28)",
        "L29": "=SUM(L5:L28)",
    }.items():
        _set_cell(ws, coordinate, formula, bold=True, fill=RED_FILL, border=MEDIUM_BORDER)

    _set_cell(ws, "D33", "=E29", fill=RED_FILL, border=MEDIUM_BORDER)
    _set_cell(ws, "E33", "=F29+G29", fill=RED_FILL, border=MEDIUM_BORDER)
    _set_cell(ws, "F33", "=G29", fill=RED_FILL, border=MEDIUM_BORDER)
    _set_cell(ws, "G33", "=F29", fill=RED_FILL, border=MEDIUM_BORDER)
    _set_cell(ws, "H33", int(summary.get("cases_cleared_in_court", 0)), border=MEDIUM_BORDER)
    _set_cell(ws, "I33", int(summary.get("transgressions_count", 0)), border=MEDIUM_BORDER)
    _set_cell(ws, "J33", int(summary.get("exempted_permit", 0)), border=MEDIUM_BORDER)
    _set_cell(ws, "K33", int(summary.get("manually_weighed", 0)), border=MEDIUM_BORDER)
    _merge_and_set(ws, "D34:K34", "weighed summary", bold=True, border=MEDIUM_BORDER)


def _add_hourly_summary_chart(ws: Worksheet) -> None:
    chart = LineChart()
    chart.title = "DAILY HOURLY DATA"
    chart.style = 2
    chart.roundedCorners = False
    chart.varyColors = False
    chart.height = 7.5
    chart.width = 11.25
    chart.legend.position = "b"
    chart.dataLabels = DataLabelList()
    chart.dataLabels.showLegendKey = False
    chart.dataLabels.showVal = False
    chart.dataLabels.showCatName = False
    chart.dataLabels.showSerName = False
    chart.dataLabels.showPercent = False
    chart.dataLabels.showBubbleSize = False

    data = Reference(ws, min_col=11, max_col=12, min_row=4, max_row=28)
    categories = Reference(ws, min_col=10, min_row=5, max_row=28)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)

    line_colors = (DARK_BLUE_LINE, MAROON_LINE)
    for series, color in zip(chart.ser, line_colors):
        series.smooth = False
        series.marker.symbol = "none"
        series.graphicalProperties.line.width = 28575
        series.graphicalProperties.line.cap = "rnd"
        series.graphicalProperties.line.round = True
        series.graphicalProperties.line.solidFill = color

    chart.x_axis.axId = 1549586240
    chart.y_axis.axId = 1549567520
    chart.x_axis.crossAx = 1549567520
    chart.y_axis.crossAx = 1549586240
    chart.x_axis.axPos = "b"
    chart.y_axis.axPos = "l"
    chart.x_axis.crosses = "autoZero"
    chart.y_axis.crosses = "autoZero"
    chart.y_axis.crossBetween = "between"
    chart.x_axis.tickLblPos = "nextTo"
    chart.y_axis.tickLblPos = "nextTo"
    chart.x_axis.number_format = "General"
    chart.y_axis.number_format = "General"
    chart.x_axis.textProperties = _chart_axis_text_properties()
    chart.y_axis.textProperties = _chart_axis_text_properties()
    chart.x_axis.majorTickMark = "none"
    chart.x_axis.minorTickMark = "none"
    chart.y_axis.majorTickMark = "none"
    chart.y_axis.minorTickMark = "none"

    chart.anchor = TwoCellAnchor(
        _from=AnchorMarker(col=13, colOff=728385, row=3, rowOff=51548),
        to=AnchorMarker(col=21, colOff=67238, row=27, rowOff=156883),
    )
    ws.add_chart(chart)


def _write_manual_header(
    ws: Worksheet,
    session,
    report_date: datetime | str,
    summary: dict[str, Any],
) -> None:
    danka_staff = _manual_value(
        session,
        "danka_staff",
        "danka_officers",
        "computer_operator",
        "computer_operators",
    )
    police_officers = _manual_value(session, "police_officers", "police")
    mileage_start = _manual_value(session, "mileage_start")
    mileage_end = _manual_value(session, "mileage_end")

    _set_detail_center_cell(ws, "B6", report_date, number_format=DATE_FORMAT)
    _set_detail_text_cell(ws, "E6", _upper(danka_staff))
    _set_detail_text_cell(ws, "G6", _upper(police_officers), vertical=None)
    _set_detail_center_cell(ws, "I6", int(summary["total_trucks_weighed"]), border=MEDIUM_BORDER, size=10)
    _set_detail_center_cell(ws, "J6", int(summary["charged_gvw_axle_trucks"]), border=MEDIUM_BORDER, size=10)
    _set_detail_center_cell(ws, "J7", int(summary["charged_dimensions_trucks"]))
    _set_detail_center_cell(ws, "K6", mileage_start if mileage_start != "" else None, border=MEDIUM_BORDER)
    _set_detail_center_cell(ws, "L6", mileage_end if mileage_end != "" else None)
    _set_detail_center_cell(
        ws,
        "M6",
        "=L6-K6" if mileage_start != "" and mileage_end != "" else None,
        border=MEDIUM_BORDER,
        size=10,
    )
    _set_detail_center_cell(
        ws,
        "N6",
        _upper(_manual_value(session, "mobile_vehicle", "vehicle_used", "vehicle")),
        border=MEDIUM_BORDER,
        size=10,
    )


def _write_detail_rows(
    ws: Worksheet,
    records: pd.DataFrame,
    session,
    report_date: datetime | str,
) -> None:
    route = _upper(_manual_value(session, "route"))
    danka_staff = _upper(
        _manual_value(
            session,
            "danka_staff",
            "danka_officers",
            "computer_operator",
            "computer_operators",
        )
    )
    police_officers = _upper(_manual_value(session, "police_officers", "police"))

    for offset, (_, record) in enumerate(records.iterrows()):
        row = 11 + offset
        gvw_excess = record["gvw_difference_kg"] if record["gvw_difference_kg"] > 0 else "-"
        axle_excess = record["excess_kg"] if record["excess_kg"] > 0 else "-"

        values = {
            f"B{row}": report_date,
            f"C{row}": _upper(record["registration"]),
            f"D{row}": _upper(record["transporter"]),
            f"E{row}": _upper(record["axle"]),
            f"F{row}": int(gvw_excess) if gvw_excess != "-" else "-",
            f"G{row}": int(axle_excess) if axle_excess != "-" else "-",
            f"H{row}": _upper(record["origin"]),
            f"I{row}": _upper(record["destination"]),
            f"J{row}": _upper(record["cargo"]),
            f"K{row}": danka_staff,
            f"L{row}": police_officers,
            f"M{row}": _upper(record["remarks"]),
            f"N{row}": route,
        }

        for coordinate, value in values.items():
            number_format = DATE_FORMAT if coordinate.startswith("B") else "General"
            column = coordinate[0]
            if column in {"B", "F", "G"}:
                _set_detail_center_cell(
                    ws,
                    coordinate,
                    value,
                    number_format=number_format,
                )
            else:
                _set_detail_text_cell(
                    ws,
                    coordinate,
                    value,
                    vertical="center" if column in {"K", "M"} else None,
                    number_format=number_format,
                )


def build_mobile_excel_report(session) -> io.BytesIO:
    if (
        "mobile_report" not in session.dataframes
        or session.sections.get("mobile_report", {}).get("status") != "ready"
    ):
        raise ValueError("Mobile report data is not ready")

    records = session.dataframes["mobile_report"].copy()
    if records.empty:
        raise ValueError("Mobile report data is empty")

    records["date_time"] = pd.to_datetime(records["date_time"], errors="coerce")
    records = records.loc[records["date_time"].notna()].copy()
    if records.empty:
        raise ValueError("Mobile report data has no valid Date Time values")

    summary = summarize_mobile_report(records)
    summary["cases_cleared_in_court"] = _manual_int(
        session,
        "cases_cleared_in_court",
        "cases_cleared_court",
        default=0,
    )
    summary["transgressions_count"] = _transgressions_count(session)
    summary["exempted_permit"] = _manual_int(session, "exempted_permit", default=0)
    summary["manually_weighed"] = _manual_int(session, "manually_weighed", default=0)

    report_date = _report_date(session, records)
    title = f"{session.station} DAILY {session.bound} REPORT".strip().upper()

    wb = Workbook()
    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
    except AttributeError:
        pass

    summary_ws = wb.active
    _setup_summary_sheet(summary_ws)
    _write_summary_rows(summary_ws, records, report_date, summary)
    _add_hourly_summary_chart(summary_ws)

    detail_ws = wb.create_sheet("Mobile Daily Report")
    _setup_detail_sheet(detail_ws, title)
    _write_manual_header(detail_ws, session, report_date, summary)
    _write_detail_rows(detail_ws, records, session, report_date)

    hidden_ws = wb.create_sheet("Sheet3")
    hidden_ws.sheet_state = "hidden"

    stream = io.BytesIO()
    wb.save(stream)
    stream.seek(0)
    return stream
