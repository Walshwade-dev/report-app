import io
from datetime import datetime
from typing import Any

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from app.services.daily_hour_chart_generator import build_daily_hour_chart_data
from app.services.daily_summary_processor import (
    DailySummaryMissingSourceError,
    build_daily_summary_from_session,
    daily_summary_rows,
)
from app.services.report_session_metrics import get_wideload_count_from_session
from app.services.traffic_census_processor import traffic_census_rows


TITLE_FILL = PatternFill("solid", fgColor="D9EAF7")
SECTION_FILL = PatternFill("solid", fgColor="FFF2CC")
HEADER_FILL = PatternFill("solid", fgColor="E2F0D9")
THIN_BORDER = Border(
    left=Side(style="thin", color="000000"),
    right=Side(style="thin", color="000000"),
    top=Side(style="thin", color="000000"),
    bottom=Side(style="thin", color="000000"),
)


DAILY_HOUR_HEADER_ROWS = [
    [
        "DATE",
        "TIME",
        "TRUCKS WEIGHED",
        "",
        "",
        "",
        "",
        "",
        "CALLED IN",
        "TOTAL OVERLOADED",
        "IMPOUNDED & PROHIBITED",
        "WARNED TRUCKS",
        "CHARGED & PROHIBITED",
        "SPECIAL RELEASE",
        "REDISTRIBUTED",
        "EXEMPTION PERMITS NOT WEIGHED",
    ],
    [
        "",
        "",
        "MULTIDECK SCALE",
        "WEIGHED SAW",
        "MANUAL",
        "HSWIM TOTAL",
        "HSWIM - CLEARED",
        "TOTAL WEIGHED",
        "(C)",
        "(Y)=A+Z+G+R",
        "(P)=Z+R+G",
        "(A)",
        "(Z)",
        "(G)",
        "(R)",
        "(E)",
    ],
    [
        "",
        "",
        "(D)",
        "(S)",
        "(M)",
        "(H)",
        "Q=H-C",
        "X=(D+S+M)",
        "",
        "",
        "",
        "",
        "",
        "",
        "",
        "",
    ],
]


def _display_report_date(report_date: str) -> str:
    try:
        return datetime.strptime(report_date, "%Y-%m-%d").strftime("%d/%m/%Y")
    except ValueError:
        return report_date


def _title(session) -> str:
    parts = [
        session.station,
        "WEIGHBRIDGE",
        session.bound,
        "DAILY REPORT",
        _display_report_date(session.report_date),
    ]
    return " ".join(str(part).strip() for part in parts if part).upper()


def _set_cell(
    ws: Worksheet,
    row: int,
    column: int,
    value: Any,
    *,
    bold: bool = False,
    fill: PatternFill | None = None,
    align: str = "center",
    border: bool = False,
    font_size: int = 10,
) -> None:
    cell = ws.cell(row=row, column=column, value=value)
    cell.font = Font(name="Arial", size=font_size, bold=bold)
    cell.alignment = Alignment(
        horizontal=align,
        vertical="center",
        wrap_text=True,
    )
    if fill is not None:
        cell.fill = fill
    if border:
        cell.border = THIN_BORDER
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        cell.number_format = "#,##0"


def _section_title(ws: Worksheet, row: int, title: str, last_column: int) -> int:
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=last_column)
    _set_cell(
        ws,
        row,
        1,
        title,
        bold=True,
        fill=SECTION_FILL,
        align="left",
        border=True,
        font_size=11,
    )
    return row + 1


def _write_table(
    ws: Worksheet,
    start_row: int,
    headers: list[str],
    rows: list[list[Any]],
    *,
    start_column: int = 1,
) -> int:
    for column_offset, header in enumerate(headers):
        _set_cell(
            ws,
            start_row,
            start_column + column_offset,
            header,
            bold=True,
            fill=HEADER_FILL,
            border=True,
        )

    for row_offset, row_values in enumerate(rows, start=1):
        for column_offset, value in enumerate(row_values):
            _set_cell(
                ws,
                start_row + row_offset,
                start_column + column_offset,
                _clean_value(value),
                border=True,
                align="left" if column_offset == 0 else "center",
            )

    return start_row + len(rows) + 1


def _clean_value(value: Any) -> Any:
    if pd.isna(value):
        return None
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


def _dataframe_rows(df: pd.DataFrame) -> list[list[Any]]:
    return [
        [_clean_value(value) for value in row]
        for row in df.itertuples(index=False, name=None)
    ]


def _numeric_int(value: Any) -> int:
    number = pd.to_numeric(value, errors="coerce")
    if pd.isna(number):
        return 0
    return int(number)


def _daily_hour_totals(daily_df: pd.DataFrame) -> dict[str, int]:
    if "DATE" in daily_df.columns:
        totals_mask = daily_df["DATE"].astype(str).str.strip().str.lower().eq("totals")
        if totals_mask.any():
            totals_row = daily_df.loc[totals_mask].iloc[-1]
        else:
            totals_row = daily_df.drop(columns=["DATE", "TIME"], errors="ignore").sum(
                numeric_only=True
            )
    else:
        totals_row = daily_df.sum(numeric_only=True)

    return {
        column.lower(): _numeric_int(totals_row.get(column, 0))
        for column in ["Q", "X", "E"]
    }


def _write_daily_hour_statistics(ws: Worksheet, row: int, daily_df: pd.DataFrame) -> int:
    columns = list(daily_df.columns)
    row = _section_title(ws, row, "DAILY AND HOURLY STATISTICS", len(columns))

    for header_offset, header_row in enumerate(DAILY_HOUR_HEADER_ROWS):
        for column_index, value in enumerate(header_row[: len(columns)], start=1):
            _set_cell(
                ws,
                row + header_offset,
                column_index,
                value,
                bold=True,
                fill=HEADER_FILL,
                border=True,
                font_size=9,
            )

    if len(columns) >= 8:
        ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=8)

    for data_offset, values in enumerate(_dataframe_rows(daily_df), start=3):
        is_total = str(values[0]).strip().lower() == "totals" if values else False
        for column_index, value in enumerate(values, start=1):
            _set_cell(
                ws,
                row + data_offset,
                column_index,
                value,
                bold=is_total,
                border=True,
                align="center",
            )

    return row + len(daily_df) + 5


def _write_daily_hour_data(ws: Worksheet, row: int, daily_df: pd.DataFrame) -> int:
    chart_df = build_daily_hour_chart_data(daily_df)
    row = _section_title(ws, row, "DAILY HOURLY DATA", 5)

    headers = [
        "Time",
        "Multideck weighed\nN=(D+S)",
        "Manually\n(M)",
        "HSWIM CLEARED\nQ=H-C",
        "Total weighed\nX=(N+M)",
    ]
    rows = [
        [
            record["TIME"],
            int(record["N"]),
            int(record["M"]),
            int(record["Q"]),
            int(record["X"]),
        ]
        for record in chart_df.to_dict(orient="records")
    ]
    rows.append(
        [
            "Total",
            int(chart_df["N"].sum()),
            int(chart_df["M"].sum()),
            int(chart_df["Q"].sum()),
            int(chart_df["X"].sum()),
        ]
    )

    return _write_table(ws, row, headers, rows) + 2


def _write_traffic_census(ws: Worksheet, row: int, session, daily_df: pd.DataFrame) -> int:
    row = _section_title(ws, row, "TRAFFIC CENSUS DATA", 7)
    traffic_census = session.manual_inputs.get("traffic_census")

    if not traffic_census:
        return _write_table(
            ws,
            row,
            ["Status"],
            [["NIL - traffic census manual input has not been captured."]],
        ) + 2

    traffic_values = dict(traffic_census)
    totals = _daily_hour_totals(daily_df)
    wideload_count = get_wideload_count_from_session(session)
    exemption_not_weighed = wideload_count if wideload_count is not None else totals["e"]
    total_traffic = (
        totals["q"]
        + totals["x"]
        + int(traffic_values["total_traffic_census"])
        + exemption_not_weighed
    )
    headers = [
        "Buses >= 3,500 kg",
        "Vehicles >= 3,500 kg but < 7,000 kg excluding buses",
        "Vehicles >= 7,000 kg excluding buses",
        "Total Traffic Census (K)",
        "Exemption permits Not weighed (E)",
        "Total Weighed (X)",
        "Total Traffic (T)=Q+X+K+E",
    ]
    rows = [
        [
            traffic_values["buses_gte_3500kg"],
            traffic_values["vehicles_3500_to_7000_excluding_buses"],
            traffic_values["vehicles_gte_7000_excluding_buses"],
            traffic_values["total_traffic_census"],
            exemption_not_weighed,
            totals["x"],
            total_traffic,
        ]
    ]

    return _write_table(ws, row, headers, rows) + 2


def _daily_summary_values(session) -> dict[str, int] | None:
    values = session.sections.get("daily_summary", {}).get("values")
    if values:
        return values

    try:
        return build_daily_summary_from_session(session)
    except DailySummaryMissingSourceError:
        return None


def _write_daily_summary(ws: Worksheet, row: int, session) -> int:
    row = _section_title(ws, row, "DAILY SUMMARY", 16)
    summary = _daily_summary_values(session)

    if not summary:
        return _write_table(
            ws,
            row,
            ["Status"],
            [["NIL - daily summary source data is not ready."]],
        )

    headers = [label for label, _ in daily_summary_rows(summary)]
    rows = [[value for _, value in daily_summary_rows(summary)]]
    return _write_table(ws, row, headers, rows)


def _write_cc_records(ws: Worksheet, session) -> None:
    ws.freeze_panes = "A6"
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=7)
    _set_cell(ws, 1, 1, _title(session), bold=True, fill=TITLE_FILL, font_size=12)
    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=4)
    _set_cell(ws, 3, 1, "TOTAL TRAFFIC CENSUS SUMMARY", bold=True, fill=SECTION_FILL)

    traffic_census = session.manual_inputs.get("traffic_census")
    if traffic_census:
        rows = [[label, value] for label, value in traffic_census_rows(traffic_census)]
    else:
        rows = [["Status", "NIL - traffic census manual input has not been captured."]]
    next_row = _write_table(ws, 5, ["Category", "Count"], rows) + 2

    cc_records = session.manual_inputs.get("cc_records")
    headers = [
        "Record",
        "Date",
        "Time",
        "Vehicle Registration",
        "Category",
        "Weight Band",
        "Count",
        "Remarks",
    ]

    if isinstance(cc_records, list) and cc_records:
        rows = [
            [
                index,
                record.get("date"),
                record.get("time"),
                record.get("vehicle_registration") or record.get("registration"),
                record.get("category"),
                record.get("weight_band"),
                record.get("count", 1),
                record.get("remarks"),
            ]
            for index, record in enumerate(cc_records, start=1)
        ]
    else:
        rows = [["NIL", "", "", "", "", "", 0, "No detailed CC records captured."]]

    _section_title(ws, next_row, "CENSUS / CC RECORDS", len(headers))
    _write_table(ws, next_row + 1, headers, rows)


def _set_reasonable_widths(ws: Worksheet) -> None:
    for column_cells in ws.columns:
        letter = get_column_letter(column_cells[0].column)
        max_length = 10
        for cell in column_cells:
            value = cell.value
            if value is None:
                continue
            max_length = max(max_length, min(len(str(value)), 38))
        ws.column_dimensions[letter].width = max_length + 2


def _apply_page_setup(ws: Worksheet) -> None:
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def build_excel_report(session) -> io.BytesIO:
    if (
        "daily_hour" not in session.dataframes
        or session.sections.get("daily_hour", {}).get("status") != "ready"
    ):
        raise ValueError("Excel report requires ready daily_hour data.")

    daily_df = session.dataframes["daily_hour"]

    workbook = Workbook()
    summary = workbook.active
    summary.title = "Summary"
    summary.freeze_panes = "A4"
    summary.merge_cells(start_row=1, start_column=1, end_row=1, end_column=16)
    _set_cell(summary, 1, 1, _title(session), bold=True, fill=TITLE_FILL, font_size=12)

    row = 3
    row = _write_daily_hour_statistics(summary, row, daily_df)
    row = _write_daily_hour_data(summary, row, daily_df)
    row = _write_traffic_census(summary, row, session, daily_df)
    _write_daily_summary(summary, row, session)

    cc_records = workbook.create_sheet("CC records")
    _write_cc_records(cc_records, session)

    for worksheet in workbook.worksheets:
        _set_reasonable_widths(worksheet)
        _apply_page_setup(worksheet)

    buffer = io.BytesIO()
    workbook.save(buffer)
    buffer.seek(0)
    return buffer
