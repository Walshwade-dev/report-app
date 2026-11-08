import calendar
import io
import re
from datetime import datetime
from typing import Any

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from app.services.report_session_store import report_session_store


FONT_ARIAL = "Arial"
FONT_CALIBRI = "Calibri"

THIN_SIDE = Side(style="thin", color="000000")
MEDIUM_SIDE = Side(style="medium", color="000000")

THIN_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)
MEDIUM_BORDER = Border(left=MEDIUM_SIDE, right=MEDIUM_SIDE, top=MEDIUM_SIDE, bottom=MEDIUM_SIDE)

FILL_ORANGE_HEADER = PatternFill("solid", fgColor="FFEC7600")
FILL_YELLOW = PatternFill("solid", fgColor="FFFFFF00")
FILL_RED = PatternFill("solid", fgColor="FFFF0000")
FILL_ORANGE_P = PatternFill("solid", fgColor="FFFFC266")  # #FFC266 requested for Column P
FILL_GRAY_DATE = PatternFill("solid", fgColor="FFD9D9D9")  # #D9D9D9 requested for Date column

DATE_FORMAT = r"[$-409]d\-mmm\-yy;@"
NUM_FORMAT = r"#,##0"

# Column widths for Sheet 1 (M-DECK THIKA BD)
STATIC_DAILY_WIDTHS = {
    "A": 13.0,
    "B": 9.75,
    "C": 10.25,
    "D": 13.25,
    "E": 11.125,
    "F": 9.875,
    "G": 10.5,
    "H": 11.5,
    "I": 9.75,
    "J": 9.875,
    "K": 13.0,
    "L": 10.375,
    "M": 8.625,
    "N": 10.75,
    "O": 13.875,
    "P": 14.125,
    "Q": 9.125,
    "R": 7.625,
    "S": 11.0,
    "T": 12.0,
    "U": 12.75,
}

# Column widths for Sheet 4 (M-DECK NAIROBI BD)
NAIROBI_DAILY_WIDTHS = {
    "A": 12.625,
    "B": 10.75,
    "C": 10.75,
    "D": 16.125,
    "E": 10.75,
    "O": 13.25,
    "P": 10.75,
    "U": 10.75,
}

# Column widths for Sheet 7 (JUJA WB DAILY SUMMARY)
CONSOLIDATED_DAILY_WIDTHS = {
    "A": 13.0,
    "B": 9.75,
    "C": 10.25,
    "D": 9.5,
    "E": 11.125,
    "F": 9.875,
    "G": 10.5,
    "H": 11.5,
    "I": 9.75,
    "J": 9.875,
    "K": 13.0,
    "L": 10.375,
    "M": 8.625,
    "N": 10.75,
    "O": 13.875,
    "P": 14.125,
    "Q": 9.125,
    "R": 8.875,
    "S": 11.0,
    "T": 12.0,
    "U": 12.75,
    "V": 9.0,
}

# Column widths for Hourly sheets (exact to reference template)
HOURLY_WIDTHS_HSWIM_THIKA = {"A": 10.375, "B": 6.125, "Z": 8.0}
HOURLY_WIDTHS_MDECK_THIKA = {"A": 12.875, "B": 5.875, "D": 6.625, "E": 5.875, "Z": 7.5, "AC": 9.0}
HOURLY_WIDTHS_HSWIM_NRB = {"A": 10.375, "B": 6.125, "Z": 8.0}
HOURLY_WIDTHS_MDECK_NRB = {"A": 14.375, "B": 9.125, "C": 6.125, "F": 6.75, "G": 6.125, "Z": 8.5}
HOURLY_WIDTHS_JUJA_HSWIM = {"A": 11.0, "B": 5.875, "F": 7.375, "G": 5.875, "U": 6.75, "V": 5.875, "Z": 7.5, "AA": 9.0}
HOURLY_WIDTHS_MDECK_SUMMARY = {"A": 13.875, "B": 6.125, "D": 7.75, "E": 6.125, "Y": 6.0, "Z": 7.5, "AA": 9.0}

HOURLY_INTERVALS = [
    "00:00 - 01:00", "01:00 - 02:00", "02:00 - 03:00", "03:00 - 04:00",
    "04:00 - 05:00", "05:00 - 06:00", "06:00 - 07:00", "07:00 - 08:00",
    "08:00 - 09:00", "09:00 - 10:00", "10:00 - 11:00", "11:00 - 12:00",
    "12:00 - 13:00", "13:00 - 14:00", "14:00 - 15:00", "15:00 - 16:00",
    "16:00 - 17:00", "17:00 - 18:00", "18:00 - 19:00", "19:00 - 20:00",
    "20:00 - 21:00", "21:00 - 22:00", "22:00 - 23:00", "23:00 - 00:00",
]


def _safe_int(val: Any, default: int = 0) -> int:
    try:
        if val is None or val == "" or str(val).strip() == "-":
            return default
        return int(float(str(val).replace(",", "").strip()))
    except (ValueError, TypeError, AttributeError):
        return default


def _extract_static_day_metrics(session: Any) -> dict[str, Any]:
    if not session:
        return {
            "Q": 0, "N": 0, "M": 0,
            "buses": 0, "v_3500_7000": 0, "w_7000": 0,
            "Y": 0, "A": 0, "Z": 0, "G": 0, "R": 0, "P": 0, "B": 0,
            "L": 0, "E": 0, "F": 0,
        }

    sections = session.sections or {}
    sec_dh = sections.get("daily_hour") or {}
    dh_summary = sec_dh.get("summary") or {}
    daily_sum_sec = sections.get("daily_summary") or {}
    sum_vals = daily_sum_sec.get("summary_values") or {}
    census_vals = (sections.get("traffic_census") or {}).get("values") or {}
    transgressions_sec = sections.get("transgressions") or {}

    q = _safe_int(dh_summary.get("Q", daily_sum_sec.get("q", 0)))
    m = _safe_int(sum_vals.get("manually_weighed_m", dh_summary.get("M", 0)))

    scale_n = sum_vals.get("weighed_scale_total_n")
    if scale_n is not None:
        n = _safe_int(scale_n)
    else:
        d_val = _safe_int(dh_summary.get("D", 0))
        s_val = _safe_int(dh_summary.get("S", 0))
        if d_val > 0 or s_val > 0:
            n = d_val + s_val
        else:
            x_val = _safe_int(dh_summary.get("X", 0))
            n = max(0, x_val - m)

    buses = _safe_int(census_vals.get("buses_gte_3500kg", 0))
    v_3500_7000 = _safe_int(census_vals.get("vehicles_3500_to_7000_excluding_buses", 0))
    w_7000 = _safe_int(census_vals.get("vehicles_gte_7000_excluding_buses", 0))

    a = _safe_int(sum_vals.get("warned_a", dh_summary.get("A", 0)))
    z = _safe_int(sum_vals.get("charged_prohibited_z", dh_summary.get("Z", 0)))
    g = _safe_int(sum_vals.get("special_release_g", dh_summary.get("G", 0)))
    r = _safe_int(sum_vals.get("vehicles_charged_but_redistributed_r", dh_summary.get("R", 0)))

    y_val = sum_vals.get("total_overload_y", dh_summary.get("Y"))
    if y_val is not None:
        y = _safe_int(y_val)
    else:
        y = a + z + g + r

    p_val = sum_vals.get("impounded_prohibited_p", dh_summary.get("P"))
    if p_val is not None:
        p = _safe_int(p_val)
    else:
        p = z + r

    b_court = _safe_int(sum_vals.get("cases_cleared_in_court_b", 0))
    if b_court == 0:
        man = session.manual_inputs or {}
        b_court = _safe_int(man.get("cases_cleared_in_court", 0))

    l_trans = _safe_int(
        transgressions_sec.get("total_transgressions",
        sum_vals.get("transgression_l", 0))
    )

    permits_not_weighed = _safe_int(
        census_vals.get("exemption_permits_not_weighed",
        sum_vals.get("exemption_permits_not_weighed", dh_summary.get("E", 0)))
    )

    sec_wideload = sections.get("wideload") or {}
    wideload_records = sec_wideload.get("records") or []
    if not wideload_records and session.dataframes:
        df_wl = session.dataframes.get("wideload")
        if df_wl is not None and hasattr(df_wl, "to_dict"):
            wideload_records = df_wl.to_dict("records")

    permits_weighed = len(wideload_records)
    if permits_weighed == 0:
        permits_weighed = _safe_int(sum_vals.get("exemption_permits_weighed", 0))

    return {
        "Q": q,
        "N": n,
        "M": m,
        "buses": buses,
        "v_3500_7000": v_3500_7000,
        "w_7000": w_7000,
        "Y": y,
        "A": a,
        "Z": z,
        "G": g,
        "R": r,
        "P": p,
        "B": b_court,
        "L": l_trans,
        "E": permits_not_weighed,
        "F": permits_weighed,
    }


def _extract_hourly_distribution(session: Any) -> dict[int, dict[str, int]]:
    res = {h: {"Q": 0, "X": 0} for h in range(24)}
    if not session:
        return res

    dh = (session.dataframes or {}).get("daily_hour")
    if dh is None:
        dh_data = (session.sections or {}).get("daily_hour", {}).get("data")
        if dh_data:
            dh = pd.DataFrame(dh_data)

    if dh is not None:
        if not isinstance(dh, pd.DataFrame):
            try:
                dh = pd.DataFrame(dh)
            except Exception:
                dh = None

    if dh is not None and not dh.empty:
        for _, row in dh.iterrows():
            t_str = str(row.get("TIME") or "").strip().replace(":", "")
            if len(t_str) >= 4 and t_str[:2].isdigit():
                h = int(t_str[:2])
                if 0 <= h < 24:
                    res[h]["Q"] = _safe_int(row.get("Q"), 0)
                    res[h]["X"] = _safe_int(row.get("X"), 0)

    return res


def _apply_box_border(
    ws,
    min_row: int,
    max_row: int,
    min_col: int,
    max_col: int,
    outer_side=MEDIUM_SIDE,
    inner_side=None,
):
    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            top = outer_side if r == min_row else inner_side
            bottom = outer_side if r == max_row else inner_side
            left = outer_side if c == min_col else inner_side
            right = outer_side if c == max_col else inner_side
            ws.cell(row=r, column=c).border = Border(
                left=left, right=right, top=top, bottom=bottom
            )


def build_monthly_juja_summary_update_excel_report(year: int, month: int) -> io.BytesIO:
    num_days = calendar.monthrange(year, month)[1]
    dates = [f"{year:04d}-{month:02d}-{day:02d}" for day in range(1, num_days + 1)]
    dates_set = set(dates)

    # Ingest sessions
    static_sessions_thika: dict[str, Any] = {}
    static_sessions_nairobi: dict[str, Any] = {}

    for s, mod_time in report_session_store.list_all_sessions():
        if not s or not s.report_date or s.report_date not in dates_set:
            continue

        st = (s.station or s.weighbridge_name or "").lower()
        b = (s.bound or "").lower()

        if "juja" in st and "mobile" not in st:
            if "thika" in b:
                static_sessions_thika[s.report_date] = s
            elif "nairobi" in b:
                static_sessions_nairobi[s.report_date] = s

    # Parse metrics & hourly distribution
    thika_metrics = {d: _extract_static_day_metrics(static_sessions_thika.get(d)) for d in dates}
    nairobi_metrics = {d: _extract_static_day_metrics(static_sessions_nairobi.get(d)) for d in dates}

    thika_hourly = {d: _extract_hourly_distribution(static_sessions_thika.get(d)) for d in dates}
    nairobi_hourly = {d: _extract_hourly_distribution(static_sessions_nairobi.get(d)) for d in dates}

    wb = Workbook()
    wb.remove(wb.active)

    align_center = Alignment(horizontal="center", vertical="center")
    align_center_wrap = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Fonts for M-DECK THIKA BD (strictly black text as requested)
    font_thika_data = Font(name=FONT_ARIAL, size=11, bold=False, color="000000")
    font_thika_date = Font(name=FONT_ARIAL, size=11, bold=True, color="000000")
    font_thika_tot = Font(name=FONT_ARIAL, size=11, bold=True, color="000000")
    font_thika_tot_u = Font(name=FONT_ARIAL, size=11, bold=False, color="000000")

    # Fonts for M-DECK NAIROBI BD (black text)
    font_nrb_data = Font(name=FONT_ARIAL, size=11, bold=False, color="000000")
    font_nrb_date = Font(name=FONT_ARIAL, size=11, bold=True, color="000000")
    font_nrb_tot = Font(name=FONT_ARIAL, size=11, bold=True, color="000000")

    # Font for JUJA WB DAILY SUMMARY (strictly black text as in reference template)
    font_wb_data = Font(name=FONT_CALIBRI, size=11, bold=False, color="000000")
    font_wb_date = Font(name=FONT_CALIBRI, size=11, bold=True, color="000000")
    font_wb_tot = Font(name=FONT_CALIBRI, size=11, bold=True, color="000000")

    # =========================================================================
    # SHEET 1: M-DECK THIKA BD
    # =========================================================================
    ws1 = wb.create_sheet(title="M-DECK THIKA BD")
    ws1.sheet_properties.tabColor = "FFA366"
    ws1.sheet_format.defaultColWidth = 9.0
    ws1.sheet_format.defaultRowHeight = 14.25
    ws1.page_setup.orientation = "landscape"
    ws1.page_setup.fitToWidth = 1
    ws1.page_setup.fitToHeight = 1
    ws1.sheet_properties.pageSetUpPr.fitToPage = True
    ws1.page_margins.left = 0.25
    ws1.page_margins.right = 0.25
    ws1.page_margins.top = 0.75
    ws1.page_margins.bottom = 0.75

    for col_letter, width in STATIC_DAILY_WIDTHS.items():
        ws1.column_dimensions[col_letter].width = width

    # Row 4: Title
    ws1.row_dimensions[4].height = 23.25
    ws1.merge_cells("A4:U4")
    t1 = ws1.cell(4, 1, "JUJA WEIGHBRIDGE DAILY SUMMARY - THIKA BOUND")
    t1.font = Font(name=FONT_ARIAL, size=18, bold=True, color="000000")
    t1.alignment = align_center
    _apply_box_border(ws1, 4, 4, 1, 21, outer_side=MEDIUM_SIDE, inner_side=None)

    ws1.row_dimensions[5].height = 15.75
    ws1.merge_cells("D5:L5")

    # Row 6: Headers
    ws1.row_dimensions[6].height = 90.0
    thika_headers = [
        "Date",
        "No. of Vehicles weighed HSWIM   (Q)",
        "No. of  Vehicles weighed Scale          (N)",
        "No. of Vehicles weighed manually (M)",
        "Total No. of Vehiclesweighed (X) =(N+M)",
        "No. of Buses ≥3500kg                                                                      (J)",
        "No of Vehicles ≥3500Kg ≤7000Kg      (V)",
        "No. of Vehicles ˃7000 excluding buses (W)",
        "Traffic Census (K)",
        "Total Traffic   (T)=(X+K+E)",
        "Total No. of Vehicles  Overloaded  (Y)            A+Z+G+R",
        "No. of Vehicles warned   (A)",
        "No. of Vehicles charged in court (Z)",
        "Special release (G)",
        "No. of vehicles given warning to offload  or redistribute      (R)",
        "No of vehicles Impounded &  prohibited (P)  Z+R",
        "No of vehicles  cleared in court (B)",
        "Transgressions (L)",
        "Not weighed (E)",
        "Weighed (F)",
        "Total No. of exemption permits",
    ]
    for c_idx, h_text in enumerate(thika_headers, start=1):
        cell = ws1.cell(6, c_idx, h_text)
        f_sz = 10 if c_idx == 19 else 11
        cell.font = Font(name=FONT_ARIAL, size=f_sz, bold=True, color="000000")
        cell.alignment = Alignment(horizontal=None if c_idx == 20 else "center", vertical="top", wrap_text=True)
        cell.border = THIN_BORDER
        if c_idx == 16:
            # P6 background color #FFC266
            cell.fill = FILL_ORANGE_P
        elif c_idx in (5, 9, 10, 21):
            # E6, I6, J6, U6 background color red
            cell.fill = FILL_RED

    # Rows 7 to 37 (31 Day Slots)
    for day_num in range(1, 32):
        cur_r = 6 + day_num
        ws1.row_dimensions[cur_r].height = 15.75 if cur_r == 11 else 15.0

        if day_num <= num_days:
            d = dates[day_num - 1]
            dt_obj = datetime.strptime(d, "%Y-%m-%d")
            mt = thika_metrics[d]

            c_dt = ws1.cell(cur_r, 1, dt_obj)
            c_dt.font = font_thika_date
            c_dt.alignment = align_center
            c_dt.number_format = DATE_FORMAT
            c_dt.border = THIN_BORDER
            c_dt.fill = FILL_GRAY_DATE

            row_vals = {
                2: mt["Q"],
                3: mt["N"],
                4: mt["M"],
                5: f"=C{cur_r}+D{cur_r}",
                6: mt["buses"],
                7: mt["v_3500_7000"],
                8: mt["w_7000"],
                9: f"=F{cur_r}+G{cur_r}+H{cur_r}",
                10: f"=B{cur_r}+E{cur_r}+I{cur_r}+S{cur_r}",
                11: mt["Y"],
                12: mt["A"],
                13: mt["Z"],
                14: mt["G"],
                15: mt["R"],
                16: mt["P"],
                17: mt["B"],
                18: mt["L"],
                19: mt["E"],
                20: mt["F"],
                21: f"=S{cur_r}+T{cur_r}",
            }
        else:
            # Slot beyond month end (e.g. Day 31 for a 30-day month like September)
            c_dt = ws1.cell(cur_r, 1, None)
            c_dt.font = font_thika_date
            c_dt.alignment = align_center
            c_dt.border = THIN_BORDER
            c_dt.fill = FILL_GRAY_DATE

            row_vals = {
                2: None, 3: None, 4: None,
                5: f"=C{cur_r}+D{cur_r}",
                6: None, 7: None, 8: None,
                9: f"=F{cur_r}+G{cur_r}+H{cur_r}",
                10: f"=B{cur_r}+E{cur_r}+I{cur_r}+S{cur_r}",
                11: None, 12: None, 13: None, 14: None, 15: None,
                16: None,
                17: None, 18: None, 19: None, 20: None,
                21: f"=S{cur_r}+T{cur_r}",
            }

        for c_idx in range(2, 22):
            val = row_vals[c_idx]
            cell = ws1.cell(cur_r, c_idx, val)
            cell.alignment = align_center
            cell.border = THIN_BORDER
            # Black text for ALL values in M-DECK THIKA BD
            cell.font = font_thika_data

            if val is not None and not str(val).startswith("="):
                cell.number_format = NUM_FORMAT
            elif str(val).startswith("="):
                cell.number_format = NUM_FORMAT

            if c_idx == 2:
                cell.fill = FILL_YELLOW
            elif c_idx in (5, 9, 10, 21):
                cell.fill = FILL_RED
            elif c_idx == 16:
                # P7:P37 background color #FFC266
                cell.fill = FILL_ORANGE_P

    # Sheet 1 TOTAL Row: Row 38
    r_s1_tot = 38
    ws1.row_dimensions[r_s1_tot].height = 15.0
    c_tot1 = ws1.cell(r_s1_tot, 1, "TOTAL")
    c_tot1.font = font_thika_date
    c_tot1.alignment = align_center
    c_tot1.border = THIN_BORDER

    for c_idx in range(2, 21):
        col_let = get_column_letter(c_idx)
        cell = ws1.cell(r_s1_tot, c_idx, f"=SUM({col_let}7:{col_let}37)")
        cell.font = font_thika_tot
        cell.alignment = align_center
        cell.border = THIN_BORDER
        cell.number_format = NUM_FORMAT
        cell.fill = FILL_RED

    cell_u1 = ws1.cell(r_s1_tot, 21, "=S38+T38")
    cell_u1.font = font_thika_tot_u
    cell_u1.alignment = align_center
    cell_u1.border = THIN_BORDER
    cell_u1.number_format = NUM_FORMAT
    cell_u1.fill = FILL_RED

    ws1.print_area = "A1:U39"

    # =========================================================================
    # SHEET 4: M-DECK NAIROBI BD
    # =========================================================================
    ws4 = wb.create_sheet(title="M-DECK NAIROBI BD")
    ws4.sheet_properties.tabColor = "FFC266"
    ws4.sheet_format.defaultColWidth = 9.0
    ws4.sheet_format.defaultRowHeight = 14.25
    ws4.page_setup.orientation = "landscape"
    ws4.page_setup.fitToWidth = 1
    ws4.page_setup.fitToHeight = 1
    ws4.sheet_properties.pageSetUpPr.fitToPage = True
    ws4.page_margins.left = 0.157
    ws4.page_margins.right = 0.157
    ws4.page_margins.top = 0.197
    ws4.page_margins.bottom = 0.197

    for col_letter, width in NAIROBI_DAILY_WIDTHS.items():
        ws4.column_dimensions[col_letter].width = width

    # Row 4: Title
    ws4.row_dimensions[4].height = 23.25
    ws4.merge_cells("A4:U4")
    t4 = ws4.cell(4, 1, "JUJA WEIGHBRIDGE DAILY SUMMARY - NAIROBI BOUND")
    t4.font = Font(name=FONT_ARIAL, size=18, bold=True, color="000000")
    t4.alignment = align_center
    _apply_box_border(ws4, 4, 4, 1, 21, outer_side=MEDIUM_SIDE, inner_side=None)

    ws4.row_dimensions[5].height = 15.75
    ws4.merge_cells("D5:L5")

    # Row 6: Headers
    ws4.row_dimensions[6].height = 90.0
    nairobi_headers = [
        "Date",
        "No. of Vehicles weighed HSWIM   (Q)",
        "No. of  Vehicles weighed Scale      (N)",
        "No. of Vehicles weighed manually (M)",
        "Total No. of Vehiclesweighed (X) =(N+M)",
        "No. of Buses ≥3500kg (J)",
        "No of Vehicles ≥3500Kg ≤7000Kg (V)",
        "No. of Vehicles ˃7000 excluding buses (W)",
        "Traffic Census (K)",
        "Total Traffic   (T)=(X+K+E)",
        "Total No. of Vehicles  Overloaded  (Y)            A+Z+G+R",
        "No. of Vehicles warned   (A)",
        "No. of Vehicles charged in court (Z)",
        "Special release (G)",
        "No. of vehicles given warning to offload  or redistribute      (R)",
        "No of vehicles Impounded &  prohibited (P)  Z+R",
        "No of vehicles  cleared in court (B)",
        "Transgressions (L)",
        "Not weighed (E)",
        "Weighed (F)",
        "Total No. of exemption permits",
    ]
    for c_idx, h_text in enumerate(nairobi_headers, start=1):
        cell = ws4.cell(6, c_idx, h_text)
        f_sz = 10 if c_idx == 19 else 11
        f_name = FONT_ARIAL if c_idx == 21 else FONT_CALIBRI
        cell.font = Font(name=f_name, size=f_sz, bold=True, color="000000")
        cell.alignment = Alignment(horizontal=None if c_idx == 20 else "center", vertical="top", wrap_text=True)
        cell.border = THIN_BORDER
        if c_idx == 16:
            cell.fill = FILL_ORANGE_P
        elif c_idx in (5, 9, 10, 21):
            cell.fill = FILL_RED

    # Rows 7 to 37 (31 Day Slots)
    for day_num in range(1, 32):
        cur_r = 6 + day_num
        ws4.row_dimensions[cur_r].height = 15.0

        if day_num <= num_days:
            d = dates[day_num - 1]
            dt_obj = datetime.strptime(d, "%Y-%m-%d")
            mn = nairobi_metrics[d]

            c_dt = ws4.cell(cur_r, 1, dt_obj)
            c_dt.font = font_nrb_date
            c_dt.alignment = align_center
            c_dt.number_format = DATE_FORMAT
            c_dt.border = THIN_BORDER
            c_dt.fill = FILL_GRAY_DATE

            row_vals = {
                2: mn["Q"],
                3: mn["N"],
                4: mn["M"],
                5: f"=C{cur_r}+D{cur_r}",
                6: mn["buses"],
                7: mn["v_3500_7000"],
                8: mn["w_7000"],
                9: f"=F{cur_r}+G{cur_r}+H{cur_r}",
                10: f"=B{cur_r}+E{cur_r}+I{cur_r}+S{cur_r}",
                11: mn["Y"],
                12: mn["A"],
                13: mn["Z"],
                14: mn["G"],
                15: mn["R"],
                16: mn["P"],
                17: mn["B"],
                18: mn["L"],
                19: mn["E"],
                20: mn["F"],
                21: f"=S{cur_r}+T{cur_r}",
            }
        else:
            c_dt = ws4.cell(cur_r, 1, None)
            c_dt.font = font_nrb_date
            c_dt.alignment = align_center
            c_dt.border = THIN_BORDER
            c_dt.fill = FILL_GRAY_DATE

            row_vals = {
                2: None, 3: None, 4: None,
                5: f"=C{cur_r}+D{cur_r}",
                6: None, 7: None, 8: None,
                9: f"=F{cur_r}+G{cur_r}+H{cur_r}",
                10: f"=B{cur_r}+E{cur_r}+I{cur_r}+S{cur_r}",
                11: None, 12: None, 13: None, 14: None, 15: None,
                16: None,
                17: None, 18: None, 19: None, 20: None,
                21: f"=S{cur_r}+T{cur_r}",
            }

        for c_idx in range(2, 22):
            val = row_vals[c_idx]
            cell = ws4.cell(cur_r, c_idx, val)
            cell.alignment = align_center
            cell.border = THIN_BORDER
            cell.font = font_nrb_data

            if val is not None and not str(val).startswith("="):
                cell.number_format = NUM_FORMAT
            elif str(val).startswith("="):
                cell.number_format = NUM_FORMAT

            if c_idx == 2:
                cell.fill = FILL_YELLOW
            elif c_idx in (5, 9, 10, 21):
                cell.fill = FILL_RED
            elif c_idx == 16:
                cell.fill = FILL_ORANGE_P

    # Sheet 4 TOTAL Row: Row 38
    r_s4_tot = 38
    ws4.row_dimensions[r_s4_tot].height = 15.0
    c_tot4 = ws4.cell(r_s4_tot, 1, "TOTAL")
    c_tot4.font = font_nrb_date
    c_tot4.alignment = align_center
    c_tot4.border = THIN_BORDER

    for c_idx in range(2, 21):
        col_let = get_column_letter(c_idx)
        cell = ws4.cell(r_s4_tot, c_idx, f"=SUM({col_let}7:{col_let}37)")
        cell.font = font_nrb_tot
        cell.alignment = align_center
        cell.border = THIN_BORDER
        cell.number_format = NUM_FORMAT
        cell.fill = FILL_RED

    cell_u4 = ws4.cell(r_s4_tot, 21, "=SUM(U7:U37)")
    cell_u4.font = font_nrb_tot
    cell_u4.alignment = align_center
    cell_u4.border = THIN_BORDER
    cell_u4.number_format = NUM_FORMAT
    cell_u4.fill = FILL_RED

    ws4.print_area = "A1:U39"

    # =========================================================================
    # SHEET 7: JUJA WB DAILY SUMMARY
    # =========================================================================
    ws7 = wb.create_sheet(title="JUJA WB DAILY SUMMARY")
    ws7.sheet_format.defaultColWidth = 9.0
    ws7.sheet_format.defaultRowHeight = 14.25
    ws7.page_setup.orientation = "landscape"
    ws7.page_setup.fitToWidth = 1
    ws7.page_setup.fitToHeight = 1
    ws7.sheet_properties.pageSetUpPr.fitToPage = True
    ws7.page_margins.left = 0.157
    ws7.page_margins.right = 0.157
    ws7.page_margins.top = 0.197
    ws7.page_margins.bottom = 0.197

    for col_letter, width in CONSOLIDATED_DAILY_WIDTHS.items():
        ws7.column_dimensions[col_letter].width = width

    # Row 4: Title
    ws7.row_dimensions[4].height = 23.25
    ws7.merge_cells("A4:U4")
    t7 = ws7.cell(4, 1, "JUJA  WEIGHBRIDGE DAILY SUMMARY")
    t7.font = Font(name=FONT_ARIAL, size=18, bold=True, color="000000")
    t7.alignment = align_center
    _apply_box_border(ws7, 4, 4, 1, 21, outer_side=MEDIUM_SIDE, inner_side=None)

    ws7.row_dimensions[5].height = 15.0
    ws7.merge_cells("D5:L5")

    # Rows 6-8: Multi-row Headers
    ws7.row_dimensions[6].height = 45.0
    ws7.row_dimensions[7].height = 15.0
    ws7.row_dimensions[8].height = 44.25

    wb_daily_headers_r6 = {
        1: "Date",
        2: "No. of Vehicles weighed HSWIM   (Q)",
        3: "No. of  Vehicles weighed Scale      (S)",
        4: "No. of Vehicles weighed manually",
        5: "Total No. of Vehiclesweighed (X)",
        6: "No. of Buses ≥3500kg",
        7: "No of Vehicles ≥3500Kg ≤7000Kg",
        8: "No. of Vehicles ˃7000 excluding buses",
        9: "Traffic Census (K)",
        10: "Total Traffic (T) ",
        11: "Total No. of Vehicles  Overloaded  (Y)            A+Z+G+R",
        12: "No. of Vehicles warned   (A)",
        13: "No. of Vehicles charged in court (Z)",
        14: "No. of Vehicles given special release (G)",
        15: "No. of vehicles given warning to offload  or redistribute      (R)",
        16: "No of vehicles Impounded &  prohibited                (P)          Z+R",
        17: "No of vehicles  cleared in court (B)",
        18: "Transgressions",
        19: "Exemption permits",
        20: "Exemption permits",
        21: "Total No. of exemption permits",
    }

    # Vertical merges for cols 1..18 and 21
    for c_idx in list(range(1, 19)) + [21]:
        ws7.merge_cells(start_row=6, start_column=c_idx, end_row=8, end_column=c_idx)
        cell = ws7.cell(6, c_idx, wb_daily_headers_r6[c_idx])
        cell.font = Font(name=FONT_CALIBRI, size=11, bold=True, color="000000")
        cell.alignment = align_center_wrap
        for r_m in range(6, 9):
            ws7.cell(r_m, c_idx).border = THIN_BORDER

    # Cols S (19) and T (20)
    c_s6 = ws7.cell(6, 19, "Exemption permits")
    c_s6.font = Font(name=FONT_CALIBRI, size=11, bold=True, color="000000")
    c_s6.alignment = align_center_wrap
    c_s6.border = THIN_BORDER

    c_t6 = ws7.cell(6, 20, "Exemption permits")
    c_t6.font = Font(name=FONT_CALIBRI, size=11, bold=True, color="000000")
    c_t6.alignment = align_center_wrap
    c_t6.border = THIN_BORDER

    c_s7 = ws7.cell(7, 19, "(E)")
    c_s7.font = Font(name=FONT_CALIBRI, size=11, bold=True, color="000000")
    c_s7.alignment = align_center
    c_s7.border = THIN_BORDER
    ws7.cell(7, 20).border = THIN_BORDER

    c_s8 = ws7.cell(8, 19, "Not Weighed")
    c_s8.font = Font(name=FONT_CALIBRI, size=11, bold=True, color="000000")
    c_s8.alignment = align_center_wrap
    c_s8.border = THIN_BORDER

    c_t8 = ws7.cell(8, 20, "Weighed")
    c_t8.font = Font(name=FONT_CALIBRI, size=11, bold=True, color="000000")
    c_t8.alignment = align_center_wrap
    c_t8.border = THIN_BORDER

    # Data Rows 9 to 39 (31 Day Slots)
    for day_num in range(1, 32):
        cur_r = 8 + day_num
        ws7.row_dimensions[cur_r].height = 15.0
        r_ref = 6 + day_num  # Corresponding row 7..37 in Sheets 1 & 4

        if day_num <= num_days:
            d = dates[day_num - 1]
            dt_obj = datetime.strptime(d, "%Y-%m-%d")
            c_dt = ws7.cell(cur_r, 1, dt_obj)
            c_dt.font = font_wb_date
            c_dt.alignment = align_center
            c_dt.number_format = DATE_FORMAT
            c_dt.border = THIN_BORDER
        else:
            c_dt = ws7.cell(cur_r, 1, None)
            c_dt.font = font_wb_date
            c_dt.alignment = align_center
            c_dt.border = THIN_BORDER

        for c_idx in range(2, 22):
            col_let = get_column_letter(c_idx)
            formula = f"='M-DECK THIKA BD'!{col_let}{r_ref}+'M-DECK NAIROBI BD'!{col_let}{r_ref}"
            cell = ws7.cell(cur_r, c_idx, formula)
            cell.font = font_wb_data
            cell.alignment = align_center
            cell.border = THIN_BORDER
            cell.number_format = NUM_FORMAT
            cell.fill = FILL_RED

    # Sheet 7 TOTAL Row: Row 40
    r_s7_tot = 40
    ws7.row_dimensions[r_s7_tot].height = 15.0
    c_tot7 = ws7.cell(r_s7_tot, 1, "TOTAL")
    c_tot7.font = Font(name=FONT_CALIBRI, size=11, bold=True, color="000000")
    c_tot7.alignment = align_center
    c_tot7.border = THIN_BORDER

    for c_idx in range(2, 22):
        col_let = get_column_letter(c_idx)
        cell = ws7.cell(r_s7_tot, c_idx, f"=SUM({col_let}9:{col_let}39)")
        cell.font = font_wb_tot
        cell.alignment = align_center
        cell.border = THIN_BORDER
        cell.number_format = NUM_FORMAT
        cell.fill = FILL_RED

    ws7.row_dimensions[41].height = 15.0
    ws7.print_area = "A1:U42"

    # =========================================================================
    # HOURLY SHEETS HELPER (31 Day Slots: Rows 2 to 32, Total Row 33, Footer Row 35)
    # =========================================================================
    def build_hourly_sheet(
        ws,
        sheet_type: str,  # "direct_q", "direct_x", "sum_q", "sum_x"
        hourly_data: dict[str, dict[int, dict[str, int]]] | None,
        source_sheet1: str | None,
        source_sheet2: str | None,
        total_label: str,
        footer_letter: str,
        footer_desc: str,
        col_widths: dict[str, float],
        orientation: str = "portrait",
        header_row_height: float = 45.0,
        total_row_height: float = 30.0,
        default_row_height: float = 14.25,
        custom_row_heights: dict[int, float] | None = None,
        data_fill: PatternFill | None = None,
        date_fill: PatternFill | None = FILL_GRAY_DATE,
        total_fill: PatternFill | None = FILL_RED,
        col_z_fill: PatternFill | None = FILL_RED,
        total_label_fill: PatternFill | None = FILL_GRAY_DATE,
        total_label_align: Alignment | None = None,
        tab_color: str | None = None,
    ):
        if tab_color:
            ws.sheet_properties.tabColor = tab_color
        ws.sheet_format.defaultColWidth = 9.0
        ws.sheet_format.defaultRowHeight = default_row_height
        ws.page_setup.orientation = orientation
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_margins.left = 0.7
        ws.page_margins.right = 0.7
        ws.page_margins.top = 0.75
        ws.page_margins.bottom = 0.75

        for col_letter, width in col_widths.items():
            ws.column_dimensions[col_letter].width = width

        # Row 1: Headers
        ws.row_dimensions[1].height = header_row_height
        c1 = ws.cell(1, 1, "DATE")
        c1.font = Font(name=FONT_CALIBRI, size=11, bold=True, color="000000")
        c1.alignment = Alignment(horizontal="center")
        c1.fill = FILL_ORANGE_HEADER
        c1.border = Border(left=THIN_SIDE, right=None, top=THIN_SIDE, bottom=THIN_SIDE)

        align_vert_center = Alignment(textRotation=90, horizontal="center", vertical="center", wrap_text=True)

        for h_idx, h_text in enumerate(HOURLY_INTERVALS, start=2):
            cell = ws.cell(1, h_idx, h_text)
            cell.font = Font(name=FONT_CALIBRI, size=11, bold=True, color="000000")
            cell.alignment = align_vert_center
            cell.fill = FILL_ORANGE_HEADER
            cell.border = THIN_BORDER

        c_tot = ws.cell(1, 26, "Total Weighed per day")
        c_tot.font = Font(name=FONT_CALIBRI, size=11, bold=True, color="000000")
        c_tot.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c_tot.fill = FILL_ORANGE_HEADER
        c_tot.border = THIN_BORDER

        # Data Rows 2 to 32 (31 Day Slots)
        for day_num in range(1, 32):
            cur_r = 1 + day_num
            if default_row_height == 14.25:
                ws.row_dimensions[cur_r].height = 15.0

            if day_num <= num_days:
                d = dates[day_num - 1]
                dt_obj = datetime.strptime(d, "%Y-%m-%d")
                c_dt = ws.cell(cur_r, 1, dt_obj)
                c_dt.font = Font(name=FONT_CALIBRI, size=11, bold=True, color="000000")
                c_dt.alignment = Alignment(horizontal="center", vertical="center")
                c_dt.number_format = DATE_FORMAT
                c_dt.border = Border(left=None, right=THIN_SIDE, top=THIN_SIDE, bottom=None)
                if date_fill:
                    c_dt.fill = date_fill

                for h in range(24):
                    col_idx = 2 + h
                    col_let = get_column_letter(col_idx)

                    if sheet_type == "direct_q":
                        val = hourly_data[d][h]["Q"] if hourly_data else 0
                    elif sheet_type == "direct_x":
                        val = hourly_data[d][h]["X"] if hourly_data else 0
                    elif sheet_type in ("sum_q", "sum_x"):
                        val = f"='{source_sheet1}'!{col_let}{cur_r}+'{source_sheet2}'!{col_let}{cur_r}"
                    else:
                        val = 0

                    cell = ws.cell(cur_r, col_idx, val)
                    cell.font = Font(name=FONT_CALIBRI, size=11, bold=False, color="000000")
                    cell.alignment = Alignment(horizontal=None, vertical=None)
                    cell.border = THIN_BORDER
                    cell.number_format = NUM_FORMAT
                    if data_fill:
                        cell.fill = data_fill

                # Col Z: Total Weighed per day formula
                cell_z = ws.cell(cur_r, 26, f"=SUM(B{cur_r}:Y{cur_r})")
                cell_z.font = Font(name=FONT_CALIBRI, size=11, bold=True, color="000000")
                cell_z.alignment = Alignment(horizontal=None, vertical=None)
                cell_z.border = THIN_BORDER
                cell_z.number_format = NUM_FORMAT
                if col_z_fill:
                    cell_z.fill = col_z_fill
            else:
                c_dt = ws.cell(cur_r, 1, None)
                c_dt.font = Font(name=FONT_CALIBRI, size=11, bold=True, color="000000")
                c_dt.alignment = Alignment(horizontal="center", vertical="center")
                c_dt.border = Border(left=None, right=THIN_SIDE, top=THIN_SIDE, bottom=None)
                if date_fill:
                    c_dt.fill = date_fill

                for h in range(24):
                    col_idx = 2 + h
                    col_let = get_column_letter(col_idx)
                    if sheet_type in ("sum_q", "sum_x"):
                        val = f"='{source_sheet1}'!{col_let}{cur_r}+'{source_sheet2}'!{col_let}{cur_r}"
                    else:
                        val = None

                    cell = ws.cell(cur_r, col_idx, val)
                    cell.font = Font(name=FONT_CALIBRI, size=11, bold=False, color="000000")
                    cell.alignment = Alignment(horizontal=None, vertical=None)
                    cell.border = THIN_BORDER
                    if data_fill:
                        cell.fill = data_fill

                cell_z = ws.cell(cur_r, 26, f"=SUM(B{cur_r}:Y{cur_r})")
                cell_z.font = Font(name=FONT_CALIBRI, size=11, bold=True, color="000000")
                cell_z.alignment = Alignment(horizontal=None, vertical=None)
                cell_z.border = THIN_BORDER
                if col_z_fill:
                    cell_z.fill = col_z_fill

        # Apply any custom row heights
        if custom_row_heights:
            for r_num, h_val in custom_row_heights.items():
                ws.row_dimensions[r_num].height = h_val

        # TOTAL Row: Row 33
        r_tot = 33
        ws.row_dimensions[r_tot].height = total_row_height
        c_tot_lbl = ws.cell(r_tot, 1, total_label)
        c_tot_lbl.font = Font(name=FONT_CALIBRI, size=11, bold=True, color="000000")
        c_tot_lbl.alignment = total_label_align or Alignment(wrap_text=True)
        c_tot_lbl.border = THIN_BORDER
        if total_label_fill:
            c_tot_lbl.fill = total_label_fill

        for col_idx in range(2, 26):
            col_let = get_column_letter(col_idx)
            cell = ws.cell(r_tot, col_idx, f"=SUM({col_let}2:{col_let}32)")
            cell.font = Font(name=FONT_CALIBRI, size=11, bold=True, color="000000")
            cell.alignment = Alignment(horizontal=None, vertical=None)
            cell.border = Border(left=THIN_SIDE, right=THIN_SIDE, top=None, bottom=THIN_SIDE)
            cell.number_format = NUM_FORMAT
            if total_fill:
                cell.fill = total_fill

        c_ztot = ws.cell(r_tot, 26, f"=SUM(Z2:Z32)")
        c_ztot.font = Font(name=FONT_CALIBRI, size=11, bold=True, color="000000")
        c_ztot.alignment = Alignment(horizontal=None, vertical=None)
        c_ztot.border = Border(left=THIN_SIDE, right=THIN_SIDE, top=None, bottom=THIN_SIDE)
        c_ztot.number_format = NUM_FORMAT
        if total_fill or col_z_fill:
            c_ztot.fill = total_fill or col_z_fill

        # Footer indicator: Row 35
        r_foot = 35
        f_lbl = ws.cell(r_foot, 1, footer_letter)
        f_lbl.font = Font(name=FONT_CALIBRI, size=11, bold=True, color="000000")
        f_desc = ws.cell(r_foot, 2, footer_desc)
        f_desc.font = Font(name=FONT_CALIBRI, size=11, bold=True, color="000000")

    # Sheet 2: HSWIM THIKA BD HOURLY  (Note the exact trailing space)
    ws2 = wb.create_sheet(title="HSWIM THIKA BD HOURLY ")
    build_hourly_sheet(
        ws2,
        sheet_type="direct_q",
        hourly_data=thika_hourly,
        source_sheet1=None,
        source_sheet2=None,
        total_label="Total Weighed per Hour (Q)",
        footer_letter="Q=",
        footer_desc="HSWIM - CLEARED",
        col_widths=HOURLY_WIDTHS_HSWIM_THIKA,
        orientation="portrait",
        header_row_height=45.0,
        total_row_height=48.0,
        default_row_height=14.25,
        total_label_align=Alignment(wrap_text=True),
        tab_color="FFA366",
    )

    # Sheet 3: M-DECK THIKA BD HOURLY
    ws3 = wb.create_sheet(title="M-DECK THIKA BD HOURLY")
    build_hourly_sheet(
        ws3,
        sheet_type="direct_x",
        hourly_data=thika_hourly,
        source_sheet1=None,
        source_sheet2=None,
        total_label="Total HSWIM cleared (Q)",
        footer_letter="x=",
        footer_desc="N+M",
        col_widths=HOURLY_WIDTHS_MDECK_THIKA,
        orientation="landscape",
        header_row_height=60.0,
        total_row_height=30.0,
        default_row_height=14.25,
        custom_row_heights={6: 16.5, 35: 25.5, 39: 15.0, 40: 15.0},
        total_label_align=Alignment(wrap_text=True),
        tab_color="FFA366",
    )

    # Sheet 5: HSWIM NRB BD HOURLY
    ws5 = wb.create_sheet(title="HSWIM NRB BD HOURLY")
    build_hourly_sheet(
        ws5,
        sheet_type="direct_q",
        hourly_data=nairobi_hourly,
        source_sheet1=None,
        source_sheet2=None,
        total_label="Total Weighed per Hour (Q)",
        footer_letter="Q=",
        footer_desc="HSWIM - CLEARED",
        col_widths=HOURLY_WIDTHS_HSWIM_NRB,
        orientation="portrait",
        header_row_height=45.0,
        total_row_height=48.0,
        default_row_height=14.25,
        total_label_align=Alignment(wrap_text=True),
        tab_color="FFC266",
    )

    # Sheet 6: M-DECK NRB BD HOURLY
    ws6 = wb.create_sheet(title="M-DECK NRB BD HOURLY")
    build_hourly_sheet(
        ws6,
        sheet_type="direct_x",
        hourly_data=nairobi_hourly,
        source_sheet1=None,
        source_sheet2=None,
        total_label="Total Weighed per Hour",
        footer_letter="x=",
        footer_desc="N+M",
        col_widths=HOURLY_WIDTHS_MDECK_NRB,
        orientation="portrait",
        header_row_height=45.0,
        total_row_height=30.0,
        default_row_height=14.25,
        custom_row_heights={19: 18.0, 39: 15.0, 40: 15.0},
        total_label_align=Alignment(wrap_text=True),
        tab_color="FFC266",
    )

    # Sheet 8: JUJA HSWIM summary
    ws8 = wb.create_sheet(title="JUJA HSWIM summary")
    build_hourly_sheet(
        ws8,
        sheet_type="sum_q",
        hourly_data=None,
        source_sheet1="HSWIM THIKA BD HOURLY ",
        source_sheet2="HSWIM NRB BD HOURLY",
        total_label="Total HSWIM cleared (Q)",
        footer_letter="Q=",
        footer_desc="HSWIM - CLEARED",
        col_widths=HOURLY_WIDTHS_JUJA_HSWIM,
        orientation="landscape",
        header_row_height=60.0,
        total_row_height=30.0,
        default_row_height=15.0,
        data_fill=FILL_RED,
        total_label_align=Alignment(wrap_text=True),
    )

    # Sheet 9: M-Decks hourly summary
    ws9 = wb.create_sheet(title="M-Decks hourly summary")
    build_hourly_sheet(
        ws9,
        sheet_type="sum_x",
        hourly_data=None,
        source_sheet1="M-DECK THIKA BD HOURLY",
        source_sheet2="M-DECK NRB BD HOURLY",
        total_label="Total Weighed per Hour",
        footer_letter="X=",
        footer_desc="N+M",
        col_widths=HOURLY_WIDTHS_MDECK_SUMMARY,
        orientation="landscape",
        header_row_height=60.0,
        total_row_height=30.0,
        default_row_height=15.0,
        custom_row_heights={16: 14.25},
        data_fill=FILL_RED,
        total_label_align=Alignment(wrap_text=True),
    )

    # Ensure sheets are in the exact reference order:
    sheet_order = [
        "M-DECK THIKA BD",
        "HSWIM THIKA BD HOURLY ",
        "M-DECK THIKA BD HOURLY",
        "M-DECK NAIROBI BD",
        "HSWIM NRB BD HOURLY",
        "M-DECK NRB BD HOURLY",
        "JUJA WB DAILY SUMMARY",
        "JUJA HSWIM summary",
        "M-Decks hourly summary",
    ]
    wb._sheets = [wb[s] for s in sheet_order]

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output
