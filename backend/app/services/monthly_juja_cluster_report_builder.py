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
DOUBLE_SIDE = Side(style="double", color="000000")

THIN_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)
MEDIUM_BORDER = Border(left=MEDIUM_SIDE, right=MEDIUM_SIDE, top=MEDIUM_SIDE, bottom=MEDIUM_SIDE)
TOTAL_STATIC_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=MEDIUM_SIDE)
TOTAL_MOBILE_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=MEDIUM_SIDE, bottom=MEDIUM_SIDE)

FILL_LIGHT_GRAY = PatternFill("solid", fgColor="F2F2F2")
FILL_RED = PatternFill("solid", fgColor="FFFF0000")
FILL_LIGHT_RED_4 = PatternFill("solid", fgColor="FFFFD7D7")
COLOR_MAGENTA = "FFBF0041"
FILL_GRAY_15 = PatternFill("solid", fgColor="D9D9D9")
FILL_GRAY_25 = PatternFill("solid", fgColor="BFBFBF")
FILL_YELLOW = PatternFill("solid", fgColor="FFFFFF00")

DATE_FORMAT = r"[$-409]d\-mmm\-yy;@"
NUM_FORMAT = r"#,##0"

STATIC_COLUMN_WIDTHS = {
    "A": 18.43,
    "B": 14.42,
    "C": 11.58,
    "D": 11.58,
    "E": 11.58,
    "F": 11.85,
    "G": 11.14,
    "H": 12.29,
    "I": 11.43,
    "J": 12.85,
    "K": 13.43,
    "L": 11.43,
    "M": 15.00,
    "N": 12.43,
    "O": 11.43,
    "P": 11.00,
    "Q": 14.28,
    "R": 14.00,
    "S": 14.00,
    "T": 11.58,
    "U": 13.71,
    "V": 15.29,
    "W": 12.85,
    "X": 13.57,
    "Y": 9.14,
}

DEPLOYMENT_COLUMN_WIDTHS = {
    "A": 19.57,
    "B": 15.85,
    "C": 20.85,
    "D": 11.58,
    "E": 9.00,
    "F": 9.14,
    "G": 13.43,
    "H": 19.00,
    "I": 19.28,
    "J": 48.57,
    "K": 51.29,
    "L": 13.14,
    "M": 215.00,
    "N": 13.71,
    "O": 13.71,
    "P": 13.00,
    "Q": 18.85,
}

CHARGED_COLUMN_WIDTHS = {
    "A": 20.00,
    "B": 16.14,
    "C": 37.58,
    "D": 10.43,
    "E": 11.14,
    "F": 17.29,
    "G": 19.42,
    "H": 19.00,
    "I": 19.28,
    "J": 56.29,
    "K": 54.71,
    "L": 24.14,
    "M": 243.71,
    "N": 9.00,
}

NOTES_TEXTS = [
    "No. of Vehicles weighed HSWIM   (Q)",
    "No. of  Vehicles weighed Scale          (N)",
    "No. of Vehicles weighed manually (M)",
    "No. of Vehicles of Vehiclesweighed (X) =(N+M)",
    "No. of Buses ≥3500kg  (J)",
    "No. of Vehicles ≥3500Kg ≤7000Kg      (V)",
    "No. of Vehicles ˃7000 excluding buses (W)",
    "No.Traffic Census (K)",
    "No.Total Traffic   (T)=(X+K+E)",
    "Total No. of Vehicles  Overloaded  (Y) = (A+Z+G+R)",
    "No. of Vehicles warned   (A)",
    "No. of Vehicles charged in court (Z)",
    "No. Special release (G)",
    "No. of vehicles given warning to offload  or redistribute     (R)",
    "No. of vehicles Impounded &  prohibited (P)  (Z+R)",
    "No. of vehicles  cleared in court (B)",
    "No. Transgressions (L)",
    "No. Exemption permits NOT weighed (E)",
    "No. Exemption permits weighed (F)",
]


def _safe_int(val: Any, default: int = 0) -> int:
    try:
        if val is None or val == "" or str(val).strip() == "-":
            return default
        return int(float(str(val).replace(",", "").strip()))
    except (ValueError, TypeError, AttributeError):
        return default


def _upper(val: Any) -> str:
    if val is None:
        return ""
    return str(val).strip().upper()


def _mobile_report_slot(bound_name: str | None) -> str:
    bound = (bound_name or "").strip().lower()
    if "2" in bound or "two" in bound:
        return "mobile_2"
    return "mobile_1"


def _extract_mobile_session_stats(session) -> dict[str, Any]:
    summary = session.sections.get("mobile_report", {}).get("summary", {})
    if not summary and hasattr(session, "dataframes") and "mobile_report" in session.dataframes:
        from app.services.mobile_report_processor import summarize_mobile_report
        df = session.dataframes["mobile_report"]
        if isinstance(df, pd.DataFrame) and not df.empty:
            summary = summarize_mobile_report(df)

    weighed = _safe_int(summary.get("total_trucks_weighed", 0))
    warned = _safe_int(summary.get("warned_trucks", 0))
    charged = _safe_int(summary.get("charged_trucks", 0))

    manual = (
        session.manual_inputs.get("mobile_report")
        or session.manual_inputs.get("extra", {}).get("mobile_report")
        or {}
    )
    cases_cleared = _safe_int(manual.get("cases_cleared_in_court", 0))
    transgressions = _safe_int(manual.get("transgressions_count", 0))
    exempted = _safe_int(manual.get("exempted_permit", 0))
    manually_weighed = _safe_int(manual.get("manually_weighed", 0))

    return {
        "weighed": weighed,
        "warned_charged": warned + charged,
        "charged": charged,
        "warned": warned,
        "cases_cleared": cases_cleared,
        "transgressions": transgressions,
        "exempted": exempted,
        "manually_weighed": manually_weighed,
    }


def _extract_static_session_metrics(session) -> dict[str, int]:
    if not session:
        return {
            "Q": 0, "H": 0, "C": 0, "N": 0, "M": 0,
            "buses": 0, "v_3500_7000": 0, "w_7000": 0,
            "Y": 0, "A": 0, "Z": 0, "G": 0, "R": 0, "P": 0,
            "B": 0, "L": 0, "E": 0, "F": 0,
        }

    sum_vals = session.sections.get("daily_summary", {}).get("values") or {}
    dh_summary = session.sections.get("daily_hour", {}).get("summary") or {}
    census_vals = (
        session.manual_inputs.get("traffic_census")
        or session.sections.get("traffic_census", {}).get("values")
        or {}
    )

    q = _safe_int(sum_vals.get("weighed_by_hswim_q", dh_summary.get("Q", 0)))
    c = _safe_int(dh_summary.get("C", 0))
    h = _safe_int(dh_summary.get("H", 0))
    if h == 0 and (q > 0 or c > 0):
        h = q + c

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
        b_court = _safe_int(man.get("cases_cleared_in_court") or man.get("cases_cleared_court", 0))

    transg_sec = session.sections.get("transgressions", {})
    l_transg = _safe_int(
        sum_vals.get(
            "transgressions_l",
            transg_sec.get("daily_transgressions_count", 0),
        )
    )

    wideload_sec = session.sections.get("wideload", {})
    e_exemp = _safe_int(
        sum_vals.get(
            "exemption_permits_not_weighed_e",
            wideload_sec.get("wideload_count", dh_summary.get("E", 0)),
        )
    )

    overloaded_sec = session.sections.get("overloaded", {})
    f_exemp = _safe_int(
        sum_vals.get(
            "exemption_permits_weighed_f",
            overloaded_sec.get("valid_permit_count", 0),
        )
    )

    return {
        "Q": q, "H": h, "C": c, "N": n, "M": m,
        "buses": buses, "v_3500_7000": v_3500_7000, "w_7000": w_7000,
        "Y": y, "A": a, "Z": z, "G": g, "R": r, "P": p,
        "B": b_court, "L": l_transg, "E": e_exemp, "F": f_exemp,
    }


def _get_manual_shifts(session) -> list[dict[str, Any]]:
    from app.services.mobile_excel_report_builder import _manual_shifts
    try:
        return _manual_shifts(session)
    except Exception:
        return []


def _get_shift_for_record(record, shifts: list[dict[str, Any]]) -> dict[str, Any]:
    from app.services.mobile_excel_report_builder import _shift_for_record
    try:
        return _shift_for_record(record, shifts)
    except Exception:
        return shifts[0] if shifts else {}


def collect_juja_cluster_monthly_data(year: int, month: int) -> dict[str, Any]:
    _, num_days = calendar.monthrange(year, month)
    dates = [f"{year:04d}-{month:02d}-{day:02d}" for day in range(1, num_days + 1)]
    dates_set = set(dates)

    # Classify all sessions in store
    static_sessions_thika: dict[str, Any] = {}
    static_sessions_nairobi: dict[str, Any] = {}
    mobile_sessions_by_date: dict[str, list[Any]] = {d: [] for d in dates}
    all_mobile_sessions: list[Any] = []

    for s, mod_time in report_session_store.list_all_sessions():
        if not s or not s.report_date or s.report_date not in dates_set:
            continue

        st = (s.station or s.weighbridge_name or "").lower()
        b = (s.bound or "").lower()
        if "juja" not in st:
            continue

        has_mobile = "mobile" in st or "mobile" in b or s.sections.get("mobile_report", {}).get("status") == "ready"

        if not has_mobile:
            if "thika" in b:
                # Keep latest ready/modified
                cur = static_sessions_thika.get(s.report_date)
                if not cur or (s.sections.get("daily_summary", {}).get("status") == "ready" and cur[1] != "ready"):
                    static_sessions_thika[s.report_date] = (s, s.sections.get("daily_summary", {}).get("status"))
            elif "nairobi" in b:
                cur = static_sessions_nairobi.get(s.report_date)
                if not cur or (s.sections.get("daily_summary", {}).get("status") == "ready" and cur[1] != "ready"):
                    static_sessions_nairobi[s.report_date] = (s, s.sections.get("daily_summary", {}).get("status"))
        else:
            mobile_sessions_by_date[s.report_date].append((s, mod_time or 0))
            all_mobile_sessions.append(s)

    # Compile Static daily metrics
    static_thika_by_date = {}
    static_nairobi_by_date = {}
    for d in dates:
        t_sess = static_sessions_thika.get(d)[0] if d in static_sessions_thika else None
        n_sess = static_sessions_nairobi.get(d)[0] if d in static_sessions_nairobi else None
        static_thika_by_date[d] = _extract_static_session_metrics(t_sess)
        static_nairobi_by_date[d] = _extract_static_session_metrics(n_sess)

    # Compile Mobile daily compliance
    mobile_daily_by_date = {}
    for d in dates:
        candidates = mobile_sessions_by_date[d]
        c1 = [item for item in candidates if _mobile_report_slot(item[0].bound) == "mobile_1"]
        c2 = [item for item in candidates if _mobile_report_slot(item[0].bound) == "mobile_2"]
        c1.sort(key=lambda x: (x[0].sections.get("mobile_report", {}).get("status") == "ready", x[1]), reverse=True)
        c2.sort(key=lambda x: (x[0].sections.get("mobile_report", {}).get("status") == "ready", x[1]), reverse=True)

        s1 = c1[0][0] if c1 else None
        s2 = c2[0][0] if c2 else None

        st1 = _extract_mobile_session_stats(s1) if s1 else {
            "weighed": 0, "warned_charged": 0, "charged": 0, "warned": 0,
            "cases_cleared": 0, "transgressions": 0, "exempted": 0, "manually_weighed": 0,
        }
        st2 = _extract_mobile_session_stats(s2) if s2 else {
            "weighed": 0, "warned_charged": 0, "charged": 0, "warned": 0,
            "cases_cleared": 0, "transgressions": 0, "exempted": 0, "manually_weighed": 0,
        }

        mobile_daily_by_date[d] = {
            "weighed": st1["weighed"] + st2["weighed"],
            "warned_charged": st1["warned_charged"] + st2["warned_charged"],
            "charged": st1["charged"] + st2["charged"],
            "warned": st1["warned"] + st2["warned"],
            "cases_cleared": st1["cases_cleared"] + st2["cases_cleared"],
            "transgressions": st1["transgressions"] + st2["transgressions"],
            "exempted": st1["exempted"] + st2["exempted"],
            "manually_weighed": st1["manually_weighed"] + st2["manually_weighed"],
        }

    # Collect Deployment and Charged Records from mobile sessions
    deployment_records = []
    # Deduplicate sessions by (report_date, bound)
    seen_keys = set()
    ordered_sessions = []
    for d in dates:
        for sess, _ in sorted(mobile_sessions_by_date[d], key=lambda x: x[0].bound or ""):
            key = (sess.report_date, sess.bound)
            if key not in seen_keys:
                seen_keys.add(key)
                ordered_sessions.append(sess)

    for sess in ordered_sessions:
        df = None
        if hasattr(sess, "dataframes") and "mobile_report" in sess.dataframes:
            df = sess.dataframes["mobile_report"]
        elif "mobile_report" in sess.sections:
            sec = sess.sections["mobile_report"]
            recs = sec.get("records") or sec.get("preview")
            if recs and isinstance(recs, list):
                try:
                    df = pd.DataFrame(recs)
                except Exception:
                    pass

        if df is None or not isinstance(df, pd.DataFrame) or df.empty:
            continue

        route = _upper(
            sess.manual_inputs.get("route")
            or sess.manual_inputs.get("mobile_report", {}).get("route")
            or ""
        )
        shifts = _get_manual_shifts(sess)

        # Iterate records in order
        num_recs = len(df)
        for idx, (_, rec) in enumerate(df.iterrows()):
            rec_date_raw = rec.get("date_time")
            dt_obj = pd.to_datetime(rec_date_raw, errors="coerce", dayfirst=True)
            if pd.isna(dt_obj):
                try:
                    dt_obj = datetime.strptime(sess.report_date, "%Y-%m-%d")
                except Exception:
                    dt_obj = datetime(year, month, 1)

            gvw_diff = _safe_int(rec.get("gvw_difference_kg"), default=0)
            axle_diff = _safe_int(rec.get("excess_kg"), default=0)
            gvw_val: int | str = gvw_diff if gvw_diff > 0 else "-"
            axle_val: int | str = axle_diff if axle_diff > 0 else "-"

            shift = _get_shift_for_record(rec, shifts)
            danka = _upper(shift.get("danka_staff") or "")
            officers = _upper(shift.get("police_officers") or "")

            is_last_rec = (idx == num_recs - 1)
            mileage_start = _safe_int(shift.get("mileage_start"), default=None) if is_last_rec else None
            mileage_end = _safe_int(shift.get("mileage_end"), default=None) if is_last_rec else None
            kms = None
            if is_last_rec and mileage_start is not None and mileage_end is not None:
                kms = max(0, mileage_end - mileage_start)
            motor_veh = _upper(shift.get("mobile_vehicle") or "") if is_last_rec else ""

            deployment_records.append({
                "date": dt_obj,
                "registration": _upper(rec.get("registration")),
                "transporter": _upper(rec.get("transporter")),
                "config": _upper(rec.get("axle")),
                "excess_gvw": gvw_val,
                "excess_axle": axle_val,
                "excess_gvw_num": gvw_diff,
                "excess_axle_num": axle_diff,
                "origin": _upper(rec.get("origin")),
                "destination": _upper(rec.get("destination")),
                "cargo": _upper(rec.get("cargo")),
                "operator": danka,
                "officer": officers,
                "remarks": _upper(rec.get("remarks")),
                "route": route,
                "mileage_start": mileage_start,
                "mileage_end": mileage_end,
                "kms": kms,
                "motor_vehicle": motor_veh,
            })

    # Charged records
    charged_records = [
        r for r in deployment_records
        if "CHARGED" in r["remarks"]
    ]

    # Charged above 2 tonnes: GVW >= 2000 OR Axle >= 2000
    above_2t_charged_records = [
        r for r in charged_records
        if r["excess_gvw_num"] >= 2000 or r["excess_axle_num"] >= 2000
    ]

    return {
        "year": year,
        "month": month,
        "month_name": calendar.month_name[month],
        "num_days": num_days,
        "dates": dates,
        "static_thika": static_thika_by_date,
        "static_nairobi": static_nairobi_by_date,
        "mobile_daily": mobile_daily_by_date,
        "deployment_records": deployment_records,
        "charged_records": charged_records,
        "above_2t_charged_records": above_2t_charged_records,
    }


def _apply_border_range(
    ws,
    min_row: int,
    max_row: int,
    min_col: int,
    max_col: int,
    border: Border,
):
    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            ws.cell(row=r, column=c).border = border


def _apply_box_border(
    ws,
    min_row: int,
    max_row: int,
    min_col: int,
    max_col: int,
    outer_side=MEDIUM_SIDE,
    inner_side=THIN_SIDE,
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


def build_monthly_juja_cluster_excel_report(year: int, month: int) -> io.BytesIO:
    data = collect_juja_cluster_monthly_data(year, month)
    dates = data["dates"]
    num_days = data["num_days"]
    month_name = data["month_name"]
    thika = data["static_thika"]
    nairobi = data["static_nairobi"]
    mobile = data["mobile_daily"]
    deployment = data["deployment_records"]
    charged = data["charged_records"]
    above_2t = data["above_2t_charged_records"]

    wb = Workbook()

    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")
    align_center_wrap = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # =========================================================================
    # SHEET 1: Juja Cluster Invoicing Report
    # =========================================================================
    ws1 = wb.active
    ws1.title = "Juja Cluster Invoicing Report"
    ws1.views.sheetView[0].showGridLines = True
    ws1.page_setup.orientation = "landscape"
    ws1.page_setup.paperSize = 9  # A4
    ws1.page_setup.fitToWidth = 1
    ws1.page_setup.fitToHeight = 1
    ws1.sheet_properties.pageSetUpPr.fitToPage = False
    ws1.page_margins.left = 0.7
    ws1.page_margins.right = 0.7
    ws1.page_margins.top = 0.75
    ws1.page_margins.bottom = 0.75

    for col_letter, width in STATIC_COLUMN_WIDTHS.items():
        ws1.column_dimensions[col_letter].width = width

    # Row 3: Title
    ws1.row_dimensions[3].height = 54.0
    ws1.merge_cells("A3:X3")
    title_cell = ws1.cell(3, 1, f"Monthly JUJA Cluster Invoicing Report - {month_name.upper()} {year}")
    title_cell.font = Font(name=FONT_ARIAL, size=36, bold=True)
    title_cell.alignment = align_center
    title_cell.fill = FILL_LIGHT_GRAY
    for c in range(1, 25):
        ws1.cell(3, c).border = MEDIUM_BORDER
        ws1.cell(3, c).fill = FILL_LIGHT_GRAY

    # Row 4, 5: spacers
    ws1.row_dimensions[4].height = 15.75
    ws1.row_dimensions[5].height = 17.25

    # Row 6: Section 1 Header
    ws1.row_dimensions[6].height = 37.5
    ws1.merge_cells("A6:X6")
    sec_cell = ws1.cell(6, 1, "JUJA Weighbridge")
    sec_cell.font = Font(name=FONT_ARIAL, size=26, bold=True)
    sec_cell.alignment = align_center
    sec_cell.fill = FILL_LIGHT_GRAY
    for c in range(1, 25):
        ws1.cell(6, c).border = MEDIUM_BORDER
        ws1.cell(6, c).fill = FILL_LIGHT_GRAY

    # Row 7, 8: spacers
    ws1.row_dimensions[7].height = 15.0

    # Row 9 & 10: Static Headers
    ws1.row_dimensions[9].height = 114.0
    ws1.row_dimensions[10].height = 14.25

    static_headers = {
        1: "Date",
        2: "Bound",
        3: "No. of Vehicles Clearedby HSWIM   (Q)",
        4: "No. of Vehicles weighed HSWIM   (H)",
        5: "Called in  (C)",
        6: "No. of  Vehicles weighed Scale      (N)",
        7: "No. of Vehicles weighed manually (M)",
        8: "Total No. of Vehiclesweighed          (X =N+M)",
        9: "No. of Buses ≥3500kg (J)",
        10: "No of Vehicles ≥3500Kg ≤7000Kg (V)",
        11: "No. of Vehicles ˃7000 excluding buses (W)",
        12: "Traffic Census (K)",
        13: "Total Traffic (T=Q+X+K+E)",
        14: "Total No. of Vehicles  Overloaded  (Y)            A+Z+G+R",
        15: "No. of Vehicles warned   (A)",
        16: "No. of Vehicles charged in court (Z)",
        17: "No. of Vehicles given special release (G)",
        18: "No. of vehicles given warning to offload  or redistribute      (R)",
        19: "No of vehicles Impounded &  prohibited                (P)          Z+R",
        20: "No of vehicles  cleared in court (B)",
        21: "Transgressions (L)",
        22: "Exemption Permits",
        23: "Exemption Permits",
        24: "Total No. of Exemption Permits",
    }

    for col_idx, h_text in static_headers.items():
        cell = ws1.cell(9, col_idx, h_text)
        cell.font = Font(name=FONT_ARIAL, size=11, bold=True)
        cell.alignment = align_center_wrap
        cell.fill = FILL_LIGHT_GRAY
        cell.border = THIN_BORDER

    for col_idx in range(1, 25):
        cell10 = ws1.cell(10, col_idx)
        cell10.font = Font(name=FONT_ARIAL, size=11, bold=True)
        cell10.alignment = align_center_wrap
        cell10.fill = FILL_LIGHT_GRAY
        cell10.border = THIN_BORDER

    ws1.cell(10, 22, "Not Weighed E")
    ws1.cell(10, 23, "Weighed F")

    _apply_box_border(ws1, 9, 10, 1, 24, outer_side=MEDIUM_SIDE, inner_side=THIN_SIDE)

    # Static Data Rows: row 11 onwards (2 rows per day: Thika & Nairobi)
    current_row = 11
    static_start_row = current_row

    for d in dates:
        r_thika = current_row
        r_nairobi = current_row + 1
        dt_obj = datetime.strptime(d, "%Y-%m-%d")

        ws1.row_dimensions[r_thika].height = 15.0
        ws1.row_dimensions[r_nairobi].height = 15.0

        # Merge Date column A
        ws1.merge_cells(start_row=r_thika, start_column=1, end_row=r_nairobi, end_column=1)
        c_date = ws1.cell(r_thika, 1, dt_obj)
        c_date.font = Font(name=FONT_ARIAL, size=11, bold=True)
        c_date.alignment = align_center
        c_date.number_format = DATE_FORMAT
        c_date.border = THIN_BORDER
        ws1.cell(r_nairobi, 1).border = THIN_BORDER

        # Bounds
        b_thika = ws1.cell(r_thika, 2, "Thika")
        b_thika.font = Font(name=FONT_ARIAL, size=11, bold=True)
        b_thika.alignment = align_center
        b_thika.border = THIN_BORDER

        b_nairobi = ws1.cell(r_nairobi, 2, "Nairobi")
        b_nairobi.font = Font(name=FONT_ARIAL, size=11, bold=True)
        b_nairobi.alignment = align_center
        b_nairobi.border = THIN_BORDER

        # Metrics for Thika & Nairobi
        mt = thika[d]
        mn = nairobi[d]

        thika_col_vals = {
            3: mt["Q"], 4: mt["H"], 5: mt["C"], 6: mt["N"], 7: mt["M"],
            8: f"=G{r_thika}+F{r_thika}",
            9: mt["buses"], 10: mt["v_3500_7000"], 11: mt["w_7000"],
            12: f"=K{r_thika}+J{r_thika}+I{r_thika}",
            13: f"=C{r_thika}+H{r_thika}+L{r_thika}+V{r_thika}",
            14: mt["Y"], 15: mt["A"], 16: mt["Z"], 17: mt["G"], 18: mt["R"], 19: mt["P"],
            20: mt["B"], 21: mt["L"], 22: mt["E"], 23: mt["F"],
            24: f"=W{r_thika}+V{r_thika}",
        }

        nairobi_col_vals = {
            3: mn["Q"], 4: mn["H"], 5: mn["C"], 6: mn["N"], 7: mn["M"],
            8: f"=G{r_nairobi}+F{r_nairobi}",
            9: mn["buses"], 10: mn["v_3500_7000"], 11: mn["w_7000"],
            12: f"=K{r_nairobi}+J{r_nairobi}+I{r_nairobi}",
            13: f"=C{r_nairobi}+H{r_nairobi}+L{r_nairobi}+V{r_nairobi}",
            14: mn["Y"], 15: mn["A"], 16: mn["Z"], 17: mn["G"], 18: mn["R"], 19: mn["P"],
            20: mn["B"], 21: mn["L"], 22: mn["E"], 23: mn["F"],
            24: f"=W{r_nairobi}+V{r_nairobi}",
        }

        for c_idx in range(3, 25):
            c_t = ws1.cell(r_thika, c_idx, thika_col_vals[c_idx])
            c_t.font = Font(name=FONT_ARIAL, size=11, bold=False)
            c_t.alignment = align_center
            c_t.border = THIN_BORDER
            c_t.number_format = NUM_FORMAT
            if c_idx < 24:
                c_t.fill = FILL_YELLOW

            c_n = ws1.cell(r_nairobi, c_idx, nairobi_col_vals[c_idx])
            c_n.font = Font(name=FONT_ARIAL, size=11, bold=False)
            c_n.alignment = align_center
            c_n.border = THIN_BORDER
            c_n.number_format = NUM_FORMAT
            if c_idx < 24:
                c_n.fill = FILL_YELLOW

        current_row += 2

    static_end_row = current_row - 1

    # Static TOTAL Row
    r_stat_tot = current_row
    ws1.row_dimensions[r_stat_tot].height = 21.75
    ws1.cell(r_stat_tot, 1, "TOTAL").font = Font(name=FONT_ARIAL, size=11, bold=True)
    ws1.cell(r_stat_tot, 1).alignment = align_center
    ws1.cell(r_stat_tot, 1).fill = FILL_LIGHT_GRAY

    ws1.cell(r_stat_tot, 2, "").fill = FILL_LIGHT_GRAY

    # Numeric totals
    for c_idx in range(3, 24):
        col_letter = get_column_letter(c_idx)
        formula = f"=SUM({col_letter}{static_start_row}:{col_letter}{static_end_row})"
        cell = ws1.cell(r_stat_tot, c_idx, formula)
        cell.alignment = align_center
        cell.number_format = NUM_FORMAT

        if c_idx in (3, 4, 5):
            cell.font = Font(name=FONT_ARIAL, size=11, bold=True, color="FFFFFF")
            cell.fill = FILL_RED
        else:
            cell.font = Font(name=FONT_ARIAL, size=11, bold=True)
            cell.fill = FILL_GRAY_15

    # Column X: Total Exemption Permits formula
    cell_x = ws1.cell(r_stat_tot, 24, f"=W{r_stat_tot}+V{r_stat_tot}")
    cell_x.font = Font(name=FONT_ARIAL, size=11, bold=True)
    cell_x.alignment = align_center
    cell_x.number_format = NUM_FORMAT
    cell_x.fill = FILL_LIGHT_GRAY

    for c_idx in range(1, 25):
        ws1.cell(r_stat_tot, c_idx).border = TOTAL_STATIC_BORDER

    current_row += 1

    # Spacers after Static Total
    ws1.row_dimensions[current_row].height = 21.75
    ws1.row_dimensions[current_row + 1].height = 14.25
    current_row += 2

    # Row: MOBILE WEIGHBRIDGE Header
    r_mob_title = current_row
    ws1.row_dimensions[r_mob_title].height = 31.5
    ws1.merge_cells(start_row=r_mob_title, start_column=1, end_row=r_mob_title, end_column=24)
    m_title = ws1.cell(r_mob_title, 1, "MOBILE WEIGHBRIDGE")
    m_title.font = Font(name=FONT_ARIAL, size=26, bold=True)
    m_title.alignment = align_center
    m_title.fill = FILL_LIGHT_GRAY
    for c in range(1, 25):
        ws1.cell(r_mob_title, c).border = MEDIUM_BORDER
        ws1.cell(r_mob_title, c).fill = FILL_LIGHT_GRAY

    current_row += 1

    # Spacers
    ws1.row_dimensions[current_row].height = 15.0
    current_row += 2

    # Mobile Weighbridge Column Headers
    r_mob_hdr = current_row
    ws1.row_dimensions[r_mob_hdr].height = 58.5

    mob_headers = [
        "Date",
        "Total Weighed (X)",
        "Warned & Charged (Y)",
        "Charged (Z)",
        "Warned (A)",
        "Cases Cleared in Court (B)",
        "  Transgressions (L)",
        "Exempted permit (E)",
        "Manually Weighed (M)",
    ]

    for col_idx, h_text in enumerate(mob_headers, start=1):
        cell = ws1.cell(r_mob_hdr, col_idx, h_text)
        cell.font = Font(name=FONT_ARIAL, size=11, bold=True)
        cell.alignment = align_center_wrap
        cell.fill = FILL_LIGHT_GRAY
        cell.border = THIN_BORDER

    _apply_box_border(ws1, r_mob_hdr, r_mob_hdr, 1, 9, outer_side=MEDIUM_SIDE, inner_side=THIN_SIDE)
    current_row += 1

    # Mobile Data Rows
    r_mob_start = current_row
    for d in dates:
        ws1.row_dimensions[current_row].height = 14.25
        dt_obj = datetime.strptime(d, "%Y-%m-%d")
        st = mobile[d]

        c_dt = ws1.cell(current_row, 1, dt_obj)
        c_dt.font = Font(name=FONT_CALIBRI, size=11, bold=True, color=COLOR_MAGENTA)
        c_dt.alignment = align_center
        c_dt.fill = FILL_LIGHT_RED_4
        c_dt.number_format = DATE_FORMAT
        c_dt.border = THIN_BORDER

        vals = [
            st["weighed"],
            st["warned_charged"],
            st["charged"],
            st["warned"],
            st["cases_cleared"],
            st["transgressions"],
            st["exempted"],
            st["manually_weighed"],
        ]

        for col_idx, v in enumerate(vals, start=2):
            c_v = ws1.cell(current_row, col_idx, v)
            c_v.font = Font(name=FONT_CALIBRI, size=11, bold=False)
            c_v.alignment = align_center
            c_v.border = THIN_BORDER
            c_v.number_format = NUM_FORMAT

        current_row += 1

    r_mob_end = current_row - 1

    # Mobile TOTAL Row
    r_mob_tot = current_row
    ws1.row_dimensions[r_mob_tot].height = 28.5

    t_lbl = ws1.cell(r_mob_tot, 1, "Total")
    t_lbl.font = Font(name=FONT_ARIAL, size=18, bold=True)
    t_lbl.alignment = align_center
    t_lbl.fill = FILL_LIGHT_GRAY
    t_lbl.border = TOTAL_MOBILE_BORDER

    for col_idx in range(2, 10):
        col_letter = get_column_letter(col_idx)
        c_tot = ws1.cell(r_mob_tot, col_idx, f"=SUM({col_letter}{r_mob_start}:{col_letter}{r_mob_end})")
        c_tot.font = Font(name=FONT_ARIAL, size=11, bold=True)
        c_tot.alignment = align_center
        c_tot.fill = FILL_GRAY_25
        c_tot.number_format = NUM_FORMAT
        c_tot.border = TOTAL_MOBILE_BORDER

    current_row += 3  # Spacers

    # Row: Monthly JUJA Cluster Summary Title
    r_sum_title = current_row
    ws1.row_dimensions[r_sum_title].height = 57.0
    ws1.merge_cells(start_row=r_sum_title, start_column=1, end_row=r_sum_title, end_column=24)
    s_title = ws1.cell(r_sum_title, 1, "Monthly JUJA  Cluster Summary")
    s_title.font = Font(name=FONT_ARIAL, size=36, bold=True)
    s_title.alignment = align_center
    s_title.fill = FILL_LIGHT_GRAY
    for c in range(1, 25):
        ws1.cell(r_sum_title, c).border = MEDIUM_BORDER
        ws1.cell(r_sum_title, c).fill = FILL_LIGHT_GRAY

    current_row += 1
    ws1.row_dimensions[current_row].height = 16.5
    ws1.row_dimensions[current_row + 1].height = 14.25
    current_row += 2

    # Summary Header Row
    r_sum_hdr = current_row
    ws1.row_dimensions[r_sum_hdr].height = 108.0

    cluster_headers = {
        2: "No. of Vehicles weighed HSWIM   (Q)",
        3: "No. of  Vehicles weighed Scale      (N)",
        4: "Called in  (C)",
        5: "No. of  Vehicles weighed Scale      (N)",
        6: "No. of Vehicles weighed manually (M)",
        7: "Total No. of Vehiclesweighed (X)",
        8: "No. of Buses ≥3500kg",
        9: "No of Vehicles ≥3500Kg ≤7000Kg",
        10: "No. of Vehicles ˃7000 excluding buses",
        11: "Traffic Census (K)",
        12: "Total Traffic   (T)",
        13: "Total No. of Vehicles  Overloaded  (Y)A+Z+G+R",
        14: "No. of Vehicles warned   (A)",
        15: "No. of Vehicles charged in court (Z)",
        16: "No. of Vehicles given special release (G)",
        17: "No. of vehicles given warning to offload  or redistribute (R)",
        18: "No of vehicles Impounded &  prohibited (P) Z+R",
        19: "No of vehicles  cleared in court (B)",
        20: "Transgressions (L)",
        21: "Exemption Permits not weighed (E)",
        22: "Exemption Permits weighed (F)",
        23: "Total No. of Exemption Permits",
    }

    ws1.cell(r_sum_hdr, 1, "").border = THIN_BORDER
    for col_idx, h_text in cluster_headers.items():
        cell = ws1.cell(r_sum_hdr, col_idx, h_text)
        cell.font = Font(name=FONT_ARIAL, size=11, bold=True)
        cell.alignment = align_center_wrap
        cell.fill = FILL_LIGHT_GRAY
        cell.border = THIN_BORDER

    _apply_box_border(ws1, r_sum_hdr, r_sum_hdr, 1, 23, outer_side=MEDIUM_SIDE, inner_side=THIN_SIDE)
    current_row += 1

    # GRAND TOTAL Row
    r_grand_tot = current_row
    ws1.row_dimensions[r_grand_tot].height = 39.0

    gt_lbl = ws1.cell(r_grand_tot, 1, "GRAND TOTAL")
    gt_lbl.font = Font(name=FONT_ARIAL, size=11, bold=True)
    gt_lbl.alignment = align_center
    gt_lbl.border = MEDIUM_BORDER

    grand_total_formulas = {
        2: f"=C{r_stat_tot}",
        3: f"=D{r_stat_tot}",
        4: f"=E{r_stat_tot}",
        5: f"=B{r_mob_tot}+F{r_stat_tot}",
        6: f"=I{r_mob_tot}+G{r_stat_tot}",
        7: f"=B{r_mob_tot}+H{r_stat_tot}",
        8: f"=I{r_stat_tot}",
        9: f"=J{r_stat_tot}",
        10: f"=K{r_stat_tot}",
        11: f"=L{r_stat_tot}",
        12: f"=M{r_stat_tot}",
        13: f"=N{r_stat_tot}",
        14: f"=O{r_stat_tot}+C{r_mob_tot}",
        15: f"=P{r_stat_tot}+D{r_mob_tot}",
        16: f"=Q{r_stat_tot}",
        17: f"=R{r_stat_tot}",
        18: f"=S{r_stat_tot}",
        19: f"=T{r_stat_tot}+F{r_mob_tot}",
        20: f"=U{r_stat_tot}",
        21: f"=V{r_stat_tot}",
        22: f"=W{r_stat_tot}",
        23: f"=X{r_stat_tot}",
    }

    for col_idx in range(2, 24):
        formula = grand_total_formulas[col_idx]
        cell = ws1.cell(r_grand_tot, col_idx, formula)
        cell.font = Font(name=FONT_ARIAL, size=14, bold=True)
        cell.alignment = align_center
        cell.number_format = NUM_FORMAT
        cell.border = MEDIUM_BORDER

        if col_idx in (2, 3, 4):
            cell.font = Font(name=FONT_ARIAL, size=14, bold=True, color="FFFFFF")
            cell.fill = FILL_RED

    current_row += 1

    # Spacers
    ws1.row_dimensions[current_row].height = 15.0
    ws1.row_dimensions[current_row + 1].height = 14.25
    current_row += 2

    # Contract Subtitle Row
    r_contract = current_row
    ws1.row_dimensions[r_contract].height = 32.25
    ws1.merge_cells(start_row=r_contract, start_column=1, end_row=r_contract, end_column=24)
    contract_text = (
        f"JUJA CLUSTER MONTHLY INVOICING REPORT                                         "
        f"KENHA/WB/MTCE/4339/2026                                                                                                         "
        f"FOR {month_name.upper()} {year}"
    )
    c_cell = ws1.cell(r_contract, 1, contract_text)
    c_cell.font = Font(name=FONT_ARIAL, size=16, bold=True)
    c_cell.alignment = align_center
    c_cell.fill = FILL_LIGHT_GRAY
    for c in range(1, 25):
        ws1.cell(r_contract, c).border = MEDIUM_BORDER
        ws1.cell(r_contract, c).fill = FILL_LIGHT_GRAY

    current_row += 1

    # Spacers before Notes
    ws1.row_dimensions[current_row].height = 22.05
    ws1.row_dimensions[current_row + 1].height = 22.05
    current_row += 2

    # Notes Title Row
    r_notes_hdr = current_row
    ws1.row_dimensions[r_notes_hdr].height = 33.85
    ws1.merge_cells(start_row=r_notes_hdr, start_column=11, end_row=r_notes_hdr, end_column=16)
    n_title = ws1.cell(r_notes_hdr, 11, "Notes")
    n_title.font = Font(name=FONT_ARIAL, size=28, bold=True)
    n_title.alignment = align_center
    n_title.fill = FILL_YELLOW
    for c in range(11, 17):
        ws1.cell(r_notes_hdr, c).border = MEDIUM_BORDER
        ws1.cell(r_notes_hdr, c).fill = FILL_YELLOW

    current_row += 1

    # 19 Notes Rows
    note_heights = [22.05, 22.05, 17.35] + [14.7] * 16
    for idx, n_txt in enumerate(NOTES_TEXTS):
        ws1.row_dimensions[current_row].height = note_heights[idx] if idx < len(note_heights) else 15.0
        ws1.merge_cells(start_row=current_row, start_column=11, end_row=current_row, end_column=16)
        n_cell = ws1.cell(current_row, 11, n_txt)
        n_cell.font = Font(name=FONT_ARIAL, size=12, bold=False)
        n_cell.alignment = Alignment(horizontal="left", vertical="bottom")
        for c in range(11, 17):
            ws1.cell(current_row, c).border = MEDIUM_BORDER

        current_row += 1

    # =========================================================================
    # SHEET 2: Deployment
    # =========================================================================
    ws2 = wb.create_sheet(title="Deployment")
    ws2.views.sheetView[0].showGridLines = True
    ws2.page_setup.orientation = "landscape"
    ws2.page_setup.paperSize = 9
    ws2.page_setup.fitToWidth = 1
    ws2.page_setup.fitToHeight = 0
    ws2.sheet_properties.pageSetUpPr.fitToPage = True
    ws2.page_margins.left = 0.7
    ws2.page_margins.right = 0.7
    ws2.page_margins.top = 0.75
    ws2.page_margins.bottom = 0.75

    for col_letter, width in DEPLOYMENT_COLUMN_WIDTHS.items():
        ws2.column_dimensions[col_letter].width = width

    ws2.row_dimensions[3].height = 26.8
    ws2.merge_cells("A3:K3")
    dep_title = ws2.cell(3, 1, "JUJA  CLUSTER MONTHLY MOBILE DEPLOYMENT SUMMARY REPORT")
    dep_title.font = Font(name=FONT_ARIAL, size=22, bold=True)
    dep_title.alignment = Alignment(horizontal="center", vertical="bottom")
    _apply_box_border(ws2, 3, 3, 1, 11, outer_side=MEDIUM_SIDE, inner_side=None)

    ws2.row_dimensions[4].height = 12.75
    ws2.row_dimensions[5].height = 12.75

    # Headers Row 6
    ws2.row_dimensions[6].height = 28.2
    dep_h6 = {
        1: "DATE  WEIGHED",
        2: "VEHICLE REG",
        3: "TRANSPORTER",
        4: "CONFIG.",
        7: "ORIGIN",
        8: "DESTIN.",
        9: "CARGO",
        10: "COMPUTER OPERATER (DANKA STAFF)",
        11: "OFFICER",
        12: "REMARKS",
        13: "ROUTE",
        17: "MOTOR VEHICLE",
    }

    for col_idx, h_text in dep_h6.items():
        cell = ws2.cell(6, col_idx, h_text)
        cell.font = Font(name=FONT_ARIAL, size=12, bold=True)
        cell.alignment = align_center_wrap
        cell.fill = FILL_LIGHT_GRAY
        cell.border = THIN_BORDER

    ws2.merge_cells("E6:F6")
    cell_ew = ws2.cell(6, 5, "EXCESS   WEIGHT KGS")
    cell_ew.font = Font(name=FONT_ARIAL, size=12, bold=True)
    cell_ew.alignment = align_center
    cell_ew.fill = FILL_LIGHT_GRAY
    ws2.cell(6, 6).fill = FILL_LIGHT_GRAY
    ws2.cell(6, 5).border = THIN_BORDER
    ws2.cell(6, 6).border = THIN_BORDER

    ws2.merge_cells("N6:P6")
    cell_mil = ws2.cell(6, 14, "MILEAGE")
    cell_mil.font = Font(name=FONT_ARIAL, size=12, bold=True)
    cell_mil.alignment = align_center
    cell_mil.fill = FILL_LIGHT_GRAY
    ws2.cell(6, 15).fill = FILL_LIGHT_GRAY
    ws2.cell(6, 16).fill = FILL_LIGHT_GRAY
    for c in range(14, 17):
        ws2.cell(6, c).border = THIN_BORDER

    # Headers Row 7 (Subheaders)
    ws2.row_dimensions[7].height = 12.75
    sub_headers = {
        5: "GVW",
        6: "AXLE",
        14: "MILEAGE START",
        15: "MILEAGE END",
        16: "KMS",
    }
    for col_idx in range(1, 18):
        c = ws2.cell(7, col_idx)
        c.fill = FILL_LIGHT_GRAY
        c.border = THIN_BORDER
        if col_idx in sub_headers:
            c.value = sub_headers[col_idx]
            c.font = Font(name=FONT_ARIAL, size=10, bold=True)
            c.alignment = align_center_wrap

    # Deployment Records (Row 8 onwards)
    dep_row = 8
    for rec in deployment:
        ws2.row_dimensions[dep_row].height = 13.8

        # Date
        c_d = ws2.cell(dep_row, 1, rec["date"])
        c_d.font = Font(name=FONT_CALIBRI, size=11, bold=False)
        c_d.alignment = align_left
        c_d.number_format = DATE_FORMAT
        c_d.border = THIN_BORDER

        # Text fields
        row_map = {
            2: rec["registration"],
            3: rec["transporter"],
            4: rec["config"],
            5: rec["excess_gvw"],
            6: rec["excess_axle"],
            7: rec["origin"],
            8: rec["destination"],
            9: rec["cargo"],
            10: rec["operator"],
            11: rec["officer"],
            12: rec["remarks"],
            13: rec["route"],
            14: rec["mileage_start"],
            15: rec["mileage_end"],
            16: rec["kms"],
            17: rec["motor_vehicle"],
        }

        for col_idx in range(2, 18):
            val = row_map[col_idx]
            cell = ws2.cell(dep_row, col_idx, val)
            cell.font = Font(name=FONT_CALIBRI, size=11, bold=False)
            cell.border = THIN_BORDER

            if col_idx in (5, 6):
                cell.alignment = align_left
                if isinstance(val, (int, float)):
                    cell.number_format = NUM_FORMAT
            elif col_idx in (14, 15, 16):
                cell.alignment = align_center
                if isinstance(val, (int, float)):
                    cell.number_format = NUM_FORMAT
            elif col_idx == 17:
                cell.alignment = align_center
            else:
                cell.alignment = align_left

        dep_row += 1

    # =========================================================================
    # SHEET 3: Charged
    # =========================================================================
    ws3 = wb.create_sheet(title="Charged")
    ws3.views.sheetView[0].showGridLines = True
    ws3.page_setup.orientation = "landscape"
    ws3.page_setup.paperSize = 9
    ws3.page_setup.fitToWidth = 1
    ws3.page_setup.fitToHeight = 1
    ws3.sheet_properties.pageSetUpPr.fitToPage = False
    ws3.page_margins.left = 0.7
    ws3.page_margins.right = 0.7
    ws3.page_margins.top = 0.75
    ws3.page_margins.bottom = 0.75

    for col_letter, width in CHARGED_COLUMN_WIDTHS.items():
        ws3.column_dimensions[col_letter].width = width

    # Table 1: All Charged Vehicles
    ws3.row_dimensions[3].height = 31.5
    ws3.merge_cells("A3:K3")
    ch_title1 = ws3.cell(3, 1, "JUJA CLUSTER MONTHLY MOBILE DEPLOYMENT SUMMARY REPORT FOR VEHICLES CHARGED ")
    ch_title1.font = Font(name=FONT_CALIBRI, size=26, bold=True)
    ch_title1.alignment = Alignment(horizontal="center", vertical="bottom")
    _apply_box_border(ws3, 3, 3, 1, 11, outer_side=MEDIUM_SIDE, inner_side=None)

    ws3.row_dimensions[4].height = 15.0
    ws3.row_dimensions[5].height = 15.0

    # Headers Row 6
    ws3.row_dimensions[6].height = 48.0
    ch_headers = {
        1: "DATE  WEIGHED",
        2: "VEHICLE REG",
        3: "TRANSPORTER",
        4: "CONFIG.",
        7: "ORIGIN",
        8: "DESTIN.",
        9: "CARGO",
        10: "COMPUTER OPERATER (DANKA STAFF)",
        11: "OFFICER",
        12: "REMARKS",
        13: "ROUTE",
    }
    for col_idx, h_text in ch_headers.items():
        cell = ws3.cell(6, col_idx, h_text)
        cell.font = Font(name=FONT_ARIAL, size=13, bold=True)
        cell.alignment = align_center_wrap
        cell.fill = FILL_LIGHT_GRAY
        cell.border = THIN_BORDER

    ws3.merge_cells("E6:F6")
    c_ew1 = ws3.cell(6, 5, "EXCESS   WEIGHT KGS")
    c_ew1.font = Font(name=FONT_ARIAL, size=13, bold=True)
    c_ew1.alignment = align_center_wrap
    c_ew1.fill = FILL_LIGHT_GRAY
    ws3.cell(6, 6).fill = FILL_LIGHT_GRAY
    ws3.cell(6, 5).border = THIN_BORDER
    ws3.cell(6, 6).border = THIN_BORDER

    ch_row = 7
    for rec in charged:
        ws3.row_dimensions[ch_row].height = 15.0

        c_d = ws3.cell(ch_row, 1, rec["date"])
        c_d.font = Font(name=FONT_ARIAL, size=11, bold=False)
        c_d.alignment = align_center
        c_d.number_format = DATE_FORMAT
        c_d.border = THIN_BORDER

        ch_map = {
            2: rec["registration"],
            3: rec["transporter"],
            4: rec["config"],
            5: rec["excess_gvw"],
            6: rec["excess_axle"],
            7: rec["origin"],
            8: rec["destination"],
            9: rec["cargo"],
            10: rec["operator"],
            11: rec["officer"],
            12: rec["remarks"],
            13: rec["route"],
        }

        for col_idx in range(2, 14):
            val = ch_map[col_idx]
            cell = ws3.cell(ch_row, col_idx, val)
            cell.font = Font(name=FONT_ARIAL, size=11, bold=False)
            cell.border = THIN_BORDER

            if col_idx in (5, 6):
                cell.alignment = align_center
                if isinstance(val, (int, float)):
                    cell.number_format = NUM_FORMAT
            elif col_idx in (2, 3, 12):
                cell.alignment = align_left
            else:
                cell.alignment = Alignment(horizontal="general", vertical="center")

        ch_row += 1

    # Table 2: Charged Above 2 Tonnes
    ch_row += 3  # Spacers
    ws3.row_dimensions[ch_row].height = 31.5
    ws3.merge_cells(start_row=ch_row, start_column=1, end_row=ch_row, end_column=11)
    ch_title2 = ws3.cell(ch_row, 1, "JUJA CLUSTER MONTHLY MOBILE REPORT FOR VEHICLES CHARGED ABOVE 2 TONNES")
    ch_title2.font = Font(name=FONT_CALIBRI, size=26, bold=True)
    ch_title2.alignment = Alignment(horizontal="center", vertical="bottom")
    _apply_box_border(ws3, ch_row, ch_row, 1, 11, outer_side=MEDIUM_SIDE, inner_side=None)

    ch_row += 1
    ws3.row_dimensions[ch_row].height = 15.0
    ws3.row_dimensions[ch_row + 1].height = 15.0
    ch_row += 2

    # Headers Table 2
    r_hdr2_1 = ch_row
    ws3.row_dimensions[r_hdr2_1].height = 48.0
    for col_idx, h_text in ch_headers.items():
        cell = ws3.cell(r_hdr2_1, col_idx, h_text)
        cell.font = Font(name=FONT_ARIAL, size=11, bold=True)
        cell.alignment = align_center_wrap
        cell.fill = FILL_LIGHT_GRAY
        cell.border = THIN_BORDER

    ws3.merge_cells(start_row=r_hdr2_1, start_column=5, end_row=r_hdr2_1, end_column=6)
    c_ew2 = ws3.cell(r_hdr2_1, 5, "EXCESS   WEIGHT KGS")
    c_ew2.font = Font(name=FONT_ARIAL, size=11, bold=True)
    c_ew2.alignment = align_center_wrap
    c_ew2.fill = FILL_LIGHT_GRAY
    ws3.cell(r_hdr2_1, 6).fill = FILL_LIGHT_GRAY
    ws3.cell(r_hdr2_1, 5).border = THIN_BORDER
    ws3.cell(r_hdr2_1, 6).border = THIN_BORDER

    ch_row += 1
    r_hdr2_2 = ch_row
    ws3.row_dimensions[r_hdr2_2].height = 15.0
    for col_idx in range(1, 14):
        c = ws3.cell(r_hdr2_2, col_idx)
        c.fill = FILL_LIGHT_GRAY
        c.border = THIN_BORDER
        if col_idx == 5:
            c.value = "GVW"
            c.font = Font(name=FONT_ARIAL, size=11, bold=True)
            c.alignment = align_center
        elif col_idx == 6:
            c.value = "AXLE"
            c.font = Font(name=FONT_ARIAL, size=11, bold=True)
            c.alignment = align_center

    ch_row += 1

    # Table 2 Records (above 2 tonnes)
    for rec in above_2t:
        ws3.row_dimensions[ch_row].height = 15.0

        c_d = ws3.cell(ch_row, 1, rec["date"])
        c_d.font = Font(name=FONT_ARIAL, size=11, bold=False)
        c_d.alignment = align_center
        c_d.number_format = DATE_FORMAT
        c_d.border = THIN_BORDER

        ch_map = {
            2: rec["registration"],
            3: rec["transporter"],
            4: rec["config"],
            5: rec["excess_gvw"],
            6: rec["excess_axle"],
            7: rec["origin"],
            8: rec["destination"],
            9: rec["cargo"],
            10: rec["operator"],
            11: rec["officer"],
            12: rec["remarks"],
            13: rec["route"],
        }

        for col_idx in range(2, 14):
            val = ch_map[col_idx]
            cell = ws3.cell(ch_row, col_idx, val)
            cell.font = Font(name=FONT_ARIAL, size=11, bold=False)
            cell.border = THIN_BORDER

            if col_idx in (5, 6):
                cell.alignment = align_center
                if isinstance(val, (int, float)):
                    cell.number_format = NUM_FORMAT
            elif col_idx in (2, 3, 12):
                cell.alignment = align_left
            else:
                cell.alignment = Alignment(horizontal="general", vertical="center")

        ch_row += 1

    # Force recalculation flags for Excel viewers
    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
    except AttributeError:
        pass

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer
