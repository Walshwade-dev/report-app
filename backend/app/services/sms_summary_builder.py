from datetime import datetime
from typing import Any
import pandas as pd

from app.services.report_session_store import ReportSession
from app.services.daily_summary_processor import (
    DailySummaryMissingSourceError,
    _transgressions_count,
    _manual_count,
    build_daily_summary_from_session,
)
from app.services.overloaded_summary import count_valid_permit_vehicles
from app.services.mobile_report_processor import summarize_mobile_report

def format_date(date_str: str) -> str:
    try:
        return datetime.strptime(date_str, "%Y-%m-%d").strftime("%d.%m.%Y")
    except Exception:
        return date_str

def format_num(val: Any) -> str:
    try:
        if val is None or pd.isna(val):
            return "0"
        num = int(float(str(val).replace(",", "")))
        return f"{num:,}"
    except Exception:
        return "0"

def get_session_column_total(session: ReportSession, section_name: str, column: str) -> int:
    if (
        section_name not in session.dataframes
        or session.sections.get(section_name, {}).get("status") != "ready"
    ):
        return 0
    df = session.dataframes[section_name]
    if "DATE" not in df.columns or column not in df.columns:
        return 0
    totals_mask = df["DATE"].astype(str).str.strip().str.lower().eq("totals")
    if not totals_mask.any():
        return 0
    try:
        return int(df.loc[totals_mask].iloc[-1].get(column, 0))
    except Exception:
        return 0

def get_static_summary_values(session: ReportSession) -> dict[str, int]:
    values = session.sections.get("daily_summary", {}).get("values")
    if isinstance(values, dict):
        return values

    try:
        return build_daily_summary_from_session(session)
    except DailySummaryMissingSourceError:
        return {}

def int_value(values: dict[str, Any], key: str, default: int = 0) -> int:
    try:
        return int(values.get(key, default) or default)
    except Exception:
        return default

def get_mobile_inputs(session: ReportSession) -> dict[str, Any]:
    inputs = session.manual_inputs.get("mobile_report")
    if isinstance(inputs, dict):
        return inputs

    extra = session.manual_inputs.get("extra")
    if isinstance(extra, dict) and isinstance(extra.get("mobile_report"), dict):
        return extra["mobile_report"]

    return {}

def build_static_sms_summary(session: ReportSession) -> str:
    station_name = str(session.weighbridge_name or session.station or "JUJA").strip().upper()
    if "WEIGHBRIDGE" in station_name:
        station_name = station_name.replace("WEIGHBRIDGE", "").strip()
    
    bound_name = str(session.bound or "").strip().upper()
    if "BOUND" not in bound_name and bound_name:
        bound_name = f"{bound_name} BOUND"
    
    date_formatted = format_date(session.report_date)
    
    summary_values = get_static_summary_values(session)

    C = get_session_column_total(session, "daily_hour", "C")
    H = get_session_column_total(session, "daily_hour", "H")
    Q = int_value(summary_values, "weighed_by_hswim_q", get_session_column_total(session, "daily_hour", "Q"))
    if H == 0:
        H = Q + C
    M = int_value(summary_values, "manually_weighed_m", get_session_column_total(session, "daily_hour", "M"))
    X = int_value(summary_values, "total_weighed_x", get_session_column_total(session, "daily_hour", "X"))
    
    traffic_census = session.manual_inputs.get("traffic_census", {})
    K = 0
    if traffic_census:
        K = int(traffic_census.get("total_traffic_census", 0))
    
    E = int_value(summary_values, "exemption_permits_not_weighed_e", get_session_column_total(session, "daily_hour", "E"))
    T = int_value(summary_values, "total_traffic_t", Q + X + K + E)
    
    Y = int_value(summary_values, "total_overload_y", get_session_column_total(session, "daily_hour", "Y"))
    A = int_value(summary_values, "warned_a", get_session_column_total(session, "daily_hour", "A"))
    Z = int_value(summary_values, "charged_prohibited_z", get_session_column_total(session, "daily_hour", "Z"))
    
    B = int_value(summary_values, "cases_cleared_in_court_b", _manual_count(session.manual_inputs, "cases_cleared_in_court", "cases_cleared_court"))
    G = int_value(summary_values, "special_release_g", get_session_column_total(session, "daily_hour", "G"))
    R = int_value(summary_values, "vehicles_charged_but_redistributed_r", get_session_column_total(session, "daily_hour", "R"))
    P = int_value(summary_values, "impounded_prohibited_p", Z + R)
    L = int_value(summary_values, "transgressions_l", _transgressions_count(session.manual_inputs))
    
    F = int_value(summary_values, "exemption_permits_weighed_f", 0)
    if F == 0 and (
        "overloaded" in session.dataframes
        and session.sections.get("overloaded", {}).get("status") == "ready"
    ):
        F = count_valid_permit_vehicles(session.dataframes["overloaded"])
        
    prepared_by = str(session.prepared_by or "").strip().upper()
    if not prepared_by:
        prepared_by = "ANASTASHA KENDA."
    else:
        if not prepared_by.endswith("."):
            prepared_by = f"{prepared_by}."

    template = (
        "DAILY REPORT\n"
        f"{station_name} {bound_name} WB\n"
        f"Date: {date_formatted}\n\n"
        f"HSWIM TOTAL(H)={format_num(H)}\n"
        f"Called in(C)={format_num(C)}\n"
        f"Weighed Hswim(Q)={format_num(Q)}\n"
        f"Manually Weighed(M)={format_num(M)}\n"
        f"T.Weighed(X)={format_num(X)}\n"
        f"T.traffic census(K)={format_num(K)}\n"
        f"T.traffic (T)={format_num(T)}\n"
        f"T.overloaded(Y)={format_num(Y)}\n"
        f"Warned(A)={format_num(A)}\n"
        f"Prohibited & Charged(Z)={format_num(Z)}\n"
        f"Cleared in Court(B)={format_num(B)}\n"
        f"Special Release(G)={format_num(G)}\n"
        f"Redistributed(R)={format_num(R)}\n"
        f"Impounded & prohibited(P)={format_num(P)}\n"
        f"Transgression(L)= {format_num(L)}\n"
        f"Exemption P not weighed(E)={format_num(E)}\n"
        f"Exemption P weighed (F)={format_num(F)}\n\n"
        f"By: {prepared_by}"
    )
    return template

def build_mobile_sms_summary(session: ReportSession) -> str:
    station_name = str(session.station or "JUJA").strip().upper()
    if "juja" in station_name.lower():
        station_name = "JUJA W/B"
    else:
        if "WEIGHBRIDGE" in station_name:
            station_name = station_name.replace("WEIGHBRIDGE", "").strip()
            
    bound_name = str(session.bound or "Mobile 1").strip()
    team_name = "TEAM ONE"
    if "2" in bound_name or "two" in bound_name.lower():
        team_name = "TEAM TWO"
        
    date_formatted = format_date(session.report_date)
    
    # Route, vehicle, kms from manual inputs
    mobile_inputs = get_mobile_inputs(session)
    route = str(mobile_inputs.get("route") or session.manual_inputs.get("route") or "—").strip().upper()
    
    vehicle = str(
        mobile_inputs.get("mobile_vehicle") or 
        mobile_inputs.get("vehicle_used") or 
        session.manual_inputs.get("mobile_vehicle") or 
        session.manual_inputs.get("vehicle_used") or 
        "—"
    ).strip().upper()
    
    # Calculate kms
    start = mobile_inputs.get("mileage_start") or session.manual_inputs.get("mileage_start")
    end = mobile_inputs.get("mileage_end") or session.manual_inputs.get("mileage_end")
    kms_val = 0
    try:
        if start is not None and end is not None:
            kms_val = int(float(str(end).replace(",", ""))) - int(float(str(start).replace(",", "")))
    except Exception:
        pass
    
    if kms_val <= 0:
        # Fallback to direct field
        try:
            kms_val = int(float(str(mobile_inputs.get("kilometers_covered") or session.manual_inputs.get("kilometers_covered") or 0).replace(",", "")))
        except Exception:
            pass
            
    # Extract mobile_report summary
    summary = {}
    above_2t = 0
    if (
        "mobile_report" in session.dataframes
        and session.sections.get("mobile_report", {}).get("status") == "ready"
    ):
        df = session.dataframes["mobile_report"]
        summary = summarize_mobile_report(df)
        try:
            above_2t = int((df["gvw_difference_kg"] > 2000).sum())
        except Exception:
            pass
            
    S = summary.get("total_records", 0)
    X = summary.get("total_trucks_weighed", 0)
    T = summary.get("total_records", 0)
    Y = summary.get("overloaded_records", 0)
    A = summary.get("warned_trucks", 0)
    
    charged = summary.get("charged_trucks", 0)
    legal = X - (A + charged)
    if legal < 0:
        legal = 0
        
    Z = charged
    B = _manual_count(session.manual_inputs, "cases_cleared_in_court", "cases_cleared_court")
    if not B:
        try:
            B = int(mobile_inputs.get("cases_cleared_in_court") or 0)
        except Exception:
            pass
            
    dimensions = summary.get("charged_dimensions_trucks", 0)
    
    E = _manual_count(session.manual_inputs, "exempted_permit", "exempted_permits")
    if not E:
        try:
            E = int(mobile_inputs.get("exempted_permit") or 0)
        except Exception:
            pass
            
    R = _manual_count(session.manual_inputs, "redistributed")
    if not R:
        try:
            R = int(mobile_inputs.get("redistributed") or 0)
        except Exception:
            pass
            
    P = _manual_count(session.manual_inputs, "impounded_prohibited", "impounded")
    if not P:
        try:
            P = int(mobile_inputs.get("impounded") or 0)
        except Exception:
            pass
            
    L = _transgressions_count(session.manual_inputs)
    if not L:
        try:
            L = int(mobile_inputs.get("transgressions_count") or 0)
        except Exception:
            pass
            
    prepared_by = str(session.prepared_by or "").strip().upper()
    if not prepared_by:
        prepared_by = "ANASTSHA KENDA."
    else:
        if not prepared_by.endswith("."):
            prepared_by = f"{prepared_by}."

    template = (
        "DAILY REPORT\n"
        f"{station_name} DAILY MOBILE REPORT_{team_name}\n"
        f"Date:{date_formatted}\n"
        f"Route:-{route}\n"
        f"scale (S)={format_num(S)}\n"
        f"Total Weighed(X)={format_num(X)}\n"
        f"Total traffic (T)={format_num(T)}\n"
        f"Total overloaded (Y)={format_num(Y)}\n"
        f"Warned (A)={format_num(A)}\n"
        f"Legal={format_num(legal)}\n"
        f"Charged & Prohibited(Z)={format_num(Z)}\n"
        f"Cleared in court (B)={format_num(B)}\n"
        f"Above 2 Tonnes ={format_num(above_2t)}\n"
        f"Charged on dimensions ={format_num(dimensions)}\n"
        f"Exemption permits(E)={format_num(E)}\n"
        f"Redistributed(R)={format_num(R)}\n"
        f"Impounded & prohibited(P)={format_num(P)}\n"
        f"Trangression (L)={format_num(L)}\n"
        f"Kilometers covered={format_num(kms_val)}KMS\n"
        f"Vehicle used:- {vehicle}\n\n"
        f"By:-{prepared_by}"
    )
    return template
