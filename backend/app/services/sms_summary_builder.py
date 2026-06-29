from datetime import datetime
from typing import Any
import pandas as pd

from app.services.report_session_store import ReportSession
from app.services.daily_summary_processor import (
    _transgressions_count,
    _manual_count,
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

def build_static_sms_summary(session: ReportSession) -> str:
    station_name = str(session.weighbridge_name or session.station or "JUJA").strip().upper()
    if "WEIGHBRIDGE" in station_name:
        station_name = station_name.replace("WEIGHBRIDGE", "").strip()
    
    bound_name = str(session.bound or "").strip().upper()
    if "BOUND" not in bound_name and bound_name:
        bound_name = f"{bound_name} BOUND"
    
    date_formatted = format_date(session.report_date)
    
    # Extract metrics
    C = get_session_column_total(session, "daily_hour", "C")
    Q = get_session_column_total(session, "daily_hour", "Q")
    H = Q + C # Or H column total
    M = get_session_column_total(session, "daily_hour", "M")
    X = get_session_column_total(session, "daily_hour", "X")
    
    # Traffic Census
    traffic_census = session.manual_inputs.get("traffic_census", {})
    K = 0
    if traffic_census:
        K = int(traffic_census.get("total_traffic_census", 0))
    
    E = get_session_column_total(session, "daily_hour", "E")
    T = Q + X + K + E
    
    Y = get_session_column_total(session, "daily_hour", "Y")
    A = get_session_column_total(session, "daily_hour", "A")
    Z = get_session_column_total(session, "daily_hour", "Z")
    
    B = _manual_count(session.manual_inputs, "cases_cleared_in_court", "cases_cleared_court")
    G = get_session_column_total(session, "daily_hour", "G")
    R = get_session_column_total(session, "daily_hour", "R")
    P = Z + R
    L = _transgressions_count(session.manual_inputs)
    
    # Exemption P weighed (F)
    F = 0
    if (
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
        f"Transgression(L)={format_num(L)}\n"
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
    mobile_inputs = session.manual_inputs.get("mobile_report", {})
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
        f"Above 2 Tonnes={format_num(above_2t)}\n"
        f"Charged on dimensions={format_num(dimensions)}\n"
        f"Exemption permits(E)={format_num(E)}\n"
        f"Redistributed(R)={format_num(R)}\n"
        f"Impounded & prohibited(P)={format_num(P)}\n"
        f"Trangression (L)={format_num(L)}\n"
        f"Kilometers covered={format_num(kms_val)}KMS\n"
        f"Vehicle used:- {vehicle}\n\n"
        f"By:-{prepared_by}"
    )
    return template
