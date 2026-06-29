import logging
from datetime import datetime

from fastapi import APIRouter, File, Form, HTTPException, UploadFile
from fastapi.responses import Response
from pydantic import BaseModel

from app.services.cleaner_core import clean_with_template
from app.services.daily_hour_processor import (
    REQUIRED_COLUMNS,
    add_daily_totals_row,
    build_daily_hour_metrics,
    distribute_wideloads,
)
from app.services.excel_report_builder import build_excel_report
from app.services.final_report_builder import build_final_report
from app.services.mobile_excel_report_builder import build_mobile_excel_report
from app.services.mobile_report_processor import (
    mobile_report_response,
    normalize_mobile_report,
    summarize_mobile_report,
)
from app.services.mobile_word_report_builder import build_mobile_word_report
from app.services.overloaded_summary import count_valid_permit_vehicles
from app.services.preview_renderer import get_cached_section_preview
from app.services.report_session_metrics import get_wideload_count_from_session
from app.services.report_upload_service import read_upload_dataframe
from app.services.report_session_store import ReportSession, report_session_store
from app.templates import impounded_prohibited, vehicle_inspection


router = APIRouter()
logger = logging.getLogger(__name__)


class ReportSessionCreate(BaseModel):
    report_date: str
    station: str | None = None
    bound: str
    weighbridge_name: str | None = None
    prepared_by: str | None = None
    confirmed_by: str | None = None


class ManualInputsUpdate(BaseModel):
    prepared_by: str | None = None
    confirmed_by: str | None = None
    weighbridge_name: str | None = None
    traffic_census: dict | None = None
    transgressions: dict | list[dict] | None = None
    extra: dict | None = None


class ReportSessionMetadataUpdate(BaseModel):
    station: str | None = None
    bound: str | None = None
    weighbridge_name: str | None = None


def daily_display_date(report_date: str) -> str:
    try:
        return datetime.strptime(report_date, "%Y-%m-%d").strftime("%d/%m/%Y")
    except ValueError:
        return report_date


def format_filename_date(date_str: str) -> str:
    parts = date_str.split("-")
    if len(parts) == 3:
        year = parts[0][-2:]
        month = parts[1]
        day = parts[2]
        return f"{day}.{month}.{year}"
    return date_str


def get_report_filename(session: ReportSession, ext: str) -> str:
    station_name = (session.station or "STATION").upper()
    if "WEIGHBRIDGE" not in station_name:
        station_name = f"{station_name} WEIGHBRIDGE"
    bound_name = (session.bound or "BOUND").upper()
    if "BOUND" not in bound_name:
        bound_name = f"{bound_name} BOUND"
    date_part = format_filename_date(session.report_date)
    return f"{station_name} {bound_name} DAILY REPORT {date_part}.{ext}"


def get_mobile_report_filename(session: ReportSession, ext: str) -> str:
    station_name = (session.station or "STATION").upper()
    if "WEIGHBRIDGE" not in station_name:
        station_name = f"{station_name} WEIGHBRIDGE"
    bound_name = (session.bound or "BOUND").upper()
    date_part = format_filename_date(session.report_date)
    return f"{station_name} MOBILE REPORT {bound_name} {date_part}.{ext}"


def serialize_session(session: ReportSession) -> dict:
    excel_report_ready = (
        "daily_hour" in session.dataframes
        and session.sections.get("daily_hour", {}).get("status") == "ready"
    )
    mobile_excel_report_ready = (
        "mobile_report" in session.dataframes
        and session.sections.get("mobile_report", {}).get("status") == "ready"
    )
    mobile_word_report_ready = mobile_excel_report_ready
    final_report = {
        "status": session.final_report_status,
        "download_url": None,
        "error": session.final_report_error,
    }

    if session.final_report_status == "ready":
        final_report["download_url"] = (
            f"/api/report-sessions/{session.report_id}/download-final-report"
        )

    return {
        "report_id": session.report_id,
        "metadata": {
            "report_date": session.report_date,
            "station": session.station,
            "bound": session.bound,
            "weighbridge_name": session.weighbridge_name,
            "prepared_by": session.prepared_by,
            "confirmed_by": session.confirmed_by,
        },
        "manual_inputs": session.manual_inputs,
        "sections": session.sections,
        "final_report": final_report,
        "excel_report": {
            "status": "ready" if excel_report_ready else "awaiting_data",
            "download_url": (
                f"/api/report-sessions/{session.report_id}/download-excel-report"
                if excel_report_ready
                else None
            ),
        },
        "mobile_excel_report": {
            "status": "ready" if mobile_excel_report_ready else "awaiting_data",
            "download_url": (
                f"/api/report-sessions/{session.report_id}/download-mobile-excel-report"
                if mobile_excel_report_ready
                else None
            ),
        },
        "mobile_word_report": {
            "status": "ready" if mobile_word_report_ready else "awaiting_data",
            "download_url": (
                f"/api/report-sessions/{session.report_id}/download-mobile-word-report"
                if mobile_word_report_ready
                else None
            ),
        },
    }


def require_session(report_id: str) -> ReportSession:
    try:
        return report_session_store.require(report_id)
    except KeyError as exc:
        raise HTTPException(status_code=404, detail="Report session not found") from exc


def update_daily_hour_wideload_count(daily_df, report_date: str, wideload_count: int):
    if set(REQUIRED_COLUMNS).issubset(set(daily_df.columns)):
        return add_daily_totals_row(
            build_daily_hour_metrics(
                daily_df,
                report_date=daily_display_date(report_date),
                wideload_count=wideload_count,
            )
        )

    if "E" not in daily_df.columns or "DATE" not in daily_df.columns:
        raise ValueError("Ready daily_hour data cannot be updated with wideload count.")

    updated_df = daily_df.copy()
    totals_mask = (
        updated_df["DATE"]
        .astype(str)
        .str.strip()
        .str.lower()
        .eq("totals")
    )

    if totals_mask.any():
        updated_df = updated_df.loc[~totals_mask].copy()

    updated_df["E"] = distribute_wideloads(wideload_count)
    return add_daily_totals_row(updated_df)


def daily_hour_total_column(session: ReportSession, column: str) -> int | None:
    if (
        "daily_hour" not in session.dataframes
        or session.sections.get("daily_hour", {}).get("status") != "ready"
    ):
        return None

    daily_df = session.dataframes["daily_hour"]

    if "DATE" not in daily_df.columns or column not in daily_df.columns:
        return None

    totals_mask = (
        daily_df["DATE"]
        .astype(str)
        .str.strip()
        .str.lower()
        .eq("totals")
    )

    if not totals_mask.any():
        return None

    return int(daily_df.loc[totals_mask].iloc[-1].get(column, 0))


def summary_card(title: str, value: int | None, source: str) -> dict:
    is_ready = value is not None

    return {
        "title": title,
        "value": value,
        "display_value": f"{value:,}" if is_ready else "—",
        "status": "ready" if is_ready else "awaiting_data",
        "subtitle": "ready" if is_ready else "awaiting data",
        "source": source,
    }


def build_summary_cards(session: ReportSession) -> dict:
    wideload_count = get_wideload_count_from_session(session)

    x_total = daily_hour_total_column(session, "X")
    y_total = daily_hour_total_column(session, "Y")
    g_total = daily_hour_total_column(session, "G")
    c_total = daily_hour_total_column(session, "C")
    z_total = daily_hour_total_column(session, "Z")
    r_total = daily_hour_total_column(session, "R")

    return {
        "report_id": session.report_id,
        "station": session.station,
        "bound": session.bound,
        "weighbridge_name": session.weighbridge_name,
        "x_total": x_total if x_total is not None else 0,
        "y_total": y_total if y_total is not None else 0,
        "g_total": g_total if g_total is not None else 0,
        "c_total": c_total if c_total is not None else 0,
        "z_total": z_total if z_total is not None else 0,
        "r_total": r_total if r_total is not None else 0,
        "cases_cleared": session.manual_inputs.get("cases_cleared_in_court", 0) or 0,
        "cards": [
            summary_card(
                "Total Weighed",
                x_total,
                "daily_hour.totals.X",
            ),
            summary_card(
                "Total Overloaded",
                y_total,
                "daily_hour.totals.Y",
            ),
            summary_card(
                "Special Released",
                g_total,
                "daily_hour.totals.G",
            ),
            summary_card(
                "Wide Loads",
                wideload_count,
                "wideload.wideload_count",
            ),
        ],
    }


@router.post("/report-sessions")
async def create_report_session(payload: ReportSessionCreate):
    station = payload.station or payload.weighbridge_name

    if not station:
        raise HTTPException(
            status_code=400,
            detail="Provide either station or weighbridge_name.",
        )

    session = report_session_store.create(
        report_date=payload.report_date,
        station=station,
        bound=payload.bound,
        weighbridge_name=payload.weighbridge_name or station,
        prepared_by=payload.prepared_by,
        confirmed_by=payload.confirmed_by,
    )
    return serialize_session(session)


@router.get("/report-sessions")
async def list_report_sessions():
    sessions = []
    for metadata_path in sorted(report_session_store.sessions_dir.glob("*.json")):
        report_id = metadata_path.stem
        try:
            session = report_session_store.get(report_id)
            if session:
                sessions.append(serialize_session(session))
        except Exception:
            pass
    return sessions


@router.get("/report-sessions/{report_id}")
async def get_report_session(report_id: str):
    return serialize_session(require_session(report_id))


@router.get("/report-sessions/{report_id}/summary-cards")
async def get_report_session_summary_cards(report_id: str):
    return build_summary_cards(require_session(report_id))


def classify_station(station_name: str | None) -> str | None:
    if not station_name:
        return None
    name = station_name.lower()
    if "juja" in name:
        return "Juja"
    if "kanyonyo" in name:
        return "Kanyonyo"
    if "athi" in name:
        return "Athi River"
    if "gilgil" in name:
        return "Gilgil"
    if "isinya" in name:
        return "Isinya"
    if "suswa" in name:
        return "Suswa"
    return None


def is_bound_a(station_code: str, bound_name: str | None) -> bool:
    if not bound_name:
        return True
    bound = bound_name.lower()
    if station_code == "Juja":
        return "thika" in bound or "bound a" in bound or "incoming" in bound
    return "bound a" in bound or "a" in bound or "incoming" in bound


@router.get("/report-sessions/analytics/dashboard")
async def get_analytics_dashboard():
    sessions = []
    for metadata_path in sorted(report_session_store.sessions_dir.glob("*.json")):
        try:
            session = report_session_store.get(metadata_path.stem)
            if session:
                sessions.append(session)
        except Exception:
            pass

    x_total = 0
    y_total = 0
    g_total = 0
    z_total = 0
    r_total = 0
    wideloads = 0
    
    mobile_weighed = 0
    mobile_warned = 0
    mobile_charged = 0

    station_names = {
        "Juja": "Juja Weighbridge",
        "Kanyonyo": "Kanyonyo",
        "Athi River": "Athi River",
        "Gilgil": "Gilgil",
        "Isinya": "Isinya",
        "Suswa": "Suswa"
    }

    stations_data = {
        code: {
            "name": name,
            "code": code,
            "traffic": {"boundA": 0, "boundB": 0},
            "cases": {"boundA": 0, "boundB": 0},
            "compliance": {
                "boundA": {"calledIn": 0, "weighed": 0, "compliant": 0},
                "boundB": {"calledIn": 0, "weighed": 0, "compliant": 0}
            }
        } for code, name in station_names.items()
    }
    
    for s in sessions:
        if s.sections.get("mobile_report", {}).get("status") == "ready":
            summary = s.sections["mobile_report"].get("summary", {})
            mobile_weighed += summary.get("total_trucks_weighed", 0)
            mobile_warned += summary.get("warned_trucks", 0)
            mobile_charged += summary.get("charged_trucks", 0)
            
        if s.sections.get("daily_hour", {}).get("status") == "ready":
            x = daily_hour_total_column(s, "X") or 0
            y = daily_hour_total_column(s, "Y") or 0
            g = daily_hour_total_column(s, "G") or 0
            z = daily_hour_total_column(s, "Z") or 0
            r = daily_hour_total_column(s, "R") or 0
            called = daily_hour_total_column(s, "C") or 0
            cases = s.manual_inputs.get("cases_cleared_in_court", 0) or 0

            x_total += x
            y_total += y
            g_total += g
            z_total += z
            r_total += r
            wideloads += get_wideload_count_from_session(s) or 0

            code = classify_station(s.station or s.weighbridge_name)
            if code and code in stations_data:
                bound_key = "boundA" if is_bound_a(code, s.bound) else "boundB"
                overload_no_permit = max(y - g, 0)
                compliant = max(called - overload_no_permit, 0)

                stations_data[code]["traffic"][bound_key] += x
                stations_data[code]["cases"][bound_key] += cases
                stations_data[code]["compliance"][bound_key]["calledIn"] += called
                stations_data[code]["compliance"][bound_key]["weighed"] += x
                stations_data[code]["compliance"][bound_key]["compliant"] += compliant

    return {
        "static": {
            "weighed": x_total,
            "overloads": max(y_total - g_total, 0),
            "minGross": g_total,
            "chargedRedist": f"{z_total} / {r_total}",
            "reportsGenerated": len([s for s in sessions if s.sections.get("daily_hour", {}).get("status") == "ready"]),
        },
        "mobile": {
            "weighed": mobile_weighed,
            "warned": mobile_warned,
            "charged": mobile_charged,
        },
        "stations": list(stations_data.values())
    }


@router.get("/report-sessions/analytics/details")
async def get_analytics_details():
    sessions = []
    for metadata_path in sorted(report_session_store.sessions_dir.glob("*.json")):
        try:
            session = report_session_store.get(metadata_path.stem)
            if session:
                sessions.append(session)
        except Exception:
            pass

    juja_sessions = [s for s in sessions if s.station and "juja" in s.station.lower()]
    
    juja_thika_traffic = 0
    juja_nairobi_traffic = 0
    juja_thika_cases = 0
    juja_nairobi_cases = 0
    juja_total_called = 0
    juja_total_compliant = 0
    juja_overloads_intercepted = 0
    
    for s in juja_sessions:
        is_thika = is_bound_a("Juja", s.bound)
        
        weighed = daily_hour_total_column(s, "X") or 0
        if is_thika:
            juja_thika_traffic += weighed
        else:
            juja_nairobi_traffic += weighed
            
        cases = s.manual_inputs.get("cases_cleared_in_court", 0) or 0
        if is_thika:
            juja_thika_cases += cases
        else:
            juja_nairobi_cases += cases
            
        called = daily_hour_total_column(s, "C") or 0
        y = daily_hour_total_column(s, "Y") or 0
        g = daily_hour_total_column(s, "G") or 0
        overload_no_permit = max(y - g, 0)
        compliant = max(called - overload_no_permit, 0)
        
        juja_total_called += called
        juja_total_compliant += compliant
        juja_overloads_intercepted += overload_no_permit

    juja_total_traffic = juja_thika_traffic + juja_nairobi_traffic
    juja_total_cases = juja_thika_cases + juja_nairobi_cases
    juja_compliance_rate = (juja_total_compliant / juja_total_called * 100) if juja_total_called > 0 else 0.0

    # Daily breakdown for Juja
    daily_traffic = {}
    daily_cases = {}
    
    for s in juja_sessions:
        try:
            day_str = s.report_date.split("-")[2]  # DD
        except Exception:
            continue
            
        is_thika = is_bound_a("Juja", s.bound)
        weighed = daily_hour_total_column(s, "X") or 0
        cases = s.manual_inputs.get("cases_cleared_in_court", 0) or 0
        
        if day_str not in daily_traffic:
            daily_traffic[day_str] = {"thikaBound": 0, "nairobiBound": 0}
        if day_str not in daily_cases:
            daily_cases[day_str] = {"thikaBound": 0, "nairobiBound": 0}
            
        if is_thika:
            daily_traffic[day_str]["thikaBound"] += weighed
            daily_cases[day_str]["thikaBound"] += cases
        else:
            daily_traffic[day_str]["nairobiBound"] += weighed
            daily_cases[day_str]["nairobiBound"] += cases

    traffic_data = []
    for day in sorted(daily_traffic.keys()):
        traffic_data.append({
            "day": day,
            "thikaBound": daily_traffic[day]["thikaBound"],
            "nairobiBound": daily_traffic[day]["nairobiBound"]
        })
        
    court_cases_data = []
    for day in sorted(daily_cases.keys()):
        court_cases_data.append({
            "day": day,
            "thikaBound": daily_cases[day]["thikaBound"],
            "nairobiBound": daily_cases[day]["nairobiBound"]
        })

    station_names = {
        "Juja": "Juja Weighbridge",
        "Kanyonyo": "Kanyonyo",
        "Athi River": "Athi River",
        "Gilgil": "Gilgil",
        "Isinya": "Isinya",
        "Suswa": "Suswa"
    }
    cross_station = {code: 0 for code in station_names}
    for s in sessions:
        code = classify_station(s.station or s.weighbridge_name)
        if code in cross_station:
            cross_station[code] += s.manual_inputs.get("cases_cleared_in_court", 0) or 0

    cross_station_data = []
    for code, name in station_names.items():
        cross_station_data.append({
            "name": name,
            "cases": cross_station[code],
            "active": code == "Juja"
        })

    return {
        "kpis": {
            "totalTraffic": juja_total_traffic,
            "thikaTraffic": juja_thika_traffic,
            "nairobiTraffic": juja_nairobi_traffic,
            "totalCourtCases": juja_total_cases,
            "thikaCourtCases": juja_thika_cases,
            "nairobiCourtCases": juja_nairobi_cases,
            "complianceRate": round(juja_compliance_rate, 1),
            "overloadsIntercepted": juja_overloads_intercepted
        },
        "trafficData": traffic_data,
        "courtCasesData": court_cases_data,
        "crossStationData": cross_station_data
    }



@router.patch("/report-sessions/{report_id}/metadata")
async def update_report_session_metadata(
    report_id: str,
    payload: ReportSessionMetadataUpdate,
):
    require_session(report_id)

    if payload.station is None and payload.bound is None and payload.weighbridge_name is None:
        raise HTTPException(
            status_code=400,
            detail="Provide station, bound, or weighbridge_name.",
        )

    updated = report_session_store.update_metadata(
        report_id,
        station=payload.station,
        bound=payload.bound,
        weighbridge_name=payload.weighbridge_name,
    )
    return serialize_session(updated)


@router.patch("/report-sessions/{report_id}/manual-inputs")
async def update_report_session_manual_inputs(
    report_id: str,
    payload: ManualInputsUpdate,
):
    require_session(report_id)

    try:
        updated = report_session_store.update_manual_inputs(
            report_id,
            prepared_by=payload.prepared_by,
            confirmed_by=payload.confirmed_by,
            weighbridge_name=payload.weighbridge_name,
            traffic_census=payload.traffic_census,
            transgressions=payload.transgressions,
            extra=payload.extra,
        )
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    return serialize_session(updated)


@router.post("/report-sessions/{report_id}/uploads/daily-hour")
async def upload_daily_hour_file(
    report_id: str,
    file: UploadFile = File(...),
    wideload_count: int = Form(0),
):
    session = require_session(report_id)

    try:
        filename, content, raw_df = await read_upload_dataframe(file)
        report_session_store.save_upload(report_id, "daily_hour", filename, content)

        saved_wideload_count = get_wideload_count_from_session(session)
        count_to_use = (
            saved_wideload_count
            if saved_wideload_count is not None
            else wideload_count
        )

        daily_df = build_daily_hour_metrics(
            raw_df,
            report_date=daily_display_date(session.report_date),
            wideload_count=count_to_use,
        )
        daily_df = add_daily_totals_row(daily_df)

        updated = report_session_store.set_section_ready(
            report_id,
            "daily_hour",
            daily_df,
            filename=filename,
            extra={"wideload_count_used": count_to_use},
        )
        return serialize_session(updated)

    except Exception as exc:
        report_session_store.set_section_error(
            report_id,
            "daily_hour",
            str(exc),
        )

        raise HTTPException(
            status_code=400,
            detail={
                "section": "daily_hour",
                "message": str(exc),
            },
        )

@router.post("/report-sessions/{report_id}/uploads/wideload")
async def upload_wideload_file(report_id: str, file: UploadFile = File(...)):
    try:
        filename, content, raw_df = await read_upload_dataframe(file)
        report_session_store.save_upload(report_id, "wideload", filename, content)

        cleaned_df = clean_with_template(raw_df, vehicle_inspection)
        wideload_count = int(len(cleaned_df))

        updated = report_session_store.set_section_ready(
            report_id,
            "wideload",
            cleaned_df,
            filename=filename,
            extra={"wideload_count": wideload_count},
        )

        session = require_session(report_id)

        if "daily_hour" in session.dataframes:
            daily_raw_df = session.dataframes["daily_hour"]

            if "DATE" in daily_raw_df.columns:
                totals_mask = daily_raw_df["DATE"].astype(str).str.lower().eq("totals")
                if totals_mask.any():
                    daily_raw_df = daily_raw_df.loc[~totals_mask].copy()

            daily_df = update_daily_hour_wideload_count(
                daily_raw_df,
                report_date=session.report_date,
                wideload_count=wideload_count,
            )

            updated = report_session_store.set_section_ready(
                report_id,
                "daily_hour",
                daily_df,
                filename=session.sections["daily_hour"].get("filename"),
                extra={"wideload_count_used": wideload_count},
            )

        return serialize_session(updated)

    except Exception as exc:
        report_session_store.set_section_error(
            report_id,
            "wideload",
            str(exc),
        )

    raise HTTPException(
        status_code=400,
        detail={
            "section": "wideload",
            "message": str(exc),
        },
    )
    

@router.post("/report-sessions/{report_id}/uploads/impounded-prohibited")
async def upload_impounded_prohibited_file(
    report_id: str,
    file: UploadFile = File(...),
):
    require_session(report_id)

    try:
        filename, content, raw_df = await read_upload_dataframe(file)
        report_session_store.save_upload(
            report_id,
            "impounded_prohibited",
            filename,
            content,
        )
        cleaned_df = clean_with_template(raw_df, impounded_prohibited)
        updated = report_session_store.set_section_ready(
            report_id,
            "impounded_prohibited",
            cleaned_df,
            filename=filename,
        )
        return serialize_session(updated)

    except Exception as exc:
        report_session_store.set_section_error(
            report_id,
            "impounded_prohibited",
            str(exc),
        )

    raise HTTPException(
        status_code=400,
        detail={
            "section": "impounded_prohibited",
            "message": str(exc),
        },
    )
    


@router.post("/report-sessions/{report_id}/uploads/overloaded")
async def upload_overloaded_file(report_id: str, file: UploadFile = File(...)):
    require_session(report_id)

    try:
        filename, content, raw_df = await read_upload_dataframe(file)
        report_session_store.save_upload(report_id, "overloaded", filename, content)
        valid_permit_count = count_valid_permit_vehicles(raw_df)
        updated = report_session_store.set_section_ready(
            report_id,
            "overloaded",
            raw_df,
            filename=filename,
            extra={"valid_permit_count": valid_permit_count},
        )
        return serialize_session(updated)

    except Exception as exc:
        report_session_store.set_section_error(
            report_id,
            "overloaded",
            str(exc),
        )

    raise HTTPException(
        status_code=400,
        detail={
            "section": "overloaded",
            "message": str(exc),
        },
    )


@router.post("/report-sessions/{report_id}/uploads/mobile-report")
async def upload_mobile_report_file(report_id: str, file: UploadFile = File(...)):
    require_session(report_id)

    try:
        filename, content, raw_df = await read_upload_dataframe(file)
        report_session_store.save_upload(
            report_id,
            "mobile_report",
            filename,
            content,
        )
        records = normalize_mobile_report(raw_df)
        summary = summarize_mobile_report(records)

        updated = report_session_store.set_section_ready(
            report_id,
            "mobile_report",
            records,
            filename=filename,
            extra={"summary": summary},
        )
        payload = serialize_session(updated)
        payload["mobile_report"] = mobile_report_response(raw_df)
        return payload

    except Exception as exc:
        report_session_store.set_section_error(
            report_id,
            "mobile_report",
            str(exc),
        )

        raise HTTPException(
            status_code=400,
            detail={
                "section": "mobile_report",
                "message": str(exc),
            },
        )


@router.post("/report-sessions/{report_id}/build-final-report")
async def build_report_session_final_report(report_id: str):
    session = require_session(report_id)
    required_sections = [
        "daily_hour",
        "wideload",
        "impounded_prohibited",
        "overloaded",
    ]
    missing = [
        section
        for section in required_sections
        if section not in session.dataframes
        or session.sections.get(section, {}).get("status") != "ready"
    ]

    if missing:
        message = f"Missing or invalid required sections: {missing}"
        updated = report_session_store.set_final_report_error(report_id, message)
        raise HTTPException(
            status_code=400,
            detail={
                "message": message,
                "missing_sections": missing,
                "session": serialize_session(updated),
            },
        )

    try:
        wideload_count = get_wideload_count_from_session(session)

        if wideload_count is None:
            wideload_count = len(session.dataframes["wideload"])

        file_stream = build_final_report(
            daily_df=session.dataframes["daily_hour"],
            wideload_df=session.dataframes["wideload"],
            impounded_prohibited_df=session.dataframes["impounded_prohibited"],
            overloaded_df=session.dataframes["overloaded"],
            report_date=session.report_date,
            station=session.station,
            bound=session.bound,
            prepared_by=session.prepared_by,
            confirmed_by=session.confirmed_by,
            traffic_census=session.manual_inputs.get("traffic_census"),
            daily_summary=session.sections.get("daily_summary", {}).get("values"),
            transgressions=session.manual_inputs.get("transgressions"),
            wideload_count=wideload_count,
        )
        updated = report_session_store.set_final_report(
            report_id,
            file_stream.read(),
        )
        return serialize_session(updated)

    except Exception as exc:
        logger.exception(
            "Failed to build final report for report session %s",
            report_id,
        )
        updated = report_session_store.set_final_report_error(report_id, str(exc))
        raise HTTPException(
            status_code=500,
            detail={
                "message": "Failed to build final report",
                "error": str(exc),
                "session": serialize_session(updated),
            },
        )


@router.get("/report-sessions/{report_id}/sections/{section_name}/preview")
async def preview_report_session_section(
    report_id: str,
    section_name: str,
    format: str = "png",
    page: int | None = None,
):
    session = require_session(report_id)

    try:
        preview = get_cached_section_preview(
            session,
            section_name,
            preview_format=format,
            page=page,
            store=report_session_store,
        )
    except (RuntimeError, ValueError) as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    disposition = "inline" if preview.inline else "attachment"

    return Response(
        content=preview.stream.getvalue(),
        media_type=preview.media_type,
        headers={"Content-Disposition": f"{disposition}; filename={preview.filename}"},
    )


@router.get("/report-sessions/{report_id}/download-final-report")
async def download_report_session_final_report(report_id: str):
    session = require_session(report_id)

    if session.final_report is None:
        raise HTTPException(status_code=404, detail="Final report is not ready")

    filename = get_report_filename(session, "docx")

    return Response(
        content=session.final_report,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )


@router.get("/report-sessions/{report_id}/download-pdf-report")
async def download_report_session_pdf_report(report_id: str):
    import io
    from app.services.preview_renderer import convert_docx_to_pdf
    session = require_session(report_id)

    if session.final_report is None:
        raise HTTPException(status_code=404, detail="Final report is not ready")

    docx_filename = get_report_filename(session, "docx")
    docx_stream = io.BytesIO(session.final_report)

    try:
        pdf_bytes, pdf_filename = convert_docx_to_pdf(docx_stream, docx_filename)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"PDF conversion failed: {str(exc)}") from exc

    return Response(
        content=pdf_bytes,
        media_type="application/pdf",
        headers={"Content-Disposition": f"attachment; filename={pdf_filename}"},
    )


@router.get("/report-sessions/{report_id}/download-excel-report")
async def download_report_session_excel_report(report_id: str):
    session = require_session(report_id)

    if (
        "daily_hour" not in session.dataframes
        or session.sections.get("daily_hour", {}).get("status") != "ready"
    ):
        raise HTTPException(status_code=400, detail="Daily hour data is not ready")

    try:
        file_stream = build_excel_report(session)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    filename = get_report_filename(session, "xlsx")

    return Response(
        content=file_stream.getvalue(),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )


@router.get("/report-sessions/{report_id}/download-mobile-excel-report")
async def download_report_session_mobile_excel_report(report_id: str):
    session = require_session(report_id)

    try:
        file_stream = build_mobile_excel_report(session)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    filename = get_mobile_report_filename(session, "xlsx")

    return Response(
        content=file_stream.getvalue(),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": f'attachment; filename="{filename}"',
            "X-Content-Type-Options": "nosniff",
        },
    )


@router.get("/report-sessions/{report_id}/download-mobile-word-report")
async def download_report_session_mobile_word_report(report_id: str):
    session = require_session(report_id)

    try:
        file_stream = build_mobile_word_report(session)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    filename = get_mobile_report_filename(session, "docx")

    return Response(
        content=file_stream.getvalue(),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={
            "Content-Disposition": f'attachment; filename="{filename}"',
            "X-Content-Type-Options": "nosniff",
        },
    )


@router.get("/report-sessions/sms-summaries/dates")
async def get_sms_summary_dates():
    dates = set()
    for metadata_path in report_session_store.sessions_dir.glob("*.json"):
        try:
            session = report_session_store.get(metadata_path.stem)
            if session and session.report_date:
                dates.add(session.report_date)
        except Exception:
            pass
    return sorted(list(dates), reverse=True)


@router.get("/report-sessions/sms-summaries/{report_date}")
async def get_sms_summaries_by_date(report_date: str):
    sessions_on_date = []
    for metadata_path in report_session_store.sessions_dir.glob("*.json"):
        try:
            session = report_session_store.get(metadata_path.stem)
            if session and session.report_date == report_date:
                sessions_on_date.append(session)
        except Exception:
            pass

    static_a = None
    static_b = None
    mobile_1 = None
    mobile_2 = None

    for s in sessions_on_date:
        bound_lower = (s.bound or "").lower()
        station_lower = (s.station or s.weighbridge_name or "").lower()
        is_mobile = "mobile" in bound_lower or "mobile" in station_lower or "mobile_report" in s.sections
        
        if is_mobile:
            if "2" in bound_lower or "two" in bound_lower:
                mobile_2 = s
            else:
                mobile_1 = s
        else:
            station_code = classify_station(s.station or s.weighbridge_name)
            if is_bound_a(station_code, s.bound):
                static_a = s
            else:
                static_b = s

    try:
        date_formatted = datetime.strptime(report_date, "%Y-%m-%d").strftime("%d.%m.%Y")
    except Exception:
        date_formatted = report_date

    from app.services.sms_summary_builder import build_static_sms_summary, build_mobile_sms_summary

    response = []

    if static_a:
        response.append({
            "slot": "static_bound_a",
            "title": f"Static: {static_a.weighbridge_name or 'JUJA'} - {static_a.bound or 'Thika Bound'}",
            "exists": True,
            "report_id": static_a.report_id,
            "text": build_static_sms_summary(static_a)
        })
    else:
        response.append({
            "slot": "static_bound_a",
            "title": "Static: JUJA - THIKA BOUND",
            "exists": False,
            "report_id": None,
            "text": f"DAILY REPORT\nJUJA THIKA BOUND WB\nDate: {date_formatted}\n\n[Awaiting report upload and processing]"
        })

    if static_b:
        response.append({
            "slot": "static_bound_b",
            "title": f"Static: {static_b.weighbridge_name or 'JUJA'} - {static_b.bound or 'Nairobi Bound'}",
            "exists": True,
            "report_id": static_b.report_id,
            "text": build_static_sms_summary(static_b)
        })
    else:
        response.append({
            "slot": "static_bound_b",
            "title": "Static: JUJA - NAIROBI BOUND",
            "exists": False,
            "report_id": None,
            "text": f"DAILY REPORT\nJUJA NAIROBI BOUND WB\nDate: {date_formatted}\n\n[Awaiting report upload and processing]"
        })

    if mobile_1:
        response.append({
            "slot": "mobile_1",
            "title": f"Mobile: {mobile_1.station or 'JUJA'} - TEAM ONE",
            "exists": True,
            "report_id": mobile_1.report_id,
            "text": build_mobile_sms_summary(mobile_1)
        })
    else:
        response.append({
            "slot": "mobile_1",
            "title": "Mobile: JUJA - TEAM ONE",
            "exists": False,
            "report_id": None,
            "text": f"DAILY REPORT\nJUJA W/B DAILY MOBILE REPORT_TEAM ONE\nDate: {date_formatted}\n\n[Awaiting report upload and processing]"
        })

    if mobile_2:
        response.append({
            "slot": "mobile_2",
            "title": f"Mobile: {mobile_2.station or 'JUJA'} - TEAM TWO",
            "exists": True,
            "report_id": mobile_2.report_id,
            "text": build_mobile_sms_summary(mobile_2)
        })
    else:
        response.append({
            "slot": "mobile_2",
            "title": "Mobile: JUJA - TEAM TWO",
            "exists": False,
            "report_id": None,
            "text": f"DAILY REPORT\nJUJA W/B DAILY MOBILE REPORT_TEAM TWO\nDate: {date_formatted}\n\n[Awaiting report upload and processing]"
        })

    return response
