import logging
import os
import secrets
from datetime import datetime

from fastapi import APIRouter, File, Form, Header, HTTPException, Query, UploadFile, Depends, BackgroundTasks
from fastapi.responses import Response
from pydantic import BaseModel

from app.db.models import User
from app.routes.auth import get_current_user
from app.core.security import decode_access_token
from app.services.report_worker import enqueue_build_final_report

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
    report_date: str | None = None
    station: str | None = None
    bound: str | None = None
    weighbridge_name: str | None = None
    prepared_by: str | None = None
    confirmed_by: str | None = None


class ReportSessionHistoryItem(BaseModel):
    report_id: str
    title: str | None = None
    report_type: str
    weighbridge_name: str | None = None
    bound_name: str | None = None
    status: str
    created_at: datetime
    updated_at: datetime
    completed_at: datetime | None = None
    has_final_report: bool
    upload_count: int
    required_uploads_completed: bool
    manual_inputs_completed: bool
    download_available: bool


def require_admin_password(
    x_admin_password: str | None,
    authorization: str | None = None
) -> None:
    # Try JWT Authentication first if header is present
    if authorization and authorization.startswith("Bearer "):
        token = authorization.split(" ")[1]
        payload = decode_access_token(token)
        if payload:
            role = payload.get("role")
            username = payload.get("sub")
            if role == "admin" or username == "admin":
                return

    configured_password = os.getenv("ADMIN_PASSWORD")

    if not configured_password:
        raise HTTPException(
            status_code=503,
            detail="Admin password is not configured.",
        )

    if not x_admin_password or not secrets.compare_digest(
        x_admin_password,
        configured_password,
    ):
        raise HTTPException(
            status_code=401,
            detail="Invalid admin password or session token.",
        )


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
    station_name = (session.station or session.weighbridge_name or "STATION").upper()
    station_name = " ".join(
        part for part in station_name.split() if part != "MOBILE"
    ).strip()
    if not station_name:
        station_name = "STATION"
    if "WEIGHBRIDGE" not in station_name:
        station_name = f"{station_name} WEIGHBRIDGE"
    date_part = format_filename_date(session.report_date)
    bound_name = (session.bound or "").lower()
    report_number = "2" if "2" in bound_name or "two" in bound_name else "1"
    return f"{station_name} MOBILE DAILY REPORT {report_number} {date_part}.{ext}"


def serialize_session(session: ReportSession) -> dict:
    excel_report_ready = (
        session.sections.get("daily_hour", {}).get("status") == "ready"
    )
    mobile_excel_report_ready = (
        session.sections.get("mobile_report", {}).get("status") == "ready"
    )
    mobile_word_report_ready = mobile_excel_report_ready
    final_report: dict[str, str | None] = {
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
        daily_hour_section = session.sections.get("daily_hour", {})
        if isinstance(daily_hour_section, dict) and "summary" in daily_hour_section:
            summary = daily_hour_section["summary"]
            if isinstance(summary, dict) and column in summary:
                try:
                    return int(summary[column])
                except Exception:
                    pass
        return None

    daily_df = session.dataframes["daily_hour"]

    if "DATE" not in daily_df.columns or column not in daily_df.columns:
        daily_hour_section = session.sections.get("daily_hour", {})
        if isinstance(daily_hour_section, dict) and "summary" in daily_hour_section:
            summary = daily_hour_section["summary"]
            if isinstance(summary, dict) and column in summary:
                try:
                    return int(summary[column])
                except Exception:
                    pass
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


def available_report_sessions() -> list[tuple[ReportSession, float]]:
    report_ids = report_session_store.list_report_ids()
    total_ids = len(report_ids)
    sessions: list[tuple[ReportSession, float]] = []

    for index, report_id in enumerate(report_ids):
        try:
            session = report_session_store.get(report_id)
            if not session:
                continue

            summary = report_session_store.report_history_summary(report_id)
            if summary and "updated_at" in summary:
                u_at = summary["updated_at"]
                if isinstance(u_at, datetime):
                    modified_at = u_at.timestamp()
                else:
                    modified_at = float(u_at)
            else:
                metadata_path = report_session_store.sessions_dir / f"{report_id}.json"
                modified_at = (
                    metadata_path.stat().st_mtime
                    if metadata_path.exists()
                    else float(total_ids - index)
                )
            sessions.append((session, modified_at))
        except Exception:
            logger.exception("Failed to load report session for analytics: %s", report_id)

    return sessions


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


def check_write_permission(current_user: User = Depends(get_current_user)):
    if current_user.role in ("duty_manager", "cluster_manager"):
        raise HTTPException(
            status_code=403,
            detail="Access denied: Your account role does not allow modifications."
        )
    return current_user


@router.post("/report-sessions")
async def create_report_session(payload: ReportSessionCreate, current_user: User = Depends(check_write_permission)):
    station = payload.station or payload.weighbridge_name

    # Override for non-admins to lock station and prepared_by
    if current_user.role != "admin":
        if current_user.station:
            station = current_user.station
        payload.prepared_by = current_user.full_name or current_user.username

    if not station:
        raise HTTPException(
            status_code=400,
            detail="Provide either station or weighbridge_name.",
        )

    session = report_session_store.create(
        report_date=payload.report_date,
        station=station,
        bound=payload.bound,
        weighbridge_name=station,
        prepared_by=payload.prepared_by,
        confirmed_by=payload.confirmed_by,
    )
    return serialize_session(session)



@router.get("/report-sessions")
async def list_report_sessions(
    status: str | None = Query(default=None),
    report_type: str | None = Query(default=None),
    limit: int = Query(default=20, ge=1, le=100),
    offset: int = Query(default=0, ge=0),
    search: str | None = Query(default=None, max_length=120),
    x_admin_password: str | None = Header(default=None),
    authorization: str | None = Header(default=None),
):
    require_admin_password(x_admin_password, authorization)
    items = []
    summaries = report_session_store.list_report_history(
        status=status,
        report_type=report_type,
        limit=limit,
        offset=offset,
        search=search,
    )

    for summary in summaries:
        history_payload = ReportSessionHistoryItem.model_validate(summary).model_dump(
            mode="json"
        )
        report_id = history_payload["report_id"]

        try:
            session = report_session_store.get(report_id)
        except Exception:
            session = None

        if session:
            payload = serialize_session(session)
        else:
            payload = {
                "report_id": report_id,
                "metadata": {
                    "report_date": None,
                    "station": history_payload["weighbridge_name"],
                    "bound": history_payload["bound_name"],
                    "weighbridge_name": history_payload["weighbridge_name"],
                    "prepared_by": None,
                    "confirmed_by": None,
                },
                "manual_inputs": {},
                "sections": {},
                "final_report": {
                    "status": "ready"
                    if history_payload["download_available"]
                    else "not_built",
                    "download_url": (
                        f"/api/report-sessions/{report_id}/download-final-report"
                        if history_payload["download_available"]
                        else None
                    ),
                    "error": None,
                },
                "excel_report": {"status": "awaiting_data", "download_url": None},
                "mobile_excel_report": {
                    "status": "awaiting_data",
                    "download_url": None,
                },
                "mobile_word_report": {
                    "status": "awaiting_data",
                    "download_url": None,
                },
            }

        payload.update(history_payload)
        items.append(payload)

    return items


@router.get("/report-sessions/{report_id}")
async def get_report_session(report_id: str):
    return serialize_session(require_session(report_id))


@router.delete("/report-sessions/{report_id}")
async def delete_report_session(
    report_id: str,
    x_admin_password: str | None = Header(default=None),
    authorization: str | None = Header(default=None),
):
    require_admin_password(x_admin_password, authorization)
    deleted = report_session_store.delete(report_id)

    if not deleted:
        raise HTTPException(status_code=404, detail="Report session not found")

    return {"status": "deleted", "report_id": report_id}


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


def is_bound_a(station_code: str | None, bound_name: str | None) -> bool:
    if not bound_name:
        return True
    bound = bound_name.lower()
    if not station_code:
        return "bound a" in bound or "incoming" in bound or "thika" in bound or "mombasa" in bound or "mwingi" in bound or "kajiado" in bound or "narok" in bound
    
    station_lower = station_code.lower()
    if "juja" in station_lower:
        return "thika" in bound or "bound a" in bound or "incoming" in bound
    elif "athi" in station_lower:
        return "mombasa" in bound or "bound a" in bound or "incoming" in bound
    elif "gilgil" in station_lower:
        return "nairobi" in bound or "bound a" in bound or "incoming" in bound
    elif "kanyonyo" in station_lower:
        return "mwingi" in bound or "bound a" in bound or "incoming" in bound
    elif "isinya" in station_lower:
        return "kajiado" in bound or "bound a" in bound or "incoming" in bound
    elif "suswa" in station_lower:
        return "narok" in bound or "bound a" in bound or "incoming" in bound

    return "bound a" in bound or "incoming" in bound or "thika" in bound or "mombasa" in bound or "mwingi" in bound or "kajiado" in bound or "narok" in bound


def mobile_report_slot(bound_name: str | None) -> str:
    bound = (bound_name or "").strip().lower()
    if "2" in bound or "two" in bound:
        return "mobile_2"
    return "mobile_1"


def mobile_report_label(slot: str) -> str:
    return "Mobile 2" if slot == "mobile_2" else "Mobile 1"


def mobile_report_manual_inputs(session: ReportSession) -> dict:
    mobile_inputs = session.manual_inputs.get("mobile_report")
    if isinstance(mobile_inputs, dict):
        return mobile_inputs

    extra_inputs = session.manual_inputs.get("extra")
    if isinstance(extra_inputs, dict) and isinstance(extra_inputs.get("mobile_report"), dict):
        return extra_inputs["mobile_report"]

    return {}


def danka_staff_names(session: ReportSession) -> list[str]:
    staff_value = str(mobile_report_manual_inputs(session).get("danka_staff") or "").strip()
    if not staff_value:
        return []

    names = []
    for part in staff_value.replace("\\", "/").split("/"):
        cleaned = " ".join(part.strip().split())
        if cleaned:
            names.append(cleaned.upper())

    return names


def danka_staff_team(session: ReportSession) -> dict | None:
    names = danka_staff_names(session)
    if not names:
        return None

    dm_name = next((name for name in names if "DM" in name.split()), names[0])
    drivers = [name for name in names if name != dm_name]

    return {
        "dm": dm_name,
        "drivers": drivers,
        "team": " / ".join([dm_name, *drivers]),
    }


def normalize_mobile_filter(value: str | None) -> str | None:
    if value is None:
        return None
    cleaned = value.strip().lower().replace("-", "_").replace(" ", "_")
    if cleaned in {"mobile_2", "2", "two"}:
        return "mobile_2"
    if cleaned in {"mobile_1", "1", "one"}:
        return "mobile_1"
    return cleaned or None


def static_bound_key(session: ReportSession) -> str:
    station_code = classify_station(session.station or session.weighbridge_name)
    return "boundA" if is_bound_a(station_code, session.bound) else "boundB"


def empty_static_kpis() -> dict:
    return {
        "label": "",
        "weighed": 0,
        "overloads": 0,
        "psvOverloads": 0,
        "minGross": 0,
        "charged": 0,
        "redistributed": 0,
        "chargedRedist": "0 / 0",
        "reportsGenerated": 0,
    }


def add_static_kpis(target: dict, session: ReportSession) -> None:
    y = daily_hour_total_column(session, "Y") or 0
    g = daily_hour_total_column(session, "G") or 0
    z = daily_hour_total_column(session, "Z") or 0
    r = daily_hour_total_column(session, "R") or 0

    target["weighed"] += daily_hour_total_column(session, "X") or 0
    target["overloads"] += max(y - g, 0)
    target["minGross"] += g
    target["charged"] += z
    target["redistributed"] += r
    target["reportsGenerated"] += 1
    target["chargedRedist"] = f"{target['charged']} / {target['redistributed']}"


@router.get("/report-sessions/analytics/dashboard")
async def get_analytics_dashboard(
    static_date: str | None = None,
    mobile_date: str | None = None,
    mobile_bound: str | None = None,
):
    sessions = []
    session_modified_at: dict[str, float] = {}
    for session, modified_at in available_report_sessions():
        sessions.append(session)
        session_modified_at[session.report_id] = modified_at

    latest_static_sessions: dict[tuple[str, str, str], tuple[ReportSession, float]] = {}
    latest_mobile_sessions: dict[tuple[str, str], tuple[ReportSession, float]] = {}

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
            slot = mobile_report_slot(s.bound)
            key = (s.report_date, slot)
            modified_at = session_modified_at.get(s.report_id, 0)
            previous = latest_mobile_sessions.get(key)
            if previous is None or modified_at >= previous[1]:
                latest_mobile_sessions[key] = (s, modified_at)
            
        if s.sections.get("daily_hour", {}).get("status") == "ready":
            code = classify_station(s.station or s.weighbridge_name)
            station_key = code or (s.station or s.weighbridge_name or "").strip().lower()
            bound_key = static_bound_key(s)
            modified_at = session_modified_at.get(s.report_id, 0)
            key = (s.report_date, station_key, bound_key)
            previous = latest_static_sessions.get(key)
            if previous is None or modified_at >= previous[1]:
                latest_static_sessions[key] = (s, modified_at)

    static_dates = sorted(
        {report_date for report_date, _, _ in latest_static_sessions},
        reverse=True,
    )
    selected_static_date = (
        static_date if static_date in static_dates else (static_dates[0] if static_dates else None)
    )

    static_by_bound = {
        "boundA": {**empty_static_kpis(), "label": "Bound A"},
        "boundB": {**empty_static_kpis(), "label": "Bound B"},
        "total": {**empty_static_kpis(), "label": "Total"},
    }

    for (report_date, _, bound_key), (s, _) in latest_static_sessions.items():
        x = daily_hour_total_column(s, "X") or 0
        y = daily_hour_total_column(s, "Y") or 0
        g = daily_hour_total_column(s, "G") or 0
        called = daily_hour_total_column(s, "C") or 0
        cases = s.manual_inputs.get("cases_cleared_in_court", 0) or 0

        if selected_static_date and report_date == selected_static_date:
            code = classify_station(s.station or s.weighbridge_name)
            if code and code in stations_data:
                overload_no_permit = max(y - g, 0)
                compliant = max(called - overload_no_permit, 0)

                stations_data[code]["traffic"][bound_key] += x
                stations_data[code]["cases"][bound_key] += cases
                stations_data[code]["compliance"][bound_key]["calledIn"] += called
                stations_data[code]["compliance"][bound_key]["weighed"] += x
                stations_data[code]["compliance"][bound_key]["compliant"] += compliant


        if selected_static_date and report_date == selected_static_date:
            if s.bound:
                static_by_bound[bound_key]["label"] = s.bound
            add_static_kpis(static_by_bound[bound_key], s)
            add_static_kpis(static_by_bound["total"], s)

    mobile_reports = [
        {
            "date": report_date,
            "bound": slot,
            "bound_label": mobile_report_label(slot),
            "label": f"{report_date} - {mobile_report_label(slot)}",
            "report_id": session.report_id,
            "updated_at": modified_at,
        }
        for (report_date, slot), (session, modified_at) in latest_mobile_sessions.items()
    ]
    mobile_reports.sort(
        key=lambda item: (
            item["date"],
            1 if item["bound"] == "mobile_2" else 0,
            item["updated_at"],
        ),
        reverse=True,
    )

    selected_mobile = None
    selected_mobile_bound = normalize_mobile_filter(mobile_bound)
    for option in mobile_reports:
        if mobile_date and option["date"] != mobile_date:
            continue
        if selected_mobile_bound and option["bound"] != selected_mobile_bound:
            continue
        selected_mobile = option
        break

    selected_session = None
    if selected_mobile:
        selected_key = (str(selected_mobile["date"]), str(selected_mobile["bound"]))
        selected_session = latest_mobile_sessions[selected_key][0]

    mobile_summary = (
        selected_session.sections["mobile_report"].get("summary", {})
        if selected_session is not None
        else {}
    )

    return {
        "static": {
            "weighed": static_by_bound["total"]["weighed"],
            "overloads": static_by_bound["total"]["overloads"],
            "psvOverloads": static_by_bound["total"]["psvOverloads"],
            "minGross": static_by_bound["total"]["minGross"],
            "chargedRedist": static_by_bound["total"]["chargedRedist"],
            "reportsGenerated": static_by_bound["total"]["reportsGenerated"],
            "dates": static_dates,
            "selectedDate": selected_static_date,
            "byBound": static_by_bound,
        },
        "mobile": {
            "weighed": mobile_summary.get("total_trucks_weighed", 0),
            "warned": mobile_summary.get("warned_trucks", 0),
            "charged": mobile_summary.get("charged_trucks", 0),
            "reports": mobile_reports,
            "selected": selected_mobile,
        },
        "stations": list(stations_data.values())
    }


@router.get("/report-sessions/analytics/details")
async def get_analytics_details():
    sessions = [session for session, _ in available_report_sessions()]

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


@router.get("/report-sessions/analytics/dms-performance")
async def get_dms_performance(date: str | None = None):
    if not date:
        date = datetime.now().strftime("%Y-%m-%d")
        
    try:
        filter_date = datetime.strptime(date, "%Y-%m-%d")
    except Exception:
        filter_date = datetime.now()

    latest_mobile_sessions: dict[tuple[str, str, str], tuple[ReportSession, float]] = {}

    for session, modified_at in available_report_sessions():
        try:
            if not session or session.sections.get("mobile_report", {}).get("status") != "ready":
                continue

            try:
                session_date = datetime.strptime(session.report_date, "%Y-%m-%d")
            except Exception:
                continue

            # Limit to sessions in the same month/year as the filter date, up to the filter date
            if session_date.year != filter_date.year or session_date.month != filter_date.month:
                continue
            if session_date > filter_date:
                continue

            station_key = (
                classify_station(session.station or session.weighbridge_name)
                or (session.station or session.weighbridge_name or "").strip().lower()
            )
            key = (session.report_date, station_key, mobile_report_slot(session.bound))
            previous = latest_mobile_sessions.get(key)
            if previous is None or modified_at >= previous[1]:
                latest_mobile_sessions[key] = (session, modified_at)
        except Exception:
            pass

    stats: dict[str, dict] = {}
    report_count = 0

    for session, _ in latest_mobile_sessions.values():
        team = danka_staff_team(session)
        if not team:
            continue

        summary = session.sections.get("mobile_report", {}).get("summary", {})
        weighed = int(summary.get("total_trucks_weighed", 0) or 0)
        charged = int(summary.get("charged_trucks", 0) or 0)
        report_count += 1

        is_current_month = False
        try:
            report_date = datetime.strptime(session.report_date, "%Y-%m-%d")
            is_current_month = report_date.year == filter_date.year and report_date.month == filter_date.month
        except Exception:
            pass

        dm_name = team["dm"]
        row = stats.setdefault(
            dm_name,
            {
                "name": dm_name,
                "surname": dm_name.split()[-1],
                "team": team["team"],
                "drivers": [],
                "weighed": 0,
                "charged": 0,
                "monthCharged": 0,
                "reports": 0,
            },
        )
        for driver in team["drivers"]:
            if driver not in row["drivers"]:
                row["drivers"].append(driver)
        row["team"] = " / ".join([dm_name, *row["drivers"]])
        row["weighed"] += weighed
        row["charged"] += charged
        row["reports"] += 1
        if is_current_month:
            row["monthCharged"] += charged

    for row in stats.values():
        row["chargeRate"] = round((row["charged"] / row["weighed"] * 100), 1) if row["weighed"] else 0

    rows = sorted(
        stats.values(),
        key=lambda item: (item["charged"], item["chargeRate"], item["weighed"], item["name"]),
        reverse=True,
    )

    return {
        "rows": rows,
        "totalCharged": sum(row["charged"] for row in rows),
        "totalWeighed": sum(row["weighed"] for row in rows),
        "reports": report_count,
    }



@router.patch("/report-sessions/{report_id}/metadata")
async def update_report_session_metadata(
    report_id: str,
    payload: ReportSessionMetadataUpdate,
    current_user: User = Depends(check_write_permission),
):
    require_session(report_id)

    # Override/lock for non-admins
    if current_user.role != "admin":
        if current_user.station:
            payload.station = current_user.station
            payload.weighbridge_name = current_user.station
        payload.prepared_by = current_user.full_name or current_user.username

    if (
        payload.report_date is None
        and payload.station is None
        and payload.bound is None
        and payload.weighbridge_name is None
        and payload.prepared_by is None
        and payload.confirmed_by is None
    ):
        raise HTTPException(
            status_code=400,
            detail="Provide at least one metadata field.",
        )

    updated = report_session_store.update_metadata(
        report_id,
        report_date=payload.report_date,
        station=payload.station,
        bound=payload.bound,
        weighbridge_name=payload.weighbridge_name,
        prepared_by=payload.prepared_by,
        confirmed_by=payload.confirmed_by,
    )
    return serialize_session(updated)


@router.patch("/report-sessions/{report_id}/manual-inputs")
async def update_report_session_manual_inputs(
    report_id: str,
    payload: ManualInputsUpdate,
    current_user: User = Depends(check_write_permission),
):
    require_session(report_id)

    # Override/lock for non-admins
    if current_user.role != "admin":
        if current_user.station:
            payload.weighbridge_name = current_user.station
        payload.prepared_by = current_user.full_name or current_user.username

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
    current_user: User = Depends(check_write_permission),
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
async def upload_wideload_file(
    report_id: str,
    file: UploadFile = File(...),
    current_user: User = Depends(check_write_permission),
):
    try:
        filename, content, raw_df = await read_upload_dataframe(file)
        report_session_store.save_upload(report_id, "wideload", filename, content)

        cleaned_df = clean_with_template(raw_df, vehicle_inspection)
        wideload_count = len(cleaned_df)

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
    current_user: User = Depends(check_write_permission),
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
async def upload_overloaded_file(
    report_id: str,
    file: UploadFile = File(...),
    current_user: User = Depends(check_write_permission),
):
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
async def upload_mobile_report_file(
    report_id: str,
    file: UploadFile = File(...),
    current_user: User = Depends(check_write_permission),
):
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
async def build_report_session_final_report(
    report_id: str,
    background_tasks: BackgroundTasks,
    current_user: User = Depends(check_write_permission),
):
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

    # Set status to processing
    updated = report_session_store.set_report_processing(report_id)

    # Enqueue background build task
    enqueue_build_final_report(report_id, background_tasks)

    return serialize_session(updated)



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
    for session, _ in available_report_sessions():
        try:
            has_sms_data = (
                session
                and session.report_date
                and (
                    session.sections.get("daily_hour", {}).get("status") == "ready"
                    or session.sections.get("mobile_report", {}).get("status") == "ready"
                )
            )
            if has_sms_data:
                dates.add(session.report_date)
        except Exception:
            pass
    return sorted(list(dates), reverse=True)


@router.get("/report-sessions/sms-summaries/{report_date}")
async def get_sms_summaries_by_date(report_date: str, station: str | None = None):
    station_val = station or "Juja"
    sessions_on_date: list[tuple[ReportSession, float]] = []
    for session, modified_at in available_report_sessions():
        try:
            if session and session.report_date == report_date:
                session_station = classify_station(session.station or session.weighbridge_name)
                matched = False
                if session_station and station_val.lower() == session_station.lower():
                    matched = True
                elif station_val.lower() in (session.station or "").lower():
                    matched = True
                elif station_val.lower() in (session.weighbridge_name or "").lower():
                    matched = True
                
                if matched:
                    sessions_on_date.append((session, modified_at))
        except Exception:
            pass

    static_a = None
    static_b = None
    mobile_1 = None
    mobile_2 = None
    static_a_modified_at = 0
    static_b_modified_at = 0
    mobile_1_modified_at = 0
    mobile_2_modified_at = 0

    for s, modified_at in sessions_on_date:
        bound_lower = (s.bound or "").lower()
        station_lower = (s.station or s.weighbridge_name or "").lower()
        is_mobile = "mobile" in bound_lower or "mobile" in station_lower or "mobile_report" in s.sections
        
        if is_mobile and s.sections.get("mobile_report", {}).get("status") == "ready":
            if mobile_report_slot(s.bound) == "mobile_2":
                if modified_at >= mobile_2_modified_at:
                    mobile_2 = s
                    mobile_2_modified_at = modified_at
            elif modified_at >= mobile_1_modified_at:
                mobile_1 = s
                mobile_1_modified_at = modified_at
        elif s.sections.get("daily_hour", {}).get("status") == "ready":
            station_code = classify_station(s.station or s.weighbridge_name)
            if is_bound_a(station_code, s.bound):
                if modified_at >= static_a_modified_at:
                    static_a = s
                    static_a_modified_at = modified_at
            elif modified_at >= static_b_modified_at:
                static_b = s
                static_b_modified_at = modified_at

    try:
        date_formatted = datetime.strptime(report_date, "%Y-%m-%d").strftime("%d.%m.%Y")
    except Exception:
        date_formatted = report_date

    from app.services.sms_summary_builder import build_static_sms_summary, build_mobile_sms_summary

    station_upper = station_val.strip().upper()
    
    # Determine default bound names based on station
    bound_a_name = "THIKA BOUND"
    bound_b_name = "NAIROBI BOUND"
    
    if "ATHI" in station_upper:
        bound_a_name = "MOMBASA BOUND"
        bound_b_name = "NAIROBI BOUND"
    elif "GILGIL" in station_upper:
        bound_a_name = "NAIROBI BOUND"
        bound_b_name = "NAKURU BOUND"
    elif "KANYONYO" in station_upper:
        bound_a_name = "MWINGI BOUND"
        bound_b_name = "THIKA BOUND"
    elif "ISINYA" in station_upper:
        bound_a_name = "KAJIADO BOUND"
        bound_b_name = "NAIROBI BOUND"
    elif "SUSWA" in station_upper:
        bound_a_name = "NAROK BOUND"
        bound_b_name = "NAIROBI BOUND"

    response = []

    if static_a:
        response.append({
            "slot": "static_bound_a",
            "title": f"Static: {static_a.weighbridge_name or station_upper} - {static_a.bound or bound_a_name}",
            "exists": True,
            "report_id": static_a.report_id,
            "text": build_static_sms_summary(static_a)
        })
    else:
        response.append({
            "slot": "static_bound_a",
            "title": f"Static: {station_upper} - {bound_a_name}",
            "exists": False,
            "report_id": None,
            "text": f"DAILY REPORT\n{station_upper} {bound_a_name} WB\nDate: {date_formatted}\n\n[Awaiting report upload and processing]"
        })

    if static_b:
        response.append({
            "slot": "static_bound_b",
            "title": f"Static: {static_b.weighbridge_name or station_upper} - {static_b.bound or bound_b_name}",
            "exists": True,
            "report_id": static_b.report_id,
            "text": build_static_sms_summary(static_b)
        })
    else:
        response.append({
            "slot": "static_bound_b",
            "title": f"Static: {station_upper} - {bound_b_name}",
            "exists": False,
            "report_id": None,
            "text": f"DAILY REPORT\n{station_upper} {bound_b_name} WB\nDate: {date_formatted}\n\n[Awaiting report upload and processing]"
        })

    if mobile_1:
        response.append({
            "slot": "mobile_1",
            "title": f"Mobile: {mobile_1.station or station_upper} - TEAM ONE",
            "exists": True,
            "report_id": mobile_1.report_id,
            "text": build_mobile_sms_summary(mobile_1)
        })
    else:
        response.append({
            "slot": "mobile_1",
            "title": f"Mobile: {station_upper} - TEAM ONE",
            "exists": False,
            "report_id": None,
            "text": f"DAILY REPORT\n{station_upper} W/B DAILY MOBILE REPORT_TEAM ONE\nDate: {date_formatted}\n\n[Awaiting report upload and processing]"
        })

    if mobile_2:
        response.append({
            "slot": "mobile_2",
            "title": f"Mobile: {mobile_2.station or station_upper} - TEAM TWO",
            "exists": True,
            "report_id": mobile_2.report_id,
            "text": build_mobile_sms_summary(mobile_2)
        })
    else:
        response.append({
            "slot": "mobile_2",
            "title": f"Mobile: {station_upper} - TEAM TWO",
            "exists": False,
            "report_id": None,
            "text": f"DAILY REPORT\n{station_upper} W/B DAILY MOBILE REPORT_TEAM TWO\nDate: {date_formatted}\n\n[Awaiting report upload and processing]"
        })

    return response
