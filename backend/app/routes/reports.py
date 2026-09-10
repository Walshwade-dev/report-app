import logging
import os
import secrets
from datetime import datetime
from typing import Any

import pandas as pd

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
from app.services.daily_hour_processor import HOURS
from app.services.excel_report_builder import build_excel_report
from app.services.final_report_builder import build_final_report
from app.services.mobile_excel_report_builder import (
    build_mobile_excel_report,
    _manual_shifts,
    _shift_cutoff_hour,
    _manual_value,
)
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
            if role in ["admin", "developer", "viewer", "user"] or username == "admin":
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

    payload = {
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

    if "mobile_report_raw" in session.dataframes:
        raw_df = session.dataframes["mobile_report_raw"]
        mobile_report_inputs = session.manual_inputs.get("mobile_report") or {}
        reweigh_tickets = mobile_report_inputs.get("reweigh_tickets") or []
        dimension_charges = mobile_report_inputs.get("dimension_charges") or []
        payload["mobile_report"] = mobile_report_response(
            raw_df,
            reweigh_tickets=reweigh_tickets,
            dimension_charges=dimension_charges,
            station=session.weighbridge_name or session.station,
        )

    return payload


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


import time

_SESSIONS_CACHE: tuple[list[tuple[ReportSession, float]], float] | None = None
_CACHE_TTL_SECONDS = 60.0


def invalidate_sessions_cache():
    global _SESSIONS_CACHE
    _SESSIONS_CACHE = None


def available_report_sessions(force_refresh: bool = False) -> list[tuple[ReportSession, float]]:
    global _SESSIONS_CACHE
    now = time.time()
    if not force_refresh and _SESSIONS_CACHE is not None:
        cached_data, timestamp = _SESSIONS_CACHE
        if now - timestamp < _CACHE_TTL_SECONDS:
            return cached_data

    sessions = report_session_store.list_all_sessions()
    _SESSIONS_CACHE = (sessions, now)
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
    invalidate_sessions_cache()
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
    if authorization and authorization.startswith("Bearer "):
        token = authorization.split(" ")[1]
        payload = decode_access_token(token)
        if payload:
            role = payload.get("role")
            if role not in ["admin", "developer"]:
                raise HTTPException(
                    status_code=403,
                    detail="Admin or Developer privileges required to delete sessions.",
                )
    
    require_admin_password(x_admin_password, authorization)
    deleted = report_session_store.delete(report_id)
    invalidate_sessions_cache()

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


def parse_danka_team(staff_value: Any) -> dict | None:
    if staff_value is None:
        return None
    cleaned_str = str(staff_value).strip()
    if not cleaned_str:
        return None
    names = []
    for part in cleaned_str.replace("\\", "/").split("/"):
        cleaned = " ".join(part.strip().split())
        if cleaned:
            names.append(cleaned.upper())
    if not names:
        return None

    dm_name = next((name for name in names if "DM" in name.split()), names[0])
    drivers = [name for name in names if name != dm_name]

    return {
        "dm": dm_name,
        "drivers": drivers,
        "team": " / ".join([dm_name, *drivers]),
    }


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
    return parse_danka_team(mobile_report_manual_inputs(session).get("danka_staff"))


def _shift_index_for_record(record: Any, shifts: list[dict[str, Any]]) -> int:
    if len(shifts) < 2:
        return 0
    date_time = None
    if isinstance(record, dict):
        date_time = record.get("date_time")
    elif hasattr(record, "get"):
        date_time = record.get("date_time")
    elif hasattr(record, "__getitem__"):
        try:
            date_time = record["date_time"]
        except Exception:
            date_time = None

    parsed_dt = pd.to_datetime(date_time, errors="coerce", dayfirst=True)
    if pd.isna(parsed_dt):
        return 0

    cutoff = _shift_cutoff_hour(shifts)
    return 0 if parsed_dt.hour < cutoff else 1


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
        "axleConfigs": {},
        "psvBreakdown": {
            "charged": 0,
            "withinAllowed": 0,
            "redistributed": 0,
            "specialRelease": 0,
        },
    }


def add_static_kpis(target: dict, session: ReportSession) -> None:
    y = daily_hour_total_column(session, "Y") or 0
    g = daily_hour_total_column(session, "G") or 0
    z = daily_hour_total_column(session, "Z") or 0
    r = daily_hour_total_column(session, "R") or 0

    # Extract overloaded dataframe or fallback to section metadata
    overloaded_df = session.dataframes.get("overloaded")
    if overloaded_df is None:
        try:
            pkl_path = report_session_store._processed_section_path(session.report_id, "overloaded")
            if pkl_path.exists():
                overloaded_df = pd.read_pickle(pkl_path)
                session.dataframes["overloaded"] = overloaded_df
        except Exception:
            pass

    psv_count = 0
    psv_breakdown = {
        "charged": 0,
        "withinAllowed": 0,
        "redistributed": 0,
        "specialRelease": 0,
    }
    session_axle_counts: dict[str, int] = {}

    if isinstance(overloaded_df, pd.DataFrame) and not overloaded_df.empty:
        # 1. PSV coaches: Use Cargo column matching 'PSV' or 'Passengers' words
        if "Cargo" in overloaded_df.columns:
            cargo_series = overloaded_df["Cargo"].fillna("").astype(str)
            psv_mask = cargo_series.str.contains(r"(?i)\b(?:psv|passengers?)\b", regex=True)
            psv_rows = overloaded_df[psv_mask]
            psv_count = len(psv_rows)

            # Rely on LastState if present for PSV state, else fallback to state.
            # PSV vehicles are given an additional allowance of 2000KG above allocated GVW.
            # If GVW overload sits within 2000kg (<= 2000kg), they are marked as within allowed GVW.
            for _, row in psv_rows.iterrows():
                last_st = str(row.get("LastState", "") or "").strip()
                st = str(row.get("state", "") or "").strip()
                eff = last_st if last_st and last_st.lower() not in ("nan", "none", "null") else st
                eff_lower = eff.lower()

                last_gvw = row.get("LastGVWOverload")
                gvw = row.get("GVWOverload")
                val = last_gvw if pd.notna(last_gvw) and str(last_gvw).strip() != "" else gvw
                try:
                    gvw_ov = float(str(val).replace(",", "").strip())
                except Exception:
                    gvw_ov = 0.0

                if "charg" in eff_lower:
                    if gvw_ov <= 2000.0:
                        psv_breakdown["withinAllowed"] += 1
                    else:
                        psv_breakdown["charged"] += 1
                elif "redistribut" in eff_lower:
                    psv_breakdown["redistributed"] += 1
                elif "releas" in eff_lower:
                    psv_breakdown["specialRelease"] += 1

        # 2. Axle configurations: Count vehicles by AxleConfig
        if "AxleConfig" in overloaded_df.columns:
            axle_series = overloaded_df["AxleConfig"].dropna().astype(str).str.strip()
            for cfg, cnt in axle_series.value_counts().items():
                cfg_clean = str(cfg).strip()
                if cfg_clean and cfg_clean.lower() not in ("nan", "none", "null"):
                    session_axle_counts[cfg_clean] = int(cnt)
    else:
        # Fallback to session section extra metadata if available
        ov_sec = session.sections.get("overloaded", {})
        psv_count = int(ov_sec.get("psv_count", 0) or 0)
        sec_breakdown = ov_sec.get("psv_breakdown")
        if isinstance(sec_breakdown, dict):
            for k in ("charged", "withinAllowed", "redistributed", "specialRelease"):
                psv_breakdown[k] = int(sec_breakdown.get(k, 0) or 0)
        sec_axles = ov_sec.get("axle_configs")
        if isinstance(sec_axles, dict):
            for cfg, cnt in sec_axles.items():
                session_axle_counts[str(cfg)] = int(cnt)

        # If still 0 and no overloaded file, fallback to traffic census if available
        if psv_count == 0 and not ov_sec.get("filename"):
            tc = session.manual_inputs.get("traffic_census") or session.sections.get("traffic_census", {}).get("values") or {}
            if isinstance(tc, dict):
                try:
                    psv_count = int(tc.get("buses_gte_3500kg", 0) or 0)
                except Exception:
                    psv_count = 0

    target["weighed"] += daily_hour_total_column(session, "X") or 0
    target["overloads"] += max(y - g, 0)
    target["psvOverloads"] += psv_count
    target["minGross"] += g
    target["charged"] += z
    target["redistributed"] += r
    target["reportsGenerated"] += 1
    target["chargedRedist"] = f"{target['charged']} / {target['redistributed']}"

    if "psvBreakdown" not in target:
        target["psvBreakdown"] = {
            "charged": 0,
            "withinAllowed": 0,
            "redistributed": 0,
            "specialRelease": 0,
        }
    for k in ("charged", "withinAllowed", "redistributed", "specialRelease"):
        target["psvBreakdown"][k] = target["psvBreakdown"].get(k, 0) + psv_breakdown.get(k, 0)

    if "axleConfigs" not in target:
        target["axleConfigs"] = {}
    for cfg, count in session_axle_counts.items():
        target["axleConfigs"][cfg] = target["axleConfigs"].get(cfg, 0) + count


@router.get("/report-sessions/analytics/dashboard")
async def get_analytics_dashboard(
    static_date: str | None = None,
    mobile_date: str | None = None,
    mobile_bound: str | None = None,
    station: str | None = None,
):
    sessions = []
    session_modified_at: dict[str, float] = {}
    for session, modified_at in available_report_sessions():
        sessions.append(session)
        session_modified_at[session.report_id] = modified_at

    target_station = classify_station(station) or (station.strip() if station else None)
    target_station_norm = target_station.lower() if target_station else None

    latest_static_sessions: dict[tuple[str, str, str], tuple[ReportSession, float]] = {}
    latest_mobile_sessions: dict[tuple[str, str, str], tuple[ReportSession, float]] = {}

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
            st_code = classify_station(s.station or s.weighbridge_name)
            s_station = st_code or (s.station or s.weighbridge_name or "").strip()
            slot = mobile_report_slot(s.bound)
            key = (s.report_date, s_station.lower(), slot)
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

    all_static_dates = sorted(
        {report_date for report_date, _, _ in latest_static_sessions},
        reverse=True,
    )
    if target_station_norm:
        station_static_dates = sorted(
            {
                report_date
                for report_date, st_key, _ in latest_static_sessions
                if st_key.lower() == target_station_norm or (classify_station(st_key) and classify_station(st_key).lower() == target_station_norm)
            },
            reverse=True,
        )
        static_dates = station_static_dates
    else:
        static_dates = all_static_dates

    selected_static_date = (
        static_date if static_date in static_dates else (static_dates[0] if static_dates else None)
    )

    static_by_bound = {
        "boundA": {**empty_static_kpis(), "label": "Bound A"},
        "boundB": {**empty_static_kpis(), "label": "Bound B"},
        "total": {**empty_static_kpis(), "label": "Total"},
    }

    # Cross-station comparison stats (Traffic comparison between stations, court cases, compliance rates)
    charts_static_date = selected_static_date or (all_static_dates[0] if all_static_dates else None)
    for (report_date, _, bound_key), (s, _) in latest_static_sessions.items():
        if charts_static_date and report_date == charts_static_date:
            code = classify_station(s.station or s.weighbridge_name)
            if code and code in stations_data:
                x = daily_hour_total_column(s, "X") or 0
                y = daily_hour_total_column(s, "Y") or 0
                g = daily_hour_total_column(s, "G") or 0
                called = daily_hour_total_column(s, "C") or 0
                cases = s.manual_inputs.get("cases_cleared_in_court", 0) or 0
                overload_no_permit = max(y - g, 0)
                compliant = max(called - overload_no_permit, 0)

                stations_data[code]["traffic"][bound_key] += x
                stations_data[code]["cases"][bound_key] += cases
                stations_data[code]["compliance"][bound_key]["calledIn"] += called
                stations_data[code]["compliance"][bound_key]["weighed"] += x
                stations_data[code]["compliance"][bound_key]["compliant"] += compliant

        # Station-indigenous static KPIs
        if selected_static_date and report_date == selected_static_date:
            session_station = classify_station(s.station or s.weighbridge_name) or (s.station or s.weighbridge_name or "").strip()
            if target_station_norm and session_station.lower() != target_station_norm:
                continue
            if s.bound:
                static_by_bound[bound_key]["label"] = s.bound
            add_static_kpis(static_by_bound[bound_key], s)
            add_static_kpis(static_by_bound["total"], s)

    # Filter mobile sessions to target station if provided
    filtered_mobile: list[tuple[str, str, str, ReportSession, float]] = []
    for (report_date, s_st, slot), (session, modified_at) in latest_mobile_sessions.items():
        if target_station_norm and s_st != target_station_norm:
            continue
        filtered_mobile.append((report_date, s_st, slot, session, modified_at))

    mobile_reports = [
        {
            "date": report_date,
            "bound": slot,
            "bound_label": mobile_report_label(slot),
            "label": f"{report_date} - {mobile_report_label(slot)}",
            "report_id": session.report_id,
            "updated_at": modified_at,
        }
        for report_date, _, slot, session, modified_at in filtered_mobile
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
        for report_date, _, slot, session, _ in filtered_mobile:
            if report_date == selected_mobile["date"] and slot == selected_mobile["bound"]:
                selected_session = session
                break

    mobile_summary = (
        selected_session.sections["mobile_report"].get("summary", {})
        if selected_session is not None
        else {}
    )

    shift_a_stats = {"weighed": 0, "warned": 0, "legal": 0, "charged": 0}
    shift_b_stats = {"weighed": 0, "warned": 0, "legal": 0, "charged": 0}

    if selected_session is not None:
        shifts = _manual_shifts(selected_session)
        df_mobile = selected_session.dataframes.get("mobile_report")
        if df_mobile is None:
            try:
                pkl_path = report_session_store._processed_section_path(selected_session.report_id, "mobile_report")
                if pkl_path.exists():
                    df_mobile = pd.read_pickle(pkl_path)
                    selected_session.dataframes["mobile_report"] = df_mobile
            except Exception:
                pass

        if isinstance(df_mobile, pd.DataFrame) and not df_mobile.empty:
            for _, r in df_mobile.iterrows():
                s_idx = _shift_index_for_record(r, shifts)
                target_shift = shift_a_stats if s_idx == 0 else shift_b_stats
                if r.get("is_weighed", True):
                    target_shift["weighed"] += 1
                rem = str(r.get("remarks", "")).strip().upper()
                if "CHARG" in rem:
                    target_shift["charged"] += 1
                elif "WARN" in rem:
                    target_shift["warned"] += 1
                elif "LEGAL" in rem:
                    target_shift["legal"] += 1

            # Dimension charges attribution if any
            raw_dim_charges = selected_session.manual_inputs.get("mobile_report", {}).get("dimension_charges", [])
            for dc in raw_dim_charges:
                if isinstance(dc, dict):
                    dc_idx = _shift_index_for_record(dc, shifts)
                    target_shift = shift_a_stats if dc_idx == 0 else shift_b_stats
                    target_shift["charged"] += 1
                    target_shift["weighed"] += 1
        else:
            total_weighed = mobile_summary.get("total_trucks_weighed", 0)
            total_warned = mobile_summary.get("warned_trucks", 0)
            total_charged = mobile_summary.get("charged_trucks", 0)
            total_legal = max(total_weighed - total_warned - total_charged, 0)
            shift_a_stats = {
                "weighed": total_weighed,
                "warned": total_warned,
                "legal": total_legal,
                "charged": total_charged,
            }

    total_mobile_weighed = shift_a_stats["weighed"] + shift_b_stats["weighed"]
    total_mobile_warned = shift_a_stats["warned"] + shift_b_stats["warned"]
    total_mobile_legal = shift_a_stats["legal"] + shift_b_stats["legal"]
    total_mobile_charged = shift_a_stats["charged"] + shift_b_stats["charged"]

    if total_mobile_weighed == 0 and mobile_summary.get("total_trucks_weighed", 0) > 0:
        total_mobile_weighed = mobile_summary.get("total_trucks_weighed", 0)
        total_mobile_warned = mobile_summary.get("warned_trucks", 0)
        total_mobile_charged = mobile_summary.get("charged_trucks", 0)
        total_mobile_legal = max(total_mobile_weighed - total_mobile_warned - total_mobile_charged, 0)
        shift_a_stats = {
            "weighed": total_mobile_weighed,
            "warned": total_mobile_warned,
            "legal": total_mobile_legal,
            "charged": total_mobile_charged,
        }

    return {
        "static": {
            "weighed": static_by_bound["total"]["weighed"],
            "overloads": static_by_bound["total"]["overloads"],
            "psvOverloads": static_by_bound["total"]["psvOverloads"],
            "minGross": static_by_bound["total"]["minGross"],
            "chargedRedist": static_by_bound["total"]["chargedRedist"],
            "reportsGenerated": static_by_bound["total"]["reportsGenerated"],
            "axleConfigs": static_by_bound["total"].get("axleConfigs", {}),
            "psvBreakdown": static_by_bound["total"].get("psvBreakdown", {}),
            "dates": static_dates,
            "selectedDate": selected_static_date,
            "byBound": static_by_bound,
        },
        "mobile": {
            "weighed": total_mobile_weighed,
            "warned": total_mobile_warned,
            "legal": total_mobile_legal,
            "charged": total_mobile_charged,
            "shifts": {
                "shiftA": {
                    "label": "Shift A (Day Shift)",
                    **shift_a_stats,
                },
                "shiftB": {
                    "label": "Shift B (Night Shift)",
                    **shift_b_stats,
                },
                "total": {
                    "weighed": total_mobile_weighed,
                    "warned": total_mobile_warned,
                    "legal": total_mobile_legal,
                    "charged": total_mobile_charged,
                },
            },
            "reports": mobile_reports,
            "selected": selected_mobile,
        },
        "stations": list(stations_data.values())
    }


@router.get("/report-sessions/analytics/details")
async def get_analytics_details(station: str | None = None):
    sessions = [session for session, _ in available_report_sessions()]

    station_names = {
        "Juja": "Juja Weighbridge",
        "Kanyonyo": "Kanyonyo",
        "Athi River": "Athi River",
        "Gilgil": "Gilgil",
        "Isinya": "Isinya",
        "Suswa": "Suswa"
    }

    target_code = classify_station(station) or (station.strip() if station else "Juja")
    target_lower = target_code.lower()

    bound_a_name = "Bound A"
    bound_b_name = "Bound B"
    if "juja" in target_lower:
        bound_a_name = "Thika Bound"
        bound_b_name = "Nairobi Bound"
    elif "athi" in target_lower:
        bound_a_name = "Mombasa Bound"
        bound_b_name = "Nairobi Bound"
    elif "gilgil" in target_lower:
        bound_a_name = "Nairobi Bound"
        bound_b_name = "Nakuru Bound"
    elif "kanyonyo" in target_lower:
        bound_a_name = "Mwingi Bound"
        bound_b_name = "Thika Bound"
    elif "isinya" in target_lower:
        bound_a_name = "Kajiado Bound"
        bound_b_name = "Nairobi Bound"
    elif "suswa" in target_lower:
        bound_a_name = "Narok Bound"
        bound_b_name = "Nairobi Bound"

    station_sessions = [
        s for s in sessions 
        if (classify_station(s.station or s.weighbridge_name) or (s.station or s.weighbridge_name or "").strip()).lower() == target_lower
    ]
    
    station_bound_a_traffic = 0
    station_bound_b_traffic = 0
    station_bound_a_cases = 0
    station_bound_b_cases = 0
    station_total_called = 0
    station_total_compliant = 0
    station_overloads_intercepted = 0
    
    for s in station_sessions:
        is_a = is_bound_a(target_code, s.bound)
        
        weighed = daily_hour_total_column(s, "X") or 0
        if is_a:
            station_bound_a_traffic += weighed
        else:
            station_bound_b_traffic += weighed
            
        cases = s.manual_inputs.get("cases_cleared_in_court", 0) or 0
        if is_a:
            station_bound_a_cases += cases
        else:
            station_bound_b_cases += cases
            
        called = daily_hour_total_column(s, "C") or 0
        y = daily_hour_total_column(s, "Y") or 0
        g = daily_hour_total_column(s, "G") or 0
        overload_no_permit = max(y - g, 0)
        compliant = max(called - overload_no_permit, 0)
        
        station_total_called += called
        station_total_compliant += compliant
        station_overloads_intercepted += overload_no_permit

    station_total_traffic = station_bound_a_traffic + station_bound_b_traffic
    station_total_cases = station_bound_a_cases + station_bound_b_cases
    station_compliance_rate = (station_total_compliant / station_total_called * 100) if station_total_called > 0 else 0.0

    # Daily breakdown for station
    daily_traffic: dict[str, dict[str, int]] = {}
    daily_cases: dict[str, dict[str, int]] = {}
    
    for s in station_sessions:
        try:
            day_str = s.report_date.split("-")[2]  # DD
        except Exception:
            continue
            
        is_a = is_bound_a(target_code, s.bound)
        weighed = daily_hour_total_column(s, "X") or 0
        cases = s.manual_inputs.get("cases_cleared_in_court", 0) or 0
        
        if day_str not in daily_traffic:
            daily_traffic[day_str] = {"boundA": 0, "boundB": 0}
        if day_str not in daily_cases:
            daily_cases[day_str] = {"boundA": 0, "boundB": 0}
            
        if is_a:
            daily_traffic[day_str]["boundA"] += weighed
            daily_cases[day_str]["boundA"] += cases
        else:
            daily_traffic[day_str]["boundB"] += weighed
            daily_cases[day_str]["boundB"] += cases

    traffic_data = []
    for day in sorted(daily_traffic.keys()):
        traffic_data.append({
            "day": day,
            "thikaBound": daily_traffic[day]["boundA"],
            "nairobiBound": daily_traffic[day]["boundB"],
            "boundA": daily_traffic[day]["boundA"],
            "boundB": daily_traffic[day]["boundB"],
        })
        
    court_cases_data = []
    for day in sorted(daily_cases.keys()):
        court_cases_data.append({
            "day": day,
            "thikaBound": daily_cases[day]["boundA"],
            "nairobiBound": daily_cases[day]["boundB"],
            "boundA": daily_cases[day]["boundA"],
            "boundB": daily_cases[day]["boundB"],
        })

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
            "active": code.lower() == target_lower
        })

    return {
        "kpis": {
            "totalTraffic": station_total_traffic,
            "thikaTraffic": station_bound_a_traffic,
            "nairobiTraffic": station_bound_b_traffic,
            "boundATraffic": station_bound_a_traffic,
            "boundBTraffic": station_bound_b_traffic,
            "totalCourtCases": station_total_cases,
            "thikaCourtCases": station_bound_a_cases,
            "nairobiCourtCases": station_bound_b_cases,
            "boundACourtCases": station_bound_a_cases,
            "boundBCourtCases": station_bound_b_cases,
            "complianceRate": round(station_compliance_rate, 1),
            "overloadsIntercepted": station_overloads_intercepted
        },
        "station": target_code,
        "stationName": station_names.get(target_code, target_code),
        "boundALabel": bound_a_name,
        "boundBLabel": bound_b_name,
        "trafficData": traffic_data,
        "courtCasesData": court_cases_data,
        "crossStationData": cross_station_data
    }


@router.get("/report-sessions/analytics/dms-performance")
async def get_dms_performance(date: str | None = None, station: str | None = None):
    requested_date = date or datetime.now().strftime("%Y-%m-%d")
    try:
        filter_date = datetime.strptime(requested_date, "%Y-%m-%d")
    except Exception:
        filter_date = datetime.now()
        requested_date = filter_date.strftime("%Y-%m-%d")

    target_station = classify_station(station) or (station.strip() if station else None)
    target_station_norm = target_station.lower() if target_station else None

    all_ready_mobile: list[tuple[ReportSession, float]] = []
    for session, modified_at in available_report_sessions():
        try:
            if session and session.sections.get("mobile_report", {}).get("status") == "ready":
                if target_station_norm:
                    st_code = classify_station(session.station or session.weighbridge_name)
                    st_name = st_code or (session.station or session.weighbridge_name or "").strip()
                    if st_name.lower() != target_station_norm:
                        continue
                all_ready_mobile.append((session, modified_at))
        except Exception:
            pass

    available_mobile_dates = sorted(
        {s.report_date for s, _ in all_ready_mobile if s.report_date},
        reverse=True,
    )

    # Check if there are any mobile sessions matching the requested month and on/before requested_date
    has_month_sessions = any(
        s.report_date
        and datetime.strptime(s.report_date, "%Y-%m-%d").year == filter_date.year
        and datetime.strptime(s.report_date, "%Y-%m-%d").month == filter_date.month
        and datetime.strptime(s.report_date, "%Y-%m-%d") <= filter_date
        for s, _ in all_ready_mobile
        if s.report_date
    )

    effective_date = requested_date
    if not has_month_sessions and available_mobile_dates:
        # Fall back to the most recent month/date that actually has ready mobile report data
        effective_date = available_mobile_dates[0]
        try:
            filter_date = datetime.strptime(effective_date, "%Y-%m-%d")
        except Exception:
            pass

    latest_mobile_sessions: dict[tuple[str, str, str], tuple[ReportSession, float]] = {}

    for session, modified_at in all_ready_mobile:
        try:
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
        shifts = _manual_shifts(session)
        default_team = danka_staff_team(session)
        if not default_team and session.prepared_by:
            default_team = {
                "dm": session.prepared_by.strip().upper(),
                "drivers": [],
                "team": session.prepared_by.strip().upper(),
            }

        shift_teams = []
        for s_idx, shift in enumerate(shifts):
            staff_val = shift.get("danka_staff")
            if not staff_val and s_idx == 1:
                staff_val = _manual_value(
                    session,
                    "shift_two_danka_staff",
                    "shiftTwoDmEntry",
                    "shift_two_staff",
                    "shiftTwoStaff",
                )
            if not staff_val and s_idx == 0:
                staff_val = _manual_value(
                    session,
                    "danka_staff",
                    "dmEntry",
                    "computer_operator",
                    "computer_operators",
                )
            team = parse_danka_team(staff_val) or default_team
            shift_teams.append(team)

        if not any(shift_teams):
            continue

        report_count += 1

        is_current_month = False
        try:
            report_date = datetime.strptime(session.report_date, "%Y-%m-%d")
            is_current_month = report_date.year == filter_date.year and report_date.month == filter_date.month
        except Exception:
            pass

        records = session.dataframes.get("mobile_report")
        if records is None or getattr(records, "empty", True):
            processed_path = report_session_store._processed_section_path(session.report_id, "mobile_report")
            if processed_path.exists():
                try:
                    records = pd.read_pickle(processed_path)
                except Exception:
                    records = None
            if records is None or getattr(records, "empty", True):
                raw_path = report_session_store._processed_section_path(session.report_id, "mobile_report_raw")
                if raw_path.exists():
                    try:
                        raw_df = pd.read_pickle(raw_path)
                        mobile_inputs = session.manual_inputs.get("mobile_report") or {}
                        records = normalize_mobile_report(
                            raw_df,
                            reweigh_tickets=mobile_inputs.get("reweigh_tickets") or [],
                            dimension_charges=mobile_inputs.get("dimension_charges") or [],
                            station=session.weighbridge_name or session.station,
                        )
                    except Exception:
                        records = None

        shift_weighed = [0] * len(shifts)
        shift_charged = [0] * len(shifts)

        if records is not None and not records.empty:
            for _, record in records.iterrows():
                s_idx = _shift_index_for_record(record, shifts)
                if s_idx >= len(shifts):
                    s_idx = len(shifts) - 1

                is_weighed = bool(record.get("is_weighed")) or (pd.to_numeric(record.get("total_gvw_kg", 0), errors="coerce") or 0) > 0
                is_charged = (
                    bool(record.get("is_gvw_axle_charge"))
                    or bool(record.get("is_dimension_charge"))
                    or str(record.get("remarks", "")).strip().upper() == "CHARGED"
                )
                if is_weighed:
                    shift_weighed[s_idx] += 1
                if is_charged:
                    shift_charged[s_idx] += 1
        else:
            summary = session.sections.get("mobile_report", {}).get("summary", {})
            total_w = int(summary.get("total_trucks_weighed", 0) or 0)
            total_c = int(summary.get("charged_trucks", 0) or 0)
            if len(shifts) == 1 or (len(shift_teams) >= 2 and shift_teams[0] == shift_teams[1]):
                shift_weighed[0] = total_w
                shift_charged[0] = total_c
            else:
                hourly = summary.get("hourly_counts", {})
                cutoff = _shift_cutoff_hour(shifts)
                w_s1 = sum(int(hourly.get(h, 0) or 0) for h in HOURS[:cutoff])
                w_s2 = total_w - w_s1
                shift_weighed[0] = max(w_s1, 0)
                if len(shift_weighed) > 1:
                    shift_weighed[1] = max(w_s2, 0)

                mobile_inputs = session.manual_inputs.get("mobile_report") or {}
                dim_charges = mobile_inputs.get("dimension_charges") or []
                for dc in dim_charges:
                    dc_idx = _shift_index_for_record(dc, shifts)
                    if dc_idx < len(shift_charged):
                        shift_charged[dc_idx] += 1
                
                remaining_c = max(total_c - sum(shift_charged), 0)
                if remaining_c > 0:
                    if total_w > 0 and len(shift_charged) > 1:
                        c_s1 = int(round(remaining_c * (shift_weighed[0] / total_w)))
                        shift_charged[0] += c_s1
                        shift_charged[1] += (remaining_c - c_s1)
                    else:
                        shift_charged[0] += remaining_c

        dms_in_session = set()
        for idx, team in enumerate(shift_teams):
            if not team:
                continue
            dm_name = team["dm"]
            row = stats.setdefault(
                dm_name,
                {
                    "name": dm_name,
                    "surname": dm_name.split()[-1] if dm_name else "",
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
            row["weighed"] += shift_weighed[idx] if idx < len(shift_weighed) else 0
            row["charged"] += shift_charged[idx] if idx < len(shift_charged) else 0
            if is_current_month:
                row["monthCharged"] += shift_charged[idx] if idx < len(shift_charged) else 0
            dms_in_session.add(dm_name)

        for dm_name in dms_in_session:
            stats[dm_name]["reports"] += 1

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
        "selectedDate": effective_date,
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

        psv_count = 0
        psv_breakdown = {
            "charged": 0,
            "withinAllowed": 0,
            "redistributed": 0,
            "specialRelease": 0,
        }
        axle_configs = {}
        if "Cargo" in raw_df.columns:
            cargo_s = raw_df["Cargo"].fillna("").astype(str)
            psv_mask = cargo_s.str.contains(r"(?i)\b(?:psv|passengers?)\b", regex=True)
            psv_df = raw_df[psv_mask]
            psv_count = len(psv_df)
            for _, r in psv_df.iterrows():
                last_st = str(r.get("LastState", "") or "").strip()
                st = str(r.get("state", "") or "").strip()
                eff = last_st if last_st and last_st.lower() not in ("nan", "none", "null") else st
                eff_l = eff.lower()

                last_gvw = r.get("LastGVWOverload")
                gvw = r.get("GVWOverload")
                val = last_gvw if pd.notna(last_gvw) and str(last_gvw).strip() != "" else gvw
                try:
                    gvw_ov = float(str(val).replace(",", "").strip())
                except Exception:
                    gvw_ov = 0.0

                if "charg" in eff_l:
                    if gvw_ov <= 2000.0:
                        psv_breakdown["withinAllowed"] += 1
                    else:
                        psv_breakdown["charged"] += 1
                elif "redistribut" in eff_l:
                    psv_breakdown["redistributed"] += 1
                elif "releas" in eff_l:
                    psv_breakdown["specialRelease"] += 1

        if "AxleConfig" in raw_df.columns:
            axle_s = raw_df["AxleConfig"].dropna().astype(str).str.strip()
            for cfg, cnt in axle_s.value_counts().items():
                cfg_c = str(cfg).strip()
                if cfg_c and cfg_c.lower() not in ("nan", "none", "null"):
                    axle_configs[cfg_c] = int(cnt)

        updated = report_session_store.set_section_ready(
            report_id,
            "overloaded",
            raw_df,
            filename=filename,
            extra={
                "valid_permit_count": valid_permit_count,
                "psv_count": psv_count,
                "psv_breakdown": psv_breakdown,
                "axle_configs": axle_configs,
            },
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
    session = require_session(report_id)

    try:
        filename, content, raw_df = await read_upload_dataframe(file)
        report_session_store.save_upload(
            report_id,
            "mobile_report",
            filename,
            content,
        )
        
        # Store raw_df pickle in session
        raw_df_path = report_session_store._processed_section_path(report_id, "mobile_report_raw")
        raw_df_path.parent.mkdir(parents=True, exist_ok=True)
        raw_df.to_pickle(raw_df_path)
        session.dataframes["mobile_report_raw"] = raw_df

        # Check existing manual inputs
        mobile_report_inputs = session.manual_inputs.get("mobile_report") or {}
        reweigh_tickets = mobile_report_inputs.get("reweigh_tickets") or []
        dimension_charges = mobile_report_inputs.get("dimension_charges") or []

        records = normalize_mobile_report(
            raw_df,
            reweigh_tickets=reweigh_tickets,
            dimension_charges=dimension_charges,
            station=session.weighbridge_name or session.station,
        )
        summary = summarize_mobile_report(records)

        updated = report_session_store.set_section_ready(
            report_id,
            "mobile_report",
            records,
            filename=filename,
            extra={"summary": summary},
        )
        payload = serialize_session(updated)
        payload["mobile_report"] = mobile_report_response(
            raw_df,
            reweigh_tickets=reweigh_tickets,
            dimension_charges=dimension_charges,
            station=session.weighbridge_name or session.station,
        )
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
async def get_sms_summary_dates(station: str | None = None):
    target_station = classify_station(station) or (station.strip() if station else None)
    target_station_norm = target_station.lower() if target_station else None

    dates = set()
    for session, _ in available_report_sessions():
        try:
            if target_station_norm:
                st_code = classify_station(session.station or session.weighbridge_name)
                st_name = st_code or (session.station or session.weighbridge_name or "").strip()
                if st_name.lower() != target_station_norm:
                    continue

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
