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


def serialize_session(session: ReportSession) -> dict:
    excel_report_ready = (
        "daily_hour" in session.dataframes
        and session.sections.get("daily_hour", {}).get("status") == "ready"
    )
    mobile_excel_report_ready = (
        "mobile_report" in session.dataframes
        and session.sections.get("mobile_report", {}).get("status") == "ready"
    )
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

    return {
        "report_id": session.report_id,
        "cards": [
            summary_card(
                "Total Weighed",
                daily_hour_total_column(session, "X"),
                "daily_hour.totals.X",
            ),
            summary_card(
                "Total Overloaded",
                daily_hour_total_column(session, "Y"),
                "daily_hour.totals.Y",
            ),
            summary_card(
                "Special Released",
                daily_hour_total_column(session, "G"),
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


@router.get("/report-sessions/{report_id}")
async def get_report_session(report_id: str):
    return serialize_session(require_session(report_id))


@router.get("/report-sessions/{report_id}/summary-cards")
async def get_report_session_summary_cards(report_id: str):
    return build_summary_cards(require_session(report_id))


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

    filename = (
        f"{session.station}_{session.bound}_{session.report_date}_daily_report.docx"
        .lower()
        .replace(" ", "_")
    )

    return Response(
        content=session.final_report,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
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

    filename = (
        f"{session.station}_{session.bound}_{session.report_date}_daily_report.xlsx"
        .lower()
        .replace(" ", "_")
    )

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

    filename = (
        f"{session.station}_{session.bound}_{session.report_date}_mobile_report.xlsx"
        .lower()
        .replace(" ", "_")
    )

    return Response(
        content=file_stream.getvalue(),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": f'attachment; filename="{filename}"',
            "X-Content-Type-Options": "nosniff",
        },
    )
