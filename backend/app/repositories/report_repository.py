import logging
import os
from collections.abc import Iterator
from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from uuid import UUID as PyUUID

from sqlalchemy import String, cast, func, or_, select
from sqlalchemy.dialects.postgresql import insert
from sqlalchemy.exc import SQLAlchemyError
from sqlalchemy.orm import Session, selectinload

from app.core.database import SessionLocal
from app.db.models import (
    Report,
    ReportManualInput,
    ReportOutput,
    ReportPreview,
    ReportUpload,
)


logger = logging.getLogger(__name__)

SECTION_NUMBERS = {
    "daily-hour-statistics": 1,
    "daily-hour-chart": 2,
    "traffic-census": 3,
    "daily_summary": 4,
    "transgressions": 5,
    "impounded-prohibited": 6,
    "impounded_prohibited": 6,
    "wideload": 7,
}

STATIC_REQUIRED_UPLOADS = {
    "daily_hour",
    "wideload",
    "impounded_prohibited",
    "overloaded",
}


def _as_uuid(report_id: str | PyUUID) -> PyUUID:
    return report_id if isinstance(report_id, PyUUID) else PyUUID(str(report_id))


def _utc_now() -> datetime:
    return datetime.now(timezone.utc)


def _int_value(value: Any) -> int:
    try:
        if value is None or value == "":
            return 0
        return int(float(str(value).replace(",", "")))
    except (TypeError, ValueError):
        return 0


def _manual_mobile_payload(manual_inputs: dict[str, Any]) -> dict[str, Any]:
    mobile_inputs = manual_inputs.get("mobile_report")
    if isinstance(mobile_inputs, dict):
        return mobile_inputs

    extra = manual_inputs.get("extra")
    if isinstance(extra, dict) and isinstance(extra.get("mobile_report"), dict):
        return extra["mobile_report"]

    return {}


def _report_type_from_payload(payload: dict[str, Any]) -> str:
    sections = payload.get("sections") or {}
    station = str(payload.get("station") or payload.get("weighbridge_name") or "")
    bound = str(payload.get("bound") or "")

    if (
        sections.get("mobile_report", {}).get("status") == "ready"
        or "mobile" in station.lower()
        or "mobile" in bound.lower()
    ):
        return "mobile_weighbridge"

    return "static_weighbridge"


def _report_title(report: Report) -> str:
    parts = [
        report.weighbridge_name or report.station,
        report.bound_name,
        report.report_date,
    ]
    title = " ".join(str(part).strip() for part in parts if part)
    return report.title or title or str(report.id)


class ReportRepository:
    def __init__(self):
        self.enabled = SessionLocal is not None

    @contextmanager
    def _session_scope(self) -> Iterator[Session]:
        if SessionLocal is None:
            raise RuntimeError("DATABASE_URL is not configured.")

        session = SessionLocal()
        try:
            yield session
            session.commit()
        except Exception:
            session.rollback()
            raise
        finally:
            session.close()

    def _run(self, operation_name: str, operation):
        if not self.enabled:
            return None

        try:
            with self._session_scope() as session:
                return operation(session)
        except SQLAlchemyError:
            logger.exception("Database operation failed: %s", operation_name)
            return None

    def save_session_snapshot(self, payload: dict[str, Any]) -> None:
        def operation(session: Session) -> None:
            report_id = _as_uuid(payload["report_id"])
            final_status = payload.get("final_report_status") or "not_built"
            status = {
                "ready": "completed",
                "error": "failed",
                "not_built": "draft",
            }.get(final_status, final_status)
            completed_at = _utc_now() if status == "completed" else None

            existing = session.get(Report, report_id)
            if existing is None:
                existing = Report(id=report_id)
                session.add(existing)

            report_type = _report_type_from_payload(payload)
            existing.report_date = payload.get("report_date")
            existing.station = payload.get("station")
            existing.weighbridge_name = payload.get("weighbridge_name")
            existing.bound_name = payload.get("bound")
            existing.report_type = report_type
            existing.title = " ".join(
                str(part).strip()
                for part in [
                    payload.get("weighbridge_name") or payload.get("station"),
                    payload.get("bound"),
                    payload.get("report_date"),
                ]
                if part
            ) or None
            existing.status = status
            existing.state_payload = payload
            existing.completed_at = completed_at
            existing.error_message = payload.get("final_report_error")

        self._run("save_session_snapshot", operation)

    def load_session_snapshot(self, report_id: str) -> dict[str, Any] | None:
        def operation(session: Session) -> dict[str, Any] | None:
            report = session.get(Report, _as_uuid(report_id))
            return report.state_payload if report else None

        return self._run("load_session_snapshot", operation)

    def list_report_ids(self) -> list[str]:
        def operation(session: Session) -> list[str]:
            rows = session.execute(
                select(Report.id).order_by(Report.updated_at.desc())
            ).all()
            return [str(row[0]) for row in rows]

        return self._run("list_report_ids", operation) or []

    def _filtered_reports_query(
        self,
        status: str | None = None,
        report_type: str | None = None,
        search: str | None = None,
    ):
        statement = select(Report)

        if status:
            statement = statement.where(Report.status == status)

        if report_type:
            statement = statement.where(Report.report_type == report_type)

        if search:
            pattern = f"%{search.strip()}%"
            statement = statement.where(
                or_(
                    cast(Report.id, String).ilike(pattern),
                    Report.title.ilike(pattern),
                    Report.weighbridge_name.ilike(pattern),
                    Report.station.ilike(pattern),
                    Report.bound_name.ilike(pattern),
                )
            )

        return statement

    def _summarize_report(self, report: Report) -> dict[str, Any]:
        upload_types = {upload.upload_type for upload in report.uploads}
        upload_count = len(upload_types)
        required_uploads_completed = STATIC_REQUIRED_UPLOADS.issubset(upload_types)
        manual_inputs_completed = report.manual_input is not None
        has_final_report = report.output is not None and bool(report.output.final_docx_path)
        download_available = False

        if (
            report.status == "completed"
            and report.output is not None
            and report.output.final_docx_path
        ):
            download_available = os.path.exists(report.output.final_docx_path)

        return {
            "report_id": str(report.id),
            "title": _report_title(report),
            "report_type": report.report_type,
            "weighbridge_name": report.weighbridge_name or report.station,
            "bound_name": report.bound_name,
            "status": report.status,
            "created_at": report.created_at,
            "updated_at": report.updated_at,
            "completed_at": report.completed_at,
            "has_final_report": has_final_report,
            "upload_count": upload_count,
            "required_uploads_completed": required_uploads_completed,
            "manual_inputs_completed": manual_inputs_completed,
            "download_available": download_available,
        }

    def list_reports(
        self,
        status: str | None = None,
        report_type: str | None = None,
        limit: int = 20,
        offset: int = 0,
        search: str | None = None,
    ) -> list[dict[str, Any]]:
        def operation(session: Session) -> list[dict[str, Any]]:
            statement = (
                self._filtered_reports_query(status, report_type, search)
                .options(
                    selectinload(Report.uploads),
                    selectinload(Report.manual_input),
                    selectinload(Report.output),
                )
                .order_by(Report.created_at.desc())
                .limit(limit)
                .offset(offset)
            )
            reports = session.execute(statement).scalars().all()
            return [self._summarize_report(report) for report in reports]

        return self._run("list_reports", operation) or []

    def count_reports(
        self,
        status: str | None = None,
        report_type: str | None = None,
        search: str | None = None,
    ) -> int:
        def operation(session: Session) -> int:
            filtered = self._filtered_reports_query(status, report_type, search).subquery()
            return int(session.execute(select(func.count()).select_from(filtered)).scalar_one())

        return self._run("count_reports", operation) or 0

    def get_report_summary(self, report_id: str) -> dict[str, Any] | None:
        def operation(session: Session) -> dict[str, Any] | None:
            statement = (
                select(Report)
                .where(Report.id == _as_uuid(report_id))
                .options(
                    selectinload(Report.uploads),
                    selectinload(Report.manual_input),
                    selectinload(Report.output),
                )
            )
            report = session.execute(statement).scalar_one_or_none()
            return self._summarize_report(report) if report else None

        return self._run("get_report_summary", operation)

    def update_report_status(
        self,
        report_id: str,
        status: str,
        error_message: str | None = None,
        completed: bool = False,
    ) -> None:
        def operation(session: Session) -> None:
            report = session.get(Report, _as_uuid(report_id))
            if report is None:
                return

            report.status = status
            report.error_message = error_message
            report.completed_at = _utc_now() if completed else None

        self._run("update_report_status", operation)

    def upsert_upload_metadata(
        self,
        report_id: str,
        upload_type: str,
        original_filename: str | None,
        file_path: Path,
        file_size_bytes: int | None = None,
        mime_type: str | None = None,
    ) -> None:
        def operation(session: Session) -> None:
            values = {
                "report_id": _as_uuid(report_id),
                "upload_type": upload_type,
                "original_filename": original_filename,
                "stored_filename": file_path.name,
                "file_path": str(file_path),
                "file_size_bytes": file_size_bytes,
                "mime_type": mime_type,
                "uploaded_at": _utc_now(),
            }
            statement = insert(ReportUpload).values(**values)
            statement = statement.on_conflict_do_update(
                constraint="uq_report_uploads_report_type",
                set_=values,
            )
            session.execute(statement)

        self._run("upsert_upload_metadata", operation)

    def upsert_manual_inputs(
        self,
        report_id: str,
        manual_inputs: dict[str, Any],
        prepared_by: str | None = None,
        approved_by: str | None = None,
        weighbridge_name: str | None = None,
        bound_name: str | None = None,
    ) -> None:
        traffic = manual_inputs.get("traffic_census") or {}
        mobile = _manual_mobile_payload(manual_inputs)
        cases = manual_inputs.get("cases_cleared_in_court")
        if cases is None:
            cases = mobile.get("cases_cleared_in_court")

        transgressions = manual_inputs.get("transgressions_count")
        if transgressions is None:
            transgressions = mobile.get("transgressions_count")

        def operation(session: Session) -> None:
            report_uuid = _as_uuid(report_id)
            existing = session.execute(
                select(ReportManualInput).where(
                    ReportManualInput.report_id == report_uuid
                )
            ).scalar_one_or_none()

            if existing is None:
                existing = ReportManualInput(report_id=report_uuid)
                session.add(existing)

            existing.prepared_by = prepared_by
            existing.approved_by = approved_by
            existing.weighbridge_name = weighbridge_name
            existing.bound_name = bound_name
            existing.cases_cleared_in_court = _int_value(cases)
            existing.total_transgressions = _int_value(transgressions)
            existing.buses_3500kg = _int_value(traffic.get("buses_gte_3500kg"))
            existing.vehicles_3500_to_7000kg = _int_value(
                traffic.get("vehicles_3500_to_7000_excluding_buses")
            )
            existing.vehicles_above_7000kg = _int_value(
                traffic.get("vehicles_gte_7000_excluding_buses")
            )
            existing.traffic_total = _int_value(traffic.get("total_traffic_census"))
            existing.payload = manual_inputs

        self._run("upsert_manual_inputs", operation)

    def save_preview_metadata(
        self,
        report_id: str,
        section_name: str,
        preview_format: str,
        file_path: Path,
        page: int | None = None,
    ) -> None:
        def operation(session: Session) -> None:
            values = {
                "report_id": _as_uuid(report_id),
                "section_name": section_name,
                "section_number": SECTION_NUMBERS.get(section_name),
                "format": preview_format,
                "page": page,
                "file_path": str(file_path),
                "generated_at": _utc_now(),
            }
            statement = insert(ReportPreview).values(**values)
            statement = statement.on_conflict_do_update(
                constraint="uq_report_previews_report_section_format_page",
                set_=values,
            )
            session.execute(statement)

        self._run("save_preview_metadata", operation)

    def save_final_output_metadata(
        self,
        report_id: str,
        final_docx_path: Path | None = None,
        final_pdf_path: Path | None = None,
    ) -> None:
        def operation(session: Session) -> None:
            values = {
                "report_id": _as_uuid(report_id),
                "final_docx_path": str(final_docx_path) if final_docx_path else None,
                "final_pdf_path": str(final_pdf_path) if final_pdf_path else None,
                "generated_at": _utc_now(),
            }
            statement = insert(ReportOutput).values(**values)
            statement = statement.on_conflict_do_update(
                index_elements=[ReportOutput.report_id],
                set_=values,
            )
            session.execute(statement)

        self._run("save_final_output_metadata", operation)

    def list_uploads_for_report(self, report_id: str) -> list[ReportUpload]:
        def operation(session: Session) -> list[ReportUpload]:
            return list(
                session.execute(
                    select(ReportUpload).where(
                        ReportUpload.report_id == _as_uuid(report_id)
                    )
                ).scalars()
            )

        return self._run("list_uploads_for_report", operation) or []

    def has_required_uploads(self, report_id: str, required_uploads: set[str]) -> bool:
        uploads = self.list_uploads_for_report(report_id)
        present = {upload.upload_type for upload in uploads}
        return required_uploads.issubset(present)

    def delete_report(self, report_id: str) -> None:
        def operation(session: Session) -> None:
            report = session.get(Report, _as_uuid(report_id))
            if report is not None:
                session.delete(report)

        self._run("delete_report", operation)
