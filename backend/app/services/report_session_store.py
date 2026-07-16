import json
import os
import shutil
import time
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from uuid import uuid4

import pandas as pd
from dotenv import load_dotenv

from app.services.daily_summary_processor import (
    DailySummaryMissingSourceError,
    build_daily_summary_from_session,
)
from app.repositories.report_repository import ReportRepository
from app.services.traffic_census_processor import normalize_traffic_census_input
from app.services.transgressions_processor import normalize_transgressions_input


load_dotenv()

STORAGE_ROOT = Path(
    os.getenv("REPORT_STORAGE_ROOT", Path(__file__).resolve().parents[1] / "storage")
)
SESSIONS_DIR = STORAGE_ROOT / "sessions"
UPLOADS_DIR = STORAGE_ROOT / "uploads"
PROCESSED_DIR = STORAGE_ROOT / "processed"
PREVIEWS_DIR = STORAGE_ROOT / "previews"
FINAL_REPORTS_DIR = STORAGE_ROOT / "final_reports"
DEFAULT_SESSION_RETENTION_HOURS = 168
PREVIEW_CACHE_VERSION = "v3"


def _default_sections() -> dict[str, dict[str, Any]]:
    return {
        "daily_hour": {"status": "missing"},
        "wideload": {"status": "missing"},
        "impounded_prohibited": {"status": "missing"},
        "overloaded": {"status": "missing"},
        "mobile_report": {"status": "missing"},
        "traffic_census": {"status": "missing"},
        "daily_summary": {"status": "missing"},
        "transgressions": {"status": "missing"},
    }


def _safe_filename(filename: str | None, fallback: str = "upload") -> str:
    safe = Path(filename or fallback).name.strip()
    return safe or fallback


def _json_safe_value(value: Any) -> Any:
    if isinstance(value, pd.Timestamp):
        return value.isoformat(sep=" ")

    return value


STATIC_REQUIRED_UPLOADS = {
    "daily_hour",
    "wideload",
    "impounded_prohibited",
    "overloaded",
}


@dataclass
class ReportSession:
    report_id: str
    report_date: str
    station: str
    bound: str
    weighbridge_name: str | None = None
    prepared_by: str | None = None
    confirmed_by: str | None = None
    manual_inputs: dict[str, Any] = field(default_factory=dict)
    sections: dict[str, dict[str, Any]] = field(default_factory=_default_sections)
    dataframes: dict[str, pd.DataFrame] = field(default_factory=dict)
    final_report: bytes | None = None
    final_report_status: str = "not_built"
    final_report_error: str | None = None


class ReportSessionStore:
    def __init__(self, storage_root: Path = STORAGE_ROOT):
        self.storage_root = storage_root
        self.sessions_dir = storage_root / "sessions"
        self.uploads_dir = storage_root / "uploads"
        self.processed_dir = storage_root / "processed"
        self.previews_dir = storage_root / "previews"
        self.final_reports_dir = storage_root / "final_reports"
        self._sessions: dict[str, ReportSession] = {}
        self.repository = ReportRepository()
        self._ensure_storage_dirs()

    def _ensure_storage_dirs(self) -> None:
        for directory in [
            self.sessions_dir,
            self.uploads_dir,
            self.processed_dir,
            self.previews_dir,
            self.final_reports_dir,
        ]:
            directory.mkdir(parents=True, exist_ok=True)

    def _is_within_storage_root(self, path: Path) -> bool:
        try:
            path.resolve(strict=False).relative_to(
                self.storage_root.resolve(strict=False)
            )
        except ValueError:
            return False
        return True

    def _remove_storage_path(self, path: Path) -> None:
        if not self._is_within_storage_root(path):
            return

        try:
            if path.is_dir() and not path.is_symlink():
                shutil.rmtree(path, ignore_errors=True)
            else:
                path.unlink(missing_ok=True)
        except FileNotFoundError:
            return

    def _remove_session_artifacts(self, report_id: str) -> None:
        paths = [
            self._session_metadata_path(report_id),
            self.uploads_dir / report_id,
            self.processed_dir / report_id,
            self.previews_dir / report_id,
            self.final_reports_dir / report_id,
        ]

        for path in paths:
            self._remove_storage_path(path)

    def delete(self, report_id: str) -> bool:
        existed = self.get(report_id) is not None
        self._remove_session_artifacts(report_id)
        self._sessions.pop(report_id, None)
        self.repository.delete_report(report_id)
        return existed

    def cleanup_expired_sessions(
        self,
        max_age_hours: int = DEFAULT_SESSION_RETENTION_HOURS,
    ) -> list[str]:
        from datetime import timedelta
        cutoff_dt = datetime.now(timezone.utc) - timedelta(hours=max_age_hours)
        deleted_report_ids: list[str] = []

        if self.repository.enabled:
            summaries = self.repository.list_reports(limit=100000)
            for summary in summaries:
                updated_at = summary.get("updated_at")
                if updated_at and updated_at < cutoff_dt:
                    report_id = summary["report_id"]
                    self._remove_session_artifacts(report_id)
                    self._sessions.pop(report_id, None)
                    self.repository.delete_report(report_id)
                    deleted_report_ids.append(report_id)
        else:
            cutoff_timestamp = time.time() - (max_age_hours * 60 * 60)
            for metadata_path in sorted(self.sessions_dir.glob("*.json")):
                if not self._is_within_storage_root(metadata_path):
                    continue

                try:
                    modified_at = metadata_path.stat().st_mtime
                except FileNotFoundError:
                    continue

                if modified_at > cutoff_timestamp:
                    continue

                report_id = metadata_path.stem
                self._remove_session_artifacts(report_id)
                self._sessions.pop(report_id, None)
                self.repository.delete_report(report_id)
                deleted_report_ids.append(report_id)

        return deleted_report_ids

    def _session_metadata_path(self, report_id: str) -> Path:
        return self.sessions_dir / f"{report_id}.json"

    def _processed_session_dir(self, report_id: str) -> Path:
        return self.processed_dir / report_id

    def _processed_section_path(self, report_id: str, section: str) -> Path:
        return self._processed_session_dir(report_id) / f"{section}.pkl"

    def _final_report_path(self, report_id: str) -> Path:
        return self.final_reports_dir / report_id / "final_report.docx"

    def _preview_session_dir(self, report_id: str) -> Path:
        return self.previews_dir / report_id

    def _write_json_atomic(self, path: Path, payload: dict[str, Any]) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        temp_path = path.with_suffix(f"{path.suffix}.tmp")
        temp_path.write_text(
            json.dumps(payload, indent=2, sort_keys=True),
            encoding="utf-8",
        )
        temp_path.replace(path)

    def _metadata_payload(self, session: ReportSession) -> dict[str, Any]:
        return {
            "report_id": session.report_id,
            "report_date": session.report_date,
            "station": session.station,
            "bound": session.bound,
            "weighbridge_name": session.weighbridge_name,
            "prepared_by": session.prepared_by,
            "confirmed_by": session.confirmed_by,
            "manual_inputs": session.manual_inputs,
            "sections": session.sections,
            "final_report_status": session.final_report_status,
            "final_report_error": session.final_report_error,
        }

    def _save_metadata(self, session: ReportSession) -> None:
        payload = self._metadata_payload(session)
        if not self.repository.enabled:
            self._write_json_atomic(
                self._session_metadata_path(session.report_id),
                payload,
            )
        self.repository.save_session_snapshot(payload)

    def _session_from_metadata_payload(self, payload: dict[str, Any]) -> ReportSession:
        report_id = payload["report_id"]
        session = ReportSession(
            report_id=report_id,
            report_date=payload["report_date"],
            station=payload["station"],
            bound=payload["bound"],
            weighbridge_name=payload.get("weighbridge_name"),
            prepared_by=payload.get("prepared_by"),
            confirmed_by=payload.get("confirmed_by"),
            manual_inputs=payload.get("manual_inputs") or {},
            sections=payload.get("sections") or _default_sections(),
            final_report_status=payload.get("final_report_status", "not_built"),
            final_report_error=payload.get("final_report_error"),
        )

        processed_dir = self._processed_session_dir(report_id)
        if processed_dir.exists():
            for dataframe_path in processed_dir.glob("*.pkl"):
                section = dataframe_path.stem
                df = pd.read_pickle(dataframe_path)
                if isinstance(df, pd.DataFrame):
                    session.dataframes[section] = df

        if session.sections.get("daily_summary", {}).get("status") != "ready":
            self._refresh_daily_summary_status(session)

        final_report_path = self._final_report_path(report_id)
        if session.final_report_status == "ready" and final_report_path.exists():
            session.final_report = final_report_path.read_bytes()
        else:
            session.final_report = None

        self._sessions[report_id] = session
        return session

    def _load_session_from_disk(self, report_id: str) -> ReportSession | None:
        metadata_path = self._session_metadata_path(report_id)

        if self.repository.enabled:
            payload = self.repository.load_session_snapshot(report_id)
        elif metadata_path.exists():
            payload = json.loads(metadata_path.read_text(encoding="utf-8"))
        else:
            payload = None

        if not payload:
            return None

        return self._session_from_metadata_payload(payload)

    def _invalidate_generated_outputs(self, session: ReportSession) -> None:
        preview_dir = self._preview_session_dir(session.report_id)
        if preview_dir.exists():
            shutil.rmtree(preview_dir)

        final_report_dir = self.final_reports_dir / session.report_id
        if final_report_dir.exists():
            shutil.rmtree(final_report_dir)

        session.final_report = None
        session.final_report_status = "not_built"
        session.final_report_error = None

    def _refresh_daily_summary_status(self, session: ReportSession) -> None:
        try:
            daily_summary = build_daily_summary_from_session(session)
        except DailySummaryMissingSourceError as exc:
            session.sections["daily_summary"] = {
                "status": "missing",
                "message": str(exc),
            }
            return

        session.sections["daily_summary"] = {
            "status": "ready",
            "preview_url": (
                f"/api/report-sessions/{session.report_id}/sections/daily_summary/preview?format=png"
            ),
            "values": daily_summary,
        }

    def create(
        self,
        report_date: str,
        station: str,
        bound: str,
        weighbridge_name: str | None = None,
        prepared_by: str | None = None,
        confirmed_by: str | None = None,
    ) -> ReportSession:
        report_id = str(uuid4())
        session = ReportSession(
            report_id=report_id,
            report_date=report_date,
            station=station,
            bound=bound,
            weighbridge_name=weighbridge_name or station,
            prepared_by=prepared_by,
            confirmed_by=confirmed_by,
            sections=_default_sections(),
        )
        self._sessions[report_id] = session
        self._save_metadata(session)

        if self.repository.enabled:
            self.repository.upsert_manual_inputs(
                report_id=report_id,
                manual_inputs=session.manual_inputs,
                prepared_by=prepared_by,
                approved_by=confirmed_by,
                weighbridge_name=weighbridge_name or station,
                bound_name=bound,
            )

        return session

    def update_metadata(
        self,
        report_id: str,
        report_date: str | None = None,
        station: str | None = None,
        bound: str | None = None,
        weighbridge_name: str | None = None,
        prepared_by: str | None = None,
        confirmed_by: str | None = None,
    ) -> ReportSession:
        session = self.require(report_id)

        changed = False

        if report_date is not None and report_date != session.report_date:
            session.report_date = report_date
            changed = True

        if station is not None and station != session.station:
            session.station = station
            changed = True

        if bound is not None and bound != session.bound:
            session.bound = bound
            changed = True

        if (
            weighbridge_name is not None
            and weighbridge_name != session.weighbridge_name
        ):
            session.weighbridge_name = weighbridge_name
            changed = True

        if prepared_by is not None and prepared_by != session.prepared_by:
            session.prepared_by = prepared_by
            changed = True

        if confirmed_by is not None and confirmed_by != session.confirmed_by:
            session.confirmed_by = confirmed_by
            changed = True

        if changed:
            self._invalidate_generated_outputs(session)
            self._save_metadata(session)

        return session

    def get(self, report_id: str) -> ReportSession | None:
        return self._sessions.get(report_id) or self._load_session_from_disk(report_id)

    def require(self, report_id: str) -> ReportSession:
        session = self.get(report_id)

        if session is None:
            raise KeyError(report_id)

        return session

    def list_report_ids(self) -> list[str]:
        if self.repository.enabled:
            return self.repository.list_report_ids()
        file_ids = [path.stem for path in self.sessions_dir.glob("*.json")]
        return list(dict.fromkeys(sorted(file_ids)))

    def _session_metadata_mtime(self, report_id: str) -> datetime:
        metadata_path = self._session_metadata_path(report_id)
        if metadata_path.exists():
            return datetime.fromtimestamp(metadata_path.stat().st_mtime, tz=timezone.utc)
        return datetime.now(timezone.utc)

    def _session_report_type(self, session: ReportSession) -> str:
        station = f"{session.station or ''} {session.weighbridge_name or ''}".lower()
        bound = (session.bound or "").lower()
        mobile_ready = session.sections.get("mobile_report", {}).get("status") == "ready"

        if mobile_ready or "mobile" in station or "mobile" in bound:
            return "mobile_weighbridge"

        return "static_weighbridge"

    def _session_title(self, session: ReportSession) -> str:
        return " ".join(
            part
            for part in [
                session.weighbridge_name or session.station,
                session.bound,
                session.report_date,
            ]
            if part
        )

    def report_history_summary(self, report_id: str) -> dict[str, Any] | None:
        database_summary = self.repository.get_report_summary(report_id)
        if database_summary:
            return database_summary

        session = self.get(report_id)
        if not session:
            return None

        upload_count = sum(
            1
            for section in STATIC_REQUIRED_UPLOADS
            if session.sections.get(section, {}).get("status") == "ready"
        )
        required_uploads_completed = upload_count == len(STATIC_REQUIRED_UPLOADS)
        manual_inputs_completed = bool(session.manual_inputs)
        final_report_path = self._final_report_path(report_id)
        has_final_report = (
            session.final_report_status == "ready" and final_report_path.exists()
        )
        updated_at = self._session_metadata_mtime(report_id)

        return {
            "report_id": report_id,
            "title": self._session_title(session),
            "report_type": self._session_report_type(session),
            "weighbridge_name": session.weighbridge_name or session.station,
            "bound_name": session.bound,
            "status": {
                "ready": "completed",
                "error": "failed",
                "not_built": "draft",
            }.get(session.final_report_status, session.final_report_status),
            "created_at": updated_at,
            "updated_at": updated_at,
            "completed_at": updated_at if has_final_report else None,
            "has_final_report": has_final_report,
            "upload_count": upload_count,
            "required_uploads_completed": required_uploads_completed,
            "manual_inputs_completed": manual_inputs_completed,
            "download_available": has_final_report,
        }

    def list_report_history(
        self,
        status: str | None = None,
        report_type: str | None = None,
        limit: int = 20,
        offset: int = 0,
        search: str | None = None,
    ) -> list[dict[str, Any]]:
        database_summaries = self.repository.list_reports(
            status=status,
            report_type=report_type,
            limit=10_000,
            offset=0,
            search=search,
        )
        summaries_by_id = {
            summary["report_id"]: summary for summary in database_summaries
        }

        summaries = list(database_summaries)
        for report_id in self.list_report_ids():
            if report_id in summaries_by_id:
                continue

            summary = self.report_history_summary(report_id)
            if not summary:
                continue

            if status and summary["status"] != status:
                continue

            if report_type and summary["report_type"] != report_type:
                continue

            if search:
                haystack = " ".join(
                    str(summary.get(key) or "")
                    for key in ("report_id", "title", "weighbridge_name", "bound_name")
                ).lower()
                if search.lower() not in haystack:
                    continue

            summaries.append(summary)

        summaries.sort(key=lambda item: item["created_at"], reverse=True)
        return summaries[offset : offset + limit]

    def count_report_history(
        self,
        status: str | None = None,
        report_type: str | None = None,
        search: str | None = None,
    ) -> int:
        database_count = self.repository.count_reports(
            status=status,
            report_type=report_type,
            search=search,
        )

        if database_count:
            return database_count

        return len(
            self.list_report_history(
                status=status,
                report_type=report_type,
                limit=10_000,
                offset=0,
                search=search,
            )
        )

    def save_upload(self, report_id: str, section: str, filename: str | None, content: bytes) -> Path:
        self.require(report_id)
        upload_dir = self.uploads_dir / report_id / section
        upload_dir.mkdir(parents=True, exist_ok=True)
        upload_path = upload_dir / _safe_filename(filename)
        upload_path.write_bytes(content)
        self.repository.upsert_upload_metadata(
            report_id=report_id,
            upload_type=section,
            original_filename=filename,
            file_path=upload_path,
            file_size_bytes=len(content),
        )
        return upload_path

    def preview_cache_path(
        self,
        report_id: str,
        section_name: str,
        preview_format: str,
        page: int | None = None,
    ) -> Path:
        extension = preview_format.strip().lower()
        cache_page = page if extension == "png" else None
        filename = f"page-{cache_page}.{extension}" if cache_page else f"preview.{extension}"
        return (
            self._preview_session_dir(report_id)
            / section_name
            / PREVIEW_CACHE_VERSION
            / filename
        )

    def read_cached_preview(
        self,
        report_id: str,
        section_name: str,
        preview_format: str,
        page: int | None = None,
    ) -> bytes | None:
        preview_path = self.preview_cache_path(report_id, section_name, preview_format, page)
        if not preview_path.exists():
            return None
        return preview_path.read_bytes()

    def write_cached_preview(
        self,
        report_id: str,
        section_name: str,
        preview_format: str,
        content: bytes,
        page: int | None = None,
    ) -> Path:
        preview_path = self.preview_cache_path(report_id, section_name, preview_format, page)
        preview_path.parent.mkdir(parents=True, exist_ok=True)
        preview_path.write_bytes(content)
        self.repository.save_preview_metadata(
            report_id=report_id,
            section_name=section_name,
            preview_format=preview_format,
            file_path=preview_path,
            page=page,
        )
        return preview_path

    def set_section_ready(
        self,
        report_id: str,
        section: str,
        dataframe: pd.DataFrame,
        filename: str | None = None,
        preview_rows: int = 5,
        extra: dict[str, Any] | None = None,
    ) -> ReportSession:
        session = self.require(report_id)
        session.dataframes[section] = dataframe

        processed_path = self._processed_section_path(report_id, section)
        processed_path.parent.mkdir(parents=True, exist_ok=True)
        dataframe.to_pickle(processed_path)

        preview_df = dataframe.head(preview_rows).astype(object)
        preview_df = preview_df.where(pd.notnull(preview_df), None)

        section_state: dict[str, Any] = {
            "status": "ready",
            "filename": filename,
            "rows": len(dataframe),
            "columns": dataframe.columns.tolist(),
            "preview": [
                {
                    key: _json_safe_value(value)
                    for key, value in record.items()
                }
                for record in preview_df.to_dict(orient="records")
            ],
        }

        if section in {"daily_hour", "wideload", "impounded_prohibited"}:
            section_state["preview_url"] = (
                f"/api/report-sessions/{report_id}/sections/{section}/preview?format=png"
            )

        if section == "daily_hour":
            section_state["preview_pages"] = [
                {
                    "label": "Daily and Hourly Statistics",
                    "url": f"/api/report-sessions/{report_id}/sections/daily-hour-statistics/preview?format=png&page=1",
                },
                {
                    "label": "Daily Hourly Data",
                    "url": f"/api/report-sessions/{report_id}/sections/daily-hour-chart/preview?format=png&page=2",
                },
            ]
            summary = {}
            if "DATE" in dataframe.columns:
                totals_mask = dataframe["DATE"].astype(str).str.strip().str.lower().eq("totals")
                if totals_mask.any():
                    totals_row = dataframe.loc[totals_mask].iloc[-1]
                    for col in dataframe.columns:
                        if col != "DATE":
                            try:
                                val = totals_row.get(col, 0)
                                if pd.isna(val) or val is None:
                                    summary[col] = 0
                                else:
                                    summary[col] = int(float(str(val).replace(",", "")))
                            except Exception:
                                summary[col] = 0
            section_state["summary"] = summary

        if extra:
            section_state.update(extra)

        session.sections[section] = section_state
        self._refresh_daily_summary_status(session)
        self._invalidate_generated_outputs(session)
        self._save_metadata(session)

        if filename and self.repository.enabled:
            self.repository.upsert_upload_metadata(
                report_id=report_id,
                upload_type=section,
                original_filename=filename,
                file_path=processed_path,
                file_size_bytes=processed_path.stat().st_size if processed_path.exists() else 0,
            )
        return session

    def set_section_error(self, report_id: str, section: str, message: str) -> ReportSession:
        session = self.require(report_id)
        session.sections[section] = {
            "status": "error",
            "message": message,
        }
        session.dataframes.pop(section, None)

        processed_path = self._processed_section_path(report_id, section)
        if processed_path.exists():
            processed_path.unlink()

        self._refresh_daily_summary_status(session)
        self._invalidate_generated_outputs(session)
        self._save_metadata(session)
        return session

    def update_manual_inputs(
        self,
        report_id: str,
        prepared_by: str | None = None,
        confirmed_by: str | None = None,
        weighbridge_name: str | None = None,
        traffic_census: dict[str, Any] | None = None,
        transgressions: dict[str, Any] | list[dict[str, Any]] | None = None,
        extra: dict[str, Any] | None = None,
    ) -> ReportSession:
        session = self.require(report_id)

        if prepared_by is not None:
            session.prepared_by = prepared_by

        if confirmed_by is not None:
            session.confirmed_by = confirmed_by

        if weighbridge_name is not None:
            session.weighbridge_name = weighbridge_name

        if traffic_census is not None:
            normalized_traffic_census = normalize_traffic_census_input(traffic_census)
            session.manual_inputs["traffic_census"] = normalized_traffic_census
            session.sections["traffic_census"] = {
                "status": "ready",
                "preview_url": (
                    f"/api/report-sessions/{report_id}/sections/traffic_census/preview?format=png"
                ),
                "values": normalized_traffic_census,
            }

        if transgressions is not None:
            normalized_transgressions = normalize_transgressions_input(transgressions)
            session.manual_inputs["transgressions"] = normalized_transgressions
            session.sections["transgressions"] = {
                "status": "ready",
                "preview_url": (
                    f"/api/report-sessions/{report_id}/sections/transgressions/preview?format=png"
                ),
                "daily_transgressions_count": len(
                    normalized_transgressions["daily_transgressions"]
                ),
                "action_report_count": len(normalized_transgressions["action_report"]),
            }

        if extra:
            session.manual_inputs.update(extra)

        self._refresh_daily_summary_status(session)
        self._invalidate_generated_outputs(session)
        self._save_metadata(session)
        self.repository.upsert_manual_inputs(
            report_id=report_id,
            manual_inputs=session.manual_inputs,
            prepared_by=session.prepared_by,
            approved_by=session.confirmed_by,
            weighbridge_name=session.weighbridge_name,
            bound_name=session.bound,
        )
        return session

    def set_report_processing(self, report_id: str) -> ReportSession:
        session = self.require(report_id)
        self.repository.update_report_status(report_id, "processing")
        return session

    def set_final_report(self, report_id: str, content: bytes) -> ReportSession:
        session = self.require(report_id)
        final_report_path = self._final_report_path(report_id)
        final_report_path.parent.mkdir(parents=True, exist_ok=True)
        final_report_path.write_bytes(content)
        self.repository.save_final_output_metadata(
            report_id=report_id,
            final_docx_path=final_report_path,
        )

        session.final_report = content
        session.final_report_status = "ready"
        session.final_report_error = None
        self._save_metadata(session)
        self.repository.update_report_status(report_id, "completed", completed=True)
        return session

    def set_final_report_error(self, report_id: str, message: str) -> ReportSession:
        session = self.require(report_id)
        final_report_dir = self.final_reports_dir / report_id
        if final_report_dir.exists():
            shutil.rmtree(final_report_dir)

        session.final_report = None
        session.final_report_status = "error"
        session.final_report_error = message
        self._save_metadata(session)
        self.repository.update_report_status(
            report_id,
            "failed",
            error_message=message,
        )
        return session


report_session_store = ReportSessionStore()
