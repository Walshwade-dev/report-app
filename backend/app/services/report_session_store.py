import json
import shutil
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any
from uuid import uuid4

import pandas as pd

from app.services.daily_summary_processor import (
    DailySummaryMissingSourceError,
    build_daily_summary_from_session,
)
from app.services.traffic_census_processor import normalize_traffic_census_input
from app.services.transgressions_processor import normalize_transgressions_input


STORAGE_ROOT = Path(__file__).resolve().parents[1] / "storage"
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
        "traffic_census": {"status": "missing"},
        "daily_summary": {"status": "missing"},
        "transgressions": {"status": "missing"},
    }


def _safe_filename(filename: str | None, fallback: str = "upload") -> str:
    safe = Path(filename or fallback).name.strip()
    return safe or fallback


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

    def cleanup_expired_sessions(
        self,
        max_age_hours: int = DEFAULT_SESSION_RETENTION_HOURS,
    ) -> list[str]:
        cutoff_timestamp = time.time() - (max_age_hours * 60 * 60)
        deleted_report_ids: list[str] = []

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
        self._write_json_atomic(
            self._session_metadata_path(session.report_id),
            self._metadata_payload(session),
        )

    def _load_session_from_disk(self, report_id: str) -> ReportSession | None:
        metadata_path = self._session_metadata_path(report_id)

        if not metadata_path.exists():
            return None

        payload = json.loads(metadata_path.read_text(encoding="utf-8"))
        session = ReportSession(
            report_id=payload["report_id"],
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
                session.dataframes[section] = pd.read_pickle(dataframe_path)

        self._refresh_daily_summary_status(session)

        final_report_path = self._final_report_path(report_id)
        if session.final_report_status == "ready" and final_report_path.exists():
            session.final_report = final_report_path.read_bytes()
        elif session.final_report_status == "ready":
            session.final_report_status = "not_built"
            session.final_report_error = "Final report file is missing from storage."
            self._save_metadata(session)

        self._sessions[report_id] = session
        return session

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
        return session

    def update_metadata(
        self,
        report_id: str,
        station: str | None = None,
        bound: str | None = None,
        weighbridge_name: str | None = None,
    ) -> ReportSession:
        session = self.require(report_id)

        changed = False

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

    def save_upload(self, report_id: str, section: str, filename: str | None, content: bytes) -> Path:
        self.require(report_id)
        upload_dir = self.uploads_dir / report_id / section
        upload_dir.mkdir(parents=True, exist_ok=True)
        upload_path = upload_dir / _safe_filename(filename)
        upload_path.write_bytes(content)
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
            "rows": int(len(dataframe)),
            "columns": dataframe.columns.tolist(),
            "preview": preview_df.to_dict(orient="records"),
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

        if extra:
            section_state.update(extra)

        session.sections[section] = section_state
        self._refresh_daily_summary_status(session)
        self._invalidate_generated_outputs(session)
        self._save_metadata(session)
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
        return session

    def set_final_report(self, report_id: str, content: bytes) -> ReportSession:
        session = self.require(report_id)
        final_report_path = self._final_report_path(report_id)
        final_report_path.parent.mkdir(parents=True, exist_ok=True)
        final_report_path.write_bytes(content)

        session.final_report = content
        session.final_report_status = "ready"
        session.final_report_error = None
        self._save_metadata(session)
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
        return session


report_session_store = ReportSessionStore()
