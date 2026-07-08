import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from app.services.report_session_store import report_session_store


def _upload_path(report_id: str, section: str, filename: str | None) -> Path | None:
    upload_dir = report_session_store.uploads_dir / report_id / section
    if not upload_dir.exists():
        return None

    if filename:
        expected_path = upload_dir / Path(filename).name
        if expected_path.exists():
            return expected_path

    return next((path for path in upload_dir.iterdir() if path.is_file()), None)


def main() -> None:
    repository = report_session_store.repository
    if not repository.enabled:
        raise SystemExit("DATABASE_URL is not configured; cannot backfill metadata.")

    reports = 0
    manual_inputs = 0
    uploads = 0
    outputs = 0

    for metadata_path in sorted(report_session_store.sessions_dir.glob("*.json")):
        payload = json.loads(metadata_path.read_text(encoding="utf-8"))
        report_id = payload["report_id"]
        repository.save_session_snapshot(payload)
        reports += 1

        manual_payload = payload.get("manual_inputs") or {}
        if manual_payload:
            repository.upsert_manual_inputs(
                report_id=report_id,
                manual_inputs=manual_payload,
                prepared_by=payload.get("prepared_by"),
                approved_by=payload.get("confirmed_by"),
                weighbridge_name=payload.get("weighbridge_name"),
                bound_name=payload.get("bound"),
            )
            manual_inputs += 1

        for section, state in (payload.get("sections") or {}).items():
            if state.get("status") != "ready":
                continue

            upload_path = _upload_path(report_id, section, state.get("filename"))
            if upload_path is None:
                continue

            repository.upsert_upload_metadata(
                report_id=report_id,
                upload_type=section,
                original_filename=state.get("filename"),
                file_path=upload_path,
                file_size_bytes=upload_path.stat().st_size,
            )
            uploads += 1

        final_report_path = report_session_store.final_reports_dir / report_id / "final_report.docx"
        if payload.get("final_report_status") == "ready" and final_report_path.exists():
            repository.save_final_output_metadata(
                report_id=report_id,
                final_docx_path=final_report_path,
            )
            outputs += 1

    print(
        "Backfilled "
        f"{reports} reports, "
        f"{manual_inputs} manual input rows, "
        f"{uploads} upload rows, "
        f"{outputs} output rows."
    )


if __name__ == "__main__":
    main()
