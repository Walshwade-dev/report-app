import os
import time

import pytest


def set_metadata_age(store, report_id: str, age_hours: int) -> None:
    metadata_path = store.sessions_dir / f"{report_id}.json"
    timestamp = time.time() - (age_hours * 60 * 60)
    os.utime(metadata_path, (timestamp, timestamp))


def write_artifact(root, report_id: str, filename: str = "artifact.txt") -> None:
    artifact_dir = root / report_id
    artifact_dir.mkdir(parents=True, exist_ok=True)
    (artifact_dir / filename).write_text("artifact", encoding="utf-8")


def test_cleanup_deletes_old_session_and_all_artifacts(temp_store):
    session = temp_store.create(
        report_date="2026-02-02",
        station="Juja",
        bound="Thika Bound",
    )
    report_id = session.report_id

    write_artifact(temp_store.uploads_dir, report_id)
    write_artifact(temp_store.processed_dir, report_id)
    write_artifact(temp_store.previews_dir, report_id)
    write_artifact(temp_store.final_reports_dir, report_id, "final_report.docx")
    set_metadata_age(temp_store, report_id, age_hours=169)

    deleted_report_ids = temp_store.cleanup_expired_sessions()

    assert deleted_report_ids == [report_id]
    assert not (temp_store.sessions_dir / f"{report_id}.json").exists()
    assert not (temp_store.uploads_dir / report_id).exists()
    assert not (temp_store.processed_dir / report_id).exists()
    assert not (temp_store.previews_dir / report_id).exists()
    assert not (temp_store.final_reports_dir / report_id).exists()
    assert temp_store.get(report_id) is None


def test_cleanup_keeps_recent_session_and_artifacts(temp_store):
    session = temp_store.create(
        report_date="2026-02-02",
        station="Juja",
        bound="Thika Bound",
    )
    report_id = session.report_id
    write_artifact(temp_store.previews_dir, report_id)
    set_metadata_age(temp_store, report_id, age_hours=1)

    deleted_report_ids = temp_store.cleanup_expired_sessions()

    assert deleted_report_ids == []
    assert (temp_store.sessions_dir / f"{report_id}.json").exists()
    assert (temp_store.previews_dir / report_id).exists()
    assert temp_store.get(report_id) is not None


def test_cleanup_ignores_missing_partial_artifact_folders(temp_store):
    session = temp_store.create(
        report_date="2026-02-02",
        station="Juja",
        bound="Thika Bound",
    )
    report_id = session.report_id
    write_artifact(temp_store.previews_dir, report_id)
    set_metadata_age(temp_store, report_id, age_hours=169)

    deleted_report_ids = temp_store.cleanup_expired_sessions()

    assert deleted_report_ids == [report_id]
    assert not (temp_store.sessions_dir / f"{report_id}.json").exists()
    assert not (temp_store.uploads_dir / report_id).exists()
    assert not (temp_store.processed_dir / report_id).exists()
    assert not (temp_store.previews_dir / report_id).exists()
    assert not (temp_store.final_reports_dir / report_id).exists()


def test_cleanup_skips_paths_that_resolve_outside_storage_root(temp_store):
    session = temp_store.create(
        report_date="2026-02-02",
        station="Juja",
        bound="Thika Bound",
    )
    report_id = session.report_id
    outside_file = temp_store.storage_root.parent / "outside.txt"
    outside_file.write_text("do not delete", encoding="utf-8")
    upload_link = temp_store.uploads_dir / report_id
    upload_link.symlink_to(outside_file)
    set_metadata_age(temp_store, report_id, age_hours=169)

    deleted_report_ids = temp_store.cleanup_expired_sessions()

    assert deleted_report_ids == [report_id]
    assert outside_file.exists()
    assert upload_link.exists()


def test_cleanup_rejects_deleted_session_reload(temp_store):
    session = temp_store.create(
        report_date="2026-02-02",
        station="Juja",
        bound="Thika Bound",
    )
    report_id = session.report_id
    set_metadata_age(temp_store, report_id, age_hours=169)

    temp_store.cleanup_expired_sessions()

    with pytest.raises(KeyError):
        temp_store.require(report_id)
