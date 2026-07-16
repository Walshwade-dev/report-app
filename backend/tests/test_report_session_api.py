import os

import pandas as pd

from app.routes import reports
from app.services.report_session_store import ReportSessionStore


ADMIN_HEADERS = {"X-Admin-Password": "test-admin-password"}


def create_session(client):
    response = client.post(
        "/api/report-sessions",
        json={
            "report_date": "2026-02-02",
            "station": "Juja",
            "bound": "Thika Bound",
            "weighbridge_name": "Juja",
        },
    )

    assert response.status_code == 200
    payload = response.json()
    assert payload["report_id"]
    return payload


def patch_manual_inputs(client, report_id: str):
    response = client.patch(
        f"/api/report-sessions/{report_id}/manual-inputs",
        json={
            "prepared_by": "Fredrick Kariuki",
            "confirmed_by": "Faith Njani",
            "traffic_census": {
                "buses_gte_3500kg": 1351,
                "vehicles_3500_to_7000_excluding_buses": 29,
                "vehicles_gte_7000_excluding_buses": 6,
            },
            "transgressions": {
                "daily_transgressions": [],
                "action_report": [],
            },
        },
    )

    assert response.status_code == 200
    return response.json()


def simulate_restart(temp_store, monkeypatch):
    restarted_store = ReportSessionStore(temp_store.storage_root)
    monkeypatch.setattr(reports, "report_session_store", restarted_store)
    return restarted_store


def full_daily_hour_df():
    return pd.DataFrame(
        [
            {
                "DATE": "Totals",
                "TIME": "",
                "D": 100,
                "S": 25,
                "M": 10,
                "H": 1010,
                "Q": 1000,
                "X": 135,
                "C": 10,
                "Y": 10,
                "P": 6,
                "A": 3,
                "Z": 2,
                "G": 1,
                "R": 4,
                "E": 7,
            }
        ]
    )


def minimal_report_dataframes():
    return {
        "daily_hour": full_daily_hour_df(),
        "wideload": pd.DataFrame({"Registration": ["KAA 001A"]}),
        "impounded_prohibited": pd.DataFrame({"Registration": ["KBB 002B"]}),
        "overloaded": pd.DataFrame(
            {"Vardict": ["Vehicle has a valid permit App-123", "No permit"]}
        ),
    }


def seed_required_sections(store: ReportSessionStore, report_id: str):
    dataframes = minimal_report_dataframes()
    store.set_section_ready(
        report_id,
        "daily_hour",
        dataframes["daily_hour"],
        filename="daily-hour.csv",
    )
    store.set_section_ready(
        report_id,
        "wideload",
        dataframes["wideload"],
        filename="wideload.csv",
        extra={"wideload_count": len(dataframes["wideload"])},
    )
    store.set_section_ready(
        report_id,
        "impounded_prohibited",
        dataframes["impounded_prohibited"],
        filename="impounded.csv",
    )
    store.set_section_ready(
        report_id,
        "overloaded",
        dataframes["overloaded"],
        filename="overloaded.csv",
        extra={"valid_permit_count": 1},
    )


def test_report_session_creation_persists_metadata(client, temp_store):
    payload = create_session(client)
    report_id = payload["report_id"]

    if temp_store.repository.enabled:
        assert temp_store.repository.load_session_snapshot(report_id) is not None
    else:
        metadata_path = temp_store.sessions_dir / f"{report_id}.json"
        assert metadata_path.exists()

    fetched = client.get(f"/api/report-sessions/{report_id}")
    assert fetched.status_code == 200
    fetched_payload = fetched.json()
    assert fetched_payload["metadata"]["station"] == "Juja"
    assert fetched_payload["metadata"]["bound"] == "Thika Bound"


def test_report_session_delete_removes_metadata_and_artifacts(client, temp_store):
    payload = create_session(client)
    report_id = payload["report_id"]
    upload_path = temp_store.save_upload(
        report_id,
        "daily_hour",
        "daily-hour.csv",
        b"DATE,TIME\n",
    )

    if temp_store.repository.enabled:
        assert temp_store.repository.load_session_snapshot(report_id) is not None
    else:
        metadata_path = temp_store.sessions_dir / f"{report_id}.json"
        assert metadata_path.exists()
    assert upload_path.exists()

    response = client.delete(
        f"/api/report-sessions/{report_id}",
        headers=ADMIN_HEADERS,
    )

    assert response.status_code == 200
    assert response.json() == {"status": "deleted", "report_id": report_id}
    if temp_store.repository.enabled:
        assert temp_store.repository.load_session_snapshot(report_id) is None
    else:
        metadata_path = temp_store.sessions_dir / f"{report_id}.json"
        assert not metadata_path.exists()
    assert not upload_path.exists()
    assert client.get(f"/api/report-sessions/{report_id}").status_code == 404


def test_report_session_history_requires_admin_password(client):
    response = client.get("/api/report-sessions")

    assert response.status_code == 401


def test_report_session_delete_requires_admin_password(client):
    report_id = create_session(client)["report_id"]

    response = client.delete(f"/api/report-sessions/{report_id}")

    assert response.status_code == 401


def test_report_session_history_listing_returns_newest_first(client, temp_store):
    older_id = create_session(client)["report_id"]
    newer_id = create_session(client)["report_id"]

    if temp_store.repository.enabled:
        from app.core.database import SessionLocal
        from app.db.models import Report
        from datetime import datetime, timezone, timedelta
        
        assert SessionLocal is not None
        with SessionLocal() as db_session:
            r1 = db_session.query(Report).filter(Report.id == older_id).first()
            r2 = db_session.query(Report).filter(Report.id == newer_id).first()
            if r1 and r2:
                r1.updated_at = datetime.now(timezone.utc) - timedelta(days=1)
                r2.updated_at = datetime.now(timezone.utc)
                db_session.commit()
    else:
        older_path = temp_store.sessions_dir / f"{older_id}.json"
        newer_path = temp_store.sessions_dir / f"{newer_id}.json"
        os.utime(older_path, (1000, 1000))
        os.utime(newer_path, (2000, 2000))

    response = client.get("/api/report-sessions", headers=ADMIN_HEADERS)

    assert response.status_code == 200
    payload = response.json()
    assert isinstance(payload, list)
    assert payload[0]["report_id"] == newer_id
    assert payload[1]["report_id"] == older_id
    assert "created_at" in payload[0]
    assert "required_uploads_completed" in payload[0]


def test_report_session_history_status_filter(client, temp_store):
    draft_id = create_session(client)["report_id"]
    completed_id = create_session(client)["report_id"]
    temp_store.set_final_report(completed_id, b"final report")

    response = client.get(
        "/api/report-sessions?status=completed",
        headers=ADMIN_HEADERS,
    )

    assert response.status_code == 200
    payload = response.json()
    assert [item["report_id"] for item in payload] == [completed_id]
    assert payload[0]["status"] == "completed"
    assert draft_id not in [item["report_id"] for item in payload]


def test_report_session_history_upload_completion(client, temp_store):
    report_id = create_session(client)["report_id"]
    seed_required_sections(temp_store, report_id)

    response = client.get(
        f"/api/report-sessions?search={report_id}",
        headers=ADMIN_HEADERS,
    )

    assert response.status_code == 200
    payload = response.json()
    assert len(payload) == 1
    assert payload[0]["upload_count"] == 4
    assert payload[0]["required_uploads_completed"] is True


def test_report_session_history_final_report_availability(client, temp_store):
    report_id = create_session(client)["report_id"]
    temp_store.set_final_report(report_id, b"final report")

    response = client.get(
        f"/api/report-sessions?search={report_id}",
        headers=ADMIN_HEADERS,
    )

    assert response.status_code == 200
    payload = response.json()
    assert len(payload) == 1
    assert payload[0]["has_final_report"] is True
    assert payload[0]["download_available"] is True


def test_report_session_metadata_update_persists_after_restart(client, temp_store, monkeypatch):
    report_id = create_session(client)["report_id"]

    response = client.patch(
        f"/api/report-sessions/{report_id}/metadata",
        json={
            "station": "ATHI RIVER",
            "bound": "MOMBASA BOUND",
            "weighbridge_name": "ATHI RIVER",
        },
    )

    assert response.status_code == 200
    payload = response.json()
    assert payload["metadata"]["station"] == "ATHI RIVER"
    assert payload["metadata"]["bound"] == "MOMBASA BOUND"
    assert payload["metadata"]["weighbridge_name"] == "ATHI RIVER"

    simulate_restart(temp_store, monkeypatch)

    recovered = client.get(f"/api/report-sessions/{report_id}")
    assert recovered.status_code == 200
    recovered_payload = recovered.json()
    assert recovered_payload["metadata"]["station"] == "ATHI RIVER"
    assert recovered_payload["metadata"]["bound"] == "MOMBASA BOUND"
    assert recovered_payload["metadata"]["weighbridge_name"] == "ATHI RIVER"


def test_manual_input_persistence_and_restart_recovery(client, temp_store, monkeypatch):
    report_id = create_session(client)["report_id"]
    payload = patch_manual_inputs(client, report_id)

    assert payload["metadata"]["prepared_by"] == "Fredrick Kariuki"
    assert payload["metadata"]["confirmed_by"] == "Faith Njani"
    assert payload["manual_inputs"]["traffic_census"]["total_traffic_census"] == 1386
    assert payload["sections"]["traffic_census"]["status"] == "ready"
    assert payload["sections"]["transgressions"]["status"] == "ready"

    simulate_restart(temp_store, monkeypatch)

    recovered = client.get(f"/api/report-sessions/{report_id}")
    assert recovered.status_code == 200
    recovered_payload = recovered.json()
    assert recovered_payload["metadata"]["prepared_by"] == "Fredrick Kariuki"
    assert recovered_payload["manual_inputs"]["traffic_census"]["total_traffic_census"] == 1386
    assert recovered_payload["sections"]["traffic_census"]["status"] == "ready"
    assert recovered_payload["sections"]["transgressions"]["status"] == "ready"


def test_preview_cache_recovers_after_restart(client, temp_store, monkeypatch):
    report_id = create_session(client)["report_id"]
    patch_manual_inputs(client, report_id)
    seed_required_sections(temp_store, report_id)

    first_preview = client.get(
        f"/api/report-sessions/{report_id}/sections/traffic-census/preview?format=docx"
    )
    assert first_preview.status_code == 200
    assert first_preview.content

    preview_path = (
        temp_store.previews_dir / report_id / "traffic_census" / "v3" / "preview.docx"
    )
    assert preview_path.exists()
    cached_bytes = preview_path.read_bytes()

    simulate_restart(temp_store, monkeypatch)

    second_preview = client.get(
        f"/api/report-sessions/{report_id}/sections/traffic-census/preview?format=docx"
    )
    assert second_preview.status_code == 200
    assert second_preview.content == cached_bytes
    assert preview_path.read_bytes() == cached_bytes


def test_final_report_persistence_and_download_after_restart(
    client,
    temp_store,
    monkeypatch,
):
    report_id = create_session(client)["report_id"]
    patch_manual_inputs(client, report_id)
    seed_required_sections(temp_store, report_id)

    build_response = client.post(f"/api/report-sessions/{report_id}/build-final-report")
    assert build_response.status_code == 200
    build_payload = build_response.json()
    assert build_payload["final_report"]["status"] == "ready"

    final_report_path = (
        temp_store.final_reports_dir / report_id / "final_report.docx"
    )
    assert final_report_path.exists()
    persisted_bytes = final_report_path.read_bytes()
    assert persisted_bytes

    simulate_restart(temp_store, monkeypatch)

    recovered = client.get(f"/api/report-sessions/{report_id}")
    assert recovered.status_code == 200
    assert recovered.json()["final_report"]["status"] == "ready"

    download = client.get(f"/api/report-sessions/{report_id}/download-final-report")
    assert download.status_code == 200
    assert download.content == persisted_bytes
    assert (
        download.headers["content-type"]
        == "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )


def test_build_final_report_reports_missing_sections(client):
    report_id = create_session(client)["report_id"]

    response = client.post(f"/api/report-sessions/{report_id}/build-final-report")

    assert response.status_code == 400
    detail = response.json()["detail"]
    assert detail["missing_sections"] == [
        "daily_hour",
        "wideload",
        "impounded_prohibited",
        "overloaded",
    ]
    assert detail["session"]["final_report"]["status"] == "error"
    assert "Missing or invalid required sections" in detail["message"]


def test_build_final_report_reports_generator_errors(
    client,
    temp_store,
    monkeypatch,
):
    report_id = create_session(client)["report_id"]
    patch_manual_inputs(client, report_id)
    seed_required_sections(temp_store, report_id)

    def fail_build_final_report(**_kwargs):
        raise RuntimeError("template failed")

    monkeypatch.setattr(reports, "build_final_report", fail_build_final_report)

    response = client.post(f"/api/report-sessions/{report_id}/build-final-report")

    assert response.status_code == 500
    detail = response.json()["detail"]
    assert detail["message"] == "Failed to build final report"
    assert detail["error"] == "template failed"
    assert detail["session"]["final_report"]["status"] == "error"


def test_is_bound_a():
    # Juja tests
    assert reports.is_bound_a("Juja", "Thika Bound") is True
    assert reports.is_bound_a("Juja", "Nairobi Bound") is False
    assert reports.is_bound_a("Juja", "Incoming") is True
    assert reports.is_bound_a("Juja", None) is True

    # Athi River tests
    assert reports.is_bound_a("Athi River", "Mombasa Bound") is True
    assert reports.is_bound_a("Athi River", "Nairobi Bound") is False

    # Gilgil tests
    assert reports.is_bound_a("Gilgil", "Nairobi Bound") is True
    assert reports.is_bound_a("Gilgil", "Nakuru Bound") is False

    # Kanyonyo tests
    assert reports.is_bound_a("Kanyonyo", "Mwingi Bound") is True
    assert reports.is_bound_a("Kanyonyo", "Thika Bound") is False

    # Isinya tests
    assert reports.is_bound_a("Isinya", "Kajiado Bound") is True
    assert reports.is_bound_a("Isinya", "Nairobi Bound") is False

    # Suswa tests
    assert reports.is_bound_a("Suswa", "Narok Bound") is True
    assert reports.is_bound_a("Suswa", "Nairobi Bound") is False

    # Fallback/None tests
    assert reports.is_bound_a(None, "Thika Bound") is True
    assert reports.is_bound_a(None, "Nairobi Bound") is False
    assert reports.is_bound_a("Unknown", "Nairobi Bound") is False
    assert reports.is_bound_a("Unknown", "Thika Bound") is True
