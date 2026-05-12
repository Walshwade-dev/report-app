from pathlib import Path

import pandas as pd

from app.routes import reports
from app.services.report_session_store import ReportSessionStore


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

    metadata_path = temp_store.sessions_dir / f"{report_id}.json"
    assert metadata_path.exists()

    fetched = client.get(f"/api/report-sessions/{report_id}")
    assert fetched.status_code == 200
    fetched_payload = fetched.json()
    assert fetched_payload["metadata"]["station"] == "Juja"
    assert fetched_payload["metadata"]["bound"] == "Thika Bound"


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
