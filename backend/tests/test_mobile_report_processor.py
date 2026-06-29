from pathlib import Path

import pandas as pd

from app.services.mobile_report_processor import (
    mobile_report_response,
    normalize_mobile_report,
)
from app.services.report_upload_service import dataframe_from_upload_bytes


FIXTURES_DIR = Path(__file__).parent / "fixtures"


def test_upload_reader_drops_repeated_weighbridge_register_header_row():
    fixture = FIXTURES_DIR / "mobile_report.csv"

    df = dataframe_from_upload_bytes(fixture.name, fixture.read_bytes())

    assert len(df) == 4
    assert df.iloc[0]["NO"] == "1"


def test_mobile_report_extracts_vehicle_records_and_summary():
    fixture = FIXTURES_DIR / "mobile_report.csv"
    df = pd.read_csv(fixture)

    records = normalize_mobile_report(df)
    payload = mobile_report_response(df)

    assert len(records) == 4
    assert records.iloc[0]["date_time"].strftime("%Y-%m-%d %H:%M") == "2026-05-12 00:36"
    assert records.iloc[0]["registration"] == "KBW781J"
    assert records.iloc[0]["total_gvw_kg"] == 37500
    assert records.iloc[0]["gvw_difference_kg"] == 7500
    assert records.iloc[0]["remarks"] == "CHARGED"
    assert records.iloc[1]["remarks"] == "WARNED"
    assert records.iloc[2]["remarks"] == "LEGAL"
    assert records.iloc[3]["remarks"] == "LEGAL"
    assert records["excess_kg"].sum() == 9300

    assert payload["summary"]["total_records"] == 4
    assert payload["summary"]["total_trucks_weighed"] == 4
    assert payload["summary"]["report_date"] == "2026-05-12"
    assert payload["summary"]["station"] == "JJM2"
    assert payload["summary"]["warned_trucks"] == 1
    assert payload["summary"]["charged_gvw_axle_trucks"] == 1
    assert payload["summary"]["charged_dimensions_trucks"] == 0
    assert payload["summary"]["overloaded_records"] == 2
    assert payload["summary"]["total_excess_kg"] == 9300
    assert payload["summary"]["mismatch_records"] == 4
    assert payload["summary"]["hourly_counts"]["0000-0100"] == 1
    assert payload["summary"]["hourly_counts"]["0600-0700"] == 3
    assert payload["data"][0]["date_time"] == "2026-05-12 00:36:00"


def test_mobile_report_recreates_worked_csv_calculations():
    fixture = FIXTURES_DIR / "mobile_report_worked.csv"
    df = pd.read_csv(fixture)

    records = normalize_mobile_report(df)

    assert records["registration"].tolist() == [
        "KBW781J",
        "KCF479D",
        "KDV441Q",
        "KDV441Q",
    ]
    assert records["total_gvw_kg"].tolist() == [37500, 29950, 6400, 17600]
    assert records["gvw_difference_kg"].tolist() == [7500, -50, -11600, -400]
    assert records["remarks"].tolist() == ["CHARGED", "LEGAL", "LEGAL", "WARNED"]


def test_mobile_report_detects_dimension_charged_remarks():
    df = pd.DataFrame(
        [
            {
                "NO": 1,
                "id": 1,
                "Date Time": "13/5/2026 08:10",
                "Ticket No.": "T1",
                "Station": "JJM2",
                "Registration": "KAA001A",
                "Axle": "2A",
                "Transporter": "Example",
                "Cargo": "Blocks",
                "Make": "Isuzu",
                "Origin": "Nairobi",
                "Destination": "Thika",
                "Deck A[KG]": 2000,
                "Deck B[KG]": 3000,
                "Deck C[KG]": 0,
                "Deck D[KG]": 0,
                "GVW [KG]": 18000,
                "Remarks": "charged (dimensions)",
                "Excess": "",
                "Excess [KG]": 0,
                "Status": "",
                "State": "",
                "Mismatch": "",
            }
        ]
    )

    records = normalize_mobile_report(df)
    payload = mobile_report_response(df)

    assert records.iloc[0]["remarks"] == "CHARGED (DIMENSIONS)"
    assert bool(records.iloc[0]["is_dimension_charge"]) is True
    assert bool(records.iloc[0]["is_gvw_axle_charge"]) is False
    assert payload["summary"]["charged_dimensions_trucks"] == 1
    assert payload["summary"]["charged_gvw_axle_trucks"] == 0


def test_upload_handles_csv_with_varying_column_counts():
    csv_content = (
        b"col1,col2,col3\n"
        b"col1,col2,col3\n"
        b"val1,val2,val3,val4\n"
    )
    df = dataframe_from_upload_bytes("test.csv", csv_content)
    assert len(df) == 1
    assert df.columns.tolist() == ["col1", "col2", "col3", "Extra_3"]
    assert df.iloc[0]["Extra_3"] == "val4"

