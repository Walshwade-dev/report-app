import io
import zipfile
from pathlib import Path
from xml.etree import ElementTree

from docx import Document
from openpyxl import load_workbook

from app.services.report_layout import (
    A4_LANDSCAPE_HEIGHT_INCHES,
    A4_LANDSCAPE_WIDTH_INCHES,
    A4_PRINTABLE_WIDTH_TWIPS,
    BOTTOM_MARGIN_INCHES,
    LEFT_MARGIN_INCHES,
    RIGHT_MARGIN_INCHES,
    TOP_MARGIN_INCHES,
)


FIXTURES_DIR = Path(__file__).parent / "fixtures"
WORD_NAMESPACE = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def fixture_upload(name: str, content_type: str = "text/csv"):
    path = FIXTURES_DIR / name
    return {
        "file": (
            path.name,
            path.read_bytes(),
            content_type,
        )
    }


def create_report_session(client) -> str:
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
    return payload["report_id"]


def patch_manual_inputs(client, report_id: str) -> None:
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
    payload = response.json()
    assert payload["sections"]["traffic_census"]["status"] == "ready"
    assert payload["sections"]["transgressions"]["status"] == "ready"


def upload_required_files(client, report_id: str) -> None:
    wideload = client.post(
        f"/api/report-sessions/{report_id}/uploads/wideload",
        files=fixture_upload("wideload.csv"),
    )
    assert wideload.status_code == 200
    wideload_payload = wideload.json()
    assert wideload_payload["sections"]["wideload"]["status"] == "ready"
    assert wideload_payload["sections"]["wideload"]["wideload_count"] == 1

    daily_hour = client.post(
        f"/api/report-sessions/{report_id}/uploads/daily-hour",
        data={"wideload_count": "999"},
        files=fixture_upload("daily_hour.csv"),
    )
    assert daily_hour.status_code == 200
    daily_hour_payload = daily_hour.json()
    assert daily_hour_payload["sections"]["daily_hour"]["status"] == "ready"
    assert daily_hour_payload["sections"]["daily_hour"]["wideload_count_used"] == 1

    impounded = client.post(
        f"/api/report-sessions/{report_id}/uploads/impounded-prohibited",
        files=fixture_upload("impounded_prohibited.csv"),
    )
    assert impounded.status_code == 200
    assert impounded.json()["sections"]["impounded_prohibited"]["status"] == "ready"

    overloaded = client.post(
        f"/api/report-sessions/{report_id}/uploads/overloaded",
        files=fixture_upload("overloaded.csv"),
    )
    assert overloaded.status_code == 200
    overloaded_payload = overloaded.json()
    assert overloaded_payload["sections"]["overloaded"]["status"] == "ready"
    assert overloaded_payload["sections"]["overloaded"]["valid_permit_count"] == 1


def assert_docx_preview(client, report_id: str, section_name: str) -> None:
    response = client.get(
        f"/api/report-sessions/{report_id}/sections/{section_name}/preview?format=docx"
    )

    assert response.status_code == 200
    assert response.content.startswith(b"PK")
    assert (
        response.headers["content-type"]
        == "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )


def docx_traffic_census_values(docx_bytes: bytes) -> dict[str, str]:
    document = Document(io.BytesIO(docx_bytes))

    for table in document.tables:
        if len(table.rows) < 2 or len(table.rows[0].cells) < 7:
            continue

        headers = [
            cell.text.replace("\n", " ").strip()
            for cell in table.rows[0].cells
        ]

        if not any("Exemption" in header and "(E)" in header for header in headers):
            continue

        values = [cell.text.strip() for cell in table.rows[1].cells]
        return {
            "k": values[3],
            "e": values[4],
            "x": values[5],
            "total_traffic": values[6],
        }

    raise AssertionError("Traffic Census table was not found in generated DOCX.")


def expected_traffic_census_values(temp_store, report_id: str) -> dict[str, str]:
    session = temp_store.require(report_id)
    daily_df = session.dataframes["daily_hour"]
    totals_row = daily_df[
        daily_df["DATE"].astype(str).str.strip().str.lower().eq("totals")
    ].iloc[-1]

    q = int(totals_row["Q"])
    x = int(totals_row["X"])
    e = int(session.sections["wideload"]["wideload_count"])
    k = int(session.manual_inputs["traffic_census"]["total_traffic_census"])

    return {
        "k": f"{k:,}",
        "e": f"{e:,}",
        "x": f"{x:,}",
        "total_traffic": f"{q + x + k + e:,}",
    }


def expected_summary_card_values(temp_store, report_id: str) -> dict[str, int]:
    session = temp_store.require(report_id)
    daily_df = session.dataframes["daily_hour"]
    totals_row = daily_df[
        daily_df["DATE"].astype(str).str.strip().str.lower().eq("totals")
    ].iloc[-1]

    return {
        "Total Weighed": int(totals_row["X"]),
        "Total Overloaded": int(totals_row["Y"]),
        "Special Released": int(totals_row["G"]),
        "Wide Loads": int(session.sections["wideload"]["wideload_count"]),
    }


def assert_a4_landscape_layout(docx_bytes: bytes) -> None:
    document = Document(io.BytesIO(docx_bytes))

    for section in document.sections:
        assert round(section.page_width.inches, 2) == A4_LANDSCAPE_WIDTH_INCHES
        assert round(section.page_height.inches, 2) == A4_LANDSCAPE_HEIGHT_INCHES
        assert round(section.left_margin.inches, 2) == LEFT_MARGIN_INCHES
        assert round(section.right_margin.inches, 2) == RIGHT_MARGIN_INCHES
        assert round(section.top_margin.inches, 2) == TOP_MARGIN_INCHES
        assert round(section.bottom_margin.inches, 2) == BOTTOM_MARGIN_INCHES


def assert_table_grids_fit_a4_printable_width(docx_bytes: bytes) -> None:
    with zipfile.ZipFile(io.BytesIO(docx_bytes)) as docx_zip:
        document_xml = docx_zip.read("word/document.xml")

    root = ElementTree.fromstring(document_xml)

    for grid in root.findall(".//w:tblGrid", WORD_NAMESPACE):
        widths = [
            int(column.attrib[f"{{{WORD_NAMESPACE['w']}}}w"])
            for column in grid.findall("w:gridCol", WORD_NAMESPACE)
        ]
        assert sum(widths) <= A4_PRINTABLE_WIDTH_TWIPS


def test_upload_fixtures_build_and_download_final_report(client, temp_store):
    report_id = create_report_session(client)
    upload_required_files(client, report_id)
    patch_manual_inputs(client, report_id)

    session_payload = client.get(f"/api/report-sessions/{report_id}").json()
    assert session_payload["sections"]["daily_summary"]["status"] == "ready"
    expected_traffic_values = expected_traffic_census_values(temp_store, report_id)
    expected_summary_values = expected_summary_card_values(temp_store, report_id)

    summary_cards = client.get(f"/api/report-sessions/{report_id}/summary-cards")
    assert summary_cards.status_code == 200
    summary_payload = summary_cards.json()
    assert summary_payload["report_id"] == report_id
    assert {
        card["title"]: card["value"]
        for card in summary_payload["cards"]
    } == expected_summary_values
    assert all(card["status"] == "ready" for card in summary_payload["cards"])

    for section_name in [
        "daily-hour-statistics",
        "traffic-census",
        "daily-summary",
        "transgressions",
        "wideload",
        "impounded-prohibited",
    ]:
        assert_docx_preview(client, report_id, section_name)

    traffic_preview = client.get(
        f"/api/report-sessions/{report_id}/sections/traffic-census/preview?format=docx"
    )
    assert traffic_preview.status_code == 200
    assert docx_traffic_census_values(traffic_preview.content) == expected_traffic_values

    build_response = client.post(f"/api/report-sessions/{report_id}/build-final-report")
    assert build_response.status_code == 200
    build_payload = build_response.json()
    assert build_payload["final_report"]["status"] == "ready"

    final_report_path = temp_store.final_reports_dir / report_id / "final_report.docx"
    assert final_report_path.exists()

    download = client.get(f"/api/report-sessions/{report_id}/download-final-report")
    assert download.status_code == 200
    assert download.content.startswith(b"PK")
    assert download.content == final_report_path.read_bytes()
    assert docx_traffic_census_values(download.content) == expected_traffic_values
    assert_a4_landscape_layout(download.content)
    assert_table_grids_fit_a4_printable_width(download.content)
    assert (
        download.headers["content-type"]
        == "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )


def test_upload_fixtures_download_excel_report(client, temp_store):
    report_id = create_report_session(client)
    upload_required_files(client, report_id)
    patch_manual_inputs(client, report_id)

    session_payload = client.get(f"/api/report-sessions/{report_id}").json()
    assert session_payload["excel_report"]["download_url"] == (
        f"/api/report-sessions/{report_id}/download-excel-report"
    )

    download = client.get(f"/api/report-sessions/{report_id}/download-excel-report")

    assert download.status_code == 200
    assert (
        download.headers["content-type"]
        == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    workbook = load_workbook(io.BytesIO(download.content))
    assert workbook.sheetnames == ["Summary", "CC records", "Hswim Weekly"]

    summary = workbook["Summary"]
    assert summary["D2"].value == "Trucks Weighed"
    assert summary["T2"].value == "TIME"
    assert summary["C33"].value == "Buses>= 3500kg  (J)"
    assert summary["H33"].value == "Exemption\npermits"
    assert summary["H35"].value == "(E)"
    assert summary["I33"].value == "Total Weighed"
    assert summary["I35"].value == "(X)"
    assert summary["J33"].value == "Total Traffic"
    assert summary["J35"].value == "(T)= (Q+X+K+E)"
    assert summary["B38"].value == "HSWIM CLEARED (Q)"
    assert len(summary._charts) == 1
    assert summary.column_dimensions["B"].width == 12.140625
    assert summary.row_dimensions[38].height == 67.5
    assert summary["D6"].fill.fgColor.rgb == "FFFFFF00"
    assert summary["H6"].fill.fgColor.rgb == "FFFF0000"
    assert summary["N40"].fill.fgColor.rgb == "00000000"
    assert summary["O40"].fill.fgColor.rgb == "00000000"
    assert summary["Q40"].fill.fgColor.rgb == "00000000"
    assert summary["P40"].value == "=Q30"
    assert summary._charts[0].x_axis.axPos == "b"
    assert "H33:H34" in {str(merged) for merged in summary.merged_cells.ranges}
    assert "I33:I34" in {str(merged) for merged in summary.merged_cells.ranges}
    assert "J33:J34" in {str(merged) for merged in summary.merged_cells.ranges}
    assert summary._charts[0].y_axis.scaling.min == -10
    assert summary._charts[0].y_axis.scaling.max == 300
    assert summary._charts[0].y_axis.majorUnit == 50
    assert summary._charts[0].x_axis.crosses == "min"
    assert all(series.smooth is False for series in summary._charts[0].series)
    assert [
        series.graphicalProperties.line.solidFill.srgbClr
        for series in summary._charts[0].series
    ] == ["4472C4", "FF0000", "70AD47", "7030A0"]
    assert summary["S48"].value == "Notes"
    assert summary["S48"].font.bold is True
    assert summary["S48"].alignment.horizontal == "left"
    assert summary["S49"].fill.fgColor.rgb == "FFFFFF00"
    assert summary["S50"].fill.fgColor.rgb == "FFFF0000"
    assert summary["S51"].fill.fgColor.rgb == "00000000"

    cc_records = workbook["CC records"]
    assert cc_records["B2"].value == "Total Traffic Census Summary"
    assert "B8:D10" not in {str(merged) for merged in cc_records.merged_cells.ranges}
    assert cc_records["B8"].value is not None
    assert cc_records["C9"].value is not None
    assert cc_records["D10"].value is not None
    assert cc_records["B14"].value == "=SUM(B6:B13)"
    assert cc_records["C14"].value == "=SUM(C6:C13)"
    assert cc_records["D14"].value == "=SUM(D6:D13)"
    assert cc_records["B14"].fill.fgColor.rgb == "FF00B050"
    assert cc_records["C14"].fill.fgColor.rgb == "FF00B050"
    assert cc_records["D14"].fill.fgColor.rgb == "FF00B050"


def test_traffic_census_preview_requires_wideload_upload(client, temp_store):
    report_id = create_report_session(client)
    patch_manual_inputs(client, report_id)

    response = client.get(
        f"/api/report-sessions/{report_id}/sections/traffic-census/preview?format=docx"
    )

    assert response.status_code == 400
    assert response.json()["detail"] == (
        "Traffic Census preview requires wideload upload first because E comes from wideload count."
    )


def test_summary_cards_report_awaiting_data_for_new_session(client, temp_store):
    report_id = create_report_session(client)

    response = client.get(f"/api/report-sessions/{report_id}/summary-cards")

    assert response.status_code == 200
    payload = response.json()
    assert [card["title"] for card in payload["cards"]] == [
        "Total Weighed",
        "Total Overloaded",
        "Special Released",
        "Wide Loads",
    ]
    assert all(card["value"] is None for card in payload["cards"])
    assert all(card["display_value"] == "—" for card in payload["cards"])
    assert all(card["status"] == "awaiting_data" for card in payload["cards"])
