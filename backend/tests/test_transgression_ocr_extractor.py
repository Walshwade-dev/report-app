import os
import unittest
from pathlib import Path
from unittest.mock import patch

from fastapi.testclient import TestClient

from app.main import app
from app.services.report_session_store import report_session_store
from app.services.transgression_ocr_extractor import (
    extract_transgression_from_file_bytes,
    parse_transgression_text,
)

SAMPLE_OCR_TEXT = """
KENYA NATIONAL HIGHWAYS AUTHORITY
Tag Ticket
Print Date: 2026-09-06T 11:30:06
Weighbridge Name: JURU
TAG ID: TAGJURU202693218384
Tagged Date & Time: 2026-09-06T11:28:55
Vehicle Registration No.: KDG096Q
Opened By: dnjoroge

Opening Reason:
Transgression using Bypass Lane JU-RU-L3 at JUJA RUIRU BOUND HSWIM on 06/09/2026 at 11:01:17 AM.
Flagged down by police officers in uniform but failed to stop. Chased, found, flagged to stop, declined and proceeded.
OB NO: 13/06/09/2026 at 1115hrs.

Transporter: JAMES GATHENYA
Axle Config: 2A
Officers On Duty: SGT ANGELO MBOGORI
"""


class TestTransgressionOcrExtractor(unittest.TestCase):

    def test_parse_transgression_text_complete(self):
        result = parse_transgression_text(SAMPLE_OCR_TEXT, fallback_station="JUJA WEIGHBRIDGE")

        daily = result["daily_transgression"]
        action = result["action_report"]

        # Check Daily Transgressions fields
        self.assertEqual(daily["regNo"], "KDG 096Q")
        self.assertEqual(daily["date"], "06/09/2026")
        self.assertEqual(daily["time"], "1101hrs")
        self.assertEqual(daily["axleConfig"], "2A")
        self.assertEqual(daily["transporter"], "JAMES GATHENYA")
        self.assertEqual(daily["policeInCharge"], "SGT ANGELO MBOGORI")
        self.assertEqual(daily["caught"], "NO")
        self.assertEqual(daily["actionTaken"], "chased not found")
        self.assertEqual(daily["nextWbReportSent"], "-")
        self.assertEqual(daily["nextWb"], "-")

        # Check Action Report fields
        self.assertEqual(action["truckNo"], "KDG 096Q")
        self.assertEqual(action["date"], "06/09/2026")
        self.assertEqual(action["timeReceived"], "1101hrs")
        self.assertEqual(action["taggedInSystem"], "YES")
        self.assertEqual(action["action1"], "Tagged in System (TAG ID: TAGJURU202693218384)")
        self.assertEqual(action["ocsReportedTo"], "YES")
        self.assertEqual(action["attachEvidence"], "TAGJURU202693218384, 13/06/09/2026")
        self.assertEqual(action["weightNoted"], "NO")

    def test_unsupported_file_extension(self):
        with self.assertRaises(ValueError):
            extract_transgression_from_file_bytes(b"dummy", filename="document.txt")

    @patch("app.services.transgression_ocr_extractor.extract_raw_text_from_file")
    def test_extract_from_file_bytes_mocked(self, mock_extract):
        mock_extract.return_value = SAMPLE_OCR_TEXT
        result = extract_transgression_from_file_bytes(b"dummy_content", filename="incident.png")

        self.assertTrue(result["success"])
        self.assertEqual(result["filename"], "incident.png")
        self.assertEqual(result["extracted"]["daily_transgression"]["regNo"], "KDG 096Q")
        self.assertEqual(result["extracted"]["action_report"]["taggedInSystem"], "YES")

    def test_real_pdf_fixture_if_present(self):
        sample_path = Path("/home/ace/.gemini/antigravity-ide/brain/c44e42cd-08b6-45e9-86ca-f612617c2487/scratch/transgression_nairobi.pdf")
        if not sample_path.exists():
            self.skipTest("Sample PDF fixture not found")

        with open(sample_path, "rb") as f:
            content = f.read()

        result = extract_transgression_from_file_bytes(content, filename="transgression_nairobi.pdf")
        self.assertTrue(result["success"])
        self.assertEqual(result["extracted"]["daily_transgression"]["regNo"], "KDG 096Q")
        self.assertEqual(result["extracted"]["daily_transgression"]["date"], "06/09/2026")
        self.assertEqual(result["extracted"]["action_report"]["taggedInSystem"], "YES")

    def test_api_endpoint(self):
        from app.db.models import User
        from app.routes.reports import check_write_permission

        session = report_session_store.create(
            report_date="2026-09-06",
            station="Juja",
            bound="Nairobi Bound",
            weighbridge_name="Juja Weighbridge",
        )
        report_id = session.report_id

        # Override permission check for unit test
        mock_user = User(username="admin", role="admin")
        app.dependency_overrides[check_write_permission] = lambda: mock_user

        try:
            client = TestClient(app)
            with patch("app.services.transgression_ocr_extractor.extract_raw_text_from_file", return_value=SAMPLE_OCR_TEXT):
                response = client.post(
                    f"/api/report-sessions/{report_id}/transgressions/ocr-extract",
                    files={"file": ("incident.pdf", b"fake_pdf_content", "application/pdf")},
                )

            self.assertEqual(response.status_code, 200)
            payload = response.json()
            self.assertTrue(payload["success"])
            self.assertEqual(payload["extracted"]["daily_transgression"]["regNo"], "KDG 096Q")
            self.assertEqual(payload["extracted"]["action_report"]["taggedInSystem"], "YES")

            # Test reset endpoint resets the session manual inputs and transgressions
            reset_resp = client.post(f"/api/report-sessions/{report_id}/reset")
            self.assertEqual(reset_resp.status_code, 200)
            reset_data = reset_resp.json()
            self.assertEqual(reset_data["manual_inputs"], {})
        finally:
            app.dependency_overrides.pop(check_write_permission, None)


if __name__ == "__main__":
    unittest.main()
