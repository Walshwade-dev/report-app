import os
import unittest
from fastapi.testclient import TestClient
from app.main import app
from app.db.models import User
from app.routes.reports import check_write_permission

class TestBatchIntakeEndToEnd(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        mock_user = User(username="admin", role="admin")
        app.dependency_overrides[check_write_permission] = lambda: mock_user
        cls.client = TestClient(app)
        cls.fixtures_dir = os.path.join(os.path.dirname(__file__), "fixtures")

    @classmethod
    def tearDownClass(cls):
        app.dependency_overrides.pop(check_write_permission, None)

    def test_full_batch_intake_and_build(self):
        # 1. Create report session
        create_res = self.client.post(
            "/api/report-sessions",
            json={
                "weighbridge_name": "ATHI RIVER",
                "station": "Athi River",
                "bound": "Nairobi Bound",
                "bound_name": "NAIROBI BOUND",
                "report_date": "2026-09-06",
                "officer_in_charge": "Officer J. Doe",
            },
        )
        self.assertEqual(create_res.status_code, 200, create_res.text)
        session_data = create_res.json()
        report_id = session_data["report_id"]
        self.assertTrue(bool(report_id))

        # 2. Ingest 4 spreadsheet files in batch
        sections = [
            ("daily_hour.csv", "daily-hour", "daily_hour"),
            ("wideload.csv", "wideload", "wideload"),
            ("impounded_prohibited.csv", "impounded-prohibited", "impounded_prohibited"),
            ("overloaded.csv", "overloaded", "overloaded"),
        ]

        for fixture_file, endpoint, section_key in sections:
            path = os.path.join(self.fixtures_dir, fixture_file)
            self.assertTrue(os.path.exists(path), f"Fixture missing: {path}")
            with open(path, "rb") as f:
                upload_res = self.client.post(
                    f"/api/report-sessions/{report_id}/uploads/{endpoint}",
                    files={"file": (fixture_file, f, "text/csv")},
                )
            self.assertEqual(
                upload_res.status_code,
                200,
                f"Failed upload for {fixture_file}: {upload_res.text}",
            )
            data = upload_res.json()
            self.assertIn(section_key, data.get("sections", {}))
            self.assertEqual(data["sections"][section_key]["status"], "ready")

        # 3. Ingest Transgression scan document for OCR
        pdf_path = os.path.join(self.fixtures_dir, "transgression_nairobi.pdf")
        self.assertTrue(os.path.exists(pdf_path), f"Fixture missing: {pdf_path}")
        with open(pdf_path, "rb") as f:
            ocr_res = self.client.post(
                f"/api/report-sessions/{report_id}/transgressions/ocr-extract",
                files={"file": ("transgression_nairobi.pdf", f, "application/pdf")},
            )
        self.assertEqual(ocr_res.status_code, 200, ocr_res.text)
        ocr_data = ocr_res.json()
        self.assertTrue(ocr_data.get("success"), ocr_data)
        extracted = ocr_data.get("extracted", {})
        daily = extracted.get("daily_transgression", {})
        action = extracted.get("action_report", {})
        self.assertTrue(bool(daily.get("regNo")), "Plate regNo extracted")

        # 4. Save manual inputs with extracted transgression details
        save_res = self.client.patch(
            f"/api/report-sessions/{report_id}/manual-inputs",
            json={
                "prepared_by": "Fredrick Kariuki",
                "confirmed_by": "Faith Njani",
                "traffic_census": {
                    "buses_gte_3500kg": 5,
                    "vehicles_3500_to_7000_excluding_buses": 10,
                    "vehicles_gte_7000_excluding_buses": 15,
                    "total_traffic_census": 30,
                },
                "transgressions": {
                    "daily_transgressions": [daily],
                    "action_report": [action],
                },
            },
        )
        self.assertEqual(save_res.status_code, 200, save_res.text)

        # 5. Build Final Report (.docx)
        build_res = self.client.post(
            f"/api/report-sessions/{report_id}/build-final-report",
        )
        self.assertEqual(build_res.status_code, 200, build_res.text)
        build_data = build_res.json()
        self.assertIn(build_data.get("final_report", {}).get("status"), ["processing", "ready"])

        # Complete worker generation
        from app.services.report_worker import run_build_final_report
        run_build_final_report(report_id)

        # 6. Verify Final Report download (.docx)
        download_docx_res = self.client.get(
            f"/api/report-sessions/{report_id}/download-final-report",
        )
        self.assertEqual(download_docx_res.status_code, 200)
        self.assertTrue(len(download_docx_res.content) > 1000)

        # 7. Verify Excel Report download (.xlsx)
        download_excel_res = self.client.get(
            f"/api/report-sessions/{report_id}/download-excel-report",
        )
        self.assertEqual(download_excel_res.status_code, 200)
        self.assertTrue(len(download_excel_res.content) > 1000)


if __name__ == "__main__":
    unittest.main()
