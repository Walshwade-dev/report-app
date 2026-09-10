import unittest
from fastapi.testclient import TestClient
from app.main import app
from app.db.models import User
from app.routes.reports import check_write_permission


class TestMobileStationLocking(unittest.TestCase):
    def test_kanyonyo_officer_mobile_session_locks_to_kanyonyo_mobile(self):
        mock_user = User(username="Tesk Kanyonyo", full_name="Team Kanyonyo", role="user", station="Kanyonyo")
        app.dependency_overrides[check_write_permission] = lambda: mock_user

        try:
            client = TestClient(app)
            response = client.post(
                "/api/report-sessions",
                json={
                    "report_date": "2026-09-10",
                    "station": "Juja mobile",  # non-admin submitted this
                    "bound": "Mobile 2",
                    "weighbridge_name": "Juja mobile",
                },
            )
            self.assertEqual(response.status_code, 200)
            data = response.json()
            self.assertEqual(data["metadata"]["station"], "Kanyonyo mobile")
            self.assertEqual(data["metadata"]["weighbridge_name"], "Kanyonyo mobile")
            self.assertEqual(data["metadata"]["prepared_by"], "Team Kanyonyo")

            # Update metadata also preserves mobile station locking
            report_id = data["report_id"]
            patch_resp = client.patch(
                f"/api/report-sessions/{report_id}/metadata",
                json={"station": "Suswa mobile", "bound": "Mobile 1"},
            )
            self.assertEqual(patch_resp.status_code, 200)
            patch_data = patch_resp.json()
            self.assertEqual(patch_data["metadata"]["station"], "Kanyonyo mobile")
            self.assertEqual(patch_data["metadata"]["weighbridge_name"], "Kanyonyo mobile")
        finally:
            app.dependency_overrides.pop(check_write_permission, None)

    def test_isinya_officer_mobile_session_locks_to_isinya_mobile(self):
        mock_user = User(username="isinya_officer", full_name="Isinya Officer", role="user", station="Isinya")
        app.dependency_overrides[check_write_permission] = lambda: mock_user

        try:
            client = TestClient(app)
            response = client.post(
                "/api/report-sessions",
                json={
                    "report_date": "2026-09-10",
                    "station": "Juja mobile",
                    "bound": "Mobile 1",
                    "weighbridge_name": "Juja mobile",
                },
            )
            self.assertEqual(response.status_code, 200)
            data = response.json()
            self.assertEqual(data["metadata"]["station"], "Isinya mobile")
            self.assertEqual(data["metadata"]["weighbridge_name"], "Isinya mobile")
        finally:
            app.dependency_overrides.pop(check_write_permission, None)

    def test_static_session_locks_to_clean_station_without_mobile(self):
        mock_user = User(username="Tesk Kanyonyo", full_name="Team Kanyonyo", role="user", station="Kanyonyo")
        app.dependency_overrides[check_write_permission] = lambda: mock_user

        try:
            client = TestClient(app)
            response = client.post(
                "/api/report-sessions",
                json={
                    "report_date": "2026-09-10",
                    "station": "JUJA",
                    "bound": "NAIROBI BOUND",
                    "weighbridge_name": "JUJA",
                },
            )
            self.assertEqual(response.status_code, 200)
            data = response.json()
            self.assertEqual(data["metadata"]["station"], "Kanyonyo")
            self.assertEqual(data["metadata"]["weighbridge_name"], "Kanyonyo")
        finally:
            app.dependency_overrides.pop(check_write_permission, None)

    def test_admin_is_not_locked(self):
        mock_admin = User(username="admin", full_name="Admin", role="admin", station=None)
        app.dependency_overrides[check_write_permission] = lambda: mock_admin

        try:
            client = TestClient(app)
            response = client.post(
                "/api/report-sessions",
                json={
                    "report_date": "2026-09-10",
                    "station": "Athiriver mobile",
                    "bound": "Mobile 2",
                    "weighbridge_name": "Athiriver mobile",
                },
            )
            self.assertEqual(response.status_code, 200)
            data = response.json()
            self.assertEqual(data["metadata"]["station"], "Athiriver mobile")
        finally:
            app.dependency_overrides.pop(check_write_permission, None)


if __name__ == "__main__":
    unittest.main()
