import unittest
from fastapi.testclient import TestClient
from app.main import app
from app.db.models import User
from app.core.security import create_access_token
from app.routes.weekly_reports import get_authenticated_user


class TestWeeklyStationLocking(unittest.TestCase):
    def test_kanyonyo_officer_fetching_mismatched_station_forbidden(self):
        mock_user = User(username="Tesk Kanyonyo", full_name="Team Kanyonyo", role="user", station="Kanyonyo")
        token = create_access_token(data={"sub": mock_user.username, "role": mock_user.role})

        client = TestClient(app)
        # Mock get_authenticated_user to return mock_user
        from unittest.mock import patch
        with patch("app.routes.weekly_reports.get_authenticated_user", return_value=mock_user):
            response = client.get(
                "/api/reports/weekly/generate",
                params={
                    "start_date": "2026-09-01",
                    "end_date": "2026-09-07",
                    "station": "JUJA",  # Officer is from Kanyonyo, asking for JUJA
                    "prepared_by": "Test Officer",
                    "approved_by": "Faith Njani",
                    "format": "excel",
                },
                headers={"Authorization": f"Bearer {token}"},
            )
            self.assertEqual(response.status_code, 403)
            self.assertIn("Access denied", response.json()["detail"])
            self.assertIn("KANYONYO", response.json()["detail"])

    def test_kanyonyo_officer_fetching_own_station_succeeds(self):
        mock_user = User(username="Tesk Kanyonyo", full_name="Team Kanyonyo", role="user", station="Kanyonyo")
        token = create_access_token(data={"sub": mock_user.username, "role": mock_user.role})

        client = TestClient(app)
        from unittest.mock import patch
        with patch("app.routes.weekly_reports.get_authenticated_user", return_value=mock_user):
            response = client.get(
                "/api/reports/weekly/generate",
                params={
                    "start_date": "2026-09-01",
                    "end_date": "2026-09-07",
                    "station": "KANYONYO",  # Matches officer station
                    "prepared_by": "Team Kanyonyo",
                    "approved_by": "Faith Njani",
                    "format": "excel",
                },
                headers={"Authorization": f"Bearer {token}"},
            )
            self.assertEqual(response.status_code, 200)
            self.assertEqual(response.headers["content-type"], "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.assertIn("KANYONYO", response.headers.get("content-disposition", ""))

    def test_admin_fetching_any_station_succeeds(self):
        mock_admin = User(username="admin", full_name="Admin", role="admin", station=None)
        token = create_access_token(data={"sub": mock_admin.username, "role": mock_admin.role})

        client = TestClient(app)
        from unittest.mock import patch
        with patch("app.routes.weekly_reports.get_authenticated_user", return_value=mock_admin):
            response = client.get(
                "/api/reports/weekly/generate",
                params={
                    "start_date": "2026-09-01",
                    "end_date": "2026-09-07",
                    "station": "ATHI RIVER",
                    "prepared_by": "Admin User",
                    "approved_by": "Faith Njani",
                    "format": "excel",
                },
                headers={"Authorization": f"Bearer {token}"},
            )
            self.assertEqual(response.status_code, 200)
            self.assertEqual(response.headers["content-type"], "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.assertIn("ATHI RIVER", response.headers.get("content-disposition", ""))


if __name__ == "__main__":
    unittest.main()
