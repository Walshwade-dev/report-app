from pathlib import Path

import pandas as pd

from app.services.cleaner_core import clean_with_template
from app.services.daily_hour_processor import (
    add_daily_totals_row,
    build_daily_hour_metrics,
)
from app.services.daily_summary_processor import build_daily_summary
from app.services.final_report_builder import build_final_report
from app.services.overloaded_summary import count_valid_permit_vehicles
from app.templates import impounded_prohibited, vehicle_inspection


BASE_DIR = Path(__file__).resolve().parent
FIXTURES_DIR = BASE_DIR / "tests" / "fixtures"
OUTPUT_PATH = BASE_DIR / "final_report_test.docx"


def fixture_path(filename: str) -> Path:
    path = FIXTURES_DIR / filename
    if not path.exists():
        raise SystemExit(f"Missing test fixture: {path}")
    return path


def main() -> None:
    daily_hour_raw = pd.read_csv(fixture_path("daily_hour.csv"))
    wideload_raw = pd.read_csv(fixture_path("wideload.csv"))
    impounded_raw = pd.read_csv(fixture_path("impounded_prohibited.csv"))
    overloaded_raw = pd.read_csv(fixture_path("overloaded.csv"))

    daily_df = build_daily_hour_metrics(
        daily_hour_raw,
        report_date="02/02/2026",
        wideload_count=1,
    )
    daily_df = add_daily_totals_row(daily_df)
    wideload_df = clean_with_template(wideload_raw, vehicle_inspection)
    impounded_df = clean_with_template(impounded_raw, impounded_prohibited)
    traffic_census = {
        "buses_gte_3500kg": 1351,
        "vehicles_3500_to_7000_excluding_buses": 29,
        "vehicles_gte_7000_excluding_buses": 6,
        "total_traffic_census": 1386,
    }
    transgressions = {
        "daily_transgressions": [],
        "action_report": [],
    }
    daily_summary = build_daily_summary(
        daily_df=daily_df,
        traffic_census=traffic_census,
        overloaded_valid_permit_count=count_valid_permit_vehicles(overloaded_raw),
        manual_inputs={"transgressions": transgressions},
    )

    file_stream = build_final_report(
        daily_df=daily_df,
        wideload_df=wideload_df,
        impounded_prohibited_df=impounded_df,
        overloaded_df=overloaded_raw,
        report_date="2026-02-02",
        station="Juja",
        bound="Thika Bound",
        prepared_by="Fredrick Kariuki",
        confirmed_by="Faith Njani",
        traffic_census=traffic_census,
        daily_summary=daily_summary,
        transgressions=transgressions,
    )

    OUTPUT_PATH.write_bytes(file_stream.read())
    print(f"Saved final smoke report to {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
