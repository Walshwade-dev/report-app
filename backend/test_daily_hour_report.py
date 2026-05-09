from pathlib import Path

import pandas as pd
from docx import Document

from app.services.daily_hour_chart_generator import add_daily_hour_chart_section
from app.services.daily_hour_generator import add_daily_hour_statistics_section
from app.services.daily_hour_processor import (
    add_daily_totals_row,
    build_daily_hour_metrics,
)
from app.services.report_layout import apply_standard_layout


BASE_DIR = Path(__file__).resolve().parent
FIXTURE_PATH = BASE_DIR / "tests" / "fixtures" / "daily_hour.csv"
OUTPUT_PATH = BASE_DIR / "daily_hour_test.docx"


def main() -> None:
    if not FIXTURE_PATH.exists():
        raise SystemExit(f"Missing daily/hour fixture: {FIXTURE_PATH}")

    raw_df = pd.read_csv(FIXTURE_PATH)

    daily_df = build_daily_hour_metrics(
        raw_df,
        report_date="02/02/2026",
        wideload_count=1,
    )
    daily_df = add_daily_totals_row(daily_df)

    doc = Document()
    apply_standard_layout(
        doc,
        report_date="2026-02-02",
        station="Juja",
        bound="Thika Bound",
    )

    add_daily_hour_statistics_section(doc, daily_df)
    doc.add_page_break()
    add_daily_hour_chart_section(doc, daily_df)

    doc.save(OUTPUT_PATH)
    print(f"Saved daily/hour smoke report to {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
