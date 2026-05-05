import pandas as pd
from docx import Document

from app.services.report_layout import apply_standard_layout
from app.services.daily_hour_processor import (
    build_daily_hour_metrics,
    add_daily_totals_row,
)
from app.services.daily_hour_generator import add_daily_hour_statistics_section


raw_df = pd.read_csv("/home/ace/Downloads/Daily Hour Statistics2026-4-4-21-59-51-291.csv")

daily_df = build_daily_hour_metrics(
    raw_df,
    report_date="02/02/2026",
    wideload_count=25,
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

doc.save("daily_hour_test.docx")
