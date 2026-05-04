import pandas as pd

from app.services.cleaner_core import clean_with_template
from app.templates import vehicle_inspection
from app.templates import impounded_prohibited
from app.services.final_report_builder import build_final_report


wideload_raw = pd.read_csv("/home/ace/Downloads/Wide Load2026-1-3-1-0-37-631.csv")
impounded_raw = pd.read_csv("/home/ace/Downloads/Impounded And Prohibited2026-1-3-1-1-10-32.csv")
overloaded_raw = pd.read_csv("/home/ace/Downloads/Impounded And Overloaded2026-1-3-1-1-24-592.csv")

wideload_df = clean_with_template(wideload_raw, vehicle_inspection)
impounded_df = clean_with_template(impounded_raw, impounded_prohibited)

file_stream = build_final_report(
    wideload_df=wideload_df,
    impounded_prohibited_df=impounded_df,
    overloaded_df=overloaded_raw,
    report_date="2026-02-02",
    station="Juja",
    bound="Thika Bound",
)

with open("final_report_test.docx", "wb") as f:
    f.write(file_stream.read())