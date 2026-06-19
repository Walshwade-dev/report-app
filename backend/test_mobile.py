import sys
import pandas as pd
import io
sys.path.append('.')

from app.services.mobile_report_processor import normalize_mobile_report, summarize_mobile_report
from app.services.mobile_excel_report_builder import build_mobile_excel_report
from app.services.mobile_word_report_builder import build_mobile_word_report

class MockSession:
    def __init__(self, df):
        self.dataframes = {"mobile_report": df}
        self.sections = {"mobile_report": {"status": "ready"}}
        self.manual_inputs = {
            "cases_cleared_in_court": 1,
            "danka_staff": "DM DUNCAN ODHIAMBO/DM ELIZABETH CHARI/DRIVER SILAS MWANGI",
            "police_officers": "CPL EMASON SAUTET / PC (w) JANE EKADELI",
            "mileage_start": "61,267",
            "mileage_end": "61,447",
            "mobile_vehicle": "KDS042Z"
        }
        self.report_date = "2026-05-06"
        self.station = "Juja"
        self.bound = "Thika"
        self.prepared_by = "Fredrick Kariuki"
        self.confirmed_by = "Faith Njani"

with open('/home/ace/.gemini/antigravity/scratch/mobile reports/Weighbridgeregister20260512061227966 edited.csv', 'rb') as f:
    df = pd.read_csv(io.BytesIO(f.read()))

records = normalize_mobile_report(df)
session = MockSession(records)

excel_buffer = build_mobile_excel_report(session)
with open('generated_excel.xlsx', 'wb') as f:
    f.write(excel_buffer.read())

word_buffer = build_mobile_word_report(session)
with open('generated_word.docx', 'wb') as f:
    f.write(word_buffer.read())

print("Done")
