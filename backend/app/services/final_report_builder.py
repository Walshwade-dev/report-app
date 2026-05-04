from docx import Document
import io
import pandas as pd

from app.services.report_layout import apply_standard_layout
from app.services.report_context import ReportContext
from app.services.wideload_generator import add_wideload_section
from app.services.impounded_prohibited_generator import add_impounded_prohibited_section
from app.services.overloaded_summary import count_valid_permit_vehicles


def build_final_report(
    wideload_df: pd.DataFrame,
    impounded_prohibited_df: pd.DataFrame,
    overloaded_df: pd.DataFrame,
    report_date: str,
    station: str,
    bound: str,
) -> io.BytesIO:
    context = ReportContext(
        report_date=report_date,
        station=station,
        bound=bound,
        wideload_count=len(wideload_df),
        overloaded_valid_permit_count=count_valid_permit_vehicles(overloaded_df),
    )

    doc = Document()

    apply_standard_layout(
        doc,
        report_date=context.report_date,
        station=context.station,
        bound=context.bound,
    )

    add_impounded_prohibited_section(doc, impounded_prohibited_df)

    doc.add_page_break()

    add_wideload_section(doc, wideload_df)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    return buffer