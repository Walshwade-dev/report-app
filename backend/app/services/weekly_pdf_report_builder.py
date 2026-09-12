import io
from app.services.weekly_excel_report_builder import build_weekly_excel_report
from app.services.preview_renderer import convert_docx_to_pdf


def build_weekly_pdf_report(
    weekly_data_by_bound: dict[str, list[dict]],
    weekly_data_combined: list[dict],
    start_date: str,
    end_date: str,
    station: str,
    prepared_by: str,
    approved_by: str,
) -> io.BytesIO:
    """
    Builds the weekly report PDF by converting the meticulously styled
    weekly Excel workbook directly to PDF via LibreOffice headless Calc.
    Ensures 1-to-1 visual fidelity with the Excel report, including:
    - 11pt Arial table text with bold dates and totals
    - 10pt bold Arial signatures and footer table
    - Cell background colors (yellow highlights on H, C, Q, T; slate on dates & totals)
    - Precise row heights and column borders
    - Landscape fitToWidth=1 page setup
    """
    excel_buffer = build_weekly_excel_report(
        weekly_data_by_bound,
        weekly_data_combined,
        start_date,
        end_date,
        station,
        prepared_by,
        approved_by,
    )
    pdf_content, _ = convert_docx_to_pdf(excel_buffer, f"{station}_weekly_report.xlsx")
    return io.BytesIO(pdf_content)
