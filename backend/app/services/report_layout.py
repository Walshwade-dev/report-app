from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Inches, Pt
from pathlib import Path
from datetime import datetime



LOGO_PATH = Path(__file__).resolve().parent.parent / "assets" / "logo.png"


def apply_standard_layout(doc, report_date, station, bound):
    section = doc.sections[0]

    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = Inches(11.69)
    section.page_height = Inches(8.27)

    section.top_margin = Inches(0.75)
    section.bottom_margin = Inches(0.55)
    section.left_margin = Inches(0.25)
    section.right_margin = Inches(0.25)

    section.header_distance = Inches(0.15)
    section.footer_distance = Inches(0.15)

    add_header(section)
    add_footer(section, report_date, station, bound)
    

def add_header(section):
    header = section.header
    paragraph = header.paragraphs[0]
    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

    if LOGO_PATH.exists():
        run = paragraph.add_run()
        run.add_picture(str(LOGO_PATH), width=Inches(1.5))


def format_report_date(date_str):
    dt = datetime.strptime(date_str, "%Y-%m-%d")

    day = dt.day
    suffix = "th" if 11 <= day <= 13 else {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")

    return dt.strftime(f"%-d{suffix} %B %Y")

def add_footer(section, report_date, station, bound):
    footer = section.footer
    paragraph = footer.paragraphs[0]
    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    # Format date (example: 02nd February 2026)
    formatted_date = format_report_date(report_date)

    footer_text = (
        "KeNHA/WB/MTCE/4339/2025        "
        f"{station.upper()} WEIGHBRIDGE {bound.upper()} DAILY REPORT {formatted_date}        "
        "Page {PAGE} of {TOTAL}"
    )

    run = paragraph.add_run(footer_text)
    run.bold = True
    run.font.size = Pt(7)
    