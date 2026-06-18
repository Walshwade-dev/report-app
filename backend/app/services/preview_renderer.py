import io
import os
import subprocess
import tempfile
from dataclasses import dataclass
from pathlib import Path

from docx import Document

from app.services.daily_hour_chart_generator import add_daily_hour_chart_section
from app.services.daily_hour_generator import add_daily_hour_statistics_section
from app.services.daily_summary_generator import add_daily_summary_section
from app.services.daily_summary_processor import build_daily_summary_from_session
from app.services.impounded_prohibited_generator import add_impounded_prohibited_section
from app.services.report_layout import apply_standard_layout
from app.services.report_session_store import ReportSession
from app.services.report_session_metrics import (
    get_traffic_census_section_values,
)
from app.services.traffic_census_generator import add_traffic_census_section
from app.services.transgressions_generator import add_transgressions_section
from app.services.wideload_generator import add_wideload_section


PREVIEW_ALIASES = {
    "daily-hour": "daily_hour",
    "daily_hour": "daily_hour",
    "daily-hour-statistics": "daily_hour",
    "daily_hour_statistics": "daily_hour",
    "daily-hour-chart": "daily_hour",
    "daily_hour_chart": "daily_hour",
    "traffic-census": "traffic_census",
    "traffic_census": "traffic_census",
    "daily-summary": "daily_summary",
    "daily_summary": "daily_summary",
    "transgressions": "transgressions",
    "wideload": "wideload",
    "wide-load": "wideload",
    "impounded-prohibited": "impounded_prohibited",
    "impounded_prohibited": "impounded_prohibited",
}


@dataclass
class PreviewResult:
    stream: io.BytesIO
    filename: str
    media_type: str
    inline: bool = True


def normalize_preview_section(section_name: str) -> str:
    normalized = section_name.strip().lower()
    normalized = PREVIEW_ALIASES.get(normalized, normalized.replace("-", "_"))
    return normalized


def default_preview_page(section_name: str) -> int:
    normalized = section_name.strip().lower()

    if normalized in {"daily-hour-chart", "daily_hour_chart"}:
        return 2

    return 1


def is_section_ready_for_preview(session: ReportSession, section: str) -> bool:
    if session.sections.get(section, {}).get("status") != "ready":
        return False

    if section == "traffic_census":
        return "traffic_census" in session.manual_inputs

    if section == "daily_summary":
        return session.sections.get(section, {}).get("status") == "ready"

    if section == "transgressions":
        return "transgressions" in session.manual_inputs

    return section in session.dataframes


def build_section_preview_docx(
    session: ReportSession,
    section_name: str,
) -> tuple[io.BytesIO, str]:
    section = normalize_preview_section(section_name)

    if not is_section_ready_for_preview(session, section):
        raise ValueError(f"Section is not ready for preview: {section_name}")

    doc = Document()
    apply_standard_layout(
        doc,
        report_date=session.report_date,
        station=session.station,
        bound=session.bound,
    )

    if section == "daily_hour":
        daily_df = session.dataframes[section]
        add_daily_hour_statistics_section(doc, daily_df)
        doc.add_page_break()
        add_daily_hour_chart_section(doc, daily_df, is_preview=True)
        filename = "daily_hour_preview.docx"

    elif section == "wideload":
        add_wideload_section(doc, session.dataframes[section])
        filename = "wideload_preview.docx"

    elif section == "impounded_prohibited":
        add_impounded_prohibited_section(doc, session.dataframes[section])
        filename = "impounded_prohibited_preview.docx"

    elif section == "traffic_census":
        traffic_values = get_traffic_census_section_values(session)

        add_traffic_census_section(
            doc,
            session.manual_inputs["traffic_census"],
            exemption_not_weighed=traffic_values["e"],
            total_weighed=traffic_values["x"],
            hswim_cleared=traffic_values["q"],
            total_traffic=traffic_values["total_traffic"],
        )

        filename = "traffic_census_preview.docx"

    elif section == "daily_summary":
        add_daily_summary_section(doc, build_daily_summary_from_session(session))
        filename = "daily_summary_preview.docx"

    elif section == "transgressions":
        add_transgressions_section(doc, session.manual_inputs["transgressions"])
        filename = "transgressions_preview.docx"

    else:
        raise ValueError(f"No preview renderer is available for section: {section_name}")

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    return buffer, filename


def preview_filename(section_name: str, preview_format: str, page: int | None = None) -> str:
    section = normalize_preview_section(section_name)
    extension = preview_format.strip().lower()
    page_suffix = f"_page_{page}" if extension == "png" and page else ""
    return f"{section}_preview{page_suffix}.{extension}"


def convert_docx_to_pdf(docx_stream: io.BytesIO, filename: str) -> tuple[bytes, str]:
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_path = Path(temp_dir)
        runtime_dir = temp_path / "runtime"
        runtime_dir.mkdir(mode=0o700)
        config_dir = temp_path / "config"
        cache_dir = temp_path / "cache"
        home_dir = temp_path / "home"
        profile_dir = temp_path / "lo_profile"

        for directory in [config_dir, cache_dir, home_dir, profile_dir]:
            directory.mkdir()

        env = os.environ.copy()
        env.update(
            {
                "HOME": str(home_dir),
                "XDG_CONFIG_HOME": str(config_dir),
                "XDG_CACHE_HOME": str(cache_dir),
                "XDG_RUNTIME_DIR": str(runtime_dir),
            }
        )

        docx_path = temp_path / filename
        docx_path.write_bytes(docx_stream.getvalue())

        result = subprocess.run(
            [
                "libreoffice",
                "--headless",
                f"-env:UserInstallation=file://{profile_dir}",
                "--convert-to",
                "pdf",
                "--outdir",
                str(temp_path),
                str(docx_path),
            ],
            capture_output=True,
            check=False,
            env=env,
            text=True,
            timeout=60,
        )

        pdf_path = docx_path.with_suffix(".pdf")

        if pdf_path.exists():
            return pdf_path.read_bytes(), pdf_path.name

        if result.returncode != 0:
            message = result.stderr.strip() or result.stdout.strip()
            raise RuntimeError(f"Failed to convert preview to PDF: {message}")

        if not pdf_path.exists():
            raise RuntimeError("Failed to convert preview to PDF: output file was not created")

        return pdf_path.read_bytes(), pdf_path.name


def convert_pdf_to_png(pdf_content: bytes, filename: str, page: int) -> tuple[bytes, str]:
    if page < 1:
        raise ValueError("Preview page must be greater than or equal to 1")

    with tempfile.TemporaryDirectory() as temp_dir:
        temp_path = Path(temp_dir)
        pdf_path = temp_path / filename
        output_prefix = temp_path / pdf_path.stem
        pdf_path.write_bytes(pdf_content)

        result = subprocess.run(
            [
                "pdftoppm",
                "-f",
                str(page),
                "-l",
                str(page),
                "-singlefile",
                "-png",
                "-r",
                "140",
                str(pdf_path),
                str(output_prefix),
            ],
            capture_output=True,
            check=False,
            text=True,
            timeout=60,
        )

        if result.returncode != 0:
            message = result.stderr.strip() or result.stdout.strip()
            raise RuntimeError(f"Failed to convert preview to PNG: {message}")

        png_path = output_prefix.with_suffix(".png")

        if not png_path.exists():
            raise RuntimeError("Failed to convert preview to PNG: output file was not created")

        return png_path.read_bytes(), png_path.name


def build_section_preview(
    session: ReportSession,
    section_name: str,
    preview_format: str = "png",
    page: int | None = None,
) -> PreviewResult:
    output_format = preview_format.strip().lower()
    page_number = page or default_preview_page(section_name)
    docx_stream, docx_filename = build_section_preview_docx(session, section_name)

    if output_format == "docx":
        return PreviewResult(
            stream=docx_stream,
            filename=docx_filename,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            inline=False,
        )

    if output_format == "pdf":
        pdf_content, pdf_filename = convert_docx_to_pdf(docx_stream, docx_filename)
        return PreviewResult(
            stream=io.BytesIO(pdf_content),
            filename=pdf_filename,
            media_type="application/pdf",
        )

    if output_format == "png":
        pdf_content, pdf_filename = convert_docx_to_pdf(docx_stream, docx_filename)
        png_content, png_filename = convert_pdf_to_png(pdf_content, pdf_filename, page_number)
        return PreviewResult(
            stream=io.BytesIO(png_content),
            filename=png_filename,
            media_type="image/png",
        )

    raise ValueError("Unsupported preview format. Use one of: png, pdf, docx")


def get_cached_section_preview(
    session: ReportSession,
    section_name: str,
    preview_format: str,
    page: int | None,
    store,
) -> PreviewResult:
    output_format = preview_format.strip().lower()
    section = normalize_preview_section(section_name)
    page_number = page or default_preview_page(section_name)

    if output_format not in {"png", "pdf", "docx"}:
        raise ValueError("Unsupported preview format. Use one of: png, pdf, docx")

    if not is_section_ready_for_preview(session, section):
        raise ValueError(f"Section is not ready for preview: {section_name}")

    if section == "traffic_census":
        get_traffic_census_section_values(session)

    cache_page = page_number if output_format == "png" else None
    cached_content = store.read_cached_preview(
        session.report_id,
        section,
        output_format,
        page=cache_page,
    )

    if cached_content is not None:
        return PreviewResult(
            stream=io.BytesIO(cached_content),
            filename=preview_filename(section, output_format, cache_page),
            media_type={
                "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "pdf": "application/pdf",
                "png": "image/png",
            }[output_format],
            inline=output_format != "docx",
        )

    docx_content = store.read_cached_preview(session.report_id, section, "docx")
    docx_filename = preview_filename(section, "docx")

    if docx_content is None:
        docx_stream, docx_filename = build_section_preview_docx(session, section)
        docx_content = docx_stream.getvalue()
        store.write_cached_preview(session.report_id, section, "docx", docx_content)

    if output_format == "docx":
        return PreviewResult(
            stream=io.BytesIO(docx_content),
            filename=docx_filename,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            inline=False,
        )

    pdf_content = store.read_cached_preview(session.report_id, section, "pdf")
    pdf_filename = preview_filename(section, "pdf")

    if pdf_content is None:
        pdf_content, pdf_filename = convert_docx_to_pdf(
            io.BytesIO(docx_content),
            docx_filename,
        )
        store.write_cached_preview(session.report_id, section, "pdf", pdf_content)

    if output_format == "pdf":
        return PreviewResult(
            stream=io.BytesIO(pdf_content),
            filename=pdf_filename,
            media_type="application/pdf",
        )

    png_content, _ = convert_pdf_to_png(pdf_content, pdf_filename, page_number)
    png_filename = preview_filename(section, "png", page_number)
    store.write_cached_preview(
        session.report_id,
        section,
        "png",
        png_content,
        page=page_number,
    )
    return PreviewResult(
        stream=io.BytesIO(png_content),
        filename=png_filename,
        media_type="image/png",
    )
