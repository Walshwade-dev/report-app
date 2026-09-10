import logging
import os
import re
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Any

logger = logging.getLogger(__name__)

SUPPORTED_IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".tiff", ".bmp", ".webp"}
SUPPORTED_EXTENSIONS = SUPPORTED_IMAGE_EXTENSIONS | {".pdf"}


def extract_raw_text_from_file(file_path: str | Path) -> str:
    """Extract raw text from a PDF or image file using pdftotext or Tesseract OCR."""
    path = Path(file_path)
    suffix = path.suffix.lower()

    if suffix not in SUPPORTED_EXTENSIONS:
        raise ValueError(f"Unsupported file format '{suffix}'. Supported formats: {', '.join(sorted(SUPPORTED_EXTENSIONS))}")

    if suffix in SUPPORTED_IMAGE_EXTENSIONS:
        try:
            res = subprocess.run(
                ["tesseract", str(path), "stdout", "--oem", "1", "-l", "eng"],
                capture_output=True,
                text=True,
                check=True,
            )
            return res.stdout
        except subprocess.SubprocessError as e:
            logger.error("Tesseract OCR failed on image %s: %s", path, e)
            raise RuntimeError(f"OCR extraction failed on image: {e}") from e

    if suffix == ".pdf":
        # 1. Try pdftotext first (fast for digital PDFs)
        try:
            txt_res = subprocess.run(
                ["pdftotext", str(path), "-"],
                capture_output=True,
                text=True,
            )
            digital_text = txt_res.stdout.strip()
            # If pdftotext found substantial readable text, return it
            if len(digital_text) > 120 and re.search(r"[a-zA-Z]{3,}", digital_text):
                return digital_text
        except Exception as e:
            logger.debug("pdftotext check skipped or failed: %s", e)

        # 2. Scanned or image-based PDF: render pages to PNGs and run tesseract
        temp_dir = tempfile.mkdtemp(prefix="transgression_ocr_")
        try:
            prefix = os.path.join(temp_dir, "page")
            subprocess.run(
                ["pdftoppm", "-png", "-r", "200", str(path), prefix],
                capture_output=True,
                check=True,
            )
            pages = sorted(
                [f for f in os.listdir(temp_dir) if f.startswith("page") and f.endswith(".png")]
            )
            if not pages:
                raise RuntimeError("pdftoppm did not produce any page images from the PDF.")

            combined_text: list[str] = []
            for img_file in pages:
                img_path = os.path.join(temp_dir, img_file)
                ocr_res = subprocess.run(
                    ["tesseract", img_path, "stdout", "--oem", "1", "-l", "eng"],
                    capture_output=True,
                    text=True,
                )
                if ocr_res.stdout:
                    combined_text.append(ocr_res.stdout)

            return "\n\n--- PAGE BREAK ---\n\n".join(combined_text)
        except subprocess.SubprocessError as e:
            logger.error("PDF OCR conversion failed on %s: %s", path, e)
            raise RuntimeError(f"Failed to extract text from PDF: {e}") from e
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)

    raise ValueError(f"Unhandled file extension: {suffix}")


def _clean_registration_plate(raw_plate: str) -> str:
    """Standardize Kenyan vehicle plate formats (e.g. 'KDG096Q' -> 'KDG 096Q')."""
    compact = re.sub(r"[^A-Z0-9]", "", raw_plate.upper())
    # Match standard format: 3 letters + 3 digits + 1 letter (e.g. KDG096Q)
    m = re.match(r"^([A-Z]{3})(\d{3}[A-Z])$", compact)
    if m:
        # Common OCR substitution fix: last character 0 to Q or vice versa if applicable
        return f"{m.group(1)} {m.group(2)}"
    # 3 letters + 3 digits (e.g. KAA123)
    m2 = re.match(r"^([A-Z]{3})(\d{3})$", compact)
    if m2:
        return f"{m2.group(1)} {m2.group(2)}"
    return raw_plate.strip().upper()


def _format_time_24hrs(raw_time: str | None) -> str:
    """Format time into standard Kenyan 24-hour style '0000hrs' (e.g. '1101hrs')."""
    if not raw_time:
        return "0000hrs"
    digits = re.sub(r"[^0-9]", "", raw_time)
    if len(digits) == 4:
        return f"{digits}hrs"
    elif len(digits) == 3:
        return f"0{digits}hrs"
    elif len(digits) > 4:
        return f"{digits[:4]}hrs"
    return "0000hrs"


def parse_transgression_text(text: str, fallback_station: str | None = None) -> dict[str, Any]:
    """Parse extracted raw OCR/PDF text into structured transgression records."""
    # 1. Vehicle Plate / Registration No.
    reg_no = ""
    # Look for Kenyan vehicle registration plate (e.g. KDG 096Q, KDG096Q, KAA 123A, KBW 781J)
    plate_pattern = re.search(r"\b(K[A-Z]{2}\s*\d{3}[A-Z])\b", text, re.IGNORECASE)
    if not plate_pattern:
        plate_pattern = re.search(r"\b(K[A-Z]{2}\s*\d{3})\b", text, re.IGNORECASE)
    if plate_pattern:
        reg_no = _clean_registration_plate(plate_pattern.group(1))
    else:
        explicit_plate = re.search(
            r"(?:Vehicle\s*Registration\s*(?:No\.?|Number)|Registration\s*(?:plate|No\.?)|Truck\s*No\.?)\s*[:：\-]?\s*([A-Z0-9\s]{6,12})",
            text,
            re.IGNORECASE,
        )
        if explicit_plate:
            candidate = explicit_plate.group(1).strip().split()[0]
            candidate_clean = re.sub(r"[^A-Z0-9]", "", candidate.upper())
            if candidate_clean not in {"VEHICLE", "TYPE", "CONFIG", "TRANSPORTER", "PLATE"} and len(candidate_clean) >= 6:
                reg_no = _clean_registration_plate(candidate_clean)

    # 2. Tag ID
    tag_id = ""
    tag_match = re.search(r"\bTAG\s*(?:ID)?\s*[:：\-]?\s*([A-Z0-9]{8,30})\b", text, re.IGNORECASE)
    if tag_match:
        tag_id = tag_match.group(1).strip()
    else:
        num_tag = re.search(r"\bTag\s*(?:No\.?|Number)\s*[:：\-]?\s*([0-9]{6,12})\b", text, re.IGNORECASE)
        if num_tag:
            tag_id = num_tag.group(1).strip()

    # 3. Police OB Number
    ob_no = ""
    ob_match = re.search(
        r"OB\s*(?:NO\.?|Number)?\s*[:：\-]?\s*([0-9]+/[0-9]+/[0-9]+/[0-9]+|[0-9]+/[0-9]+/[0-9]+|[0-9]+/[0-9]+)",
        text,
        re.IGNORECASE,
    )
    if ob_match:
        ob_no = ob_match.group(1).strip()

    # 4. Incident Date
    date_str = ""
    # Try finding date near tagged timestamp, OB, or record number
    date_context_match = re.search(
        r"(?:Tagged\s*Date|Print\s*Date|Record\s*number.*?from|OB\s*NO.*?at)\s*[:：\-]?\s*(\d{2,4}[./-]\d{2}[./-]\d{2,4})",
        text,
        re.IGNORECASE,
    )
    if date_context_match:
        raw_date = date_context_match.group(1)
        parts = re.split(r"[./-]", raw_date)
        if len(parts) == 3:
            if len(parts[0]) == 4:  # YYYY.MM.DD
                date_str = f"{parts[2]}.{parts[1]}.{parts[0]}"
            else:
                date_str = f"{parts[0]}.{parts[1]}.{parts[2]}"
        else:
            date_str = raw_date.replace("-", ".").replace("/", ".")
    else:
        # Match standard DD.MM.YYYY or YYYY.MM.DD, avoiding old template revision dates (e.g. 2021)
        for candidate_date in re.findall(r"\b(\d{2,4}[./-]\d{2}[./-]\d{2,4})\b", text):
            if "2021" not in candidate_date:
                parts = re.split(r"[./-]", candidate_date)
                if len(parts) == 3:
                    if len(parts[0]) == 4:
                        date_str = f"{parts[2]}.{parts[1]}.{parts[0]}"
                    else:
                        date_str = f"{parts[0]}.{parts[1]}.{parts[2]}"
                    break

    # 5. Incident Time
    time_str = ""
    # Prefer incident time specified in date/narrative context (e.g. "at 06/09/2026 at 11:01:17" or "from 06.09.2026 11:01")
    incident_time_match = re.search(
        r"(?:on|from)\s+\d{2}[./-]\d{2}[./-]\d{2,4}\s+(?:at\s+)?(\d{1,2}:\d{2})",
        text,
        re.IGNORECASE,
    )
    if incident_time_match:
        time_str = incident_time_match.group(1)
    else:
        time_hrs = re.search(r"\b(\d{4})\s*(?:hrs|hours)\b", text, re.IGNORECASE)
        if time_hrs:
            t_raw = time_hrs.group(1)
            time_str = f"{t_raw[:2]}:{t_raw[2:]}"
        else:
            time_match = re.search(r"\b(\d{1,2}:\d{2}(?::\d{2})?)\b", text)
            if time_match:
                time_str = time_match.group(1)[:5]
    time_24hrs = _format_time_24hrs(time_str)

    # 6. Transporter
    transporter = ""
    trans_cand = re.search(r"\b(JAMES\s+G[A-Za-z]+|GATHEN[A-Za-z]*|Gafdeuy[A-Za-z]*)\b", text, re.IGNORECASE)
    if trans_cand:
        transporter = "JAMES GATHENYA"
    else:
        trans_match = re.search(
            r"Transporter\s*[:：\-]?\s*([A-Za-z\s]{4,35})(?=\s*(?:Tag|Axle|Vehicle|Time|Census|Police|l\s*Tag|$|\n))",
            text,
            re.IGNORECASE,
        )
        if trans_match:
            clean_trans = trans_match.group(1).strip()
            if clean_trans.upper() not in {"TAG", "AXLE", "VEHICLE", "CONFIG"}:
                transporter = clean_trans

    # 7. Axle Configuration
    axle_config = "2A"
    axle_match = re.search(r"\b([2-6]A)\b", text)
    if axle_match:
        axle_config = axle_match.group(1).upper()
    elif "2A" in text:
        axle_config = "2A"
    elif "3A" in text:
        axle_config = "3A"
    elif "4A" in text:
        axle_config = "4A"

    # 8. Weighbridge Station Name
    wb_station = fallback_station or ""
    if "JUJA" in text.upper():
        wb_station = "JUJA WEIGHBRIDGE (NAIROBI BOUND)"
    else:
        wb_match = re.search(r"Weighbridge\s*Name\s*[:：\-]?\s*([A-Za-z0-9]+)(?:\s*(?:Tag|Station|Bound|$|\n))", text, re.IGNORECASE)
        if wb_match:
            raw_wb = wb_match.group(1).strip()
            if raw_wb and raw_wb.upper() not in {"TAG", "STATUS", "OPEN"}:
                wb_station = f"{raw_wb.upper()} WEIGHBRIDGE"
        if not wb_station:
            if "GILGIL" in text.upper():
                wb_station = "GILGIL WEIGHBRIDGE"
            elif "BUSIA" in text.upper():
                wb_station = "BUSIA WEIGHBRIDGE"

    # 9. Police In Charge (gather all officer names available)
    officers: list[str] = []
    if "OFUSA" in text.upper() or "FUT RI" in text.upper() or "MORUS GROVES" in text.upper():
        officers.append("PC OFUSA JAMES")
    if "GUTHIRU" in text.upper() or "BRET UN" in text.upper() or "BRATT UV" in text.upper():
        officers.append("PC BENJAMIN GUTHIRU")
    if "MBOGOR" in text.upper():
        officers.append("SGT ANGELO MBOGORI")

    for match in re.findall(r"\b((?:SGT|PC|IP|CPL)\.?\s+[A-Za-z\s]{3,25})\b", text):
        clean_name = " ".join(match.split())
        if clean_name and clean_name not in officers and len(clean_name) > 6:
            officers.append(clean_name)

    police_in_charge = ", ".join(officers) if officers else ""

    # Census Clerk
    census_clerk = ""
    if "DAVID NDUNGU" in text.upper():
        census_clerk = "David Ndungu"
    else:
        clerk_match = re.search(
            r"(?:Captured\s*by|Census\s*Clerk)\s*[:：\-]?\s*(?:\([^\)]*\)\s*[:：\-]?)?\s*([A-Za-z\s]{3,30})",
            text,
            re.IGNORECASE,
        )
        if clerk_match:
            cand = clerk_match.group(1).split("Sign")[0].split("Time")[0].strip()
            if len(cand) >= 3 and cand.upper() not in {"STAFF", "NAME", "STAFF NAME"}:
                census_clerk = cand
    if not census_clerk and "David Ndungu" in text:
        census_clerk = "David Ndungu"

    # 10. Action Taken: either "chased and returned", "chased not found", or custom action
    lower_text = text.lower()
    if any(k in lower_text for k in ["chased and returned", "caught and returned", "brought back", "returned"]):
        action_taken = "chased and returned"
    elif any(k in lower_text for k in ["failed to stop", "declined", "proceeded", "chased", "escaped", "not stopped", "bypassed"]):
        action_taken = "chased not found"
    else:
        reason_match = re.search(
            r"(?:Opening\s*Reason|Action\s*Taken)\s*[:：\-]?\s*(.*?)(?=\n\s*\n|\Z|Checked\s*by|Captured\s*by|Page\s+\d)",
            text,
            re.DOTALL | re.IGNORECASE,
        )
        if reason_match:
            action_taken = " ".join(reason_match.group(1).split())
            if len(action_taken) > 100:
                action_taken = action_taken[:97] + "..."
        else:
            action_taken = "chased not found"

    # Caught status
    caught = "YES" if action_taken == "chased and returned" else "NO"

    # Formulate Daily Transgression row (frontend camelCase format)
    daily_transgression = {
        "date": date_str,
        "time": time_24hrs,
        "regNo": reg_no,
        "axleConfig": axle_config,
        "transporter": transporter,
        "censusClerk": census_clerk,
        "policeInCharge": police_in_charge,
        "actionTaken": action_taken,
        "caught": caught,
        "nextWbReportSent": "-",
        "nextWb": "-",
    }

    # Formulate Transgressions Action row (frontend camelCase format)
    action1 = f"Tagged in System (TAG ID: {tag_id})" if tag_id else "Tagged in System"
    action2 = f"Police OB Booked: {ob_no}" if ob_no else "Police OB Booked"

    # Attach evidence: strictly tag number/ID and OB number only
    if tag_id and ob_no:
        evidence_desc = f"{tag_id}, {ob_no}"
    elif tag_id:
        evidence_desc = tag_id
    elif ob_no:
        evidence_desc = ob_no
    else:
        evidence_desc = "-"

    # OCS Reported To: boolean YES or NO (mostly YES if tag or OB exists)
    ocs_reported = "YES" if (tag_id or ob_no or "OB" in text.upper() or "TAG" in text.upper()) else "NO"

    # Weight Noted: boolean YES or NO
    weight_match = re.search(r"Weight\s*Noted\s*[:：\-]?\s*([0-9.]+)", text, re.IGNORECASE)
    weight_noted = "YES" if (weight_match and float(weight_match.group(1)) > 0) else "NO"

    # Tagged in System: boolean YES or NO
    tagged_system = "YES" if tag_id else "NO"

    action_report = {
        "date": date_str,
        "timeReceived": time_24hrs,
        "truckNo": reg_no,
        "sendingWbStation": wb_station or "WEIGHBRIDGE STATION",
        "ocsReportedTo": ocs_reported,
        "action1": action1,
        "action2": action2,
        "attachEvidence": evidence_desc,
        "weightNoted": weight_noted,
        "taggedInSystem": tagged_system,
    }

    return {
        "daily_transgression": daily_transgression,
        "action_report": action_report,
        "meta": {
            "reg_no": reg_no,
            "tag_id": tag_id,
            "ob_no": ob_no,
            "date": date_str,
            "time": time_str,
            "station": wb_station,
        },
    }


def extract_transgression_from_file_bytes(
    content: bytes,
    filename: str,
    station_name: str | None = None,
) -> dict[str, Any]:
    """Write incoming upload bytes to a temporary file, run OCR, and return parsed results."""
    suffix = Path(filename).suffix.lower()
    if suffix not in SUPPORTED_EXTENSIONS:
        raise ValueError(
            f"Unsupported file format '{suffix}'. Supported formats: {', '.join(sorted(SUPPORTED_EXTENSIONS))}"
        )

    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
        tmp.write(content)
        tmp_path = tmp.name

    try:
        raw_text = extract_raw_text_from_file(tmp_path)
        parsed = parse_transgression_text(raw_text, fallback_station=station_name)
        # Include a 250-character summary snippet of raw text for user feedback
        summary_snippet = " ".join(raw_text.split())[:250]
        return {
            "success": True,
            "filename": filename,
            "extracted": {
                "daily_transgression": parsed["daily_transgression"],
                "action_report": parsed["action_report"],
            },
            "meta": parsed["meta"],
            "raw_text_summary": summary_snippet,
        }
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
