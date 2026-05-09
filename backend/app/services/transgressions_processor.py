from typing import Any


DAILY_TRANSGRESSIONS_COLUMNS = [
    "Date",
    "Time",
    "Reg No",
    "Axle Config",
    "Transporter",
    "Census Clerk",
    "Police In charge",
    "Action Taken",
    "Caught",
    "Next WB report sent",
    "Next WB",
]

ACTION_REPORT_COLUMNS = [
    "Date",
    "Time Received",
    "Truck No.",
    "Sending WB station",
    "OCS Reported To",
    "Action 1",
    "Action 2",
    "Attach evidence if any",
    "Weight Noted",
    "Tagged in System",
]

DAILY_TRANSGRESSIONS_ALIASES = {
    "Date": ["date"],
    "Time": ["time"],
    "Reg No": ["reg_no", "registration", "vehicle_reg", "reg_number", "truck_no"],
    "Axle Config": ["axle_config", "axleconf", "axle_configuration"],
    "Transporter": ["transporter"],
    "Census Clerk": ["census_clerk", "clerk"],
    "Police In charge": ["police_in_charge", "police", "police_incharge"],
    "Action Taken": ["action_taken", "action"],
    "Caught": ["caught"],
    "Next WB report sent": [
        "next_wb_report_sent",
        "next_report_sent",
        "report_sent",
    ],
    "Next WB": ["next_wb", "next_weighbridge"],
}

ACTION_REPORT_ALIASES = {
    "Date": ["date"],
    "Time Received": ["time_received", "time"],
    "Truck No.": ["truck_no", "truck_number", "registration", "reg_no", "vehicle_reg"],
    "Sending WB station": ["sending_wb_station", "sending_station", "sending_wb"],
    "OCS Reported To": ["ocs_reported_to", "ocs_reported", "ocs"],
    "Action 1": ["action_1", "action1", "action"],
    "Action 2": ["action_2", "action2"],
    "Attach evidence if any": ["attach_evidence_if_any", "evidence", "attachment"],
    "Weight Noted": ["weight_noted", "weight"],
    "Tagged in System": ["tagged_in_system", "tagged"],
}


def _normalize_key(key: str) -> str:
    return key.strip().lower().replace(" ", "_").replace("-", "_").replace(".", "")


def _stringify_cell(value: Any) -> str:
    if value is None:
        return ""

    if isinstance(value, bool):
        return "YES" if value else "NO"

    return str(value).strip()


def _normalize_row(
    row: dict[str, Any],
    columns: list[str],
    aliases: dict[str, list[str]],
) -> dict[str, str]:
    if not isinstance(row, dict):
        raise ValueError("Transgressions rows must be objects.")

    normalized_source = {_normalize_key(str(key)): value for key, value in row.items()}
    output = {}

    for column in columns:
        candidates = [column, *aliases[column]]
        value = ""

        for candidate in candidates:
            normalized_candidate = _normalize_key(candidate)
            if normalized_candidate in normalized_source:
                value = normalized_source[normalized_candidate]
                break

        output[column] = _stringify_cell(value)

    return output


def _normalize_rows(
    rows: Any,
    field_name: str,
    columns: list[str],
    aliases: dict[str, list[str]],
) -> list[dict[str, str]]:
    if rows is None:
        return []

    if not isinstance(rows, list):
        raise ValueError(f"Transgressions field '{field_name}' must be a list.")

    return [_normalize_row(row, columns, aliases) for row in rows]


def normalize_transgressions_input(payload: Any) -> dict[str, list[dict[str, str]]]:
    if payload is None:
        payload = {}

    if isinstance(payload, list):
        payload = {"daily_transgressions": payload, "action_report": []}

    if not isinstance(payload, dict):
        raise ValueError("Transgressions input must be an object.")

    daily_rows = payload.get("daily_transgressions", payload.get("daily", []))
    action_rows = payload.get("action_report", payload.get("actions", []))

    return {
        "daily_transgressions": _normalize_rows(
            daily_rows,
            "daily_transgressions",
            DAILY_TRANSGRESSIONS_COLUMNS,
            DAILY_TRANSGRESSIONS_ALIASES,
        ),
        "action_report": _normalize_rows(
            action_rows,
            "action_report",
            ACTION_REPORT_COLUMNS,
            ACTION_REPORT_ALIASES,
        ),
    }
