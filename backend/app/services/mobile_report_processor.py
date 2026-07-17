from typing import Any

import pandas as pd

from app.services.daily_hour_processor import HOURS
from app.services.report_upload_service import drop_repeated_header_rows


WEIGHBRIDGE_REGISTER_COLUMNS = [
    "NO",
    "id",
    "Date Time",
    "Ticket No.",
    "Station",
    "Registration",
    "Axle",
    "Transporter",
    "Cargo",
    "Make",
    "Origin",
    "Destination",
    "Deck A[KG]",
    "Deck B[KG]",
    "Deck C[KG]",
    "Deck D[KG]",
    "GVW [KG]",
    "Excess",
    "Excess [KG]",
    "Status",
    "State",
    "Mismatch",
]

OPTIONAL_WEIGHBRIDGE_REGISTER_COLUMNS = [
    "Remarks",
]

MOBILE_REPORT_COLUMN_MAPPING = {
    "NO": "no",
    "id": "id",
    "Date Time": "date_time",
    "Ticket No.": "ticket_no",
    "Station": "station",
    "Registration": "registration",
    "Axle": "axle",
    "Transporter": "transporter",
    "Cargo": "cargo",
    "Make": "make",
    "Origin": "origin",
    "Destination": "destination",
    "Deck A[KG]": "deck_a_kg",
    "Deck B[KG]": "deck_b_kg",
    "Deck C[KG]": "deck_c_kg",
    "Deck D[KG]": "deck_d_kg",
    "GVW [KG]": "gvw_kg",
    "Remarks": "source_remarks",
    "Excess": "excess_type",
    "Excess [KG]": "excess_kg",
    "Status": "status",
    "State": "state",
    "Mismatch": "mismatch",
}

TEXT_COLUMNS = [
    "ticket_no",
    "station",
    "registration",
    "axle",
    "transporter",
    "cargo",
    "make",
    "origin",
    "destination",
    "source_remarks",
    "excess_type",
    "status",
    "state",
    "mismatch",
]

NUMERIC_COLUMNS = [
    "no",
    "id",
    "deck_a_kg",
    "deck_b_kg",
    "deck_c_kg",
    "deck_d_kg",
    "gvw_kg",
    "excess_kg",
]

OUTPUT_COLUMNS = [
    "no",
    "id",
    "date_time",
    "ticket_no",
    "station",
    "registration",
    "axle",
    "transporter",
    "cargo",
    "make",
    "origin",
    "destination",
    "deck_a_kg",
    "deck_b_kg",
    "deck_c_kg",
    "deck_d_kg",
    "gvw_kg",
    "total_gvw_kg",
    "gvw_difference_kg",
    "remarks",
    "source_remarks",
    "is_weighed",
    "is_dimension_charge",
    "is_gvw_axle_charge",
    "excess_type",
    "excess_kg",
    "status",
    "state",
    "mismatch",
    "hour_band",
]


def _remark(row: pd.Series) -> str:
    source_remarks = str(row.get("source_remarks", "")).strip().upper()

    if source_remarks:
        return source_remarks

    diff = row.get("gvw_difference_kg", 0)
    if diff > 2000:
        return "CHARGED"

    excess_kg = row.get("excess_kg", 0)
    if excess_kg <= 0:
        return "LEGAL"

    return "WARNED"


def _is_dimension_charge(remarks: str) -> bool:
    normalized = str(remarks).strip().lower()
    return "dimension" in normalized


def _is_gvw_axle_charge(row: pd.Series) -> bool:
    remarks = str(row["remarks"]).strip().lower()

    if _is_dimension_charge(remarks):
        return False

    return "charged" in remarks


def _hour_band(hour: int) -> str:
    return HOURS[hour]


def _single_value(values: pd.Series) -> Any:
    unique_values = [
        value
        for value in values.dropna().astype(str).str.strip().unique().tolist()
        if value
    ]

    if len(unique_values) == 1:
        return unique_values[0]

    return None


def normalize_mobile_report(
    df: pd.DataFrame,
    reweigh_tickets: list[str] = None,
    dimension_charges: list[dict[str, Any]] = None,
    station: str = None,
) -> pd.DataFrame:
    df = drop_repeated_header_rows(df)

    missing = [
        column
        for column in WEIGHBRIDGE_REGISTER_COLUMNS
        if column not in df.columns
    ]

    if missing:
        raise ValueError(f"Missing mobile report columns: {missing}")

    available_optional_columns = [
        column for column in OPTIONAL_WEIGHBRIDGE_REGISTER_COLUMNS if column in df.columns
    ]
    records = df[WEIGHBRIDGE_REGISTER_COLUMNS + available_optional_columns].rename(
        columns=MOBILE_REPORT_COLUMN_MAPPING
    )
    records = records.copy()

    if "source_remarks" not in records.columns:
        records["source_remarks"] = ""

    records["date_time"] = pd.to_datetime(
        records["date_time"],
        errors="coerce",
        dayfirst=True,
    )
    records = records.loc[records["date_time"].notna()].reset_index(drop=True)

    for column in TEXT_COLUMNS:
        records[column] = records[column].fillna("").astype(str).str.strip()

    for column in NUMERIC_COLUMNS:
        records[column] = (
            pd.to_numeric(records[column], errors="coerce")
            .fillna(0)
            .astype(int)
        )

    records = records.sort_values("date_time").reset_index(drop=True)
    records["total_gvw_kg"] = records[
        ["deck_a_kg", "deck_b_kg", "deck_c_kg", "deck_d_kg"]
    ].sum(axis=1)
    records["gvw_difference_kg"] = records["total_gvw_kg"] - records["gvw_kg"]

    # Wrap _remark to handle reweigh tickets
    def get_remark(row):
        ticket = str(row.get("ticket_no", "")).strip()
        if reweigh_tickets and ticket in reweigh_tickets:
            return "REWEIGH"
        return _remark(row)

    records["remarks"] = records.apply(get_remark, axis=1)
    records["is_weighed"] = records["total_gvw_kg"] > 0
    records["is_dimension_charge"] = records["remarks"].map(_is_dimension_charge)
    records["is_gvw_axle_charge"] = records.apply(_is_gvw_axle_charge, axis=1)
    records["hour_band"] = records["date_time"].dt.hour.map(_hour_band)

    records = records[OUTPUT_COLUMNS]

    # Append manually entered dimension charges if any
    if dimension_charges:
        manual_rows = []
        start_no = int(records["no"].max()) if not records.empty else 0
        for idx, charge in enumerate(dimension_charges):
            dt_str = charge.get("date_time") or ""
            try:
                dt = pd.to_datetime(dt_str, errors="coerce")
                if pd.isna(dt):
                    dt = pd.Timestamp.now()
            except Exception:
                dt = pd.Timestamp.now()

            gvw_diff = int(charge.get("gvw_excess") or charge.get("gvw_difference_kg") or 0)
            axle_excess = int(charge.get("axle_excess") or charge.get("excess_kg") or 0)

            row_dict = {
                "no": start_no + 1 + idx,
                "id": charge.get("id") or str(idx + 1),
                "date_time": dt,
                "ticket_no": charge.get("ticket_no") or f"MANUAL-DIM-{idx+1}",
                "station": station or charge.get("station") or "",
                "registration": str(charge.get("registration") or "").strip().upper(),
                "axle": str(charge.get("axle") or "").strip().upper(),
                "transporter": str(charge.get("transporter") or "").strip().upper(),
                "cargo": str(charge.get("cargo") or "").strip().upper(),
                "make": str(charge.get("make") or "").strip().upper(),
                "origin": str(charge.get("origin") or "").strip().upper(),
                "destination": str(charge.get("destination") or "").strip().upper(),
                "deck_a_kg": 0,
                "deck_b_kg": 0,
                "deck_c_kg": 0,
                "deck_d_kg": 0,
                "gvw_kg": 0,
                "total_gvw_kg": 0,
                "gvw_difference_kg": gvw_diff,
                "remarks": "CHARGED (DIMENSIONS)",
                "source_remarks": "CHARGED (DIMENSIONS)",
                "is_weighed": True,
                "is_dimension_charge": True,
                "is_gvw_axle_charge": False,
                "excess_type": "DIMENSION",
                "excess_kg": axle_excess,
                "status": "CHARGED",
                "state": "",
                "mismatch": "",
                "hour_band": _hour_band(dt.hour),
            }
            manual_rows.append(row_dict)

        if manual_rows:
            manual_df = pd.DataFrame(manual_rows)
            records = pd.concat([records, manual_df], ignore_index=True)

    records = records.sort_values("date_time").reset_index(drop=True)
    records["no"] = range(1, len(records) + 1)

    return records[OUTPUT_COLUMNS]


def find_duplicate_registrations(df: pd.DataFrame) -> list[dict[str, Any]]:
    df_clean = df.copy()
    reg_col = None
    for col in ["Registration", "registration"]:
        if col in df_clean.columns:
            reg_col = col
            break
    if not reg_col:
        return []

    df_clean[reg_col] = df_clean[reg_col].fillna("").astype(str).str.strip().str.upper()
    df_clean = df_clean[df_clean[reg_col] != ""]

    counts = df_clean[reg_col].value_counts()
    dup_regs = counts[counts > 1].index.tolist()

    duplicates = []
    for reg in dup_regs:
        group = df_clean[df_clean[reg_col] == reg].copy()
        dt_col = None
        for col in ["Date Time", "date_time"]:
            if col in group.columns:
                dt_col = col
                break
        if dt_col:
            group[dt_col] = pd.to_datetime(group[dt_col], errors="coerce", dayfirst=True)
            group = group.sort_values(dt_col)

        tickets = []
        for _, row in group.iterrows():
            ticket_no = str(row.get("Ticket No.", row.get("ticket_no", ""))).strip()
            date_time = str(row.get("Date Time", row.get("date_time", ""))).strip()
            if dt_col and isinstance(row.get(dt_col), pd.Timestamp):
                date_time = row[dt_col].strftime("%Y-%m-%d %H:%M:%S")
            gvw = str(row.get("GVW [KG]", row.get("gvw_kg", "0"))).strip()
            source_remarks = str(row.get("Remarks", row.get("source_remarks", ""))).strip().upper()
            
            # calculate diff
            deck_sum = sum(
                pd.to_numeric(row.get(c, 0), errors="coerce")
                for c in ["Deck A[KG]", "Deck B[KG]", "Deck C[KG]", "Deck D[KG]",
                            "deck_a_kg", "deck_b_kg", "deck_c_kg", "deck_d_kg"]
                if c in row
            )
            gvw_val = pd.to_numeric(row.get("GVW [KG]", row.get("gvw_kg", 0)), errors="coerce")
            diff_val = deck_sum - gvw_val
            is_charged = "CHARGED" in source_remarks or diff_val > 2000

            tickets.append({
                "ticket_no": ticket_no,
                "date_time": date_time,
                "gvw_kg": gvw,
                "remarks": source_remarks or ("CHARGED" if is_charged else "LEGAL")
            })
        duplicates.append({
            "registration": reg,
            "tickets": tickets
        })
    return duplicates


def summarize_mobile_report(records: pd.DataFrame) -> dict[str, Any]:
    hourly_counts = records["hour_band"].value_counts().to_dict()
    weighed_mask = records["is_weighed"] if "is_weighed" in records else (
        records[["deck_a_kg", "deck_b_kg", "deck_c_kg", "deck_d_kg"]].sum(axis=1) > 0
    )
    warned_mask = records["remarks"].str.strip().str.upper().eq("WARNED")
    charged_gvw_axle_mask = (
        records["is_gvw_axle_charge"]
        if "is_gvw_axle_charge" in records
        else records["remarks"].str.strip().str.upper().eq("CHARGED")
    )
    charged_dimensions_mask = (
        records["is_dimension_charge"]
        if "is_dimension_charge" in records
        else records["remarks"].str.contains("dimension", case=False, na=False)
    )
    charged_mask = charged_gvw_axle_mask | charged_dimensions_mask
    mismatch_mask = records["mismatch"].str.strip().str.lower().isin(
        {"mismatch", "yes", "true", "1"}
    )

    return {
        "total_records": int(len(records)),
        "total_trucks_weighed": int(weighed_mask.sum()),
        "report_date": _single_value(records["date_time"].dt.date.astype(str)),
        "station": _single_value(records["station"]),
        "warned_trucks": int(warned_mask.sum()),
        "charged_trucks": int(charged_mask.sum()),
        "charged_gvw_axle_trucks": int(charged_gvw_axle_mask.sum()),
        "charged_dimensions_trucks": int(charged_dimensions_mask.sum()),
        "overloaded_records": int((warned_mask | charged_mask).sum()),
        "total_excess_kg": int(records["excess_kg"].sum()),
        "mismatch_records": int(mismatch_mask.sum()),
        "hourly_counts": {
            hour: int(hourly_counts.get(hour, 0))
            for hour in HOURS
        },
    }


def mobile_report_response(
    df: pd.DataFrame,
    reweigh_tickets: list[str] = None,
    dimension_charges: list[dict[str, Any]] = None,
    station: str = None,
) -> dict[str, Any]:
    records = normalize_mobile_report(
        df,
        reweigh_tickets=reweigh_tickets,
        dimension_charges=dimension_charges,
        station=station,
    )
    serializable = records.copy()
    serializable["date_time"] = serializable["date_time"].dt.strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    return {
        "columns": serializable.columns.tolist(),
        "summary": summarize_mobile_report(records),
        "data": serializable.where(pd.notnull(serializable), None).to_dict(
            orient="records"
        ),
        "duplicates": find_duplicate_registrations(df),
    }
