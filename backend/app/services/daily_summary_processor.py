from typing import Any

import pandas as pd

from app.services.overloaded_summary import count_valid_permit_vehicles
from app.services.traffic_census_processor import normalize_traffic_census_input


class DailySummaryMissingSourceError(ValueError):
    pass


SUMMARY_FIELDS = [
    ("weighed_by_hswim_q", "Weighed by HSWIM (Q)"),
    ("weighed_scale_total_n", "Weighed Scale total (N)=D+S"),
    ("manually_weighed_m", "Manually Weighed (M)"),
    ("total_weighed_x", "Total weighed (X)"),
    ("total_traffic_t", "Total Traffic (T)"),
    ("total_overload_y", "Total Overload (Y)=A+Z+G+R"),
    ("warned_a", "Warned (A)"),
    ("charged_prohibited_z", "Charged & Prohibited (Z)"),
    ("special_release_g", "Special release (G)"),
    ("vehicles_charged_but_redistributed_r", "Vehicles Charged but Redistributed (R)"),
    ("impounded_prohibited_p", "Impounded & prohibited (P)=Z+R+G"),
    ("cases_cleared_in_court_b", "Cases cleared in Court (B)"),
    ("transgressions_l", "Transgressions (L)"),
    ("exemption_permits_not_weighed_e", "Exemption permits Not weighed (E)"),
    ("exemption_permits_weighed_f", "Exemption permits Weighed (F)"),
    ("exemption_permits_total", "Exemption permits Total"),
]


def _normalize_non_negative_count(value: Any, field_name: str, default: int = 0) -> int:
    if value is None or value == "":
        return default

    if isinstance(value, bool):
        raise ValueError(f"{field_name} must be a number.")

    if isinstance(value, str):
        value = value.replace(",", "").strip()

    try:
        number = float(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{field_name} must be a number.") from exc

    if number < 0:
        raise ValueError(f"{field_name} cannot be negative.")

    if not number.is_integer():
        raise ValueError(f"{field_name} must be a whole number.")

    return int(number)


def _daily_totals_row(daily_df: pd.DataFrame) -> pd.Series:
    if daily_df is None or daily_df.empty:
        raise DailySummaryMissingSourceError("Daily Summary requires ready daily_hour data.")

    if "DATE" in daily_df.columns:
        totals_mask = daily_df["DATE"].astype(str).str.strip().str.lower() == "totals"
        if totals_mask.any():
            return daily_df.loc[totals_mask].iloc[-1]

    numeric_totals = {}
    for column in daily_df.columns:
        if column in {"DATE", "TIME"}:
            continue

        numeric_totals[column] = int(pd.to_numeric(daily_df[column], errors="coerce").fillna(0).sum())

    return pd.Series(numeric_totals)


def _total_value(totals_row: pd.Series, column: str) -> int:
    if column not in totals_row:
        raise DailySummaryMissingSourceError(
            f"Daily Summary requires daily_hour total column '{column}'."
        )

    return _normalize_non_negative_count(totals_row[column], column)


def _manual_count(manual_inputs: dict[str, Any] | None, *keys: str) -> int:
    manual_inputs = manual_inputs or {}

    for key in keys:
        if key in manual_inputs:
            return _normalize_non_negative_count(manual_inputs[key], key)

    return 0


def _transgressions_count(manual_inputs: dict[str, Any] | None) -> int:
    manual_inputs = manual_inputs or {}

    for key in ("transgressions_count", "total_transgressions"):
        if key in manual_inputs:
            return _normalize_non_negative_count(manual_inputs[key], key)

    transgressions = manual_inputs.get("transgressions")

    if transgressions is None:
        return 0

    if isinstance(transgressions, list):
        return len(transgressions)

    if isinstance(transgressions, dict):
        for key in ("count", "total", "total_transgressions"):
            if key in transgressions:
                return _normalize_non_negative_count(transgressions[key], key)

        daily_transgressions = transgressions.get("daily_transgressions")
        if isinstance(daily_transgressions, list):
            return len(daily_transgressions)

    return 0


def build_daily_summary(
    daily_df: pd.DataFrame,
    traffic_census: dict[str, Any],
    overloaded_valid_permit_count: int,
    manual_inputs: dict[str, Any] | None = None,
) -> dict[str, int]:
    if not traffic_census:
        raise DailySummaryMissingSourceError("Daily Summary requires ready traffic_census data.")

    traffic = normalize_traffic_census_input(traffic_census)
    totals = _daily_totals_row(daily_df)

    d = _total_value(totals, "D")
    s = _total_value(totals, "S")
    m = _total_value(totals, "M")
    q = _total_value(totals, "Q")
    a = _total_value(totals, "A")
    z = _total_value(totals, "Z")
    g = _total_value(totals, "G")
    r = _total_value(totals, "R")
    e = _total_value(totals, "E")

    n = d + s
    x = d + s + m
    k = traffic["total_traffic_census"]
    t = q + x + k + e
    y = a + z + g + r
    p = z + r + g
    b = _manual_count(manual_inputs, "cases_cleared_in_court", "cases_cleared_court")
    l = _transgressions_count(manual_inputs)
    f = _normalize_non_negative_count(
        overloaded_valid_permit_count,
        "overloaded_valid_permit_count",
    )

    return {
        "weighed_by_hswim_q": q,
        "weighed_scale_total_n": n,
        "manually_weighed_m": m,
        "total_weighed_x": x,
        "total_traffic_t": t,
        "total_overload_y": y,
        "warned_a": a,
        "charged_prohibited_z": z,
        "special_release_g": g,
        "vehicles_charged_but_redistributed_r": r,
        "impounded_prohibited_p": p,
        "cases_cleared_in_court_b": b,
        "transgressions_l": l,
        "exemption_permits_not_weighed_e": e,
        "exemption_permits_weighed_f": f,
        "exemption_permits_total": e + f,
    }


def build_daily_summary_from_session(session) -> dict[str, int]:
    missing_sources = []

    if (
        "daily_hour" not in session.dataframes
        or session.sections.get("daily_hour", {}).get("status") != "ready"
    ):
        missing_sources.append("daily_hour")

    if (
        "traffic_census" not in session.manual_inputs
        or session.sections.get("traffic_census", {}).get("status") != "ready"
    ):
        missing_sources.append("traffic_census")

    if (
        "overloaded" not in session.dataframes
        or session.sections.get("overloaded", {}).get("status") != "ready"
    ):
        missing_sources.append("overloaded")

    if missing_sources:
        raise DailySummaryMissingSourceError(
            f"Daily Summary requires ready source data: {missing_sources}."
        )

    overloaded_valid_permit_count = session.sections.get("overloaded", {}).get(
        "valid_permit_count"
    )

    if overloaded_valid_permit_count is None:
        overloaded_valid_permit_count = count_valid_permit_vehicles(
            session.dataframes["overloaded"]
        )

    return build_daily_summary(
        daily_df=session.dataframes["daily_hour"],
        traffic_census=session.manual_inputs["traffic_census"],
        overloaded_valid_permit_count=overloaded_valid_permit_count,
        manual_inputs=session.manual_inputs,
    )


def daily_summary_rows(summary: dict[str, int]) -> list[tuple[str, int]]:
    return [(label, int(summary[key])) for key, label in SUMMARY_FIELDS]
