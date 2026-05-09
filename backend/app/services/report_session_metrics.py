from typing import Any

import pandas as pd

from app.services.traffic_census_processor import normalize_traffic_census_input


TRAFFIC_CENSUS_WIDELOAD_REQUIRED_MESSAGE = (
    "Traffic Census preview requires wideload upload first because E comes from wideload count."
)


def _safe_int(value: Any, default: int = 0) -> int:
    if value is None or value == "":
        return default

    if isinstance(value, str):
        value = value.replace(",", "").strip()

    try:
        return int(float(value))
    except (TypeError, ValueError):
        return default


def get_wideload_count_from_session(session) -> int | None:
    wideload_count = session.sections.get("wideload", {}).get("wideload_count")

    if wideload_count is None:
        return None

    return _safe_int(wideload_count)


def get_daily_hour_totals_from_dataframe(daily_df: pd.DataFrame | None) -> dict[str, int]:
    totals = {"q": 0, "x": 0, "e": 0}

    if daily_df is None or daily_df.empty or "DATE" not in daily_df.columns:
        return totals

    totals_mask = (
        daily_df["DATE"]
        .astype(str)
        .str.strip()
        .str.lower()
        .eq("totals")
    )

    if not totals_mask.any():
        return totals

    row = daily_df.loc[totals_mask].iloc[-1]
    totals["q"] = _safe_int(row.get("Q", 0))
    totals["x"] = _safe_int(row.get("X", 0))
    totals["e"] = _safe_int(row.get("E", 0))
    return totals


def get_daily_hour_totals(session) -> dict[str, int]:
    if session.sections.get("daily_hour", {}).get("status") != "ready":
        return {"q": 0, "x": 0, "e": 0}

    return get_daily_hour_totals_from_dataframe(session.dataframes.get("daily_hour"))


def get_traffic_census_section_values(session) -> dict[str, int]:
    wideload_count = get_wideload_count_from_session(session)

    if wideload_count is None:
        raise ValueError(TRAFFIC_CENSUS_WIDELOAD_REQUIRED_MESSAGE)

    traffic_census = normalize_traffic_census_input(
        session.manual_inputs.get("traffic_census") or {}
    )
    daily_totals = get_daily_hour_totals(session)
    total_census = traffic_census["total_traffic_census"]

    return {
        "e": wideload_count,
        "x": daily_totals["x"],
        "q": daily_totals["q"],
        "k": total_census,
        "total_traffic": daily_totals["q"]
        + daily_totals["x"]
        + total_census
        + wideload_count,
    }
