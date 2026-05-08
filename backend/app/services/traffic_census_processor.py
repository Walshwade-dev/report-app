import math
from typing import Any


TRAFFIC_CENSUS_FIELDS = (
    "buses_gte_3500kg",
    "vehicles_3500_to_7000_excluding_buses",
    "vehicles_gte_7000_excluding_buses",
)

TOTAL_TRAFFIC_CENSUS_FIELD = "total_traffic_census"

TRAFFIC_CENSUS_LABELS = {
    "buses_gte_3500kg": "Buses >= 3,500 kg",
    "vehicles_3500_to_7000_excluding_buses": "Vehicles 3,500 to 7,000 kg excluding buses",
    "vehicles_gte_7000_excluding_buses": "Vehicles >= 7,000 kg excluding buses",
    "total_traffic_census": "Total Traffic Census",
}


def _normalize_count(value: Any, field_name: str) -> int:
    if value is None or value == "":
        raise ValueError(f"Traffic census field '{field_name}' is required.")

    if isinstance(value, bool):
        raise ValueError(f"Traffic census field '{field_name}' must be a number.")

    if isinstance(value, str):
        clean_value = value.replace(",", "").strip()
        try:
            value = float(clean_value)
        except ValueError as exc:
            raise ValueError(
                f"Traffic census field '{field_name}' must be a number."
            ) from exc

    if not isinstance(value, (int, float)):
        raise ValueError(f"Traffic census field '{field_name}' must be a number.")

    if not math.isfinite(float(value)):
        raise ValueError(f"Traffic census field '{field_name}' must be finite.")

    if float(value) < 0:
        raise ValueError(f"Traffic census field '{field_name}' cannot be negative.")

    if not float(value).is_integer():
        raise ValueError(f"Traffic census field '{field_name}' must be a whole number.")

    return int(value)


def normalize_traffic_census_input(payload: dict[str, Any]) -> dict[str, int]:
    if not isinstance(payload, dict):
        raise ValueError("Traffic census input must be an object.")

    normalized = {
        field_name: _normalize_count(payload.get(field_name), field_name)
        for field_name in TRAFFIC_CENSUS_FIELDS
    }
    computed_total = sum(normalized.values())

    provided_total = payload.get(TOTAL_TRAFFIC_CENSUS_FIELD)

    if provided_total is None or provided_total == "":
        normalized[TOTAL_TRAFFIC_CENSUS_FIELD] = computed_total
        return normalized

    normalized_total = _normalize_count(provided_total, TOTAL_TRAFFIC_CENSUS_FIELD)

    if normalized_total != computed_total:
        raise ValueError(
            "Traffic census total_traffic_census does not match category total: "
            f"provided {normalized_total}, computed {computed_total}."
        )

    normalized[TOTAL_TRAFFIC_CENSUS_FIELD] = normalized_total
    return normalized


def traffic_census_rows(payload: dict[str, Any]) -> list[tuple[str, int]]:
    normalized = normalize_traffic_census_input(payload)
    return [
        (TRAFFIC_CENSUS_LABELS[field_name], normalized[field_name])
        for field_name in [*TRAFFIC_CENSUS_FIELDS, TOTAL_TRAFFIC_CENSUS_FIELD]
    ]
