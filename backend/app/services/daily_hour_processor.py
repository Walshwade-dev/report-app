import pandas as pd
import random

REQUIRED_COLUMNS = {
    "MultiDeck[D]": "D",
    "SingleAxle[S]": "S",
    "Manually[M]": "M",
    "HSWIM Total[H]": "H",
    "CalledIn[C]": "C",
    "WarnedTucks[A]": "A",
    "ChargedInCourt[Z]": "Z",
    "SpecialRelease[G]": "G",
    "Redistributed[R]": "R",
}

HOURS = [
    "0000-0100", "0100-0200", "0200-0300", "0300-0400",
    "0400-0500", "0500-0600", "0600-0700", "0700-0800",
    "0800-0900", "0900-1000", "1000-1100", "1100-1200",
    "1200-1300", "1300-1400", "1400-1500", "1500-1600",
    "1600-1700", "1700-1800", "1800-1900", "1900-2000",
    "2000-2100", "2100-2200", "2200-2300", "2300-0000",
]

WIDELOAD_DISTRIBUTION_START_HOUR = "0600-0700"
WIDELOAD_DISTRIBUTION_END_HOUR = "1700-1800"
WIDELOAD_DISTRIBUTION_INDEXES = list(
    range(
        HOURS.index(WIDELOAD_DISTRIBUTION_START_HOUR),
        HOURS.index(WIDELOAD_DISTRIBUTION_END_HOUR) + 1,
    )
)
WIDELOAD_PEAK_INDEXES = {
    HOURS.index("0600-0700"),
    HOURS.index("0700-0800"),
    HOURS.index("0800-0900"),
    HOURS.index("0900-1000"),
    HOURS.index("1600-1700"),
    HOURS.index("1700-1800"),
}
WIDELOAD_PEAK_MAX_PER_HOUR = 12
WIDELOAD_NORMAL_MAX_PER_HOUR = 7


def extract_daily_hour_core(df: pd.DataFrame) -> pd.DataFrame:
    """
    Extract only required columns and remove totals row
    """

    # 1. Ensure required columns exist
    missing = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing:
        raise ValueError(f"Missing columns: {missing}")

    # 2. Extract only required columns
    df = df[list(REQUIRED_COLUMNS.keys())].copy()

    # 3. Rename to clean short names
    df.rename(columns=REQUIRED_COLUMNS, inplace=True)

    # 4. Remove totals row
    # Usually last row OR contains NaN or text
    df = df.iloc[:-1]

    # 5. Reset index
    df.reset_index(drop=True, inplace=True)

    # 6. Ensure numeric values
    for col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)

    return df


def choose_weighted_index(rng, indexes: list[int], weights: dict[int, float]) -> int:
    total_weight = sum(weights[index] for index in indexes)
    threshold = rng.uniform(0, total_weight)
    running_weight = 0

    for index in indexes:
        running_weight += weights[index]

        if running_weight >= threshold:
            return index

    return indexes[-1]


def wideload_hour_cap(hour_index: int) -> int:
    if hour_index in WIDELOAD_PEAK_INDEXES:
        return WIDELOAD_PEAK_MAX_PER_HOUR

    return WIDELOAD_NORMAL_MAX_PER_HOUR


def wideload_hour_weight(rng, hour_index: int, current_value: int) -> float:
    cap = wideload_hour_cap(hour_index)
    is_peak_hour = hour_index in WIDELOAD_PEAK_INDEXES
    peak_bias = 6.0 if is_peak_hour else 0.9

    if current_value <= 2:
        band_bias = 2.1
    elif current_value <= 6:
        band_bias = 1.35
    elif is_peak_hour and current_value <= 9:
        band_bias = 0.85
    else:
        band_bias = 0.25

    remaining_capacity = max(cap - current_value, 1)
    capacity_bias = (remaining_capacity / cap) ** 1.25
    jitter = rng.uniform(0.75, 1.35)

    return peak_bias * band_bias * capacity_bias * jitter


def distribute_wideloads(wideload_count: int, rng=None) -> list[int]:
    values = [0] * 24

    if wideload_count <= 0:
        return values

    rng = rng or random.SystemRandom()
    remaining_count = wideload_count

    if wideload_count >= len(WIDELOAD_DISTRIBUTION_INDEXES):
        for hour_index in WIDELOAD_DISTRIBUTION_INDEXES:
            values[hour_index] = 1
            remaining_count -= 1

    for _ in range(remaining_count):
        available_indexes = [
            hour_index
            for hour_index in WIDELOAD_DISTRIBUTION_INDEXES
            if values[hour_index] < wideload_hour_cap(hour_index)
        ]

        if not available_indexes:
            available_indexes = WIDELOAD_DISTRIBUTION_INDEXES

        weights = {
            hour_index: wideload_hour_weight(rng, hour_index, values[hour_index])
            for hour_index in available_indexes
        }
        chosen_index = choose_weighted_index(rng, available_indexes, weights)
        values[chosen_index] += 1

    return values

def build_daily_hour_metrics(
    df: pd.DataFrame,
    report_date: str,
    wideload_count: int,
) -> pd.DataFrame:
    core_df = extract_daily_hour_core(df)

    result = pd.DataFrame()

    result["DATE"] = [report_date] + [""] * 23
    result["TIME"] = HOURS

    result["D"] = core_df["D"]
    result["S"] = core_df["S"]
    result["M"] = core_df["M"]
    result["H"] = core_df["H"]

    result["Q"] = core_df["H"] - core_df["C"]
    result["X"] = core_df["D"] + core_df["S"] + core_df["M"]

    result["C"] = core_df["C"]

    result["Y"] = (
        core_df["A"]
        + core_df["Z"]
        + core_df["G"]
        + core_df["R"]
    )

    result["P"] = core_df["Z"] + core_df["R"]

    result["A"] = core_df["A"]
    result["Z"] = core_df["Z"]
    result["G"] = core_df["G"]
    result["R"] = core_df["R"]

    result["E"] = distribute_wideloads(wideload_count)

    return result

def add_daily_totals_row(df: pd.DataFrame) -> pd.DataFrame:
    totals = {}

    for col in df.columns:
        if col in ["DATE", "TIME"]:
            totals[col] = "Totals" if col == "DATE" else ""
        else:
            totals[col] = int(df[col].sum())

    return pd.concat([df, pd.DataFrame([totals])], ignore_index=True)
