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


def distribute_wideloads(wideload_count: int) -> list[int]:
    """
    Distribute wideload count between 0600-1800 only.
    Heavier allocation earlier, lighter toward evening.
    """

    values = [0] * 24

    allowed_indexes = list(range(6, 18))  # 0600-0700 to 1700-1800

    weights = [
        12, 12, 11, 10, 9, 8,
        7, 6, 5, 4, 3, 2
    ]

    total_weight = sum(weights)

    allocation = [
        int((wideload_count * weight) / total_weight)
        for weight in weights
    ]

    remainder = wideload_count - sum(allocation)

    for _ in range(remainder):
        idx = random.choices(range(len(allowed_indexes)), weights=weights, k=1)[0]
        allocation[idx] += 1

    for i, hour_index in enumerate(allowed_indexes):
        values[hour_index] = allocation[i]

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