import pandas as pd


def count_valid_permit_vehicles(df: pd.DataFrame) -> int:
    if "Vardict" not in df.columns:
        raise ValueError("Missing required column: Vardict")

    pattern = r"(?i)\bVehicle\s+has\s+a\s+valid\s+permit\s+App-\d+\b"

    return int(
        df["Vardict"]
        .fillna("")
        .str.contains(pattern, regex=True)
        .sum()
    )