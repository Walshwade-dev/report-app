import pandas as pd


VALID_PERMIT_PATTERN = r"(?i)\bVehicle\s+has\s+a\s+valid\s+permit\s+App-\d+\b"
SPECIAL_RELEASE_PERMIT_PATTERN = (
    r"(?i)\bSpecial\s+Release\s*:\s*Vehicle\s+With\s+Permit\b"
    r".*\b(?:Truck\s+has\s+a\s+valid\s+)?Permit\s+No\s*:\s*App-\d+\b"
)


def count_valid_permit_vehicles(df: pd.DataFrame) -> int:
    if "Vardict" not in df.columns:
        raise ValueError("Missing required column: Vardict")

    vardicts = df["Vardict"].fillna("")
    valid_permit = vardicts.str.contains(VALID_PERMIT_PATTERN, regex=True)
    special_release_permit = vardicts.str.contains(
        SPECIAL_RELEASE_PERMIT_PATTERN,
        regex=True,
    )

    return int((valid_permit | special_release_permit).sum())
