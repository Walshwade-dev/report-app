import pandas as pd

def format_date(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series, errors="coerce").dt.strftime("%d/%m/%Y")

def format_comma(series: pd.Series) -> pd.Series:
    return (
        pd.to_numeric(series, errors="coerce")
        .apply(lambda x: f"{int(x):,}" if pd.notnull(x) else "")
    )

def clean_with_template(df: pd.DataFrame, template) -> pd.DataFrame:
    # Validate columns
    missing = [col for col in template.COLUMN_MAPPING.keys() if col not in df.columns]
    if missing:
        raise ValueError(f"Missing columns: {missing}")

    # Select + rename
    df = df[list(template.COLUMN_MAPPING.keys())].copy()
    df = df.rename(columns=template.COLUMN_MAPPING)

    # Apply formatting
    for col in template.DATE_COLUMNS:
        df[col] = format_date(df[col])

    for col in template.COMMA_COLUMNS:
        df[col] = format_comma(df[col])

    # Final order
    df = df[template.COLUMNS]

    return df