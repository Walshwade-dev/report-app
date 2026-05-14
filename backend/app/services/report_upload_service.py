import io

import pandas as pd
from fastapi import UploadFile


def dataframe_from_upload_bytes(filename: str, content: bytes) -> pd.DataFrame:
    lower_filename = filename.lower()
    stream = io.BytesIO(content)

    if lower_filename.endswith(".csv"):
        return drop_repeated_header_rows(pd.read_csv(stream))

    if lower_filename.endswith(".xlsx"):
        return drop_repeated_header_rows(pd.read_excel(stream))

    raise ValueError("Unsupported file format. Upload a .csv or .xlsx file.")


async def read_upload_dataframe(file: UploadFile) -> tuple[str, bytes, pd.DataFrame]:
    filename = file.filename or ""
    content = await file.read()
    dataframe = dataframe_from_upload_bytes(filename, content)
    return filename, content, dataframe


def drop_repeated_header_rows(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df

    header_labels = [str(column).strip() for column in df.columns]
    meaningful_indexes = [
        index
        for index, label in enumerate(header_labels)
        if label and not label.lower().startswith("unnamed:")
    ]

    if not meaningful_indexes:
        return df

    def is_repeated_header(row) -> bool:
        checked = 0

        for index in meaningful_indexes:
            value = row.iloc[index]

            if pd.isna(value) or str(value).strip() == "":
                continue

            checked += 1

            if str(value).strip() != header_labels[index]:
                return False

        return checked >= max(2, len(meaningful_indexes) // 2)

    repeated_header_mask = df.apply(is_repeated_header, axis=1)

    if not repeated_header_mask.any():
        return df

    return df.loc[~repeated_header_mask].reset_index(drop=True)
