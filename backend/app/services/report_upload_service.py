import io
import os

import pandas as pd
from fastapi import UploadFile

# ---------------------------------------------------------------------------
# Upload constraints
# ---------------------------------------------------------------------------

MAX_UPLOAD_BYTES: int = int(os.getenv("MAX_UPLOAD_MB", "20")) * 1024 * 1024

# Browsers sometimes send application/octet-stream for .xlsx, so we keep it
# in the allowed set.  The extension check is the authoritative gate.
ALLOWED_CONTENT_TYPES: frozenset[str] = frozenset(
    {
        "text/csv",
        "text/plain",
        "application/csv",
        "application/vnd.ms-excel",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "application/octet-stream",
    }
)

ALLOWED_EXTENSIONS: tuple[str, ...] = (".csv", ".xlsx")


# ---------------------------------------------------------------------------
# Validation helpers
# ---------------------------------------------------------------------------


def validate_upload(
    filename: str,
    content: bytes,
    content_type: str | None,
) -> None:
    """Raise ValueError with a user-friendly message if the upload is invalid."""
    lower = filename.lower()

    if not any(lower.endswith(ext) for ext in ALLOWED_EXTENSIONS):
        raise ValueError(
            "Unsupported file format. Please upload a .csv or .xlsx file."
        )

    if len(content) > MAX_UPLOAD_BYTES:
        limit_mb = MAX_UPLOAD_BYTES // (1024 * 1024)
        raise ValueError(
            f"File is too large. The maximum allowed size is {limit_mb} MB. "
            "Please reduce the file size and try again."
        )

    if content_type and content_type.split(";")[0].strip() not in ALLOWED_CONTENT_TYPES:
        raise ValueError(
            f"Unexpected file type '{content_type}'. "
            "Please upload a CSV or Excel spreadsheet."
        )


# ---------------------------------------------------------------------------
# Core parsing
# ---------------------------------------------------------------------------


def dataframe_from_upload_bytes(filename: str, content: bytes) -> pd.DataFrame:
    lower_filename = filename.lower()
    stream = io.BytesIO(content)

    if lower_filename.endswith(".csv"):
        return drop_repeated_header_rows(pd.read_csv(stream))

    if lower_filename.endswith(".xlsx"):
        return drop_repeated_header_rows(pd.read_excel(stream))

    raise ValueError("Unsupported file format. Upload a .csv or .xlsx file.")


async def read_upload_dataframe(
    file: UploadFile,
) -> tuple[str, bytes, pd.DataFrame]:
    filename = file.filename or ""
    content = await file.read()

    validate_upload(filename, content, file.content_type)

    dataframe = dataframe_from_upload_bytes(filename, content)
    return filename, content, dataframe


# ---------------------------------------------------------------------------
# Data cleaning
# ---------------------------------------------------------------------------


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
