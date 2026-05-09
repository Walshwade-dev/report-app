import io

import pandas as pd
from fastapi import UploadFile


def dataframe_from_upload_bytes(filename: str, content: bytes) -> pd.DataFrame:
    lower_filename = filename.lower()
    stream = io.BytesIO(content)

    if lower_filename.endswith(".csv"):
        return pd.read_csv(stream)

    if lower_filename.endswith(".xlsx"):
        return pd.read_excel(stream)

    raise ValueError("Unsupported file format. Upload a .csv or .xlsx file.")


async def read_upload_dataframe(file: UploadFile) -> tuple[str, bytes, pd.DataFrame]:
    filename = file.filename or ""
    content = await file.read()
    dataframe = dataframe_from_upload_bytes(filename, content)
    return filename, content, dataframe
