from fastapi import APIRouter, UploadFile, File, Form
from fastapi.responses import StreamingResponse
import pandas as pd

from app.services.cleaner_core import clean_with_template
from app.services.wideload_generator import generate_wideload_report
from app.templates import vehicle_inspection

from app.templates import impounded_prohibited
from app.services.impounded_prohibited_generator import generate_impounded_prohibited_report

router = APIRouter()


@router.post("/process-file")
async def process_file(file: UploadFile = File(...)):
    filename = file.filename

    if filename.endswith(".csv"):
        df = pd.read_csv(file.file)
    elif filename.endswith(".xlsx"):
        df = pd.read_excel(file.file)
    else:
        return {"error": "Unsupported file format"}

    try:
        cleaned_df = clean_with_template(df, vehicle_inspection)
    except Exception as e:
        return {"error": str(e)}

    return {
        "columns": cleaned_df.columns.tolist(),
        "data": cleaned_df.head(5).to_dict(orient="records"),
    }


@router.post("/download-wideload-report")
async def download_wideload_report(
    file: UploadFile = File(...),
    report_date: str = Form(...),
    station: str = Form(...),
    bound: str = Form(...),
):
    filename = file.filename

    if filename.endswith(".csv"):
        df = pd.read_csv(file.file)
    elif filename.endswith(".xlsx"):
        df = pd.read_excel(file.file)
    else:
        return {"error": "Unsupported file format"}

    try:
        cleaned_df = clean_with_template(df, vehicle_inspection)

        file_stream = generate_wideload_report(
            cleaned_df,
            report_date=report_date,
            station=station,
            bound=bound,
        )

    except Exception as e:
        return {"error": str(e)}

    return StreamingResponse(
        file_stream,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={
            "Content-Disposition": "attachment; filename=wideload_report.docx"
        },
    )


@router.post("/download-impounded-prohibited-report")
async def download_impounded_prohibited_report(
    file: UploadFile = File(...),
    report_date: str = Form(...),
    station: str = Form(...),
    bound: str = Form(...),
):
    filename = file.filename

    if filename.endswith(".csv"):
        df = pd.read_csv(file.file)
    elif filename.endswith(".xlsx"):
        df = pd.read_excel(file.file)
    else:
        return {"error": "Unsupported file format"}

    try:
        cleaned_df = clean_with_template(df, impounded_prohibited)

        file_stream = generate_impounded_prohibited_report(
            cleaned_df,
            report_date=report_date,
            station=station,
            bound=bound,
        )

    except Exception as e:
        return {"error": str(e)}

    return StreamingResponse(
        file_stream,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={
            "Content-Disposition": "attachment; filename=impounded_prohibited_report.docx"
        },
    )