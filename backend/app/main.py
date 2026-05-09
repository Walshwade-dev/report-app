import os

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from app.routes.upload import router as upload_router
from app.routes.reports import router as reports_router


app = FastAPI()

default_origins = [
    "http://localhost:3000",
    "http://127.0.0.1:3000",
]

frontend_origins = [
    origin.strip()
    for origin in os.getenv("FRONTEND_ORIGINS", "").split(",")
    if origin.strip()
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=[*default_origins, *frontend_origins],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/")
async def root():
    return {"status": "ok", "service": "report-app"}


@app.get("/health")
async def health():
    return {"status": "ok"}


app.include_router(upload_router, prefix="/api")
app.include_router(reports_router, prefix="/api")
