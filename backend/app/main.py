import os

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from sqlalchemy import text

from app.core.database import SessionLocal, is_database_configured
from app.routes.reports import router as reports_router
from app.services.report_session_store import report_session_store

app = FastAPI(
    title="Report App",
    description="Weighbridge daily report generation API",
    version="1.0.0",
)

# ---------------------------------------------------------------------------
# CORS
#
# In development (APP_ENV=development) localhost origins are added so the
# Next.js dev server can reach the backend without extra config.
#
# In production set FRONTEND_ORIGINS to a comma-separated list of the real
# frontend origins, e.g.:
#   FRONTEND_ORIGINS=https://dnkreport.netlify.app
# ---------------------------------------------------------------------------

_is_development = os.getenv("APP_ENV", "production") == "development"

_dev_origins: list[str] = (
    [
        "http://localhost:3000",
        "http://localhost:3001",
        "http://127.0.0.1:3000",
        "http://127.0.0.1:3001",
    ]
    if _is_development
    else []
)

_env_origins: list[str] = [
    origin.strip()
    for origin in os.getenv("FRONTEND_ORIGINS", "").split(",")
    if origin.strip()
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=[*_dev_origins, *_env_origins],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# ---------------------------------------------------------------------------
# Health / root
# ---------------------------------------------------------------------------


@app.get("/")
async def root():
    return {"status": "ok", "service": "report-app"}


@app.get("/health")
async def health():
    return {"status": "ok"}


@app.get("/health/persistence")
async def persistence_health():
    database_connected = False
    database_error: str | None = None

    if SessionLocal is not None:
        try:
            with SessionLocal() as session:
                session.execute(text("select 1"))
            database_connected = True
        except Exception as exc:
            database_error = exc.__class__.__name__

    storage_root = report_session_store.storage_root

    return {
        "status": "ok",
        "database": {
            "configured": is_database_configured(),
            "connected": database_connected,
            "error": database_error,
        },
        "storage": {
            "root": str(storage_root),
            "exists": storage_root.exists(),
            "sessions_dir_exists": report_session_store.sessions_dir.exists(),
            "uploads_dir_exists": report_session_store.uploads_dir.exists(),
            "processed_dir_exists": report_session_store.processed_dir.exists(),
            "final_reports_dir_exists": report_session_store.final_reports_dir.exists(),
        },
    }


# ---------------------------------------------------------------------------
# Routers
# ---------------------------------------------------------------------------

app.include_router(reports_router, prefix="/api")
