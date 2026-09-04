import os

from dotenv import load_dotenv
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
from sqlalchemy import text
load_dotenv()

from app.core.database import SessionLocal, is_database_configured
from app.routes.reports import router as reports_router
from app.routes.weekly_reports import router as weekly_reports_router
from app.routes.auth import router as auth_router, users_router
from app.routes.mobile_checklist import router as mobile_checklist_router
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


def _requires_persistence() -> bool:
    return os.getenv("APP_ENV", "production") == "production"


@app.get("/health/persistence")
async def persistence_health():
    database_connected = False
    database_error: str | None = None

    if SessionLocal is not None:
        try:
            with SessionLocal() as session:
                session.execute(text("select 1"))
                session.execute(text("select 1 from reports limit 1"))
            database_connected = True
        except Exception as exc:
            database_error = exc.__class__.__name__

    storage_root = report_session_store.storage_root
    storage_dirs_exist = {
        "sessions": report_session_store.sessions_dir.exists(),
        "uploads": report_session_store.uploads_dir.exists(),
        "processed": report_session_store.processed_dir.exists(),
        "final_reports": report_session_store.final_reports_dir.exists(),
    }
    storage_configured = bool(os.getenv("REPORT_STORAGE_ROOT"))
    persistence_required = _requires_persistence()
    persistence_ready = (
        database_connected
        and storage_configured
        and storage_root.exists()
        and all(storage_dirs_exist.values())
    )

    payload = {
        "status": "ok" if persistence_ready or not persistence_required else "error",
        "persistence_required": persistence_required,
        "database": {
            "configured": is_database_configured(),
            "connected": database_connected,
            "error": database_error,
        },
        "storage": {
            "configured": storage_configured,
            "root": str(storage_root),
            "exists": storage_root.exists(),
            "sessions_dir_exists": storage_dirs_exist["sessions"],
            "uploads_dir_exists": storage_dirs_exist["uploads"],
            "processed_dir_exists": storage_dirs_exist["processed"],
            "final_reports_dir_exists": storage_dirs_exist["final_reports"],
        },
    }

    if persistence_required and not persistence_ready:
        return JSONResponse(status_code=503, content=payload)

    return payload


# ---------------------------------------------------------------------------
# Routers
# ---------------------------------------------------------------------------

app.include_router(reports_router, prefix="/api")
app.include_router(weekly_reports_router, prefix="/api")
app.include_router(auth_router, prefix="/api")
app.include_router(users_router, prefix="/api")
app.include_router(mobile_checklist_router)
