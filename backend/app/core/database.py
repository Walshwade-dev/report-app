from collections.abc import Generator

from sqlalchemy import create_engine
from sqlalchemy.orm import DeclarativeBase, Session, sessionmaker

from app.core.config import get_settings


class Base(DeclarativeBase):
    pass


def normalize_database_url(database_url: str | None) -> str | None:
    if not database_url:
        return None

    if database_url.startswith("postgresql://"):
        return database_url.replace("postgresql://", "postgresql+psycopg://", 1)

    return database_url


settings = get_settings()
DATABASE_URL = normalize_database_url(settings.resolved_database_url)


def engine_options(database_url: str) -> dict:
    options = {
        "pool_pre_ping": True,
        "future": True,
    }

    if database_url.startswith("postgresql+psycopg://"):
        options["connect_args"] = {"connect_timeout": 3}

    return options


engine = (
    create_engine(
        DATABASE_URL,
        **engine_options(DATABASE_URL),
    )
    if DATABASE_URL
    else None
)

SessionLocal = (
    sessionmaker(
        bind=engine,
        autoflush=False,
        autocommit=False,
        expire_on_commit=False,
        class_=Session,
    )
    if engine is not None
    else None
)


def is_database_configured() -> bool:
    return SessionLocal is not None


def get_db() -> Generator[Session, None, None]:
    if SessionLocal is None:
        raise RuntimeError("DATABASE_URL is not configured.")

    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()
