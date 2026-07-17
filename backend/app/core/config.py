from functools import lru_cache

from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    app_env: str = "production"
    frontend_origins: str = ""
    database_url: str | None = None
    report_database_url: str | None = None
    report_storage_root: str | None = None
    jwt_secret_key: str = "supersecretkeychangeitinproduction"
    jwt_algorithm: str = "HS256"
    jwt_access_token_expire_minutes: int = 1440

    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        extra="ignore",
    )

    @property
    def resolved_database_url(self) -> str | None:
        return self.report_database_url or self.database_url


@lru_cache
def get_settings() -> Settings:
    return Settings()
