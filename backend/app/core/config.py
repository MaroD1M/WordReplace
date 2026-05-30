import os

from pydantic import BaseModel


class Settings(BaseModel):
    app_name: str = "WordReplace API"
    app_version: str = "2.0.5"
    database_url: str = "sqlite:////app/data/wordreplace.db"
    cors_allow_origins: list[str] = [
        "http://localhost:12344",
        "http://localhost:3000",
    ]
    max_upload_size_mb: int = 50
    run_cache_ttl_seconds: int = 1800
    run_cache_max_entries: int = 200
    export_token_secret: str = "change-this-secret"


def _parse_origins() -> list[str]:
    raw = os.getenv("CORS_ALLOW_ORIGINS", "http://localhost:12344,http://localhost:3000")
    return [i.strip() for i in raw.split(",") if i.strip()]


settings = Settings(
    app_version=os.getenv("APP_VERSION", "2.0.5"),
    cors_allow_origins=_parse_origins(),
    max_upload_size_mb=int(os.getenv("MAX_UPLOAD_SIZE_MB", "50")),
    run_cache_ttl_seconds=int(os.getenv("RUN_CACHE_TTL_SECONDS", "1800")),
    run_cache_max_entries=int(os.getenv("RUN_CACHE_MAX_ENTRIES", "200")),
    export_token_secret=os.getenv("EXPORT_TOKEN_SECRET", "change-this-secret"),
)
