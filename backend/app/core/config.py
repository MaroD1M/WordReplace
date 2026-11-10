import os

from pydantic import BaseModel

DEFAULT_EXPORT_TOKEN_SECRET = "change-this-secret"


def _env_int(name: str, default: int) -> int:
    raw = os.getenv(name)
    if raw is None or raw.strip() == "":
        return default
    try:
        return int(raw)
    except ValueError:
        return default


def _parse_list(name: str, default: list[str]) -> list[str]:
    raw = os.getenv(name)
    if raw is None:
        return list(default)
    return [item.strip() for item in raw.split(",") if item.strip()]


class Settings(BaseModel):
    app_name: str = "WordReplace API"
    app_version: str = "3.0.0"
    database_url: str = "sqlite:////app/data/wordreplace.db"
    cors_allow_origins: list[str] = [
        "http://localhost:12344",
        "http://localhost:3000",
    ]
    allowed_hosts: list[str] = ["*"]
    max_upload_size_mb: int = 50
    max_excel_rows: int = 20000
    max_excel_cells: int = 200000
    max_replace_rules: int = 500
    max_preview_rows: int = 25
    max_preview_chars: int = 20000
    run_cache_ttl_seconds: int = 1800
    run_cache_max_entries: int = 200
    run_cache_max_bytes: int = 512 * 1024 * 1024
    max_concurrent_runs: int = 1
    export_token_secret: str = DEFAULT_EXPORT_TOKEN_SECRET

    @property
    def uses_default_secret(self) -> bool:
        return self.export_token_secret == DEFAULT_EXPORT_TOKEN_SECRET


settings = Settings(
    app_version=os.getenv("APP_VERSION", "3.0.0"),
    database_url=os.getenv("DATABASE_URL", "sqlite:////app/data/wordreplace.db"),
    cors_allow_origins=_parse_list(
        "CORS_ALLOW_ORIGINS",
        ["http://localhost:12344", "http://localhost:3000"],
    ),
    allowed_hosts=_parse_list("ALLOWED_HOSTS", ["*"]),
    max_upload_size_mb=_env_int("MAX_UPLOAD_SIZE_MB", 50),
    max_excel_rows=_env_int("MAX_EXCEL_ROWS", 20000),
    max_excel_cells=_env_int("MAX_EXCEL_CELLS", 200000),
    max_replace_rules=_env_int("MAX_REPLACE_RULES", 500),
    max_preview_rows=_env_int("MAX_PREVIEW_ROWS", 25),
    max_preview_chars=_env_int("MAX_PREVIEW_CHARS", 20000),
    run_cache_ttl_seconds=_env_int("RUN_CACHE_TTL_SECONDS", 1800),
    run_cache_max_entries=_env_int("RUN_CACHE_MAX_ENTRIES", 200),
    run_cache_max_bytes=_env_int("RUN_CACHE_MAX_BYTES", 512 * 1024 * 1024),
    max_concurrent_runs=_env_int("MAX_CONCURRENT_RUNS", 1),
    export_token_secret=os.getenv("EXPORT_TOKEN_SECRET", DEFAULT_EXPORT_TOKEN_SECRET),
)
