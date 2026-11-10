from fastapi import APIRouter

from app.core.config import settings

router = APIRouter(prefix="/config", tags=["config"])


@router.get("")
def get_config() -> dict[str, int | str]:
    return {
        "app_version": settings.app_version,
        "max_upload_size_mb": settings.max_upload_size_mb,
        "max_excel_rows": settings.max_excel_rows,
        "max_replace_rules": settings.max_replace_rules,
        "max_preview_rows": settings.max_preview_rows,
        "run_cache_ttl_seconds": settings.run_cache_ttl_seconds,
    }
