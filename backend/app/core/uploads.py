from fastapi import HTTPException, UploadFile

from app.core.config import settings

ALLOWED_WORD_EXT = {".docx"}
ALLOWED_EXCEL_EXT = {".xlsx", ".xls"}


def validate_upload(upload: UploadFile, allowed_ext: set[str], label: str) -> None:
    filename = (upload.filename or "").lower()
    if not any(filename.endswith(ext) for ext in allowed_ext):
        raise HTTPException(status_code=400, detail=f"{label} 文件类型不支持")


def validate_size(content: bytes, label: str) -> None:
    max_bytes = settings.max_upload_size_mb * 1024 * 1024
    if len(content) > max_bytes:
        raise HTTPException(
            status_code=413,
            detail=f"{label} 超过 {settings.max_upload_size_mb}MB 限制",
        )
