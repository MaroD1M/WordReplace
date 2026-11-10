from fastapi import HTTPException, UploadFile

from app.core.config import settings

ALLOWED_WORD_EXT = {".docx"}
ALLOWED_EXCEL_EXT = {".xlsx", ".xls"}

ZIP_MAGIC = b"PK\x03\x04"
OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


def _max_bytes() -> int:
    return settings.max_upload_size_mb * 1024 * 1024


def validate_upload(upload: UploadFile, allowed_ext: set[str], label: str) -> None:
    filename = (upload.filename or "").lower()
    if not any(filename.endswith(ext) for ext in allowed_ext):
        raise HTTPException(status_code=400, detail=f"{label} 文件类型不支持")


def validate_size_hint(upload: UploadFile, label: str) -> None:
    size = getattr(upload, "size", None)
    if size is not None and size > _max_bytes():
        raise HTTPException(
            status_code=413,
            detail=f"{label} 超过 {settings.max_upload_size_mb}MB 限制",
        )


def validate_size(content: bytes, label: str) -> None:
    if len(content) > _max_bytes():
        raise HTTPException(
            status_code=413,
            detail=f"{label} 超过 {settings.max_upload_size_mb}MB 限制",
        )


def validate_magic(content: bytes, label: str, allowed_magics: tuple[bytes, ...]) -> None:
    if not content:
        raise HTTPException(status_code=400, detail=f"{label} 文件为空")
    if not any(content.startswith(magic) for magic in allowed_magics):
        raise HTTPException(status_code=400, detail=f"{label} 文件内容与扩展名不匹配")


async def read_upload_limited(upload: UploadFile, label: str) -> bytes:
    """Read an upload while enforcing the size limit during the read."""
    validate_size_hint(upload, label)
    max_bytes = _max_bytes()
    chunks: list[bytes] = []
    total = 0
    while True:
        chunk = await upload.read(1024 * 1024)
        if not chunk:
            break
        total += len(chunk)
        if total > max_bytes:
            raise HTTPException(
                status_code=413,
                detail=f"{label} 超过 {settings.max_upload_size_mb}MB 限制",
            )
        chunks.append(chunk)
    content = b"".join(chunks)
    validate_size(content, label)
    return content
