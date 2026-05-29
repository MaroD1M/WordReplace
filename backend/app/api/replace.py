import json

from fastapi import APIRouter, Depends, File, Form, HTTPException, UploadFile

from app.core.config import settings
from app.schemas.replace import ExecuteResponse, ReplaceSummary
from app.services.replace_service import cache_run, run_replace, sign_export_token

router = APIRouter(prefix="/replace", tags=["replace"])

ALLOWED_WORD_EXT = {".docx"}
ALLOWED_EXCEL_EXT = {".xlsx", ".xls"}


def _validate_upload(upload: UploadFile, allowed_ext: set[str], label: str) -> None:
    filename = (upload.filename or "").lower()
    if not any(filename.endswith(ext) for ext in allowed_ext):
        raise HTTPException(status_code=400, detail=f"{label} 文件类型不支持")


def _validate_size(content: bytes, label: str) -> None:
    max_bytes = settings.max_upload_size_mb * 1024 * 1024
    if len(content) > max_bytes:
        raise HTTPException(status_code=413, detail=f"{label} 超过 {settings.max_upload_size_mb}MB 限制")


@router.post("/execute", response_model=ExecuteResponse)
async def execute_replace(
    word_file: UploadFile = File(...),
    excel_file: UploadFile = File(...),
    start_row: int = Form(...),
    end_row: int = Form(...),
    file_name_column: str = Form(...),
    export_mode: str = Form("zip"),
    rules_json: str = Form(...),
    use_seq_prefix: bool = Form(False),
    seq_format: str = Form("1"),
    seq_column: str = Form(""),
    name_prefix: str = Form(""),
    name_suffix: str = Form(""),
):
    if start_row > end_row:
        raise HTTPException(status_code=400, detail="起始行不能大于结束行")
    if export_mode not in {"zip", "merge"}:
        raise HTTPException(status_code=400, detail="不支持的导出模式")
    if seq_format not in {"1", "01", "0001", "1.", "一"}:
        raise HTTPException(status_code=400, detail="不支持的序号格式")
    _validate_upload(word_file, ALLOWED_WORD_EXT, "Word")
    _validate_upload(excel_file, ALLOWED_EXCEL_EXT, "Excel")

    try:
        parsed_rules = json.loads(rules_json)
        if not isinstance(parsed_rules, list):
            raise ValueError
        rules = []
        for item in parsed_rules:
            keyword = str(item.get("keyword", "")).strip()
            excel_column = str(item.get("excel_column", "")).strip()
            if not keyword or not excel_column:
                continue
            rules.append({"keyword": keyword, "excel_column": excel_column})
    except Exception:
        raise HTTPException(status_code=400, detail="规则参数格式错误")
    if not rules:
        raise HTTPException(status_code=400, detail="请先添加至少一条有效规则")

    word_bytes = await word_file.read()
    excel_bytes = await excel_file.read()
    _validate_size(word_bytes, "Word")
    _validate_size(excel_bytes, "Excel")

    result = run_replace(
        word_bytes=word_bytes,
        excel_bytes=excel_bytes,
        rules=rules,
        start_row=start_row,
        end_row=end_row,
        file_name_column=file_name_column,
        use_seq_prefix=use_seq_prefix,
        seq_format=seq_format,
        seq_column=seq_column.strip(),
        name_prefix=name_prefix,
        name_suffix=name_suffix,
    )
    run_id = cache_run(result)
    export_token = sign_export_token(run_id)

    return ExecuteResponse(
        run_id=run_id,
        export_token=export_token,
        total=result["total"],
        success=result["success"],
        failed=result["failed"],
        replacements=result["replacements"],
        details=result["details"],
    )


@router.post("/preview", response_model=ReplaceSummary)
def preview_replace() -> ReplaceSummary:
    raise HTTPException(status_code=410, detail="请改用 /replace/execute")
