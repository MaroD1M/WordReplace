import asyncio
import json
import logging

from fastapi import APIRouter, File, Form, HTTPException, UploadFile
from fastapi.concurrency import run_in_threadpool

from app.core.config import settings
from app.core.uploads import (
    ALLOWED_EXCEL_EXT,
    ALLOWED_WORD_EXT,
    OLE_MAGIC,
    ZIP_MAGIC,
    read_upload_limited,
    validate_magic,
    validate_upload,
)
from app.schemas.replace import ExecuteResponse, ReplaceSummary
from app.services.replace_service import cache_run, run_replace, sign_export_token

router = APIRouter(prefix="/replace", tags=["replace"])
logger = logging.getLogger("wordreplace")

_RUN_SEMAPHORE = asyncio.Semaphore(max(1, settings.max_concurrent_runs))


def _parse_rules(rules_json: str) -> list[dict[str, str]]:
    try:
        parsed_rules = json.loads(rules_json)
        if not isinstance(parsed_rules, list):
            raise ValueError
    except Exception:
        raise HTTPException(status_code=400, detail="规则参数格式错误")

    if len(parsed_rules) > settings.max_replace_rules:
        raise HTTPException(
            status_code=400,
            detail=f"规则数量超过限制（最多 {settings.max_replace_rules} 条）",
        )

    rules = []
    for item in parsed_rules:
        if not isinstance(item, dict):
            continue
        keyword = str(item.get("keyword", "")).strip()
        excel_column = str(item.get("excel_column", "")).strip()
        if not keyword or not excel_column:
            continue
        if len(keyword) > 255 or len(excel_column) > 255:
            raise HTTPException(status_code=400, detail="规则关键字或列名过长")
        rules.append({"keyword": keyword, "excel_column": excel_column})
    if not rules:
        raise HTTPException(status_code=400, detail="请先添加至少一条有效规则")
    return rules


@router.post("/execute", response_model=ExecuteResponse)
async def execute_replace(
    word_file: UploadFile = File(...),
    excel_file: UploadFile = File(...),
    start_row: int = Form(..., ge=1),
    end_row: int = Form(..., ge=1),
    file_name_column: str = Form(..., min_length=1, max_length=255),
    export_mode: str = Form("zip"),
    rules_json: str = Form(...),
    use_seq_prefix: bool = Form(False),
    seq_format: str = Form("1"),
    seq_column: str = Form(""),
    name_prefix: str = Form(""),
    name_suffix: str = Form(""),
    sheet_name: str = Form(""),
    keep_placeholder_on_empty: bool = Form(False),
):
    if start_row > end_row:
        raise HTTPException(status_code=400, detail="起始行不能大于结束行")
    if export_mode not in {"zip", "merge"}:
        raise HTTPException(status_code=400, detail="不支持的导出模式")
    if seq_format not in {"1", "01", "0001", "1.", "一"}:
        raise HTTPException(status_code=400, detail="不支持的序号格式")
    validate_upload(word_file, ALLOWED_WORD_EXT, "Word")
    validate_upload(excel_file, ALLOWED_EXCEL_EXT, "Excel")

    rules = _parse_rules(rules_json)

    word_bytes = await read_upload_limited(word_file, "Word")
    excel_bytes = await read_upload_limited(excel_file, "Excel")
    validate_magic(word_bytes, "Word", (ZIP_MAGIC,))
    validate_magic(excel_bytes, "Excel", (ZIP_MAGIC, OLE_MAGIC))

    async with _RUN_SEMAPHORE:
        try:
            result = await run_in_threadpool(
                run_replace,
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
                sheet_name=sheet_name.strip(),
                keep_placeholder_on_empty=keep_placeholder_on_empty,
            )
            run_id = cache_run(result)
        except ValueError as exc:
            raise HTTPException(status_code=400, detail=str(exc))
        except Exception as exc:
            logger.exception("replace execution failed")
            raise HTTPException(status_code=400, detail=f"文件解析失败: {exc}")

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
