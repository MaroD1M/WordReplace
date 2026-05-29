from fastapi import APIRouter, Depends, File, Form, HTTPException, UploadFile
from sqlalchemy.orm import Session

from app.db.session import get_db
from app.schemas.replace import ExecuteResponse, ReplaceSummary
from app.services.replace_service import cache_run, run_replace
from app.services.rule_service import list_rules

router = APIRouter(prefix="/replace", tags=["replace"])


@router.post("/execute", response_model=ExecuteResponse)
async def execute_replace(
    word_file: UploadFile = File(...),
    excel_file: UploadFile = File(...),
    start_row: int = Form(...),
    end_row: int = Form(...),
    file_name_column: str = Form(...),
    export_mode: str = Form("zip"),
    db: Session = Depends(get_db),
):
    if start_row > end_row:
        raise HTTPException(status_code=400, detail="起始行不能大于结束行")
    if export_mode not in {"zip", "merge"}:
        raise HTTPException(status_code=400, detail="不支持的导出模式")

    rules = list_rules(db)
    if not rules:
        raise HTTPException(status_code=400, detail="请先添加至少一条规则")

    word_bytes = await word_file.read()
    excel_bytes = await excel_file.read()

    result = run_replace(
        word_bytes=word_bytes,
        excel_bytes=excel_bytes,
        rules=rules,
        start_row=start_row,
        end_row=end_row,
        file_name_column=file_name_column,
    )
    run_id = cache_run(result)

    return ExecuteResponse(
        run_id=run_id,
        total=result["total"],
        success=result["success"],
        failed=result["failed"],
        replacements=result["replacements"],
    )


@router.post("/preview", response_model=ReplaceSummary)
def preview_replace() -> ReplaceSummary:
    raise HTTPException(status_code=410, detail="请改用 /replace/execute")
