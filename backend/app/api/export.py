import io
from urllib.parse import quote

from fastapi import APIRouter, HTTPException, Query
from fastapi.responses import StreamingResponse

from app.services.replace_service import (
    delete_run_file,
    export_zip,
    get_run,
    get_run_file,
    merge_word_documents,
    verify_export_token,
)

router = APIRouter(prefix="/export", tags=["export"])


@router.get("/modes")
def export_modes() -> dict[str, list[str]]:
    return {"modes": ["zip", "merge"]}


@router.get("/zip/{run_id}")
def download_zip(run_id: str, token: str = Query(...)):
    if not verify_export_token(run_id, token):
        raise HTTPException(status_code=403, detail="导出令牌无效")
    run = get_run(run_id)
    if not run:
        raise HTTPException(status_code=404, detail="任务不存在或已过期")
    files = run["files"]
    if not files:
        raise HTTPException(status_code=400, detail="没有可导出的文件")

    data = export_zip(files)
    return StreamingResponse(
        data,
        media_type="application/zip",
        headers={"Content-Disposition": 'attachment; filename="replace_results.zip"'},
    )


@router.get("/merge/{run_id}")
def download_merge(run_id: str, token: str = Query(...)):
    if not verify_export_token(run_id, token):
        raise HTTPException(status_code=403, detail="导出令牌无效")
    run = get_run(run_id)
    if not run:
        raise HTTPException(status_code=404, detail="任务不存在或已过期")
    files = run["files"]
    if not files:
        raise HTTPException(status_code=400, detail="没有可导出的文件")

    data = merge_word_documents(files)
    return StreamingResponse(
        data,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": 'attachment; filename="replace_merged.docx"'},
    )


@router.get("/file/{run_id}/{item_id}")
def download_single_file(run_id: str, item_id: str, token: str = Query(...)):
    if not verify_export_token(run_id, token):
        raise HTTPException(status_code=403, detail="导出令牌无效")
    file = get_run_file(run_id, item_id)
    if not file:
        raise HTTPException(status_code=404, detail="文件不存在或已删除")
    # Use RFC 5987 filename* to support UTF-8 names safely across browsers/clients.
    encoded_name = quote(file.filename, safe="")
    return StreamingResponse(
        io.BytesIO(file.data.getvalue()),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={
            "Content-Disposition": f"attachment; filename*=UTF-8''{encoded_name}"
        },
    )


@router.delete("/result/{run_id}/{item_id}")
def delete_single_result(run_id: str, item_id: str, token: str = Query(...)):
    if not verify_export_token(run_id, token):
        raise HTTPException(status_code=403, detail="导出令牌无效")
    original = get_run(run_id)
    if original is None:
        raise HTTPException(status_code=404, detail="任务不存在或已过期")
    before_total = len(original.get("details", []))
    run = delete_run_file(run_id, item_id)
    if run is None:
        raise HTTPException(status_code=404, detail="任务不存在或已过期")
    deleted = len(run.get("details", [])) < before_total
    return {
        "deleted": deleted,
        "total": run["total"],
        "success": run["success"],
        "failed": run["failed"],
        "replacements": run["replacements"],
        "details": run["details"],
    }
