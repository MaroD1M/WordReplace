from fastapi import APIRouter, HTTPException
from fastapi.responses import StreamingResponse

from app.services.replace_service import export_zip, get_run, merge_word_documents

router = APIRouter(prefix="/export", tags=["export"])


@router.get("/modes")
def export_modes() -> dict[str, list[str]]:
    return {"modes": ["zip", "merge"]}


@router.get("/zip/{run_id}")
def download_zip(run_id: str):
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
def download_merge(run_id: str):
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
