import io

import pandas as pd
from docx import Document
from fastapi import APIRouter, File, HTTPException, UploadFile

from app.schemas.preview import PreviewResponse

router = APIRouter(prefix="/preview", tags=["preview"])


@router.post("", response_model=PreviewResponse)
async def preview_files(
    word_file: UploadFile = File(...),
    excel_file: UploadFile = File(...),
):
    if not (word_file.filename or "").lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="Word 文件类型不支持")
    if not any((excel_file.filename or "").lower().endswith(ext) for ext in (".xlsx", ".xls")):
        raise HTTPException(status_code=400, detail="Excel 文件类型不支持")

    word_bytes = await word_file.read()
    excel_bytes = await excel_file.read()

    try:
        doc = Document(io.BytesIO(word_bytes))
        paragraphs = [p.text.strip() for p in doc.paragraphs if p.text and p.text.strip()]
        word_text = "\n".join(paragraphs)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Word 预览失败: {exc}")

    try:
        df = pd.read_excel(io.BytesIO(excel_bytes)).fillna("")
        df.columns = [str(c).strip() for c in df.columns]
        excel_columns = list(df.columns)
        rows = []
        for _, row in df.head(25).iterrows():
            rows.append([str(row[c]) for c in excel_columns])
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Excel 预览失败: {exc}")

    return PreviewResponse(
        word_text=word_text,
        excel_columns=excel_columns,
        excel_rows=rows,
        excel_total_rows=len(df),
    )
