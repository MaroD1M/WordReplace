import io

from docx import Document
from fastapi import APIRouter, File, HTTPException, UploadFile

from app.core.excel import read_excel
from app.core.uploads import (
    ALLOWED_EXCEL_EXT,
    ALLOWED_WORD_EXT,
    validate_size,
    validate_upload,
)
from app.schemas.preview import PreviewResponse
from app.services.replace_service import iter_document_paragraphs

router = APIRouter(prefix="/preview", tags=["preview"])


@router.post("", response_model=PreviewResponse)
async def preview_files(
    word_file: UploadFile = File(...),
    excel_file: UploadFile = File(...),
):
    validate_upload(word_file, ALLOWED_WORD_EXT, "Word")
    validate_upload(excel_file, ALLOWED_EXCEL_EXT, "Excel")

    word_bytes = await word_file.read()
    excel_bytes = await excel_file.read()
    validate_size(word_bytes, "Word")
    validate_size(excel_bytes, "Excel")

    try:
        doc = Document(io.BytesIO(word_bytes))
        paragraphs = [
            paragraph.text.strip()
            for paragraph in iter_document_paragraphs(doc)
            if paragraph.text and paragraph.text.strip()
        ]
        word_text = "\n".join(paragraphs)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Word 预览失败: {exc}")

    try:
        df = read_excel(excel_bytes)
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
