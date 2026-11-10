import io

from docx import Document
from fastapi import APIRouter, File, Form, HTTPException, UploadFile

from app.core.config import settings
from app.core.excel import list_sheets, read_excel
from app.core.uploads import (
    ALLOWED_EXCEL_EXT,
    ALLOWED_WORD_EXT,
    OLE_MAGIC,
    ZIP_MAGIC,
    read_upload_limited,
    validate_magic,
    validate_upload,
)
from app.schemas.preview import PreviewResponse
from app.services.replace_service import iter_document_paragraphs, paragraph_text

router = APIRouter(prefix="/preview", tags=["preview"])


@router.post("", response_model=PreviewResponse)
async def preview_files(
    word_file: UploadFile = File(...),
    excel_file: UploadFile = File(...),
    sheet_name: str = Form(""),
):
    validate_upload(word_file, ALLOWED_WORD_EXT, "Word")
    validate_upload(excel_file, ALLOWED_EXCEL_EXT, "Excel")

    word_bytes = await read_upload_limited(word_file, "Word")
    excel_bytes = await read_upload_limited(excel_file, "Excel")
    validate_magic(word_bytes, "Word", (ZIP_MAGIC,))
    validate_magic(excel_bytes, "Excel", (ZIP_MAGIC, OLE_MAGIC))

    try:
        doc = Document(io.BytesIO(word_bytes))
        paragraphs = []
        for paragraph in iter_document_paragraphs(doc):
            text = paragraph_text(paragraph).strip()
            if text:
                paragraphs.append(text)
        word_text = "\n".join(paragraphs)
        if len(word_text) > settings.max_preview_chars:
            word_text = word_text[: settings.max_preview_chars] + "\n…（预览已截断）"
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Word 预览失败: {exc}")

    try:
        sheets = list_sheets(excel_bytes)
        selected_sheet = sheet_name.strip()
        if selected_sheet and selected_sheet not in sheets:
            raise ValueError("指定的工作表不存在")
        df = read_excel(excel_bytes, selected_sheet or 0)
        selected_sheet = selected_sheet or (sheets[0] if sheets else "")
        excel_columns = list(df.columns)
        rows = []
        for _, row in df.head(settings.max_preview_rows).iterrows():
            rows.append([str(row[column]) for column in excel_columns])
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Excel 预览失败: {exc}")

    return PreviewResponse(
        word_text=word_text,
        excel_columns=excel_columns,
        excel_rows=rows,
        excel_total_rows=len(df),
        sheet_names=sheets,
        selected_sheet=selected_sheet,
    )
