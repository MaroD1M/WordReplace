from pydantic import BaseModel


class PreviewResponse(BaseModel):
    word_text: str
    excel_columns: list[str]
    excel_rows: list[list[str]]
    excel_total_rows: int
