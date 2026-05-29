from pydantic import BaseModel, Field


class ExecuteRule(BaseModel):
    keyword: str = Field(min_length=1, max_length=255)
    excel_column: str = Field(min_length=1, max_length=255)


class ReplaceRequest(BaseModel):
    start_row: int = Field(ge=1)
    end_row: int = Field(ge=1)
    file_name_column: str = Field(min_length=1, max_length=255)
    export_mode: str = Field(pattern="^(zip|merge)$")
    seq_format: str = Field(pattern="^(1|01|0001|1\\.|一)$")
    rules: list[ExecuteRule] = Field(min_length=1)


class ReplaceRowDetail(BaseModel):
    item_id: str
    seq: int
    row_number: int
    file_name: str
    status: str
    replace_count: int
    message: str = ""


class ReplaceSummary(BaseModel):
    run_id: str
    export_token: str
    total: int
    success: int
    failed: int
    replacements: int
    details: list[ReplaceRowDetail] = []


class ExecuteResponse(ReplaceSummary):
    pass
