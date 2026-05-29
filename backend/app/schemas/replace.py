from pydantic import BaseModel, Field


class ReplaceRequest(BaseModel):
    start_row: int = Field(ge=1)
    end_row: int = Field(ge=1)
    file_name_column: str = Field(min_length=1, max_length=255)
    export_mode: str = Field(pattern="^(zip|merge)$")


class ReplaceSummary(BaseModel):
    run_id: str
    export_token: str
    total: int
    success: int
    failed: int
    replacements: int


class ExecuteResponse(ReplaceSummary):
    pass
