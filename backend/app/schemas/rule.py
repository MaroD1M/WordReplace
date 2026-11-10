from pydantic import BaseModel, ConfigDict, Field


class RuleCreate(BaseModel):
    keyword: str = Field(min_length=1, max_length=255)
    excel_column: str = Field(min_length=1, max_length=255)


class RuleUpdate(BaseModel):
    keyword: str = Field(min_length=1, max_length=255)
    excel_column: str = Field(min_length=1, max_length=255)


class RuleRead(BaseModel):
    model_config = ConfigDict(from_attributes=True)

    id: int
    keyword: str
    excel_column: str
