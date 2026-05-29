from datetime import datetime

from pydantic import BaseModel, Field

class RuleTemplateItemPayload(BaseModel):
    keyword: str = Field(min_length=1, max_length=255)
    excel_column: str = Field(min_length=1, max_length=255)


class RuleTemplateCreate(BaseModel):
    name: str = Field(min_length=1, max_length=255)
    creator: str = Field(min_length=1, max_length=255)
    description: str = Field(default="", max_length=2000)
    items: list[RuleTemplateItemPayload] = Field(default_factory=list)


class RuleTemplateUpdate(BaseModel):
    name: str | None = Field(default=None, min_length=1, max_length=255)
    creator: str | None = Field(default=None, min_length=1, max_length=255)
    description: str | None = Field(default=None, max_length=2000)
    items: list[RuleTemplateItemPayload] | None = None


class RuleTemplateItemRead(BaseModel):
    id: int
    keyword: str
    excel_column: str
    order_index: int
    is_valid: bool

    class Config:
        from_attributes = True


class RuleTemplateRead(BaseModel):
    id: int
    name: str
    creator: str
    description: str
    is_active: bool
    created_at: datetime
    updated_at: datetime
    items: list[RuleTemplateItemRead]

    class Config:
        from_attributes = True


class RuleTemplateListItem(BaseModel):
    id: int
    name: str
    creator: str
    description: str
    is_active: bool
    created_at_bj: str
    updated_at_bj: str
    item_count: int


class ApplyTemplateRequest(BaseModel):
    excel_columns: list[str] = Field(default_factory=list)
    mode: str = Field(pattern="^(replace|append)$")


class ApplyTemplateResponse(BaseModel):
    valid_items: list[RuleTemplateItemPayload]
    invalid_items: list[RuleTemplateItemPayload]
