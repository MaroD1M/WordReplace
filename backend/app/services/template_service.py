from datetime import timezone
from zoneinfo import ZoneInfo

from sqlalchemy.orm import Session, selectinload

from app.models.rule import Rule
from app.models.template import RuleTemplate, RuleTemplateItem
from app.schemas.template import (
    ApplyTemplateResponse,
    RuleTemplateCreate,
    RuleTemplateItemPayload,
    RuleTemplateUpdate,
)


def _bj_str(dt):
    return dt.astimezone(timezone.utc).astimezone(ZoneInfo("Asia/Shanghai")).strftime("%Y-%m-%d %H:%M:%S")


def ensure_legacy_rules_template(db: Session) -> None:
    exists = db.query(RuleTemplate).filter(RuleTemplate.name == "历史规则导入", RuleTemplate.is_active.is_(True)).first()
    if exists:
        return
    rules = db.query(Rule).order_by(Rule.id.asc()).all()
    if not rules:
        return
    tpl = RuleTemplate(name="历史规则导入", creator="system", description="从旧版全局规则自动迁移")
    db.add(tpl)
    db.flush()
    for idx, r in enumerate(rules):
        db.add(RuleTemplateItem(template_id=tpl.id, keyword=r.keyword, excel_column=r.excel_column, order_index=idx, is_valid=True))
    db.commit()


def list_templates(db: Session) -> list[dict]:
    rows = db.query(RuleTemplate).filter(RuleTemplate.is_active.is_(True)).order_by(RuleTemplate.updated_at.desc()).all()
    return [
        {
            "id": t.id,
            "name": t.name,
            "creator": t.creator,
            "description": t.description,
            "is_active": t.is_active,
            "created_at_bj": _bj_str(t.created_at),
            "updated_at_bj": _bj_str(t.updated_at),
            "item_count": len(t.items),
        }
        for t in rows
    ]


def get_template(db: Session, template_id: int) -> RuleTemplate | None:
    return (
        db.query(RuleTemplate)
        .options(selectinload(RuleTemplate.items))
        .filter(RuleTemplate.id == template_id, RuleTemplate.is_active.is_(True))
        .first()
    )


def create_template(db: Session, payload: RuleTemplateCreate) -> RuleTemplate:
    tpl = RuleTemplate(name=payload.name.strip(), creator=payload.creator.strip(), description=payload.description.strip())
    db.add(tpl)
    db.flush()
    for idx, item in enumerate(payload.items):
        db.add(
            RuleTemplateItem(
                template_id=tpl.id,
                keyword=item.keyword.strip(),
                excel_column=item.excel_column.strip(),
                order_index=idx,
                is_valid=True,
            )
        )
    db.commit()
    db.refresh(tpl)
    return get_template(db, tpl.id)  # type: ignore[return-value]


def update_template(db: Session, template_id: int, payload: RuleTemplateUpdate) -> RuleTemplate | None:
    tpl = get_template(db, template_id)
    if not tpl:
        return None
    if payload.name is not None:
        tpl.name = payload.name.strip()
    if payload.creator is not None:
        tpl.creator = payload.creator.strip()
    if payload.description is not None:
        tpl.description = payload.description.strip()
    if payload.items is not None:
        for it in list(tpl.items):
            db.delete(it)
        db.flush()
        for idx, item in enumerate(payload.items):
            db.add(
                RuleTemplateItem(
                    template_id=tpl.id,
                    keyword=item.keyword.strip(),
                    excel_column=item.excel_column.strip(),
                    order_index=idx,
                    is_valid=True,
                )
            )
    db.commit()
    db.refresh(tpl)
    return get_template(db, tpl.id)


def delete_template(db: Session, template_id: int) -> bool:
    tpl = get_template(db, template_id)
    if not tpl:
        return False
    tpl.is_active = False
    db.commit()
    return True


def apply_template_to_columns(db: Session, template_id: int, excel_columns: list[str]) -> ApplyTemplateResponse | None:
    tpl = get_template(db, template_id)
    if not tpl:
        return None
    valid: list[RuleTemplateItemPayload] = []
    invalid: list[RuleTemplateItemPayload] = []
    column_set = {str(c).strip() for c in excel_columns}
    for item in sorted(tpl.items, key=lambda x: x.order_index):
        payload = RuleTemplateItemPayload(keyword=item.keyword, excel_column=item.excel_column)
        if item.excel_column in column_set:
            valid.append(payload)
        else:
            invalid.append(payload)
    return ApplyTemplateResponse(valid_items=valid, invalid_items=invalid)
