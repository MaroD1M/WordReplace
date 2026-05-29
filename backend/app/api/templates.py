from fastapi import APIRouter, Depends, HTTPException, status
from sqlalchemy.orm import Session

from app.db.session import get_db
from app.schemas.template import (
    ApplyTemplateRequest,
    ApplyTemplateResponse,
    RuleTemplateCreate,
    RuleTemplateListItem,
    RuleTemplateRead,
    RuleTemplateUpdate,
)
from app.services.template_service import (
    apply_template_to_columns,
    create_template,
    delete_template,
    get_template,
    list_templates,
    update_template,
)

router = APIRouter(prefix="/rule-templates", tags=["rule-templates"])


@router.get("", response_model=list[RuleTemplateListItem])
def get_templates(db: Session = Depends(get_db)):
    return list_templates(db)


@router.get("/{template_id}", response_model=RuleTemplateRead)
def get_template_detail(template_id: int, db: Session = Depends(get_db)):
    tpl = get_template(db, template_id)
    if not tpl:
        raise HTTPException(status_code=404, detail="模板不存在")
    return tpl


@router.post("", response_model=RuleTemplateRead, status_code=status.HTTP_201_CREATED)
def post_template(payload: RuleTemplateCreate, db: Session = Depends(get_db)):
    return create_template(db, payload)


@router.put("/{template_id}", response_model=RuleTemplateRead)
def put_template(template_id: int, payload: RuleTemplateUpdate, db: Session = Depends(get_db)):
    tpl = update_template(db, template_id, payload)
    if not tpl:
        raise HTTPException(status_code=404, detail="模板不存在")
    return tpl


@router.delete("/{template_id}", status_code=status.HTTP_204_NO_CONTENT)
def remove_template(template_id: int, db: Session = Depends(get_db)):
    ok = delete_template(db, template_id)
    if not ok:
        raise HTTPException(status_code=404, detail="模板不存在")


@router.post("/{template_id}/apply", response_model=ApplyTemplateResponse)
def apply_template(template_id: int, payload: ApplyTemplateRequest, db: Session = Depends(get_db)):
    result = apply_template_to_columns(db, template_id, payload.excel_columns)
    if not result:
        raise HTTPException(status_code=404, detail="模板不存在")
    return result
