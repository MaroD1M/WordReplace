from fastapi import APIRouter, Depends, HTTPException, status
from sqlalchemy.orm import Session

from app.db.session import get_db
from app.schemas.rule import RuleCreate, RuleRead
from app.services.rule_service import create_rule, delete_rule, list_rules

router = APIRouter(prefix="/rules", tags=["rules"])


@router.get("", response_model=list[RuleRead])
def get_rules(db: Session = Depends(get_db)):
    return list_rules(db)


@router.post("", response_model=RuleRead, status_code=status.HTTP_201_CREATED)
def post_rule(payload: RuleCreate, db: Session = Depends(get_db)):
    return create_rule(db, payload)


@router.delete("/{rule_id}", status_code=status.HTTP_204_NO_CONTENT)
def remove_rule(rule_id: int, db: Session = Depends(get_db)):
    deleted = delete_rule(db, rule_id)
    if not deleted:
        raise HTTPException(status_code=404, detail="Rule not found")
