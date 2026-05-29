from sqlalchemy.orm import Session

from app.models.rule import Rule
from app.schemas.rule import RuleCreate


def list_rules(db: Session) -> list[Rule]:
    return db.query(Rule).order_by(Rule.id.asc()).all()


def create_rule(db: Session, payload: RuleCreate) -> Rule:
    entity = Rule(keyword=payload.keyword.strip(), excel_column=payload.excel_column.strip())
    db.add(entity)
    db.commit()
    db.refresh(entity)
    return entity


def delete_rule(db: Session, rule_id: int) -> bool:
    entity = db.get(Rule, rule_id)
    if not entity:
        return False
    db.delete(entity)
    db.commit()
    return True
