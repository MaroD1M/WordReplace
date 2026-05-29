# Backend (FastAPI)

## Run

```bash
cd backend
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
uvicorn app.main:app --reload --port 8000
```

## Endpoints

- `GET /health`
- `POST /preview`
- `GET /rules`
- `POST /rules`
- `PUT /rules/{id}`
- `DELETE /rules/{id}`
- `POST /replace/execute`
- `GET /export/zip/{run_id}`
- `GET /export/merge/{run_id}`
- `GET /export/file/{run_id}/{item_id}`
- `DELETE /export/result/{run_id}/{item_id}`
- `GET /rule-templates`
- `GET /rule-templates/{id}`
- `POST /rule-templates`
- `PUT /rule-templates/{id}`
- `DELETE /rule-templates/{id}`
- `POST /rule-templates/{id}/apply`

## Notes

- 规则执行采用“会话规则”方式：前端执行时通过 `rules_json` 提交当前规则集合。  
- 模板库用于复用规则；应用模板时会按当前 Excel 列名进行有效性校验，不匹配规则自动忽略。  
