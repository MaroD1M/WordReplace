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
- `GET /rules`
- `POST /rules`
- `DELETE /rules/{id}`
- `POST /replace/execute`
- `GET /export/zip/{run_id}`
- `GET /export/merge/{run_id}`
