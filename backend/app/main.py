from pathlib import Path

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

from app.api.export import router as export_router
from app.api.health import router as health_router
from app.api.replace import router as replace_router
from app.api.rules import router as rules_router
from app.core.config import settings
from app.db.base import Base
from app.db.session import engine

Base.metadata.create_all(bind=engine)

app = FastAPI(title=settings.app_name, version=settings.app_version)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(health_router)
app.include_router(rules_router)
app.include_router(export_router)
app.include_router(replace_router)

# Single-image deployment mounts exported Next.js static assets here.
FRONTEND_DIST = Path(__file__).resolve().parents[2] / "frontend_dist"
if FRONTEND_DIST.exists():
    app.mount("/_next", StaticFiles(directory=FRONTEND_DIST / "_next"), name="next-assets")

    @app.get("/", include_in_schema=False)
    def serve_index():
        return FileResponse(FRONTEND_DIST / "index.html")

    @app.get("/{full_path:path}", include_in_schema=False)
    def serve_frontend(full_path: str):
        target = FRONTEND_DIST / full_path
        if target.exists() and target.is_file():
            return FileResponse(target)
        return FileResponse(FRONTEND_DIST / "index.html")
