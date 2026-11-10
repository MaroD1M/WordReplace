import logging
from contextlib import asynccontextmanager
from pathlib import Path

from fastapi import FastAPI, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.middleware.gzip import GZipMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from starlette.middleware.trustedhost import TrustedHostMiddleware

from app.api.config import router as config_router
from app.api.export import router as export_router
from app.api.health import router as health_router
from app.api.preview import router as preview_router
from app.api.replace import router as replace_router
from app.api.rules import router as rules_router
from app.api.templates import router as templates_router
from app.core.config import settings
from app.db.base import Base
from app.db.session import SessionLocal, engine
from app.services.template_service import ensure_legacy_rules_template

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(name)s %(message)s",
)
logger = logging.getLogger("wordreplace")


@asynccontextmanager
async def lifespan(_app: FastAPI):
    Base.metadata.create_all(bind=engine)
    with SessionLocal() as db:
        ensure_legacy_rules_template(db)
    if settings.uses_default_secret:
        logger.warning(
            "EXPORT_TOKEN_SECRET is using the default value. "
            "Set a strong random secret before exposing this service publicly."
        )
    logger.info("WordReplace %s started", settings.app_version)
    yield


app = FastAPI(title=settings.app_name, version=settings.app_version, lifespan=lifespan)

app.add_middleware(GZipMiddleware, minimum_size=1024)
app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.cors_allow_origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)
if settings.allowed_hosts != ["*"]:
    app.add_middleware(TrustedHostMiddleware, allowed_hosts=settings.allowed_hosts)

app.include_router(health_router)
app.include_router(config_router)
app.include_router(rules_router)
app.include_router(templates_router)
app.include_router(export_router)
app.include_router(preview_router)
app.include_router(replace_router)

# Single-image deployment mounts exported Next.js static assets here.
# /app/app/main.py -> parents[1] is /app
FRONTEND_DIST = Path(__file__).resolve().parents[1] / "frontend_dist"
HAS_FRONTEND = FRONTEND_DIST.exists()


@app.middleware("http")
async def add_security_headers(request: Request, call_next):
    response = await call_next(request)
    response.headers.setdefault("X-Content-Type-Options", "nosniff")
    response.headers.setdefault("X-Frame-Options", "DENY")
    response.headers.setdefault("Referrer-Policy", "no-referrer")
    response.headers.setdefault("Permissions-Policy", "geolocation=(), microphone=(), camera=()")
    if HAS_FRONTEND:
        connect_src = " ".join(["'self'", *settings.cors_allow_origins])
        response.headers.setdefault(
            "Content-Security-Policy",
            "default-src 'self'; "
            "img-src 'self' data:; "
            "style-src 'self' 'unsafe-inline'; "
            "script-src 'self' 'unsafe-inline'; "
            "font-src 'self' data:; "
            f"connect-src {connect_src}",
        )
    return response


if HAS_FRONTEND:
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
else:
    @app.get("/", include_in_schema=False)
    def no_frontend_built():
        return {
            "status": "ok",
            "message": "Frontend assets not found. Build and run via Docker image, or run Next.js frontend separately.",
            "health": "/health",
        }
