"""
Main FastAPI application entry point.
"""

from pathlib import Path

from fastapi import FastAPI, Request
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from fastapi.middleware.cors import CORSMiddleware

from app.api.routes import health, curriculum, extraction, passport
from app.core.database import ensure_indexes, get_col
from app.core import config
from app.core.config import log

# ══════════════════════════════════════════════════════════════════════════════
# FASTAPI APP INITIALIZATION
# ══════════════════════════════════════════════════════════════════════════════

app = FastAPI(
    title="Skill Passport Generator",
    description="Curriculum skill extraction and student skill passport generation",
    version="2.0.0"
)

# CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Templates configuration
BASE_DIR = Path(__file__).parent.parent
templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))

# Static files (if they exist)
if (BASE_DIR / "static").exists():
    app.mount("/static", StaticFiles(directory=str(BASE_DIR / "static")), name="static")

# ══════════════════════════════════════════════════════════════════════════════
# INCLUDE ROUTERS
# ══════════════════════════════════════════════════════════════════════════════

app.include_router(health.router)
app.include_router(curriculum.router)
app.include_router(extraction.router)
app.include_router(passport.router)

# ══════════════════════════════════════════════════════════════════════════════
# ROOT ENDPOINTS
# ══════════════════════════════════════════════════════════════════════════════

@app.get("/")
async def index(request: Request):
    """Root route"""
    return {
        "name": config.APP_NAME,
        "version": config.APP_VERSION,
        "status": "running",
        "docs": "/docs" 
    }

# ══════════════════════════════════════════════════════════════════════════════
# STARTUP AND SHUTDOWN
# ══════════════════════════════════════════════════════════════════════════════

@app.on_event("startup")
async def startup_event():
    """Initialize database indexes and check curriculum data."""
    try:
        ensure_indexes()
    except Exception as e:
        log.error("MongoDB startup: %s", e)
        return

    count = get_col().count_documents({})
    log.info("curriculum collection: %d documents", count)

    if count == 0:
        log.warning("No curriculum data found in MongoDB. Please populate the curriculum collection directly.")


@app.on_event("shutdown")
async def shutdown_event():
    """Cleanup on application shutdown."""
    log.info("Application shutting down...")
