"""
Health check and system status endpoints.
"""

from fastapi import APIRouter

from app.core.config import GEMINI_API_KEY, OPENAI_API_KEY, DOCAI_PROJECT_ID, DOCAI_PROCESSOR_ID
from app.core.database import get_mongo, get_col, get_lightcast_col
from app.models.schemas import HealthResponse

router = APIRouter(prefix="/api", tags=["health"])


@router.get("/health", response_model=HealthResponse)
async def health():
    """Get system health status."""
    try:
        get_mongo().admin.command("ping")
        mongo_ok = True
    except Exception:
        mongo_ok = False

    lc_count = 0
    if mongo_ok:
        try:
            lc_count = get_lightcast_col().count_documents({})
        except Exception:
            pass

    return {
        "status": "ok",
        "gemini_key": bool(GEMINI_API_KEY),
        "openai_key": bool(OPENAI_API_KEY),
        "documentai": bool(DOCAI_PROJECT_ID and DOCAI_PROCESSOR_ID),
        "mongodb": mongo_ok,
        "curriculum_count": get_col().count_documents({}) if mongo_ok else 0,
        "lightcast_count": lc_count,
    }
