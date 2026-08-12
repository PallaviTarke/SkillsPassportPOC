"""
Curriculum-related API endpoints.
"""

import re
from typing import Optional

from fastapi import APIRouter, Query

from app.core.database import get_curriculum, get_curriculum_stats, get_col
from app.models.schemas import CurriculumResponse, CurriculumOptions

router = APIRouter(prefix="/api/curriculum", tags=["curriculum"])


@router.get("", response_model=CurriculumResponse)
async def api_curriculum(
    department: Optional[str] = Query(None),
    semester: Optional[str] = Query(None)
):
    """Get curriculum data by department and semester."""
    rows = get_curriculum(department, semester)
    return {"count": len(rows), "curriculum": rows}


@router.get("/stats")
async def api_curriculum_stats():
    """Get curriculum statistics."""
    return get_curriculum_stats()


@router.get("/options", response_model=CurriculumOptions)
async def api_curriculum_options():
    """Get available department and semester options."""
    pipeline = [
        {"$group": {
            "_id": {"department": "$department", "semester": "$semester"}
        }},
        {"$project": {"_id": 0,
            "department": "$_id.department",
            "semester": "$_id.semester"
        }},
        {"$sort": {"department": 1, "semester": 1}}
    ]
    rows = list(get_col().aggregate(pipeline))
    departments = sorted({r["department"] for r in rows if r.get("department")})
    semesters = sorted({r["semester"] for r in rows if r.get("semester")},
                       key=lambda s: (re.search(r"[0-9]+", s) or type("", (), {"group": lambda *a: s})).group(0))
    dept_sem_map: dict[str, list[str]] = {}
    for r in rows:
        d, s = r.get("department", ""), r.get("semester", "")
        if d and s:
            dept_sem_map.setdefault(d, [])
            if s not in dept_sem_map[d]:
                dept_sem_map[d].append(s)
    for d in dept_sem_map:
        dept_sem_map[d].sort(key=lambda s: (re.search(r"[0-9]+", s) or type("", (), {"group": lambda *a: s})).group(0))
    return {"departments": departments, "semesters": semesters, "dept_sem_map": dept_sem_map}
