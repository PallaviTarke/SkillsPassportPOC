"""
Skill passport generation API endpoints.
"""

import io
import uuid
from datetime import datetime, timezone

from fastapi import APIRouter, File, Form, UploadFile, HTTPException
from fastapi.responses import StreamingResponse

from app.core.config import GEMINI_MODEL, _passport_store, log
from app.core.database import get_curriculum
from app.services.passport_service import (
    extract_marksheet, build_passport_with_curriculum,
    build_passport_no_curriculum, create_passport_excel
)
from app.services.extraction_service import extract_full_text

router = APIRouter(prefix="/api/passport", tags=["passport"])


@router.post("")
async def api_generate_passport(
    marksheet: UploadFile = File(...),
    department: str = Form(""),
    semester: str = Form("")
):
    """Generate a skill passport from a marksheet PDF."""
    if not marksheet.filename:
        raise HTTPException(status_code=400, detail="No marksheet PDF. Use key 'marksheet'")

    if not marksheet.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Only PDF marksheets are supported")

    pdf_bytes = await marksheet.read()
    log.info("Marksheet: %s (%d bytes)", marksheet.filename, len(pdf_bytes))

    try:
        extracted_text, method = extract_full_text(pdf_bytes)
        log.info("Extraction via %s → %d chars", method, len(extracted_text))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Text extraction failed: {e}")

    try:
        marksheet_data = extract_marksheet(extracted_text)
    except Exception as e:
        log.error("Marksheet parse: %s", e)
        raise HTTPException(status_code=500, detail=f"Marksheet parsing failed: {e}")

    no_curriculum_mode = False
    curriculum_fallback_note = None

    dept = department.strip() or None
    sem = semester.strip() or None

    if not dept and marksheet_data.get("student_info", {}).get("department"):
        dept = marksheet_data["student_info"]["department"].strip() or None
        log.info("Auto-detected department from marksheet: %s", dept)
    if not sem and marksheet_data.get("student_info", {}).get("semester"):
        sem = marksheet_data["student_info"]["semester"].strip() or None
        log.info("Auto-detected semester from marksheet: %s", sem)

    if not dept or not sem:
        log.info("No dept/sem (user or marksheet) → marksheet-only mode (all unverified)")
        no_curriculum_mode = True
    else:
        curriculum = get_curriculum(dept, sem)
        if not curriculum:
            log.info("No curriculum rows for %s / %s → marksheet-only mode", dept, sem)
            no_curriculum_mode = True
            curriculum_fallback_note = (
                f"No curriculum found for {dept} / {sem}. "
                "Skills inferred from marksheet subjects only."
            )
        else:
            log.info("Curriculum: %d skills for %s / %s", len(curriculum), dept, sem)

    try:
        if no_curriculum_mode:
            passport = build_passport_no_curriculum(marksheet_data)
        else:
            passport = build_passport_with_curriculum(marksheet_data, curriculum)
    except Exception as e:
        log.error("Passport build: %s", e)
        raise HTTPException(status_code=500, detail=f"Passport assembly failed: {e}")

    try:
        excel_buf = create_passport_excel(passport)
        passport_id = str(uuid.uuid4())
        _passport_store.put(passport_id, excel_buf.getvalue())
        passport["_download_id"] = passport_id
    except Exception as e:
        log.warning("Passport Excel build failed: %s", e)

    lc_count = passport.get("summary", {}).get("lightcast_verified_skills", 0)
    passport["_meta"] = {
        "model": GEMINI_MODEL,
        "mode": "no_curriculum" if no_curriculum_mode else "curriculum_matched",
        "extraction_method": method,
        "department_filter": dept,
        "semester_filter": sem,
        "curriculum_skills": 0 if no_curriculum_mode else len(curriculum),
        "lightcast_verified": lc_count,
        "extracted_chars": len(extracted_text),
        "processed_at": datetime.now(timezone.utc).isoformat(),
        "warning": curriculum_fallback_note,
    }
    return passport


@router.get("/download/{passport_id}")
async def passport_download(passport_id: str):
    """Download a generated passport Excel file."""
    excel_bytes = _passport_store.pop(passport_id)
    if excel_bytes is None:
        raise HTTPException(status_code=404, detail="Download expired or not found.")

    fname = f"skill_passport_{passport_id[:8]}.xlsx"
    return StreamingResponse(
        io.BytesIO(excel_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={fname}"}
    )
