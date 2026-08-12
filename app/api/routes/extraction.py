"""
Curriculum extraction API endpoints.
"""

import io
import queue
import re
import uuid

from fastapi import APIRouter, File, Form, UploadFile, HTTPException, Query
from fastapi.responses import StreamingResponse

from app.core.config import GEMINI_API_KEY, _job_store, _result_store, _result_lock, _executor
from app.models.schemas import ExtractionStartResponse
from app.services.extraction_service import run_extraction_job

router = APIRouter(prefix="/api/extract", tags=["extraction"])


@router.post("/start", response_model=ExtractionStartResponse)
async def extract_start(
    pdf: UploadFile = File(...),
    department: str = Form(...),
    year: str = Form(...),
    semester: str = Form(...),
    custom_department: str = Form("")
):
    """Start a curriculum extraction job."""
    try:
        dept = custom_department.strip() if custom_department.strip() else department.strip()
        yr = year.strip()
        sem = semester.strip()

        if not all([dept, yr, sem]):
            raise HTTPException(status_code=400, detail="department, year, semester are required.")

        if not pdf.filename:
            raise HTTPException(status_code=400, detail="No PDF file attached.")

        if not GEMINI_API_KEY:
            raise HTTPException(status_code=500, detail="GEMINI_API_KEY is not configured.")

        pdf_bytes = await pdf.read()
        job_id = str(uuid.uuid4())
        _job_store.create(job_id)
        _executor.submit(run_extraction_job, job_id, pdf_bytes, dept, yr, sem, True)

        return {"job_id": job_id}

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@router.get("/stream/{job_id}")
async def extract_stream(job_id: str):
    """Stream extraction progress events."""
    q = _job_store.get(job_id)
    if q is None:
        raise HTTPException(status_code=404, detail="Unknown job ID")

    async def generate():
        while True:
            try:
                msg = q.get(timeout=20)
            except queue.Empty:
                yield ": keep-alive\n\n"
                continue
            if msg is None:
                _job_store.delete(job_id)
                break
            yield msg

    return StreamingResponse(
        generate(),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"}
    )


@router.get("/download/{job_id}")
async def extract_download(
    job_id: str,
    dept: str = Query("dept"),
    yr: str = Query("yr"),
    sem: str = Query("sem")
):
    """Download extracted curriculum Excel file."""
    with _result_lock:
        excel_bytes = _result_store.pop(job_id)

    if excel_bytes is None:
        raise HTTPException(status_code=404, detail="Result not found or already downloaded.")

    dept_clean = dept.replace(" ", "_")
    yr_clean = yr.replace(" ", "_")
    sem_clean = sem.replace(" ", "_")
    fname = re.sub(r"[^A-Za-z0-9_\-]", "", f"{dept_clean}_{yr_clean}_{sem_clean}_skills") + ".xlsx"

    return StreamingResponse(
        io.BytesIO(excel_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={fname}"}
    )
