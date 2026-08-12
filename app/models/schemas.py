"""
Pydantic models for request and response validation.
"""

from typing import Optional, List, Dict, Any
from pydantic import BaseModel, Field


class HealthResponse(BaseModel):
    status: str
    gemini_key: bool
    openai_key: bool
    documentai: bool
    mongodb: bool
    curriculum_count: int
    lightcast_count: int


class CurriculumResponse(BaseModel):
    count: int
    curriculum: List[Dict[str, Any]]


class CurriculumOptions(BaseModel):
    departments: List[str]
    semesters: List[str]
    dept_sem_map: Dict[str, List[str]]


class ExtractionStartResponse(BaseModel):
    job_id: str


class PassportMetadata(BaseModel):
    model: str
    mode: str
    extraction_method: str
    department_filter: Optional[str]
    semester_filter: Optional[str]
    curriculum_skills: int
    lightcast_verified: int
    extracted_chars: int
    processed_at: str
    warning: Optional[str] = None


class PassportSummary(BaseModel):
    total_skills: int
    skills_verified: int
    skills_inferred: int
    lightcast_verified_skills: int
    average_score_percentage: Optional[float]
    top_skills: List[str]
    improvement_areas: List[str]


class StudentInfo(BaseModel):
    name: Optional[str] = None
    roll_no: Optional[str] = None
    department: Optional[str] = None
    semester: Optional[str] = None
    year: Optional[str] = None
    college: Optional[str] = None
    sgpa: Optional[float] = None
