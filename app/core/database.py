"""
Database operations module for MongoDB and LightCast helpers.
"""

from __future__ import annotations

import logging
import re
from datetime import datetime, timezone
from threading import Lock

import numpy as np
from pymongo import MongoClient, ASCENDING, UpdateOne
from pymongo.errors import BulkWriteError

from app.core.config import (
    MONGO_URI, MONGO_DB_NAME,
    OPENAI_API_KEY, OPENAI_EMBED_MODEL,
    LIGHTCAST_DB_NAME, LIGHTCAST_COLLECTION,
    get_openai_client, _call_gemini_json, log,
)

# ══════════════════════════════════════════════════════════════════════════════
# OPENAI EMBEDDINGS
# ══════════════════════════════════════════════════════════════════════════════

def generate_openai_embeddings(texts: list) -> list:
    """Batch-generate OpenAI embeddings."""
    if not texts or not OPENAI_API_KEY:
        return [[] for _ in texts]
    client = get_openai_client()
    cleaned = [t.replace("\n", " ").strip() or "skill" for t in texts]
    response = client.embeddings.create(input=cleaned, model=OPENAI_EMBED_MODEL)
    return [item.embedding for item in response.data]

# ══════════════════════════════════════════════════════════════════════════════
# LIGHTCAST HELPERS
# ══════════════════════════════════════════════════════════════════════════════

_lightcast_cache = None
_lightcast_cache_lock = Lock()

def get_lightcast_col():
    return get_mongo()[LIGHTCAST_DB_NAME][LIGHTCAST_COLLECTION]

def _get_lightcast_docs() -> list:
    """Load (and cache) all LightCast docs that have embeddings."""
    global _lightcast_cache
    with _lightcast_cache_lock:
        if _lightcast_cache is None:
            log.info("Loading LightCast docs from %s.%s…", LIGHTCAST_DB_NAME, LIGHTCAST_COLLECTION)
            try:
                col = get_lightcast_col()
                _lightcast_cache = list(col.find(
                    {"description_embedding": {"$exists": True, "$not": {"$size": 0}}},
                    {"_id": 0, "id": 1, "description": 1, "skill_domain": 1,
                    "mckinsey_category": 1, "mckinsey_subcategory": 1,
                    "mckinsey_sub_subcategory": 1, "description_embedding": 1}
                ))
                log.info("Loaded %d LightCast docs", len(_lightcast_cache))
            except Exception as e:
                log.error("Failed to load LightCast docs: %s", e)
                _lightcast_cache = []
        return _lightcast_cache

def _cosine_top_k(query_emb: list, docs: list, top_k: int = 15) -> list:
    """Return top_k LightCast docs ranked by cosine similarity to query_emb."""
    if not docs or not query_emb:
        return []
    q = np.array(query_emb, dtype=np.float32)
    q_norm = float(np.linalg.norm(q))
    if q_norm < 1e-9:
        return []

    scored = []
    for doc in docs:
        emb = doc.get("description_embedding", [])
        if not emb:
            continue
        v = np.array(emb, dtype=np.float32)
        v_norm = float(np.linalg.norm(v))
        sim = float(np.dot(q, v) / (q_norm * v_norm)) if v_norm > 1e-9 else 0.0
        scored.append({
            "lightcast_id": doc.get("id", ""),
            "talent_skill": doc.get("description", ""),
            "skill_domain": doc.get("skill_domain", ""),
            "mckinsey_category": doc.get("mckinsey_category", ""),
            "mckinsey_subcategory": doc.get("mckinsey_subcategory", ""),
            "mckinsey_sub_subcategory": doc.get("mckinsey_sub_subcategory", ""),
            "similarity": round(sim, 4),
        })
    scored.sort(key=lambda x: x["similarity"], reverse=True)
    return scored[:top_k]

# ══════════════════════════════════════════════════════════════════════════════
# MONGODB
# ══════════════════════════════════════════════════════════════════════════════

_mongo_client: MongoClient = None

def get_mongo() -> MongoClient:
    global _mongo_client
    if _mongo_client is None:
        _mongo_client = MongoClient(MONGO_URI, serverSelectionTimeoutMS=5000)
        log.info("MongoDB → %s / %s", MONGO_URI, MONGO_DB_NAME)
    return _mongo_client

def get_col(name: str = "curriculum"):
    return get_mongo()[MONGO_DB_NAME][name]

def ensure_indexes():
    col = get_col()
    col.create_index([("department", ASCENDING), ("semester", ASCENDING)])
    col.create_index([("skill", ASCENDING)])
    col.create_index([("lightcast_source", ASCENDING)])
    col.create_index([("mckinsey_category", ASCENDING)])
    col.create_index([("lightcast_skills.lightcast_id", ASCENDING)])
    col.create_index(
        [("department", ASCENDING), ("semester", ASCENDING),
         ("subject", ASCENDING), ("skill", ASCENDING)],
        unique=True, name="uq_skill_per_subject",
    )
    log.info("Indexes ensured on 'curriculum'")

def save_extracted_to_mongo(rows: list[dict], replace_filter: dict = None) -> dict:
    """Upsert engine1-extracted (and LightCast-mapped) skills into MongoDB."""
    col = get_col()
    if replace_filter:
        deleted = col.delete_many(replace_filter).deleted_count
        log.info("Cleared %d docs for %s before saving", deleted, replace_filter)

    valid_rows = []
    for row in rows:
        department = str(row.get("department", "")).strip()
        semester = str(row.get("semester", "")).strip()
        subject = str(row.get("subject", "")).strip()
        skill = str(row.get("skill", "")).strip()
        proficiency = str(row.get("proficiency", "")).strip()
        if all([department, semester, subject, skill, proficiency]):
            valid_rows.append(row)

    emb_map: dict[str, list] = {}
    if OPENAI_API_KEY and valid_rows:
        unique_skills = list({str(r.get("skill", "")).strip() for r in valid_rows})
        log.info("Generating skill embeddings for %d unique skills…", len(unique_skills))
        try:
            embeddings = generate_openai_embeddings(unique_skills)
            emb_map = dict(zip(unique_skills, embeddings))
        except Exception as e:
            log.warning("Embedding generation failed: %s — proceeding without embeddings", e)

    operations = []
    for row in valid_rows:
        department = str(row.get("department", "")).strip()
        year = str(row.get("year", "")).strip()
        semester = str(row.get("semester", "")).strip()
        subject = str(row.get("subject", "")).strip()
        skill = str(row.get("skill", "")).strip()
        proficiency = str(row.get("proficiency", "")).strip()
        credits = str(row.get("credits", "")).strip()
        total_hours = str(row.get("total_hours", "")).strip()
        bloom_level = str(row.get("bloom_level", "")).strip()
        bloom_explanation = str(row.get("bloom_explanation", "")).strip()
        rationale = str(row.get("proficiency_rationale", "")).strip()
        kws = row.get("raw_skill_keywords", [])
        raw_kws = [str(k).strip() for k in kws if str(k).strip()] if isinstance(kws, list) else []
        embedding = emb_map.get(skill, [])

        lc_skills_raw = row.get("lightcast_skills", [])
        lc_skills = []
        seen_lc_ids = set()
        for m in lc_skills_raw:
            if not isinstance(m, dict):
                continue
            lc_name = str(m.get("lightcast_skill", "")).strip()
            if not lc_name:
                continue
            lc_id = str(m.get("lightcast_id", "")).strip()
            dedup_key = lc_id if lc_id else lc_name.lower()
            if dedup_key in seen_lc_ids:
                continue
            seen_lc_ids.add(dedup_key)
            try:
                conf = float(m.get("match_confidence", 0.0))
            except (ValueError, TypeError):
                conf = 0.0
            lc_skills.append({
                "lightcast_id": lc_id,
                "lightcast_skill": lc_name,
                "match_confidence": conf,
            })

        lc_source = str(row.get("lightcast_source", "unmatched")).strip()
        mck_cat = str(row.get("mckinsey_category", "")).strip()
        mck_sub = str(row.get("mckinsey_subcategory", "")).strip()
        mck_subsub = str(row.get("mckinsey_sub_subcategory", "")).strip()

        doc_set = {
            "department": department, "year": year,
            "semester": semester, "subject": subject,
            "skill": skill, "proficiency": proficiency,
            "credits": credits, "total_hours": total_hours,
            "bloom_level": bloom_level, "bloom_explanation": bloom_explanation,
            "proficiency_rationale": rationale,
            "raw_skill_keywords": raw_kws,
            "mckinsey_category": mck_cat,
            "mckinsey_subcategory": mck_sub,
            "mckinsey_sub_subcategory": mck_subsub,
            "lightcast_skills": lc_skills,
            "lightcast_source": lc_source,
            "extracted_at": datetime.now(timezone.utc),
        }
        if embedding:
            doc_set["skill_embedding"] = embedding

        operations.append(UpdateOne(
            {"department": department, "semester": semester,
             "subject": subject, "skill": skill},
            {"$set": doc_set},
            upsert=True,
        ))

    if not operations:
        return {"inserted": 0, "modified": 0, "total_ops": 0}

    try:
        res = col.bulk_write(operations, ordered=False)
        lc_verified = sum(1 for r in valid_rows if r.get("lightcast_source") == "verified")
        log.info(
            "Saved %d skills (ins=%d mod=%d) — %d LightCast-verified, %d unmatched",
            len(operations), res.upserted_count, res.modified_count,
            lc_verified, len(valid_rows) - lc_verified,
        )
        return {
            "inserted": res.upserted_count,
            "modified": res.modified_count,
            "total_ops": len(operations),
            "lc_verified": lc_verified,
            "lc_unmatched": len(valid_rows) - lc_verified,
        }
    except BulkWriteError as bwe:
        return {"error": str(bwe.details)}


def _normalise_semester(sem: str) -> str:
    """Normalise varied semester strings to canonical form."""
    if not sem:
        return sem
    sem = sem.strip()
    m = re.search(r"([0-9]+)", sem)
    if m:
        n = m.group(1)
        lsem = sem.lower()
        if "sem" in lsem or "sem." in lsem or lsem.startswith("s"):
            return f"Semester {n}"
        if "year" in lsem or "yr" in lsem:
            return sem
    return sem


def get_curriculum(department: str = None, semester: str = None) -> list[dict]:
    """Fetch curriculum rows with case-insensitive matching. Excludes embedding fields from response."""
    import re as _re
    query: dict = {}
    if department:
        dept_escaped = _re.escape(department)
        query["department"] = {"$regex": f"^{dept_escaped}$", "$options": "i"}
    if semester:
        sem_escaped = _re.escape(semester)
        query["semester"] = {"$regex": f"^{sem_escaped}$", "$options": "i"}

    # Exclude _id and embedding fields from response
    projection = {"_id": 0, "skill_embedding": 0, "description_embedding": 0}

    rows = list(get_col().find(query, projection).sort([("subject", 1), ("skill", 1)]))
    if rows:
        return rows

    if semester:
        norm_sem = _normalise_semester(semester)
        if norm_sem != semester:
            log.info("Semester normalised: '%s' → '%s' — retrying DB lookup", semester, norm_sem)
            query2: dict = {}
            if department:
                dept_escaped = _re.escape(department)
                query2["department"] = {"$regex": f"^{dept_escaped}$", "$options": "i"}
            norm_escaped = _re.escape(norm_sem)
            query2["semester"] = {"$regex": f"^{norm_escaped}$", "$options": "i"}
            rows = list(get_col().find(query2, projection).sort([("subject", 1), ("skill", 1)]))

    return rows

def get_curriculum_stats() -> dict:
    pipeline = [
        {"$group": {
            "_id": {"department": "$department", "semester": "$semester"},
            "skill_count": {"$sum": 1},
            "subject_count": {"$addToSet": "$subject"},
            "lc_verified": {"$sum": {"$cond": [{"$eq": ["$lightcast_source", "verified"]}, 1, 0]}},
        }},
        {"$project": {
            "_id": 0,
            "department": "$_id.department",
            "semester": "$_id.semester",
            "skill_count": 1,
            "subject_count": {"$size": "$subject_count"},
            "lc_verified": 1,
        }},
        {"$sort": {"department": 1, "semester": 1}},
    ]
    return {
        "total_skills": get_col().count_documents({}),
        "lc_verified": get_col().count_documents({"lightcast_source": "verified"}),
        "lc_unmatched": get_col().count_documents({"lightcast_source": "unmatched"}),
        "breakdown": list(get_col().aggregate(pipeline)),
    }
