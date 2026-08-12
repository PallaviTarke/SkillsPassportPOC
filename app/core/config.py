"""
Core configuration module for the Skill Passport application.
"""

from __future__ import annotations

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

import logging
import os
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor
from threading import Lock
import queue

from google import genai
from google.genai import types as genai_types

# ── Logging ────────────────────────────────────────────────────────────────────
logging.basicConfig(level=logging.INFO, format="%(levelname)s | %(message)s")
log = logging.getLogger(__name__)

# ── Paths ──────────────────────────────────────────────────────────────────────
BASE_DIR = Path(__file__).parent.parent.parent
KEY_PATH = BASE_DIR / "service-account.json"

# ── Environment Variables ──────────────────────────────────────────────────────
MONGO_URI = os.getenv("MONGO_URI", "mongodb://localhost:27017")
MONGO_DB_NAME = os.getenv("MONGO_DB_NAME", "skill_passport")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")
DOCAI_PROJECT_ID = os.getenv("DOCAI_PROJECT_ID", "")
DOCAI_LOCATION = os.getenv("DOCAI_LOCATION", "us")
DOCAI_PROCESSOR_ID = os.getenv("DOCAI_PROCESSOR_ID", "")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")
LIGHTCAST_DB_NAME = os.getenv("LIGHTCAST_DB_NAME", "Aadhar_VC")
LIGHTCAST_COLLECTION = os.getenv("LIGHTCAST_COLLECTION", "LightCast")

# ── Model Configuration ────────────────────────────────────────────────────────
GEMINI_MODEL = "gemini-2.0-flash"
OPENAI_EMBED_MODEL = "text-embedding-ada-002"

# ── Application Configuration ──────────────────────────────────────────────────
APP_NAME = "Skill Passport Generator"
APP_VERSION = "1.0.0"
MAX_UPLOAD_SIZE = 200 * 1024 * 1024  # 200 MB

# ── Gemini Client ──────────────────────────────────────────────────────────────
_genai_client = None

def get_genai_client():
    global _genai_client
    if _genai_client is None:
        if not GEMINI_API_KEY:
            raise RuntimeError("GEMINI_API_KEY is not set.")
        _genai_client = genai.Client(api_key=GEMINI_API_KEY)
    return _genai_client

# ── OpenAI Client ──────────────────────────────────────────────────────────────
_openai_client = None

def get_openai_client():
    global _openai_client
    if _openai_client is None:
        if not OPENAI_API_KEY:
            raise RuntimeError("OPENAI_API_KEY is not set.")
        from openai import OpenAI
        _openai_client = OpenAI(api_key=OPENAI_API_KEY)
    return _openai_client

# ── Thread Pool ────────────────────────────────────────────────────────────────
_executor = ThreadPoolExecutor(max_workers=8)

# ── SSE Job Store ──────────────────────────────────────────────────────────────
class JobStore:
    def __init__(self):
        self._lock = Lock()
        self._jobs = {}

    def create(self, job_id):
        q = queue.Queue()
        with self._lock:
            self._jobs[job_id] = q
        return q

    def get(self, job_id):
        with self._lock:
            return self._jobs.get(job_id)

    def delete(self, job_id):
        with self._lock:
            self._jobs.pop(job_id, None)

_job_store = JobStore()


class _TTLStore:
    """Thread-safe key→bytes store with TTL eviction."""
    def __init__(self, ttl_seconds: int = 600, max_entries: int = 200):
        self._lock = Lock()
        self._store = {}
        self._ttl = ttl_seconds
        self._max = max_entries

    def put(self, key: str, data: bytes) -> None:
        import time
        expiry = time.time() + self._ttl
        with self._lock:
            now = time.time()
            self._store = {k: v for k, v in self._store.items() if v[1] > now}
            if len(self._store) >= self._max:
                oldest = sorted(self._store.items(), key=lambda x: x[1][1])[:len(self._store) - self._max + 1]
                for k, _ in oldest:
                    self._store.pop(k, None)
            self._store[key] = (data, expiry)

    def pop(self, key: str) -> bytes | None:
        import time
        with self._lock:
            entry = self._store.pop(key, None)
            if entry and entry[1] > time.time():
                return entry[0]
            return None

_result_store = _TTLStore(ttl_seconds=600)
_passport_store = _TTLStore(ttl_seconds=600)
_result_lock = Lock()
_passport_lock = Lock()

# ── Utility Functions ──────────────────────────────────────────────────────────
def sse_msg(event, data):
    import json
    return f"event: {event}\ndata: {json.dumps(data)}\n\n"


def _gemini_with_retry(fn, max_attempts: int = 4):
    """Call fn() with exponential back-off on transient Gemini errors."""
    import time
    for attempt in range(max_attempts):
        try:
            return fn()
        except Exception as exc:
            msg = str(exc).lower()
            retryable = any(x in msg for x in ("503", "429", "rate", "overloaded", "timeout", "unavailable"))
            if retryable and attempt < max_attempts - 1:
                wait = 2 ** attempt
                log.warning("Gemini transient error (attempt %d/%d): %s — retrying in %ds",
                            attempt + 1, max_attempts, exc, wait)
                time.sleep(wait)
            else:
                raise


def _call_gemini_json(prompt: str, system: str = None, max_tokens: int = 4096) -> dict | list:
    """Minimal Gemini call that returns parsed JSON, with retry on transient errors."""
    import json
    import re

    cfg = genai_types.GenerateContentConfig(
        temperature=0.05,
        response_mime_type="application/json",
        max_output_tokens=max_tokens,
    )
    if system:
        cfg.system_instruction = system

    def _call():
        resp = get_genai_client().models.generate_content(
            model=GEMINI_MODEL, contents=prompt, config=cfg,
        )
        raw = resp.text.strip()
        raw = re.sub(r'^```(?:json)?\s*', '', raw, flags=re.MULTILINE)
        raw = re.sub(r'\s*```\s*$', '', raw, flags=re.MULTILINE)
        return json.loads(raw.strip())

    return _gemini_with_retry(_call)
