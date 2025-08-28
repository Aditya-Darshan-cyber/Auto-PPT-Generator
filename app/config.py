# app/config.py
"""
Global configuration knobs (env-overridable) for Auto PPT Generator.

Goals:
- Out-of-the-box works with AI Pipe on Hugging Face Spaces.
- Single source of truth for limits (uploads, text size, bullets, slides).
- Centralize LLM network settings (timeouts, retries).
- Safe defaults; production-friendly via environment overrides.
"""

from __future__ import annotations

import os
from typing import Set, List

# ---------------- helpers ----------------

def _env_int(name: str, default: int) -> int:
    try:
        return int(os.getenv(name, str(default)))
    except Exception:
        return default

def _env_float(name: str, default: float) -> float:
    try:
        return float(os.getenv(name, str(default)))
    except Exception:
        return default

def _env_bool(name: str, default: bool) -> bool:
    val = os.getenv(name)
    if val is None:
        return default
    return str(val).strip().lower() in {"1", "true", "yes", "on"}

def _env_csv_set(name: str, default_items: List[str]) -> Set[str]:
    raw = os.getenv(name)
    if not raw:
        return {x.strip().lower() for x in default_items}
    parts = [p.strip().lower() for p in raw.split(",")]
    return {p for p in parts if p}

# ---------------- files & uploads ----------------

# Max upload size (MB) for template files
MAX_FILE_MB: int = _env_int("MAX_FILE_MB", 20)

# Which file extensions we accept for PowerPoint templates
ALLOWED_EXTS: Set[str] = _env_csv_set("ALLOWED_EXTS", [".pptx", ".potx"])

# Zip safety (prevents zip-bombs and bogus PPTX)
MAX_ZIP_ENTRIES: int = _env_int("MAX_ZIP_ENTRIES", 2000)
MAX_ZIP_MEMBER_MB: int = _env_int("MAX_ZIP_MEMBER_MB", 50)

# Template image extraction limits
MAX_TEMPLATE_IMAGES: int = _env_int("MAX_TEMPLATE_IMAGES", 20)
MAX_TEMPLATE_IMAGE_MB: int = _env_int("MAX_TEMPLATE_IMAGE_MB", 5)

# ---------------- LLM & routing ----------------

# Default provider & model
DEFAULT_PROVIDER: str = os.getenv("DEFAULT_PROVIDER", "openai")
DEFAULT_MODEL: str = os.getenv("OPENAI_MODEL", "gpt-4.1-mini")

# OpenAI-compatible base URL (AI Pipe by default)
OPENAI_BASE_URL: str = os.getenv("OPENAI_BASE_URL", "https://aipipe.org/openai/v1")

# Network behavior for LLM calls
LLM_TIMEOUT_SECS: float = _env_float("LLM_TIMEOUT_SECS", 60.0)
LLM_MAX_RETRIES: int = _env_int("LLM_MAX_RETRIES", 3)

# ---------------- Text & outline constraints ----------------

# Max input text size to accept/process (characters)
MAX_TEXT_CHARS: int = _env_int("MAX_TEXT_CHARS", 40_000)

# Slide/outline content limits (keep in sync across parser/schemas/builder)
MAX_BULLETS_PER_SLIDE: int = _env_int("MAX_BULLETS_PER_SLIDE", 7)
MAX_TITLE_CHARS: int = _env_int("MAX_TITLE_CHARS", 200)
MAX_BULLET_CHARS: int = _env_int("MAX_BULLET_CHARS", 200)
MAX_NOTES_CHARS: int = _env_int("MAX_NOTES_CHARS", 600)
MAX_TOTAL_SLIDES: int = _env_int("MAX_TOTAL_SLIDES", 60)

# ---------------- CORS ----------------

# Comma-separated list of allowed origins for the API (use "*" for demos)
CORS_ALLOW_ORIGINS: Set[str] = _env_csv_set("CORS_ALLOW_ORIGINS", ["*"])

# Whether to allow credentials (cookies). For this app, defaults to False.
CORS_ALLOW_CREDENTIALS: bool = _env_bool("CORS_ALLOW_CREDENTIALS", False)

# Methods/headers are wide-open for demos; override in production if needed
CORS_ALLOW_METHODS: Set[str] = _env_csv_set("CORS_ALLOW_METHODS", ["*"])
CORS_ALLOW_HEADERS: Set[str] = _env_csv_set("CORS_ALLOW_HEADERS", ["*"])
