# app/schemas.py
"""
Pydantic schemas with strong validation and coercion for:
- OutlineSlide: title, bullets, layout, optional notes
- Outline: deck title, slides, estimated_slide_count

Key features:
- Canonicalize/validate layout names against an allowlist
- Sanitize and trim title/bullets/notes
- Deduplicate bullets, cap bullets per slide
- Auto-fix estimated_slide_count
- Forbid unknown fields to keep schema tight
"""

from __future__ import annotations

import os
import re
from typing import List, Optional, Literal

from pydantic import BaseModel, Field, ConfigDict, field_validator, model_validator

# ---------------- Config / constants ----------------

# Keep this aligned with builder/parser defaults
MAX_BULLETS_PER_SLIDE = int(os.getenv("MAX_BULLETS_PER_SLIDE", "7"))
MAX_TITLE_CHARS = int(os.getenv("MAX_TITLE_CHARS", "200"))
MAX_BULLET_CHARS = int(os.getenv("MAX_BULLET_CHARS", "200"))
MAX_NOTES_CHARS = int(os.getenv("MAX_NOTES_CHARS", "600"))
MAX_TOTAL_SLIDES = int(os.getenv("MAX_TOTAL_SLIDES", "60"))

# Canonical layout names (must match what the builder knows)
_CANONICAL_LAYOUTS = [
    "auto",
    "Title and Content",
    "Two Content",
    "Content with Caption",
    "Picture with Caption",
    "Blank",
]
LayoutName = Literal[
    "auto",
    "Title and Content",
    "Two Content",
    "Content with Caption",
    "Picture with Caption",
    "Blank",
]

# Map common variants to canonical names (case-insensitive)
_LAYOUT_ALIASES = {
    "auto": "auto",
    "title and content": "Title and Content",
    "two content": "Two Content",
    "two-content": "Two Content",
    "twocontents": "Two Content",
    "content with caption": "Content with Caption",
    "picture with caption": "Picture with Caption",
    "blank": "Blank",
}

# ---------------- helpers ----------------

_WS_RE = re.compile(r"\s+")
_CTRL_KEEP = {"\n", "\t"}

def _strip_controls(s: str) -> str:
    return "".join(ch for ch in (s or "") if (ch in _CTRL_KEEP) or (ord(ch) >= 32))

def _collapse_ws(s: str) -> str:
    return _WS_RE.sub(" ", (s or "").strip())

def _clean_text(s: str, limit: int) -> str:
    s = _strip_controls(s)
    s = _collapse_ws(s)
    return s if len(s) <= limit else s[: max(0, limit - 1)].rstrip() + "…"

def _dedup_keep_order(items: List[str]) -> List[str]:
    seen = set()
    out: List[str] = []
    for it in items:
        if it not in seen:
            seen.add(it)
            out.append(it)
    return out

def _coerce_bullets(value) -> List[str]:
    # Accept list of strings primarily; if given a string, split by newlines.
    if value is None:
        return []
    if isinstance(value, str):
        raw = [ln.strip() for ln in value.splitlines()]
    elif isinstance(value, list):
        raw = []
        for x in value:
            raw.extend(str(x).splitlines())
        raw = [ln.strip() for ln in raw]
    else:
        # Unexpected type: coerce to string
        raw = [str(value).strip()]

    # Clean, drop empties, trim and dedupe
    cleaned = [_clean_text(b, MAX_BULLET_CHARS) for b in raw if b and b.strip()]
    cleaned = _dedup_keep_order([b for b in cleaned if b])
    # Cap bullet count
    return cleaned[:MAX_BULLETS_PER_SLIDE]

def _canonical_layout(name: str | None) -> LayoutName:
    if not name:
        return "auto"
    key = str(name).strip().lower()
    return _LAYOUT_ALIASES.get(key, "auto")  # default to auto if unknown

# ---------------- Schemas ----------------

class OutlineSlide(BaseModel):
    model_config = ConfigDict(extra="forbid")

    title: str = Field(..., description="Slide title (will be trimmed)", min_length=1)
    bullets: List[str] = Field(default_factory=list, description="List of bullet strings")
    layout: LayoutName = Field(default="auto", description="Preferred layout hint")
    notes: Optional[str] = Field(default=None, description="Optional speaker notes")

    # --- validators ---

    @field_validator("title", mode="before")
    @classmethod
    def _v_title(cls, v: str) -> str:
        v = _clean_text(str(v or "").strip(), MAX_TITLE_CHARS)
        return v or "Slide"

    @field_validator("bullets", mode="before")
    @classmethod
    def _v_bullets_before(cls, v):
        return _coerce_bullets(v)

    @field_validator("bullets")
    @classmethod
    def _v_bullets_after(cls, v: List[str]) -> List[str]:
        # Final trim + filter empties + cap
        out = [_clean_text(b, MAX_BULLET_CHARS) for b in v if b and b.strip()]
        out = _dedup_keep_order(out)
        return out[:MAX_BULLETS_PER_SLIDE]

    @field_validator("layout", mode="before")
    @classmethod
    def _v_layout(cls, v) -> LayoutName:
        return _canonical_layout(v)

    @field_validator("notes", mode="before")
    @classmethod
    def _v_notes(cls, v: Optional[str]) -> Optional[str]:
        if v is None:
            return None
        cleaned = _clean_text(str(v), MAX_NOTES_CHARS)
        return cleaned if cleaned else None


class Outline(BaseModel):
    model_config = ConfigDict(extra="forbid")

    title: str = Field(..., description="Deck title", min_length=1, max_length=MAX_TITLE_CHARS)
    slides: List[OutlineSlide] = Field(..., description="Slides in order")
    estimated_slide_count: Optional[int] = Field(
        default=None,
        description="Approximate slide count; recomputed if missing or inconsistent",
    )

    # --- validators ---

    @field_validator("title", mode="before")
    @classmethod
    def _v_title(cls, v: str) -> str:
        v = _clean_text(str(v or "").strip(), MAX_TITLE_CHARS)
        return v or "Presentation"

    @field_validator("slides", mode="before")
    @classmethod
    def _v_slides_before(cls, v):
        # Accept a single slide object, coerce to list
        if isinstance(v, dict):
            return [v]
        return v

    @field_validator("slides")
    @classmethod
    def _v_slides_after(cls, slides: List[OutlineSlide]) -> List[OutlineSlide]:
        # Drop fully empty slides (no title and no bullets) — rare after slide-level validators
        filtered: List[OutlineSlide] = []
        for s in slides:
            if (s.title and s.title.strip()) or (s.bullets and any(b.strip() for b in s.bullets)):
                filtered.append(s)
        if not filtered:
            # Ensure at least one minimal slide
            filtered = [OutlineSlide(title="Overview", bullets=[], layout="auto")]
        # Cap total slides to protect builder
        return filtered[:MAX_TOTAL_SLIDES]

    @model_validator(mode="after")
    def _m_estimated_count(self) -> "Outline":
        # Compute or fix estimated_slide_count
        true_count = len(self.slides or [])
        if self.estimated_slide_count is None:
            self.estimated_slide_count = true_count
        else:
            try:
                n = int(self.estimated_slide_count)
            except Exception:
                n = true_count
            n = max(1, min(MAX_TOTAL_SLIDES, n))
            self.estimated_slide_count = n
        return self
