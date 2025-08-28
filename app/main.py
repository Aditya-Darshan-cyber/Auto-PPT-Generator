# app/main.py
"""
FastAPI back end for 'Auto PPT Generator':
- Accepts text/guidance + user-supplied LLM token + PPTX template, returns a generated .pptx
- Safe-by-default: never stores or logs API keys, validates PPTX structure, clamps input sizes
- HF Spaces + AI Pipe friendly (OpenAI-compatible base URL), CORS enabled for split hosting
"""

from __future__ import annotations

import io
import os
from datetime import datetime
from typing import Optional

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import HTMLResponse, StreamingResponse, Response, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware

from .pptx_builder import build_presentation
from .parser import heuristic_outline
from .llm_clients import plan_slides_via_llm
from .schemas import Outline
from .config import MAX_FILE_MB, ALLOWED_EXTS, DEFAULT_MODEL, DEFAULT_PROVIDER
from .template_utils import is_safe_pptx  # NEW: zip safety + content type checks

# ---------------- Config ----------------

# Limit raw text length to protect both LLM and fallback parser.
MAX_TEXT_CHARS = int(os.getenv("MAX_TEXT_CHARS", "40000"))

# Default to OpenAI-compatible provider (AI Pipe) unless caller overrides.
DEFAULT_BASE_URL = os.getenv("OPENAI_BASE_URL", "https://aipipe.org/openai/v1")

app = FastAPI(title="Auto_PPT_Generator", version="1.1.0", docs_url="/docs")

# If you host the UI separately, this enables cross-origin calls (demo-friendly).
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # For demos. In production, restrict this.
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Serve static front-end (if present)
static_path = os.path.join(os.path.dirname(__file__), "..", "web")
if os.path.isdir(static_path):
    app.mount("/assets", StaticFiles(directory=static_path), name="assets")


# ---------------- Helpers ----------------

def _bool_from_form(val: Optional[str | bool]) -> bool:
    """Robust bool coercion from HTML form values."""
    if isinstance(val, bool):
        return val
    if val is None:
        return False
    s = str(val).strip().lower()
    return s in {"1", "true", "yes", "on"}

def _clamp_text(s: str) -> str:
    """Clamp text to MAX_TEXT_CHARS; avoids oversized payloads."""
    if not s:
        return ""
    if len(s) <= MAX_TEXT_CHARS:
        return s
    return s[:MAX_TEXT_CHARS]

def _safe_filename(base: str) -> str:
    """Remove risky characters from filenames."""
    keep = "-_.()abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
    return "".join(ch if ch in keep else "_" for ch in base)


# ---------------- Routes ----------------

@app.get("/", response_class=HTMLResponse)
def index() -> HTMLResponse:
    """Serve the client app if packaged; otherwise a minimal placeholder."""
    index_html_path = os.path.join(static_path, "index.html")
    if os.path.isfile(index_html_path):
        with open(index_html_path, "r", encoding="utf-8") as f:
            return HTMLResponse(f.read())
    return HTMLResponse("<h1>Auto PPT Generator API</h1><p>POST /api/generate or /api/preview_outline</p>")

@app.head("/")
def head_root() -> Response:
    return Response(status_code=204)

@app.get("/healthz")
def healthz():
    return {"ok": True, "ts": datetime.utcnow().isoformat() + "Z"}


@app.post("/api/preview_outline")
async def preview_outline(
    text: str = Form(..., description="Raw text or markdown input"),
    guidance: Optional[str] = Form(None, description="Optional one-line guidance"),
    provider: str = Form(DEFAULT_PROVIDER, description="LLM provider (OpenAI-compatible recommended)"),
    model: Optional[str] = Form(DEFAULT_MODEL, description="Model name"),
    api_key: Optional[str] = Form(None, description="User-supplied LLM token (never stored)"),
    base_url: Optional[str] = Form(None, description="OpenAI-compatible base URL (defaults to AI Pipe)"),
    include_notes: Optional[str] = Form("false", description="true/false"),
):
    """
    Return just the outline JSON (for preview). Uses LLM if token provided; otherwise heuristic fallback.
    """
    # Clamp text
    text = _clamp_text(text or "")
    use_notes = _bool_from_form(include_notes)

    # Try LLM first if token supplied; otherwise fallback.
    try:
        if api_key:
            outline_dict = plan_slides_via_llm(
                text=text,
                guidance=guidance or "",
                provider=provider or DEFAULT_PROVIDER,
                api_key=api_key,
                model=model or DEFAULT_MODEL,
                base_url=(base_url or DEFAULT_BASE_URL),
                include_notes=use_notes,
            )
        else:
            outline_dict = heuristic_outline(text=text, guidance=guidance or "", include_notes=use_notes)
    except Exception:
        # Silent fallback: never expose token or provider error details
        outline_dict = heuristic_outline(text=text, guidance=guidance or "", include_notes=use_notes)

    # Pydantic validation (raises if malformed; handled by FastAPI)
    outline = Outline(**outline_dict)
    return JSONResponse(outline.model_dump())


@app.post("/api/generate")
async def generate_pptx(
    text: str = Form(..., description="Raw text or markdown input"),
    guidance: Optional[str] = Form(None, description="Optional one-line guidance"),
    provider: str = Form(DEFAULT_PROVIDER, description="LLM provider (OpenAI-compatible recommended)"),
    model: Optional[str] = Form(DEFAULT_MODEL, description="Model name"),
    api_key: Optional[str] = Form(None, description="User-supplied LLM token (never stored)"),
    base_url: Optional[str] = Form(None, description="OpenAI-compatible base URL (defaults to AI Pipe)"),
    include_notes: Optional[str] = Form("false", description="true/false"),
    template: UploadFile = File(..., description="PowerPoint template or presentation (.pptx or .potx)"),
):
    """
    Build and return a .pptx based on the provided text/guidance/template.
    - Uses LLM if token present; otherwise heuristic parser.
    - Applies the uploaded template's layouts/colors/fonts (via builder).
    """
    # ---------- Validate & read template ----------
    name = template.filename or "template.pptx"
    ext = os.path.splitext(name.lower())[1]
    if ext not in ALLOWED_EXTS:
        raise HTTPException(status_code=400, detail=f"Unsupported file type: {ext}. Allowed: {', '.join(ALLOWED_EXTS)}")

    contents = await template.read()
    size_mb = len(contents) / (1024 * 1024)
    if size_mb > MAX_FILE_MB:
        raise HTTPException(status_code=413, detail=f"Template too large ({size_mb:.1f} MB). Max is {MAX_FILE_MB} MB.")

    # Structural / safety checks: ensure it's a legitimate PPTX/POTX and not a zip bomb
    try:
        if not is_safe_pptx(contents):
            raise HTTPException(status_code=400, detail="Invalid or unsafe PowerPoint file.")
    except HTTPException:
        raise
    except Exception:
        # If validator itself failed (rare), be conservative
        raise HTTPException(status_code=400, detail="Could not validate the uploaded template file.")

    # ---------- Build outline (LLM → fallback) ----------
    text = _clamp_text(text or "")
    use_notes = _bool_from_form(include_notes)

    try:
        if api_key:
            outline_dict = plan_slides_via_llm(
                text=text,
                guidance=guidance or "",
                provider=provider or DEFAULT_PROVIDER,
                api_key=api_key,
                model=model or DEFAULT_MODEL,
                base_url=(base_url or DEFAULT_BASE_URL),
                include_notes=use_notes,
            )
        else:
            outline_dict = heuristic_outline(text=text, guidance=guidance or "", include_notes=use_notes)
    except Exception:
        # Defensive fallback if provider errors, never leak key
        outline_dict = heuristic_outline(text=text, guidance=guidance or "", include_notes=use_notes)

    # Validate structure (pydantic)
    outline = Outline(**outline_dict)

    # ---------- Generate PPTX ----------
    try:
        pptx_bytes = build_presentation(outline=outline, template_bytes=contents, subtitle=(guidance or None))
    except Exception as e:
        # Hide internals but give a short message
        raise HTTPException(status_code=500, detail="Failed to build PowerPoint from the provided template.")

    # ---------- Stream back as a file download ----------
    ts = datetime.utcnow().strftime("%Y%m%d-%H%M%S")
    base = _safe_filename(f"Auto_PPT_Generator-{ts}")
    filename = f"{base}.pptx"
    headers = {"Content-Disposition": f'attachment; filename="{filename}"'}

    return StreamingResponse(
        io.BytesIO(pptx_bytes),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers=headers,
    )
