# app/pptx_builder.py
"""
Builds a PowerPoint (.pptx) from an outline and an uploaded template.

Key features:
- Template style transfer: applies theme fonts/colors when present.
- Smarter layout handling: respects requested layout; splits "Two Content" bullets into two columns.
- Sub-bullet rendering: strings starting with a secondary bullet marker render as level-1 bullets.
- Defensive: works even if placeholders/layouts are missing or template is unusual.
"""

from __future__ import annotations

from io import BytesIO
from typing import List, Optional, Tuple

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.dml.color import RGBColor

from .schemas import Outline
from .template_utils import (
    extract_template_images,
    find_preferred_layout,
    get_theme_style,  # NEW: theme colors/fonts
)

# ---------------- Tunables ----------------

TITLE_FALLBACK_SIZE_PT = 32
BODY_FALLBACK_SIZE_PT = 18

# Recognize sub-bullets (the parser uses a simple "  • " prefix for nesting).
SUB_BULLET_PREFIX = "  • "

# ---------------- Helpers ----------------

def _strip_control_chars(s: str) -> str:
    # Remove control chars except \t \n
    return "".join(ch for ch in (s or "") if ch == "\t" or ch == "\n" or ord(ch) >= 32)

def _rgb_from_hex(hex6: Optional[str]) -> Optional[RGBColor]:
    if not hex6 or len(hex6) < 6:
        return None
    try:
        r = int(hex6[0:2], 16)
        g = int(hex6[2:4], 16)
        b = int(hex6[4:6], 16)
        return RGBColor(r, g, b)
    except Exception:
        return None

def _apply_font_to_runs(text_frame, *, name: Optional[str], size_pt: Optional[int], color: Optional[RGBColor]):
    # Apply font to all runs across all paragraphs in a text_frame.
    for p in text_frame.paragraphs:
        for r in p.runs:
            if name:
                r.font.name = name
            if size_pt:
                r.font.size = Pt(size_pt)
            if color:
                r.font.color.rgb = color

def _title_placeholder(slide):
    # Prefer TITLE or CENTER_TITLE placeholders if present.
    for shp in slide.placeholders:
        try:
            ptype = shp.placeholder_format.type
            if ptype in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE):
                return shp
        except Exception:
            continue
    # Fallback: first placeholder if any
    return slide.placeholders[0] if slide.placeholders else None

def _subtitle_placeholder(slide):
    for shp in slide.placeholders:
        try:
            if shp.placeholder_format.type == PP_PLACEHOLDER.SUBTITLE:
                return shp
        except Exception:
            continue
    return None

def _content_placeholders(slide):
    """Return a list of text-capable content placeholders (BODY/CONTENT/SUBTITLE but not TITLE)."""
    out = []
    for shp in slide.placeholders:
        try:
            ptype = shp.placeholder_format.type
            if ptype in (PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.CONTENT, PP_PLACEHOLDER.SUBTITLE):
                # Ensure we can treat it as a text frame
                _ = shp.text_frame  # may raise if not text-capable
                out.append(shp)
        except Exception:
            continue
    return out

def _picture_placeholders(slide):
    holders = []
    for shp in slide.placeholders:
        try:
            if shp.placeholder_format.type in (PP_PLACEHOLDER.PICTURE, PP_PLACEHOLDER.CONTENT):
                holders.append(shp)
        except Exception:
            continue
    return holders

def _first_picture_placeholder(slide):
    pics = _picture_placeholders(slide)
    return pics[0] if pics else None

def _bullet_level_and_text(text: str) -> Tuple[int, str]:
    """
    Detect whether this bullet should be level-0 or level-1.
    The parser formats sub-bullets with a '  • ' prefix; detect this and set level=1.
    """
    s = _strip_control_chars(text or "").rstrip()
    if not s:
        return 0, ""
    if s.startswith(SUB_BULLET_PREFIX):
        return 1, s[len(SUB_BULLET_PREFIX):].strip()
    # Be tolerant of a literal '• ' at the start too.
    if s.strip().startswith("•") and not s.startswith("•"):  # e.g., "   • text"
        # Remove first '•'
        cleaned = s[s.index("•") + 1 :].strip()
        return 1, cleaned
    if s.startswith("• "):
        return 1, s[2:].strip()
    return 0, s

# ---------------- Writers ----------------

def _set_title(slide, title_text: str, theme: Optional[dict]):
    ph = _title_placeholder(slide)
    if ph is None:
        return
    ph.text = _strip_control_chars(title_text or "")

    # Title styling: prefer major font + accent1 color; fallback to dk1
    font_name = (theme or {}).get("fonts", {}).get("major") or (theme or {}).get("fonts", {}).get("minor")
    color_hex = (theme or {}).get("colors", {}).get("accent1") or (theme or {}).get("colors", {}).get("dk1")
    color = _rgb_from_hex(color_hex)
    tf = ph.text_frame
    tf.word_wrap = True
    _apply_font_to_runs(tf, name=font_name, size_pt=TITLE_FALLBACK_SIZE_PT, color=color)

def _set_subtitle_if_present(slide, subtitle_text: Optional[str], theme: Optional[dict]):
    if not subtitle_text:
        return
    ph = _subtitle_placeholder(slide)
    if ph is None:
        return
    ph.text = _strip_control_chars(subtitle_text)
    font_name = (theme or {}).get("fonts", {}).get("minor") or (theme or {}).get("fonts", {}).get("major")
    color_hex = (theme or {}).get("colors", {}).get("dk1") or (theme or {}).get("colors", {}).get("lt1")
    color = _rgb_from_hex(color_hex)
    tf = ph.text_frame
    tf.word_wrap = True
    _apply_font_to_runs(tf, name=font_name, size_pt=BODY_FALLBACK_SIZE_PT, color=color)

def _set_bullets_single(tf, bullets: List[str], theme: Optional[dict]):
    """Write bullets into a single text frame, respecting sub-bullet levels."""
    tf.clear()
    tf.word_wrap = True
    if not bullets:
        return

    # First bullet
    lvl, txt = _bullet_level_and_text(bullets[0])
    p0 = tf.paragraphs[0]
    p0.level = max(0, min(4, lvl))
    p0.text = txt

    # Others
    for b in bullets[1:]:
        lvl, txt = _bullet_level_and_text(b)
        if not txt:
            continue
        para = tf.add_paragraph()
        para.level = max(0, min(4, lvl))
        para.text = txt

    # Apply theme
    body_font = (theme or {}).get("fonts", {}).get("minor") or (theme or {}).get("fonts", {}).get("major")
    body_color_hex = (theme or {}).get("colors", {}).get("dk1")
    body_color = _rgb_from_hex(body_color_hex)
    _apply_font_to_runs(tf, name=body_font, size_pt=BODY_FALLBACK_SIZE_PT, color=body_color)

def _set_bullets(slide, bullets: List[str], theme: Optional[dict]):
    """
    Fill bullets into available content placeholders.
    - If two or more content placeholders exist and there are many bullets, split across first two.
    - Otherwise, write into the first content placeholder.
    """
    placeholders = _content_placeholders(slide)
    if not placeholders:
        return

    # Decide split for "Two Content" layouts when two text placeholders are present.
    if len(placeholders) >= 2 and len(bullets) >= 6:
        half = (len(bullets) + 1) // 2
        left, right = bullets[:half], bullets[half:]
        _set_bullets_single(placeholders[0].text_frame, left, theme)
        _set_bullets_single(placeholders[1].text_frame, right, theme)
        return

    # Single placeholder
    _set_bullets_single(placeholders[0].text_frame, bullets, theme)

def _insert_picture(slide, image_bytes: bytes) -> bool:
    """
    Insert picture using a picture-capable placeholder if available.
    Fallback: anchor on the right side within slide bounds.
    """
    ph = _first_picture_placeholder(slide)
    if ph is not None:
        try:
            ph.insert_picture(BytesIO(image_bytes))
            return True
        except Exception:
            pass
    # Fallback: place a right-aligned image at a reasonable size
    try:
        # Size ~ one third of width; maintain aspect ratio auto-handled by python-pptx when only width provided.
        slide_width = slide.part.slide_width
        slide_height = slide.part.slide_height
        width_in = max(2.5, (slide_width / 914400) * 0.33)  # EMU to inches (914400 per inch)
        left_in = max(0.0, (slide_width / 914400) - width_in - 0.5)
        top_in = 1.0
        slide.shapes.add_picture(BytesIO(image_bytes), Inches(left_in), Inches(top_in), width=Inches(width_in))
        return True
    except Exception:
        return False

# ---------------- Public Builder ----------------

def build_presentation(outline: Outline, template_bytes: bytes, *, subtitle: Optional[str] = None) -> bytes:
    """
    Build a .pptx as bytes from the given outline and template (.pptx/.potx bytes).
    - Adds a title slide when possible.
    - Applies theme fonts/colors.
    - Writes content slides per requested layout with graceful fallbacks.
    - Optionally sets a subtitle on title slide (unused by API by default).
    """
    prs = Presentation(BytesIO(template_bytes))
    theme = get_theme_style(template_bytes) or {"colors": {}, "fonts": {}}  # NEW
    template_images = extract_template_images(template_bytes)
    img_idx = 0

    # ----- Title slide -----
    title_layout = find_preferred_layout(prs, ["Title Slide", "Title Only", "Section Header", "Title"])
    if title_layout is not None:
        title_slide = prs.slides.add_slide(title_layout)
        _set_title(title_slide, outline.title, theme)
        _set_subtitle_if_present(title_slide, subtitle, theme)
    else:
        # Fallback: first layout
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        _set_title(slide, outline.title, theme)
        _set_subtitle_if_present(slide, subtitle, theme)

    # ----- Content slides -----
    for s in outline.slides:
        requested = (s.layout or "auto").strip().lower()
        # Try to honor requested layout; otherwise reasonable fallback ordering.
        if requested == "auto":
            layout = find_preferred_layout(
                prs,
                ["Title and Content", "Two Content", "Content with Caption", "Picture with Caption", "Blank"],
            )
        else:
            layout = find_preferred_layout(
                prs,
                [s.layout, "Title and Content", "Two Content", "Content with Caption", "Picture with Caption", "Blank"],
            )
        if layout is None:
            layout = prs.slide_layouts[1] if len(prs.slide_layouts) > 1 else prs.slide_layouts[0]

        slide = prs.slides.add_slide(layout)

        # Title + bullets with theme
        _set_title(slide, s.title, theme)
        _set_bullets(slide, list(s.bullets or []), theme)

        # Reuse template images politely:
        # Only insert if a picture placeholder exists OR the chosen layout implies a picture.
        layout_name = getattr(layout, "name", "") or ""
        wants_picture = "picture" in layout_name.lower() or _first_picture_placeholder(slide) is not None
        if template_images and wants_picture:
            inserted = _insert_picture(slide, template_images[img_idx % len(template_images)])
            if inserted:
                img_idx += 1

        # Speaker notes (optional)
        if getattr(s, "notes", None) is not None:
            try:
                notes_slide = slide.notes_slide
                notes_tf = notes_slide.notes_text_frame
                notes_tf.clear()
                notes_tf.text = _strip_control_chars(s.notes or "")
            except Exception:
                # Some very old templates may not support notes properly; ignore silently.
                pass

    # ----- Save -----
    bio = BytesIO()
    prs.save(bio)
    return bio.getvalue()
