# app/template_utils.py
"""
Utilities for working with uploaded PowerPoint templates:
- Safe image extraction from ppt/media (raster formats only)
- Theme parsing (colors + major/minor fonts) from ppt/theme/theme*.xml
- Robust layout selection by name and placeholder capabilities
- PPTX safety checks to avoid zip-bombs / corrupt files
- Slide dimension helper (EMU, inches, cm)
"""

from __future__ import annotations

import hashlib
import os
import zipfile
from io import BytesIO
from typing import Dict, List, Optional, Tuple
import xml.etree.ElementTree as ET

from pptx import Presentation
from pptx.slide import SlideLayout
from pptx.enum.shapes import PP_PLACEHOLDER

# ---------------- Limits / Env ----------------

MAX_TEMPLATE_IMAGES = int(os.getenv("MAX_TEMPLATE_IMAGES", "20"))
MAX_TEMPLATE_IMAGE_MB = int(os.getenv("MAX_TEMPLATE_IMAGE_MB", "5"))
MAX_ZIP_ENTRIES = int(os.getenv("MAX_ZIP_ENTRIES", "2000"))
MAX_ZIP_MEMBER_MB = int(os.getenv("MAX_ZIP_MEMBER_MB", "50"))

RASTER_EXTS = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff"}

# Unit constants
EMU_PER_INCH = 914400
CM_PER_INCH = 2.54

# ---------------- Safety ----------------

def is_safe_pptx(template_bytes: bytes, max_entries: int = MAX_ZIP_ENTRIES, max_member_mb: int = MAX_ZIP_MEMBER_MB) -> bool:
    """
    Basic PPTX safety checks:
      - Valid ZIP
      - Contains [Content_Types].xml (typical for OOXML packages)
      - Reasonable number of entries
      - Reasonable per-member uncompressed size
    """
    from zipfile import ZipFile, BadZipFile

    try:
        with ZipFile(BytesIO(template_bytes)) as z:
            names = z.namelist()
            if "[Content_Types].xml" not in names:
                return False
            if len(names) > max_entries:
                return False
            limit_bytes = max_member_mb * 1024 * 1024
            for info in z.infolist():
                # file_size is uncompressed size
                if info.file_size > limit_bytes:
                    return False
        return True
    except BadZipFile:
        return False
    except Exception:
        return False

# ---------------- Images ----------------

def extract_template_images(template_bytes: bytes) -> List[bytes]:
    """
    Return raw bytes for raster images in ppt/media/*.
    - Filters to raster formats the builder can reliably insert.
    - Deduplicates by SHA-256 to avoid repeats.
    - Respects MAX_TEMPLATE_IMAGES and per-image size limit.
    """
    images: List[bytes] = []
    seen_hashes = set()
    per_image_limit = MAX_TEMPLATE_IMAGE_MB * 1024 * 1024

    with zipfile.ZipFile(BytesIO(template_bytes)) as z:
        for name in sorted(z.namelist()):
            if not name.startswith("ppt/media/"):
                continue
            ext = os.path.splitext(name)[1].lower()
            if ext not in RASTER_EXTS:
                continue
            with z.open(name) as f:
                data = f.read()
                if len(data) > per_image_limit:
                    continue
                h = hashlib.sha256(data).hexdigest()
                if h in seen_hashes:
                    continue
                images.append(data)
                seen_hashes.add(h)
                if len(images) >= MAX_TEMPLATE_IMAGES:
                    break
    return images

# ---------------- Theme parsing ----------------

def get_theme_style(template_bytes: bytes) -> Dict[str, Dict[str, str]]:
    """
    Parse theme colors and fonts from ppt/theme/theme*.xml.
    Returns:
      {
        "colors": { "dk1": "000000", "lt1": "FFFFFF", "accent1": "FFAA00", ... },
        "fonts":  { "major": "Calibri Light", "minor": "Calibri" }
      }
    Missing fields are omitted.
    """
    try:
        with zipfile.ZipFile(BytesIO(template_bytes)) as z:
            # Typically "ppt/theme/theme1.xml"; fall back to any matching theme file.
            theme_name = next((n for n in z.namelist() if n.startswith("ppt/theme/theme")), None)
            if not theme_name:
                return {"colors": {}, "fonts": {}}
            xml_bytes = z.read(theme_name)
    except Exception:
        return {"colors": {}, "fonts": {}}

    try:
        root = ET.fromstring(xml_bytes)
        ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}

        # Colors
        colors: Dict[str, str] = {}
        for tag in ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6"]:
            node = root.find(f".//a:clrScheme/a:{tag}", ns)
            if node is None:
                continue
            srgb = node.find(".//a:srgbClr", ns)
            if srgb is not None and "val" in srgb.attrib:
                colors[tag] = srgb.attrib["val"]

        # Fonts
        fonts: Dict[str, str] = {}
        major = root.find(".//a:fontScheme/a:majorFont/a:latin", ns)
        minor = root.find(".//a:fontScheme/a:minorFont/a:latin", ns)
        if major is not None and "typeface" in major.attrib:
            fonts["major"] = major.attrib["typeface"]
        if minor is not None and "typeface" in minor.attrib:
            fonts["minor"] = minor.attrib["typeface"]

        return {"colors": colors, "fonts": fonts}
    except Exception:
        return {"colors": {}, "fonts": {}}

# ---------------- Layout selection ----------------

def _layout_capabilities(layout: SlideLayout) -> Tuple[int, int]:
    """
    Return (#text-capable placeholders, #picture-capable placeholders) for a layout.
    We detect 'text-capable' by the presence of a text_frame on the placeholder.
    'picture-capable' approximated by placeholder type PICTURE or CONTENT.
    """
    text_capable = 0
    picture_capable = 0
    try:
        for shp in layout.placeholders:
            try:
                ptype = shp.placeholder_format.type
            except Exception:
                continue
            # Text-capable
            try:
                _ = shp.text_frame
                text_capable += 1
            except Exception:
                pass
            # Picture-capable
            if ptype in (PP_PLACEHOLDER.PICTURE, PP_PLACEHOLDER.CONTENT):
                picture_capable += 1
    except Exception:
        pass
    return text_capable, picture_capable

def _name_match_score(candidate: str, target: str) -> int:
    """
    Very light name similarity:
      - exact (case-insensitive) → 3
      - contains → 2
      - otherwise → 0
    """
    c = (candidate or "").strip().lower()
    t = (target or "").strip().lower()
    if not c or not t:
        return 0
    if c == t:
        return 3
    if t in c:
        return 2
    return 0

def _capability_ok(target: str, text_count: int, pic_count: int) -> bool:
    t = (target or "").lower()
    if "two content" in t:
        return text_count >= 2
    if "picture" in t:
        return pic_count >= 1
    if "content" in t or "caption" in t:
        return text_count >= 1
    return True

def find_preferred_layout(prs: Presentation, preferred_names: List[str]) -> Optional[SlideLayout]:
    """
    Find a layout that best matches any of the preferred names using:
      1) Exact (case-insensitive) name match.
      2) Contains/fuzzy match + capability check (text/picture placeholders).
    Returns the first good candidate; None if nothing is usable.
    """
    if not preferred_names:
        return None

    # Pass 1: exact case-insensitive match
    lowered = [p.lower() for p in preferred_names]
    for layout in prs.slide_layouts:
        try:
            if layout.name and layout.name.lower() in lowered:
                # Check capabilities roughly fit the preference
                txt, pic = _layout_capabilities(layout)
                target = preferred_names[lowered.index(layout.name.lower())]
                if _capability_ok(target, txt, pic):
                    return layout
        except Exception:
            continue

    # Pass 2: fuzzy contains + capability score
    best: Tuple[int, Optional[SlideLayout]] = (0, None)  # (score, layout)
    for layout in prs.slide_layouts:
        lname = getattr(layout, "name", "") or ""
        txt, pic = _layout_capabilities(layout)
        score = 0
        for pref in preferred_names:
            score = max(score, _name_match_score(lname, pref))
            # Penalize if capability doesn't fit at all
            if score > 0 and not _capability_ok(pref, txt, pic):
                score = 0
        if score > best[0]:
            best = (score, layout)

    return best[1] if best[0] > 0 else None

# ---------------- Dimensions ----------------

def get_ppt_dimensions(prs: Presentation) -> Dict[str, float]:
    """
    Return slide dimensions in multiple units for convenience.

    Args:
        prs: python-pptx Presentation instance.

    Returns:
        {
          "width_emu": int,
          "height_emu": int,
          "width_in": float,
          "height_in": float,
          "width_cm": float,
          "height_cm": float,
        }
    """
    w_emu = int(prs.slide_width)
    h_emu = int(prs.slide_height)
    w_in = w_emu / EMU_PER_INCH
    h_in = h_emu / EMU_PER_INCH
    w_cm = w_in * CM_PER_INCH
    h_cm = h_in * CM_PER_INCH
    return {
        "width_emu": w_emu,
        "height_emu": h_emu,
        "width_in": w_in,
        "height_in": h_in,
        "width_cm": w_cm,
        "height_cm": h_cm,
    }
