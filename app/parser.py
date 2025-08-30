# app/parser.py
"""
Heuristic fallback parser that maps raw text/markdown into a slide outline
if LLM is unavailable or fails. This parser aims to be:
- Markdown-aware (headings, lists, blockquotes, code fences)
- Archetype-aware (investor pitch, SOP, sales, research, lesson/quiz)
- Privacy-conscious (scrubs obvious secrets / PII)
- Layout-aware (choose reasonable layouts)
- Robust for empty/short/very long inputs
"""

from __future__ import annotations

import re
from typing import Dict, List, Any, Optional
from markdown_it import MarkdownIt

# ---------------- Tunables ----------------

MAX_BULLETS_PER_SLIDE = 7
MAX_CHARS_PER_BULLET = 160
MAX_CHARS_PER_SLIDE = 800  # soft budget; overflow splits into "(cont.)" slide(s)
MAX_SLIDES = 40
MIN_SLIDES = 3
DEFAULT_WORDS_PER_SLIDE = (60, 110)  # (min, max) band for estimating slide counts
SUB_BULLET_PREFIX = "  • "

ALLOWED_LAYOUTS = {
    "auto",
    "Title and Content",
    "Two Content",
    "Content with Caption",
    "Picture with Caption",
    "Blank",
}

# ---------------- Precompiled regexes ----------------

RE_WORD = re.compile(r"\w+")
RE_EMAIL = re.compile(r"\b[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}\b")
RE_URL = re.compile(r"\bhttps?://\S+\b")
RE_OPENAI_KEY = re.compile(r"\bsk-[A-Za-z0-9]{16,}\b")
RE_HEX_SECRET = re.compile(r"\b[a-fA-F0-9]{32,128}\b")
RE_PHONE = re.compile(r"\b(?:\+?\d[\d\-\s]{7,}\d)\b")
RE_CCARD = re.compile(r"\b(?:\d[ -]*?){13,19}\b")  # loose credit-card-ish matcher
RE_IMG_MD = re.compile(r"!\[.*?\]\(.*?\)")
RE_LINK_MD = re.compile(r"\[([^\]]+)\]\([^)]+\)")
RE_HTML_TAG = re.compile(r"<[^>]+>")

# Multilingual/robust sentence splitter (., !, ?, Chinese/Japanese punctuation)
RE_SENTENCES = re.compile(r"(?<=[。！？!?\.])\s+|(?<=\.)\s+|(?<=\?)\s+|(?<=!)\s+")

# ---------------- Utilities ----------------

def _word_count(s: str) -> int:
    return len(RE_WORD.findall(s or ""))

def _collapse_ws(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())

def _truncate(s: str, n: int) -> str:
    s = _collapse_ws(s or "")
    return s if len(s) <= n else s[: max(0, n - 1)].rstrip() + "…"

def _chunks(lst: List[str], n: int):
    for i in range(0, len(lst), n):
        yield lst[i : i + n]

def _dedup_preserve_order(items: List[str]) -> List[str]:
    seen = set()
    out: List[str] = []
    for x in items:
        if x not in seen:
            seen.add(x)
            out.append(x)
    return out

def _strip_markup(text: str) -> str:
    """Remove MD images/links and HTML tags but keep visible label text."""
    t = RE_IMG_MD.sub("", text or "")
    t = RE_LINK_MD.sub(r"\1", t)
    t = RE_HTML_TAG.sub("", t)
    return t

def _scrub_sensitive(text: str) -> str:
    """Redact obvious secrets/PII unless essential."""
    if not text:
        return text
    t = text
    t = RE_EMAIL.sub("[…]", t)
    t = RE_URL.sub("[…]", t)
    t = RE_OPENAI_KEY.sub("[…]", t)
    t = RE_HEX_SECRET.sub("[…]", t)
    t = RE_PHONE.sub("[…]", t)
    # Be conservative on credit cards; only scrub when multiple groups present
    if len(re.findall(r"\d", t)) >= 12:
        t = RE_CCARD.sub("[…]", t)
    return t

def _likely_legal(text: str) -> bool:
    return bool(re.search(r"\b(policy|compliance|gdpr|hipaa|terms|contract|license|liability)\b", text, re.I))

def _likely_medical(text: str) -> bool:
    return bool(re.search(r"\b(clinical|diagnos|treatment|adverse|contraindication|guideline|prescrib)\b", text, re.I))

def _has_meaningful_notes(include_notes: bool, notes_text: str) -> bool:
    return include_notes and bool((notes_text or "").strip())

# ---------------- Archetypes ----------------

def _detect_archetype(guidance: str) -> Optional[str]:
    g = (guidance or "").lower()
    if "investor" in g or "pitch" in g:
        return "investor"
    if "sop" in g or "runbook" in g or "standard operating" in g:
        return "sop"
    if "sales" in g:
        return "sales"
    if "research" in g or "conference" in g or "talk" in g or "paper" in g:
        return "research"
    if "lesson" in g or "quiz" in g or "lecture" in g or "teaching" in g:
        return "lesson"
    return None

def _archetype_sections(kind: str) -> List[str]:
    if kind == "investor":
        return [
            "Problem", "Solution", "Market", "Product", "Moat",
            "Go-To-Market", "Traction", "Business Model", "Competition",
            "Team", "Financials", "Ask", "Roadmap"
        ]
    if kind == "sop":
        return ["Purpose", "Scope", "Prerequisites", "Procedure", "Validation/Checks", "Rollback/Recovery", "Contact/On-call"]
    if kind == "sales":
        return ["Overview", "Value Proposition", "ROI/Impact", "Case Studies", "Pricing", "Call to Action"]
    if kind == "research":
        return ["Background", "Methods", "Results", "Limitations", "Future Work", "References/Acknowledgements"]
    if kind == "lesson":
        return ["Objectives", "Key Concepts", "Examples", "Practice Questions", "Summary"]
    return []

def _keyword_bucket(sentence: str, sections: List[str]) -> int:
    """Heuristic mapping of a sentence to a section index via keywords; else round-robin by hash."""
    s = (sentence or "").lower()
    hints = {
        "problem": ["problem", "pain", "gap"],
        "solution": ["solution", "approach", "proposal"],
        "market": ["market", "tam", "sam", "som"],  # fixed acronyms
        "product": ["product", "feature", "prototype", "architecture"],
        "moat": ["moat", "defensib", "advantage", "ip", "patent"],
        "go-to-market": ["gtm", "marketing", "sales", "channel"],
        "traction": ["traction", "revenue", "users", "growth"],
        "business model": ["pricing", "business model", "subscription", "margin"],
        "competition": ["competitor", "competition", "alternative"],
        "team": ["team", "hiring", "founder"],
        "financials": ["financial", "projection", "cost", "profit", "loss", "burn"],
        "ask": ["ask", "raise", "fund", "investment"],
        "roadmap": ["roadmap", "timeline", "milestone"],
        "purpose": ["purpose", "objective"],
        "scope": ["scope", "coverage"],
        "prerequisites": ["prerequisite", "requirement", "dependency"],
        "procedure": ["step", "procedure", "instruction"],
        "validation/checks": ["validate", "check", "verify"],
        "rollback/recovery": ["rollback", "recovery", "restore"],
        "contact/on-call": ["contact", "on-call", "escalation"],
        "value proposition": ["value", "benefit", "advantage", "roi"],
        "roi/impact": ["roi", "impact", "benefit"],
        "case studies": ["case", "study", "example"],
        "call to action": ["cta", "contact", "next step", "trial"],
        "background": ["background", "intro", "motivation"],
        "methods": ["method", "algorithm", "procedure"],
        "results": ["result", "finding", "outcome"],
        "limitations": ["limit", "constraint", "threat"],
        "future work": ["future", "next", "expand"],
        "references/acknowledgements": ["reference", "cite", "acknowledgement"],
        "objectives": ["objective", "goal", "outcome"],
        "key concepts": ["concept", "definition", "theory"],
        "examples": ["example", "illustration"],
        "practice questions": ["question", "quiz", "mcq"],
        "summary": ["summary", "conclusion", "recap"],
    }
    for i, sec in enumerate(sections):
        need = hints.get(sec.lower(), [])
        if any(k in s for k in need):
            return i
    return hash(s) % max(1, len(sections))

# ---------------- Guidance influences ----------------

def _layout_bias_from_guidance(guidance: str) -> str:
    g = (guidance or "").lower()
    if any(k in g for k in ("visual", "image", "design-heavy", "poster")):
        return "Picture with Caption"
    if any(k in g for k in ("executive", "summary", "tl;dr")):
        return "Content with Caption"
    if any(k in g for k in ("technical", "deep dive", "details")):
        return "Two Content"
    return "auto"

def _bullet_target_from_guidance(guidance: str) -> int:
    g = (guidance or "").lower()
    if any(k in g for k in ("executive", "brief", "summary")):
        return 3
    if any(k in g for k in ("technical", "detailed", "thorough")):
        return 6
    return 5

# ---------------- Char-budget enforcement ----------------

def _split_by_char_budget(title: str, bullets: List[str]) -> List[Dict[str, Any]]:
    """
    If bullets collectively exceed MAX_CHARS_PER_SLIDE, split into multiple slides.
    """
    out: List[Dict[str, Any]] = []
    buf: List[str] = []
    running = 0
    for b in bullets:
        b = (b or "").strip()
        if not b:
            continue
        blen = len(b)
        if running + blen > MAX_CHARS_PER_SLIDE and buf:
            out.append({"title": title, "bullets": buf})
            buf, running = [], 0
        buf.append(b)
        running += blen
    if buf:
        out.append({"title": title, "bullets": buf})
    # add (cont.) to subsequent titles
    for i in range(1, len(out)):
        out[i]["title"] = f"{title} (cont.)"
    return out

# ---------------- Notes generation ----------------

def _generate_notes_from_bullets(bullets: List[str]) -> str:
    """
    Simple 1–2 sentence synthesis from bullets for speaker notes.
    """
    if not bullets:
        return ""
    core = bullets[: min(4, len(bullets))]
    sent = "; ".join(_collapse_ws(b) for b in core)
    if len(core) >= 3:
        return _truncate(f"Key points: {sent}.", 380)
    return _truncate(sent, 380)

# ---------------- Core Parser ----------------

def heuristic_outline(text: str, guidance: str = "", include_notes: bool = False) -> Dict[str, Any]:
    """
    Build a slide outline without the LLM:
    - Prefer headings as slide titles.
    - Use list items, blockquotes, and paragraphs as bullets.
    - If no headings, derive sectioned slides from archetype (if any) or sentence chunking.
    - Enforce per-slide char budgets and sensible layouts.
    """
    raw = text or ""
    md = MarkdownIt()
    tokens = md.parse(raw)

    slides: List[Dict[str, Any]] = []
    current_title: Optional[str] = None
    current_bullets: List[str] = []
    list_level = 0

    # Guidance biases
    guidance_layout_bias = _layout_bias_from_guidance(guidance)
    guidance_bullet_target = _bullet_target_from_guidance(guidance)

    def flush_slide():
        nonlocal current_title, current_bullets
        if not (current_title or current_bullets):
            return
        # Clean bullets
        bullets = [
            _scrub_sensitive(_truncate(_strip_markup(b), MAX_CHARS_PER_BULLET))
            for b in current_bullets
            if b and b.strip()
        ]
        bullets = _dedup_preserve_order([b for b in bullets if b.strip()])

        # Enforce bullet cap & char budget, possibly splitting into continuation slides
        if not bullets:
            bullets = []
        split = _split_by_char_budget(current_title or "Overview", bullets)
        for idx, part in enumerate(split):
            chosen_layout = "Two Content" if len(part["bullets"]) > MAX_BULLETS_PER_SLIDE else "auto"
            # Apply gentle guidance bias if "auto"
            if chosen_layout == "auto" and guidance_layout_bias in ALLOWED_LAYOUTS:
                chosen_layout = guidance_layout_bias
            slide: Dict[str, Any] = {
                "title": part["title"] if idx == 0 else f"{(current_title or 'Overview')} (cont.)",
                "bullets": part["bullets"][:MAX_BULLETS_PER_SLIDE],
                "layout": chosen_layout if chosen_layout in ALLOWED_LAYOUTS else "auto",
            }
            if include_notes:
                slide["notes"] = _generate_notes_from_bullets(slide["bullets"])
            slides.append(slide)

        current_title, current_bullets = None, []

    # Pass 1: Markdown-aware extraction
    i = 0
    L = len(tokens)
    while i < L:
        t = tokens[i]

        if t.type == "heading_open":
            flush_slide()
            # Next inline contains the heading text
            if i + 1 < L and tokens[i + 1].type == "inline":
                current_title = _truncate(_strip_markup(tokens[i + 1].content), 80)
                i += 2
                continue

        if t.type in ("bullet_list_open", "ordered_list_open"):
            list_level += 1

        elif t.type in ("bullet_list_close", "ordered_list_close"):
            list_level = max(0, list_level - 1)

        elif t.type == "list_item_open":
            # Eat everything until list_item_close and capture inline text
            j = i + 1
            text_buf = ""
            while j < L and tokens[j].type not in ("list_item_close", "list_item_open"):
                if tokens[j].type == "inline":
                    text_buf += " " + tokens[j].content
                j += 1
            text_buf = _collapse_ws(_strip_markup(text_buf))
            if text_buf:
                prefix = SUB_BULLET_PREFIX if list_level >= 2 else ""
                current_bullets.append(prefix + text_buf)
            i = j
            continue

        elif t.type == "blockquote_open":
            # Capture the quoted inline text as a bullet prefixed with “Quote:”
            j = i + 1
            quote_buf = ""
            while j < L and tokens[j].type != "blockquote_close":
                if tokens[j].type == "inline":
                    quote_buf += " " + tokens[j].content
                j += 1
            quote_clean = _collapse_ws(_strip_markup(quote_buf))
            if quote_clean:
                current_bullets.append(f"Quote: {quote_clean}")
            i = j
            continue

        elif t.type == "table_open":
            # Summarize tables as lines captured until table_close
            j = i + 1
            rows: List[str] = []
            row = ""
            while j < L and tokens[j].type != "table_close":
                if tokens[j].type == "inline":
                    row = _collapse_ws(_strip_markup(tokens[j].content))
                    if row:
                        rows.append(row)
                j += 1
            if rows:
                for r in rows[: guidance_bullet_target + 2]:
                    current_bullets.append(_truncate(r, MAX_CHARS_PER_BULLET))
            i = j
            continue

        elif t.type == "paragraph_open":
            if i + 1 < L and tokens[i + 1].type == "inline":
                content = _strip_markup(tokens[i + 1].content)
                stripped = _collapse_ws(content)
                if stripped:
                    current_bullets.append(stripped)

        elif t.type == "fence":  # code block
            lang = (t.info or "").strip() or "code"
            code_lines = (t.content or "").splitlines()
            approx = len(code_lines)
            current_bullets.append(f"Code: {lang} block (~{approx} lines) – summary unavailable")

        i += 1

    flush_slide()

    # If we produced nothing meaningful, try archetype bucketing or sentence chunking
    have_real_titles = any(s["title"] and s["title"] not in ("Overview", "Slide") for s in slides)
    archetype = _detect_archetype(guidance)

    if not slides or not have_real_titles:
        sentences = [s for s in RE_SENTENCES.split(_collapse_ws(raw)) if s and not re.fullmatch(r"[.!?]+", s)]
        # Archetype mapping first
        if archetype:
            sections = _archetype_sections(archetype)
            bucketed: List[List[str]] = [[] for _ in sections] if sections else []
            if sections:
                for s in sentences:
                    bucketed[_keyword_bucket(s, sections)].append(_truncate(s, MAX_CHARS_PER_BULLET))
                slides = []
                for sec, group in zip(sections, bucketed):
                    if not group:
                        continue
                    bullets = _dedup_preserve_order(group)
                    # char-budget split
                    split_parts = _split_by_char_budget(sec, bullets)
                    for part in split_parts:
                        chosen_layout = "Two Content" if len(part["bullets"]) > MAX_BULLETS_PER_SLIDE else guidance_layout_bias
                        if chosen_layout not in ALLOWED_LAYOUTS:
                            chosen_layout = "auto"
                        slides.append({
                            "title": part["title"],
                            "bullets": part["bullets"][:MAX_BULLETS_PER_SLIDE],
                            "layout": chosen_layout,
                            **({"notes": _generate_notes_from_bullets(part['bullets'])} if include_notes else {}),
                        })
        # Generic chunking if still empty
        if not slides:
            total_words = _word_count(raw)
            avg_words_per_slide = sum(DEFAULT_WORDS_PER_SLIDE) // 2
            approx_slides = max(MIN_SLIDES, min(25, total_words // max(1, avg_words_per_slide)))
            approx_slides = max(MIN_SLIDES, min(MAX_SLIDES, approx_slides))
            if not sentences:
                sentences = [raw] if raw.strip() else []
            group_size = max(1, len(sentences) // max(1, approx_slides))
            slides = []
            for idx, group in enumerate(_chunks(sentences, group_size), 1):
                bul = []
                for sent in group:
                    if sent:
                        bul.append(_truncate(_strip_markup(sent), MAX_CHARS_PER_BULLET))
                if bul:
                    # enforce char budget again
                    parts = _split_by_char_budget(f"Section {idx}", bul)
                    for p in parts:
                        slides.append({
                            "title": p["title"],
                            "bullets": _dedup_preserve_order(p["bullets"])[:MAX_BULLETS_PER_SLIDE],
                            "layout": guidance_layout_bias if guidance_layout_bias in ALLOWED_LAYOUTS else "auto",
                            **({"notes": _generate_notes_from_bullets(p['bullets'])} if include_notes else {}),
                        })

    # Disclaimers: legal/medical
    raw_lower = (raw or "").lower()
    if _likely_legal(raw_lower) or _likely_medical(raw_lower):
        disclaimer = "Informational only; not legal/medical advice."
        if include_notes and slides:
            slides[0]["notes"] = _truncate(
                ((slides[0].get("notes", "") + " ").strip() + disclaimer).strip(), 400
            )
        elif slides:
            last_bul = slides[-1].get("bullets", [])
            if disclaimer not in last_bul:
                last_bul.append(disclaimer)
            slides[-1]["bullets"] = last_bul

    # Final cleanup pass (normalize layouts, enforce min/max slide count)
    cleaned: List[Dict[str, Any]] = []
    for s in slides:
        title = _truncate(s.get("title") or "Slide", 80)
        bullets = [b for b in (s.get("bullets") or []) if b and b.strip()]
        bullets = [_scrub_sensitive(_truncate(b, MAX_CHARS_PER_BULLET)) for b in bullets]
        bullets = _dedup_preserve_order(bullets)

        layout = s.get("layout") or "auto"
        if layout not in ALLOWED_LAYOUTS:
            layout = "auto"
        if _has_meaningful_notes(include_notes, s.get("notes", "")) and layout == "auto":
            layout = "Content with Caption"
        if len(bullets) > MAX_BULLETS_PER_SLIDE and layout != "Two Content":
            layout = "Two Content"

        out = {"title": title, "bullets": bullets[:MAX_BULLETS_PER_SLIDE], "layout": layout}
        if include_notes:
            notes = s.get("notes") or _generate_notes_from_bullets(out["bullets"])
            out["notes"] = _truncate(notes, 400)
        # drop truly empty slides
        if out["title"] or out["bullets"]:
            cleaned.append(out)

    # Ensure at least MIN_SLIDES (split dense slides first, then pad title-only)
    def _ensure_min_slides(slides_in: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        if len(slides_in) >= MIN_SLIDES:
            return slides_in
        out: List[Dict[str, Any]] = []
        for s in slides_in:
            out.append(s)
            # split if too many bullets and we need more slides
            while len(out) < MIN_SLIDES and len(s.get("bullets", [])) > 3:
                extra = s["bullets"][3:]
                s["bullets"] = s["bullets"][:3]
                cont = {"title": f"{s['title']} (cont.)", "bullets": extra[:3], "layout": s["layout"]}
                if include_notes:
                    cont["notes"] = _generate_notes_from_bullets(cont["bullets"])
                out.append(cont)
        while len(out) < MIN_SLIDES:
            pad_idx = len(out) + 1
            pad = {"title": f"Slide {pad_idx}", "bullets": [], "layout": "Blank"}
            if include_notes:
                pad["notes"] = ""
            out.append(pad)
        return out

    cleaned = _ensure_min_slides(cleaned)[:MAX_SLIDES]

    # Title selection
    deck_title = None
    for s in cleaned:
        if s["title"] not in ("Overview", "Slide", "Section 1"):
            deck_title = s["title"]
            break
    if not deck_title:
        deck_title = f"Generated Presentation — {_truncate(guidance, 60)}" if guidance else "Generated Presentation"

    return {
        "title": deck_title,
        "slides": cleaned,
        "estimated_slide_count": len(cleaned),
    }
