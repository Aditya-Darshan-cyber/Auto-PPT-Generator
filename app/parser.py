# app/parser.py
"""
Heuristic fallback parser that maps raw text/markdown into a slide outline
if LLM is unavailable or fails. This parser aims to be:
- Markdown-aware (headings, lists, code fences)
- Archetype-aware (investor pitch, SOP, sales, research, lesson/quiz)
- Privacy-conscious (scrubs obvious secrets / PII)
- Layout-aware (choose reasonable layouts)
- Robust for empty/short/very long inputs
"""

import re
from typing import Dict, List, Any, Optional
from markdown_it import MarkdownIt

# ---------------- Tunables ----------------

MAX_BULLETS_PER_SLIDE = 7
MAX_CHARS_PER_BULLET = 160
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

# ---------------- Utilities ----------------

def _word_count(s: str) -> int:
    return len(re.findall(r"\w+", s or ""))

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
    out = []
    for x in items:
        if x not in seen:
            seen.add(x)
            out.append(x)
    return out

def _scrub_sensitive(text: str) -> str:
    """Redact obvious secrets/PII unless essential."""
    if not text:
        return text
    t = text
    # Emails
    t = re.sub(r"\b[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}\b", "[…]", t)
    # URLs
    t = re.sub(r"\bhttps?://\S+\b", "[…]", t)
    # OpenAI-like keys / tokens (e.g., sk-... long)
    t = re.sub(r"\bsk-[A-Za-z0-9]{16,}\b", "[…]", t)
    # Long hex secrets (32–128 hex chars)
    t = re.sub(r"\b[a-fA-F0-9]{32,128}\b", "[…]", t)
    # Phone-like sequences (simple heuristic)
    t = re.sub(r"\b(?:\+?\d[\d\-\s]{7,}\d)\b", "[…]", t)
    return t

def _likely_legal(text: str) -> bool:
    return bool(re.search(r"\b(policy|compliance|gdpr|hipaa|terms|contract|license|liability)\b", text, re.I))

def _likely_medical(text: str) -> bool:
    return bool(re.search(r"\b(clinical|diagnos|treatment|adverse|contraindication|guideline|prescrib)\b", text, re.I))

def _has_meaningful_notes(include_notes: bool, notes_text: str) -> bool:
    return include_notes and bool(notes_text.strip())

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
    s = sentence.lower()
    # Simple hints
    hints = {
        "problem": ["problem", "pain", "gap"],
        "solution": ["solution", "approach", "proposal"],
        "market": ["market", "tamar", "sam", "som"],
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
    # Try exact section keyword match
    for i, sec in enumerate(sections):
        need = hints.get(sec.lower(), [])
        if any(k in s for k in need):
            return i
    # Fallback: stable round-robin-ish bucketing
    return hash(s) % max(1, len(sections))

# ---------------- Core Parser ----------------

def heuristic_outline(text: str, guidance: str = "", include_notes: bool = False) -> Dict[str, Any]:
    """
    Build a slide outline without the LLM:
    - Prefer headings as slide titles.
    - Use list items and paragraphs as bullets.
    - If no headings, derive sectioned slides from archetype (if any) or sentence chunking.
    """
    raw = text or ""
    md = MarkdownIt()
    tokens = md.parse(raw)

    slides: List[Dict[str, Any]] = []
    current_title: Optional[str] = None
    current_bullets: List[str] = []
    list_level = 0
    heading_stack: List[int] = []  # track heading levels for better titling

    def flush_slide():
        nonlocal current_title, current_bullets
        if current_title or current_bullets:
            # Dedup, truncate, scrub
            bullets = [
                _scrub_sensitive(_truncate(b, MAX_CHARS_PER_BULLET))
                for b in current_bullets
                if b and b.strip()
            ]
            bullets = _dedup_preserve_order([b for b in bullets if b.strip()])
            # Pick layout
            layout = "Two Content" if len(bullets) > MAX_BULLETS_PER_SLIDE else "auto"
            slide: Dict[str, Any] = {
                "title": current_title or "Overview",
                "bullets": bullets[:MAX_BULLETS_PER_SLIDE],
                "layout": layout if layout in ALLOWED_LAYOUTS else "auto",
            }
            if include_notes:
                slide["notes"] = ""
            slides.append(slide)
        current_title, current_bullets = None, []

    # Pass 1: Markdown-aware extraction
    i = 0
    L = len(tokens)
    while i < L:
        t = tokens[i]
        if t.type == "heading_open":
            # t.tag is h1..h6
            flush_slide()
            level = int(t.tag[-1]) if t.tag[-1].isdigit() else 2
            heading_stack.append(level)
            # Next inline should contain the title text
            if i + 1 < L and tokens[i + 1].type == "inline":
                current_title = _truncate(tokens[i + 1].content, 80)
                i += 2
                continue

        if t.type in ("bullet_list_open", "ordered_list_open"):
            list_level += 1

        elif t.type in ("bullet_list_close", "ordered_list_close"):
            list_level = max(0, list_level - 1)

        elif t.type == "list_item_open":
            # next inline paragraph contains content
            # markdown-it often emits: list_item_open -> paragraph_open -> inline -> paragraph_close -> list_item_close
            # We sniff forward for inline text
            j = i + 1
            text_buf = ""
            while j < L and tokens[j].type not in ("list_item_close", "list_item_open"):
                if tokens[j].type == "inline":
                    text_buf += " " + tokens[j].content
                j += 1
            text_buf = _collapse_ws(text_buf)
            if text_buf:
                prefix = SUB_BULLET_PREFIX if list_level >= 2 else ""
                # Remove image/link-only syntaxes
                text_buf = re.sub(r"!\[.*?\]\(.*?\)", "", text_buf)
                text_buf = re.sub(r"\[([^\]]+)\]\([^)]+\)", r"\1", text_buf)
                current_bullets.append(prefix + text_buf)
            i = j
            continue

        elif t.type == "paragraph_open":
            # paragraph appears; we'll read its inline sibling
            if i + 1 < L and tokens[i + 1].type == "inline":
                content = tokens[i + 1].content
                # Ignore pure images or links
                stripped = re.sub(r"!\[.*?\]\(.*?\)", "", content)
                stripped = re.sub(r"\[([^\]]+)\]\([^)]+\)", r"\1", stripped).strip()
                if stripped:
                    current_bullets.append(stripped)
            # skip paragraph_close handled by loop

        elif t.type == "fence":  # code block
            lang = (t.info or "").strip() or "code"
            code_lines = (t.content or "").splitlines()
            approx = len(code_lines)
            current_bullets.append(f"Code: {lang} block (~{approx} lines) – summary unavailable")

        i += 1

    flush_slide()

    # If we produced nothing or only generic titles, try archetype or chunking
    have_real_titles = any(s["title"] and s["title"] != "Overview" for s in slides)
    archetype = _detect_archetype(guidance)

    if not slides or not have_real_titles:
        # Build content sentences
        sentences = re.split(r"(?<=[。！？!?\.])\s+|(?<=\.)\s+|(?<=\?)\s+|(?<=!)\s+", _collapse_ws(raw))
        sentences = [s for s in sentences if s and not re.fullmatch(r"[.!?]+", s)]
        # Use archetype sections if possible
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
                    bullets = _dedup_preserve_order(group)[:MAX_BULLETS_PER_SLIDE]
                    slides.append({
                        "title": sec,
                        "bullets": bullets,
                        "layout": "Two Content" if len(bullets) > MAX_BULLETS_PER_SLIDE else "auto",
                        **({"notes": ""} if include_notes else {}),
                    })
        # If still empty or no archetype, sentence chunking
        if not slides:
            total_words = _word_count(raw)
            avg_words_per_slide = sum(DEFAULT_WORDS_PER_SLIDE) // 2
            approx_slides = max(MIN_SLIDES, min(25, total_words // max(1, avg_words_per_slide)))
            approx_slides = max(MIN_SLIDES, min(MAX_SLIDES, approx_slides))
            group_size = max(1, len(sentences) // max(1, approx_slides))
            slides = []
            for idx, group in enumerate(_chunks(sentences, group_size), 1):
                bul = []
                for sent in group:
                    if sent:
                        bul.append(_truncate(sent, MAX_CHARS_PER_BULLET))
                if bul:
                    slides.append({
                        "title": f"Section {idx}",
                        "bullets": _dedup_preserve_order(bul)[:MAX_BULLETS_PER_SLIDE],
                        "layout": "auto",
                        **({"notes": ""} if include_notes else {}),
                    })

    # Disclaimers: legal/medical
    raw_lower = raw.lower()
    if _likely_legal(raw_lower) or _likely_medical(raw_lower):
        disclaimer = "Informational only; not legal/medical advice."
        if include_notes and slides:
            # Put once in the first slide notes for visibility
            slides[0]["notes"] = (slides[0].get("notes", "") + (" " if slides[0].get("notes") else "") + disclaimer).strip()
        else:
            # Append to last slide bullets if notes not used
            if slides:
                last_bul = slides[-1].get("bullets", [])
                if disclaimer not in last_bul:
                    last_bul.append(disclaimer)
                slides[-1]["bullets"] = last_bul

    # Final layout pass + bullet cleanup
    cleaned_slides: List[Dict[str, Any]] = []
    for s in slides:
        title = _truncate(s.get("title") or "Slide", 80)
        bullets = [b for b in s.get("bullets", []) if b.strip()]
        bullets = [_scrub_sensitive(_truncate(b, MAX_CHARS_PER_BULLET)) for b in bullets]
        bullets = _dedup_preserve_order(bullets)
        layout = s.get("layout") or "auto"
        if layout not in ALLOWED_LAYOUTS:
            layout = "auto"
        # Promote layout to 'Content with Caption' if meaningful notes are present
        if _has_meaningful_notes(include_notes, s.get("notes", "")) and layout == "auto":
            layout = "Content with Caption"
        if len(bullets) > MAX_BULLETS_PER_SLIDE and layout != "Two Content":
            layout = "Two Content"
        out = {"title": title, "bullets": bullets[:MAX_BULLETS_PER_SLIDE], "layout": layout}
        if include_notes:
            out["notes"] = _truncate(s.get("notes", ""), 400)
        cleaned_slides.append(out)

    # Title selection
    deck_title = None
    # Use the first H1/H2 we encountered (current_title captured during parse), else build from guidance
    # Since we don't retain heading levels beyond slide titles, we infer from first non-generic slide title.
    for s in cleaned_slides:
        if s["title"] not in ("Overview", "Slide", "Section 1"):
            deck_title = s["title"]
            break
    if not deck_title:
        base = "Generated Presentation"
        deck_title = f"{base} — {_truncate(guidance, 60)}" if guidance else base

    # Trim total slides to MAX_SLIDES
    cleaned_slides = cleaned_slides[:MAX_SLIDES]

    return {
        "title": deck_title,
        "slides": cleaned_slides,
        "estimated_slide_count": len(cleaned_slides),
    }
