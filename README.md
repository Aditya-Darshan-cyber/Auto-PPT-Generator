---
title: Auto PPT Generator
emoji: 📊
colorFrom: blue
colorTo: indigo
sdk: docker
pinned: false
license: mit
---

# Auto PPT Generator — Turn Text/Markdown into a Themed PowerPoint 🎯

Small FastAPI app that turns bulk **text or markdown** into a fully formatted **PowerPoint (.pptx)** that matches an **uploaded template’s** look & feel — fonts, theme colors, layouts, and (optionally) reuses images already inside the template.  
No AI image generation.

- **LLM planner (optional):** via **AI Pipe** (OpenAI-compatible). Users bring their own AI Pipe token.
- **Model:** `gpt-4.1-mini` (default).
- **Fallback:** deterministic, markdown-aware parser when no token is provided or the LLM fails.
- **Privacy:** Keys are **never stored or logged**. Templates are validated to avoid zip-bombs.

---

## ✨ Features

- Paste **text/markdown** + optional one-line **guidance** (e.g., “investor pitch deck”).
- Upload a `.pptx`/`.potx` **template**; the deck inherits theme **fonts**, **colors**, and **layouts**.
- **Optional speaker notes** generation.
- **Preview** the outline JSON before building the PPTX.
- **Heuristic fallback** parser ensures it works without an LLM token.

---

## 📦 Project Structure

.
├── app/
│   ├── init.py
│   ├── config.py
│   ├── llm_clients.py
│   ├── main.py
│   ├── parser.py
│   ├── pptx_builder.py
│   ├── schemas.py
│   └── template_utils.py
├── web/
│   └── index.html # (optional UI; you can add later)
├── writeup.md # 200–300 word short write-up
├── requirements.txt
├── Dockerfile
├── README.md
└── LICENSE # MIT

---

## 🚀 Quick Start (Local)

**Prereqs:** Python 3.11+

```bash
python -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt

# (optional) defaults for AI Pipe (OpenAI-compatible)
export OPENAI_BASE_URL="https://aipipe.org/openai/v1"
export OPENAI_MODEL="gpt-4.1-mini"

uvicorn app.main:app --host 0.0.0.0 --port 7860
# open http://localhost:7860
````

If you include a `web/index.html`, it will be served at `/`.
Otherwise, use the API directly (see below) or visit `/docs` for Swagger.

### 🐳 Run with Docker (local)

```bash
docker build -t Auto_PPT_Generator .
docker run --rm -p 7860:7860 \
  -e OPENAI_BASE_URL="https://aipipe.org/openai/v1" \
  -e OPENAI_MODEL="gpt-4.1-mini" \
  Auto_PPT_Generator
# open http://localhost:7860
```

### 🤝 Deploy on Hugging Face Spaces (Docker)

This repo includes a Dockerfile. Create a Docker Space, connect this repo, and it will build and run automatically.

The server listens on `$PORT` (Spaces sets it). Default is `7860`.

Users paste their AI Pipe token in the UI (or pass in API). No server secrets needed.

---

## 🔌 AI Pipe (OpenAI-compatible)

Get a token at [https://aipipe.org/login](https://aipipe.org/login).

The app defaults to `OPENAI_BASE_URL=https://aipipe.org/openai/v1` and `OPENAI_MODEL=gpt-4.1-mini`.

Users paste their token client-side (or pass it in the API as `api_key`).

No keys are stored or logged.

---

## 🧪 API

### Health

```http
GET /healthz
```

### Preview outline (no file)

```http
POST /api/preview_outline
```

Form fields:

* `text` (str, required)
* `guidance` (str, optional)
* `api_key` (str, optional)           # AI Pipe token; if omitted → heuristic parser
* `provider` (str, default `"openai"`)
* `model` (str, default `"gpt-4.1-mini"`)
* `base_url` (str, default `"https://aipipe.org/openai/v1"`)
* `include_notes` (bool-ish, default `"false"`)

**curl example**

```bash
curl -s -X POST http://localhost:7860/api/preview_outline \
  -F 'text=# Title\n\n- point A\n- point B' \
  -F 'guidance=investor pitch deck' \
  -F 'include_notes=true' | jq .
```

### Generate PPTX (multipart with template)

```http
POST /api/generate
```

Form fields:

* `text` (str, required)
* `guidance` (str, optional)
* `api_key` (str, optional)           # AI Pipe token; if omitted → heuristic parser
* `provider/model/base_url/include_notes` (same as preview)
* `template` (file, required)         # .pptx or .potx, ≤ MAX\_FILE\_MB

**curl example**

```bash
curl -s -X POST http://localhost:7860/api/generate \
  -F 'text=# Intro\n\nSome content.' \
  -F 'guidance=research talk' \
  -F 'include_notes=false' \
  -F 'template=@your-template.pptx' \
  -o deck.pptx
```

---

## ⚙️ Configuration

Environment variables (sane defaults baked in):

```
Var                   Default                         Purpose
OPENAI_BASE_URL       https://aipipe.org/openai/v1    OpenAI-compatible endpoint via AI Pipe
OPENAI_MODEL          gpt-4.1-mini                    LLM used by the planner
MAX_FILE_MB           20                              Max template size (MB)
MAX_TEXT_CHARS        40000                           Clamp for input text length
LLM_TIMEOUT_SECS      60                              Request timeout for LLM calls
LLM_MAX_RETRIES       3                               Retry attempts on network/5xx
```

Additional limits (in `app/config.py`):

* `MAX_BULLETS_PER_SLIDE` (default 7)
* `MAX_TITLE_CHARS`, `MAX_BULLET_CHARS`, `MAX_NOTES_CHARS`
* `MAX_TOTAL_SLIDES` (default 60)
* Zip safety knobs: `MAX_ZIP_ENTRIES`, `MAX_ZIP_MEMBER_MB`
* Template image extraction limits: `MAX_TEMPLATE_IMAGES`, `MAX_TEMPLATE_IMAGE_MB`

Accepted templates: `.pptx`, `.potx`
Security checks: verifies PPTX structure and guards against zip-bombs.

---

## 🧠 How it works (high level)

**Planner (LLM path):** Uses a strict JSON prompt (via AiPipe → OpenAI-compatible API) to produce slide objects with titles, concise bullets, optional notes, and layout hints. It’s archetype-aware (e.g., investor pitch, SOP, research talk) and enforces no new image generation.

**Fallback parser:** Markdown-aware; handles headings, lists, and code fences. When headings are missing, it chunks sentences and, if guidance implies an archetype, buckets content into that structure.

**Builder:** Loads the uploaded template, extracts theme fonts/colors (`ppt/theme/theme*.xml`), applies them to titles & bullets, honors layout hints, and reuses `ppt/media/*` images only where a picture placeholder exists.

See `writeup.md` for the 200–300 word description.

---

## 🔒 Privacy & Safety

* API keys/tokens are never logged or stored.
* PPTX files are validated (structure, entry counts, member sizes).
* Secrets/PII in text (URLs, emails, tokens) are scrubbed in the fallback path.

---

## 🐞 Troubleshooting

**“Invalid or unsafe PowerPoint file.”**
Ensure you upload a real `.pptx`/`.potx` (not a renamed file), ≤ `MAX_FILE_MB`, and not password-protected.

**LLM errors or empty JSON:**
The app auto-falls back to the heuristic parser; check your AiPipe token and network if you want LLM output.

**Template looks ignored:**
Try a different layout. The builder applies theme fonts/colors and picks the closest available layout by capability.

---

## 📄 License

MIT — see LICENSE.
