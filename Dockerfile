# ------- base image -------
FROM python:3.11-slim

# System deps (runtime libs for Pillow) + curl for healthcheck
RUN apt-get update && apt-get install -y --no-install-recommends \
    curl \
    libjpeg62-turbo \
    libpng16-16 \
    libtiff5 \
    libfreetype6 \
    liblcms2-2 \
    libopenjp2-7 \
    libwebp7 \
  && rm -rf /var/lib/apt/lists/*

# Env: fast, quiet Python
ENV PIP_DISABLE_PIP_VERSION_CHECK=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

# Defaults for AI Pipe (users still paste their token in the UI)
ENV OPENAI_BASE_URL="https://aipipe.org/openai/v1" \
    DEFAULT_PROVIDER="openai" \
    OPENAI_MODEL="gpt-4.1-mini" \
    PORT=7860

WORKDIR /app

# Install deps first for layer caching
COPY requirements.txt /app/requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

# Copy source
COPY app /app/app
COPY web /app/web
COPY README.md /app/README.md
COPY writeup.md /app/writeup.md

# Expose (HF/Render respect $PORT)
EXPOSE 7860

# Healthcheck
HEALTHCHECK --interval=30s --timeout=5s --start-period=20s --retries=3 \
  CMD curl -fsS "http://127.0.0.1:${PORT}/healthz" || exit 1

# Run the API
CMD ["bash", "-lc", "uvicorn app.main:app --host 0.0.0.0 --port ${PORT}"]
