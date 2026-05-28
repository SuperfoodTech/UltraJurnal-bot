# =============================================================================
#  Dockerfile — UltraJurnal.bot
#  Discord PDF Converter Bot
# =============================================================================

# ── Base image: Python slim (Debian-based) ───────────────────────────────────
FROM python:3.12-slim

# ── Install system dependencies ───────────────────────────────────────────────
# ghostscript  : wajib untuk camelot-py (parsing PDF)
# libmagic1    : deteksi MIME type
# build-essential: compile beberapa paket Python (jika perlu)
RUN apt-get update && apt-get install -y --no-install-recommends \
        ghostscript \
        libmagic1 \
        curl \
        build-essential \
    && rm -rf /var/lib/apt/lists/*

# ── Install uv ────────────────────────────────────────────────────────────────
COPY --from=ghcr.io/astral-sh/uv:latest /uv /usr/local/bin/uv

# ── Set working directory ─────────────────────────────────────────────────────
WORKDIR /app

# ── Copy dependency files terlebih dahulu (layer caching) ────────────────────
COPY pyproject.toml uv.lock ./

# ── Install Python dependencies ───────────────────────────────────────────────
# --frozen     : gunakan versi persis dari uv.lock
# --no-dev     : skip dev dependencies
# --system     : install ke Python sistem (bukan .venv)
RUN uv sync --frozen --no-dev --system

# ── Copy source code ──────────────────────────────────────────────────────────
COPY converter.py discord_bot.py ./

# ── Buat folder output ────────────────────────────────────────────────────────
RUN mkdir -p /app/Hasil

# ── Non-root user untuk keamanan ─────────────────────────────────────────────
RUN useradd -m -u 1000 botuser && chown -R botuser:botuser /app
USER botuser

# ── Environment ───────────────────────────────────────────────────────────────
ENV PYTHONUNBUFFERED=1

# ── Entrypoint ────────────────────────────────────────────────────────────────
CMD ["python", "discord_bot.py"]
