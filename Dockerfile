# =============================================================================
#  Dockerfile — UltraJurnal.bot
#  Discord PDF Converter Bot
# =============================================================================

# ── Base image: Python slim (Debian-based) ───────────────────────────────────
FROM python:3.12-slim

# ── Install system dependencies ───────────────────────────────────────────────
# ghostscript   : wajib untuk camelot-py (parsing PDF)
# libmagic1     : deteksi MIME type (python-magic)
# libmagic-dev  : header untuk compile python-magic
# libgomp1      : OpenMP — dibutuhkan numpy/pandas di beberapa arch
# build-essential: compile ekstensi C Python jika diperlukan
RUN apt-get update && apt-get install -y --no-install-recommends \
        ghostscript \
        libmagic1 \
        libmagic-dev \
        libgomp1 \
        curl \
        build-essential \
    && rm -rf /var/lib/apt/lists/*

# ── Install uv (pin versi untuk stabilitas) ──────────────────────────────────
COPY --from=ghcr.io/astral-sh/uv:0.7.8 /uv /usr/local/bin/uv

# ── Set working directory ─────────────────────────────────────────────────────
WORKDIR /app

# ── Copy dependency files terlebih dahulu (layer caching) ────────────────────
COPY pyproject.toml uv.lock ./

# ── Install Python dependencies ke virtual environment ───────────────────────
# Gunakan venv eksplisit (/app/.venv) agar lebih portable di Docker.
# --frozen     : gunakan versi persis dari uv.lock (no resolution)
# --no-dev     : skip dev dependencies
ENV VIRTUAL_ENV=/app/.venv
RUN uv venv $VIRTUAL_ENV && \
    uv sync --frozen --no-dev

# ── Tambahkan .venv/bin ke PATH ──────────────────────────────────────────────
ENV PATH="$VIRTUAL_ENV/bin:$PATH"

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
