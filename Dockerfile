FROM python:3.11-slim-bookworm

ENV PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1 \
    PLAYWRIGHT_BROWSERS_PATH=/home/user/.cache/ms-playwright \
    HOST=0.0.0.0 \
    PORT=7860

# HF Spaces require a non-root user with UID 1000.
RUN useradd -m -u 1000 user

# System libraries needed by Chromium (matches `playwright install-deps chromium`).
RUN apt-get update && apt-get install -y --no-install-recommends \
        ca-certificates \
        fonts-liberation \
        libasound2 \
        libatk-bridge2.0-0 \
        libatk1.0-0 \
        libatspi2.0-0 \
        libcairo2 \
        libcups2 \
        libdbus-1-3 \
        libdrm2 \
        libgbm1 \
        libglib2.0-0 \
        libnspr4 \
        libnss3 \
        libpango-1.0-0 \
        libx11-6 \
        libxcb1 \
        libxcomposite1 \
        libxdamage1 \
        libxext6 \
        libxfixes3 \
        libxkbcommon0 \
        libxrandr2 \
        wget \
    && rm -rf /var/lib/apt/lists/*

USER user
ENV HOME=/home/user PATH=/home/user/.local/bin:$PATH
WORKDIR /home/user/app

COPY --chown=user:user requirements.txt ./
RUN pip install --user --no-cache-dir -r requirements.txt

# Install only the Chromium browser (smaller than full install).
RUN python -m playwright install chromium

COPY --chown=user:user . .

RUN mkdir -p /home/user/app/downloads

EXPOSE 7860

CMD ["python", "app.py"]
