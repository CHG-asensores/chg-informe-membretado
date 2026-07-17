# ── Base ──────────────────────────────────────────────────────────────
FROM python:3.11-slim

# Dependencias de sistema para Playwright / Chromium
RUN apt-get update && apt-get install -y \
    wget curl gnupg ca-certificates \
    libglib2.0-0 libnss3 libnspr4 libatk1.0-0 libatk-bridge2.0-0 \
    libcups2 libdrm2 libdbus-1-3 libxcb1 libxkbcommon0 libx11-6 \
    libxcomposite1 libxdamage1 libxext6 libxfixes3 libxrandr2 \
    libgbm1 libpango-1.0-0 libcairo2 libasound2 libatspi2.0-0 \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Instalar dependencias Python
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Instalar Chromium para Playwright
RUN playwright install chromium

# Copiar código fuente
COPY . .

# Puerto expuesto (Railway usa $PORT)
EXPOSE 5000

# Gunicorn: 1 worker (Playwright no es thread-safe).
# Timeout 300s: leer la OT por API implica ~24 descargas de fotos desde S3
# (antes venian dentro del PDF, cero red). Van en paralelo, pero si el
# operador arrastra varios PDFs de golpe, 120s se quedaban cortos.
CMD gunicorn app:app --workers 1 --timeout 300 --bind 0.0.0.0:${PORT:-5000}
