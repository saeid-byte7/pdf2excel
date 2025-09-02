# Base Python image (fungerar på Apple Silicon & Intel)
FROM python:3.11-slim

# System dependencies: poppler, ghostscript, tesseract OCR, ocrmypdf (+ språkpaket eng/swe)
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    poppler-utils \
    ghostscript \
    tesseract-ocr \
    tesseract-ocr-eng \
    tesseract-ocr-swe \
    ocrmypdf \
 && rm -rf /var/lib/apt/lists/*

# Snabbare, renare Python-körning
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

# Arbetskatalog
WORKDIR /app

# Installera Python-paket först (bättre layer-cache)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Kopiera in koden
COPY . .

# Kör appen på port 8080 (matcha med -p vid 'docker run')
ENV PORT=8080
CMD ["uvicorn","app:app","--host","0.0.0.0","--port","8080","--workers","2","--timeout-keep-alive","5"]
