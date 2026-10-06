# Образ веб-сервиса переводчика. Собирается командой docker compose up -d --build (см. README).
FROM python:3.11-slim

RUN apt-get update \
 && apt-get install -y --no-install-recommends tesseract-ocr tesseract-ocr-hun \
 && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY src src
COPY server server
COPY manage.py translate_komplekt.py ./

ENV HUT_DATA=/app/data HUT_PROJECTS=/app/projects PYTHONUNBUFFERED=1 PYTHONIOENCODING=utf-8
EXPOSE 8010
HEALTHCHECK --interval=30s --timeout=5s CMD python -c "import urllib.request; urllib.request.urlopen('http://127.0.0.1:8010/health', timeout=4)"
CMD ["uvicorn", "server.app:app", "--host", "0.0.0.0", "--port", "8010", "--proxy-headers"]
