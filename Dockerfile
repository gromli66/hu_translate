# Образ веб-сервиса переводчика. Собирается командой docker compose up -d --build (см. README).
FROM python:3.11-slim

RUN apt-get update \
 && apt-get install -y --no-install-recommends tesseract-ocr tesseract-ocr-hun tesseract-ocr-rus tesseract-ocr-eng \
 && rm -rf /var/lib/apt/lists/*

# PDF перевода печатается из готового DOCX (LibreOffice без окна, оглавление и поля — через UNO системного python3).
# Шрифты: Tahoma — метрическая копия из fonts-wine, Times New Roman — Liberation Serif, Consolas/Verdana — DejaVu,
# Open Sans — свой (руководства OAH); прочие неустановленные — по классу (docker/fonts.conf)
RUN apt-get update \
 && apt-get install -y --no-install-recommends libreoffice-writer-nogui python3-uno \
    fonts-wine fonts-liberation fonts-dejavu-core fonts-crosextra-carlito fonts-crosextra-caladea fonts-opensymbol \
    fonts-open-sans \
 && rm -rf /var/lib/apt/lists/* \
 && mkdir -p /usr/share/fonts/truetype/wine \
 && ln -s /usr/share/wine/fonts/tahoma.ttf /usr/share/fonts/truetype/wine/
COPY docker/fonts.conf /etc/fonts/local.conf
RUN fc-cache -f

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
