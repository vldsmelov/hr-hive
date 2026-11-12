FROM python:3.13.12-slim-bookworm@sha256:a58daefb915e1e03ad48f3ca4df8832065412c5c35cacb9d39f4229184de12b6
ENV PYTHONUNBUFFERED=1 PYTHONUTF8=1 PYTHONIOENCODING=utf-8 LANG=C.UTF-8 TZ=Europe/Moscow HR_DATA=/data
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-writer libreoffice-impress libreoffice-calc tesseract-ocr tesseract-ocr-rus \
    tesseract-ocr-eng fonts-dejavu-core fonts-liberation2 fontconfig \
    && rm -rf /var/lib/apt/lists/*
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY app ./app
COPY rules ./rules
COPY tools ./tools
ENTRYPOINT ["python", "-m", "app.cli"]
CMD ["status"]
