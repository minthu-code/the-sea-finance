FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

WORKDIR /app

RUN apt-get update && apt-get install -y \
    tesseract-ocr \
    libtesseract-dev \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

RUN mkdir -p /data /data/exports

ENV DB_PATH=/data/exhibitledger.db
ENV EXPORT_DIR=/data/exports
ENV DEFAULT_EXHIBITION=SHWEDAGON2024

EXPOSE 10000

CMD ["python", "main.py"]
