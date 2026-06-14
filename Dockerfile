FROM python:3.11-slim

# LibreOffice + poppler (pdf2image needs poppler-utils)
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    libreoffice-writer \
    libreoffice-impress \
    libreoffice-calc \
    poppler-utils \
    tesseract-ocr \
    tesseract-ocr-eng \
    tesseract-ocr-hin \
    ocrmypdf \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY main.py .

# Railway uses PORT env variable
ENV PORT=8000

EXPOSE 8000

CMD ["python", "main.py"]
