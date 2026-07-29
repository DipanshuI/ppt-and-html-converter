FROM python:3.12-slim

# Install LibreOffice Headless and Poppler Utilities
RUN apt-get update && apt-get install -y \
    libreoffice \
    poppler-utils \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 8095

CMD ["python", "app.py"]
