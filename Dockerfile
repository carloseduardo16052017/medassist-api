FROM python:3.11-slim

WORKDIR /app

# Instalar dependências do sistema + LibreOffice para conversão PDF
RUN apt-get update && apt-get install -y --no-install-recommends \
    gcc \
    libreoffice \
    libreoffice-writer \
    fonts-liberation \
    fonts-dejavu \
    && rm -rf /var/lib/apt/lists/*

# Copiar e instalar dependências Python
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copiar código
COPY . .

# Railway injeta PORT automaticamente
ENV PORT=8000
ENV HOME=/tmp

EXPOSE $PORT

CMD uvicorn main:app --host 0.0.0.0 --port $PORT
