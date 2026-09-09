FROM python:3.11-slim

# Dependências do sistema: Tesseract OCR + pdftk + poppler
RUN apt-get update && apt-get install -y \
    tesseract-ocr \
    tesseract-ocr-por \
    pdftk \
    poppler-utils \
    curl \
    && rm -rf /var/lib/apt/lists/*

# Node.js 20
RUN curl -fsSL https://deb.nodesource.com/setup_20.x | bash - && \
    apt-get install -y nodejs && \
    rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Python deps
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Node deps
COPY package.json .
RUN npm install

# Código
COPY . .

EXPOSE 3002

CMD ["node", "server.js"]
