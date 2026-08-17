# Imagem base oficial do Python
FROM python:3.11-slim

# Evita que o Python grave arquivos .pyc no disco e habilita stdout sem buffer
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1
ENV DISPLAY=:0

# Instala dependências do sistema operacional necessárias para:
# 1. WeasyPrint (Pango, Cairo, GDK-Pixbuf, Fontes)
# 2. CustomTkinter / Tkinter (python3-tk, X11 libraries, Xvfb para execução headless se necessário)
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    python3-dev \
    python3-tk \
    tk-dev \
    tcl-dev \
    libpangocairo-1.0-0 \
    libcairo2 \
    libgdk-pixbuf-2.0-0 \
    libffi-dev \
    shared-mime-info \
    fontconfig \
    fonts-dejavu-core \
    fonts-liberation \
    libx11-6 \
    libxext6 \
    libxrender1 \
    xvfb \
    && rm -rf /var/lib/apt/lists/*

# Define o diretório de trabalho no container
WORKDIR /app

# Copia e instala as dependências do Python
COPY requirements.txt .
RUN pip install --no-cache-dir --upgrade pip && \
    pip install --no-cache-dir -r requirements.txt

# Copia o código-fonte da aplicação
COPY . .

# Comando padrão para iniciar a aplicação
CMD ["python", "main.py"]
