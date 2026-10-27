# Imagem base oficial do Python
FROM python:3.11-slim

# Evita .pyc no disco, habilita stdout sem buffer e define fuso/locale
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1 \
    DISPLAY=:0 \
    TZ=America/Sao_Paulo \
    LANG=C.UTF-8

# Dependências de sistema (somente runtime — todas as libs Python usadas possuem wheels):
# 1. WeasyPrint: Pango, HarfBuzz, Cairo, GDK-Pixbuf, fontes, shared-mime-info
# 2. CustomTkinter / Tkinter: Tcl/Tk + bibliotecas X11
# 3. ca-certificates: necessário para baixar as fontes do Google Fonts durante a geração do PDF
RUN apt-get update && apt-get install -y --no-install-recommends \
    libpango-1.0-0 \
    libpangocairo-1.0-0 \
    libpangoft2-1.0-0 \
    libharfbuzz0b \
    libcairo2 \
    libgdk-pixbuf-2.0-0 \
    shared-mime-info \
    fontconfig \
    fonts-dejavu-core \
    fonts-liberation \
    libtk8.6 \
    libtcl8.6 \
    libx11-6 \
    libxext6 \
    libxrender1 \
    libxft2 \
    ca-certificates \
    tzdata \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Instala as dependências Python primeiro (aproveita cache de camadas)
COPY requirements.txt .
RUN pip install --upgrade pip && pip install -r requirements.txt

# Copia o código-fonte (src/, Logo Revinn/, main.py...)
COPY . .

# Pasta padrão para salvar os relatórios gerados
RUN mkdir -p /output
VOLUME ["/output"]

CMD ["python", "main.py"]
