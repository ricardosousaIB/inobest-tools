# Usa uma imagem base Python oficial
FROM python:3.9-slim-buster

# Instala dependências de sistema para PyMuPDF (e outras bibliotecas comuns)
RUN apt-get update && apt-get install -y \
    build-essential \
    libssl-dev \
    libffi-dev \
    python3-dev \
    # Dependências específicas para MuPDF/fitz
    libopenjp2-7 \
    libjpeg-dev \
    zlib1g-dev \
    # Limpeza
    && rm -rf /var/lib/apt/lists/*

# Define o diretório de trabalho dentro do container
WORKDIR /app

# Copia o ficheiro de requisitos e instala as dependências
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copia o resto do código da aplicação para o diretório de trabalho
COPY . .

# Expõe a porta que o Streamlit usa (padrão é 8501)
EXPOSE 8501

# Comando para executar a aplicação Streamlit quando o container iniciar
CMD ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]
