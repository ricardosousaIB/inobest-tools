# Usa uma imagem base Python oficial (mais genérica que slim-buster)
FROM python:3.9-slim

# Instala apenas as dependências de sistema ABSOLUTAMENTE mínimas para PyMuPDF
# As rodas (wheels) do PyMuPDF geralmente não precisam de build-essential
# mas podem precisar de libjpeg-dev e zlib1g-dev para funcionalidades de imagem/compressão
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
    libjpeg62-turbo-dev \
    zlib1g-dev \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Instala PyMuPDF separadamente (sempre bom para isolar)
RUN pip install --no-cache-dir PyMuPDF

# Define o diretório de trabalho dentro do container
WORKDIR /app

# Copia o ficheiro de requisitos (agora SEM PyMuPDF) e instala as restantes dependências
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copia o resto do código da aplicação para o diretório de trabalho
COPY . .

# Expõe a porta que o Streamlit usa (padrão é 8501)
EXPOSE 8501

# Comando para executar a aplicação Streamlit quando o container iniciar
CMD ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]
