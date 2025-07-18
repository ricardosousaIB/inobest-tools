# Usa uma imagem base Python oficial mais completa
FROM python:3.9-buster

# Instala dependências de sistema para PyMuPDF
# build-essential, libssl-dev, libffi-dev, python3-dev já devem estar na imagem buster
# mas libjpeg-dev e zlib1g-dev são específicos para a compilação de imagens/compressão
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
    libjpeg-dev \
    zlib1g-dev \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Instala PyMuPDF separadamente
# Isso pode ajudar a isolar problemas de instalação específicos do PyMuPDF
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
