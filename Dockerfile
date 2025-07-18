# Usa uma imagem base Python oficial mais completa
FROM python:3.9-buster

# Instala dependências de sistema para PyMuPDF (e outras bibliotecas comuns)
# A imagem 'buster' já deve ter muitas destas, mas vamos garantir as essenciais
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
    libjpeg-dev \
    zlib1g-dev \
    # build-essential, libssl-dev, libffi-dev, python3-dev já devem estar na imagem buster
    # mas podemos deixá-los para garantir, ou remover se o build for bem-sucedido sem eles
    && apt-get clean \
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
