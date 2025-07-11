# Use uma imagem base Python oficial
FROM python:3.9-slim-buster

# Define o diretório de trabalho dentro do contêiner
WORKDIR /app

# Copia o ficheiro de requisitos e instala as dependências
# Usamos COPY --chown para garantir que o utilizador 'appuser' (se criado) tenha permissões
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copia o restante do código da aplicação para o diretório de trabalho
# O '.' no final significa copiar tudo do diretório atual do host para /app no contêiner
COPY . .

# Expõe a porta padrão do Streamlit
EXPOSE 8501

# Comando para rodar a aplicação Streamlit
# Certifique-se de que 'app.py' é o nome do seu ficheiro principal Streamlit
CMD ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]
