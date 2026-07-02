FROM python:3.9-slim

WORKDIR /app

# Instala dependências do sistema necessárias para o Streamlit
RUN apt-get update && apt-get install -y \
    build-essential \
    curl \
    && rm -rf /var/lib/apt/lists/*

# Copia os arquivos do projeto
COPY . .

# Instala as bibliotecas do Python
RUN pip3 install --no-cache-dir -r requirements.txt

EXPOSE 8501

# Comando para rodar o seu app
CMD ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]
