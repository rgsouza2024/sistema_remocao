# Dockerfile
# Container oficial para a API de Remoção

# Imagem base leve (Python 3.12)
FROM python:3.12-slim

# Define diretório de trabalho
WORKDIR /app

# Variáveis de ambiente
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

# Instala dependências do sistema (se necessárias)
# RUN apt-get update && apt-get install -y gcc ...

# Copia e instala dependências Python
COPY requirements_api.txt .
RUN pip install --no-cache-dir -r requirements_api.txt

# Copia o código da aplicação
COPY api.py .
COPY motor_remocao.py .

# Expõe a porta 8000
EXPOSE 8000

# Comando de inicialização
CMD ["uvicorn", "api:app", "--host", "0.0.0.0", "--port", "8000"]
