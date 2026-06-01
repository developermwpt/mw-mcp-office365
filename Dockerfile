FROM python:3.12-slim

WORKDIR /app

# Instalar curl para o healthcheck
RUN apt-get update && apt-get install -y --no-install-recommends curl && rm -rf /var/lib/apt/lists/*

# Copiar só o necessário para o install (aproveita layer cache)
COPY pyproject.toml ./
COPY src/ ./src/

# Install não-editável (sem caveats de .pth com caminhos com espaços)
RUN pip install --no-cache-dir .

# Pasta de dados para o SQLite (montada como volume)
RUN mkdir -p /data && chmod 700 /data

# Por defeito o servidor escuta em 0.0.0.0 dentro do container; o porto só é
# exposto para 127.0.0.1 no host via docker-compose (não fica público).
ENV BIND_HOST=0.0.0.0
ENV TOKEN_STORE_PATH=/data/tokens.db

EXPOSE 8000

HEALTHCHECK --interval=30s --timeout=5s --retries=3 \
    CMD curl -sf http://localhost:8000/healthz || exit 1

CMD ["mcp-o365"]
