FROM python:3.12-slim

RUN apt-get update && \
    apt-get install -y --no-install-recommends pandoc fonts-noto-cjk && \
    rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY pyproject.toml README.md ./
COPY src/ src/

RUN pip install --no-cache-dir ".[web]"

EXPOSE 8080

ENTRYPOINT ["marp-pptx"]
CMD ["serve", "--host", "0.0.0.0", "--port", "8080"]
