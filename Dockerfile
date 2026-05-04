FROM python:3.12-alpine3.22 AS builder

WORKDIR /build

# Upgrade base packages and install build deps for binary wheels
RUN apk upgrade --no-cache && \
    apk add --no-cache --virtual .build-deps \
      gcc g++ musl-dev libxml2-dev libxslt-dev

COPY requirements.txt ./
RUN pip install --upgrade pip setuptools wheel && \
    pip wheel --no-cache-dir --wheel-dir /wheels -r requirements.txt


FROM python:3.12-alpine3.22

WORKDIR /app

# Keep runtime image small while applying latest security patches
RUN apk upgrade --no-cache && \
    apk add --no-cache libxml2 libxslt && \
    addgroup -S app && adduser -S -G app app

COPY requirements.txt ./
COPY --from=builder /wheels /wheels
RUN pip install --no-cache-dir --no-index --find-links=/wheels -r requirements.txt && \
    rm -rf /wheels

COPY app/ ./app/
RUN chown -R app:app /app
USER app

EXPOSE 8501

HEALTHCHECK --interval=30s --timeout=10s --start-period=10s --retries=3 \
    CMD python -c "import requests; response = requests.get('http://localhost:8501/_stcore/health', timeout=5); response.raise_for_status()"

ENTRYPOINT ["streamlit", "run", "app/main.py", "--server.port=8501", "--server.address=0.0.0.0"]
