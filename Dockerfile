# syntax=docker/dockerfile:1.7

FROM node:20-alpine AS frontend-builder
WORKDIR /frontend
ARG APP_VERSION=2.0.4
COPY frontend/package.json frontend/pnpm-lock.yaml ./
RUN corepack enable && corepack prepare pnpm@10.17.1 --activate && pnpm install --frozen-lockfile --ignore-scripts
COPY frontend ./
ENV NEXT_PUBLIC_API_BASE=""
ENV NEXT_PUBLIC_APP_VERSION=${APP_VERSION}
RUN pnpm build

FROM python:3.12-slim AS runtime
WORKDIR /app
ARG APP_VERSION=2.0.4
ENV APP_VERSION=${APP_VERSION}
ENV NEXT_PUBLIC_APP_VERSION=${APP_VERSION}

COPY backend/requirements.txt ./requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

COPY backend/app ./app
COPY --from=frontend-builder /frontend/out ./frontend_dist

EXPOSE 8000
CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8000"]
