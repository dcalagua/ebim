# ═══════════════════════════════════════════════════════════════════════════
#  Imagen unificada: frontend ebim (React/Vite) + API NIT Bolivia (FastAPI)
#
#  Un solo contenedor sirve todo desde una única URL:
#    /                 → UI del Estimador ebim (React)
#    /nit/{numero}     → API de consulta de NIT (con X-API-Key)
#    /verificar-ci     → verificación de identidad (requiere convenio SEGIP)
#    /health · /docs
#
#  Etapa 1 (Node): compila el frontend de ESTE repo (`npm run build` → dist/).
#  Etapa 2 (Python): API de NIT (FastAPI + Firefox de Playwright) que además
#  sirve ese dist/ como archivos estáticos. El backend real de ebim (auth, base
#  de datos, IA y email) sigue corriendo en Supabase, no aquí.
# ═══════════════════════════════════════════════════════════════════════════

# ─── Etapa 1: build del frontend ─────────────────────────────────────────────
# Node 22: Vite 8 requiere Node >= 20.19 / >= 22.12.
FROM node:22-slim AS frontend

# Variables públicas de Supabase que Vite "hornea" en el bundle en tiempo de
# build. En Railway/Render se definen como variables del servicio y se pasan
# como build args automáticamente. La anon key es pública de frontend por diseño.
ARG VITE_SUPABASE_URL
ARG VITE_SUPABASE_ANON_KEY
ENV VITE_SUPABASE_URL=$VITE_SUPABASE_URL \
    VITE_SUPABASE_ANON_KEY=$VITE_SUPABASE_ANON_KEY

WORKDIR /app
COPY package.json package-lock.json ./
RUN npm ci
COPY . .
RUN npm run build
# Resultado: /app/dist

# ─── Etapa 2: API de NIT + servido del frontend ──────────────────────────────
# Imagen base ligera de Python sobre Debian: Firefox de Playwright requiere
# bibliotecas del sistema (las instala `playwright install --with-deps`).
FROM python:3.12-slim

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PIP_NO_CACHE_DIR=1 \
    PLAYWRIGHT_BROWSERS_PATH=/ms-playwright \
    FRONTEND_DIR=/app/frontend_dist

WORKDIR /app

COPY backend/requirements.txt .
RUN pip install -r requirements.txt \
    && playwright install --with-deps firefox

# Código del backend (API NIT + scraper + verificación SEGIP).
COPY backend/ ./

# Build estático del frontend producido en la etapa 1.
COPY --from=frontend /app/dist /app/frontend_dist

EXPOSE 8000

# Railway y Render inyectan el puerto en $PORT; localmente cae a 8000.
CMD ["sh", "-c", "uvicorn api_nit_bolivia:app --host 0.0.0.0 --port ${PORT:-8000}"]
