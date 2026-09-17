# Backend — API de consulta NIT Bolivia

Este directorio contiene el **backend** que se despliega junto al frontend de
ebim en un mismo contenedor. Expone una API HTTP (FastAPI) que consulta el
**estado de un NIT** en el portal SIAT del Servicio de Impuestos Nacionales de
Bolivia, automatizando un navegador Firefox (Playwright).

En producción, el `Dockerfile` de la raíz del repo compila el frontend (Vite) y
levanta esta API en el mismo proceso: FastAPI sirve la SPA en `/` y la API en
sus rutas.

## Archivos

| Archivo | Qué es |
|---|---|
| `api_nit_bolivia.py` | App FastAPI: endpoints y servido del frontend estático |
| `consulta_nit_bolivia_playwright.py` | Motor de scraping (Playwright + Firefox) |
| `verificacion_segip.py` | Cliente de verificación de identidad SEGIP (esqueleto — requiere convenio) |
| `requirements.txt` | Dependencias Python |
| `CONVENIO-SEGIP.md` | Pasos para gestionar el convenio de verificación de CI con SEGIP |

## Endpoints

| Método | Ruta | Auth | Descripción |
|---|---|---|---|
| `GET` | `/health` | — | Estado de la API y del navegador interno |
| `GET` | `/nit/{numero}` | `X-API-Key` | Datos del contribuyente para ese NIT |
| `POST` | `/verificar-ci` | `X-API-Key` | Verifica identidad contra SEGIP (requiere convenio; si no, `501`) |

## Autenticación

Las consultas exigen el header `X-API-Key`, validado contra la variable de
entorno **`API_KEY`**. `/health` queda abierto (lo usan Railway/Render como
healthcheck). **La API se niega a arrancar si `API_KEY` no está configurada.**

```bash
curl -H "X-API-Key: tu-clave" https://<tu-app>/nit/555162024
```

## Variables de entorno

| Variable | Cuándo | Para qué |
|---|---|---|
| `API_KEY` | runtime | Clave del header `X-API-Key`; sin ella la API no arranca |
| `VITE_SUPABASE_URL` | build | URL de Supabase que Vite hornea en el frontend |
| `VITE_SUPABASE_ANON_KEY` | build | Anon key pública de Supabase para el frontend |
| `SEGIP_API_URL` / `SEGIP_API_TOKEN` | runtime (opcional) | Solo con convenio SEGIP activo (ver `CONVENIO-SEGIP.md`) |

## Correr solo el backend en local (sin Docker)

```bash
cd backend
pip install -r requirements.txt
python -m playwright install firefox
export API_KEY="una-clave-de-prueba"
uvicorn api_nit_bolivia:app --host 0.0.0.0 --port 8000
```

Sin un build del frontend presente, la API responde igual en sus rutas; solo no
sirve la SPA en `/`. En el contenedor de producción, el frontend sí se incluye.
