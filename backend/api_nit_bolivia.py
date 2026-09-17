"""
API REST (FastAPI) - Consulta Estado NIT Bolivia
Portal: https://siat.impuestos.gob.bo/rnc/public/consultas-estado-nit

Expone el scraper de consulta_nit_bolivia_playwright.py como una API HTTP.
Mantiene un único navegador Firefox (Playwright) abierto durante toda la
vida del proceso —cargar el portal toma ~10s por el chequeo SSO de
Keycloak— y reutiliza esa misma pestaña para cada consulta, serializadas
con un lock (el formulario no admite consultas concurrentes en paralelo).

Endpoints:
    GET /health       -> estado de la API y del navegador interno
    GET /nit/{numero} -> datos del contribuyente para ese NIT

Ejecutar:
    uvicorn api_nit_bolivia:app --host 0.0.0.0 --port 8000

Requisitos:
    pip install fastapi "uvicorn[standard]" playwright openpyxl
    python -m playwright install firefox
"""

import asyncio
import os
import re
from contextlib import asynccontextmanager

from fastapi import Depends, FastAPI, HTTPException, Path
from fastapi.responses import FileResponse
from fastapi.security import APIKeyHeader
from playwright.async_api import async_playwright
from pydantic import BaseModel

from consulta_nit_bolivia_playwright import URL, consultar_nit
from verificacion_segip import (
    ResultadoVerificacion,
    SegipNoConfigurado,
    esta_configurado as segip_configurado,
    verificar_identidad,
)

# ─── Autenticación por API key ────────────────────────────────────────────────
#
# La API mantiene un único navegador Firefox y serializa las consultas con un
# lock, así que es un recurso compartido y limitado: sin autenticación,
# cualquiera con la URL pública podría monopolizarlo, agotar la RAM del plan
# gratuito o provocar que el SIN bloquee la IP del servidor. Por eso se exige
# una clave en el header `X-API-Key`, comparada contra la variable de entorno
# `API_KEY`.
#
# El endpoint `/health` queda SIN protección a propósito: Railway y Render lo
# usan como healthcheck y no pueden enviar la clave.

API_KEY = os.environ.get("API_KEY", "").strip()
NOMBRE_HEADER_API_KEY = "X-API-Key"
_api_key_header = APIKeyHeader(name=NOMBRE_HEADER_API_KEY, auto_error=False)


async def verificar_api_key(clave: str = Depends(_api_key_header)):
    """Valida el header `X-API-Key` contra la variable de entorno `API_KEY`."""
    if not clave or clave != API_KEY:
        raise HTTPException(
            status_code=401,
            detail=f"API key inválida o ausente. Envíe el header '{NOMBRE_HEADER_API_KEY}'.",
        )


# ─── Estado del navegador (compartido durante la vida del proceso) ───────────

estado_navegador = {"playwright": None, "browser": None, "context": None, "page": None}
lock_consulta = asyncio.Lock()


RECURSOS_PESADOS = re.compile(
    r"\.(png|jpe?g|gif|svg|webp|woff2?|ttf|otf|mp4|webm)(\?.*)?$", re.IGNORECASE
)


async def _bloquear_recursos_pesados(route):
    await route.abort()


async def _nuevo_navegador():
    """Crea (o recrea) browser + contexto + página, y carga el portal.

    Se usa tanto al arrancar como para recuperarse de un crash: en planes
    gratuitos con poca RAM, Firefox puede morir a media consulta
    ("Page crashed" / "Target page... has been closed"), y la única forma
    de seguir sirviendo peticiones es descartar el navegador muerto y
    levantar uno nuevo desde cero (reusar la página fallida solo produce
    más errores en cascada).
    """
    pw = estado_navegador["playwright"]
    for clave in ("context", "browser"):
        objeto = estado_navegador.get(clave)
        if objeto is not None:
            try:
                await objeto.close()
            except Exception:
                pass

    # Firefox: en Chrome/Chromium el chequeo SSO de Keycloak nunca responde
    # y la SPA termina redirigiendo fuera de la página de consulta.
    browser = await pw.firefox.launch(
        headless=True,
        firefox_user_prefs={
            "browser.cache.disk.enable": False,
            "browser.cache.memory.enable": False,
            "browser.sessionhistory.max_total_viewers": 0,
            "media.autoplay.default": 5,
        },
    )
    context = await browser.new_context(locale="es-BO")
    page = await context.new_page()
    # Bloquear imágenes/fuentes/video reduce bastante el consumo de RAM de
    # Firefox; el formulario se ubica por selector/texto, no por apariencia.
    await page.route(RECURSOS_PESADOS, _bloquear_recursos_pesados)
    estado_navegador.update(browser=browser, context=context, page=page)

    await page.goto(URL, timeout=60_000)
    await page.wait_for_selector("input[id^='mat-input']", timeout=60_000)
    await page.wait_for_timeout(1000)
    return page


@asynccontextmanager
async def lifespan(app: FastAPI):
    # Sin API_KEY configurada la API quedaría abierta a internet: preferimos
    # fallar ruidosamente al arrancar antes que publicar un servicio sin
    # autenticación por descuido.
    if not API_KEY:
        raise RuntimeError(
            "La variable de entorno API_KEY no está configurada. "
            "Defínela con una clave secreta antes de iniciar la API "
            "(p. ej. API_KEY=una-clave-larga-y-aleatoria)."
        )

    pw = await async_playwright().start()
    estado_navegador["playwright"] = pw
    await _nuevo_navegador()
    yield

    for clave in ("context", "browser"):
        objeto = estado_navegador.get(clave)
        if objeto is not None:
            try:
                await objeto.close()
            except Exception:
                pass
    await pw.stop()
    estado_navegador.update(playwright=None, browser=None, context=None, page=None)


app = FastAPI(
    title="API Consulta NIT Bolivia",
    description="Consulta el estado de un NIT en el portal SIAT (RNC) del Servicio de Impuestos Nacionales de Bolivia.",
    version="1.0.0",
    lifespan=lifespan,
)


# ─── Modelos de respuesta ─────────────────────────────────────────────────────

class Contribuyente(BaseModel):
    nit: str
    consultado_en: str
    razon_social: str
    estado: str
    estado_actividad: str
    tipo_contribuyente: str
    regimen_contribuyente: str


class ErrorRespuesta(BaseModel):
    detail: str


class VerificacionCIRequest(BaseModel):
    ci: str
    fecha_nacimiento: str  # AAAA-MM-DD
    complemento: str | None = None


class VerificacionCIRespuesta(BaseModel):
    coincide: bool
    detalle: str


class SaludRespuesta(BaseModel):
    status: str
    navegador_listo: bool
    verificacion_ci_disponible: bool


# ─── Endpoints ────────────────────────────────────────────────────────────────

@app.get("/health", response_model=SaludRespuesta, summary="Verifica que la API y el navegador interno funcionan")
async def health():
    page = estado_navegador["page"]
    listo = page is not None and not page.is_closed()
    return SaludRespuesta(
        status="ok" if listo else "degradado",
        navegador_listo=listo,
        verificacion_ci_disponible=segip_configurado(),
    )


@app.get(
    "/nit/{numero}",
    response_model=Contribuyente,
    responses={
        401: {"model": ErrorRespuesta},
        404: {"model": ErrorRespuesta},
        502: {"model": ErrorRespuesta},
    },
    summary="Consulta el estado de un NIT",
    dependencies=[Depends(verificar_api_key)],
)
async def obtener_nit(
    numero: str = Path(..., pattern=r"^\d+$", description="Número de NIT a consultar (solo dígitos)")
):
    if estado_navegador["page"] is None:
        raise HTTPException(status_code=503, detail="El navegador interno aún no está listo")

    async with lock_consulta:
        try:
            resultado = await consultar_nit(estado_navegador["page"], numero)
            if resultado.get("estado", "").startswith("ERROR"):
                raise RuntimeError(resultado["estado"])
        except Exception:
            # El navegador pudo haberse caído (crash por memoria, sesión
            # vencida, "Target page... has been closed", etc.): se relanza
            # desde cero y se reintenta la consulta una sola vez.
            try:
                page = await _nuevo_navegador()
                resultado = await consultar_nit(page, numero)
            except Exception as e:
                raise HTTPException(
                    status_code=502,
                    detail=f"ERROR: no se pudo recuperar el navegador interno ({e})",
                )

    if "razon_social" in resultado:
        return Contribuyente(**resultado)

    detalle = resultado.get("estado", "No se pudo obtener información del NIT")
    if detalle.startswith("ERROR"):
        raise HTTPException(status_code=502, detail=detalle)
    raise HTTPException(status_code=404, detail=detalle)


@app.post(
    "/verificar-ci",
    response_model=VerificacionCIRespuesta,
    responses={
        401: {"model": ErrorRespuesta},
        501: {"model": ErrorRespuesta},
        502: {"model": ErrorRespuesta},
    },
    summary="Verifica una identidad contra SEGIP (requiere convenio)",
    dependencies=[Depends(verificar_api_key)],
)
async def verificar_ci(datos: VerificacionCIRequest):
    """Verifica si una CI + fecha de nacimiento corresponden a un registro válido.

    No devuelve datos personales: responde únicamente coincide / no coincide.
    Requiere un convenio activo con SEGIP (variables SEGIP_API_URL y
    SEGIP_API_TOKEN configuradas). Ver CONVENIO-SEGIP.md.
    """
    try:
        resultado: ResultadoVerificacion = await verificar_identidad(
            ci=datos.ci,
            fecha_nacimiento=datos.fecha_nacimiento,
            complemento=datos.complemento,
        )
    except SegipNoConfigurado as e:
        # 501 Not Implemented: el endpoint existe pero el convenio aún no está activo.
        raise HTTPException(status_code=501, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=502, detail=f"ERROR al verificar con SEGIP: {e}")

    return VerificacionCIRespuesta(coincide=resultado.coincide, detalle=resultado.detalle)


# ─── Frontend (SPA ebim) ──────────────────────────────────────────────────────
#
# Si existe un build del frontend de ebim (React/Vite) en FRONTEND_DIR, FastAPI
# lo sirve como archivos estáticos en la misma app: la UI del Estimador queda en
# la raíz y la API sigue en sus rutas (/nit, /verificar-ci, /health, /docs).
# Servir estáticos no agrega ningún proceso ni consumo de RAM relevante.
#
# Estas rutas se registran AL FINAL, así que nunca "tapan" a las de la API ni a
# /docs · /openapi.json: FastAPI evalúa las rutas en orden de registro y las
# explícitas (definidas antes) siempre ganan. El catch-all solo atiende lo que
# no matchea ninguna ruta de la API, devolviendo el archivo pedido o, si no
# existe, el index.html (para el ruteo del lado del cliente).

FRONTEND_DIR = os.path.abspath(os.environ.get("FRONTEND_DIR", "frontend_dist"))

if os.path.isdir(FRONTEND_DIR) and os.path.isfile(os.path.join(FRONTEND_DIR, "index.html")):
    _INDEX_HTML = os.path.join(FRONTEND_DIR, "index.html")

    @app.get("/", include_in_schema=False)
    async def _spa_root():
        return FileResponse(_INDEX_HTML)

    @app.get("/{ruta:path}", include_in_schema=False)
    async def _spa_catchall(ruta: str):
        # Sirve el archivo estático si existe (assets, favicon, etc.); si no,
        # devuelve index.html para que la SPA maneje la ruta en el navegador.
        # Se normaliza la ruta para impedir path traversal fuera de FRONTEND_DIR.
        candidato = os.path.normpath(os.path.join(FRONTEND_DIR, ruta))
        if candidato.startswith(FRONTEND_DIR + os.sep) and os.path.isfile(candidato):
            return FileResponse(candidato)
        return FileResponse(_INDEX_HTML)
