"""
Cliente de verificación de identidad contra SEGIP — ESQUELETO / punto de integración.

⚠️  IMPORTANTE — leer antes de usar:

SEGIP **no** ofrece ningún endpoint público para obtener los datos personales
de alguien a partir de su cédula de identidad (CI). El acceso a su servicio de
verificación de identidad está reservado a instituciones que firmaron un
**convenio**, respaldado por la Ley N° 145 del Servicio General de
Identificación Personal. Sin ese convenio y sus credenciales oficiales, este
módulo NO realiza ninguna consulta: queda como punto de integración listo para
conectar el día que el servicio esté habilitado.

Los pasos para gestionar el convenio están en `CONVENIO-SEGIP.md`.

Modelo de uso correcto — **VERIFICACIÓN, no extracción**:
    No se trata de "dame los datos de esta CI", sino de "¿estos datos que ya
    tengo, con consentimiento del titular, corresponden a un registro válido?".
    Se envían datos que la institución ya posee (CI + complemento + fecha de
    nacimiento) y SEGIP responde coincide / no coincide. No se obtienen ni se
    almacenan datos personales nuevos.

Configuración (variables de entorno, una vez firmado el convenio):
    SEGIP_API_URL    -> URL del endpoint de verificación que provea SEGIP
    SEGIP_API_TOKEN  -> credencial/token entregado por SEGIP en el convenio

El formato exacto del request/response lo define SEGIP en el convenio; las
funciones de abajo dejan marcado con TODO el punto donde adaptarlo.
"""

import os
from dataclasses import dataclass
from typing import Optional

SEGIP_API_URL = os.environ.get("SEGIP_API_URL", "").strip()
SEGIP_API_TOKEN = os.environ.get("SEGIP_API_TOKEN", "").strip()


class SegipNoConfigurado(RuntimeError):
    """Se intentó verificar sin convenio/credenciales de SEGIP configuradas."""


@dataclass
class ResultadoVerificacion:
    coincide: bool
    detalle: str


def esta_configurado() -> bool:
    """True solo si hay URL y token de SEGIP configurados (convenio activo)."""
    return bool(SEGIP_API_URL and SEGIP_API_TOKEN)


async def verificar_identidad(
    ci: str,
    fecha_nacimiento: str,
    complemento: Optional[str] = None,
) -> ResultadoVerificacion:
    """Verifica una identidad contra el servicio oficial de SEGIP.

    Args:
        ci: Número de cédula de identidad (solo dígitos).
        fecha_nacimiento: Fecha de nacimiento del titular (formato AAAA-MM-DD).
        complemento: Complemento alfanumérico de la CI, si aplica.

    Returns:
        ResultadoVerificacion con `coincide` (bool) y un `detalle` informativo.

    Raises:
        SegipNoConfigurado: si todavía no hay convenio/credenciales configuradas.
    """
    if not esta_configurado():
        raise SegipNoConfigurado(
            "Verificación de CI no disponible: falta el convenio con SEGIP. "
            "Configura las variables SEGIP_API_URL y SEGIP_API_TOKEN una vez "
            "habilitado el servicio oficial. Ver CONVENIO-SEGIP.md."
        )

    # ── Punto de integración real ────────────────────────────────────────────
    # TODO(convenio): ajustar endpoint, payload, cabeceras y parseo de la
    # respuesta al contrato que defina SEGIP. Lo de abajo es un EJEMPLO
    # ilustrativo, no el formato definitivo.
    import httpx  # import diferido: solo se necesita si el convenio está activo

    async with httpx.AsyncClient(timeout=30) as client:
        respuesta = await client.post(
            SEGIP_API_URL,
            headers={"Authorization": f"Bearer {SEGIP_API_TOKEN}"},
            json={
                "ci": ci,
                "complemento": complemento,
                "fecha_nacimiento": fecha_nacimiento,
            },
        )
        respuesta.raise_for_status()
        datos = respuesta.json()

    # TODO(convenio): mapear al esquema real de respuesta de SEGIP.
    coincide = bool(datos.get("coincide"))
    detalle = datos.get("detalle", "Coincide" if coincide else "No coincide")
    return ResultadoVerificacion(coincide=coincide, detalle=detalle)
