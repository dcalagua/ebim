# Convenio con SEGIP para verificación de identidad (CI)

Esta guía explica **por qué** la consulta por cédula de identidad (CI) no es un
scraping como el de NIT, y **qué pasos** seguir para habilitar la verificación
de identidad por la vía oficial y legal: un convenio con el **SEGIP** (Servicio
General de Identificación Personal).

> Resumen en una línea: no existe una fuente pública para "ingresar una CI y
> traer los datos personales de alguien". La única vía legítima es un servicio
> de **verificación** (coincide / no coincide) habilitado por SEGIP mediante
> convenio.

---

## 1. Por qué hace falta un convenio (y no se puede scrapear)

- **El NIT es un registro público** pensado para ser consultado por cualquiera
  → por eso el scraping del portal del SIN es legítimo.
- **La CI y los datos personales NO son públicos.** SEGIP no expone ningún
  formulario donde, metiendo una CI cualquiera, devuelva nombre, dirección,
  fecha de nacimiento, etc. No es que sea difícil de scrapear: **no existe esa
  pantalla pública**.
- El acceso a la información de identidad está **reservado a instituciones con
  convenio**, respaldado por la **Ley N° 145** del Servicio General de
  Identificación Personal.

## 2. Qué servicio se solicita

El servicio relevante es la **verificación de identidad** de SEGIP (familia
"verID"). En lugar de devolver datos personales, valida datos que tu institución
**ya posee con consentimiento del titular**:

- Número de **CI**
- **Complemento** (cuando aplica)
- **Fecha de nacimiento**

…y responde **coincide / no coincide**. Esto es lo que sirve para KYC /
onboarding sin violar la privacidad del ciudadano.

## 3. ⚠️ Restricciones de uso (leer antes de invertir tiempo)

Según las propias condiciones de SEGIP, el acceso a sus servicios de identidad:

- **No puede usarse con fines de lucro.**
- **No puede tercerizarse** ni revenderse a terceros (ni siquiera entre
  instancias de la misma institución).
- **No puede almacenarse** en servidores o servicios externos al centro de
  datos nacional de SEGIP.

**Implicación práctica para tu caso:** si tu desarrollo es un producto comercial
que **revende** verificación de identidad a clientes, probablemente **no
calificás** para un convenio en esos términos, o el convenio deberá estructurarse
de otra forma (p. ej. siendo la institución usuaria final quien firma, no el
proveedor de software). Conviene **plantear el modelo de negocio explícitamente
a SEGIP** desde el primer contacto para no avanzar en una dirección inviable.

## 4. Pasos para gestionar el convenio

> Estos son los pasos generales del proceso interinstitucional boliviano. El
> detalle fino (formularios, instancia exacta) lo confirma SEGIP en el primer
> contacto.

1. **Definir la entidad solicitante y la base legal.** SEGIP firma convenios con
   instituciones que tengan respaldo legal para validar identidad (entidades
   públicas, financieras reguladas, etc.). Identificá bajo qué figura entra tu
   organización.
2. **Contacto inicial con SEGIP** (ver sección 6) para solicitar información del
   servicio de verificación y los requisitos vigentes. Planteá tu caso de uso y
   modelo de negocio desde el inicio (ver restricciones, sección 3).
3. **Carta/solicitud formal de convenio** dirigida a la Dirección Ejecutiva de
   SEGIP, describiendo: institución, finalidad del uso, volumen estimado de
   consultas, medidas de seguridad y tratamiento de datos.
4. **Evaluación de SEGIP** de la solicitud y los respaldos legales.
5. **Firma del convenio interinstitucional**, que define alcance, obligaciones,
   confidencialidad y condiciones técnicas.
6. **Habilitación técnica:** SEGIP entrega los datos de integración —URL del
   servicio, credenciales/token y el contrato de la API (formato de request y
   response)—.
7. **Integración y pruebas** contra el ambiente que indique SEGIP.

## 5. Qué hacer cuando tengas el convenio (integración técnica)

Este repositorio ya deja el **punto de integración listo**:

- Módulo: [`verificacion_segip.py`](verificacion_segip.py) — cliente de
  verificación (esqueleto). Mientras no esté configurado, no realiza ninguna
  consulta.
- Endpoint: `POST /verificar-ci` en [`api_nit_bolivia.py`](api_nit_bolivia.py),
  protegido con la misma API key (`X-API-Key`).

Para activarlo, una vez firmado el convenio:

1. Configurá las variables de entorno con lo que entregue SEGIP:
   ```bash
   export SEGIP_API_URL="https://<endpoint-que-provea-segip>"
   export SEGIP_API_TOKEN="<credencial-entregada-por-segip>"
   ```
   (En Railway/Render se cargan en el panel de variables, igual que `API_KEY`.)
2. Ajustá en `verificacion_segip.py` los `TODO(convenio)`: el formato exacto del
   request, las cabeceras de autenticación y el parseo de la respuesta, según el
   contrato real que defina SEGIP (el código actual es un ejemplo ilustrativo).
3. Verificá el estado en `GET /health`:
   ```json
   { "status": "ok", "navegador_listo": true, "verificacion_ci_disponible": true }
   ```
   `verificacion_ci_disponible: true` indica que las credenciales están cargadas.

Comportamiento del endpoint según el estado del convenio:

| Situación | Respuesta de `POST /verificar-ci` |
|---|---|
| Sin convenio configurado | `501` — `"Verificación de CI no disponible: falta el convenio con SEGIP..."` |
| Falta/!nválida la API key | `401` |
| Convenio activo, datos verificados | `200` — `{ "coincide": true, "detalle": "Coincide" }` |
| Error al contactar a SEGIP | `502` |

Ejemplo de llamada (con convenio activo):
```bash
curl -X POST https://<tu-servicio>.up.railway.app/verificar-ci \
  -H "X-API-Key: <tu-api-key>" \
  -H "Content-Type: application/json" \
  -d '{"ci":"1234567","fecha_nacimiento":"1990-01-01","complemento":null}'
```

## 6. Contacto SEGIP

- **Sitio oficial:** https://www.segip.gob.bo/
- **Línea nacional de consultas:** 800 10 11 02
- **WhatsApp (orientación):** 67198896
- **Oficinas centrales:** La Paz, Bolivia (dirección y horarios en el sitio).
- **Marco legal:** Ley N° 145 del Servicio General de Identificación Personal.

> Recomendación: hacé el primer contacto por la línea/WhatsApp para que te
> indiquen la unidad y el procedimiento vigente de convenios, y pedí por escrito
> los requisitos actualizados antes de redactar la solicitud formal.

---

## Nota legal

Bolivia tiene en trámite un **anteproyecto de Ley de Protección de Datos
Personales** (AGETIC, 2024–2025), y la Constitución ya protege la privacidad
(art. 130, acción de protección de privacidad / habeas data). Tratar datos
personales por fuera de un convenio oficial — por ejemplo, scrapeando o
almacenando datos de identidad — implica riesgo legal. La vía del convenio no es
solo la correcta: es la única sostenible.
