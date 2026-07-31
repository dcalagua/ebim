import React, { useState, useMemo, useEffect, useRef } from "react";
import ExcelJS from "exceljs";
import {
  Document, Packer, Paragraph, TextRun, AlignmentType,
  Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, VerticalAlign,
} from "docx";
import mammoth from "mammoth";
import {
  Sparkles, Plus, Trash2, Save, Download, History, FileSpreadsheet,
  AlertTriangle, CheckCircle2, RotateCcw, Loader2, X, Paperclip,
  FileText, FileType, Image as ImageIcon, LogOut, Mail, Target, SlidersHorizontal, ChevronDown
} from "lucide-react";
import { supabase } from "./lib/supabaseClient";
import { store } from "./lib/store";
import ebimLogoCompleto from "./assets/brand/ebim-logo-completo.png";

/* =========================================================================
   EBIM — Estimador de Costos de Consultoría TI
   Moneda de cotización: USD. Tarifas referenciales por país (editables).
   ========================================================================= */

const EBIM_PROFILE = `GRUPO EBIM SAC (RUC 20602517986, San Isidro - Lima, Perú; presencia en Perú, Ecuador y Bolivia; ~11-50 personas; +7-10 años, ISO 9001:2015). Consultora de TI — IMPORTANTE sobre SAP: EBIM NACIÓ dando soporte SAP R/3 a nivel TÉCNICO (ABAP/BASIS) y FUNCIONAL (SD MM FI CO PP PM QM PS HR), y ese sigue siendo su fuerte principal — pero vía AMS (soporte y mejora continua a sistemas SAP ya implementados), Staffing (consultores dedicados) y Bolsa de Horas (paquetes de horas prepagadas), NO vía implementaciones greenfield/rollouts nuevos de SAP. Además es implementador de Odoo ERP (línea activa) y tiene foco fuerte en desarrollo de IA/agentes. Desarrolla y vende software propio como SaaS (no se cotiza en este estimador, es otro modelo de negocio): eSUPPLIER (portal de proveedores con IA), GMAO (mantenimiento/EAM), eExpense (portal de rendiciones y gastos). Servicios que SÍ se cotizan aquí: soporte/AMS/staffing/bolsa de horas SAP (técnico y funcional), implementación Odoo ERP, desarrollo de software web/móvil a medida, IA/agentes/chatbots, arquitectura cloud (AWS/Azure/GCP), BI/QLIK, IoT, outsourcing de personal TI, mesa de ayuda 24/7. Metodologías ágiles. Perfiles típicos: Arquitecto TI/Consultor Senior (funcional o técnico ABAP/BASIS), Consultor Semi-Senior, Analista de Procesos/Capacidades TI y PM/Coordinador.`;

// Inteligencia de mercado curada por Lucho (Analista de Pricing) — referencia ESTABLE de competencia
// PE/EC, para que "análisis de competencia" no dependa de un web_search distinto cada corrida (eso
// era una fuente real de inconsistencia de precio). Actualizar cuando Lucho entregue un nuevo reporte.
const MARKET_INTEL = `INTELIGENCIA DE MERCADO PE/EC (investigada por Lucho, Analista de Pricing EBIM — 2026-06-30, usar como referencia fija, NO la reinvestigues con web_search):

PERÚ:
- Soporte SAP técnico+funcional / AMS / Staffing / Bolsa de horas (NO implementaciones nuevas): competidores confirmados por EBIM (conocimiento directo de mercado, no solo búsqueda) — CSTI Corp (17 años, SAP Gold Partner, ~150 consultores, fuerte en staffing/outsourcing/support/projects, PE+EC+CO+CR+US), Omnia Solution (SAP Gold Partner, 25 años, adquirida por EPI-USE/GroupElephant — ahora respaldada por un grupo global, ojo con esto), AYESA (SAP Gold Partner, Centro de Excelencia SAP de 750 profesionales, proyectos grandes como Enel/Cálidda — jugador GRANDE, no boutique), Gestión y Sistemas, Innoteam, Altamira Technology, Medialab LA, Novis. EBIM nació en este nicho (SAP R/3 desde el origen) y compite ahí, no contra integradores globales tipo IBM/Accenture/NTT — aunque AYESA y Omnia (ya con EPI-USE) empiezan a jugar en esa escala mayor, hay que diferenciarse por agilidad/cercanía frente a ellos, no por tamaño.
- Odoo ERP: IT Grupo (líder, 30+ profesionales certificados), Ganemo, OZ Solutions, E&M Sistemas (localización SUNAT). Precios referenciales: implementaciones básicas USD 900-3,000, proyectos medianos/grandes USD 10,000-50,000.
- IA aplicada/agentes: boutique ALEF AI Solutions, Miss Yera Consulting, IAConsultor.pe, Automaxia, NEO Consulting. 34% de empresas peruanas ya usan IA/automatización, inversión proyectada 3.9x en 2026 (la más alta de LatAm). Sin líder consolidado.

ECUADOR:
- Soporte SAP / AMS / proyectos / integraciones: **SEIDOR es el competidor PRINCIPAL confirmado por EBIM en Ecuador** — SAP Platinum Partner con presencia en 45 países, +9,000 clientes globales, Centro de Excelencia SAP AMS (soporte funcional, correctivo, evolutivo, autorizaciones), Cloud/SuccessFactors/HANA/Business One/AWS/Azure. SEIDOR SÍ hace implementaciones (EBIM no) pero compite de igual a igual con EBIM en AMS, proyectos de mejora e integraciones — esa es la diferencia clave a comunicar: "EBIM no implementa, pero en soporte/AMS/integraciones somos alternativa real frente a un jugador del tamaño de SEIDOR, con más cercanía y velocidad de respuesta". Sypsoft360 y Heinsohn Ecuador son partners de SAP Business One para PyMEs (add-ons, soporte técnico), un nicho distinto y más pequeño que SEIDOR.
- Odoo ERP: TresCloud (Gold Partner #1, Quito/Guayaquil/Cuenca, se posiciona explícitamente con "IA conectada a la operación" — EBIM debe igualar o superar ese discurso), IntiTecnología (Gold, localización propia), NextGen. Precios referenciales: implementaciones básicas 40-80h, proyectos avanzados 100-250h, costo típico primer año USD 5,000-9,000; también hay modelos SaaS gestionado tipo ~USD 299/mes todo incluido. Localización SRI/IESS es tabla de entrada obligatoria.
- IA aplicada/agentes: ecosistema boutique real — Zetri (Guayaquil), Baigency, Innovación IA, ToGrow, Vex AI, Agencia IA. ~40% de empresas ecuatorianas proyectadas con IA incorporada a fin de 2025; 87% de profesionales ya usa IA en el trabajo; solo 12-18% tiene gobierno de IA formal; sectores líderes banca/retail/agroindustria.
- Presencia de EBIM en Ecuador: CONFIRMADA y real (no es atención remota desde Lima) — sitio dedicado grupoebim.com/ec/, vacantes activas en Guayaquil, operación en Guayaquil/Quito/Cuenca/Manta. El discurso de "equipo local en Ecuador" es defendible.

GENERAL:
- Hallazgo clave: NINGÚN competidor SAP/Odoo boutique investigado publica tarifas por hora — la opacidad de precios es la norma. El Estimador mismo (con números claros y rápidos) ya es un diferenciador de venta.
- Diferenciador ganador: la COMBINACIÓN soporte SAP técnico+funcional (AMS/staffing/bolsa de horas) + implementación Odoo + IA aplicada + productos SaaS propios con integración nativa a SAP/Oracle — ningún competidor listado combina las cuatro cosas.
Usa esta base para "analisisCompetencia" según el país del cliente; solo usa web_search para investigar al CLIENTE específico (perfilCliente), no para redescubrir el panorama competitivo general de PE/EC — eso ya está resuelto arriba y no debe variar entre corridas.`;

// EBIM opera en Perú y Ecuador. Tarifas por país en USD/hora.
// cost = costo real del consultor (cargado); sale = precio de venta competitivo de EBIM.
// minRent = rentabilidad neta mínima aceptable; maxDesc = descuento competitivo máx.
const COUNTRIES = {
  PE: { name: "Perú", flag: "🇵🇪", flagBg: "linear-gradient(90deg,#D91023 0 33.34%,#fff 33.34% 66.66%,#D91023 66.66% 100%)", cur: "PEN", fx: 3.75, minRent: 0.20, maxDesc: 0.10, note: "Mercado base de EBIM. Competencia local fuerte en SAP y desarrollo; diferénciate por experiencia y entregables accionables.",
        rates: { senior: { cost: 26, sale: 62.5 }, semi: { cost: 17, sale: 41.67 }, analista: { cost: 11, sale: 25 } } },
  EC: { name: "Ecuador", flag: "🇪🇨", flagBg: "linear-gradient(180deg,#FFDD00 0 50%,#034EA2 50% 75%,#EF3340 75% 100%)", emblem: true, cur: "USD", fx: 1, minRent: 0.25, maxDesc: 0.05, note: "Economía dolarizada: precios en USD directos, sin riesgo cambiario. Menos competencia SAP local que Perú — política más exigente (piso de rentabilidad más alto, menos margen de descuento) para capturar esa ventaja.",
        rates: { senior: { cost: 24, sale: 58 }, semi: { cost: 16, sale: 38 }, analista: { cost: 10, sale: 23 } } },
};

// Ícono de bandera dibujado con CSS (evita depender de la fuente de emoji de banderas, ausente en Windows nativo).
// Ecuador lleva un pequeño emblema al centro: sin él, sus franjas amarillo/azul/rojo son indistinguibles de la bandera de Colombia.
function FlagIcon({ country, size = 16 }) {
  if (!country) return null;
  const h = Math.round(size * 0.72);
  return (
    <span className="flag-icon" title={country.name} style={{ width: size, height: h, background: country.flagBg, position: "relative", display: "inline-block" }}>
      {country.emblem && (
        <span style={{ position: "absolute", top: "50%", left: "50%", width: Math.round(size * 0.3), height: Math.round(size * 0.3), transform: "translate(-50%,-50%)", borderRadius: "50%", background: "#fff", boxShadow: "0 0 0 1px rgba(0,0,0,.3) inset" }} />
      )}
    </span>
  );
}

const TIERS = [["senior", "Senior / Arquitecto / PM"], ["semi", "Semi-Senior"], ["analista", "Analista"]];
const cloneRates = (r) => ({ senior: { ...r.senior }, semi: { ...r.semi }, analista: { ...r.analista } });

// Paleta de marca EBIM — única fuente de verdad para la app y el HTML de la propuesta (que vive en su propio documento/iframe).
const BRAND = {
  bg: "#EBEEF2", surface: "#FFFFFF", ink: "#15202E", muted: "#5E6E81", line: "#D9E0E8", lineSoft: "#E7ECF1",
  accent: "#056769", accentSoft: "#E2EFF0", gold: "#9A6B12", ok: "#0F8A5F", warn: "#B7791F", bad: "#C13B3B",
  th: "#F4F8F9",
};

// Resuelve el país detectado por la IA a una clave válida (solo Perú/Ecuador)
const COUNTRY_ALIASES = { PERU: "PE", "PERÚ": "PE", ECUADOR: "EC" };
function resolveCountryKey(val) {
  if (!val) return null;
  const v = String(val).trim().toUpperCase();
  let key = null;
  if (COUNTRIES[v]) key = v;
  else if (COUNTRY_ALIASES[v]) key = COUNTRY_ALIASES[v];
  else { const byName = Object.entries(COUNTRIES).find(([, c]) => v.includes(c.name.toUpperCase())); key = byName ? byName[0] : null; }
  return key && COUNTRIES[key] ? key : null;
}

const fmtUSD = (n) =>
  "$" + (Number.isFinite(n) ? n : 0).toLocaleString("en-US", { minimumFractionDigits: 0, maximumFractionDigits: 0 });
const fmtLocal = (n, cur) =>
  (Number.isFinite(n) ? n : 0).toLocaleString("es-PE", { minimumFractionDigits: 0, maximumFractionDigits: 0 }) + " " + cur;
const pct = (n) => (Number.isFinite(n) ? (n * 100).toFixed(1) : "0.0") + "%";
const esc = (s) => String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

// Convierte un ArrayBuffer a base64 en chunks (evita reventar el call stack de String.fromCharCode
// con archivos grandes) — usado para adjuntar el Excel al correo desde el Historial.
function arrayBufferToBase64(buf) {
  let binary = "";
  const bytes = new Uint8Array(buf);
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) {
    binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
  }
  return btoa(binary);
}

// Correlativo desde 1 (no aleatorio): el próximo número es la cantidad de cotizaciones existentes + 1.
const newReqCode = (existing = []) => {
  let n = existing.length + 1;
  let c = "REQ-" + String(n).padStart(6, "0");
  while (existing.includes(c)) { n++; c = "REQ-" + String(n).padStart(6, "0"); }
  return c;
};

// `hours` mapea id de integrante del equipo → horas asignadas a ese entregable (no por tier genérico,
// para que cada columna de la tabla corresponda a una persona/rol real, ej. "PM" o "Arquitecto TI").
const blankDeliverable = (name = "") => ({ id: crypto.randomUUID(), name, hours: {}, start: 0, dur: 1 });

const defaultTeam = () => [
  { id: crypto.randomUUID(), perfil: "Arquitecto TI / Consultor Senior", tier: "senior", rol: "Lidera diagnóstico, arquitectura y roadmap" },
  { id: crypto.randomUUID(), perfil: "Consultor Semi-Senior TI", tier: "semi", rol: "Levantamiento técnico: apps, plataformas, APIs" },
  { id: crypto.randomUUID(), perfil: "Analista de Procesos / Cap. TI", tier: "analista", rol: "Relevamiento de equipo y brechas" },
  { id: crypto.randomUUID(), perfil: "PM / Coordinador de Proyecto", tier: "senior", rol: "Gestión, cronograma y presentación ejecutiva" },
];
const defaultSchedule = () => [
  { id: crypto.randomUUID(), hito: "Inicio del Requerimiento", pct: 0.35 },
  { id: crypto.randomUUID(), hito: "Inicio de Pruebas Integrales", pct: 0.35 },
  { id: crypto.randomUUID(), hito: "Entrega / Pase a Producción", pct: 0.30 },
];

// Un "escenario" es una variante cotizable dentro de la misma estimación (p. ej. "Presencial" vs.
// "Remoto — 1 semana"): tiene su propio equipo/entregables/cronograma/descuento y por lo tanto su
// propio precio, pero comparte cliente, país, tarifas y gastos de estructura con los demás.
const MAX_SCENARIOS = 3;
const blankScenario = (name = "Escenario 1") => ({
  id: crypto.randomUUID(),
  name,
  team: defaultTeam(),
  deliverables: [blankDeliverable("1. Kick-off, planificación y accesos")],
  schedule: defaultSchedule(),
  discount: 0,
  reasoning: "",
});

const defaultEstimation = (countryKey = "PE", code) => {
  const c = COUNTRIES[countryKey] || COUNTRIES.PE;
  const scenario = blankScenario("Escenario 1");
  return {
    code: code || newReqCode(),
    client: "",
    project: "",
    country: countryKey,
    fx: c.fx,
    rates: cloneRates(c.rates),
    context: "",
    insight: { resumenRequerimiento: "", perfilCliente: "", competencia: "", valor: [], ia: "", estrategiaCierre: "", hechosClave: [] },
    scenarios: [scenario],
    activeScenarioId: scenario.id,
    validDays: 30,
    margin: 0.35,
    adm: 0.045,
    com: 0.05,
    mkt: 0.01,
    savedAt: null,
  };
};

// Compatibilidad con entregables guardados antes de que las horas se capturaran por integrante real
// del equipo (antes eran hS/hM/hA por tier). Reparte cada bucket de tier al primer integrante de ese
// tier — mismo total de horas y de costo (la tarifa es por tier), el reparto exacto entre personas que
// comparten tier queda como mejor esfuerzo y el usuario puede reajustarlo a mano.
function migrateDeliverableHours(d, team) {
  if (d.hours) return d;
  const hours = {};
  const firstByTier = (tier) => team.find((t) => t.tier === tier);
  const sr = firstByTier("senior"), se = firstByTier("semi"), an = firstByTier("analista");
  if (sr && d.hS) hours[sr.id] = (hours[sr.id] || 0) + (+d.hS || 0);
  if (se && d.hM) hours[se.id] = (hours[se.id] || 0) + (+d.hM || 0);
  if (an && d.hA) hours[an.id] = (hours[an.id] || 0) + (+d.hA || 0);
  const { hS: _hS, hM: _hM, hA: _hA, ...rest } = d;
  return { ...rest, hours };
}

// Compatibilidad con estimaciones guardadas antes de que existieran los escenarios (team/deliverables/
// schedule/discount vivían en la raíz de `est`, como un único escenario implícito).
function normalizeEstimation(raw) {
  if (!raw) return raw;
  if (Array.isArray(raw.scenarios) && raw.scenarios.length) {
    const scenarios = raw.scenarios.map((s) => ({ ...s, deliverables: s.deliverables.map((d) => migrateDeliverableHours(d, s.team)) }));
    const validActive = raw.activeScenarioId && scenarios.some((s) => s.id === raw.activeScenarioId);
    return { ...raw, scenarios, activeScenarioId: validActive ? raw.activeScenarioId : scenarios[0].id };
  }
  const team = raw.team || defaultTeam();
  const scenario = {
    id: crypto.randomUUID(),
    name: "Escenario 1",
    team,
    deliverables: (raw.deliverables || [blankDeliverable("1. Kick-off, planificación y accesos")]).map((d) => migrateDeliverableHours(d, team)),
    schedule: raw.schedule || defaultSchedule(),
    discount: raw.discount || 0,
    reasoning: "",
  };
  const { team: _team, deliverables: _deliverables, schedule: _schedule, discount: _discount, ...rest } = raw;
  return { ...rest, scenarios: [scenario], activeScenarioId: scenario.id };
}

/* ---------- cálculo central (costo país vs precio de venta competitivo), por escenario ---------- */
function compute(est, scenario) {
  const r = est.rates;
  const team = scenario.team;
  const rows = scenario.deliverables.map((d) => {
    const hours = d.hours || {};
    let hrs = 0, cost = 0, sale = 0;
    team.forEach((t) => {
      const h = hours[t.id] || 0;
      hrs += h;
      cost += h * r[t.tier].cost;
      sale += h * r[t.tier].sale;
    });
    return { ...d, hrs, cost, sale };
  });
  const CO = rows.reduce((s, x) => s + x.cost, 0);          // costo operativo (costo país)
  const PVO = rows.reduce((s, x) => s + x.sale, 0);          // precio de venta competitivo (país), solo horas×tarifa
  const totalHrs = rows.reduce((s, x) => s + x.hrs, 0);
  // Total de horas por integrante real del equipo (para la fila TOTAL de la tabla, una columna por persona).
  const totalByTeam = team.map((t) => ({ id: t.id, perfil: t.perfil, hrs: rows.reduce((s, x) => s + ((x.hours || {})[t.id] || 0), 0) }));
  const grossMargin = PVO > 0 ? (PVO - CO) / PVO : 0;        // margen bruto resultante (solo labor, sin gastos de estructura)
  // Gastos sobre CO (costo país), no sobre PVO: si no, un descuento alto los "encoge" y
  // la rentabilidad aparente queda inflada justo cuando más importa que sea realista (decisión de Lia).
  const gAdm = est.adm * CO, gCom = est.com * CO, gMkt = est.mkt * CO;
  const gastos = gAdm + gCom + gMkt;
  const floorPrice = est.margin < 1 ? (CO + gastos) / (1 - est.margin) : (CO + gastos); // piso por margen objetivo, ya con gastos incluidos
  // Full cost recovery (decisión de Dennis 2026-07-01): Adm/Comercial/MKT se suman al precio antes
  // del descuento, en vez de solo restarse de la utilidad — así el cliente cubre esos gastos de
  // estructura y no quedan absorbidos silenciosamente del margen de EBIM.
  const PVOconGastos = PVO + gastos;
  const PVfinal = PVOconGastos * (1 - (scenario.discount || 0));
  const utilidad = PVfinal - CO - gastos;
  const rent = PVfinal > 0 ? utilidad / PVfinal : 0;
  const weeksTotal = rows.reduce((m, x) => Math.max(m, (x.start || 0) + (x.dur || 1)), 1);
  return { rows, CO, PVO, PVOconGastos, totalHrs, totalByTeam, weeksTotal, grossMargin, floorPrice, gAdm, gCom, gMkt, gastos, PVfinal, utilidad, rent };
}

function scenarioRow(calc, d, minRent) {
  const price = calc.PVOconGastos * (1 - d);
  const util = price - calc.CO - calc.gastos;
  const rent = price > 0 ? util / price : 0;
  let light = "ok", txt = "Aceptable — sobre el mínimo";
  if (rent < 0) { light = "bad"; txt = "No recomendado — rentabilidad negativa"; }
  else if (rent < minRent) { light = "warn"; txt = "Revisar — bajo el mínimo del país"; }
  else if (d === 0) { txt = "Precio óptimo — mantener si es posible"; }
  return { d, price, util, rent, light, txt };
}

// Input numérico para porcentajes: mantiene el texto que el usuario está tecleando en estado local
// y solo lo reformatea (a "35.0") al perder foco — si se reformatea en cada onChange (como hacía antes
// con value={(x*100).toFixed(1)} directo), el valor "salta" mientras se escribe y parece que no deja editar.
function PctInput({ value, onChange, decimals = 1, step = 0.5 }) {
  const [text, setText] = useState((value * 100).toFixed(decimals));
  const [focused, setFocused] = useState(false);
  useEffect(() => { if (!focused) setText((value * 100).toFixed(decimals)); }, [value, focused, decimals]);
  return (
    <span className="pct-input-wrap">
      <input
        className="num"
        type="number"
        step={step}
        value={text}
        onFocus={() => setFocused(true)}
        onChange={(e) => {
          setText(e.target.value);
          const n = parseFloat(e.target.value);
          if (Number.isFinite(n)) onChange(n / 100);
        }}
        onBlur={() => { setFocused(false); setText((value * 100).toFixed(decimals)); }}
      />
      <span className="pct-suffix">%</span>
    </span>
  );
}

// Campo que se ve como texto en negrita (no como una caja de input) y se edita haciendo clic encima,
// como el título de un documento — usado en la barra de Cliente/Proyecto/T.C./Validez para que la
// barra se sienta menos "formulario" y más a la lectura directa del dato.
function InlineLabel({ value, onChange, placeholder = "", type = "text", percent = false, decimals = 1, multiline = false, className = "", inputStyle, step, min }) {
  const [editing, setEditing] = useState(false);
  const ref = useRef(null);
  useEffect(() => { if (editing) { ref.current?.focus(); if (!multiline) ref.current?.select(); } }, [editing, multiline]);
  const commit = (raw) => {
    if (percent) onChange((parseFloat(raw) || 0) / 100);
    else onChange(type === "number" ? +raw : raw);
    setEditing(false);
  };
  if (editing) {
    if (multiline) {
      return (
        <textarea
          ref={ref}
          className={`inline-edit-input inline-edit-textarea ${className}`}
          defaultValue={value}
          style={inputStyle}
          onBlur={(e) => commit(e.target.value)}
          onKeyDown={(e) => { if (e.key === "Escape") setEditing(false); }}
        />
      );
    }
    return (
      <input
        ref={ref}
        className={`inline-edit-input ${className}`}
        type={percent ? "number" : type}
        step={step}
        min={min}
        defaultValue={percent ? (value * 100).toFixed(decimals) : value}
        style={inputStyle}
        onBlur={(e) => commit(e.target.value)}
        onKeyDown={(e) => { if (e.key === "Enter") e.currentTarget.blur(); if (e.key === "Escape") setEditing(false); }}
      />
    );
  }
  const isEmpty = !percent && (value === "" || value == null);
  const display = percent ? `${(value * 100).toFixed(decimals)}%` : String(value);
  return (
    <span className={`inline-label ${multiline ? "inline-label-multiline" : ""} ${isEmpty ? "placeholder" : ""} ${className}`} tabIndex={0}
      onClick={() => setEditing(true)}
      onKeyDown={(e) => { if (e.key === "Enter") setEditing(true); }}>
      {isEmpty ? placeholder : display}
    </span>
  );
}

// Modal centrado reutilizable para los paneles secundarios del menú lateral (Resumen, Estrategia,
// Parámetros, Inteligencia Comercial, Uso interno) — evita repetir el mismo overlay+card 5 veces y
// permite ajustar el tamaño/espaciado de todos a la vez.
function ModalCard({ onClose, title, maxWidth = 640, cardStyle, children }) {
  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="card modal-card" onClick={(e) => e.stopPropagation()} style={{ maxWidth, ...cardStyle }}>
        <h3 style={{ justifyContent: "space-between" }}>
          {typeof title === "string" ? <span>{title}</span> : title}
          <button className="iconbtn" aria-label="Cerrar" title="Cerrar" onClick={onClose}><X size={16} /></button>
        </h3>
        {children}
      </div>
    </div>
  );
}

/* ---------- Propuesta comercial (cara al cliente) ---------- */
function defaultProse(est) {
  return {
    entendimiento: `${est.client || "Su organización"} requiere ${est.project || "un servicio de consultoría de TI"}. En GRUPO EBIM entendemos su impacto en la operación y proponemos un enfoque orientado a resultados medibles.`,
    enfoque: `Trabajamos con metodologías ágiles y un equipo especializado, con entregables claros en cada fase y acompañamiento continuo hasta la adopción.`,
    diferenciadores: (est.insight?.valor?.length ? est.insight.valor : ["Equipo certificado con más de 7 años de experiencia", "Entregables accionables, no solo diagnósticos", "Acompañamiento real en la adopción"]),
    porQueEbim: `Con más de 7 años y presencia en Perú y Ecuador, GRUPO EBIM combina experiencia en SAP, desarrollo a medida, cloud e inteligencia artificial para entregar soluciones que generan valor real.`,
    cierre: `Quedamos atentos para iniciar a la brevedad y convertirnos en su aliado estratégico de tecnología.`,
  };
}

/* ---------- Estilos para Excel (ExcelJS) ---------- */
const xArgb = (hex) => "FF" + hex.replace("#", "").toUpperCase();
const X_THIN = { style: "thin", color: { argb: xArgb(BRAND.line) } };
const X_BORDER = { top: X_THIN, left: X_THIN, bottom: X_THIN, right: X_THIN };
const X_FMT_USD = '"$"#,##0';
const X_FMT_NUM = "#,##0";
const X_FMT_PCT = "0.0%";
const X_LIGHT_FILL = { ok: "E6F5EE", warn: "FBF1DD", bad: "FBE6E6" };
const X_LIGHT_FONT = { ok: BRAND.ok, warn: BRAND.warn, bad: BRAND.bad };

function xlTitle(ws, text, cols) {
  const row = ws.addRow([text]);
  ws.mergeCells(row.number, 1, row.number, cols);
  row.height = 24;
  const cell = row.getCell(1);
  cell.font = { bold: true, size: 12, color: { argb: "FFFFFFFF" } };
  cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: xArgb(BRAND.ink) } };
  cell.alignment = { vertical: "middle", indent: 1 };
  return row;
}
function xlSubtitle(ws, text, cols) {
  const row = ws.addRow([text]);
  ws.mergeCells(row.number, 1, row.number, cols);
  row.getCell(1).font = { italic: true, size: 10, color: { argb: xArgb(BRAND.muted) } };
  row.getCell(1).alignment = { indent: 1 };
  return row;
}
function xlSection(ws, text, cols) {
  const row = ws.addRow([text]);
  ws.mergeCells(row.number, 1, row.number, cols);
  row.height = 20;
  const cell = row.getCell(1);
  cell.font = { bold: true, size: 11, color: { argb: xArgb(BRAND.accent) } };
  cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: xArgb(BRAND.accentSoft) } };
  cell.alignment = { vertical: "middle", indent: 1 };
  return row;
}
function xlHeader(ws, values) {
  const row = ws.addRow(values);
  row.eachCell({ includeEmpty: true }, (cell) => {
    cell.font = { bold: true, size: 10, color: { argb: xArgb(BRAND.muted) } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: xArgb(BRAND.th) } };
    cell.border = X_BORDER;
    cell.alignment = { vertical: "middle" };
  });
  return row;
}
function xlRow(ws, values, { formats = [], bold = false, fillHex = null } = {}) {
  const row = ws.addRow(values);
  row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
    cell.border = X_BORDER;
    if (bold) cell.font = { ...(cell.font || {}), bold: true };
    if (fillHex) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: xArgb(fillHex) } };
    const fmt = formats[colNumber - 1];
    if (fmt) cell.numFmt = fmt;
  });
  return row;
}

// Convierte el logo importado (URL de asset) a data-URL una sola vez, para que la propuesta
// descargada sea un HTML autocontenido (sin depender de que el servidor siga corriendo).
let logoDataUrlCache = null;
async function getLogoDataUrl() {
  if (logoDataUrlCache) return logoDataUrlCache;
  const res = await fetch(ebimLogoCompleto);
  const blob = await res.blob();
  logoDataUrlCache = await new Promise((resolve) => {
    const r = new FileReader();
    r.onload = () => resolve(r.result);
    r.readAsDataURL(blob);
  });
  return logoDataUrlCache;
}

function buildProposalHTML(est, scenarioCalcs, c, prose, logoDataUrl) {
  const today = new Date().toLocaleDateString("es-PE", { day: "2-digit", month: "long", year: "numeric" });
  const multi = scenarioCalcs.length > 1;
  const hechos = (est.insight?.hechosClave || []).filter(Boolean).map((h) => `<li>${esc(h)}</li>`).join("");
  const valor = (prose.diferenciadores || []).filter(Boolean).map((v) => `<li>${esc(v)}</li>`).join("");
  const iaBlock = est.insight?.ia && !/no aplica/i.test(est.insight.ia)
    ? `<section class="ia-callout"><h2>✨ Innovación con Inteligencia Artificial</h2><p>${esc(est.insight.ia)}</p></section>` : "";

  // Alcance/Equipo/Cronograma/Inversión se repiten por escenario cuando hay más de uno a comparar
  // (ej. Presencial vs. Remoto); con un solo escenario el resultado es idéntico al de antes.
  const scenarioSections = scenarioCalcs.map(({ scenario, calc }) => {
    // Rango de negociación en vez de un número rígido (decisión de Lia): techo = precio ya configurado,
    // piso = lo más bajo que EBIM puede llegar sin perder margen, dado el descuento máximo competitivo del país.
    const precioTecho = calc.PVfinal;
    const precioPiso = Math.min(precioTecho, calc.PVOconGastos * (1 - c.maxDesc));
    const hayRango = precioTecho - precioPiso > 1;
    const weeks = calc.weeksTotal || Math.max(1, Math.ceil(calc.totalHrs / (Math.max(1, scenario.team.filter((t) => t.perfil).length) * 32)));
    const gantt = scenario.deliverables.filter((d) => d.name).map((d) => {
      const s = d.start || 0, du = d.dur || 1;
      return `<tr><td style="font-size:13px">${esc(d.name)}</td><td style="width:58%"><div style="position:relative;height:18px;background:#EEF3F5;border-radius:5px"><div style="position:absolute;left:${(s / weeks) * 100}%;width:${(du / weeks) * 100}%;top:0;bottom:0;background:#0B5563;border-radius:5px"></div></div></td><td class="r" style="white-space:nowrap;font-size:12px;color:#5E6E81">Sem ${s + 1}–${s + du}</td></tr>`;
    }).join("");
    const deliv = scenario.deliverables.filter((d) => d.name).map((d) => `<li>${esc(d.name)}</li>`).join("");
    const team = scenario.team.filter((t) => t.perfil).map((t) => `<tr><td><b>${esc(t.perfil)}</b></td><td>${esc(t.rol)}</td></tr>`).join("");
    const pay = scenario.schedule.map((s, i) => `<tr><td>${i + 1}</td><td>${esc(s.hito)}</td><td class="r">${(s.pct * 100).toFixed(0)}%</td><td class="r">${fmtUSD(calc.PVfinal * s.pct)}</td></tr>`).join("");
    const heading = multi ? `<div class="eyebrow" style="margin-top:36px">Opción — ${esc(scenario.name)}</div>` : "";
    return `${heading}
  <section><h2>Alcance y entregables${multi ? " — " + esc(scenario.name) : ""}</h2><ul>${deliv || "<li>Por definir</li>"}</ul>
    ${!multi && hechos ? `<p style="margin-top:14px;font-size:13px;color:${BRAND.muted}"><b>Dimensionado sobre la base de:</b></p><ul style="font-size:13px;color:${BRAND.muted}">${hechos}</ul>` : ""}
  </section>
  <section><h2>Equipo asignado${multi ? " — " + esc(scenario.name) : ""}</h2><table><thead><tr><th>Perfil</th><th>Responsabilidad</th></tr></thead><tbody>${team || "<tr><td>Por definir</td><td></td></tr>"}</tbody></table></section>
  <section><h2>Cronograma de trabajo${multi ? " — " + esc(scenario.name) : ""}</h2><p>Duración estimada de <b>${weeks} semana${weeks > 1 ? "s" : ""}</b>, organizada por entregable con hitos de avance y validación con su equipo:</p>
    <table><thead><tr><th>Entregable</th><th>Avance en el tiempo</th><th class="r">Semanas</th></tr></thead><tbody>${gantt || "<tr><td>Por definir</td><td></td><td></td></tr>"}</tbody></table></section>
  <section><h2>Inversión${multi ? " — " + esc(scenario.name) : ""}</h2>
    <div class="invest"><div><div class="label">Inversión — llave en mano</div><small>No incluye IGV${hayRango ? " · rango según alcance final acordado" : ""}</small></div><div style="text-align:right">
      ${hayRango
        ? `<div class="amt">USD ${fmtUSD(precioPiso).replace("$", "")} – ${fmtUSD(precioTecho).replace("$", "")}</div><small>≈ ${fmtLocal(precioPiso * est.fx, c.cur)} – ${fmtLocal(precioTecho * est.fx, c.cur)}</small>`
        : `<div class="amt">USD ${fmtUSD(precioTecho).replace("$", "")}</div><small>≈ ${fmtLocal(precioTecho * est.fx, c.cur)}</small>`}
    </div></div>
    <table style="margin-top:16px"><thead><tr><th>N°</th><th>Hito de pago</th><th class="r">%</th><th class="r">Monto USD</th></tr></thead><tbody>${pay}</tbody></table>
  </section>`;
  }).join("\n");

  return `<!doctype html><html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Propuesta ${esc(est.client)} — ${esc(est.code)}</title>
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@600;700&family=Inter:wght@400;500;600&display=swap');
*{box-sizing:border-box} body{margin:0;font-family:'Inter',system-ui,sans-serif;color:${BRAND.ink};line-height:1.55;background:${BRAND.surface}}
.page{max-width:820px;margin:0 auto;padding:48px 56px}
.disp{font-family:'Space Grotesk',sans-serif}
.cover{min-height:88vh;display:flex;flex-direction:column;justify-content:center;border-left:6px solid ${BRAND.accent};padding-left:36px}
.brand{display:flex;align-items:center;gap:12px;margin-bottom:40px}
.eyebrow{letter-spacing:.18em;text-transform:uppercase;font-size:12px;color:${BRAND.accent};font-weight:600;margin-bottom:14px}
.cover h1{font-family:'Space Grotesk';font-size:40px;line-height:1.1;margin:0 0 18px;font-weight:700}
.cover .meta{color:${BRAND.muted};font-size:15px}
.cover .meta b{color:${BRAND.ink}}
.req{display:inline-block;margin-top:26px;font-family:monospace;background:${BRAND.ink};color:#fff;padding:7px 14px;border-radius:7px;font-weight:600}
section{margin:30px 0;page-break-inside:avoid}
h2{font-family:'Space Grotesk';font-size:20px;color:${BRAND.accent};border-bottom:2px solid ${BRAND.accentSoft};padding-bottom:7px;margin:0 0 12px}
ul{margin:8px 0;padding-left:20px} li{margin:5px 0}
table{width:100%;border-collapse:collapse;margin-top:8px;font-size:14px}
td,th{padding:9px 10px;border-bottom:1px solid ${BRAND.lineSoft};text-align:left;vertical-align:top}
th{background:${BRAND.th};font-size:12px;text-transform:uppercase;letter-spacing:.04em;color:${BRAND.muted}}
.r{text-align:right}
.invest{background:${BRAND.accent};color:#fff;border-radius:14px;padding:26px 30px;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:14px}
.invest .label{letter-spacing:.14em;text-transform:uppercase;font-size:12px;opacity:.85}
.invest .amt{font-family:'Space Grotesk';font-size:38px;font-weight:700;line-height:1}
.invest small{opacity:.8}
.foot{margin-top:46px;border-top:1px solid ${BRAND.lineSoft};padding-top:18px;color:${BRAND.muted};font-size:12.5px;display:flex;justify-content:space-between;flex-wrap:wrap;gap:8px}
.tag{display:inline-block;background:${BRAND.accentSoft};color:${BRAND.accent};border-radius:20px;padding:3px 12px;font-size:12px;font-weight:600;margin:0 6px 6px 0}
.highlight{background:${BRAND.accentSoft};border-left:4px solid ${BRAND.accent};border-radius:0 12px 12px 0;padding:22px 26px}
.highlight h2{border-bottom:none;padding-bottom:0;margin-bottom:14px}
.checklist{list-style:none;margin:0;padding:0}
.checklist li{padding:6px 0 6px 28px;position:relative;font-size:14.5px}
.checklist li::before{content:"✓";position:absolute;left:0;top:6px;color:${BRAND.accent};font-weight:700}
.ia-callout{background:${BRAND.ink};color:#fff;border-radius:12px;padding:22px 26px}
.ia-callout h2{color:#fff;border-bottom:none;padding-bottom:0;margin-bottom:10px}
@media print{.page{padding:24px 30px}.cover{min-height:94vh}}
</style></head><body>
<div class="page">
  <div class="cover">
    <div class="brand"><img src="${logoDataUrl}" alt="EBIM" style="height:32px"/><div style="color:${BRAND.muted};font-size:13px">GRUPO EBIM SAC · Consultoría en Tecnologías de la Información</div></div>
    <div class="eyebrow">Propuesta de servicios profesionales</div>
    <h1>${esc(est.project || "Servicio de consultoría TI")}</h1>
    <div class="meta">Preparada para <b>${esc(est.client || "—")}</b> · ${esc(c.name)}<br>${today}</div>
    <div class="req">${esc(est.code)}</div>
  </div>

  <section><h2>Resumen ejecutivo</h2><p>${esc(prose.entendimiento)}</p></section>
  ${est.insight?.perfilCliente ? `<section><h2>Entendimiento de su negocio</h2><p>${esc(est.insight.perfilCliente)}</p></section>` : ""}
  <section><h2>Nuestro enfoque</h2><p>${esc(prose.enfoque)}</p></section>
  ${multi && hechos ? `<section><h2>Dimensionado sobre la base de</h2><ul style="font-size:13px;color:${BRAND.muted}">${hechos}</ul></section>` : ""}
  ${scenarioSections}
  <section class="highlight"><h2>★ Valor agregado y diferenciadores</h2><ul class="checklist">${valor}</ul></section>
  ${iaBlock}
  <section><h2>Por qué GRUPO EBIM</h2><p>${esc(prose.porQueEbim)}</p>
    <div><span class="tag">Soporte y AMS SAP</span><span class="tag">Implementación Odoo ERP</span><span class="tag">Desarrollo a medida</span><span class="tag">Cloud AWS · Azure · GCP</span><span class="tag">Inteligencia Artificial</span><span class="tag">+7 años</span></div></section>
  <section><h2>Siguientes pasos</h2><p>${esc(prose.cierre)}</p></section>
  <div class="foot"><span>GRUPO EBIM SAC · Consultoría TI · Lima, Perú · www.grupoebim.com</span><span>${esc(est.code)} · Oferta válida por ${est.validDays || 30} días</span></div>
</div></body></html>`;
}

/* ---------- Documentos Word (docx) ---------- */
const DOCX_INK = BRAND.ink.replace("#", "");
const DOCX_ACCENT = BRAND.accent.replace("#", "");
const DOCX_MUTED = BRAND.muted.replace("#", "");
const DOCX_LINE = BRAND.line.replace("#", "");
const DOCX_ACCENT_SOFT = BRAND.accentSoft.replace("#", "");
const DOCX_GOLD = BRAND.gold.replace("#", "");

function docxHeading(text) {
  return new Paragraph({
    spacing: { before: 260, after: 120 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: DOCX_ACCENT_SOFT } },
    children: [new TextRun({ text, bold: true, color: DOCX_ACCENT, size: 24 })],
  });
}
function docxSubheading(text) {
  return new Paragraph({ spacing: { before: 120, after: 60 }, children: [new TextRun({ text, bold: true, size: 20, color: DOCX_INK })] });
}
function docxP(text) {
  return new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text: String(text ?? "—"), size: 22 })] });
}
function docxNote(text) {
  return new Paragraph({ spacing: { before: 80, after: 60 }, children: [new TextRun({ text, italics: true, size: 18, color: DOCX_MUTED })] });
}
function docxBullet(text) {
  return new Paragraph({ bullet: { level: 0 }, spacing: { after: 60 }, children: [new TextRun({ text: String(text ?? ""), size: 22 })] });
}
function docxBulletSmall(text) {
  return new Paragraph({ bullet: { level: 0 }, spacing: { after: 40 }, children: [new TextRun({ text: String(text ?? ""), size: 18, color: DOCX_MUTED })] });
}
function docxCell(text, { header = false, width, align = AlignmentType.LEFT } = {}) {
  return new TableCell({
    width: width ? { size: width, type: WidthType.PERCENTAGE } : undefined,
    shading: header ? { type: ShadingType.CLEAR, fill: "F4F8F9" } : undefined,
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 80, bottom: 80, left: 100, right: 100 },
    children: [new Paragraph({ alignment: align, children: [new TextRun({ text: String(text ?? ""), bold: header, size: 20, color: header ? DOCX_MUTED : DOCX_INK })] })],
  });
}
function docxTable(headerRow, rows, widths, aligns) {
  const al = aligns || headerRow.map((_, i) => (i === 0 ? AlignmentType.LEFT : AlignmentType.RIGHT));
  const border = { style: BorderStyle.SINGLE, size: 2, color: DOCX_LINE };
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: { top: border, bottom: border, left: border, right: border, insideHorizontal: border, insideVertical: border },
    rows: [
      new TableRow({ children: headerRow.map((h, i) => docxCell(h, { header: true, width: widths?.[i], align: al[i] })) }),
      ...rows.map((r) => new TableRow({ children: r.map((v, i) => docxCell(v, { width: widths?.[i], align: al[i] })) })),
    ],
  });
}
function docxReqBadge(code) {
  return new Paragraph({
    shading: { type: ShadingType.CLEAR, fill: DOCX_INK },
    spacing: { after: 320 },
    children: [new TextRun({ text: `  ${code}  `, bold: true, color: "FFFFFF", size: 20, font: "Consolas" })],
  });
}

// Propuesta comercial descargable en Word — cara al cliente, mismo contenido que la vista previa HTML.
function buildProposalDocx(est, scenarioCalcs, c, prose) {
  const today = new Date().toLocaleDateString("es-PE", { day: "2-digit", month: "long", year: "numeric" });
  const multi = scenarioCalcs.length > 1;
  const hechos = (est.insight?.hechosClave || []).filter(Boolean);
  const valor = (prose.diferenciadores || []).filter(Boolean);
  const children = [];

  children.push(new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: "PROPUESTA DE SERVICIOS PROFESIONALES", bold: true, color: DOCX_ACCENT, size: 20, characterSpacing: 20 })] }));
  children.push(new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text: est.project || "Servicio de consultoría TI", bold: true, size: 44, color: DOCX_INK })] }));
  children.push(new Paragraph({ spacing: { after: 100 }, children: [new TextRun({ text: `Preparada para ${est.client || "—"} · ${c.name}  ·  ${today}`, size: 22, color: DOCX_MUTED })] }));
  children.push(docxReqBadge(est.code));

  children.push(docxHeading("Resumen ejecutivo"));
  children.push(docxP(prose.entendimiento));

  if (est.insight?.perfilCliente) {
    children.push(docxHeading("Entendimiento de su negocio"));
    children.push(docxP(est.insight.perfilCliente));
  }

  children.push(docxHeading("Nuestro enfoque"));
  children.push(docxP(prose.enfoque));

  if (multi && hechos.length) {
    children.push(docxHeading("Dimensionado sobre la base de"));
    hechos.forEach((h) => children.push(docxBulletSmall(h)));
  }

  // Alcance/Equipo/Cronograma/Inversión se repiten por escenario cuando hay más de uno a comparar
  // (ej. Presencial vs. Remoto); con un solo escenario el resultado es idéntico al de antes.
  scenarioCalcs.forEach(({ scenario, calc }) => {
    const precioTecho = calc.PVfinal;
    const precioPiso = Math.min(precioTecho, calc.PVOconGastos * (1 - c.maxDesc));
    const hayRango = precioTecho - precioPiso > 1;
    const weeks = calc.weeksTotal || Math.max(1, Math.ceil(calc.totalHrs / (Math.max(1, scenario.team.filter((t) => t.perfil).length) * 32)));
    const deliverables = scenario.deliverables.filter((d) => d.name);
    const team = scenario.team.filter((t) => t.perfil);
    const suffix = multi ? ` — ${scenario.name}` : "";

    if (multi) {
      children.push(new Paragraph({
        shading: { type: ShadingType.CLEAR, fill: DOCX_ACCENT_SOFT },
        spacing: { before: 300, after: 100 },
        children: [new TextRun({ text: `OPCIÓN — ${scenario.name.toUpperCase()}`, bold: true, color: DOCX_ACCENT, size: 22, characterSpacing: 10 })],
      }));
    }

    children.push(docxHeading("Alcance y entregables" + suffix));
    if (deliverables.length) deliverables.forEach((d) => children.push(docxBullet(d.name)));
    else children.push(docxP("Por definir"));
    if (!multi && hechos.length) {
      children.push(docxNote("Dimensionado sobre la base de:"));
      hechos.forEach((h) => children.push(docxBulletSmall(h)));
    }

    children.push(docxHeading("Equipo asignado" + suffix));
    children.push(team.length
      ? docxTable(["Perfil", "Responsabilidad"], team.map((t) => [t.perfil, t.rol || ""]), [40, 60], [AlignmentType.LEFT, AlignmentType.LEFT])
      : docxP("Por definir"));

    children.push(docxHeading("Cronograma de trabajo" + suffix));
    children.push(docxP(`Duración estimada de ${weeks} semana${weeks > 1 ? "s" : ""}, organizada por entregable con hitos de avance y validación con su equipo:`));
    if (deliverables.length) {
      children.push(docxTable(["Entregable", "Semanas"], deliverables.map((d) => [d.name, `Sem ${(d.start || 0) + 1}–${(d.start || 0) + (d.dur || 1)}`]), [70, 30]));
    }

    children.push(docxHeading("Inversión" + suffix));
    const montoTexto = hayRango
      ? `USD ${fmtUSD(precioPiso).replace("$", "")} – ${fmtUSD(precioTecho).replace("$", "")}`
      : `USD ${fmtUSD(precioTecho).replace("$", "")}`;
    children.push(new Paragraph({ shading: { type: ShadingType.CLEAR, fill: DOCX_ACCENT }, spacing: { after: 40 }, children: [new TextRun({ text: "INVERSIÓN — LLAVE EN MANO", bold: true, color: "FFFFFF", size: 18, characterSpacing: 15 })] }));
    children.push(new Paragraph({
      shading: { type: ShadingType.CLEAR, fill: DOCX_ACCENT },
      spacing: { after: 220 },
      children: [
        new TextRun({ text: montoTexto, bold: true, color: "FFFFFF", size: 40 }),
        new TextRun({ text: `   (No incluye IGV${hayRango ? " · rango según alcance final acordado" : ""})`, color: "FFFFFF", size: 16 }),
      ],
    }));
    if (scenario.schedule.length) {
      children.push(docxTable(
        ["N°", "Hito de pago", "%", "Monto USD"],
        scenario.schedule.map((s, i) => [String(i + 1), s.hito, pct(s.pct), fmtUSD(calc.PVfinal * s.pct)]),
        [8, 52, 15, 25],
        [AlignmentType.LEFT, AlignmentType.LEFT, AlignmentType.RIGHT, AlignmentType.RIGHT]
      ));
    }
  });

  children.push(docxHeading("Valor agregado y diferenciadores"));
  if (valor.length) valor.forEach((v) => children.push(docxBullet(v)));
  else children.push(docxP("—"));

  if (est.insight?.ia && !/no aplica/i.test(est.insight.ia)) {
    children.push(new Paragraph({ shading: { type: ShadingType.CLEAR, fill: DOCX_INK }, spacing: { before: 220, after: 60 }, children: [new TextRun({ text: "✨ Innovación con Inteligencia Artificial", bold: true, color: "FFFFFF", size: 22 })] }));
    children.push(new Paragraph({ shading: { type: ShadingType.CLEAR, fill: DOCX_INK }, spacing: { after: 220 }, children: [new TextRun({ text: est.insight.ia, color: "FFFFFF", size: 20 })] }));
  }

  children.push(docxHeading("Por qué GRUPO EBIM"));
  children.push(docxP(prose.porQueEbim));
  children.push(new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: "Soporte y AMS SAP · Implementación Odoo ERP · Desarrollo a medida · Cloud AWS · Azure · GCP · Inteligencia Artificial · +7 años", italics: true, size: 18, color: DOCX_ACCENT })] }));

  children.push(docxHeading("Siguientes pasos"));
  children.push(docxP(prose.cierre));

  children.push(new Paragraph({
    border: { top: { style: BorderStyle.SINGLE, size: 4, color: DOCX_LINE } },
    spacing: { before: 300 },
    children: [new TextRun({ text: `GRUPO EBIM SAC · Consultoría TI · Lima, Perú · www.grupoebim.com  ·  ${est.code} · Oferta válida por ${est.validDays || 30} días`, size: 16, color: DOCX_MUTED })],
  }));

  return new Document({ creator: "GRUPO EBIM SAC", title: `Propuesta ${est.client} — ${est.code}`, sections: [{ properties: {}, children }] });
}

// Documento interno aparte: resumen del proyecto + estrategia comercial + inteligencia comercial completa.
function buildIntelDocx(est, scenarioCalcs, c) {
  const today = new Date().toLocaleDateString("es-PE", { day: "2-digit", month: "long", year: "numeric" });
  const hechos = (est.insight?.hechosClave || []).filter(Boolean);
  const valor = (est.insight?.valor || []).filter(Boolean);
  const multi = scenarioCalcs.length > 1;
  const children = [];

  children.push(new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text: "USO INTERNO EBIM — NO COMPARTIR CON EL CLIENTE", bold: true, color: DOCX_GOLD, size: 18, characterSpacing: 15 })] }));
  children.push(new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text: "Informe de Inteligencia Comercial y Estrategia", bold: true, size: 40, color: DOCX_INK })] }));
  children.push(new Paragraph({ spacing: { after: 100 }, children: [new TextRun({ text: `${est.project || "Proyecto"} · ${est.client || "Cliente"} · ${c.name}  ·  ${today}`, size: 22, color: DOCX_MUTED })] }));
  children.push(docxReqBadge(est.code));

  if (multi) {
    children.push(docxHeading("Comparación de escenarios"));
    children.push(docxTable(
      ["Escenario", "Precio USD", "Rentabilidad", "Horas"],
      scenarioCalcs.map(({ scenario, calc }) => [scenario.name, fmtUSD(calc.PVfinal), pct(calc.rent), String(calc.totalHrs)]),
      [40, 20, 20, 20]
    ));
  }

  children.push(docxHeading("Resumen del proyecto"));
  children.push(docxP(est.insight?.resumenRequerimiento || est.context || "—"));
  if (hechos.length) {
    children.push(docxNote("Cifras clave usadas para dimensionar el alcance:"));
    hechos.forEach((h) => children.push(docxBulletSmall(h)));
  }

  children.push(docxHeading("Estrategia comercial (cierre)"));
  children.push(docxP(est.insight?.estrategiaCierre || "—"));

  children.push(docxHeading("Inteligencia comercial"));
  children.push(docxSubheading("Perfil del cliente"));
  children.push(docxP(est.insight?.perfilCliente || "—"));
  children.push(docxSubheading("Análisis de competencia"));
  children.push(docxP(est.insight?.competencia || "—"));
  children.push(docxSubheading("Valor agregado a proponer"));
  if (valor.length) valor.forEach((v) => children.push(docxBullet(v)));
  else children.push(docxP("—"));
  children.push(docxSubheading("Oportunidad de IA / automatización"));
  children.push(docxP(est.insight?.ia || "—"));

  children.push(new Paragraph({
    border: { top: { style: BorderStyle.SINGLE, size: 4, color: DOCX_LINE } },
    spacing: { before: 300 },
    children: [new TextRun({ text: "GRUPO EBIM SAC · Documento de uso interno — no compartir con el cliente.", size: 16, color: DOCX_MUTED })],
  }));

  return new Document({ creator: "GRUPO EBIM SAC", title: `Inteligencia comercial ${est.client} — ${est.code}`, sections: [{ properties: {}, children }] });
}

// Libro de Excel de una estimación — función pura reutilizable tanto por el botón "Excel" (estimación
// activa) como por el envío de correo desde el Historial (estimaciones guardadas arbitrarias).
function buildEstimationWorkbook(est, c, scenarioCalcs) {
  const wb = new ExcelJS.Workbook();
  wb.creator = "GRUPO EBIM SAC";
  wb.created = new Date();
  const multi = est.scenarios.length > 1;

  // Nombres de hoja únicos y ≤31 caracteres (límite de Excel)
  const usedSheetNames = new Set();
  function sheetName(base, scenarioLabel) {
    let name = (multi ? `${base} — ${scenarioLabel}` : base).slice(0, 31);
    let i = 2;
    while (usedSheetNames.has(name)) { name = `${base.slice(0, 24)} (${i})`; i++; }
    usedSheetNames.add(name);
    return name;
  }

  // Hoja de comparación — solo si hay más de un escenario que comparar
  if (multi) {
    const wsCmp = wb.addWorksheet(sheetName("Comparación", ""));
    wsCmp.columns = [{ width: 26 }, { width: 16 }, { width: 14 }, { width: 14 }, { width: 70 }];
    xlTitle(wsCmp, `COMPARACIÓN DE ESCENARIOS — ${est.code} | ${est.client || "Cliente"} | ${c.name}`, 5);
    wsCmp.addRow([]);
    xlHeader(wsCmp, ["Escenario", "Precio USD", "Rentabilidad", "Horas totales", "Razón del precio"]);
    scenarioCalcs.forEach(({ scenario, calc: sc }) => {
      xlRow(wsCmp, [scenario.name, +sc.PVfinal.toFixed(2), +sc.rent.toFixed(4), sc.totalHrs, scenario.reasoning || "—"], { formats: [null, X_FMT_USD, X_FMT_PCT, X_FMT_NUM] }).getCell(5).alignment = { wrapText: true };
    });
  }

  scenarioCalcs.forEach(({ scenario, calc: sc }) => {
    const discScen = [0, 0.05, 0.10, 0.15].map((d) => scenarioRow(sc, d, c.minRent));

    // Hoja: Estimación
    const ws1 = wb.addWorksheet(sheetName("Estimación", scenario.name));
    const nCols = 1 + scenario.team.length + 4; // Entregable + 1 col por integrante + Total/Costo/Venta/Venta-local
    ws1.columns = [{ width: 46 }, ...scenario.team.map(() => ({ width: 16 })), { width: 12 }, { width: 14 }, { width: 14 }, { width: 16 }];
    xlTitle(ws1, `HOJA DE COSTOS — ${est.code}  |  ${est.project || "Proyecto"}  |  ${est.client || "Cliente"}${multi ? "  |  " + scenario.name : ""}`, nCols);
    xlSubtitle(ws1, `Elaborado por: GRUPO EBIM SAC  |  País: ${c.name}  |  Moneda: USD  |  T.C. ref: ${est.fx} ${c.cur}/USD`, nCols);
    ws1.addRow([]);
    xlSection(ws1, "SUPUESTOS — TODO EN USD", nCols);
    xlHeader(ws1, ["Perfil", "Costo país (USD/h)", "Precio venta (USD/h)"]);
    TIERS.forEach(([k, label]) => xlRow(ws1, [label, est.rates[k].cost, est.rates[k].sale], { formats: [null, X_FMT_USD, X_FMT_USD] }));
    xlRow(ws1, ["Margen objetivo / piso (% s/ PV)", est.margin, ""], { formats: [null, X_FMT_PCT] });
    xlRow(ws1, ["Gastos Adm. + Comercial + MKT (% s/ Costo País)", est.adm + est.com + est.mkt, ""], { formats: [null, X_FMT_PCT] });
    ws1.addRow([]);
    xlSection(ws1, "EQUIPO DEL PROYECTO", nCols);
    xlHeader(ws1, ["Perfil", "Venta USD/h", "Rol en el proyecto"]);
    scenario.team.forEach((t) => xlRow(ws1, [t.perfil, est.rates[t.tier].sale, t.rol], { formats: [null, X_FMT_USD] }));
    ws1.addRow([]);
    xlSection(ws1, "DETALLE DE HORAS Y COSTOS POR ENTREGABLE", nCols);
    xlHeader(ws1, ["Entregable / Actividad", ...scenario.team.map((t) => t.perfil || "(sin nombre)"), "Total Hrs", "Costo USD", "Venta USD", "Venta " + c.cur]);
    const numFmts1 = [null, ...scenario.team.map(() => X_FMT_NUM), X_FMT_NUM, X_FMT_USD, X_FMT_USD, X_FMT_NUM];
    sc.rows.forEach((d) => xlRow(ws1, [d.name, ...scenario.team.map((t) => (d.hours || {})[t.id] || 0), d.hrs, +d.cost.toFixed(2), +d.sale.toFixed(2), +(d.sale * est.fx).toFixed(2)], { formats: numFmts1 }));
    xlRow(ws1, ["TOTAL", ...sc.totalByTeam.map((t) => t.hrs), sc.totalHrs, +sc.CO.toFixed(2), +sc.PVO.toFixed(2), +(sc.PVO * est.fx).toFixed(2)], { formats: numFmts1, bold: true, fillHex: BRAND.th });
    ws1.addRow([]);
    xlSection(ws1, "RESUMEN COMERCIAL Y PRECIO DE VENTA", nCols);
    xlHeader(ws1, ["Concepto", "USD", "% / Factor", c.cur]);
    xlRow(ws1, ["Costo Total Operaciones (costo país)", +sc.CO.toFixed(2), "—", +(sc.CO * est.fx).toFixed(2)], { formats: [null, X_FMT_USD, null, X_FMT_NUM] });
    xlRow(ws1, ["Precio de Venta Operaciones (competitivo país)", +sc.PVO.toFixed(2), "—", +(sc.PVO * est.fx).toFixed(2)], { formats: [null, X_FMT_USD, null, X_FMT_NUM] });
    xlRow(ws1, ["Margen bruto resultante", "—", +sc.grossMargin.toFixed(4), "—"], { formats: [null, null, X_FMT_PCT] });
    xlRow(ws1, ["Gastos Administrativos", +sc.gAdm.toFixed(2), est.adm, +(sc.gAdm * est.fx).toFixed(2)], { formats: [null, X_FMT_USD, X_FMT_PCT, X_FMT_NUM] });
    xlRow(ws1, ["Gestión Comercial — Carmen", +sc.gCom.toFixed(2), est.com, +(sc.gCom * est.fx).toFixed(2)], { formats: [null, X_FMT_USD, X_FMT_PCT, X_FMT_NUM] });
    xlRow(ws1, ["Gestión MKT", +sc.gMkt.toFixed(2), est.mkt, +(sc.gMkt * est.fx).toFixed(2)], { formats: [null, X_FMT_USD, X_FMT_PCT, X_FMT_NUM] });
    xlRow(ws1, ["Total Gastos Adm. y Comisiones", +sc.gastos.toFixed(2), "—", +(sc.gastos * est.fx).toFixed(2)], { formats: [null, X_FMT_USD, null, X_FMT_NUM], bold: true });
    xlRow(ws1, ["Descuento comercial (%)", "—", scenario.discount || 0, "—"], { formats: [null, null, X_FMT_PCT] });
    xlRow(ws1, ["PRECIO DE VENTA FINAL (Sin IGV)", +sc.PVfinal.toFixed(2), "", +(sc.PVfinal * est.fx).toFixed(2)], { formats: [null, X_FMT_USD, null, X_FMT_NUM], bold: true, fillHex: BRAND.accentSoft });
    xlRow(ws1, ["Utilidad Neta", +sc.utilidad.toFixed(2), "—", +(sc.utilidad * est.fx).toFixed(2)], { formats: [null, X_FMT_USD, null, X_FMT_NUM] });
    xlRow(ws1, ["Rentabilidad Neta (%)", +sc.rent.toFixed(4), "—", "—"], { formats: [null, X_FMT_PCT], bold: true });
    ws1.addRow([]);
    xlSection(ws1, "CRONOGRAMA DE PAGOS", nCols);
    xlHeader(ws1, ["N°", "Hito", "% del Total", "Importe USD", "Importe " + c.cur]);
    scenario.schedule.forEach((s, i) => xlRow(ws1, [i + 1, s.hito, s.pct, +(sc.PVfinal * s.pct).toFixed(2), +(sc.PVfinal * s.pct * est.fx).toFixed(2)], { formats: [X_FMT_NUM, null, X_FMT_PCT, X_FMT_USD, X_FMT_NUM] }));
    xlRow(ws1, ["", "TOTAL", scenario.schedule.reduce((a, b) => a + b.pct, 0), +sc.PVfinal.toFixed(2), +(sc.PVfinal * est.fx).toFixed(2)], { formats: [null, null, X_FMT_PCT, X_FMT_USD, X_FMT_NUM], bold: true, fillHex: BRAND.th });
    ws1.views = [{ state: "frozen", ySplit: 2 }];

    // Hoja: Sensibilidad de descuento (uso interno)
    const ws2 = wb.addWorksheet(sheetName("Sensibilidad desc.", scenario.name));
    ws2.columns = [{ width: 18 }, { width: 12 }, { width: 14 }, { width: 16 }, { width: 14 }, { width: 14 }, { width: 44 }];
    xlTitle(ws2, `SENSIBILIDAD DE DESCUENTO — ${est.code} | ${est.client || "Cliente"}${multi ? " | " + scenario.name : ""} | USO INTERNO EBIM`, 7);
    xlSubtitle(ws2, `País: ${c.name}  |  Rentabilidad mínima aceptable: ${pct(c.minRent)}  |  Descuento sugerido máx.: ${pct(c.maxDesc)}`, 7);
    ws2.addRow([]);
    xlHeader(ws2, ["Escenario", "Descuento %", "Precio USD", "Precio " + c.cur, "Utilidad USD", "Rentabilidad", "Recomendación"]);
    discScen.forEach((s) => {
      const row = xlRow(ws2, [
        s.d === 0 ? "Sin descuento" : "Descuento " + (s.d * 100) + "%",
        s.d, +s.price.toFixed(2), +(s.price * est.fx).toFixed(2), +s.util.toFixed(2), +s.rent.toFixed(4),
        (s.light === "ok" ? "OK — " : s.light === "warn" ? "REVISAR — " : "NO — ") + s.txt,
      ], { formats: [null, X_FMT_PCT, X_FMT_USD, X_FMT_NUM, X_FMT_USD, X_FMT_PCT] });
      const light = row.getCell(7);
      light.fill = { type: "pattern", pattern: "solid", fgColor: { argb: xArgb(X_LIGHT_FILL[s.light]) } };
      light.font = { bold: true, color: { argb: xArgb(X_LIGHT_FONT[s.light]) } };
    });
    ws2.addRow([]);
    xlSubtitle(ws2, "Nota: hoja de uso interno. El cliente recibe solo el precio llave en mano, sin desglose de descuentos ni márgenes.", 7);

    // Hoja: Cronograma de trabajo (Gantt)
    const ws3 = wb.addWorksheet(sheetName("Cronograma", scenario.name));
    ws3.columns = [{ width: 46 }, { width: 14 }, { width: 16 }, { width: 12 }, { width: 10 }, { width: 14 }, { width: 14 }];
    xlTitle(ws3, `CRONOGRAMA DE TRABAJO — ${est.code} | ${est.client || "Cliente"}${multi ? " | " + scenario.name : ""}`, 7);
    xlSubtitle(ws3, `Duración total estimada: ${sc.weeksTotal} semana(s)`, 7);
    ws3.addRow([]);
    xlHeader(ws3, ["Entregable", "Semana inicio", "Duración (sem)", "Semana fin", "Total hrs", "Costo USD", "Venta USD"]);
    sc.rows.forEach((d) => xlRow(ws3, [d.name, (d.start || 0) + 1, d.dur || 1, (d.start || 0) + (d.dur || 1), d.hrs, +d.cost.toFixed(2), +d.sale.toFixed(2)], { formats: [null, X_FMT_NUM, X_FMT_NUM, X_FMT_NUM, X_FMT_NUM, X_FMT_USD, X_FMT_USD] }));
  });

  // Hoja: Estrategia comercial (a nivel proyecto — cliente, competencia, valor agregado, IA — no varía por escenario)
  const ws4 = wb.addWorksheet(sheetName("Estrategia comercial", ""));
  ws4.columns = [{ width: 110 }];
  xlTitle(ws4, `ESTRATEGIA COMERCIAL — ${est.code} | ${est.client || "Cliente"} | ${c.name}`, 1);
  ws4.addRow([]);
  xlSection(ws4, "Perfil del cliente / qué ofrecerle", 1);
  ws4.addRow([est.insight?.perfilCliente || "—"]).alignment = { wrapText: true };
  ws4.addRow([]);
  xlSection(ws4, "Análisis de competencia", 1);
  ws4.addRow([est.insight?.competencia || "—"]).alignment = { wrapText: true };
  ws4.addRow([]);
  xlSection(ws4, "Valor agregado a proponer", 1);
  (est.insight?.valor?.length ? est.insight.valor : ["—"]).forEach((v) => ws4.addRow(["• " + v]));
  ws4.addRow([]);
  xlSection(ws4, "Oportunidad de IA / automatización", 1);
  ws4.addRow([est.insight?.ia || "—"]).alignment = { wrapText: true };
  ws4.addRow([]);
  xlSection(ws4, "★ Estrategia de cierre (uso interno)", 1);
  ws4.addRow([est.insight?.estrategiaCierre || "—"]).alignment = { wrapText: true };

  return wb;
}

// Genera los 3 adjuntos (Excel, Propuesta, Inteligencia Comercial) de una cotización guardada del
// Historial, para el envío por correo. `item` ya viene normalizado (con `.scenarios`) porque
// refreshHistory() aplica normalizeEstimation() a todo lo que lee del store. No se llama a la IA para
// la prosa de la propuesta (se usa la plantilla determinística defaultProse) — evitar una llamada a
// Claude por cada cotización histórica que se reenvía, esto es documentación interna, no la versión
// final pulida para el cliente (esa se genera aparte con el botón "Propuesta").
async function buildAttachmentsForEstimation(item) {
  const country = COUNTRIES[item.country] || COUNTRIES.PE;
  const scenarioCalcs = item.scenarios.map((s) => ({ scenario: s, calc: compute(item, s) }));
  const clientSlug = (item.client || "cliente").replace(/\s+/g, "_");

  const wb = buildEstimationWorkbook(item, country, scenarioCalcs);
  const xlsxBuf = await wb.xlsx.writeBuffer();

  const proposalDoc = buildProposalDocx(item, scenarioCalcs, country, defaultProse(item));
  const intelDoc = buildIntelDocx(item, scenarioCalcs, country);

  return [
    {
      filename: `${item.code}_Excel_${clientSlug}.xlsx`,
      contentBase64: arrayBufferToBase64(xlsxBuf),
      contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    },
    {
      filename: `Propuesta_${item.code}_${clientSlug}.docx`,
      contentBase64: await Packer.toBase64String(proposalDoc),
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    },
    {
      filename: `Inteligencia_Comercial_${item.code}_${clientSlug}.docx`,
      contentBase64: await Packer.toBase64String(intelDoc),
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    },
  ];
}

/* ============================ APP ============================ */
export default function App() {
  const [est, setEst] = useState(() => defaultEstimation());
  const [history, setHistory] = useState([]);
  const [showHistory, setShowHistory] = useState(false);
  const [showResumen, setShowResumen] = useState(false);
  const [showEstrategia, setShowEstrategia] = useState(false);
  const [showParametros, setShowParametros] = useState(false);
  const [showIntel, setShowIntel] = useState(false);
  const [showSensibilidad, setShowSensibilidad] = useState(false);
  const [showChat, setShowChat] = useState(false);
  const [historySearch, setHistorySearch] = useState("");
  const [historyCountryFilter, setHistoryCountryFilter] = useState("");
  const [historyTab, setHistoryTab] = useState("pendientes");
  const [selectedCodes, setSelectedCodes] = useState(() => new Set());
  const [sendingEmail, setSendingEmail] = useState(false);
  const [aiLoading, setAiLoading] = useState(false);
  const [aiError, setAiError] = useState("");
  // Adjuntos por escenario: cada escenario puede tener sus propios documentos fuente (ej. BBP
  // "Presencial" vs. acta "Remoto"), para que el análisis con IA use solo los de cada uno.
  const [attachmentsByScenario, setAttachmentsByScenario] = useState({});
  const [parsing, setParsing] = useState(false);
  const [proposalLoading, setProposalLoading] = useState(false);
  const [proposalHTML, setProposalHTML] = useState("");
  const [proposalProse, setProposalProse] = useState(null);
  const [showProposal, setShowProposal] = useState(false);
  const [intelTab, setIntelTab] = useState("competencia");
  const [chatMessages, setChatMessages] = useState([]);
  const proposalFrame = useRef(null);
  const fileRef = useRef(null);
  const [dragOver, setDragOver] = useState(false);
  const [toast, setToast] = useState("");
  const [countryMenuOpen, setCountryMenuOpen] = useState(false);
  const countryMenuRef = useRef(null);
  useEffect(() => {
    if (!countryMenuOpen) return;
    const onClickOutside = (e) => { if (countryMenuRef.current && !countryMenuRef.current.contains(e.target)) setCountryMenuOpen(false); };
    document.addEventListener("mousedown", onClickOutside);
    return () => document.removeEventListener("mousedown", onClickOutside);
  }, [countryMenuOpen]);
  const [downloadMenuOpen, setDownloadMenuOpen] = useState(false);
  const downloadMenuRef = useRef(null);
  useEffect(() => {
    if (!downloadMenuOpen) return;
    const onClickOutside = (e) => { if (downloadMenuRef.current && !downloadMenuRef.current.contains(e.target)) setDownloadMenuOpen(false); };
    document.addEventListener("mousedown", onClickOutside);
    return () => document.removeEventListener("mousedown", onClickOutside);
  }, [downloadMenuOpen]);
  const countrySelected = !!COUNTRIES[est.country];
  const country = COUNTRIES[est.country] || COUNTRIES.PE;
  // Ecuador cotiza en USD (economía dolarizada): mostrar una columna/línea "en moneda local" que es
  // literalmente el mismo número en USD es puro ruido — solo tiene sentido cuando difieren (ej. Perú/PEN).
  const hasLocalCurrency = country.cur !== "USD";
  const activeScenario = useMemo(() => est.scenarios.find((s) => s.id === est.activeScenarioId) || est.scenarios[0], [est.scenarios, est.activeScenarioId]);
  // Adjuntos del escenario que está activo en el composer/pestañas (donde se suben los documentos).
  const attachments = attachmentsByScenario[est.activeScenarioId] || [];
  const calc = useMemo(() => compute(est, activeScenario), [est, activeScenario]);
  // Un cálculo por escenario, para las pestañas de comparación y las exportaciones (Excel/Word).
  const scenarioCalcs = useMemo(() => est.scenarios.map((s) => ({ scenario: s, calc: compute(est, s) })), [est]);

  useEffect(() => { refreshHistory(); }, []);
  const flash = (m) => { setToast(m); setTimeout(() => setToast(""), 2600); };

  async function refreshHistory() {
    const keys = await store.list();
    const items = [];
    for (const k of keys) {
      const v = await store.get(k);
      if (v) { try { items.push(normalizeEstimation(JSON.parse(v))); } catch (e) {} }
    }
    items.sort((a, b) => (b.savedAt || 0) - (a.savedAt || 0));
    setHistory(items);
  }

  const up = (patch) => setEst((e) => ({ ...e, ...patch }));
  // Actualiza solo el escenario activo (equipo/entregables/cronograma/descuento) sin tocar lo compartido.
  const upScenario = (patch) => setEst((e) => ({
    ...e,
    scenarios: e.scenarios.map((s) => (s.id === e.activeScenarioId ? { ...s, ...patch } : s)),
  }));

  // Vía manual para crear una segunda opción (ej. "Remoto") sin depender de que la IA la entienda a la
  // primera: duplica el escenario activo con su equipo/entregables/cronograma como punto de partida.
  function addScenario() {
    if (est.scenarios.length >= MAX_SCENARIOS) { flash(`Máximo ${MAX_SCENARIOS} escenarios por estimación`); return; }
    const src = activeScenario;
    const clone = {
      ...src,
      id: crypto.randomUUID(),
      name: `${src.name} (copia)`,
      team: src.team.map((t) => ({ ...t, id: crypto.randomUUID() })),
      deliverables: src.deliverables.map((d) => ({ ...d, id: crypto.randomUUID() })),
      schedule: src.schedule.map((s) => ({ ...s, id: crypto.randomUUID() })),
      reasoning: "",
    };
    setEst((e) => ({ ...e, scenarios: [...e.scenarios, clone], activeScenarioId: clone.id }));
  }

  function renameScenario(id, name) {
    setEst((e) => ({ ...e, scenarios: e.scenarios.map((s) => (s.id === id ? { ...s, name } : s)) }));
  }

  function removeScenario(id) {
    if (est.scenarios.length <= 1) { flash("Debe quedar al menos un escenario"); return; }
    setEst((e) => {
      const scenarios = e.scenarios.filter((s) => s.id !== id);
      const activeScenarioId = e.activeScenarioId === id ? scenarios[0].id : e.activeScenarioId;
      return { ...e, scenarios, activeScenarioId };
    });
  }

  function changeCountry(key) {
    const c = COUNTRIES[key];
    setEst((e) => ({ ...e, country: key, fx: c.fx, rates: cloneRates(c.rates) }));
  }

  function newEstimation() {
    setEst(defaultEstimation(est.country, newReqCode(history.map((h) => h.code))));
    setAttachmentsByScenario({});
    setAiError("");
    setChatMessages([]);
    flash("Nueva estimación iniciada");
  }

  async function saveEstimation() {
    if (!countrySelected) { flash("Selecciona el país primero"); return; }
    if (!est.client.trim()) { flash("Falta el nombre del cliente"); return; }
    const savedAt = Date.now();
    const totals = scenarioCalcs.map(({ scenario, calc: c }) => ({ name: scenario.name, PVfinal: c.PVfinal, rent: c.rent }));
    const attachmentNames = [...new Set(Object.values(attachmentsByScenario).flat().filter((a) => a.kind !== "error").map((a) => a.name))];
    const payload = { ...est, savedAt, expiresAt: savedAt + (est.validDays || 30) * 86400000, attachmentNames, totals };
    await store.set("est:" + est.code, JSON.stringify(payload));
    await refreshHistory();
    flash("Guardado " + est.code);
  }

  function loadEstimation(item) {
    setEst(normalizeEstimation({ ...item })); setAttachmentsByScenario({}); setShowHistory(false); setChatMessages([]);
    flash("Cargada " + item.code);
  }

  async function deleteEstimation(code) {
    await store.del("est:" + code); await refreshHistory();
  }

  const clientCount = (name) =>
    history.filter((h) => (h.client || "").trim().toLowerCase() === (name || "").trim().toLowerCase()).length;

  function toggleSelected(code) {
    setSelectedCodes((prev) => {
      const next = new Set(prev);
      if (next.has(code)) next.delete(code); else next.add(code);
      return next;
    });
  }

  const SEND_TO = "lcondori@grupoebim.com";

  // Junta Excel + Propuesta + Inteligencia Comercial de cada cotización seleccionada en UN solo correo
  // (decisión del usuario: un envío por selección, no uno por cotización) y las marca como enviadas.
  async function sendSelectedByEmail() {
    const items = history.filter((h) => selectedCodes.has(h.code));
    if (!items.length) return;
    setSendingEmail(true);
    try {
      const attachmentGroups = await Promise.all(items.map((item) => buildAttachmentsForEstimation(item)));
      const attachments = attachmentGroups.flat();
      const bodyText = `Se adjuntan ${items.length} cotización(es) de GRUPO EBIM:\n\n` +
        items.map((it) => `- ${it.code} · ${it.client || "(sin cliente)"} · ${it.project || "(sin proyecto)"}`).join("\n") +
        `\n\nCada una incluye: Excel de costos, Propuesta comercial (Word) e Informe de Inteligencia Comercial (Word).`;
      const subject = `Cotizaciones EBIM (${items.length}): ${items.map((it) => it.client || it.code).join(", ")}`;

      const { data, error: fnError } = await supabase.functions.invoke("send-report-email", {
        body: { to: SEND_TO, subject, bodyText, attachments },
      });
      if (fnError || data?.error) throw new Error(data?.error || fnError.message);

      const sentAt = Date.now();
      for (const item of items) {
        await store.set("est:" + item.code, JSON.stringify({ ...item, sentAt, sentTo: SEND_TO }));
      }
      await refreshHistory();
      setSelectedCodes(new Set());
      setHistoryTab("enviados");
      flash(`${items.length} cotización(es) enviada(s) a ${SEND_TO}`);
    } catch (err) {
      flash("No se pudo enviar el correo (" + (err?.message || "error") + ")");
    } finally {
      setSendingEmail(false);
    }
  }

  /* ---------- Adjuntos: BBP, transcripciones Teams, actas, PDF, imágenes ---------- */
  const TEXT_EXT = ["txt", "md", "markdown", "vtt", "srt", "csv", "tsv", "json", "log", "html"];
  const IMG_MIME = { png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", webp: "image/webp" };
  const MAX_TEXT = 60000; // tope de caracteres por documento que se envía a la IA

  const toBase64 = (file) =>
    new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(String(r.result).split(",")[1]); r.onerror = rej; r.readAsDataURL(file); });

  // Limpia las marcas de tiempo y hablantes repetidos de transcripciones VTT/SRT para reducir ruido
  function cleanTranscript(t) {
    return t
      .replace(/^WEBVTT.*$/gim, "")
      .replace(/^\d+\s*$/gm, "")
      .replace(/\d{2}:\d{2}:\d{2}[.,]\d{3}\s*-->\s*\d{2}:\d{2}:\d{2}[.,]\d{3}.*$/gm, "")
      .replace(/\n{3,}/g, "\n\n")
      .trim();
  }

  async function addFiles(fileList) {
    const files = Array.from(fileList || []);
    if (!files.length) return;
    setParsing(true); setAiError("");
    const out = [];
    for (const f of files) {
      const ext = (f.name.split(".").pop() || "").toLowerCase();
      try {
        if (ext === "docx") {
          const buf = await f.arrayBuffer();
          const { value } = await mammoth.extractRawText({ arrayBuffer: buf });
          out.push({ id: crypto.randomUUID(), name: f.name, kind: "docx", text: (value || "").slice(0, MAX_TEXT) });
        } else if (ext === "pdf") {
          out.push({ id: crypto.randomUUID(), name: f.name, kind: "pdf", base64: await toBase64(f), mediaType: "application/pdf" });
        } else if (IMG_MIME[ext]) {
          out.push({ id: crypto.randomUUID(), name: f.name, kind: "image", base64: await toBase64(f), mediaType: IMG_MIME[ext] });
        } else if (TEXT_EXT.includes(ext) || f.type.startsWith("text")) {
          let txt = await f.text();
          if (ext === "vtt" || ext === "srt") txt = cleanTranscript(txt);
          out.push({ id: crypto.randomUUID(), name: f.name, kind: "text", text: txt.slice(0, MAX_TEXT) });
        } else {
          out.push({ id: crypto.randomUUID(), name: f.name, kind: "error", text: "Formato no soportado en el navegador. Pega el contenido como texto." });
        }
      } catch (err) {
        out.push({ id: crypto.randomUUID(), name: f.name, kind: "error", text: "No se pudo leer (" + (err?.message || "error") + ")" });
      }
    }
    const sid = est.activeScenarioId;
    setAttachmentsByScenario((m) => ({ ...m, [sid]: [...(m[sid] || []), ...out] }));
    setParsing(false);
    flash(out.length + " documento(s) adjuntado(s)" + (est.scenarios.length > 1 ? ` a "${activeScenario.name}"` : ""));
  }

  const removeAttachment = (id) => {
    const sid = est.activeScenarioId;
    setAttachmentsByScenario((m) => ({ ...m, [sid]: (m[sid] || []).filter((x) => x.id !== id) }));
  };

  /* ---------- IA: generar borrador ---------- */
  // Convierte un escenario devuelto por la IA (forma "sc" del JSON) a la forma que usa la app,
  // sin asignarle id — cada llamador decide si es un escenario nuevo o si reemplaza uno existente.
  function scenarioFromAI(sc, fallbackName) {
    let acc = 0;
    const team = (sc.equipo || []).map((t) => ({ id: crypto.randomUUID(), perfil: t.perfil, tier: ["senior", "semi", "analista"].includes(t.tier) ? t.tier : "semi", rol: t.rol || "" }));
    return {
      name: sc.nombre || fallbackName,
      team,
      deliverables: (sc.entregables || []).map((d) => {
        const dur = Math.max(1, Math.round(+d.semanas || 1));
        const horasArr = Array.isArray(d.horas) ? d.horas : [];
        const hours = {};
        team.forEach((t, idx) => { hours[t.id] = +horasArr[idx] || 0; });
        const item = { id: crypto.randomUUID(), name: d.nombre, hours, start: acc, dur };
        acc += dur;
        return item;
      }),
      schedule: (sc.cronograma && sc.cronograma.length ? sc.cronograma : defaultSchedule().map((s) => ({ hito: s.hito, pct: s.pct }))).map((s) => ({ id: crypto.randomUUID(), hito: s.hito, pct: +s.pct || 0 })),
      discount: 0,
      reasoning: sc.razonPrecio || "",
    };
  }

  async function generateDraft() {
    // Adjuntos propios de cada escenario (no globales): si 2+ escenarios tienen sus propios
    // documentos, se analiza cada uno POR SEPARADO en la misma llamada ("modo multi-escenario").
    const attForScenario = (id) => (attachmentsByScenario[id] || []).filter((a) => a.kind !== "error");
    const allUsable = Object.values(attachmentsByScenario).flat().filter((a) => a.kind !== "error");
    if (!est.context.trim() && allUsable.length === 0) { setAiError("Pega el contexto o adjunta al menos un documento (BBP, acta, transcripción…)."); return; }
    const scenariosWithDocs = est.scenarios.filter((s) => attForScenario(s.id).length > 0);
    const perScenarioMode = scenariosWithDocs.length >= 2;

    const instructionText = est.context.trim() || (perScenarioMode
      ? `Analiza documentos por separado para ${scenariosWithDocs.length} escenarios (${scenariosWithDocs.map((s) => s.name).join(", ")}).`
      : `Analiza ${allUsable.length} documento(s) adjunto(s).`);
    setChatMessages((m) => [...m, { id: crypto.randomUUID(), role: "user", text: instructionText, ts: Date.now() }]);
    setAiLoading(true); setAiError("");
    const c = COUNTRIES[est.country];
    const hist = history.slice(0, 10).map((h) => `- ${h.project || "(sin título)"} [${h.client}, ${COUNTRIES[h.country]?.name || h.country}] ${h.insight?.ia ? "· IA: " + h.insight.ia.slice(0, 80) : ""}`).join("\n") || "- (sin historial aún)";
    const sys = `Eres la parte operativo-comercial senior de tecnología de GRUPO EBIM: actúas como ESTRATEGA COMERCIAL que arma propuestas ganadoras, rentables y a la medida del mercado. EBIM opera en PERÚ y ECUADOR (todo en USD); enfócate en esos dos mercados. Conoces el negocio:
${EBIM_PROFILE}
${MARKET_INTEL}
Aprendes del historial de propuestas de EBIM (reutiliza patrones, perfiles, precios y valor agregado que funcionaron):
${hist}

En CADA estimación SIEMPRE debes:
1) Investigar al CLIENTE ESPECÍFICO con web_search (a qué se dedica, tamaño, contexto). El análisis de competencia general de Perú/Ecuador ya está resuelto en INTELIGENCIA DE MERCADO arriba — no lo vuelvas a investigar, solo aplícalo y adapta la mención de diferenciación al segmento real del cliente (SAP, Odoo, IA, o combinación).
2) Recomendar valor agregado concreto que EBIM puede sumar (entregables o servicios extra que justifican el precio y abren cross-sell).
3) Identificar dónde vender IA/automatización si aplica (copilotos, chatbots, automatización de procesos, analítica/BI, agentes); si no aplica, dilo.
4) Dar una ESTRATEGIA DE CIERRE: cómo defender y posicionar el precio que resulte de TUS PROPIAS horas/entregables (nunca inventes ni menciones un monto distinto — no conoces el precio final exacto en USD porque lo calcula la app a partir de las horas que tú propongas; habla en términos relativos: "sostener el precio sin descuento salvo...", "enfatizar X para justificar el valor", no en cifras absolutas), qué enfatizar ante este cliente, y riesgos a cuidar.
5) Si el usuario pide EXPLÍCITAMENTE comparar alternativas (p. ej. "presencial vs remoto", "con IA y sin IA", "una versión de 1 semana y otra completa"), genera un ESCENARIO por cada alternativa pedida (máximo 3), cada uno con nombre corto y diferenciado, su propio equipo/entregables/cronograma y su propia razón de precio. Si el usuario NO pide comparar alternativas, genera un solo escenario.

El usuario de EBIM puede incluir INSTRUCCIONES o consideraciones (p. ej. "agrega un consultor de seguridad", "hazlo más competitivo", "incluye fase 2"): aplícalas. Si recibes una PROPUESTA ACTUAL en JSON (con uno o más escenarios), modifica ESA base y conserva lo no afectado, en lugar de empezar de cero — si el usuario no pide agregar/quitar escenarios, conserva la misma cantidad y nombres que ya existían.

Tarea: a partir del CONTEXTO y de los DOCUMENTOS ADJUNTOS (BBP/Business Blueprint, transcripciones de Teams, actas, propuestas, diagramas): (a) DETECTA automáticamente el tipo de servicio/proyecto, el CLIENTE y el PAÍS del cliente; (b) propón el equipo y el desglose de entregables con horas por perfil para ese país, para cada escenario que corresponda. Si hay un BBP, deriva los entregables de los procesos y gaps; si hay transcripciones, extrae alcance, supuestos y compromisos. Usa web_search para validar cliente, competencia o tarifas locales. Devuelve EXCLUSIVAMENTE un objeto JSON válido (sin texto ni markdown) con esta forma exacta y conciso:
{"cliente":"string (nombre del cliente detectado, o '')","pais":"string (PE o EC; o el nombre del país)","tipoProyecto":"string","resumenRequerimiento":"string (RESUMEN EN LENGUAJE SIMPLE de qué necesita el cliente y por qué — no quién es el cliente ni de qué se dedica, sino CUÁL ES EL PROBLEMA/NECESIDAD que trae y qué se le va a entregar en términos llanos, como si se lo explicaras a alguien que no leyó los documentos; 3-4 frases, sin jerga de venta)","hechosClave":["string (cifras/datos LITERALES tomados de los documentos que usaste para dimensionar el alcance — ej. 'Catálogo de 15,000 SKUs iniciales', '5 usuarios en panel admin' — copia el número exacto del documento, nunca lo redondees ni lo inventes; 3-6 items)"],"perfilCliente":"string (investiga al cliente con web_search: a qué se dedica, industria, tamaño y qué ofrecerle; 2-3 frases)","escenarios":[{"nombre":"string (corto y diferenciado, ej. 'Presencial', 'Remoto — 1 semana'; si el usuario NO pidió comparar alternativas, usa el nombre del proyecto o 'Propuesta única')","equipo":[{"perfil":"string","tier":"senior|semi|analista","rol":"string"}],"entregables":[{"nombre":"string","horas":[number,"..."],"semanas":number}],"cronograma":[{"hito":"string","pct":number}],"razonPrecio":"string (2-3 frases: qué impulsa el precio de ESTE escenario — horas/entregables clave, supuestos de modalidad, riesgos; sin inventar montos en USD)"}],"analisisCompetencia":"string (2-3 frases)","valorAgregado":["string","string"],"oportunidadIA":"string (1-2 frases; o 'No aplica')","estrategiaCierre":"string (cómo defender el precio resultante en términos relativos, sin inventar un monto en USD; qué enfatizar y riesgos; 2-3 frases)"}
Reglas: "horas" es un array de números EN EL MISMO ORDEN Y LARGO que el array "equipo" de ese escenario — horas[i] son las horas que le tocan a equipo[i] en ese entregable (ej. si equipo=[Arquitecto,PM,Analista], horas=[8,4,2] significa 8h Arquitecto + 4h PM + 2h Analista). Nunca agrupes por tier: cada integrante del equipo tiene su propia posición en "horas" aunque comparta tier con otro. semanas=duración del entregable en semanas (entero ≥1). El país solo puede ser Perú (PE) o Ecuador (EC). Máximo 6 entregables por escenario, numéralos. El cronograma de cada escenario suma pct=1.0 (típico 0.35/0.35/0.30). "escenarios" tiene 1 elemento por defecto; solo tiene más de uno si el usuario pidió explícitamente comparar alternativas (máx. 3). Sé breve para no exceder el límite de tokens. Si el mensaje incluye una PROPUESTA ACTUAL, trata el texto del usuario como INSTRUCCIONES DE MODIFICACIÓN: aplícalas sobre esa propuesta y conserva todo lo que el usuario no pida cambiar (devuelve igualmente el JSON completo).
CONSISTENCIA (importante): para el mismo alcance/contexto, dos corridas NO deberían producir precios finales muy distintos entre sí — eso rompe la confianza del comercial en la herramienta. Antes de fijar las horas, estima primero la complejidad y el tamaño real del alcance (número de procesos/módulos/integraciones descritos, cantidad de usuarios, plazos mencionados) y deriva las horas de ahí de forma metódica, no de una sensación distinta cada vez. Usa como referencia órdenes de magnitud típicos de EBIM para consultoría TI LatAm (un diagnóstico acotado ronda 80-150h totales; una implementación mediana con desarrollo, 400-900h; un programa multi-módulo, 1000h+) y ajusta según lo que el contexto realmente pida, no por defecto.
ALCANCE (cuando los documentos son transcripciones/actas, no un BBP cerrado): incluye en las horas SOLO lo que quedó como compromiso o requerimiento explícito en la conversación; si algo se menciona como idea, posibilidad futura o "fase 2/nice to have", NO lo sumes al alcance de esta cotización — mencionarlo como oportunidad de venta futura en valorAgregado si aplica, pero no lo cotices. Esto es clave para que el alcance (y por lo tanto el precio) no varíe según cuánto de la charla decidas incluir cada vez.
CALIBRACIÓN DE HORAS POR TIPO DE TRABAJO (crítico — sesgo conocido a corregir): las estimaciones de LLM sistemáticamente SUBESTIMAN el trabajo "de riesgo" (integraciones, migraciones, trabajo sobre sistemas legados/de terceros) frente al trabajo "de construcción limpia" (features nuevas desde cero) — y subestiman las pruebas. Corrige eso así:
- Entregables de INTEGRACIÓN con sistemas legados/terceros (ERPs viejos como AX 2012/SAP ECC/Oracle antiguos, APIs de terceros sin documentación clara, migraciones de datos): añade un colchón de 30-50% sobre tu primera estimación "limpia" — casi siempre hay reglas de negocio no documentadas, calidad de datos irregular, o falta de API moderna que obliga a desarrollo adicional.
- PRUEBAS/QA/UAT: no lo estimes de forma aislada — calcúlalo como 15-25% de la suma de horas de TODOS los demás entregables de desarrollo (más alto si hay IA, integraciones o arquitectura nueva que agregan superficie de riesgo; más bajo solo si el alcance es muy chico/simple). Ajusta el entregable de pruebas a ese rango antes de finalizar.
- Entregables de FEATURES NUEVAS bien acotadas (UI, lógica de negocio estándar, reportes) no necesitan colchón — ahí las estimaciones "limpias" suelen ser razonables.
AISLAMIENTO DE web_search (crítico para estabilidad de precio): equipo, entregables, horas por integrante y semanas se derivan ÚNICA Y EXCLUSIVAMENTE del CONTEXTO y los DOCUMENTOS ADJUNTOS — nunca de lo que encuentres con web_search ni de tu conocimiento general de "proyectos similares". web_search es solo para enriquecer perfilCliente, analisisCompetencia y oportunidadIA (texto cualitativo); sus resultados varían entre corridas por naturaleza, así que si dejas que influyan en el dimensionamiento del alcance, el precio final deja de ser estable. Primero fija el alcance y las horas leyendo solo los documentos; después, sin tocar esas horas, usa web_search para el análisis de mercado.
PRECISIÓN NUMÉRICA (crítico): antes de estimar horas, releé los documentos y extrae en "hechosClave" las cifras EXACTAS que mencionan (cantidad de SKUs/usuarios/sedes/registros, plazos, etc.), copiándolas TAL CUAL aparecen en el texto — nunca las redondees, aproximes ni cambies de una corrida a otra (si el documento dice "15,000 SKUs", es 15,000, no 20,000 ni 45,000). Basa las horas en esas cifras extraídas, no en una impresión general del tamaño del proyecto. Un error de cifra aquí es la causa más común de que el precio final salte entre corridas.`;
    const sysFinal = perScenarioMode
      ? sys + `\n\nMODO MULTI-ESCENARIO CON DOCUMENTOS SEPARADOS (anula el punto 5 de arriba para esta corrida): el mensaje del usuario trae ${scenariosWithDocs.length} bloques "===== ESCENARIO: <nombre> =====", cada uno con sus propios documentos. Devuelve EXACTAMENTE un elemento en "escenarios" por bloque, EN EL MISMO ORDEN, con "nombre" igual al del bloque, derivando equipo/entregables/horas/cronograma/razonPrecio de cada uno ÚNICA Y EXCLUSIVAMENTE de los documentos de SU PROPIO bloque — nunca mezcles información entre bloques.`
      : sys;

    // Prompt de usuario: si hay documentos separados por escenario, se segmenta en bloques
    // "===== ESCENARIO: nombre =====" (uno por escenario con adjuntos propios); si no, es el
    // flujo compartido de siempre (un solo set de documentos para toda la propuesta).
    let content;
    if (perScenarioMode) {
      const header = `País sugerido: ${c.name}. Tarifas EBIM en USD/h — Senior (costo ${c.rates.senior.cost} / venta ${c.rates.senior.sale}), Semi (costo ${c.rates.semi.cost} / venta ${c.rates.semi.sale}), Analista (costo ${c.rates.analista.cost} / venta ${c.rates.analista.sale}).
CONTEXTO E INSTRUCCIONES DEL USUARIO (EBIM, aplica a todos los escenarios abajo):\n${est.context || "(ver documentos de cada escenario)"}`;
      const blocks = [{ type: "text", text: header }];
      scenariosWithDocs.forEach((s) => {
        const docs = attForScenario(s.id);
        const textDocs = docs.filter((a) => a.kind === "text" || a.kind === "docx");
        const mediaDocs = docs.filter((a) => a.kind === "pdf" || a.kind === "image");
        const hasSceneContent = s.team.some((t) => t.perfil) || s.deliverables.some((d) => d.name);
        const currentProposal = hasSceneContent
          ? `\nPROPUESTA ACTUAL DE ESTE ESCENARIO (aplica las instrucciones compartidas sobre esta base; conserva lo no afectado):\n${JSON.stringify({
              equipo: s.team.map((t) => ({ perfil: t.perfil, tier: t.tier, rol: t.rol })),
              entregables: s.deliverables.map((d) => ({ nombre: d.name, horas: s.team.map((t) => (d.hours || {})[t.id] || 0), semanas: d.dur })),
            })}`
          : "";
        let block = `\n\n===== ESCENARIO: ${s.name} =====${currentProposal}`;
        if (textDocs.length) {
          block += "\n\n----- DOCUMENTOS -----\n" + textDocs.map((a) => `\n--- ${a.name} ---\n${a.text}`).join("\n");
        }
        blocks.push({ type: "text", text: block });
        mediaDocs.forEach((a) => {
          blocks.push(a.kind === "pdf"
            ? { type: "document", source: { type: "base64", media_type: a.mediaType, data: a.base64 } }
            : { type: "image", source: { type: "base64", media_type: a.mediaType, data: a.base64 } });
        });
      });
      content = blocks;
    } else {
      const hasContent = est.scenarios.some((s) => s.team.some((t) => t.perfil) || s.deliverables.some((d) => d.name));
      const currentProposal = hasContent
        ? `\n\nPROPUESTA ACTUAL (aplica sobre esta base las modificaciones pedidas; conserva lo no afectado):\n${JSON.stringify({
            escenarios: est.scenarios.map((s) => ({
              nombre: s.name,
              equipo: s.team.map((t) => ({ perfil: t.perfil, tier: t.tier, rol: t.rol })),
              entregables: s.deliverables.map((d) => ({ nombre: d.name, horas: s.team.map((t) => (d.hours || {})[t.id] || 0), semanas: d.dur })),
            })),
          })}`
        : "";
      const usr = `País sugerido: ${c.name}. Tarifas EBIM en USD/h — Senior (costo ${c.rates.senior.cost} / venta ${c.rates.senior.sale}), Semi (costo ${c.rates.semi.cost} / venta ${c.rates.semi.sale}), Analista (costo ${c.rates.analista.cost} / venta ${c.rates.analista.sale}).
CONTEXTO E INSTRUCCIONES DEL USUARIO (EBIM):\n${est.context || "(ver documentos adjuntos)"}${currentProposal}`;

      // Texto de los documentos legibles (BBP en Word, transcripciones, actas, CSV…)
      const textDocs = allUsable.filter((a) => a.kind === "text" || a.kind === "docx");
      let docsText = "";
      if (textDocs.length) {
        docsText = "\n\n===== DOCUMENTOS ADJUNTOS =====\n" +
          textDocs.map((a) => `\n----- ${a.name} -----\n${a.text}`).join("\n");
      }
      // PDFs e imágenes se envían como bloques nativos para que la IA los lea
      const mediaBlocks = allUsable
        .filter((a) => a.kind === "pdf" || a.kind === "image")
        .map((a) => a.kind === "pdf"
          ? { type: "document", source: { type: "base64", media_type: a.mediaType, data: a.base64 } }
          : { type: "image", source: { type: "base64", media_type: a.mediaType, data: a.base64 } });

      content = [{ type: "text", text: usr + docsText }, ...mediaBlocks];
    }

    try {
      const { data, error: fnError } = await supabase.functions.invoke("claude-proxy", {
        body: {
          model: "claude-sonnet-4-6",
          max_tokens: 8000,
          temperature: 0.1,
          system: sysFinal,
          messages: [{ role: "user", content }],
          tools: [{ type: "web_search_20250305", name: "web_search" }],
        },
      });
      if (fnError) throw fnError;
      const text = (data.content || []).map((b) => (b.type === "text" ? b.text : "")).join("\n");
      const clean = text.replace(/```json/gi, "").replace(/```/g, "").trim();
      const start = clean.indexOf("{"), end = clean.lastIndexOf("}");
      const parsed = JSON.parse(clean.slice(start, end + 1));
      const detKey = resolveCountryKey(parsed.pais);
      const cc = COUNTRIES[detKey] || country;
      const ratesForCalc = detKey ? cloneRates(cc.rates) : est.rates;
      const estForCalc = { ...est, rates: ratesForCalc };
      const paisNombre = COUNTRIES[detKey]?.name || parsed.pais || "no identificado";

      let newScenarios, priceSummary, confirmText;
      if (perScenarioMode) {
        // Un resultado de la IA por bloque enviado, EN EL MISMO ORDEN — se reemplaza solo el
        // contenido de esos escenarios (conservando su id), los demás quedan intactos.
        const rawScenarios = Array.isArray(parsed.escenarios) ? parsed.escenarios : [];
        const updatedById = {};
        scenariosWithDocs.forEach((orig, i) => {
          updatedById[orig.id] = { ...scenarioFromAI(rawScenarios[i] || {}, orig.name), id: orig.id };
        });
        newScenarios = est.scenarios.map((s) => updatedById[s.id] || s);
        priceSummary = scenariosWithDocs.map((orig) => `${updatedById[orig.id].name}: ${fmtUSD(compute(estForCalc, updatedById[orig.id]).PVfinal)}`).join(" · ");
        const skipped = est.scenarios.length - scenariosWithDocs.length;
        confirmText = `Detecté: ${parsed.tipoProyecto || "servicio"} para ${parsed.cliente || "cliente no identificado"} (${paisNombre}). Analicé documentos por separado para ${scenariosWithDocs.length} escenario${scenariosWithDocs.length > 1 ? "s" : ""} — ${priceSummary}.${skipped > 0 ? ` ${skipped} escenario${skipped > 1 ? "s" : ""} sin documentos nuevos ${skipped > 1 ? "quedaron" : "quedó"} sin cambios.` : ""} Revisa el equipo, los entregables y ajusta lo que haga falta.`;
      } else {
        const rawScenarios = Array.isArray(parsed.escenarios) && parsed.escenarios.length ? parsed.escenarios.slice(0, MAX_SCENARIOS) : [{}];
        newScenarios = rawScenarios.map((sc, i) => ({ ...scenarioFromAI(sc, `Escenario ${i + 1}`), id: crypto.randomUUID() }));
        priceSummary = newScenarios.map((sc) => `${sc.name}: ${fmtUSD(compute(estForCalc, sc).PVfinal)}`).join(" · ");
        confirmText = `Detecté: ${parsed.tipoProyecto || "servicio"} para ${parsed.cliente || "cliente no identificado"} (${paisNombre}). Generé ${newScenarios.length} escenario${newScenarios.length > 1 ? "s" : ""} — ${priceSummary}. Revisa el equipo, los entregables y ajusta lo que haga falta.`;
      }

      setEst((e) => ({
        ...e,
        client: e.client || parsed.cliente || "",
        country: detKey || e.country,
        fx: detKey ? cc.fx : e.fx,
        rates: detKey ? cloneRates(cc.rates) : e.rates,
        project: e.project || parsed.tipoProyecto || "",
        context: "",
        scenarios: newScenarios,
        activeScenarioId: perScenarioMode ? e.activeScenarioId : newScenarios[0].id,
        insight: {
          resumenRequerimiento: parsed.resumenRequerimiento || "",
          perfilCliente: parsed.perfilCliente || "",
          competencia: parsed.analisisCompetencia || "",
          valor: Array.isArray(parsed.valorAgregado) ? parsed.valorAgregado : (parsed.valorAgregado ? [parsed.valorAgregado] : []),
          ia: parsed.oportunidadIA || "",
          estrategiaCierre: parsed.estrategiaCierre || "",
          hechosClave: Array.isArray(parsed.hechosClave) ? parsed.hechosClave : [],
        },
      }));
      setChatMessages((m) => [...m, { id: crypto.randomUUID(), role: "assistant", text: confirmText, ts: Date.now() }]);
      flash("Detectado y generado — revisa la propuesta");
    } catch (err) {
      const msg = "No se pudo generar el borrador automáticamente. Puedes cargar los entregables manualmente. (" + (err?.message || "error") + ")";
      setAiError(msg);
      setChatMessages((m) => [...m, { id: crypto.randomUUID(), role: "assistant", text: msg, ts: Date.now() }]);
    } finally { setAiLoading(false); }
  }

  /* ---------- Excel (con formato, vía ExcelJS) ---------- */
  async function exportExcel() {
    if (!countrySelected) { flash("Selecciona el país primero"); return; }
    const wb = buildEstimationWorkbook(est, country, scenarioCalcs);
    const buf = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = `${est.code}_${(est.client || "cliente").replace(/\s+/g, "_")}.xlsx`;
    a.click(); URL.revokeObjectURL(url);
    flash("Excel descargado");
  }


  /* ---------- Propuesta comercial al cliente ---------- */
  async function generateProposal() {
    if (!countrySelected) { flash("Selecciona el país primero"); return; }
    if (!est.client.trim() || !est.scenarios.some((sc) => sc.deliverables.some((d) => d.name))) { flash("Completa cliente y entregables primero"); return; }
    setProposalLoading(true);
    let prose = defaultProse(est);
    try {
      const c = country;
      const sys = `Eres redactor de propuestas comerciales GANADORAS de GRUPO EBIM. Escribe en español, tono profesional, cálido y persuasivo, centrado en el valor y los resultados para el cliente (no en lo técnico interno). NUNCA menciones costos internos, márgenes, descuentos ni tarifas. Devuelve EXCLUSIVAMENTE un JSON válido sin markdown: {"entendimiento":"2-3 frases que demuestran que entendemos su necesidad","enfoque":"2-3 frases sobre cómo lo abordamos","diferenciadores":["3 a 5 bullets de valor"],"porQueEbim":"2-3 frases convincentes","cierre":"1-2 frases de llamada a la acción"}.`;
      const entregablesTxt = est.scenarios.map((sc) => sc.deliverables.filter((d) => d.name).map((d) => d.name).join(", ")).join(" | ");
      const usr = `Cliente: ${est.client}. País: ${c.name}. Servicio: ${est.project}. Perfil del cliente: ${est.insight?.perfilCliente || "—"}. Entregables: ${entregablesTxt}. Valor agregado: ${(est.insight?.valor || []).join("; ")}. Oportunidad IA: ${est.insight?.ia || ""}. Contexto: ${(est.context || "").slice(0, 1500)}`;
      const { data, error: fnError } = await supabase.functions.invoke("claude-proxy", {
        body: { model: "claude-sonnet-4-6", max_tokens: 1000, system: sys, messages: [{ role: "user", content: usr }] },
      });
      if (fnError) throw fnError;
      const text = (data.content || []).map((b) => (b.type === "text" ? b.text : "")).join("\n");
      const clean = text.replace(/```json/gi, "").replace(/```/g, "").trim();
      const p = JSON.parse(clean.slice(clean.indexOf("{"), clean.lastIndexOf("}") + 1));
      prose = {
        entendimiento: p.entendimiento || prose.entendimiento,
        enfoque: p.enfoque || prose.enfoque,
        diferenciadores: Array.isArray(p.diferenciadores) && p.diferenciadores.length ? p.diferenciadores : prose.diferenciadores,
        porQueEbim: p.porQueEbim || prose.porQueEbim,
        cierre: p.cierre || prose.cierre,
      };
    } catch (err) { /* usa prose por defecto */ }
    const logoDataUrl = await getLogoDataUrl();
    setProposalProse(prose);
    setProposalHTML(buildProposalHTML(est, scenarioCalcs, country, prose, logoDataUrl));
    setShowProposal(true);
    setProposalLoading(false);
    flash("Propuesta comercial lista");
  }

  async function downloadProposalDocx() {
    const doc = buildProposalDocx(est, scenarioCalcs, country, proposalProse || defaultProse(est));
    const blob = await Packer.toBlob(doc);
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = `Propuesta_${est.code}_${(est.client || "cliente").replace(/\s+/g, "_")}.docx`;
    a.click(); URL.revokeObjectURL(url);
    flash("Propuesta Word descargada");
  }

  async function downloadIntelDocx() {
    if (!countrySelected) { flash("Selecciona el país primero"); return; }
    const doc = buildIntelDocx(est, scenarioCalcs, country);
    const blob = await Packer.toBlob(doc);
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = `Inteligencia_Comercial_${est.code}_${(est.client || "cliente").replace(/\s+/g, "_")}.docx`;
    a.click(); URL.revokeObjectURL(url);
    flash("Informe de inteligencia comercial descargado");
  }

  const discountScen = [0, 0.05, 0.10, 0.15].map((d) => scenarioRow(calc, d, country.minRent));
  const rentColor = calc.rent < 0 ? "var(--bad)" : calc.rent < country.minRent ? "var(--warn)" : "var(--ok)";

  /* ============================ RENDER ============================ */
  return (
    <div className="ebim-root">
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600&family=JetBrains+Mono:wght@500;600&display=swap');
        .ebim-root{--bg:${BRAND.bg};--surface:${BRAND.surface};--ink:${BRAND.ink};--muted:${BRAND.muted};--line:${BRAND.line};--accent:${BRAND.accent};--accent-soft:${BRAND.accentSoft};--gold:${BRAND.gold};--ok:${BRAND.ok};--warn:${BRAND.warn};--bad:${BRAND.bad};
          font-family:'Inter',system-ui,sans-serif;color:var(--ink);background:var(--bg);min-height:100vh;font-size:14px;line-height:1.45;}
        .ebim-root *{box-sizing:border-box;}
        .mono{font-family:'JetBrains Mono',ui-monospace,monospace;font-variant-numeric:tabular-nums;}
        .disp{font-family:'Space Grotesk',sans-serif;}
        .ebim-root{height:100vh;overflow:hidden;}
        .wrap{max-width:100%;margin:0;padding:16px 22px 0;display:flex;flex-direction:column;height:100vh;}
        .metabar{display:flex;align-items:center;gap:14px;flex-wrap:wrap;padding-bottom:12px;margin-bottom:12px;border-bottom:1px solid var(--line);flex-shrink:0;}
        .worklayout{display:flex;gap:18px;flex:1;min-height:0;padding-bottom:18px;}
        @media(max-width:980px){.ebim-root{height:auto;overflow:visible;}.wrap{height:auto;}.worklayout{flex-direction:column;}}
        .chatlog{flex:1;overflow-y:auto;padding:18px;display:flex;flex-direction:column;gap:12px;}
        .msg{max-width:90%;padding:9px 13px;border-radius:12px;font-size:13.5px;line-height:1.5;white-space:pre-wrap;}
        .msg.user{align-self:flex-end;background:var(--accent);color:#fff;border-bottom-right-radius:3px;}
        .msg.assistant{align-self:flex-start;background:var(--bg);border:1px solid var(--line);border-bottom-left-radius:3px;}
        .composer{border-top:1px solid var(--line);padding:12px 14px;flex-shrink:0;background:var(--surface);max-height:56vh;overflow-y:auto;}
        .composer textarea{min-height:64px;max-height:150px;}
        .result-pane{flex:1;min-width:0;overflow-y:auto;padding:0 88px;}
        @media(max-width:980px){.result-pane{padding:0 20px;}}
        .result-pane .grid{padding-bottom:24px;}
        .side-rail{position:fixed;right:0;top:50%;transform:translateY(-50%);z-index:20;display:flex;flex-direction:column;gap:6px;}
        .side-rail-left{left:0;right:auto;}
        .side-rail-left .side-rail-btn{border-right:1px solid var(--line);border-left:none;border-radius:0 10px 10px 0;box-shadow:2px 2px 10px rgba(0,0,0,.08);}
        .side-rail-btn{background:var(--surface);border:1px solid var(--line);border-right:none;border-radius:10px 0 0 10px;padding:10px;cursor:pointer;color:var(--muted);display:flex;flex-direction:column;align-items:center;gap:4px;box-shadow:-2px 2px 10px rgba(0,0,0,.08);width:100%;box-sizing:border-box;}
        .side-rail-btn:hover{color:var(--accent);border-color:var(--accent);}
        .side-rail-btn span{white-space:nowrap;font-size:9.5px;font-weight:700;text-transform:uppercase;letter-spacing:.04em;}
        .modal-overlay{position:fixed;inset:0;background:rgba(20,32,46,.55);z-index:45;display:flex;align-items:center;justify-content:center;padding:28px;}
        .modal-card{width:100%;max-height:88vh;overflow-y:auto;padding:30px 34px;box-shadow:0 24px 64px rgba(0,0,0,.32);}
        .modal-card h3{margin:0 0 22px;padding-bottom:16px;border-bottom:1px solid var(--line);font-size:13.5px;}
        .modal-card > *+*{margin-top:16px;}
        .modal-card label{margin-bottom:8px;}
        .modal-card .hint{line-height:1.55;}
        .modal-card .note{padding:14px 16px;line-height:1.6;gap:10px;}
        .modal-card textarea{padding:12px 14px;line-height:1.65;}
        .modal-card .tabs{gap:28px;margin-bottom:22px;}
        .modal-card .row4{gap:20px;}
        .modal-card .row4 label{font-size:10.5px;}
        .modal-section-label{font-family:'Space Grotesk',sans-serif;font-size:11px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:var(--accent);}
        .modal-section-label:not(:first-child){padding-top:6px;border-top:1px solid var(--line);}
        .param-field{background:var(--bg);border:1px solid var(--line);border-radius:10px;padding:12px 14px;height:100%;}
        .param-field label{margin-bottom:6px;}
        .param-field .inline-label{font-size:19px;padding:2px 4px;margin-left:-4px;}
        .param-field .inline-edit-input{font-size:19px;width:90px;}
        .modal-card table{margin-top:4px;}
        .modal-card table th{white-space:nowrap;}
        .modal-card table th,.modal-card table td{padding:13px 15px;font-size:13.5px;}
        .modal-card .tabpanel textarea{min-height:190px;}
        .chat-modal{width:100%;max-width:720px;height:82vh;display:flex;flex-direction:column;padding:22px 26px 0;box-shadow:0 24px 64px rgba(0,0,0,.32);overflow:hidden;}
        .chat-modal h3{flex-shrink:0;margin:0 0 18px;padding-bottom:14px;border-bottom:1px solid var(--line);font-size:13.5px;}
        .chat-modal .chatlog{flex:1;min-height:0;padding:0 0 18px;}
        .chat-modal .composer{border-top:1px solid var(--line);margin:0 -26px;padding:16px 26px 22px;flex-shrink:0;max-height:50vh;}
        .scenario-tabs{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:2px;align-items:center;}
        .scenario-tab{display:flex;align-items:center;gap:7px;border:1.5px solid var(--line);background:var(--surface);border-radius:10px;padding:7px 10px;cursor:pointer;}
        .scenario-tab.active{border-color:var(--accent);background:var(--accent-soft);}
        .scenario-tab input{border:none;background:transparent;padding:0;width:auto;max-width:140px;font-weight:600;font-size:13px;color:var(--ink);}
        .scenario-tab .sc-price{font-family:'JetBrains Mono',monospace;font-size:12px;color:var(--muted);white-space:nowrap;}
        .scenario-tab.active .sc-price{color:var(--accent);}
        .scenario-pick-btn{border:1.5px solid var(--line);background:var(--surface);border-radius:10px;padding:6px 11px;cursor:pointer;font-weight:600;font-size:12.5px;color:var(--muted);}
        .scenario-pick-btn:hover{border-color:var(--accent);color:var(--accent);}
        .scenario-pick-btn.active{border-color:var(--accent);background:var(--accent-soft);color:var(--accent);}
        .reasoning-box{margin-top:12px;font-size:13.5px;line-height:1.55;background:var(--accent-soft);border-radius:8px;padding:10px 12px;color:var(--ink);}
        .section-label{font-family:'Space Grotesk',sans-serif;font-size:11px;font-weight:700;letter-spacing:.12em;text-transform:uppercase;color:var(--accent);margin:10px 0 -8px 2px;scroll-margin-top:64px;}
        .section-label:first-child{margin-top:0;}
        .tabs{display:flex;gap:22px;border-bottom:1px solid var(--line);margin-bottom:18px;}
        .tab{font-family:'Space Grotesk',sans-serif;font-size:13px;font-weight:600;color:var(--muted);background:none;border:none;padding:0 0 12px;cursor:pointer;position:relative;transition:.12s;}
        .tab:hover{color:var(--ink);}
        .tab.active{color:var(--accent);}
        .tab.active::after{content:"";position:absolute;left:0;right:0;bottom:-1px;height:2px;background:var(--accent);border-radius:2px 2px 0 0;}
        .tabpanel textarea{min-height:220px;font-size:14.5px;line-height:1.6;}
        .topbar{display:flex;align-items:center;gap:14px;flex-wrap:wrap;border-bottom:1px solid var(--line);padding-bottom:16px;margin-bottom:20px;}
        .brand{display:flex;align-items:center;gap:10px;}
        .brand-client{display:block;font-family:'Space Grotesk',sans-serif;font-weight:700;font-size:32px;line-height:1.15;color:var(--ink);letter-spacing:-.4px;padding:2px 6px;}
        .brand-project{display:block;font-size:13px;font-weight:600;color:var(--muted);line-height:1.3;padding:1px 6px;}
        .reqchip{font-family:'JetBrains Mono',monospace;background:var(--ink);color:#fff;padding:6px 12px;border-radius:7px;font-weight:600;letter-spacing:.5px;}
        .btn{display:inline-flex;align-items:center;gap:7px;border:1px solid var(--line);background:var(--surface);color:var(--ink);padding:8px 13px;border-radius:8px;font-weight:600;font-size:13px;cursor:pointer;transition:.12s;}
        .btn:hover{border-color:var(--accent);color:var(--accent);}
        .btn.primary{background:var(--accent);color:#fff;border-color:var(--accent);}
        .btn.primary:hover{filter:brightness(1.08);color:#fff;}
        .btn.ghost{background:transparent;}
        .grid{display:grid;gap:20px;}
        .card{background:var(--surface);border:1px solid var(--line);border-radius:12px;padding:20px 22px;}
        .card h3{font-family:'Space Grotesk',sans-serif;font-size:12px;letter-spacing:.10em;text-transform:uppercase;color:var(--muted);margin:0 0 12px;display:flex;align-items:center;gap:8px;}
        label{display:block;font-size:11px;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.05em;margin-bottom:5px;}
        input,select,textarea{width:100%;border:1px solid var(--line);border-radius:8px;padding:8px 10px;font-size:14px;font-family:inherit;background:#fff;color:var(--ink);}
        textarea{field-sizing:content;min-height:80px;max-height:480px;overflow-y:auto;resize:vertical;}
        input:focus,select:focus,textarea:focus{outline:none;border-color:var(--accent);box-shadow:0 0 0 3px var(--accent-soft);}
        select.invalid{border-color:var(--bad);box-shadow:0 0 0 3px rgba(193,59,59,.08);}
        .num{font-family:'JetBrains Mono',monospace;text-align:right;}
        .pct-input-wrap{position:relative;display:block;}
        .pct-input-wrap input{padding-right:24px;}
        .pct-input-wrap .pct-suffix{position:absolute;right:10px;top:50%;transform:translateY(-50%);font-size:12px;color:var(--muted);font-family:'JetBrains Mono',monospace;pointer-events:none;}
        table{width:100%;border-collapse:collapse;}
        th{font-size:10.5px;text-transform:uppercase;letter-spacing:.05em;color:var(--muted);text-align:right;padding:7px 8px;border-bottom:1px solid var(--line);font-weight:600;}
        th.l,td.l{text-align:left;}
        td{padding:9px 10px;border-bottom:1px solid #EEF1F4;font-size:13px;}
        td .num{width:100%;border:1px solid transparent;background:transparent;padding:4px 6px;border-radius:6px;}
        td .num:hover{border-color:var(--line);}
        td .num:focus{background:#fff;}
        tbody tr:nth-child(even) td{background:#FAFBFC;}
        tbody tr.total td{font-weight:700;border-top:2px solid var(--ink);border-bottom:none;background:#F8FAFB;}
        .row2{display:grid;grid-template-columns:1fr 1fr;gap:14px;}
        .row3{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;}
        .row4{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;}
        @media(max-width:760px){.row2,.row3,.row4{grid-template-columns:1fr;}}
        .price{font-family:'Space Grotesk',sans-serif;font-size:46px;font-weight:700;letter-spacing:-1px;line-height:1;}
        .price-sticky{position:sticky;top:0;z-index:15;background:var(--surface);border:1px solid var(--line);border-radius:10px;padding:10px 18px;display:flex;align-items:center;gap:14px;box-shadow:0 6px 18px rgba(20,32,46,.1);}
        .price-sticky-label{font-size:11px;text-transform:uppercase;letter-spacing:.08em;color:var(--gold);font-weight:600;}
        .price-sticky-value{font-family:'Space Grotesk',sans-serif;font-size:22px;font-weight:700;letter-spacing:-.3px;color:var(--ink);}
        .price-sticky-rent{margin-left:auto;font-family:'JetBrains Mono',monospace;font-size:12px;font-weight:700;}
        .pricelabel{font-size:11px;text-transform:uppercase;letter-spacing:.12em;color:var(--gold);font-weight:600;margin-bottom:6px;}
        .gauge{height:9px;border-radius:6px;background:#EAEFF2;overflow:hidden;margin-top:14px;}
        .gauge > i{display:block;height:100%;border-radius:6px;}
        .kv{display:flex;justify-content:space-between;padding:7px 0;border-bottom:1px solid #EEF1F4;}
        .kv:last-child{border-bottom:none;}
        .kv b{font-family:'JetBrains Mono',monospace;}
        .pill{display:inline-flex;align-items:center;gap:5px;font-size:11px;font-weight:600;padding:3px 9px;border-radius:20px;}
        .pill.ok{background:#E6F5EE;color:var(--ok);} .pill.warn{background:#FBF1DD;color:var(--warn);} .pill.bad{background:#FBE6E6;color:var(--bad);}
        .iconbtn{border:none;background:transparent;color:var(--muted);cursor:pointer;padding:4px;border-radius:6px;}
        .iconbtn:hover{color:var(--bad);background:#FBE6E6;}
        .btn:focus-visible, .iconbtn:focus-visible, input:focus-visible, select:focus-visible, textarea:focus-visible{outline:2px solid var(--accent);outline-offset:2px;}
        .toast{position:fixed;bottom:22px;left:50%;transform:translateX(-50%);background:var(--ink);color:#fff;padding:11px 20px;border-radius:10px;font-weight:600;font-size:13px;z-index:50;box-shadow:0 8px 24px rgba(0,0,0,.18);}
        .drawer{position:fixed;inset:0;background:rgba(20,32,46,.4);z-index:40;display:flex;justify-content:flex-end;}
        .drawer .panel{width:min(440px,100%);background:var(--bg);height:100%;overflow:auto;padding:20px;}
        .histcard{background:#fff;border:1px solid var(--line);border-radius:10px;padding:12px 14px;margin-bottom:10px;cursor:pointer;}
        .histcard:hover{border-color:var(--accent);}
        .note{font-size:12px;color:var(--muted);background:var(--accent-soft);border-radius:8px;padding:9px 12px;display:flex;gap:8px;}
        .aiwrap{background:linear-gradient(180deg,#F3FAFA,#fff);border:1px solid var(--accent-soft);}
        .labelchip{display:inline-flex;align-items:center;gap:8px;font-size:11px;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.05em;}
        .labelchip select{width:auto;min-width:160px;padding:5px 9px;font-size:13px;font-weight:500;text-transform:none;letter-spacing:0;color:var(--ink);}
        .flag-icon{display:inline-block;flex-shrink:0;border-radius:2px;box-shadow:0 0 0 1px rgba(0,0,0,.15) inset;}
        .side-rail-item{position:relative;}
        .side-rail-btn.invalid{border-color:var(--bad);box-shadow:0 0 0 3px rgba(193,59,59,.08);color:var(--bad);}
        .country-picker-menu{position:absolute;top:calc(100% + 4px);left:0;z-index:30;background:#fff;border:1px solid var(--line);border-radius:8px;box-shadow:0 12px 32px rgba(20,32,46,.16);min-width:180px;padding:4px;}
        .country-picker-menu-rail{top:0;left:100%;margin-left:6px;}
        .country-picker-option{display:flex;align-items:center;gap:8px;width:100%;border:none;background:transparent;padding:7px 9px;border-radius:6px;font-size:13px;font-weight:500;color:var(--ink);cursor:pointer;text-align:left;}
        .country-picker-option:hover{background:var(--accent-soft);}
        .country-picker-option.active{color:var(--accent);font-weight:700;}
        .topbar-dropdown{position:relative;}
        .topbar-dropdown-menu{position:absolute;top:calc(100% + 4px);right:0;z-index:30;background:#fff;border:1px solid var(--line);border-radius:8px;box-shadow:0 12px 32px rgba(20,32,46,.16);min-width:230px;padding:4px;}
        .topbar-dropdown-option{display:flex;align-items:center;gap:8px;width:100%;border:none;background:transparent;padding:8px 10px;border-radius:6px;font-size:13px;font-weight:600;color:var(--ink);cursor:pointer;text-align:left;}
        .topbar-dropdown-option:hover{background:var(--accent-soft);}
        .inline-label{font-family:'Space Grotesk',sans-serif;font-weight:700;font-size:14px;color:var(--ink);text-transform:none;letter-spacing:0;cursor:pointer;border-radius:6px;padding:3px 6px;border:1px solid transparent;}
        .inline-label:hover{background:var(--accent-soft);color:var(--accent);}
        .inline-label.placeholder{font-weight:500;color:var(--muted);}
        .inline-label-multiline{display:block;white-space:pre-wrap;font-family:'Inter',sans-serif;font-weight:400;font-size:13px;line-height:1.55;min-height:22px;}
        .inline-edit-input{font-family:'Space Grotesk',sans-serif;font-weight:700;font-size:14px;color:var(--ink);border:1px solid var(--accent);border-radius:6px;padding:3px 6px;background:#fff;width:auto;min-width:120px;}
        .inline-edit-textarea{font-family:'Inter',sans-serif;font-weight:400;font-size:13px;line-height:1.55;width:100%;min-height:60px;resize:vertical;}
        .card.compact{padding:14px 18px;}
        .card.compact .row3{gap:12px;}
        .card.compact label{margin-bottom:3px;}
        .hint{font-size:11.5px;color:var(--muted);margin-top:4px;}
        .attach-btn{border:1px solid var(--line);}
        .composer.over{background:var(--accent-soft);}
        .filelist{display:grid;grid-template-columns:repeat(auto-fill,minmax(220px,1fr));gap:8px;margin-top:10px;}
        .filechip{display:flex;align-items:center;gap:9px;background:#fff;border:1px solid var(--line);border-radius:9px;padding:8px 10px;font-size:12.5px;}
        .filechip.err{border-color:#F0C9C9;background:#FCF1F1;color:var(--bad);}
        .filechip .iconbtn:hover{color:var(--bad);}
        .team-list{display:flex;flex-direction:column;gap:10px;}
        .team-row{border:1px solid var(--line);border-radius:10px;padding:12px 14px;background:#fff;}
        .team-row-top{display:flex;gap:10px;align-items:center;margin-bottom:8px;}
        .team-perfil{flex:1;font-weight:600;min-width:0;}
        .team-tier{width:auto;flex-shrink:0;min-width:150px;}
        .team-rol{font-size:13px;min-height:44px;}
      `}</style>

      <div className="wrap">
        {/* Topbar */}
        <div className="topbar">
          <div className="brand">
            <img src={ebimLogoCompleto} alt="EBIM" style={{ height: 36 }} />
            <div style={{ borderLeft: "1px solid var(--line)", paddingLeft: 10 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 9 }}>
                {countrySelected && <FlagIcon country={country} size={26} />}
                <InlineLabel className="brand-client" value={est.client} onChange={(v) => up({ client: v })} placeholder="Nombre del cliente" />
              </div>
              <InlineLabel className="brand-project" value={est.project} onChange={(v) => up({ project: v })} placeholder="Nombre del proyecto" />
            </div>
          </div>
          <div style={{ marginLeft: "auto", display: "flex", gap: 8, flexWrap: "wrap" }}>
            <button className="btn ghost" onClick={newEstimation}><RotateCcw size={15} /> Nueva</button>
            <button className="btn" onClick={saveEstimation}><Save size={15} /> Guardar</button>
            <div className="topbar-dropdown" ref={downloadMenuRef}>
              <button className="btn primary" onClick={() => setDownloadMenuOpen((o) => !o)}>
                <Download size={15} /> Descargar <ChevronDown size={13} />
              </button>
              {downloadMenuOpen && (
                <div className="topbar-dropdown-menu">
                  <button type="button" className="topbar-dropdown-option" onClick={() => { setDownloadMenuOpen(false); exportExcel(); }}>
                    <Download size={14} /> Excel
                  </button>
                  <button type="button" className="topbar-dropdown-option" onClick={() => { setDownloadMenuOpen(false); downloadIntelDocx(); }} title="Documento aparte: resumen del proyecto, estrategia comercial e inteligencia comercial">
                    <FileType size={14} /> Inteligencia Comercial
                  </button>
                </div>
              )}
            </div>
            <button className="btn ghost" aria-label="Cerrar sesión" title="Cerrar sesión" onClick={() => supabase.auth.signOut()}><LogOut size={15} /></button>
          </div>
        </div>

        {/* Barra de datos estructurales compartidos entre escenarios (no son parte de la conversación) */}
        {(!countrySelected || hasLocalCurrency) && (
          <div className="metabar">
            {!countrySelected && <span className="hint" style={{ color: "var(--bad)" }}><AlertTriangle size={12} style={{ display: "inline", verticalAlign: "-2px" }} /> Selecciona el país para continuar.</span>}
            {hasLocalCurrency && (
              <span className="labelchip">T.C. ({country.cur}/USD)
                <InlineLabel type="number" step="0.01" value={est.fx} onChange={(v) => up({ fx: v })} inputStyle={{ minWidth: 70, width: 70 }} />
              </span>
            )}
          </div>
        )}

        <div className="worklayout">
          {/* Panel derecho: resultado — precio primero, luego el detalle editable */}
          <section className="result-pane">
            <div className="grid">
              {/* Pestañas de escenario — solo si hay más de una opción a comparar (ej. Presencial vs. Remoto) */}
              {est.scenarios.length > 1 && (
                <div className="scenario-tabs">
                  {est.scenarios.map((s) => {
                    const scCalc = scenarioCalcs.find((x) => x.scenario.id === s.id)?.calc;
                    const isActive = s.id === est.activeScenarioId;
                    return (
                      <div key={s.id} className={"scenario-tab" + (isActive ? " active" : "")} onClick={() => up({ activeScenarioId: s.id })}>
                        <input value={s.name} onChange={(e) => renameScenario(s.id, e.target.value)} onClick={(e) => e.stopPropagation()} aria-label="Nombre del escenario" />
                        <span className="sc-price">{fmtUSD(scCalc.PVfinal)}</span>
                        <button className="iconbtn" aria-label={`Eliminar escenario ${s.name}`} title="Eliminar escenario" onClick={(e) => { e.stopPropagation(); removeScenario(s.id); }}><X size={13} /></button>
                      </div>
                    );
                  })}
                  <button className="btn ghost" onClick={addScenario} disabled={est.scenarios.length >= MAX_SCENARIOS} title="Duplica el escenario activo como punto de partida"><Plus size={14} /> Añadir escenario</button>
                </div>
              )}
              {/* Hero precio — primero lo que más importa: cuánto y por qué */}
              <div className="card">
                <div className="pricelabel">★ Precio de venta final (sin IGV) · USD{est.scenarios.length > 1 ? ` · ${activeScenario.name}` : ""}</div>
                <div style={{ display: "flex", alignItems: "baseline", justifyContent: "space-between", flexWrap: "wrap", gap: 10 }}>
                  <div className="price">{fmtUSD(calc.PVfinal)}</div>
                  <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
                    {est.client.trim() && clientCount(est.client) > 0 && (
                      <span className="hint" style={{ margin: 0 }}>Este cliente tiene <b>{clientCount(est.client)}</b> cotización(es) guardada(s).</span>
                    )}
                    <span className="reqchip">{est.code}</span>
                    <span className="labelchip">Validez (días)
                      <InlineLabel type="number" step="1" min="1" value={est.validDays ?? 30} onChange={(v) => up({ validDays: Math.max(1, v) })} inputStyle={{ minWidth: 56, width: 56 }} />
                    </span>
                  </div>
                </div>
                <div style={{ fontSize: 13, color: "var(--muted)", marginTop: 6 }} className="mono">{hasLocalCurrency ? `≈ ${fmtLocal(calc.PVfinal * est.fx, country.cur)} · ` : ""}{activeScenario.discount > 0 ? `incluye ${pct(activeScenario.discount)} descuento` : "sin descuento"}</div>
                {activeScenario.reasoning && <div className="reasoning-box"><b>Por qué este precio:</b> {activeScenario.reasoning}</div>}
                <div className="gauge"><i style={{ width: Math.max(3, Math.min(100, (calc.rent / 0.4) * 100)) + "%", background: rentColor }} /></div>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginTop: 8, fontSize: 12, gap: 10, flexWrap: "wrap" }}>
                  <span style={{ color: "var(--muted)" }}>Rentabilidad neta</span>
                  {countrySelected && <span className="hint" style={{ margin: 0 }}>{country.name}: rentabilidad mínima <b>{pct(country.minRent)}</b> · descuento máx. <b>{pct(country.maxDesc)}</b></span>}
                  <span className="mono" style={{ fontWeight: 700, color: rentColor }}>{pct(calc.rent)} {calc.rent < country.minRent ? "· bajo el mínimo" : "· ok"}</span>
                </div>
                <div style={{ marginTop: 18, borderTop: "1px solid var(--line)", paddingTop: 6 }}>
                  <div className="kv"><span>Costo operativo (país)</span><b>{fmtUSD(calc.CO)}</b></div>
                  <div className="kv"><span>Precio venta (horas)</span><b>{fmtUSD(calc.PVO)}</b></div>
                  <div className="kv"><span>Margen bruto (horas)</span><b>{pct(calc.grossMargin)}</b></div>
                  <div className="kv"><span>+ Gastos adm. (Carmen, MKT)</span><b>{fmtUSD(calc.gastos)}</b></div>
                  <div className="kv"><span>= Precio con gastos incluidos</span><b>{fmtUSD(calc.PVOconGastos)}</b></div>
                  <div className="kv"><span>− Descuento aplicado</span><b>−{fmtUSD(calc.PVOconGastos * activeScenario.discount)}</b></div>
                  <div className="kv"><span>Utilidad neta</span><b style={{ color: rentColor }}>{fmtUSD(calc.utilidad)}</b></div>
                </div>
              </div>

              {/* Barra compacta del precio: sticky, se "revela" al hacer scroll más allá del hero completo y queda fija arriba mientras se navega el detalle. */}
              {calc.PVfinal > 0 && (
                <div className="price-sticky">
                  <span className="price-sticky-label">Precio venta final{est.scenarios.length > 1 ? ` · ${activeScenario.name}` : ""}</span>
                  <span className="price-sticky-value">{fmtUSD(calc.PVfinal)}</span>
                  <span className="price-sticky-rent" style={{ color: rentColor }}>{pct(calc.rent)} rent.</span>
                </div>
              )}

          <div className="section-label">Supuestos y equipo</div>
          {/* Supuestos / tarifas */}
          <div className="card">
            <h3>Supuestos · costo país vs. precio de venta (USD/hora)</h3>
            <table>
              <thead><tr><th className="l">Perfil</th><th>Costo país</th><th>Precio venta</th><th>Margen unit.</th></tr></thead>
              <tbody>
                {TIERS.map(([key, label]) => {
                  const t = est.rates[key];
                  const m = t.sale > 0 ? (t.sale - t.cost) / t.sale : 0;
                  return (
                    <tr key={key}>
                      <td className="l">{label}</td>
                      <td><input className="num" type="number" value={t.cost} onChange={(e) => up({ rates: { ...est.rates, [key]: { ...t, cost: +e.target.value } } })} /></td>
                      <td><input className="num" type="number" value={t.sale} onChange={(e) => up({ rates: { ...est.rates, [key]: { ...t, sale: +e.target.value } } })} /></td>
                      <td className="num" style={{ color: m < 0.4 ? "var(--warn)" : "var(--ok)" }}>{pct(m)}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
            <div className="hint">El costo y el precio se cargan según el país elegido. Edítalos para afinar a la realidad de cada mercado. Todo en USD.</div>
          </div>

          {/* Equipo */}
          <div className="card">
            <h3>Equipo del proyecto</h3>
            <div className="team-list">
              {activeScenario.team.map((t, i) => (
                <div key={t.id} className="team-row">
                  <div className="team-row-top">
                    <InlineLabel className="team-perfil" value={t.perfil} placeholder="Perfil (ej. Arquitecto TI)"
                      onChange={(v) => { const team = [...activeScenario.team]; team[i] = { ...t, perfil: v }; upScenario({ team }); }} />
                    <select className="team-tier" value={t.tier} title={`Tarifa: costo ${fmtUSD(est.rates[t.tier].cost)}/h · venta ${fmtUSD(est.rates[t.tier].sale)}/h`}
                      onChange={(e) => { const team = [...activeScenario.team]; team[i] = { ...t, tier: e.target.value }; upScenario({ team }); }}>
                      <option value="senior">Senior · {fmtUSD(est.rates.senior.sale)}/h</option>
                      <option value="semi">Semi · {fmtUSD(est.rates.semi.sale)}/h</option>
                      <option value="analista">Analista · {fmtUSD(est.rates.analista.sale)}/h</option>
                    </select>
                    <button className="iconbtn" aria-label={`Quitar perfil ${t.perfil || ""}`} title="Quitar perfil" onClick={() => upScenario({ team: activeScenario.team.filter((x) => x.id !== t.id) })}><Trash2 size={15} /></button>
                  </div>
                  <InlineLabel className="team-rol" multiline value={t.rol} placeholder="Rol en el proyecto — qué hace este perfil en la estimación"
                    onChange={(v) => { const team = [...activeScenario.team]; team[i] = { ...t, rol: v }; upScenario({ team }); }} />
                </div>
              ))}
            </div>
            <button className="btn ghost" style={{ marginTop: 10 }} onClick={() => upScenario({ team: [...activeScenario.team, { id: crypto.randomUUID(), perfil: "", tier: "semi", rol: "" }] })}><Plus size={15} /> Añadir perfil</button>
          </div>

          {/* Entregables */}
          <div className="card">
            <h3>Detalle de horas y costos por entregable</h3>
            <div className="hint" style={{ marginBottom: 8 }}>Una columna por cada integrante de "Equipo del proyecto" — la tarifa de cada uno depende de su tier (ver tooltip).</div>
            <div style={{ overflowX: "auto" }}>
              <table style={{ minWidth: 460 + activeScenario.team.length * 84 }}>
                <thead><tr>
                  <th className="l" style={{ minWidth: 200 }}>Entregable / actividad</th>
                  {activeScenario.team.map((t) => (
                    <th key={t.id} style={{ width: 84 }} title={`${t.perfil || "(sin nombre)"} — ${TIERS.find(([k]) => k === t.tier)?.[1] || t.tier}. Tarifa actual: costo ${fmtUSD(est.rates[t.tier].cost)}/h · venta ${fmtUSD(est.rates[t.tier].sale)}/h`}>{t.perfil || "(sin nombre)"}</th>
                  ))}
                  <th style={{ width: 70 }} title="Suma de horas de todo el equipo para este entregable.">Total hrs</th>
                  <th style={{ width: 90 }} title="Costo real para EBIM: horas de cada integrante × su tarifa de costo país.">Costo USD</th>
                  <th style={{ width: 90 }} title="Precio de venta competitivo: horas de cada integrante × su tarifa de venta, antes de gastos Adm/Comercial/MKT y del descuento comercial.">Venta USD</th>
                  <th style={{ width: 36 }}></th>
                </tr></thead>
                <tbody>
                  {calc.rows.map((d, i) => (
                    <tr key={d.id}>
                      <td className="l"><InlineLabel value={d.name} placeholder="Describe el entregable" onChange={(v) => { const dd = [...activeScenario.deliverables]; dd[i] = { ...activeScenario.deliverables[i], name: v }; upScenario({ deliverables: dd }); }} /></td>
                      {activeScenario.team.map((t) => (
                        <td key={t.id}><input className="num" type="number" value={activeScenario.deliverables[i].hours?.[t.id] || ""} onChange={(e) => {
                          const dd = [...activeScenario.deliverables];
                          dd[i] = { ...dd[i], hours: { ...dd[i].hours, [t.id]: +e.target.value } };
                          upScenario({ deliverables: dd });
                        }} /></td>
                      ))}
                      <td className="num">{d.hrs}</td>
                      <td className="num" style={{ color: "var(--muted)" }}>{fmtUSD(d.cost)}</td>
                      <td className="num">{fmtUSD(d.sale)}</td>
                      <td><button className="iconbtn" aria-label={`Quitar entregable ${d.name || ""}`} title="Quitar entregable" onClick={() => upScenario({ deliverables: activeScenario.deliverables.filter((x) => x.id !== d.id) })}><Trash2 size={15} /></button></td>
                    </tr>
                  ))}
                  <tr className="total">
                    <td className="l">TOTAL</td>
                    {calc.totalByTeam.map((t) => <td key={t.id} className="num">{t.hrs}</td>)}
                    <td className="num">{calc.totalHrs}</td><td className="num" style={{ color: "var(--muted)" }}>{fmtUSD(calc.CO)}</td><td className="num">{fmtUSD(calc.PVO)}</td><td></td>
                  </tr>
                </tbody>
              </table>
            </div>
            <button className="btn ghost" style={{ marginTop: 10 }} onClick={() => upScenario({ deliverables: [...activeScenario.deliverables, blankDeliverable()] })}><Plus size={15} /> Añadir entregable</button>
          </div>

          <div className="section-label">Planificación</div>
          {/* Cronograma de trabajo (Gantt) */}
          <div className="card">
            <h3>Cronograma de trabajo · Gantt por entregable</h3>
            <div style={{ overflowX: "auto" }}>
              <table style={{ minWidth: 520 }}>
                <thead>
                  <tr>
                    <th className="l" style={{ minWidth: 200 }}>Entregable</th>
                    <th style={{ width: 70 }}>Inicio</th>
                    <th style={{ width: 70 }}>Sem.</th>
                    <th className="l">Línea de tiempo (semanas 1–{calc.weeksTotal})</th>
                  </tr>
                </thead>
                <tbody>
                  {activeScenario.deliverables.map((d, i) => {
                    const start = d.start || 0, dur = d.dur || 1;
                    return (
                      <tr key={d.id}>
                        <td className="l" style={{ fontSize: 12.5 }}>{d.name || <span style={{ color: "var(--muted)" }}>Entregable {i + 1}</span>}</td>
                        <td><input className="num" type="number" min="0" value={start} onChange={(e) => { const dd = [...activeScenario.deliverables]; dd[i] = { ...d, start: Math.max(0, +e.target.value) }; upScenario({ deliverables: dd }); }} /></td>
                        <td><input className="num" type="number" min="1" value={dur} onChange={(e) => { const dd = [...activeScenario.deliverables]; dd[i] = { ...d, dur: Math.max(1, +e.target.value) }; upScenario({ deliverables: dd }); }} /></td>
                        <td>
                          <div style={{ position: "relative", height: 20, background: "#F1F4F7", borderRadius: 5 }}>
                            <div style={{ position: "absolute", left: (start / calc.weeksTotal) * 100 + "%", width: (dur / calc.weeksTotal) * 100 + "%", top: 0, bottom: 0, background: "var(--accent)", borderRadius: 5, display: "flex", alignItems: "center", justifyContent: "center", color: "#fff", fontSize: 10, fontWeight: 600 }}>
                              {dur}s
                            </div>
                          </div>
                        </td>
                      </tr>
                    );
                  })}
                  <tr className="total">
                    <td className="l">DURACIÓN TOTAL</td><td></td>
                    <td className="num">{calc.weeksTotal}</td>
                    <td className="l mono" style={{ fontWeight: 700 }}>{calc.weeksTotal} semana{calc.weeksTotal > 1 ? "s" : ""}</td>
                  </tr>
                </tbody>
              </table>
            </div>
            <div className="hint">La IA propone la duración de cada entregable y los encadena. Edita "Inicio" para solapar fases (trabajo en paralelo) y "Sem." para ajustar la duración.</div>
          </div>

          {/* Cronograma de pagos */}
          <div className="card">
            <h3>Cronograma de pagos</h3>
            <table>
              <thead><tr><th className="l" style={{ width: 40 }}>N°</th><th className="l">Hito</th><th style={{ width: 110 }}>% del total</th><th>Importe USD</th>{hasLocalCurrency && <th>Importe {country.cur}</th>}<th style={{ width: 36 }}></th></tr></thead>
              <tbody>
                {activeScenario.schedule.map((s, i) => (
                  <tr key={s.id}>
                    <td className="l mono">{i + 1}</td>
                    <td className="l"><InlineLabel value={s.hito} placeholder="Nombre del hito" onChange={(v) => { const sc = [...activeScenario.schedule]; sc[i] = { ...s, hito: v }; upScenario({ schedule: sc }); }} /></td>
                    <td><PctInput value={s.pct} decimals={0} step={1} onChange={(v) => { const sc = [...activeScenario.schedule]; sc[i] = { ...s, pct: v }; upScenario({ schedule: sc }); }} /></td>
                    <td className="num">{fmtUSD(calc.PVfinal * s.pct)}</td>
                    {hasLocalCurrency && <td className="num">{fmtLocal(calc.PVfinal * s.pct * est.fx, country.cur)}</td>}
                    <td><button className="iconbtn" aria-label={`Quitar hito ${s.hito || ""}`} title="Quitar hito" onClick={() => upScenario({ schedule: activeScenario.schedule.filter((x) => x.id !== s.id) })}><Trash2 size={15} /></button></td>
                  </tr>
                ))}
                <tr className="total">
                  <td></td><td className="l">TOTAL</td>
                  <td className="num" style={{ color: Math.abs(activeScenario.schedule.reduce((a, b) => a + b.pct, 0) - 1) > 0.001 ? "var(--bad)" : "inherit" }}>{(activeScenario.schedule.reduce((a, b) => a + b.pct, 0) * 100).toFixed(0)}%</td>
                  <td className="num">{fmtUSD(calc.PVfinal * activeScenario.schedule.reduce((a, b) => a + b.pct, 0))}</td>
                  {hasLocalCurrency && <td className="num">{fmtLocal(calc.PVfinal * est.fx * activeScenario.schedule.reduce((a, b) => a + b.pct, 0), country.cur)}</td>}<td></td>
                </tr>
              </tbody>
            </table>
            <button className="btn ghost" style={{ marginTop: 10 }} onClick={() => upScenario({ schedule: [...activeScenario.schedule, { id: crypto.randomUUID(), hito: "", pct: 0 }] })}><Plus size={15} /> Añadir hito</button>
          </div>

            </div>
          </section>
        </div>
      </div>

      {/* Menú lateral izquierdo: el prompt/chat con la IA, se abre al medio en vez de ocupar una columna fija */}
      <div className="side-rail side-rail-left">
        <div className="side-rail-item" ref={countryMenuRef}>
          <button className={"side-rail-btn" + (countrySelected ? "" : " invalid")} onClick={() => setCountryMenuOpen((o) => !o)} title="Seleccionar país" aria-label="Seleccionar país">
            {countrySelected ? <FlagIcon country={country} size={16} /> : <AlertTriangle size={16} />}
            <span>País</span>
          </button>
          {countryMenuOpen && (
            <div className="country-picker-menu country-picker-menu-rail">
              {Object.entries(COUNTRIES).map(([k, c]) => (
                <button type="button" key={k} className={"country-picker-option" + (k === est.country ? " active" : "")}
                  onClick={() => { changeCountry(k); setCountryMenuOpen(false); }}>
                  <FlagIcon country={c} /> {c.name}
                </button>
              ))}
            </div>
          )}
        </div>
        <button className="side-rail-btn" onClick={() => setShowChat(true)} title="Abrir Borrador inteligente (chat con la IA)" aria-label="Abrir Borrador inteligente">
          <Sparkles size={16} />
          <span>Prompt</span>
        </button>
        <button className="side-rail-btn" onClick={() => setShowHistory(true)} title="Ver historial de cotizaciones" aria-label="Ver historial de cotizaciones">
          <History size={16} />
          <span>Historial ({history.length})</span>
        </button>
        <button className="side-rail-btn" onClick={generateProposal} disabled={proposalLoading} title="Generar propuesta" aria-label="Generar propuesta">
          {proposalLoading ? <Loader2 size={16} className="spin" /> : <FileText size={16} />}
          <span>Propuesta</span>
        </button>
        {est.scenarios.length === 1 && (
          <button className="side-rail-btn" onClick={addScenario} title="Comparar otra opción (ej. remoto)" aria-label="Comparar otra opción">
            <Plus size={16} />
            <span>Comparar</span>
          </button>
        )}
      </div>

      {/* Menú lateral: paneles secundarios que se abren al medio en vez de ocupar una columna fija.
          Solo tiene sentido una vez que hay un precio calculado — antes de eso (estimación recién
          creada, sin equipo/entregables) no hay nada que revisar en Resumen/Estrategia/Intel/etc. */}
      <div className="side-rail">
        {calc.PVfinal > 0 && (
          <>
            <button className="side-rail-btn" onClick={() => setShowResumen(true)} title="Ver resumen del requerimiento" aria-label="Ver resumen del requerimiento">
              <FileText size={16} />
              <span>Resumen</span>
            </button>
            <button className="side-rail-btn" onClick={() => setShowEstrategia(true)} title="Ver estrategia de cierre recomendada" aria-label="Ver estrategia de cierre recomendada">
              <Target size={16} />
              <span>Estrategia</span>
            </button>
            <button className="side-rail-btn" onClick={() => setShowParametros(true)} title="Ver parámetros comerciales" aria-label="Ver parámetros comerciales">
              <SlidersHorizontal size={16} />
              <span>Parámetros</span>
            </button>
            <button className="side-rail-btn" onClick={() => setShowIntel(true)} title="Ver inteligencia comercial" aria-label="Ver inteligencia comercial">
              <Sparkles size={16} />
              <span>Intel.</span>
            </button>
            <button className="side-rail-btn" onClick={() => setShowSensibilidad(true)} title="Ver sensibilidad de descuento (uso interno)" aria-label="Ver sensibilidad de descuento">
              <AlertTriangle size={16} />
              <span>Uso interno</span>
            </button>
          </>
        )}
      </div>

      {showChat && (
        <div className="modal-overlay" onClick={() => setShowChat(false)}>
          <div className="card chat-modal" onClick={(e) => e.stopPropagation()}>
            <h3 style={{ justifyContent: "space-between" }}>
              <span style={{ display: "flex", alignItems: "center", gap: 8 }}><Sparkles size={14} /> Borrador inteligente</span>
              <button className="iconbtn" aria-label="Cerrar" title="Cerrar" onClick={() => setShowChat(false)}><X size={16} /></button>
            </h3>
            <div className="chatlog" style={chatMessages.length === 0 && !aiLoading ? { flex: "0 0 auto", margin: "auto 0" } : undefined}>
              {chatMessages.map((m) => <div key={m.id} className={"msg " + m.role}>{m.text}</div>)}
              {aiLoading && <div className="msg assistant"><Loader2 size={13} className="spin" style={{ display: "inline", verticalAlign: "-2px" }} /> Analizando…</div>}
            </div>
            <div
              className={"composer" + (dragOver ? " over" : "")}
              onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
              onDragLeave={() => setDragOver(false)}
              onDrop={(e) => { e.preventDefault(); setDragOver(false); addFiles(e.dataTransfer.files); }}
            >
              {est.scenarios.length > 1 && (
                <>
                  <label>Documentos para el escenario</label>
                  <div className="scenario-tabs" style={{ marginBottom: 10 }}>
                    {est.scenarios.map((s) => {
                      const n = (attachmentsByScenario[s.id] || []).filter((a) => a.kind !== "error").length;
                      return (
                        <button type="button" key={s.id} className={"scenario-pick-btn" + (s.id === est.activeScenarioId ? " active" : "")}
                          onClick={() => up({ activeScenarioId: s.id })}>
                          {s.name}{n > 0 ? ` · ${n} doc.` : ""}
                        </button>
                      );
                    })}
                  </div>
                </>
              )}
              <label>Tu prompt / instrucciones</label>
              <textarea value={est.context} onChange={(e) => up({ context: e.target.value })}
                placeholder="Escribe aquí como si hablaras conmigo: 'estima este BBP', 'agrega capacitación', 'dame una opción presencial y otra remota de 1 semana'… Adjunta documentos con el clip y pulsa Enviar." />
              <div className="hint" style={{ marginTop: 6 }}>
                {est.scenarios.length > 1
                  ? `Cada envío aplica tus instrucciones sobre la propuesta actual en lugar de empezar de cero. Arrastra archivos aquí o usa el clip para adjuntarlos al escenario seleccionado arriba ("${activeScenario.name}") — si adjuntas documentos propios a 2 o más escenarios, se analizan por separado y obtienes resultados por cada uno al enviar.`
                  : `Cada envío aplica tus instrucciones sobre la propuesta actual (si ya existe) en lugar de empezar de cero. Arrastra archivos aquí o usa el clip para adjuntar.`}
              </div>
              {parsing && <div className="hint" style={{ marginTop: 6 }}><Loader2 size={12} className="spin" style={{ display: "inline", verticalAlign: "-2px" }} /> Procesando documentos…</div>}

              {attachments.length > 0 && (
                <div className="filelist">
                  {attachments.map((a) => {
                    const Icon = a.kind === "image" ? ImageIcon : a.kind === "pdf" ? FileType : FileText;
                    const meta = a.kind === "error" ? a.text
                      : a.kind === "pdf" ? "PDF · lo lee la IA"
                      : a.kind === "image" ? "imagen · la lee la IA"
                      : (a.text?.length || 0).toLocaleString() + " caracteres";
                    return (
                      <div key={a.id} className={"filechip" + (a.kind === "error" ? " err" : "")}>
                        <Icon size={15} />
                        <div style={{ minWidth: 0, flex: 1 }}>
                          <div style={{ fontWeight: 600, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{a.name}</div>
                          <div className="hint">{meta}</div>
                        </div>
                        <button className="iconbtn" aria-label={`Quitar ${a.name}`} title="Quitar archivo" onClick={() => removeAttachment(a.id)}><X size={14} /></button>
                      </div>
                    );
                  })}
                </div>
              )}

              <div style={{ display: "flex", alignItems: "center", gap: 8, marginTop: 10 }}>
                <button className="btn ghost attach-btn" aria-label="Adjuntar documentos" title="Adjuntar BBP, transcripciones, actas… Word/PDF/imágenes/texto" onClick={() => fileRef.current?.click()}>
                  <Paperclip size={16} />
                </button>
                <input ref={fileRef} type="file" multiple style={{ display: "none" }}
                  accept=".docx,.pdf,.txt,.md,.markdown,.vtt,.srt,.csv,.tsv,.json,.log,.html,.png,.jpg,.jpeg,.gif,.webp"
                  onChange={(e) => { addFiles(e.target.files); e.target.value = ""; }} />
                <button className="btn primary" onClick={generateDraft} disabled={aiLoading || parsing} style={{ flex: 1, justifyContent: "center" }}>
                  {aiLoading ? <Loader2 size={15} className="spin" /> : <Sparkles size={15} />}
                  {aiLoading ? "Analizando…" : "Enviar"}
                </button>
              </div>
              {aiError && <div className="note" style={{ marginTop: 10, background: "#FBE6E6", color: "var(--bad)" }}><AlertTriangle size={15} /> {aiError}</div>}
            </div>
          </div>
        </div>
      )}

      {showResumen && (
        <ModalCard onClose={() => setShowResumen(false)} maxWidth={680} cardStyle={{ borderLeft: "4px solid var(--accent)" }}
          title={<span style={{ display: "flex", alignItems: "center", gap: 8 }}><FileText size={14} /> Resumen del requerimiento — qué necesita el cliente</span>}>
          <textarea
            value={est.insight.resumenRequerimiento}
            onChange={(e) => up({ insight: { ...est.insight, resumenRequerimiento: e.target.value } })}
            placeholder="Qué problema/necesidad trae el cliente y qué se le va a entregar, en lenguaje simple — se completa al generar el borrador."
            style={{ fontSize: 15, lineHeight: 1.7, minHeight: 280 }}
          />
          <div className="hint">Léelo primero: es el resumen llano de la necesidad real, sin jerga — distinto del perfil del cliente (a qué se dedica).</div>
        </ModalCard>
      )}

      {showEstrategia && (
        <ModalCard onClose={() => setShowEstrategia(false)} maxWidth={680} cardStyle={{ borderColor: "var(--gold)", background: "#FFFDF6" }}
          title="★ Estrategia de cierre recomendada">
          <textarea value={est.insight.estrategiaCierre} onChange={(e) => up({ insight: { ...est.insight, estrategiaCierre: e.target.value } })}
            placeholder="Postura de precio para ganar siendo rentable, qué enfatizar ante este cliente y riesgos a cuidar…"
            style={{ borderColor: "var(--gold)", background: "#fff", fontSize: 14.5, lineHeight: 1.7, minHeight: 220 }} />
          <div className="hint">La IA la completa al generar el borrador (uso interno, aplica a todos los escenarios). Puedes editarla libremente.</div>
        </ModalCard>
      )}

      {showParametros && (
        <ModalCard onClose={() => setShowParametros(false)} maxWidth={920} title="Parámetros comerciales (editables)">
          <div className="modal-section-label">Estructura de costos — compartida entre todos los escenarios</div>
          <div className="hint">Estos gastos SÍ se suman al precio final del cliente (full cost recovery) — el precio antes de descuento es horas×tarifa + Adm/Comercial/MKT, así tu margen queda protegido en vez de absorber estos gastos de tu utilidad.</div>
          <div className="row4">
            <div className="param-field" title="Margen objetivo / piso sobre el precio de venta">
              <label>Margen objetivo</label>
              <InlineLabel percent value={est.margin} onChange={(v) => up({ margin: v })} decimals={1} />
            </div>
            <div className="param-field">
              <label>Gastos administrativos</label>
              <InlineLabel percent value={est.adm} onChange={(v) => up({ adm: v })} decimals={1} />
            </div>
            <div className="param-field" title="Gestión comercial — Carmen">
              <label>Gestión comercial</label>
              <InlineLabel percent value={est.com} onChange={(v) => up({ com: v })} decimals={1} />
            </div>
            <div className="param-field">
              <label>Gestión MKT</label>
              <InlineLabel percent value={est.mkt} onChange={(v) => up({ mkt: v })} decimals={1} />
            </div>
          </div>

          <div className="modal-section-label">Descuento — solo {activeScenario.name}</div>
          <div className="param-field" style={{ display: "inline-block", minWidth: 200 }}>
            <label>Descuento comercial</label>
            <InlineLabel percent value={activeScenario.discount} onChange={(v) => upScenario({ discount: v })} decimals={0} />
            {activeScenario.discount > country.maxDesc && <div className="hint" style={{ color: "var(--warn)" }}>Por encima del descuento competitivo sugerido para {country.name} ({pct(country.maxDesc)}).</div>}
          </div>

          <div className="note" style={{ background: calc.grossMargin < est.margin ? "#FBF1DD" : "var(--accent-soft)", color: calc.grossMargin < est.margin ? "var(--warn)" : "var(--accent)" }}>
            {calc.grossMargin < est.margin ? <AlertTriangle size={15} style={{ flexShrink: 0 }} /> : <CheckCircle2 size={15} style={{ flexShrink: 0 }} />}
            <span>Margen bruto del precio competitivo: <b>{pct(calc.grossMargin)}</b> (precio venta {fmtUSD(calc.PVO)} vs. costo país {fmtUSD(calc.CO)}). Piso por margen objetivo: <b>{fmtUSD(calc.floorPrice)}</b>. {calc.grossMargin < est.margin ? "El precio de mercado queda por debajo de tu margen objetivo — decide si compites igual o subes precio." : "El precio competitivo cumple tu margen objetivo."}</span>
          </div>
        </ModalCard>
      )}

      {showIntel && (
        <ModalCard onClose={() => setShowIntel(false)} maxWidth={680}
          title={<span style={{ display: "flex", alignItems: "center", gap: 8 }}><Sparkles size={14} /> Inteligencia comercial · {country.name}</span>}>
          {est.insight.hechosClave?.length > 0 && (
            <div className="note" style={{ marginBottom: 16, alignItems: "flex-start" }}>
              <CheckCircle2 size={15} style={{ flexShrink: 0, marginTop: 1 }} />
              <span>
                <b>Verifica estas cifras contra tus documentos</b> — son las que la IA usó para dimensionar el alcance:
                <ul style={{ margin: "6px 0 0", paddingLeft: 18 }}>
                  {est.insight.hechosClave.map((h, i) => <li key={i}>{h}</li>)}
                </ul>
              </span>
            </div>
          )}
          <div className="tabs">
            {[
              ["competencia", "Análisis de competencia"],
              ["valor", "Valor agregado"],
              ["ia", "Oportunidad de IA"],
            ].map(([id, label]) => (
              <button key={id} type="button" className={"tab" + (intelTab === id ? " active" : "")} onClick={() => setIntelTab(id)}>{label}</button>
            ))}
          </div>
          <div className="tabpanel">
            {intelTab === "competencia" && (
              <div>
                <label>Quién compite y cómo diferenciarse</label>
                <textarea value={est.insight.competencia} onChange={(e) => up({ insight: { ...est.insight, competencia: e.target.value } })}
                  placeholder="Quién compite en este país, rango de precios típico y cómo diferenciarse…" />
              </div>
            )}
            {intelTab === "valor" && (
              <div>
                <label>Valor agregado a proponer (una idea por línea)</label>
                <textarea value={(est.insight.valor || []).join("\n")} onChange={(e) => up({ insight: { ...est.insight, valor: e.target.value.split("\n").filter((x) => x.trim()) } })}
                  placeholder="Entregables o servicios extra que justifican el precio…" />
              </div>
            )}
            {intelTab === "ia" && (
              <div>
                <label>Oportunidad de IA / automatización</label>
                <textarea value={est.insight.ia} onChange={(e) => up({ insight: { ...est.insight, ia: e.target.value } })}
                  placeholder="Copilotos, chatbots, automatización, BI, agentes… o 'No aplica'." />
              </div>
            )}
          </div>
          <div className="hint">La IA completa esto al generar el borrador; siempre investiga al cliente, analiza la competencia (Perú/Ecuador), sugiere valor agregado y propone IA.</div>
        </ModalCard>
      )}

      {showSensibilidad && (
        <ModalCard onClose={() => setShowSensibilidad(false)} maxWidth={1040} title={`Sensibilidad de descuento · alineado a ${country.name}`}>
          <div style={{ overflowX: "auto" }}>
            <table style={{ minWidth: 760 }}>
              <thead><tr><th className="l">Escenario</th><th>Desc.</th><th>Precio USD</th>{hasLocalCurrency && <th>Precio {country.cur}</th>}<th>Utilidad</th><th>Rentabilidad</th><th className="l">Recomendación</th></tr></thead>
              <tbody>
                {discountScen.map((s) => (
                  <tr key={s.d}>
                    <td className="l" style={{ whiteSpace: "nowrap" }}>{s.d === 0 ? "Sin descuento" : `Descuento ${s.d * 100}%`}</td>
                    <td className="num">{(s.d * 100).toFixed(0)}%</td>
                    <td className="num">{fmtUSD(s.price)}</td>
                    {hasLocalCurrency && <td className="num">{fmtLocal(s.price * est.fx, country.cur)}</td>}
                    <td className="num">{fmtUSD(s.util)}</td>
                    <td className="num" style={{ color: s.light === "ok" ? "var(--ok)" : s.light === "warn" ? "var(--warn)" : "var(--bad)", fontWeight: 600 }}>{pct(s.rent)}</td>
                    <td className="l"><span className={"pill " + s.light}>{s.light === "ok" ? <CheckCircle2 size={12} /> : <AlertTriangle size={12} />}{s.txt}</span></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div className="note"><AlertTriangle size={15} style={{ flexShrink: 0 }} /><span>Hoja de uso interno. El cliente recibe únicamente el precio llave en mano, sin desglose de descuentos ni márgenes. Mínimo aceptable para {country.name}: <b>{pct(country.minRent)}</b> de rentabilidad neta.</span></div>
        </ModalCard>
      )}

      {/* Drawer historial */}
      {showHistory && (
        <div className="drawer" onClick={() => setShowHistory(false)}>
          <div className="panel" onClick={(e) => e.stopPropagation()}>
            <div style={{ display: "flex", alignItems: "center", marginBottom: 16 }}>
              <h3 className="disp" style={{ fontSize: 16, margin: 0 }}>Historial de cotizaciones</h3>
              <button className="iconbtn" aria-label="Cerrar historial" title="Cerrar" style={{ marginLeft: "auto", color: "var(--ink)" }} onClick={() => setShowHistory(false)}><X size={18} /></button>
            </div>
            {history.length === 0 && <div className="note">Aún no hay cotizaciones guardadas. Usa <b>Guardar</b> para registrar la actual.</div>}
            {history.length > 0 && (() => {
              const pendientes = history.filter((h) => !h.sentAt);
              const enviados = history.filter((h) => h.sentAt);
              const base = historyTab === "enviados" ? enviados : pendientes;
              const filtered = base.filter((h) =>
                (!historyCountryFilter || h.country === historyCountryFilter) &&
                (!historySearch.trim() || (h.client || "").toLowerCase().includes(historySearch.trim().toLowerCase()))
              );
              return (
                <>
                  <div className="tabs" style={{ marginBottom: 14 }}>
                    <button type="button" className={"tab" + (historyTab === "pendientes" ? " active" : "")} onClick={() => setHistoryTab("pendientes")}>Pendientes ({pendientes.length})</button>
                    <button type="button" className={"tab" + (historyTab === "enviados" ? " active" : "")} onClick={() => setHistoryTab("enviados")}>Enviados ({enviados.length})</button>
                  </div>
                  <div style={{ display: "flex", gap: 8, marginBottom: 14 }}>
                    <input placeholder="Buscar por cliente…" value={historySearch} onChange={(e) => setHistorySearch(e.target.value)} style={{ flex: 1 }} />
                    <select value={historyCountryFilter} onChange={(e) => setHistoryCountryFilter(e.target.value)} style={{ width: 140 }}>
                      <option value="">Todos los países</option>
                      {Object.entries(COUNTRIES).map(([k, c]) => <option key={k} value={k}>{c.flag} {c.name}</option>)}
                    </select>
                  </div>
                  {historyTab === "pendientes" && selectedCodes.size > 0 && (
                    <div className="note" style={{ marginBottom: 14, background: "var(--accent-soft)", color: "var(--accent)", alignItems: "center" }}>
                      <span style={{ flex: 1 }}><b>{selectedCodes.size}</b> seleccionada{selectedCodes.size > 1 ? "s" : ""}</span>
                      <button className="btn primary" style={{ fontSize: 12, padding: "6px 10px" }} onClick={sendSelectedByEmail} disabled={sendingEmail}>
                        {sendingEmail ? <Loader2 size={13} className="spin" /> : <Mail size={13} />} Enviar por correo a {SEND_TO}
                      </button>
                      <button className="btn ghost" style={{ fontSize: 12, padding: "6px 10px" }} onClick={() => setSelectedCodes(new Set())} disabled={sendingEmail}>Cancelar</button>
                    </div>
                  )}
                  {filtered.length === 0 && <div className="note">{historyTab === "enviados" ? "Aún no se envió ninguna cotización por correo." : "Sin resultados para ese filtro."}</div>}
                  {filtered.map((h) => {
                    const expired = h.expiresAt && h.expiresAt < Date.now();
                    const expSoon = !expired && h.expiresAt && h.expiresAt - Date.now() < 5 * 86400000;
                    return (
                      <div key={h.code} className="histcard" onClick={() => loadEstimation(h)}>
                        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                          {historyTab === "pendientes" && (
                            <input type="checkbox" aria-label={`Seleccionar ${h.code}`} checked={selectedCodes.has(h.code)}
                              onClick={(e) => e.stopPropagation()} onChange={() => toggleSelected(h.code)} style={{ width: "auto" }} />
                          )}
                          <span className="mono" style={{ fontWeight: 700, fontSize: 13 }}>{h.code}</span>
                          <span style={{ marginLeft: "auto", fontSize: 11, color: "var(--muted)" }}>{h.savedAt ? new Date(h.savedAt).toLocaleDateString() : ""}</span>
                          <button className="iconbtn" aria-label={`Eliminar cotización ${h.code}`} title="Eliminar" onClick={(e) => { e.stopPropagation(); deleteEstimation(h.code); }}><Trash2 size={14} /></button>
                        </div>
                        <div style={{ fontWeight: 600, marginTop: 3 }}>{h.project || "(sin título)"}</div>
                        <div style={{ fontSize: 12, color: "var(--muted)" }}>{COUNTRIES[h.country]?.flag} {h.client} · {(() => {
                          const totalsArr = Array.isArray(h.totals) ? h.totals : (h.totals ? [{ name: "Escenario 1", ...h.totals }] : []);
                          if (!totalsArr.length) return "—";
                          if (totalsArr.length === 1) return `${fmtUSD(totalsArr[0].PVfinal)} · rent. ${pct(totalsArr[0].rent)}`;
                          const prices = totalsArr.map((t) => t.PVfinal);
                          return `${fmtUSD(Math.min(...prices))}–${fmtUSD(Math.max(...prices))} · ${totalsArr.length} escenarios`;
                        })()}</div>
                        {h.sentAt ? (
                          <div style={{ fontSize: 11, marginTop: 4, color: "var(--ok)" }}>
                            <Mail size={11} style={{ display: "inline", verticalAlign: "-1px" }} /> Enviado el {new Date(h.sentAt).toLocaleDateString()} {new Date(h.sentAt).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" })} a {h.sentTo}
                          </div>
                        ) : h.expiresAt && (
                          <div style={{ fontSize: 11, marginTop: 4, color: expired ? "var(--bad)" : expSoon ? "var(--warn)" : "var(--muted)" }}>
                            {expired ? "Vencida el " : "Vence el "}{new Date(h.expiresAt).toLocaleDateString()}
                          </div>
                        )}
                      </div>
                    );
                  })}
                </>
              );
            })()}
            {history.length > 0 && (
              <div className="note" style={{ marginTop: 8 }}>
                Clientes registrados: {[...new Set(history.map((h) => h.client).filter(Boolean))].map((cl) => `${cl} (${clientCount(cl)})`).join(" · ")}
              </div>
            )}
          </div>
        </div>
      )}

      {showProposal && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(20,32,46,.55)", zIndex: 45, display: "flex", flexDirection: "column", padding: "24px" }} onClick={() => setShowProposal(false)}>
          <div onClick={(e) => e.stopPropagation()} style={{ background: "#fff", borderRadius: 14, overflow: "hidden", maxWidth: 900, width: "100%", margin: "0 auto", height: "100%", display: "flex", flexDirection: "column", boxShadow: "0 20px 60px rgba(0,0,0,.3)" }}>
            <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "12px 16px", borderBottom: "1px solid var(--line)" }}>
              <FileText size={17} style={{ color: "var(--accent)" }} />
              <b className="disp">Propuesta comercial · {est.client || "Cliente"}</b>
              <div style={{ marginLeft: "auto", display: "flex", gap: 8 }}>
                <button className="btn" onClick={() => proposalFrame.current?.contentWindow?.print()}><FileText size={15} /> Imprimir / PDF</button>
                <button className="btn primary" onClick={downloadProposalDocx}><Download size={15} /> Descargar Word</button>
                <button className="iconbtn" aria-label="Cerrar propuesta" title="Cerrar" style={{ color: "var(--ink)" }} onClick={() => setShowProposal(false)}><X size={18} /></button>
              </div>
            </div>
            <iframe ref={proposalFrame} title="Propuesta" srcDoc={proposalHTML} style={{ flex: 1, border: "none", width: "100%" }} />
          </div>
        </div>
      )}

      {toast && <div className="toast">{toast}</div>}
      <style>{`.spin{animation:sp 1s linear infinite}@keyframes sp{to{transform:rotate(360deg)}}`}</style>
    </div>
  );
}
