import React, { useState, useMemo, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import mammoth from "mammoth";
import {
  Globe, Sparkles, Plus, Trash2, Save, Download, History, FileSpreadsheet,
  AlertTriangle, CheckCircle2, RotateCcw, Building2, Loader2, X, Paperclip,
  FileText, FileType, Image as ImageIcon, UploadCloud, LogOut
} from "lucide-react";
import { supabase } from "./lib/supabaseClient";
import { store } from "./lib/store";

/* =========================================================================
   EBIM — Estimador de Costos de Consultoría TI
   Moneda de cotización: USD. Tarifas referenciales por país (editables).
   ========================================================================= */

const EBIM_PROFILE = `GRUPO EBIM SAC (RUC 20602517986, San Isidro - Lima, Perú; presencia en Perú y Ecuador; ~11-50 personas; +7 años). Consultora de TI. Servicios: Consultoría e implementación SAP S/4HANA (módulos SD MM FI CO PP PM QM PS HR TRM TMS VMS IM WM, BTP, FIORI, BASIS, ABAP), desarrollo de software web/móvil a medida, arquitectura cloud (AWS/Azure/GCP) y migraciones, BI/QLIK, IA/chatbots, IoT, outsourcing de personal TI, mesa de ayuda 24/7, soporte y AMS. Trabaja con metodologías ágiles. Perfiles típicos: Arquitecto TI/Consultor Senior, Consultor Semi-Senior, Analista de Procesos/Capacidades TI y PM/Coordinador.`;

// EBIM opera en Perú y Ecuador. Tarifas por país en USD/hora.
// cost = costo real del consultor (cargado); sale = precio de venta competitivo de EBIM.
// minRent = rentabilidad neta mínima aceptable; maxDesc = descuento competitivo máx.
const COUNTRIES = {
  PE: { name: "Perú", flag: "🇵🇪", cur: "PEN", fx: 3.75, minRent: 0.20, maxDesc: 0.10, note: "Mercado base de EBIM. Competencia local fuerte en SAP y desarrollo; diferénciate por experiencia y entregables accionables.",
        rates: { senior: { cost: 26, sale: 62.5 }, semi: { cost: 17, sale: 41.67 }, analista: { cost: 11, sale: 25 } } },
  EC: { name: "Ecuador", flag: "🇪🇨", cur: "USD", fx: 1, minRent: 0.20, maxDesc: 0.10, note: "Economía dolarizada: precios en USD directos, sin riesgo cambiario. Menos competencia SAP local que Perú; oportunidad de posicionar valor regional.",
        rates: { senior: { cost: 24, sale: 58 }, semi: { cost: 16, sale: 38 }, analista: { cost: 10, sale: 23 } } },
};

const TIERS = [["senior", "Senior / Arquitecto / PM"], ["semi", "Semi-Senior"], ["analista", "Analista"]];
const cloneRates = (r) => ({ senior: { ...r.senior }, semi: { ...r.semi }, analista: { ...r.analista } });

// Paleta de marca EBIM — única fuente de verdad para la app y el HTML de la propuesta (que vive en su propio documento/iframe).
const BRAND = {
  bg: "#EBEEF2", surface: "#FFFFFF", ink: "#15202E", muted: "#5E6E81", line: "#D9E0E8", lineSoft: "#E7ECF1",
  accent: "#0B5563", accentSoft: "#E2EFF0", gold: "#9A6B12", ok: "#0F8A5F", warn: "#B7791F", bad: "#C13B3B",
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
const fmtUSD2 = (n) =>
  "$" + (Number.isFinite(n) ? n : 0).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const fmtLocal = (n, cur) =>
  (Number.isFinite(n) ? n : 0).toLocaleString("es-PE", { minimumFractionDigits: 0, maximumFractionDigits: 0 }) + " " + cur;
const pct = (n) => (Number.isFinite(n) ? (n * 100).toFixed(1) : "0.0") + "%";
const esc = (s) => String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

const newReqCode = (existing = []) => {
  let c;
  do { c = "REQ-" + String(Math.floor(100000 + Math.random() * 900000)); } while (existing.includes(c));
  return c;
};

const blankDeliverable = (name = "") => ({ id: crypto.randomUUID(), name, hS: 0, hM: 0, hA: 0, start: 0, dur: 1 });

const defaultEstimation = (countryKey = "PE", code) => {
  const c = COUNTRIES[countryKey];
  return {
    code: code || newReqCode(),
    client: "",
    project: "",
    country: countryKey,
    fx: c.fx,
    rates: cloneRates(c.rates),
    context: "",
    insight: { perfilCliente: "", competencia: "", valor: [], ia: "", estrategiaCierre: "" },
    team: [
      { id: crypto.randomUUID(), perfil: "Arquitecto TI / Consultor Senior", tier: "senior", rol: "Lidera diagnóstico, arquitectura y roadmap" },
      { id: crypto.randomUUID(), perfil: "Consultor Semi-Senior TI", tier: "semi", rol: "Levantamiento técnico: apps, plataformas, APIs" },
      { id: crypto.randomUUID(), perfil: "Analista de Procesos / Cap. TI", tier: "analista", rol: "Relevamiento de equipo y brechas" },
      { id: crypto.randomUUID(), perfil: "PM / Coordinador de Proyecto", tier: "senior", rol: "Gestión, cronograma y presentación ejecutiva" },
    ],
    deliverables: [
      blankDeliverable("1. Kick-off, planificación y accesos"),
    ],
    margin: 0.35,
    adm: 0.045,
    com: 0.05,
    mkt: 0.01,
    discount: 0,
    schedule: [
      { id: crypto.randomUUID(), hito: "Inicio del Requerimiento", pct: 0.35 },
      { id: crypto.randomUUID(), hito: "Inicio de Pruebas Integrales", pct: 0.35 },
      { id: crypto.randomUUID(), hito: "Entrega / Pase a Producción", pct: 0.30 },
    ],
    savedAt: null,
  };
};

/* ---------- cálculo central (costo país vs precio de venta competitivo) ---------- */
function compute(est) {
  const r = est.rates;
  const rows = est.deliverables.map((d) => {
    const hS = d.hS || 0, hM = d.hM || 0, hA = d.hA || 0;
    const cost = hS * r.senior.cost + hM * r.semi.cost + hA * r.analista.cost;
    const sale = hS * r.senior.sale + hM * r.semi.sale + hA * r.analista.sale;
    return { ...d, hrs: hS + hM + hA, cost, sale };
  });
  const CO = rows.reduce((s, x) => s + x.cost, 0);          // costo operativo (costo país)
  const PVO = rows.reduce((s, x) => s + x.sale, 0);          // precio de venta competitivo (país)
  const totalHrs = rows.reduce((s, x) => s + x.hrs, 0);
  const grossMargin = PVO > 0 ? (PVO - CO) / PVO : 0;        // margen bruto resultante
  const floorPrice = est.margin < 1 ? CO / (1 - est.margin) : CO; // piso por margen objetivo
  const gAdm = est.adm * PVO, gCom = est.com * PVO, gMkt = est.mkt * PVO;
  const gastos = gAdm + gCom + gMkt;
  const PVfinal = PVO * (1 - est.discount);
  const utilidad = PVfinal - CO - gastos;
  const rent = PVfinal > 0 ? utilidad / PVfinal : 0;
  const weeksTotal = rows.reduce((m, x) => Math.max(m, (x.start || 0) + (x.dur || 1)), 1);
  return { rows, CO, PVO, totalHrs, weeksTotal, grossMargin, floorPrice, gAdm, gCom, gMkt, gastos, PVfinal, utilidad, rent };
}

function scenarioRow(calc, d, minRent) {
  const price = calc.PVO * (1 - d);
  const util = price - calc.CO - calc.gastos;
  const rent = price > 0 ? util / price : 0;
  let light = "ok", txt = "Aceptable — sobre el mínimo";
  if (rent < 0) { light = "bad"; txt = "No recomendado — rentabilidad negativa"; }
  else if (rent < minRent) { light = "warn"; txt = "Revisar — bajo el mínimo del país"; }
  else if (d === 0) { txt = "Precio óptimo — mantener si es posible"; }
  return { d, price, util, rent, light, txt };
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

function buildProposalHTML(est, calc, c, prose) {
  const today = new Date().toLocaleDateString("es-PE", { day: "2-digit", month: "long", year: "numeric" });
  const weeks = calc.weeksTotal || Math.max(1, Math.ceil(calc.totalHrs / (Math.max(1, est.team.filter((t) => t.perfil).length) * 32)));
  const gantt = est.deliverables.filter((d) => d.name).map((d) => {
    const s = d.start || 0, du = d.dur || 1;
    return `<tr><td style="font-size:13px">${esc(d.name)}</td><td style="width:58%"><div style="position:relative;height:18px;background:#EEF3F5;border-radius:5px"><div style="position:absolute;left:${(s / weeks) * 100}%;width:${(du / weeks) * 100}%;top:0;bottom:0;background:#0B5563;border-radius:5px"></div></div></td><td class="r" style="white-space:nowrap;font-size:12px;color:#5E6E81">Sem ${s + 1}–${s + du}</td></tr>`;
  }).join("");
  const deliv = est.deliverables.filter((d) => d.name).map((d) => `<li>${esc(d.name)}</li>`).join("");
  const team = est.team.filter((t) => t.perfil).map((t) => `<tr><td><b>${esc(t.perfil)}</b></td><td>${esc(t.rol)}</td></tr>`).join("");
  const pay = est.schedule.map((s, i) => `<tr><td>${i + 1}</td><td>${esc(s.hito)}</td><td class="r">${(s.pct * 100).toFixed(0)}%</td><td class="r">${fmtUSD(calc.PVfinal * s.pct)}</td></tr>`).join("");
  const valor = (prose.diferenciadores || []).filter(Boolean).map((v) => `<li>${esc(v)}</li>`).join("");
  const iaBlock = est.insight?.ia && !/no aplica/i.test(est.insight.ia)
    ? `<section><h2>Innovación con Inteligencia Artificial</h2><p>${esc(est.insight.ia)}</p></section>` : "";
  return `<!doctype html><html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Propuesta ${esc(est.client)} — ${esc(est.code)}</title>
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@600;700&family=Inter:wght@400;500;600&display=swap');
*{box-sizing:border-box} body{margin:0;font-family:'Inter',system-ui,sans-serif;color:${BRAND.ink};line-height:1.55;background:${BRAND.surface}}
.page{max-width:820px;margin:0 auto;padding:48px 56px}
.disp{font-family:'Space Grotesk',sans-serif}
.cover{min-height:88vh;display:flex;flex-direction:column;justify-content:center;border-left:6px solid ${BRAND.accent};padding-left:36px}
.brand{display:flex;align-items:center;gap:12px;margin-bottom:40px}
.mark{width:46px;height:46px;border-radius:10px;background:${BRAND.accent};color:#fff;display:grid;place-items:center;font-weight:700;font-family:'Space Grotesk';font-size:22px}
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
@media print{.page{padding:24px 30px}.cover{min-height:94vh}}
</style></head><body>
<div class="page">
  <div class="cover">
    <div class="brand"><div class="mark">E</div><div><div class="disp" style="font-weight:700;font-size:18px">GRUPO EBIM</div><div style="color:${BRAND.muted};font-size:13px">Consultoría en Tecnologías de la Información</div></div></div>
    <div class="eyebrow">Propuesta de servicios profesionales</div>
    <h1>${esc(est.project || "Servicio de consultoría TI")}</h1>
    <div class="meta">Preparada para <b>${esc(est.client || "—")}</b> · ${esc(c.name)}<br>${today}</div>
    <div class="req">${esc(est.code)}</div>
  </div>

  <section><h2>Resumen ejecutivo</h2><p>${esc(prose.entendimiento)}</p></section>
  ${est.insight?.perfilCliente ? `<section><h2>Entendimiento de su negocio</h2><p>${esc(est.insight.perfilCliente)}</p></section>` : ""}
  <section><h2>Nuestro enfoque</h2><p>${esc(prose.enfoque)}</p></section>
  <section><h2>Alcance y entregables</h2><ul>${deliv || "<li>Por definir</li>"}</ul></section>
  <section><h2>Equipo asignado</h2><table><thead><tr><th>Perfil</th><th>Responsabilidad</th></tr></thead><tbody>${team || "<tr><td>Por definir</td><td></td></tr>"}</tbody></table></section>
  <section><h2>Cronograma de trabajo</h2><p>Duración estimada de <b>${weeks} semana${weeks > 1 ? "s" : ""}</b>, organizada por entregable con hitos de avance y validación con su equipo:</p>
    <table><thead><tr><th>Entregable</th><th>Avance en el tiempo</th><th class="r">Semanas</th></tr></thead><tbody>${gantt || "<tr><td>Por definir</td><td></td><td></td></tr>"}</tbody></table></section>
  <section><h2>Valor agregado y diferenciadores</h2><ul>${valor}</ul></section>
  ${iaBlock}
  <section><h2>Por qué GRUPO EBIM</h2><p>${esc(prose.porQueEbim)}</p>
    <div><span class="tag">SAP S/4HANA</span><span class="tag">Desarrollo a medida</span><span class="tag">Cloud AWS · Azure · GCP</span><span class="tag">Inteligencia Artificial</span><span class="tag">+7 años</span></div></section>
  <section><h2>Inversión</h2>
    <div class="invest"><div><div class="label">Inversión total — llave en mano</div><small>No incluye IGV</small></div><div style="text-align:right"><div class="amt">USD ${fmtUSD(calc.PVfinal).replace("$", "")}</div><small>≈ ${fmtLocal(calc.PVfinal * est.fx, c.cur)}</small></div></div>
    <table style="margin-top:16px"><thead><tr><th>N°</th><th>Hito de pago</th><th class="r">%</th><th class="r">Monto USD</th></tr></thead><tbody>${pay}</tbody></table>
  </section>
  <section><h2>Siguientes pasos</h2><p>${esc(prose.cierre)}</p></section>
  <div class="foot"><span>GRUPO EBIM SAC · Consultoría TI · Lima, Perú · www.grupoebim.com</span><span>${esc(est.code)} · Oferta válida por 30 días</span></div>
</div></body></html>`;
}

/* ============================ APP ============================ */
export default function App() {
  const [est, setEst] = useState(() => defaultEstimation());
  const [history, setHistory] = useState([]);
  const [showHistory, setShowHistory] = useState(false);
  const [aiLoading, setAiLoading] = useState(false);
  const [aiError, setAiError] = useState("");
  const [attachments, setAttachments] = useState([]);
  const [parsing, setParsing] = useState(false);
  const [detected, setDetected] = useState(null);
  const [proposalLoading, setProposalLoading] = useState(false);
  const [proposalHTML, setProposalHTML] = useState("");
  const [showProposal, setShowProposal] = useState(false);
  const proposalFrame = useRef(null);
  const fileRef = useRef(null);
  const [dragOver, setDragOver] = useState(false);
  const [toast, setToast] = useState("");
  const calc = useMemo(() => compute(est), [est]);
  const country = COUNTRIES[est.country];

  useEffect(() => { refreshHistory(); }, []);
  const flash = (m) => { setToast(m); setTimeout(() => setToast(""), 2600); };

  async function refreshHistory() {
    const keys = await store.list();
    const items = [];
    for (const k of keys) {
      const v = await store.get(k);
      if (v) { try { items.push(JSON.parse(v)); } catch (e) {} }
    }
    items.sort((a, b) => (b.savedAt || 0) - (a.savedAt || 0));
    setHistory(items);
  }

  const up = (patch) => setEst((e) => ({ ...e, ...patch }));

  function changeCountry(key) {
    const c = COUNTRIES[key];
    setEst((e) => ({ ...e, country: key, fx: c.fx, rates: cloneRates(c.rates) }));
  }

  function newEstimation() {
    setEst(defaultEstimation(est.country, newReqCode(history.map((h) => h.code))));
    setAttachments([]);
    setDetected(null);
    setAiError("");
    flash("Nueva estimación iniciada");
  }

  async function saveEstimation() {
    if (!est.client.trim()) { flash("Falta el nombre del cliente"); return; }
    const payload = { ...est, savedAt: Date.now(), attachmentNames: attachments.filter((a) => a.kind !== "error").map((a) => a.name), totals: { CO: calc.CO, PVfinal: calc.PVfinal, rent: calc.rent } };
    await store.set("est:" + est.code, JSON.stringify(payload));
    await refreshHistory();
    flash("Guardado " + est.code);
  }

  function loadEstimation(item) {
    setEst({ ...item }); setAttachments([]); setDetected(null); setShowHistory(false);
    flash("Cargada " + item.code);
  }

  async function deleteEstimation(code) {
    await store.del("est:" + code); await refreshHistory();
  }

  const clientCount = (name) =>
    history.filter((h) => (h.client || "").trim().toLowerCase() === (name || "").trim().toLowerCase()).length;

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
    setAttachments((a) => [...a, ...out]);
    setParsing(false);
    flash(out.length + " documento(s) adjuntado(s)");
  }

  const removeAttachment = (id) => setAttachments((a) => a.filter((x) => x.id !== id));

  /* ---------- IA: generar borrador ---------- */
  async function generateDraft() {
    const usable = attachments.filter((a) => a.kind !== "error");
    if (!est.context.trim() && usable.length === 0) { setAiError("Pega el contexto o adjunta al menos un documento (BBP, acta, transcripción…)."); return; }
    setAiLoading(true); setAiError("");
    const c = COUNTRIES[est.country];
    const hist = history.slice(0, 10).map((h) => `- ${h.project || "(sin título)"} [${h.client}, ${COUNTRIES[h.country]?.name || h.country}] ${h.insight?.ia ? "· IA: " + h.insight.ia.slice(0, 80) : ""}`).join("\n") || "- (sin historial aún)";
    const sys = `Eres la parte operativo-comercial senior de tecnología de GRUPO EBIM: actúas como ESTRATEGA COMERCIAL que arma propuestas ganadoras, rentables y a la medida del mercado. EBIM opera en PERÚ y ECUADOR (todo en USD); enfócate en esos dos mercados. Conoces el negocio:
${EBIM_PROFILE}
Aprendes del historial de propuestas de EBIM (reutiliza patrones, perfiles, precios y valor agregado que funcionaron):
${hist}

En CADA estimación SIEMPRE debes:
1) Investigar al cliente (con web_search) y analizar la competencia local del país (Perú o Ecuador): quién compite, rango de precios típico y cómo diferenciarse para no competir solo por precio.
2) Recomendar valor agregado concreto que EBIM puede sumar (entregables o servicios extra que justifican el precio y abren cross-sell).
3) Identificar dónde vender IA/automatización si aplica (copilotos, chatbots, automatización de procesos, analítica/BI, agentes); si no aplica, dilo.
4) Dar una ESTRATEGIA DE CIERRE: postura de precio recomendada para ganar el trato siendo rentable, qué enfatizar ante este cliente, y riesgos a cuidar.

El usuario de EBIM puede incluir INSTRUCCIONES o consideraciones (p. ej. "agrega un consultor de seguridad", "hazlo más competitivo", "incluye fase 2"): aplícalas. Si recibes una PROPUESTA ACTUAL en JSON, modifica ESA base y conserva lo no afectado, en lugar de empezar de cero.

Tarea: a partir del CONTEXTO y de los DOCUMENTOS ADJUNTOS (BBP/Business Blueprint, transcripciones de Teams, actas, propuestas, diagramas): (a) DETECTA automáticamente el tipo de servicio/proyecto, el CLIENTE y el PAÍS del cliente; (b) propón el equipo y el desglose de entregables con horas por perfil para ese país. Si hay un BBP, deriva los entregables de los procesos y gaps; si hay transcripciones, extrae alcance, supuestos y compromisos. Usa web_search para validar cliente, competencia o tarifas locales. Devuelve EXCLUSIVAMENTE un objeto JSON válido (sin texto ni markdown) con esta forma exacta y conciso:
{"cliente":"string (nombre del cliente detectado, o '')","pais":"string (PE o EC; o el nombre del país)","tipoProyecto":"string","perfilCliente":"string (investiga al cliente con web_search: a qué se dedica, industria, tamaño y qué ofrecerle; 2-3 frases)","equipo":[{"perfil":"string","tier":"senior|semi|analista","rol":"string"}],"entregables":[{"nombre":"string","hS":number,"hM":number,"hA":number,"semanas":number}],"cronograma":[{"hito":"string","pct":number}],"analisisCompetencia":"string (2-3 frases)","valorAgregado":["string","string"],"oportunidadIA":"string (1-2 frases; o 'No aplica')","estrategiaCierre":"string (postura de precio para ganar siendo rentable, qué enfatizar y riesgos; 2-3 frases)"}
Reglas: hS=horas Senior/Arquitecto/PM, hM=horas Semi-Senior, hA=horas Analista, semanas=duración del entregable en semanas (entero ≥1). El país solo puede ser Perú (PE) o Ecuador (EC). Máximo 6 entregables, numéralos. El cronograma suma pct=1.0 (típico 0.35/0.35/0.30). Sé breve para no exceder el límite de tokens. Si el mensaje incluye una PROPUESTA ACTUAL, trata el texto del usuario como INSTRUCCIONES DE MODIFICACIÓN: aplícalas sobre esa propuesta y conserva todo lo que el usuario no pida cambiar (devuelve igualmente el JSON completo).`;
    const currentProposal = (est.team.some((t) => t.perfil) || est.deliverables.some((d) => d.name))
      ? `\n\nPROPUESTA ACTUAL (aplica sobre esta base las modificaciones pedidas; conserva lo no afectado):\n${JSON.stringify({ equipo: est.team.map((t) => ({ perfil: t.perfil, tier: t.tier, rol: t.rol })), entregables: est.deliverables.map((d) => ({ nombre: d.name, hS: d.hS, hM: d.hM, hA: d.hA, semanas: d.dur })) })}`
      : "";
    const usr = `País sugerido: ${c.name}. Tarifas EBIM en USD/h — Senior (costo ${c.rates.senior.cost} / venta ${c.rates.senior.sale}), Semi (costo ${c.rates.semi.cost} / venta ${c.rates.semi.sale}), Analista (costo ${c.rates.analista.cost} / venta ${c.rates.analista.sale}).
CONTEXTO E INSTRUCCIONES DEL USUARIO (EBIM):\n${est.context || "(ver documentos adjuntos)"}${currentProposal}`;

    // Texto de los documentos legibles (BBP en Word, transcripciones, actas, CSV…)
    const textDocs = usable.filter((a) => a.kind === "text" || a.kind === "docx");
    let docsText = "";
    if (textDocs.length) {
      docsText = "\n\n===== DOCUMENTOS ADJUNTOS =====\n" +
        textDocs.map((a) => `\n----- ${a.name} -----\n${a.text}`).join("\n");
    }
    // PDFs e imágenes se envían como bloques nativos para que la IA los lea
    const mediaBlocks = usable
      .filter((a) => a.kind === "pdf" || a.kind === "image")
      .map((a) => a.kind === "pdf"
        ? { type: "document", source: { type: "base64", media_type: a.mediaType, data: a.base64 } }
        : { type: "image", source: { type: "base64", media_type: a.mediaType, data: a.base64 } });

    const content = [{ type: "text", text: usr + docsText }, ...mediaBlocks];

    try {
      const { data, error: fnError } = await supabase.functions.invoke("claude-proxy", {
        body: {
          model: "claude-sonnet-4-6",
          max_tokens: 1000,
          system: sys,
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
      setEst((e) => ({
        ...e,
        client: e.client || parsed.cliente || "",
        country: detKey || e.country,
        fx: detKey ? cc.fx : e.fx,
        rates: detKey ? cloneRates(cc.rates) : e.rates,
        project: e.project || parsed.tipoProyecto || "",
        team: (parsed.equipo || []).map((t) => ({ id: crypto.randomUUID(), perfil: t.perfil, tier: ["senior", "semi", "analista"].includes(t.tier) ? t.tier : "semi", rol: t.rol || "" })),
        deliverables: (() => {
          let acc = 0;
          return (parsed.entregables || []).map((d) => {
            const dur = Math.max(1, Math.round(+d.semanas || 1));
            const item = { id: crypto.randomUUID(), name: d.nombre, hS: +d.hS || 0, hM: +d.hM || 0, hA: +d.hA || 0, start: acc, dur };
            acc += dur;
            return item;
          });
        })(),
        schedule: (parsed.cronograma && parsed.cronograma.length ? parsed.cronograma : e.schedule.map((s) => ({ hito: s.hito, pct: s.pct }))).map((s) => ({ id: crypto.randomUUID(), hito: s.hito, pct: +s.pct || 0 })),
        insight: {
          perfilCliente: parsed.perfilCliente || "",
          competencia: parsed.analisisCompetencia || "",
          valor: Array.isArray(parsed.valorAgregado) ? parsed.valorAgregado : (parsed.valorAgregado ? [parsed.valorAgregado] : []),
          ia: parsed.oportunidadIA || "",
          estrategiaCierre: parsed.estrategiaCierre || "",
        },
      }));
      setDetected({ cliente: parsed.cliente || "", pais: detKey, paisNombre: COUNTRIES[detKey]?.name || parsed.pais || "", tipo: parsed.tipoProyecto || "" });
      flash("Detectado y generado — revisa la propuesta");
    } catch (err) {
      setAiError("No se pudo generar el borrador automáticamente. Puedes cargar los entregables manualmente. (" + (err?.message || "error") + ")");
    } finally { setAiLoading(false); }
  }

  /* ---------- Excel ---------- */
  function exportExcel() {
    const c = country;
    const scen = [0, 0.05, 0.10, 0.15].map((d) => scenarioRow(calc, d, c.minRent));
    const wb = XLSX.utils.book_new();

    // Hoja 1: Estimación
    const A = [];
    A.push([`HOJA DE COSTOS — ${est.code}  |  ${est.project || "Proyecto"}  |  ${est.client || "Cliente"}`]);
    A.push([`Elaborado por: GRUPO EBIM SAC  |  País: ${c.name}  |  Moneda: USD  |  T.C. ref: ${est.fx} ${c.cur}/USD`]);
    A.push([]);
    A.push(["SUPUESTOS — TODO EN USD"]);
    A.push(["Perfil", "Costo país (USD/h)", "Precio venta (USD/h)"]);
    TIERS.forEach(([k, label]) => A.push([label, est.rates[k].cost, est.rates[k].sale]));
    A.push(["Margen objetivo / piso (% s/ PV)", est.margin, ""]);
    A.push(["Gastos Adm. + Comercial + MKT (% s/ PV)", est.adm + est.com + est.mkt, ""]);
    A.push([]);
    A.push(["EQUIPO DEL PROYECTO"]);
    A.push(["Perfil", "Venta USD/h", "Rol en el proyecto"]);
    est.team.forEach((t) => A.push([t.perfil, est.rates[t.tier].sale, t.rol]));
    A.push([]);
    A.push(["DETALLE DE HORAS Y COSTOS POR ENTREGABLE"]);
    A.push(["Entregable / Actividad", "Hrs Senior", "Hrs Semi", "Hrs Analista", "Total Hrs", "Costo USD", "Venta USD", "Venta " + c.cur]);
    calc.rows.forEach((d) => A.push([d.name, d.hS, d.hM, d.hA, d.hrs, +d.cost.toFixed(2), +d.sale.toFixed(2), +(d.sale * est.fx).toFixed(2)]));
    A.push(["TOTAL", "", "", "", calc.totalHrs, +calc.CO.toFixed(2), +calc.PVO.toFixed(2), +(calc.PVO * est.fx).toFixed(2)]);
    A.push([]);
    A.push(["RESUMEN COMERCIAL Y PRECIO DE VENTA"]);
    A.push(["Concepto", "USD", "% / Factor", c.cur]);
    A.push(["Costo Total Operaciones (costo país)", +calc.CO.toFixed(2), "—", +(calc.CO * est.fx).toFixed(2)]);
    A.push(["Precio de Venta Operaciones (competitivo país)", +calc.PVO.toFixed(2), "—", +(calc.PVO * est.fx).toFixed(2)]);
    A.push(["Margen bruto resultante", "—", +calc.grossMargin.toFixed(4), "—"]);
    A.push(["Gastos Administrativos", +calc.gAdm.toFixed(2), est.adm, +(calc.gAdm * est.fx).toFixed(2)]);
    A.push(["Gestión Comercial — Carmen", +calc.gCom.toFixed(2), est.com, +(calc.gCom * est.fx).toFixed(2)]);
    A.push(["Gestión MKT", +calc.gMkt.toFixed(2), est.mkt, +(calc.gMkt * est.fx).toFixed(2)]);
    A.push(["Total Gastos Adm. y Comisiones", +calc.gastos.toFixed(2), "—", +(calc.gastos * est.fx).toFixed(2)]);
    A.push(["Descuento comercial (%)", "—", est.discount, "—"]);
    A.push(["PRECIO DE VENTA FINAL (Sin IGV)", +calc.PVfinal.toFixed(2), "", +(calc.PVfinal * est.fx).toFixed(2)]);
    A.push(["Utilidad Neta", +calc.utilidad.toFixed(2), "—", +(calc.utilidad * est.fx).toFixed(2)]);
    A.push(["Rentabilidad Neta (%)", +calc.rent.toFixed(4), "—", "—"]);
    A.push([]);
    A.push(["CRONOGRAMA DE PAGOS"]);
    A.push(["N°", "Hito", "% del Total", "Importe USD", "Importe " + c.cur]);
    est.schedule.forEach((s, i) => A.push([i + 1, s.hito, s.pct, +(calc.PVfinal * s.pct).toFixed(2), +(calc.PVfinal * s.pct * est.fx).toFixed(2)]));
    A.push(["", "TOTAL", est.schedule.reduce((a, b) => a + b.pct, 0), +calc.PVfinal.toFixed(2), +(calc.PVfinal * est.fx).toFixed(2)]);
    const ws1 = XLSX.utils.aoa_to_sheet(A);
    ws1["!cols"] = [{ wch: 46 }, { wch: 18 }, { wch: 18 }, { wch: 16 }, { wch: 12 }, { wch: 14 }, { wch: 14 }, { wch: 16 }];
    XLSX.utils.book_append_sheet(wb, ws1, "Estimación");

    // Hoja 2: Escenarios de Descuento
    const B = [];
    B.push([`ESCENARIOS DE DESCUENTO — ${est.code} | ${est.client || "Cliente"} | USO INTERNO EBIM`]);
    B.push([`País: ${c.name}  |  Rentabilidad mínima aceptable: ${pct(c.minRent)}  |  Descuento sugerido máx.: ${pct(c.maxDesc)}`]);
    B.push([]);
    B.push(["Escenario", "Descuento %", "Precio USD", "Precio " + c.cur, "Utilidad USD", "Rentabilidad", "Recomendación"]);
    scen.forEach((s) => B.push([
      s.d === 0 ? "Sin descuento" : "Descuento " + (s.d * 100) + "%",
      s.d, +s.price.toFixed(2), +(s.price * est.fx).toFixed(2), +s.util.toFixed(2), +s.rent.toFixed(4),
      (s.light === "ok" ? "OK — " : s.light === "warn" ? "REVISAR — " : "NO — ") + s.txt,
    ]));
    B.push([]);
    B.push(["Nota: hoja de uso interno. El cliente recibe solo el precio llave en mano, sin desglose de descuentos ni márgenes."]);
    const ws2 = XLSX.utils.aoa_to_sheet(B);
    ws2["!cols"] = [{ wch: 18 }, { wch: 12 }, { wch: 14 }, { wch: 16 }, { wch: 14 }, { wch: 14 }, { wch: 44 }];
    XLSX.utils.book_append_sheet(wb, ws2, "Escenarios de Descuento");

    // Hoja 3: Cronograma de trabajo (Gantt)
    const G = [];
    G.push([`CRONOGRAMA DE TRABAJO — ${est.code} | ${est.client || "Cliente"}`]);
    G.push([`Duración total estimada: ${calc.weeksTotal} semana(s)`]);
    G.push([]);
    G.push(["Entregable", "Semana inicio", "Duración (sem)", "Semana fin", "Total hrs", "Costo USD", "Venta USD"]);
    calc.rows.forEach((d) => G.push([d.name, (d.start || 0) + 1, d.dur || 1, (d.start || 0) + (d.dur || 1), d.hrs, +d.cost.toFixed(2), +d.sale.toFixed(2)]));
    const wsg = XLSX.utils.aoa_to_sheet(G);
    wsg["!cols"] = [{ wch: 46 }, { wch: 14 }, { wch: 16 }, { wch: 12 }, { wch: 10 }, { wch: 14 }, { wch: 14 }];
    XLSX.utils.book_append_sheet(wb, wsg, "Cronograma");

    // Hoja 4: Estrategia comercial (cliente, competencia, valor agregado, IA)
    const C3 = [];
    C3.push([`ESTRATEGIA COMERCIAL — ${est.code} | ${est.client || "Cliente"} | ${c.name}`]);
    C3.push([]);
    C3.push(["Perfil del cliente / qué ofrecerle"]);
    C3.push([est.insight?.perfilCliente || "—"]);
    C3.push([]);
    C3.push(["Análisis de competencia"]);
    C3.push([est.insight?.competencia || "—"]);
    C3.push([]);
    C3.push(["Valor agregado a proponer"]);
    (est.insight?.valor?.length ? est.insight.valor : ["—"]).forEach((v) => C3.push(["• " + v]));
    C3.push([]);
    C3.push(["Oportunidad de IA / automatización"]);
    C3.push([est.insight?.ia || "—"]);
    C3.push([]);
    C3.push(["★ Estrategia de cierre (uso interno)"]);
    C3.push([est.insight?.estrategiaCierre || "—"]);
    const ws3 = XLSX.utils.aoa_to_sheet(C3);
    ws3["!cols"] = [{ wch: 110 }];
    XLSX.utils.book_append_sheet(wb, ws3, "Estrategia comercial");

    const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    const blob = new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = `${est.code}_${(est.client || "cliente").replace(/\s+/g, "_")}.xlsx`;
    a.click(); URL.revokeObjectURL(url);
    flash("Excel descargado");
  }

  /* ---------- Propuesta comercial al cliente ---------- */
  async function generateProposal() {
    if (!est.client.trim() || !est.deliverables.some((d) => d.name)) { flash("Completa cliente y entregables primero"); return; }
    setProposalLoading(true);
    let prose = defaultProse(est);
    try {
      const c = country;
      const sys = `Eres redactor de propuestas comerciales GANADORAS de GRUPO EBIM. Escribe en español, tono profesional, cálido y persuasivo, centrado en el valor y los resultados para el cliente (no en lo técnico interno). NUNCA menciones costos internos, márgenes, descuentos ni tarifas. Devuelve EXCLUSIVAMENTE un JSON válido sin markdown: {"entendimiento":"2-3 frases que demuestran que entendemos su necesidad","enfoque":"2-3 frases sobre cómo lo abordamos","diferenciadores":["3 a 5 bullets de valor"],"porQueEbim":"2-3 frases convincentes","cierre":"1-2 frases de llamada a la acción"}.`;
      const usr = `Cliente: ${est.client}. País: ${c.name}. Servicio: ${est.project}. Perfil del cliente: ${est.insight?.perfilCliente || "—"}. Entregables: ${est.deliverables.filter((d) => d.name).map((d) => d.name).join("; ")}. Valor agregado: ${(est.insight?.valor || []).join("; ")}. Oportunidad IA: ${est.insight?.ia || ""}. Contexto: ${(est.context || "").slice(0, 1500)}`;
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
    setProposalHTML(buildProposalHTML(est, compute(est), country, prose));
    setShowProposal(true);
    setProposalLoading(false);
    flash("Propuesta comercial lista");
  }

  function downloadProposal() {
    const blob = new Blob([proposalHTML], { type: "text/html" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = `Propuesta_${est.code}_${(est.client || "cliente").replace(/\s+/g, "_")}.html`;
    a.click(); URL.revokeObjectURL(url);
  }

  const scen = [0, 0.05, 0.10, 0.15].map((d) => scenarioRow(calc, d, country.minRent));
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
        .wrap{max-width:1180px;margin:0 auto;padding:22px 18px 80px;}
        .topbar{display:flex;align-items:center;gap:14px;flex-wrap:wrap;border-bottom:1px solid var(--line);padding-bottom:16px;margin-bottom:20px;}
        .brand{display:flex;align-items:center;gap:10px;}
        .brandmark{width:34px;height:34px;border-radius:8px;background:var(--accent);color:#fff;display:grid;place-items:center;font-weight:700;}
        .reqchip{font-family:'JetBrains Mono',monospace;background:var(--ink);color:#fff;padding:6px 12px;border-radius:7px;font-weight:600;letter-spacing:.5px;}
        .btn{display:inline-flex;align-items:center;gap:7px;border:1px solid var(--line);background:var(--surface);color:var(--ink);padding:8px 13px;border-radius:8px;font-weight:600;font-size:13px;cursor:pointer;transition:.12s;}
        .btn:hover{border-color:var(--accent);color:var(--accent);}
        .btn.primary{background:var(--accent);color:#fff;border-color:var(--accent);}
        .btn.primary:hover{filter:brightness(1.08);color:#fff;}
        .btn.ghost{background:transparent;}
        .grid{display:grid;gap:16px;}
        .card{background:var(--surface);border:1px solid var(--line);border-radius:12px;padding:16px 18px;}
        .card h3{font-family:'Space Grotesk',sans-serif;font-size:12px;letter-spacing:.10em;text-transform:uppercase;color:var(--muted);margin:0 0 12px;display:flex;align-items:center;gap:8px;}
        label{display:block;font-size:11px;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.05em;margin-bottom:5px;}
        input,select,textarea{width:100%;border:1px solid var(--line);border-radius:8px;padding:8px 10px;font-size:14px;font-family:inherit;background:#fff;color:var(--ink);}
        input:focus,select:focus,textarea:focus{outline:none;border-color:var(--accent);box-shadow:0 0 0 3px var(--accent-soft);}
        .num{font-family:'JetBrains Mono',monospace;text-align:right;}
        table{width:100%;border-collapse:collapse;}
        th{font-size:10.5px;text-transform:uppercase;letter-spacing:.05em;color:var(--muted);text-align:right;padding:7px 8px;border-bottom:1px solid var(--line);font-weight:600;}
        th.l,td.l{text-align:left;}
        td{padding:6px 8px;border-bottom:1px solid #EEF1F4;font-size:13px;}
        td .num{width:100%;border:1px solid transparent;background:transparent;padding:4px 6px;border-radius:6px;}
        td .num:hover{border-color:var(--line);}
        td .num:focus{background:#fff;}
        tr.total td{font-weight:700;border-top:2px solid var(--ink);border-bottom:none;background:#F8FAFB;}
        .row2{display:grid;grid-template-columns:1fr 1fr;gap:14px;}
        .row3{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;}
        .row4{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;}
        @media(max-width:760px){.row2,.row3,.row4{grid-template-columns:1fr;}}
        .hero{display:grid;grid-template-columns:1.3fr 1fr;gap:18px;}
        @media(max-width:760px){.hero{grid-template-columns:1fr;}}
        .price{font-family:'Space Grotesk',sans-serif;font-size:46px;font-weight:700;letter-spacing:-1px;line-height:1;}
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
        .hint{font-size:11.5px;color:var(--muted);margin-top:4px;}
        .dropzone{display:flex;align-items:center;gap:12px;border:1.5px dashed var(--line);border-radius:10px;padding:14px 16px;cursor:pointer;color:var(--muted);background:#fff;transition:.12s;}
        .dropzone:hover,.dropzone.over{border-color:var(--accent);color:var(--accent);background:var(--accent-soft);}
        .filelist{display:grid;grid-template-columns:repeat(auto-fill,minmax(220px,1fr));gap:8px;margin-top:10px;}
        .filechip{display:flex;align-items:center;gap:9px;background:#fff;border:1px solid var(--line);border-radius:9px;padding:8px 10px;font-size:12.5px;}
        .filechip.err{border-color:#F0C9C9;background:#FCF1F1;color:var(--bad);}
        .filechip .iconbtn:hover{color:var(--bad);}
      `}</style>

      <div className="wrap">
        {/* Topbar */}
        <div className="topbar">
          <div className="brand">
            <div className="brandmark disp">E</div>
            <div>
              <div className="disp" style={{ fontWeight: 700, fontSize: 16 }}>EBIM · Estimador de Costos</div>
              <div style={{ fontSize: 11, color: "var(--muted)" }}>Consultoría TI · cotización en USD</div>
            </div>
          </div>
          <span className="reqchip">{est.code}</span>
          <div style={{ marginLeft: "auto", display: "flex", gap: 8, flexWrap: "wrap" }}>
            <button className="btn ghost" onClick={newEstimation}><RotateCcw size={15} /> Nueva</button>
            <button className="btn" onClick={() => setShowHistory(true)}><History size={15} /> Historial ({history.length})</button>
            <button className="btn" onClick={saveEstimation}><Save size={15} /> Guardar</button>
            <button className="btn" onClick={generateProposal} disabled={proposalLoading}>
              {proposalLoading ? <Loader2 size={15} className="spin" /> : <FileText size={15} />} Propuesta
            </button>
            <button className="btn primary" onClick={exportExcel}><Download size={15} /> Excel</button>
            <button className="btn ghost" aria-label="Cerrar sesión" title="Cerrar sesión" onClick={() => supabase.auth.signOut()}><LogOut size={15} /></button>
          </div>
        </div>

        <div className="grid">
          {/* Identificación */}
          <div className="card">
            <h3><Building2 size={14} /> Identificación del requerimiento</h3>
            <div className="row4">
              <div>
                <label>Cliente</label>
                <input value={est.client} onChange={(e) => up({ client: e.target.value })} placeholder="p. ej. BESCO S.A.C." />
                {est.client.trim() && clientCount(est.client) > 0 && (
                  <div className="hint">Este cliente tiene <b>{clientCount(est.client)}</b> cotización(es) guardada(s).</div>
                )}
              </div>
              <div>
                <label>Proyecto</label>
                <input value={est.project} onChange={(e) => up({ project: e.target.value })} placeholder="Diagnóstico de Arquitectura TI" />
              </div>
              <div>
                <label><Globe size={11} style={{ display: "inline", marginRight: 4 }} /> País (mercado)</label>
                <select value={est.country} onChange={(e) => changeCountry(e.target.value)}>
                  {Object.entries(COUNTRIES).map(([k, c]) => <option key={k} value={k}>{c.flag} {c.name}</option>)}
                </select>
              </div>
              <div>
                <label>T.C. referencial ({country.cur}/USD)</label>
                <input className="num" type="number" step="0.01" value={est.fx} onChange={(e) => up({ fx: +e.target.value })} />
                <div className="hint">Informativo. La cotización es siempre en USD.</div>
              </div>
            </div>
            <div className="note" style={{ marginTop: 12 }}>
              <Globe size={15} style={{ flexShrink: 0, marginTop: 1 }} />
              <span><b>{country.name}:</b> {country.note} Rentabilidad neta mínima sugerida <b>{pct(country.minRent)}</b>, descuento máx. competitivo <b>{pct(country.maxDesc)}</b>.</span>
            </div>
          </div>

          {/* IA */}
          <div className="card aiwrap">
            <h3><Sparkles size={14} /> Borrador inteligente · tu prompt</h3>
            <label>Tu prompt / instrucciones (alcance, modificaciones, consideraciones…)</label>
            <textarea rows={4} value={est.context} onChange={(e) => up({ context: e.target.value })}
              placeholder="Escribe aquí como si hablaras conmigo: 'estima este BBP', 'agrega capacitación', 'sube las horas del PM', 'aplica 8% de descuento', 'enfócalo en migración cloud'… Adjunta documentos abajo y pulsa Analizar." />
            <div className="hint" style={{ marginTop: 6 }}>Cada vez que pulses <b>Analizar documentos y generar propuesta</b>, aplico tus instrucciones sobre la propuesta actual (si ya existe) en lugar de empezar de cero.</div>

            <label style={{ marginTop: 14 }}>Documentos del requerimiento</label>
            <div
              className={"dropzone" + (dragOver ? " over" : "")}
              onClick={() => fileRef.current?.click()}
              onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
              onDragLeave={() => setDragOver(false)}
              onDrop={(e) => { e.preventDefault(); setDragOver(false); addFiles(e.dataTransfer.files); }}
            >
              <UploadCloud size={20} />
              <div>
                <b>Adjunta BBP, transcripciones de Teams, actas, propuestas…</b>
                <div className="hint">Arrastra aquí o haz clic. Word (.docx), PDF, imágenes, y texto (.txt .md .vtt .srt .csv .json).</div>
              </div>
              <input ref={fileRef} type="file" multiple style={{ display: "none" }}
                accept=".docx,.pdf,.txt,.md,.markdown,.vtt,.srt,.csv,.tsv,.json,.log,.html,.png,.jpg,.jpeg,.gif,.webp"
                onChange={(e) => { addFiles(e.target.files); e.target.value = ""; }} />
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

            <div style={{ display: "flex", alignItems: "center", gap: 12, marginTop: 12, flexWrap: "wrap" }}>
              <button className="btn primary" onClick={generateDraft} disabled={aiLoading || parsing}>
                {aiLoading ? <Loader2 size={15} className="spin" /> : <Sparkles size={15} />}
                {aiLoading ? "Analizando…" : "Analizar documentos y generar propuesta"}
              </button>
              <span className="hint">Detecta automáticamente servicio, cliente y país, y completa equipo, entregables, cronograma y estrategia. Aprende del historial de EBIM.</span>
            </div>
            {detected && (
              <div className="note" style={{ marginTop: 10, background: "var(--accent-soft)", color: "var(--accent)" }}>
                <CheckCircle2 size={15} style={{ flexShrink: 0 }} />
                <span>Detectado automáticamente — Servicio: <b>{detected.tipo || "—"}</b> · Cliente: <b>{detected.cliente || "no identificado"}</b> · País: <b>{detected.paisNombre || "no identificado"}</b>{!detected.pais && " (no está en la lista; el país no se cambió)"}. Puedes corregir cualquier campo.</span>
              </div>
            )}
            {aiError && <div className="note" style={{ marginTop: 10, background: "#FBE6E6", color: "var(--bad)" }}><AlertTriangle size={15} /> {aiError}</div>}
          </div>

          {/* Inteligencia comercial */}
          <div className="card">
            <h3><Sparkles size={14} /> Inteligencia comercial · {country.name}</h3>
            <div style={{ marginBottom: 12 }}>
              <label>Perfil del cliente · a qué se dedica y qué ofrecerle (investigado por IA)</label>
              <textarea rows={2} value={est.insight.perfilCliente} onChange={(e) => up({ insight: { ...est.insight, perfilCliente: e.target.value } })}
                placeholder="La IA investiga al cliente: industria, tamaño, contexto y oportunidades de servicio…" />
            </div>
            <div className="row3">
              <div>
                <label>Análisis de competencia</label>
                <textarea rows={5} value={est.insight.competencia} onChange={(e) => up({ insight: { ...est.insight, competencia: e.target.value } })}
                  placeholder="Quién compite en este país, rango de precios típico y cómo diferenciarse…" />
              </div>
              <div>
                <label>Valor agregado a proponer (una idea por línea)</label>
                <textarea rows={5} value={(est.insight.valor || []).join("\n")} onChange={(e) => up({ insight: { ...est.insight, valor: e.target.value.split("\n").filter((x) => x.trim()) } })}
                  placeholder="Entregables o servicios extra que justifican el precio…" />
              </div>
              <div>
                <label>Oportunidad de IA / automatización</label>
                <textarea rows={5} value={est.insight.ia} onChange={(e) => up({ insight: { ...est.insight, ia: e.target.value } })}
                  placeholder="Copilotos, chatbots, automatización, BI, agentes… o 'No aplica'." />
              </div>
            </div>
            <div style={{ marginTop: 12 }}>
              <label>★ Estrategia de cierre (recomendación del estratega — uso interno)</label>
              <textarea rows={3} value={est.insight.estrategiaCierre} onChange={(e) => up({ insight: { ...est.insight, estrategiaCierre: e.target.value } })}
                placeholder="Postura de precio para ganar siendo rentable, qué enfatizar ante este cliente y riesgos a cuidar…"
                style={{ borderColor: "var(--gold)", background: "#FFFDF6" }} />
            </div>
            <div className="hint">La IA completa esto al generar el borrador; siempre investiga al cliente, analiza la competencia (Perú/Ecuador), sugiere valor agregado, propone IA y recomienda cómo cerrar. Puedes editarlo libremente.</div>
          </div>

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
            <table>
              <thead><tr><th className="l">Perfil</th><th className="l" style={{ width: 130 }}>Tarifa (tier)</th><th className="l">Rol</th><th style={{ width: 36 }}></th></tr></thead>
              <tbody>
                {est.team.map((t, i) => (
                  <tr key={t.id}>
                    <td className="l"><input value={t.perfil} onChange={(e) => { const team = [...est.team]; team[i] = { ...t, perfil: e.target.value }; up({ team }); }} /></td>
                    <td className="l">
                      <select value={t.tier} onChange={(e) => { const team = [...est.team]; team[i] = { ...t, tier: e.target.value }; up({ team }); }}>
                        <option value="senior">Senior · venta {fmtUSD2(est.rates.senior.sale)}</option>
                        <option value="semi">Semi · venta {fmtUSD2(est.rates.semi.sale)}</option>
                        <option value="analista">Analista · venta {fmtUSD2(est.rates.analista.sale)}</option>
                      </select>
                    </td>
                    <td className="l"><input value={t.rol} onChange={(e) => { const team = [...est.team]; team[i] = { ...t, rol: e.target.value }; up({ team }); }} /></td>
                    <td><button className="iconbtn" aria-label={`Quitar perfil ${t.perfil || ""}`} title="Quitar perfil" onClick={() => up({ team: est.team.filter((x) => x.id !== t.id) })}><Trash2 size={15} /></button></td>
                  </tr>
                ))}
              </tbody>
            </table>
            <button className="btn ghost" style={{ marginTop: 10 }} onClick={() => up({ team: [...est.team, { id: crypto.randomUUID(), perfil: "", tier: "semi", rol: "" }] })}><Plus size={15} /> Añadir perfil</button>
          </div>

          {/* Entregables */}
          <div className="card">
            <h3>Detalle de horas y costos por entregable</h3>
            <table>
              <thead><tr>
                <th className="l">Entregable / actividad</th>
                <th>Hrs Sr</th><th>Hrs Semi</th><th>Hrs Anl</th><th>Total hrs</th><th>Costo USD</th><th>Venta USD</th><th style={{ width: 36 }}></th>
              </tr></thead>
              <tbody>
                {calc.rows.map((d, i) => (
                  <tr key={d.id}>
                    <td className="l"><input value={d.name} onChange={(e) => { const dd = [...est.deliverables]; dd[i] = { ...est.deliverables[i], name: e.target.value }; up({ deliverables: dd }); }} placeholder="Describe el entregable" /></td>
                    {["hS", "hM", "hA"].map((f) => (
                      <td key={f}><input className="num" type="number" value={est.deliverables[i][f] || ""} onChange={(e) => { const dd = [...est.deliverables]; dd[i] = { ...est.deliverables[i], [f]: +e.target.value }; up({ deliverables: dd }); }} /></td>
                    ))}
                    <td className="num">{d.hrs}</td>
                    <td className="num" style={{ color: "var(--muted)" }}>{fmtUSD2(d.cost)}</td>
                    <td className="num">{fmtUSD2(d.sale)}</td>
                    <td><button className="iconbtn" aria-label={`Quitar entregable ${d.name || ""}`} title="Quitar entregable" onClick={() => up({ deliverables: est.deliverables.filter((x) => x.id !== d.id) })}><Trash2 size={15} /></button></td>
                  </tr>
                ))}
                <tr className="total">
                  <td className="l">TOTAL</td><td></td><td></td><td></td>
                  <td className="num">{calc.totalHrs}</td><td className="num" style={{ color: "var(--muted)" }}>{fmtUSD2(calc.CO)}</td><td className="num">{fmtUSD2(calc.PVO)}</td><td></td>
                </tr>
              </tbody>
            </table>
            <button className="btn ghost" style={{ marginTop: 10 }} onClick={() => up({ deliverables: [...est.deliverables, blankDeliverable()] })}><Plus size={15} /> Añadir entregable</button>
          </div>

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
                  {est.deliverables.map((d, i) => {
                    const start = d.start || 0, dur = d.dur || 1;
                    return (
                      <tr key={d.id}>
                        <td className="l" style={{ fontSize: 12.5 }}>{d.name || <span style={{ color: "var(--muted)" }}>Entregable {i + 1}</span>}</td>
                        <td><input className="num" type="number" min="0" value={start} onChange={(e) => { const dd = [...est.deliverables]; dd[i] = { ...d, start: Math.max(0, +e.target.value) }; up({ deliverables: dd }); }} /></td>
                        <td><input className="num" type="number" min="1" value={dur} onChange={(e) => { const dd = [...est.deliverables]; dd[i] = { ...d, dur: Math.max(1, +e.target.value) }; up({ deliverables: dd }); }} /></td>
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

          {/* Parámetros comerciales */}
          <div className="card">
            <h3>Parámetros comerciales (editables)</h3>
            <div className="row4">
              <div><label>Margen objetivo / piso (% s/ PV)</label><input className="num" type="number" step="0.5" value={(est.margin * 100).toFixed(1)} onChange={(e) => up({ margin: +e.target.value / 100 })} /></div>
              <div><label>Gastos administrativos</label><input className="num" type="number" step="0.1" value={(est.adm * 100).toFixed(1)} onChange={(e) => up({ adm: +e.target.value / 100 })} /></div>
              <div><label>Gestión comercial — Carmen</label><input className="num" type="number" step="0.1" value={(est.com * 100).toFixed(1)} onChange={(e) => up({ com: +e.target.value / 100 })} /></div>
              <div><label>Gestión MKT</label><input className="num" type="number" step="0.1" value={(est.mkt * 100).toFixed(1)} onChange={(e) => up({ mkt: +e.target.value / 100 })} /></div>
            </div>
            <div style={{ marginTop: 12, maxWidth: 260 }}>
              <label>Descuento comercial</label>
              <input className="num" type="number" step="1" value={(est.discount * 100).toFixed(0)} onChange={(e) => up({ discount: +e.target.value / 100 })} />
              {est.discount > country.maxDesc && <div className="hint" style={{ color: "var(--warn)" }}>Por encima del descuento competitivo sugerido para {country.name} ({pct(country.maxDesc)}).</div>}
            </div>
            <div className="note" style={{ marginTop: 14, background: calc.grossMargin < est.margin ? "#FBF1DD" : "var(--accent-soft)", color: calc.grossMargin < est.margin ? "var(--warn)" : "var(--accent)" }}>
              {calc.grossMargin < est.margin ? <AlertTriangle size={15} style={{ flexShrink: 0 }} /> : <CheckCircle2 size={15} style={{ flexShrink: 0 }} />}
              <span>Margen bruto del precio competitivo: <b>{pct(calc.grossMargin)}</b> (precio venta {fmtUSD(calc.PVO)} vs. costo país {fmtUSD(calc.CO)}). Piso por margen objetivo: <b>{fmtUSD(calc.floorPrice)}</b>. {calc.grossMargin < est.margin ? "El precio de mercado queda por debajo de tu margen objetivo — decide si compites igual o subes precio." : "El precio competitivo cumple tu margen objetivo."}</span>
            </div>
          </div>

          {/* Hero precio */}
          <div className="card hero">
            <div>
              <div className="pricelabel">★ Precio de venta final (sin IGV) · USD</div>
              <div className="price">{fmtUSD(calc.PVfinal)}</div>
              <div style={{ fontSize: 13, color: "var(--muted)", marginTop: 6 }} className="mono">≈ {fmtLocal(calc.PVfinal * est.fx, country.cur)} · {est.discount > 0 ? `incluye ${pct(est.discount)} descuento` : "sin descuento"}</div>
              <div className="gauge"><i style={{ width: Math.max(3, Math.min(100, (calc.rent / 0.4) * 100)) + "%", background: rentColor }} /></div>
              <div style={{ display: "flex", justifyContent: "space-between", marginTop: 8, fontSize: 12 }}>
                <span style={{ color: "var(--muted)" }}>Rentabilidad neta</span>
                <span className="mono" style={{ fontWeight: 700, color: rentColor }}>{pct(calc.rent)} {calc.rent < country.minRent ? "· bajo el mínimo" : "· ok"}</span>
              </div>
            </div>
            <div>
              <div className="kv"><span>Costo operativo (país)</span><b>{fmtUSD2(calc.CO)}</b></div>
              <div className="kv"><span>Precio venta operaciones</span><b>{fmtUSD2(calc.PVO)}</b></div>
              <div className="kv"><span>Margen bruto</span><b>{pct(calc.grossMargin)}</b></div>
              <div className="kv"><span>Gastos adm. (Carmen, MKT)</span><b>{fmtUSD2(calc.gastos)}</b></div>
              <div className="kv"><span>Descuento aplicado</span><b>−{fmtUSD2(calc.PVO * est.discount)}</b></div>
              <div className="kv"><span>Utilidad neta</span><b style={{ color: rentColor }}>{fmtUSD2(calc.utilidad)}</b></div>
            </div>
          </div>

          {/* Escenarios de descuento */}
          <div className="card">
            <h3>Escenarios de descuento · alineado a {country.name}</h3>
            <table>
              <thead><tr><th className="l">Escenario</th><th>Desc.</th><th>Precio USD</th><th>Precio {country.cur}</th><th>Utilidad</th><th>Rentabilidad</th><th className="l">Recomendación</th></tr></thead>
              <tbody>
                {scen.map((s) => (
                  <tr key={s.d}>
                    <td className="l">{s.d === 0 ? "Sin descuento" : `Descuento ${s.d * 100}%`}</td>
                    <td className="num">{(s.d * 100).toFixed(0)}%</td>
                    <td className="num">{fmtUSD(s.price)}</td>
                    <td className="num">{fmtLocal(s.price * est.fx, country.cur)}</td>
                    <td className="num">{fmtUSD(s.util)}</td>
                    <td className="num" style={{ color: s.light === "ok" ? "var(--ok)" : s.light === "warn" ? "var(--warn)" : "var(--bad)", fontWeight: 600 }}>{pct(s.rent)}</td>
                    <td className="l"><span className={"pill " + s.light}>{s.light === "ok" ? <CheckCircle2 size={12} /> : <AlertTriangle size={12} />}{s.txt}</span></td>
                  </tr>
                ))}
              </tbody>
            </table>
            <div className="note" style={{ marginTop: 12 }}><AlertTriangle size={15} style={{ flexShrink: 0 }} /><span>Hoja de uso interno. El cliente recibe únicamente el precio llave en mano, sin desglose de descuentos ni márgenes. Mínimo aceptable para {country.name}: <b>{pct(country.minRent)}</b> de rentabilidad neta.</span></div>
          </div>

          {/* Cronograma de pagos */}
          <div className="card">
            <h3>Cronograma de pagos</h3>
            <table>
              <thead><tr><th className="l" style={{ width: 40 }}>N°</th><th className="l">Hito</th><th style={{ width: 110 }}>% del total</th><th>Importe USD</th><th>Importe {country.cur}</th><th style={{ width: 36 }}></th></tr></thead>
              <tbody>
                {est.schedule.map((s, i) => (
                  <tr key={s.id}>
                    <td className="l mono">{i + 1}</td>
                    <td className="l"><input value={s.hito} onChange={(e) => { const sc = [...est.schedule]; sc[i] = { ...s, hito: e.target.value }; up({ schedule: sc }); }} /></td>
                    <td><input className="num" type="number" value={(s.pct * 100).toFixed(0)} onChange={(e) => { const sc = [...est.schedule]; sc[i] = { ...s, pct: +e.target.value / 100 }; up({ schedule: sc }); }} /></td>
                    <td className="num">{fmtUSD(calc.PVfinal * s.pct)}</td>
                    <td className="num">{fmtLocal(calc.PVfinal * s.pct * est.fx, country.cur)}</td>
                    <td><button className="iconbtn" aria-label={`Quitar hito ${s.hito || ""}`} title="Quitar hito" onClick={() => up({ schedule: est.schedule.filter((x) => x.id !== s.id) })}><Trash2 size={15} /></button></td>
                  </tr>
                ))}
                <tr className="total">
                  <td></td><td className="l">TOTAL</td>
                  <td className="num" style={{ color: Math.abs(est.schedule.reduce((a, b) => a + b.pct, 0) - 1) > 0.001 ? "var(--bad)" : "inherit" }}>{(est.schedule.reduce((a, b) => a + b.pct, 0) * 100).toFixed(0)}%</td>
                  <td className="num">{fmtUSD(calc.PVfinal * est.schedule.reduce((a, b) => a + b.pct, 0))}</td>
                  <td></td><td></td>
                </tr>
              </tbody>
            </table>
            <button className="btn ghost" style={{ marginTop: 10 }} onClick={() => up({ schedule: [...est.schedule, { id: crypto.randomUUID(), hito: "", pct: 0 }] })}><Plus size={15} /> Añadir hito</button>
          </div>
        </div>
      </div>

      {/* Drawer historial */}
      {showHistory && (
        <div className="drawer" onClick={() => setShowHistory(false)}>
          <div className="panel" onClick={(e) => e.stopPropagation()}>
            <div style={{ display: "flex", alignItems: "center", marginBottom: 16 }}>
              <h3 className="disp" style={{ fontSize: 16, margin: 0 }}>Historial de cotizaciones</h3>
              <button className="iconbtn" aria-label="Cerrar historial" title="Cerrar" style={{ marginLeft: "auto", color: "var(--ink)" }} onClick={() => setShowHistory(false)}><X size={18} /></button>
            </div>
            {history.length === 0 && <div className="note">Aún no hay cotizaciones guardadas. Usa <b>Guardar</b> para registrar la actual.</div>}
            {history.map((h) => (
              <div key={h.code} className="histcard" onClick={() => loadEstimation(h)}>
                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                  <span className="mono" style={{ fontWeight: 700, fontSize: 13 }}>{h.code}</span>
                  <span style={{ marginLeft: "auto", fontSize: 11, color: "var(--muted)" }}>{h.savedAt ? new Date(h.savedAt).toLocaleDateString() : ""}</span>
                  <button className="iconbtn" aria-label={`Eliminar cotización ${h.code}`} title="Eliminar" onClick={(e) => { e.stopPropagation(); deleteEstimation(h.code); }}><Trash2 size={14} /></button>
                </div>
                <div style={{ fontWeight: 600, marginTop: 3 }}>{h.project || "(sin título)"}</div>
                <div style={{ fontSize: 12, color: "var(--muted)" }}>{COUNTRIES[h.country]?.flag} {h.client} · {h.totals ? fmtUSD(h.totals.PVfinal) : ""} · rent. {h.totals ? pct(h.totals.rent) : "—"}</div>
              </div>
            ))}
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
                <button className="btn primary" onClick={downloadProposal}><Download size={15} /> Descargar</button>
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
