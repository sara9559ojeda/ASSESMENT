"use strict";

/**
 * ╔══════════════════════════════════════════════════════════════════════════╗
 * ║  2_POWER_POINT.js — Generador de Presentación CMMI RBKT                ║
 * ║  v2.1: Fix safeStr — sanitización defensiva de outputs del LLM         ║
 * ╚══════════════════════════════════════════════════════════════════════════╝
 */

const pptxgen = require("pptxgenjs");
const { buildHeatMap } = require("./heatmap_pptx");

// ══════════════════════════════════════════════════════════════════════════════
// SECCIÓN 1: PARSEO DEL PAYLOAD
// ══════════════════════════════════════════════════════════════════════════════

const payload     = JSON.parse(process.argv[2]);
const data        = payload.data;
const heatmapData = payload.heatmapData;
const outPath     = payload.outPath;
const logoB64     = payload.logoB64 || "";
const iconB64     = payload.iconB64 || "";
const analysis    = payload.analysis || {};

// Sub-objetos del análisis Llama
const execSummary      = analysis.executive_summary       || null;
const criticalAnalysis = analysis.critical_analysis       || [];
const dashAnalysis     = analysis.dashboard_analysis      || {};
const heatAnalysis     = analysis.heatmap_analysis        || {};
const conclusions      = analysis.conclusions             || {};
const globalRecs       = analysis.global_recommendations  || [];

// ══════════════════════════════════════════════════════════════════════════════
// SECCIÓN 2: PALETA DE COLORES Y CONSTANTES
// ══════════════════════════════════════════════════════════════════════════════

const C = {
  deepBlue:  "0D2B55",
  blue800:   "1A3A6A",
  brandBlue: "2563A8",
  accent:    "4DA3D4",
  surface:   "F1F6FB",
  white:     "FFFFFF",
  textDark:  "2E3A4E",
  textDeep:  "1C2332",
  grey600:   "6B7A8D",
  grey300:   "C8D4DE",
  grey100:   "E2ECF5",
  red:       "C43030",
  orange:    "E06030",
  yellow:    "D49A10",
  green:     "1A8A5E",
  teal:      "0E7490",
  purple:    "6D28D9",
};

const NIST_ORDER = ["GOBERNAR", "IDENTIFICAR", "PROTEGER", "DETECTAR", "RESPONDER", "RECUPERAR"];

const PRIORITY_COLORS = { "ALTA": C.red, "MEDIA": C.orange, "BAJA": C.green };

function orderByNist(processes) {
  return [...processes].sort((a, b) => {
    const nameA = a.name.toUpperCase();
    const nameB = b.name.toUpperCase();
    const idxA = NIST_ORDER.indexOf(nameA);
    const idxB = NIST_ORDER.indexOf(nameB);
    if (idxA === -1 && idxB === -1) return 0;
    if (idxA === -1) return 1;
    if (idxB === -1) return -1;
    return idxA - idxB;
  });
}

const mColor = pct =>
  pct < 30 ? C.red : pct < 50 ? C.orange : pct < 70 ? C.yellow : C.green;

const mLabel = pct =>
  pct < 30 ? "Inexistente"
  : pct < 50 ? "Inicial"
  : pct < 70 ? "En progreso"
  : "Optimizada";

const mkShadow = () => ({
  type: "outer", blur: 6, offset: 2, angle: 135, color: "000000", opacity: 0.10
});

// ══════════════════════════════════════════════════════════════════════════════
// SECCIÓN 2.5: HELPER DEFENSIVO — safeStr
// Convierte cualquier valor a string seguro para slide.addText().
// El LLM puede devolver objetos, arrays o undefined en lugar de strings;
// pptxgenjs lanza "forEach is not a function" si recibe un objeto plano.
// ══════════════════════════════════════════════════════════════════════════════

function safeStr(val) {
  if (val === null || val === undefined) return "";
  if (typeof val === "string") return val;
  if (Array.isArray(val)) return val.map(safeStr).join("\n");
  if (typeof val === "object") {
    // LLM a veces devuelve {titulo, descripcion} en lugar de un string
    if (val.descripcion) return (val.titulo ? val.titulo + ": " : "") + val.descripcion;
    if (val.titulo)      return val.titulo;
    if (val.text)        return String(val.text);
    return JSON.stringify(val);
  }
  return String(val);
}

// Sanitiza un array de strings devuelto por el LLM:
// filtra nulls, convierte objetos a string, elimina vacíos.
function safeStrArray(arr, fallback) {
  if (!Array.isArray(arr) || arr.length === 0) return fallback || [];
  return arr.map(safeStr).filter(Boolean);
}

// ══════════════════════════════════════════════════════════════════════════════
// SECCIÓN 3: HELPERS REUTILIZABLES
// ══════════════════════════════════════════════════════════════════════════════

function addLogo(slide, b64, x, y, w, h) {
  if (!b64) return;
  slide.addImage({ data: "image/png;base64," + b64, x, y, w, h });
}

function addHeader(slide, pres, title, pageNum, total) {
  slide.background = { color: C.surface };
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 13.33, h: 0.82,
    fill: { color: C.deepBlue }, line: { color: C.deepBlue, width: 0 }
  });
  if (logoB64) addLogo(slide, logoB64, 0.25, 0.09, 1.55, 0.64);
  slide.addText("NIST CSF 2.0  ·  ISO/IEC 27001:2022  ·  CMMI  ·  CONFIDENCIAL", {
    x: 2.0, y: 0.08, w: 8, h: 0.22,
    fontSize: 7, color: C.accent, charSpacing: 3, fontFace: "Calibri", margin: 0
  });
  slide.addText(title, {
    x: 2.0, y: 0.28, w: 8.8, h: 0.46,
    fontSize: 20, bold: true, color: C.white, fontFace: "Calibri",
    valign: "middle", margin: 0
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 11.5, y: 0.10, w: 1.65, h: 0.60,
    fill: { color: C.blue800 }, line: { color: C.blue800, width: 0 }
  });
  slide.addText(`${pageNum} / ${total}`, {
    x: 11.5, y: 0.10, w: 1.65, h: 0.60,
    fontSize: 11, color: C.accent, fontFace: "Calibri",
    align: "center", valign: "middle", margin: 0
  });
}

function addFooter(slide, pres) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 7.16, w: 13.33, h: 0.34,
    fill: { color: C.grey100 }, line: { color: C.grey300, width: 0.5 }
  });
  slide.addText("RBKT · Dashboard de Ciberseguridad · NIST CSF 2.0 + ISO/IEC 27001:2022 · Confidencial · 2025", {
    x: 0.4, y: 7.19, w: 10, h: 0.26,
    fontSize: 7.5, color: C.grey600, fontFace: "Calibri", valign: "middle", margin: 0
  });
  if (iconB64) addLogo(slide, iconB64, 12.6, 7.17, 0.32, 0.32);
}

function addKPICard(slide, pres, x, y, w, h, label, value, accent, sub) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w, h,
    fill: { color: C.white }, line: { color: C.grey300, width: 0.5 }, shadow: mkShadow()
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w, h: 0.06,
    fill: { color: accent }, line: { color: accent, width: 0 }
  });
  slide.addText(safeStr(value), {
    x: x+0.1, y: y+0.10, w: w-0.2, h: h*0.48,
    fontSize: 22, bold: true, color: accent, fontFace: "Calibri",
    valign: "bottom", margin: 0
  });
  slide.addText(safeStr(label), {
    x: x+0.1, y: y+h*0.52, w: w-0.2, h: h*0.28,
    fontSize: 8.5, bold: true, color: C.textDark, fontFace: "Calibri",
    valign: "top", margin: 0
  });
  if (sub) {
    slide.addText(safeStr(sub), {
      x: x+0.1, y: y+h*0.76, w: w-0.2, h: h*0.22,
      fontSize: 7.5, color: C.grey600, fontFace: "Calibri",
      valign: "top", margin: 0
    });
  }
}

function secLabel(slide, x, y, text) {
  slide.addText(safeStr(text), {
    x, y, w: 12.5, h: 0.26,
    fontSize: 8, bold: true, color: C.grey600,
    fontFace: "Calibri", charSpacing: 1, margin: 0
  });
}

/**
 * Dibuja un bloque de análisis Llama con borde lateral de color e ícono de comillas.
 */
function addAnalysisBlock(slide, pres, x, y, w, h, text, accentColor, labelText) {
  const safeText = safeStr(text);
  if (!safeText) return;
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w, h,
    fill: { color: C.white }, line: { color: C.grey300, width: 0.5 }, shadow: mkShadow()
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w: 0.06, h,
    fill: { color: accentColor }, line: { color: accentColor, width: 0 }
  });
  if (labelText) {
    slide.addText(safeStr(labelText), {
      x: x + 0.14, y: y + 0.06, w: w - 0.2, h: 0.20,
      fontSize: 7, bold: true, color: accentColor, fontFace: "Calibri",
      charSpacing: 1, margin: 0
    });
  }
  const textY = labelText ? y + 0.26 : y + 0.10;
  const textH = labelText ? h - 0.30 : h - 0.14;
  slide.addText(safeText, {
    x: x + 0.14, y: textY, w: w - 0.22, h: textH,
    fontSize: 8.8, color: C.textDark, fontFace: "Calibri",
    wrap: true, valign: "top", margin: 0, lineSpacingMultiple: 1.15
  });
}

function addCriticalityScale(slide, pres, pct, x, y, w, h) {
  slide.addText("ESCALA DE CRITICIDAD GLOBAL", {
    x, y, w, h: 0.22,
    fontSize: 8, bold: true, color: C.grey600, fontFace: "Calibri", charSpacing: 1, margin: 0,
  });
  const barY = y + 0.30, barH = 0.52, labelH = 0.26;
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y: barY - 0.10, w, h: barH + labelH + 0.60,
    fill: { color: C.white }, line: { color: C.grey300, width: 0.5 }, shadow: mkShadow(),
  });
  const SEGS = 20, segW = (w - 0.20) / SEGS;
  const gradX0 = x + 0.10, gradY = barY + 0.14;
  function hexToRgb(hex) {
    return [parseInt(hex.slice(0,2),16), parseInt(hex.slice(2,4),16), parseInt(hex.slice(4,6),16)];
  }
  function rgbToHex(r,g,b) {
    return [r,g,b].map(v=>Math.round(Math.max(0,Math.min(255,v))).toString(16).padStart(2,"0")).join("").toUpperCase();
  }
  function lerpColorScale(hexA, hexB, t) {
    const a = hexToRgb(hexA), b = hexToRgb(hexB);
    return rgbToHex(a[0]+(b[0]-a[0])*t, a[1]+(b[1]-a[1])*t, a[2]+(b[2]-a[2])*t);
  }
  const colorStops = [C.red, C.orange, C.yellow, C.green];
  for (let i = 0; i < SEGS; i++) {
    const t = i / (SEGS - 1);
    const stopIdx = t * (colorStops.length - 1);
    const lo = Math.floor(stopIdx), hi = Math.min(lo+1, colorStops.length-1);
    const segColor = lerpColorScale(colorStops[lo], colorStops[hi], stopIdx - lo);
    slide.addShape(pres.shapes.RECTANGLE, {
      x: gradX0+i*segW, y: gradY, w: segW+0.01, h: barH-0.10,
      fill: { color: segColor }, line: { color: segColor, width: 0 },
    });
  }
  const gradW = segW * SEGS;
  const markerX = gradX0 + (pct/100) * gradW;
  const markerColor = mColor(pct);
  slide.addShape(pres.shapes.RECTANGLE, {
    x: markerX-0.02, y: gradY-0.01, w: 0.04, h: barH-0.08,
    fill: { color: C.white }, line: { color: C.white, width: 0 },
  });
  const pinR = 0.28, pinX = markerX - pinR/2, pinY = gradY - pinR - 0.04;
  slide.addShape(pres.shapes.OVAL, {
    x: pinX, y: pinY, w: pinR, h: pinR,
    fill: { color: markerColor }, line: { color: C.white, width: 1.5 },
    shadow: { type: "outer", blur: 4, offset: 1, angle: 135, color: "000000", opacity: 0.20 },
  });
  slide.addText(`${pct.toFixed(0)}%`, {
    x: pinX, y: pinY, w: pinR, h: pinR,
    fontSize: 7.5, bold: true, color: C.white, fontFace: "Calibri",
    align: "center", valign: "middle", margin: 0,
  });
  const lblY = gradY + barH - 0.08;
  slide.addText("◀  Crítico", {
    x: gradX0, y: lblY, w: 1.2, h: labelH,
    fontSize: 7.5, bold: true, color: C.red, fontFace: "Calibri", align: "left", valign: "top", margin: 0
  });
  slide.addText("Óptimo  ▶", {
    x: gradX0 + gradW - 1.2, y: lblY, w: 1.2, h: labelH,
    fontSize: 7.5, bold: true, color: C.green, fontFace: "Calibri", align: "right", valign: "top", margin: 0
  });
  slide.addText("50%", {
    x: gradX0 + gradW/2 - 0.3, y: lblY, w: 0.6, h: labelH,
    fontSize: 7, color: C.grey600, fontFace: "Calibri", align: "center", valign: "top", margin: 0
  });
  const stateLabel = pct < 30 ? "CRÍTICO" : pct < 50 ? "DEFICIENTE" : pct < 70 ? "EN PROGRESO" : "ÓPTIMO";
  slide.addText(`Estado global: ${stateLabel}`, {
    x: gradX0, y: lblY + labelH + 0.04, w: gradW, h: 0.22,
    fontSize: 8, bold: true, color: markerColor, fontFace: "Calibri", align: "center", valign: "top", margin: 0,
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// SECCIÓN 4: SLIDE — PORTADA
// ══════════════════════════════════════════════════════════════════════════════

function buildCover(pres, total) {
  const slide = pres.addSlide();
  slide.background = { color: C.deepBlue };
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 5.45, w: 13.33, h: 2.05,
    fill: { color: C.blue800 }, line: { color: C.blue800, width: 0 }
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 5.40, w: 13.33, h: 0.10,
    fill: { color: C.brandBlue }, line: { color: C.brandBlue, width: 0 }
  });
  slide.addText("RBKT CYBERSECURITY", {
    x: 0.7, y: 0.55, w: 9.5, h: 0.32,
    fontSize: 11, color: C.accent, charSpacing: 6, fontFace: "Calibri", bold: true, margin: 0
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.7, y: 1.00, w: 4.2, h: 0.35,
    fill: { color: C.brandBlue }, line: { color: C.brandBlue, width: 0 }
  });
  slide.addText("NIST CSF 2.0  ·  ISO/IEC 27001:2022  ·  CMMI", {
    x: 0.7, y: 1.00, w: 4.2, h: 0.35,
    fontSize: 8.5, color: C.white, fontFace: "Calibri",
    align: "center", valign: "middle", charSpacing: 1, margin: 0
  });
  slide.addText("Ciberseguridad\nCMMI", {
    x: 0.7, y: 1.52, w: 9.8, h: 2.20,
    fontSize: 68, bold: true, color: C.white,
    fontFace: "Calibri", margin: 0, lineSpacingMultiple: 0.85
  });
  slide.addText("Evaluación de Procesos y Nivel de Madurez", {
    x: 0.7, y: 3.82, w: 9.8, h: 0.55,
    fontSize: 15, color: C.accent, fontFace: "Calibri", margin: 0
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.7, y: 4.48, w: 9.8, h: 0.035,
    fill: { color: C.brandBlue }, line: { color: C.brandBlue, width: 0 }
  });
  slide.addText("Clasificación: Confidencial  ·  Versión 2.1  ·  2025", {
    x: 0.7, y: 5.60, w: 9, h: 0.36,
    fontSize: 9, color: C.accent, fontFace: "Calibri", margin: 0
  });
  slide.addText("Área de Ciberseguridad — RBKT", {
    x: 0.7, y: 6.05, w: 9, h: 0.36,
    fontSize: 10.5, color: "90A8C0", fontFace: "Calibri", margin: 0
  });
  slide.addText("1 / " + total, {
    x: 10.5, y: 6.80, w: 2.5, h: 0.40,
    fontSize: 9.5, color: C.grey600, fontFace: "Calibri", align: "right", margin: 0
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// SECCIÓN 5: SLIDE — LEYENDA CMMI
// ══════════════════════════════════════════════════════════════════════════════

function buildLegend(pres, pageNum, total) {
  const slide = pres.addSlide();
  addHeader(slide, pres, "Marco de Referencia — Niveles CMMI", pageNum, total);
  addFooter(slide, pres);
  const items = [
    { color: C.red,    label: "Nivel 1 — Inexistente", desc: "No existen procesos formales ni controles documentados. Las actividades se realizan de manera ad-hoc o directamente no se realizan." },
    { color: C.orange, label: "Nivel 2 — Inicial",     desc: "Existen prácticas básicas pero son informales, inconsistentes y dependientes de individuos clave. No hay estandarización ni documentación mínima." },
    { color: C.yellow, label: "Nivel 3 — Definida",    desc: "Los procesos están documentados, estandarizados y comunicados. Existe una base formal y repetible para la gestión de seguridad." },
    { color: C.green,  label: "Nivel 4 — Optimizada",  desc: "Los procesos son medidos, controlados y mejorados de forma continua. Gestión proactiva orientada a datos y mejora continua." },
  ];
  items.forEach((item, i) => {
    const cy = 1.20 + i * 1.42;
    slide.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: cy, w: 12.5, h: 1.25, fill: { color: C.white }, line: { color: C.grey300, width: 0.5 }, shadow: mkShadow() });
    slide.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: cy, w: 0.5, h: 1.25, fill: { color: item.color }, line: { color: item.color, width: 0 } });
    slide.addShape(pres.shapes.RECTANGLE, { x: 1.1, y: cy+0.22, w: 1.7, h: 0.80, fill: { color: item.color }, line: { color: item.color, width: 0 } });
    slide.addText(item.label, { x: 3.0, y: cy+0.10, w: 9.6, h: 0.38, fontSize: 14, bold: true, color: C.deepBlue, fontFace: "Calibri", margin: 0 });
    slide.addText(item.desc,  { x: 3.0, y: cy+0.50, w: 9.6, h: 0.68, fontSize: 10, color: C.textDark, fontFace: "Calibri", wrap: true, margin: 0 });
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// SECCIÓN 6: SLIDE — RESUMEN EJECUTIVO (con análisis Llama)
// ══════════════════════════════════════════════════════════════════════════════

function buildExecutive(pres, data, pageNum, total) {
  const { processes, globalAvg } = data;
  const slide = pres.addSlide();
  addHeader(slide, pres, "Resumen Ejecutivo", pageNum, total);
  addFooter(slide, pres);
  const scoreColor = mColor(globalAvg);

  // Score global
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.35, y: 1.02, w: 3.5, h: 1.5, fill: { color: C.white }, line: { color: C.grey300, width: 0.5 }, shadow: mkShadow() });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.35, y: 1.02, w: 3.5, h: 0.06, fill: { color: scoreColor }, line: { color: scoreColor, width: 0 } });
  slide.addText("PUNTUACIÓN GLOBAL", { x: 0.50, y: 1.10, w: 3.2, h: 0.28, fontSize: 8, bold: true, color: C.grey600, fontFace: "Calibri", charSpacing: 1, margin: 0 });
  slide.addText(`${globalAvg.toFixed(1)}%`, { x: 0.50, y: 1.36, w: 2.2, h: 0.90, fontSize: 54, bold: true, color: scoreColor, fontFace: "Calibri", valign: "middle", margin: 0 });
  slide.addText(mLabel(globalAvg), { x: 2.5, y: 1.90, w: 1.2, h: 0.38, fontSize: 9, color: scoreColor, bold: true, fontFace: "Calibri", valign: "middle", margin: 0 });

  // Lista de procesos por criticidad
  const cardH = 0.58;
  processes.forEach((p, i) => {
    const cy = 2.72 + i * (cardH + 0.08);
    slide.addShape(pres.shapes.RECTANGLE, { x: 0.35, y: cy, w: 3.5, h: cardH, fill: { color: C.white }, line: { color: C.grey300, width: 0.5 }, shadow: mkShadow() });
    slide.addShape(pres.shapes.RECTANGLE, { x: 0.35, y: cy, w: 0.06, h: cardH, fill: { color: mColor(p.avgPct) }, line: { color: mColor(p.avgPct), width: 0 } });
    slide.addText(`#${i+1}  ${p.name}`, { x: 0.50, y: cy+0.06, w: 2.3, h: 0.28, fontSize: 9.5, bold: true, color: C.textDark, fontFace: "Calibri", margin: 0 });
    slide.addText(mLabel(p.avgPct), { x: 0.50, y: cy+0.30, w: 2.0, h: 0.22, fontSize: 8, color: C.grey600, fontFace: "Calibri", margin: 0 });
    slide.addText(`${p.avgPct.toFixed(1)}%`, { x: 2.60, y: cy+0.08, w: 1.1, h: 0.42, fontSize: 18, bold: true, color: mColor(p.avgPct), fontFace: "Calibri", align: "right", valign: "middle", margin: 0 });
  });

  const RX = 4.1, RW = 9.0;

  // Top 2 riesgos identificados
  secLabel(slide, RX, 1.02, "⚠  PRINCIPALES RIESGOS IDENTIFICADOS");
  processes.slice(0, 2).forEach((p, i) => {
    const ry = 1.32 + i * 1.55;
    slide.addShape(pres.shapes.RECTANGLE, { x: RX, y: ry, w: RW, h: 1.38, fill: { color: C.white }, line: { color: C.grey300, width: 0.5 }, shadow: mkShadow() });
    slide.addShape(pres.shapes.RECTANGLE, { x: RX, y: ry, w: 0.06, h: 1.38, fill: { color: mColor(p.avgPct) }, line: { color: mColor(p.avgPct), width: 0 } });
    const worstSub = p.subs.reduce((a, b) => a.pctg < b.pctg ? a : b);
    slide.addText(`RIESGO #${i+1}  —  ${p.name}`, { x: RX+0.2, y: ry+0.10, w: RW-0.3, h: 0.30, fontSize: 11, bold: true, color: C.deepBlue, fontFace: "Calibri", margin: 0 });
    slide.addText(`Cumplimiento: ${p.avgPct.toFixed(1)}%   ·   Subproceso crítico: ${worstSub.subproceso} (${worstSub.pctg.toFixed(1)}%)`,
      { x: RX+0.2, y: ry+0.44, w: RW-0.3, h: 0.28, fontSize: 9.5, color: C.grey600, fontFace: "Calibri", margin: 0 });
    const dominant = Object.entries(p.maturityCounts).filter(([,v])=>v>0).sort((a,b)=>b[1]-a[1])[0][0];
    slide.addText(`Nivel dominante: ${dominant}. Ausencia o informalidad de controles clave. Requiere atención prioritaria.`,
      { x: RX+0.2, y: ry+0.78, w: RW-0.3, h: 0.46, fontSize: 9, color: C.textDark, fontFace: "Calibri", wrap: true, margin: 0 });
  });

  // Análisis Llama: Resumen Ejecutivo
  const safeExec = safeStr(execSummary);
  if (safeExec) {
    secLabel(slide, RX, 4.48, "💡  RESUMEN EJECUTIVO — ANÁLISIS EXPERTO");
    slide.addShape(pres.shapes.RECTANGLE, { x: RX, y: 4.76, w: RW, h: 2.40, fill: { color: C.white }, line: { color: C.grey300, width: 0.5 }, shadow: mkShadow() });
    slide.addShape(pres.shapes.RECTANGLE, { x: RX, y: 4.76, w: 0.06, h: 2.40, fill: { color: C.brandBlue }, line: { color: C.brandBlue, width: 0 } });
    slide.addText("\u201C", { x: RX+0.14, y: 4.78, w: 0.38, h: 0.36, fontSize: 28, bold: true, color: C.accent, fontFace: "Calibri", valign: "top", margin: 0 });
    slide.addText(safeExec, { x: RX+0.22, y: 4.88, w: RW-0.35, h: 2.20, fontSize: 8.8, color: C.textDark, fontFace: "Calibri", wrap: true, valign: "top", margin: 0, lineSpacingMultiple: 1.15 });
  } else {
    // Fallback a hallazgos derivados
    const best = processes[processes.length - 1];
    secLabel(slide, RX, 4.48, "🔍  HALLAZGOS CLAVE");
    slide.addShape(pres.shapes.RECTANGLE, { x: RX, y: 4.76, w: RW, h: 2.12, fill: { color: C.white }, line: { color: C.grey300, width: 0.5 }, shadow: mkShadow() });
    slide.addShape(pres.shapes.RECTANGLE, { x: RX, y: 4.76, w: 0.06, h: 2.12, fill: { color: C.brandBlue }, line: { color: C.brandBlue, width: 0 } });
    const findings = [
      `Proceso de mayor madurez: ${best.name} con ${best.avgPct.toFixed(1)}% de cumplimiento promedio.`,
      `${processes[0].name} presenta el menor índice (${processes[0].avgPct.toFixed(1)}%), requiriendo atención prioritaria e inversión inmediata en controles formales.`,
      `El índice global del ${globalAvg.toFixed(1)}% indica que la organización opera con brechas materiales en múltiples dominios del NIST CSF 2.0, exponiendo activos críticos a riesgos no mitigados.`,
    ];
    const richFindings = findings.map((f, i) => ([
      { text: `${i+1}.  `, options: { bold: true, color: C.brandBlue, fontSize: 9.5, fontFace: "Calibri" } },
      { text: f + (i < findings.length-1 ? "\n" : ""), options: { color: C.textDark, fontSize: 9.5, fontFace: "Calibri" } },
    ])).flat();
    slide.addText(richFindings, { x: RX+0.22, y: 4.84, w: RW-0.35, h: 1.96 });
  }
}

// ══════════════════════════════════════════════════════════════════════════════
// SECCIÓN 7: SLIDE — DASHBOARD GENERAL (con análisis Llama)
// ══════════════════════════════════════════════════════════════════════════════

function buildDashboard(pres, data, pageNum, total, orderedProcesses) {
  const { globalAvg } = data;
  const processes = orderedProcesses;
  const slide = pres.addSlide();
  addHeader(slide, pres, "Dashboard General — Visión NIST CSF 2.0", pageNum, total);
  addFooter(slide, pres);

  // Radar de cumplimiento
  secLabel(slide, 0.35, 1.00, "CUMPLIMIENTO POR PROCESO (%) — NIST CSF 2.0");
  slide.addChart(pres.charts.RADAR, [{
    name:   "Cumplimiento",
    labels: processes.map(p => p.name),
    values: processes.map(p => p.avgPct),
  }], {
    x: 0.35, y: 1.28, w: 6.0, h: 3.7,
    chartColors: [C.brandBlue],
    chartArea:   { fill: { color: C.white }, roundedCorners: false },
    catAxisLabelColor: C.grey600, valAxisLabelColor: C.grey600,
    valGridLine: { color: C.grey100, size: 0.5 },
    showLegend: false, radarStyle: "standard",
    valAxisMinVal: 0, valAxisMaxVal: 100, valAxisMajorUnit: 25,
    showValue: true, dataLabelFontSize: 8,
    dataLabelColor: C.brandBlue, dataLabelFormatCode: '0"%"',
  });

  addCriticalityScale(slide, pres, globalAvg, 6.55, 4.40, 6.60, 1.70);

  // Barras apiladas por nivel de madurez
  secLabel(slide, 6.55, 1.00, "PREGUNTAS POR NIVEL DE MADUREZ (ISO 27001:2022)");
  const levels = ["1-Inexistente", "2-Inicial", "3-Definida", "4-Optimizada"];
  slide.addChart(pres.charts.BAR, levels.map(lvl => ({
    name:   lvl.split("-")[1],
    labels: processes.map(p => p.name),
    values: processes.map(p => p.maturityCounts[lvl] || 0),
  })), {
    x: 6.55, y: 1.28, w: 6.50, h: 2.90,
    barDir: "col", barGrouping: "stacked",
    chartColors: [C.red, C.orange, C.yellow, C.green],
    chartArea: { fill: { color: C.white }, roundedCorners: false },
    catAxisLabelColor: C.grey600, valAxisLabelColor: C.grey600,
    valGridLine: { color: C.grey100, size: 0.5 }, catGridLine: { style: "none" },
    showLegend: true, legendPos: "b", legendFontSize: 8,
    showValue: true, dataLabelPosition: "ctr",
    dataLabelColor: C.white, dataLabelFontSize: 9, dataLabelFontBold: true,
    dataLabelFormatCode: "#",
  });

  // KPI cards por proceso
  secLabel(slide, 0.35, 5.56, "ÍNDICE DE CUMPLIMIENTO POR PROCESO");
  const cardW = (13.33 - 0.35 - 0.35 - 0.12 * 4) / 5;
  processes.forEach((p, i) => {
    addKPICard(slide, pres,
      0.35 + i * (cardW + 0.12), 5.82, cardW, 1.1,
      p.name, `${p.avgPct.toFixed(1)}%`, mColor(p.avgPct), mLabel(p.avgPct)
    );
  });

  // ── Análisis Llama del Dashboard ──────────────────────────────────────────
  if (dashAnalysis.radar_insight || dashAnalysis.global_posture) {
    const radarText = [
      safeStr(dashAnalysis.radar_insight),
      safeStr(dashAnalysis.kpi_insight)
    ].filter(Boolean).join("  ·  ");

    if (radarText) {
      slide.addShape(pres.shapes.RECTANGLE, {
        x: 0.35, y: 4.42, w: 6.02, h: 1.06,
        fill: { color: C.surface }, line: { color: C.grey300, width: 0.5 }
      });
      slide.addShape(pres.shapes.RECTANGLE, {
        x: 0.35, y: 4.42, w: 0.05, h: 1.06,
        fill: { color: C.brandBlue }, line: { color: C.brandBlue, width: 0 }
      });
      slide.addText("🔍 INTERPRETACIÓN DEL RADAR — NIST CSF 2.0", {
        x: 0.48, y: 4.46, w: 5.80, h: 0.18,
        fontSize: 6.5, bold: true, color: C.brandBlue, fontFace: "Calibri", charSpacing: 1, margin: 0
      });
      slide.addText(radarText, {
        x: 0.48, y: 4.66, w: 5.80, h: 0.78,
        fontSize: 7.8, color: C.textDark, fontFace: "Calibri",
        wrap: true, valign: "top", margin: 0, lineSpacingMultiple: 1.1
      });
    }
  }
}

// ══════════════════════════════════════════════════════════════════════════════
// SECCIÓN 8: SLIDE — PROCESO INDIVIDUAL (con análisis Llama profundo)
// ══════════════════════════════════════════════════════════════════════════════

function generateInsight(proc) {
  const sorted   = [...proc.subs].sort((a, b) => a.pctg - b.pctg);
  const worst    = sorted[0];
  const best     = sorted[sorted.length - 1];
  const counts   = proc.maturityCounts;
  const dominant = Object.entries(counts).filter(([,v])=>v>0).sort((a,b)=>b[1]-a[1])[0][0];
  const inexCount = counts["1-Inexistente"] || 0;
  const optCount  = counts["4-Optimizada"]  || 0;
  let t = `El proceso ${proc.name} presenta un cumplimiento promedio de ${proc.avgPct.toFixed(1)}%, con nivel dominante "${dominant}". `;
  if (inexCount > 0) t += `Se detectaron ${inexCount} pregunta(s) en estado Inexistente. "${worst.subproceso}" es el subproceso más crítico (${worst.pctg.toFixed(1)}%). `;
  else t += `El subproceso de menor desempeño es "${worst.subproceso}" con ${worst.pctg.toFixed(1)}%. `;
  if (optCount > 0) t += `"${best.subproceso}" alcanza ${best.pctg.toFixed(1)}%, demostrando capacidad instalada. `;
  else t += `El mejor rendimiento lo tiene "${best.subproceso}" con ${best.pctg.toFixed(1)}%. `;
  const gap = best.pctg - worst.pctg;
  t += gap > 0 ? `La brecha interna de ${gap.toFixed(1)} puntos evidencia madurez desigual dentro del proceso.` : `Todos los subprocesos presentan el mismo nivel de cumplimiento.`;
  return t;
}

function buildProcess(pres, proc, rank, pageNum, total, allProcesses) {
  const slide = pres.addSlide();
  slide.background = { color: C.white };
  const pc = mColor(proc.avgPct);
  const PW = 3.1;

  // Panel lateral izquierdo (color del proceso)
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: PW, h: 7.5, fill: { color: pc }, line: { color: pc, width: 0 } });
  if (logoB64) addLogo(slide, logoB64, 0.18, 0.12, 2.60, 1.06);
  slide.addText(`PRIORIDAD #${rank} · CMMI RBKT`, { x: 0.22, y: 1.26, w: PW-0.3, h: 0.25, fontSize: 7.5, color: "FFFFFF", fontFace: "Calibri", charSpacing: 2, margin: 0 });
  slide.addText(proc.name, { x: 0.22, y: 1.52, w: PW-0.3, h: 0.60, fontSize: 26, bold: true, color: C.white, fontFace: "Calibri", margin: 0 });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.22, y: 2.16, w: PW-0.44, h: 0.03, fill: { color: "FFFFFF", transparency: 55 }, line: { color: "FFFFFF", width: 0 } });
  slide.addText(`${proc.avgPct.toFixed(1)}%`, { x: 0.22, y: 2.24, w: PW-0.3, h: 1.20, fontSize: 64, bold: true, color: C.white, fontFace: "Calibri", margin: 0 });
  slide.addText(`Nivel: ${mLabel(proc.avgPct)}`, { x: 0.22, y: 3.46, w: PW-0.3, h: 0.32, fontSize: 10, color: "FFFFFF", fontFace: "Calibri", margin: 0 });

  const vals = proc.subs.map(s => s.pctg);
  const maxV = Math.max(...vals), minV = Math.min(...vals);
  const metrics = [
    ["Mejor subproceso", `${maxV.toFixed(1)}%`],
    ["Peor subproceso",  `${minV.toFixed(1)}%`],
    ["Brecha interna",   `${(maxV - minV).toFixed(1)} pts`],
    ["Subprocesos",      `${proc.subs.length}`],
  ];
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.22, y: 3.88, w: PW-0.44, h: 0.03, fill: { color: "FFFFFF", transparency: 55 }, line: { color: "FFFFFF", width: 0 } });
  metrics.forEach(([lbl, val], i) => {
    const my = 4.00 + i * 0.40;
    slide.addText(lbl, { x: 0.22, y: my, w: PW-0.52, h: 0.22, fontSize: 8, color: "FFFFFF", fontFace: "Calibri", margin: 0 });
    slide.addText(val, { x: 0.22, y: my, w: PW-0.52, h: 0.22, fontSize: 9, bold: true, color: C.white, fontFace: "Calibri", align: "right", margin: 0 });
  });

  // Índice lateral de todos los procesos
  const procAnalysis = criticalAnalysis.find(item => item.process === proc.name);
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.22, y: 5.68, w: PW-0.44, h: 0.03, fill: { color: "FFFFFF", transparency: 55 }, line: { color: "FFFFFF", width: 0 } });
  slide.addText("TODOS LOS PROCESOS", { x: 0.22, y: 5.80, w: PW-0.3, h: 0.22, fontSize: 7, color: "FFFFFF", fontFace: "Calibri", charSpacing: 1, margin: 0 });
  allProcesses.forEach((p, i) => {
    const isCurrent = p.name === proc.name;
    const py = 6.06 + i * 0.28;
    slide.addShape(pres.shapes.RECTANGLE, { x: 0.22, y: py, w: 0.15, h: 0.15, fill: { color: "FFFFFF", transparency: isCurrent ? 0 : 65 }, line: { color: "FFFFFF", width: 0 } });
    slide.addText(p.name, { x: 0.44, y: py, w: PW-0.62, h: 0.20, fontSize: 7.5, color: C.white, bold: isCurrent, fontFace: "Calibri", margin: 0 });
    slide.addText(`${p.avgPct.toFixed(1)}%`, { x: 0.44, y: py, w: PW-0.62, h: 0.20, fontSize: 7.5, color: C.white, bold: isCurrent, fontFace: "Calibri", align: "right", margin: 0 });
  });

  // Panel derecho: gráficos + análisis
  const RX = PW + 0.05, RW = 13.33 - RX - 0.3;
  slide.addShape(pres.shapes.RECTANGLE, { x: RX, y: 0, w: 13.33-RX, h: 0.55, fill: { color: C.deepBlue }, line: { color: C.deepBlue, width: 0 } });
  slide.addText("ANÁLISIS PROFUNDO — NIST CSF 2.0 + ISO/IEC 27001:2022", { x: RX+0.25, y: 0, w: RW-0.5, h: 0.55, fontSize: 9, color: C.accent, fontFace: "Calibri", charSpacing: 1, valign: "middle", margin: 0 });
  slide.addText(`${pageNum} / ${total}`, { x: RX+0.25, y: 0, w: RW-0.3, h: 0.55, fontSize: 9, color: C.grey600, fontFace: "Calibri", align: "right", valign: "middle", margin: 0 });

  // Gráfico de barras horizontales por subproceso
  secLabel(slide, RX+0.25, 0.68, "CUMPLIMIENTO POR SUBPROCESO (%)");
  slide.addChart(pres.charts.BAR, [{
    name:   "Cumplimiento (%)",
    labels: proc.subs.map(s => s.subproceso),
    values: proc.subs.map(s => s.pctg),
  }], {
    x: RX+0.15, y: 0.94, w: RW+0.1, h: 2.4,
    barDir: "bar",
    chartColors: proc.subs.map(s => mColor(s.pctg)),
    chartArea: { fill: { color: C.white }, roundedCorners: false },
    catAxisLabelColor: C.grey600, valAxisLabelColor: C.grey600,
    valGridLine: { color: C.grey100, size: 0.5 }, catGridLine: { style: "none" },
    showValue: true, dataLabelPosition: "outEnd",
    dataLabelColor: C.textDeep, dataLabelFontSize: 8,
    showLegend: false, valAxisMaxVal: 100,
  });

  // Gráfico de distribución de niveles de madurez
  secLabel(slide, RX+0.25, 3.42, "DISTRIBUCIÓN POR NIVEL DE MADUREZ (N° PREGUNTAS)");
  const orden = ["1-Inexistente", "2-Inicial", "3-Definida", "4-Optimizada"];
  slide.addChart(pres.charts.BAR, [{
    name:   "Preguntas",
    labels: orden.map(l => l.split("-")[1]),
    values: orden.map(l => proc.maturityCounts[l] || 0),
  }], {
    x: RX+0.15, y: 3.66, w: RW+0.1, h: 1.40,
    barDir: "col",
    chartColors: [C.red, C.orange, C.yellow, C.green],
    chartArea: { fill: { color: C.white }, roundedCorners: false },
    catAxisLabelColor: C.grey600, valAxisLabelColor: C.grey600,
    valGridLine: { color: C.grey100, size: 0.5 }, catGridLine: { style: "none" },
    showValue: true, dataLabelPosition: "outEnd",
    dataLabelColor: C.textDeep, dataLabelFontSize: 9, showLegend: false,
  });

  // ── Panel de análisis Llama expandido ─────────────────────────────────────
  secLabel(slide, RX+0.25, 5.14, "🔍 ANÁLISIS EXPERTO — NIST CSF 2.0 + ISO/IEC 27001:2022");

  if (procAnalysis) {
    const panelY = 5.38;
    const col1X  = RX + 0.15;
    const col1W  = (RW - 0.15) * 0.52;
    const col2X  = col1X + col1W + 0.08;
    const col2W  = (RW - 0.15) * 0.46;
    const panelH = 1.80;

    // Columna 1: por qué es crítico + referencias normativas
    slide.addShape(pres.shapes.RECTANGLE, { x: col1X, y: panelY, w: col1W, h: panelH, fill: { color: C.white }, line: { color: C.grey300, width: 0.5 }, shadow: mkShadow() });
    slide.addShape(pres.shapes.RECTANGLE, { x: col1X, y: panelY, w: 0.05, h: panelH, fill: { color: pc }, line: { color: pc, width: 0 } });

    slide.addText("POR QUÉ ES CRÍTICO", { x: col1X+0.12, y: panelY+0.06, w: col1W-0.18, h: 0.16, fontSize: 6.5, bold: true, color: pc, fontFace: "Calibri", charSpacing: 1, margin: 0 });
    slide.addText(safeStr(procAnalysis.why_critical), { x: col1X+0.12, y: panelY+0.24, w: col1W-0.18, h: 0.64, fontSize: 7.8, color: C.textDark, fontFace: "Calibri", wrap: true, valign: "top", margin: 0, lineSpacingMultiple: 1.1 });

    // Referencias normativas
    const nistRefs = (procAnalysis.normative_references && Array.isArray(procAnalysis.normative_references.nist_csf))
      ? procAnalysis.normative_references.nist_csf.map(safeStr).join("  ·  ") : "";
    const isoRefs  = (procAnalysis.normative_references && Array.isArray(procAnalysis.normative_references.iso_27001))
      ? procAnalysis.normative_references.iso_27001.map(safeStr).join("  ·  ") : "";
    slide.addShape(pres.shapes.RECTANGLE, { x: col1X+0.10, y: panelY+0.92, w: col1W-0.20, h: 0.02, fill: { color: C.grey300 }, line: { color: C.grey300, width: 0 } });
    slide.addText("NIST CSF 2.0:", { x: col1X+0.12, y: panelY+0.98, w: 0.85, h: 0.18, fontSize: 6.5, bold: true, color: C.brandBlue, fontFace: "Calibri", margin: 0 });
    slide.addText(nistRefs, { x: col1X+0.97, y: panelY+0.98, w: col1W-1.10, h: 0.18, fontSize: 6.5, color: C.textDark, fontFace: "Calibri", margin: 0 });
    slide.addText("ISO 27001:", { x: col1X+0.12, y: panelY+1.18, w: 0.85, h: 0.18, fontSize: 6.5, bold: true, color: C.teal, fontFace: "Calibri", margin: 0 });
    slide.addText(isoRefs, { x: col1X+0.97, y: panelY+1.18, w: col1W-1.10, h: 0.18, fontSize: 6.5, color: C.textDark, fontFace: "Calibri", margin: 0 });

    // Métricas de éxito
    const successMetrics = safeStrArray(procAnalysis.success_metrics);
    if (successMetrics.length > 0) {
      slide.addText("MÉTRICAS DE ÉXITO", { x: col1X+0.12, y: panelY+1.40, w: col1W-0.18, h: 0.16, fontSize: 6.5, bold: true, color: C.green, fontFace: "Calibri", charSpacing: 1, margin: 0 });
      const metricsText = successMetrics.map((m, i) => `${i+1}. ${m}`).join("\n");
      slide.addText(metricsText, { x: col1X+0.12, y: panelY+1.58, w: col1W-0.18, h: 0.20, fontSize: 6.5, color: C.textDark, fontFace: "Calibri", wrap: true, valign: "top", margin: 0 });
    }

    // Columna 2: acciones recomendadas
    slide.addShape(pres.shapes.RECTANGLE, { x: col2X, y: panelY, w: col2W, h: panelH, fill: { color: C.white }, line: { color: C.grey300, width: 0.5 }, shadow: mkShadow() });
    slide.addShape(pres.shapes.RECTANGLE, { x: col2X, y: panelY, w: 0.05, h: panelH, fill: { color: C.green }, line: { color: C.green, width: 0 } });

    slide.addText("ACCIONES RECOMENDADAS", { x: col2X+0.12, y: panelY+0.06, w: col2W-0.18, h: 0.16, fontSize: 6.5, bold: true, color: C.green, fontFace: "Calibri", charSpacing: 1, margin: 0 });

    const actions = safeStrArray(procAnalysis.recommended_actions);
    const actText = actions.map((a, i) => `${i+1}. ${a}`).join("\n");
    slide.addText(actText, { x: col2X+0.12, y: panelY+0.26, w: col2W-0.18, h: 1.10, fontSize: 7.4, color: C.textDark, fontFace: "Calibri", wrap: true, valign: "top", margin: 0, lineSpacingMultiple: 1.2 });

    // Cómo implementar
    const howTo = safeStr(procAnalysis.how_to_implement);
    if (howTo) {
      slide.addShape(pres.shapes.RECTANGLE, { x: col2X+0.10, y: panelY+1.38, w: col2W-0.20, h: 0.02, fill: { color: C.grey300 }, line: { color: C.grey300, width: 0 } });
      slide.addText("🛠️ IMPLEMENTACIÓN:", { x: col2X+0.12, y: panelY+1.44, w: col2W-0.18, h: 0.16, fontSize: 6.5, bold: true, color: C.orange, fontFace: "Calibri", margin: 0 });
      slide.addText(howTo, { x: col2X+0.12, y: panelY+1.62, w: col2W-0.18, h: 0.16, fontSize: 6.5, color: C.textDark, fontFace: "Calibri", wrap: true, valign: "top", margin: 0 });
    }

  } else {
    // Fallback: insight generado algorítmicamente
    const insightText = generateInsight(proc);
    slide.addShape(pres.shapes.RECTANGLE, { x: RX+0.15, y: 5.38, w: RW+0.1, h: 1.80, fill: { color: C.surface }, line: { color: C.grey300, width: 0.5 } });
    slide.addShape(pres.shapes.RECTANGLE, { x: RX+0.15, y: 5.38, w: 0.06, h: 1.80, fill: { color: pc }, line: { color: pc, width: 0 } });
    slide.addText(insightText, { x: RX+0.35, y: 5.44, w: RW-0.15, h: 1.68, fontSize: 9.5, color: C.textDark, fontFace: "Calibri", wrap: true, valign: "top", margin: 0 });
  }

  addFooter(slide, pres);
}

// ══════════════════════════════════════════════════════════════════════════════
// SECCIÓN 9: SLIDE — INFORME GENERAL DE CONCLUSIONES
// ══════════════════════════════════════════════════════════════════════════════

function buildConclusions(pres, data, heatmapData, pageNum, total) {
  const { processes, globalAvg } = data;
  const riskAvg = heatmapData.globalAvg;
  const slide = pres.addSlide();
  addHeader(slide, pres, "Informe General — Conclusiones y Hoja de Ruta", pageNum, total);
  addFooter(slide, pres);

  const scoreColor = mColor(globalAvg);
  const nivel = globalAvg < 30 ? "CRÍTICO" : globalAvg < 50 ? "DEFICIENTE" : globalAvg < 70 ? "EN PROGRESO" : "AVANZADO";

  // ── Fila 1: Diagnóstico general + KPIs ───────────────────────────────────
  const diagText = safeStr(conclusions.overall_diagnosis) ||
    `La organización alcanza un índice global de madurez del ${globalAvg.toFixed(1)}% bajo NIST CSF 2.0 e ISO/IEC 27001:2022, en estado ${nivel}. Existen brechas estructurales que requieren atención inmediata.`;

  slide.addShape(pres.shapes.RECTANGLE, { x: 0.35, y: 1.02, w: 5.80, h: 1.80, fill: { color: C.white }, line: { color: C.grey300, width: 0.5 }, shadow: mkShadow() });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.35, y: 1.02, w: 0.06, h: 1.80, fill: { color: scoreColor }, line: { color: scoreColor, width: 0 } });
  slide.addText("DIAGNÓSTICO INTEGRAL — NIST CSF 2.0 + ISO/IEC 27001:2022", { x: 0.50, y: 1.07, w: 5.55, h: 0.18, fontSize: 6.5, bold: true, color: scoreColor, fontFace: "Calibri", charSpacing: 1, margin: 0 });
  slide.addText(diagText, { x: 0.50, y: 1.27, w: 5.55, h: 1.48, fontSize: 8.6, color: C.textDark, fontFace: "Calibri", wrap: true, valign: "top", margin: 0, lineSpacingMultiple: 1.15 });

  // KPI Cards
  const kpis = [
    { label: "Madurez Global",      value: `${globalAvg.toFixed(1)}%`, color: scoreColor,       sub: nivel },
    { label: "Riesgo Inherente",    value: `${riskAvg.toFixed(1)}%`,   color: mColor(riskAvg),  sub: "Índice global" },
    { label: "Procesos Evaluados",  value: `${processes.length}`,      color: C.brandBlue,      sub: "NIST CSF 2.0" },
  ];
  kpis.forEach((k, i) => {
    addKPICard(slide, pres, 6.30 + i * 2.36, 1.02, 2.20, 1.80, k.label, k.value, k.color, k.sub);
  });

  // ── Fila 2: Hallazgos clave ───────────────────────────────────────────────
  secLabel(slide, 0.35, 2.98, "🔎  HALLAZGOS CLAVE DE LA EVALUACIÓN");

  // FIX PRINCIPAL: sanitizar key_findings con safeStrArray para evitar
  // que objetos devueltos por el LLM lleguen a slide.addText()
  const findings = safeStrArray(conclusions.key_findings, [
    `Proceso más maduro: ${processes[processes.length-1].name} (${processes[processes.length-1].avgPct.toFixed(1)}%)`,
    `Proceso más crítico: ${processes[0].name} (${processes[0].avgPct.toFixed(1)}%)`,
    "Patrón transversal: predominio de niveles 1-Inexistente y 2-Inicial",
    `Riesgo sistémico: índice inherente del ${riskAvg.toFixed(1)}% sin controles mitigadores formales`,
  ]);

  slide.addShape(pres.shapes.RECTANGLE, { x: 0.35, y: 3.22, w: 5.80, h: 2.20, fill: { color: C.white }, line: { color: C.grey300, width: 0.5 }, shadow: mkShadow() });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.35, y: 3.22, w: 0.06, h: 2.20, fill: { color: C.accent }, line: { color: C.accent, width: 0 } });
  findings.forEach((f, i) => {
    const fy = 3.30 + i * 0.52;
    slide.addShape(pres.shapes.OVAL, { x: 0.52, y: fy+0.07, w: 0.18, h: 0.18, fill: { color: C.accent }, line: { color: C.accent, width: 0 } });
    slide.addText(`${i+1}`, { x: 0.52, y: fy+0.07, w: 0.18, h: 0.18, fontSize: 7, bold: true, color: C.white, fontFace: "Calibri", align: "center", valign: "middle", margin: 0 });
    // f ya es string garantizado por safeStrArray
    slide.addText(f, { x: 0.78, y: fy+0.04, w: 5.24, h: 0.44, fontSize: 8.4, color: C.textDark, fontFace: "Calibri", wrap: true, valign: "top", margin: 0, lineSpacingMultiple: 1.1 });
  });

  // Gap analysis NIST / ISO
  secLabel(slide, 6.30, 2.98, "📋  BRECHAS DE CUMPLIMIENTO NORMATIVO");
  slide.addShape(pres.shapes.RECTANGLE, { x: 6.30, y: 3.22, w: 6.72, h: 2.20, fill: { color: C.white }, line: { color: C.grey300, width: 0.5 }, shadow: mkShadow() });

  const nistGap = safeStr(conclusions.compliance_gap_summary && conclusions.compliance_gap_summary.nist_csf_gap) ||
    "Las funciones de Gobernar e Identificar presentan las mayores brechas bajo NIST CSF 2.0.";
  const isoGap = safeStr(conclusions.compliance_gap_summary && conclusions.compliance_gap_summary.iso_27001_gap) ||
    "Los dominios de control de activos y gestión de accesos requieren atención inmediata bajo ISO/IEC 27001:2022.";

  // Badge NIST
  slide.addShape(pres.shapes.RECTANGLE, { x: 6.42, y: 3.30, w: 1.80, h: 0.34, fill: { color: C.brandBlue }, line: { color: C.brandBlue, width: 0 } });
  slide.addText("NIST CSF 2.0", { x: 6.42, y: 3.30, w: 1.80, h: 0.34, fontSize: 8.5, bold: true, color: C.white, fontFace: "Calibri", align: "center", valign: "middle", margin: 0 });
  slide.addText(nistGap, { x: 6.42, y: 3.68, w: 6.40, h: 0.72, fontSize: 8.6, color: C.textDark, fontFace: "Calibri", wrap: true, valign: "top", margin: 0, lineSpacingMultiple: 1.1 });

  // Separador
  slide.addShape(pres.shapes.RECTANGLE, { x: 6.42, y: 4.44, w: 6.40, h: 0.02, fill: { color: C.grey300 }, line: { color: C.grey300, width: 0 } });

  // Badge ISO
  slide.addShape(pres.shapes.RECTANGLE, { x: 6.42, y: 4.50, w: 1.80, h: 0.34, fill: { color: C.teal }, line: { color: C.teal, width: 0 } });
  slide.addText("ISO/IEC 27001:2022", { x: 6.42, y: 4.50, w: 1.80, h: 0.34, fontSize: 8.5, bold: true, color: C.white, fontFace: "Calibri", align: "center", valign: "middle", margin: 0 });
  slide.addText(isoGap, { x: 6.42, y: 4.88, w: 6.40, h: 0.50, fontSize: 8.6, color: C.textDark, fontFace: "Calibri", wrap: true, valign: "top", margin: 0, lineSpacingMultiple: 1.1 });

  // ── Fila 3: Recomendaciones estratégicas ──────────────────────────────────
  secLabel(slide, 0.35, 5.58, "🚀  RECOMENDACIONES ESTRATÉGICAS — HOJA DE RUTA");

  // FIX: sanitizar cada campo de cada recomendación
  const rawRecs = Array.isArray(conclusions.strategic_recommendations)
    ? conclusions.strategic_recommendations : [];

  const recs = rawRecs.length > 0 ? rawRecs : [
    { titulo: "Programa de remediación prioritaria", descripcion: "Implementar plan formal con sponsor ejecutivo y KPIs en 90 días. NIST GV.OC-01 + ISO A.5.1.", prioridad: "ALTA", timeline: "0-90 días" },
    { titulo: "Marco de gestión de riesgos continuo", descripcion: "Establecer proceso formal trimestral de evaluación de riesgos. NIST GV.RM-01 + ISO A.5.36.", prioridad: "ALTA", timeline: "90-180 días" },
    { titulo: "Roadmap de madurez a 12 meses", descripcion: `Elevar índice del ${globalAvg.toFixed(1)}% al 60% con hitos trimestrales medibles. NIST ID.IM-01.`, prioridad: "MEDIA", timeline: "180-365 días" },
  ];

  const recW = (13.33 - 0.35 - 0.35 - 0.12 * 2) / 3;
  recs.forEach((rec, i) => {
    const rx = 0.35 + i * (recW + 0.12);
    const ry = 5.82;
    const rh = 1.28;

    // FIX: garantizar que todos los campos sean strings antes de addText
    const recTitulo     = safeStr(rec.titulo);
    const recDescripcion = safeStr(rec.descripcion);
    const recPrioridad  = safeStr(rec.prioridad).toUpperCase() || "MEDIA";
    const recTimeline   = safeStr(rec.timeline);
    const prioColor     = PRIORITY_COLORS[recPrioridad] || C.brandBlue;

    slide.addShape(pres.shapes.RECTANGLE, { x: rx, y: ry, w: recW, h: rh, fill: { color: C.white }, line: { color: C.grey300, width: 0.5 }, shadow: mkShadow() });
    slide.addShape(pres.shapes.RECTANGLE, { x: rx, y: ry, w: recW, h: 0.06, fill: { color: prioColor }, line: { color: prioColor, width: 0 } });

    // Badge de prioridad + timeline
    slide.addShape(pres.shapes.RECTANGLE, { x: rx+0.10, y: ry+0.12, w: 0.68, h: 0.22, fill: { color: prioColor }, line: { color: prioColor, width: 0 } });
    slide.addText(recPrioridad,  { x: rx+0.10, y: ry+0.12, w: 0.68,       h: 0.22, fontSize: 6.5, bold: true, color: C.white,    fontFace: "Calibri", align: "center", valign: "middle", margin: 0 });
    slide.addText(recTimeline,   { x: rx+0.84, y: ry+0.12, w: recW-0.94,  h: 0.22, fontSize: 6.5, color: C.grey600,  fontFace: "Calibri", valign: "middle", margin: 0 });
    slide.addText(recTitulo,     { x: rx+0.10, y: ry+0.38, w: recW-0.20,  h: 0.26, fontSize: 9,   bold: true, color: C.textDeep, fontFace: "Calibri", wrap: true, margin: 0 });
    slide.addText(recDescripcion,{ x: rx+0.10, y: ry+0.66, w: recW-0.20,  h: 0.58, fontSize: 7.6, color: C.textDark, fontFace: "Calibri", wrap: true, valign: "top", margin: 0, lineSpacingMultiple: 1.1 });
  });

  // ── Mensaje de cierre ─────────────────────────────────────────────────────
  const closingMsg = safeStr(conclusions.closing_message);
  if (closingMsg) {
    slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 7.10, w: 13.33, h: 0.08, fill: { color: scoreColor }, line: { color: scoreColor, width: 0 } });
  }
}

// ══════════════════════════════════════════════════════════════════════════════
// SECCIÓN 10: SLIDE — ANÁLISIS HEATMAP (con análisis Llama)
// ══════════════════════════════════════════════════════════════════════════════

function buildHeatmapAnalysisSlide(pres, heatmapData, pageNum, total) {
  if (!heatAnalysis || Object.keys(heatAnalysis).length === 0) return;

  const slide = pres.addSlide();
  addHeader(slide, pres, "Análisis del Mapa de Riesgos — NIST CSF + ISO 27001", pageNum, total);
  addFooter(slide, pres);

  const risks     = heatmapData.processes;
  const globalAvg = heatmapData.globalAvg;
  const critical  = risks.filter(r => r.ri >= 6);
  const high      = risks.filter(r => r.ri >= 4 && r.ri < 6);
  const medium    = risks.filter(r => r.ri < 4);

  // KPI cards de distribución de riesgos
  const riskKPIs = [
    { label: "Riesgos Críticos (RI ≥ 6)", value: `${critical.length}`, color: C.red,           sub: "Acción inmediata" },
    { label: "Riesgos Altos (RI 4-5)",    value: `${high.length}`,     color: C.orange,        sub: "Prioridad alta" },
    { label: "Riesgos Medios/Bajos",      value: `${medium.length}`,   color: C.green,         sub: "Monitoreo continuo" },
    { label: "Madurez Promedio",          value: `${globalAvg.toFixed(1)}%`, color: mColor(globalAvg), sub: "Índice de controles" },
  ];
  riskKPIs.forEach((k, i) => {
    addKPICard(slide, pres, 0.35 + i * 3.20, 1.02, 3.04, 1.10, k.label, k.value, k.color, k.sub);
  });

  // Bloques de análisis Llama en 2 columnas
  const leftX = 0.35, leftW = 6.25, rightX = 6.73, rightW = 6.25;

  addAnalysisBlock(slide, pres, leftX, 2.28, leftW, 1.20,
    safeStr(heatAnalysis.heatmap_reading),
    C.brandBlue, "📊 CÓMO INTERPRETAR EL MAPA DE RIESGOS"
  );
  addAnalysisBlock(slide, pres, rightX, 2.28, rightW, 1.20,
    safeStr(heatAnalysis.critical_zone_analysis),
    C.red, "🚨 ANÁLISIS DE ZONA CRÍTICA (RI ≥ 6)"
  );
  addAnalysisBlock(slide, pres, leftX, 3.56, leftW, 0.90,
    safeStr(heatAnalysis.risk_clusters),
    C.orange, "🔗 PATRONES Y CLUSTERS DE RIESGO"
  );
  addAnalysisBlock(slide, pres, rightX, 3.56, rightW, 0.90,
    safeStr(heatAnalysis.remediation_priority),
    C.teal, "🎯 ESTRATEGIA DE PRIORIZACIÓN"
  );
  addAnalysisBlock(slide, pres, leftX, 4.54, leftW, 0.70,
    safeStr(heatAnalysis.executive_risk_message),
    C.deepBlue, "💼 MENSAJE PARA LA JUNTA DIRECTIVA"
  );

  // Controles NIST/ISO
  const isoControls = safeStrArray(heatAnalysis.iso_nist_controls);
  if (isoControls.length > 0) {
    secLabel(slide, rightX, 4.54, "🛡️  CONTROLES NIST CSF / ISO 27001 PRIORITARIOS");
    slide.addShape(pres.shapes.RECTANGLE, { x: rightX, y: 4.78, w: rightW, h: 0.46, fill: { color: C.white }, line: { color: C.grey300, width: 0.5 }, shadow: mkShadow() });
    slide.addShape(pres.shapes.RECTANGLE, { x: rightX, y: 4.78, w: 0.05, h: 0.46, fill: { color: C.teal }, line: { color: C.teal, width: 0 } });
    const ctrlText = isoControls.map((c, i) => `${i+1}. ${c}`).join("\n");
    slide.addText(ctrlText, { x: rightX+0.14, y: 4.82, w: rightW-0.20, h: 0.38, fontSize: 7.4, color: C.textDark, fontFace: "Calibri", wrap: true, valign: "top", margin: 0, lineSpacingMultiple: 1.1 });
  }

  // Roadmap summary
  const roadmapText = safeStr(conclusions.roadmap_summary);
  if (roadmapText) {
    secLabel(slide, leftX, 5.30, "🗓️  RESUMEN DEL ROADMAP DE MEJORA");
    addAnalysisBlock(slide, pres, leftX, 5.54, 12.98, 0.90, roadmapText, C.green, null);
  }

  // Recomendaciones globales
  const safeGlobalRecs = safeStrArray(globalRecs);
  if (safeGlobalRecs.length > 0) {
    secLabel(slide, leftX, 6.52, "📌  RECOMENDACIONES GLOBALES");
    const recText = safeGlobalRecs.map((r, i) => `${i+1}. ${r}`).join("   |   ");
    slide.addShape(pres.shapes.RECTANGLE, { x: leftX, y: 6.74, w: 12.98, h: 0.36, fill: { color: C.surface }, line: { color: C.grey300, width: 0.5 } });
    slide.addText(recText, { x: leftX+0.10, y: 6.76, w: 12.78, h: 0.32, fontSize: 7.4, color: C.textDark, fontFace: "Calibri", wrap: true, valign: "middle", margin: 0 });
  }
}

// ══════════════════════════════════════════════════════════════════════════════
// SECCIÓN 11: FUNCIÓN PRINCIPAL
// ══════════════════════════════════════════════════════════════════════════════

async function main() {
  if (!heatmapData || !Array.isArray(heatmapData.processes)) {
    console.error("❌ heatmapData inválido o ausente en el payload.");
    process.exit(1);
  }

  const { processes } = data;
  const nistOrdered = orderByNist(processes);

  // Conteo total de slides:
  // 1: portada | 2: leyenda | 3: ejecutivo | 4: dashboard
  // N: procesos (uno por proceso en orden NIST)
  // N+1: conclusiones | N+2: análisis heatmap | N+3: heatmap
  const TOTAL = 4 + nistOrdered.length + 3;

  const pres = new pptxgen();
  pres.layout  = "LAYOUT_WIDE";
  pres.title   = "CMMI Cybersecurity Report — RBKT";
  pres.author  = "RBKT Cybersecurity";
  pres.company = "RBKT";

  let page = 1;

  buildCover(pres, TOTAL);                                      // slide 1
  page++;
  buildLegend(pres, page, TOTAL);                               // slide 2
  page++;
  buildExecutive(pres, data, page, TOTAL);                      // slide 3
  page++;
  buildDashboard(pres, data, page, TOTAL, nistOrdered);         // slide 4
  page++;

  nistOrdered.forEach((p, i) => {
    buildProcess(pres, p, i + 1, page, TOTAL, nistOrdered);
    page++;
  });

  buildConclusions(pres, data, heatmapData, page, TOTAL);       // slide N+1
  page++;
  buildHeatmapAnalysisSlide(pres, heatmapData, page, TOTAL);    // slide N+2
  page++;
  buildHeatMap(pres, heatmapData, page, TOTAL, logoB64);        // slide N+3

  await pres.writeFile({ fileName: outPath });
  console.log("OK:" + outPath);
}

main().catch(e => { console.error(e); process.exit(1); });