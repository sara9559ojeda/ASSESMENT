// ═════════════════════════════════════════════════════════════════════════════
// buildHeatMap — Gaussian Heatmap para PptxGenJS
// heatmap_pptx.js  (v5 — centrado, ejes con Bajo/Medio/Alto, números reposicionados)
// ═════════════════════════════════════════════════════════════════════════════

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
};

const NOMBRES_RIESGOS = {
  1:  "Gestión estratégica",
  2:  "Planificación y seguimiento",
  3:  "Gestión de requisitos",
  4:  "Gestión de proyectos",
  5:  "Aseguramiento de calidad",
  6:  "Gestión de configuración",
  7:  "Medición y análisis",
  8:  "Gestión de proveedores",
  9:  "Acuerdos con proveedores",
  10: "Infraestructura",
  11: "Recursos humanos",
  12: "Gestión del conocimiento",
  13: "Comunicación organizacional",
  14: "Gestión del desempeño",
  15: "Mejora continua",
  16: "Gestión de cambios",
  17: "Análisis de decisiones",
  18: "Riesgos operacionales",
  19: "Gobierno corporativo",
};

const HEAT_T = [0.00, 0.15, 0.30, 0.46, 0.60, 0.75, 0.88, 1.00];
const HEAT_C = [
  [ 15, 100,  15],
  [ 60, 170,  20],
  [150, 210,   0],
  [240, 210,   0],
  [255, 140,   0],
  [230,  55,   0],
  [185,  10,   0],
  [130,   0,   0],
];

function clampInt(v) { return Math.max(0, Math.min(255, Math.round(v))); }
function toHex(v)    { return clampInt(v).toString(16).padStart(2, "0").toUpperCase(); }

function lerpColor(t) {
  t = Math.max(0.0, Math.min(1.0, t));
  let i = 0;
  while (i < HEAT_T.length - 2 && HEAT_T[i + 1] <= t) i++;
  const f = (t - HEAT_T[i]) / (HEAT_T[i + 1] - HEAT_T[i]);
  const r = clampInt(HEAT_C[i][0] + (HEAT_C[i+1][0] - HEAT_C[i][0]) * f);
  const g = clampInt(HEAT_C[i][1] + (HEAT_C[i+1][1] - HEAT_C[i][1]) * f);
  const b = clampInt(HEAT_C[i][2] + (HEAT_C[i+1][2] - HEAT_C[i][2]) * f);
  return toHex(r) + toHex(g) + toHex(b);
}

function brightenHex(hex, factor = 1.25) {
  const r = clampInt(parseInt(hex.slice(0, 2), 16) * factor);
  const g = clampInt(parseInt(hex.slice(2, 4), 16) * factor);
  const b = clampInt(parseInt(hex.slice(4, 6), 16) * factor);
  return toHex(r) + toHex(g) + toHex(b);
}

const mColor = pct =>
  pct < 30 ? C.red : pct < 50 ? C.orange : pct < 70 ? C.yellow : C.green;

const mLabel = pct =>
  pct < 30 ? "Inexistente" : pct < 50 ? "Inicial" : pct < 70 ? "En progreso" : "Optimizada";

// ─────────────────────────────────────────────────────────────────────────────
// Geometría
// Mapa de calor centrado horizontalmente dentro del recuadro contenedor.
// Se deja más margen izquierdo para eje Y + label PROBABILIDAD, y el ancho
// del plot se reduce para que quepa simétricamente.
// ─────────────────────────────────────────────────────────────────────────────
const GEO = {
  boxX:   0.28,
  boxY:   0.58,
  boxW:   8.72,
  boxH:   6.66,
  // plotX aumentado para acomodar label PROBABILIDAD + números + labels Bajo/Medio/Alto
  plotX:  1.60,
  // plotY con más espacio arriba para no pegarse al borde
  plotY:  0.92,
  // plotW reducido para centrar: boxX + plotX_offset + plotW + right_margin ≈ boxX + boxW
  // 0.28 + 1.60 + plotW + 0.20 ≈ 0.28 + 8.72  →  plotW ≈ 7.00
  plotW:  7.00,
  plotH:  5.20,
  tableX: 9.18,
  tableW:  3.96,
  slideW: 13.33,
  slideH:  7.50,
};

const RES = 50;

// Etiquetas de cuadrante para los ejes
const AXIS_LABELS = ["Bajo", "Medio", "Alto"];

function drawDiagonalLine(slide, pres, x1, y1, x2, y2, lineOpts) {
  const bx    = Math.min(x1, x2);
  const by    = Math.min(y1, y2);
  const bw    = Math.abs(x2 - x1);
  const bh    = Math.abs(y2 - y1);
  const flipH = x2 < x1;
  const flipV = y2 < y1;
  if (bw < 0.01 && bh < 0.01) return;
  slide.addShape(pres.shapes.LINE, {
    x: bx, y: by,
    w: Math.max(bw, 0.01),
    h: Math.max(bh, 0.01),
    flipH, flipV,
    line: lineOpts,
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// TABLA LATERAL — sin leyenda de colores al pie
// ─────────────────────────────────────────────────────────────────────────────
function drawRiskTable(slide, pres, allItems) {
  const { tableX, tableW, boxY, boxH } = GEO;
  const panelY = boxY;
  const panelH = boxH;

  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: tableX, y: panelY, w: tableW, h: panelH,
    fill: { color: "0A1828" },
    line: { color: C.brandBlue, width: 0.8 },
    rectRadius: 0.05,
  });

  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: tableX, y: panelY, w: tableW, h: 0.30,
    fill: { color: C.blue800 },
    line: { color: C.brandBlue, width: 0 },
    rectRadius: 0.05,
  });

  slide.addText("REGISTRO DE RIESGOS", {
    x: tableX + 0.10, y: panelY + 0.02, w: tableW - 0.20, h: 0.24,
    fontSize: 7.5, bold: true, color: C.accent,
    fontFace: "Calibri", charSpacing: 2,
    align: "left", valign: "middle", margin: 0,
  });

  slide.addShape(pres.shapes.RECTANGLE, {
    x: tableX + 0.10, y: panelY + 0.28, w: tableW - 0.20, h: 0.005,
    fill: { color: C.brandBlue, transparency: 40 },
    line: { color: C.brandBlue, width: 0 },
  });

  const startY = panelY + 0.32;
  const itemH  = (panelH - 0.36) / 19;

  const colorMap = {};
  allItems.forEach(it => { colorMap[it.label] = it.avgPct; });

  for (let rid = 1; rid <= 19; rid++) {
    const label    = `R${rid}`;
    const name     = NOMBRES_RIESGOS[rid] || "";
    const iy = startY + (rid - 1) * itemH;

    if (rid % 2 === 0) {
      slide.addShape(pres.shapes.RECTANGLE, {
        x: tableX + 0.06, y: iy, w: tableW - 0.12, h: itemH,
        fill: { color: "FFFFFF", transparency: 93 },
        line: { color: "FFFFFF", width: 0 },
      });
    }

    slide.addText(label, {
      x: tableX + 0.10, y: iy, w: 0.34, h: itemH,
      fontSize: 10, bold: true, color: C.accent,
      fontFace: "Calibri", align: "left", valign: "middle", margin: 0,
    });

    slide.addText(name, {
      x: tableX + 0.44, y: iy, w: tableW - 0.54, h: itemH,
      fontSize: 10, color: "DDDDDD",
      fontFace: "Calibri", align: "left", valign: "middle", margin: 0,
    });
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// FUNCIÓN PRINCIPAL
// logoB64 se recibe como parámetro adicional desde 2_POWER_POINT.js
// ─────────────────────────────────────────────────────────────────────────────
function buildHeatMap(pres, heatmapData, pageNum, total, logoB64) {
  if (!heatmapData || !Array.isArray(heatmapData.processes)) {
    throw new Error(
      "buildHeatMap: heatmapData inválido. Recibido: " + JSON.stringify(heatmapData)
    );
  }

  const processes = heatmapData.processes;
  const slide = pres.addSlide();

  // ── Fondo de marca RBKT ───────────────────────────────────────────────────
  slide.background = { color: C.deepBlue };

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: GEO.slideW, h: 0.52,
    fill: { color: C.blue800 },
    line: { color: C.blue800, width: 0 },
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0.50, w: GEO.slideW, h: 0.04,
    fill: { color: C.brandBlue },
    line: { color: C.brandBlue, width: 0 },
  });

  // ── Logo RBKT centrado arriba ─────────────────────────────────────────────
  if (logoB64) {
    slide.addImage({
      data: "image/png;base64," + logoB64,
      x: GEO.slideW / 2 - 1.10,
      y: 0.04,
      w: 2.20,
      h: 0.44,
    });
  } else {
    slide.addText("RBKT CYBERSECURITY", {
      x: 0, y: 0.06, w: GEO.slideW, h: 0.38,
      fontSize: 13, bold: true, color: C.white, fontFace: "Calibri",
      align: "center", valign: "middle", charSpacing: 4, margin: 0,
    });
  }

  // ── Número de página ──────────────────────────────────────────────────────
  slide.addText(`${pageNum} / ${total}`, {
    x: GEO.slideW - 1.40, y: 0.09, w: 1.20, h: 0.28,
    fontSize: 9, color: C.grey600, fontFace: "Calibri",
    align: "right", valign: "middle", margin: 0,
  });

  // ── Recuadro contenedor ───────────────────────────────────────────────────
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: GEO.boxX, y: GEO.boxY, w: GEO.boxW, h: GEO.boxH,
    fill: { color: "091827" },
    line: { color: C.brandBlue, width: 0.4, transparency: 55 },
    rectRadius: 0.05,
  });

  const { plotX, plotY, plotW, plotH } = GEO;

  // ── Gradiente diagonal ────────────────────────────────────────────────────
  const step  = 100.0 / RES;
  const cellW = plotW / RES;
  const cellH = plotH / RES;

  for (let i = 0; i < RES; i++) {
    for (let j = 0; j < RES; j++) {
      const t     = (i * step + j * step) / 200.0;
      const color = lerpColor(t);
      slide.addShape(pres.shapes.RECTANGLE, {
        x: plotX + i * cellW,
        y: plotY + plotH - (j + 1) * cellH,
        w: cellW + 0.002,
        h: cellH + 0.002,
        fill: { color },
        line: { color, width: 0 },
      });
    }
  }

  // ── Grid 3×3 ──────────────────────────────────────────────────────────────
  for (const frac of [1/3, 2/3]) {
    slide.addShape(pres.shapes.RECTANGLE, {
      x: plotX + frac * plotW - 0.005, y: plotY,
      w: 0.010, h: plotH,
      fill: { color: "000000", transparency: 25 },
      line: { color: "000000", width: 0 },
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x: plotX, y: plotY + frac * plotH - 0.005,
      w: plotW, h: 0.010,
      fill: { color: "000000", transparency: 25 },
      line: { color: "000000", width: 0 },
    });
  }

  // ── Borde del plot ────────────────────────────────────────────────────────
  slide.addShape(pres.shapes.RECTANGLE, {
    x: plotX, y: plotY, w: plotW, h: plotH,
    fill: { type: "none" },
    line: { color: "FFFFFF", width: 0.6, transparency: 70 },
  });

  // ── Líneas de zona de riesgo ──────────────────────────────────────────────
  drawDiagonalLine(slide, pres,
    plotX + (1/3) * plotW, plotY + plotH,
    plotX + plotW,          plotY + (2/3) * plotH,
    { color: "FFFFFF", width: 0.5, dashType: "dash", transparency: 75 }
  );
  drawDiagonalLine(slide, pres,
    plotX + (2/3) * plotW, plotY + plotH,
    plotX,                  plotY + (1/3) * plotH,
    { color: "FFFFFF", width: 0.5, dashType: "dash", transparency: 75 }
  );

  slide.addText("LOW RISK", {
    x: plotX + 0.06, y: plotY + plotH - 0.26, w: 1.1, h: 0.22,
    fontSize: 7.5, bold: true, color: "3CC83C", fontFace: "Calibri",
    align: "left", valign: "bottom", margin: 0, transparency: 40,
  });
  slide.addText("CRITICAL RISK", {
    x: plotX + plotW - 1.20, y: plotY + 0.04, w: 1.18, h: 0.22,
    fontSize: 7.5, bold: true, color: "DC3232", fontFace: "Calibri",
    align: "right", valign: "top", margin: 0, transparency: 30,
  });

  // ── Ejes: label texto (Bajo/Medio/Alto) + número centrados en cada tercio ─
  //
  // EJE X (Impacto) — label encima, número abajo
  //   labelY  = plotY + plotH + 0.06   → primera línea: texto Bajo/Medio/Alto
  //   numY    = plotY + plotH + 0.24   → segunda línea: número 1/2/3
  //
  // EJE Y (Probabilidad) — label a la izquierda del número, ambos al lado del eje
  //   El número va más a la izquierda que antes; el label va aún más a la izquierda.

  const axisLabelFontSize = 7.5;
  const axisNumFontSize   = 11;

  [0, 1, 2].forEach(i => {
    // ── Eje X ───────────────────────────────────────────────────────────────
    const cx = plotX + (i + 0.5) * (plotW / 3);

    // Label Bajo/Medio/Alto — encima del número
    slide.addText(AXIS_LABELS[i], {
      x: cx - 0.30, y: plotY + plotH + 0.06, w: 0.60, h: 0.18,
      fontSize: axisLabelFontSize, bold: false, color: "6688AA",
      fontFace: "Calibri", align: "center", valign: "top", margin: 0,
    });

    // Número 1/2/3 — debajo del label
    slide.addText(String(i + 1), {
      x: cx - 0.22, y: plotY + plotH + 0.22, w: 0.44, h: 0.22,
      fontSize: axisNumFontSize, bold: true, color: "8899AA", fontFace: "Calibri",
      align: "center", valign: "top", margin: 0,
    });

    // ── Eje Y ───────────────────────────────────────────────────────────────
    const cy = plotY + plotH - (i + 0.5) * (plotH / 3);

    // Label Bajo/Medio/Alto — encima del número (en vertical = a la derecha del número
    // en pantalla, pero aquí los ponemos en línea apilada):
    // Usamos dos columnas de texto a la izquierda del eje:
    //   columna derecha (más cerca del eje): número
    //   columna izquierda (más lejos del eje): label texto

    // Label texto — más a la izquierda
    slide.addText(AXIS_LABELS[i], {
      x: plotX - 1.10, y: cy - 0.11, w: 0.56, h: 0.22,
      fontSize: axisLabelFontSize, bold: false, color: "6688AA",
      fontFace: "Calibri", align: "right", valign: "middle", margin: 0,
    });

    // Número — justo al lado del eje, desplazado más a la izquierda que antes
    slide.addText(String(i + 1), {
      x: plotX - 0.52, y: cy - 0.13, w: 0.36, h: 0.26,
      fontSize: axisNumFontSize, bold: true, color: "8899AA", fontFace: "Calibri",
      align: "right", valign: "middle", margin: 0,
    });
  });

  // ── Label IMPACTO ─────────────────────────────────────────────────────────
  // Más grande y más cercano al eje para mejor visibilidad
  slide.addText("IMPACTO  →", {
    x: plotX, y: plotY + plotH + 0.44, w: plotW, h: 0.28,
    fontSize: 13, color: "9AAABB", fontFace: "Calibri",
    align: "center", valign: "top", margin: 0, bold: true, charSpacing: 2,
  });

  // ── Label PROBABILIDAD — bounding box ajustado al texto ──────────────────
  // Con rotate:270 el bounding box se dibuja horizontal pero PptxGenJS lo rota
  // alrededor de su centro. Para que quede centrado en el eje Y del plot:
  //   centro visual deseado: cx = GEO.boxX + 0.20, cy = plotY + plotH/2
  // El bounding box tiene:
  //   w = ancho del texto legible (≈ 1.90 in para "PROBABILIDAD  →" a 13pt)
  //   h = alto visible del texto (≈ 0.28 in)
  // Posición del bounding box (top-left antes de rotar):
  //   x = cx - w/2 = GEO.boxX + 0.20 - 0.95
  //   y = cy - h/2 = plotY + plotH/2 - 0.14
  const probW = 1.90;
  const probH = 0.28;
  const probCX = GEO.boxX + 0.20;
  const probCY = plotY + plotH / 2;
  slide.addText("PROBABILIDAD  →", {
    x: probCX - probW / 2,
    y: probCY - probH / 2,
    w: probW,
    h: probH,
    fontSize: 13, color: "9AAABB", fontFace: "Calibri",
    align: "center", valign: "middle", margin: 0, bold: true, charSpacing: 2,
    rotate: 270,
  });

  // ── Agrupar por celda 3×3 ─────────────────────────────────────────────────
  const celdas   = new Map();
  const allItems = [];

  processes.forEach((proc, idx) => {
    if (!proc || typeof proc !== "object") return;
    const xPct   = typeof proc.xPct   === "number" ? proc.xPct   : 50;
    const yPct   = typeof proc.yPct   === "number" ? proc.yPct   : 50;
    const avgPct = typeof proc.avgPct === "number" ? proc.avgPct : 0;
    const label  = proc.label || `R${idx + 1}`;

    allItems.push({ label, avgPct });

    const col = xPct >= 67 ? 2 : xPct >= 34 ? 1 : 0;
    const row = yPct >= 67 ? 2 : yPct >= 34 ? 1 : 0;
    const key = `${col}_${row}`;

    if (!celdas.has(key)) celdas.set(key, { col, row, items: [] });
    celdas.get(key).items.push({ label, avgPct });
  });

  const PILL_W = 0.46;
  const PILL_H = 0.26;

  // ── Renderizar nodos y pills en abanico ───────────────────────────────────
  celdas.forEach(({ col, row, items }) => {
    const n = items.length;
    if (n === 0) return;

    // Centro exacto de la celda
    const cellCX = plotX + (col + 0.5) * (plotW / 3);
    const cellCY = plotY + plotH - (row + 0.5) * (plotH / 3);

    const avgGroup   = items.reduce((s, it) => s + it.avgPct, 0) / n;
    const groupColor = mColor(avgGroup);

    // Anillo
    const ringR = 0.20 + (n - 1) * 0.022;
    slide.addShape(pres.shapes.OVAL, {
      x: cellCX - ringR, y: cellCY - ringR,
      w: ringR * 2, h: ringR * 2,
      fill: { color: groupColor, transparency: 72 },
      line: { color: groupColor, width: 1.2 },
    });

    // Sombra del dot
    const dotR = 0.13;
    slide.addShape(pres.shapes.OVAL, {
      x: cellCX - dotR / 2 + 0.005, y: cellCY - dotR / 2 + 0.005,
      w: dotR, h: dotR,
      fill: { color: "000000", transparency: 45 },
      line: { color: "000000", width: 0 },
    });
    // Dot central
    slide.addShape(pres.shapes.OVAL, {
      x: cellCX - dotR / 2, y: cellCY - dotR / 2,
      w: dotR, h: dotR,
      fill: { color: groupColor },
      line: { color: "FFFFFF", width: 1.4 },
    });

    // Badge de conteo
    if (n > 1) {
      const bR = 0.12;
      slide.addShape(pres.shapes.OVAL, {
        x: cellCX + ringR * 0.55, y: cellCY - ringR * 1.10,
        w: bR * 2, h: bR * 2,
        fill: { color: "0D2B55" },
        line: { color: groupColor, width: 1.0 },
      });
      slide.addText(String(n), {
        x: cellCX + ringR * 0.55, y: cellCY - ringR * 1.10,
        w: bR * 2, h: bR * 2,
        fontSize: 8, bold: true, color: "EEEEEE",
        fontFace: "Calibri", align: "center", valign: "middle", margin: 0,
      });
    }

    // Pills distribuidas en círculo completo (360°) empezando desde arriba (-π/2)
    const FAN_SHORT = ringR + 0.26;
    const FAN_LONG  = ringR + 0.46;

    items.forEach((item, i) => {
      const angle = -Math.PI / 2 + (2 * Math.PI * i) / n;
      const fanR = (i % 2 === 0) ? FAN_LONG : FAN_SHORT;

      const anchorX = cellCX + Math.cos(angle) * (ringR + 0.03);
      const anchorY = cellCY + Math.sin(angle) * (ringR + 0.03);

      const tipX = cellCX + Math.cos(angle) * fanR;
      const tipY = cellCY + Math.sin(angle) * fanR;

      let lx = tipX - PILL_W / 2;
      let ly = tipY - PILL_H / 2;
      lx = Math.max(plotX + 0.03, Math.min(lx, plotX + plotW - PILL_W - 0.03));
      ly = Math.max(plotY + 0.03, Math.min(ly, plotY + plotH - PILL_H - 0.03));

      const pillCX    = lx + PILL_W / 2;
      const pillCY    = ly + PILL_H / 2;
      const pillColor = mColor(item.avgPct);

      drawDiagonalLine(slide, pres,
        anchorX, anchorY,
        pillCX, pillCY,
        { color: "FFFFFF", width: 0.7, transparency: 35 }
      );

      slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: lx, y: ly, w: PILL_W, h: PILL_H,
        fill: { color: "060E1A" },
        line: { color: pillColor, width: 1.1 },
        rectRadius: 0.05,
        shadow: { type: "outer", blur: 6, offset: 2, angle: 135, color: "000000", opacity: 0.60 },
      });

      slide.addText(item.label, {
        x: lx, y: ly, w: PILL_W, h: PILL_H,
        fontSize: 10, bold: true, color: "FFFFFF",
        fontFace: "Calibri", align: "center", valign: "middle", margin: 0,
      });
    });
  });

  // ── Tabla lateral ─────────────────────────────────────────────────────────
  drawRiskTable(slide, pres, allItems);

  // ── Footer ────────────────────────────────────────────────────────────────
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 7.16, w: GEO.slideW, h: 0.34,
    fill: { color: C.blue800 },
    line: { color: C.brandBlue, width: 0.5 },
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 7.14, w: GEO.slideW, h: 0.03,
    fill: { color: C.brandBlue },
    line: { color: C.brandBlue, width: 0 },
  });
  slide.addText("RBKT · Dashboard de Ciberseguridad · Confidencial · 2025", {
    x: 0.4, y: 7.19, w: 10, h: 0.26,
    fontSize: 7.5, color: C.grey600, fontFace: "Calibri",
    valign: "middle", margin: 0,
  });
}

module.exports = { buildHeatMap, lerpColor, mColor, mLabel, brightenHex };