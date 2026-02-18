/* Respuesta | Validador + Generador Word (cliente, sin backend)
 * Validador reforzado: INFORME / RUTA / DIMENSION / VALIDACIÓN
 * + cruce DAEI + comparativo histórico 2025.
 */

// ------------------ PDF.js worker ------------------
if (window.pdfjsLib) {
  pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
} else {
  console.warn("PDF.js no cargó; extracción automática del PDF deshabilitada.");
}

// ------------------ Estado ------------------
const state = {
  memoPdf: null,
  modelXlsx: null,
  refXlsx: null,
  histXlsx: null,
  cuadrosXlsx: null,
  tplDocx: null,

  memoText: "",
  memoFields: {
    memo_nro: "",
    memo_fecha: "",
    para: "",
    de: "",
    asunto: "",
    anio_lectivo: "",
    firma: "",
  },

  validation: null,
  modelData: null,
  compareData: null,
};

const $ = (id) => document.getElementById(id);

// ------------------ Utilitarios UI ------------------
function setMeta(el, file) {
  if (!el) return;
  if (!file) {
    el.classList.remove("ok");
    el.textContent = "Sin archivo";
    return;
  }
  el.classList.add("ok");
  el.textContent = `${file.name} • ${(file.size / 1024).toFixed(1)} KB`;
}

function safeBindDrop({ boxId, inputId, metaId, accept, onFile }) {
  const box = $(boxId);
  const input = $(inputId);
  const meta = $(metaId);

  if (!box || !input || !meta) {
    console.warn("No se pudo enlazar dropbox:", { boxId, inputId, metaId });
    return;
  }

  const handle = async (file) => {
    if (!file) return;
    // accept es informativo; el input ya filtra.
    setMeta(meta, file);
    try {
      await onFile(file);
    } catch (e) {
      console.error("Error procesando archivo", file?.name, e);
      showRuntimeNote(`Error leyendo ${file?.name}: ${e?.message || e}`);
    }
  };

  input.addEventListener("change", (e) => handle(e.target.files?.[0] || null));

  box.addEventListener("dragover", (e) => {
    e.preventDefault();
    box.classList.add("dragover");
  });
  box.addEventListener("dragleave", () => box.classList.remove("dragover"));
  box.addEventListener("drop", async (e) => {
    e.preventDefault();
    box.classList.remove("dragover");
    const file = e.dataTransfer.files?.[0];
    await handle(file);
  });
}

function showRuntimeNote(msg) {
  const el = $("runtimeNote");
  if (!el) return;
  el.style.display = "block";
  el.textContent = msg;
}

function setPill(type, text) {
  const p = $("statusPill");
  if (!p) return;
  p.className = `pill ${type}`;
  p.textContent = text;
}

// ------------------ Lecturas base ------------------
async function readAsArrayBuffer(file) {
  return await file.arrayBuffer();
}

function norm(v) {
  return (v ?? "").toString().trim();
}

function normUpper(v) {
  return norm(v).toUpperCase();
}

function isFilled(v) {
  if (v === null || v === undefined) return false;
  if (typeof v === "string") return v.trim() !== "";
  return true;
}

function getCell(sheet, a1) {
  const cell = sheet?.[a1];
  return cell ? cell.v : null;
}

function countFilledCells(sheet) {
  let count = 0;
  for (const k of Object.keys(sheet || {})) {
    if (k.startsWith("!")) continue;
    const v = sheet[k]?.v;
    if (isFilled(v)) count++;
  }
  return count;
}

function findSheetByName(wb, expected) {
  const e = normUpper(expected);
  for (const n of wb.SheetNames) {
    if (normUpper(n) === e) return n;
  }
  // fallback: contiene
  for (const n of wb.SheetNames) {
    if (normUpper(n).includes(e)) return n;
  }
  return null;
}

function toNumber(v) {
  if (typeof v === "number" && Number.isFinite(v)) return v;
  const s = norm(v).replace(/\./g, "").replace(/,/g, "."); // tolera separadores
  const n = parseFloat(s);
  return Number.isFinite(n) ? n : null;
}

function excelTimeToMinutes(v) {
  // Excel guarda horas como fracción del día.
  if (typeof v === "number" && Number.isFinite(v)) {
    const minutes = v * 24 * 60;
    return minutes;
  }
  const s = norm(v);
  if (!s) return null;
  // hh:mm o hh:mm:ss
  const m = s.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (!m) return null;
  const hh = parseInt(m[1], 10);
  const mm = parseInt(m[2], 10);
  const ss = m[3] ? parseInt(m[3], 10) : 0;
  if (!Number.isFinite(hh) || !Number.isFinite(mm) || !Number.isFinite(ss)) return null;
  return hh * 60 + mm + ss / 60;
}

function hasND(v) {
  const s = normUpper(v);
  return s.includes("#N/D") || s.includes("N/D") || s.includes("#N/A") || s.includes("N/A");
}

// ------------------ PDF parsing (básico) ------------------
async function extractPdfText(file) {
  if (!window.pdfjsLib) return "";
  const buf = await readAsArrayBuffer(file);
  const loadingTask = pdfjsLib.getDocument({ data: buf });
  const pdf = await loadingTask.promise;
  let text = "";
  const maxPages = Math.min(pdf.numPages, 3); // suficiente para encabezado
  for (let p = 1; p <= maxPages; p++) {
    const page = await pdf.getPage(p);
    const content = await page.getTextContent();
    const pageText = content.items.map((it) => it.str).join(" ");
    text += "\n" + pageText;
  }
  return text;
}

function guessMemoFields(text) {
  const t = text.replace(/\s+/g, " ").trim();
  const out = { memo_nro: "", memo_fecha: "", para: "", de: "", asunto: "" };

  // Memorando Nro.
  const mNro = t.match(/Memorando\s+Nro\.?\s*([A-Z0-9\-]+\-M)/i);
  if (mNro) out.memo_nro = mNro[1].trim();

  // Fecha (muy variable; dejamos lo que sigue a "fecha")
  const mFecha = t.match(/de\s+fecha\s+([0-9]{1,2}\s+de\s+[A-Za-zÁÉÍÓÚáéíóúñÑ]+\s+de\s+20\d{2})/);
  if (mFecha) out.memo_fecha = mFecha[1].trim();

  // Asunto
  const mAsunto = t.match(/Asunto\s*:\s*([^\n]{10,140})/i);
  if (mAsunto) out.asunto = mAsunto[1].trim();

  return out;
}

// ------------------ Excel: imágenes (mapa) ------------------
async function listMediaImages(excelFile) {
  if (!window.JSZip) return null;
  const buf = await readAsArrayBuffer(excelFile);
  const zip = await JSZip.loadAsync(buf);
  const media = Object.keys(zip.files).filter((p) => p.startsWith("xl/media/") && !zip.files[p].dir);
  return media;
}

// ------------------ Parseo RUTA ------------------

function parseRutaSheet(shRuta) {
  const distrito = normUpper(getCell(shRuta, "B3"));
  const amieEje = normUpper(getCell(shRuta, "B4"));
  const coordX = getCell(shRuta, "H6");
  const coordY = getCell(shRuta, "H7");

  const matrix = XLSX.utils.sheet_to_json(shRuta, { header: 1, defval: "" });

  // encontrar encabezado de la tabla de rutas (buscamos "RUTA" + "BENEFICIARIOS POR RUTA" + "AMIE IE FUSIONADA")
  const headerIdx = findHeaderRow(matrix, ["RUTA", "BENEFICIARIOS", "AMIE"]);
  const startRow = headerIdx >= 0 ? headerIdx + 1 : 9; // fallback: fila 10 (0-index 9)

  const header = headerIdx >= 0 ? matrix[headerIdx].map((v)=>normUpper(v)) : [];
  const idxRuta = headerIdx >= 0 ? header.findIndex((x)=> x==="RUTA") : 0;
  const idxAmieAfc = headerIdx >= 0 ? header.findIndex((x)=> x.includes("AMIE") && x.includes("FUSION")) : 2;
  const idxTotal = headerIdx >= 0 ? header.findIndex((x)=> x.includes("BENEFICIARIOS") && x.includes("RUTA")) : 3;
  const idxIni = headerIdx >= 0 ? header.findIndex((x)=> x.includes("INICIAL")) : 5;
  const idxEgb = headerIdx >= 0 ? header.findIndex((x)=> x.includes("EGB")) : 6;
  const idxBach = headerIdx >= 0 ? header.findIndex((x)=> x.includes("BACH")) : 7;
  const idxDist = headerIdx >= 0 ? header.findIndex((x)=> x.includes("DISTANCIA") && x.includes("LINEAL")) : header.findIndex((x)=> x.includes("DISTANCIA") && x.includes("LINEAL"));
  const idxX = headerIdx >= 0 ? header.findIndex((x)=> x.includes("COORDENADA") && x.endsWith("X")) : -1;
  const idxY = headerIdx >= 0 ? header.findIndex((x)=> x.includes("COORDENADA") && x.endsWith("Y")) : -1;

  const rows = [];
  const afcSet = new Set();
  let totalBenef = 0;
  const coordIssues = [];
  const distIssues = [];

  for (let r = startRow; r < matrix.length; r++) {
    const row = matrix[r];
    const rutaVal = idxRuta >=0 ? row[idxRuta] : row[0];
    if (String(rutaVal).trim() === "") break;

    const ruta = toNumber(rutaVal) ?? String(rutaVal).trim();
    const afc = idxAmieAfc >=0 ? normUpper(row[idxAmieAfc]) : "";
    const total = idxTotal >=0 ? toNumber(row[idxTotal]) : null;
    const ini = idxIni >=0 ? toNumber(row[idxIni]) : null;
    const egb = idxEgb >=0 ? toNumber(row[idxEgb]) : null;
    const bach = idxBach >=0 ? toNumber(row[idxBach]) : null;

    const x = idxX >=0 ? toNumber(row[idxX]) : null;
    const y = idxY >=0 ? toNumber(row[idxY]) : null;

    const dist = idxDist >=0 ? toNumber(row[idxDist]) : null;

    if (afc) afcSet.add(afc);
    if (total !== null) totalBenef += total;

    if (idxX >=0 || idxY >=0) {
      if (x === null || y === null) coordIssues.push({ fila: r + 1, ruta, afc, x: row[idxX], y: row[idxY] });
    }

    if (idxDist >=0 && dist === null) distIssues.push({ fila: r + 1, ruta, afc });

    rows.push({ fila: r + 1, ruta, afc, total, ini, egb, bach, x, y, dist });
  }

  return {
    distrito,
    amieEje,
    coordX,
    coordY,
    rows,
    afcList: Array.from(afcSet),
    totalBenef,
    coordIssues,
    distIssues,
  };
}

// ------------------ Parseo DIMENSION ------------------
function findHeaderRow(matrix, requiredTokens) {
  const tokens = requiredTokens.map((t) => normUpper(t));
  for (let i = 0; i < Math.min(matrix.length, 120); i++) {
    const row = (matrix[i] || []).map((v) => normUpper(v));
    const ok = tokens.every((tk) => row.some((c) => c.includes(tk)));
    if (ok) return i;
  }
  return -1;
}

function parseDimensionSheet(shDim) {
  const matrix = XLSX.utils.sheet_to_json(shDim, { header: 1, defval: "" });

  // Header principal de la tabla: debe contener RUTA y PRESUPUESTO ANUAL, y columnas de horas.
  const headerIdx = findHeaderRow(matrix, ["RUTA", "PRESUPUESTO ANUAL"]);
  if (headerIdx < 0) {
    return { error: "No se encontró encabezado de dimensionamiento (no se detecta 'RUTA' y 'PRESUPUESTO ANUAL')." };
  }

  const header = matrix[headerIdx].map((v) => normUpper(v));
  const col = (token) => header.findIndex((h) => h.includes(normUpper(token)));

  const idxRuta = col("RUTA");
  const idxPresM = col("PRESUPUESTO MENSUAL");
  const idxPresA = col("PRESUPUESTO ANUAL");
  const idxHs1 = col("HORA DE SALIDA 1");
  const idxHl1 = col("HORA DE LLEGADA 1");
  const idxHs2 = col("HORA DE SALIDA 2");
  const idxHl2 = col("HORA DE LLEGADA 2");
  const idxTi = col("TIEMPO INICIO");
  const idxTr = col("TIEMPO RETORNO");
  const idxPorc = col("OPTIMIZACI");

  const idxCostoUnit = col("COSTO UNITARIO");
  const idxCostoEst = col("COSTO ESTUDIANT");

  const rows = [];
  for (let i = headerIdx + 1; i < Math.min(matrix.length, headerIdx + 2000); i++) {
    const row = matrix[i] || [];
    const ruta = row[idxRuta];
    if (!isFilled(ruta)) continue;
    // saltar filas tipo "NO APLICA" si aparece un bloque, pero lo dejamos como fila también.

    const presM = idxPresM >= 0 ? toNumber(row[idxPresM]) : null;
    const presA = idxPresA >= 0 ? toNumber(row[idxPresA]) : null;

    const hs1 = idxHs1 >= 0 ? row[idxHs1] : null;
    const hl1 = idxHl1 >= 0 ? row[idxHl1] : null;
    const hs2 = idxHs2 >= 0 ? row[idxHs2] : null;
    const hl2 = idxHl2 >= 0 ? row[idxHl2] : null;

    const ti = idxTi >= 0 ? row[idxTi] : null;
    const tr = idxTr >= 0 ? row[idxTr] : null;

    const porc = idxPorc >= 0 ? row[idxPorc] : null;

    const costoUnit = idxCostoUnit >= 0 ? toNumber(row[idxCostoUnit]) : null;
    const costoEst = idxCostoEst >= 0 ? toNumber(row[idxCostoEst]) : null;

    rows.push({
      i,
      ruta,
      presM,
      presA,
      hs1,
      hl1,
      hs2,
      hl2,
      ti,
      tr,
      porc,
      costoUnit,
      costoEst,
    });
  }

  // Validaciones
  const timeIssues = [];
  for (const r of rows) {
    // Solo evaluar filas que parezcan rutas numéricas
    const rutaNum = toNumber(r.ruta);
    if (!rutaNum) continue;

    const t_hs1 = excelTimeToMinutes(r.hs1);
    const t_hl1 = excelTimeToMinutes(r.hl1);
    const t_hs2 = excelTimeToMinutes(r.hs2);
    const t_hl2 = excelTimeToMinutes(r.hl2);

    if ([t_hs1, t_hl1, t_hs2, t_hl2].some((x) => x === null)) {
      timeIssues.push({ ruta: r.ruta, motivo: "Faltan horas (HS1/HL1/HS2/HL2)" });
      continue;
    }

    if (!(t_hs1 < t_hl1 && t_hl1 < t_hs2 && t_hs2 < t_hl2)) {
      timeIssues.push({ ruta: r.ruta, motivo: "Orden de horas incoherente" });
    }

    // TI/TR deben ser parseables y >0
    const tiMin = excelTimeToMinutes(r.ti);
    const trMin = excelTimeToMinutes(r.tr);
    if (tiMin === null || trMin === null || tiMin <= 0 || trMin <= 0) {
      timeIssues.push({ ruta: r.ruta, motivo: "TI/TR faltante o inválido" });
    }

    // Presupuesto anual vs mensual (heurística: anual ≈ mensual * 10)
    if (r.presM !== null && r.presA !== null) {
      const ratio = r.presA / r.presM;
      if (!(ratio > 9.5 && ratio < 10.5)) {
        timeIssues.push({ ruta: r.ruta, motivo: `Pres_A no consistente con Pres_M (ratio ${ratio.toFixed(2)})` });
      }
    } else {
      timeIssues.push({ ruta: r.ruta, motivo: "Presupuesto mensual/anual faltante" });
    }

    // Optimización 0-120% (heurística)
    const p = toNumber(r.porc);
    if (p !== null && (p < 0 || p > 120)) {
      timeIssues.push({ ruta: r.ruta, motivo: `Porcentaje optimización fuera de rango: ${p}` });
    }

    // Costos numéricos (si existen)
    if (r.costoUnit === null) {
      timeIssues.push({ ruta: r.ruta, motivo: "Costo unitario faltante/no numérico" });
    }
    if (r.costoEst === null) {
      // costo por estudiante puede estar vacío en algunos formatos, pero lo marcamos como WARN a nivel de checklist.
    }
  }

  const presATotal = rows.reduce((s, r) => s + (r.presA || 0), 0);

  return {
    headerIdx,
    rows,
    timeIssues,
    presATotal,
  };
}

// ------------------ Parseo VALIDACIÓN ------------------
function parseValidacionSheet(shVal) {
  const matrix = XLSX.utils.sheet_to_json(shVal, { header: 1, defval: "" });
  const headerIdx = findHeaderRow(matrix, ["RUTA", "TEST IDA", "TEST RETORNO"]);
  if (headerIdx < 0) {
    return { error: "No se encontró encabezado con 'RUTA' + 'Test Ida' + 'Test Retorno'." };
  }
  const header = matrix[headerIdx].map((v) => normUpper(v));
  const col = (token) => header.findIndex((h) => h.includes(normUpper(token)));

  const idxRuta = col("RUTA");
  const idxHs = col("HORA DE SALIDA");
  const idxHl = col("HORA DE LLEGAD");
  const idxTestIda = col("TEST IDA");
  const idxTestRet = col("TEST RETORNO");

  const rows = [];
  for (let i = headerIdx + 1; i < Math.min(matrix.length, headerIdx + 5000); i++) {
    const row = matrix[i] || [];
    const ruta = row[idxRuta];
    if (!isFilled(ruta)) continue;

    const hs = idxHs >= 0 ? row[idxHs] : null;
    const hl = idxHl >= 0 ? row[idxHl] : null;

    const testIda = idxTestIda >= 0 ? normUpper(row[idxTestIda]) : "";
    const testRet = idxTestRet >= 0 ? normUpper(row[idxTestRet]) : "";

    rows.push({ i, ruta, hs, hl, testIda, testRet });
  }

  const notTrue = rows.filter((r) => {
    const idaOk = r.testIda === "VERDADERO" || r.testIda === "TRUE" || r.testIda === "SI" || r.testIda === "SÍ";
    const retOk = r.testRet === "VERDADERO" || r.testRet === "TRUE" || r.testRet === "SI" || r.testRet === "SÍ";
    return !(idaOk && retOk);
  });

  const timeBad = rows.filter((r) => {
    const tHs = excelTimeToMinutes(r.hs);
    const tHl = excelTimeToMinutes(r.hl);
    if (tHs === null || tHl === null) return true;
    return !(tHs < tHl);
  });

  return {
    headerIdx,
    rows,
    notTrueCount: notTrue.length,
    timeBadCount: timeBad.length,
  };
}

// ------------------ DAEI ------------------
function parseDAEIRanges(refWb) {
  const sheetName = refWb.SheetNames[0];
  const sh = refWb.Sheets[sheetName];
  const matrix = XLSX.utils.sheet_to_json(sh, { header: 1, defval: null });
  let headerRow = -1;
  let idxAmie = -1, idxMin = -1, idxMax = -1, idxEst = -1;

  for (let i = 0; i < Math.min(matrix.length, 50); i++) {
    const row = matrix[i] || [];
    const rowNorm = row.map((c) => normUpper(c));
    const a = rowNorm.findIndex((c) => c === "AMIE" || c.includes("AMIE"));
    const mi = rowNorm.findIndex((c) => c.includes("MINIMO_F") || c.includes("MÍNIMO_F"));
    const ma = rowNorm.findIndex((c) => c.includes("MAXIMO_F") || c.includes("MÁXIMO_F"));
    const es = rowNorm.findIndex((c) => c.includes("ESTIMADO") && c.includes("2026"));
    if (a >= 0 && mi >= 0 && ma >= 0) {
      headerRow = i;
      idxAmie = a;
      idxMin = mi;
      idxMax = ma;
      idxEst = es;
      break;
    }
  }

  if (headerRow < 0) {
    return { error: "No se detecta encabezado con AMIE + Mínimo_F + Máximo_F." };
  }

  const ranges = {};
  for (let i = headerRow + 1; i < matrix.length; i++) {
    const row = matrix[i] || [];
    const amie = normUpper(row[idxAmie]);
    if (!amie) continue;
    const min = toNumber(row[idxMin]);
    const max = toNumber(row[idxMax]);
    const est = idxEst >= 0 ? toNumber(row[idxEst]) : null;
    if (min === null || max === null) continue;
    ranges[amie] = { min, max, est };
  }

  return { ranges };
}

// ------------------ Histórico ------------------

function parseHistorico(wbHist, amieEje) {
  const sheetName = wbHist.SheetNames[0];
  const sh = wbHist.Sheets[sheetName];
  const matrix = XLSX.utils.sheet_to_json(sh, { header: 1, defval: "" });

  // Header esperado (según base histórica 2025): incluye "AMIE EJE" (col J) y "AMIE AFC" y "Beneficiarios"
  const headerIdx = findHeaderRow(matrix, ["AMIE EJE", "AMIE AFC", "BENEFICIARIOS"]);
  const out = { sheetName, matchCount: 0, rutas: null, ben: null, presA: null, matches: [] };

  if (headerIdx < 0) return out;

  const header = matrix[headerIdx].map((v) => normUpper(v));
  const idxAmieEje = header.findIndex((x) => x === "AMIE EJE" || x === "AMIE_EJE" || x === "AMIEEJE");
  const idxAfc = header.findIndex((x) => x === "AMIE AFC" || x === "AMIE_AFC" || x === "AMIEAFC");
  const idxBen = header.findIndex((x) => x === "BENEFICIARIOS" || x === "BENEFICIARIO" || x === "BENEFICIARIOS TOTAL");

  const afcSet = new Set();
  let benSum = 0;

  for (let r = headerIdx + 1; r < matrix.length; r++) {
    const row = matrix[r];
    const eje = idxAmieEje >= 0 ? String(row[idxAmieEje] || "").trim().toUpperCase() : "";
    if (!eje || eje !== String(amieEje || "").trim().toUpperCase()) continue;

    const afc = idxAfc >= 0 ? String(row[idxAfc] || "").trim().toUpperCase() : "";
    if (afc) afcSet.add(afc);

    const b = idxBen >= 0 ? toNumber(row[idxBen]) : null;
    if (b !== null) benSum += b;

    out.matches.push({ row: r + 1, afc, ben: b });
  }

  out.matchCount = out.matches.length;
  out.rutas = afcSet.size || null;  // proxy: #AMIE AFC únicos (equivale a rutas en muchos casos)
  out.ben = out.matchCount ? benSum : null;
  return out;
}

// ------------------ Checklist rendering ------------------

function renderChecks(checks, extras = []) {
  const box = $("resultBox");
  if (!box) return;

  const counts = {
    OK: checks.filter(c=>c.status==="OK").length,
    WARN: checks.filter(c=>c.status==="WARN").length,
    FAIL: checks.filter(c=>c.status==="FAIL").length,
  };
  const overall = counts.FAIL ? "Con fallas" : counts.WARN ? "Con observaciones" : "OK";

  const groupOrder = [
    ["INS", "Insumos"],
    ["SHT", "Hojas obligatorias"],
    ["DAT", "Datos por hoja"],
    ["INF", "INFORME"],
    ["RUT", "RUTA"],
    ["DIM", "DIMENSION"],
    ["VAL", "VALIDACIÓN"],
    ["DAEI", "Rangos DAEI"],
    ["CMP", "Comparativo histórico 2025"],
  ];

  const codeToGroup = (codigo) => {
    const p = String(codigo||"").split("-")[0].toUpperCase();
    if (["PDF","XLS","REF","HIS","CQX"].includes(p)) return "INS";
    return p;
  };

  const groups = {};
  for (const c of checks) {
    const g = codeToGroup(c.codigo);
    if (!groups[g]) groups[g] = [];
    groups[g].push(c);
  }

  const badge = (status) => {
    const cls = status==="OK" ? "b-ok" : status==="WARN" ? "b-warn" : "b-fail";
    return `<span class="badge ${cls}">${status}</span>`;
  };

  const rowHtml = (c) => {
    const cls = c.status === "OK" ? "ok" : c.status === "WARN" ? "warn" : "fail";
    const detail = (c.detalle||"").toString().trim();
    const detailHtml = detail ? `<div class="checkDetail">${escapeHtml(detail)}</div>` : "";
    return `
      <div class="checkRow2 ${cls}">
        <div class="checkLeft">
          <div class="checkCode2">${escapeHtml(c.codigo)}</div>
          <div class="checkTitle2">${escapeHtml(c.descripcion)}</div>
          ${detailHtml}
        </div>
        <div class="checkRight">${badge(c.status)}</div>
      </div>`;
  };

  // Highlights (si existe modelData)
  const hi = [];
  if (state?.modelData?.ids?.distrito) hi.push({ k: "Distrito", v: state.modelData.ids.distrito });
  if (state?.modelData?.ids?.amieEje) hi.push({ k: "AMIE Eje", v: state.modelData.ids.amieEje });
  if (state?.modelData?.rutas?.total != null) hi.push({ k: "Rutas", v: String(state.modelData.rutas.total) });
  if (state?.modelData?.rutas?.beneficiariosTotal != null) hi.push({ k: "Beneficiarios", v: String(state.modelData.rutas.beneficiariosTotal) });
  if (state?.modelData?.dimension?.presATotal != null) hi.push({ k: "Presupuesto anual (Σ)", v: formatMoney(state.modelData.dimension.presATotal) });

  // DAEI highlight
  if (state?.modelData?.daei) {
    const d = state.modelData.daei;
    hi.push({ k: "DAEI rango", v: `${d.minimo_f}–${d.maximo_f}` });
  }

  // Histórico highlight
  if (state?.compareData && state?.modelData?.rutas?.total != null) {
    const h = state.compareData;
    if (h.matchCount != null) hi.push({ k: "Histórico coincidencias", v: String(h.matchCount) });
    if (h.rutas != null) hi.push({ k: "Histórico (rutas proxy)", v: String(h.rutas) });
  }

  const highlightsHtml = hi.length ? `
    <div class="panel">
      <div class="panelTitle">Resumen</div>
      <div class="kpiGrid">
        ${hi.map(x=>`<div class="kpi"><div class="kpiK">${escapeHtml(x.k)}</div><div class="kpiV">${escapeHtml(x.v)}</div></div>`).join("")}
      </div>
    </div>` : "";

  const summaryHtml = `
    <div class="panel">
      <div class="panelTitle">Resultado de validación</div>
      <div class="summaryRow">
        <span class="pill">${escapeHtml(overall)}</span>
        <div class="summaryCounts">
          <span class="count ok">OK: ${counts.OK}</span>
          <span class="count warn">WARN: ${counts.WARN}</span>
          <span class="count fail">FAIL: ${counts.FAIL}</span>
        </div>
      </div>
    </div>`;

  const groupsHtml = groupOrder
    .filter(([g]) => (groups[g] && groups[g].length))
    .map(([g, title]) => {
      const items = groups[g];
      const gCounts = {
        OK: items.filter(c=>c.status==="OK").length,
        WARN: items.filter(c=>c.status==="WARN").length,
        FAIL: items.filter(c=>c.status==="FAIL").length,
      };
      const tag = gCounts.FAIL ? "b-fail" : gCounts.WARN ? "b-warn" : "b-ok";
      return `
        <details class="acc" open>
          <summary>
            <span class="accTitle">${escapeHtml(title)}</span>
            <span class="accMeta">
              <span class="badge ${tag}">${gCounts.FAIL? "FAIL" : gCounts.WARN ? "WARN" : "OK"}</span>
              <span class="accCounts">(${gCounts.OK} OK · ${gCounts.WARN} WARN · ${gCounts.FAIL} FAIL)</span>
            </span>
          </summary>
          <div class="accBody">
            ${items.map(rowHtml).join("")}
          </div>
        </details>`;
    })
    .join("");

  const extrasHtml = extras.length
    ? `<details class="acc" open>
        <summary>
          <span class="accTitle">Notas técnicas</span>
          <span class="accMeta"><span class="badge b-warn">INFO</span></span>
        </summary>
        <div class="accBody">
          ${extras.map((t) => `<div class="noteLine">• ${escapeHtml(t)}</div>`).join("")}
        </div>
      </details>`
    : "";

  box.innerHTML = summaryHtml + highlightsHtml + groupsHtml + extrasHtml;
}


function overallStatus(checks) {
  const hasFail = checks.some((c) => c.status === "FAIL");
  if (hasFail) return "FAIL";
  const hasWarn = checks.some((c) => c.status === "WARN");
  if (hasWarn) return "WARN";
  return "OK";
}

// ------------------ Validación principal ------------------
async function validateAll() {
  const checks = [];

  // 1) Insumos
  checks.push({
    codigo: "PDF-01",
    descripcion: "Memorando PDF cargado",
    status: state.memoPdf ? "OK" : "FAIL",
    detalle: state.memoPdf ? state.memoPdf.name : "No se cargó el PDF.",
  });

  if (!state.modelXlsx) {
    checks.push({
      codigo: "XLS-01",
      descripcion: "Excel de modelamiento cargado",
      status: "FAIL",
      detalle: "No se cargó el Excel (.xlsx/.xlsm).",
    });
    return { checks };
  }
  checks.push({
    codigo: "XLS-01",
    descripcion: "Excel de modelamiento cargado",
    status: "OK",
    detalle: state.modelXlsx.name,
  });

  checks.push({
    codigo: "REF-01",
    descripcion: "Referencia Beneficiarios (DAEI) cargada",
    status: state.refXlsx ? "OK" : "WARN",
    detalle: state.refXlsx ? state.refXlsx.name : "No se cargó (recomendado).",
  });

  checks.push({
    codigo: "HIS-01",
    descripcion: "Base histórica (2025) cargada",
    status: state.histXlsx ? "OK" : "WARN",
    detalle: state.histXlsx ? state.histXlsx.name : "No se cargó (opcional).",
  });

  checks.push({
    codigo: "CQX-01",
    descripcion: "Cuadros (Quipux) cargados",
    status: state.cuadrosXlsx ? "OK" : "WARN",
    detalle: state.cuadrosXlsx ? state.cuadrosXlsx.name : "No se cargó (opcional).",
  });

  if (!window.XLSX) {
    checks.push({
      codigo: "SYS-01",
      descripcion: "Librería XLSX disponible",
      status: "FAIL",
      detalle: "No cargó SheetJS (xlsx.full.min.js).",
    });
    return { checks };
  }

  // 2) Leer modelamiento
  const modelBuf = await readAsArrayBuffer(state.modelXlsx);
  const wb = XLSX.read(modelBuf, { type: "array" });

  // 3) Hojas mínimas
  const shInforme = findSheetByName(wb, "INFORME");
  const shRutaName = findSheetByName(wb, "RUTA");
  const shDimName = findSheetByName(wb, "DIMENSION");
  const shValName = findSheetByName(wb, "VALIDACIÓN") || findSheetByName(wb, "VALIDACION");

  const required = [
    { code: "SHT-INF", name: "INFORME", found: shInforme },
    { code: "SHT-RUT", name: "RUTA", found: shRutaName },
    { code: "SHT-DIM", name: "DIMENSION", found: shDimName },
    { code: "SHT-VAL", name: "VALIDACIÓN", found: shValName },
  ];

  for (const r of required) {
    checks.push({
      codigo: r.code,
      descripcion: `Hoja obligatoria: ${r.name}`,
      status: r.found ? "OK" : "FAIL",
      detalle: r.found ? `Detectada: ${r.found}` : "No encontrada.",
    });
  }

  if (!shInforme || !shRutaName || !shDimName || !shValName) {
    return { checks };
  }

  const shInf = wb.Sheets[shInforme];
  const shRuta = wb.Sheets[shRutaName];
  const shDim = wb.Sheets[shDimName];
  const shVal = wb.Sheets[shValName];

  // 4) Hoja con contenido
  const filledInf = countFilledCells(shInf);
  const filledRuta = countFilledCells(shRuta);
  const filledDim = countFilledCells(shDim);
  const filledVal = countFilledCells(shVal);

  checks.push({ codigo: "DAT-INF", descripcion: "INFORME: hoja con datos", status: filledInf >= 30 ? "OK" : "FAIL", detalle: `Celdas con contenido: ${filledInf}` });
  checks.push({ codigo: "DAT-RUT", descripcion: "RUTA: hoja con datos", status: filledRuta >= 30 ? "OK" : "FAIL", detalle: `Celdas con contenido: ${filledRuta}` });
  checks.push({ codigo: "DAT-DIM", descripcion: "DIMENSION: hoja con datos", status: filledDim >= 30 ? "OK" : "FAIL", detalle: `Celdas con contenido: ${filledDim}` });
  checks.push({ codigo: "DAT-VAL", descripcion: "VALIDACIÓN: hoja con datos", status: filledVal >= 20 ? "OK" : "FAIL", detalle: `Celdas con contenido: ${filledVal}` });

  // 5) INFORME – campos críticos + mapa
  const fechaInforme = getCell(shInf, "C2");
  checks.push({ codigo: "INF-01", descripcion: "INFORME: Fecha (C2)", status: isFilled(fechaInforme) ? "OK" : "FAIL", detalle: isFilled(fechaInforme) ? `${fechaInforme}` : "Falta fecha en C2." });

  const devNombre = getCell(shInf, "A105");
  const devCargo = getCell(shInf, "C105");
  const revNombre = getCell(shInf, "A108");
  const revCargo = getCell(shInf, "C108");
  checks.push({ codigo: "INF-02", descripcion: "INFORME: Desarrollo (Nombre/Cargo)", status: isFilled(devNombre) && isFilled(devCargo) ? "OK" : "FAIL", detalle: isFilled(devNombre) ? `${devNombre}` : "Falta A105/C105" });
  checks.push({ codigo: "INF-03", descripcion: "INFORME: Revisión (Nombre/Cargo)", status: isFilled(revNombre) && isFilled(revCargo) ? "OK" : "FAIL", detalle: isFilled(revNombre) ? `${revNombre}` : "Falta A108/C108" });

  // Buscar rótulo MAPEO DE RUTAS
  const infMatrix = XLSX.utils.sheet_to_json(shInf, { header: 1, defval: "" });
  const hasMapeo = infMatrix.some((row) => (row || []).some((v) => normUpper(v).includes("MAPEO DE RUTAS")));
  checks.push({ codigo: "INF-04", descripcion: "INFORME: sección 'MAPEO DE RUTAS'", status: hasMapeo ? "OK" : "FAIL", detalle: hasMapeo ? "Se detectó el rótulo." : "No se detectó el rótulo." });

  const media = await listMediaImages(state.modelXlsx);
  if (media === null) {
    checks.push({ codigo: "INF-05", descripcion: "INFORME: mapa como imagen (xl/media)", status: "WARN", detalle: "No se pudo verificar imágenes (JSZip no cargó)." });
  } else {
    checks.push({ codigo: "INF-05", descripcion: "INFORME: mapa como imagen (xl/media)", status: media.length ? "OK" : "FAIL", detalle: media.length ? `Imágenes detectadas: ${media.length}` : "No se detectaron imágenes (se espera mapa)." });
  }

  // 6) RUTA – cabecera, rutas, coords, distancias
  const rutaParsed = parseRutaSheet(shRuta);

  const headOk = rutaParsed.distrito && rutaParsed.amieEje;
  checks.push({ codigo: "RUT-01", descripcion: "RUTA: Distrito (B3) y AMIE Eje (B4)", status: headOk ? "OK" : "FAIL", detalle: `Distrito=${rutaParsed.distrito || "(vacío)"} • AMIE=${rutaParsed.amieEje || "(vacío)"}` });

  checks.push({ codigo: "RUT-02", descripcion: "RUTA: coordenadas cabecera (H6/H7)", status: rutaParsed.coordHeaderOk ? "OK" : "FAIL", detalle: rutaParsed.coordHeaderOk ? `X=${rutaParsed.coordHeader.x} • Y=${rutaParsed.coordHeader.y}` : "Faltan o no son numéricas." });

  checks.push({ codigo: "RUT-03", descripcion: "RUTA: total rutas y beneficiarios", status: rutaParsed.rows.length ? "OK" : "FAIL", detalle: `Rutas=${rutaParsed.rows.length} • Beneficiarios=${rutaParsed.totalBenef.toFixed(0)}` });

  checks.push({ codigo: "RUT-04", descripcion: "RUTA: consistencia D = F+G+H", status: rutaParsed.inconsistCount === 0 ? "OK" : "WARN", detalle: rutaParsed.inconsistCount === 0 ? "Sin inconsistencias." : `Filas con inconsistencia: ${rutaParsed.inconsistCount}` });

  checks.push({ codigo: "RUT-05", descripcion: "RUTA: coordenadas por ruta (X/Y en tabla)", status: rutaParsed.coordRowsMissingCount === 0 ? "OK" : "FAIL", detalle: rutaParsed.coordRowsMissingCount === 0 ? "Todas las rutas tienen X/Y." : `Rutas con X/Y faltante o #N/D: ${rutaParsed.coordRowsMissingCount}` });

  checks.push({ codigo: "RUT-06", descripcion: "RUTA: distancia lineal numérica", status: rutaParsed.distLinealBadCount === 0 ? "OK" : "WARN", detalle: rutaParsed.distLinealBadCount === 0 ? "OK" : `Rutas con distancia lineal inválida: ${rutaParsed.distLinealBadCount}` });

  // 7) DIMENSION – horas/costos/presupuesto
  const dimParsed = parseDimensionSheet(shDim);
  if (dimParsed.error) {
    checks.push({ codigo: "DIM-00", descripcion: "DIMENSION: estructura de tabla", status: "FAIL", detalle: dimParsed.error });
  } else {
    const issues = dimParsed.timeIssues;
    checks.push({ codigo: "DIM-01", descripcion: "DIMENSION: coherencia horas/costos/presupuestos", status: issues.length === 0 ? "OK" : "WARN", detalle: issues.length === 0 ? "Sin novedades." : `Observaciones: ${issues.length} (ver comparativo/notas)` });
    checks.push({ codigo: "DIM-02", descripcion: "DIMENSION: presupuesto anual total calculado", status: dimParsed.presATotal > 0 ? "OK" : "WARN", detalle: `Σ Presupuesto anual (aprox): ${dimParsed.presATotal.toLocaleString("es-EC", { maximumFractionDigits: 2 })}` });
  }

  // 8) VALIDACIÓN – VERDADERO + horas
  const valParsed = parseValidacionSheet(shVal);
  if (valParsed.error) {
    checks.push({ codigo: "VAL-00", descripcion: "VALIDACIÓN: estructura de tabla", status: "FAIL", detalle: valParsed.error });
  } else {
    checks.push({ codigo: "VAL-01", descripcion: "VALIDACIÓN: Test Ida/Test Retorno en VERDADERO", status: valParsed.notTrueCount === 0 ? "OK" : "FAIL", detalle: valParsed.notTrueCount === 0 ? "Todo VERDADERO" : `Filas no VERDADERO: ${valParsed.notTrueCount}` });
    checks.push({ codigo: "VAL-02", descripcion: "VALIDACIÓN: hora salida < llegada", status: valParsed.timeBadCount === 0 ? "OK" : "WARN", detalle: valParsed.timeBadCount === 0 ? "OK" : `Filas con horas inválidas: ${valParsed.timeBadCount}` });
  }

  // 9) DAEI – rango vs beneficiarios
  let daeiParsed = null;
  if (state.refXlsx && rutaParsed.amieEje) {
    const refWb = XLSX.read(await readAsArrayBuffer(state.refXlsx), { type: "array" });
    daeiParsed = parseDAEIRanges(refWb);
    if (daeiParsed.error) {
      checks.push({ codigo: "DAEI-00", descripcion: "DAEI: estructura válida (AMIE/Mínimo_F/Máximo_F)", status: "FAIL", detalle: daeiParsed.error });
    } else {
      const rng = daeiParsed.ranges[rutaParsed.amieEje];
      if (!rng) {
        checks.push({ codigo: "DAEI-01", descripcion: "DAEI: AMIE encontrado", status: "WARN", detalle: `No se encontró AMIE ${rutaParsed.amieEje} en la base DAEI.` });
      } else {
        const ok = rutaParsed.totalBenef >= rng.min && rutaParsed.totalBenef <= rng.max;
        checks.push({
          codigo: "DAEI-02",
          descripcion: "DAEI: beneficiarios dentro del rango",
          status: ok ? "OK" : "WARN",
          detalle: `Total=${rutaParsed.totalBenef.toFixed(0)} • Rango=${rng.min}–${rng.max}`,
        });
      }
    }
  }

  // 10) Comparativo histórico
  let compareExtras = [];
  if (state.histXlsx && rutaParsed.amieEje) {
    const histWb = XLSX.read(await readAsArrayBuffer(state.histXlsx), { type: "array" });
    const hist = parseHistorico(histWb, rutaParsed.amieEje);
    state.compareData = hist;

    if (hist.matchCount === 0) {
      checks.push({ codigo: "CMP-01", descripcion: "Histórico: existe registro AMIE Eje", status: "WARN", detalle: `No se halló AMIE ${rutaParsed.amieEje} en histórico.` });
    } else {
      checks.push({ codigo: "CMP-01", descripcion: "Histórico: existe registro AMIE Eje", status: "OK", detalle: `Coincidencias: ${hist.matchCount}` });

      if (hist.rutas !== null) {
        const delta = rutaParsed.rows.length - hist.rutas;
        const st = delta === 0 ? "OK" : delta > 0 ? "WARN" : "WARN";
        checks.push({ codigo: "CMP-02", descripcion: "Comparación rutas vs año anterior", status: st, detalle: `Histórico=${hist.rutas} • Actual=${rutaParsed.rows.length} • Δ=${delta}` });
        if (delta !== 0) compareExtras.push(`Rutas: histórico ${hist.rutas} vs actual ${rutaParsed.rows.length} (Δ ${delta}).`);
      } else {
        checks.push({ codigo: "CMP-02", descripcion: "Comparación rutas vs año anterior", status: "WARN", detalle: "No se pudo leer total rutas en histórico." });
      }

      if (hist.ben !== null) {
        const deltaB = rutaParsed.totalBenef - hist.ben;
        compareExtras.push(`Beneficiarios: histórico ${hist.ben} vs actual ${rutaParsed.totalBenef.toFixed(0)} (Δ ${deltaB.toFixed(0)}).`);
      }

      if (hist.presA !== null && !dimParsed.error) {
        const deltaP = dimParsed.presATotal - hist.presA;
        compareExtras.push(`Presupuesto anual (aprox): histórico ${hist.presA} vs actual ${dimParsed.presATotal.toFixed(2)} (Δ ${deltaP.toFixed(2)}).`);
      }
    }
  }

  // Notas de DIMENSION (si hay)
  if (!dimParsed.error && dimParsed.timeIssues.length) {
    compareExtras.push("Observaciones DIMENSION:");
    const max = Math.min(dimParsed.timeIssues.length, 12);
    for (let i = 0; i < max; i++) {
      const it = dimParsed.timeIssues[i];
      compareExtras.push(`Ruta ${it.ruta}: ${it.motivo}`);
    }
    if (dimParsed.timeIssues.length > max) compareExtras.push(`(y ${dimParsed.timeIssues.length - max} más…)`);
  }

  // Construir modelData mínimo para uso posterior
  state.modelData = {
    ids: { distrito: rutaParsed.distrito, amieEje: rutaParsed.amieEje },
    rutas: { total: rutaParsed.rows.length, beneficiariosTotal: rutaParsed.totalBenef, afcList: rutaParsed.afcList },
    dimension: dimParsed.error ? null : { presATotal: dimParsed.presATotal },
    validacion: valParsed.error ? null : { notTrueCount: valParsed.notTrueCount, timeBadCount: valParsed.timeBadCount },
  };

  return { checks, compareExtras };
}

// ------------------ Botones ------------------
async function onValidate() {
  setPill("neutral", "Validando…");
  $("btnGenerate").disabled = true;
  const res = await validateAll();
  state.validation = res;

  const st = overallStatus(res.checks);
  if (st === "OK") {
    setPill("ok", "OK");
    $("btnGenerate").disabled = false;
  } else if (st === "WARN") {
    setPill("warn", "Advertencias");
    $("btnGenerate").disabled = false; // permite generar, pero con advertencias
  } else {
    setPill("fail", "Con fallas");
  }

  renderChecks(res.checks, res.compareExtras || []);
}

function onReset() {
  Object.assign(state, {
    memoPdf: null,
    modelXlsx: null,
    refXlsx: null,
    histXlsx: null,
    cuadrosXlsx: null,
    tplDocx: null,
    memoText: "",
    validation: null,
    modelData: null,
    compareData: null,
  });

  // limpia inputs
  ["memoFile", "modelFile", "refFile", "histFile", "cuadrosFile", "tplFile"].forEach((id) => {
    const el = $(id);
    if (el) el.value = "";
  });
  setMeta($("memoMeta"), null);
  setMeta($("modelMeta"), null);
  setMeta($("refMeta"), null);
  setMeta($("histMeta"), null);
  setMeta($("cuadrosMeta"), null);
  // tpl meta se deja como texto fijo

  setPill("neutral", "Sin ejecutar");
  $("resultBox").innerHTML = '<div class="placeholder">Ejecuta <b>Validar</b> para ver el checklist.</div>';
  $("btnGenerate").disabled = true;

  showRuntimeNote("");
  const rn = $("runtimeNote");
  if (rn) rn.style.display = "none";
}

// ------------------ Inicialización ------------------
function initUI() {
  // binds dropboxes
  safeBindDrop({
    boxId: "dropMemo",
    inputId: "memoFile",
    metaId: "memoMeta",
    accept: ".pdf",
    onFile: async (file) => {
      state.memoPdf = file;
      // parse PDF
      try {
        const text = await extractPdfText(file);
        state.memoText = text;
        const guessed = guessMemoFields(text);
        // no sobreescribir si el usuario ya escribió algo
        if (!$("fMemoNro").value) $("fMemoNro").value = guessed.memo_nro || "";
        if (!$("fMemoFecha").value) $("fMemoFecha").value = guessed.memo_fecha || "";
        if (!$("fAsunto").value) $("fAsunto").value = guessed.asunto || "";
      } catch (e) {
        console.warn("No se pudo extraer texto del PDF:", e);
      }
    },
  });

  safeBindDrop({
    boxId: "dropModel",
    inputId: "modelFile",
    metaId: "modelMeta",
    accept: ".xlsx,.xlsm",
    onFile: async (file) => {
      state.modelXlsx = file;
    },
  });

  safeBindDrop({
    boxId: "dropRef",
    inputId: "refFile",
    metaId: "refMeta",
    accept: ".xlsx",
    onFile: async (file) => {
      state.refXlsx = file;
    },
  });

  safeBindDrop({
    boxId: "dropHist",
    inputId: "histFile",
    metaId: "histMeta",
    accept: ".xlsx",
    onFile: async (file) => {
      state.histXlsx = file;
    },
  });

  safeBindDrop({
    boxId: "dropCuadros",
    inputId: "cuadrosFile",
    metaId: "cuadrosMeta",
    accept: ".xlsx",
    onFile: async (file) => {
      state.cuadrosXlsx = file;
    },
  });

  safeBindDrop({
    boxId: "dropTpl",
    inputId: "tplFile",
    metaId: "tplMeta",
    accept: ".docx",
    onFile: async (file) => {
      state.tplDocx = file;
      setMeta($("tplMeta"), file);
    },
  });

  // buttons
  const btnValidate = $("btnValidate");
  if (btnValidate) btnValidate.addEventListener("click", onValidate);

  const btnReset = $("btnReset");
  if (btnReset) btnReset.addEventListener("click", onReset);

  // sync memo fields
  [
    ["fMemoNro", "memo_nro"],
    ["fMemoFecha", "memo_fecha"],
    ["fPara", "para"],
    ["fDe", "de"],
    ["fAsunto", "asunto"],
    ["fAnioLectivo", "anio_lectivo"],
    ["fFirma", "firma"],
  ].forEach(([id, key]) => {
    const el = $(id);
    if (!el) return;
    el.addEventListener("input", () => {
      state.memoFields[key] = el.value;
    });
  });

  // initial state
  setPill("neutral", "Sin ejecutar");
}

document.addEventListener("DOMContentLoaded", initUI);

function excelDateToStr(v) {
  // Excel serial number to YYYY-MM-DD (best-effort). If string, return as-is.
  if (typeof v === "number" && isFinite(v)) {
    // Excel epoch 1899-12-30
    const utc_days = Math.floor(v - 25569);
    const utc_value = utc_days * 86400; 
    const date_info = new Date(utc_value * 1000);
    const y = date_info.getUTCFullYear();
    const m = String(date_info.getUTCMonth() + 1).padStart(2, "0");
    const d = String(date_info.getUTCDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }
  const s = String(v ?? "").trim();
  return s;
}


function escapeHtml(str){
  return String(str ?? "").replace(/[&<>"']/g, (m)=>({ "&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#039;" }[m]));
}
function formatMoney(v){
  const n = toNumber(v);
  if (n===null) return "";
  return n.toLocaleString("es-EC", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}


