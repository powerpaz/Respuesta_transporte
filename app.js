/* Respuesta | Validador + Generador Word (cliente, sin backend) */

if (window.pdfjsLib) {
  pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
} else {
  console.warn("PDF.js no cargó; se desactiva extracción automática del PDF.");
}


const state = {
  memoPdf: null,
  modelXlsx: null,
  refXlsx: null,
  tplDocx: null,
  memoText: "",
  validation: null,
};

const $ = (id) => document.getElementById(id);

function setMeta(el, file){
  if(!file){
    el.classList.remove("ok");
    el.textContent = "Sin archivo";
    return;
  }
  el.classList.add("ok");
  el.textContent = `${file.name} • ${(file.size/1024).toFixed(1)} KB`;
}

function bindDrop(boxId, inputId, metaId, onFile){
  const box = $(boxId);
  const input = $(inputId);
  const meta = $(metaId);

  const handle = async (file) => {
    if(!file) return;
    setMeta(meta, file);
    await onFile(file);
  };

  input.addEventListener("change", (e)=>handle(e.target.files?.[0] || null));

  box.addEventListener("dragover", (e)=>{ e.preventDefault(); box.classList.add("dragover"); });
  box.addEventListener("dragleave", ()=> box.classList.remove("dragover"));
  box.addEventListener("drop", async (e)=>{
    e.preventDefault();
    box.classList.remove("dragover");
    const file = e.dataTransfer.files?.[0];
    await handle(file);
  });
}

async function readAsArrayBuffer(file){
  return await file.arrayBuffer();
}

function isFilled(v){
  if(v === null || v === undefined) return false;
  if(typeof v === "string") return v.trim() !== "" && v.trim().toUpperCase() !== "NOMBRE";
  return true;
}

function countFilledCells(sheet){
  let count = 0;
  for (const k of Object.keys(sheet)){
    if(k.startsWith("!")) continue;
    const v = sheet[k]?.v;
    if(isFilled(v)) count++;
  }
  return count;
}

async function hasAnyMediaImage(excelFile){
  if(!window.JSZip) return null;
  const buf = await readAsArrayBuffer(excelFile);
  const zip = await JSZip.loadAsync(buf);
  const media = Object.keys(zip.files).filter(p => p.startsWith("xl/media/"));
  return media.length > 0 ? media : [];
}

function getNumber(v){
  const n = (typeof v === "number") ? v : parseFloat((v||"").toString().replace(",", "."));
  return Number.isFinite(n) ? n : null;
}

function parseRutaBeneficiarios(sheetRuta){
  let total = 0;
  let rowIssues = 0;
  let rows = 0;

  for(let r=10; r<2000; r++){
    const ruta = getCell(sheetRuta, `A${r}`);
    if(ruta === null || ruta === undefined || ruta === "") break;
    rows++;

    const dTot = getNumber(getCell(sheetRuta, `D${r}`)) ?? 0;
    const fIni = getNumber(getCell(sheetRuta, `F${r}`)) ?? 0;
    const gEgb = getNumber(getCell(sheetRuta, `G${r}`)) ?? 0;
    const hBac = getNumber(getCell(sheetRuta, `H${r}`)) ?? 0;

    const sum = fIni + gEgb + hBac;
    if(Math.abs(sum - dTot) > 0.5){
      rowIssues++;
    }
    total += dTot;
  }
  return {total, rows, rowIssues};
}

function parseDAEIRanges(refWb){
  const sheetName = refWb.SheetNames[0];
  const sh = refWb.Sheets[sheetName];
  const matrix = XLSX.utils.sheet_to_json(sh, {header:1, defval:null});
  let headerRow = -1;
  let idxAmie=-1, idxMin=-1, idxMax=-1;

  for(let i=0;i<matrix.length;i++){
    const row = matrix[i] || [];
    const norm = row.map(v => normName(v));
    if(norm.includes("AMIE") && (norm.includes("MINIMO_F") || norm.includes("MINIMO F")) && (norm.includes("MAXIMO_F") || norm.includes("MAXIMO F"))){
      headerRow=i;
      idxAmie = norm.indexOf("AMIE");
      idxMin  = norm.indexOf("MINIMO_F"); if(idxMin<0) idxMin = norm.indexOf("MINIMO F");
      idxMax  = norm.indexOf("MAXIMO_F"); if(idxMax<0) idxMax = norm.indexOf("MAXIMO F");
      break;
    }
  }
  if(headerRow<0) return {ranges:null, error:"No se encontró encabezado con AMIE / Mínimo_F / Máximo_F en la referencia DAEI."};

  const ranges = {};
  for(let i=headerRow+1;i<matrix.length;i++){
    const row = matrix[i] || [];
    const amie = (row[idxAmie]||"").toString().trim();
    if(!amie) continue;
    const min = getNumber(row[idxMin]);
    const max = getNumber(row[idxMax]);
    if(min===null || max===null) continue;
    ranges[amie] = {min, max};
  }
  return {ranges, error:null};
}


async function extractPdfText(file){
  if(!window.pdfjsLib){
    return "";
  }
  const data = await readAsArrayBuffer(file);
  const pdf = await pdfjsLib.getDocument({data}).promise;
  let full = "";
  for(let i=1; i<=pdf.numPages; i++){
    const page = await pdf.getPage(i);
    const content = await page.getTextContent();
    const strings = content.items.map(it => it.str);
    full += strings.join(" ") + "\n";
  }
  return full;
}

function guessMemoFields(text){
  const out = { memo_nro:"", memo_fecha:"", asunto:"" };

  // Memorando Nro.
  const nroMatch = text.match(/\b[A-ZÁÉÍÓÚÜÑ]{3,}-[A-ZÁÉÍÓÚÜÑ0-9]{2,}-\d{4}-\d{4,}-M\b/);
  if(nroMatch) out.memo_nro = nroMatch[0];

  // Fecha (muy flexible)
  const dateMatch = text.match(/\b(\d{1,2}\s+de\s+[A-Za-zÁÉÍÓÚÜÑáéíóúñ]+(?:\s+de)?\s+\d{4})\b/);
  if(dateMatch) out.memo_fecha = dateMatch[1];

  // Asunto
  const asuntoMatch = text.match(/ASUNTO\s*[:\-]\s*(.+)/i);
  if(asuntoMatch) out.asunto = asuntoMatch[1].trim().slice(0, 180);

  return out;
}

function normalizeHeader(s){
  return String(s||"").trim().toUpperCase().replace(/\s+/g," ");
}

function normName(s){
  return (s||"")
    .toString()
    .normalize("NFD")
    .replace(/\p{Diacritic}/gu, "")
    .toUpperCase()
    .trim();
}

function findSheetByName(wb, name){
  const target = normName(name);
  const hit = wb.SheetNames.find(n => normName(n) === target)
           || wb.SheetNames.find(n => normName(n).includes(target));
  return hit || null;
}


function getCell(sheet, addr){
  const cell = sheet[addr];
  if(!cell) return null;
  return cell.v;
}

function asNumber(v){
  if(v === null || v === undefined || v === "") return null;
  const n = Number(String(v).replace(",", "."));
  return Number.isFinite(n) ? n : null;
}

function renderChecklist(checks){
  const rows = checks.map(c=>{
    const cls = c.status === "OK" ? "ok" : (c.status === "WARN" ? "warn" : "bad");
    return `<tr>
      <td class="k">${c.codigo}</td>
      <td>${c.descripcion}</td>
      <td>${c.detalle || ""}</td>
      <td><span class="tag ${cls}">${c.status}</span></td>
    </tr>`;
  }).join("");

  return `<table>
    <thead>
      <tr>
        <th>Código</th>
        <th>Validación</th>
        <th>Detalle</th>
        <th>Estado</th>
      </tr>
    </thead>
    <tbody>${rows}</tbody>
  </table>`;
}

function overallStatus(checks){
  if(checks.some(c=>c.status==="FAIL")) return "FAIL";
  if(checks.some(c=>c.status==="WARN")) return "WARN";
  return "OK";
}

function setPill(status){
  const pill = $("statusPill");
  pill.classList.remove("neutral","ok","bad","warn");
  if(status==="OK"){ pill.classList.add("ok"); pill.textContent="VALIDADO"; }
  else if(status==="WARN"){ pill.classList.add("warn"); pill.textContent="VALIDADO CON OBSERVACIONES"; }
  else if(status==="FAIL"){ pill.classList.add("bad"); pill.textContent="CON OBSERVACIONES"; }
  else { pill.classList.add("neutral"); pill.textContent="Sin ejecutar"; }
}


function showGenMsg(opts){
  const el = $("genMsg");
  if(!opts){
    el.style.display = "none";
    el.className = "genMsg";
    el.innerHTML = "";
    return;
  }
  const { type, title, lines } = opts;
  el.style.display = "block";
  el.className = "genMsg " + (type || "");
  const items = (lines || []).map(t=>`<li>${escapeHtml(t)}</li>`).join("");
  el.innerHTML = `<b>${escapeHtml(title || "")}</b>${items ? `<ul>${items}</ul>` : ""}`;
}


function buildResumenText(checks){
  const lines = checks.map(c=>{
    const mark = c.status==="OK" ? "✓" : (c.status==="WARN" ? "!" : "✗");
    return `${mark} [${c.codigo}] ${c.descripcion}${c.detalle ? " — " + c.detalle : ""}`;
  });
  return lines.join("\n");
}

async function validateAll(){
  const checks = [];

  // ========= 1) Insumos =========
  if(!state.memoPdf){
    checks.push({codigo:"PDF-01", descripcion:"Memorando PDF cargado", status:"FAIL", detalle:"No se cargó el PDF."});
  } else {
    checks.push({codigo:"PDF-01", descripcion:"Memorando PDF cargado", status:"OK", detalle: state.memoPdf.name});
  }

  if(!state.modelXlsx){
    checks.push({codigo:"XLS-01", descripcion:"Excel de modelamiento cargado", status:"FAIL", detalle:"No se cargó el Excel (.xlsx/.xlsm)."});
    return {checks};
  } else {
    checks.push({codigo:"XLS-01", descripcion:"Excel de modelamiento cargado", status:"OK", detalle: state.modelXlsx.name});
  }

  // ========= 2) Leer libro =========
  if(!window.XLSX){
    checks.push({codigo:"SYS-01", descripcion:"Librería XLSX disponible", status:"FAIL", detalle:"No cargó SheetJS (xlsx.full.min.js). Revisa conectividad/CDN."});
    return {checks};
  }

  const modelBuf = await readAsArrayBuffer(state.modelXlsx);
  const wb = XLSX.read(modelBuf, {type:"array"});

  // ========= 3) Hojas obligatorias =========
  const required = [
    {key:"INFORME", label:"INFORME"},
    {key:"RUTA", label:"RUTA"},
    {key:"DIMENSION", label:"DIMENSION"},
    {key:"VALIDACION", label:"VALIDACIÓN / VALIDACION"},
    {key:"DICCIONARIO", label:"DICCIONARIO DE VARIABLES"},
    {key:"GIEE", label:"GIEE"},
  ];

  const sheets = {};
  for(const r of required){
    let name = null;
    if(r.key==="VALIDACION"){
      name = findSheetByName(wb, "VALIDACIÓN") || findSheetByName(wb, "VALIDACION");
    } else if(r.key==="DICCIONARIO"){
      name = findSheetByName(wb, "DICCIONARIO DE VARIABLES");
    } else {
      name = findSheetByName(wb, r.label);
    }
    sheets[r.key]=name;

    checks.push({
      codigo:`SHT-${r.key}`,
      descripcion:`Hoja obligatoria: ${r.label}`,
      status: name ? "OK":"FAIL",
      detalle: name ? `Detectada: ${name}` : "No encontrada."
    });
  }

  const missingCritical = ["INFORME","RUTA","DIMENSION","VALIDACION","DICCIONARIO","GIEE"].filter(k => !sheets[k]);
  if(missingCritical.length){
    checks.push({codigo:"SHT-STOP", descripcion:"Estructura mínima del modelamiento", status:"FAIL", detalle:`Faltan hojas: ${missingCritical.join(", ")}`});
    return {checks};
  }

  // ========= 4) Datos llenos =========
  for(const k of Object.keys(sheets)){
    const sh = wb.Sheets[sheets[k]];
    const filled = countFilledCells(sh);
    checks.push({
      codigo:`DAT-${k}`,
      descripcion:`Hoja ${sheets[k]} con datos`,
      status: filled >= 20 ? "OK" : "FAIL",
      detalle: `Celdas con contenido detectadas: ${filled} (mínimo esperado: 20).`
    });
  }

  // ========= 5) INFORME =========
  const shInf = wb.Sheets[sheets["INFORME"]];
  const fechaInforme = getCell(shInf, "C2");
  checks.push({
    codigo:"INF-01",
    descripcion:"INFORME: Fecha de Informe (C2)",
    status: isFilled(fechaInforme) ? "OK":"FAIL",
    detalle: isFilled(fechaInforme) ? `Fecha: ${fechaInforme}` : "No se detectó fecha en C2."
  });

  const devNombre = getCell(shInf, "A105");
  const devCargo  = getCell(shInf, "C105");
  const revNombre = getCell(shInf, "A108");
  const revCargo  = getCell(shInf, "C108");
  const aprNombre = getCell(shInf, "A111");
  const aprCargo  = getCell(shInf, "C111");

  checks.push({codigo:"INF-02", descripcion:"INFORME: Desarrollo del documento (Nombre/Cargo)", status: (isFilled(devNombre)&&isFilled(devCargo))?"OK":"FAIL", detalle: (isFilled(devNombre)&&isFilled(devCargo)) ? `${devNombre} • ${devCargo}` : "Falta Nombre o Cargo en Desarrollo (A105/C105)."});
  checks.push({codigo:"INF-03", descripcion:"INFORME: Revisión del documento (Nombre/Cargo)", status: (isFilled(revNombre)&&isFilled(revCargo))?"OK":"FAIL", detalle: (isFilled(revNombre)&&isFilled(revCargo)) ? `${revNombre} • ${revCargo}` : "Falta Nombre o Cargo en Revisión (A108/C108)."});
  checks.push({codigo:"INF-04", descripcion:"INFORME: Aprobación del documento (Nombre/Cargo)", status: (isFilled(aprNombre)&&isFilled(aprCargo))?"OK":"FAIL", detalle: (isFilled(aprNombre)&&isFilled(aprCargo)) ? `${aprNombre} • ${aprCargo}` : "Falta Nombre o Cargo en Aprobación (A111/C111)."});

  const infMatrix = XLSX.utils.sheet_to_json(shInf, {header:1, defval:""});
  const hasMapeoText = infMatrix.some(row => row.some(v => normName(v).includes("MAPEO DE RUTAS")));
  checks.push({codigo:"INF-05", descripcion:"INFORME: sección MAPEO DE RUTAS presente", status: hasMapeoText?"OK":"FAIL", detalle: hasMapeoText ? "Se detectó el rótulo 'MAPEO DE RUTAS'." : "No se detectó 'MAPEO DE RUTAS' en la hoja INFORME."});

  const media = await hasAnyMediaImage(state.modelXlsx);
  if(media === null){
    checks.push({codigo:"INF-06", descripcion:"INFORME: mapa como imagen (xl/media)", status:"WARN", detalle:"No se pudo verificar imágenes (JSZip no cargó)."});
  } else {
    checks.push({codigo:"INF-06", descripcion:"INFORME: mapa como imagen (xl/media)", status: media.length ? "OK":"FAIL", detalle: media.length ? `Imágenes detectadas: ${media.length}` : "No se detectaron imágenes en xl/media/ (se espera el mapa como imagen)."});
  }

  // ========= 6) RUTA =========
  const shRuta = wb.Sheets[sheets["RUTA"]];
  const distritoCod = getCell(shRuta, "B3");
  const distritoNom = getCell(shRuta, "E3");
  const amieEje     = getCell(shRuta, "B4");
  const ieEjeNom    = getCell(shRuta, "E4");
  const provincia   = getCell(shRuta, "B6");
  const parroquia   = getCell(shRuta, "E6");
  const canton      = getCell(shRuta, "B7");
  const coordX      = getCell(shRuta, "H6");
  const coordY      = getCell(shRuta, "H7");

  const rutaHeadOk = [distritoCod,distritoNom,amieEje,ieEjeNom,provincia,parroquia,canton,coordX,coordY].every(isFilled);
  checks.push({codigo:"RUT-01", descripcion:"RUTA: cabecera completa (Distrito/AMIE/Ubicación/Coordenadas)", status: rutaHeadOk?"OK":"FAIL", detalle: rutaHeadOk ? `AMIE IE EJE: ${amieEje}` : "Faltan datos en cabecera (B3/E3/B4/E4/B6/E6/B7/H6/H7)."});

  const br = parseRutaBeneficiarios(shRuta);
  checks.push({
    codigo:"RUT-02",
    descripcion:"RUTA: Beneficiarios por ruta (tabla desde fila 10)",
    status: br.rows>0 ? (br.rowIssues===0 ? "OK":"WARN") : "FAIL",
    detalle: br.rows>0 ? `Filas: ${br.rows} • Total beneficiarios (col D): ${br.total.toFixed(0)} • Filas con inconsistencia D ≠ (F+G+H): ${br.rowIssues}` : "No se detectaron filas de rutas desde la fila 10."
  });

  // ========= 7) DIMENSION =========
  const shDim = wb.Sheets[sheets["DIMENSION"]];
  const dimHeaderOk = isFilled(getCell(shDim,"A12")) && isFilled(getCell(shDim,"B12")) && isFilled(getCell(shDim,"E12"));
  const dimHasRow = isFilled(getCell(shDim,"A13")) || isFilled(getCell(shDim,"B13")) || isFilled(getCell(shDim,"C13"));
  checks.push({codigo:"DIM-01", descripcion:"DIMENSION: encabezados y al menos 1 fila de dimensionamiento", status: (dimHeaderOk && dimHasRow)?"OK":"FAIL", detalle: (dimHeaderOk && dimHasRow) ? "Encabezados OK y datos presentes." : "Faltan encabezados en fila 12 o no hay datos en filas siguientes."});

  // ========= 8) VALIDACIÓN =========
  const shVal = wb.Sheets[sheets["VALIDACION"]];
  const valHasHeader = isFilled(getCell(shVal,"A1")) && isFilled(getCell(shVal,"B1"));
  const valHasRow = isFilled(getCell(shVal,"A2")) && isFilled(getCell(shVal,"B2"));
  checks.push({codigo:"VAL-01", descripcion:"VALIDACIÓN: encabezado y filas con datos", status: (valHasHeader && valHasRow) ? "OK":"FAIL", detalle: (valHasHeader && valHasRow) ? "Datos de validación presentes." : "No se detectan filas con datos en VALIDACIÓN."});

  // ========= 9) Rangos DAEI =========
  if(state.refXlsx && isFilled(amieEje)){
    const refWb = XLSX.read(await readAsArrayBuffer(state.refXlsx), {type:"array"});
    const parsed = parseDAEIRanges(refWb);
    if(parsed.error){
      checks.push({codigo:"RNG-00", descripcion:"Referencia DAEI: estructura válida (AMIE/Mínimo_F/Máximo_F)", status:"FAIL", detalle: parsed.error});
    } else {
      const rng = parsed.ranges[amieEje];
      if(!rng){
        checks.push({codigo:"RNG-01", descripcion:"Rangos DAEI para AMIE IE EJE", status:"WARN", detalle:`No se encontró AMIE ${amieEje} en la referencia DAEI.`});
      } else {
        const ok = br.total >= rng.min && br.total <= rng.max;
        checks.push({codigo:"RNG-02", descripcion:"Beneficiarios del modelamiento dentro del rango DAEI (Mínimo_F–Máximo_F)", status: ok?"OK":"FAIL", detalle:`AMIE ${amieEje} • Total(modelamiento): ${br.total.toFixed(0)} • Rango DAEI: ${rng.min}–${rng.max}`});
      }
    }
  } else {
    checks.push({codigo:"RNG-02", descripcion:"Rangos DAEI (validación recomendada)", status:"WARN", detalle:"Carga la referencia DAEI para validar rangos por AMIE."});
  }

  return {checks};
}


async function loadTemplateBytes(){
  if(state.tplDocx){
    return new Uint8Array(await readAsArrayBuffer(state.tplDocx));
  }
  // plantilla incluida
  const res = await fetch("templates/plantilla_respuesta.docx");
  const ab = await res.arrayBuffer();
  return new Uint8Array(ab);
}

async function generateDocx(){
  // Reglas: debe existir validación, insumos core y no tener FAIL
  if(!state.modelXlsx || !state.memoPdf){
    showGenMsg({type:"fail", title:"No se puede generar la respuesta", lines:[
      "Carga el Memorando (PDF) y el Modelamiento (Excel) y ejecuta Validar."
    ]});
    return;
  }
  if(!state.validation){
    showGenMsg({type:"fail", title:"No se puede generar la respuesta", lines:[
      "Primero ejecuta Validar para construir el checklist."
    ]});
    return;
  }

  const checks = state.validation.checks;
  const status = overallStatus(checks);
  const hasFail = checks.some(c => c.level === "FAIL");
  if(hasFail){
    showGenMsg({type:"fail", title:"No se puede generar la respuesta", lines:[
      "Existen hallazgos tipo FAIL en el checklist.",
      "Corrige el modelamiento y vuelve a ejecutar Validar."
    ]});
    return;
  }
  const resumen = buildResumenText(checks);

  const resultado_general = status==="OK" ? "VALIDADO" : (status==="WARN" ? "VALIDADO CON OBSERVACIONES" : "CON OBSERVACIONES");
  const conclusion = status==="OK"
    ? "Con base en la validación automática realizada, la documentación y anexos cumplen los criterios revisados, por lo que se procede conforme a lo solicitado."
    : "Con base en la validación automática realizada, se identifican observaciones que deben ser ajustadas por la unidad remitente previo a continuar el trámite."

  const data = {
    memo_nro: $("fMemoNro").value || "(s/n)",
    memo_fecha: $("fMemoFecha").value || "(s/f)",
    para: $("fPara").value || "(s/d)",
    de: $("fDe").value || "(s/d)",
    asunto: $("fAsunto").value || "(s/a)",
    resultado_general,
    resumen_texto: resumen,
    conclusion,
    firma: $("fFirma").value || "(falta firma)"
  };

  const content = await loadTemplateBytes();
  const zip = new PizZip(content);
  const doc = new window.docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
  doc.render(data);

  const out = doc.getZip().generate({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
  const fname = `Respuesta_${data.memo_nro.replace(/[^\w\-]+/g,"_")}.docx`;
  saveAs(out, fname);

  const inherited = [];
  inherited.push(`Memorando (PDF): ${state.memoPdf.name}`);
  inherited.push(`Modelamiento (Excel): ${state.modelXlsx.name}`);
  inherited.push(state.refXlsx ? `Referencia Beneficiarios (Excel): ${state.refXlsx.name}` : "Referencia Beneficiarios (Excel): (no cargado)");
  inherited.push(state.tplDocx ? `Plantilla (DOCX): ${state.tplDocx.name}` : "Plantilla (DOCX): plantilla incluida / oficial");
  inherited.push(`Estado de validación: ${resultado_general}`);

  showGenMsg({type:"ok", title:"Respuesta generada correctamente", lines:[
    `Archivo: ${fname}`,
    "La respuesta hereda los datos del memorando + modelamiento (y referencias si fueron cargadas).",
    ...inherited
  ]});
}

function resetAll(){
  state.memoPdf = null;
  state.modelXlsx = null;
  state.refXlsx = null;
  state.tplDocx = null;
  state.memoText = "";
  state.validation = null;

  $("memoFile").value="";
  $("modelFile").value="";
  $("refFile").value="";
  $("tplFile").value="";

  setMeta($("memoMeta"), null);
  setMeta($("modelMeta"), null);
  setMeta($("refMeta"), null);
  $("tplMeta").innerHTML = 'Incluida: <code>templates/plantilla_respuesta.docx</code>';

  $("resultBox").innerHTML = `<div class="placeholder">Ejecuta <b>Validar</b> para ver el checklist.</div>`;
  setPill("NONE");
  $("btnGenerate").disabled = true;
  showGenMsg(null);

  // limpiar campos memo
  ["fMemoNro","fMemoFecha","fPara","fDe","fAsunto","fFirma"].forEach(id=>$(id).value="");
}

bindDrop("dropMemo", "memoFile", "memoMeta", async (file)=>{
  state.memoPdf = file;
  state.memoText = await extractPdfText(file);
  const g = guessMemoFields(state.memoText);
  if(g.memo_nro) $("fMemoNro").value = g.memo_nro;
  if(g.memo_fecha) $("fMemoFecha").value = g.memo_fecha;
  if(g.asunto) $("fAsunto").value = g.asunto;
});

bindDrop("dropModel", "modelFile", "modelMeta", async (file)=>{ state.modelXlsx = file; });
bindDrop("dropRef", "refFile", "refMeta", async (file)=>{ state.refXlsx = file; });
bindDrop("dropTpl", "tplFile", "tplMeta", async (file)=>{
  state.tplDocx = file;
  $("tplMeta").classList.add("ok");
  $("tplMeta").textContent = `${file.name} • ${(file.size/1024).toFixed(1)} KB`;
});

$("btnValidate").addEventListener("click", async ()=>{
  const res = await validateAll();
  state.validation = res;

  const checks = res.checks;
  const status = overallStatus(checks);
  setPill(status);

  $("resultBox").innerHTML = renderChecklist(checks);

  // habilitar generar SOLO si: hay PDF + Excel y NO existen FAIL
  const hasCore = !!state.modelXlsx && !!state.memoPdf;
  const hasFail = checks.some(c => c.level === "FAIL");
  $("btnGenerate").disabled = !(hasCore && !hasFail);

  // limpiar mensaje de generación
  showGenMsg(null);


  // autocompletar "Para" / "De" si no están (placeholder)
  if(!$("fPara").value) $("fPara").value = "Coordinación Zonal / Dirección Distrital";
  if(!$("fDe").value) $("fDe").value = "Dirección / Analista responsable";
  if(!$("fFirma").value) $("fFirma").value = "________________________";
});

$("btnGenerate").addEventListener("click", generateDocx);
$("btnReset").addEventListener("click", resetAll);

// inicio
resetAll();
