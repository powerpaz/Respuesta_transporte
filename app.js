/* Respuesta | Validador + Generador Word (cliente, sin backend) */

pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/pdf.worker.min.js";

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

async function extractPdfText(file){
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

function findSheetByName(wb, name){
  const target = name.toUpperCase();
  const hit = wb.SheetNames.find(n => n.toUpperCase() === target)
           || wb.SheetNames.find(n => n.toUpperCase().includes(target));
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

function buildResumenText(checks){
  const lines = checks.map(c=>{
    const mark = c.status==="OK" ? "✓" : (c.status==="WARN" ? "!" : "✗");
    return `${mark} [${c.codigo}] ${c.descripcion}${c.detalle ? " — " + c.detalle : ""}`;
  });
  return lines.join("\n");
}

async function validateAll(){
  const checks = [];

  // PDF
  if(!state.memoPdf){
    checks.push({codigo:"PDF-01", descripcion:"Memorando PDF cargado", status:"FAIL", detalle:"No se cargó el PDF."});
  } else {
    checks.push({codigo:"PDF-01", descripcion:"Memorando PDF cargado", status:"OK"});
  }

  // Modelamiento
  if(!state.modelXlsx){
    checks.push({codigo:"XLS-01", descripcion:"Excel de modelamiento cargado", status:"FAIL", detalle:"No se cargó el Excel."});
    return {checks}; // sin Excel no hay más
  } else {
    checks.push({codigo:"XLS-01", descripcion:"Excel de modelamiento cargado", status:"OK"});
  }

  const wb = XLSX.read(await readAsArrayBuffer(state.modelXlsx), {type:"array"});
  const rutaName = findSheetByName(wb, "RUTA");
  const dimName  = findSheetByName(wb, "DIMENSION");

  checks.push({
    codigo:"XLS-02",
    descripcion:"Hoja RUTA existe",
    status: rutaName ? "OK":"FAIL",
    detalle: rutaName ? `Detectada: ${rutaName}` : "No se encontró hoja RUTA."
  });
  checks.push({
    codigo:"XLS-03",
    descripcion:"Hoja DIMENSION existe",
    status: dimName ? "OK":"FAIL",
    detalle: dimName ? `Detectada: ${dimName}` : "No se encontró hoja DIMENSION."
  });

  const baseSheetName = rutaName || wb.SheetNames[0];
  const sheet = wb.Sheets[baseSheetName];

  // AMIE / beneficiarios (estructura conocida)
  const amie = getCell(sheet, "B4");
  checks.push({
    codigo:"BEN-01",
    descripcion:"AMIE en celda B4",
    status: amie ? "OK":"FAIL",
    detalle: amie ? `AMIE: ${amie}` : "No hay dato en B4."
  });

  const bIni = getCell(sheet, "F10");
  const bEgb = getCell(sheet, "G10");
  const bBac = getCell(sheet, "H10");

  const hasBen = (bIni!==null && bIni!==undefined) || (bEgb!==null && bEgb!==undefined) || (bBac!==null && bBac!==undefined);
  checks.push({
    codigo:"BEN-02",
    descripcion:"Beneficiarios reportados (F10/G10/H10)",
    status: hasBen ? "OK":"FAIL",
    detalle: hasBen ? `Inicial:${bIni ?? "-"} • EGB:${bEgb ?? "-"} • Bach:${bBac ?? "-"}` : "No hay valores en F10/G10/H10."
  });

  // Validación vs referencia
  if(state.refXlsx && amie){
    const refWb = XLSX.read(await readAsArrayBuffer(state.refXlsx), {type:"array"});
    const refSheetName = refWb.SheetNames[0];
    const refSheet = refWb.Sheets[refSheetName];

    // convertir a json usando primera fila como encabezados
    const rows = XLSX.utils.sheet_to_json(refSheet, {defval:""});
    const hdrs = rows.length ? Object.keys(rows[0]).map(normalizeHeader) : [];

    const colAMIE = Object.keys(rows[0]||{}).find(k => normalizeHeader(k)==="AMIE" || normalizeHeader(k).includes("AMIE"));
    const colIni  = Object.keys(rows[0]||{}).find(k => normalizeHeader(k).includes("INICIAL"));
    const colEgb  = Object.keys(rows[0]||{}).find(k => normalizeHeader(k)==="EGB" || normalizeHeader(k).includes("EGB"));
    const colBac  = Object.keys(rows[0]||{}).find(k => normalizeHeader(k).includes("BACH"));

    if(!colAMIE || !colIni || !colEgb || !colBac){
      checks.push({
        codigo:"BEN-03",
        descripcion:"Referencia de beneficiarios (columnas)",
        status:"FAIL",
        detalle:"La referencia debe tener columnas AMIE, INICIAL, EGB, BACHILLERATO (o nombres equivalentes)."
      });
    } else {
      const row = rows.find(r => String(r[colAMIE]).trim() === String(amie).trim());
      if(!row){
        checks.push({
          codigo:"BEN-04",
          descripcion:"AMIE existe en tabla de referencia",
          status:"FAIL",
          detalle:`No se encontró AMIE ${amie} en el Excel de referencia.`
        });
      } else {
        const expIni = row[colIni];
        const expEgb = row[colEgb];
        const expBac = row[colBac];

        const okIni = String(expIni).trim() === String(bIni ?? "").trim();
        const okEgb = String(expEgb).trim() === String(bEgb ?? "").trim();
        const okBac = String(expBac).trim() === String(bBac ?? "").trim();

        const allOk = okIni && okEgb && okBac;

        checks.push({
          codigo:"BEN-05",
          descripcion:"Beneficiarios coinciden con referencia",
          status: allOk ? "OK":"FAIL",
          detalle:`Ref: Inicial:${expIni} • EGB:${expEgb} • Bach:${expBac}`
        });
      }
    }
  } else if(!state.refXlsx){
    checks.push({
      codigo:"BEN-03",
      descripcion:"Validación contra referencia",
      status:"WARN",
      detalle:"No se cargó el Excel de referencia; se valida solo presencia de datos."
    });
  }

  // Distancias (heurística)
  try{
    const rutaSheetName = rutaName || baseSheetName;
    const rutaSheet = wb.Sheets[rutaSheetName];
    const data = XLSX.utils.sheet_to_json(rutaSheet, {defval:""});
    if(data.length){
      const keys = Object.keys(data[0]);
      const distKey = keys.find(k => /dist/i.test(k) && /km|kil/i.test(k))
                  || keys.find(k => /dist/i.test(k))
                  || keys.find(k => /km/i.test(k));
      if(distKey){
        const dists = data.map(r=>asNumber(r[distKey])).filter(v=>v!==null);
        if(dists.length){
          const bad = dists.filter(v=>v < 2.5);
          checks.push({
            codigo:"RUT-01",
            descripcion:"Distancias mínimas (>= 2.5 km) si aplica",
            status: bad.length ? "FAIL":"OK",
            detalle: bad.length ? `Se detectaron ${bad.length} registros con ${distKey} < 2.5 (ej: ${bad.slice(0,3).join(", ")}).` : `Campo detectado: ${distKey}.`
          });
        } else {
          checks.push({codigo:"RUT-01", descripcion:"Distancias mínimas (>= 2.5 km) si aplica", status:"WARN", detalle:`Se detectó columna ${distKey}, pero sin valores numéricos.`});
        }
      } else {
        checks.push({codigo:"RUT-01", descripcion:"Distancias mínimas (>= 2.5 km) si aplica", status:"WARN", detalle:"No se detectó una columna de distancia (heurística)."});
      }
    } else {
      checks.push({codigo:"RUT-00", descripcion:"Hoja RUTA con registros", status:"WARN", detalle:"No se detectaron filas en RUTA (sheet_to_json vacío)."});
    }
  }catch(e){
    checks.push({codigo:"RUT-99", descripcion:"Lectura de distancias", status:"WARN", detalle:"No se pudo evaluar distancias automáticamente."});
  }

  return {checks, amie};
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
  if(!state.validation) return;

  const checks = state.validation.checks;
  const status = overallStatus(checks);
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
  saveAs(out, `Respuesta_${data.memo_nro.replace(/[^\w\-]+/g,"_")}.docx`);
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

  // habilitar generar si al menos hay PDF y Excel (aunque haya observaciones igual se genera)
  const hasCore = !!state.modelXlsx && !!state.memoPdf;
  $("btnGenerate").disabled = !hasCore;

  // autocompletar "Para" / "De" si no están (placeholder)
  if(!$("fPara").value) $("fPara").value = "Coordinación Zonal / Dirección Distrital";
  if(!$("fDe").value) $("fDe").value = "Dirección / Analista responsable";
  if(!$("fFirma").value) $("fFirma").value = "________________________";
});

$("btnGenerate").addEventListener("click", generateDocx);
$("btnReset").addEventListener("click", resetAll);

// inicio
resetAll();
