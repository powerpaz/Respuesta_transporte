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
  cuadrosXlsx: null,
  tplDocx: null,
  baseHistXlsx: null,
  memoText: "",
  validation: null,
  modelData: null,
  cumplimiento: null,
  historico: null,
  comparativo: null,
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


/** Construye un objeto modelData desde el XLSM/XLSX de modelamiento (SheetJS workbook) */
function normAmie(v){
  return (v||"").toString().trim().toUpperCase();
}

function fmtMoney(v){
  if(v===null || v===undefined || isNaN(v)) return "N/A";
  try{ return new Intl.NumberFormat("es-EC", {minimumFractionDigits:2, maximumFractionDigits:2}).format(Number(v)); } catch(e){ return String(v); }
}

function buildModelData(wb){
  const rutaName = findSheetByName(wb, "RUTA");
  if(!rutaName) return { error: "No se encontró la hoja RUTA.", modelData:null };

  const shRuta = wb.Sheets[rutaName];
  const distrito = normAmie(getCell(shRuta, "B3"));
  const amieEje  = normAmie(getCell(shRuta, "B4"));
  const x = getNumber(getCell(shRuta, "H6"));
  const y = getNumber(getCell(shRuta, "H7"));

  // Tabla RUTA desde fila 10: A=rutaId, B=AMIE eje, C=AMIE AFC, D=Total, E=Estado, F/G/H = Inicial/EGB/Bach
  const porRuta = [];
  let totalPlanificado = 0;
  let rowIssues = 0;

  for(let r=10; r<2000; r++){
    const rutaId = getCell(shRuta, `A${r}`);
    if(rutaId === null || rutaId === undefined || rutaId === "") break;

    const amieAfc = normAmie(getCell(shRuta, `C${r}`));
    const dTot = getNumber(getCell(shRuta, `D${r}`)) ?? 0;
    const fIni = getNumber(getCell(shRuta, `F${r}`)) ?? 0;
    const gEgb = getNumber(getCell(shRuta, `G${r}`)) ?? 0;
    const hBac = getNumber(getCell(shRuta, `H${r}`)) ?? 0;

    const sum = fIni + gEgb + hBac;
    const okDesagregacion = Math.abs(sum - dTot) <= 0.5;
    if(!okDesagregacion) rowIssues++;

    porRuta.push({
      fila: r,
      rutaId,
      amieAfc: amieAfc || null,
      total: dTot,
      inicial: fIni,
      egb: gEgb,
      bach: hBac,
      okDesagregacion
    });

    totalPlanificado += dTot;
  }

  
  // DIMENSION (costos/presupuesto)
  const dimName = findSheetByName(wb, "DIMENSION");
  let presupuestoAnual = null;
  let costoUnitarioProm = null;
  let presupuestoPorRuta = {}; // rutaId -> suma PRES_A
  if(dimName){
    const shD = wb.Sheets[dimName];
    let sumAnual = 0;
    let sumCostU = 0;
    let cntCostU = 0;

    for(let r=15; r<5000; r++){
      const rutaId = getCell(shD, `A${r}`);
      if(rutaId === null || rutaId === undefined || rutaId === "") break;

      const costU = getNumber(getCell(shD, `G${r}`));
      if(costU !== null){ sumCostU += costU; cntCostU++; }

      const presA = getNumber(getCell(shD, `W${r}`));
      if(presA !== null){
        sumAnual += presA;
        const rid = (rutaId||"").toString().trim();
        if(rid){ presupuestoPorRuta[rid] = (presupuestoPorRuta[rid]||0) + presA; }
      }
    }

    if(sumAnual>0) presupuestoAnual = sumAnual;
    if(cntCostU>0) costoUnitarioProm = sumCostU/cntCostU;
  }
// Fusiones (opcional): AMIE EJE + AMIE AFC
  const fusName = findSheetByName(wb, "Fusiones") || findSheetByName(wb, "FUSIONES");
  let afcList = [];
  if(fusName){
    const shF = wb.Sheets[fusName];
    const matrix = XLSX.utils.sheet_to_json(shF, {header:1, defval:null});
    let headerRow = -1;
    let idxEje=-1, idxAfc=-1;

    for(let i=0;i<matrix.length;i++){
      const row = matrix[i] || [];
      const norm = row.map(v => normName(v));
      const ejeIdx = norm.indexOf("AMIE EJE");
      const afcIdx = norm.indexOf("AMIE AFC");
      if(ejeIdx>=0 && afcIdx>=0){
        headerRow=i; idxEje=ejeIdx; idxAfc=afcIdx; break;
      }
    }
    if(headerRow>=0 && amieEje){
      for(let i=headerRow+1;i<matrix.length;i++){
        const row = matrix[i] || [];
        const eje = normAmie(row[idxEje]);
        const afc = normAmie(row[idxAfc]);
        if(!eje || !afc) continue;
        if(eje === amieEje) afcList.push(afc);
      }
      // fallback: si la hoja está enorme y trae repetidos
      afcList = Array.from(new Set(afcList));
    }
  }

  // Fallback: si no hay Fusiones, usar AMIE AFC detectados en la tabla RUTA
  if(afcList.length===0){
    const fromRuta = porRuta.map(r => r.amieAfc).filter(Boolean);
    afcList = Array.from(new Set(fromRuta));
  }

  const validacionModelamiento = {
    rutasDetectadas: porRuta.length,
    reglaDvsFGH: { ok: rowIssues===0, okCount: porRuta.length - rowIssues, failCount: rowIssues },
    coordenadasOk: (x!==null && y!==null),
    hojasObligatoriasOk: true, // esto lo confirma validateAll
    estadoGlobal: (rowIssues===0 ? "OK" : "WARN")
  };

  const modelData = {
    ids: { distrito, amieEje },
    ubicacion: { utm: { x, y } },
    beneficiarios: { totalPlanificado, porRuta },
    fusiones: { totalAfc: afcList.length, afcList },
    costos: { presupuestoAnual, costoUnitarioProm, presupuestoPorRuta },
    validacionModelamiento
  };

  return { error:null, modelData };
}

function extractJustificacionFromMemo(text){
  const t = (text||"").replace(/\s+/g," ").trim();
  if(!t) return "";
  // busca un tramo típico de justificación (sin ponerse exquisito)
  const m = t.match(/(Se\s+consider[óo][^\.]{0,240}\.)(?!\S)/i) || t.match(/(Inicial\s*\(\s*3\s*años\s*\)[^\.]{0,240}\.)(?!\S)/i);
  if(m) return m[1].trim();
  // fallback: frase corta alrededor de "Inicial"
  const k = t.toLowerCase().indexOf("inicial");
  if(k>=0){
    return t.slice(Math.max(0,k-120), Math.min(t.length,k+180)).trim();
  }
  return "";
}

function buildCumplimientoCriterios(modelData, daeiRanges, memoText){
  if(!modelData || !daeiRanges) return null;
  const amie = modelData.ids.amieEje;
  const r = daeiRanges[amie];
  if(!r) return { cumple:null, etiqueta:"SIN DATO DAEI", detalle:`No se encontró AMIE ${amie} en la referencia DAEI.` };

  const plan = modelData.beneficiarios.totalPlanificado;
  const diffMax = (r.max!=null) ? (plan - r.max) : null;
  const just = extractJustificacionFromMemo(memoText);

  if(r.min!=null && r.max!=null && plan>=r.min && plan<=r.max){
    return {
      cumple:true,
      etiqueta:"CUMPLE",
      detalle:`Planificado ${plan} dentro del rango DAEI [${r.min}–${r.max}].`,
      justificacion:""
    };
  }

  if(r.max!=null && plan>r.max){
    return {
      cumple:false,
      etiqueta:"NO CUMPLE",
      detalle:`Planificado ${plan} > Máximo_F ${r.max} (dif. +${diffMax}).`,
      justificacion: just || ""
    };
  }

  if(r.min!=null && plan<r.min){
    return {
      cumple:false,
      etiqueta:"NO CUMPLE",
      detalle:`Planificado ${plan} < Mínimo_F ${r.min}.`,
      justificacion:""
    };
  }

  return { cumple:null, etiqueta:"REVISAR", detalle:"No se pudo evaluar el rango DAEI.", justificacion: just || "" };
}

function buildQuipuxTables(modelData, cumplimiento, hasMemo, comparativo){
  const distrito = modelData.ids.distrito;
  const amie_eje = modelData.ids.amieEje;
  const informe = hasMemo ? "X" : "";
  const listado = (comparativo && comparativo.nuevaRuta) ? "SI" : "N/A";
  const rev_doc_rows = [{ distrito, amie_eje, informe, listado }];
  const criteriosTxt = cumplimiento ? (cumplimiento.cumple ? "Si" : "No") : "N/A";
  const valTxt = "Si";
  const val_rows = modelData.fusiones.afcList.map(afc => ({ distrito, amie_eje, amie_afc: afc, beneficiarios: modelData.beneficiarios.totalPlanificado, criterios: criteriosTxt, validacion: valTxt }));
  return { rev_doc_rows, val_rows };
}


function fillCuadrosWorkbook(cwb, modelData, cumplimiento){
  const wsValid = cwb.Sheets["4. VALIDACIÓN"] || cwb.Sheets["4. VALIDACION"];
  const wsRev   = cwb.Sheets["3. REVISIÓN Y DOCUMENTACIÓN"] || cwb.Sheets["3. REVISION Y DOCUMENTACION"];
  if(!wsValid || !wsRev){
    return {error:"El archivo de Cuadros no contiene las hojas esperadas: '3. REVISIÓN Y DOCUMENTACIÓN' y '4. VALIDACIÓN'."};
  }

  const {rev_doc_rows, val_rows} = buildQuipuxTables(modelData, cumplimiento, true, state.comparativo);

  // --- Hoja 4. VALIDACIÓN ---
  const startRow = 2; // cabecera está en fila 1
  // limpiar rango previo (A2:F200) por seguridad
  for(let r=startRow; r<200; r++){
    for(const col of ["A","B","C","D","E","F"]){
      const addr = `${col}${r}`;
      if(wsValid[addr]) delete wsValid[addr];
    }
  }

  val_rows.forEach((row,i)=>{
    const r = startRow + i;
    wsValid[`A${r}`] = { t:"s", v: row.distrito };
    wsValid[`B${r}`] = { t:"s", v: row.amie_eje };
    wsValid[`C${r}`] = { t:"s", v: row.amie_afc };
    wsValid[`D${r}`] = { t:"n", v: Number(row.beneficiarios)||0 };
    wsValid[`E${r}`] = { t:"s", v: row.criterios };
    wsValid[`F${r}`] = { t:"s", v: row.validacion };
  });

  // merges para que se vea como el formato Quipux (opcional pero recomendado)
  const endRow = startRow + val_rows.length - 1;
  wsValid["!merges"] = wsValid["!merges"] || [];
  // reset merges en el área A-F por si el template trae otros merges
  wsValid["!merges"] = wsValid["!merges"].filter(m => !(m.s.r>=1 && m.s.r<=endRow-1 && m.s.c<=5)); // conservador

  if(val_rows.length>1){
    const mrg = (c)=>({s:{r:startRow-1,c}, e:{r:endRow-1,c}});
    // A,B,D,E,F (0,1,3,4,5)
    wsValid["!merges"].push(mrg(0), mrg(1), mrg(3), mrg(4), mrg(5));
  }

  // actualizar rango !ref
  const last = startRow + Math.max(0,val_rows.length-1);
  wsValid["!ref"] = `A1:F${Math.max(1,last)}`;

  // --- Hoja 3. REVISIÓN Y DOCUMENTACIÓN ---
  // limpiar A2:D100
  for(let r=2;r<100;r++){
    for(const col of ["A","B","C","D"]){
      const addr=`${col}${r}`;
      if(wsRev[addr]) delete wsRev[addr];
    }
  }
  const rr = rev_doc_rows[0];
  wsRev["A2"] = {t:"s", v: rr.distrito};
  wsRev["B2"] = {t:"s", v: rr.amie_eje};
  wsRev["C2"] = {t:"s", v: rr.informe};
  wsRev["D2"] = {t:"s", v: rr.listado};
  wsRev["!ref"] = "A1:D2";

  return {error:null};
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


function parseHistoricoBase(baseWb){
  const sheetName = baseWb.SheetNames[0];
  const sh = baseWb.Sheets[sheetName];
  const matrix = XLSX.utils.sheet_to_json(sh, {header:1, defval:null});
  let headerRow=-1;
  let idxDistrito=-1, idxAmie=-1, idxBenef=-1, idxPresA=-1, idxAnio=-1;
  for(let i=0;i<matrix.length;i++){
    const row = matrix[i]||[];
    const norm = row.map(v=>normName(v));
    if(norm.includes("DISTRITO") && norm.includes("AMIE EJE")){
      headerRow=i;
      idxDistrito = norm.indexOf("DISTRITO");
      idxAmie = norm.indexOf("AMIE EJE");
      idxBenef = norm.indexOf("BENEFICIARIOS");
      idxPresA = norm.indexOf("PRESPUESTO ANUAL");
      if(idxPresA<0) idxPresA = norm.indexOf("PRESUPUESTO ANUAL");
      idxAnio = norm.indexOf("AÑO LECTIVO VALIDADO");
      if(idxAnio<0) idxAnio = norm.indexOf("ANO LECTIVO VALIDADO");
      if(idxAnio<0) idxAnio = norm.indexOf("AÑO");
      break;
    }
  }
  if(headerRow<0 || idxDistrito<0 || idxAmie<0){
    return {byKey:null, error:"No se encontró encabezado con 'Distrito' y 'AMIE EJE' en la base histórica."};
  }
  const byKey = {};
  for(let i=headerRow+1;i<matrix.length;i++){
    const row = matrix[i]||[];
    const distrito = normAmie(row[idxDistrito]);
    const amieEje = normAmie(row[idxAmie]);
    if(!distrito || !amieEje) continue;
    const key = `${distrito}__${amieEje}`;
    const benef = (idxBenef>=0 ? (getNumber(row[idxBenef]) ?? 0) : 0);
    const presA = (idxPresA>=0 ? (getNumber(row[idxPresA]) ?? 0) : 0);
    const anio = (idxAnio>=0 ? (row[idxAnio]||"").toString().trim() : "");
    if(!byKey[key]) byKey[key] = { distrito, amieEje, totalBenef:0, totalPresA:0, rows:0, anioRef:anio };
    byKey[key].totalBenef += benef;
    byKey[key].totalPresA += presA;
    byKey[key].rows += 1;
    if(anio) byKey[key].anioRef = anio;
  }
  return {byKey, error:null};
}

function buildComparativo(modelData, historicoByKey){
  if(!modelData || !historicoByKey) return null;
  const key = `${modelData.ids.distrito}__${modelData.ids.amieEje}`;
  const h = historicoByKey[key];
  if(!h) return { tieneHistorico:false };
  const actualPresA = (modelData.costos && modelData.costos.presupuestoAnual) ? modelData.costos.presupuestoAnual : null;
  const anteriorPresA = h.totalPresA || null;
  const actualBenef = modelData.beneficiarios.totalPlanificado;
  const anteriorBenef = h.totalBenef || null;
  const actualRutas = modelData.validacionModelamiento.rutasDetectadas;
  const anteriorRutas = h.rows || null;
  const nuevaRuta = (anteriorRutas!==null && actualRutas>anteriorRutas);
  let variacionPresPct = null;
  if(actualPresA!==null && anteriorPresA!==null && anteriorPresA>0){
    variacionPresPct = ((actualPresA-anteriorPresA)/anteriorPresA)*100;
  }
  return {tieneHistorico:true, anioRef:(h.anioRef||""), anterior:{rutas:anteriorRutas, beneficiarios:anteriorBenef, presupuestoAnual:anteriorPresA}, actual:{rutas:actualRutas, beneficiarios:actualBenef, presupuestoAnual:actualPresA}, nuevaRuta, variacionPresPct};
}

function buildComparativoTexto(comp){
  if(!comp || !comp.tieneHistorico) return "";
  const presAAct = fmtMoney(comp.actual.presupuestoAnual);
  const presAAnt = fmtMoney(comp.anterior.presupuestoAnual);
  const pct = (comp.variacionPresPct===null || comp.variacionPresPct===undefined) ? "" : ` (${comp.variacionPresPct.toFixed(2)}%)`;
  const nr = comp.nuevaRuta ? " Se identifican nuevas rutas respecto al histórico." : "";
  return `Comparativo ${comp.anioRef || "histórico"}: rutas ${comp.anterior.rutas}→${comp.actual.rutas}; beneficiarios ${comp.anterior.beneficiarios}→${comp.actual.beneficiarios}; presupuesto anual ${presAAnt}→${presAAct}${pct}.${nr}`;
}

function buildNotasTecnicas(comp, cumplimiento){
  const notas = [];
  if(cumplimiento && !cumplimiento.cumple){
    notas.push(`Beneficiarios planificados superan el rango DAEI (${cumplimiento.detalle}). Se requiere justificación en el informe suscrito por el nivel desconcentrado.`);
    if(cumplimiento.justificacion) notas.push(`Justificación: ${cumplimiento.justificacion}`);
  }
  if(comp && comp.nuevaRuta){
    notas.push("Al existir nuevas rutas/actualización, corresponde adjuntar el listado de beneficiarios (según lineamientos)." );
  }
  return notas.join("\n");
}


function parseCuadrosQuipux(wb){
  try{
    const sh1 = wb.Sheets["Formato cuadro 1"] || wb.Sheets["Formato Cuadro 1"];
    const sh2 = wb.Sheets["Formato Cuadro 2"] || wb.Sheets["Formato cuadro 2"];
    if(!sh1 || !sh2){
      return { error: "No se encontraron las hojas 'Formato cuadro 1' y/o 'Formato Cuadro 2' en el archivo de Cuadros (Quipux)." };
    }

    // --- Cuadro 1: Revisión de documentación (A:D) ---
    const a1 = XLSX.utils.sheet_to_json(sh1, {header:1, blankrows:false});
    if(!a1 || a1.length<2){ return { error: "El Cuadro 1 (Formato cuadro 1) no contiene filas de datos." }; }
    const hdr1 = (a1[0]||[]).slice(0,4).map(v => (v||"").toString().trim().toLowerCase());
    // Validación mínima de encabezados
    if(!(hdr1[0]||"").includes("distrito") || !(hdr1[1]||"").includes("amie")){
      return { error: "El Cuadro 1 no tiene el encabezado esperado (Distrito / AMIE UE Eje / ...)." };
    }
    let lastDistrito = "";
    let lastAmieEje = "";
    const rev_doc_rows = [];
    for(let i=1;i<a1.length;i++){
      const row = a1[i] || [];
      const distrito = (row[0]??"").toString().trim() || lastDistrito;
      const amie_eje = (row[1]??"").toString().trim() || lastAmieEje;
      const informe = (row[2]??"").toString().trim();
      const listado = (row[3]??"").toString().trim();
      if(!distrito && !amie_eje && !informe && !listado) continue;
      if(distrito) lastDistrito = distrito;
      if(amie_eje) lastAmieEje = amie_eje;
      rev_doc_rows.push({ distrito, amie_eje, informe, listado });
    }
    if(rev_doc_rows.length===0){ return { error: "El Cuadro 1 no contiene filas válidas para integrar en la respuesta." }; }

    // --- Cuadro 2: Validación de modelamientos (A:F) ---
    const a2 = XLSX.utils.sheet_to_json(sh2, {header:1, blankrows:false});
    if(!a2 || a2.length<2){ return { error: "El Cuadro 2 (Formato Cuadro 2) no contiene filas de datos." }; }
    const hdr2 = (a2[0]||[]).slice(0,6).map(v => (v||"").toString().trim().toLowerCase());
    if(!(hdr2[0]||"").includes("distrito") || !(hdr2[1]||"").includes("amie")){
      return { error: "El Cuadro 2 no tiene el encabezado esperado (Distrito / AMIE UE Eje / ...)." };
    }
    lastDistrito = "";
    lastAmieEje = "";
    const val_rows = [];
    for(let i=1;i<a2.length;i++){
      const row = a2[i] || [];
      const distrito = (row[0]??"").toString().trim() || lastDistrito;
      const amie_eje = (row[1]??"").toString().trim() || lastAmieEje;
      const amie_fc = (row[2]??"").toString().trim();
      const benef_est = (row[3]??"").toString().trim();
      const cumple = (row[4]??"").toString().trim();
      const validacion = (row[5]??"").toString().trim();
      if(!distrito && !amie_eje && !amie_fc && !benef_est && !cumple && !validacion) continue;
      if(distrito) lastDistrito = distrito;
      if(amie_eje) lastAmieEje = amie_eje;
      // Filtra filas "en blanco" (por ejemplo solo continuidad sin amie_fc)
      if(!amie_fc && !benef_est && !cumple && !validacion) continue;
      val_rows.push({ distrito, amie_eje, amie_fc, benef_est, cumple, validacion });
    }
    const modelamientos_total = val_rows.length;
    const distrito_principal = (val_rows[0]?.distrito || rev_doc_rows[0]?.distrito || "").toString();

    return { rev_doc_rows, val_rows, modelamientos_total, distrito_principal };
  } catch(e){
    return { error: "Error al leer Cuadros (Quipux): " + (e?.message || e) };
  }
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

  
  // Referencia DAEI (opcional)
  if(!state.refXlsx){
    checks.push({codigo:"REF-01", descripcion:"Referencia Beneficiarios (DAEI) cargada", status:"WARN", detalle:"No se cargó la referencia DAEI (recomendado para validar rangos)."});
  } else {
    checks.push({codigo:"REF-01", descripcion:"Referencia Beneficiarios (DAEI) cargada", status:"OK", detalle: state.refXlsx.name});
  }

  // Cuadros (Quipux) (opcional pero para heredar tablas)
  if(!state.cuadrosXlsx){
    checks.push({codigo:"CQX-01", descripcion:"Cuadros (Quipux) cargados (tablas para el Word)", status:"WARN", detalle:"No se cargó el Excel de Cuadros (Quipux)."});
  } else {
    checks.push({codigo:"CQX-01", descripcion:"Cuadros (Quipux) cargados (tablas para el Word)", status:"OK", detalle: state.cuadrosXlsx.name});
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

  
// --- Construir modelData (para heredar a Cuadros Quipux) ---
const builtModel = buildModelData(wb);
if(builtModel.error){
  checks.push({codigo:"MDL-01", descripcion:"ModelData: extracción estructurada (Distrito/AMIE/Beneficiarios/Fusiones)", status:"FAIL", detalle: builtModel.error});
  state.modelData = null;
  state.cumplimiento = null;
} else {
  state.modelData = builtModel.modelData;
  // comparativo (si base histórica está cargada)
  if(state.baseHistXlsx){
    try{
      const bbuf = await fileToArrayBuffer(state.baseHistXlsx);
      const bwb = XLSX.read(bbuf, {type:"array"});
      const ph = parseHistoricoBase(bwb);
      if(!ph.error){ state.historico = ph.byKey; state.comparativo = buildComparativo(state.modelData, ph.byKey); }
    } catch(e){}
  }

  checks.push({codigo:"MDL-01", descripcion:"ModelData: extracción estructurada (Distrito/AMIE/Beneficiarios/Fusiones)", status:"OK", detalle:`Distrito ${state.modelData.ids.distrito} • AMIE EJE ${state.modelData.ids.amieEje} • Beneficiarios ${state.modelData.beneficiarios.totalPlanificado.toFixed(0)} • AFC ${state.modelData.fusiones.totalAfc}`});
}

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

  
// ========= 9) Rangos DAEI (cruce) =========
if(state.refXlsx && state.modelData && isFilled(state.modelData.ids.amieEje)){
  const amieMDL = state.modelData.ids.amieEje;
  const totalMDL = state.modelData.beneficiarios.totalPlanificado;

  const refWb = XLSX.read(await readAsArrayBuffer(state.refXlsx), {type:"array"});
  const parsed = parseDAEIRanges(refWb);
  if(parsed.error){
    checks.push({codigo:"RNG-00", descripcion:"Referencia DAEI: estructura válida (AMIE/Mínimo_F/Máximo_F)", status:"FAIL", detalle: parsed.error});
    state.cumplimiento = null;
  } else {
    const rng = parsed.ranges[amieMDL];
    if(!rng){
      checks.push({codigo:"RNG-01", descripcion:"Rangos DAEI para AMIE IE EJE", status:"WARN", detalle:`No se encontró AMIE ${amieMDL} en la referencia DAEI.`});
      state.cumplimiento = { cumple:null, etiqueta:"SIN DATO DAEI", detalle:`No se encontró AMIE ${amieMDL} en la referencia DAEI.`, justificacion:"" };
    } else {
      state.cumplimiento = buildCumplimientoCriterios(state.modelData, parsed.ranges, state.memoText);
      const ok = state.cumplimiento && state.cumplimiento.cumple === true;

      const detBase = `AMIE ${amieMDL} • Total(modelamiento): ${totalMDL.toFixed(0)} • Rango DAEI: ${rng.min}–${rng.max}`;
      const detExtra = state.cumplimiento?.detalle ? ` • ${state.cumplimiento.detalle}` : "";
      checks.push({
        codigo:"RNG-02",
        descripcion:"Beneficiarios del modelamiento vs rango DAEI (Mínimo_F–Máximo_F)",
        status: ok ? "OK" : "WARN",
        detalle: detBase + detExtra
      });
    }
  }
}

// ========= 10) Cuadros Quipux (herencia desde modelamiento) =========
// En tu UX, el Excel de Cuadros viene como formato vacío; se llena a partir del modelamiento + DAEI + Memorando.
state.cuadrosData = null;

if(state.modelData){
  const builtTables = buildQuipuxTables(state.modelData, state.cumplimiento, !!state.memoPdf);
  state.cuadrosData = builtTables;

  // Si el usuario cargó el archivo de Cuadros, validamos que existan las hojas esperadas
  if(state.cuadrosXlsx){
    const cwb = XLSX.read(await readAsArrayBuffer(state.cuadrosXlsx), {type:"array"});
    const hasValid = !!(cwb.Sheets["4. VALIDACIÓN"] || cwb.Sheets["4. VALIDACION"]);
    const hasRev   = !!(cwb.Sheets["3. REVISIÓN Y DOCUMENTACIÓN"] || cwb.Sheets["3. REVISION Y DOCUMENTACION"]);
    checks.push({
      codigo:"CQX-02",
      descripcion:"Cuadros (Quipux): formato esperado (3. REVISIÓN… y 4. VALIDACIÓN)",
      status: (hasValid && hasRev) ? "OK" : "WARN",
      detalle: (hasValid && hasRev) ? "Hojas detectadas en el formato Quipux." : "No se detectaron una o ambas hojas en el archivo de Cuadros; igual se integrarán tablas al Word."
    });
  } else {
    checks.push({codigo:"CQX-02", descripcion:"Cuadros (Quipux): tablas generadas desde el modelamiento", status:"OK", detalle:`Revisión filas: ${builtTables.rev_doc_rows.length} • Validación filas: ${builtTables.val_rows.length}`});
  }
}

// Guardar meta útil para el Word
state.validationMeta = {
  amieEje: state.modelData?.ids?.amieEje || amieEje || "",
  rutas_total: state.modelData?.beneficiarios?.porRuta?.length || br.rows || 0,
  beneficiarios_total: state.modelData?.beneficiarios?.totalPlanificado || br.total || 0
};

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
  // Nota: el checklist usa la propiedad `status` (OK/WARN/FAIL)
  const hasFail = checks.some(c => c.status === "FAIL");
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

  const meta = state.validationMeta || {};
  const cqx = state.cuadrosData || {rev_doc_rows:[], val_rows:[], modelamientos_total:0, distrito_principal:""};

  const data = {
    memo_nro: $("fMemoNro").value || "(s/n)",
    memo_fecha: $("fMemoFecha").value || "(s/f)",
    para: $("fPara").value || "(s/d)",
    de: $("fDe").value || "(s/d)",
    asunto: $("fAsunto").value || "(s/a)",
    resultado_general,
    resumen_texto: resumen,
    conclusion,
    firma: $("fFirma").value || "(falta firma)",

    // Herencia de Cuadros (Quipux)
    rev_doc_rows: cqx.rev_doc_rows,
    val_rows: cqx.val_rows,

    // Meta para textos dinámicos
    modelamientos_total: cqx.modelamientos_total || 0,
    distrito_principal: cqx.distrito_principal || "",
    rutas_total: meta.rutas_total || 0,
    anio_lectivo: $("fAnioLectivo") ? ($("fAnioLectivo").value || "2026-2027") : "2026-2027"
  };

  const content = await loadTemplateBytes();
  const zip = new PizZip(content);
  const doc = new window.docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
  doc.render(data);

  const out = doc.getZip().generate({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
  const fname = `Respuesta_${data.memo_nro.replace(/[^\w\-]+/g,"_")}.docx`;
  saveAs(out, fname);

// Si se cargó el archivo de Cuadros (Quipux), lo llenamos y lo descargamos también
if(state.cuadrosXlsx && state.modelData){
  try{
    const cwb = XLSX.read(await readAsArrayBuffer(state.cuadrosXlsx), {type:"array"});
    const filled = fillCuadrosWorkbook(cwb, state.modelData, state.cumplimiento);
    if(!filled.error){
      const wbout = XLSX.write(cwb, {bookType:"xlsx", type:"array"});
      const blob = new Blob([wbout], {type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
      const xname = `Cuadros_Quipux_Llenos_${data.memo_nro.replace(/[^\w\-]+/g,"_")}.xlsx`;
      saveAs(blob, xname);
    }
  } catch(e){
    console.warn("No se pudo generar el Excel de Cuadros llenos:", e);
  }
}


  const inherited = [];
  inherited.push(`Memorando (PDF): ${state.memoPdf.name}`);
  inherited.push(`Modelamiento (Excel): ${state.modelXlsx.name}`);
  inherited.push(state.refXlsx ? `Referencia Beneficiarios (Excel): ${state.refXlsx.name}` : "Referencia Beneficiarios (Excel): (no cargado)");
  inherited.push(state.cuadrosXlsx ? `Cuadros (Quipux) (Excel): ${state.cuadrosXlsx.name}` : "Cuadros (Quipux) (Excel): (no cargado)");
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
  state.cuadrosXlsx = null;
  state.tplDocx = null;
  state.memoText = "";
  state.validation = null;

  $("memoFile").value="";
  $("modelFile").value="";
  $("refFile").value="";
  $("cuadrosFile").value="";
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
bindDrop("dropBase", "baseFile", "baseMeta", async (file)=>{ state.baseHistXlsx = file; });

bindDrop("dropCuadros", "cuadrosFile", "cuadrosMeta", async (file)=>{ state.cuadrosXlsx = file; });
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
  const hasFail = checks.some(c => c.status === "FAIL");
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
