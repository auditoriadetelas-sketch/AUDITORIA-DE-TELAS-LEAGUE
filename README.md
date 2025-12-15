<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width,initial-scale=1" />
  <title>Auditoría de Tela</title>

  <style>
    body{font-family:Arial, sans-serif;background:#f0f3f9;padding:18px;color:#222}
    h2{margin-top:20px;font-size:1.1rem;color:#0b63d1}
    table{width:100%;border-collapse:collapse;margin-top:10px;font-size:0.9rem;background:#ffffff}
    th,td{border:1px solid #ddd;padding:6px;text-align:center}
    th{background:#e6f0ff;color:#222;font-weight:bold}
    input,textarea,select{
      width:100%;border:1px solid #ccc;border-radius:4px;padding:4px;font-size:0.9rem;background:#fff
    }
    textarea{resize:vertical;min-height:40px}
    button{
      padding:8px 12px;border:none;border-radius:4px;
      cursor:pointer;font-weight:bold;margin-right:5px;transition:background 0.3s
    }
    .add{background:#0b63d1;color:white;margin-top:10px}
    .add:hover{background:#084c9f}
    .del{background:#dc3545;color:white}
    .del:hover{background:#c82333}
    .totales{margin-top:15px;display:flex;gap:15px;flex-wrap:wrap}
    .box{
      background:#ffffff;border:1px solid #ccc;border-radius:6px;
      padding:8px 12px;box-shadow:0 1px 3px rgba(0,0,0,0.05);
    }
    label{font-weight:bold;font-size:0.9rem;display:block;margin-top:6px;color:#333}
    .form-header{
      display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));
      gap:15px;margin-bottom:20px;background:#ffffff;padding:15px;
      border:1px solid #ccc;border-radius:6px;box-shadow:0 1px 3px rgba(0,0,0,0.05);
    }
  </style>

  <script src="https://cdn.sheetjs.com/xlsx-latest/package/dist/xlsx.full.min.js"></script>
</head>

<body>

<h2>📊 Información de Auditoría</h2>

<div class="form-header">
  <div><label>Lote:<input type="text" id="loteGlobal"></label></div>
  <div><label>Part Number:<input type="text" id="partNumber"></label></div>
  <div><label>Part Color:<input type="text" id="partColor"></label></div>

  <div>
    <label>Proveedor:
      <select id="proveedor">
        <option value="">Seleccione</option>
        <option>INNOV TEXTILES</option>
        <option>GAMATEX</option>
        <option>TEXPASA</option>
        <option>RAYONES</option>
      </select>
    </label>
  </div>

  <div>
    <label>Auditor:
      <select id="auditor">
        <option value="">Seleccione</option>
        <option>René Rodríguez</option>
        <option>Ezequiel Rivera</option>
        <option>Felipe Betancourt</option>
      </select>
    </label>
  </div>

  <div><label>Total Rollos:<input type="text" id="totalRollos"></label></div>
  <div><label>Total Yardas:<input type="text" id="totalYardasGlobal"></label></div>

  <div>
    <label>Resultados de auditoría:
      <select id="resultado">
        <option value="">Seleccione</option>
        <option>Aprobado</option>
        <option>Rechazado</option>
        <option>Certificar</option>
      </select>
    </label>
  </div>
</div>

<h2>🧵 Rollos</h2>

<table id="tabla">
  <thead>
    <tr>
      <th># Rollo</th>
      <th>Yardas Label</th>
      <th>Yardas Reales</th>
      <th>Ancho Estándar</th>
      <th>Ancho Real</th>
      <th>Yds Revisadas</th>
      <th>Puntos por Rollo</th>
      <th>Rate por Rollo</th>
      <th>Peso Estándar</th>
      <th>Peso Real</th>
      <th>% (+/-) Peso</th>
      <th>% Yardas Faltantes</th>
      <th>Obs.</th>
      <th>Acción</th>
    </tr>
  </thead>
  <tbody id="tbody"></tbody>
</table>

<button class="add" id="addRow">➕ Añadir fila</button>
<button class="add" id="exportXLSX">💾 Exportar a XLSX</button>

<div class="totales">
  <div class="box">Total Yardas Reales: <span id="tYardas">0</span></div>
  <div class="box">Puntos Penalizados Global: <span id="tPuntos">0</span></div>
  <div class="box">Rate Penalización Global: <span id="tRate">0</span></div>
</div>

<h2>🚫 Defectos por Rollo</h2>

<table id="tablaDefectos">
  <thead>
    <tr>
      <th>Rollo</th>
      <th>Cantidad Defectos</th>
      <th>Puntos (1–4)</th>
      <th>Código Defecto</th>
      <th>Acción</th>
    </tr>
  </thead>
  <tbody id="tbodyDef"></tbody>
</table>

<button class="add" id="addDefecto">➕ Añadir defecto</button>

<h2>📏 Anchos por Rollo</h2>

<table id="tablaAnchos">
  <thead>
    <tr>
      <th>Rollo</th>
      <th>Ancho 1</th>
      <th>Ancho 2</th>
      <th>Ancho 3</th>
      <th>Acción</th>
    </tr>
  </thead>
  <tbody id="tbodyAnchos"></tbody>
</table>

<button class="add" id="addAncho">➕ Añadir Ancho</button>

<h2>🧪 Pruebas Adicionales</h2>

<div class="form-header">
  <div><label>Match:<select id="match"><option value="">-</option><option>Aceptable</option><option>Rechazado</option></select></label></div>
  <div><label>Stretch:<select id="stretch"><option value="">-</option><option>Aceptable</option><option>Rechazado</option></select></label></div>
  <div><label>Handfeel:<select id="handfeel"><option value="">-</option><option>Aceptable</option><option>Rechazado</option></select></label></div>
  <div><label>Pilling:<select id="pilling"><option value="">-</option><option>Aceptable</option><option>Rechazado</option></select></label></div>
  <div><label>Brushing:<select id="brushing"><option value="">-</option><option>Aceptable</option><option>Rechazado</option></select></label></div>
</div>

<script>
/* ------------------------
    TABLA DE ROLLOS
-------------------------*/
const tbody=document.getElementById("tbody");
let idx=0;

function addRow(data={}) {
  idx++;
  const tr=document.createElement("tr");
  tr.dataset.rollo=idx;

  tr.innerHTML = `
    <td><input class="rolloNum" value="${data.rollo||idx}"></td>
    <td><input type="number" class="label" value="${data.label||''}"></td>
    <td><input type="number" class="real" value="${data.real||''}"></td>
    <td><input type="number" class="anchoStd" value="${data.anchoStd||''}"></td>
    <td><input type="number" class="anchoReal" value="${data.anchoReal||''}"></td>
    <td><input type="number" class="revisadas" readonly></td>
    <td><input type="number" class="puntosRollo" readonly value="0"></td>
    <td><input type="number" class="rateRollo" readonly value="0"></td>
    <td><input type="number" class="pesoStd" value="${data.pesoStd||''}"></td>
    <td><input type="number" class="pesoReal" value="${data.pesoReal||''}"></td>
    <td><input class="pct" readonly></td>
    <td><input class="faltantes" readonly></td>
    <td><textarea>${data.obs||''}</textarea></td>
    <td><button class="del">Eliminar</button></td>
  `;

  tbody.appendChild(tr);

  tr.querySelectorAll(".real, .label, .pesoStd, .pesoReal, .anchoReal, .anchoStd")
    .forEach(inp=>inp.addEventListener("input",()=>calcRow(tr)));

  tr.querySelector(".del").addEventListener("click",()=>{
    tr.remove();
    calcTotals();
    updateDefectos();
  });

  calcRow(tr);
}

function calcRow(tr){
  const real=parseFloat(tr.querySelector(".real").value)||0;
  const label=parseFloat(tr.querySelector(".label").value)||0;
  const pesoStd=parseFloat(tr.querySelector(".pesoStd").value)||0;
  const pesoReal=parseFloat(tr.querySelector(".pesoReal").value)||0;
  const anchoReal=parseFloat(tr.querySelector(".anchoReal").value)||1;
  const puntos=parseFloat(tr.querySelector(".puntosRollo").value)||0;

  tr.querySelector(".revisadas").value=real.toFixed(2);

  tr.querySelector(".faltantes").value =
    label>0 ? (((real/label-1)*100).toFixed(2)+"%") : "0.00%";

  tr.querySelector(".pct").value =
    pesoStd>0 ? (((pesoReal/pesoStd-1)*100).toFixed(2)+"%") : "0.00%";

  tr.querySelector(".rateRollo").value =
    (real>0 && anchoReal>0)
      ? ((puntos*36)/(real*anchoReal)*100).toFixed(2)
      : 0;

  calcTotals();
}

function calcTotals(){
  let yardas=0,puntos=0;
  const filas=[...tbody.querySelectorAll("tr")];

  filas.forEach(tr=>{
    yardas+=parseFloat(tr.querySelector(".real").value)||0;
    puntos+=parseFloat(tr.querySelector(".puntosRollo").value)||0;
  });

  document.getElementById("tYardas").textContent=yardas.toFixed(2);
  document.getElementById("tPuntos").textContent=puntos.toFixed(2);

  let anchoMin=Infinity;
  filas.forEach(r=>{
    const a=parseFloat(r.querySelector(".anchoReal").value);
    if(a && a > 0 && a < anchoMin) anchoMin=a;
  });
  
  const anchoRef = anchoMin < Infinity ? anchoMin : 1;

  let rate = (yardas>0 && anchoRef>0)
    ? (puntos*36)/(yardas*anchoRef)*100 : 0;

  document.getElementById("tRate").textContent=rate.toFixed(2);
}

document.getElementById("addRow").addEventListener("click",()=>addRow());
addRow();

/* ------------------------
    DEFECTOS
-------------------------*/
const tbodyDef=document.getElementById("tbodyDef");

document.getElementById("addDefecto").addEventListener("click",()=>addDefecto());

function addDefecto(data={}) {
  const tr=document.createElement("tr");

  tr.innerHTML = `
    <td><input class="defRollo" value="${data.rollo||''}"></td>
    <td><input type="number" class="defCantidad" value="${data.cantidad||''}"></td>
    <td>
      <select class="defPuntos">
        <option value="1" ${data.puntos==1?'selected':''}>1</option>
        <option value="2" ${data.puntos==2?'selected':''}>2</option>
        <option value="3" ${data.puntos==3?'selected':''}>3</option>
        <option value="4" ${data.puntos==4?'selected':''}>4</option>
      </select>
    </td>
    <td><input value="${data.codigo||''}"></td>
    <td><button class="del">Eliminar</button></td>
  `;

  tbodyDef.appendChild(tr);

  tr.querySelector(".del").addEventListener("click",()=>{
    tr.remove();
    updateDefectos();
  });

  tr.querySelectorAll("input, select").forEach(i=>i.addEventListener("input",updateDefectos));
}

function updateDefectos(){
  tbody.querySelectorAll(".puntosRollo").forEach(i=>i.value=0);

  tbodyDef.querySelectorAll("tr").forEach(tr=>{
    const rollo=tr.querySelector(".defRollo").value.trim();
    const cant=parseFloat(tr.querySelector(".defCantidad").value)||0;
    const pts=parseFloat(tr.querySelector(".defPuntos").value)||1;
    const total=cant*pts;

    if(!rollo) return;

    const fila=[...tbody.querySelectorAll("tr")]
      .find(r=>r.querySelector(".rolloNum").value.trim()==rollo);

    if(fila){
      const puntosInput = fila.querySelector(".puntosRollo");
      puntosInput.value = (parseFloat(puntosInput.value)||0)+total;
    }
  });

  tbody.querySelectorAll("tr").forEach(calcRow);
  calcTotals();
}

/* ------------------------
    ANCHOS
-------------------------*/
const tbodyAnchos=document.getElementById("tbodyAnchos");

document.getElementById("addAncho").addEventListener("click",()=>addAncho());

function addAncho(data={}) {
  const tr=document.createElement("tr");

  tr.innerHTML = `
    <td><input class="anchoRollo" value="${data.rollo||''}"></td>
    <td><input type="number" class="ancho1" value="${data.ancho1||''}"></td>
    <td><input type="number" class="ancho2" value="${data.ancho2||''}"></td>
    <td><input type="number" class="ancho3" value="${data.ancho3||''}"></td>
    <td><button class="del">Eliminar</button></td>
  `;

  tbodyAnchos.appendChild(tr);

  tr.querySelector(".del").addEventListener("click",()=>tr.remove());
}

/* ------------------------
    EXPORTAR XLSX (Modificado a estructura lineal)
-------------------------*/
document.getElementById("exportXLSX").addEventListener("click",()=>{

  const wb=XLSX.utils.book_new();

  // 1. Recolección de datos globales
  const globalData = {
    Auditor: document.getElementById("auditor").value,
    Proveedor: document.getElementById("proveedor").value,
    Lote: document.getElementById("loteGlobal").value,
    PartNumber: document.getElementById("partNumber").value,
    PartColor: document.getElementById("partColor").value,
    TotalRollos: document.getElementById("totalRollos").value,
    TotalYardasGlobal: document.getElementById("totalYardasGlobal").value,
    Resultado: document.getElementById("resultado").value,
    tYardas: document.getElementById("tYardas").textContent,
    tPuntos: document.getElementById("tPuntos").textContent,
    tRate: document.getElementById("tRate").textContent,
    Match: document.getElementById("match").value,
    Handfeel: document.getElementById("handfeel").value,
    Brushing: document.getElementById("brushing").value,
    Stretch: document.getElementById("stretch").value,
    Pilling: document.getElementById("pilling").value,
  };
  
  // 2. Definición del orden de las columnas (según el usuario)
  const headerOrder = [
    "Auditor:", "Proveedor:", "Lote:", "Part Number:", "Part Color:", 
    "Total Rollos:", "Total Yardas (Doc):", "# Rollo", "Yardas Label", 
    "Yardas Reales", "Ancho Estándar", "Ancho Real", "Yds Revisadas", 
    "Puntos por Rollo", "Rate por Rollo", "Peso Estándar", "Peso Real", 
    "% (+/-) Peso", "% Yardas Faltantes", "Obs.", "Resultado:", 
    "Total Yardas Reales:", "Puntos Penalizados Global:", "Rate Penalización Global:", 
    "Match:", "Handfeel:", "Brushing:", "Stretch:", "Pilling:"
  ];

  const dataRollos = [headerOrder]; // Fila 1: Encabezados

  // 3. Procesamiento de los datos de cada rollo
  tbody.querySelectorAll("tr").forEach(tr=>{
    // Obtener los datos específicos del rollo actual
    const rolloData = {
      "# Rollo": tr.querySelector(".rolloNum").value,
      "Yardas Label": tr.querySelector(".label").value,
      "Yardas Reales": tr.querySelector(".real").value,
      "Ancho Estándar": tr.querySelector(".anchoStd").value,
      "Ancho Real": tr.querySelector(".anchoReal").value,
      "Yds Revisadas": tr.querySelector(".revisadas").value,
      "Puntos por Rollo": tr.querySelector(".puntosRollo").value,
      "Rate por Rollo": tr.querySelector(".rateRollo").value,
      "Peso Estándar": tr.querySelector(".pesoStd").value,
      "Peso Real": tr.querySelector(".pesoReal").value,
      "% (+/-) Peso": tr.querySelector(".pct").value,
      "% Yardas Faltantes": tr.querySelector(".faltantes").value,
      "Obs.": tr.querySelector("textarea").value,
    };
    
    // Mapear los datos a la fila, siguiendo el orden definido
    const row = headerOrder.map(header => {
      // Campos de Rollo
      if (rolloData.hasOwnProperty(header)) {
        return rolloData[header];
      } 
      // Campos Globales (usando los nombres sin los dos puntos)
      else {
        const key = header.replace(/:/g, '');
        if (key === "Total Yardas (Doc)") return globalData.TotalYardasGlobal;
        if (key === "Total Yardas Reales") return globalData.tYardas;
        if (key === "Puntos Penalizados Global") return globalData.tPuntos;
        if (key === "Rate Penalización Global") return globalData.tRate;
        if (globalData.hasOwnProperty(key)) {
            return globalData[key];
        }
        return ''; // Vacío si no se encuentra
      }
    });

    dataRollos.push(row);
  });

  // Generación de la Pestaña Rollos_Lineal
  const wsRollos=XLSX.utils.aoa_to_sheet(dataRollos);
  XLSX.utils.book_append_sheet(wb,wsRollos,"Rollos_Lineal");


  // --- Pestaña Defectos (sin cambios en la estructura) ---
  let wsDef=[["Rollo","Cantidad Defectos","Puntos (1–4)","Código Defecto"]];
  tbodyDef.querySelectorAll("tr").forEach(tr=>{
    wsDef.push([
      tr.querySelector(".defRollo").value,
      tr.querySelector(".defCantidad").value,
      tr.querySelector(".defPuntos").value,
      tr.querySelector("input:not(.defRollo)").value
    ]);
  });
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(wsDef),"Defectos");

  // --- Pestaña Anchos (sin cambios en la estructura) ---
  let wsAnchos=[["Rollo","Ancho 1","Ancho 2","Ancho 3"]];
  tbodyAnchos.querySelectorAll("tr").forEach(tr=>{
    wsAnchos.push([
      tr.querySelector(".anchoRollo").value,
      tr.querySelector(".ancho1").value,
      tr.querySelector(".ancho2").value,
      tr.querySelector(".ancho3").value,
    ]);
  });
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(wsAnchos),"Anchos");

  // --- Generación del Archivo ---
  let fileName=globalData.Lote.replace(/[^a-z0-9]/gi,'_').toLowerCase()||"auditoria_tela";
  XLSX.writeFile(wb,`${fileName}.xlsx`);
});
</script>

</body>
</html>
