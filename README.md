<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Auditoría de Tela</title>

<style>
body{font-family:Arial,sans-serif;background:#f0f3f9;padding:18px;color:#222}
h2{color:#0b63d1;margin-top:20px}
table{width:100%;border-collapse:collapse;background:#fff;margin-top:10px}
th,td{border:1px solid #ccc;padding:6px;text-align:center}
th{background:#e6f0ff}
input,select,textarea{width:100%;padding:4px;border:1px solid #ccc;border-radius:4px}
textarea{resize:vertical}
button{padding:7px 12px;border:none;border-radius:4px;font-weight:bold;cursor:pointer}
.add{background:#0b63d1;color:#fff;margin-top:8px}
.del{background:#dc3545;color:#fff}
.box{background:#fff;border:1px solid #ccc;border-radius:6px;padding:8px 12px}
.flex{display:flex;gap:15px;flex-wrap:wrap}
.form-header{
display:grid;
grid-template-columns:repeat(auto-fit,minmax(200px,1fr));
gap:15px;
background:#fff;
padding:15px;
border-radius:6px;
border:1px solid #ccc
}
</style>

<script src="https://cdn.sheetjs.com/xlsx-latest/package/dist/xlsx.full.min.js"></script>
</head>

<body>

<h2>📋 Información de Auditoría</h2>
<div class="form-header">
<div><label>Lote<input id="lote"></label></div>
<div><label>Part Number<input id="pn"></label></div>
<div><label>Part Color<input id="color"></label></div>

<div><label>Proveedor
<select id="prov">
<option value="">Seleccione</option>
<option>INNOV TEXTILES</option>
<option>TEXPASA</option>
<option>RAYONES</option>
<option>NUEVO MUNDO</option>
</select></label></div>

<div><label>Auditor
<select id="aud">
<option value="">Seleccione</option>
<option>RENE RODRIGUEZ</option>
<option>EZEQUIEL RIVERA</option>
<option>FELIPE BETANCOURT</option>
</select></label></div>

<div><label>Total Rollos<input id="trollos" type="number"></label></div>
<div><label>Total Yardas<input id="tyardas" type="number"></label></div>
</div>

<h2>🧵 Rollos</h2>
<table>
<thead>
<tr>
<th># Rollo</th>
<th>Yds Label</th>
<th>Yds Reales</th>
<th>Ancho Std</th>
<th>Ancho Real</th>
<th>Puntos</th>
<th>Rate</th>
<th>Obs</th>
<th></th>
</tr>
</thead>
<tbody id="rollos"></tbody>
</table>
<button class="add" onclick="addRollo()">➕ Añadir Rollo</button>

<div class="flex">
<div class="box">Total Yardas Label: <span id="tLabel">0</span></div>
<div class="box">Total Yardas Reales: <span id="tReal">0</span></div>
<div class="box">% Faltante: <span id="pctFalt">0%</span></div>
<div class="box">Puntos Penalizados Global: <span id="tPuntos">0</span></div>
<div class="box">Rate Penalización Global: <span id="rateGlobal">0</span></div>
</div>

<h2>🚫 Defectos por Rollo</h2>
<table>
<thead>
<tr>
<th>Rollo</th>
<th>Cantidad</th>
<th>Puntos (1–4)</th>
<th>Código Defecto</th>
<th></th>
</tr>
</thead>
<tbody id="defectos"></tbody>
</table>
<button class="add" onclick="addDefecto()">➕ Añadir defecto</button>

<h2>📏 Anchos por Rollo</h2>
<table>
<thead>
<tr>
<th>Rollo</th>
<th>Ancho 1</th>
<th>Ancho 2</th>
<th>Ancho 3</th>
<th></th>
</tr>
</thead>
<tbody id="anchos"></tbody>
</table>
<button class="add" onclick="addAncho()">➕ Añadir Anchos</button>

<h2>🧪 Pruebas Adicionales</h2>
<div class="form-header">
<div><label>MATCH<select id="MATCH"><option>N/A</option><option>Aceptable</option><option>Rechazado</option></select></label></div>
<div><label>STRETCH<select id="STRETCH"><option>N/A</option><option>Aceptable</option><option>Rechazado</option></select></label></div>
<div><label>HANDFEEL<select id="HANDFEEL"><option>N/A</option><option>Aceptable</option><option>Rechazado</option></select></label></div>
<div><label>PILLING<select id="PILLING"><option>N/A</option><option>Aceptable</option><option>Rechazado</option></select></label></div>
<div><label>BRUSHING<select id="BRUSHING"><option>N/A</option><option>Aceptable</option><option>Rechazado</option></select></label></div>
</div>

<button class="add" onclick="exportar()">💾 Exportar XLSX</button>

<script>
function addRollo(){
const tr=document.createElement("tr");
tr.innerHTML=`
<td><input class="r"></td>
<td><input class="l" type="number"></td>
<td><input class="re" type="number"></td>
<td><input class="as" type="number"></td>
<td><input class="ar" type="number"></td>
<td><input class="p" readonly></td>
<td><input class="rate" readonly></td>
<td><textarea></textarea></td>
<td><button class="del" onclick="this.closest('tr').remove();calc()">X</button></td>`;
document.getElementById("rollos").appendChild(tr);
tr.querySelectorAll("input").forEach(i=>i.oninput=calc);
}
addRollo();

function addDefecto(){
const tr=document.createElement("tr");
tr.innerHTML=`
<td><input class="dr"></td>
<td><input class="dc" type="number"></td>
<td><select class="dp"><option>1</option><option>2</option><option>3</option><option>4</option></select></td>
<td><input class="cod"></td>
<td><button class="del" onclick="this.closest('tr').remove();calc()">X</button></td>`;
document.getElementById("defectos").appendChild(tr);
tr.querySelectorAll("input,select").forEach(i=>i.oninput=calc);
}

function addAncho(){
const tr=document.createElement("tr");
tr.innerHTML=`
<td><input></td>
<td><input type="number"></td>
<td><input type="number"></td>
<td><input type="number"></td>
<td><button class="del" onclick="this.closest('tr').remove()">X</button></td>`;
document.getElementById("anchos").appendChild(tr);
}

function calc(){
let tL=0,tR=0,tP=0;
const map={};

document.querySelectorAll("#defectos tr").forEach(tr=>{
const r=tr.querySelector(".dr").value;
const c=+tr.querySelector(".dc").value||0;
const p=+tr.querySelector(".dp").value||0;
if(r) map[r]=(map[r]||0)+(c*p);
tP+=c*p;
});

document.querySelectorAll("#rollos tr").forEach(tr=>{
const r=tr.querySelector(".r").value;
const l=+tr.querySelector(".l").value||0;
const re=+tr.querySelector(".re").value||0;
const ar=+tr.querySelector(".ar").value||0;

tL+=l; tR+=re;
const pts=map[r]||0;
tr.querySelector(".p").value=pts;
tr.querySelector(".rate").value=(re&&ar)?((pts*36)/(re*ar)*100).toFixed(2):0;
});

tLabel.textContent=tL.toFixed(2);
tReal.textContent=tR.toFixed(2);
pctFalt.textContent=tL?(((tR/tL-1)*100).toFixed(2)+"%"):"0%";
tPuntos.textContent=tP.toFixed(2);
rateGlobal.textContent=(tR?((tP*36)/tR).toFixed(2):0);
}

function exportar(){
const fecha=new Date().toISOString().slice(0,10);
const wb=XLSX.utils.book_new();

/* ROLLOS */
const headers=[
"FECHA",
"Auditor","Proveedor","Lote","Part Number","Part Color",
"Total Yds Label","Total Yds Reales","% Faltante",
"Puntos Globales","Rate Global",
"MATCH","STRETCH","HANDFEEL","PILLING","BRUSHING",
"# Rollo","Yds Label","Yds Reales","Ancho Std","Ancho Real","Puntos","Rate","Obs"
];
const data=[headers];

const rollos=[...document.querySelectorAll("#rollos tr")];
if(rollos.length){
const tr=rollos[0];
data.push([
fecha,
aud.value,prov.value,lote.value,pn.value,color.value,
tLabel.textContent,tReal.textContent,pctFalt.textContent,
tPuntos.textContent,rateGlobal.textContent,
MATCH.value,STRETCH.value,HANDFEEL.value,PILLING.value,BRUSHING.value,
tr.querySelector(".r").value,
tr.querySelector(".l").value,
tr.querySelector(".re").value,
tr.querySelector(".as").value,
tr.querySelector(".ar").value,
tr.querySelector(".p").value,
tr.querySelector(".rate").value,
tr.querySelector("textarea").value
]);
}

rollos.slice(1).forEach(tr=>{
data.push([
"",
"","","","","","","","","","","","","","","",
tr.querySelector(".r").value,
tr.querySelector(".l").value,
tr.querySelector(".re").value,
tr.querySelector(".as").value,
tr.querySelector(".ar").value,
tr.querySelector(".p").value,
tr.querySelector(".rate").value,
tr.querySelector("textarea").value
]);
});

XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(data),"Rollos");

/* DEFECTOS */
const dData=[["Rollo","Cantidad","Puntos","Código Defecto"]];
document.querySelectorAll("#defectos tr").forEach(tr=>{
dData.push([
tr.querySelector(".dr").value,
tr.querySelector(".dc").value,
tr.querySelector(".dp").value,
tr.querySelector(".cod").value
]);
});
XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(dData),"Defectos");

XLSX.writeFile(wb,`${lote.value||"auditoria"}_${fecha}.xlsx`);
}
</script>

</body>
</html>
