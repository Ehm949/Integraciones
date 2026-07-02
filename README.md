[index (8).html](https://github.com/user-attachments/files/29611593/index.8.html)
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Conciliación de Fleteros · Sol Firma</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
  :root{
    --navy-950:#0a0e17;
    --navy-900:#0f1521;
    --navy-800:#151d2e;
    --navy-700:#1c2740;
    --navy-600:#28365a;
    --line:#26314a;
    --amber:#d99a3f;
    --amber-bright:#f0b44f;
    --text:#e8ecf4;
    --text-dim:#8b96ad;
    --text-faint:#5c6680;
    --ok:#4caf82;
    --warn:#e0a838;
    --bad:#d9645f;
    --mono:'JetBrains Mono','SF Mono',Consolas,monospace;
    --sans:'Inter',-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;
  }
  *{box-sizing:border-box;}
  body{
    margin:0;
    background:var(--navy-950);
    background-image:
      radial-gradient(circle at 15% 0%, rgba(217,154,63,0.06), transparent 45%),
      radial-gradient(circle at 85% 100%, rgba(76,175,130,0.04), transparent 40%);
    color:var(--text);
    font-family:var(--sans);
    min-height:100vh;
    -webkit-font-smoothing:antialiased;
  }
  .wrap{max-width:1180px;margin:0 auto;padding:28px 24px 60px;}

  header{
    display:flex;justify-content:space-between;align-items:flex-end;
    border-bottom:1px solid var(--line);
    padding-bottom:16px;margin-bottom:20px;
    gap:20px;flex-wrap:wrap;
  }
  .brand-eyebrow{
    font-family:var(--mono);font-size:11px;letter-spacing:.18em;
    color:var(--amber);text-transform:uppercase;margin-bottom:8px;
  }
  h1{font-size:26px;margin:0;font-weight:650;letter-spacing:-.01em;}
  header p{color:var(--text-dim);margin:6px 0 0;font-size:13.5px;max-width:520px;line-height:1.5;}
  .badge{
    font-family:var(--mono);font-size:11px;color:var(--text-faint);
    border:1px solid var(--line);padding:6px 12px;border-radius:20px;
    white-space:nowrap;
  }

  /* Dropzone */
  .dropzone{
    border:1.5px dashed var(--navy-600);
    border-radius:14px;
    background:linear-gradient(180deg, var(--navy-900), var(--navy-800));
    padding:22px 24px;
    text-align:center;
    cursor:pointer;
    transition:.18s border-color, .18s background;
    margin-bottom:18px;
  }
  .dropzone.drag{border-color:var(--amber);background:var(--navy-800);}
  .dropzone svg{width:30px;height:30px;stroke:var(--amber);margin-bottom:10px;}
  .dropzone .dz-title{font-size:15px;font-weight:600;margin-bottom:4px;}
  .dropzone .dz-sub{font-size:12.5px;color:var(--text-faint);}
  .dropzone input{display:none;}
  .file-chip{
    display:inline-flex;align-items:center;gap:8px;
    background:var(--navy-700);border:1px solid var(--line);
    padding:6px 10px 6px 12px;border-radius:8px;font-size:12.5px;
    margin:4px 5px 0 0;font-family:var(--mono);color:var(--text-dim);
  }
  .file-chip button{
    background:none;border:none;color:var(--text-faint);cursor:pointer;
    font-size:14px;line-height:1;padding:0 2px;
  }
  .file-chip button:hover{color:var(--bad);}

  /* Summary cards */
  .cards{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:20px;}
  @media(max-width:820px){.cards{grid-template-columns:repeat(2,1fr);}}
  .card{
    background:var(--navy-900);border:1px solid var(--line);
    border-radius:12px;padding:12px 16px;
  }
  .card .label{font-size:10.5px;color:var(--text-faint);text-transform:uppercase;letter-spacing:.08em;margin-bottom:5px;}
  .card .value{font-family:var(--mono);font-size:19px;font-weight:600;}
  .card.owe .value{color:var(--bad);}
  .card.paid .value{color:var(--ok);}
  .card.pending-count .value{color:var(--amber-bright);}

  /* Controls */
  .controls{
    display:flex;justify-content:space-between;align-items:center;
    margin-bottom:14px;gap:12px;flex-wrap:wrap;
  }
  .tabs{display:flex;gap:6px;flex-wrap:wrap;}
  .tab{
    font-family:var(--mono);font-size:12px;padding:7px 13px;border-radius:8px;
    border:1px solid var(--line);background:var(--navy-900);color:var(--text-dim);
    cursor:pointer;transition:.15s;
  }
  .tab.active{background:var(--amber);color:var(--navy-950);border-color:var(--amber);font-weight:600;}
  .tab:hover:not(.active){border-color:var(--navy-600);color:var(--text);}

  .right-controls{display:flex;gap:8px;align-items:center;}
  select, .btn{
    font-family:var(--sans);font-size:12.5px;
    background:var(--navy-800);color:var(--text);
    border:1px solid var(--line);border-radius:8px;padding:8px 12px;
    cursor:pointer;
  }
  .btn.primary{background:var(--amber);color:var(--navy-950);border-color:var(--amber);font-weight:600;}
  .btn.primary:hover{background:var(--amber-bright);}
  .btn:disabled{opacity:.4;cursor:not-allowed;}

  /* Table */
  .table-wrap{
    border:1px solid var(--line);border-radius:12px;overflow:hidden;
    background:var(--navy-900);
  }
  table{width:100%;border-collapse:collapse;font-size:13px;}
  thead th{
    text-align:left;font-family:var(--mono);font-size:10.5px;
    text-transform:uppercase;letter-spacing:.06em;color:var(--text-faint);
    padding:11px 14px;border-bottom:1px solid var(--line);background:var(--navy-800);
    position:sticky;top:0;
  }
  tbody td{padding:5px 12px;border-bottom:1px solid var(--navy-800);vertical-align:top;line-height:1.3;}
  tbody tr:last-child td{border-bottom:none;}
  tbody tr:hover{background:var(--navy-800);}
  table.compact thead th{padding:8px 12px;}
  table.compact tbody td{padding:4.5px 12px;font-size:12.3px;}
  .mono{font-family:var(--mono);}
  .num{text-align:right;font-family:var(--mono);white-space:nowrap;}
  .status-pill{
    display:inline-block;font-family:var(--mono);font-size:10px;
    padding:2px 8px;border-radius:20px;font-weight:600;letter-spacing:.02em;
    white-space:nowrap;
  }
  .st-pendiente{background:rgba(217,100,95,.14);color:var(--bad);}
  .st-excedente{background:rgba(224,168,56,.14);color:var(--warn);}
  .st-ok{background:rgba(76,175,130,.14);color:var(--ok);}
  .fletero-name{font-weight:600;font-size:12px;}
  .ref-cell{font-size:11.5px;line-height:1.35;}
  .ref-cell .num-part{font-family:var(--mono);color:var(--text);}
  .ref-cell .date-part{color:var(--text-faint);font-family:var(--mono);}
  .ref-cell .empty-ref{color:var(--text-faint);}
  .ref-line + .ref-line{margin-top:1px;}

  .empty-state{
    text-align:center;padding:60px 20px;color:var(--text-faint);
    font-size:13.5px;
  }
  footer{
    margin-top:40px;text-align:center;color:var(--text-faint);
    font-size:11.5px;font-family:var(--mono);
  }
  .section-title{
    font-size:13px;font-weight:600;color:var(--text-dim);
    margin:26px 0 10px;text-transform:uppercase;letter-spacing:.05em;
  }
  .loading{color:var(--amber);font-family:var(--mono);font-size:12.5px;text-align:center;padding:20px;}
</style>
</head>
<body>
<div class="wrap">

  <header>
    <div>
      <div class="brand-eyebrow">Sol Firma · Cuentas por Pagar</div>
      <h1>Conciliación de Fleteros</h1>
      <p>Sube el Excel contable (Diarios y Egresos) y el sistema calcula automáticamente cuánto se le debe a cada fletero, cruzando cada factura contra su pago por folio.</p>
    </div>
    <div class="badge" id="statusBadge">Sin archivos cargados</div>
  </header>

  <div class="dropzone" id="dropzone">
    <input type="file" id="fileInput" accept=".xlsx,.xls" multiple>
    <svg viewBox="0 0 24 24" fill="none" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3v13"/><path d="M7 8l5-5 5 5"/><path d="M4 17v3a2 2 0 002 2h12a2 2 0 002-2v-3"/></svg>
    <div class="dz-title">Arrastra tu(s) archivo(s) .xlsx aquí, o haz clic para seleccionar</div>
    <div class="dz-sub">Acepta el formato de diario contable con columnas Tipo (Dr/Eg), Número, Fecha, Concepto, Debe, Haber, Saldo. Puedes cargar varios archivos a la vez.</div>
    <div id="fileList" style="margin-top:12px;"></div>
  </div>

  <div id="loadingMsg" style="display:none;" class="loading">Procesando archivo(s)…</div>

  <div id="results" style="display:none;">

    <div class="cards">
      <div class="card owe">
        <div class="label">Total que debemos</div>
        <div class="value" id="cardOwe">$0.00</div>
      </div>
      <div class="card paid">
        <div class="label">Total facturado (Dr)</div>
        <div class="value" id="cardInvoiced">$0.00</div>
      </div>
      <div class="card paid">
        <div class="label">Total pagado (Eg)</div>
        <div class="value" id="cardPaid">$0.00</div>
      </div>
      <div class="card pending-count">
        <div class="label">Facturas pendientes</div>
        <div class="value" id="cardPending">0</div>
      </div>
    </div>

    <div class="controls">
      <div class="tabs" id="fleteroTabs"></div>
      <div class="right-controls">
        <select id="statusFilter">
          <option value="pendiente">Solo pendientes / excedentes</option>
          <option value="all">Mostrar todo (incluye saldados)</option>
        </select>
        <button class="btn primary" id="exportBtn">Exportar reporte .xlsx</button>
      </div>
    </div>

    <div class="table-wrap">
      <table class="compact">
        <thead>
          <tr>
            <th>Fletero</th>
            <th>Folio</th>
            <th>Diario (No. / Fecha)</th>
            <th>Egreso (No. / Fecha)</th>
            <th class="num">Facturado</th>
            <th class="num">Pagado</th>
            <th class="num">Diferencia</th>
            <th>Estatus</th>
          </tr>
        </thead>
        <tbody id="resultsBody"></tbody>
      </table>
    </div>

    <div id="emptyState" class="empty-state" style="display:none;">
      No hay diferencias pendientes con este filtro — todo está saldado 🎉
    </div>
  </div>

  <footer>Sol Firma · Herramienta interna de validación de fletes · Procesamiento 100% local en el navegador</footer>
</div>

<script>
/* ============================================================
   CONCILIACIÓN DE FLETEROS
   Lógica:
   - Cada bloque de cuenta empieza en una fila con Col A = '-'
     y texto "Cuenta : <codigo> <NOMBRE FLETERO>" en Col B.
   - Filas de datos: Col B = 'Dr' (factura/diario) o 'Eg' (pago/egreso)
   - Col E (Concepto) = "<folio> <NOMBRE FISCAL> <uuid-opcional>"
     El folio (primer token) es la llave de cruce entre Dr y Eg.
   - Dr: monto en columna Haber (index 7) = facturado
   - Eg: monto en columna Debe   (index 6) = pagado
   - Diferencia = facturado - pagado
     > 0  => les debemos (pendiente de pago)
     < 0  => pagamos de más (excedente)
     = 0  => saldado
   ============================================================ */

let allRecords = []; // {fletero, folio, uuid, facturado, pagado}
let loadedFiles = [];

const dropzone = document.getElementById('dropzone');
const fileInput = document.getElementById('fileInput');
const fileListEl = document.getElementById('fileList');
const statusBadge = document.getElementById('statusBadge');
const loadingMsg = document.getElementById('loadingMsg');
const resultsEl = document.getElementById('results');

dropzone.addEventListener('click', (e) => { if(e.target.tagName !== 'BUTTON') fileInput.click(); });
dropzone.addEventListener('dragover', (e)=>{e.preventDefault();dropzone.classList.add('drag');});
dropzone.addEventListener('dragleave', ()=>dropzone.classList.remove('drag'));
dropzone.addEventListener('drop', (e)=>{
  e.preventDefault();dropzone.classList.remove('drag');
  handleFiles(e.dataTransfer.files);
});
fileInput.addEventListener('change', (e)=>handleFiles(e.target.files));

function handleFiles(fileListArg){
  const files = Array.from(fileListArg).filter(f=>/\.xlsx?$/i.test(f.name));
  if(files.length===0) return;
  files.forEach(f=>{
    if(!loadedFiles.find(lf=>lf.name===f.name && lf.size===f.size)){
      loadedFiles.push(f);
    }
  });
  renderFileChips();
  processAllFiles();
}

function renderFileChips(){
  fileListEl.innerHTML = loadedFiles.map((f,i)=>`
    <span class="file-chip">${escapeHtml(f.name)}
      <button data-idx="${i}" onclick="removeFile(${i})">×</button>
    </span>`).join('');
}
function removeFile(idx){
  loadedFiles.splice(idx,1);
  renderFileChips();
  if(loadedFiles.length===0){
    resultsEl.style.display='none';
    statusBadge.textContent='Sin archivos cargados';
    allRecords=[];
  } else {
    processAllFiles();
  }
}

function escapeHtml(s){
  return String(s).replace(/[&<>"']/g, c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}

async function processAllFiles(){
  loadingMsg.style.display='block';
  resultsEl.style.display='none';
  await new Promise(r=>setTimeout(r,30)); // let UI paint

  allRecords = [];
  for(const file of loadedFiles){
    try{
      const data = await file.arrayBuffer();
      const wb = XLSX.read(data, {type:'array'});
      wb.SheetNames.forEach(sheetName=>{
        const ws = wb.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(ws, {header:1, raw:true, defval:null});
        parseSheet(rows, allRecords);
      });
    }catch(err){
      console.error('Error procesando', file.name, err);
      alert('No se pudo leer el archivo "'+file.name+'". Verifica que sea un .xlsx válido.\n\n'+err.message);
    }
  }

  loadingMsg.style.display='none';
  if(allRecords.length===0){
    statusBadge.textContent='No se detectaron registros con este formato';
    return;
  }
  buildAndRender();
}

function parseSheet(rows, out){
  let currentFletero = null;
  // map: fletero -> folio -> {facturado, pagado, uuid, diarios:[], egresos:[]}
  const localMap = {};

  for(let i=0;i<rows.length;i++){
    const r = rows[i] || [];
    const colA = r[0];
    const colB = r[1];
    const colC = r[2]; // Numero
    const colD = r[3]; // Fecha
    const colE = r[4];
    const colG = num(r[6]); // Debe
    const colH = num(r[7]); // Haber

    // Detect account header row: Col A === '-' and Col B contains "Cuenta"
    if(colA === '-' && typeof colB === 'string' && /cuenta/i.test(colB)){
      currentFletero = extractFleteroName(colB);
      continue;
    }

    if(!currentFletero) continue;
    if(colB !== 'Dr' && colB !== 'Eg') continue;
    if(!colE || typeof colE !== 'string') continue;

    const parsed = parseConcepto(colE);
    if(!parsed.folio) continue;

    if(!localMap[currentFletero]) localMap[currentFletero] = {};
    if(!localMap[currentFletero][parsed.folio]){
      localMap[currentFletero][parsed.folio] = {facturado:0, pagado:0, uuid:'', diarios:[], egresos:[]};
    }
    const entry = localMap[currentFletero][parsed.folio];
    const numero = colC !== null && colC !== undefined ? String(colC).trim() : '';
    const fecha = colD !== null && colD !== undefined ? String(colD).trim() : '';

    if(colB === 'Dr'){
      entry.facturado += colH;
      if(parsed.uuid) entry.uuid = parsed.uuid;
      if(numero && !entry.diarios.some(d=>d.numero===numero && d.fecha===fecha)){
        entry.diarios.push({numero, fecha});
      }
    } else if(colB === 'Eg'){
      entry.pagado += colG;
      if(numero && !entry.egresos.some(d=>d.numero===numero && d.fecha===fecha)){
        entry.egresos.push({numero, fecha});
      }
    }
  }

  // flatten into out[]
  Object.keys(localMap).forEach(fletero=>{
    Object.keys(localMap[fletero]).forEach(folio=>{
      const e = localMap[fletero][folio];
      out.push({
        fletero, folio,
        uuid: e.uuid,
        diarios: e.diarios,
        egresos: e.egresos,
        facturado: e.facturado,
        pagado: e.pagado,
        diferencia: round2(e.facturado - e.pagado)
      });
    });
  });
}

function extractFleteroName(text){
  // "Cuenta : 2101-3071-0000-0000 MARIN GOMEZ MARIA GUADALUPE"
  const parts = text.split(':');
  const tail = (parts[1] || text).trim();
  // remove leading account code (letters/numbers/dashes) if present
  const m = tail.match(/^[\d\-A-Za-z]+\s+(.*)$/);
  return (m ? m[1] : tail).trim().toUpperCase();
}

function parseConcepto(text){
  const t = text.trim();
  const uuidMatch = t.match(/[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}/);
  const tokens = t.split(/\s+/);
  const folio = tokens[0] || '';
  return {
    folio: folio,
    uuid: uuidMatch ? uuidMatch[0] : ''
  };
}

function num(v){
  if(v === null || v === undefined || v === '') return 0;
  const n = typeof v === 'number' ? v : parseFloat(String(v).replace(/,/g,''));
  return isNaN(n) ? 0 : n;
}
function round2(n){ return Math.round(n*100)/100; }
function money(n){
  const sign = n < 0 ? '-' : '';
  return sign + '$' + Math.abs(n).toLocaleString('es-MX', {minimumFractionDigits:2, maximumFractionDigits:2});
}

function refCellHtml(refs){
  if(!refs || refs.length === 0) return '<span class="empty-ref">—</span>';
  return refs.map(r=>
    `<div class="ref-line"><span class="num-part">#${escapeHtml(r.numero)}</span> <span class="date-part">${escapeHtml(r.fecha || '')}</span></div>`
  ).join('');
}
function refCellText(refs){
  if(!refs || refs.length === 0) return '';
  return refs.map(r=>`#${r.numero} (${r.fecha||''})`).join(', ');
}

let currentFleteroFilter = 'ALL';

function buildAndRender(){
  statusBadge.textContent = allRecords.length + ' folios detectados en ' + loadedFiles.length + ' archivo(s)';
  resultsEl.style.display = 'block';

  // Fletero tabs
  const fleteros = [...new Set(allRecords.map(r=>r.fletero))].sort();
  const tabsEl = document.getElementById('fleteroTabs');
  tabsEl.innerHTML = ['<button class="tab active" data-f="ALL">Todos los fleteros</button>']
    .concat(fleteros.map(f=>`<button class="tab" data-f="${escapeHtml(f)}">${escapeHtml(f)}</button>`))
    .join('');
  tabsEl.querySelectorAll('.tab').forEach(btn=>{
    btn.addEventListener('click', ()=>{
      tabsEl.querySelectorAll('.tab').forEach(b=>b.classList.remove('active'));
      btn.classList.add('active');
      currentFleteroFilter = btn.dataset.f;
      renderTable();
    });
  });

  document.getElementById('statusFilter').onchange = renderTable;
  document.getElementById('exportBtn').onclick = exportReport;

  renderTable();
}

function folioSortKey(folio){
  // Numeric-aware sort: pulls out the numeric portion so "001355" sorts before "001367", etc.
  const n = parseInt(String(folio).replace(/\D/g,''), 10);
  return isNaN(n) ? Infinity : n;
}

function getFilteredRecords(){
  let recs = allRecords;
  if(currentFleteroFilter !== 'ALL'){
    recs = recs.filter(r=>r.fletero === currentFleteroFilter);
  }
  const statusFilter = document.getElementById('statusFilter').value;
  if(statusFilter === 'pendiente'){
    recs = recs.filter(r=>Math.abs(r.diferencia) > 0.005);
  }
  return recs.slice().sort((a,b)=>{
    if(a.fletero !== b.fletero) return a.fletero.localeCompare(b.fletero);
    const ka = folioSortKey(a.folio), kb = folioSortKey(b.folio);
    if(ka !== kb) return ka - kb;
    return String(a.folio).localeCompare(String(b.folio));
  });
}

function renderTable(){
  const recs = getFilteredRecords();
  const tbody = document.getElementById('resultsBody');
  const emptyState = document.getElementById('emptyState');

  if(recs.length === 0){
    tbody.innerHTML = '';
    emptyState.style.display = 'block';
  } else {
    emptyState.style.display = 'none';
    tbody.innerHTML = recs.map(r=>{
      let statusHtml;
      if(r.diferencia > 0.005){
        statusHtml = '<span class="status-pill st-pendiente">DEBEMOS</span>';
      } else if(r.diferencia < -0.005){
        statusHtml = '<span class="status-pill st-excedente">PAGO DE MÁS</span>';
      } else {
        statusHtml = '<span class="status-pill st-ok">SALDADO</span>';
      }
      return `<tr>
        <td class="fletero-name">${escapeHtml(r.fletero)}</td>
        <td class="mono">${escapeHtml(r.folio)}</td>
        <td class="ref-cell">${refCellHtml(r.diarios)}</td>
        <td class="ref-cell">${refCellHtml(r.egresos)}</td>
        <td class="num">${money(r.facturado)}</td>
        <td class="num">${money(r.pagado)}</td>
        <td class="num">${money(r.diferencia)}</td>
        <td>${statusHtml}</td>
      </tr>`;
    }).join('');
  }

  // Summary cards reflect the currently selected fletero scope (all statuses, not the status filter)
  let scopeRecs = allRecords;
  if(currentFleteroFilter !== 'ALL'){
    scopeRecs = scopeRecs.filter(r=>r.fletero === currentFleteroFilter);
  }
  const totalInvoiced = scopeRecs.reduce((s,r)=>s+r.facturado,0);
  const totalPaid = scopeRecs.reduce((s,r)=>s+r.pagado,0);
  const totalOwe = round2(totalInvoiced - totalPaid);
  const pendingCount = scopeRecs.filter(r=>Math.abs(r.diferencia)>0.005).length;

  document.getElementById('cardOwe').textContent = money(totalOwe);
  document.getElementById('cardInvoiced').textContent = money(totalInvoiced);
  document.getElementById('cardPaid').textContent = money(totalPaid);
  document.getElementById('cardPending').textContent = pendingCount;
}

function exportReport(){
  const recs = getFilteredRecords();
  const exportRows = recs.map(r=>({
    'Fletero': r.fletero,
    'Folio': r.folio,
    'Diario (No. / Fecha)': refCellText(r.diarios),
    'Egreso (No. / Fecha)': refCellText(r.egresos),
    'Folio Fiscal (UUID)': r.uuid,
    'Facturado': r.facturado,
    'Pagado': r.pagado,
    'Diferencia': r.diferencia,
    'Estatus': r.diferencia > 0.005 ? 'DEBEMOS' : (r.diferencia < -0.005 ? 'PAGO DE MAS' : 'SALDADO')
  }));
  const ws = XLSX.utils.json_to_sheet(exportRows);
  ws['!cols'] = [{wch:32},{wch:10},{wch:22},{wch:22},{wch:38},{wch:14},{wch:14},{wch:14},{wch:14}];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Conciliacion');
  const today = new Date().toISOString().slice(0,10);
  XLSX.writeFile(wb, `Conciliacion_Fleteros_${today}.xlsx`);
}
</script>
</body>
</html>
