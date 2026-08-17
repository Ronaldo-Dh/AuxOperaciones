const els = {
    fileGram: document.getElementById('fileGram'),
    fileBase: document.getElementById('fileBase'),
    dropGram: document.getElementById('dropGram'),
    dropBase: document.getElementById('dropBase'),
    fnameGram: document.getElementById('fnameGram'),
    fnameBase: document.getElementById('fnameBase'),
    fmetaGram: document.getElementById('fmetaGram'),
    fmetaBase: document.getElementById('fmetaBase'),
    btnProcess: document.getElementById('btnProcess'),
    btnDownload: document.getElementById('btnDownload'),
    btnReset: document.getElementById('btnReset'),
    log: document.getElementById('log'),
    statusPill: document.getElementById('statusPill'),
    stats: document.getElementById('stats'),
    statGramRows: document.getElementById('statGramRows'),
    statBaseRows: document.getElementById('statBaseRows'),
    statEnCurso: document.getElementById('statEnCurso'),
    statMatch: document.getElementById('statMatch'),
    resultsHead: document.getElementById('resultsHead'),
    tableWrap: document.getElementById('tableWrap'),
    tbody: document.getElementById('tbody'),
    emptyMsg: document.getElementById('emptyMsg'),
    filterBox: document.getElementById('filterBox'),
    loadingOverlay: document.getElementById('loadingOverlay'),
loadingMessage: document.getElementById('loadingMessage'),
loadingDetail: document.getElementById('loadingDetail'),
loadingProgressBar: document.getElementById('loadingProgressBar'),
loadingPercent: document.getElementById('loadingPercent'),
  };
  
let gramFile = null, baseFile = null;
let resultRows = [];

const worker = new Worker('./worker.js');
  
  function fmtBytes(b){
    if(b>1024*1024) return (b/1024/1024).toFixed(1)+' MB';
    return (b/1024).toFixed(0)+' KB';
  }
  
  function wireDrop(dropEl, inputEl, onFile){
    dropEl.addEventListener('dragover', e=>{e.preventDefault(); dropEl.classList.add('drag');});
    dropEl.addEventListener('dragleave', ()=>dropEl.classList.remove('drag'));
    dropEl.addEventListener('drop', e=>{
      e.preventDefault(); dropEl.classList.remove('drag');
      if(e.dataTransfer.files.length){ inputEl.files = e.dataTransfer.files; onFile(e.dataTransfer.files[0]); }
    });
    inputEl.addEventListener('change', ()=>{ if(inputEl.files.length) onFile(inputEl.files[0]); });
  }
  
  wireDrop(els.dropGram, els.fileGram, f=>{
    gramFile = f;
    els.fnameGram.textContent = f.name;
    els.fmetaGram.textContent = fmtBytes(f.size);
    els.dropGram.classList.add('filled');
    updateProcessBtn();
  });
  wireDrop(els.dropBase, els.fileBase, f=>{
    baseFile = f;
    els.fnameBase.textContent = f.name;
    els.fmetaBase.textContent = fmtBytes(f.size);
    els.dropBase.classList.add('filled');
    updateProcessBtn();
  });

  
  function updateProcessBtn(){
    els.btnProcess.disabled = !(gramFile && baseFile);
  }

  function showLoading(){
    document.body.classList.add('processing');
    els.loadingOverlay.classList.add('active');
  
    setLoading(
      0,
      'Preparando procesamiento…',
      'Inicializando lectura de archivos.'
    );
  }
  
  function hideLoading(){
    document.body.classList.remove('processing');
    els.loadingOverlay.classList.remove('active');
  }
  
  function setLoading(percent, message, detail=''){
    percent = Math.max(0, Math.min(100, percent));
  
    els.loadingProgressBar.style.width = percent + '%';
    els.loadingPercent.textContent = Math.round(percent) + '%';
    els.loadingMessage.textContent = message;
    els.loadingDetail.textContent = detail;
  }

  function logLine(msg, kind){
    const row = document.createElement('div');
    row.className = 'ln ' + (kind||'run');
    const t = document.createElement('span');
    t.className = 't';
    t.textContent = new Date().toLocaleTimeString('es-PE', {hour12:false});
    const m = document.createElement('span');
    m.className = 'msg';
    m.textContent = msg;
    row.appendChild(t); row.appendChild(m);
    els.log.appendChild(row);
    els.log.scrollTop = els.log.scrollHeight;
  }
  function tick(){ return new Promise(r=>setTimeout(r, 0)); }
  function workerRead(action, buffer, sheets = []) {

    return new Promise((resolve, reject) => {
  
      const handler = (e) => {
  
        const data = e.data;
  
        if (data.type === 'progress') {
  
          setLoading(
            data.percent,
            data.message,
            data.detail
          );
  
          return;
        }
  
        if (data.type === 'error') {
  
          worker.removeEventListener('message', handler);
  
          reject(new Error(data.message));
  
          return;
        }
  
        if (action === 'readGram' && data.type === 'gram') {
  
          worker.removeEventListener('message', handler);
  
          resolve(data.aoa);
  
          return;
        }
  
        if (action === 'readBase' && data.type === 'base') {
  
          worker.removeEventListener('message', handler);
  
          resolve(data.sheets);
  
          return;
        }
  
      };
  
      worker.addEventListener('message', handler);
  
      worker.postMessage(
        {
          action,
          buffer,
          sheets
        },
        [buffer]
      );
  
    });
  
  }

  
  // ---------- helpers ----------
  function normHeader(s){
    return (s==null?'':String(s))
      .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
      .replace(/\s+/g,' ').trim().toUpperCase();
  }
  function findCol(headers, includesAny){
    for(let i=0;i<headers.length;i++){
      const h = headers[i];
      if(includesAny.some(tok => h.includes(tok))) return i;
    }
    return -1;
  }
  function parseSpanishDate(v){
    if(v==null || v==='') return null;
    if(v instanceof Date) return v;
    if(typeof v === 'number'){
      // excel serial date
      const d = XLSX.SSF.parse_date_code(v);
      if(!d) return null;
      return new Date(d.y, d.m-1, d.d, d.H||0, d.M||0, Math.floor(d.S||0));
    }
    const s = String(v).trim();
    const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
    if(!m) return null;
    const [, dd, mm, yyyy, hh, mi, ss] = m;
    return new Date(+yyyy, +mm-1, +dd, +(hh||0), +(mi||0), +(ss||0));
  }
  function isNumericST(v){
    if(v==null) return false;
    if(typeof v === 'number') return isFinite(v);
    const s = String(v).trim();
    if(s==='') return false;
    return /^\d+(\.\d+)?$/.test(s);
  }
  function toIntST(v){
    if(typeof v === 'number') return Math.round(v);
    return parseInt(String(v).trim(), 10);
  }
  function excelDateFmt(d){
    return d ? d : null;
  }
  
  // ---------- main  ----------
  async function processFiles(){

    els.btnProcess.disabled = true;
    els.btnDownload.disabled = true;
  
    showLoading();
  
    els.log.innerHTML = '';
    els.stats.style.display = 'none';
    els.resultsHead.style.display = 'none';
    els.tableWrap.style.display = 'none';
    els.emptyMsg.style.display = 'none';
  
    els.statusPill.textContent = 'procesando…';
    resultRows = [];
  
    try{
      // ---- ETAPA 1: read GRAM EAM ---- / Observación:en el 10 porciento se laguea - ta para corregir
      setLoading(
        10,
        'Leyendo reporte GRAM EAM',
        'Cargando y analizando el archivo de Solicitudes de Trabajo…'
      );
      logLine('Leyendo reporte GRAM EAM…'); await tick();
      const gramBuf = await gramFile.arrayBuffer();

      const gramAoa = await workerRead(
       'readGram',
      gramBuf
    );


      logLine('GRAM EAM leído: ' + gramAoa.length + ' filas en bruto.', 'ok'); await tick(); //parche
      setLoading(
        20,
        'Analizando GRAM EAM',
        gramAoa.length.toLocaleString('es-PE') + ' filas encontradas.'
      );

      await tick();


  
      // find header row (contains SOLICITUD) / Up
      let headerIdx = -1, headers = [];
      for(let i=0;i<Math.min(gramAoa.length, 30);i++){
        const row = gramAoa[i] || [];
        const norm = row.map(normHeader);
        if(norm.some(h=>h.includes('SOLICITUD TRABAJO') || h==='SOLICITUD')){
          headerIdx = i; headers = norm; break;
        }
      }
      if(headerIdx === -1) throw new Error('No se encontró la fila de encabezados (con "SOLICITUD TRABAJO") en el reporte GRAM EAM.');
      logLine('Encabezados detectados en fila ' + (headerIdx+1) + '.', 'ok'); await tick();
  
      const idxSolicitud = findCol(headers, ['SOLICITUD TRABAJO','SOLICITUD']);
      const idxFechaFinal = findCol(headers, ['FECHA FINALIZACION PT','FINALIZACION PT']);
      const idxFechaInicio = findCol(headers, ['FECHA INICIO REAL','INICIO REAL']);
      const idxFechaFin = findCol(headers, ['FECHA FIN REAL','FIN REAL']);
      const idxUsuarioPT = findCol(headers, ['USUARIO PT']);
      if([idxSolicitud, idxFechaFinal, idxFechaInicio, idxFechaFin, idxUsuarioPT].includes(-1)){
        throw new Error('No se pudieron mapear todas las columnas requeridas (Solicitud, Fecha Final, Fecha Inicio, Fecha Fin, Usuario PT) en GRAM EAM.');
      }
      logLine('Columnas mapeadas → Solicitud, Fecha Final, Fecha Inicio, Fecha Fin, Usuario PT.', 'ok'); await tick();
  
      // build clean GRAM rows
      logLine('Convirtiendo Solicitud a número y fechas a formato fecha…'); await tick();
      const gramRows = [];
      let skippedNoSolicitud = 0;
      for(let i=headerIdx+1; i<gramAoa.length; i++){
        const row = gramAoa[i];
        if(!row) continue;
        const rawSol = row[idxSolicitud];
        if(rawSol==null || String(rawSol).trim()==='') continue;
        const sol = parseInt(String(rawSol).trim(), 10);
        if(isNaN(sol)){ skippedNoSolicitud++; continue; }
        gramRows.push({
          solicitud: sol,
          fechaFinal: parseSpanishDate(row[idxFechaFinal]),
          fechaInicio: parseSpanishDate(row[idxFechaInicio]),
          fechaFin: parseSpanishDate(row[idxFechaFin]),
          usuarioPT: row[idxUsuarioPT]==null ? '' : String(row[idxUsuarioPT]).trim(),
        });
      }
      logLine('GRAM EAM listo: ' + gramRows.length + ' filas válidas' + (skippedNoSolicitud? (' ('+skippedNoSolicitud+' sin N° de solicitud, omitidas)'):'') + '.', 'ok'); await tick();
      setLoading(
        35,
        'GRAM EAM procesado',
        gramRows.length.toLocaleString('es-PE') + ' Solicitudes de Trabajo válidas.'
      );
      
      await tick();


      // ---- STEP 2: read BASE DE SEGUIMIENTO (BD 2022-2026) ----
      setLoading(
        40,
        'Leyendo Base de Seguimiento',
        'Procesando las hojas BD 2022 – BD 2026…'
      );
      
      logLine('Leyendo Base de Seguimiento (hojas BD 2022–2026)…'); await tick();
      const wantedSheets = ['BD 2022','BD 2023','BD 2024','BD 2025','BD 2026'];

      const baseBuf = await baseFile.arrayBuffer();

const baseSheets = await workerRead(
  'readBase',
  baseBuf,
  wantedSheets
);

const availableSheets =
  wantedSheets.filter(n => baseSheets[n]);




      if(availableSheets.length===0) throw new Error('No se encontraron las hojas BD 2022–2026 en la Base de Seguimiento.');
      logLine('Hojas encontradas: ' + availableSheets.join(', ') + '.', 'ok'); await tick();
      setLoading(
        50,
        'Combinando bases históricas',
        availableSheets.join(', ') + ' encontradas.'
      );
      
      await tick();


  
      let combined = [];
      let baseHeaders = null;
      let idxStatus=-1, idxStPt=-1, idxCodSge=-1;

      for(const sheetName of availableSheets){

        const aoa = baseSheets[sheetName];
      
        if(!aoa || !aoa.length) continue;
      
        const hdr = aoa[0].map(normHeader);
      
        if(!baseHeaders){
      
          baseHeaders = hdr;
      
          idxStatus = findCol(hdr, ['STATUS']);
          idxStPt = findCol(hdr, ['ST/PT']);
          idxCodSge = findCol(hdr, ['COD_SGE','COD SGE']);
      
          if([idxStatus, idxStPt, idxCodSge].includes(-1)){
      
            throw new Error(
              'No se pudieron mapear las columnas STATUS / ST-PT / COD_SGE en la hoja ' +
              sheetName +
              '.'
            );
      
          }
      
        }
      
        let added = 0;
      
        for(let i=1; i<aoa.length; i++){
      
          const row = aoa[i];
      
          if(!row || row[0] == null) continue;
      
          combined.push(row);
      
          added++;
        }
      
        logLine(
          '  ' + sheetName +
          ': ' + added +
          ' filas copiadas hacia BD 2026 (combinado).',
          'ok'
        );
      
        await tick();
      }





      logLine('Total combinado BD 2022–2026: ' + combined.length + ' filas.', 'ok'); await tick();
      setLoading(
        60,
        'Base de Seguimiento cargada',
        combined.length.toLocaleString('es-PE') + ' filas combinadas.'
      );
      
      await tick();


    
      // ---- STEP 3: pivot-equivalent — filter STATUS = EN CURSO, keep numeric ST/PT only ----


    setLoading(
     65,
      'Filtrando ST "EN CURSO"',
      'Limpiando ST/PT y descartando valores vacíos o GYM…'
      );


      logLine('Filtrando STATUS = "EN CURSO" y limpiando columna ST/PT (fuera vacíos y "GYM…")…'); await tick();
      const enCursoMap = new Map();
      let enCursoTotal = 0, enCursoClean = 0;
      for(const row of combined){
        const status = normHeader(row[idxStatus]);
        if(status !== 'EN CURSO') continue;
        enCursoTotal++;
        const stptRaw = row[idxStPt];
        if(!isNumericST(stptRaw)) continue; // descarta vacíos y "GYM..."
        const stNum = toIntST(stptRaw);
        enCursoClean++;
        enCursoMap.set(stNum, row[idxCodSge]);
      }
      logLine('"EN CURSO": ' + enCursoTotal + ' filas → ' + enCursoClean + ' con ST numérica limpia (usadas para el cruce).', 'ok'); await tick();
      setLoading(
        78,
        'ST "EN CURSO" preparada',
        enCursoClean.toLocaleString('es-PE') + ' ST numéricas listas para el cruce.'
      );
      
      await tick();




      // ---- STEP 4: XLOOKUP-equivalent + filter out #N/A ----
      setLoading(
        82,
        'Cruzando Solicitudes de Trabajo',
        'Ejecutando el equivalente a BUSCARX…'
      );
      logLine('Cruzando cada Solicitud de Trabajo (BUSCARX) contra la lista limpia "EN CURSO"…'); await tick();
      for(const r of gramRows){
        if(enCursoMap.has(r.solicitud)){
          resultRows.push({
            codSge: enCursoMap.get(r.solicitud),
            solicitud: r.solicitud,
            fechaFinal: r.fechaFinal,//parche gaaaaaaaaaaaaaas
            fechaInicio: r.fechaInicio,
            fechaFin: r.fechaFin,
            usuarioPT: r.usuarioPT,
          });
        }
      }
      logLine('Filtro aplicado (se descartan los #N/A). Coincidencias finales: ' + resultRows.length + '.', 'ok'); await tick();
      setLoading(
        94,
        'Generando resultado',
        resultRows.length.toLocaleString('es-PE') + ' coincidencias encontradas.'
      );


      // ---- render ----
      els.statGramRows.textContent = gramRows.length.toLocaleString('es-PE');
      els.statBaseRows.textContent = combined.length.toLocaleString('es-PE');
      els.statEnCurso.textContent = enCursoClean.toLocaleString('es-PE');
      els.statMatch.textContent = resultRows.length.toLocaleString('es-PE');
      els.stats.style.display = 'grid';
  
      renderTable(resultRows);
      els.resultsHead.style.display = 'flex';
      els.tableWrap.style.display = resultRows.length ? 'block' : 'none';
      els.emptyMsg.style.display = resultRows.length ? 'none' : 'block';
  
      els.btnDownload.disabled = resultRows.length === 0;
      els.statusPill.textContent = 'listo · ' + resultRows.length + ' coincidencias';

      setLoading(
        100,
        'Proceso completado',
        'Resultado listo. Puedes revisar o descargar el Excel.'
      );
      
      els.statusPill.textContent =
        'listo · ' + resultRows.length + ' coincidencias';

      logLine('Proceso completo.', 'ok');
    }catch(err){
      console.error(err);
      logLine('Error: ' + err.message, 'err');
      els.statusPill.textContent = 'error';
    }finally{
     
  setTimeout(() => {
    hideLoading();
  }, 700);

  els.btnProcess.disabled = false;
    }
  }
  
  function fmtDate(d){
    if(!d) return '';
    const p = n=>String(n).padStart(2,'0');
    return p(d.getDate())+'/'+p(d.getMonth()+1)+'/'+d.getFullYear()+' '+p(d.getHours())+':'+p(d.getMinutes())+':'+p(d.getSeconds());
  }
  
  function renderTable(rows){
    els.tbody.innerHTML = '';
    const frag = document.createDocumentFragment();
    const MAX_RENDER = 3000;
    rows.slice(0, MAX_RENDER).forEach(r=>{
      const tr = document.createElement('tr');
      tr.innerHTML =
        '<td class="mono">'+(r.codSge??'')+'</td>'+
        '<td class="num">'+r.solicitud+'</td>'+
        '<td class="mono">'+fmtDate(r.fechaFinal)+'</td>'+
        '<td class="mono">'+fmtDate(r.fechaInicio)+'</td>'+
        '<td class="mono">'+fmtDate(r.fechaFin)+'</td>'+
        '<td>'+(r.usuarioPT||'')+'</td>';
      frag.appendChild(tr);
    });
    els.tbody.appendChild(frag);
    if(rows.length > MAX_RENDER){
      const tr = document.createElement('tr');
      tr.innerHTML = '<td colspan="6" style="text-align:center;color:var(--muted);padding:14px;">… mostrando primeras '+MAX_RENDER+' de '+rows.length+' filas. Descarga el Excel para ver todo.</td>';
      els.tbody.appendChild(tr);
    }
  }
  
  els.filterBox.addEventListener('input', ()=>{
    const q = els.filterBox.value.trim().toLowerCase();
    if(!q){ renderTable(resultRows); return; }
    const filtered = resultRows.filter(r =>
      String(r.codSge||'').toLowerCase().includes(q) ||
      String(r.solicitud).includes(q) ||
      String(r.usuarioPT||'').toLowerCase().includes(q)
    );
    renderTable(filtered);
  });
  
  function downloadResult(){
    const header = ['COD_SGE','Solicitud de trabajo','Fecha Final','Fecha Inicio','Fecha Fin','Usuario PT'];
    const aoa = [
      ['RESULTADO — ST finalizadas en GRAM EAM que figuran "EN CURSO" en Base de Seguimiento'],
      header,
    ];
    resultRows.forEach(r=>{
      aoa.push([
        r.codSge ?? '',
        r.solicitud,
        r.fechaFinal ? new Date(r.fechaFinal.getTime() + 36 * 1000) : '',
        r.fechaInicio ?? '',
        r.fechaFin ?? '',
        r.usuarioPT ?? '',
      ]);
    });
    const ws = XLSX.utils.aoa_to_sheet(aoa, { cellDates: true });

    
    ws['!cols'] = [{wch:20},{wch:18},{wch:20},{wch:20},{wch:20},{wch:22}];
    // date formatting for columns C,D,E from row 3 onward
    const dateCols = [2,3,4]; // 0-indexed: C,D,E
    for(let r=2; r<aoa.length; r++){
      dateCols.forEach(c=>{
        const addr = XLSX.utils.encode_cell({r, c});
        const cell = ws[addr];
        if(cell && cell.v instanceof Date){
          cell.t = 'd';
          cell.z = 'dd/mm/yyyy hh:mm:ss';
        }
      });
    }
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Resultado ST-PT');
    const stamp = new Date().toISOString().slice(0,10);
    XLSX.writeFile(wb, 'Resultado_ST_PT_' + stamp + '.xlsx');
  }
  
  els.btnProcess.addEventListener('click', processFiles);
  els.btnDownload.addEventListener('click', downloadResult);
  els.btnReset.addEventListener('click', ()=>{
    gramFile = null; baseFile = null; resultRows = [];
    els.fileGram.value=''; els.fileBase.value='';
    els.fnameGram.textContent=''; els.fnameBase.textContent='';
    els.fmetaGram.textContent=''; els.fmetaBase.textContent='';
    els.dropGram.classList.remove('filled'); els.dropBase.classList.remove('filled');
    els.log.innerHTML=''; els.stats.style.display='none';
    els.resultsHead.style.display='none'; els.tableWrap.style.display='none'; els.emptyMsg.style.display='none';
    els.btnDownload.disabled = true; els.btnProcess.disabled = true;
    els.statusPill.textContent = 'esperando archivos';
    els.filterBox.value='';
  });