// ============================================================
// BALANCEO DE CARGAS WO v6 — Programacion CQ, Grupo Peñaflor
// ============================================================

const state = {
  planData: [], entregaData: [], insumoData: {},
  planRawRows: {}, planHeaders: [],
  semanaMap: {}, expanded: {},
  movedToN2: {}, movedToN1: {},
  notes: {}, historial: [],
  editedCodEq: {}, editedEntregas: {},
  filtroSemana: 'Todos',
  showChart: false, showCompare: false,
  config: {
    semanaInicio: '',
    capacidades: { N:0, 'N+1':0, 'N+2':0, 'N+3':0 },
    minPct: 80, maxPct: 100,
    merma: 1000, factorCajas: 9
  },
  importMode: 'plan', rawHeaders: [], rawData: [], colMap: {},
  nextId: 1,
  planLoaded: false, entregaLoaded: false, insumoLoaded: false
};

// ── LOCALSTORAGE ──────────────────────────────────────────
function toJSON(obj) {
  return JSON.stringify(obj, (k,v) => v instanceof Date ? {__d: v.toISOString()} : v);
}
function fromJSON(str) {
  return JSON.parse(str, (k,v) => (v && v.__d) ? new Date(v.__d) : v);
}
function autoSave() {
  try {
    localStorage.setItem('bwo-v6', toJSON({
      planData: state.planData, entregaData: state.entregaData,
      insumoData: state.insumoData, planRawRows: state.planRawRows,
      planHeaders: state.planHeaders, semanaMap: state.semanaMap,
      movedToN2: state.movedToN2, movedToN1: state.movedToN1,
      notes: state.notes, historial: state.historial,
      editedCodEq: state.editedCodEq, editedEntregas: state.editedEntregas,
      config: state.config,
      planLoaded: state.planLoaded, entregaLoaded: state.entregaLoaded,
      insumoLoaded: state.insumoLoaded, ts: new Date().toISOString()
    }));
    const t = new Date().toLocaleTimeString('es-AR',{hour:'2-digit',minute:'2-digit'});
    setSaveIndicator('Guardado ' + t);
  } catch(e) { console.warn('Save error', e); }
}
function manualSave() {
  autoSave();
  const btn = document.getElementById('btn-save');
  if (btn) { btn.textContent = '✓ Guardado'; setTimeout(() => btn.textContent = '💾 Guardar', 2000); }
}
function loadSession() {
  try {
    const raw = localStorage.getItem('bwo-v6');
    if (!raw) return null;
    const s = fromJSON(raw);
    const fields = ['planData','entregaData','insumoData','planRawRows','planHeaders',
      'semanaMap','movedToN2','movedToN1','notes','historial','editedCodEq',
      'editedEntregas','planLoaded','entregaLoaded','insumoLoaded'];
    fields.forEach(f => { if (s[f] !== undefined) state[f] = s[f]; });
    if (s.config) Object.assign(state.config, s.config);
    return s.ts || null;
  } catch(e) { return null; }
}
function setSaveIndicator(msg) {
  const el = document.getElementById('save-indicator');
  if (el) el.textContent = msg;
}

// ── INIT ──────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  const savedTs = loadSession();
  renderConfig();
  renderSemanaFilter();

  if (savedTs) {
    if (state.planLoaded)    updateStatus('plan',     state.planData.length + ' filas (sesion)');
    if (state.entregaLoaded) updateStatus('entregas', state.entregaData.length + ' filas (sesion)');
    if (state.insumoLoaded)  updateStatus('insumo',   Object.keys(state.insumoData).length + ' WOs (sesion)');
    setSaveIndicator('Sesion del ' + new Date(savedTs).toLocaleDateString('es-AR'));
    renderConfig();
    render();
  }

  document.getElementById('filePlan')
    .addEventListener('change', e => handleFileUpload(e,'plan'));
  document.getElementById('fileEntrega')
    .addEventListener('change', e => handleFileUpload(e,'entregas'));
  document.getElementById('fileInsumo')
    .addEventListener('change', e => handleFileUpload(e,'insumo'));

  document.querySelectorAll('.tab').forEach(tab => {
    tab.addEventListener('click', () => {
      document.querySelectorAll('.tab').forEach(t => t.classList.remove('on'));
      document.querySelectorAll('.panel').forEach(p => p.classList.remove('on'));
      tab.classList.add('on');
      const panel = document.getElementById('panel-' + tab.dataset.tab);
      if (panel) panel.classList.add('on');
      if (tab.dataset.tab === 'historial') renderHistorial();
    });
  });

  // Sincronizacion en tiempo real con bodegas.html
  window.addEventListener('storage', e => {
    if (e.key === 'bwo-shared-entregas' && (state.planLoaded || state.entregaLoaded)) {
      render();
    }
  });
});

// ── FILE IMPORT ───────────────────────────────────────────
function handleFileUpload(e, mode) {
  const file = e.target.files[0];
  if (!file) return;
  state.importMode = mode;
  const reader = new FileReader();
  reader.onload = function(evt) {
    const wb    = XLSX.read(evt.target.result, { type:'binary' });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const rows  = XLSX.utils.sheet_to_json(sheet, { header:1, defval:'', raw:true });
    if (rows.length < 2) { alert('Archivo vacío.'); return; }
    state.rawHeaders = rows[0].map(h => String(h).trim());
    state.rawData    = rows.slice(1).filter(r => r.some(c => c !== ''));
    showColumnMapper(state.rawHeaders, autoDetect(state.rawHeaders, mode), mode);
  };
  reader.readAsBinaryString(file);
  e.target.value = '';
}

function autoDetect(headers, mode) {
  const m = { codEq:'', semana:'', lts:'', wo:'', estado:'', fecha:'', etiqueta:'', caja:'' };
  headers.forEach(h => {
    const hl = h.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    if (!m.codEq    && /codeq|equivalente/.test(hl))      m.codEq    = h;
    if (!m.semana   && /\bsemana\b/.test(hl))              m.semana   = h;
    if (!m.lts      && /^lts$/.test(hl))                   m.lts      = h;
    if (!m.wo       && /^orden$/.test(hl))                 m.wo       = h;
    if (!m.estado   && /^estado$/.test(hl))                m.estado   = h;
    if (!m.fecha    && /feinicio|fechainicio/.test(hl))    m.fecha    = h;
    if (!m.etiqueta && /etiqueta/.test(hl))                m.etiqueta = h;
    if (!m.caja     && /\bcaja\b/.test(hl))                m.caja     = h;
  });
  if (!m.codEq)    m.codEq    = headers.find(h => /codeq/i.test(h))    || '';
  if (!m.semana)   m.semana   = headers.find(h => /semana/i.test(h))   || '';
  if (!m.lts)      m.lts      = headers.find(h => /^lts$/i.test(h))    || '';
  if (!m.wo)       m.wo       = headers.find(h => /orden/i.test(h))    || '';
  if (!m.etiqueta) m.etiqueta = headers.find(h => /etiqueta/i.test(h)) || '';
  if (!m.caja)     m.caja     = headers.find(h => /caja/i.test(h))     || '';
  return m;
}

function showColumnMapper(headers, detected, mode) {
  const byMode = {
    plan: [
      { key:'codEq',  label:'CodEq',             req:true  },
      { key:'semana', label:'Semana',             req:true  },
      { key:'lts',    label:'Volumen (Lts)',       req:true  },
      { key:'wo',     label:'Nro. WO (Orden)',    req:true  },
      { key:'estado', label:'Estado',             req:false },
      { key:'fecha',  label:'Fecha inicio',       req:false }
    ],
    entregas: [
      { key:'codEq',  label:'CodEq',             req:true },
      { key:'semana', label:'Semana',             req:true },
      { key:'lts',    label:'Volumen (Lts)',       req:true }
    ],
    insumo: [
      { key:'wo',       label:'Nro. WO (Orden)', req:true },
      { key:'etiqueta', label:'Fecha etiqueta',  req:true },
      { key:'caja',     label:'Fecha caja',      req:true }
    ]
  };
  const titles = { plan:'Columnas del Plan', entregas:'Columnas de Entregas', insumo:'Columnas de Insumos' };
  document.getElementById('modal-title').textContent    = titles[mode];
  document.getElementById('modal-desc').textContent     = 'Mapeá las columnas del archivo:';
  document.getElementById('modal-confirm-btn').textContent = 'Confirmar e importar';
  document.getElementById('modal-confirm-btn').onclick  = confirmImport;
  document.getElementById('col-map-fields').innerHTML   =
    (byMode[mode]||[]).map(f => `
      <div class="col-map-row">
        <label>${f.label}${f.req ? '' : ' <span style="color:#bbb;font-size:11px">(opc.)</span>'}</label>
        <select id="map-${f.key}">
          ${!f.req ? '<option value="">— no usar —</option>' : '<option value="">— seleccionar —</option>'}
          ${headers.map(h => `<option value="${h}"${h === detected[f.key] ? ' selected' : ''}>${h}</option>`).join('')}
        </select>
      </div>`).join('');
  document.getElementById('modal-overlay').classList.remove('hidden');
}

function confirmImport() {
  const mode = state.importMode;
  const req  = { plan:['codEq','semana','lts','wo'], entregas:['codEq','semana','lts'], insumo:['wo','etiqueta','caja'] };
  const m    = {};
  ['codEq','semana','lts','wo','estado','fecha','etiqueta','caja'].forEach(f => {
    const el = document.getElementById('map-'+f);
    if (el) m[f] = el.value;
  });
  if ((req[mode]||[]).some(f => !m[f])) { alert('Asigná todas las columnas obligatorias.'); return; }
  cerrarModal();
  if (mode === 'insumo') {
    state.insumoData   = processInsumoData(state.rawData, m);
    state.insumoLoaded = true;
    updateStatus('insumo', Object.keys(state.insumoData).length + ' WOs cargadas');
  } else {
    const processed = processData(state.rawData, m);
    if (mode === 'plan') {
      state.planData = processed; state.planLoaded = true;
      state.movedToN2 = {}; state.movedToN1 = {}; state.expanded = {};
      updateStatus('plan', processed.length + ' filas cargadas');
    } else {
      state.entregaData = processed; state.entregaLoaded = true;
      updateStatus('entregas', processed.length + ' filas cargadas');
    }
    buildSemanaMap();
  }
  autoSave(); render();
}

function cerrarModal() {
  document.getElementById('modal-overlay').classList.add('hidden');
}

// ── PROCESS DATA ──────────────────────────────────────────
function processData(rows, m) {
  const H  = state.rawHeaders;
  const ix = k => (m[k] ? H.indexOf(m[k]) : -1);
  if (state.importMode === 'plan') {
    state.planHeaders = H.slice();
    state.planRawRows = {};
  }
  return rows.map((r,i) => {
    const woRaw = ix('wo') >= 0 ? r[ix('wo')] : '';
    const wo    = woRaw instanceof Date ? String(woRaw.getTime())
                : String(woRaw || 'WO-'+(i+1)).trim();
    if (state.importMode === 'plan') state.planRawRows[wo] = r;
    return {
      id:        state.nextId++, wo,
      codEq:     String(ix('codEq')  >= 0 ? r[ix('codEq')]  : '').trim(),
      semanaRaw: String(ix('semana') >= 0 ? r[ix('semana')] : '').trim(),
      lts:       parseFloat(ix('lts') >= 0 ? r[ix('lts')] : 0) || 0,
      estado:    String(ix('estado') >= 0 ? r[ix('estado')] : '').trim(),
      fecha:     fmtFecha(ix('fecha') >= 0 ? r[ix('fecha')] : '')
    };
  }).filter(r => r.codEq && r.lts > 0);
}

function processInsumoData(rows, m) {
  const H  = state.rawHeaders;
  const ix = k => (m[k] ? H.indexOf(m[k]) : -1);
  const res = {};
  rows.forEach(r => {
    const woRaw = ix('wo') >= 0 ? r[ix('wo')] : '';
    const wo    = woRaw instanceof Date ? String(woRaw.getTime()) : String(woRaw||'').trim();
    if (!wo) return;
    res[wo] = {
      etiqueta: parseExcelDate(ix('etiqueta') >= 0 ? r[ix('etiqueta')] : ''),
      caja:     parseExcelDate(ix('caja')     >= 0 ? r[ix('caja')]     : '')
    };
  });
  return res;
}

// ── DATE UTILS ────────────────────────────────────────────
function parseExcelDate(val) {
  if (!val && val !== 0) return null;
  if (val instanceof Date) return isNaN(val) ? null : val;
  const s = String(val).trim();
  if (!s) return null;
  const n = parseFloat(s);
  if (!isNaN(n) && n > 1000 && n < 100000)
    return new Date((n - 25569) * 86400000);
  const dmy = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (dmy) return new Date(+dmy[3], +dmy[2]-1, +dmy[1]);
  const ymd = s.match(/^(\d{4})[\/\-](\d{2})[\/\-](\d{2})/);
  if (ymd) return new Date(+ymd[1], +ymd[2]-1, +ymd[3]);
  return null;
}
function getSemanaStartDate(semLabel) {
  if (!state.config.semanaInicio) return null;
  const base   = new Date(state.config.semanaInicio + 'T00:00:00');
  const offset = {N:0,'N+1':7,'N+2':14,'N+3':21}[semLabel];
  if (offset === undefined) return null;
  const d = new Date(base);
  d.setDate(d.getDate() + offset);
  return d;
}
function getSemaforoColor(woStr, semLabel) {
  const ins = state.insumoData[woStr];
  if (!ins) return 'none';
  const ss   = getSemanaStartDate(semLabel);
  const isOk = d => !d || (ss ? d <= ss : true);
  const eOk  = isOk(ins.etiqueta), cOk = isOk(ins.caja);
  return (eOk && cOk) ? 'green' : (eOk || cOk) ? 'yellow' : 'red';
}
function getSemaforoText(woStr, semLabel) {
  const ins = state.insumoData[woStr];
  if (!ins) return '';
  const ss   = getSemanaStartDate(semLabel);
  const isOk = d => !d || (ss ? d <= ss : true);
  const fmt  = d => d ? fmtDateShort(d) : 'sin fecha';
  return `Etiq: ${fmt(ins.etiqueta)}${isOk(ins.etiqueta)?'✓':'✗'} · Caja: ${fmt(ins.caja)}${isOk(ins.caja)?'✓':'✗'}`;
}
function getArrivalInfo(woStr, semLabel) {
  const ins = state.insumoData[woStr];
  if (!ins) return null;
  const ss   = getSemanaStartDate(semLabel);
  const isOk = d => !d || (ss ? d <= ss : true);
  const bad  = [];
  if (!isOk(ins.etiqueta)) bad.push('Etiq: '+(ins.etiqueta?fmtDateShort(ins.etiqueta):'?'));
  if (!isOk(ins.caja))     bad.push('Caja: '+(ins.caja?fmtDateShort(ins.caja):'?'));
  if (bad.length) return { text: bad.join(' · '), ok: false };
  const ok = [];
  if (ins.etiqueta) ok.push('Etiq: '+fmtDateShort(ins.etiqueta));
  if (ins.caja)     ok.push('Caja: '+fmtDateShort(ins.caja));
  return { text: ok.length ? ok.join(' · ') : 'Sin fecha (OK)', ok: true };
}

// ── WEEK NORMALIZATION ────────────────────────────────────
function buildSemanaMap() {
  const all    = new Set([...state.planData.map(r=>r.semanaRaw),...state.entregaData.map(r=>r.semanaRaw)]);
  const sorted = [...all].filter(Boolean).sort((a,b) => {
    const na=parseFloat(a),nb=parseFloat(b);
    return !isNaN(na)&&!isNaN(nb) ? na-nb : a.localeCompare(b);
  });
  const labels = ['N','N+1','N+2','N+3'];
  state.semanaMap = {};
  sorted.forEach((w,i) => { state.semanaMap[w] = labels[i]||'N+3'; });
  const txt = sorted.map((w,i) => `${labels[i]||'N+3'}=sem.${w}`).join(' · ');
  document.getElementById('status-semanas').textContent = txt;
}

function updateStatus(type, text) {
  const map    = { plan:'status-plan', entregas:'status-entrega', insumo:'status-insumo' };
  const labels = { plan:'Plan: ', entregas:'Entregas: ', insumo:'Insumos: ' };
  const el     = document.getElementById(map[type]);
  if (!el) return;
  el.querySelector('.status-dot').className = 'status-dot dot-on';
  el.querySelector('span:last-child').textContent = labels[type] + text;
}

// ── MOVEMENTS ────────────────────────────────────────────
function getCodEq(row) {
  return state.editedCodEq[row.id] || row.codEq;
}
function getSemana(row) {
  if (state.movedToN1[row.id]) return 'N+1';
  if (state.movedToN2[row.id]) return 'N+2';
  return state.semanaMap[row.semanaRaw] || null;
}
function getOrigSemana(row) {
  return state.semanaMap[row.semanaRaw] || null;
}
function logMovimiento(wo, codEq, de, a, accion) {
  state.historial.unshift({ ts: new Date().toISOString(), wo, codEq, de, a, accion });
  if (state.historial.length > 150) state.historial.pop();
  const el = document.getElementById('hist-badge');
  if (el) el.textContent = state.historial.length;
}
function moveToN2(woId) {
  const w = state.planData.find(r => r.id === woId);
  state.movedToN2[woId] = true;
  if (w) logMovimiento(w.wo, getCodEq(w), 'N+3', 'N+2', 'adelantar');
  autoSave(); render();
}
function moveToN1(woId) {
  const w = state.planData.find(r => r.id === woId);
  state.movedToN1[woId] = true;
  if (w) logMovimiento(w.wo, getCodEq(w), 'N+2', 'N+1', 'adelantar');
  autoSave(); render();
}
function revertirA(woId) {
  const w = state.planData.find(r => r.id === woId);
  if (!w) return;
  if (state.movedToN1[woId]) {
    logMovimiento(w.wo, getCodEq(w), 'N+1', 'N+2', 'revertir');
    delete state.movedToN1[woId];
  } else if (state.movedToN2[woId]) {
    logMovimiento(w.wo, getCodEq(w), 'N+2', 'N+3', 'revertir');
    delete state.movedToN2[woId];
  }
  autoSave(); render();
}

// ── INLINE EDITS ──────────────────────────────────────────
function editNote(woId, woNum, evt) {
  if (evt) evt.stopPropagation();
  openModal('Nota — WO ' + woNum,
    'La nota es opcional y visible al expandir la WO.',
    `<textarea id="note-ta" style="width:100%;height:90px;padding:8px 10px;border:.5px solid #d0d0c8;border-radius:6px;font-size:13px;font-family:inherit;resize:vertical" placeholder="Escribe una nota...">${state.notes[woId]||''}</textarea>`,
    'Guardar nota',
    () => {
      const val = (document.getElementById('note-ta')?.value || '').trim();
      if (val) state.notes[woId] = val; else delete state.notes[woId];
      cerrarModal(); autoSave(); render();
    }
  );
}
function editCodEq(woId, woNum, cur, evt) {
  if (evt) evt.stopPropagation();
  openModal('Editar CodEq — WO ' + woNum, '',
    `<div class="col-map-row"><label>Codigo equivalente</label>
     <input id="codeq-in" type="text" value="${cur}" style="flex:1;padding:7px 10px;border:.5px solid #d0d0c8;border-radius:6px;font-size:13px"></div>`,
    'Guardar',
    () => {
      const val = (document.getElementById('codeq-in')?.value || '').trim();
      if (val) state.editedCodEq[woId] = val;
      cerrarModal(); autoSave(); render();
    }
  );
}
function editEntrega(cod, sem, cur, evt) {
  if (evt) evt.stopPropagation();
  const key = cod + '__' + sem;
  openModal('Editar entrega — ' + cod + ' · Sem. ' + sem, '',
    `<div class="col-map-row"><label>Volumen de entrega (Lts)</label>
     <input id="entr-in" type="number" value="${cur}" style="flex:1;padding:7px 10px;border:.5px solid #d0d0c8;border-radius:6px;font-size:13px"></div>`,
    'Guardar',
    () => {
      const val = parseFloat(document.getElementById('entr-in')?.value) || 0;
      if (val > 0) state.editedEntregas[key] = val; else delete state.editedEntregas[key];
      cerrarModal(); autoSave(); render();
    }
  );
}
function openModal(title, desc, fields, btnLabel, onConfirm) {
  document.getElementById('modal-title').textContent    = title;
  document.getElementById('modal-desc').textContent     = desc;
  document.getElementById('col-map-fields').innerHTML   = fields;
  document.getElementById('modal-confirm-btn').textContent = btnLabel;
  document.getElementById('modal-confirm-btn').onclick  = onConfirm;
  document.getElementById('modal-overlay').classList.remove('hidden');
}

// ── EXPORT ────────────────────────────────────────────────
function showExportModal() {
  const movedIds = [...Object.keys(state.movedToN2),...Object.keys(state.movedToN1)].map(Number);
  if (!movedIds.length) { alert('No hay WOs movidas para exportar.'); return; }
  if (!state.planHeaders.length) { alert('Primero importá el Plan.'); return; }
  const autoC = state.planHeaders.find(h => /cantordenada/i.test(h)) || '';
  openModal('Exportar WOs balanceadas',
    movedIds.length + ' WOs movidas. ¿Cuál columna es CantOrdenadaPT?',
    `<div class="col-map-row"><label>CantOrdenadaPT</label>
     <select id="map-export-cant">
       <option value="">— seleccionar —</option>
       ${state.planHeaders.map(h=>`<option value="${h}"${h===autoC?' selected':''}>${h}</option>`).join('')}
     </select></div>`,
    'Exportar Excel', doExport
  );
}
function doExport() {
  const cantCol = document.getElementById('map-export-cant')?.value;
  if (!cantCol) { alert('Seleccioná la columna CantOrdenadaPT.'); return; }
  cerrarModal();
  const cantIdx  = state.planHeaders.indexOf(cantCol);
  const movedIds = new Set([
    ...Object.keys(state.movedToN2).map(Number),
    ...Object.keys(state.movedToN1).map(Number)
  ]);
  const wos = state.planData.filter(r => movedIds.has(r.id));
  const wsData = [
    ['Orden','CodEq','Semana original','Semana nueva','Lts','Cajas','Estado','Fecha inicio','CantOrdenadaPT'],
    ...wos.map(r => {
      const rawRow = state.planRawRows[r.wo] || [];
      const cant   = cantIdx >= 0 ? rawRow[cantIdx] : '';
      const cajas  = state.config.factorCajas > 0 ? Math.round(r.lts / state.config.factorCajas) : '';
      return [r.wo, getCodEq(r), getOrigSemana(r)||r.semanaRaw, getSemana(r)||'N+2', r.lts, cajas, r.estado, r.fecha, cant];
    })
  ];
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(wsData);
  ws['!cols'] = [14,16,14,12,10,8,8,14,14].map(w=>({wch:w}));
  XLSX.utils.book_append_sheet(wb, ws, 'WOs Balanceadas');
  XLSX.writeFile(wb, 'WOs_balanceadas.xlsx');
}

// ── DEMO DATA ─────────────────────────────────────────────
function cargarDemoData() {
  state.planData = [
    {id:1,  wo:'1129121',codEq:'C1079',       semanaRaw:'22',lts:7141.5,estado:'40',fecha:'25/5/2026'},
    {id:2,  wo:'1134019',codEq:'C1079',       semanaRaw:'22',lts:7012.1,estado:'40',fecha:'25/5/2026'},
    {id:3,  wo:'1134020',codEq:'A1039/A1497', semanaRaw:'22',lts:5200.0,estado:'40',fecha:'25/5/2026'},
    {id:4,  wo:'1134021',codEq:'A1039/A1497', semanaRaw:'23',lts:5100.0,estado:'40',fecha:'2/6/2026'},
    {id:5,  wo:'1134022',codEq:'C1079',       semanaRaw:'23',lts:8500.0,estado:'40',fecha:'2/6/2026'},
    {id:6,  wo:'1134023',codEq:'B2201',       semanaRaw:'23',lts:3800.0,estado:'40',fecha:'2/6/2026'},
    {id:7,  wo:'1134024',codEq:'A1039/A1497', semanaRaw:'24',lts:5050.0,estado:'20',fecha:'9/6/2026'},
    {id:8,  wo:'1134025',codEq:'B2201',       semanaRaw:'24',lts:3600.0,estado:'20',fecha:'9/6/2026'},
    {id:9,  wo:'1134026',codEq:'C1079',       semanaRaw:'25',lts:7200.0,estado:'20',fecha:'16/6/2026'},
    {id:10, wo:'1134027',codEq:'D3310',       semanaRaw:'25',lts:4100.0,estado:'20',fecha:'16/6/2026'},
    {id:11, wo:'1134028',codEq:'A1039/A1497', semanaRaw:'25',lts:4900.0,estado:'20',fecha:'16/6/2026'},
  ];
  state.entregaData = [
    {id:20,codEq:'C1079',       semanaRaw:'22',lts:16000.0},
    {id:21,codEq:'A1039/A1497', semanaRaw:'22',lts:4800.0},
    {id:22,codEq:'A1039/A1497', semanaRaw:'23',lts:5300.0},
    {id:23,codEq:'C1079',       semanaRaw:'23',lts:7200.0},
    {id:24,codEq:'B2201',       semanaRaw:'23',lts:3600.0},
    {id:25,codEq:'A1039/A1497', semanaRaw:'24',lts:9500.0},
    {id:26,codEq:'B2201',       semanaRaw:'24',lts:3400.0},
    {id:27,codEq:'C1079',       semanaRaw:'25',lts:7500.0},
    {id:28,codEq:'D3310',       semanaRaw:'25',lts:3200.0},
  ];
  state.insumoData = {
    '1134026': { etiqueta: new Date(2026,5,20), caja: new Date(2026,5,20) },
    '1134027': { etiqueta: null,                caja: new Date(2026,5,20) },
    '1134028': { etiqueta: null,                caja: null                },
  };
  state.planLoaded=true; state.entregaLoaded=true; state.insumoLoaded=true;
  state.movedToN2={}; state.movedToN1={}; state.expanded={}; state.notes={};
  state.historial=[]; state.editedCodEq={}; state.editedEntregas={};
  state.planHeaders = ['Orden','CodEq','Semana','Lts','Estado','FeInicio','CantOrdenadaPT'];
  state.planRawRows = {};
  state.planData.forEach(r => {
    state.planRawRows[r.wo] = [r.wo,r.codEq,r.semanaRaw,r.lts,r.estado,r.fecha,Math.round(r.lts/6/0.75)];
  });
  state.nextId = 60;
  if (!state.config.semanaInicio)   state.config.semanaInicio = '2026-05-25';
  if (!state.config.capacidades.N) state.config.capacidades = {N:160000,'N+1':160000,'N+2':160000,'N+3':160000};
  updateStatus('plan','11 filas (demo)');
  updateStatus('entregas','9 filas (demo)');
  updateStatus('insumo','3 WOs (demo)');
  buildSemanaMap(); renderConfig(); autoSave(); render();
}

// ── PIVOT ─────────────────────────────────────────────────
function buildPivot() {
  const pivot = {};
  state.planData.forEach(r => {
    const sem = getSemana(r);
    if (!sem) return;
    const cod = getCodEq(r);
    if (!pivot[cod]) pivot[cod] = {};
    if (!pivot[cod][sem]) pivot[cod][sem] = { planVol:0, entregaVol:0, wos:[] };
    pivot[cod][sem].planVol += r.lts;
    pivot[cod][sem].wos.push({ ...r, semLabel:sem, codEqEff:cod });
  });
  // Merma por CodEq por semana
  Object.keys(pivot).forEach(cod =>
    Object.keys(pivot[cod]).forEach(sem => { pivot[cod][sem].planVol += state.config.merma; })
  );
  // Entregas desde Excel
  state.entregaData.forEach(r => {
    const sem = state.semanaMap[r.semanaRaw];
    if (!sem || !pivot[r.codEq]) return;
    if (!pivot[r.codEq][sem]) pivot[r.codEq][sem] = { planVol:0, entregaVol:0, wos:[] };
    pivot[r.codEq][sem].entregaVol += r.lts;
  });
  // Volúmenes ingresados desde bodegas.html (sincronizacion en tiempo real)
  try {
    const sharedRaw = localStorage.getItem('bwo-shared-entregas');
    if (sharedRaw) {
      const shared = JSON.parse(sharedRaw);
      Object.entries(shared).forEach(([key, vol]) => {
        const [cod, sem] = key.split('|||');
        if (pivot[cod]?.[sem] && parseFloat(vol) > 0) {
          pivot[cod][sem].entregaVol = parseFloat(vol);
          pivot[cod][sem].fromBodega = true;
        }
      });
    }
  } catch(e) {}
  // Ediciones manuales del mismo Balanceo
  Object.entries(state.editedEntregas).forEach(([key, val]) => {
    const [cod, sem] = key.split('__');
    if (pivot[cod]?.[sem]) {
      pivot[cod][sem].entregaVol = val;
      pivot[cod][sem].fromBodega = false;
    }
  });
  // Calcular %
  Object.values(pivot).forEach(cd =>
    Object.values(cd).forEach(cell => {
      cell.pct = cell.entregaVol > 0 ? Math.round(cell.planVol / cell.entregaVol * 100) : null;
    })
  );
  return pivot;
}

// ── RENDER MASTER ─────────────────────────────────────────
function render() {
  const pivot = buildPivot();
  renderMetrics(pivot);
  renderSemanaFilter();
  renderPivotTable(pivot);
  renderAlertas(pivot);
  const hb = document.getElementById('hist-badge');
  if (hb) hb.textContent = state.historial.length;
}

// ── METRICS ───────────────────────────────────────────────
function renderMetrics(pivot) {
  let totalPlanV = 0;
  Object.values(pivot).forEach(cd => Object.values(cd).forEach(c => totalPlanV += c.planVol));
  const totalCap = Object.values(state.config.capacidades).reduce((a,v) => a+(+v||0), 0);
  const utilPct  = totalCap > 0 ? Math.round(totalPlanV / totalCap * 100) : 0;
  const utilCls  = utilPct >= 95 ? 'danger' : utilPct >= 70 ? 'ok' : 'warn';
  const alerts   = computeAlerts(pivot);
  const na       = alerts.dups.length + alerts.deficit.length + alerts.surplus.length;
  let ipCnt = 0;
  if (state.insumoLoaded) {
    state.planData.forEach(r => {
      const sem = getSemana(r);
      if (sem) { const c = getSemaforoColor(r.wo, sem); if (c==='red'||c==='yellow') ipCnt++; }
    });
  }
  document.getElementById('m-util').textContent   = state.planLoaded ? utilPct+'%' : '—';
  document.getElementById('m-util').className     = 'metric-value ' + (state.planLoaded ? utilCls : '');
  document.getElementById('m-lts').textContent    = state.planLoaded ? fmtLts(totalPlanV) : '—';
  document.getElementById('m-alerts').textContent = (state.planLoaded||state.entregaLoaded) ? na : '—';
  document.getElementById('m-alerts').className   = 'metric-value ' + (na>0?'warn':'ok');
  document.getElementById('m-insumo').textContent = state.insumoLoaded ? ipCnt : '—';
  document.getElementById('m-insumo').className   = 'metric-value ' + (ipCnt>0?'danger':'ok');
  document.getElementById('alert-badge').textContent = na;
}

// ── FILTER ────────────────────────────────────────────────
function renderSemanaFilter() {
  document.getElementById('semana-filter').innerHTML =
    ['Todos','N','N+1','N+2','N+3'].map(p =>
      `<button class="fpill${p==='N+3'?' n3':''}${p===state.filtroSemana?' on':''}"
         onclick="setFiltro('${p}')">${p==='Todos'?'Todas las semanas':'Semana '+p}</button>`
    ).join('');
}
function setFiltro(val) {
  state.filtroSemana = val;
  renderSemanaFilter();
  renderPivotTable(buildPivot());
}

// ── CHART & COMPARE ───────────────────────────────────────
function toggleChart() {
  state.showChart = !state.showChart;
  document.getElementById('btn-chart')?.classList.toggle('active', state.showChart);
  renderPivotTable(buildPivot());
}
function toggleCompare() {
  state.showCompare = !state.showCompare;
  document.getElementById('btn-compare')?.classList.toggle('active', state.showCompare);
  renderPivotTable(buildPivot());
}

function renderChartHTML(pivot) {
  const semanas = ['N','N+1','N+2','N+3'];
  const caps    = state.config.capacidades;
  const totalCap = semanas.reduce((a,s) => a+(+caps[s]||0), 0);
  if (!totalCap) return '<div style="font-size:12px;color:#aaa;padding:8px">Configura las capacidades para ver el gráfico.</div>';
  const bars = semanas.map(s => {
    const cap   = +caps[s]||0;
    const planV = Object.values(pivot).reduce((a,cd) => a+(cd[s]?cd[s].planVol:0), 0);
    const pct   = cap > 0 ? Math.round(planV/cap*100) : 0;
    const col   = pct>=95?'#E24B4A':pct>=75?'#EF9F27':'#1D9E75';
    const cajas = state.config.factorCajas>0 ? ` / ${fmtNum(Math.round(planV/state.config.factorCajas))} cj` : '';
    return `<div class="chart-row">
      <div class="chart-lbl">Sem. ${s}</div>
      <div class="chart-bar-outer"><div class="chart-fill" style="width:${Math.min(pct,100)}%;background:${col}"></div></div>
      <div class="chart-pct" style="color:${col}">${pct}%</div>
      <div class="chart-nums">${fmtLts(planV)}${cajas}</div>
    </div>`;
  }).join('');
  const totalPlanV = semanas.reduce((a,s) =>
    a+Object.values(pivot).reduce((b,cd) => b+(cd[s]?cd[s].planVol:0), 0), 0);
  const tPct = totalCap>0 ? Math.round(totalPlanV/totalCap*100) : 0;
  const tCol = tPct>=95?'#E24B4A':tPct>=75?'#EF9F27':'#1D9E75';
  return `<div class="chart-wrap">
    <div class="chart-hdr">
      <span class="chart-ttl">Utilizacion de fraccionamiento</span>
      <span class="chart-total-val" style="color:${tCol}">${tPct}% total &nbsp;|&nbsp; ${fmtLts(totalPlanV)} Lts</span>
    </div>${bars}
  </div>`;
}

function renderCompareHTML() {
  const moved = state.planData.filter(r => state.movedToN1[r.id]||state.movedToN2[r.id]);
  if (!moved.length) return '<div style="font-size:12px;color:#aaa;padding:8px">No hay WOs movidas aún.</div>';
  return `<div class="compare-wrap">
    <div class="compare-ttl">Plan balanceado — ${moved.length} WO${moved.length>1?'s':''} movidas</div>
    <table class="compare-table">
      <thead><tr><th>WO</th><th>Codigo vino</th><th>Semana original</th><th>Semana nueva</th></tr></thead>
      <tbody>
        ${moved.map(r => `<tr>
          <td class="wo-num">${r.wo}</td>
          <td>${getCodEq(r)}</td>
          <td><span class="sem-badge sem-orig">${getOrigSemana(r)||r.semanaRaw}</span></td>
          <td><span class="sem-badge sem-moved">&#8593; ${getSemana(r)}</span></td>
        </tr>`).join('')}
      </tbody>
    </table>
  </div>`;
}

// ── PIVOT TABLE ───────────────────────────────────────────
function renderPivotTable(pivot) {
  const semanas = ['N','N+1','N+2','N+3'];
  let codigos   = Object.keys(pivot);

  document.getElementById('chart-section').innerHTML   = state.showChart   ? renderChartHTML(pivot)  : '';
  document.getElementById('compare-section').innerHTML = state.showCompare ? renderCompareHTML()      : '';

  if (!codigos.length) {
    document.getElementById('pivot-container').innerHTML =
      `<div class="empty-state"><div class="empty-icon">&#128202;</div>
       <div class="empty-title">Sin datos</div>
       <div class="empty-sub">${state.planLoaded?'No hay WOs.':'Importá el Plan.'}</div></div>`;
    return;
  }

  codigos.sort((a,b) => {
    if (state.filtroSemana !== 'Todos') {
      const aH = pivot[a][state.filtroSemana]?1:0, bH = pivot[b][state.filtroSemana]?1:0;
      if (aH!==bH) return bH-aH;
    }
    return a.localeCompare(b);
  });

  const alerts    = computeAlerts(pivot);
  const dupCods   = new Set(alerts.dups.map(d=>d.cod));
  const surpN2Set = new Set(alerts.surplus.filter(s=>s.moveTo==='N+2').map(s=>s.cod));
  const surpN1Set = new Set(alerts.surplus.filter(s=>s.moveTo==='N+1').map(s=>s.cod));

  const rows = codigos.map(cod => {
    const isExp  = !!state.expanded[cod];
    const isDup  = dupCods.has(cod);
    const isTop  = state.filtroSemana!=='Todos' && pivot[cod][state.filtroSemana];
    const hasN2S = surpN2Set.has(cod);
    const hasN1S = surpN1Set.has(cod);

    const cells = semanas.map(s => {
      const cell = pivot[cod][s];
      if (!cell) return `<td class="pcell-empty pcell${s==='N+3'?' pcell-n3':''}">—</td>`;
      const p = cell.pct;
      let cls = 'pcell'+(s==='N+3'?' pcell-n3':'');
      let pTxt = '—';
      if (p!==null && (state.entregaLoaded||cell.fromBodega)) {
        pTxt = p+'%';
        cls += p>state.config.maxPct?' cell-warn':p<state.config.minPct?' cell-surplus':' cell-ok';
      } else {
        cls += ' cell-nodata';
      }
      const cajas = state.config.factorCajas>0 ? Math.round(cell.planVol/state.config.factorCajas) : 0;
      const cajasStr = cajas>0 ? ` / ${fmtNum(cajas)}cj` : '';
      const isEdited = state.editedEntregas[cod+'__'+s] !== undefined;
      const bodBadge = cell.fromBodega
        ? ' <span style="font-size:9px;background:#E6F1FB;color:#185FA5;padding:1px 4px;border-radius:3px;font-weight:600">BOD</span>'
        : '';
      let semaHtml = '';
      if (state.insumoLoaded) {
        const cs = cell.wos.map(w=>getSemaforoColor(w.wo,s));
        const cnt={green:0,yellow:0,red:0,none:0}; cs.forEach(c=>cnt[c]++);
        const agg = cnt.red>0?'red':cnt.yellow>0?'yellow':cnt.none===cs.length?'none':'green';
        semaHtml = `<div class="cell-sema"><span class="sema-dot sema-${agg}" title="${cnt.green}✓ ${cnt.yellow}~ ${cnt.red}✗"></span></div>`;
      }
      return `<td class="${cls}" ondblclick="editEntrega('${cod}','${s}',${cell.entregaVol},event)">
        <div class="cell-pct">${pTxt}</div>
        <div class="cell-vols">WO:${fmtLts(cell.planVol)}${cajasStr}${(state.entregaLoaded||cell.fromBodega)?' / E:'+fmtLts(cell.entregaVol)+(bodBadge||(isEdited?' ✏':'')):''}</div>
        ${semaHtml}
      </td>`;
    }).join('');

    const allWOs  = semanas.flatMap(s => (pivot[cod][s]?pivot[cod][s].wos:[]));

    const subRows = allWOs.map(w => {
      const effSem   = w.semLabel;
      const isMoved  = state.movedToN1[w.id]||state.movedToN2[w.id];
      const canMoveN2 = effSem==='N+3' && hasN2S && !state.movedToN2[w.id];
      const canMoveN1 = effSem==='N+2' && hasN1S && !state.movedToN1[w.id];
      const isN3      = effSem==='N+3';
      const note      = state.notes[w.id];
      const hasCodEdit = state.editedCodEq[w.id];
      let trCls = '';
      let arrivalHtml = '';
      if (isN3 && state.insumoLoaded) {
        const sColor = getSemaforoColor(w.wo,'N+3');
        trCls = sColor==='green'?'sub-ok':(sColor==='red'||sColor==='yellow')?'sub-nok':'';
        const info = getArrivalInfo(w.wo,'N+3');
        if (info) arrivalHtml = `<span class="arrival-date${info.ok?'':' nok'}">${info.text}</span>`;
      }
      const cajas = state.config.factorCajas>0 ? ` / ${Math.round(w.lts/state.config.factorCajas)} cj` : '';
      return `<tr class="${trCls}" onclick="event.stopPropagation()">
        <td style="padding-left:44px">
          ${isN3&&state.insumoLoaded?`<span class="sema-dot sema-${getSemaforoColor(w.wo,'N+3')}" title="${getSemaforoText(w.wo,'N+3')}"></span> `:''}
          <span class="wo-num">${w.wo}</span>
          ${hasCodEdit?`<span style="font-size:10px;color:#185FA5;margin-left:5px">✏ ${getCodEq(w)}</span>`:''}
          ${arrivalHtml}
        </td>
        <td><span class="sem-badge${isN3?' sem-n3':''}">${effSem}</span>
          ${isMoved?`<span style="font-size:9px;color:#aaa"> (orig. ${getOrigSemana(w)||'?'})</span>`:''}
        </td>
        <td>${fmtLts(w.lts)} Lts${cajas}</td>
        <td>${w.estado||'—'}</td>
        <td>${w.fecha||'—'}</td>
        <td style="white-space:nowrap">
          ${canMoveN2?`<button class="btn-mover"    onclick="moveToN2(${w.id})">&#8593; N+2</button> `:''}
          ${canMoveN1?`<button class="btn-mover-n1" onclick="moveToN1(${w.id})">&#8593; N+1</button> `:''}
          ${isMoved  ?`<button class="btn-revert"   onclick="revertirA(${w.id})">&#8595; Revertir</button> `:''}
          <button class="btn-note"    onclick="editNote(${w.id},'${w.wo}',event)" title="Nota">&#128203;</button>
          <button class="btn-edit-sm" onclick="editCodEq(${w.id},'${w.wo}','${getCodEq(w)}',event)" title="Editar CodEq">&#9998;</button>
        </td>
      </tr>
      ${note?`<tr class="note-row" onclick="event.stopPropagation()">
        <td colspan="6"><span style="font-size:10px;color:#aaa">Nota:</span> <span class="note-text">${note}</span></td>
      </tr>`:''}`;
    }).join('');

    const subSection = isExp
      ? `<tr class="sub-row"><td colspan="5"><table class="sub-table">
           <thead><tr><th>WO</th><th>Semana</th><th>Lts</th><th>Estado</th><th>Fecha</th><th></th></tr></thead>
           <tbody>${subRows}</tbody>
         </table></td></tr>`
      : '';

    return `
      <tr class="cod-row" onclick="toggleExpand('${codSafe(cod)}')">
        <td><div class="cod-cell">
          <span class="expand-icon${isExp?' open':''}">&#9658;</span>
          <span class="cod-name">${cod}</span>
          ${isDup ?'<span class="badge b-dup">3+ sem.</span>'    :''}
          ${hasN2S?'<span class="badge b-surplus">Sobrante N+2</span>':''}
          ${hasN1S?'<span class="badge b-surplus">Sobrante N+1</span>':''}
          ${isTop ?'<span class="badge b-top">&#9650;</span>'    :''}
        </div></td>
        ${cells}
      </tr>${subSection}`;
  }).join('');

  document.getElementById('pivot-container').innerHTML = `
    ${renderCapSummary(pivot)}
    <div class="pivot-wrap">
      <table class="pivot-table">
        <thead><tr>
          <th class="th-left">Codigo vino</th>
          ${semanas.map(s=>`<th class="th-sem${s==='N+3'?' th-n3':''}">
            ${s==='N+3'?'N+3 <span class="n3-tag-sm">+vis.</span>':'Sem. '+s}
          </th>`).join('')}
        </tr></thead>
        <tbody>${rows}</tbody>
      </table>
      <div class="pivot-legend">
        <span class="leg-item"><span class="leg-dot ok-dot"></span>En rango</span>
        <span class="leg-item"><span class="leg-dot warn-dot"></span>Deficit</span>
        <span class="leg-item"><span class="leg-dot surplus-dot"></span>Sobrante</span>
        ${state.insumoLoaded?`
          <span class="leg-item"><span class="sema-dot sema-green"></span>Insumos OK</span>
          <span class="leg-item"><span class="sema-dot sema-yellow"></span>Insumos parcial</span>
          <span class="leg-item"><span class="sema-dot sema-red"></span>Insumos falta</span>`:''}
        <span class="leg-item" style="color:#bbb;font-size:11px">
          % = WO/Entrega &nbsp;·&nbsp; +${fmtLts(state.config.merma)} merma &nbsp;·&nbsp; BOD = desde Bodegas &nbsp;·&nbsp; Doble-clic E: editar entrega
        </span>
      </div>
    </div>`;
}

function renderCapSummary(pivot) {
  const semanas  = ['N','N+1','N+2','N+3'];
  const caps     = state.config.capacidades;
  const totalCap = semanas.reduce((a,s)=>a+(+caps[s]||0),0);
  if (!totalCap) return '';
  const cards = semanas.map(s => {
    const cap   = +caps[s]||0;
    const planV = Object.values(pivot).reduce((a,cd)=>a+(cd[s]?cd[s].planVol:0),0);
    const pct   = cap>0?Math.round(planV/cap*100):0;
    const cls   = pct>=95?'danger':pct>=70?'ok':'warn';
    const fCls  = pct>=95?'fill-danger':pct>=70?'fill-ok':'fill-warn';
    const cajas = state.config.factorCajas>0 ? `<div style="font-size:10px;color:#aaa">${fmtNum(Math.round(planV/state.config.factorCajas))} cj</div>` : '';
    return `<div class="cap-card">
      <div class="cap-sem">Sem. ${s}</div>
      <div class="cap-bar-wrap"><div class="cap-bar-fill ${fCls}" style="width:${Math.min(pct,100)}%"></div></div>
      <div class="cap-nums">${fmtLts(planV)} / ${fmtLts(cap)} Lts</div>
      ${cajas}
      <div class="cap-pct ${cls}">${cap>0?pct+'%':'—'}</div>
    </div>`;
  }).join('');
  const tPlanV = semanas.reduce((a,s)=>a+Object.values(pivot).reduce((b,cd)=>b+(cd[s]?cd[s].planVol:0),0),0);
  const tPct   = totalCap>0?Math.round(tPlanV/totalCap*100):0;
  const tCls   = tPct>=95?'danger':tPct>=70?'ok':'warn';
  const tFill  = tPct>=95?'fill-danger':tPct>=70?'fill-ok':'fill-warn';
  const tCajas = state.config.factorCajas>0 ? `<div style="font-size:10px;color:#aaa">${fmtNum(Math.round(tPlanV/state.config.factorCajas))} cj</div>` : '';
  return `<div class="cap-summary">
    <div class="cap-summary-title">Capacidad de fraccionamiento por semana</div>
    <div class="cap-grid">
      ${cards}
      <div class="cap-card total">
        <div class="cap-sem">TOTAL</div>
        <div class="cap-bar-wrap"><div class="cap-bar-fill ${tFill}" style="width:${Math.min(tPct,100)}%"></div></div>
        <div class="cap-nums">${fmtLts(tPlanV)} / ${fmtLts(totalCap)} Lts</div>
        ${tCajas}
        <div class="cap-pct ${tCls}">${tPct}%</div>
      </div>
    </div>
  </div>`;
}

// ── HISTORIAL ─────────────────────────────────────────────
function renderHistorial() {
  const el = document.getElementById('historial-container');
  if (!el) return;
  if (!state.historial.length) {
    el.innerHTML = `<div class="empty-state"><div class="empty-icon">&#128203;</div>
      <div class="empty-title">Sin movimientos</div>
      <div class="empty-sub">Los movimientos se registran aqui automaticamente.</div></div>`;
    return;
  }
  const rows = state.historial.map(h => {
    const d   = new Date(h.ts);
    const ts  = d.toLocaleDateString('es-AR')+' '+d.toLocaleTimeString('es-AR',{hour:'2-digit',minute:'2-digit'});
    const icon= h.accion==='revertir'?'&#8617;':'&#8593;';
    const cls = h.accion==='revertir'?'hist-revert':'hist-move';
    return `<div class="hist-item ${cls}">
      <span class="hist-icon">${icon}</span>
      <div class="hist-content">
        <div class="hist-desc"><strong>${h.wo}</strong> (${h.codEq}) &nbsp;&rarr;&nbsp; ${h.de} &rarr; ${h.a}</div>
        <div class="hist-time">${ts} &middot; ${h.accion}</div>
      </div>
    </div>`;
  }).join('');
  el.innerHTML = `<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:12px">
    <span style="font-size:11px;font-weight:500;color:#888;text-transform:uppercase;letter-spacing:.05em">${state.historial.length} movimientos</span>
    <button class="btn-secondary" style="font-size:11px;padding:4px 10px" onclick="limpiarHistorial()">Limpiar</button>
  </div>${rows}`;
}
function limpiarHistorial() {
  if (!confirm('¿Limpiar el historial?')) return;
  state.historial = [];
  const hb = document.getElementById('hist-badge');
  if (hb) hb.textContent = 0;
  autoSave(); renderHistorial();
}

// ── ALERTS ────────────────────────────────────────────────
function computeAlerts(pivot) {
  const semanas = ['N','N+1','N+2','N+3'];
  const dups=[],deficit=[],surplus=[],insumoProb=[];
  Object.keys(pivot).forEach(cod => {
    const present = semanas.filter(s => pivot[cod][s]?.planVol > 0);
    if (present.length >= 3) dups.push({ cod, sems: present.join(', ') });
    semanas.forEach(s => {
      const cell = pivot[cod][s];
      if (cell?.pct > state.config.maxPct) deficit.push({ cod, semana:s, pct:cell.pct });
    });
    const n2 = pivot[cod]['N+2'];
    if (n2?.entregaVol>0 && n2.planVol<n2.entregaVol) {
      const wos = (pivot[cod]['N+3']||{wos:[]}).wos.filter(w=>!state.movedToN2[w.id]);
      if (wos.length) surplus.push({ cod, sobrante:n2.entregaVol-n2.planVol, wos, moveTo:'N+2', moveFn:'moveToN2' });
    }
    const n1 = pivot[cod]['N+1'];
    if (n1?.entregaVol>0 && n1.planVol<n1.entregaVol) {
      const wos = (pivot[cod]['N+2']||{wos:[]}).wos.filter(w=>!state.movedToN1[w.id]);
      if (wos.length) surplus.push({ cod, sobrante:n1.entregaVol-n1.planVol, wos, moveTo:'N+1', moveFn:'moveToN1' });
    }
    if (state.insumoLoaded) {
      semanas.forEach(s => {
        pivot[cod][s]?.wos.forEach(w => {
          const col = getSemaforoColor(w.wo, s);
          if (col==='red'||col==='yellow')
            insumoProb.push({wo:w.wo, cod, semana:s, color:col, text:getSemaforoText(w.wo,s)});
        });
      });
    }
  });
  return { dups, deficit, surplus, insumoProb };
}

function renderAlertas(pivot) {
  const a = computeAlerts(pivot);
  let html = '';
  if (a.surplus.length) {
    html += `<div class="alert-section">
      <div class="alert-section-title">&#128200; Sobrante — adelantar ordenes (${a.surplus.length})</div>
      ${a.surplus.map(s=>`<div class="alert-card surplus">
        <div class="alert-icon">&#128199;</div>
        <div style="flex:1">
          <div class="alert-title">${s.cod} — Sobrante ${fmtLts(s.sobrante)} Lts en ${s.moveTo==='N+1'?'N+1':'N+2'}</div>
          <div class="alert-sub">WOs candidatas a adelantar a ${s.moveTo}:</div>
          <div class="alert-wo-list">
            ${s.wos.map(w=>{
              const sColor = state.insumoLoaded?getSemaforoColor(w.wo,w.semLabel):'none';
              const info   = state.insumoLoaded?getArrivalInfo(w.wo,w.semLabel):null;
              return `<div class="alert-wo-row">
                ${state.insumoLoaded?`<span class="sema-dot sema-${sColor}"></span>`:''}
                <span><strong>${w.wo}</strong> &middot; ${fmtLts(w.lts)} Lts
                  ${info?`<span class="arrival-date${info.ok?'':' nok'}">${info.text}</span>`:''}
                </span>
                <button class="${s.moveTo==='N+1'?'btn-mover-n1':'btn-mover'}" onclick="${s.moveFn}(${w.id})">&#8593; Mover a ${s.moveTo}</button>
              </div>`;
            }).join('')}
          </div>
        </div>
      </div>`).join('')}
    </div>`;
  }
  if (a.insumoProb.length) {
    const rojos=a.insumoProb.filter(x=>x.color==='red');
    const ams  =a.insumoProb.filter(x=>x.color==='yellow');
    html += `<div class="alert-section">
      <div class="alert-section-title">&#128994; Semaforo insumos (${a.insumoProb.length} WOs)</div>
      ${rojos.length?`<div class="alert-card danger"><div class="alert-icon">&#128308;</div>
        <div><div class="alert-title">Sin insumos completos (${rojos.length})</div>
        ${rojos.map(x=>`<div style="font-size:12px;margin-top:4px">
          <span class="sema-dot sema-red" style="display:inline-block;vertical-align:middle;margin-right:5px"></span>
          <strong>${x.wo}</strong> &middot; ${x.cod} &middot; Sem.${x.semana} &middot; <span style="color:#888">${x.text}</span>
        </div>`).join('')}</div></div>`:''}
      ${ams.length?`<div class="alert-card warn"><div class="alert-icon">&#128993;</div>
        <div><div class="alert-title">Insumos parciales (${ams.length})</div>
        ${ams.map(x=>`<div style="font-size:12px;margin-top:4px">
          <span class="sema-dot sema-yellow" style="display:inline-block;vertical-align:middle;margin-right:5px"></span>
          <strong>${x.wo}</strong> &middot; ${x.cod} &middot; Sem.${x.semana} &middot; <span style="color:#888">${x.text}</span>
        </div>`).join('')}</div></div>`:''}
    </div>`;
  }
  if (a.dups.length) {
    html += `<div class="alert-section">
      <div class="alert-section-title">&#9888; Mismo codigo 3+ semanas (${a.dups.length})</div>
      ${a.dups.map(d=>`<div class="alert-card warn"><div class="alert-icon">&#128260;</div>
        <div><div class="alert-title">${d.cod}</div><div class="alert-sub">Semanas: ${d.sems}</div></div>
      </div>`).join('')}
    </div>`;
  }
  if (a.deficit.length) {
    html += `<div class="alert-section">
      <div class="alert-section-title">&#128314; Deficit WOs &gt; Entrega (${a.deficit.length})</div>
      ${a.deficit.map(d=>`<div class="alert-card danger"><div class="alert-icon">&#128202;</div>
        <div><div class="alert-title">${d.cod} &middot; Sem. ${d.semana}</div>
        <div class="alert-sub">WOs = ${d.pct}% de la entrega (max: ${state.config.maxPct}%)</div></div>
      </div>`).join('')}
    </div>`;
  }
  if (!html) html = `<div class="alert-card ok"><div class="alert-icon">&#9989;</div>
    <div><div class="alert-title">Sin alertas activas</div><div class="alert-sub">El plan esta balanceado.</div></div></div>`;
  document.getElementById('alertas-container').innerHTML = html;
}

// ── CONFIG ────────────────────────────────────────────────
function renderConfig() {
  const caps  = state.config.capacidades;
  const total = Object.values(caps).reduce((a,v)=>a+(+v||0),0);
  document.getElementById('config-container').innerHTML = `
    <div class="config-section">
      <div class="config-title">Configuracion de semanas</div>
      <div class="config-grid">
        <div class="config-card">
          <div class="config-card-title">Inicio de semana N</div>
          <div class="field-row">
            <label>Fecha de inicio (lunes de semana N)</label>
            <input type="date" id="cfg-fecha" value="${state.config.semanaInicio}">
            <span class="hint">N+1=+7 dias · N+2=+14 · N+3=+21</span>
          </div>
        </div>
        <div class="config-card">
          <div class="config-card-title">Merma y conversion</div>
          <div class="field-row">
            <label>Merma (Lts por CodEq por semana)</label>
            <input type="number" id="cfg-merma" value="${state.config.merma}">
          </div>
          <div class="field-row">
            <label>Factor cajas (Lts por caja)</label>
            <input type="number" id="cfg-cajas" value="${state.config.factorCajas}">
            <span class="hint">Ej: 9 Lts = 1 caja (12 botellas 750ml)</span>
          </div>
        </div>
      </div>
    </div>
    <div class="config-section">
      <div class="config-title">Capacidad de fraccionamiento por semana</div>
      <div class="config-card">
        <div class="cap-inputs">
          ${['N','N+1','N+2','N+3'].map(s=>`
            <div class="cap-input-item">
              <label>Semana ${s} (Lts)</label>
              <input type="number" id="cfg-cap-${s.replace('+','p')}" value="${caps[s]||0}" placeholder="ej: 150000">
            </div>`).join('')}
        </div>
        <div class="total-cap">
          <span>Total 4 semanas</span>
          <span id="cfg-cap-total-display"><strong>${fmtLts(total)} Lts</strong></span>
        </div>
      </div>
    </div>
    <div class="config-section">
      <div class="config-title">Rangos de alerta (%)</div>
      <div class="config-grid">
        <div class="config-card">
          <div class="config-card-title">% Minimo (sobrante)</div>
          <div class="field-row">
            <label>Celda azul si WO/Entrega es menor a</label>
            <input type="number" id="cfg-min" value="${state.config.minPct}">
          </div>
        </div>
        <div class="config-card">
          <div class="config-card-title">% Maximo (deficit)</div>
          <div class="field-row">
            <label>Celda roja si WO/Entrega supera</label>
            <input type="number" id="cfg-max" value="${state.config.maxPct}">
          </div>
        </div>
      </div>
    </div>
    <div style="display:flex;align-items:center">
      <button class="btn-primary" onclick="saveConfig()">Guardar configuracion</button>
      <span id="cfg-ok" class="save-feedback"></span>
    </div>`;
  ['N','N+1','N+2','N+3'].forEach(s => {
    document.getElementById('cfg-cap-'+s.replace('+','p'))?.addEventListener('input', updateCapTotal);
  });
}
function updateCapTotal() {
  const t = ['N','N+1','N+2','N+3'].reduce((a,s) => {
    const el = document.getElementById('cfg-cap-'+s.replace('+','p'));
    return a + (+el?.value||0);
  }, 0);
  const el = document.getElementById('cfg-cap-total-display');
  if (el) el.innerHTML = `<strong>${fmtLts(t)} Lts</strong>`;
}
function saveConfig() {
  state.config.semanaInicio = document.getElementById('cfg-fecha')?.value   || state.config.semanaInicio;
  state.config.merma        = parseFloat(document.getElementById('cfg-merma')?.value)  || 1000;
  state.config.factorCajas  = parseFloat(document.getElementById('cfg-cajas')?.value)  || 9;
  state.config.minPct       = parseFloat(document.getElementById('cfg-min')?.value)    || 80;
  state.config.maxPct       = parseFloat(document.getElementById('cfg-max')?.value)    || 100;
  ['N','N+1','N+2','N+3'].forEach(s => {
    const el = document.getElementById('cfg-cap-'+s.replace('+','p'));
    if (el) state.config.capacidades[s] = parseFloat(el.value)||0;
  });
  const fb = document.getElementById('cfg-ok');
  if (fb) fb.textContent = '✓ Guardado';
  setTimeout(() => { const e=document.getElementById('cfg-ok'); if(e) e.textContent=''; }, 2500);
  autoSave();
  if (state.planLoaded||state.entregaLoaded) render();
}

// ── UTILS ─────────────────────────────────────────────────
function fmtLts(v) {
  if (!v) return '0';
  if (v>=1000000) return (v/1000000).toFixed(1)+'M';
  if (v>=1000)    return (Math.round(v/100)/10).toFixed(1)+'k';
  return Math.round(v).toString();
}
function fmtNum(n) {
  if (n>=1000) return (n/1000).toFixed(1)+'k';
  return Math.round(n).toString();
}
function fmtFecha(val) {
  if (!val) return '';
  if (val instanceof Date && !isNaN(val)) return val.toLocaleDateString('es-AR');
  const s = String(val).trim();
  const n = parseFloat(s);
  if (!isNaN(n) && n>1000 && n<100000) return new Date((n-25569)*86400000).toLocaleDateString('es-AR');
  return s.split(' ')[0];
}
function fmtDateShort(d) {
  if (!d||!(d instanceof Date)||isNaN(d)) return '—';
  return d.toLocaleDateString('es-AR');
}
function codSafe(cod) { return cod.replace(/[^a-zA-Z0-9]/g,'_'); }
function toggleExpand(safeKey) {
  const cod = Object.keys(buildPivot()).find(c => codSafe(c)===safeKey);
  if (!cod) return;
  state.expanded[cod] = !state.expanded[cod];
  render();
}
