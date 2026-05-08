// ============================================================
// BALANCEO DE CARGAS WO v4
// ============================================================

const state = {
  planData:      [],
  entregaData:   [],
  insumoData:    {},   // { "WO_NUM": { etiqueta: Date|null, caja: Date|null } }
  semanaMap:     {},
  expanded:      {},
  movedToN2:     {},
  filtroSemana:  'Todos',
  config: {
    semanaInicio:   '',        // "YYYY-MM-DD" lunes de semana N
    capacidades:    { N:0, 'N+1':0, 'N+2':0, 'N+3':0 },
    minPct: 80,
    maxPct: 100,
    merma:  1000               // Lts por CodEq por semana
  },
  importMode:    'plan',
  rawHeaders:    [],
  rawData:       [],
  colMap:        {},
  nextId:        1,
  planLoaded:    false,
  entregaLoaded: false,
  insumoLoaded:  false
};

// ── INIT ──────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  renderConfig();
  renderSemanaFilter();
  document.getElementById('filePlan')
    .addEventListener('change', e => handleFileUpload(e, 'plan'));
  document.getElementById('fileEntrega')
    .addEventListener('change', e => handleFileUpload(e, 'entregas'));
  document.getElementById('fileInsumo')
    .addEventListener('change', e => handleFileUpload(e, 'insumo'));
  document.querySelectorAll('.tab').forEach(tab => {
    tab.addEventListener('click', () => {
      document.querySelectorAll('.tab').forEach(t => t.classList.remove('on'));
      document.querySelectorAll('.panel').forEach(p => p.classList.remove('on'));
      tab.classList.add('on');
      document.getElementById('panel-' + tab.dataset.tab).classList.add('on');
    });
  });
});

// ── FILE IMPORT ───────────────────────────────────────────
function handleFileUpload(e, mode) {
  const file = e.target.files[0];
  if (!file) return;
  state.importMode = mode;
  const reader = new FileReader();
  reader.onload = function(evt) {
    const wb    = XLSX.read(evt.target.result, { type:'binary', cellDates:true });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const rows  = XLSX.utils.sheet_to_json(sheet, { header:1, defval:'' });
    if (rows.length < 2) { alert('El archivo parece estar vacío.'); return; }
    state.rawHeaders = rows[0].map(h => String(h).trim());
    state.rawData    = rows.slice(1).filter(r => r.some(c => c !== ''));
    showColumnMapper(state.rawHeaders, autoDetect(state.rawHeaders, mode), mode);
  };
  reader.readAsBinaryString(file);
  e.target.value = '';
}

function autoDetect(headers, mode) {
  const map = { codEq:'', semana:'', lts:'', wo:'', estado:'', fecha:'', etiqueta:'', caja:'' };
  headers.forEach(h => {
    const hl = h.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'');
    if (!map.codEq    && /codeq|cod.eq|equivalente/.test(hl))         map.codEq    = h;
    if (!map.semana   && /\bsemana\b/.test(hl))                        map.semana   = h;
    if (!map.lts      && /^lts$/.test(hl))                             map.lts      = h;
    if (!map.wo       && /^orden$/.test(hl))                           map.wo       = h;
    if (!map.estado   && /^estado$/.test(hl))                          map.estado   = h;
    if (!map.fecha    && /feinicio|fechainicio/.test(hl))              map.fecha    = h;
    if (!map.etiqueta && /etiqueta|label/.test(hl))                    map.etiqueta = h;
    if (!map.caja     && /\bcaja\b|box/.test(hl))                      map.caja     = h;
  });
  // Fallbacks
  if (!map.codEq)    map.codEq    = headers.find(h => /codeq/i.test(h))         || '';
  if (!map.semana)   map.semana   = headers.find(h => /semana/i.test(h))         || '';
  if (!map.lts)      map.lts      = headers.find(h => /^lts$/i.test(h))          || '';
  if (!map.wo)       map.wo       = headers.find(h => /orden/i.test(h))          || '';
  if (!map.etiqueta) map.etiqueta = headers.find(h => /etiqueta/i.test(h))       || '';
  if (!map.caja)     map.caja     = headers.find(h => /caja/i.test(h))           || '';
  return map;
}

function showColumnMapper(headers, detected, mode) {
  const fieldsByMode = {
    plan: [
      { key:'codEq',  label:'Código equivalente (CodEq)', req:true  },
      { key:'semana', label:'Semana',                      req:true  },
      { key:'lts',    label:'Volumen (Lts)',                req:true  },
      { key:'wo',     label:'Número de WO (Orden)',        req:true  },
      { key:'estado', label:'Estado',                      req:false },
      { key:'fecha',  label:'Fecha de inicio (FeInicio)',  req:false }
    ],
    entregas: [
      { key:'codEq',  label:'Código equivalente (CodEq)', req:true },
      { key:'semana', label:'Semana',                      req:true },
      { key:'lts',    label:'Volumen (Lts)',                req:true }
    ],
    insumo: [
      { key:'wo',       label:'Número de WO (Orden)',       req:true  },
      { key:'etiqueta', label:'Fecha llegada etiqueta',     req:true  },
      { key:'caja',     label:'Fecha llegada caja',         req:true  }
    ]
  };
  const titles = { plan:'Columnas del Plan', entregas:'Columnas de Entregas', insumo:'Columnas de Insumos' };
  document.getElementById('modal-title').textContent = titles[mode];
  document.getElementById('modal-desc').textContent  = 'Seleccioná las columnas del archivo:';
  document.getElementById('col-map-fields').innerHTML =
    (fieldsByMode[mode] || []).map(f => `
      <div class="col-map-row">
        <label>${f.label}${f.req ? '' : ' <span style="color:#bbb;font-size:11px">(opc.)</span>'}</label>
        <select id="map-${f.key}">
          ${!f.req ? '<option value="">— no usar —</option>' : '<option value="">— seleccionar —</option>'}
          ${headers.map(h =>
            `<option value="${h}"${h === detected[f.key] ? ' selected' : ''}>${h}</option>`
          ).join('')}
        </select>
      </div>`).join('');
  document.getElementById('modal-overlay').classList.remove('hidden');
}

function confirmImport() {
  const mode = state.importMode;
  const reqByMode = { plan:['codEq','semana','lts','wo'], entregas:['codEq','semana','lts'], insumo:['wo','etiqueta','caja'] };
  const newMap = {};
  ['codEq','semana','lts','wo','estado','fecha','etiqueta','caja'].forEach(f => {
    const el = document.getElementById('map-' + f);
    if (el) newMap[f] = el.value;
  });
  if ((reqByMode[mode] || []).some(f => !newMap[f])) {
    alert('Asigná todas las columnas obligatorias.');
    return;
  }
  cerrarModal();

  if (mode === 'insumo') {
    state.insumoData   = processInsumoData(state.rawData, newMap);
    state.insumoLoaded = true;
    updateStatus('insumo', Object.keys(state.insumoData).length + ' WOs cargadas');
  } else {
    const processed = processData(state.rawData, newMap);
    if (mode === 'plan') {
      state.planData   = processed;
      state.planLoaded = true;
      state.movedToN2  = {};
      state.expanded   = {};
      updateStatus('plan', processed.length + ' filas cargadas');
    } else {
      state.entregaData   = processed;
      state.entregaLoaded = true;
      updateStatus('entregas', processed.length + ' filas cargadas');
    }
    buildSemanaMap();
  }
  render();
}

function cerrarModal() {
  document.getElementById('modal-overlay').classList.add('hidden');
}

function processData(rows, m) {
  const H = state.rawHeaders;
  const ix = k => m[k] ? H.indexOf(m[k]) : -1;
  const get = (r, k) => ix(k) >= 0 ? r[ix(k)] : '';
  return rows.map((r, i) => ({
    id:        state.nextId++,
    wo:        String(get(r,'wo')  || 'WO-' + (i+1)).trim(),
    codEq:     String(get(r,'codEq')  || '').trim(),
    semanaRaw: String(get(r,'semana') || '').trim(),
    lts:       parseFloat(get(r,'lts')) || 0,
    estado:    String(get(r,'estado') || '').trim(),
    fecha:     fmtFecha(get(r,'fecha'))
  })).filter(r => r.codEq && r.lts > 0);
}

function processInsumoData(rows, m) {
  const H = state.rawHeaders;
  const ix = k => m[k] ? H.indexOf(m[k]) : -1;
  const result = {};
  rows.forEach(r => {
    const wo  = String(ix('wo') >= 0 ? r[ix('wo')] : '').trim();
    if (!wo) return;
    result[wo] = {
      etiqueta: parseExcelDate(ix('etiqueta') >= 0 ? r[ix('etiqueta')] : ''),
      caja:     parseExcelDate(ix('caja')     >= 0 ? r[ix('caja')]     : '')
    };
  });
  return result;
}

// ── DATE PARSING ──────────────────────────────────────────
function parseExcelDate(val) {
  if (!val && val !== 0) return null;
  if (val instanceof Date) return val;
  const s = String(val).trim();
  if (!s) return null;
  // Excel serial number
  const n = parseFloat(s);
  if (!isNaN(n) && n > 1000) return new Date((n - 25569) * 86400000);
  // DD/MM/YYYY or D/M/YYYY
  const dmy = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (dmy) return new Date(+dmy[3], +dmy[2]-1, +dmy[1]);
  // YYYY-MM-DD
  const ymd = s.match(/^(\d{4})[\/\-](\d{2})[\/\-](\d{2})/);
  if (ymd) return new Date(+ymd[1], +ymd[2]-1, +ymd[3]);
  return null;
}

function getSemanaStartDate(semLabel) {
  if (!state.config.semanaInicio) return null;
  const base   = new Date(state.config.semanaInicio + 'T00:00:00');
  const offset = { 'N':0, 'N+1':7, 'N+2':14, 'N+3':21 }[semLabel];
  if (offset === undefined) return null;
  const d = new Date(base);
  d.setDate(d.getDate() + offset);
  return d;
}

function getSemaforoColor(woStr, semLabel) {
  const ins = state.insumoData[woStr];
  if (!ins) return 'none';                     // sin dato de insumo
  const semStart = getSemanaStartDate(semLabel);
  const isOk = (date) => !date || (semStart ? date <= semStart : true);
  const etiqOk = isOk(ins.etiqueta);
  const cajaOk = isOk(ins.caja);
  if (etiqOk && cajaOk)  return 'green';
  if (etiqOk || cajaOk)  return 'yellow';
  return 'red';
}

function getSemaforoText(woStr, semLabel) {
  const ins = state.insumoData[woStr];
  if (!ins) return '';
  const semStart = getSemanaStartDate(semLabel);
  const fmt = d => d ? fmtDateShort(d) : 'sin fecha';
  const isOk = d => !d || (semStart ? d <= semStart : true);
  return `Etiqueta: ${fmt(ins.etiqueta)}${isOk(ins.etiqueta)?'✓':'✗'} · Caja: ${fmt(ins.caja)}${isOk(ins.caja)?'✓':'✗'}`;
}

// ── WEEK NORMALIZATION ────────────────────────────────────
function buildSemanaMap() {
  const all = new Set([
    ...state.planData.map(r => r.semanaRaw),
    ...state.entregaData.map(r => r.semanaRaw)
  ]);
  const sorted = [...all].filter(Boolean).sort((a, b) => {
    const na = parseFloat(a), nb = parseFloat(b);
    return !isNaN(na) && !isNaN(nb) ? na - nb : a.localeCompare(b);
  });
  const labels = ['N','N+1','N+2','N+3'];
  state.semanaMap = {};
  sorted.forEach((w, i) => { state.semanaMap[w] = labels[i] || 'N+3'; });
  const txt = sorted.map((w, i) => `${labels[i]||'N+3'}=sem.${w}`).join(' · ');
  document.getElementById('status-semanas').textContent = txt;
}

function getSemana(row) {
  return state.movedToN2[row.id] ? 'N+2' : (state.semanaMap[row.semanaRaw] || null);
}

function updateStatus(type, text) {
  const map = { plan:'status-plan', entregas:'status-entrega', insumo:'status-insumo' };
  const labels = { plan:'Plan: ', entregas:'Entregas: ', insumo:'Insumos: ' };
  const el = document.getElementById(map[type]);
  if (!el) return;
  el.querySelector('.status-dot').className = 'status-dot dot-on';
  el.querySelector('span:last-child').textContent = labels[type] + text;
}

function moveToN2(woId) {
  state.movedToN2[woId] = true;
  render();
}

function toggleExpand(safeKey) {
  const cod = Object.keys(buildPivot()).find(c => codSafe(c) === safeKey);
  if (!cod) return;
  state.expanded[cod] = !state.expanded[cod];
  render();
}

// ── DEMO DATA ─────────────────────────────────────────────
function cargarDemoData() {
  state.planData = [
    { id:1,  wo:'1129121', codEq:'C1079',       semanaRaw:'22', lts:7141.5, estado:'40', fecha:'25/5/2026' },
    { id:2,  wo:'1134019', codEq:'C1079',       semanaRaw:'22', lts:7012.1, estado:'40', fecha:'25/5/2026' },
    { id:3,  wo:'1134020', codEq:'A1039/A1497', semanaRaw:'22', lts:5200.0, estado:'40', fecha:'25/5/2026' },
    { id:4,  wo:'1134021', codEq:'A1039/A1497', semanaRaw:'23', lts:5100.0, estado:'40', fecha:'2/6/2026'  },
    { id:5,  wo:'1134022', codEq:'C1079',       semanaRaw:'23', lts:8500.0, estado:'40', fecha:'2/6/2026'  },
    { id:6,  wo:'1134023', codEq:'B2201',       semanaRaw:'23', lts:3800.0, estado:'40', fecha:'2/6/2026'  },
    { id:7,  wo:'1134024', codEq:'A1039/A1497', semanaRaw:'24', lts:5050.0, estado:'20', fecha:'9/6/2026'  },
    { id:8,  wo:'1134025', codEq:'B2201',       semanaRaw:'24', lts:3600.0, estado:'20', fecha:'9/6/2026'  },
    { id:9,  wo:'1134026', codEq:'C1079',       semanaRaw:'25', lts:7200.0, estado:'20', fecha:'16/6/2026' },
    { id:10, wo:'1134027', codEq:'D3310',       semanaRaw:'25', lts:4100.0, estado:'20', fecha:'16/6/2026' },
    { id:11, wo:'1134028', codEq:'A1039/A1497', semanaRaw:'25', lts:4900.0, estado:'20', fecha:'16/6/2026' },
  ];
  state.entregaData = [
    { id:20, codEq:'C1079',       semanaRaw:'22', lts:16000.0 },
    { id:21, codEq:'A1039/A1497', semanaRaw:'22', lts:4800.0  },
    { id:22, codEq:'A1039/A1497', semanaRaw:'23', lts:5300.0  },
    { id:23, codEq:'C1079',       semanaRaw:'23', lts:7200.0  },
    { id:24, codEq:'B2201',       semanaRaw:'23', lts:3600.0  },
    { id:25, codEq:'A1039/A1497', semanaRaw:'24', lts:9500.0  },
    { id:26, codEq:'B2201',       semanaRaw:'24', lts:3400.0  },
    { id:27, codEq:'C1079',       semanaRaw:'25', lts:7500.0  },
    { id:28, codEq:'D3310',       semanaRaw:'25', lts:3200.0  },
  ];
  state.insumoData = {
    '1134021': { etiqueta: new Date(2026,5,1),  caja: new Date(2026,5,1)  },  // verde N+1
    '1134022': { etiqueta: new Date(2026,5,8),  caja: new Date(2026,5,1)  },  // amarillo N+1
    '1134026': { etiqueta: new Date(2026,5,20), caja: new Date(2026,5,20) },  // rojo N+3
    '1134028': { etiqueta: null,                caja: null                },  // verde N+3
  };
  state.planLoaded    = true;
  state.entregaLoaded = true;
  state.insumoLoaded  = true;
  state.movedToN2 = {};
  state.expanded  = {};
  state.nextId    = 60;
  if (!state.config.semanaInicio) state.config.semanaInicio = '2026-05-25';
  if (!state.config.capacidades.N) {
    state.config.capacidades = { N:160000, 'N+1':160000, 'N+2':160000, 'N+3':160000 };
  }
  updateStatus('plan',     '11 filas cargadas (demo)');
  updateStatus('entregas',  '9 filas cargadas (demo)');
  updateStatus('insumo',    '4 WOs cargadas (demo)');
  buildSemanaMap();
  renderConfig();
  render();
}

// ── PIVOT ─────────────────────────────────────────────────
function buildPivot() {
  const pivot = {};
  // Solo CodEq con WOs en plan
  state.planData.forEach(r => {
    const sem = getSemana(r);
    if (!sem) return;
    if (!pivot[r.codEq]) pivot[r.codEq] = {};
    if (!pivot[r.codEq][sem]) pivot[r.codEq][sem] = { planVol:0, entregaVol:0, wos:[] };
    pivot[r.codEq][sem].planVol += r.lts;
    pivot[r.codEq][sem].wos.push({ ...r, semLabel:sem });
  });
  // Merma: +config.merma por CodEq por semana que tenga WOs
  Object.keys(pivot).forEach(cod => {
    Object.keys(pivot[cod]).forEach(sem => {
      pivot[cod][sem].planVol += state.config.merma;
    });
  });
  // Entregas (solo para CodEq ya en pivot)
  state.entregaData.forEach(r => {
    const sem = state.semanaMap[r.semanaRaw];
    if (!sem || !pivot[r.codEq]) return;
    if (!pivot[r.codEq][sem]) pivot[r.codEq][sem] = { planVol:0, entregaVol:0, wos:[] };
    pivot[r.codEq][sem].entregaVol += r.lts;
  });
  // % = planVol / entregaVol * 100
  Object.values(pivot).forEach(codData => {
    Object.values(codData).forEach(cell => {
      cell.pct = cell.entregaVol > 0
        ? Math.round(cell.planVol / cell.entregaVol * 100)
        : null;
    });
  });
  return pivot;
}

// ── RENDER MASTER ─────────────────────────────────────────
function render() {
  const pivot = buildPivot();
  renderMetrics(pivot);
  renderSemanaFilter();
  renderPivotTable(pivot);
  renderAlertas(pivot);
}

// ── METRICS ───────────────────────────────────────────────
function renderMetrics(pivot) {
  const totalPlan = state.planData.reduce((a, r) => a + r.lts, 0)
                  + Object.keys(pivot).length * Object.values(
                      Object.fromEntries(
                        ['N','N+1','N+2','N+3'].map(s => [s, new Set(
                          state.planData.filter(r => getSemana(r)===s).map(r=>r.codEq)
                        ).size * state.config.merma])
                      )
                    ).reduce((a,v)=>a+v,0);

  // Simpler: totalPlan = sum of all planVol across pivot cells
  let totalPlanV = 0;
  Object.values(pivot).forEach(cd => Object.values(cd).forEach(c => totalPlanV += c.planVol));

  const totalCap  = Object.values(state.config.capacidades).reduce((a,v) => a + (+v||0), 0);
  const utilPct   = totalCap > 0 ? Math.round(totalPlanV / totalCap * 100) : 0;
  const utilCls   = utilPct >= 95 ? 'danger' : utilPct >= 70 ? 'ok' : 'warn';

  const alerts    = computeAlerts(pivot);
  const na        = alerts.dups.length + alerts.deficit.length + alerts.surplus.length;

  // Insumo problems count
  let insumoProb = 0;
  if (state.insumoLoaded) {
    state.planData.forEach(r => {
      const sem = getSemana(r);
      if (!sem) return;
      if (getSemaforoColor(r.wo, sem) === 'red') insumoProb++;
    });
  }

  document.getElementById('m-util').textContent    = state.planLoaded ? utilPct + '%' : '—';
  document.getElementById('m-util').className      = 'metric-value ' + (state.planLoaded ? utilCls : '');
  document.getElementById('m-lts').textContent     = state.planLoaded ? fmtLts(totalPlanV) : '—';
  document.getElementById('m-alerts').textContent  = (state.planLoaded||state.entregaLoaded) ? na : '—';
  document.getElementById('m-alerts').className    = 'metric-value ' + (na > 0 ? 'warn' : 'ok');
  document.getElementById('m-insumo').textContent  = state.insumoLoaded ? insumoProb : '—';
  document.getElementById('m-insumo').className    = 'metric-value ' + (insumoProb > 0 ? 'danger' : 'ok');
  document.getElementById('alert-badge').textContent = na;
}

// ── SEMANA FILTER ─────────────────────────────────────────
function renderSemanaFilter() {
  const pills = ['Todos','N','N+1','N+2','N+3'];
  document.getElementById('semana-filter').innerHTML = pills.map(p => `
    <button class="fpill${p==='N+3'?' n3':''}${p===state.filtroSemana?' on':''}"
      onclick="setFiltro('${p}')">${p === 'Todos' ? 'Todas las semanas' : 'Semana ' + p}</button>
  `).join('');
}

function setFiltro(val) {
  state.filtroSemana = val;
  renderSemanaFilter();
  renderPivotTable(buildPivot());
}

// ── PIVOT TABLE ───────────────────────────────────────────
function renderPivotTable(pivot) {
  const semanas = ['N','N+1','N+2','N+3'];
  let codigos   = Object.keys(pivot);

  if (!codigos.length) {
    document.getElementById('pivot-container').innerHTML =
      `<div class="empty-state"><div class="empty-icon">&#128202;</div>
       <div class="empty-title">Sin datos</div>
       <div class="empty-sub">${state.planLoaded ? 'No hay códigos con WOs.' : 'Importá el Plan.'}</div></div>`;
    return;
  }

  // SORT: filtroSemana primero, luego alfabético
  codigos.sort((a, b) => {
    if (state.filtroSemana !== 'Todos') {
      const aH = pivot[a][state.filtroSemana] ? 1 : 0;
      const bH = pivot[b][state.filtroSemana] ? 1 : 0;
      if (aH !== bH) return bH - aH;
    }
    return a.localeCompare(b);
  });

  const alerts     = computeAlerts(pivot);
  const dupCods    = new Set(alerts.dups.map(d => d.cod));
  const surpCods   = new Set(alerts.surplus.map(s => s.cod));

  const rows = codigos.map(cod => {
    const isExp     = !!state.expanded[cod];
    const isDup     = dupCods.has(cod);
    const hasSurplus = surpCods.has(cod);
    const isTop     = state.filtroSemana !== 'Todos' && pivot[cod][state.filtroSemana];

    const cells = semanas.map(s => {
      const cell = pivot[cod][s];
      if (!cell) return `<td class="pcell-empty pcell${s==='N+3'?' pcell-n3':''}">—</td>`;
      const p   = cell.pct;
      let cls   = 'pcell' + (s==='N+3'?' pcell-n3':'');
      let pTxt  = '—';
      if (p !== null && state.entregaLoaded) {
        pTxt = p + '%';
        cls += p > state.config.maxPct ? ' cell-warn' : p < state.config.minPct ? ' cell-surplus' : ' cell-ok';
      } else {
        cls += ' cell-nodata';
      }
      // Aggregate semaphore for this cell
      let semaHtml = '';
      if (state.insumoLoaded) {
        const colors = cell.wos.map(w => getSemaforoColor(w.wo, s));
        const counts = { green:0, yellow:0, red:0, none:0 };
        colors.forEach(c => counts[c]++);
        const agg = counts.red > 0 ? 'red' : counts.yellow > 0 ? 'yellow' : counts.none === colors.length ? 'none' : 'green';
        semaHtml = `<div class="cell-sema"><span class="sema-dot sema-${agg}" title="Insumos: ${counts.green}✓ ${counts.yellow}~ ${counts.red}✗"></span></div>`;
      }
      return `<td class="${cls}">
        <div class="cell-pct">${pTxt}</div>
        <div class="cell-vols">WO: ${fmtLts(cell.planVol)}${state.entregaLoaded?' / Ent: '+fmtLts(cell.entregaVol):''}</div>
        ${semaHtml}
      </td>`;
    }).join('');

    // Sub-rows: WOs expandidas
    const allWOs    = semanas.flatMap(s => (pivot[cod][s] ? pivot[cod][s].wos : []));
    const surpN2    = hasSurplus;
    const subRows   = allWOs.map(w => {
      const isN3     = w.semLabel === 'N+3';
      const moved    = !!state.movedToN2[w.id];
      const sColor   = state.insumoLoaded ? getSemaforoColor(w.wo, w.semLabel) : 'none';
      const sTip     = state.insumoLoaded ? getSemaforoText(w.wo, w.semLabel) : '';
      return `<tr>
        <td style="padding-left:44px">
          <span class="sema-dot sema-${sColor}" title="${sTip}"></span>
          <span class="wo-num" style="margin-left:6px">${w.wo}</span>
        </td>
        <td><span class="sem-badge${isN3?' sem-n3':''}">${w.semLabel}</span></td>
        <td>${fmtLts(w.lts)} Lts</td>
        <td>${w.estado||'—'}</td>
        <td>${w.fecha||'—'}</td>
        <td>
          ${moved ? '<span class="moved-tag">&#10003; Movida</span>' : ''}
          ${isN3 && surpN2 && !moved
            ? `<button class="btn-mover" onclick="moveToN2(${w.id})">&#8593; Mover a N+2</button>`
            : ''}
        </td>
      </tr>`;
    }).join('');

    const subSection = isExp
      ? `<tr class="sub-row"><td colspan="5">
           <table class="sub-table">
             <thead><tr>
               <th>WO</th><th>Semana</th><th>Lts (+ merma)</th><th>Estado</th><th>Fecha inicio</th><th></th>
             </tr></thead>
             <tbody>${subRows}</tbody>
           </table>
         </td></tr>`
      : '';

    return `
      <tr class="cod-row" onclick="toggleExpand('${codSafe(cod)}')">
        <td><div class="cod-cell">
          <span class="expand-icon${isExp?' open':''}">&#9658;</span>
          <span class="cod-name">${cod}</span>
          ${isDup    ? '<span class="badge b-dup">3+ sem.</span>'    : ''}
          ${hasSurplus ? '<span class="badge b-surplus">Sobrante N+2</span>' : ''}
          ${isTop     ? '<span class="badge b-top">&#9650; filtrada</span>' : ''}
        </div></td>
        ${cells}
      </tr>
      ${subSection}`;
  }).join('');

  const capSummary = renderCapSummary(pivot);

  document.getElementById('pivot-container').innerHTML = `
    <div class="pivot-wrap">
      <table class="pivot-table">
        <thead><tr>
          <th class="th-left">Codigo vino</th>
          ${semanas.map(s =>
            `<th class="th-sem${s==='N+3'?' th-n3':''}">
              ${s==='N+3' ? 'N+3 <span class="n3-tag-sm">+vis.</span>' : 'Sem. '+s}
            </th>`
          ).join('')}
        </tr></thead>
        <tbody>${rows}</tbody>
      </table>
      <div class="pivot-legend">
        <span class="leg-item"><span class="leg-dot ok-dot"></span>En rango</span>
        <span class="leg-item"><span class="leg-dot warn-dot"></span>Deficit (&gt;${state.config.maxPct}%)</span>
        <span class="leg-item"><span class="leg-dot surplus-dot"></span>Sobrante (&lt;${state.config.minPct}%)</span>
        ${state.insumoLoaded ? `
          <span class="leg-item"><span class="sema-dot sema-green"></span>Insumos OK</span>
          <span class="leg-item"><span class="sema-dot sema-yellow"></span>Insumos parcial</span>
          <span class="leg-item"><span class="sema-dot sema-red"></span>Insumos falta</span>` : ''}
        <span class="leg-item" style="font-size:11px;color:#bbb">% = WO Lts / Entrega Lts &nbsp;·&nbsp; Lts incluye ${fmtLts(state.config.merma)} merma</span>
      </div>
    </div>
    ${capSummary}`;
}

function renderCapSummary(pivot) {
  const semanas = ['N','N+1','N+2','N+3'];
  const caps    = state.config.capacidades;
  const totalCap = semanas.reduce((a, s) => a + (+caps[s]||0), 0);
  if (!totalCap) return '';

  const cards = semanas.map(s => {
    const cap   = +caps[s] || 0;
    const planV = Object.values(pivot).reduce((a, cd) => a + (cd[s] ? cd[s].planVol : 0), 0);
    const pct   = cap > 0 ? Math.round(planV / cap * 100) : 0;
    const cls   = pct >= 95 ? 'danger' : pct >= 70 ? 'ok' : 'warn';
    const fillCls = pct >= 95 ? 'fill-danger' : pct >= 70 ? 'fill-ok' : 'fill-warn';
    return `<div class="cap-card">
      <div class="cap-sem">Sem. ${s}</div>
      <div class="cap-bar-wrap"><div class="cap-bar-fill ${fillCls}" style="width:${Math.min(pct,100)}%"></div></div>
      <div class="cap-nums">${fmtLts(planV)} / ${fmtLts(cap)} Lts</div>
      <div class="cap-pct ${cls}">${cap > 0 ? pct+'%' : '—'}</div>
    </div>`;
  }).join('');

  const totalPlanV = semanas.reduce((a, s) =>
    a + Object.values(pivot).reduce((b, cd) => b + (cd[s] ? cd[s].planVol : 0), 0), 0);
  const totalPct = totalCap > 0 ? Math.round(totalPlanV / totalCap * 100) : 0;
  const tCls     = totalPct >= 95 ? 'danger' : totalPct >= 70 ? 'ok' : 'warn';
  const tFill    = totalPct >= 95 ? 'fill-danger' : totalPct >= 70 ? 'fill-ok' : 'fill-warn';
  const totalCard = `<div class="cap-card total">
    <div class="cap-sem">TOTAL</div>
    <div class="cap-bar-wrap"><div class="cap-bar-fill ${tFill}" style="width:${Math.min(totalPct,100)}%"></div></div>
    <div class="cap-nums">${fmtLts(totalPlanV)} / ${fmtLts(totalCap)} Lts</div>
    <div class="cap-pct ${tCls}">${totalPct}%</div>
  </div>`;

  return `<div class="cap-summary">
    <div class="cap-summary-title">Capacidad de fraccionamiento por semana</div>
    <div class="cap-grid">${cards}${totalCard}</div>
  </div>`;
}

// ── ALERTS ────────────────────────────────────────────────
function computeAlerts(pivot) {
  const semanas = ['N','N+1','N+2','N+3'];
  const codigos = Object.keys(pivot);
  const dups = [], deficit = [], surplus = [];

  codigos.forEach(cod => {
    const present = semanas.filter(s => pivot[cod][s] && pivot[cod][s].planVol > 0);
    if (present.length >= 3) dups.push({ cod, sems: present.join(', ') });

    semanas.forEach(s => {
      const cell = pivot[cod][s];
      if (!cell || cell.pct === null) return;
      if (cell.pct > state.config.maxPct)
        deficit.push({ cod, semana:s, pct:cell.pct });
    });

    const n2 = pivot[cod]['N+2'];
    if (n2 && n2.entregaVol > 0 && n2.planVol < n2.entregaVol) {
      const wosN3 = (pivot[cod]['N+3']||{wos:[]}).wos.filter(w => !state.movedToN2[w.id]);
      if (wosN3.length > 0)
        surplus.push({ cod, sobrante: n2.entregaVol - n2.planVol, wosN3 });
    }
  });

  // Insumo problems
  const insumoProb = [];
  if (state.insumoLoaded) {
    state.planData.forEach(r => {
      const sem = getSemana(r);
      if (!sem) return;
      const col = getSemaforoColor(r.wo, sem);
      if (col === 'red' || col === 'yellow')
        insumoProb.push({ wo:r.wo, cod:r.codEq, semana:sem, color:col, text:getSemaforoText(r.wo, sem) });
    });
  }

  return { dups, deficit, surplus, insumoProb };
}

function renderAlertas(pivot) {
  const a = computeAlerts(pivot);
  let html = '';

  if (a.surplus.length) {
    html += `<div class="alert-section">
      <div class="alert-section-title">&#128200; Sobrante en N+2 — sugerir adelantar N+3 (${a.surplus.length})</div>
      ${a.surplus.map(s => `<div class="alert-card surplus">
        <div class="alert-icon">&#128199;</div>
        <div style="flex:1">
          <div class="alert-title">${s.cod} — Sobrante: ${fmtLts(s.sobrante)} Lts en N+2</div>
          <div class="alert-sub">WOs de N+3 disponibles para adelantar:</div>
          <div class="alert-wo-list">
            ${s.wosN3.map(w => `<div class="alert-wo-row">
              <span><strong>${w.wo}</strong> &middot; ${fmtLts(w.lts)} Lts &middot; ${w.fecha||''}</span>
              <button class="btn-mover" onclick="moveToN2(${w.id})">&#8593; Mover a N+2</button>
            </div>`).join('')}
          </div>
        </div>
      </div>`).join('')}
    </div>`;
  }

  if (a.insumoProb.length) {
    const rojos    = a.insumoProb.filter(x => x.color === 'red');
    const amarillos = a.insumoProb.filter(x => x.color === 'yellow');
    html += `<div class="alert-section">
      <div class="alert-section-title">&#128994; Semaforo de insumos (${a.insumoProb.length} WOs con problemas)</div>
      ${rojos.length ? `<div class="alert-card danger">
        <div class="alert-icon">&#128308;</div>
        <div>
          <div class="alert-title">Sin insumos completos (${rojos.length} WOs)</div>
          ${rojos.map(x => `<div style="font-size:12px;margin-top:4px">
            <span class="sema-dot sema-red" style="display:inline-block;vertical-align:middle;margin-right:6px"></span>
            <strong>${x.wo}</strong> &middot; ${x.cod} &middot; Sem. ${x.semana} &middot; ${x.text}
          </div>`).join('')}
        </div>
      </div>` : ''}
      ${amarillos.length ? `<div class="alert-card warn">
        <div class="alert-icon">&#128993;</div>
        <div>
          <div class="alert-title">Insumos parciales (${amarillos.length} WOs)</div>
          ${amarillos.map(x => `<div style="font-size:12px;margin-top:4px">
            <span class="sema-dot sema-yellow" style="display:inline-block;vertical-align:middle;margin-right:6px"></span>
            <strong>${x.wo}</strong> &middot; ${x.cod} &middot; Sem. ${x.semana} &middot; ${x.text}
          </div>`).join('')}
        </div>
      </div>` : ''}
    </div>`;
  }

  if (a.dups.length) {
    html += `<div class="alert-section">
      <div class="alert-section-title">&#9888; Mismo codigo en 3+ semanas (${a.dups.length})</div>
      ${a.dups.map(d => `<div class="alert-card warn">
        <div class="alert-icon">&#128260;</div>
        <div><div class="alert-title">${d.cod}</div>
        <div class="alert-sub">Semanas: ${d.sems}</div></div>
      </div>`).join('')}
    </div>`;
  }

  if (a.deficit.length) {
    html += `<div class="alert-section">
      <div class="alert-section-title">&#128314; Deficit de vino (WOs &gt; Entrega) (${a.deficit.length})</div>
      ${a.deficit.map(d => `<div class="alert-card danger">
        <div class="alert-icon">&#128202;</div>
        <div><div class="alert-title">${d.cod} &middot; Semana ${d.semana}</div>
        <div class="alert-sub">WOs = ${d.pct}% de la entrega (max: ${state.config.maxPct}%)</div></div>
      </div>`).join('')}
    </div>`;
  }

  if (!html) html = `<div class="alert-card ok">
    <div class="alert-icon">&#9989;</div>
    <div><div class="alert-title">Sin alertas activas</div>
    <div class="alert-sub">El plan esta correctamente balanceado.</div></div>
  </div>`;

  document.getElementById('alertas-container').innerHTML = html;
}

// ── CONFIG ────────────────────────────────────────────────
function renderConfig() {
  const caps = state.config.capacidades;
  const total = Object.values(caps).reduce((a,v) => a + (+v||0), 0);
  document.getElementById('config-container').innerHTML = `
    <div class="config-section">
      <div class="config-title">Configuracion de semanas</div>
      <div class="config-grid">
        <div class="config-card">
          <div class="config-card-title">Inicio de semana N</div>
          <div class="field-row">
            <label>Fecha de inicio (lunes de semana N)</label>
            <input type="date" id="cfg-fecha" value="${state.config.semanaInicio}">
            <span class="hint">Las semanas van de lunes a domingo. N+1 = +7 dias, N+2 = +14, N+3 = +21</span>
          </div>
        </div>
        <div class="config-card">
          <div class="config-card-title">Merma de linea</div>
          <div class="field-row">
            <label>Lts por codigo equivalente por semana</label>
            <input type="number" id="cfg-merma" value="${state.config.merma}">
            <span class="hint">Se suma al total de WOs de cada CodEq en cada semana</span>
          </div>
        </div>
      </div>
    </div>

    <div class="config-section">
      <div class="config-title">Capacidad de fraccionamiento por semana</div>
      <div class="config-card">
        <div class="cap-inputs">
          ${['N','N+1','N+2','N+3'].map(s => `
            <div class="cap-input-item">
              <label>Semana ${s} (Lts)</label>
              <input type="number" id="cfg-cap-${s.replace('+','p')}" value="${caps[s]||0}" placeholder="ej: 150000">
            </div>
          `).join('')}
        </div>
        <div class="total-cap">
          <span>Total 4 semanas</span>
          <span id="cfg-cap-total-display"><strong>${fmtLts(total)} Lts</strong></span>
        </div>
      </div>
    </div>

    <div class="config-section">
      <div class="config-title">Rangos de alerta</div>
      <div class="config-grid">
        <div class="config-card">
          <div class="config-card-title">% Minimo (sobrante)</div>
          <div class="field-row">
            <label>Alerta si WO Lts / Entrega Lts es menor a</label>
            <input type="number" id="cfg-min" value="${state.config.minPct}">
            <span class="hint">Celda azul = hay mas entrega que WOs</span>
          </div>
        </div>
        <div class="config-card">
          <div class="config-card-title">% Maximo (deficit)</div>
          <div class="field-row">
            <label>Alerta si WO Lts / Entrega Lts supera</label>
            <input type="number" id="cfg-max" value="${state.config.maxPct}">
            <span class="hint">Celda roja = hay mas WOs que entrega</span>
          </div>
        </div>
      </div>
    </div>

    <div style="display:flex;align-items:center">
      <button class="btn-primary" onclick="saveConfig()">Guardar configuracion</button>
      <span id="cfg-ok" class="save-feedback"></span>
    </div>`;

  // Update total dinamicamente
  ['N','N+1','N+2','N+3'].forEach(s => {
    const el = document.getElementById('cfg-cap-' + s.replace('+','p'));
    if (el) el.addEventListener('input', updateCapTotal);
  });
}

function updateCapTotal() {
  const total = ['N','N+1','N+2','N+3'].reduce((a, s) => {
    const el = document.getElementById('cfg-cap-' + s.replace('+','p'));
    return a + (el ? (+el.value||0) : 0);
  }, 0);
  const el = document.getElementById('cfg-cap-total-display');
  if (el) el.innerHTML = `<strong>${fmtLts(total)} Lts</strong>`;
}

function saveConfig() {
  state.config.semanaInicio = document.getElementById('cfg-fecha')?.value  || state.config.semanaInicio;
  state.config.merma        = parseFloat(document.getElementById('cfg-merma')?.value) || 1000;
  state.config.minPct       = parseFloat(document.getElementById('cfg-min')?.value)   || 80;
  state.config.maxPct       = parseFloat(document.getElementById('cfg-max')?.value)   || 100;
  ['N','N+1','N+2','N+3'].forEach(s => {
    const el = document.getElementById('cfg-cap-' + s.replace('+','p'));
    if (el) state.config.capacidades[s] = parseFloat(el.value) || 0;
  });
  document.getElementById('cfg-ok').textContent = '&#10003; Guardado';
  setTimeout(() => {
    const el = document.getElementById('cfg-ok');
    if (el) el.textContent = '';
  }, 2500);
  if (state.planLoaded || state.entregaLoaded) render();
}

// ── UTILS ─────────────────────────────────────────────────
function fmtLts(v) {
  if (!v) return '0';
  if (v >= 1000000) return (v/1000000).toFixed(1)+'M';
  if (v >= 1000)    return (Math.round(v/100)/10).toFixed(1)+'k';
  return Math.round(v).toString();
}

function fmtFecha(val) {
  if (!val) return '';
  if (val instanceof Date) return val.toLocaleDateString('es-AR');
  return String(val).split(' ')[0];
}

function fmtDateShort(d) {
  if (!d || !(d instanceof Date)) return '—';
  return d.toLocaleDateString('es-AR');
}

function codSafe(cod) {
  return cod.replace(/[^a-zA-Z0-9]/g, '_');
}
