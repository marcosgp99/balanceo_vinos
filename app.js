// ============================================================
// BALANCEO DE CARGAS WO — Logica principal
// ============================================================

// --- ESTADO DE LA APLICACION ---
const state = {
  wos: [],
  config: {
    MI:      { cap: 1100, minPct: 60, maxPct: 115 },
    StComex: { cap: 700,  minPct: 60, maxPct: 110 },
    Estibas: { cap: 400,  minPct: 0,  maxPct: 105 }
  },
  selWeek: 'N',
  rawHeaders: [],
  rawData: [],
  colMap: { wo:'', cod:'', semana:'', vol:'', tipo:'' },
  nextId: 1
};

// --- INICIALIZACION ---
document.addEventListener('DOMContentLoaded', () => {
  renderConfig();
  renderWeekPills();

  document.getElementById('fileInput')
    .addEventListener('change', handleFileUpload);

  document.querySelectorAll('.tab').forEach(tab => {
    tab.addEventListener('click', () => {
      state.activeTab = tab.dataset.tab;
      document.querySelectorAll('.tab').forEach(t => t.classList.remove('on'));
      document.querySelectorAll('.panel').forEach(p => p.classList.remove('on'));
      tab.classList.add('on');
      document.getElementById('panel-' + tab.dataset.tab).classList.add('on');
    });
  });
});

// --- IMPORTAR EXCEL ---
function handleFileUpload(e) {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function(evt) {
    const wb = XLSX.read(evt.target.result, { type: 'binary' });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    if (rows.length < 2) { alert('Archivo vacio o sin datos.'); return; }
    state.rawHeaders = rows[0].map(h => String(h).trim());
    state.rawData = rows.slice(1).filter(r => r.some(c => c !== ''));
    const detected = autoDetectColumns(state.rawHeaders);
    state.colMap = detected;
    showColumnMapper(state.rawHeaders, detected);
  };
  reader.readAsBinaryString(file);
  e.target.value = '';
}

function autoDetectColumns(headers) {
  const map = { wo:'', cod:'', semana:'', vol:'', tipo:'' };
  headers.forEach(h => {
    const hl = h.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'');
    if (!map.wo     && /\b(wo|orden|order|nro|numero|num)\b/.test(hl))         map.wo = h;
    if (!map.cod    && /\b(vino|cod|codigo|wine|producto|prod|sku)\b/.test(hl)) map.cod = h;
    if (!map.semana && /\b(semana|week|fecha|date|periodo|sem)\b/.test(hl))     map.semana = h;
    if (!map.vol    && /\b(vol|volumen|cantidad|hl|litros|lts|qty)\b/.test(hl)  ) map.vol = h;
    if (!map.tipo   && /\b(tipo|type|categoria|bodega|clase)\b/.test(hl))       map.tipo = h;
  });
  return map;
}

function showColumnMapper(headers, detected) {
  const fields = [
    { key:'wo',     label:'Numero de WO'       },
    { key:'cod',    label:'Codigo de vino'     },
    { key:'semana', label:'Semana / Fecha'     },
    { key:'vol',    label:'Volumen (hl)'       },
    { key:'tipo',   label:'Tipo (MI/StComex/Estibas)' }
  ];
  document.getElementById('col-map-fields').innerHTML = fields.map(f => `
    <div class="col-map-row">
      <label>${f.label}</label>
      <select id="map-${f.key}">
        <option value="">— seleccionar —</option>
        ${headers.map(h => `<option value="${h}" ${h===detected[f.key]?'selected':''}>${h}</option>`).join('')}
      </select>
    </div>
  `).join('');
  document.getElementById('modal-overlay').classList.remove('hidden');
}

function confirmImport() {
  const fields = ['wo','cod','semana','vol','tipo'];
  const newMap = {};
  let valid = true;
  fields.forEach(f => {
    const val = document.getElementById('map-' + f).value;
    if (!val) valid = false;
    newMap[f] = val;
  });
  if (!valid) { alert('Por favor asigna todas las columnas.'); return; }
  state.colMap = newMap;
  state.wos = processExcelData(state.rawData, newMap);
  cerrarModal();
  render();
}

function cerrarModal() {
  document.getElementById('modal-overlay').classList.add('hidden');
}

function processExcelData(rows, colMap) {
  const headers = state.rawHeaders;
  const idx = {};
  Object.keys(colMap).forEach(k => { idx[k] = headers.indexOf(colMap[k]); });

  const semanaVals = [...new Set(
    rows.map(r => String(r[idx.semana] || '').trim()).filter(Boolean)
  )].sort((a,b) => {
    const na = parseFloat(a), nb = parseFloat(b);
    return (!isNaN(na) && !isNaN(nb)) ? na - nb : a.localeCompare(b);
  });

  const semanaMap = {};
  const labels = ['N','N+1','N+2','N+3'];
  semanaVals.forEach((s,i) => { semanaMap[s] = labels[i] || 'N+3'; });

  return rows.map((r, i) => {
    const tipoRaw = String(r[idx.tipo] || '').trim().toLowerCase();
    let tipo = 'MI';
    if (tipoRaw.includes('estiba')) tipo = 'Estibas';
    else if (tipoRaw.includes('comex') || tipoRaw.includes('stc')) tipo = 'StComex';

    return {
      id: state.nextId++,
      wo:     String(r[idx.wo]  || 'WO-' + (i+1)).trim(),
      cod:    String(r[idx.cod] || '').trim(),
      semana: semanaMap[String(r[idx.semana]||'').trim()] || 'N',
      vol:    parseFloat(r[idx.vol]) || 0,
      tipo,
      prioridad: 'Normal',
      pulled: false
    };
  }).filter(w => w.cod);
}

// --- DATOS DE EJEMPLO ---
function cargarDemoData() {
  state.wos = [
    {id:1, wo:'WO-1001', cod:'VIN-032', tipo:'MI',      semana:'N',   vol:950,  prioridad:'Alta',   pulled:false},
    {id:2, wo:'WO-1002', cod:'VIN-017', tipo:'MI',      semana:'N',   vol:820,  prioridad:'Normal', pulled:false},
    {id:3, wo:'WO-1003', cod:'VIN-055', tipo:'StComex', semana:'N',   vol:650,  prioridad:'Normal', pulled:false},
    {id:4, wo:'WO-1004', cod:'VIN-032', tipo:'MI',      semana:'N+1', vol:1050, prioridad:'Normal', pulled:false},
    {id:5, wo:'WO-1005', cod:'VIN-044', tipo:'MI',      semana:'N+1', vol:760,  prioridad:'Normal', pulled:false},
    {id:6, wo:'WO-1006', cod:'VIN-017', tipo:'StComex', semana:'N+1', vol:490,  prioridad:'Baja',   pulled:false},
    {id:7, wo:'WO-1007', cod:'VIN-032', tipo:'MI',      semana:'N+2', vol:1020, prioridad:'Normal', pulled:false},
    {id:8, wo:'WO-1008', cod:'VIN-062', tipo:'StComex', semana:'N+2', vol:530,  prioridad:'Normal', pulled:false},
    {id:9, wo:'WO-1009', cod:'VIN-081', tipo:'Estibas', semana:'N+2', vol:310,  prioridad:'Baja',   pulled:false},
    {id:10,wo:'WO-1010', cod:'VIN-032', tipo:'MI',      semana:'N+3', vol:990,  prioridad:'Normal', pulled:false},
    {id:11,wo:'WO-1011', cod:'VIN-099', tipo:'MI',      semana:'N+3', vol:870,  prioridad:'Alta',   pulled:false},
  ];
  state.nextId = 20;
  render();
}

// --- OCUPACION ---
function getOcc(w) {
  const cap = (state.config[w.tipo] || {cap:1000}).cap;
  return cap > 0 ? Math.round(w.vol / cap * 100) : 0;
}
function occFillClass(p) { return p>=100?'occ-danger':p>=85?'occ-warn':'occ-ok'; }
function occColor(p)     { return p>=100?'#A32D2D':p>=85?'#BA7517':'#0F6E56'; }

// --- ALERTAS ---
function getAlerts() {
  const a = { duplicados:[], sobrecap:[], estibas:[], prioAlta:[], n3disponibles:[] };

  const byCode = {};
  state.wos.forEach(w => {
    if (!byCode[w.cod]) byCode[w.cod] = new Set();
    byCode[w.cod].add(w.semana);
  });
  Object.entries(byCode).forEach(([cod, sems]) => {
    if (sems.size >= 3) a.duplicados.push({ cod, sems:[...sems].join(', ') });
  });

  state.wos.forEach(w => {
    const cfg = state.config[w.tipo]; if (!cfg) return;
    const p = getOcc(w);
    if (p > cfg.maxPct) a.sobrecap.push({ wo:w.wo, cod:w.cod, tipo:w.tipo, p, max:cfg.maxPct });
    if (cfg.minPct > 0 && p < cfg.minPct) a.sobrecap.push({ wo:w.wo, cod:w.cod, tipo:w.tipo, p, max:cfg.minPct, bajo:true });
  });

  state.wos.filter(w => w.tipo==='Estibas').forEach(w =>
    a.estibas.push({ wo:w.wo, cod:w.cod, semana:w.semana })
  );

  state.wos.filter(w => w.prioridad==='Alta').forEach(w =>
    a.prioAlta.push({ wo:w.wo, cod:w.cod, semana:w.semana, tipo:w.tipo })
  );

  a.n3disponibles = state.wos.filter(w => w.semana==='N+3' && !w.pulled);

  return a;
}

function countAlerts(a) {
  return a.duplicados.length + a.sobrecap.length + a.estibas.length;
}

// --- RENDER MASTER ---
function render() {
  renderMetrics();
  renderWeekPills();
  renderPlanTable();
  renderAlertas();
}

// --- METRICAS ---
function renderMetrics() {
  const vis = state.wos.filter(w => w.semana !== 'N+3');
  const n3  = state.wos.filter(w => w.semana === 'N+3' && !w.pulled);
  const alerts = getAlerts();
  const na = countAlerts(alerts);
  const avgOcc = vis.length ? Math.round(vis.reduce((a,w)=>a+getOcc(w),0)/vis.length) : 0;
  const occCls = avgOcc>=100?'danger':avgOcc>=85?'warn':'ok';

  document.getElementById('m-wos').textContent    = vis.length;
  document.getElementById('m-wos').className      = 'metric-value';
  document.getElementById('m-occ').textContent    = vis.length ? avgOcc+'%' : '—';
  document.getElementById('m-occ').className      = 'metric-value ' + occCls;
  document.getElementById('m-alerts').textContent = na;
  document.getElementById('m-alerts').className   = 'metric-value ' + (na>0?'warn':'ok');
  document.getElementById('m-n3').textContent     = n3.length;
  document.getElementById('alert-badge').textContent = na;
}

// --- WEEK PILLS ---
function renderWeekPills() {
  const weeks = ['N','N+1','N+2','N+3'];
  document.getElementById('week-pills').innerHTML = weeks.map(w => {
    const cnt = state.wos.filter(x => x.semana===w).length;
    return `<button class="wpill ${w==='N+3'?'n3':''} ${w===state.selWeek?'on':''}"
      onclick="selectWeek('${w}')">
      Sem. ${w} <span style="font-size:10px;opacity:.6">(${cnt})</span>
      ${w==='N+3'?'<span class="n3-tag">+visibilidad</span>':''}
    </button>`;
  }).join('');
}

function selectWeek(w) {
  state.selWeek = w;
  renderWeekPills();
  renderPlanTable();
}

// --- TABLA PLAN ---
function renderPlanTable() {
  const container = document.getElementById('plan-table-container');
  const prioOrder = { Alta:0, Normal:1, Baja:2 };
  const list = state.wos
    .filter(w => w.semana === state.selWeek)
    .sort((a,b) => prioOrder[a.prioridad] - prioOrder[b.prioridad]);

  if (!list.length) {
    container.innerHTML = `<div class="empty-state">
      <div class="empty-icon">📋</div>
      <div class="empty-title">Sin WOs en semana ${state.selWeek}</div>
      <div class="empty-sub">No hay ordenes asignadas a esta semana.</div>
    </div>`;
    return;
  }

  const alerts  = getAlerts();
  const dupCods = new Set(alerts.duplicados.map(d => d.cod));
  const overWOs = new Set(alerts.sobrecap.map(d => d.wo));
  const isN3    = state.selWeek === 'N+3';

  container.innerHTML = `<div class="wo-table-wrap">
    <table class="wo-table">
      <thead><tr>
        <th>Nro. WO</th>
        <th>Codigo vino</th>
        <th>Tipo</th>
        <th>Volumen</th>
        <th>Ocupacion</th>
        <th>Prioridad</th>
        <th>Accion</th>
      </tr></thead>
      <tbody>
        ${list.map(w => {
          const p = getOcc(w);
          const rc = overWOs.has(w.wo)?'row-over':dupCods.has(w.cod)?'row-dup':w.prioridad==='Alta'?'row-alta':'';
          return `<tr class="${rc}">
            <td><strong>${w.wo}</strong></td>
            <td>${w.cod}${dupCods.has(w.cod)?'<span class="badge b-dup">3 sem.</span>':''}</td>
            <td><span class="badge ${tipoBadge(w.tipo)}">${w.tipo}</span></td>
            <td>${w.vol.toLocaleString()} hl</td>
            <td>
              <div class="occ-wrap">
                <div class="occ-bar"><div class="occ-fill ${occFillClass(p)}" style="width:${Math.min(p,100)}%"></div></div>
                <span class="occ-label" style="color:${occColor(p)}">${p}%</span>
              </div>
            </td>
            <td>
              <select class="prio-select" onchange="setPriority(${w.id},this.value)">
                <option ${w.prioridad==='Alta'?'selected':''}>Alta</option>
                <option ${w.prioridad==='Normal'?'selected':''}>Normal</option>
                <option ${w.prioridad==='Baja'?'selected':''}>Baja</option>
              </select>
            </td>
            <td>
              ${isN3 && !w.pulled ? `<button class="btn-sm" onclick="pullFromN3(${w.id})">↑ Usar en plan</button>` : ''}
              ${w.pulled ? '<span class="btn-pulled">✓ Incluida</span>' : ''}
              ${!isN3 ? `<button class="btn-sm del" onclick="eliminarWO(${w.id})" title="Eliminar">✕</button>` : ''}
            </td>
          </tr>`;
        }).join('')}
      </tbody>
    </table>
  </div>`;
}

function tipoBadge(tipo) {
  return tipo==='MI'?'b-mi':tipo==='StComex'?'b-stc':'b-est';
}

// --- ALERTAS TAB ---
function renderAlertas() {
  const a = getAlerts();
  let html = '';

  if (a.duplicados.length) {
    html += `<div class="alert-section">
      <div class="alert-section-title">⚠ Mismo vino en 3+ semanas (${a.duplicados.length})</div>
      ${a.duplicados.map(d => `
        <div class="alert-card warn">
          <div class="alert-icon">🔄</div>
          <div><div class="alert-title">${d.cod} repetido</div>
          <div class="alert-sub">Semanas: ${d.sems} · Considera usar WO de N+3 anticipadamente.</div></div>
        </div>`).join('')}
    </div>`;
  }

  if (a.sobrecap.length) {
    html += `<div class="alert-section">
      <div class="alert-section-title">📊 Ocupacion fuera de rango (${a.sobrecap.length})</div>
      ${a.sobrecap.map(d => `
        <div class="alert-card danger">
          <div class="alert-icon">${d.bajo?'🔽':'🔺'}</div>
          <div><div class="alert-title">${d.wo} — ${d.cod} · ${d.tipo}</div>
          <div class="alert-sub">Ocupacion: ${d.p}% · ${d.bajo?'Minimo':'Maximo'}: ${d.max}%</div></div>
        </div>`).join('')}
    </div>`;
  }

  if (a.estibas.length) {
    html += `<div class="alert-section">
      <div class="alert-section-title">⛔ WOs en Estibas — evitar (${a.estibas.length})</div>
      ${a.estibas.map(e => `
        <div class="alert-card danger">
          <div class="alert-icon">🏗</div>
          <div><div class="alert-title">${e.wo} — ${e.cod} · Semana ${e.semana}</div>
          <div class="alert-sub">Estibas es prioridad 3. Reasignar a MI o StComex si es posible.</div></div>
        </div>`).join('')}
    </div>`;
  }

  if (a.prioAlta.length) {
    html += `<div class="alert-section">
      <div class="alert-section-title">🔺 Ordenes con prioridad alta (${a.prioAlta.length})</div>
      ${a.prioAlta.map(p => `
        <div class="alert-card info">
          <div class="alert-icon">⭐</div>
          <div><div class="alert-title">${p.wo} — ${p.cod}</div>
          <div class="alert-sub">Semana ${p.semana} · Tipo ${p.tipo}</div></div>
        </div>`).join('')}
    </div>`;
  }

  if (a.n3disponibles.length) {
    html += `<div class="alert-section">
      <div class="alert-section-title">📅 Disponibles en N+3 para adelantar (${a.n3disponibles.length})</div>
      <div class="alert-card info">
        <div class="alert-icon">👁</div>
        <div><div class="alert-title">${a.n3disponibles.length} WO${a.n3disponibles.length>1?'s':''} disponible${a.n3disponibles.length>1?'s':''}</div>
        <div class="alert-sub">${a.n3disponibles.map(w=>w.wo).join(', ')} · Usa el boton "Usar en plan" en semana N+3.</div></div>
      </div>
    </div>`;
  }

  if (!html) html = `<div class="alert-card ok">
    <div class="alert-icon">✅</div>
    <div><div class="alert-title">Sin alertas activas</div>
    <div class="alert-sub">El plan esta correctamente balanceado.</div></div>
  </div>`;

  document.getElementById('alertas-container').innerHTML = html;
}

// --- CONFIG TAB ---
function renderConfig() {
  document.getElementById('config-container').innerHTML = `
    <div class="config-section">
      <div class="config-title">Capacidades y tolerancias por tipo</div>
      <div class="config-grid">
        ${['MI','StComex','Estibas'].map(t => {
          const cfg = state.config[t];
          return `<div class="config-card">
            <div class="config-card-title"><span class="badge ${tipoBadge(t)}">${t}</span></div>
            <div class="field-row">
              <label>Capacidad base (hl)</label>
              <input type="number" id="cfg-cap-${t}" value="${cfg.cap}">
            </div>
            <div class="field-row">
              <label>% minimo de ocupacion</label>
              <input type="number" id="cfg-min-${t}" value="${cfg.minPct}">
              <span class="hint">Alerta si la WO esta por debajo de este %</span>
            </div>
            <div class="field-row">
              <label>% maximo de ocupacion</label>
              <input type="number" id="cfg-max-${t}" value="${cfg.maxPct}">
              <span class="hint">Alerta si la WO supera este %</span>
            </div>
          </div>`;
        }).join('')}
      </div>
      <div style="display:flex;align-items:center;margin-top:16px">
        <button class="btn-primary" onclick="saveConfig()">Guardar configuracion</button>
        <span id="cfg-ok" class="save-feedback"></span>
      </div>
    </div>
    <div class="config-section">
      <div class="config-title">Proceso de balanceo</div>
      <div style="display:flex;flex-direction:column;gap:8px;max-width:520px">
        <div class="alert-card ok"><div class="alert-icon">1</div>
          <div><div class="alert-title"><span class="badge b-mi">MI</span> Prioridad 1</div>
          <div class="alert-sub">Tolerancia: +15% / -10% sobre capacidad base</div></div></div>
        <div class="alert-card info"><div class="alert-icon">2</div>
          <div><div class="alert-title"><span class="badge b-stc">StComex</span> Prioridad 2</div>
          <div class="alert-sub">Tolerancia: +10% / -10% sobre capacidad base</div></div></div>
        <div class="alert-card danger"><div class="alert-icon">3</div>
          <div><div class="alert-title"><span class="badge b-est">Estibas</span> Prioridad 3 — evitar</div>
          <div class="alert-sub">Usar solo como ultimo recurso. Genera alerta automatica.</div></div></div>
      </div>
    </div>`;
}

function saveConfig() {
  ['MI','StComex','Estibas'].forEach(t => {
    state.config[t].cap    = parseFloat(document.getElementById('cfg-cap-'+t).value) || state.config[t].cap;
    state.config[t].minPct = parseFloat(document.getElementById('cfg-min-'+t).value) || 0;
    state.config[t].maxPct = parseFloat(document.getElementById('cfg-max-'+t).value) || state.config[t].maxPct;
  });
  const fb = document.getElementById('cfg-ok');
  fb.textContent = '✓ Guardado';
  setTimeout(() => fb.textContent='', 2500);
  render();
}

// --- ACCIONES ---
function pullFromN3(id) {
  const w = state.wos.find(x => x.id===id);
  if (w) { w.semana='N+2'; w.pulled=true; }
  render();
}

function eliminarWO(id) {
  if (!confirm('Eliminar esta WO del plan?')) return;
  state.wos = state.wos.filter(w => w.id!==id);
  render();
}

function setPriority(id, prio) {
  const w = state.wos.find(x => x.id===id);
  if (w) w.prioridad = prio;
  renderMetrics();
  renderPlanTable();
  renderAlertas();
}
