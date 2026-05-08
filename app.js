// ============================================================
// BALANCEO DE CARGAS WO v2 — Vista pivot por codigo equivalente
// ============================================================

const state = {
  planData:    [],
  entregaData: [],
  semanaMap:   {},
  config: {
    tipos: {
      MI:      { minPct: 70, maxPct: 115 },
      StComex: { minPct: 70, maxPct: 110 },
      Estibas: { minPct: 0,  maxPct: 105 }
    }
  },
  importMode:    'plan',
  rawHeaders:    [],
  rawData:       [],
  colMap:        { codEq:'', semana:'', lts:'', tipo:'' },
  nextId:        1,
  planLoaded:    false,
  entregaLoaded: false
};

// --- INIT ---
document.addEventListener('DOMContentLoaded', () => {
  renderConfig();

  document.getElementById('filePlan')
    .addEventListener('change', e => handleFileUpload(e, 'plan'));
  document.getElementById('fileEntrega')
    .addEventListener('change', e => handleFileUpload(e, 'entregas'));

  document.querySelectorAll('.tab').forEach(tab => {
    tab.addEventListener('click', () => {
      document.querySelectorAll('.tab').forEach(t => t.classList.remove('on'));
      document.querySelectorAll('.panel').forEach(p => p.classList.remove('on'));
      tab.classList.add('on');
      document.getElementById('panel-' + tab.dataset.tab).classList.add('on');
    });
  });
});

// --- IMPORTAR EXCEL ---
function handleFileUpload(e, mode) {
  const file = e.target.files[0];
  if (!file) return;
  state.importMode = mode;
  const reader = new FileReader();
  reader.onload = function(evt) {
    const wb = XLSX.read(evt.target.result, { type: 'binary' });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    if (rows.length < 2) { alert('El archivo parece estar vacio.'); return; }
    state.rawHeaders = rows[0].map(h => String(h).trim());
    state.rawData    = rows.slice(1).filter(r => r.some(c => c !== ''));
    const detected   = autoDetectColumns(state.rawHeaders);
    state.colMap     = detected;
    showColumnMapper(state.rawHeaders, detected, mode);
  };
  reader.readAsBinaryString(file);
  e.target.value = '';
}

function autoDetectColumns(headers) {
  const map = { codEq:'', semana:'', lts:'', tipo:'' };
  headers.forEach(h => {
    const hl = h.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    if (!map.codEq  && /codeq|cod.eq|equivalente|equiv/.test(hl))     map.codEq  = h;
    if (!map.semana && /\bsemana\b|\bweek\b|\bsem\b/.test(hl))        map.semana = h;
    if (!map.lts    && /^lts$|^litros$|^volumen$|^vol$/.test(hl))     map.lts    = h;
    if (!map.tipo   && /^tipo$|ordentipo|^type$/.test(hl))             map.tipo   = h;
  });
  // Fallbacks sin word boundary (SAP export puede variar)
  if (!map.codEq)  map.codEq  = headers.find(h => /codeq/i.test(h))  || '';
  if (!map.semana) map.semana = headers.find(h => /semana/i.test(h))  || '';
  if (!map.lts)    map.lts    = headers.find(h => /^lts$/i.test(h))   || '';
  if (!map.tipo)   map.tipo   = headers.find(h => /tipo/i.test(h))    || '';
  return map;
}

function showColumnMapper(headers, detected, mode) {
  const isPlan = mode === 'plan';
  document.getElementById('modal-title').textContent =
    isPlan ? 'Columnas del Plan' : 'Columnas de Entregas';
  document.getElementById('modal-desc').textContent =
    isPlan
      ? 'Selecciona las columnas del archivo de Plan:'
      : 'Selecciona las columnas del archivo de Entregas:';

  const fields = [
    { key:'codEq',  label:'Codigo equivalente (CodEq)', req:true  },
    { key:'semana', label:'Semana',                      req:true  },
    { key:'lts',    label:'Volumen (Lts)',                req:true  },
    { key:'tipo',   label:'Tipo (MI / StComex / Estibas)', req:false }
  ];

  document.getElementById('col-map-fields').innerHTML = fields.map(f => `
    <div class="col-map-row">
      <label>${f.label}${f.req ? '' : ' <span style="color:#bbb;font-size:11px">(opcional)</span>'}</label>
      <select id="map-${f.key}">
        ${!f.req ? '<option value="">— no usar —</option>' : '<option value="">— seleccionar —</option>'}
        ${headers.map(h =>
          `<option value="${h}" ${h === detected[f.key] ? 'selected' : ''}>${h}</option>`
        ).join('')}
      </select>
    </div>
  `).join('');

  document.getElementById('modal-overlay').classList.remove('hidden');
}

function confirmImport() {
  const newMap = {};
  ['codEq','semana','lts','tipo'].forEach(f => {
    const el = document.getElementById('map-' + f);
    if (el) newMap[f] = el.value;
  });
  if (!newMap.codEq || !newMap.semana || !newMap.lts) {
    alert('Asigna las columnas obligatorias: CodEq, Semana y Lts.');
    return;
  }
  state.colMap = newMap;
  const processed = processData(state.rawData, newMap);

  if (state.importMode === 'plan') {
    state.planData   = processed;
    state.planLoaded = true;
    updateStatus('plan', processed.length + ' filas cargadas');
  } else {
    state.entregaData   = processed;
    state.entregaLoaded = true;
    updateStatus('entregas', processed.length + ' filas cargadas');
  }
  cerrarModal();
  buildSemanaMap();
  render();
}

function cerrarModal() {
  document.getElementById('modal-overlay').classList.add('hidden');
}

function processData(rows, colMap) {
  const headers = state.rawHeaders;
  const idx = {};
  Object.keys(colMap).forEach(k => {
    idx[k] = colMap[k] ? headers.indexOf(colMap[k]) : -1;
  });

  return rows.map((r, i) => {
    const tipoRaw = idx.tipo >= 0 ? String(r[idx.tipo] || '').toLowerCase() : '';
    let tipo = 'MI';
    if (tipoRaw.includes('estiba'))                             tipo = 'Estibas';
    else if (tipoRaw.includes('comex') || tipoRaw.includes('stc')) tipo = 'StComex';

    return {
      id:        state.nextId++,
      codEq:     String(r[idx.codEq]  || '').trim(),
      semanaRaw: String(r[idx.semana] || '').trim(),
      lts:       parseFloat(r[idx.lts]) || 0,
      tipo
    };
  }).filter(r => r.codEq && r.lts > 0);
}

// --- NORMALIZACION DE SEMANAS ---
function buildSemanaMap() {
  const allWeeks = new Set([
    ...state.planData.map(r => r.semanaRaw),
    ...state.entregaData.map(r => r.semanaRaw)
  ]);
  const sorted = [...allWeeks].filter(Boolean).sort((a, b) => {
    const na = parseFloat(a), nb = parseFloat(b);
    return (!isNaN(na) && !isNaN(nb)) ? na - nb : a.localeCompare(b);
  });
  const labels = ['N', 'N+1', 'N+2', 'N+3'];
  state.semanaMap = {};
  sorted.forEach((w, i) => { state.semanaMap[w] = labels[i] || 'N+3'; });

  const str = sorted.map((w, i) => `${labels[i] || 'N+3'}=sem.${w}`).join(' · ');
  document.getElementById('status-semanas').textContent = str;
  document.getElementById('status-semanas').style.color = '#888';
  document.getElementById('status-semanas').style.fontSize = '11px';
}

function updateStatus(type, text) {
  const id   = type === 'plan' ? 'status-plan' : 'status-entrega';
  const el   = document.getElementById(id);
  const label = type === 'plan' ? 'Plan: ' : 'Entregas: ';
  el.querySelector('.status-dot').className = 'status-dot dot-on';
  el.querySelector('span:last-child').textContent = label + text;
}

// --- DATOS DE EJEMPLO ---
function cargarDemoData() {
  state.planData = [
    { id:1,  codEq:'C1079',       semanaRaw:'22', lts:7141.5, tipo:'MI'      },
    { id:2,  codEq:'C1079',       semanaRaw:'22', lts:7012.1, tipo:'MI'      },
    { id:3,  codEq:'A1039/A1497', semanaRaw:'22', lts:5200.0, tipo:'MI'      },
    { id:4,  codEq:'A1039/A1497', semanaRaw:'23', lts:5100.0, tipo:'MI'      },
    { id:5,  codEq:'C1079',       semanaRaw:'23', lts:8500.0, tipo:'StComex' },
    { id:6,  codEq:'B2201',       semanaRaw:'23', lts:3800.0, tipo:'MI'      },
    { id:7,  codEq:'A1039/A1497', semanaRaw:'24', lts:5050.0, tipo:'MI'      },
    { id:8,  codEq:'B2201',       semanaRaw:'24', lts:3600.0, tipo:'MI'      },
    { id:9,  codEq:'C1079',       semanaRaw:'25', lts:7200.0, tipo:'MI'      },
    { id:10, codEq:'D3310',       semanaRaw:'25', lts:4100.0, tipo:'Estibas' },
  ];
  state.entregaData = [
    { id:11, codEq:'C1079',       semanaRaw:'22', lts:12800.0, tipo:'MI'      },
    { id:12, codEq:'A1039/A1497', semanaRaw:'22', lts:4950.0,  tipo:'MI'      },
    { id:13, codEq:'A1039/A1497', semanaRaw:'23', lts:5300.0,  tipo:'MI'      },
    { id:14, codEq:'C1079',       semanaRaw:'23', lts:7200.0,  tipo:'StComex' },
    { id:15, codEq:'B2201',       semanaRaw:'23', lts:4200.0,  tipo:'MI'      },
    { id:16, codEq:'A1039/A1497', semanaRaw:'24', lts:4800.0,  tipo:'MI'      },
    { id:17, codEq:'B2201',       semanaRaw:'24', lts:3400.0,  tipo:'MI'      },
    { id:18, codEq:'C1079',       semanaRaw:'25', lts:7500.0,  tipo:'MI'      },
    { id:19, codEq:'D3310',       semanaRaw:'25', lts:3200.0,  tipo:'Estibas' },
  ];
  state.planLoaded    = true;
  state.entregaLoaded = true;
  state.nextId        = 30;
  updateStatus('plan',     '10 filas cargadas (demo)');
  updateStatus('entregas',  '9 filas cargadas (demo)');
  buildSemanaMap();
  render();
}

// --- PIVOT ---
function buildPivot() {
  const pivot = {};

  state.planData.forEach(r => {
    const sem = state.semanaMap[r.semanaRaw];
    if (!sem) return;
    if (!pivot[r.codEq]) pivot[r.codEq] = {};
    if (!pivot[r.codEq][sem]) pivot[r.codEq][sem] = { planVol:0, entregaVol:0, tipo: r.tipo };
    pivot[r.codEq][sem].planVol += r.lts;
  });

  state.entregaData.forEach(r => {
    const sem = state.semanaMap[r.semanaRaw];
    if (!sem) return;
    if (!pivot[r.codEq]) pivot[r.codEq] = {};
    if (!pivot[r.codEq][sem]) pivot[r.codEq][sem] = { planVol:0, entregaVol:0, tipo:'' };
    pivot[r.codEq][sem].entregaVol += r.lts;
  });

  Object.values(pivot).forEach(codData => {
    Object.values(codData).forEach(cell => {
      cell.pct = cell.planVol > 0
        ? Math.round(cell.entregaVol / cell.planVol * 100)
        : null;
    });
  });

  return pivot;
}

// --- RENDER MASTER ---
function render() {
  const pivot = buildPivot();
  renderMetrics(pivot);
  renderPivotTable(pivot);
  renderAlertas(pivot);
}

// --- METRICAS ---
function renderMetrics(pivot) {
  const semanas = ['N','N+1','N+2','N+3'];
  const codigos = Object.keys(pivot);
  let sum = 0, cnt = 0;
  codigos.forEach(cod => semanas.forEach(s => {
    const c = pivot[cod][s];
    if (c && c.pct !== null) { sum += c.pct; cnt++; }
  }));
  const avg    = cnt > 0 ? Math.round(sum / cnt) : 0;
  const occCls = avg >= 100 ? 'danger' : avg >= 90 ? 'warn' : 'ok';
  const alerts = computeAlerts(pivot);
  const na     = alerts.dups.length + alerts.over.length + alerts.estibas.length;
  const n3cnt  = codigos.filter(c => pivot[c]['N+3']).length;

  document.getElementById('m-cods').textContent    = codigos.length || '—';
  document.getElementById('m-occ').textContent     = cnt > 0 ? avg + '%' : '—';
  document.getElementById('m-occ').className       = 'metric-value ' + (cnt > 0 ? occCls : '');
  document.getElementById('m-alerts').textContent  = (state.planLoaded || state.entregaLoaded) ? na : '—';
  document.getElementById('m-alerts').className    = 'metric-value ' + (na > 0 ? 'warn' : 'ok');
  document.getElementById('m-n3').textContent      = n3cnt || '—';
  document.getElementById('alert-badge').textContent = na;
}

// --- TABLA PIVOT ---
function renderPivotTable(pivot) {
  const semanas = ['N','N+1','N+2','N+3'];
  const codigos = Object.keys(pivot).sort();

  if (!codigos.length) {
    const msg = !state.planLoaded
      ? 'Importa el Plan usando el boton "Importar Plan".'
      : 'Importa las Entregas para calcular el % de ocupacion.';
    document.getElementById('pivot-container').innerHTML =
      `<div class="empty-state"><div class="empty-icon">&#128202;</div>
       <div class="empty-title">Sin datos</div>
       <div class="empty-sub">${msg}</div></div>`;
    return;
  }

  const alerts  = computeAlerts(pivot);
  const dupCods = new Set(alerts.dups.map(d => d.cod));

  let rows = codigos.map(cod => {
    const isDup   = dupCods.has(cod);
    const anyCell = Object.values(pivot[cod])[0] || {};
    const tipo    = anyCell.tipo || '';

    const cells = semanas.map(s => {
      const cell = pivot[cod][s];
      if (!cell) return `<td class="td-empty">—</td>`;

      const p   = cell.pct;
      const onlyPlan = cell.planVol > 0 && !state.entregaLoaded;
      let cls = 'td-cell' + (s === 'N+3' ? ' td-n3' : '');
      if (p !== null) {
        if (p >= 100) cls += ' cell-over';
        else if (p >= 90) cls += ' cell-warn';
        else cls += ' cell-ok';
      } else if (onlyPlan) {
        cls += ' cell-noent';
      }

      const pctTxt  = p !== null && state.entregaLoaded ? p + '%' : onlyPlan ? '—' : '—';
      const planTxt = cell.planVol    > 0 ? 'P: ' + fmtLts(cell.planVol)    : '';
      const entTxt  = cell.entregaVol > 0 ? ' / E: ' + fmtLts(cell.entregaVol) : '';

      return `<td class="${cls}">
        <div class="cell-pct">${pctTxt}</div>
        <div class="cell-vols">${planTxt}${state.entregaLoaded ? entTxt : ''}</div>
      </td>`;
    }).join('');

    return `<tr class="${isDup ? 'dup-row' : ''}">
      <td class="td-cod">
        <span class="cod-name">${cod}</span>
        ${isDup ? '<span class="badge b-dup">3+ sem.</span>' : ''}
      </td>
      <td class="td-tipo"><span class="badge ${tipoBadge(tipo)}">${tipo || '—'}</span></td>
      ${cells}
    </tr>`;
  }).join('');

  document.getElementById('pivot-container').innerHTML = `
    <div class="pivot-wrap">
      <table class="pivot-table">
        <thead><tr>
          <th class="th-left td-cod">Codigo vino</th>
          <th class="th-left td-tipo">Tipo</th>
          ${semanas.map(s => `<th class="th-sem ${s==='N+3'?'th-n3':''}">
            ${s === 'N+3' ? 'N+3 <span class="n3-tag-sm">+vis.</span>' : 'Sem. ' + s}
          </th>`).join('')}
        </tr></thead>
        <tbody>${rows}</tbody>
      </table>
      <div class="pivot-legend">
        <span class="leg-item"><span class="leg-dot ok-dot"></span>Dentro del rango</span>
        <span class="leg-item"><span class="leg-dot warn-dot"></span>Ocupacion &ge; 90%</span>
        <span class="leg-item"><span class="leg-dot over-dot"></span>Ocupacion &ge; 100%</span>
        <span class="leg-item" style="font-size:11px;color:#bbb">P = Plan &nbsp;·&nbsp; E = Entrega</span>
      </div>
    </div>`;
}

function fmtLts(v) {
  return v >= 1000 ? (Math.round(v / 100) / 10).toFixed(1) + 'k' : Math.round(v) + '';
}

function tipoBadge(tipo) {
  return tipo === 'MI' ? 'b-mi' : tipo === 'StComex' ? 'b-stc' : tipo === 'Estibas' ? 'b-est' : 'b-mi';
}

// --- ALERTAS ---
function computeAlerts(pivot) {
  const semanas = ['N','N+1','N+2','N+3'];
  const codigos = Object.keys(pivot);
  const dups = [], over = [], estibas = [];

  codigos.forEach(cod => {
    const present = semanas.filter(s => pivot[cod][s] && pivot[cod][s].planVol > 0);
    if (present.length >= 3) dups.push({ cod, sems: present.join(', ') });

    semanas.forEach(s => {
      const cell = pivot[cod][s];
      if (!cell) return;
      if (cell.tipo === 'Estibas') estibas.push({ cod, semana: s });
      if (cell.pct !== null) {
        const cfg = state.config.tipos[cell.tipo] || { maxPct:115, minPct:70 };
        if (cell.pct > cfg.maxPct)
          over.push({ cod, semana:s, pct:cell.pct, max:cfg.maxPct, tipo:cell.tipo });
      }
    });
  });

  return { dups, over, estibas };
}

function renderAlertas(pivot) {
  const a = computeAlerts(pivot);
  let html = '';

  if (a.dups.length) {
    html += `<div class="alert-section">
      <div class="alert-section-title">&#9888; Mismo codigo en 3+ semanas (${a.dups.length})</div>
      ${a.dups.map(d => `<div class="alert-card warn">
        <div class="alert-icon">&#128260;</div>
        <div><div class="alert-title">${d.cod}</div>
        <div class="alert-sub">Aparece en semanas: ${d.sems}</div></div>
      </div>`).join('')}
    </div>`;
  }

  if (a.over.length) {
    html += `<div class="alert-section">
      <div class="alert-section-title">&#128314; Sobre capacidad maxima (${a.over.length})</div>
      ${a.over.map(d => `<div class="alert-card danger">
        <div class="alert-icon">&#128202;</div>
        <div><div class="alert-title">${d.cod} &middot; Semana ${d.semana} &middot; ${d.tipo}</div>
        <div class="alert-sub">Ocupacion: ${d.pct}% &middot; Maximo configurado: ${d.max}%</div></div>
      </div>`).join('')}
    </div>`;
  }

  if (a.estibas.length) {
    html += `<div class="alert-section">
      <div class="alert-section-title">&#9940; Codigos en Estibas &mdash; evitar (${a.estibas.length})</div>
      ${a.estibas.map(e => `<div class="alert-card danger">
        <div class="alert-icon">&#128295;</div>
        <div><div class="alert-title">${e.cod} &middot; Semana ${e.semana}</div>
        <div class="alert-sub">Estibas es prioridad 3. Reasignar a MI o StComex si es posible.</div></div>
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

// --- CONFIG ---
function renderConfig() {
  document.getElementById('config-container').innerHTML = `
    <div class="config-section">
      <div class="config-title">Tolerancias por tipo</div>
      <div class="config-grid">
        ${['MI','StComex','Estibas'].map(t => {
          const cfg = state.config.tipos[t];
          return `<div class="config-card">
            <div class="config-card-title"><span class="badge ${tipoBadge(t)}">${t}</span></div>
            <div class="field-row">
              <label>% minimo de ocupacion</label>
              <input type="number" id="cfg-min-${t}" value="${cfg.minPct}">
              <span class="hint">Alerta si el codigo esta por debajo</span>
            </div>
            <div class="field-row">
              <label>% maximo de ocupacion</label>
              <input type="number" id="cfg-max-${t}" value="${cfg.maxPct}">
              <span class="hint">Alerta si el codigo supera este valor</span>
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
          <div class="alert-sub">Tolerancia: +15% / -10%</div></div></div>
        <div class="alert-card info"><div class="alert-icon">2</div>
          <div><div class="alert-title"><span class="badge b-stc">StComex</span> Prioridad 2</div>
          <div class="alert-sub">Tolerancia: +10% / -10%</div></div></div>
        <div class="alert-card danger"><div class="alert-icon">3</div>
          <div><div class="alert-title"><span class="badge b-est">Estibas</span> Prioridad 3 &mdash; evitar</div>
          <div class="alert-sub">Solo como ultimo recurso.</div></div></div>
      </div>
    </div>`;
}

function saveConfig() {
  ['MI','StComex','Estibas'].forEach(t => {
    state.config.tipos[t].minPct = parseFloat(document.getElementById('cfg-min-'+t).value) || 0;
    state.config.tipos[t].maxPct = parseFloat(document.getElementById('cfg-max-'+t).value) || 115;
  });
  const fb = document.getElementById('cfg-ok');
  fb.textContent = '✓ Guardado';
  setTimeout(() => fb.textContent = '', 2500);
  if (state.planLoaded || state.entregaLoaded) render();
}
