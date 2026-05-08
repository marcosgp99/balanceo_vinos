// ============================================================
// BALANCEO DE CARGAS WO v3
// ============================================================

const state = {
  planData:      [],   // { id, wo, codEq, semanaRaw, lts, estado, fecha }
  entregaData:   [],   // { id, codEq, semanaRaw, lts }
  semanaMap:     {},   // { "22": "N", "23": "N+1", ... }
  expanded:      {},   // { codEq: true/false }
  movedToN2:     {},   // { woId: true } — WOs movidas de N+3 a N+2
  config: {
    capacidadTotal: 500000,  // Lts totales de fraccionamiento
    minPct: 80,
    maxPct: 100
  },
  importMode:    'plan',
  rawHeaders:    [],
  rawData:       [],
  colMap:        {},
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

// --- EXCEL IMPORT ---
function handleFileUpload(e, mode) {
  const file = e.target.files[0];
  if (!file) return;
  state.importMode = mode;
  const reader = new FileReader();
  reader.onload = function(evt) {
    const wb   = XLSX.read(evt.target.result, { type: 'binary' });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const rows  = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    if (rows.length < 2) { alert('El archivo parece estar vacio.'); return; }
    state.rawHeaders = rows[0].map(h => String(h).trim());
    state.rawData    = rows.slice(1).filter(r => r.some(c => c !== ''));
    showColumnMapper(state.rawHeaders, autoDetect(state.rawHeaders, mode), mode);
  };
  reader.readAsBinaryString(file);
  e.target.value = '';
}

function autoDetect(headers, mode) {
  const map = { codEq:'', semana:'', lts:'', wo:'', estado:'', fecha:'' };
  headers.forEach(h => {
    const hl = h.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'');
    if (!map.codEq  && /codeq|cod.eq|equivalente/.test(hl))        map.codEq  = h;
    if (!map.semana && /\bsemana\b/.test(hl))                       map.semana = h;
    if (!map.lts    && /^lts$/.test(hl))                            map.lts    = h;
    if (!map.wo     && /^orden$/.test(hl))                          map.wo     = h;
    if (!map.estado && /^estado$/.test(hl))                         map.estado = h;
    if (!map.fecha  && /feinicio|fechainicio/.test(hl))             map.fecha  = h;
  });
  // Fallbacks
  if (!map.codEq)  map.codEq  = headers.find(h => /codeq/i.test(h))      || '';
  if (!map.semana) map.semana = headers.find(h => /semana/i.test(h))      || '';
  if (!map.lts)    map.lts    = headers.find(h => /^lts$/i.test(h))       || '';
  if (!map.wo)     map.wo     = headers.find(h => /orden/i.test(h))       || '';
  if (!map.estado) map.estado = headers.find(h => /estado/i.test(h))      || '';
  if (!map.fecha)  map.fecha  = headers.find(h => /feinicio|fecha/i.test(h)) || '';
  return map;
}

function showColumnMapper(headers, detected, mode) {
  const isPlan = mode === 'plan';
  document.getElementById('modal-title').textContent =
    isPlan ? 'Columnas del Plan' : 'Columnas de Entregas';
  document.getElementById('modal-desc').textContent =
    isPlan
      ? 'Mapeá las columnas del archivo de Plan:'
      : 'Mapeá las columnas del archivo de Entregas:';

  const fields = isPlan
    ? [
        { key:'codEq',  label:'Código equivalente (CodEq)', req:true  },
        { key:'semana', label:'Semana',                      req:true  },
        { key:'lts',    label:'Volumen (Lts)',                req:true  },
        { key:'wo',     label:'Número de WO (Orden)',        req:true  },
        { key:'estado', label:'Estado',                      req:false },
        { key:'fecha',  label:'Fecha de inicio (FeInicio)',  req:false }
      ]
    : [
        { key:'codEq',  label:'Código equivalente (CodEq)', req:true  },
        { key:'semana', label:'Semana',                      req:true  },
        { key:'lts',    label:'Volumen (Lts)',                req:true  }
      ];

  document.getElementById('col-map-fields').innerHTML = fields.map(f => `
    <div class="col-map-row">
      <label>${f.label}${f.req ? '' : ' <span style="color:#bbb;font-size:11px">(opcional)</span>'}</label>
      <select id="map-${f.key}">
        ${!f.req ? '<option value="">— no usar —</option>' : '<option value="">— seleccionar —</option>'}
        ${headers.map(h =>
          `<option value="${h}"${h === detected[f.key] ? ' selected' : ''}>${h}</option>`
        ).join('')}
      </select>
    </div>
  `).join('');

  state.colMap = detected;
  document.getElementById('modal-overlay').classList.remove('hidden');
}

function confirmImport() {
  const isPlan = state.importMode === 'plan';
  const required = isPlan ? ['codEq','semana','lts','wo'] : ['codEq','semana','lts'];
  const newMap = {};
  ['codEq','semana','lts','wo','estado','fecha'].forEach(f => {
    const el = document.getElementById('map-' + f);
    if (el) newMap[f] = el.value;
  });
  if (required.some(f => !newMap[f])) {
    alert('Asigná todas las columnas obligatorias.');
    return;
  }
  state.colMap = newMap;
  const processed = processData(state.rawData, newMap);

  if (isPlan) {
    state.planData   = processed;
    state.planLoaded = true;
    state.movedToN2  = {};   // reset al reimportar
    state.expanded   = {};
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
  return rows.map((r, i) => ({
    id:        state.nextId++,
    wo:        idx.wo     >= 0 ? String(r[idx.wo]     || '').trim() : 'WO-' + (i+1),
    codEq:     idx.codEq  >= 0 ? String(r[idx.codEq]  || '').trim() : '',
    semanaRaw: idx.semana >= 0 ? String(r[idx.semana] || '').trim() : '',
    lts:       idx.lts    >= 0 ? (parseFloat(r[idx.lts]) || 0) : 0,
    estado:    idx.estado >= 0 ? String(r[idx.estado] || '').trim() : '',
    fecha:     idx.fecha  >= 0 ? fmtFecha(String(r[idx.fecha] || '').trim()) : ''
  })).filter(r => r.codEq && r.lts > 0);
}

function fmtFecha(raw) {
  if (!raw) return '';
  // "25/5/2026 00:00" → "25/5/2026"
  return raw.split(' ')[0];
}

// --- WEEK NORMALIZATION ---
function buildSemanaMap() {
  const all = new Set([
    ...state.planData.map(r => r.semanaRaw),
    ...state.entregaData.map(r => r.semanaRaw)
  ]);
  const sorted = [...all].filter(Boolean).sort((a, b) => {
    const na = parseFloat(a), nb = parseFloat(b);
    return (!isNaN(na) && !isNaN(nb)) ? na - nb : a.localeCompare(b);
  });
  const labels = ['N','N+1','N+2','N+3'];
  state.semanaMap = {};
  sorted.forEach((w, i) => { state.semanaMap[w] = labels[i] || 'N+3'; });

  const txt = sorted.map((w, i) => `${labels[i] || 'N+3'}=sem.${w}`).join(' · ');
  document.getElementById('status-semanas').textContent = txt;
}

function updateStatus(type, text) {
  const id = type === 'plan' ? 'status-plan' : 'status-entrega';
  const el = document.getElementById(id);
  el.querySelector('.status-dot').className = 'status-dot dot-on';
  el.querySelector('span:last-child').textContent =
    (type === 'plan' ? 'Plan: ' : 'Entregas: ') + text;
}

// --- SEMANA EFECTIVA (considera WOs movidas) ---
function getSemana(row) {
  if (state.movedToN2[row.id]) {
    return 'N+2';
  }
  return state.semanaMap[row.semanaRaw] || null;
}

// --- MOVER WO DE N+3 A N+2 ---
function moveToN2(woId) {
  state.movedToN2[woId] = true;
  render();
}

// --- DEMO DATA ---
function cargarDemoData() {
  state.planData = [
    { id:1,  wo:'1129121', codEq:'C1079',       semanaRaw:'22', lts:7141.5,  estado:'40', fecha:'25/5/2026' },
    { id:2,  wo:'1134019', codEq:'C1079',       semanaRaw:'22', lts:7012.1,  estado:'40', fecha:'25/5/2026' },
    { id:3,  wo:'1134020', codEq:'A1039/A1497', semanaRaw:'22', lts:5200.0,  estado:'40', fecha:'25/5/2026' },
    { id:4,  wo:'1134021', codEq:'A1039/A1497', semanaRaw:'23', lts:5100.0,  estado:'40', fecha:'1/6/2026'  },
    { id:5,  wo:'1134022', codEq:'C1079',       semanaRaw:'23', lts:8500.0,  estado:'40', fecha:'1/6/2026'  },
    { id:6,  wo:'1134023', codEq:'B2201',       semanaRaw:'23', lts:3800.0,  estado:'40', fecha:'1/6/2026'  },
    { id:7,  wo:'1134024', codEq:'A1039/A1497', semanaRaw:'24', lts:5050.0,  estado:'20', fecha:'8/6/2026'  },
    { id:8,  wo:'1134025', codEq:'B2201',       semanaRaw:'24', lts:3600.0,  estado:'20', fecha:'8/6/2026'  },
    { id:9,  wo:'1134026', codEq:'C1079',       semanaRaw:'25', lts:7200.0,  estado:'20', fecha:'15/6/2026' },
    { id:10, wo:'1134027', codEq:'D3310',       semanaRaw:'25', lts:4100.0,  estado:'20', fecha:'15/6/2026' },
    { id:11, wo:'1134028', codEq:'A1039/A1497', semanaRaw:'25', lts:4900.0,  estado:'20', fecha:'15/6/2026' },
  ];
  state.entregaData = [
    { id:20, codEq:'C1079',       semanaRaw:'22', lts:16000.0 },
    { id:21, codEq:'A1039/A1497', semanaRaw:'22', lts:4800.0  },
    { id:22, codEq:'A1039/A1497', semanaRaw:'23', lts:5300.0  },
    { id:23, codEq:'C1079',       semanaRaw:'23', lts:7200.0  },
    { id:24, codEq:'B2201',       semanaRaw:'23', lts:3600.0  },
    { id:25, codEq:'A1039/A1497', semanaRaw:'24', lts:9500.0  },  // surplus para demo
    { id:26, codEq:'B2201',       semanaRaw:'24', lts:3400.0  },
    { id:27, codEq:'C1079',       semanaRaw:'25', lts:7500.0  },
    { id:28, codEq:'D3310',       semanaRaw:'25', lts:3200.0  },
  ];
  state.planLoaded    = true;
  state.entregaLoaded = true;
  state.movedToN2     = {};
  state.expanded      = {};
  state.nextId        = 50;
  updateStatus('plan',     '11 filas cargadas (demo)');
  updateStatus('entregas',  '9 filas cargadas (demo)');
  buildSemanaMap();
  render();
}

// --- PIVOT ---
function buildPivot() {
  const pivot   = {};   // pivot[codEq][semana] = { planVol, entregaVol, wos:[] }
  const semanas = ['N','N+1','N+2','N+3'];

  // Solo CodEq que tengan WOs en planData
  state.planData.forEach(r => {
    const sem = getSemana(r);
    if (!sem) return;
    if (!pivot[r.codEq]) pivot[r.codEq] = {};
    if (!pivot[r.codEq][sem]) pivot[r.codEq][sem] = { planVol:0, entregaVol:0, wos:[] };
    pivot[r.codEq][sem].planVol += r.lts;
    pivot[r.codEq][sem].wos.push({ ...r, semLabel: sem });
  });

  // Entregas: solo para CodEq que ya existan en pivot (tienen WOs)
  state.entregaData.forEach(r => {
    const sem = state.semanaMap[r.semanaRaw];
    if (!sem || !pivot[r.codEq]) return;
    if (!pivot[r.codEq][sem]) pivot[r.codEq][sem] = { planVol:0, entregaVol:0, wos:[] };
    pivot[r.codEq][sem].entregaVol += r.lts;
  });

  // Calcular %: planVol / entregaVol * 100
  Object.values(pivot).forEach(codData => {
    Object.values(codData).forEach(cell => {
      if (cell.entregaVol > 0) {
        cell.pct = Math.round(cell.planVol / cell.entregaVol * 100);
      } else if (cell.planVol > 0) {
        cell.pct = null;  // hay WOs pero sin entrega registrada
      } else {
        cell.pct = null;
      }
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
  const totalPlanLts = state.planData.reduce((a, r) => a + r.lts, 0);
  const utilPct      = state.config.capacidadTotal > 0
    ? Math.round(totalPlanLts / state.config.capacidadTotal * 100)
    : 0;
  const utilCls = utilPct >= 95 ? 'danger' : utilPct >= 75 ? 'ok' : 'warn';

  const alerts = computeAlerts(pivot);
  const na     = alerts.dups.length + alerts.deficit.length + alerts.surplus.length;
  const n3cnt  = Object.keys(pivot).filter(c => pivot[c]['N+3']).length;

  document.getElementById('m-util').textContent   = state.planLoaded ? utilPct + '%' : '—';
  document.getElementById('m-util').className     = 'metric-value ' + (state.planLoaded ? utilCls : '');
  document.getElementById('m-lts').textContent    = state.planLoaded ? fmtLts(totalPlanLts) : '—';
  document.getElementById('m-alerts').textContent = (state.planLoaded || state.entregaLoaded) ? na : '—';
  document.getElementById('m-alerts').className   = 'metric-value ' + (na > 0 ? 'warn' : 'ok');
  document.getElementById('m-n3').textContent     = n3cnt || '—';
  document.getElementById('alert-badge').textContent = na;
}

// --- PIVOT TABLE ---
function renderPivotTable(pivot) {
  const semanas = ['N','N+1','N+2','N+3'];
  const codigos = Object.keys(pivot).sort();

  if (!codigos.length) {
    const msg = !state.planLoaded
      ? 'Importa el Plan usando el boton "Importar Plan".'
      : 'No hay codigos con WOs para mostrar.';
    document.getElementById('pivot-container').innerHTML =
      `<div class="empty-state"><div class="empty-icon">&#128202;</div>
       <div class="empty-title">Sin datos</div><div class="empty-sub">${msg}</div></div>`;
    return;
  }

  const alerts  = computeAlerts(pivot);
  const dupCods = new Set(alerts.dups.map(d => d.cod));
  const surplusCods = new Set(alerts.surplus.map(s => s.cod));

  let rows = codigos.map(cod => {
    const isExp     = !!state.expanded[cod];
    const isDup     = dupCods.has(cod);
    const hasSurplus = surplusCods.has(cod);
    const hasCells  = semanas.some(s => pivot[cod][s]);

    if (!hasCells) return '';

    const cells = semanas.map(s => {
      const cell = pivot[cod][s];
      if (!cell) return `<td class="pcell-empty pcell${s==='N+3'?' pcell-n3':''}">—</td>`;

      const p    = cell.pct;
      let cls    = 'pcell' + (s === 'N+3' ? ' pcell-n3' : '');
      let pctTxt = '—';

      if (p !== null && state.entregaLoaded) {
        pctTxt = p + '%';
        if (p > state.config.maxPct) cls += ' cell-warn';      // deficit
        else if (p < state.config.minPct) cls += ' cell-surplus'; // surplus grande
        else cls += ' cell-ok';
      } else if (p === null && cell.planVol > 0) {
        cls += ' cell-nodata';
      } else {
        cls += ' cell-ok';
      }

      const planTxt = fmtLts(cell.planVol);
      const entTxt  = state.entregaLoaded && cell.entregaVol > 0 ? fmtLts(cell.entregaVol) : '—';

      return `<td class="${cls}">
        <div class="cell-pct">${pctTxt}</div>
        <div class="cell-vols">WO: ${planTxt}${state.entregaLoaded ? ' / Ent: ' + entTxt : ''}</div>
      </td>`;
    }).join('');

    // Sub-rows con WOs individuales
    const allWOs = semanas.flatMap(s => (pivot[cod][s] ? pivot[cod][s].wos : []));
    const surplusN2 = hasSurplus; // si hay surplus en N+2 para este cod

    const subRows = allWOs.map(w => {
      const isN3   = w.semLabel === 'N+3';
      const moved  = state.movedToN2[w.id];
      const showMoveBtn = isN3 && surplusN2 && !moved;
      return `<tr>
        <td class="wo-num">${w.wo}</td>
        <td><span class="sem-badge${isN3 ? ' sem-n3' : ''}">${w.semLabel}</span></td>
        <td>${fmtLts(w.lts)} Lts</td>
        <td>${w.estado || '—'}</td>
        <td>${w.fecha  || '—'}</td>
        <td>
          ${moved ? '<span class="moved-tag">&#10003; Movida a N+2</span>' : ''}
          ${showMoveBtn ? `<button class="btn-mover" onclick="moveToN2(${w.id})">&#8593; Mover a N+2</button>` : ''}
        </td>
      </tr>`;
    }).join('');

    const subSection = isExp ? `
      <tr class="sub-row" id="sub-${codSafe(cod)}">
        <td colspan="5">
          <table class="sub-table">
            <thead><tr>
              <th>WO</th><th>Semana</th><th>Lts</th><th>Estado</th><th>Fecha inicio</th><th></th>
            </tr></thead>
            <tbody>${subRows}</tbody>
          </table>
        </td>
      </tr>` : '';

    return `
      <tr class="cod-row" onclick="toggleExpand('${codSafe(cod)}')">
        <td colspan="1">
          <div class="cod-cell">
            <span class="expand-icon${isExp ? ' open' : ''}">&#9658;</span>
            <span class="cod-name">${cod}</span>
            ${isDup ? '<span class="badge b-dup">3+ sem.</span>' : ''}
            ${hasSurplus ? '<span class="badge b-surplus">Sobrante N+2</span>' : ''}
          </div>
        </td>
        ${cells}
      </tr>
      ${subSection}`;
  }).join('');

  document.getElementById('pivot-container').innerHTML = `
    <div class="pivot-wrap">
      <table class="pivot-table">
        <thead><tr>
          <th class="th-left" style="min-width:180px">Codigo vino</th>
          ${semanas.map(s =>
            `<th class="th-sem${s==='N+3'?' th-n3':''}">
              ${s === 'N+3' ? 'N+3 <span class="n3-tag-sm">+vis.</span>' : 'Sem. ' + s}
            </th>`
          ).join('')}
        </tr></thead>
        <tbody>${rows}</tbody>
      </table>
      <div class="pivot-legend">
        <span class="leg-item"><span class="leg-dot ok-dot"></span>En rango (${state.config.minPct}–${state.config.maxPct}%)</span>
        <span class="leg-item"><span class="leg-dot warn-dot"></span>Deficit (&gt;${state.config.maxPct}%)</span>
        <span class="leg-item"><span class="leg-dot surplus-dot"></span>Sobrante (&lt;${state.config.minPct}%)</span>
        <span class="leg-item" style="font-size:11px;color:#bbb">% = WO Lts / Entrega Lts &nbsp;·&nbsp; Clic en codigo para ver WOs</span>
      </div>
    </div>`;
}

function codSafe(cod) {
  return cod.replace(/[^a-zA-Z0-9]/g, '_');
}

function toggleExpand(codSafeVal) {
  // Reconstruir cod original buscando en pivot
  const pivot   = buildPivot();
  const original = Object.keys(pivot).find(c => codSafe(c) === codSafeVal);
  if (!original) return;
  state.expanded[original] = !state.expanded[original];
  renderPivotTable(pivot);
  renderAlertas(pivot);
}

// --- ALERTAS ---
function computeAlerts(pivot) {
  const semanas = ['N','N+1','N+2','N+3'];
  const codigos = Object.keys(pivot);
  const dups = [], deficit = [], surplus = [];

  codigos.forEach(cod => {
    // 1. Mismo código 3+ semanas
    const present = semanas.filter(s => pivot[cod][s] && pivot[cod][s].planVol > 0);
    if (present.length >= 3) dups.push({ cod, sems: present.join(', ') });

    // 2. Deficit o surplus
    semanas.forEach(s => {
      const cell = pivot[cod][s];
      if (!cell || cell.pct === null) return;
      if (cell.pct > state.config.maxPct)
        deficit.push({ cod, semana:s, pct:cell.pct });
    });

    // 3. Surplus en N+2 → sugerir WOs de N+3
    const n2 = pivot[cod]['N+2'];
    if (n2 && n2.entregaVol > 0 && n2.planVol < n2.entregaVol) {
      const sobrante = n2.entregaVol - n2.planVol;
      const wosN3 = (pivot[cod]['N+3'] || { wos:[] }).wos
        .filter(w => !state.movedToN2[w.id]);
      if (wosN3.length > 0) {
        surplus.push({ cod, sobrante, wosN3 });
      }
    }
  });

  return { dups, deficit, surplus };
}

function renderAlertas(pivot) {
  const a = computeAlerts(pivot);
  let html = '';

  // Surplus N+2 con sugerencia de adelantar N+3
  if (a.surplus.length) {
    html += `<div class="alert-section">
      <div class="alert-section-title">&#128200; Sobrante en N+2 — sugerencia de adelantar (${a.surplus.length})</div>
      ${a.surplus.map(s => `
        <div class="alert-card surplus">
          <div class="alert-icon">&#128199;</div>
          <div style="flex:1">
            <div class="alert-title">${s.cod} — Sobrante en N+2: ${fmtLts(s.sobrante)} Lts</div>
            <div class="alert-sub">Las siguientes WOs de N+3 pueden adelantarse para aprovechar el sobrante:</div>
            <div class="alert-wo-list">
              ${s.wosN3.map(w => `
                <div class="alert-wo-row">
                  <span><strong>${w.wo}</strong> &middot; ${fmtLts(w.lts)} Lts &middot; ${w.fecha || ''}</span>
                  <button class="btn-mover" onclick="moveToN2(${w.id})">&#8593; Mover a N+2</button>
                </div>
              `).join('')}
            </div>
          </div>
        </div>
      `).join('')}
    </div>`;
  }

  // Duplicados 3+ semanas
  if (a.dups.length) {
    html += `<div class="alert-section">
      <div class="alert-section-title">&#9888; Mismo codigo en 3+ semanas (${a.dups.length})</div>
      ${a.dups.map(d => `
        <div class="alert-card warn">
          <div class="alert-icon">&#128260;</div>
          <div>
            <div class="alert-title">${d.cod}</div>
            <div class="alert-sub">Aparece en semanas: ${d.sems}</div>
          </div>
        </div>
      `).join('')}
    </div>`;
  }

  // Deficit
  if (a.deficit.length) {
    html += `<div class="alert-section">
      <div class="alert-section-title">&#128314; Deficit de vino (WOs &gt; Entrega) (${a.deficit.length})</div>
      ${a.deficit.map(d => `
        <div class="alert-card danger">
          <div class="alert-icon">&#128202;</div>
          <div>
            <div class="alert-title">${d.cod} &middot; Semana ${d.semana}</div>
            <div class="alert-sub">WOs representan el ${d.pct}% de la entrega (maximo: ${state.config.maxPct}%)</div>
          </div>
        </div>
      `).join('')}
    </div>`;
  }

  if (!html) html = `<div class="alert-card ok">
    <div class="alert-icon">&#9989;</div>
    <div>
      <div class="alert-title">Sin alertas activas</div>
      <div class="alert-sub">El plan esta correctamente balanceado.</div>
    </div>
  </div>`;

  document.getElementById('alertas-container').innerHTML = html;
}

// --- CONFIG ---
function renderConfig() {
  document.getElementById('config-container').innerHTML = `
    <div class="config-section">
      <div class="config-title">Capacidad de fraccionamiento</div>
      <div class="config-grid">
        <div class="config-card">
          <div class="config-card-title">Capacidad total</div>
          <div class="field-row">
            <label>Litros totales de fraccionamiento</label>
            <input type="number" id="cfg-cap" value="${state.config.capacidadTotal}" placeholder="ej: 500000">
            <span class="hint">Se usa para calcular el % de utilizacion en el dashboard</span>
          </div>
        </div>
        <div class="config-card">
          <div class="config-card-title">Rangos de alerta (%)</div>
          <div class="field-row">
            <label>% minimo (debajo = sobrante)</label>
            <input type="number" id="cfg-min" value="${state.config.minPct}" placeholder="ej: 80">
            <span class="hint">Celda azul si WO Lts / Entrega Lts esta por debajo</span>
          </div>
          <div class="field-row">
            <label>% maximo (encima = deficit)</label>
            <input type="number" id="cfg-max" value="${state.config.maxPct}" placeholder="ej: 100">
            <span class="hint">Celda roja si WO Lts / Entrega Lts supera este valor</span>
          </div>
        </div>
      </div>
      <div style="display:flex;align-items:center;margin-top:16px">
        <button class="btn-primary" onclick="saveConfig()">Guardar configuracion</button>
        <span id="cfg-ok" class="save-feedback"></span>
      </div>
    </div>`;
}

function saveConfig() {
  state.config.capacidadTotal = parseFloat(document.getElementById('cfg-cap').value) || state.config.capacidadTotal;
  state.config.minPct         = parseFloat(document.getElementById('cfg-min').value) || state.config.minPct;
  state.config.maxPct         = parseFloat(document.getElementById('cfg-max').value) || state.config.maxPct;
  document.getElementById('cfg-ok').textContent = '&#10003; Guardado';
  setTimeout(() => document.getElementById('cfg-ok').textContent = '', 2500);
  if (state.planLoaded || state.entregaLoaded) render();
}

// --- UTILS ---
function fmtLts(v) {
  if (!v) return '0';
  if (v >= 1000000) return (v / 1000000).toFixed(1) + 'M';
  if (v >= 1000)    return (Math.round(v / 100) / 10).toFixed(1) + 'k';
  return Math.round(v).toString();
}
