/* ── Auditoría Equipo de Ventas · app.js ── */

// ── Criteria definitions ───────────────────────────────────────────────────
const CRITERIA = [
  { key: 'c1', col: 'C1_Nombre',        label: '¿Usó el nombre del paciente desde el primer mensaje?' },
  { key: 'c2', col: 'C2_SPIN',          label: '¿Hizo al menos 2 preguntas abiertas (SPIN) antes de dar información?' },
  { key: 'c3', col: 'C3_Valor',         label: '¿Presentó el valor antes del precio (Trío del Valor)?' },
  { key: 'c4', col: 'C4_SiguientePaso', label: '¿La conversación terminó con un siguiente paso concreto y fechado?' },
  { key: 'c5', col: 'C5_Respuesta2min', label: '¿Respondió dentro de los 2 minutos establecidos?' },
  { key: 'c6', col: 'C6_HubSpot',       label: '¿Documentó correctamente en HubSpot?' },
  { key: 'c7', col: 'C7_Objeciones',    label: '¿Aplicó el protocolo de objeciones (Reconocer→Validar→Informar→Invitar)?' },
  { key: 'c8', col: 'C8_InfoMedica',    label: '¿Evitó dar información médica sin confirmación del equipo clínico?' },
  { key: 'c9', col: 'C9_Emocion',       label: 'Si el paciente mostró emoción, ¿la reconoció antes de pasar a información?', optional: true },
  { key: 'c10', col: 'C10_EVRAG',       label: '¿Aplicó el protocolo EVRA+G en caso de frustración o molestia?', optional: true },
];

const INCIDENTS = [
  { key: 'i1', col: 'I1_PrecioSinValor',      label: 'Presentó el precio sin presentar valor primero' },
  { key: 'i2', col: 'I2_Lead24h',             label: 'Lead con interés real lleva más de 24 hs sin respuesta' },
  { key: 'i3', col: 'I3_PlantillaGenerica',   label: 'Se usó la plantilla genérica completa sin personalización' },
  { key: 'i4', col: 'I4_FrustracionNoValidada', label: 'El paciente expresó frustración y no fue validado emocionalmente' },
  { key: 'i5', col: 'I5_AltoValorSinSupervisora', label: 'Se cotizó FIV/PGT u otro tratamiento de alto valor sin involucrar a la supervisora' },
  { key: 'i6', col: 'I6_DealPerdidoPronto',   label: 'Cerró el deal como perdido antes de los 2 intentos establecidos' },
  { key: 'i7', col: 'I7_InfoClinicaSinMedico', label: 'Se dio información clínica sin confirmación del equipo médico' },
  { key: 'i8', col: 'I8_NoDocumentoHubSpot',  label: 'No documentó la conversación en HubSpot al cierre del día' },
];

// ── State ──────────────────────────────────────────────────────────────────
let allRows = [];
let filtered = [];
let charts = {};
let activeTab = 'cumplimiento';

// ── File handling ──────────────────────────────────────────────────────────
function handleFile(evt) {
  const file = evt.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb = XLSX.read(e.target.result, { type: 'array', cellDates: true });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(ws, { defval: '' });
      processData(json, file.name);
    } catch (err) {
      showToast('Error al leer el archivo: ' + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
  evt.target.value = '';
}

function loadSampleData() {
  processData(generateSampleData(), 'datos_ejemplo');
}

// ── Process data ───────────────────────────────────────────────────────────
function processData(rows, filename) {
  if (!rows.length) { showToast('El archivo está vacío.'); return; }

  const headers = Object.keys(rows[0]).map(h => h.trim());

  // Find columns flexibly
  const findCol = (colName, aliases) => {
    // exact match first
    const h = headers.find(h => h.toLowerCase() === colName.toLowerCase());
    if (h) return h;
    // alias match
    if (aliases) {
      const a = headers.find(h => aliases.some(al => h.toLowerCase().includes(al.toLowerCase())));
      if (a) return a;
    }
    return null;
  };

  const dateCol  = findCol('Fecha', ['fecha', 'date', 'día']);
  const agentCol = findCol('Agente', ['agente', 'agent', 'asesor', 'vendedor', 'nombre']);
  const chatCol  = findCol('Chat_ID', ['chat', 'id', 'paciente', 'folio']);

  if (!agentCol) { showToast('No se encontró columna "Agente". Revisa la plantilla.'); return; }

  allRows = rows.map((r, i) => {
    const dateRaw = dateCol ? r[dateCol] : null;
    const dateVal = parseDate(dateRaw) || `Fila ${i+2}`;
    const agent   = String(r[agentCol] || 'Sin nombre').trim();
    const chatId  = chatCol ? String(r[chatCol] || `Chat ${i+1}`).trim() : `Chat ${i+1}`;
    const week    = dateToWeek(dateVal);

    const criteria = {};
    CRITERIA.forEach(c => {
      const col = findCol(c.col, [c.key]);
      const val = col ? String(r[col] || '').trim().toLowerCase() : '';
      if (val === 'n/a' || val === 'na' || val === '') {
        criteria[c.key] = c.optional ? null : false;
      } else {
        criteria[c.key] = ['sí','si','yes','1','true','✓','x'].includes(val);
      }
    });

    const incidents = {};
    INCIDENTS.forEach(inc => {
      const col = findCol(inc.col, [inc.key]);
      const val = col ? String(r[col] || '').trim().toLowerCase() : '';
      incidents[inc.key] = ['sí','si','yes','1','true','✓','x'].includes(val);
    });

    // compliance = % of applicable criteria that passed
    const applicable = CRITERIA.filter(c => criteria[c.key] !== null);
    const passed     = applicable.filter(c => criteria[c.key] === true);
    const compliance = applicable.length ? Math.round(passed.length / applicable.length * 100) : 0;

    return { date: dateVal, week, agent, chatId, criteria, incidents, compliance };
  });

  showToast(`✓ ${allRows.length} auditorías cargadas`);
  buildFilterOptions();
  applyFilters();

  document.getElementById('uploadZone').style.display = 'none';
  document.getElementById('dashboard').style.display = 'block';
  document.getElementById('tabBar').style.display = 'block';
}

function parseDate(v) {
  if (!v) return null;
  if (v instanceof Date) return v.toISOString().slice(0,10);
  const d = new Date(v);
  if (!isNaN(d)) return d.toISOString().slice(0,10);
  return String(v);
}

function dateToWeek(dateStr) {
  const d = new Date(dateStr);
  if (isNaN(d)) return dateStr;
  const jan4 = new Date(d.getFullYear(), 0, 4);
  const week = Math.ceil(((d - jan4) / 86400000 + jan4.getDay() + 1) / 7);
  return `Semana ${week} · ${d.getFullYear()}`;
}

function weekSortKey(weekLabel) {
  // "Semana 23 · 2025" → "2025-023"
  const m = weekLabel.match(/Semana (\d+) · (\d+)/);
  if (!m) return weekLabel;
  return `${m[2]}-${m[1].padStart(3,'0')}`;
}

// ── Filters ────────────────────────────────────────────────────────────────
function buildFilterOptions() {
  const weeks  = [...new Set(allRows.map(r => r.week))].sort((a,b) => weekSortKey(a) > weekSortKey(b) ? 1 : -1);
  const agents = [...new Set(allRows.map(r => r.agent))].sort();

  const wSel = document.getElementById('filterWeek');
  wSel.innerHTML = '<option value="">Todas las semanas</option>';
  weeks.forEach(w => { const o = document.createElement('option'); o.value = o.textContent = w; wSel.appendChild(o); });

  // default to latest week
  if (weeks.length) wSel.value = weeks[weeks.length - 1];

  const aSel = document.getElementById('filterAgent');
  aSel.innerHTML = '<option value="">Todos los agentes</option>';
  agents.forEach(a => { const o = document.createElement('option'); o.value = o.textContent = a; aSel.appendChild(o); });
}

function applyFilters() {
  const week  = document.getElementById('filterWeek').value;
  const agent = document.getElementById('filterAgent').value;
  filtered = allRows.filter(r =>
    (!week  || r.week  === week) &&
    (!agent || r.agent === agent)
  );
  document.getElementById('recordCount').textContent =
    `${filtered.length} chat${filtered.length !== 1 ? 's' : ''} auditado${filtered.length !== 1 ? 's' : ''}`;
  renderAll();
}

function clearFilters() {
  document.getElementById('filterWeek').value  = '';
  document.getElementById('filterAgent').value = '';
  applyFilters();
}

// ── Tab switching ──────────────────────────────────────────────────────────
function switchTab(tab) {
  activeTab = tab;
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  document.getElementById('tab-' + tab).classList.add('active');
  document.getElementById('panelCumplimiento').style.display = tab === 'cumplimiento' ? '' : 'none';
  document.getElementById('panelIncidencias').style.display  = tab === 'incidencias'  ? '' : 'none';
}

// ── Render all ─────────────────────────────────────────────────────────────
function renderAll() {
  renderKPIs();
  renderAgentCards();
  renderCriteriaChart();
  renderTrendChart();
  renderIncidentSummary();
  renderAgentIncidents();
}

// ── KPIs ───────────────────────────────────────────────────────────────────
function renderKPIs() {
  const agents  = [...new Set(filtered.map(r => r.agent))].length;
  const total   = filtered.length;
  const avgComp = total ? Math.round(filtered.reduce((s,r) => s + r.compliance, 0) / total) : 0;
  const fullPass = filtered.filter(r => r.compliance === 100).length;
  const incCount = filtered.reduce((s,r) => s + INCIDENTS.filter(i => r.incidents[i.key]).length, 0);

  const cls = avgComp >= 80 ? 'success' : avgComp >= 60 ? 'warning' : 'danger';

  document.getElementById('kpiGrid').innerHTML = `
    <div class="kpi-card ${cls}">
      <div class="kpi-label">Cumplimiento promedio</div>
      <div class="kpi-value">${avgComp}<span style="font-size:16px;font-weight:600">%</span></div>
      <div class="kpi-sub">De todos los criterios</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">Chats auditados</div>
      <div class="kpi-value">${total}</div>
      <div class="kpi-sub">${agents} agente${agents !== 1 ? 's' : ''}</div>
    </div>
    <div class="kpi-card success">
      <div class="kpi-label">Chats 100% conformes</div>
      <div class="kpi-value">${fullPass}</div>
      <div class="kpi-sub">${total ? Math.round(fullPass/total*100) : 0}% del total</div>
    </div>
    <div class="kpi-card ${incCount > 0 ? 'danger' : 'success'}">
      <div class="kpi-label">Total incidencias</div>
      <div class="kpi-value">${incCount}</div>
      <div class="kpi-sub">En el período</div>
    </div>
  `;
}

// ── Agent cards ────────────────────────────────────────────────────────────
function renderAgentCards() {
  const agents = [...new Set(filtered.map(r => r.agent))].sort();
  const grid = document.getElementById('agentGrid');

  grid.innerHTML = agents.map(agent => {
    const rows = filtered.filter(r => r.agent === agent);
    const avg  = Math.round(rows.reduce((s,r) => s + r.compliance, 0) / rows.length);
    const pass = rows.filter(r => r.compliance === 100).length;
    const fail = rows.length - pass;
    const cls  = avg >= 80 ? 'success' : avg >= 60 ? 'warning' : 'danger';

    // top missed criteria
    const missed = CRITERIA.map(c => {
      const applicable = rows.filter(r => r.criteria[c.key] !== null);
      const failed     = applicable.filter(r => r.criteria[c.key] === false);
      return { label: c.label, pct: applicable.length ? Math.round(failed.length / applicable.length * 100) : 0 };
    }).filter(c => c.pct > 0).sort((a,b) => b.pct - a.pct).slice(0, 3);

    const missedHtml = missed.length
      ? missed.map(m => `<div class="missed-item">${truncate(m.label, 52)} <span>${m.pct}%</span></div>`).join('')
      : `<div style="font-size:12px;color:var(--success)">✓ Sin criterios fallidos</div>`;

    return `
      <div class="agent-card">
        <div class="agent-card-header">
          <div>
            <div class="agent-name">${agent}</div>
            <div class="agent-week">${rows.length} chat${rows.length !== 1 ? 's' : ''} auditado${rows.length !== 1 ? 's' : ''}</div>
          </div>
          <div class="compliance-big ${cls}">${avg}%</div>
        </div>
        <div class="compliance-bar-wrap">
          <div class="compliance-bar-label"><span>Cumplimiento</span><span>${pass}/${rows.length} conformes</span></div>
          <div class="compliance-bar-track">
            <div class="compliance-bar-fill ${cls}" style="width:${avg}%"></div>
          </div>
        </div>
        <div class="agent-stats">
          <div class="agent-stat">
            <div class="agent-stat-val" style="color:var(--success)">${pass}</div>
            <div class="agent-stat-lbl">Conformes</div>
          </div>
          <div class="agent-stat">
            <div class="agent-stat-val" style="color:var(--danger)">${fail}</div>
            <div class="agent-stat-lbl">Con fallas</div>
          </div>
          <div class="agent-stat">
            <div class="agent-stat-val" style="color:var(--warning)">${rows.reduce((s,r) => s + INCIDENTS.filter(i => r.incidents[i.key]).length, 0)}</div>
            <div class="agent-stat-lbl">Incidencias</div>
          </div>
        </div>
        <div class="missed-criteria">
          <div class="missed-label">Criterios más fallidos</div>
          ${missedHtml}
        </div>
      </div>`;
  }).join('');
}

// ── Criteria chart (horizontal bars) ──────────────────────────────────────
function renderCriteriaChart() {
  destroyChart('chartCriteria');
  if (!filtered.length) return;

  const data = CRITERIA.map(c => {
    const applicable = filtered.filter(r => r.criteria[c.key] !== null);
    const failed     = applicable.filter(r => r.criteria[c.key] === false);
    return applicable.length ? Math.round(failed.length / applicable.length * 100) : 0;
  });

  const shortLabels = [
    'Usó nombre','Preguntas SPIN','Valor antes precio',
    'Siguiente paso','Respuesta 2 min','Documentó HubSpot',
    'Protocolo objeciones','Info médica','Reconoció emoción','EVRA+G',
  ];

  charts['chartCriteria'] = new Chart(document.getElementById('chartCriteria'), {
    type: 'bar',
    data: {
      labels: shortLabels,
      datasets: [{
        label: '% chats con falla',
        data,
        backgroundColor: data.map(v => v >= 40 ? '#dc262688' : v >= 20 ? '#d9770688' : '#16a34a44'),
        borderColor:     data.map(v => v >= 40 ? '#dc2626'   : v >= 20 ? '#d97706'   : '#16a34a'),
        borderWidth: 2, borderRadius: 5,
      }],
    },
    options: {
      indexAxis: 'y',
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        x: { min: 0, max: 100, ticks: { callback: v => v + '%', font: { size: 11 } }, grid: { color: '#e2e8f0' } },
        y: { ticks: { font: { size: 11 } }, grid: { display: false } },
      },
    },
  });
}

// ── Trend chart ────────────────────────────────────────────────────────────
function renderTrendChart() {
  destroyChart('chartTrend');
  const weeks  = [...new Set(allRows.map(r => r.week))].sort((a,b) => weekSortKey(a) > weekSortKey(b) ? 1 : -1);
  const agents = [...new Set(filtered.map(r => r.agent))].sort();
  const PALETTE = ['#2563eb','#7c3aed','#db2777','#d97706','#16a34a','#0891b2'];

  const datasets = agents.slice(0,6).map((agent, i) => ({
    label: agent,
    data: weeks.map(w => {
      const rows = allRows.filter(r => r.agent === agent && r.week === w);
      return rows.length ? Math.round(rows.reduce((s,r) => s + r.compliance, 0) / rows.length) : null;
    }),
    borderColor: PALETTE[i % PALETTE.length],
    backgroundColor: PALETTE[i % PALETTE.length] + '22',
    tension: .3, fill: false, pointRadius: 4, spanGaps: true,
  }));

  charts['chartTrend'] = new Chart(document.getElementById('chartTrend'), {
    type: 'line',
    data: { labels: weeks.map(w => w.replace(' · ', '\n')), datasets },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { position: 'top', labels: { font: { size: 11 }, boxWidth: 12 } } },
      scales: {
        y: { min: 0, max: 100, ticks: { callback: v => v + '%', font: { size: 11 } }, grid: { color: '#e2e8f0' } },
        x: { ticks: { font: { size: 10 } }, grid: { display: false } },
      },
    },
  });
}

// ── Incident summary cards ─────────────────────────────────────────────────
function renderIncidentSummary() {
  document.getElementById('incidentSummary').innerHTML = INCIDENTS.map(inc => {
    const count   = filtered.filter(r => r.incidents[inc.key]).length;
    const agents  = [...new Set(filtered.filter(r => r.incidents[inc.key]).map(r => r.agent))];
    const agentTxt = agents.length ? agents.join(', ') : '—';
    return `
      <div class="incident-card ${count === 0 ? 'zero' : ''}">
        <div class="incident-count">${count}</div>
        <div class="incident-info">
          <div class="incident-title">${inc.label}</div>
          <div class="incident-agents">${count ? 'Agente(s): ' + agentTxt : 'Sin incidencias'}</div>
        </div>
      </div>`;
  }).join('');
}

// ── Per-agent incident detail ──────────────────────────────────────────────
function renderAgentIncidents() {
  const agents = [...new Set(filtered.map(r => r.agent))].sort();
  document.getElementById('agentIncidents').innerHTML = agents.map(agent => {
    const rows = filtered.filter(r => r.agent === agent);
    const agentIncs = INCIDENTS.map(inc => {
      const chats = rows.filter(r => r.incidents[inc.key]);
      return { inc, chats };
    }).filter(x => x.chats.length > 0);

    const totalIncs = agentIncs.reduce((s, x) => s + x.chats.length, 0);

    const bodyHtml = agentIncs.length
      ? agentIncs.map(x => `
          <div class="incident-row">
            <div class="incident-row-dot"></div>
            <div class="incident-row-text">
              <strong>${x.inc.label}</strong><br>
              <span class="incident-row-chats">Chats: ${x.chats.map(r => r.chatId).join(', ')} · ${x.chats.map(r => r.date).join(', ')}</span>
            </div>
            <div style="font-size:13px;font-weight:700;color:var(--warning);min-width:24px;text-align:right">${x.chats.length}</div>
          </div>`)
        .join('')
      : `<div class="no-incidents">✓ Sin incidencias en el período seleccionado</div>`;

    return `
      <div class="agent-incident-block">
        <div class="agent-incident-header">
          <div class="agent-incident-name">${agent}</div>
          <div class="agent-incident-count ${totalIncs === 0 ? 'zero' : ''}">${totalIncs} incidencia${totalIncs !== 1 ? 's' : ''}</div>
        </div>
        ${bodyHtml}
      </div>`;
  }).join('');
}

// ── Helpers ────────────────────────────────────────────────────────────────
function destroyChart(id) {
  if (charts[id]) { charts[id].destroy(); delete charts[id]; }
}

function truncate(str, n) {
  return str.length > n ? str.slice(0, n) + '…' : str;
}

// ── Template download ──────────────────────────────────────────────────────
function downloadTemplate() {
  const headers = ['Fecha','Agente','Chat_ID',
    ...CRITERIA.map(c => c.col),
    ...INCIDENTS.map(i => i.col),
  ];
  const example = [
    ['2024-06-03','María López','Chat_001','Sí','Sí','Sí','Sí','Sí','Sí','Sí','Sí','N/A','N/A','No','No','No','No','No','No','No','No'],
    ['2024-06-03','Carlos Ruiz','Chat_002','No','Sí','No','Sí','No','Sí','Sí','Sí','Sí','N/A','Sí','No','No','No','No','No','No','No'],
    ['2024-06-04','Ana García', 'Chat_003','Sí','No','Sí','No','Sí','Sí','Sí','Sí','N/A','Sí','No','No','Sí','No','No','No','No','Sí'],
  ];
  const ws = XLSX.utils.aoa_to_sheet([headers, ...example]);
  ws['!cols'] = headers.map(() => ({ wch: 26 }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Auditorías');
  XLSX.writeFile(wb, 'plantilla_auditoria.xlsx');
}

// ── Sample data ────────────────────────────────────────────────────────────
function generateSampleData() {
  const agents = ['María López','Carlos Ruiz','Ana García','Luis Torres'];
  const rows = [];
  const today = new Date();

  for (let d = 20; d >= 0; d--) {
    const date = new Date(today);
    date.setDate(today.getDate() - d);
    const dateStr = date.toISOString().slice(0,10);

    agents.forEach((agent, ai) => {
      if (Math.random() < 0.35) return;
      const numChats = Math.ceil(Math.random() * 3);
      for (let c = 0; c < numChats; c++) {
        const baseCompliance = 0.6 + ai * 0.08;
        const row = { Fecha: dateStr, Agente: agent, Chat_ID: `${agent.split(' ')[0]}_${dateStr}_${c+1}` };
        CRITERIA.forEach(cr => {
          if (cr.optional) row[cr.col] = Math.random() < 0.4 ? (Math.random() < baseCompliance ? 'Sí' : 'No') : 'N/A';
          else row[cr.col] = Math.random() < baseCompliance ? 'Sí' : 'No';
        });
        INCIDENTS.forEach(inc => {
          row[inc.col] = Math.random() < (0.15 - ai * 0.03) ? 'Sí' : 'No';
        });
        rows.push(row);
      }
    });
  }
  return rows;
}

// ── Toast ──────────────────────────────────────────────────────────────────
function showToast(msg) {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.classList.add('show');
  setTimeout(() => t.classList.remove('show'), 3500);
}
