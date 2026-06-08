/* ── Auditoría Equipo de Ventas · app.js ── */

// ── State ──────────────────────────────────────────────────────────────────
let allRows = [];       // raw parsed rows
let filtered = [];      // rows after filters
let categoryCols = [];  // detected category column names
let charts = {};        // Chart.js instances keyed by canvas id
let sortCol = null;
let sortAsc = true;

// ── Column name aliases (case-insensitive) ─────────────────────────────────
const COL = {
  date:  ['fecha', 'date', 'día', 'dia'],
  agent: ['agente', 'agent', 'nombre', 'asesor', 'vendedor'],
  total: ['puntaje total', 'total', 'score', 'puntaje', 'calificación', 'calificacion', 'nota'],
};

// ── Colour palette ────────────────────────────────────────────────────────
const PALETTE = [
  '#2563eb','#7c3aed','#db2777','#d97706','#16a34a',
  '#0891b2','#9333ea','#ea580c','#65a30d','#0284c7',
];

// ── Entry points ───────────────────────────────────────────────────────────
function handleFile(evt) {
  const file = evt.target.files[0];
  if (!file) return;
  const ext = file.name.split('.').pop().toLowerCase();
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
  // reset input so same file can be re-loaded
  evt.target.value = '';
}

function loadSampleData() {
  const sample = generateSampleData();
  processData(sample, 'datos_ejemplo');
}

// ── Process & normalise data ───────────────────────────────────────────────
function processData(rows, filename) {
  if (!rows.length) { showToast('El archivo está vacío.'); return; }

  const headers = Object.keys(rows[0]);

  // detect column roles
  const dateCol  = findCol(headers, COL.date);
  const agentCol = findCol(headers, COL.agent);
  const totalCol = findCol(headers, COL.total);

  if (!agentCol) { showToast('No se encontró columna de agente. Revisa la plantilla.'); return; }

  // category cols = everything that is not date/agent/total
  const reserved = [dateCol, agentCol, totalCol].filter(Boolean).map(s => s.toLowerCase());
  categoryCols = headers.filter(h => !reserved.includes(h.toLowerCase()));

  // normalise rows
  allRows = rows.map((r, i) => {
    const dateRaw = dateCol ? r[dateCol] : null;
    const dateVal = parseDate(dateRaw) || `Fila ${i+2}`;
    const agent   = String(r[agentCol] || 'Sin nombre').trim();
    const cats    = {};
    let sumCats = 0, countCats = 0;
    categoryCols.forEach(c => {
      const v = parseFloat(r[c]);
      cats[c] = isNaN(v) ? null : v;
      if (!isNaN(v)) { sumCats += v; countCats++; }
    });
    let total = totalCol ? parseFloat(r[totalCol]) : NaN;
    if (isNaN(total) && countCats) total = +(sumCats / countCats).toFixed(1);
    return { date: dateVal, agent, cats, total: isNaN(total) ? null : total, _raw: r };
  });

  showToast(`✓ ${allRows.length} auditorías cargadas desde "${filename}"`);
  buildFilterOptions();
  applyFilters();
  document.getElementById('uploadZone').style.display = 'none';
  document.getElementById('dashboard').style.display = 'block';
}

function findCol(headers, aliases) {
  for (const h of headers)
    if (aliases.some(a => h.toLowerCase().trim() === a)) return h;
  // partial match fallback
  for (const h of headers)
    if (aliases.some(a => h.toLowerCase().includes(a))) return h;
  return null;
}

function parseDate(v) {
  if (!v) return null;
  if (v instanceof Date) return v.toISOString().slice(0,10);
  const d = new Date(v);
  if (!isNaN(d)) return d.toISOString().slice(0,10);
  return String(v);
}

// ── Filters ────────────────────────────────────────────────────────────────
function buildFilterOptions() {
  const agents = [...new Set(allRows.map(r => r.agent))].sort();
  const sel = document.getElementById('filterAgent');
  sel.innerHTML = '<option value="">Todos</option>';
  agents.forEach(a => {
    const opt = document.createElement('option');
    opt.value = opt.textContent = a;
    sel.appendChild(opt);
  });
}

function applyFilters() {
  const agent = document.getElementById('filterAgent').value;
  const from  = document.getElementById('filterFrom').value;
  const to    = document.getElementById('filterTo').value;

  filtered = allRows.filter(r => {
    if (agent && r.agent !== agent) return false;
    if (from && r.date < from) return false;
    if (to   && r.date > to)   return false;
    return true;
  });

  document.getElementById('recordCount').textContent =
    `${filtered.length} registro${filtered.length !== 1 ? 's' : ''}`;

  renderKPIs();
  renderCharts();
  renderTable();
}

function clearFilters() {
  document.getElementById('filterAgent').value = '';
  document.getElementById('filterFrom').value  = '';
  document.getElementById('filterTo').value    = '';
  applyFilters();
}

// ── KPIs ───────────────────────────────────────────────────────────────────
function renderKPIs() {
  const totals  = filtered.map(r => r.total).filter(v => v !== null);
  const avg     = totals.length ? +(totals.reduce((a,b) => a+b, 0) / totals.length).toFixed(1) : 0;
  const agents  = new Set(filtered.map(r => r.agent)).size;
  const audits  = filtered.length;
  const passing = totals.filter(v => v >= 70).length;
  const pct     = totals.length ? Math.round(passing / totals.length * 100) : 0;

  const scoreClass = avg >= 80 ? 'success' : avg >= 60 ? 'warning' : 'danger';

  document.getElementById('kpiGrid').innerHTML = `
    <div class="kpi-card ${scoreClass}">
      <div class="kpi-label">Puntaje promedio</div>
      <div class="kpi-value">${avg}<span style="font-size:18px;font-weight:600">/100</span></div>
      <div class="kpi-sub">En el período seleccionado</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">Total auditorías</div>
      <div class="kpi-value">${audits}</div>
      <div class="kpi-sub">Chats revisados</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">Agentes evaluados</div>
      <div class="kpi-value">${agents}</div>
      <div class="kpi-sub">Agentes únicos</div>
    </div>
    <div class="kpi-card ${pct >= 70 ? 'success' : pct >= 50 ? 'warning' : 'danger'}">
      <div class="kpi-label">Tasa de aprobación</div>
      <div class="kpi-value">${pct}<span style="font-size:18px;font-weight:600">%</span></div>
      <div class="kpi-sub">Puntaje ≥ 70 pts</div>
    </div>
  `;
}

// ── Charts ─────────────────────────────────────────────────────────────────
function destroyChart(id) {
  if (charts[id]) { charts[id].destroy(); delete charts[id]; }
}

function renderCharts() {
  renderAgentChart();
  renderTrendChart();
  renderCategoryChart();
  renderDistChart();
}

function renderAgentChart() {
  destroyChart('chartAgents');
  const agents = [...new Set(filtered.map(r => r.agent))].sort();
  const data = agents.map(a => {
    const rows = filtered.filter(r => r.agent === a);
    const vals = rows.map(r => r.total).filter(v => v !== null);
    return vals.length ? +(vals.reduce((x,y) => x+y, 0) / vals.length).toFixed(1) : 0;
  });
  charts['chartAgents'] = new Chart(document.getElementById('chartAgents'), {
    type: 'bar',
    data: {
      labels: agents,
      datasets: [{
        label: 'Puntaje promedio',
        data,
        backgroundColor: agents.map((_,i) => PALETTE[i % PALETTE.length] + 'cc'),
        borderColor:     agents.map((_,i) => PALETTE[i % PALETTE.length]),
        borderWidth: 2, borderRadius: 6,
      }],
    },
    options: chartOpts({ min: 0, max: 100, yLabel: 'Puntaje' }),
  });
}

function renderTrendChart() {
  destroyChart('chartTrend');
  const dates = [...new Set(filtered.map(r => r.date))].sort();
  const agents = [...new Set(filtered.map(r => r.agent))].sort();
  const datasets = agents.slice(0, 6).map((a, i) => {
    const color = PALETTE[i % PALETTE.length];
    return {
      label: a,
      data: dates.map(d => {
        const rows = filtered.filter(r => r.agent === a && r.date === d);
        const vals = rows.map(r => r.total).filter(v => v !== null);
        return vals.length ? +(vals.reduce((x,y) => x+y,0)/vals.length).toFixed(1) : null;
      }),
      borderColor: color, backgroundColor: color + '22',
      tension: .3, fill: false,
      pointRadius: 4, pointHoverRadius: 6,
      spanGaps: true,
    };
  });
  charts['chartTrend'] = new Chart(document.getElementById('chartTrend'), {
    type: 'line',
    data: { labels: dates, datasets },
    options: chartOpts({ min: 0, max: 100, yLabel: 'Puntaje', legend: true }),
  });
}

function renderCategoryChart() {
  destroyChart('chartCategories');
  if (!categoryCols.length) return;
  const avgs = categoryCols.map(c => {
    const vals = filtered.map(r => r.cats[c]).filter(v => v !== null);
    return vals.length ? +(vals.reduce((a,b) => a+b,0)/vals.length).toFixed(1) : 0;
  });
  charts['chartCategories'] = new Chart(document.getElementById('chartCategories'), {
    type: 'radar',
    data: {
      labels: categoryCols,
      datasets: [{
        label: 'Promedio',
        data: avgs,
        backgroundColor: '#2563eb22',
        borderColor: '#2563eb',
        pointBackgroundColor: '#2563eb',
        borderWidth: 2, pointRadius: 4,
      }],
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      scales: { r: { min: 0, max: 100, ticks: { stepSize: 20, font: { size: 10 } }, pointLabels: { font: { size: 11 } } } },
      plugins: { legend: { display: false } },
    },
  });
}

function renderDistChart() {
  destroyChart('chartDist');
  const buckets = ['0–39','40–59','60–69','70–79','80–89','90–100'];
  const counts  = [0,0,0,0,0,0];
  filtered.forEach(r => {
    if (r.total === null) return;
    const v = r.total;
    if (v < 40) counts[0]++;
    else if (v < 60) counts[1]++;
    else if (v < 70) counts[2]++;
    else if (v < 80) counts[3]++;
    else if (v < 90) counts[4]++;
    else counts[5]++;
  });
  charts['chartDist'] = new Chart(document.getElementById('chartDist'), {
    type: 'bar',
    data: {
      labels: buckets,
      datasets: [{
        label: 'Auditorías',
        data: counts,
        backgroundColor: ['#dc262688','#ea580c88','#d9770688','#2563eb88','#16a34a88','#16a34acc'],
        borderColor:     ['#dc2626','#ea580c','#d97706','#2563eb','#16a34a','#16a34a'],
        borderWidth: 2, borderRadius: 6,
      }],
    },
    options: chartOpts({ yLabel: 'Cantidad', integer: true }),
  });
}

function chartOpts({ min, max, yLabel, legend, integer } = {}) {
  return {
    responsive: true, maintainAspectRatio: false,
    plugins: { legend: { display: !!legend, position: 'top', labels: { font: { size: 11 }, boxWidth: 12 } } },
    scales: {
      y: {
        min, max,
        ticks: {
          font: { size: 11 },
          callback: integer ? (v => Number.isInteger(v) ? v : '') : undefined,
        },
        title: { display: !!yLabel, text: yLabel, font: { size: 11 } },
        grid: { color: '#e2e8f0' },
      },
      x: { ticks: { font: { size: 11 } }, grid: { display: false } },
    },
  };
}

// ── Table ──────────────────────────────────────────────────────────────────
function renderTable() {
  const cols = ['date', 'agent', ...categoryCols, 'total'];
  const labels = {
    date: 'Fecha', agent: 'Agente', total: 'Puntaje Total',
  };
  const getLabel = c => labels[c] || c;

  // Header
  document.getElementById('tableHead').innerHTML =
    '<tr>' + cols.map(c =>
      `<th onclick="sortTable('${c}')">${getLabel(c)} ${sortCol===c ? (sortAsc?'▲':'▼') : ''}</th>`
    ).join('') + '</tr>';

  // Sort
  let rows = [...filtered];
  if (sortCol) {
    rows.sort((a, b) => {
      const va = sortCol === 'date' ? a.date : sortCol === 'agent' ? a.agent : sortCol === 'total' ? a.total : (a.cats[sortCol] ?? -1);
      const vb = sortCol === 'date' ? b.date : sortCol === 'agent' ? b.agent : sortCol === 'total' ? b.total : (b.cats[sortCol] ?? -1);
      if (va < vb) return sortAsc ? -1 : 1;
      if (va > vb) return sortAsc ?  1 : -1;
      return 0;
    });
  }

  // Body
  document.getElementById('tableBody').innerHTML = rows.map(r =>
    '<tr>' + cols.map(c => {
      if (c === 'date')  return `<td>${r.date}</td>`;
      if (c === 'agent') return `<td><strong>${r.agent}</strong></td>`;
      if (c === 'total') return `<td>${scoreBadge(r.total)}</td>`;
      const v = r.cats[c];
      return `<td>${v !== null ? scoreBadge(v) : '—'}</td>`;
    }).join('') + '</tr>'
  ).join('');
}

function scoreBadge(v) {
  if (v === null || v === undefined) return '—';
  const cls = v >= 80 ? 'badge-success' : v >= 60 ? 'badge-warning' : 'badge-danger';
  return `<span class="badge ${cls}">${v}</span>`;
}

function sortTable(col) {
  if (sortCol === col) sortAsc = !sortAsc;
  else { sortCol = col; sortAsc = true; }
  renderTable();
}

// ── Export ─────────────────────────────────────────────────────────────────
function exportFiltered() {
  const ws = XLSX.utils.json_to_sheet(filtered.map(r => ({
    Fecha: r.date, Agente: r.agent,
    ...Object.fromEntries(categoryCols.map(c => [c, r.cats[c]])),
    'Puntaje Total': r.total,
  })));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Auditorías');
  XLSX.writeFile(wb, 'auditoria_exportada.xlsx');
}

// ── Template download ──────────────────────────────────────────────────────
function downloadTemplate() {
  const headers = ['Fecha','Agente','Saludo y Presentación','Manejo de Objeciones','Lenguaje y Tono','Seguimiento al Cliente','Cierre de Venta','Puntaje Total'];
  const example = [
    ['2024-06-01','María López',85,90,80,75,88,84],
    ['2024-06-01','Carlos Ruiz',70,65,78,80,72,73],
    ['2024-06-02','María López',88,92,85,90,95,90],
  ];
  const ws = XLSX.utils.aoa_to_sheet([headers, ...example]);
  ws['!cols'] = headers.map(() => ({ wch: 22 }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Auditorías');
  XLSX.writeFile(wb, 'plantilla_auditoria.xlsx');
}

// ── Sample data generator ──────────────────────────────────────────────────
function generateSampleData() {
  const agents = ['María López','Carlos Ruiz','Ana García','Luis Torres','Sofía Méndez'];
  const cats   = ['Saludo y Presentación','Manejo de Objeciones','Lenguaje y Tono','Seguimiento al Cliente','Cierre de Venta'];
  const rows   = [];
  const today  = new Date();

  for (let d = 29; d >= 0; d--) {
    const date = new Date(today);
    date.setDate(today.getDate() - d);
    const dateStr = date.toISOString().slice(0,10);

    agents.forEach((agent, ai) => {
      if (Math.random() < 0.3) return; // not every agent every day
      const base = 65 + ai * 5;
      const catScores = {};
      let sum = 0;
      cats.forEach(c => {
        const v = Math.min(100, Math.max(40, Math.round(base + (Math.random()*30 - 15))));
        catScores[c] = v; sum += v;
      });
      const total = Math.round(sum / cats.length);
      rows.push({ Fecha: dateStr, Agente: agent, ...catScores, 'Puntaje Total': total });
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
