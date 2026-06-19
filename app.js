/* Prospecting Map — Leaflet + SheetJS
 *
 * Reads a spreadsheet matching the provided format:
 *   Full name | Email | Role | Company | State | Zip Code |
 *   Beckhoff Proficiency | Course name | Enrolled | Started |
 *   Completed | Score | Course progress
 * (also handles the shorter Sheet2 format: Name | Email | Role |
 *  Company | State | Beckhoff Proficiency | Zip Code)
 *
 * Rows are placed on the map by Zip Code (centroid lookup),
 * falling back to the state centroid when the zip is unknown.
 */

// ---------------------------------------------------------------------------
// Header mapping — tolerant to casing / minor wording differences.
// ---------------------------------------------------------------------------
const FIELD_ALIASES = {
  name:        ['full name', 'name'],
  email:       ['email', 'e-mail'],
  role:        ['role', 'title', 'job title'],
  company:     ['company', 'organization', 'organisation'],
  state:       ['state'],
  zip:         ['zip code', 'zip', 'zipcode', 'postal code'],
  proficiency: ['beckhoff proficiency', 'proficiency'],
  course:      ['course name', 'course'],
  enrolled:    ['enrolled'],
  started:     ['started'],
  completed:   ['completed'],
  score:       ['score'],
  progress:    ['course progress', 'progress'],
};

const STATE_NAMES = {
  AL: 'Alabama', AK: 'Alaska', AZ: 'Arizona', AR: 'Arkansas', CA: 'California',
  CO: 'Colorado', CT: 'Connecticut', DE: 'Delaware', DC: 'District of Columbia',
  FL: 'Florida', GA: 'Georgia', HI: 'Hawaii', ID: 'Idaho', IL: 'Illinois',
  IN: 'Indiana', IA: 'Iowa', KS: 'Kansas', KY: 'Kentucky', LA: 'Louisiana',
  ME: 'Maine', MD: 'Maryland', MA: 'Massachusetts', MI: 'Michigan', MN: 'Minnesota',
  MS: 'Mississippi', MO: 'Missouri', MT: 'Montana', NE: 'Nebraska', NV: 'Nevada',
  NH: 'New Hampshire', NJ: 'New Jersey', NM: 'New Mexico', NY: 'New York',
  NC: 'North Carolina', ND: 'North Dakota', OH: 'Ohio', OK: 'Oklahoma', OR: 'Oregon',
  PA: 'Pennsylvania', RI: 'Rhode Island', SC: 'South Carolina', SD: 'South Dakota',
  TN: 'Tennessee', TX: 'Texas', UT: 'Utah', VT: 'Vermont', VA: 'Virginia',
  WA: 'Washington', WV: 'West Virginia', WI: 'Wisconsin', WY: 'Wyoming', PR: 'Puerto Rico',
};
const ABBR_SET = new Set(Object.keys(STATE_NAMES));
const NAME_TO_ABBR = Object.fromEntries(
  Object.entries(STATE_NAMES).map(([a, n]) => [n.toLowerCase(), a]));

// Normalize a state value (abbreviation or full name) to a 2-letter code.
function toAbbr(s) {
  const v = String(s || '').trim();
  if (!v) return '';
  if (ABBR_SET.has(v.toUpperCase())) return v.toUpperCase();
  return NAME_TO_ABBR[v.toLowerCase()] || v.toUpperCase();
}

const PROFICIENCY_COLORS = {
  'expert':       '#22c55e',
  'advanced':     '#84cc16',
  'intermediate': '#eab308',
  'beginner':     '#f97316',
  'novice':       '#f97316',
  'none':         '#ef4444',
};
const DEFAULT_COLOR = '#38bdf8';

// ---------------------------------------------------------------------------
// State
// ---------------------------------------------------------------------------
let ZIPS = {};
let STATES = {};
let allRows = [];          // normalized rows with geo info
let map, markerLayer, stateChart, stateLayer;
const selectedStates = new Set();   // abbreviations selected for download

// ---------------------------------------------------------------------------
// Boot
// ---------------------------------------------------------------------------
init();

async function init() {
  initMap();
  await loadGeoData();
  wireUI();
  updateSelectionUI();
  setStatus('Load your spreadsheet to plot prospects.');
}

function initMap() {
  map = L.map('map', { worldCopyJump: true }).setView([39.5, -98.35], 4); // US center
  L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
    maxZoom: 19,
    attribution: '&copy; OpenStreetMap contributors',
  }).addTo(map);
  markerLayer = L.featureGroup().addTo(map); // individual markers, no clustering
}

async function loadGeoData() {
  try {
    const [z, s, geo] = await Promise.all([
      fetch('data/zipcodes.min.json').then(r => r.json()),
      fetch('data/states.min.json').then(r => r.json()),
      fetch('data/us-states.min.json').then(r => r.json()),
    ]);
    ZIPS = z; STATES = s;
    addStateLayer(geo);
  } catch (e) {
    console.error('Could not load geo lookup data', e);
    setStatus('Warning: geo lookup data failed to load.');
  }
}

// Clickable state boundaries: click a state to select it for download.
function addStateLayer(geo) {
  stateLayer = L.geoJSON(geo, {
    style: styleState,
    onEachFeature: (feature, layer) => {
      const abbr = feature.properties.abbr;
      layer.on({
        click: () => toggleState(abbr),
        mouseover: () => { if (!selectedStates.has(abbr)) layer.setStyle({ fillOpacity: 0.15 }); },
        mouseout: () => stateLayer.resetStyle(layer),
      });
      layer.bindTooltip(feature.properties.name, { sticky: true });
    },
  }).addTo(map);
  stateLayer.bringToBack(); // keep markers clickable on top
}

function styleState(feature) {
  const selected = selectedStates.has(feature.properties.abbr);
  return {
    color: selected ? '#38bdf8' : '#64748b',
    weight: selected ? 2 : 1,
    fillColor: '#38bdf8',
    fillOpacity: selected ? 0.35 : 0.01, // ~invisible but keeps the area clickable
  };
}

function toggleState(abbr) {
  if (selectedStates.has(abbr)) selectedStates.delete(abbr);
  else selectedStates.add(abbr);
  if (stateLayer) stateLayer.setStyle(styleState);
  updateSelectionUI();
}

// ---------------------------------------------------------------------------
// Spreadsheet parsing
// ---------------------------------------------------------------------------
function parseWorkbook(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: 'array' });
  // Pick the sheet that has the most data rows.
  let best = null;
  for (const name of wb.SheetNames) {
    const json = XLSX.utils.sheet_to_json(wb.Sheets[name], { defval: '' });
    if (!best || json.length > best.length) best = json;
  }
  return best || [];
}

function buildHeaderMap(sampleRow) {
  // Map canonical field -> actual key present in the row.
  const keys = Object.keys(sampleRow);
  const map = {};
  for (const [canon, aliases] of Object.entries(FIELD_ALIASES)) {
    const hit = keys.find(k => aliases.includes(String(k).trim().toLowerCase()));
    if (hit) map[canon] = hit;
  }
  return map;
}

function normalizeRows(rawRows) {
  if (!rawRows.length) return [];
  const hmap = buildHeaderMap(rawRows[0]);
  const get = (row, canon) => (hmap[canon] !== undefined ? row[hmap[canon]] : '');

  return rawRows.map((row, i) => {
    const zipRaw = String(get(row, 'zip') ?? '').trim();
    const zip = zipRaw ? zipRaw.split('-')[0].padStart(5, '0').slice(0, 5) : '';
    const state = String(get(row, 'state') ?? '').trim().toUpperCase();

    return {
      _i: i,
      _raw: row,
      stateAbbr: toAbbr(state),
      name: String(get(row, 'name') ?? '').trim(),
      email: String(get(row, 'email') ?? '').trim(),
      role: String(get(row, 'role') ?? '').trim(),
      company: String(get(row, 'company') ?? '').trim(),
      state,
      zip,
      proficiency: String(get(row, 'proficiency') ?? '').trim(),
      course: String(get(row, 'course') ?? '').trim(),
      enrolled: String(get(row, 'enrolled') ?? '').trim(),
      started: String(get(row, 'started') ?? '').trim(),
      completed: String(get(row, 'completed') ?? '').trim(),
      score: String(get(row, 'score') ?? '').trim(),
      progress: parseProgress(get(row, 'progress')),
      _geo: geocode(zip, state, i),
    };
  });
}

function parseProgress(v) {
  if (v === '' || v == null) return null;
  const n = parseFloat(String(v).replace('%', '').trim());
  return isNaN(n) ? null : n;
}

// Place a row: prefer zip centroid, else state centroid (with deterministic
// jitter so co-located prospects don't perfectly overlap).
function geocode(zip, state, idx) {
  if (zip && ZIPS[zip]) return { lat: ZIPS[zip][0], lng: ZIPS[zip][1], approx: false };
  if (state && STATES[state]) {
    const j = jitter(idx);
    return { lat: STATES[state][0] + j[0], lng: STATES[state][1] + j[1], approx: true };
  }
  return null;
}

function jitter(idx) {
  // ~ +/- 0.35 deg, deterministic per index
  const a = (Math.sin(idx * 12.9898) * 43758.5453) % 1;
  const b = (Math.sin(idx * 78.233) * 12345.6789) % 1;
  return [(a - 0.5) * 0.7, (b - 0.5) * 0.7];
}

// ---------------------------------------------------------------------------
// Load + render
// ---------------------------------------------------------------------------
function loadRows(rawRows) {
  allRows = normalizeRows(rawRows);
  populateFilters();
  buildLegend();
  applyFilters();
  updateSelectionUI();
}

function applyFilters() {
  const f = readFilters();
  const rows = allRows.filter(r => matches(r, f));

  markerLayer.clearLayers();
  let placed = 0;
  for (const r of rows) {
    if (!r._geo) continue;
    placed++;
    makeMarker(r).addTo(markerLayer);
  }

  const unplaced = rows.length - placed;
  document.getElementById('summary').textContent =
    `${rows.length} prospect${rows.length === 1 ? '' : 's'} shown` +
    (unplaced ? ` (${unplaced} without a mappable location)` : '');

  if (placed) {
    map.fitBounds(markerLayer.getBounds().pad(0.15), { maxZoom: 11 });
  }

  updateChart(rows);
}

function matches(r, f) {
  if (f.search) {
    const hay = `${r.name} ${r.email} ${r.company}`.toLowerCase();
    if (!hay.includes(f.search)) return false;
  }
  if (f.state && r.state !== f.state) return false;
  if (f.role && r.role !== f.role) return false;
  if (f.company && r.company !== f.company) return false;
  if (f.proficiency && r.proficiency !== f.proficiency) return false;
  if (f.course && r.course !== f.course) return false;
  if (f.minProgress > 0 && (r.progress == null || r.progress < f.minProgress)) return false;
  if (f.completedOnly && !isCompleted(r)) return false;
  return true;
}

function isCompleted(r) {
  const c = r.completed.toLowerCase();
  if (['yes', 'true', 'y', '1', 'complete', 'completed'].includes(c)) return true;
  if (c && !['no', 'false', 'n', '0', ''].includes(c)) return true; // a date counts as completed
  return r.progress === 100;
}

function makeMarker(r) {
  const color = colorFor(r.proficiency);
  const icon = L.divIcon({
    className: '',
    html: `<div class="pin" style="background:${color}"></div>`,
    iconSize: [18, 18],
    iconAnchor: [9, 18],
    popupAnchor: [0, -16],
  });
  const m = L.marker([r._geo.lat, r._geo.lng], { icon, title: r.name });
  m.bindPopup(popupHtml(r));
  return m;
}

function popupHtml(r) {
  const row = (k, v) => v ? `<tr><td class="k">${k}</td><td>${escapeHtml(v)}</td></tr>` : '';
  const loc = [r.zip, r.state].filter(Boolean).join(' · ') + (r._geo && r._geo.approx ? ' (approx.)' : '');
  return `<div class="popup">
    <h3>${escapeHtml(r.name || '(no name)')}</h3>
    <table>
      ${row('Role', r.role)}
      ${row('Company', r.company)}
      ${row('Email', r.email)}
      ${row('Location', loc)}
      ${row('Proficiency', r.proficiency)}
      ${row('Course', r.course)}
      ${r.progress != null ? row('Progress', r.progress + '%') : ''}
      ${row('Score', r.score)}
      ${row('Completed', r.completed)}
    </table>
  </div>`;
}

function colorFor(prof) {
  return PROFICIENCY_COLORS[String(prof).trim().toLowerCase()] || DEFAULT_COLOR;
}

// ---------------------------------------------------------------------------
// Filters UI
// ---------------------------------------------------------------------------
function uniqueSorted(field) {
  return [...new Set(allRows.map(r => r[field]).filter(Boolean))]
    .sort((a, b) => String(a).localeCompare(String(b)));
}

function populateFilters() {
  fillSelect('filterState', uniqueSorted('state'), 'All states');
  fillSelect('filterRole', uniqueSorted('role'), 'All roles');
  fillSelect('filterCompany', uniqueSorted('company'), 'All companies');
  fillSelect('filterProficiency', uniqueSorted('proficiency'), 'All proficiencies');
  fillSelect('filterCourse', uniqueSorted('course'), 'All courses');

  // Hide course-related filters if the data has no course info.
  const hasCourses = uniqueSorted('course').length > 0;
  const hasProgress = allRows.some(r => r.progress != null);
  toggle('courseField', hasCourses);
  toggle('progressField', hasProgress);
  toggle('completedField', allRows.some(r => r.completed) || hasProgress);
}

function fillSelect(id, values, allLabel) {
  const sel = document.getElementById(id);
  sel.innerHTML = `<option value="">${allLabel}</option>` +
    values.map(v => `<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`).join('');
}

function readFilters() {
  return {
    search: document.getElementById('search').value.trim().toLowerCase(),
    state: document.getElementById('filterState').value,
    role: document.getElementById('filterRole').value,
    company: document.getElementById('filterCompany').value,
    proficiency: document.getElementById('filterProficiency').value,
    course: document.getElementById('filterCourse').value,
    minProgress: Number(document.getElementById('filterProgress').value),
    completedOnly: document.getElementById('filterCompleted').checked,
  };
}

function buildLegend() {
  const present = uniqueSorted('proficiency');
  const ul = document.getElementById('legend');
  const items = present.length ? present : Object.keys(PROFICIENCY_COLORS);
  ul.innerHTML = items.map(p =>
    `<li><span class="dot" style="background:${colorFor(p)}"></span>${escapeHtml(p)}</li>`
  ).join('') + `<li><span class="dot" style="background:${DEFAULT_COLOR}"></span>Other / unspecified</li>`;
}

// ---------------------------------------------------------------------------
// Pie chart — prospects by state (reflects the active filters)
// ---------------------------------------------------------------------------
const CHART_COLORS = [
  '#38bdf8', '#22c55e', '#eab308', '#f97316', '#ef4444', '#a855f7',
  '#ec4899', '#14b8a6', '#84cc16', '#f59e0b', '#6366f1', '#06b6d4',
];

function updateChart(rows) {
  const counts = {};
  for (const r of rows) {
    const s = r.state || 'Unknown';
    counts[s] = (counts[s] || 0) + 1;
  }
  const entries = Object.entries(counts).sort((a, b) => b[1] - a[1]);

  const emptyMsg = document.getElementById('chartEmpty');
  const canvas = document.getElementById('stateChart');
  if (!entries.length) {
    if (stateChart) { stateChart.destroy(); stateChart = null; }
    emptyMsg.style.display = '';
    canvas.style.display = 'none';
    return;
  }
  emptyMsg.style.display = 'none';
  canvas.style.display = '';

  const labels = entries.map(e => e[0]);
  const data = entries.map(e => e[1]);
  const colors = labels.map((_, i) => CHART_COLORS[i % CHART_COLORS.length]);

  if (stateChart) {
    stateChart.data.labels = labels;
    stateChart.data.datasets[0].data = data;
    stateChart.data.datasets[0].backgroundColor = colors;
    stateChart.update();
    return;
  }

  stateChart = new Chart(canvas.getContext('2d'), {
    type: 'pie',
    data: { labels, datasets: [{ data, backgroundColor: colors, borderColor: '#1e293b', borderWidth: 1 }] },
    options: {
      responsive: true,
      plugins: {
        legend: { position: 'bottom', labels: { color: '#e2e8f0', boxWidth: 12, font: { size: 11 } } },
        tooltip: {
          callbacks: {
            label: ctx => {
              const total = ctx.dataset.data.reduce((a, b) => a + b, 0);
              const pct = total ? Math.round((ctx.parsed / total) * 100) : 0;
              return `${ctx.label}: ${ctx.parsed} (${pct}%)`;
            },
          },
        },
      },
    },
  });
}

// ---------------------------------------------------------------------------
// State selection + download
// ---------------------------------------------------------------------------
function countInState(abbr) {
  return allRows.reduce((n, r) => n + (r.stateAbbr === abbr ? 1 : 0), 0);
}

function updateSelectionUI() {
  const chips = document.getElementById('selectedChips');
  const btn = document.getElementById('downloadBtn');
  const clear = document.getElementById('clearStates');
  const list = [...selectedStates].sort();

  if (!list.length) {
    chips.innerHTML = '<span class="muted small">No states selected. Click states on the map.</span>';
  } else {
    chips.innerHTML = list.map(a => {
      const c = countInState(a);
      return `<span class="chip" data-abbr="${a}">${a} <small>${c}</small> <span class="x">&times;</span></span>`;
    }).join('');
    chips.querySelectorAll('.chip').forEach(el =>
      el.addEventListener('click', () => toggleState(el.dataset.abbr)));
  }

  const total = list.reduce((n, a) => n + countInState(a), 0);
  btn.disabled = !list.length || total === 0;
  btn.textContent = list.length
    ? `Download ${total} row${total === 1 ? '' : 's'} (${list.length} state${list.length === 1 ? '' : 's'})`
    : 'Download selected states';
  clear.style.display = list.length ? '' : 'none';
}

function downloadSelected() {
  const rows = allRows.filter(r => selectedStates.has(r.stateAbbr)).map(r => r._raw);
  if (!rows.length) return;
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Prospects');
  const name = `prospects_${[...selectedStates].sort().join('-')}.xlsx`;
  XLSX.writeFile(wb, name);
}

function clearSelectedStates() {
  selectedStates.clear();
  if (stateLayer) stateLayer.setStyle(styleState);
  updateSelectionUI();
}

// ---------------------------------------------------------------------------
// Wiring
// ---------------------------------------------------------------------------
function wireUI() {
  ['filterState', 'filterRole', 'filterCompany', 'filterProficiency', 'filterCourse',
   'filterCompleted'].forEach(id =>
    document.getElementById(id).addEventListener('change', applyFilters));

  document.getElementById('search').addEventListener('input', debounce(applyFilters, 200));

  document.getElementById('downloadBtn').addEventListener('click', downloadSelected);
  document.getElementById('clearStates').addEventListener('click', clearSelectedStates);

  const prog = document.getElementById('filterProgress');
  prog.addEventListener('input', () => {
    document.getElementById('progressVal').textContent = prog.value + '%';
    applyFilters();
  });

  document.getElementById('resetFilters').addEventListener('click', () => {
    document.getElementById('search').value = '';
    ['filterState', 'filterRole', 'filterCompany', 'filterProficiency', 'filterCourse'].forEach(id =>
      document.getElementById(id).value = '');
    prog.value = 0;
    document.getElementById('progressVal').textContent = '0%';
    document.getElementById('filterCompleted').checked = false;
    applyFilters();
  });

  // File input + drag & drop
  const drop = document.getElementById('fileDrop');
  const input = document.getElementById('fileInput');
  input.addEventListener('change', e => { if (e.target.files[0]) handleFile(e.target.files[0]); });
  ['dragover', 'dragenter'].forEach(ev =>
    drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('dragover'); }));
  ['dragleave', 'drop'].forEach(ev =>
    drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('dragover'); }));
  drop.addEventListener('drop', e => {
    const file = e.dataTransfer.files[0];
    if (file) handleFile(file);
  });
}

function handleFile(file) {
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const rows = parseWorkbook(e.target.result);
      if (!rows.length) { setStatus('No rows found in that file.'); return; }
      loadRows(rows);
      setStatus(`Loaded ${rows.length} rows from “${file.name}”.`);
    } catch (err) {
      console.error(err);
      setStatus('Could not read that file. Use the .xlsx format shown in the README.');
    }
  };
  reader.readAsArrayBuffer(file);
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------
function setStatus(msg) { document.getElementById('dataStatus').textContent = msg; }
function toggle(id, show) { document.getElementById(id).style.display = show ? '' : 'none'; }

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, c =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function debounce(fn, ms) {
  let t;
  return (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), ms); };
}
