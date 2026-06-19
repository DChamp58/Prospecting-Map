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
let map, cluster;

// ---------------------------------------------------------------------------
// Boot
// ---------------------------------------------------------------------------
init();

async function init() {
  initMap();
  await loadGeoData();
  wireUI();
  loadRows(sampleData());
  setStatus('Showing sample data — load your spreadsheet to replace it.');
}

function initMap() {
  map = L.map('map', { worldCopyJump: true }).setView([39.5, -98.35], 4); // US center
  L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
    maxZoom: 19,
    attribution: '&copy; OpenStreetMap contributors',
  }).addTo(map);
  cluster = L.markerClusterGroup({ maxClusterRadius: 45 });
  map.addLayer(cluster);
}

async function loadGeoData() {
  try {
    const [z, s] = await Promise.all([
      fetch('data/zipcodes.min.json').then(r => r.json()),
      fetch('data/states.min.json').then(r => r.json()),
    ]);
    ZIPS = z; STATES = s;
  } catch (e) {
    console.error('Could not load geo lookup data', e);
    setStatus('Warning: geo lookup data failed to load.');
  }
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
}

function applyFilters() {
  const f = readFilters();
  const rows = allRows.filter(r => matches(r, f));

  cluster.clearLayers();
  const markers = [];
  let placed = 0;
  for (const r of rows) {
    if (!r._geo) continue;
    placed++;
    markers.push(makeMarker(r));
  }
  cluster.addLayers(markers);

  const unplaced = rows.length - placed;
  document.getElementById('summary').textContent =
    `${rows.length} prospect${rows.length === 1 ? '' : 's'} shown` +
    (unplaced ? ` (${unplaced} without a mappable location)` : '');

  if (markers.length) {
    map.fitBounds(cluster.getBounds().pad(0.15), { maxZoom: 11 });
  }
}

function matches(r, f) {
  if (f.search) {
    const hay = `${r.name} ${r.email} ${r.company}`.toLowerCase();
    if (!hay.includes(f.search)) return false;
  }
  if (f.states.length && !f.states.includes(r.state)) return false;
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
  fillMultiSelect('filterState', uniqueSorted('state'));
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

function fillMultiSelect(id, values) {
  const sel = document.getElementById(id);
  sel.innerHTML = values.map(v => `<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`).join('');
}

function readFilters() {
  return {
    search: document.getElementById('search').value.trim().toLowerCase(),
    states: [...document.getElementById('filterState').selectedOptions].map(o => o.value),
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
// Wiring
// ---------------------------------------------------------------------------
function wireUI() {
  ['filterState', 'filterRole', 'filterCompany', 'filterProficiency', 'filterCourse',
   'filterCompleted'].forEach(id =>
    document.getElementById(id).addEventListener('change', applyFilters));

  document.getElementById('search').addEventListener('input', debounce(applyFilters, 200));

  const prog = document.getElementById('filterProgress');
  prog.addEventListener('input', () => {
    document.getElementById('progressVal').textContent = prog.value + '%';
    applyFilters();
  });

  document.getElementById('resetFilters').addEventListener('click', () => {
    document.getElementById('search').value = '';
    [...document.getElementById('filterState').options].forEach(o => o.selected = false);
    ['filterRole', 'filterCompany', 'filterProficiency', 'filterCourse'].forEach(id =>
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

// ---------------------------------------------------------------------------
// Sample data (matches the provided Sheet1 format)
// ---------------------------------------------------------------------------
function sampleData() {
  return [
    { 'Full name': 'Jane Cooper', Email: 'jane@acme.com', Role: 'Controls Engineer', Company: 'Acme Automation', State: 'TX', 'Zip Code': '75201', 'Beckhoff Proficiency': 'Intermediate', 'Course name': 'TwinCAT 3 Basics', Enrolled: '2026-01-10', Started: '2026-01-12', Completed: '2026-02-01', Score: '88', 'Course progress': 100 },
    { 'Full name': 'Marcus Lee', Email: 'marcus@nordic.io', Role: 'Automation Lead', Company: 'Nordic Systems', State: 'MN', 'Zip Code': '55401', 'Beckhoff Proficiency': 'Advanced', 'Course name': 'TwinCAT 3 Basics', Enrolled: '2026-02-03', Started: '2026-02-05', Completed: '', Score: '', 'Course progress': 60 },
    { 'Full name': 'Priya Patel', Email: 'priya@westcoast.com', Role: 'Project Manager', Company: 'West Coast Integrators', State: 'CA', 'Zip Code': '94103', 'Beckhoff Proficiency': 'Beginner', 'Course name': 'PLC Fundamentals', Enrolled: '2026-03-01', Started: '', Completed: '', Score: '', 'Course progress': 0 },
    { 'Full name': 'Tom Becker', Email: 'tom@greatlakes.com', Role: 'Controls Engineer', Company: 'Great Lakes Mfg', State: 'MI', 'Zip Code': '48226', 'Beckhoff Proficiency': 'Expert', 'Course name': 'Motion Control', Enrolled: '2025-11-15', Started: '2025-11-16', Completed: '2025-12-20', Score: '95', 'Course progress': 100 },
    { 'Full name': 'Sara Nguyen', Email: 'sara@gulfauto.com', Role: 'Sales Engineer', Company: 'Gulf Automation', State: 'FL', 'Zip Code': '33101', 'Beckhoff Proficiency': 'Intermediate', 'Course name': 'PLC Fundamentals', Enrolled: '2026-04-10', Started: '2026-04-11', Completed: '', Score: '', 'Course progress': 35 },
    { 'Full name': 'David Kim', Email: 'david@empire.com', Role: 'Automation Lead', Company: 'Empire Controls', State: 'NY', 'Zip Code': '10001', 'Beckhoff Proficiency': 'Advanced', 'Course name': 'Motion Control', Enrolled: '2026-01-22', Started: '2026-01-25', Completed: '', Score: '', 'Course progress': 75 },
    { 'Full name': 'Emily Carter', Email: 'emily@rockymtn.com', Role: 'Controls Engineer', Company: 'Rocky Mountain Robotics', State: 'CO', 'Zip Code': '80202', 'Beckhoff Proficiency': 'Beginner', 'Course name': 'TwinCAT 3 Basics', Enrolled: '2026-05-02', Started: '2026-05-03', Completed: '', Score: '', 'Course progress': 20 },
    { 'Full name': 'Luis Romero', Email: 'luis@deserttech.com', Role: 'Service Technician', Company: 'Desert Tech', State: 'AZ', 'Zip Code': '85004', 'Beckhoff Proficiency': 'None', 'Course name': 'PLC Fundamentals', Enrolled: '2026-06-01', Started: '', Completed: '', Score: '', 'Course progress': 0 },
  ];
}
