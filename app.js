(() => {
  'use strict';

  const state = { rows: [], filtered: [], page: 1, pageSize: 50, selected: null, source: '' };
  const $ = (id) => document.getElementById(id);
  const fields = ['Phone Number', 'Saved Name', 'City', 'Gender', 'Person Type', 'Contact Category', 'Role/Tags', 'Video Seen', 'Video Sightings', 'Manual Review', 'Country/Region', 'International', 'Classification Confidence', 'Number Confidence', 'Classification Basis', 'OCR Evidence', 'Notes'];

  function text(value) { return value == null ? '' : String(value).trim(); }
  function escapeHtml(value) { return text(value).replace(/[&<>"']/g, (ch) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch])); }
  function normalizeKey(key) { return text(key).replace(/\s+/g, ' ').trim(); }
  function get(row, key) {
    if (Object.prototype.hasOwnProperty.call(row, key)) return text(row[key]);
    const wanted = normalizeKey(key).toLowerCase();
    const found = Object.keys(row).find((candidate) => normalizeKey(candidate).toLowerCase() === wanted);
    return found ? text(row[found]) : '';
  }
  function showToast(message) { const node = $('toast'); node.textContent = message; node.classList.add('show'); window.clearTimeout(showToast.timer); showToast.timer = window.setTimeout(() => node.classList.remove('show'), 2600); }
  function isReview(row) { return /^yes\b/i.test(get(row, 'Manual Review')); }
  function isVideo(row) { return /^yes$/i.test(get(row, 'Video Seen')) || Number(get(row, 'Video Sightings')) > 0; }
  function isSaved(row) { return Boolean(get(row, 'CSV Source') || get(row, 'Saved Name')); }

  function parseCsv(source) {
    const rows = []; let row = []; let cell = ''; let quoted = false;
    for (let i = 0; i < source.length; i += 1) {
      const ch = source[i]; const next = source[i + 1];
      if (ch === '"' && quoted && next === '"') { cell += '"'; i += 1; continue; }
      if (ch === '"') { quoted = !quoted; continue; }
      if (ch === ',' && !quoted) { row.push(cell); cell = ''; continue; }
      if ((ch === '\n' || ch === '\r') && !quoted) { if (ch === '\r' && next === '\n') i += 1; row.push(cell); if (row.some((value) => value !== '')) rows.push(row); row = []; cell = ''; continue; }
      cell += ch;
    }
    if (cell || row.length) { row.push(cell); if (row.some((value) => value !== '')) rows.push(row); }
    if (!rows.length) return [];
    const headers = rows.shift().map(normalizeKey);
    return rows.map((values) => headers.reduce((out, key, index) => { out[key] = values[index] || ''; return out; }, {}));
  }

  async function parseFile(file) {
    const lower = file.name.toLowerCase();
    if (lower.endsWith('.json')) {
      const parsed = JSON.parse(await file.text());
      if (Array.isArray(parsed)) return parsed;
      if (Array.isArray(parsed.rows)) return parsed.rows;
      if (Array.isArray(parsed.records)) return parsed.records;
      throw new Error('JSON must contain an array, rows, or records.');
    }
    if (lower.endsWith('.csv')) return parseCsv(await file.text());
    if (!window.XLSX) throw new Error('Workbook parser did not load. Use CSV or retry with an internet connection.');
    const workbook = window.XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: false });
    const sheetName = workbook.SheetNames.includes('Master Mapping') ? 'Master Mapping' : workbook.SheetNames[0];
    return window.XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: '' });
  }

  function setRows(rows, source) {
    state.rows = rows.map((row) => Object.keys(row).reduce((out, key) => { out[normalizeKey(key)] = row[key]; return out; }, {}));
    state.source = source;
    state.page = 1; state.selected = null;
    $('sourceStatus').textContent = `${source} - ${state.rows.length.toLocaleString()} rows`;
    $('emptyState').classList.add('hidden'); $('workspace').classList.remove('hidden');
    populateFilters(); updateKpis(); applyFilters();
  }

  function populateFilters() {
    const category = [...new Set(state.rows.map((row) => get(row, 'Contact Category')).filter(Boolean))].sort((a, b) => a.localeCompare(b));
    const city = [...new Set(state.rows.flatMap((row) => get(row, 'City').split(/[,;]\s*/).filter(Boolean)))].sort((a, b) => a.localeCompare(b));
    $('categoryFilter').innerHTML = '<option value="">All</option>' + category.map((value) => `<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`).join('');
    $('cityFilter').innerHTML = '<option value="">All</option>' + city.map((value) => `<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`).join('');
  }
  function updateKpis() {
    $('kpiTotal').textContent = state.rows.length.toLocaleString();
    $('kpiSaved').textContent = state.rows.filter(isSaved).length.toLocaleString();
    $('kpiVideo').textContent = state.rows.filter(isVideo).length.toLocaleString();
    $('kpiReview').textContent = state.rows.filter(isReview).length.toLocaleString();
    $('kpiInternational').textContent = state.rows.filter((row) => /^yes/i.test(get(row, 'International'))).length.toLocaleString();
  }
  function matches(row) {
    const query = text($('searchInput').value).toLowerCase();
    const category = $('categoryFilter').value; const city = $('cityFilter').value;
    if (query && !Object.values(row).some((value) => text(value).toLowerCase().includes(query))) return false;
    if (category && get(row, 'Contact Category') !== category) return false;
    if (city && !get(row, 'City').split(/[,;]\s*/).includes(city)) return false;
    if ($('reviewOnly').checked && !isReview(row)) return false;
    if ($('videoOnly').checked && !isVideo(row)) return false;
    return true;
  }
  function applyFilters() {
    state.filtered = state.rows.filter(matches); state.page = Math.min(state.page, Math.max(1, Math.ceil(state.filtered.length / state.pageSize))); renderTable();
  }
  function renderTable() {
    const start = (state.page - 1) * state.pageSize; const visible = state.filtered.slice(start, start + state.pageSize);
    $('resultCount').textContent = `${state.filtered.length.toLocaleString()} rows`; $('pageLabel').textContent = `Page ${state.filtered.length ? state.page : 0} / ${Math.max(1, Math.ceil(state.filtered.length / state.pageSize))}`;
    $('prevBtn').disabled = state.page <= 1; $('nextBtn').disabled = state.page >= Math.ceil(state.filtered.length / state.pageSize) || !state.filtered.length;
    $('rowsBody').innerHTML = visible.map((row, index) => {
      const rowId = state.rows.indexOf(row); const selected = row === state.selected ? ' selected' : '';
      const review = isReview(row); const video = isVideo(row);
      return `<tr class="${selected}" data-row-id="${rowId}"><td class="phone-cell">${escapeHtml(get(row, 'Phone Number') || get(row, 'Phone 1 - Value'))}</td><td>${escapeHtml(get(row, 'Saved Name') || get(row, 'First Name')) || '<span class="muted-cell">blank</span>'}</td><td>${escapeHtml(get(row, 'City')) || '<span class="muted-cell">blank</span>'}</td><td>${escapeHtml(get(row, 'Gender')) || '<span class="muted-cell">blank</span>'}</td><td>${escapeHtml(get(row, 'Person Type') || get(row, 'Contact Category')) || '<span class="muted-cell">blank</span>'}</td><td>${video ? `<span class="pill good">${escapeHtml(get(row, 'Video Sightings') || 'seen')}</span>` : '<span class="pill">No</span>'}</td><td>${review ? '<span class="pill warn">Review</span>' : '<span class="pill good">Clear</span>'}</td></tr>`;
    }).join('');
    $('rowsBody').querySelectorAll('tr').forEach((tr) => tr.addEventListener('click', () => { state.selected = state.rows[Number(tr.dataset.rowId)]; renderTable(); renderDetail(state.selected); }));
  }
  function renderDetail(row) {
    if (!row) { $('detailPanel').innerHTML = '<div class="detail-empty">Select a row to inspect its evidence.</div>'; return; }
    const title = get(row, 'Saved Name') || get(row, 'Phone Number') || 'Unlabelled row';
    const detailFields = fields.filter((key) => get(row, key));
    $('detailPanel').innerHTML = `<div class="detail-head"><h2>${escapeHtml(title)}</h2><div class="detail-phone">${escapeHtml(get(row, 'Phone Number'))}</div></div><div class="detail-body"><dl class="detail-grid">${detailFields.map((key) => { const value = get(row, key); const wide = ['OCR Evidence', 'Notes', 'Classification Basis'].includes(key); return `<div class="detail-field ${wide ? 'detail-wide' : ''}"><dt>${escapeHtml(key)}</dt><dd class="${wide ? 'evidence' : ''}">${escapeHtml(value)}</dd></div>`; }).join('')}</dl></div>`;
  }
  function csvCell(value) { const out = text(value).replace(/"/g, '""'); return /[",\n\r]/.test(out) ? `"${out}"` : out; }
  function exportCsv() {
    if (!state.filtered.length) { showToast('No rows match the current filters.'); return; }
    const headers = [...new Set(state.filtered.flatMap((row) => Object.keys(row)))];
    const csv = [headers.map(csvCell).join(','), ...state.filtered.map((row) => headers.map((key) => csvCell(row[key])).join(','))].join('\r\n');
    const link = document.createElement('a'); link.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' })); link.download = 'contact-mapping-filtered.csv'; link.click(); URL.revokeObjectURL(link.href); showToast(`Exported ${state.filtered.length.toLocaleString()} rows.`);
  }

  $('fileInput').addEventListener('change', async (event) => {
    const files = [...event.target.files]; if (!files.length) return;
    try { const parsed = []; for (const file of files) parsed.push(...await parseFile(file)); setRows(parsed, files.map((file) => file.name).join(', ')); showToast(`Loaded ${parsed.length.toLocaleString()} rows.`); }
    catch (error) { showToast(error.message || 'Could not read that file.'); }
    event.target.value = '';
  });
  ['searchInput', 'categoryFilter', 'cityFilter', 'reviewOnly', 'videoOnly'].forEach((id) => $(id).addEventListener('input', () => { state.page = 1; applyFilters(); }));
  $('prevBtn').addEventListener('click', () => { if (state.page > 1) { state.page -= 1; renderTable(); } });
  $('nextBtn').addEventListener('click', () => { if (state.page < Math.ceil(state.filtered.length / state.pageSize)) { state.page += 1; renderTable(); } });
  $('exportBtn').addEventListener('click', exportCsv);
  $('resetBtn').addEventListener('click', () => { state.rows = []; state.filtered = []; state.selected = null; $('workspace').classList.add('hidden'); $('emptyState').classList.remove('hidden'); $('sourceStatus').textContent = 'No file loaded'; $('fileInput').value = ''; showToast('Workspace cleared.'); });
})();
