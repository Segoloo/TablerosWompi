/**
 * ╔══════════════════════════════════════════════════════════════════╗
 * ║  rollos_detalles_tas.js — Tab "Detalles TAs"                   ║
 * ║                                                                  ║
 * ║  Fuente: rollos_detalles.json.gz (generado por prueba.py via   ║
 * ║          TablerosUpdater.py Módulo C)                           ║
 * ║                                                                  ║
 * ║  Columnas:                                                       ║
 * ║  COD. SITIO / FECHA PLANEADA INICIO / FECHA PLANEADA ENTREGA /  ║
 * ║  FECHA DE ENTREGA / CANTIDAD / CODIGO DE TAREA / GUIA / FO /   ║
 * ║  ESTADO / ESTADO TRANSPORTADORA / PROYECTO /                    ║
 * ║  DIAS INVENTARIO RESTANTES                                      ║
 * ║                                                                  ║
 * ║  Funcionalidades:                                                ║
 * ║  · Tabla paginada con ordenamiento por columna                  ║
 * ║  · Filtros individuales por cada columna                        ║
 * ║  · Filtros de texto con sugerencias (autocomplete)             ║
 * ║  · Filtro numérico de rango para DIAS INVENTARIO RESTANTES     ║
 * ║  · Export a Excel del resultado completo o filtrado             ║
 * ║  · Carga lazy: solo descarga rollos_detalles.json.gz cuando    ║
 * ║    el tab es visitado por primera vez                           ║
 * ╚══════════════════════════════════════════════════════════════════╝
 */
'use strict';

// ──────────────────────────────────────────────────────────────────
//  ESTADO DEL MÓDULO
// ──────────────────────────────────────────────────────────────────
const DT = {
  raw:         [],   // todas las filas del JSON
  filtered:    [],   // filas tras aplicar filtros
  page:        1,
  pageSize:    50,
  sortCol:     -1,
  sortDir:     1,
  loading:     false,
  loaded:      false,
  error:       null,
  columns: [
    'COD. SITIO',
    'FECHA PLANEADA INICIO',
    'FECHA PLANEADA ENTREGA',
    'FECHA DE ENTREGA',
    'CANTIDAD',
    'CODIGO DE TAREA',
    'GUIA',
    'FO',
    'ESTADO',
    'ESTADO TRANSPORTADORA',
    'PROYECTO',
    'DIAS INVENTARIO RESTANTES',
  ],
  // Estado de filtros: un objeto por columna { text, min, max }
  filters: {},
};

// URL del JSON (mismo origen que otros archivos del repo)
const DT_JSON_URL = new URL('rollos_detalles.json.gz?t=' + Date.now(), window.location.href).href;


// ──────────────────────────────────────────────────────────────────
//  CARGA DEL JSON (Web Worker inline para no bloquear el hilo)
// ──────────────────────────────────────────────────────────────────
const _DT_WORKER_SRC = `
self.onmessage = async function(e) {
  const url = e.data;
  try {
    const res = await fetch(url);
    if (!res.ok) throw new Error('HTTP ' + res.status);
    const buf = await res.arrayBuffer();
    const ds  = new DecompressionStream('gzip');
    const writer = ds.writable.getWriter();
    writer.write(new Uint8Array(buf));
    writer.close();
    const reader = ds.readable.getReader();
    const chunks = [];
    let totalLen = 0;
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      chunks.push(value);
      totalLen += value.length;
    }
    const merged = new Uint8Array(totalLen);
    let offset = 0;
    for (const c of chunks) { merged.set(c, offset); offset += c.length; }
    const text = new TextDecoder().decode(merged);
    const payload = JSON.parse(text);
    self.postMessage({ ok: true, payload });
  } catch(err) {
    self.postMessage({ ok: false, error: err.message });
  }
};
`;

function _dtLoadData() {
  return new Promise((resolve, reject) => {
    const blob   = new Blob([_DT_WORKER_SRC], { type: 'application/javascript' });
    const url    = URL.createObjectURL(blob);
    const worker = new Worker(url);
    worker.onmessage = function(e) {
      URL.revokeObjectURL(url);
      worker.terminate();
      if (e.data.ok) {
        resolve(e.data.payload);
      } else {
        reject(new Error(e.data.error));
      }
    };
    worker.onerror = function(err) {
      URL.revokeObjectURL(url);
      worker.terminate();
      reject(err);
    };
    worker.postMessage(DT_JSON_URL);
  });
}

// ──────────────────────────────────────────────────────────────────
//  HELPERS
// ──────────────────────────────────────────────────────────────────
function _dtNorm(v) {
  if (v == null) return '';
  return String(v).trim().toUpperCase();
}

function _dtFmt(v) {
  if (v == null || v === '') return '—';
  return String(v);
}

function _dtApplyFilters() {
  let rows = DT.raw.slice();
  for (const col of DT.columns) {
    const f = DT.filters[col];
    if (!f) continue;
    if (col === 'DIAS INVENTARIO RESTANTES') {
      const min = f.min !== '' && f.min != null ? parseFloat(f.min) : null;
      const max = f.max !== '' && f.max != null ? parseFloat(f.max) : null;
      if (min !== null || max !== null) {
        rows = rows.filter(r => {
          const v = parseFloat(r[col]);
          if (isNaN(v)) return false;
          if (min !== null && v < min) return false;
          if (max !== null && v > max) return false;
          return true;
        });
      }
    } else if (f.text) {
      const needle = f.text.trim().toUpperCase();
      if (needle) {
        rows = rows.filter(r => _dtNorm(r[col]).includes(needle));
      }
    }
  }
  DT.filtered = rows;
  DT.page = 1;
}

function _dtSort(rows) {
  if (DT.sortCol < 0) return rows;
  const col = DT.columns[DT.sortCol];
  const dir = DT.sortDir;
  const isNum = col === 'CANTIDAD' || col === 'DIAS INVENTARIO RESTANTES';
  return rows.slice().sort((a, b) => {
    let va = a[col], vb = b[col];
    if (isNum) {
      va = parseFloat(va);
      vb = parseFloat(vb);
      if (isNaN(va)) va = -Infinity;
      if (isNaN(vb)) vb = -Infinity;
      return (va - vb) * dir;
    }
    va = _dtNorm(va);
    vb = _dtNorm(vb);
    return va < vb ? -dir : va > vb ? dir : 0;
  });
}

function _dtPagedRows() {
  const sorted = _dtSort(DT.filtered);
  const start  = (DT.page - 1) * DT.pageSize;
  return sorted.slice(start, start + DT.pageSize);
}

function _dtTotalPages() {
  return Math.max(1, Math.ceil(DT.filtered.length / DT.pageSize));
}

// Valores únicos de una columna para autocomplete (máx 300)
function _dtUniqueVals(col) {
  const set = new Set();
  for (const r of DT.raw) {
    const v = r[col];
    if (v != null && v !== '') set.add(String(v));
    if (set.size >= 300) break;
  }
  return Array.from(set).sort();
}

// ──────────────────────────────────────────────────────────────────
//  INYECTAR ESTILOS
// ──────────────────────────────────────────────────────────────────
function _dtInjectStyles() {
  if (document.getElementById('dt-styles')) return;
  const s = document.createElement('style');
  s.id = 'dt-styles';
  s.textContent = `
    /* ── Detalles TAs ─────────────────────────────── */
    #dt-root {
      font-family: 'Outfit', sans-serif;
      color: #e2e8f0;
    }
    #dt-header-bar {
      display: flex;
      align-items: center;
      gap: 12px;
      padding: 0 4px 18px;
      flex-wrap: wrap;
    }
    #dt-title {
      font-family: 'Syne', sans-serif;
      font-size: 18px;
      font-weight: 700;
      color: #f1f5f9;
      letter-spacing: .2px;
      flex: 1;
    }
    #dt-count-badge {
      font-size: 11px;
      color: #475569;
      font-family: 'JetBrains Mono', monospace;
      white-space: nowrap;
    }
    #dt-export-btn {
      padding: 7px 18px;
      background: rgba(176,242,174,.1);
      border: 1px solid rgba(176,242,174,.3);
      border-radius: 20px;
      color: #B0F2AE;
      font-size: 12px;
      font-family: 'Outfit', sans-serif;
      font-weight: 600;
      cursor: pointer;
      transition: background .2s, border-color .2s;
      white-space: nowrap;
    }
    #dt-export-btn:hover { background: rgba(176,242,174,.2); border-color: rgba(176,242,174,.55); }
    #dt-clear-btn {
      padding: 7px 14px;
      background: rgba(255,255,255,.05);
      border: 1px solid rgba(255,255,255,.1);
      border-radius: 20px;
      color: #94a3b8;
      font-size: 12px;
      font-family: 'Outfit', sans-serif;
      cursor: pointer;
      transition: background .2s;
      white-space: nowrap;
    }
    #dt-clear-btn:hover { background: rgba(255,92,92,.1); border-color: rgba(255,92,92,.3); color: #FF5C5C; }

    /* tabla wrapper */
    #dt-table-wrap {
      width: 100%;
      overflow-x: auto;
      border-radius: 14px;
      border: 1px solid rgba(255,255,255,.07);
      background: rgba(255,255,255,.015);
    }
    #dt-table-wrap::-webkit-scrollbar { height: 5px; }
    #dt-table-wrap::-webkit-scrollbar-thumb { background: rgba(176,242,174,.25); border-radius: 3px; }

    #dt-table {
      width: 100%;
      border-collapse: collapse;
      min-width: 1300px;
    }
    #dt-table thead th {
      background: rgba(255,255,255,.04);
      font-family: 'Outfit', sans-serif;
      font-size: 11px;
      font-weight: 700;
      color: #94a3b8;
      text-transform: uppercase;
      letter-spacing: .4px;
      padding: 0;
      border-bottom: 1px solid rgba(255,255,255,.08);
      white-space: nowrap;
      position: sticky;
      top: 0;
      z-index: 2;
    }
    /* header cell inner */
    .dt-th-inner {
      display: flex;
      flex-direction: column;
      gap: 4px;
      padding: 10px 10px 8px;
    }
    .dt-th-label {
      display: flex;
      align-items: center;
      gap: 5px;
      cursor: pointer;
      user-select: none;
      white-space: nowrap;
    }
    .dt-th-label:hover { color: #B0F2AE; }
    .dt-sort-icon { font-size: 10px; color: #475569; }
    .dt-sort-asc .dt-sort-icon,
    .dt-sort-desc .dt-sort-icon { color: #B0F2AE; }

    /* filtro input */
    .dt-filter-wrap { position: relative; }
    .dt-filter-input {
      width: 100%;
      background: rgba(255,255,255,.05);
      border: 1px solid rgba(255,255,255,.1);
      border-radius: 6px;
      padding: 4px 8px;
      font-size: 11px;
      color: #e2e8f0;
      font-family: 'Outfit', sans-serif;
      outline: none;
      box-sizing: border-box;
      min-width: 80px;
      transition: border-color .2s;
    }
    .dt-filter-input:focus { border-color: rgba(176,242,174,.5); background: rgba(176,242,174,.04); }
    .dt-filter-input.active { border-color: rgba(176,242,174,.7); color: #B0F2AE; background: rgba(176,242,174,.07); }
    .dt-filter-range-wrap { display: flex; gap: 4px; }
    .dt-filter-range-wrap .dt-filter-input { min-width: 52px; }

    /* autocomplete dropdown */
    .dt-ac-list {
      position: absolute;
      top: 100%;
      left: 0;
      right: 0;
      z-index: 9999;
      background: #1c1b18;
      border: 1px solid rgba(176,242,174,.3);
      border-radius: 8px;
      margin-top: 3px;
      list-style: none;
      padding: 0;
      max-height: 200px;
      overflow-y: auto;
      box-shadow: 0 8px 28px rgba(0,0,0,.7);
      scrollbar-width: thin;
      scrollbar-color: rgba(176,242,174,.3) transparent;
    }
    .dt-ac-list::-webkit-scrollbar { width: 4px; }
    .dt-ac-list::-webkit-scrollbar-thumb { background: rgba(176,242,174,.3); border-radius: 2px; }
    .dt-ac-list li {
      padding: 6px 12px;
      font-size: 11px;
      font-family: 'Outfit', sans-serif;
      color: #e2e8f0;
      cursor: pointer;
      border-bottom: 1px solid rgba(255,255,255,.04);
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }
    .dt-ac-list li:last-child { border-bottom: none; }
    .dt-ac-list li:hover, .dt-ac-list li.dt-ac-active { background: rgba(176,242,174,.12); color: #B0F2AE; }

    /* body rows */
    #dt-table tbody tr {
      border-bottom: 1px solid rgba(255,255,255,.04);
      transition: background .15s;
    }
    #dt-table tbody tr:hover { background: rgba(176,242,174,.04); }
    #dt-table tbody td {
      padding: 8px 10px;
      font-size: 12px;
      color: #cbd5e1;
      vertical-align: middle;
      white-space: nowrap;
      max-width: 220px;
      overflow: hidden;
      text-overflow: ellipsis;
    }
    #dt-table tbody td.dt-null { color: #334155; }

    /* estado badges */
    .dt-badge {
      display: inline-block;
      padding: 2px 9px;
      border-radius: 20px;
      font-size: 10px;
      font-weight: 700;
      letter-spacing: .3px;
      white-space: nowrap;
    }
    .dt-badge-entregado     { background: rgba(176,242,174,.15); color: #B0F2AE; border: 1px solid rgba(176,242,174,.3); }
    .dt-badge-transito      { background: rgba(153,209,252,.12); color: #99D1FC; border: 1px solid rgba(153,209,252,.3); }
    .dt-badge-alistamiento  { background: rgba(223,255,97,.1);   color: #DFFF61; border: 1px solid rgba(223,255,97,.3); }
    .dt-badge-devolucion    { background: rgba(255,92,92,.1);    color: #FF5C5C; border: 1px solid rgba(255,92,92,.3); }
    .dt-badge-asignada      { background: rgba(255,192,77,.1);   color: #FFC04D; border: 1px solid rgba(255,192,77,.3); }
    .dt-badge-completada    { background: rgba(0,130,90,.15);    color: #00C87A; border: 1px solid rgba(0,200,122,.3); }
    .dt-badge-default       { background: rgba(255,255,255,.06); color: #94a3b8; border: 1px solid rgba(255,255,255,.1); }

    /* días: semáforo de color */
    .dt-dias-ok       { color: #B0F2AE; font-weight: 700; }
    .dt-dias-warn     { color: #DFFF61; font-weight: 600; }
    .dt-dias-alert    { color: #FFC04D; font-weight: 600; }
    .dt-dias-critical { color: #FF5C5C; font-weight: 700; }

    /* paginación */
    #dt-pagination {
      display: flex;
      align-items: center;
      gap: 6px;
      padding: 12px 16px;
      border-top: 1px solid rgba(255,255,255,.06);
      flex-wrap: wrap;
    }
    .dt-page-btn {
      padding: 4px 10px;
      background: rgba(255,255,255,.05);
      border: 1px solid rgba(255,255,255,.1);
      border-radius: 6px;
      color: #94a3b8;
      font-size: 11px;
      font-family: 'Outfit', sans-serif;
      cursor: pointer;
      transition: background .2s;
    }
    .dt-page-btn:hover:not(:disabled) { background: rgba(176,242,174,.1); border-color: rgba(176,242,174,.3); color: #B0F2AE; }
    .dt-page-btn.active { background: rgba(176,242,174,.15); border-color: rgba(176,242,174,.4); color: #B0F2AE; font-weight: 700; }
    .dt-page-btn:disabled { opacity: .35; cursor: default; }
    #dt-page-info { font-size: 11px; color: #475569; font-family: 'JetBrains Mono', monospace; margin-left: auto; }

    /* loading / error */
    #dt-loading, #dt-error-msg {
      display: flex;
      align-items: center;
      justify-content: center;
      min-height: 260px;
      font-size: 14px;
      font-family: 'Outfit', sans-serif;
      gap: 14px;
      flex-direction: column;
    }
    @keyframes dt-spin { to { transform: rotate(360deg); } }
    .dt-spinner {
      width: 38px; height: 38px; border-radius: 50%;
      border: 3px solid rgba(176,242,174,.15);
      border-top-color: #B0F2AE;
      animation: dt-spin .7s linear infinite;
    }
  `;
  document.head.appendChild(s);
}

// ──────────────────────────────────────────────────────────────────
//  RENDER PRINCIPAL
// ──────────────────────────────────────────────────────────────────
function _dtRenderRoot(mode) {
  const root = document.getElementById('dt-root');
  if (!root) return;

  if (mode === 'loading') {
    root.innerHTML = `
      <div id="dt-loading">
        <div class="dt-spinner"></div>
        <div style="color:#94a3b8;">Cargando <code style="color:#DFFF61;font-size:12px;">rollos_detalles.json.gz</code>…</div>
      </div>`;
    return;
  }
  if (mode === 'error') {
    root.innerHTML = `
      <div id="dt-error-msg">
        <div style="font-size:28px;">⚠️</div>
        <div style="color:#FF5C5C;font-weight:700;">No se pudo cargar rollos_detalles.json.gz</div>
        <div style="color:#64748b;font-size:12px;max-width:420px;text-align:center;line-height:1.6;">${DT.error || 'Error desconocido'}</div>
        <button onclick="window.renderDetallesTas(true)" style="
          margin-top:8px;padding:7px 20px;background:rgba(255,92,92,.1);
          border:1px solid rgba(255,92,92,.3);border-radius:20px;
          color:#FF5C5C;font-size:12px;cursor:pointer;">
          Reintentar
        </button>
      </div>`;
    return;
  }

  // Modo normal: tabla completa
  _dtApplyFilters();
  const rows  = _dtPagedRows();
  const total = DT.filtered.length;
  const tp    = _dtTotalPages();
  const cols  = DT.columns;

  root.innerHTML = `
    <div id="dt-header-bar">
      <div id="dt-title">📑 Detalles TAs</div>
      <div id="dt-count-badge">${total.toLocaleString('es-CO')} registros encontrados · ${DT.raw.length.toLocaleString('es-CO')} total</div>
      <button id="dt-clear-btn" onclick="window._dtClearFilters()">✕ Limpiar filtros</button>
      <button id="dt-export-btn" onclick="window._dtExport()">⬇ Exportar Excel</button>
    </div>
    <div id="dt-table-wrap">
      <table id="dt-table">
        <thead>
          <tr>${cols.map((col, i) => _dtThHTML(col, i)).join('')}</tr>
        </thead>
        <tbody>
          ${rows.length === 0
            ? `<tr><td colspan="${cols.length}" style="text-align:center;padding:40px;color:#475569;">Sin resultados con los filtros actuales.</td></tr>`
            : rows.map(_dtRowHTML).join('')}
        </tbody>
      </table>
    </div>
    <div id="dt-pagination">${_dtPaginationHTML(tp)}</div>
  `;

  // Bind autocomplete
  cols.forEach((col, i) => {
    if (col !== 'CANTIDAD' && col !== 'DIAS INVENTARIO RESTANTES') {
      _dtBindAutocomplete(col, i);
    }
  });
}

// ── Header cell ───────────────────────────────────────────────────
function _dtThHTML(col, i) {
  const sortClass = DT.sortCol === i
    ? (DT.sortDir === 1 ? 'dt-sort-asc' : 'dt-sort-desc')
    : '';
  const sortIcon = DT.sortCol === i
    ? (DT.sortDir === 1 ? '▲' : '▼')
    : '⇅';

  let filterHtml = '';
  if (col === 'DIAS INVENTARIO RESTANTES') {
    const f    = DT.filters[col] || {};
    const fMin = f.min != null ? f.min : '';
    const fMax = f.max != null ? f.max : '';
    filterHtml = `
      <div class="dt-filter-wrap dt-filter-range-wrap">
        <input class="dt-filter-input${fMin !== '' ? ' active' : ''}" type="number"
          placeholder="Min" value="${fMin}"
          oninput="window._dtSetRangeFilter('${col}','min',this.value)"
          style="min-width:52px;">
        <input class="dt-filter-input${fMax !== '' ? ' active' : ''}" type="number"
          placeholder="Max" value="${fMax}"
          oninput="window._dtSetRangeFilter('${col}','max',this.value)"
          style="min-width:52px;">
      </div>`;
  } else {
    const fVal = (DT.filters[col] || {}).text || '';
    filterHtml = `
      <div class="dt-filter-wrap" id="dt-fw-${i}">
        <input class="dt-filter-input${fVal ? ' active' : ''}" type="text"
          placeholder="Filtrar…" value="${fVal.replace(/"/g, '&quot;')}"
          autocomplete="off"
          id="dt-fi-${i}"
          oninput="window._dtSetTextFilter('${col}',this.value,${i})"
          onfocus="window._dtShowAC(${i})"
          onblur="setTimeout(()=>window._dtHideAC(${i}),180)">
        <ul class="dt-ac-list" id="dt-ac-${i}" style="display:none;"></ul>
      </div>`;
  }

  return `
    <th>
      <div class="dt-th-inner ${sortClass}">
        <div class="dt-th-label" onclick="window._dtToggleSort(${i})">
          <span>${col}</span>
          <span class="dt-sort-icon">${sortIcon}</span>
        </div>
        ${filterHtml}
      </div>
    </th>`;
}

// ── Row ───────────────────────────────────────────────────────────
function _dtRowHTML(row) {
  return `<tr>${DT.columns.map(col => {
    const v = row[col];
    if (col === 'ESTADO') {
      return `<td>${_dtBadgeEstado(v)}</td>`;
    }
    if (col === 'DIAS INVENTARIO RESTANTES') {
      return `<td>${_dtDiasHTML(v)}</td>`;
    }
    if (v == null || v === '') {
      return `<td class="dt-null">—</td>`;
    }
    return `<td title="${String(v).replace(/"/g, '&quot;')}">${String(v)}</td>`;
  }).join('')}</tr>`;
}

function _dtBadgeEstado(v) {
  if (!v) return '<span class="dt-badge dt-badge-default">—</span>';
  const u = String(v).toUpperCase();
  let cls = 'dt-badge-default';
  if (u === 'ENTREGADO')                             cls = 'dt-badge-entregado';
  else if (u === 'EN TRANSITO')                      cls = 'dt-badge-transito';
  else if (u === 'EN ALISTAMIENTO')                  cls = 'dt-badge-alistamiento';
  else if (u.includes('DEVOL') || u === 'DEVOLUCIÓN') cls = 'dt-badge-devolucion';
  else if (u === 'ASIGNADA')                         cls = 'dt-badge-asignada';
  else if (u === 'COMPLETADA')                       cls = 'dt-badge-completada';
  return `<span class="dt-badge ${cls}">${v}</span>`;
}

function _dtDiasHTML(v) {
  const n = parseFloat(v);
  if (isNaN(n) || v == null || v === '') return '<span class="dt-null">—</span>';
  let cls = 'dt-dias-ok';
  if (n < 0)        cls = 'dt-dias-critical';
  else if (n <= 15) cls = 'dt-dias-alert';
  else if (n <= 30) cls = 'dt-dias-warn';
  return `<span class="${cls}">${Math.round(n)}</span>`;
}

// ── Paginación ────────────────────────────────────────────────────
function _dtPaginationHTML(tp) {
  if (tp <= 1) return '';
  const p = DT.page;
  let html = `<button class="dt-page-btn" onclick="window._dtGoPage(1)" ${p===1?'disabled':''}>«</button>`;
  html    += `<button class="dt-page-btn" onclick="window._dtGoPage(${p-1})" ${p===1?'disabled':''}>‹</button>`;
  const start = Math.max(1, p - 2);
  const end   = Math.min(tp, p + 2);
  if (start > 1) html += `<button class="dt-page-btn" onclick="window._dtGoPage(1)">1</button>`;
  if (start > 2) html += `<span style="color:#475569;font-size:11px;padding:0 2px;">…</span>`;
  for (let i = start; i <= end; i++) {
    html += `<button class="dt-page-btn${i===p?' active':''}" onclick="window._dtGoPage(${i})">${i}</button>`;
  }
  if (end < tp - 1) html += `<span style="color:#475569;font-size:11px;padding:0 2px;">…</span>`;
  if (end < tp)     html += `<button class="dt-page-btn" onclick="window._dtGoPage(${tp})">${tp}</button>`;
  html    += `<button class="dt-page-btn" onclick="window._dtGoPage(${p+1})" ${p===tp?'disabled':''}>›</button>`;
  html    += `<button class="dt-page-btn" onclick="window._dtGoPage(${tp})" ${p===tp?'disabled':''}>»</button>`;
  html    += `<span id="dt-page-info">Pág ${p} / ${tp} · ${DT.filtered.length.toLocaleString('es-CO')} filas</span>`;
  return html;
}

// ──────────────────────────────────────────────────────────────────
//  AUTOCOMPLETE
// ──────────────────────────────────────────────────────────────────
let _dtAcCache = {};

function _dtBindAutocomplete(col, i) {
  // Cache de valores únicos al primer uso de cada columna
  if (!_dtAcCache[col]) {
    _dtAcCache[col] = _dtUniqueVals(col);
  }
}

window._dtShowAC = function(i) {
  const col  = DT.columns[i];
  const inp  = document.getElementById(`dt-fi-${i}`);
  const list = document.getElementById(`dt-ac-${i}`);
  if (!inp || !list) return;
  if (!_dtAcCache[col]) _dtAcCache[col] = _dtUniqueVals(col);
  _dtRenderAC(i, col, inp.value);
  list.style.display = 'block';
};

window._dtHideAC = function(i) {
  const list = document.getElementById(`dt-ac-${i}`);
  if (list) list.style.display = 'none';
};

function _dtRenderAC(i, col, query) {
  const list = document.getElementById(`dt-ac-${i}`);
  if (!list) return;
  const needle = query.trim().toUpperCase();
  const all    = _dtAcCache[col] || [];
  const matches = needle
    ? all.filter(v => String(v).toUpperCase().includes(needle)).slice(0, 60)
    : all.slice(0, 60);
  list.innerHTML = matches.map(v =>
    `<li onmousedown="window._dtPickAC(${i},'${String(v).replace(/'/g,"\\'")}')">
       ${String(v)}
     </li>`
  ).join('') || `<li style="color:#475569;pointer-events:none;">Sin resultados</li>`;
}

window._dtPickAC = function(i, val) {
  const col = DT.columns[i];
  if (!DT.filters[col]) DT.filters[col] = {};
  DT.filters[col].text = val;
  _dtApplyFilters();
  _dtRenderRoot('table');
};

// ──────────────────────────────────────────────────────────────────
//  CALLBACKS DESDE EL DOM
// ──────────────────────────────────────────────────────────────────
window._dtSetTextFilter = function(col, val, i) {
  if (!DT.filters[col]) DT.filters[col] = {};
  DT.filters[col].text = val;
  _dtApplyFilters();
  // Actualizar solo tabla y paginación sin re-render completo de cabeceras
  const tbody = document.querySelector('#dt-table tbody');
  const pg    = document.getElementById('dt-pagination');
  const cnt   = document.getElementById('dt-count-badge');
  const rows  = _dtPagedRows();
  if (tbody) {
    tbody.innerHTML = rows.length === 0
      ? `<tr><td colspan="${DT.columns.length}" style="text-align:center;padding:40px;color:#475569;">Sin resultados.</td></tr>`
      : rows.map(_dtRowHTML).join('');
  }
  if (pg)  pg.innerHTML  = _dtPaginationHTML(_dtTotalPages());
  if (cnt) cnt.textContent = `${DT.filtered.length.toLocaleString('es-CO')} registros encontrados · ${DT.raw.length.toLocaleString('es-CO')} total`;
  // Re-render AC en el campo activo
  const inp = document.getElementById(`dt-fi-${i}`);
  if (inp) { inp.className = 'dt-filter-input' + (val ? ' active' : ''); }
  if (document.getElementById(`dt-ac-${i}`)) _dtRenderAC(i, col, val);
};

window._dtSetRangeFilter = function(col, side, val) {
  if (!DT.filters[col]) DT.filters[col] = {};
  DT.filters[col][side] = val;
  _dtApplyFilters();
  const tbody = document.querySelector('#dt-table tbody');
  const pg    = document.getElementById('dt-pagination');
  const cnt   = document.getElementById('dt-count-badge');
  const rows  = _dtPagedRows();
  if (tbody) tbody.innerHTML = rows.length === 0
    ? `<tr><td colspan="${DT.columns.length}" style="text-align:center;padding:40px;color:#475569;">Sin resultados.</td></tr>`
    : rows.map(_dtRowHTML).join('');
  if (pg)  pg.innerHTML  = _dtPaginationHTML(_dtTotalPages());
  if (cnt) cnt.textContent = `${DT.filtered.length.toLocaleString('es-CO')} registros encontrados · ${DT.raw.length.toLocaleString('es-CO')} total`;
};

window._dtToggleSort = function(i) {
  if (DT.sortCol === i) {
    DT.sortDir = DT.sortDir === 1 ? -1 : 1;
  } else {
    DT.sortCol = i;
    DT.sortDir = 1;
  }
  _dtRenderRoot('table');
};

window._dtGoPage = function(p) {
  const tp = _dtTotalPages();
  DT.page = Math.max(1, Math.min(tp, p));
  const tbody = document.querySelector('#dt-table tbody');
  const pg    = document.getElementById('dt-pagination');
  const rows  = _dtPagedRows();
  if (tbody) tbody.innerHTML = rows.length === 0
    ? `<tr><td colspan="${DT.columns.length}" style="text-align:center;padding:40px;color:#475569;">Sin resultados.</td></tr>`
    : rows.map(_dtRowHTML).join('');
  if (pg) pg.innerHTML = _dtPaginationHTML(tp);
};

window._dtClearFilters = function() {
  DT.filters = {};
  DT.sortCol = -1;
  DT.sortDir = 1;
  DT.page    = 1;
  _dtRenderRoot('table');
};

// ──────────────────────────────────────────────────────────────────
//  EXPORT A EXCEL
// ──────────────────────────────────────────────────────────────────
window._dtExport = function() {
  if (!window.XLSX) {
    alert('La librería XLSX no está disponible. Recarga la página.');
    return;
  }
  const rows  = DT.filtered.length > 0 ? DT.filtered : DT.raw;
  const data  = [DT.columns, ...rows.map(r => DT.columns.map(c => {
    const v = r[c];
    if (v == null) return '';
    const n = parseFloat(v);
    return (c === 'CANTIDAD' || c === 'DIAS INVENTARIO RESTANTES') && !isNaN(n) ? n : String(v);
  }))];
  const ws    = XLSX.utils.aoa_to_sheet(data);
  // Anchos de columna
  ws['!cols'] = DT.columns.map(c => ({ wch: Math.max(c.length + 2, 16) }));
  const wb    = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Detalles TAs');
  const fname = `detalles_tas_${new Date().toISOString().slice(0,10)}.xlsx`;
  XLSX.writeFile(wb, fname);
};

// ──────────────────────────────────────────────────────────────────
//  PUNTO DE ENTRADA PÚBLICO
// ──────────────────────────────────────────────────────────────────
window.renderDetallesTas = async function(forceReload = false) {
  _dtInjectStyles();

  // Ya cargado y no forzar reload → solo re-render
  if (DT.loaded && !forceReload) {
    _dtRenderRoot('table');
    return;
  }

  // Cargando en paralelo → esperar
  if (DT.loading && !forceReload) {
    _dtRenderRoot('loading');
    return;
  }

  // Reintentar desde cero
  if (forceReload) {
    DT.loaded  = false;
    DT.error   = null;
    DT.raw     = [];
    DT.filtered = [];
    DT.filters = {};
  }

  DT.loading = true;
  DT.error   = null;
  _dtRenderRoot('loading');

  try {
    const payload = await _dtLoadData();
    DT.raw     = Array.isArray(payload.filas) ? payload.filas : [];
    DT.loaded  = true;
    DT.loading = false;
    _dtApplyFilters();

    // Actualizar dot de loading overlay si sigue visible
    const dot = document.getElementById('dl-dot-detalles-tas');
    const sub = document.getElementById('dl-sub-detalles-tas');
    if (dot) dot.className = 'dl-item-dot done';
    if (sub) sub.textContent = DT.raw.length.toLocaleString('es-CO') + ' filas ✓';

    _dtRenderRoot('table');
  } catch (err) {
    console.error('[DetallesTas] Error al cargar:', err);
    DT.loading = false;
    DT.error   = err.message || String(err);
    _dtRenderRoot('error');
    // Marcar dot como error
    const dot = document.getElementById('dl-dot-detalles-tas');
    const sub = document.getElementById('dl-sub-detalles-tas');
    if (dot) dot.className = 'dl-item-dot error';
    if (sub) sub.textContent = 'Error al cargar ✗';
  }
};