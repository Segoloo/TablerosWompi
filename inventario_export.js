/**
 * ╔══════════════════════════════════════════════════════════════════╗
 * ║  inventario_export.js — Tab "Export Completo de Stock"          ║
 * ║  Dashboard Inventario Wompi v2                                   ║
 * ║  Exporta stock_wompi_filtrado.json.gz → Excel completo          ║
 * ╚══════════════════════════════════════════════════════════════════╝
 */
'use strict';

// ══════════════════════════════════════════════════════════════════
//  INIT — Monta el panel cuando se activa el tab
// ══════════════════════════════════════════════════════════════════
window.initExportStock = async function () {
  const panel = document.getElementById('panel-inv-export');
  if (!panel) return;
  if (panel.dataset.mounted === '1') return; // ya montado

  panel.dataset.mounted = '1';
  panel.innerHTML = _buildExportUI();
  _attachExportEvents();
  await _loadPreviewStats();
};

// ══════════════════════════════════════════════════════════════════
//  UI — HTML del panel
// ══════════════════════════════════════════════════════════════════
function _buildExportUI() {
  return `
  <div style="max-width:960px;margin:0 auto;padding:0 4px 60px;">

    <!-- ── Header ── -->
    <div style="margin-bottom:32px;">
      <div style="display:flex;align-items:center;gap:14px;margin-bottom:10px;">
        <div style="width:46px;height:46px;background:rgba(34,197,94,.12);border-radius:14px;
                    display:flex;align-items:center;justify-content:center;font-size:24px;flex-shrink:0;">📥</div>
        <div>
          <div style="font-family:'Syne',sans-serif;font-size:22px;font-weight:800;
                      color:#f1f5f9;letter-spacing:-.3px;">Export Completo de Stock</div>
          <div style="font-size:13px;color:#64748b;margin-top:2px;font-family:'Outfit',sans-serif;">
            Exporta <code style="color:#DFFF61;background:rgba(255,255,255,.06);
            padding:1px 6px;border-radius:4px;">stock_wompi_filtrado.json.gz</code>
            completo a Excel — todas las filas, sin filtros.
          </div>
        </div>
      </div>
    </div>

    <!-- ── Stats rápidas ── -->
    <div id="exp-stats-row" style="display:grid;grid-template-columns:repeat(auto-fill,minmax(160px,1fr));
         gap:12px;margin-bottom:28px;">
      <div class="exp-stat-card" id="exp-stat-total">
        <div class="exp-stat-val" id="exp-stat-total-val">—</div>
        <div class="exp-stat-lbl">Total filas</div>
      </div>
      <div class="exp-stat-card" id="exp-stat-refs">
        <div class="exp-stat-val" id="exp-stat-refs-val">—</div>
        <div class="exp-stat-lbl">Referencias únicas</div>
      </div>
      <div class="exp-stat-card" id="exp-stat-bodegas">
        <div class="exp-stat-val" id="exp-stat-bodegas-val">—</div>
        <div class="exp-stat-lbl">Ubicaciones</div>
      </div>
      <div class="exp-stat-card" id="exp-stat-uds">
        <div class="exp-stat-val" id="exp-stat-uds-val">—</div>
        <div class="exp-stat-lbl">Total unidades</div>
      </div>
    </div>

    <!-- ── Opciones de export ── -->
    <div style="background:rgba(255,255,255,.025);border:1px solid rgba(255,255,255,.07);
                border-radius:20px;padding:28px 32px;margin-bottom:20px;">

      <div style="font-family:'Syne',sans-serif;font-size:13px;font-weight:700;
                  color:#f1f5f9;margin-bottom:20px;letter-spacing:.2px;">⚙️ Opciones de exportación</div>

      <div style="display:grid;grid-template-columns:1fr 1fr;gap:20px;margin-bottom:20px;">

        <!-- Columnas a incluir -->
        <div>
          <div style="font-size:12px;color:#94a3b8;font-family:'Outfit',sans-serif;
                      font-weight:600;margin-bottom:10px;text-transform:uppercase;letter-spacing:.5px;">
            Columnas a incluir
          </div>
          <div id="exp-cols-list" style="display:flex;flex-direction:column;gap:7px;">
            <!-- generado por JS -->
          </div>
        </div>

        <!-- Filtro rápido por categoría -->
        <div>
          <div style="font-size:12px;color:#94a3b8;font-family:'Outfit',sans-serif;
                      font-weight:600;margin-bottom:10px;text-transform:uppercase;letter-spacing:.5px;">
            Filtro rápido (opcional)
          </div>

          <div style="margin-bottom:12px;">
            <div style="font-size:11px;color:#64748b;margin-bottom:5px;font-family:'Outfit',sans-serif;">Categoría</div>
            <select id="exp-filter-cat" style="width:100%;background:#0f172a;border:1px solid rgba(255,255,255,.12);
              border-radius:8px;color:#e2e8f0;font-size:12px;font-family:'Outfit',sans-serif;
              padding:8px 10px;outline:none;">
              <option value="">Todas las categorías</option>
              <option value="Rollos">Rollos</option>
              <option value="Pin pad">Pin pad</option>
              <option value="Forros">Forros</option>
              <option value="Accesorios">Accesorios</option>
              <option value="SIM">SIM</option>
              <option value="Datáfonos">Datáfonos</option>
              <option value="KIT POP VP">KIT POP VP</option>
            </select>
          </div>

          <div style="margin-bottom:12px;">
            <div style="font-size:11px;color:#64748b;margin-bottom:5px;font-family:'Outfit',sans-serif;">Negocio</div>
            <select id="exp-filter-neg" style="width:100%;background:#0f172a;border:1px solid rgba(255,255,255,.12);
              border-radius:8px;color:#e2e8f0;font-size:12px;font-family:'Outfit',sans-serif;
              padding:8px 10px;outline:none;">
              <option value="">Todos (VP + CB)</option>
              <option value="VP">Solo VP</option>
              <option value="CB">Solo CB</option>
            </select>
          </div>

          <div>
            <div style="font-size:11px;color:#64748b;margin-bottom:5px;font-family:'Outfit',sans-serif;">Tipo de ubicación</div>
            <select id="exp-filter-tipo" style="width:100%;background:#0f172a;border:1px solid rgba(255,255,255,.12);
              border-radius:8px;color:#e2e8f0;font-size:12px;font-family:'Outfit',sans-serif;
              padding:8px 10px;outline:none;">
              <option value="">Todos los tipos</option>
              <option value="Warehouse">Bodega (Warehouse)</option>
              <option value="Site">Comercio (Site)</option>
              <option value="Staff">Técnico (Staff)</option>
              <option value="Supplier">OPL (Supplier)</option>
            </select>
          </div>
        </div>
      </div>

      <!-- Opciones adicionales -->
      <div style="display:flex;flex-wrap:wrap;gap:16px;padding-top:16px;
                  border-top:1px solid rgba(255,255,255,.06);">
        <label style="display:flex;align-items:center;gap:7px;cursor:pointer;">
          <input type="checkbox" id="exp-opt-cols-extra" checked
            style="accent-color:#22c55e;width:14px;height:14px;">
          <span style="font-size:12px;color:#cbd5e1;font-family:'Outfit',sans-serif;">
            Agregar columnas calculadas (Categoría, Negocio)
          </span>
        </label>
        <label style="display:flex;align-items:center;gap:7px;cursor:pointer;">
          <input type="checkbox" id="exp-opt-freeze" checked
            style="accent-color:#22c55e;width:14px;height:14px;">
          <span style="font-size:12px;color:#cbd5e1;font-family:'Outfit',sans-serif;">
            Congelar fila de encabezados
          </span>
        </label>
        <label style="display:flex;align-items:center;gap:7px;cursor:pointer;">
          <input type="checkbox" id="exp-opt-autofilter" checked
            style="accent-color:#22c55e;width:14px;height:14px;">
          <span style="font-size:12px;color:#cbd5e1;font-family:'Outfit',sans-serif;">
            Auto-filtros en encabezados
          </span>
        </label>
        <label style="display:flex;align-items:center;gap:7px;cursor:pointer;">
          <input type="checkbox" id="exp-opt-resumen" checked
            style="accent-color:#22c55e;width:14px;height:14px;">
          <span style="font-size:12px;color:#cbd5e1;font-family:'Outfit',sans-serif;">
            Incluir hoja de Resumen
          </span>
        </label>
      </div>
    </div>

    <!-- ── Botón principal de export ── -->
    <div style="display:flex;align-items:center;gap:14px;flex-wrap:wrap;margin-bottom:20px;">
      <button id="exp-btn-export"
        onclick="window.invExportCompleto()"
        style="padding:14px 32px;background:linear-gradient(135deg,rgba(34,197,94,.18),rgba(16,185,129,.12));
               border:1px solid rgba(34,197,94,.4);border-radius:14px;
               color:#4ade80;font-size:14px;font-weight:700;cursor:pointer;
               font-family:'Outfit',sans-serif;transition:all .2s;
               display:flex;align-items:center;gap:10px;"
        onmouseover="this.style.background='linear-gradient(135deg,rgba(34,197,94,.28),rgba(16,185,129,.20))';this.style.borderColor='rgba(34,197,94,.7)'"
        onmouseout="this.style.background='linear-gradient(135deg,rgba(34,197,94,.18),rgba(16,185,129,.12))';this.style.borderColor='rgba(34,197,94,.4)'">
        <span style="font-size:18px;">📊</span>
        Exportar Stock Completo (.xlsx)
      </button>

      <button id="exp-btn-csv"
        onclick="window.invExportCompletoCSV()"
        style="padding:14px 24px;background:rgba(153,209,252,.08);
               border:1px solid rgba(153,209,252,.3);border-radius:14px;
               color:#99D1FC;font-size:13px;font-weight:600;cursor:pointer;
               font-family:'Outfit',sans-serif;transition:all .2s;
               display:flex;align-items:center;gap:8px;"
        onmouseover="this.style.background='rgba(153,209,252,.16)'"
        onmouseout="this.style.background='rgba(153,209,252,.08)'">
        <span>🗒️</span> Exportar CSV
      </button>
    </div>

    <!-- ── Barra de progreso ── -->
    <div id="exp-progress-wrap" style="display:none;margin-bottom:20px;">
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:6px;">
        <span id="exp-progress-label" style="font-size:12px;color:#94a3b8;font-family:'Outfit',sans-serif;">
          Preparando export...
        </span>
        <span id="exp-progress-pct" style="font-size:12px;color:#DFFF61;font-family:'JetBrains Mono',monospace;"></span>
      </div>
      <div style="height:6px;background:rgba(255,255,255,.06);border-radius:10px;overflow:hidden;">
        <div id="exp-progress-bar"
          style="height:100%;width:0%;background:linear-gradient(90deg,#22c55e,#4ade80);
                 border-radius:10px;transition:width .3s ease;">
        </div>
      </div>
    </div>

    <!-- ── Log / estado ── -->
    <div id="exp-log" style="display:none;background:rgba(0,0,0,.25);border:1px solid rgba(255,255,255,.06);
         border-radius:12px;padding:16px 18px;font-family:'JetBrains Mono',monospace;
         font-size:11px;color:#64748b;line-height:1.8;max-height:200px;overflow-y:auto;">
    </div>

    <!-- ── Vista previa ── -->
    <div style="margin-top:28px;">
      <div style="font-family:'Syne',sans-serif;font-size:13px;font-weight:700;
                  color:#f1f5f9;margin-bottom:12px;letter-spacing:.2px;">
        🔎 Vista previa — primeras 20 filas del JSON
      </div>
      <div id="exp-preview-wrap" style="overflow-x:auto;border-radius:12px;
           border:1px solid rgba(255,255,255,.07);max-height:320px;overflow-y:auto;">
        <div style="padding:24px;text-align:center;color:#475569;
                    font-family:'Outfit',sans-serif;font-size:12px;">
          Cargando vista previa...
        </div>
      </div>
    </div>

  </div>

  <style>
    .exp-stat-card {
      background: rgba(255,255,255,.03);
      border: 1px solid rgba(255,255,255,.07);
      border-radius: 14px;
      padding: 16px 18px;
      transition: border-color .2s;
    }
    .exp-stat-card:hover { border-color: rgba(34,197,94,.25); }
    .exp-stat-val {
      font-family: 'JetBrains Mono', monospace;
      font-size: 22px;
      font-weight: 700;
      color: #4ade80;
      margin-bottom: 4px;
    }
    .exp-stat-lbl {
      font-family: 'Outfit', sans-serif;
      font-size: 11px;
      color: #64748b;
      text-transform: uppercase;
      letter-spacing: .4px;
    }
    #exp-preview-wrap table {
      width: 100%;
      border-collapse: collapse;
      font-family: 'Outfit', sans-serif;
      font-size: 11px;
    }
    #exp-preview-wrap th {
      position: sticky;
      top: 0;
      background: #0d1929;
      color: #94a3b8;
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: .4px;
      padding: 8px 12px;
      border-bottom: 1px solid rgba(255,255,255,.1);
      white-space: nowrap;
      z-index: 1;
    }
    #exp-preview-wrap td {
      padding: 7px 12px;
      border-bottom: 1px solid rgba(255,255,255,.04);
      color: #cbd5e1;
      white-space: nowrap;
      max-width: 220px;
      overflow: hidden;
      text-overflow: ellipsis;
    }
    #exp-preview-wrap tr:hover td { background: rgba(255,255,255,.03); }
    .exp-col-toggle {
      display: flex;
      align-items: center;
      gap: 8px;
      cursor: pointer;
    }
    .exp-col-toggle input { accent-color: #22c55e; width: 13px; height: 13px; cursor: pointer; }
    .exp-col-toggle span {
      font-size: 11px;
      color: #94a3b8;
      font-family: 'Outfit', sans-serif;
    }
  </style>`;
}

// ══════════════════════════════════════════════════════════════════
//  Columnas canónicas del JSON (todas las del stock)
// ══════════════════════════════════════════════════════════════════
const EXP_COLUMNS = [
  { key: 'Nombre',                    label: 'Referencia / Nombre',      default: true  },
  { key: 'Cantidad',                  label: 'Cantidad',                 default: true  },
  { key: 'Número de serie',           label: 'Número de serie',          default: true  },
  { key: 'Código de ubicación',       label: 'Código de ubicación',      default: true  },
  { key: 'Nombre de la ubicación',    label: 'Nombre de la ubicación',   default: true  },
  { key: 'Tipo de ubicación',         label: 'Tipo de ubicación',        default: true  },
  { key: 'Subtipo',                   label: 'Subtipo',                  default: true  },
  { key: 'Código de comercio',        label: 'Código de comercio',       default: true  },
  { key: 'Lote',                      label: 'Lote',                     default: false },
  { key: 'Código de producto',        label: 'Código de producto',       default: false },
  { key: 'Estado',                    label: 'Estado',                   default: false },
  { key: 'Fecha de vencimiento',      label: 'Fecha de vencimiento',     default: false },
  { key: 'Propietario',               label: 'Propietario',              default: false },
];

// ── Renderizar checkboxes de columnas ─────────────────────────────
function _attachExportEvents() {
  const list = document.getElementById('exp-cols-list');
  if (!list) return;

  list.innerHTML = EXP_COLUMNS.map((col, i) => `
    <label class="exp-col-toggle">
      <input type="checkbox" id="exp-col-${i}" data-key="${col.key}" ${col.default ? 'checked' : ''}>
      <span>${col.label}</span>
    </label>`).join('');
}

// ══════════════════════════════════════════════════════════════════
//  Cargar stats y vista previa
// ══════════════════════════════════════════════════════════════════
async function _loadPreviewStats() {
  // Esperar a que INV_RAW esté disponible (puede estar cargando)
  let attempts = 0;
  while (!window.INV_RAW && attempts < 30) {
    await new Promise(r => setTimeout(r, 300));
    attempts++;
  }

  const raw = window.INV_RAW || [];
  if (!raw.length) {
    _expLog('⚠ No se encontraron datos en INV_RAW. Asegúrate de haber abierto el tablero de Inventario primero.');
    return;
  }

  // Stats
  const totalFilas  = raw.length;
  const totalUds    = raw.reduce((a, r) => a + (parseInt(r['Cantidad']) || 0), 0);
  const refs        = new Set(raw.map(r => r['Nombre']).filter(Boolean)).size;
  const locs        = new Set(raw.map(r => r['Nombre de la ubicación']).filter(Boolean)).size;

  document.getElementById('exp-stat-total-val').textContent  = totalFilas.toLocaleString('es-CO');
  document.getElementById('exp-stat-refs-val').textContent   = refs.toLocaleString('es-CO');
  document.getElementById('exp-stat-bodegas-val').textContent = locs.toLocaleString('es-CO');
  document.getElementById('exp-stat-uds-val').textContent    = totalUds.toLocaleString('es-CO');

  // Vista previa (primeras 20 filas)
  _renderPreview(raw.slice(0, 20));
}

function _renderPreview(rows) {
  const wrap = document.getElementById('exp-preview-wrap');
  if (!wrap || !rows.length) return;

  // Detectar todas las claves presentes en las primeras filas
  const allKeys = [...new Set(rows.flatMap(r => Object.keys(r)))];

  wrap.innerHTML = `<table>
    <thead><tr>${allKeys.map(k => `<th title="${k}">${k}</th>`).join('')}</tr></thead>
    <tbody>
      ${rows.map(r => `<tr>${allKeys.map(k => {
        const v = r[k] !== undefined && r[k] !== null ? String(r[k]) : '';
        return `<td title="${v}">${v || '<span style="color:#334155">—</span>'}</td>`;
      }).join('')}</tr>`).join('')}
    </tbody>
  </table>`;
}

// ══════════════════════════════════════════════════════════════════
//  Helpers UI
// ══════════════════════════════════════════════════════════════════
function _expSetProgress(pct, label) {
  const wrap = document.getElementById('exp-progress-wrap');
  const bar  = document.getElementById('exp-progress-bar');
  const lbl  = document.getElementById('exp-progress-label');
  const pctEl = document.getElementById('exp-progress-pct');
  if (!wrap) return;
  wrap.style.display = 'block';
  if (bar)   bar.style.width = pct + '%';
  if (lbl)   lbl.textContent = label || '';
  if (pctEl) pctEl.textContent = pct + '%';
}

function _expLog(msg) {
  const log = document.getElementById('exp-log');
  if (!log) return;
  log.style.display = 'block';
  const ts = new Date().toLocaleTimeString('es-CO');
  log.innerHTML += `<div><span style="color:#334155;">[${ts}]</span> ${msg}</div>`;
  log.scrollTop = log.scrollHeight;
}

function _expSetBtn(loading) {
  const btn = document.getElementById('exp-btn-export');
  const csv = document.getElementById('exp-btn-csv');
  if (btn) { btn.disabled = loading; btn.style.opacity = loading ? '.5' : '1'; }
  if (csv) { csv.disabled = loading; csv.style.opacity = loading ? '.5' : '1'; }
}

// ══════════════════════════════════════════════════════════════════
//  Obtener datos filtrados para el export
// ══════════════════════════════════════════════════════════════════
function _getExportRows() {
  const raw = window.INV_RAW || [];
  const cat  = (document.getElementById('exp-filter-cat')?.value  || '').trim();
  const neg  = (document.getElementById('exp-filter-neg')?.value  || '').trim();
  const tipo = (document.getElementById('exp-filter-tipo')?.value || '').trim();

  let rows = raw;
  if (cat)  rows = rows.filter(r => invCategoria(r['Nombre']) === cat);
  if (neg)  rows = rows.filter(r => invNegocio(r['Subtipo'])   === neg);
  if (tipo) rows = rows.filter(r => (r['Tipo de ubicación']||'').trim() === tipo);
  return rows;
}

// ── Columnas seleccionadas ────────────────────────────────────────
function _getSelectedCols() {
  return EXP_COLUMNS.filter((_, i) => {
    const cb = document.getElementById(`exp-col-${i}`);
    return cb ? cb.checked : false;
  });
}

// ══════════════════════════════════════════════════════════════════
//  EXPORT XLSX
// ══════════════════════════════════════════════════════════════════
window.invExportCompleto = async function () {
  if (typeof XLSX === 'undefined') {
    alert('⚠ Librería XLSX no disponible. Recarga la página.');
    return;
  }

  _expSetBtn(true);
  _expLog('🚀 Iniciando export completo de stock...');
  _expSetProgress(5, 'Preparando datos...');

  await new Promise(r => setTimeout(r, 30)); // ceder al hilo UI

  try {
    const rows       = _getExportRows();
    const selCols    = _getSelectedCols();
    const addCalc    = document.getElementById('exp-opt-cols-extra')?.checked ?? true;
    const doFreeze   = document.getElementById('exp-opt-freeze')?.checked    ?? true;
    const doFilter   = document.getElementById('exp-opt-autofilter')?.checked ?? true;
    const doResumen  = document.getElementById('exp-opt-resumen')?.checked   ?? true;

    _expLog(`📦 Filas a exportar: ${rows.length.toLocaleString('es-CO')}`);
    _expLog(`📋 Columnas seleccionadas: ${selCols.length}${addCalc ? ' + 2 calculadas' : ''}`);
    _expSetProgress(15, 'Construyendo filas del Excel...');

    await new Promise(r => setTimeout(r, 20));

    // ── Construir array de objetos para SheetJS ───────────────────
    const CHUNK = 5000;
    const exportData = [];
    const total = rows.length;

    for (let i = 0; i < total; i++) {
      const r   = rows[i];
      const obj = {};

      selCols.forEach(col => {
        // Alias: algunos campos tienen variantes de nombre
        let val = r[col.key];
        if (val === undefined || val === null) {
          // Intentar alias comunes
          if (col.key === 'Número de serie') val = r['Numero de serie'] ?? r['Serial'];
          if (col.key === 'Código de comercio') val = r['codigo_de_comercio'];
        }
        obj[col.label] = val !== undefined && val !== null ? val : '';
      });

      if (addCalc) {
        obj['Categoría']  = invCategoria(r['Nombre']);
        obj['Negocio']    = invNegocio(r['Subtipo']);
      }

      exportData.push(obj);

      // Actualizar progreso cada CHUNK filas
      if ((i + 1) % CHUNK === 0 || i === total - 1) {
        const pct = Math.round(15 + ((i + 1) / total) * 55);
        _expSetProgress(pct, `Procesando fila ${(i+1).toLocaleString('es-CO')} de ${total.toLocaleString('es-CO')}...`);
        await new Promise(r2 => setTimeout(r2, 0)); // yield
      }
    }

    _expSetProgress(72, 'Creando hoja Excel...');
    await new Promise(r => setTimeout(r, 20));

    const wb = XLSX.utils.book_new();

    // ── Hoja principal: Stock Completo ────────────────────────────
    const ws = XLSX.utils.json_to_sheet(exportData);

    // Ancho de columnas automático
    if (exportData.length) {
      const headers = Object.keys(exportData[0]);
      ws['!cols'] = headers.map(h => ({
        wch: Math.min(Math.max(h.length + 3, 14), 45)
      }));
    }

    // Freeze row 1
    if (doFreeze) {
      ws['!freeze'] = { xSplit: 0, ySplit: 1, topLeftCell: 'A2', activePane: 'bottomLeft', state: 'frozen' };
    }

    // Auto-filter
    if (doFilter && exportData.length) {
      const lastCol = _colLetter(Object.keys(exportData[0]).length);
      ws['!autofilter'] = { ref: `A1:${lastCol}1` };
    }

    XLSX.utils.book_append_sheet(wb, ws, 'Stock Completo');
    _expLog('✅ Hoja "Stock Completo" creada.');

    // ── Hoja de Resumen ───────────────────────────────────────────
    if (doResumen) {
      _expSetProgress(82, 'Generando hoja de resumen...');
      await new Promise(r => setTimeout(r, 10));
      const wsRes = _buildResumenSheet(rows);
      XLSX.utils.book_append_sheet(wb, wsRes, 'Resumen');
      _expLog('✅ Hoja "Resumen" creada.');
    }

    _expSetProgress(92, 'Escribiendo archivo...');
    await new Promise(r => setTimeout(r, 20));

    // Nombre del archivo
    const ts      = new Date().toISOString().slice(0, 10);
    const catTag  = (document.getElementById('exp-filter-cat')?.value || '').replace(/\s/g,'_') || 'Completo';
    const negTag  = (document.getElementById('exp-filter-neg')?.value || '');
    const parts   = ['Stock_Wompi', catTag, negTag, ts].filter(Boolean);
    const fname   = parts.join('_') + '.xlsx';

    XLSX.writeFile(wb, fname);

    _expSetProgress(100, '¡Export completado!');
    _expLog(`🎉 Archivo <strong style="color:#4ade80">${fname}</strong> descargado — ${rows.length.toLocaleString('es-CO')} filas.`);

    setTimeout(() => {
      document.getElementById('exp-progress-wrap').style.display = 'none';
    }, 3500);

  } catch (err) {
    console.error('[ExportStock]', err);
    _expLog(`❌ Error: ${err.message}`);
    _expSetProgress(0, 'Error en el export');
  }

  _expSetBtn(false);
};

// ══════════════════════════════════════════════════════════════════
//  EXPORT CSV (alternativa ligera)
// ══════════════════════════════════════════════════════════════════
window.invExportCompletoCSV = function () {
  if (typeof XLSX === 'undefined') {
    alert('⚠ Librería XLSX no disponible.');
    return;
  }

  _expSetBtn(true);
  _expLog('🗒️ Generando CSV...');

  try {
    const rows    = _getExportRows();
    const selCols = _getSelectedCols();
    const addCalc = document.getElementById('exp-opt-cols-extra')?.checked ?? true;

    const exportData = rows.map(r => {
      const obj = {};
      selCols.forEach(col => {
        let val = r[col.key];
        if (val === undefined || val === null) {
          if (col.key === 'Número de serie') val = r['Numero de serie'] ?? r['Serial'];
          if (col.key === 'Código de comercio') val = r['codigo_de_comercio'];
        }
        obj[col.label] = val !== undefined && val !== null ? val : '';
      });
      if (addCalc) {
        obj['Categoría'] = invCategoria(r['Nombre']);
        obj['Negocio']   = invNegocio(r['Subtipo']);
      }
      return obj;
    });

    const ws  = XLSX.utils.json_to_sheet(exportData);
    const csv = XLSX.utils.sheet_to_csv(ws);
    const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8;' });
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    a.href     = url;
    a.download = `Stock_Wompi_${new Date().toISOString().slice(0,10)}.csv`;
    a.click();
    URL.revokeObjectURL(url);

    _expLog(`✅ CSV descargado — ${rows.length.toLocaleString('es-CO')} filas.`);

  } catch (err) {
    _expLog(`❌ Error CSV: ${err.message}`);
  }

  _expSetBtn(false);
};

// ══════════════════════════════════════════════════════════════════
//  Hoja de Resumen
// ══════════════════════════════════════════════════════════════════
function _buildResumenSheet(rows) {
  const ts = new Date().toLocaleString('es-CO');
  const resData = [];

  // ── KPIs generales ───────────────────────────────────────────────
  const totalUds  = rows.reduce((a, r) => a + (parseInt(r['Cantidad'])||0), 0);
  resData.push({ 'Campo': '── RESUMEN GENERAL ──', 'Valor': '' });
  resData.push({ 'Campo': 'Fecha de export',    'Valor': ts });
  resData.push({ 'Campo': 'Total filas',        'Valor': rows.length });
  resData.push({ 'Campo': 'Total unidades',     'Valor': totalUds });
  resData.push({ 'Campo': 'Refs. únicas',       'Valor': new Set(rows.map(r=>r['Nombre']).filter(Boolean)).size });
  resData.push({ 'Campo': 'Ubicaciones únicas', 'Valor': new Set(rows.map(r=>r['Nombre de la ubicación']).filter(Boolean)).size });
  resData.push({ 'Campo': '', 'Valor': '' });

  // ── Por categoría ─────────────────────────────────────────────
  resData.push({ 'Campo': '── POR CATEGORÍA ──', 'Valor': '' });
  const catMap = {};
  rows.forEach(r => {
    const c = invCategoria(r['Nombre']);
    if (!catMap[c]) catMap[c] = { filas: 0, uds: 0 };
    catMap[c].filas++;
    catMap[c].uds += parseInt(r['Cantidad']) || 0;
  });
  Object.entries(catMap).sort((a,b) => b[1].uds - a[1].uds).forEach(([cat, d]) => {
    resData.push({ 'Campo': cat, 'Valor': d.uds, 'Filas': d.filas });
  });
  resData.push({ 'Campo': '', 'Valor': '' });

  // ── Por negocio ──────────────────────────────────────────────
  resData.push({ 'Campo': '── POR NEGOCIO ──', 'Valor': '' });
  const negMap = {};
  rows.forEach(r => {
    const n = invNegocio(r['Subtipo']);
    if (!negMap[n]) negMap[n] = { filas: 0, uds: 0 };
    negMap[n].filas++;
    negMap[n].uds += parseInt(r['Cantidad']) || 0;
  });
  Object.entries(negMap).forEach(([neg, d]) => {
    resData.push({ 'Campo': neg, 'Valor': d.uds, 'Filas': d.filas });
  });
  resData.push({ 'Campo': '', 'Valor': '' });

  // ── Por tipo de ubicación ────────────────────────────────────
  resData.push({ 'Campo': '── POR TIPO DE UBICACIÓN ──', 'Valor': '' });
  const tipoMap = {};
  rows.forEach(r => {
    const t = (r['Tipo de ubicación']||'Sin tipo').trim() || 'Sin tipo';
    if (!tipoMap[t]) tipoMap[t] = { filas: 0, uds: 0 };
    tipoMap[t].filas++;
    tipoMap[t].uds += parseInt(r['Cantidad']) || 0;
  });
  Object.entries(tipoMap).sort((a,b) => b[1].uds - a[1].uds).forEach(([t, d]) => {
    resData.push({ 'Campo': t, 'Valor': d.uds, 'Filas': d.filas });
  });

  const ws = XLSX.utils.json_to_sheet(resData);
  ws['!cols'] = [{ wch: 38 }, { wch: 18 }, { wch: 12 }];
  return ws;
}

// ══════════════════════════════════════════════════════════════════
//  Helpers
// ══════════════════════════════════════════════════════════════════
function _colLetter(n) {
  let s = '';
  while (n > 0) {
    const mod = (n - 1) % 26;
    s = String.fromCharCode(65 + mod) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}