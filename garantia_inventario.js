/**
 * ╔══════════════════════════════════════════════════════════════════╗
 * ║  garantia_inventario.js                                         ║
 * ║  Tab: "Estado garantía en Datáfonos y tiempo de vida de         ║
 * ║        rollos en bodega"                                         ║
 * ║  Sección 1 — Garantía datáfonos                                 ║
 * ║  Sección 2 — Días de inventario de rollos en almacén            ║
 * ╚══════════════════════════════════════════════════════════════════╝
 */

'use strict';

// ══════════════════════════════════════════════════════════════════
//  CONSTANTES
// ══════════════════════════════════════════════════════════════════

const GI_REFS_DATAFONO = new Set([
  'PINPAD DESK 1700 - INGENICO',
  'DATAFONO EX6000 - INGENICO',
  'DATAFONO EX4000 - INGENICO',
  'DATAFONO DX4000 PORTATIL - INGENICO',
  'DATAFONO DX4000 ESCRITORIO - INGENICO',
]);

// 42 bodegas hardcodeadas
const GI_BODEGAS = new Set([
  "ALMACEN WOMPI MEDELLIN","ALMACEN WOMPI BOGOTA","ALMACEN WOMPI BUCARAMANGA",
  "ALMACEN WOMPI CALI","ALMACEN WOMPI VILLAVICENCIO","ALMACEN WOMPI CUCUTA",
  "ALMACEN WOMPI PEREIRA","ALMACEN WOMPI NEIVA","ALMACEN WOMPI IBAGUE",
  "ALMACEN WOMPI TUNJA","ALMACEN WOMPI MONTERIA","ALMACEN WOMPI SANTA MARTA",
  "ALMACEN WOMPI VALLEDUPAR","ALMACEN WOMPI CARTAGENA","ALMACEN WOMPI FLORENCIA",
  "ALMACEN WOMPI POPAYAN","ALMACEN WOMPI MANIZALES","ALMACEN WOMPI YOPAL",
  "ALMACEN WOMPI APARTADO","ALMACEN WOMPI QUIBDO","ALMACEN WOMPI PASTO",
  "ALMACEN WOMPI SINCELEJO","ALMACEN WOMPI BARRANQUILLA","ALMACEN WOMPI ARMENIA",
  "ALMACEN BAJAS WOMPI","ALMACEN WOMPI ALISTAMIENTO MEDELLIN",
  "ALMACEN WOMPI VP MEDELLIN | ALQUILER","ALMACEN WOMPI VP BOGOTA | ALQUILER",
  "ALMACEN WOMPI VP CALI | ALQUILER","ALMACEN WOMPI VP BUCARAMANGA | ALQUILER",
  "ALMACEN WOMPI VP PEREIRA | ALQUILER","ALMACEN WOMPI VP BARRANQUILLA | ALQUILER",
  "ALMACEN WOMPI VP MONTERIA | ALQUILER","ALMACEN WOMPI VP MEDELLIN | VENTA",
  "ALMACEN WOMPI VP BOGOTA | VENTA","ALMACEN WOMPI VP CALI | VENTA",
  "ALMACEN WOMPI VP BUCARAMANGA | VENTA","ALMACEN WOMPI VP PEREIRA | VENTA",
  "ALMACEN WOMPI VP BARRANQUILLA | VENTA","ALMACEN WOMPI VP MONTERIA | VENTA",
  "ALMACEN WOMPI VP ALISTAMIENTO MEDELLIN","ALMACEN INGENICO - PROVEEDOR WOMPI",
]);

const GI_COLORS = {
  vigente:    '#B0F2AE',
  porVencer:  '#FFC04D',
  vencida:    '#FF5C5C',
  sinFecha:   '#64748b',
  nuevo:      '#99D1FC',
  usado:      '#C084FC',
  palette:    ['#B0F2AE','#99D1FC','#DFFF61','#C084FC','#FFC04D','#FF5C5C','#F49D6E','#7B8CDE'],
};

// ── Estado ─────────────────────────────────────────────────────────
let GI_CHARTS       = {};
let GI_DF_ALL       = [];   // datáfonos procesados
let GI_DF_FILTERED  = [];
let GI_DF_PAGE      = 1;
let GI_ACC_ALL      = [];   // acumulado por referencia
let GI_ACC_FILTERED = [];
let GI_RO_ALL       = [];   // rollos procesados
let GI_RO_FILTERED  = [];
let GI_RO_PAGE      = 1;
const GI_PAGE_SIZE  = 50;

// Filtros datáfonos
let GIF_SERIAL  = '';
let GIF_REUSO   = '';
let GIF_ESTADO  = '';
// Filtros rollos
let GIR_BODEGA  = '';
let GIR_ESTADO  = '';

// ══════════════════════════════════════════════════════════════════
//  HELPERS
// ══════════════════════════════════════════════════════════════════

function _giRaw() {
  return window.INV_RAW || [];
}

function _giParseDate(str) {
  if (!str) return null;
  // DD/MM/YYYY
  const m = String(str).trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
  // fallback ISO
  const d = new Date(str);
  return isNaN(d) ? null : d;
}

function _giDiffDays(date) {
  if (!date) return null;
  const now = new Date(); now.setHours(0,0,0,0);
  const d   = new Date(date); d.setHours(0,0,0,0);
  return Math.round((d - now) / 86400000);
}

function _giEstadoGarantia(dias) {
  if (dias === null || dias === undefined) return 'Sin fecha';
  if (dias < 0)   return 'Vencida';
  if (dias <= 60) return 'Por vencer (≤ 60 días)';
  return 'Vigente';
}

function _giEstadoBadge(estado) {
  const map = {
    'Vigente':               { bg:'rgba(176,242,174,.15)', color:'#B0F2AE', icon:'✅' },
    'Por vencer (≤ 60 días)':{ bg:'rgba(255,192,77,.15)',  color:'#FFC04D', icon:'⚠️' },
    'Vencida':               { bg:'rgba(255,92,92,.15)',   color:'#FF5C5C', icon:'❌' },
    'Sin fecha':             { bg:'rgba(100,116,139,.15)', color:'#94a3b8', icon:'—'  },
  };
  const s = map[estado] || map['Sin fecha'];
  return `<span style="display:inline-flex;align-items:center;gap:4px;padding:3px 10px;border-radius:20px;background:${s.bg};color:${s.color};font-size:11px;font-weight:600;">${s.icon} ${estado}</span>`;
}

function _giParseAtributos(raw) {
  if (!raw) return {};
  // "WOMPI, FC: 28/01/2025, FR: 27/08/2025, FG: 27/08/2026, PU: 745608.39, FA: 15/09/2025"
  const out = {};
  const parts = String(raw).split(',');
  for (const part of parts) {
    const kv = part.trim();
    const colon = kv.indexOf(':');
    if (colon < 0) continue;
    const key = kv.substring(0, colon).trim().toUpperCase();
    const val = kv.substring(colon + 1).trim();
    out[key] = val;
  }
  return out;
}

function _giFmtCOP(v) {
  if (!v && v !== 0) return '—';
  return '$' + Number(v).toLocaleString('es-CO', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
}

function _giFmtDate(d) {
  if (!d) return '—';
  return d.toLocaleDateString('es-CO', { day:'2-digit', month:'2-digit', year:'numeric' });
}

function _giDestroyChart(id) {
  if (GI_CHARTS[id]) { try { GI_CHARTS[id].destroy(); } catch(e){} delete GI_CHARTS[id]; }
}

function _giMkChart(id, cfg) {
  _giDestroyChart(id);
  const ctx = document.getElementById(id);
  if (!ctx) return;
  GI_CHARTS[id] = new Chart(ctx, cfg);
}

// ══════════════════════════════════════════════════════════════════
//  PROCESADO DE DATOS
// ══════════════════════════════════════════════════════════════════

function _giBuildDatafonos() {
  const raw = _giRaw();
  const rows = [];
  for (const r of raw) {
    const nombre = (r['Nombre'] || r['nombre'] || '').trim();
    if (!GI_REFS_DATAFONO.has(nombre)) continue;
    const attr = _giParseAtributos(r['Atributos'] || r['atributos'] || '');
    const fc = _giParseDate(attr['FC']);
    const fr = _giParseDate(attr['FR']);
    const fg = _giParseDate(attr['FG']);
    const pu = parseFloat((attr['PU'] || '').replace(',','.')) || null;
    const diasGarantia = _giDiffDays(fg);
    const estadoGarantia = _giEstadoGarantia(diasGarantia);
    const diasInventario = fr ? Math.round((new Date() - fr) / 86400000) : null;
    rows.push({
      referencia:     nombre,
      serial:         (r['Número de serie'] || r['numero_de_serie'] || '').trim(),
      fc, fr, fg, pu,
      diasGarantia, estadoGarantia, diasInventario,
      reusado: r['tipo_de_operacion'] === 'return',
      cantidad: parseInt(r['Cantidad']) || 1,
    });
  }
  GI_DF_ALL = rows;
}

function _giBuildAcumulado() {
  const map = {};
  for (const r of GI_DF_ALL) {
    if (!map[r.referencia]) map[r.referencia] = { referencia: r.referencia, unidades: 0, acumulado: 0 };
    const qty = r.cantidad;
    map[r.referencia].unidades  += qty;
    map[r.referencia].acumulado += (r.pu || 0) * qty;
  }
  GI_ACC_ALL = Object.values(map).sort((a,b) => b.acumulado - a.acumulado);
}

function _giBuildRollos() {
  const raw = _giRaw();
  const rows = [];
  const RE_ROLLO = /^ROLLOS\s+WOMPI\s+(\d{6})\s*$/i;
  for (const r of raw) {
    const nombre   = (r['Nombre'] || r['nombre'] || '').trim();
    const ubicacion = (r['Nombre de la ubicación'] || r['nombre_de_la_ubicacion'] || '').trim();
    if (!GI_BODEGAS.has(ubicacion)) continue;
    const m = nombre.match(RE_ROLLO);
    if (!m) continue;
    // DDMMYY → fecha ingreso
    const code = m[1]; // "301225"
    const dd = parseInt(code.substring(0,2));
    const mm = parseInt(code.substring(2,4));
    const yy = parseInt(code.substring(4,6)) + 2000;
    const fechaIngreso = new Date(yy, mm - 1, dd);
    const fechaVence   = new Date(yy, mm - 1 + 6, dd);
    const diasRestantes = _giDiffDays(fechaVence);
    const estadoRollo   = diasRestantes === null ? 'Sin fecha' : diasRestantes < 0 ? 'Vencido' : diasRestantes <= 30 ? 'Por vencer (≤ 30 días)' : 'Vigente';
    rows.push({
      nombre, ubicacion,
      cantidad: parseInt(r['Cantidad']) || 0,
      fechaIngreso, fechaVence, diasRestantes, estadoRollo,
    });
  }
  // Agrupar por bodega + nombre (mismo lote)
  const map = {};
  for (const r of rows) {
    const key = r.ubicacion + '||' + r.nombre;
    if (!map[key]) {
      map[key] = { ...r };
    } else {
      map[key].cantidad += r.cantidad;
    }
  }
  GI_RO_ALL = Object.values(map).sort((a,b) => a.diasRestantes - b.diasRestantes);
}

// ══════════════════════════════════════════════════════════════════
//  FILTROS
// ══════════════════════════════════════════════════════════════════

function _giApplyFiltersDF() {
  GI_DF_FILTERED = GI_DF_ALL.filter(r => {
    if (GIF_SERIAL && !r.serial.toLowerCase().includes(GIF_SERIAL.toLowerCase())) return false;
    if (GIF_REUSO === 'nuevo'  &&  r.reusado) return false;
    if (GIF_REUSO === 'usado'  && !r.reusado) return false;
    if (GIF_ESTADO && r.estadoGarantia !== GIF_ESTADO) return false;
    return true;
  });
  GI_DF_PAGE = 1;
}

function _giApplyFiltersRO() {
  GI_RO_FILTERED = GI_RO_ALL.filter(r => {
    if (GIR_BODEGA && r.ubicacion !== GIR_BODEGA) return false;
    if (GIR_ESTADO && r.estadoRollo !== GIR_ESTADO) return false;
    return true;
  });
  GI_RO_PAGE = 1;
}

// ── Acumulado no tiene filtros propios (usa mismos que df) ─────────
function _giApplyFiltersACC() {
  // Recalcular acumulado solo sobre registros filtrados
  const map = {};
  for (const r of GI_DF_FILTERED) {
    if (!map[r.referencia]) map[r.referencia] = { referencia: r.referencia, unidades: 0, acumulado: 0 };
    const qty = r.cantidad;
    map[r.referencia].unidades  += qty;
    map[r.referencia].acumulado += (r.pu || 0) * qty;
  }
  GI_ACC_FILTERED = Object.values(map).sort((a,b) => b.acumulado - a.acumulado);
}

// ══════════════════════════════════════════════════════════════════
//  RENDER PRINCIPAL
// ══════════════════════════════════════════════════════════════════

window.renderGarantiaInventario = async function() {
  const panel = document.getElementById('panel-garantia-inventario');
  if (!panel) return;

  // Asegurar datos cargados
  if (!window.INV_RAW || !window.INV_RAW.length) {
    if (typeof loadInventarioData === 'function') await loadInventarioData();
  }

  _giBuildDatafonos();
  _giBuildAcumulado();
  _giBuildRollos();
  _giApplyFiltersDF();
  _giApplyFiltersACC();
  _giApplyFiltersRO();

  panel.innerHTML = _giLayoutHTML();
  _giBindFilters();
  _giRenderAll();
};

// ══════════════════════════════════════════════════════════════════
//  HTML LAYOUT
// ══════════════════════════════════════════════════════════════════

function _giLayoutHTML() {
  return `
<div style="padding:4px 0 60px;">

  <!-- ════ SECCIÓN 1: GARANTÍA DATÁFONOS ════ -->
  <div class="section-label fade-up" style="color:#B0F2AE;font-size:16px;margin-bottom:4px;">🔒 Estado de Garantía en Datáfonos</div>
  <div style="font-size:12px;color:#64748b;margin-bottom:24px;">Datáfonos en inventario · Garantía calculada a partir de FG en atributos</div>

  <!-- KPI Strip -->
  <div id="gi-kpi-strip" style="display:flex;flex-wrap:wrap;gap:12px;margin-bottom:28px;"></div>

  <!-- Filtros datáfonos -->
  <div class="filters-bar" style="margin-bottom:24px;">
    <div class="filters-title">Filtros — Datáfonos</div>
    <div class="filters-row" style="flex-wrap:wrap;gap:12px;align-items:flex-end;">
      <div class="filter-group" style="min-width:200px;">
        <label>Número de Serie</label>
        <input type="text" id="gi-f-serial" placeholder="🔍 Buscar serial..." style="width:100%;box-sizing:border-box;" oninput="giFilterSerial(this.value)">
      </div>
      <div class="filter-group">
        <label>Condición</label>
        <select id="gi-f-reuso" onchange="giFilterReuso(this.value)">
          <option value="">Todos</option>
          <option value="nuevo">🆕 Nuevo</option>
          <option value="usado">♻️ Usado (Return)</option>
        </select>
      </div>
      <div class="filter-group">
        <label>Estado Garantía</label>
        <select id="gi-f-estado" onchange="giFilterEstado(this.value)">
          <option value="">Todos</option>
          <option value="Vigente">✅ Vigente</option>
          <option value="Por vencer (≤ 60 días)">⚠️ Por vencer (≤ 60 días)</option>
          <option value="Vencida">❌ Vencida</option>
          <option value="Sin fecha">— Sin fecha</option>
        </select>
      </div>
      <div class="filters-actions">
        <button class="btn-apply" onclick="giApplyAndRender()">✦ Aplicar</button>
        <button class="btn-reset" onclick="giResetFilters()">↺ Reset</button>
      </div>
    </div>
  </div>

  <!-- Gráficas datáfonos -->
  <div id="gi-charts-df" style="display:grid;grid-template-columns:repeat(3,1fr);gap:18px;margin-bottom:32px;">
    <div class="chart-card"><div class="chart-header"><div class="chart-title">Estado de Garantía</div><div class="chart-sub">Distribución por estado</div></div><div class="chart-wrap"><canvas id="gi-chart-estado"></canvas></div></div>
    <div class="chart-card"><div class="chart-header"><div class="chart-title">Garantía por Referencia</div><div class="chart-sub">Unidades por referencia y estado</div></div><div class="chart-wrap"><canvas id="gi-chart-ref-estado"></canvas></div></div>
    <div class="chart-card"><div class="chart-header"><div class="chart-title">Nuevo vs Usado</div><div class="chart-sub">Condición del equipo</div></div><div class="chart-wrap"><canvas id="gi-chart-reuso"></canvas></div></div>
  </div>

  <!-- Tabla 1: Detalle por serial -->
  <div class="section-label fade-up" style="color:#B0F2AE;">Detalle por Serial</div>
  <div style="background:linear-gradient(145deg,rgba(10,26,18,.95),rgba(8,20,14,.9));border:1px solid rgba(176,242,174,.12);border-radius:18px;overflow:hidden;margin-bottom:32px;box-shadow:0 4px 20px rgba(0,0,0,.4);">
    <div style="display:flex;align-items:center;justify-content:space-between;padding:16px 20px;border-bottom:1px solid rgba(176,242,174,.08);">
      <div>
        <div style="font-family:'Syne',sans-serif;font-size:13px;font-weight:700;color:#f1f5f9;">Tabla 1 — Referencia · Serial · Fechas · Garantía · Días en Inventario</div>
        <div style="font-size:11px;color:#7A7674;margin-top:2px;" id="gi-t1-count">0 registros</div>
      </div>
      <button onclick="giExportT1()" style="background:rgba(176,242,174,.08);border:1px solid rgba(176,242,174,.2);color:#B0F2AE;padding:7px 14px;border-radius:8px;cursor:pointer;font-size:12px;font-family:'Outfit',sans-serif;">⬇ Excel</button>
    </div>
    <div id="gi-t1-wrap" style="overflow-x:auto;max-height:480px;"></div>
    <div style="display:flex;align-items:center;justify-content:flex-end;padding:10px 20px;border-top:1px solid rgba(176,242,174,.07);">
      <div id="gi-t1-pag" style="display:flex;gap:6px;flex-wrap:wrap;"></div>
    </div>
  </div>

  <!-- Tabla 2: Acumulado por referencia -->
  <div class="section-label fade-up" style="color:#DFFF61;">$$ Acumulado por Referencia</div>
  <div style="background:linear-gradient(145deg,rgba(10,18,10,.95),rgba(8,14,8,.9));border:1px solid rgba(223,255,97,.12);border-radius:18px;overflow:hidden;margin-bottom:40px;box-shadow:0 4px 20px rgba(0,0,0,.4);">
    <div style="display:flex;align-items:center;justify-content:space-between;padding:16px 20px;border-bottom:1px solid rgba(223,255,97,.08);">
      <div>
        <div style="font-family:'Syne',sans-serif;font-size:13px;font-weight:700;color:#f1f5f9;">Tabla 2 — Referencia · Precio Acumulado · Unidades Totales</div>
        <div style="font-size:11px;color:#7A7674;margin-top:2px;" id="gi-t2-count">0 referencias</div>
      </div>
      <button onclick="giExportT2()" style="background:rgba(223,255,97,.08);border:1px solid rgba(223,255,97,.2);color:#DFFF61;padding:7px 14px;border-radius:8px;cursor:pointer;font-size:12px;font-family:'Outfit',sans-serif;">⬇ Excel</button>
    </div>
    <div id="gi-t2-wrap" style="overflow-x:auto;"></div>
    <!-- Gráfica precio acumulado -->
    <div style="padding:20px;border-top:1px solid rgba(223,255,97,.07);">
      <div style="font-family:'Syne',sans-serif;font-size:12px;font-weight:600;color:#DFFF61;margin-bottom:12px;">💰 Precio Acumulado por Referencia (COP)</div>
      <div style="max-height:220px;"><canvas id="gi-chart-acumulado"></canvas></div>
    </div>
  </div>

  <!-- ════ SECCIÓN 2: ROLLOS EN ALMACÉN ════ -->
  <div class="section-label fade-up" style="color:#99D1FC;font-size:16px;margin-bottom:4px;">🎞️ Días de Inventario de Rollos en Almacén</div>
  <div style="font-size:12px;color:#64748b;margin-bottom:24px;">Rollos WOMPI en las 42 bodegas · Vida útil máxima: 6 meses desde ingreso</div>

  <!-- KPI Strip rollos -->
  <div id="gi-kpi-rollos" style="display:flex;flex-wrap:wrap;gap:12px;margin-bottom:28px;"></div>

  <!-- Filtros rollos -->
  <div class="filters-bar" style="margin-bottom:24px;">
    <div class="filters-title">Filtros — Rollos</div>
    <div class="filters-row" style="flex-wrap:wrap;gap:12px;align-items:flex-end;">
      <div class="filter-group" style="min-width:260px;">
        <label>Bodega</label>
        <select id="gi-r-bodega" onchange="giFilterRolloBodega(this.value)">
          <option value="">Todas las bodegas</option>
          ${[...GI_BODEGAS].sort().map(b => `<option value="${b}">${b}</option>`).join('')}
        </select>
      </div>
      <div class="filter-group">
        <label>Estado</label>
        <select id="gi-r-estado" onchange="giFilterRolloEstado(this.value)">
          <option value="">Todos</option>
          <option value="Vigente">✅ Vigente</option>
          <option value="Por vencer (≤ 30 días)">⚠️ Por vencer (≤ 30 días)</option>
          <option value="Vencido">❌ Vencido</option>
        </select>
      </div>
      <div class="filters-actions">
        <button class="btn-apply" onclick="giApplyRollosAndRender()">✦ Aplicar</button>
        <button class="btn-reset" onclick="giResetRolloFilters()">↺ Reset</button>
      </div>
    </div>
  </div>

  <!-- Gráficas rollos -->
  <div id="gi-charts-ro" style="display:grid;grid-template-columns:repeat(3,1fr);gap:18px;margin-bottom:32px;">
    <div class="chart-card"><div class="chart-header"><div class="chart-title">Estado de Rollos</div><div class="chart-sub">Distribución por estado de vencimiento</div></div><div class="chart-wrap"><canvas id="gi-chart-ro-estado"></canvas></div></div>
    <div class="chart-card"><div class="chart-header"><div class="chart-title">Rollos por Bodega</div><div class="chart-sub">Top bodegas con más stock</div></div><div class="chart-wrap"><canvas id="gi-chart-ro-bodega"></canvas></div></div>
    <div class="chart-card"><div class="chart-header"><div class="chart-title">Días Restantes</div><div class="chart-sub">Distribución de días para vencimiento</div></div><div class="chart-wrap"><canvas id="gi-chart-ro-dias"></canvas></div></div>
  </div>

  <!-- Tabla rollos -->
  <div class="section-label fade-up" style="color:#99D1FC;">Rollos en Almacén</div>
  <div style="background:linear-gradient(145deg,rgba(8,18,28,.95),rgba(6,14,22,.9));border:1px solid rgba(153,209,252,.12);border-radius:18px;overflow:hidden;margin-bottom:32px;box-shadow:0 4px 20px rgba(0,0,0,.4);">
    <div style="display:flex;align-items:center;justify-content:space-between;padding:16px 20px;border-bottom:1px solid rgba(153,209,252,.08);">
      <div>
        <div style="font-family:'Syne',sans-serif;font-size:13px;font-weight:700;color:#f1f5f9;">Tabla 3 — Almacén · Cantidad · Fecha Ingreso · Vencimiento · Días Restantes</div>
        <div style="font-size:11px;color:#7A7674;margin-top:2px;" id="gi-t3-count">0 lotes</div>
      </div>
      <button onclick="giExportT3()" style="background:rgba(153,209,252,.08);border:1px solid rgba(153,209,252,.2);color:#99D1FC;padding:7px 14px;border-radius:8px;cursor:pointer;font-size:12px;font-family:'Outfit',sans-serif;">⬇ Excel</button>
    </div>
    <div id="gi-t3-wrap" style="overflow-x:auto;max-height:500px;"></div>
    <div style="display:flex;align-items:center;justify-content:flex-end;padding:10px 20px;border-top:1px solid rgba(153,209,252,.07);">
      <div id="gi-t3-pag" style="display:flex;gap:6px;flex-wrap:wrap;"></div>
    </div>
  </div>

</div>
`;
}

// ══════════════════════════════════════════════════════════════════
//  BIND FILTROS
// ══════════════════════════════════════════════════════════════════

function _giBindFilters() {
  // Already wired via inline onchange/oninput — no additional binding needed
}

window.giFilterSerial  = v => { GIF_SERIAL = v; };
window.giFilterReuso   = v => { GIF_REUSO  = v; };
window.giFilterEstado  = v => { GIF_ESTADO = v; };

window.giApplyAndRender = () => {
  GIF_SERIAL = (document.getElementById('gi-f-serial')?.value || '').trim();
  GIF_REUSO  =  document.getElementById('gi-f-reuso')?.value  || '';
  GIF_ESTADO =  document.getElementById('gi-f-estado')?.value || '';
  _giApplyFiltersDF();
  _giApplyFiltersACC();
  _giRenderDFSection();
};

window.giResetFilters = () => {
  GIF_SERIAL = ''; GIF_REUSO = ''; GIF_ESTADO = '';
  ['gi-f-serial','gi-f-reuso','gi-f-estado'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = '';
  });
  _giApplyFiltersDF();
  _giApplyFiltersACC();
  _giRenderDFSection();
};

window.giFilterRolloBodega = v => { GIR_BODEGA = v; };
window.giFilterRolloEstado = v => { GIR_ESTADO = v; };

window.giApplyRollosAndRender = () => {
  GIR_BODEGA = document.getElementById('gi-r-bodega')?.value || '';
  GIR_ESTADO = document.getElementById('gi-r-estado')?.value || '';
  _giApplyFiltersRO();
  _giRenderRollosSection();
};

window.giResetRolloFilters = () => {
  GIR_BODEGA = ''; GIR_ESTADO = '';
  ['gi-r-bodega','gi-r-estado'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = '';
  });
  _giApplyFiltersRO();
  _giRenderRollosSection();
};

// ══════════════════════════════════════════════════════════════════
//  RENDER COMPLETO
// ══════════════════════════════════════════════════════════════════

function _giRenderAll() {
  _giRenderKPIs();
  _giRenderDFSection();
  _giRenderKPIsRollos();
  _giRenderRollosSection();
}

// ── KPIs datáfonos ─────────────────────────────────────────────────
function _giRenderKPIs() {
  const rows = GI_DF_ALL;
  const total     = rows.length;
  const vigente   = rows.filter(r => r.estadoGarantia === 'Vigente').length;
  const porVencer = rows.filter(r => r.estadoGarantia === 'Por vencer (≤ 60 días)').length;
  const vencida   = rows.filter(r => r.estadoGarantia === 'Vencida').length;
  const sinFecha  = rows.filter(r => r.estadoGarantia === 'Sin fecha').length;
  const reusados  = rows.filter(r => r.reusado).length;
  const kpis = [
    { label:'Total Datáfonos', val: total,     color:'#DFFF61', icon:'📱' },
    { label:'Garantía Vigente', val: vigente,   color:'#B0F2AE', icon:'✅' },
    { label:'Por Vencer ≤60d', val: porVencer,  color:'#FFC04D', icon:'⚠️' },
    { label:'Garantía Vencida', val: vencida,   color:'#FF5C5C', icon:'❌' },
    { label:'Sin Fecha Garantía',val: sinFecha,  color:'#94a3b8', icon:'—'  },
    { label:'Equipos Reusados',  val: reusados,  color:'#C084FC', icon:'♻️' },
  ];
  const strip = document.getElementById('gi-kpi-strip');
  if (!strip) return;
  strip.innerHTML = kpis.map(k => `
    <div style="background:linear-gradient(145deg,rgba(10,26,18,.95),rgba(8,20,14,.9));
      border:1px solid rgba(${k.color === '#DFFF61' ? '223,255,97' : k.color === '#B0F2AE' ? '176,242,174' : k.color === '#FFC04D' ? '255,192,77' : k.color === '#FF5C5C' ? '255,92,92' : k.color === '#C084FC' ? '192,132,252' : '148,163,184'},.2);
      border-radius:14px;padding:16px 20px;min-width:140px;flex:1;">
      <div style="font-size:20px;margin-bottom:6px;">${k.icon}</div>
      <div style="font-family:'JetBrains Mono',monospace;font-size:22px;font-weight:700;color:${k.color};">${k.val.toLocaleString('es-CO')}</div>
      <div style="font-size:11px;color:#64748b;margin-top:4px;">${k.label}</div>
    </div>`).join('');
}

// ── KPIs rollos ─────────────────────────────────────────────────────
function _giRenderKPIsRollos() {
  const rows = GI_RO_ALL;
  const totalLotes  = rows.length;
  const totalUnids  = rows.reduce((s,r) => s + r.cantidad, 0);
  const vigente     = rows.filter(r => r.estadoRollo === 'Vigente').length;
  const porVencer   = rows.filter(r => r.estadoRollo === 'Por vencer (≤ 30 días)').length;
  const vencido     = rows.filter(r => r.estadoRollo === 'Vencido').length;
  const kpis = [
    { label:'Lotes en Bodegas',    val: totalLotes,  color:'#99D1FC', icon:'📦' },
    { label:'Rollos Totales',      val: totalUnids,  color:'#DFFF61', icon:'🎞️' },
    { label:'Lotes Vigentes',      val: vigente,     color:'#B0F2AE', icon:'✅' },
    { label:'Lotes Por Vencer',    val: porVencer,   color:'#FFC04D', icon:'⚠️' },
    { label:'Lotes Vencidos',      val: vencido,     color:'#FF5C5C', icon:'❌' },
  ];
  const strip = document.getElementById('gi-kpi-rollos');
  if (!strip) return;
  strip.innerHTML = kpis.map(k => `
    <div style="background:linear-gradient(145deg,rgba(8,18,28,.95),rgba(6,14,22,.9));
      border:1px solid rgba(153,209,252,.2);
      border-radius:14px;padding:16px 20px;min-width:140px;flex:1;">
      <div style="font-size:20px;margin-bottom:6px;">${k.icon}</div>
      <div style="font-family:'JetBrains Mono',monospace;font-size:22px;font-weight:700;color:${k.color};">${k.val.toLocaleString('es-CO')}</div>
      <div style="font-size:11px;color:#64748b;margin-top:4px;">${k.label}</div>
    </div>`).join('');
}

// ══════════════════════════════════════════════════════════════════
//  RENDER SECCIÓN DATÁFONOS
// ══════════════════════════════════════════════════════════════════

function _giRenderDFSection() {
  _giRenderChartsDF();
  _giRenderT1();
  _giRenderT2();
}

function _giRenderChartsDF() {
  const rows = GI_DF_FILTERED;

  // Chart 1: Estado de garantía (donut)
  const estadoMap = { 'Vigente':0, 'Por vencer (≤ 60 días)':0, 'Vencida':0, 'Sin fecha':0 };
  rows.forEach(r => estadoMap[r.estadoGarantia] = (estadoMap[r.estadoGarantia]||0)+1);
  _giMkChart('gi-chart-estado', {
    type: 'doughnut',
    data: {
      labels: Object.keys(estadoMap),
      datasets: [{ data: Object.values(estadoMap),
        backgroundColor:[GI_COLORS.vigente, GI_COLORS.porVencer, GI_COLORS.vencida, GI_COLORS.sinFecha],
        borderWidth:0, hoverOffset:8 }]
    },
    options: { plugins:{ legend:{ labels:{ color:'#94a3b8', font:{ family:"'Outfit'" } } } }, cutout:'65%' }
  });

  // Chart 2: Garantía por referencia (bar apilada)
  const refs = [...GI_REFS_DATAFONO];
  const vigData = refs.map(ref => rows.filter(r=>r.referencia===ref&&r.estadoGarantia==='Vigente').length);
  const pvData  = rows.filter(r=>r.estadoGarantia==='Por vencer (≤ 60 días)');
  const vData   = rows.filter(r=>r.estadoGarantia==='Vencida');
  _giMkChart('gi-chart-ref-estado', {
    type:'bar',
    data:{
      labels: refs.map(r => r.replace(' - INGENICO','').substring(0,20)),
      datasets:[
        { label:'Vigente',    data: refs.map(ref=>rows.filter(r=>r.referencia===ref&&r.estadoGarantia==='Vigente').length), backgroundColor:GI_COLORS.vigente },
        { label:'Por vencer', data: refs.map(ref=>rows.filter(r=>r.referencia===ref&&r.estadoGarantia==='Por vencer (≤ 60 días)').length), backgroundColor:GI_COLORS.porVencer },
        { label:'Vencida',    data: refs.map(ref=>rows.filter(r=>r.referencia===ref&&r.estadoGarantia==='Vencida').length), backgroundColor:GI_COLORS.vencida },
      ]
    },
    options:{
      indexAxis:'y', plugins:{ legend:{ labels:{ color:'#94a3b8' } } },
      scales:{ x:{ stacked:true, ticks:{ color:'#64748b' }, grid:{ color:'rgba(255,255,255,.05)' } },
               y:{ stacked:true, ticks:{ color:'#94a3b8', font:{ size:10 } }, grid:{ display:false } } }
    }
  });

  // Chart 3: Nuevo vs Usado (pie)
  const nuevo = rows.filter(r=>!r.reusado).length;
  const usado = rows.filter(r=> r.reusado).length;
  _giMkChart('gi-chart-reuso', {
    type:'doughnut',
    data:{ labels:['Nuevo','Usado (Return)'], datasets:[{ data:[nuevo,usado], backgroundColor:[GI_COLORS.nuevo, GI_COLORS.usado], borderWidth:0, hoverOffset:8 }] },
    options:{ plugins:{ legend:{ labels:{ color:'#94a3b8' } } }, cutout:'60%' }
  });
}

function _giRenderT1() {
  const rows = GI_DF_FILTERED;
  const total = rows.length;
  const start = (GI_DF_PAGE - 1) * GI_PAGE_SIZE;
  const page  = rows.slice(start, start + GI_PAGE_SIZE);

  const countEl = document.getElementById('gi-t1-count');
  if (countEl) countEl.textContent = `${total.toLocaleString('es-CO')} registros`;

  const wrap = document.getElementById('gi-t1-wrap');
  if (!wrap) return;

  if (!total) { wrap.innerHTML = '<div style="padding:32px;text-align:center;color:#64748b;">Sin registros para los filtros seleccionados</div>'; return; }

  const TH = (t, align='left') =>
    `<th style="padding:10px 14px;font-family:'Syne',sans-serif;font-size:11px;font-weight:700;color:#94a3b8;white-space:nowrap;text-align:${align};border-bottom:1px solid rgba(176,242,174,.1);background:rgba(0,0,0,.3);position:sticky;top:0;z-index:1;">${t}</th>`;

  const TD = (v, color='#e2e8f0', align='left', mono=false) =>
    `<td style="padding:9px 14px;font-size:12px;color:${color};text-align:${align};${mono?"font-family:'JetBrains Mono',monospace;":''}">${v}</td>`;

  const rows_html = page.map(r => {
    const diasG = r.diasGarantia !== null ? (r.diasGarantia < 0 ? `<span style="color:#FF5C5C">${r.diasGarantia}d</span>` : `<span style="color:${r.diasGarantia<=60?'#FFC04D':'#B0F2AE'}">${r.diasGarantia}d</span>`) : '—';
    const diasI = r.diasInventario !== null ? `${r.diasInventario}d` : '—';
    const cond  = r.reusado ? `<span style="color:#C084FC;font-size:11px;">♻️ Usado</span>` : `<span style="color:#99D1FC;font-size:11px;">🆕 Nuevo</span>`;
    return `<tr style="border-bottom:1px solid rgba(255,255,255,.04);" onmouseover="this.style.background='rgba(176,242,174,.04)'" onmouseout="this.style.background=''">
      ${TD(r.referencia.replace(' - INGENICO',''), '#f1f5f9')}
      ${TD(r.serial || '—', '#B0F2AE', 'left', true)}
      ${TD(_giFmtDate(r.fc), '#94a3b8')}
      ${TD(_giFmtDate(r.fr), '#94a3b8')}
      ${TD(_giFmtDate(r.fg), '#f1f5f9')}
      ${TD(r.pu ? _giFmtCOP(r.pu) : '—', '#DFFF61', 'right', true)}
      <td style="padding:9px 14px;">${_giEstadoBadge(r.estadoGarantia)}</td>
      <td style="padding:9px 14px;font-size:12px;text-align:right;">${diasG}</td>
      <td style="padding:9px 14px;font-size:12px;text-align:right;color:#99D1FC;">${diasI}</td>
      <td style="padding:9px 14px;">${cond}</td>
    </tr>`;
  }).join('');

  wrap.innerHTML = `<table style="width:100%;border-collapse:collapse;">
    <thead><tr>${TH('Referencia')}${TH('Núm. Serie')}${TH('F. Compra')}${TH('F. Recibo')}${TH('F. Garantía')}${TH('Precio Unit.','right')}${TH('Estado Garantía')}${TH('Días Restantes','right')}${TH('Días en Inventario','right')}${TH('Condición')}</tr></thead>
    <tbody>${rows_html}</tbody>
  </table>`;

  _giRenderPag('gi-t1-pag', total, GI_DF_PAGE, GI_PAGE_SIZE, p => { GI_DF_PAGE = p; _giRenderT1(); }, '#B0F2AE');
}

function _giRenderT2() {
  const rows = GI_ACC_FILTERED;
  const countEl = document.getElementById('gi-t2-count');
  if (countEl) countEl.textContent = `${rows.length} referencias`;

  const wrap = document.getElementById('gi-t2-wrap');
  if (!wrap) return;

  if (!rows.length) { wrap.innerHTML = '<div style="padding:24px;text-align:center;color:#64748b;">Sin datos</div>'; return; }

  const TH = (t, align='left') =>
    `<th style="padding:10px 14px;font-family:'Syne',sans-serif;font-size:11px;font-weight:700;color:#94a3b8;text-align:${align};border-bottom:1px solid rgba(223,255,97,.1);background:rgba(0,0,0,.3);">${t}</th>`;

  const body = rows.map((r,i) => `
    <tr style="border-bottom:1px solid rgba(255,255,255,.04);" onmouseover="this.style.background='rgba(223,255,97,.04)'" onmouseout="this.style.background=''">
      <td style="padding:9px 14px;font-size:12px;color:#f1f5f9;">${r.referencia.replace(' - INGENICO','')}</td>
      <td style="padding:9px 14px;font-family:'JetBrains Mono',monospace;font-size:13px;color:#DFFF61;text-align:right;">${_giFmtCOP(r.acumulado)}</td>
      <td style="padding:9px 14px;font-family:'JetBrains Mono',monospace;font-size:13px;color:#99D1FC;text-align:right;">${r.unidades.toLocaleString('es-CO')}</td>
    </tr>`).join('');

  wrap.innerHTML = `<table style="width:100%;border-collapse:collapse;">
    <thead><tr>${TH('Referencia')}${TH('Precio Acumulado','right')}${TH('Unidades Totales','right')}</tr></thead>
    <tbody>${body}</tbody>
  </table>`;

  // Gráfica acumulado
  _giMkChart('gi-chart-acumulado', {
    type: 'bar',
    data: {
      labels: rows.map(r => r.referencia.replace(' - INGENICO','').substring(0,22)),
      datasets: [{ label:'Precio Acumulado COP', data: rows.map(r => r.acumulado),
        backgroundColor: GI_COLORS.palette.slice(0, rows.length), borderWidth:0, borderRadius:6 }]
    },
    options: {
      indexAxis:'y',
      plugins:{ legend:{ display:false },
        tooltip:{ callbacks:{ label: ctx => ' ' + _giFmtCOP(ctx.raw) } } },
      scales:{ x:{ ticks:{ color:'#64748b', callback: v => _giFmtCOP(v) }, grid:{ color:'rgba(255,255,255,.05)' } },
               y:{ ticks:{ color:'#94a3b8', font:{ size:11 } }, grid:{ display:false } } }
    }
  });
}

// ══════════════════════════════════════════════════════════════════
//  RENDER SECCIÓN ROLLOS
// ══════════════════════════════════════════════════════════════════

function _giRenderRollosSection() {
  _giRenderChartsRollos();
  _giRenderT3();
}

function _giRenderChartsRollos() {
  const rows = GI_RO_FILTERED;

  // Chart: Estado rollos
  const estadoMap = { 'Vigente':0, 'Por vencer (≤ 30 días)':0, 'Vencido':0 };
  rows.forEach(r => estadoMap[r.estadoRollo] = (estadoMap[r.estadoRollo]||0) + r.cantidad);
  _giMkChart('gi-chart-ro-estado', {
    type:'doughnut',
    data:{ labels:Object.keys(estadoMap), datasets:[{ data:Object.values(estadoMap),
      backgroundColor:['#B0F2AE','#FFC04D','#FF5C5C'], borderWidth:0, hoverOffset:8 }] },
    options:{ plugins:{ legend:{ labels:{ color:'#94a3b8' } } }, cutout:'60%' }
  });

  // Chart: Top bodegas (bar)
  const bodMap = {};
  rows.forEach(r => { bodMap[r.ubicacion] = (bodMap[r.ubicacion]||0) + r.cantidad; });
  const sorted = Object.entries(bodMap).sort((a,b)=>b[1]-a[1]).slice(0,10);
  _giMkChart('gi-chart-ro-bodega', {
    type:'bar',
    data:{ labels: sorted.map(([b]) => b.replace('ALMACEN WOMPI ','').replace('ALMACEN ','')),
      datasets:[{ label:'Rollos', data: sorted.map(([,v])=>v), backgroundColor:'#99D1FC', borderWidth:0, borderRadius:6 }] },
    options:{
      indexAxis:'y',
      plugins:{ legend:{ display:false } },
      scales:{ x:{ ticks:{ color:'#64748b' }, grid:{ color:'rgba(255,255,255,.05)' } },
               y:{ ticks:{ color:'#94a3b8', font:{ size:10 } }, grid:{ display:false } } }
    }
  });

  // Chart: Distribución días restantes (histogram buckets)
  const buckets = { '<0 (Vencido)':0, '0–30d':0, '31–60d':0, '61–90d':0, '>90d':0 };
  rows.forEach(r => {
    const d = r.diasRestantes;
    if (d < 0)       buckets['<0 (Vencido)'] += r.cantidad;
    else if (d<=30)  buckets['0–30d']  += r.cantidad;
    else if (d<=60)  buckets['31–60d'] += r.cantidad;
    else if (d<=90)  buckets['61–90d'] += r.cantidad;
    else             buckets['>90d']   += r.cantidad;
  });
  _giMkChart('gi-chart-ro-dias', {
    type:'bar',
    data:{ labels:Object.keys(buckets), datasets:[{ label:'Rollos', data:Object.values(buckets),
      backgroundColor:['#FF5C5C','#FFC04D','#FFC04D','#B0F2AE','#B0F2AE'], borderWidth:0, borderRadius:6 }] },
    options:{ plugins:{ legend:{ display:false } },
      scales:{ x:{ ticks:{ color:'#94a3b8' }, grid:{ display:false } },
               y:{ ticks:{ color:'#64748b' }, grid:{ color:'rgba(255,255,255,.05)' } } } }
  });
}

function _giRenderT3() {
  const rows = GI_RO_FILTERED;
  const total = rows.length;
  const start = (GI_RO_PAGE - 1) * GI_PAGE_SIZE;
  const page  = rows.slice(start, start + GI_PAGE_SIZE);

  const countEl = document.getElementById('gi-t3-count');
  if (countEl) countEl.textContent = `${total.toLocaleString('es-CO')} lotes`;

  const wrap = document.getElementById('gi-t3-wrap');
  if (!wrap) return;

  if (!total) { wrap.innerHTML = '<div style="padding:32px;text-align:center;color:#64748b;">Sin rollos en las bodegas con los filtros actuales</div>'; return; }

  const TH = (t, align='left') =>
    `<th style="padding:10px 14px;font-family:'Syne',sans-serif;font-size:11px;font-weight:700;color:#94a3b8;white-space:nowrap;text-align:${align};border-bottom:1px solid rgba(153,209,252,.1);background:rgba(0,0,0,.3);position:sticky;top:0;z-index:1;">${t}</th>`;

  const estadoBadge = (e) => {
    const map = {
      'Vigente':                { bg:'rgba(176,242,174,.15)', color:'#B0F2AE', icon:'✅' },
      'Por vencer (≤ 30 días)': { bg:'rgba(255,192,77,.15)',  color:'#FFC04D', icon:'⚠️' },
      'Vencido':                { bg:'rgba(255,92,92,.15)',   color:'#FF5C5C', icon:'❌' },
    };
    const s = map[e] || { bg:'rgba(100,116,139,.15)', color:'#94a3b8', icon:'—' };
    return `<span style="display:inline-flex;align-items:center;gap:4px;padding:3px 10px;border-radius:20px;background:${s.bg};color:${s.color};font-size:11px;font-weight:600;">${s.icon} ${e}</span>`;
  };

  const rows_html = page.map(r => {
    const diasColor = r.diasRestantes < 0 ? '#FF5C5C' : r.diasRestantes <= 30 ? '#FFC04D' : '#B0F2AE';
    return `<tr style="border-bottom:1px solid rgba(255,255,255,.04);" onmouseover="this.style.background='rgba(153,209,252,.04)'" onmouseout="this.style.background=''">
      <td style="padding:9px 14px;font-size:12px;color:#f1f5f9;max-width:220px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;" title="${r.ubicacion}">${r.ubicacion}</td>
      <td style="padding:9px 14px;font-size:12px;color:#99D1FC;">${r.nombre}</td>
      <td style="padding:9px 14px;font-family:'JetBrains Mono',monospace;font-size:12px;color:#DFFF61;text-align:right;">${r.cantidad.toLocaleString('es-CO')}</td>
      <td style="padding:9px 14px;font-size:12px;color:#94a3b8;">${_giFmtDate(r.fechaIngreso)}</td>
      <td style="padding:9px 14px;font-size:12px;color:#f1f5f9;">${_giFmtDate(r.fechaVence)}</td>
      <td style="padding:9px 14px;">${estadoBadge(r.estadoRollo)}</td>
      <td style="padding:9px 14px;font-family:'JetBrains Mono',monospace;font-size:13px;font-weight:700;color:${diasColor};text-align:right;">${r.diasRestantes !== null ? r.diasRestantes + 'd' : '—'}</td>
    </tr>`;
  }).join('');

  wrap.innerHTML = `<table style="width:100%;border-collapse:collapse;">
    <thead><tr>${TH('Nombre del Almacén')}${TH('Lote (Nombre)')}${TH('Cantidad','right')}${TH('Fecha Ingreso')}${TH('Fecha Vencimiento')}${TH('Estado')}${TH('Días Restantes','right')}</tr></thead>
    <tbody>${rows_html}</tbody>
  </table>`;

  _giRenderPag('gi-t3-pag', total, GI_RO_PAGE, GI_PAGE_SIZE, p => { GI_RO_PAGE = p; _giRenderT3(); }, '#99D1FC');
}

// ══════════════════════════════════════════════════════════════════
//  PAGINACIÓN
// ══════════════════════════════════════════════════════════════════

function _giRenderPag(containerId, total, current, pageSize, onPage, accentColor) {
  const pages = Math.ceil(total / pageSize);
  const el = document.getElementById(containerId);
  if (!el || pages <= 1) { if (el) el.innerHTML = ''; return; }

  const btnStyle = (active) =>
    `style="padding:5px 11px;border-radius:6px;font-size:12px;cursor:pointer;border:1px solid rgba(255,255,255,.1);
    background:${active ? accentColor : 'rgba(255,255,255,.05)'};color:${active ? '#0a1a12' : '#94a3b8'};font-family:'Outfit',sans-serif;"`;

  let html = '';
  if (current > 1)   html += `<button ${btnStyle(false)} onclick="(${onPage.toString()})(${current-1})">‹</button>`;
  const range = [];
  for (let i = Math.max(1,current-2); i<=Math.min(pages,current+2); i++) range.push(i);
  range.forEach(p => html += `<button ${btnStyle(p===current)} onclick="(${onPage.toString()})(${p})">${p}</button>`);
  if (current < pages) html += `<button ${btnStyle(false)} onclick="(${onPage.toString()})(${current+1})">›</button>`;
  el.innerHTML = html;
}

// ══════════════════════════════════════════════════════════════════
//  EXPORTAR EXCEL
// ══════════════════════════════════════════════════════════════════

window.giExportT1 = () => {
  if (!window.XLSX) { alert('Librería Excel no disponible'); return; }
  const data = GI_DF_FILTERED.map(r => ({
    'Referencia': r.referencia,
    'Número de Serie': r.serial,
    'Fecha Compra': r.fc ? _giFmtDate(r.fc) : '',
    'Fecha Recibo': r.fr ? _giFmtDate(r.fr) : '',
    'Fecha Garantía': r.fg ? _giFmtDate(r.fg) : '',
    'Precio Unitario COP': r.pu || '',
    'Estado Garantía': r.estadoGarantia,
    'Días Restantes Garantía': r.diasGarantia !== null ? r.diasGarantia : '',
    'Días en Inventario': r.diasInventario !== null ? r.diasInventario : '',
    'Condición': r.reusado ? 'Usado (Return)' : 'Nuevo',
  }));
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Garantía Datáfonos');
  XLSX.writeFile(wb, 'garantia_datafonos.xlsx');
};

window.giExportT2 = () => {
  if (!window.XLSX) { alert('Librería Excel no disponible'); return; }
  const data = GI_ACC_FILTERED.map(r => ({
    'Referencia': r.referencia,
    'Precio Acumulado COP': r.acumulado,
    'Unidades Totales': r.unidades,
  }));
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Acumulado Referencia');
  XLSX.writeFile(wb, 'acumulado_referencia.xlsx');
};

window.giExportT3 = () => {
  if (!window.XLSX) { alert('Librería Excel no disponible'); return; }
  const data = GI_RO_FILTERED.map(r => ({
    'Nombre del Almacén': r.ubicacion,
    'Lote (Nombre)': r.nombre,
    'Cantidad': r.cantidad,
    'Fecha Ingreso': _giFmtDate(r.fechaIngreso),
    'Fecha Vencimiento': _giFmtDate(r.fechaVence),
    'Estado': r.estadoRollo,
    'Días Restantes': r.diasRestantes !== null ? r.diasRestantes : '',
  }));
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Rollos Almacén');
  XLSX.writeFile(wb, 'rollos_almacen.xlsx');
};