/**
 * ╔══════════════════════════════════════════════════════════════════╗
 * ║  inventario.js — Dashboard Inventario Wompi v2                  ║
 * ║  KPIs unificados · Gráficas · Tabla Top Referencias             ║
 * ╚══════════════════════════════════════════════════════════════════╝
 */

'use strict';

// ── Estado global ─────────────────────────────────────────────────
let INV_RAW      = null;
let INV_FILTERED = [];
let INV_CHARTS   = {};   // instancias Chart.js activas

// ── 42 Bodegas Wompi ─────────────────────────────────────────────
const INV_BODEGAS = new Set([
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

// ── Categorización ────────────────────────────────────────────────
function invCategoria(nombre) {
  if (!nombre) return 'KIT POP VP';
  let r = nombre.toUpperCase().trim()
    .replace(/\u00A0/g, ' ').replace(/  +/g, ' ')
    .replace('DX 4000','DX4000').replace('EX 4000','EX4000');

  if (r.includes('ROLLO'))                                                       return 'Rollos';
  if (r.includes('PINPAD')||r.includes('PIN PAD')||r.includes('DESK 1700'))     return 'Pin pad';
  if (r.includes('FORRO'))                                                       return 'Forros';
  if (r.includes('PROTECTOR')||r.includes('PANTALLA')||r.includes('VIDRIO')||
      r.includes('TEMPLADO')||r.includes('MICA'))                                return 'Accesorios';
  if (r.includes('MAGIC BOX')||r.includes('MAGICBOX'))                          return 'Accesorios';
  if (r.includes('USB')||r.includes('RS232')||r.includes('CONVERTER'))          return 'Accesorios';
  if (r.includes('SIM'))                                                         return 'SIM';
  if (r.includes('KIT')||r.includes('STICKER'))                                 return 'KIT POP VP';
  if (r.includes('DATAFONO')||r.includes('DX4000')||
      r.includes('EX4000')||r.includes('EX6000'))                               return 'Datáfonos';
  return 'KIT POP VP';
}

// ── Negocio ───────────────────────────────────────────────────────
function invNegocio(subtipo) {
  const s = (subtipo || '').trim().toUpperCase();
  if (s === 'WOMPI VP' || s === 'EQUIPO VP' || s === 'VP') return 'VP';
  return 'CB';
}

// ── Patrón GW ─────────────────────────────────────────────────────
const GW_RE = /^GW\d+$/i;

// ── Helpers numéricos ─────────────────────────────────────────────
function sumCantidad(rows) {
  return rows.reduce((acc, r) => acc + (parseInt(r['Cantidad']) || 0), 0);
}
function fmtN(n)   { return n.toLocaleString('es-CO'); }
function fmtPct(num, den) {
  if (!den) return '0.0%';
  return (num / den * 100).toFixed(1) + '%';
}

// ── Paleta de colores ─────────────────────────────────────────────
const INV_PALETTE = {
  bodega:   '#B0F2AE',
  comercio: '#99D1FC',
  tecnico:  '#FFC04D',
  gestores: '#C084FC',
  ingenico: '#F87171',
  opl:      '#FB923C',
  total:    '#DFFF61',
};

// ══════════════════════════════════════════════════════════════════
//  CARGA DEL JSON.GZ
// ══════════════════════════════════════════════════════════════════
async function loadInventarioData() {
  if (INV_RAW) return;
  try {
    const res = await fetch('stock_wompi_filtrado.json.gz?t=' + Date.now());
    if (!res.ok) throw new Error('HTTP ' + res.status);
    const buf    = await res.arrayBuffer();
    const ds     = new DecompressionStream('gzip');
    const writer = ds.writable.getWriter();
    writer.write(new Uint8Array(buf));
    writer.close();
    const reader = ds.readable.getReader();
    const chunks = [];
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      chunks.push(value);
    }
    const total  = chunks.reduce((s, c) => s + c.length, 0);
    const merged = new Uint8Array(total);
    let off = 0;
    for (const c of chunks) { merged.set(c, off); off += c.length; }
    INV_RAW = JSON.parse(new TextDecoder().decode(merged));
    console.log('[Inventario] ' + INV_RAW.length + ' filas cargadas');
  } catch (e) {
    console.error('[Inventario] Error cargando datos:', e);
    INV_RAW = [];
  }
}

// ══════════════════════════════════════════════════════════════════
//  POBLAR FILTROS
// ══════════════════════════════════════════════════════════════════
function _invPopulateFilters() {
  if (!INV_RAW || !INV_RAW.length) return;
  const cats = ['Todas','Rollos','Pin pad','Forros','Accesorios','SIM','KIT POP VP','Datáfonos'];
  _invSetSelect('inv-f-categoria', cats);
  const refs = [...new Set(INV_RAW.map(r => r['Nombre']).filter(Boolean))].sort();
  _invSetSelect('inv-f-referencia', ['Todas', ...refs]);
  const bods = [...INV_BODEGAS].sort();
  _invSetSelect('inv-f-bodega', ['Todas', ...bods]);
}

function _invSetSelect(id, opts) {
  const el = document.getElementById(id);
  if (!el) return;
  const cur = el.value;
  el.innerHTML = opts.map(o => '<option value="' + o + '">' + o + '</option>').join('');
  if (opts.includes(cur)) el.value = cur;
}

// ══════════════════════════════════════════════════════════════════
//  APLICAR FILTROS
// ══════════════════════════════════════════════════════════════════
function invApplyFilters() {
  if (!INV_RAW) { INV_FILTERED = []; return; }

  const negocio    = (document.getElementById('inv-f-negocio')?.value    || '');
  const categoria  = (document.getElementById('inv-f-categoria')?.value  || '');
  const referencia = (document.getElementById('inv-f-referencia')?.value || '');
  const bodega     = (document.getElementById('inv-f-bodega')?.value     || '');

  INV_FILTERED = INV_RAW.filter(r => {
    if (negocio    && negocio    !== 'Todos' && invNegocio(r['Subtipo'])                 !== negocio)    return false;
    if (categoria  && categoria  !== 'Todas' && invCategoria(r['Nombre'])                !== categoria)  return false;
    if (referencia && referencia !== 'Todas' && r['Nombre']                              !== referencia) return false;
    if (bodega     && bodega     !== 'Todas' && (r['Nombre de la ubicación']||'').trim() !== bodega.trim()) return false;
    return true;
  });

  _invRenderAll();
}

window.invApplyFilters = invApplyFilters;
window.invResetFilters = function() {
  ['inv-f-negocio','inv-f-categoria','inv-f-referencia','inv-f-bodega'].forEach(function(id) {
    const el = document.getElementById(id);
    if (el) el.selectedIndex = 0;
  });
  invApplyFilters();
};

// ══════════════════════════════════════════════════════════════════
//  RENDER ALL
// ══════════════════════════════════════════════════════════════════
function _invRenderAll() {
  _invRenderKPIs();
  _invRenderCharts();
  _invRenderTable();
}

// ══════════════════════════════════════════════════════════════════
//  RENDER KPIs — UNIFICADOS
// ══════════════════════════════════════════════════════════════════
function _invRenderKPIs() {
  const rows = INV_FILTERED;
  const total = sumCantidad(rows);

  const unBodega   = sumCantidad(rows.filter(r => INV_BODEGAS.has((r['Nombre de la ubicación']||'').trim())));
  const unComercio = sumCantidad(rows.filter(r => (r['Tipo de ubicación']||'').trim() === 'Site'));
  const unTecnico  = sumCantidad(rows.filter(r => (r['Tipo de ubicación']||'').trim() === 'Staff'));
  const unGW       = sumCantidad(rows.filter(r => GW_RE.test((r['Código de ubicación']||'').trim())));
  const unIngenico = sumCantidad(rows.filter(r => (r['Nombre de la ubicación']||'').trim() === 'ALMACEN INGENICO - PROVEEDOR WOMPI'));
  const unOPL      = sumCantidad(rows.filter(r => (r['Tipo de ubicación']||'').trim() === 'Supplier'));

  const kpis = [
    { label:'TOTAL INVENTARIO', value:fmtN(total),      sub1:'100% del stock',               sub2:'42 bodegas Wompi',       color:INV_PALETTE.total,    icon:'📦', wide:true },
    { label:'EN BODEGA',        value:fmtN(unBodega),   sub1:fmtPct(unBodega,total),         sub2:fmtN(unBodega)+' uds',   color:INV_PALETTE.bodega,   icon:'🏪' },
    { label:'EN COMERCIO',      value:fmtN(unComercio), sub1:fmtPct(unComercio,total),       sub2:fmtN(unComercio)+' uds', color:INV_PALETTE.comercio, icon:'🏬' },
    { label:'TÉC. LINEACOM',    value:fmtN(unTecnico),  sub1:fmtPct(unTecnico,total),        sub2:fmtN(unTecnico)+' uds',  color:INV_PALETTE.tecnico,  icon:'🔧' },
    { label:'GEST. & EMPL.',    value:fmtN(unGW),       sub1:fmtPct(unGW,total),             sub2:fmtN(unGW)+' uds',       color:INV_PALETTE.gestores, icon:'👤' },
    { label:'INGENICO',         value:fmtN(unIngenico), sub1:fmtPct(unIngenico,total),       sub2:fmtN(unIngenico)+' uds', color:INV_PALETTE.ingenico, icon:'🔌' },
    { label:'OPL',              value:fmtN(unOPL),      sub1:fmtPct(unOPL,total),            sub2:fmtN(unOPL)+' uds',      color:INV_PALETTE.opl,      icon:'🚚' },
  ];

  const grid = document.getElementById('inv-kpi-grid');
  if (!grid) return;

  grid.innerHTML = kpis.map(function(k) {
    return (
      '<div class="kpi-card inv-kpi-v2" style="' +
        'border-top:3px solid ' + k.color + ';' +
        'background:linear-gradient(160deg,rgba(255,255,255,0.04) 0%,rgba(255,255,255,0.01) 100%);' +
        'border-radius:16px;padding:20px 22px 18px;position:relative;overflow:hidden;' +
        'transition:transform 0.2s ease,box-shadow 0.2s ease;' +
        (k.wide ? 'grid-column:span 2;' : '') +
      '" ' +
      'onmouseover="this.style.transform=\'translateY(-3px)\';this.style.boxShadow=\'0 12px 40px rgba(0,0,0,0.35)\'" ' +
      'onmouseout="this.style.transform=\'\';this.style.boxShadow=\'\'">' +
        // Glow accent
        '<div style="position:absolute;top:-20px;right:-20px;width:80px;height:80px;border-radius:50%;background:' + k.color + ';opacity:0.07;pointer-events:none;"></div>' +
        // Header row
        '<div style="display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:10px;">' +
          '<span style="font-size:10px;font-weight:700;letter-spacing:1px;color:var(--muted);text-transform:uppercase;line-height:1.3;">' + k.label + '</span>' +
          '<span style="font-size:18px;line-height:1;">' + k.icon + '</span>' +
        '</div>' +
        // Main value
        '<div style="font-size:32px;font-weight:800;color:' + k.color + ';font-family:\'Syne\',sans-serif;line-height:1;margin-bottom:12px;letter-spacing:-0.5px;">' + k.value + '</div>' +
        // Sub tags
        '<div style="display:flex;align-items:center;gap:7px;flex-wrap:wrap;">' +
          '<span style="display:inline-block;background:' + k.color + '22;color:' + k.color + ';font-size:11px;font-weight:700;padding:2px 9px;border-radius:20px;font-family:\'JetBrains Mono\',monospace;">' + k.sub1 + '</span>' +
          '<span style="font-size:11px;color:var(--muted);">' + k.sub2 + '</span>' +
        '</div>' +
      '</div>'
    );
  }).join('');
}

// ══════════════════════════════════════════════════════════════════
//  RENDER CHARTS
// ══════════════════════════════════════════════════════════════════
function _invDestroyCharts() {
  Object.values(INV_CHARTS).forEach(function(c) { try { c.destroy(); } catch(_) {} });
  INV_CHARTS = {};
}

function _invRenderCharts() {
  const rows  = INV_FILTERED;
  const total = sumCantidad(rows);

  const chartsGrid = document.getElementById('inv-charts-grid');
  if (!chartsGrid) return;

  _invDestroyCharts();

  if (!total) {
    chartsGrid.innerHTML = '<div style="color:var(--muted);padding:32px;text-align:center;">Sin datos para mostrar gráficas</div>';
    return;
  }

  // ── Datos ──────────────────────────────────────────────────────
  const unBodega   = sumCantidad(rows.filter(function(r){ return INV_BODEGAS.has((r['Nombre de la ubicación']||'').trim()); }));
  const unComercio = sumCantidad(rows.filter(function(r){ return (r['Tipo de ubicación']||'').trim() === 'Site'; }));
  const unTecnico  = sumCantidad(rows.filter(function(r){ return (r['Tipo de ubicación']||'').trim() === 'Staff'; }));
  const unGW       = sumCantidad(rows.filter(function(r){ return GW_RE.test((r['Código de ubicación']||'').trim()); }));
  const unIngenico = sumCantidad(rows.filter(function(r){ return (r['Nombre de la ubicación']||'').trim() === 'ALMACEN INGENICO - PROVEEDOR WOMPI'; }));
  const unOPL      = sumCantidad(rows.filter(function(r){ return (r['Tipo de ubicación']||'').trim() === 'Supplier'; }));

  const catMap = {};
  rows.forEach(function(r) {
    const cat = invCategoria(r['Nombre']);
    const qty = parseInt(r['Cantidad']) || 0;
    catMap[cat] = (catMap[cat] || 0) + qty;
  });

  const negMap = { CB: 0, VP: 0 };
  rows.forEach(function(r) {
    const neg = invNegocio(r['Subtipo']);
    const qty = parseInt(r['Cantidad']) || 0;
    negMap[neg] = (negMap[neg] || 0) + qty;
  });

  const bodMap = {};
  rows.forEach(function(r) {
    const bod = (r['Nombre de la ubicación'] || 'Sin Nombre').trim();
    const qty = parseInt(r['Cantidad']) || 0;
    bodMap[bod] = (bodMap[bod] || 0) + qty;
  });
  const topBodegas = Object.entries(bodMap).sort(function(a,b){ return b[1]-a[1]; }).slice(0, 10);

  // ── Chart defaults ─────────────────────────────────────────────
  const TOOLTIP_OPTS = {
    backgroundColor: 'rgba(15,23,42,0.95)',
    titleColor: '#B0F2AE', bodyColor: '#e2e8f0',
    borderColor: 'rgba(176,242,174,0.2)', borderWidth: 1, padding: 12,
  };
  const LEGEND_OPTS = { labels: { color: '#94a3b8', font: { family: 'Outfit', size: 12 }, padding: 16 } };
  const XGRID = { color: 'rgba(255,255,255,0.05)' };
  const YTICK = { color: '#94a3b8', font: { family: 'Outfit', size: 12 } };
  const XTICK_MONO = { color: '#64748b', font: { family: 'JetBrains Mono', size: 11 }, callback: function(v){ return v.toLocaleString('es-CO'); } };

  // ── Card builder ───────────────────────────────────────────────
  function makeCard(canvasId, title, sub) {
    const div = document.createElement('div');
    div.style.cssText = 'background:linear-gradient(160deg,rgba(255,255,255,0.04) 0%,rgba(255,255,255,0.01) 100%);border:1px solid rgba(255,255,255,0.06);border-radius:20px;padding:22px 24px 20px;position:relative;overflow:hidden;';
    div.innerHTML =
      '<div style="margin-bottom:16px;">' +
        '<div style="font-size:13px;font-weight:700;color:#f1f5f9;letter-spacing:0.3px;font-family:\'Outfit\',sans-serif;">' + title + '</div>' +
        (sub ? '<div style="font-size:11px;color:var(--muted);margin-top:3px;">' + sub + '</div>' : '') +
      '</div>' +
      '<canvas id="' + canvasId + '" style="max-height:260px;"></canvas>';
    return div;
  }

  chartsGrid.innerHTML = '';

  // ── 1. Donut: distribución por ubicación ──────────────────────
  const c1 = makeCard('inv-c-ubicacion', 'Distribución por Ubicación', 'Unidades según destino');
  chartsGrid.appendChild(c1);
  INV_CHARTS['ubicacion'] = new Chart(document.getElementById('inv-c-ubicacion'), {
    type: 'doughnut',
    data: {
      labels: ['Bodega', 'Comercio', 'Técnico', 'Gest./Empl.', 'Ingenico', 'OPL'],
      datasets: [{
        data: [unBodega, unComercio, unTecnico, unGW, unIngenico, unOPL],
        backgroundColor: [
          INV_PALETTE.bodega+'CC', INV_PALETTE.comercio+'CC',
          INV_PALETTE.tecnico+'CC', INV_PALETTE.gestores+'CC',
          INV_PALETTE.ingenico+'CC', INV_PALETTE.opl+'CC',
        ],
        borderColor: 'rgba(0,0,0,0)', borderWidth: 0, hoverOffset: 8,
      }]
    },
    options: {
      cutout: '68%',
      plugins: {
        legend: LEGEND_OPTS,
        tooltip: Object.assign({}, TOOLTIP_OPTS, {
          callbacks: {
            label: function(ctx) {
              var v = ctx.parsed;
              var pct = total ? (v / total * 100).toFixed(1) : 0;
              return ' ' + ctx.label + ': ' + v.toLocaleString('es-CO') + ' uds (' + pct + '%)';
            }
          }
        })
      }
    }
  });

  // ── 2. Bar horizontal: por categoría ─────────────────────────
  const catEntries = Object.entries(catMap).sort(function(a,b){ return b[1]-a[1]; });
  const catColors  = ['#B0F2AE','#99D1FC','#DFFF61','#C084FC','#FFC04D','#F87171','#FB923C','#7BC8FB'];
  const c2 = makeCard('inv-c-categoria', 'Unidades por Categoría', 'Desglose por tipo de producto');
  chartsGrid.appendChild(c2);
  INV_CHARTS['categoria'] = new Chart(document.getElementById('inv-c-categoria'), {
    type: 'bar',
    data: {
      labels: catEntries.map(function(e){ return e[0]; }),
      datasets: [{
        label: 'Unidades',
        data: catEntries.map(function(e){ return e[1]; }),
        backgroundColor: catColors.slice(0, catEntries.length).map(function(c){ return c+'BB'; }),
        borderColor: catColors.slice(0, catEntries.length),
        borderWidth: 2, borderRadius: 8, borderSkipped: false,
      }]
    },
    options: {
      indexAxis: 'y',
      plugins: { legend: { display: false }, tooltip: TOOLTIP_OPTS },
      scales: {
        x: { grid: XGRID, ticks: XTICK_MONO },
        y: { grid: { display: false }, ticks: YTICK }
      }
    }
  });

  // ── 3. Bar vertical: CB vs VP ─────────────────────────────────
  const c3 = makeCard('inv-c-negocio', 'CB vs VP', 'Unidades por tipo de negocio');
  chartsGrid.appendChild(c3);
  INV_CHARTS['negocio'] = new Chart(document.getElementById('inv-c-negocio'), {
    type: 'bar',
    data: {
      labels: ['CB (Wompi)', 'VP (Valor Plus)'],
      datasets: [{
        data: [negMap.CB, negMap.VP],
        backgroundColor: ['#99D1FCBB', '#C084FCBB'],
        borderColor: ['#99D1FC', '#C084FC'],
        borderWidth: 2, borderRadius: 10, borderSkipped: false,
      }]
    },
    options: {
      plugins: {
        legend: { display: false },
        tooltip: Object.assign({}, TOOLTIP_OPTS, {
          callbacks: {
            label: function(ctx) {
              var v = ctx.parsed.y;
              var pct = total ? (v / total * 100).toFixed(1) : 0;
              return ' ' + v.toLocaleString('es-CO') + ' uds (' + pct + '%)';
            }
          }
        })
      },
      scales: {
        x: { grid: { display: false }, ticks: { color: '#94a3b8', font: { family: 'Outfit', size: 13, weight: '600' } } },
        y: { grid: XGRID, ticks: XTICK_MONO }
      }
    }
  });

  // ── 4. Bar horizontal: Top 10 bodegas (full width) ────────────
  const shortName = function(n) {
    return n.replace('ALMACEN WOMPI VP ', '').replace('ALMACEN WOMPI ', '').replace('ALMACEN ', '');
  };
  const topLabels = topBodegas.map(function(b){ return shortName(b[0]); });
  const topVals   = topBodegas.map(function(b){ return b[1]; });
  const maxVal    = Math.max.apply(null, topVals);
  const gradColors = topVals.map(function(v) {
    var ratio = maxVal ? v / maxVal : 0;
    if (ratio > 0.7) return '#B0F2AECC';
    if (ratio > 0.4) return '#99D1FCCC';
    return '#7BC8FBCC';
  });

  const c4 = makeCard('inv-c-bodegas', 'Top 10 Bodegas por Stock', 'Ubicaciones con mayor inventario actual');
  c4.style.gridColumn = '1 / -1';
  chartsGrid.appendChild(c4);
  INV_CHARTS['bodegas'] = new Chart(document.getElementById('inv-c-bodegas'), {
    type: 'bar',
    data: {
      labels: topLabels,
      datasets: [{
        label: 'Unidades',
        data: topVals,
        backgroundColor: gradColors,
        borderColor: gradColors.map(function(c){ return c.replace('CC',''); }),
        borderWidth: 2, borderRadius: 6, borderSkipped: false,
      }]
    },
    options: {
      indexAxis: 'y',
      plugins: { legend: { display: false }, tooltip: TOOLTIP_OPTS },
      scales: {
        x: { grid: XGRID, ticks: XTICK_MONO },
        y: { grid: { display: false }, ticks: { color: '#e2e8f0', font: { family: 'Outfit', size: 11 } } }
      }
    }
  });
}

// ══════════════════════════════════════════════════════════════════
//  RENDER TABLE — Top 20 referencias
// ══════════════════════════════════════════════════════════════════
function _invRenderTable() {
  const rows  = INV_FILTERED;
  const total = sumCantidad(rows);
  const container = document.getElementById('inv-table-container');
  if (!container) return;

  const refMap = {};
  rows.forEach(function(r) {
    const nombre = (r['Nombre'] || 'Sin nombre').trim();
    const cat    = invCategoria(nombre);
    const qty    = parseInt(r['Cantidad']) || 0;
    if (!refMap[nombre]) refMap[nombre] = { nombre: nombre, cat: cat, total: 0 };
    refMap[nombre].total += qty;
  });

  const sorted = Object.values(refMap).sort(function(a,b){ return b.total - a.total; }).slice(0, 20);

  const catColor = {
    'Datáfonos': '#99D1FC', 'Rollos': '#B0F2AE', 'Pin pad': '#FFC04D',
    'Forros': '#C084FC', 'Accesorios': '#FB923C', 'SIM': '#F87171',
    'KIT POP VP': '#DFFF61',
  };

  container.innerHTML =
    '<table style="width:100%;border-collapse:separate;border-spacing:0;font-family:\'Outfit\',sans-serif;font-size:13px;">' +
      '<thead>' +
        '<tr style="border-bottom:2px solid rgba(255,255,255,0.08);">' +
          '<th style="padding:12px 16px;text-align:left;font-size:10px;letter-spacing:1px;font-weight:700;color:#64748b;text-transform:uppercase;">#</th>' +
          '<th style="padding:12px 16px;text-align:left;font-size:10px;letter-spacing:1px;font-weight:700;color:#64748b;text-transform:uppercase;">Referencia</th>' +
          '<th style="padding:12px 16px;text-align:left;font-size:10px;letter-spacing:1px;font-weight:700;color:#64748b;text-transform:uppercase;">Categoría</th>' +
          '<th style="padding:12px 16px;text-align:right;font-size:10px;letter-spacing:1px;font-weight:700;color:#64748b;text-transform:uppercase;">Unidades</th>' +
          '<th style="padding:12px 16px;text-align:right;font-size:10px;letter-spacing:1px;font-weight:700;color:#64748b;text-transform:uppercase;">% Total</th>' +
          '<th style="padding:12px 24px 12px 16px;text-align:left;font-size:10px;letter-spacing:1px;font-weight:700;color:#64748b;text-transform:uppercase;">Distribución</th>' +
        '</tr>' +
      '</thead>' +
      '<tbody>' +
        sorted.map(function(r, i) {
          var pct   = total ? (r.total / total * 100) : 0;
          var color = catColor[r.cat] || '#94a3b8';
          var bg    = i % 2 === 0 ? 'rgba(255,255,255,0.015)' : 'transparent';
          var bgOut = bg;
          return (
            '<tr style="background:' + bg + ';transition:background 0.15s;" ' +
                'onmouseover="this.style.background=\'rgba(176,242,174,0.06)\'" ' +
                'onmouseout="this.style.background=\'' + bgOut + '\'">' +
              '<td style="padding:11px 16px;color:#475569;font-family:\'JetBrains Mono\',monospace;font-size:11px;font-weight:600;">' + String(i+1).padStart(2,'0') + '</td>' +
              '<td style="padding:11px 16px;color:#e2e8f0;font-weight:500;max-width:320px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;" title="' + r.nombre + '">' + r.nombre + '</td>' +
              '<td style="padding:11px 16px;">' +
                '<span style="display:inline-block;background:' + color + '22;color:' + color + ';font-size:10px;font-weight:700;padding:2px 9px;border-radius:20px;letter-spacing:0.3px;">' + r.cat + '</span>' +
              '</td>' +
              '<td style="padding:11px 16px;text-align:right;font-family:\'JetBrains Mono\',monospace;font-size:13px;font-weight:700;color:' + color + ';">' + r.total.toLocaleString('es-CO') + '</td>' +
              '<td style="padding:11px 16px;text-align:right;font-family:\'JetBrains Mono\',monospace;font-size:12px;color:#94a3b8;">' + pct.toFixed(1) + '%</td>' +
              '<td style="padding:11px 24px 11px 16px;min-width:140px;">' +
                '<div style="background:rgba(255,255,255,0.05);border-radius:4px;height:6px;overflow:hidden;">' +
                  '<div style="width:' + Math.min(pct * 5, 100) + '%;height:100%;background:' + color + ';border-radius:4px;transition:width 0.6s ease;"></div>' +
                '</div>' +
              '</td>' +
            '</tr>'
          );
        }).join('') +
      '</tbody>' +
    '</table>';
}

// ══════════════════════════════════════════════════════════════════
//  ENTRY POINT
// ══════════════════════════════════════════════════════════════════
async function renderInventarioPrincipal() {
  const panel = document.getElementById('panel-inv-principal');
  if (!panel) return;

  if (!INV_RAW) {
    const grid = document.getElementById('inv-kpi-grid');
    if (grid) grid.innerHTML = '<div class="loading"><div class="spinner"></div><span>Cargando inventario...</span></div>';
    await loadInventarioData();
    _invPopulateFilters();
  }

  INV_FILTERED = INV_RAW ? INV_RAW.slice() : [];
  _invRenderAll();
}

window.renderInventarioPrincipal = renderInventarioPrincipal;

// ══════════════════════════════════════════════════════════════════
//  CARGA ANTICIPADA — en paralelo con dashboard.js
//  Notifica a dashboard.js cuando termina para desbloquear la pantalla
// ══════════════════════════════════════════════════════════════════
(async function _invEarlyLoad() {
  try {
    await loadInventarioData();
    _invPopulateFilters();
  } catch (e) {
    console.warn('[Inventario] Early load error:', e);
  } finally {
    if (typeof window._setInventarioLoaded === 'function') {
      window._setInventarioLoaded();
    }
  }
})();