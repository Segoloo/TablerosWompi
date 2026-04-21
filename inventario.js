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
    window.INV_RAW = INV_RAW;
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
    { label:'TOTAL INVENTARIO', value:fmtN(total),      sub1:'100% del stock',               sub2:'42 bodegas Wompi',       color:INV_PALETTE.total,    icon:'📦', wide:true,  drillRows: rows,                                                                                                                                     drillTitle:'Total Inventario' },
    { label:'EN BODEGA',        value:fmtN(unBodega),   sub1:fmtPct(unBodega,total),         sub2:fmtN(unBodega)+' uds',   color:INV_PALETTE.bodega,   icon:'🏪', drillRows: rows.filter(function(r){ return INV_BODEGAS.has((r['Nombre de la ubicación']||'').trim()); }),                                                     drillTitle:'Stock en Bodega' },
    { label:'EN COMERCIO',      value:fmtN(unComercio), sub1:fmtPct(unComercio,total),       sub2:fmtN(unComercio)+' uds', color:INV_PALETTE.comercio, icon:'🏬', drillRows: rows.filter(function(r){ return (r['Tipo de ubicación']||'').trim() === 'Site'; }),                                                                  drillTitle:'Stock en Comercio (Site)' },
    { label:'TÉC. LINEACOM',    value:fmtN(unTecnico),  sub1:fmtPct(unTecnico,total),        sub2:fmtN(unTecnico)+' uds',  color:INV_PALETTE.tecnico,  icon:'🔧', drillRows: rows.filter(function(r){ return (r['Tipo de ubicación']||'').trim() === 'Staff'; }),                                                                 drillTitle:'Stock Técnicos Lineacom (Staff)' },
    { label:'GEST. & EMPL. WOMPI',    value:fmtN(unGW),       sub1:fmtPct(unGW,total),             sub2:fmtN(unGW)+' uds',       color:INV_PALETTE.gestores, icon:'👤', drillRows: rows.filter(function(r){ return GW_RE.test((r['Código de ubicación']||'').trim()); }),                                                              drillTitle:'Stock Gestores & Empleados (GW)' },
    { label:'INGENICO',         value:fmtN(unIngenico), sub1:fmtPct(unIngenico,total),       sub2:fmtN(unIngenico)+' uds', color:INV_PALETTE.ingenico, icon:'🔌', drillRows: rows.filter(function(r){ return (r['Nombre de la ubicación']||'').trim() === 'ALMACEN INGENICO - PROVEEDOR WOMPI'; }),                             drillTitle:'Stock Ingenico (Proveedor)' },
    { label:'OPL',              value:fmtN(unOPL),      sub1:fmtPct(unOPL,total),            sub2:fmtN(unOPL)+' uds',      color:INV_PALETTE.opl,      icon:'🚚', drillRows: rows.filter(function(r){ return (r['Tipo de ubicación']||'').trim() === 'Supplier'; }),                                                             drillTitle:'Stock OPL (Supplier)' },
    { label:'42 BODEGAS',       value:'42',             sub1:'Bodegas Wompi',                sub2:'Ver distribución',       color:'#67e8f9',             icon:'🗺️', drillRows: null, drillTitle:'Bodegas', special:'bodegas' },
  ];

  const grid = document.getElementById('inv-kpi-grid');
  if (!grid) return;

  grid.innerHTML = kpis.map(function(k, idx) {
    // Guardamos las rows de cada KPI en un objeto global indexado para poder llamarlas desde onclick inline
    window._invKpiDrillData = window._invKpiDrillData || [];
    window._invKpiDrillData[idx] = { rows: k.drillRows, title: k.drillTitle, special: k.special || null };
    return (
      '<div class="kpi-card inv-kpi-v2 fade-up" style="' +
        'background:linear-gradient(145deg,rgba(10,26,18,.95) 0%,rgba(8,20,14,.9) 100%);' +
        'border:1px solid rgba(176,242,174,.1);border-top:2px solid ' + k.color + ';' +
        'border-radius:var(--radius,18px);padding:22px 20px;' +
        'position:relative;overflow:hidden;' +
        'box-shadow:0 4px 20px rgba(0,0,0,.4),inset 0 1px 0 rgba(255,255,255,.03);' +
        'transition:all .25s cubic-bezier(.4,0,.2,1);' +
        'cursor:pointer;' +
        (k.wide ? 'grid-column:span 2;' : '') +
      '" ' +
      'onclick="(window._invKpiDrillData[' + idx + '].special===\'bodegas\') ? invOpenBodegasModal() : invOpenDrillModal(window._invKpiDrillData[' + idx + '].title, window._invKpiDrillData[' + idx + '].rows)" ' +
      'title="🔍 Ver detalle de ' + k.label + '" ' +
      'onmouseover="this.style.transform=\'translateY(-5px)\';this.style.borderColor=\'' + k.color + '55\';this.style.boxShadow=\'0 12px 40px rgba(0,0,0,.5),0 0 30px ' + k.color + '22\'" ' +
      'onmouseout="this.style.transform=\'\';this.style.borderColor=\'rgba(176,242,174,.1)\';this.style.boxShadow=\'0 4px 20px rgba(0,0,0,.4)\'">' +
        // Glow orb
        '<div style="position:absolute;top:-24px;right:-24px;width:88px;height:88px;border-radius:50%;background:' + k.color + ';opacity:0.06;pointer-events:none;filter:blur(12px);"></div>' +
        // Click hint badge
        '<div style="position:absolute;top:10px;right:12px;font-size:9px;font-family:\'Outfit\',sans-serif;color:' + k.color + ';opacity:0.55;letter-spacing:.5px;text-transform:uppercase;font-weight:600;">Ver detalle ›</div>' +
        // Icon
        '<span style="font-size:18px;margin-bottom:14px;display:block;">' + k.icon + '</span>' +
        // Label — VP system style
        '<div style="font-size:10px;font-weight:600;color:var(--muted,#7A7674);text-transform:uppercase;letter-spacing:1px;margin-bottom:10px;line-height:1.3;font-family:\'Outfit\',sans-serif;">' + k.label + '</div>' +
        // Value — JetBrains Mono like VP
        '<div style="font-family:\'JetBrains Mono\',monospace;font-size:34px;font-weight:700;color:' + k.color + ';line-height:1;letter-spacing:-1.5px;animation:countUp .4s cubic-bezier(.34,1.56,.64,1);text-shadow:0 0 28px ' + k.color + '66;margin-bottom:10px;">' + k.value + '</div>' +
        // Sub tags
        '<div style="display:flex;align-items:center;gap:7px;flex-wrap:wrap;">' +
          '<span style="display:inline-block;background:' + k.color + '22;color:' + k.color + ';font-size:11px;font-weight:700;padding:2px 9px;border-radius:20px;font-family:\'JetBrains Mono\',monospace;letter-spacing:.3px;">' + k.sub1 + '</span>' +
          '<span style="font-size:11px;color:var(--muted,#7A7674);font-family:\'Outfit\',sans-serif;">' + k.sub2 + '</span>' +
        '</div>' +
        // Progress bar for percentage values
        (k.sub1.includes('%') && k.sub1 !== '100.0%' ? (
          '<div style="margin-top:14px;">' +
            '<div style="height:3px;background:rgba(255,255,255,.06);border-radius:2px;overflow:hidden;">' +
              '<div style="width:' + Math.min(parseFloat(k.sub1), 100) + '%;height:100%;background:linear-gradient(90deg,' + k.color + '88,' + k.color + ');border-radius:2px;transition:width 1s cubic-bezier(.4,0,.2,1);"></div>' +
            '</div>' +
          '</div>'
        ) : '') +
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

  // ── Datos — calculados TODOS antes de tocar el DOM ─────────────
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

  // catEntries declarado en el scope de la función (usado por chart 2 y chart 6 polar)
  const catEntries = Object.entries(catMap).sort(function(a,b){ return b[1]-a[1]; });

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

  // ── Chart defaults — VP system ─────────────────────────────────
  const TOOLTIP_OPTS = {
    backgroundColor: 'rgba(24,23,21,.95)',
    titleColor: '#B0F2AE', bodyColor: '#FAFAFA',
    borderColor: 'rgba(176,242,174,.2)', borderWidth: 1, padding: 12,
    titleFont: { family: 'Syne', size: 13, weight: '700' },
    bodyFont:  { family: 'Outfit', size: 12 },
  };
  const LEGEND_OPTS = {
    labels: {
      color: '#FAFAFA',
      font: { family: 'Outfit', size: 12 },
      padding: 16,
      boxWidth: 12,
    }
  };
  const XGRID   = { color: 'rgba(255,255,255,.05)' };
  const YTICK   = { color: '#7A7674', font: { family: 'Outfit', size: 12 }, border: { display: false } };
  const XTICK_MONO = {
    color: '#7A7674',
    font: { family: 'JetBrains Mono', size: 11 },
    callback: function(v){ return v.toLocaleString('es-CO'); },
    border: { display: false }
  };

  // ── Card builder — VP style ────────────────────────────────────
  function makeCard(canvasId, title, sub) {
    const div = document.createElement('div');
    div.style.cssText = [
      'background:linear-gradient(145deg,rgba(10,26,18,.95) 0%,rgba(8,20,14,.9) 100%)',
      'border:1px solid rgba(176,242,174,.1)',
      'border-radius:var(--radius,18px)',
      'padding:22px 24px 20px',
      'position:relative',
      'overflow:hidden',
      'box-shadow:0 4px 20px rgba(0,0,0,.4),inset 0 1px 0 rgba(255,255,255,.03)',
      'transition:all .25s cubic-bezier(.4,0,.2,1)',
    ].join(';');
    div.innerHTML =
      '<div style="position:absolute;top:0;left:0;right:0;height:2px;background:linear-gradient(90deg,var(--verde-menta,#B0F2AE),var(--azul-cielo,#99D1FC),transparent);opacity:.6;"></div>' +
      '<div style="margin-bottom:16px;">' +
        '<div style="font-family:\'Syne\',sans-serif;font-size:13px;font-weight:700;color:#f1f5f9;letter-spacing:.3px;">' + title + '</div>' +
        (sub ? '<div style="font-size:11px;color:var(--muted,#7A7674);margin-top:3px;font-family:\'Outfit\',sans-serif;">' + sub + '</div>' : '') +
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
      labels: ['CB (Wompi)', 'VP (Venta Presente)'],
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
        x: { grid: { display: false }, ticks: { color: '#FAFAFA', font: { family: 'Outfit', size: 13, weight: '600' } }, border: { display: false } },
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
        y: { grid: { display: false }, ticks: { color: '#FAFAFA', font: { family: 'Outfit', size: 11 }, border: { display: false } } }
      }
    }
  });

}

// ══════════════════════════════════════════════════════════════════
//  MODAL DRILLDOWN — abre tabla con filas del KPI seleccionado
// ══════════════════════════════════════════════════════════════════
function invOpenDrillModal(title, rows) {
  // ── Limpiar modal anterior ─────────────────────────────────────
  var prev = document.getElementById('inv-drill-modal');
  if (prev) prev.remove();

  // ── Estado interno ─────────────────────────────────────────────
  var _allRows      = rows || [];
  var _filtered     = _allRows.slice();
  var _page         = 1;
  var PAGE_SIZE     = 50;
  var _title        = title || 'Detalle';

  var CAT_COLOR = {
    'Datáfonos':'#99D1FC','Rollos':'#B0F2AE','Pin pad':'#FFC04D',
    'Forros':'#C084FC','Accesorios':'#FB923C','SIM':'#F87171','KIT POP VP':'#DFFF61',
  };

  // ── Helpers ────────────────────────────────────────────────────
  function badge(text, color) {
    var s = document.createElement('span');
    s.style.cssText = 'display:inline-block;background:'+color+'22;color:'+color+';font-size:10px;font-weight:700;padding:2px 9px;border-radius:20px;white-space:nowrap;';
    s.textContent = text;
    return s;
  }
  function el(tag, css, text) {
    var e = document.createElement(tag);
    if (css)  e.style.cssText = css;
    if (text !== undefined) e.textContent = text;
    return e;
  }

  // ── Render tabla ───────────────────────────────────────────────
  function renderTable() {
    var pages = Math.max(1, Math.ceil(_filtered.length / PAGE_SIZE));
    if (_page > pages) _page = pages;
    var slice = _filtered.slice((_page-1)*PAGE_SIZE, _page*PAGE_SIZE);

    // Actualizar contadores
    countEl.textContent = _filtered.length + ' registros' +
      (_filtered.length !== _allRows.length ? ' (de ' + _allRows.length + ' totales)' : '');
    footerCountEl.textContent = 'Página ' + _page + ' / ' + pages + '  ·  ' + _filtered.length + ' registros';

    // Limpiar cuerpo
    tbody.innerHTML = '';

    if (!slice.length) {
      var emptyRow = document.createElement('tr');
      var emptyTd  = el('td', 'text-align:center;padding:48px;color:#7A7674;font-family:\'Outfit\',sans-serif;', 'Sin registros para los filtros aplicados');
      emptyTd.colSpan = 9;
      emptyRow.appendChild(emptyTd);
      tbody.appendChild(emptyRow);
    } else {
      slice.forEach(function(r, i) {
        var cat      = invCategoria(r['Nombre']);
        var neg      = invNegocio(r['Subtipo']);
        var catColor = CAT_COLOR[cat] || '#94a3b8';
        var negColor = neg === 'VP' ? '#C084FC' : '#99D1FC';
        var bgBase   = i % 2 === 0 ? 'rgba(176,242,174,.015)' : 'transparent';

        var tr = document.createElement('tr');
        tr.style.cssText = 'background:'+bgBase+';transition:background .12s;';
        tr.addEventListener('mouseover', function(){ tr.style.background='rgba(176,242,174,.05)'; });
        tr.addEventListener('mouseout',  function(){ tr.style.background=bgBase; });

        function td(css, content) {
          var cell = el('td', 'padding:9px 12px;' + (css||''));
          if (typeof content === 'string' || typeof content === 'number') {
            cell.textContent = content;
          } else if (content) {
            cell.appendChild(content);
          }
          return cell;
        }

        var refTd = td('color:#FAFAFA;font-weight:500;max-width:260px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;');
        refTd.title = r['Nombre'] || '';
        refTd.textContent = r['Nombre'] || '—';

        var serialTd = td('color:#a5f3fc;font-family:\'JetBrains Mono\',monospace;font-size:11px;max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;');
        serialTd.title = r['Número de serie'] || r['Numero de serie'] || r['Serial'] || '';
        serialTd.textContent = r['Número de serie'] || r['Numero de serie'] || r['Serial'] || '—';

        tr.appendChild(refTd);
        tr.appendChild(serialTd);
        tr.appendChild(td('', badge(cat, catColor)));
        tr.appendChild(td('', badge(neg, negColor)));
        var bodTd = td('color:#cbd5e1;font-size:11px;max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;');
        bodTd.title = r['Nombre de la ubicación'] || '';
        bodTd.textContent = r['Nombre de la ubicación'] || '—';
        tr.appendChild(bodTd);
        tr.appendChild(td('color:#7A7674;font-size:11px;', r['Tipo de ubicación'] || '—'));
        tr.appendChild(td('color:#7A7674;font-family:\'JetBrains Mono\',monospace;font-size:11px;', r['Código de ubicación'] || '—'));
        var qtyTd = td('text-align:right;font-family:\'JetBrains Mono\',monospace;font-size:14px;font-weight:700;color:'+catColor+';');
        qtyTd.textContent = (parseInt(r['Cantidad'])||0).toLocaleString('es-CO');
        tr.appendChild(qtyTd);
        tr.appendChild(td('color:#7A7674;font-size:11px;', r['Subtipo'] || '—'));

        tbody.appendChild(tr);
      });
    }

    // Paginación
    pagEl.innerHTML = '';
    var btnCSS = 'padding:5px 11px;border-radius:7px;cursor:pointer;font-size:12px;font-family:\'Outfit\',sans-serif;border:1px solid rgba(176,242,174,.18);transition:all .15s;';
    function mkPageBtn(label, targetPage, active) {
      var b = el('button', btnCSS + (active
        ? 'background:#B0F2AE;color:#0a1a12;font-weight:700;'
        : 'background:rgba(255,255,255,.06);color:#94a3b8;'));
      b.textContent = label;
      b.addEventListener('click', function(){ _page = targetPage; renderTable(); });
      return b;
    }
    if (_page > 1)  pagEl.appendChild(mkPageBtn('‹ Ant', _page-1, false));
    var s = Math.max(1, _page-2), e = Math.min(pages, _page+2);
    for (var pg = s; pg <= e; pg++) pagEl.appendChild(mkPageBtn(String(pg), pg, pg===_page));
    if (_page < pages) pagEl.appendChild(mkPageBtn('Sig ›', _page+1, false));
  }

  // ── Filtrar ────────────────────────────────────────────────────
  function applyFilter() {
    var search = (searchInput.value || '').toLowerCase().trim();
    var bod    = selBodega.value;
    var cat    = selCat.value;
    var neg    = selNeg.value;
    _filtered = _allRows.filter(function(r) {
      if (bod && (r['Nombre de la ubicación']||'').trim() !== bod) return false;
      if (cat && invCategoria(r['Nombre']) !== cat) return false;
      if (neg && invNegocio(r['Subtipo'])  !== neg) return false;
      if (search) {
        var hit = Object.values(r).some(function(v){ return String(v||'').toLowerCase().includes(search); });
        if (!hit) return false;
      }
      return true;
    });
    _page = 1;
    renderTable();
  }

  // ── Exportar Excel ─────────────────────────────────────────────
  function exportExcel() {
    if (!_filtered.length) { alert('Sin datos para exportar.'); return; }
    if (typeof XLSX === 'undefined') { alert('XLSX no disponible.'); return; }
    var exportRows = _filtered.map(function(r) {
      return {
        'Referencia':       r['Nombre'] || '',
        'N.° de Serie':     r['Número de serie'] || r['Numero de serie'] || r['Serial'] || '',
        'Categoría':        invCategoria(r['Nombre']),
        'Negocio':          invNegocio(r['Subtipo']),
        'Bodega':           r['Nombre de la ubicación'] || '',
        'Tipo Ubicación':   r['Tipo de ubicación'] || '',
        'Cód. Ubicación':   r['Código de ubicación'] || '',
        'Cantidad':         parseInt(r['Cantidad']) || 0,
        'Subtipo':          r['Subtipo'] || '',
      };
    });
    var wb = XLSX.utils.book_new();
    var ws = XLSX.utils.json_to_sheet(exportRows);
    ws['!cols'] = Object.keys(exportRows[0]).map(function(k){
      return { wch: Math.max(k.length+2, ...exportRows.slice(0,50).map(function(rx){ return String(rx[k]||'').length; })) };
    });
    XLSX.utils.book_append_sheet(wb, ws, _title.slice(0,31));
    XLSX.writeFile(wb, 'Inventario_' + _title.replace(/[^a-zA-Z0-9]/g,'_') + '_' + new Date().toISOString().slice(0,10) + '.xlsx');
  }

  // ══════════════════════════════════════════════════════════════
  //  CONSTRUCCIÓN DEL DOM
  // ══════════════════════════════════════════════════════════════

  // Overlay
  var overlay = el('div', 'position:fixed;inset:0;z-index:99999;display:flex;align-items:center;justify-content:center;background:rgba(0,0,0,.80);backdrop-filter:blur(8px);padding:20px;animation:fadeIn .18s ease;');
  overlay.id = 'inv-drill-modal';
  overlay.addEventListener('click', function(e){ if (e.target === overlay) overlay.remove(); });

  // Panel
  var panel = el('div',
    'background:#181715;border:1px solid rgba(176,242,174,.18);border-radius:18px;' +
    'width:100%;max-width:1140px;max-height:90vh;display:flex;flex-direction:column;' +
    'box-shadow:0 32px 96px rgba(0,0,0,.9);animation:slideUp .2s cubic-bezier(.34,1.2,.64,1);overflow:hidden;'
  );

  // ── Barra de color superior ────────────────────────────────────
  var topBar = el('div','height:3px;background:linear-gradient(90deg,#B0F2AE,#99D1FC,#C084FC);flex-shrink:0;');
  panel.appendChild(topBar);

  // ── Header ────────────────────────────────────────────────────
  var header = el('div','display:flex;align-items:center;justify-content:space-between;padding:20px 24px;border-bottom:1px solid rgba(176,242,174,.1);flex-shrink:0;gap:16px;');

  var headerLeft = el('div','');
  var titleEl    = el('div','font-family:\'Syne\',sans-serif;font-size:17px;font-weight:800;color:#B0F2AE;letter-spacing:-.2px;', _title);
  var countEl    = el('div','font-size:12px;color:#7A7674;margin-top:4px;font-family:\'Outfit\',sans-serif;');
  headerLeft.appendChild(titleEl);
  headerLeft.appendChild(countEl);

  var headerRight = el('div','display:flex;gap:8px;align-items:center;flex-shrink:0;');
  var btnExport   = el('button',
    'background:rgba(176,242,174,.08);border:1px solid rgba(176,242,174,.22);color:#B0F2AE;' +
    'padding:7px 16px;border-radius:8px;cursor:pointer;font-size:12px;font-family:\'Outfit\',sans-serif;' +
    'display:flex;align-items:center;gap:6px;transition:all .2s;',
    '⬇ Excel'
  );
  btnExport.addEventListener('click', exportExcel);
  btnExport.addEventListener('mouseover', function(){ btnExport.style.background='rgba(176,242,174,.16)'; });
  btnExport.addEventListener('mouseout',  function(){ btnExport.style.background='rgba(176,242,174,.08)'; });

  var btnClose = el('button',
    'background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.1);color:#7A7674;' +
    'width:34px;height:34px;border-radius:8px;cursor:pointer;font-size:20px;' +
    'display:flex;align-items:center;justify-content:center;line-height:1;transition:all .2s;'
  );
  btnClose.innerHTML = '&times;';
  btnClose.addEventListener('click', function(){ overlay.remove(); });
  btnClose.addEventListener('mouseover', function(){ btnClose.style.background='rgba(255,80,80,.15)';btnClose.style.color='#f87171'; });
  btnClose.addEventListener('mouseout',  function(){ btnClose.style.background='rgba(255,255,255,.06)';btnClose.style.color='#7A7674'; });

  headerRight.appendChild(btnExport);
  headerRight.appendChild(btnClose);
  header.appendChild(headerLeft);
  header.appendChild(headerRight);
  panel.appendChild(header);

  // ── Barra de filtros ───────────────────────────────────────────
  var filterBar = el('div','padding:12px 24px;border-bottom:1px solid rgba(176,242,174,.07);flex-shrink:0;display:flex;gap:10px;flex-wrap:wrap;background:rgba(0,0,0,.18);align-items:center;');

  var searchInput = el('input','flex:1;min-width:180px;background:rgba(255,255,255,.05);border:1px solid rgba(176,242,174,.18);border-radius:8px;color:#FAFAFA;padding:8px 12px;font-size:13px;font-family:\'Outfit\',sans-serif;outline:none;');
  searchInput.placeholder = '🔍 Buscar en todos los campos...';
  searchInput.addEventListener('input', applyFilter);

  var selCSS = 'background:#1a1916;border:1px solid rgba(176,242,174,.18);border-radius:8px;color:#FAFAFA;padding:8px 10px;font-size:12px;font-family:\'Outfit\',sans-serif;cursor:pointer;';

  var selBodega = el('select', selCSS + 'min-width:160px;');
  var selCat    = el('select', selCSS + 'min-width:140px;');
  var selNeg    = el('select', selCSS + 'min-width:120px;');

  // Poblar bodegas únicas del subconjunto
  var uniqBodegas = [...new Set(_allRows.map(function(r){ return (r['Nombre de la ubicación']||'').trim(); }).filter(Boolean))].sort();
  selBodega.appendChild(new Option('Todas las bodegas',''));
  uniqBodegas.forEach(function(b){ selBodega.appendChild(new Option(b,b)); });

  // Poblar categorías únicas
  var uniqCats = [...new Set(_allRows.map(function(r){ return invCategoria(r['Nombre']); }))].sort();
  selCat.appendChild(new Option('Todas las categorías',''));
  uniqCats.forEach(function(c){ selCat.appendChild(new Option(c,c)); });

  selNeg.appendChild(new Option('CB + VP',''));
  selNeg.appendChild(new Option('CB (Wompi)','CB'));
  selNeg.appendChild(new Option('VP','VP'));

  [selBodega, selCat, selNeg].forEach(function(s){ s.addEventListener('change', applyFilter); });

  filterBar.appendChild(searchInput);
  filterBar.appendChild(selBodega);
  filterBar.appendChild(selCat);
  filterBar.appendChild(selNeg);
  panel.appendChild(filterBar);

  // ── Tabla scrollable ───────────────────────────────────────────
  var tableWrap = el('div','overflow:auto;flex:1;');

  var table = el('table','width:100%;border-collapse:collapse;font-size:12px;font-family:\'Outfit\',sans-serif;');

  var thead = document.createElement('thead');
  var theadRow = document.createElement('tr');
  theadRow.style.cssText = 'position:sticky;top:0;background:#181715;z-index:2;';
  ['Referencia','N.° de Serie','Categoría','Negocio','Bodega','Tipo Ubic.','Cód. Ubic.','Cantidad','Subtipo'].forEach(function(h){
    var th = el('th','padding:10px 12px;text-align:left;color:#B0F2AE;font-weight:700;font-size:10px;letter-spacing:1px;text-transform:uppercase;border-bottom:1px solid rgba(176,242,174,.15);white-space:nowrap;font-family:\'Syne\',sans-serif;');
    th.textContent = h;
    theadRow.appendChild(th);
  });
  // Ajustar alineación de Cantidad
  theadRow.children[7].style.textAlign = 'right';
  thead.appendChild(theadRow);
  table.appendChild(thead);

  var tbody = document.createElement('tbody');
  table.appendChild(tbody);
  tableWrap.appendChild(table);
  panel.appendChild(tableWrap);

  // ── Footer paginación ──────────────────────────────────────────
  var footer = el('div','display:flex;align-items:center;justify-content:space-between;padding:11px 24px;border-top:1px solid rgba(176,242,174,.08);flex-shrink:0;background:rgba(0,0,0,.15);gap:12px;');
  var footerCountEl = el('span','font-size:12px;color:#7A7674;font-family:\'Outfit\',sans-serif;');
  var pagEl         = el('div','display:flex;gap:6px;flex-wrap:wrap;');
  footer.appendChild(footerCountEl);
  footer.appendChild(pagEl);
  panel.appendChild(footer);

  overlay.appendChild(panel);

  // Añadir animaciones si no existen
  if (!document.getElementById('inv-drill-keyframes')) {
    var style = document.createElement('style');
    style.id = 'inv-drill-keyframes';
    style.textContent = '@keyframes fadeIn{from{opacity:0}to{opacity:1}}@keyframes slideUp{from{opacity:0;transform:translateY(24px) scale(.97)}to{opacity:1;transform:none}}';
    document.head.appendChild(style);
  }

  document.body.appendChild(overlay);

  // Render inicial
  renderTable();
}

// Alias global para backwards-compat (ya no se usa desde fuera pero por si acaso)
window.invDrillFilter   = function(){};
window.invDrillGoPage   = function(){};
window.invExportDrillExcel = function(){};

// ══════════════════════════════════════════════════════════════════
//  EXPORT TOP REFERENCIAS
// ══════════════════════════════════════════════════════════════════
window.invExportTopExcel = function() {
  var rows = INV_FILTERED;
  var total = sumCantidad(rows);
  var refMap = {};
  rows.forEach(function(r) {
    var nombre = (r['Nombre'] || 'Sin nombre').trim();
    var cat    = invCategoria(nombre);
    var qty    = parseInt(r['Cantidad']) || 0;
    if (!refMap[nombre]) refMap[nombre] = { Referencia: nombre, Categoría: cat, Unidades: 0 };
    refMap[nombre].Unidades += qty;
  });
  var sorted = Object.values(refMap).sort(function(a,b){ return b.Unidades - a.Unidades; }).slice(0, 20);
  sorted = sorted.map(function(r){ return Object.assign({}, r, { '% Total': total ? (r.Unidades/total*100).toFixed(1)+'%' : '0%' }); });
  if (!sorted.length) { alert('Sin datos.'); return; }
  if (typeof XLSX === 'undefined') { alert('XLSX no disponible.'); return; }
  var wb = XLSX.utils.book_new();
  var ws = XLSX.utils.json_to_sheet(sorted);
  XLSX.utils.book_append_sheet(wb, ws, 'Top Referencias');
  XLSX.writeFile(wb, 'Top_Referencias_Inventario_' + new Date().toISOString().slice(0,10) + '.xlsx');
};


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
        '<tr style="border-bottom:1px solid rgba(176,242,174,.1);">' +
          '<th style="padding:12px 16px;text-align:left;font-family:\'Syne\',sans-serif;font-size:9px;letter-spacing:2px;font-weight:700;color:rgba(176,242,174,.6);text-transform:uppercase;">#</th>' +
          '<th style="padding:12px 16px;text-align:left;font-family:\'Syne\',sans-serif;font-size:9px;letter-spacing:2px;font-weight:700;color:rgba(176,242,174,.6);text-transform:uppercase;">Referencia</th>' +
          '<th style="padding:12px 16px;text-align:left;font-family:\'Syne\',sans-serif;font-size:9px;letter-spacing:2px;font-weight:700;color:rgba(176,242,174,.6);text-transform:uppercase;">Categoría</th>' +
          '<th style="padding:12px 16px;text-align:right;font-family:\'Syne\',sans-serif;font-size:9px;letter-spacing:2px;font-weight:700;color:rgba(176,242,174,.6);text-transform:uppercase;">Unidades</th>' +
          '<th style="padding:12px 16px;text-align:right;font-family:\'Syne\',sans-serif;font-size:9px;letter-spacing:2px;font-weight:700;color:rgba(176,242,174,.6);text-transform:uppercase;">% Total</th>' +
          '<th style="padding:12px 24px 12px 16px;text-align:left;font-family:\'Syne\',sans-serif;font-size:9px;letter-spacing:2px;font-weight:700;color:rgba(176,242,174,.6);text-transform:uppercase;">Distribución</th>' +
        '</tr>' +
      '</thead>' +
      '<tbody>' +
        sorted.map(function(r, i) {
          var pct   = total ? (r.total / total * 100) : 0;
          var color = catColor[r.cat] || '#94a3b8';
          var bg    = i % 2 === 0 ? 'rgba(176,242,174,.015)' : 'transparent';
          return (
            '<tr style="background:' + bg + ';transition:background 0.15s;" ' +
                'onmouseover="this.style.background=\'rgba(176,242,174,0.05)\'" ' +
                'onmouseout="this.style.background=\'' + bg + '\'">' +
              '<td style="padding:11px 16px;color:rgba(176,242,174,.4);font-family:\'JetBrains Mono\',monospace;font-size:11px;font-weight:600;">' + String(i+1).padStart(2,'0') + '</td>' +
              '<td style="padding:11px 16px;color:#FAFAFA;font-weight:500;font-family:\'Outfit\',sans-serif;max-width:320px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;" title="' + r.nombre + '">' + r.nombre + '</td>' +
              '<td style="padding:11px 16px;">' +
                '<span style="display:inline-block;background:' + color + '22;color:' + color + ';font-size:10px;font-weight:700;padding:2px 9px;border-radius:20px;letter-spacing:0.3px;font-family:\'Outfit\',sans-serif;">' + r.cat + '</span>' +
              '</td>' +
              '<td style="padding:11px 16px;text-align:right;font-family:\'JetBrains Mono\',monospace;font-size:14px;font-weight:700;color:' + color + ';text-shadow:0 0 16px ' + color + '66;">' + r.total.toLocaleString('es-CO') + '</td>' +
              '<td style="padding:11px 16px;text-align:right;font-family:\'JetBrains Mono\',monospace;font-size:12px;color:#7A7674;">' + pct.toFixed(1) + '%</td>' +
              '<td style="padding:11px 24px 11px 16px;min-width:140px;">' +
                '<div style="background:rgba(255,255,255,.05);border-radius:4px;height:4px;overflow:hidden;">' +
                  '<div style="width:' + Math.min(pct * 5, 100) + '%;height:100%;background:linear-gradient(90deg,' + color + '88,' + color + ');border-radius:4px;transition:width 0.8s cubic-bezier(.4,0,.2,1);"></div>' +
                '</div>' +
              '</td>' +
            '</tr>'
          );
        }).join('') +
      '</tbody>' +
    '</table>';
}

// ══════════════════════════════════════════════════════════════════
//  MODAL BODEGAS — muestra las 42 bodegas hardcodeadas con su stock
// ══════════════════════════════════════════════════════════════════
function invOpenBodegasModal() {
  var prev = document.getElementById('inv-bodegas-modal');
  if (prev) prev.remove();

  // Calcular stock por bodega desde los datos filtrados actuales
  var stockMap = {};
  (INV_FILTERED || []).forEach(function(r) {
    var bod = (r['Nombre de la ubicación'] || '').trim();
    if (bod) stockMap[bod] = (stockMap[bod] || 0) + (parseInt(r['Cantidad']) || 0);
  });

  // Construir lista de las 42 bodegas con stock (0 si no hay datos)
  var bodegaList = [...INV_BODEGAS].map(function(b) {
    return { nombre: b, stock: stockMap[b] || 0 };
  }).sort(function(a, b) { return b.stock - a.stock; });

  var totalBodStock = bodegaList.reduce(function(s, b) { return s + b.stock; }, 0);

  // Helper
  function el(tag, css, text) {
    var e = document.createElement(tag);
    if (css)  e.style.cssText = css;
    if (text !== undefined) e.textContent = text;
    return e;
  }

  // Nombre corto para display
  function shortName(n) {
    return n.replace('ALMACEN WOMPI VP ', '').replace('ALMACEN WOMPI ', '').replace('ALMACEN ', '');
  }

  // Tipo de bodega
  function bodTipo(n) {
    if (n.includes('| ALQUILER')) return { label: 'VP Alquiler', color: '#C084FC' };
    if (n.includes('| VENTA'))   return { label: 'VP Venta',    color: '#F87171' };
    if (n.includes('BAJAS'))     return { label: 'Bajas',       color: '#FB923C' };
    if (n.includes('INGENICO'))  return { label: 'Proveedor',   color: '#FFC04D' };
    if (n.includes('ALISTAMIENTO')) return { label: 'Alist.',   color: '#DFFF61' };
    return { label: 'Almacén',   color: '#B0F2AE' };
  }

  // Overlay
  var overlay = el('div', 'position:fixed;inset:0;z-index:99999;display:flex;align-items:center;justify-content:center;background:rgba(0,0,0,.82);backdrop-filter:blur(8px);padding:20px;animation:fadeIn .18s ease;');
  overlay.id = 'inv-bodegas-modal';
  overlay.addEventListener('click', function(e) { if (e.target === overlay) overlay.remove(); });

  // Panel
  var panel = el('div',
    'background:#181715;border:1px solid rgba(103,232,249,.22);border-radius:18px;' +
    'width:100%;max-width:1100px;max-height:90vh;display:flex;flex-direction:column;' +
    'box-shadow:0 32px 96px rgba(0,0,0,.9);animation:slideUp .2s cubic-bezier(.34,1.2,.64,1);overflow:hidden;'
  );

  // Barra color
  var topBar = el('div', 'height:3px;background:linear-gradient(90deg,#67e8f9,#B0F2AE,#C084FC);flex-shrink:0;');
  panel.appendChild(topBar);

  // Header
  var header = el('div', 'display:flex;align-items:center;justify-content:space-between;padding:20px 24px;border-bottom:1px solid rgba(103,232,249,.12);flex-shrink:0;gap:16px;');
  var hLeft  = el('div', '');
  var hTitle = el('div', 'font-family:\'Syne\',sans-serif;font-size:17px;font-weight:800;color:#67e8f9;letter-spacing:-.2px;', '🗺️  42 Bodegas Wompi');
  var hSub   = el('div', 'font-size:12px;color:#7A7674;margin-top:4px;font-family:\'Outfit\',sans-serif;',
    'Stock total en bodegas: ' + totalBodStock.toLocaleString('es-CO') + ' uds  ·  42 ubicaciones');
  hLeft.appendChild(hTitle);
  hLeft.appendChild(hSub);

  var btnClose = el('button',
    'background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.1);color:#7A7674;' +
    'width:34px;height:34px;border-radius:8px;cursor:pointer;font-size:20px;' +
    'display:flex;align-items:center;justify-content:center;line-height:1;transition:all .2s;'
  );
  btnClose.innerHTML = '&times;';
  btnClose.addEventListener('click', function() { overlay.remove(); });
  btnClose.addEventListener('mouseover', function() { btnClose.style.background='rgba(255,80,80,.15)';btnClose.style.color='#f87171'; });
  btnClose.addEventListener('mouseout',  function() { btnClose.style.background='rgba(255,80,80,0)';btnClose.style.color='#7A7674'; });

  header.appendChild(hLeft);
  header.appendChild(btnClose);
  panel.appendChild(header);

  // Barra búsqueda
  var filterBar = el('div', 'padding:12px 24px;border-bottom:1px solid rgba(103,232,249,.07);flex-shrink:0;display:flex;gap:10px;flex-wrap:wrap;background:rgba(0,0,0,.18);align-items:center;');
  var searchInput = el('input', 'flex:1;min-width:200px;background:rgba(255,255,255,.05);border:1px solid rgba(103,232,249,.2);border-radius:8px;color:#FAFAFA;padding:8px 12px;font-size:13px;font-family:\'Outfit\',sans-serif;outline:none;');
  searchInput.placeholder = '🔍 Buscar bodega...';

  var selTipo = el('select', 'background:#1a1916;border:1px solid rgba(103,232,249,.2);border-radius:8px;color:#FAFAFA;padding:8px 10px;font-size:12px;font-family:\'Outfit\',sans-serif;cursor:pointer;min-width:140px;');
  [['Todos los tipos',''],['Almacén','Almacén'],['VP Alquiler','VP Alquiler'],['VP Venta','VP Venta'],['Bajas','Bajas'],['Proveedor','Proveedor'],['Alist.','Alist.']].forEach(function(o){
    selTipo.appendChild(new Option(o[0], o[1]));
  });

  filterBar.appendChild(searchInput);
  filterBar.appendChild(selTipo);
  panel.appendChild(filterBar);

  // Grid de bodegas
  var gridWrap = el('div', 'overflow:auto;flex:1;padding:20px 24px;');
  var grid = el('div', 'display:grid;grid-template-columns:repeat(auto-fill,minmax(300px,1fr));gap:12px;');
  gridWrap.appendChild(grid);
  panel.appendChild(gridWrap);

  var maxStock = bodegaList[0] ? bodegaList[0].stock : 1;

  function renderCards(list) {
    grid.innerHTML = '';
    if (!list.length) {
      var empty = el('div', 'grid-column:1/-1;text-align:center;padding:40px;color:#7A7674;font-family:\'Outfit\',sans-serif;', 'Sin bodegas para los filtros aplicados');
      grid.appendChild(empty);
      return;
    }
    list.forEach(function(b, i) {
      var tipo  = bodTipo(b.nombre);
      var pct   = maxStock ? (b.stock / maxStock * 100) : 0;
      var short = shortName(b.nombre);

      var card = el('div',
        'background:linear-gradient(145deg,rgba(10,26,18,.95),rgba(8,20,14,.9));' +
        'border:1px solid rgba(103,232,249,.1);border-top:2px solid ' + tipo.color + ';' +
        'border-radius:14px;padding:16px 18px;position:relative;overflow:hidden;' +
        'transition:all .2s cubic-bezier(.4,0,.2,1);cursor:default;'
      );
      card.addEventListener('mouseover', function() {
        card.style.transform = 'translateY(-3px)';
        card.style.borderColor = tipo.color + '55';
        card.style.boxShadow = '0 8px 32px rgba(0,0,0,.5),0 0 20px ' + tipo.color + '18';
      });
      card.addEventListener('mouseout', function() {
        card.style.transform = '';
        card.style.borderColor = 'rgba(103,232,249,.1)';
        card.style.boxShadow = '';
      });

      // Glow orb
      var orb = el('div', 'position:absolute;top:-20px;right:-20px;width:70px;height:70px;border-radius:50%;background:' + tipo.color + ';opacity:.05;filter:blur(10px);pointer-events:none;');
      card.appendChild(orb);

      // Rank badge
      var rank = el('div', 'position:absolute;top:10px;right:12px;font-size:9px;font-family:\'JetBrains Mono\',monospace;color:' + tipo.color + ';opacity:.5;font-weight:700;', '#' + String(i+1).padStart(2,'0'));
      card.appendChild(rank);

      // Tipo badge
      var tipoBadge = el('span',
        'display:inline-block;background:' + tipo.color + '22;color:' + tipo.color + ';' +
        'font-size:9px;font-weight:700;padding:2px 8px;border-radius:20px;' +
        'font-family:\'Outfit\',sans-serif;letter-spacing:.3px;margin-bottom:10px;',
        tipo.label
      );
      card.appendChild(tipoBadge);

      // Nombre
      var nameEl = el('div',
        'font-family:\'Outfit\',sans-serif;font-size:13px;font-weight:600;color:#FAFAFA;' +
        'margin-bottom:12px;line-height:1.3;padding-right:28px;',
        short
      );
      nameEl.title = b.nombre;
      card.appendChild(nameEl);

      // Stock
      var stockRow = el('div', 'display:flex;align-items:baseline;gap:6px;margin-bottom:10px;');
      var stockNum = el('div',
        'font-family:\'JetBrains Mono\',monospace;font-size:22px;font-weight:700;color:' + tipo.color + ';' +
        'line-height:1;letter-spacing:-1px;text-shadow:0 0 20px ' + tipo.color + '55;',
        b.stock.toLocaleString('es-CO')
      );
      var stockLbl = el('div', 'font-size:10px;color:#7A7674;font-family:\'Outfit\',sans-serif;', 'unidades');
      stockRow.appendChild(stockNum);
      stockRow.appendChild(stockLbl);
      card.appendChild(stockRow);

      // Barra de progreso
      var barWrap = el('div', 'background:rgba(255,255,255,.05);border-radius:4px;height:4px;overflow:hidden;');
      var barFill = el('div',
        'height:100%;border-radius:4px;background:linear-gradient(90deg,' + tipo.color + '66,' + tipo.color + ');' +
        'transition:width .8s cubic-bezier(.4,0,.2,1);',
        ''
      );
      barFill.style.width = Math.max(b.stock > 0 ? 4 : 0, Math.min(pct, 100)) + '%';
      barWrap.appendChild(barFill);
      card.appendChild(barWrap);

      grid.appendChild(card);
    });
  }

  function applyBodFilter() {
    var search = (searchInput.value || '').toLowerCase().trim();
    var tipo   = selTipo.value;
    var filtered = bodegaList.filter(function(b) {
      if (tipo && bodTipo(b.nombre).label !== tipo) return false;
      if (search && !b.nombre.toLowerCase().includes(search) && !shortName(b.nombre).toLowerCase().includes(search)) return false;
      return true;
    });
    renderCards(filtered);
  }

  searchInput.addEventListener('input',  applyBodFilter);
  selTipo.addEventListener('change', applyBodFilter);

  // Render inicial
  renderCards(bodegaList);

  overlay.appendChild(panel);
  document.body.appendChild(overlay);

  // Keyframes (reutiliza los mismos del drill modal si ya existen)
  if (!document.getElementById('inv-drill-keyframes')) {
    var style = document.createElement('style');
    style.id = 'inv-drill-keyframes';
    style.textContent = '@keyframes fadeIn{from{opacity:0}to{opacity:1}}@keyframes slideUp{from{opacity:0;transform:translateY(24px) scale(.97)}to{opacity:1;transform:none}}';
    document.head.appendChild(style);
  }
}

window.invOpenBodegasModal = invOpenBodegasModal;

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
window.loadInventarioData = loadInventarioData;
window.invGetRaw = function() { return INV_RAW; };

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