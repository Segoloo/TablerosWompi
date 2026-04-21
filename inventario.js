/**
 * ╔══════════════════════════════════════════════════════════════════╗
 * ║  inventario.js — Dashboard Inventario Wompi                     ║
 * ║  Carga stock_wompi_filtrado.json.gz                             ║
 * ║  KPIs: Totales, Bodega, Comercio, Técnico, Gestores, Ingenico,  ║
 * ║         OPL + filtros Negocio / Categoría / Referencia / Bodega ║
 * ╚══════════════════════════════════════════════════════════════════╝
 */

'use strict';

// ── Estado global del módulo ──────────────────────────────────────
let INV_RAW      = null;   // datos crudos del .json.gz
let INV_FILTERED = [];     // después de aplicar filtros UI

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

// ── Categorización por nombre (lógica DAX traducida a JS) ─────────
function invCategoria(nombre) {
  if (!nombre) return 'KIT POP VP';
  // normalizar
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

// ── Negocio por subtipo ───────────────────────────────────────────
function invNegocio(subtipo) {
  const s = (subtipo || '').trim().toUpperCase();
  if (s === 'WOMPI')    return 'CB';
  if (s === 'WOMPI VP') return 'VP';
  if (s === 'EQUIPO VP')return 'VP';
  if (s === 'VP')       return 'VP';
  return 'CB';
}

// ── Patrón GW (gestores/empleados wompi) ─────────────────────────
const GW_RE = /^GW\d+$/i;

// ── Sumar cantidad de un arreglo de filas ────────────────────────
function sumCantidad(rows) {
  return rows.reduce((acc, r) => acc + (parseInt(r['Cantidad']) || 0), 0);
}

// ── Formato número ────────────────────────────────────────────────
function fmtN(n) {
  return n.toLocaleString('es-CO');
}
function fmtPct(num, den) {
  if (!den) return '0.0%';
  return (num / den * 100).toFixed(1) + '%';
}

// ══════════════════════════════════════════════════════════════════
//  CARGA DEL JSON.GZ
// ══════════════════════════════════════════════════════════════════
async function loadInventarioData() {
  if (INV_RAW) return; // ya cargado
  try {
    const res = await fetch(`stock_wompi_filtrado.json.gz?t=${Date.now()}`);
    if (!res.ok) throw new Error('HTTP ' + res.status);
    const buf       = await res.arrayBuffer();
    const ds        = new DecompressionStream('gzip');
    const writer    = ds.writable.getWriter();
    writer.write(new Uint8Array(buf));
    writer.close();
    const reader    = ds.readable.getReader();
    const chunks    = [];
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
    console.log(`[Inventario] ${INV_RAW.length} filas cargadas`);
  } catch (e) {
    console.error('[Inventario] Error cargando datos:', e);
    INV_RAW = [];
  }
}

// ══════════════════════════════════════════════════════════════════
//  POBLAR FILTROS
// ══════════════════════════════════════════════════════════════════
function _invPopulateFilters() {
  if (!INV_RAW?.length) return;

  // Categorías fijas
  const cats = ['Todas','Rollos','Pin pad','Forros','Accesorios','SIM','KIT POP VP','Datáfonos'];
  _invSetSelect('inv-f-categoria', cats);

  // Referencias únicas (columna Nombre)
  const refs = [...new Set(INV_RAW.map(r => r['Nombre']).filter(Boolean))].sort();
  _invSetSelect('inv-f-referencia', ['Todas', ...refs]);

  // Bodegas
  const bods = [...INV_BODEGAS].sort();
  _invSetSelect('inv-f-bodega', ['Todas', ...bods]);
}

function _invSetSelect(id, opts) {
  const el = document.getElementById(id);
  if (!el) return;
  const cur = el.value;
  el.innerHTML = opts.map(o => `<option value="${o}">${o}</option>`).join('');
  if (opts.includes(cur)) el.value = cur;
}

// ══════════════════════════════════════════════════════════════════
//  APLICAR FILTROS
// ══════════════════════════════════════════════════════════════════
function invApplyFilters() {
  if (!INV_RAW) { INV_FILTERED = []; return; }

  const negocio   = document.getElementById('inv-f-negocio')?.value   || '';
  const categoria = document.getElementById('inv-f-categoria')?.value || '';
  const referencia= document.getElementById('inv-f-referencia')?.value|| '';
  const bodega    = document.getElementById('inv-f-bodega')?.value    || '';

  INV_FILTERED = INV_RAW.filter(r => {
    if (negocio   && negocio   !== 'Todos' && invNegocio(r['Subtipo'])    !== negocio)    return false;
    if (categoria && categoria !== 'Todas' && invCategoria(r['Nombre'])   !== categoria)  return false;
    if (referencia&& referencia!== 'Todas' && r['Nombre']                 !== referencia) return false;
    if (bodega    && bodega    !== 'Todas' && (r['Nombre de la ubicación']||'').trim() !== bodega.trim()) return false;
    return true;
  });

  _invRenderKPIs();
}

window.invApplyFilters  = invApplyFilters;
window.invResetFilters  = function() {
  ['inv-f-negocio','inv-f-categoria','inv-f-referencia','inv-f-bodega'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.selectedIndex = 0;
  });
  invApplyFilters();
};

// ══════════════════════════════════════════════════════════════════
//  RENDER KPIs
// ══════════════════════════════════════════════════════════════════
function _invRenderKPIs() {
  const rows = INV_FILTERED;

  // ── Totales ──
  const total = sumCantidad(rows);

  // ── En Bodega ──
  const enBodega = rows.filter(r => INV_BODEGAS.has((r['Nombre de la ubicación']||'').trim()));
  const unBodega = sumCantidad(enBodega);

  // ── En Comercio ──
  const enComercio = rows.filter(r => (r['Tipo de ubicación']||'').trim() === 'Site');
  const unComercio = sumCantidad(enComercio);

  // ── Técnico Lineacom ──
  const enTecnico = rows.filter(r => (r['Tipo de ubicación']||'').trim() === 'Staff');
  const unTecnico = sumCantidad(enTecnico);

  // ── Gestores / Empleados Wompi (GW###) ──
  const enGW  = rows.filter(r => GW_RE.test((r['Código de ubicación']||'').trim()));
  const unGW  = sumCantidad(enGW);

  // ── Ingenico ──
  const enIngenico = rows.filter(r => (r['Nombre de la ubicación']||'').trim() === 'ALMACEN INGENICO - PROVEEDOR WOMPI');
  const unIngenico = sumCantidad(enIngenico);

  // ── OPL (Supplier) ──
  const enOPL  = rows.filter(r => (r['Tipo de ubicación']||'').trim() === 'Supplier');
  const unOPL  = sumCantidad(enOPL);

  const kpis = [
    { label:'UNIDADES TOTALES',        value: fmtN(total),       sub:'Suma de cantidades', color:'var(--verde-menta)', icon:'📦' },
    { label:'# BODEGAS',               value: '42',              sub:'Almacenes Wompi',    color:'var(--azul-cielo)',  icon:'🏭' },
    { label:'UNIDADES EN BODEGA',       value: fmtN(unBodega),    sub: fmtPct(unBodega,total)+' del total', color:'var(--verde-menta)', icon:'🏪' },
    { label:'% EN BODEGA',              value: fmtPct(unBodega,total), sub: fmtN(unBodega)+' uds',          color:'var(--verde-lima)',  icon:'📊' },
    { label:'UNIDADES EN COMERCIO',     value: fmtN(unComercio),  sub: fmtPct(unComercio,total)+' del total', color:'var(--azul-cielo)', icon:'🏬' },
    { label:'% EN COMERCIO',            value: fmtPct(unComercio,total), sub: fmtN(unComercio)+' uds',       color:'var(--azul-cielo)', icon:'📊' },
    { label:'UNIDADES EN TÉC. LINEACOM',value: fmtN(unTecnico),   sub: fmtPct(unTecnico,total)+' del total', color:'#FFC04D',           icon:'🔧' },
    { label:'% EN TÉC. LINEACOM',       value: fmtPct(unTecnico,total), sub: fmtN(unTecnico)+' uds',        color:'#FFC04D',           icon:'📊' },
    { label:'UNIDS. GEST. Y EMPL_WOMPI',value: fmtN(unGW),        sub: fmtPct(unGW,total)+' del total',     color:'#C084FC',           icon:'👤' },
    { label:'% GEST Y EMPLEADO WOMPI',  value: fmtPct(unGW,total),sub: fmtN(unGW)+' uds',                   color:'#C084FC',           icon:'📊' },
    { label:'UNIDADES EN INGENICO',     value: fmtN(unIngenico),  sub: fmtPct(unIngenico,total)+' del total', color:'#F87171',         icon:'🔌' },
    { label:'% EN INGENICO',            value: fmtPct(unIngenico,total), sub: fmtN(unIngenico)+' uds',       color:'#F87171',         icon:'📊' },
    { label:'UNIDADES EN OPL',          value: fmtN(unOPL),       sub: fmtPct(unOPL,total)+' del total',    color:'#FB923C',           icon:'🚚' },
    { label:'% EN OPL',                 value: fmtPct(unOPL,total),sub: fmtN(unOPL)+' uds',                 color:'#FB923C',           icon:'📊' },
  ];

  const grid = document.getElementById('inv-kpi-grid');
  if (!grid) return;
  grid.innerHTML = kpis.map(k => `
    <div class="kpi-card inv-kpi" style="border-top:3px solid ${k.color};">
      <div class="kpi-icon" style="font-size:22px;margin-bottom:6px;">${k.icon}</div>
      <div class="kpi-label" style="font-size:10px;font-weight:700;letter-spacing:.8px;color:var(--muted);text-transform:uppercase;">${k.label}</div>
      <div class="kpi-value" style="font-size:28px;font-weight:800;color:${k.color};font-family:'Syne',sans-serif;line-height:1.1;margin:6px 0 4px;">${k.value}</div>
      <div class="kpi-sub" style="font-size:11px;color:var(--muted);">${k.sub}</div>
    </div>
  `).join('');
}

// ══════════════════════════════════════════════════════════════════
//  ENTRY POINT — llamado desde dashboard.js _selectBoardTab
// ══════════════════════════════════════════════════════════════════
async function renderInventarioPrincipal() {
  const panel = document.getElementById('panel-inv-principal');
  if (!panel) return;

  // Si aún no hay datos, mostrar spinner y cargar
  if (!INV_RAW) {
    document.getElementById('inv-kpi-grid').innerHTML =
      '<div class="loading"><div class="spinner"></div><span>Cargando inventario...</span></div>';
    await loadInventarioData();
    _invPopulateFilters();
  }

  INV_FILTERED = INV_RAW ? [...INV_RAW] : [];
  _invRenderKPIs();
}

window.renderInventarioPrincipal = renderInventarioPrincipal;