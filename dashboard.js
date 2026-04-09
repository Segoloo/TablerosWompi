/**
 * ╔══════════════════════════════════════════════════════════════════╗
 * ║  dashboard.js — Lógica de KPIs, filtros y renderizado           ║
 * ║  LINEACOM · Dashboard Tracking VP Wompi                         ║
 * ║                                                                  ║
 * ║  Lógica de cálculo alineada 100% con vp.py:                     ║
 * ║  · Filtro global: solo registros desde 2026-03-01               ║
 * ║  · VT : TIPO DE SOLICITUD FACTURACIÓN === exacto (VISITA...)    ║
 * ║  · OPLG: TIPO DE SOLICITUD FACTURACIÓN === exacto (ENVIO...)    ║
 * ║  · Devueltos: contiene DEVOLUCION | DEVUELTO | REMITENTE        ║
 * ║  · n_alistados: total - en_alistamiento                         ║
 * ╚══════════════════════════════════════════════════════════════════╝
 */

'use strict';

// ══════════════════════════════════════════════════════════════════
//  ESTADO GLOBAL
// ══════════════════════════════════════════════════════════════════
let RAW_DATA      = [];   // filas crudas del JSON (todas)
let FILTERED      = [];   // filas después de filtros de UI
let chartInstances = {};
let tablePage     = 1;
let sortCol       = -1;
let sortDir       = 1;
let tableSearchTerm = '';
let filteredForTable = [];

const TABLE_PAGE_SIZE = 50;

// Fecha de corte global — permite datos históricos completos
// Originalmente solo marzo 2026, ahora se puede configurar desde el filtro de UI
// Si no hay filtro de fecha, se muestran TODOS los datos
const MARCH_2026 = new Date(2020, 0, 1);   // fecha mínima muy antigua para no filtrar historial

// ══════════════════════════════════════════════════════════════════
//  CONSTANTES VT / OPLG  (coinciden exactamente con vp.py)
// ══════════════════════════════════════════════════════════════════
const VT_EXACT   = "VISITA DATAFONO+KIT POP+CAPACITACION";
const OPLG_EXACT = "ENVIO DATAFONO+KIT POP";

// ══════════════════════════════════════════════════════════════════
//  PALETA
// ══════════════════════════════════════════════════════════════════
const WOMPI_COLORS = [
  '#B0F2AE','#99D1FC','#DFFF61','#00C87A',
  '#FF5C5C','#FFC04D','#7B8CDE','#F49D6E',
  '#C4F0C4','#7BC8FB','#A8E6CF','#FFD3B6',
];

const CHART_OPTS = {
  tooltip: {
    backgroundColor:'rgba(24,23,21,.95)',
    titleColor:'#B0F2AE', bodyColor:'#FAFAFA',
    borderColor:'rgba(176,242,174,.2)', borderWidth:1, padding:12,
  },
};

// ══════════════════════════════════════════════════════════════════
//  AUTENTICACIÓN
// ══════════════════════════════════════════════════════════════════
const USERS = { wompi: 'tracking2025', lineacom: 'VP2025*' };

function doLogin() {
  const u = document.getElementById('login-user').value.trim();
  const p = document.getElementById('login-pass').value.trim();
  if (USERS[u] && USERS[u] === p) {
    document.getElementById('login-screen').classList.add('hidden');
    document.getElementById('app').classList.add('visible');
    initDashboard();
  } else {
    const err = document.getElementById('login-error');
    err.style.display = 'block';
    setTimeout(() => err.style.display = 'none', 3000);
  }
}

function doLogout() {
  document.getElementById('app').classList.remove('visible');
  document.getElementById('login-screen').classList.remove('hidden');
  document.getElementById('login-user').value = '';
  document.getElementById('login-pass').value = '';
}

// ══════════════════════════════════════════════════════════════════
//  CARGA DE DATOS
// ══════════════════════════════════════════════════════════════════
function loadData() {
  const url = `data.json?t=${Date.now()}`;
  fetch(url)
    .then(r => r.ok ? r.json() : Promise.reject(r.status))
    .then(payload => {
      // data.py ya no envía filas vacías, pero por si acaso:
      RAW_DATA = (payload.rows || []).filter(r => Object.values(r).some(v => v !== '' && v !== null));
      // DEBUG: imprimir columnas del primer row para verificar nombres exactos
      if (RAW_DATA.length > 0) {
        console.log('[Dashboard] Columnas disponibles en data.json:', Object.keys(RAW_DATA[0]));
        // Buscar columnas que contengan "novedad" o "causal"
        const novCols = Object.keys(RAW_DATA[0]).filter(k => {
          const kl = k.toLowerCase();
          return kl.includes('novedad') || kl.includes('causal') || kl.includes('responsable');
        });
        console.log('[Dashboard] Columnas de novedad detectadas:', novCols);
        // Buscar columnas que contengan "devoluci"
        const devCols = Object.keys(RAW_DATA[0]).filter(k => k.toLowerCase().includes('devoluci'));
        console.log('[Dashboard] Columnas de devolución detectadas:', devCols);
        // Muestra de valores de novedad en primeras 5 filas
        const sample = RAW_DATA.slice(0,5).map(r => {
          const obj = {};
          novCols.forEach(c => { obj[c] = r[c]; });
          return obj;
        });
        console.log('[Dashboard] Muestra novedades (primeras 5 filas):', sample);
      }
      document.getElementById('last-update').textContent =
        `Actualizado: ${payload.generado || '—'}`;
      populateFilters();
      applyDateFilter();   // aplica el filtro global de marzo 2026
      renderAll();
    })
    .catch(() => {
      // Demo data si no hay data.json
      RAW_DATA = getDemoData();
      populateFilters();
      applyDateFilter();
      renderAll();
    });
}

// ══════════════════════════════════════════════════════════════════
//  HELPER: acceso a columna insensible a mayúsculas / variantes
// ══════════════════════════════════════════════════════════════════
function getCol(row, ...keys) {
  for (const k of keys) {
    if (row[k] !== undefined && row[k] !== null) return String(row[k]);
    // búsqueda insensible
    const kl = k.toUpperCase();
    for (const rk of Object.keys(row)) {
      if (rk.toUpperCase() === kl) return String(row[rk]);
    }
  }
  return '';
}

// ══════════════════════════════════════════════════════════════════
//  PARSEO DE FECHAS
// ══════════════════════════════════════════════════════════════════
function parseDate(s) {
  if (!s || typeof s !== 'string') return null;
  s = s.trim();
  if (!s || s === 'NAN' || s === 'nan') return null;

  const fmts = [
    // dd/mm/yyyy
    s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/) ?
      new Date(+RegExp.$3, +RegExp.$2 - 1, +RegExp.$1) : null,
    // yyyy-mm-dd
    s.match(/^(\d{4})-(\d{2})-(\d{2})/) ?
      new Date(+RegExp.$1, +RegExp.$2 - 1, +RegExp.$3) : null,
    // dd-mm-yyyy
    s.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/) ?
      new Date(+RegExp.$3, +RegExp.$2 - 1, +RegExp.$1) : null,
  ];

  for (const d of fmts) {
    if (d && !isNaN(d.getTime()) && d.getFullYear() > 2000 && d.getFullYear() < 2100) return d;
  }
  // fallback nativo
  const fb = new Date(s);
  return (!isNaN(fb.getTime()) && fb.getFullYear() > 2000 && fb.getFullYear() < 2100) ? fb : null;
}

function diffDays(a, b) {
  return Math.floor((b - a) / 86400000);
}

// ══════════════════════════════════════════════════════════════════
//  HELPER: busca el primer valor de novedad real en una fila,
//  buscando por coincidencia parcial en el nombre de la columna.
//  Excluye columnas de estado/fecha/transporte que no son novedades.
// ══════════════════════════════════════════════════════════════════
const NOV_KEY_INCLUDES  = ['novedad','novedades','causal','responsable incump'];
const NOV_KEY_EXCLUDES  = ['estado','fecha','solicitud','comercio','guia','transpo','datafon','serial','departam','ciudad','tipolog','cumple','referencia','tipo de sol','id com'];

function findNovedad(r) {
  // 1. Primero buscar en columnas exactas conocidas (más rápido)
  const exactCols = [
    'NOVEDADES','novedades','NOVEDAD','novedad',
    'CAUSAL INCU','causal incu','CAUSAL INC','causal inc',
    'RESPONSABLE INCUMPLIMIENTO','responsable incumplimiento',
    'CAUSAL INCUMPLIMIENTO','causal incumplimiento',
  ];
  for (const col of exactCols) {
    const v = getCol(r, col).trim();
    if (v && v !== '0' && v.toLowerCase() !== 'nan' && v !== '') return { col, val: v };
  }
  // 2. Búsqueda fuzzy: recorrer TODAS las claves del row buscando las que
  //    contengan palabras clave de novedad y NO sean columnas de estado/info general
  for (const k of Object.keys(r)) {
    const kl = k.toLowerCase();
    const isNovCol = NOV_KEY_INCLUDES.some(n => kl.includes(n));
    const isExcluded = NOV_KEY_EXCLUDES.some(x => kl.includes(x));
    if (isNovCol && !isExcluded) {
      const v = String(r[k] || '').trim();
      if (v && v !== '0' && v.toLowerCase() !== 'nan' && v !== '') return { col: k, val: v };
    }
  }
  return null;
}

// ══════════════════════════════════════════════════════════════════
//  HELPER: obtener FECHA LIMITE DE ENTREGA robustamente
//  Busca primero por nombres exactos, luego fuzzy por nombre de columna.
// ══════════════════════════════════════════════════════════════════
function getFechaLimite(r) {
  const exactos = [
    'FECHA LIMITE DE ENTREGA','fecha limite de entrega',
    'FECHA LÍMITE DE ENTREGA','fecha límite de entrega',
    'FECHA LIMITE','fecha limite',
  ];
  for (const col of exactos) {
    const v = getCol(r, col);
    if (v) { const d = parseDate(v); if (d) return d; }
  }
  // Fuzzy: columna que contenga "limite" o "límite" pero NO "solicitud" ni "comercio"
  for (const k of Object.keys(r)) {
    const kl = k.toLowerCase();
    if ((kl.includes('limite') || kl.includes('límite')) &&
        !kl.includes('solicitud') && !kl.includes('comercio')) {
      const d = parseDate(String(r[k] || ''));
      if (d) return d;
    }
  }
  return null;
}

// ══════════════════════════════════════════════════════════════════
//  FILTRO GLOBAL: solo datos desde marzo 2026  (=== vp.py)
// ══════════════════════════════════════════════════════════════════
function applyDateFilter() {
  // Datos históricos completos — sin corte de fecha global
  // El filtro de fecha se aplica desde los filtros de UI (Fecha Desde / Hasta / Mes)
  FILTERED = RAW_DATA.slice();
}

// ══════════════════════════════════════════════════════════════════
//  FILTROS DE UI (estado, tipo, departamento, etc.)
// ══════════════════════════════════════════════════════════════════
function populateFilters() {
  const sets = {
    'f-estado':        new Set(),
    'f-tipo-envio':    new Set(),
    'f-material':      new Set(),
    'f-departamento':  new Set(),
    'f-ciudad':        new Set(),
    'f-transportadora':new Set(),
    'f-mes':           new Set(),
  };

  RAW_DATA.forEach(r => {
    const est  = getCol(r, 'ESTADO DATAFONO', 'estado datafono');
    const tipo = getCol(r, 'TIPO DE SOLICITUD FACTURACIÓN', 'TIPO DE SOLICITUD FACTURACION',
                           'tipo de solicitud facturacion');
    const mat  = getCol(r, 'REFERENCIA DEL DATAFONO', 'REFERENCIA DEL DATAFONOS', 'referencia del datafono');
    const dep  = getCol(r, 'Departamento', 'DEPARTAMENTO', 'departamento');
    const ciu  = getCol(r, 'Ciudad', 'CIUDAD', 'ciudad');
    const tra  = getCol(r, 'TRANSPORTADORA', 'Transportadora', 'transportadora');
    const fs   = getCol(r, 'FECHA DE SOLICITUD', 'fecha de solicitud');

    if (est) sets['f-estado'].add(est);
    if (tipo) sets['f-tipo-envio'].add(tipo);
    if (mat) sets['f-material'].add(mat);
    if (dep) sets['f-departamento'].add(dep);
    if (ciu) sets['f-ciudad'].add(ciu);
    if (tra) sets['f-transportadora'].add(tra);

    const d = parseDate(fs);
    if (d) {
      const mes = d.toLocaleString('es-CO', { month: 'long', year: 'numeric' });
      sets['f-mes'].add(mes);
    }
  });

  for (const [id, vals] of Object.entries(sets)) {
    const sel = document.getElementById(id);
    if (!sel) continue;
    sel.innerHTML = '<option value="">Todos</option>' +
      [...vals].sort().map(v => `<option value="${v}">${v}</option>`).join('');
  }
}

function applyFilters() {
  const estado   = document.getElementById('f-estado')?.value;
  const tipoEnvio= document.getElementById('f-tipo-envio')?.value;
  const material = document.getElementById('f-material')?.value;
  const depto    = document.getElementById('f-departamento')?.value;
  const ciudad   = document.getElementById('f-ciudad')?.value;
  const transp   = document.getElementById('f-transportadora')?.value;
  const mes      = document.getElementById('f-mes')?.value;
  const guia     = document.getElementById('f-guia')?.value?.trim().toUpperCase();
  const idSitio  = document.getElementById('f-idsitio')?.value?.trim().toUpperCase();
  const desde    = document.getElementById('f-fecha-desde')?.value;
  const hasta    = document.getElementById('f-fecha-hasta')?.value;
  const limDesde = document.getElementById('f-limite-desde')?.value;
  const limHasta = document.getElementById('f-limite-hasta')?.value;

  // Primero aplicar filtro de fecha global (marzo 2026)
  applyDateFilter();

  FILTERED = FILTERED.filter(r => {
    const rEst  = getCol(r, 'ESTADO DATAFONO', 'estado datafono').toUpperCase();
    const rTipo = getCol(r, 'TIPO DE SOLICITUD FACTURACIÓN', 'TIPO DE SOLICITUD FACTURACION',
                            'tipo de solicitud facturacion').toUpperCase();
    const rMat  = getCol(r, 'REFERENCIA DEL DATAFONO', 'REFERENCIA DEL DATAFONOS',
                            'referencia del datafono').toUpperCase();
    const rDep  = getCol(r, 'Departamento', 'DEPARTAMENTO', 'departamento').toUpperCase();
    const rCiu  = getCol(r, 'Ciudad', 'CIUDAD', 'ciudad').toUpperCase();
    const rTra  = getCol(r, 'TRANSPORTADORA', 'Transportadora', 'transportadora').toUpperCase();
    const rGui  = getCol(r, 'NÚMERO DE GUIA', 'NUMERO DE GUIA', 'numero de guia',
                            'Numero de Guia').toUpperCase();
    const rId   = getCol(r, 'ID Comercio', 'id comercio', 'Id Comercio').toUpperCase();
    const fs    = getCol(r, 'FECHA DE SOLICITUD', 'fecha de solicitud');
    const fd    = parseDate(fs);
    const fl    = parseDate(getCol(r, 'FECHA LIMITE DE ENTREGA', 'fecha limite de entrega'));

    if (estado   && rEst  !== estado.toUpperCase())  return false;
    if (tipoEnvio&& rTipo !== tipoEnvio.toUpperCase()) return false;
    if (material && rMat  !== material.toUpperCase()) return false;
    if (depto    && rDep  !== depto.toUpperCase())    return false;
    if (ciudad   && rCiu  !== ciudad.toUpperCase())   return false;
    if (transp   && rTra  !== transp.toUpperCase())   return false;
    if (guia     && !rGui.includes(guia))             return false;
    if (idSitio  && !rId.includes(idSitio))           return false;

    if (mes && fd) {
      const rMes = fd.toLocaleString('es-CO', { month:'long', year:'numeric' });
      if (rMes !== mes) return false;
    }
    // Rango fecha solicitud
    if (desde && fd && fd < new Date(desde)) return false;
    if (hasta && fd && fd > new Date(hasta + 'T23:59:59')) return false;
    // Rango fecha límite de entrega
    if (limDesde && fl && fl < new Date(limDesde)) return false;
    if (limHasta && fl && fl > new Date(limHasta + 'T23:59:59')) return false;

    return true;
  });

  tablePage = 1;
  renderAll();
}

function resetFilters() {
  ['f-estado','f-tipo-envio','f-material','f-departamento',
   'f-ciudad','f-transportadora','f-mes'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = '';
  });
  ['f-guia','f-idsitio'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = '';
  });
  ['f-fecha-desde','f-fecha-hasta','f-limite-desde','f-limite-hasta'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = '';
  });

  applyDateFilter();
  tablePage = 1;
  renderAll();
}

// ══════════════════════════════════════════════════════════════════
//  COMPUTE KPIs  ← lógica idéntica a vp.py compute_kpis()
// ══════════════════════════════════════════════════════════════════
function computeKPIs(data) {
  // 1. Excluir cancelados (igual que vp.py)
  const df = data.filter(r =>
    getCol(r, 'ESTADO DATAFONO', 'estado datafono').toUpperCase() !== 'CANCELADO'
  );
  const cancelados = data.length - df.length;
  const total = df.length;

  // 2. Conteo por estado (upper case)
  const ec = {};
  df.forEach(r => {
    const e = getCol(r, 'ESTADO DATAFONO', 'estado datafono').toUpperCase() || 'SIN ESTADO';
    ec[e] = (ec[e] || 0) + 1;
  });

  const entregados      = ec['ENTREGADO']       || 0;
  const en_transito     = ec['EN TRANSITO']     || ec['EN TRÁNSITO'] || 0;
  const programados     = ec['PROGRAMADO']      || ec['VISITA PROGRAMADA'] || 0;
  const en_alistamiento = ec['EN ALISTAMIENTO'] || 0;

  // 3. Devueltos: buscar en TODAS las columnas de la fila cualquier variante de devolución
  //    tanto en el VALOR como en el NOMBRE de la columna
  function isDevolucion(r) {
    for (const [k, v] of Object.entries(r)) {
      const kUp = k.toUpperCase();
      const vUp = String(v || '').toUpperCase();
      // Detectar por valor
      if (vUp.includes('DEVOLUCI') || vUp.includes('DEVOLUCION') ||
          vUp.includes('DEVOLUCIÓN') || vUp.includes('DEVUELTO') ||
          vUp.includes('REMITENTE')) return true;
      // Detectar por nombre de columna (ej: "ESTADO DEVOLUCION", "MOTIVO DEVOLUCION")
      if (kUp.includes('DEVOLUCI') || kUp.includes('DEVOLUCION') || kUp.includes('DEVOLUCIÓN')) {
        if (vUp && vUp !== '' && vUp !== '0' && vUp !== 'NAN') return true;
      }
    }
    return false;
  }
  const devueltos = df.filter(isDevolucion).length;

  // 4. n_alistados = total - en_alistamiento  (igual que vp.py)
  const n_alistados = total - en_alistamiento;

  // 5. VT: coincidencia EXACTA con el valor de vp.py
  const vtRows = df.filter(r =>
    getCol(r,
      'TIPO DE SOLICITUD FACTURACIÓN',
      'TIPO DE SOLICITUD FACTURACION',
      'tipo de solicitud facturacion'
    ).toUpperCase() === VT_EXACT.toUpperCase()
  );

  // 6. OPLG: coincidencia EXACTA con el valor de vp.py
  const olRows = df.filter(r =>
    getCol(r,
      'TIPO DE SOLICITUD FACTURACIÓN',
      'TIPO DE SOLICITUD FACTURACION',
      'tipo de solicitud facturacion'
    ).toUpperCase() === OPLG_EXACT.toUpperCase()
  );

  const entVT = vtRows.filter(r =>
    getCol(r, 'ESTADO DATAFONO', 'estado datafono').toUpperCase() === 'ENTREGADO'
  ).length;

  const entOL = olRows.filter(r =>
    getCol(r, 'ESTADO DATAFONO', 'estado datafono').toUpperCase() === 'ENTREGADO'
  ).length;

  // 7. Programados VT (igual que vp.py: estado === "VISITA PROGRAMADA")
  const programados_vt = vtRows.filter(r =>
    getCol(r, 'ESTADO DATAFONO', 'estado datafono').toUpperCase() === 'VISITA PROGRAMADA'
  ).length;

  // 8. ANS Oportunidad
  const entDf = df.filter(r =>
    getCol(r, 'ESTADO DATAFONO', 'estado datafono').toUpperCase() === 'ENTREGADO'
  );
  const cumpleOport = entDf.filter(r =>
    getCol(r, 'CUMPLE ANS', 'cumple ans', 'Cumple Ans').toUpperCase() === 'SI'
  ).length;

  const pctOport   = entDf.length ? Math.round(cumpleOport / entDf.length * 100) : 0;
  const pctCalidad = entregados   ? Math.round((entregados - devueltos) / entregados * 100) : 100;

  // 9. Vencen hoy / vencidas (incluye VT y OPLG)
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const vencenHoyRows = df.filter(r => {
    const lim = parseDate(getCol(r, 'FECHA LIMITE DE ENTREGA', 'fecha limite de entrega'));
    const est = getCol(r, 'ESTADO DATAFONO', 'estado datafono').toUpperCase();
    if (!lim) return false;
    const limD = new Date(lim); limD.setHours(0, 0, 0, 0);
    return limD.getTime() === today.getTime() && est !== 'ENTREGADO' && est !== 'CANCELADO';
  });
  const vencenHoy = vencenHoyRows.length;

  // Vencidas: TODOS los registros (VT + OPLG) no entregados con fecha limite pasada
  const vencidasRows = df.filter(r => {
    const lim = parseDate(getCol(r, 'FECHA LIMITE DE ENTREGA', 'fecha limite de entrega'));
    const est = getCol(r, 'ESTADO DATAFONO', 'estado datafono').toUpperCase();
    return lim && lim < today && est !== 'ENTREGADO' && est !== 'CANCELADO';
  });
  const vencidas = vencidasRows.length;

  // 10. Primer intento: entregados sin novedades ni causal ni responsable de incumplimiento
  const FAILED_COLS = ['NOVEDADES','novedades','CAUSAL INCU','causal incu','RESPONSABLE INCUMPLIMIENTO','responsable incumplimiento','CAUSAL INC','causal inc'];
  const primerIntentoRows = entDf.filter(r => {
    for (const col of FAILED_COLS) {
      const v = getCol(r, col);
      if (v && v.trim() !== '' && v.trim() !== '0') return false;
    }
    return true;
  });
  const primerIntento = primerIntentoRows.length;

  const pct = (a, b) => b ? Math.round(a / b * 100) : 0;

  return {
    total, cancelados, entregados,
    en_transito, programados, en_alistamiento,
    n_alistados,
    devueltos,
    totalVT:      vtRows.length,
    entVT,
    programados_vt,
    totalOL:      olRows.length,
    entOL,
    pctEntregado:    pct(entregados, total),
    pctTransito:     pct(en_transito, total),
    pctAlistamiento: pct(en_alistamiento, total),
    pctNAlistados:   pct(n_alistados, total),
    pctVT:           pct(entVT, vtRows.length),
    pctOL:           pct(entOL, olRows.length),
    pctOport, pctCalidad,
    vencenHoy, vencenHoyRows,
    vencidas, vencidasRows,
    primerIntento, primerIntentoRows,
    pctPrimerIntento: pct(primerIntento, entregados),
    ec,
    entregadosRows: entDf,
    vtRows, olRows,
    devueltosRows: df.filter(isDevolucion),
  };
}

// ══════════════════════════════════════════════════════════════════
//  MODAL DRILLDOWN
// ══════════════════════════════════════════════════════════════════
function openDrillModal(title, rows, cols) {
  let modal = document.getElementById('drill-modal');
  if (!modal) {
    modal = document.createElement('div');
    modal.id = 'drill-modal';
    modal.style.cssText = `
      position:fixed;inset:0;z-index:99999;display:flex;align-items:center;justify-content:center;
      background:rgba(0,0,0,.75);backdrop-filter:blur(6px);padding:20px;
    `;
    modal.innerHTML = `
      <div style="background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);
        width:100%;max-width:1000px;max-height:85vh;display:flex;flex-direction:column;
        box-shadow:0 24px 80px rgba(0,0,0,.8);">
        <div style="display:flex;align-items:center;justify-content:space-between;padding:20px 24px;
          border-bottom:1px solid var(--border);flex-shrink:0;">
          <div>
            <div id="drill-modal-title" style="font-family:'Syne',sans-serif;font-size:18px;font-weight:700;color:var(--verde-menta)"></div>
            <div id="drill-modal-count" style="font-size:12px;color:var(--muted);margin-top:2px"></div>
          </div>
          <div style="display:flex;gap:8px;align-items:center">
            <button onclick="exportDrillExcel()" style="background:var(--surface2);border:1px solid var(--border);color:var(--verde-menta);padding:7px 14px;border-radius:8px;cursor:pointer;font-size:12px;font-family:'Outfit',sans-serif">⬇ Excel</button>
            <button onclick="document.getElementById('drill-modal').remove()" style="background:var(--surface2);border:1px solid var(--border);color:var(--muted);width:32px;height:32px;border-radius:8px;cursor:pointer;font-size:16px;display:flex;align-items:center;justify-content:center">×</button>
          </div>
        </div>
        <div id="drill-modal-body" style="overflow:auto;flex:1;padding:0 4px 4px"></div>
      </div>`;
    modal.addEventListener('click', e => { if (e.target === modal) modal.remove(); });
    document.body.appendChild(modal);
  }

  window._drillData  = rows;
  window._drillCols  = cols;
  document.getElementById('drill-modal-title').textContent = title;
  document.getElementById('drill-modal-count').textContent = `${rows.length} registros`;

  const DRILL_COLS = cols || [
    { label:'Comercio',        fn:r=>getCol(r,'Nombre del comercio','nombre del comercio','NOMBRE DEL COMERCIO') },
    { label:'ID Sitio',        fn:r=>getCol(r,'ID Comercio','id comercio') },
    { label:'Guía',            fn:r=>getCol(r,'NÚMERO DE GUIA','NUMERO DE GUIA','numero de guia') },
    { label:'Fecha Límite',    fn:r=>{ const d=parseDate(getCol(r,'FECHA LIMITE DE ENTREGA','fecha limite de entrega')); return d?d.toLocaleDateString('es-CO'):'—'; } },
    { label:'Transportadora',  fn:r=>getCol(r,'TRANSPORTADORA','Transportadora','transportadora') },
    { label:'Estado',          fn:r=>getCol(r,'ESTADO DATAFONO','estado datafono'), isStatus:true },
    { label:'Novedad',         fn:r=>getCol(r,'NOVEDADES','novedades')||getCol(r,'CAUSAL INCU','causal incu')||'—' },
    { label:'Tipo',            fn:r=>getCol(r,'TIPO DE SOLICITUD FACTURACIÓN','TIPO DE SOLICITUD FACTURACION','tipo de solicitud facturacion') },
    { label:'Departamento',    fn:r=>getCol(r,'Departamento','DEPARTAMENTO','departamento') },
  ];

  const body = document.getElementById('drill-modal-body');
  if (!rows.length) {
    body.innerHTML = '<div style="text-align:center;padding:48px;color:var(--muted)">Sin registros</div>';
    return;
  }
  body.innerHTML = `<table style="width:100%;border-collapse:collapse;font-size:12px">
    <thead style="position:sticky;top:0;background:var(--surface)">
      <tr>${DRILL_COLS.map(c=>`<th style="padding:10px 12px;text-align:left;color:var(--verde-menta);font-weight:600;border-bottom:1px solid var(--border);white-space:nowrap">${c.label}</th>`).join('')}</tr>
    </thead>
    <tbody>${rows.map((r,i)=>`<tr style="background:${i%2?'transparent':'rgba(255,255,255,.02)'}">
      ${DRILL_COLS.map(c=>{
        const v = c.fn(r)||'—';
        return c.isStatus ? `<td style="padding:8px 12px"><span class="status-pill ${statusClass(v)}">${v}</span></td>`
                          : `<td style="padding:8px 12px;color:var(--blanco)">${v}</td>`;
      }).join('')}
    </tr>`).join('')}</tbody>
  </table>`;
}

function exportDrillExcel() {
  if (!window._drillData || !window._drillData.length) return;
  const DRILL_COLS = window._drillCols || [
    { label:'Comercio',        fn:r=>getCol(r,'Nombre del comercio','nombre del comercio','NOMBRE DEL COMERCIO') },
    { label:'ID Sitio',        fn:r=>getCol(r,'ID Comercio','id comercio') },
    { label:'Guía',            fn:r=>getCol(r,'NÚMERO DE GUIA','NUMERO DE GUIA','numero de guia') },
    { label:'Fecha Límite',    fn:r=>{ const d=parseDate(getCol(r,'FECHA LIMITE DE ENTREGA','fecha limite de entrega')); return d?d.toLocaleDateString('es-CO'):'—'; } },
    { label:'Transportadora',  fn:r=>getCol(r,'TRANSPORTADORA','Transportadora','transportadora') },
    { label:'Estado',          fn:r=>getCol(r,'ESTADO DATAFONO','estado datafono') },
    { label:'Novedad',         fn:r=>getCol(r,'NOVEDADES','novedades')||getCol(r,'CAUSAL INCU','causal incu')||'' },
    { label:'Tipo',            fn:r=>getCol(r,'TIPO DE SOLICITUD FACTURACIÓN','TIPO DE SOLICITUD FACTURACION','tipo de solicitud facturacion') },
    { label:'Departamento',    fn:r=>getCol(r,'Departamento','DEPARTAMENTO','departamento') },
  ];
  const data = window._drillData.map(r=>{ const o={}; DRILL_COLS.forEach(c=>{ o[c.label]=c.fn(r)||''; }); return o; });
  exportToExcel(data, 'KPI_Drilldown');
}

// ══════════════════════════════════════════════════════════════════
//  RENDER ALL
// ══════════════════════════════════════════════════════════════════
function renderAll() {
  const k = computeKPIs(FILTERED);
  renderKPIs(k);
  renderCharts(k);
  renderDeptTable();
  renderANSAlerts(k);
  renderMainTable();
  const badge = document.getElementById('topbar-badge');
  if (badge) badge.textContent = `${FILTERED.length} registros`;
  const fd = document.getElementById('footer-date');
  if (fd) fd.textContent = new Date().toLocaleDateString('es-CO', { dateStyle: 'long' });
}

// ══════════════════════════════════════════════════════════════════
//  RENDER KPI CARDS
// ══════════════════════════════════════════════════════════════════
function renderKPIs(k) {
  const grid = document.getElementById('kpi-grid');
  if (!grid) return;

  const cards = [
    { label:'Total Solicitados',  value:k.total,          color:'green', icon:'📦',
      sub:`Sin cancelados (${k.cancelados} cancelados)`, rows: FILTERED },
    { label:'Alistados',          value:`${k.n_alistados} (${k.pctNAlistados}%)`,
      color:'lime',  icon:'⚙️',  sub:'Total – en alistamiento', pct:k.pctNAlistados, pctColor:'lime',
      rows: FILTERED.filter(r=>getCol(r,'ESTADO DATAFONO','estado datafono').toUpperCase()!=='EN ALISTAMIENTO'&&getCol(r,'ESTADO DATAFONO','estado datafono').toUpperCase()!=='CANCELADO') },
    { label:'Entregados',         value:k.entregados,     color:'selva', icon:'✅',
      sub:`${k.pctEntregado}%`, pct:k.pctEntregado,
      rows: k.entregadosRows },
    { label:'En Tránsito',        value:k.en_transito,    color:'blue',  icon:'🚚',
      sub:`${k.pctTransito}%`, pct:k.pctTransito, pctColor:'blue',
      rows: FILTERED.filter(r=>{ const e=getCol(r,'ESTADO DATAFONO','estado datafono').toUpperCase(); return e==='EN TRANSITO'||e==='EN TRÁNSITO'; }) },
    { label:'En Alistamiento',    value:k.en_alistamiento,color:'lime',  icon:'🔧',
      sub:`${k.pctAlistamiento}%`, pct:k.pctAlistamiento, pctColor:'lime',
      rows: FILTERED.filter(r=>getCol(r,'ESTADO DATAFONO','estado datafono').toUpperCase()==='EN ALISTAMIENTO') },
    { label:'Devueltos',          value:k.devueltos,      color:'danger',icon:'↩️',
      sub:`${k.total ? Math.round(k.devueltos/k.total*100) : 0}% del total`,
      pct: k.total ? Math.round(k.devueltos/k.total*100) : 0,
      pctColor:'danger', alert:k.devueltos > 0, rows: k.devueltosRows },
    { label:'% Oportunidad ANS',  value:k.pctOport+'%',   color:'green', icon:'🎯',
      sub:'Entregas en plazo', pct:k.pctOport,
      rows: k.entregadosRows.filter(r=>getCol(r,'CUMPLE ANS','cumple ans').toUpperCase()==='SI') },
    { label:'% Calidad',          value:k.pctCalidad+'%', color:'blue',  icon:'💎',
      sub:'Sin devoluciones', pct:k.pctCalidad, pctColor:'blue',
      rows: k.entregadosRows.filter(r=>{ const e=getCol(r,'ESTADO DATAFONO','estado datafono').toUpperCase(); return !e.includes('DEVOLUCION')&&!e.includes('DEVUELTO')&&!e.includes('REMITENTE'); }) },
    { label:'Visita Técnica',
      value:`${k.entVT} ejec / ${k.programados_vt} prog / ${k.totalVT} total`,
      color:'lime', icon:'🔧', sub:`${k.pctVT}% ejecutado`, pct:k.pctVT, pctColor:'lime',
      isVT: true, vtEjec: k.entVT, vtProg: k.programados_vt, vtTotal: k.totalVT, vtPct: k.pctVT,
      rows: k.vtRows },
    { label:'Op. Logístico',      value:`${k.entOL}/${k.totalOL}`, color:'blue', icon:'📮',
      sub:`${k.pctOL}% entregado`, pct:k.pctOL, rows: k.olRows },
    { label:'Vencen Hoy',         value:k.vencenHoy,      color:'warn',  icon:'⏰',
      sub:'Sin entregar, límite hoy', alert:k.vencenHoy > 0, rows: k.vencenHoyRows },
    { label:'Vencidas ANS',       value:k.vencidas,       color:'danger',icon:'🚨',
      sub:'Fuera de ANS (VT + OPLG)', alert:k.vencidas > 0, rows: k.vencidasRows },
    { label:'1er Intento',        value:k.primerIntento,  color:'selva', icon:'🎯',
      sub:`${k.pctPrimerIntento}% del entregado`, pct:k.pctPrimerIntento, rows: k.primerIntentoRows },
  ];

  grid.innerHTML = cards.map((c, i) => {
    const clickAttr = c.rows ? `onclick="openDrillModal('${c.label}', window._kpiRows[${i}])" style="cursor:pointer"` : '';
    if (c.isVT) {
      return `
    <div class="kpi-card lime vt-special fade-up" ${clickAttr} style="animation-delay:${i*.04}s;cursor:pointer" title="Ver listado">
      ${c.alert ? '<div class="kpi-alert-badge"></div>' : ''}
      <div class="kpi-drill-hint">Ver listado ↗</div>
      <div style="display:flex;align-items:flex-start;justify-content:space-between;gap:12px">
        <div style="flex:1">
          <span class="kpi-icon">${c.icon}</span>
          <div class="kpi-label">${c.label}</div>
          <div style="display:flex;gap:16px;margin:10px 0;flex-wrap:wrap">
            <div style="text-align:center">
              <div style="font-family:'JetBrains Mono',monospace;font-size:26px;font-weight:700;color:var(--verde-lima);line-height:1;text-shadow:0 0 20px rgba(223,255,97,.4)">${c.vtEjec}</div>
              <div style="font-size:9px;color:var(--muted);text-transform:uppercase;letter-spacing:1px;margin-top:4px">Ejecutadas</div>
            </div>
            <div style="text-align:center;opacity:.7">
              <div style="font-family:'JetBrains Mono',monospace;font-size:26px;font-weight:700;color:var(--azul-cielo);line-height:1">${c.vtProg}</div>
              <div style="font-size:9px;color:var(--muted);text-transform:uppercase;letter-spacing:1px;margin-top:4px">Programadas</div>
            </div>
            <div style="text-align:center;opacity:.6">
              <div style="font-family:'JetBrains Mono',monospace;font-size:26px;font-weight:700;color:var(--verde-menta);line-height:1">${c.vtTotal}</div>
              <div style="font-size:9px;color:var(--muted);text-transform:uppercase;letter-spacing:1px;margin-top:4px">Total</div>
            </div>
          </div>
          <div class="progress-wrap">
            <div class="progress-label"><span>Ejecución</span><span style="color:var(--verde-lima);font-weight:700">${c.vtPct}%</span></div>
            <div class="progress-track" style="height:5px">
              <div class="progress-fill lime" style="width:${Math.min(c.vtPct,100)}%"></div>
            </div>
          </div>
        </div>
        <div style="position:relative;width:90px;height:90px;flex-shrink:0">
          <canvas id="kpi-vt-donut" width="90" height="90"></canvas>
          <div style="position:absolute;inset:0;display:flex;flex-direction:column;align-items:center;justify-content:center;pointer-events:none">
            <div style="font-family:'JetBrains Mono',monospace;font-size:16px;font-weight:700;color:var(--verde-lima);line-height:1">${c.vtPct}%</div>
          </div>
        </div>
      </div>
    </div>`;
    }
    return `
    <div class="kpi-card ${c.color} fade-up" ${clickAttr} style="animation-delay:${i*.04}s;${c.rows?'cursor:pointer':''}" title="${c.rows?'Ver listado':''}">
      ${c.alert ? '<div class="kpi-alert-badge"></div>' : ''}
      ${c.rows ? '<div class="kpi-drill-hint">Ver listado ↗</div>' : ''}
      <span class="kpi-icon">${c.icon}</span>
      <div class="kpi-label">${c.label}</div>
      <div class="kpi-value ${c.color}">${c.value}</div>
      <div class="kpi-sub">${c.sub}</div>
      ${c.pct !== undefined ? `
        <div class="progress-wrap">
          <div class="progress-track">
            <div class="progress-fill ${c.pctColor||'green'}" style="width:${Math.min(c.pct,100)}%"></div>
          </div>
        </div>` : ''}
    </div>`;
  }).join('');

  // Guardar rows en window para onclick
  window._kpiRows = cards.map(c => c.rows || []);

  // Render VT mini donut
  requestAnimationFrame(() => {
    const vtCanvas = document.getElementById('kpi-vt-donut');
    if (vtCanvas) {
      const vtEjec = cards.find(c => c.isVT);
      if (vtEjec && window.Chart) {
        destroyChart('kpi-vt-donut');
        chartInstances['kpi-vt-donut'] = new Chart(vtCanvas, {
          type: 'doughnut',
          data: {
            datasets: [{
              data: [vtEjec.vtEjec, Math.max(0, vtEjec.vtTotal - vtEjec.vtEjec)],
              backgroundColor: ['#DFFF61','rgba(223,255,97,.12)'],
              borderWidth: 0,
              borderRadius: 4,
            }]
          },
          options: {
            cutout: '70%', responsive: false, animation: { duration: 1000 },
            plugins: { legend: { display: false }, tooltip: { enabled: false } }
          }
        });
      }
    }
  });
}

// ══════════════════════════════════════════════════════════════════
//  BUILD HELPERS
// ══════════════════════════════════════════════════════════════════
function buildDailyMap(data) {
  const map = {};
  data.forEach(r => {
    const d = parseDate(getCol(r, 'FECHA DE SOLICITUD', 'fecha de solicitud'));
    if (!d) return;
    const key = d.toISOString().slice(0, 10);
    map[key] = (map[key] || 0) + 1;
  });
  const sorted = Object.keys(map).sort();
  return {
    labels: sorted.map(k => { const [y,m,d] = k.split('-'); return `${d}/${m}`; }),
    values: sorted.map(k => map[k]),
  };
}

function buildDeptData(data) {
  const map = {};
  data.forEach(r => {
    const dep = getCol(r, 'Departamento', 'DEPARTAMENTO', 'departamento') || 'Sin datos';
    const tip = getCol(r, 'TIPOLOGIA', 'Tipologia', 'tipologia').toUpperCase();
    if (!map[dep]) map[dep] = { principal:0, intermedia:0, lejana:0 };
    if      (tip === 'PRINCIPAL')  map[dep].principal++;
    else if (tip === 'INTERMEDIA') map[dep].intermedia++;
    else if (tip === 'LEJANA')     map[dep].lejana++;
    else                           map[dep].principal++;
  });
  const sorted = Object.entries(map)
    .sort((a, b) => (b[1].principal+b[1].intermedia+b[1].lejana) -
                    (a[1].principal+a[1].intermedia+a[1].lejana));
  return {
    labels:    sorted.map(([k]) => k),
    principal: sorted.map(([,v]) => v.principal),
    intermedia:sorted.map(([,v]) => v.intermedia),
    lejana:    sorted.map(([,v]) => v.lejana),
  };
}

// ══════════════════════════════════════════════════════════════════
//  CHARTS
// ══════════════════════════════════════════════════════════════════
function destroyChart(id) {
  if (chartInstances[id]) { chartInstances[id].destroy(); delete chartInstances[id]; }
}

function renderCharts(k) {
  destroyChart('estados');
  const estadoLabels = Object.keys(k.ec);
  const estadoVals   = Object.values(k.ec);
  const ctxE = document.getElementById('chart-estados');
  if (ctxE) {
    chartInstances['estados'] = new Chart(ctxE, {
      type:'doughnut',
      data:{ labels:estadoLabels, datasets:[{ data:estadoVals,
        backgroundColor:WOMPI_COLORS, borderColor:'#181715', borderWidth:3, hoverOffset:10 }] },
      options:{ responsive:true, maintainAspectRatio:false, cutout:'65%',
        plugins:{ legend:{position:'right',labels:{color:'#FAFAFA',font:{size:11},boxWidth:12,padding:10}},
                  tooltip:CHART_OPTS.tooltip } },
    });
  }

  destroyChart('ans');
  const ctxA = document.getElementById('chart-ans');
  if (ctxA) {
    chartInstances['ans'] = new Chart(ctxA, {
      type:'bar',
      data:{ labels:['% Oportunidad','% Calidad','% VT','% OPLG'],
        datasets:[{ data:[k.pctOport,k.pctCalidad,k.pctVT,k.pctOL],
          backgroundColor:['#B0F2AE','#99D1FC','#DFFF61','#00C87A'],
          borderRadius:8, borderSkipped:false }] },
      options:{ responsive:true, maintainAspectRatio:false, indexAxis:'y',
        scales:{
          x:{ grid:{color:'rgba(255,255,255,.05)'}, ticks:{color:'#7A7674',callback:v=>v+'%'},
              max:100, border:{display:false} },
          y:{ grid:{display:false}, ticks:{color:'#FAFAFA',font:{size:12}}, border:{display:false} },
        },
        plugins:{ legend:{display:false},
                  tooltip:{...CHART_OPTS.tooltip,callbacks:{label:c=>` ${c.parsed.x}%`}} } },
    });
  }

  destroyChart('dias');
  const dias  = buildDailyMap(FILTERED);
  const ctxD  = document.getElementById('chart-dias');
  if (ctxD) {
    chartInstances['dias'] = new Chart(ctxD, {
      type:'line',
      data:{ labels:dias.labels, datasets:[{
        label:'Solicitudes', data:dias.values,
        borderColor:'#B0F2AE', backgroundColor:'rgba(176,242,174,.08)',
        fill:true, tension:.4, pointBackgroundColor:'#B0F2AE',
        pointRadius:3, pointHoverRadius:6, borderWidth:2.5 }] },
      options:{ responsive:true, maintainAspectRatio:false,
        scales:{
          x:{ grid:{color:'rgba(255,255,255,.04)'}, ticks:{color:'#7A7674',maxTicksLimit:12,maxRotation:45}, border:{display:false} },
          y:{ grid:{color:'rgba(255,255,255,.04)'}, ticks:{color:'#7A7674'}, border:{display:false} },
        },
        plugins:{ legend:{display:false}, tooltip:CHART_OPTS.tooltip } },
    });
  }

  destroyChart('tipos');
  const ctxT = document.getElementById('chart-tipos');
  if (ctxT) {
    chartInstances['tipos'] = new Chart(ctxT, {
      type:'bar',
      data:{ labels:['Visita Técnica','Op. Logístico'],
        datasets:[
          { label:'Total',     data:[k.totalVT,k.totalOL], backgroundColor:'rgba(255,255,255,.07)', borderRadius:6, borderSkipped:false },
          { label:'Entregado', data:[k.entVT,k.entOL],     backgroundColor:['#DFFF61','#99D1FC'],   borderRadius:6, borderSkipped:false },
        ] },
      options:{ responsive:true, maintainAspectRatio:false,
        scales:{
          x:{ grid:{display:false}, ticks:{color:'#FAFAFA'}, border:{display:false} },
          y:{ grid:{color:'rgba(255,255,255,.04)'}, ticks:{color:'#7A7674'}, border:{display:false} },
        },
        plugins:{ legend:{labels:{color:'#FAFAFA',font:{size:11},boxWidth:12}}, tooltip:CHART_OPTS.tooltip } },
    });
  }

  destroyChart('dept');
  const deptData = buildDeptData(FILTERED);
  const ctxDep   = document.getElementById('chart-dept');
  if (ctxDep) {
    chartInstances['dept'] = new Chart(ctxDep, {
      type:'bar',
      data:{ labels:deptData.labels.slice(0,15),
        datasets:[
          { label:'Principal',  data:deptData.principal.slice(0,15),  backgroundColor:'#99D1FC', borderRadius:4, borderSkipped:false },
          { label:'Intermedia', data:deptData.intermedia.slice(0,15), backgroundColor:'#B0F2AE', borderRadius:4, borderSkipped:false },
          { label:'Lejana',     data:deptData.lejana.slice(0,15),     backgroundColor:'#DFFF61', borderRadius:4, borderSkipped:false },
        ] },
      options:{ responsive:true, maintainAspectRatio:false,
        scales:{
          x:{ stacked:true, grid:{display:false}, ticks:{color:'#FAFAFA',maxRotation:45}, border:{display:false} },
          y:{ stacked:true, grid:{color:'rgba(255,255,255,.04)'}, ticks:{color:'#7A7674'}, border:{display:false} },
        },
        plugins:{ legend:{labels:{color:'#FAFAFA',font:{size:11},boxWidth:12}}, tooltip:CHART_OPTS.tooltip } },
    });
  }
}

function renderDevCharts() {
  function isDevRow(r) {
    for (const [k, v] of Object.entries(r)) {
      const kUp = k.toUpperCase();
      const vUp = String(v || '').toUpperCase();
      if (vUp.includes('DEVOLUCI') || vUp.includes('DEVOLUCION') ||
          vUp.includes('DEVOLUCIÓN') || vUp.includes('DEVUELTO') ||
          vUp.includes('REMITENTE')) return true;
      if (kUp.includes('DEVOLUCI') || kUp.includes('DEVOLUCION') || kUp.includes('DEVOLUCIÓN')) {
        if (vUp && vUp !== '' && vUp !== '0' && vUp !== 'NAN') return true;
      }
    }
    return false;
  }
  const devRows = FILTERED.filter(isDevRow);
  const byTransp = {}, byCausal = {};
  devRows.forEach(r => {
    const t = getCol(r,'TRANSPORTADORA','Transportadora','transportadora') || 'Sin datos';
    byTransp[t] = (byTransp[t]||0) + 1;
    const c = getCol(r,'CAUSAL INC','causal inc','Causal Inc','NOVEDADES','novedades') || 'Sin datos';
    byCausal[c] = (byCausal[c]||0) + 1;
  });
  destroyChart('dev-transp');
  const ctDT = document.getElementById('chart-dev-transp');
  if (ctDT) chartInstances['dev-transp'] = new Chart(ctDT,{
    type:'bar',
    data:{labels:Object.keys(byTransp),datasets:[{data:Object.values(byTransp),backgroundColor:'#FF5C5C',borderRadius:6,borderSkipped:false}]},
    options:{responsive:true,maintainAspectRatio:false,scales:{x:{grid:{display:false},ticks:{color:'#FAFAFA'},border:{display:false}},y:{grid:{color:'rgba(255,255,255,.04)'},ticks:{color:'#7A7674'},border:{display:false}}},plugins:{legend:{display:false},tooltip:CHART_OPTS.tooltip}},
  });
  destroyChart('dev-motivo');
  const ctDM = document.getElementById('chart-dev-motivo');
  if (ctDM) chartInstances['dev-motivo'] = new Chart(ctDM,{
    type:'doughnut',
    data:{labels:Object.keys(byCausal),datasets:[{data:Object.values(byCausal),backgroundColor:WOMPI_COLORS,borderColor:'#181715',borderWidth:3}]},
    options:{responsive:true,maintainAspectRatio:false,cutout:'60%',plugins:{legend:{position:'right',labels:{color:'#FAFAFA',font:{size:10},boxWidth:10,padding:8}},tooltip:CHART_OPTS.tooltip}},
  });
}

// ══════════════════════════════════════════════════════════════════
//  DEPT TABLE
// ══════════════════════════════════════════════════════════════════
function renderDeptTable() {
  const d = buildDeptData(FILTERED);
  const tbody = document.getElementById('dept-tbody');
  if (!tbody) return;
  let totP=0, totI=0, totL=0;
  // Calcular el máximo total para las mini barras
  const maxTotal = d.labels.length
    ? Math.max(...d.labels.map((_,i) => d.principal[i]+d.intermedia[i]+d.lejana[i]))
    : 1;

  tbody.innerHTML = d.labels.map((dep, i) => {
    const p=d.principal[i], in_=d.intermedia[i], l=d.lejana[i], t=p+in_+l;
    totP+=p; totI+=in_; totL+=l;
    const barW = Math.round((t / maxTotal) * 120); // max 120px
    const pBarW = t ? Math.round((p/t)*barW) : 0;
    const iBarW = t ? Math.round((in_/t)*barW) : 0;
    const lBarW = t ? barW - pBarW - iBarW : 0;
    return `<tr>
      <td>
        <div style="display:flex;flex-direction:column;gap:5px">
          <strong style="color:var(--blanco)">${dep}</strong>
          <div style="display:flex;gap:2px;height:4px;border-radius:2px;overflow:hidden;width:${barW}px;min-width:20px">
            ${pBarW ? `<div style="width:${pBarW}px;background:var(--azul-cielo);border-radius:2px 0 0 2px"></div>` : ''}
            ${iBarW ? `<div style="width:${iBarW}px;background:var(--verde-menta)"></div>` : ''}
            ${lBarW ? `<div style="width:${lBarW}px;background:var(--verde-lima);border-radius:0 2px 2px 0"></div>` : ''}
          </div>
        </div>
      </td>
      <td><span class="dept-val-principal">${p}</span></td>
      <td><span class="dept-val-intermedia">${in_}</span></td>
      <td><span class="dept-val-lejana">${l}</span></td>
      <td><span class="dept-val-total">${t}</span></td>
    </tr>`;
  }).join('') +
    `<tr class="dept-total">
      <td><strong>TOTAL</strong></td>
      <td><span class="dept-val-principal">${totP}</span></td>
      <td><span class="dept-val-intermedia">${totI}</span></td>
      <td><span class="dept-val-lejana">${totL}</span></td>
      <td><span class="dept-val-total">${totP+totI+totL}</span></td>
    </tr>`;
}

// ══════════════════════════════════════════════════════════════════
//  ANS ALERTS
// ══════════════════════════════════════════════════════════════════
function renderANSAlerts(k) {
  const grid = document.getElementById('ans-alerts-grid');
  if (!grid) return;
  const now = new Date();

  // Guías sin cambios = vencidas ANS con novedad (usa findNovedad robusto)
  const today2 = new Date(); today2.setHours(0,0,0,0);
  const guiasEstRows = FILTERED.filter(r => {
    const est = getCol(r,'ESTADO DATAFONO','estado datafono').toUpperCase();
    if (est === 'ENTREGADO' || est === 'CANCELADO') return false;
    const fLim = parseDate(getCol(r,'FECHA LIMITE DE ENTREGA','fecha limite de entrega'));
    if (!fLim) return false;
    const limDay = new Date(fLim); limDay.setHours(0,0,0,0);
    return limDay <= today2 && findNovedad(r) !== null;
  });

  // Intentos fallidos = EN TRÁNSITO con alguna novedad (usa findNovedad robusto)
  const fallidosRows = FILTERED.filter(r => {
    const est = getCol(r,'ESTADO DATAFONO','estado datafono').toUpperCase();
    if (est !== 'EN TRANSITO' && est !== 'EN TRÁNSITO') return false;
    return findNovedad(r) !== null;
  });

  const pctOL = k.total ? Math.round(k.totalOL/k.total*100) : 0;
  const pctVT = k.total ? Math.round(k.totalVT/k.total*100) : 0;

  const alerts = [
    { label:'Vencen Hoy',           value:k.vencenHoy,           type:k.vencenHoy>0?'':'ok',    sub:'Límite hoy sin entregar',               rows:k.vencenHoyRows },
    { label:'Vencidas ANS',         value:k.vencidas,             type:k.vencidas>0?'':'ok',     sub:'Fuera de plazo (VT + OPLG)',             rows:k.vencidasRows },
    { label:'1er Intento',          value:k.pctPrimerIntento+'%', type:'ok',                     sub:`${k.primerIntento} de ${k.entregados}`,  rows:k.primerIntentoRows },
    { label:'Guías sin Cambios',       value:guiasEstRows.length,    type:guiasEstRows.length>0?'warn':'ok',  sub:'Vencidas ANS con novedad',      rows:guiasEstRows },
    { label:'Intentos Fallidos',       value:fallidosRows.length,    type:fallidosRows.length>0?'warn':'ok',  sub:'En tránsito con novedad',        rows:fallidosRows },
    { label:'% Op. Logístico',      value:pctOL+'%',              type:'info',                   sub:`${k.totalOL} vía OPLG`,                 rows:k.olRows },
    { label:'% Visita Técnica',     value:pctVT+'%',              type:'info',                   sub:`${k.totalVT} gestionadas por VT`,       rows:k.vtRows },
    { label:'Devoluciones',         value:k.devueltos,            type:k.devueltos>0?'warn':'ok',sub:'Total devueltos',                        rows:k.devueltosRows },
  ];

  grid.innerHTML = alerts.map(a => `
    <div class="alert-card ${a.type}" onclick="openDrillModal('${a.label}', window._ansAlertRows[${alerts.indexOf(a)}])"
      style="cursor:pointer" title="Ver listado">
      <div class="kpi-drill-hint" style="font-size:9px;color:var(--muted);text-align:right;margin-bottom:2px">Ver ↗</div>
      <div class="alert-label">${a.label}</div>
      <div class="alert-value">${a.value}</div>
      <div class="alert-sub">${a.sub}</div>
    </div>
  `).join('');

  window._ansAlertRows = alerts.map(a => a.rows || []);

  renderBacklog();
  renderStalledGuias();
  renderFallidos();
}

// ══════════════════════════════════════════════════════════════════
//  BACKLOG
// ══════════════════════════════════════════════════════════════════
function renderBacklog() {
  const hrs    = parseInt(document.getElementById('f-backlog-window')?.value || 24);
  const now    = new Date();
  const nowDay = new Date(now); nowDay.setHours(0,0,0,0);
  const cutoff = new Date(nowDay.getTime() + hrs * 3600000);
  const wrap   = document.getElementById('backlog-wrap');
  if (!wrap) return;

  // Incluye guías cuya fecha límite está entre AHORA y el cutoff (futuras próximas a vencer)
  // O que vencen HOY (fecha límite == hoy, sin importar la hora exacta)
  const atRisk = FILTERED.filter(r => {
    const lim = parseDate(getCol(r,'FECHA LIMITE DE ENTREGA','fecha limite de entrega'));
    const est = getCol(r,'ESTADO DATAFONO','estado datafono').toUpperCase();
    if (!lim || est === 'ENTREGADO' || est === 'CANCELADO') return false;
    const limDay = new Date(lim); limDay.setHours(23,59,59,999);
    // Vencen en las próximas Xh (desde hoy 0:00 hasta cutoff)
    return limDay >= nowDay && lim <= cutoff;
  });

  if (!atRisk.length) {
    wrap.innerHTML = '<div class="empty-state"><div class="icon">✅</div><p>Sin registros en riesgo para la ventana seleccionada</p></div>';
    return;
  }
  wrap.innerHTML = `<table>
    <thead><tr><th>Comercio</th><th>ID Sitio</th><th>Guía</th><th>Fecha Límite</th><th>Transportadora</th><th>Estado</th><th>Tipo</th><th>Riesgo</th></tr></thead>
    <tbody>${atRisk.map(r => {
      const lim  = parseDate(getCol(r,'FECHA LIMITE DE ENTREGA','fecha limite de entrega'));
      const limDay = lim ? new Date(lim) : null;
      if (limDay) limDay.setHours(23,59,59,999);
      const hLeft= limDay ? Math.round((limDay - now)/3600000) : 0;
      const urgColor = hLeft <= 0 ? 'var(--danger)' : hLeft <= 24 ? 'var(--warning)' : 'var(--azul-cielo)';
      const urgLabel = hLeft <= 0 ? 'VENCE HOY' : hLeft <= 24 ? `${hLeft}h` : `${Math.ceil(hLeft/24)}d`;
      return `<tr>
        <td>${getCol(r,'Nombre del comercio','nombre del comercio','NOMBRE DEL COMERCIO')||'—'}</td>
        <td style="font-family:'JetBrains Mono',monospace;font-size:11px">${getCol(r,'ID Comercio','id comercio')||'—'}</td>
        <td style="font-family:'JetBrains Mono',monospace;font-size:11px">${getCol(r,'NÚMERO DE GUIA','NUMERO DE GUIA','numero de guia')||'—'}</td>
        <td>${lim?lim.toLocaleDateString('es-CO'):'—'}</td>
        <td>${getCol(r,'TRANSPORTADORA','Transportadora','transportadora')||'—'}</td>
        <td><span class="status-pill ${statusClass(getCol(r,'ESTADO DATAFONO','estado datafono'))}">${getCol(r,'ESTADO DATAFONO','estado datafono')||'—'}</span></td>
        <td style="font-size:10px;color:var(--muted)">${getCol(r,'TIPO DE SOLICITUD FACTURACIÓN','TIPO DE SOLICITUD FACTURACION','tipo de solicitud facturacion')||'—'}</td>
        <td><span class="risk-badge" style="background:rgba(255,255,255,.07);color:${urgColor};border:1px solid ${urgColor};border-radius:6px;padding:3px 8px;font-size:11px;font-weight:700;font-family:'JetBrains Mono',monospace">${urgLabel}</span></td>
      </tr>`;
    }).join('')}</tbody>
  </table>`;
}

// ══════════════════════════════════════════════════════════════════
//  GUÍAS ESTANCADAS
// ══════════════════════════════════════════════════════════════════
function renderStalledGuias() {
  const wrap = document.getElementById('guias-estancadas-wrap');
  if (!wrap) return;
  const now = new Date(); now.setHours(0,0,0,0);

  // Guías SIN CAMBIOS = vencidas ANS (fecha límite pasada, no entregadas/canceladas)
  // que tienen algún valor en columna NOVEDADES / NOVEDAD / CAUSAL (cualquier variante).
  // Días sin cambios = días desde FECHA LIMITE DE ENTREGA (no desde solicitud).
  const stalled = FILTERED.filter(r => {
    const est = getCol(r,'ESTADO DATAFONO','estado datafono').toUpperCase();
    if (est === 'ENTREGADO' || est === 'CANCELADO') return false;
    const fLim = getFechaLimite(r);
    if (!fLim) return false;
    const limDay = new Date(fLim); limDay.setHours(0,0,0,0);
    if (limDay > now) return false;            // aún no vencida
    return findNovedad(r) !== null;
  });

  if (!stalled.length) {
    wrap.innerHTML = '<div class="empty-state"><div class="icon">✅</div><p>Sin guías sin cambios</p></div>';
    return;
  }
  wrap.innerHTML = `<table>
    <thead><tr>
      <th>Comercio</th><th>Guía</th><th>Fecha Límite</th>
      <th style="color:var(--warning)">Días sin cambios</th>
      <th>Novedad</th>
      <th>Transportadora</th><th>Estado</th><th>Tipo</th>
    </tr></thead>
    <tbody>${stalled.sort((a,b)=>{
      const da = getFechaLimite(a) || new Date(0);
      const db = getFechaLimite(b) || new Date(0);
      return da - db;
    }).map(r=>{
      const fLim  = getFechaLimite(r);
      const limDay = fLim ? new Date(fLim) : null;
      if (limDay) limDay.setHours(0,0,0,0);
      const dias = limDay ? diffDays(limDay, now) : null;
      const cls  = dias===null?'ok':dias>=7?'crit':dias>=3?'warn':'ok';
      const label = dias === null ? '—'
        : dias === 1 ? '1 día sin cambios'
        : `${dias} días sin cambios`;
      const novedad = (findNovedad(r) || {val:'—'}).val;
      return `<tr>
        <td>${getCol(r,'Nombre del comercio','nombre del comercio','NOMBRE DEL COMERCIO')||'—'}</td>
        <td style="font-family:'JetBrains Mono',monospace;font-size:11px">${getCol(r,'NÚMERO DE GUIA','NUMERO DE GUIA','numero de guia')||'—'}</td>
        <td style="color:var(--muted)">${fLim?fLim.toLocaleDateString('es-CO'):'—'}</td>
        <td><span class="days-stalled ${cls}">${label}</span></td>
        <td style="color:var(--warning);font-size:11px">${novedad}</td>
        <td>${getCol(r,'TRANSPORTADORA','Transportadora','transportadora')||'—'}</td>
        <td><span class="status-pill ${statusClass(getCol(r,'ESTADO DATAFONO','estado datafono'))}">${getCol(r,'ESTADO DATAFONO','estado datafono')||'—'}</span></td>
        <td style="font-size:10px;color:var(--muted)">${getCol(r,'TIPO DE SOLICITUD FACTURACIÓN','TIPO DE SOLICITUD FACTURACION','tipo de solicitud facturacion')||'—'}</td>
      </tr>`;
    }).join('')}</tbody>
  </table>`;
}

// ══════════════════════════════════════════════════════════════════
//  FALLIDOS
// ══════════════════════════════════════════════════════════════════
function renderFallidos() {
  const wrap = document.getElementById('fallidos-wrap');
  if (!wrap) return;

  // Intento fallido = EN TRÁNSITO con alguna novedad real registrada.
  // Usamos findNovedad() que busca robustamente por nombre/valor de columna,
  // EXCLUYENDO columnas de estado/fecha que generan falsos positivos.
  const fallidos = FILTERED.filter(r => {
    const est = getCol(r,'ESTADO DATAFONO','estado datafono').toUpperCase();
    if (est !== 'EN TRANSITO' && est !== 'EN TRÁNSITO') return false;
    return findNovedad(r) !== null;
  });

  if (!fallidos.length) {
    wrap.innerHTML = '<div class="empty-state"><div class="icon">✅</div><p>Sin intentos fallidos</p></div>';
    return;
  }
  wrap.innerHTML = `<table>
    <thead><tr>
      <th>Comercio</th><th>Guía</th>
      <th style="color:var(--danger)">Novedad</th>
      <th>Columna</th>
      <th>Transportadora</th><th>Estado</th><th>Tipo</th>
    </tr></thead>
    <tbody>${fallidos.map(r=>{
      const info = findNovedad(r) || { col: '—', val: '—' };
      return `<tr>
        <td>${getCol(r,'Nombre del comercio','nombre del comercio','NOMBRE DEL COMERCIO')||'—'}</td>
        <td style="font-family:'JetBrains Mono',monospace;font-size:11px">${getCol(r,'NÚMERO DE GUIA','NUMERO DE GUIA','numero de guia')||'—'}</td>
        <td style="color:var(--danger);font-weight:600">${info.val}</td>
        <td style="font-size:10px;color:var(--muted)">${info.col}</td>
        <td>${getCol(r,'TRANSPORTADORA','Transportadora','transportadora')||'—'}</td>
        <td><span class="status-pill ${statusClass(getCol(r,'ESTADO DATAFONO','estado datafono'))}">${getCol(r,'ESTADO DATAFONO','estado datafono')||'—'}</span></td>
        <td style="font-size:10px;color:var(--muted)">${getCol(r,'TIPO DE SOLICITUD FACTURACIÓN','TIPO DE SOLICITUD FACTURACION','tipo de solicitud facturacion')||'—'}</td>
      </tr>`;
    }).join('')}</tbody>
  </table>`;
}

// ══════════════════════════════════════════════════════════════════
//  TABLA PRINCIPAL
// ══════════════════════════════════════════════════════════════════
const TABLE_COLS = [
  { label:'Comercio',        fn:r=>getCol(r,'Nombre del comercio','nombre del comercio','NOMBRE DEL COMERCIO') },
  { label:'ID Sitio',        fn:r=>getCol(r,'ID Comercio','id comercio','Id Comercio') },
  { label:'Material',        fn:r=>getCol(r,'REFERENCIA DEL DATAFONO','REFERENCIA DEL DATAFONOS','referencia del datafono') },
  { label:'Num. Serie',      fn:r=>getCol(r,'SERIAL DATAFÓNOS','SERIAL DATAFONOS','serial datafonos','Serial Datafono') },
  { label:'Fecha Solicitud', fn:r=>{ const d=parseDate(getCol(r,'FECHA DE SOLICITUD','fecha de solicitud')); return d?d.toLocaleDateString('es-CO'):'—'; } },
  { label:'Fecha Límite',    fn:r=>{ const d=parseDate(getCol(r,'FECHA LIMITE DE ENTREGA','fecha limite de entrega')); return d?d.toLocaleDateString('es-CO'):'—'; } },
  { label:'Fecha Entrega',   fn:r=>{ const d=parseDate(getCol(r,'FECHA ENTREGA AL COMERCIO','fecha entrega al comercio')); return d?d.toLocaleDateString('es-CO'):'—'; } },
  { label:'Tipo Envío',      fn:r=>getCol(r,'TIPO DE SOLICITUD','tipo de solicitud') },
  { label:'Transportadora',  fn:r=>getCol(r,'TRANSPORTADORA','Transportadora','transportadora') },
  { label:'Guía',            fn:r=>getCol(r,'NÚMERO DE GUIA','NUMERO DE GUIA','numero de guia') },
  { label:'Estado',          fn:r=>getCol(r,'ESTADO DATAFONO','estado datafono'), isStatus:true },
  { label:'Estado Guía',     fn:r=>getCol(r,'ESTADO GUIA','estado guia'), isStatus:true },
  { label:'Cumple ANS',      fn:r=>getCol(r,'CUMPLE ANS','cumple ans') },
  { label:'Departamento',    fn:r=>getCol(r,'Departamento','DEPARTAMENTO','departamento') },
  { label:'Ciudad',          fn:r=>getCol(r,'Ciudad','CIUDAD','ciudad') },
];

function statusClass(v) {
  const s = (v||'').toUpperCase();
  if (s === 'ENTREGADO')                         return 'status-entregado';
  if (s.includes('TRANSITO')||s.includes('TRÁNSITO')) return 'status-transito';
  if (s.includes('ALISTAMIENTO'))                return 'status-alistamiento';
  if (s.includes('DEVOLU')||s.includes('REMIT')) return 'status-devolucion';
  if (s === 'CANCELADO')                         return 'status-cancelado';
  return 'status-default';
}

function renderMainTable() {
  filteredForTable = FILTERED.filter(r => {
    if (!tableSearchTerm) return true;
    const s = tableSearchTerm.toLowerCase();
    return TABLE_COLS.some(c => (c.fn(r)||'').toLowerCase().includes(s));
  });
  const total = filteredForTable.length;
  const pages = Math.max(1, Math.ceil(total/TABLE_PAGE_SIZE));
  tablePage   = Math.min(tablePage, pages);
  const start = (tablePage-1)*TABLE_PAGE_SIZE;
  const slice = filteredForTable.slice(start, start+TABLE_PAGE_SIZE);
  const wrap  = document.getElementById('main-table-wrap');
  if (!wrap) return;
  wrap.innerHTML = `<table>
    <thead><tr>${TABLE_COLS.map((c,i)=>`<th onclick="sortTable(${i})" class="${sortCol===i?'sorted':''}">${c.label}<span class="sort-icon">${sortCol===i?(sortDir>0?'▲':'▼'):'⬍'}</span></th>`).join('')}</tr></thead>
    <tbody>${slice.length ? slice.map(r=>`<tr>${TABLE_COLS.map(c=>{const v=c.fn(r)||'—';return c.isStatus?`<td><span class="status-pill ${statusClass(v)}">${v}</span></td>`:`<td>${v}</td>`;}).join('')}</tr>`).join('') : `<tr><td colspan="${TABLE_COLS.length}" style="text-align:center;padding:40px;color:var(--muted)">Sin resultados</td></tr>`}</tbody>
  </table>`;
  const tc = document.getElementById('table-count');
  if (tc) tc.textContent = `${total} registros`;
  renderPagination(pages);
}

function tableSearch(v) { tableSearchTerm = v; tablePage = 1; renderMainTable(); }

function sortTable(col) {
  if (sortCol===col) sortDir*=-1; else { sortCol=col; sortDir=1; }
  const fn = TABLE_COLS[col].fn;
  FILTERED.sort((a,b)=>{ const va=fn(a)||'',vb=fn(b)||''; return va.localeCompare(vb,'es',{numeric:true})*sortDir; });
  tablePage = 1;
  renderMainTable();
}

function renderPagination(pages) {
  const pg = document.getElementById('table-pagination');
  if (!pg) return;
  let html = `<button class="page-btn" onclick="goPage(${tablePage-1})" ${tablePage===1?'disabled':''}>‹</button>`;
  const range = [];
  for (let i=1; i<=pages; i++) {
    if (i===1||i===pages||Math.abs(i-tablePage)<=2) range.push(i);
    else if (range[range.length-1]!=='…') range.push('…');
  }
  range.forEach(p => {
    if (p==='…') html+=`<span style="padding:4px 6px;color:var(--muted);display:inline-flex;align-items:center">…</span>`;
    else html+=`<button class="page-btn ${p===tablePage?'active':''}" onclick="goPage(${p})">${p}</button>`;
  });
  html += `<button class="page-btn" onclick="goPage(${tablePage+1})" ${tablePage===pages?'disabled':''}>›</button>`;
  pg.innerHTML = html;
}

function goPage(p) {
  const pages = Math.ceil(filteredForTable.length/TABLE_PAGE_SIZE);
  if (p>=1 && p<=pages) { tablePage=p; renderMainTable(); }
}

// ══════════════════════════════════════════════════════════════════
//  EXPORT
// ══════════════════════════════════════════════════════════════════
function exportMainExcel() {
  const data = filteredForTable.map(r=>{ const obj={}; TABLE_COLS.forEach(c=>{obj[c.label]=c.fn(r)||'';}); return obj; });
  exportToExcel(data, 'Tracking_VP_Wompi_Detalle');
}
function exportDeptExcel() {
  const d = buildDeptData(FILTERED);
  const data = d.labels.map((dep,i)=>({Departamento:dep,Principal:d.principal[i],Intermedia:d.intermedia[i],Lejana:d.lejana[i],Total:d.principal[i]+d.intermedia[i]+d.lejana[i]}));
  exportToExcel(data, 'Tracking_VP_Departamentos');
}
function exportBacklogExcel() {
  const hrs    = parseInt(document.getElementById('f-backlog-window')?.value||24);
  const now    = new Date();
  const nowDay = new Date(now); nowDay.setHours(0,0,0,0);
  const cutoff = new Date(nowDay.getTime()+hrs*3600000);
  const data   = FILTERED.filter(r=>{
    const lim = parseDate(getCol(r,'FECHA LIMITE DE ENTREGA','fecha limite de entrega'));
    const est = getCol(r,'ESTADO DATAFONO','estado datafono').toUpperCase();
    if (!lim || est==='ENTREGADO' || est==='CANCELADO') return false;
    const limDay = new Date(lim); limDay.setHours(23,59,59,999);
    return limDay >= nowDay && lim <= cutoff;
  }).map(r=>({
    Comercio:       getCol(r,'Nombre del comercio','nombre del comercio','NOMBRE DEL COMERCIO'),
    'ID Sitio':     getCol(r,'ID Comercio','id comercio'),
    Guía:           getCol(r,'NÚMERO DE GUIA','NUMERO DE GUIA','numero de guia'),
    'Fecha Límite': getCol(r,'FECHA LIMITE DE ENTREGA','fecha limite de entrega'),
    Transportadora: getCol(r,'TRANSPORTADORA','Transportadora'),
    Estado:         getCol(r,'ESTADO DATAFONO','estado datafono'),
    Tipo:           getCol(r,'TIPO DE SOLICITUD FACTURACIÓN','TIPO DE SOLICITUD FACTURACION','tipo de solicitud facturacion'),
  }));
  exportToExcel(data, `Backlog_Riesgo_${hrs}h`);
}

function exportToExcel(data, filename) {
  if (!data.length) { alert('Sin datos para exportar.'); return; }
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(data);
  const range = XLSX.utils.decode_range(ws['!ref']);
  for (let C=range.s.c; C<=range.e.c; C++) {
    const ref = XLSX.utils.encode_cell({r:0,c:C});
    if (!ws[ref]) continue;
    ws[ref].s = { fill:{patternType:'solid',fgColor:{rgb:'2C2A29'}}, font:{color:{rgb:'B0F2AE'},bold:true}, alignment:{horizontal:'center'} };
  }
  ws['!cols'] = Object.keys(data[0]||{}).map(k=>({wch:Math.max(k.length,...data.slice(0,50).map(r=>String(r[k]||'').length))}));
  XLSX.utils.book_append_sheet(wb, ws, 'Datos');

  // Hoja KPIs
  const k = computeKPIs(FILTERED);
  const summaryData = [
    {KPI:'Total Solicitados',     Valor:k.total},
    {KPI:'Alistados',             Valor:`${k.n_alistados} (${k.pctNAlistados}%)`},
    {KPI:'Entregados',            Valor:`${k.entregados} (${k.pctEntregado}%)`},
    {KPI:'En Tránsito',           Valor:k.en_transito},
    {KPI:'En Alistamiento',       Valor:k.en_alistamiento},
    {KPI:'Devueltos',             Valor:k.devueltos},
    {KPI:'Visita Técnica',        Valor:`${k.entVT} ejec / ${k.programados_vt} prog / ${k.totalVT} total (${k.pctVT}%)`},
    {KPI:'Op. Logístico',         Valor:`${k.entOL}/${k.totalOL} (${k.pctOL}%)`},
    {KPI:'% Oportunidad ANS',     Valor:`${k.pctOport}%`},
    {KPI:'% Calidad',             Valor:`${k.pctCalidad}%`},
    {KPI:'Vencen Hoy',            Valor:k.vencenHoy},
    {KPI:'Vencidas ANS',          Valor:k.vencidas},
    {KPI:'Generado',              Valor:new Date().toLocaleString('es-CO')},
  ];
  const ws2 = XLSX.utils.json_to_sheet(summaryData);
  ws2['!cols'] = [{wch:28},{wch:32}];
  XLSX.utils.book_append_sheet(wb, ws2, 'KPIs');
  XLSX.writeFile(wb, `${filename}_${new Date().toISOString().slice(0,10)}.xlsx`);
}

function exportPDF() { window.print(); }

// ══════════════════════════════════════════════════════════════════
//  TABS
// ══════════════════════════════════════════════════════════════════
function showTab(tab) {
  ['tracking','detalle','tabla','rollos'].forEach(t => {
    const panel = document.getElementById('panel-'+t);
    const btn   = document.getElementById('tab-'+t);
    if (panel) panel.style.display = t===tab ? 'block' : 'none';
    if (btn)   btn.classList.toggle('active', t===tab);
  });
  if (tab==='detalle') { renderDevCharts(); renderBacklog(); renderStalledGuias(); renderFallidos(); }
  if (tab==='tabla')   renderMainTable();
  if (tab==='rollos')  renderRollosTab();
}

// ══════════════════════════════════════════════════════════════════
//  INIT
// ══════════════════════════════════════════════════════════════════
function initDashboard() { loadData(); loadRollosData(); }

// ══════════════════════════════════════════════════════════════════
//  DEMO DATA
// ══════════════════════════════════════════════════════════════════
function getDemoData() {
  const estados = ['ENTREGADO','ENTREGADO','ENTREGADO','EN TRANSITO','EN TRANSITO','EN ALISTAMIENTO','PROGRAMADO','DEVOLUCION'];
  const deptos  = ['ANTIOQUIA','CUNDINAMARCA','VALLE DEL CAUCA','SANTANDER','ATLANTICO','BOLIVAR','CORDOBA','NARIÑO','TOLIMA','HUILA'];
  // Tipos EXACTOS igual a vp.py para que VT y OPLG funcionen
  const tipos   = [VT_EXACT, OPLG_EXACT, 'ENVIO DATAFONO - VENTA'];
  const transps = ['COORDINADORA','SERVIENTREGA','DEPRISA','ENVIA','TCC'];
  const tipols  = ['PRINCIPAL','INTERMEDIA','LEJANA'];
  const rows    = [];
  for (let i = 0; i < 320; i++) {
    // Fechas distribuidas desde octubre 2025 hasta abril 2026 para datos históricos reales
    const monthOffset = Math.floor(i / 53);  // ~53 registros por mes, 6 meses
    const startMonth  = new Date(2025, 9, 1); // Octubre 2025
    const fSol = new Date(startMonth.getFullYear(), startMonth.getMonth() + monthOffset, 1+(i%28));
    const fLim = new Date(fSol.getTime() + 7*86400000);
    const est  = estados[i%estados.length];
    const fEnt = est==='ENTREGADO' ? new Date(fSol.getTime()+(3+Math.floor(Math.random()*4))*86400000) : null;
    rows.push({
      'ID Comercio':                    String(100000+i),
      'Nombre del comercio':            `COMERCIO DEMO ${i+1}`,
      'Departamento':                   deptos[i%deptos.length],
      'Ciudad':                         deptos[i%deptos.length]+' CAPITAL',
      'REFERENCIA DEL DATAFONO':        i%2===0?'EX6000':'LANE 3000',
      'SERIAL DATAFÓNOS':               `248KKU${600000+i}`,
      'TIPO DE SOLICITUD':              tipos[i%tipos.length],
      'TIPO DE SOLICITUD FACTURACIÓN':  tipos[i%tipos.length],
      'FECHA DE SOLICITUD':             fSol.toLocaleDateString('es-CO'),
      'FECHA LIMITE DE ENTREGA':        fLim.toLocaleDateString('es-CO'),
      'TRANSPORTADORA':                 transps[i%transps.length],
      'NÚMERO DE GUIA':                 `FO-26-${200000+i}`,
      'ESTADO DATAFONO':                est,
      'FECHA ENTREGA AL COMERCIO':      fEnt?fEnt.toLocaleDateString('es-CO'):'',
      'FECHA DE ENTREGA':               fEnt?fEnt.toLocaleDateString('es-CO'):'',
      'CUMPLE ANS':                     fEnt&&fEnt<=fLim?'SI':'NO',
      'ESTADO GUIA':                    est==='ENTREGADO'?'ENTREGADO':'EN TRANSITO',
      'TIPOLOGIA':                      tipols[i%tipols.length],
      'NOVEDADES':                      i%15===0?'CLIENTE AUSENTE':i%20===0?'DIRECCION INVALIDA':'',
    });
  }
  return rows;
}
// ══════════════════════════════════════════════════════════════════
//  ROLLOS WOMPI — módulo independiente
//  Carga data_rollos.json.gz (gzip → JSON), sin tocar nada de lo anterior
// ══════════════════════════════════════════════════════════════════

let ROLLOS_RAW      = null;   // payload completo del .json.gz
let ROLLOS_DETALLE  = [];     // array de detalle filtrado
let ROLLOS_COMERCIO = [];     // array de comercio filtrado
let rollosDetallePage  = 1;
let rollosComercioPage = 1;
const ROLLOS_PAGE_SIZE = 50;

// ── Carga y descompresión ─────────────────────────────────────────
async function loadRollosData() {
  try {
    const res = await fetch(`data_rollos.json.gz?t=${Date.now()}`);
    if (!res.ok) throw new Error('no file');
    const buf  = await res.arrayBuffer();
    // Descomprimir gzip usando DecompressionStream (disponible en browsers modernos)
    const ds   = new DecompressionStream('gzip');
    const writer = ds.writable.getWriter();
    writer.write(new Uint8Array(buf));
    writer.close();
    const out  = await new Response(ds.readable).arrayBuffer();
    const text = new TextDecoder().decode(out);
    ROLLOS_RAW = JSON.parse(text);
    console.log('[Rollos] Cargado OK — detalle:', ROLLOS_RAW.detalle?.length, 'filas');
    initRollosFilters();
    // si el tab rollos ya está activo, renderizarlo
    if (document.getElementById('tab-rollos')?.classList.contains('active')) renderRollosTab();
  } catch(e) {
    console.warn('[Rollos] No se pudo cargar data_rollos.json.gz:', e.message);
  }
}

// ── Inicializar select-filters con valores únicos ─────────────────
function initRollosFilters() {
  if (!ROLLOS_RAW) return;
  const detalle  = ROLLOS_RAW.detalle  || [];
  const comercio = ROLLOS_RAW.comercio || [];

  // Estado
  const estados = [...new Set(detalle.map(r => r.estado).filter(Boolean))].sort();
  const fEst = document.getElementById('rf-estado');
  if (fEst) fEst.innerHTML = '<option value="">Todos</option>' + estados.map(e=>`<option value="${e}">${e}</option>`).join('');

  // Año
  const anios = [...new Set(detalle.map(r => r.anio).filter(Boolean))].sort().reverse();
  const fAnio = document.getElementById('rf-anio');
  if (fAnio) fAnio.innerHTML = '<option value="">Todos</option>' + anios.map(a=>`<option value="${a}">${a}</option>`).join('');

  // Mes
  const meses = [...new Set(detalle.map(r => r.mes).filter(Boolean))].sort();
  const fMes = document.getElementById('rf-mes');
  if (fMes) fMes.innerHTML = '<option value="">Todos</option>' + meses.map(m=>`<option value="${m}">${String(m).padStart(2,'0')}</option>`).join('');

  // Comercio estado
  const comEst = [...new Set(comercio.map(r => r.estado).filter(Boolean))].sort();
  const fComEst = document.getElementById('rf-com-estado');
  if (fComEst) fComEst.innerHTML = '<option value="">Todos</option>' + comEst.map(e=>`<option value="${e}">${e}</option>`).join('');

  // Comercio tipo
  const comTipo = [...new Set(comercio.map(r => r.tipo_envio).filter(Boolean))].sort();
  const fComTipo = document.getElementById('rf-com-tipo');
  if (fComTipo) fComTipo.innerHTML = '<option value="">Todos</option>' + comTipo.map(t=>`<option value="${t}">${t}</option>`).join('');

  // Comercio mes (basado en detalle)
  const fComMes = document.getElementById('rf-com-mes');
  if (fComMes) fComMes.innerHTML = '<option value="">Todos</option>' + meses.map(m=>`<option value="${m}">${String(m).padStart(2,'0')}</option>`).join('');

  // Aplicar datos sin filtros
  ROLLOS_DETALLE  = detalle.slice();
  ROLLOS_COMERCIO = comercio.slice();
}

// ── Render principal ──────────────────────────────────────────────
function renderRollosTab() {
  if (!ROLLOS_RAW) {
    document.getElementById('rollos-kpi-grid').innerHTML = '<div class="empty-state"><div class="icon">⚠️</div><p>data_rollos.json.gz no disponible</p></div>';
    return;
  }
  renderRollosKPIs();
  renderRollosDetalleTable();
  renderRollosComercioTable();
  renderRollosRefTable();
  renderRollosDeptChart();
}

// ── KPIs ──────────────────────────────────────────────────────────
function renderRollosKPIs() {
  const k = ROLLOS_RAW.kpis || {};
  const grid = document.getElementById('rollos-kpi-grid');
  if (!grid) return;

  const cards = [
    { label:'Rollos Alistamiento',   value: k.rollos_alistamiento ?? '—', icon:'🔧', color:'lime',
      sub: `${k.tareas_alistamiento ?? 0} tareas` },
    { label:'Rollos en Tránsito',    value: k.rollos_transito ?? '—',     icon:'🚚', color:'blue',
      sub: `${k.tareas_transito ?? 0} tareas` },
    { label:'Rollos Entregados',     value: k.rollos_entregados ?? '—',   icon:'✅', color:'selva',
      sub: `${k.tareas_entregados ?? 0} tareas` },
    { label:'Rollos Devueltos',      value: k.rollos_devueltos ?? '—',    icon:'↩️', color:'danger',
      sub: `${k.tareas_devueltos ?? 0} tareas` },
    { label:'Total Solicitados',     value: k.total_rollos_solicitados ?? '—', icon:'📦', color:'green',
      sub: `${k.total_tareas_solicitadas ?? 0} tareas totales` },
    { label:'% Cumplimiento Entrega',value: (k.pct_sla ?? 0) + '%',       icon:'🎯', color:'green',
      sub: `${k.sla_cumple ?? 0} / ${k.sla_total ?? 0} en plazo`, pct: k.pct_sla ?? 0 },
    { label:'% Oportunidad',         value: (k.pct_sla ?? 0) + '%',       icon:'⏱️', color:'blue',
      sub: 'Entregas en plazo SLA', pct: k.pct_sla ?? 0 },
    { label:'% Calidad',             value: (k.pct_calidad ?? 0) + '%',   icon:'💎', color:'blue',
      sub: `${k.calidad_exitoso ?? 0} exitosos / ${k.calidad_total ?? 0}`, pct: k.pct_calidad ?? 0 },
  ];

  grid.innerHTML = cards.map((c,i) => `
    <div class="kpi-card ${c.color} fade-up" style="animation-delay:${i*.04}s">
      <span class="kpi-icon">${c.icon}</span>
      <div class="kpi-label">${c.label}</div>
      <div class="kpi-value">${c.value}</div>
      ${c.pct !== undefined ? `
        <div class="progress-wrap">
          <div class="progress-track"><div class="progress-fill ${c.color}" style="width:${Math.min(c.pct,100)}%"></div></div>
        </div>` : ''}
      <div class="kpi-sub">${c.sub}</div>
    </div>`).join('');
}

// ── Helpers tabla ─────────────────────────────────────────────────
function statusPill(val) {
  const v = (val||'').toUpperCase();
  let cls = 'status-default';
  if (v.includes('ENTREGADO'))      cls = 'status-entregado';
  else if (v.includes('TRANSITO'))  cls = 'status-transito';
  else if (v.includes('ALISTAM'))   cls = 'status-alistamiento';
  else if (v.includes('DEVOLUC'))   cls = 'status-devolucion';
  else if (v.includes('CANCELADO')) cls = 'status-cancelado';
  return `<span class="status-pill ${cls}">${val||'—'}</span>`;
}

function diasBadge(d) {
  const n = parseFloat(d);
  if (isNaN(n) || d === '') return '—';
  const cls = n < 0 ? 'crit' : n < 7 ? 'warn' : 'ok';
  return `<span class="days-stalled ${cls}">${Math.round(n)}</span>`;
}

function mkPagination(containerId, page, pages, setPageFn) {
  const pg = document.getElementById(containerId);
  if (!pg) return;
  if (pages <= 1) { pg.innerHTML = ''; return; }
  let html = `<button class="page-btn" onclick="${setPageFn}(${page-1})" ${page===1?'disabled':''}>‹</button>`;
  const range = [];
  for (let i=1; i<=pages; i++) {
    if (i===1||i===pages||Math.abs(i-page)<=1) range.push(i);
    else if (range[range.length-1]!=='…') range.push('…');
  }
  range.forEach(p => {
    if (p==='…') html+=`<span style="padding:4px 6px;color:var(--muted);display:inline-flex;align-items:center">…</span>`;
    else html+=`<button class="page-btn ${p===page?'active':''}" onclick="${setPageFn}(${p})">${p}</button>`;
  });
  html += `<button class="page-btn" onclick="${setPageFn}(${page+1})" ${page===pages?'disabled':''}>›</button>`;
  pg.innerHTML = html;
}

// ── Tabla Detalles ────────────────────────────────────────────────
const DETALLE_COLS = [
  { label:'Cod. Sitio',           fn: r => r.cod_sitio || '—' },
  { label:'F. Plan Inicio',       fn: r => r.fecha_plan_inicio || '—' },
  { label:'F. Plan Entrega',      fn: r => r.fecha_plan_fin || '—' },
  { label:'F. Entrega',           fn: r => r.fecha_entrega || '—' },
  { label:'Pendientes',           fn: r => (r.estado||'').toUpperCase().includes('PENDIENTE') ? '●' : '' },
  { label:'Cantidad',             fn: r => r.cantidad ?? '—' },
  { label:'Cod. Tarea',           fn: r => r.codigo_tarea || '—' },
  { label:'Cod. Ubicación',       fn: r => r.cod_ubicacion || '—' },
  { label:'Guía',                 fn: r => r.guia || '—' },
  { label:'FO',                   fn: r => r.FO || '—' },
  { label:'Estado',               fn: r => r.estado || '—', isStatus: true },
  { label:'Estado Transportadora',fn: r => r.estado_transportadora || '—' },
  { label:'Proyecto',             fn: r => r.proyecto || '—' },
  { label:'Días Inv. Restantes',  fn: r => r.dias_inventario_restantes ?? '—', isDias: true },
];

function renderRollosDetalleTable() {
  const wrap  = document.getElementById('rollos-detalle-wrap');
  const count = document.getElementById('rollos-detalle-count');
  if (!wrap) return;
  const data  = ROLLOS_DETALLE;
  const pages = Math.ceil(data.length / ROLLOS_PAGE_SIZE);
  const slice = data.slice((rollosDetallePage-1)*ROLLOS_PAGE_SIZE, rollosDetallePage*ROLLOS_PAGE_SIZE);

  if (!data.length) {
    wrap.innerHTML = '<div class="empty-state"><div class="icon">📭</div><p>Sin registros</p></div>';
    if (count) count.textContent = '0 registros';
    return;
  }

  wrap.innerHTML = `<table><thead><tr>
    ${DETALLE_COLS.map(c=>`<th>${c.label}</th>`).join('')}
  </tr></thead><tbody>
    ${slice.map(r => `<tr>${DETALLE_COLS.map(c => {
      const v = c.fn(r);
      if (c.isStatus) return `<td>${statusPill(v)}</td>`;
      if (c.isDias)   return `<td>${diasBadge(v)}</td>`;
      return `<td>${v}</td>`;
    }).join('')}</tr>`).join('')}
  </tbody></table>`;

  if (count) count.textContent = `${data.length} registros`;
  mkPagination('rollos-detalle-pagination', rollosDetallePage, pages, 'goRollosDetallePage');
}

function goRollosDetallePage(p) {
  const pages = Math.ceil(ROLLOS_DETALLE.length / ROLLOS_PAGE_SIZE);
  if (p >= 1 && p <= pages) { rollosDetallePage = p; renderRollosDetalleTable(); }
}

// ── Tabla Comercio ────────────────────────────────────────────────
const COMERCIO_COLS = [
  { label:'Cod. Comercio', fn: r => r.cod_comercio || '—' },
  { label:'Cantidad',      fn: r => r.cantidad ?? '—' },
  { label:'Nombre Sitio',  fn: r => r.nombre_sitio || '—' },
  { label:'Dirección',     fn: r => r.direccion || '—' },
  { label:'Ciudad',        fn: r => r.ciudad || '—' },
  { label:'Departamento',  fn: r => r.departamento || '—' },
  { label:'Estado',        fn: r => r.estado || '—', isStatus: true },
  { label:'Tipo Envío',    fn: r => r.tipo_envio || '—' },
];

function renderRollosComercioTable() {
  const wrap  = document.getElementById('rollos-comercio-wrap');
  const count = document.getElementById('rollos-comercio-count');
  if (!wrap) return;
  const data  = ROLLOS_COMERCIO;
  const pages = Math.ceil(data.length / ROLLOS_PAGE_SIZE);
  const slice = data.slice((rollosComercioPage-1)*ROLLOS_PAGE_SIZE, rollosComercioPage*ROLLOS_PAGE_SIZE);

  if (!data.length) {
    wrap.innerHTML = '<div class="empty-state"><div class="icon">📭</div><p>Sin registros</p></div>';
    if (count) count.textContent = '0 registros';
    return;
  }

  wrap.innerHTML = `<table><thead><tr>
    ${COMERCIO_COLS.map(c=>`<th>${c.label}</th>`).join('')}
  </tr></thead><tbody>
    ${slice.map(r => `<tr>${COMERCIO_COLS.map(c => {
      const v = c.fn(r);
      return c.isStatus ? `<td>${statusPill(v)}</td>` : `<td>${v}</td>`;
    }).join('')}</tr>`).join('')}
  </tbody></table>`;

  if (count) count.textContent = `${data.length} registros`;
  mkPagination('rollos-comercio-pagination', rollosComercioPage, pages, 'goRollosComercioPage');
}

function goRollosComercioPage(p) {
  const pages = Math.ceil(ROLLOS_COMERCIO.length / ROLLOS_PAGE_SIZE);
  if (p >= 1 && p <= pages) { rollosComercioPage = p; renderRollosComercioTable(); }
}

// ── Tabla Referencias ─────────────────────────────────────────────
function renderRollosRefTable() {
  const wrap  = document.getElementById('rollos-ref-wrap');
  const count = document.getElementById('rollos-ref-count');
  if (!wrap) return;
  const data = ROLLOS_RAW.referencias || [];

  if (!data.length) {
    wrap.innerHTML = '<div class="empty-state"><div class="icon">📭</div><p>Sin referencias</p></div>';
    if (count) count.textContent = '0 registros';
    return;
  }

  wrap.innerHTML = `<table><thead><tr>
    <th>Referencia</th><th>Cantidad</th><th>Estado</th><th>Departamento</th><th>Ciudad</th><th>Material</th>
  </tr></thead><tbody>
    ${data.map(r=>`<tr>
      <td>${r.referencia||'—'}</td>
      <td style="font-family:'JetBrains Mono',monospace;font-weight:600;color:var(--azul-cielo)">${r.cantidad??'—'}</td>
      <td>${statusPill(r.estado)}</td>
      <td>${r.departamento||'—'}</td>
      <td>${r.ciudad||'—'}</td>
      <td>${r.material||'—'}</td>
    </tr>`).join('')}
  </tbody></table>`;

  if (count) count.textContent = `${data.length} referencias`;
}

// ── Gráfica departamento ──────────────────────────────────────────
function renderRollosDeptChart() {
  const canvas = document.getElementById('chart-rollos-depto');
  if (!canvas || !ROLLOS_RAW) return;
  const por_dep = (ROLLOS_RAW.por_departamento || []).slice(0, 15);
  if (chartInstances['rollos-depto']) { chartInstances['rollos-depto'].destroy(); }
  chartInstances['rollos-depto'] = new Chart(canvas, {
    type: 'bar',
    data: {
      labels: por_dep.map(d => d.departamento),
      datasets: [{
        label: 'Rollos',
        data: por_dep.map(d => d.total_rollos),
        backgroundColor: por_dep.map((_,i) => WOMPI_COLORS[i % WOMPI_COLORS.length] + 'BB'),
        borderColor: por_dep.map((_,i) => WOMPI_COLORS[i % WOMPI_COLORS.length]),
        borderWidth: 1, borderRadius: 6,
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false, indexAxis: 'y',
      plugins: { legend:{ display:false }, tooltip: CHART_OPTS.tooltip },
      scales: {
        x: { ticks:{ color:'#7A7674' }, grid:{ color:'rgba(255,255,255,.05)' } },
        y: { ticks:{ color:'#FAFAFA', font:{ size:11 } }, grid:{ display:false } },
      },
    }
  });
}

// ── Filtros Detalle ───────────────────────────────────────────────
function applyRollosFilters() {
  if (!ROLLOS_RAW) return;
  const codigo = (document.getElementById('rf-codigo-tarea')?.value||'').trim().toUpperCase();
  const guia   = (document.getElementById('rf-guia')?.value||'').trim().toUpperCase();
  const estado = (document.getElementById('rf-estado')?.value||'').toUpperCase();
  const desde  = document.getElementById('rf-fecha-desde')?.value;
  const hasta  = document.getElementById('rf-fecha-hasta')?.value;
  const anio   = document.getElementById('rf-anio')?.value;
  const mes    = document.getElementById('rf-mes')?.value;

  ROLLOS_DETALLE = (ROLLOS_RAW.detalle || []).filter(r => {
    if (codigo && !(r.codigo_tarea||'').toUpperCase().includes(codigo)) return false;
    if (guia   && !(r.guia||'').toUpperCase().includes(guia))           return false;
    if (estado && (r.estado||'').toUpperCase() !== estado)              return false;
    if (anio   && String(r.anio) !== anio)                              return false;
    if (mes    && String(r.mes).padStart(2,'0') !== mes.padStart(2,'0')) return false;
    if (desde || hasta) {
      const fp = r.fecha_plan_fin ? new Date(r.fecha_plan_fin.substring(0,10)) : null;
      if (desde && fp && fp < new Date(desde)) return false;
      if (hasta && fp && fp > new Date(hasta + 'T23:59:59')) return false;
    }
    return true;
  });
  rollosDetallePage = 1;
  renderRollosDetalleTable();
}

function resetRollosFilters() {
  ['rf-codigo-tarea','rf-guia','rf-fecha-desde','rf-fecha-hasta'].forEach(id => {
    const el = document.getElementById(id); if (el) el.value = '';
  });
  ['rf-estado','rf-anio','rf-mes'].forEach(id => {
    const el = document.getElementById(id); if (el) el.value = '';
  });
  ROLLOS_DETALLE = (ROLLOS_RAW?.detalle || []).slice();
  rollosDetallePage = 1;
  renderRollosDetalleTable();
}

// ── Filtros Comercio ──────────────────────────────────────────────
function applyComercioFilters() {
  if (!ROLLOS_RAW) return;
  const cod   = (document.getElementById('rf-cod-comercio')?.value||'').trim().toUpperCase();
  const est   = (document.getElementById('rf-com-estado')?.value||'').toUpperCase();
  const tipo  = (document.getElementById('rf-com-tipo')?.value||'').toUpperCase();
  const mes   = document.getElementById('rf-com-mes')?.value;

  ROLLOS_COMERCIO = (ROLLOS_RAW.comercio || []).filter(r => {
    if (cod  && !(r.cod_comercio||'').toUpperCase().includes(cod))  return false;
    if (est  && (r.estado||'').toUpperCase() !== est)               return false;
    if (tipo && (r.tipo_envio||'').toUpperCase() !== tipo)          return false;
    // mes: filtrar desde detalle — si hay mes, filtrar comercios que tengan tareas en ese mes
    if (mes) {
      const cod_c = r.cod_comercio;
      const hasInMes = (ROLLOS_RAW.detalle||[]).some(d =>
        d.cod_sitio === cod_c && String(d.mes).padStart(2,'0') === mes.padStart(2,'0'));
      if (!hasInMes) return false;
    }
    return true;
  });
  rollosComercioPage = 1;
  renderRollosComercioTable();
  renderRollosDeptChart();
}

function resetComercioFilters() {
  ['rf-cod-comercio'].forEach(id => { const el=document.getElementById(id); if(el) el.value=''; });
  ['rf-com-estado','rf-com-tipo','rf-com-mes'].forEach(id => { const el=document.getElementById(id); if(el) el.value=''; });
  ROLLOS_COMERCIO = (ROLLOS_RAW?.comercio||[]).slice();
  rollosComercioPage = 1;
  renderRollosComercioTable();
  renderRollosDeptChart();
}

// ── Exportar Excel (Rollos) ───────────────────────────────────────
function exportRollosDetalleExcel() {
  const data = ROLLOS_DETALLE.map(r => {
    const o = {};
    DETALLE_COLS.forEach(c => { o[c.label] = c.fn(r); });
    return o;
  });
  _exportExcelRollos(data, 'Rollos_Detalle');
}

function exportComercioExcel() {
  const data = ROLLOS_COMERCIO.map(r => {
    const o = {};
    COMERCIO_COLS.forEach(c => { o[c.label] = c.fn(r); });
    return o;
  });
  _exportExcelRollos(data, 'Rollos_Comercios');
}

function exportReferenciasExcel() {
  const data = (ROLLOS_RAW?.referencias || []).map(r => ({
    Referencia: r.referencia, Cantidad: r.cantidad, Estado: r.estado,
    Departamento: r.departamento, Ciudad: r.ciudad, Material: r.material,
  }));
  _exportExcelRollos(data, 'Rollos_Referencias');
}

function _exportExcelRollos(data, filename) {
  if (!data.length) { alert('Sin datos para exportar.'); return; }
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(data);
  ws['!cols'] = Object.keys(data[0]||{}).map(k => ({ wch: Math.max(k.length, ...data.slice(0,50).map(r=>String(r[k]||'').length)) }));
  XLSX.utils.book_append_sheet(wb, ws, 'Datos');
  XLSX.writeFile(wb, `${filename}_${new Date().toISOString().slice(0,10)}.xlsx`);
}