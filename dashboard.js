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

// Fecha de corte global — igual que vp.py: solo desde 2026-03-01
const MARCH_2026 = new Date(2024, 2, 1);   // mes 2 = marzo (0-indexed)

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
    if (d && !isNaN(d.getTime()) && d.getFullYear() > 2000) return d;
  }
  // fallback nativo
  const fb = new Date(s);
  return (!isNaN(fb.getTime()) && fb.getFullYear() > 2000) ? fb : null;
}

function diffDays(a, b) {
  return Math.floor((b - a) / 86400000);
}

// ══════════════════════════════════════════════════════════════════
//  FILTRO GLOBAL: solo datos desde marzo 2026  (=== vp.py)
// ══════════════════════════════════════════════════════════════════
function applyDateFilter() {
  FILTERED = RAW_DATA.filter(r => {
    const fs = getCol(r, 'FECHA DE SOLICITUD', 'fecha de solicitud');
    if (!fs || fs === '') return true;   // sin fecha: pasa (igual que vp.py)
    const d = parseDate(fs);
    if (!d) return true;
    return d >= MARCH_2026;
  });
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
    if (desde && fd && fd < new Date(desde)) return false;
    if (hasta && fd && fd > new Date(hasta + 'T23:59:59')) return false;

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
  const fd = document.getElementById('f-fecha-desde');
  const fh = document.getElementById('f-fecha-hasta');
  if (fd) fd.value = '';
  if (fh) fh.value = '';

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

  // 3. Devueltos: DEVOLUCION | DEVUELTO | REMITENTE  (igual que vp.py)
  const devueltos = df.filter(r => {
    const e = getCol(r, 'ESTADO DATAFONO', 'estado datafono').toUpperCase();
    return e.includes('DEVOLUCION') || e.includes('DEVUELTO') || e.includes('REMITENTE');
  }).length;

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

  // 9. Vencen hoy / vencidas
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const vencenHoy = df.filter(r => {
    const lim = parseDate(getCol(r, 'FECHA LIMITE DE ENTREGA', 'fecha limite de entrega'));
    const est = getCol(r, 'ESTADO DATAFONO', 'estado datafono').toUpperCase();
    if (!lim) return false;
    const limD = new Date(lim); limD.setHours(0, 0, 0, 0);
    return limD.getTime() === today.getTime() && est !== 'ENTREGADO' && est !== 'CANCELADO';
  }).length;

  const vencidas = df.filter(r => {
    const lim = parseDate(getCol(r, 'FECHA LIMITE DE ENTREGA', 'fecha limite de entrega'));
    const est = getCol(r, 'ESTADO DATAFONO', 'estado datafono').toUpperCase();
    return lim && lim < today && est !== 'ENTREGADO' && est !== 'CANCELADO';
  }).length;

  // 10. Primer intento (entregados sin novedades — igual que index.html previo)
  const primerIntento = entDf.filter(r => {
    const nov = getCol(r, 'NOVEDADES', 'novedades', 'Novedades');
    return !nov || nov.trim() === '' || nov.trim() === '0';
  }).length;

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
    vencenHoy, vencidas,
    primerIntento,
    pctPrimerIntento: pct(primerIntento, entregados),
    ec,
  };
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
      sub:`Sin cancelados (${k.cancelados} cancelados)` },
    { label:'Alistados',          value:`${k.n_alistados} (${k.pctNAlistados}%)`,
      color:'lime',  icon:'⚙️',  sub:'Total – en alistamiento', pct:k.pctNAlistados, pctColor:'lime' },
    { label:'Entregados',         value:k.entregados,     color:'selva', icon:'✅',
      sub:`${k.pctEntregado}%`, pct:k.pctEntregado },
    { label:'En Tránsito',        value:k.en_transito,    color:'blue',  icon:'🚚',
      sub:`${k.pctTransito}%`, pct:k.pctTransito, pctColor:'blue' },
    { label:'En Alistamiento',    value:k.en_alistamiento,color:'lime',  icon:'🔧',
      sub:`${k.pctAlistamiento}%`, pct:k.pctAlistamiento, pctColor:'lime' },
    { label:'Devueltos',          value:k.devueltos,      color:'danger',icon:'↩️',
      sub:`${k.total ? Math.round(k.devueltos/k.total*100) : 0}% del total`,
      pct: k.total ? Math.round(k.devueltos/k.total*100) : 0,
      pctColor:'danger', alert:k.devueltos > 0 },
    { label:'% Oportunidad ANS',  value:k.pctOport+'%',   color:'green', icon:'🎯',
      sub:'Entregas en plazo', pct:k.pctOport },
    { label:'% Calidad',          value:k.pctCalidad+'%', color:'blue',  icon:'💎',
      sub:'Sin devoluciones', pct:k.pctCalidad, pctColor:'blue' },
    { label:'Visita Técnica',
      value:`${k.entVT} ejec / ${k.programados_vt} prog / ${k.totalVT} total`,
      color:'lime', icon:'🔧', sub:`${k.pctVT}% ejecutado`, pct:k.pctVT, pctColor:'lime' },
    { label:'Op. Logístico',      value:`${k.entOL}/${k.totalOL}`, color:'blue', icon:'📮',
      sub:`${k.pctOL}% entregado`, pct:k.pctOL },
    { label:'Vencen Hoy',         value:k.vencenHoy,      color:'warn',  icon:'⏰',
      sub:'Sin entregar, límite hoy', alert:k.vencenHoy > 0 },
    { label:'Vencidas ANS',       value:k.vencidas,       color:'danger',icon:'🚨',
      sub:'Fuera de ANS', alert:k.vencidas > 0 },
    { label:'1er Intento',        value:k.primerIntento,  color:'selva', icon:'🎯',
      sub:`${k.pctPrimerIntento}% del entregado`, pct:k.pctPrimerIntento },
  ];

  grid.innerHTML = cards.map((c, i) => `
    <div class="kpi-card ${c.color} fade-up" style="animation-delay:${i*.04}s">
      ${c.alert ? '<div class="kpi-alert-badge"></div>' : ''}
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
    </div>
  `).join('');
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
  const devRows = FILTERED.filter(r => {
    const e = getCol(r,'ESTADO DATAFONO','estado datafono').toUpperCase();
    return e.includes('DEVOLUCION') || e.includes('DEVUELTO') || e.includes('REMITENTE');
  });
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
  tbody.innerHTML = d.labels.map((dep, i) => {
    const p=d.principal[i], in_=d.intermedia[i], l=d.lejana[i], t=p+in_+l;
    totP+=p; totI+=in_; totL+=l;
    return `<tr><td><strong>${dep}</strong></td><td>${p}</td><td>${in_}</td><td>${l}</td><td><strong>${t}</strong></td></tr>`;
  }).join('') +
    `<tr class="dept-total"><td><strong>TOTAL</strong></td><td>${totP}</td><td>${totI}</td><td>${totL}</td><td><strong>${totP+totI+totL}</strong></td></tr>`;
}

// ══════════════════════════════════════════════════════════════════
//  ANS ALERTS
// ══════════════════════════════════════════════════════════════════
function renderANSAlerts(k) {
  const grid = document.getElementById('ans-alerts-grid');
  if (!grid) return;
  const now = new Date();

  const guiasEst = FILTERED.filter(r => {
    const est = getCol(r,'ESTADO DATAFONO','estado datafono').toUpperCase();
    if (est === 'ENTREGADO' || est === 'CANCELADO') return false;
    const fEnt = parseDate(getCol(r,'FECHA DE ENTREGA','fecha de entrega'));
    return fEnt !== null && diffDays(fEnt, now) >= 1;
  }).length;

  const FAILED_WORDS = ['INVALIDA','INVALIDO','NO ESTÁ','NO ESTA','CERRADO','CIERRE','RECHAZO','CLIENTE AUSENTE'];
  const fallidos = FILTERED.filter(r => {
    const nov = (getCol(r,'NOVEDADES','novedades') + ' ' + getCol(r,'CAUSAL INC','causal inc')).toUpperCase();
    return FAILED_WORDS.some(w => nov.includes(w));
  }).length;

  const pctOL = k.total ? Math.round(k.totalOL/k.total*100) : 0;
  const pctVT = k.total ? Math.round(k.totalVT/k.total*100) : 0;

  const alerts = [
    { label:'Vencen Hoy',           value:k.vencenHoy,           type:k.vencenHoy>0?'':'ok',    sub:'Límite hoy sin entregar' },
    { label:'Vencidas ANS',         value:k.vencidas,             type:k.vencidas>0?'':'ok',     sub:'Fuera de plazo' },
    { label:'1er Intento',          value:k.pctPrimerIntento+'%', type:'ok',                     sub:`${k.primerIntento} de ${k.entregados}` },
    { label:'Guías Estancadas >24h',value:guiasEst,               type:guiasEst>0?'warn':'ok',   sub:'Sin eventos registrados' },
    { label:'Intentos Fallidos',    value:fallidos,               type:fallidos>0?'warn':'ok',   sub:'Dir. inválida, ausente, cierre, rechazo' },
    { label:'% Op. Logístico',      value:pctOL+'%',              type:'info',                   sub:`${k.totalOL} vía OPLG` },
    { label:'% Visita Técnica',     value:pctVT+'%',              type:'info',                   sub:`${k.totalVT} gestionadas por VT` },
    { label:'Devoluciones',         value:k.devueltos,            type:k.devueltos>0?'warn':'ok',sub:'Total devueltos' },
  ];

  grid.innerHTML = alerts.map(a => `
    <div class="alert-card ${a.type}">
      <div class="alert-label">${a.label}</div>
      <div class="alert-value">${a.value}</div>
      <div class="alert-sub">${a.sub}</div>
    </div>
  `).join('');

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
  const cutoff = new Date(now.getTime() + hrs * 3600000);
  const wrap   = document.getElementById('backlog-wrap');
  if (!wrap) return;

  const atRisk = FILTERED.filter(r => {
    const lim = parseDate(getCol(r,'FECHA LIMITE DE ENTREGA','fecha limite de entrega'));
    const est = getCol(r,'ESTADO DATAFONO','estado datafono').toUpperCase();
    return lim && est !== 'ENTREGADO' && est !== 'CANCELADO' && lim >= now && lim <= cutoff;
  });

  if (!atRisk.length) {
    wrap.innerHTML = '<div class="empty-state"><div class="icon">✅</div><p>Sin registros en riesgo para la ventana seleccionada</p></div>';
    return;
  }
  wrap.innerHTML = `<table>
    <thead><tr><th>Comercio</th><th>ID Sitio</th><th>Guía</th><th>Fecha Límite</th><th>Transportadora</th><th>Estado</th><th>Riesgo</th></tr></thead>
    <tbody>${atRisk.map(r => {
      const lim  = parseDate(getCol(r,'FECHA LIMITE DE ENTREGA','fecha limite de entrega'));
      const hLeft= lim ? Math.round((lim-now)/3600000) : 0;
      return `<tr>
        <td>${getCol(r,'Nombre del comercio','nombre del comercio','NOMBRE DEL COMERCIO')||'—'}</td>
        <td style="font-family:'JetBrains Mono',monospace;font-size:11px">${getCol(r,'ID Comercio','id comercio')||'—'}</td>
        <td style="font-family:'JetBrains Mono',monospace;font-size:11px">${getCol(r,'NÚMERO DE GUIA','NUMERO DE GUIA','numero de guia')||'—'}</td>
        <td>${lim?lim.toLocaleDateString('es-CO'):'—'}</td>
        <td>${getCol(r,'TRANSPORTADORA','Transportadora','transportadora')||'—'}</td>
        <td><span class="status-pill ${statusClass(getCol(r,'ESTADO DATAFONO','estado datafono'))}">${getCol(r,'ESTADO DATAFONO','estado datafono')||'—'}</span></td>
        <td><span class="risk-badge ${hLeft<=24?'risk-24':'risk-48'}">${hLeft}h</span></td>
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
  const now = new Date();

  const stalled = FILTERED.filter(r => {
    const est  = getCol(r,'ESTADO DATAFONO','estado datafono').toUpperCase();
    if (est === 'ENTREGADO' || est === 'CANCELADO') return false;
    const fEnt = parseDate(getCol(r,'FECHA DE ENTREGA','fecha de entrega'));
    return fEnt !== null && diffDays(fEnt, now) >= 1;
  });

  if (!stalled.length) {
    wrap.innerHTML = '<div class="empty-state"><div class="icon">✅</div><p>Sin guías estancadas</p></div>';
    return;
  }
  wrap.innerHTML = `<table>
    <thead><tr>
      <th>Comercio</th><th>Guía</th><th>Último Evento</th>
      <th style="color:var(--warning)">Días Sin Cambios</th>
      <th>Transportadora</th><th>Estado</th>
    </tr></thead>
    <tbody>${stalled.sort((a,b)=>{
      const da=parseDate(getCol(a,'FECHA DE ENTREGA','fecha de entrega'));
      const db=parseDate(getCol(b,'FECHA DE ENTREGA','fecha de entrega'));
      return (da||0)-(db||0);
    }).map(r=>{
      const fEnt = parseDate(getCol(r,'FECHA DE ENTREGA','fecha de entrega'));
      const dias = fEnt ? diffDays(fEnt,now) : null;
      const cls  = dias===null?'ok':dias>=7?'crit':dias>=3?'warn':'ok';
      return `<tr>
        <td>${getCol(r,'Nombre del comercio','nombre del comercio','NOMBRE DEL COMERCIO')||'—'}</td>
        <td style="font-family:'JetBrains Mono',monospace;font-size:11px">${getCol(r,'NÚMERO DE GUIA','NUMERO DE GUIA','numero de guia')||'—'}</td>
        <td style="color:var(--muted)">${fEnt?fEnt.toLocaleDateString('es-CO'):'—'}</td>
        <td><span class="days-stalled ${cls}">${dias!==null?dias+' día'+(dias!==1?'s':''):'—'}</span></td>
        <td>${getCol(r,'TRANSPORTADORA','Transportadora','transportadora')||'—'}</td>
        <td><span class="status-pill ${statusClass(getCol(r,'ESTADO DATAFONO','estado datafono'))}">${getCol(r,'ESTADO DATAFONO','estado datafono')||'—'}</span></td>
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
  const FAILED_WORDS = ['INVALIDA','INVALIDO','NO ESTÁ','NO ESTA','CERRADO','CIERRE','RECHAZO','CLIENTE AUSENTE'];
  const fallidos = FILTERED.filter(r => {
    const nov = (getCol(r,'NOVEDADES','novedades') + ' ' + getCol(r,'CAUSAL INC','causal inc')).toUpperCase();
    return FAILED_WORDS.some(w => nov.includes(w));
  });
  if (!fallidos.length) {
    wrap.innerHTML = '<div class="empty-state"><div class="icon">✅</div><p>Sin intentos fallidos</p></div>';
    return;
  }
  wrap.innerHTML = `<table>
    <thead><tr><th>Comercio</th><th>Guía</th><th style="color:var(--danger)">Motivo</th><th>Transportadora</th><th>Estado</th></tr></thead>
    <tbody>${fallidos.map(r=>`<tr>
      <td>${getCol(r,'Nombre del comercio','nombre del comercio','NOMBRE DEL COMERCIO')||'—'}</td>
      <td style="font-family:'JetBrains Mono',monospace;font-size:11px">${getCol(r,'NÚMERO DE GUIA','NUMERO DE GUIA','numero de guia')||'—'}</td>
      <td style="color:var(--danger)">${getCol(r,'NOVEDADES','novedades')||getCol(r,'CAUSAL INC','causal inc')||'—'}</td>
      <td>${getCol(r,'TRANSPORTADORA','Transportadora','transportadora')||'—'}</td>
      <td><span class="status-pill ${statusClass(getCol(r,'ESTADO DATAFONO','estado datafono'))}">${getCol(r,'ESTADO DATAFONO','estado datafono')||'—'}</span></td>
    </tr>`).join('')}</tbody>
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
  const cutoff = new Date(now.getTime()+hrs*3600000);
  const data   = FILTERED.filter(r=>{
    const lim = parseDate(getCol(r,'FECHA LIMITE DE ENTREGA','fecha limite de entrega'));
    const est = getCol(r,'ESTADO DATAFONO','estado datafono').toUpperCase();
    return lim && est!=='ENTREGADO' && est!=='CANCELADO' && lim>=now && lim<=cutoff;
  }).map(r=>({
    Comercio:       getCol(r,'Nombre del comercio','nombre del comercio','NOMBRE DEL COMERCIO'),
    'ID Sitio':     getCol(r,'ID Comercio','id comercio'),
    Guía:           getCol(r,'NÚMERO DE GUIA','NUMERO DE GUIA','numero de guia'),
    'Fecha Límite': getCol(r,'FECHA LIMITE DE ENTREGA','fecha limite de entrega'),
    Transportadora: getCol(r,'TRANSPORTADORA','Transportadora'),
    Estado:         getCol(r,'ESTADO DATAFONO','estado datafono'),
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
  ['tracking','detalle','tabla'].forEach(t => {
    const panel = document.getElementById('panel-'+t);
    const btn   = document.getElementById('tab-'+t);
    if (panel) panel.style.display = t===tab ? 'block' : 'none';
    if (btn)   btn.classList.toggle('active', t===tab);
  });
  if (tab==='detalle') { renderDevCharts(); renderBacklog(); renderStalledGuias(); renderFallidos(); }
  if (tab==='tabla')   renderMainTable();
}

// ══════════════════════════════════════════════════════════════════
//  INIT
// ══════════════════════════════════════════════════════════════════
function initDashboard() { loadData(); }

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
    // Fechas desde marzo 2026 para pasar el filtro global
    const fSol = new Date(2026, 2 + Math.floor(i/40), 1+(i%28));
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