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

// ── MULTI-SELECT HELPER ───────────────────────────────────────────
// ── Multiselect con buscador y renderizado virtual ────────────────
// Umbral: ≥ este número de opciones → muestra input de búsqueda
const MS_SEARCH_THRESHOLD = 10;
const MS_VIRTUAL_MAX      = 120; // máx items renderizados en el DOM a la vez

window._setupMS = function(id, vals) {
  const container = document.getElementById(id);
  if (!container) return;
  const placeholder = container.dataset.placeholder || 'Todos';
  const useSearch   = vals.length >= MS_SEARCH_THRESHOLD;

  container._msAllVals      = vals;          // lista original completa
  container._selectedValues = [];
  container._msFilteredVals = vals.slice();  // lista actualmente visible

  container.innerHTML = `
    <div class="ms-trigger">${placeholder}</div>
    <div class="ms-dropdown">
      ${useSearch ? `
      <div class="ms-search-wrap">
        <input class="ms-search-input" type="text" placeholder="🔍 Buscar..." autocomplete="off"
               onclick="event.stopPropagation()"
               oninput="window._msFilterItems('${id}', this.value)">
        <span class="ms-search-count"></span>
      </div>` : ''}
      <div class="ms-items-wrap"></div>
      <div class="ms-actions">
        <button class="ms-btn" onclick="event.stopPropagation(); window._msAction('${id}', 'clear')">Limpiar</button>
        <button class="ms-btn ms-btn-apply" onclick="event.stopPropagation(); window._msAction('${id}', 'apply')">Hecho</button>
      </div>
    </div>
  `;

  const trigger = container.querySelector('.ms-trigger');
  trigger.onclick = (e) => {
    e.stopPropagation();
    const isOpen = container.classList.contains('open');
    document.querySelectorAll('.ms-container').forEach(c => c.classList.remove('open'));
    if (!isOpen) {
      container.classList.add('open');
      // Renderizar items al abrir (lazy)
      _msRenderItems(container);
      // Foco en búsqueda si aplica
      const si = container.querySelector('.ms-search-input');
      if (si) setTimeout(() => si.focus(), 50);
    }
  };
};

// Renderiza hasta MS_VIRTUAL_MAX items del filtrado actual
function _msRenderItems(container) {
  const wrap     = container.querySelector('.ms-items-wrap');
  if (!wrap) return;
  const filtered = container._msFilteredVals || [];
  const selected = new Set(container._selectedValues || []);
  const slice    = filtered.slice(0, MS_VIRTUAL_MAX);
  const hasMore  = filtered.length > MS_VIRTUAL_MAX;

  wrap.innerHTML = slice.map(v => {
    const key  = String(v).toUpperCase();
    const isSel= selected.has(key);
    return `<div class="ms-item${isSel ? ' selected' : ''}" data-val="${key}">
      <input type="checkbox"${isSel ? ' checked' : ''}>
      <span class="ms-item-label">${v}</span>
    </div>`;
  }).join('') + (hasMore
    ? `<div class="ms-more-hint">… ${(filtered.length - MS_VIRTUAL_MAX).toLocaleString('es-CO')} más — refiná la búsqueda</div>`
    : '');

  // Actualizar contador de búsqueda
  const countEl = container.querySelector('.ms-search-count');
  if (countEl) {
    const id = container.id;
    const all = (container._msAllVals || []).length;
    countEl.textContent = filtered.length < all
      ? `${filtered.length.toLocaleString('es-CO')} de ${all.toLocaleString('es-CO')}`
      : `${all.toLocaleString('es-CO')} opciones`;
  }

  // Re-bindear eventos de click en los items renderizados
  wrap.querySelectorAll('.ms-item').forEach(item => {
    item.onclick = (e) => {
      e.stopPropagation();
      item.classList.toggle('selected');
      const cb  = item.querySelector('input');
      cb.checked = !cb.checked;
      const val  = item.dataset.val;
      const sels = container._selectedValues;
      if (cb.checked) { if (!sels.includes(val)) sels.push(val); }
      else            { container._selectedValues = sels.filter(v => v !== val); }
      _updateMSLabel(container);
    };
  });
}

window._msFilterItems = function(id, term) {
  const container = document.getElementById(id);
  if (!container) return;
  const t = (term || '').toLowerCase().trim();
  container._msFilteredVals = t
    ? (container._msAllVals || []).filter(v => String(v).toLowerCase().includes(t))
    : (container._msAllVals || []).slice();
  _msRenderItems(container);
};

function _updateMSLabel(container) {
  const trigger = container.querySelector('.ms-trigger');
  const vals    = container._selectedValues;
  if (vals.length === 0)      trigger.textContent = container.dataset.placeholder || 'Todos';
  else if (vals.length === 1) trigger.textContent = vals[0];
  else                        trigger.textContent = `${vals.length} seleccionados`;
}

window._msAction = function(id, action) {
  const container = document.getElementById(id);
  if (!container) return;
  if (action === 'clear') {
    container._selectedValues = [];
    container._msFilteredVals = (container._msAllVals || []).slice();
    // Limpiar búsqueda
    const si = container.querySelector('.ms-search-input');
    if (si) si.value = '';
    _msRenderItems(container);
    _updateMSLabel(container);
  } else if (action === 'apply') {
    container.classList.remove('open');
  }
};

window._msGetSels = function(id) {
  const el = document.getElementById(id);
  if (el && el._selectedValues && el._selectedValues.length > 0) return el._selectedValues;
  const val = el?.value;
  return (val && val !== '') ? [val.toUpperCase()] : null;
};

document.addEventListener('click', () => document.querySelectorAll('.ms-container').forEach(c => c.classList.remove('open')));

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
      _mainLoaded = true;
      _updateLoadingUI();
    })
    .catch(() => {
      // Demo data si no hay data.json
      RAW_DATA = getDemoData();
      populateFilters();
      applyDateFilter();
      renderAll();
      _mainLoaded = true;
      _updateLoadingUI();
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

    // Si hay algún filtro de fecha de solicitud activo, excluir registros sin fecha
    const hayFiltroFecha = mes || desde || hasta;
    if (hayFiltroFecha && !fd) return false;

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
    const e = getCol(r, 'ESTADO DATAFONO', 'estado datafono').toUpperCase().trim() || 'SIN ESTADO';
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

  // pctOport: incumplimientos = entregados con RESPONSABLE INCUMPLIMIENTO === LINEACOM.
  // No depende del campo CUMPLE ANS (puede estar vacío); usa directamente el responsable.
  const noCumpleLineacom = entDf.filter(r =>
    getCol(r, 'RESPONSABLE INCUMPLIMIENTO', 'responsable incumplimiento').trim().toUpperCase().includes('LINEACOM')
  ).length;
  // DEBUG: ver valores reales del campo en entregados
  const _respValores = [...new Set(entDf.map(r => getCol(r, 'RESPONSABLE INCUMPLIMIENTO', 'responsable incumplimiento')))];
  console.log('[ANS Debug] RESPONSABLE INCUMPLIMIENTO valores únicos en entregados:', _respValores);
  console.log('[ANS Debug] entDf:', entDf.length, '| noCumpleLineacom:', noCumpleLineacom);
  const pctOport = entDf.length ? Math.round((entDf.length - noCumpleLineacom) / entDf.length * 100) : 0;

  const pctCalidad = entregados   ? Math.round((entregados - devueltos) / entregados * 100) : 100;

  // 9. Vencen hoy / vencidas (incluye VT y OPLG)
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const vencenHoyRows = df.filter(r => {
    const lim = parseDate(getCol(r, 'FECHA LIMITE DE ENTREGA', 'fecha limite de entrega'));
    const est = getCol(r, 'ESTADO DATAFONO', 'estado datafono').toUpperCase().trim();
    if (!lim) return false;
    if (est === 'ENTREGADO' || est.includes('ENTREGADO') || est === 'CANCELADO') return false;
    const limD = new Date(lim); limD.setHours(0, 0, 0, 0);
    return limD.getTime() === today.getTime();
  });
  const vencenHoy = vencenHoyRows.length;

  // Vencidas: TODOS los registros (VT + OPLG) no entregados con fecha limite pasada
  const vencidasRows = df.filter(r => {
    const lim = parseDate(getCol(r, 'FECHA LIMITE DE ENTREGA', 'fecha limite de entrega'));
    const est = getCol(r, 'ESTADO DATAFONO', 'estado datafono').toUpperCase().trim();
    if (!lim) return false;
    if (est === 'ENTREGADO' || est.includes('ENTREGADO') || est === 'CANCELADO') return false;
    return lim < today;
  });
  const vencidas = vencidasRows.length;

  // Incumplimientos Totales (historico completo):
  // ENTREGADO con FECHA ENTREGA AL COMERCIO > FECHA LIMITE DE ENTREGA
  // + no entregados con fecha limite vencida
  const incumplimientosRows = df.filter(r => {
    const est = getCol(r, 'ESTADO DATAFONO', 'estado datafono').toUpperCase();
    const lim = getFechaLimite(r);
    if (!lim || est === 'CANCELADO') return false;
    const limD = new Date(lim); limD.setHours(0,0,0,0);
    if (est === 'ENTREGADO') {
      const fe = parseDate(getCol(r, 'FECHA ENTREGA AL COMERCIO', 'fecha entrega al comercio'));
      if (!fe) return false;
      const feD = new Date(fe); feD.setHours(0,0,0,0);
      return feD > limD;
    }
    return limD < today;
  });
  const incumplimientos = incumplimientosRows.length;

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
    pctOport, pctCalidad, cumpleOport, noCumpleOport: noCumpleLineacom,
    vencenHoy, vencenHoyRows,
    vencidas, vencidasRows,
    incumplimientos, incumplimientosRows,
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
  if (badge) badge.textContent = `App Activa`;
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
      sub:`${k.entregados - k.noCumpleOport} cumplen · ${k.noCumpleOport} no cumplen LINEACOM (de ${k.entregados} entregados)`,
      pct:k.pctOport,
      rows: k.entregadosRows.filter(r=>getCol(r,'CUMPLE ANS','cumple ans').toUpperCase()!=='SI' && getCol(r,'RESPONSABLE INCUMPLIMIENTO','responsable incumplimiento').toUpperCase()==='LINEACOM') },
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
      sub:'Sin entregar y fuera de ANS', alert:k.vencidas > 0, rows: k.vencidasRows },
    { label:'Incumplimientos Totales', value:k.incumplimientos, color:'danger', icon:'📛',
      sub:'Todos los estados (incl. entregados tardíos)', alert:k.incumplimientos > 0,
      rows: k.incumplimientosRows },
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
    const est = getCol(r,'ESTADO DATAFONO','estado datafono').toUpperCase().trim();
    if (!lim) return false;
    if (est === 'ENTREGADO' || est.includes('ENTREGADO') || est === 'CANCELADO') return false;
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
    const est = getCol(r,'ESTADO DATAFONO','estado datafono').toUpperCase().trim();
    if (!lim) return false;
    if (est==='ENTREGADO' || est.includes('ENTREGADO') || est==='CANCELADO') return false;
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
//  TABS (mantiene compatibilidad original)
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
//  SIDEBAR NAVIGATION
// ══════════════════════════════════════════════════════════════════
let _currentBoard = null;
let _currentTab   = null;

function _toggleSidebar() {
  const sidebar = document.getElementById('sidebar');
  if (sidebar) sidebar.classList.toggle('collapsed');
}

function _selectBoard(board) {
  _currentBoard = board;

  const section = document.getElementById('sb-board-' + board);
  const isOpen  = section?.classList.contains('open');

  // Collapse all boards
  ['datafonos','rollos','inventario'].forEach(b => {
    document.getElementById('sb-board-' + b)?.classList.remove('open');
    document.getElementById('sb-btn-' + b)?.classList.remove('active');
  });

  // If wasn't open, open it and navigate to first tab
  if (!isOpen) {
    section?.classList.add('open');
    document.getElementById('sb-btn-' + board)?.classList.add('active');

    const label = document.getElementById('topbar-board-label');
    if (label) label.textContent = board === 'datafonos' ? 'Tracking VP · Datafonos' : board === 'inventario' ? 'Inventario Wompi' : 'Tracking Rollos Trim y VP';

    if (board === 'datafonos') _selectBoardTab('datafonos', 'tracking');
    else if (board === 'inventario') _selectBoardTab('inventario', 'inv-principal');
    else _selectBoardTab('rollos', 'rollos-main');
  }
  // If was already open (toggle close) — just collapse, keep current panel visible (don't redirect to home)
}

function _selectBoardTab(board, tab) {
  _currentBoard = board;
  _currentTab   = tab;

  // Ensure the board is expanded and marked active
  ['datafonos','rollos','inventario'].forEach(b => {
    document.getElementById('sb-board-' + b)?.classList.remove('open');
    document.getElementById('sb-btn-'   + b)?.classList.remove('active');
  });
  document.getElementById('sb-board-' + board)?.classList.add('open');
  document.getElementById('sb-btn-'   + board)?.classList.add('active');

  // Topbar label
  const label = document.getElementById('topbar-board-label');
  if (label) label.textContent = board === 'datafonos' ? 'Tracking VP · Datafonos' : board === 'inventario' ? 'Inventario Wompi' : 'Tracking Rollos Trim y VP';

  document.querySelectorAll('.sidebar-tab').forEach(el => el.classList.remove('active'));
  const activeEl = document.getElementById('sb-tab-' + tab);
  if (activeEl) activeEl.classList.add('active');

  // Si salimos de incumplimientos, limpiar su contenido para evitar que quede visible
  if (tab !== 'incumplimientos') {
    const resumen = document.getElementById('incump-resumen');
    const tableWrap = document.getElementById('incump-table-wrap');
    const pagination = document.getElementById('incump-pagination');
    if (resumen) resumen.innerHTML = '';
    if (tableWrap) tableWrap.innerHTML = '';
    if (pagination) pagination.innerHTML = '';
  }

  _showAllPanels(tab);

  if (tab === 'detalle')         { renderDevCharts(); renderBacklog(); renderStalledGuias(); renderFallidos(); }
  if (tab === 'tabla')           renderMainTable();
  if (tab === 'incumplimientos') renderIncumplimientosTab();
  if (tab === 'rollos-main')     renderRollosTab();
  if (tab === 'rollos-detalle')  {
    if (ROLLOS_RAW) renderRollosDetalleTable();
    else setTimeout(() => { if (ROLLOS_RAW) renderRollosDetalleTable(); }, 2000);
  }
  if (tab === 'rollos-comercio') {
    // rollos_comercio_v2.js sobrescribe renderRollosComercioTable y maneja
    // internamente la espera de TABLERO_ROLLOS_FILAS — no necesita ROLLOS_RAW.
    if (typeof window.renderRollosComercioTable === 'function') window.renderRollosComercioTable();
    if (typeof window.renderRollosInvComercio === 'function') window.renderRollosInvComercio();
  }
  if (tab === 'rollos-inventario') { if (typeof window.renderRollosInventario === 'function') window.renderRollosInventario(); }
  if (tab === 'inv-principal')   renderInventarioPrincipal();
  if (tab === 'inv-detalles')    renderInventarioDetalles();
  if (tab === 'estado-materiales' && typeof window.renderEstadoMateriales === 'function') window.renderEstadoMateriales();
  if (tab === 'garantia-inventario' && typeof window.renderGarantiaInventario === 'function') window.renderGarantiaInventario();
  if (tab === 'puntos-reorden'      && typeof window.renderPuntosReorden      === 'function') window.renderPuntosReorden();
}

function _showAllPanels(activeTab) {
  const panelMap = {
    'tracking':          'panel-tracking',
    'detalle':           'panel-detalle',
    'tabla':             'panel-tabla',
    'incumplimientos':   'panel-incump',
    'rollos-main':       'panel-rollos',
    'rollos-detalle':    'panel-rollos-detalle',
    'rollos-comercio':   'panel-rollos-comercio',
    'rollos-inventario': 'panel-rollos-inventario',
    'inv-principal':     'panel-inv-principal',
    'inv-detalles':      'panel-inv-detalles',
    'estado-materiales': 'panel-estado-materiales',
    'garantia-inventario': 'panel-garantia-inventario',
    'puntos-reorden':      'panel-puntos-reorden',
  };
  // Hide all
  ['panel-home','panel-tracking','panel-detalle','panel-tabla','panel-incump',
   'panel-rollos','panel-rollos-detalle','panel-rollos-comercio','panel-rollos-inventario','panel-inv-principal','panel-inv-detalles',
   'panel-estado-materiales','panel-garantia-inventario','panel-puntos-reorden'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.style.display = 'none';
  });
  // Show target or home
  if (activeTab && panelMap[activeTab]) {
    const el = document.getElementById(panelMap[activeTab]);
    if (el) el.style.display = 'block';
  } else {
    const home = document.getElementById('panel-home');
    if (home) home.style.display = 'flex';
  }
}

window._sidebarReady   = true;
window._toggleSidebar  = _toggleSidebar;
window._selectBoard    = _selectBoard;
window._selectBoardTab = _selectBoardTab;

// Wire toggle button — script loads after DOM so just grab it directly
(function() {
  const btn = document.getElementById('sidebar-toggle-btn');
  if (btn) btn.addEventListener('click', _toggleSidebar);
  // Show welcome panel on load
  _showAllPanels(null);
})();

// (sidebar selectBoardTab handled by _selectBoardTab above)

// ══════════════════════════════════════════════════════════════════
//  DATA LOADING OVERLAY
// ══════════════════════════════════════════════════════════════════
let _mainLoaded       = false;
let _rollosLoaded     = false;
let _inventarioLoaded = false;

// Exponer para que inventario.js lo llame cuando termine su carga
window._setInventarioLoaded = function() {
  _inventarioLoaded = true;
  _updateLoadingUI();
};

function _updateLoadingUI() {
  const mainDone  = _mainLoaded;
  const rollosDone= _rollosLoaded;
  const invDone   = _inventarioLoaded;

  const dotMain   = document.getElementById('dl-dot-main');
  const dotRollos = document.getElementById('dl-dot-rollos');
  const dotInv    = document.getElementById('dl-dot-inventario');
  const msg       = document.getElementById('dl-msg');
  const fill      = document.getElementById('dl-progress-fill');

  if (dotMain)   dotMain.className   = 'dl-item-dot ' + (mainDone   ? 'done' : 'loading');
  if (dotRollos) dotRollos.className = 'dl-item-dot ' + (rollosDone ? 'done' : 'loading');
  if (dotInv)    dotInv.className    = 'dl-item-dot ' + (invDone    ? 'done' : 'loading');

  const loaded = [mainDone, rollosDone, invDone].filter(Boolean).length;
  if (fill) fill.style.width = Math.round(loaded / 3 * 100) + '%';

  const allDone = mainDone && rollosDone && invDone;
  if (!mainDone) {
    if (msg) msg.textContent = 'Cargando datos principales...';
  } else if (!rollosDone) {
    if (msg) msg.textContent = 'data.json ✓ · Descomprimiendo rollos...';
  } else if (!invDone) {
    if (msg) msg.textContent = 'Rollos ✓ · Cargando inventario Wompi...';
  } else {
    if (msg) msg.textContent = '¡Listo! Iniciando dashboard...';
    setTimeout(() => {
      const overlay = document.getElementById('data-loading-overlay');
      if (overlay) overlay.classList.add('hidden');
    }, 600);
  }
}

// ══════════════════════════════════════════════════════════════════
//  INIT
// ══════════════════════════════════════════════════════════════════
function initDashboard() {
  // Mark both as loading
  const dotMain   = document.getElementById('dl-dot-main');
  const dotRollos = document.getElementById('dl-dot-rollos');
  const dotInv    = document.getElementById('dl-dot-inventario');
  if (dotMain)   dotMain.className   = 'dl-item-dot loading';
  if (dotRollos) dotRollos.className = 'dl-item-dot loading';
  if (dotInv)    dotInv.className    = 'dl-item-dot loading';

  loadData();
  loadRollosData();
  // inventario.js se carga en paralelo y llama window._setInventarioLoaded() al terminar
}

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
//  ROLLOS WOMPI — módulo independiente v2 (con filtros globales)
//  Carga data_rollos.json.gz (gzip → JSON), sin tocar nada de lo anterior
// ══════════════════════════════════════════════════════════════════

var ROLLOS_RAW       = null;   // payload completo del .json.gz
let ROLLOS_FILTERED  = [];     // detalle filtrado por filtros globales
let ROLLOS_DETALLE   = [];     // detalle filtrado por búsqueda de tabla
let ROLLOS_COMERCIO  = [];     // comercio filtrado
let rollosDetallePage  = 1;
let rollosComercioPage = 1;
const ROLLOS_PAGE_SIZE = 15;

// ── Carga y descompresión ─────────────────────────────────────────
async function loadRollosData() {
  try {
    const res = await fetch(`data_rollos.json.gz?t=${Date.now()}`);
    if (!res.ok) throw new Error('no file');
    const buf  = await res.arrayBuffer();
    const ds   = new DecompressionStream('gzip');
    const writer = ds.writable.getWriter();
    writer.write(new Uint8Array(buf));
    writer.close();
    const out  = await new Response(ds.readable).arrayBuffer();
    const text = new TextDecoder().decode(out);
    ROLLOS_RAW = JSON.parse(text);

    // ── EXCLUSIÓN PERMANENTE REDEBAN (Requerimiento) ──────────────
    if (ROLLOS_RAW.detalle) {
      // Exclusión Redeban
      ROLLOS_RAW.detalle = ROLLOS_RAW.detalle.filter(r => 
        !(r.proyecto || '').toUpperCase().includes('REDEBAN')
      );
      // Transformación Tipo Flujo
      ROLLOS_RAW.detalle.forEach(r => {
        if (r.tipo_flujo) {
          r.tipo_flujo = String(r.tipo_flujo).replace(/P-TA-/gi, '').trim();
        }
      });
    }
    if (ROLLOS_RAW.comercio) {
      ROLLOS_RAW.comercio = ROLLOS_RAW.comercio.filter(r => 
        !(r.proyecto || '').toUpperCase().includes('REDEBAN')
      );
    }
    if (ROLLOS_RAW.calculos) {
      ROLLOS_RAW.calculos = ROLLOS_RAW.calculos.filter(r => 
        !(r.proyecto || '').toUpperCase().includes('REDEBAN')
      );
    }
    // ──────────────────────────────────────────────────────────────

    console.log('[Rollos] Cargado OK — detalle filtrado:', ROLLOS_RAW.detalle?.length, 'filas');
    initRollosGlobalFilters();
    applyRollosGlobalFilters();
    if (document.getElementById('tab-rollos')?.classList.contains('active')) renderRollosTab();
    _rollosLoaded = true;
    console.log('[Rollos] Llamando a initRollosInventario...');
    if (typeof window.initRollosInventario === 'function') window.initRollosInventario();
    else console.warn('[Rollos] window.initRollosInventario no es una función!');
    _updateLoadingUI();
  } catch(e) {
    console.warn('[Rollos] No se pudo cargar data_rollos.json.gz:', e.message);

    // ── Fallback: construir ROLLOS_RAW desde data_tablero_rollos.json.gz ──
    // data_rollos.json.gz fue descontinuado. Los datos equivalentes están en
    // TABLERO_ROLLOS_FILAS, cargado por rollos_inventario.js via _loadTablRollos().
    // Esperamos que esa carga termine (puede estar aún en vuelo) antes de continuar.
    console.log('[Rollos] Esperando data_tablero_rollos.json.gz...');
    if (typeof window.loadTablRollos === 'function') {
      await window.loadTablRollos();   // no-op si ya cargó; espera si está en vuelo
    } else {
      // rollos_inventario.js aún no parseado — esperar hasta 30 s
      await new Promise(resolve => {
        const t0 = Date.now();
        const poll = () => {
          if (typeof window.loadTablRollos === 'function' || Date.now() - t0 > 30000)
            return resolve();
          setTimeout(poll, 300);
        };
        poll();
      });
      if (typeof window.loadTablRollos === 'function') await window.loadTablRollos();
    }

    const filas = window.TABLERO_ROLLOS_FILAS || [];

    if (filas.length > 0) {
      // Derivar calculos: una fila por sitio único (tomando la primera con cal_saldo_dias != 0)
      const calcMap = new Map();
      const detalle = [];
      filas.forEach(f => {
        const codSitio = (f.codigo_sitio || f.cal_codigo_sitio || '').trim();
        const codMO    = (f.cal_codigo_mo || '').trim();
        const key      = codSitio || codMO;
        // Detalle: cada fila es un movimiento
        // Mapear campos de TABLERO_ROLLOS_FILAS al formato que espera computeRollosKPIs / DETALLE_COLS
        const _flujoLimpio = (f.flujo || '').replace(/P-TA-/gi, '').trim();
        const _estTarea    = (f.estado_tarea || '').toUpperCase();
        detalle.push({
          tarea                 : f.tarea || '',
          codigo_tarea          : f.tarea || '',           // alias para computeRollosKPIs
          cod_sitio             : codSitio,
          nombre_sitio          : f.nombre_sitio || '',
          departamento          : f.departamento || '',
          ciudad                : f.Ciudad || f.ciudad || '',
          proyecto              : f.proyecto || '',
          subproyecto           : f.subproyecto || '',
          tipo_flujo            : _flujoLimpio,
          flujo_raw             : f.flujo || '',
          codigo_material       : f.codigo_material || '',
          nombre_material       : f.nombre_material || '',
          Cantidad              : f.Cantidad || 0,
          cantidad              : parseFloat(f.Cantidad || 0),
          estado                : f.estado_tarea || '',    // para status pills
          estado_tarea          : f.estado_tarea || '',
          estado_transportadora : '',                      // no disponible en esta fuente
          // estado_ans inferido: Completada → CUMPLE (heurístico)
          estado_ans            : _estTarea === 'COMPLETADA' ? 'CUMPLE' : '',
          estado_resultado      : _estTarea === 'COMPLETADA' ? 'EXITOSO' : '',
          fecha_confirmacion    : f.fecha_confirmacion || '',
          fecha_entrega         : f.fecha_confirmacion || f.tarea_fecha_fin || '',
          fecha_entrega_raw     : f.tarea_fecha_fin || '',
          fecha_plan_inicio     : f.plan_inicio || '',
          fecha_plan_fin        : f.plan_fin || '',
          guia                  : f.guia || f.guia_raw || '',
          transportadora        : f.transportadora || '',
          nombre_ubicacion_origen  : f.nombre_ubicacion_origen || '',
          nombre_ubicacion_destino : f.nombre_ubicacion_destino || '',
          nombre_plantilla_tarea   : f.nombre_plantilla_tarea || '',
          codigo_operacion         : f.codigo_operacion || '',
          tipologia                : f.tipologia || '',
          red_asociada             : f.red_asociada || '',
          nit                      : f.nit || '',
          // campos cal_* del corresponsal para enriquecer la fila de tarea
          cal_codigo_mo            : (f.cal_codigo_mo || '').trim(),
          cal_saldo_rollos         : parseFloat(f.cal_saldo_rollos || 0),
          cal_saldo_dias           : parseFloat(f.cal_saldo_dias || 0),
          cal_punto_reorden        : parseFloat(f.cal_punto_reorden || 0),
          cal_prom_mensual         : parseFloat(f.cal_promedio_mensual || 0),
          cal_estado_punto         : f.cal_estado_punto || '',
          cal_fecha_abst           : f.cal_fecha_abst_1 || '',
          oportunidad           : '',
          FO                    : '',
          dias_inventario_restantes: '',
        });
        // Calculos: un registro por sitio con join real
        if (!calcMap.has(key) && parseFloat(f.cal_saldo_dias || 0) !== 0) {
          calcMap.set(key, {
            id                          : f.cal_id || '',
            tarea                       : f.tarea || '',
            codigo_mo                   : (f.cal_codigo_mo || '').trim(),
            codigo_sitio                : codSitio || codMO,
            estado_punto                : f.cal_estado_punto || '',
            promedio_mensual            : parseFloat(f.cal_promedio_mensual || 0),
            rollos_promedio_mes         : parseFloat(f.cal_rollos_promedio_mes || 0),
            periodo_abast_e5            : parseFloat(f.cal_periodo_abast_e5 || 0),
            valor_busqueda              : f.cal_valor_busqueda || '',
            rollos_periodo_abast_e5     : parseFloat(f.cal_rollos_periodo_abast_e5 || 0),
            rollos_anio_e5              : parseFloat(f.cal_rollos_anio_e5 || 0),
            punto_reorden               : parseFloat(f.cal_punto_reorden || 0),
            fecha_apertura_final        : f.cal_fecha_apertura_final || '',
            fecha_abst_1                : f.cal_fecha_abst_1 || '',
            rollos_entregados_mig_apert : parseFloat(f.cal_rollos_entregados_mig_apert || 0),
            trx_desde_migra_apert       : parseFloat(f.cal_trx_desde_migra_apert || 0),
            rollos_consumidos_migr_apert: parseFloat(f.cal_rollos_consumidos_migr_apert || 0),
            saldo_rollos                : parseFloat(f.cal_saldo_rollos || 0),
            saldo_dias                  : parseFloat(f.cal_saldo_dias || 0),
            saldo                       : parseFloat(f.cal_saldo || 0),
            // metadata para enriquecer
            nombre_sitio : f.nombre_sitio || '',
            departamento : f.departamento || '',
            ciudad       : f.Ciudad || f.ciudad || '',
            proyecto     : f.proyecto || '',
          });
        }
      });

      const calculos = [...calcMap.values()].filter(r =>
        !(r.proyecto || '').toUpperCase().includes('REDEBAN')
      );
      const detalleFiltered = detalle.filter(r =>
        !(r.proyecto || '').toUpperCase().includes('REDEBAN')
      );

      ROLLOS_RAW = { detalle: detalleFiltered, comercio: [], calculos };
      window.ROLLOS_RAW = ROLLOS_RAW;

      console.log('[Rollos] Fallback TABLERO_ROLLOS_FILAS → ROLLOS_RAW construido:',
        calculos.length, 'sitios en calculos,', detalleFiltered.length, 'filas en detalle');

      initRollosGlobalFilters();
      applyRollosGlobalFilters();
      if (document.getElementById('tab-rollos')?.classList.contains('active')) renderRollosTab();
      console.log('[Rollos] Llamando a initRollosInventario (fallback)...');
      if (typeof window.initRollosInventario === 'function') window.initRollosInventario();
      else console.warn('[Rollos] window.initRollosInventario no es una función!');
    } else {
      console.warn('[Rollos] TABLERO_ROLLOS_FILAS también vacío — tablero de rollos sin datos.');
    }

    _rollosLoaded = true;
    _updateLoadingUI();
  }
}

// ── Inicializar selects de filtros GLOBALES ───────────────────────
function initRollosGlobalFilters() {
  // Los filtros del panel "Rollos Wompi" usan INV_RAW (inventario)
  // para reflejar los mismos datos que muestra la sección de KPIs de stock
  const invRaw = window.INV_RAW;
  const getCat = window.invCategoria || (n => ((n||'').toUpperCase().includes('ROLLO') ? 'Rollos' : 'Otro'));

  if (invRaw && invRaw.length) {
    const soloRollos = invRaw.filter(r => getCat(r['Nombre']) === 'Rollos');

    const uniqInv = (key) => [...new Set(soloRollos.map(r => (r[key]||'').trim()).filter(Boolean))].sort();

    const populate = (id, vals) => {
      const el = document.getElementById(id);
      if (!el) return;
      if (el.classList.contains('ms-container')) {
        window._setupMS(id, vals);
      } else {
        el.innerHTML = '<option value="">Todos</option>' + vals.map(v=>`<option value="${v}">${v}</option>`).join('');
      }
    };

    // Reutilizar los slots de filtros existentes con valores del inventario
    populate('rg-estado',      uniqInv('Tipo de ubicación'));      // Tipo ubic como "estado"
    populate('rg-departamento', []);                                // No aplica en inv
    populate('rg-tipo-flujo',  []);                                 // No aplica en inv
    populate('rg-material',    uniqInv('Nombre'));                  // Referencia de rollo
    populate('rg-proyecto',    uniqInv('Nombre de la ubicación')); // Bodega/ubicación

    // Actualizar labels de los filtros para que tengan sentido con inventario
    const labelMap = {
      'rg-estado':      'Tipo Ubicación',
      'rg-departamento': null,   // ocultar
      'rg-tipo-flujo':  null,    // ocultar
      'rg-material':    'Referencia',
      'rg-proyecto':    'Ubicación',
    };
    Object.entries(labelMap).forEach(([id, labelText]) => {
      const group = document.getElementById(id)?.closest('.filter-group');
      if (!group) return;
      if (labelText === null) {
        group.style.display = 'none';
      } else {
        group.style.display = '';
        const lbl = group.querySelector('label');
        if (lbl) lbl.textContent = labelText;
      }
    });

    // Ocultar el filtro de fecha (no aplica a inventario estático)
    const dateGroup = document.querySelector('.filter-group--date-range');
    if (dateGroup) dateGroup.style.display = 'none';

  } else if (ROLLOS_RAW) {
    // Fallback: si inventario aún no cargó, usar ROLLOS_RAW (comportamiento anterior)
    const det = ROLLOS_RAW.detalle || [];
    const uniq = (key) => [...new Set(det.map(r => r[key]).filter(Boolean))].sort();
    const populate = (id, vals) => {
      const el = document.getElementById(id);
      if (!el) return;
      if (el.classList.contains('ms-container')) window._setupMS(id, vals);
      else el.innerHTML = '<option value="">Todos</option>' + vals.map(v=>`<option value="${v}">${v}</option>`).join('');
    };
    populate('rg-estado',      uniq('estado'));
    populate('rg-departamento',uniq('departamento'));
    populate('rg-tipo-flujo',  uniq('tipo_flujo'));
    populate('rg-material',    uniq('material'));
    populate('rg-proyecto',    uniq('proyecto'));
  }

  if (!ROLLOS_RAW) return;
  const det = ROLLOS_RAW.detalle || [];
  const uniq = (key) => [...new Set(det.map(r => r[key]).filter(Boolean))].sort();
  const populate = (id, vals) => {
    const el = document.getElementById(id);
    if (!el) return;
    if (el.classList.contains('ms-container')) window._setupMS(id, vals);
    else el.innerHTML = '<option value="">Todos</option>' + vals.map(v=>`<option value="${v}">${v}</option>`).join('');
  };

  // Filtros tabla detalle (internos, sin cambios)
  const anios     = [...new Set(det.map(r=>r.anio).filter(Boolean))].sort().reverse();
  const meses     = [...new Set(det.map(r=>r.mes).filter(Boolean))].sort();
  const tiposFlujo= uniq('tipo_flujo');
  const deptos    = uniq('departamento');
  const ciudades  = uniq('ciudad');
  const proyectos = uniq('proyecto');
  const estados   = uniq('estado');

  const materiales    = uniq('nombre_material');
  const plantillas    = uniq('nombre_plantilla_tarea');
  const estadosPunto  = uniq('cal_estado_punto');

  populate('rf-estado',           estados);
  populate('rf-tipo-flujo-det',   tiposFlujo);
  populate('rf-departamento-det', deptos);
  populate('rf-ciudad-det',       ciudades);
  populate('rf-proyecto-det',     proyectos);
  populate('rf-anio',             anios.map(String));
  populate('rf-mes',              meses.map(m => String(m).padStart(2,'0')));
  populate('rf-nombre-material-det', materiales);
  populate('rf-plantilla-tarea-det', plantillas);
  populate('rf-estado-punto-det',    estadosPunto);

  // Filtros tabla comercio
  const comercio = ROLLOS_RAW.comercio || [];
  const comEst   = [...new Set(comercio.map(r=>r.estado).filter(Boolean))].sort();
  const comTipo  = [...new Set(comercio.map(r=>r.tipo_envio).filter(Boolean))].sort();
  const comDeptos= [...new Set(comercio.map(r=>r.departamento).filter(Boolean))].sort();
  populate('rf-com-estado',       comEst);
  populate('rf-com-tipo',         comTipo);
  populate('rf-com-departamento', comDeptos);

  // Inicializar referencias filtradas
  ROLLOS_REF_FILTERED = (ROLLOS_RAW.referencias || []).slice();
}

// ── Aplicar filtros GLOBALES — recalcula TODO ─────────────────────
function applyRollosGlobalFilters() {
  if (!ROLLOS_RAW && !window.INV_RAW) return;

  // ── Filtros de inventario (afectan KPIs de stock) ─────────────
  const selTipoUbic  = window._msGetSels('rg-estado');    // mapeado a Tipo Ubicación
  const selReferencia= window._msGetSels('rg-material');  // mapeado a Nombre (referencia)
  const selUbicacion = window._msGetSels('rg-proyecto');  // mapeado a Nombre de la ubicación

  const hayFiltroInv = selTipoUbic || selReferencia || selUbicacion;

  if (window.INV_RAW && hayFiltroInv) {
    // Filtrar INV_RAW con los filtros de inventario y re-renderizar KPIs
    const getCat = window.invCategoria || (n => ((n||'').toUpperCase().includes('ROLLO') ? 'Rollos' : 'Otro'));
    window._invRawOverride = window.INV_RAW.filter(r => {
      if (getCat(r['Nombre']) !== 'Rollos') return false;
      if (selTipoUbic   && !selTipoUbic.includes((r['Tipo de ubicación']||'').trim().toUpperCase()))    return false;
      if (selReferencia && !selReferencia.includes((r['Nombre']||'').trim().toUpperCase()))               return false;
      if (selUbicacion  && !selUbicacion.includes((r['Nombre de la ubicación']||'').trim().toUpperCase())) return false;
      return true;
    });
    // Temporalmente sobreescribir INV_RAW para re-render (restaurar después)
    const origInvRaw = window.INV_RAW;
    window.INV_RAW = window._invRawOverride;
    if (typeof window.renderRollosInvBodegaKPIs === 'function') window.renderRollosInvBodegaKPIs();
    window.INV_RAW = origInvRaw;
  } else if (!hayFiltroInv && typeof window.renderRollosInvBodegaKPIs === 'function') {
    // Sin filtros: re-renderizar con datos completos
    window.renderRollosInvBodegaKPIs();
  }

  // ── Filtros ROLLOS_RAW (gráficas y tablas de tracking) ──────
  if (!ROLLOS_RAW) return;
  const det = ROLLOS_RAW.detalle || [];

  // Para las gráficas de tracking seguimos usando los campos de ROLLOS_RAW
  // (los filtros visibles en el panel son ahora de inventario, pero las gráficas
  //  de tendencias/estados se filtran sin restricción por defecto)
  ROLLOS_FILTERED = det.slice(); // sin filtro extra sobre rollos_raw

  // Actualizar sumario de filtros
  const activos = [];
  if (selTipoUbic)   activos.push(`Tipo Ubic: ${selTipoUbic.length > 1 ? selTipoUbic.length+' items' : selTipoUbic[0]}`);
  if (selReferencia) activos.push(`Ref: ${selReferencia.length > 1 ? selReferencia.length+' items' : selReferencia[0]}`);
  if (selUbicacion)  activos.push(`Ubic: ${selUbicacion.length > 1 ? selUbicacion.length+' items' : selUbicacion[0]}`);
  
  const summary = document.getElementById('rg-filter-summary');
  if (summary) {
    summary.textContent = activos.length
      ? `🔍 Filtros activos (Inventario): ${activos.join('  ·  ')}`
      : `Sin filtros activos — mostrando datos completos del inventario`;
  }

  ROLLOS_COMERCIO = (ROLLOS_RAW.comercio || []).slice();
  ROLLOS_DETALLE = ROLLOS_FILTERED.slice();
  ROLLOS_REF_FILTERED = (ROLLOS_RAW?.referencias || []).slice();
  rollosDetallePage  = 1;
  rollosComercioPage = 1;
  rollosRefPage      = 1;

  renderRollosKPIs();
  renderRollosKPIsTareas();
  renderRollosANSRow();
  renderRollosCharts();
  renderRollosDetalleTable();
  renderRollosComercioTable();
  renderRollosRefTable();
}

function resetRollosGlobalFilters() {
  ['rg-fecha-desde','rg-fecha-hasta'].forEach(id => { const el=document.getElementById(id); if(el) el.value=''; });
  ['rg-estado','rg-departamento','rg-tipo-flujo','rg-material','rg-proyecto'].forEach(id => {
    const el = document.getElementById(id);
    if (el && el.classList.contains('ms-container')) window._msAction(id, 'clear');
    else if (el) el.value = '';
  });
  // Restaurar KPIs de inventario sin filtros
  window._invRawOverride = null;
  if (typeof window.renderRollosInvBodegaKPIs === 'function') window.renderRollosInvBodegaKPIs();
  applyRollosGlobalFilters();
}

// ── Calcular KPIs desde ROLLOS_FILTERED ──────────────────────────
function computeRollosKPIs() {
  const det   = ROLLOS_FILTERED;
  const kRaw  = ROLLOS_RAW?.kpis || {};

  // Conteos por estado (desde detalle filtrado)
  let rollos_alistamiento = 0, tareas_alistamiento = 0;
  let rollos_transito = 0, tareas_transito = 0;
  let rollos_entregados = 0, tareas_entregados = 0;
  let rollos_devueltos = 0, tareas_devueltos = 0;
  let total_rollos = 0, total_tareas = 0;

  const tareasSet = new Set();
  det.forEach(r => {
    const est = (r.estado||'').toUpperCase();
    const qty = parseFloat(r.cantidad) || 0;
    const tarea = r.codigo_tarea;
    total_rollos += qty;
    if (tarea && !tareasSet.has(tarea)) { tareasSet.add(tarea); total_tareas++; }

    // Mapeo de estados reales del campo `estado` en data_rollos.json:
    //   "Abierta"                   → En Alistamiento (pendiente de despacho)
    //   "En proceso"                → En Tránsito (en camino)
    //   "Completada"                → Entregado
    //   "Completada con pendientes" → Entregado (con observaciones)
    // Las devoluciones se detectan via estado_transportadora
    const estTransp = (r.estado_transportadora||'').toUpperCase();
    const esDev = estTransp.includes('DEVOLUC') || estTransp.includes('DEVUELTO') || estTransp.includes('REMITENTE');
    if (esDev)                                   { rollos_devueltos += qty; tareas_devueltos++; }
    else if (est === 'ABIERTA')                  { rollos_alistamiento += qty; tareas_alistamiento++; }
    else if (est === 'EN PROCESO')               { rollos_transito += qty; tareas_transito++; }
    else if (est === 'COMPLETADA' || est === 'COMPLETADA CON PENDIENTES') { rollos_entregados += qty; tareas_entregados++; }
  });

  const pct = (n, d) => d > 0 ? Math.round(n / d * 100) : 0;

  // ── ANS Oportunidad (replica DAX: ESTADO_ANS_ROLLOS_SIMPLE) ─────
  // DAX: % Oportunidad = Rollos Entregados (ANS_SI) / Total_Tareas_Solicitadas
  // = tareas con estado_ans==="CUMPLE" / total_tareas (DISTINCTCOUNT codigo_tarea)
  // Usamos tareas únicas (Map) para respetar el DISTINCTCOUNT del DAX
  const tareasANSMap = new Map(); // tarea -> estado_ans
  det.forEach(r => {
    const tarea = r.codigo_tarea; if (!tarea) return;
    if (!tareasANSMap.has(tarea)) tareasANSMap.set(tarea, (r.estado_ans || '').toUpperCase());
  });
  let sla_cumple = 0, sla_nc = 0;
  tareasANSMap.forEach(ans => {
    if (ans === 'CUMPLE')    sla_cumple++;
    else if (ans === 'INCUMPLE') sla_nc++;
  });
  const sla_total = sla_cumple + sla_nc;
  // % Oportunidad = CUMPLE / Total_Tareas_Solicitadas (no solo las evaluadas)
  const pct_sla = pct(sla_cumple, total_tareas);

  // ── Tareas únicas por estado ──────────────────────────────────────
  const tareasEstadoMap = new Map();
  det.forEach(r => {
    const tarea = r.codigo_tarea; if (!tarea) return;
    if (!tareasEstadoMap.has(tarea)) tareasEstadoMap.set(tarea, { est: (r.estado||'').toUpperCase(), transp: (r.estado_transportadora||'').toUpperCase() });
  });
  let t_alistamiento = 0, t_transito = 0, t_entregados = 0, t_devueltos = 0;
  tareasEstadoMap.forEach(({est, transp}) => {
    const esDev = transp.includes('DEVOLUC') || transp.includes('DEVUELTO') || transp.includes('REMITENTE');
    if (esDev)                                                         t_devueltos++;
    else if (est === 'ABIERTA')                                        t_alistamiento++;
    else if (est === 'EN PROCESO')                                     t_transito++;
    else if (est === 'COMPLETADA' || est === 'COMPLETADA CON PENDIENTES') t_entregados++;
  });
  const t_total = tareasEstadoMap.size;
  const pct_t_entrega    = pct(t_entregados, t_total);
  const pct_t_alistam    = pct(t_alistamiento, t_total);
  const pct_t_transito   = pct(t_transito, t_total);
  const pct_t_devolucion = pct(t_devueltos, t_total);

  // Calidad: exitoso vs cancelado (informativo, para subtítulos)
  const cal_exitoso   = det.filter(r => (r.estado_resultado||'').toUpperCase().includes('EXITOSO')).length;
  const cal_cancelado = det.filter(r => { const e=(r.estado_resultado||'').toUpperCase(); return e.includes('CANCELADO'); }).length;
  const cal_pendiente = det.filter(r => { const e=(r.estado_resultado||'').toUpperCase().trim(); return !e || (!e.includes('EXITOSO') && !e.includes('CANCELADO')); }).length;
  const cal_total     = cal_exitoso + cal_cancelado;
  // % Calidad ANS (DAX exacto): 1 - % Devoluciones
  const pct_calidad = Math.max(0, 100 - pct(tareas_devueltos, total_tareas));

  // Calidad por tareas únicas
  const tareasCalMap = new Map();
  det.forEach(r => {
    const tarea = r.codigo_tarea; if (!tarea) return;
    if (tareasCalMap.has(tarea)) return;
    tareasCalMap.set(tarea, (r.estado_resultado||'').toUpperCase().trim());
  });
  let cal_t_exitoso = 0, cal_t_cancelado = 0, cal_t_pendiente = 0;
  tareasCalMap.forEach(est => {
    if (est.includes('EXITOSO'))      cal_t_exitoso++;
    else if (est.includes('CANCELADO')) cal_t_cancelado++;
    else                                cal_t_pendiente++;
  });
  const cal_t_total   = cal_t_exitoso + cal_t_cancelado;
  const pct_cal_tareas = pct(cal_t_exitoso, cal_t_total);

  return {
    rollos_alistamiento: Math.round(rollos_alistamiento), tareas_alistamiento,
    rollos_transito: Math.round(rollos_transito), tareas_transito,
    rollos_entregados: Math.round(rollos_entregados), tareas_entregados,
    rollos_devueltos: Math.round(rollos_devueltos), tareas_devueltos,
    total_rollos: Math.round(total_rollos), total_tareas,
    pct_entrega: pct(rollos_entregados, total_rollos),
    pct_alistamiento: pct(rollos_alistamiento, total_rollos),
    pct_transito: pct(rollos_transito, total_rollos),
    pct_devolucion: pct(rollos_devueltos, total_rollos),
    sla_cumple, sla_nc, sla_total, pct_sla,
    t_alistamiento, t_transito, t_entregados, t_devueltos, t_total,
    pct_t_entrega, pct_t_alistam, pct_t_transito, pct_t_devolucion,
    cal_exitoso, cal_cancelado, cal_pendiente, cal_total, pct_calidad,
    cal_t_exitoso, cal_t_cancelado, cal_t_pendiente, cal_t_total, pct_cal_tareas,
    // % Calidad KPI Rollos = datáfonos devueltos / datáfonos entregados
    pct_calidad_nueva: pct(rollos_devueltos, rollos_entregados),
  };
}

// ── Render KPIs ───────────────────────────────────────────────────
function renderRollosKPIs() {
  const k    = computeRollosKPIs();
  const grid = document.getElementById('rollos-kpi-grid');
  if (!grid) return;

  const cards = [
    { label:'Total Rollos',        value: k.total_rollos.toLocaleString('es-CO'), icon:'📦', color:'green',
      sub: `${k.total_tareas} tareas` },
    { label:'Rollos Entregados',   value: k.rollos_entregados.toLocaleString('es-CO'), icon:'✅', color:'selva',
      sub: `${k.pct_entrega}% del total · ${k.tareas_entregados} tareas`, pct: k.pct_entrega, pctColor:'green' },
    { label:'En Alistamiento',     value: k.rollos_alistamiento.toLocaleString('es-CO'), icon:'🔧', color:'lime',
      sub: `${k.pct_alistamiento}% del total · ${k.tareas_alistamiento} tareas`, pct: k.pct_alistamiento, pctColor:'lime' },
    { label:'En Tránsito',         value: k.rollos_transito.toLocaleString('es-CO'), icon:'🚚', color:'blue',
      sub: `${k.pct_transito}% del total · ${k.tareas_transito} tareas`, pct: k.pct_transito, pctColor:'blue' },
    { label:'Devoluciones',        value: k.rollos_devueltos.toLocaleString('es-CO'), icon:'↩️', color:'danger',
      sub: `${k.pct_devolucion}% del total · ${k.tareas_devueltos} tareas`, pct: k.pct_devolucion, pctColor:'danger' },
    { label:'ANS Oportunidad',     value: k.pct_sla + '%', icon:'🎯', color: k.pct_sla >= 80 ? 'selva' : k.pct_sla >= 60 ? 'warn' : 'danger',
      sub: `${k.sla_cumple} cumple / ${k.sla_total} evaluados`, pct: k.pct_sla, pctColor: k.pct_sla >= 80 ? 'green' : 'warn' },
    { label:'% Calidad',           value: (100 - k.pct_calidad_nueva) + '%', icon:'💎', color:'blue',
      sub: `${k.rollos_devueltos.toLocaleString('es-CO')} devueltos / ${k.rollos_entregados.toLocaleString('es-CO')} entregados`, pct: Math.max(0, 100 - k.pct_calidad_nueva), pctColor:'blue' },
  ];

  grid.innerHTML = cards.map((c, i) => `
    <div class="rkpi-card ${c.color} fade-up" style="animation-delay:${i*.05}s">
      <span class="rkpi-icon">${c.icon}</span>
      <div class="rkpi-label">${c.label}</div>
      <div class="rkpi-value" style="color:var(--${c.color === 'green' ? 'verde-menta' : c.color === 'selva' ? 'verde-selva' : c.color === 'lime' ? 'verde-lima' : c.color === 'blue' ? 'azul-cielo' : c.color === 'danger' ? 'danger' : c.color === 'warn' ? 'warning' : 'blanco'})">${c.value}</div>
      <div class="rkpi-sub">${c.sub}</div>
      ${c.pct !== undefined ? `
      <div class="rkpi-pct-row">
        <div class="rkpi-pct-bar"><div class="rkpi-pct-fill ${c.pctColor||c.color}" style="width:${Math.min(c.pct,100)}%"></div></div>
        <span class="rkpi-pct-label" style="color:var(--${c.pctColor === 'green' ? 'verde-menta' : c.pctColor === 'lime' ? 'verde-lima' : c.pctColor === 'blue' ? 'azul-cielo' : c.pctColor === 'danger' ? 'danger' : 'warning'})">${c.pct}%</span>
      </div>` : ''}
    </div>`).join('');
}

// ── KPIs por Tareas ───────────────────────────────────────────────
function renderRollosKPIsTareas() {
  const k    = computeRollosKPIs();
  const grid = document.getElementById('rollos-kpi-tareas-grid');
  if (!grid) return;

  const colorVal = c => c === 'green' ? 'verde-menta' : c === 'selva' ? 'verde-selva' : c === 'lime' ? 'verde-lima' : c === 'blue' ? 'azul-cielo' : c === 'danger' ? 'danger' : c === 'warn' ? 'warning' : 'blanco';
  const colorBar = c => c === 'green' ? 'verde-menta' : c === 'lime' ? 'verde-lima' : c === 'blue' ? 'azul-cielo' : c === 'danger' ? 'danger' : 'warning';

  const cards = [
    { label:'Total Tareas',          value: k.t_total.toLocaleString('es-CO'), icon:'📋', color:'green',
      sub: `universo de tareas únicas` },
    { label:'Tareas Entregadas',     value: k.t_entregados.toLocaleString('es-CO'), icon:'✅', color:'selva',
      sub: `${k.pct_t_entrega}% del total`, pct: k.pct_t_entrega, pctColor:'green' },
    { label:'En Alistamiento',       value: k.t_alistamiento.toLocaleString('es-CO'), icon:'🔧', color:'lime',
      sub: `${k.pct_t_alistam}% del total`, pct: k.pct_t_alistam, pctColor:'lime' },
    { label:'En Tránsito',           value: k.t_transito.toLocaleString('es-CO'), icon:'🚚', color:'blue',
      sub: `${k.pct_t_transito}% del total`, pct: k.pct_t_transito, pctColor:'blue' },
    { label:'Devoluciones',          value: k.t_devueltos.toLocaleString('es-CO'), icon:'↩️', color:'danger',
      sub: `${k.pct_t_devolucion}% del total`, pct: k.pct_t_devolucion, pctColor:'danger' },
    { label:'% Entrega Tareas',      value: k.pct_t_entrega + '%', icon:'📊', color:'green',
      sub: `${k.t_entregados} / ${k.t_total} tareas`, pct: k.pct_t_entrega, pctColor:'green' },
    { label:'ANS por Fechas',        value: k.pct_sla + '%', icon:'🗓️', color: k.pct_sla >= 80 ? 'selva' : k.pct_sla >= 60 ? 'warn' : 'danger',
      sub: `${k.sla_cumple} cumple / ${k.sla_total} evaluados`, pct: k.pct_sla, pctColor: k.pct_sla >= 80 ? 'green' : 'warn' },
    { label:'% Calidad Tareas',      value: k.pct_cal_tareas + '%', icon:'💎', color:'blue',
      sub: `${k.cal_t_exitoso} exitosos / ${k.cal_t_total} evaluados (${k.cal_t_pendiente} pendientes)`, pct: k.pct_cal_tareas, pctColor:'blue' },
  ];

  grid.innerHTML = cards.map((c, i) => `
    <div class="rkpi-card ${c.color} fade-up" style="animation-delay:${i*.05}s">
      <span class="rkpi-icon">${c.icon}</span>
      <div class="rkpi-label">${c.label}</div>
      <div class="rkpi-value" style="color:var(--${colorVal(c.color)})">${c.value}</div>
      <div class="rkpi-sub">${c.sub}</div>
      ${c.pct !== undefined ? `
      <div class="rkpi-pct-row">
        <div class="rkpi-pct-bar"><div class="rkpi-pct-fill ${c.pctColor||c.color}" style="width:${Math.min(c.pct,100)}%"></div></div>
        <span class="rkpi-pct-label" style="color:var(--${colorBar(c.pctColor||c.color)})">${c.pct}%</span>
      </div>` : ''}
    </div>`).join('');
}

// ── ANS Big Row ────────────────────────────────────────────────────
function renderRollosANSRow() {
  const k   = computeRollosKPIs();
  const row = document.getElementById('rollos-ans-row');
  if (!row) return;

  const pctNc  = 100 - k.pct_sla;
  const pctDev = k.pct_devolucion;
  const pctOk  = k.pct_entrega;

  row.innerHTML = `
    <div class="ans-big-card" style="border-top:3px solid var(--verde-menta)">
      <canvas id="ans-mini-entrega" width="120" height="120" style="width:120px;height:120px;margin-bottom:12px"></canvas>
      <div class="ans-big-pct" style="color:var(--verde-menta)">${k.pct_entrega}%</div>
      <div class="ans-big-label">Tasa de Entrega</div>
      <div class="ans-detail-row">
        <div class="ans-detail-item"><div class="ans-detail-val" style="color:var(--verde-menta)">${k.rollos_entregados.toLocaleString()}</div><div class="ans-detail-lbl">Entregados</div></div>
        <div class="ans-detail-item"><div class="ans-detail-val" style="color:var(--muted)">${k.total_rollos.toLocaleString()}</div><div class="ans-detail-lbl">Total</div></div>
      </div>
    </div>
    <div class="ans-big-card" style="border-top:3px solid var(--azul-cielo)">
      <canvas id="ans-mini-sla" width="120" height="120" style="width:120px;height:120px;margin-bottom:12px"></canvas>
      <div class="ans-big-pct" style="color:var(--azul-cielo)">${k.pct_sla}%</div>
      <div class="ans-big-label">Cumplimiento ANS</div>
      <div class="ans-detail-row">
        <div class="ans-detail-item"><div class="ans-detail-val" style="color:var(--verde-menta)">${k.sla_cumple}</div><div class="ans-detail-lbl">Cumple</div></div>
        <div class="ans-detail-item"><div class="ans-detail-val" style="color:var(--danger)">${k.sla_nc}</div><div class="ans-detail-lbl">No Cumple</div></div>
        <div class="ans-detail-item"><div class="ans-detail-val" style="color:var(--muted)">${k.sla_total}</div><div class="ans-detail-lbl">Total eval.</div></div>
      </div>
    </div>
    <div class="ans-big-card" style="border-top:3px solid ${k.pct_calidad_nueva > 10 ? 'var(--danger)' : k.pct_calidad_nueva > 5 ? 'var(--warning)' : 'var(--verde-menta)'}">
      <canvas id="ans-mini-devol" width="120" height="120" style="width:120px;height:120px;margin-bottom:12px"></canvas>
      <div class="ans-big-pct" style="color:${k.pct_calidad_nueva > 10 ? 'var(--danger)' : k.pct_calidad_nueva > 5 ? 'var(--warning)' : 'var(--verde-menta)'}">${100 - k.pct_calidad_nueva}%</div>
      <div class="ans-big-label">% Calidad</div>
      <div class="ans-detail-row">
        <div class="ans-detail-item"><div class="ans-detail-val" style="color:var(--danger)">${k.rollos_devueltos.toLocaleString()}</div><div class="ans-detail-lbl">Devueltos</div></div>
        <div class="ans-detail-item"><div class="ans-detail-val" style="color:var(--verde-menta)">${k.rollos_entregados.toLocaleString()}</div><div class="ans-detail-lbl">Entregados</div></div>
      </div>
    </div>`;

  // Dibujar mini donuts (después de insertar el DOM)
  requestAnimationFrame(() => {
    _miniDonut('ans-mini-entrega', k.pct_entrega, '#B0F2AE', '#1a2a1a');
    _miniDonut('ans-mini-sla',     k.pct_sla,     '#99D1FC', '#1a1f2a');
    _miniDonut('ans-mini-devol',   100 - k.pct_calidad_nueva,
      k.pct_calidad_nueva <= 5 ? '#B0F2AE' : k.pct_calidad_nueva <= 10 ? '#FFC04D' : '#FF5C5C', '#1a1a1a');
  });
}

function _miniDonut(id, pct, color, bg) {
  const canvas = document.getElementById(id);
  if (!canvas) return;
  const ctx = canvas.getContext('2d');
  const cx = 60, cy = 60, r = 48, lw = 10;
  ctx.clearRect(0, 0, 120, 120);
  // background arc
  ctx.beginPath(); ctx.arc(cx, cy, r, 0, Math.PI*2);
  ctx.strokeStyle = bg || 'rgba(255,255,255,.06)'; ctx.lineWidth = lw;
  ctx.stroke();
  // value arc
  const angle = (pct / 100) * Math.PI * 2 - Math.PI / 2;
  ctx.beginPath(); ctx.arc(cx, cy, r, -Math.PI/2, angle);
  ctx.strokeStyle = color; ctx.lineWidth = lw;
  ctx.lineCap = 'round'; ctx.stroke();
}

// ── Render todas las gráficas ─────────────────────────────────────
function renderRollosCharts() {
  const det = ROLLOS_FILTERED;

  _destroyChart('rollos-estados');
  _destroyChart('rollos-tipo-flujo');
  _destroyChart('rollos-sla');
  _destroyChart('rollos-depto');
  _destroyChart('rollos-calidad');
  _destroyChart('rollos-mensual');
  _destroyChart('rollos-cumplimiento-mes');
  _destroyChart('rollos-oportunidad');
  _destroyChart('rollos-funnel');
  _destroyChart('rollos-t-mensual');

  // 1. DONA — Estados
  {
    const stateMap = {};
    det.forEach(r => {
      const e = r.estado || 'SIN ESTADO';
      stateMap[e] = (stateMap[e] || 0) + (parseFloat(r.cantidad) || 0);
    });
    const labels = Object.keys(stateMap);
    const data   = labels.map(k => Math.round(stateMap[k]));
    const colors = labels.map(l => {
      const u = l.toUpperCase();
      if (u === 'ENTREGADO')          return '#B0F2AE';
      if (u.includes('TRANSITO'))     return '#99D1FC';
      if (u.includes('ALISTAMIENTO')) return '#DFFF61';
      if (u.includes('DEVOLUC'))      return '#FF5C5C';
      return '#7B8CDE';
    });
    _buildDona('chart-rollos-estados', labels, data, colors, 'Rollos por estado');
  }

  // 2. DONA — Tipo flujo (solo rollos_TRIM, excluye VP)
  {
    const map = {};
    det.forEach(r => {
      const flujo = r.tipo_flujo || 'Sin tipo';
      if (flujo.toUpperCase().includes('VP')) return;
      map[flujo] = (map[flujo]||0) + (parseFloat(r.cantidad)||0);
    });
    const labels = Object.keys(map);
    const data   = labels.map(k => Math.round(map[k]));
    _buildDona('chart-rollos-tipo-flujo', labels, data, WOMPI_COLORS, 'Rollos por tipo flujo');
  }

  // 3. DONA — SLA
  {
    const k = computeRollosKPIs();
    _buildDona('chart-rollos-sla',
      ['Cumple ANS', 'No Cumple'],
      [k.sla_cumple, k.sla_nc],
      ['#B0F2AE', '#FF5C5C'],
      'ANS Oportunidad'
    );
  }

  // 4. BARRAS AGRUPADAS — Departamento (Rollos + Tareas)
  {
    const depRollos = {}, depTareas = {};
    const tareasSeenDep = new Set();
    det.forEach(r => {
      const d = r.departamento; if (!d) return;
      depRollos[d] = (depRollos[d] || 0) + (parseFloat(r.cantidad) || 0);
      const key = `${r.codigo_tarea}||${d}`;
      if (!tareasSeenDep.has(key)) { tareasSeenDep.add(key); depTareas[d] = (depTareas[d] || 0) + 1; }
    });
    const sorted = Object.entries(depRollos).sort((a,b)=>b[1]-a[1]).slice(0,15);
    const canvas = document.getElementById('chart-rollos-depto');
    if (canvas) {
      chartInstances['rollos-depto'] = new Chart(canvas, {
        type: 'bar',
        data: {
          labels: sorted.map(d => d[0]),
          datasets: [
            { label:'Rollos', data: sorted.map(d => Math.round(d[1])),
              backgroundColor:'rgba(176,242,174,.65)', borderColor:'#B0F2AE', borderWidth:1, borderRadius:5, yAxisID:'y' },
            { label:'Tareas', data: sorted.map(d => depTareas[d[0]] || 0),
              backgroundColor:'rgba(153,209,252,.55)', borderColor:'#99D1FC', borderWidth:1, borderRadius:5, yAxisID:'y1' },
          ]
        },
        options: {
          responsive: true, maintainAspectRatio: false, indexAxis: 'y',
          plugins: { legend:{ labels:{ color:'#FAFAFA', font:{size:11} } }, tooltip: CHART_OPTS.tooltip },
          scales: {
            x:  { ticks:{ color:'#7A7674' }, grid:{ color:'rgba(255,255,255,.05)' }, position:'bottom' },
            y:  { ticks:{ color:'#FAFAFA', font:{ size:11 } }, grid:{ display:false } },
            y1: { ticks:{ color:'#99D1FC', font:{size:10} }, grid:{ display:false }, position:'right' },
          },
        }
      });
    }
  }

  // 5. BARRAS — Top Proyectos (rollos por proyecto)
  {
    const map = {};
    det.forEach(r => { const p = r.proyecto || r.nombre_proyecto || 'Sin proyecto'; map[p] = (map[p]||0) + (parseFloat(r.cantidad)||0); });
    const sorted = Object.entries(map).sort((a,b)=>b[1]-a[1]).slice(0,8);
    const canvas = document.getElementById('chart-rollos-calidad');
    if (canvas) {
      chartInstances['rollos-calidad'] = new Chart(canvas, {
        type: 'bar',
        data: {
          labels: sorted.map(x => x[0].length > 20 ? x[0].substring(0,20)+'…' : x[0]),
          datasets: [{
            data: sorted.map(x => Math.round(x[1])),
            backgroundColor: sorted.map((_,i) => WOMPI_COLORS[i % WOMPI_COLORS.length] + 'BB'),
            borderColor:     sorted.map((_,i) => WOMPI_COLORS[i % WOMPI_COLORS.length]),
            borderWidth: 1, borderRadius: 8,
          }]
        },
        options: {
          responsive: true, maintainAspectRatio: false,
          plugins: { legend:{ display:false }, tooltip: CHART_OPTS.tooltip },
          scales: {
            x: { ticks:{ color:'#FAFAFA', font:{size:10}, maxRotation:35 }, grid:{ display:false } },
            y: { ticks:{ color:'#7A7674' }, grid:{ color:'rgba(255,255,255,.05)' } },
          }
        }
      });
    }
  }

  // 6. LÍNEA — Tendencia mensual
  {
    const mesMap = {};
    det.forEach(r => {
      const fp = String(r.fecha_plan_fin || '');
      if (fp.length < 7) return;
      const k = fp.substring(0,7);
      if (!mesMap[k]) mesMap[k] = { tareas:0, rollos:0 };
      mesMap[k].tareas++;
      mesMap[k].rollos += parseFloat(r.cantidad) || 0;
    });
    const periodos = Object.keys(mesMap).sort();
    const canvas = document.getElementById('chart-rollos-mensual');
    if (canvas && periodos.length) {
      chartInstances['rollos-mensual'] = new Chart(canvas, {
        type: 'bar',
        data: {
          labels: periodos,
          datasets: [
            { label:'Rollos', data: periodos.map(p=>Math.round(mesMap[p].rollos)),
              backgroundColor:'rgba(176,242,174,.5)', borderColor:'#B0F2AE', borderWidth:2, borderRadius:4, type:'bar', yAxisID:'y' },
            { label:'Tareas', data: periodos.map(p=>mesMap[p].tareas),
              borderColor:'#99D1FC', backgroundColor:'rgba(153,209,252,.15)', borderWidth:2, fill:true, tension:.4, type:'line', yAxisID:'y1' },
          ]
        },
        options: {
          responsive: true, maintainAspectRatio: false,
          plugins: { legend:{ labels:{ color:'#FAFAFA', font:{size:11} } }, tooltip: CHART_OPTS.tooltip },
          scales: {
            x: { ticks:{ color:'#7A7674', maxRotation:45 }, grid:{ display:false } },
            y:  { ticks:{ color:'#7A7674' }, grid:{ color:'rgba(255,255,255,.05)' }, position:'left' },
            y1: { ticks:{ color:'#99D1FC' }, grid:{ display:false }, position:'right' },
          }
        }
      });
    }
  }

  // 7. BARRAS APILADAS — Cumplimiento por mes
  {
    const mesMap = {};
    det.forEach(r => {
      const fp = String(r.fecha_plan_fin || '');
      if (fp.length < 7) return;
      const k = fp.substring(0,7);
      if (!mesMap[k]) mesMap[k] = { entregado:0, devuelto:0, otros:0 };
      const est = (r.estado||'').toUpperCase();
      if (est === 'ENTREGADO')              mesMap[k].entregado++;
      else if (est.includes('DEVOLUC'))     mesMap[k].devuelto++;
      else                                  mesMap[k].otros++;
    });
    const periodos = Object.keys(mesMap).sort();
    const canvas = document.getElementById('chart-rollos-cumplimiento-mes');
    if (canvas && periodos.length) {
      chartInstances['rollos-cumplimiento-mes'] = new Chart(canvas, {
        type: 'bar',
        data: {
          labels: periodos,
          datasets: [
            { label:'Entregado',    data: periodos.map(p=>mesMap[p].entregado),    backgroundColor:'rgba(176,242,174,.7)', borderRadius:3 },
            { label:'Devuelto',     data: periodos.map(p=>mesMap[p].devuelto),     backgroundColor:'rgba(255,92,92,.65)',  borderRadius:3 },
            { label:'En Proceso',   data: periodos.map(p=>mesMap[p].otros),        backgroundColor:'rgba(153,209,252,.45)',borderRadius:3 },
          ]
        },
        options: {
          responsive: true, maintainAspectRatio: false,
          plugins: { legend:{ labels:{ color:'#FAFAFA', font:{size:11} } }, tooltip: CHART_OPTS.tooltip },
          scales: {
            x: { stacked:true, ticks:{ color:'#7A7674', maxRotation:45 }, grid:{ display:false } },
            y: { stacked:true, ticks:{ color:'#7A7674' }, grid:{ color:'rgba(255,255,255,.05)' } },
          }
        }
      });
    }
  }

  // 8. LÍNEA — Devoluciones por mes
  {
    const mesMap = {};
    det.forEach(r => {
      const fp = String(r.fecha_plan_fin || '');
      if (fp.length < 7) return;
      const mk = fp.substring(0, 7);
      if (!mesMap[mk]) mesMap[mk] = { devueltos: 0, entregados: 0 };
      const est = (r.estado || '').toUpperCase();
      const qty = parseFloat(r.cantidad) || 0;
      if (est.includes('DEVOLUC'))  mesMap[mk].devueltos  += qty;
      else if (est === 'ENTREGADO') mesMap[mk].entregados += qty;
    });
    const periodos = Object.keys(mesMap).sort();
    const canvas = document.getElementById('chart-rollos-oportunidad');
    if (canvas && periodos.length) {
      chartInstances['rollos-oportunidad'] = new Chart(canvas, {
        type: 'bar',
        data: {
          labels: periodos,
          datasets: [
            { label:'Devueltos', data: periodos.map(p => Math.round(mesMap[p].devueltos)),
              backgroundColor:'rgba(255,92,92,.6)', borderColor:'#FF5C5C', borderWidth:1, borderRadius:4 },
            { label:'Entregados', data: periodos.map(p => Math.round(mesMap[p].entregados)),
              backgroundColor:'rgba(176,242,174,.4)', borderColor:'#B0F2AE', borderWidth:1, borderRadius:4 },
          ]
        },
        options: {
          responsive: true, maintainAspectRatio: false,
          plugins: { legend:{ labels:{ color:'#FAFAFA', font:{size:11} } }, tooltip: CHART_OPTS.tooltip },
          scales: {
            x: { ticks:{ color:'#7A7674', maxRotation:45, font:{size:10} }, grid:{ display:false } },
            y: { ticks:{ color:'#7A7674' }, grid:{ color:'rgba(255,255,255,.05)' } },
          }
        }
      });
    } else if (canvas) {
      _buildDona('chart-rollos-oportunidad', ['Sin datos'], [1], ['#3a3f4a'], 'Devoluciones');
    }
  }

  // 10. EMBUDO — Proceso
  {
    const k = computeRollosKPIs();
    const canvas = document.getElementById('chart-rollos-funnel');
    if (canvas) {
      const stages = [
        { label:'Solicitados', value: k.total_rollos,          color:'rgba(176,242,174,.6)' },
        { label:'Alistados',   value: k.total_rollos - k.rollos_alistamiento, color:'rgba(223,255,97,.6)' },
        { label:'En Tránsito', value: k.rollos_transito,       color:'rgba(153,209,252,.6)' },
        { label:'Entregados',  value: k.rollos_entregados,     color:'rgba(0,130,90,.7)' },
      ];
      chartInstances['rollos-funnel'] = new Chart(canvas, {
        type: 'bar',
        data: {
          labels: stages.map(s=>s.label),
          datasets: [{
            data: stages.map(s=>s.value),
            backgroundColor: stages.map(s=>s.color),
            borderColor: stages.map(s=>s.color.replace(',.6)',',1)').replace(',.7)',',1)')),
            borderWidth:1, borderRadius:8,
          }]
        },
        options: {
          responsive: true, maintainAspectRatio: false,
          plugins: { legend:{display:false}, tooltip: CHART_OPTS.tooltip },
          scales: {
            x: { ticks:{ color:'#FAFAFA' }, grid:{ display:false } },
            y: { ticks:{ color:'#7A7674' }, grid:{ color:'rgba(255,255,255,.05)' } },
          }
        }
      });
    }
  }

  // ── GRÁFICA TENDENCIA MENSUAL TAREAS (movida a Row 4) ───────────

  // T3. LÍNEA — Tendencia mensual por tareas únicas
  {
    const mesMap = {};
    const seen = new Set();
    det.forEach(r => {
      const tarea = r.codigo_tarea;
      const fp = String(r.fecha_plan_fin || '');
      if (fp.length < 7) return;
      const key = `${tarea}||${fp.substring(0,7)}`;
      if (seen.has(key)) return; seen.add(key);
      const k = fp.substring(0,7);
      if (!mesMap[k]) mesMap[k] = 0;
      mesMap[k]++;
    });
    const periodos = Object.keys(mesMap).sort();
    const canvas = document.getElementById('chart-rollos-t-mensual');
    if (canvas && periodos.length) {
      chartInstances['rollos-t-mensual'] = new Chart(canvas, {
        type: 'bar',
        data: {
          labels: periodos,
          datasets: [{
            label: 'Tareas',
            data: periodos.map(p => mesMap[p]),
            backgroundColor: 'rgba(153,209,252,.5)', borderColor: '#99D1FC',
            borderWidth: 2, borderRadius: 4,
          }]
        },
        options: {
          responsive: true, maintainAspectRatio: false,
          plugins: { legend:{ labels:{ color:'#FAFAFA', font:{size:11} } }, tooltip: CHART_OPTS.tooltip },
          scales: {
            x: { ticks:{ color:'#7A7674', maxRotation:45 }, grid:{ display:false } },
            y: { ticks:{ color:'#7A7674' }, grid:{ color:'rgba(255,255,255,.05)' } },
          }
        }
      });
    }
  }

}

function _buildDona(id, labels, data, colors, title) {
  const canvas = document.getElementById(id);
  if (!canvas) return;
  if (chartInstances[id]) { chartInstances[id].destroy(); }
  chartInstances[id] = new Chart(canvas, {
    type: 'doughnut',
    data: { labels, datasets: [{ data, backgroundColor: labels.map((_,i)=>(colors[i%colors.length]||'#7B8CDE')+'CC'), borderColor:'rgba(0,0,0,.0)', hoverOffset:8 }] },
    options: {
      responsive: true, maintainAspectRatio: false, cutout: '65%',
      plugins: {
        legend: { position:'bottom', labels:{ color:'#FAFAFA', font:{size:11}, padding:10, boxWidth:12 } },
        tooltip: CHART_OPTS.tooltip,
        title: { display:false }
      }
    }
  });
}

function _destroyChart(key) {
  if (chartInstances[key]) { chartInstances[key].destroy(); delete chartInstances[key]; }
}

// ── Render principal del tab ──────────────────────────────────────
function renderRollosTab() {
  if (!ROLLOS_RAW) {
    const el = document.getElementById('rollos-kpi-grid');
    if (el) el.innerHTML =
      '<div class="empty-state"><div class="icon">⏳</div><p>Cargando datos de rollos... Si persiste, recarga la página.</p></div>';
    // Reintentar en 2 s por si el fallback async aún no terminó
    setTimeout(() => { if (ROLLOS_RAW) renderRollosTab(); }, 2000);
    return;
  }
  renderRollosKPIs();
  renderRollosKPIsTareas();
  renderRollosANSRow();
  renderRollosCharts();
  renderRollosDetalleTable();
  renderRollosComercioTable();
  renderRollosRefTable();
  if (typeof window.renderRollosInvBodegaKPIs === 'function') window.renderRollosInvBodegaKPIs();
}

// ── Helpers ────────────────────────────────────────────────────────
function statusPill(val) {
  const v = (val||'').toUpperCase();
  let cls = 'status-default';
  if (v === 'ENTREGADO')                       cls = 'status-entregado';
  else if (v.includes('TRANSITO'))             cls = 'status-transito';
  else if (v.includes('ALISTAM'))              cls = 'status-alistamiento';
  else if (v.includes('DEVOLUC'))              cls = 'status-devolucion';
  else if (v === 'CANCELADO')                  cls = 'status-cancelado';
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

// ── Tabla Detalle por Tarea ────────────────────────────────────────
// Columnas que exponen TODOS los datos relevantes de una tarea y su MO
// Fuente: TABLERO_ROLLOS_FILAS (data_tablero_rollos.json.gz)
const DETALLE_COLS = [
  // ── Identificación de la tarea ─────────────────────────────────
  { label:'Tarea / MO',            fn: r => r.codigo_tarea || '—',            group: 'tarea' },
  { label:'Cód. Operación',        fn: r => r.codigo_operacion || '—',        group: 'tarea' },
  { label:'Plantilla Tarea',       fn: r => r.nombre_plantilla_tarea || '—',  group: 'tarea' },
  { label:'Estado Tarea',          fn: r => r.estado_tarea || '—', isStatus: true, group: 'tarea' },
  // ── Sitio / Corresponsal ────────────────────────────────────────
  { label:'Cód. Sitio',            fn: r => r.cod_sitio || '—',               group: 'sitio' },
  { label:'Nombre Sitio',          fn: r => r.nombre_sitio || '—',            group: 'sitio' },
  { label:'Destino',               fn: r => r.nombre_ubicacion_destino || '—',group: 'sitio' },
  { label:'Origen',                fn: r => r.nombre_ubicacion_origen || '—', group: 'sitio' },
  { label:'MO Corresponsal',       fn: r => r.cal_codigo_mo || '—',           group: 'sitio' },
  // ── Ubicación ───────────────────────────────────────────────────
  { label:'Departamento',          fn: r => r.departamento || '—',            group: 'geo' },
  { label:'Ciudad',                fn: r => r.ciudad || '—',                  group: 'geo' },
  // ── Proyecto ────────────────────────────────────────────────────
  { label:'Proyecto',              fn: r => r.proyecto || '—',                group: 'proyecto' },
  { label:'Subproyecto',           fn: r => r.subproyecto || '—',             group: 'proyecto' },
  { label:'Tipo Flujo',            fn: r => r.tipo_flujo || '—',              group: 'proyecto' },
  // ── Material ────────────────────────────────────────────────────
  { label:'Cód. Material',         fn: r => r.codigo_material || '—',         group: 'material' },
  { label:'Material',              fn: r => r.nombre_material || '—',         group: 'material' },
  { label:'Cantidad',              fn: r => r.cantidad != null ? Number(r.cantidad).toLocaleString('es-CO') : '—', isNum: true, group: 'material' },
  // ── Fechas ──────────────────────────────────────────────────────
  { label:'Plan Inicio',           fn: r => r.fecha_plan_inicio || '—',       group: 'fechas' },
  { label:'Plan Fin',              fn: r => r.fecha_plan_fin || '—',          group: 'fechas' },
  { label:'Fecha Entrega',         fn: r => r.fecha_entrega_raw || r.fecha_entrega || '—', group: 'fechas' },
  // ── Logística ───────────────────────────────────────────────────
  { label:'Guía',                  fn: r => r.guia || '—',                    group: 'logistica' },
  { label:'Transportadora',        fn: r => r.transportadora || '—',          group: 'logistica' },
  // ── Stock del corresponsal (cal_*) ──────────────────────────────
  { label:'Saldo Rollos',          fn: r => r.cal_saldo_rollos != null && r.cal_saldo_rollos !== 0 ? Number(r.cal_saldo_rollos).toLocaleString('es-CO') : '—', isNum: true, group: 'stock' },
  { label:'Saldo Días',            fn: r => r.cal_saldo_dias != null && r.cal_saldo_dias !== 0 ? Number(r.cal_saldo_dias).toFixed(0) : '—', isDias: true, group: 'stock' },
  { label:'P. Reorden',            fn: r => r.cal_punto_reorden != null && r.cal_punto_reorden !== 0 ? Number(r.cal_punto_reorden).toLocaleString('es-CO') : '—', group: 'stock' },
  { label:'Prom/Mes',              fn: r => r.cal_prom_mensual != null && r.cal_prom_mensual !== 0 ? Number(r.cal_prom_mensual).toLocaleString('es-CO', {maximumFractionDigits:1}) : '—', group: 'stock' },
  { label:'Estado Punto',          fn: r => r.cal_estado_punto || '—',        group: 'stock' },
  { label:'Próx. Abast.',          fn: r => r.cal_fecha_abst || '—',          group: 'stock' },
];

// Búsqueda inline en la tabla detalle
let _rdtSearch = '';

window._rdtSetSearch = function(val) {
  _rdtSearch = (val || '').toLowerCase().trim();
  rollosDetallePage = 1;
  renderRollosDetalleTable();
};

// Colores de grupo para el header
const _RDT_GROUP_COLORS = {
  tarea:    { bg: 'rgba(153,209,252,.12)', border: 'rgba(153,209,252,.3)', text: '#99D1FC' },
  sitio:    { bg: 'rgba(176,242,174,.10)', border: 'rgba(176,242,174,.3)', text: '#B0F2AE' },
  geo:      { bg: 'rgba(255,255,255,.04)', border: 'rgba(255,255,255,.1)', text: '#94a3b8' },
  proyecto: { bg: 'rgba(223,255,97,.08)',  border: 'rgba(223,255,97,.25)', text: '#DFFF61' },
  material: { bg: 'rgba(244,157,110,.10)', border: 'rgba(244,157,110,.3)', text: '#F49D6E' },
  fechas:   { bg: 'rgba(192,132,252,.08)', border: 'rgba(192,132,252,.3)', text: '#C084FC' },
  logistica:{ bg: 'rgba(255,192,77,.08)', border: 'rgba(255,192,77,.3)',  text: '#FFC04D' },
  stock:    { bg: 'rgba(255,92,92,.08)',   border: 'rgba(255,92,92,.25)', text: '#FF5C5C' },
};

function _rdtStatusCell(val) {
  const v = (val || '').toUpperCase();
  let col = '#94a3b8', bg = 'rgba(255,255,255,.06)';
  if (v === 'COMPLETADA')                   { col = '#B0F2AE'; bg = 'rgba(176,242,174,.1)'; }
  else if (v === 'COMPLETADA CON PENDIENTES') { col = '#DFFF61'; bg = 'rgba(223,255,97,.08)'; }
  else if (v === 'EN PROCESO')              { col = '#99D1FC'; bg = 'rgba(153,209,252,.1)'; }
  else if (v === 'ABIERTA')                 { col = '#FFC04D'; bg = 'rgba(255,192,77,.1)'; }
  else if (v.includes('DEVOLUC') || v.includes('DEVUELTO')) { col = '#FF5C5C'; bg = 'rgba(255,92,92,.1)'; }
  else if (v === 'CANCELADO' || v === 'CANCELADA') { col = '#64748b'; bg = 'rgba(100,116,139,.1)'; }
  return `<span style="display:inline-block;padding:2px 9px;border-radius:20px;font-size:10px;font-weight:600;color:${col};background:${bg};border:1px solid ${col}44;white-space:nowrap;">${val || '—'}</span>`;
}

function _rdtDiasCell(val) {
  const n = parseFloat(val);
  if (isNaN(n) || val === '—') return `<span style="color:#475569;">—</span>`;
  let col = '#B0F2AE'; // ok
  if (n < 30)  col = '#FF5C5C';  // crítico
  else if (n < 60) col = '#FFC04D'; // alerta
  else if (n < 90) col = '#DFFF61'; // atención
  return `<span style="font-family:'JetBrains Mono',monospace;font-size:11px;font-weight:700;color:${col};">${Math.round(n)}d</span>`;
}

function renderRollosDetalleTable() {
  const wrap  = document.getElementById('rollos-detalle-wrap');
  const count = document.getElementById('rollos-detalle-count');
  if (!wrap) return;

  // Aplicar búsqueda inline si hay término
  let data = ROLLOS_DETALLE;
  if (_rdtSearch) {
    const t = _rdtSearch;
    data = data.filter(r => {
      return (r.codigo_tarea||'').toLowerCase().includes(t) ||
             (r.cod_sitio||'').toLowerCase().includes(t) ||
             (r.nombre_sitio||'').toLowerCase().includes(t) ||
             (r.nombre_ubicacion_destino||'').toLowerCase().includes(t) ||
             (r.nombre_material||'').toLowerCase().includes(t) ||
             (r.guia||'').toLowerCase().includes(t) ||
             (r.proyecto||'').toLowerCase().includes(t) ||
             (r.ciudad||'').toLowerCase().includes(t) ||
             (r.departamento||'').toLowerCase().includes(t) ||
             (r.cal_codigo_mo||'').toLowerCase().includes(t) ||
             (r.transportadora||'').toLowerCase().includes(t) ||
             (r.estado_tarea||'').toLowerCase().includes(t);
    });
  }

  const pages  = Math.max(1, Math.ceil(data.length / ROLLOS_PAGE_SIZE));
  const slice  = data.slice((rollosDetallePage-1)*ROLLOS_PAGE_SIZE, rollosDetallePage*ROLLOS_PAGE_SIZE);

  if (!ROLLOS_DETALLE.length) {
    wrap.innerHTML = '<div class="empty-state"><div class="icon">📭</div><p>Sin registros</p></div>';
    if (count) count.textContent = '0 registros';
    return;
  }

  // Construir grupos de header
  const groupSpans = {};
  DETALLE_COLS.forEach(c => {
    groupSpans[c.group] = (groupSpans[c.group] || 0) + 1;
  });
  const groupOrder = ['tarea','sitio','geo','proyecto','material','fechas','logistica','stock'];
  const groupLabels = {
    tarea: '🔖 TAREA', sitio: '🏪 SITIO / CORRESPONSAL', geo: '📍 UBICACIÓN',
    proyecto: '📁 PROYECTO', material: '📦 MATERIAL', fechas: '📅 FECHAS',
    logistica: '🚚 LOGÍSTICA', stock: '📊 STOCK MO',
  };

  const groupHeaderCells = groupOrder.map(g => {
    const gc = _RDT_GROUP_COLORS[g];
    const span = groupSpans[g] || 0;
    if (!span) return '';
    return `<th colspan="${span}" style="padding:6px 10px;font-size:9px;font-weight:800;letter-spacing:.8px;text-align:center;color:${gc.text};background:${gc.bg};border-bottom:1px solid ${gc.border};border-right:1px solid rgba(255,255,255,.05);white-space:nowrap;">${groupLabels[g]}</th>`;
  }).join('');

  const colHeaderCells = DETALLE_COLS.map(c => {
    const gc = _RDT_GROUP_COLORS[c.group];
    return `<th style="padding:8px 10px;font-size:9px;color:#64748b;font-weight:700;letter-spacing:.4px;white-space:nowrap;text-align:${c.isNum||c.isDias?'right':'left'};border-bottom:2px solid ${gc.border};border-right:1px solid rgba(255,255,255,.03);">${c.label}</th>`;
  }).join('');

  const rows = slice.map((r, idx) => {
    const bg = idx % 2 === 0 ? 'transparent' : 'rgba(255,255,255,.013)';
    const cells = DETALLE_COLS.map(c => {
      const v = c.fn(r);
      let inner;
      if (c.isStatus)  inner = _rdtStatusCell(v);
      else if (c.isDias) inner = _rdtDiasCell(v);
      else if (c.isNum)  inner = `<span style="font-family:'JetBrains Mono',monospace;font-size:11px;font-weight:600;color:#B0F2AE;">${v}</span>`;
      else               inner = `<span style="color:${v==='—'?'#334155':'#e2e8f0'}">${v}</span>`;
      const align = (c.isNum || c.isDias) ? 'right' : 'left';
      return `<td style="padding:7px 10px;font-size:11px;border-bottom:1px solid rgba(255,255,255,.04);border-right:1px solid rgba(255,255,255,.02);text-align:${align};white-space:nowrap;max-width:220px;overflow:hidden;text-overflow:ellipsis;" title="${String(c.fn(r)).replace(/"/g,'&quot;')}">${inner}</td>`;
    }).join('');
    return `<tr style="background:${bg};transition:background .12s;"
      onmouseover="this.style.background='rgba(176,242,174,.04)'"
      onmouseout="this.style.background='${bg}'">${cells}</tr>`;
  }).join('');

  wrap.innerHTML = `
    <div style="overflow-x:auto;">
      <table style="width:100%;border-collapse:collapse;font-family:'Outfit',sans-serif;font-size:12px;min-width:1800px;">
        <thead style="position:sticky;top:0;z-index:4;">
          <tr style="background:rgba(6,12,20,.98);">${groupHeaderCells}</tr>
          <tr style="background:rgba(6,12,20,.95);">${colHeaderCells}</tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
    </div>`;

  const totalStr = _rdtSearch
    ? `${data.length.toLocaleString('es-CO')} de ${ROLLOS_DETALLE.length.toLocaleString('es-CO')} registros`
    : `${ROLLOS_DETALLE.length.toLocaleString('es-CO')} registros`;
  if (count) count.textContent = totalStr;
  mkPagination('rollos-detalle-pagination', rollosDetallePage, pages, 'goRollosDetallePage');
}

function goRollosDetallePage(p) {
  const pages = Math.ceil(ROLLOS_DETALLE.length / ROLLOS_PAGE_SIZE);
  if (p >= 1 && p <= pages) { rollosDetallePage = p; renderRollosDetalleTable(); }
}

function applyRollosDetalleSearch() {
  if (!ROLLOS_RAW) return;
  const codigo  = (document.getElementById('rf-codigo-tarea')?.value||'').trim().toUpperCase();
  const codSitio= (document.getElementById('rf-cod-sitio')?.value||'').trim().toUpperCase();
  const guia    = (document.getElementById('rf-guia')?.value||'').trim().toUpperCase();
  
  const selEstado     = window._msGetSels('rf-estado');
  const selFlujo      = window._msGetSels('rf-tipo-flujo-det');
  const selDepto      = window._msGetSels('rf-departamento-det');
  const selCiudad     = window._msGetSels('rf-ciudad-det');
  const selProyecto   = window._msGetSels('rf-proyecto-det');
  const selAnio       = window._msGetSels('rf-anio');
  const selMes        = window._msGetSels('rf-mes');
  const selMaterial   = window._msGetSels('rf-nombre-material-det');
  const selPlantilla  = window._msGetSels('rf-plantilla-tarea-det');
  const selEstadoPunto= window._msGetSels('rf-estado-punto-det');

  ROLLOS_DETALLE = ROLLOS_FILTERED.filter(r => {
    if (codigo   && !(r.codigo_tarea||'').toUpperCase().includes(codigo))    return false;
    if (codSitio && !(r.cod_sitio||'').toUpperCase().includes(codSitio))     return false;
    if (guia     && !(r.guia||'').toUpperCase().includes(guia))              return false;
    
    if (selEstado    && !selEstado.includes((r.estado||'').toUpperCase()))                                 return false;
    if (selFlujo     && !selFlujo.includes((r.tipo_flujo||'').toUpperCase()))                             return false;
    if (selDepto     && !selDepto.includes((r.departamento||'').toUpperCase()))                           return false;
    if (selCiudad    && !selCiudad.includes((r.ciudad||'').toUpperCase()))                                return false;
    if (selProyecto  && !selProyecto.includes((r.proyecto||'').toUpperCase()))                            return false;
    if (selAnio      && !selAnio.includes(String(r.anio)))                                                return false;
    if (selMes       && !selMes.includes(String(r.mes).padStart(2,'0')))                                  return false;
    if (selMaterial  && !selMaterial.includes((r.nombre_material||'').toUpperCase()))                     return false;
    if (selPlantilla && !selPlantilla.includes((r.nombre_plantilla_tarea||'').toUpperCase()))             return false;
    if (selEstadoPunto && !selEstadoPunto.includes((r.cal_estado_punto||'').toUpperCase()))               return false;
    
    return true;
  });
  rollosDetallePage = 1;
  renderRollosDetalleTable();
}

function resetRollosDetalleSearch() {
  ['rf-codigo-tarea','rf-guia','rf-cod-sitio'].forEach(id => { const el=document.getElementById(id); if(el) el.value=''; });
  ['rf-estado','rf-anio','rf-mes','rf-tipo-flujo-det','rf-departamento-det','rf-ciudad-det','rf-proyecto-det','rf-nombre-material-det','rf-plantilla-tarea-det','rf-estado-punto-det'].forEach(id => {
    const el = document.getElementById(id);
    if (el && el.classList.contains('ms-container')) window._msAction(id, 'clear');
    else if (el) el.value = '';
  });
  ROLLOS_DETALLE = ROLLOS_FILTERED.slice();
  rollosDetallePage = 1;
  renderRollosDetalleTable();
}

// ── Tabla Comercio ─────────────────────────────────────────────────
const COMERCIO_COLS = [
  { label:'Cod. Comercio',  fn: r => r.cod_comercio || '—' },
  { label:'Nombre Sitio',   fn: r => r.nombre_sitio || '—' },
  { label:'Dirección',      fn: r => r.direccion || '—' },
  { label:'Ciudad',         fn: r => r.ciudad || '—' },
  { label:'Departamento',   fn: r => r.departamento || '—' },
  { label:'Estado',         fn: r => r.estado || '—', isStatus: true },
  { label:'Tipo Envío',     fn: r => r.tipo_envio || '—' },
  { label:'Cantidad',       fn: r => r.cantidad ?? '—' },
  { label:'Tareas',         fn: r => r.tareas ?? '—' },
];

function renderRollosComercioTable() {
  const wrap  = document.getElementById('rollos-comercio-wrap');
  const count = document.getElementById('rollos-comercio-count');
  if (!wrap) return;
  const data  = ROLLOS_COMERCIO;
  const pages = Math.max(1, Math.ceil(data.length / ROLLOS_PAGE_SIZE));
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

function applyComercioFilters() {
  if (!ROLLOS_RAW) return;
  const cod    = (document.getElementById('rf-cod-comercio')?.value||'').trim().toUpperCase();
  const nombre = (document.getElementById('rf-nombre-sitio')?.value||'').trim().toUpperCase();
  const ciudad = (document.getElementById('rf-com-ciudad')?.value||'').trim().toUpperCase();
  
  const selDepto = window._msGetSels('rf-com-departamento');
  const selEst   = window._msGetSels('rf-com-estado');
  const selTipo  = window._msGetSels('rf-com-tipo');

  const sitiosFiltrados = new Set(ROLLOS_FILTERED.map(r => r.cod_sitio));

  ROLLOS_COMERCIO = (ROLLOS_RAW.comercio || []).filter(r => {
    if (cod    && !(r.cod_comercio||'').toUpperCase().includes(cod))    return false;
    if (nombre && !(r.nombre_sitio||'').toUpperCase().includes(nombre)) return false;
    if (ciudad && !(r.ciudad||'').toUpperCase().includes(ciudad))       return false;
    
    if (selDepto && !selDepto.includes((r.departamento||'').toUpperCase())) return false;
    if (selEst   && !selEst.includes((r.estado||'').toUpperCase()))       return false;
    if (selTipo  && !selTipo.includes((r.tipo_envio||'').toUpperCase()))    return false;

    const hayFiltros = ROLLOS_FILTERED.length < (ROLLOS_RAW.detalle||[]).length;
    if (hayFiltros && !sitiosFiltrados.has(r.cod_comercio)) return false;
    return true;
  });
  rollosComercioPage = 1;
  renderRollosComercioTable();
}

function resetComercioFilters() {
  ['rf-cod-comercio','rf-nombre-sitio','rf-com-ciudad'].forEach(id => { const el=document.getElementById(id); if(el) el.value=''; });
  ['rf-com-estado','rf-com-tipo','rf-com-departamento'].forEach(id => {
    const el = document.getElementById(id);
    if (el && el.classList.contains('ms-container')) window._msAction(id, 'clear');
    else if (el) el.value = '';
  });
  const sitiosFiltrados = new Set(ROLLOS_FILTERED.map(r => r.cod_sitio));
  ROLLOS_COMERCIO = (ROLLOS_RAW?.comercio||[]).filter(r =>
    ROLLOS_FILTERED.length === (ROLLOS_RAW.detalle||[]).length || sitiosFiltrados.has(r.cod_comercio)
  );
  rollosComercioPage = 1;
  renderRollosComercioTable();
}

// ── Tabla Referencias ─────────────────────────────────────────────
let ROLLOS_REF_FILTERED = [];
let rollosRefPage = 1;

function renderRollosRefTable() {
  const wrap  = document.getElementById('rollos-ref-wrap');
  const count = document.getElementById('rollos-ref-count');
  if (!wrap) return;
  const allData = ROLLOS_RAW?.referencias || [];

  // Si no está inicializado aún, inicializar con todos
  if (!ROLLOS_REF_FILTERED.length && allData.length) ROLLOS_REF_FILTERED = allData.slice();

  // Poblar selects de filtros referencias si están vacíos
  const refEstEl  = document.getElementById('rf-ref-estado');
  const refDepEl  = document.getElementById('rf-ref-departamento');
  const refCiuEl  = document.getElementById('rf-ref-ciudad');
  if (refEstEl && refEstEl.options.length <= 1) {
    const estados = [...new Set(allData.map(r=>r.estado).filter(Boolean))].sort();
    refEstEl.innerHTML = '<option value="">Todos</option>' + estados.map(e=>`<option value="${e}">${e}</option>`).join('');
  }
  if (refDepEl && refDepEl.options.length <= 1) {
    const deptos = [...new Set(allData.map(r=>r.departamento).filter(Boolean))].sort();
    refDepEl.innerHTML = '<option value="">Todos</option>' + deptos.map(d=>`<option value="${d}">${d}</option>`).join('');
  }
  if (refCiuEl && refCiuEl.options.length <= 1) {
    const ciudades = [...new Set(allData.map(r=>r.ciudad).filter(Boolean))].sort();
    refCiuEl.innerHTML = '<option value="">Todos</option>' + ciudades.map(c=>`<option value="${c}">${c}</option>`).join('');
  }

  const data  = ROLLOS_REF_FILTERED;
  const pages = Math.max(1, Math.ceil(data.length / ROLLOS_PAGE_SIZE));
  const slice = data.slice((rollosRefPage-1)*ROLLOS_PAGE_SIZE, rollosRefPage*ROLLOS_PAGE_SIZE);

  if (!data.length) {
    wrap.innerHTML = '<div class="empty-state"><div class="icon">📭</div><p>Sin referencias</p></div>';
    if (count) count.textContent = '0 registros';
    return;
  }

  wrap.innerHTML = `<table><thead><tr>
    <th>Referencia (Material)</th><th>Cantidad</th><th>Estado</th><th>Departamento</th><th>Ciudad</th>
  </tr></thead><tbody>
    ${slice.map(r=>`<tr>
      <td>${r.material||'—'}</td>
      <td style="font-family:'JetBrains Mono',monospace;font-weight:600;color:var(--azul-cielo)">${r.cantidad??'—'}</td>
      <td>${statusPill(r.estado)}</td>
      <td>${r.departamento||'—'}</td>
      <td>${r.ciudad||'—'}</td>
    </tr>`).join('')}
  </tbody></table>`;

  if (count) count.textContent = `${data.length} referencias`;
  mkPagination('rollos-ref-pagination', rollosRefPage, pages, 'goRollosRefPage');
}

function goRollosRefPage(p) {
  const pages = Math.ceil(ROLLOS_REF_FILTERED.length / ROLLOS_PAGE_SIZE);
  if (p >= 1 && p <= pages) { rollosRefPage = p; renderRollosRefTable(); }
}

function applyRefFilters() {
  if (!ROLLOS_RAW) return;
  const nombre = (document.getElementById('rf-ref-nombre')?.value||'').trim().toUpperCase();
  const estado = (document.getElementById('rf-ref-estado')?.value||'').toUpperCase();
  const depto  = (document.getElementById('rf-ref-departamento')?.value||'').toUpperCase();
  const ciudad = (document.getElementById('rf-ref-ciudad')?.value||'').toUpperCase();

  ROLLOS_REF_FILTERED = (ROLLOS_RAW.referencias || []).filter(r => {
    const ref = (r.material||'').toUpperCase();
    if (nombre && !ref.includes(nombre)) return false;
    if (estado && (r.estado||'').toUpperCase() !== estado) return false;
    if (depto  && (r.departamento||'').toUpperCase() !== depto) return false;
    if (ciudad && (r.ciudad||'').toUpperCase() !== ciudad) return false;
    return true;
  });
  rollosRefPage = 1;
  renderRollosRefTable();
}

function resetRefFilters() {
  ['rf-ref-nombre'].forEach(id => { const el=document.getElementById(id); if(el) el.value=''; });
  ['rf-ref-estado','rf-ref-departamento','rf-ref-ciudad'].forEach(id => { const el=document.getElementById(id); if(el) el.value=''; });
  ROLLOS_REF_FILTERED = (ROLLOS_RAW?.referencias||[]).slice();
  rollosRefPage = 1;
  renderRollosRefTable();
}

// ── Exportar Excel (Rollos) ───────────────────────────────────────
function exportRollosDetalleExcel() {
  const src = _rdtSearch
    ? ROLLOS_DETALLE.filter(r => {
        const t = _rdtSearch;
        return (r.codigo_tarea||'').toLowerCase().includes(t) ||
               (r.cod_sitio||'').toLowerCase().includes(t) ||
               (r.nombre_sitio||'').toLowerCase().includes(t) ||
               (r.nombre_ubicacion_destino||'').toLowerCase().includes(t) ||
               (r.nombre_material||'').toLowerCase().includes(t) ||
               (r.guia||'').toLowerCase().includes(t) ||
               (r.proyecto||'').toLowerCase().includes(t) ||
               (r.ciudad||'').toLowerCase().includes(t) ||
               (r.cal_codigo_mo||'').toLowerCase().includes(t);
      })
    : ROLLOS_DETALLE;
  const data = src.map(r => ({
    'Tarea / MO'           : r.codigo_tarea || '',
    'Cód. Operación'       : r.codigo_operacion || '',
    'Plantilla Tarea'      : r.nombre_plantilla_tarea || '',
    'Estado Tarea'         : r.estado_tarea || '',
    'Cód. Sitio'           : r.cod_sitio || '',
    'Nombre Sitio'         : r.nombre_sitio || '',
    'Destino'              : r.nombre_ubicacion_destino || '',
    'Origen'               : r.nombre_ubicacion_origen || '',
    'MO Corresponsal'      : r.cal_codigo_mo || '',
    'Departamento'         : r.departamento || '',
    'Ciudad'               : r.ciudad || '',
    'Proyecto'             : r.proyecto || '',
    'Subproyecto'          : r.subproyecto || '',
    'Tipo Flujo'           : r.tipo_flujo || '',
    'Cód. Material'        : r.codigo_material || '',
    'Material'             : r.nombre_material || '',
    'Cantidad'             : r.cantidad || 0,
    'Plan Inicio'          : r.fecha_plan_inicio || '',
    'Plan Fin'             : r.fecha_plan_fin || '',
    'Fecha Entrega'        : r.fecha_entrega_raw || r.fecha_entrega || '',
    'Guía'                 : r.guia || '',
    'Transportadora'       : r.transportadora || '',
    'Saldo Rollos'         : r.cal_saldo_rollos || '',
    'Saldo Días'           : r.cal_saldo_dias || '',
    'Punto Reorden'        : r.cal_punto_reorden || '',
    'Prom/Mes'             : r.cal_prom_mensual || '',
    'Estado Punto'         : r.cal_estado_punto || '',
    'Próx. Abast.'         : r.cal_fecha_abst || '',
  }));
  _exportExcelRollos(data, 'Rollos_Detalle_Tareas');
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
  const data = (ROLLOS_REF_FILTERED.length ? ROLLOS_REF_FILTERED : ROLLOS_RAW?.referencias || []).map(r => ({
    'Referencia (Material)': r.material,
    Cantidad: r.cantidad, Estado: r.estado,
    Departamento: r.departamento, Ciudad: r.ciudad,
  }));
  _exportExcelRollos(data, 'Rollos_Referencias');
}

// ══════════════════════════════════════════════════════════════════
//  TAB INCUMPLIMIENTOS ANS
// ══════════════════════════════════════════════════════════════════
let incumpPage = 1;
let INCUMP_DATA = [];
let incumpRespFilter = 'todos';
const INCUMP_PAGE_SIZE = 50;

function renderIncumplimientosTab() {
  const k = computeKPIs(FILTERED);
  const rows = k.incumplimientosRows;
  INCUMP_DATA = rows;
  incumpPage = 1;

  // ── Resumen KPIs de incumplimientos ──
  const resumenEl = document.getElementById('incump-resumen');
  if (resumenEl) {
    const today = new Date(); today.setHours(0,0,0,0);

    // KPIs de cumplimiento ANS: solo registros donde responsable === LINEACOM
    const lineacomRows = rows.filter(r =>
      (r['RESPONSABLE INCUMPLIMIENTO'] || '').trim().toUpperCase() === 'LINEACOM'
    );

    // Conteo por responsable y causal (sobre el total de incumplimientos)
    const porResp = {};
    const porCausal = {};

    // Días retraso: solo LINEACOM
    let totalDias = 0, conDias = 0, tardios = 0;

    lineacomRows.forEach(r => {
      const est = (r['ESTADO DATAFONO'] || '').toUpperCase();
      const lim = getFechaLimite(r);
      const fe  = parseDate(r['FECHA ENTREGA AL COMERCIO'] || '');
      if (est === 'ENTREGADO' && lim && fe) {
        const limD = new Date(lim); limD.setHours(0,0,0,0);
        const feD  = new Date(fe);  feD.setHours(0,0,0,0);
        const dias = Math.round((feD - limD) / 86400000);
        if (dias > 0) { totalDias += dias; conDias++; tardios++; }
      } else if (lim) {
        conDias++;
      }
    });

    // Barras por responsable y causal: sobre todos los incumplimientos
    rows.forEach(r => {
      const resp = (r['RESPONSABLE INCUMPLIMIENTO'] || '(Sin responsable)').trim() || '(Sin responsable)';
      const caus = (r['CAUSAL INCU'] || '(Sin causal)').trim() || '(Sin causal)';
      porResp[resp]   = (porResp[resp]   || 0) + 1;
      porCausal[caus] = (porCausal[caus] || 0) + 1;
    });

    const avgDias  = conDias ? Math.round(totalDias / (tardios || 1)) : 0;
    const lcTotal  = lineacomRows.length;
    const lcTardios = tardios;

    const topResp   = Object.entries(porResp).sort((a,b)=>b[1]-a[1]);
    const topCausal = Object.entries(porCausal).sort((a,b)=>b[1]-a[1]);

    resumenEl.innerHTML = `
      <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:16px;margin-bottom:28px">
        <div class="kpi-card danger" style="padding:20px">
          <span class="kpi-icon">📛</span>
          <div class="kpi-label">Total Incumplimientos</div>
          <div class="kpi-value danger">${rows.length}</div>
          <div class="kpi-sub">${rows.filter(r=>(r['ESTADO DATAFONO']||'').toUpperCase()==='ENTREGADO').length} entregados tardíos · ${rows.filter(r=>(r['ESTADO DATAFONO']||'').toUpperCase()!=='ENTREGADO').length} sin entregar</div>
        </div>
        <div class="kpi-card danger" style="padding:20px">
          <span class="kpi-icon">🏢</span>
          <div class="kpi-label">Incumplimientos LINEACOM</div>
          <div class="kpi-value danger">${lcTotal}</div>
          <div class="kpi-sub">${rows.length ? Math.round(lcTotal/rows.length*100) : 0}% del total · ${lcTardios} tardíos</div>
        </div>
        <div class="kpi-card warn" style="padding:20px">
          <span class="kpi-icon">⏱️</span>
          <div class="kpi-label">Promedio Días Retraso</div>
          <div class="kpi-value warn">${avgDias} días</div>
          <div class="kpi-sub">solo LINEACOM · entregas tardías</div>
        </div>
      </div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:24px">
        <div style="background:var(--surface2);border:1px solid var(--border);border-radius:var(--radius);padding:20px">
          <div style="font-family:'Syne',sans-serif;font-weight:700;color:var(--verde-menta);margin-bottom:14px;font-size:13px;letter-spacing:.5px;text-transform:uppercase">Por Responsable</div>
          ${topResp.map(([r,n])=>`
            <div style="display:flex;align-items:center;gap:10px;margin-bottom:10px">
              <div style="flex:1;font-size:13px;color:var(--blanco)">${r}</div>
              <div style="font-family:'JetBrains Mono',monospace;font-weight:700;color:var(--danger);min-width:32px;text-align:right">${n}</div>
              <div style="width:120px;height:6px;background:rgba(255,255,255,.07);border-radius:3px;overflow:hidden">
                <div style="height:100%;width:${rows.length?Math.round(n/rows.length*100):0}%;background:var(--danger);border-radius:3px"></div>
              </div>
              <div style="font-size:11px;color:var(--muted);min-width:30px">${rows.length?Math.round(n/rows.length*100):0}%</div>
            </div>`).join('')}
        </div>
        <div style="background:var(--surface2);border:1px solid var(--border);border-radius:var(--radius);padding:20px">
          <div style="font-family:'Syne',sans-serif;font-weight:700;color:var(--azul-cielo);margin-bottom:14px;font-size:13px;letter-spacing:.5px;text-transform:uppercase">Por Causal</div>
          ${topCausal.map(([c,n])=>`
            <div style="display:flex;align-items:center;gap:10px;margin-bottom:10px">
              <div style="flex:1;font-size:13px;color:var(--blanco)">${c}</div>
              <div style="font-family:'JetBrains Mono',monospace;font-weight:700;color:var(--azul-cielo);min-width:32px;text-align:right">${n}</div>
              <div style="width:120px;height:6px;background:rgba(255,255,255,.07);border-radius:3px;overflow:hidden">
                <div style="height:100%;width:${rows.length?Math.round(n/rows.length*100):0}%;background:var(--azul-cielo);border-radius:3px"></div>
              </div>
              <div style="font-size:11px;color:var(--muted);min-width:30px">${rows.length?Math.round(n/rows.length*100):0}%</div>
            </div>`).join('')}
        </div>
      </div>`;
  }

  // ── Botones filtro por responsable ──
  const respFiltersEl = document.getElementById('incump-resp-filters');
  if (respFiltersEl) {
    // Obtener responsables únicos del total de incumplimientos
    const respUnicos = ['todos', ...new Set(
      rows.map(r => (r['RESPONSABLE INCUMPLIMIENTO'] || '(Sin responsable)').trim() || '(Sin responsable)')
    ).values()];

    // Contar por responsable
    const respCount = {};
    rows.forEach(r => {
      const v = (r['RESPONSABLE INCUMPLIMIENTO'] || '(Sin responsable)').trim() || '(Sin responsable)';
      respCount[v] = (respCount[v] || 0) + 1;
    });

    respFiltersEl.innerHTML = respUnicos.map(v => {
      const isAll    = v === 'todos';
      const isActive = incumpRespFilter === v;
      const label    = isAll ? `Todos (${rows.length})` : `${v} (${respCount[v] || 0})`;
      const colorMap = { 'LINEACOM': 'var(--azul-cielo)', 'USUARIO': 'var(--warning)' };
      const color    = isAll ? 'var(--verde-menta)' : (colorMap[v.toUpperCase()] || 'var(--muted)');
      const bg       = isActive ? color : 'rgba(255,255,255,.05)';
      const textColor = isActive ? 'var(--negro-cib)' : color;
      const border   = isActive ? color : 'rgba(255,255,255,.12)';
      return `<button onclick="setIncumpRespFilter('${v.replace(/'/g,"\\'")}')"
        style="padding:5px 14px;border-radius:20px;border:1px solid ${border};
               background:${bg};color:${textColor};font-size:12px;font-weight:600;
               cursor:pointer;transition:all .2s;font-family:'Outfit',sans-serif;white-space:nowrap">
        ${label}
      </button>`;
    }).join('');
  }

  _renderIncumpTable();
}

function setIncumpRespFilter(v) {
  incumpRespFilter = v;
  incumpPage = 1;
  // Re-render solo los pills y la tabla (sin recalcular todo el resumen)
  const respFiltersEl = document.getElementById('incump-resp-filters');
  if (respFiltersEl) {
    respFiltersEl.querySelectorAll('button').forEach(btn => {
      const isActive = btn.textContent.trim().startsWith(v === 'todos' ? 'Todos' : v);
      // Re-render simple: llamar renderIncumplimientosTab refresca todo limpiamente
    });
  }
  renderIncumplimientosTab();
}

function _renderIncumpTable() {
  const wrap  = document.getElementById('incump-table-wrap');
  const count = document.getElementById('incump-count');
  if (!wrap) return;

  const today = new Date(); today.setHours(0,0,0,0);

  // Aplicar filtro por responsable
  const data = incumpRespFilter === 'todos'
    ? INCUMP_DATA
    : INCUMP_DATA.filter(r =>
        ((r['RESPONSABLE INCUMPLIMIENTO'] || '(Sin responsable)').trim() || '(Sin responsable)') === incumpRespFilter
      );

  const pages = Math.max(1, Math.ceil(data.length / INCUMP_PAGE_SIZE));
  const slice = data.slice((incumpPage-1)*INCUMP_PAGE_SIZE, incumpPage*INCUMP_PAGE_SIZE);

  if (!data.length) {
    wrap.innerHTML = '<div class="empty-state"><div class="icon">🎉</div><p>Sin incumplimientos registrados</p></div>';
    if (count) count.textContent = '0 registros';
    return;
  }

  const cols = [
    { label:'Comercio',        fn: r => r['Nombre del comercio'] || '—' },
    { label:'ID Sitio',        fn: r => r['ID Comercio'] || '—' },
    { label:'Ciudad',          fn: r => (r['Ciudad'] || '—') + ', ' + (r['Departamento'] || '') },
    { label:'Estado',          fn: r => r['ESTADO DATAFONO'] || '—', isStatus: true },
    { label:'Tipo Solicitud',  fn: r => r['TIPO DE SOLICITUD FACTURACIÓN'] || r['TIPO DE SOLICITUD'] || '—' },
    { label:'Fecha Límite',    fn: r => r['FECHA LIMITE DE ENTREGA'] || '—' },
    { label:'Fecha Entrega',   fn: r => r['FECHA ENTREGA AL COMERCIO'] || '—' },
    { label:'Días Retraso',    fn: r => {
      const est = (r['ESTADO DATAFONO']||'').toUpperCase();
      const lim = getFechaLimite(r);
      if (!lim) return '—';
      const limD = new Date(lim); limD.setHours(0,0,0,0);
      if (est === 'ENTREGADO') {
        const fe = parseDate(r['FECHA ENTREGA AL COMERCIO']||'');
        if (!fe) return '—';
        const feD = new Date(fe); feD.setHours(0,0,0,0);
        return Math.round((feD - limD) / 86400000);
      }
      return Math.round((today - limD) / 86400000);
    }, isDias: true },
    { label:'Responsable',     fn: r => r['RESPONSABLE INCUMPLIMIENTO'] || '(Sin responsable)', isResp: true },
    { label:'Causal',          fn: r => r['CAUSAL INCU'] || '(Sin causal)' },
    { label:'Transportadora',  fn: r => r['TRANSPORTADORA'] || '—' },
    { label:'Guía',            fn: r => r['NÚMERO DE GUIA'] || '—' },
    { label:'Novedades',       fn: r => (r['NOVEDADES']||'').slice(0,120) + ((r['NOVEDADES']||'').length > 120 ? '…' : '') || '—' },
  ];

  wrap.innerHTML = `<table><thead><tr>
    ${cols.map(c=>`<th style="white-space:nowrap">${c.label}</th>`).join('')}
  </tr></thead><tbody>
    ${slice.map((r,i) => `<tr style="background:${i%2?'transparent':'rgba(255,255,255,.02)'}">
      ${cols.map(c => {
        const v = c.fn(r);
        if (c.isStatus) return `<td>${statusPill(String(v))}</td>`;
        if (c.isDias) {
          const n = parseInt(v);
          if (isNaN(n)) return `<td>—</td>`;
          const cls = n > 10 ? 'crit' : n > 3 ? 'warn' : 'ok';
          return `<td><span class="days-stalled ${cls}">${n}d</span></td>`;
        }
        if (c.isResp) {
          const cls = String(v).toUpperCase() === 'LINEACOM' ? 'var(--azul-cielo)' : String(v).toUpperCase() === 'USUARIO' ? 'var(--warning)' : 'var(--muted)';
          return `<td><span style="font-weight:600;color:${cls};font-size:11px;padding:3px 8px;background:rgba(255,255,255,.06);border-radius:4px">${v}</span></td>`;
        }
        return `<td style="max-width:200px;white-space:normal;font-size:12px">${v}</td>`;
      }).join('')}
    </tr>`).join('')}
  </tbody></table>`;

  if (count) count.textContent = `${data.length} incumplimientos${incumpRespFilter !== 'todos' ? ` · filtrando por: ${incumpRespFilter}` : ''}`;
  mkPagination('incump-pagination', incumpPage, pages, 'goIncumpPage');
}

function goIncumpPage(p) {
  const filtered = incumpRespFilter === 'todos'
    ? INCUMP_DATA
    : INCUMP_DATA.filter(r => ((r['RESPONSABLE INCUMPLIMIENTO']||'(Sin responsable)').trim()||'(Sin responsable)') === incumpRespFilter);
  const pages = Math.ceil(filtered.length / INCUMP_PAGE_SIZE);
  if (p >= 1 && p <= pages) { incumpPage = p; _renderIncumpTable(); }
}

function exportIncumpExcel() {
  if (!INCUMP_DATA.length) { alert('Sin datos para exportar.'); return; }
  const today = new Date(); today.setHours(0,0,0,0);
  const data = INCUMP_DATA.map(r => {
    const est = (r['ESTADO DATAFONO']||'').toUpperCase();
    const lim = getFechaLimite(r);
    let dias = '';
    if (lim) {
      const limD = new Date(lim); limD.setHours(0,0,0,0);
      if (est === 'ENTREGADO') {
        const fe = parseDate(r['FECHA ENTREGA AL COMERCIO']||'');
        if (fe) { const feD = new Date(fe); feD.setHours(0,0,0,0); dias = Math.round((feD-limD)/86400000); }
      } else {
        dias = Math.round((today-limD)/86400000);
      }
    }
    return {
      'Comercio':              r['Nombre del comercio'] || '',
      'ID Sitio':              r['ID Comercio'] || '',
      'Ciudad':                r['Ciudad'] || '',
      'Departamento':          r['Departamento'] || '',
      'Estado':                r['ESTADO DATAFONO'] || '',
      'Tipo Solicitud':        r['TIPO DE SOLICITUD FACTURACIÓN'] || r['TIPO DE SOLICITUD'] || '',
      'Fecha Límite':          r['FECHA LIMITE DE ENTREGA'] || '',
      'Fecha Entrega':         r['FECHA ENTREGA AL COMERCIO'] || '',
      'Días Retraso':          dias,
      'Responsable':           r['RESPONSABLE INCUMPLIMIENTO'] || '',
      'Causal Incumplimiento': r['CAUSAL INCU'] || '',
      'Transportadora':        r['TRANSPORTADORA'] || '',
      'Guía':                  r['NÚMERO DE GUIA'] || '',
      'Novedades':             r['NOVEDADES'] || '',
    };
  });
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(data);
  ws['!cols'] = Object.keys(data[0]).map(k => ({ wch: Math.max(k.length+2, ...data.slice(0,50).map(r=>String(r[k]||'').length)) }));
  XLSX.utils.book_append_sheet(wb, ws, 'Incumplimientos');
  XLSX.writeFile(wb, `Incumplimientos_ANS_${new Date().toISOString().slice(0,10)}.xlsx`);
}

function _exportExcelRollos(data, filename) {
  if (!data.length) { alert('Sin datos para exportar.'); return; }
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(data);
  ws['!cols'] = Object.keys(data[0]||{}).map(k => ({ wch: Math.max(k.length, ...data.slice(0,50).map(r=>String(r[k]||'').length)) }));
  XLSX.utils.book_append_sheet(wb, ws, 'Datos');
  XLSX.writeFile(wb, `${filename}_${new Date().toISOString().slice(0,10)}.xlsx`);
}