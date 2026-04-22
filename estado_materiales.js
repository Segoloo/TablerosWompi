/**
 * ╔══════════════════════════════════════════════════════════════════╗
 * ║  estado_materiales.js — Tab Estado de Materiales                ║
 * ║  Sección 1: Estado de Datáfonos en Bodega                       ║
 * ║  Sección 2: Estado de SIMCards                                  ║
 * ╚══════════════════════════════════════════════════════════════════╝
 */

'use strict';

// ── Referencias de datáfonos ──────────────────────────────────────
const EM_DATAFONO_REFS = new Set([
  'PINPAD DESK 1700 - INGENICO',
  'DATAFONO EX6000 - INGENICO',
  'DATAFONO EX4000 - INGENICO',
  'DATAFONO DX4000 PORTATIL - INGENICO',
  'DATAFONO DX4000 ESCRITORIO - INGENICO',
]);

const EM_SIM_REF = 'SIM LTE MULTI 256K P17 POST QR - CLARO';

// ── Posiciones de depósito de interés ────────────────────────────
const EM_POS_DANIO      = 'EN DAÑO';
const EM_POS_INC        = 'DESINSTALADO-INCIDENTE';
const EM_POS_CIERRE     = 'DESINSTALADO-CIERRE';

// ── Colores del sistema ───────────────────────────────────────────
const EM_C = {
  verde:   '#B0F2AE',
  azul:    '#99D1FC',
  lima:    '#DFFF61',
  purple:  '#C084FC',
  danger:  '#F87171',
  warn:    '#FFC04D',
  orange:  '#FB923C',
  teal:    '#67E8F9',
  muted:   '#7A7674',
  bg:      'rgba(10,26,18,.95)',
  card:    'linear-gradient(145deg,rgba(10,26,18,.95),rgba(8,20,14,.9))',
  border:  'rgba(176,242,174,.12)',
};

// ── Estado global del módulo ──────────────────────────────────────
let EM_CHARTS = {};
let EM_DF_ALL  = [];   // todos los datáfonos en bodegas
let EM_SIM_ALL = [];   // todas las simcards
let EM_DF_PAGE = 1;
let EM_SIM_PAGE = 1;
let EM_DF_SEARCH = '';
let EM_SIM_SEARCH = '';
const EM_PAGE_SIZE = 50;

// ─────────────────────────────────────────────────────────────────
//  HELPERS
// ─────────────────────────────────────────────────────────────────
function _emGetPos(row) {
  return (row['Posición en depósito'] || '').trim().toUpperCase();
}

function _emGetComment(row) {
  return (row['Comentarios'] || '').trim();
}

function _emGetRef(row) {
  return (row['Nombre'] || '').trim();
}

function _emGetSerial(row) {
  return (row['Número de serie'] || row['Serial'] || row['SERIAL'] || '').trim();
}

function _emGetBodega(row) {
  return (row['Nombre de la ubicación'] || '').trim();
}

function _emGetAtributos(row) {
  return (row['Atributos'] || '').trim();
}

/**
 * Clasifica el estado de un datáfono según sus campos.
 * Retorna: 'DISPONIBLE' | 'ASOCIADO' | 'EN DAÑO' | 'DES. CIERRE' | 'DES. INCIDENTE'
 */
function _emDfEstado(row) {
  const pos = _emGetPos(row);
  if (pos === EM_POS_DANIO)   return 'EN DAÑO';
  if (pos === EM_POS_CIERRE)  return 'DES. CIERRE';
  if (pos === EM_POS_INC)     return 'DES. INCIDENTE';

  const comment = _emGetComment(row);
  if (!comment) return 'DISPONIBLE';

  // Extrae el primer segmento antes de ' | '
  const parts = comment.split('|');
  const firstNum = (parts[0] || '').trim();
  if (firstNum === '99999' || firstNum === '') return 'DISPONIBLE';
  // Si hay número y no es 99999 → ASOCIADO
  if (/^\d+$/.test(firstNum)) return 'ASOCIADO';
  return 'DISPONIBLE';
}

/**
 * Extrae el tipo de daño desde Comentarios cuando empieza con 99999.
 * Ej: "99999 | LECTOR TARJETAS | 18/03/2026" → "LECTOR TARJETAS"
 */
function _emDaño(row) {
  const comment = _emGetComment(row);
  if (!comment) return null;
  const parts = comment.split('|').map(s => s.trim());
  if (parts[0] !== '99999' || parts.length < 2) return null;
  const tipo = parts[1] || '';
  const exclude = ['DESASOCIADO', 'DISPONIBLE', ''];
  if (exclude.includes(tipo.toUpperCase())) return null;
  return tipo.toUpperCase();
}

/**
 * SIM: determina si está activada y extrae fecha.
 * Busca "FA:DD/MM/AAAA" en Atributos.
 */
function _emSimEstado(row) {
  const attr = _emGetAtributos(row);
  const match = attr.match(/FA:(\d{1,2}\/\d{1,2}\/\d{4})/i);
  if (match) return { activa: true, fecha: match[1] };
  return { activa: false, fecha: null };
}

function _emDestroyChart(id) {
  if (EM_CHARTS[id]) { EM_CHARTS[id].destroy(); delete EM_CHARTS[id]; }
}

// ─────────────────────────────────────────────────────────────────
//  TOOLTIP / LEGEND defaults (copian estilo inventario.js)
// ─────────────────────────────────────────────────────────────────
const EM_TT = {
  backgroundColor: 'rgba(24,23,21,.95)',
  titleColor: '#B0F2AE', bodyColor: '#FAFAFA',
  borderColor: 'rgba(176,242,174,.2)', borderWidth: 1, padding: 12,
  titleFont: { family: 'Syne', size: 13, weight: '700' },
  bodyFont:  { family: 'Outfit', size: 12 },
};
const EM_LEG = {
  labels: { color: '#FAFAFA', font: { family: 'Outfit', size: 11 }, padding: 12, boxWidth: 12 }
};

// ─────────────────────────────────────────────────────────────────
//  CARD BUILDER
// ─────────────────────────────────────────────────────────────────
function _emCard(canvasId, title, sub, accentColor) {
  const color = accentColor || EM_C.verde;
  const div = document.createElement('div');
  div.style.cssText = [
    'background:' + EM_C.card,
    'border:1px solid rgba(176,242,174,.1)',
    'border-top:2px solid ' + color,
    'border-radius:18px',
    'padding:22px 24px 20px',
    'position:relative',
    'overflow:hidden',
    'box-shadow:0 4px 20px rgba(0,0,0,.4),inset 0 1px 0 rgba(255,255,255,.03)',
  ].join(';');
  div.innerHTML =
    '<div style="position:absolute;top:-20px;right:-20px;width:70px;height:70px;border-radius:50%;background:' + color + ';opacity:.05;filter:blur(12px);pointer-events:none;"></div>' +
    '<div style="margin-bottom:16px;">' +
      '<div style="font-family:\'Syne\',sans-serif;font-size:13px;font-weight:700;color:#f1f5f9;letter-spacing:.3px;">' + title + '</div>' +
      (sub ? '<div style="font-size:11px;color:' + EM_C.muted + ';margin-top:3px;font-family:\'Outfit\',sans-serif;">' + sub + '</div>' : '') +
    '</div>' +
    '<canvas id="' + canvasId + '" style="max-height:260px;"></canvas>';
  return div;
}

function _emKpiCard(icon, label, value, sub, color, span2) {
  const c = color || EM_C.verde;
  const div = document.createElement('div');
  div.style.cssText = [
    'background:' + EM_C.card,
    'border:1px solid rgba(176,242,174,.1)',
    'border-top:2px solid ' + c,
    'border-radius:18px',
    'padding:20px',
    'position:relative',
    'overflow:hidden',
    'box-shadow:0 4px 20px rgba(0,0,0,.4)',
    span2 ? 'grid-column:span 2' : '',
  ].filter(Boolean).join(';');
  div.innerHTML =
    '<div style="position:absolute;top:-20px;right:-20px;width:70px;height:70px;border-radius:50%;background:' + c + ';opacity:.05;filter:blur(12px);pointer-events:none;"></div>' +
    '<span style="font-size:20px;display:block;margin-bottom:10px;">' + icon + '</span>' +
    '<div style="font-size:10px;font-weight:700;color:' + EM_C.muted + ';text-transform:uppercase;letter-spacing:1px;margin-bottom:8px;font-family:\'Outfit\',sans-serif;">' + label + '</div>' +
    '<div style="font-family:\'JetBrains Mono\',monospace;font-size:32px;font-weight:700;color:' + c + ';line-height:1;text-shadow:0 0 24px ' + c + '55;margin-bottom:8px;">' + value + '</div>' +
    '<span style="background:' + c + '22;color:' + c + ';font-size:10px;font-weight:700;padding:2px 8px;border-radius:20px;font-family:\'JetBrains Mono\',monospace;">' + sub + '</span>';
  return div;
}

function _emSectionTitle(text, color, emoji) {
  const div = document.createElement('div');
  div.style.cssText = 'font-family:\'Syne\',sans-serif;font-size:16px;font-weight:800;color:' + (color||EM_C.verde) + ';letter-spacing:-.3px;margin:36px 0 18px;display:flex;align-items:center;gap:10px;border-left:3px solid ' + (color||EM_C.verde) + ';padding-left:14px;';
  div.innerHTML = (emoji ? '<span style="font-size:18px">' + emoji + '</span>' : '') + text;
  return div;
}

function _emDivider() {
  const d = document.createElement('div');
  d.style.cssText = 'height:1px;background:linear-gradient(90deg,rgba(176,242,174,.2),transparent);margin:32px 0;';
  return d;
}

function _emTableWrap(id, maxH) {
  const d = document.createElement('div');
  d.id = id;
  d.style.cssText = 'overflow-x:auto;max-height:' + (maxH||'480px') + ';';
  return d;
}

// ─────────────────────────────────────────────────────────────────
//  PAGINATION
// ─────────────────────────────────────────────────────────────────
function _emMkPag(containerId, current, total, callbackName) {
  const el = document.getElementById(containerId);
  if (!el) return;
  if (total <= 1) { el.innerHTML = ''; return; }
  const maxShow = 7;
  let pages = [];
  if (total <= maxShow) {
    for (let i = 1; i <= total; i++) pages.push(i);
  } else {
    pages = [1];
    if (current > 3) pages.push('...');
    for (let i = Math.max(2, current-1); i <= Math.min(total-1, current+1); i++) pages.push(i);
    if (current < total-2) pages.push('...');
    pages.push(total);
  }
  el.innerHTML = pages.map(p =>
    p === '...'
      ? '<span style="color:' + EM_C.muted + ';padding:0 4px;">…</span>'
      : '<button onclick="' + callbackName + '(' + p + ')" style="background:' + (p===current ? EM_C.verde : 'rgba(255,255,255,.06)') + ';color:' + (p===current ? '#0a1a12' : '#FAFAFA') + ';border:none;border-radius:6px;padding:5px 10px;cursor:pointer;font-family:\'JetBrains Mono\',monospace;font-size:11px;font-weight:700;transition:all .15s;">' + p + '</button>'
  ).join('');
}

// ─────────────────────────────────────────────────────────────────
//  MAIN RENDER ENTRY POINT
// ─────────────────────────────────────────────────────────────────
function renderEstadoMateriales() {
  const panel = document.getElementById('panel-estado-materiales');
  if (!panel) return;

  const raw = window.INV_RAW;
  if (!raw || !raw.length) {
    panel.innerHTML = '<div style="display:flex;align-items:center;justify-content:center;padding:80px;gap:16px;"><div class="spinner"></div><span style="color:var(--muted);font-family:\'Outfit\',sans-serif;">Cargando datos de inventario...</span></div>';
    // Retry after data loads
    const retry = setInterval(() => {
      if (window.INV_RAW && window.INV_RAW.length) {
        clearInterval(retry);
        renderEstadoMateriales();
      }
    }, 500);
    return;
  }

  // ── Filtrar datáfonos en bodegas ───────────────────────────────
  EM_DF_ALL = raw.filter(r => {
    const ref = _emGetRef(r);
    const bod = _emGetBodega(r);
    return EM_DATAFONO_REFS.has(ref) && window.INV_BODEGAS && window.INV_BODEGAS.has(bod);
  });

  // ── Filtrar SIMcards ───────────────────────────────────────────
  EM_SIM_ALL = raw.filter(r => _emGetRef(r) === EM_SIM_REF);

  // ── Construir panel ───────────────────────────────────────────
  panel.innerHTML = '';

  // SECCIÓN 1: DATÁFONOS
  panel.appendChild(_emSectionTitle('ESTADO DE DATÁFONOS EN BODEGA', EM_C.verde, '📱'));
  _emRenderDfKPIs(panel);
  _emRenderDfCharts(panel);
  _emRenderBodegaTable(panel);
  _emRenderDañoChart(panel);
  _emRenderDfTable(panel);

  panel.appendChild(_emDivider());

  // SECCIÓN 2: SIMCARDS
  panel.appendChild(_emSectionTitle('ESTADO DE SIMCARDS', EM_C.teal, '📡'));
  _emRenderSimKPIs(panel);
  _emRenderSimCharts(panel);
  _emRenderSimTable(panel);
}

// ─────────────────────────────────────────────────────────────────
//  SECCIÓN 1 — KPIs DATÁFONOS
// ─────────────────────────────────────────────────────────────────
function _emRenderDfKPIs(panel) {
  const rows = EM_DF_ALL;
  const total       = rows.length;
  const disponibles = rows.filter(r => _emDfEstado(r) === 'DISPONIBLE').length;
  const asociados   = rows.filter(r => _emDfEstado(r) === 'ASOCIADO').length;
  const enDanio     = rows.filter(r => _emDfEstado(r) === 'EN DAÑO').length;
  const desCierre   = rows.filter(r => _emDfEstado(r) === 'DES. CIERRE').length;
  const desInc      = rows.filter(r => _emDfEstado(r) === 'DES. INCIDENTE').length;

  const grid = document.createElement('div');
  grid.style.cssText = 'display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:14px;margin-bottom:28px;';

  const kpis = [
    { icon:'📦', label:'TOTAL BODEGAS',       value: total,       sub: '100% datáfonos',   color: EM_C.lima },
    { icon:'✅', label:'DISPONIBLES',          value: disponibles, sub: pctStr(disponibles,total), color: EM_C.verde },
    { icon:'🔗', label:'ASOCIADOS',            value: asociados,   sub: pctStr(asociados,total),   color: EM_C.azul },
    { icon:'💥', label:'EN DAÑO',              value: enDanio,     sub: pctStr(enDanio,total),     color: EM_C.danger },
    { icon:'🔒', label:'DES. CIERRE',          value: desCierre,   sub: pctStr(desCierre,total),   color: EM_C.warn },
    { icon:'🚨', label:'DES. INCIDENTE',       value: desInc,      sub: pctStr(desInc,total),      color: EM_C.orange },
  ];

  kpis.forEach(k => {
    grid.appendChild(_emKpiCard(k.icon, k.label, k.value.toLocaleString('es-CO'), k.sub, k.color));
  });
  panel.appendChild(grid);
}

function pctStr(n, d) {
  if (!d) return '0.0%';
  return (n / d * 100).toFixed(1) + '%';
}

// ─────────────────────────────────────────────────────────────────
//  SECCIÓN 1 — GRÁFICAS DATÁFONOS
// ─────────────────────────────────────────────────────────────────
function _emRenderDfCharts(panel) {
  const rows  = EM_DF_ALL;

  // Contar por posición de depósito
  const posCounts = {
    'EN DAÑO': rows.filter(r => _emGetPos(r) === EM_POS_DANIO).length,
    'DES. INCIDENTE': rows.filter(r => _emGetPos(r) === EM_POS_INC).length,
    'DES. CIERRE': rows.filter(r => _emGetPos(r) === EM_POS_CIERRE).length,
  };

  // Por referencia × estado
  const REFS_SHORT = {
    'PINPAD DESK 1700 - INGENICO':       'PINPAD 1700',
    'DATAFONO EX6000 - INGENICO':        'EX6000',
    'DATAFONO EX4000 - INGENICO':        'EX4000',
    'DATAFONO DX4000 PORTATIL - INGENICO':'DX4000 PORT.',
    'DATAFONO DX4000 ESCRITORIO - INGENICO':'DX4000 ESCR.',
  };

  const refStats = {};
  [...EM_DATAFONO_REFS].forEach(ref => {
    refStats[ref] = { 'EN DAÑO':0, 'DES. INCIDENTE':0, 'DES. CIERRE':0, 'DISPONIBLE':0, 'ASOCIADO':0 };
  });
  rows.forEach(r => {
    const ref = _emGetRef(r);
    const est = _emDfEstado(r);
    if (refStats[ref] && refStats[ref][est] !== undefined) refStats[ref][est]++;
  });

  const chartsGrid = document.createElement('div');
  chartsGrid.style.cssText = 'display:grid;grid-template-columns:repeat(2,1fr);gap:18px;margin-bottom:28px;';

  // Donut: distribución por posición de depósito
  const c1 = _emCard('em-c-pos-dona', 'Distribución por Estado en Depósito', 'Solo EN DAÑO / DES. INC. / DES. CIERRE', EM_C.danger);
  chartsGrid.appendChild(c1);

  // Bar: por referencia × estado
  const c2 = _emCard('em-c-ref-bar', 'Estado por Referencia de Datáfono', 'Desglose de los 3 estados clave', EM_C.azul);
  chartsGrid.appendChild(c2);

  panel.appendChild(chartsGrid);

  // Render Chart 1: Donut posición
  requestAnimationFrame(() => {
    _emDestroyChart('em-c-pos-dona');
    const ctx1 = document.getElementById('em-c-pos-dona');
    if (ctx1) {
      EM_CHARTS['em-c-pos-dona'] = new Chart(ctx1, {
        type: 'doughnut',
        data: {
          labels: Object.keys(posCounts),
          datasets: [{
            data: Object.values(posCounts),
            backgroundColor: [EM_C.danger+'CC', EM_C.orange+'CC', EM_C.warn+'CC'],
            borderColor: 'rgba(0,0,0,0)', borderWidth: 0, hoverOffset: 8,
          }]
        },
        options: {
          cutout: '68%',
          plugins: {
            legend: EM_LEG,
            tooltip: Object.assign({}, EM_TT, {
              callbacks: {
                label: ctx => ' ' + ctx.label + ': ' + ctx.parsed.toLocaleString('es-CO') + ' uds'
              }
            })
          }
        }
      });
    }

    // Chart 2: stacked bar por referencia
    _emDestroyChart('em-c-ref-bar');
    const ctx2 = document.getElementById('em-c-ref-bar');
    if (ctx2) {
      const refLabels = [...EM_DATAFONO_REFS].map(r => REFS_SHORT[r] || r);
      EM_CHARTS['em-c-ref-bar'] = new Chart(ctx2, {
        type: 'bar',
        data: {
          labels: refLabels,
          datasets: [
            { label:'EN DAÑO',       data: [...EM_DATAFONO_REFS].map(r => refStats[r]['EN DAÑO']),       backgroundColor: EM_C.danger+'BB', borderColor: EM_C.danger, borderWidth:1, borderRadius:4 },
            { label:'DES. INCIDENTE',data: [...EM_DATAFONO_REFS].map(r => refStats[r]['DES. INCIDENTE']),backgroundColor: EM_C.orange+'BB', borderColor: EM_C.orange, borderWidth:1, borderRadius:4 },
            { label:'DES. CIERRE',   data: [...EM_DATAFONO_REFS].map(r => refStats[r]['DES. CIERRE']),   backgroundColor: EM_C.warn+'BB',   borderColor: EM_C.warn,   borderWidth:1, borderRadius:4 },
          ]
        },
        options: {
          plugins: { legend: EM_LEG, tooltip: EM_TT },
          scales: {
            x: { stacked: false, grid:{display:false}, ticks:{color:EM_C.muted,font:{family:'Outfit',size:10}} },
            y: { grid:{color:'rgba(255,255,255,.05)'}, ticks:{color:EM_C.muted,font:{family:'Outfit',size:11}} }
          }
        }
      });
    }
  });
}

// ─────────────────────────────────────────────────────────────────
//  SECCIÓN 1 — TABLA POR BODEGA
// ─────────────────────────────────────────────────────────────────
function _emRenderBodegaTable(panel) {
  // Construir mapa por bodega
  const bodMap = {};
  EM_DF_ALL.forEach(r => {
    const bod = _emGetBodega(r);
    if (!bodMap[bod]) bodMap[bod] = { asociado:0, danio:0, cierre:0, incidente:0, disponible:0, total:0 };
    const est = _emDfEstado(r);
    bodMap[bod].total++;
    if (est === 'ASOCIADO')      bodMap[bod].asociado++;
    else if (est === 'EN DAÑO')  bodMap[bod].danio++;
    else if (est === 'DES. CIERRE') bodMap[bod].cierre++;
    else if (est === 'DES. INCIDENTE') bodMap[bod].incidente++;
    else if (est === 'DISPONIBLE') bodMap[bod].disponible++;
  });

  const rows = Object.entries(bodMap).sort((a,b) => b[1].total - a[1].total);
  const maxVal = rows[0] ? rows[0][1].total : 1;

  const wrap = document.createElement('div');
  wrap.style.cssText = 'background:' + EM_C.card + ';border:1px solid ' + EM_C.border + ';border-radius:18px;overflow:hidden;margin-bottom:28px;box-shadow:0 4px 20px rgba(0,0,0,.4);';

  wrap.innerHTML = `
    <div style="display:flex;align-items:center;justify-content:space-between;padding:16px 20px;border-bottom:1px solid rgba(176,242,174,.08);">
      <div>
        <div style="font-family:'Syne',sans-serif;font-size:13px;font-weight:700;color:#f1f5f9;">📊 Estado por Bodega</div>
        <div style="font-size:11px;color:${EM_C.muted};margin-top:2px;">${rows.length} bodegas con datáfonos</div>
      </div>
    </div>
    <div style="overflow-x:auto;max-height:460px;">
      <table style="width:100%;border-collapse:collapse;">
        <thead>
          <tr style="background:rgba(0,0,0,.3);">
            <th style="padding:10px 16px;text-align:left;font-family:'Outfit',sans-serif;font-size:11px;font-weight:700;color:${EM_C.muted};text-transform:uppercase;letter-spacing:1px;white-space:nowrap;">Bodega</th>
            <th style="padding:10px 12px;text-align:center;font-size:11px;font-weight:700;color:${EM_C.azul};text-transform:uppercase;letter-spacing:.5px;white-space:nowrap;">Asociados</th>
            <th style="padding:10px 12px;text-align:center;font-size:11px;font-weight:700;color:${EM_C.danger};text-transform:uppercase;letter-spacing:.5px;white-space:nowrap;">EN DAÑO</th>
            <th style="padding:10px 12px;text-align:center;font-size:11px;font-weight:700;color:${EM_C.warn};text-transform:uppercase;letter-spacing:.5px;white-space:nowrap;">DES. CIERRE</th>
            <th style="padding:10px 12px;text-align:center;font-size:11px;font-weight:700;color:${EM_C.orange};text-transform:uppercase;letter-spacing:.5px;white-space:nowrap;">DES. INCIDENTE</th>
            <th style="padding:10px 12px;text-align:center;font-size:11px;font-weight:700;color:${EM_C.verde};text-transform:uppercase;letter-spacing:.5px;white-space:nowrap;">Disponible</th>
            <th style="padding:10px 12px;text-align:center;font-size:11px;font-weight:700;color:${EM_C.lima};text-transform:uppercase;letter-spacing:.5px;white-space:nowrap;">Total</th>
            <th style="padding:10px 16px;font-size:11px;font-weight:700;color:${EM_C.muted};text-transform:uppercase;letter-spacing:.5px;min-width:100px;">Distribución</th>
          </tr>
        </thead>
        <tbody>
          ${rows.map(([bod, s], i) => {
            const pct = maxVal ? (s.total / maxVal * 100) : 0;
            return `<tr style="border-top:1px solid rgba(255,255,255,.04);background:${i%2?'transparent':'rgba(255,255,255,.015)'};">
              <td style="padding:9px 16px;font-family:'Outfit',sans-serif;font-size:12px;font-weight:600;color:#e2e8f0;white-space:nowrap;">${bod.replace('ALMACEN WOMPI ','').replace('ALMACEN ','')} <span style="font-size:9px;color:${EM_C.muted};font-weight:400;">· ${bod.includes('VP')?(bod.includes('ALQUILER')?'VP Alquiler':'VP Venta'):(bod.includes('BAJA')?'Bajas':'Almacén')}</span></td>
              <td style="padding:9px 12px;text-align:center;font-family:'JetBrains Mono',monospace;font-size:12px;color:${EM_C.azul};">${s.asociado||'—'}</td>
              <td style="padding:9px 12px;text-align:center;font-family:'JetBrains Mono',monospace;font-size:12px;color:${s.danio?EM_C.danger:EM_C.muted};">${s.danio||'—'}</td>
              <td style="padding:9px 12px;text-align:center;font-family:'JetBrains Mono',monospace;font-size:12px;color:${s.cierre?EM_C.warn:EM_C.muted};">${s.cierre||'—'}</td>
              <td style="padding:9px 12px;text-align:center;font-family:'JetBrains Mono',monospace;font-size:12px;color:${s.incidente?EM_C.orange:EM_C.muted};">${s.incidente||'—'}</td>
              <td style="padding:9px 12px;text-align:center;font-family:'JetBrains Mono',monospace;font-size:12px;color:${EM_C.verde};">${s.disponible||'—'}</td>
              <td style="padding:9px 12px;text-align:center;font-family:'JetBrains Mono',monospace;font-size:13px;font-weight:700;color:${EM_C.lima};">${s.total}</td>
              <td style="padding:9px 16px;">
                <div style="background:rgba(255,255,255,.06);border-radius:4px;height:6px;overflow:hidden;min-width:80px;">
                  <div style="height:100%;border-radius:4px;background:linear-gradient(90deg,${EM_C.verde}88,${EM_C.verde});width:${Math.max(pct,2)}%;transition:width .6s ease;"></div>
                </div>
              </td>
            </tr>`;
          }).join('')}
        </tbody>
      </table>
    </div>`;

  panel.appendChild(wrap);
}

// ─────────────────────────────────────────────────────────────────
//  SECCIÓN 1 — GRÁFICA DE TIPOS DE DAÑO
// ─────────────────────────────────────────────────────────────────
function _emRenderDañoChart(panel) {
  // Solo datáfonos con 99999 y algún daño en comentarios
  const daños = {};
  EM_DF_ALL.forEach(r => {
    const d = _emDaño(r);
    if (d) daños[d] = (daños[d] || 0) + 1;
  });
  const sorted = Object.entries(daños).sort((a,b) => b[1]-a[1]);

  if (!sorted.length) return;

  const chartsGrid = document.createElement('div');
  chartsGrid.style.cssText = 'display:grid;grid-template-columns:1fr;gap:18px;margin-bottom:28px;';

  const c3 = _emCard('em-c-dano', 'Tipos de Daño (desde Comentarios)', 'Solo equipos con código 99999 y daño registrado', EM_C.danger);
  c3.style.maxWidth = '640px';
  chartsGrid.appendChild(c3);
  panel.appendChild(chartsGrid);

  const danoPalette = [EM_C.danger, EM_C.orange, EM_C.warn, EM_C.purple, '#F49D6E', '#FF5C5C', '#C084FC', EM_C.azul];

  requestAnimationFrame(() => {
    _emDestroyChart('em-c-dano');
    const ctx = document.getElementById('em-c-dano');
    if (ctx) {
      EM_CHARTS['em-c-dano'] = new Chart(ctx, {
        type: 'bar',
        data: {
          labels: sorted.map(e => e[0]),
          datasets: [{
            label: 'Cantidad',
            data: sorted.map(e => e[1]),
            backgroundColor: sorted.map((_,i) => danoPalette[i % danoPalette.length] + 'BB'),
            borderColor: sorted.map((_,i) => danoPalette[i % danoPalette.length]),
            borderWidth: 2, borderRadius: 6, borderSkipped: false,
          }]
        },
        options: {
          indexAxis: 'y',
          plugins: { legend:{display:false}, tooltip: EM_TT },
          scales: {
            x: { grid:{color:'rgba(255,255,255,.05)'}, ticks:{color:EM_C.muted,font:{family:'JetBrains Mono',size:11}} },
            y: { grid:{display:false}, ticks:{color:'#FAFAFA',font:{family:'Outfit',size:11}} }
          }
        }
      });
    }
  });
}

// ─────────────────────────────────────────────────────────────────
//  SECCIÓN 1 — TABLA COMPLETA DATÁFONOS
// ─────────────────────────────────────────────────────────────────
function _emRenderDfTable(panel) {
  const wrap = document.createElement('div');
  wrap.style.cssText = 'background:' + EM_C.card + ';border:1px solid ' + EM_C.border + ';border-radius:18px;overflow:hidden;margin-bottom:28px;box-shadow:0 4px 20px rgba(0,0,0,.4);';

  wrap.innerHTML = `
    <div style="display:flex;align-items:center;justify-content:space-between;padding:16px 20px;border-bottom:1px solid rgba(176,242,174,.08);">
      <div>
        <div style="font-family:'Syne',sans-serif;font-size:13px;font-weight:700;color:#f1f5f9;">🔍 Tabla Completa de Datáfonos en Bodega</div>
        <div style="font-size:11px;color:${EM_C.muted};margin-top:2px;" id="em-df-count">${EM_DF_ALL.length} registros</div>
      </div>
    </div>
    <div style="padding:10px 20px;border-bottom:1px solid rgba(176,242,174,.07);background:rgba(0,0,0,.15);display:flex;gap:10px;flex-wrap:wrap;">
      <input type="text" id="em-df-search-inp" placeholder="🔍 Buscar serial, referencia, bodega, estado, comentario..." oninput="emDfSearch(this.value)"
        style="flex:1;min-width:250px;background:rgba(255,255,255,.05);border:1px solid rgba(176,242,174,.2);border-radius:8px;color:#FAFAFA;padding:8px 14px;font-size:13px;font-family:'Outfit',sans-serif;outline:none;">
      <select id="em-df-filter-ref" onchange="emDfSearch(document.getElementById('em-df-search-inp').value)"
        style="background:#1a1916;border:1px solid rgba(176,242,174,.2);border-radius:8px;color:#FAFAFA;padding:8px 12px;font-size:12px;font-family:'Outfit',sans-serif;cursor:pointer;">
        <option value="">Todas las referencias</option>
        ${[...EM_DATAFONO_REFS].map(r=>`<option value="${r}">${r.replace(' - INGENICO','')}</option>`).join('')}
      </select>
      <select id="em-df-filter-est" onchange="emDfSearch(document.getElementById('em-df-search-inp').value)"
        style="background:#1a1916;border:1px solid rgba(176,242,174,.2);border-radius:8px;color:#FAFAFA;padding:8px 12px;font-size:12px;font-family:'Outfit',sans-serif;cursor:pointer;">
        <option value="">Todos los estados</option>
        <option>DISPONIBLE</option><option>ASOCIADO</option><option>EN DAÑO</option><option>DES. CIERRE</option><option>DES. INCIDENTE</option>
      </select>
    </div>
    <div id="em-df-table-wrap" style="overflow-x:auto;max-height:480px;"></div>
    <div style="display:flex;align-items:center;justify-content:space-between;padding:10px 20px;border-top:1px solid rgba(176,242,174,.07);">
      <span id="em-df-count-footer" style="font-size:11px;color:${EM_C.muted};font-family:'Outfit',sans-serif;"></span>
      <div id="em-df-pag" style="display:flex;gap:6px;flex-wrap:wrap;"></div>
    </div>`;

  panel.appendChild(wrap);
  emDfSearch('');
}

window.emDfSearch = function(q) {
  EM_DF_SEARCH = (q || '').toLowerCase().trim();
  EM_DF_PAGE = 1;
  _emRenderDfTableBody();
};

window.emDfGoPage = function(p) {
  EM_DF_PAGE = p;
  _emRenderDfTableBody();
};

function _emDfFiltered() {
  const refFilter = (document.getElementById('em-df-filter-ref')?.value || '');
  const estFilter = (document.getElementById('em-df-filter-est')?.value || '');
  return EM_DF_ALL.filter(r => {
    const ref = _emGetRef(r);
    const est = _emDfEstado(r);
    const serial = _emGetSerial(r);
    const bod    = _emGetBodega(r);
    const com    = _emGetComment(r);
    if (refFilter && ref !== refFilter) return false;
    if (estFilter && est !== estFilter) return false;
    if (EM_DF_SEARCH) {
      const haystack = [ref, serial, bod, com, est].join(' ').toLowerCase();
      if (!haystack.includes(EM_DF_SEARCH)) return false;
    }
    return true;
  });
}

function _emRenderDfTableBody() {
  const wrap = document.getElementById('em-df-table-wrap');
  const countEl = document.getElementById('em-df-count');
  const countFooter = document.getElementById('em-df-count-footer');
  if (!wrap) return;

  const filtered = _emDfFiltered();
  const total  = filtered.length;
  const pages  = Math.max(1, Math.ceil(total / EM_PAGE_SIZE));
  EM_DF_PAGE   = Math.min(EM_DF_PAGE, pages);
  const slice  = filtered.slice((EM_DF_PAGE-1)*EM_PAGE_SIZE, EM_DF_PAGE*EM_PAGE_SIZE);

  if (countEl) countEl.textContent = total + ' registros' + (EM_DF_SEARCH ? ' (filtrados)' : '');
  if (countFooter) countFooter.textContent = `Mostrando ${(EM_DF_PAGE-1)*EM_PAGE_SIZE+1}–${Math.min(EM_DF_PAGE*EM_PAGE_SIZE,total)} de ${total}`;

  const estColor = { 'DISPONIBLE': EM_C.verde, 'ASOCIADO': EM_C.azul, 'EN DAÑO': EM_C.danger, 'DES. CIERRE': EM_C.warn, 'DES. INCIDENTE': EM_C.orange };

  wrap.innerHTML = `<table style="width:100%;border-collapse:collapse;">
    <thead><tr style="background:rgba(0,0,0,.3);">
      <th style="padding:9px 14px;text-align:left;font-size:10px;font-weight:700;color:${EM_C.muted};text-transform:uppercase;letter-spacing:1px;font-family:'Outfit',sans-serif;white-space:nowrap;">Serial</th>
      <th style="padding:9px 14px;font-size:10px;font-weight:700;color:${EM_C.muted};text-transform:uppercase;letter-spacing:1px;font-family:'Outfit',sans-serif;white-space:nowrap;">Referencia</th>
      <th style="padding:9px 14px;font-size:10px;font-weight:700;color:${EM_C.muted};text-transform:uppercase;letter-spacing:1px;font-family:'Outfit',sans-serif;white-space:nowrap;">Bodega</th>
      <th style="padding:9px 14px;font-size:10px;font-weight:700;color:${EM_C.muted};text-transform:uppercase;letter-spacing:1px;font-family:'Outfit',sans-serif;white-space:nowrap;">Estado</th>
      <th style="padding:9px 14px;font-size:10px;font-weight:700;color:${EM_C.muted};text-transform:uppercase;letter-spacing:1px;font-family:'Outfit',sans-serif;white-space:nowrap;">Posición Depósito</th>
      <th style="padding:9px 14px;font-size:10px;font-weight:700;color:${EM_C.muted};text-transform:uppercase;letter-spacing:1px;font-family:'Outfit',sans-serif;white-space:nowrap;">Comentarios</th>
    </tr></thead>
    <tbody>
      ${slice.map((r,i) => {
        const est   = _emDfEstado(r);
        const color = estColor[est] || EM_C.muted;
        const serial= _emGetSerial(r) || '—';
        const ref   = _emGetRef(r).replace(' - INGENICO','');
        const bod   = _emGetBodega(r).replace('ALMACEN WOMPI ','').replace('ALMACEN ','');
        const pos   = _emGetPos(r) || '—';
        const com   = _emGetComment(r) || '—';
        return `<tr style="border-top:1px solid rgba(255,255,255,.04);background:${i%2?'transparent':'rgba(255,255,255,.015)'};">
          <td style="padding:8px 14px;font-family:'JetBrains Mono',monospace;font-size:11px;color:#e2e8f0;white-space:nowrap;">${serial}</td>
          <td style="padding:8px 14px;font-family:'Outfit',sans-serif;font-size:11px;color:#94a3b8;white-space:nowrap;">${ref}</td>
          <td style="padding:8px 14px;font-family:'Outfit',sans-serif;font-size:11px;color:#cbd5e1;white-space:nowrap;">${bod}</td>
          <td style="padding:8px 14px;"><span style="background:${color}22;color:${color};font-size:10px;font-weight:700;padding:2px 8px;border-radius:20px;font-family:'JetBrains Mono',monospace;letter-spacing:.3px;white-space:nowrap;">${est}</span></td>
          <td style="padding:8px 14px;font-family:'Outfit',sans-serif;font-size:11px;color:${EM_C.muted};white-space:nowrap;">${pos}</td>
          <td style="padding:8px 14px;font-family:'Outfit',sans-serif;font-size:11px;color:#94a3b8;max-width:240px;white-space:normal;line-height:1.4;">${com}</td>
        </tr>`;
      }).join('')}
      ${!slice.length ? `<tr><td colspan="6" style="text-align:center;padding:40px;color:${EM_C.muted};font-family:'Outfit',sans-serif;">Sin resultados para los filtros aplicados</td></tr>` : ''}
    </tbody>
  </table>`;

  _emMkPag('em-df-pag', EM_DF_PAGE, pages, 'emDfGoPage');
}

// ─────────────────────────────────────────────────────────────────
//  SECCIÓN 2 — KPIs SIMCARDS
// ─────────────────────────────────────────────────────────────────
function _emRenderSimKPIs(panel) {
  const rows    = EM_SIM_ALL;
  const total   = rows.length;
  const activas = rows.filter(r => _emSimEstado(r).activa).length;
  const sinAct  = total - activas;

  // Sitio con más SIMs
  const sitioMap = {};
  rows.forEach(r => {
    const s = _emGetBodega(r);
    sitioMap[s] = (sitioMap[s] || 0) + 1;
  });
  const topSitio = Object.entries(sitioMap).sort((a,b)=>b[1]-a[1])[0];

  const grid = document.createElement('div');
  grid.style.cssText = 'display:grid;grid-template-columns:repeat(auto-fill,minmax(200px,1fr));gap:14px;margin-bottom:28px;';

  const kpis = [
    { icon:'📡', label:'TOTAL SIMCARDS',    value: total,   sub: '100% del inventario', color: EM_C.teal },
    { icon:'✅', label:'ACTIVADAS',          value: activas, sub: pctStr(activas,total),  color: EM_C.verde },
    { icon:'⏸️', label:'SIN ACTIVAR',        value: sinAct,  sub: pctStr(sinAct,total),   color: EM_C.warn },
    { icon:'🏆', label:'SITIO TOP',          value: topSitio ? topSitio[1] : 0, sub: topSitio ? topSitio[0].replace('ALMACEN WOMPI ','').replace('ALMACEN ','').substring(0,18) : '—', color: EM_C.purple },
  ];
  kpis.forEach(k => grid.appendChild(_emKpiCard(k.icon, k.label, k.value.toLocaleString('es-CO'), k.sub, k.color)));
  panel.appendChild(grid);
}

// ─────────────────────────────────────────────────────────────────
//  SECCIÓN 2 — GRÁFICAS SIMCARDS
// ─────────────────────────────────────────────────────────────────
function _emRenderSimCharts(panel) {
  const rows = EM_SIM_ALL;

  // Donut activas/sin activar
  const activas = rows.filter(r => _emSimEstado(r).activa).length;
  const sinAct  = rows.length - activas;

  // Top 10 sitios por cantidad de SIMs
  const sitioMap = {};
  rows.forEach(r => {
    const s = _emGetBodega(r) || 'Sin ubicación';
    sitioMap[s] = (sitioMap[s] || 0) + 1;
  });
  const topSitios = Object.entries(sitioMap).sort((a,b)=>b[1]-a[1]).slice(0,12);

  // Activaciones por mes
  const mesMap = {};
  rows.forEach(r => {
    const { activa, fecha } = _emSimEstado(r);
    if (!activa || !fecha) return;
    const parts = fecha.split('/');
    if (parts.length < 3) return;
    const key = parts[2] + '-' + parts[1].padStart(2,'0');
    mesMap[key] = (mesMap[key] || 0) + 1;
  });
  const mesSorted = Object.entries(mesMap).sort((a,b) => a[0].localeCompare(b[0]));

  const chartsGrid = document.createElement('div');
  chartsGrid.style.cssText = 'display:grid;grid-template-columns:repeat(3,1fr);gap:18px;margin-bottom:28px;';

  const c1 = _emCard('em-c-sim-dona',    'Activadas vs Sin Activar',         '% del total de SIMcards',              EM_C.teal);
  const c2 = _emCard('em-c-sim-sitios',  'Top Sitios por Cantidad de SIMs',  'Ubicaciones con mayor stock de SIMs',  EM_C.purple);
  const c3 = _emCard('em-c-sim-meses',   'Activaciones por Mes',             'Evolución de activaciones en el tiempo', EM_C.verde);
  c3.style.gridColumn = 'span 3';

  chartsGrid.appendChild(c1);
  chartsGrid.appendChild(c2);
  panel.appendChild(chartsGrid);

  // Chart de activaciones (full width)
  const chartLine = document.createElement('div');
  chartLine.style.cssText = 'margin-bottom:28px;';
  chartLine.appendChild(c3);
  panel.appendChild(chartLine);

  requestAnimationFrame(() => {
    // Donut activadas
    _emDestroyChart('em-c-sim-dona');
    const d1 = document.getElementById('em-c-sim-dona');
    if (d1) {
      EM_CHARTS['em-c-sim-dona'] = new Chart(d1, {
        type: 'doughnut',
        data: {
          labels: ['Activadas', 'Sin Activar'],
          datasets: [{
            data: [activas, sinAct],
            backgroundColor: [EM_C.verde+'CC', EM_C.warn+'CC'],
            borderColor: 'rgba(0,0,0,0)', borderWidth: 0, hoverOffset: 8,
          }]
        },
        options: {
          cutout: '68%',
          plugins: {
            legend: EM_LEG,
            tooltip: Object.assign({}, EM_TT, {
              callbacks: { label: ctx => ' ' + ctx.label + ': ' + ctx.parsed.toLocaleString('es-CO') + ' SIMs (' + pctStr(ctx.parsed, rows.length) + ')' }
            })
          }
        }
      });
    }

    // Bar sitios
    _emDestroyChart('em-c-sim-sitios');
    const d2 = document.getElementById('em-c-sim-sitios');
    if (d2) {
      EM_CHARTS['em-c-sim-sitios'] = new Chart(d2, {
        type: 'bar',
        data: {
          labels: topSitios.map(e => e[0].replace('ALMACEN WOMPI ','').replace('ALMACEN ','')),
          datasets: [{
            label: 'SIMcards',
            data: topSitios.map(e => e[1]),
            backgroundColor: EM_C.purple + 'BB',
            borderColor: EM_C.purple,
            borderWidth: 2, borderRadius: 6,
          }]
        },
        options: {
          indexAxis: 'y',
          plugins: { legend:{display:false}, tooltip: EM_TT },
          scales: {
            x: { grid:{color:'rgba(255,255,255,.05)'}, ticks:{color:EM_C.muted,font:{family:'JetBrains Mono',size:10}} },
            y: { grid:{display:false}, ticks:{color:'#FAFAFA',font:{family:'Outfit',size:10}} }
          }
        }
      });
    }

    // Line activaciones por mes
    _emDestroyChart('em-c-sim-meses');
    const d3 = document.getElementById('em-c-sim-meses');
    if (d3 && mesSorted.length) {
      EM_CHARTS['em-c-sim-meses'] = new Chart(d3, {
        type: 'bar',
        data: {
          labels: mesSorted.map(e => {
            const [y,m] = e[0].split('-');
            const meses = ['','Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
            return meses[parseInt(m)] + ' ' + y;
          }),
          datasets: [{
            label: 'Activaciones',
            data: mesSorted.map(e => e[1]),
            backgroundColor: EM_C.verde + '88',
            borderColor: EM_C.verde,
            borderWidth: 2, borderRadius: 4,
            fill: true,
          }]
        },
        options: {
          plugins: { legend:{display:false}, tooltip: EM_TT },
          scales: {
            x: { grid:{display:false}, ticks:{color:EM_C.muted,font:{family:'Outfit',size:11}} },
            y: { grid:{color:'rgba(255,255,255,.05)'}, ticks:{color:EM_C.muted,font:{family:'JetBrains Mono',size:11}} }
          }
        }
      });
    }
  });
}

// ─────────────────────────────────────────────────────────────────
//  SECCIÓN 2 — TABLA SIMCARDS
// ─────────────────────────────────────────────────────────────────
function _emRenderSimTable(panel) {
  const wrap = document.createElement('div');
  wrap.style.cssText = 'background:' + EM_C.card + ';border:1px solid rgba(103,232,249,.15);border-top:2px solid ' + EM_C.teal + ';border-radius:18px;overflow:hidden;margin-bottom:32px;box-shadow:0 4px 20px rgba(0,0,0,.4);';

  wrap.innerHTML = `
    <div style="display:flex;align-items:center;justify-content:space-between;padding:16px 20px;border-bottom:1px solid rgba(103,232,249,.08);">
      <div>
        <div style="font-family:'Syne',sans-serif;font-size:13px;font-weight:700;color:#f1f5f9;">🔍 Tabla Completa de SIMcards</div>
        <div style="font-size:11px;color:${EM_C.muted};margin-top:2px;" id="em-sim-count">${EM_SIM_ALL.length} registros</div>
      </div>
    </div>
    <div style="padding:10px 20px;border-bottom:1px solid rgba(103,232,249,.07);background:rgba(0,0,0,.15);display:flex;gap:10px;flex-wrap:wrap;">
      <input type="text" id="em-sim-search-inp" placeholder="🔍 Buscar serial, sitio, atributos, fecha activación..."
        oninput="emSimSearch(this.value)"
        style="flex:1;min-width:250px;background:rgba(255,255,255,.05);border:1px solid rgba(103,232,249,.2);border-radius:8px;color:#FAFAFA;padding:8px 14px;font-size:13px;font-family:'Outfit',sans-serif;outline:none;">
      <select id="em-sim-filter-est" onchange="emSimSearch(document.getElementById('em-sim-search-inp').value)"
        style="background:#1a1916;border:1px solid rgba(103,232,249,.2);border-radius:8px;color:#FAFAFA;padding:8px 12px;font-size:12px;font-family:'Outfit',sans-serif;cursor:pointer;">
        <option value="">Todas (activadas + sin activar)</option>
        <option value="activa">✅ Activadas</option>
        <option value="sin">⏸️ Sin Activar</option>
      </select>
    </div>
    <div id="em-sim-table-wrap" style="overflow-x:auto;max-height:480px;"></div>
    <div style="display:flex;align-items:center;justify-content:space-between;padding:10px 20px;border-top:1px solid rgba(103,232,249,.07);">
      <span id="em-sim-count-footer" style="font-size:11px;color:${EM_C.muted};font-family:'Outfit',sans-serif;"></span>
      <div id="em-sim-pag" style="display:flex;gap:6px;flex-wrap:wrap;"></div>
    </div>`;

  panel.appendChild(wrap);
  emSimSearch('');
}

window.emSimSearch = function(q) {
  EM_SIM_SEARCH = (q || '').toLowerCase().trim();
  EM_SIM_PAGE = 1;
  _emRenderSimTableBody();
};

window.emSimGoPage = function(p) {
  EM_SIM_PAGE = p;
  _emRenderSimTableBody();
};

function _emSimFiltered() {
  const estFilter = (document.getElementById('em-sim-filter-est')?.value || '');
  return EM_SIM_ALL.filter(r => {
    const { activa } = _emSimEstado(r);
    if (estFilter === 'activa' && !activa) return false;
    if (estFilter === 'sin'    &&  activa) return false;
    if (EM_SIM_SEARCH) {
      const serial = _emGetSerial(r);
      const bod    = _emGetBodega(r);
      const attr   = _emGetAtributos(r);
      const haystack = [serial, bod, attr].join(' ').toLowerCase();
      if (!haystack.includes(EM_SIM_SEARCH)) return false;
    }
    return true;
  });
}

function _emRenderSimTableBody() {
  const wrap = document.getElementById('em-sim-table-wrap');
  const countEl = document.getElementById('em-sim-count');
  const footer  = document.getElementById('em-sim-count-footer');
  if (!wrap) return;

  const filtered = _emSimFiltered();
  const total  = filtered.length;
  const pages  = Math.max(1, Math.ceil(total / EM_PAGE_SIZE));
  EM_SIM_PAGE  = Math.min(EM_SIM_PAGE, pages);
  const slice  = filtered.slice((EM_SIM_PAGE-1)*EM_PAGE_SIZE, EM_SIM_PAGE*EM_PAGE_SIZE);

  if (countEl) countEl.textContent = total + ' registros';
  if (footer)  footer.textContent  = `Mostrando ${(EM_SIM_PAGE-1)*EM_PAGE_SIZE+1}–${Math.min(EM_SIM_PAGE*EM_PAGE_SIZE,total)} de ${total}`;

  wrap.innerHTML = `<table style="width:100%;border-collapse:collapse;">
    <thead><tr style="background:rgba(0,0,0,.3);">
      <th style="padding:9px 14px;text-align:left;font-size:10px;font-weight:700;color:${EM_C.muted};text-transform:uppercase;letter-spacing:1px;font-family:'Outfit',sans-serif;white-space:nowrap;">Serial</th>
      <th style="padding:9px 14px;font-size:10px;font-weight:700;color:${EM_C.muted};text-transform:uppercase;letter-spacing:1px;font-family:'Outfit',sans-serif;white-space:nowrap;">Nombre de Sitio / Ubicación</th>
      <th style="padding:9px 14px;font-size:10px;font-weight:700;color:${EM_C.muted};text-transform:uppercase;letter-spacing:1px;font-family:'Outfit',sans-serif;white-space:nowrap;">Estado</th>
      <th style="padding:9px 14px;font-size:10px;font-weight:700;color:${EM_C.muted};text-transform:uppercase;letter-spacing:1px;font-family:'Outfit',sans-serif;white-space:nowrap;">Fecha Activación</th>
      <th style="padding:9px 14px;font-size:10px;font-weight:700;color:${EM_C.muted};text-transform:uppercase;letter-spacing:1px;font-family:'Outfit',sans-serif;white-space:nowrap;">Atributos</th>
    </tr></thead>
    <tbody>
      ${slice.map((r,i) => {
        const { activa, fecha } = _emSimEstado(r);
        const serial = _emGetSerial(r) || '—';
        const bod    = _emGetBodega(r) || '—';
        const attr   = _emGetAtributos(r) || '—';
        const colorEst = activa ? EM_C.verde : EM_C.warn;
        const label    = activa ? 'ACTIVADA' : 'SIN ACTIVAR';
        return `<tr style="border-top:1px solid rgba(255,255,255,.04);background:${i%2?'transparent':'rgba(255,255,255,.015)'};">
          <td style="padding:8px 14px;font-family:'JetBrains Mono',monospace;font-size:11px;color:#e2e8f0;white-space:nowrap;">${serial}</td>
          <td style="padding:8px 14px;font-family:'Outfit',sans-serif;font-size:11px;color:#cbd5e1;">${bod}</td>
          <td style="padding:8px 14px;"><span style="background:${colorEst}22;color:${colorEst};font-size:10px;font-weight:700;padding:2px 8px;border-radius:20px;font-family:'JetBrains Mono',monospace;white-space:nowrap;">${label}</span></td>
          <td style="padding:8px 14px;font-family:'JetBrains Mono',monospace;font-size:11px;color:${fecha ? EM_C.teal : EM_C.muted};white-space:nowrap;">${fecha || '—'}</td>
          <td style="padding:8px 14px;font-family:'Outfit',sans-serif;font-size:10px;color:#94a3b8;max-width:260px;white-space:normal;line-height:1.4;">${attr}</td>
        </tr>`;
      }).join('')}
      ${!slice.length ? `<tr><td colspan="5" style="text-align:center;padding:40px;color:${EM_C.muted};font-family:'Outfit',sans-serif;">Sin resultados para los filtros aplicados</td></tr>` : ''}
    </tbody>
  </table>`;

  _emMkPag('em-sim-pag', EM_SIM_PAGE, pages, 'emSimGoPage');
}

// ─────────────────────────────────────────────────────────────────
//  EXPOSE
// ─────────────────────────────────────────────────────────────────
window.renderEstadoMateriales = renderEstadoMateriales;