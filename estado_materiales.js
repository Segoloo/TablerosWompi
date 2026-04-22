/**
 * ╔══════════════════════════════════════════════════════════════════╗
 * ║  estado_materiales.js — Tab "Estado de Materiales"             ║
 * ║  Datáfonos en bodega · SIMCards · Análisis de estado           ║
 * ╚══════════════════════════════════════════════════════════════════╝
 */

'use strict';

// ══════════════════════════════════════════════════════════════════
//  CONSTANTES
// ══════════════════════════════════════════════════════════════════

const EM_REFS_DATAFONO = new Set([
  'PINPAD DESK 1700 - INGENICO',
  'DATAFONO EX6000 - INGENICO',
  'DATAFONO EX4000 - INGENICO',
  'DATAFONO DX4000 PORTATIL - INGENICO',
  'DATAFONO DX4000 ESCRITORIO - INGENICO',
]);

const EM_REF_SIMCARD = 'SIM LTE MULTI 256K P17 POST QR - CLARO';

const EM_ESTADOS_DEPOSITO = new Set([
  'EN DAÑO',
  'DESINSTALADO-INCIDENTE',
  'DESINSTALADO-CIERRE',
]);

const EM_COLORS = {
  danio:       '#FF5C5C',
  incidente:   '#FFC04D',
  cierre:      '#99D1FC',
  asociado:    '#C084FC',
  disponible:  '#B0F2AE',
  activada:    '#B0F2AE',
  sinActivar:  '#FF5C5C',
  palette: ['#B0F2AE','#99D1FC','#DFFF61','#C084FC','#FFC04D','#FF5C5C','#F49D6E','#7B8CDE','#00C87A','#A8E6CF'],
};

// ── Estado módulo ─────────────────────────────────────────────────
let EM_CHARTS = {};
let EM_DF_ALL   = [];   // datáfonos en bodegas
let EM_SIM_ALL  = [];   // simcards
let EM_DF_PAGE  = 1;
let EM_SIM_PAGE = 1;
const EM_PAGE_SIZE = 50;

// ── Filtro de búsqueda tabla datáfonos ───────────────────────────
let EM_DF_SEARCH = '';
let EM_DF_FILTERED = [];
let EM_SIM_SEARCH  = '';
let EM_SIM_FILTERED = [];

// ══════════════════════════════════════════════════════════════════
//  HELPERS
// ══════════════════════════════════════════════════════════════════

function _emGetRaw() {
  if (window.INV_RAW && window.INV_RAW.length) return window.INV_RAW;
  if (typeof window.invGetRaw === 'function') return window.invGetRaw();
  return [];
}

function _emDestroyChart(id) {
  if (EM_CHARTS[id]) { try { EM_CHARTS[id].destroy(); } catch(e){} delete EM_CHARTS[id]; }
}

/** Determina el estado de un datáfono en bodega */
function _emEstadoDatafono(row) {
  const pos  = (row['Posición en depósito'] || '').trim().toUpperCase();
  const com  = (row['Comentarios'] || '').trim();

  if (pos === 'EN DAÑO')                 return 'DAÑADO';
  if (pos === 'DESINSTALADO-INCIDENTE')  return 'DES. INCIDENTE';
  if (pos === 'DESINSTALADO-CIERRE')     return 'DES. CIERRE';

  // Determinar asociado / disponible a partir de comentarios
  if (!com) return 'DISPONIBLE';
  const parts = com.split('|');
  const numero = (parts[0] || '').trim();
  if (numero === '99999') return 'DISPONIBLE';
  if (/^\d+$/.test(numero) && numero !== '99999') return 'ASOCIADO';
  return 'DISPONIBLE';
}

/** Extrae el tipo de daño del campo Comentarios (para datáfonos EN DAÑO) */
function _emTipoDanio(row) {
  const com = (row['Comentarios'] || '').trim();
  if (!com) return 'SIN DESCRIPCIÓN';
  const parts = com.split('|');
  if (parts.length >= 2) {
    return (parts[1] || '').trim().toUpperCase() || 'SIN DESCRIPCIÓN';
  }
  return 'SIN DESCRIPCIÓN';
}

/** Determina si una simcard está activada */
function _emSimActivada(row) {
  const attr = (row['Atributos'] || '').toUpperCase();
  return /FA:\d{2}\/\d{2}\/\d{4}/.test(attr);
}

/** Extrae la fecha de activación de una simcard */
function _emSimFechaActivacion(row) {
  const attr = row['Atributos'] || '';
  const m = attr.match(/FA:(\d{2}\/\d{2}\/\d{4})/i);
  return m ? m[1] : '—';
}

/** Crea un chart donut estándar */
function _emDonut(canvasId, labels, data, colors, title) {
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;
  _emDestroyChart(canvasId);
  const total = data.reduce((a, b) => a + b, 0);
  EM_CHARTS[canvasId] = new Chart(canvas, {
    type: 'doughnut',
    data: {
      labels,
      datasets: [{ data, backgroundColor: colors, borderWidth: 2, borderColor: '#181715', hoverBorderColor: '#fff', hoverOffset: 6 }]
    },
    options: {
      cutout: '65%',
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: 'bottom', labels: { color: '#cbd5e1', font: { size: 11, family: "'Outfit', sans-serif" }, padding: 12, boxWidth: 12 } },
        tooltip: {
          backgroundColor: 'rgba(24,23,21,.95)',
          titleColor: '#B0F2AE', bodyColor: '#FAFAFA',
          borderColor: 'rgba(176,242,174,.2)', borderWidth: 1, padding: 12,
          callbacks: { label: ctx => ` ${ctx.label}: ${ctx.parsed.toLocaleString('es-CO')} (${total ? (ctx.parsed/total*100).toFixed(1) : 0}%)` }
        },
        ...(title ? { title: { display: false } } : {}),
      }
    }
  });
}

/** Crea un chart de barras horizontal */
function _emBarH(canvasId, labels, data, colors, label) {
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;
  _emDestroyChart(canvasId);
  const maxVal = Math.max(...data, 1);
  EM_CHARTS[canvasId] = new Chart(canvas, {
    type: 'bar',
    data: {
      labels,
      datasets: [{ label, data,
        backgroundColor: Array.isArray(colors) ? colors : labels.map((_,i) => EM_COLORS.palette[i % EM_COLORS.palette.length] + 'BB'),
        borderColor:     Array.isArray(colors) ? colors.map(c => c) : labels.map((_,i) => EM_COLORS.palette[i % EM_COLORS.palette.length]),
        borderWidth: 1, borderRadius: 6,
      }]
    },
    options: {
      indexAxis: 'y',
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: { backgroundColor: 'rgba(24,23,21,.95)', titleColor: '#B0F2AE', bodyColor: '#FAFAFA', borderColor: 'rgba(176,242,174,.2)', borderWidth: 1, padding: 10 }
      },
      scales: {
        x: { ticks: { color: '#7A7674', font: { size: 10 } }, grid: { color: 'rgba(255,255,255,.05)' } },
        y: { ticks: { color: '#e2e8f0', font: { size: 11, family: "'Outfit', sans-serif" } }, grid: { display: false } }
      }
    }
  });
}

/** Crea chart de barras agrupadas */
function _emBarGrouped(canvasId, labels, datasets) {
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;
  _emDestroyChart(canvasId);
  EM_CHARTS[canvasId] = new Chart(canvas, {
    type: 'bar',
    data: { labels, datasets },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { labels: { color: '#cbd5e1', font: { size: 11 }, padding: 10, boxWidth: 12 } },
        tooltip: { backgroundColor: 'rgba(24,23,21,.95)', titleColor: '#B0F2AE', bodyColor: '#FAFAFA', borderColor: 'rgba(176,242,174,.2)', borderWidth: 1, padding: 10 }
      },
      scales: {
        x: { ticks: { color: '#cbd5e1', font: { size: 10 }, maxRotation: 30 }, grid: { display: false } },
        y: { ticks: { color: '#7A7674' }, grid: { color: 'rgba(255,255,255,.05)' } }
      }
    }
  });
}

// ══════════════════════════════════════════════════════════════════
//  PROCESAMIENTO DE DATOS
// ══════════════════════════════════════════════════════════════════

function _emProcesar() {
  const raw = _emGetRaw();
  if (!raw || !raw.length) { EM_DF_ALL = []; EM_SIM_ALL = []; return; }

  const bodegas = typeof INV_BODEGAS !== 'undefined' ? INV_BODEGAS : window.INV_BODEGAS;

  // ── Datáfonos en bodegas ─────────────────────────────────────
  EM_DF_ALL = raw.filter(r => {
    const nombre = (r['Nombre'] || '').trim();
    const ubic   = (r['Nombre de la ubicación'] || '').trim();
    return EM_REFS_DATAFONO.has(nombre) && bodegas && bodegas.has(ubic);
  });

  // ── SIMCards ─────────────────────────────────────────────────
  EM_SIM_ALL = raw.filter(r => {
    const nombre = (r['Nombre'] || '').trim();
    return nombre === EM_REF_SIMCARD;
  });

  EM_DF_FILTERED  = EM_DF_ALL.slice();
  EM_SIM_FILTERED = EM_SIM_ALL.slice();
}

// ══════════════════════════════════════════════════════════════════
//  SECCIÓN DATÁFONOS
// ══════════════════════════════════════════════════════════════════

function _emRenderDatafonos() {
  const data = EM_DF_ALL;

  // ── KPIs resumen ─────────────────────────────────────────────
  const counts = { DISPONIBLE: 0, ASOCIADO: 0, DAÑADO: 0, 'DES. INCIDENTE': 0, 'DES. CIERRE': 0 };
  data.forEach(r => { const e = _emEstadoDatafono(r); counts[e] = (counts[e]||0) + 1; });
  const total = data.length;

  const kpiEl = document.getElementById('em-df-kpis');
  if (kpiEl) {
    const kpis = [
      { label: 'Total en Bodegas', val: total, color: '#DFFF61', icon: '📦' },
      { label: 'Disponibles', val: counts['DISPONIBLE'], color: '#B0F2AE', icon: '✅' },
      { label: 'Asociados', val: counts['ASOCIADO'], color: '#C084FC', icon: '🔗' },
      { label: 'En Daño', val: counts['DAÑADO'], color: '#FF5C5C', icon: '🔴' },
      { label: 'Des. Incidente', val: counts['DES. INCIDENTE'], color: '#FFC04D', icon: '⚠️' },
      { label: 'Des. Cierre', val: counts['DES. CIERRE'], color: '#99D1FC', icon: '🔵' },
    ];
    kpiEl.innerHTML = kpis.map(k => `
      <div style="background:linear-gradient(145deg,rgba(10,26,18,.95),rgba(8,20,14,.9));
        border:1px solid ${k.color}22;border-top:2px solid ${k.color};border-radius:14px;
        padding:18px 20px;min-width:130px;flex:1;">
        <div style="font-size:20px;margin-bottom:6px">${k.icon}</div>
        <div style="font-family:'Outfit',sans-serif;font-size:11px;color:#7A7674;letter-spacing:.4px;text-transform:uppercase;margin-bottom:6px">${k.label}</div>
        <div style="font-family:'JetBrains Mono',monospace;font-size:26px;font-weight:700;color:${k.color};line-height:1;text-shadow:0 0 20px ${k.color}55">${k.val.toLocaleString('es-CO')}</div>
        <div style="font-size:10px;color:#64748b;margin-top:4px">${total ? (k.val/total*100).toFixed(1) : 0}% del total</div>
      </div>`).join('');
  }

  // ── Gráfica 1: Dona — estados depósito de interés ──────────
  const depositoCounts = { 'EN DAÑO': 0, 'DESINSTALADO-INCIDENTE': 0, 'DESINSTALADO-CIERRE': 0 };
  data.forEach(r => {
    const pos = (r['Posición en depósito'] || '').trim().toUpperCase();
    if (EM_ESTADOS_DEPOSITO.has(pos)) depositoCounts[pos]++;
  });
  const donaLabels  = Object.keys(depositoCounts);
  const donaData    = Object.values(depositoCounts);
  const donaColors  = [EM_COLORS.danio, EM_COLORS.incidente, EM_COLORS.cierre];
  _emDonut('em-chart-deposito', donaLabels, donaData, donaColors, 'Posición Depósito');

  // ── Gráfica 2: Barras — por referencia × estado depósito ───
  const refStates = {};
  [...EM_REFS_DATAFONO].forEach(ref => { refStates[ref] = { 'EN DAÑO': 0, 'DESINSTALADO-INCIDENTE': 0, 'DESINSTALADO-CIERRE': 0 }; });
  data.forEach(r => {
    const nombre = (r['Nombre'] || '').trim();
    const pos    = (r['Posición en depósito'] || '').trim().toUpperCase();
    if (refStates[nombre] && EM_ESTADOS_DEPOSITO.has(pos)) refStates[nombre][pos]++;
  });

  // Abreviar nombres de referencias
  const _abbrev = n => n.replace('DATAFONO ','').replace(' - INGENICO','').replace('PORTATIL','PORT.').replace('ESCRITORIO','ESC.');
  const refLabels = [...EM_REFS_DATAFONO].map(_abbrev);
  const estadosKeys = ['EN DAÑO','DESINSTALADO-INCIDENTE','DESINSTALADO-CIERRE'];
  const estadosColors = [EM_COLORS.danio, EM_COLORS.incidente, EM_COLORS.cierre];
  const estadosLabels = ['En Daño', 'Des. Incidente', 'Des. Cierre'];

  const datasets = estadosKeys.map((est, i) => ({
    label: estadosLabels[i],
    data:  [...EM_REFS_DATAFONO].map(ref => refStates[ref][est]),
    backgroundColor: estadosColors[i] + 'BB',
    borderColor:     estadosColors[i],
    borderWidth: 1, borderRadius: 5,
  }));
  _emBarGrouped('em-chart-ref-estado', refLabels, datasets);

  // ── Gráfica 3: Dona — tipos de daño ─────────────────────────
  const danioRows = data.filter(r => (r['Posición en depósito']||'').trim().toUpperCase() === 'EN DAÑO');
  const danioMap  = {};
  danioRows.forEach(r => {
    const tipo = _emTipoDanio(r);
    danioMap[tipo] = (danioMap[tipo] || 0) + 1;
  });
  const danioSorted = Object.entries(danioMap).sort((a,b) => b[1]-a[1]);
  _emBarH('em-chart-tipos-danio',
    danioSorted.map(d => d[0]),
    danioSorted.map(d => d[1]),
    danioSorted.map((_,i) => EM_COLORS.palette[i % EM_COLORS.palette.length] + 'BB'),
    'Cantidad'
  );

  // ── Tabla por bodega ─────────────────────────────────────────
  const bodegas = typeof INV_BODEGAS !== 'undefined' ? INV_BODEGAS : window.INV_BODEGAS;
  const bodegaMap = {};
  data.forEach(r => {
    const b = (r['Nombre de la ubicación'] || 'Sin bodega').trim();
    if (!bodegaMap[b]) bodegaMap[b] = { DISPONIBLE:0, ASOCIADO:0, DAÑADO:0, 'DES. CIERRE':0, 'DES. INCIDENTE':0 };
    const est = _emEstadoDatafono(r);
    bodegaMap[b][est] = (bodegaMap[b][est] || 0) + 1;
  });
  const bodegaRows = Object.entries(bodegaMap)
    .map(([nombre, c]) => ({ nombre, ...c, total: Object.values(c).reduce((a,b)=>a+b,0) }))
    .sort((a,b) => b.total - a.total);

  const bodTablaEl = document.getElementById('em-tabla-bodegas-wrap');
  if (bodTablaEl) {
    if (!bodegaRows.length) {
      bodTablaEl.innerHTML = '<div style="text-align:center;padding:32px;color:#7A7674;font-family:\'Outfit\',sans-serif;">Sin datos</div>';
    } else {
      bodTablaEl.innerHTML = `<table style="width:100%;border-collapse:collapse;font-size:12px;">
        <thead>
          <tr style="background:rgba(0,0,0,.3)">
            ${['Bodega','Disponible','Asociado','Dañado','Des. Cierre','Des. Incidente','Total'].map(h =>
              `<th style="padding:10px 14px;text-align:left;font-family:'Syne',sans-serif;font-size:10px;font-weight:700;
                color:#B0F2AE;letter-spacing:1px;text-transform:uppercase;white-space:nowrap;
                border-bottom:1px solid rgba(176,242,174,.15);">${h}</th>`
            ).join('')}
          </tr>
        </thead>
        <tbody>
          ${bodegaRows.map((b, i) => {
            const bg = i%2===0 ? 'rgba(176,242,174,.012)' : 'transparent';
            const _pill = (n, color) => `<span style="background:${color}22;color:${color};font-family:'JetBrains Mono',monospace;
              font-size:12px;font-weight:700;padding:3px 10px;border-radius:10px;display:inline-block;">${n}</span>`;
            return `<tr style="background:${bg}" onmouseover="this.style.background='rgba(176,242,174,.04)'" onmouseout="this.style.background='${bg}'">
              <td style="padding:9px 14px;color:#e2e8f0;font-size:11px;max-width:260px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="${b.nombre}">${b.nombre}</td>
              <td style="padding:9px 14px;text-align:center">${_pill(b['DISPONIBLE']||0, '#B0F2AE')}</td>
              <td style="padding:9px 14px;text-align:center">${_pill(b['ASOCIADO']||0, '#C084FC')}</td>
              <td style="padding:9px 14px;text-align:center">${_pill(b['DAÑADO']||0, '#FF5C5C')}</td>
              <td style="padding:9px 14px;text-align:center">${_pill(b['DES. CIERRE']||0, '#99D1FC')}</td>
              <td style="padding:9px 14px;text-align:center">${_pill(b['DES. INCIDENTE']||0, '#FFC04D')}</td>
              <td style="padding:9px 14px;text-align:center"><span style="font-family:'JetBrains Mono',monospace;font-weight:700;color:#DFFF61;font-size:13px;">${b.total}</span></td>
            </tr>`;
          }).join('')}
        </tbody>
      </table>`;
    }
  }

  // ── Tabla completa datáfonos ─────────────────────────────────
  _emApplyDfSearch('');
}

function _emApplyDfSearch(query) {
  EM_DF_SEARCH = (query || '').toLowerCase().trim();
  EM_DF_FILTERED = EM_DF_ALL.filter(r => {
    if (!EM_DF_SEARCH) return true;
    const serial = (r['Número de serie']||r['Numero de serie']||r['Serial']||'').toLowerCase();
    const nombre = (r['Nombre']||'').toLowerCase();
    const ubic   = (r['Nombre de la ubicación']||'').toLowerCase();
    const pos    = (r['Posición en depósito']||'').toLowerCase();
    const com    = (r['Comentarios']||'').toLowerCase();
    return serial.includes(EM_DF_SEARCH) || nombre.includes(EM_DF_SEARCH) || ubic.includes(EM_DF_SEARCH) || pos.includes(EM_DF_SEARCH) || com.includes(EM_DF_SEARCH);
  });
  EM_DF_PAGE = 1;
  _emRenderDfTabla();
}

function _emRenderDfTabla() {
  const wrap  = document.getElementById('em-df-tabla-wrap');
  const count = document.getElementById('em-df-tabla-count');
  const pagEl = document.getElementById('em-df-tabla-pag');
  if (!wrap) return;

  const data  = EM_DF_FILTERED;
  const pages = Math.max(1, Math.ceil(data.length / EM_PAGE_SIZE));
  const slice = data.slice((EM_DF_PAGE-1)*EM_PAGE_SIZE, EM_DF_PAGE*EM_PAGE_SIZE);

  if (count) count.textContent = `${data.length.toLocaleString('es-CO')} registros`;

  if (!slice.length) {
    wrap.innerHTML = '<div style="text-align:center;padding:40px;color:#7A7674;font-family:\'Outfit\',sans-serif;">Sin resultados</div>';
    if (pagEl) pagEl.innerHTML = '';
    return;
  }

  const COLS = [
    { h:'Serial',         fn: r => r['Número de serie']||r['Numero de serie']||r['Serial']||'—', mono:true, color:'#a5f3fc' },
    { h:'Referencia',     fn: r => r['Nombre']||'—', color:'#e2e8f0' },
    { h:'Bodega',         fn: r => r['Nombre de la ubicación']||'—', color:'#cbd5e1' },
    { h:'Posición Dep.',  fn: r => r['Posición en depósito']||'—', color:'#94a3b8' },
    { h:'Estado',         fn: r => _emEstadoDatafono(r), isEstado: true },
    { h:'Comentarios',    fn: r => r['Comentarios']||'—', color:'#94a3b8' },
  ];

  const _estadoStyle = est => {
    const m = { 'DISPONIBLE': ['#B0F2AE','rgba(176,242,174,.12)'], 'ASOCIADO': ['#C084FC','rgba(192,132,252,.12)'],
      'DAÑADO': ['#FF5C5C','rgba(255,92,92,.12)'], 'DES. INCIDENTE': ['#FFC04D','rgba(255,192,77,.12)'], 'DES. CIERRE': ['#99D1FC','rgba(153,209,252,.12)'] };
    const [col, bg] = m[est] || ['#7A7674','rgba(255,255,255,.06)'];
    return `background:${bg};color:${col};font-size:10px;font-weight:700;padding:3px 10px;border-radius:12px;
      border:1px solid ${col}44;display:inline-block;font-family:'Outfit',sans-serif;white-space:nowrap;`;
  };

  wrap.innerHTML = `<table style="width:100%;border-collapse:collapse;font-size:12px;">
    <thead>
      <tr style="background:rgba(0,0,0,.3);">
        ${COLS.map(c => `<th style="padding:10px 14px;text-align:left;font-family:'Syne',sans-serif;font-size:10px;
          font-weight:700;color:#B0F2AE;letter-spacing:1px;text-transform:uppercase;
          border-bottom:1px solid rgba(176,242,174,.15);white-space:nowrap;">${c.h}</th>`).join('')}
      </tr>
    </thead>
    <tbody>
      ${slice.map((r, i) => {
        const bg = i%2===0 ? 'rgba(176,242,174,.012)' : 'transparent';
        return `<tr style="background:${bg}" onmouseover="this.style.background='rgba(176,242,174,.04)'" onmouseout="this.style.background='${bg}'">
          ${COLS.map(c => {
            const v = c.fn(r);
            if (c.isEstado) return `<td style="padding:9px 14px"><span style="${_estadoStyle(v)}">${v}</span></td>`;
            const st = `padding:9px 14px;${c.mono?`font-family:'JetBrains Mono',monospace;font-size:11px;`:'font-size:11px;'}color:${c.color||'#e2e8f0'};max-width:220px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;`;
            return `<td style="${st}" title="${String(v).replace(/"/g,'&quot;')}">${v}</td>`;
          }).join('')}
        </tr>`;
      }).join('')}
    </tbody>
  </table>`;

  _emRenderPag(pagEl, EM_DF_PAGE, pages, p => { EM_DF_PAGE = p; _emRenderDfTabla(); }, '#B0F2AE');
}

// ══════════════════════════════════════════════════════════════════
//  SECCIÓN SIMCARDS
// ══════════════════════════════════════════════════════════════════

function _emRenderSimcards() {
  const data = EM_SIM_ALL;

  const activas    = data.filter(r => _emSimActivada(r)).length;
  const sinActivar = data.length - activas;

  // ── KPIs ─────────────────────────────────────────────────────
  const kpiEl = document.getElementById('em-sim-kpis');
  if (kpiEl) {
    const kpis = [
      { label:'Total SIMCards', val: data.length, color:'#DFFF61', icon:'📡' },
      { label:'Activadas',      val: activas,      color:'#B0F2AE', icon:'✅' },
      { label:'Sin Activar',    val: sinActivar,   color:'#FF5C5C', icon:'❌' },
      { label:'% Activación',   val: data.length ? Math.round(activas/data.length*100)+'%' : '0%', color:'#99D1FC', icon:'📊' },
    ];
    kpiEl.innerHTML = kpis.map(k => `
      <div style="background:linear-gradient(145deg,rgba(10,26,18,.95),rgba(8,20,14,.9));
        border:1px solid ${k.color}22;border-top:2px solid ${k.color};border-radius:14px;
        padding:18px 20px;min-width:130px;flex:1;">
        <div style="font-size:20px;margin-bottom:6px">${k.icon}</div>
        <div style="font-family:'Outfit',sans-serif;font-size:11px;color:#7A7674;letter-spacing:.4px;text-transform:uppercase;margin-bottom:6px">${k.label}</div>
        <div style="font-family:'JetBrains Mono',monospace;font-size:26px;font-weight:700;color:${k.color};line-height:1;text-shadow:0 0 20px ${k.color}55">${typeof k.val === 'number' ? k.val.toLocaleString('es-CO') : k.val}</div>
      </div>`).join('');
  }

  // ── Gráfica 1: Dona activadas vs no activadas ───────────────
  _emDonut('em-chart-sim-estado',
    ['Activadas', 'Sin Activar'],
    [activas, sinActivar],
    [EM_COLORS.activada, EM_COLORS.sinActivar],
    'Estado SIMCards'
  );

  // ── Gráfica 2: Top sitios con más SIMCards ──────────────────
  const sitioMap = {};
  data.forEach(r => {
    const s = (r['Nombre de la ubicación'] || 'Sin ubicación').trim();
    sitioMap[s] = (sitioMap[s] || 0) + (parseInt(r['Cantidad'])||1);
  });
  const topSitios = Object.entries(sitioMap).sort((a,b)=>b[1]-a[1]).slice(0,15);
  _emBarH('em-chart-sim-sitios',
    topSitios.map(s => s[0].length > 28 ? s[0].substring(0,28)+'…' : s[0]),
    topSitios.map(s => s[1]),
    null,
    'SIMCards'
  );

  // ── Gráfica 3: Sitios con más SIMCards ACTIVADAS ────────────
  const sitioActivMap = {};
  data.filter(r => _emSimActivada(r)).forEach(r => {
    const s = (r['Nombre de la ubicación'] || 'Sin ubicación').trim();
    sitioActivMap[s] = (sitioActivMap[s] || 0) + (parseInt(r['Cantidad'])||1);
  });
  const topActiv = Object.entries(sitioActivMap).sort((a,b)=>b[1]-a[1]).slice(0,12);
  _emBarH('em-chart-sim-activadas',
    topActiv.map(s => s[0].length > 28 ? s[0].substring(0,28)+'…' : s[0]),
    topActiv.map(s => s[1]),
    topActiv.map(() => '#B0F2AEBB'),
    'Activadas'
  );

  // ── Gráfica 4: Distribución por fecha de activación (mensual) ─
  const mesMap = {};
  data.filter(r => _emSimActivada(r)).forEach(r => {
    const fa = _emSimFechaActivacion(r);
    if (fa === '—') return;
    const parts = fa.split('/');
    if (parts.length < 3) return;
    const key = `${parts[2]}-${parts[1]}`;  // YYYY-MM
    mesMap[key] = (mesMap[key]||0) + 1;
  });
  const mesSorted = Object.entries(mesMap).sort((a,b)=>a[0].localeCompare(b[0]));
  if (mesSorted.length) {
    const canvasMes = document.getElementById('em-chart-sim-mensual');
    if (canvasMes) {
      _emDestroyChart('em-chart-sim-mensual');
      EM_CHARTS['em-chart-sim-mensual'] = new Chart(canvasMes, {
        type: 'bar',
        data: {
          labels: mesSorted.map(([k]) => { const [y,m] = k.split('-'); return `${m}/${y}`; }),
          datasets: [{
            label: 'Activaciones',
            data:  mesSorted.map(([,v]) => v),
            backgroundColor: '#B0F2AEAA',
            borderColor: '#B0F2AE',
            borderWidth: 1, borderRadius: 6,
          }]
        },
        options: {
          responsive: true, maintainAspectRatio: false,
          plugins: {
            legend: { display: false },
            tooltip: { backgroundColor: 'rgba(24,23,21,.95)', titleColor: '#B0F2AE', bodyColor: '#FAFAFA', borderColor: 'rgba(176,242,174,.2)', borderWidth: 1, padding: 10 }
          },
          scales: {
            x: { ticks: { color: '#cbd5e1', font: { size: 10 }, maxRotation: 30 }, grid: { display: false } },
            y: { ticks: { color: '#7A7674' }, grid: { color: 'rgba(255,255,255,.05)' } }
          }
        }
      });
    }
  }

  // ── Tabla SIMCards ───────────────────────────────────────────
  _emApplySimSearch('');
}

function _emApplySimSearch(query) {
  EM_SIM_SEARCH = (query || '').toLowerCase().trim();
  EM_SIM_FILTERED = EM_SIM_ALL.filter(r => {
    if (!EM_SIM_SEARCH) return true;
    const serial = (r['Número de serie']||r['Numero de serie']||r['Serial']||'').toLowerCase();
    const ubic   = (r['Nombre de la ubicación']||'').toLowerCase();
    const attr   = (r['Atributos']||'').toLowerCase();
    const cod    = (r['Código de ubicación']||'').toLowerCase();
    return serial.includes(EM_SIM_SEARCH) || ubic.includes(EM_SIM_SEARCH) || attr.includes(EM_SIM_SEARCH) || cod.includes(EM_SIM_SEARCH);
  });
  EM_SIM_PAGE = 1;
  _emRenderSimTabla();
}

function _emRenderSimTabla() {
  const wrap  = document.getElementById('em-sim-tabla-wrap');
  const count = document.getElementById('em-sim-tabla-count');
  const pagEl = document.getElementById('em-sim-tabla-pag');
  if (!wrap) return;

  const data  = EM_SIM_FILTERED;
  const pages = Math.max(1, Math.ceil(data.length / EM_PAGE_SIZE));
  const slice = data.slice((EM_SIM_PAGE-1)*EM_PAGE_SIZE, EM_SIM_PAGE*EM_PAGE_SIZE);

  if (count) count.textContent = `${data.length.toLocaleString('es-CO')} registros`;

  if (!slice.length) {
    wrap.innerHTML = '<div style="text-align:center;padding:40px;color:#7A7674;font-family:\'Outfit\',sans-serif;">Sin resultados</div>';
    if (pagEl) pagEl.innerHTML = '';
    return;
  }

  wrap.innerHTML = `<table style="width:100%;border-collapse:collapse;font-size:12px;">
    <thead>
      <tr style="background:rgba(0,0,0,.3);">
        ${['Serial','Ubicación','Cód. Ubic.','Estado','Fecha Activación','Atributos'].map(h =>
          `<th style="padding:10px 14px;text-align:left;font-family:'Syne',sans-serif;font-size:10px;
            font-weight:700;color:#99D1FC;letter-spacing:1px;text-transform:uppercase;
            border-bottom:1px solid rgba(153,209,252,.15);white-space:nowrap;">${h}</th>`
        ).join('')}
      </tr>
    </thead>
    <tbody>
      ${slice.map((r, i) => {
        const activa = _emSimActivada(r);
        const fa     = _emSimFechaActivacion(r);
        const serial = r['Número de serie']||r['Numero de serie']||r['Serial']||'—';
        const ubic   = r['Nombre de la ubicación']||'—';
        const cod    = r['Código de ubicación']||'—';
        const attr   = r['Atributos']||'—';
        const bg = i%2===0 ? 'rgba(153,209,252,.012)' : 'transparent';
        const estadoPill = activa
          ? `<span style="background:rgba(176,242,174,.12);color:#B0F2AE;font-size:10px;font-weight:700;padding:3px 10px;border-radius:12px;border:1px solid rgba(176,242,174,.3);display:inline-block;font-family:'Outfit',sans-serif;">✅ ACTIVADA</span>`
          : `<span style="background:rgba(255,92,92,.12);color:#FF5C5C;font-size:10px;font-weight:700;padding:3px 10px;border-radius:12px;border:1px solid rgba(255,92,92,.3);display:inline-block;font-family:'Outfit',sans-serif;">❌ SIN ACTIVAR</span>`;
        return `<tr style="background:${bg}" onmouseover="this.style.background='rgba(153,209,252,.04)'" onmouseout="this.style.background='${bg}'">
          <td style="padding:9px 14px;font-family:'JetBrains Mono',monospace;font-size:11px;color:#a5f3fc;">${serial}</td>
          <td style="padding:9px 14px;font-size:11px;color:#cbd5e1;max-width:220px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="${ubic}">${ubic}</td>
          <td style="padding:9px 14px;font-family:'JetBrains Mono',monospace;font-size:11px;color:#67e8f9;">${cod}</td>
          <td style="padding:9px 14px;">${estadoPill}</td>
          <td style="padding:9px 14px;font-family:'JetBrains Mono',monospace;font-size:11px;color:${activa?'#B0F2AE':'#64748b'};">${fa}</td>
          <td style="padding:9px 14px;font-size:11px;color:#94a3b8;max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="${attr}">${attr}</td>
        </tr>`;
      }).join('')}
    </tbody>
  </table>`;

  _emRenderPag(pagEl, EM_SIM_PAGE, pages, p => { EM_SIM_PAGE = p; _emRenderSimTabla(); }, '#99D1FC');
}

// ══════════════════════════════════════════════════════════════════
//  PAGINACIÓN
// ══════════════════════════════════════════════════════════════════
function _emRenderPag(el, cur, total, onPage, accent) {
  if (!el) return;
  el.innerHTML = '';
  if (total <= 1) return;
  const btnCSS = `padding:5px 11px;border-radius:7px;cursor:pointer;font-size:12px;
    font-family:'Outfit',sans-serif;border:1px solid rgba(255,255,255,.1);transition:all .15s;`;
  function mk(label, page, active) {
    const b = document.createElement('button');
    b.style.cssText = btnCSS + (active
      ? `background:${accent};color:#0a1a12;font-weight:700;border-color:${accent};`
      : 'background:rgba(255,255,255,.06);color:#94a3b8;');
    b.textContent = label;
    b.addEventListener('click', () => onPage(page));
    return b;
  }
  if (cur > 1) el.appendChild(mk('‹', cur-1, false));
  const s = Math.max(1, cur-2), e = Math.min(total, cur+2);
  for (let p = s; p <= e; p++) el.appendChild(mk(String(p), p, p === cur));
  if (cur < total) el.appendChild(mk('›', cur+1, false));
}

// ══════════════════════════════════════════════════════════════════
//  RENDER HTML PRINCIPAL
// ══════════════════════════════════════════════════════════════════
function _emBuildHTML() {
  const panel = document.getElementById('panel-estado-materiales');
  if (!panel) return;

  const _sectionTitle = (text, color, icon) => `
    <div style="display:flex;align-items:center;gap:12px;margin:36px 0 20px;padding-bottom:10px;
      border-bottom:1px solid ${color}22;">
      <span style="font-size:22px;">${icon}</span>
      <div>
        <div style="font-family:'Syne',sans-serif;font-size:17px;font-weight:800;color:${color};letter-spacing:-.3px;">${text}</div>
      </div>
    </div>`;

  const _chartCard = (title, sub, canvasId, h = '260px', extra = '') => `
    <div style="background:linear-gradient(145deg,rgba(10,26,18,.95),rgba(8,20,14,.9));
      border:1px solid rgba(103,232,249,.1);border-radius:18px;padding:20px 22px;${extra}">
      <div style="font-family:'Syne',sans-serif;font-size:13px;font-weight:700;color:#f1f5f9;margin-bottom:4px">${title}</div>
      ${sub ? `<div style="font-size:11px;color:#7A7674;margin-bottom:14px">${sub}</div>` : '<div style="margin-bottom:14px"></div>'}
      <div style="position:relative;height:${h}"><canvas id="${canvasId}"></canvas></div>
    </div>`;

  panel.innerHTML = `
  <div style="padding:0 4px 40px;">

    <!-- ══ SECCIÓN DATÁFONOS ══ -->
    ${_sectionTitle('ESTADO DE DATÁFONOS EN BODEGA', '#B0F2AE', '📱')}

    <!-- KPIs datáfonos -->
    <div id="em-df-kpis" style="display:flex;flex-wrap:wrap;gap:12px;margin-bottom:28px;">
      <div class="loading"><div class="spinner"></div><span>Cargando...</span></div>
    </div>

    <!-- Gráficas row 1 -->
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:18px;margin-bottom:20px;">
      ${_chartCard('Posición en Depósito', 'Solo estados: EN DAÑO · DES. INCIDENTE · DES. CIERRE', 'em-chart-deposito', '240px')}
      ${_chartCard('Estados por Referencia', 'Cantidad de datáfonos por referencia y estado de depósito', 'em-chart-ref-estado', '240px')}
    </div>

    <!-- Gráfica tipos de daño (ancho completo) -->
    <div style="margin-bottom:24px;">
      ${_chartCard('Tipos de Daño', 'Extraído del campo Comentarios (solo datáfonos EN DAÑO)', 'em-chart-tipos-danio', '280px', 'margin-bottom:0')}
    </div>

    <!-- Tabla por bodega -->
    <div style="background:linear-gradient(145deg,rgba(10,26,18,.95),rgba(8,20,14,.9));
      border:1px solid rgba(176,242,174,.12);border-radius:18px;overflow:hidden;margin-bottom:28px;
      box-shadow:0 4px 20px rgba(0,0,0,.4);">
      <div style="padding:16px 20px;border-bottom:1px solid rgba(176,242,174,.08);display:flex;align-items:center;justify-content:space-between;">
        <div>
          <div style="font-family:'Syne',sans-serif;font-size:13px;font-weight:700;color:#f1f5f9;">Estado por Bodega</div>
          <div style="font-size:11px;color:#7A7674;margin-top:2px;">Disponible · Asociado · Dañado · Des. Cierre · Des. Incidente</div>
        </div>
      </div>
      <div id="em-tabla-bodegas-wrap" style="overflow-x:auto;max-height:480px;">
        <div class="loading"><div class="spinner"></div></div>
      </div>
    </div>

    <!-- Tabla completa datáfonos -->
    <div style="background:linear-gradient(145deg,rgba(10,26,18,.95),rgba(8,20,14,.9));
      border:1px solid rgba(176,242,174,.12);border-radius:18px;overflow:hidden;margin-bottom:40px;
      box-shadow:0 4px 20px rgba(0,0,0,.4);">
      <div style="padding:16px 20px;border-bottom:1px solid rgba(176,242,174,.08);display:flex;align-items:center;justify-content:space-between;">
        <div>
          <div style="font-family:'Syne',sans-serif;font-size:13px;font-weight:700;color:#f1f5f9;">Tabla Completa — Datáfonos en Bodegas</div>
          <div style="font-size:11px;color:#7A7674;margin-top:2px;" id="em-df-tabla-count">0 registros</div>
        </div>
        <button onclick="emExportDfExcel()"
          style="background:rgba(176,242,174,.08);border:1px solid rgba(176,242,174,.2);color:#B0F2AE;
          padding:7px 14px;border-radius:8px;cursor:pointer;font-size:12px;font-family:'Outfit',sans-serif;transition:all .2s;">⬇ Excel</button>
      </div>
      <div style="padding:10px 20px;border-bottom:1px solid rgba(176,242,174,.07);background:rgba(0,0,0,.15);">
        <input type="text" oninput="_emApplyDfSearch(this.value)"
          placeholder="🔍 Buscar por serial, referencia, bodega, estado, comentarios..."
          style="width:100%;box-sizing:border-box;background:rgba(255,255,255,.05);border:1px solid rgba(176,242,174,.2);
          border-radius:8px;color:#FAFAFA;padding:8px 14px;font-size:13px;font-family:'Outfit',sans-serif;outline:none;">
      </div>
      <div id="em-df-tabla-wrap" style="overflow-x:auto;max-height:520px;">
        <div class="loading"><div class="spinner"></div></div>
      </div>
      <div style="display:flex;align-items:center;justify-content:flex-end;padding:10px 20px;border-top:1px solid rgba(176,242,174,.07);">
        <div id="em-df-tabla-pag" style="display:flex;gap:6px;flex-wrap:wrap;"></div>
      </div>
    </div>

    <!-- ══ SECCIÓN SIMCARDS ══ -->
    ${_sectionTitle('ESTADO DE SIMCARDS', '#99D1FC', '📡')}

    <!-- KPIs simcards -->
    <div id="em-sim-kpis" style="display:flex;flex-wrap:wrap;gap:12px;margin-bottom:28px;">
      <div class="loading"><div class="spinner"></div><span>Cargando...</span></div>
    </div>

    <!-- Gráficas SIM row 1 -->
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:18px;margin-bottom:20px;">
      ${_chartCard('Activadas vs Sin Activar', 'FA:DD/MM/AAAA en Atributos = activada', 'em-chart-sim-estado', '240px')}
      ${_chartCard('Sitios con más SIMCards Activadas', 'Top 12 por ubicación', 'em-chart-sim-activadas', '240px')}
    </div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:18px;margin-bottom:24px;">
      ${_chartCard('Top Sitios por SIMCards', 'Total por ubicación (activadas + sin activar)', 'em-chart-sim-sitios', '280px')}
      ${_chartCard('Activaciones por Mes', 'Distribución temporal de fechas FA', 'em-chart-sim-mensual', '280px')}
    </div>

    <!-- Tabla SIMCards -->
    <div style="background:linear-gradient(145deg,rgba(10,26,18,.95),rgba(8,20,14,.9));
      border:1px solid rgba(153,209,252,.12);border-radius:18px;overflow:hidden;margin-bottom:40px;
      box-shadow:0 4px 20px rgba(0,0,0,.4);">
      <div style="padding:16px 20px;border-bottom:1px solid rgba(153,209,252,.08);display:flex;align-items:center;justify-content:space-between;">
        <div>
          <div style="font-family:'Syne',sans-serif;font-size:13px;font-weight:700;color:#f1f5f9;">Tabla Completa — SIMCards</div>
          <div style="font-size:11px;color:#7A7674;margin-top:2px;" id="em-sim-tabla-count">0 registros</div>
        </div>
        <button onclick="emExportSimExcel()"
          style="background:rgba(153,209,252,.08);border:1px solid rgba(153,209,252,.2);color:#99D1FC;
          padding:7px 14px;border-radius:8px;cursor:pointer;font-size:12px;font-family:'Outfit',sans-serif;transition:all .2s;">⬇ Excel</button>
      </div>
      <div style="padding:10px 20px;border-bottom:1px solid rgba(153,209,252,.07);background:rgba(0,0,0,.15);">
        <input type="text" oninput="_emApplySimSearch(this.value)"
          placeholder="🔍 Buscar por serial, nombre de sitio, atributos, código..."
          style="width:100%;box-sizing:border-box;background:rgba(255,255,255,.05);border:1px solid rgba(153,209,252,.2);
          border-radius:8px;color:#FAFAFA;padding:8px 14px;font-size:13px;font-family:'Outfit',sans-serif;outline:none;">
      </div>
      <div id="em-sim-tabla-wrap" style="overflow-x:auto;max-height:520px;">
        <div class="loading"><div class="spinner"></div></div>
      </div>
      <div style="display:flex;align-items:center;justify-content:flex-end;padding:10px 20px;border-top:1px solid rgba(153,209,252,.07);">
        <div id="em-sim-tabla-pag" style="display:flex;gap:6px;flex-wrap:wrap;"></div>
      </div>
    </div>

  </div>`;
}

// ══════════════════════════════════════════════════════════════════
//  EXPORTS EXCEL
// ══════════════════════════════════════════════════════════════════
window.emExportDfExcel = function() {
  if (!EM_DF_FILTERED.length) { alert('Sin datos.'); return; }
  if (typeof XLSX === 'undefined') { alert('XLSX no disponible.'); return; }
  const rows = EM_DF_FILTERED.map(r => ({
    'Serial':              r['Número de serie']||r['Numero de serie']||r['Serial']||'',
    'Referencia':          r['Nombre']||'',
    'Bodega':              r['Nombre de la ubicación']||'',
    'Código Ubicación':    r['Código de ubicación']||'',
    'Posición Depósito':   r['Posición en depósito']||'',
    'Estado':              _emEstadoDatafono(r),
    'Tipo Daño':           (r['Posición en depósito']||'').trim().toUpperCase() === 'EN DAÑO' ? _emTipoDanio(r) : '',
    'Comentarios':         r['Comentarios']||'',
    'Atributos':           r['Atributos']||'',
  }));
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(rows);
  ws['!cols'] = Object.keys(rows[0]).map(k => ({ wch: Math.max(k.length+2, 14) }));
  XLSX.utils.book_append_sheet(wb, ws, 'Datáfonos Bodegas');
  XLSX.writeFile(wb, `Datafonos_Bodegas_${new Date().toISOString().slice(0,10)}.xlsx`);
};

window.emExportSimExcel = function() {
  if (!EM_SIM_FILTERED.length) { alert('Sin datos.'); return; }
  if (typeof XLSX === 'undefined') { alert('XLSX no disponible.'); return; }
  const rows = EM_SIM_FILTERED.map(r => ({
    'Serial':              r['Número de serie']||r['Numero de serie']||r['Serial']||'',
    'Ubicación':           r['Nombre de la ubicación']||'',
    'Código Ubicación':    r['Código de ubicación']||'',
    'Estado':              _emSimActivada(r) ? 'ACTIVADA' : 'SIN ACTIVAR',
    'Fecha Activación':    _emSimFechaActivacion(r),
    'Atributos':           r['Atributos']||'',
  }));
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(rows);
  ws['!cols'] = Object.keys(rows[0]).map(k => ({ wch: Math.max(k.length+2, 14) }));
  XLSX.utils.book_append_sheet(wb, ws, 'SIMCards');
  XLSX.writeFile(wb, `SIMCards_${new Date().toISOString().slice(0,10)}.xlsx`);
};

// ══════════════════════════════════════════════════════════════════
//  ENTRY POINT
// ══════════════════════════════════════════════════════════════════
async function renderEstadoMateriales() {
  const panel = document.getElementById('panel-estado-materiales');
  if (!panel) return;

  let raw = _emGetRaw();
  if (!raw || !raw.length) {
    panel.innerHTML = '<div class="loading"><div class="spinner"></div><span>Cargando datos de inventario...</span></div>';
    if (typeof window.loadInventarioData === 'function') {
      await window.loadInventarioData();
    }
    raw = _emGetRaw();
  }

  if (!raw || !raw.length) {
    panel.innerHTML = '<div class="loading"><div class="spinner"></div><span>No hay datos disponibles. Asegúrate de que stock_wompi_filtrado.json.gz esté cargado.</span></div>';
    return;
  }

  // Destruir charts anteriores
  Object.keys(EM_CHARTS).forEach(id => { try { EM_CHARTS[id].destroy(); } catch(e){} });
  EM_CHARTS = {};

  // Construir HTML
  _emBuildHTML();

  // Procesar datos
  _emProcesar();

  // Renderizar secciones
  requestAnimationFrame(() => {
    _emRenderDatafonos();
    _emRenderSimcards();
  });
}

// Exponer funciones que se llaman desde HTML inline
window.renderEstadoMateriales = renderEstadoMateriales;
window._emApplyDfSearch  = _emApplyDfSearch;
window._emApplySimSearch = _emApplySimSearch;