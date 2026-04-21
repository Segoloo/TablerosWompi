/**
 * ╔══════════════════════════════════════════════════════════════════╗
 * ║  inventario_detalles.js — Tab "Detalles" Inventario Wompi       ║
 * ║  Filtros avanzados · Alertas · 2 Tablas · Gráficas              ║
 * ╚══════════════════════════════════════════════════════════════════╝
 */

'use strict';

// ══════════════════════════════════════════════════════════════════
//  LÓGICA UBICACIÓNv3  (replica fórmula DAX)
// ══════════════════════════════════════════════════════════════════
const GW_EW_SET = new Set([
  "GW060","GW065","GW068","GW069","GW070","GW071",
  "GW072","GW073","GW074","GW075","GW076","GW077",
  "GW078","GW079","GW080","GW081","GW082",
  "GW083","GW087","GW002","GW090","GW094",
  "GW084","GW085","GW086","GW093","GW095",
  "GW092","GW106"
]);

function invUbicacionV3(row) {
  const tipo = (row['Tipo de ubicación'] || '').trim();
  const cod  = (row['Código de ubicación'] || '').trim().toUpperCase();
  const pos  = (row['Posición en depósito'] || row['posicion_en_deposito'] || '').trim().toUpperCase();

  if (tipo === 'Site' || tipo === 'Network Element') return 'En corresponsal';
  if (tipo === 'Staff')    return 'Gestor LineaCom';
  if (tipo === 'Supplier') return 'En operador Logistico';
  if (pos === 'ENVIADO TERMINAL-TERMINAL' || pos === 'ENVIADO OPERADOR LOGISTICO') return 'En distribución';
  if (cod === 'VW-0278')   return 'En Ingenico';
  if (cod.startsWith('GW')) {
    if (GW_EW_SET.has(cod)) return 'Empleados Wompi';
    return 'Gestor Wompi';
  }
  return 'En bodega';
}

// ══════════════════════════════════════════════════════════════════
//  LÓGICA ALERTA (>8 días en Gestor LineaCom o Gestor Wompi)
// ══════════════════════════════════════════════════════════════════
function invGetAlerta(row, ubicacion) {
  if (ubicacion !== 'Gestor LineaCom' && ubicacion !== 'Gestor Wompi') return false;
  const rawFecha = row['Fecha de la última edición'] || row['Fecha confirmación'] || row['fecha_confirmacion'] || '';
  if (!rawFecha) return false;
  const fecha = new Date(rawFecha);
  if (isNaN(fecha.getTime())) return false;
  const dias = Math.floor((Date.now() - fecha.getTime()) / 86400000);
  return dias > 8;
}

function invFechaConfirmacion(row) {
  return row['Fecha de la última edición'] || row['Fecha confirmación'] || row['fecha_confirmacion'] || '';
}

function invDiasTranscurridos(row) {
  const raw = invFechaConfirmacion(row);
  if (!raw) return null;
  const f = new Date(raw);
  if (isNaN(f.getTime())) return null;
  return Math.floor((Date.now() - f.getTime()) / 86400000);
}

function invFmtFecha(raw) {
  if (!raw) return '—';
  const f = new Date(raw);
  if (isNaN(f.getTime())) return raw;
  return f.toLocaleDateString('es-CO', { day:'2-digit', month:'2-digit', year:'numeric' });
}

// ══════════════════════════════════════════════════════════════════
//  ESTADO DEL TAB DETALLES
// ══════════════════════════════════════════════════════════════════
let DET_FILTERED  = [];
let DET_PAGE_T1   = 1;
let DET_PAGE_T2   = 1;
const DET_PAGE_SIZE = 50;
let DET_CHARTS    = {};
let DET_SORT_COL  = null;
let DET_SORT_DIR  = 1;

// ══════════════════════════════════════════════════════════════════
//  POBLAR FILTROS DEL TAB DETALLES
// ══════════════════════════════════════════════════════════════════
function detPopulateFilters() {
  const raw = window.INV_RAW || [];
  if (!raw.length) return;

  // Tipo de negocio — reutiliza invNegocio global
  // Categoría — reutiliza invCategoria global

  // Ubicación V3 única
  const ubicaciones = [...new Set(raw.map(r => invUbicacionV3(r)))].sort();
  _detSetSelect('det-f-ubicacion', ['Todas', ...ubicaciones]);

  // Nombre de la ubicación
  const nombres = [...new Set(raw.map(r => (r['Nombre de la ubicación']||'').trim()).filter(Boolean))].sort();
  _detSetSelect('det-f-nombre-ubic', ['Todas', ...nombres]);

  // Código de comercio
  const codCom = [...new Set(raw.map(r => (r['Código de comercio'] || r['codigo_de_comercio'] || r['Código de ubicación']||'').trim()).filter(Boolean))].sort();
  _detSetSelect('det-f-cod-comercio', ['Todos', ...codCom]);

  // Número de serie
  const seriales = [...new Set(raw.map(r => (r['Número de serie']||r['Numero de serie']||r['Serial']||'').trim()).filter(Boolean))].sort();
  _detSetSelect('det-f-serial', ['Todos', ...seriales.slice(0,500)]);

  // Referencia (Nombre)
  const refs = [...new Set(raw.map(r => (r['Nombre']||'').trim()).filter(Boolean))].sort();
  _detSetSelect('det-f-referencia', ['Todas', ...refs]);
}

function _detSetSelect(id, opts) {
  const el = document.getElementById(id);
  if (!el) return;
  const cur = el.value;
  el.innerHTML = opts.map(o => `<option value="${o}">${o}</option>`).join('');
  if (opts.includes(cur)) el.value = cur;
}

// ══════════════════════════════════════════════════════════════════
//  APLICAR FILTROS DETALLES
// ══════════════════════════════════════════════════════════════════
function detApplyFilters() {
  const raw = window.INV_RAW || [];
  if (!raw.length) { DET_FILTERED = []; detRenderAll(); return; }

  const negocio    = document.getElementById('det-f-negocio')?.value    || '';
  const categoria  = document.getElementById('det-f-categoria')?.value  || '';
  const ubicacion  = document.getElementById('det-f-ubicacion')?.value  || '';
  const nombreUbic = document.getElementById('det-f-nombre-ubic')?.value || '';
  const codCom     = document.getElementById('det-f-cod-comercio')?.value || '';
  const alerta     = document.getElementById('det-f-alerta')?.value     || '';
  const serial     = document.getElementById('det-f-serial')?.value     || '';
  const referencia = document.getElementById('det-f-referencia')?.value || '';

  DET_FILTERED = raw.filter(r => {
    if (negocio   && negocio   !== 'Todos' && invNegocio(r['Subtipo'])                       !== negocio)   return false;
    if (categoria && categoria !== 'Todas' && invCategoria(r['Nombre'])                       !== categoria) return false;
    const ubV3 = invUbicacionV3(r);
    if (ubicacion && ubicacion !== 'Todas' && ubV3 !== ubicacion) return false;
    if (nombreUbic && nombreUbic !== 'Todas' && (r['Nombre de la ubicación']||'').trim()     !== nombreUbic) return false;
    const codVal = (r['Código de comercio']||r['codigo_de_comercio']||r['Código de ubicación']||'').trim();
    if (codCom    && codCom    !== 'Todos' && codVal !== codCom) return false;
    const tieneAlerta = invGetAlerta(r, ubV3);
    if (alerta === 'con_alerta'    && !tieneAlerta) return false;
    if (alerta === 'sin_alerta'    && tieneAlerta)  return false;
    const serialVal = (r['Número de serie']||r['Numero de serie']||r['Serial']||'').trim();
    if (serial    && serial    !== 'Todos' && serialVal !== serial) return false;
    if (referencia && referencia !== 'Todas' && (r['Nombre']||'').trim() !== referencia) return false;
    return true;
  });

  DET_PAGE_T1 = 1;
  DET_PAGE_T2 = 1;
  detRenderAll();
}

window.detApplyFilters = detApplyFilters;
window.detResetFilters = function() {
  ['det-f-negocio','det-f-categoria','det-f-ubicacion','det-f-nombre-ubic',
   'det-f-cod-comercio','det-f-alerta','det-f-serial','det-f-referencia'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.selectedIndex = 0;
  });
  detApplyFilters();
};

// ══════════════════════════════════════════════════════════════════
//  RENDER ALL
// ══════════════════════════════════════════════════════════════════
function detRenderAll() {
  _detRenderKPIStrip();
  _detRenderCharts();
  _detRenderTabla1();
  _detRenderTabla2();
}

// ── Strip de KPIs de alerta ────────────────────────────────────────
function _detRenderKPIStrip() {
  const strip = document.getElementById('det-kpi-strip');
  if (!strip) return;
  const rows = DET_FILTERED;
  const total = rows.length;
  let conAlerta = 0, sinAlerta = 0, gestorLC = 0, gestorWP = 0;
  rows.forEach(r => {
    const ub = invUbicacionV3(r);
    const al = invGetAlerta(r, ub);
    if (al) conAlerta++; else sinAlerta++;
    if (ub === 'Gestor LineaCom') gestorLC++;
    if (ub === 'Gestor Wompi')    gestorWP++;
  });
  const totalQty = rows.reduce((s,r) => s + (parseInt(r['Cantidad'])||0), 0);

  strip.innerHTML = [
    { icon:'📦', label:'Total Registros',     val: total.toLocaleString('es-CO'),     color:'#DFFF61' },
    { icon:'⚡', label:'Uds. en Inventario',  val: totalQty.toLocaleString('es-CO'),  color:'#B0F2AE' },
    { icon:'🚨', label:'Con Alerta (>8d)',     val: conAlerta.toLocaleString('es-CO'), color:'#F87171' },
    { icon:'✅', label:'Sin Alerta',           val: sinAlerta.toLocaleString('es-CO'), color:'#B0F2AE' },
    { icon:'🔧', label:'Gestor LineaCom',      val: gestorLC.toLocaleString('es-CO'),  color:'#FFC04D' },
    { icon:'👤', label:'Gestor Wompi',         val: gestorWP.toLocaleString('es-CO'),  color:'#C084FC' },
  ].map(k => `
    <div style="background:linear-gradient(145deg,rgba(10,26,18,.95),rgba(8,20,14,.9));
      border:1px solid rgba(255,255,255,.07);border-top:2px solid ${k.color};
      border-radius:14px;padding:16px 18px;flex:1;min-width:140px;
      box-shadow:0 4px 20px rgba(0,0,0,.35);">
      <div style="font-size:18px;margin-bottom:8px">${k.icon}</div>
      <div style="font-size:9px;font-weight:700;color:#7A7674;text-transform:uppercase;letter-spacing:1px;
        font-family:'Syne',sans-serif;margin-bottom:6px">${k.label}</div>
      <div style="font-family:'JetBrains Mono',monospace;font-size:26px;font-weight:700;
        color:${k.color};line-height:1;text-shadow:0 0 20px ${k.color}55">${k.val}</div>
    </div>`).join('');
}

// ══════════════════════════════════════════════════════════════════
//  GRÁFICAS DEL TAB DETALLES
// ══════════════════════════════════════════════════════════════════
function _detDestroyCharts() {
  Object.values(DET_CHARTS).forEach(c => { try { c.destroy(); } catch(_) {} });
  DET_CHARTS = {};
}

function _detRenderCharts() {
  const grid = document.getElementById('det-charts-grid');
  if (!grid) return;
  _detDestroyCharts();

  const rows = DET_FILTERED;
  if (!rows.length) {
    grid.innerHTML = '<div style="color:#7A7674;padding:32px;text-align:center;font-family:\'Outfit\',sans-serif;">Sin datos para mostrar gráficas</div>';
    return;
  }

  // Agregar datos
  const ubicMap = {}, negMap = { CB: 0, VP: 0 }, catMap = {}, alertMap = { 'Con alerta': 0, 'Sin alerta': 0 };
  rows.forEach(r => {
    const ub  = invUbicacionV3(r);
    const neg = invNegocio(r['Subtipo']);
    const cat = invCategoria(r['Nombre']);
    const qty = parseInt(r['Cantidad']) || 0;
    const al  = invGetAlerta(r, ub);
    ubicMap[ub]  = (ubicMap[ub] || 0) + qty;
    negMap[neg]  = (negMap[neg] || 0) + qty;
    catMap[cat]  = (catMap[cat] || 0) + qty;
    if (al) alertMap['Con alerta']++; else alertMap['Sin alerta']++;
  });

  const TOOLTIP = {
    backgroundColor:'rgba(24,23,21,.95)',titleColor:'#DFFF61',bodyColor:'#FAFAFA',
    borderColor:'rgba(223,255,97,.2)',borderWidth:1,padding:12,
    titleFont:{family:'Syne',size:13,weight:'700'},
    bodyFont:{family:'Outfit',size:12},
  };
  const LEGEND = { labels:{ color:'#FAFAFA',font:{family:'Outfit',size:11},padding:14,boxWidth:11 } };
  const XGRID  = { color:'rgba(255,255,255,.05)' };
  const YTICK  = { color:'#7A7674',font:{family:'Outfit',size:11},border:{display:false} };
  const XTICK  = { color:'#7A7674',font:{family:'JetBrains Mono',size:10},border:{display:false},
    callback: v => v.toLocaleString('es-CO') };

  function card(canvasId, title, sub, span) {
    const d = document.createElement('div');
    d.style.cssText = `background:linear-gradient(145deg,rgba(10,26,18,.95),rgba(8,20,14,.9));
      border:1px solid rgba(223,255,97,.1);border-radius:18px;padding:22px 24px 18px;
      position:relative;overflow:hidden;box-shadow:0 4px 20px rgba(0,0,0,.4);
      ${span ? 'grid-column:span '+span+';' : ''}`;
    d.innerHTML = `<div style="position:absolute;top:0;left:0;right:0;height:2px;
      background:linear-gradient(90deg,#DFFF61,#B0F2AE,#99D1FC);opacity:.6;"></div>
      <div style="margin-bottom:14px;">
        <div style="font-family:'Syne',sans-serif;font-size:13px;font-weight:700;color:#f1f5f9;">${title}</div>
        ${sub ? `<div style="font-size:11px;color:#7A7674;margin-top:3px;font-family:'Outfit',sans-serif;">${sub}</div>` : ''}
      </div>
      <canvas id="${canvasId}" style="max-height:240px;"></canvas>`;
    return d;
  }

  grid.innerHTML = '';

  // ── 1. Donut: distribución por UbicacionV3
  const ubEntries = Object.entries(ubicMap).sort((a,b) => b[1]-a[1]);
  const UB_COLORS = ['#B0F2AE','#99D1FC','#FFC04D','#C084FC','#F87171','#FB923C','#DFFF61','#67e8f9'];
  const c1 = card('det-c-ubicacion','Distribución por Ubicación V3','Unidades por tipo de ubicación calculado');
  grid.appendChild(c1);
  DET_CHARTS.ubicacion = new Chart(document.getElementById('det-c-ubicacion'), {
    type: 'doughnut',
    data: {
      labels: ubEntries.map(e => e[0]),
      datasets: [{ data: ubEntries.map(e => e[1]),
        backgroundColor: UB_COLORS.slice(0, ubEntries.length).map(c => c+'BB'),
        borderColor: 'rgba(0,0,0,0)', borderWidth: 0, hoverOffset: 8 }]
    },
    options: { cutout:'65%', plugins:{ legend: LEGEND, tooltip: TOOLTIP } }
  });

  // ── 2. Donut: Con alerta vs Sin alerta
  const c2 = card('det-c-alertas','Estado de Alertas','Registros con permanencia > 8 días en gestores');
  grid.appendChild(c2);
  DET_CHARTS.alertas = new Chart(document.getElementById('det-c-alertas'), {
    type: 'doughnut',
    data: {
      labels: ['Con alerta','Sin alerta'],
      datasets: [{ data: [alertMap['Con alerta'], alertMap['Sin alerta']],
        backgroundColor: ['#F8717188','#B0F2AE88'],
        borderColor: ['#F87171','#B0F2AE'], borderWidth: 2, hoverOffset: 8 }]
    },
    options: { cutout:'60%', plugins:{ legend: LEGEND, tooltip: TOOLTIP } }
  });

  // ── 3. Bar: CB vs VP
  const c3 = card('det-c-negocio','Negocio CB vs VP','Unidades por tipo de negocio');
  grid.appendChild(c3);
  DET_CHARTS.negocio = new Chart(document.getElementById('det-c-negocio'), {
    type: 'bar',
    data: {
      labels: ['CB (Wompi)', 'VP'],
      datasets: [{ data: [negMap.CB, negMap.VP],
        backgroundColor: ['#99D1FC99','#C084FC99'],
        borderColor: ['#99D1FC','#C084FC'], borderWidth: 2, borderRadius: 10, borderSkipped: false }]
    },
    options: {
      plugins:{ legend:{display:false}, tooltip: TOOLTIP },
      scales:{ x:{grid:{display:false},ticks:{color:'#FAFAFA',font:{family:'Outfit',size:13,weight:'600'}},border:{display:false}}, y:{grid:XGRID,ticks:XTICK} }
    }
  });

  // ── 4. Bar horizontal: por categoría (span 2)
  const catEntries = Object.entries(catMap).sort((a,b) => b[1]-a[1]);
  const CAT_COLORS = ['#B0F2AE','#99D1FC','#DFFF61','#C084FC','#FFC04D','#F87171','#FB923C','#67e8f9'];
  const c4 = card('det-c-categoria','Unidades por Categoría','Distribución del inventario filtrado', 2);
  grid.appendChild(c4);
  DET_CHARTS.categoria = new Chart(document.getElementById('det-c-categoria'), {
    type: 'bar',
    data: {
      labels: catEntries.map(e => e[0]),
      datasets: [{ label:'Unidades', data: catEntries.map(e => e[1]),
        backgroundColor: CAT_COLORS.slice(0, catEntries.length).map(c => c+'BB'),
        borderColor: CAT_COLORS.slice(0, catEntries.length),
        borderWidth: 2, borderRadius: 8, borderSkipped: false }]
    },
    options: {
      indexAxis:'y',
      plugins:{ legend:{display:false}, tooltip: TOOLTIP },
      scales:{ x:{grid:XGRID,ticks:XTICK}, y:{grid:{display:false},ticks:YTICK} }
    }
  });

  // ── 5. Bar apilado: alertas por ubicación (full width)
  const ubAlertData = {};
  rows.forEach(r => {
    const ub = invUbicacionV3(r);
    if (!ubAlertData[ub]) ubAlertData[ub] = { con:0, sin:0 };
    if (invGetAlerta(r, ub)) ubAlertData[ub].con++; else ubAlertData[ub].sin++;
  });
  const ubAlertEntries = Object.entries(ubAlertData).sort((a,b) => b[1].con - a[1].con);
  const c5 = card('det-c-alerta-ubic','Alertas por Ubicación','Registros con/sin alerta desglosados por tipo de ubicación V3', 3);
  grid.appendChild(c5);
  DET_CHARTS.alertaUbic = new Chart(document.getElementById('det-c-alerta-ubic'), {
    type: 'bar',
    data: {
      labels: ubAlertEntries.map(e => e[0]),
      datasets: [
        { label:'Con alerta', data: ubAlertEntries.map(e => e[1].con),
          backgroundColor: '#F8717166', borderColor: '#F87171', borderWidth: 2, borderRadius: 6, borderSkipped: false },
        { label:'Sin alerta', data: ubAlertEntries.map(e => e[1].sin),
          backgroundColor: '#B0F2AE44', borderColor: '#B0F2AE', borderWidth: 2, borderRadius: 6, borderSkipped: false },
      ]
    },
    options: {
      plugins:{ legend: LEGEND, tooltip: TOOLTIP },
      scales:{
        x:{ stacked:true, grid:{display:false}, ticks:{color:'#FAFAFA',font:{family:'Outfit',size:11}}, border:{display:false} },
        y:{ stacked:true, grid:XGRID, ticks:XTICK }
      }
    }
  });
}

// ══════════════════════════════════════════════════════════════════
//  TABLA 1 — Referencia / Total Uds / Uds con alerta ≥8d
// ══════════════════════════════════════════════════════════════════
function _detRenderTabla1() {
  const wrap = document.getElementById('det-tabla1-wrap');
  const cntEl = document.getElementById('det-tabla1-count');
  const pagEl  = document.getElementById('det-tabla1-pag');
  if (!wrap) return;

  // Agrupar por referencia
  const refMap = {};
  DET_FILTERED.forEach(r => {
    const ref  = (r['Nombre']||'Sin nombre').trim();
    const cat  = invCategoria(ref);
    const qty  = parseInt(r['Cantidad'])||0;
    const ub   = invUbicacionV3(r);
    const al   = invGetAlerta(r, ub);
    if (!refMap[ref]) refMap[ref] = { ref, cat, total: 0, conAlerta: 0 };
    refMap[ref].total += qty;
    if (al) refMap[ref].conAlerta += qty;
  });

  const sorted = Object.values(refMap).sort((a,b) => b.total - a.total);
  const pages  = Math.max(1, Math.ceil(sorted.length / DET_PAGE_SIZE));
  if (DET_PAGE_T1 > pages) DET_PAGE_T1 = pages;
  const slice  = sorted.slice((DET_PAGE_T1-1)*DET_PAGE_SIZE, DET_PAGE_T1*DET_PAGE_SIZE);

  if (cntEl) cntEl.textContent = sorted.length + ' referencias';

  const CAT_COLOR = {
    'Datáfonos':'#99D1FC','Rollos':'#B0F2AE','Pin pad':'#FFC04D',
    'Forros':'#C084FC','Accesorios':'#FB923C','SIM':'#F87171','KIT POP VP':'#DFFF61',
  };

  wrap.innerHTML = `
    <table style="width:100%;border-collapse:collapse;font-size:12px;font-family:'Outfit',sans-serif;">
      <thead>
        <tr style="position:sticky;top:0;background:#181715;z-index:2;">
          ${['#','REFERENCIA','CATEGORÍA','TOTAL UDS','UDS CON ALERTA ≥8d','% ALERTA'].map(h =>
            `<th style="padding:10px 14px;text-align:${h==='TOTAL UDS'||h==='UDS CON ALERTA ≥8d'||h==='% ALERTA'?'right':'left'};
              color:#DFFF61;font-weight:700;font-size:9px;letter-spacing:1.2px;text-transform:uppercase;
              border-bottom:1px solid rgba(223,255,97,.15);white-space:nowrap;font-family:'Syne',sans-serif;">${h}</th>`
          ).join('')}
        </tr>
      </thead>
      <tbody>
        ${!slice.length ? `<tr><td colspan="6" style="text-align:center;padding:40px;color:#7A7674;">Sin datos</td></tr>` :
          slice.map((r, i) => {
            const color = CAT_COLOR[r.cat] || '#94a3b8';
            const pct   = r.total ? (r.conAlerta / r.total * 100).toFixed(1) : '0.0';
            const bg    = i%2===0 ? 'rgba(223,255,97,.012)' : 'transparent';
            return `<tr style="background:${bg};transition:background .12s;"
              onmouseover="this.style.background='rgba(223,255,97,.04)'"
              onmouseout="this.style.background='${bg}'">
              <td style="padding:9px 14px;color:rgba(223,255,97,.35);font-family:'JetBrains Mono',monospace;font-size:10px;">${String((DET_PAGE_T1-1)*DET_PAGE_SIZE+i+1).padStart(2,'0')}</td>
              <td style="padding:9px 14px;color:#FAFAFA;font-weight:500;max-width:320px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="${r.ref}">${r.ref}</td>
              <td style="padding:9px 14px;"><span style="display:inline-block;background:${color}22;color:${color};font-size:10px;font-weight:700;padding:2px 9px;border-radius:20px;">${r.cat}</span></td>
              <td style="padding:9px 14px;text-align:right;font-family:'JetBrains Mono',monospace;font-weight:700;color:${color};font-size:13px;">${r.total.toLocaleString('es-CO')}</td>
              <td style="padding:9px 14px;text-align:right;font-family:'JetBrains Mono',monospace;font-weight:700;color:${r.conAlerta > 0 ? '#F87171' : '#B0F2AE'};font-size:13px;">${r.conAlerta.toLocaleString('es-CO')}</td>
              <td style="padding:9px 14px;text-align:right;">
                <div style="display:flex;align-items:center;justify-content:flex-end;gap:8px;">
                  <div style="background:rgba(255,255,255,.05);border-radius:3px;height:4px;width:60px;overflow:hidden;">
                    <div style="height:100%;width:${Math.min(parseFloat(pct),100)}%;background:${parseFloat(pct)>50?'#F87171':'#B0F2AE'};border-radius:3px;"></div>
                  </div>
                  <span style="font-family:'JetBrains Mono',monospace;font-size:11px;color:#7A7674;">${pct}%</span>
                </div>
              </td>
            </tr>`;
          }).join('')
        }
      </tbody>
    </table>`;

  _detRenderPag(pagEl, DET_PAGE_T1, pages, p => { DET_PAGE_T1 = p; _detRenderTabla1(); }, '#DFFF61');
}

// ══════════════════════════════════════════════════════════════════
//  TABLA 2 — Código Ubic / Nombre Ubic / Serial / Cantidad / Fecha Conf / Alerta
// ══════════════════════════════════════════════════════════════════
function _detRenderTabla2() {
  const wrap  = document.getElementById('det-tabla2-wrap');
  const cntEl = document.getElementById('det-tabla2-count');
  const pagEl = document.getElementById('det-tabla2-pag');
  if (!wrap) return;

  const rows   = DET_FILTERED;
  const pages  = Math.max(1, Math.ceil(rows.length / DET_PAGE_SIZE));
  if (DET_PAGE_T2 > pages) DET_PAGE_T2 = pages;
  const slice  = rows.slice((DET_PAGE_T2-1)*DET_PAGE_SIZE, DET_PAGE_T2*DET_PAGE_SIZE);

  if (cntEl) cntEl.textContent = rows.length + ' registros';

  wrap.innerHTML = `
    <table style="width:100%;border-collapse:collapse;font-size:12px;font-family:'Outfit',sans-serif;">
      <thead>
        <tr style="position:sticky;top:0;background:#181715;z-index:2;">
          ${['CÓD. UBICACIÓN','NOMBRE UBICACIÓN','N.° SERIE','REFERENCIA','CANTIDAD','FECHA CONF.','ALERTA'].map(h =>
            `<th style="padding:10px 14px;text-align:${h==='CANTIDAD'?'right':'left'};
              color:#B0F2AE;font-weight:700;font-size:9px;letter-spacing:1.2px;text-transform:uppercase;
              border-bottom:1px solid rgba(176,242,174,.15);white-space:nowrap;font-family:'Syne',sans-serif;">${h}</th>`
          ).join('')}
        </tr>
      </thead>
      <tbody>
        ${!slice.length ? `<tr><td colspan="7" style="text-align:center;padding:40px;color:#7A7674;">Sin datos</td></tr>` :
          slice.map((r, i) => {
            const ub      = invUbicacionV3(r);
            const alerta  = invGetAlerta(r, ub);
            const dias    = invDiasTranscurridos(r);
            const serial  = r['Número de serie']||r['Numero de serie']||r['Serial']||'—';
            const codUbic = r['Código de ubicación']||'—';
            const nomUbic = r['Nombre de la ubicación']||'—';
            const fecha   = invFmtFecha(invFechaConfirmacion(r));
            const qty     = parseInt(r['Cantidad'])||0;
            const nombre  = r['Nombre']||'—';
            const bg      = i%2===0 ? 'rgba(176,242,174,.012)' : 'transparent';
            return `<tr style="background:${bg};transition:background .12s;"
              onmouseover="this.style.background='rgba(176,242,174,.04)'"
              onmouseout="this.style.background='${bg}'">
              <td style="padding:9px 14px;color:#67e8f9;font-family:'JetBrains Mono',monospace;font-size:11px;">${codUbic}</td>
              <td style="padding:9px 14px;color:#cbd5e1;font-size:11px;max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="${nomUbic}">${nomUbic}</td>
              <td style="padding:9px 14px;color:#a5f3fc;font-family:'JetBrains Mono',monospace;font-size:11px;max-width:160px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="${serial}">${serial}</td>
              <td style="padding:9px 14px;color:#FAFAFA;font-size:11px;max-width:220px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="${nombre}">${nombre}</td>
              <td style="padding:9px 14px;text-align:right;font-family:'JetBrains Mono',monospace;font-weight:700;color:#B0F2AE;font-size:13px;">${qty.toLocaleString('es-CO')}</td>
              <td style="padding:9px 14px;color:#7A7674;font-size:11px;white-space:nowrap;">
                ${fecha}
                ${dias !== null ? `<span style="font-size:10px;color:${dias>8?'#F87171':'#7A7674'};margin-left:6px;font-family:'JetBrains Mono',monospace;">(${dias}d)</span>` : ''}
              </td>
              <td style="padding:9px 14px;">
                ${alerta
                  ? `<span style="display:inline-flex;align-items:center;gap:4px;background:#F8717122;color:#F87171;
                      font-size:10px;font-weight:700;padding:3px 10px;border-radius:20px;border:1px solid #F8717133;
                      white-space:nowrap;">🚨 Con alerta</span>`
                  : `<span style="display:inline-flex;align-items:center;gap:4px;background:#B0F2AE15;color:#B0F2AE;
                      font-size:10px;font-weight:700;padding:3px 10px;border-radius:20px;border:1px solid #B0F2AE22;
                      white-space:nowrap;">✅ Sin alerta</span>`
                }
              </td>
            </tr>`;
          }).join('')
        }
      </tbody>
    </table>`;

  _detRenderPag(pagEl, DET_PAGE_T2, pages, p => { DET_PAGE_T2 = p; _detRenderTabla2(); }, '#B0F2AE');
}

// ── Paginación helper ─────────────────────────────────────────────
function _detRenderPag(el, cur, total, onPage, accent) {
  if (!el) return;
  el.innerHTML = '';
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
//  EXPORT EXCEL TABLA 2
// ══════════════════════════════════════════════════════════════════
window.detExportTabla2Excel = function() {
  if (!DET_FILTERED.length) { alert('Sin datos.'); return; }
  if (typeof XLSX === 'undefined') { alert('XLSX no disponible.'); return; }
  const exportRows = DET_FILTERED.map(r => {
    const ub    = invUbicacionV3(r);
    const alerta = invGetAlerta(r, ub);
    const dias   = invDiasTranscurridos(r);
    return {
      'Código Ubicación':   r['Código de ubicación']||'',
      'Nombre Ubicación':   r['Nombre de la ubicación']||'',
      'N.° de Serie':       r['Número de serie']||r['Numero de serie']||r['Serial']||'',
      'Referencia':         r['Nombre']||'',
      'Cantidad':           parseInt(r['Cantidad'])||0,
      'Fecha Confirmación': invFmtFecha(invFechaConfirmacion(r)),
      'Días transcurridos': dias !== null ? dias : '',
      'Alerta':             alerta ? 'Con alerta' : 'Sin alerta',
      'Ubicación V3':       ub,
      'Negocio':            invNegocio(r['Subtipo']),
      'Categoría':          invCategoria(r['Nombre']),
    };
  });
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(exportRows);
  ws['!cols'] = Object.keys(exportRows[0]).map(k => ({ wch: Math.max(k.length+2, 12) }));
  XLSX.utils.book_append_sheet(wb, ws, 'Detalles Inventario');
  XLSX.writeFile(wb, `Detalles_Inventario_${new Date().toISOString().slice(0,10)}.xlsx`);
};

window.detExportTabla1Excel = function() {
  if (!DET_FILTERED.length) { alert('Sin datos.'); return; }
  if (typeof XLSX === 'undefined') { alert('XLSX no disponible.'); return; }
  const refMap = {};
  DET_FILTERED.forEach(r => {
    const ref = (r['Nombre']||'Sin nombre').trim();
    const cat = invCategoria(ref);
    const qty = parseInt(r['Cantidad'])||0;
    const ub  = invUbicacionV3(r);
    const al  = invGetAlerta(r, ub);
    if (!refMap[ref]) refMap[ref] = { Referencia: ref, Categoría: cat, 'Total Uds': 0, 'Uds con Alerta': 0 };
    refMap[ref]['Total Uds'] += qty;
    if (al) refMap[ref]['Uds con Alerta'] += qty;
  });
  const rows = Object.values(refMap).sort((a,b) => b['Total Uds'] - a['Total Uds'])
    .map(r => ({ ...r, '% Alerta': r['Total Uds'] ? (r['Uds con Alerta']/r['Total Uds']*100).toFixed(1)+'%' : '0%' }));
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(rows);
  XLSX.utils.book_append_sheet(wb, ws, 'Referencias Detalles');
  XLSX.writeFile(wb, `Referencias_Detalles_${new Date().toISOString().slice(0,10)}.xlsx`);
};

// ══════════════════════════════════════════════════════════════════
//  ENTRY POINT
// ══════════════════════════════════════════════════════════════════
async function renderInventarioDetalles() {
  const panel = document.getElementById('panel-inv-detalles');
  if (!panel) return;

  const loadingEl = document.getElementById('det-loading');
  const contentEl = document.getElementById('det-content');

  // Si los datos aún no están, esperarlos
  if (!window.INV_RAW) {
    if (loadingEl) loadingEl.style.display = 'flex';
    if (contentEl) contentEl.style.display = 'none';
    if (typeof window.loadInventarioData === 'function') {
      await window.loadInventarioData();
    }
  }

  // Siempre ocultar loading y mostrar contenido
  if (loadingEl) loadingEl.style.display = 'none';
  if (contentEl) contentEl.style.display = '';

  DET_FILTERED = (window.INV_RAW || []).slice();
  DET_PAGE_T1 = 1;
  DET_PAGE_T2 = 1;
  detPopulateFilters();
  detRenderAll();
}

window.renderInventarioDetalles = renderInventarioDetalles;