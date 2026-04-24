/**
 * ╔══════════════════════════════════════════════════════════════════╗
 * ║  rollos_inventario.js — Módulo Inventario de Rollos Wompi       ║
 * ║  Fuente: wompi_tablero_rollos_calculos (via data_rollos.json.gz) ║
 * ║                                                                  ║
 * ║  Indicadores:                                                    ║
 * ║  · Cobertura por corresponsal (meses)                            ║
 * ║  · Consumo real diario / semanal / mensual                       ║
 * ║  · Consumo real vs proyectado                                    ║
 * ║  · Inventario total y por corresponsal                           ║
 * ║  · Rotación de inventario                                        ║
 * ║  · SLA 3 meses cobertura                                         ║
 * ║  · Punto de reorden                                              ║
 * ║  · Alertas automáticas                                           ║
 * ║  · Riesgo de quiebre                                             ║
 * ╚══════════════════════════════════════════════════════════════════╝
 */

'use strict';

// ──────────────────────────────────────────────────────────────────
//  ESTADO DEL MÓDULO
// ──────────────────────────────────────────────────────────────────
let RI_DATA = [];   // rows from calculos (one per sitio/corresponsal)
let RI_FILTERED = [];   // after global filters
let riPage = 1;
let riSearchTerm = '';
let riSortCol = -1;
let riSortDir = 1;
const RI_PAGE_SIZE = 50;

// Chart instances for this module
let riCharts = {};

// ──────────────────────────────────────────────────────────────────
//  BOOTSTRAP — called once ROLLOS_RAW is ready
// ──────────────────────────────────────────────────────────────────
window.initRollosInventario = function () {
    console.log('[Rollos-Inv] Iniciando módulo...');
    if (!window.ROLLOS_RAW) {
        console.warn('[Rollos-Inv] ROLLOS_RAW no disponible todavía.');
        return;
    }

    // The calculos data lives at ROLLOS_RAW.calculos (array)
    // Fallback to comercio if calculos is empty
    let calc = window.ROLLOS_RAW.calculos || [];
    if (!calc.length && window.ROLLOS_RAW.comercio) {
        console.log('[Rollos-Inv] Usando fallback: comercio');
        calc = window.ROLLOS_RAW.comercio;
    }
    console.log('[Rollos-Inv] Filas a procesar:', calc.length);

    // Merge with sitio metadata from detalle if needed
    const detBySitio = new Map();
    (window.ROLLOS_RAW.detalle || []).forEach(r => {
        if (r.cod_sitio && !detBySitio.has(r.cod_sitio)) {
            detBySitio.set(r.cod_sitio, {
                nombre_sitio: r.nombre_sitio || '',
                departamento: r.departamento || '',
                ciudad: r.ciudad || '',
                proyecto: r.proyecto || '',
            });
        }
    });

    RI_DATA = calc.map(r => {
        const sitId = r.codigo_sitio || r.cod_sitio || r.cod_comercio || '';
        const meta = detBySitio.get(sitId) || {};
        const saldoDias = parseFloat(r.saldo_dias || r.cal_saldo_dias || 0);
        const saldoMeses = saldoDias / 30;
        const promMes = parseFloat(r.promedio_mensual || r.cal_promedio_mensual || 0);
        const promDia = promMes / 30;
        const promSem = promDia * 7;
        const saldoRollos = parseFloat(r.saldo_rollos || r.cal_saldo_rollos || 0);
        const puntoReorden = parseFloat(r.punto_reorden || r.cal_punto_reorden || 0);
        const periodoAbast = parseFloat(r.periodo_abast_e5 || r.cal_periodo_abast_e5 || 0);
        const rollEntregados = parseFloat(r.rollos_entregados_mig_apert || r.cal_rollos_entregados_mig_apert || 0);
        const rollConsumidos = parseFloat(r.rollos_consumidos_migr_apert || r.cal_rollos_consumidos_migr_apert || 0);
        const trxDesde = parseFloat(r.trx_desde_migra_apert || r.cal_trx_desde_migra_apert || 0);
        const rollProyectados = parseFloat(r.rollos_periodo_abast_e5 || r.cal_rollos_periodo_abast_e5 || 0);
        const rollAnio = parseFloat(r.rollos_anio_e5 || r.cal_rollos_anio_e5 || 0);
        const valorBusqueda = parseFloat(r.valor_busqueda || r.cal_valor_busqueda || 0);
        const indVariacion = parseFloat(r.ind_variacion_consumo_pct || 0);
        const indAlerta = r.ind_alerta_umbral || '';
        const indRotacion = parseFloat(r.ind_rotacion_inventario || 0);
        const indRiesgo = r.ind_riesgo_quiebre || '';
        const estadoPunto = r.estado_punto || r.cal_estado_punto || '';
        const fechaApertura = r.fecha_apertura_final || r.cal_fecha_apertura_final || '';
        const fechaAbst = r.fecha_abst_1 || r.cal_fecha_abst_1 || '';

        // Cobertura semaforo
        let coberturaSem = 'ok';
        if (saldoMeses < 1) coberturaSem = 'critico';
        else if (saldoMeses < 2) coberturaSem = 'alerta';
        else if (saldoMeses < 3) coberturaSem = 'warn';

        // SLA 3 meses
        const slaCumple = saldoMeses >= 3;

        // Punto reorden check
        const bajoPuntoReorden = saldoRollos <= puntoReorden && puntoReorden > 0;

        // Variación consumo
        const variacionAbs = Math.abs(indVariacion);
        const variacionLabel = indVariacion > 0 ? `+${indVariacion.toFixed(1)}%` : `${indVariacion.toFixed(1)}%`;

        return {
            // IDs
            id: r.id || '',
            tarea: r.tarea || '',
            codigo_mo: r.codigo_mo || r.cal_codigo_mo || '',
            codigo_sitio: r.codigo_sitio || r.cal_codigo_mo || '',
            // Metadata enriquecida
            nombre_sitio: meta.nombre_sitio,
            departamento: meta.departamento,
            ciudad: meta.ciudad,
            proyecto: meta.proyecto,
            estado_punto: estadoPunto,
            fecha_apertura: fechaApertura,
            fecha_abst: fechaAbst,
            // Consumo
            prom_mensual: promMes,
            prom_diario: promDia,
            prom_semanal: promSem,
            rollos_prom_mes: parseFloat(r.rollos_promedio_mes || r.cal_rollos_promedio_mes || 0),
            // Inventario
            saldo_rollos: saldoRollos,
            saldo_dias: saldoDias,
            saldo_meses: saldoMeses,
            saldo_valor: parseFloat(r.saldo || r.cal_saldo || 0),
            // Proyecciones
            periodo_abast: periodoAbast,
            rollos_proyect: rollProyectados,
            rollos_anio: rollAnio,
            valor_busqueda: valorBusqueda,
            punto_reorden: puntoReorden,
            // Ejecutado
            roll_entregados: rollEntregados,
            roll_consumidos: rollConsumidos,
            trx_desde: trxDesde,
            // Indicadores
            ind_variacion: indVariacion,
            variacion_label: variacionLabel,
            variacion_abs: variacionAbs,
            ind_alerta: indAlerta,
            ind_rotacion: indRotacion,
            ind_riesgo: indRiesgo,
            // Semaforos
            cobertura_sem: coberturaSem,
            sla_cumple: slaCumple,
            bajo_reorden: bajoPuntoReorden,
        };
    }).filter(r => r.saldo_rollos > 0 || r.prom_mensual > 0); // skip ghost rows
    console.log('[Rollos-Inv] Corresponsales procesados:', RI_DATA.length);

    RI_FILTERED = RI_DATA.slice();
};

// ──────────────────────────────────────────────────────────────────
//  RENDER PRINCIPAL
// ──────────────────────────────────────────────────────────────────
window.renderRollosInventario = function () {
    if (!RI_DATA.length && window.ROLLOS_RAW) window.initRollosInventario();
    RI_FILTERED = RI_DATA.slice();
    riApplySearch('');

    _renderRIKPIs();
    _renderRIAlerts();
    _renderRICharts();
    _renderRITable();
};

// ──────────────────────────────────────────────────────────────────
//  KPIs STRIP
// ──────────────────────────────────────────────────────────────────
function _renderRIKPIs() {
    const el = document.getElementById('ri-kpi-strip');
    if (!el) return;

    const d = RI_FILTERED;
    const total = d.length;
    const slaOk = d.filter(r => r.sla_cumple).length;
    const criticos = d.filter(r => r.cobertura_sem === 'critico').length;
    const alertas = d.filter(r => r.cobertura_sem === 'alerta').length;
    const warn = d.filter(r => r.cobertura_sem === 'warn').length;
    const enRiesgo = d.filter(r => r.ind_riesgo === 'ALTO' || r.ind_riesgo === 'CRÍTICO').length;
    const bajoPto = d.filter(r => r.bajo_reorden).length;
    const totalSaldo = d.reduce((s, r) => s + r.saldo_rollos, 0);
    const avgMeses = total > 0 ? d.reduce((s, r) => s + r.saldo_meses, 0) / total : 0;
    const pctSla = total > 0 ? Math.round(slaOk / total * 100) : 0;
    const totalConsumo = d.reduce((s, r) => s + r.prom_mensual, 0);
    const rotProm = total > 0 ? d.reduce((s, r) => s + r.ind_rotacion, 0) / total : 0;

    const kpis = [
        { label: 'Total Corresponsales', value: total.toLocaleString('es-CO'), icon: '🏪', color: '#99D1FC', bg: 'rgba(153,209,252,.08)' },
        { label: 'Saldo Total Rollos', value: Math.round(totalSaldo).toLocaleString('es-CO'), icon: '📦', color: '#B0F2AE', bg: 'rgba(176,242,174,.08)' },
        { label: 'Cobertura Promedio', value: avgMeses.toFixed(1) + ' meses', icon: '📅', color: '#DFFF61', bg: 'rgba(223,255,97,.08)' },
        { label: 'Consumo Mensual Total', value: Math.round(totalConsumo).toLocaleString('es-CO') + ' rollos', icon: '📊', color: '#F49D6E', bg: 'rgba(244,157,110,.08)' },
        { label: 'SLA ≥3 Meses', value: pctSla + '%', sub: `${slaOk}/${total}`, icon: '🎯', color: pctSla >= 70 ? '#B0F2AE' : pctSla >= 50 ? '#DFFF61' : '#FF5C5C', bg: 'rgba(176,242,174,.06)', pct: pctSla },
        { label: 'En Riesgo Quiebre', value: enRiesgo.toLocaleString(), icon: '🚨', color: enRiesgo > 0 ? '#FF5C5C' : '#B0F2AE', bg: 'rgba(255,92,92,.08)' },
        { label: 'Bajo Punto Reorden', value: bajoPto.toLocaleString(), icon: '⚠️', color: bajoPto > 0 ? '#FFC04D' : '#B0F2AE', bg: 'rgba(255,192,77,.08)' },
        { label: 'Rotación Prom.', value: rotProm.toFixed(2) + 'x', icon: '🔄', color: '#7B8CDE', bg: 'rgba(123,140,222,.08)' },
    ];

    el.innerHTML = kpis.map(k => `
    <div style="background:${k.bg};border:1px solid ${k.color}22;border-radius:14px;padding:16px 18px;min-width:160px;flex:1;position:relative;overflow:hidden;transition:transform .2s,box-shadow .2s;"
         onmouseover="this.style.transform='translateY(-2px)';this.style.boxShadow='0 8px 24px rgba(0,0,0,.4)'"
         onmouseout="this.style.transform='';this.style.boxShadow=''">
      <div style="font-size:20px;margin-bottom:6px;">${k.icon}</div>
      <div style="font-family:'JetBrains Mono',monospace;font-size:20px;font-weight:700;color:${k.color};line-height:1;">${k.value}</div>
      ${k.sub ? `<div style="font-size:10px;color:#64748b;margin-top:2px;">${k.sub}</div>` : ''}
      ${k.pct !== undefined ? `
        <div style="margin-top:8px;background:rgba(255,255,255,.06);border-radius:4px;height:4px;overflow:hidden;">
          <div style="width:${k.pct}%;height:100%;background:${k.color};border-radius:4px;transition:width .6s ease;"></div>
        </div>` : ''}
      <div style="font-size:10px;color:#7A7674;margin-top:6px;font-family:'Outfit',sans-serif;">${k.label}</div>
    </div>`).join('');
}

// ──────────────────────────────────────────────────────────────────
//  ALERTAS AUTOMÁTICAS
// ──────────────────────────────────────────────────────────────────
function _renderRIAlerts() {
    const el = document.getElementById('ri-alerts-container');
    if (!el) return;

    const criticos = RI_FILTERED.filter(r => r.cobertura_sem === 'critico').sort((a, b) => a.saldo_meses - b.saldo_meses).slice(0, 10);
    const sinSla = RI_FILTERED.filter(r => !r.sla_cumple).length;
    const bajoPto = RI_FILTERED.filter(r => r.bajo_reorden).length;
    const riesgoAlto = RI_FILTERED.filter(r => r.ind_riesgo === 'ALTO' || r.ind_riesgo === 'CRÍTICO').length;

    let html = `
    <div style="display:flex;gap:12px;flex-wrap:wrap;margin-bottom:18px;">
      ${sinSla > 0 ? `<div style="background:rgba(255,92,92,.12);border:1px solid rgba(255,92,92,.35);border-radius:10px;padding:10px 16px;font-family:'Outfit',sans-serif;font-size:12px;color:#FF5C5C;">
        🚨 <strong>${sinSla}</strong> corresponsal${sinSla > 1 ? 'es' : ''} sin cumplir SLA 3 meses</div>` : ''}
      ${bajoPto > 0 ? `<div style="background:rgba(255,192,77,.12);border:1px solid rgba(255,192,77,.35);border-radius:10px;padding:10px 16px;font-family:'Outfit',sans-serif;font-size:12px;color:#FFC04D;">
        ⚠️ <strong>${bajoPto}</strong> bajo punto de reorden</div>` : ''}
      ${riesgoAlto > 0 ? `<div style="background:rgba(255,92,92,.12);border:1px solid rgba(255,92,92,.35);border-radius:10px;padding:10px 16px;font-family:'Outfit',sans-serif;font-size:12px;color:#FF5C5C;">
        🔴 <strong>${riesgoAlto}</strong> en riesgo de quiebre alto/crítico</div>` : ''}
      ${sinSla === 0 && bajoPto === 0 && riesgoAlto === 0 ? `<div style="background:rgba(176,242,174,.08);border:1px solid rgba(176,242,174,.25);border-radius:10px;padding:10px 16px;font-family:'Outfit',sans-serif;font-size:12px;color:#B0F2AE;">
        ✅ Sin alertas activas — todos los indicadores en rango normal</div>` : ''}
    </div>`;

    if (criticos.length > 0) {
        html += `
    <div style="margin-bottom:8px;font-family:'Syne',sans-serif;font-size:12px;font-weight:700;color:#FF5C5C;letter-spacing:.5px;text-transform:uppercase;">
      🚨 Corresponsales Críticos — Cobertura &lt; 1 Mes
    </div>
    <div style="display:flex;gap:10px;flex-wrap:wrap;">
      ${criticos.map(r => `
        <div style="background:rgba(255,92,92,.08);border:1px solid rgba(255,92,92,.3);border-radius:10px;padding:10px 14px;min-width:160px;font-family:'Outfit',sans-serif;">
          <div style="font-size:10px;color:#64748b;">${r.codigo_sitio}</div>
          <div style="font-size:12px;font-weight:600;color:#f1f5f9;margin:2px 0;">${r.nombre_sitio || 'Sin nombre'}</div>
          <div style="font-size:18px;font-weight:700;color:#FF5C5C;">${r.saldo_meses.toFixed(1)} <span style="font-size:11px;">meses</span></div>
          <div style="font-size:10px;color:#FF8888;">Saldo: ${Math.round(r.saldo_rollos)} rollos</div>
          <div style="margin-top:6px;background:rgba(255,92,92,.15);border-radius:3px;height:3px;overflow:hidden;">
            <div style="width:${Math.min(r.saldo_meses / 3 * 100, 100)}%;height:100%;background:#FF5C5C;border-radius:3px;"></div>
          </div>
        </div>`).join('')}
    </div>`;
    }

    el.innerHTML = html;
}

// ──────────────────────────────────────────────────────────────────
//  GRÁFICAS
// ──────────────────────────────────────────────────────────────────
function _renderRICharts() {
    _destroyRICharts();
    _chartCoberturaBuckets();
    _chartConsumoVsProyectado();
    _chartRotacion();
    _chartTopSaldos();
    _chartSLADistrib();
    _chartCoberturaHistogram();
}

function _destroyRICharts() {
    Object.values(riCharts).forEach(c => { try { c.destroy(); } catch (e) { } });
    riCharts = {};
}

const RICHART_OPTS = {
    plugins: { tooltip: { backgroundColor: 'rgba(24,23,21,.95)', titleColor: '#99D1FC', bodyColor: '#FAFAFA', borderColor: 'rgba(153,209,252,.2)', borderWidth: 1, padding: 12 } },
    animation: { duration: 600 },
};

function _chartCoberturaBuckets() {
    const ctx = document.getElementById('ri-chart-cobertura-buckets');
    if (!ctx) return;
    const buckets = { '<1m': 0, '1-2m': 0, '2-3m': 0, '3-6m': 0, '>6m': 0 };
    RI_FILTERED.forEach(r => {
        const m = r.saldo_meses;
        if (m < 1) buckets['<1m']++;
        else if (m < 2) buckets['1-2m']++;
        else if (m < 3) buckets['2-3m']++;
        else if (m < 6) buckets['3-6m']++;
        else buckets['>6m']++;
    });
    riCharts.cobBuckets = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: Object.keys(buckets),
            datasets: [{
                label: 'Corresponsales', data: Object.values(buckets),
                backgroundColor: ['#FF5C5C', '#FFC04D', '#DFFF61', '#B0F2AE', '#99D1FC'],
                borderRadius: 6, borderWidth: 0
            }]
        },
        options: {
            ...RICHART_OPTS, responsive: true, maintainAspectRatio: false,
            scales: {
                x: { grid: { color: 'rgba(255,255,255,.04)' }, ticks: { color: '#94a3b8', font: { size: 11 } } },
                y: { grid: { color: 'rgba(255,255,255,.04)' }, ticks: { color: '#94a3b8', font: { size: 11 } } }
            },
            plugins: {
                ...RICHART_OPTS.plugins, legend: { display: false },
                title: { display: false }
            }
        }
    });
}

function _chartConsumoVsProyectado() {
    const ctx = document.getElementById('ri-chart-consumo-vs-proy');
    if (!ctx) return;
    // Top 12 sitios por consumo mensual
    const top = RI_FILTERED.filter(r => r.prom_mensual > 0).sort((a, b) => b.prom_mensual - a.prom_mensual).slice(0, 12);
    riCharts.consumoProy = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: top.map(r => r.codigo_sitio || r.nombre_sitio || r.codigo_mo || 'N/A'),
            datasets: [
                { label: 'Consumo Real (mes)', data: top.map(r => r.prom_mensual), backgroundColor: 'rgba(176,242,174,.7)', borderRadius: 4, borderWidth: 0 },
                { label: 'Proyectado (período)', data: top.map(r => r.rollos_proyect > 0 ? r.rollos_proyect / Math.max(r.periodo_abast, 1) : 0), backgroundColor: 'rgba(153,209,252,.5)', borderRadius: 4, borderWidth: 0 },
            ]
        },
        options: {
            ...RICHART_OPTS, responsive: true, maintainAspectRatio: false,
            scales: {
                x: { grid: { color: 'rgba(255,255,255,.04)' }, ticks: { color: '#94a3b8', font: { size: 9 }, maxRotation: 45 } },
                y: { grid: { color: 'rgba(255,255,255,.04)' }, ticks: { color: '#94a3b8', font: { size: 11 } } }
            },
            plugins: { ...RICHART_OPTS.plugins, legend: { labels: { color: '#94a3b8', font: { size: 10 } } } }
        }
    });
}

function _chartRotacion() {
    const ctx = document.getElementById('ri-chart-rotacion');
    if (!ctx) return;
    const buckets = { '0-0.5': 0, '0.5-1': 0, '1-2': 0, '2-3': 0, '>3': 0 };
    RI_FILTERED.forEach(r => {
        const v = r.ind_rotacion;
        if (v <= 0.5) buckets['0-0.5']++;
        else if (v <= 1) buckets['0.5-1']++;
        else if (v <= 2) buckets['1-2']++;
        else if (v <= 3) buckets['2-3']++;
        else buckets['>3']++;
    });
    riCharts.rotacion = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels: Object.keys(buckets),
            datasets: [{
                data: Object.values(buckets),
                backgroundColor: ['#FF5C5C', '#FFC04D', '#DFFF61', '#B0F2AE', '#99D1FC'],
                borderWidth: 0, hoverOffset: 8
            }]
        },
        options: {
            ...RICHART_OPTS, responsive: true, maintainAspectRatio: false, cutout: '62%',
            plugins: { ...RICHART_OPTS.plugins, legend: { position: 'bottom', labels: { color: '#94a3b8', font: { size: 10 }, padding: 10 } } }
        }
    });
}

function _chartTopSaldos() {
    const ctx = document.getElementById('ri-chart-top-saldos');
    if (!ctx) return;
    const top = RI_FILTERED.filter(r => r.saldo_rollos > 0).sort((a, b) => b.saldo_rollos - a.saldo_rollos).slice(0, 15);
    riCharts.topSaldos = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: top.map(r => r.codigo_sitio || r.codigo_mo || 'N/A'),
            datasets: [{
                label: 'Saldo Rollos',
                data: top.map(r => r.saldo_rollos),
                backgroundColor: top.map(r =>
                    r.cobertura_sem === 'critico' ? 'rgba(255,92,92,.75)' :
                        r.cobertura_sem === 'alerta' ? 'rgba(255,192,77,.75)' :
                            r.cobertura_sem === 'warn' ? 'rgba(223,255,97,.65)' :
                                'rgba(176,242,174,.7)'),
                borderRadius: 4, borderWidth: 0
            }]
        },
        options: {
            ...RICHART_OPTS, responsive: true, maintainAspectRatio: false, indexAxis: 'y',
            scales: {
                x: { grid: { color: 'rgba(255,255,255,.04)' }, ticks: { color: '#94a3b8', font: { size: 10 } } },
                y: { grid: { color: 'rgba(255,255,255,.04)' }, ticks: { color: '#94a3b8', font: { size: 9 } } }
            },
            plugins: { ...RICHART_OPTS.plugins, legend: { display: false } }
        }
    });
}

function _chartSLADistrib() {
    const ctx = document.getElementById('ri-chart-sla');
    if (!ctx) return;
    const slaOk = RI_FILTERED.filter(r => r.sla_cumple).length;
    const slaNo = RI_FILTERED.length - slaOk;
    riCharts.sla = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels: ['Cumple SLA ≥3m', 'No Cumple'],
            datasets: [{
                data: [slaOk, slaNo],
                backgroundColor: ['rgba(176,242,174,.8)', 'rgba(255,92,92,.7)'],
                borderWidth: 0, hoverOffset: 8
            }]
        },
        options: {
            ...RICHART_OPTS, responsive: true, maintainAspectRatio: false, cutout: '65%',
            plugins: {
                ...RICHART_OPTS.plugins,
                legend: { position: 'bottom', labels: { color: '#94a3b8', font: { size: 11 }, padding: 12 } }
            }
        }
    });
}

function _chartCoberturaHistogram() {
    const ctx = document.getElementById('ri-chart-cobertura-hist');
    if (!ctx) return;
    // Distribución de cobertura: agrupar por rango de 0.5 meses hasta 12
    const bins = {};
    for (let i = 0; i <= 12; i += 0.5) bins[i.toFixed(1)] = 0;
    RI_FILTERED.forEach(r => {
        const bin = Math.min(Math.floor(r.saldo_meses * 2) / 2, 12).toFixed(1);
        if (bins[bin] !== undefined) bins[bin]++;
        else bins['12.0'] = (bins['12.0'] || 0) + 1;
    });
    riCharts.cobHist = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: Object.keys(bins),
            datasets: [{
                label: 'Corresponsales', data: Object.values(bins),
                backgroundColor: Object.keys(bins).map(k => {
                    const v = parseFloat(k);
                    return v < 1 ? 'rgba(255,92,92,.7)' : v < 2 ? 'rgba(255,192,77,.7)' : v < 3 ? 'rgba(223,255,97,.65)' : 'rgba(176,242,174,.7)';
                }),
                borderRadius: 3, borderWidth: 0
            }]
        },
        options: {
            ...RICHART_OPTS, responsive: true, maintainAspectRatio: false,
            scales: {
                x: {
                    grid: { color: 'rgba(255,255,255,.04)' }, ticks: { color: '#94a3b8', font: { size: 9 } },
                    title: { display: true, text: 'Meses de cobertura', color: '#64748b', font: { size: 10 } }
                },
                y: { grid: { color: 'rgba(255,255,255,.04)' }, ticks: { color: '#94a3b8', font: { size: 10 } } }
            },
            plugins: {
                ...RICHART_OPTS.plugins, legend: { display: false },
                annotation: {
                    annotations: {
                        sla: {
                            type: 'line', xMin: 6, xMax: 6, borderColor: '#B0F2AE', borderWidth: 2, borderDash: [4, 4],
                            label: { content: 'SLA 3m', display: true, color: '#B0F2AE', font: { size: 9 } }
                        }
                    }
                }
            }
        }
    });
}

// ──────────────────────────────────────────────────────────────────
//  TABLA PRINCIPAL
// ──────────────────────────────────────────────────────────────────
function riApplySearch(term) {
    riSearchTerm = (term || '').toLowerCase();
    riPage = 1;
    const src = RI_FILTERED;
    const filtered = riSearchTerm
        ? src.filter(r =>
            (r.codigo_sitio || '').toLowerCase().includes(riSearchTerm) ||
            (r.nombre_sitio || '').toLowerCase().includes(riSearchTerm) ||
            (r.codigo_mo || '').toLowerCase().includes(riSearchTerm) ||
            (r.departamento || '').toLowerCase().includes(riSearchTerm) ||
            (r.ciudad || '').toLowerCase().includes(riSearchTerm) ||
            (r.estado_punto || '').toLowerCase().includes(riSearchTerm))
        : src;
    _renderRITableData(filtered);
}
window.riApplySearch = riApplySearch;

function _renderRITable() {
    _renderRITableData(RI_FILTERED);
}

function _renderRITableData(data) {
    const wrap = document.getElementById('ri-table-wrap');
    const countEl = document.getElementById('ri-table-count');
    const pagEl = document.getElementById('ri-table-pagination');
    if (!wrap) return;

    const totalPages = Math.max(1, Math.ceil(data.length / RI_PAGE_SIZE));
    if (riPage > totalPages) riPage = 1;

    if (countEl) countEl.textContent = `${data.length.toLocaleString('es-CO')} corresponsales`;

    const start = (riPage - 1) * RI_PAGE_SIZE;
    const slice = data.slice(start, start + RI_PAGE_SIZE);

    const semColor = sem =>
        sem === 'critico' ? '#FF5C5C' : sem === 'alerta' ? '#FFC04D' : sem === 'warn' ? '#DFFF61' : '#B0F2AE';
    const semBg = sem =>
        sem === 'critico' ? 'rgba(255,92,92,.12)' : sem === 'alerta' ? 'rgba(255,192,77,.10)' : sem === 'warn' ? 'rgba(223,255,97,.08)' : 'rgba(176,242,174,.06)';
    const semIcon = sem =>
        sem === 'critico' ? '🔴' : sem === 'alerta' ? '🟠' : sem === 'warn' ? '🟡' : '🟢';

    const fmt = (n, d = 0) => typeof n === 'number' && !isNaN(n) ? n.toLocaleString('es-CO', { maximumFractionDigits: d }) : '—';

    wrap.innerHTML = `
    <table style="width:100%;border-collapse:collapse;font-family:'Outfit',sans-serif;font-size:12px;">
      <thead>
        <tr style="background:rgba(0,0,0,.3);border-bottom:1px solid rgba(153,209,252,.15);">
          <th style="padding:10px 12px;text-align:left;color:#64748b;font-weight:600;font-size:11px;white-space:nowrap;">CORRESPONSAL</th>
          <th style="padding:10px 8px;text-align:center;color:#64748b;font-weight:600;font-size:11px;">COBERTURA</th>
          <th style="padding:10px 8px;text-align:right;color:#64748b;font-weight:600;font-size:11px;">SALDO ROLLOS</th>
          <th style="padding:10px 8px;text-align:right;color:#64748b;font-weight:600;font-size:11px;">CONS. MES</th>
          <th style="padding:10px 8px;text-align:right;color:#64748b;font-weight:600;font-size:11px;">CONS. DÍA</th>
          <th style="padding:10px 8px;text-align:right;color:#64748b;font-weight:600;font-size:11px;">CONS. SEM.</th>
          <th style="padding:10px 8px;text-align:right;color:#64748b;font-weight:600;font-size:11px;">PROYECT.</th>
          <th style="padding:10px 8px;text-align:right;color:#64748b;font-weight:600;font-size:11px;">PTO. REORDEN</th>
          <th style="padding:10px 8px;text-align:center;color:#64748b;font-weight:600;font-size:11px;">VARIACIÓN</th>
          <th style="padding:10px 8px;text-align:center;color:#64748b;font-weight:600;font-size:11px;">ROTACIÓN</th>
          <th style="padding:10px 8px;text-align:center;color:#64748b;font-weight:600;font-size:11px;">SLA 3M</th>
          <th style="padding:10px 8px;text-align:center;color:#64748b;font-weight:600;font-size:11px;">RIESGO</th>
        </tr>
      </thead>
      <tbody>
        ${slice.map((r, i) => `
          <tr style="border-bottom:1px solid rgba(255,255,255,.04);background:${i % 2 === 0 ? 'transparent' : 'rgba(255,255,255,.015)'};transition:background .15s;"
              onmouseover="this.style.background='rgba(153,209,252,.06)'"
              onmouseout="this.style.background='${i % 2 === 0 ? 'transparent' : 'rgba(255,255,255,.015)'}'"
          >
            <td style="padding:9px 12px;">
              <div style="color:#f1f5f9;font-weight:600;font-size:11px;">${r.codigo_sitio || r.codigo_mo || '—'}</div>
              <div style="color:#64748b;font-size:10px;">${r.nombre_sitio || ''}</div>
              <div style="color:#475569;font-size:9px;">${r.departamento ? r.departamento + (r.ciudad ? ' · ' + r.ciudad : '') : ''}</div>
            </td>
            <td style="padding:9px 8px;text-align:center;">
              <div style="display:inline-flex;align-items:center;gap:5px;background:${semBg(r.cobertura_sem)};border:1px solid ${semColor(r.cobertura_sem)}44;border-radius:20px;padding:4px 10px;">
                <span style="font-size:9px;">${semIcon(r.cobertura_sem)}</span>
                <span style="color:${semColor(r.cobertura_sem)};font-weight:700;font-size:12px;">${r.saldo_meses.toFixed(1)}m</span>
              </div>
              <div style="color:#475569;font-size:9px;margin-top:2px;">${fmt(r.saldo_dias)} días</div>
            </td>
            <td style="padding:9px 8px;text-align:right;">
              <div style="color:#f1f5f9;font-weight:600;">${fmt(r.saldo_rollos)}</div>
              ${r.bajo_reorden ? `<div style="color:#FFC04D;font-size:9px;">⚠ bajo reorden</div>` : ''}
            </td>
            <td style="padding:9px 8px;text-align:right;color:#99D1FC;">${fmt(r.prom_mensual, 1)}</td>
            <td style="padding:9px 8px;text-align:right;color:#7B8CDE;">${fmt(r.prom_diario, 2)}</td>
            <td style="padding:9px 8px;text-align:right;color:#7B8CDE;">${fmt(r.prom_semanal, 1)}</td>
            <td style="padding:9px 8px;text-align:right;color:#94a3b8;">${fmt(r.rollos_proyect)}</td>
            <td style="padding:9px 8px;text-align:right;">
              <span style="color:${r.bajo_reorden ? '#FFC04D' : '#64748b'};">${fmt(r.punto_reorden)}</span>
            </td>
            <td style="padding:9px 8px;text-align:center;">
              <span style="color:${r.ind_variacion > 15 ? '#FF5C5C' : r.ind_variacion < -15 ? '#FFC04D' : '#94a3b8'};font-family:'JetBrains Mono',monospace;font-size:11px;">${r.variacion_label}</span>
            </td>
            <td style="padding:9px 8px;text-align:center;color:#7B8CDE;font-family:'JetBrains Mono',monospace;font-size:11px;">${fmt(r.ind_rotacion, 2)}x</td>
            <td style="padding:9px 8px;text-align:center;">
              ${r.sla_cumple
            ? `<span style="color:#B0F2AE;font-size:14px;" title="Cumple SLA ≥3 meses">✓</span>`
            : `<span style="color:#FF5C5C;font-size:14px;" title="No cumple SLA 3 meses">✗</span>`}
            </td>
            <td style="padding:9px 8px;text-align:center;">
              ${r.ind_riesgo
            ? `<span style="font-size:10px;padding:2px 8px;border-radius:10px;background:${r.ind_riesgo === 'CRÍTICO' || r.ind_riesgo === 'ALTO' ? 'rgba(255,92,92,.15)' : r.ind_riesgo === 'MEDIO' ? 'rgba(255,192,77,.12)' : 'rgba(176,242,174,.08)'};color:${r.ind_riesgo === 'CRÍTICO' || r.ind_riesgo === 'ALTO' ? '#FF5C5C' : r.ind_riesgo === 'MEDIO' ? '#FFC04D' : '#B0F2AE'};">${r.ind_riesgo}</span>`
            : `<span style="color:#475569;font-size:11px;">—</span>`}
            </td>
          </tr>`).join('')}
      </tbody>
    </table>`;

    // Pagination
    if (pagEl) {
        let phtml = '';
        const maxShow = 7;
        const half = Math.floor(maxShow / 2);
        let start_ = Math.max(1, riPage - half);
        let end_ = Math.min(totalPages, start_ + maxShow - 1);
        if (end_ - start_ < maxShow - 1) start_ = Math.max(1, end_ - maxShow + 1);

        if (riPage > 1) phtml += `<button onclick="riGoPage(${riPage - 1})" style="${_riPagBtnStyle(false)}">‹</button>`;
        if (start_ > 1) { phtml += `<button onclick="riGoPage(1)" style="${_riPagBtnStyle(false)}">1</button>`; if (start_ > 2) phtml += `<span style="color:#475569;padding:0 4px;">…</span>`; }
        for (let p = start_; p <= end_; p++) phtml += `<button onclick="riGoPage(${p})" style="${_riPagBtnStyle(p === riPage)}">${p}</button>`;
        if (end_ < totalPages) { if (end_ < totalPages - 1) phtml += `<span style="color:#475569;padding:0 4px;">…</span>`; phtml += `<button onclick="riGoPage(${totalPages})" style="${_riPagBtnStyle(false)}">${totalPages}</button>`; }
        if (riPage < totalPages) phtml += `<button onclick="riGoPage(${riPage + 1})" style="${_riPagBtnStyle(false)}">›</button>`;

        pagEl.innerHTML = phtml;
    }
}

function _riPagBtnStyle(active) {
    return `padding:5px 10px;border-radius:6px;border:1px solid ${active ? 'rgba(153,209,252,.5)' : 'rgba(255,255,255,.08)'};background:${active ? 'rgba(153,209,252,.12)' : 'rgba(255,255,255,.03)'};color:${active ? '#99D1FC' : '#94a3b8'};cursor:pointer;font-family:'JetBrains Mono',monospace;font-size:11px;`;
}

window.riGoPage = function (p) { riPage = p; riApplySearch(riSearchTerm); };

// ──────────────────────────────────────────────────────────────────
//  EXPORT EXCEL
// ──────────────────────────────────────────────────────────────────
window.exportRIExcel = function () {
    if (!window.XLSX) { alert('XLSX library not loaded'); return; }
    const rows = RI_FILTERED.map(r => ({
        'Código Sitio': r.codigo_sitio || r.codigo_mo,
        'Nombre Sitio': r.nombre_sitio,
        'Departamento': r.departamento,
        'Ciudad': r.ciudad,
        'Proyecto': r.proyecto,
        'Estado Punto': r.estado_punto,
        'Saldo Rollos': r.saldo_rollos,
        'Saldo Días': r.saldo_dias,
        'Cobertura (meses)': +r.saldo_meses.toFixed(2),
        'Consumo Mensual': r.prom_mensual,
        'Consumo Diario': +r.prom_diario.toFixed(2),
        'Consumo Semanal': +r.prom_semanal.toFixed(2),
        'Rollos Proyect.': r.rollos_proyect,
        'Punto Reorden': r.punto_reorden,
        'Período Abast. (días)': r.periodo_abast,
        'Variación Consumo %': r.ind_variacion,
        'Rotación Inventario': r.ind_rotacion,
        'Riesgo Quiebre': r.ind_riesgo,
        'Cumple SLA 3m': r.sla_cumple ? 'Sí' : 'No',
        'Bajo Punto Reorden': r.bajo_reorden ? 'Sí' : 'No',
        'Rollos Entregados': r.roll_entregados,
        'Rollos Consumidos': r.roll_consumidos,
        'Fecha Apertura': r.fecha_apertura,
        'Fecha Abast.': r.fecha_abst,
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Inventario Rollos');
    XLSX.writeFile(wb, `inventario_rollos_${new Date().toISOString().slice(0, 10)}.xlsx`);
};