/**
 * ╔══════════════════════════════════════════════════════════════════╗
 * ║  rollos_inventario.js — Módulo Inventario de Rollos Wompi       ║
 * ║  Fuente: wompi_tablero_rollos_calculos (via data_rollos.json.gz) ║
 * ║          + data_tablero_rollos.json.gz (tablero completo enriq.) ║
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
//  TABLERO ROLLOS COMPLETO — data_tablero_rollos.json.gz
//  Todas las filas de v_rollos_tablero_wompi + cal_* de calculos
// ──────────────────────────────────────────────────────────────────
window.TABLERO_ROLLOS_RAW = null;   // payload completo del JSON
window.TABLERO_ROLLOS_FILAS = [];   // array de filas listo para consultar

/**
 * Carga data_tablero_rollos.json.gz y lo guarda en memoria.
 * No modifica ningún render ni funcionalidad existente.
 * Expone window.TABLERO_ROLLOS_FILAS y window.TABLERO_ROLLOS_RAW.
 */
async function _loadTablRollos() {
    if (window.TABLERO_ROLLOS_RAW) return; // ya cargado
    try {
        const res = await fetch('data_tablero_rollos.json.gz?t=' + Date.now());
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

        const payload = JSON.parse(new TextDecoder().decode(merged));
        window.TABLERO_ROLLOS_RAW   = payload;
        window.TABLERO_ROLLOS_FILAS = payload.filas || [];

        console.log(
            '[TablRollos] Cargado:',
            window.TABLERO_ROLLOS_FILAS.length, 'filas |',
            'join_sitio:', payload.pct_join_sitio + '%',
            '| join_cal:', payload.pct_join_calculos + '%',
            '| cols/fila:', Object.keys(window.TABLERO_ROLLOS_FILAS[0] || {}).length
        );
    } catch (e) {
        // El archivo puede no existir aún en el servidor; silenciar para no
        // interrumpir el dashboard existente.
        console.warn('[TablRollos] No se pudo cargar data_tablero_rollos.json.gz:', e.message);
        window.TABLERO_ROLLOS_RAW   = { filas: [] };
        window.TABLERO_ROLLOS_FILAS = [];
    }
}

// Iniciar carga en paralelo tan pronto como se parsee este script
_loadTablRollos();

// Exponer para que otros módulos puedan re-intentar o esperar
window.loadTablRollos = _loadTablRollos;

// ──────────────────────────────────────────────────────────────────
//  BOOTSTRAP — called once ROLLOS_RAW is ready
// ──────────────────────────────────────────────────────────────────
window.initRollosInventario = function () {
    console.log('[Rollos-Inv] Iniciando módulo...');
    
    // Ahora que usamos 'var' en dashboard.js, debería estar disponible globalmente
    const raw = window.ROLLOS_RAW || (typeof ROLLOS_RAW !== 'undefined' ? ROLLOS_RAW : null);

    if (!raw) {
        console.warn('[Rollos-Inv] ROLLOS_RAW no disponible todavía.');
        return;
    }


    // The calculos data lives at raw.calculos (array)
    // Fallback to comercio if calculos is empty
    let calc = raw.calculos || [];
    if (!calc.length && raw.comercio) {
        console.log('[Rollos-Inv] Usando fallback: comercio');
        calc = raw.comercio;
    }
    console.log('[Rollos-Inv] Filas a procesar:', calc.length);

    // Merge with sitio metadata from detalle if needed
    const detBySitio = new Map();
    (ROLLOS_RAW.detalle || []).forEach(r => {
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
    if (!RI_DATA.length && ROLLOS_RAW) window.initRollosInventario();
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
//  TOTAL ROLLOS POR COMERCIO — desde INV_RAW (inventario Wompi)
//  Equivale a: tab Inventario Wompi → Categoría: Rollos → Tipo: Site
// ──────────────────────────────────────────────────────────────────
window.renderRollosInvComercio = function () {
    const el = document.getElementById('ri-inv-comercio-section');
    if (!el) return;

    const raw = window.INV_RAW;
    if (!raw || !raw.length) {
        el.innerHTML = `<div style="color:#64748b;font-size:12px;padding:12px 0;">⚠ Inventario Wompi no disponible aún. Vuelve a abrir este tab en unos segundos.</div>`;
        return;
    }

    // Función de categoría — usa la global de inventario.js si está disponible
    const getCat = window.invCategoria
        ? window.invCategoria
        : (nombre) => ((nombre || '').toUpperCase().includes('ROLLO') ? 'Rollos' : 'Otro');

    // Filtrar solo rollos en comercios (Site)
    const rollosEnComercio = raw.filter(r =>
        getCat(r['Nombre']) === 'Rollos' &&
        (r['Tipo de ubicación'] || '').trim() === 'Site'
    );

    if (!rollosEnComercio.length) {
        el.innerHTML = `
        <div style="margin-top:8px;padding:20px;background:rgba(255,255,255,.03);border:1px solid rgba(255,255,255,.07);border-radius:12px;color:#64748b;font-size:12px;">
          🏬 <strong style="color:#94a3b8;">Stock de Rollos por Comercio (Inventario)</strong> — Sin rollos en ubicaciones tipo Site en el inventario actual.
        </div>`;
        return;
    }

    // Agrupar por comercio (Nombre de la ubicación)
    const byComercio = new Map();
    rollosEnComercio.forEach(r => {
        const key  = (r['Nombre de la ubicación'] || 'Sin nombre').trim();
        const cod  = (r['Código de ubicación']    || '').trim();
        const qty  = parseInt(r['Cantidad'])       || 0;
        if (!byComercio.has(key)) byComercio.set(key, { nombre: key, cod, total: 0, refs: new Map() });
        const entry = byComercio.get(key);
        entry.total += qty;
        const refKey = (r['Nombre'] || 'Sin ref').trim();
        entry.refs.set(refKey, (entry.refs.get(refKey) || 0) + qty);
    });

    // Ordenar por total desc
    const sorted         = [...byComercio.values()].sort((a, b) => b.total - a.total);
    const grandTotal     = sorted.reduce((s, r) => s + r.total, 0);
    const promedio       = sorted.length ? Math.round(grandTotal / sorted.length) : 0;
    const fmt            = n => n.toLocaleString('es-CO');

    el.innerHTML = `
    <div style="margin-top:4px;padding:24px;background:rgba(176,242,174,.03);border:1px solid rgba(176,242,174,.1);border-radius:16px;">

      <!-- Encabezado -->
      <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:12px;margin-bottom:20px;">
        <div>
          <div style="font-family:'Syne',sans-serif;font-size:15px;font-weight:700;color:#B0F2AE;letter-spacing:.4px;">
            🏬 Stock de Rollos por Comercio — Inventario Wompi
          </div>
          <div style="font-size:11px;color:#475569;margin-top:3px;">
            Fuente: inventario.json · categoría <strong style="color:#B0F2AE">Rollos</strong> · ubicaciones tipo <strong style="color:#99D1FC">Site</strong>
          </div>
        </div>
        <button onclick="window.exportRollosInvComercioExcel()"
          style="background:rgba(176,242,174,.1);border:1px solid rgba(176,242,174,.3);border-radius:8px;color:#B0F2AE;font-family:'Outfit',sans-serif;font-size:12px;padding:7px 14px;cursor:pointer;">
          ⬇ Excel
        </button>
      </div>

      <!-- KPIs -->
      <div style="display:flex;gap:14px;flex-wrap:wrap;margin-bottom:20px;">
        <div style="background:rgba(176,242,174,.07);border:1px solid rgba(176,242,174,.2);border-radius:12px;padding:14px 20px;min-width:140px;">
          <div style="font-family:'JetBrains Mono',monospace;font-size:22px;font-weight:700;color:#B0F2AE;">${fmt(grandTotal)}</div>
          <div style="font-size:10px;color:#64748b;margin-top:4px;">Total rollos en comercios</div>
        </div>
        <div style="background:rgba(153,209,252,.07);border:1px solid rgba(153,209,252,.2);border-radius:12px;padding:14px 20px;min-width:140px;">
          <div style="font-family:'JetBrains Mono',monospace;font-size:22px;font-weight:700;color:#99D1FC;">${fmt(sorted.length)}</div>
          <div style="font-size:10px;color:#64748b;margin-top:4px;">Comercios con rollos</div>
        </div>
        <div style="background:rgba(244,157,110,.07);border:1px solid rgba(244,157,110,.2);border-radius:12px;padding:14px 20px;min-width:140px;">
          <div style="font-family:'JetBrains Mono',monospace;font-size:22px;font-weight:700;color:#F49D6E;">${fmt(promedio)}</div>
          <div style="font-size:10px;color:#64748b;margin-top:4px;">Promedio por comercio</div>
        </div>
        <div style="background:rgba(255,92,92,.07);border:1px solid rgba(255,92,92,.2);border-radius:12px;padding:14px 20px;min-width:140px;">
          <div style="font-family:'JetBrains Mono',monospace;font-size:22px;font-weight:700;color:#FF5C5C;">${fmt(sorted.filter(r => r.total === 0).length)}</div>
          <div style="font-size:10px;color:#64748b;margin-top:4px;">Comercios sin stock</div>
        </div>
      </div>

      <!-- Buscador -->
      <div style="margin-bottom:14px;">
        <input id="ri-inv-com-search" type="text" placeholder="🔍 Buscar comercio o código..."
          oninput="window._riInvComSearch(this.value)"
          style="background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.1);border-radius:8px;padding:8px 14px;color:#f1f5f9;font-family:'Outfit',sans-serif;font-size:13px;width:300px;outline:none;transition:border-color .2s;"
          onfocus="this.style.borderColor='rgba(176,242,174,.4)'"
          onblur="this.style.borderColor='rgba(255,255,255,.1)'">
        <span id="ri-inv-com-count" style="margin-left:12px;font-size:11px;color:#475569;font-family:'JetBrains Mono',monospace;">${sorted.length} comercios</span>
      </div>

      <!-- Tabla -->
      <div id="ri-inv-com-table-wrap" style="overflow-x:auto;border-radius:10px;border:1px solid rgba(255,255,255,.06);"></div>
    </div>`;

    // Guardar datos para búsqueda, paginación y export
    window._riInvComData    = sorted;
    window._riInvComTotal   = grandTotal;
    window._riInvComPage    = 1;
    window._riInvComCurrent = sorted; // dataset activo (filtrado o completo)

    window._riInvComSearch = function (term) {
        const t        = (term || '').toLowerCase();
        const filtered = t
            ? sorted.filter(r => r.nombre.toLowerCase().includes(t) || r.cod.toLowerCase().includes(t))
            : sorted;
        window._riInvComCurrent = filtered;
        window._riInvComPage    = 1;
        const countEl  = document.getElementById('ri-inv-com-count');
        if (countEl) countEl.textContent = filtered.length + ' comercios';
        _renderRiInvComTable(filtered, grandTotal, 1);
    };

    window.riInvComGoPage = function (p) {
        window._riInvComPage = p;
        _renderRiInvComTable(window._riInvComCurrent, window._riInvComTotal, p);
        // scroll suave al inicio de la tabla
        const wrap = document.getElementById('ri-inv-com-table-wrap');
        if (wrap) wrap.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    };

    _renderRiInvComTable(sorted, grandTotal, 1);
};

const RI_INV_COM_PAGE_SIZE = 100;

function _renderRiInvComTable(data, grandTotal, page) {
    const wrap = document.getElementById('ri-inv-com-table-wrap');
    if (!wrap) return;

    page = page || window._riInvComPage || 1;
    const totalPages = Math.max(1, Math.ceil(data.length / RI_INV_COM_PAGE_SIZE));
    if (page > totalPages) page = totalPages;
    window._riInvComPage = page;

    const start  = (page - 1) * RI_INV_COM_PAGE_SIZE;
    const slice  = data.slice(start, start + RI_INV_COM_PAGE_SIZE);

    const fmt    = n => n.toLocaleString('es-CO');
    const maxVal = data.length ? data[0].total : 1;

    if (!data.length) {
        wrap.innerHTML = `<div style="color:#64748b;padding:24px;text-align:center;font-size:13px;">Sin resultados para la búsqueda</div>`;
        return;
    }

    const barColor = qty =>
        qty >= 10 ? '#B0F2AE' : qty >= 5 ? '#DFFF61' : qty >= 2 ? '#FFC04D' : '#FF5C5C';

    const refsHtml = refs =>
        [...refs.entries()]
            .sort((a, b) => b[1] - a[1])
            .slice(0, 3)
            .map(([ref, qty]) => {
                const label = ref.length > 22 ? ref.substring(0, 22) + '…' : ref;
                return `<span title="${ref}" style="font-size:9px;background:rgba(176,242,174,.07);border:1px solid rgba(176,242,174,.16);border-radius:10px;padding:2px 7px;color:#94a3b8;white-space:nowrap;">${qty}u · ${label}</span>`;
            }).join(' ');

    // ── Paginación ─────────────────────────────────────────────────
    const pagBtnStyle = active =>
        `padding:5px 10px;border-radius:6px;border:1px solid ${active ? 'rgba(176,242,174,.5)' : 'rgba(255,255,255,.08)'};background:${active ? 'rgba(176,242,174,.12)' : 'rgba(255,255,255,.03)'};color:${active ? '#B0F2AE' : '#94a3b8'};cursor:pointer;font-family:'JetBrains Mono',monospace;font-size:11px;`;

    function buildPagHtml() {
        if (totalPages <= 1) return '';
        let h = `<div style="display:flex;align-items:center;gap:6px;padding:12px 14px;border-top:1px solid rgba(255,255,255,.06);flex-wrap:wrap;">`;
        h += `<span style="font-size:11px;color:#475569;font-family:'JetBrains Mono',monospace;margin-right:6px;">Pág ${page}/${totalPages} · ${data.length} registros</span>`;
        if (page > 1) h += `<button onclick="riInvComGoPage(${page-1})" style="${pagBtnStyle(false)}">‹</button>`;
        // rango visible
        const half = 3, st = Math.max(1, page-half), en = Math.min(totalPages, page+half);
        if (st > 1) { h += `<button onclick="riInvComGoPage(1)" style="${pagBtnStyle(false)}">1</button>`; if (st > 2) h += `<span style="color:#475569;padding:0 3px;">…</span>`; }
        for (let p = st; p <= en; p++) h += `<button onclick="riInvComGoPage(${p})" style="${pagBtnStyle(p===page)}">${p}</button>`;
        if (en < totalPages) { if (en < totalPages-1) h += `<span style="color:#475569;padding:0 3px;">…</span>`; h += `<button onclick="riInvComGoPage(${totalPages})" style="${pagBtnStyle(false)}">${totalPages}</button>`; }
        if (page < totalPages) h += `<button onclick="riInvComGoPage(${page+1})" style="${pagBtnStyle(false)}">›</button>`;
        h += `</div>`;
        return h;
    }

    const pagHtml = buildPagHtml();

    wrap.innerHTML = `
    ${pagHtml}
    <div style="overflow-x:auto;">
    <table style="width:100%;border-collapse:collapse;font-family:'Outfit',sans-serif;font-size:12px;">
      <thead>
        <tr style="background:rgba(0,0,0,.35);border-bottom:1px solid rgba(176,242,174,.15);">
          <th style="padding:10px 14px;text-align:left;color:#64748b;font-weight:600;font-size:10px;letter-spacing:.5px;">#</th>
          <th style="padding:10px 14px;text-align:left;color:#64748b;font-weight:600;font-size:10px;letter-spacing:.5px;">COMERCIO</th>
          <th style="padding:10px 8px;text-align:left;color:#64748b;font-weight:600;font-size:10px;letter-spacing:.5px;">CÓD. UBIC.</th>
          <th style="padding:10px 8px;text-align:right;color:#64748b;font-weight:600;font-size:10px;letter-spacing:.5px;">ROLLOS</th>
          <th style="padding:10px 8px;text-align:right;color:#64748b;font-weight:600;font-size:10px;letter-spacing:.5px;">% TOTAL</th>
          <th style="padding:10px 14px;text-align:left;color:#64748b;font-weight:600;font-size:10px;letter-spacing:.5px;min-width:140px;">DISTRIBUCIÓN</th>
          <th style="padding:10px 8px;text-align:left;color:#64748b;font-weight:600;font-size:10px;letter-spacing:.5px;">REFERENCIAS (top 3)</th>
        </tr>
      </thead>
      <tbody>
        ${slice.map((r, i) => {
            const globalIdx = start + i;
            const pct    = grandTotal > 0 ? (r.total / grandTotal * 100) : 0;
            const barW   = maxVal  > 0 ? (r.total / maxVal  * 100) : 0;
            const col    = barColor(r.total);
            const bgBase = i % 2 === 0 ? 'transparent' : 'rgba(255,255,255,.015)';
            return `
            <tr style="border-bottom:1px solid rgba(255,255,255,.04);background:${bgBase};transition:background .15s;"
                onmouseover="this.style.background='rgba(176,242,174,.05)'"
                onmouseout="this.style.background='${bgBase}'">
              <td style="padding:9px 14px;color:#334155;font-family:'JetBrains Mono',monospace;font-size:10px;">${globalIdx + 1}</td>
              <td style="padding:9px 14px;color:#f1f5f9;font-weight:600;">${r.nombre}</td>
              <td style="padding:9px 8px;color:#64748b;font-family:'JetBrains Mono',monospace;font-size:10px;">${r.cod || '—'}</td>
              <td style="padding:9px 8px;text-align:right;">
                <span style="font-family:'JetBrains Mono',monospace;font-size:14px;font-weight:700;color:${col};">${fmt(r.total)}</span>
              </td>
              <td style="padding:9px 8px;text-align:right;color:#94a3b8;font-family:'JetBrains Mono',monospace;font-size:11px;">${pct.toFixed(1)}%</td>
              <td style="padding:9px 14px;">
                <div style="background:rgba(255,255,255,.06);border-radius:4px;height:6px;overflow:hidden;min-width:100px;">
                  <div style="width:${barW.toFixed(1)}%;height:100%;background:${col};border-radius:4px;"></div>
                </div>
              </td>
              <td style="padding:9px 8px;">
                <div style="display:flex;gap:4px;flex-wrap:wrap;">${refsHtml(r.refs) || '<span style="color:#334155;font-size:10px;">—</span>'}</div>
              </td>
            </tr>`;
        }).join('')}
      </tbody>
      <tfoot>
        <tr style="background:rgba(176,242,174,.06);border-top:2px solid rgba(176,242,174,.2);">
          <td colspan="3" style="padding:10px 14px;font-family:'Syne',sans-serif;font-size:11px;font-weight:700;color:#B0F2AE;letter-spacing:.5px;">
            TOTAL — ${data.length} comercios
          </td>
          <td style="padding:10px 8px;text-align:right;font-family:'JetBrains Mono',monospace;font-size:14px;font-weight:700;color:#B0F2AE;">
            ${fmt(data.reduce((s,r) => s + r.total, 0))}
          </td>
          <td style="padding:10px 8px;text-align:right;font-family:'JetBrains Mono',monospace;font-size:11px;color:#475569;">100%</td>
          <td colspan="2"></td>
        </tr>
      </tfoot>
    </table>
    </div>
    ${pagHtml}`;
}

// ──────────────────────────────────────────────────────────────────
//  KPIs ROLLOS EN BODEGA — se inyectan en el panel "Rollos Wompi"
//  Misma estructura de cards que renderRollosKPIs pero con datos
//  de INV_RAW filtrados por categoría Rollos + ubicación Bodega
// ──────────────────────────────────────────────────────────────────
window.renderRollosInvBodegaKPIs = function () {
    const el = document.getElementById('rollos-inv-bodega-kpis');
    if (!el) return;

    const raw = window.INV_RAW;
    if (!raw || !raw.length) {
        el.innerHTML = `<div style="color:#64748b;font-size:12px;padding:8px 0;">⚠ Inventario no disponible aún.</div>`;
        return;
    }

    const INV_BODEGAS_SET = window.INV_BODEGAS instanceof Set && window.INV_BODEGAS.size > 0
        ? window.INV_BODEGAS
        : new Set([
            "ALMACEN WOMPI MEDELLIN","ALMACEN WOMPI BOGOTA","ALMACEN WOMPI BUCARAMANGA",
            "ALMACEN WOMPI CALI","ALMACEN WOMPI VILLAVICENCIO","ALMACEN WOMPI CUCUTA",
            "ALMACEN WOMPI PEREIRA","ALMACEN WOMPI NEIVA","ALMACEN WOMPI IBAGUE",
            "ALMACEN WOMPI TUNJA","ALMACEN WOMPI MONTERIA","ALMACEN WOMPI SANTA MARTA",
            "ALMACEN WOMPI VALLEDUPAR","ALMACEN WOMPI CARTAGENA","ALMACEN WOMPI FLORENCIA",
            "ALMACEN WOMPI POPAYAN","ALMACEN WOMPI MANIZALES","ALMACEN WOMPI YOPAL",
            "ALMACEN WOMPI APARTADO","ALMACEN WOMPI PASTO",
            "ALMACEN WOMPI SINCELEJO","ALMACEN WOMPI BARRANQUILLA","ALMACEN WOMPI ARMENIA",
            "ALMACEN BAJAS WOMPI","ALMACEN INGENICO - PROVEEDOR WOMPI",
          ]);
    const getCat = window.invCategoria || (n => ((n||'').toUpperCase().includes('ROLLO') ? 'Rollos' : 'Otro'));
    const sumQty = rows => rows.reduce((s, r) => s + (parseInt(r['Cantidad']) || 0), 0);
    const pct    = (n, d) => d > 0 ? Math.round(n / d * 100) : 0;
    const fmt    = n => n.toLocaleString('es-CO');

    // Solo rollos
    const soloRollos  = raw.filter(r => getCat(r['Nombre']) === 'Rollos');
    const totalRollos = sumQty(soloRollos);

    // Por tipo de ubicación
    const enBodega   = sumQty(soloRollos.filter(r => INV_BODEGAS_SET.size
        ? INV_BODEGAS_SET.has((r['Nombre de la ubicación']||'').trim())
        : (r['Tipo de ubicación']||'').trim() === 'Warehouse'));
    const enComercio = sumQty(soloRollos.filter(r => (r['Tipo de ubicación']||'').trim() === 'Site'));
    const enTecnico  = sumQty(soloRollos.filter(r => (r['Tipo de ubicación']||'').trim() === 'Staff'));
    const enOPL      = sumQty(soloRollos.filter(r => (r['Tipo de ubicación']||'').trim() === 'Supplier'));
    // Bodegas únicas con rollos
    const bodegasConRollos = new Set(
        soloRollos
            .filter(r => INV_BODEGAS_SET.size
                ? INV_BODEGAS_SET.has((r['Nombre de la ubicación']||'').trim())
                : (r['Tipo de ubicación']||'').trim() === 'Warehouse')
            .map(r => (r['Nombre de la ubicación']||'').trim())
            .filter(Boolean)
    );

    // Referencias únicas
    const refsUnicas = new Set(soloRollos.map(r => r['Nombre']).filter(Boolean));

    const cards = [
        {
            label: 'Total Rollos Inventario', value: fmt(totalRollos), icon: '📦',
            color: '#B0F2AE', bg: 'rgba(176,242,174,.07)', border: 'rgba(176,242,174,.2)',
            sub: `${refsUnicas.size} referencia${refsUnicas.size !== 1 ? 's' : ''} distintas`,
            drillRows: soloRollos, drillTitle: 'Total Rollos Inventario'
        },
        {
            label: 'En Bodega Wompi', value: fmt(enBodega), icon: '🏪',
            color: '#99D1FC', bg: 'rgba(153,209,252,.07)', border: 'rgba(153,209,252,.2)',
            sub: `${pct(enBodega, totalRollos)}% del total · ${bodegasConRollos.size} bodega${bodegasConRollos.size !== 1 ? 's' : ''}`,
            pct: pct(enBodega, totalRollos), pctColor: '#99D1FC',
            drillRows: soloRollos.filter(r => INV_BODEGAS_SET.size
                ? INV_BODEGAS_SET.has((r['Nombre de la ubicación']||'').trim())
                : (r['Tipo de ubicación']||'').trim() === 'Warehouse'),
            drillTitle: 'En Bodega Wompi'
        },
        {
            label: 'En Comercio (Site)', value: fmt(enComercio), icon: '🏬',
            color: '#DFFF61', bg: 'rgba(223,255,97,.06)', border: 'rgba(223,255,97,.18)',
            sub: `${pct(enComercio, totalRollos)}% del total`,
            pct: pct(enComercio, totalRollos), pctColor: '#DFFF61',
            drillRows: soloRollos.filter(r => (r['Tipo de ubicación']||'').trim() === 'Site'),
            drillTitle: 'En Comercio (Site)'
        },
        {
            label: 'Tec. Lineacom (Staff)', value: fmt(enTecnico), icon: '🔧',
            color: '#F49D6E', bg: 'rgba(244,157,110,.07)', border: 'rgba(244,157,110,.2)',
            sub: `${pct(enTecnico, totalRollos)}% del total`,
            pct: pct(enTecnico, totalRollos), pctColor: '#F49D6E',
            drillRows: soloRollos.filter(r => (r['Tipo de ubicación']||'').trim() === 'Staff'),
            drillTitle: 'Tec. Lineacom (Staff)'
        },
        {
            label: 'OPL (Supplier)', value: fmt(enOPL), icon: '🚚',
            color: '#C084FC', bg: 'rgba(192,132,252,.07)', border: 'rgba(192,132,252,.18)',
            sub: `${pct(enOPL, totalRollos)}% del total`,
            pct: pct(enOPL, totalRollos), pctColor: '#C084FC',
            drillRows: soloRollos.filter(r => (r['Tipo de ubicación']||'').trim() === 'Supplier'),
            drillTitle: 'OPL (Supplier)'
        },
    ];

    // Guardar cards para acceso desde el modal
    window._rollInvBodegaCards = cards;

    el.innerHTML = `
    <div style="padding-top:22px;border-top:1px solid rgba(176,242,174,.12);margin-bottom:6px;">
      <div style="font-family:'Syne',sans-serif;font-size:13px;font-weight:700;color:#B0F2AE;letter-spacing:.6px;text-transform:uppercase;margin-bottom:14px;">
        📦 Stock Actual de Rollos — Inventario Wompi
        <span style="font-family:'Outfit',sans-serif;font-size:10px;font-weight:400;color:#475569;text-transform:none;letter-spacing:0;margin-left:10px;">Haz clic en un KPI para ver el detalle</span>
      </div>
      <div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(190px,1fr));gap:14px;">
        ${cards.map((c, idx) => `
        <div data-kpi-idx="${idx}" style="background:${c.bg};border:1px solid ${c.border};border-radius:14px;padding:16px 18px;position:relative;overflow:hidden;transition:transform .2s,box-shadow .2s;cursor:pointer;"
             onclick="window._openRollInvBodegaModal(${idx})"
             onmouseover="this.style.transform='translateY(-2px)';this.style.boxShadow='0 8px 24px rgba(0,0,0,.4)';this.style.borderColor='${c.color}66'"
             onmouseout="this.style.transform='';this.style.boxShadow='';this.style.borderColor='${c.border}'">
          <div style="position:absolute;top:10px;right:10px;font-size:9px;color:${c.color};opacity:.6;font-family:'Outfit',sans-serif;">Ver detalle ›</div>
          <div style="font-size:20px;margin-bottom:6px;">${c.icon}</div>
          <div style="font-family:'JetBrains Mono',monospace;font-size:20px;font-weight:700;color:${c.color};line-height:1;">${c.value}</div>
          <div style="font-size:10px;color:#7A7674;margin-top:6px;font-family:'Outfit',sans-serif;">${c.label}</div>
          <div style="font-size:10px;color:#475569;margin-top:3px;">${c.sub}</div>
          ${c.pct !== undefined ? `
          <div style="margin-top:8px;background:rgba(255,255,255,.06);border-radius:4px;height:4px;overflow:hidden;">
            <div style="width:${Math.min(c.pct,100)}%;height:100%;background:${c.pctColor};border-radius:4px;transition:width .6s ease;"></div>
          </div>` : ''}
        </div>`).join('')}
      </div>
    </div>`;

    // ── Modal drilldown para KPIs de Rollos Inventario ───────────────
    if (!document.getElementById('roll-inv-modal')) {
        const modal = document.createElement('div');
        modal.id = 'roll-inv-modal';
        modal.style.cssText = 'display:none;position:fixed;inset:0;z-index:9999;background:rgba(0,0,0,.75);backdrop-filter:blur(4px);align-items:center;justify-content:center;padding:20px;';
        modal.innerHTML = `
        <div id="roll-inv-modal-box" style="background:#0d1520;border:1px solid rgba(176,242,174,.2);border-radius:20px;width:min(960px,100%);max-height:85vh;display:flex;flex-direction:column;overflow:hidden;box-shadow:0 32px 80px rgba(0,0,0,.7);">
          <div style="padding:20px 24px 16px;border-bottom:1px solid rgba(255,255,255,.07);display:flex;align-items:center;justify-content:space-between;flex-shrink:0;">
            <div>
              <div id="roll-inv-modal-title" style="font-family:'Syne',sans-serif;font-size:16px;font-weight:700;color:#f1f5f9;letter-spacing:.4px;"></div>
              <div id="roll-inv-modal-count" style="font-family:'JetBrains Mono',monospace;font-size:11px;color:#475569;margin-top:3px;"></div>
            </div>
            <div style="display:flex;align-items:center;gap:10px;">
              <input id="roll-inv-modal-search" type="text" placeholder="🔍 Buscar referencia, ubicación..."
                oninput="window._rollInvModalSearch(this.value)"
                style="background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.1);border-radius:8px;padding:7px 12px;color:#f1f5f9;font-family:'Outfit',sans-serif;font-size:12px;width:240px;outline:none;"
                onfocus="this.style.borderColor='rgba(176,242,174,.4)'"
                onblur="this.style.borderColor='rgba(255,255,255,.1)'">
              <button onclick="window._rollInvModalExcel()"
                style="background:rgba(176,242,174,.1);border:1px solid rgba(176,242,174,.25);border-radius:8px;color:#B0F2AE;font-family:'Outfit',sans-serif;font-size:12px;padding:7px 13px;cursor:pointer;">⬇ Excel</button>
              <button onclick="window._closeRollInvModal()"
                style="background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.1);border-radius:8px;color:#94a3b8;font-family:'Outfit',sans-serif;font-size:13px;padding:7px 13px;cursor:pointer;line-height:1;">✕</button>
            </div>
          </div>
          <div id="roll-inv-modal-body" style="overflow-y:auto;flex:1;padding:0;"></div>
          <div id="roll-inv-modal-pag" style="padding:12px 20px;border-top:1px solid rgba(255,255,255,.06);display:flex;align-items:center;gap:6px;flex-wrap:wrap;flex-shrink:0;"></div>
        </div>`;
        document.body.appendChild(modal);
        modal.addEventListener('click', e => { if (e.target === modal) window._closeRollInvModal(); });
    }

    const MODAL_PAGE_SIZE = 50;
    let _modalAllRows = [], _modalFilteredRows = [], _modalPage = 1, _modalColor = '#B0F2AE', _modalTitle = '';

    window._openRollInvBodegaModal = function(idx) {
        const card = (window._rollInvBodegaCards || [])[idx];
        if (!card || !card.drillRows) return;
        _modalAllRows = card.drillRows;
        _modalFilteredRows = _modalAllRows.slice();
        _modalPage = 1;
        _modalColor = card.color;
        _modalTitle = card.drillTitle;
        const modal = document.getElementById('roll-inv-modal');
        const titleEl = document.getElementById('roll-inv-modal-title');
        const searchEl = document.getElementById('roll-inv-modal-search');
        if (titleEl) titleEl.innerHTML = `${card.icon} ${card.label} <span style="color:${card.color}">${card.value}</span>`;
        if (searchEl) searchEl.value = '';
        modal.style.display = 'flex';
        document.body.style.overflow = 'hidden';
        _rollInvModalRender();
    };

    window._closeRollInvModal = function() {
        const modal = document.getElementById('roll-inv-modal');
        if (modal) modal.style.display = 'none';
        document.body.style.overflow = '';
    };

    window._rollInvModalSearch = function(term) {
        const t = (term || '').toLowerCase();
        _modalFilteredRows = t
            ? _modalAllRows.filter(r =>
                (r['Nombre']||'').toLowerCase().includes(t) ||
                (r['Nombre de la ubicación']||'').toLowerCase().includes(t) ||
                (r['Código de ubicación']||'').toLowerCase().includes(t) ||
                (r['Tipo de ubicación']||'').toLowerCase().includes(t))
            : _modalAllRows.slice();
        _modalPage = 1;
        _rollInvModalRender();
    };

    window._rollInvModalGoPage = function(p) {
        _modalPage = p;
        _rollInvModalRender();
        const body = document.getElementById('roll-inv-modal-body');
        if (body) body.scrollTop = 0;
    };

    window._rollInvModalExcel = function() {
        if (!window.XLSX) { alert('XLSX no disponible'); return; }
        const rows = _modalFilteredRows.map(r => ({
            'Referencia': r['Nombre'] || '',
            'Ubicación': r['Nombre de la ubicación'] || '',
            'Código Ubicación': r['Código de ubicación'] || '',
            'Tipo Ubicación': r['Tipo de ubicación'] || '',
            'Cantidad': parseInt(r['Cantidad']) || 0,
        }));
        const ws = window.XLSX.utils.json_to_sheet(rows);
        const wb = window.XLSX.utils.book_new();
        window.XLSX.utils.book_append_sheet(wb, ws, _modalTitle.substring(0, 31));
        window.XLSX.writeFile(wb, `rollos_${_modalTitle.replace(/\s+/g,'_').toLowerCase()}_${new Date().toISOString().slice(0,10)}.xlsx`);
    };

    function _rollInvModalRender() {
        const countEl = document.getElementById('roll-inv-modal-count');
        const body = document.getElementById('roll-inv-modal-body');
        const pagEl = document.getElementById('roll-inv-modal-pag');
        if (!body) return;

        const total = _modalFilteredRows.length;
        const totalPages = Math.max(1, Math.ceil(total / MODAL_PAGE_SIZE));
        if (_modalPage > totalPages) _modalPage = 1;
        const start = (_modalPage - 1) * MODAL_PAGE_SIZE;
        const slice = _modalFilteredRows.slice(start, start + MODAL_PAGE_SIZE);

        const totalQty = _modalFilteredRows.reduce((s, r) => s + (parseInt(r['Cantidad']) || 0), 0);
        if (countEl) countEl.textContent = `${total.toLocaleString('es-CO')} registros · ${totalQty.toLocaleString('es-CO')} unidades`;

        const fmt = n => Number(n).toLocaleString('es-CO');
        const tipoColor = t => t === 'Site' ? '#DFFF61' : t === 'Staff' ? '#F49D6E' : t === 'Supplier' ? '#C084FC' : '#99D1FC';

        body.innerHTML = total === 0
            ? `<div style="padding:40px;text-align:center;color:#475569;font-family:'Outfit',sans-serif;">Sin resultados para la búsqueda</div>`
            : `<div style="overflow-x:auto;">
            <table style="width:100%;border-collapse:collapse;font-family:'Outfit',sans-serif;font-size:12px;">
              <thead>
                <tr style="background:rgba(0,0,0,.4);border-bottom:1px solid rgba(176,242,174,.15);position:sticky;top:0;z-index:1;">
                  <th style="padding:10px 16px;text-align:left;color:#64748b;font-weight:600;font-size:10px;letter-spacing:.5px;">#</th>
                  <th style="padding:10px 16px;text-align:left;color:#64748b;font-weight:600;font-size:10px;letter-spacing:.5px;">REFERENCIA</th>
                  <th style="padding:10px 12px;text-align:left;color:#64748b;font-weight:600;font-size:10px;letter-spacing:.5px;">UBICACIÓN</th>
                  <th style="padding:10px 12px;text-align:left;color:#64748b;font-weight:600;font-size:10px;letter-spacing:.5px;">CÓD. UBIC.</th>
                  <th style="padding:10px 12px;text-align:center;color:#64748b;font-weight:600;font-size:10px;letter-spacing:.5px;">TIPO</th>
                  <th style="padding:10px 16px;text-align:right;color:#64748b;font-weight:600;font-size:10px;letter-spacing:.5px;">CANTIDAD</th>
                </tr>
              </thead>
              <tbody>
                ${slice.map((r, i) => {
                    const qty = parseInt(r['Cantidad']) || 0;
                    const tipo = (r['Tipo de ubicación'] || '').trim();
                    const bgBase = i % 2 === 0 ? 'transparent' : 'rgba(255,255,255,.015)';
                    return `<tr style="border-bottom:1px solid rgba(255,255,255,.04);background:${bgBase};transition:background .12s;"
                         onmouseover="this.style.background='rgba(176,242,174,.04)'"
                         onmouseout="this.style.background='${bgBase}'">
                      <td style="padding:8px 16px;color:#334155;font-family:'JetBrains Mono',monospace;font-size:10px;">${start + i + 1}</td>
                      <td style="padding:8px 16px;color:#f1f5f9;font-weight:600;max-width:220px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;" title="${(r['Nombre']||'').replace(/"/g,'&quot;')}">${r['Nombre'] || '—'}</td>
                      <td style="padding:8px 12px;color:#94a3b8;max-width:200px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;" title="${(r['Nombre de la ubicación']||'').replace(/"/g,'&quot;')}">${r['Nombre de la ubicación'] || '—'}</td>
                      <td style="padding:8px 12px;color:#64748b;font-family:'JetBrains Mono',monospace;font-size:10px;">${r['Código de ubicación'] || '—'}</td>
                      <td style="padding:8px 12px;text-align:center;">
                        <span style="font-size:10px;padding:2px 9px;border-radius:10px;background:${tipoColor(tipo)}18;color:${tipoColor(tipo)};font-weight:600;">${tipo || '—'}</span>
                      </td>
                      <td style="padding:8px 16px;text-align:right;font-family:'JetBrains Mono',monospace;font-size:13px;font-weight:700;color:${_modalColor};">${fmt(qty)}</td>
                    </tr>`;
                }).join('')}
              </tbody>
              <tfoot>
                <tr style="background:rgba(176,242,174,.05);border-top:2px solid rgba(176,242,174,.2);">
                  <td colspan="5" style="padding:10px 16px;font-family:'Syne',sans-serif;font-size:11px;font-weight:700;color:#B0F2AE;letter-spacing:.4px;">
                    SUBTOTAL PÁGINA ${_modalPage}/${totalPages}
                  </td>
                  <td style="padding:10px 16px;text-align:right;font-family:'JetBrains Mono',monospace;font-size:13px;font-weight:700;color:#B0F2AE;">
                    ${fmt(slice.reduce((s,r)=>s+(parseInt(r['Cantidad'])||0),0))}
                  </td>
                </tr>
              </tfoot>
            </table>
            </div>`;

        // Paginación
        if (pagEl) {
            const btnStyle = act => `padding:5px 10px;border-radius:6px;border:1px solid ${act?`rgba(176,242,174,.5)`:'rgba(255,255,255,.08)'};background:${act?'rgba(176,242,174,.12)':'rgba(255,255,255,.03)'};color:${act?'#B0F2AE':'#94a3b8'};cursor:pointer;font-family:'JetBrains Mono',monospace;font-size:11px;`;
            let h = `<span style="font-size:11px;color:#475569;font-family:'JetBrains Mono',monospace;margin-right:8px;">Pág ${_modalPage}/${totalPages}</span>`;
            if (_modalPage > 1) h += `<button onclick="window._rollInvModalGoPage(${_modalPage-1})" style="${btnStyle(false)}">‹</button>`;
            const half=3, st=Math.max(1,_modalPage-half), en=Math.min(totalPages,_modalPage+half);
            if (st>1){h+=`<button onclick="window._rollInvModalGoPage(1)" style="${btnStyle(false)}">1</button>`;if(st>2)h+=`<span style="color:#475569;padding:0 3px;">…</span>`;}
            for(let p=st;p<=en;p++) h+=`<button onclick="window._rollInvModalGoPage(${p})" style="${btnStyle(p===_modalPage)}">${p}</button>`;
            if(en<totalPages){if(en<totalPages-1)h+=`<span style="color:#475569;padding:0 3px;">…</span>`;h+=`<button onclick="window._rollInvModalGoPage(${totalPages})" style="${btnStyle(false)}">${totalPages}</button>`;}
            if(_modalPage<totalPages) h+=`<button onclick="window._rollInvModalGoPage(${_modalPage+1})" style="${btnStyle(false)}">›</button>`;
            pagEl.innerHTML = h;
        }
    }
};

// Export Excel de la sección de inventario por comercio
window.exportRollosInvComercioExcel = function () {
    if (!window.XLSX) { alert('XLSX library not loaded'); return; }
    const data = window._riInvComData || [];
    if (!data.length) { alert('Sin datos para exportar'); return; }
    const rows = data.map((r, i) => ({
        '#': i + 1,
        'Comercio (Nombre Ubicación)': r.nombre,
        'Código Ubicación': r.cod,
        'Total Rollos': r.total,
        '% del Total': window._riInvComTotal > 0 ? +(r.total / window._riInvComTotal * 100).toFixed(2) : 0,
        'Referencias (detalle)': [...r.refs.entries()].map(([ref, qty]) => `${qty}u: ${ref}`).join(' | '),
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Rollos por Comercio');
    XLSX.writeFile(wb, `rollos_por_comercio_inv_${new Date().toISOString().slice(0, 10)}.xlsx`);
};

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