/**
 * ╔══════════════════════════════════════════════════════════════════╗
 * ║  rollos_comercio_v2.js — Detalle por Comercio (nueva versión)   ║
 * ║                                                                  ║
 * ║  Reemplaza el panel rollos-comercio con:                         ║
 * ║  1. Buscador de comercio → stock (inventario) + historial MOs   ║
 * ║  2. Cobertura en meses por corresponsal                          ║
 * ║  3. Consumo real diario / semanal / mensual                      ║
 * ║  4. Consumo real vs proyectado + variación                       ║
 * ║  5. Inventario por bodegas y por corresponsal                    ║
 * ║  6. Rotación de inventario                                       ║
 * ║  7. SLA 3 meses + alertas + riesgo quiebre                       ║
 * ║  8. Punto de reorden + alertas automáticas                       ║
 * ║  9. Registro de quiebres de stock                                ║
 * ╚══════════════════════════════════════════════════════════════════╝
 */

'use strict';

// ── Estado del módulo ──────────────────────────────────────────────
let RCV2_READY   = false;
let RCV2_SITIOS  = new Map(); // codigo_sitio → { meta, calculos, movimientos[] }
let RCV2_CURRENT = null;      // sitio seleccionado actualmente

// ── Helpers ────────────────────────────────────────────────────────
const rcv2Fmt  = n  => (parseFloat(n) || 0).toLocaleString('es-CO', { maximumFractionDigits: 1 });
const rcv2FmtI = n  => Math.round(parseFloat(n) || 0).toLocaleString('es-CO');
const rcv2P    = n  => parseFloat(n) || 0;
const rcv2Date = s  => {
  if (!s) return null;
  const d = new Date(s);
  return isNaN(d) ? null : d;
};
const rcv2FmtDate = s => {
  const d = rcv2Date(s);
  if (!d) return '—';
  return d.toLocaleDateString('es-CO', { day:'2-digit', month:'short', year:'numeric' });
};

// ── Semáforo cobertura ─────────────────────────────────────────────
function rcv2Semaforo(meses) {
  if (meses >= 3)  return { col:'#B0F2AE', label:'OK',       emoji:'🟢' };
  if (meses >= 2)  return { col:'#DFFF61', label:'ATENCIÓN',  emoji:'🟡' };
  if (meses >= 1)  return { col:'#FFC04D', label:'ALERTA',    emoji:'🟠' };
  return                  { col:'#FF5C5C', label:'CRÍTICO',   emoji:'🔴' };
}

// ── Construir índice de sitios desde TABLERO_ROLLOS_FILAS ──────────
function rcv2BuildIndex() {
  const filas = window.TABLERO_ROLLOS_FILAS || [];
  RCV2_SITIOS.clear();

  filas.forEach(f => {
    const codSitio = (f.codigo_sitio || f.cal_codigo_sitio || '').trim();
    const codMO    = (f.cal_codigo_mo || '').trim();
    const key      = codSitio || codMO;
    if (!key) return;

    if (!RCV2_SITIOS.has(key)) {
      // Extraer calculos (tomados de la primera fila con datos cal_)
      const cal = {
        codigo_sitio    : key,
        codigo_mo       : codMO,
        estado_punto    : f.cal_estado_punto    || '',
        prom_mensual    : rcv2P(f.cal_promedio_mensual),
        rollos_prom_mes : rcv2P(f.cal_rollos_promedio_mes),
        periodo_abast   : rcv2P(f.cal_periodo_abast_e5),
        rollos_periodo  : rcv2P(f.cal_rollos_periodo_abast_e5),
        punto_reorden   : rcv2P(f.cal_punto_reorden),
        saldo_rollos    : rcv2P(f.cal_saldo_rollos),
        saldo_dias      : rcv2P(f.cal_saldo_dias),
        saldo_valor     : rcv2P(f.cal_saldo),
        rollos_entregados : rcv2P(f.cal_rollos_entregados_mig_apert),
        rollos_consumidos : rcv2P(f.cal_rollos_consumidos_migr_apert),
        trx_desde       : rcv2P(f.cal_trx_desde_migra_apert),
        fecha_apertura  : f.cal_fecha_apertura_final || '',
        fecha_abst      : f.cal_fecha_abst_1 || '',
        rollos_anio     : rcv2P(f.cal_rollos_anio_e5),
        join_nivel      : f.join_nivel || 0,
      };
      const meta = {
        nombre_sitio : f.nombre_sitio || f.nombre_ubicacion_destino || key,
        departamento : f.departamento || '',
        ciudad       : f.Ciudad || f.ciudad || '',
        nit          : f.nit || '',
        tipologia    : f.tipologia || '',
        proyecto     : f.proyecto || '',
        red          : f.red_asociada || '',
      };
      RCV2_SITIOS.set(key, { cal, meta, movs: [] });
    }

    // Agregar movimiento (cada fila del tablero es un movimiento de MO)
    const sitio = RCV2_SITIOS.get(key);
    if (f.tarea) {
      sitio.movs.push({
        tarea          : f.tarea,
        fecha          : f.fecha_confirmacion || f.tarea_fecha_fin || '',
        fecha_fin      : f.tarea_fecha_fin || '',
        plan_inicio    : f.plan_inicio || '',
        plan_fin       : f.plan_fin || '',
        estado         : f.estado_tarea || '',
        flujo          : f.flujo || '',
        cantidad       : rcv2P(f.Cantidad),
        material       : f.nombre_material || '',
        codigo_material: f.codigo_material || '',
        guia           : f.guia || f.guia_raw || '',
        transportadora : f.transportadora || '',
        origen         : f.nombre_ubicacion_origen || '',
        destino        : f.nombre_ubicacion_destino || '',
        plantilla      : f.nombre_plantilla_tarea || '',
        subproyecto    : f.subproyecto || '',
        cod_op         : f.codigo_operacion || '',
      });
    }
  });

  // Agregar stock de inventario (stock_wompi_filtrado) si está disponible
  const invRaw = window.INV_RAW || [];
  if (invRaw.length) {
    invRaw.forEach(r => {
      if ((window.invCategoria ? window.invCategoria(r['Nombre']) : '') !== 'Rollos') return;
      const codUbic = (r['Código de ubicación'] || '').trim();
      const nomUbic = (r['Nombre de la ubicación'] || '').trim();
      const qty     = parseInt(r['Cantidad']) || 0;
      // buscar por código o nombre
      let target = RCV2_SITIOS.get(codUbic);
      if (!target) {
        // intentar match por nombre normalizado
        for (const [k, v] of RCV2_SITIOS) {
          if (v.meta.nombre_sitio.toUpperCase() === nomUbic.toUpperCase()) {
            target = v; break;
          }
        }
      }
      if (target) {
        target.inv_stock = (target.inv_stock || 0) + qty;
        if (!target.inv_refs) target.inv_refs = new Map();
        const ref = (r['Nombre'] || 'Sin ref').trim();
        target.inv_refs.set(ref, (target.inv_refs.get(ref) || 0) + qty);
        target.inv_ubicacion = target.inv_ubicacion || nomUbic;
      }
    });
  }

  RCV2_READY = true;
  console.log('[RCV2] Índice construido:', RCV2_SITIOS.size, 'sitios');
}

// ── Autocompletar búsqueda ─────────────────────────────────────────
function rcv2Suggest(term) {
  const t  = term.trim().toUpperCase();
  const ul = document.getElementById('rcv2-suggestions');
  if (!ul) return;
  if (!t || t.length < 2) { ul.style.display = 'none'; return; }

  const matches = [];
  for (const [key, s] of RCV2_SITIOS) {
    const n = s.meta.nombre_sitio.toUpperCase();
    const c = s.meta.ciudad.toUpperCase();
    if (key.includes(t) || n.includes(t) || c.includes(t)) {
      matches.push({ key, s });
      if (matches.length >= 12) break;
    }
  }

  if (!matches.length) { ul.style.display = 'none'; return; }

  const sem = m => rcv2Semaforo(m.s.cal.saldo_dias / 30);
  ul.innerHTML = matches.map(m => {
    const sg = sem(m);
    const co = (m.s.cal.saldo_dias / 30).toFixed(1);
    return `<li onclick="rcv2Select('${m.key}')" style="padding:10px 16px;cursor:pointer;border-bottom:1px solid rgba(255,255,255,.05);display:flex;align-items:center;justify-content:space-between;gap:12px;transition:background .15s;" onmouseover="this.style.background='rgba(176,242,174,.08)'" onmouseout="this.style.background='transparent'">
      <div>
        <div style="font-size:12px;font-weight:600;color:#f1f5f9;">${m.s.meta.nombre_sitio}</div>
        <div style="font-size:10px;color:#64748b;margin-top:2px;">${m.key} · ${m.s.meta.ciudad}</div>
      </div>
      <div style="font-size:11px;color:${sg.col};font-family:'JetBrains Mono',monospace;white-space:nowrap;">${sg.emoji} ${co}m</div>
    </li>`;
  }).join('');
  ul.style.display = 'block';
}

// ── Seleccionar un comercio ────────────────────────────────────────
window.rcv2Select = function(key) {
  const ul = document.getElementById('rcv2-suggestions');
  if (ul) ul.style.display = 'none';
  const inp = document.getElementById('rcv2-search-input');
  const sitio = RCV2_SITIOS.get(key);
  if (!sitio) return;
  if (inp) inp.value = sitio.meta.nombre_sitio;
  RCV2_CURRENT = { key, ...sitio };
  rcv2Render(key, sitio);
};

// ── Render principal del comercio seleccionado ─────────────────────
function rcv2Render(key, sitio) {
  const el = document.getElementById('rcv2-detail-panel');
  if (!el) return;

  const { cal, meta, movs } = sitio;
  const saldoMeses   = cal.saldo_dias / 30;
  const sg           = rcv2Semaforo(saldoMeses);
  const promDia      = cal.prom_mensual / 30;
  const promSem      = promDia * 7;
  const slaOk        = saldoMeses >= 3;
  const bajoPR       = cal.saldo_rollos > 0 && cal.saldo_rollos <= cal.punto_reorden;
  const invStock     = sitio.inv_stock || 0;
  const invRefs      = sitio.inv_refs ? [...sitio.inv_refs.entries()].sort((a,b)=>b[1]-a[1]) : [];

  // Calcular consumo real desde movimientos (filas con cantidad > 0 y estado entregado)
  const movsOrd = [...movs].sort((a,b)=> {
    const da = rcv2Date(a.fecha) || new Date(0);
    const db = rcv2Date(b.fecha) || new Date(0);
    return db - da;
  });

  // Consumo por mes desde historial
  const consumoPorMes = new Map();
  movsOrd.forEach(m => {
    const d = rcv2Date(m.fecha);
    if (!d || !m.cantidad) return;
    const k = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;
    consumoPorMes.set(k, (consumoPorMes.get(k) || 0) + m.cantidad);
  });
  const mesesOrdenados = [...consumoPorMes.entries()].sort((a,b)=>a[0].localeCompare(b[0]));
  const ultMeses = mesesOrdenados.slice(-6);

  // Consumo últimos 30 días
  const hoy      = new Date();
  const hace30   = new Date(hoy - 30*864e5);
  const hace7    = new Date(hoy - 7*864e5);
  const consReal30 = movsOrd.filter(m=>rcv2Date(m.fecha)>=hace30).reduce((s,m)=>s+m.cantidad,0);
  const consReal7  = movsOrd.filter(m=>rcv2Date(m.fecha)>=hace7).reduce((s,m)=>s+m.cantidad,0);

  // Variación consumo real vs proyectado
  const proyMensual = cal.prom_mensual;
  const variacion   = proyMensual > 0 ? ((consReal30 - proyMensual) / proyMensual * 100) : 0;
  const varLabel    = variacion > 0 ? `+${variacion.toFixed(1)}%` : `${variacion.toFixed(1)}%`;
  const varCol      = Math.abs(variacion) < 10 ? '#B0F2AE' : Math.abs(variacion) < 25 ? '#FFC04D' : '#FF5C5C';

  // Rotación de inventario
  const rotacion = cal.saldo_rollos > 0 && cal.prom_mensual > 0
    ? (cal.rollos_consumidos / cal.saldo_rollos).toFixed(2)
    : '—';

  // Quiebres (movimientos con stock vacío o estado que indica falta)
  const quiebres = movsOrd.filter(m =>
    m.estado && (m.estado.toLowerCase().includes('quiebre') ||
    m.estado.toLowerCase().includes('falta') ||
    m.estado.toLowerCase().includes('sin stock') ||
    m.estado.toLowerCase().includes('devuelto'))
  );

  // ── Construir HTML ─────────────────────────────────────────────────
  const secStyle = 'background:rgba(255,255,255,.025);border:1px solid rgba(255,255,255,.07);border-radius:14px;padding:20px 24px;margin-bottom:16px;';
  const kpiStyle = (col) => `background:rgba(0,0,0,.25);border:1px solid ${col}33;border-radius:12px;padding:14px 18px;min-width:130px;flex:1;`;
  const chip     = (txt, col) => `<span style="font-size:10px;background:${col}18;border:1px solid ${col}44;border-radius:20px;padding:3px 10px;color:${col};font-family:'JetBrains Mono',monospace;">${txt}</span>`;

  // Alerta riesgo quiebre
  let alertaHtml = '';
  if (cal.saldo_rollos === 0) {
    alertaHtml = `<div style="background:#FF5C5C18;border:1px solid #FF5C5C66;border-radius:10px;padding:12px 16px;margin-bottom:16px;display:flex;align-items:center;gap:10px;"><span style="font-size:18px;">🚨</span><div><div style="font-size:13px;font-weight:700;color:#FF5C5C;">QUIEBRE DE STOCK</div><div style="font-size:11px;color:#94a3b8;margin-top:2px;">Este corresponsal no tiene rollos disponibles actualmente.</div></div></div>`;
  } else if (bajoPR) {
    alertaHtml = `<div style="background:#FFC04D18;border:1px solid #FFC04D66;border-radius:10px;padding:12px 16px;margin-bottom:16px;display:flex;align-items:center;gap:10px;"><span style="font-size:18px;">⚠️</span><div><div style="font-size:13px;font-weight:700;color:#FFC04D;">POR DEBAJO DEL PUNTO DE REORDEN</div><div style="font-size:11px;color:#94a3b8;margin-top:2px;">Saldo (${rcv2FmtI(cal.saldo_rollos)} rollos) ≤ Punto de reorden (${rcv2FmtI(cal.punto_reorden)} rollos). Generar orden urgente.</div></div></div>`;
  } else if (!slaOk) {
    alertaHtml = `<div style="background:#DFFF6118;border:1px solid #DFFF6166;border-radius:10px;padding:12px 16px;margin-bottom:16px;display:flex;align-items:center;gap:10px;"><span style="font-size:18px;">📋</span><div><div style="font-size:13px;font-weight:700;color:#DFFF61;">SLA INCUMPLIDO — Cobertura &lt; 3 meses</div><div style="font-size:11px;color:#94a3b8;margin-top:2px;">Cobertura actual: ${saldoMeses.toFixed(1)} meses. Se requieren ${((3-saldoMeses)*cal.prom_mensual/30).toFixed(0)} rollos adicionales para cumplir el SLA.</div></div></div>`;
  }

  // Barras de consumo por mes
  const maxCons = Math.max(...ultMeses.map(m=>m[1]), 1);
  const barsHtml = ultMeses.length ? ultMeses.map(([mes, qty]) => {
    const pct   = (qty / maxCons * 100).toFixed(0);
    const label = mes.split('-').reverse().join('/');
    return `<div style="flex:1;min-width:60px;text-align:center;">
      <div style="font-family:'JetBrains Mono',monospace;font-size:10px;color:#94a3b8;margin-bottom:4px;">${rcv2FmtI(qty)}</div>
      <div style="height:80px;display:flex;align-items:flex-end;justify-content:center;">
        <div style="width:28px;background:linear-gradient(180deg,#B0F2AE,#00825A);border-radius:4px 4px 0 0;height:${pct}%;min-height:4px;transition:height .3s;"></div>
      </div>
      <div style="font-size:9px;color:#64748b;margin-top:4px;">${label}</div>
    </div>`;
  }).join('') : '<div style="color:#475569;font-size:12px;padding:20px;">Sin historial de consumo disponible</div>';

  // Tabla de movimientos/MOs
  const MAX_MOVS = 50;
  const movsHtml = movsOrd.slice(0, MAX_MOVS).map((m, i) => {
    const est     = m.estado || '—';
    const eColor  = est.toLowerCase().includes('entregado') || est.toLowerCase().includes('completado') ? '#B0F2AE'
                  : est.toLowerCase().includes('tránsito') || est.toLowerCase().includes('transito') ? '#99D1FC'
                  : est.toLowerCase().includes('cancelado') || est.toLowerCase().includes('devuelto') ? '#FF5C5C'
                  : '#94a3b8';
    return `<tr style="border-bottom:1px solid rgba(255,255,255,.04);${i%2?'background:rgba(255,255,255,.012)':''}">
      <td style="padding:8px 12px;font-family:'JetBrains Mono',monospace;font-size:10px;color:#64748b;">${m.tarea||'—'}</td>
      <td style="padding:8px 10px;font-size:11px;color:#cbd5e1;">${rcv2FmtDate(m.fecha)}</td>
      <td style="padding:8px 10px;font-size:11px;color:#e2e8f0;">${m.material||'—'}</td>
      <td style="padding:8px 10px;font-family:'JetBrains Mono',monospace;font-size:12px;color:#B0F2AE;text-align:right;">${m.cantidad||0}</td>
      <td style="padding:8px 10px;"><span style="font-size:10px;color:${eColor};background:${eColor}18;border:1px solid ${eColor}33;border-radius:20px;padding:2px 8px;white-space:nowrap;">${est}</span></td>
      <td style="padding:8px 10px;font-size:10px;color:#475569;">${m.guia||'—'}</td>
      <td style="padding:8px 10px;font-size:10px;color:#475569;">${m.transportadora||'—'}</td>
    </tr>`;
  }).join('');

  // Referencias de inventario
  const refsHtml = invRefs.length
    ? invRefs.map(([ref, qty]) => `<div style="display:flex;justify-content:space-between;align-items:center;padding:6px 0;border-bottom:1px solid rgba(255,255,255,.04);">
        <span style="font-size:11px;color:#cbd5e1;">${ref}</span>
        <span style="font-family:'JetBrains Mono',monospace;font-size:12px;color:#B0F2AE;">${rcv2FmtI(qty)}</span>
      </div>`).join('')
    : '<div style="color:#475569;font-size:11px;padding:8px 0;">Sin datos de inventario por referencia</div>';

  // Quiebres registrados
  const quiebresHtml = quiebres.length
    ? quiebres.map(m => `<div style="display:flex;align-items:center;gap:10px;padding:8px 0;border-bottom:1px solid rgba(255,92,92,.08);">
        <span style="font-size:14px;">🚨</span>
        <div><div style="font-size:11px;color:#FF5C5C;">${m.estado}</div><div style="font-size:10px;color:#475569;">${m.tarea} · ${rcv2FmtDate(m.fecha)}</div></div>
      </div>`).join('')
    : '<div style="color:#475569;font-size:11px;padding:8px 0;">Sin quiebres registrados</div>';

  el.innerHTML = `
  <!-- Header comercio -->
  <div style="${secStyle}display:flex;align-items:flex-start;justify-content:space-between;flex-wrap:wrap;gap:16px;">
    <div>
      <div style="font-family:'Syne',sans-serif;font-size:18px;font-weight:800;color:#f1f5f9;letter-spacing:.3px;">${meta.nombre_sitio}</div>
      <div style="font-size:11px;color:#64748b;margin-top:4px;display:flex;gap:10px;flex-wrap:wrap;">
        <span>📍 ${meta.ciudad}${meta.departamento ? ', '+meta.departamento : ''}</span>
        ${meta.nit ? `<span>NIT: ${meta.nit}</span>` : ''}
        ${meta.tipologia ? `<span>Tipo: ${meta.tipologia}</span>` : ''}
        <span>Sitio: <code style="font-family:'JetBrains Mono',monospace;color:#99D1FC;">${key}</code></span>
      </div>
      <div style="margin-top:10px;display:flex;gap:8px;flex-wrap:wrap;">
        ${chip(sg.emoji + ' Cobertura: ' + saldoMeses.toFixed(1) + ' meses', sg.col)}
        ${chip(slaOk ? '✓ SLA Cumplido' : '✗ SLA Incumplido', slaOk ? '#B0F2AE' : '#FF5C5C')}
        ${bajoPR ? chip('⬇ Bajo Punto Reorden', '#FFC04D') : ''}
        ${cal.estado_punto ? chip(cal.estado_punto, '#99D1FC') : ''}
        ${chip('Unión N'+cal.join_nivel, '#475569')}
      </div>
    </div>
    <div style="text-align:right;">
      <div style="font-family:'JetBrains Mono',monospace;font-size:36px;font-weight:700;color:${sg.col};">${rcv2FmtI(cal.saldo_rollos)}</div>
      <div style="font-size:10px;color:#64748b;">rollos (cálculos)</div>
      ${invStock ? `<div style="font-family:'JetBrains Mono',monospace;font-size:22px;font-weight:700;color:#99D1FC;margin-top:4px;">${rcv2FmtI(invStock)}</div><div style="font-size:10px;color:#64748b;">rollos (inventario Wompi)</div>` : ''}
    </div>
  </div>

  ${alertaHtml}

  <!-- KPIs inventario & consumo -->
  <div style="display:flex;flex-wrap:wrap;gap:12px;margin-bottom:16px;">
    <div style="${kpiStyle(sg.col)}">
      <div style="font-size:9px;color:#64748b;letter-spacing:.5px;margin-bottom:6px;">COBERTURA</div>
      <div style="font-family:'JetBrains Mono',monospace;font-size:24px;font-weight:700;color:${sg.col};">${saldoMeses.toFixed(1)}<span style="font-size:12px;font-weight:400;"> m</span></div>
      <div style="font-size:10px;color:#475569;margin-top:2px;">${rcv2FmtI(cal.saldo_dias)} días</div>
    </div>
    <div style="${kpiStyle('#99D1FC')}">
      <div style="font-size:9px;color:#64748b;letter-spacing:.5px;margin-bottom:6px;">CONSUMO MENSUAL</div>
      <div style="font-family:'JetBrains Mono',monospace;font-size:24px;font-weight:700;color:#99D1FC;">${rcv2Fmt(cal.prom_mensual)}</div>
      <div style="font-size:10px;color:#475569;margin-top:2px;">proyectado / mes</div>
    </div>
    <div style="${kpiStyle('#DFFF61')}">
      <div style="font-size:9px;color:#64748b;letter-spacing:.5px;margin-bottom:6px;">REAL ÚLTIMOS 30d</div>
      <div style="font-family:'JetBrains Mono',monospace;font-size:24px;font-weight:700;color:#DFFF61;">${rcv2FmtI(consReal30)}</div>
      <div style="font-size:10px;color:${varCol};margin-top:2px;">vs proyectado: <strong>${varLabel}</strong></div>
    </div>
    <div style="${kpiStyle('#FFC04D')}">
      <div style="font-size:9px;color:#64748b;letter-spacing:.5px;margin-bottom:6px;">REAL ÚLTIMOS 7d</div>
      <div style="font-family:'JetBrains Mono',monospace;font-size:24px;font-weight:700;color:#FFC04D;">${rcv2FmtI(consReal7)}</div>
      <div style="font-size:10px;color:#475569;margin-top:2px;">${rcv2Fmt(promDia)}/día proy.</div>
    </div>
    <div style="${kpiStyle('#C084FC')}">
      <div style="font-size:9px;color:#64748b;letter-spacing:.5px;margin-bottom:6px;">PUNTO REORDEN</div>
      <div style="font-family:'JetBrains Mono',monospace;font-size:24px;font-weight:700;color:#C084FC;">${rcv2FmtI(cal.punto_reorden)}</div>
      <div style="font-size:10px;color:#475569;margin-top:2px;">rollos mínimos</div>
    </div>
    <div style="${kpiStyle('#F87171')}">
      <div style="font-size:9px;color:#64748b;letter-spacing:.5px;margin-bottom:6px;">ROTACIÓN INV.</div>
      <div style="font-family:'JetBrains Mono',monospace;font-size:24px;font-weight:700;color:#F87171;">${rotacion}x</div>
      <div style="font-size:10px;color:#475569;margin-top:2px;">consumidos/saldo</div>
    </div>
  </div>

  <!-- Fila: Consumo histórico + Inventario desglosado -->
  <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:16px;">

    <!-- Consumo histórico (barras) -->
    <div style="${secStyle}">
      <div style="font-size:12px;font-weight:700;color:#B0F2AE;margin-bottom:14px;">📈 Consumo Real por Mes (últimos 6)</div>
      <div style="display:flex;gap:8px;align-items:flex-end;height:120px;">${barsHtml}</div>
      <div style="margin-top:14px;display:flex;gap:16px;flex-wrap:wrap;">
        <div><div style="font-size:9px;color:#64748b;">PROMEDIO SEMANAL PROY.</div><div style="font-family:'JetBrains Mono',monospace;font-size:13px;color:#DFFF61;">${rcv2Fmt(promSem)} rollos</div></div>
        <div><div style="font-size:9px;color:#64748b;">PERÍODO ABASTECIMIENTO</div><div style="font-family:'JetBrains Mono',monospace;font-size:13px;color:#99D1FC;">${rcv2Fmt(cal.periodo_abast)} días</div></div>
        <div><div style="font-size:9px;color:#64748b;">ROLLOS/AÑO PROY.</div><div style="font-family:'JetBrains Mono',monospace;font-size:13px;color:#FFC04D;">${rcv2FmtI(cal.rollos_anio)}</div></div>
      </div>
    </div>

    <!-- Inventario por referencia -->
    <div style="${secStyle}">
      <div style="font-size:12px;font-weight:700;color:#99D1FC;margin-bottom:14px;">📦 Stock Inventario por Referencia</div>
      ${invStock
        ? `<div style="margin-bottom:12px;display:flex;align-items:center;gap:8px;">
            <span style="font-family:'JetBrains Mono',monospace;font-size:20px;font-weight:700;color:#B0F2AE;">${rcv2FmtI(invStock)}</span>
            <span style="font-size:10px;color:#64748b;">rollos totales en inventario Wompi</span>
           </div>
           <div style="max-height:140px;overflow-y:auto;">${refsHtml}</div>`
        : `<div style="color:#475569;font-size:11px;padding:12px 0;">Stock de inventario Wompi no disponible para este corresponsal.<br><small>Carga el módulo de Inventario para cruzar datos.</small></div>`
      }
      <div style="margin-top:14px;border-top:1px solid rgba(255,255,255,.06);padding-top:12px;">
        <div style="font-size:9px;color:#64748b;margin-bottom:6px;">DESDE CÁLCULOS (wompi_tablero_rollos_calculos)</div>
        <div style="display:flex;gap:12px;flex-wrap:wrap;">
          <div><span style="font-size:9px;color:#64748b;">Entregados:</span> <span style="font-family:'JetBrains Mono',monospace;font-size:11px;color:#B0F2AE;">${rcv2FmtI(cal.rollos_entregados)}</span></div>
          <div><span style="font-size:9px;color:#64748b;">Consumidos:</span> <span style="font-family:'JetBrains Mono',monospace;font-size:11px;color:#FFC04D;">${rcv2FmtI(cal.rollos_consumidos)}</span></div>
          <div><span style="font-size:9px;color:#64748b;">TRX desde apertura:</span> <span style="font-family:'JetBrains Mono',monospace;font-size:11px;color:#99D1FC;">${rcv2FmtI(cal.trx_desde)}</span></div>
        </div>
        ${cal.fecha_apertura ? `<div style="margin-top:6px;font-size:10px;color:#64748b;">Apertura: <strong style="color:#94a3b8;">${rcv2FmtDate(cal.fecha_apertura)}</strong> · Próx. abast: <strong style="color:#94a3b8;">${rcv2FmtDate(cal.fecha_abst)}</strong></div>` : ''}
      </div>
    </div>
  </div>

  <!-- Quiebres registrados -->
  ${quiebres.length ? `<div style="${secStyle}border-color:rgba(255,92,92,.2);">
    <div style="font-size:12px;font-weight:700;color:#FF5C5C;margin-bottom:12px;">🚨 Quiebres / Devoluciones Registradas (${quiebres.length})</div>
    <div style="max-height:150px;overflow-y:auto;">${quiebresHtml}</div>
  </div>` : ''}

  <!-- Historial de MOs -->
  <div style="${secStyle}">
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:14px;flex-wrap:wrap;gap:8px;">
      <div style="font-size:12px;font-weight:700;color:#f1f5f9;">📋 Historial de Órdenes de Movimiento (MOs)</div>
      <div style="display:flex;gap:8px;align-items:center;">
        <span style="font-size:10px;color:#64748b;font-family:'JetBrains Mono',monospace;">${movs.length} registros total</span>
        ${movs.length > MAX_MOVS ? `<span style="font-size:10px;color:#FFC04D;">mostrando ${MAX_MOVS} más recientes</span>` : ''}
      </div>
    </div>
    ${movsOrd.length ? `
    <div style="overflow-x:auto;border-radius:8px;border:1px solid rgba(255,255,255,.06);">
      <table style="width:100%;border-collapse:collapse;font-family:'Outfit',sans-serif;">
        <thead>
          <tr style="background:rgba(0,0,0,.4);">
            <th style="padding:8px 12px;font-size:9px;color:#475569;font-weight:600;letter-spacing:.5px;text-align:left;">TAREA</th>
            <th style="padding:8px 10px;font-size:9px;color:#475569;font-weight:600;letter-spacing:.5px;text-align:left;">FECHA</th>
            <th style="padding:8px 10px;font-size:9px;color:#475569;font-weight:600;letter-spacing:.5px;text-align:left;">MATERIAL</th>
            <th style="padding:8px 10px;font-size:9px;color:#475569;font-weight:600;letter-spacing:.5px;text-align:right;">CANT.</th>
            <th style="padding:8px 10px;font-size:9px;color:#475569;font-weight:600;letter-spacing:.5px;text-align:left;">ESTADO</th>
            <th style="padding:8px 10px;font-size:9px;color:#475569;font-weight:600;letter-spacing:.5px;text-align:left;">GUÍA</th>
            <th style="padding:8px 10px;font-size:9px;color:#475569;font-weight:600;letter-spacing:.5px;text-align:left;">TRANSP.</th>
          </tr>
        </thead>
        <tbody>${movsHtml}</tbody>
      </table>
    </div>` : '<div style="color:#475569;font-size:12px;padding:16px 0;">Sin órdenes de movimiento registradas para este corresponsal.</div>'}
  </div>`;
}

// ── Inicializar el panel de búsqueda ──────────────────────────────
window.initRollosComercioV2 = function() {
  // Esperar a que los datos estén listos
  const _try = (attempt) => {
    const filas = window.TABLERO_ROLLOS_FILAS || [];
    if (!filas.length && attempt < 20) {
      setTimeout(() => _try(attempt + 1), 500);
      return;
    }
    rcv2BuildIndex();
    rcv2RenderSearchUI();
  };
  _try(0);
};

// ── Renderizar la UI de búsqueda ───────────────────────────────────
function rcv2RenderSearchUI() {
  const panel = document.getElementById('panel-rollos-comercio');
  if (!panel) return;

  // Reemplazar el contenido completo del panel
  panel.innerHTML = `
  <!-- Header del panel -->
  <div style="margin-bottom:24px;">
    <div style="font-family:'Syne',sans-serif;font-size:22px;font-weight:800;color:#f1f5f9;letter-spacing:.3px;">🏪 Detalle por Comercio — Análisis Completo</div>
    <div style="font-size:12px;color:#475569;margin-top:4px;">${RCV2_SITIOS.size.toLocaleString('es-CO')} corresponsales indexados · Fuente: v_rollos_tablero_wompi + calculos</div>
  </div>

  <!-- Buscador -->
  <div style="position:relative;max-width:560px;margin-bottom:28px;">
    <div style="display:flex;gap:10px;align-items:center;">
      <div style="position:relative;flex:1;">
        <input id="rcv2-search-input" type="text"
          placeholder="🔍 Buscar corresponsal por nombre, código o ciudad..."
          oninput="rcv2Suggest(this.value)"
          onfocus="rcv2Suggest(this.value)"
          autocomplete="off"
          style="width:100%;background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.12);border-radius:10px;padding:12px 18px;color:#f1f5f9;font-family:'Outfit',sans-serif;font-size:14px;outline:none;box-sizing:border-box;transition:border-color .2s;"
          onfocus="this.style.borderColor='rgba(176,242,174,.5)'"
          onblur="setTimeout(()=>{const ul=document.getElementById('rcv2-suggestions');if(ul)ul.style.display='none'},200)">
        <ul id="rcv2-suggestions" style="display:none;position:absolute;top:calc(100% + 4px);left:0;right:0;background:#1a1a1a;border:1px solid rgba(255,255,255,.1);border-radius:10px;list-style:none;margin:0;padding:0;z-index:9999;max-height:300px;overflow-y:auto;box-shadow:0 8px 32px rgba(0,0,0,.5);"></ul>
      </div>
      <button onclick="rcv2ClearSearch()" style="background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.1);border-radius:10px;padding:11px 16px;color:#94a3b8;font-family:'Outfit',sans-serif;font-size:12px;cursor:pointer;white-space:nowrap;">✕ Limpiar</button>
    </div>
    <div style="font-size:10px;color:#475569;margin-top:6px;">Tip: escribe mínimo 2 caracteres para ver sugerencias</div>
  </div>

  <!-- Panel de detalle (vacío hasta que se seleccione un comercio) -->
  <div id="rcv2-detail-panel">
    <div style="background:rgba(255,255,255,.02);border:1px solid rgba(255,255,255,.06);border-radius:14px;padding:40px;text-align:center;color:#475569;">
      <div style="font-size:36px;margin-bottom:12px;">🔍</div>
      <div style="font-size:14px;font-weight:600;color:#64748b;">Busca un corresponsal para ver su análisis completo</div>
      <div style="font-size:12px;margin-top:6px;">Stock actual · Historial de MOs · Cobertura · Consumo · Alertas</div>
    </div>
  </div>

  <!-- Tabla resumen global (cobertura de todos los sitios) -->
  <div style="margin-top:32px;">
    <div style="font-family:'Syne',sans-serif;font-size:15px;font-weight:700;color:#f1f5f9;margin-bottom:16px;">📊 Cobertura Global — Todos los Corresponsales</div>
    <div id="rcv2-global-table"></div>
  </div>`;

  // Renderizar tabla global
  rcv2RenderGlobalTable();
}

// ── Tabla global de cobertura ──────────────────────────────────────
function rcv2RenderGlobalTable(filterTerm) {
  const el = document.getElementById('rcv2-global-table');
  if (!el) return;

  let rows = [...RCV2_SITIOS.entries()].map(([key, s]) => ({
    key, ...s,
    meses: s.cal.saldo_dias / 30,
  })).sort((a,b) => a.meses - b.meses); // más críticos primero

  if (filterTerm) {
    const t = filterTerm.toUpperCase();
    rows = rows.filter(r =>
      r.key.includes(t) ||
      r.meta.nombre_sitio.toUpperCase().includes(t) ||
      r.meta.ciudad.toUpperCase().includes(t) ||
      r.meta.departamento.toUpperCase().includes(t)
    );
  }

  // Stats rápidos
  const criticos  = rows.filter(r => r.meses < 1).length;
  const alertas   = rows.filter(r => r.meses >= 1 && r.meses < 2).length;
  const warns     = rows.filter(r => r.meses >= 2 && r.meses < 3).length;
  const oks       = rows.filter(r => r.meses >= 3).length;
  const slaIncump = rows.filter(r => r.meses < 3).length;

  const PAGE_SIZE = 100;
  const maxShown  = Math.min(rows.length, PAGE_SIZE);
  const shown     = rows.slice(0, maxShown);

  el.innerHTML = `
  <!-- Stats -->
  <div style="display:flex;flex-wrap:wrap;gap:10px;margin-bottom:16px;">
    ${[
      ['🔴 Críticos (<1m)',  criticos,  '#FF5C5C'],
      ['🟠 Alerta (<2m)',    alertas,   '#FFC04D'],
      ['🟡 Atención (<3m)', warns,     '#DFFF61'],
      ['🟢 OK (≥3m)',        oks,       '#B0F2AE'],
      ['✗ SLA Incumplido',  slaIncump, '#FF5C5C'],
    ].map(([l,n,c]) => `<div style="background:${c}12;border:1px solid ${c}33;border-radius:10px;padding:10px 16px;display:flex;align-items:center;gap:8px;">
      <span style="font-family:'JetBrains Mono',monospace;font-size:18px;font-weight:700;color:${c};">${n}</span>
      <span style="font-size:10px;color:#64748b;">${l}</span>
    </div>`).join('')}
    <div style="flex:1;display:flex;align-items:center;justify-content:flex-end;">
      <input type="text" placeholder="Filtrar tabla..." oninput="rcv2RenderGlobalTable(this.value)"
        style="background:rgba(255,255,255,.04);border:1px solid rgba(255,255,255,.1);border-radius:8px;padding:7px 12px;color:#f1f5f9;font-family:'Outfit',sans-serif;font-size:12px;outline:none;width:200px;">
    </div>
  </div>

  <!-- Tabla -->
  <div style="overflow-x:auto;border-radius:10px;border:1px solid rgba(255,255,255,.06);">
    <table style="width:100%;border-collapse:collapse;font-family:'Outfit',sans-serif;font-size:12px;">
      <thead>
        <tr style="background:rgba(0,0,0,.4);">
          ${['#','CORRESPONSAL','CIUDAD','COBERTURA','SALDO ROL.','PROM/MES','P.REORDEN','SLA','ESTADO'].map(h=>
            `<th style="padding:9px 12px;font-size:9px;color:#475569;font-weight:700;letter-spacing:.5px;text-align:${h==='SALDO ROL.'||h==='PROM/MES'||h==='P.REORDEN'?'right':'left'};">${h}</th>`
          ).join('')}
        </tr>
      </thead>
      <tbody>
        ${shown.map((r, i) => {
          const sg    = rcv2Semaforo(r.meses);
          const sla   = r.meses >= 3;
          const bajPR = r.cal.saldo_rollos > 0 && r.cal.saldo_rollos <= r.cal.punto_reorden;
          return `<tr style="border-bottom:1px solid rgba(255,255,255,.04);${i%2?'background:rgba(255,255,255,.012)':''}cursor:pointer;transition:background .15s;"
            onclick="rcv2Select('${r.key}')"
            onmouseover="this.style.background='rgba(176,242,174,.05)'"
            onmouseout="this.style.background='${i%2?'rgba(255,255,255,.012)':'transparent'}'">
            <td style="padding:8px 12px;font-size:10px;color:#475569;">${i+1}</td>
            <td style="padding:8px 12px;"><div style="font-size:11px;font-weight:600;color:#e2e8f0;">${r.meta.nombre_sitio}</div><div style="font-size:9px;color:#475569;font-family:'JetBrains Mono',monospace;">${r.key}</div></td>
            <td style="padding:8px 10px;font-size:11px;color:#94a3b8;">${r.meta.ciudad||'—'}</td>
            <td style="padding:8px 12px;">
              <div style="display:flex;align-items:center;gap:6px;">
                <div style="width:${Math.min(r.meses/6*60,60).toFixed(0)}px;height:4px;background:${sg.col};border-radius:2px;"></div>
                <span style="font-family:'JetBrains Mono',monospace;font-size:12px;font-weight:700;color:${sg.col};">${r.meses.toFixed(1)}m</span>
              </div>
            </td>
            <td style="padding:8px 10px;font-family:'JetBrains Mono',monospace;font-size:12px;color:#B0F2AE;text-align:right;">${rcv2FmtI(r.cal.saldo_rollos)}</td>
            <td style="padding:8px 10px;font-family:'JetBrains Mono',monospace;font-size:11px;color:#94a3b8;text-align:right;">${rcv2Fmt(r.cal.prom_mensual)}</td>
            <td style="padding:8px 10px;font-family:'JetBrains Mono',monospace;font-size:11px;color:${bajPR?'#FFC04D':'#64748b'};text-align:right;">${rcv2FmtI(r.cal.punto_reorden)}${bajPR?' ⬇':''}</td>
            <td style="padding:8px 10px;"><span style="font-size:10px;color:${sla?'#B0F2AE':'#FF5C5C'};background:${sla?'#B0F2AE':'#FF5C5C'}18;border:1px solid ${sla?'#B0F2AE':'#FF5C5C'}33;border-radius:20px;padding:2px 8px;">${sla?'✓ OK':'✗ INC'}</span></td>
            <td style="padding:8px 10px;font-size:10px;color:#64748b;">${r.cal.estado_punto||'—'}</td>
          </tr>`;
        }).join('')}
      </tbody>
    </table>
    ${rows.length > PAGE_SIZE ? `<div style="padding:10px 16px;font-size:11px;color:#475569;border-top:1px solid rgba(255,255,255,.06);">Mostrando ${maxShown} de ${rows.length} registros. Usa el buscador para filtrar.</div>` : ''}
  </div>`;
}

// ── Limpiar búsqueda ───────────────────────────────────────────────
window.rcv2ClearSearch = function() {
  const inp = document.getElementById('rcv2-search-input');
  if (inp) inp.value = '';
  RCV2_CURRENT = null;
  const det = document.getElementById('rcv2-detail-panel');
  if (det) det.innerHTML = `
    <div style="background:rgba(255,255,255,.02);border:1px solid rgba(255,255,255,.06);border-radius:14px;padding:40px;text-align:center;color:#475569;">
      <div style="font-size:36px;margin-bottom:12px;">🔍</div>
      <div style="font-size:14px;font-weight:600;color:#64748b;">Busca un corresponsal para ver su análisis completo</div>
      <div style="font-size:12px;margin-top:6px;">Stock actual · Historial de MOs · Cobertura · Consumo · Alertas</div>
    </div>`;
};

// ── Hook al tab rollos-comercio ────────────────────────────────────
// dashboard.js llama renderRollosComercioTable() cuando se abre el tab.
// Sobrescribimos para inicializar nuestro módulo en su lugar.
window.renderRollosComercioTable = function() {
  if (!RCV2_READY) {
    window.initRollosComercioV2();
  }
  // También llamar renderRollosInvComercio si existe (para compatibilidad)
  if (typeof window._origRenderRollosInvComercio === 'function') {
    // no llamar — el panel es ahora v2
  }
};

// ── Exponer función de búsqueda global ────────────────────────────
window.rcv2Suggest = rcv2Suggest;
window.rcv2RenderGlobalTable = rcv2RenderGlobalTable;

// ── Auto-iniciar cuando el tab se abra ────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  // Si el tab rollos-comercio ya está activo, inicializar
  const panel = document.getElementById('panel-rollos-comercio');
  if (panel && panel.style.display !== 'none') {
    window.initRollosComercioV2();
  }
});

console.log('[RCV2] rollos_comercio_v2.js cargado');