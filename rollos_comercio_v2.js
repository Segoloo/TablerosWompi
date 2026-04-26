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
let _rcv2Initializing = false; // evita lanzar múltiples init en paralelo

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

  // Mapa auxiliar: codigo_sitio → cal_codigo_mo, para resolver alias entre claves.
  // Cuando una fila tiene ambos campos, registramos que codigo_sitio es alias de cal_codigo_mo.
  // Así evitamos crear entradas duplicadas (una por codigo_sitio, otra por cal_codigo_mo)
  // para el mismo corresponsal físico, lo que inflaba los conteos de KPIs.
  const aliasMap = new Map(); // codigo_sitio → cal_codigo_mo (clave canónica)
  filas.forEach(f => {
    const codSitio = (f.codigo_sitio || f.cal_codigo_sitio || '').trim();
    const codMO    = (f.cal_codigo_mo || '').trim();
    if (codSitio && codMO && codSitio !== codMO && !aliasMap.has(codSitio)) {
      aliasMap.set(codSitio, codMO);
    }
  });
  // Resuelve cualquier clave a su forma canónica (cal_codigo_mo tiene prioridad)
  const canonical = k => aliasMap.get(k) || k;

  // IMPORTANTE: se itera todo el array.
  // Los cal_* solo se asignan desde filas con join_nivel > 0 Y cal_saldo_dias != 0,
  // para evitar que sitios queden con saldo=0 por tomar una fila sin join match.
  // La primera fila válida que aparezca por sitio gana; las siguientes se ignoran para cal.

  filas.forEach(f => {
    const codSitio = (f.codigo_sitio || f.cal_codigo_sitio || '').trim();
    const codMO    = (f.cal_codigo_mo || '').trim();
    // Usar siempre la clave canónica: prioriza cal_codigo_mo para evitar duplicados
    const key      = canonical(codMO || codSitio);
    if (!key) return;

    if (!RCV2_SITIOS.has(key)) {
      const meta = {
        nombre_sitio : f.nombre_sitio || f.nombre_ubicacion_destino || key,
        departamento : f.departamento || '',
        ciudad       : f.Ciudad || f.ciudad || '',
        nit          : f.nit || '',
        tipologia    : f.tipologia || '',
        proyecto     : f.proyecto || '',
        red          : f.red_asociada || '',
      };
      // cal vacío por defecto — se llenará cuando haya join_nivel > 0
      const calVacio = {
        codigo_sitio: key, codigo_mo: codMO, join_nivel: 0,
        estado_punto: '', prom_mensual: 0, rollos_prom_mes: 0,
        periodo_abast: 0, rollos_periodo: 0, punto_reorden: 0,
        saldo_rollos: 0, saldo_dias: 0, saldo_valor: 0,
        rollos_entregados: 0, rollos_consumidos: 0, trx_desde: 0,
        fecha_apertura: '', fecha_abst: '', rollos_anio: 0,
      };
      RCV2_SITIOS.set(key, { cal: calVacio, meta, movs: [], _calSet: false });
    }

    // Intentar actualizar cal si esta fila tiene join real y el sitio aún no lo tiene.
    // IMPORTANTE: la condición ya NO excluye saldo_dias=0, porque un corresponsal con
    // stock agotado (saldo=0) es el caso más crítico y debe aparecer en la tabla global.
    // _calSet = true siempre que join_nivel > 0, sin importar el saldo.
    const sitio = RCV2_SITIOS.get(key);
    const nivel = parseInt(f.join_nivel) || 0;
    if (!sitio._calSet && nivel > 0) {
      sitio.cal = {
        codigo_sitio    : key,
        codigo_mo       : (f.cal_codigo_mo || '').trim(),
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
        join_nivel      : nivel,
      };
      sitio._calSet = true;
    }

    // Agregar movimiento (cada fila del tablero es un movimiento de MO)
    const sitioMov = RCV2_SITIOS.get(key);
    if (f.tarea) {
      sitioMov.movs.push({
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
  _rcv2Initializing = false;
  const conCal   = [...RCV2_SITIOS.values()].filter(s => s._calSet).length;
  const sinCal   = RCV2_SITIOS.size - conCal;
  console.log('[RCV2] Índice construido:', RCV2_SITIOS.size, 'sitios únicos |',
    conCal, 'con cálculos reales |', sinCal, 'sin join (excluidos de tabla global)');
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
    <div style="font-size:12px;color:#475569;margin-top:4px;">${[...RCV2_SITIOS.values()].filter(s=>s._calSet).length.toLocaleString('es-CO')} corresponsales con datos de cálculo · ${RCV2_SITIOS.size.toLocaleString('es-CO')} sitios totales indexados</div>
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


// ══════════════════════════════════════════════════════════════════
//  MODAL DRILL-DOWN KPI — Cobertura Global
// ══════════════════════════════════════════════════════════════════

// Columnas del modal de corresponsales
const RCV2_MODAL_COLS = [
  { label:'#',              fn:(r,i)=>i+1,                                         w:40  },
  { label:'Código',         fn:r=>r.key,                                            w:110 },
  { label:'Corresponsal',   fn:r=>r.meta.nombre_sitio,                              w:220 },
  { label:'Ciudad',         fn:r=>r.meta.ciudad||'—',                               w:110 },
  { label:'Departamento',   fn:r=>r.meta.departamento||'—',                         w:130 },
  { label:'Proyecto',       fn:r=>r.meta.proyecto||'—',                             w:110 },
  { label:'Cobertura (m)',  fn:r=>r.meses!==null?r.meses.toFixed(2):'Sin datos',    w:110, isNum:true },
  { label:'Saldo Rollos',   fn:r=>r._calSet?r.cal.saldo_rollos:'-',                 w:100, isNum:true },
  { label:'Saldo Días',     fn:r=>r._calSet?r.cal.saldo_dias:'-',                   w:90,  isNum:true },
  { label:'Prom/Mes',       fn:r=>r._calSet?r.cal.prom_mensual.toFixed(1):'-',      w:90,  isNum:true },
  { label:'P.Reorden',      fn:r=>r._calSet?r.cal.punto_reorden:'-',                w:90,  isNum:true },
  { label:'Período Abast.', fn:r=>r._calSet?r.cal.periodo_abast:'-',               w:110, isNum:true },
  { label:'Estado Punto',   fn:r=>r._calSet?r.cal.estado_punto||'—':'—',            w:120 },
  { label:'Fecha Abast.',   fn:r=>r._calSet?rcv2FmtDate(r.cal.fecha_abst)||'—':'—', w:110 },
  { label:'SLA',            fn:r=>r.meses===null?'—':r.meses>=3?'✓ OK':'✗ INC',    w:70  },
];

let _rcv2ModalRows  = [];
let _rcv2ModalTitle = '';
let _rcv2ModalPage  = 1;
const RCV2_MODAL_PAGE = 100;

function rcv2OpenKpiModal(title, rows) {
  _rcv2ModalRows  = rows;
  _rcv2ModalTitle = title;
  _rcv2ModalPage  = 1;

  let modal = document.getElementById('rcv2-kpi-modal');
  if (!modal) {
    modal = document.createElement('div');
    modal.id = 'rcv2-kpi-modal';
    modal.style.cssText = 'position:fixed;inset:0;z-index:99999;display:flex;align-items:center;justify-content:center;background:rgba(0,0,0,.8);backdrop-filter:blur(6px);padding:20px;';
    modal.addEventListener('click', e => { if (e.target === modal) modal.remove(); });
    modal.innerHTML = `
      <div style="background:#0f1923;border:1px solid rgba(255,255,255,.1);border-radius:16px;
        width:100%;max-width:1300px;max-height:90vh;display:flex;flex-direction:column;
        box-shadow:0 32px 96px rgba(0,0,0,.9);">
        <!-- Header -->
        <div style="display:flex;align-items:center;justify-content:space-between;
          padding:20px 24px;border-bottom:1px solid rgba(255,255,255,.08);flex-shrink:0;">
          <div>
            <div id="rcv2-modal-title" style="font-family:'Syne',sans-serif;font-size:18px;font-weight:700;color:#B0F2AE;"></div>
            <div id="rcv2-modal-count" style="font-size:12px;color:#475569;margin-top:3px;"></div>
          </div>
          <div style="display:flex;gap:8px;align-items:center;">
            <!-- Búsqueda inline -->
            <input id="rcv2-modal-search" type="text" placeholder="🔍 Filtrar..." autocomplete="off"
              oninput="rcv2ModalFilter(this.value)"
              style="background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.1);border-radius:8px;
                padding:7px 12px;color:#f1f5f9;font-family:'Outfit',sans-serif;font-size:12px;outline:none;width:180px;">
            <button onclick="rcv2ModalExportExcel()"
              style="background:rgba(176,242,174,.1);border:1px solid rgba(176,242,174,.3);color:#B0F2AE;
                padding:7px 14px;border-radius:8px;cursor:pointer;font-size:12px;font-family:'Outfit',sans-serif;
                display:flex;align-items:center;gap:6px;transition:background .2s;"
              onmouseover="this.style.background='rgba(176,242,174,.2)'"
              onmouseout="this.style.background='rgba(176,242,174,.1)'">
              ⬇ Excel
            </button>
            <button onclick="document.getElementById('rcv2-kpi-modal').remove()"
              style="background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.1);color:#94a3b8;
                width:32px;height:32px;border-radius:8px;cursor:pointer;font-size:18px;
                display:flex;align-items:center;justify-content:center;line-height:1;">×</button>
          </div>
        </div>
        <!-- Body tabla -->
        <div id="rcv2-modal-body" style="overflow:auto;flex:1;padding:4px;"></div>
        <!-- Footer paginación -->
        <div id="rcv2-modal-footer" style="padding:12px 20px;border-top:1px solid rgba(255,255,255,.06);
          display:flex;align-items:center;justify-content:space-between;flex-shrink:0;gap:8px;flex-wrap:wrap;">
          <span id="rcv2-modal-pag-info" style="font-size:11px;color:#475569;font-family:'JetBrains Mono',monospace;"></span>
          <div id="rcv2-modal-pag-btns" style="display:flex;gap:4px;flex-wrap:wrap;"></div>
        </div>
      </div>`;
    document.body.appendChild(modal);
  } else {
    modal.style.display = 'flex';
    const si = document.getElementById('rcv2-modal-search');
    if (si) si.value = '';
  }

  window._rcv2ModalFiltered = rows;
  _rcv2ModalRenderBody();
}

function rcv2ModalFilter(term) {
  const t = term.trim().toUpperCase();
  window._rcv2ModalFiltered = t
    ? _rcv2ModalRows.filter(r =>
        r.key.toUpperCase().includes(t) ||
        r.meta.nombre_sitio.toUpperCase().includes(t) ||
        (r.meta.ciudad||'').toUpperCase().includes(t) ||
        (r.meta.departamento||'').toUpperCase().includes(t) ||
        (r.meta.proyecto||'').toUpperCase().includes(t) ||
        (r._calSet && (r.cal.estado_punto||'').toUpperCase().includes(t))
      )
    : _rcv2ModalRows;
  _rcv2ModalPage = 1;
  _rcv2ModalRenderBody();
}

function _rcv2ModalRenderBody() {
  const titleEl = document.getElementById('rcv2-modal-title');
  const countEl = document.getElementById('rcv2-modal-count');
  const body    = document.getElementById('rcv2-modal-body');
  const pagInfo = document.getElementById('rcv2-modal-pag-info');
  const pagBtns = document.getElementById('rcv2-modal-pag-btns');
  if (!body) return;

  const filtered = window._rcv2ModalFiltered || _rcv2ModalRows;
  const total    = filtered.length;
  const pages    = Math.max(1, Math.ceil(total / RCV2_MODAL_PAGE));
  if (_rcv2ModalPage > pages) _rcv2ModalPage = 1;
  const start    = (_rcv2ModalPage - 1) * RCV2_MODAL_PAGE;
  const slice    = filtered.slice(start, start + RCV2_MODAL_PAGE);

  if (titleEl) titleEl.textContent = _rcv2ModalTitle;
  if (countEl) countEl.textContent = total !== _rcv2ModalRows.length
    ? `${total.toLocaleString('es-CO')} de ${_rcv2ModalRows.length.toLocaleString('es-CO')} corresponsales`
    : `${total.toLocaleString('es-CO')} corresponsales`;

  if (!total) {
    body.innerHTML = '<div style="text-align:center;padding:60px;color:#475569;font-family:\'Outfit\',sans-serif;">Sin registros para este filtro</div>';
  } else {
    const semColor = r => {
      if (r.meses === null) return '#475569';
      if (r.meses < 1)  return '#FF5C5C';
      if (r.meses < 2)  return '#FFC04D';
      if (r.meses < 3)  return '#DFFF61';
      return '#B0F2AE';
    };

    const hdrs = RCV2_MODAL_COLS.map(c =>
      `<th style="padding:9px 12px;font-size:9px;color:#475569;font-weight:700;letter-spacing:.5px;
        white-space:nowrap;text-align:${c.isNum?'right':'left'};
        border-bottom:2px solid rgba(176,242,174,.15);position:sticky;top:0;
        background:#0a1520;z-index:2;">${c.label}</th>`
    ).join('');

    const trows = slice.map((r, idx) => {
      const bg    = idx % 2 === 0 ? 'transparent' : 'rgba(255,255,255,.013)';
      const sc    = semColor(r);
      const cells = RCV2_MODAL_COLS.map((c, ci) => {
        const v   = c.fn(r, start + idx);
        const num = c.isNum && r._calSet && r.meses !== null;
        let inner;
        if (ci === 6 && r.meses !== null) {
          // Cobertura: chip de color
          inner = `<span style="font-family:'JetBrains Mono',monospace;font-size:12px;font-weight:700;color:${sc};">${v}m</span>`;
        } else if (ci === 14) {
          // SLA
          const slaOk = v === '✓ OK';
          const slaNo = v === '✗ INC';
          inner = slaOk || slaNo
            ? `<span style="font-size:10px;padding:2px 8px;border-radius:20px;
                color:${slaOk?'#B0F2AE':'#FF5C5C'};
                background:${slaOk?'#B0F2AE':'#FF5C5C'}18;
                border:1px solid ${slaOk?'#B0F2AE':'#FF5C5C'}33;">${v}</span>`
            : `<span style="color:#475569;">—</span>`;
        } else if (num) {
          inner = `<span style="font-family:'JetBrains Mono',monospace;font-size:11px;color:#99D1FC;">${Number(v).toLocaleString('es-CO')}</span>`;
        } else {
          inner = `<span style="color:${v==='—'||v==='-'?'#334155':'#e2e8f0'}">${v}</span>`;
        }
        return `<td style="padding:7px 12px;font-size:11px;border-bottom:1px solid rgba(255,255,255,.04);
          white-space:nowrap;text-align:${c.isNum?'right':'left'};
          max-width:${c.w||200}px;overflow:hidden;text-overflow:ellipsis;">${inner}</td>`;
      }).join('');
      return `<tr style="background:${bg};transition:background .1s;"
        onmouseover="this.style.background='rgba(176,242,174,.04)'"
        onmouseout="this.style.background='${bg}'">${cells}</tr>`;
    }).join('');

    body.innerHTML = `
      <div style="overflow-x:auto;">
        <table style="width:100%;border-collapse:collapse;font-family:'Outfit',sans-serif;font-size:12px;min-width:1400px;">
          <thead><tr>${hdrs}</tr></thead>
          <tbody>${trows}</tbody>
        </table>
      </div>`;
  }

  // Paginación
  if (pagInfo) pagInfo.textContent = `Pág ${_rcv2ModalPage} / ${pages}  ·  ${start+1}–${Math.min(start+RCV2_MODAL_PAGE,total)} de ${total}`;
  if (pagBtns) {
    const btn = (p, label, active) =>
      `<button onclick="_rcv2ModalPage=${p};_rcv2ModalRenderBody()"
        style="padding:4px 10px;border-radius:6px;font-size:11px;cursor:pointer;
          border:1px solid ${active?'rgba(176,242,174,.5)':'rgba(255,255,255,.08)'};
          background:${active?'rgba(176,242,174,.12)':'rgba(255,255,255,.03)'};
          color:${active?'#B0F2AE':'#94a3b8'};font-family:'JetBrains Mono',monospace;">${label}</button>`;
    let h = '';
    if (_rcv2ModalPage > 1) h += btn(_rcv2ModalPage-1, '‹', false);
    const half=2, st=Math.max(1,_rcv2ModalPage-half), en=Math.min(pages,_rcv2ModalPage+half);
    if (st>1){ h+=btn(1,'1',false); if(st>2) h+=`<span style="color:#475569;padding:0 4px;">…</span>`; }
    for(let p=st;p<=en;p++) h+=btn(p,p,p===_rcv2ModalPage);
    if(en<pages){ if(en<pages-1) h+=`<span style="color:#475569;padding:0 4px;">…</span>`; h+=btn(pages,pages,false); }
    if(_rcv2ModalPage<pages) h+=btn(_rcv2ModalPage+1,'›',false);
    pagBtns.innerHTML = h;
  }
}

function rcv2ModalExportExcel() {
  const data = (window._rcv2ModalFiltered || _rcv2ModalRows);
  if (!data.length) { alert('Sin datos para exportar'); return; }
  if (!window.XLSX) { alert('XLSX no disponible'); return; }

  const rows = data.map((r, i) => {
    const obj = {};
    RCV2_MODAL_COLS.forEach(c => {
      // Saltar columna '#' en Excel — se numera sola
      if (c.label === '#') return;
      let v = c.fn(r, i);
      // Limpiar emojis de SLA para Excel
      if (c.label === 'SLA') v = v === '✓ OK' ? 'OK' : v === '✗ INC' ? 'INCUMPLIDO' : '—';
      obj[c.label] = v;
    });
    return obj;
  });

  const ws = window.XLSX.utils.json_to_sheet(rows);
  // Ancho de columnas
  ws['!cols'] = Object.keys(rows[0]||{}).map(k => ({
    wch: Math.max(k.length, ...rows.slice(0,50).map(r=>String(r[k]||'').length), 10)
  }));

  // Hoja resumen KPIs del modal
  const totalRows  = _rcv2ModalRows.length;
  const conDatos   = _rcv2ModalRows.filter(r=>r._calSet).length;
  const criticos   = _rcv2ModalRows.filter(r=>r.meses!==null&&r.meses<1).length;
  const alertas    = _rcv2ModalRows.filter(r=>r.meses!==null&&r.meses>=1&&r.meses<2).length;
  const warns      = _rcv2ModalRows.filter(r=>r.meses!==null&&r.meses>=2&&r.meses<3).length;
  const oks        = _rcv2ModalRows.filter(r=>r.meses!==null&&r.meses>=3).length;
  const sinDatos   = _rcv2ModalRows.filter(r=>r.meses===null).length;
  const slaIncump  = criticos+alertas+warns;

  const summary = [
    {KPI:'Grupo',            Valor:_rcv2ModalTitle},
    {KPI:'Total',            Valor:totalRows},
    {KPI:'Con datos stock',  Valor:conDatos},
    {KPI:'Críticos (<1m)',   Valor:criticos},
    {KPI:'Alerta (<2m)',     Valor:alertas},
    {KPI:'Atención (<3m)',   Valor:warns},
    {KPI:'OK (≥3m)',         Valor:oks},
    {KPI:'SLA Incumplido',   Valor:slaIncump},
    {KPI:'Sin datos stock',  Valor:sinDatos},
  ];
  const wsSummary = window.XLSX.utils.json_to_sheet(summary);
  wsSummary['!cols'] = [{wch:20},{wch:15}];

  const wb = window.XLSX.utils.book_new();
  window.XLSX.utils.book_append_sheet(wb, ws, 'Corresponsales');
  window.XLSX.utils.book_append_sheet(wb, wsSummary, 'Resumen KPI');

  const safeTitle = _rcv2ModalTitle.replace(/[^a-zA-Z0-9_\-áéíóúüñÁÉÍÓÚÜÑ ]/g,'').trim().replace(/\s+/g,'_');
  window.XLSX.writeFile(wb, `cobertura_${safeTitle}_${new Date().toISOString().slice(0,10)}.xlsx`);
}

// Exponer globalmente
window.rcv2OpenKpiModal     = rcv2OpenKpiModal;
window.rcv2ModalFilter      = rcv2ModalFilter;
window._rcv2ModalRenderBody = _rcv2ModalRenderBody;
window.rcv2ModalExportExcel = rcv2ModalExportExcel;

// ── Tabla global de cobertura ──────────────────────────────────────
function rcv2RenderGlobalTable(filterTerm) {
  const el = document.getElementById('rcv2-global-table');
  if (!el) return;

  // Incluir todos los sitios indexados.
  // Los que tienen _calSet=true tienen datos reales de stock/cálculo.
  // Los que no tienen _calSet se muestran como "Sin datos" (no se ocultan,
  // porque son corresponsales reales sin match en la tabla de cálculos).
  let rows = [...RCV2_SITIOS.entries()]
    .map(([key, s]) => ({
      key, ...s,
      meses: s._calSet ? s.cal.saldo_dias / 30 : null,  // null = sin datos
    }))
    .sort((a, b) => {
      // Sin datos van al final; con datos se ordenan más críticos primero
      if (a.meses === null && b.meses === null) return 0;
      if (a.meses === null) return 1;
      if (b.meses === null) return -1;
      return a.meses - b.meses;
    });

  if (filterTerm) {
    const t = filterTerm.toUpperCase();
    rows = rows.filter(r =>
      r.key.includes(t) ||
      r.meta.nombre_sitio.toUpperCase().includes(t) ||
      r.meta.ciudad.toUpperCase().includes(t) ||
      r.meta.departamento.toUpperCase().includes(t)
    );
  }

  // Stats rápidos — cada sitio se cuenta exactamente una vez
  // meses===null significa sin datos de cálculo (no se clasifican en semáforo)
  const total     = rows.length;
  const sinDatos  = rows.filter(r => r.meses === null).length;
  const criticos  = rows.filter(r => r.meses !== null && r.meses < 1).length;
  const alertas   = rows.filter(r => r.meses !== null && r.meses >= 1 && r.meses < 2).length;
  const warns     = rows.filter(r => r.meses !== null && r.meses >= 2 && r.meses < 3).length;
  const oks       = rows.filter(r => r.meses !== null && r.meses >= 3).length;
  // criticos + alertas + warns + oks + sinDatos debe ser === total
  const slaIncump = criticos + alertas + warns; // solo los que tienen datos y < 3m

  // Guardar grupos para el modal drill-down
  window._rcv2KpiGroups = {
    total:    rows,
    criticos: rows.filter(r => r.meses !== null && r.meses < 1),
    alertas:  rows.filter(r => r.meses !== null && r.meses >= 1 && r.meses < 2),
    warns:    rows.filter(r => r.meses !== null && r.meses >= 2 && r.meses < 3),
    oks:      rows.filter(r => r.meses !== null && r.meses >= 3),
    sla:      rows.filter(r => r.meses !== null && r.meses < 3),
    sinDatos: rows.filter(r => r.meses === null),
  };

  const PAGE_SIZE = 100;
  const maxShown  = Math.min(rows.length, PAGE_SIZE);
  const shown     = rows.slice(0, maxShown);

  el.innerHTML = `
  <!-- Stats -->
  <div style="display:flex;flex-wrap:wrap;gap:10px;margin-bottom:16px;">
    ${[
      ['📍 Total',           total,     '#99D1FC', 'total'],
      ['🔴 Críticos (<1m)',  criticos,  '#FF5C5C', 'criticos'],
      ['🟠 Alerta (<2m)',    alertas,   '#FFC04D', 'alertas'],
      ['🟡 Atención (<3m)', warns,     '#DFFF61', 'warns'],
      ['🟢 OK (≥3m)',        oks,       '#B0F2AE', 'oks'],
      ['✗ SLA Incumplido',  slaIncump, '#FF5C5C', 'sla'],
      ['⚪ Sin datos stock', sinDatos,  '#475569', 'sinDatos'],
    ].map(([l,n,c,grp]) => `<div
        onclick="rcv2OpenKpiModal('${l.replace(/'/g,"\\'")} (${n})', window._rcv2KpiGroups['${grp}'] || [])"
        title="Click para ver corresponsales"
        style="background:${c}12;border:1px solid ${c}33;border-radius:10px;padding:10px 16px;
          display:flex;align-items:center;gap:8px;cursor:pointer;transition:background .18s,transform .15s;"
        onmouseover="this.style.background='${c}22';this.style.transform='translateY(-2px)'"
        onmouseout="this.style.background='${c}12';this.style.transform='translateY(0)'">
      <span style="font-family:'JetBrains Mono',monospace;font-size:18px;font-weight:700;color:${c};">${n}</span>
      <div style="line-height:1.3;">
        <div style="font-size:10px;color:#64748b;">${l}</div>
        <div style="font-size:9px;color:${c};opacity:.6;">Ver listado →</div>
      </div>
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
          const noData = r.meses === null;
          const sg    = noData ? { col: '#475569' } : rcv2Semaforo(r.meses);
          const sla   = !noData && r.meses >= 3;
          const bajPR = !noData && r.cal.saldo_rollos > 0 && r.cal.saldo_rollos <= r.cal.punto_reorden;
          const cobCell = noData
            ? `<span style="font-size:10px;color:#475569;font-style:italic;">Sin datos</span>`
            : `<div style="display:flex;align-items:center;gap:6px;">
                <div style="width:${Math.min(r.meses/6*60,60).toFixed(0)}px;height:4px;background:${sg.col};border-radius:2px;"></div>
                <span style="font-family:'JetBrains Mono',monospace;font-size:12px;font-weight:700;color:${sg.col};">${r.meses.toFixed(1)}m</span>
               </div>`;
          const slaCell = noData
            ? `<span style="font-size:10px;color:#475569;">—</span>`
            : `<span style="font-size:10px;color:${sla?'#B0F2AE':'#FF5C5C'};background:${sla?'#B0F2AE':'#FF5C5C'}18;border:1px solid ${sla?'#B0F2AE':'#FF5C5C'}33;border-radius:20px;padding:2px 8px;">${sla?'✓ OK':'✗ INC'}</span>`;
          return `<tr style="border-bottom:1px solid rgba(255,255,255,.04);${i%2?'background:rgba(255,255,255,.012)':''}${noData?'opacity:.55;':''}cursor:pointer;transition:background .15s;"
            onclick="rcv2Select('${r.key}')"
            onmouseover="this.style.background='rgba(176,242,174,.05)'"
            onmouseout="this.style.background='${i%2?'rgba(255,255,255,.012)':'transparent'}'">
            <td style="padding:8px 12px;font-size:10px;color:#475569;">${i+1}</td>
            <td style="padding:8px 12px;"><div style="font-size:11px;font-weight:600;color:#e2e8f0;">${r.meta.nombre_sitio}</div><div style="font-size:9px;color:#475569;font-family:'JetBrains Mono',monospace;">${r.key}</div></td>
            <td style="padding:8px 10px;font-size:11px;color:#94a3b8;">${r.meta.ciudad||'—'}</td>
            <td style="padding:8px 12px;">${cobCell}</td>
            <td style="padding:8px 10px;font-family:'JetBrains Mono',monospace;font-size:12px;color:#B0F2AE;text-align:right;">${noData?'—':rcv2FmtI(r.cal.saldo_rollos)}</td>
            <td style="padding:8px 10px;font-family:'JetBrains Mono',monospace;font-size:11px;color:#94a3b8;text-align:right;">${noData?'—':rcv2Fmt(r.cal.prom_mensual)}</td>
            <td style="padding:8px 10px;font-family:'JetBrains Mono',monospace;font-size:11px;color:${bajPR?'#FFC04D':'#64748b'};text-align:right;">${noData?'—':rcv2FmtI(r.cal.punto_reorden)+(bajPR?' ⬇':'')}</td>
            <td style="padding:8px 10px;">${slaCell}</td>
            <td style="padding:8px 10px;font-size:10px;color:#64748b;">${noData?'—':r.cal.estado_punto||'—'}</td>
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
  console.log('[RCV2] renderRollosComercioTable | RCV2_READY:', RCV2_READY,
    '| filas:', (window.TABLERO_ROLLOS_FILAS || []).length);

  if (RCV2_READY) {
    rcv2RenderGlobalTable();
    return;
  }

  // Si los datos ya están disponibles, construir índice de forma síncrona ahora
  const filas = window.TABLERO_ROLLOS_FILAS || [];
  if (filas.length > 0) {
    console.log('[RCV2] Datos listos, construyendo índice...');
    rcv2BuildIndex();
    rcv2RenderSearchUI();
    return;
  }

  // Datos aún no disponibles — esperar con polling
  if (_rcv2Initializing) return;
  _rcv2Initializing = true;
  window.initRollosComercioV2();
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