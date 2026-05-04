'use strict';
/**
 * ╔══════════════════════════════════════════════════════════════════╗
 * ║  puntos_reorden.js — Tab "Puntos de Reorden"                    ║
 * ║  Rollos · Datáfonos CB · Pinpads CB · Datáfonos VP             ║
 * ╚══════════════════════════════════════════════════════════════════╝
 */

// ── Configuración de puntos de reorden ───────────────────────────
// Valores extraídos del archivo Punto_reorden_cb_-_wompi.xlsx
const PR_CONFIG = [
  {
    id:        'rollos',
    titulo:    'Punto de Reorden · Rollos',
    icono:     '🎞️',
    // Hoja PR-Rollos del Excel
    punto:     530611,   // Punto Reorden final (con stock de seguridad)
    solicitud: 897958,   // Pedido sugerido
    stockMin:  530611,
    stockMax:  null,
    categoria: 'Rollos',
    negocio:   null,     // CB + VP
    acento:    '#DFFF61',
    formula: {
      leadTime:         1.5,      // meses
      mesesAbast:       3,
      demandaMes:       244897.6, // Total Rollos Trim / 3
      demandaTrim:      734692.8, // rollos / trimestre
      puntoSinSeg:      367346.4, // Lead Time × demanda mensual
      diasSeguridad:    20,       // días de stock de seguridad
      stockSeguridad:   163265.1,
      puntoFinal:       530611.5,
      pedido:           897957.9,
      componentes: [
        { label: 'Distribución masiva (trim)',  valor: 650448,   unidad: 'rollos' },
        { label: 'Aperturas (trim)',            valor: 19200,    unidad: 'rollos' },
        { label: 'Órdenes de cambio (trim)',    valor: 65044.8,  unidad: 'rollos' },
      ]
    }
  },
  {
    id:        'datafonos-cb',
    titulo:    'Punto de Reorden · Datáfonos CB',
    icono:     '📱',
    // Hoja PR-Dataf-CB del Excel
    punto:     610,      // Punto Reorden Datafonos CB (redondeado de 610.5)
    solicitud: 685,
    stockMin:  58,
    stockMax:  1210,
    categoria: 'Datáfonos',
    negocio:   'CB',
    acento:    '#99D1FC',
    formula: {
      leadTime:         4,        // meses
      colchon:          0.20,     // 20% buffer de seguridad
      aperturasMes:     200,
      datafonosPorAp:   1.2,
      demandaAperturas: 240,
      incidentesMes:    0,
      ocMes:            0,
      totalRequeridos:  288,      // con colchón
      cierresMes:       150,
      datafonosPorCierre: 1.1,
      recuperacion:     0.90,
      reusoMes:         149,      // recuperados de cierres
      datafonosNuevosMes: 91,     // nuevos necesarios (sin reuso)
      puntoFinal:       610.5,
      pedido:           685,
      componentes: [
        { label: 'Demanda aperturas / mes',   valor: 240,  unidad: 'uds' },
        { label: 'Colchón de seguridad 20%',  valor: 48,   unidad: 'uds' },
        { label: 'Recuperación de cierres',   valor: -149, unidad: 'uds' },
        { label: 'Lead Time',                 valor: 4,    unidad: 'meses' },
      ]
    }
  },
  {
    id:        'pinpads-cb',
    titulo:    'Punto de Reorden · Pinpads CB',
    icono:     '🔢',
    // Hoja PR-Pinpad-CB del Excel
    punto:     180,
    solicitud: 198,
    stockMin:  10,
    stockMax:  202,
    categoria: 'Pin pad',
    negocio:   'CB',
    acento:    '#C084FC',
    formula: {
      leadTime:         4,        // meses
      colchon:          0.20,
      aperturasMes:     200,
      pctConPinpad:     0.20,     // 20% de aperturas llevan pinpad
      demandaAperturas: 40,
      incidentesMes:    0,
      ocMes:            0,
      totalRequeridos:  48,       // con colchón
      cierresMes:       200,
      pinpadsPorCierre: 0.10,
      recuperacion:     0.90,
      reusoMes:         18,
      puntoFinal:       180,
      pedido:           198,
      componentes: [
        { label: 'Demanda aperturas / mes (20%)', valor: 40, unidad: 'uds' },
        { label: 'Colchón de seguridad 20%',      valor: 8,  unidad: 'uds' },
        { label: 'Recuperación de cierres',       valor: -18, unidad: 'uds' },
        { label: 'Lead Time',                     valor: 4,  unidad: 'meses' },
      ]
    }
  },
  {
    id:        'datafonos-vp',
    titulo:    'Punto de Reorden · Datáfonos VP',
    icono:     '💳',
    // Hoja PR-Dataf-VP del Excel
    punto:     3906,
    solicitud: 3915,
    stockMin:  132,
    stockMax:  2772,
    categoria: 'Datáfonos',
    negocio:   'VP',
    acento:    '#B0F2AE',
    formula: {
      leadTime:         4,        // meses
      colchon:          0.20,
      instalacionesMes: 500,      // proyección Wompi
      vpPorInstalacion: 1.1,
      demandaInstalaciones: 550,
      totalRequeridos:  660,      // con colchón
      cierresMes:       10,
      vpPorCierre:      0.10,
      recuperacion:     0.90,
      reusoMes:         9,
      puntoFinal:       3906,
      pedido:           3915,
      componentes: [
        { label: 'Demanda instalaciones / mes', valor: 550, unidad: 'uds' },
        { label: 'Colchón de seguridad 20%',    valor: 110, unidad: 'uds' },
        { label: 'Recuperación de cierres',     valor: -9,  unidad: 'uds' },
        { label: 'Lead Time',                   valor: 4,   unidad: 'meses' },
      ]
    }
  },
];

// ── Obtener datos crudos ──────────────────────────────────────────
function _prGetRaw() {
  if (window.INV_RAW && window.INV_RAW.length) return window.INV_RAW;
  return [];
}

// ── Calcular Stock LineaCom para una categoría/negocio dado ───────
// Equivale a UBICACIONv3 IN {"En bodega", "En distribución", "Gestor LineaCom"}
// + filtro de categoría y negocio.
function _prStockLineaCom(rows, categoria, negocio) {
  const UBICACIONES_OK = new Set(['En bodega', 'En distribución', 'Gestor LineaCom']);
  return rows
    .filter(function(r) {
      // Filtro UBICACIONv3
      var ub = (typeof invUbicacionV3 === 'function') ? invUbicacionV3(r) : _prUbicacionV3Fallback(r);
      if (!UBICACIONES_OK.has(ub)) return false;
      // Filtro categoría
      if (categoria) {
        var cat = (typeof invCategoria === 'function') ? invCategoria(r['Nombre']) : '';
        if (cat !== categoria) return false;
      }
      // Filtro negocio
      if (negocio) {
        var neg = (typeof invNegocio === 'function') ? invNegocio(r['Subtipo']) : '';
        if (neg !== negocio) return false;
      }
      return true;
    })
    .reduce(function(acc, r) { return acc + (parseInt(r['Cantidad']) || 0); }, 0);
}

// Fallback por si invUbicacionV3 aún no está disponible
function _prUbicacionV3Fallback(row) {
  var tipo = (row['Tipo de ubicación'] || '').trim();
  var cod  = (row['Código de ubicación'] || '').trim().toUpperCase();
  var pos  = (row['Posición en depósito'] || '').trim().toUpperCase();
  if (tipo === 'Site' || tipo === 'Network Element') return 'En corresponsal';
  if (tipo === 'Staff')    return 'Gestor LineaCom';
  if (tipo === 'Supplier') return 'En operador Logistico';
  if (pos === 'ENVIADO TERMINAL-TERMINAL' || pos === 'ENVIADO OPERADOR LOGISTICO') return 'En distribución';
  if (cod.startsWith('GW')) return 'Gestor Wompi';
  return 'En bodega';
}

// ── Lógica de estado ──────────────────────────────────────────────
function _prEstado(ratio) {
  if (ratio === 0)    return { emoji: '⚫', label: 'DESABASTECIDO', cls: 'pr-estado-negro' };
  if (ratio < 0.95)   return { emoji: '🔴', label: 'CRÍTICO',       cls: 'pr-estado-rojo'  };
  if (ratio < 1.3)    return { emoji: '🟡', label: 'PRECAUCIÓN',    cls: 'pr-estado-amarillo' };
  return               { emoji: '🟢', label: 'OK',           cls: 'pr-estado-verde' };
}

function _prMensaje(ratio) {
  if (ratio >= 1.5)  return 'Tranquilo, cuentas con suficiente inventario. Cuentas con al menos 1,5 veces el valor del punto de reorden.';
  if (ratio >= 1.3)  return 'Es hora de empezar a estar atento al stock. Cuentas con al menos 1,3 veces el valor del punto de reorden.';
  if (ratio >= 1.05) return 'Ya casi debes realizar pedido para reabastecer. Cuentas con al menos 1,05 veces el valor del punto de reorden.';
  if (ratio >= 0.95) return 'Alerta: solicitar reabastecimiento.';
  if (ratio >= 0.8)  return 'Alerta: sigues consumiendo el material y no ha habido abastecimiento. Estás cerca del 80% del punto de reorden.';
  if (ratio >= 0.5)  return 'Alerta: sigues consumiendo el material y no ha habido abastecimiento. Estás cerca del 50% del punto de reorden.';
  if (ratio >= 0.3)  return 'Alerta crítica: sigues consumiendo el material y no ha habido abastecimiento. Estás cerca del 30% del punto de reorden.';
  if (ratio >= 0.2)  return 'Alerta crítica: sigues consumiendo el material y no ha habido abastecimiento. Estás cerca del 20% del punto de reorden.';
  if (ratio >= 0.1)  return 'Alerta crítica: sigues consumiendo el material y no ha habido abastecimiento. Estás cerca del 10% del punto de reorden.';
  if (ratio  > 0)    return 'Alerta crítica: sigues consumiendo el material y no ha habido abastecimiento. Stock casi en cero.';
  return 'Estamos desabastecidos.';
}

// ── Formatear número ──────────────────────────────────────────────
function _prFmt(n) { return n.toLocaleString('es-CO'); }
function _prPct(n) { return (n * 100).toFixed(1) + '%'; }

// ── Render principal ──────────────────────────────────────────────
window.renderPuntosReorden = function() {
  var panel = document.getElementById('panel-puntos-reorden');
  if (!panel) return;

  var raw = _prGetRaw();
  if (!raw.length) {
    panel.innerHTML = '<div style="color:#94a3b8;padding:40px;text-align:center;font-family:\'Outfit\',sans-serif;">⏳ Cargando datos de inventario…</div>';
    // Reintentar en 800ms por si el JSON aún está cargando
    setTimeout(function() {
      if (window.INV_RAW && window.INV_RAW.length) window.renderPuntosReorden();
    }, 800);
    return;
  }

  // Calcular métricas para cada sección
  var secciones = PR_CONFIG.map(function(cfg) {
    var stock = _prStockLineaCom(raw, cfg.categoria, cfg.negocio);
    var ratio = cfg.punto > 0 ? stock / cfg.punto : 0;
    var estado  = _prEstado(ratio);
    var mensaje = _prMensaje(ratio);
    var pct     = Math.min(ratio, 2); // cap a 200% para la barra visual
    return { cfg: cfg, stock: stock, ratio: ratio, pct: pct, estado: estado, mensaje: mensaje };
  });

  // ── Resumen global ────────────────────────────────────────────
  var criticos    = secciones.filter(function(s){ return s.estado.cls === 'pr-estado-rojo' || s.estado.cls === 'pr-estado-negro'; }).length;
  var precaucion  = secciones.filter(function(s){ return s.estado.cls === 'pr-estado-amarillo'; }).length;
  var ok          = secciones.filter(function(s){ return s.estado.cls === 'pr-estado-verde'; }).length;

  var resumenColor = criticos > 0 ? '#FF5C5C' : precaucion > 0 ? '#FFC04D' : '#B0F2AE';
  var resumenEmoji = criticos > 0 ? '🔴' : precaucion > 0 ? '🟡' : '🟢';
  var resumenTexto = criticos > 0
    ? criticos + (criticos === 1 ? ' sección en estado crítico o desabastecida' : ' secciones en estado crítico o desabastecidas')
    : precaucion > 0
    ? precaucion + (precaucion === 1 ? ' sección en precaución' : ' secciones en precaución')
    : 'Todos los puntos de reorden están en niveles OK';

  // ── HTML ──────────────────────────────────────────────────────
  var html = '<div class="pr-wrap">';

  // Encabezado de página — mismo patrón que las otras tabs
  html += '<div style="margin-bottom:20px;">';
  html += '  <div class="section-label fade-up" style="color:#DFFF61;font-size:16px;margin-bottom:4px;">🎯 Puntos de Reorden</div>';
  html += '  <div style="font-size:12px;color:#64748b;margin-bottom:0;">Stock LineaCom vs. umbral de reabastecimiento&nbsp;&nbsp;·&nbsp;&nbsp;Ubicaciones:&nbsp;<span style="color:#94a3b8;">En bodega</span>&nbsp;·&nbsp;<span style="color:#94a3b8;">En distribución</span>&nbsp;·&nbsp;<span style="color:#94a3b8;">Gestor LineaCom</span></div>';
  html += '</div>';

  // Banner resumen global — patrón filters-bar del sistema
  html += '<div class="filters-bar fade-up" style="margin-bottom:24px;border-color:' + resumenColor + '33;background:' + resumenColor + '08;display:flex;align-items:center;gap:14px;padding:14px 18px;border-radius:14px;border:1px solid;">';
  html += '  <div style="font-size:26px;flex-shrink:0;">' + resumenEmoji + '</div>';
  html += '  <div style="display:flex;flex-direction:column;gap:8px;flex:1;">';
  html += '    <span style="font-family:\'Syne\',sans-serif;font-size:13px;font-weight:700;color:' + resumenColor + ';">' + resumenTexto + '</span>';
  html += '    <span style="display:flex;flex-wrap:wrap;gap:8px;">';
  html += '      <span class="pr-chip pr-chip-verde">🟢 OK: ' + ok + '</span>';
  html += '      <span class="pr-chip pr-chip-amarillo">🟡 Precaución: ' + precaucion + '</span>';
  html += '      <span class="pr-chip pr-chip-rojo">🔴 Crítico / ⚫ Desabast.: ' + criticos + '</span>';
  html += '    </span>';
  html += '  </div>';
  html += '</div>';

  // Grid de tarjetas
  html += '<div class="pr-grid">';

  secciones.forEach(function(s) {
    var cfg    = s.cfg;
    var estado = s.estado;
    var barW   = Math.min(s.pct * 50, 100); // 100% de la barra = 2× el punto reorden
    var barColor;
    if (estado.cls === 'pr-estado-verde')    barColor = '#B0F2AE';
    else if (estado.cls === 'pr-estado-amarillo') barColor = '#FFC04D';
    else if (estado.cls === 'pr-estado-rojo')  barColor = '#FF5C5C';
    else                                      barColor = '#64748b';

    // Marcadores en la barra: 100% (punto reorden) y 130% (zona OK)
    var marker100 = 50;   // 100% del punto = mitad de la barra (que va hasta 200%)
    var marker130 = 65;   // 130%

    html += '<div class="pr-card pr-card-' + estado.cls.replace('pr-estado-','') + '" style="--acento:' + cfg.acento + ';--bar-color:' + barColor + ';">';

    // ── Cabecera con fondo glassmorphism del color acento ──────────
    html += '  <div class="pr-card-header">';
    html += '    <div class="pr-card-header-left">';
    html += '      <div class="pr-card-icon-wrap"><span class="pr-card-icon">' + cfg.icono + '</span></div>';
    html += '      <div>';
    html += '        <div class="pr-card-title">' + cfg.titulo + '</div>';
    html += '        <div class="pr-card-subtitle">Umbral de reabastecimiento</div>';
    html += '      </div>';
    html += '    </div>';
    html += '    <div class="pr-estado-badge ' + estado.cls + '">' + estado.emoji + '&nbsp;' + estado.label + '</div>';
    html += '  </div>';

    // ── KPIs con glow ──────────────────────────────────────────────
    html += '  <div class="pr-kpi-row">';
    // KPI grande: Stock LineaCom (protagonista)
    html += '    <div class="pr-kpi pr-kpi-hero">';
    html += '      <div class="pr-kpi-label">STOCK LINEACOM</div>';
    html += '      <div class="pr-kpi-value pr-kpi-hero-val" style="color:' + barColor + ';text-shadow:0 0 20px ' + barColor + '55;">' + _prFmt(s.stock) + '</div>';
    html += '      <div class="pr-kpi-sub">unidades en inventario</div>';
    html += '    </div>';
    // KPI: Punto de reorden
    html += '    <div class="pr-kpi">';
    html += '      <div class="pr-kpi-label">PUNTO REORDEN</div>';
    html += '      <div class="pr-kpi-value" style="color:' + cfg.acento + ';text-shadow:0 0 16px ' + cfg.acento + '44;">' + _prFmt(cfg.punto) + '</div>';
    html += '      <div class="pr-kpi-sub">umbral mínimo</div>';
    html += '    </div>';
    // KPI: % cubierto
    html += '    <div class="pr-kpi">';
    html += '      <div class="pr-kpi-label">COBERTURA</div>';
    html += '      <div class="pr-kpi-value" style="color:' + barColor + ';">' + _prPct(s.ratio) + '</div>';
    html += '      <div class="pr-kpi-sub">vs. punto de reorden</div>';
    html += '    </div>';
    // KPI: Solicitud sugerida
    if (cfg.solicitud) {
      html += '    <div class="pr-kpi">';
      html += '      <div class="pr-kpi-label">SOLICITUD</div>';
      html += '      <div class="pr-kpi-value" style="color:#FFC04D;">' + _prFmt(cfg.solicitud) + '</div>';
      html += '      <div class="pr-kpi-sub">pedido sugerido</div>';
      html += '    </div>';
    }
    html += '  </div>';

    // ── Desglose de fórmula ────────────────────────────────────────
    if (cfg.formula) {
      var f = cfg.formula;
      html += '  <details class="pr-formula">';
      html += '    <summary class="pr-formula-summary">📐 Ver fórmula de cálculo</summary>';
      html += '    <div class="pr-formula-body">';

      // Fórmula rollos
      if (cfg.id === 'rollos') {
        html += '      <div class="pr-formula-eq">';
        html += '        <span class="pr-formula-title">Punto Reorden = (Demanda mensual × Lead Time) + Stock de Seguridad</span>';
        html += '        <div class="pr-formula-steps">';
        html += '          <div class="pr-formula-step"><span class="pr-fs-label">Demanda trimestral</span><span class="pr-fs-val">' + _prFmt(Math.round(f.demandaTrim)) + ' rollos</span></div>';
        html += '          <div class="pr-formula-step pr-formula-step-indent">';
        if (f.componentes) f.componentes.forEach(function(c) {
          html += '<div class="pr-formula-comp"><span>' + c.label + '</span><span>' + _prFmt(Math.round(c.valor)) + ' ' + c.unidad + '</span></div>';
        });
        html += '          </div>';
        html += '          <div class="pr-formula-step"><span class="pr-fs-label">Demanda mensual</span><span class="pr-fs-val">' + _prFmt(Math.round(f.demandaTrim / 3)) + ' rollos / mes</span></div>';
        html += '          <div class="pr-formula-step"><span class="pr-fs-label">Lead Time cadena</span><span class="pr-fs-val">' + f.leadTime + ' meses</span></div>';
        html += '          <div class="pr-formula-step"><span class="pr-fs-label">P. Reorden sin seguridad</span><span class="pr-fs-val">' + _prFmt(Math.round(f.puntoSinSeg)) + ' rollos</span></div>';
        html += '          <div class="pr-formula-step"><span class="pr-fs-label">Stock de seguridad (' + f.diasSeguridad + ' días)</span><span class="pr-fs-val">' + _prFmt(Math.round(f.stockSeguridad)) + ' rollos</span></div>';
        html += '          <div class="pr-formula-step pr-formula-step-total"><span>PUNTO DE REORDEN FINAL</span><span style="color:' + cfg.acento + ';">' + _prFmt(Math.round(f.puntoFinal)) + ' rollos</span></div>';
        html += '          <div class="pr-formula-step pr-formula-step-pedido"><span>Pedido sugerido (' + f.mesesAbast + ' meses)</span><span style="color:#FFC04D;">' + _prFmt(Math.round(f.pedido)) + ' rollos</span></div>';
        html += '        </div>';
        html += '      </div>';
      } else {
        // Fórmula dispositivos (CB y VP)
        html += '      <div class="pr-formula-eq">';
        html += '        <span class="pr-formula-title">Punto Reorden = (Demanda neta × Lead Time) + Stock Mínimo</span>';
        html += '        <div class="pr-formula-steps">';

        if (cfg.id === 'rollos') {
          // ya manejado arriba
        } else if (cfg.id === 'datafonos-cb' || cfg.id === 'pinpads-cb') {
          html += '          <div class="pr-formula-step"><span class="pr-fs-label">Aperturas prom. mes</span><span class="pr-fs-val">' + f.aperturasMes + ' aperturas</span></div>';
          if (cfg.id === 'pinpads-cb') {
            html += '          <div class="pr-formula-step pr-formula-step-indent"><div class="pr-formula-comp"><span>% aperturas con pinpad</span><span>' + _prPct(f.pctConPinpad) + '</span></div></div>';
          }
          html += '          <div class="pr-formula-step"><span class="pr-fs-label">Demanda aperturas / mes</span><span class="pr-fs-val">' + f.demandaAperturas + ' uds</span></div>';
          html += '          <div class="pr-formula-step"><span class="pr-fs-label">Colchón de seguridad (20%)</span><span class="pr-fs-val">→ Total requeridos: ' + f.totalRequeridos + ' uds / mes</span></div>';
          html += '          <div class="pr-formula-step"><span class="pr-fs-label">Recuperación de cierres</span><span class="pr-fs-val">− ' + f.reusoMes + ' uds / mes (90% de ' + (cfg.id === 'datafonos-cb' ? f.cierresMes + ' cierres × ' + f.datafonosPorCierre : f.cierresMes + ' cierres × ' + f.pinpadsPorCierre) + ')</span></div>';
        } else if (cfg.id === 'datafonos-vp') {
          html += '          <div class="pr-formula-step"><span class="pr-fs-label">Instalaciones prom. mes (Wompi)</span><span class="pr-fs-val">' + f.instalacionesMes + ' instalaciones</span></div>';
          html += '          <div class="pr-formula-step"><span class="pr-fs-label">VP por instalación</span><span class="pr-fs-val">× ' + f.vpPorInstalacion + '</span></div>';
          html += '          <div class="pr-formula-step"><span class="pr-fs-label">Demanda instalaciones / mes</span><span class="pr-fs-val">' + f.demandaInstalaciones + ' uds</span></div>';
          html += '          <div class="pr-formula-step"><span class="pr-fs-label">Colchón de seguridad (20%)</span><span class="pr-fs-val">→ Total requeridos: ' + f.totalRequeridos + ' uds / mes</span></div>';
          html += '          <div class="pr-formula-step"><span class="pr-fs-label">Recuperación de cierres</span><span class="pr-fs-val">− ' + f.reusoMes + ' uds / mes (' + f.recuperacion * 100 + '% × ' + f.cierresMes + ' cierres)</span></div>';
        }

        html += '          <div class="pr-formula-step"><span class="pr-fs-label">Lead Time entrega</span><span class="pr-fs-val">' + f.leadTime + ' meses</span></div>';
        html += '          <div class="pr-formula-step pr-formula-step-total"><span>PUNTO DE REORDEN FINAL</span><span style="color:' + cfg.acento + ';">' + _prFmt(Math.round(f.puntoFinal)) + ' uds</span></div>';
        html += '          <div class="pr-formula-step pr-formula-step-pedido"><span>Solicitud sugerida</span><span style="color:#FFC04D;">' + _prFmt(Math.round(f.pedido)) + ' uds</span></div>';

        // Stock Min / Max
        if (cfg.stockMin != null) {
          html += '          <div class="pr-formula-step pr-formula-step-minmax"><span>Stock Mín / Máx</span><span>' + _prFmt(cfg.stockMin) + ' / ' + (cfg.stockMax != null ? _prFmt(cfg.stockMax) : '—') + ' uds</span></div>';
        }
        html += '        </div>';
        html += '      </div>';
      }

      html += '    </div>'; // /pr-formula-body
      html += '  </details>';
    } // /if (cfg.formula)

    // ── Barra de progreso mejorada ─────────────────────────────────
    html += '  <div class="pr-bar-wrap">';
    html += '    <div class="pr-bar-header">';
    html += '      <span class="pr-bar-header-label">Stock vs. umbral</span>';
    html += '      <span class="pr-bar-header-pct" style="color:' + barColor + ';">' + _prPct(s.ratio) + '</span>';
    html += '    </div>';
    html += '    <div class="pr-bar-track">';
    html += '      <div class="pr-bar-fill" style="width:' + barW + '%;background:linear-gradient(90deg,' + barColor + 'bb,' + barColor + ');box-shadow:0 0 12px ' + barColor + '44;"></div>';
    // Marcador 100%
    html += '      <div class="pr-bar-marker" style="left:' + marker100 + '%;" title="Punto de Reorden (100%)">';
    html += '        <div class="pr-bar-marker-line"></div>';
    html += '        <div class="pr-bar-marker-label">Reorden</div>';
    html += '      </div>';
    // Marcador 130%
    html += '      <div class="pr-bar-marker" style="left:' + marker130 + '%;" title="Zona OK (130%)">';
    html += '        <div class="pr-bar-marker-line pr-bar-marker-line-ok"></div>';
    html += '        <div class="pr-bar-marker-label pr-bar-marker-label-ok">OK</div>';
    html += '      </div>';
    html += '    </div>';
    html += '    <div class="pr-bar-labels">';
    html += '      <span>0%</span><span>50%</span><span>100%</span><span>150%</span><span>200%</span>';
    html += '    </div>';
    html += '  </div>';

    // ── Mensaje de alerta ──────────────────────────────────────────
    html += '  <div class="pr-mensaje ' + estado.cls + '">';
    html += '    <span class="pr-mensaje-emoji">' + estado.emoji + '</span>';
    html += '    <span>' + s.mensaje + '</span>';
    html += '  </div>';

    html += '</div>'; // /pr-card
  });

  html += '</div>'; // /pr-grid
  html += '</div>'; // /pr-wrap

  panel.innerHTML = html;
};

// ── Estilos ───────────────────────────────────────────────────────
(function _prInjectStyles() {
  if (document.getElementById('pr-styles')) return;
  var style = document.createElement('style');
  style.id = 'pr-styles';
  style.textContent = `
    /* ── Wrapper ── */
    .pr-wrap {
      font-family: 'Outfit', 'Inter', sans-serif;
      padding: 8px 4px 48px;
      max-width: 1400px;
    }

    /* ── Encabezado página — estilos delegados a .section-label (dashboard global) ── */

    /* ── Banner global — usa .filters-bar del sistema ── */
    .pr-chip {
      font-size: 11px;
      font-weight: 600;
      padding: 3px 10px;
      border-radius: 20px;
      letter-spacing: 0.2px;
    }
    .pr-chip-verde    { background: #B0F2AE22; color: #B0F2AE; border: 1px solid #B0F2AE44; }
    .pr-chip-amarillo { background: #FFC04D22; color: #FFC04D; border: 1px solid #FFC04D44; }
    .pr-chip-rojo     { background: #FF5C5C22; color: #FF5C5C; border: 1px solid #FF5C5C44; }

    /* ── Grid ── */
    .pr-grid {
      display: grid;
      grid-template-columns: repeat(auto-fill, minmax(560px, 1fr));
      gap: 20px;
    }
    @media (max-width: 700px) {
      .pr-grid { grid-template-columns: 1fr; }
    }

    /* ── Tarjeta ── */
    .pr-card {
      background: #0f1623;
      border-radius: 16px;
      border: 1px solid rgba(255,255,255,0.07);
      padding: 20px 22px 18px;
      position: relative;
      overflow: hidden;
      transition: transform 0.2s, box-shadow 0.2s;
    }
    .pr-card::before {
      content: '';
      position: absolute;
      top: 0; left: 0; right: 0;
      height: 3px;
      background: var(--acento, #64748b);
      border-radius: 16px 16px 0 0;
    }
    .pr-card:hover {
      transform: translateY(-2px);
      box-shadow: 0 8px 32px rgba(0,0,0,0.35);
    }
    /* Tints por estado */
    .pr-card-verde    { border-color: #B0F2AE22; }
    .pr-card-amarillo { border-color: #FFC04D22; }
    .pr-card-rojo     { border-color: #FF5C5C22; }
    .pr-card-negro    { border-color: #64748b44; }

    /* ── Cabecera tarjeta ── */
    .pr-card-header {
      display: flex;
      align-items: center;
      gap: 10px;
      margin-bottom: 18px;
      flex-wrap: wrap;
    }
    .pr-card-icon { font-size: 20px; flex-shrink: 0; }
    .pr-card-title {
      font-family: 'Syne', sans-serif;
      font-size: 14px;
      font-weight: 700;
      color: #e2e8f0;
      flex: 1;
    }

    /* ── Badge de estado ── */
    .pr-estado-badge {
      font-size: 11px;
      font-weight: 700;
      padding: 4px 12px;
      border-radius: 20px;
      letter-spacing: 0.5px;
      text-transform: uppercase;
      flex-shrink: 0;
    }
    .pr-estado-verde    { background: #B0F2AE1a; color: #B0F2AE; border: 1px solid #B0F2AE44; }
    .pr-estado-amarillo { background: #FFC04D1a; color: #FFC04D; border: 1px solid #FFC04D44; }
    .pr-estado-rojo     { background: #FF5C5C1a; color: #FF5C5C; border: 1px solid #FF5C5C44; }
    .pr-estado-negro    { background: #64748b1a; color: #94a3b8;  border: 1px solid #64748b44; }

    /* ── Fila de KPIs ── */
    .pr-kpi-row {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(100px, 1fr));
      gap: 12px;
      margin-bottom: 18px;
    }
    .pr-kpi {
      background: rgba(255,255,255,0.04);
      border-radius: 10px;
      padding: 10px 12px 8px;
      border: 1px solid rgba(255,255,255,0.06);
    }
    .pr-kpi-label {
      font-size: 9px;
      font-weight: 700;
      letter-spacing: 0.8px;
      text-transform: uppercase;
      color: #64748b;
      margin-bottom: 5px;
    }
    .pr-kpi-value {
      font-family: 'JetBrains Mono', monospace;
      font-size: 18px;
      font-weight: 700;
      color: #f1f5f9;
      line-height: 1;
    }
    .pr-kpi-alerta { font-family: inherit !important; }

    /* ── Barra de progreso ── */
    .pr-bar-wrap { margin-bottom: 14px; }
    .pr-bar-track {
      position: relative;
      height: 14px;
      background: rgba(255,255,255,0.07);
      border-radius: 8px;
      overflow: visible;
      margin-bottom: 4px;
    }
    .pr-bar-fill {
      height: 100%;
      border-radius: 8px;
      transition: width 0.8s cubic-bezier(.4,0,.2,1);
      position: relative;
    }
    .pr-bar-fill::after {
      content: '';
      position: absolute;
      top: 0; left: 0; right: 0; bottom: 0;
      background: linear-gradient(90deg, transparent 0%, rgba(255,255,255,0.25) 100%);
      border-radius: 8px;
    }
    .pr-bar-marker {
      position: absolute;
      top: -4px;
      transform: translateX(-50%);
      display: flex;
      flex-direction: column;
      align-items: center;
      pointer-events: none;
      z-index: 2;
    }
    .pr-bar-marker-line {
      width: 2px;
      height: 22px;
      background: rgba(255,255,255,0.35);
      border-radius: 1px;
    }
    .pr-bar-marker-line-ok {
      background: #B0F2AE77;
    }
    .pr-bar-marker-label {
      font-size: 9px;
      color: #64748b;
      white-space: nowrap;
      margin-top: 3px;
      font-weight: 600;
    }
    .pr-bar-marker-label-ok { color: #B0F2AE99; }
    .pr-bar-labels {
      display: flex;
      justify-content: space-between;
      font-size: 9px;
      color: #475569;
      padding: 0 1px;
      margin-top: 22px;
    }
    .pr-bar-labels em { font-style: normal; color: #64748b; }

    /* ── Mensaje ── */
    .pr-mensaje {
      display: flex;
      align-items: flex-start;
      gap: 10px;
      padding: 10px 14px;
      border-radius: 10px;
      font-size: 12px;
      line-height: 1.5;
      margin-top: 2px;
    }
    .pr-mensaje-emoji { font-size: 16px; flex-shrink: 0; margin-top: 1px; }
    .pr-mensaje.pr-estado-verde    { background: #B0F2AE0d; color: #9de09a; }
    .pr-mensaje.pr-estado-amarillo { background: #FFC04D0d; color: #d4a03f; }
    .pr-mensaje.pr-estado-rojo     { background: #FF5C5C0d; color: #e07070; }
    .pr-mensaje.pr-estado-negro    { background: #64748b0d; color: #94a3b8; }

    /* ── Fórmula desplegable ── */
    .pr-formula {
      margin-top: 12px;
      border-radius: 10px;
      background: rgba(255,255,255,0.025);
      border: 1px solid rgba(255,255,255,0.06);
      overflow: hidden;
    }
    .pr-formula-summary {
      cursor: pointer;
      padding: 9px 14px;
      font-size: 11px;
      font-weight: 600;
      color: #64748b;
      font-family: 'Outfit', sans-serif;
      letter-spacing: 0.3px;
      list-style: none;
      user-select: none;
      transition: color 0.2s, background 0.2s;
    }
    .pr-formula-summary::-webkit-details-marker { display: none; }
    .pr-formula[open] .pr-formula-summary {
      color: #94a3b8;
      background: rgba(255,255,255,0.03);
      border-bottom: 1px solid rgba(255,255,255,0.06);
    }
    .pr-formula-body {
      padding: 14px 16px 16px;
    }
    .pr-formula-title {
      display: block;
      font-family: 'JetBrains Mono', monospace;
      font-size: 10px;
      color: #475569;
      margin-bottom: 12px;
      letter-spacing: 0.2px;
      font-style: italic;
    }
    .pr-formula-steps {
      display: flex;
      flex-direction: column;
      gap: 5px;
    }
    .pr-formula-step {
      display: flex;
      justify-content: space-between;
      align-items: center;
      font-size: 11px;
      padding: 5px 8px;
      border-radius: 6px;
      gap: 10px;
    }
    .pr-formula-step:hover { background: rgba(255,255,255,0.03); }
    .pr-fs-label { color: #64748b; }
    .pr-fs-val {
      font-family: 'JetBrains Mono', monospace;
      color: #94a3b8;
      font-size: 11px;
      text-align: right;
      flex-shrink: 0;
    }
    .pr-formula-step-indent {
      flex-direction: column;
      align-items: stretch;
      padding-left: 16px;
      gap: 2px;
    }
    .pr-formula-comp {
      display: flex;
      justify-content: space-between;
      font-size: 10.5px;
      color: #475569;
      padding: 2px 4px;
    }
    .pr-formula-comp span:last-child {
      font-family: 'JetBrains Mono', monospace;
      color: #64748b;
    }
    .pr-formula-step-total {
      margin-top: 4px;
      border-top: 1px solid rgba(255,255,255,0.07);
      padding-top: 8px;
      font-weight: 700;
      font-size: 12px;
      color: #e2e8f0;
    }
    .pr-formula-step-pedido {
      font-size: 11px;
      color: #94a3b8;
    }
    .pr-formula-step-minmax {
      font-size: 10px;
      color: #475569;
      border-top: 1px dashed rgba(255,255,255,0.05);
      margin-top: 2px;
      padding-top: 6px;
    }
  `;
  document.head.appendChild(style);
})();