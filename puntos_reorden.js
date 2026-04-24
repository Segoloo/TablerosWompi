'use strict';
/**
 * ╔══════════════════════════════════════════════════════════════════╗
 * ║  puntos_reorden.js — Tab "Puntos de Reorden"                    ║
 * ║  Rollos · Datáfonos CB · Pinpads CB · Datáfonos VP             ║
 * ╚══════════════════════════════════════════════════════════════════╝
 */

// ── Configuración de puntos de reorden ───────────────────────────
const PR_CONFIG = [
  {
    id:        'rollos',
    titulo:    'Punto de Reorden · Rollos',
    icono:     '🎞️',
    punto:     400000,
    categoria: 'Rollos',
    negocio:   null,          // CB + VP
    acento:    '#DFFF61',
  },
  {
    id:        'datafonos-cb',
    titulo:    'Punto de Reorden · Datáfonos CB',
    icono:     '📱',
    punto:     2322,
    categoria: 'Datáfonos',
    negocio:   'CB',
    acento:    '#99D1FC',
  },
  {
    id:        'pinpads-cb',
    titulo:    'Punto de Reorden · Pinpads CB',
    icono:     '🔢',
    punto:     455,
    categoria: 'Pin pad',
    negocio:   'CB',
    acento:    '#C084FC',
  },
  {
    id:        'datafonos-vp',
    titulo:    'Punto de Reorden · Datáfonos VP',
    icono:     '💳',
    punto:     4500,
    categoria: 'Datáfonos',
    negocio:   'VP',
    acento:    '#B0F2AE',
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
    html += '  </div>';

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
      grid-template-columns: repeat(4, 1fr);
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
  `;
  document.head.appendChild(style);
})();