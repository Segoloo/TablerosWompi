/**
 * ╔══════════════════════════════════════════════════════════════════╗
 * ║  consulta_masiva_rollos.js — Consulta Masiva por Código Punto   ║
 * ║                                                                  ║
 * ║  Permite subir un Excel con códigos de punto y obtener toda     ║
 * ║  la información de rollos disponible en TABLERO_ROLLOS_FILAS    ║
 * ║  y RCV2_SITIOS para cada código consultado.                     ║
 * ║                                                                  ║
 * ║  Funcionalidades:                                                ║
 * ║  · Upload Excel / CSV con columna de códigos de punto           ║
 * ║  · Descargar plantilla Excel de demostración                    ║
 * ║  · KPIs agregados del set consultado                            ║
 * ║  · Tabla resumen con todos los indicadores de rollos            ║
 * ║  · Detalle de MOs (movimientos) por punto                       ║
 * ║  · Conteo de rollos por corresponsal                            ║
 * ║  · Export Excel completo del reporte                            ║
 * ╚══════════════════════════════════════════════════════════════════╝
 */

'use strict';

// ── Estado del módulo ──────────────────────────────────────────────
let CMR_RESULTS   = [];   // resultados procesados de la consulta
let CMR_CODES     = [];   // códigos ingresados por el usuario
let cmrPage       = 1;
let cmrDetailPage = {};   // paginación por código en detalle de MOs
const CMR_PAGE_SIZE = 25;
const CMR_MO_PAGE_SIZE = 10;

// ── Helpers ────────────────────────────────────────────────────────
const cmrFmt  = n => (parseFloat(n) || 0).toLocaleString('es-CO', { maximumFractionDigits: 1 });
const cmrFmtI = n => Math.round(parseFloat(n) || 0).toLocaleString('es-CO');
const cmrP    = n => parseFloat(n) || 0;
const cmrDate = s => { if (!s) return null; const d = new Date(s); return isNaN(d) ? null : d; };
const cmrFmtDate = s => { const d = cmrDate(s); if (!d) return '—'; return d.toLocaleDateString('es-CO', { day:'2-digit', month:'short', year:'numeric' }); };
const cmrSem  = m => {
  if (m >= 3) return { col:'#B0F2AE', label:'OK',      emoji:'🟢', bg:'rgba(176,242,174,.1)' };
  if (m >= 2) return { col:'#DFFF61', label:'ATENCIÓN', emoji:'🟡', bg:'rgba(223,255,97,.1)' };
  if (m >= 1) return { col:'#FFC04D', label:'ALERTA',   emoji:'🟠', bg:'rgba(255,192,77,.1)' };
  return             { col:'#FF5C5C', label:'CRÍTICO',  emoji:'🔴', bg:'rgba(255,92,92,.1)' };
};

// ──────────────────────────────────────────────────────────────────
//  PLANTILLA DEMO
// ──────────────────────────────────────────────────────────────────
window.cmrDescargarPlantilla = function () {
  if (!window.XLSX) { alert('Librería XLSX no disponible'); return; }

  const ejemplos = [
    { 'CODIGO_PUNTO': 'SITIO001', 'DESCRIPCION (opcional)': 'Ejemplo corresponsal 1' },
    { 'CODIGO_PUNTO': 'SITIO002', 'DESCRIPCION (opcional)': 'Ejemplo corresponsal 2' },
    { 'CODIGO_PUNTO': 'SITIO003', 'DESCRIPCION (opcional)': 'Ejemplo corresponsal 3' },
    { 'CODIGO_PUNTO': 'SITIO004', 'DESCRIPCION (opcional)': '' },
    { 'CODIGO_PUNTO': 'SITIO005', 'DESCRIPCION (opcional)': '' },
  ];

  // Agregar algunos códigos reales si están disponibles
  const sitios = window.RCV2_SITIOS;
  if (sitios && sitios.size > 0) {
    ejemplos.splice(0, ejemplos.length); // limpiar ejemplos ficticios
    let i = 0;
    for (const [key] of sitios) {
      ejemplos.push({ 'CODIGO_PUNTO': key, 'DESCRIPCION (opcional)': '' });
      if (++i >= 5) break;
    }
  }

  const ws = window.XLSX.utils.json_to_sheet(ejemplos);
  // Ancho de columnas
  ws['!cols'] = [{ wch: 22 }, { wch: 35 }];
  const wb = window.XLSX.utils.book_new();
  window.XLSX.utils.book_append_sheet(wb, ws, 'Consulta Masiva');

  // Hoja de instrucciones
  const instrRows = [
    { 'INSTRUCCIONES': '═══════════ INSTRUCCIONES ═══════════' },
    { 'INSTRUCCIONES': '1. Coloca los códigos de punto en la columna CODIGO_PUNTO (hoja anterior).' },
    { 'INSTRUCCIONES': '2. La columna DESCRIPCION es opcional — puedes dejarla vacía.' },
    { 'INSTRUCCIONES': '3. Acepta códigos cal_codigo_sitio o codigo_ubicacion_destino.' },
    { 'INSTRUCCIONES': '4. El archivo puede tener .xlsx, .xls o .csv (una columna).' },
    { 'INSTRUCCIONES': '5. No incluyas encabezados adicionales; la primera fila es el header.' },
    { 'INSTRUCCIONES': '6. Máximo 500 códigos por consulta.' },
  ];
  const wsI = window.XLSX.utils.json_to_sheet(instrRows);
  wsI['!cols'] = [{ wch: 60 }];
  window.XLSX.utils.book_append_sheet(wb, wsI, 'Instrucciones');

  window.XLSX.writeFile(wb, 'plantilla_consulta_masiva_rollos.xlsx');
};

// ──────────────────────────────────────────────────────────────────
//  PROCESAR ARCHIVO SUBIDO
// ──────────────────────────────────────────────────────────────────
window.cmrHandleFile = function (input) {
  const file = input.files[0];
  if (!file) return;

  const statusEl = document.getElementById('cmr-file-status');
  if (statusEl) statusEl.textContent = `Procesando: ${file.name}…`;

  const ext = file.name.split('.').pop().toLowerCase();

  if (ext === 'csv') {
    const reader = new FileReader();
    reader.onload = e => {
      const lines = e.target.result.split(/\r?\n/).filter(l => l.trim());
      // Detectar si hay header
      const first = lines[0].trim().toUpperCase();
      const startIdx = (first.includes('CODIGO') || first.includes('PUNTO') || first === 'ID') ? 1 : 0;
      const codes = lines.slice(startIdx).map(l => l.split(/[,;|\t]/)[0].trim().toUpperCase()).filter(c => c);
      _cmrRunQuery(codes, file.name);
    };
    reader.readAsText(file);
  } else {
    // Excel via XLSX
    if (!window.XLSX) { alert('Librería XLSX no cargada'); return; }
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb = window.XLSX.read(e.target.result, { type: 'binary' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = window.XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
        if (!rows.length) { alert('El archivo está vacío.'); return; }

        // Detectar columna CODIGO_PUNTO (o la primera columna)
        let colIdx = 0;
        const header = rows[0];
        if (Array.isArray(header)) {
          const ci = header.findIndex(h => String(h).toUpperCase().includes('CODIGO') || String(h).toUpperCase().includes('PUNTO') || String(h).toUpperCase() === 'ID');
          if (ci >= 0) colIdx = ci;
        }
        const startRow = (typeof header[colIdx] === 'string' && isNaN(header[colIdx])) ? 1 : 0;
        const codes = rows.slice(startRow).map(r => String(r[colIdx] || '').trim().toUpperCase()).filter(c => c);
        _cmrRunQuery(codes, file.name);
      } catch(err) {
        alert('Error leyendo el archivo: ' + err.message);
      }
    };
    reader.readAsBinaryString(file);
  }
};

// ──────────────────────────────────────────────────────────────────
//  EJECUTAR CONSULTA
// ──────────────────────────────────────────────────────────────────
function _cmrRunQuery(codes, fileName) {
  if (!codes.length) {
    alert('No se encontraron códigos en el archivo.');
    return;
  }
  if (codes.length > 500) {
    if (!confirm(`Se encontraron ${codes.length} códigos. Se procesarán los primeros 500. ¿Continuar?`)) return;
    codes = codes.slice(0, 500);
  }

  const statusEl = document.getElementById('cmr-file-status');
  if (statusEl) statusEl.innerHTML = `<span style="color:#B0F2AE;">✓ ${fileName}</span> · <strong style="color:#DFFF61;">${codes.length}</strong> códigos leídos`;

  CMR_CODES = [...new Set(codes)]; // deduplicar
  cmrPage = 1;
  cmrDetailPage = {};
  CMR_RESULTS = CMR_CODES.map(code => _cmrBuildRecord(code));

  _cmrRenderAll();
}

// ──────────────────────────────────────────────────────────────────
//  CONSTRUIR REGISTRO POR CÓDIGO
// ──────────────────────────────────────────────────────────────────
function _cmrBuildRecord(code) {
  const rec = {
    code,
    found: false,
    meta: {},
    cal: {},
    movs: [],
    inv_stock: 0,
    inv_refs: [],
    // KPIs calculados
    saldo_meses: 0,
    saldo_meses_str: '0.0',
    sem: cmrSem(0),
    prom_mensual: 0,
    prom_diario: 0,
    prom_semanal: 0,
    saldo_rollos: 0,
    saldo_dias: 0,
    punto_reorden: 0,
    sla_cumple: false,
    bajo_reorden: false,
    ind_riesgo: '',
    rotacion: 0,
    rollos_entregados: 0,
    rollos_consumidos: 0,
    fecha_apertura: '',
    fecha_abst: '',
    // MOs únicas
    mos_unicas: new Set(),
    // Rollos por corresponsal (MO)
    rollos_por_mo: new Map(),
    // Inventario (stock_wompi)
    inv_por_tipo: new Map(),
  };

  // ── Buscar en RCV2_SITIOS (índice construido por rollos_comercio_v2) ──
  const sitios = window.RCV2_SITIOS;
  let sitio = null;
  if (sitios) {
    sitio = sitios.get(code);
    if (!sitio) {
      // Búsqueda case-insensitive y por nombre
      for (const [k, v] of sitios) {
        if (k.toUpperCase() === code) { sitio = v; break; }
        if (v.meta && v.meta.nombre_sitio && v.meta.nombre_sitio.toUpperCase() === code) { sitio = v; break; }
        // También buscar por codigo_ubic_destino
        if (v.meta && v.meta.codigo_ubic_destino && v.meta.codigo_ubic_destino.toUpperCase() === code) { sitio = v; break; }
      }
    }
  }

  if (sitio) {
    rec.found = true;
    rec.meta  = sitio.meta || {};
    const cal  = sitio.cal || {};
    rec.cal    = cal;

    rec.saldo_rollos    = cmrP(cal.saldo_rollos);
    rec.saldo_dias      = cmrP(cal.saldo_dias);
    rec.saldo_meses     = rec.saldo_dias / 30;
    rec.saldo_meses_str = rec.saldo_meses.toFixed(1);
    rec.sem             = cmrSem(rec.saldo_meses);
    rec.prom_mensual    = cmrP(cal.prom_mensual);
    rec.prom_diario     = rec.prom_mensual / 30;
    rec.prom_semanal    = rec.prom_diario * 7;
    rec.punto_reorden   = cmrP(cal.punto_reorden);
    rec.sla_cumple      = rec.saldo_meses >= 3;
    rec.bajo_reorden    = rec.saldo_rollos > 0 && rec.saldo_rollos <= rec.punto_reorden && rec.punto_reorden > 0;
    rec.rollos_entregados = cmrP(cal.rollos_entregados);
    rec.rollos_consumidos = cmrP(cal.rollos_consumidos);
    rec.fecha_apertura  = cal.fecha_apertura || '';
    rec.fecha_abst      = cal.fecha_abst || '';
    rec.rotacion        = rec.saldo_rollos > 0 && rec.prom_mensual > 0
      ? (rec.rollos_consumidos / rec.saldo_rollos)
      : 0;
    rec.ind_riesgo      = cal.ind_riesgo || '';

    // Inventario desde sitio
    rec.inv_stock = sitio.inv_stock || 0;
    if (sitio.inv_refs) {
      rec.inv_refs = [...sitio.inv_refs.entries()].sort((a,b) => b[1]-a[1]);
    }

    // Movimientos (MOs)
    rec.movs = sitio.movs || [];
    rec.mos_unicas = new Set(rec.movs.map(m => m.tarea).filter(Boolean));

    // Agrupar rollos por MO
    const rPorMO = new Map();
    rec.movs.forEach(m => {
      if (!m.tarea) return;
      const prev = rPorMO.get(m.tarea) || { tarea: m.tarea, total: 0, movs: [] };
      prev.total += m.cantidad || 0;
      prev.movs.push(m);
      rPorMO.set(m.tarea, prev);
    });
    rec.rollos_por_mo = rPorMO;

    // Agrupar inventario por tipo de ubicación desde INV_RAW
    const invRaw = window.INV_RAW || [];
    invRaw.forEach(r => {
      const cat = window.invCategoria ? window.invCategoria(r['Nombre']) : '';
      if (cat !== 'Rollos') return;
      const codUbic = (r['Código de ubicación'] || '').trim().toUpperCase();
      const nomUbic = (r['Nombre de la ubicación'] || '').trim().toUpperCase();
      const matches = codUbic === code || nomUbic === (rec.meta.nombre_sitio || '').toUpperCase();
      if (!matches) return;
      const tipo = (r['Tipo de ubicación'] || 'Sin tipo').trim();
      const qty  = parseInt(r['Cantidad']) || 0;
      const ref  = (r['Nombre'] || 'Sin ref').trim();
      const prev = rec.inv_por_tipo.get(tipo) || { total: 0, items: [] };
      prev.total += qty;
      prev.items.push({ ref, qty, cod: r['Código de ubicación'] || '' });
      rec.inv_por_tipo.set(tipo, prev);
    });

  } else {
    // Intentar al menos buscar en TABLERO_ROLLOS_FILAS directo
    const filas = window.TABLERO_ROLLOS_FILAS || [];
    const match = filas.find(f => {
      const cs = (f.cal_codigo_sitio || f.codigo_ubicacion_destino || '').trim().toUpperCase();
      const cm = (f.cal_codigo_mo || '').trim().toUpperCase();
      return cs === code || cm === code;
    });
    if (match) {
      rec.found = true;
      rec.meta = {
        nombre_sitio: match.nombre_ubicacion_destino || match.cal_nombre_sitio || code,
        departamento: match.departamento || '',
        ciudad: match.Ciudad || match.ciudad || '',
        proyecto: match.proyecto || '',
      };
    }
  }

  return rec;
}

// ──────────────────────────────────────────────────────────────────
//  RENDER COMPLETO
// ──────────────────────────────────────────────────────────────────
function _cmrRenderAll() {
  _cmrRenderKPIs();
  _cmrRenderTable();
  // Mostrar sección de resultados
  const resultSection = document.getElementById('cmr-results-section');
  if (resultSection) resultSection.style.display = 'block';

  // Scroll suave a resultados
  setTimeout(() => {
    resultSection?.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }, 120);
}

// ──────────────────────────────────────────────────────────────────
//  KPIs AGREGADOS
// ──────────────────────────────────────────────────────────────────
function _cmrRenderKPIs() {
  const el = document.getElementById('cmr-kpi-strip');
  if (!el) return;

  const all     = CMR_RESULTS;
  const found   = all.filter(r => r.found);
  const notFound = all.filter(r => !r.found);
  const total   = all.length;

  const totalRollos   = found.reduce((s, r) => s + r.saldo_rollos, 0);
  const totalInvStock = found.reduce((s, r) => s + r.inv_stock, 0);
  const avgMeses      = found.length > 0 ? found.reduce((s, r) => s + r.saldo_meses, 0) / found.length : 0;
  const slaOk         = found.filter(r => r.sla_cumple).length;
  const pctSla        = found.length > 0 ? Math.round(slaOk / found.length * 100) : 0;
  const criticos      = found.filter(r => r.saldo_meses < 1).length;
  const bajoPR        = found.filter(r => r.bajo_reorden).length;
  const totalMOs      = found.reduce((s, r) => s + r.mos_unicas.size, 0);
  const totalConsumoMes = found.reduce((s, r) => s + r.prom_mensual, 0);

  const kpis = [
    { label: 'Puntos Consultados',   value: total.toLocaleString('es-CO'),            icon: '🔍', color: '#99D1FC',  bg: 'rgba(153,209,252,.1)' },
    { label: 'Puntos Encontrados',   value: `${found.length}/${total}`,               icon: '✅', color: '#B0F2AE',  bg: 'rgba(176,242,174,.1)' },
    { label: 'No Encontrados',       value: notFound.length.toLocaleString('es-CO'),  icon: '❓', color: notFound.length > 0 ? '#FFC04D' : '#B0F2AE', bg: 'rgba(255,192,77,.08)' },
    { label: 'Saldo Total Rollos',   value: cmrFmtI(totalRollos),                     icon: '📦', color: '#B0F2AE',  bg: 'rgba(176,242,174,.08)' },
    { label: 'Stock Inventario',     value: cmrFmtI(totalInvStock),                   icon: '🏭', color: '#C084FC',  bg: 'rgba(192,132,252,.08)' },
    { label: 'Cobertura Promedio',   value: avgMeses.toFixed(1) + ' meses',           icon: '📅', color: '#DFFF61',  bg: 'rgba(223,255,97,.08)' },
    { label: 'Consumo Mensual Total',value: cmrFmtI(totalConsumoMes) + ' rollos',    icon: '📊', color: '#F49D6E',  bg: 'rgba(244,157,110,.08)' },
    { label: 'SLA ≥3 Meses',         value: pctSla + '%', sub: `${slaOk}/${found.length}`, icon: '🎯', color: pctSla >= 70 ? '#B0F2AE' : pctSla >= 50 ? '#DFFF61' : '#FF5C5C', pct: pctSla, bg: 'rgba(176,242,174,.06)' },
    { label: 'Críticos <1 Mes',      value: criticos.toLocaleString('es-CO'),         icon: '🚨', color: criticos > 0 ? '#FF5C5C' : '#B0F2AE', bg: 'rgba(255,92,92,.08)' },
    { label: 'Bajo Punto Reorden',   value: bajoPR.toLocaleString('es-CO'),           icon: '⚠️', color: bajoPR > 0 ? '#FFC04D' : '#B0F2AE', bg: 'rgba(255,192,77,.08)' },
    { label: 'Total MOs Asociadas',  value: totalMOs.toLocaleString('es-CO'),         icon: '📋', color: '#7B8CDE',  bg: 'rgba(123,140,222,.08)' },
  ];

  el.innerHTML = kpis.map(k => `
    <div style="background:${k.bg};border:1px solid ${k.color}22;border-radius:14px;padding:14px 16px;min-width:140px;flex:1;position:relative;overflow:hidden;transition:transform .2s,box-shadow .2s;"
         onmouseover="this.style.transform='translateY(-2px)';this.style.boxShadow='0 8px 24px rgba(0,0,0,.4)'"
         onmouseout="this.style.transform='';this.style.boxShadow=''">
      <div style="font-size:18px;margin-bottom:5px;">${k.icon}</div>
      <div style="font-family:'JetBrains Mono',monospace;font-size:18px;font-weight:700;color:${k.color};line-height:1;">${k.value}</div>
      ${k.sub ? `<div style="font-size:10px;color:#64748b;margin-top:2px;">${k.sub}</div>` : ''}
      ${k.pct !== undefined ? `<div style="margin-top:7px;background:rgba(255,255,255,.06);border-radius:4px;height:3px;overflow:hidden;"><div style="width:${k.pct}%;height:100%;background:${k.color};border-radius:4px;transition:width .6s ease;"></div></div>` : ''}
      <div style="font-size:10px;color:#7A7674;margin-top:5px;font-family:'Outfit',sans-serif;">${k.label}</div>
    </div>`).join('');
}

// ──────────────────────────────────────────────────────────────────
//  TABLA DE RESULTADOS
// ──────────────────────────────────────────────────────────────────
function _cmrRenderTable() {
  const wrap = document.getElementById('cmr-table-wrap');
  const pagEl = document.getElementById('cmr-table-pagination');
  const countEl = document.getElementById('cmr-table-count');
  if (!wrap) return;

  const total = CMR_RESULTS.length;
  const totalPages = Math.max(1, Math.ceil(total / CMR_PAGE_SIZE));
  if (cmrPage > totalPages) cmrPage = 1;
  const start = (cmrPage - 1) * CMR_PAGE_SIZE;
  const slice = CMR_RESULTS.slice(start, start + CMR_PAGE_SIZE);

  if (countEl) countEl.textContent = `${total} puntos consultados · ${CMR_RESULTS.filter(r => r.found).length} encontrados`;

  const thStyle = 'padding:10px 14px;text-align:left;color:#64748b;font-weight:700;font-size:10px;letter-spacing:.6px;white-space:nowrap;border-bottom:1px solid rgba(255,255,255,.07);';
  const tdStyle = (col) => `padding:9px 14px;color:${col||'#e2e8f0'};font-family:'Outfit',sans-serif;font-size:12px;vertical-align:top;border-bottom:1px solid rgba(255,255,255,.04);`;

  wrap.innerHTML = `<div style="overflow-x:auto;">
  <table style="width:100%;border-collapse:collapse;font-size:12px;">
    <thead style="position:sticky;top:0;z-index:2;">
      <tr style="background:rgba(0,0,0,.55);backdrop-filter:blur(4px);">
        <th style="${thStyle}">#</th>
        <th style="${thStyle}">CÓDIGO PUNTO</th>
        <th style="${thStyle}">NOMBRE CORRESPONSAL</th>
        <th style="${thStyle}">CIUDAD</th>
        <th style="${thStyle}">ESTADO</th>
        <th style="${thStyle}">SALDO ROLLOS</th>
        <th style="${thStyle}">STOCK INV.</th>
        <th style="${thStyle}">COBERTURA</th>
        <th style="${thStyle}">CONS. MENSUAL</th>
        <th style="${thStyle}">CONS. DIARIO</th>
        <th style="${thStyle}">PTO. REORDEN</th>
        <th style="${thStyle}">SLA 3M</th>
        <th style="${thStyle}">RIESGO</th>
        <th style="${thStyle}">MOs</th>
        <th style="${thStyle}">ENT. / CONS.</th>
        <th style="${thStyle}">F. APERTURA</th>
        <th style="${thStyle}">F. ABAST.</th>
        <th style="${thStyle}">DETALLE</th>
      </tr>
    </thead>
    <tbody>
      ${slice.map((r, i) => {
        const idx = start + i + 1;
        if (!r.found) {
          return `<tr style="background:rgba(255,92,92,.04);">
            <td style="${tdStyle('#475569')}">${idx}</td>
            <td style="${tdStyle('#FF5C5C')};font-family:'JetBrains Mono',monospace;font-size:11px;font-weight:600;">${r.code}</td>
            <td colspan="16" style="${tdStyle('#64748b')}"><span style="background:rgba(255,92,92,.12);color:#FF5C5C;border-radius:6px;padding:2px 10px;font-size:11px;">❌ No encontrado en el sistema</span></td>
          </tr>`;
        }
        const sg = r.sem;
        const bgRow = i % 2 === 0 ? 'transparent' : 'rgba(255,255,255,.012)';
        const slaCol = r.sla_cumple ? '#B0F2AE' : '#FF5C5C';
        const riesgoCol = r.ind_riesgo === 'ALTO' ? '#FF5C5C' : r.ind_riesgo === 'CRÍTICO' ? '#FF0000' : r.ind_riesgo ? '#FFC04D' : '#64748b';
        return `<tr style="background:${bgRow};transition:background .12s;"
                    onmouseover="this.style.background='rgba(176,242,174,.04)'"
                    onmouseout="this.style.background='${bgRow}'">
          <td style="${tdStyle('#475569')};font-size:10px;">${idx}</td>
          <td style="${tdStyle('#DFFF61')};font-family:'JetBrains Mono',monospace;font-size:11px;font-weight:700;">${r.code}</td>
          <td style="${tdStyle('#f1f5f9')};font-weight:600;max-width:180px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;" title="${(r.meta.nombre_sitio||'').replace(/"/g,'&quot;')}">${r.meta.nombre_sitio || '—'}</td>
          <td style="${tdStyle('#94a3b8')};white-space:nowrap;">${r.meta.ciudad || '—'}</td>
          <td style="padding:9px 14px;vertical-align:top;border-bottom:1px solid rgba(255,255,255,.04);">
            <span style="background:${sg.bg};color:${sg.col};border-radius:8px;padding:2px 9px;font-size:10px;font-weight:700;white-space:nowrap;">${sg.emoji} ${sg.label}</span>
          </td>
          <td style="${tdStyle(sg.col)};font-family:'JetBrains Mono',monospace;font-weight:700;">${cmrFmtI(r.saldo_rollos)}</td>
          <td style="${tdStyle('#C084FC')};font-family:'JetBrains Mono',monospace;">${cmrFmtI(r.inv_stock) || '—'}</td>
          <td style="${tdStyle(sg.col)};font-family:'JetBrains Mono',monospace;font-weight:700;">${r.saldo_meses_str}m</td>
          <td style="${tdStyle('#99D1FC')};font-family:'JetBrains Mono',monospace;">${cmrFmtI(r.prom_mensual)}</td>
          <td style="${tdStyle('#7B8CDE')};font-family:'JetBrains Mono',monospace;">${r.prom_diario.toFixed(1)}</td>
          <td style="${tdStyle('#FFC04D')};font-family:'JetBrains Mono',monospace;">${cmrFmtI(r.punto_reorden) || '—'}${r.bajo_reorden ? ' ⚠️' : ''}</td>
          <td style="padding:9px 14px;vertical-align:top;border-bottom:1px solid rgba(255,255,255,.04);">
            <span style="color:${slaCol};font-size:11px;font-weight:700;">${r.sla_cumple ? '✓ Sí' : '✗ No'}</span>
          </td>
          <td style="padding:9px 14px;vertical-align:top;border-bottom:1px solid rgba(255,255,255,.04);">
            <span style="color:${riesgoCol};font-size:10px;font-weight:700;">${r.ind_riesgo || '—'}</span>
          </td>
          <td style="${tdStyle('#B0F2AE')};font-family:'JetBrains Mono',monospace;">${r.mos_unicas.size}</td>
          <td style="${tdStyle('#94a3b8')};font-family:'JetBrains Mono',monospace;font-size:11px;">
            <span style="color:#B0F2AE;">${cmrFmtI(r.rollos_entregados)}</span> / <span style="color:#F49D6E;">${cmrFmtI(r.rollos_consumidos)}</span>
          </td>
          <td style="${tdStyle('#64748b')};white-space:nowrap;font-size:11px;">${cmrFmtDate(r.fecha_apertura)}</td>
          <td style="${tdStyle('#64748b')};white-space:nowrap;font-size:11px;">${cmrFmtDate(r.fecha_abst)}</td>
          <td style="padding:9px 14px;vertical-align:top;border-bottom:1px solid rgba(255,255,255,.04);">
            <button onclick="window.cmrToggleDetail('${r.code}')"
              style="padding:4px 12px;border-radius:8px;border:1px solid rgba(176,242,174,.3);background:rgba(176,242,174,.07);color:#B0F2AE;font-size:11px;cursor:pointer;font-family:'Outfit',sans-serif;transition:all .15s;"
              onmouseover="this.style.background='rgba(176,242,174,.18)'"
              onmouseout="this.style.background='rgba(176,242,174,.07)'">
              Ver MOs
            </button>
          </td>
        </tr>
        <tr id="cmr-detail-row-${r.code}" style="display:none;">
          <td colspan="18" style="padding:0;border-bottom:2px solid rgba(176,242,174,.15);">
            <div id="cmr-detail-panel-${r.code}" style="padding:20px 24px 24px;background:rgba(176,242,174,.025);">
            </div>
          </td>
        </tr>`;
      }).join('')}
    </tbody>
  </table>
  </div>`;

  // Paginación
  if (pagEl) {
    const bs = act => `padding:5px 10px;border-radius:6px;border:1px solid ${act?'rgba(176,242,174,.5)':'rgba(255,255,255,.08)'};background:${act?'rgba(176,242,174,.12)':'rgba(255,255,255,.03)'};color:${act?'#B0F2AE':'#94a3b8'};cursor:pointer;font-family:'JetBrains Mono',monospace;font-size:11px;margin:2px;`;
    let h = `<span style="font-size:11px;color:#475569;font-family:'JetBrains Mono',monospace;margin-right:8px;">Pág ${cmrPage}/${totalPages}</span>`;
    if (cmrPage > 1) h += `<button onclick="window.cmrGoPage(${cmrPage-1})" style="${bs(false)}">‹</button>`;
    const half=3, st=Math.max(1,cmrPage-half), en=Math.min(totalPages,cmrPage+half);
    if (st>1){h+=`<button onclick="window.cmrGoPage(1)" style="${bs(false)}">1</button>`;if(st>2)h+=`<span style="color:#475569;padding:0 4px;">…</span>`;}
    for(let p=st;p<=en;p++) h+=`<button onclick="window.cmrGoPage(${p})" style="${bs(p===cmrPage)}">${p}</button>`;
    if(en<totalPages){if(en<totalPages-1)h+=`<span style="color:#475569;padding:0 4px;">…</span>`;h+=`<button onclick="window.cmrGoPage(${totalPages})" style="${bs(false)}">${totalPages}</button>`;}
    if(cmrPage<totalPages) h+=`<button onclick="window.cmrGoPage(${cmrPage+1})" style="${bs(false)}">›</button>`;
    pagEl.innerHTML = h;
  }
}

window.cmrGoPage = function(p) {
  const total = CMR_RESULTS.length;
  const totalPages = Math.max(1, Math.ceil(total / CMR_PAGE_SIZE));
  if (p >= 1 && p <= totalPages) { cmrPage = p; _cmrRenderTable(); }
};

// ──────────────────────────────────────────────────────────────────
//  DETALLE EXPANDIBLE — MOs POR PUNTO
// ──────────────────────────────────────────────────────────────────
window.cmrToggleDetail = function(code) {
  const row = document.getElementById(`cmr-detail-row-${code}`);
  if (!row) return;
  const isOpen = row.style.display !== 'none';
  if (isOpen) {
    row.style.display = 'none';
    return;
  }
  row.style.display = 'table-row';
  _cmrRenderDetail(code);
};

function _cmrRenderDetail(code) {
  const panel = document.getElementById(`cmr-detail-panel-${code}`);
  if (!panel) return;
  const rec = CMR_RESULTS.find(r => r.code === code);
  if (!rec || !rec.found) { panel.innerHTML = '<p style="color:#64748b;">Sin datos.</p>'; return; }

  // ── Resumen del sitio ──────────────────────────────────────────
  const sg = rec.sem;
  let html = `
    <div style="display:grid;grid-template-columns:1fr 1fr 1fr 1fr;gap:12px;margin-bottom:20px;">
      <div style="background:rgba(0,0,0,.3);border-radius:10px;padding:12px 14px;border:1px solid rgba(255,255,255,.06);">
        <div style="font-size:10px;color:#64748b;font-family:'Outfit',sans-serif;margin-bottom:4px;">NOMBRE</div>
        <div style="font-size:13px;font-weight:600;color:#f1f5f9;">${rec.meta.nombre_sitio || code}</div>
        <div style="font-size:10px;color:#64748b;margin-top:2px;">${rec.meta.ciudad || ''} · ${rec.meta.departamento || ''}</div>
      </div>
      <div style="background:${sg.bg};border-radius:10px;padding:12px 14px;border:1px solid ${sg.col}22;">
        <div style="font-size:10px;color:#64748b;font-family:'Outfit',sans-serif;margin-bottom:4px;">COBERTURA</div>
        <div style="font-size:22px;font-weight:700;color:${sg.col};font-family:'JetBrains Mono',monospace;">${rec.saldo_meses_str}<span style="font-size:12px;">m</span></div>
        <div style="font-size:10px;color:${sg.col};margin-top:2px;">${sg.emoji} ${sg.label} · Saldo: ${cmrFmtI(rec.saldo_rollos)} rollos</div>
      </div>
      <div style="background:rgba(153,209,252,.06);border-radius:10px;padding:12px 14px;border:1px solid rgba(153,209,252,.15);">
        <div style="font-size:10px;color:#64748b;font-family:'Outfit',sans-serif;margin-bottom:4px;">CONSUMO</div>
        <div style="font-size:14px;font-weight:700;color:#99D1FC;font-family:'JetBrains Mono',monospace;">${cmrFmtI(rec.prom_mensual)}<span style="font-size:10px;color:#64748b;"> /mes</span></div>
        <div style="font-size:10px;color:#64748b;margin-top:2px;">${rec.prom_diario.toFixed(1)}/día · ${cmrFmt(rec.prom_semanal)}/sem</div>
      </div>
      <div style="background:rgba(192,132,252,.06);border-radius:10px;padding:12px 14px;border:1px solid rgba(192,132,252,.15);">
        <div style="font-size:10px;color:#64748b;font-family:'Outfit',sans-serif;margin-bottom:4px;">STOCK INVENTARIO</div>
        <div style="font-size:22px;font-weight:700;color:#C084FC;font-family:'JetBrains Mono',monospace;">${cmrFmtI(rec.inv_stock)}</div>
        <div style="font-size:10px;color:#64748b;margin-top:2px;">${rec.inv_refs.length} ref${rec.inv_refs.length !== 1 ? 's' : ''} · ${rec.meta.proyecto || ''}</div>
      </div>
    </div>`;

  // ── Rollos por MO ──────────────────────────────────────────────
  const mosArr = [...rec.rollos_por_mo.entries()].sort((a,b) => b[1].total - a[1].total);
  if (mosArr.length > 0) {
    const page = cmrDetailPage[code] || 1;
    const totalMOPages = Math.max(1, Math.ceil(mosArr.length / CMR_MO_PAGE_SIZE));
    const moSlice = mosArr.slice((page-1)*CMR_MO_PAGE_SIZE, page*CMR_MO_PAGE_SIZE);

    html += `
      <div style="margin-bottom:20px;">
        <div style="font-family:'Syne',sans-serif;font-size:12px;font-weight:700;color:#B0F2AE;letter-spacing:.5px;text-transform:uppercase;margin-bottom:10px;">
          📋 MOs / Movimientos — ${rec.mos_unicas.size} únicas · ${rec.movs.length} registros
        </div>
        <div style="overflow-x:auto;">
        <table style="width:100%;border-collapse:collapse;font-size:11px;">
          <thead>
            <tr style="background:rgba(0,0,0,.4);border-bottom:1px solid rgba(176,242,174,.12);">
              <th style="padding:8px 12px;text-align:left;color:#64748b;font-weight:600;font-size:10px;letter-spacing:.4px;">TAREA / MO</th>
              <th style="padding:8px 10px;text-align:right;color:#64748b;font-weight:600;font-size:10px;">ROLLOS</th>
              <th style="padding:8px 10px;text-align:left;color:#64748b;font-weight:600;font-size:10px;">ESTADO</th>
              <th style="padding:8px 10px;text-align:left;color:#64748b;font-weight:600;font-size:10px;">FLUJO</th>
              <th style="padding:8px 10px;text-align:left;color:#64748b;font-weight:600;font-size:10px;">MATERIAL</th>
              <th style="padding:8px 10px;text-align:left;color:#64748b;font-weight:600;font-size:10px;">FECHA</th>
              <th style="padding:8px 10px;text-align:left;color:#64748b;font-weight:600;font-size:10px;">GUÍA</th>
              <th style="padding:8px 10px;text-align:left;color:#64748b;font-weight:600;font-size:10px;">ORIGEN</th>
            </tr>
          </thead>
          <tbody>
            ${moSlice.map(([tarea, moData], ii) => {
              const lastMov = moData.movs[0];
              const bgRow = ii % 2 === 0 ? 'transparent' : 'rgba(255,255,255,.01)';
              // Si hay múltiples movimientos, mostrar la primera fila con rowspan y las demás
              if (moData.movs.length <= 1) {
                const m = moData.movs[0] || {};
                return `<tr style="background:${bgRow};border-bottom:1px solid rgba(255,255,255,.04);">
                  <td style="padding:7px 12px;color:#DFFF61;font-family:'JetBrains Mono',monospace;font-size:10px;max-width:160px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;" title="${tarea}">${tarea}</td>
                  <td style="padding:7px 10px;text-align:right;font-family:'JetBrains Mono',monospace;font-weight:700;color:#B0F2AE;">${cmrFmtI(moData.total)}</td>
                  <td style="padding:7px 10px;color:#94a3b8;font-size:10px;">${m.estado || '—'}</td>
                  <td style="padding:7px 10px;color:#7B8CDE;font-size:10px;">${m.flujo || '—'}</td>
                  <td style="padding:7px 10px;color:#f1f5f9;font-size:10px;max-width:120px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;" title="${m.material||''}">${m.material || '—'}</td>
                  <td style="padding:7px 10px;color:#64748b;white-space:nowrap;font-size:10px;">${cmrFmtDate(m.fecha)}</td>
                  <td style="padding:7px 10px;color:#64748b;font-size:10px;max-width:100px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;" title="${m.guia||''}">${m.guia || '—'}</td>
                  <td style="padding:7px 10px;color:#64748b;font-size:10px;max-width:120px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;" title="${m.origen||''}">${m.origen || '—'}</td>
                </tr>`;
              } else {
                // Múltiples movimientos: primera fila de resumen + sub-filas
                return `<tr style="background:rgba(176,242,174,.03);border-bottom:1px solid rgba(176,242,174,.08);">
                  <td style="padding:7px 12px;color:#DFFF61;font-family:'JetBrains Mono',monospace;font-size:10px;" colspan="8">
                    <span style="font-weight:700;">${tarea}</span>
                    <span style="color:#64748b;font-size:10px;margin-left:10px;">${moData.movs.length} movimientos</span>
                    <span style="background:rgba(176,242,174,.12);color:#B0F2AE;border-radius:6px;padding:1px 8px;font-size:10px;margin-left:8px;">Total: ${cmrFmtI(moData.total)} rollos</span>
                  </td>
                </tr>
                ${moData.movs.slice(0, 5).map((m, mi) => `
                  <tr style="background:${mi%2===0?'rgba(0,0,0,.15)':'rgba(0,0,0,.08)'};border-bottom:1px solid rgba(255,255,255,.03);">
                    <td style="padding:5px 12px 5px 24px;color:#475569;font-size:10px;">└ mov ${mi+1}</td>
                    <td style="padding:5px 10px;text-align:right;font-family:'JetBrains Mono',monospace;color:#B0F2AE;font-size:11px;">${cmrFmtI(m.cantidad)}</td>
                    <td style="padding:5px 10px;color:#94a3b8;font-size:10px;">${m.estado || '—'}</td>
                    <td style="padding:5px 10px;color:#7B8CDE;font-size:10px;">${m.flujo || '—'}</td>
                    <td style="padding:5px 10px;color:#f1f5f9;font-size:10px;max-width:120px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${m.material || '—'}</td>
                    <td style="padding:5px 10px;color:#64748b;white-space:nowrap;font-size:10px;">${cmrFmtDate(m.fecha)}</td>
                    <td style="padding:5px 10px;color:#64748b;font-size:10px;">${m.guia || '—'}</td>
                    <td style="padding:5px 10px;color:#64748b;font-size:10px;max-width:120px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${m.origen || '—'}</td>
                  </tr>`).join('')}
                ${moData.movs.length > 5 ? `<tr><td colspan="8" style="padding:4px 24px;color:#475569;font-size:10px;">… y ${moData.movs.length - 5} movimientos más</td></tr>` : ''}`;
              }
            }).join('')}
          </tbody>
        </table>
        </div>
        ${totalMOPages > 1 ? `<div style="margin-top:8px;display:flex;align-items:center;gap:6px;flex-wrap:wrap;">
          <span style="font-size:10px;color:#475569;font-family:'JetBrains Mono',monospace;">MOs pág ${page}/${totalMOPages}</span>
          ${page > 1 ? `<button onclick="window.cmrMOPage('${code}',${page-1})" style="padding:3px 8px;border-radius:5px;border:1px solid rgba(255,255,255,.1);background:rgba(255,255,255,.04);color:#94a3b8;cursor:pointer;font-size:10px;">‹</button>` : ''}
          ${page < totalMOPages ? `<button onclick="window.cmrMOPage('${code}',${page+1})" style="padding:3px 8px;border-radius:5px;border:1px solid rgba(255,255,255,.1);background:rgba(255,255,255,.04);color:#94a3b8;cursor:pointer;font-size:10px;">›</button>` : ''}
        </div>` : ''}
      </div>`;
  } else {
    html += `<div style="padding:16px;background:rgba(0,0,0,.2);border-radius:10px;color:#475569;font-size:12px;margin-bottom:16px;">Sin movimientos/MOs registrados para este punto.</div>`;
  }

  // ── Stock de inventario por referencias ──────────────────────────
  if (rec.inv_refs.length > 0) {
    html += `
      <div style="margin-bottom:16px;">
        <div style="font-family:'Syne',sans-serif;font-size:12px;font-weight:700;color:#C084FC;letter-spacing:.5px;text-transform:uppercase;margin-bottom:10px;">
          🏭 Stock Inventario (wompi_inventario) — ${cmrFmtI(rec.inv_stock)} unidades
        </div>
        <div style="display:flex;flex-wrap:wrap;gap:8px;">
          ${rec.inv_refs.map(([ref, qty]) => `
            <div style="background:rgba(192,132,252,.08);border:1px solid rgba(192,132,252,.2);border-radius:8px;padding:8px 14px;min-width:160px;">
              <div style="font-size:10px;color:#9f7aea;font-weight:700;font-family:'JetBrains Mono',monospace;">${cmrFmtI(qty)} uds</div>
              <div style="font-size:11px;color:#f1f5f9;margin-top:2px;max-width:180px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;" title="${ref}">${ref}</div>
            </div>`).join('')}
        </div>
      </div>`;
  }

  panel.innerHTML = html;
}

window.cmrMOPage = function(code, p) {
  cmrDetailPage[code] = p;
  _cmrRenderDetail(code);
};

// ──────────────────────────────────────────────────────────────────
//  EXPORT EXCEL REPORTE COMPLETO
// ──────────────────────────────────────────────────────────────────
window.cmrExportExcel = function () {
  if (!window.XLSX) { alert('Librería XLSX no disponible'); return; }
  if (!CMR_RESULTS.length) { alert('No hay resultados para exportar. Carga un archivo primero.'); return; }

  const wb = window.XLSX.utils.book_new();

  // ── Hoja 1: Resumen General ──────────────────────────────────────
  const resumen = CMR_RESULTS.map((r, i) => ({
    '#': i + 1,
    'Código Punto': r.code,
    'Encontrado': r.found ? 'Sí' : 'No',
    'Nombre Corresponsal': r.meta.nombre_sitio || '',
    'Departamento': r.meta.departamento || '',
    'Ciudad': r.meta.ciudad || '',
    'Proyecto': r.meta.proyecto || '',
    'Estado Punto': r.cal.estado_punto || '',
    'Saldo Rollos': Math.round(r.saldo_rollos),
    'Stock Inventario': r.inv_stock,
    'Saldo Días': Math.round(r.saldo_dias),
    'Cobertura (meses)': +r.saldo_meses.toFixed(2),
    'Semáforo': r.sem.label,
    'Consumo Mensual': Math.round(r.prom_mensual),
    'Consumo Diario': +r.prom_diario.toFixed(2),
    'Consumo Semanal': +r.prom_semanal.toFixed(2),
    'Punto Reorden': Math.round(r.punto_reorden),
    'Bajo Punto Reorden': r.bajo_reorden ? 'Sí' : 'No',
    'SLA 3 Meses': r.sla_cumple ? 'Sí' : 'No',
    'Riesgo Quiebre': r.ind_riesgo || '',
    'Rotación Inventario': +r.rotacion.toFixed(3),
    'Rollos Entregados': Math.round(r.rollos_entregados),
    'Rollos Consumidos': Math.round(r.rollos_consumidos),
    'MOs Únicas': r.mos_unicas.size,
    'Total Movimientos': r.movs.length,
    'Fecha Apertura': r.fecha_apertura,
    'Fecha Abastecimiento': r.fecha_abst,
    'Periodo Abast. (días)': cmrP(r.cal.periodo_abast),
    'Rollos Proyectados': Math.round(cmrP(r.cal.rollos_periodo)),
    'Rollos/Año Estimados': Math.round(cmrP(r.cal.rollos_anio)),
  }));
  const wsRes = window.XLSX.utils.json_to_sheet(resumen);
  wsRes['!cols'] = Object.keys(resumen[0] || {}).map((k, i) => ({ wch: i < 3 ? 8 : Math.min(Math.max(k.length + 2, 12), 30) }));
  window.XLSX.utils.book_append_sheet(wb, wsRes, 'Resumen General');

  // ── Hoja 2: Detalle MOs ──────────────────────────────────────────
  const moRows = [];
  CMR_RESULTS.forEach(r => {
    if (!r.found || !r.movs.length) return;
    r.movs.forEach(m => {
      moRows.push({
        'Código Punto': r.code,
        'Nombre Corresponsal': r.meta.nombre_sitio || '',
        'Ciudad': r.meta.ciudad || '',
        'Tarea / MO': m.tarea,
        'Cantidad Rollos': m.cantidad || 0,
        'Material': m.material || '',
        'Código Material': m.codigo_material || '',
        'Estado': m.estado || '',
        'Flujo': m.flujo || '',
        'Fecha Confirmación': m.fecha,
        'Fecha Fin': m.fecha_fin,
        'Plan Inicio': m.plan_inicio,
        'Plan Fin': m.plan_fin,
        'Guía': m.guia || '',
        'Transportadora': m.transportadora || '',
        'Origen': m.origen || '',
        'Destino': m.destino || '',
        'Plantilla': m.plantilla || '',
        'Subproyecto': m.subproyecto || '',
        'Cód. Operación': m.cod_op || '',
      });
    });
  });
  if (moRows.length) {
    const wsMO = window.XLSX.utils.json_to_sheet(moRows);
    wsMO['!cols'] = Object.keys(moRows[0] || {}).map((k, i) => ({ wch: Math.min(Math.max(k.length + 2, 12), 32) }));
    window.XLSX.utils.book_append_sheet(wb, wsMO, 'Detalle MOs');
  }

  // ── Hoja 3: Stock Inventario ─────────────────────────────────────
  const invRows = [];
  CMR_RESULTS.forEach(r => {
    if (!r.found) return;
    if (r.inv_refs.length === 0) {
      invRows.push({ 'Código Punto': r.code, 'Nombre Corresponsal': r.meta.nombre_sitio || '', 'Ciudad': r.meta.ciudad || '', 'Referencia': '—', 'Cantidad': 0, 'Total Stock Punto': r.inv_stock });
    } else {
      r.inv_refs.forEach(([ref, qty]) => {
        invRows.push({ 'Código Punto': r.code, 'Nombre Corresponsal': r.meta.nombre_sitio || '', 'Ciudad': r.meta.ciudad || '', 'Referencia': ref, 'Cantidad': qty, 'Total Stock Punto': r.inv_stock });
      });
    }
  });
  if (invRows.length) {
    const wsInv = window.XLSX.utils.json_to_sheet(invRows);
    wsInv['!cols'] = Object.keys(invRows[0] || {}).map(k => ({ wch: Math.min(Math.max(k.length + 2, 14), 35) }));
    window.XLSX.utils.book_append_sheet(wb, wsInv, 'Stock Inventario');
  }

  // ── Hoja 4: Rollos por MO ────────────────────────────────────────
  const rollsMO = [];
  CMR_RESULTS.forEach(r => {
    if (!r.found) return;
    for (const [tarea, moData] of r.rollos_por_mo) {
      rollsMO.push({
        'Código Punto': r.code,
        'Nombre Corresponsal': r.meta.nombre_sitio || '',
        'Ciudad': r.meta.ciudad || '',
        'Tarea / MO': tarea,
        'Total Rollos en MO': moData.total,
        'Cantidad Movimientos': moData.movs.length,
        'Último Flujo': (moData.movs[0] || {}).flujo || '',
        'Último Estado': (moData.movs[0] || {}).estado || '',
        'Última Fecha': (moData.movs[0] || {}).fecha || '',
      });
    }
  });
  if (rollsMO.length) {
    const wsRM = window.XLSX.utils.json_to_sheet(rollsMO);
    wsRM['!cols'] = Object.keys(rollsMO[0] || {}).map(k => ({ wch: Math.min(Math.max(k.length + 2, 14), 35) }));
    window.XLSX.utils.book_append_sheet(wb, wsRM, 'Rollos por MO');
  }

  // ── Hoja 5: No Encontrados ───────────────────────────────────────
  const noFound = CMR_RESULTS.filter(r => !r.found).map((r, i) => ({ '#': i+1, 'Código Consultado': r.code, 'Observación': 'No encontrado en TABLERO_ROLLOS_FILAS ni RCV2_SITIOS' }));
  if (noFound.length) {
    const wsNF = window.XLSX.utils.json_to_sheet(noFound);
    wsNF['!cols'] = [{ wch: 5 }, { wch: 22 }, { wch: 55 }];
    window.XLSX.utils.book_append_sheet(wb, wsNF, 'No Encontrados');
  }

  const fecha = new Date().toISOString().slice(0, 10);
  window.XLSX.writeFile(wb, `reporte_masivo_rollos_${fecha}.xlsx`);
};

// ──────────────────────────────────────────────────────────────────
//  LIMPIAR CONSULTA
// ──────────────────────────────────────────────────────────────────
window.cmrLimpiar = function () {
  CMR_RESULTS = [];
  CMR_CODES   = [];
  cmrPage     = 1;
  cmrDetailPage = {};
  const statusEl = document.getElementById('cmr-file-status');
  if (statusEl) statusEl.innerHTML = '';
  const fileInput = document.getElementById('cmr-file-input');
  if (fileInput) fileInput.value = '';
  const resultSection = document.getElementById('cmr-results-section');
  if (resultSection) resultSection.style.display = 'none';
};

// ──────────────────────────────────────────────────────────────────
//  INIT — llamado al activar el tab
// ──────────────────────────────────────────────────────────────────
window.renderConsultaMasivaRollos = function () {
  // Asegurar que el panel esté visible (defensa ante race conditions de navegación)
  const panel = document.getElementById('panel-rollos-consulta-masiva');
  if (panel && panel.style.display === 'none') panel.style.display = 'block';

  // Si RCV2_SITIOS no fue construido (el tab de Comercio nunca fue visitado),
  // iniciarlo ahora para que las búsquedas funcionen.
  if (!window.RCV2_SITIOS && typeof window.initRollosComercioV2 === 'function') {
    try { window.initRollosComercioV2(); } catch(e) { console.warn('[CMR] initRollosComercioV2:', e); }
  }

  // Re-renderizar resultados si hay consulta previa.
  if (CMR_RESULTS.length) {
    _cmrRenderKPIs();
    _cmrRenderTable();
  }
};