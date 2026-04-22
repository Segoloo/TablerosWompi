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

// SIMCards de prueba a excluir
const EM_SIM_PRUEBA = new Set([
  'SIMVP00001','SIMVP00002','SIMVP00003','SIMVP00004',
  'SIMP00001','SIMP00002','SIMP00003',
  '12341234','234234000000000'
]);

// Bodegas a excluir del KPI "Total en Bodegas"
const EM_BODEGAS_EXCLUIR_TOTAL = new Set([
  'ALMACEN BAJAS WOMPI','ALMACÉN BAJAS WOMPI',
  'ALMACEN INGENICO - PROVEEDOR WOMPI'
]);

const EM_COLORS = {
  danio:      '#FF5C5C',
  incidente:  '#FFC04D',
  cierre:     '#99D1FC',
  asociado:   '#C084FC',
  disponible: '#B0F2AE',
  activada:   '#B0F2AE',
  sinActivar: '#FF5C5C',
  total:      '#DFFF61',
  palette: ['#B0F2AE','#99D1FC','#DFFF61','#C084FC','#FFC04D','#FF5C5C','#F49D6E','#7B8CDE','#00C87A','#A8E6CF'],
};

let EM_CHARTS    = {};
let EM_DF_ALL    = [];
let EM_SIM_ALL   = [];
let EM_DF_PAGE   = 1;
let EM_SIM_PAGE  = 1;
const EM_PAGE_SIZE = 50;
let EM_DF_SEARCH    = '';
let EM_DF_FILTERED  = [];
let EM_SIM_SEARCH   = '';
let EM_SIM_FILTERED = [];
let _EM_KPI_ROWS    = [];

// Filtros globales de Estado Materiales
let EM_F_NEGOCIO  = '';   // 'CB' | 'VP' | ''
let EM_F_UBICV3   = '';   // valor de invUbicacionV3 | ''
let EM_F_BODEGA   = '';   // nombre de bodega exacto | ''
let EM_F_SERIAL   = '';   // número de serie exacto | ''
let EM_F_REF      = '';   // referencia exacta | ''

// ══════════════════════════════════════════════════════════════════
//  HELPERS DE DATOS
// ══════════════════════════════════════════════════════════════════

function _emGetRaw() {
  if (window.INV_RAW && window.INV_RAW.length) return window.INV_RAW;
  if (typeof window.invGetRaw === 'function') return window.invGetRaw();
  return [];
}

function _emDestroyChart(id) {
  if (EM_CHARTS[id]) { try { EM_CHARTS[id].destroy(); } catch(e){} delete EM_CHARTS[id]; }
}

// ── Lógica estado datáfono ────────────────────────────────────────
// Precedencia:
//   1. Posición depósito DESINSTALADO-INCIDENTE → DES. INCIDENTE
//   2. Posición depósito DESINSTALADO-CIERRE    → DES. CIERRE
//   3. Comentarios vacíos                        → DISPONIBLE
//   4. Primer segmento ≠ 99999                  → ASOCIADO
//   5. 99999 | DESASOCIADO                       → DISPONIBLE
//   6. 99999 | MIGRACION*                        → DISPONIBLE
//   7. 99999 | <otro texto>                      → DAÑADO
//   8. resto                                     → DISPONIBLE

function _emEstadoDatafono(row) {
  const pos = (row['Posición en depósito'] || '').trim().toUpperCase();
  const ubi = (row['Nombre de la ubicación'] || '').trim().toUpperCase();
  const com = (row['Comentarios'] || '').trim();

  // Datáfonos en ALMACEN BAJAS WOMPI o ALMACEN INGENICO - PROVEEDOR WOMPI → estado BAJAS WOMPI (no suman en disponible)
  if (ubi.includes('ALMACEN BAJAS WOMPI') || ubi.includes('ALMACÉN BAJAS WOMPI') || ubi.includes('ALMACEN INGENICO - PROVEEDOR WOMPI')) return 'BAJAS WOMPI';

  // EN DAÑO viene directamente de la columna Posición en depósito
  if (pos === 'EN DAÑO') return 'DAÑADO';
  if (pos === 'DESINSTALADO-INCIDENTE') return 'DES. INCIDENTE';
  if (pos === 'DESINSTALADO-CIERRE')    return 'DES. CIERRE';

  if (!com) return 'DISPONIBLE';

  const parts  = com.split('|');
  const numero = (parts[0] || '').trim();
  const texto  = (parts[1] || '').trim().toUpperCase();

  if (/^\d+$/.test(numero) && numero !== '99999') return 'ASOCIADO';

  if (numero === '99999') {
    if (!texto)                         return 'DISPONIBLE';
    if (texto === 'DESASOCIADO')        return 'DISPONIBLE';
    if (texto.includes('MIGRACION'))    return 'DISPONIBLE';
    return 'DISPONIBLE'; // ya no se mapea a DAÑADO desde comentarios
  }

  return 'DISPONIBLE';
}

// Extrae tipo de daño de comentarios: solo 99999 | <texto> donde texto ≠ DESASOCIADO/MIGRACION
function _emTipoDanio(row) {
  const com = (row['Comentarios'] || '').trim();
  if (!com) return null;
  const parts  = com.split('|');
  const numero = (parts[0] || '').trim();
  const texto  = (parts[1] || '').trim().toUpperCase();
  if (numero !== '99999') return null;
  if (!texto || texto === 'DESASOCIADO' || texto.includes('MIGRACION')) return null;
  return texto;
}

// SIMCard helpers
function _emSimActivada(row) {
  // Acepta FA:DD/MM/AAAA o FA: DD/MM/AAAA (con espacio tras los dos puntos), año parcial o ausente
  return /FA:\s*\d{2}\/\d{2}(\/\d{0,4})?/i.test(row['Atributos'] || '');
}
function _emSimFechaActivacion(row) {
  const m = (row['Atributos'] || '').match(/FA:\s*(\d{2}\/\d{2}(?:\/\d{0,4})?)/i);
  return m ? m[1] : '—';
}

// ══════════════════════════════════════════════════════════════════
//  CHART HELPERS
// ══════════════════════════════════════════════════════════════════

const _EM_TOOLTIP = {
  backgroundColor: 'rgba(24,23,21,.95)',
  titleColor: '#B0F2AE', bodyColor: '#FAFAFA',
  borderColor: 'rgba(176,242,174,.2)', borderWidth: 1, padding: 12,
  titleFont: { family: 'Syne', size: 13, weight: '700' },
  bodyFont:  { family: 'Outfit', size: 12 },
};

function _emDonut(canvasId, labels, data, colors) {
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;
  _emDestroyChart(canvasId);
  const total = data.reduce((a, b) => a + b, 0);
  EM_CHARTS[canvasId] = new Chart(canvas, {
    type: 'doughnut',
    data: { labels, datasets: [{ data, backgroundColor: colors, borderWidth: 2, borderColor: '#181715', hoverOffset: 6 }] },
    options: {
      cutout: '65%', responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { position: 'bottom', labels: { color: '#cbd5e1', font: { size: 11, family: "'Outfit',sans-serif" }, padding: 12, boxWidth: 12 } },
        tooltip: { ..._EM_TOOLTIP, callbacks: { label: ctx => ` ${ctx.label}: ${ctx.parsed.toLocaleString('es-CO')} (${total?(ctx.parsed/total*100).toFixed(1):0}%)` } },
      }
    }
  });
}

function _emBarH(canvasId, labels, data, colorsArr) {
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;
  _emDestroyChart(canvasId);
  EM_CHARTS[canvasId] = new Chart(canvas, {
    type: 'bar',
    data: { labels, datasets: [{ data,
      backgroundColor: colorsArr || labels.map((_,i) => EM_COLORS.palette[i%EM_COLORS.palette.length]+'BB'),
      borderColor:     colorsArr || labels.map((_,i) => EM_COLORS.palette[i%EM_COLORS.palette.length]),
      borderWidth: 1, borderRadius: 6 }]
    },
    options: {
      indexAxis: 'y', responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false }, tooltip: _EM_TOOLTIP },
      scales: {
        x: { ticks: { color: '#7A7674', font: { size: 10 } }, grid: { color: 'rgba(255,255,255,.05)' } },
        y: { ticks: { color: '#e2e8f0', font: { size: 11, family: "'Outfit',sans-serif" } }, grid: { display: false } }
      }
    }
  });
}

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
        legend: { labels: { color: '#cbd5e1', font: { size: 11, family: "'Outfit',sans-serif" }, padding: 10, boxWidth: 12 } },
        tooltip: _EM_TOOLTIP,
      },
      scales: {
        x: { ticks: { color: '#cbd5e1', font: { size: 10 }, maxRotation: 30 }, grid: { display: false } },
        y: { ticks: { color: '#7A7674' }, grid: { color: 'rgba(255,255,255,.05)' } }
      }
    }
  });
}

// ══════════════════════════════════════════════════════════════════
//  MODAL DRILL-DOWN DATÁFONOS
// ══════════════════════════════════════════════════════════════════

function _emOpenModal(title, rows, accent) {
  const prev = document.getElementById('em-drill-overlay');
  if (prev) prev.remove();

  if (!document.getElementById('em-kf')) {
    const st = document.createElement('style');
    st.id = 'em-kf';
    st.textContent = '@keyframes emFadeIn{from{opacity:0}to{opacity:1}}@keyframes emSlideUp{from{opacity:0;transform:translateY(20px) scale(.97)}to{opacity:1;transform:none}}';
    document.head.appendChild(st);
  }

  const ov = document.createElement('div');
  ov.id = 'em-drill-overlay';
  ov.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,.78);z-index:9000;display:flex;align-items:center;justify-content:center;padding:24px;animation:emFadeIn .18s ease;';

  const pnl = document.createElement('div');
  pnl.style.cssText = `background:#181715;border:1px solid rgba(255,255,255,.1);border-top:3px solid ${accent||'#B0F2AE'};border-radius:20px;width:min(1100px,96vw);max-height:88vh;display:flex;flex-direction:column;animation:emSlideUp .22s cubic-bezier(.4,0,.2,1);overflow:hidden;`;

  // Header
  const hdr = document.createElement('div');
  hdr.style.cssText = 'display:flex;align-items:center;justify-content:space-between;padding:18px 24px;border-bottom:1px solid rgba(255,255,255,.07);flex-shrink:0;';
  hdr.innerHTML = `<div>
    <div style="font-family:'Syne',sans-serif;font-size:15px;font-weight:800;color:#f1f5f9;">${title}</div>
    <div style="font-size:11px;color:#7A7674;margin-top:2px;font-family:'Outfit',sans-serif;" id="em-modal-count">${rows.length.toLocaleString('es-CO')} registros</div>
  </div>`;
  const cb = document.createElement('button');
  cb.innerHTML = '✕';
  cb.style.cssText = 'background:transparent;border:none;color:#7A7674;font-size:18px;cursor:pointer;padding:6px 10px;border-radius:8px;transition:all .15s;';
  cb.onmouseover = () => { cb.style.background='rgba(255,80,80,.12)'; cb.style.color='#f87171'; };
  cb.onmouseout  = () => { cb.style.background='transparent';         cb.style.color='#7A7674'; };
  cb.onclick = () => ov.remove();
  hdr.appendChild(cb);
  pnl.appendChild(hdr);

  // Barra búsqueda + export
  const bar = document.createElement('div');
  bar.style.cssText = 'padding:12px 24px;border-bottom:1px solid rgba(255,255,255,.05);display:flex;gap:10px;background:rgba(0,0,0,.18);flex-shrink:0;';
  const si = document.createElement('input');
  si.placeholder = '🔍 Buscar por serial, referencia, bodega, estado, comentarios...';
  si.style.cssText = `flex:1;background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.12);border-radius:8px;color:#FAFAFA;padding:8px 12px;font-size:13px;font-family:'Outfit',sans-serif;outline:none;`;
  bar.appendChild(si);
  const eb = document.createElement('button');
  eb.textContent = '⬇ Excel';
  eb.style.cssText = `background:rgba(176,242,174,.08);border:1px solid rgba(176,242,174,.2);color:#B0F2AE;padding:8px 14px;border-radius:8px;cursor:pointer;font-size:12px;font-family:'Outfit',sans-serif;white-space:nowrap;`;
  eb.onclick = () => _emExportModalExcel(title, filtered);
  bar.appendChild(eb);
  pnl.appendChild(bar);

  // Tabla wrap
  const tw = document.createElement('div');
  tw.style.cssText = 'overflow:auto;flex:1;';

  const COLS = [
    { h:'Serial',        fn:r=>r['Número de serie']||r['Numero de serie']||r['Serial']||'—', mono:true, col:'#a5f3fc' },
    { h:'Referencia',    fn:r=>r['Nombre']||'—', col:'#e2e8f0' },
    { h:'Bodega',        fn:r=>r['Nombre de la ubicación']||'—', col:'#cbd5e1' },
    { h:'Cód. Ubic.',   fn:r=>r['Código de ubicación']||'—', mono:true, col:'#67e8f9' },
    { h:'Pos. Depósito', fn:r=>r['Posición en depósito']||'—', col:'#94a3b8' },
    { h:'Estado',        fn:r=>_emEstadoDatafono(r), isEstado:true },
    { h:'Comentarios',   fn:r=>r['Comentarios']||'—', col:'#7A7674' },
  ];

  const _epill = est => {
    const m = { 'DISPONIBLE':['#B0F2AE','rgba(176,242,174,.12)'], 'ASOCIADO':['#C084FC','rgba(192,132,252,.12)'],
      'DAÑADO':['#FF5C5C','rgba(255,92,92,.12)'], 'DES. INCIDENTE':['#FFC04D','rgba(255,192,77,.12)'], 'DES. CIERRE':['#99D1FC','rgba(153,209,252,.12)'] };
    const [col, bg] = m[est]||['#7A7674','rgba(255,255,255,.06)'];
    return `<span style="background:${bg};color:${col};font-size:10px;font-weight:700;padding:3px 10px;border-radius:12px;border:1px solid ${col}44;display:inline-block;font-family:'Outfit',sans-serif;white-space:nowrap;">${est}</span>`;
  };

  let filtered = rows.slice();
  let pg = 1;
  const PG = 50;

  function rt() {
    const pages = Math.max(1, Math.ceil(filtered.length/PG));
    if (pg>pages) pg=pages;
    const slice = filtered.slice((pg-1)*PG, pg*PG);
    const cntEl = pnl.querySelector('#em-modal-count');
    if (cntEl) cntEl.textContent = `${filtered.length.toLocaleString('es-CO')} registros`;

    tw.innerHTML = `<table style="width:100%;border-collapse:collapse;font-size:12px;font-family:'Outfit',sans-serif;">
      <thead><tr style="background:#181715;position:sticky;top:0;z-index:2;">
        ${COLS.map(c=>`<th style="padding:10px 14px;text-align:left;font-family:'Syne',sans-serif;font-size:9px;font-weight:700;color:${accent||'#B0F2AE'};letter-spacing:1px;text-transform:uppercase;border-bottom:1px solid rgba(255,255,255,.08);white-space:nowrap;">${c.h}</th>`).join('')}
      </tr></thead>
      <tbody>${!slice.length?`<tr><td colspan="${COLS.length}" style="text-align:center;padding:40px;color:#7A7674;">Sin registros</td></tr>`:
        slice.map((r,i)=>{
          const bg=i%2===0?'rgba(255,255,255,.015)':'transparent';
          return `<tr style="background:${bg}" onmouseover="this.style.background='rgba(255,255,255,.04)'" onmouseout="this.style.background='${bg}'">
            ${COLS.map(c=>{
              const v=c.fn(r);
              if(c.isEstado) return `<td style="padding:9px 14px;">${_epill(v)}</td>`;
              return `<td style="padding:9px 14px;${c.mono?`font-family:'JetBrains Mono',monospace;font-size:11px;`:'font-size:11px;'}color:${c.col||'#e2e8f0'};max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="${String(v).replace(/"/g,'&quot;')}">${v}</td>`;
            }).join('')}
          </tr>`;
        }).join('')}
      </tbody>
    </table>`;

    // Paginación
    let pb = tw.nextElementSibling;
    if (!pb || !pb._emPag) {
      pb = document.createElement('div');
      pb._emPag = true;
      pb.style.cssText = 'display:flex;justify-content:flex-end;align-items:center;gap:6px;padding:10px 24px;border-top:1px solid rgba(255,255,255,.05);flex-shrink:0;';
      pnl.appendChild(pb);
    }
    pb.innerHTML = '';
    const bs = `padding:5px 11px;border-radius:7px;cursor:pointer;font-size:12px;font-family:'Outfit',sans-serif;border:1px solid rgba(255,255,255,.1);transition:all .15s;`;
    const mk = (l,p,a) => { const b=document.createElement('button');b.style.cssText=bs+(a?`background:${accent};color:#0a1a12;font-weight:700;`:'background:rgba(255,255,255,.06);color:#94a3b8;');b.textContent=l;b.onclick=()=>{pg=p;rt();};return b; };
    if(pg>1) pb.appendChild(mk('‹',pg-1,false));
    for(let p=Math.max(1,pg-2);p<=Math.min(pages,pg+2);p++) pb.appendChild(mk(String(p),p,p===pg));
    if(pg<pages) pb.appendChild(mk('›',pg+1,false));
    const lbl=document.createElement('span');lbl.style.cssText='font-size:11px;color:#7A7674;font-family:"Outfit",sans-serif;margin-left:6px;';lbl.textContent=`${filtered.length.toLocaleString('es-CO')} registros`;pb.appendChild(lbl);
  }

  si.oninput = () => {
    const q = si.value.toLowerCase().trim();
    filtered = !q ? rows.slice() : rows.filter(r => {
      const s=(r['Número de serie']||r['Numero de serie']||r['Serial']||'').toLowerCase();
      const n=(r['Nombre']||'').toLowerCase();
      const u=(r['Nombre de la ubicación']||'').toLowerCase();
      const c=(r['Comentarios']||'').toLowerCase();
      const p=(r['Posición en depósito']||'').toLowerCase();
      const e=_emEstadoDatafono(r).toLowerCase();
      return s.includes(q)||n.includes(q)||u.includes(q)||c.includes(q)||p.includes(q)||e.includes(q);
    });
    pg = 1; rt();
  };

  pnl.appendChild(tw);
  ov.appendChild(pnl);
  document.body.appendChild(ov);
  ov.addEventListener('click', e => { if(e.target===ov) ov.remove(); });
  document.addEventListener('keydown', function esc(e){ if(e.key==='Escape'){ov.remove();document.removeEventListener('keydown',esc);} });
  rt();
}

function _emExportModalExcel(title, rows) {
  if (!rows.length||typeof XLSX==='undefined'){alert('Sin datos o XLSX no disponible.');return;}
  const data=rows.map(r=>({
    'Serial':r['Número de serie']||r['Numero de serie']||r['Serial']||'',
    'Referencia':r['Nombre']||'',
    'Bodega':r['Nombre de la ubicación']||'',
    'Código Ubicación':r['Código de ubicación']||'',
    'Posición Depósito':r['Posición en depósito']||'',
    'Estado':_emEstadoDatafono(r),
    'Tipo Daño':_emTipoDanio(r)||'',
    'Comentarios':r['Comentarios']||'',
    'Atributos':r['Atributos']||'',
  }));
  const wb=XLSX.utils.book_new();
  const ws=XLSX.utils.json_to_sheet(data);
  ws['!cols']=Object.keys(data[0]).map(k=>({wch:Math.max(k.length+2,14)}));
  XLSX.utils.book_append_sheet(wb,ws,title.substring(0,31));
  XLSX.writeFile(wb,`${title.replace(/[^a-z0-9]/gi,'_')}_${new Date().toISOString().slice(0,10)}.xlsx`);
}

// Modal SIMCards
function _emOpenSimModal(title, rows, accent) {
  const prev=document.getElementById('em-drill-overlay'); if(prev) prev.remove();
  const ov=document.createElement('div');
  ov.id='em-drill-overlay';
  ov.style.cssText='position:fixed;inset:0;background:rgba(0,0,0,.78);z-index:9000;display:flex;align-items:center;justify-content:center;padding:24px;animation:emFadeIn .18s ease;';
  const pnl=document.createElement('div');
  pnl.style.cssText=`background:#181715;border:1px solid rgba(255,255,255,.1);border-top:3px solid ${accent||'#99D1FC'};border-radius:20px;width:min(900px,96vw);max-height:88vh;display:flex;flex-direction:column;animation:emSlideUp .22s cubic-bezier(.4,0,.2,1);overflow:hidden;`;
  const hdr=document.createElement('div');
  hdr.style.cssText='display:flex;align-items:center;justify-content:space-between;padding:18px 24px;border-bottom:1px solid rgba(255,255,255,.07);flex-shrink:0;';
  hdr.innerHTML=`<div><div style="font-family:'Syne',sans-serif;font-size:15px;font-weight:800;color:#f1f5f9;">${title}</div><div style="font-size:11px;color:#7A7674;margin-top:2px;font-family:'Outfit',sans-serif;" id="em-sim-modal-count">${rows.length.toLocaleString('es-CO')} registros</div></div>`;
  const cb=document.createElement('button');cb.innerHTML='✕';cb.style.cssText='background:transparent;border:none;color:#7A7674;font-size:18px;cursor:pointer;padding:6px 10px;border-radius:8px;';cb.onclick=()=>ov.remove();hdr.appendChild(cb);pnl.appendChild(hdr);
  const bar=document.createElement('div');bar.style.cssText='padding:12px 24px;border-bottom:1px solid rgba(255,255,255,.05);display:flex;gap:10px;background:rgba(0,0,0,.18);flex-shrink:0;';
  const si=document.createElement('input');si.placeholder='🔍 Buscar serial, sitio, atributos...';si.style.cssText=`flex:1;background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.12);border-radius:8px;color:#FAFAFA;padding:8px 12px;font-size:13px;font-family:'Outfit',sans-serif;outline:none;`;bar.appendChild(si);
  const eb=document.createElement('button');eb.textContent='⬇ Excel';eb.style.cssText=`background:rgba(153,209,252,.08);border:1px solid rgba(153,209,252,.2);color:#99D1FC;padding:8px 14px;border-radius:8px;cursor:pointer;font-size:12px;font-family:'Outfit',sans-serif;white-space:nowrap;`;
  eb.onclick=()=>{if(!filtered.length||typeof XLSX==='undefined'){alert('Sin datos.');return;}const d=filtered.map(r=>({'Serial':r['Número de serie']||r['Numero de serie']||r['Serial']||'','Ubicación':r['Nombre de la ubicación']||'','Cód. Ubic.':r['Código de ubicación']||'','Estado':_emSimActivada(r)?'ACTIVADA':'SIN ACTIVAR','Fecha Activación':_emSimFechaActivacion(r),'Atributos':r['Atributos']||''}));const wb=XLSX.utils.book_new();const ws=XLSX.utils.json_to_sheet(d);ws['!cols']=Object.keys(d[0]).map(k=>({wch:Math.max(k.length+2,14)}));XLSX.utils.book_append_sheet(wb,ws,'SIMCards');XLSX.writeFile(wb,`SIMCards_${title.replace(/\s+/g,'_')}_${new Date().toISOString().slice(0,10)}.xlsx`);};
  bar.appendChild(eb);pnl.appendChild(bar);
  const tw=document.createElement('div');tw.style.cssText='overflow:auto;flex:1;';
  let filtered=rows.slice(),pg=1;const PG=50;
  function rt(){
    const pages=Math.max(1,Math.ceil(filtered.length/PG));if(pg>pages)pg=pages;const slice=filtered.slice((pg-1)*PG,pg*PG);
    const cntEl=pnl.querySelector('#em-sim-modal-count');if(cntEl)cntEl.textContent=`${filtered.length.toLocaleString('es-CO')} registros`;
    tw.innerHTML=`<table style="width:100%;border-collapse:collapse;font-size:12px;font-family:'Outfit',sans-serif;">
      <thead><tr style="background:#181715;position:sticky;top:0;z-index:2;">
        ${['Serial','Ubicación','Cód. Ubic.','Estado','Fecha Activación','Atributos'].map(h=>`<th style="padding:10px 14px;text-align:left;font-family:'Syne',sans-serif;font-size:9px;font-weight:700;color:${accent||'#99D1FC'};letter-spacing:1px;text-transform:uppercase;border-bottom:1px solid rgba(255,255,255,.08);white-space:nowrap;">${h}</th>`).join('')}
      </tr></thead>
      <tbody>${slice.map((r,i)=>{const act=_emSimActivada(r);const bg=i%2===0?'rgba(255,255,255,.015)':'transparent';const pill=act?`<span style="background:rgba(176,242,174,.12);color:#B0F2AE;font-size:10px;font-weight:700;padding:3px 10px;border-radius:12px;border:1px solid rgba(176,242,174,.3);display:inline-block;font-family:'Outfit',sans-serif;">✅ ACTIVADA</span>`:`<span style="background:rgba(255,92,92,.12);color:#FF5C5C;font-size:10px;font-weight:700;padding:3px 10px;border-radius:12px;border:1px solid rgba(255,92,92,.3);display:inline-block;font-family:'Outfit',sans-serif;">❌ SIN ACTIVAR</span>`;
      return `<tr style="background:${bg}" onmouseover="this.style.background='rgba(255,255,255,.04)'" onmouseout="this.style.background='${bg}'">
        <td style="padding:9px 14px;font-family:'JetBrains Mono',monospace;font-size:11px;color:#a5f3fc;">${r['Número de serie']||r['Numero de serie']||r['Serial']||'—'}</td>
        <td style="padding:9px 14px;font-size:11px;color:#cbd5e1;max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="${r['Nombre de la ubicación']||''}">${r['Nombre de la ubicación']||'—'}</td>
        <td style="padding:9px 14px;font-family:'JetBrains Mono',monospace;font-size:11px;color:#67e8f9;">${r['Código de ubicación']||'—'}</td>
        <td style="padding:9px 14px;">${pill}</td>
        <td style="padding:9px 14px;font-family:'JetBrains Mono',monospace;font-size:11px;color:${act?'#B0F2AE':'#64748b'};">${_emSimFechaActivacion(r)}</td>
        <td style="padding:9px 14px;font-size:11px;color:#94a3b8;max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="${r['Atributos']||''}">${r['Atributos']||'—'}</td>
      </tr>`}).join('')}
      </tbody></table>`;
    let pb=tw.nextElementSibling;if(!pb||!pb._emPag){pb=document.createElement('div');pb._emPag=true;pb.style.cssText='display:flex;justify-content:flex-end;align-items:center;gap:6px;padding:10px 24px;border-top:1px solid rgba(255,255,255,.05);flex-shrink:0;';pnl.appendChild(pb);}
    pb.innerHTML='';const bs=`padding:5px 11px;border-radius:7px;cursor:pointer;font-size:12px;font-family:'Outfit',sans-serif;border:1px solid rgba(255,255,255,.1);`;
    const mk=(l,p,a)=>{const b=document.createElement('button');b.style.cssText=bs+(a?`background:${accent};color:#0a1a12;font-weight:700;`:'background:rgba(255,255,255,.06);color:#94a3b8;');b.textContent=l;b.onclick=()=>{pg=p;rt();};return b;};
    if(pg>1)pb.appendChild(mk('‹',pg-1,false));for(let p=Math.max(1,pg-2);p<=Math.min(pages,pg+2);p++)pb.appendChild(mk(String(p),p,p===pg));if(pg<pages)pb.appendChild(mk('›',pg+1,false));
    const lbl=document.createElement('span');lbl.style.cssText='font-size:11px;color:#7A7674;font-family:"Outfit",sans-serif;margin-left:6px;';lbl.textContent=`${filtered.length.toLocaleString('es-CO')} registros`;pb.appendChild(lbl);
  }
  si.oninput=()=>{const q=si.value.toLowerCase().trim();filtered=!q?rows.slice():rows.filter(r=>{const s=(r['Número de serie']||r['Numero de serie']||r['Serial']||'').toLowerCase();const u=(r['Nombre de la ubicación']||'').toLowerCase();const a=(r['Atributos']||'').toLowerCase();return s.includes(q)||u.includes(q)||a.includes(q);});pg=1;rt();};
  pnl.appendChild(tw);ov.appendChild(pnl);document.body.appendChild(ov);
  ov.addEventListener('click',e=>{if(e.target===ov)ov.remove();});
  document.addEventListener('keydown',function esc(e){if(e.key==='Escape'){ov.remove();document.removeEventListener('keydown',esc);}});
  rt();
}

// ══════════════════════════════════════════════════════════════════
//  PROCESAMIENTO
// ══════════════════════════════════════════════════════════════════

function _emProcesar() {
  const raw = _emGetRaw();
  if (!raw || !raw.length) { EM_DF_ALL=[]; EM_SIM_ALL=[]; return; }
  const bodegas = (typeof INV_BODEGAS!=='undefined'?INV_BODEGAS:null)||window.INV_BODEGAS;
  EM_DF_ALL  = raw.filter(r => EM_REFS_DATAFONO.has((r['Nombre']||'').trim()) && bodegas && bodegas.has((r['Nombre de la ubicación']||'').trim()));
  // Excluir SIMCards de prueba
  EM_SIM_ALL = raw.filter(r => {
    if ((r['Nombre']||'').trim() !== EM_REF_SIMCARD) return false;
    const serial = (r['Número de serie']||r['Numero de serie']||r['Serial']||'').trim().toUpperCase();
    // Excluir seriales de prueba (exacto y también el numérico 2.34234E+11 → 234234000000000)
    if (EM_SIM_PRUEBA.has(serial)) return false;
    // Excluir variantes numéricas del serial de prueba 2,34234E+11
    if (/^2\.?34234[e\+]?/i.test(serial) || serial === '234234000000000' || serial === '23423400000000') return false;
    return true;
  });
  _emApplyGlobalFilters();
}

// Aplica los filtros globales (negocio, ubicaciónV3, bodega, serial, referencia)
function _emApplyGlobalFilters() {
  const raw = _emGetRaw() || [];
  const bodegas = (typeof INV_BODEGAS!=='undefined'?INV_BODEGAS:null)||window.INV_BODEGAS;

  // Reconstruir DF_ALL con filtros globales aplicados
  let dfBase = raw.filter(r => EM_REFS_DATAFONO.has((r['Nombre']||'').trim()) && bodegas && bodegas.has((r['Nombre de la ubicación']||'').trim()));
  // Reconstruir SIM_ALL con filtros globales aplicados (excepto referencia fija)
  let simBase = raw.filter(r => {
    if ((r['Nombre']||'').trim() !== EM_REF_SIMCARD) return false;
    const serial = (r['Número de serie']||r['Numero de serie']||r['Serial']||'').trim().toUpperCase();
    if (EM_SIM_PRUEBA.has(serial)) return false;
    if (/^2\.?34234[e\+]?/i.test(serial) || serial === '234234000000000' || serial === '23423400000000') return false;
    return true;
  });

  // Función de filtro común
  function applyFilters(arr) {
    return arr.filter(r => {
      if (EM_F_NEGOCIO) {
        const neg = (typeof invNegocio === 'function') ? invNegocio(r['Subtipo']) : (()=>{const s=(r['Subtipo']||'').trim().toUpperCase();return(s==='WOMPI VP'||s==='EQUIPO VP'||s==='VP')?'VP':'CB';})();
        if (neg !== EM_F_NEGOCIO) return false;
      }
      if (EM_F_UBICV3) {
        const ubv3 = (typeof invUbicacionV3 === 'function') ? invUbicacionV3(r) : '';
        if (ubv3 !== EM_F_UBICV3) return false;
      }
      if (EM_F_BODEGA) {
        if ((r['Nombre de la ubicación']||'').trim() !== EM_F_BODEGA) return false;
      }
      if (EM_F_SERIAL) {
        const s = (r['Número de serie']||r['Numero de serie']||r['Serial']||'').trim();
        if (s !== EM_F_SERIAL) return false;
      }
      if (EM_F_REF) {
        if ((r['Nombre']||'').trim() !== EM_F_REF) return false;
      }
      return true;
    });
  }

  EM_DF_ALL  = applyFilters(dfBase);
  EM_SIM_ALL = applyFilters(simBase);
  EM_DF_FILTERED  = EM_DF_ALL.slice();
  EM_SIM_FILTERED = EM_SIM_ALL.slice();
}

// ══════════════════════════════════════════════════════════════════
//  RENDER DATÁFONOS
// ══════════════════════════════════════════════════════════════════

function _emRenderDatafonos() {
  const data = EM_DF_ALL;

  // Segmentar
  const seg = { DISPONIBLE:[], ASOCIADO:[], DAÑADO:[], 'DES. INCIDENTE':[], 'DES. CIERRE':[], 'BAJAS WOMPI':[] };
  data.forEach(r => { const e=_emEstadoDatafono(r); if(!seg[e]) seg[e]=[]; seg[e].push(r); });
  const total = data.length;

  // KPIs
  // "Total en Bodegas" excluye ALMACEN BAJAS WOMPI y ALMACEN INGENICO - PROVEEDOR WOMPI
  const dataParaTotal = data.filter(r => {
    const ubi = (r['Nombre de la ubicación']||'').trim();
    for (const excl of EM_BODEGAS_EXCLUIR_TOTAL) {
      if (ubi.toUpperCase().includes(excl.toUpperCase())) return false;
    }
    return true;
  });
  _EM_KPI_ROWS = [
    { title:'Total en Bodegas',  rows:dataParaTotal,           color:'#DFFF61', icon:'📦' },
    { title:'Disponibles',       rows:seg['DISPONIBLE']||[],   color:'#B0F2AE', icon:'✅' },
    { title:'Asociados',         rows:seg['ASOCIADO']||[],     color:'#C084FC', icon:'🔗' },
    { title:'En Daño',           rows:seg['DAÑADO']||[],       color:'#FF5C5C', icon:'🔴' },
    { title:'Des. Incidente',    rows:seg['DES. INCIDENTE']||[], color:'#FFC04D', icon:'⚠️' },
    { title:'Des. Cierre',       rows:seg['DES. CIERRE']||[],  color:'#99D1FC', icon:'🔵' },
    { title:'Bajas Wompi',       rows:seg['BAJAS WOMPI']||[],  color:'#F49D6E', icon:'⬇️' },
  ];
  window._EM_KPI_ROWS = _EM_KPI_ROWS;

  const kpiEl = document.getElementById('em-df-kpis');
  if (kpiEl) {
    kpiEl.innerHTML = _EM_KPI_ROWS.map((k,idx)=>{
      const n=k.rows.length;
      const pct=total?(n/total*100).toFixed(1):'0.0';
      return `<div onclick="window._emKpiClick(${idx})"
        style="background:linear-gradient(145deg,rgba(10,26,18,.95),rgba(8,20,14,.9));
          border:1px solid ${k.color}22;border-top:2px solid ${k.color};border-radius:14px;
          padding:18px 20px;min-width:130px;flex:1;cursor:pointer;position:relative;overflow:hidden;
          box-shadow:0 4px 20px rgba(0,0,0,.35);transition:all .22s cubic-bezier(.4,0,.2,1);"
        onmouseover="this.style.transform='translateY(-4px)';this.style.boxShadow='0 12px 36px rgba(0,0,0,.55),0 0 24px ${k.color}22'"
        onmouseout="this.style.transform='';this.style.boxShadow='0 4px 20px rgba(0,0,0,.35)'">
        <div style="position:absolute;top:0;left:0;right:0;height:100%;background:radial-gradient(ellipse at top right,${k.color}08,transparent 70%);pointer-events:none;"></div>
        <div style="position:absolute;top:10px;right:11px;font-size:9px;color:${k.color};opacity:.55;font-family:'Outfit',sans-serif;font-weight:600;letter-spacing:.5px;">VER ›</div>
        <div style="font-size:20px;margin-bottom:8px;">${k.icon}</div>
        <div style="font-family:'Syne',sans-serif;font-size:10px;font-weight:700;color:#7A7674;text-transform:uppercase;letter-spacing:.8px;margin-bottom:8px;padding-right:20px;line-height:1.3;">${k.title}</div>
        <div style="font-family:'JetBrains Mono',monospace;font-size:28px;font-weight:700;color:${k.color};line-height:1;text-shadow:0 0 24px ${k.color}55;margin-bottom:4px;">${n.toLocaleString('es-CO')}</div>
        <div style="padding-top:10px;border-top:1px solid rgba(255,255,255,.06);">
          <span style="background:${k.color}18;color:${k.color};font-size:10px;font-weight:700;padding:2px 9px;border-radius:20px;font-family:'JetBrains Mono',monospace;">${pct}% del total</span>
        </div>
      </div>`;
    }).join('');
  }

  // Dona: posición en depósito — ahora EN DAÑO viene directamente del campo
  const depMap={'EN DAÑO':0,'DESINSTALADO-INCIDENTE':0,'DESINSTALADO-CIERRE':0};
  data.forEach(r=>{const p=(r['Posición en depósito']||'').trim().toUpperCase();if(depMap[p]!==undefined)depMap[p]++;});
  _emDonut('em-chart-deposito',Object.keys(depMap),Object.values(depMap),[EM_COLORS.danio,EM_COLORS.incidente,EM_COLORS.cierre]);

  // Barras agrupadas: referencia × posición depósito
  const refS={};[...EM_REFS_DATAFONO].forEach(r=>{refS[r]={'EN DAÑO':0,'DESINSTALADO-INCIDENTE':0,'DESINSTALADO-CIERRE':0};});
  data.forEach(r=>{const n=(r['Nombre']||'').trim();const p=(r['Posición en depósito']||'').trim().toUpperCase();if(refS[n]&&depMap[p]!==undefined)refS[n][p]++;});
  const _ab=n=>n.replace('DATAFONO ','').replace(' - INGENICO','').replace('PORTATIL','PORT.').replace('ESCRITORIO','ESC.').replace('PINPAD ','');
  _emBarGrouped('em-chart-ref-estado',[...EM_REFS_DATAFONO].map(_ab),[
    {label:'En Daño (Dep.)',   data:[...EM_REFS_DATAFONO].map(r=>refS[r]['EN DAÑO']),               backgroundColor:EM_COLORS.danio+'BB',     borderColor:EM_COLORS.danio,     borderWidth:1,borderRadius:5},
    {label:'Des. Incidente',   data:[...EM_REFS_DATAFONO].map(r=>refS[r]['DESINSTALADO-INCIDENTE']),backgroundColor:EM_COLORS.incidente+'BB', borderColor:EM_COLORS.incidente, borderWidth:1,borderRadius:5},
    {label:'Des. Cierre',      data:[...EM_REFS_DATAFONO].map(r=>refS[r]['DESINSTALADO-CIERRE']),   backgroundColor:EM_COLORS.cierre+'BB',    borderColor:EM_COLORS.cierre,    borderWidth:1,borderRadius:5},
  ]);

  // Barras tipos de daño — de comentarios (top 15, resto agrupa en "Otros")
  const danioMap={};
  data.forEach(r=>{const t=_emTipoDanio(r);if(t)danioMap[t]=(danioMap[t]||0)+1;});
  const danioSorted=Object.entries(danioMap).sort((a,b)=>b[1]-a[1]);
  const cv=document.getElementById('em-chart-tipos-danio');
  if(danioSorted.length){
    const TOP_N=15;
    let labels, values;
    if(danioSorted.length<=TOP_N){
      labels=danioSorted.map(d=>d[0]);
      values=danioSorted.map(d=>d[1]);
    } else {
      const top=danioSorted.slice(0,TOP_N);
      const otrosSum=danioSorted.slice(TOP_N).reduce((s,[,v])=>s+v,0);
      labels=[...top.map(d=>d[0]),'OTROS ('+danioSorted.length+'+ tipos)'];
      values=[...top.map(d=>d[1]),otrosSum];
    }
    // Truncar etiquetas largas
    const labelsShort=labels.map(l=>l.length>35?l.substring(0,35)+'…':l);
    const colors=values.map((_,i)=>EM_COLORS.palette[i%EM_COLORS.palette.length]+'BB');
    // Calcular altura dinámica según número de barras
    const chartH=Math.max(240, labels.length*32+60);
    if(cv&&cv.parentElement){
      cv.parentElement.style.height=chartH+'px';
      cv.style.height='100%';
    }
    _emDestroyChart('em-chart-tipos-danio');
    if(cv){
      EM_CHARTS['em-chart-tipos-danio']=new Chart(cv,{
        type:'bar',
        data:{labels:labelsShort,datasets:[{data:values,backgroundColor:colors,
          borderColor:values.map((_,i)=>EM_COLORS.palette[i%EM_COLORS.palette.length]),
          borderWidth:1,borderRadius:5}]},
        options:{
          indexAxis:'y',responsive:true,maintainAspectRatio:false,
          plugins:{
            legend:{display:false},
            tooltip:{..._EM_TOOLTIP,callbacks:{label:ctx=>` ${ctx.parsed.x.toLocaleString('es-CO')} datáfonos`}}
          },
          scales:{
            x:{ticks:{color:'#7A7674',font:{size:10}},grid:{color:'rgba(255,255,255,.05)'}},
            y:{ticks:{color:'#e2e8f0',font:{size:10,family:"'Outfit',sans-serif"}},grid:{display:false},
               afterFit:ctx=>{ctx.width=Math.min(ctx.width,280);}}
          }
        }
      });
    }
  } else {
    if(cv&&cv.parentElement){
      cv.style.display='none';
      const msg=document.createElement('div');
      msg.style.cssText='text-align:center;padding:48px;color:#7A7674;font-family:\'Outfit\',sans-serif;font-size:13px;';
      msg.textContent='Sin registros de daño en comentarios';
      cv.parentElement.appendChild(msg);
    }
  }

  // Tabla bodegas — usa _emEstadoDatafono (lógica correcta), incluye BAJAS WOMPI
  const bodMap={};
  data.forEach(r=>{const b=(r['Nombre de la ubicación']||'Sin bodega').trim();const e=_emEstadoDatafono(r);if(!bodMap[b])bodMap[b]={DISPONIBLE:0,ASOCIADO:0,DAÑADO:0,'DES. CIERRE':0,'DES. INCIDENTE':0,'BAJAS WOMPI':0};if(bodMap[b][e]!==undefined)bodMap[b][e]++;else bodMap[b][e]=1;});
  const bodRows=Object.entries(bodMap).map(([nombre,c])=>({nombre,...c,total:Object.values(c).reduce((a,b)=>a+b,0)})).sort((a,b)=>b.total-a.total);
  // Totales globales
  const bodTotales={DISPONIBLE:0,ASOCIADO:0,DAÑADO:0,'DES. CIERRE':0,'DES. INCIDENTE':0,'BAJAS WOMPI':0,total:0};
  bodRows.forEach(b=>{bodTotales.DISPONIBLE+=b['DISPONIBLE']||0;bodTotales.ASOCIADO+=b['ASOCIADO']||0;bodTotales.DAÑADO+=b['DAÑADO']||0;bodTotales['DES. CIERRE']+=b['DES. CIERRE']||0;bodTotales['DES. INCIDENTE']+=b['DES. INCIDENTE']||0;bodTotales['BAJAS WOMPI']+=b['BAJAS WOMPI']||0;bodTotales.total+=b.total;});
  const bodEl=document.getElementById('em-tabla-bodegas-wrap');
  if(bodEl){
    if(!bodRows.length){bodEl.innerHTML='<div style="text-align:center;padding:32px;color:#7A7674;font-family:\'Outfit\',sans-serif;">Sin datos</div>';}
    else{
      const _p=(n,c)=>`<span style="background:${c}22;color:${c};font-family:'JetBrains Mono',monospace;font-size:12px;font-weight:700;padding:3px 10px;border-radius:10px;display:inline-block;min-width:28px;text-align:center;">${n}</span>`;
      bodEl.innerHTML=`<table style="width:100%;border-collapse:collapse;font-size:12px;">
        <thead><tr style="background:rgba(0,0,0,.3);position:sticky;top:0;z-index:2;">
          ${['Bodega','Disponible','Asociado','Dañado','Des. Cierre','Des. Incidente','Bajas Wompi','Total'].map((h,hi)=>`<th style="padding:10px 14px;text-align:${hi===0?'left':'center'};font-family:'Syne',sans-serif;font-size:9px;font-weight:700;color:#B0F2AE;letter-spacing:1px;text-transform:uppercase;border-bottom:1px solid rgba(176,242,174,.15);white-space:nowrap;">${h}</th>`).join('')}
        </tr></thead>
        <tbody>${bodRows.map((b,i)=>{const bg=i%2===0?'rgba(176,242,174,.012)':'transparent';return `<tr style="background:${bg}" onmouseover="this.style.background='rgba(176,242,174,.04)'" onmouseout="this.style.background='${bg}'">
          <td style="padding:9px 14px;color:#e2e8f0;font-size:11px;max-width:260px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="${b.nombre}">${b.nombre}</td>
          <td style="padding:9px 14px;text-align:center;">${_p(b['DISPONIBLE']||0,'#B0F2AE')}</td>
          <td style="padding:9px 14px;text-align:center;">${_p(b['ASOCIADO']||0,'#C084FC')}</td>
          <td style="padding:9px 14px;text-align:center;">${_p(b['DAÑADO']||0,'#FF5C5C')}</td>
          <td style="padding:9px 14px;text-align:center;">${_p(b['DES. CIERRE']||0,'#99D1FC')}</td>
          <td style="padding:9px 14px;text-align:center;">${_p(b['DES. INCIDENTE']||0,'#FFC04D')}</td>
          <td style="padding:9px 14px;text-align:center;">${_p(b['BAJAS WOMPI']||0,'#F49D6E')}</td>
          <td style="padding:9px 14px;text-align:center;"><span style="font-family:'JetBrains Mono',monospace;font-weight:700;color:#DFFF61;font-size:13px;">${b.total}</span></td>
        </tr>`;}).join('')}
        <tr style="background:rgba(223,255,97,.05);border-top:2px solid rgba(223,255,97,.2);">
          <td style="padding:10px 14px;font-family:'Syne',sans-serif;font-size:10px;font-weight:700;color:#DFFF61;text-transform:uppercase;letter-spacing:.5px;">TOTAL</td>
          <td style="padding:10px 14px;text-align:center;">${_p(bodTotales.DISPONIBLE,'#B0F2AE')}</td>
          <td style="padding:10px 14px;text-align:center;">${_p(bodTotales.ASOCIADO,'#C084FC')}</td>
          <td style="padding:10px 14px;text-align:center;">${_p(bodTotales.DAÑADO,'#FF5C5C')}</td>
          <td style="padding:10px 14px;text-align:center;">${_p(bodTotales['DES. CIERRE'],'#99D1FC')}</td>
          <td style="padding:10px 14px;text-align:center;">${_p(bodTotales['DES. INCIDENTE'],'#FFC04D')}</td>
          <td style="padding:10px 14px;text-align:center;">${_p(bodTotales['BAJAS WOMPI'],'#F49D6E')}</td>
          <td style="padding:10px 14px;text-align:center;"><span style="font-family:'JetBrains Mono',monospace;font-weight:700;color:#DFFF61;font-size:14px;">${bodTotales.total}</span></td>
        </tr>
        </tbody></table>`;
    }
  }

  _emApplyDfSearch('');
}

window._emKpiClick = function(idx) {
  const k=(window._EM_KPI_ROWS||[])[idx];
  if(k) _emOpenModal(k.title, k.rows, k.color);
};

// ══════════════════════════════════════════════════════════════════
//  TABLA DATÁFONOS
// ══════════════════════════════════════════════════════════════════

function _emApplyDfSearch(query) {
  EM_DF_SEARCH=(query||'').toLowerCase().trim();
  EM_DF_FILTERED=EM_DF_ALL.filter(r=>{
    if(!EM_DF_SEARCH) return true;
    const s=(r['Número de serie']||r['Numero de serie']||r['Serial']||'').toLowerCase();
    const n=(r['Nombre']||'').toLowerCase();
    const u=(r['Nombre de la ubicación']||'').toLowerCase();
    const p=(r['Posición en depósito']||'').toLowerCase();
    const c=(r['Comentarios']||'').toLowerCase();
    const e=_emEstadoDatafono(r).toLowerCase();
    return s.includes(EM_DF_SEARCH)||n.includes(EM_DF_SEARCH)||u.includes(EM_DF_SEARCH)||p.includes(EM_DF_SEARCH)||c.includes(EM_DF_SEARCH)||e.includes(EM_DF_SEARCH);
  });
  EM_DF_PAGE=1; _emRenderDfTabla();
}

function _emRenderDfTabla() {
  const wrap=document.getElementById('em-df-tabla-wrap');
  const cnt=document.getElementById('em-df-tabla-count');
  const pag=document.getElementById('em-df-tabla-pag');
  if(!wrap) return;
  const data=EM_DF_FILTERED;
  const pages=Math.max(1,Math.ceil(data.length/EM_PAGE_SIZE));
  const slice=data.slice((EM_DF_PAGE-1)*EM_PAGE_SIZE,EM_DF_PAGE*EM_PAGE_SIZE);
  if(cnt) cnt.textContent=`${data.length.toLocaleString('es-CO')} registros`;
  const _ep=est=>{const m={'DISPONIBLE':['#B0F2AE','rgba(176,242,174,.12)'],'ASOCIADO':['#C084FC','rgba(192,132,252,.12)'],'DAÑADO':['#FF5C5C','rgba(255,92,92,.12)'],'DES. INCIDENTE':['#FFC04D','rgba(255,192,77,.12)'],'DES. CIERRE':['#99D1FC','rgba(153,209,252,.12)'],'BAJAS WOMPI':['#F49D6E','rgba(244,157,110,.12)']};const[col,bg]=m[est]||['#7A7674','rgba(255,255,255,.06)'];return `<span style="background:${bg};color:${col};font-size:10px;font-weight:700;padding:3px 10px;border-radius:12px;border:1px solid ${col}44;display:inline-block;font-family:'Outfit',sans-serif;white-space:nowrap;">${est}</span>`;};
  if(!slice.length){wrap.innerHTML='<div style="text-align:center;padding:40px;color:#7A7674;font-family:\'Outfit\',sans-serif;">Sin resultados</div>';if(pag)pag.innerHTML='';return;}
  const COLS=[{h:'Serial',fn:r=>r['Número de serie']||r['Numero de serie']||r['Serial']||'—',mono:true,col:'#a5f3fc'},{h:'Referencia',fn:r=>r['Nombre']||'—',col:'#e2e8f0'},{h:'Bodega',fn:r=>r['Nombre de la ubicación']||'—',col:'#cbd5e1'},{h:'Pos. Depósito',fn:r=>r['Posición en depósito']||'—',col:'#94a3b8'},{h:'Estado',fn:r=>_emEstadoDatafono(r),isEstado:true},{h:'Comentarios',fn:r=>r['Comentarios']||'—',col:'#7A7674'}];
  wrap.innerHTML=`<table style="width:100%;border-collapse:collapse;font-size:12px;font-family:'Outfit',sans-serif;">
    <thead><tr style="background:#181715;position:sticky;top:0;z-index:2;">
      ${COLS.map(c=>`<th style="padding:10px 14px;text-align:left;font-family:'Syne',sans-serif;font-size:9px;font-weight:700;color:#B0F2AE;letter-spacing:1px;text-transform:uppercase;border-bottom:1px solid rgba(176,242,174,.15);white-space:nowrap;">${c.h}</th>`).join('')}
    </tr></thead>
    <tbody>${slice.map((r,i)=>{const bg=i%2===0?'rgba(176,242,174,.012)':'transparent';return`<tr style="background:${bg}" onmouseover="this.style.background='rgba(176,242,174,.04)'" onmouseout="this.style.background='${bg}'">
      ${COLS.map(c=>{const v=c.fn(r);if(c.isEstado)return`<td style="padding:9px 14px;">${_ep(v)}</td>`;return`<td style="padding:9px 14px;${c.mono?`font-family:'JetBrains Mono',monospace;font-size:11px;`:'font-size:11px;'}color:${c.col||'#e2e8f0'};max-width:220px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="${String(v).replace(/"/g,'&quot;')}">${v}</td>`;}).join('')}
    </tr>`;}).join('')}
    </tbody></table>`;
  _emRenderPag(pag,EM_DF_PAGE,pages,p=>{EM_DF_PAGE=p;_emRenderDfTabla();},'#B0F2AE');
}

// ══════════════════════════════════════════════════════════════════
//  RENDER SIMCARDS
// ══════════════════════════════════════════════════════════════════

function _emRenderSimcards() {
  const data=EM_SIM_ALL;
  const activas=data.filter(r=>_emSimActivada(r));
  const sinAct=data.filter(r=>!_emSimActivada(r));
  const total=data.length;

  // KPIs SIM
  const simKpis=[
    {title:'Total SIMCards',rows:data,    color:'#DFFF61',icon:'📡'},
    {title:'Activadas',     rows:activas, color:'#B0F2AE',icon:'✅'},
    {title:'Sin Activar',   rows:sinAct,  color:'#FF5C5C',icon:'❌'},
  ];
  window._EM_SIM_KPI_ROWS=simKpis;

  const kpiEl=document.getElementById('em-sim-kpis');
  if(kpiEl){
    kpiEl.innerHTML=[...simKpis.map((k,idx)=>{
      const n=k.rows.length,pct=total?(n/total*100).toFixed(1):'0.0';
      return `<div onclick="window._emSimKpiClick(${idx})" style="background:linear-gradient(145deg,rgba(10,26,18,.95),rgba(8,20,14,.9));border:1px solid ${k.color}22;border-top:2px solid ${k.color};border-radius:14px;padding:18px 20px;min-width:130px;flex:1;cursor:pointer;position:relative;overflow:hidden;box-shadow:0 4px 20px rgba(0,0,0,.35);transition:all .22s cubic-bezier(.4,0,.2,1);" onmouseover="this.style.transform='translateY(-4px)';this.style.boxShadow='0 12px 36px rgba(0,0,0,.55)'" onmouseout="this.style.transform='';this.style.boxShadow='0 4px 20px rgba(0,0,0,.35)'">
        <div style="position:absolute;top:10px;right:11px;font-size:9px;color:${k.color};opacity:.55;font-family:'Outfit',sans-serif;font-weight:600;letter-spacing:.5px;">VER ›</div>
        <div style="font-size:20px;margin-bottom:8px;">${k.icon}</div>
        <div style="font-family:'Syne',sans-serif;font-size:10px;font-weight:700;color:#7A7674;text-transform:uppercase;letter-spacing:.8px;margin-bottom:8px;">${k.title}</div>
        <div style="font-family:'JetBrains Mono',monospace;font-size:28px;font-weight:700;color:${k.color};line-height:1;text-shadow:0 0 24px ${k.color}55;margin-bottom:4px;">${n.toLocaleString('es-CO')}</div>
        <div style="padding-top:10px;border-top:1px solid rgba(255,255,255,.06);">
          <span style="background:${k.color}18;color:${k.color};font-size:10px;font-weight:700;padding:2px 9px;border-radius:20px;font-family:'JetBrains Mono',monospace;">${pct}% del total</span>
        </div>
      </div>`;
    }),
    `<div style="background:linear-gradient(145deg,rgba(10,26,18,.95),rgba(8,20,14,.9));border:1px solid rgba(153,209,252,.22);border-top:2px solid #99D1FC;border-radius:14px;padding:18px 20px;min-width:130px;flex:1;box-shadow:0 4px 20px rgba(0,0,0,.35);">
      <div style="font-size:20px;margin-bottom:8px;">📊</div>
      <div style="font-family:'Syne',sans-serif;font-size:10px;font-weight:700;color:#7A7674;text-transform:uppercase;letter-spacing:.8px;margin-bottom:8px;">% Activación</div>
      <div style="font-family:'JetBrains Mono',monospace;font-size:28px;font-weight:700;color:#99D1FC;line-height:1;text-shadow:0 0 24px #99D1FC55;margin-bottom:4px;">${total?Math.round(activas.length/total*100):0}%</div>
      <div style="padding-top:10px;border-top:1px solid rgba(255,255,255,.06);"><span style="background:rgba(153,209,252,.12);color:#99D1FC;font-size:10px;font-weight:700;padding:2px 9px;border-radius:20px;font-family:'JetBrains Mono',monospace;">${activas.length} / ${total}</span></div>
    </div>`].join('');
  }
  window._emSimKpiClick=idx=>{const k=(window._EM_SIM_KPI_ROWS||[])[idx];if(k)_emOpenSimModal(k.title,k.rows,k.color);};

  _emDonut('em-chart-sim-estado',['Activadas','Sin Activar'],[activas.length,sinAct.length],[EM_COLORS.activada,EM_COLORS.sinActivar]);

  const sitioActMap={};activas.forEach(r=>{const s=(r['Nombre de la ubicación']||'Sin ubicación').trim();sitioActMap[s]=(sitioActMap[s]||0)+(parseInt(r['Cantidad'])||1);});
  const topAct=Object.entries(sitioActMap).sort((a,b)=>b[1]-a[1]).slice(0,12);
  _emBarH('em-chart-sim-activadas',topAct.map(s=>s[0].length>28?s[0].substring(0,28)+'…':s[0]),topAct.map(s=>s[1]),topAct.map(()=>'#B0F2AEBB'));

  const sitMap={};data.forEach(r=>{const s=(r['Nombre de la ubicación']||'Sin ubicación').trim();sitMap[s]=(sitMap[s]||0)+(parseInt(r['Cantidad'])||1);});
  const topSit=Object.entries(sitMap).sort((a,b)=>b[1]-a[1]).slice(0,15);
  _emBarH('em-chart-sim-sitios',topSit.map(s=>s[0].length>28?s[0].substring(0,28)+'…':s[0]),topSit.map(s=>s[1]),null);

  const mesMap={};activas.forEach(r=>{const fa=_emSimFechaActivacion(r);if(fa==='—')return;const pts=fa.split('/');if(pts.length<3)return;const k=`${pts[2]}-${pts[1]}`;mesMap[k]=(mesMap[k]||0)+1;});
  const mesSorted=Object.entries(mesMap).sort((a,b)=>a[0].localeCompare(b[0]));
  if(mesSorted.length){
    const c=document.getElementById('em-chart-sim-mensual');
    if(c){_emDestroyChart('em-chart-sim-mensual');EM_CHARTS['em-chart-sim-mensual']=new Chart(c,{type:'bar',data:{labels:mesSorted.map(([k])=>{const[y,m]=k.split('-');return`${m}/${y}`;}),datasets:[{data:mesSorted.map(([,v])=>v),backgroundColor:'#B0F2AEAA',borderColor:'#B0F2AE',borderWidth:1,borderRadius:6}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:_EM_TOOLTIP},scales:{x:{ticks:{color:'#cbd5e1',font:{size:10},maxRotation:30},grid:{display:false}},y:{ticks:{color:'#7A7674'},grid:{color:'rgba(255,255,255,.05)'}}}}}); }
  }

  // Tabla por sitio/comercio: Nombre ubicación · Activas · Inactivas · Total
  _emRenderSimPorSitio();

  _emApplySimSearch('');
}

// ══════════════════════════════════════════════════════════════════
//  TABLA SIMCARDS POR SITIO/COMERCIO
// ══════════════════════════════════════════════════════════════════
function _emRenderSimPorSitio() {
  const wrap = document.getElementById('em-sim-por-sitio-wrap');
  if (!wrap) return;
  const data = EM_SIM_ALL;

  // Agrupar por nombre de ubicación
  const sitioMap = {};
  data.forEach(r => {
    const s = (r['Nombre de la ubicación'] || 'Sin ubicación').trim();
    if (!sitioMap[s]) sitioMap[s] = { activas: 0, inactivas: 0 };
    if (_emSimActivada(r)) sitioMap[s].activas++;
    else                   sitioMap[s].inactivas++;
  });

  const rows = Object.entries(sitioMap)
    .map(([nombre, c]) => ({ nombre, activas: c.activas, inactivas: c.inactivas, total: c.activas + c.inactivas }))
    .sort((a, b) => b.total - a.total);

  // Totales globales
  const totActivas   = rows.reduce((s, r) => s + r.activas,   0);
  const totInactivas = rows.reduce((s, r) => s + r.inactivas, 0);
  const totTotal     = totActivas + totInactivas;

  // Actualizar KPI de sumatoria en el encabezado de la tabla
  const sumEl = document.getElementById('em-sim-sitio-suma');
  if (sumEl) {
    sumEl.innerHTML = `
      <span style="margin-right:18px;"><span style="color:#B0F2AE;font-family:'JetBrains Mono',monospace;font-weight:700;">${totActivas.toLocaleString('es-CO')}</span><span style="color:#64748b;font-size:10px;"> activadas</span></span>
      <span style="margin-right:18px;"><span style="color:#FF5C5C;font-family:'JetBrains Mono',monospace;font-weight:700;">${totInactivas.toLocaleString('es-CO')}</span><span style="color:#64748b;font-size:10px;"> sin activar</span></span>
      <span><span style="color:#DFFF61;font-family:'JetBrains Mono',monospace;font-weight:700;">${totTotal.toLocaleString('es-CO')}</span><span style="color:#64748b;font-size:10px;"> total</span></span>`;
  }

  if (!rows.length) {
    wrap.innerHTML = '<div style="text-align:center;padding:32px;color:#7A7674;font-family:\'Outfit\',sans-serif;">Sin datos</div>';
    return;
  }

  const _p = (n, c) => `<span style="background:${c}22;color:${c};font-family:'JetBrains Mono',monospace;font-size:12px;font-weight:700;padding:3px 10px;border-radius:10px;display:inline-block;min-width:28px;text-align:center;">${n.toLocaleString('es-CO')}</span>`;

  wrap.innerHTML = `<table style="width:100%;border-collapse:collapse;font-size:12px;">
    <thead><tr style="background:rgba(0,0,0,.3);position:sticky;top:0;z-index:2;">
      ${['Comercio / Ubicación','Activadas','Sin Activar','Total'].map((h,hi)=>`<th style="padding:10px 14px;text-align:${hi===0?'left':'center'};font-family:'Syne',sans-serif;font-size:9px;font-weight:700;color:#99D1FC;letter-spacing:1px;text-transform:uppercase;border-bottom:1px solid rgba(153,209,252,.15);white-space:nowrap;">${h}</th>`).join('')}
    </tr></thead>
    <tbody>
      ${rows.map((r, i) => {
        const bg = i % 2 === 0 ? 'rgba(153,209,252,.012)' : 'transparent';
        const pct = r.total ? Math.round(r.activas / r.total * 100) : 0;
        return `<tr style="background:${bg}" onmouseover="this.style.background='rgba(153,209,252,.04)'" onmouseout="this.style.background='${bg}'">
          <td style="padding:9px 14px;color:#cbd5e1;font-size:11px;max-width:280px;">
            <div style="display:flex;flex-direction:column;gap:4px;">
              <span style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="${r.nombre}">${r.nombre}</span>
              <div style="height:3px;background:rgba(255,255,255,.08);border-radius:2px;overflow:hidden;width:120px;">
                <div style="height:100%;width:${pct}%;background:#B0F2AE;border-radius:2px;"></div>
              </div>
            </div>
          </td>
          <td style="padding:9px 14px;text-align:center;">${_p(r.activas, '#B0F2AE')}</td>
          <td style="padding:9px 14px;text-align:center;">${_p(r.inactivas, '#FF5C5C')}</td>
          <td style="padding:9px 14px;text-align:center;"><span style="font-family:'JetBrains Mono',monospace;font-weight:700;color:#DFFF61;font-size:13px;">${r.total.toLocaleString('es-CO')}</span></td>
        </tr>`;
      }).join('')}
      <tr style="background:rgba(153,209,252,.05);border-top:2px solid rgba(153,209,252,.2);">
        <td style="padding:10px 14px;font-family:'Syne',sans-serif;font-size:10px;font-weight:700;color:#99D1FC;text-transform:uppercase;letter-spacing:.5px;">TOTAL (${rows.length} puntos)</td>
        <td style="padding:10px 14px;text-align:center;">${_p(totActivas, '#B0F2AE')}</td>
        <td style="padding:10px 14px;text-align:center;">${_p(totInactivas, '#FF5C5C')}</td>
        <td style="padding:10px 14px;text-align:center;"><span style="font-family:'JetBrains Mono',monospace;font-weight:700;color:#DFFF61;font-size:14px;">${totTotal.toLocaleString('es-CO')}</span></td>
      </tr>
    </tbody>
  </table>`;
}

// ══════════════════════════════════════════════════════════════════
//  TABLA SIMCARDS
// ══════════════════════════════════════════════════════════════════

function _emApplySimSearch(query) {
  EM_SIM_SEARCH=(query||'').toLowerCase().trim();
  EM_SIM_FILTERED=EM_SIM_ALL.filter(r=>{if(!EM_SIM_SEARCH)return true;const s=(r['Número de serie']||r['Numero de serie']||r['Serial']||'').toLowerCase();const u=(r['Nombre de la ubicación']||'').toLowerCase();const a=(r['Atributos']||'').toLowerCase();const c=(r['Código de ubicación']||'').toLowerCase();return s.includes(EM_SIM_SEARCH)||u.includes(EM_SIM_SEARCH)||a.includes(EM_SIM_SEARCH)||c.includes(EM_SIM_SEARCH);});
  EM_SIM_PAGE=1;_emRenderSimTabla();
}

function _emRenderSimTabla() {
  const wrap=document.getElementById('em-sim-tabla-wrap');const cnt=document.getElementById('em-sim-tabla-count');const pag=document.getElementById('em-sim-tabla-pag');
  if(!wrap)return;
  const data=EM_SIM_FILTERED;const pages=Math.max(1,Math.ceil(data.length/EM_PAGE_SIZE));const slice=data.slice((EM_SIM_PAGE-1)*EM_PAGE_SIZE,EM_SIM_PAGE*EM_PAGE_SIZE);
  if(cnt)cnt.textContent=`${data.length.toLocaleString('es-CO')} registros`;
  if(!slice.length){wrap.innerHTML='<div style="text-align:center;padding:40px;color:#7A7674;font-family:\'Outfit\',sans-serif;">Sin resultados</div>';if(pag)pag.innerHTML='';return;}
  wrap.innerHTML=`<table style="width:100%;border-collapse:collapse;font-size:12px;font-family:'Outfit',sans-serif;">
    <thead><tr style="background:#181715;position:sticky;top:0;z-index:2;">
      ${['Serial','Ubicación','Cód. Ubic.','Estado','Fecha Activación','Atributos','Comentarios'].map(h=>`<th style="padding:10px 14px;text-align:left;font-family:'Syne',sans-serif;font-size:9px;font-weight:700;color:#99D1FC;letter-spacing:1px;text-transform:uppercase;border-bottom:1px solid rgba(153,209,252,.15);white-space:nowrap;">${h}</th>`).join('')}
    </tr></thead>
    <tbody>${slice.map((r,i)=>{const act=_emSimActivada(r);const bg=i%2===0?'rgba(153,209,252,.012)':'transparent';const pill=act?`<span style="background:rgba(176,242,174,.12);color:#B0F2AE;font-size:10px;font-weight:700;padding:3px 10px;border-radius:12px;border:1px solid rgba(176,242,174,.3);display:inline-block;font-family:'Outfit',sans-serif;">✅ ACTIVADA</span>`:`<span style="background:rgba(255,92,92,.12);color:#FF5C5C;font-size:10px;font-weight:700;padding:3px 10px;border-radius:12px;border:1px solid rgba(255,92,92,.3);display:inline-block;font-family:'Outfit',sans-serif;">❌ SIN ACTIVAR</span>`;
    return`<tr style="background:${bg}" onmouseover="this.style.background='rgba(153,209,252,.04)'" onmouseout="this.style.background='${bg}'">
      <td style="padding:9px 14px;font-family:'JetBrains Mono',monospace;font-size:11px;color:#a5f3fc;">${r['Número de serie']||r['Numero de serie']||r['Serial']||'—'}</td>
      <td style="padding:9px 14px;font-size:11px;color:#cbd5e1;max-width:220px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="${r['Nombre de la ubicación']||''}">${r['Nombre de la ubicación']||'—'}</td>
      <td style="padding:9px 14px;font-family:'JetBrains Mono',monospace;font-size:11px;color:#67e8f9;">${r['Código de ubicación']||'—'}</td>
      <td style="padding:9px 14px;">${pill}</td>
      <td style="padding:9px 14px;font-family:'JetBrains Mono',monospace;font-size:11px;color:${act?'#B0F2AE':'#64748b'};">${_emSimFechaActivacion(r)}</td>
      <td style="padding:9px 14px;font-size:11px;color:#94a3b8;max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="${r['Atributos']||''}">${r['Atributos']||'—'}</td>
      <td style="padding:9px 14px;font-size:11px;color:#7A7674;max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="${r['Comentarios']||''}">${r['Comentarios']||'—'}</td>
    </tr>`;}).join('')}
    </tbody></table>`;
  _emRenderPag(pag,EM_SIM_PAGE,pages,p=>{EM_SIM_PAGE=p;_emRenderSimTabla();},'#99D1FC');
}

// ══════════════════════════════════════════════════════════════════
//  PAGINACIÓN
// ══════════════════════════════════════════════════════════════════
function _emRenderPag(el, cur, total, onPage, accent) {
  if(!el)return;el.innerHTML='';if(total<=1)return;
  const bs=`padding:5px 11px;border-radius:7px;cursor:pointer;font-size:12px;font-family:'Outfit',sans-serif;border:1px solid rgba(255,255,255,.1);transition:all .15s;`;
  const mk=(l,p,a)=>{const b=document.createElement('button');b.style.cssText=bs+(a?`background:${accent};color:#0a1a12;font-weight:700;border-color:${accent};`:'background:rgba(255,255,255,.06);color:#94a3b8;');b.textContent=l;b.addEventListener('click',()=>onPage(p));return b;};
  if(cur>1)el.appendChild(mk('‹',cur-1,false));
  const s=Math.max(1,cur-2),e=Math.min(total,cur+2);
  for(let p=s;p<=e;p++)el.appendChild(mk(String(p),p,p===cur));
  if(cur<total)el.appendChild(mk('›',cur+1,false));
}

// ══════════════════════════════════════════════════════════════════
//  HTML PRINCIPAL
// ══════════════════════════════════════════════════════════════════
function _emBuildHTML() {
  const panel=document.getElementById('panel-estado-materiales');
  if(!panel) return;

  // Inyectar CSS para forzar modo oscuro en los <select> al desplegar opciones
  if (!document.getElementById('em-select-dark-style')) {
    const st = document.createElement('style');
    st.id = 'em-select-dark-style';
    st.textContent = `
      #em-f-negocio, #em-f-ubicv3 {
        background-color: #1a1f2e !important;
        color: #e2e8f0 !important;
        color-scheme: dark;
      }
      #em-f-negocio option, #em-f-ubicv3 option {
        background-color: #1a1f2e !important;
        color: #e2e8f0 !important;
      }
      #em-f-negocio:focus, #em-f-ubicv3:focus {
        background-color: #1a1f2e !important;
        color: #e2e8f0 !important;
      }
    `;
    document.head.appendChild(st);
  }

  const _cc=(title,sub,cid,h='260px',extra='')=>`
    <div style="background:linear-gradient(145deg,rgba(10,26,18,.95),rgba(8,20,14,.9));border:1px solid rgba(223,255,97,.1);border-radius:18px;padding:22px 24px 18px;position:relative;overflow:hidden;box-shadow:0 4px 20px rgba(0,0,0,.4);${extra}">
      <div style="position:absolute;top:0;left:0;right:0;height:2px;background:linear-gradient(90deg,#DFFF61,#B0F2AE,#99D1FC);opacity:.5;"></div>
      <div style="margin-bottom:14px;">
        <div style="font-family:'Syne',sans-serif;font-size:13px;font-weight:700;color:#f1f5f9;">${title}</div>
        ${sub?`<div style="font-size:11px;color:#7A7674;margin-top:3px;font-family:'Outfit',sans-serif;">${sub}</div>`:''}
      </div>
      <div style="position:relative;height:${h};"><canvas id="${cid}"></canvas></div>
    </div>`;

  panel.innerHTML=`<div style="padding:0 4px 40px;">

    <!-- ══ FILTROS GLOBALES ══ -->
    <div id="em-filtros-globales" style="background:linear-gradient(145deg,rgba(10,26,18,.97),rgba(8,20,14,.95));border:1px solid rgba(176,242,174,.18);border-radius:16px;padding:18px 22px;margin-bottom:28px;box-shadow:0 4px 24px rgba(0,0,0,.4);">
      <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:10px;margin-bottom:14px;">
        <div style="font-family:'Syne',sans-serif;font-size:12px;font-weight:700;color:#B0F2AE;letter-spacing:.8px;text-transform:uppercase;">🔽 Filtros — Aplican a todo: Datáfonos y SIMCards</div>
        <button onclick="emResetGlobalFilters()" style="background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.12);color:#94a3b8;padding:5px 14px;border-radius:8px;cursor:pointer;font-size:11px;font-family:'Outfit',sans-serif;transition:all .2s;" onmouseover="this.style.background='rgba(176,242,174,.1)';this.style.color='#B0F2AE'" onmouseout="this.style.background='rgba(255,255,255,.06)';this.style.color='#94a3b8'">↺ Limpiar filtros</button>
      </div>
      <div style="display:flex;flex-wrap:wrap;gap:12px;align-items:flex-end;">

        <!-- Tipo de Negocio -->
        <div style="display:flex;flex-direction:column;gap:5px;min-width:140px;">
          <label style="font-size:10px;font-weight:700;color:#7A7674;font-family:'Syne',sans-serif;letter-spacing:.5px;text-transform:uppercase;">Tipo Negocio</label>
          <select id="em-f-negocio" onchange="emApplyGlobalFilters()" style="background:#1a1f2e;border:1px solid rgba(176,242,174,.2);border-radius:8px;color:#e2e8f0;padding:7px 10px;font-size:12px;font-family:'Outfit',sans-serif;outline:none;cursor:pointer;color-scheme:dark;">
            <option value="">Todos</option>
            <option value="CB">CB</option>
            <option value="VP">VP</option>
          </select>
        </div>

        <!-- Ubicación V3 -->
        <div style="display:flex;flex-direction:column;gap:5px;min-width:180px;">
          <label style="font-size:10px;font-weight:700;color:#7A7674;font-family:'Syne',sans-serif;letter-spacing:.5px;text-transform:uppercase;">Ubicación (V3)</label>
          <select id="em-f-ubicv3" onchange="emApplyGlobalFilters()" style="background:#1a1f2e;border:1px solid rgba(176,242,174,.2);border-radius:8px;color:#e2e8f0;padding:7px 10px;font-size:12px;font-family:'Outfit',sans-serif;outline:none;cursor:pointer;color-scheme:dark;">
            <option value="">Todas</option>
            <option value="En bodega">En bodega</option>
            <option value="En corresponsal">En corresponsal</option>
            <option value="Gestor LineaCom">Gestor LineaCom</option>
            <option value="Gestor Wompi">Gestor Wompi</option>
            <option value="Empleados Wompi">Empleados Wompi</option>
            <option value="En operador Logistico">En operador Logístico</option>
            <option value="En distribución">En distribución</option>
            <option value="En Ingenico">En Ingenico</option>
          </select>
        </div>

        <!-- Bodega -->
        <div style="display:flex;flex-direction:column;gap:5px;min-width:200px;flex:1;">
          <label style="font-size:10px;font-weight:700;color:#7A7674;font-family:'Syne',sans-serif;letter-spacing:.5px;text-transform:uppercase;">Bodega</label>
          <div style="position:relative;">
            <input id="em-f-bodega" type="text" placeholder="Todas las bodegas..." autocomplete="off"
              style="width:100%;box-sizing:border-box;background:rgba(255,255,255,.06);border:1px solid rgba(176,242,174,.2);border-radius:8px;color:#e2e8f0;padding:7px 10px;font-size:12px;font-family:'Outfit',sans-serif;outline:none;"
              oninput="_emAcBodega(this.value)" onblur="_emAcBodegaHide()">
            <div id="em-f-bodega-ac" style="display:none;position:absolute;top:100%;left:0;right:0;background:#181715;border:1px solid rgba(176,242,174,.2);border-radius:0 0 8px 8px;max-height:200px;overflow-y:auto;z-index:100;font-size:12px;font-family:'Outfit',sans-serif;"></div>
          </div>
        </div>

        <!-- Número de Serie -->
        <div style="display:flex;flex-direction:column;gap:5px;min-width:160px;flex:1;">
          <label style="font-size:10px;font-weight:700;color:#7A7674;font-family:'Syne',sans-serif;letter-spacing:.5px;text-transform:uppercase;">Número de Serie</label>
          <input id="em-f-serial" type="text" placeholder="Buscar serial..." autocomplete="off"
            style="width:100%;box-sizing:border-box;background:rgba(255,255,255,.06);border:1px solid rgba(176,242,174,.2);border-radius:8px;color:#e2e8f0;padding:7px 10px;font-size:12px;font-family:'Outfit',sans-serif;outline:none;"
            oninput="_emGfSerialInput(this.value)">
        </div>

        <!-- Referencia -->
        <div style="display:flex;flex-direction:column;gap:5px;min-width:200px;flex:1;">
          <label style="font-size:10px;font-weight:700;color:#7A7674;font-family:'Syne',sans-serif;letter-spacing:.5px;text-transform:uppercase;">Referencia</label>
          <div style="position:relative;">
            <input id="em-f-ref" type="text" placeholder="Todas las referencias..." autocomplete="off"
              style="width:100%;box-sizing:border-box;background:rgba(255,255,255,.06);border:1px solid rgba(176,242,174,.2);border-radius:8px;color:#e2e8f0;padding:7px 10px;font-size:12px;font-family:'Outfit',sans-serif;outline:none;"
              oninput="_emAcRef(this.value)" onblur="_emAcRefHide()">
            <div id="em-f-ref-ac" style="display:none;position:absolute;top:100%;left:0;right:0;background:#181715;border:1px solid rgba(176,242,174,.2);border-radius:0 0 8px 8px;max-height:180px;overflow-y:auto;z-index:100;font-size:12px;font-family:'Outfit',sans-serif;"></div>
          </div>
        </div>

      </div>
      <div id="em-filtros-tag" style="margin-top:10px;display:flex;flex-wrap:wrap;gap:6px;min-height:20px;"></div>
    </div>

    <div class="section-label fade-up" style="color:#B0F2AE;">ESTADO DE DATÁFONOS EN BODEGA</div>

    <div id="em-df-kpis" style="display:flex;flex-wrap:wrap;gap:12px;margin-bottom:28px;">
      <div class="loading"><div class="spinner"></div><span>Cargando...</span></div>
    </div>

    <div style="display:grid;grid-template-columns:1fr 1fr;gap:18px;margin-bottom:20px;">
      ${_cc('Posición en Depósito','Solo estados: EN DAÑO · DES. INCIDENTE · DES. CIERRE','em-chart-deposito','240px')}
      ${_cc('Estados por Referencia (Posición Depósito)','Datáfonos por referencia y estado de depósito','em-chart-ref-estado','240px')}
    </div>

    <div style="margin-bottom:24px;">
      ${_cc('Tipos de Daño','Desde campo Comentarios: 99999 | TIPO (excluye DESASOCIADO y MIGRACION CB) · Top 15, resto agrupado','em-chart-tipos-danio','360px')}
    </div>

    <div class="section-label fade-up" style="color:#B0F2AE;">Estado por Bodega</div>
    <div style="background:linear-gradient(145deg,rgba(10,26,18,.95),rgba(8,20,14,.9));border:1px solid rgba(176,242,174,.12);border-radius:18px;overflow:hidden;margin-bottom:28px;box-shadow:0 4px 20px rgba(0,0,0,.4);">
      <div style="padding:14px 20px;border-bottom:1px solid rgba(176,242,174,.08);">
        <div style="font-family:'Syne',sans-serif;font-size:13px;font-weight:700;color:#f1f5f9;">Resumen por Bodega</div>
        <div style="font-size:11px;color:#7A7674;margin-top:2px;">Disponible · Asociado · Dañado · Des. Cierre · Des. Incidente</div>
      </div>
      <div id="em-tabla-bodegas-wrap" style="overflow-x:auto;max-height:480px;"></div>
    </div>

    <div class="section-label fade-up" style="color:#B0F2AE;">Tabla Completa — Datáfonos en Bodegas</div>
    <div style="background:linear-gradient(145deg,rgba(10,26,18,.95),rgba(8,20,14,.9));border:1px solid rgba(176,242,174,.12);border-radius:18px;overflow:hidden;margin-bottom:40px;box-shadow:0 4px 20px rgba(0,0,0,.4);">
      <div style="display:flex;align-items:center;justify-content:space-between;padding:16px 20px;border-bottom:1px solid rgba(176,242,174,.08);">
        <div>
          <div style="font-family:'Syne',sans-serif;font-size:13px;font-weight:700;color:#f1f5f9;">Serial · Referencia · Bodega · Estado · Comentarios</div>
          <div style="font-size:11px;color:#7A7674;margin-top:2px;" id="em-df-tabla-count">0 registros</div>
        </div>
        <button onclick="emExportDfExcel()" style="background:rgba(176,242,174,.08);border:1px solid rgba(176,242,174,.2);color:#B0F2AE;padding:7px 14px;border-radius:8px;cursor:pointer;font-size:12px;font-family:'Outfit',sans-serif;transition:all .2s;">⬇ Excel</button>
      </div>
      <div style="padding:10px 20px;border-bottom:1px solid rgba(176,242,174,.07);background:rgba(0,0,0,.15);">
        <input type="text" oninput="_emApplyDfSearch(this.value)" placeholder="🔍 Buscar por serial, referencia, bodega, estado, comentarios..."
          style="width:100%;box-sizing:border-box;background:rgba(255,255,255,.05);border:1px solid rgba(176,242,174,.2);border-radius:8px;color:#FAFAFA;padding:8px 14px;font-size:13px;font-family:'Outfit',sans-serif;outline:none;">
      </div>
      <div id="em-df-tabla-wrap" style="overflow-x:auto;max-height:520px;"></div>
      <div style="display:flex;justify-content:flex-end;padding:10px 20px;border-top:1px solid rgba(176,242,174,.07);">
        <div id="em-df-tabla-pag" style="display:flex;gap:6px;flex-wrap:wrap;"></div>
      </div>
    </div>

    <div class="section-label fade-up" style="color:#99D1FC;">ESTADO DE SIMCARDS</div>

    <div id="em-sim-kpis" style="display:flex;flex-wrap:wrap;gap:12px;margin-bottom:28px;">
      <div class="loading"><div class="spinner"></div><span>Cargando...</span></div>
    </div>

    <div style="display:grid;grid-template-columns:1fr 1fr;gap:18px;margin-bottom:20px;">
      ${_cc('Activadas vs Sin Activar','FA:DD/MM/AAAA en columna Atributos = activada','em-chart-sim-estado','240px')}
      ${_cc('Sitios con más SIMCards Activadas','Top 12 por nombre de ubicación','em-chart-sim-activadas','240px')}
    </div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:18px;margin-bottom:24px;">
      ${_cc('Top Sitios por Total SIMCards','Activadas + sin activar por ubicación','em-chart-sim-sitios','280px')}
      ${_cc('Activaciones por Mes','Distribución temporal según fechas FA','em-chart-sim-mensual','280px')}
    </div>

    <div class="section-label fade-up" style="color:#99D1FC;">SIMCards por Comercio / Punto</div>
    <div style="background:linear-gradient(145deg,rgba(10,26,18,.95),rgba(8,20,14,.9));border:1px solid rgba(153,209,252,.12);border-radius:18px;overflow:hidden;margin-bottom:28px;box-shadow:0 4px 20px rgba(0,0,0,.4);">
      <div style="padding:14px 20px;border-bottom:1px solid rgba(153,209,252,.08);display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:10px;">
        <div>
          <div style="font-family:'Syne',sans-serif;font-size:13px;font-weight:700;color:#f1f5f9;">Resumen por Punto de Venta · Activadas · Sin Activar · Total</div>
          <div style="font-size:11px;color:#7A7674;margin-top:2px;">Sumatoria de todas las SIMCards agrupadas por comercio</div>
        </div>
        <div id="em-sim-sitio-suma" style="display:flex;align-items:center;gap:6px;font-size:12px;"></div>
      </div>
      <div id="em-sim-por-sitio-wrap" style="overflow-x:auto;max-height:480px;"></div>
    </div>

    <div class="section-label fade-up" style="color:#99D1FC;">Tabla Completa — SIMCards</div>
    <div style="background:linear-gradient(145deg,rgba(10,26,18,.95),rgba(8,20,14,.9));border:1px solid rgba(153,209,252,.12);border-radius:18px;overflow:hidden;margin-bottom:40px;box-shadow:0 4px 20px rgba(0,0,0,.4);">
      <div style="display:flex;align-items:center;justify-content:space-between;padding:16px 20px;border-bottom:1px solid rgba(153,209,252,.08);">
        <div>
          <div style="font-family:'Syne',sans-serif;font-size:13px;font-weight:700;color:#f1f5f9;">Serial · Ubicación · Estado · Fecha Activación · Atributos</div>
          <div style="font-size:11px;color:#7A7674;margin-top:2px;" id="em-sim-tabla-count">0 registros</div>
        </div>
        <button onclick="emExportSimExcel()" style="background:rgba(153,209,252,.08);border:1px solid rgba(153,209,252,.2);color:#99D1FC;padding:7px 14px;border-radius:8px;cursor:pointer;font-size:12px;font-family:'Outfit',sans-serif;transition:all .2s;">⬇ Excel</button>
      </div>
      <div style="padding:10px 20px;border-bottom:1px solid rgba(153,209,252,.07);background:rgba(0,0,0,.15);">
        <input type="text" oninput="_emApplySimSearch(this.value)" placeholder="🔍 Buscar por serial, nombre de sitio, atributos, código..."
          style="width:100%;box-sizing:border-box;background:rgba(255,255,255,.05);border:1px solid rgba(153,209,252,.2);border-radius:8px;color:#FAFAFA;padding:8px 14px;font-size:13px;font-family:'Outfit',sans-serif;outline:none;">
      </div>
      <div id="em-sim-tabla-wrap" style="overflow-x:auto;max-height:520px;"></div>
      <div style="display:flex;justify-content:flex-end;padding:10px 20px;border-top:1px solid rgba(153,209,252,.07);">
        <div id="em-sim-tabla-pag" style="display:flex;gap:6px;flex-wrap:wrap;"></div>
      </div>
    </div>

  </div>`;
}

// ══════════════════════════════════════════════════════════════════
//  EXPORTS EXCEL
// ══════════════════════════════════════════════════════════════════
window.emExportDfExcel=function(){
  if(!EM_DF_FILTERED.length){alert('Sin datos.');return;}
  if(typeof XLSX==='undefined'){alert('XLSX no disponible.');return;}
  const rows=EM_DF_FILTERED.map(r=>({'Serial':r['Número de serie']||r['Numero de serie']||r['Serial']||'','Referencia':r['Nombre']||'','Bodega':r['Nombre de la ubicación']||'','Código Ubicación':r['Código de ubicación']||'','Posición Depósito':r['Posición en depósito']||'','Estado':_emEstadoDatafono(r),'Tipo Daño':_emTipoDanio(r)||'','Comentarios':r['Comentarios']||'','Atributos':r['Atributos']||''}));
  const wb=XLSX.utils.book_new();const ws=XLSX.utils.json_to_sheet(rows);ws['!cols']=Object.keys(rows[0]).map(k=>({wch:Math.max(k.length+2,14)}));XLSX.utils.book_append_sheet(wb,ws,'Datáfonos Bodegas');XLSX.writeFile(wb,`Datafonos_Bodegas_${new Date().toISOString().slice(0,10)}.xlsx`);
};
window.emExportSimExcel=function(){
  if(!EM_SIM_FILTERED.length){alert('Sin datos.');return;}
  if(typeof XLSX==='undefined'){alert('XLSX no disponible.');return;}
  const rows=EM_SIM_FILTERED.map(r=>({'Serial':r['Número de serie']||r['Numero de serie']||r['Serial']||'','Ubicación':r['Nombre de la ubicación']||'','Código Ubicación':r['Código de ubicación']||'','Estado':_emSimActivada(r)?'ACTIVADA':'SIN ACTIVAR','Fecha Activación':_emSimFechaActivacion(r),'Atributos':r['Atributos']||'','Comentarios':r['Comentarios']||''}));
  const wb=XLSX.utils.book_new();const ws=XLSX.utils.json_to_sheet(rows);ws['!cols']=Object.keys(rows[0]).map(k=>({wch:Math.max(k.length+2,14)}));XLSX.utils.book_append_sheet(wb,ws,'SIMCards');XLSX.writeFile(wb,`SIMCards_${new Date().toISOString().slice(0,10)}.xlsx`);
};

// ══════════════════════════════════════════════════════════════════
//  FILTROS GLOBALES — funciones auxiliares
// ══════════════════════════════════════════════════════════════════

// Autocomplete para bodega
let _emAcBodegaTimer = null;
window._emAcBodega = function(val) {
  clearTimeout(_emAcBodegaTimer);
  _emAcBodegaTimer = setTimeout(() => {
    const ac = document.getElementById('em-f-bodega-ac');
    if (!ac) return;
    const q = val.trim().toLowerCase();
    const raw = _emGetRaw() || [];
    const bodegas = (typeof INV_BODEGAS !== 'undefined' ? INV_BODEGAS : null) || window.INV_BODEGAS;
    const opts = [...new Set(raw.filter(r => bodegas && bodegas.has((r['Nombre de la ubicación']||'').trim())).map(r => (r['Nombre de la ubicación']||'').trim()).filter(Boolean))].sort();
    const matches = q ? opts.filter(o => o.toLowerCase().includes(q)) : opts;
    if (!matches.length) { ac.style.display = 'none'; return; }
    ac.innerHTML = matches.slice(0, 30).map(o => `<div style="padding:7px 12px;cursor:pointer;color:#e2e8f0;border-bottom:1px solid rgba(255,255,255,.04);" onmousedown="event.preventDefault();_emAcBodegaSelect('${o.replace(/'/g,"\\'")}')" onmouseover="this.style.background='rgba(176,242,174,.08)'" onmouseout="this.style.background=''">${o}</div>`).join('');
    ac.style.display = 'block';
  }, 120);
};
window._emAcBodegaSelect = function(val) {
  const inp = document.getElementById('em-f-bodega');
  if (inp) inp.value = val;
  const ac = document.getElementById('em-f-bodega-ac');
  if (ac) ac.style.display = 'none';
  emApplyGlobalFilters();
};
window._emAcBodegaHide = function() {
  setTimeout(() => { const ac = document.getElementById('em-f-bodega-ac'); if (ac) ac.style.display = 'none'; }, 200);
};

// Autocomplete para referencia
let _emAcRefTimer = null;
window._emAcRef = function(val) {
  clearTimeout(_emAcRefTimer);
  _emAcRefTimer = setTimeout(() => {
    const ac = document.getElementById('em-f-ref-ac');
    if (!ac) return;
    const q = val.trim().toLowerCase();
    const raw = _emGetRaw() || [];
    const allRefs = [...new Set(raw.map(r => (r['Nombre']||'').trim()).filter(Boolean))].sort();
    const matches = q ? allRefs.filter(o => o.toLowerCase().includes(q)) : allRefs;
    if (!matches.length) { ac.style.display = 'none'; return; }
    ac.innerHTML = matches.slice(0, 20).map(o => `<div style="padding:7px 12px;cursor:pointer;color:#e2e8f0;border-bottom:1px solid rgba(255,255,255,.04);" onmousedown="event.preventDefault();_emAcRefSelect('${o.replace(/'/g,"\\'")}')" onmouseover="this.style.background='rgba(176,242,174,.08)'" onmouseout="this.style.background=''">${o}</div>`).join('');
    ac.style.display = 'block';
  }, 120);
};
window._emAcRefSelect = function(val) {
  const inp = document.getElementById('em-f-ref');
  if (inp) inp.value = val;
  const ac = document.getElementById('em-f-ref-ac');
  if (ac) ac.style.display = 'none';
  emApplyGlobalFilters();
};
window._emAcRefHide = function() {
  setTimeout(() => { const ac = document.getElementById('em-f-ref-ac'); if (ac) ac.style.display = 'none'; }, 200);
};

// Serial — debounce directo
let _emGfSerialTimer = null;
window._emGfSerialInput = function(val) {
  clearTimeout(_emGfSerialTimer);
  _emGfSerialTimer = setTimeout(() => emApplyGlobalFilters(), 350);
};

// Renderiza los tags activos
function _emRenderFilterTags() {
  const tag = document.getElementById('em-filtros-tag');
  if (!tag) return;
  const tags = [];
  if (EM_F_NEGOCIO) tags.push({ label: `Negocio: ${EM_F_NEGOCIO}`, clear: () => { EM_F_NEGOCIO=''; const el=document.getElementById('em-f-negocio'); if(el)el.value=''; } });
  if (EM_F_UBICV3)  tags.push({ label: `Ubic. V3: ${EM_F_UBICV3}`, clear: () => { EM_F_UBICV3=''; const el=document.getElementById('em-f-ubicv3'); if(el)el.value=''; } });
  if (EM_F_BODEGA)  tags.push({ label: `Bodega: ${EM_F_BODEGA}`, clear: () => { EM_F_BODEGA=''; const el=document.getElementById('em-f-bodega'); if(el)el.value=''; } });
  if (EM_F_SERIAL)  tags.push({ label: `Serial: ${EM_F_SERIAL}`, clear: () => { EM_F_SERIAL=''; const el=document.getElementById('em-f-serial'); if(el)el.value=''; } });
  if (EM_F_REF)     tags.push({ label: `Ref: ${EM_F_REF}`, clear: () => { EM_F_REF=''; const el=document.getElementById('em-f-ref'); if(el)el.value=''; } });
  tag.innerHTML = tags.map((t,i)=>`<span style="display:inline-flex;align-items:center;gap:5px;background:rgba(176,242,174,.1);border:1px solid rgba(176,242,174,.25);border-radius:20px;padding:3px 10px 3px 12px;font-size:11px;color:#B0F2AE;font-family:'Outfit',sans-serif;" data-tag-idx="${i}">${t.label}<button onclick="_emClearTag(${i})" style="background:transparent;border:none;color:#B0F2AE;cursor:pointer;font-size:13px;line-height:1;padding:0 0 1px 2px;" title="Quitar filtro">×</button></span>`).join('');
  window._emTagsClear = tags.map(t=>t.clear);
}
window._emClearTag = function(i) {
  if (window._emTagsClear && window._emTagsClear[i]) { window._emTagsClear[i](); emApplyGlobalFilters(); }
};

// Función principal de aplicación de filtros globales
window.emApplyGlobalFilters = function() {
  EM_F_NEGOCIO = (document.getElementById('em-f-negocio')?.value||'').trim();
  EM_F_UBICV3  = (document.getElementById('em-f-ubicv3')?.value||'').trim();
  EM_F_BODEGA  = (document.getElementById('em-f-bodega')?.value||'').trim();
  EM_F_SERIAL  = (document.getElementById('em-f-serial')?.value||'').trim();
  EM_F_REF     = (document.getElementById('em-f-ref')?.value||'').trim();

  _emRenderFilterTags();
  _emApplyGlobalFilters(); // reconstruye EM_DF_ALL y EM_SIM_ALL con filtros
  EM_DF_SEARCH = ''; EM_SIM_SEARCH = '';
  EM_DF_PAGE = 1; EM_SIM_PAGE = 1;
  _emRenderDatafonos();
  _emRenderSimcards();
};

window.emResetGlobalFilters = function() {
  EM_F_NEGOCIO=''; EM_F_UBICV3=''; EM_F_BODEGA=''; EM_F_SERIAL=''; EM_F_REF='';
  ['em-f-negocio','em-f-ubicv3'].forEach(id=>{const el=document.getElementById(id);if(el)el.value='';});
  ['em-f-bodega','em-f-serial','em-f-ref'].forEach(id=>{const el=document.getElementById(id);if(el)el.value='';});
  _emRenderFilterTags();
  _emApplyGlobalFilters();
  EM_DF_SEARCH=''; EM_SIM_SEARCH=''; EM_DF_PAGE=1; EM_SIM_PAGE=1;
  _emRenderDatafonos(); _emRenderSimcards();
};

// ══════════════════════════════════════════════════════════════════
//  ENTRY POINT
// ══════════════════════════════════════════════════════════════════
async function renderEstadoMateriales() {
  const panel=document.getElementById('panel-estado-materiales');
  if(!panel) return;
  let raw=_emGetRaw();
  if(!raw||!raw.length){
    panel.innerHTML='<div class="loading"><div class="spinner"></div><span>Cargando inventario...</span></div>';
    if(typeof window.loadInventarioData==='function') await window.loadInventarioData();
    raw=_emGetRaw();
  }
  if(!raw||!raw.length){
    panel.innerHTML='<div style="text-align:center;padding:60px;color:#7A7674;font-family:\'Outfit\',sans-serif;">No hay datos disponibles.</div>';
    return;
  }
  Object.keys(EM_CHARTS).forEach(id=>{try{EM_CHARTS[id].destroy();}catch(e){}});
  EM_CHARTS={};
  EM_F_NEGOCIO=''; EM_F_UBICV3=''; EM_F_BODEGA=''; EM_F_SERIAL=''; EM_F_REF='';
  _emBuildHTML();
  _emProcesar();
  requestAnimationFrame(()=>{ _emRenderDatafonos(); _emRenderSimcards(); });
}

window.renderEstadoMateriales=renderEstadoMateriales;
window._emApplyDfSearch=_emApplyDfSearch;
window._emApplySimSearch=_emApplySimSearch;
window._emRenderFilterTags=_emRenderFilterTags;