/**
 * login_tracker.js — Registro y visualización de accesos al dashboard
 * Guarda en Firebase Realtime DB: nombre, cargo, foto, timestamp
 * Muestra en panel-home quién ha entrado y hace cuánto (global, todos los PCs)
 * Desarrollado para Dashboard Unificado Wompi × Linea Comunicaciones
 */
'use strict';

const LoginTracker = (() => {

  // ════════════════════════════════════════════════════════
  //  FIREBASE CONFIG
  // ════════════════════════════════════════════════════════
  const FIREBASE_CONFIG = {
    apiKey:            "AIzaSyAvG6kJ8arHJbybxSyeB4zMHza6nxhzVcg",
    authDomain:        "tableroswompi.firebaseapp.com",
    databaseURL:       "https://tableroswompi-default-rtdb.firebaseio.com",
    projectId:         "tableroswompi",
    storageBucket:     "tableroswompi.firebasestorage.app",
    messagingSenderId: "675251150612",
    appId:             "1:675251150612:web:be9b9473fd2293152eec01",
    measurementId:     "G-S1YC8WDC55"
  };

  const DB_PATH    = 'dashboard_logins';
  const MAX_SHOW   = 40;   // máx entradas visibles en el panel

  let _db = null, _fbRef = null, _fbGet = null, _fbSet = null, _fbQuery = null,
      _orderByChild = null, _limitToLast = null;
  let _panelRendered = false;
  let _refreshInterval = null;

  // ── Inicializar Firebase (módulos ES via CDN) ────────────────────
  async function _initFirebase() {
    if (_db) return true;
    try {
      const { initializeApp, getApps } = await import(
        'https://www.gstatic.com/firebasejs/10.12.0/firebase-app.js'
      );
      const { getDatabase, ref, get, set, query, orderByChild, limitToLast } = await import(
        'https://www.gstatic.com/firebasejs/10.12.0/firebase-database.js'
      );
      const app = getApps().length ? getApps()[0] : initializeApp(FIREBASE_CONFIG);
      _db           = getDatabase(app);
      _fbRef        = ref;
      _fbGet        = get;
      _fbSet        = set;
      _fbQuery      = query;
      _orderByChild = orderByChild;
      _limitToLast  = limitToLast;
      return true;
    } catch (e) {
      console.warn('[LoginTracker] Firebase init error:', e.message);
      return false;
    }
  }

  // ── Tiempo relativo legible ──────────────────────────────────────
  function _timeAgo(ts) {
    const diff = Date.now() - ts;
    const m = Math.floor(diff / 60000);
    if (m < 1)  return 'Ahora mismo';
    if (m < 60) return `Hace ${m} min`;
    const h = Math.floor(m / 60);
    if (h < 24) return `Hace ${h}h ${m % 60}min`;
    const d = Math.floor(h / 24);
    return `Hace ${d} día${d !== 1 ? 's' : ''}`;
  }

  // ── Guardar login en Firebase ────────────────────────────────────
  async function registrar(profile) {
    const ok = await _initFirebase();
    if (!ok) return;
    try {
      const key = `${Date.now()}_${Math.random().toString(36).slice(2, 7)}`;
      await _fbSet(_fbRef(_db, `${DB_PATH}/${key}`), {
        ts:     Date.now(),
        nombre: profile.nombre || 'Desconocido',
        cargo:  profile.cargo  || '',
        email:  profile.email  || '',
        foto:   profile.foto   || ''   // base64 completo desde Microsoft Graph
      });
      console.log('[LoginTracker] ✅ Acceso registrado:', profile.email);
    } catch (e) {
      console.warn('[LoginTracker] Error al guardar:', e.message);
    }
  }

  // ── Cargar entradas de Firebase y renderizar chips ───────────────
  async function _loadAndRender() {
    const list    = document.getElementById('lt-list');
    const countEl = document.getElementById('lt-count');
    if (!list) return;

    try {
      const q    = _fbQuery(_fbRef(_db, DB_PATH), _orderByChild('ts'), _limitToLast(MAX_SHOW));
      const snap = await _fbGet(q);

      if (!snap.exists()) {
        list.innerHTML = `
          <div style="font-size:12px;color:#475569;font-family:'Outfit',sans-serif;
                      padding:12px 0;width:100%;text-align:center;">
            Aún no hay accesos registrados.
          </div>`;
        return;
      }

      // Ordenar más reciente primero
      const entries = [];
      snap.forEach(child => entries.push(child.val()));
      entries.sort((a, b) => b.ts - a.ts);

      if (countEl) {
        countEl.textContent = `· ${entries.length} acceso${entries.length !== 1 ? 's' : ''}`;
      }

      list.innerHTML = entries.map(e => {
        const initials = (e.nombre || 'U')
          .split(' ').map(w => w[0]).join('').substring(0, 2).toUpperCase();

        const avatarHTML = e.foto
          ? `<img class="lt-avatar" src="${e.foto}" alt="${initials}"
               onerror="this.outerHTML='<div class=\\'lt-initials\\'>${initials}</div>'">`
          : `<div class="lt-initials">${initials}</div>`;

        // Hora legible local
        const fechaLocal = new Date(e.ts).toLocaleString('es-CO', {
          day: '2-digit', month: '2-digit',
          hour: '2-digit', minute: '2-digit'
        });

        return `
          <div class="lt-chip" title="${e.email ? e.email : ''}">
            ${avatarHTML}
            <div>
              <div class="lt-name">${e.nombre || 'Desconocido'}</div>
              <div class="lt-meta">
                ${_timeAgo(e.ts)}${e.cargo ? ' · ' + e.cargo : ''}
              </div>
            </div>
          </div>`;
      }).join('');

    } catch (e) {
      console.warn('[LoginTracker] Error al leer Firebase:', e.message);
      const list = document.getElementById('lt-list');
      if (list) list.innerHTML = `
        <div style="font-size:12px;color:#EF4444;font-family:'Outfit',sans-serif;padding:8px 0;">
          Error al conectar con Firebase.
        </div>`;
    }
  }

  // ── Inyectar estilos de los chips ───────────────────────────────
  function _injectStyles() {
    if (document.getElementById('lt-styles')) return;
    const style = document.createElement('style');
    style.id = 'lt-styles';
    style.textContent = `
      @keyframes lt-pulse {
        0%,100% { opacity:1; transform:scale(1); }
        50%      { opacity:.4; transform:scale(1.4); }
      }
      @keyframes lt-fadein {
        from { opacity:0; transform:translateY(8px); }
        to   { opacity:1; transform:translateY(0); }
      }
      #lt-section {
        animation: lt-fadein .4s ease both;
      }
      .lt-chip {
        display:flex; align-items:center; gap:9px;
        padding:7px 14px 7px 7px;
        background:rgba(255,255,255,.04);
        border:1px solid rgba(255,255,255,.08);
        border-radius:40px;
        transition:background .2s, border-color .2s, transform .15s;
        cursor:default;
        animation: lt-fadein .3s ease both;
      }
      .lt-chip:hover {
        background:rgba(176,242,174,.06);
        border-color:rgba(176,242,174,.22);
        transform:translateY(-1px);
      }
      .lt-avatar {
        width:32px; height:32px; border-radius:50%;
        object-fit:cover; flex-shrink:0;
        border:1.5px solid rgba(176,242,174,.25);
      }
      .lt-initials {
        width:32px; height:32px; border-radius:50%;
        background:linear-gradient(135deg,#3B82F6,#8B5CF6);
        display:flex; align-items:center; justify-content:center;
        font-size:11px; font-weight:700; color:#fff; flex-shrink:0;
        border:1.5px solid rgba(139,92,246,.4);
      }
      .lt-name {
        font-family:'Outfit',sans-serif;
        font-size:12px; font-weight:600; color:#e2e8f0; line-height:1.2;
        white-space:nowrap;
      }
      .lt-meta {
        font-family:'JetBrains Mono',monospace;
        font-size:10px; color:#475569; margin-top:1px;
        white-space:nowrap;
      }
    `;
    document.head.appendChild(style);
  }

  // ── Crear y montar la sección en panel-home ──────────────────────
  async function renderPanel() {
    if (_panelRendered) return;

    const panelHome = document.getElementById('panel-home');
    if (!panelHome) return;

    _injectStyles();

    // Crear sección
    const section = document.createElement('div');
    section.id = 'lt-section';
    section.style.cssText = `
      width:100%; max-width:900px; margin-top:4px;
      background:rgba(255,255,255,.025);
      border:1px solid rgba(255,255,255,.07);
      border-radius:20px; padding:24px 28px;
    `;
    section.innerHTML = `
      <div style="display:flex;align-items:center;gap:10px;margin-bottom:18px;flex-wrap:wrap;">
        <div style="
          width:8px;height:8px;border-radius:50%;
          background:#B0F2AE;box-shadow:0 0 10px #B0F2AE55;
          animation:lt-pulse 2s infinite;flex-shrink:0;">
        </div>
        <span style="
          font-family:'Syne',sans-serif;font-size:13px;
          font-weight:700;color:#f1f5f9;letter-spacing:.3px;">
          Accesos Recientes
        </span>
        <span id="lt-count" style="
          font-size:11px;color:#475569;
          font-family:'Outfit',sans-serif;margin-left:2px;">
        </span>
        <button id="lt-refresh-btn" onclick="window.LoginTracker._refresh()" style="
          margin-left:auto;padding:4px 12px;
          background:rgba(176,242,174,.07);
          border:1px solid rgba(176,242,174,.18);
          border-radius:20px;color:#B0F2AE;font-size:11px;
          font-family:'Outfit',sans-serif;cursor:pointer;
          transition:background .2s;"
          onmouseover="this.style.background='rgba(176,242,174,.14)'"
          onmouseout="this.style.background='rgba(176,242,174,.07)'">
          ↺ Actualizar
        </button>
      </div>
      <div id="lt-list" style="
        display:flex;flex-wrap:wrap;gap:8px;min-height:48px;
        align-items:flex-start;align-content:flex-start;">
        <div style="font-size:12px;color:#475569;font-family:'Outfit',sans-serif;
                    padding:8px 0;width:100%;text-align:center;">
          Conectando con Firebase...
        </div>
      </div>
    `;
    panelHome.appendChild(section);
    _panelRendered = true;

    // Inicializar Firebase y cargar datos
    const ok = await _initFirebase();
    if (!ok) {
      document.getElementById('lt-list').innerHTML = `
        <div style="font-size:12px;color:#EF4444;font-family:'Outfit',sans-serif;
                    padding:8px 0;width:100%;">
          ⚠ No se pudo conectar a Firebase. Verifica la consola.
        </div>`;
      return;
    }

    await _loadAndRender();

    // Refrescar cada 30 segundos automáticamente
    if (_refreshInterval) clearInterval(_refreshInterval);
    _refreshInterval = setInterval(_loadAndRender, 30000);
  }

  // Expuesta para el botón de actualizar
  async function _refresh() {
    const btn = document.getElementById('lt-refresh-btn');
    if (btn) { btn.textContent = '↺ ...'; btn.disabled = true; }
    await _loadAndRender();
    if (btn) { btn.textContent = '↺ Actualizar'; btn.disabled = false; }
  }

  return { registrar, renderPanel, _refresh };

})();

window.LoginTracker = LoginTracker;