'use strict';
/**
 * ╔══════════════════════════════════════════════════════════════════╗
 * ║  notas_reorden.js — Notas colaborativas por Punto de Reorden   ║
 * ║  Persiste en Firebase Realtime DB (misma base que login_tracker)║
 * ║  Path: dashboard_notas_reorden/{seccionId}/{noteKey}            ║
 * ║  Desarrollado para Dashboard Unificado Wompi × Linea Comunicac. ║
 * ╚══════════════════════════════════════════════════════════════════╝
 *
 * INTEGRACIÓN:
 *   1. Cargar este script en index.html DESPUÉS de login_tracker.js
 *   2. En puntos_reorden.js, al final de cada pr-card (antes del cierre
 *      de </div> de la card), agregar:
 *        html += NotasReorden.cardHTML(cfg.id, cfg.acento);
 *   3. Al final de renderPuntosReorden(), después de panel.innerHTML = html:
 *        NotasReorden.init();
 */

const NotasReorden = (() => {

  // ════════════════════════════════════════════════════════
  //  FIREBASE CONFIG — igual que login_tracker.js
  // ════════════════════════════════════════════════════════
  const FIREBASE_CONFIG = {
    apiKey:            'AIzaSyAvG6kJ8arHJbybxSyeB4zMHza6nxhzVcg',
    authDomain:        'tableroswompi.firebaseapp.com',
    databaseURL:       'https://tableroswompi-default-rtdb.firebaseio.com',
    projectId:         'tableroswompi',
    storageBucket:     'tableroswompi.firebasestorage.app',
    messagingSenderId: '675251150612',
    appId:             '1:675251150612:web:be9b9473fd2293152eec01',
    measurementId:     'G-S1YC8WDC55'
  };

  const DB_PATH  = 'dashboard_notas_reorden';   // nodo raíz en Firebase
  const MAX_SHOW = 30;                           // máx notas visibles por sección

  let _db = null, _fbRef = null, _fbGet = null, _fbSet = null, _fbRemove = null;

  // ── Inicializar Firebase (reutiliza instancia si ya existe) ──────
  async function _initFirebase() {
    if (_db) return true;
    try {
      const { initializeApp, getApps } = await import(
        'https://www.gstatic.com/firebasejs/10.12.0/firebase-app.js'
      );
      const { getDatabase, ref, get, set, remove } = await import(
        'https://www.gstatic.com/firebasejs/10.12.0/firebase-database.js'
      );
      const app = getApps().length ? getApps()[0] : initializeApp(FIREBASE_CONFIG);
      _db       = getDatabase(app);
      _fbRef    = ref;
      _fbGet    = get;
      _fbSet    = set;
      _fbRemove = remove;
      return true;
    } catch (e) {
      console.warn('[NotasReorden] Firebase init error:', e.message);
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

  // ── Nombre del usuario activo (de _msUserProfile) ───────────────
  function _currentUser() {
    const p = window._msUserProfile || {};
    return {
      nombre: p.nombre || 'Anónimo',
      cargo:  p.cargo  || '',
      email:  p.email  || ''
    };
  }

  // ── Guardar nota en Firebase ─────────────────────────────────────
  async function _guardarNota(seccionId, texto) {
    const ok = await _initFirebase();
    if (!ok) throw new Error('Firebase no disponible');
    const user = _currentUser();
    const key  = `${Date.now()}_${Math.random().toString(36).slice(2, 6)}`;
    await _fbSet(_fbRef(_db, `${DB_PATH}/${seccionId}/${key}`), {
      ts:     Date.now(),
      texto:  texto.trim(),
      nombre: user.nombre,
      cargo:  user.cargo,
      email:  user.email
    });
    return key;
  }

  // ── Eliminar nota de Firebase ────────────────────────────────────
  async function _eliminarNota(seccionId, noteKey) {
    const ok = await _initFirebase();
    if (!ok) throw new Error('Firebase no disponible');
    await _fbRemove(_fbRef(_db, `${DB_PATH}/${seccionId}/${noteKey}`));
  }

  // ── Cargar y renderizar notas para una sección ───────────────────
  async function _cargarNotas(seccionId) {
    const lista = document.getElementById(`nr-list-${seccionId}`);
    if (!lista) return;

    lista.innerHTML = `<div class="nr-loading">Cargando notas…</div>`;

    const ok = await _initFirebase();
    if (!ok) {
      lista.innerHTML = `<div class="nr-error">⚠ No se pudo conectar a Firebase.</div>`;
      return;
    }

    try {
      const snap = await _fbGet(_fbRef(_db, `${DB_PATH}/${seccionId}`));

      if (!snap.exists()) {
        lista.innerHTML = `<div class="nr-empty">Aún no hay notas. ¡Sé el primero en agregar una!</div>`;
        return;
      }

      const raw = snap.val();
      const notas = Object.entries(raw)
        .filter(([, v]) => v && v.ts && v.texto)
        .sort(([, a], [, b]) => b.ts - a.ts)
        .slice(0, MAX_SHOW);

      const currentEmail = window._msUserProfile?.email || '';

      lista.innerHTML = notas.map(([key, n]) => {
        const initials = (n.nombre || 'U')
          .split(' ').map(w => w[0]).join('').substring(0, 2).toUpperCase();
        const esPropio = currentEmail && n.email === currentEmail;
        const fecha = new Date(n.ts).toLocaleString('es-CO', {
          day: '2-digit', month: '2-digit', year: '2-digit',
          hour: '2-digit', minute: '2-digit'
        });

        return `
          <div class="nr-nota" id="nr-nota-${key}" data-key="${key}" data-seccion="${seccionId}">
            <div class="nr-nota-header">
              <div class="nr-avatar">${initials}</div>
              <div class="nr-nota-meta">
                <span class="nr-nota-autor">${n.nombre || 'Desconocido'}</span>
                ${n.cargo ? `<span class="nr-nota-cargo">${n.cargo}</span>` : ''}
                <span class="nr-nota-tiempo" title="${fecha}">${_timeAgo(n.ts)} · ${fecha}</span>
              </div>
              ${esPropio ? `
                <button class="nr-btn-delete" title="Eliminar nota"
                  onclick="NotasReorden._borrar('${seccionId}','${key}')"
                  aria-label="Eliminar">✕</button>
              ` : ''}
            </div>
            <div class="nr-nota-texto">${_escapeHtml(n.texto)}</div>
          </div>`;
      }).join('');

    } catch (e) {
      console.warn('[NotasReorden] Error al leer:', e.message);
      lista.innerHTML = `<div class="nr-error">⚠ Error al cargar notas. Intenta de nuevo.</div>`;
    }
  }

  // ── Escapa HTML para evitar XSS ──────────────────────────────────
  function _escapeHtml(str) {
    return (str || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  }

  // ── Manejar envío de nota ────────────────────────────────────────
  async function _enviar(seccionId) {
    const input  = document.getElementById(`nr-input-${seccionId}`);
    const btn    = document.getElementById(`nr-btn-${seccionId}`);
    const errEl  = document.getElementById(`nr-err-${seccionId}`);
    if (!input || !btn) return;

    const texto = (input.value || '').trim();
    if (!texto) {
      errEl.textContent = 'Escribe una nota antes de guardar.';
      errEl.style.display = 'block';
      input.focus();
      return;
    }
    if (texto.length > 500) {
      errEl.textContent = 'Máximo 500 caracteres.';
      errEl.style.display = 'block';
      return;
    }
    errEl.style.display = 'none';

    btn.disabled = true;
    btn.textContent = 'Guardando…';
    input.disabled = true;

    try {
      await _guardarNota(seccionId, texto);
      input.value = '';
      // Actualizar contador de caracteres
      const counter = document.getElementById(`nr-counter-${seccionId}`);
      if (counter) counter.textContent = '0 / 500';
      await _cargarNotas(seccionId);
    } catch (e) {
      errEl.textContent = '⚠ Error al guardar. Verifica tu conexión.';
      errEl.style.display = 'block';
      console.warn('[NotasReorden] Error al guardar:', e.message);
    } finally {
      btn.disabled = false;
      btn.textContent = 'Guardar nota';
      input.disabled = false;
      input.focus();
    }
  }

  // ── Borrar nota (solo el autor puede ver el botón) ───────────────
  async function _borrar(seccionId, key) {
    if (!confirm('¿Eliminar esta nota?')) return;
    try {
      await _eliminarNota(seccionId, key);
      const el = document.getElementById(`nr-nota-${key}`);
      if (el) {
        el.style.opacity = '0';
        el.style.transform = 'translateX(20px)';
        setTimeout(() => el.remove(), 300);
      }
      // Re-check si quedó vacío
      setTimeout(() => _cargarNotas(seccionId), 350);
    } catch (e) {
      alert('No se pudo eliminar la nota. Intenta de nuevo.');
      console.warn('[NotasReorden] Error al borrar:', e.message);
    }
  }

  // ── HTML del bloque de notas (se inyecta dentro de cada pr-card) ─
  // @param seccionId  — id del PR_CONFIG (ej: 'rollos', 'datafonos-cb')
  // @param acento     — color hex del acento de la card (ej: '#DFFF61')
  function cardHTML(seccionId, acento) {
    return `
      <div class="nr-section" style="--nr-acento:${acento};">
        <div class="nr-section-header">
          <span class="nr-section-icon">📝</span>
          <span class="nr-section-title">Notas del equipo</span>
          <button class="nr-refresh-btn" onclick="NotasReorden._recargar('${seccionId}')"
            title="Actualizar notas">↺</button>
        </div>

        <!-- Lista de notas -->
        <div class="nr-list" id="nr-list-${seccionId}">
          <div class="nr-loading">Conectando…</div>
        </div>

        <!-- Formulario de nueva nota -->
        <div class="nr-form">
          <textarea
            id="nr-input-${seccionId}"
            class="nr-textarea"
            placeholder='Ej: "Vienen 5 000 datáfonos, entrega estimada el 20 de mayo."'
            maxlength="500"
            rows="3"
            oninput="
              var c=document.getElementById('nr-counter-${seccionId}');
              if(c) c.textContent=this.value.length+' / 500';
            "
            onkeydown="if((event.ctrlKey||event.metaKey)&&event.key==='Enter'){NotasReorden._enviar('${seccionId}');}"
          ></textarea>
          <div class="nr-form-footer">
            <span id="nr-counter-${seccionId}" class="nr-counter">0 / 500</span>
            <span id="nr-err-${seccionId}" class="nr-error-msg" style="display:none;"></span>
            <button id="nr-btn-${seccionId}" class="nr-submit-btn"
              onclick="NotasReorden._enviar('${seccionId}')">
              Guardar nota
            </button>
          </div>
        </div>

        <div class="nr-hint">Ctrl + Enter para guardar rápido · Solo tú puedes borrar tus notas</div>
      </div>
    `;
  }

  // ── Inicializar todas las secciones después del render ───────────
  function init() {
    // Inyectar estilos (idempotente)
    _injectStyles();
    // Cargar notas para cada sección que tenga un contenedor
    const secciones = ['rollos', 'datafonos-cb', 'pinpads-cb', 'datafonos-vp'];
    secciones.forEach(id => {
      const el = document.getElementById(`nr-list-${id}`);
      if (el) _cargarNotas(id);
    });
  }

  // ── Recargar notas de una sección (botón ↺) ──────────────────────
  async function _recargar(seccionId) {
    const btn = document.querySelector(`#nr-list-${seccionId}`)
      ?.closest('.nr-section')
      ?.querySelector('.nr-refresh-btn');
    if (btn) { btn.textContent = '↺…'; btn.disabled = true; }
    await _cargarNotas(seccionId);
    if (btn) { btn.textContent = '↺'; btn.disabled = false; }
  }

  // ── Inyectar estilos ─────────────────────────────────────────────
  function _injectStyles() {
    if (document.getElementById('nr-styles')) return;
    const style = document.createElement('style');
    style.id = 'nr-styles';
    style.textContent = `
      /* ── Sección contenedora ── */
      @keyframes nr-fadein {
        from { opacity: 0; transform: translateY(6px); }
        to   { opacity: 1; transform: translateY(0); }
      }
      .nr-section {
        margin-top: 16px;
        padding-top: 16px;
        border-top: 1px solid rgba(255,255,255,0.07);
        animation: nr-fadein 0.35s ease both;
      }

      /* ── Cabecera de la sección ── */
      .nr-section-header {
        display: flex;
        align-items: center;
        gap: 8px;
        margin-bottom: 12px;
      }
      .nr-section-icon { font-size: 15px; }
      .nr-section-title {
        font-family: 'Syne', sans-serif;
        font-size: 12px;
        font-weight: 700;
        color: var(--nr-acento, #B0F2AE);
        letter-spacing: 0.3px;
        flex: 1;
      }
      .nr-refresh-btn {
        background: none;
        border: 1px solid rgba(255,255,255,0.1);
        border-radius: 12px;
        color: #475569;
        font-size: 12px;
        padding: 2px 8px;
        cursor: pointer;
        transition: color 0.2s, border-color 0.2s;
      }
      .nr-refresh-btn:hover {
        color: var(--nr-acento, #B0F2AE);
        border-color: var(--nr-acento, #B0F2AE);
      }

      /* ── Lista de notas ── */
      .nr-list {
        display: flex;
        flex-direction: column;
        gap: 8px;
        max-height: 280px;
        overflow-y: auto;
        margin-bottom: 12px;
        padding-right: 4px;
        scrollbar-width: thin;
        scrollbar-color: rgba(255,255,255,0.08) transparent;
      }
      .nr-list::-webkit-scrollbar { width: 4px; }
      .nr-list::-webkit-scrollbar-thumb { background: rgba(255,255,255,0.08); border-radius: 4px; }

      /* ── Mensajes de estado ── */
      .nr-loading, .nr-empty, .nr-error {
        font-family: 'Outfit', sans-serif;
        font-size: 12px;
        text-align: center;
        padding: 12px 0;
        color: #475569;
      }
      .nr-error { color: #EF4444; }

      /* ── Nota individual ── */
      .nr-nota {
        background: rgba(255,255,255,0.03);
        border: 1px solid rgba(255,255,255,0.06);
        border-radius: 10px;
        padding: 10px 12px;
        transition: background 0.2s, opacity 0.3s, transform 0.3s;
        animation: nr-fadein 0.25s ease both;
      }
      .nr-nota:hover { background: rgba(255,255,255,0.05); }

      .nr-nota-header {
        display: flex;
        align-items: center;
        gap: 8px;
        margin-bottom: 7px;
      }
      .nr-avatar {
        width: 26px;
        height: 26px;
        border-radius: 50%;
        background: linear-gradient(135deg, #3B82F6, #8B5CF6);
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 9px;
        font-weight: 700;
        color: #fff;
        flex-shrink: 0;
        border: 1.5px solid rgba(139,92,246,0.3);
      }
      .nr-nota-meta {
        display: flex;
        flex-wrap: wrap;
        align-items: baseline;
        gap: 4px;
        flex: 1;
        min-width: 0;
      }
      .nr-nota-autor {
        font-family: 'Outfit', sans-serif;
        font-size: 11.5px;
        font-weight: 700;
        color: #e2e8f0;
        white-space: nowrap;
      }
      .nr-nota-cargo {
        font-family: 'Outfit', sans-serif;
        font-size: 10px;
        color: #475569;
        background: rgba(255,255,255,0.05);
        border-radius: 8px;
        padding: 1px 6px;
        white-space: nowrap;
      }
      .nr-nota-tiempo {
        font-family: 'JetBrains Mono', monospace;
        font-size: 9.5px;
        color: #475569;
        white-space: nowrap;
      }
      .nr-btn-delete {
        background: none;
        border: none;
        color: #475569;
        font-size: 11px;
        cursor: pointer;
        padding: 2px 5px;
        border-radius: 6px;
        transition: color 0.2s, background 0.2s;
        flex-shrink: 0;
        line-height: 1;
      }
      .nr-btn-delete:hover {
        color: #EF4444;
        background: rgba(239,68,68,0.1);
      }
      .nr-nota-texto {
        font-family: 'Outfit', sans-serif;
        font-size: 12.5px;
        color: #cbd5e1;
        line-height: 1.55;
        white-space: pre-wrap;
        word-break: break-word;
        padding-left: 34px; /* alinear bajo el nombre */
      }

      /* ── Formulario ── */
      .nr-form { display: flex; flex-direction: column; gap: 6px; }
      .nr-textarea {
        width: 100%;
        box-sizing: border-box;
        background: rgba(255,255,255,0.04);
        border: 1px solid rgba(255,255,255,0.1);
        border-radius: 10px;
        color: #e2e8f0;
        font-family: 'Outfit', sans-serif;
        font-size: 12.5px;
        line-height: 1.5;
        padding: 10px 12px;
        resize: vertical;
        min-height: 72px;
        transition: border-color 0.2s, box-shadow 0.2s;
        outline: none;
      }
      .nr-textarea::placeholder { color: #334155; }
      .nr-textarea:focus {
        border-color: var(--nr-acento, #B0F2AE);
        box-shadow: 0 0 0 3px color-mix(in srgb, var(--nr-acento, #B0F2AE) 12%, transparent);
      }
      .nr-textarea:disabled { opacity: 0.5; }

      .nr-form-footer {
        display: flex;
        align-items: center;
        gap: 8px;
      }
      .nr-counter {
        font-family: 'JetBrains Mono', monospace;
        font-size: 10px;
        color: #334155;
        flex: 1;
      }
      .nr-error-msg {
        font-family: 'Outfit', sans-serif;
        font-size: 11px;
        color: #EF4444;
      }
      .nr-submit-btn {
        background: linear-gradient(135deg,
          color-mix(in srgb, var(--nr-acento, #B0F2AE) 18%, transparent),
          color-mix(in srgb, var(--nr-acento, #B0F2AE) 10%, transparent));
        border: 1px solid color-mix(in srgb, var(--nr-acento, #B0F2AE) 35%, transparent);
        border-radius: 20px;
        color: var(--nr-acento, #B0F2AE);
        font-family: 'Outfit', sans-serif;
        font-size: 11.5px;
        font-weight: 600;
        padding: 6px 16px;
        cursor: pointer;
        transition: background 0.2s, transform 0.15s, box-shadow 0.2s;
        white-space: nowrap;
        flex-shrink: 0;
      }
      .nr-submit-btn:hover:not(:disabled) {
        background: color-mix(in srgb, var(--nr-acento, #B0F2AE) 22%, transparent);
        box-shadow: 0 0 16px color-mix(in srgb, var(--nr-acento, #B0F2AE) 25%, transparent);
        transform: translateY(-1px);
      }
      .nr-submit-btn:disabled { opacity: 0.5; cursor: not-allowed; transform: none; }

      /* ── Hint inferior ── */
      .nr-hint {
        font-family: 'JetBrains Mono', monospace;
        font-size: 9.5px;
        color: #1e293b;
        margin-top: 6px;
        letter-spacing: 0.1px;
      }
    `;
    document.head.appendChild(style);
  }

  // ── API pública ──────────────────────────────────────────────────
  return { cardHTML, init, _enviar, _borrar, _recargar };

})();

window.NotasReorden = NotasReorden;