'use strict';
// ══════════════════════════════════════════════════════════════════
//  AUTH — Microsoft Azure AD (MSAL)
//  Solo permite acceso a cuentas @lineacom.co
// ══════════════════════════════════════════════════════════════════

const MSAL_CONFIG = {
  auth: {
    clientId: 'febe226c-0265-4fb2-b34e-3beebbb9fee8',
    authority: 'https://login.microsoftonline.com/af1a17b2-5d34-4f58-8b6c-6b94c6cd87ea',
    redirectUri: window.location.origin + window.location.pathname,
    navigateToLoginRequestUrl: true
  },
  cache: { cacheLocation: 'sessionStorage', storeAuthStateInCookie: false }
};

const MS_SCOPES        = ['openid','profile','email','User.Read'];
const ALLOWED_DOMAIN   = 'lineacom.co';
const SESSION_DURATION = 1 * 60 * 60 * 1000; // 1 hora

let _msalInstance = null;
let _sessionTimer = null;

// ── Cargar MSAL dinámicamente ─────────────────────────────────────
async function _loadMSAL() {
  if (window.msal) return window.msal;
  const URLS = [
    'https://alcdn.msauth.net/browser/2.38.3/js/msal-browser.min.js',
    'https://alcdn.msftauth.net/browser/2.38.3/js/msal-browser.min.js',
    'https://cdn.jsdelivr.net/npm/@azure/msal-browser@2.38.3/lib/msal-browser.min.js',
    'https://unpkg.com/@azure/msal-browser@2.38.3/lib/msal-browser.min.js'
  ];
  for (const url of URLS) {
    try {
      await new Promise((resolve, reject) => {
        const s = document.createElement('script');
        s.src = url; s.onload = resolve; s.onerror = reject;
        document.head.appendChild(s);
      });
      if (window.msal) return window.msal;
    } catch (_) {}
  }
  throw new Error('No se pudo cargar MSAL. Verifica tu conexión a Internet.');
}

// ── Mostrar error en login ────────────────────────────────────────
function _showLoginError(msg) {
  const e = document.getElementById('login-error');
  e.innerHTML = '⚠ ' + msg;
  e.style.display = 'block';
  const btn = document.getElementById('msLoginBtn');
  if (btn) { btn.disabled = false; btn.innerHTML = _msBtnHTML(); }
}

function _msBtnHTML() {
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 23 23" style="width:20px;height:20px;flex-shrink:0;"><rect x="1" y="1" width="10" height="10" fill="#f25022"/><rect x="12" y="1" width="10" height="10" fill="#7fba00"/><rect x="1" y="12" width="10" height="10" fill="#00a4ef"/><rect x="12" y="12" width="10" height="10" fill="#ffb900"/></svg> Iniciar sesión con Microsoft`;
}

// ── Login con Microsoft ───────────────────────────────────────────
async function doMicrosoftLogin() {
  const btn = document.getElementById('msLoginBtn');
  if (btn) { btn.disabled = true; btn.textContent = 'Conectando con Microsoft...'; }
  document.getElementById('login-error').style.display = 'none';

  try {
    const msal = await _loadMSAL();
    _msalInstance = new msal.PublicClientApplication(MSAL_CONFIG);
    await _msalInstance.initialize();

    const result = await _msalInstance.loginPopup({ scopes: MS_SCOPES, prompt: 'select_account' });
    const email  = result.account?.username || '';
    const domain = email.split('@')[1]?.toLowerCase();

    if (domain !== ALLOWED_DOMAIN) {
      await _msalInstance.logoutPopup({ account: result.account }).catch(() => {});
      _showLoginError(`Acceso denegado. Solo cuentas @${ALLOWED_DOMAIN}.<br><small style="opacity:0.7">Tu cuenta: ${email}</small>`);
      return;
    }

    // Obtener perfil desde Microsoft Graph
    let displayName = result.account.name || email, jobTitle = '', photo = null;
    try {
      const graphToken = await _msalInstance.acquireTokenSilent({ scopes: ['User.Read'], account: result.account });
      const profileRes = await fetch('https://graph.microsoft.com/v1.0/me?$select=displayName,jobTitle,mail', { headers: { Authorization: `Bearer ${graphToken.accessToken}` } });
      if (profileRes.ok) { const p = await profileRes.json(); displayName = p.displayName || displayName; jobTitle = p.jobTitle || ''; }
      const photoRes = await fetch('https://graph.microsoft.com/v1.0/me/photo/$value', { headers: { Authorization: `Bearer ${graphToken.accessToken}` } });
      if (photoRes.ok) {
        const blob = await photoRes.blob();
        photo = await new Promise(res => { const r = new FileReader(); r.onload = () => res(r.result); r.readAsDataURL(blob); });
      }
    } catch (graphErr) { console.warn('[Auth] Graph parcial:', graphErr.message); }

    window._msUserProfile = { nombre: displayName, cargo: jobTitle, email, foto: photo };

    // ── Registrar acceso en Firebase (global para todos los PCs) ──
    if (window.LoginTracker) {
      window.LoginTracker.registrar(window._msUserProfile).catch(() => {});
    }

    _enterApp();

  } catch (err) {
    if (err.errorCode === 'user_cancelled') _showLoginError('Inicio de sesión cancelado.');
    else { console.error('[Auth]', err); _showLoginError('Error al conectar con Microsoft. Intenta de nuevo.'); }
  }
}

// ── Mostrar la app tras login exitoso ─────────────────────────────
function _enterApp() {
  const p = window._msUserProfile || {};
  const topbarRight = document.querySelector('.topbar-right');
  if (topbarRight && !document.getElementById('ms-user-chip')) {
    const initials = (p.nombre || 'U').split(' ').map(w => w[0]).join('').substring(0, 2).toUpperCase();
    const avatar   = p.foto
      ? `<img src="${p.foto}" style="width:28px;height:28px;border-radius:50%;object-fit:cover;flex-shrink:0;" onerror="this.outerHTML='<div style=\\'width:28px;height:28px;border-radius:50%;background:linear-gradient(135deg,#3B82F6,#8B5CF6);display:flex;align-items:center;justify-content:center;font-size:0.65rem;font-weight:700;color:#fff;\\'>${initials}</div>'">`
      : `<div style="width:28px;height:28px;border-radius:50%;background:linear-gradient(135deg,#3B82F6,#8B5CF6);display:flex;align-items:center;justify-content:center;font-size:0.65rem;font-weight:700;color:#fff;flex-shrink:0;">${initials}</div>`;
    const chip = document.createElement('div');
    chip.id = 'ms-user-chip';
    chip.title = 'Cerrar sesión';
    chip.onclick = () => { if (confirm('¿Cerrar sesión?')) doLogout(); };
    chip.style.cssText = 'display:flex;align-items:center;gap:8px;cursor:pointer;padding:4px 10px 4px 6px;border-radius:20px;background:rgba(255,255,255,0.07);border:1px solid rgba(255,255,255,0.1);transition:background 0.2s;';
    chip.innerHTML = `${avatar}<div style="line-height:1.2;"><div style="font-size:12px;font-weight:600;color:#f8fafc;">${p.nombre || ''}</div>${p.cargo ? `<div style="font-size:10px;color:#94a3b8;">${p.cargo}</div>` : ''}</div>`;
    const logoutBtn = topbarRight.querySelector('.btn-logout');
    if (logoutBtn) topbarRight.insertBefore(chip, logoutBtn);
    else topbarRight.appendChild(chip);
  }

  document.getElementById('login-screen').classList.add('hidden');
  document.getElementById('app').classList.add('visible');
  clearTimeout(_sessionTimer);
  _sessionTimer = setTimeout(() => { alert('⏰ Sesión expirada. Inicia sesión nuevamente.'); doLogout(); }, SESSION_DURATION);
  initDashboard();

  // ── Montar panel de accesos en home (con pequeño delay para que el DOM esté listo) ──
  setTimeout(() => {
    if (window.LoginTracker) window.LoginTracker.renderPanel();
  }, 400);
}

// ── Logout ────────────────────────────────────────────────────────
window.doLogout = async function() {
  clearTimeout(_sessionTimer);
  window._msUserProfile = null;
  const chip = document.getElementById('ms-user-chip');
  if (chip) chip.remove();
  if (_msalInstance) {
    try {
      const accounts = _msalInstance.getAllAccounts();
      if (accounts.length) await _msalInstance.logoutPopup({ account: accounts[0] }).catch(() => {});
    } catch(e) {}
  }
  document.getElementById('app').classList.remove('visible');
  document.getElementById('login-screen').classList.remove('hidden');
  document.getElementById('login-error').style.display = 'none';
  const btn = document.getElementById('msLoginBtn');
  if (btn) { btn.disabled = false; btn.innerHTML = _msBtnHTML(); }
};