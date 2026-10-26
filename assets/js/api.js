/**
 * KK22 - API Helper
 * Wrapper untuk fetch ke Apps Script.
 *
 * PENTING: Apps Script tidak handle CORS preflight OPTIONS.
 * Solusi: gunakan Content-Type: text/plain (bukan application/json).
 * Apps Script tetap bisa parse body via e.postData.contents.
 */

window.KK22_API = (function() {
  const cfg = window.KK22_CONFIG;

  async function call(action, params) {
    const body = Object.assign({ action: action }, params || {});
    try {
      const res = await fetch(cfg.API_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify(body)
      });
      if (!res.ok) throw new Error('HTTP ' + res.status);
      const data = await res.json();
      return data;
    } catch (err) {
      return { ok: false, error: 'Network error: ' + err.message };
    }
  }

  function getToken() {
    try { return localStorage.getItem(cfg.STORAGE_KEY_TOKEN) || ''; }
    catch (e) { return ''; }
  }

  function setToken(token) {
    try { localStorage.setItem(cfg.STORAGE_KEY_TOKEN, token); } catch (e) {}
  }

  function getUser() {
    try {
      const raw = localStorage.getItem(cfg.STORAGE_KEY_USER);
      return raw ? JSON.parse(raw) : null;
    } catch (e) { return null; }
  }

  function setUser(user) {
    try { localStorage.setItem(cfg.STORAGE_KEY_USER, JSON.stringify(user)); } catch (e) {}
  }

  function clearAuth() {
    try {
      localStorage.removeItem(cfg.STORAGE_KEY_TOKEN);
      localStorage.removeItem(cfg.STORAGE_KEY_USER);
    } catch (e) {}
  }

  async function login(noHp, pin) {
    const r = await call('login', { no_hp: noHp, pin: pin });
    if (r.ok && r.data) {
      setToken(r.data.token);
      setUser(r.data.user);
    }
    return r;
  }

  async function logout() {
    const token = getToken();
    if (token) await call('logout', { token: token });
    clearAuth();
  }

  async function whoami() {
    const token = getToken();
    if (!token) return { ok: false, error: 'No token' };
    return call('whoami', { token: token });
  }

  return {
    call: call,
    login: login,
    logout: logout,
    whoami: whoami,
    getToken: getToken,
    setToken: setToken,
    getUser: getUser,
    setUser: setUser,
    clearAuth: clearAuth
  };
})();