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

  // ========== PRODUK CACHE (localStorage) ==========
  // Strategi: fetch full list sekali (tanpa q), filter q di client.
  // Hasil: CALL 2+ instant 0 ms (bypass 1s Apps Script overhead).
  function _produkCacheKey() {
    const u = getUser();
    return cfg.STORAGE_KEY_PRODUK_CACHE + '_' + (u && u.role ? u.role : 'guest');
  }

  function getProdukCache() {
    try {
      const raw = localStorage.getItem(_produkCacheKey());
      if (!raw) return null;
      const obj = JSON.parse(raw);
      if (!obj || !obj.ts || !Array.isArray(obj.items)) return null;
      if (Date.now() - obj.ts > cfg.TTL_PRODUK_MS) return null;
      return obj;
    } catch (e) { return null; }
  }

  function setProdukCache(items) {
    try {
      localStorage.setItem(_produkCacheKey(), JSON.stringify({ ts: Date.now(), items: items }));
    } catch (e) {}
  }

  function clearProdukCache() {
    try {
      const u = getUser();
      const role = (u && u.role) ? u.role : 'guest';
      localStorage.removeItem(cfg.STORAGE_KEY_PRODUK_CACHE + '_' + role);
    } catch (e) {}
  }

  function _filterProduk(items, q) {
    if (!q) return items;
    const ql = q.toLowerCase();
    return items.filter(function(p) {
      return String(p.nama_produk || '').toLowerCase().indexOf(ql) !== -1;
    });
  }

  async function getProduk(q, forceRefresh) {
    const token = getToken();
    if (!token) return { ok: false, error: 'No token' };
    const ql = String(q || '').trim();
    const cache = forceRefresh ? null : getProdukCache();

    if (cache) {
      const filtered = _filterProduk(cache.items, ql);
      return { ok: true, data: { total: filtered.length, q: ql || null, items: filtered, _cached: true } };
    }

    const r = await call('getProduk', { token: token, q: '' });
    if (!r.ok || !r.data) return r;
    setProdukCache(r.data.items);
    const filtered = _filterProduk(r.data.items, ql);
    return { ok: true, data: { total: filtered.length, q: ql || null, items: filtered, _cached: false } };
  }

  // ========== BUAT PESANAN ==========
  async function buatPesanan(items, tglKirim, catatan, namaReseller) {
    const token = getToken();
    if (!token) return { ok: false, error: 'No token' };
    const body = {
      token: token,
      items: items,
      tgl_kirim: tglKirim
    };
    if (catatan) body.catatan = catatan;
    if (namaReseller) body.nama_reseller = namaReseller;
    return call('buatPesanan', body);
  }

  // ========== GET PESANAN (list) ==========
  async function getPesanan() {
    const token = getToken();
    if (!token) return { ok: false, error: 'No token' };
    return call('getPesanan', { token: token });
  }

  async function getDetailPesanan(idPesanan) {
    const token = getToken();
    if (!token) return { ok: false, error: 'No token' };
    return call('getDetailPesanan', { token: token, id_pesanan: idPesanan });
  }

  // ========== GET RESELLER LIST (autocomplete) ==========
  // Cache 24 jam di localStorage untuk instant load berikutnya.
  async function getResellerList(forceRefresh) {
    const token = getToken();
    if (!token) return { ok: false, error: 'No token' };

    const KEY = 'kk22_reseller_cache';
    const TS  = 'kk22_reseller_cache_ts';
    const TTL = 24 * 60 * 60 * 1000;

    try {
      const ts = Number(localStorage.getItem(TS) || 0);
      if (!forceRefresh && ts && (Date.now() - ts < TTL)) {
        const raw = localStorage.getItem(KEY);
        if (raw) {
          const items = JSON.parse(raw);
          if (Array.isArray(items) && items.length > 0) {
            return { ok: true, data: { items: items, _localCached: true } };
          }
        }
      }
    } catch (e) {}

    const r = await call('getResellerList', { token: token });
    if (r.ok && r.data && Array.isArray(r.data.items)) {
      try {
        localStorage.setItem(KEY, JSON.stringify(r.data.items));
        localStorage.setItem(TS, String(Date.now()));
      } catch (e) {}
    }
    return r;
  }

  function clearResellerCache() {
    try {
      localStorage.removeItem('kk22_reseller_cache');
      localStorage.removeItem('kk22_reseller_cache_ts');
    } catch (e) {}
  }

  // ========== ADMIN: LOCK / UNLOCK ITEM PESANAN ==========
  async function lockItemPesanan(idPesanan, idDetailArray) {
    const token = getToken();
    if (!token) return { ok: false, error: 'No token' };
    return call('lockItemPesanan', {
      token: token,
      id_pesanan: idPesanan,
      id_detail: idDetailArray
    });
  }

  async function unlockItemPesanan(idPesanan, idDetail, alasan) {
    const token = getToken();
    if (!token) return { ok: false, error: 'No token' };
    return call('unlockItemPesanan', {
      token: token,
      id_pesanan: idPesanan,
      id_detail: idDetail,
      alasan: alasan
    });
  }

  // ========== ADMIN: TOGGLE PRODUK AKTIF ==========
  async function toggleProdukAktif(idProduk, aktif) {
    const token = getToken();
    if (!token) return { ok: false, error: 'No token' };
    return call('toggleProdukAktif', {
      token: token,
      id_produk: idProduk,
      aktif: !!aktif
    });
  }

  // ========== ADMIN: BULK LOCK / UNLOCK ==========
  async function getProdukInPOs(idPesananList) {
    const token = getToken();
    if (!token) return { ok: false, error: 'No token' };
    return call('getProdukInPOs', { token: token, id_pesanan_list: idPesananList });
  }

  async function lockItemsBulk(idPesananList, produkFilter) {
    const token = getToken();
    if (!token) return { ok: false, error: 'No token' };
    const body = { token: token, id_pesanan_list: idPesananList };
    if (Array.isArray(produkFilter) && produkFilter.length > 0) body.produk_filter = produkFilter;
    return call('lockItemsBulk', body);
  }

  async function unlockItemsBulk(idPesananList, alasan, produkFilter) {
    const token = getToken();
    if (!token) return { ok: false, error: 'No token' };
    const body = { token: token, id_pesanan_list: idPesananList, alasan: alasan };
    if (Array.isArray(produkFilter) && produkFilter.length > 0) body.produk_filter = produkFilter;
    return call('unlockItemsBulk', body);
  }

  // ========== ADMIN: HARGA KHUSUS (F9.1) ==========
  async function getHargaKhusus(idReseller) {
    const token = getToken();
    if (!token) return { ok: false, error: 'No token' };
    return call('getHargaKhusus', { token: token, id_reseller: idReseller });
  }

  async function setHargaKhusus(idReseller, namaProduk, hargaKhusus, catatan) {
    const token = getToken();
    if (!token) return { ok: false, error: 'No token' };
    return call('setHargaKhusus', {
      token: token,
      id_reseller: idReseller,
      nama_produk: namaProduk,
      harga_khusus: hargaKhusus,
      catatan: catatan || ''
    });
  }

  async function deleteHargaKhusus(idReseller, namaProduk) {
    const token = getToken();
    if (!token) return { ok: false, error: 'No token' };
    return call('deleteHargaKhusus', {
      token: token,
      id_reseller: idReseller,
      nama_produk: namaProduk
    });
  }

  // ========== ADMIN: PRICE LISTS (F9.1b) ==========
  async function getPriceLists() {
    const token = getToken();
    if (!token) return { ok: false, error: 'No token' };
    return call('getPriceLists', { token: token });
  }

  async function getPriceListDetail(idPriceList) {
    const token = getToken();
    if (!token) return { ok: false, error: 'No token' };
    return call('getPriceListDetail', { token: token, id_price_list: idPriceList });
  }

  async function createPriceList(namaPaket, deskripsi, sourceType, sourceId, items) {
    const token = getToken();
    if (!token) return { ok: false, error: 'No token' };
    const body = {
      token: token,
      nama_paket: namaPaket,
      deskripsi: deskripsi || '',
      source_type: sourceType || 'normal'
    };
    if (sourceId) body.source_id = sourceId;
    if (Array.isArray(items) && items.length > 0) body.items = items;
    return call('createPriceList', body);
  }

  async function renamePriceList(idPriceList, namaBaru, deskripsi) {
    const token = getToken();
    if (!token) return { ok: false, error: 'No token' };
    const body = { token: token, id_price_list: idPriceList, nama_baru: namaBaru };
    if (deskripsi !== undefined && deskripsi !== null) body.deskripsi = deskripsi;
    return call('renamePriceList', body);
  }

  async function togglePriceListAktif(idPriceList, aktif) {
    const token = getToken();
    if (!token) return { ok: false, error: 'No token' };
    return call('togglePriceListAktif', { token: token, id_price_list: idPriceList, aktif: !!aktif });
  }

  async function deletePriceList(idPriceList) {
    const token = getToken();
    if (!token) return { ok: false, error: 'No token' };
    return call('deletePriceList', { token: token, id_price_list: idPriceList });
  }

  async function setPriceListItem(idPriceList, namaProduk, harga) {
    const token = getToken();
    if (!token) return { ok: false, error: 'No token' };
    return call('setPriceListItem', {
      token: token,
      id_price_list: idPriceList,
      nama_produk: namaProduk,
      harga: Number(harga)
    });
  }

  async function bulkSetPriceListItems(idPriceList, items) {
    const token = getToken();
    if (!token) return { ok: false, error: 'No token' };
    return call('bulkSetPriceListItems', {
      token: token,
      id_price_list: idPriceList,
      items: items || []
    });
  }

  async function assignPriceListToReseller(idReseller, idPriceList) {
    const token = getToken();
    if (!token) return { ok: false, error: 'No token' };
    return call('assignPriceListToReseller', {
      token: token,
      id_reseller: idReseller,
      id_price_list: idPriceList || ''
    });
  }

  async function getResellerAssignments() {
    const token = getToken();
    if (!token) return { ok: false, error: 'No token' };
    return call('getResellerAssignments', { token: token });
  }

  // ========== ADMIN: USER MANAGEMENT ==========
  async function createUser(noHp, pin, role, namaLengkap, idReseller) {
    const token = getToken();
    if (!token) return { ok: false, error: 'No token' };
    return call('createUser', {
      token: token,
      no_hp: noHp,
      pin: pin,
      role: role,
      nama_lengkap: namaLengkap,
      id_reseller: idReseller || ''
    });
  }

  async function listUsers() {
    const token = getToken();
    if (!token) return { ok: false, error: 'No token' };
    return call('listUsers', { token: token });
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
    clearAuth: clearAuth,
    getProduk: getProduk,
    buatPesanan: buatPesanan,
    getPesanan: getPesanan,
    getDetailPesanan: getDetailPesanan,
    getProdukCache: getProdukCache,
    clearProdukCache: clearProdukCache,
    getResellerList: getResellerList,
    clearResellerCache: clearResellerCache,
    lockItemPesanan: lockItemPesanan,
    unlockItemPesanan: unlockItemPesanan,
    toggleProdukAktif: toggleProdukAktif,
    createUser: createUser,
    listUsers: listUsers,
    getPriceLists: getPriceLists,
    getPriceListDetail: getPriceListDetail,
    createPriceList: createPriceList,
    renamePriceList: renamePriceList,
    togglePriceListAktif: togglePriceListAktif,
    deletePriceList: deletePriceList,
    setPriceListItem: setPriceListItem,
    bulkSetPriceListItems: bulkSetPriceListItems,
    assignPriceListToReseller: assignPriceListToReseller,
    getResellerAssignments: getResellerAssignments,
    getHargaKhusus: getHargaKhusus,
    setHargaKhusus: setHargaKhusus,
    deleteHargaKhusus: deleteHargaKhusus,
    getProdukInPOs: getProdukInPOs,
    lockItemsBulk: lockItemsBulk,
    unlockItemsBulk: unlockItemsBulk
  };
})();