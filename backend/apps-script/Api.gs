/**
 * KK22 - API
 * Endpoint utama: doGet & doPost (dipanggil dari frontend via fetch)
 *
 * Action yang didukung:
 *   - ping      : health check
 *   - login     : { no_hp, pin } -> { token, user }
 *   - logout    : { token }
 *   - whoami    : { token } -> { user }
 */

// ============================================================
// ROUTER
// ============================================================

function doGet(e) {
  return jsonResponse_({
    ok: true,
    service: 'KK22 Backend',
    time: nowJakarta_(),
    version: String(readConfig_('versi_aplikasi') || 'unknown')
  });
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse_({ ok: false, error: 'Body kosong' });
    }
    const body = JSON.parse(e.postData.contents);
    const action = String(body.action || '');

    switch (action) {
      case 'ping':    return jsonResponse_({ ok: true, time: nowJakarta_() });
      case 'login':   return handleLogin_(body);
      case 'logout':  return handleLogout_(body);
      case 'whoami':  return handleWhoami_(body);
      case 'getProduk': return handleGetProduk_(body);
      case 'buatPesanan':        return handleBuatPesanan_(body);
      case 'getPesanan':         return handleGetPesanan_(body);
      case 'getDetailPesanan':   return handleGetDetailPesanan_(body);
      case 'updatePesananItems': return handleUpdatePesananItems_(body);
      case 'lockItemPesanan':    return handleLockItemPesanan_(body);
      case 'unlockItemPesanan':  return handleUnlockItemPesanan_(body);
      case 'toggleProdukAktif':  return handleToggleProdukAktif_(body);
      case 'getProdukInPOs':     return handleGetProdukInPOs_(body);
      case 'lockItemsBulk':      return handleLockItemsBulk_(body);
      case 'unlockItemsBulk':    return handleUnlockItemsBulk_(body);
      case 'getHargaKhusus':      return handleGetHargaKhusus_(body);
      case 'setHargaKhusus':      return handleSetHargaKhusus_(body);
      case 'deleteHargaKhusus':   return handleDeleteHargaKhusus_(body);
      case 'getPriceLists':             return handleGetPriceLists_(body);
      case 'getPriceListDetail':        return handleGetPriceListDetail_(body);
      case 'createPriceList':           return handleCreatePriceList_(body);
      case 'renamePriceList':           return handleRenamePriceList_(body);
      case 'togglePriceListAktif':      return handleTogglePriceListAktif_(body);
      case 'deletePriceList':           return handleDeletePriceList_(body);
      case 'setPriceListItem':          return handleSetPriceListItem_(body);
      case 'bulkSetPriceListItems':     return handleBulkSetPriceListItems_(body);
      case 'assignPriceListToReseller': return handleAssignPriceListToReseller_(body);
case 'bulkAssignPriceList':       return handleBulkAssignPriceList_(body);
      case 'getResellerAssignments':    return handleGetResellerAssignments_(body);
      case 'getResellerList':    return handleGetResellerList_(body);
      case 'createUser':         return handleCreateUser_(body);
      case 'listUsers':          return handleListUsers_(body);
      default:        return jsonResponse_({ ok: false, error: 'Action tidak dikenal: ' + action });
    }
  } catch (err) {
    return jsonResponse_({ ok: false, error: 'Server error: ' + err.message });
  }
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// LOGIN
// ============================================================

function handleLogin_(body) {
  const noHp = normalizePhone_(body.no_hp || '');
  const pin  = String(body.pin || '');

  if (!noHp || noHp.length < 10) {
    return jsonResponse_({ ok: false, error: 'Nomor HP tidak valid' });
  }
  if (!pin || pin.length < CONFIG.PIN.MIN || pin.length > CONFIG.PIN.MAX) {
    return jsonResponse_({ ok: false, error: 'PIN tidak valid' });
  }

  const user = findUserByPhone_(noHp);
  if (!user) {
    return jsonResponse_({ ok: false, error: 'Nomor HP atau PIN salah' });
  }
  if (String(user.aktif).toUpperCase() !== 'TRUE') {
    return jsonResponse_({ ok: false, error: 'Akun tidak aktif. Hubungi admin.' });
  }

  const computed = pinHash_(pin, user.pin_salt);
  if (computed !== user.pin_hash) {
    return jsonResponse_({ ok: false, error: 'Nomor HP atau PIN salah' });
  }

  const token = createSession_(user);
  logAktivitas_(user, 'login', 'Login via API');

  return jsonResponse_({
    ok: true,
    data: {
      token: token,
      user: {
        id_user: user.id_user,
        no_hp: user.no_hp,
        role: user.role,
        nama_lengkap: user.nama_lengkap,
        id_reseller: user.id_reseller
      }
    }
  });
}

// ============================================================
// SESSION
// ============================================================

function createSession_(user) {
  const sh = ensureSessionsSheet_();
  const token = Utilities.getUuid().replace(/-/g, '');
  const now = new Date();
  const expired = new Date(now.getTime() + AUTH.SESSION_DAYS * 24 * 60 * 60 * 1000);
  const fmt = function(d) { return Utilities.formatDate(d, CONFIG.TZ, 'yyyy-MM-dd HH:mm:ss'); };
  sh.appendRow([token, user.id_user, user.role, fmt(expired), fmt(now), fmt(now)]);
  return token;
}

function validateToken_(token) {
  if (!token) return null;
  // Delegate ke versi cached (2 menit TTL).
  return validateSessionCached_(token);
}

function handleWhoami_(body) {
  const user = validateToken_(body.token);
  if (!user) {
    return jsonResponse_({ ok: false, error: 'Token tidak valid atau sudah expired' });
  }
  return jsonResponse_({
    ok: true,
    data: {
      id_user: user.id_user,
      no_hp: user.no_hp,
      role: user.role,
      nama_lengkap: user.nama_lengkap,
      id_reseller: user.id_reseller
    }
  });
}

function handleLogout_(body) {
  const token = body.token;
  if (!token) {
    return jsonResponse_({ ok: false, error: 'Token kosong' });
  }
  const sh = ensureSessionsSheet_();
  const last = sh.getLastRow();
  if (last >= 2) {
    const data = sh.getRange(2, 1, last - 1, 1).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]) === String(token)) {
        sh.deleteRow(i + 2);
        cacheBustSession_(token);
        return jsonResponse_({ ok: true });
      }
    }
  }
  return jsonResponse_({ ok: true });
}

function getUserById_(idUser) {
  const sheet = getSheet_(CONFIG.SHEETS.USERS);
  const last = sheet.getLastRow();
  if (last < 2) return null;
  const data = sheet.getRange(2, 1, last - 1, 10).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(idUser)) {
      return {
        id_user: data[i][0],
        no_hp: data[i][1],
        pin_hash: data[i][2],
        pin_salt: data[i][3],
        role: data[i][4],
        id_reseller: data[i][5],
        nama_lengkap: data[i][6],
        aktif: data[i][7],
        rowIndex: i + 2
      };
    }
  }
  return null;
}

// ============================================================
// LOG & CONFIG
// ============================================================

function logAktivitas_(user, aksi, detail) {
  try {
    const sh = getSheet_(CONFIG.SHEETS.LOG_AKTIVITAS);
    const last = sh.getLastRow();
    const id = 'KK22-LOG-' + pad4_(last);
    sh.appendRow([id, nowJakarta_(), user ? user.id_user : '-', aksi, detail || '', '']);
  } catch (e) { /* silent */ }
}

function readConfig_(key) {
  try {
    const sh = getSheet_(CONFIG.SHEETS.CONFIG);
    const last = sh.getLastRow();
    if (last < 2) return null;
    const data = sh.getRange(2, 1, last - 1, 2).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]) === key) return data[i][1];
    }
  } catch (e) {}
  return null;
}
// ============================================================
// GET PRODUK (katalog untuk reseller)
// ============================================================

function getIdPriceListForReseller_(idReseller) {
  if (!idReseller) return '';
  const cacheKey = 'reseller_pl_' + idReseller;
  const cached = cacheGet_(cacheKey);
  if (cached !== null && cached !== undefined) return cached;
  const rInfo = findResellerById_(idReseller);
  const idPL = rInfo ? (rInfo.id_price_list || '') : '';
  cachePut_(cacheKey, idPL, 300);
  return idPL;
}

function handleGetProduk_(body) {
  const user = validateToken_(body.token);
  if (!user) {
    return jsonResponse_({ ok: false, error: 'Token tidak valid atau sudah expired. Silakan login ulang.' });
  }

  const q = String(body.q || '').trim().toLowerCase();
  const isAdmin = (user.role === 'admin' || user.role === 'owner');
  const all = getProdukListCached_();

  let hargaOverride = {};
  let priceListOverride = {};
  if (!isAdmin && user.role === 'reseller' && user.id_reseller) {
    hargaOverride = getHargaOverrideForReseller_(user.id_reseller);
    const idPL = getIdPriceListForReseller_(user.id_reseller);
    if (idPL) {
      priceListOverride = getPriceListOverrideForReseller_(idPL);
    }
  }

  const items = [];
  for (let i = 0; i < all.length; i++) {
    const p = all[i];
    if (!isAdmin && !p.aktif) continue;
    if (q && String(p.nama_produk || '').toLowerCase().indexOf(q) === -1) continue;
    const key = String(p.nama_produk || '').toLowerCase();
    const customOvr = hargaOverride[key];
    const plOvr = priceListOverride[key];
    if (customOvr) {
      // Layer 3: custom override per reseller
      const cloned = Object.assign({}, p);
      cloned.harga_default = p.harga_jual;
      cloned.harga_jual = customOvr.harga_khusus;
      cloned.harga_khusus = true;
      items.push(cloned);
    } else if (plOvr) {
      // Layer 2: paket harga (price list)
      const cloned2 = Object.assign({}, p);
      cloned2.harga_default = p.harga_jual;
      cloned2.harga_jual = plOvr.harga;
      cloned2.harga_price_list = true;
      items.push(cloned2);
    } else {
      items.push(p);
    }
  }

  return jsonResponse_({
    ok: true,
    data: {
      total: items.length,
      q: q || null,
      items: items
    }
  });
}

function readConfigCached_(key) {
  let all = cacheGet_('config_all');
  if (!all) {
    all = {};
    try {
      const sh = getSheet_(CONFIG.SHEETS.CONFIG);
      const last = sh.getLastRow();
      if (last >= 2) {
        const data = sh.getRange(2, 1, last - 1, 2).getValues();
        for (let i = 0; i < data.length; i++) {
          all[String(data[i][0])] = data[i][1];
        }
      }
    } catch (e) {}
    cachePut_('config_all', all, CACHE.TTL_CONFIG);
  }
  return all[key] || null;
}