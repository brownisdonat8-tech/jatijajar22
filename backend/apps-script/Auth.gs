/**
 * KK22 - Auth
 * Login, register, session management.
 *
 * Sheet SESSIONS akan otomatis dibuat saat pertama kali dipakai.
 */

const AUTH = {
  SESSION_DAYS: 30,
  SHEET_SESSIONS: 'SESSIONS'
};

// ============================================================
// SHEET SESSIONS (auto-create)
// ============================================================

function ensureSessionsSheet_() {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(AUTH.SHEET_SESSIONS);
  if (!sh) {
    sh = ss.insertSheet(AUTH.SHEET_SESSIONS);
    const headers = ['token', 'id_user', 'role', 'expired_at', 'created_at', 'last_used'];
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#0b5394')
      .setFontColor('#ffffff');
    sh.setFrozenRows(1);
    Logger.log('[OK] Sheet SESSIONS dibuat otomatis.');
  }
  return sh;
}

// ============================================================
// SETUP ADMIN PERTAMA
// ============================================================
// GANTI nilai di bawah SEBELUM jalankan setupAdminPertama().
// Setelah dijalankan, GANTI PIN di sini (jangan biarkan PIN asli tertinggal di kode).

const ADMIN_PERTAMA = {
  noHp:        '628123456789',   // GANTI - format 62...
  pin:         '1234',            // GANTI - 4 sampai 6 digit
  namaLengkap: 'Owner KK22'       // GANTI - nama kamu
};

function setupAdminPertama() {
  const noHp = normalizePhone_(ADMIN_PERTAMA.noHp);
  const pin  = String(ADMIN_PERTAMA.pin);
  const nama = ADMIN_PERTAMA.namaLengkap;

  if (!noHp || noHp.length < 10) {
    Logger.log('ERROR: noHp tidak valid. Hasil normalisasi: "' + noHp + '"');
    return;
  }
  if (pin.length < CONFIG.PIN.MIN || pin.length > CONFIG.PIN.MAX) {
    Logger.log('ERROR: PIN harus ' + CONFIG.PIN.MIN + '-' + CONFIG.PIN.MAX + ' digit.');
    return;
  }

  const existing = findUserByPhone_(noHp);
  if (existing) {
    Logger.log('ERROR: nomor HP sudah terdaftar sebagai:');
    Logger.log('       ' + existing.nama_lengkap + ' (' + existing.role + ')');
    return;
  }

  const user = createUser_({
    noHp: noHp,
    pin: pin,
    role: 'owner',
    namaLengkap: nama,
    idReseller: ''
  });

  Logger.log('===== ADMIN BERHASIL DIBUAT =====');
  Logger.log('id_user   : ' + user.id_user);
  Logger.log('no_hp     : ' + user.no_hp + '  (input asli: ' + ADMIN_PERTAMA.noHp + ')');
  Logger.log('PIN       : ' + pin + '   <<< SIMPAN, TIDAK BISA DILIHAT LAGI');
  Logger.log('nama      : ' + user.nama_lengkap);
  Logger.log('role      : owner');
  Logger.log('');
  Logger.log('PENTING: Setelah ini, buka file Auth.gs dan GANTI nilai ADMIN_PERTAMA.pin');
  Logger.log('         supaya PIN asli tidak tertinggal di kode.');
}

// ============================================================
// USER HELPERS
// ============================================================

function createUser_(opts) {
  const sheet = getSheet_(CONFIG.SHEETS.USERS);
  const lastRow = sheet.getLastRow();
  const newRow = lastRow + 1;
  const idUser = CONFIG.PREFIX.USER + pad4_(lastRow);
  const salt = makeSalt_();
  const hash = pinHash_(opts.pin, salt);
  const now = nowJakarta_();

  const row = [
    idUser,
    opts.noHp,
    hash,
    salt,
    opts.role,
    opts.idReseller || '',
    opts.namaLengkap,
    'TRUE',
    now,
    now
  ];

  sheet.getRange(newRow, 1, 1, row.length).setValues([row]);

  return {
    id_user: idUser,
    no_hp: opts.noHp,
    role: opts.role,
    nama_lengkap: opts.namaLengkap
  };
}

function findUserByPhone_(noHp) {
  const sheet = getSheet_(CONFIG.SHEETS.USERS);
  const last = sheet.getLastRow();
  if (last < 2) return null;
  const data = sheet.getRange(2, 1, last - 1, 10).getValues();
  for (let i = 0; i < data.length; i++) {
    const r = data[i];
    if (String(r[1]) === String(noHp)) {
      return {
        id_user: r[0],
        no_hp: r[1],
        pin_hash: r[2],
        pin_salt: r[3],
        role: r[4],
        id_reseller: r[5],
        nama_lengkap: r[6],
        aktif: r[7],
        tgl_dibuat: r[8],
        tgl_update_terakhir: r[9],
        rowIndex: i + 2
      };
    }
  }
  return null;
}

function previewUsers() {
  const sheet = getSheet_(CONFIG.SHEETS.USERS);
  const last = sheet.getLastRow();
  if (last < 2) {
    Logger.log('Belum ada user.');
    return;
  }
  const data = sheet.getRange(2, 1, last - 1, 10).getValues();
  Logger.log('===== DAFTAR USER (' + data.length + ') =====');
  data.forEach(function(r, i) {
    Logger.log((i+1) + '. [' + r[4] + '] ' + r[6] + ' | ' + r[1] + ' | aktif=' + r[7]);
  });
}

// ============================================================
// TEST SESSION SHEET
// ============================================================

function testSessionsSheet() {
  const sh = ensureSessionsSheet_();
  Logger.log('Sheet SESSIONS siap. Total baris: ' + sh.getLastRow());
}
// ============================================================
// RESET PIN (untuk admin/owner/reseller yang lupa PIN)
// ============================================================
// Cara pakai:
//   1. Edit nilai RESET_PIN di bawah
//   2. Jalankan resetPinUser() dari dropdown
//   3. Login ulang di API dengan PIN baru

const RESET_PIN = {
  noHp:     '6282114188277',   // <-- no_hp admin kamu
  pinBaru:  '123456'            // <-- GANTI jadi PIN baru 4-6 karakter
};

function resetPinUser() {
  const noHp = normalizePhone_(RESET_PIN.noHp);
  const pinBaru = String(RESET_PIN.pinBaru);

  if (!noHp || noHp.length < 10) {
    Logger.log('ERROR: noHp tidak valid: "' + noHp + '"');
    return;
  }
  if (pinBaru.length < CONFIG.PIN.MIN || pinBaru.length > CONFIG.PIN.MAX) {
    Logger.log('ERROR: PIN baru harus ' + CONFIG.PIN.MIN + '-' + CONFIG.PIN.MAX + ' karakter. Dapat: ' + pinBaru.length);
    return;
  }

  const user = findUserByPhone_(noHp);
  if (!user) {
    Logger.log('ERROR: nomor HP tidak terdaftar: ' + noHp);
    Logger.log('       Cek no_hp di sheet USERS.');
    return;
  }

  const sheet = getSheet_(CONFIG.SHEETS.USERS);
  const salt = makeSalt_();
  const hash = pinHash_(pinBaru, salt);
  const now = nowJakarta_();

  // Update kolom C (pin_hash, index 3), D (pin_salt, index 4), J (tgl_update, index 10)
  sheet.getRange(user.rowIndex, 3).setValue(hash);
  sheet.getRange(user.rowIndex, 4).setValue(salt);
  sheet.getRange(user.rowIndex, 10).setValue(now);

  // INVALIDATE cache user (PIN berubah)
  cacheInvalidate_(['user_' + user.id_user]);

  // Hapus semua session lama user ini (paksa login ulang)
  const sess = ensureSessionsSheet_();
  const last = sess.getLastRow();
  if (last >= 2) {
    const data = sess.getRange(2, 1, last - 1, 2).getValues();
    for (let i = data.length - 1; i >= 0; i--) {
      if (String(data[i][1]) === String(user.id_user)) {
        sess.deleteRow(i + 2);
      }
    }
  }

  Logger.log('===== RESET PIN BERHASIL =====');
  Logger.log('id_user    : ' + user.id_user);
  Logger.log('no_hp      : ' + user.no_hp);
  Logger.log('nama       : ' + user.nama_lengkap);
  Logger.log('role       : ' + user.role);
  Logger.log('PIN baru   : ' + pinBaru + '   <<< SIMPAN BAIK-BAIK');
  Logger.log('');
  Logger.log('PENTING: Setelah reset, buka Auth.gs dan GANTI RESET_PIN.pinBaru');
  Logger.log('         supaya PIN tidak tertinggal di kode.');
  Logger.log('');
  Logger.log('Sekarang TEST LOGIN LAGI via PowerShell dengan PIN baru.');
}
// ============================================================
// SESSION VALIDATION (CACHED)
// ============================================================
// Cache 2 menit untuk hindari baca sheet SESSIONS setiap request.
// Trade-off: kalau admin logout user paksa, ada delay max 2 menit.

function validateSessionCached_(token) {
  if (!token) return null;

  const cacheKey = 'session_' + token;
  let cached = cacheGet_(cacheKey);
  if (cached) return cached;

  const sh = ensureSessionsSheet_();
  const last = sh.getLastRow();
  if (last < 2) return null;

  const data = sh.getRange(2, 1, last - 1, 6).getValues();
  const now = new Date();
  for (let i = 0; i < data.length; i++) {
    const r = data[i];
    if (String(r[0]) === String(token)) {
      const exp = new Date(String(r[3]).replace(' ', 'T'));
      if (exp.getTime() < now.getTime()) return null;
      try {
        sh.getRange(i + 2, 6).setValue(
          Utilities.formatDate(now, CONFIG.TZ, 'yyyy-MM-dd HH:mm:ss')
        );
      } catch (e) {}
      const user = {
        id_user: String(r[1]),
        role: String(r[2])
      };
      // Enrich dengan data user lengkap (dari cache kalau ada)
      const userFull = getCachedUser_(user.id_user);
      if (userFull) {
        cachePut_(cacheKey, userFull, CACHE.TTL_SESSION);
        return userFull;
      }
      cachePut_(cacheKey, user, CACHE.TTL_SESSION);
      return user;
    }
  }
  return null;
}

// ============================================================
// USER MANAGEMENT (endpoint: createUser, listUsers)
// ============================================================

function handleCreateUser_(body) {
  const user = validateToken_(body.token);
  if (!user) return jsonResponse_({ ok: false, error: 'Token tidak valid atau expired.' });
  if (user.role !== 'owner' && user.role !== 'admin') {
    return jsonResponse_({ ok: false, error: 'Hanya owner/admin yang bisa membuat user baru.' });
  }

  const noHpRaw = String(body.no_hp || '').trim();
  const pin = String(body.pin || '').trim();
  const namaLengkap = String(body.nama_lengkap || '').trim();
  const role = String(body.role || '').trim().toLowerCase();
  const idReseller = String(body.id_reseller || '').trim();

  if (!noHpRaw) return jsonResponse_({ ok: false, error: 'no_hp wajib diisi.' });
  const noHp = normalizePhone_(noHpRaw);
  if (!noHp || noHp.length < 10) return jsonResponse_({ ok: false, error: 'no_hp tidak valid (hasil normalisasi: "' + noHp + '").' });

  if (!namaLengkap) return jsonResponse_({ ok: false, error: 'nama_lengkap wajib diisi.' });
  if (pin.length < CONFIG.PIN.MIN || pin.length > CONFIG.PIN.MAX) {
    return jsonResponse_({ ok: false, error: 'PIN harus ' + CONFIG.PIN.MIN + '-' + CONFIG.PIN.MAX + ' digit.' });
  }

  const validRoles = ['owner', 'admin', 'reseller'];
  if (validRoles.indexOf(role) === -1) {
    return jsonResponse_({ ok: false, error: 'Role tidak valid. Pilih: ' + validRoles.join(', ') });
  }

  // Guard: hanya owner yang bisa membuat user role owner
  if (role === 'owner' && user.role !== 'owner') {
    return jsonResponse_({ ok: false, error: 'Hanya owner yang bisa membuat user role owner.' });
  }

  if (role === 'reseller' && !idReseller) {
    return jsonResponse_({ ok: false, error: 'id_reseller wajib untuk role reseller.' });
  }

  const existing = findUserByPhone_(noHp);
  if (existing) {
    return jsonResponse_({ ok: false, error: 'Nomor HP sudah terdaftar: ' + existing.nama_lengkap + ' (' + existing.role + ')' });
  }

  try {
    const newUser = createUser_({
      noHp: noHp,
      pin: pin,
      role: role,
      namaLengkap: namaLengkap,
      idReseller: role === 'reseller' ? idReseller : ''
    });
    try { logAktivitas_(user, 'createUser', 'Buat user ' + newUser.id_user + ' role=' + role + ' no_hp=' + noHp); } catch (e) {}
    return jsonResponse_({ ok: true, data: { user: newUser } });
  } catch (e) {
    return jsonResponse_({ ok: false, error: 'Gagal membuat user: ' + e.message });
  }
}

function handleListUsers_(body) {
  const user = validateToken_(body.token);
  if (!user) return jsonResponse_({ ok: false, error: 'Token tidak valid atau expired.' });
  if (user.role !== 'owner' && user.role !== 'admin') {
    return jsonResponse_({ ok: false, error: 'Hanya owner/admin yang bisa lihat daftar user.' });
  }

  const sh = getSheet_(CONFIG.SHEETS.USERS);
  const last = sh.getLastRow();
  const items = [];
  if (last >= 2) {
    const data = sh.getRange(2, 1, last - 1, 10).getValues();
    for (let i = 0; i < data.length; i++) {
      const r = data[i];
      const idUser = String(r[0] || '');
      if (!idUser) continue;
      items.push({
        id_user: idUser,
        no_hp: String(r[1] || ''),
        role: String(r[4] || ''),
        id_reseller: String(r[5] || ''),
        nama_lengkap: String(r[6] || ''),
        aktif: String(r[7] || '').toUpperCase() === 'TRUE',
        tgl_dibuat: String(r[8] || '')
      });
    }
  }

  const roleOrder = { owner: 0, admin: 1, reseller: 2 };
  items.sort(function(a, b) {
    const ra = roleOrder[a.role] != null ? roleOrder[a.role] : 99;
    const rb = roleOrder[b.role] != null ? roleOrder[b.role] : 99;
    if (ra !== rb) return ra - rb;
    return String(a.id_user).localeCompare(String(b.id_user));
  });

  return jsonResponse_({ ok: true, data: { total: items.length, items: items } });
}

function getCachedUser_(idUser) {
  const cacheKey = 'user_' + idUser;
  let cached = cacheGet_(cacheKey);
  if (cached) return cached;

  const sheet = getSheet_(CONFIG.SHEETS.USERS);
  const last = sheet.getLastRow();
  if (last < 2) return null;

  const data = sheet.getRange(2, 1, last - 1, 10).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(idUser)) {
      const user = {
        id_user: String(data[i][0]),
        no_hp: String(data[i][1]),
        role: String(data[i][4]),
        id_reseller: String(data[i][5] || ''),
        nama_lengkap: String(data[i][6]),
        aktif: data[i][7]
      };
      cachePut_(cacheKey, user, CACHE.TTL_SESSION);
      return user;
    }
  }
  return null;
}
