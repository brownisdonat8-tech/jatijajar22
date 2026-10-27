/**
 * KK22 - Pesanan
 * Endpoint: buatPesanan, getPesanan, getDetailPesanan
 *
 * Aturan bisnis:
 *   - 1 reseller hanya boleh punya 1 PO aktif per tanggal kirim
 *   - Kalau sudah ada PO aktif dengan tanggal kirim sama -> auto-merge:
 *       * Item baru dengan produk belum ada -> tambah baris (pending)
 *       * Item baru dengan produk sudah ada & masih pending -> gabung qty
 *       * Item baru dengan produk sudah ada & sudah locked -> tolak item itu
 *   - Reseller tidak bisa pesan produk yang aktif = FALSE
 *   - Harga di-snapshot dari master produk saat pesan
 */

// ============================================================
// BUAT PESANAN
// ============================================================

function handleBuatPesanan_(body) {
  const user = validateToken_(body.token);
  if (!user) {
    return jsonResponse_({ ok: false, error: 'Token tidak valid atau sudah expired. Silakan login ulang.' });
  }

  // Tentukan reseller target
  let reseller;
  if (user.role === 'reseller') {
    if (!user.id_reseller) return jsonResponse_({ ok: false, error: 'Akun reseller tidak memiliki ID reseller.' });
    reseller = findResellerById_(user.id_reseller);
    if (!reseller) return jsonResponse_({ ok: false, error: 'Data reseller tidak ditemukan.' });
  } else if (user.role === 'admin' || user.role === 'owner') {
    if (!body.nama_reseller) return jsonResponse_({ ok: false, error: 'Admin/owner harus isi nama_reseller.' });
    reseller = findResellerByName_(body.nama_reseller);
    if (!reseller) return jsonResponse_({ ok: false, error: 'Reseller tidak ditemukan: ' + body.nama_reseller });
  } else {
    return jsonResponse_({ ok: false, error: 'Role tidak diizinkan.' });
  }

  if (String(reseller.aktif).toUpperCase() !== 'TRUE') {
    return jsonResponse_({ ok: false, error: 'Reseller tidak aktif.' });
  }

  // Validasi items
  const rawItems = Array.isArray(body.items) ? body.items : [];
  if (rawItems.length === 0) {
    return jsonResponse_({ ok: false, error: 'Minimal 1 item dalam pesanan.' });
  }

  // Tentukan tgl kirim
  const tglKirim = String(body.tgl_kirim || '').trim() || defaultTglKirim_();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(tglKirim)) {
    return jsonResponse_({ ok: false, error: 'Format tgl_kirim harus YYYY-MM-DD.' });
  }

  // Enrich & validasi items terhadap master produk
  const validasi = validasiItems_(rawItems);
  if (!validasi.ok) return jsonResponse_({ ok: false, error: validasi.error });

  // Cari PO existing dengan tgl kirim sama
  const existingPO = findPOAktif_(reseller.id_reseller, tglKirim);

  if (existingPO) {
    return mergeKeExistingPO_(existingPO, validasi.items, user, body);
  } else {
    return createNewPO_(reseller, tglKirim, validasi.items, user, body);
  }
}

// ============================================================
// VALIDASI ITEMS
// ============================================================

function validasiItems_(rawItems) {
  const sh = getSheet_(CONFIG.SHEETS.PRODUK);
  const last = sh.getLastRow();
  if (last < 2) return { ok: false, error: 'Master produk kosong.' };

  const data = sh.getRange(2, 1, last - 1, 20).getValues();
  const produkMap = {};
  data.forEach(function(r) {
    const nama = String(r[0] || '').trim();
    if (!nama) return;
    produkMap[nama.toLowerCase()] = {
      nama_produk: nama,
      kode_sku:    String(r[2] || ''),
      satuan:      String(r[5] || ''),
      harga_jual:  Number(r[9]) || 0,
      aktif:       String(r[17] || '').toUpperCase() === 'TRUE'
    };
  });

  const items = [];
  for (let i = 0; i < rawItems.length; i++) {
    const it = rawItems[i];
    const nama = String(it.nama_produk || '').trim();
    const qty  = Number(it.qty) || 0;

    if (!nama) return { ok: false, error: 'Item ke-' + (i+1) + ': nama_produk kosong.' };
    if (qty <= 0) return { ok: false, error: 'Item ke-' + (i+1) + ' (' + nama + '): qty harus > 0.' };

    const master = produkMap[nama.toLowerCase()];
    if (!master) return { ok: false, error: 'Item ke-' + (i+1) + ': produk tidak ditemukan di master: ' + nama };
    if (!master.aktif) return { ok: false, error: 'Item ke-' + (i+1) + ': produk sedang tidak tersedia (nonaktif): ' + nama };
    if (master.harga_jual <= 0) return { ok: false, error: 'Item ke-' + (i+1) + ': harga jual produk belum diset: ' + nama };

    items.push({
      nama_produk: master.nama_produk,
      kode_sku:    master.kode_sku,
      satuan:      master.satuan,
      qty:         qty,
      harga_saat_pesan: master.harga_jual,
      subtotal:    qty * master.harga_jual,
      catatan:     String(it.catatan || '').trim()
    });
  }
  return { ok: true, items: items };
}

// ============================================================
// CREATE NEW PO
// ============================================================

function createNewPO_(reseller, tglKirim, items, user, body) {
  const sh = getSheet_(CONFIG.SHEETS.PESANAN);
  const lastRow = sh.getLastRow();
  const newRow = lastRow + 1;
  const idPesanan = CONFIG.PREFIX.PESANAN + pad4_(lastRow);
  const customerRefNo = CONFIG.PREFIX.ORDER + pad4_(lastRow);
  const orderNumber = generateOrderNumber_(reseller.nama_reseller, tglKirim);

  const now = nowJakarta_();
  const tglPesan = todayJakarta_();

  const totalItem = items.length;
  const totalQty = items.reduce(function(s, x) { return s + x.qty; }, 0);
  const totalHarga = items.reduce(function(s, x) { return s + x.subtotal; }, 0);

  // Header
  const headerRow = [
    idPesanan,
    orderNumber,
    customerRefNo,
    reseller.nama_reseller,
    reseller.id_reseller,
    tglPesan,
    tglKirim,
    '',                     // shipping_date (diisi nanti admin)
    '',                     // ship_via
    '',                     // tracking_no
    'submitted',
    totalItem,
    totalQty,
    totalHarga,
    String(body.catatan || '').trim(),
    String(body.memo || '').trim(),
    '',
    now,
    user.id_user,
    now
  ];
  sh.getRange(newRow, 1, 1, headerRow.length).setValues([headerRow]);

  // Detail
  const detailSh = getSheet_(CONFIG.SHEETS.DETAIL_PESANAN);
  const detailLastRow = detailSh.getLastRow();
  const detailRows = [];
  let counter = detailLastRow;

  items.forEach(function(it, idx) {
    counter++;
    const idDetail = 'KK22-DTL-' + pad4_(counter);
    detailRows.push([
      idDetail,
      idPesanan,
      orderNumber,
      it.nama_produk,
      it.kode_sku,
      it.qty,
      it.satuan,
      it.harga_saat_pesan,
      0,
      it.subtotal,
      it.catatan,
      'pending',
      '',
      ''
    ]);
  });

  if (detailRows.length > 0) {
    detailSh.getRange(detailLastRow + 1, 1, detailRows.length, 14).setValues(detailRows);
  }

  logAktivitas_(user, 'buatPesanan', 'PO baru: ' + orderNumber + ' (total ' + totalQty + ' item)');

  return jsonResponse_({
    ok: true,
    data: {
      mode: 'new',
      id_pesanan: idPesanan,
      order_number: orderNumber,
      customer_ref_no: customerRefNo,
      nama_reseller: reseller.nama_reseller,
      tgl_kirim: tglKirim,
      status: 'submitted',
      total_item: totalItem,
      total_qty: totalQty,
      total_harga: totalHarga
    }
  });
}

// ============================================================
// MERGE KE EXISTING PO
// ============================================================

function mergeKeExistingPO_(po, newItems, user, body) {
  const headerSh = getSheet_(CONFIG.SHEETS.PESANAN);
  const detailSh = getSheet_(CONFIG.SHEETS.DETAIL_PESANAN);

  // Baca semua detail PO existing
  const lastDetail = detailSh.getLastRow();
  const detailData = detailSh.getRange(2, 1, Math.max(0, lastDetail - 1), 14).getValues();

  // Index item existing by nama_produk (lowercase)
  const existingIndex = {};
  for (let i = 0; i < detailData.length; i++) {
    if (String(detailData[i][1]) !== String(po.id_pesanan)) continue;
    const key = String(detailData[i][3]).toLowerCase();
    existingIndex[key] = {
      rowIndex: i + 2,
      qty: Number(detailData[i][5]) || 0,
      status: String(detailData[i][11] || 'pending').toLowerCase(),
      harga: Number(detailData[i][7]) || 0,
      subtotal: Number(detailData[i][9]) || 0,
      data: detailData[i]
    };
  }

  const accepted = [];    // {nama_produk, qty, aksi: 'added'|'merged'}
  const rejected = [];    // {nama_produk, qty, alasan}

  // Proses tiap item baru
  newItems.forEach(function(it) {
    const key = it.nama_produk.toLowerCase();
    const existing = existingIndex[key];

    if (existing) {
      if (existing.status === 'pending') {
        // Merge qty
        const newQty = existing.qty + it.qty;
        const newSubtotal = newQty * existing.harga;
        detailSh.getRange(existing.rowIndex, 6).setValue(newQty);
        detailSh.getRange(existing.rowIndex, 10).setValue(newSubtotal);
        accepted.push({ nama_produk: it.nama_produk, qty: it.qty, aksi: 'merged' });
      } else {
        // locked/shipped/done -> tolak
        rejected.push({ nama_produk: it.nama_produk, qty: it.qty, alasan: 'sudah ' + existing.status + ', hubungi WA admin' });
      }
    } else {
      // Tambah baris baru (pending)
      const counter = detailSh.getLastRow();
      const idDetail = 'KK22-DTL-' + pad4_(counter + 1);
      detailSh.appendRow([
        idDetail,
        po.id_pesanan,
        po.order_number,
        it.nama_produk,
        it.kode_sku,
        it.qty,
        it.satuan,
        it.harga_saat_pesan,
        0,
        it.subtotal,
        it.catatan,
        'pending',
        '',
        ''
      ]);
      accepted.push({ nama_produk: it.nama_produk, qty: it.qty, aksi: 'added' });
    }
  });

  // Recalculate header
  const rekap = rekapDetail_(po.id_pesanan);

  // Update status header
  let newStatus = 'submitted';
  if (rekap.total_item_locked > 0 && rekap.total_item_pending > 0) newStatus = 'partial_locked';
  else if (rekap.total_item_locked === rekap.total_item && rekap.total_item > 0) newStatus = 'locked';

  headerSh.getRange(po.rowIndex, 11).setValue(newStatus);
  headerSh.getRange(po.rowIndex, 12).setValue(rekap.total_item);
  headerSh.getRange(po.rowIndex, 13).setValue(rekap.total_qty);
  headerSh.getRange(po.rowIndex, 14).setValue(rekap.total_harga);
  headerSh.getRange(po.rowIndex, 19).setValue(user.id_user);
  headerSh.getRange(po.rowIndex, 20).setValue(nowJakarta_());

  logAktivitas_(user, 'mergePesanan', 'Merge ke ' + po.order_number + ' (acc ' + accepted.length + ', reject ' + rejected.length + ')');

  return jsonResponse_({
    ok: true,
    data: {
      mode: 'merge',
      id_pesanan: po.id_pesanan,
      order_number: po.order_number,
      customer_ref_no: po.customer_ref_no,
      nama_reseller: po.nama_reseller,
      tgl_kirim: po.tgl_kirim,
      status: newStatus,
      accepted: accepted,
      rejected: rejected,
      total_item: rekap.total_item,
      total_qty: rekap.total_qty,
      total_harga: rekap.total_harga
    }
  });
}

// ============================================================
// GET PESANAN (LIST)
// ============================================================

function handleGetPesanan_(body) {
  const user = validateToken_(body.token);
  if (!user) return jsonResponse_({ ok: false, error: 'Token tidak valid atau sudah expired.' });

  const sh = getSheet_(CONFIG.SHEETS.PESANAN);
  const last = sh.getLastRow();
  if (last < 2) return jsonResponse_({ ok: true, data: { total: 0, items: [] } });

  const data = sh.getRange(2, 1, last - 1, 20).getValues();
  const filterNama = body.nama_reseller ? String(body.nama_reseller).toLowerCase() : null;
  const filterStatus = body.status ? String(body.status).toLowerCase() : null;

  const list = [];
  for (let i = 0; i < data.length; i++) {
    const r = data[i];
    const namaReseller = String(r[3] || '');

    // Isolasi: reseller hanya lihat punya sendiri
    if (user.role === 'reseller' && String(r[4]) !== String(user.id_reseller)) continue;

    // Admin filter
    if (filterNama && namaReseller.toLowerCase().indexOf(filterNama) === -1) continue;
    if (filterStatus && String(r[10]).toLowerCase() !== filterStatus) continue;

    list.push({
      id_pesanan:    String(r[0] || ''),
      order_number:  String(r[1] || ''),
      customer_ref_no: String(r[2] || ''),
      nama_reseller: namaReseller,
      tgl_pesan:     toYMD_(r[5]),
      tgl_kirim:     toYMD_(r[6]),
      tgl_kirim_aktual: toYMD_(r[7]),
      status:        String(r[10] || ''),
      total_item:    Number(r[11]) || 0,
      total_qty:     Number(r[12]) || 0,
      total_harga:   Number(r[13]) || 0,
      tgl_dibuat:    String(r[17] || '')
    });
  }

  // Sort by tgl_dibuat DESC
  list.sort(function(a, b) { return String(b.tgl_dibuat).localeCompare(String(a.tgl_dibuat)); });

  return jsonResponse_({ ok: true, data: { total: list.length, items: list } });
}

// ============================================================
// GET DETAIL PESANAN
// ============================================================

function handleGetDetailPesanan_(body) {
  const user = validateToken_(body.token);
  if (!user) return jsonResponse_({ ok: false, error: 'Token tidak valid atau sudah expired.' });

  const idPesanan = String(body.id_pesanan || '').trim();
  if (!idPesanan) return jsonResponse_({ ok: false, error: 'id_pesanan wajib diisi.' });

  const sh = getSheet_(CONFIG.SHEETS.PESANAN);
  const last = sh.getLastRow();
  if (last < 2) return jsonResponse_({ ok: false, error: 'Tidak ada pesanan.' });

  const data = sh.getRange(2, 1, last - 1, 20).getValues();
  let header = null;
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]) === idPesanan) {
      header = {
        id_pesanan: String(data[i][0]),
        order_number: String(data[i][1]),
        customer_ref_no: String(data[i][2]),
        nama_reseller: String(data[i][3]),
        id_reseller: String(data[i][4]),
        tgl_pesan: toYMD_(data[i][5]),
        tgl_kirim: toYMD_(data[i][6]),
        tgl_kirim_aktual: toYMD_(data[i][7]),
        status: String(data[i][10]),
        total_item: Number(data[i][11]) || 0,
        total_qty: Number(data[i][12]) || 0,
        total_harga: Number(data[i][13]) || 0,
        catatan: String(data[i][14] || ''),
        memo: String(data[i][15] || '')
      };
      break;
    }
  }
  if (!header) return jsonResponse_({ ok: false, error: 'Pesanan tidak ditemukan.' });

  // Isolasi reseller
  if (user.role === 'reseller' && header.id_reseller !== String(user.id_reseller)) {
    return jsonResponse_({ ok: false, error: 'Anda tidak berhak melihat pesanan ini.' });
  }

  // Detail items
  const detailSh = getSheet_(CONFIG.SHEETS.DETAIL_PESANAN);
  const dLast = detailSh.getLastRow();
  const items = [];
  if (dLast >= 2) {
    const dData = detailSh.getRange(2, 1, dLast - 1, 14).getValues();
    for (let i = 0; i < dData.length; i++) {
      if (String(dData[i][1]) !== idPesanan) continue;
      items.push({
        id_detail:     String(dData[i][0] || ''),
        nama_produk:   String(dData[i][3] || ''),
        kode_sku:      String(dData[i][4] || ''),
        qty:           Number(dData[i][5]) || 0,
        satuan:        String(dData[i][6] || ''),
        harga_saat_pesan: Number(dData[i][7]) || 0,
        subtotal:      Number(dData[i][9]) || 0,
        catatan:       String(dData[i][10] || ''),
        status_item:   String(dData[i][11] || 'pending'),
        locked_at:     String(dData[i][12] || ''),
        locked_by:     String(dData[i][13] || '')
      });
    }
  }

  return jsonResponse_({ ok: true, data: { header: header, items: items } });
}

// ============================================================
// HELPERS
// ============================================================

function findResellerById_(idReseller) {
  const sh = getSheet_(CONFIG.SHEETS.RESELLER);
  const last = sh.getLastRow();
  if (last < 2) return null;
  const data = sh.getRange(2, 1, last - 1, 43).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][38]) === String(idReseller)) {
      return {
        id_reseller: String(data[i][38]),
        nama_reseller: String(data[i][1] || ''),
        aktif: data[i][39]
      };
    }
  }
  return null;
}

function findResellerByName_(nama) {
  const sh = getSheet_(CONFIG.SHEETS.RESELLER);
  const last = sh.getLastRow();
  if (last < 2) return null;
  const data = sh.getRange(2, 1, last - 1, 43).getValues();
  const key = String(nama).trim().toLowerCase();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][1]).trim().toLowerCase() === key) {
      return {
        id_reseller: String(data[i][38]),
        nama_reseller: String(data[i][1] || ''),
        aktif: data[i][39]
      };
    }
  }
  return null;
}

// ============================================================
// GET RESELLER LIST (untuk autocomplete form pesanan)
// Admin/owner only. Return compact [id, nama] aktif saja.
// ============================================================
function handleGetResellerList_(body) {
  const user = validateToken_(body.token);
  if (!user) {
    return jsonResponse_({ ok: false, error: 'Token tidak valid atau expired.' });
  }
  if (user.role !== 'admin' && user.role !== 'owner') {
    return jsonResponse_({ ok: false, error: 'Hanya admin/owner yang bisa lihat daftar reseller.' });
  }

  const cached = cacheGet_('reseller_list');
  if (cached && cached.length !== undefined) {
    return jsonResponse_({ ok: true, data: { items: cached, _cached: true } });
  }

  const sh = getSheet_(CONFIG.SHEETS.RESELLER);
  const last = sh.getLastRow();
  const items = [];
  if (last >= 2) {
    const data = sh.getRange(2, 1, last - 1, 43).getValues();
    for (let i = 0; i < data.length; i++) {
      const r = data[i];
      const aktif = String(r[39] || '').toUpperCase() === 'TRUE';
      if (!aktif) continue;
      const idR = String(r[38] || '').trim();
      const namaR = String(r[1] || '').trim();
      if (!idR || !namaR) continue;
      items.push([idR, namaR]);
    }
  }

  try {
    Logger.log('reseller_list: items=' + items.length + ' jsonSize=' + JSON.stringify(items).length + ' bytes');
  } catch (e) {}
  const ok = cachePut_('reseller_list', items, 3600);
  if (!ok) Logger.log('reseller_list: CACHE PUT FAILED (size over limit) - frontend will cache locally');

  return jsonResponse_({ ok: true, data: { items: items, _cached: false } });
}

function cacheBustResellerList_() {
  cacheInvalidate_(['reseller_list']);
}

// ============================================================
// GET RESELLER LIST (untuk autocomplete form pesanan)
// Admin/owner only. Return compact [id, nama] aktif saja.
// ============================================================
function handleGetResellerList_(body) {
  const user = validateToken_(body.token);
  if (!user) {
    return jsonResponse_({ ok: false, error: 'Token tidak valid atau expired.' });
  }
  if (user.role !== 'admin' && user.role !== 'owner') {
    return jsonResponse_({ ok: false, error: 'Hanya admin/owner yang bisa lihat daftar reseller.' });
  }

  const cached = cacheGet_('reseller_list');
  if (cached && cached.length !== undefined) {
    return jsonResponse_({ ok: true, data: { items: cached, _cached: true } });
  }

  const sh = getSheet_(CONFIG.SHEETS.RESELLER);
  const last = sh.getLastRow();
  const items = [];
  if (last >= 2) {
    const data = sh.getRange(2, 1, last - 1, 43).getValues();
    for (let i = 0; i < data.length; i++) {
      const r = data[i];
      const aktif = String(r[39] || '').toUpperCase() === 'TRUE';
      if (!aktif) continue;
      const idR = String(r[38] || '').trim();
      const namaR = String(r[1] || '').trim();
      if (!idR || !namaR) continue;
      items.push([idR, namaR]);
    }
  }

  try {
    Logger.log('reseller_list: items=' + items.length + ' jsonSize=' + JSON.stringify(items).length + ' bytes');
  } catch (e) {}
  const ok = cachePut_('reseller_list', items, 3600);
  if (!ok) Logger.log('reseller_list: CACHE PUT FAILED (size over limit) - frontend will cache locally');

  return jsonResponse_({ ok: true, data: { items: items, _cached: false } });
}

function cacheBustResellerList_() {
  cacheInvalidate_(['reseller_list']);
}

function findPOAktif_(idReseller, tglKirim) {
  const sh = getSheet_(CONFIG.SHEETS.PESANAN);
  const last = sh.getLastRow();
  if (last < 2) return null;
  const data = sh.getRange(2, 1, last - 1, 20).getValues();
  for (let i = 0; i < data.length; i++) {
    const st = String(data[i][10]).toLowerCase();
    if (st !== 'submitted' && st !== 'partial_locked') continue;
    if (String(data[i][4]) !== String(idReseller)) continue;
    if (toYMD_(data[i][6]) !== String(tglKirim)) continue;
    return {
      rowIndex: i + 2,
      id_pesanan: String(data[i][0]),
      order_number: String(data[i][1]),
      customer_ref_no: String(data[i][2]),
      nama_reseller: String(data[i][3]),
      tgl_kirim: String(data[i][6]),
      status: st
    };
  }
  return null;
}

function rekapDetail_(idPesanan) {
  const sh = getSheet_(CONFIG.SHEETS.DETAIL_PESANAN);
  const last = sh.getLastRow();
  let total_item = 0, total_item_locked = 0, total_item_pending = 0, total_qty = 0, total_harga = 0;
  if (last >= 2) {
    const data = sh.getRange(2, 1, last - 1, 14).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][1]) !== String(idPesanan)) continue;
      const st = String(data[i][11] || 'pending').toLowerCase();
      total_item++;
      if (st === 'pending') total_item_pending++;
      else if (st === 'locked') total_item_locked++;
      total_qty += Number(data[i][5]) || 0;
      total_harga += Number(data[i][9]) || 0;
    }
  }
  return {
    total_item: total_item,
    total_item_pending: total_item_pending,
    total_item_locked: total_item_locked,
    total_qty: total_qty,
    total_harga: total_harga
  };
}

function defaultTglKirim_() {
  const d = new Date();
  d.setDate(d.getDate() + 1);
  return Utilities.formatDate(d, CONFIG.TZ, 'yyyy-MM-dd');
}

function generateOrderNumber_(namaReseller, tglKirim) {
  // Format: YYYYMMDD(kirim)_YYYYMMDD(pesan)_HHMM_NamaReseller
  const now = new Date();
  const tglPesan = Utilities.formatDate(now, CONFIG.TZ, 'yyyyMMdd');
  const hhmm = Utilities.formatDate(now, CONFIG.TZ, 'HHmm');
  const tglKirimCompact = String(tglKirim).replace(/-/g, '');
  const namaBersih = normalizeNamaUntukOrderNumber_(namaReseller);

  let base = tglKirimCompact + '_' + tglPesan + '_' + hhmm + '_' + namaBersih;

  // Cek tabrakan
  const sh = getSheet_(CONFIG.SHEETS.PESANAN);
  const last = sh.getLastRow();
  const existing = {};
  if (last >= 2) {
    const data = sh.getRange(2, 2, last - 1, 1).getValues();
    data.forEach(function(r){ if (r[0]) existing[String(r[0])] = true; });
  }

  if (!existing[base]) return base;

  // Suffix -02, -03, ...
  for (let n = 2; n <= 99; n++) {
    const suffix = '-' + (n < 10 ? '0' + n : n);
    const candidate = base + suffix;
    if (!existing[candidate]) return candidate;
  }
  // Fallback random
  return base + '-' + Utilities.getUuid().substring(0, 4);
}

function normalizeNamaUntukOrderNumber_(nama) {
  // Hapus spasi, karakter non-alphanumeric (kecuali - dan _), max 25 char
  let s = String(nama || '').trim();
  s = s.replace(/[^A-Za-z0-9_-]/g, '');
  if (s.length > 25) s = s.substring(0, 25);
  if (!s) s = 'TanpaNama';
  return s;
}
// ============================================================
// HELPER: Normalize tanggal ke string yyyy-MM-dd
// ============================================================
// Google Sheets auto-convert string "2026-10-28" menjadi Date object.
// Saat getValues(), nilai Date dikembalikan. Kita perlu normalize.

function toYMD_(val) {
  if (val === null || val === undefined || val === '') return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, CONFIG.TZ, 'yyyy-MM-dd');
  }
  return String(val);
}
// ============================================================
// CACHED: List produk untuk getProduk
// ============================================================

function getProdukListCached_() {
  // Cache compact form (short field names) untuk muat di CacheService 95KB limit.
  // Map: i=id_produk, n=nama_produk, s=satuan, h=harga_jual,
  //      sa=stok_available, g=gambar_url, a=aktif
  const cached = cacheGet_('produk_list');
  if (cached && cached.length !== undefined) {
    return cached.map(function(p){
      return {
        id_produk: p.i, nama_produk: p.n, satuan: p.s, harga_jual: p.h,
        stok_available: p.sa, gambar_url: p.g, aktif: p.a
      };
    });
  }

  const sh = getSheet_(CONFIG.SHEETS.PRODUK);
  const last = sh.getLastRow();
  const raw = [];
  if (last >= 2) {
    const data = sh.getRange(2, 1, last - 1, 20).getValues();
    for (let i = 0; i < data.length; i++) {
      const r = data[i];
      const nama = String(r[0] || '').trim();
      if (!nama) continue;
      if (String(r[13] || '').trim() !== 'Barang Jadi') continue;

      raw.push({
        i:  String(r[14] || ''),
        n:  nama,
        s:  String(r[5] || ''),
        h:  Number(r[9]) || 0,
        sa: (Number(r[4]) || 0) - (Number(r[16]) || 0),
        g:  String(r[15] || ''),
        a:  String(r[17] || '').toUpperCase() === 'TRUE'
      });
    }
  }

  try {
    Logger.log('produk_list: items=' + raw.length + ' jsonSize=' + JSON.stringify(raw).length + ' bytes');
  } catch (e) {}
  const ok = cachePut_('produk_list', raw, CACHE.TTL_PRODUK);
  if (!ok) Logger.log('produk_list: CACHE PUT FAILED (size over limit)');

  return raw.map(function(p){
    return {
      id_produk: p.i, nama_produk: p.n, satuan: p.s, harga_jual: p.h,
      stok_available: p.sa, gambar_url: p.g, aktif: p.a
    };
  });
}