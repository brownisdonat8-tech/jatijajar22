/**
 * KK22 - Admin Pesanan
 * Endpoint khusus admin/owner:
 *   - lockItemPesanan   : lock 1 atau lebih item per PO
 *   - unlockItemPesanan : unlock 1 item (wajib alasan)
 *   - toggleProdukAktif : aktif/nonaktifkan produk (soft lock katalog)
 */

function requireAdmin_(user) {
  return user && (user.role === 'admin' || user.role === 'owner');
}

// ============================================================
// LOCK ITEM PESANAN
// ============================================================

function handleLockItemPesanan_(body) {
  const user = validateToken_(body.token);
  if (!user) return jsonResponse_({ ok: false, error: 'Token tidak valid atau sudah expired.' });
  if (!requireAdmin_(user)) return jsonResponse_({ ok: false, error: 'Hanya admin/owner yang bisa lock item.' });

  const idPesanan = String(body.id_pesanan || '').trim();
  const listIdDetail = Array.isArray(body.id_detail) ? body.id_detail.map(String) : [];
  if (!idPesanan) return jsonResponse_({ ok: false, error: 'id_pesanan wajib diisi.' });
  if (listIdDetail.length === 0) return jsonResponse_({ ok: false, error: 'id_detail (array) wajib diisi, minimal 1.' });

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    const shD = getSheet_(CONFIG.SHEETS.DETAIL_PESANAN);
    const lastD = shD.getLastRow();
    if (lastD < 2) return jsonResponse_({ ok: false, error: 'Tidak ada detail pesanan.' });

    const data = shD.getRange(2, 1, lastD - 1, 14).getValues();
    const now = nowJakarta_();
    const locked = [];
    const skipped = [];
    const targetSet = {};
    listIdDetail.forEach(function(id) { targetSet[id] = true; });
    const seen = {};

    for (let i = 0; i < data.length; i++) {
      const idDetail = String(data[i][0]);
      if (!targetSet[idDetail]) continue;
      seen[idDetail] = true;

      if (String(data[i][1]) !== idPesanan) {
        skipped.push({ id_detail: idDetail, alasan: 'tidak termasuk PO ini' });
        continue;
      }
      const st = String(data[i][11] || 'pending').toLowerCase();
      if (st === 'locked') { skipped.push({ id_detail: idDetail, alasan: 'sudah locked' }); continue; }
      if (st !== 'pending') { skipped.push({ id_detail: idDetail, alasan: 'status: ' + st }); continue; }

      shD.getRange(i + 2, 12).setValue('locked');
      shD.getRange(i + 2, 13).setValue(now);
      shD.getRange(i + 2, 14).setValue(user.id_user);
      locked.push(idDetail);
    }

    const notFound = listIdDetail.filter(function(id) { return !seen[id]; });

    const headerResult = updateHeaderStatusByRekap_(idPesanan);

    logAktivitas_(user, 'lockItemPesanan',
      'PO ' + idPesanan + ': lock ' + locked.length + ', skip ' + skipped.length);

    return jsonResponse_({
      ok: true,
      data: {
        id_pesanan: idPesanan,
        locked: locked,
        skipped: skipped,
        not_found: notFound,
        header_status: headerResult.status,
        total_item_pending: headerResult.total_item_pending,
        total_item_locked: headerResult.total_item_locked
      }
    });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// ============================================================
// UNLOCK ITEM PESANAN
// ============================================================

function handleUnlockItemPesanan_(body) {
  const user = validateToken_(body.token);
  if (!user) return jsonResponse_({ ok: false, error: 'Token tidak valid atau sudah expired.' });
  if (!requireAdmin_(user)) return jsonResponse_({ ok: false, error: 'Hanya admin/owner yang bisa unlock item.' });

  const idPesanan = String(body.id_pesanan || '').trim();
  const idDetail  = String(body.id_detail || '').trim();
  const alasan    = String(body.alasan || '').trim();

  if (!idPesanan) return jsonResponse_({ ok: false, error: 'id_pesanan wajib diisi.' });
  if (!idDetail)  return jsonResponse_({ ok: false, error: 'id_detail wajib diisi.' });
  if (!alasan)    return jsonResponse_({ ok: false, error: 'alasan unlock wajib diisi (untuk audit).' });

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    const shD = getSheet_(CONFIG.SHEETS.DETAIL_PESANAN);
    const lastD = shD.getLastRow();
    if (lastD < 2) return jsonResponse_({ ok: false, error: 'Tidak ada detail pesanan.' });

    const data = shD.getRange(2, 1, lastD - 1, 14).getValues();
    let rowIndex = -1;
    let currentStatus = '';
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]) === idDetail && String(data[i][1]) === idPesanan) {
        rowIndex = i + 2;
        currentStatus = String(data[i][11] || '').toLowerCase();
        break;
      }
    }
    if (rowIndex === -1) return jsonResponse_({ ok: false, error: 'Item tidak ditemukan atau bukan bagian dari PO ini.' });
    if (currentStatus !== 'locked') {
      return jsonResponse_({ ok: false, error: 'Item bukan berstatus locked (saat ini: ' + currentStatus + ').' });
    }

    shD.getRange(rowIndex, 12).setValue('pending');
    shD.getRange(rowIndex, 13).setValue('');
    shD.getRange(rowIndex, 14).setValue('');

    const headerResult = updateHeaderStatusByRekap_(idPesanan);

    logAktivitas_(user, 'unlockItemPesanan',
      'PO ' + idPesanan + ', item ' + idDetail + ', alasan: ' + alasan);

    return jsonResponse_({
      ok: true,
      data: {
        id_pesanan: idPesanan,
        id_detail: idDetail,
        alasan: alasan,
        header_status: headerResult.status,
        total_item_pending: headerResult.total_item_pending,
        total_item_locked: headerResult.total_item_locked
      }
    });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// ============================================================
// TOGGLE PRODUK AKTIF
// ============================================================

function handleToggleProdukAktif_(body) {
  const user = validateToken_(body.token);
  if (!user) return jsonResponse_({ ok: false, error: 'Token tidak valid atau sudah expired.' });
  if (!requireAdmin_(user)) return jsonResponse_({ ok: false, error: 'Hanya admin/owner yang bisa mengubah status produk.' });

  const idProduk = String(body.id_produk || '').trim();
  const aktifBaru = body.aktif;
  if (!idProduk) return jsonResponse_({ ok: false, error: 'id_produk wajib diisi.' });
  if (typeof aktifBaru !== 'boolean') return jsonResponse_({ ok: false, error: 'Field aktif harus boolean (true/false).' });

  const sh = getSheet_(CONFIG.SHEETS.PRODUK);
  const last = sh.getLastRow();
  if (last < 2) return jsonResponse_({ ok: false, error: 'Master produk kosong.' });

  const data = sh.getRange(2, 1, last - 1, 20).getValues();
  let foundIdx = -1;
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][14]) === idProduk) { foundIdx = i; break; }
  }
  if (foundIdx === -1) return jsonResponse_({ ok: false, error: 'Produk tidak ditemukan: ' + idProduk });

  const rowIndex = foundIdx + 2;
  const namaProduk = String(data[foundIdx][0]);
  const aktifLama = String(data[foundIdx][17] || '').toUpperCase() === 'TRUE';

  if (aktifLama === aktifBaru) {
    return jsonResponse_({
      ok: true,
      data: {
        message: 'Produk sudah dalam status ' + (aktifBaru ? 'aktif' : 'nonaktif') + '. Tidak ada perubahan.',
        id_produk: idProduk,
        nama_produk: namaProduk,
        aktif: aktifBaru,
        pending_items_affected: 0
      }
    });
  }

  // Hitung pending items terdampak (kalau nonaktifkan)
  let affectedCount = 0;
  if (!aktifBaru) {
    const shD = getSheet_(CONFIG.SHEETS.DETAIL_PESANAN);
    const lastD = shD.getLastRow();
    if (lastD >= 2) {
      const dData = shD.getRange(2, 1, lastD - 1, 14).getValues();
      const namaLower = namaProduk.toLowerCase();
      for (let i = 0; i < dData.length; i++) {
        if (String(dData[i][3]).toLowerCase() === namaLower &&
            String(dData[i][11]).toLowerCase() === 'pending') {
          affectedCount++;
        }
      }
    }
  }

  sh.getRange(rowIndex, 18).setValue(aktifBaru ? 'TRUE' : 'FALSE');
  sh.getRange(rowIndex, 20).setValue(nowJakarta_());

  logAktivitas_(user, 'toggleProdukAktif',
    namaProduk + ' -> ' + (aktifBaru ? 'aktif' : 'nonaktif'));

  return jsonResponse_({
    ok: true,
    data: {
      id_produk: idProduk,
      nama_produk: namaProduk,
      aktif_sebelum: aktifLama,
      aktif_sekarang: aktifBaru,
      pending_items_affected: affectedCount,
      catatan: aktifBaru
        ? 'Produk kembali tersedia di katalog.'
        : (affectedCount > 0
            ? affectedCount + ' item di PO pending masih pakai produk ini. Lock manual per-item kalau perlu.'
            : 'Tidak ada PO pending yang terdampak.')
    }
  });
}

// ============================================================
// HELPER: Update Header Status by Rekap
// ============================================================

function updateHeaderStatusByRekap_(idPesanan) {
  const rekap = rekapDetail_(idPesanan);
  const shP = getSheet_(CONFIG.SHEETS.PESANAN);
  const lastP = shP.getLastRow();
  const empty = { status: 'unknown', total_item: 0, total_item_pending: 0, total_item_locked: 0 };
  if (lastP < 2) return empty;

  const dataP = shP.getRange(2, 1, lastP - 1, 20).getValues();
  let rowIndex = -1;
  let currentStatus = 'submitted';
  for (let i = 0; i < dataP.length; i++) {
    if (String(dataP[i][0]) === idPesanan) {
      rowIndex = i + 2;
      currentStatus = String(dataP[i][10]).toLowerCase();
      break;
    }
  }
  if (rowIndex === -1) return empty;

  // Jangan ubah kalau sudah shipped/done/cancelled
  if (currentStatus === 'shipped' || currentStatus === 'done' || currentStatus === 'cancelled') {
    return {
      status: currentStatus,
      total_item: rekap.total_item,
      total_item_pending: rekap.total_item_pending,
      total_item_locked: rekap.total_item_locked
    };
  }

  let newStatus = 'submitted';
  if (rekap.total_item_pending === 0 && rekap.total_item_locked > 0) newStatus = 'locked';
  else if (rekap.total_item_pending > 0 && rekap.total_item_locked > 0) newStatus = 'partial_locked';
  else newStatus = 'submitted';

  shP.getRange(rowIndex, 11).setValue(newStatus);
  shP.getRange(rowIndex, 12).setValue(rekap.total_item);
  shP.getRange(rowIndex, 13).setValue(rekap.total_qty);
  shP.getRange(rowIndex, 14).setValue(rekap.total_harga);
  shP.getRange(rowIndex, 20).setValue(nowJakarta_());

  return {
    status: newStatus,
    total_item: rekap.total_item,
    total_item_pending: rekap.total_item_pending,
    total_item_locked: rekap.total_item_locked
  };
}