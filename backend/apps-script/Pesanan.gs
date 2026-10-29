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
// UPDATE PESANAN ITEMS (F7.5 ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€šÃ‚Â edit PO selama masih pending/partial_locked)
// Reseller bisa edit PO sendiri; admin/owner bisa edit semua.
// Locked items: readonly (tidak boleh diubah).
// items[] = FINAL list item PENDING saja (locked tidak dikirim, tetap ada).
// ============================================================

function handleUpdatePesananItems_(body) {
  const user = validateToken_(body.token);
  if (!user) return jsonResponse_({ ok: false, error: 'Token tidak valid atau expired.' });

  const idPesanan = String(body.id_pesanan || '').trim();
  if (!idPesanan) return jsonResponse_({ ok: false, error: 'id_pesanan wajib diisi.' });

  const incomingItems = Array.isArray(body.items) ? body.items : [];

  // --- Verify header ---
  const shP = getSheet_(CONFIG.SHEETS.PESANAN);
  const lastP = shP.getLastRow();
  if (lastP < 2) return jsonResponse_({ ok: false, error: 'Tidak ada pesanan.' });
  const dataP = shP.getRange(2, 1, lastP - 1, 20).getValues();

  let headerRowIdx = -1;
  let orderNumber = '';
  let headerResellerId = '';
  let headerStatus = '';
  for (let i = 0; i < dataP.length; i++) {
    if (String(dataP[i][0]) === idPesanan) {
      headerRowIdx = i + 2;
      orderNumber = String(dataP[i][1] || '');
      headerResellerId = String(dataP[i][4] || '');
      headerStatus = String(dataP[i][10] || '').toLowerCase();
      break;
    }
  }
  if (headerRowIdx === -1) return jsonResponse_({ ok: false, error: 'Pesanan tidak ditemukan.' });

  // --- Auth: reseller hanya PO sendiri ---
  if (user.role === 'reseller') {
    if (String(headerResellerId) !== String(user.id_reseller)) {
      return jsonResponse_({ ok: false, error: 'Anda tidak berhak mengedit pesanan ini.' });
    }
  } else if (user.role !== 'admin' && user.role !== 'owner') {
    return jsonResponse_({ ok: false, error: 'Role tidak diizinkan.' });
  }

  // --- Status guard ---
  if (headerStatus !== 'submitted' && headerStatus !== 'partial_locked') {
    return jsonResponse_({ ok: false, error: 'PO tidak bisa diedit (status: ' + headerStatus + '). Hubungi admin.' });
  }

  // --- Validate incoming items ---
  let validItems = [];
  if (incomingItems.length > 0) {
    const v = validasiItemsForUpdate_(incomingItems);
    if (!v.ok) return jsonResponse_({ ok: false, error: v.error });
    validItems = v.items;
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    // --- Read existing DETAIL rows for this PO ---
    const shD = getSheet_(CONFIG.SHEETS.DETAIL_PESANAN);
    const lastD = shD.getLastRow();
    let lockedItems = [];
    let pendingRowIndices = [];
    if (lastD >= 2) {
      const dataD = shD.getRange(2, 1, lastD - 1, 14).getValues();
      for (let i = 0; i < dataD.length; i++) {
        if (String(dataD[i][1]) !== idPesanan) continue;
        const st = String(dataD[i][11] || 'pending').toLowerCase();
        if (st === 'pending') {
          pendingRowIndices.push(i + 2);
        } else {
          lockedItems.push({
            qty: Number(dataD[i][5]) || 0,
            subtotal: Number(dataD[i][9]) || 0,
            status_item: st
          });
        }
      }
    }

    const lockedCount = lockedItems.length;

    // --- HARD DELETE existing pending rows (dari bawah ke atas) ---
    pendingRowIndices.sort(function(a, b) { return b - a; });
    pendingRowIndices.forEach(function(rowIdx) { shD.deleteRow(rowIdx); });

    // --- Insert new pending rows ---
    let maxDetailNum = 0;
    const lastDNow = shD.getLastRow();
    if (lastDNow >= 2) {
      const dataDNow = shD.getRange(2, 1, lastDNow - 1, 1).getValues();
      dataDNow.forEach(function(row) {
        const m = String(row[0] || '').match(/KK22-DTL-(\d+)/);
        if (m) {
          const n = parseInt(m[1], 10);
          if (n > maxDetailNum) maxDetailNum = n;
        }
      });
    }

    const newRows = [];
    validItems.forEach(function(it) {
      maxDetailNum++;
      const idDetail = 'KK22-DTL-' + pad4_(maxDetailNum);
      const qty = Number(it.qty) || 0;
      const harga = Number(it.harga_jual) || 0;
      const subtotal = qty * harga;
      newRows.push([
        idDetail, idPesanan, orderNumber,
        it.nama_produk, it.kode_sku || '', qty, it.satuan || '',
        harga, '', subtotal, '', 'pending', '', ''
      ]);
    });

    if (newRows.length > 0) {
      const startRow = shD.getLastRow() + 1;
      shD.getRange(startRow, 1, newRows.length, 14).setValues(newRows);
    }

    // --- Recalc totals ---
    let totalQty = 0, totalHarga = 0;
    lockedItems.forEach(function(it) {
      totalQty += it.qty;
      totalHarga += it.subtotal;
    });
    newRows.forEach(function(r) {
      totalQty += Number(r[5]) || 0;
      totalHarga += Number(r[9]) || 0;
    });
    const totalItemCount = lockedCount + newRows.length;

    // --- Determine new header status ---
    let newStatus = headerStatus;
    if (totalItemCount === 0) newStatus = 'cancelled';
    else if (lockedCount > 0 && newRows.length > 0) newStatus = 'partial_locked';
    else if (lockedCount > 0 && newRows.length === 0) newStatus = 'locked';
    else newStatus = 'submitted';

    const now = nowJakarta_();
    shP.getRange(headerRowIdx, 11).setValue(newStatus);
    shP.getRange(headerRowIdx, 12).setValue(totalItemCount);
    shP.getRange(headerRowIdx, 13).setValue(totalQty);
    shP.getRange(headerRowIdx, 14).setValue(totalHarga);
    shP.getRange(headerRowIdx, 19).setValue(user.id_user);
    shP.getRange(headerRowIdx, 20).setValue(now);

    try {
      logAktivitas_(user, 'updatePesananItems',
        'PO ' + idPesanan + ': ' + newRows.length + ' pending baru, ' +
        lockedCount + ' locked readonly, status=' + newStatus);
    } catch (e) {}

    return jsonResponse_({
      ok: true,
      data: {
        id_pesanan: idPesanan,
        updated_items: newRows.length,
        locked_items: lockedCount,
        total_item: totalItemCount,
        total_qty: totalQty,
        total_harga: totalHarga,
        header_status: newStatus,
        cancelled: newStatus === 'cancelled'
      }
    });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function validasiItemsForUpdate_(rawItems) {
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
      kode_sku: String(r[2] || ''),
      satuan: String(r[5] || ''),
      harga_jual: Number(r[9]) || 0,
      aktif: String(r[17] || '').toUpperCase() === 'TRUE'
    };
  });

  const items = [];
  const seen = {};
  for (let i = 0; i < rawItems.length; i++) {
    const it = rawItems[i];
    const nama = String(it.nama_produk || '').trim();
    const qty = Number(it.qty) || 0;
    if (!nama) return { ok: false, error: 'Item ke-' + (i + 1) + ': nama_produk kosong.' };
    const key = nama.toLowerCase();
    if (seen[key]) return { ok: false, error: 'Produk duplikat: ' + nama };
    seen[key] = true;

    const produk = produkMap[key];
    if (!produk) return { ok: false, error: 'Produk tidak ditemukan: ' + nama };
    if (!produk.aktif) return { ok: false, error: 'Produk nonaktif: ' + nama };
    if (qty <= 0) return { ok: false, error: 'Qty harus > 0: ' + nama };

    items.push({
      nama_produk: nama,
      kode_sku: produk.kode_sku,
      satuan: produk.satuan,
      harga_jual: produk.harga_jual,
      qty: qty
    });
  }
  return { ok: true, items: items };
}

// ============================================================
// F9.1 ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â HARGA KHUSUS PER RESELLER
// ============================================================

function getHargaOverrideForReseller_(idReseller) {
  if (!idReseller) return {};
  const cacheKey = 'harga_override_' + idReseller;
  const cached = cacheGet_(cacheKey);
  if (cached && typeof cached === 'object') return cached;

  const sh = getSheet_(CONFIG.SHEETS.RESELLER_HARGA);
  if (!sh) return {};
  const last = sh.getLastRow();
  if (last < 2) {
    cachePut_(cacheKey, {}, 300);
    return {};
  }

  const data = sh.getRange(2, 1, last - 1, 7).getValues();
  const map = {};
  for (let i = 0; i < data.length; i++) {
    const r = data[i];
    if (String(r[0]) !== String(idReseller)) continue;
    const nama = String(r[1] || '').trim();
    const harga = Number(r[2]) || 0;
    if (!nama || harga <= 0) continue;
    map[nama.toLowerCase()] = {
      harga_khusus: harga,
      catatan: String(r[3] || '')
    };
  }
  cachePut_(cacheKey, map, 300);
  return map;
}

function handleGetHargaKhusus_(body) {
  const user = validateToken_(body.token);
  if (!user) return jsonResponse_({ ok: false, error: 'Token tidak valid atau expired.' });
  if (!requireAdmin_(user)) return jsonResponse_({ ok: false, error: 'Hanya admin/owner.' });

  const idReseller = String(body.id_reseller || '').trim();
  if (!idReseller) return jsonResponse_({ ok: false, error: 'id_reseller wajib diisi.' });

  const sh = getSheet_(CONFIG.SHEETS.RESELLER_HARGA);
  if (!sh) return jsonResponse_({ ok: false, error: 'Sheet RESELLER_HARGA belum ada.' });
  const last = sh.getLastRow();
  const items = [];
  if (last >= 2) {
    const data = sh.getRange(2, 1, last - 1, 7).getValues();
    for (let i = 0; i < data.length; i++) {
      const r = data[i];
      if (String(r[0]) !== idReseller) continue;
      items.push({
        id_reseller: String(r[0]),
        nama_produk: String(r[1]),
        harga_khusus: Number(r[2]) || 0,
        catatan: String(r[3] || ''),
        tgl_dibuat: String(r[4] || ''),
        diupdate_oleh: String(r[5] || ''),
        tgl_update_terakhir: String(r[6] || '')
      });
    }
  }
  items.sort(function(a, b) { return a.nama_produk.localeCompare(b.nama_produk); });
  return jsonResponse_({ ok: true, data: { total: items.length, items: items } });
}

function handleSetHargaKhusus_(body) {
  const user = validateToken_(body.token);
  if (!user) return jsonResponse_({ ok: false, error: 'Token tidak valid atau expired.' });
  if (!requireAdmin_(user)) return jsonResponse_({ ok: false, error: 'Hanya admin/owner.' });

  const idReseller = String(body.id_reseller || '').trim();
  const namaProduk = String(body.nama_produk || '').trim();
  const hargaKhusus = Number(body.harga_khusus);
  const catatan = String(body.catatan || '').trim();

  if (!idReseller) return jsonResponse_({ ok: false, error: 'id_reseller wajib diisi.' });
  if (!namaProduk) return jsonResponse_({ ok: false, error: 'nama_produk wajib diisi.' });
  if (!hargaKhusus || hargaKhusus <= 0) return jsonResponse_({ ok: false, error: 'harga_khusus harus angka > 0.' });

  const res = findResellerById_(idReseller);
  if (!res) return jsonResponse_({ ok: false, error: 'Reseller tidak ditemukan: ' + idReseller });

  const shP = getSheet_(CONFIG.SHEETS.PRODUK);
  const lastP = shP.getLastRow();
  let produkFound = null;
  if (lastP >= 2) {
    const dataP = shP.getRange(2, 1, lastP - 1, 20).getValues();
    for (let i = 0; i < dataP.length; i++) {
      if (String(dataP[i][0]).toLowerCase() === namaProduk.toLowerCase()) {
        produkFound = {
          nama: String(dataP[i][0]),
          harga_default: Number(dataP[i][9]) || 0
        };
        break;
      }
    }
  }
  if (!produkFound) return jsonResponse_({ ok: false, error: 'Produk tidak ditemukan: ' + namaProduk });

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    const sh = getSheet_(CONFIG.SHEETS.RESELLER_HARGA);
    if (!sh) return jsonResponse_({ ok: false, error: 'Sheet RESELLER_HARGA belum ada.' });
    const last = sh.getLastRow();
    let foundRowIdx = -1;
    if (last >= 2) {
      const data = sh.getRange(2, 1, last - 1, 7).getValues();
      for (let i = 0; i < data.length; i++) {
        if (String(data[i][0]) === idReseller &&
            String(data[i][1]).toLowerCase() === namaProduk.toLowerCase()) {
          foundRowIdx = i + 2;
          break;
        }
      }
    }

    const now = nowJakarta_();
    if (foundRowIdx > 0) {
      sh.getRange(foundRowIdx, 3).setValue(hargaKhusus);
      sh.getRange(foundRowIdx, 4).setValue(catatan);
      sh.getRange(foundRowIdx, 6).setValue(user.id_user);
      sh.getRange(foundRowIdx, 7).setValue(now);
    } else {
      const newRow = sh.getLastRow() + 1;
      sh.getRange(newRow, 1, 1, 7).setValues([[
        idReseller, namaProduk, hargaKhusus, catatan, now, user.id_user, now
      ]]);
    }

    cacheInvalidate_(['harga_override_' + idReseller]);

    try {
      logAktivitas_(user, 'setHargaKhusus',
        'Reseller ' + idReseller + ', produk ' + namaProduk +
        ', harga=' + hargaKhusus + ' (default=' + produkFound.harga_default + ')');
    } catch (e) {}

    return jsonResponse_({
      ok: true,
      data: {
        mode: foundRowIdx > 0 ? 'updated' : 'created',
        id_reseller: idReseller,
        nama_produk: produkFound.nama,
        harga_khusus: hargaKhusus,
        harga_default: produkFound.harga_default
      }
    });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function handleDeleteHargaKhusus_(body) {
  const user = validateToken_(body.token);
  if (!user) return jsonResponse_({ ok: false, error: 'Token tidak valid atau expired.' });
  if (!requireAdmin_(user)) return jsonResponse_({ ok: false, error: 'Hanya admin/owner.' });

  const idReseller = String(body.id_reseller || '').trim();
  const namaProduk = String(body.nama_produk || '').trim();
  if (!idReseller) return jsonResponse_({ ok: false, error: 'id_reseller wajib diisi.' });
  if (!namaProduk) return jsonResponse_({ ok: false, error: 'nama_produk wajib diisi.' });

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    const sh = getSheet_(CONFIG.SHEETS.RESELLER_HARGA);
    if (!sh) return jsonResponse_({ ok: false, error: 'Sheet RESELLER_HARGA belum ada.' });
    const last = sh.getLastRow();
    if (last < 2) return jsonResponse_({ ok: false, error: 'Belum ada data harga khusus.' });

    const data = sh.getRange(2, 1, last - 1, 7).getValues();
    let foundRowIdx = -1;
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]) === idReseller &&
          String(data[i][1]).toLowerCase() === namaProduk.toLowerCase()) {
        foundRowIdx = i + 2;
        break;
      }
    }
    if (foundRowIdx === -1) return jsonResponse_({ ok: false, error: 'Harga khusus tidak ditemukan.' });

    sh.deleteRow(foundRowIdx);
    cacheInvalidate_(['harga_override_' + idReseller]);

    try {
      logAktivitas_(user, 'deleteHargaKhusus',
        'Reseller ' + idReseller + ', produk ' + namaProduk);
    } catch (e) {}

    return jsonResponse_({ ok: true, data: { id_reseller: idReseller, nama_produk: namaProduk } });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// ============================================================
// F9.1b ÃƒÂ¢Ã¢â€šÂ¬Ã¢â‚¬Â PRICE LISTS (PAKET HARGA)
// ============================================================

function getNextPriceListId_() {
  const sh = getSheet_(CONFIG.SHEETS.PRICE_LISTS);
  const last = sh.getLastRow();
  if (last < 2) return 'KK22-PL-0001';
  const data = sh.getRange(2, 1, last - 1, 1).getValues();
  let maxNum = 0;
  data.forEach(function(r) {
    const m = String(r[0] || '').match(/-PL-(\d+)/);
    if (m) { const n = parseInt(m[1], 10); if (n > maxNum) maxNum = n; }
  });
  return 'KK22-PL-' + pad4_(maxNum + 1);
}

function getMaxPriceListItemNum_() {
  const sh = getSheet_(CONFIG.SHEETS.PRICE_LIST_ITEMS);
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const data = sh.getRange(2, 1, last - 1, 1).getValues();
  let maxNum = 0;
  data.forEach(function(r) {
    const m = String(r[0] || '').match(/-PLI-(\d+)/);
    if (m) { const n = parseInt(m[1], 10); if (n > maxNum) maxNum = n; }
  });
  return maxNum;
}

function getPriceListById_(idPriceList) {
  const sh = getSheet_(CONFIG.SHEETS.PRICE_LISTS);
  const last = sh.getLastRow();
  if (last < 2) return null;
  const data = sh.getRange(2, 1, last - 1, 7).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(idPriceList)) {
      return {
        id_price_list: String(data[i][0]),
        nama_paket: String(data[i][1] || ''),
        deskripsi: String(data[i][2] || ''),
        aktif: String(data[i][3] || '').toUpperCase() === 'TRUE',
        tgl_dibuat: String(data[i][4] || ''),
        diupdate_oleh: String(data[i][5] || ''),
        tgl_update_terakhir: String(data[i][6] || ''),
        rowIndex: i + 2
      };
    }
  }
  return null;
}

function getPriceListOverrideForReseller_(idPriceList) {
  if (!idPriceList) return {};
  const cacheKey = 'price_list_override_' + idPriceList;
  const cached = cacheGet_(cacheKey);
  if (cached && typeof cached === 'object') return cached;

  const sh = getSheet_(CONFIG.SHEETS.PRICE_LIST_ITEMS);
  if (!sh) return {};
  const last = sh.getLastRow();
  if (last < 2) {
    cachePut_(cacheKey, {}, 300);
    return {};
  }
  const data = sh.getRange(2, 1, last - 1, 6).getValues();
  const map = {};
  for (let i = 0; i < data.length; i++) {
    const r = data[i];
    if (String(r[1]) !== String(idPriceList)) continue;
    const nama = String(r[2] || '').trim();
    const harga = Number(r[3]) || 0;
    if (!nama || harga <= 0) continue;
    map[nama.toLowerCase()] = { harga: harga };
  }
  cachePut_(cacheKey, map, 300);
  return map;
}

function getDefaultHargaMap_() {
  const sh = getSheet_(CONFIG.SHEETS.PRODUK);
  const last = sh.getLastRow();
  if (last < 2) return {};
  const data = sh.getRange(2, 1, last - 1, 20).getValues();
  const map = {};
  for (let i = 0; i < data.length; i++) {
    const r = data[i];
    const nama = String(r[0] || '').trim();
    if (!nama) continue;
    map[nama.toLowerCase()] = {
      nama_produk: nama,
      harga: Number(r[9]) || 0,
      aktif: String(r[17] || '').toUpperCase() === 'TRUE'
    };
  }
  return map;
}

function handleGetPriceLists_(body) {
  const user = validateToken_(body.token);
  if (!user) return jsonResponse_({ ok: false, error: 'Token tidak valid atau expired.' });
  if (!requireAdmin_(user)) return jsonResponse_({ ok: false, error: 'Hanya admin/owner.' });

  const sh = getSheet_(CONFIG.SHEETS.PRICE_LISTS);
  const last = sh.getLastRow();
  const items = [];
  if (last >= 2) {
    const data = sh.getRange(2, 1, last - 1, 7).getValues();
    const shI = getSheet_(CONFIG.SHEETS.PRICE_LIST_ITEMS);
    const lastI = shI.getLastRow();
    const countMap = {};
    if (lastI >= 2) {
      const dataI = shI.getRange(2, 1, lastI - 1, 2).getValues();
      dataI.forEach(function(r) {
        const id = String(r[1]);
        countMap[id] = (countMap[id] || 0) + 1;
      });
    }
    for (let i = 0; i < data.length; i++) {
      const r = data[i];
      items.push({
        id_price_list: String(r[0]),
        nama_paket: String(r[1] || ''),
        deskripsi: String(r[2] || ''),
        aktif: String(r[3] || '').toUpperCase() === 'TRUE',
        total_item: countMap[String(r[0])] || 0,
        tgl_dibuat: String(r[4] || ''),
        diupdate_oleh: String(r[5] || ''),
        tgl_update_terakhir: String(r[6] || '')
      });
    }
  }
  items.sort(function(a, b) { return a.nama_paket.localeCompare(b.nama_paket); });
  return jsonResponse_({ ok: true, data: { total: items.length, items: items } });
}

function handleGetPriceListDetail_(body) {
  const user = validateToken_(body.token);
  if (!user) return jsonResponse_({ ok: false, error: 'Token tidak valid atau expired.' });
  if (!requireAdmin_(user)) return jsonResponse_({ ok: false, error: 'Hanya admin/owner.' });

  const idPriceList = String(body.id_price_list || '').trim();
  if (!idPriceList) return jsonResponse_({ ok: false, error: 'id_price_list wajib diisi.' });

  const header = getPriceListById_(idPriceList);
  if (!header) return jsonResponse_({ ok: false, error: 'Paket tidak ditemukan.' });

  const produkMap = getDefaultHargaMap_();

  const shI = getSheet_(CONFIG.SHEETS.PRICE_LIST_ITEMS);
  const lastI = shI.getLastRow();
  const itemsMap = {};
  if (lastI >= 2) {
    const dataI = shI.getRange(2, 1, lastI - 1, 6).getValues();
    dataI.forEach(function(r) {
      if (String(r[1]) !== idPriceList) return;
      const nama = String(r[2] || '').trim();
      if (!nama) return;
      itemsMap[nama.toLowerCase()] = Number(r[3]) || 0;
    });
  }

  const items = [];
  Object.keys(produkMap).forEach(function(key) {
    const p = produkMap[key];
    if (!p.aktif) return;
    items.push({
      nama_produk: p.nama_produk,
      harga_default: p.harga,
      harga_paket: itemsMap[key] || 0
    });
  });
  items.sort(function(a, b) { return a.nama_produk.localeCompare(b.nama_produk); });

  return jsonResponse_({ ok: true, data: { header: header, total: items.length, items: items } });
}

function handleCreatePriceList_(body) {
  const user = validateToken_(body.token);
  if (!user) return jsonResponse_({ ok: false, error: 'Token tidak valid atau expired.' });
  if (!requireAdmin_(user)) return jsonResponse_({ ok: false, error: 'Hanya admin/owner.' });

  const namaPaket = String(body.nama_paket || '').trim();
  const deskripsi = String(body.deskripsi || '').trim();
  const sourceType = String(body.source_type || 'normal').toLowerCase();

  if (!namaPaket) return jsonResponse_({ ok: false, error: 'nama_paket wajib diisi.' });
  const validTypes = ['normal', 'price_list', 'reseller', 'excel', 'empty'];
  if (validTypes.indexOf(sourceType) === -1) {
    return jsonResponse_({ ok: false, error: 'source_type tidak valid. Pilih: ' + validTypes.join(', ') });
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    const produkMap = getDefaultHargaMap_();
    const newPriceListId = getNextPriceListId_();
    const now = nowJakarta_();
    const itemsMap = {};

    if (sourceType === 'normal') {
      Object.keys(produkMap).forEach(function(key) {
        const p = produkMap[key];
        if (p.aktif) itemsMap[key] = p.harga;
      });
    } else if (sourceType === 'price_list') {
      const srcId = String(body.source_id || '').trim();
      if (!srcId) return jsonResponse_({ ok: false, error: 'source_id (id_price_list) wajib untuk source price_list.' });
      const src = getPriceListById_(srcId);
      if (!src) return jsonResponse_({ ok: false, error: 'Paket sumber tidak ditemukan.' });
      const shSrc = getSheet_(CONFIG.SHEETS.PRICE_LIST_ITEMS);
      const lastSrc = shSrc.getLastRow();
      if (lastSrc >= 2) {
        const dataSrc = shSrc.getRange(2, 1, lastSrc - 1, 6).getValues();
        dataSrc.forEach(function(r) {
          if (String(r[1]) === srcId) {
            itemsMap[String(r[2]).toLowerCase()] = Number(r[3]) || 0;
          }
        });
      }
    } else if (sourceType === 'reseller') {
      const srcReseller = String(body.source_id || '').trim();
      if (!srcReseller) return jsonResponse_({ ok: false, error: 'source_id (id_reseller) wajib untuk source reseller.' });
      const rInfo = findResellerById_(srcReseller);
      if (!rInfo) return jsonResponse_({ ok: false, error: 'Reseller sumber tidak ditemukan.' });

      Object.keys(produkMap).forEach(function(key) {
        const p = produkMap[key];
        if (p.aktif) itemsMap[key] = p.harga;
      });

      if (rInfo.id_price_list) {
        const shPL = getSheet_(CONFIG.SHEETS.PRICE_LIST_ITEMS);
        const lastPL = shPL.getLastRow();
        if (lastPL >= 2) {
          const dataPL = shPL.getRange(2, 1, lastPL - 1, 6).getValues();
          dataPL.forEach(function(r) {
            if (String(r[1]) === rInfo.id_price_list) {
              itemsMap[String(r[2]).toLowerCase()] = Number(r[3]) || 0;
            }
          });
        }
      }

      const customMap = getHargaOverrideForReseller_(srcReseller);
      Object.keys(customMap).forEach(function(key) {
        itemsMap[key] = customMap[key].harga_khusus;
      });
    } else if (sourceType === 'excel') {
      const rawItems = Array.isArray(body.items) ? body.items : [];
      if (rawItems.length === 0) return jsonResponse_({ ok: false, error: 'items[] wajib untuk source excel.' });
      rawItems.forEach(function(it) {
        const nama = String(it.nama_produk || '').trim();
        const harga = Number(it.harga);
        if (!nama || !harga || harga <= 0) return;
        const key = nama.toLowerCase();
        if (produkMap[key]) {
          itemsMap[key] = harga;
        }
      });
    } else if (sourceType === 'empty') {
      // skip
    }

    const shPL = getSheet_(CONFIG.SHEETS.PRICE_LISTS);
    const newRow = shPL.getLastRow() + 1;
    shPL.getRange(newRow, 1, 1, 7).setValues([[
      newPriceListId, namaPaket, deskripsi, 'TRUE', now, user.id_user, now
    ]]);

    const itemKeys = Object.keys(itemsMap);
    if (itemKeys.length > 0) {
      const shItem = getSheet_(CONFIG.SHEETS.PRICE_LIST_ITEMS);
      const maxNum = getMaxPriceListItemNum_();
      const rows = [];
      let counter = maxNum;
      itemKeys.forEach(function(key) {
        counter++;
        rows.push([
          'KK22-PLI-' + pad4_(counter),
          newPriceListId,
          produkMap[key].nama_produk,
          itemsMap[key],
          now,
          now
        ]);
      });
      const startRow = shItem.getLastRow() + 1;
      shItem.getRange(startRow, 1, rows.length, 6).setValues(rows);
    }

    cacheInvalidate_(['price_list_override_' + newPriceListId]);

    try {
      logAktivitas_(user, 'createPriceList',
        'Paket ' + newPriceListId + ' (' + namaPaket + ') source=' + sourceType + ', ' + itemKeys.length + ' item');
    } catch (e) {}

    return jsonResponse_({
      ok: true,
      data: {
        id_price_list: newPriceListId,
        nama_paket: namaPaket,
        deskripsi: deskripsi,
        source_type: sourceType,
        total_item: itemKeys.length
      }
    });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function handleRenamePriceList_(body) {
  const user = validateToken_(body.token);
  if (!user) return jsonResponse_({ ok: false, error: 'Token tidak valid atau expired.' });
  if (!requireAdmin_(user)) return jsonResponse_({ ok: false, error: 'Hanya admin/owner.' });

  const idPriceList = String(body.id_price_list || '').trim();
  const namaBaru = String(body.nama_baru || '').trim();
  const deskripsi = String(body.deskripsi || '').trim();
  if (!idPriceList) return jsonResponse_({ ok: false, error: 'id_price_list wajib diisi.' });
  if (!namaBaru) return jsonResponse_({ ok: false, error: 'nama_baru wajib diisi.' });

  const header = getPriceListById_(idPriceList);
  if (!header) return jsonResponse_({ ok: false, error: 'Paket tidak ditemukan.' });

  const sh = getSheet_(CONFIG.SHEETS.PRICE_LISTS);
  const now = nowJakarta_();
  sh.getRange(header.rowIndex, 2).setValue(namaBaru);
  if (deskripsi !== null && deskripsi !== undefined && body.deskripsi !== undefined) {
    sh.getRange(header.rowIndex, 3).setValue(deskripsi);
  }
  sh.getRange(header.rowIndex, 6).setValue(user.id_user);
  sh.getRange(header.rowIndex, 7).setValue(now);

  try {
    logAktivitas_(user, 'renamePriceList',
      'Paket ' + idPriceList + ': "' + header.nama_paket + '" -> "' + namaBaru + '"');
  } catch (e) {}

  return jsonResponse_({ ok: true, data: { id_price_list: idPriceList, nama_paket: namaBaru } });
}

function handleTogglePriceListAktif_(body) {
  const user = validateToken_(body.token);
  if (!user) return jsonResponse_({ ok: false, error: 'Token tidak valid atau expired.' });
  if (!requireAdmin_(user)) return jsonResponse_({ ok: false, error: 'Hanya admin/owner.' });

  const idPriceList = String(body.id_price_list || '').trim();
  const aktifBaru = body.aktif;
  if (!idPriceList) return jsonResponse_({ ok: false, error: 'id_price_list wajib diisi.' });
  if (typeof aktifBaru !== 'boolean') return jsonResponse_({ ok: false, error: 'Field aktif harus boolean.' });

  const header = getPriceListById_(idPriceList);
  if (!header) return jsonResponse_({ ok: false, error: 'Paket tidak ditemukan.' });

  const sh = getSheet_(CONFIG.SHEETS.PRICE_LISTS);
  const now = nowJakarta_();
  sh.getRange(header.rowIndex, 4).setValue(aktifBaru ? 'TRUE' : 'FALSE');
  sh.getRange(header.rowIndex, 6).setValue(user.id_user);
  sh.getRange(header.rowIndex, 7).setValue(now);

  try {
    logAktivitas_(user, 'togglePriceListAktif', 'Paket ' + idPriceList + ' aktif=' + aktifBaru);
  } catch (e) {}

  return jsonResponse_({ ok: true, data: { id_price_list: idPriceList, aktif: aktifBaru } });
}

function handleDeletePriceList_(body) {
  const user = validateToken_(body.token);
  if (!user) return jsonResponse_({ ok: false, error: 'Token tidak valid atau expired.' });
  if (!requireAdmin_(user)) return jsonResponse_({ ok: false, error: 'Hanya admin/owner.' });

  const idPriceList = String(body.id_price_list || '').trim();
  if (!idPriceList) return jsonResponse_({ ok: false, error: 'id_price_list wajib diisi.' });

  const header = getPriceListById_(idPriceList);
  if (!header) return jsonResponse_({ ok: false, error: 'Paket tidak ditemukan.' });

  const shR = getSheet_(CONFIG.SHEETS.RESELLER);
  const lastR = shR.getLastRow();
  let usedCount = 0;
  if (lastR >= 2) {
    const dataR = shR.getRange(2, 1, lastR - 1, 44).getValues();
    for (let i = 0; i < dataR.length; i++) {
      if (String(dataR[i][43]) === idPriceList) usedCount++;
    }
  }
  if (usedCount > 0) {
    return jsonResponse_({ ok: false, error: 'Paket masih dipakai oleh ' + usedCount + ' reseller. Lepas dulu sebelum hapus.' });
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    const shI = getSheet_(CONFIG.SHEETS.PRICE_LIST_ITEMS);
    const lastI = shI.getLastRow();
    if (lastI >= 2) {
      const dataI = shI.getRange(2, 1, lastI - 1, 6).getValues();
      const toDelete = [];
      for (let i = 0; i < dataI.length; i++) {
        if (String(dataI[i][1]) === idPriceList) toDelete.push(i + 2);
      }
      toDelete.sort(function(a, b) { return b - a; });
      toDelete.forEach(function(rowIdx) { shI.deleteRow(rowIdx); });
    }

    const shPL = getSheet_(CONFIG.SHEETS.PRICE_LISTS);
    shPL.deleteRow(header.rowIndex);

    cacheInvalidate_(['price_list_override_' + idPriceList]);

    try {
      logAktivitas_(user, 'deletePriceList',
        'Paket ' + idPriceList + ' (' + header.nama_paket + ') dihapus');
    } catch (e) {}

    return jsonResponse_({ ok: true, data: { id_price_list: idPriceList, nama_paket: header.nama_paket } });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function handleSetPriceListItem_(body) {
  const user = validateToken_(body.token);
  if (!user) return jsonResponse_({ ok: false, error: 'Token tidak valid atau expired.' });
  if (!requireAdmin_(user)) return jsonResponse_({ ok: false, error: 'Hanya admin/owner.' });

  const idPriceList = String(body.id_price_list || '').trim();
  const namaProduk = String(body.nama_produk || '').trim();
  const harga = Number(body.harga);
  if (!idPriceList) return jsonResponse_({ ok: false, error: 'id_price_list wajib diisi.' });
  if (!namaProduk) return jsonResponse_({ ok: false, error: 'nama_produk wajib diisi.' });
  if (!harga || harga <= 0) return jsonResponse_({ ok: false, error: 'harga harus angka > 0.' });

  const header = getPriceListById_(idPriceList);
  if (!header) return jsonResponse_({ ok: false, error: 'Paket tidak ditemukan.' });

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    const shI = getSheet_(CONFIG.SHEETS.PRICE_LIST_ITEMS);
    const lastI = shI.getLastRow();
    let foundRow = -1;
    let foundExistingHarga = 0;
    if (lastI >= 2) {
      const dataI = shI.getRange(2, 1, lastI - 1, 6).getValues();
      for (let i = 0; i < dataI.length; i++) {
        if (String(dataI[i][1]) === idPriceList &&
            String(dataI[i][2]).toLowerCase() === namaProduk.toLowerCase()) {
          foundRow = i + 2;
          foundExistingHarga = Number(dataI[i][3]) || 0;
          break;
        }
      }
    }

    const now = nowJakarta_();
    if (foundRow > 0) {
      shI.getRange(foundRow, 4).setValue(harga);
      shI.getRange(foundRow, 6).setValue(now);
    } else {
      const maxNum = getMaxPriceListItemNum_();
      const newRow = shI.getLastRow() + 1;
      shI.getRange(newRow, 1, 1, 6).setValues([[
        'KK22-PLI-' + pad4_(maxNum + 1),
        idPriceList, namaProduk, harga, now, now
      ]]);
    }

    cacheInvalidate_(['price_list_override_' + idPriceList]);

    try {
      logAktivitas_(user, 'setPriceListItem',
        'Paket ' + idPriceList + ', produk ' + namaProduk + ', harga=' + harga +
        (foundRow > 0 ? ' (sebelumnya ' + foundExistingHarga + ')' : ' (baru)'));
    } catch (e) {}

    return jsonResponse_({
      ok: true,
      data: {
        mode: foundRow > 0 ? 'updated' : 'created',
        id_price_list: idPriceList,
        nama_produk: namaProduk,
        harga: harga
      }
    });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function handleBulkSetPriceListItems_(body) {
  const user = validateToken_(body.token);
  if (!user) return jsonResponse_({ ok: false, error: 'Token tidak valid atau expired.' });
  if (!requireAdmin_(user)) return jsonResponse_({ ok: false, error: 'Hanya admin/owner.' });

  const idPriceList = String(body.id_price_list || '').trim();
  if (!idPriceList) return jsonResponse_({ ok: false, error: 'id_price_list wajib diisi.' });
  const rawItems = Array.isArray(body.items) ? body.items : [];
  if (rawItems.length === 0) return jsonResponse_({ ok: false, error: 'items[] wajib.' });

  const header = getPriceListById_(idPriceList);
  if (!header) return jsonResponse_({ ok: false, error: 'Paket tidak ditemukan.' });

  const produkMap = getDefaultHargaMap_();

  const validItems = [];
  rawItems.forEach(function(it) {
    const nama = String(it.nama_produk || '').trim();
    const harga = Number(it.harga);
    if (!nama || !harga || harga <= 0) return;
    const key = nama.toLowerCase();
    if (!produkMap[key]) return;
    validItems.push({ key: key, nama: produkMap[key].nama_produk, harga: harga });
  });

  if (validItems.length === 0) {
    return jsonResponse_({ ok: false, error: 'Tidak ada item valid untuk disimpan.' });
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    const shI = getSheet_(CONFIG.SHEETS.PRICE_LIST_ITEMS);
    const lastI = shI.getLastRow();
    const existingMap = {};
    if (lastI >= 2) {
      const dataI = shI.getRange(2, 1, lastI - 1, 6).getValues();
      for (let i = 0; i < dataI.length; i++) {
        if (String(dataI[i][1]) !== idPriceList) continue;
        existingMap[String(dataI[i][2]).toLowerCase()] = i + 2;
      }
    }

    const now = nowJakarta_();
    let updated = 0;
    let created = 0;
    let nextNum = getMaxPriceListItemNum_();
    const newRows = [];

    validItems.forEach(function(it) {
      const existingRow = existingMap[it.key];
      if (existingRow > 0) {
        shI.getRange(existingRow, 4).setValue(it.harga);
        shI.getRange(existingRow, 6).setValue(now);
        updated++;
      } else {
        nextNum++;
        newRows.push([
          'KK22-PLI-' + pad4_(nextNum),
          idPriceList, it.nama, it.harga, now, now
        ]);
        created++;
      }
    });

    if (newRows.length > 0) {
      const startRow = shI.getLastRow() + 1;
      shI.getRange(startRow, 1, newRows.length, 6).setValues(newRows);
    }

    cacheInvalidate_(['price_list_override_' + idPriceList]);

    try {
      logAktivitas_(user, 'bulkSetPriceListItems',
        'Paket ' + idPriceList + ': updated=' + updated + ', created=' + created);
    } catch (e) {}

    return jsonResponse_({
      ok: true,
      data: { id_price_list: idPriceList, updated: updated, created: created, total: validItems.length }
    });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function handleAssignPriceListToReseller_(body) {
  const user = validateToken_(body.token);
  if (!user) return jsonResponse_({ ok: false, error: 'Token tidak valid atau expired.' });
  if (!requireAdmin_(user)) return jsonResponse_({ ok: false, error: 'Hanya admin/owner.' });

  const idReseller = String(body.id_reseller || '').trim();
  const idPriceList = String(body.id_price_list || '').trim();
  if (!idReseller) return jsonResponse_({ ok: false, error: 'id_reseller wajib diisi.' });

  const shR = getSheet_(CONFIG.SHEETS.RESELLER);
  const lastR = shR.getLastRow();
  if (lastR < 2) return jsonResponse_({ ok: false, error: 'Tidak ada reseller.' });

  const dataR = shR.getRange(2, 1, lastR - 1, 44).getValues();
  let foundRowIdx = -1;
  let currentIdPL = '';
  for (let i = 0; i < dataR.length; i++) {
    if (String(dataR[i][38]) === idReseller) {
      foundRowIdx = i + 2;
      currentIdPL = String(dataR[i][43] || '');
      break;
    }
  }
  if (foundRowIdx === -1) return jsonResponse_({ ok: false, error: 'Reseller tidak ditemukan.' });

  if (idPriceList) {
    const header = getPriceListById_(idPriceList);
    if (!header) return jsonResponse_({ ok: false, error: 'Paket tidak ditemukan.' });
    if (!header.aktif) return jsonResponse_({ ok: false, error: 'Paket nonaktif. Aktifkan dulu.' });
  }

  shR.getRange(foundRowIdx, 44).setValue(idPriceList || '');

  try {
    logAktivitas_(user, 'assignPriceListToReseller',
      'Reseller ' + idReseller + ': "' + currentIdPL + '" -> "' + idPriceList + '"');
  } catch (e) {}

  return jsonResponse_({
    ok: true,
    data: { id_reseller: idReseller, id_price_list_lama: currentIdPL, id_price_list_baru: idPriceList }
  });
}

function handleGetResellerAssignments_(body) {
  const user = validateToken_(body.token);
  if (!user) return jsonResponse_({ ok: false, error: 'Token tidak valid atau expired.' });
  if (!requireAdmin_(user)) return jsonResponse_({ ok: false, error: 'Hanya admin/owner.' });

  const shR = getSheet_(CONFIG.SHEETS.RESELLER);
  const lastR = shR.getLastRow();
  const list = [];
  if (lastR >= 2) {
    const dataR = shR.getRange(2, 1, lastR - 1, 44).getValues();

    const shH = getSheet_(CONFIG.SHEETS.RESELLER_HARGA);
    const lastH = shH.getLastRow();
    const customCountMap = {};
    if (lastH >= 2) {
      const dataH = shH.getRange(2, 1, lastH - 1, 2).getValues();
      dataH.forEach(function(r) {
        const idR = String(r[0]);
        customCountMap[idR] = (customCountMap[idR] || 0) + 1;
      });
    }

    const shPL = getSheet_(CONFIG.SHEETS.PRICE_LISTS);
    const lastPL = shPL.getLastRow();
    const priceListMap = {};
    if (lastPL >= 2) {
      const dataPL = shPL.getRange(2, 1, lastPL - 1, 7).getValues();
      dataPL.forEach(function(r) {
        priceListMap[String(r[0])] = { nama_paket: String(r[1] || ''), aktif: String(r[3] || '').toUpperCase() === 'TRUE' };
      });
    }

    for (let i = 0; i < dataR.length; i++) {
      const r = dataR[i];
      const idR = String(r[38] || '');
      if (!idR) continue;
      const idPL = String(r[43] || '');
      const plInfo = priceListMap[idPL] || null;
      list.push({
        id_reseller: idR,
        nama_reseller: String(r[1] || ''),
        aktif: String(r[39] || '').toUpperCase() === 'TRUE',
        id_price_list: idPL,
        nama_paket: plInfo ? plInfo.nama_paket : '',
        paket_aktif: plInfo ? plInfo.aktif : null,
        custom_override_count: customCountMap[idR] || 0
      });
    }
  }
  list.sort(function(a, b) { return a.nama_reseller.localeCompare(b.nama_reseller); });
  return jsonResponse_({ ok: true, data: { total: list.length, items: list } });
}

function handleBulkAssignPriceList_(body) {
  const user = validateToken_(body.token);
  if (!user) return jsonResponse_({ ok: false, error: 'Token tidak valid atau expired.' });
  if (!requireAdmin_(user)) return jsonResponse_({ ok: false, error: 'Hanya admin/owner.' });

  const idResellerList = Array.isArray(body.id_reseller_list) ? body.id_reseller_list.map(String) : [];
  const idPriceList = String(body.id_price_list || '').trim();

  if (idResellerList.length === 0) return jsonResponse_({ ok: false, error: 'id_reseller_list wajib (min 1).' });
  if (idResellerList.length > 100) return jsonResponse_({ ok: false, error: 'Maksimal 100 reseller per batch.' });

  if (idPriceList) {
    const header = getPriceListById_(idPriceList);
    if (!header) return jsonResponse_({ ok: false, error: 'Paket tidak ditemukan.' });
    if (!header.aktif) return jsonResponse_({ ok: false, error: 'Paket nonaktif. Aktifkan dulu.' });
  }

  const targetSet = {};
  idResellerList.forEach(function(id) { targetSet[id] = true; });

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    const shR = getSheet_(CONFIG.SHEETS.RESELLER);
    const lastR = shR.getLastRow();
    if (lastR < 2) return jsonResponse_({ ok: false, error: 'Tidak ada reseller.' });

    const dataR = shR.getRange(2, 1, lastR - 1, 44).getValues();
    const updated = [];
    const notFound = [];

    for (let i = 0; i < dataR.length; i++) {
      const idR = String(dataR[i][38] || '');
      if (!idR || !targetSet[idR]) continue;
      const currentPL = String(dataR[i][43] || '');
      if (currentPL === idPriceList) {
        notFound.push({ id_reseller: idR, reason: 'sudah pakai paket ini' });
        continue;
      }
      shR.getRange(i + 2, 44).setValue(idPriceList || '');
      updated.push(idR);
    }

    idResellerList.forEach(function(id) {
      if (updated.indexOf(id) === -1 && !notFound.some(function(x) { return x.id_reseller === id; })) {
        notFound.push({ id_reseller: id, reason: 'tidak ditemukan' });
      }
    });

    updated.forEach(function(id) {
      cacheInvalidate_(['reseller_pl_' + id]);
    });

    try {
      logAktivitas_(user, 'bulkAssignPriceList',
        'Bulk assign paket="' + idPriceList + '" ke ' + updated.length + ' reseller (skip ' + notFound.length + ')');
    } catch (e) {}

    return jsonResponse_({
      ok: true,
      data: {
        id_price_list: idPriceList,
        total_updated: updated.length,
        total_skipped: notFound.length,
        updated_list: updated,
        skipped_list: notFound
      }
    });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// ============================================================
// F9.7A — PROGRAMS (PROGRAM HARGA / DISKON) — CRUD
// ============================================================

function getProgramById_(idProgram) {
  const sh = getSheet_(CONFIG.SHEETS.PROGRAMS);
  const last = sh.getLastRow();
  if (last < 2) return null;
  const data = sh.getRange(2, 1, last - 1, 21).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(idProgram)) {
      return {
        id_program: String(data[i][0]),
        nama_program: String(data[i][1] || ''),
        deskripsi: String(data[i][2] || ''),
        tgl_mulai: String(data[i][3] || ''),
        tgl_selesai: String(data[i][4] || ''),
        aktif: String(data[i][5] || '').toUpperCase() === 'TRUE',
        mode_approval: String(data[i][6] || 'auto'),
        visibility: String(data[i][7] || 'visible'),
        kombinasi_mode: String(data[i][8] || 'exclusive'),
        kalkulasi_mode: String(data[i][9] || 'sequential'),
        priority: Number(data[i][10]) || 10,
        target_type: String(data[i][11] || 'all'),
        target_ids: String(data[i][12] || '[]'),
        basis_harga: String(data[i][13] || 'auto'),
        max_diskon_per_order: Number(data[i][14]) || 0,
        max_usage_per_customer: Number(data[i][15]) || 0,
        max_usage_total: Number(data[i][16]) || 0,
        current_usage: Number(data[i][17]) || 0,
        tgl_dibuat: String(data[i][18] || ''),
        diupdate_oleh: String(data[i][19] || ''),
        tgl_update_terakhir: String(data[i][20] || ''),
        rowIndex: i + 2
      };
    }
  }
  return null;
}

function getProgramRules_(idProgram) {
  const sh = getSheet_(CONFIG.SHEETS.PROGRAM_RULES);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const data = sh.getRange(2, 1, last - 1, 9).getValues();
  const rules = [];
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][1]) !== String(idProgram)) continue;
    rules.push({
      id_rule: String(data[i][0]),
      id_program: String(data[i][1]),
      urutan: Number(data[i][2]) || 1,
      rule_type: String(data[i][3] || ''),
      scope: String(data[i][4] || 'all'),
      scope_ids: String(data[i][5] || '[]'),
      parameter: String(data[i][6] || '{}'),
      tgl_dibuat: String(data[i][7] || ''),
      tgl_update_terakhir: String(data[i][8] || ''),
      rowIndex: i + 2
    });
  }
  rules.sort(function(a, b) { return a.urutan - b.urutan; });
  return rules;
}

function getNextProgramId_() {
  const sh = getSheet_(CONFIG.SHEETS.PROGRAMS);
  const last = sh.getLastRow();
  if (last < 2) return 'KK22-PRG-0001';
  const data = sh.getRange(2, 1, last - 1, 1).getValues();
  let maxNum = 0;
  data.forEach(function(r) {
    const m = String(r[0] || '').match(/-PRG-(\d+)/);
    if (m) { const n = parseInt(m[1], 10); if (n > maxNum) maxNum = n; }
  });
  return 'KK22-PRG-' + pad4_(maxNum + 1);
}

function getMaxProgramRuleNum_() {
  const sh = getSheet_(CONFIG.SHEETS.PROGRAM_RULES);
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const data = sh.getRange(2, 1, last - 1, 1).getValues();
  let maxNum = 0;
  data.forEach(function(r) {
    const m = String(r[0] || '').match(/-PRL-(\d+)/);
    if (m) { const n = parseInt(m[1], 10); if (n > maxNum) maxNum = n; }
  });
  return maxNum;
}

const VALID_PROGRAM_RULE_TYPES = [
  'item_fixed_price', 'item_discount_pct', 'item_discount_fixed',
  'item_tiered_qty', 'item_bxgy', 'item_bundle',
  'order_discount_pct', 'order_discount_fixed',
  'order_tiered_total', 'order_tiered_qty',
  'order_min_qty', 'order_min_value',
  'order_free_shipping', 'gift_with_purchase'
];

function validateProgramRules_(rules) {
  if (!Array.isArray(rules) || rules.length === 0) return { ok: false, error: 'Minimal 1 rule.' };
  for (let i = 0; i < rules.length; i++) {
    const rt = String(rules[i].rule_type || '').toLowerCase();
    if (VALID_PROGRAM_RULE_TYPES.indexOf(rt) === -1) {
      return { ok: false, error: 'Rule ke-' + (i + 1) + ': rule_type tidak valid: ' + rt };
    }
  }
  return { ok: true };
}

function handleGetPrograms_(body) {
  const user = validateToken_(body.token);
  if (!user) return jsonResponse_({ ok: false, error: 'Token tidak valid atau expired.' });
  if (!requireAdmin_(user)) return jsonResponse_({ ok: false, error: 'Hanya admin/owner.' });

  const sh = getSheet_(CONFIG.SHEETS.PROGRAMS);
  const last = sh.getLastRow();
  const items = [];
  if (last >= 2) {
    const data = sh.getRange(2, 1, last - 1, 21).getValues();

    const shR = getSheet_(CONFIG.SHEETS.PROGRAM_RULES);
    const lastR = shR.getLastRow();
    const ruleCount = {};
    if (lastR >= 2) {
      const dataR = shR.getRange(2, 1, lastR - 1, 2).getValues();
      dataR.forEach(function(r) {
        const idP = String(r[1]);
        ruleCount[idP] = (ruleCount[idP] || 0) + 1;
      });
    }

    for (let i = 0; i < data.length; i++) {
      const r = data[i];
      items.push({
        id_program: String(r[0]),
        nama_program: String(r[1] || ''),
        deskripsi: String(r[2] || ''),
        tgl_mulai: String(r[3] || ''),
        tgl_selesai: String(r[4] || ''),
        aktif: String(r[5] || '').toUpperCase() === 'TRUE',
        mode_approval: String(r[6] || 'auto'),
        visibility: String(r[7] || 'visible'),
        kombinasi_mode: String(r[8] || 'exclusive'),
        kalkulasi_mode: String(r[9] || 'sequential'),
        priority: Number(r[10]) || 10,
        target_type: String(r[11] || 'all'),
        basis_harga: String(r[13] || 'auto'),
        total_rules: ruleCount[String(r[0])] || 0,
        current_usage: Number(r[17]) || 0,
        tgl_dibuat: String(r[18] || '')
      });
    }
  }
  items.sort(function(a, b) { return b.priority - a.priority || a.nama_program.localeCompare(b.nama_program); });
  return jsonResponse_({ ok: true, data: { total: items.length, items: items } });
}

function handleGetProgramDetail_(body) {
  const user = validateToken_(body.token);
  if (!user) return jsonResponse_({ ok: false, error: 'Token tidak valid atau expired.' });
  if (!requireAdmin_(user)) return jsonResponse_({ ok: false, error: 'Hanya admin/owner.' });

  const idProgram = String(body.id_program || '').trim();
  if (!idProgram) return jsonResponse_({ ok: false, error: 'id_program wajib diisi.' });

  const header = getProgramById_(idProgram);
  if (!header) return jsonResponse_({ ok: false, error: 'Program tidak ditemukan.' });

  const rules = getProgramRules_(idProgram);
  return jsonResponse_({ ok: true, data: { header: header, rules: rules } });
}

function handleCreateProgram_(body) {
  const user = validateToken_(body.token);
  if (!user) return jsonResponse_({ ok: false, error: 'Token tidak valid atau expired.' });
  if (!requireAdmin_(user)) return jsonResponse_({ ok: false, error: 'Hanya admin/owner.' });

  const namaProgram = String(body.nama_program || '').trim();
  if (!namaProgram) return jsonResponse_({ ok: false, error: 'nama_program wajib diisi.' });

  const deskripsi = String(body.deskripsi || '').trim();
  const tglMulai = String(body.tgl_mulai || '').trim();
  const tglSelesai = String(body.tgl_selesai || '').trim();
  const modeApproval = String(body.mode_approval || 'auto').toLowerCase();
  const visibility = String(body.visibility || 'visible').toLowerCase();
  const kombinasiMode = String(body.kombinasi_mode || 'exclusive').toLowerCase();
  const kalkulasiMode = String(body.kalkulasi_mode || 'sequential').toLowerCase();
  const priority = Number(body.priority) || 10;
  const targetType = String(body.target_type || 'all').toLowerCase();
  const targetIds = Array.isArray(body.target_ids) ? body.target_ids : [];
  const basisHarga = String(body.basis_harga || 'auto').toLowerCase();
  const maxDiskon = Number(body.max_diskon_per_order) || 0;
  const maxUsagePerCustomer = Number(body.max_usage_per_customer) || 0;
  const maxUsageTotal = Number(body.max_usage_total) || 0;
  const rules = Array.isArray(body.rules) ? body.rules : [];

  if (['auto','manual'].indexOf(modeApproval) === -1) return jsonResponse_({ ok: false, error: 'mode_approval harus: auto/manual' });
  if (['visible','auto_apply'].indexOf(visibility) === -1) return jsonResponse_({ ok: false, error: 'visibility harus: visible/auto_apply' });
  if (['exclusive','stackable'].indexOf(kombinasiMode) === -1) return jsonResponse_({ ok: false, error: 'kombinasi_mode harus: exclusive/stackable' });
  if (['sequential','parallel','best_for_customer'].indexOf(kalkulasiMode) === -1) return jsonResponse_({ ok: false, error: 'kalkulasi_mode harus: sequential/parallel/best_for_customer' });
  if (['all','filter','custom'].indexOf(targetType) === -1) return jsonResponse_({ ok: false, error: 'target_type harus: all/filter/custom' });
  if (['auto','normal','paket','custom'].indexOf(basisHarga) === -1) return jsonResponse_({ ok: false, error: 'basis_harga harus: auto/normal/paket/custom' });
  if (tglMulai && !/^\d{4}-\d{2}-\d{2}$/.test(tglMulai)) return jsonResponse_({ ok: false, error: 'tgl_mulai harus YYYY-MM-DD' });
  if (tglSelesai && !/^\d{4}-\d{2}-\d{2}$/.test(tglSelesai)) return jsonResponse_({ ok: false, error: 'tgl_selesai harus YYYY-MM-DD' });

  const v = validateProgramRules_(rules);
  if (!v.ok) return jsonResponse_({ ok: false, error: v.error });

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    const newId = getNextProgramId_();
    const now = nowJakarta_();
    const targetIdsJson = JSON.stringify(targetIds);

    const sh = getSheet_(CONFIG.SHEETS.PROGRAMS);
    const newRow = sh.getLastRow() + 1;
    sh.getRange(newRow, 1, 1, 21).setValues([[
      newId, namaProgram, deskripsi,
      tglMulai, tglSelesai, 'TRUE',
      modeApproval, visibility, kombinasiMode, kalkulasiMode, priority,
      targetType, targetIdsJson, basisHarga,
      maxDiskon, maxUsagePerCustomer, maxUsageTotal, 0,
      now, user.id_user, now
    ]]);

    const shR = getSheet_(CONFIG.SHEETS.PROGRAM_RULES);
    const maxNum = getMaxProgramRuleNum_();
    const rows = [];
    rules.forEach(function(r, idx) {
      const scopeIds = Array.isArray(r.scope_ids) ? r.scope_ids : [];
      const parameter = r.parameter || {};
      rows.push([
        'KK22-PRL-' + pad4_(maxNum + idx + 1),
        newId, Number(r.urutan) || (idx + 1),
        String(r.rule_type || '').toLowerCase(),
        String(r.scope || 'all').toLowerCase(),
        JSON.stringify(scopeIds),
        JSON.stringify(parameter),
        now, now
      ]);
    });
    if (rows.length > 0) {
      const startRow = shR.getLastRow() + 1;
      shR.getRange(startRow, 1, rows.length, 9).setValues(rows);
    }

    try {
      logAktivitas_(user, 'createProgram',
        'Program ' + newId + ' (' + namaProgram + '), ' + rules.length + ' rule, target=' + targetType);
    } catch (e) {}

    return jsonResponse_({
      ok: true,
      data: { id_program: newId, nama_program: namaProgram, total_rules: rules.length, mode_approval: modeApproval }
    });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function handleUpdateProgram_(body) {
  const user = validateToken_(body.token);
  if (!user) return jsonResponse_({ ok: false, error: 'Token tidak valid atau expired.' });
  if (!requireAdmin_(user)) return jsonResponse_({ ok: false, error: 'Hanya admin/owner.' });

  const idProgram = String(body.id_program || '').trim();
  if (!idProgram) return jsonResponse_({ ok: false, error: 'id_program wajib diisi.' });

  const header = getProgramById_(idProgram);
  if (!header) return jsonResponse_({ ok: false, error: 'Program tidak ditemukan.' });

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    const sh = getSheet_(CONFIG.SHEETS.PROGRAMS);
    const now = nowJakarta_();
    const colMap = {
      nama_program: 2, deskripsi: 3, tgl_mulai: 4, tgl_selesai: 5,
      mode_approval: 7, visibility: 8, kombinasi_mode: 9, kalkulasi_mode: 10,
      priority: 11, target_type: 12, target_ids: 13, basis_harga: 14,
      max_diskon_per_order: 15, max_usage_per_customer: 16, max_usage_total: 17
    };
    const updated = [];
    Object.keys(colMap).forEach(function(k) {
      if (body[k] === undefined) return;
      let v = body[k];
      if (k === 'target_ids') v = JSON.stringify(Array.isArray(v) ? v : []);
      sh.getRange(header.rowIndex, colMap[k]).setValue(v);
      updated.push(k);
    });
    sh.getRange(header.rowIndex, 20).setValue(user.id_user);
    sh.getRange(header.rowIndex, 21).setValue(now);

    try {
      logAktivitas_(user, 'updateProgram', 'Program ' + idProgram + ' diupdate: ' + updated.join(', '));
    } catch (e) {}

    return jsonResponse_({ ok: true, data: { id_program: idProgram, fields_updated: updated } });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function handleToggleProgramAktif_(body) {
  const user = validateToken_(body.token);
  if (!user) return jsonResponse_({ ok: false, error: 'Token tidak valid atau expired.' });
  if (!requireAdmin_(user)) return jsonResponse_({ ok: false, error: 'Hanya admin/owner.' });

  const idProgram = String(body.id_program || '').trim();
  const aktifBaru = body.aktif;
  if (!idProgram) return jsonResponse_({ ok: false, error: 'id_program wajib diisi.' });
  if (typeof aktifBaru !== 'boolean') return jsonResponse_({ ok: false, error: 'aktif harus boolean.' });

  const header = getProgramById_(idProgram);
  if (!header) return jsonResponse_({ ok: false, error: 'Program tidak ditemukan.' });

  const sh = getSheet_(CONFIG.SHEETS.PROGRAMS);
  const now = nowJakarta_();
  sh.getRange(header.rowIndex, 6).setValue(aktifBaru ? 'TRUE' : 'FALSE');
  sh.getRange(header.rowIndex, 20).setValue(user.id_user);
  sh.getRange(header.rowIndex, 21).setValue(now);

  try {
    logAktivitas_(user, 'toggleProgramAktif', 'Program ' + idProgram + ' aktif=' + aktifBaru);
  } catch (e) {}

  return jsonResponse_({ ok: true, data: { id_program: idProgram, aktif: aktifBaru } });
}

function handleSetProgramRules_(body) {
  const user = validateToken_(body.token);
  if (!user) return jsonResponse_({ ok: false, error: 'Token tidak valid atau expired.' });
  if (!requireAdmin_(user)) return jsonResponse_({ ok: false, error: 'Hanya admin/owner.' });

  const idProgram = String(body.id_program || '').trim();
  const rules = Array.isArray(body.rules) ? body.rules : [];
  if (!idProgram) return jsonResponse_({ ok: false, error: 'id_program wajib diisi.' });

  const header = getProgramById_(idProgram);
  if (!header) return jsonResponse_({ ok: false, error: 'Program tidak ditemukan.' });

  const v = validateProgramRules_(rules);
  if (!v.ok) return jsonResponse_({ ok: false, error: v.error });

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    const shR = getSheet_(CONFIG.SHEETS.PROGRAM_RULES);
    const lastR = shR.getLastRow();
    if (lastR >= 2) {
      const dataR = shR.getRange(2, 1, lastR - 1, 2).getValues();
      const toDelete = [];
      for (let i = 0; i < dataR.length; i++) {
        if (String(dataR[i][1]) === idProgram) toDelete.push(i + 2);
      }
      toDelete.sort(function(a, b) { return b - a; });
      toDelete.forEach(function(rowIdx) { shR.deleteRow(rowIdx); });
    }

    const now = nowJakarta_();
    const maxNum = getMaxProgramRuleNum_();
    const rows = [];
    rules.forEach(function(r, idx) {
      const scopeIds = Array.isArray(r.scope_ids) ? r.scope_ids : [];
      const parameter = r.parameter || {};
      rows.push([
        'KK22-PRL-' + pad4_(maxNum + idx + 1),
        idProgram, Number(r.urutan) || (idx + 1),
        String(r.rule_type || '').toLowerCase(),
        String(r.scope || 'all').toLowerCase(),
        JSON.stringify(scopeIds),
        JSON.stringify(parameter),
        now, now
      ]);
    });
    if (rows.length > 0) {
      const startRow = shR.getLastRow() + 1;
      shR.getRange(startRow, 1, rows.length, 9).setValues(rows);
    }

    try {
      logAktivitas_(user, 'setProgramRules', 'Program ' + idProgram + ': ' + rules.length + ' rules di-set');
    } catch (e) {}

    return jsonResponse_({ ok: true, data: { id_program: idProgram, total_rules: rules.length } });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function handleDeleteProgram_(body) {
  const user = validateToken_(body.token);
  if (!user) return jsonResponse_({ ok: false, error: 'Token tidak valid atau expired.' });
  if (!requireAdmin_(user)) return jsonResponse_({ ok: false, error: 'Hanya admin/owner.' });

  const idProgram = String(body.id_program || '').trim();
  if (!idProgram) return jsonResponse_({ ok: false, error: 'id_program wajib diisi.' });

  const header = getProgramById_(idProgram);
  if (!header) return jsonResponse_({ ok: false, error: 'Program tidak ditemukan.' });

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    const shR = getSheet_(CONFIG.SHEETS.PROGRAM_RULES);
    const lastR = shR.getLastRow();
    if (lastR >= 2) {
      const dataR = shR.getRange(2, 1, lastR - 1, 2).getValues();
      const toDelete = [];
      for (let i = 0; i < dataR.length; i++) {
        if (String(dataR[i][1]) === idProgram) toDelete.push(i + 2);
      }
      toDelete.sort(function(a, b) { return b - a; });
      toDelete.forEach(function(rowIdx) { shR.deleteRow(rowIdx); });
    }

    const sh = getSheet_(CONFIG.SHEETS.PROGRAMS);
    sh.deleteRow(header.rowIndex);

    try {
      logAktivitas_(user, 'deleteProgram', 'Program ' + idProgram + ' (' + header.nama_program + ') dihapus');
    } catch (e) {}

    return jsonResponse_({ ok: true, data: { id_program: idProgram, nama_program: header.nama_program } });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// ============================================================
// HELPERS
// ============================================================

function findResellerById_(idReseller) {
  const sh = getSheet_(CONFIG.SHEETS.RESELLER);
  const last = sh.getLastRow();
  if (last < 2) return null;
  const data = sh.getRange(2, 1, last - 1, 44).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][38]) === String(idReseller)) {
      return {
        id_reseller: String(data[i][38]),
        nama_reseller: String(data[i][1] || ''),
        aktif: data[i][39],
        id_price_list: String(data[i][43] || '')
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