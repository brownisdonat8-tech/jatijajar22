/**
 * KK22 - Maintenance
 * Fungsi-fungsi untuk pemeliharaan data.
 * Jalankan manual dari Apps Script editor, BUKAN via API.
 */

// ============================================================
// RESET STOK NEGATIF KE NOL
// ============================================================
// Konteks: data awal dari Mekari banyak stok negatif (legacy).
// Karena sistem baru TIDAK memakai stok sebagai pembatas pesanan
// (produksi dilakukan setelah PO terkumpul), stok negatif
// di-reset ke 0 agar tampilan bersih.
//
// Fungsi ini:
//   - Scan kolom E (Stock) sheet PRODUK
//   - Kalau < 0, ubah jadi 0
//   - Catat ke STOK_LOG untuk audit
//   - Laporkan ringkasan di log

function resetStokNegatifKeNol() {
  const sh = getSheet_(CONFIG.SHEETS.PRODUK);
  const last = sh.getLastRow();
  if (last < 2) {
    Logger.log('Sheet PRODUK kosong.');
    return;
  }

  const data = sh.getRange(2, 1, last - 1, 14).getValues();
  const now = nowJakarta_();
  const logRows = [];

  let count = 0;
  let totalNegatif = 0;
  let totalHilang = 0;

  for (let i = 0; i < data.length; i++) {
    const nama = String(data[i][0] || '').trim();
    const stok = Number(data[i][4]) || 0;
    if (!nama) continue;
    if (stok >= 0) continue;

    totalNegatif += stok;
    totalHilang += Math.abs(stok);
    count++;

    // Update kolom E (Stock) = 0
    sh.getRange(i + 2, 5).setValue(0);

    // Catat ke STOK_LOG
    logRows.push([
      'KK22-LOG-' + pad4_(count),
      now,
      nama,
      Math.abs(stok),
      stok,
      0,
      'Reset stok negatif ke nol (migrasi legacy)',
      'system'
    ]);
  }

  // Bulk insert ke STOK_LOG
  if (logRows.length > 0) {
    const logSheet = getSheet_(CONFIG.SHEETS.STOK_LOG);
    const logLast = logSheet.getLastRow();
    // Fix: gunakan nomor urut dari row terakhir
    const startId = logLast;
    for (let k = 0; k < logRows.length; k++) {
      logRows[k][0] = 'KK22-LOG-' + pad4_(startId + k + 1);
    }
    logSheet.getRange(logSheet.getLastRow() + 1, 1, logRows.length, 8).setValues(logRows);
  }

  Logger.log('===== RESET STOK SELESAI =====');
  Logger.log('Total produk di-reset : ' + count);
  Logger.log('Total stok negatif (sum) : ' + totalNegatif);
  Logger.log('Total kekurangan yang dihapus : ' + totalHilang);
  Logger.log('Semua stok < 0 sudah di-set jadi 0.');
  Logger.log('');
  Logger.log('Catatan: audit trail tersimpan di sheet STOK_LOG.');
}

// ============================================================
// CEK RINGKASAN STOK SETELAH RESET
// ============================================================

function ringkasanStok() {
  const sh = getSheet_(CONFIG.SHEETS.PRODUK);
  const last = sh.getLastRow();
  if (last < 2) return;

  const data = sh.getRange(2, 1, last - 1, 14).getValues();
  let nol = 0, positif = 0, negatif = 0, kosong = 0;

  for (let i = 0; i < data.length; i++) {
    const stok = data[i][4];
    if (stok === '' || stok === null) { kosong++; }
    else if (Number(stok) === 0) { nol++; }
    else if (Number(stok) > 0) { positif++; }
    else if (Number(stok) < 0) { negatif++; }
  }

  Logger.log('===== RINGKASAN STOK PRODUK =====');
  Logger.log('Total baris       : ' + data.length);
  Logger.log('Stok > 0 (ada)    : ' + positif);
  Logger.log('Stok = 0 (kosong) : ' + nol);
  Logger.log('Stok < 0 (minus)  : ' + negatif + (negatif > 0 ? '  <-- MASIH ADA, jalankan resetStokNegatifKeNol()' : ''));
  Logger.log('Stok kosong       : ' + kosong);
}
// ============================================================
// CLEANUP SEMUA PESANAN (HAPUS DATA TEST)
// ============================================================
// PERINGATAN: Menghapus SEMUA baris di sheet PESANAN & DETAIL_PESANAN.
// Hanya untuk development/testing. JANGAN jalankan di production.

function cleanupSemuaPesanan() {
  const shP = getSheet_(CONFIG.SHEETS.PESANAN);
  const shD = getSheet_(CONFIG.SHEETS.DETAIL_PESANAN);

  const lastP = shP.getLastRow();
  const lastD = shD.getLastRow();

  if (lastP >= 2) {
    shP.deleteRows(2, lastP - 1);
    Logger.log('[OK] Hapus ' + (lastP - 1) + ' baris PESANAN');
  } else {
    Logger.log('[i] Sheet PESANAN sudah kosong.');
  }

  if (lastD >= 2) {
    shD.deleteRows(2, lastD - 1);
    Logger.log('[OK] Hapus ' + (lastD - 1) + ' baris DETAIL_PESANAN');
  } else {
    Logger.log('[i] Sheet DETAIL_PESANAN sudah kosong.');
  }

  Logger.log('');
  Logger.log('===== CLEANUP SELESAI =====');
  Logger.log('Semua data pesanan test sudah dihapus.');
  Logger.log('Sheet PESANAN sekarang kosong (hanya header).');
  Logger.log('Sheet DETAIL_PESANAN sekarang kosong (hanya header).');
}