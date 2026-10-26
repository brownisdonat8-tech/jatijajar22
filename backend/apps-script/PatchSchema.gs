/**
 * KK22 - Patch Schema
 * Menambahkan kolom baru ke sheet yang sudah ada (tidak menghapus data).
 * Jalankan SEKALI saja setiap kali dibutuhkan.
 */

/**
 * Tambah 3 kolom ke sheet DETAIL_PESANAN:
 *   - status_item  (pending / locked / shipped / done)
 *   - locked_at    (timestamp saat di-lock admin)
 *   - locked_by    (id_user admin yang lock)
 */
function patchDetailPesananAddStatusItem() {
  const sh = getSheet_(CONFIG.SHEETS.DETAIL_PESANAN);
  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(function(h){ return String(h).trim(); });

  // Cek apakah sudah ada
  const hasStatusItem = headers.indexOf('status_item') !== -1;
  const hasLockedAt   = headers.indexOf('locked_at') !== -1;
  const hasLockedBy   = headers.indexOf('locked_by') !== -1;

  if (hasStatusItem && hasLockedAt && hasLockedBy) {
    Logger.log('[i] Kolom sudah lengkap. Tidak ada perubahan.');
    return;
  }

  // Backup dulu: log semua header saat ini
  Logger.log('===== SEBELUM PATCH =====');
  Logger.log('Kolom terakhir: ' + lastCol);
  Logger.log('Headers: ' + headers.join(' | '));

  // Tambahkan yang belum ada
  let addCol = lastCol;
  if (!hasStatusItem) {
    addCol++;
    sh.getRange(1, addCol).setValue('status_item')
      .setFontWeight('bold')
      .setBackground('#0b5394')
      .setFontColor('#ffffff');
    Logger.log('[OK] Kolom status_item ditambahkan di kolom ' + addCol);
  }
  if (!hasLockedAt) {
    addCol++;
    sh.getRange(1, addCol).setValue('locked_at')
      .setFontWeight('bold')
      .setBackground('#0b5394')
      .setFontColor('#ffffff');
    Logger.log('[OK] Kolom locked_at ditambahkan di kolom ' + addCol);
  }
  if (!hasLockedBy) {
    addCol++;
    sh.getRange(1, addCol).setValue('locked_by')
      .setFontWeight('bold')
      .setBackground('#0b5394')
      .setFontColor('#ffffff');
    Logger.log('[OK] Kolom locked_by ditambahkan di kolom ' + addCol);
  }

  Logger.log('===== PATCH SELESAI =====');
  Logger.log('Kolom baru: ' + (addCol - lastCol));
}

/**
 * Verifikasi: tampilkan struktur sheet DETAIL_PESANAN
 */
function verifikasiDetailPesanan() {
  const sh = getSheet_(CONFIG.SHEETS.DETAIL_PESANAN);
  const lastCol = sh.getLastColumn();
  const lastRow = sh.getLastRow();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];

  Logger.log('===== STRUKTUR DETAIL_PESANAN =====');
  Logger.log('Jumlah kolom : ' + lastCol);
  Logger.log('Jumlah baris : ' + lastRow);
  Logger.log('');
  Logger.log('--- Header ---');
  for (let i = 0; i < headers.length; i++) {
    Logger.log('  [' + (i+1) + '] ' + headers[i]);
  }
}