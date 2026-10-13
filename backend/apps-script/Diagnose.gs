/**
 * KK22 - Diagnose
 * Fungsi bantu untuk cek struktur file sumber.
 */

function diagnoseReseller() {
  const ss = SpreadsheetApp.openById(MIGRATE_SOURCES.RESELLER);
  const sh = ss.getSheets()[0];
  const name = sh.getName();
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();

  Logger.log('===== DIAGNOSE FILE RESELLER =====');
  Logger.log('Sheet pertama    : ' + name);
  Logger.log('Jumlah baris     : ' + lastRow);
  Logger.log('Jumlah kolom     : ' + lastCol);

  if (lastRow < 2 || lastCol < 1) {
    Logger.log('File kosong atau tidak ada data.');
    return;
  }

  // Tampilkan semua header + index
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  Logger.log('--- HEADER DENGAN INDEX ---');
  for (let i = 0; i < headers.length; i++) {
    Logger.log('  [' + i + '] ' + headers[i]);
  }

  // Tampilkan 3 baris pertama dengan isi kolom yang penting
  Logger.log('--- SAMPLE 3 BARIS PERTAMA ---');
  const sampleEnd = Math.min(4, lastRow);
  const sample = sh.getRange(1, 1, sampleEnd, lastCol).getValues();
  for (let r = 1; r < sample.length; r++) {
    Logger.log('  Baris ' + (r+1) + ':');
    for (let c = 0; c < lastCol; c++) {
      const val = sample[r][c];
      if (val !== '' && val !== null && val !== undefined) {
        const v = String(val);
        Logger.log('    [' + c + '] ' + headers[c] + ' = ' + v.substring(0, 60));
      }
    }
  }

  // Hitung jumlah baris yang punya kata "customer" di tiap kolom
  Logger.log('--- CARI KOLOM YANG MENGANDUNG "customer" ---');
  const allData = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  for (let c = 0; c < lastCol; c++) {
    let count = 0;
    for (let r = 0; r < allData.length; r++) {
      const v = String(allData[r][c] || '').toLowerCase();
      if (v.indexOf('customer') !== -1) count++;
    }
    if (count > 0) {
      Logger.log('  Kolom [' + c + '] ' + headers[c] + ' -> ' + count + ' baris mengandung "customer"');
    }
  }
}