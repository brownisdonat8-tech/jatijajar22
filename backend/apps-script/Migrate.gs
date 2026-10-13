/**
 * KK22 - Migrate
 * Baca file XLSX Mekari (yang sudah di-convert ke Google Sheets) dan
 * migrate ke sheet PRODUK / RESELLER.
 *
 * Cara pakai:
 *   1. Jalankan previewProduk()    -> lihat ringkasan
 *   2. Kalau OK, jalankan migrateProduk()
 *   3. Jalankan previewReseller()  -> lihat ringkasan
 *   4. Kalau OK, jalankan migrateReseller()
 */

const MIGRATE_SOURCES = {
  PRODUK:   '1Nej89ghERKkvHru8oUZ6A2HudoWO4Gzz',
  RESELLER: '1ni1dO3o3Fafjm4gKHs4xWRyiFh7B5Nkd'
};

// ============================================================
// PRODUK
// ============================================================

function previewProduk() {
  const data = readProdukSource_();
  const seen = {};
  const dupes = [];
  data.forEach(function(r) {
    const key = String(r[0]).trim().toLowerCase();
    if (seen[key]) dupes.push(r[0]);
    else seen[key] = true;
  });
  Logger.log('===== PREVIEW MIGRASI PRODUK =====');
  Logger.log('Total baris terbaca   : ' + data.length);
  Logger.log('Duplikat nama produk  : ' + dupes.length);
  if (dupes.length > 0) {
    Logger.log('Contoh duplikat (10)  : ' + dupes.slice(0, 10).join(' || '));
  }
  Logger.log('--- Contoh 5 baris pertama ---');
  data.slice(0, 5).forEach(function(r) {
    Logger.log('  [' + r[2] + '] ' + r[0] + '  (' + r[13] + ')');
  });
  Logger.log('--- Contoh 5 baris terakhir ---');
  data.slice(-5).forEach(function(r) {
    Logger.log('  [' + r[2] + '] ' + r[0] + '  (' + r[13] + ')');
  });
  Logger.log('');
  Logger.log('Kalau OK, jalankan migrateProduk().');
}

function readProdukSource_() {
  const ss = SpreadsheetApp.openById(MIGRATE_SOURCES.PRODUK);
  const sh = ss.getSheets()[0];
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];

  const header = values[0].map(function(x){return String(x).trim();});
  const idx = {
    name:      header.indexOf('Name*'),
    desc:      header.indexOf('Description'),
    code:      header.indexOf('ProductCode'),
    barcode:   header.indexOf('#Barcode'),
    stock:     header.indexOf('Stock'),
    unit:      header.indexOf('*Unit'),
    buyPrice:  header.indexOf('BuyPrice'),
    buyAcc:    header.indexOf('DefaultBuyAccountCode'),
    buyTax:    header.indexOf('DefaultBuyTaxName'),
    sellPrice: header.indexOf('SellPrice'),
    sellAcc:   header.indexOf('DefaultSellAccountCode'),
    sellTax:   header.indexOf('DefaultSellTaxName'),
    minStock:  header.indexOf('MinimumStock'),
    category:  header.indexOf('ProductCategory')
  };

  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    const nama = String(r[idx.name] || '').trim();
    const kategori = String(r[idx.category] || '').trim();

    if (!nama) continue;
    if (kategori !== 'Barang Jadi') continue;
    if (isNamaLama_(nama)) continue;

    rows.push([
      nama,
      r[idx.desc] || '',
      r[idx.code] || '',
      r[idx.barcode] || '',
      Number(r[idx.stock]) || 0,
      r[idx.unit] || '',
      Number(r[idx.buyPrice]) || 0,
      r[idx.buyAcc] || '',
      r[idx.buyTax] || '',
      Number(r[idx.sellPrice]) || 0,
      r[idx.sellAcc] || '',
      r[idx.sellTax] || '',
      Number(r[idx.minStock]) || 0,
      kategori
    ]);
  }
  return rows;
}

function isNamaLama_(nama) {
  const prefixes = ['NAMA LAMA#', 'Reject#', 'ERP#', '#ERP', 'tes ', 'tes-'];
  for (let i = 0; i < prefixes.length; i++) {
    if (nama.indexOf(prefixes[i]) === 0) return true;
  }
  if (nama.length > 0 && nama.charAt(nama.length - 1) === '.') return true;
  return false;
}

function migrateProduk() {
  const data = readProdukSource_();
  if (data.length === 0) { Logger.log('Tidak ada data untuk migrate.'); return; }

  const sheet = getSheet_(CONFIG.SHEETS.PRODUK);
  const now = nowJakarta_();
  const lastRow = sheet.getLastRow();
  const startRow = lastRow + 1;

  const existing = {};
  if (lastRow >= 2) {
    const names = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    names.forEach(function(r){ if (r[0]) existing[String(r[0]).trim().toLowerCase()] = true; });
  }

  const output = [];
  const skipped = [];
  let counter = lastRow;

  data.forEach(function(r) {
    const key = String(r[0]).trim().toLowerCase();
    if (existing[key]) { skipped.push(r[0]); return; }
    existing[key] = true;
    counter++;
    const id = CONFIG.PREFIX.PRODUK + pad4_(counter);
    output.push(r.concat([id, '', 0, 'TRUE', now, now]));
  });

  if (output.length > 0) {
    sheet.getRange(startRow, 1, output.length, output[0].length).setValues(output);
  }

  Logger.log('===== MIGRASI PRODUK SELESAI =====');
  Logger.log('Berhasil migrate : ' + output.length + ' produk');
  Logger.log('Skip (duplikat)  : ' + skipped.length);
  if (skipped.length > 0) {
    Logger.log('Contoh skip (10) : ' + skipped.slice(0, 10).join(' || '));
  }
}

function pad4_(n) {
  const s = String(n);
  return s.length >= 4 ? s : ('0000' + s).slice(-4);
}

// ============================================================
// RESELLER
// ============================================================

function previewReseller() {
  const data = readResellerSource_();
  const seen = {};
  const dupes = [];
  data.forEach(function(r) {
    const key = String(r[1]).trim().toLowerCase();
    if (seen[key]) dupes.push(r[1]);
    else seen[key] = true;
  });
  Logger.log('===== PREVIEW MIGRASI RESELLER =====');
  Logger.log('Total baris terbaca   : ' + data.length);
  Logger.log('Duplikat nama         : ' + dupes.length);
  if (dupes.length > 0) {
    Logger.log('Contoh duplikat (10)  : ' + dupes.slice(0, 10).join(' || '));
  }
  Logger.log('--- Contoh 5 baris pertama ---');
  data.slice(0, 5).forEach(function(r) {
    Logger.log('  ' + r[1] + ' | Mobile: ' + r[8] + ' | Type: ' + r[0]);
  });
  Logger.log('--- Contoh 5 baris terakhir ---');
  data.slice(-5).forEach(function(r) {
    Logger.log('  ' + r[1] + ' | Mobile: ' + r[8] + ' | Type: ' + r[0]);
  });
  Logger.log('');
  Logger.log('Kalau OK, jalankan migrateReseller().');
}

function readResellerSource_() {
  const ss = SpreadsheetApp.openById(MIGRATE_SOURCES.RESELLER);
  const sh = ss.getSheets()[0];
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];

  const header = values[0].map(function(x){return String(x).trim();});
  const totalCols = header.length;

  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    const displayName = String(r[1] || '').trim();
    const tipe = String(r[0] || '').trim().toLowerCase();

    if (!displayName) continue;
    if (tipe.indexOf('customer') === -1) continue;

    const row = [];
    for (let c = 0; c < totalCols; c++) {
      row.push(r[c] !== undefined ? r[c] : '');
    }
    rows.push(row);
  }
  return rows;
}

function migrateReseller() {
  const data = readResellerSource_();
  if (data.length === 0) { Logger.log('Tidak ada data untuk migrate.'); return; }

  const sheet = getSheet_(CONFIG.SHEETS.RESELLER);
  const now = nowJakarta_();
  const lastRow = sheet.getLastRow();
  const startRow = lastRow + 1;

  const existing = {};
  if (lastRow >= 2) {
    const names = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    names.forEach(function(r){ if (r[0]) existing[String(r[0]).trim().toLowerCase()] = true; });
  }

  const output = [];
  const skipped = [];
  let counter = lastRow;

  data.forEach(function(r) {
    const key = String(r[1]).trim().toLowerCase();
    if (existing[key]) { skipped.push(r[1]); return; }
    existing[key] = true;
    counter++;
    const id = CONFIG.PREFIX.RESELLER + pad4_(counter);
    output.push(r.concat([id, 'TRUE', '', now, now]));
  });

  if (output.length > 0) {
    sheet.getRange(startRow, 1, output.length, output[0].length).setValues(output);
  }

  Logger.log('===== MIGRASI RESELLER SELESAI =====');
  Logger.log('Berhasil migrate : ' + output.length + ' reseller');
  Logger.log('Skip (duplikat)  : ' + skipped.length);
  if (skipped.length > 0) {
    Logger.log('Contoh skip (10) : ' + skipped.slice(0, 10).join(' || '));
  }
}