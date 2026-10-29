/**
 * KK22 Database - Setup
 * Jalankan setupAllSheets() SEKALI untuk inisialisasi.
 * PERINGATAN: akan mengosongkan sheet kalau sudah ada.
 */

function setupAllSheets() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  // 1. USERS
  createSheet_(ss, CONFIG.SHEETS.USERS, [
    'id_user', 'no_hp', 'pin_hash', 'pin_salt', 'role', 'id_reseller',
    'nama_lengkap', 'aktif', 'tgl_dibuat', 'tgl_update_terakhir'
  ]);

  // 2. RESELLER - MIRROR MEKARI (38 kolom) + TAMBAHAN KITA
  createSheet_(ss, CONFIG.SHEETS.RESELLER, [
    // === 38 kolom dari Mekari Contact Export (JANGAN DIUBAH) ===
    'Type', 'DisplayName', 'CompanyName', 'Title', 'FirstName', 'MiddleName', 'LastName',
    'Email', 'Mobile', 'Phone', 'Fax',
    'BillingAddress', 'BillingAddressNo', 'BillingAddressRT', 'BillingAddressRW',
    'BillingAddressPostCode', 'BillingAddressKelurahan', 'BillingAddressKecamatan',
    'BillingAddressCity', 'BillingAddressProvince', 'ShippingAddress',
    'TaxNumber', '#IdentityType', '#IdentityNumber', 'OtherDetail',
    'ContactGroup (User ; for multiple)',
    'BankID1', 'BranchBankName1', 'NameOnBankAccount1', 'BankAccountNumber1',
    'BankID2', 'BranchBankName2', 'NameOnBankAccount2', 'BankAccountNumber2',
    'BankID3', 'BranchBankName3', 'NameOnBankAccount3', 'BankAccountNumber3',
    // === TAMBAHAN KITA ===
    'id_reseller', 'aktif', 'catatan_internal', 'tgl_dibuat', 'tgl_update_terakhir'
  ]);

  // 3. PRODUK - MIRROR MEKARI (14 kolom) + TAMBAHAN KITA
  createSheet_(ss, CONFIG.SHEETS.PRODUK, [
    // === 14 kolom dari Mekari Product Export (JANGAN DIUBAH) ===
    'Name*', 'Description', 'ProductCode', '#Barcode', 'Stock', '*Unit',
    'BuyPrice', 'DefaultBuyAccountCode', 'DefaultBuyTaxName',
    'SellPrice', 'DefaultSellAccountCode', 'DefaultSellTaxName',
    'MinimumStock', 'ProductCategory',
    // === TAMBAHAN KITA ===
    'id_produk', 'gambar_url', 'stok_reserved', 'aktif', 'tgl_dibuat', 'tgl_update_terakhir'
  ]);

  // 4. STOK_LOG
  createSheet_(ss, CONFIG.SHEETS.STOK_LOG, [
    'id_log', 'waktu', 'nama_produk', 'perubahan', 'stok_sebelum',
    'stok_sesudah', 'alasan', 'oleh_user'
  ]);

  // 5. PESANAN
  createSheet_(ss, CONFIG.SHEETS.PESANAN, [
    'id_pesanan', 'order_number', 'customer_ref_no', 'nama_reseller', 'id_reseller',
    'order_date', 'due_date', 'shipping_date', 'ship_via', 'tracking_no',
    'status', 'total_item', 'total_qty', 'total_harga',
    'catatan', 'memo', 'alasan_tolak',
    'tgl_dibuat', 'diupdate_oleh', 'tgl_update_terakhir'
  ]);

  // 6. DETAIL_PESANAN
  createSheet_(ss, CONFIG.SHEETS.DETAIL_PESANAN, [
    'id_detail', 'id_pesanan', 'order_number', 'nama_produk', 'kode_sku',
    'qty', 'satuan', 'harga_saat_pesan', 'product_discount_rate', 'subtotal', 'catatan'
  ]);

  // 7. LOG_AKTIVITAS
  createSheet_(ss, CONFIG.SHEETS.LOG_AKTIVITAS, [
    'id_log', 'waktu', 'user', 'aksi', 'detail', 'ip_address'
  ]);

  // 8. CONFIG
  createSheet_(ss, CONFIG.SHEETS.CONFIG, [
    'key', 'value', 'keterangan', 'tgl_update'
  ]);

  // 9. JURNAL_MAPPING
  createSheet_(ss, CONFIG.SHEETS.JURNAL_MAPPING, [
    'key', 'value', 'keterangan', 'tgl_update'
  ]);

  // 10. RESELLER_HARGA - Harga khusus per reseller (F9.1)
  createSheet_(ss, CONFIG.SHEETS.RESELLER_HARGA, [
    'id_reseller', 'nama_produk', 'harga_khusus',
    'catatan', 'tgl_dibuat', 'diupdate_oleh', 'tgl_update_terakhir'
  ]);

  // Hapus Sheet1 default kalau ada
  const sheet1 = ss.getSheetByName('Sheet1');
  if (sheet1 && ss.getSheets().length > 1) {
    try { ss.deleteSheet(sheet1); } catch (e) {}
  }

  // Seed data default
  seedConfig_(ss);
  seedJurnalMapping_(ss);

  Logger.log('Setup selesai! 9 sheet berhasil dibuat.');
  Logger.log('Cek: ' + ss.getUrl());
}

function createSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
  } else {
    sh.clear();
  }
  const range = sh.getRange(1, 1, 1, headers.length);
  range.setValues([headers]);
  range.setFontWeight('bold');
  range.setBackground('#0b5394');
  range.setFontColor('#ffffff');
  sh.setFrozenRows(1);
  return sh;
}

function seedConfig_(ss) {
  const sh = ss.getSheetByName(CONFIG.SHEETS.CONFIG);
  const now = nowJakarta_();
  const data = [
    ['versi_aplikasi',      '0.1.0', 'Versi rilis',                    now],
    ['mode_maintenance',    'false', 'true = tolak semua request',     now],
    ['deadline_order_jam',  '15:00', 'Batas jam pesan harian',         now]
  ];
  sh.getRange(2, 1, data.length, 4).setValues(data);
}

function seedJurnalMapping_(ss) {
  const sh = ss.getSheetByName(CONFIG.SHEETS.JURNAL_MAPPING);
  const now = nowJakarta_();
  const data = [
    ['default_currency_code',       'IDR',      'Mata uang untuk export SO',           now],
    ['default_warehouse_name',      '',         'Nama gudang di Mekari (kosong=default)', now],
    ['default_sell_account_code',   '4-40000',  'Kode akun jual default',              now],
    ['default_tax_name',            '',         'Nama pajak (kosong=tidak ada)',       now],
    ['default_tax_rate',            '0',        'Rate pajak (%)',                      now],
    ['default_ship_via',            '',         'Kurir default',                       now],
    ['default_payment_method',      'Cash',     'Metode pembayaran default',           now],
    ['default_paid_to_account',     '',         'Kode akun penerimaan default',        now]
  ];
  sh.getRange(2, 1, data.length, 4).setValues(data);
}

/**
 * Test koneksi ke Spreadsheet
 * Jalankan ini dulu untuk verifikasi Config benar
 */
// ============================================================
// F9.1 - MANUAL: Bikin sheet RESELLER_HARGA (jalankan sekali dari editor)
// ============================================================
function addResellerHargaSheet() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const name = CONFIG.SHEETS.RESELLER_HARGA;
  const existing = ss.getSheetByName(name);
  if (existing) {
    Logger.log('Sheet ' + name + ' sudah ada. Skip.');
    return;
  }
  createSheet_(ss, name, [
    'id_reseller', 'nama_produk', 'harga_khusus',
    'catatan', 'tgl_dibuat', 'diupdate_oleh', 'tgl_update_terakhir'
  ]);
  Logger.log('Sheet ' + name + ' berhasil dibuat dengan 7 kolom header.');
}

// ============================================================
// F9.1b - MANUAL: Setup 2 sheet Price List + kolom RESELLER.id_price_list
// Jalankan sekali dari Apps Script editor. Idempotent (aman run 2x).
// ============================================================
function addPriceListSheets() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  // 1) PRICE_LISTS
  const name1 = CONFIG.SHEETS.PRICE_LISTS;
  if (!ss.getSheetByName(name1)) {
    createSheet_(ss, name1, [
      'id_price_list', 'nama_paket', 'deskripsi',
      'aktif', 'tgl_dibuat', 'diupdate_oleh', 'tgl_update_terakhir'
    ]);
    Logger.log('âœ… Sheet ' + name1 + ' dibuat (7 kolom)');
  } else {
    Logger.log('â„¹ Sheet ' + name1 + ' sudah ada. Skip.');
  }

  // 2) PRICE_LIST_ITEMS
  const name2 = CONFIG.SHEETS.PRICE_LIST_ITEMS;
  if (!ss.getSheetByName(name2)) {
    createSheet_(ss, name2, [
      'id_price_list_item', 'id_price_list', 'nama_produk',
      'harga', 'tgl_dibuat', 'tgl_update_terakhir'
    ]);
    Logger.log('âœ… Sheet ' + name2 + ' dibuat (6 kolom)');
  } else {
    Logger.log('â„¹ Sheet ' + name2 + ' sudah ada. Skip.');
  }
}

function addResellerIdPriceListColumn() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sh = ss.getSheetByName(CONFIG.SHEETS.RESELLER);
  if (!sh) {
    Logger.log('âŒ Sheet RESELLER tidak ditemukan');
    return;
  }
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const alreadyHas = headers.some(function(h) { return String(h).toLowerCase() === 'id_price_list'; });
  if (alreadyHas) {
    Logger.log('â„¹ Kolom id_price_list sudah ada. Skip.');
    return;
  }
  const newCol = headers.length + 1;
  sh.getRange(1, newCol).setValue('id_price_list');
  Logger.log('âœ… Kolom id_price_list ditambahkan di kolom ke-' + newCol + ' (' + (headers.length + 1) + ')');
}

// ============================================================
// F9.7A - MANUAL: Setup 2 sheet Program + kolom PESANAN untuk program
// Jalankan sekali dari Apps Script editor. Idempotent (aman run 2x).
// ============================================================
function addProgramSheets() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  // 1) PROGRAMS
  const name1 = CONFIG.SHEETS.PROGRAMS;
  if (!ss.getSheetByName(name1)) {
    createSheet_(ss, name1, [
      'id_program', 'nama_program', 'deskripsi',
      'tgl_mulai', 'tgl_selesai', 'aktif',
      'mode_approval', 'visibility', 'kombinasi_mode', 'kalkulasi_mode', 'priority',
      'target_type', 'target_ids', 'basis_harga',
      'max_diskon_per_order', 'max_usage_per_customer', 'max_usage_total', 'current_usage',
      'tgl_dibuat', 'diupdate_oleh', 'tgl_update_terakhir'
    ]);
    Logger.log('✅ Sheet ' + name1 + ' dibuat (21 kolom)');
  } else {
    Logger.log('ℹ Sheet ' + name1 + ' sudah ada. Skip.');
  }

  // 2) PROGRAM_RULES
  const name2 = CONFIG.SHEETS.PROGRAM_RULES;
  if (!ss.getSheetByName(name2)) {
    createSheet_(ss, name2, [
      'id_rule', 'id_program', 'urutan',
      'rule_type', 'scope', 'scope_ids', 'parameter',
      'tgl_dibuat', 'tgl_update_terakhir'
    ]);
    Logger.log('✅ Sheet ' + name2 + ' dibuat (9 kolom)');
  } else {
    Logger.log('ℹ Sheet ' + name2 + ' sudah ada. Skip.');
  }
}

function addPesananProgramColumns() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sh = ss.getSheetByName(CONFIG.SHEETS.PESANAN);
  if (!sh) {
    Logger.log('❌ Sheet PESANAN tidak ditemukan');
    return;
  }
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const existingLower = headers.map(function(h) { return String(h).toLowerCase(); });
  const newCols = [
    'id_program', 'program_nama', 'program_snapshot',
    'approval_status', 'approved_by', 'approved_at', 'rejection_reason'
  ];
  let nextCol = headers.length + 1;
  const added = [];
  newCols.forEach(function(col) {
    if (existingLower.indexOf(col.toLowerCase()) === -1) {
      sh.getRange(1, nextCol).setValue(col);
      added.push(col);
      nextCol++;
    }
  });
  if (added.length === 0) {
    Logger.log('ℹ Semua kolom PESANAN sudah ada. Skip.');
  } else {
    Logger.log('✅ Kolom PESANAN ditambahkan: ' + added.join(', '));
  }
}

function testConnection() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  Logger.log('Terhubung ke : ' + ss.getName());
  Logger.log('URL          : ' + ss.getUrl());
  Logger.log('Jumlah sheet : ' + ss.getSheets().length);
  Logger.log('Daftar sheet : ' + ss.getSheets().map(function(s){return s.getName();}).join(', '));
}