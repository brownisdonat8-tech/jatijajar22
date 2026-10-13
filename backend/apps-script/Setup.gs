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
function testConnection() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  Logger.log('Terhubung ke : ' + ss.getName());
  Logger.log('URL          : ' + ss.getUrl());
  Logger.log('Jumlah sheet : ' + ss.getSheets().length);
  Logger.log('Daftar sheet : ' + ss.getSheets().map(function(s){return s.getName();}).join(', '));
}