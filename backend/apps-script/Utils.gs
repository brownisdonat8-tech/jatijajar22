/**
 * KK22 Database - Utils
 * Helper functions yang dipakai di seluruh kode.
 */

function getSpreadsheet_() {
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

function getSheet_(name) {
  const sh = getSpreadsheet_().getSheetByName(name);
  if (!sh) throw new Error('Sheet tidak ditemukan: ' + name);
  return sh;
}

/**
 * Normalisasi nomor HP ke format 62xxxxxxxxxx
 * Contoh: 0812xxx -> 62812xxx; +62 812-xxx -> 62812xxx
 */
function normalizePhone_(phone) {
  if (phone === null || phone === undefined) return '';
  let s = String(phone).replace(/[^0-9]/g, '');
  if (s === '') return '';
  if (s.indexOf('62') === 0) return s;
  if (s.indexOf('0') === 0)  return '62' + s.substring(1);
  if (s.indexOf('8') === 0)  return '62' + s;
  return s;
}

function nowJakarta_() {
  return Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyy-MM-dd HH:mm:ss');
}

function todayJakarta_() {
  return Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyy-MM-dd');
}

/**
 * Generate PIN hash (SHA-256 + salt) -> hex string
 */
function pinHash_(pin, salt) {
  const raw = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(pin) + String(salt),
    Utilities.Charset.UTF_8
  );
  let hex = '';
  for (let i = 0; i < raw.length; i++) {
    const b = (raw[i] & 0xFF).toString(16);
    hex += (b.length === 1 ? '0' + b : b);
  }
  return hex;
}

/**
 * Generate salt 16 karakter dari UUID
 */
function makeSalt_() {
  return Utilities.getUuid().replace(/-/g, '').substring(0, 16);
}

/**
 * Cari baris terakhir yang kolom [colIndex] tidak kosong
 */
function getLastDataRow_(sheet, colIndex) {
  const last = sheet.getLastRow();
  if (last < 2) return 1;
  const values = sheet.getRange(2, colIndex, last - 1, 1).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] !== '' && values[i][0] !== null) return i + 2;
  }
  return 1;
}