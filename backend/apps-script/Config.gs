/**
 * KK22 Database - Config
 * Semua konfigurasi global di sini.
 * JANGAN commit file ini kalau ada data sensitif (nanti kita pisah).
 */

const CONFIG = {
  // === SPREADSHEET ===
  SPREADSHEET_ID: '1cBHDpm1pJWbmxbxExBpTTW4asYgyoTGx5XbITWXvD7M',

  // === NAMA SHEET ===
  SHEETS: {
    USERS:          'USERS',
    RESELLER:       'RESELLER',
    PRODUK:         'PRODUK',
    STOK_LOG:       'STOK_LOG',
    PESANAN:        'PESANAN',
    DETAIL_PESANAN: 'DETAIL_PESANAN',
    LOG_AKTIVITAS:  'LOG_AKTIVITAS',
    CONFIG:         'CONFIG',
    JURNAL_MAPPING: 'JURNAL_MAPPING',
    RESELLER_HARGA: 'RESELLER_HARGA',
    PRICE_LISTS:      'PRICE_LISTS',
    PRICE_LIST_ITEMS: 'PRICE_LIST_ITEMS'
  },

  // === PREFIX ID ===
  PREFIX: {
    USER:     'KK22-USR-',
    RESELLER: 'KK22-RSL-',
    PRODUK:   'KK22-PRD-',
    PESANAN:  'KK22-PSN-',
    ORDER:    'KK22-ORD-'
  },

  // === GLOBAL ===
  TZ:       'Asia/Jakarta',
  CURRENCY: 'IDR',

  // === PIN ===
  PIN: { MIN: 4, MAX: 6 }
};