# Roadmap Kunang-Kunang 22

| Fase | Isi | Status |
|------|-----|--------|
| 0 | Cek PC + struktur folder + git init | Sedang jalan |
| 1 | Desain schema Sheets + bikin Spreadsheet | Belum |
| 2 | Apps Script: Config, Utils, Auth | Belum |
| 3 | API Produk + Stok | Belum |
| 4 | API Pesanan | Belum |
| 5 | GitHub repo + Pages + frontend login | Belum |
| 6 | Frontend reseller: katalog + pesan | Belum |
| 7 | Frontend reseller: riwayat + status | Belum |
| 8 | Dashboard admin: pesanan | Belum |
| 9 | Dashboard admin: produk, stok, reseller | Belum |
| 10 | Dashboard owner: laporan | Belum |
| 11 | Notifikasi Telegram + WA fallback | Belum |
| 12 | Export Jurnal Mekari + XLSX | Belum |
| 13 | Auto-backup + PWA + offline | Belum |
| 14 | User manual + BACKUP_ARSITEKTUR.txt | Belum |
| 15 | Uji end-to-end | Belum |

## Keputusan Desain Penting
- Login   : No HP + PIN (SHA-256 + salt)
- Role    : owner (read-only), admin (CRUD), reseller (pesanan sendiri)
- Stok    : real-time locking pakai LockService
- Harga   : disnapshot saat pesan (histori terjaga)
- Config  : semua ID/URL/kunci di Config.gs (gampang pindah server)
- Adapter : kontrak JSON bersih, siap migrasi ke server lokal