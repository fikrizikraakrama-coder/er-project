# er-project
System AR Develop — Sistem AR CSU (Google Apps Script).

## Struktur
- `src/` — berkas proyek Apps Script (salin apa adanya ke editor Apps Script). `PdfLib.html` tidak ada di repo ini.
- `src/Engine.gs` adalah **sumber tunggal** perhitungan. `src/EngineClient.html` dibuat darinya: `npm run sync-engine`.
- `test/` — uji ENGINE di luar Sheets: `npm test` (Node 18+, tanpa dependensi).

## Pemasangan v2.8.1
1. Salin semua berkas `src/` ke proyek Apps Script, lalu **Terapkan ▸ Kelola deployment ▸ Versi baru**.
   Aplikasi menampilkan peringatan merah bila versi Engine server dan aplikasi berbeda.
2. Setiap pengguna akan diminta izin baru (**kirim email atas nama Anda**) untuk notifikasi keputusan.
   Bila `appsscript.json` memuat `oauthScopes`, tambahkan `https://www.googleapis.com/auth/script.send_mail`.
3. Pemilik spreadsheet menjalankan **AR CSU ▸ Aktifkan otomatisasi** sekali (menambah trigger 5 menit `rebuildIfDirty`).
4. Parameter ▸ Pengguna & Peran: pastikan email PEMBUAT dan PENYETUJU benar (sebaiknya akun kantor). Tidak ada akun yang ditulis di kode.
5. Parameter (kolom A–C) yang baru:
   | Kunci | Isi |
   |---|---|
   | `TUTUP_BUKU_TANGGAL` | Bawaan `3`: bulan lalu otomatis tutup buku tanggal 3 pukul 00.00 WIB (laporan ke pusat). Setelah itu invoice, penerimaan, revisi & pembatalan bertanggal bulan itu ditolak. |
   | `TUTUP_BUKU_OTOMATIS` | Isi `TIDAK` hanya bila tutup otomatis ingin dimatikan. |
   | `PERIODE_DITUTUP_SAMPAI` | Opsional, lewat menu/aplikasi **Tutup periode**: menutup lebih awal (tidak bisa memundurkan tutup otomatis). |
   | `PERIODE_KOREKSI_SAMPAI`, `_OLEH`, `_ALASAN` | Diisi Section Head di aplikasi (Otorisasi Dokumen ▸ Izinkan koreksi), maks. 7 hari. |
   | `NOTIF_PERSETUJUAN_CC` | Email tambahan penerima notifikasi keputusan (opsional, pisahkan dengan koma). |

## Tutup buku & koreksi
- Snapshot Analisis Bulanan dibuat otomatis setelah tutup buku (jadwal harian 06.00 pada tanggal tutup buku).
- Koreksi periode tertutup: Section Head membuka izin (maks. 7 hari) di Otorisasi Dokumen AR → pembuat menyimpan dengan alasan →
  data ditandai `[KOREKSI PERIODE TERTUTUP …]`, email ke Section Head, dan snapshot bulan itu otomatis direvisi (Rev.01, Rev.02, …; arsip lama tetap).
