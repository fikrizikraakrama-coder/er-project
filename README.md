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

## Jejak edit langsung
Invoice & penerimaan dicatat lewat aplikasi. Ketikan / tempel manual pada kolom terkontrol Monitoring Invoice (kolom isian aplikasi)
dan Log Penerimaan (semua kecuali KET, NO. DOK SAP, BUKTI POTONG) dicatat di sheet **Jejak Edit Langsung** dan dikirim ke email Section Head.
Baris yang hilang (dihapus / ID diubah) terdeteksi pada pembaruan berikutnya. Perlu trigger `onChangeInstalled` (menu **Aktifkan otomatisasi**).

## Surat Penagihan (permohonan percepatan pembayaran)
Menu aplikasi **Surat Penagihan** (khusus PEMBUAT):
- Alert per pelanggan yang punya invoice lewat jatuh tempo (invoice tanpa jatuh tempo — belum submit / tanda terima — tidak ikut).
- Surat berjenjang otomatis: ke-1 Permohonan Percepatan Pembayaran → ke-2 Permohonan Pembayaran Kedua → ke-3 Penegasan Penyelesaian
  Pembayaran (merujuk nomor & tanggal surat sebelumnya). Surat berikutnya disarankan setelah tenggat lewat (tanpa tenggat: `SURAT_JEDA_HARI`,
  bawaan 7 hari) bila belum ada tanggapan dan invoice belum lunas. Setelah surat ke-3: alert **eskalasi ke manajemen**.
- Status surat: **SELESAI – LUNAS** / **SEBAGIAN DIBAYAR** otomatis dari penerimaan; **Catat tanggapan** (bentuk, ringkasan, janji bayar →
  surat berikutnya ditahan sampai janji); **Tutup surat** dengan alasan (sengketa, invoice dibatalkan/diganti).
  Peringatan dini: TENGGAT SURAT TERLEWAT, JANJI BAYAR TERLEWAT, ESKALASI PENAGIHAN.
- Satu surat per pelanggan; pelanggan multi-proyek → tabel ringkasan per proyek, rincian per invoice di lampiran.
- Diisi pengguna: nomor surat, tanggal surat, Kepada Yth., tenggat (tanpa / 3 / 7 hari / tanggal), rekening (centang dari Parameter), kontak.
- Penandatangan: peran `MENGETAHUI` di Parameter (Manager CSU). Tembusan: `SURAT_TEMBUSAN` (dipisah `;`).
- **Simpan PDF ke Drive & catat surat** → folder arsip `Surat Penagihan/<tahun>/<bulan>` dan sheet **Log Surat Penagihan**.
  Tenggat yang terlewat dengan invoice masih terbuka muncul sebagai peringatan dini **TENGGAT SURAT TERLEWAT**.
