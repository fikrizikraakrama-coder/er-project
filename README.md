# er-project
System AR Develop — Sistem AR CSU (Google Apps Script).

## Struktur
- `src/` — berkas proyek Apps Script (salin apa adanya ke editor Apps Script). `PdfLib.html` tidak ada di repo ini.
- `src/Engine.gs` adalah **sumber tunggal** perhitungan. `src/EngineClient.html` dibuat darinya: `npm run sync-engine`.
- `test/` — uji ENGINE di luar Sheets: `npm test` (Node 18+, tanpa dependensi).

## Pemasangan v2.8.6
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

## v2.8.6: Document filing, file Pajak, Monitoring Invoice ringkas, menu rapi
- **Perbaikan Surat Penagihan**: tampilan surat kembali normal (nama gaya layar aplikasi v2.8.4 bertabrakan dengan gaya surat). Garis putus-putus yang ikut muncul di tabel hitungan invoice cetak juga hilang.
- **Document filing** (Invoice Register ▸ *Document filing…* atau tombol di lembar invoice): unggah **Signed billing pack** (invoice, kwitansi, BAP, sertifikat pembayaran ber-TTD; PDF/JPG/PNG maks. 10 MB) dan **Tax invoice (e-Faktur)** (PDF Coretax).
  Nomor & tanggal faktur dibaca otomatis dan dicek terhadap DPP/PPN invoice dan referensi nomor invoice. File disimpan di folder arsip ▸ *Berkas Tagihan & Faktur Pajak* ▸ tahun ▸ bulan,
  dinamai otomatis (mis. `4C.0183.26 - INV Prog 4 - Metland … Sign.Rev.00.pdf`), versi lama pindah ke *Revisi lama*. Link masuk kolom **BERKAS TAGIHAN (TTD)** & **FAKTUR PAJAK (PDF)**
  (sebelumnya DOKUMEN INVOICE / DOKUMEN FAKTUR PAJAK), Rekap Invoice Keluar & file Pajak. Pembaca PDF (pdf.js, cdnjs) dimuat hanya saat mengunggah faktur; bila gagal, nomor & tanggal bisa diketik.
- **File Pajak** "REKAP INVOICE KELUAR & BUKU PENJUALAN — CSU (PAJAK)": tab KELENGKAPAN DOKUMEN, REKAP tahun, BUK PEN per bulan (tanggal faktur). Dibuat lewat menu
  *Pengaturan ▸ Aktifkan file Pajak*, lalu diperbarui tiap jam bersama file Akuntansi & P&E. Email Pajak: Parameter ▸ Pengguna & Peran, peran **PAJAK (LIHAT)** (akses lihat otomatis).
- **Monitoring Invoice**: kolom rincian hitungan, dokumen & faktur, alokasi penerimaan dan kolom sistem dikelompokkan (tombol +/− di atas kolom), tidak dihapus. Dibuat otomatis sekali pada pembaruan laporan berikutnya.
- **Filter di sheet isian & log** (Monitoring Invoice, Log Penerimaan, Log Persetujuan, Log Surat Penagihan, master, dll.): setiap pembaruan laporan, rentang filter aktif dilebarkan sampai baris terakhir sheet
  (pilihan filter tetap), sehingga invoice/penerimaan baru dari aplikasi ikut tersaring. Tampilan filter (filter view) ikut dilebarkan bila layanan lanjutan **Google Sheets API** diaktifkan
  (Apps Script ▸ Layanan ▸ + ▸ Google Sheets API). Data & laporan selalu membaca semua baris — filter hanya memengaruhi tampilan.
- **Menu AR CSU**: menu harian di atas, sisanya di sub-menu (Analisis bulanan · Akuntansi, P&E & Pajak · Data proyek · Pengaturan). *Terapkan pembaruan v2.6* dan *Ekspor Buku Penjualan* dihapus.

## Aplikasi v2.8.5: Overview, form bertahap & bahasa Inggris
- **Bahasa aplikasi Inggris** (istilah finance: Overview, Billing, Cash Receipts, Month-End Close, Approvals; Pending approval, Resubmit required, …).
  **Dokumen cetak/PDF tetap Bahasa Indonesia**, sheet Google tidak berubah, angka tetap format Indonesia (5.224.132.582).
- **Overview** layar penuh (AR Officer & Section Head): pilihan bulan (bawaan bulan berjalan, otomatis berganti tiap awal bulan termasuk ganti tahun; 12 bulan ke belakang),
  6 angka utama (piutang terbuka, lewat jatuh tempo, invoice keluar, penerimaan vs target INKASO, pencapaian RA, DSO), invoice keluar bulan itu + status persetujuan,
  RA–RI Termin bulan itu + daftar RA yang belum cair, tren 12 bulan, umur piutang, daftar **Action needed** (AR Officer) / **Awaiting your decision** + **Attention this month** (Section Head, dengan tombol **Print overview**).
  Data dari `apiOverview(bulan)` (hitungan ringan + cache 5 menit, gugur setiap ada penyimpanan). Saldo bulan lalu = posisi akhir bulan.
- **Issue Invoice** & **Cash Receipts**: kartu bertahap 1·2·3, info proyek (kontrak, sisa UM, retensi, PPh), hitungan langsung (DPP, PPN, retensi, PPh, kas yang diharapkan),
  ringkasan **Received / Applied / Unapplied** (harus 0), tombol tetap **Save** dan **Save & submit for approval** (sekali klik: simpan lalu langsung diajukan ke Section Head).
- **Tampilan HP (iPhone, layar ≤ 700 px)**: menu di bawah layar, daftar approval berupa kartu, ringkasan angka pengajuan sebelum dokumen lengkap (*View full document*),
  tombol Reject / Approve menempel di bawah, Overview 2 kolom + cari pelanggan/invoice. Laptop & iPad tidak berubah. Di iPhone: buka link lewat Safari ▸ Share ▸ Add to Home Screen.
- Lembar pengesahan penerimaan: kolom *Dasar alokasi* dihapus. Pindah buku menampilkan nilai yang dipindahkan (bukan 0).

## Aplikasi v2.8.4: menu atas, Status Persetujuan & kecepatan
- Menu di bar atas per kelompok alur piutang (Ringkasan · Penagihan ▾ · Penerimaan Kas · Tutup Buku · Persetujuan); Section Head: Otorisasi · Dashboard · Tutup Buku.
- Status Persetujuan (Pembuat): Belum diajukan → Perlu diajukan ulang → Menunggu Section Head per jenis, dihitung sejak Parameter
  `MULAI_PERSETUJUAN` (bawaan 2026-10-01), tanpa data migrasi. Otorisasi (Section Head): urut terlama, umur pengajuan, Setujui & lanjut.
- Lembar pengesahan: kepala dokumen, tab lampiran (lembar pengesahan · kartu piutang · usulan jurnal), bilah keputusan dengan cek kelengkapan.
- Kecepatan: aplikasi & lembar pengesahan memakai hitungan ringan (tanpa laporan) + cache 5 menit yang gugur setiap ada penyimpanan;
  setelah simpan, laporan & file turunan diperbarui di latar belakang (`apiAfterSave`), jadwal 5 menit tetap sebagai cadangan.

## Proyek ARSIP (selesai sebelum sistem dipakai)
1. Master Proyek kolom **AKTIF = ARSIP** untuk proyek yang sudah selesai & lunas sebelum sistem dibangun → invoice, penerimaan dan
   tagbrut proyek itu langsung diabaikan seluruh sistem (laporan, Cek Data, aplikasi, Kartu Piutang; memorial lama dibiarkan).
2. Menu **AR CSU ▸ Pindahkan proyek ARSIP ke sheet arsip…** memindahkan barisnya dari Monitoring Invoice, Log Penerimaan & Input Tagbrut
   (kolom tagbrut saja) ke sheet **Arsip Data Lama** (hanya baca) lalu menghapusnya dari sheet kerja — tidak ada baris tersembunyi.
   Ringkasan ditampilkan dulu; bila data menunjukkan masih ada sisa tagihan, harus diketik PINDAHKAN. Tercatat di Jejak Edit Langsung + email Section Head.
3. **Kembalikan proyek dari arsip…** (per NPP) menaruh barisnya kembali di bawah sheet asal; lalu ubah AKTIF ke YA / SELESAI.

## Surat Penagihan (permohonan percepatan pembayaran)
Menu aplikasi **Surat Penagihan** (khusus PEMBUAT):
- Alert per pelanggan yang punya invoice lewat jatuh tempo (invoice tanpa jatuh tempo — belum submit / tanda terima — tidak ikut).
- Surat berjenjang otomatis: ke-1 Permohonan Percepatan Pembayaran → ke-2 Permohonan Pembayaran Kedua → ke-3 Penegasan Penyelesaian
  Pembayaran (merujuk nomor & tanggal surat sebelumnya). Surat berikutnya disarankan setelah tenggat lewat (tanpa tenggat: `SURAT_JEDA_HARI`,
  bawaan 7 hari) bila belum ada tanggapan dan invoice belum lunas. Setelah surat ke-3: alert **eskalasi ke manajemen**.
- Status surat: **SELESAI – LUNAS** / **SEBAGIAN DIBAYAR** otomatis dari penerimaan; **Catat tanggapan** (bentuk, ringkasan, janji bayar →
  surat berikutnya ditahan sampai janji); **Tutup surat** dengan alasan (sengketa, invoice dibatalkan/diganti).
  Peringatan dini: TENGGAT SURAT TERLEWAT, JANJI BAYAR TERLEWAT, ESKALASI PENAGIHAN.
- Revisi surat yang sudah disimpan: **Revisi (belum dikirim)** → nomor sama, tercatat Rev.N, PDF lama ke "Revisi lama";
  **Ganti (sudah diterima pelanggan)** → nomor + huruf (212 → 212A → 212B), surat lama berstatus DIGANTI (invoice tetap ditagih),
  surat pengganti tetap di tingkat yang sama dan memuat "Surat ini menggantikan surat kami Nomor … tanggal …". Keduanya wajib alasan
  dan dikirim ke email Section Head. Surat yang hanya dicetak (belum disimpan) cukup diubah lalu dicetak ulang.
- Satu surat per pelanggan; pelanggan multi-proyek → tabel ringkasan per proyek, rincian per invoice di lampiran.
  Isi mengalir ke halaman berikut bila panjang (kop di setiap halaman, judul tabel diulang, tanda tangan tidak terpotong).
- Diisi pengguna: nomor urut surat (format `SURAT_NOMOR_FORMAT`, bawaan `KU.03.03/WB-4C.{no}/{tahun}`), tanggal surat, Kepada Yth., tenggat (tanpa / 3 / 7 hari / tanggal), rekening (centang dari Parameter), kontak.
- Penandatangan: peran `MENGETAHUI` di Parameter (Manager CSU). Tembusan: `SURAT_TEMBUSAN` (dipisah `;`).
- **Simpan PDF ke Drive & catat surat** → folder arsip `Surat Penagihan/<tahun>/<bulan>` dan sheet **Log Surat Penagihan**.
  Tenggat yang terlewat dengan invoice masih terbuka muncul sebagai peringatan dini **TENGGAT SURAT TERLEWAT**.

## Arsip PDF
Jalankan sekali **AR CSU ▸ Siapkan folder arsip PDF** (pemilik) agar `ARSIP_FOLDER_ID` terisi dan semua akun mengarsip ke folder yang sama.

## Revisi invoice / kwitansi / BAP
Saat invoice yang sudah ada disimpan ulang (dengan alasan), sistem membandingkan data lama-baru dan menentukan jenisnya sendiri:
- **REVISI NILAI** (tanggal, DPP, PPN, UM, retensi, pelanggan, proyek …): Rev naik, ajukan ulang persetujuan, kena tutup buku, tanda faktur pengganti bila DPP/PPN berubah.
- **REVISI INVOICE/KWITANSI** (uraian, invoice ke-, penandatangan, isian form lain; nilai tetap): Rev naik, ajukan ulang persetujuan; tidak kena tutup buku.
- **REVISI DOKUMEN BAP** (tanggal BAP, progress, Nomor/Tanggal SPK di BAP): Rev & persetujuan tetap, cukup tercatat.
Semua tercatat di sheet Riwayat Revisi (kolom JENIS REVISI & PERUBAHAN: nilai lama → baru). Dokumen dicetak tanpa tanda revisi.
Nomor/Tanggal SPK di BAP bisa diisi per invoice (kosong = Master Proyek) untuk koreksi typo tanpa mengubah dokumen lain.

## Format tanggal di sheet
Menu **AR CSU ▸ Format tanggal Indonesia (dd/mm/yyyy)…**: lokal file Indonesia (05/10/2026 = 5 Oktober) dan validasi tanggal sah pada kolom tanggal manual. Cek Data menandai tanggal dokumen sebelum tanggal invoice atau di masa depan (kemungkinan tanggal & bulan tertukar).
