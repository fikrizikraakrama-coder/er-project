# er-project
System AR Develop — Sistem AR CSU (Google Apps Script).

## Struktur
- `src/` — berkas proyek Apps Script (salin apa adanya ke editor Apps Script). `PdfLib.html` tidak ada di repo ini.
- `src/Engine.gs` adalah **sumber tunggal** perhitungan. `src/EngineClient.html` dibuat darinya: `npm run sync-engine`.
- `test/` — uji ENGINE di luar Sheets: `npm test` (Node 18+, tanpa dependensi).

## Pemasangan v2.9.0 (sama dengan v2.8.7)
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

## v2.8.7: Kompensasi utang–piutang & PPh sesuai bukti potong
- **Cash Receipts ▸ Compensation** (jenis transaksi `KOMPENSASI` di Log Penerimaan): piutang dilunasi dengan utang perusahaan kepada pelanggan
  (mis. klaim pemakaian solar yang dipotong dari pembayaran progres). Isi tanggal, **nomor BA / memo** (jadi REF, wajib), invoice, nominal, uraian.
  - Mengurangi piutang dan tampil di kartu piutang sebagai *"Dikompensasi dengan utang ke pelanggan — …"*.
  - **Tidak dihitung uang masuk**: tidak masuk INKASO, RA-RI, Overview "received", maupun "Diterima (kas)" di kartu piutang.
  - Jurnal usulan & memorial: **Utang Usaha (kompensasi) D / Piutang Usaha K** (FB01 · SA). Akun `UTANG_KOMPENSASI` ditambahkan otomatis ke
    Parameter ▸ Pemetaan Akun SAP saat kompensasi pertama disimpan, dengan GL kosong — **isi GL & kode memorial dari Akuntansi**, lalu hapus kata "konfirmasi". Selama belum diisi, Cek Data mengingatkan.
  - Lembar pengesahan berjudul *LEMBAR PENGESAHAN KOMPENSASI PIUTANG* (lampiran: BA/memo, bukti potong, kartu piutang, jurnal FB01).
- Baris KOMPENSASI tidak pernah memikul PPh / adm: bila kompensasi yang menuntaskan invoice, PPh (dan nomor bupot) diakui pada **penerimaan kas terakhir**
  (jurnal F-28), kompensasi hanya Utang kompensasi / Piutang (FB01).
- **PPh sesuai bukti potong** (kolom baru `PPh DIPOTONG SESUAI BUPOT (OPSIONAL)` di Log Penerimaan, dibuat otomatis saat pertama dipakai): di aplikasi muncul saat *paid in full* dicentang
  (transfer) atau di tab Compensation. Bila diisi, PPh diakui **sebesar bupot** (bukan hitungan sistem); selisih terhadap hitungan sistem ditulis di keterangan kartu piutang.
  Kosongkan untuk invoice normal — deteksi otomatis tetap berjalan seperti sebelumnya.
- **ADJUSTMENT DPP (DASAR PPN SAJA)** — kolom isian baru di Monitoring Invoice (ditambahkan otomatis di ujung kanan pada pembaruan laporan berikutnya).
  Untuk tambahan DPP di faktur yang hanya menaikkan dasar PPN dan tidak dibayar (BAPP butir *Adjustment DPP*). DPP faktur, PPN, file Pajak & Buku Penjualan tetap;
  piutang, dasar PPh, status & kartu piutang memakai **DPP − adjustment** (sama dengan pencatatan Akuntansi di LK). Isi = DPP faktur − DPP tagihan BAPP.
  Mengisinya mengubah sidik invoice → **ajukan ulang** paket invoice ke Section Head.
- Contoh 4C.0141/2026 (WRK, progres 12): ADJUSTMENT DPP 365.569.857 (8.069.185.245 − 7.703.615.388) → piutang Netto 7.258.953.518 (= LK Sep-26), PPh 192.362.268 (= BAPP).
  Setor tunai 5.766.264.127 (transfer) + kompensasi klaim solar 2.187.937.500 → **LUNAS**, tanpa tutup manual.

- **Persetujuan lewat tanda tangan basah** (Parameter `PERSETUJUAN_TTD_BASAH`, bawaan `INV,PAY,KP,REK`; isi `TIDAK` untuk kembali ke persetujuan di aplikasi):
  pengajuan invoice, penerimaan, kartu piutang & rekap bulanan langsung tercatat **DISETUJUI** saat Submit. Log Persetujuan: PUTUS OLEH kosong,
  CATATAN `[TTD BASAH] …` — tidak ada cap elektronik atas nama Section Head; bukti persetujuan = tanda tangan basah pada lembar cetak.
  Validasi rencana cair, pembatalan invoice, penutupan manual & pembatalan lembar validasi tetap diputuskan Section Head di aplikasi
  (bisa ditambahkan: `TUTUP`, `BATAL`). Pengajuan lama yang masih menunggu: tombol *Mark … pending as approved (wet signature)* di Approvals
  atau menu *Pengaturan ▸ Sahkan pengajuan menunggu (TTD basah)* (yang datanya berubah dilewati).
- Tanggal di kotak tanda tangan (Dibuat, Diperiksa & disetujui, Disetujui, Diterima Akuntansi, "Jakarta, …" di kartu piutang) terisi tanggal pengajuan.
- Cetak kartu piutang lebih ringkas; kartu & jurnal landscape tercetak selebar halaman setelah lembar portrait.

- **Menu baru v2.9.0 (mengikuti alur AR SAP, kode T-code tampil kecil sebagai petunjuk)**: Overview · **Customers** (Customer Master `BP`, Customer Balances `FD10N`) ·
  **Invoices** (Create Invoice `FB70`, Invoice List `FBL5N`, Cancel Invoice `FB08`) · **Payments** (Post Incoming Payment `F-28`, Payment List, Reverse Payment `FBRA`) ·
  **Collections** (Collection Plan, Dunning Letters `F150`, Credit Limits `UKM_BP`) · **Month-End** (AR Closing Checklist, Documents & Handover) · Approvals.
- **Overview baru (v2.9.0)**: sapaan & sisa hari kerja, 5 kartu KPI dengan grafik mini (saldo piutang + tren, overdue + cincin, uang masuk vs RA, invoice terbit, DSO),
  panel RA–RI ringkas (donat + 2 bar, sisa yang harus ditagih, **kebutuhan per hari kerja**, item RA yang sudah cair), **Action needed** (hanya yang perlu dikerjakan,
  urut dari yang paling penting, langsung ke menunya), tren 12 bulan (invoice vs uang masuk + garis saldo piutang), umur piutang berwarna, pelanggan terbesar,
  rencana cair yang belum masuk ("x d late / in x days") dan invoice terbit bulan ini. Tabel invoice yang dulu berantakan sudah diperbaiki.
- **Customer Balances (v2.9.0)**: saldo per pelanggan per umur (belum ada JT, belum jatuh tempo, 1–30, 31–90, > 90 hari) + penerimaan terakhir; klik pelanggan →
  proyek & tombol kartu piutang. Bisa dicetak.
- **Credit Limits (v2.9.0)**: kolom baru **BATAS KREDIT (Rp, OPSIONAL)** di Master Pelanggan (dibuat otomatis; diisi lewat Customer Master, perubahan wajib dasar &
  diemailkan ke Section Head). Layar membandingkan batas vs exposure (saldo invoice aktif) dengan status OVER LIMIT / ABOVE 80% / OK. Create Invoice memberi
  **peringatan** bila invoice baru membuat pelanggan melewati batas (boleh disimpan dengan alasan; tercetak di lembar pengesahan).
- **Reverse Payment (v2.9.0)**: penerimaan yang salah dibalik lewat aplikasi (Payments ▸ Reverse Payment atau tombol ↺ Reverse di Payment List) dengan alasan
  (salah invoice / salah nominal / input ganda / transfer dikembalikan / lainnya). Baris Log Penerimaan **tidak dihapus**: setelah disahkan (TTD basah seperti
  penerimaan, jenis `PEMBALIKAN PENERIMAAN` di Log Persetujuan) baris ditandai **↺ DIBALIK** di kolom CEK, tidak dihitung lagi, invoice kembali terbuka,
  memorial Akuntansi otomatis dibatalkan / dikoreksi, dan bila penerimaan itu sudah diserahkan ke Akuntansi, pembalikannya masuk lembar serah terima
  penerimaan berikutnya agar posting SAP dibalik (FB08). Lembar Pembalikan Penerimaan bisa dicetak & diarsipkan. Penerimaan yang benar lalu dicatat ulang.
- **Cancel Invoice** kini juga ada di menu Invoices (pilih invoice dari daftar → lembar pembatalan seperti sebelumnya).
- **Issue Invoice bertahap (v2.8.8)**: 1 Customer & project · 2 Amount · 3 Wording · 4 Preview & save, dengan ringkasan angka yang selalu
  terlihat (progres − UM = DPP + PPN − retensi − PPh = kas diterima). Langkah 3 = **pola uraian per proyek** dengan isian otomatis
  `{jenis} {progres} {termin} {terminR} {deskripsi} {bapp} {tglbapp} {kumulatif}` (tombol sisip; "Save as this project's pattern" menyimpan ke
  kolom baru POLA URAIAN INVOICE di Master Proyek, dibuat otomatis) dan **uraian kwitansi terpisah** bila perlu ("Untuk pembayaran …").
  No./tgl BAPP bisa diisi di langkah 2. Langkah 4 = ringkasan & pemeriksaan sebelum Save / Save & submit.
- **Master Data (v2.8.8)**: pelanggan, proyek & adendum diisi lewat aplikasi (tersimpan di Master Pelanggan / Master Proyek / Riwayat Adendum).
  Daftar dengan pencarian & saringan (berelasi / pihak ketiga / data belum lengkap), detail per pelanggan (profil, proyek, riwayat) dan per proyek (kontrak &
  syarat penagihan, adendum, riwayat). Pemeriksaan otomatis (NPWP 16 digit & ganda, nama mirip, NPP ganda, retensi/UM/tarif PPh wajar…). Kolom **sensitif**
  (nilai kontrak, UM, retensi, PPh, template, kelompok, jenis, NPWP) wajib diisi **dasar perubahan** — dicatat di sheet **Riwayat Master Data** dan diemailkan ke
  Section Head; adendum selalu wajib dasar. NPP & nama pelanggan dikunci (ganti NPP lewat menu). **+ New customer** = wizard 3 langkah (pelanggan →
  proyek pertama → uraian invoice & cek) yang menyimpan keduanya sekaligus.
- **Month-End Close ▸ AR Closing Checklist (v2.8.8)**: 10 langkah tutup buku piutang diperiksa otomatis dari data — cut-off penagihan, nota kredit &
  penyesuaian, penerapan penerimaan (termasuk penerimaan tanpa tanggal), rekonsiliasi per pelanggan (invoice lebih bayar / PPN ≠ 11% × DPP = merah),
  umur piutang, sengketa & kurang bayar, CKPN, AR vs GL (Data LK), serah terima ke Akuntansi, laporan akhir & tutup buku. Cincin kesiapan, alur berwarna,
  laporan akhir AR, tombol menuju perbaikan, **cetak checklist** untuk ditandatangani, dan tombol **Close <bulan>** yang hanya aktif bila tidak ada langkah merah.
- **Pemeriksaan input (v2.8.8)**: form Issue Invoice & Cash Receipts menampilkan panel *Checks* — ⛔ error (simpan terkunci: progres kumulatif
  > 100% kontrak, tanggal kosong/masa depan, retensi melebihi sisa, nilai 0, bayar sebelum tanggal invoice…), ⚠ peringatan (nomor progres melompat/ganda,
  PPN ≠ 11% × DPP, uraian masih "…", invoice jadi lebih bayar, ref bank kosong…) yang boleh disimpan dengan **alasan** (tercatat di DATA FORM / KETERANGAN dan
  tercetak di lembar pengesahan), ℹ info. Kolom bermasalah diberi bingkai merah/kuning dan tombol "Use …" mengisi nilai yang disarankan.
  Data di sheet juga diperiksa: invoice yang PPN-nya tidak cocok dengan DPP (mis. DPP sudah dikurangi retensi/adjustment) muncul di Cek Data (KONSISTENSI).
- **Menu Collection Plan** (Rencana Inkaso, layar penuh): pilih bulan, alur status (Draft → Submitted → Rejected/Approved), banner penolakan
  beserta alasannya, ringkasan ✓/◇/⏸/→ dengan bar terhadap target inkaso, cari & saring per kelompok (tunggakan, jatuh tempo bulan ini, …),
  centang beberapa invoice → *Set selected to*, Quick fill, bar bawah berisi total & tombol Save draft / Submit. Data & pengajuan sama dengan
  Lembar Validasi sebelumnya; keputusan Section Head tetap di Approvals. Bulan yang memakai RA Manual menampilkan keterangannya.
- **Serah terima ke Akuntansi dipisah per jenis**: tab *Invoices issued* / *Cash receipts*, nomor **STD-INV-yyyy-mm-###** dan **STD-PNR-yyyy-mm-###**,
  lembar cetak berbeda (invoice: No. invoice, uraian, DPP, No. faktur pajak · penerimaan: No. PNR, tgl uang masuk, jenis, rekening/ref, nilai) dengan
  pratinjau langsung saat mencentang. Lembar lama STD-yyyy-mm-### tetap bisa dicetak ulang.
- **Peran MANAJEMEN (LIHAT)** (Parameter ▸ Pengguna & Peran): link aplikasi sama, akses dicek dari email; menu Executive Summary · Overview ·
  Collection Plan · Reports, semua hanya baca (server menolak setiap simpan/ajukan/setujui dari akun ini). Executive Summary: 6 angka utama, tren
  tagihan vs uang masuk 12 bulan, umur piutang, top pelanggan, progres tagihan proyek, peringatan otomatis, tata kelola; bisa dicetak/PDF dan nyaman
  di HP. Link pintas `…/exec?v=manajemen` membuka Executive Summary (juga tersedia di menu Section Head).
- **Sheet RA Manual** (dibuat otomatis saat *Perbarui semua laporan*, terisi 20 baris RA TERMIN OKTOBER 2026): bila suatu bulan ada di sheet ini, daftar
  itulah RA resmi bulan itu — **RA-RI Termin, Pareto, Overview & Kinerja** mengikuti; Lembar Validasi aplikasi bulan itu diabaikan. NILAI RA = DPP − PPh
  (cicilan: isi nilai cicilan yang direncanakan), NO INVOICE boleh singkat (0085/2026). RI tetap dari Log Penerimaan (penerimaan yang sudah dicatat langsung
  terhitung); invoice yang cair tetapi tidak ada di daftar masuk **DI LUAR RA**. Mulai November 2026 RA kembali diajukan lewat aplikasi.
- **Cash Receipts ▸ Receipts Register** (menu baru): daftar penerimaan dipindah dari halaman input ke halaman sendiri — cari (no. PNR / invoice /
  pelanggan / ref bank, semua bulan), filter bulan & jenis (transfer, pindah buku/kompensasi, belum disahkan, belum diserahkan ke Akuntansi), klik baris →
  lembar pengesahan, tombol **Print** mencetak **Register Penerimaan Piutang** per bulan. Halaman *Record Receipt* kini hanya form input.
- **Tanda terima ke Akuntansi: invoice dicatat sebesar DPP (tanpa PPN)** — kolom *DPP Invoice* dan *Penerimaan* dipisah, masing-masing dengan total;
  lembar lama yang dicetak ulang juga menampilkan DPP.
- **Serah terima dokumen ke Akuntansi** (Month-End Close ▸ *Handover to Accounting*): daftar invoice & penerimaan yang sudah **disahkan** tetapi belum diserahkan
  (dengan penanda No. Dok SAP & verifikasi memorial). Pilih → *Create handover sheet* → cetak **Tanda Terima Dokumen ke Akuntansi** (STD-yyyy-mm-###), Akuntansi
  menandatangani & mencentang setiap dokumen. Tercatat di sheet **Serah Terima Akuntansi** (dibuat otomatis) dan kolom **SERAH TERIMA AKUNTANSI** di Monitoring Invoice &
  Log Penerimaan. Dokumen yang disahkan ulang setelah diserahkan muncul lagi ("serahkan lagi"). Belum diserahkan > 3 hari → Cek Data. Salah cetak: *Cancel* (baris tetap, STATUS BATAL).
  Alur pemeriksaan per dokumen: Disahkan → Diserahkan → VERIF memorial Akuntansi → No. Dok SAP → Rekonsiliasi Saldo.

- **Posisi kartu piutang selalu hari ini**: Kartu Piutang (E5 kosong = hari ini) dan file P&E (HARI INI / IKUT RINGKASAN) mengikuti tanggal terbaru.
  Tanggal pilihan manual hanya sementara — setiap pagi 06.00 dikembalikan ke hari ini. Selama posisi bukan hari ini, muncul baris merah
  "⚠ POSISI PER … — BUKAN HARI INI" beserta jumlah penerimaan / invoice setelah tanggal itu. Status P&E tidak lagi tertahan di "MEMPERBARUI…".

- **Sheet baru Rekap Tagihan Proyek** (dibuat otomatis saat pembaruan laporan): per proyek nilai kontrak setelah adendum, sudah ditagih (nilai progres),
  % tertagih, belum ditagih & status; ringkasan total; diagram *sudah vs belum ditagih (Rp miliar)* dan *% tertagih per proyek*.
- Sinkron file Akuntansi / P&E / Pajak: tetap dikirim walau sistem sibuk atau ada laporan yang gagal; proyek dari simpan yang bertabrakan dengan sinkron lain dititipkan dan dikirim sinkron berikutnya.
- Pembaruan laporan: satu laporan yang gagal tidak lagi menghentikan laporan lain; pesan "Reports not fully updated" di aplikasi menampilkan penyebabnya.

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
