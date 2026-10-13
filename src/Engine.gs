/**
 * ENGINE v2 — perhitungan inti Sistem AR Concrete Services Unit, PT Wijaya Karya Beton Tbk.
 * Murni JavaScript (tanpa SpreadsheetApp) supaya bisa diuji di luar Google Sheets dan dipakai juga oleh aplikasi (EngineClient).
 * Sumber data: Master Pelanggan, Master Proyek, Riwayat Adendum, Parameter, Monitoring Invoice, Log Penerimaan,
 * Input Tagbrut, Data LK (Tempel), Log Persetujuan. Semua sheet laporan dihitung ulang dari sumber ini.
 * Versi 2 (Oktober 2026): alokasi Netto/PPN baru, kartu piutang buku besar, faktur pajak kode 04, usulan jurnal SAP,
 * CKPN (matriks umur + NPV), monitoring tagbrut, rekon ke LK, persetujuan digital.
 * Dirancang dan dikembangkan oleh Fikri Zikra Akrama (AR Officer, Concrete Services Unit), 2026.
 * Penggandaan sistem, kode, maupun format laporan tanpa izin pengembang tidak diperkenankan.
 */
var ENGINE = (function () {
  'use strict';
  var VERSION = '2.8.1 · 13-10-2026';

  // ------------------------------------------------------------------ SKEMA SHEET
  // [kunci, JUDUL KOLOM, tipe, lebar, format]; tipe: app = diisi aplikasi, manual = diisi pengguna (kuning), auto = dihitung mesin (hijau)
  // Semua sheet dibaca BERDASARKAN JUDUL KOLOM, jadi urutan kolom boleh berbeda (file lama tetap terbaca).
  var REG = {
    sheet: 'Monitoring Invoice', headerRow: 7, firstRow: 8,
    cols: [
      ['no', 'NO', 'auto', 5], ['pelanggan', 'NAMA PELANGGAN', 'app', 30], ['npp', 'NPP', 'app', 14], ['proyek', 'NAMA PROYEK', 'auto', 30],
      ['tgl', 'TANGGAL INVOICE', 'app', 12, 'date'], ['nomor', 'NOMOR INVOICE', 'app', 26], ['jenis', 'JENIS TAGIHAN', 'app', 12],
      ['invke', 'INVOICE KE-', 'app', 16], ['template', 'TEMPLATE', 'app', 12],
      ['bruto', 'NILAI PROGRES / BRUTO', 'app', 16, 'money'], ['umk', 'PENGEMBALIAN UM', 'app', 15, 'money'],
      ['retensi', 'RETENSI', 'app', 14, 'money'], ['retpos', 'POSISI RETENSI', 'app', 11],
      ['dpp', 'DPP (FAKTUR)', 'app', 16, 'money'], ['ppn', 'PPN', 'app', 15, 'money'], ['total', 'TOTAL INVOICE (DPP + PPN)', 'auto', 16, 'money'],
      ['netto', 'NETTO TERCATAT', 'auto', 16, 'money'], ['tarif', 'TARIF PPh', 'app', 8, 'pct'], ['dasar', 'DASAR PPh', 'app', 14],
      ['pph', 'PPh', 'auto', 14, 'money'], ['tagih', 'NILAI DI INVOICE (DIBAYAR)', 'app', 16, 'money'],
      ['kas', 'SEHARUSNYA DITERIMA (KAS)', 'auto', 16, 'money'], ['ttd', 'PENANDATANGAN', 'app', 20], ['rev', 'REVISI KE', 'app', 7, 'int'],
      ['uraian', 'URAIAN TAGIHAN (UNTUK FAKTUR)', 'app', 40],
      ['status_inv', 'STATUS DOKUMEN (AKTIF / BATAL)', 'auto', 16], ['status', 'STATUS PEMBAYARAN', 'auto', 26], ['submit', 'TGL SUBMIT DOKUMEN', 'manual', 12, 'date'],
      ['tt', 'TGL TANDA TERIMA (KHUSUS FAJARPUTERA)', 'manual', 13, 'date'], ['top', 'TOP (HARI)', 'manual', 7, 'int'],
      ['jt', 'TGL JATUH TEMPO', 'auto', 12, 'date'], ['jt_akt', 'JT AKUNTANSI (TGL INVOICE + 30)', 'auto', 12, 'date'],
      ['statdok', 'STATUS DOKUMEN TAGIHAN', 'manual', 20],
      ['no_fp', 'NOMOR FAKTUR PAJAK', 'manual', 19], ['tgl_fp', 'TANGGAL FAKTUR PAJAK', 'manual', 12, 'date'],
      ['link_inv', 'DOKUMEN INVOICE', 'manual', 16], ['link_fp', 'DOKUMEN FAKTUR PAJAK', 'manual', 16],
      ['sap_doc', 'NO. DOK SAP (INVOICE)', 'manual', 13],
      ['validasi', 'USULAN RENCANA CAIR', 'manual', 18], ['rencana', 'RENCANA TGL CAIR', 'manual', 12, 'date'],
      ['validasi_ok', 'VALIDASI SECTION HEAD', 'auto', 22],
      ['catatan', 'CATATAN AR (OPSIONAL)', 'manual', 30], ['tutup', 'TUTUP MANUAL', 'manual', 13], ['alasan_tutup', 'ALASAN TUTUP MANUAL', 'manual', 30],
      ['ba_no', 'NO. BA / MEMO PENAGIHAN KEKURANGAN', 'manual', 24], ['ba_tgl', 'TGL BA / MEMO', 'manual', 12, 'date'],
      ['ke_ppn', 'DITERIMA → PPN', 'auto', 15, 'money'], ['ke_net', 'DITERIMA → NETTO', 'auto', 16, 'money'],
      ['diterima', 'TOTAL DITERIMA', 'auto', 16, 'money'], ['pot_pph', 'POT. PPh', 'auto', 14, 'money'], ['adm', 'ADM BANK', 'auto', 10, 'money'],
      ['sel_pph', 'SELISIH PPh', 'auto', 13, 'money'], ['sisa', 'SISA PIUTANG (KAS)', 'auto', 16, 'money'],
      ['tgl_bayar', 'TGL BAYAR TERAKHIR', 'auto', 12, 'date'], ['umur', 'UMUR (HARI)', 'auto', 7, 'int'],
      ['persetujuan', 'PERSETUJUAN PAKET INVOICE', 'auto', 24], ['link_paket', 'ARSIP PAKET (LINK)', 'app', 18],
      ['sumber', 'SUMBER DATA', 'app', 18], ['dibuat', 'DIBUAT / DIUBAH', 'app', 17], ['json', 'DATA FORM (JANGAN DIUBAH)', 'app', 10],
      ['memorial', 'MEMORIAL AKUNTANSI', 'auto', 22]
    ]
  };
  var REG_ALIAS = { validasi: ['VALIDASI MANAJER'], status_inv: ['STATUS INVOICE'], catatan: ['CATATAN / PENJELASAN'] };
  var LOG = {
    sheet: 'Log Penerimaan', headerRow: 5, firstRow: 6,
    cols: [
      ['no', 'NO', 'auto', 5], ['idp', 'ID PENERIMAAN', 'auto', 11], ['tgl', 'TANGGAL BAYAR', 'manual', 12, 'date'], ['ref', 'REF. TRANSFER / MUTASI', 'manual', 18],
      ['transfer', 'NOMINAL TRANSFER (BILA DIPECAH)', 'manual', 16, 'money'], ['inv', 'NOMOR INVOICE', 'manual', 26],
      ['nominal', 'NOMINAL UNTUK INVOICE INI', 'manual', 16, 'money'], ['alokasi', 'ALOKASI (KOSONG = OTOMATIS)', 'manual', 12],
      ['lunas', 'PELUNASAN TERAKHIR? (✓)', 'manual', 11], ['ket', 'KETERANGAN', 'manual', 30], ['sap_doc', 'NO. DOK SAP (PENERIMAAN)', 'manual', 13],
      ['rekening', 'REKENING PENERIMAAN', 'manual', 22], ['jenis_trx', 'JENIS TRANSAKSI (KOSONG = TRANSFER)', 'manual', 14], ['bupot', 'NO. BUKTI POTONG PPh (OPSIONAL)', 'manual', 18],
      ['pelanggan', 'NAMA PELANGGAN', 'auto', 28], ['npp', 'NPP', 'auto', 13],
      ['ke_ppn', '→ PPN', 'auto', 15, 'money'], ['ke_net', '→ NETTO', 'auto', 16, 'money'], ['aturan', 'DASAR ALOKASI', 'auto', 30],
      ['pph_cat', 'POT. PPh DIAKUI', 'auto', 14, 'money'], ['adm_cat', 'ADM BANK DIAKUI', 'auto', 11, 'money'],
      ['cicil', 'PEMBAYARAN KE-', 'auto', 9, 'int'], ['persetujuan', 'PERSETUJUAN PAKET', 'auto', 22], ['cek', 'CEK', 'auto', 34],
      ['memorial', 'MEMORIAL AKUNTANSI', 'auto', 22],
      ['request_id', 'ID PERMINTAAN (SISTEM)', 'app', 20], ['request_hash', 'SIDIK PERMINTAAN (SISTEM)', 'app', 20],
      ['request_user', 'PENCATAT (SISTEM)', 'app', 26]
    ]
  };
  var MPEL = {
    sheet: 'Master Pelanggan', headerRow: 2, firstRow: 3,
    cols: [['nama', 'NAMA PELANGGAN (KAPITAL)', 30], ['inv', 'NAMA DI INVOICE (KEPADA YTH)', 30], ['npwp', 'NPWP', 18],
      ['nitku', 'NITKU PEMBELI (OPSIONAL)', 24], ['kode_bp', 'KODE BUSINESS PARTNER (SAP)', 14],
      ['a1', 'ALAMAT BARIS 1', 34], ['a2', 'ALAMAT BARIS 2', 34], ['a3', 'ALAMAT BARIS 3', 30], ['jenis', 'JENIS PELANGGAN', 11],
      ['kelompok', 'KELOMPOK', 15], ['dasar_jt', 'DASAR JATUH TEMPO', 17], ['adm', 'BATAS ADM BANK (Rp)', 12], ['top_ret', 'TOP RETENSI (HARI)', 11],
      ['wakil', 'NAMA WAKIL (UNTUK BAP)', 20], ['jab', 'JABATAN WAKIL (UNTUK BAP)', 20], ['aktif', 'AKTIF', 7], ['catatan', 'CATATAN', 30]]
  };
  var MPRO = {
    sheet: 'Master Proyek', headerRow: 2, firstRow: 3,
    cols: [['npp', 'NPP', 14], ['pelanggan', 'NAMA PELANGGAN', 30], ['singkat', 'NAMA PROYEK (SINGKAT)', 30], ['kode', 'JENIS PEKERJAAN', 12],
      ['deskripsi', 'DESKRIPSI DI INVOICE ("Atas ...")', 50], ['pekerjaan', 'NAMA PEKERJAAN (BAP / KARTU PIUTANG)', 40], ['lokasi', 'LOKASI', 24],
      ['kontrak', 'NO. KONTRAK / SPK', 26], ['tgl_kontrak', 'TGL KONTRAK / SPK', 12], ['nilai', 'NILAI KONTRAK AWAL (EXCL. PPN)', 17],
      ['add_tambah', 'TOTAL ADENDUM TAMBAH (OTOMATIS)', 16], ['add_kurang', 'TOTAL ADENDUM KURANG (OTOMATIS)', 16],
      ['nilai_ok', 'NILAI OK / SETELAH ADENDUM (OTOMATIS)', 18], ['um', 'UANG MUKA DITERIMA (Rp)', 16],
      ['ret_pct', '% RETENSI', 8], ['ret_pos', 'POSISI RETENSI', 11], ['tarif', 'TARIF PPh (KOSONG = OTOMATIS)', 11],
      ['dasar_pph', 'DASAR PPh', 15], ['tampil_pph', 'TAMPILKAN PPh DI INVOICE', 10], ['template', 'TEMPLATE INVOICE', 13],
      ['top', 'TOP (HARI)', 7], ['ttd_manual', 'PENANDATANGAN MANUAL (OPSIONAL)', 20],
      ['akun_peng', 'AKUN PENGAKUAN PENDAPATAN (TAGBRUT / HAD)', 14], ['kode_jasa', 'KODE BARANG/JASA CORETAX', 12],
      ['pic', 'PIC', 12], ['aktif', 'AKTIF', 7], ['catatan', 'CATATAN', 40],
      ['um_awal', 'UM SUDAH DIKEMBALIKAN SEBELUM SISTEM (Rp)', 18],
      ['nama_pe', 'NAMA PROYEK (P&E)', 50], ['npp_produk', 'NPP PRODUK (TAMPIL DI KARTU)', 34],
      ['npp_adendum', 'NPP ADENDUM TERPISAH (TAGIHAN TETAP DI NPP INI)', 30]]
  };
  var MPRO_ALIAS = { nilai: ['NILAI KONTRAK (EXCL. PPN)'], nilai_ok: ['NILAI OK (OTOMATIS)'], addendum_lama: ['NILAI ADDENDUM TERAKHIR'] };
  var ADD = {
    sheet: 'Riwayat Adendum', headerRow: 5, firstRow: 6,
    cols: [['npp', 'NPP', 14], ['ke', 'ADENDUM KE-', 9, 'int'], ['nodok', 'NO. DOKUMEN ADENDUM', 30], ['tgl', 'TANGGAL BERLAKU', 13, 'date'],
      ['jenis', 'JENIS (TAMBAH / KURANG)', 13], ['nilai', 'NILAI ADENDUM (Rp, EXCL. PPN)', 19, 'money'], ['catatan', 'CATATAN', 60]]
  };
  // Input Tagbrut: dua tabel berdampingan — pengakuan pendapatan (kiri) dan catatan tindak lanjut (kanan)
  var TAG = {
    sheet: 'Input Tagbrut', headerRow: 6, firstRow: 7,
    cols: [['npp', 'NPP', 14], ['tgl', 'TANGGAL PENGAKUAN', 13, 'date'], ['nodok', 'NO. BUKTI SAP (BILA ADA)', 14], ['nilai', 'NILAI PENGAKUAN (Rp)', 17, 'money'],
      ['sumber', 'SUMBER', 16], ['ket', 'KETERANGAN', 34]]
  };
  var NOTE = {
    sheet: 'Input Tagbrut', headerRow: 6, firstRow: 7,
    cols: [['npp', 'NPP / PELANGGAN (CATATAN)', 22], ['tgl', 'TANGGAL CATATAN', 12, 'date'], ['pic', 'PIC', 12], ['catatan', 'CATATAN TINDAK LANJUT', 60]]
  };
  // surat permohonan percepatan pembayaran yang sudah diterbitkan (diisi aplikasi; dasar alert & peringatan tenggat)
  var SURAT = {
    sheet: 'Log Surat Penagihan', headerRow: 5, firstRow: 6,
    cols: [['nomor', 'NOMOR SURAT', 'auto', 26], ['tgl', 'TANGGAL SURAT', 'auto', 12, 'date'], ['pelanggan', 'PELANGGAN', 'auto', 30],
      ['invoices', 'INVOICE (DIPISAH ;)', 'auto', 44], ['sisa', 'TOTAL SISA TAGIHAN', 'auto', 16, 'money'], ['tenggat', 'TENGGAT PEMBAYARAN', 'auto', 12, 'date'],
      ['kepada', 'KEPADA YTH.', 'auto', 30], ['rekening', 'REKENING DICANTUMKAN', 'auto', 30], ['kontak', 'KONTAK', 'auto', 30],
      ['oleh', 'DIBUAT OLEH', 'auto', 24], ['waktu', 'WAKTU', 'auto', 16], ['link', 'ARSIP PDF', 'auto', 30]]
  };
  var APPR = {
    sheet: 'Log Persetujuan', headerRow: 5, firstRow: 6,
    cols: [['id', 'ID', 'auto', 11], ['waktu', 'WAKTU PENGAJUAN', 'auto', 16], ['jenis', 'JENIS DOKUMEN', 'auto', 20], ['ref', 'REFERENSI', 'auto', 26],
      ['label', 'URAIAN', 'auto', 40], ['nilai', 'NILAI (Rp)', 'auto', 16, 'money'], ['sidik', 'SIDIK DATA', 'auto', 14], ['oleh', 'DIAJUKAN OLEH', 'auto', 24],
      ['status', 'STATUS', 'auto', 16], ['putus_oleh', 'DIPUTUS OLEH', 'auto', 24], ['putus_waktu', 'WAKTU KEPUTUSAN', 'auto', 16],
      ['catatan', 'CATATAN', 'auto', 30], ['kode', 'KODE VERIFIKASI', 'auto', 14], ['data', 'DATA SNAPSHOT (JANGAN DIUBAH)', 'auto', 12]]
  };
  var MONTHS = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
  var OPT = {
    validasi: ['✓ CAIR', '◇ CADANGAN', '→ BULAN DEPAN', '⏸ HOLD'],
    status_inv: ['AKTIF', 'BATAL'], /* v2.6: STATUS DOKUMEN otomatis — pembatalan lewat aplikasi */ tutup: ['CLOSE MANUAL'], lunas: ['✓'], alokasi: ['OTOMATIS', 'NETTO', 'PPN'], jenis_trx: ['PINDAH BUKU'], aktif_pro: ['YA', 'TIDAK', 'SELESAI'],
    jenis: ['PROGRES', 'UANG MUKA', 'RETENSI', 'LAINNYA'], jenis_add: ['TAMBAH', 'KURANG'], template: ['UMUM', 'IKK', 'KIDOH', 'FAJARPUTERA', 'WIKA-WITON'],
    ret_pos: ['AWAL', 'AKHIR', 'TIDAK ADA'], dasar_pph: ['SETELAH RETENSI', 'SEBELUM RETENSI'], yatidak: ['YA', 'TIDAK'],
    jenis_pel: ['WIKA', 'BUMN', 'SWASTA', 'JO'], kelompok: ['PIHAK BERELASI', 'PIHAK KETIGA'], dasar_jt: ['TGL SUBMIT', 'TGL TANDA TERIMA'],
    akun_peng: ['TAGBRUT', 'HAD'], sumber_tag: ['P&E MINGGUAN', 'LK BULANAN', 'SALDO AWAL LK', 'KOREKSI'],
    peran: ['PEMBUAT', 'PENYETUJU', 'MENGETAHUI (TTD BASAH)', 'PENERIMA', 'P&E (LIHAT)']
  };
  // jenis dokumen persetujuan
  var DOC = { INV: 'PAKET INVOICE', PAY: 'PAKET PENERIMAAN', VAL: 'VALIDASI RENCANA CAIR', REK: 'REKAP BULANAN AKUNTANSI', KP: 'KARTU PIUTANG', TUTUP: 'PENUTUPAN MANUAL', BATAL: 'PEMBATALAN INVOICE', VALX: 'PEMBATALAN LEMBAR VALIDASI' };
  // rekening penerimaan termin (bawaan; dapat diubah di Parameter ▸ REKENING PENERIMAAN)
  // memorial akuntansi (format DOKUMEN MEMORIAL bagian Akuntansi) — disimpan di sheet "Memorial AR", dikirim ke file Akuntansi
  var MEMO = {
    sheet: 'Memorial AR', headerRow: 5, firstRow: 6,
    cols: [['no', 'NO', 5], ['akun', 'AKUN', 9], ['uraian', 'URAIAN TAMBAHAN', 70], ['npp', 'NO. NPP', 13], ['pelanggan', 'PELANGGAN', 28], ['produk', 'PRODUK', 45],
      ['pajak', 'PAJAK', 20], ['inv', 'NO INVOICE', 24], ['debet', 'DEBET', 15, 'money'], ['kredit', 'KREDIT', 15, 'money'], ['jd', 'JUMLAH (DEBET)', 16, 'money'],
      ['jk', 'JUMLAH (KREDIT)', 16, 'money'], ['selisih', 'SELISIH', 10, 'money'], ['status', 'STATUS', 20], ['verif', 'VERIFIKASI AKUNTANSI', 12], ['kode', 'KODE MEMORIAL', 16],
      ['bulan', 'BULAN (yyyy-mm)', 9], ['jenis', 'JENIS (INV / TRM)', 7], ['sumber', 'SUMBER (OTOMATIS / MANUAL)', 10], ['sidik', 'SIDIK DATA', 15],
      ['verif_at', 'DIVERIFIKASI PADA', 15], ['catatan', 'CATATAN SISTEM', 40], ['ref', 'REFERENSI (JANGAN DIUBAH)', 18]]
  };
  /** kode akun ringkas di DOKUMEN MEMORIAL Akuntansi (bisa diubah di Parameter ▸ PEMETAAN AKUN SAP kolom T) */
  var MEMO_KODE = { PIUTANG_USAHA: '11311', PIUTANG_RETENSI: '11312', PIUTANG_PPN: '11313', TAGBRUT: '11341', HAD: '11323', PPN_KELUARAN: '21163', UANG_MUKA_PROYEK: '21211' };
  var MEMO_LABEL = { PIUTANG_USAHA: 'PIUTANG USAHA', PIUTANG_RETENSI: 'PIUTANG RETENSI', PIUTANG_PPN: 'PIUTANG PPN', TAGBRUT: 'TAGBRUT', HAD: 'HAD', PPN_KELUARAN: 'UTANG PPNK',
    UANG_MUKA_PROYEK: 'POT UM', PPH_FINAL: 'PPh FINAL 4(2)', PPH_23: 'PPh 23', BEBAN_ADM_BANK: 'BIAYA ADM BANK', BANK: 'BANK' };
  // referensi daftar NPP dari P&E (RESUME OP) untuk pemeriksaan "Selaras P&E"
  var PEREF = { sheet: 'Referensi P&E', headerRow: 5, firstRow: 6,
    cols: [['npp', 'NPP', 14], ['jenis', 'JENIS (JASA / PRODUK / ADENDUM)', 12], ['induk', 'NPP INDUK (UNTUK PRODUK / ADENDUM)', 14], ['nama', 'NAMA DI P&E', 60], ['ringkas', 'NAMA RINGKAS PRODUK', 16], ['nilai', 'NILAI (P&E)', 17, 'money'], ['catatan', 'CATATAN', 40]] };
  var REK_DEFAULT = [['BNI', '149394879', 'BANK_BNI'], ['MANDIRI', '0060097045854', 'BANK_MANDIRI'], ['BSI', '7148973088', 'BANK_BSI'], ['BTN', '0024101300010876', 'BANK_BTN']];

  // ------------------------------------------------------------------ UTIL
  function isNum(v) { return typeof v === 'number' && isFinite(v); }
  function N(v) {
    if (isNum(v)) return v;
    if (typeof v === 'string') { var s = v.replace(/[^\d,.\-]/g, ''); if (!s) return 0; s = s.replace(/\./g, '').replace(',', '.'); var x = parseFloat(s); return isFinite(x) ? x : 0; }
    return 0;
  }
  function S(v) { return v === null || v === undefined ? '' : String(v).trim(); }
  function U(v) { return S(v).toUpperCase(); }
  function rnd(x) { return Math.round(x + (x >= 0 ? 1e-9 : -1e-9)); }
  function flo(x) { return Math.floor(x + 1e-7); }
  function pad(n, w) { n = String(n); while (n.length < w) n = '0' + n; return n; }
  // tanggal disimpan sebagai string 'yyyy-mm-dd'
  function D(v) {
    if (v === null || v === undefined || v === '') return '';
    if (Object.prototype.toString.call(v) === '[object Date]') {
      if (isNaN(v.getTime())) return '';
      return v.getFullYear() + '-' + pad(v.getMonth() + 1, 2) + '-' + pad(v.getDate(), 2);
    }
    if (isNum(v)) { // serial Excel
      var ms = Math.round((v - 25569) * 86400000), d = new Date(ms);
      return d.getUTCFullYear() + '-' + pad(d.getUTCMonth() + 1, 2) + '-' + pad(d.getUTCDate(), 2);
    }
    var s = S(v), m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) return m[1] + '-' + pad(m[2], 2) + '-' + pad(m[3], 2);
    m = s.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})$/);
    if (m) return m[3] + '-' + pad(m[2], 2) + '-' + pad(m[1], 2);
    return '';
  }
  function dnum(s) { if (!s) return NaN; var p = s.split('-'); return Date.UTC(+p[0], +p[1] - 1, +p[2]) / 86400000; }
  function addDays(s, n) { var d = new Date((dnum(s) + n) * 86400000); return d.getUTCFullYear() + '-' + pad(d.getUTCMonth() + 1, 2) + '-' + pad(d.getUTCDate(), 2); }
  function diffDays(a, b) { return dnum(a) - dnum(b); }
  function monthStart(y, m) { return y + '-' + pad(m, 2) + '-01'; }
  function monthEnd(y, m) { var d = new Date(Date.UTC(y, m, 0)); return y + '-' + pad(m, 2) + '-' + pad(d.getUTCDate(), 2); }
  function prevMonthEnd(s) { var y = +s.substring(0, 4), m = +s.substring(5, 7); return m === 1 ? monthEnd(y - 1, 12) : monthEnd(y, m - 1); }
  function monthLabel(y, m) { return MONTHS[m - 1] + ' ' + y; }
  function fmtRp(x) { var n = rnd(Math.abs(x)), s = String(n).replace(/\B(?=(\d{3})+(?!\d))/g, '.'); return (x < 0 ? '-' : '') + s; }
  function pctTxt(x) { return (Math.round(x * 10000) / 100).toString().replace('.', ',') + '%'; }
  function tglPanjang(s) { if (!s) return ''; var p = s.split('-'); return (+p[2]) + ' ' + MONTHS[+p[1] - 1] + ' ' + p[0]; }
  function nppValid(n) { return /^\d{3}C \d{4} [A-Z]{2}$/.test(S(n)); }
  function nppKey(n) { return U(n).replace(/\s+/g, ''); }
  /** '244C0817BF' / '244C 0817 BF' → '244C 0817 BF' */
  function normNpp(n) { var m = U(n).replace(/\s+/g, '').match(/^(\d{3}C)(\d{4})([A-Z]{2})$/); return m ? m[1] + ' ' + m[2] + ' ' + m[3] : S(n); }
  /** "244C 0817 BF (Spun pile); 244C 0816 DF (Half slab)" → [{npp, label}] */
  function nppList(s) {
    var out = [];
    S(s).split(/[;\n]+/).forEach(function (x) {
      var m = S(x).match(/(\d{3}C\s*\d{4}\s*[A-Za-z]{2})\s*(?:\(([^)]*)\)|[-–:]\s*(.+))?/);
      if (m) out.push({ npp: normNpp(m[1]), label: S(m[2] || m[3]) });
    });
    return out;
  }
  function nppListTxt(L) { return (L || []).map(function (x) { return x.npp + (x.label ? ' (' + x.label + ')' : ''); }).join(' · '); }
  function normInv(s) {
    var t = U(s).replace(/\s+/g, ''), m = t.match(/KU\.?02\.?02\/?WB[\.\-]?4C[\.\-]?(\d{1,4})([A-Z]?)\/(\d{4})(#\d+)?/);
    if (!m) return S(s).toUpperCase();
    return 'KU.02.02/WB-4C.' + pad(+m[1], 4) + m[2] + '/' + m[3] + (m[4] || '');
  }
  function digits(s) { return S(s).replace(/\D/g, ''); }
  /** NPWP 15 digit lama → 16 digit (awalan 0); 16 digit dibiarkan */
  function npwp16(s) { var d = digits(s); if (d.length === 15) d = '0' + d; return d; }
  function rowsToObjects(rows, cols, firstRow) {
    var out = [];
    for (var i = 0; i < (rows || []).length; i++) {
      var o = { _row: firstRow + i }, any = false;
      for (var c = 0; c < cols.length; c++) { o[cols[c][0]] = rows[i][c]; if (rows[i][c] !== '' && rows[i][c] !== null && rows[i][c] !== undefined) any = true; }
      o._empty = !any; out.push(o);
    }
    return out;
  }
  function colIndex(schema, key) { for (var i = 0; i < schema.cols.length; i++) if (schema.cols[i][0] === key) return i; return -1; }
  function sumBy(arr, f) { var t = 0; arr.forEach(function (x) { t += f(x) || 0; }); return t; }
  /** hash teks pendek & deterministik (cyrb53) — dipakai sebagai "sidik data" dokumen */
  function hash(str) {
    var h1 = 0xdeadbeef, h2 = 0x41c6ce57;
    for (var i = 0, ch; i < str.length; i++) { ch = str.charCodeAt(i); h1 = Math.imul(h1 ^ ch, 2654435761); h2 = Math.imul(h2 ^ ch, 1597334677); }
    h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507) ^ Math.imul(h2 ^ (h2 >>> 13), 3266489909);
    h2 = Math.imul(h2 ^ (h2 >>> 16), 2246822507) ^ Math.imul(h1 ^ (h1 >>> 13), 3266489909);
    var v = 4294967296 * (2097151 & h2) + (h1 >>> 0);
    return pad(v.toString(16).toUpperCase(), 14);
  }

  // ------------------------------------------------------------------ PARAMETER
  /** kv: A:C · signers: E:G · targets: I:K · akun: O:S · users: U:Y · ckpn: AA:AD (lihat sheet Parameter) */
  function readParams(kv, signers, targets, akun, users, ckpn, rek) {
    var p = { TARIF_JT: 0.0265, TARIF_JL: 0.02, TARIF_JF: 0.0265, ADM_DEFAULT: 25000, TOL_PEMBULATAN: 1000, TOL_LEBIH: 1000, KOTA: 'Jakarta',
      JT_AKUNTANSI_HARI: 30, CKPN_DISKONTO: 0.0833, CKPN_PIHAK_BERELASI: 'TIDAK', TAGBRUT_SALDO_AWAL_PER: '2026-09-30', TAGBRUT_BATAS_HARI: 30,
      FAKTUR_KODE_TRANSAKSI: '04', FAKTUR_PENGINGAT_HARI: 10, SAP_COMPANY_CODE: 'B000', SAP_KODE_UNIT: 'BE00100000', MEMORIAL_MULAI: '2026-10-01', PE_ABAIKAN: '254C 1584 JL',
      KINERJA_LANCAR_HARI: 15, KINERJA_PERHATIAN_HARI: 60, KINERJA_KURANG_HARI: 90, EWS_LEWAT_JT_HARI: 60, EWS_CICILAN_HENTI_HARI: 30, EWS_OVERDUE_PER_PELANGGAN: 2, EWS_WAITING_SUBMIT_HARI: 14, EWS_KONSENTRASI_PCT: 30 };
    (kv || []).forEach(function (r) { var k = U(r[0]); if (k) p[k] = (isNum(r[1]) ? r[1] : (Object.prototype.toString.call(r[1]) === '[object Date]' ? D(r[1]) : S(r[1]))); });
    p.signers = (signers || []).filter(function (r) { return S(r[1]); }).map(function (r) {
      return { batas: isNum(r[0]) ? r[0] : (S(r[0]) === '' ? Infinity : N(r[0]) || Infinity), nama: S(r[1]), jabatan: S(r[2]) };
    }).sort(function (a, b) { return a.batas - b.batas; });
    p.targets = {};
    (targets || []).forEach(function (r) { if (isNum(r[0]) && isNum(r[1]) && isNum(r[2])) p.targets[r[0] + '-' + pad(r[1], 2)] = r[2]; });
    p.akun = {};
    (akun || []).forEach(function (r) { var k = U(r[0]); if (k) p.akun[k] = { nama: S(r[1]), rel: S(r[2]), tiga: S(r[3]) || S(r[2]), ket: S(r[4]), memo: S(r[5]) }; });
    p.users = (users || []).filter(function (r) { return S(r[1]); }).map(function (r) { return { email: S(r[0]).toLowerCase(), nama: S(r[1]), jabatan: S(r[2]), nip: S(r[3]), peran: U(r[4]) }; });
    p.ckpn = (ckpn || []).filter(function (r) { return S(r[0]) && isNum(r[1]); }).map(function (r) {
      return { kelas: S(r[0]), dari: r[1], sampai: isNum(r[2]) ? r[2] : Infinity, tarif: N(r[3]) };
    }).sort(function (a, b) { return a.dari - b.dari; });
    if (!p.ckpn.length) p.ckpn = CKPN_DEFAULT.map(function (r) { return { kelas: r[0], dari: r[1], sampai: r[2], tarif: r[3] }; });
    var rk = (rek || []).filter(function (r) { return S(r[0]) && digits(r[1]); });
    if (!rk.length) rk = REK_DEFAULT;
    p.rek = rk.map(function (r) { return { bank: U(r[0]), no: digits(r[1]), akun: U(r[2]) || ('BANK_' + U(r[0]).replace(/[^A-Z]/g, '')), label: U(r[0]) + ' ' + digits(r[1]) }; });
    return p;
  }
  var CKPN_DEFAULT = [['A', 0, 90, 0.0003], ['B', 91, 180, 0.0008], ['C', 181, 270, 0.0017], ['D', 271, 360, 0.0031], ['E', 361, 450, 0.0055],
    ['F', 451, 540, 0.0088], ['G', 541, 630, 0.0141], ['H', 631, 720, 0.0286], ['I', 721, 810, 0.0507], ['J', 811, 900, 0.0918],
    ['K', 901, 990, 0.1774], ['L', 991, 1080, 0.428], ['M', 1081, Infinity, 1]];
  /** akun GL per kode & kelompok pelanggan */
  function gl(params, kode, berelasi) { var a = params.akun[kode]; if (!a) return ''; return berelasi ? a.rel : a.tiga; }
  /** rekening dari teks Log Penerimaan ("BNI 149394879", "149394879", "BNI") */
  function rekOf(params, v) {
    var s = U(v), d = digits(v); if (!s) return null;
    return (params.rek || []).filter(function (r) { return (d.length >= 6 && (r.no.indexOf(d) >= 0 || d.indexOf(r.no) >= 0)) || (!d && s.indexOf(r.bank) === 0); })[0] || null;
  }
  function glName(params, kode) { var a = params.akun[kode]; return a ? a.nama : kode; }
  function userBy(params, email) { email = S(email).toLowerCase(); return params.users.filter(function (u) { return u.email && u.email === email; })[0] || null; }
  function userByRole(params, role) { return params.users.filter(function (u) { return u.peran.indexOf(role) === 0; })[0] || null; }
  function tarifDefault(params, npp) {
    var k = S(npp).slice(-2).toUpperCase();
    var t = params['TARIF_' + k];
    return isNum(t) ? t : 0.0265;
  }
  function signerFor(params, nilaiOk, manual) {
    if (S(manual)) {
      var hit = params.signers.filter(function (s) { return U(s.nama) === U(manual); })[0];
      return hit ? { nama: hit.nama, jabatan: hit.jabatan } : { nama: S(manual), jabatan: '' };
    }
    if (!isNum(nilaiOk) || nilaiOk <= 0) return null;
    for (var i = 0; i < params.signers.length; i++) if (params.signers[i].batas === null || params.signers[i].batas === undefined || nilaiOk <= params.signers[i].batas) return { nama: params.signers[i].nama, jabatan: params.signers[i].jabatan };
    var last = params.signers[params.signers.length - 1];
    return last ? { nama: last.nama, jabatan: last.jabatan } : null;
  }
  /** delegasi persetujuan sementara (Parameter DELEGASI_*): akun pengganti boleh memutuskan a.n. Section Head sampai tanggal batas, hanya untuk jenis tertentu */
  function delegasiFor(params, email, kind, today) {
    var e = S(email).toLowerCase(), de = S(params.DELEGASI_EMAIL).toLowerCase(), until = D(params.DELEGASI_SAMPAI);
    if (!e || !de || e !== de || !until || (today && today > until)) return null;
    // pemisahan tugas: delegasi tidak pernah jatuh ke akun PEMBUAT (pembuat tidak boleh menyetujui paket yang ia susun)
    var du = userBy(params, e); if (du && /^PEMBUAT/.test(U(du.peran))) return null;
    var J = U(params.DELEGASI_JENIS || 'PAY,VAL').split(/[,;\s]+/).filter(String);
    var kk = U(kind) === 'VALX' ? 'VAL' : U(kind); // pembatalan lembar validasi = bagian dari validasi rencana cair
    if (kind && J.indexOf(kk) < 0 && J.indexOf('SEMUA') < 0) return null;
    var sh = userByRole(params, 'PENYETUJU') || {};
    return { atas: sh.nama || 'Section Head', jabatan: sh.jabatan || 'Section Head', sampai: until, jenis: J, dasar: S(params.DELEGASI_DASAR) };
  }
  function isDelegasi(rec) { return /^\[DELEGASI/.test(S(rec && rec.catatan)); }
  function pphLabel(npp, tarif) { return (S(npp).slice(-2).toUpperCase() === 'JL' ? 'PPh 23 ' : 'PPh Final 4(2) ') + pctTxt(tarif); }
  function pphGlKode(npp) { return S(npp).slice(-2).toUpperCase() === 'JL' ? 'PPH_23' : 'PPH_FINAL'; }

  // ------------------------------------------------------------------ MASTER
  function buildMaster(pelRows, proRows, params, addRows) {
    var pel = {}, pro = {}, issues = [], adds = {};
    rowsToObjects(addRows || [], ADD.cols, ADD.firstRow).forEach(function (a) {
      if (a._empty) return;
      var n = S(a.npp), j = U(a.jenis), v = N(a.nilai), ref = 'baris ' + a._row + (n ? ' (' + n + ')' : '');
      if (!n) { issues.push(['ADENDUM', ADD.sheet, ref, 'NPP kosong', 'Isi NPP proyek']); return; }
      if (j !== 'TAMBAH' && j !== 'KURANG') { issues.push(['ADENDUM', ADD.sheet, ref, 'Jenis adendum harus TAMBAH atau KURANG', 'Pilih dari daftar']); return; }
      if (!(v > 0)) { issues.push(['ADENDUM', ADD.sheet, ref, 'Nilai adendum kosong / nol', 'Isi nilai adendum (angka positif)']); return; }
      var t = D(a.tgl);
      if (!t) issues.push(['ADENDUM', ADD.sheet, ref, 'Tanggal berlaku kosong — dianggap berlaku sejak awal kontrak', 'Isi TANGGAL BERLAKU agar invoice sebelum adendum tetap memakai nilai lama']);
      (adds[n] = adds[n] || []).push({ ke: a.ke, nodok: S(a.nodok), tgl: t, jenis: j, nilai: v });
    });
    rowsToObjects(pelRows, MPEL.cols, MPEL.firstRow).forEach(function (o) {
      if (o._empty || !S(o.nama)) return;
      var key = U(o.nama);
      if (pel[key]) issues.push(['MASTER', MPEL.sheet, 'baris ' + o._row, 'Nama pelanggan ganda: ' + key, 'Hapus salah satu baris']);
      if (S(o.nama) !== U(o.nama)) issues.push(['MASTER', MPEL.sheet, 'baris ' + o._row, 'Nama pelanggan belum huruf kapital: ' + S(o.nama), 'Tulis dengan huruf kapital']);
      o.nama = key; o.adm = isNum(o.adm) ? o.adm : N(params.ADM_DEFAULT); o.jenis = U(o.jenis) || 'SWASTA';
      o.dasar_jt = U(o.dasar_jt) || 'TGL SUBMIT'; o.top_ret = isNum(o.top_ret) && N(o.top_ret) > 0 ? N(o.top_ret) : ''; o.aktif = U(o.aktif) || 'YA';
      o.kelompok = U(o.kelompok) || (o.jenis === 'WIKA' || o.jenis === 'JO' ? 'PIHAK BERELASI' : 'PIHAK KETIGA');
      o.berelasi = o.kelompok.indexOf('BERELASI') >= 0;
      o.npwp16 = npwp16(o.npwp); o.kode_bp = S(o.kode_bp);
      o.nitku = digits(o.nitku) || (o.npwp16 ? o.npwp16 + '000000' : '');
      pel[key] = o;
    });
    rowsToObjects(proRows, MPRO.cols.concat([['addendum_lama', '', 0]]), MPRO.firstRow).forEach(function (o) {
      if (o._empty || !S(o.npp)) return;
      var n = S(o.npp);
      if (!nppValid(n)) issues.push(['MASTER', MPRO.sheet, 'baris ' + o._row, 'Format NPP tidak standar: "' + n + '"', 'Gunakan format 254C 0823 JT (kode, spasi, 4 digit, spasi, 2 huruf)']);
      if (pro[n]) issues.push(['MASTER', MPRO.sheet, 'baris ' + o._row, 'NPP ganda: ' + n, 'Satu NPP hanya boleh satu baris']);
      o.npp = n; o.pelanggan = U(o.pelanggan);
      if (!o.pelanggan) issues.push(['MASTER', MPRO.sheet, n, 'Proyek belum punya pelanggan', 'Isi NAMA PELANGGAN']);
      else if (!pel[o.pelanggan]) issues.push(['MASTER', MPRO.sheet, n, 'Pelanggan "' + o.pelanggan + '" tidak ada di Master Pelanggan', 'Tambahkan pelanggan atau perbaiki namanya']);
      o.nilai_awal = N(o.nilai);
      o.adds = (adds[n] || []).sort(function (x, y) { return (x.tgl || '') < (y.tgl || '') ? -1 : (x.tgl || '') > (y.tgl || '') ? 1 : N(x.ke) - N(y.ke); });
      o.add_tambah = sumBy(o.adds, function (x) { return x.jenis === 'TAMBAH' ? x.nilai : 0; });
      o.add_kurang = sumBy(o.adds, function (x) { return x.jenis === 'KURANG' ? x.nilai : 0; });
      o.nilai_ok = o.nilai_awal + o.add_tambah - o.add_kurang;
      if (!o.adds.length && N(o.addendum_lama) > 0) {
        o.nilai_ok = N(o.addendum_lama);
        issues.push(['MASTER', MPRO.sheet, n, 'Masih memakai kolom lama "Nilai Addendum Terakhir"', 'Jalankan menu AR CSU ▸ Siapkan Riwayat Adendum']);
      }
      if (o.adds.length && !o.nilai_awal) issues.push(['ADENDUM', MPRO.sheet, n, 'Ada adendum tetapi nilai kontrak awal kosong', 'Isi NILAI KONTRAK AWAL di Master Proyek']);
      if (o.adds.length && o.nilai_ok <= 0) issues.push(['ADENDUM', MPRO.sheet, n, 'Nilai kontrak setelah adendum ≤ 0', 'Periksa nilai adendum kurang']);
      o.tarif_eff = isNum(o.tarif) && o.tarif > 0 ? o.tarif : tarifDefault(params, n);
      o.ret_pct = N(o.ret_pct); o.ret_pos = U(o.ret_pos) || (o.ret_pct ? 'AWAL' : 'TIDAK ADA');
      o.dasar_pph = U(o.dasar_pph) || 'SETELAH RETENSI'; o.tampil_pph = U(o.tampil_pph) || 'YA';
      o.template = U(o.template) || 'UMUM'; o.top = isNum(o.top) ? o.top : 30; o.um = N(o.um); o.um_awal = N(o.um_awal); o.aktif = U(o.aktif) || 'YA'; o.selesai = o.aktif === 'SELESAI';
      o.akun_peng = U(o.akun_peng) === 'HAD' ? 'HAD' : (U(o.akun_peng) === 'TAGBRUT' ? 'TAGBRUT' : (n.slice(-2) === 'JL' ? 'HAD' : 'TAGBRUT'));
      o.kode_jasa = S(o.kode_jasa);
      o.singkat = S(o.singkat) || n;
      o.nama_pe = S(o.nama_pe); o.produk = nppList(o.npp_produk); o.adendumNpp = nppList(o.npp_adendum);
      if (/^\d{3}C 0000 [A-Z]{2}$/.test(n)) issues.push(['MASTER', MPRO.sheet, n, 'NPP sementara "' + n + '" — tidak unik, kartu & lampiran bisa tercampur dengan proyek lain yang memakai NPP sama', 'Ganti dengan NPP sementara yang unik (mis. ' + n.replace(' 0000 ', ' 9001 ') + '), lalu AR CSU ▸ Ganti NPP setelah NPP resmi terbit']);
      var sg = signerFor(params, o.nilai_ok, o.ttd_manual);
      o.signer = sg;
      // pemeriksaan kelengkapan untuk penagihan berikutnya dilakukan di extraChecks (perlu data invoice)
      o._noSigner = !sg;
      pro[n] = o;
    });
    Object.keys(adds).forEach(function (n) { if (!pro[n]) issues.push(['ADENDUM', ADD.sheet, n, 'NPP adendum tidak ada di Master Proyek', 'Perbaiki NPP']); });
    // NPP produk / adendum milik proyek lain: bukan proyek AR (tidak muncul di daftar & kartu)
    Object.keys(pro).forEach(function (n) {
      pro[n].produk.forEach(function (x) { if (pro[x.npp] && x.npp !== n) { pro[x.npp]._produkOf = n; } });
      pro[n].adendumNpp.forEach(function (x) { if (pro[x.npp] && x.npp !== n) { pro[x.npp]._adendumOf = n; } });
    });
    return { pel: pel, pro: pro, issues: issues };
  }
  /** nilai kontrak yang berlaku pada tanggal tertentu (adendum dengan tanggal berlaku ≤ tgl, atau tanpa tanggal) */
  function kontrakAt(p, tgl) {
    var awal = N(p.nilai_awal !== undefined ? p.nilai_awal : p.nilai), t = 0, k = 0, list = [];
    (p.adds || p.adendum || []).forEach(function (a) {
      if (a.tgl && tgl && a.tgl > tgl) return;
      list.push(a); if (a.jenis === 'TAMBAH') t += N(a.nilai); else k += N(a.nilai);
    });
    if (!list.length && !(p.adds || p.adendum || []).length && N(p.nilai_ok) && N(p.nilai_ok) !== awal) return { awal: N(p.nilai_ok), tambah: 0, kurang: 0, net: N(p.nilai_ok), list: [] };
    return { awal: awal, tambah: t, kurang: k, net: awal + t - k, list: list };
  }

  // ------------------------------------------------------------------ PERHITUNGAN INVOICE
  function invoiceAmounts(inv, proj, params) {
    var dpp = rnd(N(inv.dpp)), ppn = rnd(N(inv.ppn)), ret = rnd(N(inv.retensi)), umk = rnd(N(inv.umk));
    var retpos = U(inv.retpos) || 'TIDAK ADA';
    var netto = dpp - (retpos === 'AKHIR' ? ret : 0);
    var tarif = isNum(inv.tarif) && inv.tarif > 0 ? inv.tarif : (proj ? proj.tarif_eff : tarifDefault(params, inv.npp));
    var dasar = U(inv.dasar) || (proj ? proj.dasar_pph : 'SETELAH RETENSI');
    var base = dasar === 'SEBELUM RETENSI' ? dpp : netto;
    var tpl = U(inv.template) || (proj ? proj.template : 'UMUM');
    var pph = tpl === 'KIDOH' ? flo(tarif * base) : rnd(tarif * base);
    var altBase = dasar === 'SEBELUM RETENSI' ? netto : dpp;
    var pphAlt = tpl === 'KIDOH' ? flo(tarif * altBase) : rnd(tarif * altBase);
    var jenis = U(inv.jenis) || 'PROGRES';
    // nilai bruto (pengurang tagbrut): progres sebelum potongan UM & retensi
    var bruto = N(inv.bruto) > 0 ? rnd(N(inv.bruto)) : netto + ret + umk;
    return { dpp: dpp, ppn: ppn, ret: ret, umk: umk, retpos: retpos, total: dpp + ppn, netto: netto, tarif: tarif, dasar: dasar, pph: pph, pphAlt: pphAlt,
      kas: netto + ppn - pph, bruto: bruto, jenis: jenis, tpl: tpl };
  }

  // ------------------------------------------------------------------ ALOKASI PEMBAYARAN
  /** Alokasi berurutan per tanggal. Aturan OTOMATIS:
   *  1) nominal = sisa PPN → PPN · 2) nominal = sisa Netto − PPh (± adm bank) → Netto
   *  3) nominal < sisa PPN → Netto (PPN tidak dicicil) · 4) selebihnya: PPN dilunasi sekaligus, sisa → Netto.
   *  Kolom ALOKASI (NETTO / PPN) mengalahkan aturan otomatis. Baris migrasi membaca keterangan lama ("Netto" / "PPn").
   *  PPh & adm bank diakui pada pembayaran yang menyelesaikan porsi Netto. */
  /** aturan alokasi satu pembayaran (dipakai mesin & pratinjau di aplikasi) */
  function allocRule(remP, remNK, remN, x, al, admMax, TOLR, TOLL) {
    var toP = 0, rule = '';
    if (al === 'PPN') { toP = Math.min(x, remP); rule = toP < x ? 'Ditetapkan PPN; kelebihan → Netto' : 'Ditetapkan manual: PPN'; }
    // data lama berketerangan "Netto" yang nominalnya ternyata melunasi Netto (setelah PPh & adm) sekaligus PPN
    else if (al === 'NETTO' && remP > 0 && x >= remNK + remP - admMax - TOLR) { toP = remP; rule = 'Ditetapkan Netto, tetapi nominal melunasi Netto & PPN → PPN dilunasi, sisa → Netto'; }
    else if (al === 'NETTO') { toP = x > remN + TOLL ? Math.min(remP, x - remN) : 0; rule = toP ? 'Ditetapkan Netto; kelebihan → PPN' : 'Ditetapkan manual: Netto'; }
    else if (remP <= 0) { toP = 0; rule = 'PPN sudah lunas → Netto'; }
    else if (Math.abs(x - remP) <= TOLR) { toP = Math.min(x, remP); rule = 'Nominal = sisa PPN → PPN'; }
    else if (remNK > TOLR && x <= remNK + TOLR && x >= remNK - admMax - TOLR) { toP = 0; rule = 'Nominal = sisa Netto − PPh → Netto'; }
    else if (x < remP && remNK > TOLR) { toP = 0; rule = 'Nominal < sisa PPN → Netto (PPN tidak dicicil)'; }
    else if (x < remP) { toP = x; rule = 'Netto sudah selesai → PPN'; }
    else { toP = remP; rule = 'PPN dilunasi sekaligus, sisa → Netto'; }
    return { toP: toP, rule: rule };
  }
  function allocate(inv, params) {
    var a = inv._a, pel = inv._pel, admMax = pel ? N(pel.adm) : N(params.ADM_DEFAULT), TOLR = N(params.TOL_PEMBULATAN), TOLL = N(params.TOL_LEBIH);
    var keP = 0, keN = 0, settle = null;
    inv._pays.forEach(function (p, i) {
      var x = rnd(p.nominal), remP = Math.max(0, a.ppn - keP), remNK = (a.netto - a.pph) - keN, remN = Math.max(0, a.netto - keN), toP = 0, rule = '';
      if (p._pb && x < 0) { toP = 0; rule = 'Pindah buku ke ' + (p._other || 'invoice lain') + ' (kelebihan bayar dialihkan)'; }
      else {
        var al = U(p.alokasi); if (al === 'OTOMATIS') al = '';
        if (!al && p._migr) {
          var k = U(p.ket);
          if (/PPN/.test(k) && !/NET/.test(k)) al = 'PPN'; else if (/NET/.test(k) && !/PPN/.test(k)) al = 'NETTO';
          if (al) rule = 'Data lama (keterangan "' + S(p.ket).replace(/^Migrasi:\s*/i, '') + '") → ';
        }
        var ar = allocRule(remP, remNK, remN, x, al, admMax, TOLR, TOLL);
        toP = ar.toP; rule = rule ? rule + (al === 'PPN' ? (toP < x ? 'PPN; kelebihan → Netto' : 'PPN') : (toP ? 'Netto; kelebihan → PPN' : 'Netto')) : ar.rule;
        if (p._pb) rule = 'Pindah buku dari ' + (p._other || 'invoice lain') + ' · ' + rule;
      }
      keP += toP; keN += x - toP;
      p._keP = toP; p._keN = x - toP; p._rule = rule; p._ke = i + 1; p._pph = 0; p._adm = 0; p._sel = 0;
      // pindah buku keluar membuka kembali porsi Netto: pengakuan PPh/adm sebelumnya dihitung ulang
      if (x < 0 && settle) { var sp = inv._pays[settle.i]; sp._pph = 0; sp._adm = 0; sp._sel = 0; settle = null; }
      if (!settle && a.netto > 0 && p._keN !== 0) {
        var ns = a.netto - keN;
        if (ns <= a.pph + admMax + TOLR) {
          settle = { i: i, tgl: p.tgl, pot: ns >= a.pph ? a.pph : Math.max(0, ns), adm: ns >= a.pph ? ns - a.pph : 0, sel: 0 };
        } else if (a.ret > 0 && a.retpos === 'AKHIR' && a.pphAlt > a.pph && ns <= a.pphAlt + admMax + TOLR && ns >= a.pphAlt - TOLR) {
          settle = { i: i, tgl: p.tgl, pot: a.pph, sel: a.pphAlt - a.pph, adm: Math.max(0, ns - a.pphAlt) };
        }
        if (settle) { p._pph = settle.pot; p._adm = settle.adm; p._sel = settle.sel; }
      }
    });
    // pelunasan ditandai (✓) tetapi Netto kurang: PPh tetap diakui pada pembayaran pelunasan; sisanya = kurang bayar
    if (!settle && a.netto > 0) {
      var fl = inv._pays.filter(function (p) { return S(p.lunas); });
      if (fl.length) { var lp = fl[fl.length - 1], ns2 = a.netto - keN; settle = { i: inv._pays.indexOf(lp), tgl: lp.tgl, pot: Math.max(0, Math.min(a.pph, ns2)), adm: 0, sel: 0, kurang: true }; lp._pph = settle.pot; }
    }
    inv._settle = settle;
    return { keP: keP, keN: keN, admMax: admMax };
  }

  /** posisi satu invoice per tanggal tertentu (untuk CKPN, rekon, kartu per tanggal) */
  function invAt(inv, C, params) {
    var a = inv._a, TOL = N(params.TOL_PEMBULATAN);
    var P = inv._pays.filter(function (p) { return p.tgl && p.tgl <= C; });
    if (P.length !== inv._pays.length) {
      P = P.map(function (p) { var q = {}; Object.keys(p).forEach(function (k) { q[k] = p[k]; }); return q; });
      var tmp = { _a: a, _pel: inv._pel, _pays: P }; allocate(tmp, params);
    }
    var keN = sumBy(P, function (p) { return p._keN; }), keP = sumBy(P, function (p) { return p._keP; });
    var pot = sumBy(P, function (p) { return p._pph + p._adm + p._sel; });
    // sama dengan praktik SAP unit: pembayaran sebagian mengurangi piutang sebesar kas; PPh & adm diakui saat porsi Netto selesai
    var sN = a.netto - keN - pot, sP = a.ppn - keP;
    var r = inv._res;
    if ((r.closed && !inv._closedManual && r.tgl_bayar && r.tgl_bayar <= C) || (inv._closedManual && (!r.tgl_tutup || r.tgl_tutup <= C))) { sN = 0; sP = 0; }
    if (Math.abs(sN) <= TOL) sN = 0; if (Math.abs(sP) <= TOL) sP = 0;
    return { keN: keN, keP: keP, saldoN: rnd(sN), saldoP: rnd(sP), paid: P };
  }

  // ------------------------------------------------------------------ SIDIK DATA (untuk persetujuan)
  function fpInvoice(inv) { var a = inv._a; return hash(['INV', normInv(inv.nomor), N(inv.rev), a.dpp, a.ppn, a.netto, a.pph, a.ret, S(inv.npp), inv.tgl, inv.pelanggan].join('|')); }
  function payKey(p) { return S(p.ref) && U(p.ref) !== 'MIGRASI' ? 'REF ' + S(p.ref) : p.idp; }
  function fpPayGroup(rows) {
    return hash('PAY|' + rows.slice().sort(function (x, y) { return x.idp < y.idp ? -1 : 1; }).map(function (p) {
      return [p.idp, p.tgl, normInv(p.inv), rnd(p.nominal), p._keP || 0, p._keN || 0, p._pph || 0, p._adm || 0].join(',');
    }).join(';'));
  }
  function apprStatusText(rec, fpNow) {
    if (!rec) return '';
    var st = U(rec.status), same = !fpNow || S(rec.sidik) === fpNow, tg = S(rec.putus_waktu).substring(0, 10);
    if (st === 'DISETUJUI') return same ? '✓ DISETUJUI ' + tg + (isDelegasi(rec) ? ' · a.n. Section Head oleh ' + (rec.putus_nama || rec.putus_oleh) + ' (delegasi)' : (rec.putus_nama ? ' · ' + rec.putus_nama : '')) : '⚠ DATA BERUBAH SETELAH DISETUJUI — AJUKAN ULANG';
    if (st === 'MENUNGGU') return same ? '⏳ MENUNGGU PERSETUJUAN' : '⚠ DATA BERUBAH — AJUKAN ULANG';
    if (st === 'DITOLAK') return '✗ DITOLAK' + (S(rec.catatan) ? ': ' + S(rec.catatan) : '');
    return st;
  }

  // ------------------------------------------------------------------ HITUNG SEMUA
  function compute(input) {
    var params = readParams(input.paramKV, input.paramSigners, input.paramTargets, input.paramAkun, input.paramUsers, input.paramCkpn, input.paramRek);
    var today = D(input.today) || D(new Date());
    var M = buildMaster(input.pelanggan, input.proyek, params, input.adendum);
    var issues = M.issues.slice();
    if (input.ssTz && input.ssTz !== 'Asia/Jakarta') issues.unshift(['SISTEM', 'Setelan file', 'Zona waktu: ' + input.ssTz, 'Zona waktu file bukan (GMT+07:00) Jakarta — tanggal di laporan bisa bergeser 1 hari', 'Jalankan menu AR CSU ▸ Perbaiki zona waktu & tanggal']);
    var jam = /\d{4}-\d{2}-\d{2} \d{2}:\d{2}/;
    var invs = rowsToObjects(input.register, REG.cols, REG.firstRow);
    // v2.6: pembatalan invoice lewat aplikasi (persetujuan Section Head) → STATUS DOKUMEN otomatis
    var batalOk = {}, batalPend = {}, batalLama = null;
    if (params.BATAL_LAMA !== undefined) { batalLama = {}; S(params.BATAL_LAMA).split(/[;\n]+/).forEach(function (x) { x = S(x); if (x && x !== '-') batalLama[normInv(x)] = 1; }); }
    rowsToObjects(input.approvals || [], APPR.cols, APPR.firstRow).forEach(function (r) {
      if (r._empty || U(r.jenis) !== DOC.BATAL) return; var k = S(r.ref), st = U(r.status);
      if (st === 'DISETUJUI') batalOk[k] = { id: S(r.id), tgl: S(r.putus_waktu).substring(0, 10) }; else if (st === 'MENUNGGU') batalPend[k] = S(r.id); else if (st === 'DITOLAK' || st === 'DIBATALKAN') delete batalPend[k];
    });
    var pays = rowsToObjects(input.log, LOG.cols, LOG.firstRow);
    var byNo = {};
    invs.forEach(function (inv) {
      if (inv._empty || !S(inv.nomor)) { inv._skip = true; return; }
      inv.nomor = S(inv.nomor);
      var key = normInv(inv.nomor);
      if (byNo[key]) { issues.push(['REGISTER', REG.sheet, inv.nomor, 'Nomor invoice ganda (baris ' + byNo[key]._row + ' & ' + inv._row + ')', 'Nomor harus unik; gunakan fitur revisi untuk mengubah invoice']); }
      else byNo[key] = inv;
      if (inv.nomor.indexOf('#') >= 0 && !(batalLama ? batalLama[key] : /^BATAL/.test(U(inv.status_inv))) && !batalOk[key]) issues.push(['REGISTER', REG.sheet, inv.nomor, 'Nomor invoice ditandai ganda dari data lama (#)', 'Cek arsip, ganti dengan nomor yang benar']);
      if (!/WB-4C\./.test(inv.nomor)) issues.push(['REGISTER', REG.sheet, inv.nomor, 'Format nomor tidak standar (harus KU.02.02/WB-4C.xxxx/tahun)', 'Jalankan menu AR CSU ▸ Seragamkan format nomor invoice']);
      inv.pelanggan = U(inv.pelanggan); inv.npp = S(inv.npp);
      inv._pro = M.pro[inv.npp] || null; inv._pel = M.pel[inv.pelanggan] || null;
      if (!inv._pro) issues.push(['REGISTER', REG.sheet, inv.nomor, 'NPP "' + inv.npp + '" tidak ada di Master Proyek', 'Tambahkan proyek di Master Proyek atau perbaiki NPP']);
      if (!inv._pel) issues.push(['REGISTER', REG.sheet, inv.nomor, 'Pelanggan "' + inv.pelanggan + '" tidak ada di Master Pelanggan', 'Perbaiki nama pelanggan']);
      if (inv._pro && inv._pel && inv._pro.pelanggan && inv._pro.pelanggan !== inv.pelanggan) issues.push(['REGISTER', REG.sheet, inv.nomor, 'Pelanggan invoice (' + inv.pelanggan + ') beda dengan pelanggan proyek ' + inv.npp + ' (' + inv._pro.pelanggan + ')', 'Periksa NPP atau pelanggan']);
      if (jam.test(S(inv.tgl))) issues.push(['REGISTER', REG.sheet, inv.nomor, 'Tanggal invoice tersimpan dengan jam (' + S(inv.tgl) + ') — kemungkinan bergeser karena zona waktu', 'Jalankan menu AR CSU ▸ Perbaiki zona waktu & tanggal']);
      inv.tgl = D(inv.tgl); inv.submit = D(inv.submit); inv.tt = D(inv.tt); inv.rencana = D(inv.rencana); inv.tgl_fp = D(inv.tgl_fp);
      inv.no_fp = digits(inv.no_fp); inv.sap_doc = S(inv.sap_doc); inv.uraian = S(inv.uraian);
      if (!inv.tgl) issues.push(['REGISTER', REG.sheet, inv.nomor, 'Tanggal invoice kosong / tidak terbaca', 'Isi TANGGAL INVOICE']);
      else if (inv.tgl > today) issues.push(['REGISTER', REG.sheet, inv.nomor, 'Tanggal invoice di masa depan (' + inv.tgl + ')', 'Kemungkinan salah ketik tahun']);
      inv._a = invoiceAmounts(inv, inv._pro, params);
      inv._batalInfo = batalOk[key] || null; inv._batal = !!inv._batalInfo || (batalLama ? !!batalLama[key] : /^BATAL/.test(U(inv.status_inv))); inv._batalPend = inv._batal ? '' : (batalPend[key] || '');
      inv._berelasi = inv._pel ? inv._pel.berelasi : false;
      inv._pays = [];
    });
    // ---- ID penerimaan (stabil: ID yang sudah ada dipertahankan)
    var maxId = 0;
    pays.forEach(function (p) { var m = S(p.idp).match(/^PNR-(\d+)$/); if (m) maxId = Math.max(maxId, +m[1]); });
    // ---- pembayaran
    var refSum = {}, pbRef = {}, rekMulai = D(params.REKENING_WAJIB_MULAI) || '2026-10-09', newFrom = isNum(input.newFrom) ? input.newFrom : Infinity;
    pays.forEach(function (p) {
      if (p._empty) { p._skip = true; return; }
      p._new = (p._row - LOG.firstRow) >= newFrom;
      p._pb = U(p.jenis_trx) === 'PINDAH BUKU';
      p._rek = rekOf(params, p.rekening);
      if (!S(p.idp)) p.idp = 'PNR-' + pad(++maxId, 5); else p.idp = S(p.idp);
      if (jam.test(S(p.tgl))) issues.push(['PENERIMAAN', LOG.sheet, 'baris ' + p._row, 'Tanggal bayar tersimpan dengan jam (' + S(p.tgl) + ') — kemungkinan bergeser karena zona waktu', 'Jalankan menu AR CSU ▸ Perbaiki zona waktu & tanggal']);
      p.tgl = D(p.tgl); p.inv = S(p.inv); p.nominal = N(p.nominal); p._cek = []; p.sap_doc = S(p.sap_doc);
      p._migr = U(p.ref) === 'MIGRASI' || /^MIGRASI/.test(U(p.ket));
      var al = U(p.alokasi); if (al && OPT.alokasi.indexOf(al) < 0) p._cek.push('⚠ ALOKASI HARUS OTOMATIS / NETTO / PPN');
      var key = normInv(p.inv), inv = byNo[key];
      p._inv = inv && !inv._skip ? inv : null;
      if (!p.inv) p._cek.push('⚠ NOMOR INVOICE KOSONG');
      else if (!p._inv) p._cek.push('⚠ NOMOR INVOICE TIDAK DITEMUKAN');
      else if (p._inv._batal) p._cek.push('⚠ INVOICE BERSTATUS BATAL');
      if (!p.tgl) p._cek.push('⚠ TANGGAL KOSONG');
      if (p._pb ? !p.nominal : !(p.nominal > 0)) p._cek.push(p._pb ? '⚠ NOMINAL PINDAH BUKU KOSONG' : '⚠ NOMINAL KOSONG / NEGATIF (nominal negatif hanya untuk PINDAH BUKU)');
      if (p._pb) { if (!S(p.ref)) p._cek.push('⚠ PINDAH BUKU WAJIB PUNYA REF (mis. PB-0001)'); else (pbRef[S(p.ref)] = pbRef[S(p.ref)] || []).push(p); }
      else if (!p._migr && p.tgl && p.tgl >= rekMulai) {
        if (!S(p.rekening)) p._cek.push('⚠ REKENING PENERIMAAN KOSONG');
        else if (!p._rek) p._cek.push('⚠ REKENING "' + S(p.rekening) + '" TIDAK ADA DI DAFTAR REKENING PENERIMAAN');
      }
      if (p._inv && p.tgl && p._inv.tgl && p.tgl < p._inv.tgl) p._cek.push('⚠ TGL BAYAR SEBELUM TGL INVOICE');
      if (p.tgl && p.tgl > today) p._cek.push('⚠ TGL BAYAR DI MASA DEPAN');
      if (S(p.ref) && !p._pb) { var r = refSum[S(p.ref)] || (refSum[S(p.ref)] = { sum: 0, transfer: 0, rows: [] }); r.sum += p.nominal; if (N(p.transfer) > 0) r.transfer = N(p.transfer); r.rows.push(p); }
      if (p._inv && !p._inv._batal && (p.nominal > 0 || (p._pb && p.nominal))) p._inv._pays.push(p);
    });
    Object.keys(pbRef).forEach(function (k) {
      var L = pbRef[k], tot = sumBy(L, function (p) { return p.nominal; });
      L.forEach(function (p) { p._other = L.filter(function (q) { return q !== p && q._inv; }).map(function (q) { return q._inv.nomor.replace(/^KU\.02\.02\/WB-4C\./, ''); }).join(', '); });
      if (Math.abs(tot) > 1 || L.length < 2) L.forEach(function (p) { p._cek.push('⚠ PINDAH BUKU "' + k + '" TIDAK SEIMBANG (jumlah ' + fmtRp(tot) + '; harus ada baris − di invoice asal dan + di invoice tujuan)'); });
    });
    Object.keys(refSum).forEach(function (k) {
      var r = refSum[k];
      if (r.transfer > 0 && Math.abs(r.transfer - r.sum) > 1) r.rows.forEach(function (p) { p._cek.push('⚠ PECAHAN REF "' + k + '" = ' + fmtRp(r.sum) + ' ≠ TRANSFER ' + fmtRp(r.transfer)); });
    });
    pays.forEach(function (p) { if (!p._skip && p._cek.length) issues.push(['PEMBAYARAN', LOG.sheet, 'baris ' + p._row + (p.inv ? ' (' + p.inv + ')' : ''), p._cek.join('; '), 'Perbaiki baris di Log Penerimaan']); });

    // ---- persetujuan
    var appr = rowsToObjects(input.approvals || [], APPR.cols, APPR.firstRow).filter(function (r) { return !r._empty && S(r.id); });
    appr.forEach(function (r) {
      r.jenis = U(r.jenis); r.ref = S(r.ref); r.status = U(r.status);
      var u = userBy(params, r.putus_oleh); r.putus_nama = u ? u.nama : S(r.putus_oleh);
      var u2 = userBy(params, r.oleh); r.oleh_nama = u2 ? u2.nama : S(r.oleh);
      r.waktu = S(r.waktu); r.putus_waktu = S(r.putus_waktu);
      try { r._data = S(r.data) ? JSON.parse(S(r.data)) : null; } catch (e) { r._data = null; }
    });
    function latest(jenis, ref) { var L = appr.filter(function (r) { return r.jenis === jenis && r.ref === ref && r.status !== 'DIBATALKAN'; }); return L.length ? L[L.length - 1] : null; }
    // ---- alokasi & status per invoice
    var TOLR = N(params.TOL_PEMBULATAN), TOLL = N(params.TOL_LEBIH);
    invs.forEach(function (inv) {
      if (inv._skip) return;
      var a = inv._a, pel = inv._pel, pro = inv._pro;
      inv._pays.sort(function (x, y) { return (x.tgl || '9999') < (y.tgl || '9999') ? -1 : (x.tgl || '9999') > (y.tgl || '9999') ? 1 : x._row - y._row; });
      var al = allocate(inv, params), admMax = al.admMax, keP = al.keP, keN = al.keN, last = '', lunasFlag = false;
      inv._pays.forEach(function (p) { if (p.tgl && p.tgl > last) last = p.tgl; if (S(p.lunas)) lunasFlag = true; });
      var R = keP + keN, sisaBr = a.netto + a.ppn - R; // sisa terhadap tagihan sebelum PPh
      var res = { ke_ppn: keP, ke_net: keN, diterima: R, pot_pph: 0, adm: 0, sel_pph: 0, sisa: 0, tgl_bayar: last, status: '', closed: false };
      // tagihan retensi memakai TOP RETENSI pelanggan bila diisi (mis. PT Fajarputera Dinasti = 100 hari)
      var isRet = U(inv.jenis) === 'RETENSI' || /RETENSI/.test(U(inv.invke));
      var top = isNum(inv.top) ? inv.top : (isRet && pel && pel.top_ret) ? pel.top_ret : (pro ? pro.top : 30);
      var basisTT = pel && pel.dasar_jt === 'TGL TANDA TERIMA';
      var basis = basisTT ? inv.tt : inv.submit;
      res.jt = basis ? addDays(basis, top) : '';
      res.jt_akt = inv.tgl ? addDays(inv.tgl, N(params.JT_AKUNTANSI_HARI) || 30) : '';
      res.umur = inv.tgl ? Math.max(0, diffDays(today, inv.tgl)) : '';
      var stl = inv._settle, remPPN = a.ppn - keP;
      if (inv._batal) { res.status = 'BATAL'; res.closed = true; }
      else if (R > 0 && sisaBr <= a.pph + admMax + TOLR && sisaBr >= -TOLL) {
        res.closed = true;
        if (sisaBr >= a.pph) { res.pot_pph = a.pph; res.adm = sisaBr - a.pph; }
        else if (sisaBr >= 0) { res.pot_pph = sisaBr; }
        res.status = 'LUNAS';
      } else if (R > 0 && a.ret > 0 && a.retpos === 'AKHIR' && a.pphAlt !== a.pph && sisaBr <= a.pphAlt + admMax + TOLR && sisaBr >= a.pphAlt - TOLR) {
        res.closed = true; res.pot_pph = a.pph; res.sel_pph = a.pphAlt - a.pph; res.adm = Math.max(0, sisaBr - a.pphAlt);
        res.status = 'CLOSE – SELISIH PPh ' + fmtRp(res.sel_pph);
      } else if (sisaBr < -TOLL) {
        res.status = 'LEBIH BAYAR ' + fmtRp(-sisaBr); res.closed = true; res.sisa = sisaBr;
      } else if (R > 0) {
        res.sisa = Math.max(0, sisaBr - a.pph);
        if (stl && !stl.kurang && remPPN > TOLR) res.status = 'NETTO LUNAS – PPN BELUM DIBAYAR';
        else if (lunasFlag && S(inv.ba_no)) res.status = 'KURANG BAYAR – DITAGIH VIA BA/MEMO';
        else res.status = lunasFlag ? 'KURANG BAYAR ' + fmtRp(sisaBr - a.pph) : 'CICILAN BERJALAN';
      } else {
        res.sisa = a.kas;
        if (!inv.submit && !(basisTT && inv.tt)) res.status = 'WAITING SUBMIT';
        else if (basisTT && !inv.tt) res.status = 'BELUM KELUAR TANDA TERIMA';
        else if (res.jt && today > res.jt) res.status = 'OVERDUE';
        else if (res.jt && today === res.jt) res.status = 'JATUH TEMPO HARI INI';
        else res.status = 'BELUM JATUH TEMPO';
      }
      if (res.closed && !inv._batal && res.status.indexOf('LEBIH') !== 0) res.sisa = 0;
      if (!res.closed && (res.status === 'CICILAN BERJALAN' || res.status === 'NETTO LUNAS – PPN BELUM DIBAYAR') && res.jt && today > res.jt) res.status += ' (OVERDUE)';
      // penutupan manual: wajib alasan & persetujuan Section Head; sebelum disetujui piutang tetap terbuka
      inv._closedManual = false;
      if (!inv._batal && U(inv.tutup) === 'CLOSE MANUAL' && !res.closed) {
        inv._tutupSisa = res.sisa; inv._tutupFp = hash(['TUTUP', normInv(inv.nomor), rnd(res.sisa), U(inv.alasan_tutup)].join('|'));
        var tr = latest(DOC.TUTUP, normInv(inv.nomor)), ok = tr && tr.status === 'DISETUJUI' && S(tr.sidik) === inv._tutupFp;
        if (!S(inv.alasan_tutup)) { res.status = 'CLOSE MANUAL – ALASAN BELUM DIISI'; issues.push(['TUTUP MANUAL', REG.sheet, inv.nomor, 'TUTUP MANUAL diisi tanpa alasan', 'Isi ALASAN TUTUP MANUAL, lalu ajukan penutupan di aplikasi (Persetujuan ▸ Penutupan manual)']); }
        else if (ok) { res.status = 'CLOSE MANUAL'; res.closed = true; res.sisa = 0; inv._closedManual = true; res.tgl_tutup = D(tr.putus_waktu); res.tutup_sisa = inv._tutupSisa; }
        else { res.status = 'CLOSE MANUAL – ' + (tr && tr.status === 'MENUNGGU' && S(tr.sidik) === inv._tutupFp ? 'MENUNGGU PERSETUJUAN' : 'BELUM DIAJUKAN'); issues.push(['TUTUP MANUAL', REG.sheet, inv.nomor, 'Penutupan manual (sisa ' + fmtRp(res.sisa) + ') ' + (/MENUNGGU/.test(res.status) ? 'menunggu persetujuan Section Head' : 'belum diajukan'), 'Ajukan / ingatkan di aplikasi ▸ Persetujuan ▸ Penutupan manual. Sebelum disetujui piutang tetap terbuka.']); }
      }
      // invoice lunas tetapi porsi Netto belum pernah "selesai" (mis. alokasi manual) → PPh diakui pada pembayaran terakhir
      if (res.closed && !inv._batal && /^(LUNAS|CLOSE –)/.test(res.status) && !inv._settle && inv._pays.length) {
        var lp = inv._pays[inv._pays.length - 1];
        lp._pph = res.pot_pph; lp._adm = res.adm; lp._sel = res.sel_pph;
        inv._settle = { i: inv._pays.length - 1, tgl: lp.tgl, pot: res.pot_pph, adm: res.adm, sel: res.sel_pph };
      }
      if (/^KURANG BAYAR \d/.test(res.status)) issues.push(['STATUS', REG.sheet, inv.nomor, res.status + ' (pelunasan sudah ditandai)', 'Bila kekurangan ditagih lewat berita acara / memo: isi NO. BA / MEMO. Bila dikompensasi invoice lain: catat PINDAH BUKU. Bila direlakan: TUTUP MANUAL + alasan.']);
      if (res.status === 'KURANG BAYAR – DITAGIH VIA BA/MEMO' && /BELUM DIISI/i.test(S(inv.ba_no))) issues.push(['STATUS', REG.sheet, inv.nomor, 'Nomor BA / memo penagihan kekurangan belum diisi', 'Isi NO. BA / MEMO dan TGL BA / MEMO']);
      if (res.status.indexOf('LEBIH BAYAR') === 0) issues.push(['STATUS', REG.sheet, inv.nomor, res.status, 'Bila kelebihan dipakai untuk invoice lain, catat PINDAH BUKU (aplikasi ▸ Pembayaran ▸ Pindah buku); bila salah input, perbaiki Log Penerimaan']);
      res.rencana_net = res.closed ? 0 : Math.max(0, a.netto - a.pph - keN);
      res.rencana_kas = res.closed ? 0 : Math.max(0, res.sisa);
      inv._res = res;
      inv._fp = fpInvoice(inv);
    });

    var payGroups = {};
    pays.forEach(function (p) { if (!p._skip && p._inv) (payGroups[payKey(p)] = payGroups[payKey(p)] || []).push(p); });
    var VI = valIndex(appr), valMap = VI.valMap; // nomor → keputusan Section Head yang berlaku (v2.5: per invoice)

    // ---- tagbrut & data LK
    // tagbrut AR hanya proyek jasa (NPP di Master Proyek); baris NPP lain (material — ranah Akuntansi/P&E) diabaikan
    var tagAll = rowsToObjects(input.tagbrut || [], TAG.cols, TAG.firstRow).filter(function (r) { return !r._empty && S(r.npp); }).map(function (r) {
      return { npp: S(r.npp), tgl: D(r.tgl), nodok: S(r.nodok), nilai: N(r.nilai), sumber: S(r.sumber), ket: S(r.ket), _row: r._row };
    });
    var tagRows = tagAll.filter(function (r) { return !!M.pro[r.npp]; }), tagSkip = tagAll.length - tagRows.length;
    tagRows.forEach(function (r) { if (!r.tgl) issues.push(['TAGBRUT', TAG.sheet, 'baris ' + r._row, 'Tanggal pengakuan kosong', 'Isi tanggal']); });
    var notes = rowsToObjects(input.notes || [], NOTE.cols, NOTE.firstRow).filter(function (r) { return !r._empty && S(r.catatan); }).map(function (r) {
      return { npp: U(r.npp), tgl: D(r.tgl), pic: S(r.pic), catatan: S(r.catatan) };
    });
    var lk = parseLK(input.lk || {});

    var ctx = { params: params, today: today, M: M, invs: invs.filter(function (i) { return !i._skip; }), pays: pays.filter(function (p) { return !p._skip; }),
      surat: rowsToObjects(input.surat || [], SURAT.cols, SURAT.firstRow).filter(function (r) { return !r._empty && S(r.nomor); }).map(function (r) { r.tgl = D(r.tgl); r.tenggat = D(r.tenggat); r.pelanggan = U(r.pelanggan); return r; }),
      _input: input, issues: issues, sel: input.selectors || {}, appr: appr, latest: latest, payGroups: payGroups, valMap: valMap, valDec: VI.byRef, valFirst: VI.first, tag: tagRows, tagSkip: tagSkip, notes: notes, lk: lk };

    // ---- memorial akuntansi (v2.4): blok baru / revisi / koreksi dihitung dari data + sheet Memorial AR
    if (!input.light && input.memo !== undefined) { try { ctx.memo = memorial(ctx, input.memo); } catch (e) { issues.push(['MEMORIAL', MEMO.sheet, '-', 'Memorial gagal dihitung: ' + e.message, 'Laporkan ke pengembang']); ctx.memo = null; } }
    var mst = ctx.memo ? ctx.memo.status : {};
    // ---- kolom otomatis register & log
    var regOut = invs.map(function (inv, idx) {
      if (inv._skip) return null;
      var a = inv._a, r = inv._res;
      var bi = inv._batalInfo, stDok = inv._batal ? 'BATAL' + (bi ? ' · ' + bi.tgl.split('-').reverse().join('-') + ' · ' + bi.id : '') : (inv._batalPend ? 'AKTIF · ⏳ pembatalan menunggu (' + inv._batalPend + ')' : 'AKTIF');
      return { status_inv: stDok, no: idx + 1, proyek: inv._pro ? inv._pro.singkat : S(inv.proyek), total: a.total, netto: a.netto, pph: a.pph, kas: a.kas, jt: r.jt, jt_akt: r.jt_akt,
        ke_ppn: r.ke_ppn, ke_net: r.ke_net, diterima: r.diterima, pot_pph: r.pot_pph, adm: r.adm, sel_pph: r.sel_pph, sisa: r.sisa,
        tgl_bayar: r.tgl_bayar, status: r.status, umur: r.closed ? '' : r.umur, validasi_ok: valMap[normInv(inv.nomor)] || '',
        persetujuan: inv._batal ? '' : apprStatusText(latest(DOC.INV, normInv(inv.nomor)), inv._fp), memorial: mst['INV|' + normInv(inv.nomor)] || '' };
    });
    var n = 0;
    var logOut = pays.map(function (p) {
      if (p._skip) return null;
      n++;
      var g = payGroups[payKey(p)], st = g ? apprStatusText(latest(DOC.PAY, payKey(p)), fpPayGroup(g)) : '';
      return { no: n, idp: p.idp, pelanggan: p._inv ? p._inv.pelanggan : '', npp: p._inv ? p._inv.npp : '', ke_ppn: p._inv ? (p._keP || 0) : '', ke_net: p._inv ? (p._keN || 0) : '',
        aturan: p._rule || '', pph_cat: p._pph || p._sel ? (p._pph || 0) + (p._sel || 0) : '', adm_cat: p._adm || '', cicil: p._ke || '', persetujuan: st,
        cek: p._cek.length ? p._cek.join('; ') : '✓ OK', memorial: mst[p._pb ? 'PB|' + S(p.ref) : 'TRM|' + p.idp] || '' };
    });
    if (input.light) return { regOut: regOut, logOut: logOut, reports: {}, issues: issues, params: params, master: M, today: today, ctx: ctx };
    // ---- cek data tambahan (faktur, SAP, persetujuan, parameter)
    extraChecks(ctx);
    peCheck(ctx, input.peref);
    if (ctx.memo) ctx.memo.belum.forEach(function (b) { if (b.tgl && diffDays(today, b.tgl) > 7) issues.push(['MEMORIAL', MEMO.sheet, b.kode, 'Memorial belum diverifikasi Akuntansi (' + diffDays(today, b.tgl) + ' hari)', 'Ingatkan Mbak Elmarta mencentang VERIF di file Memorial AR']); });
    var reports = {
      'Pareto': repPareto(ctx), 'Rekap Invoice Keluar': repRekapKeluar(ctx), 'Kartu Piutang': repKartu(ctx),
      'Monitoring UM & Retensi': repUmRet(ctx), 'INKASO': repInkaso(ctx), 'RA-RI Termin': repRARI(ctx), 'Aging Piutang': repAging(ctx),
      'Monitoring Tagbrut': repTagbrut(ctx), 'CKPN': repCkpn(ctx), 'Usulan Jurnal SAP': repJurnal(ctx), 'Rekon SAP': repRekon(ctx)
    };
    try { reports['Kinerja & Peringatan Dini'] = repKinerja(ctx); } catch (e) { issues.push(['KINERJA', 'Kinerja & Peringatan Dini', '-', 'Gagal dihitung: ' + e.message, 'Laporkan ke pengembang']); }
    reports['Dashboard Summary'] = repDashboard(ctx);
    reports['Cek Data'] = repCek(ctx);
    return { regOut: regOut, logOut: logOut, reports: reports, issues: issues, params: params, master: M, today: today, ctx: ctx };
  }

  // ------------------------------------------------------------------ PEMILIH BULAN & BLOK LAPORAN
  function parseMonthSel(v, today, allowAll) {
    var s = S(v), y = +today.substring(0, 4), m = +today.substring(5, 7);
    if (allowAll && U(s) === 'SEMUA BULAN') return { all: true, y: y, m: m, start: monthStart(y, 1), end: monthEnd(y, 12), label: 'Semua bulan ' + y };
    for (var i = 0; i < 12; i++) {
      var re = new RegExp('^' + MONTHS[i] + '\\s+(\\d{4})$', 'i'), mm = s.match(re);
      if (mm) return { y: +mm[1], m: i + 1, start: monthStart(+mm[1], i + 1), end: monthEnd(+mm[1], i + 1), label: monthLabel(+mm[1], i + 1) };
    }
    var mm2 = s.match(/^(\d{1,2})[\/\-\.](\d{4})$/), mm3 = s.match(/^(\d{4})-(\d{2})$/);
    if (mm3) { y = +mm3[1]; m = +mm3[2]; }
    else if (mm2 && +mm2[1] >= 1 && +mm2[1] <= 12) { y = +mm2[2]; m = +mm2[1]; }
    else if (Object.prototype.toString.call(v) === '[object Date]' || /^\d{4}-\d{2}-\d{2}/.test(s) || /^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4}$/.test(s)) { var d = D(v) || D(s); if (d) { y = +d.substring(0, 4); m = +d.substring(5, 7); } }
    else { var nm = U(s).replace(/\s+/g, ' '); for (var k = 0; k < 12; k++) { var ab = MONTHS[k].substring(0, 3).toUpperCase(), r2 = nm.match(new RegExp('^' + ab + '[A-Z]*\\.?\\s*(\\d{4})$')); if (r2) { y = +r2[1]; m = k + 1; break; } } }
    if (S(v) && U(s).indexOf('OTOMATIS') < 0) return { y: y, m: m, start: monthStart(y, m), end: monthEnd(y, m), label: monthLabel(y, m) };
    return { y: y, m: m, start: monthStart(y, m), end: monthEnd(y, m), label: monthLabel(y, m) + ' (otomatis)', auto: true };
  }
  /** pemilih "per akhir bulan" (CKPN, tagbrut): OTOMATIS = akhir bulan lalu */
  function parseCutoff(v, today) {
    var sel = parseMonthSel(v, today, false);
    if (sel.auto) { var pe = prevMonthEnd(today), y = +pe.substring(0, 4), m = +pe.substring(5, 7); return { y: y, m: m, start: monthStart(y, m), end: pe, label: monthLabel(y, m) + ' (otomatis: akhir bulan lalu)' }; }
    return sel;
  }
  function block(r, c, rows, opt) { opt = opt || {}; return { r: r, c: c || 1, rows: rows, styles: opt.styles || [], fmt: opt.fmt || {}, edit: opt.edit || null, merge: opt.merge || [], span: opt.span || 1, noWidth: !!opt.noWidth, group: !!opt.group }; }
  function sty(rowOff, s, n) { return { r: rowOff, s: s, n: n || 1 }; }
  function openInv(ctx) { return ctx.invs.filter(function (i) { return !i._batal && !i._res.closed; }); }
  function planDate(inv) { return inv.rencana || inv._res.jt || ''; }
  function invLabel(inv) { return S(inv.invke) || S(inv.jenis) || 'Tagihan'; }
  function uraianOf(inv) {
    if (inv.uraian) return inv.uraian;
    var p = inv._pro; return (S(inv.invke) ? 'Tagihan ' + S(inv.invke) + ' ' : '') + (p ? S(p.deskripsi) : '');
  }

  // ------------------------------------------------------------------ DATA LK (TEMPEL)
  /** pu: lampiran 1104111000/1104211000 kolom B:O · ppn: 1104112000/1104212000 kolom B:P · ret: 1105111000/1105211000 kolom B:N */
  function parseLK(lk) {
    var out = { periode: D(lk.periode), pu: [], ppn: [], ret: [] };
    function walk(rows, map, target) {
      var pel = '', npp = '', bp = '';
      (rows || []).forEach(function (r) {
        if (!r) return;
        var c1 = S(r[1]);
        if (/^(SUB|GRAND|JUMLAH|TOTAL)/i.test(c1) || /^(SUB|GRAND)/i.test(S(r[0]))) { pel = ''; npp = ''; bp = ''; return; }
        if (c1) pel = U(c1); if (S(r[map.npp])) npp = nppKey(r[map.npp]); if (map.bp !== undefined && S(r[map.bp])) bp = S(r[map.bp]);
        var doc = S(r[map.doc]), jml = N(r[map.jml]);
        if (!doc || !jml) return;
        target.push({ pelanggan: pel, npp: npp, bp: bp, tgl: D(r[map.tgl]), doc: doc.replace(/\.0$/, ''), jumlah: rnd(jml), saldo: rnd(N(r[map.saldo])), faktur: map.fp !== undefined ? digits(r[map.fp]) : '' });
      });
    }
    walk(lk.pu, { npp: 4, bp: 2, tgl: 6, doc: 7, jml: 8, saldo: 13 }, out.pu);
    walk(lk.ppn, { npp: 4, bp: 2, fp: 5, tgl: 7, doc: 8, jml: 9, saldo: 14 }, out.ppn);
    walk(lk.ret, { npp: 3, bp: 2, tgl: 5, doc: 6, jml: 7, saldo: 11 }, out.ret);
    return out;
  }

  // ------------------------------------------------------------------ KARTU PIUTANG (BUKU BESAR PEMBANTU)
  function kartuData(ctx, npp, asOf) {
    ctx = contextAt(ctx, asOf);
    var pro = ctx.M.pro[npp], C = asOf || ctx.today;
    if (!pro) return null;
    var pel = ctx.M.pel[pro.pelanggan] || {};
    // NPP + pelanggan: NPP sementara yang (keliru) dipakai dua proyek tidak membuat kartu tercampur
    var L = ctx.invs.filter(function (i) { return i.npp === npp && !i._batal && (!i.tgl || i.tgl <= C) && (!pro.pelanggan || i.pelanggan === pro.pelanggan); })
      .sort(function (a, b) { return (a.tgl || '') < (b.tgl || '') ? -1 : (a.tgl || '') > (b.tgl || '') ? 1 : (a.nomor < b.nomor ? -1 : 1); });
    var groups = [], T = { tag: 0, debit: 0, kredit: 0, kas: 0, pot: 0, pb: 0, kasN: 0, kasP: 0, pph: 0, adm: 0, adj: 0, saldo: 0, sN: 0, sP: 0, dpp: 0, ppn: 0, brutoProg: 0, umk: 0, ret: 0 };
    L.forEach(function (i) {
      // debit/kredit per komponen (Netto, PPN, PPh, adm), tanpa saldo berjalan; tiap invoice ditutup baris "Sisa tagihan"
      var a = i._a, rows = [], sN = a.netto, sP = a.ppn, sh = i.nomor;
      var desc = 'Tagihan ' + invLabel(i) + ' — ' + sh + ' (DPP ' + fmtRp(a.dpp) + ' + PPN ' + fmtRp(a.ppn) + (a.retpos === 'AKHIR' && a.ret ? '; retensi ' + fmtRp(a.ret) + ' ditahan' : '') + ')';
      var row = function (o) { o.nomor = i.nomor; o.debit = o.debit || 0; o.kredit = o.kredit || 0; o.saldo = o.saldo === undefined ? '' : o.saldo; o.ket = o.ket || ''; o.ref = o.ref || ''; o.sap = o.sap || ''; rows.push(o); return o; };
      row({ kind: 'inv', tgl: i.tgl, uraian: desc, ref: S(i.no_fp) ? 'FP ' + i.no_fp : '', sap: i.sap_doc, debit: a.netto + a.ppn });
      var nN = i._pays.filter(function (p) { return p._keN > 0 && !p._pb && p.tgl && p.tgl <= C; }).length, kN = 0, pphNanti = true;
      i._pays.forEach(function (p) {
        if (!p.tgl || p.tgl > C) return;
        var ref = S(p.ref) && U(p.ref) !== 'MIGRASI' ? S(p.ref) : (p._migr ? 'data lama' : (p._rek ? p._rek.bank + ' ' : '') + 'mutasi ' + p.tgl.split('-').reverse().join('-'));
        var bk = p._rek ? ' · ' + p._rek.bank : '', nw = !!p._new, o = function (x) { x.tgl = p.tgl; x.sap = p.sap_doc; x.baru = nw; if (x.ref === undefined) x.ref = ref; return row(x); };
        if (p._pb) {
          if (p._keN < 0) { sN -= p._keN; o({ kind: 'pb', uraian: 'Kelebihan bayar dialihkan ke ' + (p._other || 'invoice lain'), debit: -p._keN }); }
          if (p._keN > 0) { sN -= p._keN; o({ kind: 'pb', uraian: 'Diterima melalui pindah buku dari ' + (p._other || 'invoice lain') + ' (Netto)', kredit: p._keN }); }
          if (p._keP > 0) { sP -= p._keP; o({ kind: 'pb', uraian: 'Diterima melalui pindah buku dari ' + (p._other || 'invoice lain') + ' (PPN)', kredit: p._keP }); }
        } else {
          if (p._keN > 0) { kN++; sN -= p._keN; o({ kind: 'pay', uraian: 'Penerimaan Netto' + bk + (nN > 1 ? ' (cicilan ke-' + kN + ')' : ''), kredit: p._keN }); }
          if (p._keP > 0) { sP -= p._keP; o({ kind: 'pay', uraian: 'Penerimaan PPN' + bk, kredit: p._keP }); }
        }
        if (p._pph) { sN -= p._pph; pphNanti = false; o({ kind: 'pot', uraian: 'Dipotong ' + pphLabel(i.npp, a.tarif), ref: S(p.bupot) ? 'Bupot ' + S(p.bupot) : 'bukti potong', kredit: p._pph }); }
        if (p._sel) { sN -= p._sel; pphNanti = false; o({ kind: 'pot', uraian: 'Selisih PPh (dasar potong pelanggan berbeda)', ref: S(p.bupot) ? 'Bupot ' + S(p.bupot) : 'bukti potong', kredit: p._sel }); }
        if (p._adm) { sN -= p._adm; o({ kind: 'pot', uraian: 'Biaya administrasi bank', kredit: p._adm }); }
      });
      if (i._closedManual && (!i._res.tgl_tutup || i._res.tgl_tutup <= C) && Math.abs(sN + sP) > 0) {
        var k = sN + sP; sN = 0; sP = 0;
        row({ kind: 'adj', tgl: i._res.tgl_tutup || '', uraian: 'Sisa ditutup manual (disetujui Section Head)', kredit: k, ket: S(i.alasan_tutup) });
      }
      var st0 = invAt(i, C, ctx.params), status = C >= ctx.today ? i._res.status : (st0.saldoN + st0.saldoP > 0 ? 'TERBUKA per ' + tglPanjang(C) : 'LUNAS per ' + tglPanjang(C));
      var sisa = Math.abs(sN + sP) <= 1 ? 0 : sN + sP, ket;
      if (!sisa) ket = /^(LUNAS|CLOSE)/.test(status) ? status.replace(/ per .*$/, '') : 'LUNAS';
      else if (/BA\/MEMO/.test(status)) ket = 'Kurang bayar — ditagih via berita acara / memo ' + S(i.ba_no) + (D(i.ba_tgl) ? ' tgl ' + tglPanjang(D(i.ba_tgl)) : '') + '; tetap piutang sampai dibayar';
      else if (/PPN BELUM/.test(status)) ket = 'Netto sudah lunas — PPN Rp ' + fmtRp(sP) + ' belum dibayar' + (/OVERDUE/.test(status) ? ' (lewat jatuh tempo)' : '');
      else if (sisa < 0) ket = 'Lebih bayar Rp ' + fmtRp(-sisa);
      else {
        var kasSisa = Math.max(0, sisa - (pphNanti ? a.pph : 0)), lbl = /CICILAN/.test(status) ? 'Cicilan berjalan' : status.charAt(0) + status.slice(1).toLowerCase();
        ket = lbl + ' — akan dibayar pelanggan ' + fmtRp(kasSisa) + (pphNanti && a.pph ? '; PPh ' + fmtRp(a.pph) + ' dipotong saat pelunasan' : '');
      }
      row({ kind: 'sisa', tgl: '', uraian: 'Sisa tagihan ' + sh, saldo: sisa, ket: ket });
      // bagian A: rincian tagihan (nilai progres asli, potongan UM, retensi, DPP, PPN)
      var prog = a.jenis === 'PROGRES' || a.jenis === 'LAINNYA', pp = pro || {};
      var lama = prog && !(N(i.bruto) > 0) && ((pp.um > 0 || pp.um_awal > 0 || pp.ret_pct > 0) || /^MIGRASI/.test(U(i.sumber)) && (a.umk || a.ret));
      var A = { nomor: i.nomor, tgl: i.tgl, invke: invLabel(i), jenis: a.jenis, progres: prog ? a.bruto : 0, umk: a.umk, ret: a.ret, retpos: a.retpos, dpp: a.dpp, ppn: a.ppn,
        total: a.total, dicatat: a.netto + a.ppn, lama: !!lama, fp: S(i.no_fp) };
      groups.push({ inv: i, rows: rows, status: status, saldo: sN + sP, A: A });
      T.debit += a.netto + a.ppn; T.tag += a.netto + a.ppn; T.dpp += a.dpp; T.ppn += a.ppn;
      rows.forEach(function (r) { if (r.kind === 'inv' || r.kind === 'sisa') return; T.kredit += r.kredit; T.debit += r.debit;
        if (r.kind === 'pay') T.kas += r.kredit; else if (r.kind === 'pb') T.pb += r.kredit - r.debit; else T.pot += r.kredit; });
      T.saldo += sN + sP; T.sN += sN; T.sP += sP;
      if (a.jenis === 'PROGRES' || a.jenis === 'LAINNYA') T.brutoProg += a.bruto;
      T.umk += a.umk;
    });
    var k = kontrakAt(pro, C), ret = retensiRows(ctx).filter(function (r) { return r[1] === npp; })[0];
    T.ret = ret ? ret[9] : 0;
    return { pro: pro, pel: pel, asOf: C, kontrak: k, groups: groups, T: T, sisaKontrak: Math.max(0, k.net - T.brutoProg), pctTagih: k.net ? T.brutoProg / k.net : 0 };
  }

  // ------------------------------------------------------------------ USULAN JURNAL SAP
  function jl(ctx, kode, berelasi, pk, debit, kredit, extra) {
    var o = { kode: kode, gl: gl(ctx.params, kode, berelasi), nama: glName(ctx.params, kode), pk: pk, debit: rnd(debit || 0), kredit: rnd(kredit || 0) };
    Object.keys(extra || {}).forEach(function (k) { o[k] = extra[k]; });
    return o;
  }
  function jInvoice(ctx, inv) {
    var a = inv._a, b = inv._berelasi, pel = inv._pel || {}, pro = inv._pro || {}, x = { bp: pel.kode_bp || '', npp: nppKey(inv.npp) }, L = [];
    var cust = function (kode, d, k) { var o = jl(ctx, kode, b, d ? '01' : '15', d, k, x); o.cust = true; return o; };
    var akunPeng = pro.akun_peng === 'HAD' ? 'HAD' : 'TAGBRUT';
    if (a.jenis === 'UANG MUKA') {
      L.push(cust('PIUTANG_USAHA', a.dpp, 0));
      if (a.ppn) L.push(jl(ctx, 'PIUTANG_PPN', b, '40', a.ppn, 0, x));
      L.push(jl(ctx, 'UANG_MUKA_PROYEK', b, '50', 0, a.dpp, x));
    } else if (a.jenis === 'RETENSI') {
      L.push(cust('PIUTANG_USAHA', a.netto, 0));
      if (a.ppn) L.push(jl(ctx, 'PIUTANG_PPN', b, '40', a.ppn, 0, x));
      L.push(jl(ctx, 'PIUTANG_RETENSI', b, '50', 0, a.netto, x));
    } else {
      L.push(cust('PIUTANG_USAHA', a.netto, 0));
      if (a.ret) L.push(jl(ctx, 'PIUTANG_RETENSI', b, '40', a.ret, 0, x));
      if (a.umk) L.push(jl(ctx, 'UANG_MUKA_PROYEK', b, '40', a.umk, 0, x));
      if (a.ppn) L.push(jl(ctx, 'PIUTANG_PPN', b, '40', a.ppn, 0, x));
      L.push(jl(ctx, akunPeng, b, '50', 0, a.netto + a.ret + a.umk, x));
    }
    if (a.ppn) L.push(jl(ctx, 'PPN_KELUARAN', b, '50', 0, a.ppn, x));
    var teks = ('TAGIHAN ' + invLabel(inv) + ' ' + (pro.singkat || inv.npp)).toUpperCase().substring(0, 50);
    L.forEach(function (l) { l.teks = teks; });
    return { kejadian: 'Invoice keluar', tgl: inv.tgl, ref: inv.nomor, tcode: 'FBVB', doctype: 'DR', sap: inv.sap_doc, header: S(inv.no_fp), lines: L, inv: inv };
  }
  function jPayment(ctx, p) {
    var inv = p._inv, a = inv._a, b = inv._berelasi, pel = inv._pel || {}, x = { bp: pel.kode_bp || '', npp: nppKey(inv.npp) }, L = [];
    if (p._pb) { if (p._keN < 0) { var o0 = jl(ctx, 'PIUTANG_USAHA', b, '01', -p._keN, 0, x); o0.cust = true; L.push(o0); } }
    else L.push(jl(ctx, p._rek && ctx.params.akun[p._rek.akun] ? p._rek.akun : 'BANK', b, '40', p.nominal, 0, x));
    if (p._pph || p._sel) L.push(jl(ctx, pphGlKode(inv.npp), b, '40', (p._pph || 0) + (p._sel || 0), 0, x));
    if (p._adm) L.push(jl(ctx, 'BEBAN_ADM_BANK', b, '40', p._adm, 0, x));
    if (p._keP) L.push(jl(ctx, 'PIUTANG_PPN', b, '50', 0, p._keP, x));
    var cu = Math.max(0, p._keN || 0) + (p._pph || 0) + (p._sel || 0) + (p._adm || 0);
    if (cu) { var o = jl(ctx, 'PIUTANG_USAHA', b, '15', 0, cu, x); o.cust = true; L.push(o); }
    var teks = ((p._pb ? 'PINDAH BUKU ' : 'TERIMA ') + invLabel(inv) + ' ' + (inv._pro ? inv._pro.singkat : inv.npp)).toUpperCase().substring(0, 50);
    L.forEach(function (l) { l.teks = teks; });
    if (p._pb) return { kejadian: 'Pindah buku', tgl: p.tgl, ref: inv.nomor + ' · ' + S(p.ref), tcode: 'FB01', doctype: 'SA', sap: p.sap_doc, lines: L, pay: p };
    return { kejadian: 'Penerimaan' + (p._rek ? ' ' + p._rek.bank : ''), tgl: p.tgl, ref: inv.nomor + ' · ' + p.idp + (S(p.ref) && U(p.ref) !== 'MIGRASI' ? ' · ' + S(p.ref) : ''), tcode: 'F-28', doctype: 'DZ', sap: p.sap_doc, lines: L, pay: p };
  }
  /** jurnal untuk sekumpulan baris penerimaan: baris PINDAH BUKU dengan REF sama digabung menjadi satu jurnal (seimbang) */
  function jPays(ctx, pays) {
    var out = [], pb = {};
    pays.forEach(function (p) {
      if (!p._pb) { out.push(jPayment(ctx, p)); return; }
      var j = jPayment(ctx, p), k = S(p.ref);
      if (!pb[k]) { pb[k] = { kejadian: 'Pindah buku / kompensasi', tgl: p.tgl, ref: S(p.ref) + ' · ' + p._inv.nomor, tcode: 'FB01', doctype: 'SA', sap: p.sap_doc, lines: [], pay: p }; out.push(pb[k]); }
      else pb[k].ref += ' → ' + p._inv.nomor;
      pb[k].lines = pb[k].lines.concat(j.lines);
    });
    return out;
  }

  // ------------------------------------------------------------------ RETENSI
  function retensiRows(ctx) {
    if (ctx._retRows) return ctx._retRows;
    var out = [];
    Object.keys(ctx.M.pro).sort().forEach(function (npp) {
      var p = ctx.M.pro[npp], L = ctx.invs.filter(function (i) { return i.npp === npp && !i._batal; });
      var awal = sumBy(L, function (i) { return i._a.retpos === 'AWAL' && i._a.jenis !== 'RETENSI' ? i._a.ret : 0; });
      var akhir = sumBy(L, function (i) { return i._a.retpos === 'AKHIR' && i._a.jenis !== 'RETENSI' ? i._a.ret : 0; });
      var ditagih = sumBy(L, function (i) { return i._a.jenis === 'RETENSI' ? i._a.dpp : 0; });
      if (!awal && !akhir && !ditagih && !p.ret_pct) return;
      out.push([p.pelanggan, npp, p.singkat, p.ret_pos, p.ret_pct || '', awal, akhir, awal + akhir, ditagih, Math.max(0, awal + akhir - ditagih),
        (p.ret_pct && !awal && !akhir) ? 'Retensi invoice lama belum tercatat — isi kolom RETENSI di Monitoring Invoice' :
          (ditagih > awal + akhir + 1 ? 'Retensi ditagih melebihi yang tercatat ditahan — lengkapi % RETENSI / kolom RETENSI invoice lama' : '')]);
    });
    ctx._retRows = out;
    return out;
  }
  /** potongan retensi yang masih tertahan per tanggal (FIFO; invoice retensi melepas yang tertua) */
  function retPieces(ctx, C) {
    var out = [];
    Object.keys(ctx.M.pro).forEach(function (npp) {
      var L = ctx.invs.filter(function (i) { return i.npp === npp && !i._batal && i.tgl && i.tgl <= C; });
      var pcs = L.filter(function (i) { return i._a.jenis !== 'RETENSI' && i._a.ret > 0; }).sort(function (a, b) { return a.tgl < b.tgl ? -1 : 1; })
        .map(function (i) { return { inv: i, tgl: i.tgl, saldo: i._a.ret }; });
      var rel = sumBy(L, function (i) { return i._a.jenis === 'RETENSI' ? i._a.dpp : 0; });
      pcs.forEach(function (pc) { var t = Math.min(pc.saldo, rel); pc.saldo -= t; rel -= t; if (pc.saldo > 0) out.push(pc); });
    });
    return out;
  }

  // ------------------------------------------------------------------ TAGBRUT / HAD
  function tagAt(ctx, C) {
    var seed = D(ctx.params.TAGBRUT_SALDO_AWAL_PER) || '0000-00-00', by = {};
    ctx.tag.forEach(function (r) { if (!r.tgl || r.tgl > C) return; (by[r.npp] = by[r.npp] || { lines: [], inv: [] }).lines.push({ npp: r.npp, tgl: r.tgl, nodok: r.nodok, nilai: r.nilai, sumber: r.sumber, ket: r.ket }); });
    ctx.invs.forEach(function (i) {
      if (i._batal || !i.tgl || i.tgl <= seed || i.tgl > C) return;
      if (i._a.jenis !== 'PROGRES' && i._a.jenis !== 'LAINNYA') return;
      if (!by[i.npp] && !ctx.tag.some(function (r) { return r.npp === i.npp; })) return; // proyek tanpa data pengakuan tidak dimonitor
      (by[i.npp] = by[i.npp] || { lines: [], inv: [] }).inv.push(i);
    });
    var res = {};
    Object.keys(by).forEach(function (npp) {
      var o = by[npp], pos = o.lines.filter(function (l) { return l.nilai > 0; }).sort(function (a, b) { return a.tgl < b.tgl ? -1 : 1; });
      var K = sumBy(o.lines, function (l) { return l.nilai < 0 ? -l.nilai : 0; }) + sumBy(o.inv, function (i) { return i._a.bruto; });
      var peng = sumBy(pos, function (l) { return l.nilai; }), K0 = K;
      var rem = pos.map(function (l) { var t = Math.min(l.nilai, K); K -= t; return { npp: npp, tgl: l.tgl, nodok: l.nodok, nilai: l.nilai, saldo: l.nilai - t, sumber: l.sumber }; }).filter(function (l) { return l.saldo > 0.5; });
      var pro = ctx.M.pro[npp] || {}, pel = ctx.M.pel[pro.pelanggan] || {};
      rem.forEach(function (l) { l.umur = Math.max(0, diffDays(C, l.tgl)); });
      res[npp] = { npp: npp, pro: pro, pel: pel, akun: pro.akun_peng || 'TAGBRUT', berelasi: !!pel.berelasi, pengakuan: peng, ditagih: sumBy(o.inv, function (i) { return i._a.bruto; }),
        lines: rem, sisa: sumBy(rem, function (l) { return l.saldo; }), hdd: K > 0.5 ? K : 0, oldest: rem.length ? rem[0].tgl : '', inv: o.inv, consumed: K0 };
    });
    return res;
  }

  // ------------------------------------------------------------------ CKPN
  function kelasOf(ctx, umur) {
    if (umur < 0) return { kelas: 'Belum JT', tarif: 0 };
    for (var i = 0; i < ctx.params.ckpn.length; i++) { var k = ctx.params.ckpn[i]; if (umur >= k.dari && umur <= k.sampai) return { kelas: k.kelas, tarif: k.tarif }; }
    var last = ctx.params.ckpn[ctx.params.ckpn.length - 1]; return { kelas: last.kelas, tarif: last.tarif };
  }
  function npv(saldo, umur, r) { return saldo / Math.pow(1 + r / 365, Math.max(0, umur)); }
  function ckpnAt(ctx, C) {
    var key = C; ctx._ck = ctx._ck || {}; if (ctx._ck[key]) return ctx._ck[key];
    var P = ctx.params, rel = U(P.CKPN_PIHAK_BERELASI) === 'YA', r = N(P.CKPN_DISKONTO) || 0.0833, items = [];
    ctx.invs.forEach(function (i) {
      if (i._batal || !i.tgl || i.tgl > C) return;
      var st = invAt(i, C, P); if (st.saldoN <= 0) return;
      var umur = diffDays(C, i._res.jt_akt), k = kelasOf(ctx, umur), tarif = i._berelasi && !rel ? 0 : k.tarif;
      items.push({ key: 'PU|' + normInv(i.nomor), akun: 'PIUTANG_USAHA', berelasi: i._berelasi, pelanggan: i.pelanggan, bp: i._pel ? i._pel.kode_bp : '', npp: i.npp,
        ref: i.nomor, doc: i.sap_doc, tgl: i.tgl, jumlah: i._a.netto, saldo: st.saldoN, jt: i._res.jt_akt, umur: umur, kelas: k.kelas + (i._berelasi && !rel ? ' (berelasi)' : ''), tarif: tarif, ckpn: rnd(st.saldoN * tarif) });
    });
    retPieces(ctx, C).forEach(function (pc) {
      var i = pc.inv, umur = Math.max(0, diffDays(C, pc.tgl)), b = i._berelasi && !rel, v = npv(pc.saldo, umur, r);
      items.push({ key: 'RT|' + normInv(i.nomor), akun: 'PIUTANG_RETENSI', berelasi: i._berelasi, pelanggan: i.pelanggan, bp: i._pel ? i._pel.kode_bp : '', npp: i.npp,
        ref: i.nomor, doc: i.sap_doc, tgl: pc.tgl, jumlah: i._a.ret, saldo: pc.saldo, umur: umur, npv: b ? pc.saldo : v, ckpn: b ? 0 : rnd(pc.saldo - v) });
    });
    var tg = tagAt(ctx, C);
    Object.keys(tg).forEach(function (npp) {
      var t = tg[npp], b = t.berelasi && !rel;
      t.lines.forEach(function (l) {
        var v = npv(l.saldo, l.umur, r);
        items.push({ key: 'TG|' + npp + '|' + l.tgl + '|' + l.nodok, akun: t.akun === 'HAD' ? 'HAD' : 'TAGBRUT', berelasi: t.berelasi, pelanggan: t.pro.pelanggan || '', bp: t.pel.kode_bp || '', npp: npp,
          ref: l.nodok || l.sumber, doc: l.nodok, tgl: l.tgl, jumlah: l.nilai, saldo: l.saldo, umur: l.umur, npv: b ? l.saldo : v, ckpn: b ? 0 : rnd(l.saldo - v) });
      });
    });
    ctx._ck[key] = items;
    return items;
  }
  var CK_GROUPS = [['PIUTANG_USAHA', 'Piutang Usaha'], ['PIUTANG_RETENSI', 'Piutang Retensi'], ['TAGBRUT', 'Tagihan Bruto (Pendapatan yg Diperhitungkan)'], ['HAD', 'Hasil Akan Diterima (Piutang Prestasi)']];
  var CK_KODE = { PIUTANG_USAHA: 'PIUTANG_USAHA', PIUTANG_RETENSI: 'RETENSI', TAGBRUT: 'TAGBRUT', HAD: 'HAD' };
  function ckpnMovement(ctx, C) {
    var cur = ckpnAt(ctx, C), prevC = prevMonthEnd(C), prev = ckpnAt(ctx, prevC), pm = {}, cm = {}, g = {};
    // tagbrut/HAD baru punya data sejak saldo awal: sebelum itu pergerakan tidak dihitung (dianggap sudah dibukukan Akuntansi)
    var seed = D(ctx.params.TAGBRUT_SALDO_AWAL_PER) || '0000-00-00', seedGap = prevC < seed && C >= seed;
    if (seedGap) {
      prev = prev.filter(function (x) { return x.akun !== 'TAGBRUT' && x.akun !== 'HAD'; })
        .concat(cur.filter(function (x) { return x.akun === 'TAGBRUT' || x.akun === 'HAD'; }));
    }
    prev.forEach(function (x) { pm[x.key] = x; }); cur.forEach(function (x) { cm[x.key] = x; });
    function grp(x) { var k = x.akun + '|' + (x.berelasi ? 'R' : 'K'); return g[k] || (g[k] = { akun: x.akun, berelasi: x.berelasi, saldo: 0, ckpn: 0, prev: 0, naik: 0, turun: 0, n: 0 }); }
    cur.forEach(function (x) { var o = grp(x), pv = pm[x.key] ? pm[x.key].ckpn : 0; o.saldo += x.saldo; o.ckpn += x.ckpn; o.n++; if (x.ckpn > pv) o.naik += x.ckpn - pv; else o.turun += pv - x.ckpn; });
    prev.forEach(function (x) { var o = grp(x); o.prev += x.ckpn; if (!cm[x.key]) o.turun += x.ckpn; });
    var list = Object.keys(g).map(function (k) { return g[k]; }).sort(function (a, b) { var ia = CK_GROUPS.map(function (z) { return z[0]; }).indexOf(a.akun), ib = CK_GROUPS.map(function (z) { return z[0]; }).indexOf(b.akun); return ia - ib || (a.berelasi ? -1 : 1); });
    return { C: C, prevC: prevC, groups: list, cur: cur, prev: prev, seedGap: seedGap };
  }
  function jCkpn(ctx, C) {
    var mv = ckpnMovement(ctx, C), J = [];
    mv.groups.forEach(function (o) {
      var k = CK_KODE[o.akun], lbl = (o.berelasi ? 'PB ' : 'PK ') + k.replace('_', ' ');
      if (o.naik > 0.5) J.push({ kejadian: 'CKPN — kenaikan', tgl: C, ref: lbl, tcode: 'F-02', doctype: 'SA', sap: '',
        lines: [jl(ctx, 'BEBAN_CKPN_' + k, o.berelasi, '40', o.naik, 0, { teks: 'BEBAN CKPN ' + lbl }), jl(ctx, 'CKPN_' + k, o.berelasi, '50', 0, o.naik, { teks: 'BEBAN CKPN ' + lbl })] });
      if (o.turun > 0.5) J.push({ kejadian: 'CKPN — pemulihan', tgl: C, ref: lbl, tcode: 'F-02', doctype: 'SA', sap: '',
        lines: [jl(ctx, 'CKPN_' + k, o.berelasi, '40', o.turun, 0, { teks: 'PEMULIHAN CKPN ' + lbl }), jl(ctx, 'PEMULIHAN_CKPN_' + k, o.berelasi, '50', 0, o.turun, { teks: 'PEMULIHAN CKPN ' + lbl })] });
    });
    return J;
  }

  // ------------------------------------------------------------------ FAKTUR PAJAK (CORETAX, KODE 04)
  function nsfpParts(s) { var d = digits(s); return d.length === 17 ? { kode: d.substring(0, 2), status: d.substring(2, 4), seri: d.substring(4), full: d } : { kode: '', status: '', seri: '', full: d }; }

  // ------------------------------------------------------------------ DATA PAKET & PERSETUJUAN
  function invPackage(ctx, nomor) {
    var k = normInv(nomor), inv = ctx.invs.filter(function (i) { return normInv(i.nomor) === k; })[0];
    if (!inv) return null;
    var rec = ctx.latest(DOC.INV, k);
    return { inv: plainInv(inv), kartu: kartuPlain(ctx, inv.npp, ctx.today), jurnal: [jPlain(jInvoice(ctx, inv))], faktur: fakturDraft(ctx, inv),
      fp: inv._fp, appr: rec ? apprPlain(rec, inv._fp) : null, ref: k, jenis: DOC.INV, label: 'Paket invoice ' + inv.nomor + ' — ' + inv.pelanggan, nilai: inv._a.total };
  }
  function payPackage(ctx, key) {
    var g = ctx.payGroups[key];
    if (!g || !g.length) return null;
    var fp = fpPayGroup(g), rec = ctx.latest(DOC.PAY, key), npps = {};
    g.forEach(function (p) { npps[p._inv.npp] = 1; });
    return { key: key, rows: g.map(function (p) { var i = p._inv; return { idp: p.idp, tgl: p.tgl, ref: S(p.ref), transfer: N(p.transfer), nomor: i.nomor, pelanggan: i.pelanggan, npp: i.npp, proyek: i._pro ? i._pro.singkat : '',
        invke: invLabel(i), nominal: p.nominal, keN: p._keN, keP: p._keP, pph: (p._pph || 0) + (p._sel || 0), adm: p._adm || 0, aturan: p._rule, cicil: p._ke, ket: S(p.ket), sap: p.sap_doc,
        status: i._res.status, sisa: i._res.sisa, kas: i._a.kas, rekening: p._rek ? p._rek.label : S(p.rekening), pb: !!p._pb }; }),
      kartu: Object.keys(npps).map(function (n) { return kartuPlain(ctx, n, ctx.today); }), jurnal: jPays(ctx, g).map(jPlain),
      fp: fp, appr: rec ? apprPlain(rec, fp) : null, ref: key, jenis: DOC.PAY, label: 'Paket penerimaan ' + key + ' — ' + g[0]._inv.pelanggan, nilai: sumBy(g, function (p) { return p.nominal; }) };
  }
  /** kartu piutang yang disetujui: sidik = ringkasan saldo per invoice pada tanggal posisi */
  function kpFp(ctx, npp, asOf) {
    var k = kartuData(ctx, npp, D(asOf) || ctx.today); if (!k) return '';
    return hash(['KP', npp, k.asOf, rnd(k.T.debit), rnd(k.T.kredit), rnd(k.T.saldo), k.groups.map(function (g) { return normInv(g.inv.nomor) + ':' + rnd(g.saldo); }).join(';')].join('|'));
  }
  function kpPackage(ctx, ref) {
    var x = S(ref).split('|'), npp = x[0], asOf = D(x[1]) || ctx.today, pro = ctx.M.pro[npp];
    if (!pro) return null;
    if (asOf > ctx.today) asOf = ctx.today;
    var r = npp + '|' + asOf, fp = kpFp(ctx, npp, asOf), rec = ctx.latest(DOC.KP, r), k = kartuPlain(ctx, npp, asOf);
    return { ref: r, npp: npp, asOf: asOf, kartu: k, fp: fp, appr: rec ? apprPlain(rec, fp) : null, jenis: DOC.KP, label: 'Kartu piutang ' + npp + ' ' + pro.singkat + ' per ' + tglPanjang(asOf), nilai: k.T.saldo };
  }
  function tutupPackage(ctx, nomor) {
    var k = normInv(nomor), inv = ctx.invs.filter(function (i) { return normInv(i.nomor) === k; })[0];
    if (!inv || U(inv.tutup) !== 'CLOSE MANUAL' || !inv._tutupFp) return null;
    var rec = ctx.latest(DOC.TUTUP, k);
    return { ref: k, inv: plainInv(inv), alasan: S(inv.alasan_tutup), sisa: inv._tutupSisa, kartu: kartuPlain(ctx, inv.npp, ctx.today), fp: inv._tutupFp,
      appr: rec ? apprPlain(rec, inv._tutupFp) : null, jenis: DOC.TUTUP, label: 'Penutupan manual ' + inv.nomor + ' — ' + inv.pelanggan, nilai: inv._tutupSisa };
  }
  /** monitoring uang muka & retensi per NPP */
  function umRetRows(ctx) {
    if (ctx._umr) return ctx._umr;
    var TOL = N(ctx.params.TOL_PEMBULATAN) || 1000, ret = {}, out = [];
    retensiRows(ctx).forEach(function (r) { ret[r[1]] = r; });
    Object.keys(ctx.M.pro).sort().forEach(function (npp) {
      var p = ctx.M.pro[npp], L = ctx.invs.filter(function (i) { return i.npp === npp && !i._batal; });
      var umInv = sumBy(L, function (i) { return i._a.jenis === 'UANG MUKA' ? i._a.dpp : 0; }), basis = p.um > 0 ? p.um : umInv, awal = p.um_awal || 0;
      var kembali = sumBy(L, function (i) { return i._a.umk; }), sisa = basis - awal - kembali;
      var prog = sumBy(L, function (i) { return i._a.jenis === 'PROGRES' || i._a.jenis === 'LAINNYA' ? i._a.bruto : 0; }), pctProg = p.nilai_ok ? prog / p.nilai_ok : '';
      var lastUm = L.filter(function (i) { return i._a.umk > 0; }).sort(function (a, b) { return a.tgl < b.tgl ? -1 : 1; }).pop();
      // invoice progres hasil migrasi tidak menyimpan potongan UM → sisa UM belum bisa dihitung sampai saldo awal diisi
      var migr = L.filter(function (i) { return (i._a.jenis === 'PROGRES' || i._a.jenis === 'LAINNYA') && /^MIGRASI/.test(U(i.sumber)) && !N(i.umk) && !N(i.bruto); });
      var lengkap = !(basis > 0 && migr.length && !awal);
      var umSt = '';
      if ((basis || kembali || awal) && !p.selesai) {
        if (!lengkap) umSt = '⚠ DATA LAMA BELUM LENGKAP: ' + migr.length + ' invoice migrasi tanpa potongan UM — isi UM SUDAH DIKEMBALIKAN SEBELUM SISTEM di Master Proyek';
        else if (sisa < -TOL) umSt = '⚠ PENGEMBALIAN UM MELEBIHI UM (' + fmtRp(-sisa) + ')';
        else if (Math.abs(sisa) <= TOL) umSt = '✓ UM SUDAH HABIS DIKEMBALIKAN';
        else if (isNum(pctProg) && pctProg >= 0.999) umSt = '⚠ PROGRES 100% TETAPI UM MASIH SISA';
        else umSt = 'BERJALAN';
        if (p.um > 0 && umInv > 0 && Math.abs(p.um - umInv) > TOL) umSt += ' · UM di Master (' + fmtRp(p.um) + ') ≠ invoice UM (' + fmtRp(umInv) + ')';
      }
      var r = ret[npp], batas = p.ret_pct && p.nilai_ok ? rnd(p.ret_pct * p.nilai_ok) : 0, retSt = '';
      var retInv = L.filter(function (i) { return i._a.jenis === 'RETENSI'; }), dup = {};
      retInv.forEach(function (i) { var k = rnd(i._a.dpp); dup[k] = (dup[k] || []).concat([i.nomor.replace(/^KU\.02\.02\/WB-4C\./, '')]); });
      var dupTxt = Object.keys(dup).filter(function (k) { return dup[k].length > 1; }).map(function (k) { return dup[k].join(' & '); }).join('; ');
      if (r) {
        if (r[7] <= TOL && r[8] > TOL) retSt = 'DATA LAMA: retensi yang ditahan di invoice lama tidak tercatat (sudah ditagih ' + fmtRp(r[8]) + ')';
        else if (r[8] > r[7] + TOL) retSt = '⚠ DITAGIH MELEBIHI YANG DITAHAN';
        else if (batas && r[7] > batas + TOL) retSt = '⚠ DITAHAN MELEBIHI ' + pctTxt(p.ret_pct) + ' NILAI KONTRAK';
        else if (r[7] > 0 && r[9] <= TOL) retSt = '✓ RETENSI SUDAH DITAGIH SEMUA';
        else retSt = r[10] ? '⚠ ' + r[10] : 'BERJALAN';
        if (dupTxt) retSt = '⚠ INVOICE RETENSI NILAI SAMA: ' + dupTxt + ' — cek kemungkinan ganda';
      }
      if (!umSt && !r) return;
      out.push({ npp: npp, pro: p, lengkap: lengkap, pelanggan: p.pelanggan, singkat: p.singkat, nilai_ok: p.nilai_ok, umMaster: p.um, umInv: umInv, basis: basis, awal: awal, kembali: kembali, sisa: sisa,
        pctUm: basis ? (awal + kembali) / basis : '', pctProg: pctProg, umSt: umSt, lastUm: lastUm ? lastUm.nomor : '', ret: r, batas: batas, retSt: retSt, aktif: p.aktif });
    });
    return (ctx._umr = out);
  }
  function rekapData(ctx, ym) {
    var sel = parseMonthSel(ym, ctx.today, false), ref = sel.y + '-' + pad(sel.m, 2);
    var invs = ctx.invs.filter(function (i) { return !i._batal && i.tgl >= sel.start && i.tgl <= sel.end; }).sort(function (a, b) { return a.tgl < b.tgl ? -1 : 1; });
    var pays = ctx.pays.filter(function (p) { return p._inv && p.tgl >= sel.start && p.tgl <= sel.end && (p.nominal > 0 || p._pb); }).sort(function (a, b) { return a.tgl < b.tgl ? -1 : 1; });
    var mv = ckpnMovement(ctx, sel.end);
    var J = invs.map(function (i) { return jPlain(jInvoice(ctx, i)); }).concat(jPays(ctx, pays).map(jPlain)).concat(jCkpn(ctx, sel.end).map(jPlain));
    var rec = ctx.latest(DOC.REK, ref);
    var data = { ref: ref, label: monthLabel(sel.y, sel.m), start: sel.start, end: sel.end,
      invoices: invs.map(function (i) { return { nomor: i.nomor, tgl: i.tgl, pelanggan: i.pelanggan, npp: i.npp, invke: invLabel(i), dpp: i._a.dpp, ppn: i._a.ppn, total: i._a.total, netto: i._a.netto, ret: i._a.ret, umk: i._a.umk, no_fp: i.no_fp, sap: i.sap_doc, appr: apprStatusText(ctx.latest(DOC.INV, normInv(i.nomor)), i._fp) }; }),
      pays: pays.map(function (p) { var i = p._inv; return { idp: p.idp, tgl: p.tgl, nomor: i.nomor, pelanggan: i.pelanggan, npp: i.npp, nominal: p.nominal, keN: p._keN, keP: p._keP, pph: (p._pph || 0) + (p._sel || 0), adm: p._adm || 0, ref: S(p.ref), sap: p.sap_doc, rekening: p._pb ? 'PINDAH BUKU' : (p._rek ? p._rek.label : S(p.rekening)) }; }),
      ckpn: mv.groups.map(function (o) { return { akun: o.akun, berelasi: o.berelasi, saldo: o.saldo, ckpn: o.ckpn, prev: o.prev, naik: o.naik, turun: o.turun }; }),
      jurnal: J, totDebit: sumBy(J, function (j) { return sumBy(j.lines, function (l) { return l.debit; }); }), totKredit: sumBy(J, function (j) { return sumBy(j.lines, function (l) { return l.kredit; }); }) };
    data.fp = hash('REK|' + ref + '|' + data.invoices.length + ':' + rnd(sumBy(data.invoices, function (x) { return x.total; })) + '|' + data.pays.length + ':' + rnd(sumBy(data.pays, function (x) { return x.nominal; })) + '|' + rnd(data.totDebit));
    data.appr = rec ? apprPlain(rec, data.fp) : null; data.jenis = DOC.REK;
    return data;
  }
  function plainInv(i) {
    var a = i._a, r = i._res;
    return { nomor: i.nomor, tgl: i.tgl, pelanggan: i.pelanggan, npp: i.npp, proyek: i._pro ? i._pro.singkat : '', invke: invLabel(i), jenis: a.jenis, template: a.tpl, uraian: uraianOf(i),
      bruto: a.bruto, umk: a.umk, ret: a.ret, retpos: a.retpos, dpp: a.dpp, ppn: a.ppn, total: a.total, netto: a.netto, tarif: a.tarif, pph: a.pph, kas: a.kas,
      ttd: S(i.ttd), rev: N(i.rev), no_fp: i.no_fp, tgl_fp: i.tgl_fp, sap: i.sap_doc, status: r.status, sisa: r.sisa, diterima: r.diterima, jt: r.jt, jt_akt: r.jt_akt,
      kelompok: i._pel ? i._pel.kelompok : '', kode_bp: i._pel ? i._pel.kode_bp : '' };
  }
  function kartuPlain(ctx, npp, asOf) {
    ctx = contextAt(ctx, asOf);
    var k = kartuData(ctx, npp, asOf); if (!k) return null;
    var p = k.pro, pel = k.pel;
    return { npp: npp, asOf: k.asOf, pelanggan: p.pelanggan, proyek: p.singkat, pekerjaan: S(p.pekerjaan) || p.singkat, nama_pe: p.nama_pe || S(p.pekerjaan) || p.singkat,
      produk: nppListTxt(p.produk), adendumNpp: nppListTxt(p.adendumNpp), A: k.groups.map(function (g) { return g.A; }), lokasi: S(p.lokasi), kontrak: S(p.kontrak), tgl_kontrak: D(p.tgl_kontrak),
      nilai_awal: k.kontrak.awal, tambah: k.kontrak.tambah, kurang: k.kontrak.kurang, nilai_ok: k.kontrak.net, template: p.template, tarif: p.tarif_eff, ret_pct: p.ret_pct, ret_pos: p.ret_pos, um: umBasisOf(ctx, npp),
      signer: p.signer ? p.signer.nama : '', kode_bp: pel.kode_bp || '', npwp: pel.npwp16 || '', kelompok: pel.kelompok || '', jenis_pel: pel.jenis || '',
      T: k.T, sisaKontrak: k.sisaKontrak, pctTagih: k.pctTagih,
      groups: k.groups.map(function (g) { return { nomor: g.inv.nomor, status: g.status, saldo: g.saldo, rows: g.rows.map(function (r) { return { kind: r.kind, tgl: r.tgl, uraian: r.uraian, ref: r.ref, sap: r.sap, debit: r.debit, kredit: r.kredit, saldo: r.saldo, baru: !!r.baru, ket: r.ket || '' }; }) }; }) };
  }
  function jPlain(j) { return { kejadian: j.kejadian, tgl: j.tgl, ref: j.ref, tcode: j.tcode, doctype: j.doctype, sap: j.sap || '', header: j.header || '', lines: j.lines.map(function (l) { return { pk: l.pk, gl: l.gl, nama: l.nama, bp: l.cust ? l.bp : '', npp: l.npp || '', debit: l.debit, kredit: l.kredit, teks: l.teks }; }) }; }
  function apprPlain(r, fpNow) {
    return { id: S(r.id), waktu: r.waktu, jenis: r.jenis, ref: r.ref, status: r.status, oleh: S(r.oleh), oleh_nama: r.oleh_nama, putus_oleh: S(r.putus_oleh), putus_nama: r.putus_nama,
      putus_waktu: r.putus_waktu, catatan: S(r.catatan), kode: S(r.kode), sidik: S(r.sidik), berubah: !!(fpNow && S(r.sidik) !== fpNow), teks: apprStatusText(r, fpNow) };
  }
  function approvalsPlain(ctx) {
    return ctx.appr.map(function (r) {
      var fpNow = null;
      if (r.jenis === DOC.INV) { var inv = ctx.invs.filter(function (i) { return normInv(i.nomor) === r.ref; })[0]; fpNow = inv ? inv._fp : 'X'; }
      else if (r.jenis === DOC.PAY) { var g = ctx.payGroups[r.ref]; fpNow = g ? fpPayGroup(g) : 'X'; }
      else if (r.jenis === DOC.KP) { var xp = S(r.ref).split('|'); fpNow = ctx.M.pro[xp[0]] ? kpFp(ctx, xp[0], xp[1]) : 'X'; }
      else if (r.jenis === DOC.TUTUP) { var it = ctx.invs.filter(function (i) { return normInv(i.nomor) === r.ref; })[0]; fpNow = it && it._tutupFp ? it._tutupFp : 'X'; }
      var o = apprPlain(r, fpNow); o.label = S(r.label); o.nilai = N(r.nilai); return o;
    });
  }

  // ------------------------------------------------------------------ CEK TAMBAHAN
  function extraChecks(ctx) {
    var P = ctx.params, I = ctx.issues, mulai = D(P.FAKTUR_CEK_MULAI) || '2026-09-01', hari = N(P.FAKTUR_PENGINGAT_HARI) || 10;
    ctx.invs.forEach(function (i) {
      if (i._batal || !i.tgl) return;
      if (i.tgl >= mulai && i._a.ppn > 0 && !i.no_fp && diffDays(ctx.today, i.tgl) > hari) I.push(['FAKTUR', REG.sheet, i.nomor, 'Nomor faktur pajak belum dicatat (' + diffDays(ctx.today, i.tgl) + ' hari sejak invoice)', 'Buat faktur di Coretax (draft ada di aplikasi), lalu isi NOMOR & TANGGAL FAKTUR PAJAK']);
      if (i.no_fp && i.no_fp.length !== 17) I.push(['FAKTUR', REG.sheet, i.nomor, 'Nomor faktur pajak harus 17 digit (tertulis ' + i.no_fp.length + ' digit)', 'Periksa NOMOR FAKTUR PAJAK']);
      var st = apprStatusText(ctx.latest(DOC.INV, normInv(i.nomor)), i._fp);
      if (/DATA BERUBAH/.test(st)) I.push(['PERSETUJUAN', REG.sheet, i.nomor, 'Invoice berubah setelah diajukan/disetujui', 'Ajukan ulang paket invoice di aplikasi (menu Persetujuan)']);
    });
    ctx.appr.forEach(function (r) {
      if (r.status === 'MENUNGGU' && D(r.waktu) && diffDays(ctx.today, D(r.waktu)) > 3) I.push(['PERSETUJUAN', APPR.sheet, r.id + ' · ' + r.ref, 'Menunggu persetujuan lebih dari 3 hari (' + r.jenis + ')', 'Ingatkan Section Head untuk memeriksa di aplikasi']);
    });
    ['PIUTANG_USAHA', 'PIUTANG_PPN', 'PIUTANG_RETENSI', 'TAGBRUT', 'HAD', 'PPN_KELUARAN', 'UANG_MUKA_PROYEK', 'PPH_FINAL', 'PPH_23', 'BANK', 'BEBAN_ADM_BANK'].forEach(function (k) {
      if (!P.akun[k] || !(P.akun[k].rel || P.akun[k].tiga)) I.push(['PARAMETER', 'Parameter', 'Pemetaan Akun ' + k, 'Akun GL belum diisi', 'Isi di tabel PEMETAAN AKUN SAP (sheet Parameter)']);
    });
    var TOL = N(P.TOL_PEMBULATAN) || 1000, retS = {}; retensiRows(ctx).forEach(function (r) { retS[r[1]] = r[9]; });
    Object.keys(ctx.M.pro).forEach(function (n) {
      var p = ctx.M.pro[n], L = ctx.invs.filter(function (i) { return i.npp === n && !i._batal; });
      var b = sumBy(L.filter(function (i) { return i._a.jenis === 'PROGRES' || i._a.jenis === 'LAINNYA'; }), function (i) { return i._a.bruto; });
      if (p.nilai_ok && b > p.nilai_ok * 1.0005) I.push(['KONTRAK', MPRO.sheet, n, 'Total tagihan progres (' + fmtRp(b) + ') melebihi nilai kontrak setelah adendum (' + fmtRp(p.nilai_ok) + ')', 'Periksa adendum di Riwayat Adendum atau nilai tagihan']);
      var open = L.filter(function (i) { return !i._res.closed; });
      if (p.selesai) {
        if (open.length) I.push(['MASTER', MPRO.sheet, n, 'Proyek ditandai SELESAI tetapi masih ada ' + open.length + ' invoice terbuka (Rp ' + fmtRp(sumBy(open, function (i) { return i._res.sisa; })) + ')', 'Status SELESAI hanya untuk proyek yang piutangnya sudah 0 — kembalikan AKTIF = YA']);
        return;
      }
      if (p.aktif === 'TIDAK') return;
      // kelengkapan master hanya dicek bila proyek masih akan ditagih (sisa kontrak / retensi belum ditagih / nilai kontrak belum diisi)
      var perlu = !p.nilai_ok || (p.nilai_ok - b) > TOL || (retS[n] || 0) > TOL;
      if (!perlu) return;
      if (!S(p.deskripsi)) I.push(['MASTER', MPRO.sheet, n, 'Deskripsi invoice kosong (proyek masih akan ditagih)', 'Isi DESKRIPSI DI INVOICE. Bila tidak ada tagihan baru: AKTIF = TIDAK (atau SELESAI bila sudah lunas semua)']);
      if (p.template === 'FAJARPUTERA' && (!S(p.kontrak) || !S(p.lokasi))) I.push(['MASTER', MPRO.sheet, n, 'Data BAP belum lengkap (No. SPK / Lokasi)', 'Lengkapi NO. KONTRAK/SPK, TGL, dan LOKASI']);
      if (p._noSigner) I.push(['MASTER', MPRO.sheet, n, 'Nilai kontrak kosong — penandatangan tidak bisa ditentukan otomatis', 'Isi NILAI KONTRAK atau PENANDATANGAN MANUAL']);
    });
    umRetRows(ctx).forEach(function (u) { if (/^⚠ PENGEMBALIAN UM MELEBIHI/.test(u.umSt)) I.push(['UANG MUKA', 'Monitoring UM & Retensi', u.npp, u.umSt, 'Periksa potongan UM di invoice terakhir (' + (u.lastUm || '-') + ') atau isi saldo awal UM di Master Proyek']);
      if (/NILAI SAMA/.test(u.retSt)) I.push(['RETENSI', 'Monitoring UM & Retensi', u.npp, u.retSt, 'Batalkan invoice retensi yang ganda (STATUS INVOICE = BATAL)']); });
    var rekNo = (P.rek || []).filter(function (r) { var a = P.akun[r.akun]; return !a || !(a.rel || a.tiga); }).map(function (r) { return r.bank; });
    if (rekNo.length) I.push(['PARAMETER', 'Parameter', 'Rekening penerimaan', 'Akun GL rekening ' + rekNo.join(', ') + ' belum diisi (usulan jurnal penerimaan memakai akun kosong)', 'Konfirmasi ke Akuntansi (Mbak Elmarta), isi di PEMETAAN AKUN SAP kode ' + (P.rek || []).map(function (r) { return r.akun; }).join(' / ')]);
    if (!userByRole(P, 'PENYETUJU')) I.push(['PARAMETER', 'Parameter', 'Pengguna & Peran', 'Belum ada pengguna dengan peran PENYETUJU', 'Isi tabel PENGGUNA & PERAN (email Section Head)']);
    var pelUsed = {};
    ctx.invs.forEach(function (i) { if (!i._batal && i.tgl >= mulai && i._pel) pelUsed[i._pel.nama] = i._pel; });
    Object.keys(pelUsed).forEach(function (k) {
      var p = pelUsed[k];
      if (!p.npwp16 || p.npwp16.length !== 16) I.push(['MASTER', MPEL.sheet, k, 'NPWP belum 16 digit (dipakai di faktur pajak)', 'Lengkapi NPWP']);
      if (!p.kode_bp) I.push(['MASTER', MPEL.sheet, k, 'Kode Business Partner SAP kosong (dipakai di usulan jurnal)', 'Isi dari lampiran LK / tanyakan Akuntansi']);
    });
    // v2.4: kelengkapan nilai progres / pot. UM / retensi (bagian A kartu piutang)
    var lamaBy = {};
    ctx.invs.forEach(function (i) {
      if (i._batal) return;
      var a = i._a, pr = i._pro || {}, prog = a.jenis === 'PROGRES' || a.jenis === 'LAINNYA', TOLB = 5;
      if (prog && !(N(i.bruto) > 0) && (pr.um > 0 || pr.um_awal > 0 || pr.ret_pct > 0)) (lamaBy[i.npp] = lamaBy[i.npp] || []).push(i.nomor.replace(/^KU\.02\.02\/WB-4C\./, ''));
      if (N(i.bruto) > 0 && prog) {
        var exp = a.retpos === 'AWAL' ? a.bruto - a.umk - a.ret : a.bruto - a.umk;
        if (Math.abs(exp - a.dpp) > TOLB) I.push(['DATA INVOICE', REG.sheet, i.nomor, 'Nilai progres − pot. UM' + (a.retpos === 'AWAL' ? ' − retensi' : '') + ' = ' + fmtRp(exp) + ' tidak sama dengan DPP ' + fmtRp(a.dpp), 'Periksa NILAI PROGRES / BRUTO, PENGEMBALIAN UM, RETENSI, POSISI RETENSI']);
      }
      if (a.ret > 0 && a.retpos !== 'AWAL' && a.retpos !== 'AKHIR' && a.jenis !== 'RETENSI') I.push(['DATA INVOICE', REG.sheet, i.nomor, 'Retensi diisi tetapi POSISI RETENSI kosong', 'Isi AWAL (dipotong sebelum DPP) atau AKHIR (ditahan dari pembayaran). AKHIR mengubah Netto tercatat.']);
    });
    Object.keys(lamaBy).forEach(function (n) { var L = lamaBy[n]; I.push(['DATA LAMA', REG.sheet, n, L.length + ' invoice lama belum menyimpan nilai progres / pot. UM / retensi: ' + L.slice(0, 12).join(', ') + (L.length > 12 ? ', …' : ''), 'Isi manual kolom NILAI PROGRES / BRUTO, PENGEMBALIAN UM, RETENSI (& POSISI RETENSI) di Monitoring Invoice sesuai invoice aslinya']); });
    if (P.akun.BEBAN_ADM_BANK && /KONFIRMASI/i.test(P.akun.BEBAN_ADM_BANK.ket)) I.push(['PARAMETER', 'Parameter', 'BEBAN_ADM_BANK', 'Akun beban adm bank masih perlu dikonfirmasi Akuntansi', 'Konfirmasi ke Akuntansi, lalu hapus kata "konfirmasi" di keterangan']);
  }

  // ------------------------------------------------------------------ DASHBOARD
  function repDashboard(ctx) {
    var act = ctx.invs.filter(function (i) { return !i._batal; }), open = openInv(ctx);
    var od = open.filter(function (i) { return /OVERDUE/.test(i._res.status); });
    var ba = open.filter(function (i) { return /BA\/MEMO/.test(i._res.status); }), ppnB = open.filter(function (i) { return /PPN BELUM DIBAYAR/.test(i._res.status); });
    var retSisa = retensiRows(ctx).reduce(function (t, r) { return t + r[9]; }, 0);
    var y = +ctx.today.substring(0, 4), pe = prevMonthEnd(ctx.today);
    var tg = tagAt(ctx, ctx.today), tgS = sumBy(Object.keys(tg), function (k) { return tg[k].sisa; });
    var ck = ckpnAt(ctx, pe), pend = ctx.appr.filter(function (r) { return r.status === 'MENUNGGU'; });
    var mulai = D(ctx.params.FAKTUR_CEK_MULAI) || '2026-09-01';
    var noFp = act.filter(function (i) { return i.tgl >= mulai && i._a.ppn > 0 && !i.no_fp; });
    var kpi = [
      ['RINGKASAN PER ' + tglPanjang(ctx.today).toUpperCase(), '', ''],
      ['Indikator', 'Jumlah', 'Nilai (Rp)'],
      ['Invoice aktif (semua tahun)', act.length, sumBy(act, function (i) { return i._a.total; })],
      ['Invoice terbit tahun ' + y, act.filter(function (i) { return i.tgl && +i.tgl.substring(0, 4) === y; }).length, sumBy(act.filter(function (i) { return i.tgl && +i.tgl.substring(0, 4) === y; }), function (i) { return i._a.total; })],
      ['Sudah diterima (semua pembayaran, kas)', ctx.pays.filter(function (p) { return p._inv && !p._pb; }).length, sumBy(ctx.pays, function (p) { return p._inv && !p._pb ? p.nominal : 0; })],
      ['Piutang terbuka (sisa kas, incl. PPN)', open.length, sumBy(open, function (i) { return i._res.sisa; })],
      ['  · rencana cair (DPP − PPh, tanpa PPN)', open.length, sumBy(open, function (i) { return i._res.rencana_net; })],
      ['  · sudah lewat jatuh tempo', od.length, sumBy(od, function (i) { return i._res.sisa; })],
      ['  · kekurangan ditagih via BA/memo', ba.length, sumBy(ba, function (i) { return i._res.sisa; })],
      ['  · Netto lunas, PPN belum dibayar', ppnB.length, sumBy(ppnB, function (i) { return i._res.sisa; })],
      ['Piutang retensi belum ditagih', '', retSisa],
      ['Tagbrut / HAD belum ditagih (hari ini)', Object.keys(tg).filter(function (k) { return tg[k].sisa > 0; }).length, tgS],
      ['CKPN per ' + tglPanjang(pe), ck.length, sumBy(ck, function (x) { return x.ckpn; })],
      ['Faktur pajak belum dicatat (sejak ' + tglPanjang(mulai) + ')', noFp.length, sumBy(noFp, function (i) { return i._a.ppn; })],
      ['Menunggu persetujuan Section Head', pend.length, sumBy(pend, function (r) { return N(r.nilai); })],
      ['Data yang perlu dicek (sheet Cek Data)', ctx.issues.length, '']
    ];
    var st = {};
    act.forEach(function (i) { var k = i._res.status.replace(/ [\d\.\-]+$/, '').replace(/ Rp.*$/, ''); k = k.replace(/ (\d{1,3}(\.\d{3})*)$/, ''); var o = st[k] || (st[k] = [0, 0, 0]); o[0]++; o[1] += i._a.total; o[2] += i._res.sisa; });
    var order = ['WAITING SUBMIT', 'BELUM KELUAR TANDA TERIMA', 'BELUM JATUH TEMPO', 'JATUH TEMPO HARI INI', 'OVERDUE', 'CICILAN BERJALAN', 'CICILAN BERJALAN (OVERDUE)', 'NETTO LUNAS – PPN BELUM DIBAYAR', 'NETTO LUNAS – PPN BELUM DIBAYAR (OVERDUE)',
      'KURANG BAYAR – DITAGIH VIA BA/MEMO', 'KURANG BAYAR', 'CLOSE MANUAL – BELUM DIAJUKAN', 'CLOSE MANUAL – MENUNGGU PERSETUJUAN', 'LEBIH BAYAR', 'CLOSE – SELISIH PPh', 'LUNAS', 'CLOSE MANUAL'];
    var keys = Object.keys(st).sort(function (a, b) { var ia = order.indexOf(a), ib = order.indexOf(b); return (ia < 0 ? 99 : ia) - (ib < 0 ? 99 : ib); });
    var srows = [['STATUS PEMBAYARAN', 'Jumlah Invoice', 'Nilai Invoice (Rp)', 'Sisa Piutang (Rp)']];
    keys.forEach(function (k) { srows.push([k, st[k][0], st[k][1], st[k][2]]); });
    srows.push(['TOTAL', sumBy(keys, function (k) { return st[k][0]; }), sumBy(keys, function (k) { return st[k][1]; }), sumBy(keys, function (k) { return st[k][2]; })]);
    var bp = {};
    open.forEach(function (i) { var o = bp[i.pelanggan] || (bp[i.pelanggan] = [0, 0, 0, i._pel ? i._pel.jenis : '']); o[0]++; o[1] += i._res.sisa; if (/OVERDUE/.test(i._res.status)) o[2] += i._res.sisa; });
    var prow = [['PIUTANG TERBUKA PER PELANGGAN', 'Jenis', 'Invoice Terbuka', 'Sisa Piutang (Rp)', 'Lewat Jatuh Tempo (Rp)']];
    Object.keys(bp).sort(function (a, b) { return bp[b][1] - bp[a][1]; }).forEach(function (k) { prow.push([k, bp[k][3], bp[k][0], bp[k][1], bp[k][2]]); });
    prow.push(['TOTAL', '', sumBy(Object.keys(bp), function (k) { return bp[k][0]; }), sumBy(Object.keys(bp), function (k) { return bp[k][1]; }), sumBy(Object.keys(bp), function (k) { return bp[k][2]; })]);
    // ---- v2.4: kartu angka + data grafik (posisi tetap agar grafik bawaan Sheets selalu membaca sel yang sama)
    var y0 = +ctx.today.substring(0, 4), m0 = +ctx.today.substring(5, 7), ly = m0 === 1 ? y0 - 1 : y0, lm = m0 === 1 ? 12 : m0 - 1;
    var lmS = monthStart(ly, lm), lmE = monthEnd(ly, lm), selNow = parseMonthSel('', ctx.today, false);
    var sisaAll = sumBy(open, function (i) { return i._res.sisa; }), odSisa = sumBy(od, function (i) { return i._res.sisa; });
    var cair = open.filter(function (i) { var d = i._res.jt; return !d || d <= selNow.end; });
    var trmLm = sumBy(ctx.pays, function (p) { return p._inv && !p._pb && p.tgl >= lmS && p.tgl <= lmE ? p.nominal : 0; });
    var tagLm = sumBy(act.filter(function (i) { return i.tgl >= lmS && i.tgl <= lmE; }), function (i) { return i._a.total; });
    var memoBelum = ctx.memo ? ctx.memo.belum.length : 0, pctOd = sisaAll ? odSisa / sisaAll : 0, pctTr = tagLm ? trmLm / tagLm : 0;
    var M3 = function (x) { return 'Rp ' + (x / 1e9).toFixed(2).replace('.', ',') + ' M'; };
    var tiles = [
      ['PIUTANG TERBUKA', '', '', 'LEWAT JATUH TEMPO', '', '', 'RENCANA CAIR ' + MONTHS[m0 - 1].toUpperCase() + ' ' + y0, '', '', 'DITERIMA ' + MONTHS[lm - 1].toUpperCase() + ' ' + ly, '', '', 'MENUNGGU PERSETUJUAN', '', ''],
      [M3(sisaAll), '', '', M3(odSisa), '', '', M3(sumBy(cair, function (i) { return i._res.rencana_net; })), '', '', M3(trmLm), '', '', String(pend.length), '', ''],
      [open.length + ' invoice · incl. PPN', '', '', od.length + ' invoice', '', '', cair.length + ' invoice · DPP − PPh (incl. tunggakan)', '', '', 'ditagih bulan itu ' + M3(tagLm), '', '', 'memorial belum diverifikasi: ' + memoBelum, '', ''],
      ['', '', '', (pctOd > 0.3 ? '▲ ' : '● ') + pctTxt(pctOd) + ' dari piutang', '', '', '', '', '', (pctTr >= 0.8 ? '✓ ' : '● ') + pctTxt(pctTr) + ' dari yang ditagih', '', '', pend.length ? '● perlu diputuskan' : '✓ tidak ada antrean', '', '']];
    var merges = []; for (var tr = 0; tr < 4; tr++) for (var tc = 0; tc < 5; tc++) merges.push([tr, tc * 3, 1, 3]);
    var tileB = block(6, 2, tiles, { styles: [sty(0, 'tileL'), sty(1, 'tileV'), sty(2, 'tileS'), sty(3, 'tileF')], merge: merges, noWidth: true });
    tileB.flags = [pctOd > 0.3 ? 'R' : 'A', pctTr >= 0.8 ? 'G' : 'A', pend.length ? 'A' : 'G'];
    // v2.5: baris tile kedua — kinerja piutang (detail di sheet Kinerja & Peringatan Dini)
    var KT = kinerjaTiles(ctx), tileB2 = block(11, 2, KT.rows, { styles: [sty(0, 'tileL'), sty(1, 'tileV'), sty(2, 'tileS'), sty(3, 'tileF')], merge: merges, noWidth: true });
    tileB2.flags = KT.flags;
    // data grafik di kolom R (18) ke kanan
    var D1 = [['DATA GRAFIK (otomatis) — Rp miliar'], ['BULAN', 'TAGIHAN TERBIT (INCL. PPN)', 'PENERIMAAN KAS']];
    for (var k1 = 11; k1 >= 0; k1--) {
      var yy = y0 + Math.floor((m0 - 1 - k1) / 12), mm = ((m0 - 1 - k1) % 12 + 12) % 12 + 1, s0 = monthStart(yy, mm), s1 = monthEnd(yy, mm);
      D1.push([MONTHS[mm - 1].substring(0, 3) + ' ' + String(yy).substring(2), rnd(sumBy(act.filter(function (i) { return i.tgl >= s0 && i.tgl <= s1; }), function (i) { return i._a.total; }) / 1e7) / 100,
        rnd(sumBy(ctx.pays, function (p) { return p._inv && !p._pb && p.tgl >= s0 && p.tgl <= s1 ? p.nominal : 0; }) / 1e7) / 100]);
    }
    var ag = [0, 0, 0, 0]; open.forEach(function (i) { var d = N(i._res.umur); ag[d <= 30 ? 0 : d <= 60 ? 1 : d <= 90 ? 2 : 3] += i._res.sisa; });
    var D2 = [['UMUR PIUTANG (SEJAK TGL INVOICE)', 'SISA PIUTANG'], ['0–30 hari', ag[0] / 1e9], ['31–60 hari', ag[1] / 1e9], ['61–90 hari', ag[2] / 1e9], ['> 90 hari', ag[3] / 1e9]];
    var D3 = [['PROYEKSI RENCANA CAIR (JATUH TEMPO)', 'DPP − PPh']];
    for (var k3 = 0; k3 < 6; k3++) {
      var y3 = y0 + Math.floor((m0 - 1 + k3) / 12), m3 = (m0 - 1 + k3) % 12 + 1, a3 = monthStart(y3, m3), b3 = monthEnd(y3, m3);
      D3.push([MONTHS[m3 - 1].substring(0, 3) + ' ' + String(y3).substring(2) + (k3 === 0 ? ' (+tunggakan)' : ''), sumBy(open.filter(function (i) { var d = i._res.jt; return k3 === 0 ? (d && d <= b3) : (d && d >= a3 && d <= b3); }), function (i) { return i._res.rencana_net; }) / 1e9]);
    }
    D3.push(['Belum ada JT', sumBy(open.filter(function (i) { return !i._res.jt; }), function (i) { return i._res.rencana_net; }) / 1e9]);
    var pk = {}; open.forEach(function (i) { var k = i.pelanggan.replace(/^PT\.?\s+/, '').replace(/\s*\(PERSERO\)/, '').replace(/,?\s*TBK\.?$/, ''), o = pk[k] || (pk[k] = [0, 0]); if (/OVERDUE/.test(i._res.status)) o[0] += i._res.sisa; else o[1] += i._res.sisa; });
    var top8 = Object.keys(pk).sort(function (a, b) { return (pk[b][0] + pk[b][1]) - (pk[a][0] + pk[a][1]); }).slice(0, 8);
    var D4 = [['PELANGGAN (8 TERBESAR)', 'LEWAT JATUH TEMPO', 'BELUM LEWAT JATUH TEMPO']];
    for (var k4 = 0; k4 < 8; k4++) { var nm = top8[k4]; D4.push(nm ? [nm.length > 30 ? nm.substring(0, 29) + '…' : nm, pk[nm][0] / 1e9, pk[nm][1] / 1e9] : ['', '', '']); }
    var cat = [['Lewat jatuh tempo', /OVERDUE/], ['Belum jatuh tempo', /BELUM JATUH|HARI INI/], ['Waiting submit', /WAITING/], ['Belum keluar tanda terima', /TANDA TERIMA/], ['Cicilan berjalan', /CICILAN/],
      ['PPN belum dibayar', /PPN BELUM/], ['Kurang bayar via BA/memo', /BA\/MEMO/], ['Lainnya', /./]], cv = cat.map(function () { return [0, 0]; });
    open.forEach(function (i) { for (var c = 0; c < cat.length; c++) if (cat[c][1].test(i._res.status)) { cv[c][0]++; cv[c][1] += i._res.sisa; break; } });
    var D5 = [['STATUS PIUTANG TERBUKA', 'SISA PIUTANG', 'JUMLAH INVOICE']].concat(cat.map(function (c, ix) { return [c[0] + ' (' + cv[ix][0] + ')', cv[ix][1] / 1e9, cv[ix][0]]; }));
    var g2 = { 1: 'num2', 2: 'num2' };
    var dB = [block(6, 18, D1, { styles: [sty(0, 'title'), sty(1, 'hdr')], fmt: g2, noWidth: true }), block(21, 18, D2, { styles: [sty(0, 'hdr')], fmt: g2, noWidth: true }),
      block(27, 18, D3, { styles: [sty(0, 'hdr')], fmt: g2, noWidth: true }), block(36, 18, D4, { styles: [sty(0, 'hdr')], fmt: g2, noWidth: true }),
      block(46, 18, D5, { styles: [sty(0, 'hdr')], fmt: { 1: 'num2', 2: 'int' }, noWidth: true })];
    var charts = [
      { key: 'dash1', type: 'COLUMN', title: 'Tagihan terbit vs penerimaan kas per bulan (Rp miliar)', range: [7, 18, 13, 3], row: 16, col: 2, w: 1180, h: 300, colors: ['#2a78d6', '#eb6834'], legend: 'top' },
      { key: 'dash2', type: 'COLUMN', title: 'Umur piutang terbuka (Rp miliar)', range: [21, 18, 5, 2], row: 32, col: 2, w: 585, h: 290, colors: ['#256abf'], legend: 'none', labels: true },
      { key: 'dash3', type: 'COLUMN', title: 'Proyeksi rencana cair, DPP − PPh (Rp miliar)', range: [27, 18, 8, 2], row: 32, col: 9, w: 585, h: 290, colors: ['#2a78d6'], legend: 'none', labels: true },
      { key: 'dash4', type: 'BAR', title: '8 pelanggan dengan piutang terbesar (Rp miliar)', range: [36, 18, 9, 3], row: 48, col: 2, w: 585, h: 330, colors: ['#d03b3b', '#86b6ef'], legend: 'top', stacked: true },
      { key: 'dash5', type: 'BAR', title: 'Piutang terbuka per status (Rp miliar)', range: [46, 18, 9, 2], row: 48, col: 9, w: 585, h: 330, colors: ['#1F3864'], legend: 'none', labels: true }];
    // tabel angka (seperti sebelumnya) di bawah grafik
    var r0 = 66, b1 = block(r0, 2, kpi, { styles: [sty(0, 'title'), sty(1, 'hdr'), sty(5, 'bold'), sty(14, 'warn')], fmt: { 1: 'int', 2: 'money' }, span: 4 });
    var r1 = r0 + kpi.length + 2, b2 = block(r1, 2, srows, { styles: [sty(0, 'hdr'), sty(srows.length - 1, 'total')], fmt: { 1: 'int', 2: 'money', 3: 'money' }, span: 4 });
    var r2 = r1 + srows.length + 2, b3 = block(r2, 2, prow, { styles: [sty(0, 'hdr'), sty(prow.length - 1, 'total')], fmt: { 2: 'int', 3: 'money', 4: 'money' }, span: 4 });
    var W = { 1: 2, 17: 2, 18: 30, 19: 16, 20: 16 }; for (var wc = 2; wc <= 16; wc++) W[wc] = 12;
    return { from: 6, blocks: [tileB, tileB2].concat(dB).concat([b1, b2, b3]), charts: charts, widths: W };
  }

  // ------------------------------------------------------------------ PARETO & VALIDASI RENCANA CAIR
  function agingCat(days) { return days <= 30 ? '0 – 30 hari' : days <= 60 ? '31 – 60 hari' : days <= 90 ? '61 – 90 hari' : '> 90 hari'; }
  function repPareto(ctx) {
    var sel = parseMonthSel(ctx.sel['Pareto'], ctx.today, false), pref = sel.y + '-' + pad(sel.m, 2), vdM = (ctx.valDec || {})[pref], pendM = ctx.latest(DOC.VAL, pref);
    var open = openInv(ctx);
    var list = open.filter(function (i) { var d = i._res.jt; return !d || d <= sel.end; });
    list.sort(function (a, b) { var da = a._res.jt || '9999', db = b._res.jt || '9999'; return da < db ? -1 : da > db ? 1 : (b._res.rencana_net - a._res.rencana_net); });
    var tung = list.filter(function (i) { return i._res.jt && i._res.jt < sel.start; });
    var bln = list.filter(function (i) { return i._res.jt && i._res.jt >= sel.start && i._res.jt <= sel.end; });
    var est = list.filter(function (i) { return !i._res.jt; });
    // keputusan Section Head untuk bulan terpilih (kolom terkunci; diisi otomatis saat Lembar Validasi disetujui)
    var dk = function (i) { var x = vdM && vdM.map[normInv(i.nomor)]; return x ? x.d : ''; };
    var vk = function (i) { var x = vdM && vdM.map[normInv(i.nomor)]; if (x) return DEC[x.d]; return pendM && pendM.status === 'MENUNGGU' ? '⏳ menunggu keputusan' : ''; };
    var va = function (i) { var x = vdM && vdM.map[normInv(i.nomor)]; return x && x.d === 'H' ? x.a : ''; };
    var appr = list.filter(function (i) { return dk(i) === 'C'; }), wait = list.filter(function (i) { return dk(i) === 'S'; });
    var usul = list.filter(function (i) { return dk(i) === 'H' || dk(i) === 'N'; });
    var belum = list.filter(function (i) { return !dk(i); });
    var f = function (i) { return i._res.rencana_net; }, inf = function (L) { return L.length + ' inv · Rp ' + fmtRp(sumBy(L, f)); };
    var sum = [['PERIODE', 'JUMLAH INVOICE HARUS CAIR', 'TOTAL RENCANA CAIR (DPP − PPh, TANPA PPN)', 'TUNGGAKAN (JT SEBELUM BULAN INI)', 'JATUH TEMPO DI BULAN INI', 'ESTIMASI (JT BELUM ADA)',
      '✓ CAIR (VALIDASI SECTION HEAD F&HR)', '◇ CADANGAN', '⏸ HOLD · → BULAN DEPAN', vdM ? '☐ TIDAK ADA DI LEMBAR VALIDASI' : (pendM && pendM.status === 'MENUNGGU' ? '⏳ MENUNGGU VALIDASI SECTION HEAD F&HR' : '☐ BELUM DIVALIDASI SECTION HEAD F&HR'), 'PERKIRAAN KAS MASUK — SELURUH PIUTANG TERBUKA (INCL. PPN, SETELAH PPh)'],
      [sel.label, list.length, sumBy(list, f), sumBy(tung, f), sumBy(bln, f), sumBy(est, f), inf(appr), inf(wait), inf(usul), inf(belum), sumBy(open, function (i) { return i._res.sisa; })]];
    var blocks = [block(8, 1, sum, { styles: [sty(0, 'hdr'), sty(1, 'kpi')], fmt: { 1: 'int', 2: 'money', 3: 'money', 4: 'money', 5: 'money', 10: 'money' }, span: 2 })];
    var ag = { '0 – 30 hari': [0, 0, 0], '31 – 60 hari': [0, 0, 0], '61 – 90 hari': [0, 0, 0], '> 90 hari': [0, 0, 0] };
    list.forEach(function (i) { var o = ag[agingCat(N(i._res.umur))]; o[0]++; o[1] += i._res.rencana_net; o[2] += i._res.sisa; });
    var agr = [['UMUR PIUTANG', 'Jumlah', 'Rencana Cair (DPP − PPh, tanpa PPN)', 'Perkiraan Kas Masuk (incl. PPN, setelah PPh)']];
    Object.keys(ag).forEach(function (k) { agr.push([k, ag[k][0], ag[k][1], ag[k][2]]); });
    agr.push(['TOTAL', list.length, sumBy(list, f), sumBy(list, function (i) { return i._res.sisa; })]);
    var pj = [['PROYEKSI 6 BULAN (JATUH TEMPO)', 'Jumlah', 'Rencana Cair (DPP − PPh, tanpa PPN)']];
    for (var k = 0; k < 6; k++) {
      var yy = sel.y + Math.floor((sel.m - 1 + k) / 12), mm = (sel.m - 1 + k) % 12 + 1, s0 = monthStart(yy, mm), s1 = monthEnd(yy, mm);
      var inm = open.filter(function (i) { var d = i._res.jt; return k === 0 ? (d && d <= s1) : (d && d >= s0 && d <= s1); });
      pj.push([monthLabel(yy, mm) + (k === 0 ? ' (+ tunggakan)' : ''), inm.length, sumBy(inm, f)]);
    }
    pj.push(['TOTAL 6 BULAN', sumBy(pj.slice(1), function (r) { return r[1]; }), sumBy(pj.slice(1), function (r) { return r[2]; })]);
    blocks.push(block(12, 1, agr, { styles: [sty(0, 'hdr'), sty(agr.length - 1, 'total')], fmt: { 1: 'int', 2: 'money', 3: 'money' }, span: 2 }));
    blocks.push(block(12, 7, pj, { styles: [sty(0, 'hdr'), sty(pj.length - 1, 'total')], fmt: { 1: 'int', 2: 'money' } }));
    var top = 12 + Math.max(agr.length, pj.length) + 2;
    var hdr = ['NO', 'NAMA PELANGGAN', 'NOMOR INVOICE', 'TGL INVOICE', 'INVOICE KE-', 'NPP', 'NAMA PROYEK', 'DPP / NETTO', 'PPh', 'RENCANA CAIR (DPP − PPh, TANPA PPN)', 'PERKIRAAN KAS MASUK (INCL. PPN, SETELAH PPh)',
      'TGL JATUH TEMPO', 'UMUR (HARI)', 'LEWAT JT (HARI)', 'KATEGORI UMUR', 'STATUS', 'USULAN AR 🔒', 'RENCANA TGL CAIR 🔒', 'VALIDASI SECTION HEAD F&HR 🔒', 'ALASAN HOLD 🔒', 'CATATAN AR 🔒'];
    var rows = [['DAFTAR TAGIHAN YANG HARUS CAIR s.d. ' + sel.label.toUpperCase() + '  —  usulan AR, rencana tgl cair & catatan diisi di aplikasi ▸ Persetujuan ▸ Lembar Validasi Rencana Cair. Semua kolom 🔒 terisi otomatis (ketikan di sheet dikembalikan).'], hdr];
    list.forEach(function (i, ix) {
      var jt = i._res.jt, lewat = jt && ctx.today > jt ? diffDays(ctx.today, jt) : 0;
      rows.push([ix + 1, i.pelanggan, i.nomor, i.tgl, S(i.invke), i.npp, i._pro ? i._pro.singkat : '', i._a.netto, i._a.pph, i._res.rencana_net, i._res.sisa,
        jt, i._res.umur, lewat, agingCat(N(i._res.umur)), i._res.status, S(i.validasi), i.rencana, vk(i), va(i), S(i.catatan)]);
    });
    rows.push(['', 'TOTAL', '', '', '', '', '', sumBy(list, function (i) { return i._a.netto; }), sumBy(list, function (i) { return i._a.pph; }), sumBy(list, f), sumBy(list, function (i) { return i._res.sisa; })]);
    blocks.push(block(top, 1, rows, { styles: [sty(0, 'title'), sty(1, 'hdr'), sty(rows.length - 1, 'total')],
      fmt: { 3: 'date', 7: 'money', 8: 'money', 9: 'money', 10: 'money', 11: 'date', 12: 'int', 13: 'int', 17: 'date' },
      edit: { keyCol: 2, firstDataOff: 2, lastDataOff: rows.length - 2, cols: {}, lock: [16, 17, 18, 19, 20] } }));
    // bukti validasi bulan terpilih
    var ref = sel.y + '-' + pad(sel.m, 2), rec = ctx.latest(DOC.VAL, ref), r2 = top + rows.length + 2;
    var vb = [['BUKTI VALIDASI RENCANA CAIR — ' + monthLabel(sel.y, sel.m).toUpperCase()]];
    if (!rec || !rec._data) vb.push(['Belum ada Lembar Validasi yang diajukan untuk bulan ini. Ajukan dari aplikasi ▸ Persetujuan ▸ Validasi rencana cair.']);
    else {
      vb.push(['ID ' + rec.id + ' · diajukan ' + rec.waktu + ' oleh ' + rec.oleh_nama + ' · status: ' + rec.status + (rec.putus_nama ? ' oleh ' + rec.putus_nama + ' ' + rec.putus_waktu : '') + (rec.kode ? ' · kode verifikasi ' + rec.kode : '') + (S(rec.catatan) ? ' · catatan: ' + S(rec.catatan) : '')]);
      var v2b = rec._data.v === 2, its = rec._data.items.filter(function (it) { var d = v2b ? (S(it.dec) || S(it.def)) : 'C'; return d === 'C' || d === 'S'; });
      if (v2b) { var cnt = {}; rec._data.items.forEach(function (it) { var d = S(it.dec) || S(it.def) || 'N'; cnt[d] = (cnt[d] || 0) + 1; }); vb.push([(rec.status === 'DISETUJUI' ? 'Keputusan: ' : 'Usulan (belum diputuskan): ') + DEC_ORDER.map(function (d) { return DEC[d] + ' ' + (cnt[d] || 0); }).join(' · ') + ' invoice. Rincian ✓ dan ◇ di bawah; lengkapnya di sheet RA-RI Termin dan arsip PDF.']); }
      vb.push(['NO', 'NOMOR INVOICE', 'PELANGGAN', 'NPP', 'INVOICE KE-', v2b ? 'VALIDASI SECTION HEAD F&HR' : 'RENCANA TGL CAIR', 'RENCANA CAIR (DPP − PPh)', 'REALISASI DI BULAN INI (KAS)', 'STATUS SAAT INI']);
      var tA = 0, tR = 0;
      its.forEach(function (it, ix) {
        var inv = ctx.invs.filter(function (i) { return normInv(i.nomor) === normInv(it.nomor); })[0];
        var real = inv ? sumBy(inv._pays, function (p) { return p.tgl >= sel.start && p.tgl <= sel.end && !p._pb ? p.nominal : 0; }) : 0;
        vb.push([ix + 1, it.nomor, it.pelanggan, it.npp, it.invke, v2b ? DEC[S(it.dec) || S(it.def)] : D(it.rencana), N(it.nilai), real, inv ? inv._res.status : '—']); tA += N(it.nilai); tR += real;
      });
      vb.push(['', 'TOTAL', '', '', '', '', tA, tR, tA ? 'Realisasi ' + pctTxt(tR / tA) : '']);
    }
    var hdrAt = 0; vb.forEach(function (r, ix) { if (r[0] === 'NO') hdrAt = ix; });
    blocks.push(block(r2, 1, vb, { styles: [sty(0, 'title')].concat(hdrAt ? [sty(hdrAt, 'hdr'), sty(vb.length - 1, 'total')] : []), fmt: hdrAt ? { 5: 'date', 6: 'money', 7: 'money' } : {} }));
    // tagihan kekurangan yang ditagih di luar invoice (berita acara / memo antar perusahaan)
    var BA = open.filter(function (i) { return /BA\/MEMO/.test(i._res.status); }).sort(function (a, b) { return a.tgl < b.tgl ? -1 : 1; }), r3 = r2 + vb.length + 2;
    var br = [['TAGIHAN KEKURANGAN NON-INVOICE (DITAGIH VIA BERITA ACARA / MEMO) — tetap piutang sampai dibayar; catat pembayarannya ke nomor invoice asal'],
      ['NO', 'NAMA PELANGGAN', 'NOMOR INVOICE ASAL', 'TGL INVOICE', 'NPP', 'NO. BA / MEMO', 'TGL BA / MEMO', 'KEKURANGAN (SISA)', 'UMUR INVOICE (HARI)', 'CATATAN']];
    BA.forEach(function (i, ix) { br.push([ix + 1, i.pelanggan, i.nomor, i.tgl, i.npp, S(i.ba_no), D(i.ba_tgl), i._res.sisa, i._res.umur, S(i.catatan)]); });
    if (!BA.length) br.push(['', 'Tidak ada']);
    br.push(['', 'TOTAL', '', '', '', '', '', sumBy(BA, function (i) { return i._res.sisa; })]);
    blocks.push(block(r3, 1, br, { styles: [sty(0, 'title'), sty(1, 'hdr'), sty(br.length - 1, 'total')], fmt: { 3: 'date', 6: 'date', 7: 'money', 8: 'int' } }));
    return { from: 8, blocks: blocks, wrapCols: [20] };
  }

  // ------------------------------------------------------------------ REKAP INVOICE KELUAR & BUKU PENJUALAN
  function repRekapKeluar(ctx) {
    var sel = parseMonthSel(ctx.sel['Rekap Invoice Keluar'], ctx.today, true), y = sel.y;
    var act = ctx.invs.filter(function (i) { return !i._batal && i.tgl; });
    var rows = [['REKAP INVOICE KELUAR TAHUN ' + y], ['BULAN', 'JUMLAH INVOICE', 'DPP', 'PPN', 'TOTAL INVOICE', 'RETENSI', 'NETTO TERCATAT', 'SUDAH LUNAS', 'MASIH TERBUKA', 'FAKTUR TERCATAT']];
    for (var m = 1; m <= 12; m++) {
      var s0 = monthStart(y, m), s1 = monthEnd(y, m), L = act.filter(function (i) { return i.tgl >= s0 && i.tgl <= s1; });
      rows.push([monthLabel(y, m), L.length, sumBy(L, function (i) { return i._a.dpp; }), sumBy(L, function (i) { return i._a.ppn; }), sumBy(L, function (i) { return i._a.total; }),
        sumBy(L, function (i) { return i._a.ret; }), sumBy(L, function (i) { return i._a.netto; }), L.filter(function (i) { return i._res.closed; }).length, L.filter(function (i) { return !i._res.closed; }).length,
        L.filter(function (i) { return i.no_fp; }).length]);
    }
    var body = rows.slice(2);
    rows.push(['TOTAL ' + y].concat([1, 2, 3, 4, 5, 6, 7, 8, 9].map(function (j) { return sumBy(body, function (r) { return r[j]; }); })));
    var blocks = [block(8, 1, rows, { styles: [sty(0, 'title'), sty(1, 'hdr'), sty(rows.length - 1, 'total')], fmt: { 1: 'int', 2: 'money', 3: 'money', 4: 'money', 5: 'money', 6: 'money', 7: 'int', 8: 'int', 9: 'int' }, span: 2 })];
    var L = act.filter(function (i) { return i.tgl >= sel.start && i.tgl <= sel.end; });
    L.sort(function (a, b) { return a.tgl < b.tgl ? -1 : a.tgl > b.tgl ? 1 : (a.nomor < b.nomor ? -1 : 1); });
    var r2 = 8 + rows.length + 2;
    var lr = [['DAFTAR INVOICE KELUAR — ' + sel.label.toUpperCase()], ['NO', 'TGL INVOICE', 'NOMOR INVOICE', 'NAMA PELANGGAN', 'NPP', 'NAMA PROYEK', 'INVOICE KE-', 'DPP', 'PPN', 'TOTAL INVOICE', 'RETENSI',
      'NOMOR FAKTUR PAJAK', 'TGL FAKTUR', 'NO. DOK SAP', 'PENANDATANGAN', 'PERSETUJUAN PAKET', 'STATUS PEMBAYARAN']];
    L.forEach(function (i, ix) { lr.push([ix + 1, i.tgl, i.nomor, i.pelanggan, i.npp, i._pro ? i._pro.singkat : '', S(i.invke), i._a.dpp, i._a.ppn, i._a.total, i._a.ret,
      i.no_fp, i.tgl_fp, i.sap_doc, S(i.ttd), apprStatusText(ctx.latest(DOC.INV, normInv(i.nomor)), i._fp), i._res.status]); });
    lr.push(['', 'TOTAL', '', '', '', '', '', sumBy(L, function (i) { return i._a.dpp; }), sumBy(L, function (i) { return i._a.ppn; }), sumBy(L, function (i) { return i._a.total; }), sumBy(L, function (i) { return i._a.ret; })]);
    blocks.push(block(r2, 1, lr, { styles: [sty(0, 'title'), sty(1, 'hdr'), sty(lr.length - 1, 'total')], fmt: { 1: 'date', 7: 'money', 8: 'money', 9: 'money', 10: 'money', 12: 'date' } }));
    // buku penjualan (format BUK PEN bagian Pajak) — berdasarkan tanggal faktur
    var BP = bukuPenjualan(ctx, sel.start, sel.end), r3 = r2 + lr.length + 2;
    var bk = [['BUKU PENJUALAN (FORMAT BUK PEN BAGIAN PAJAK) — ' + sel.label.toUpperCase() + ' · berdasarkan tanggal faktur; harga jual = DPP pasti dari invoice · ekspor: menu AR CSU ▸ Ekspor Buku Penjualan'],
      ['NO', 'KODE FP', 'KD STATUS', 'NOMOR SERI FP', 'NOMOR FAKTUR', 'TANGGAL FP', 'NPWP', 'NAMA PELANGGAN', 'URAIAN FAKTUR PAJAK', 'HARGA JUAL', 'DPP NILAI LAIN', 'PPN', 'NO SAP', 'NOMOR INVOICE', 'CEK']];
    BP.forEach(function (x, ix) { bk.push([ix + 1, x.kode, x.status, x.seri, x.nomorFP, x.tglFP, x.npwp, x.nama, x.uraian, x.harga, x.dppnl, x.ppn, x.sap, x.nomor, x.cek]); });
    bk.push(['', 'TOTAL', '', '', '', '', '', '', '', sumBy(BP, function (x) { return x.harga; }), sumBy(BP, function (x) { return x.dppnl; }), sumBy(BP, function (x) { return x.ppn; })]);
    blocks.push(block(r3, 1, bk, { styles: [sty(0, 'title'), sty(1, 'hdr'), sty(bk.length - 1, 'total')], fmt: { 5: 'date', 9: 'money', 10: 'money', 11: 'money' } }));
    return { from: 8, blocks: blocks };
  }

  // ------------------------------------------------------------------ KARTU PIUTANG
  function repKartu(ctx) {
    var npp = S(ctx.sel['Kartu Piutang']);
    if (!npp || !ctx.M.pro[npp]) npp = Object.keys(ctx.M.pro)[0] || '';
    if (!npp) return { from: 7, blocks: [block(7, 1, [['Belum ada proyek di Master Proyek']])] };
    var asOf = D(ctx.sel['Kartu Piutang|asof']) || ctx.today; if (asOf > ctx.today) asOf = ctx.today;
    var kb = kartuBlocks(ctx, npp, 7, asOf), last = kb.blocks[kb.blocks.length - 1], r = last.r + last.rows.length + 2;
    kb.blocks.push(kartuSignBlock(ctx, npp, asOf, r));
    return kb;
  }
  /** tempat, tanggal persetujuan, tanda tangan AR Officer & Section Head (tanpa NIP) */
  function kartuSignBlock(ctx, npp, asOf, r) {
    var P = ctx.params, mk = userByRole(P, 'PEMBUAT') || {}, sh = userByRole(P, 'PENYETUJU') || {}, ref = npp + '|' + asOf;
    var rec = ctx.latest(DOC.KP, ref), fp = kpFp(ctx, npp, asOf), ok = rec && rec.status === 'DISETUJUI' && S(rec.sidik) === fp;
    var sp = function (a, b) { var x = ['', a, '', '', '', '', '', b]; return x; };
    // v2.5: urutan dokumen internal — yang menyetujui di kiri, pembuat di kanan
    var stAR = rec && (ok || rec.status === 'MENUNGGU') ? '✓ Diajukan secara elektronik ' + rec.waktu + ' WIB' : '';
    var stSH = ok ? (isDelegasi(rec) ? '✓ DISETUJUI a.n. SECTION HEAD oleh ' + (rec.putus_nama || rec.putus_oleh) + ' (DELEGASI) · ' : '✓ DISETUJUI SECARA ELEKTRONIK · ') + rec.putus_waktu + ' WIB · kode ' + S(rec.kode) : (rec && rec.status === 'MENUNGGU' && S(rec.sidik) === fp ? '⏳ menunggu persetujuan elektronik' : '');
    var rows = [sp('', (S(P.KOTA) || 'Jakarta') + ', ' + (ok ? tglPanjang(D(rec.putus_waktu)) : '...............................')), sp('Disetujui oleh,', 'Dibuat oleh,'),
      sp(stSH, stAR), sp('', ''), sp('', ''), sp(sh.nama || '(..............................)', mk.nama || '(..............................)'), sp(sh.jabatan || 'Section Head of Finance & HR', mk.jabatan || 'AR Officer')];
    if (!ok) {
      var any = ctx.appr.filter(function (a) { return a.jenis === DOC.KP && S(a.ref).indexOf(npp + '|') === 0 && a.status === 'DISETUJUI'; }).pop();
      rows.push([any ? 'Persetujuan elektronik terakhir untuk proyek ini: posisi per ' + tglPanjang(S(any.ref).split('|')[1]) + ' (disetujui ' + any.putus_waktu + '). Isi POSISI PER di E5 dengan tanggal tersebut untuk menampilkan capnya, atau ajukan kartu posisi ini di aplikasi ▸ Persetujuan ▸ Kartu piutang.'
        : 'Belum ada persetujuan elektronik untuk kartu posisi ini — tanda tangan basah, atau ajukan dari aplikasi ▸ Persetujuan ▸ Kartu piutang.']);
    }
    return block(r, 1, rows, { styles: [sty(1, 'bold'), sty(2, 'stamp'), sty(5, 'signname'), ok ? null : sty(7, 'note')].filter(Boolean) });
  }
  /** retensi dalam kalimat yang mudah dipahami pembaca kartu */
  /** uang muka proyek: Master Proyek, atau jumlah invoice uang muka bila master kosong */
  function umBasisOf(ctx, npp) { var p = ctx.M.pro[npp] || {}; if (p.um > 0) return p.um; return sumBy(ctx.invs.filter(function (i) { return i.npp === npp && !i._batal && i._a.jenis === 'UANG MUKA'; }), function (i) { return i._a.dpp; }); }
  function retTxt(pct, pos) { if (!pct) return 'Tidak ada'; return pctTxt(pct) + ' dari progres, ' + (U(pos) === 'AKHIR' ? 'ditagih di akhir' : 'dipotong sebelum DPP'); }
  function kartuBlocks(ctx, npp, top, asOf) {
    ctx = contextAt(ctx, asOf);
    var k = kartuData(ctx, npp, asOf), p = k.pro, pel = k.pel, kk = k.kontrak, T = k.T;
    var nT = p.adds.filter(function (a) { return a.jenis === 'TAMBAH'; }).length, nK = p.adds.filter(function (a) { return a.jenis === 'KURANG'; }).length;
    var um = umBasisOf(ctx, npp);
    var info = [
      [p.pelanggan + '  ·  ' + (p.nama_pe || p.singkat) + '  ·  POSISI PER ' + tglPanjang(k.asOf).toUpperCase()],
      ['NPP', npp, '', 'Nilai kontrak awal', '', kk.awal, '', 'Uang muka', um ? 'Rp ' + fmtRp(um) : 'Tidak ada'],
      ['No. Kontrak/SPK', S(p.kontrak), '', 'Adendum tambah (' + nT + '×)', '', kk.tambah || 0, '', 'Retensi', retTxt(p.ret_pct, p.ret_pos)],
      ['Tgl. Kontrak', tglPanjang(D(p.tgl_kontrak)), '', 'Adendum kurang (' + nK + '×)', '', kk.kurang ? -kk.kurang : 0, '', 'PPh', pphLabel(npp, p.tarif_eff)],
      ['Proyek', p.nama_pe || S(p.pekerjaan) || p.singkat, '', 'Nilai setelah adendum', '', kk.net, '', 'Penandatangan invoice', p.signer ? p.signer.nama : '-'],
      ['Lokasi', S(p.lokasi), '', 'PPN 11%', '', rnd(kk.net * 0.11), '', 'Kelompok', (pel.kelompok || '') + (pel.jenis ? ' · ' + pel.jenis : '')],
      ['Kode BP · NPWP', (pel.kode_bp || '-') + ' · ' + (pel.npwp16 || '-'), '', 'Nilai + PPN', '', rnd(kk.net * 1.11), '', 'NPP produk', nppListTxt(p.produk) || '-']
    ];
    if (p.adendumNpp.length) info.push(['', '', '', '', '', '', '', 'NPP adendum', nppListTxt(p.adendumNpp) + ' — ditagih lewat NPP ini']);
    var b1 = block(top, 1, info, { styles: [sty(0, 'title'), sty(1, 'lbl', info.length - 1)], fmt: { 5: 'money' } });
    var sm = [['RINGKASAN', 'DITAGIH (NILAI PROGRES)', '% DARI KONTRAK', 'DPP + PPN DITAGIH', 'DITERIMA (KAS)', 'PPh & ADM DIPOTONG', 'SISA TAGIHAN', 'SISA NETTO', 'SISA PPN', 'RETENSI DITAHAN', 'SISA KONTRAK BELUM DITAGIH'],
      ['per ' + tglPanjang(k.asOf), T.brutoProg, k.pctTagih, T.debit, T.kas, T.pot, T.saldo, T.sN, T.sP, T.ret, k.sisaKontrak]];
    var r2 = top + info.length + 1, b2 = block(r2, 1, sm, { styles: [sty(0, 'hdr'), sty(1, 'kpi')], fmt: { 1: 'money', 2: 'pct', 3: 'money', 4: 'money', 5: 'money', 6: 'money', 7: 'money0', 8: 'money0', 9: 'money', 10: 'money' } });
    // A. rincian tagihan per invoice
    var ra = [['A. RINCIAN TAGIHAN PER INVOICE — nilai progres asli, potongan uang muka & retensi, dasar pengenaan pajak'],
      ['TGL INVOICE', 'NO. INVOICE · TAGIHAN', 'NILAI PROGRES', 'POT. UANG MUKA', 'RETENSI', 'DPP', 'PPN', 'TOTAL INVOICE (DPP + PPN)', 'PIUTANG DICATAT']], sa = [sty(0, 'title'), sty(1, 'hdr')], TA = [0, 0, 0, 0, 0, 0, 0], anyLama = false, anyAkhir = false;
    k.groups.forEach(function (g) {
      var A = g.A, lb = A.nomor + ' · ' + A.invke + (A.jenis === 'UANG MUKA' ? ' (uang muka)' : A.jenis === 'RETENSI' ? ' (pelepasan retensi)' : '') + (A.lama ? ' *' : '') + (A.retpos === 'AKHIR' && A.ret ? ' **' : '');
      ra.push([A.tgl, lb, A.progres || '', A.umk || '', A.ret || '', A.dpp, A.ppn, A.total, A.dicatat]);
      TA[0] += A.progres; TA[1] += A.umk; TA[2] += A.ret; TA[3] += A.dpp; TA[4] += A.ppn; TA[5] += A.total; TA[6] += A.dicatat;
      if (A.lama) anyLama = true; if (A.retpos === 'AKHIR' && A.ret) anyAkhir = true;
    });
    if (!k.groups.length) ra.push(['', 'Belum ada invoice untuk proyek ini']);
    ra.push(['', 'TOTAL'].concat(TA)); sa.push(sty(ra.length - 1, 'total'));
    var na = 'Nilai progres − pot. uang muka − retensi (bila dipotong sebelum DPP) = DPP. Piutang dicatat = DPP + PPN − retensi yang ditahan dari pembayaran.';
    if (anyAkhir) na += ' ** retensi ditahan dari pembayaran (ditagih di akhir).';
    if (anyLama) na += ' * invoice lama: nilai progres / pot. UM belum dilengkapi — isi kolom NILAI PROGRES / BRUTO, PENGEMBALIAN UM, RETENSI di Monitoring Invoice.';
    ra.push([na]); sa.push(sty(ra.length - 1, 'note'));
    var r3 = r2 + sm.length + 1, b3 = block(r3, 1, ra, { styles: sa, fmt: { 0: 'date', 2: 'money', 3: 'money', 4: 'money', 5: 'money', 6: 'money', 7: 'money', 8: 'money' } });
    // B. mutasi piutang (debit / kredit, saldo per invoice)
    var rows = [['B. MUTASI PIUTANG — debit = tagihan, kredit = penerimaan & potongan; saldo di baris "Sisa tagihan" tiap invoice'],
      ['TANGGAL', 'URAIAN', 'REF / BUKTI', 'NO. DOK SAP', 'DEBIT (Rp)', 'KREDIT (Rp)', 'SALDO (Rp)', 'KETERANGAN']], styles = [sty(0, 'title'), sty(1, 'hdr')];
    var z = function (x) { return x ? x : ''; };
    k.groups.forEach(function (g) {
      g.rows.forEach(function (r) {
        rows.push([r.tgl, r.uraian, r.ref, r.sap, z(r.debit), z(r.kredit), r.kind === 'sisa' ? r.saldo : '', r.ket]);
        if (r.kind === 'inv') styles.push(sty(rows.length - 1, 'band')); else if (r.kind === 'sisa') styles.push(sty(rows.length - 1, 'sisa'));
      });
    });
    if (!k.groups.length) rows.push(['', 'Belum ada invoice untuk proyek ini']);
    rows.push(['', 'TOTAL PROYEK', '', '', T.debit, T.kredit, T.saldo, 'Sisa Netto ' + fmtRp(T.sN) + ' · Sisa PPN ' + fmtRp(T.sP)]); styles.push(sty(rows.length - 1, 'total'));
    rows.push(['Debit = tagihan (dan kelebihan bayar yang dialihkan ke invoice lain). Kredit = penerimaan Netto & PPN, PPh dipotong pelanggan, biaya adm bank, pindah buku masuk. Saldo = sisa tagihan tiap invoice.']); styles.push(sty(rows.length - 1, 'note'));
    var b4 = block(r3 + ra.length + 1, 1, rows, { styles: styles, fmt: { 0: 'date', 4: 'money', 5: 'money', 6: 'money0' } });
    return { from: top, blocks: [b1, b2, b3, b4], wrapCols: [2, 8], minW: { 1: 15 } };
  }

  // ------------------------------------------------------------------ MONITORING UANG MUKA & RETENSI
  function repUmRet(ctx) {
    var R = umRetRows(ctx), um = R.filter(function (u) { return u.umSt; }), rt = R.filter(function (u) { return u.ret; });
    var a = [['A. UANG MUKA — sisa UM yang masih boleh dipotong. Saat membuat invoice, aplikasi membatasi potongan UM sampai sisa ini.'],
      ['NO', 'PELANGGAN', 'NPP', 'NAMA PROYEK', 'STATUS PROYEK', 'NILAI KONTRAK (OK)', 'UM DI MASTER', 'UM DITAGIH (INVOICE UM)', 'UM DASAR PENGEMBALIAN', 'DIKEMBALIKAN SEBELUM SISTEM', 'DIKEMBALIKAN VIA INVOICE', 'SISA UM', '% UM DIKEMBALIKAN', '% PROGRES DITAGIH', 'INVOICE TERAKHIR MEMOTONG UM', 'STATUS']];
    um.forEach(function (u, ix) { a.push([ix + 1, u.pelanggan, u.npp, u.singkat, u.aktif, u.nilai_ok, u.umMaster, u.umInv, u.basis, u.awal, u.kembali, u.lengkap ? u.sisa : '(isi saldo awal)', u.lengkap ? u.pctUm : '', u.pctProg, u.lastUm, u.umSt]); });
    if (!um.length) a.push(['', 'Belum ada proyek dengan uang muka']);
    a.push(['', 'TOTAL', '', '', '', '', '', sumBy(um, function (u) { return u.umInv; }), sumBy(um, function (u) { return u.basis; }), sumBy(um, function (u) { return u.awal; }), sumBy(um, function (u) { return u.kembali; }), sumBy(um, function (u) { return u.lengkap ? u.sisa : 0; })]);
    var b = [['B. RETENSI — retensi yang ditahan pelanggan vs yang sudah ditagih'],
      ['PELANGGAN', 'NPP', 'NAMA PROYEK', 'POSISI RETENSI', '% RETENSI', 'BATAS (% × NILAI KONTRAK)', 'DITAHAN (DIPOTONG DI AWAL)', 'DITAHAN (DIPOTONG DI AKHIR)', 'TOTAL DITAHAN', 'SUDAH DITAGIH (INVOICE RETENSI)', 'SISA BELUM DITAGIH', 'STATUS']];
    rt.forEach(function (u) { var r = u.ret; b.push([r[0], r[1], r[2], r[3], r[4], u.batas || '', r[5], r[6], r[7], r[8], r[9], u.retSt]); });
    b.push(['TOTAL', '', '', '', '', '', sumBy(rt, function (u) { return u.ret[5]; }), sumBy(rt, function (u) { return u.ret[6]; }), sumBy(rt, function (u) { return u.ret[7]; }), sumBy(rt, function (u) { return u.ret[8]; }), sumBy(rt, function (u) { return u.ret[9]; }), '']);
    var r2 = 6 + a.length + 2;
    return { from: 6, blocks: [block(6, 1, a, { styles: [sty(0, 'title'), sty(1, 'hdr'), sty(a.length - 1, 'total')], fmt: { 5: 'money', 6: 'money', 7: 'money', 8: 'money', 9: 'money', 10: 'money', 11: 'money', 12: 'pct', 13: 'pct' } }),
      block(r2, 1, b, { styles: [sty(0, 'title'), sty(1, 'hdr'), sty(b.length - 1, 'total')], fmt: { 4: 'pct', 5: 'money', 6: 'money', 7: 'money', 8: 'money', 9: 'money', 10: 'money' }, span: 2 })], wrapCols: [16] };
  }
  function repRetensi(ctx) {
    var rows = [['PELANGGAN', 'NPP', 'NAMA PROYEK', 'POSISI RETENSI', '% RETENSI', 'DITAHAN (DIPOTONG DI AWAL)', 'DITAHAN (DIPOTONG DI AKHIR)', 'TOTAL DITAHAN', 'SUDAH DITAGIH (INVOICE RETENSI)', 'SISA BELUM DITAGIH', 'CATATAN']];
    var R = retensiRows(ctx);
    rows = rows.concat(R);
    rows.push(['TOTAL', '', '', '', '', sumBy(R, function (r) { return r[5]; }), sumBy(R, function (r) { return r[6]; }), sumBy(R, function (r) { return r[7]; }), sumBy(R, function (r) { return r[8]; }), sumBy(R, function (r) { return r[9]; }), '']);
    return { from: 6, blocks: [block(6, 1, rows, { styles: [sty(0, 'hdr'), sty(rows.length - 1, 'total')], fmt: { 4: 'pct', 5: 'money', 6: 'money', 7: 'money', 8: 'money', 9: 'money' } })] };
  }

  // ------------------------------------------------------------------ INKASO & REKAP CAIR
  function repInkaso(ctx) {
    var sel = parseMonthSel(ctx.sel['INKASO'], ctx.today, false), y = sel.y;
    var P = ctx.pays.filter(function (p) { return p._inv && p.tgl && p.nominal > 0 && !p._pb; });
    var inM = P.filter(function (p) { return p.tgl >= sel.start && p.tgl <= sel.end; }).sort(function (a, b) { return a.tgl < b.tgl ? -1 : a.tgl > b.tgl ? 1 : a._row - b._row; });
    var tg = function (m) { return ctx.params.targets[y + '-' + pad(m, 2)] || 0; };
    var target = tg(sel.m), real = sumBy(inM, function (p) { return p.nominal; });
    var ytdT = 0, ytdR = 0; for (var m0 = 1; m0 <= sel.m; m0++) ytdT += tg(m0);
    ytdR = sumBy(P.filter(function (p) { return p.tgl >= monthStart(y, 1) && p.tgl <= sel.end; }), function (p) { return p.nominal; });
    var head = [['', 'BULAN ' + MONTHS[sel.m - 1].toUpperCase() + ' ' + y, '%', 's.d. ' + MONTHS[sel.m - 1].toUpperCase(), '%'],
      ['TARGET INKASO', target, target ? 1 : '', ytdT, ytdT ? 1 : ''], ['REALISASI', real, target ? real / target : '', ytdR, ytdT ? ytdR / ytdT : ''],
      ['KURANG / LEBIH', real - target, target ? (real - target) / target : '', ytdR - ytdT, ytdT ? (ytdR - ytdT) / ytdT : ''],
      ['  · porsi Netto', sumBy(inM, function (p) { return p._keN; }), '', '', ''], ['  · porsi PPN (wajib disetor)', sumBy(inM, function (p) { return p._keP; }), '', '', '']];
    var J = {}; inM.forEach(function (p) { var j = (p._inv._pel && p._inv._pel.jenis) || 'SWASTA', o = J[j] || (J[j] = [0, 0, 0, 0]); o[0]++; o[1] += p.nominal; o[2] += p._keN; o[3] += p._keP; });
    var RK = {}; inM.forEach(function (p) { var k = p._rek ? p._rek.label : (S(p.rekening) || '(belum diisi)'), o = RK[k] || (RK[k] = [0, 0]); o[0]++; o[1] += p.nominal; });
    var rk = [['PER REKENING PENERIMAAN — ' + sel.label.toUpperCase(), 'TRANSAKSI', 'TOTAL DITERIMA', '% DARI REALISASI']];
    (ctx.params.rek || []).forEach(function (r) { var o = RK[r.label] || [0, 0]; rk.push([r.label, o[0], o[1], real ? o[1] / real : 0]); delete RK[r.label]; });
    Object.keys(RK).forEach(function (k) { rk.push([k, RK[k][0], RK[k][1], real ? RK[k][1] / real : 0]); });
    rk.push(['TOTAL', inM.length, real, real ? 1 : 0]);
    var jr = [['PER JENIS PELANGGAN — ' + sel.label.toUpperCase(), 'TRANSAKSI', 'TOTAL DITERIMA', '→ NETTO', '→ PPN', '% DARI REALISASI']];
    ['WIKA', 'BUMN', 'SWASTA', 'JO'].forEach(function (j) { var o = J[j] || [0, 0, 0, 0]; jr.push([j, o[0], o[1], o[2], o[3], real ? o[1] / real : 0]); });
    jr.push(['TOTAL', inM.length, real, sumBy(inM, function (p) { return p._keN; }), sumBy(inM, function (p) { return p._keP; }), real ? 1 : 0]);
    var yr = [['TREN ' + y], ['BULAN', 'TARGET', 'REALISASI', '%', 'KURANG / LEBIH', '→ NETTO', '→ PPN', 'TRANSAKSI', 'INVOICE LUNAS']];
    for (var m = 1; m <= 12; m++) {
      var s0 = monthStart(y, m), s1 = monthEnd(y, m), L = P.filter(function (p) { return p.tgl >= s0 && p.tgl <= s1; }), rr = sumBy(L, function (p) { return p.nominal; }), t = tg(m);
      var lun = ctx.invs.filter(function (i) { return i._res.closed && !i._batal && i._res.tgl_bayar >= s0 && i._res.tgl_bayar <= s1; }).length;
      yr.push([monthLabel(y, m), t, rr, t ? rr / t : '', rr - t, sumBy(L, function (p) { return p._keN; }), sumBy(L, function (p) { return p._keP; }), L.length, lun]);
    }
    var b = yr.slice(2); yr.push(['TOTAL ' + y, sumBy(b, function (r) { return r[1]; }), sumBy(b, function (r) { return r[2]; }), '', sumBy(b, function (r) { return r[4]; }), sumBy(b, function (r) { return r[5]; }), sumBy(b, function (r) { return r[6]; }), sumBy(b, function (r) { return r[7]; }), sumBy(b, function (r) { return r[8]; })]);
    var dr = [['RINCIAN PENERIMAAN — PERIODE 01 s/d ' + monthEnd(y, sel.m).substring(8) + ' ' + MONTHS[sel.m - 1].toUpperCase() + ' ' + y],
      ['NO', 'TGL BAYAR', 'NAMA PELANGGAN', 'JENIS', 'NPP', 'NOMOR INVOICE', 'INVOICE KE-', 'NOMINAL DITERIMA', '→ NETTO', '→ PPN', 'POT. PPh DIAKUI', 'ADM BANK', 'STATUS INVOICE', 'REF. TRANSFER', 'DASAR ALOKASI', 'KETERANGAN', 'REKENING PENERIMAAN']];
    inM.forEach(function (p, ix) { var i = p._inv; dr.push([ix + 1, p.tgl, i.pelanggan, (i._pel && i._pel.jenis) || '', i.npp, i.nomor, invLabel(i), p.nominal, p._keN, p._keP, (p._pph || 0) + (p._sel || 0), p._adm || 0, i._res.status, S(p.ref), p._rule, S(p.ket), p._rek ? p._rek.label : S(p.rekening)]); });
    dr.push(['', 'TOTAL', '', '', '', '', '', real, sumBy(inM, function (p) { return p._keN; }), sumBy(inM, function (p) { return p._keP; }), sumBy(inM, function (p) { return (p._pph || 0) + (p._sel || 0); }), sumBy(inM, function (p) { return p._adm || 0; })]);
    var r1 = 8, r2 = r1 + head.length + 1, r3 = r2 + Math.max(jr.length, rk.length) + 1, r4 = r3 + yr.length + 1;
    return { from: 8, wrapCols: [16], blocks: [
      block(r1, 1, head, { styles: [sty(0, 'hdr'), sty(1, 'kpi', 3)], fmt: { 1: 'money', 2: 'pct', 3: 'money', 4: 'pct' }, span: 2 }),
      block(r2, 1, jr, { styles: [sty(0, 'hdr'), sty(jr.length - 1, 'total')], fmt: { 1: 'int', 2: 'money', 3: 'money', 4: 'money', 5: 'pct' }, span: 2 }),
      block(r2, 9, rk, { styles: [sty(0, 'hdr'), sty(rk.length - 1, 'total')], fmt: { 1: 'int', 2: 'money', 3: 'pct' }, span: 2 }),
      block(r3, 1, yr, { styles: [sty(0, 'title'), sty(1, 'hdr'), sty(yr.length - 1, 'total')], fmt: { 1: 'money', 2: 'money', 3: 'pct', 4: 'money', 5: 'money', 6: 'money', 7: 'int', 8: 'int' }, span: 2 }),
      block(r4, 1, dr, { styles: [sty(0, 'title'), sty(1, 'hdr'), sty(dr.length - 1, 'total')], fmt: { 1: 'date', 7: 'money', 8: 'money', 9: 'money', 10: 'money', 11: 'money' } })] };
  }


  // ------------------------------------------------------------------ AGING
  function repAging(ctx) {
    var open = openInv(ctx), bp = {};
    open.forEach(function (i) {
      var k = i.pelanggan, o = bp[k] || (bp[k] = { kel: i._pel ? S(i._pel.kelompok) : '', jenis: i._pel ? i._pel.jenis : '', v: [0, 0, 0, 0, 0, 0, 0] });
      var s = i._res.sisa, jt = i._res.jt, idx;
      if (!jt) idx = 0; else if (ctx.today <= jt) idx = 1; else { var d = diffDays(ctx.today, jt); idx = d <= 30 ? 2 : d <= 90 ? 3 : d <= 180 ? 4 : 5; }
      o.v[idx] += s; o.v[6] += s;
    });
    var rows = [['NO', 'PELANGGAN', 'KELOMPOK', 'JENIS', 'TOTAL PIUTANG', 'BELUM SUBMIT / TANPA JT', 'BELUM JATUH TEMPO', 'LEWAT JT 0–30 HARI', '31–90 HARI', '91–180 HARI', '> 180 HARI']];
    var keys = Object.keys(bp).sort(function (a, b) { return bp[b].v[6] - bp[a].v[6]; }), T = [0, 0, 0, 0, 0, 0, 0];
    keys.forEach(function (k, ix) { var v = bp[k].v; rows.push([ix + 1, k, bp[k].kel, bp[k].jenis, v[6], v[0], v[1], v[2], v[3], v[4], v[5]]); for (var j = 0; j < 7; j++) T[j] += v[j]; });
    rows.push(['', 'TOTAL', '', '', T[6], T[0], T[1], T[2], T[3], T[4], T[5]]);
    var g = {};
    keys.forEach(function (k) { var j = bp[k].jenis || 'SWASTA', o = g[j] || (g[j] = [0, 0, 0, 0, 0, 0, 0]); for (var x = 0; x < 7; x++) o[x] += bp[k].v[x]; });
    var gr = [['PER JENIS PELANGGAN', 'TOTAL', '%', 'BELUM SUBMIT', 'BELUM JT', '0–30', '31–90', '91–180', '> 180']];
    ['WIKA', 'BUMN', 'SWASTA', 'JO'].forEach(function (j) { var o = g[j] || [0, 0, 0, 0, 0, 0, 0]; gr.push([j, o[6], T[6] ? o[6] / T[6] : 0, o[0], o[1], o[2], o[3], o[4], o[5]]); });
    gr.push(['TOTAL', T[6], T[6] ? 1 : 0, T[0], T[1], T[2], T[3], T[4], T[5]]);
    var r2 = 6 + rows.length + 2;
    return { from: 6, blocks: [block(6, 1, rows, { styles: [sty(0, 'hdr'), sty(rows.length - 1, 'total')], fmt: { 4: 'money', 5: 'money', 6: 'money', 7: 'money', 8: 'money', 9: 'money', 10: 'money' } }),
      block(r2, 2, gr, { styles: [sty(0, 'hdr'), sty(gr.length - 1, 'total')], fmt: { 1: 'money', 2: 'pct', 3: 'money', 4: 'money', 5: 'money', 6: 'money', 7: 'money', 8: 'money' } })] };
  }

  // ------------------------------------------------------------------ MONITORING TAGBRUT
  function repTagbrut(ctx) {
    var v = ctx.sel['Monitoring Tagbrut'], sel = parseMonthSel(v, ctx.today, false), C = sel.auto ? ctx.today : sel.end;
    var tg = tagAt(ctx, C), r = N(ctx.params.CKPN_DISKONTO) || 0.0833, rel = U(ctx.params.CKPN_PIHAK_BERELASI) === 'YA', batas = N(ctx.params.TAGBRUT_BATAS_HARI) || 30;
    var lastNote = {}; ctx.notes.forEach(function (n) { var k = U(n.npp); if (!lastNote[k] || (n.tgl || '') >= (lastNote[k].tgl || '')) lastNote[k] = n; });
    var keys = Object.keys(tg).sort(function (a, b) { return tg[b].sisa - tg[a].sisa; });
    var rows = [['TAGBRUT / HAD PER PROYEK — POSISI ' + tglPanjang(C).toUpperCase() + (sel.auto ? ' (HARI INI)' : '')],
      ['NO', 'PELANGGAN', 'NPP', 'NAMA PROYEK', 'AKUN', 'KELOMPOK', 'PENGAKUAN (P&E / LK)', 'DITAGIH SETELAH SALDO AWAL', 'SISA BELUM DITAGIH', 'DITAGIH MELEBIHI PENGAKUAN', 'PENGAKUAN TERTUA BELUM DITAGIH', 'UMUR (HARI)', 'CKPN (NPV)', 'STATUS', 'PIC', 'CATATAN TERAKHIR']];
    var tot = [0, 0, 0, 0, 0];
    keys.forEach(function (k, ix) {
      var t = tg[k], umur = t.oldest ? diffDays(C, t.oldest) : '', ck = (t.berelasi && !rel) ? 0 : sumBy(t.lines, function (l) { return rnd(l.saldo - npv(l.saldo, l.umur, r)); });
      var st = t.sisa > 0 ? (umur > batas ? '⚠ SEGERA TERBITKAN TAGIHAN' : 'BELUM DITAGIH') : (t.hdd > 0 ? 'DITAGIH MELEBIHI PENGAKUAN' : '✓ SUDAH DITAGIH');
      var nt = lastNote[U(k)] || lastNote[U(t.pro.pelanggan)] || {};
      rows.push([ix + 1, t.pro.pelanggan || '', k, t.pro.singkat || '', t.akun, t.berelasi ? 'PIHAK BERELASI' : 'PIHAK KETIGA', t.pengakuan, t.ditagih, t.sisa, t.hdd, t.oldest, umur, ck, st, nt.pic || '', nt.catatan ? (nt.tgl ? tglPanjang(nt.tgl) + ': ' : '') + nt.catatan : '']);
      tot[0] += t.pengakuan; tot[1] += t.ditagih; tot[2] += t.sisa; tot[3] += t.hdd; tot[4] += ck;
    });
    if (!keys.length) rows.push(['', 'Belum ada data pengakuan pendapatan. Isi sheet Input Tagbrut (mingguan dari P&E, rekon akhir bulan dengan LK).']);
    if (ctx.tagSkip) rows.splice(1, 0, ['Catatan: hanya tagbrut JASA (NPP di Master Proyek) yang dimonitor AR. ' + ctx.tagSkip + ' baris Input Tagbrut ber-NPP di luar Master Proyek (mis. material — ranah Akuntansi / P&E) diabaikan; selisih dengan LK = porsi material.']);
    rows.push(['', 'TOTAL', '', '', '', '', tot[0], tot[1], tot[2], tot[3], '', '', tot[4]]);
    var det = [['RINCIAN PENGAKUAN YANG BELUM DITAGIH (FIFO: tagihan menghabiskan pengakuan tertua lebih dulu)'], ['NPP', 'PELANGGAN', 'AKUN', 'TGL PENGAKUAN', 'NO. BUKTI', 'SUMBER', 'NILAI PENGAKUAN', 'SISA BELUM DITAGIH', 'UMUR (HARI)', 'NILAI NPV', 'CKPN']];
    keys.forEach(function (k) { var t = tg[k]; t.lines.forEach(function (l) { var b = t.berelasi && !rel, v2 = npv(l.saldo, l.umur, r); det.push([k, t.pro.pelanggan || '', t.akun, l.tgl, l.nodok, l.sumber, l.nilai, l.saldo, l.umur, b ? l.saldo : v2, b ? 0 : rnd(l.saldo - v2)]); }); });
    var r2 = 8 + rows.length + 2;
    return { from: 8, blocks: [block(8, 1, rows, { styles: [sty(0, 'title'), sty(1, 'hdr'), sty(rows.length - 1, 'total')], fmt: { 6: 'money', 7: 'money', 8: 'money', 9: 'money', 10: 'date', 11: 'int', 12: 'money' } }),
      block(r2, 1, det, { styles: [sty(0, 'title'), sty(1, 'hdr')], fmt: { 3: 'date', 6: 'money', 7: 'money', 8: 'int', 9: 'money', 10: 'money' }, span: 2 })], wrapCols: [16] };
  }

  // ------------------------------------------------------------------ CKPN
  function repCkpn(ctx) {
    var sel = parseCutoff(ctx.sel['CKPN'], ctx.today), C = sel.end, mv = ckpnMovement(ctx, C), P = ctx.params;
    var sm = [['CKPN PER ' + tglPanjang(C).toUpperCase() + ' — dibandingkan ' + tglPanjang(mv.prevC)],
      ['AKUN', 'KELOMPOK', 'GL AKUN', 'GL CADANGAN', 'SALDO', 'CKPN PER ' + C, 'CKPN BULAN LALU', 'KENAIKAN → BEBAN', 'GL BEBAN', 'PENURUNAN → PEMULIHAN', 'GL PEMULIHAN', 'BERSIH (BEBAN − PEMULIHAN)']];
    mv.groups.forEach(function (o) {
      var k = CK_KODE[o.akun], nm = CK_GROUPS.filter(function (g) { return g[0] === o.akun; })[0][1];
      sm.push([nm, o.berelasi ? 'PIHAK BERELASI' : 'PIHAK KETIGA', gl(P, o.akun === 'PIUTANG_USAHA' ? 'PIUTANG_USAHA' : o.akun === 'PIUTANG_RETENSI' ? 'PIUTANG_RETENSI' : o.akun, o.berelasi), gl(P, 'CKPN_' + k, o.berelasi),
        o.saldo, o.ckpn, o.prev, o.naik, gl(P, 'BEBAN_CKPN_' + k, o.berelasi), o.turun, gl(P, 'PEMULIHAN_CKPN_' + k, o.berelasi), o.naik - o.turun]);
    });
    sm.push(['TOTAL', '', '', '', sumBy(mv.groups, function (o) { return o.saldo; }), sumBy(mv.groups, function (o) { return o.ckpn; }), sumBy(mv.groups, function (o) { return o.prev; }), sumBy(mv.groups, function (o) { return o.naik; }), '', sumBy(mv.groups, function (o) { return o.turun; }), '', sumBy(mv.groups, function (o) { return o.naik - o.turun; })]);
    var blocks = [block(8, 1, sm, { styles: [sty(0, 'title'), sty(1, 'hdr'), sty(sm.length - 1, 'total')], fmt: { 4: 'money', 5: 'money', 6: 'money', 7: 'money', 9: 'money', 11: 'money' }, span: 2 })];
    var note = [['METODE (mengikuti lampiran LK unit): Piutang usaha — matriks umur dari JT akuntansi (tgl invoice + ' + (N(P.JT_AKUNTANSI_HARI) || 30) + ' hari), tarif per kelas di Parameter; belum jatuh tempo 0%. ' +
      'Retensi, tagbrut & HAD — NPV: saldo ÷ (1 + ' + pctTxt(N(P.CKPN_DISKONTO) || 0.0833) + '/365)^umur. Pihak berelasi: ' + (U(P.CKPN_PIHAK_BERELASI) === 'YA' ? 'dihitung' : 'tidak dicadangkan (0)') + '. ' +
      'Saldo piutang usaha = Netto − penerimaan Netto − PPh/adm yang diakui saat porsi Netto selesai (sama dengan saldo SAP). Kenaikan per dokumen → beban, penurunan → pemulihan.' + (mv.seedGap ? ' Tagbrut/HAD: saldo awal ' + tglPanjang(D(P.TAGBRUT_SALDO_AWAL_PER)) + ', pergerakan baru dihitung mulai bulan berikutnya.' : '')]];
    var r = 8 + sm.length + 1; blocks.push(block(r, 1, note, { styles: [sty(0, 'note')] })); r += 2;
    var pu = mv.cur.filter(function (x) { return x.akun === 'PIUTANG_USAHA'; }).sort(function (a, b) { return a.pelanggan < b.pelanggan ? -1 : a.pelanggan > b.pelanggan ? 1 : (a.tgl < b.tgl ? -1 : 1); });
    var d1 = [['RINCIAN PIUTANG USAHA (MATRIKS UMUR)'], ['NO', 'PELANGGAN', 'KODE BP', 'NPP', 'NOMOR INVOICE', 'NO. DOK SAP', 'TGL TERBIT', 'JUMLAH (NETTO)', 'SALDO', 'JT AKUNTANSI', 'UMUR DARI JT (HARI)', 'KELAS', 'TARIF', 'CKPN']];
    pu.forEach(function (x, ix) { d1.push([ix + 1, x.pelanggan, x.bp, x.npp, x.ref, x.doc, x.tgl, x.jumlah, x.saldo, x.jt, x.umur, x.kelas, x.tarif, x.ckpn]); });
    d1.push(['', 'TOTAL', '', '', '', '', '', sumBy(pu, function (x) { return x.jumlah; }), sumBy(pu, function (x) { return x.saldo; }), '', '', '', '', sumBy(pu, function (x) { return x.ckpn; })]);
    blocks.push(block(r, 1, d1, { styles: [sty(0, 'title'), sty(1, 'hdr'), sty(d1.length - 1, 'total')], fmt: { 6: 'date', 7: 'money', 8: 'money', 9: 'date', 10: 'int', 12: 'pct', 13: 'money' } })); r += d1.length + 2;
    var rt = mv.cur.filter(function (x) { return x.akun === 'PIUTANG_RETENSI'; });
    var d2 = [['RINCIAN PIUTANG RETENSI (NPV)'], ['NO', 'PELANGGAN', 'KODE BP', 'NPP', 'NOMOR INVOICE (ASAL RETENSI)', 'TGL', 'RETENSI', 'SALDO BELUM DITAGIH', 'UMUR (HARI)', 'NILAI NPV', 'CKPN']];
    rt.forEach(function (x, ix) { d2.push([ix + 1, x.pelanggan, x.bp, x.npp, x.ref, x.tgl, x.jumlah, x.saldo, x.umur, x.npv, x.ckpn]); });
    d2.push(['', 'TOTAL', '', '', '', '', sumBy(rt, function (x) { return x.jumlah; }), sumBy(rt, function (x) { return x.saldo; }), '', sumBy(rt, function (x) { return x.npv; }), sumBy(rt, function (x) { return x.ckpn; })]);
    blocks.push(block(r, 1, d2, { styles: [sty(0, 'title'), sty(1, 'hdr'), sty(d2.length - 1, 'total')], fmt: { 5: 'date', 6: 'money', 7: 'money', 8: 'int', 9: 'money', 10: 'money' } })); r += d2.length + 2;
    var tb = mv.cur.filter(function (x) { return x.akun === 'TAGBRUT' || x.akun === 'HAD'; });
    var d3 = [['RINCIAN TAGIHAN BRUTO & HASIL AKAN DITERIMA (NPV) — sumber: sheet Input Tagbrut'], ['NO', 'AKUN', 'PELANGGAN', 'NPP', 'TGL PENGAKUAN', 'NO. BUKTI', 'PENGAKUAN', 'SALDO BELUM DITAGIH', 'UMUR (HARI)', 'NILAI NPV', 'CKPN']];
    tb.forEach(function (x, ix) { d3.push([ix + 1, x.akun, x.pelanggan, x.npp, x.tgl, x.doc, x.jumlah, x.saldo, x.umur, x.npv, x.ckpn]); });
    d3.push(['', 'TOTAL', '', '', '', '', sumBy(tb, function (x) { return x.jumlah; }), sumBy(tb, function (x) { return x.saldo; }), '', sumBy(tb, function (x) { return x.npv; }), sumBy(tb, function (x) { return x.ckpn; })]);
    blocks.push(block(r, 1, d3, { styles: [sty(0, 'title'), sty(1, 'hdr'), sty(d3.length - 1, 'total')], fmt: { 4: 'date', 6: 'money', 7: 'money', 8: 'int', 9: 'money', 10: 'money' } }));
    return { from: 8, blocks: blocks };
  }

  // ------------------------------------------------------------------ USULAN JURNAL SAP
  function journalsFor(ctx, sel) {
    var J = [];
    ctx.invs.filter(function (i) { return !i._batal && i.tgl >= sel.start && i.tgl <= sel.end; }).sort(function (a, b) { return a.tgl < b.tgl ? -1 : 1; }).forEach(function (i) { J.push(jInvoice(ctx, i)); });
    jPays(ctx, ctx.pays.filter(function (p) { return p._inv && !p._inv._batal && p.tgl >= sel.start && p.tgl <= sel.end && (p.nominal > 0 || p._pb); }).sort(function (a, b) { return a.tgl < b.tgl ? -1 : 1; })).forEach(function (j) { J.push(j); });
    if (sel.end < ctx.today) jCkpn(ctx, sel.end).forEach(function (j) { J.push(j); });
    return J;
  }
  function repJurnal(ctx) {
    var sel = parseMonthSel(ctx.sel['Usulan Jurnal SAP'], ctx.today, false), J = journalsFor(ctx, sel), P = ctx.params;
    var head = [['USULAN JURNAL SAP — ' + sel.label.toUpperCase() + '  ·  Company Code ' + S(P.SAP_COMPANY_CODE) + ' · Unit ' + S(P.SAP_KODE_UNIT) + '  ·  diserahkan ke Akuntansi untuk diposting; nomor dokumen SAP dicatat balik di Monitoring Invoice / Log Penerimaan']];
    var tb = {}; J.forEach(function (j) { j.lines.forEach(function (l) { var k = l.gl || l.kode, o = tb[k] || (tb[k] = { gl: l.gl, nama: l.nama, d: 0, k: 0 }); o.d += l.debit; o.k += l.kredit; }); });
    var tr = [['RINGKASAN PER AKUN'], ['AKUN GL', 'NAMA AKUN', 'DEBIT', 'KREDIT', 'MUTASI BERSIH']];
    Object.keys(tb).sort().forEach(function (k) { var o = tb[k]; tr.push([o.gl || '(belum dipetakan)', o.nama, o.d, o.k, o.d - o.k]); });
    var TD = sumBy(Object.keys(tb), function (k) { return tb[k].d; }), TK = sumBy(Object.keys(tb), function (k) { return tb[k].k; });
    tr.push(['TOTAL', Math.abs(TD - TK) < 1 ? '✓ seimbang' : '⚠ TIDAK SEIMBANG', TD, TK, TD - TK]);
    var rows = [['RINCIAN JURNAL'], ['NO JURNAL', 'TANGGAL', 'KEJADIAN', 'REFERENSI', 'T-CODE', 'DOC TYPE', 'PK', 'AKUN GL', 'NAMA AKUN', 'KODE BP', 'NPP', 'DEBIT', 'KREDIT', 'TEKS', 'NO. DOK SAP', 'HEADER (NO. FAKTUR)']], styles = [sty(0, 'title'), sty(1, 'hdr')];
    J.forEach(function (j, ix) {
      var no = 'J' + pad(ix + 1, 3), d = 0, k = 0;
      j.lines.forEach(function (l, li) { rows.push([li === 0 ? no : '', li === 0 ? j.tgl : '', li === 0 ? j.kejadian : '', li === 0 ? j.ref : '', j.tcode, j.doctype, l.pk, l.gl || '(belum dipetakan)', l.nama, l.cust ? l.bp : '', l.npp || '', l.debit || '', l.kredit || '', l.teks || '', li === 0 ? (j.sap || '') : '', li === 0 ? (j.header || '') : '']); d += l.debit; k += l.kredit; if (li === 0) styles.push(sty(rows.length - 1, 'band')); });
      if (Math.abs(d - k) > 1) ctx.issues.push(['JURNAL', 'Usulan Jurnal SAP', no + ' ' + j.ref, 'Jurnal tidak seimbang (D ' + fmtRp(d) + ' / K ' + fmtRp(k) + ')', 'Periksa data invoice / pembayaran']);
    });
    if (!J.length) rows.push(['', '', 'Tidak ada transaksi di bulan ini']);
    var r2 = 8 + head.length + 1, r3 = r2 + tr.length + 2;
    return { from: 8, blocks: [block(8, 1, head, { styles: [sty(0, 'note')] }), block(r2, 1, tr, { styles: [sty(0, 'title'), sty(1, 'hdr'), sty(tr.length - 1, 'total')], fmt: { 2: 'money', 3: 'money', 4: 'money' }, span: 2 }),
      block(r3, 1, rows, { styles: styles, fmt: { 1: 'date', 11: 'money', 12: 'money' } })] };
  }

  // ------------------------------------------------------------------ REKON SAP (vs LAMPIRAN LK)
  /** pasangkan dokumen LK dengan invoice: (1) no. dok SAP + nilai, (2) NPP + nilai (tanggal terdekat), (3) no. dok saja, (4) NPP + tanggal sama */
  function pairLK(items, invs, valFn) {
    var used = {}, pairOf = {}, ix = function (i) { return normInv(i.nomor); };
    function set(i, it, how) { used[ix(i)] = 1; it._p = 1; pairOf[ix(i)] = { it: it, how: how }; }
    items.forEach(function (it) { var c = invs.filter(function (i) { return !used[ix(i)] && i.sap_doc && i.sap_doc === it.doc && Math.abs(valFn(i) - it.jumlah) <= 1; }); if (c.length) set(c[0], it, 'doc'); });
    items.filter(function (it) { return !it._p; }).sort(function (a, b) { return (a.tgl || '') < (b.tgl || '') ? -1 : 1; }).forEach(function (it) {
      var c = invs.filter(function (i) { return !used[ix(i)] && Math.abs(valFn(i) - it.jumlah) <= 1 && (it.npp ? nppKey(i.npp) === it.npp : i.tgl === it.tgl); });
      c.sort(function (x, y) { return Math.abs(diffDays(x.tgl, it.tgl || x.tgl)) - Math.abs(diffDays(y.tgl, it.tgl || y.tgl)); });
      if (c.length) set(c[0], it, 'nilai');
    });
    items.filter(function (it) { return !it._p; }).forEach(function (it) { var c = invs.filter(function (i) { return !used[ix(i)] && i.sap_doc && i.sap_doc === it.doc && (!it.npp || nppKey(i.npp) === it.npp); }); if (c.length) set(c[0], it, 'beda'); });
    items.filter(function (it) { return !it._p; }).forEach(function (it) { var c = invs.filter(function (i) { return !used[ix(i)] && it.npp && nppKey(i.npp) === it.npp && i.tgl === it.tgl; }); if (c.length === 1) set(c[0], it, 'beda'); });
    var rest = items.filter(function (it) { return !it._p; }); items.forEach(function (it) { delete it._p; });
    return { pairOf: pairOf, rest: rest };
  }
  function rekon(ctx) {
    if (ctx._rekon !== undefined) return ctx._rekon;
    var lk = ctx.lk, C = lk.periode; if (!C) return (ctx._rekon = null);
    var out = { C: C, pu: [], ppn: [], ret: [], suggest: [] }, TOL = N(ctx.params.TOL_PEMBULATAN);
    var invs = ctx.invs.filter(function (i) { return !i._batal && i.tgl && i.tgl <= C; }).sort(function (a, b) { return a.tgl < b.tgl ? -1 : 1; });
    var vPU = function (i) { return i._a.jenis === 'UANG MUKA' ? i._a.dpp : i._a.netto; }, vPP = function (i) { return i._a.ppn; };
    var mPU = pairLK(lk.pu, invs, vPU), mPP = pairLK(lk.ppn, invs.filter(function (i) { return i._a.ppn > 0; }), vPP);
    invs.forEach(function (i) {
      var st = invAt(i, C, ctx.params), k = normInv(i.nomor), m = mPU.pairOf[k], mp = mPP.pairOf[k];
      if (m && i.sap_doc !== m.it.doc && m.how === 'nilai') out.suggest.push({ nomor: i.nomor, doc: m.it.doc, npp: i.npp, tgl: i.tgl, lama: i.sap_doc });
      if (m || st.saldoN) out.pu.push({ inv: i, lk: m ? m.it : null, sis: st.saldoN, lks: m ? m.it.saldo : 0, val: vPU(i) });
      if (mp || st.saldoP) out.ppn.push({ inv: i, lk: mp ? mp.it : null, sis: st.saldoP, lks: mp ? mp.it.saldo : 0, val: vPP(i) });
    });
    mPU.rest.filter(function (x) { return x.saldo; }).forEach(function (x) { out.pu.push({ inv: null, lk: x, sis: 0, lks: x.saldo }); });
    mPP.rest.filter(function (x) { return x.saldo; }).forEach(function (x) { out.ppn.push({ inv: null, lk: x, sis: 0, lks: x.saldo }); });
    var rs = {}; retPieces(ctx, C).forEach(function (pc) { var k = nppKey(pc.inv.npp); rs[k] = (rs[k] || 0) + pc.saldo; });
    var rl = {}; lk.ret.forEach(function (x) { rl[x.npp] = (rl[x.npp] || 0) + x.saldo; });
    Object.keys(rs).concat(Object.keys(rl)).filter(function (k, ix, a) { return a.indexOf(k) === ix; }).sort().forEach(function (k) { out.ret.push({ npp: k, sis: rnd(rs[k] || 0), lks: rnd(rl[k] || 0) }); });
    [out.pu, out.ppn].forEach(function (L) { L.forEach(function (r) {
      r.sel = r.sis - r.lks; r.ok = Math.abs(r.sel) <= TOL;
      var nb = r.lk && r.inv && Math.abs(r.lk.jumlah - r.val) > 1, npb = r.lk && r.inv && r.lk.npp && r.lk.npp !== nppKey(r.inv.npp);
      r.ket = !r.inv ? 'Ada di LK, tidak ada di sistem' : !r.lk ? (r.sis ? 'Ada di sistem, belum ada di LK' : '') :
        [nb ? 'Nilai terbit beda (sistem ' + fmtRp(r.val) + ' / LK ' + fmtRp(r.lk.jumlah) + ')' : '', npb ? 'NPP di SAP berbeda (' + r.lk.npp + ')' : '',
          r.ok ? '' : (r.inv._closedManual && r.lks ? 'Ditutup manual di AR (disetujui), masih terbuka di SAP — minta Akuntansi menutup' : 'Selisih saldo')].filter(String).join('; ') || '✓ cocok';
      if (npb || nb) r.ok = false;
    }); });
    out.ret.forEach(function (r) { r.sel = r.sis - r.lks; r.ok = Math.abs(r.sel) <= TOL; });
    return (ctx._rekon = out);
  }
  /** catatan rekon PPh: pelanggan memotong PPh dengan dasar berbeda (invoice sudah close) */
  function pphNotes(ctx) {
    var L = ctx.invs.filter(function (i) { return !i._batal && i._res.sel_pph; }).sort(function (a, b) { return a.tgl < b.tgl ? -1 : 1; });
    var rows = [['CATATAN REKON PPh — pelanggan memotong PPh dengan dasar berbeda (invoice sudah CLOSE; cocokkan dengan bukti potong saat rekon PPh)'],
      ['NO', 'PELANGGAN', 'NOMOR INVOICE', 'TGL INVOICE', 'NPP', 'PPh SESUAI INVOICE', 'SELISIH (DIPOTONG LEBIH)', 'TOTAL DIPOTONG PELANGGAN', 'TGL BAYAR']];
    L.forEach(function (i, ix) { rows.push([ix + 1, i.pelanggan, i.nomor, i.tgl, i.npp, i._res.pot_pph, i._res.sel_pph, i._res.pot_pph + i._res.sel_pph, i._res.tgl_bayar]); });
    if (!L.length) rows.push(['', 'Tidak ada']);
    rows.push(['', 'TOTAL', '', '', '', sumBy(L, function (i) { return i._res.pot_pph; }), sumBy(L, function (i) { return i._res.sel_pph; }), sumBy(L, function (i) { return i._res.pot_pph + i._res.sel_pph; })]);
    return rows;
  }
  function repRekon(ctx) {
    var R = rekon(ctx), pn = pphNotes(ctx), pnB = function (r) { return block(r, 1, pn, { styles: [sty(0, 'title'), sty(1, 'hdr'), sty(pn.length - 1, 'total')], fmt: { 3: 'date', 5: 'money', 6: 'money', 7: 'money', 8: 'date' } }); };
    if (!R) return { from: 8, blocks: [block(8, 1, [['Belum ada data LK. Tempel lampiran LK di sheet "Data LK (Tempel)" dan isi PERIODE LK, lalu perbarui laporan.']], { styles: [sty(0, 'note')] }), pnB(11)] };
    ctx.rekonSuggest = R.suggest;
    function sumr(L, f) { return sumBy(L, f); }
    var sm = [['REKON SALDO SISTEM AR vs LAMPIRAN LK PER ' + tglPanjang(R.C).toUpperCase()], ['AKUN', 'SALDO SISTEM AR', 'SALDO LK (SAP)', 'SELISIH', 'JML DOKUMEN', '✓ COCOK', 'SELISIH NILAI', 'HANYA DI SISTEM', 'HANYA DI LK']];
    [['Piutang Usaha (1104x1)', R.pu], ['Piutang PPN K (1104x2)', R.ppn]].forEach(function (x) {
      var L = x[1]; sm.push([x[0], sumr(L, function (r) { return r.sis; }), sumr(L, function (r) { return r.lks; }), sumr(L, function (r) { return r.sel; }), L.length, L.filter(function (r) { return r.ok && r.inv && r.lk; }).length,
        L.filter(function (r) { return !r.ok && r.inv && r.lk; }).length, L.filter(function (r) { return !r.lk && r.sis; }).length, L.filter(function (r) { return !r.inv; }).length]);
    });
    sm.push(['Piutang Retensi (1105x1, per NPP)', sumr(R.ret, function (r) { return r.sis; }), sumr(R.ret, function (r) { return r.lks; }), sumr(R.ret, function (r) { return r.sel; }), R.ret.length, R.ret.filter(function (r) { return r.ok; }).length, R.ret.filter(function (r) { return !r.ok; }).length, '', '']);
    var blocks = [block(8, 1, sm, { styles: [sty(0, 'title'), sty(1, 'hdr')], fmt: { 1: 'money', 2: 'money', 3: 'money', 4: 'int', 5: 'int', 6: 'int', 7: 'int', 8: 'int' }, span: 2 })], r = 8 + sm.length + 2;
    function det(title, L) {
      var rows = [[title], ['NO', 'PELANGGAN', 'NPP', 'NOMOR INVOICE', 'TGL', 'NO. DOK SAP (SISTEM)', 'NO. DOK SAP (LK)', 'SALDO SISTEM', 'SALDO LK', 'SELISIH', 'KETERANGAN', 'STATUS DI SISTEM']], n = 0;
      L.filter(function (x) { return !x.ok || !x.inv || !x.lk; }).forEach(function (x) {
        var i = x.inv, l = x.lk || {};
        rows.push([++n, i ? i.pelanggan : l.pelanggan, i ? i.npp : l.npp, i ? i.nomor : '', i ? i.tgl : l.tgl, i ? i.sap_doc : '', l.doc || '', x.sis, x.lks, x.sel, x.ket, i ? i._res.status : '']);
      });
      if (!n) rows.push(['', '✓ Semua dokumen cocok']);
      blocks.push(block(r, 1, rows, { styles: [sty(0, 'title'), sty(1, 'hdr')], fmt: { 4: 'date', 7: 'money', 8: 'money', 9: 'money' } })); r += rows.length + 2;
    }
    det('PIUTANG USAHA — DOKUMEN YANG PERLU DICEK', R.pu);
    det('PIUTANG PPN K — DOKUMEN YANG PERLU DICEK', R.ppn);
    var rr = [['PIUTANG RETENSI PER NPP'], ['NPP', 'SALDO SISTEM', 'SALDO LK', 'SELISIH', 'KETERANGAN']];
    R.ret.forEach(function (x) { rr.push([x.npp, x.sis, x.lks, x.sel, x.ok ? '✓ cocok' : (!x.lks ? 'Tidak ada di LK' : !x.sis ? 'Tidak ada di sistem — lengkapi RETENSI invoice lama' : 'Selisih')]); });
    blocks.push(block(r, 1, rr, { styles: [sty(0, 'title'), sty(1, 'hdr')], fmt: { 1: 'money', 2: 'money', 3: 'money' }, span: 2 })); r += rr.length + 2;
    var sg = [['SARAN NO. DOK SAP UNTUK INVOICE YANG BELUM DIISI (cocok nilai & NPP) — isi otomatis lewat menu AR CSU ▸ Isi No. Dok SAP dari Data LK'], ['NOMOR INVOICE', 'NPP', 'TGL', 'NO. DOK SAP (LK)']];
    R.suggest.forEach(function (s) { sg.push([s.nomor, s.npp, s.tgl, s.doc]); });
    if (!R.suggest.length) sg.push(['—']);
    blocks.push(block(r, 1, sg, { styles: [sty(0, 'title'), sty(1, 'hdr')], fmt: { 2: 'date' }, span: 2 })); r += sg.length + 2;
    blocks.push(pnB(r));
    var nSel = R.pu.filter(function (x) { return !x.ok; }).length + R.ppn.filter(function (x) { return !x.ok; }).length;
    if (nSel) ctx.issues.push(['REKON', 'Rekon SAP', 'per ' + R.C, nSel + ' dokumen piutang tidak cocok dengan lampiran LK', 'Lihat sheet Rekon SAP; konfirmasi ke Akuntansi']);
    return { from: 8, blocks: blocks };
  }

  // ------------------------------------------------------------------ CEK DATA
  function repCek(ctx) {
    var rows = [['NO', 'KATEGORI', 'SHEET', 'REFERENSI', 'MASALAH', 'SARAN PERBAIKAN']];
    ctx.issues.forEach(function (x, ix) { rows.push([ix + 1, x[0], x[1], x[2], x[3], x[4]]); });
    if (rows.length === 1) rows.push(['', '✓', '', '', 'Tidak ada data yang perlu dicek', '']);
    return { from: 6, blocks: [block(6, 1, rows, { styles: [sty(0, 'hdr')] })], wrapCols: [5, 6] };
  }

  // ================================================================== v2.4: MEMORIAL AKUNTANSI, RINGKASAN P&E, TEMPEL SAP, SELARAS P&E
  var BLN3 = ['JAN', 'FEB', 'MAR', 'APR', 'MEI', 'JUN', 'JUL', 'AGU', 'SEP', 'OKT', 'NOV', 'DES'];
  function tglMemo(s) { if (!s) return ''; var p = s.split('-'); return p[2] + ' ' + BLN3[+p[1] - 1] + ' ' + p[0]; }
  /** 04002600367072394 → 0400 2600 36707 2394 (format di dokumen memorial) */
  function fpFmt(v) { var d = digits(v); return d.length === 17 ? d.substr(0, 4) + ' ' + d.substr(4, 4) + ' ' + d.substr(8, 5) + ' ' + d.substr(13, 4) : d; }
  function memoCode(ctx, kode) { var a = ctx.params.akun[kode]; return (a && a.memo) || MEMO_KODE[kode] || ''; }
  function memoLabel(ctx, kode) { if (/^BANK_/.test(kode)) return 'BANK ' + kode.slice(5); return MEMO_LABEL[kode] || U(glName(ctx.params, kode)); }
  function lblMemo(i) { return U(invLabel(i)).replace(/\bKE-/g, 'ke-'); }
  function invShort(nomor) { return S(nomor).replace(/^KU\.02\.02\/WB-/, ''); }
  function produkOf(inv) { var p = inv._pro || {}; return U(p.nama_pe || p.pekerjaan || p.singkat || inv.npp).replace(/\s+/g, ' '); }
  function yymm(tgl) { return S(tgl).substring(2, 4) + S(tgl).substring(5, 7); }
  function memoKodeInv(inv) { var m = normInv(inv.nomor).match(/4C\.(\d{4}[A-Z]?)\/\d{4}(#(\d+))?/); return 'MI-' + yymm(inv.tgl) + '-' + (m ? m[1] + (m[3] ? '-' + m[3] : '') : hash(inv.nomor).slice(-6)); }
  function memoRowsFromLines(ctx, lines, base, pajakOf) {
    return lines.map(function (l) {
      var code = memoCode(ctx, l.kode), lbl = memoLabel(ctx, l.kode);
      return { akun: code || 'MOHON KONFIRMASI', uraian: (lbl + ' ' + base).replace(/\s+/g, ' ').trim(), debet: l.debit || 0, kredit: l.kredit || 0, pajak: pajakOf(l.kode) || '',
        catatan: code ? '' : 'Kode akun memorial ' + l.kode + ' belum diisi (Parameter ▸ Pemetaan Akun SAP kolom T)' };
    });
  }
  function memoFp(rows) { return hash('MF|' + rows.map(function (r) { return [r.akun, rnd(N(r.debet)), rnd(N(r.kredit))].join(','); }).join(';')); }
  function memoInfo(rows) { return hash('MI|' + rows.map(function (r) { return [r.uraian, r.pajak, r.npp, r.pelanggan, r.produk, r.inv].join(','); }).join(';')); }
  /** blok memorial hasil hitung (invoice keluar & penerimaan sejak MEMORIAL_MULAI) */
  function memoGenerate(ctx) {
    var mulai = D(ctx.params.MEMORIAL_MULAI) || '2026-10-01', out = [];
    // tanggal di masa depan (salah ketik) tidak dibuatkan memorial sampai diperbaiki
    ctx.invs.filter(function (i) { return !i._batal && i.tgl && i.tgl >= mulai && i.tgl <= ctx.today; }).forEach(function (i) {
      var base = lblMemo(i) + ' ' + produkOf(i) + ' ' + invShort(i.nomor) + ', ' + tglMemo(i.tgl);
      var rows = memoRowsFromLines(ctx, jInvoice(ctx, i).lines, base, function (k) { return k === 'PIUTANG_PPN' || k === 'PPN_KELUARAN' ? fpFmt(i.no_fp) : ''; });
      rows.forEach(function (r) { r.npp = i.npp; r.pelanggan = i.pelanggan; r.produk = produkOf(i); r.inv = i.nomor; });
      out.push({ ref: 'INV|' + normInv(i.nomor), jenis: 'INV', kode: memoKodeInv(i), bulan: i.tgl.substring(0, 7), tgl: i.tgl, rows: rows, sidik: memoFp(rows), info: memoInfo(rows) });
    });
    var pays = ctx.pays.filter(function (p) { return p._inv && !p._inv._batal && p.tgl && p.tgl >= mulai && p.tgl <= ctx.today && (p.nominal > 0 || p._pb); })
      .sort(function (a, b) { return a.tgl < b.tgl ? -1 : a.tgl > b.tgl ? 1 : a._row - b._row; });
    var pbDone = {};
    pays.forEach(function (p) {
      var L = p._pb ? pays.filter(function (q) { return q._pb && S(q.ref) === S(p.ref); }) : [p];
      if (p._pb) { if (pbDone[S(p.ref)]) return; pbDone[S(p.ref)] = 1; }
      var rows = [];
      L.forEach(function (q) {
        var i = q._inv, base = (q._pb ? 'PINDAH BUKU ' : 'PENERIMAAN ') + lblMemo(i) + ' ' + produkOf(i) + ' ' + invShort(i.nomor) + ', ' + tglMemo(q.tgl);
        var rr = memoRowsFromLines(ctx, jPayment(ctx, q).lines, base, function (k) { return k === 'PIUTANG_PPN' ? fpFmt(i.no_fp) : (k === 'PPH_FINAL' || k === 'PPH_23') && S(q.bupot) ? 'BUPOT ' + S(q.bupot) : ''; });
        rr.forEach(function (r) { r.npp = i.npp; r.pelanggan = i.pelanggan; r.produk = produkOf(i); r.inv = i.nomor; });
        rows = rows.concat(rr);
      });
      var deb = rows.filter(function (r) { return r.debet; }), kre = rows.filter(function (r) { return !r.debet; });
      rows = deb.concat(kre);
      var kode = 'MT-' + yymm(p.tgl) + '-' + (p._pb ? S(p.ref).toUpperCase().replace(/[^A-Z0-9\-]/g, '') : (S(p.idp).match(/(\d+)$/) || ['', p.idp])[1]);
      out.push({ ref: p._pb ? 'PB|' + S(p.ref) : 'TRM|' + p.idp, jenis: 'TRM', kode: kode, bulan: p.tgl.substring(0, 7), tgl: p.tgl, rows: rows, sidik: memoFp(rows), info: memoInfo(rows) });
    });
    return out;
  }
  /** baca blok memorial tersimpan (sheet Memorial AR) */
  function memoParse(rows) {
    var blocks = [], cur = null;
    rowsToObjects(rows || [], MEMO.cols, MEMO.firstRow).forEach(function (o) {
      if (o._empty) return;
      var kode = S(o.kode);
      if (!kode && !cur) return;
      if (kode && (!cur || cur.kode !== kode)) {
        cur = { kode: kode, ref: S(o.ref), jenis: U(o.jenis) || (/^MT/.test(kode) ? 'TRM' : 'INV'), bulan: S(o.bulan).substring(0, 7), sumber: U(o.sumber) || 'OTOMATIS', sidik: S(o.sidik),
          status: S(o.status), verif: /^(✓|TRUE|YA|1)/i.test(S(o.verif)), verif_at: S(o.verif_at), rows: [], info: '' };
        blocks.push(cur);
      } else if (!kode) cur.sumber = 'MANUAL';
      cur.rows.push({ akun: S(o.akun), uraian: S(o.uraian), npp: S(o.npp), pelanggan: S(o.pelanggan), produk: S(o.produk), pajak: S(o.pajak), inv: S(o.inv),
        debet: N(o.debet) || 0, kredit: N(o.kredit) || 0, catatan: S(o.catatan) });
    });
    blocks.forEach(function (b) { b.info = memoInfo(b.rows); if (!b.ref) b.ref = 'MANUAL|' + b.kode; });
    return blocks;
  }
  function memoBase(b) { return /-K\d+$/.test(b.kode) ? 'KOREKSI' : 'BARU'; }
  function memoEffect(b) { return b.rows.filter(function (r) { return U(r.catatan) !== 'PEMBALIK'; }); }
  function memoReverse(rows) { return rows.map(function (r) { var x = JSON.parse(JSON.stringify(r)); var d = x.debet; x.debet = x.kredit; x.kredit = d; x.uraian = 'PEMBALIK ' + x.uraian; x.catatan = 'PEMBALIK'; return x; }); }
  /** gabungkan blok tersimpan dengan hasil hitung: blok baru, revisi (belum diverifikasi), koreksi (sudah diverifikasi), batal */
  function memoMerge(ctx, stored, gen) {
    var today = ctx.today, tg = today.split('-').reverse().join('-'), byRef = {}, order = [], G = {}, log = [];
    stored.forEach(function (b) { if (!byRef[b.ref]) { byRef[b.ref] = []; order.push(b.ref); } byRef[b.ref].push(b); });
    gen.forEach(function (g) { G[g.ref] = g; if (!byRef[g.ref]) { byRef[g.ref] = []; order.push(g.ref); } });
    var final = [];
    order.forEach(function (ref) {
      var L = byRef[ref], g = G[ref] || null;
      if (/^MANUAL\|/.test(ref)) { L.forEach(function (b) { final.push(b); }); return; }
      var live = L.filter(function (b) { return !/DIKOREKSI|^BATAL/.test(U(b.status)); }), eff = live.length ? live[live.length - 1] : null;
      var kN = L.filter(function (b) { return /-K\d+$/.test(b.kode); }).length;
      if (!L.length) { if (g) { L.push({ kode: g.kode, ref: ref, jenis: g.jenis, bulan: g.bulan, sumber: 'OTOMATIS', sidik: g.sidik, status: 'BARU', verif: false, verif_at: '', rows: g.rows, info: g.info, _new: 1 }); log.push('baru ' + g.kode); } }
      else if (!eff) { if (g && L.every(function (b) { return /^BATAL/.test(U(b.status)); })) { var b0 = L[L.length - 1]; b0.status = 'BARU'; b0.rows = g.rows; b0.sidik = g.sidik; b0.info = g.info; log.push('aktif lagi ' + b0.kode); } }
      else if (!g) {
        if (!eff.verif) { if (!/^BATAL/.test(U(eff.status))) { eff.status = 'BATAL ' + tg; log.push('batal ' + eff.kode); } }
        else if (!/PEMBATALAN/.test(U(eff.status))) {
          var kb = eff.kode.replace(/-K\d+$/, '') + '-K' + (kN + 1);
          L.push({ kode: kb, ref: ref, jenis: eff.jenis, bulan: today.substring(0, 7), sumber: 'OTOMATIS', sidik: 'BATAL', status: 'KOREKSI (PEMBATALAN)', verif: false, verif_at: '', rows: memoReverse(memoEffect(eff)), _new: 1 });
          eff.status = 'DIKOREKSI → ' + kb; log.push('pembatalan ' + kb);
        }
      } else if (g.sidik === eff.sidik) {
        if (!eff.verif && eff.sumber === 'OTOMATIS' && !/-K\d+$/.test(eff.kode) && g.info !== eff.info) { eff.rows = g.rows; eff.info = g.info; log.push('info ' + eff.kode); }
      } else if (!eff.verif) {
        if (eff.sumber === 'MANUAL') { if (!/BERUBAH/.test(eff.status)) { eff.status = '⚠ INVOICE BERUBAH — CEK MANUAL'; log.push('cek manual ' + eff.kode); } }
        else if (/-K\d+$/.test(eff.kode)) {
          var prev = L.filter(function (b) { return /DIKOREKSI/.test(U(b.status)); }).pop();
          eff.rows = (prev ? memoReverse(memoEffect(prev)) : []).concat(g.rows); eff.sidik = g.sidik; eff.info = g.info; eff.status = 'KOREKSI · DIREVISI ' + tg; log.push('revisi koreksi ' + eff.kode);
        } else { eff.rows = g.rows; eff.sidik = g.sidik; eff.info = g.info; eff.status = 'DIREVISI ' + tg; log.push('revisi ' + eff.kode); }
      } else {
        var kk = eff.kode.replace(/-K\d+$/, '') + '-K' + (kN + 1);
        L.push({ kode: kk, ref: ref, jenis: eff.jenis, bulan: today.substring(0, 7), sumber: 'OTOMATIS', sidik: g.sidik, status: 'KOREKSI', verif: false, verif_at: '', rows: memoReverse(memoEffect(eff)).concat(g.rows), _new: 1 });
        eff.status = 'DIKOREKSI → ' + kk; log.push('koreksi ' + kk);
      }
      L.forEach(function (b) { b.tgl = g ? g.tgl : (b.tgl || ''); final.push(b); });
    });
    final.sort(function (a, b) { return a.bulan < b.bulan ? -1 : a.bulan > b.bulan ? 1 : a.jenis < b.jenis ? -1 : a.jenis > b.jenis ? 1 : a.kode < b.kode ? -1 : a.kode > b.kode ? 1 : 0; });
    return { blocks: final, log: log };
  }
  /** baris sheet Memorial AR (urutan kolom MEMO) */
  function memoSheetRows(blocks) {
    var out = [], no = {}, keyOf = function (b) { return b.bulan + '|' + b.jenis; };
    blocks.forEach(function (b) {
      var k = keyOf(b); no[k] = (no[k] || 0) + 1;
      var td = sumBy(b.rows, function (r) { return N(r.debet); }), tk = sumBy(b.rows, function (r) { return N(r.kredit); });
      var base = S(b.status).replace(/^✓ DIVERIFIKASI( · )?/, '');
      var st = b.verif ? '✓ DIVERIFIKASI' + (/DIKOREKSI/.test(U(base)) ? ' · ' + base : '') : (base || memoBase(b));
      b.rows.forEach(function (r, ix) {
        var last = ix === b.rows.length - 1;
        out.push([ix === 0 ? no[k] : '', r.akun, r.uraian, r.npp, r.pelanggan, r.produk, r.pajak, r.inv, r.debet || '', r.kredit || '', last ? td : '', last ? tk : '', last ? td - tk : '',
          ix === 0 ? st : '', ix === 0 ? (b.verif ? '✓' : '') : '', b.kode, b.bulan, b.jenis, b.sumber, ix === 0 ? b.sidik : '', ix === 0 ? b.verif_at : '', r.catatan || '', b.ref]);
      });
    });
    return out;
  }
  function memoStatusText(b) {
    if (!b) return '';
    if (b.verif) return '✓ DIVERIFIKASI ' + S(b.verif_at) + ' · ' + b.kode;
    return (S(b.status).replace(/^✓ DIVERIFIKASI( · )?/, '') || memoBase(b)) + ' — belum diverifikasi · ' + b.kode;
  }
  /** status memorial per invoice / penerimaan (blok terakhir yang berlaku) */
  function memoStatusMap(blocks) {
    var m = {};
    blocks.forEach(function (b) { if (!/DIKOREKSI/.test(U(b.status)) || !m[b.ref]) m[b.ref] = b; });
    var out = {}; Object.keys(m).forEach(function (k) { out[k] = memoStatusText(m[k]); });
    return out;
  }
  function memorial(ctx, storedRows) {
    var stored = memoParse(storedRows), gen = memoGenerate(ctx), mg = memoMerge(ctx, stored, gen);
    var rows = memoSheetRows(mg.blocks), norm = function (L) { return (L || []).filter(function (r) { return r.some(function (v) { return S(v) !== ''; }); }).map(function (r) { return r.map(function (v) { return typeof v === 'number' ? String(rnd(v)) : S(v); }).join('\u0001'); }).join('\u0002'); };
    return { blocks: mg.blocks, rows: rows, changed: norm(rows) !== norm(storedRows), log: mg.log, status: memoStatusMap(mg.blocks),
      belum: mg.blocks.filter(function (b) { return !b.verif && !/DIKOREKSI|^BATAL/.test(U(b.status)); }) };
  }

  // ------------------------------------------------------------------ RINGKASAN UNTUK P&E
  function proyekAR(ctx) {
    return Object.keys(ctx.M.pro).filter(function (n) { var p = ctx.M.pro[n]; return !p._produkOf && !p._adendumOf; });
  }
  function peSummary(ctx) {
    var umr = {}, ret = {}, today = ctx.today, out = [];
    umRetRows(ctx).forEach(function (u) { umr[u.npp] = u; });
    retensiRows(ctx).forEach(function (r) { ret[r[1]] = r[9]; });
    proyekAR(ctx).forEach(function (n) {
      var p = ctx.M.pro[n], k = kartuData(ctx, n, today); if (!k) return;
      if (!k.groups.length && (p.aktif !== 'YA' || p.selesai)) return;
      var T = k.T, b = { belum: 0, d30: 0, d60: 0, d90: 0, lebih: 0, ba: 0 }, nOd = 0, nBjt = 0, lastInv = '', lastPay = '';
      k.groups.forEach(function (g) {
        var i = g.inv, sal = g.saldo; if (i.tgl > lastInv) lastInv = i.tgl;
        i._pays.forEach(function (q) { if (!q._pb && q.tgl && q.tgl <= today && q.tgl > lastPay) lastPay = q.tgl; });
        if (!(sal > 1)) return;
        if (/BA\/MEMO/.test(i._res.status)) { b.ba += sal; return; }
        var jt = i._res.jt, d = jt ? diffDays(today, jt) : -1;
        if (d <= 0) { b.belum += sal; nBjt++; } else { nOd++; if (d <= 30) b.d30 += sal; else if (d <= 60) b.d60 += sal; else if (d <= 90) b.d90 += sal; else b.lebih += sal; }
      });
      var u = umr[n], st = [];
      if (Math.abs(T.saldo) <= 1) st.push(k.groups.length ? 'Lunas semua' : 'Belum ada tagihan');
      if (nOd) st.push(nOd + ' inv lewat jatuh tempo'); if (nBjt) st.push(nBjt + ' inv belum jatuh tempo'); if (b.ba > 1) st.push('kekurangan via BA/memo');
      out.push({ npp: n, produk: nppListTxt(p.produk), adendum: nppListTxt(p.adendumNpp), pelanggan: p.pelanggan, proyek: p.nama_pe || S(p.pekerjaan) || p.singkat, singkat: p.singkat,
        kontrak: k.kontrak.net, progres: T.brutoProg, pct: k.pctTagih, umk: T.umk, ret: sumBy(k.groups, function (g) { return g.A.jenis === 'RETENSI' ? 0 : g.A.ret; }), dpp: T.dpp,
        kas: T.kas, sisa: T.saldo, aging: b, sisaUM: u && u.basis ? (u.lengkap ? u.sisa : '') : '', retBelum: ret[n] || 0, lastInv: lastInv, lastPay: lastPay,
        status: st.join(' · '), selesai: !!p.selesai, aktif: p.aktif });
    });
    out.sort(function (a, b) { return (a.selesai ? 1 : 0) - (b.selesai ? 1 : 0) || (a.pelanggan < b.pelanggan ? -1 : a.pelanggan > b.pelanggan ? 1 : a.npp < b.npp ? -1 : 1); });
    return out;
  }
  function peSummaryBlocks(ctx, top) {
    var L = peSummary(ctx), act = L.filter(function (x) { return !x.selesai; }), done = L.filter(function (x) { return x.selesai; });
    var hdr = ['NO', 'NPP (AR)', 'NPP PRODUK / ADENDUM', 'PELANGGAN', 'NAMA PROYEK (P&E)', 'NILAI KONTRAK (EXCL. PPN)', 'NILAI PROGRES DITAGIH', '% PROGRES DITAGIH', 'POT. UANG MUKA', 'RETENSI',
      'DPP DITAGIH', 'DITERIMA (KAS)', 'SISA PIUTANG', 'BELUM JATUH TEMPO', 'LEWAT 1–30 HARI', 'LEWAT 31–60 HARI', 'LEWAT 61–90 HARI', 'LEWAT > 90 HARI', 'KEKURANGAN VIA BA / MEMO',
      'SISA UANG MUKA', 'RETENSI BELUM DITAGIH', 'TAGIHAN TERAKHIR', 'PENERIMAAN TERAKHIR', 'STATUS'];
    var row = function (x, ix) { var a = x.aging; return [ix + 1, x.npp, [x.produk, x.adendum ? 'adendum: ' + x.adendum : ''].filter(String).join(' · '), x.pelanggan, x.proyek, x.kontrak, x.progres, x.pct, x.umk, x.ret, x.dpp, x.kas, x.sisa,
      a.belum, a.d30, a.d60, a.d90, a.lebih, a.ba, x.sisaUM, x.retBelum, x.lastInv, x.lastPay, x.status]; };
    var tot = function (L2, lbl) { var f = function (k) { return sumBy(L2, k); }; return ['', lbl, '', '', '', f(function (x) { return x.kontrak; }), f(function (x) { return x.progres; }), '', f(function (x) { return x.umk; }), f(function (x) { return x.ret; }),
      f(function (x) { return x.dpp; }), f(function (x) { return x.kas; }), f(function (x) { return x.sisa; }), f(function (x) { return x.aging.belum; }), f(function (x) { return x.aging.d30; }), f(function (x) { return x.aging.d60; }),
      f(function (x) { return x.aging.d90; }), f(function (x) { return x.aging.lebih; }), f(function (x) { return x.aging.ba; }), f(function (x) { return N(x.sisaUM); }), f(function (x) { return x.retBelum; }), '', '', '']; };
    var fmt = { 5: 'money', 6: 'money', 7: 'pct', 8: 'money', 9: 'money', 10: 'money', 11: 'money', 12: 'money0', 13: 'money', 14: 'money', 15: 'money', 16: 'money', 17: 'money', 18: 'money', 19: 'money', 20: 'money', 21: 'date', 22: 'date' };
    var A = [['RINGKASAN PIUTANG PROYEK — CONCRETE SERVICES UNIT · per ' + tglPanjang(ctx.today)], hdr], sa = [sty(0, 'title'), sty(1, 'hdr')];
    act.forEach(function (x, ix) { A.push(row(x, ix)); });
    A.push(tot(act, 'TOTAL PROYEK AKTIF')); sa.push(sty(A.length - 1, 'total'));
    A.push(['Sisa piutang = DPP + PPN − penerimaan − PPh/adm yang sudah dipotong (sama dengan kartu piutang). Umur dihitung dari tanggal jatuh tempo; tagihan tanpa jatuh tempo (belum submit) masuk "belum jatuh tempo".']); sa.push(sty(A.length - 1, 'note'));
    var blocks = [block(top, 1, A, { styles: sa, fmt: fmt })];
    if (done.length) {
      var B2 = [['PROYEK SELESAI (' + done.length + ')'], hdr], sb = [sty(0, 'title'), sty(1, 'hdr')];
      done.forEach(function (x, ix) { B2.push(row(x, ix)); });
      blocks.push(block(top + A.length + 1, 1, B2, { styles: sb, fmt: fmt, group: true }));
    }
    return { from: top, blocks: blocks, wrapCols: [3, 5, 24], list: L };
  }

  // ------------------------------------------------------------------ TEMPEL SAP (export FBL5N / daftar dokumen) → No. Dok SAP
  /** rows: tabel tempelan (baris pertama yang berisi judul kolom dikenali otomatis) */
  function sapParse(rows) {
    var H = -1, map = {};
    var pats = { doc: /(document ?n|doc\.? ?no|no\.? ?dok|nomor dok|belegn|^dokumen$|document number)/i, ref: /(^reference|referen|xblnr)/i, asg: /(assignment|zuordn|penugasan|alokasi)/i,
      txt: /(^text|teks|keterangan|uraian)/i, amt: /(amount|jumlah|nilai|amt|betrag)/i, tgl: /(posting|pstng|tgl|tanggal|doc\.? ?date|date)/i, typ: /(doc\.? ?type|^type|jenis dok|^ty$|^typ)/i };
    for (var i = 0; i < Math.min((rows || []).length, 30) && H < 0; i++) {
      var m = {}, hit = 0;
      (rows[i] || []).forEach(function (c, j) { var t = S(c); if (!t) return; Object.keys(pats).forEach(function (k) { if (m[k] === undefined && pats[k].test(t)) { m[k] = j; hit++; } }); });
      if (m.doc !== undefined && m.amt !== undefined && hit >= 3) { H = i; map = m; }
    }
    if (H < 0) return { ok: false, rows: [] };
    var out = [];
    rows.slice(H + 1).forEach(function (r) {
      var doc = S(r[map.doc]).replace(/\.0$/, ''); if (!/^\d{6,}$/.test(doc)) return;
      var g = function (k) { return map[k] === undefined ? '' : r[map[k]]; };
      out.push({ doc: doc, ref: S(g('ref')), asg: S(g('asg')), txt: S(g('txt')), amt: Math.abs(rnd(N(g('amt')))), tgl: D(g('tgl')), typ: U(g('typ')) });
    });
    return { ok: true, rows: out, header: H + 1 };
  }
  /** cocokkan dokumen SAP dengan invoice (No. Invoice + nilai) dan penerimaan (No. Invoice + nominal + tanggal ±10 hari) */
  function sapMatch(ctx, sap) {
    var inv = [], pay = [], used = {};
    var hasInv = function (x, nomor) { var k = normInv(nomor), sh = invShort(nomor).replace(/^4C\./, ''); return [x.ref, x.asg, x.txt].some(function (t) { var u = U(t).replace(/\s+/g, ''); return u && (normInv(t) === k || u.indexOf(sh) >= 0); }); };
    ctx.invs.filter(function (i) { return !i._batal; }).forEach(function (i) {
      var a = i._a, vals = [a.netto, a.total, a.dpp, a.netto + a.ppn];
      var c = sap.rows.filter(function (x) { return !used[x.doc] && hasInv(x, i.nomor) && vals.some(function (v) { return Math.abs(v - x.amt) <= 1; }) && !/^(DZ|ZP|KZ)$/.test(x.typ); });
      if (!c.length) return;
      used[c[0].doc] = 1;
      inv.push({ nomor: i.nomor, doc: c[0].doc, lama: i.sap_doc, sama: i.sap_doc === c[0].doc });
    });
    ctx.pays.filter(function (p) { return p._inv && p.nominal > 0 && !p._pb; }).forEach(function (p) {
      var c = sap.rows.filter(function (x) { return !used[x.doc] && Math.abs(x.amt - p.nominal) <= 1 && (hasInv(x, p._inv.nomor) || /^(DZ|ZP)$/.test(x.typ)) && (!x.tgl || !p.tgl || Math.abs(diffDays(x.tgl, p.tgl)) <= 10); });
      c.sort(function (a, b) { return (hasInv(b, p._inv.nomor) ? 1 : 0) - (hasInv(a, p._inv.nomor) ? 1 : 0); });
      if (!c.length || (!hasInv(c[0], p._inv.nomor) && c.length > 1)) return;
      used[c[0].doc] = 1;
      pay.push({ idp: p.idp, row: p._row, nomor: p._inv.nomor, tgl: p.tgl, nominal: p.nominal, doc: c[0].doc, lama: p.sap_doc, sama: p.sap_doc === c[0].doc });
    });
    return { inv: inv, pay: pay, sisa: sap.rows.filter(function (x) { return !used[x.doc]; }) };
  }

  // ------------------------------------------------------------------ SELARAS P&E
  function peCheck(ctx, refRows) {
    var I = ctx.issues, ab = {}, ref = {};
    S(ctx.params.PE_ABAIKAN).split(/[;,\n]+/).forEach(function (x) { if (S(x)) ab[normNpp(x)] = 1; });
    rowsToObjects(refRows || [], PEREF.cols, PEREF.firstRow).forEach(function (r) { if (r._empty || !S(r.npp)) return; var n = normNpp(r.npp); ref[n] = { npp: n, jenis: U(r.jenis) || 'JASA', induk: normNpp(r.induk), nama: S(r.nama), nilai: N(r.nilai) }; });
    if (!Object.keys(ref).length) return;
    var pro = ctx.M.pro, mapped = {};
    Object.keys(pro).forEach(function (n) { pro[n].produk.concat(pro[n].adendumNpp).forEach(function (x) { mapped[x.npp] = n; }); });
    Object.keys(ref).forEach(function (n) {
      var r = ref[n]; if (ab[n]) return;
      if (r.jenis === 'JASA' && !pro[n]) I.push(['SELARAS P&E', PEREF.sheet, n, 'NPP ada di daftar P&E (' + r.nama + ') tetapi belum ada di Master Proyek', 'Daftarkan proyek bila akan ditagih, atau tulis NPP ini di Parameter PE_ABAIKAN']);
      if (r.jenis !== 'JASA' && !mapped[n]) I.push(['SELARAS P&E', MPRO.sheet, r.induk || n, 'NPP ' + r.jenis.toLowerCase() + ' ' + n + ' belum dicantumkan di proyek induknya' + (r.induk ? ' ' + r.induk : ''), 'Isi kolom NPP PRODUK / NPP ADENDUM TERPISAH di Master Proyek']);
    });
    Object.keys(pro).forEach(function (n) {
      var p = pro[n]; if (p._produkOf || p._adendumOf || p.selesai || p.aktif === 'TIDAK' || ab[n]) return;
      if (!ref[n] && !/ 0000 /.test(n)) I.push(['SELARAS P&E', MPRO.sheet, n, 'Proyek AR tidak ada di daftar NPP P&E', 'Konfirmasi ke P&E (mungkin NPP berbeda / belum terdaftar)']);
    });
  }

  // ================================================================== v2.5: VALIDASI RENCANA CAIR PER INVOICE, RA/RI GABUNGAN, KINERJA & PERINGATAN DINI
  /** keputusan Section Head per invoice */
  var DEC = { C: '✓ CAIR', S: '◇ CADANGAN', N: '→ BULAN DEPAN', H: '⏸ HOLD' }, DEC_ORDER = ['C', 'S', 'H', 'N'];
  var VGRP = { A: 'Tunggakan (jatuh tempo sebelum awal bulan)', B: 'Jatuh tempo bulan ini', C: 'Jatuh tempo belum ada (waiting submit / belum tanda terima)', D: 'Jatuh tempo bulan depan dan seterusnya' };
  /** isi kolom USULAN AR → kode (nilai lama "✓ MASUK RENCANA" / "✗ TUNDA" tetap dikenali) */
  function usulCode(v) {
    var u = U(v); if (!u) return '';
    if (/HOLD|TAHAN/.test(u)) return 'H';
    if (/CADANG/.test(u)) return 'S';
    if (/BULAN DEPAN|TUNDA|TIDAK MASUK|GESER/.test(u)) return 'N';
    if (/CAIR|MASUK/.test(u)) return 'C';
    return '';
  }
  /** sidik keputusan: ikut ditandatangani (kolom SIDIK = sidik data / sidik keputusan) */
  function valDecHash(items) { return hash('DEC|' + (items || []).map(function (x) { return normInv(x.nomor) + ':' + S(x.dec) + ':' + S(x.alasan); }).join(';')); }
  function valSidik(fp0, items) { return S(fp0) + '/' + valDecHash(items); }

  /** indeks keputusan yang berlaku per bulan (revisi terakhir yang disetujui) + teks kolom VALIDASI SECTION HEAD */
  function valIndex(appr) {
    var byRef = {}, pend = {}, valMap = {}, first = {};
    appr.forEach(function (r) {
      if (r.jenis !== DOC.VAL || !r._data || !r._data.items) return;
      var v2 = r._data.v === 2;
      if (r.status === 'DISETUJUI') {
        r._decOk = !v2 || S(r.sidik).split('/')[1] === valDecHash(r._data.items);
        if (!r._decOk) return;
        var m = {};
        r._data.items.forEach(function (it) { m[normInv(it.nomor)] = { d: v2 ? (S(it.dec) || 'N') : 'C', a: S(it.alasan), it: it }; });
        var o = { rec: r, map: m, v2: v2 }; if (!first[r.ref]) first[r.ref] = o; byRef[r.ref] = o;
      } else if (r.status === 'MENUNGGU') pend[r.ref] = r;
    });
    Object.keys(byRef).sort().forEach(function (ref) {
      var o = byRef[ref], lbl = o.rec._data.label || ref;
      Object.keys(o.map).forEach(function (k) { var x = o.map[k]; valMap[k] = DEC[x.d] + ' · ' + lbl + (x.d === 'H' && x.a ? ' — ' + x.a : ''); });
    });
    Object.keys(pend).forEach(function (ref) {
      var r = pend[ref]; r._data.items.forEach(function (it) { var k = normInv(it.nomor); if (!valMap[k]) valMap[k] = '⏳ MENUNGGU · ' + (r._data.label || ref); });
    });
    return { byRef: byRef, first: first, pend: pend, valMap: valMap };
  }
  function valPrev(ctx, ref) { var refs = Object.keys(ctx.valDec || {}).filter(function (r) { return r < ref; }).sort(); return refs.length ? ctx.valDec[refs[refs.length - 1]] : null; }
  /** posisi invoice per awal bulan (sebelum pembayaran apa pun di bulan itu) */
  function awalBulan(i, start) {
    var pb = 0, pn = 0; i._pays.forEach(function (p) { if (p.tgl && p.tgl < start) { pb += p.nominal; pn += p._keN || 0; } });
    return { kas: Math.max(0, i._a.kas - pb), net: Math.max(0, i._a.netto - i._a.pph - pn) };
  }
  /** invoice yang masih terbuka per awal bulan + invoice yang terbit selama bulan itu */
  function valCandidates(ctx, sel) {
    return ctx.invs.filter(function (i) {
      if (i._batal || !i.tgl || i.tgl > sel.end) return false;
      if (i._res.closed && !i._closedManual && i._res.tgl_bayar && i._res.tgl_bayar < sel.start) return false;
      if (i._closedManual && i._res.tgl_tutup && i._res.tgl_tutup < sel.start) return false;
      return awalBulan(i, sel.start).kas > 0.5;
    });
  }
  function paidInMonth(i, sel) { var L = i._pays.filter(function (p) { return p.tgl && p.tgl >= sel.start && p.tgl <= sel.end && !p._pb; }), t = ''; L.forEach(function (p) { if (p.tgl > t) t = p.tgl; }); return { n: sumBy(L, function (p) { return p.nominal; }), tgl: t, L: L }; }

  function validasiData(ctx, ym, opt) {
    var sel = parseMonthSel(ym, ctx.today, false), ref = sel.y + '-' + pad(sel.m, 2);
    var rec = ctx.latest(DOC.VAL, ref), prev = valPrev(ctx, ref), snap = {}, v2rec = rec && rec._data && rec._data.v === 2;
    // opt.fresh: susun dari isian terbaru di Monitoring Invoice (dipakai saat mengajukan), bukan dari lembar yang sudah ada
    if (v2rec && !(opt && opt.fresh)) rec._data.items.forEach(function (it) { snap[normInv(it.nomor)] = it; });
    var prevLbl = prev ? (prev.rec._data.label || prev.rec.ref) : '', same = (ctx.valDec || {})[ref], sameV2 = !!(same && same.v2);
    var items = valCandidates(ctx, sel).map(function (i) {
      var k = normInv(i.nomor), aw = awalBulan(i, sel.start), jt = i._res.jt, pd = paidInMonth(i, sel);
      var grp = !jt ? 'C' : jt < sel.start ? 'A' : jt <= sel.end ? 'B' : 'D';
      var ua = usulCode(i.validasi), ue = ua, carry = '', pv = prev && prev.map[k], prevTxt = '';
      if (pv) {
        prevTxt = 'bulan lalu: ' + DEC[pv.d] + (pv.a ? ' — ' + pv.a : '');
        // "→ BULAN DEPAN" bulan lalu otomatis menjadi usulan ✓ CAIR bulan ini (kecuali AR sudah mengganti usulannya)
        if (pv.d === 'N' && prev.v2 && (S(pv.it.ua) || S(pv.it.u)) && (!ua || ua === S(pv.it.ua))) { ue = 'C'; carry = prevLbl; }
      }
      // revisi di bulan yang sama: pilihan awal = keputusan yang berlaku (Section Head cukup mengubah yang perlu)
      var px = sameV2 && same.map[k], sn = snap[k], rtgl = D(i.rencana), cat = S(i.catatan), pf = '';
      // v2.7: invoice yang sudah cair bulan ini tanpa usulan → ✓ CAIR dengan tanggal penerimaan
      if (!ue && pd.n > 0) { ue = 'C'; }
      if (pd.n > 0 && (ue === 'C' || ue === 'S') && !(rtgl >= sel.start && rtgl <= sel.end)) rtgl = pd.tgl;
      var def = ue || (px ? px.d : 'N'), defA = px && px.d === 'H' ? px.a : '';
      // lembar yang sudah diajukan: nilai, usulan & pilihan awal mengikuti isi saat diajukan (yang diputuskan = yang diajukan)
      if (sn) { ua = S(sn.ua); ue = S(sn.u); carry = S(sn.carry); prevTxt = S(sn.prev) || prevTxt; def = S(sn.def) || def; defA = S(sn.defA); rtgl = S(sn.rtgl); cat = S(sn.cat); pf = S(sn.p); }
      return { nomor: i.nomor, pelanggan: i.pelanggan, npp: i.npp, proyek: i._pro ? i._pro.singkat : '', invke: invLabel(i), tgl: i.tgl, jt: jt, rencana: planDate(i), grp: sn ? (S(sn.grp) || grp) : grp,
        nilai: sn ? N(sn.nilai) : aw.net, sisa: sn ? N(sn.sisa) : aw.kas, status: i._res.status, catatan: cat, rtgl: rtgl, p: pf, cur: px ? px.d : '', curP: px ? S(px.it.p) : '', curT: px ? S(px.it.rtgl) : '', curC: px ? S(px.it.cat) : '', regU: usulCode(i.validasi), regT: D(i.rencana), regC: S(i.catatan), ua: ua, usul: ue, carry: carry, prev: prevTxt, paid: pd.n, paidTgl: pd.tgl,
        def: def, defA: defA, dec: sn ? (S(sn.dec) || def) : def, alasan: sn ? (S(sn.dec) ? S(sn.alasan) : defA) : defA, inSnap: !!sn };
    });
    // invoice yang ada di lembar yang diajukan tetapi tidak lagi memenuhi syarat (mis. dibatalkan) tetap ditampilkan agar keputusannya utuh
    var have = {}; items.forEach(function (x) { have[normInv(x.nomor)] = 1; });
    if (v2rec && !(opt && opt.fresh)) rec._data.items.forEach(function (it) { if (!have[normInv(it.nomor)]) { var x = JSON.parse(JSON.stringify(it)); x.usul = S(it.u); x.dec = S(it.dec) || S(it.def) || 'N'; x.alasan = S(it.dec) ? S(it.alasan) : S(it.defA); x.paid = 0; x.gone = true; x.inSnap = true; items.push(x); } });
    var GO = { A: 0, B: 1, C: 2, D: 3 };
    items.sort(function (a, b) { return (GO[a.grp] - GO[b.grp]) || ((a.jt || '9999') < (b.jt || '9999') ? -1 : (a.jt || '9999') > (b.jt || '9999') ? 1 : b.nilai - a.nilai); });
    var tot = {}; DEC_ORDER.forEach(function (d) { var L = items.filter(function (x) { return x.dec === d; }); tot[d] = { n: L.length, nilai: sumBy(L, function (x) { return x.nilai; }), sisa: sumBy(L, function (x) { return x.sisa; }) }; });
    var fp = hash('VAL2|' + ref + '|' + items.filter(function (x) { return !x.gone; }).map(function (x) { return normInv(x.nomor) + ':' + rnd(x.nilai) + ':' + x.usul; }).join(';'));
    var nApproved = ctx.appr.filter(function (r) { return r.jenis === DOC.VAL && r.ref === ref && r.status === 'DISETUJUI'; }).length;
    var a = rec ? apprPlain(rec, null) : null;
    if (a && rec.status === 'DISETUJUI' && rec._decOk === false) { a.berubah = true; a.teks = '⚠ DATA KEPUTUSAN DIUBAH DI LOG PERSETUJUAN — TIDAK BERLAKU'; }
    return { ref: ref, label: monthLabel(sel.y, sel.m), start: sel.start, end: sel.end, items: items, tot: tot, total: tot.C.nilai, totalSisa: tot.C.sisa, nilai: tot.C.nilai,
      fp: fp, appr: a, jenis: DOC.VAL, v: 2, legacy: !!(rec && rec._data && !v2rec), rev: v2rec ? N(rec._data.rev) : nApproved, revNext: nApproved,
      berubah: !!(v2rec && rec._data.fp0 && rec._data.fp0 !== fp), notProposed: items.filter(function (x) { return !x.ua && !x.carry; }).length, today: ctx.today,
      dasar: v2rec ? S(rec._data.dasar) : '', tglRapat: v2rec ? S(rec._data.tglRapat) : '', curId: same ? same.rec.id : '', curRev: same && same.v2 ? N(same.rec._data.rev) : '', curAt: same ? S(same.rec.putus_waktu) : '',
      recs: ctx.appr.filter(function (r) { return r.jenis === DOC.VAL && r.ref === ref && (r.status === 'DISETUJUI' || r.status === 'MENUNGGU'); }).length };
  }
  /** isi DATA SNAPSHOT saat diajukan (dipakai Code.gs) */
  function valSnapshot(d, opt) {
    opt = opt || {}; var pm = opt.p || {}, pusat = U(opt.dasar) === 'PUSAT';
    (opt.items || []).forEach(function (x) { if (x.p) pm[normInv(x.nomor)] = x.p; });
    return { v: 2, label: d.label, rev: d.revNext, fp0: d.fp, nilai: d.tot.C.nilai, dasar: S(opt.dasar), tglRapat: S(opt.tglRapat), items: d.items.filter(function (x) { return !x.gone; }).map(function (x) {
      var p = S(pm[normInv(x.nomor)]); if (pusat && x.cur === 'C' && x.usul !== 'C') p = '-'; if (p === '+' && (x.usul !== 'C' && x.usul !== 'S')) p = '';
      return { nomor: x.nomor, pelanggan: x.pelanggan, npp: x.npp, proyek: x.proyek, invke: x.invke, tgl: x.tgl, jt: x.jt, rencana: x.rencana, grp: x.grp, nilai: rnd(x.nilai), sisa: rnd(x.sisa),
        status: x.status, ua: x.ua, u: x.usul, carry: x.carry, prev: x.prev, def: x.def, defA: x.defA, paid: rnd(x.paid), paidTgl: x.paidTgl, rtgl: x.rtgl, cat: x.catatan, p: p, cur: x.cur };
    }) };
  }
  /** v2.7: periksa usulan AR sebelum diajukan (tanggal rencana cair wajib untuk ✓ / ◇, di dalam bulan yang diajukan) */
  function valCheckUsulan(d, opt) {
    var bad = d.items.filter(function (x) { return !x.gone && (x.usul === 'C' || x.usul === 'S') && !(x.rtgl >= d.start && x.rtgl <= d.end); });
    if (bad.length) return bad.length + ' invoice ✓ CAIR / ◇ CADANGAN belum ada rencana tanggal cair di ' + d.label + ' (' + bad.slice(0, 3).map(function (x) { return x.nomor; }).join(', ') + (bad.length > 3 ? ', …' : '') + ').';
    if (U(opt && opt.dasar) === 'PUSAT' && !D(opt.tglRapat)) return 'Isi tanggal rapat pusat (dasar revisi: hasil rapat pusat).';
    return '';
  }
  /** v2.7: pembatalan lembar validasi satu bulan (mis. uji coba / urutan input keliru) */
  function valxPackage(ctx, ref) {
    var sel = parseMonthSel(ref, ctx.today, false); ref = sel.y + '-' + pad(sel.m, 2);
    var L = ctx.appr.filter(function (r) { return r.jenis === DOC.VAL && r.ref === ref && (r.status === 'DISETUJUI' || r.status === 'MENUNGGU'); });
    var rec = ctx.latest(DOC.VALX, ref), pend = rec && rec.status === 'MENUNGGU';
    var fp = hash('VALX|' + ref + '|' + L.map(function (r) { return r.id + ':' + r.status; }).join(';'));
    var d = rec && rec._data ? rec._data : {};
    var list = (rec && rec.status !== 'MENUNGGU' && d.list) ? d.list : L.map(function (r) { return { id: S(r.id), label: S(r.label), status: r.status, waktu: r.waktu, putus: r.putus_waktu, kode: S(r.kode) }; });
    return { ref: ref, label: 'Pembatalan lembar validasi ' + monthLabel(sel.y, sel.m), nilai: 0, jenis: DOC.VALX, fp: fp, list: list, alasan: S(d.alasan), appr: rec ? apprPlain(rec, rec.status === 'MENUNGGU' ? fp : null) : null,
      canSubmit: !!L.length && !pend, blok: !L.length ? 'Tidak ada lembar validasi ' + monthLabel(sel.y, sel.m) + ' yang bisa dibatalkan.' : pend ? 'Pembatalan sedang menunggu keputusan Section Head.' : '' };
  }
  /** terapkan keputusan Section Head ke snapshot; kembalikan pesan kesalahan bila tidak lengkap */
  function valApplyDecisions(data, dec) {
    var m = {}; (dec || []).forEach(function (x) { m[normInv(x.nomor)] = x; });
    var miss = [], noReason = [];
    data.items.forEach(function (it) {
      var x = m[normInv(it.nomor)], d = x ? U(x.d) : '';
      if (!DEC[d]) { miss.push(it.nomor); return; }
      it.dec = d; it.alasan = d === 'H' ? S(x.a) : '';
      if (d === 'H' && !it.alasan) noReason.push(it.nomor);
    });
    if (miss.length) return 'Keputusan belum lengkap untuk ' + miss.length + ' invoice (' + miss.slice(0, 3).join(', ') + (miss.length > 3 ? ', …' : '') + ').';
    if (noReason.length) return 'Isi alasan HOLD untuk ' + noReason.length + ' invoice (' + noReason.slice(0, 3).join(', ') + (noReason.length > 3 ? ', …' : '') + ').';
    data.nilai = sumBy(data.items.filter(function (x) { return x.dec === 'C'; }), function (x) { return N(x.nilai); });
    return '';
  }

  // ------------------------------------------------------------------ RA / RI TERMIN (satu tabel: ✓ RA · ◇ CADANGAN · ✚ DI LUAR RA)
  function repRARI(ctx) {
    var sel = parseMonthSel(ctx.sel['RA-RI Termin'], ctx.today, false), ref = sel.y + '-' + pad(sel.m, 2), vd = (ctx.valDec || {})[ref], v0 = (ctx.valFirst || {})[ref];
    var byK = {}; ctx.invs.forEach(function (i) { if (!i._batal) byK[normInv(i.nomor)] = i; });
    var cand = valCandidates(ctx, sel), ra = [], cad = [], keluar = [], note;
    var awalC = {}; if (v0) Object.keys(v0.map).forEach(function (k) { if (v0.map[k].d === 'C') awalC[k] = 1; });
    var rv = function (o) { return o && o.v2 ? ' Rev.' + pad(N(o.rec._data.rev), 2) : ''; };
    var dasar = vd && vd.v2 ? U(vd.rec._data.dasar) : '', tRapat = vd && vd.v2 ? D(vd.rec._data.tglRapat) : '';
    if (vd) {
      Object.keys(vd.map).forEach(function (k) { var i = byK[k]; if (!i) return; if (vd.map[k].d === 'C') ra.push(i); else if (vd.map[k].d === 'S') cad.push(i); });
      Object.keys(awalC).forEach(function (k) { if (byK[k] && !(vd.map[k] && vd.map[k].d === 'C')) keluar.push(byK[k]); });
      note = 'RA AWAL = ' + v0.rec.id + rv(v0) + ' (disetujui ' + v0.rec.putus_waktu + ', dasar laporan ke pusat) · RA FINAL = ' + vd.rec.id + rv(vd) + ' (disetujui ' + vd.rec.putus_waktu + (isDelegasi(vd.rec) ? ', a.n. Section Head (delegasi)' : '') + ')' +
        (dasar === 'PUSAT' ? ' · dasar revisi: rapat pusat ' + (tRapat ? tglPanjang(tRapat) : '') : dasar === 'INTERNAL' ? ' · dasar revisi: koreksi internal' : '') + (vd.v2 ? '' : ' · lembar format lama: semua invoice di lembar dianggap ✓ CAIR') + '. Nilai RA = sisa per awal bulan.';
    } else {
      ra = cand.filter(function (i) { var d = planDate(i); return d && d <= sel.end; });
      note = 'Belum ada Lembar Validasi yang disetujui Section Head untuk ' + monthLabel(sel.y, sel.m) + ' — sementara RA = invoice dengan jatuh tempo / rencana tgl cair s.d. akhir bulan.';
    }
    var byTgl = function (a, b) { return a.tgl < b.tgl ? -1 : a.tgl > b.tgl ? 1 : 0; };
    ra.sort(byTgl); cad.sort(byTgl); keluar.sort(byTgl);
    var inSet = {}; ra.concat(cad).concat(keluar).forEach(function (i) { inSet[normInv(i.nomor)] = 1; });
    var luar = ctx.invs.filter(function (i) { return !i._batal && !inSet[normInv(i.nomor)] && paidInMonth(i, sel).n > 0; }).sort(byTgl);
    var hdr = ['NO', 'PELANGGAN', 'PROYEK', 'NPP', 'NO INVOICE', 'TAGIHAN', 'KATEGORI', 'SUMBER', 'RA AWAL (KAS INCL. PPN)', 'RA FINAL (KAS INCL. PPN)', 'RA FINAL (DPP − PPh)', 'RA TGL CAIR', 'RI NOMINAL CAIR', 'RI TGL CAIR', '% RI / RA FINAL', 'CATATAN'];
    var rows = [['RA TERMIN ' + MONTHS[sel.m - 1].toUpperCase() + ' ' + sel.y + '  |  RI TERMIN ' + MONTHS[sel.m - 1].toUpperCase() + ' ' + sel.y], [note], hdr], st = [sty(0, 'title'), sty(1, 'note'), sty(2, 'hdr')], no = 0, T = {};
    var tglTxt = function (s) { return s ? s.split('-').reverse().join('-') : ''; };
    var sumber = function (i) {
      var k = normInv(i.nomor), it = vd && vd.map[k] ? vd.map[k].it : null, p = it ? S(it.p) : '';
      if (!vd || !v0 || v0 === vd) return 'usulan unit';
      if (awalC[k]) return 'usulan unit';
      return p === '+' || dasar === 'PUSAT' ? '➕ rapat pusat' + (tRapat ? ' ' + tglTxt(tRapat).substring(0, 5) : '') : '➕ revisi internal';
    };
    var sec = function (code, title, list) {
      st.push(sty(rows.length, 'band')); rows.push([title + ' · ' + list.length + ' invoice']);
      var A0 = 0, A = 0, B = 0, Cc = 0;
      list.forEach(function (i) {
        no++; var k = normInv(i.nomor), pd = paidInMonth(i, sel), aw = awalBulan(i, sel.start), x = vd && vd.map[k], it = x ? x.it : null, cat = [];
        var inFinal = code === 'C', inAwal = !!awalC[k] || (!v0 && code === 'C');
        var ra0 = inAwal && code !== 'S' && code !== 'X' ? aw.kas : 0, raK = inFinal ? aw.kas : 0, raN = inFinal ? aw.net : 0;
        if (code === 'X') cat.push('bukan RA' + (x ? ' · validasi Section Head F&HR: ' + DEC[x.d] + (x.a ? ' (' + x.a + ')' : '') : ''));
        if (code === 'K') cat.push('tidak dihitung di RA final');
        if (pd.n && (code === 'C' || code === 'K')) cat.push('realisasi tercatat ' + tglTxt(pd.tgl));
        if (it && S(it.cat)) cat.push(S(it.cat));
        else if (S(i.catatan)) cat.push(S(i.catatan));
        var src = code === 'C' ? sumber(i) : code === 'K' ? (dasar === 'PUSAT' ? '➖ digeser rapat pusat' : '➖ revisi internal') : code === 'S' ? 'cadangan' : '';
        var rt = it && S(it.rtgl) ? S(it.rtgl) : planDate(i);
        rows.push([no, i.pelanggan, i._pro ? i._pro.singkat : '', i.npp, i.nomor, invLabel(i), DEC_KAT[code], src, inAwal && code !== 'S' && code !== 'X' ? ra0 : '', inFinal ? raK : '', inFinal ? raN : '', code === 'C' ? rt : '', pd.n || '', pd.tgl, inFinal && raK ? pd.n / raK : '', cat.join(' · ')]);
        A0 += ra0; A += raK; B += raN; Cc += pd.n;
      });
      st.push(sty(rows.length, 'total')); rows.push(['', 'Subtotal ' + DEC_KAT[code], '', '', '', '', '', '', A0 || (code === 'C' || code === 'K' ? 0 : ''), code === 'C' ? A : '', code === 'C' ? B : '', '', Cc, '', code === 'C' && A ? Cc / A : '', '']);
      T[code] = { a0: A0, ra: A, ran: B, ri: Cc };
    };
    sec('C', '✓ RA FINAL — DIRENCANAKAN CAIR' + (vd ? ' (VALIDASI SECTION HEAD F&HR' + rv(vd) + ')' : ' (JATUH TEMPO S.D. AKHIR BULAN)'), ra);
    if (keluar.length) sec('K', '➖ DIKELUARKAN DARI RA SAAT REVISI' + (dasar === 'PUSAT' ? ' (RAPAT PUSAT)' : ''), keluar); else T.K = { a0: 0, ra: 0, ran: 0, ri: 0 };
    sec('S', '◇ CADANGAN — TIDAK DITARGETKAN, DIKEJAR BILA MEMUNGKINKAN', cad);
    sec('X', '✚ DI LUAR RA — CAIR TANPA RENCANA', luar);
    var tri = T.C.ri + T.K.ri + T.S.ri + T.X.ri, a0 = T.C.a0 + T.K.a0;
    st.push(sty(rows.length, 'total')); rows.push(['', 'TOTAL RA / RI ' + monthLabel(sel.y, sel.m).toUpperCase(), '', '', '', '', '', '', a0, T.C.ra, T.C.ran, '', tri, '', T.C.ra ? tri / T.C.ra : '', '']);
    // RA awal dalam DPP − PPh, tambahan & dikeluarkan
    var a0n = 0, tamK = 0, tamN = 0, kelK = 0, kelN = 0, riAwal = 0;
    ra.forEach(function (i) { var k = normInv(i.nomor), aw = awalBulan(i, sel.start); if (awalC[k] || !v0) a0n += aw.net; else { tamK += aw.kas; tamN += aw.net; } });
    keluar.forEach(function (i) { var aw = awalBulan(i, sel.start); a0n += aw.net; kelK += aw.kas; kelN += aw.net; });
    ra.concat(keluar).forEach(function (i) { if (awalC[normInv(i.nomor)] || !v0) riAwal += paidInMonth(i, sel).n; });
    var r2 = 8 + rows.length + 2, nT = ra.length - (v0 ? ra.filter(function (i) { return awalC[normInv(i.nomor)]; }).length : ra.length);
    var sm = [['RINGKASAN ' + sel.label.toUpperCase(), 'KAS (INCL. PPN, SETELAH PPh)', 'DPP − PPh'],
      ['RA AWAL' + (v0 ? ' (' + v0.rec.id + rv(v0) + ', dilaporkan unit ke pusat)' : ''), a0, a0n],
      ['+ Tambahan' + (dasar === 'PUSAT' ? ' rapat pusat' : ' revisi') + ' (' + nT + ' invoice)', tamK, tamN],
      ['− Dikeluarkan saat revisi (' + keluar.length + ' invoice)', kelK, kelN],
      ['RA FINAL' + (vd ? ' (' + vd.rec.id + rv(vd) + ')' : ''), T.C.ra, T.C.ran],
      ['RI atas ✓ RA final', T.C.ri, ''], ['RI dari ◇ cadangan', T.S.ri, ''], ['RI ✚ di luar RA', T.X.ri + T.K.ri, ''], ['TOTAL RI', tri, ''],
      ['PENCAPAIAN RA FINAL = RI atas RA final ÷ RA final (angka utama)', T.C.ra ? T.C.ri / T.C.ra : '', ''],
      ['PENCAPAIAN RA AWAL = RI atas RA awal ÷ RA awal (pembanding)', a0 ? riAwal / a0 : '', ''],
      ['TOTAL RI ÷ RA FINAL (termasuk cadangan & di luar RA)', T.C.ra ? tri / T.C.ra : '', ''],
      ['RA awal = lembar pertama yang disetujui bulan ini (yang dilaporkan ke pusat); RA final = revisi terakhir yang disetujui. Bila belum ada revisi, keduanya sama.' + (sel.end >= ctx.today ? ' Bulan berjalan: angka masih bergerak sampai akhir bulan.' : ''), '', '']];
    var smB = block(r2, 1, sm, { styles: [sty(0, 'hdr'), sty(4, 'bold'), sty(8, 'bold'), sty(9, 'kpi'), sty(10, 'kpi'), sty(12, 'note')], fmt: { 1: 'money', 2: 'money' }, span: 4 });
    smB.cellFmt = [[9, 1, 'pct'], [10, 1, 'pct'], [11, 1, 'pct']];
    return { from: 8, blocks: [block(8, 1, rows, { styles: st, fmt: { 8: 'money', 9: 'money', 10: 'money', 11: 'date', 12: 'money', 13: 'date', 14: 'pct' } }), smB], wrapCols: [16], minW: { 16: 30 } };
  }
  var DEC_KAT = { C: '✓ RA', S: '◇ CADANGAN', X: '✚ DI LUAR RA', K: '➖ DIKELUARKAN' };

  // ------------------------------------------------------------------ KINERJA PIUTANG & PERINGATAN DINI
  function kinerjaData(ctx) {
    var P = ctx.params, today = ctx.today, act = ctx.invs.filter(function (i) { return !i._batal && i.tgl; });
    var num = function (k, d) { var v = N(P[k]); return v > 0 ? v : d; };
    var TH = { lancar: num('KINERJA_LANCAR_HARI', 15), perhatian: num('KINERJA_PERHATIAN_HARI', 60), kurang: num('KINERJA_KURANG_HARI', 90), lewat: num('EWS_LEWAT_JT_HARI', 60),
      henti: num('EWS_CICILAN_HENTI_HARI', 30), odPel: num('EWS_OVERDUE_PER_PELANGGAN', 2), waiting: num('EWS_WAITING_SUBMIT_HARI', 14), konsen: num('EWS_KONSENTRASI_PCT', 30) };
    var bal = function (d) {
      return sumBy(act, function (i) {
        if (i.tgl > d) return 0;
        var r = i._res;
        if ((r.closed && !i._closedManual && r.tgl_bayar && r.tgl_bayar <= d) || (i._closedManual && r.tgl_tutup && r.tgl_tutup <= d)) return 0;
        var paid = 0; i._pays.forEach(function (p) { if (p.tgl && p.tgl <= d) paid += p.nominal; });
        return Math.max(0, i._a.kas - paid);
      });
    };
    var tag = function (s, e) { return sumBy(act, function (i) { return i.tgl >= s && i.tgl <= e ? i._a.kas : 0; }); };
    var y0 = +today.substring(0, 4), m0 = +today.substring(5, 7), M = [];
    for (var k = 0; k < 24; k++) { var yy = y0 + Math.floor((m0 - 1 - k) / 12), mm = ((m0 - 1 - k) % 12 + 12) % 12 + 1; M.push({ y: yy, m: mm, s: monthStart(yy, mm), e: k === 0 ? today : monthEnd(yy, mm) }); }
    M.forEach(function (x) { x.bal = bal(x.e); x.tag = tag(x.s, x.e); });
    var dsoAt = function (k, n) {
      var L = M.slice(k, k + n); if (L.length < n) return null;
      var avg = sumBy(L, function (x) { return x.bal; }) / n, tg = sumBy(L, function (x) { return x.tag; }), days = diffDays(L[0].e, L[n - 1].s) + 1;
      return tg > 0 ? { dso: avg / tg * days, to: tg / avg * 365 / days, avg: avg, tag: tg } : null;
    };
    var d12 = dsoAt(0, 12), d3 = dsoAt(0, 3), w0 = M[11].s;
    var trend = []; for (var t = 11; t >= 0; t--) { var a12 = dsoAt(t, 12), a3 = dsoAt(t, 3); trend.push([MONTHS[M[t].m - 1].substring(0, 3) + ' ' + String(M[t].y).substring(2), a12 ? rnd(a12.dso) : '', a3 ? rnd(a3.dso) : '', M[t].bal, M[t].tag]); }
    var first = ''; act.forEach(function (i) { if (!first || i.tgl < first) first = i.tgl; });
    // per pelanggan
    var open = openInv(ctx), totOpen = sumBy(open, function (i) { return i._res.sisa; }), pel = {};
    var P0 = function (n) { return pel[n] || (pel[n] = { nama: n, open: 0, od: 0, nOd: 0, w: 0, wd: 0, paid: 0, ontime: 0, paidJt: 0 }); };
    open.forEach(function (i) { var o = P0(i.pelanggan); o.open += i._res.sisa; if (i._res.jt && today > i._res.jt) { o.od += i._res.sisa; o.nOd++; var dd = diffDays(today, i._res.jt); o.w += i._res.sisa; o.wd += i._res.sisa * dd; } });
    ctx.pays.forEach(function (p) {
      var i = p._inv; if (!i || p._pb || !(p.nominal > 0) || !p.tgl || p.tgl < w0 || p.tgl > today) return;
      var o = P0(i.pelanggan); o.paid += p.nominal;
      var jt = i._res.jt; if (!jt) return;
      o.paidJt += p.nominal; if (p.tgl <= jt) o.ontime += p.nominal;
      o.w += p.nominal; o.wd += p.nominal * Math.max(0, diffDays(p.tgl, jt));
    });
    var stat = function (d) { return d === null ? '— (belum ada data)' : d <= TH.lancar ? 'LANCAR' : d <= TH.perhatian ? 'PERHATIAN' : d <= TH.kurang ? 'KURANG LANCAR' : 'BERMASALAH'; };
    var pelL = Object.keys(pel).map(function (n) { var o = pel[n]; o.late = o.w > 0 ? o.wd / o.w : null; o.pct = o.paidJt > 0 ? o.ontime / o.paidJt : null; o.status = stat(o.late); return o; })
      .filter(function (o) { return o.open > 0.5 || o.paid > 0; }).sort(function (a, b) { return b.open - a.open || b.paid - a.paid; });
    var allPaidJt = sumBy(pelL, function (o) { return o.paidJt; }), allOn = sumBy(pelL, function (o) { return o.ontime; });
    // peringatan dini
    var E = [], add = function (lv, kat, pelg, ref, sinyal, tind) { E.push({ lv: lv, kat: kat, pel: pelg, ref: ref, sinyal: sinyal, tind: tind }); };
    open.forEach(function (i) {
      var jt = i._res.jt, r = i._res;
      if (jt && diffDays(today, jt) > TH.lewat) add('MERAH', 'LEWAT JATUH TEMPO', i.pelanggan, i.nomor, invLabel(i) + ' ' + (i._pro ? i._pro.singkat : i.npp) + ' lewat jatuh tempo ' + diffDays(today, jt) + ' hari · sisa Rp ' + fmtRp(r.sisa), 'Eskalasi ke Section Head; surat penagihan / pertemuan dengan pelanggan');
      if (/CICILAN/.test(r.status) && r.tgl_bayar && diffDays(today, r.tgl_bayar) > TH.henti) add('KUNING', 'CICILAN BERHENTI', i.pelanggan, i.nomor, 'Pembayaran terakhir ' + diffDays(today, r.tgl_bayar) + ' hari lalu (' + tglPanjang(r.tgl_bayar) + ') · sisa Rp ' + fmtRp(r.sisa), 'Minta jadwal pelunasan tertulis');
      if (/WAITING SUBMIT/.test(r.status) && diffDays(today, i.tgl) > TH.waiting) add('KUNING', 'WAITING SUBMIT', i.pelanggan, i.nomor, 'Invoice belum disubmit ' + diffDays(today, i.tgl) + ' hari sejak tanggal invoice', 'Segera submit dokumen tagihan ke pelanggan');
    });
    pelL.forEach(function (o) {
      if (o.nOd >= TH.odPel) add('MERAH', 'BANYAK INVOICE OVERDUE', o.nama, o.nOd + ' invoice', o.nOd + ' invoice lewat jatuh tempo · Rp ' + fmtRp(o.od), 'Rapat penagihan; minta komitmen jadwal bayar');
      if (totOpen > 0 && o.open / totOpen * 100 > TH.konsen) add('MERAH', 'KONSENTRASI PIUTANG', o.nama, pctTxt(o.open / totOpen), 'Piutang satu pelanggan ' + pctTxt(o.open / totOpen) + ' dari seluruh piutang terbuka (batas ' + TH.konsen + '%)', 'Pantau mingguan; laporkan ke Section Head');
    });
    Object.keys(ctx.M.pro).forEach(function (n) {
      var p = ctx.M.pro[n]; if (!p.selesai) return;
      var L = open.filter(function (i) { return i.npp === n; });
      if (L.length) add('KUNING', 'PROYEK SELESAI BERSALDO', p.pelanggan, n, 'Proyek ' + p.singkat + ' berstatus SELESAI tetapi masih ada ' + L.length + ' invoice terbuka · Rp ' + fmtRp(sumBy(L, function (i) { return i._res.sisa; })), 'Tagih, atau tutup manual dengan alasan');
    });
    retensiRows(ctx).forEach(function (r) { var p = ctx.M.pro[r[1]]; if (p && p.selesai && r[9] > 0.5) add('KUNING', 'RETENSI BELUM DITAGIH', r[0], r[1], 'Proyek SELESAI, retensi belum ditagih Rp ' + fmtRp(r[9]), 'Cek masa pemeliharaan; terbitkan invoice retensi'); });
    // surat permohonan dengan tenggat yang sudah lewat, invoice di dalamnya masih terbuka
    (ctx.surat || []).forEach(function (r) {
      if (!r.tenggat || r.tenggat >= today) return;
      var keys = S(r.invoices).split(/\s*;\s*/).map(normInv), L = open.filter(function (i) { return keys.indexOf(normInv(i.nomor)) >= 0 && i._res.sisa > 1; });
      if (L.length) add('MERAH', 'TENGGAT SURAT TERLEWAT', r.pelanggan, S(r.nomor), 'Tenggat ' + tglPanjang(r.tenggat) + ' lewat · ' + L.length + ' invoice belum lunas · sisa kas Rp ' + fmtRp(sumBy(L, function (i) { return i._res.sisa; })), 'Hubungi pelanggan; pertimbangkan surat kedua / eskalasi ke Manager');
    });
    var pm = M[1], prevRef = pm.y + '-' + pad(pm.m, 2), pv = (ctx.valDec || {})[prevRef];
    if (pv) {
      var byK = {}; ctx.invs.forEach(function (i) { byK[normInv(i.nomor)] = i; });
      Object.keys(pv.map).forEach(function (k) {
        if (pv.map[k].d !== 'C') return; var i = byK[k]; if (!i) return;
        var got = sumBy(i._pays, function (p) { return p.tgl >= pm.s && p.tgl <= pm.e && !p._pb ? p.nominal : 0; });
        if (got < 0.5) add('KUNING', 'RENCANA CAIR TIDAK TERCAPAI', i.pelanggan, i.nomor, '✓ CAIR ' + monthLabel(pm.y, pm.m) + ' tetapi tidak ada penerimaan di bulan itu', 'Bahas di validasi bulan ini');
      });
    }
    E.sort(function (a, b) { return (a.lv === 'MERAH' ? 0 : 1) - (b.lv === 'MERAH' ? 0 : 1) || (a.kat < b.kat ? -1 : a.kat > b.kat ? 1 : 0); });
    return { TH: TH, d12: d12, d3: d3, trend: trend, first: first, pel: pelL, ontimePct: allPaidJt > 0 ? allOn / allPaidJt : null, ews: E,
      merah: E.filter(function (e) { return e.lv === 'MERAH'; }).length, kuning: E.filter(function (e) { return e.lv === 'KUNING'; }).length, totOpen: totOpen };
  }
  function repKinerja(ctx) {
    var K = ctx._kin || (ctx._kin = kinerjaData(ctx)), TH = K.TH, hari = function (x) { return x ? fmtRp(x.dso) + ' hari' : '—'; };
    var tiles = [['DSO ROLLING 12 BULAN', 'DSO ROLLING 3 BULAN', 'RECEIVABLES TURNOVER (12 BULAN)', 'TEPAT WAKTU (12 BULAN)', 'PERINGATAN MERAH', 'PERINGATAN KUNING'],
      [hari(K.d12), hari(K.d3), K.d12 ? (Math.round(K.d12.to * 10) / 10).toString().replace('.', ',') + '×' : '—', K.ontimePct === null ? '—' : pctTxt(K.ontimePct), K.merah, K.kuning]];
    var blocks = [block(6, 1, tiles, { styles: [sty(0, 'hdr'), sty(1, 'kpi')], span: 2, noWidth: true })];
    blocks.push(block(8, 1, [['Basis nilai: piutang kas (incl. PPN, setelah PPh). DSO = rata-rata saldo piutang akhir bulan ÷ tagihan terbit × jumlah hari; turnover = 365 ÷ DSO. Tepat waktu = nilai yang dibayar ≤ jatuh tempo. Posisi per ' + tglPanjang(ctx.today) +
      (K.first ? ' · data invoice mulai ' + tglPanjang(K.first) + (diffDays(ctx.today, K.first) < 730 ? ' (histori < 2 tahun: angka bulan-bulan awal kurang akurat)' : '') : '') + '.']], { styles: [sty(0, 'note')] }));
    var tr = [['TREN DSO (12 BULAN TERAKHIR)', 'DSO 12 BLN (HARI)', 'DSO 3 BLN (HARI)', 'SALDO PIUTANG AKHIR BULAN', 'TAGIHAN TERBIT']].concat(K.trend);
    blocks.push(block(10, 1, tr, { styles: [sty(0, 'hdr')], fmt: { 1: 'int', 2: 'int', 3: 'money', 4: 'money' }, noWidth: true }));
    var r1 = 25, pr = [['STATUS KINERJA PEMBAYARAN PER PELANGGAN', 'PIUTANG TERBUKA', 'LEWAT JATUH TEMPO', 'INVOICE OVERDUE', 'RATA-RATA TERLAMBAT (HARI)', 'TEPAT WAKTU (12 BLN)', 'DIBAYAR 12 BLN', 'STATUS']];
    K.pel.forEach(function (o) { pr.push([o.nama, o.open, o.od, o.nOd, o.late === null ? '' : rnd(o.late), o.pct === null ? '' : o.pct, o.paid, o.status]); });
    pr.push(['TOTAL', sumBy(K.pel, function (o) { return o.open; }), sumBy(K.pel, function (o) { return o.od; }), sumBy(K.pel, function (o) { return o.nOd; }), '', K.ontimePct === null ? '' : K.ontimePct, sumBy(K.pel, function (o) { return o.paid; }), '']);
    pr.push(['Rata-rata terlambat = hari bayar setelah jatuh tempo, ditimbang nilai (pembayaran 12 bulan + invoice yang masih lewat jatuh tempo dihitung sampai hari ini). Ambang (Parameter): LANCAR ≤ ' + TH.lancar + ' · PERHATIAN ≤ ' + TH.perhatian + ' · KURANG LANCAR ≤ ' + TH.kurang + ' · BERMASALAH > ' + TH.kurang + ' hari. Status ini ukuran internal, bukan kategori CKPN.']);
    blocks.push(block(r1, 1, pr, { styles: [sty(0, 'hdr'), sty(pr.length - 2, 'total'), sty(pr.length - 1, 'note')], fmt: { 1: 'money', 2: 'money', 3: 'int', 4: 'int', 5: 'pct', 6: 'money' } }));
    var r2 = r1 + pr.length + 2, er = [['PERINGATAN DINI (EWS) — ' + K.merah + ' MERAH · ' + K.kuning + ' KUNING'], ['TINGKAT', 'KATEGORI', 'PELANGGAN', 'REFERENSI', 'SINYAL', 'TINDAKAN']];
    if (!K.ews.length) er.push(['✓', 'Tidak ada peringatan', '', '', '', '']);
    K.ews.forEach(function (e) { er.push([(e.lv === 'MERAH' ? '● ' : '● ') + e.lv, e.kat, e.pel, e.ref, e.sinyal, e.tind]); });
    er.push(['Ambang (Parameter): lewat jatuh tempo > ' + TH.lewat + ' hari · cicilan berhenti > ' + TH.henti + ' hari · ≥ ' + TH.odPel + ' invoice overdue per pelanggan · WAITING SUBMIT > ' + TH.waiting + ' hari · konsentrasi > ' + TH.konsen + '% · ✓ CAIR bulan lalu tidak cair · proyek SELESAI bersaldo / retensi belum ditagih.']);
    var eb = block(r2, 1, er, { styles: [sty(0, 'title'), sty(1, 'hdr'), sty(er.length - 1, 'note')] });
    eb.levels = K.ews.map(function (e, ix) { return [2 + ix, e.lv]; });
    blocks.push(eb);
    var charts = [{ key: 'kin1', type: 'LINE', title: 'Tren DSO (hari)', range: [10, 1, 13, 3], row: 10, col: 6, w: 620, h: 270, colors: ['#1F3864', '#eb6834'], legend: 'top' }];
    return { from: 6, blocks: blocks, charts: charts, wrapCols: [5, 6], minW: { 1: 34, 5: 46, 6: 34 } };
  }
  /** 5 tile tambahan di Dashboard (baris kedua) */
  function kinerjaTiles(ctx) {
    var K = ctx._kin || (ctx._kin = kinerjaData(ctx)), sel = parseMonthSel('', ctx.today, false), ref = sel.y + '-' + pad(sel.m, 2), vd = (ctx.valDec || {})[ref];
    var raK = 0, ri = 0;
    if (vd) Object.keys(vd.map).forEach(function (k) { if (vd.map[k].d !== 'C') return; var i = ctx.invs.filter(function (x) { return normInv(x.nomor) === k; })[0]; if (!i) return; raK += awalBulan(i, sel.start).kas; ri += paidInMonth(i, sel).n; });
    var pRA = vd && raK ? ri / raK : null, worse = K.d12 && K.d3 ? K.d3.dso > K.d12.dso : false;
    var t = [
      ['RECEIVABLES TURNOVER', '', '', 'DSO (ROLLING 12 BULAN)', '', '', 'TEPAT WAKTU (12 BULAN)', '', '', 'PENCAPAIAN RA ' + MONTHS[sel.m - 1].toUpperCase(), '', '', 'PERINGATAN DINI', '', ''],
      [K.d12 ? (Math.round(K.d12.to * 10) / 10).toString().replace('.', ',') + '×' : '—', '', '', K.d12 ? fmtRp(K.d12.dso) + ' hari' : '—', '', '', K.ontimePct === null ? '—' : pctTxt(K.ontimePct), '', '', pRA === null ? '—' : pctTxt(pRA), '', '', K.merah + ' merah · ' + K.kuning + ' kuning', '', ''],
      ['12 bulan terakhir', '', '', K.d3 ? '3 bulan: ' + fmtRp(K.d3.dso) + ' hari' : '', '', '', 'nilai dibayar ≤ jatuh tempo', '', '', vd ? 'RI ' + 'Rp ' + (ri / 1e9).toFixed(2).replace('.', ',') + ' M dari RA Rp ' + (raK / 1e9).toFixed(2).replace('.', ',') + ' M' : 'belum ada validasi Section Head', '', '', 'sheet Kinerja & Peringatan Dini', '', ''],
      ['', '', '', K.d12 && K.d3 ? (worse ? '▲ memburuk 3 bulan terakhir' : '▼ membaik 3 bulan terakhir') : '', '', '', '', '', '', pRA === null ? '' : (pRA >= 0.8 ? '✓ ' : '● ') + 'bulan berjalan', '', '', K.merah ? '● perlu tindakan' : '✓ tidak ada sinyal merah', '', '']];
    return { rows: t, flags: [worse ? 'R' : 'G', pRA === null ? 'A' : pRA >= 0.8 ? 'G' : 'A', K.merah ? 'R' : K.kuning ? 'A' : 'G'] };
  }

  // ================================================================== v2.6: PEMBATALAN INVOICE (lewat aplikasi, persetujuan Section Head)
  function batalPackage(ctx, nomor) {
    var k = normInv(nomor), inv = ctx.invs.filter(function (i) { return normInv(i.nomor) === k; })[0];
    if (!inv) return null;
    var pays = ctx.pays.filter(function (p) { return p._inv && normInv(p._inv.nomor) === k && (p.nominal || p._pb); });
    var paySum = sumBy(pays, function (p) { return p.nominal; });
    var rec = ctx.latest(DOC.BATAL, k), pend = rec && rec.status === 'MENUNGGU', done = inv._batal;
    var fp = hash('BATAL|' + k + '|' + pays.length + ':' + rnd(paySum));
    var checks = [];
    if (done) checks.push({ s: 'ok', t: 'Invoice sudah berstatus BATAL' + (inv._batalInfo ? ' (' + inv._batalInfo.id + ')' : '') });
    else checks.push({ s: pays.length ? 'bad' : 'ok', t: pays.length ? 'Sudah ada ' + pays.length + ' pembayaran / pindah buku (Rp ' + fmtRp(paySum) + ', terakhir ' + tglPanjang(pays.map(function (p) { return p.tgl; }).sort().pop() || '') + '). Pindahbukukan ke invoice pengganti (Pembayaran ▸ Pindah buku) atau hapus bila salah input, lalu ajukan lagi.' : 'Belum ada pembayaran tercatat' });
    if (!done) checks.push({ s: pend ? 'bad' : 'ok', t: pend ? 'Pembatalan ' + rec.id + ' sedang menunggu keputusan Section Head' : 'Status dokumen AKTIF, tidak ada pembatalan yang sedang menunggu' });
    if (inv._a.ppn > 0) checks.push({ s: 'warn', t: inv.no_fp ? 'Faktur pajak ' + fpFmt(inv.no_fp) + ': batalkan / ganti di Coretax (sistem tidak bisa melakukannya)' : 'Faktur pajak belum dicatat — bila sudah terbit, batalkan / ganti di Coretax' });
    var mAll = ctx.memo ? ctx.memo.blocks.filter(function (b) { return b.ref === 'INV|' + k; }) : [];
    if (done && mAll.length) { var kb = mAll.filter(function (b) { return /PEMBATALAN/.test(U(b.status)); }).pop(), bb = mAll.filter(function (b) { return /^BATAL/.test(U(b.status)); }).pop();
      checks.push({ s: 'ok', t: kb ? 'Memorial pembatalan ' + kb.kode + ' (jurnal pembalik) dikirim ke file Akuntansi' : bb ? 'Memorial ' + bb.kode + ' ditandai BATAL di file Akuntansi' : 'Memorial diperbarui di file Akuntansi' }); }
    var mb = done ? [] : mAll.filter(function (b) { return !/DIKOREKSI|^BATAL/.test(U(b.status)); });
    if (mb.length) { var vb = mb.filter(function (b) { return b.verif; }); checks.push({ s: 'warn', t: vb.length ? 'Memorial ' + vb[vb.length - 1].kode + ' sudah diverifikasi Akuntansi → dibuat blok KOREKSI (PEMBATALAN) berisi jurnal pembalik' : 'Memorial ' + mb[mb.length - 1].kode + ' belum diverifikasi → ditandai BATAL di file Akuntansi' }); }
    var pk = ctx.latest(DOC.INV, k);
    if (pk && pk.status === 'DISETUJUI') checks.push({ s: 'warn', t: 'Paket invoice ' + pk.id + ' yang sudah disetujui menjadi tidak berlaku' });
    var canSubmit = !done && !pend && !pays.length;
    var d = rec && rec._data ? rec._data : {};
    return { ref: k, inv: plainInv(inv), label: 'Pembatalan invoice ' + inv.nomor + ' — ' + inv.pelanggan, nilai: inv._a.total, jenis: DOC.BATAL, fp: fp,
      checks: checks, canSubmit: canSubmit, done: done, pays: pays.length, alasan: S(d.alasan), pengganti: S(d.pengganti), appr: rec ? apprPlain(rec, done ? null : fp) : null,
      blok: canSubmit ? '' : (done ? 'Invoice sudah dibatalkan.' : pend ? 'Pembatalan sedang menunggu keputusan Section Head.' : 'Pembatalan tidak bisa diajukan: invoice sudah ada pembayarannya.') };
  }

  // ------------------------------------------------------------------ FAKTUR (fungsi murni, dipakai mesin & aplikasi)
  function fakturFields(o) {
    var P = o.params || {}, dpp = rnd(N(o.dpp)), dppnl = rnd(dpp * 11 / 12), ppnCt = rnd(dppnl * 0.12), tg = o.tgl ? o.tgl.split('-') : null, kode = S(P.FAKTUR_KODE_TRANSAKSI) || '04';
    var tgl = tg ? tg[2] + '/' + tg[1] + '/' + tg[0] : '', jasa = S(o.kodeJasa) || S(P.FAKTUR_KODE_JASA), sat = S(P.FAKTUR_SATUAN) || 'Lainnya';
    var f = [
      ['Jenis Faktur', 'Normal', 'Normal'], ['Kode Transaksi', kode + ' — DPP Nilai Lain', kode], ['Tanggal Faktur', tgl, tgl], ['Referensi', S(o.nomor), S(o.nomor)],
      ['NPWP Pembeli', S(o.npwp16), S(o.npwp16)], ['NITKU Pembeli', S(o.nitku), S(o.nitku)], ['Nama Pembeli', S(o.nama), S(o.nama)], ['Alamat Pembeli', S(o.alamat), S(o.alamat)],
      ['Jenis Barang/Jasa', 'Jasa', 'Jasa'], ['Kode Barang/Jasa', jasa || '(isi di Master Proyek)', jasa], ['Nama Barang/Jasa', S(o.uraian), S(o.uraian)], ['Satuan', sat, sat],
      ['Harga Satuan', dpp, String(dpp)], ['Kuantitas', 1, '1'], ['Potongan Harga', 0, '0'], ['DPP (Harga Jual)', dpp, String(dpp)],
      ['DPP Nilai Lain (11/12)', dppnl, String(dppnl)], ['Tarif PPN', '12%', '12'], ['PPN', ppnCt, String(ppnCt)], ['PPnBM', 0, '0']];
    var warn = [];
    if (N(o.ppn) && ppnCt !== rnd(N(o.ppn))) warn.push('PPN hitungan Coretax (' + fmtRp(ppnCt) + ') beda ' + fmtRp(ppnCt - rnd(N(o.ppn))) + ' dari PPN di invoice (' + fmtRp(N(o.ppn)) + '). Periksa sebelum posting.');
    if (!o.npwp16 || S(o.npwp16).length !== 16) warn.push('NPWP pembeli belum 16 digit — lengkapi di Master Pelanggan.');
    if (!jasa) warn.push('Kode barang/jasa Coretax belum diisi (Master Proyek atau Parameter FAKTUR_KODE_JASA).');
    return { fields: f, dpp: dpp, dppnl: dppnl, ppn: ppnCt, ppnInv: rnd(N(o.ppn)), warn: warn };
  }
  function fakturDraft(ctx, inv) {
    var pel = inv._pel || {}, pro = inv._pro || {};
    var r = fakturFields({ params: ctx.params, tgl: inv.tgl, nomor: inv.nomor, npwp16: pel.npwp16, nitku: pel.nitku, nama: S(pel.inv) || S(pel.nama) || inv.pelanggan,
      alamat: [pel.a1, pel.a2, pel.a3].map(S).filter(String).join(', '), kodeJasa: pro.kode_jasa, uraian: uraianOf(inv), dpp: inv._a.dpp, ppn: inv._a.ppn });
    r.nomorFP = inv.no_fp; r.tglFP = inv.tgl_fp; return r;
  }
  /** Buku Penjualan format BUK PEN bagian Pajak (berdasarkan tanggal faktur; invoice tanpa faktur memakai tanggal invoice) */
  function bukuPenjualan(ctx, start, end) {
    return ctx.invs.filter(function (i) { var t = i.tgl_fp || i.tgl; return !i._batal && i._a.ppn > 0 && t >= start && t <= end; })
      .sort(function (a, b) { var ta = a.tgl_fp || a.tgl, tb = b.tgl_fp || b.tgl; return ta < tb ? -1 : ta > tb ? 1 : (a.nomor < b.nomor ? -1 : 1); })
      .map(function (i) { var ns = nsfpParts(i.no_fp), fd = fakturDraft(ctx, i), pel = i._pel || {};
        return { kode: ns.kode || S(ctx.params.FAKTUR_KODE_TRANSAKSI) || '04', status: ns.status, seri: ns.seri, nomorFP: ns.full, tglFP: i.tgl_fp, npwp: pel.npwp16 || '', nama: S(pel.inv) || i.pelanggan,
          uraian: uraianOf(i), harga: i._a.dpp, dppnl: fd.dppnl, ppn: i._a.ppn, sap: i.sap_doc, nomor: i.nomor, tgl: i.tgl, cek: !i.no_fp ? 'Faktur belum dicatat' : (fd.ppn !== i._a.ppn ? 'PPN beda ' + fmtRp(fd.ppn - i._a.ppn) + ' vs Coretax' : '✓') }; });
  }

  // ------------------------------------------------------------------ DATA UNTUK APLIKASI
  function roles(params) {
    var r = {}; ['PEMBUAT', 'PENYETUJU', 'MENGETAHUI', 'PENERIMA'].forEach(function (k) { var u = userByRole(params, k); r[k] = u ? { nama: u.nama, jabatan: u.jabatan, nip: u.nip } : null; });
    return r;
  }
  function bootstrap(res, email) {
    var ctx = res.ctx, M = res.master, P = res.params, me = userBy(P, email) || { email: S(email).toLowerCase(), nama: S(email), jabatan: '', nip: '', peran: '' };
    var invoices = ctx.invs.map(function (i) { var a = i._a, r = i._res, pel = i._pel;
      return { nomor: i.nomor, pelanggan: i.pelanggan, npp: i.npp, tgl: i.tgl, invke: invLabel(i), proyek: i._pro ? i._pro.singkat : '', total: a.total, kas: a.kas, netto: a.netto, ppn: a.ppn, pph: a.pph,
        keP: r.ke_ppn, keN: r.ke_net, diterima: r.diterima, sisa: r.sisa, status: r.status, rev: N(i.rev), status_inv: i._batal ? 'BATAL' : 'AKTIF', batalPend: i._batalPend || '', json: !!S(i.json), sap: i.sap_doc, no_fp: i.no_fp,
        admMax: pel ? N(pel.adm) : N(P.ADM_DEFAULT), appr: i._batal ? '' : apprStatusText(ctx.latest(DOC.INV, normInv(i.nomor)), i._fp),
        jenis: a.jenis, umk: a.umk, ret: a.ret, dpp: a.dpp, pot_pph: r.pot_pph, closed: !!r.closed, tutup: U(i.tutup) === 'CLOSE MANUAL', alasan_tutup: S(i.alasan_tutup), ba_no: S(i.ba_no) }; });
    var payments = ctx.pays.filter(function (p) { return p._inv; }).slice(-300).map(function (p) { var i = p._inv, k = payKey(p), g = ctx.payGroups[k];
      return { key: k, idp: p.idp, tgl: p.tgl, ref: S(p.ref), nomor: i.nomor, pelanggan: i.pelanggan, nominal: p.nominal, keN: p._keN, keP: p._keP, aturan: p._rule, pb: !!p._pb, rekening: p._rek ? p._rek.label : S(p.rekening),
        appr: g ? apprStatusText(ctx.latest(DOC.PAY, k), fpPayGroup(g)) : '' }; });
    var retSisa = {}; retensiRows(ctx).forEach(function (r) { retSisa[r[1]] = r[9]; });
    var umr = {}; umRetRows(ctx).forEach(function (u) { umr[u.npp] = u; });
    var pbMax = 0; ctx.pays.forEach(function (p) { var m = U(p.ref).match(/^PB-(\d+)$/); if (m) pbMax = Math.max(pbMax, +m[1]); });
    var pel = Object.keys(M.pel).map(function (k) { var p = M.pel[k]; return { nama: p.nama, inv: S(p.inv) || p.nama, npwp: S(p.npwp), npwp16: p.npwp16, nitku: p.nitku, kode_bp: p.kode_bp,
      alamat: [p.a1, p.a2, p.a3].map(S).filter(String), jenis: p.jenis, kelompok: p.kelompok, dasar_jt: p.dasar_jt, wakil: S(p.wakil), jab: S(p.jab), aktif: p.aktif, adm: p.adm }; });
    var pro = Object.keys(M.pro).filter(function (k) { return !M.pro[k]._produkOf && !M.pro[k]._adendumOf; }).map(function (k) { var p = M.pro[k]; return { npp: p.npp, pelanggan: p.pelanggan, singkat: p.singkat, deskripsi: S(p.deskripsi), nilai: p.nilai_awal,
      nama_pe: p.nama_pe, produk: nppListTxt(p.produk), selesai: !!p.selesai,
      adendum: p.adds.map(function (a) { return { ke: a.ke, nodok: a.nodok, tgl: a.tgl, jenis: a.jenis, nilai: a.nilai }; }),
      pekerjaan: S(p.pekerjaan), lokasi: S(p.lokasi), kontrak: S(p.kontrak), tgl_kontrak: D(p.tgl_kontrak), nilai_ok: p.nilai_ok,
      um: p.um, ret_pct: p.ret_pct, ret_pos: p.ret_pos, tarif: p.tarif_eff, dasar_pph: p.dasar_pph, tampil_pph: p.tampil_pph, template: p.template, top: p.top,
      signer: p.signer, aktif: p.aktif, retSisa: retSisa[p.npp] || 0, kode_jasa: p.kode_jasa, akun_peng: p.akun_peng,
      umBasis: umr[p.npp] ? umr[p.npp].basis : p.um, umSisa: umr[p.npp] ? (umr[p.npp].lengkap ? umr[p.npp].sisa : '') : p.um, umLengkap: umr[p.npp] ? umr[p.npp].lengkap : true, umKembali: umr[p.npp] ? umr[p.npp].awal + umr[p.npp].kembali : 0,
      retDitahan: umr[p.npp] && umr[p.npp].ret ? umr[p.npp].ret[7] : 0, retBatas: umr[p.npp] ? umr[p.npp].batas : 0 }; });
    var params = {}; Object.keys(P).forEach(function (k) { if (['signers', 'targets', 'akun', 'users', 'ckpn'].indexOf(k) < 0) params[k] = P[k]; });
    var pendTutup = ctx.invs.filter(function (i) { return U(i.tutup) === 'CLOSE MANUAL' && /^CLOSE MANUAL –/.test(i._res.status); }).map(function (i) { return { nomor: i.nomor, pelanggan: i.pelanggan, sisa: i._res.sisa, alasan: S(i.alasan_tutup), status: i._res.status }; });
    var y = +ctx.today.substring(0, 4), m = +ctx.today.substring(5, 7), months = [];
    for (var k2 = -12; k2 <= 2; k2++) { var yy = y + Math.floor((m - 1 + k2) / 12), mm = ((m - 1 + k2) % 12 + 12) % 12 + 1; months.push({ ref: yy + '-' + pad(mm, 2), label: monthLabel(yy, mm) }); }
    var dl = delegasiFor(P, me.email, null, ctx.today);
    return { version: VERSION, today: ctx.today, user: me, roles: roles(P), delegasi: dl, suratAlerts: suratAlerts(ctx), pelanggan: pel, proyek: pro, params: params, signers: P.signers, invoices: invoices, payments: payments,
      approvals: approvalsPlain(ctx).slice(-200), months: months, DOC: DOC, rek: P.rek, nextPB: 'PB-' + pad(pbMax + 1, 4), pendTutup: pendTutup };
  }
  function packageFor(ctx, kind, ref, opt) {
    var d = kind === 'INV' ? invPackage(ctx, ref) : kind === 'PAY' ? payPackage(ctx, ref) : kind === 'VAL' ? validasiData(ctx, ref, opt) : kind === 'REK' ? rekapData(ctx, ref) :
      kind === 'KP' ? kpPackage(ctx, ref) : kind === 'TUTUP' ? tutupPackage(ctx, ref) : kind === 'BATAL' ? batalPackage(ctx, ref) : kind === 'VALX' ? valxPackage(ctx, ref) : null;
    if (d) { d.kind = kind; d.roles = roles(ctx.params); d.today = ctx.today; d.params = { TTD_MODE: S(ctx.params.TTD_MODE) || 'KEDUANYA', KOTA: S(ctx.params.KOTA) || 'Jakarta', SAP_COMPANY_CODE: ctx.params.SAP_COMPANY_CODE, SAP_KODE_UNIT: ctx.params.SAP_KODE_UNIT }; }
    return d;
  }
  // ------------------------------------------------------------------ PERHITUNGAN FORM INVOICE (dipakai aplikasi)
  /** f: {template, jenis, nilaiKontrak, progres, umDiterima, retPct, retPos, tarif, dasar, tampilPPh, items[], dpp, pctUM, ppnOn}
   *  menghasilkan semua angka cetak + angka register. */
  /** angka ketikan form (1.234.567 / 1.234,5 / 1234.5) — sama dengan num() di aplikasi */
  function numForm(v) { if (isNum(v)) return v; var s = S(v); if (!s) return 0; s = s.replace(/[^\d,.\-]/g, ''); if (/,\d{1,2}$/.test(s) && s.indexOf('.') >= 0) s = s.replace(/\./g, '').replace(',', '.'); else if (/^\d{1,3}(\.\d{3})+$/.test(s)) s = s.replace(/\./g, ''); else s = s.replace(/,/g, ''); return parseFloat(s) || 0; }
  /** input calcInvoice dari data proyek (bentuk bootstrap) + isian form; dipakai aplikasi (pratinjau) dan server (pemeriksaan saat simpan)
   *  exUmk = potongan UM invoice yang sedang direvisi (dikembalikan ke sisa UM) */
  function invoiceInput(p, form, exUmk) {
    var umSisa = (p.umLengkap === false || p.umSisa === '' || p.umSisa === null || p.umSisa === undefined) ? undefined : (+p.umSisa || 0) + (+exUmk || 0);
    var k = kontrakAt({ nilai_awal: p.nilai, adds: p.adendum || [], nilai_ok: p.nilai_ok }, form.tgl);
    return { umSisa: umSisa, template: p.template, jenis: form.jenis, nilaiKontrak: k.net, progres: numForm(form.progres), umDiterima: p.um, retPct: p.ret_pct, retPos: p.ret_pos, tarif: p.tarif,
      dasar: p.dasar_pph, tampilPPh: p.tampil_pph, items: (form.items || []).map(function (x) { return { uraian: x.uraian, nilai: numForm(x.nilai), bulan: numForm(x.bulan) }; }), dpp: numForm(form.dpp), ppnOn: form.ppnOn };
  }
  function calcInvoice(f) {
    // Meniru rumus form Excel: angka mentah (desimal) dipakai berantai, dibulatkan saat ditampilkan / di sel ber-ROUND.
    var T = U(f.template) || 'UMUM', J = U(f.jenis) || 'PROGRES', K = N(f.nilaiKontrak), tarif = N(f.tarif) || 0.0265;
    var ROUND = function (x) { return rnd(x); };
    var ppnOf = function (x) { return ROUND(x * 11 / 12 * 0.12); };
    var o = { template: T, jenis: J, nilaiKontrak: K, tarif: tarif, lines: {} };
    if (T === 'WIKA-WITON') {
      var items = (f.items || []).map(function (it) { return { uraian: S(it.uraian), nilai: N(it.nilai), bulan: N(it.bulan), total: N(it.nilai) * N(it.bulan) }; })
        .filter(function (it) { return it.uraian || it.total; });
      var g22 = sumBy(items, function (x) { return x.total; }), g23 = g22 * 0.11;
      o.items = items; o.jumlah = ROUND(g22); o.ppn = ROUND(g23); o.total = ROUND(g22 + g23);
      o.dpp = o.jumlah; o.bruto = o.jumlah; o.umk = 0; o.retensi = 0; o.retPos = 'TIDAK ADA'; o.netto = o.dpp; o.pph = ROUND(tarif * o.dpp);
      o.bayar = o.total; o.showPPh = false; return o;
    }
    if (T === 'FAJARPUTERA') {
      var i17 = N(f.dpp), i18 = i17 * 0.11;
      o.dpp = ROUND(i17); o.ppn = ROUND(i18); o.total = ROUND(i17 + i18); o.bruto = o.dpp; o.umk = 0; o.retensi = 0; o.retPos = 'TIDAK ADA';
      o.netto = o.dpp; o.pph = ROUND(tarif * o.dpp); o.bayar = o.total; o.showPPh = false; return o;
    }
    if (J !== 'PROGRES') {
      var b = N(f.dpp), p = f.ppnOn === false ? 0 : b * 11 / 12 * 0.12, pph0 = T === 'KIDOH' ? flo(tarif * b) : ROUND(tarif * b);
      o.dpp = ROUND(b); o.ppn = f.ppnOn === false ? 0 : ROUND(p); o.total = ROUND(b + o.ppn); o.bruto = o.dpp; o.umk = 0; o.retensi = 0; o.retPos = 'TIDAK ADA';
      o.netto = o.dpp; o.pph = pph0;
      o.showPPh = U(f.tampilPPh) !== 'TIDAK' && T !== 'IKK';
      o.bayar = o.showPPh ? ROUND(b + o.ppn - pph0) : o.total;
      if (T === 'KIDOH' && o.showPPh) o.bayar = flo(b + o.ppn - pph0);
      return o;
    }
    var prog = N(f.progres), pct = K ? prog / K : 0, UM = N(f.umDiterima), rp = N(f.retPct);
    o.pct = pct; o.umDiterima = UM; o.retPct = rp; o.bruto = ROUND(prog); o.progres = ROUND(prog);
    // potongan UM tidak boleh melebihi sisa UM proyek (f.umSisa, bila diketahui)
    var hasCap = f.umSisa !== undefined && f.umSisa !== null && f.umSisa !== '' && UM > 0, cap = hasCap ? Math.max(0, N(f.umSisa)) : Infinity;
    var capUM = function (x) { if (x > cap + 0.5) { o.umCapped = true; o.umAsli = ROUND(x); return cap; } return x; };
    o.umSisa = hasCap ? cap : '';
    if (T === 'IKK') {
      var umPct = K ? UM / K : 0, k19 = capUM(umPct * prog), h21 = prog - k19, h22 = ppnOf(h21), h23 = ROUND(h21 + h22), h24 = rp * prog;
      o.umPct = umPct; o.umk = ROUND(k19); o.dpp = ROUND(h21); o.ppn = h22; o.total = h23; o.retensi = ROUND(h24); o.retPos = rp ? 'AKHIR' : 'TIDAK ADA';
      o.netto = o.dpp - o.retensi;
      var baseI = U(f.dasar) === 'SEBELUM RETENSI' ? o.dpp : o.netto;
      o.pph = ROUND(tarif * baseI); o.showPPh = false;
      o.bayar = h23 - o.retensi; return o;
    }
    if (T === 'KIDOH') {
      var k23 = capUM(pct * UM), h25 = prog - k23, h26 = ppnOf(h25), h27 = h25 + h26, h28 = prog * rp, h29 = flo(tarif * h25);
      o.umk = ROUND(k23); o.dpp = ROUND(h25); o.ppn = h26; o.total = ROUND(h27); o.retensi = ROUND(h28); o.retPos = rp ? 'AKHIR' : 'TIDAK ADA';
      o.netto = o.dpp - o.retensi;
      o.showPPh = U(f.tampilPPh) !== 'TIDAK';
      o.pph = U(f.dasar) === 'SETELAH RETENSI' ? flo(tarif * (h25 - h28)) : h29;
      o.bayar = flo(h27 - h28 - (o.showPPh ? o.pph : 0)); return o;
    }
    // UMUM / MIP — retensi dipotong di awal (sebelum DPP)
    var k19u = prog * rp, k20 = prog - k19u, k23u = capUM(pct * UM), h25u = k20 - k23u, h26u = ppnOf(h25u), h27u = h25u + h26u, h28u = ROUND(tarif * h25u);
    o.retensi = ROUND(k19u); o.retPos = rp ? 'AWAL' : 'TIDAK ADA'; o.progNet = ROUND(k20); o.umk = ROUND(k23u);
    o.dpp = ROUND(h25u); o.ppn = h26u; o.total = ROUND(h27u); o.netto = o.dpp; o.pph = h28u;
    o.showPPh = U(f.tampilPPh) !== 'TIDAK';
    o.bayar = o.showPPh ? ROUND(h27u - h28u) : ROUND(h27u);
    return o;
  }

  // ------------------------------------------------------------------ TERBILANG
  function terbilang(n) {
    n = Math.floor(Math.abs(n));
    var h = ['', 'Satu', 'Dua', 'Tiga', 'Empat', 'Lima', 'Enam', 'Tujuh', 'Delapan', 'Sembilan', 'Sepuluh', 'Sebelas'];
    function t(x) {
      if (x < 12) return h[x];
      if (x < 20) return h[x - 10] + ' Belas';
      if (x < 100) return h[Math.floor(x / 10)] + ' Puluh' + (x % 10 ? ' ' + h[x % 10] : '');
      if (x < 200) return 'Seratus' + (x - 100 ? ' ' + t(x - 100) : '');
      if (x < 1000) return h[Math.floor(x / 100)] + ' Ratus' + (x % 100 ? ' ' + t(x % 100) : '');
      if (x < 2000) return 'Seribu' + (x - 1000 ? ' ' + t(x - 1000) : '');
      if (x < 1e6) return t(Math.floor(x / 1000)) + ' Ribu' + (x % 1000 ? ' ' + t(x % 1000) : '');
      if (x < 1e9) return t(Math.floor(x / 1e6)) + ' Juta' + (x % 1e6 ? ' ' + t(x % 1e6) : '');
      if (x < 1e12) return t(Math.floor(x / 1e9)) + ' Milyar' + (x % 1e9 ? ' ' + t(x % 1e9) : '');
      return t(Math.floor(x / 1e12)) + ' Triliun' + (x % 1e12 ? ' ' + t(x % 1e12) : '');
    }
    return n === 0 ? 'Nol' : t(n).replace(/\s+/g, ' ').trim();
  }
  function toRoman(n) { var m = [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']], r = ''; m.forEach(function (p) { while (n >= p[0]) { r += p[1]; n -= p[0]; } }); return r; }

  // v2.8: historical view is a recomputation from dated source rows, never a writeback.
  function computeAt(input, cutoff) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(S(cutoff)) || !isFinite(dnum(cutoff)) || addDays(cutoff, 0) !== cutoff) throw new Error('Tanggal cutoff tidak valid.');
    var x = JSON.parse(JSON.stringify(input));
    var ai = function(k) { return colIndex(APPR, k); }, ri = function(k) { return colIndex(REG, k); };
    var futureBatal = {}, restoreVal = {};
    (x.approvals || []).forEach(function(r) {
      if (U(r[ai("jenis")]) !== DOC.VALX || U(r[ai("status")]) !== "DISETUJUI" || S(r[ai("putus_waktu")]).slice(0,10) <= cutoff) return;
      var data; try { data = JSON.parse(S(r[ai("data")])); } catch(e) { return; }
      (data.list || []).forEach(function(v) { restoreVal[S(v.id)] = v; });
    });
    (x.approvals || []).forEach(function(r) { var v = restoreVal[S(r[ai("id")])]; if (v && U(r[ai("jenis")]) === DOC.VAL && U(r[ai("status")]) === "DIBATALKAN") { r[ai("status")] = v.status; r[ai("putus_waktu")] = v.putus || ""; r[ai("kode")] = v.kode || ""; } });
    (x.approvals || []).forEach(function(r) { if (U(r[ai('jenis')]) === DOC.BATAL && U(r[ai('status')]) === 'DISETUJUI' && S(r[ai('putus_waktu')]).slice(0,10) > cutoff) futureBatal[normInv(r[ai('ref')])] = true; });
    // peristiwa dokumen yang terjadi setelah cutoff belum ada pada posisi itu (mis. invoice 20-09 disubmit 05-10 → per 30-09 masih WAITING SUBMIT)
    var after = function(r, k) { var d = D(r[ri(k)]); return !!d && d > cutoff; };
    x.register = (x.register || []).map(function(r) {
      if (futureBatal[normInv(r[ri('nomor')])]) r[ri('status_inv')] = 'AKTIF';
      ['submit', 'tt'].forEach(function(k) { if (after(r, k)) r[ri(k)] = ''; });
      if (after(r, 'tgl_fp')) { r[ri('tgl_fp')] = ''; r[ri('no_fp')] = ''; }
      if (after(r, 'ba_tgl')) { r[ri('ba_tgl')] = ''; r[ri('ba_no')] = ''; }
      return r;
    }).filter(function(r) { var d = D(r[ri('tgl')]); return !d || d <= cutoff; });
    x.log = (x.log || []).filter(function(r) { var d = D(r[colIndex(LOG,'tgl')]); return d && d <= cutoff; });
    x.approvals = (x.approvals || []).filter(function(r) { return S(r[ai('waktu')]).slice(0,10) <= cutoff; }).map(function(r) { if (S(r[ai('putus_waktu')]).slice(0,10) > cutoff) { r[ai('status')] = 'MENUNGGU'; r[ai('putus_waktu')] = ''; r[ai('putus_oleh')] = ''; r[ai('kode')] = ''; } return r; });
    x.adendum = (x.adendum || []).filter(function(r) { var d = D(r[colIndex(ADD,'tgl')]); return !d || d <= cutoff; });
    x.tagbrut = (x.tagbrut || []).filter(function(r) { var d = D(r[colIndex(TAG,'tgl')]); return !d || d <= cutoff; });
    x.notes = (x.notes || []).filter(function(r) { var d = D(r[colIndex(NOTE,'tgl')]); return !d || d <= cutoff; });
    x.today = cutoff; x.light = true; delete x.memo;
    x.selectors = x.selectors || {}; ['INKASO','RA-RI Termin','Pareto'].forEach(function(k) { x.selectors[k] = cutoff.slice(0,7); });
    return compute(x);
  }
  // posisi lampau di-cache per tanggal pada ctx: approvalsPlain / paket KP tidak lagi menghitung ulang seluruh sistem untuk tanggal yang sama
  function contextAt(ctx, cutoff) {
    if (!(cutoff && cutoff < ctx.today && ctx._input)) return ctx;
    var c = ctx._atCache || (ctx._atCache = {});
    return c[cutoff] || (c[cutoff] = computeAt(ctx._input, cutoff).ctx);
  }
  function reconciliationRows(ctx) {
    var out = [];
    proyekAR(ctx).forEach(function(n) {
      var k = kartuData(ctx, n, ctx.today); if (!k) return;
      k.groups.forEach(function(g) {
        var i = g.inv, pp = sumBy(i._pays, function(p) { return (p._pph || 0) + (p._sel || 0); });
        var diff = rnd(g.saldo - i._res.sisa), bad = i._res.closed && Math.abs(g.saldo) > 1;
        out.push([i.pelanggan, n, i.nomor, i._res.status, i._a.netto + i._a.ppn, i._res.diterima, i._a.pph, pp, rnd(g.saldo), i._res.sisa, diff,
          bad ? 'PERIKSA: status selesai tetapi buku pembantu bersaldo' : Math.abs(diff) > 1 ? 'Perbedaan basis kas / potongan belum diakui — cocokkan bukti' : 'Saldo sistem konsisten; cocokkan LK/SAP',
          i._pays.map(function(p) { return S(p.bupot); }).filter(Boolean).join('; '), i.sap_doc || '']);
      });
    });
    return out;
  }
  /** tanggal terakhir periode tertutup per hari ini: tutup buku bulan lalu jatuh pada tanggal TUTUP_BUKU_TANGGAL (bawaan 3) pukul 00.00 WIB.
   *  mis. 02-10 → tertutup s.d. 31-08; 03-10 → s.d. 30-09. PERIODE_DITUTUP_SAMPAI hanya bisa memajukan, tidak memundurkan. */
  function periodeTutup(params, today) {
    today = D(today); if (!today) return '';
    var hari = Math.max(1, Math.min(28, Math.floor(N(params.TUTUP_BUKU_TANGGAL)) || 3));
    var auto = U(params.TUTUP_BUKU_OTOMATIS) === 'TIDAK' ? '' : (+today.substring(8, 10) >= hari ? prevMonthEnd(today) : prevMonthEnd(prevMonthEnd(today)));
    var man = D(params.PERIODE_DITUTUP_SAMPAI) || '';
    return man > auto ? man : auto;
  }
  // ------------------------------------------------------------------ SURAT PERMOHONAN PERCEPATAN PEMBAYARAN
  /** invoice yang lewat jatuh tempo per hari ini dan masih bersisa (invoice tanpa jatuh tempo — belum submit / tanda terima — tidak ikut) */
  function suratOverdue(i, today) {
    var r = i._res, a = i._a;
    return !i._batal && !r.closed && !!r.jt && r.jt < today && (a.netto + a.ppn - r.diterima) > 1;
  }
  function suratItem(i, today) {
    var a = i._a, r = i._res, pel = i._pel, tt = pel && pel.dasar_jt === 'TGL TANDA TERIMA';
    var terima = tt ? D(i.tt) : D(i.submit);
    return { nomor: i.nomor, invke: invLabel(i), npp: i.npp, tgl: i.tgl, terima: terima || '', basisTT: !!tt, top: terima && r.jt ? diffDays(r.jt, terima) : '', jt: r.jt,
      lewat: diffDays(today, r.jt), dpp: rnd(a.dpp), ppn: rnd(a.ppn), ret: a.retpos === 'AKHIR' ? rnd(a.ret || 0) : 0, tagihan: rnd(a.netto + a.ppn), pph: rnd(a.pph), pphLabel: pphLabel(i.npp, a.tarif),
      diterima: rnd(r.diterima), sisa: rnd(Math.max(0, a.netto + a.ppn - r.diterima)), kas: rnd(Math.max(0, r.sisa)), no_fp: S(i.no_fp), status: r.status };
  }
  /** data surat untuk satu pelanggan: invoice lewat jatuh tempo dikelompokkan per proyek */
  function suratData(ctx, pelanggan) {
    var key = U(pelanggan), today = ctx.today, P = ctx.M.pel[key] || null, groups = {}, order = [];
    ctx.invs.forEach(function (i) {
      if (U(i.pelanggan) !== key || !suratOverdue(i, today)) return;
      if (!groups[i.npp]) {
        var pro = ctx.M.pro[i.npp] || {};
        groups[i.npp] = { npp: i.npp, singkat: S(pro.singkat), pekerjaan: S(pro.pekerjaan) || S(pro.nama_pe) || S(pro.singkat), kontrak: S(pro.kontrak), tgl_kontrak: D(pro.tgl_kontrak) || '', invoices: [] };
        order.push(i.npp);
      }
      groups[i.npp].invoices.push(suratItem(i, today));
    });
    var projects = order.sort().map(function (n) { var g = groups[n]; g.invoices.sort(function (a, b) { return a.jt < b.jt ? -1 : a.jt > b.jt ? 1 : (a.nomor < b.nomor ? -1 : 1); }); return g; });
    var letters = ctx.surat.filter(function (r) { return r.pelanggan === key; }).map(function (r) { return { nomor: S(r.nomor), tgl: r.tgl, tenggat: r.tenggat, invoices: S(r.invoices), sisa: N(r.sisa), kepada: S(r.kepada), kontak: S(r.kontak), rekening: S(r.rekening), link: S(r.link) }; });
    return { today: today, pelanggan: key, nama: P ? (S(P.inv) || P.nama) : key, alamat: P ? [P.a1, P.a2, P.a3].map(S).filter(String) : [], projects: projects, letters: letters };
  }
  /** pelanggan yang perlu dikirimi surat: punya invoice lewat jatuh tempo yang belum masuk surat dalam SURAT_ULANG_HARI terakhir (bawaan 14) */
  function suratAlerts(ctx) {
    var today = ctx.today, ulang = N(ctx.params.SURAT_ULANG_HARI) || 14, by = {};
    ctx.invs.forEach(function (i) {
      if (!suratOverdue(i, today)) return;
      var k = U(i.pelanggan), o = by[k] || (by[k] = { pelanggan: k, n: 0, sisa: 0, maxLewat: 0, proyek: {}, baru: 0 });
      var it = suratItem(i, today), nk = normInv(i.nomor);
      o.n++; o.sisa += it.sisa; o.maxLewat = Math.max(o.maxLewat, it.lewat); o.proyek[i.npp] = 1;
      var covered = ctx.surat.some(function (r) { return r.pelanggan === k && r.tgl && diffDays(today, r.tgl) <= ulang && S(r.invoices).split(/\s*;\s*/).map(normInv).indexOf(nk) >= 0; });
      if (!covered) o.baru++;
    });
    return Object.keys(by).map(function (k) {
      var o = by[k], L = ctx.surat.filter(function (r) { return r.pelanggan === k; }).sort(function (a, b) { return a.tgl < b.tgl ? 1 : -1; })[0];
      o.proyek = Object.keys(o.proyek).length; o.sisa = rnd(o.sisa); o.perlu = o.baru > 0;
      o.terakhir = L ? { nomor: S(L.nomor), tgl: L.tgl, tenggat: L.tenggat } : null;
      return o;
    }).sort(function (a, b) { return (b.perlu ? 1 : 0) - (a.perlu ? 1 : 0) || b.sisa - a.sisa; });
  }
  /** PPh yang sudah diakui pada pelunasan tetapi nomor bukti potong belum dicatat — tanpa bupot, PPh tidak bisa dikreditkan */
  function bupotPendingRows(ctx) {
    var out = [];
    ctx.pays.forEach(function(p) {
      var i = p._inv, x = (p._pph || 0) + (p._sel || 0);
      if (!i || i._batal || x <= 0 || S(p.bupot)) return;
      out.push([i.pelanggan, i.npp, i.nomor, p.tgl, S(p.ref), pphLabel(i.npp, i._a.tarif), rnd(x), diffDays(ctx.today, p.tgl), p.idp || '']);
    });
    return out.sort(function(a, b) { return b[7] - a[7]; });
  }
  function monthlyMetrics(input, cutoff) {
    var res = computeAt(input, cutoff), ctx = res.ctx, start = cutoff.slice(0,7) + '-01', prev = computeAt(input, prevMonthEnd(cutoff)), customers = {}, saldo = 0, opening = 0, overdue = 0, cash = 0, receipt = 0, buckets = [0,0,0,0,0,0];
    function ledger(c, collect) { var total = 0; proyekAR(c).forEach(function(n) { var k = kartuData(c,n,c.today); if (!k) return; total += k.T.saldo; if (collect) k.groups.forEach(function(g) { var i = g.inv, s = g.saldo; if (s <= 1) return; customers[i.pelanggan] = (customers[i.pelanggan] || 0) + s; }); }); return rnd(total); }
    saldo = ledger(ctx,true); opening = ledger(prev.ctx,false);
    ctx.invs.forEach(function(i) { if (i._batal || i._res.closed) return; var s = Math.max(0,i._res.sisa), jt = i._res.jt, d = jt ? diffDays(cutoff,jt) : null; cash += s; if (d > 0) overdue += s; var j = d === null ? 0 : d <= 0 ? 1 : d <= 30 ? 2 : d <= 90 ? 3 : d <= 180 ? 4 : 5; buckets[j] += s; });
    ctx.pays.forEach(function(p) { if (p._inv && !p._inv._batal && !p._pb && p.nominal > 0 && p.tgl >= start) receipt += p.nominal; });
    var rar = repRARI(ctx), sr = rar.blocks[1].rows;
    var top = Object.keys(customers).map(function(k) { return { pelanggan:k, saldo:rnd(customers[k]) }; }).sort(function(a,b) { return b.saldo-a.saldo; }).slice(0,5);
    return { cutoff:cutoff, nextMonth:addDays(cutoff,1).slice(0,7), opening:opening, saldo:saldo, perubahan:saldo-opening, kasTerbuka:rnd(cash), overdue:rnd(overdue), penerimaan:rnd(receipt),
      targetInkaso:ctx.params.targets[cutoff.slice(0,7)] || 0, raAwal:sr[1][1], raFinal:sr[4][1], riAtasRA:sr[5][1], totalRI:sr[8][1], aging:buckets.map(rnd), top:top,
      issues:res.issues.length, bedaStatus:reconciliationRows(ctx).filter(function(r) { return /^PERIKSA:/.test(r[11]); }).length,
      basis:'Saldo buku pembantu = tagihan dikurangi kas dan potongan diakui. Aging memakai sisa kas. Berdasarkan data dan master yang tersedia saat snapshot dibuat; bukan arsip master historis.' };
  }

  return {
    computeAt: computeAt, invoiceInput: invoiceInput, numForm: numForm, monthlyMetrics: monthlyMetrics, reconciliationRows: reconciliationRows, periodeTutup: periodeTutup, suratData: suratData, suratAlerts: suratAlerts, SURAT: SURAT, bupotPendingRows: bupotPendingRows, invAt: invAt, VERSION: VERSION, REG: REG, REG_ALIAS: REG_ALIAS, LOG: LOG, MPEL: MPEL, MPRO: MPRO, MPRO_ALIAS: MPRO_ALIAS, ADD: ADD, TAG: TAG, NOTE: NOTE, APPR: APPR, OPT: OPT, DOC: DOC,
    MONTHS: MONTHS, CKPN_DEFAULT: CKPN_DEFAULT, kontrakAt: kontrakAt,
    compute: compute, calcInvoice: calcInvoice, umRetRows: umRetRows, kpFp: kpFp, rekOf: rekOf, kartuSignBlock: kartuSignBlock, invoiceAmounts: invoiceAmounts, kartuBlocks: kartuBlocks, kartuPlain: kartuPlain,
    invPackage: invPackage, payPackage: payPackage, packageFor: packageFor, bootstrap: bootstrap, allocRule: allocRule, fakturFields: fakturFields, bukuPenjualan: bukuPenjualan, roles: roles, validasiData: validasiData, rekapData: rekapData, approvalsPlain: approvalsPlain, fakturDraft: fakturDraft,
    fpInvoice: fpInvoice, fpPayGroup: fpPayGroup, payKey: payKey, userBy: userBy, userByRole: userByRole, hash: hash, rekon: rekon,
    readParams: readParams, buildMaster: buildMaster, signerFor: signerFor, tarifDefault: tarifDefault, pphLabel: pphLabel,
    terbilang: terbilang, toRoman: toRoman, retTxt: retTxt, normInv: normInv, nppValid: nppValid, nppKey: nppKey, npwp16: npwp16, D: D, N: N, S: S, U: U, rnd: rnd, flo: flo,
    monthLabel: monthLabel, colIndex: colIndex, fmtRp: fmtRp, tglPanjang: tglPanjang, prevMonthEnd: prevMonthEnd, addDays: addDays, parseMonthSel: parseMonthSel,
    MEMO: MEMO, PEREF: PEREF, MEMO_KODE: MEMO_KODE, memorial: memorial, memoParse: memoParse, memoSheetRows: memoSheetRows, peSummary: peSummary, peSummaryBlocks: peSummaryBlocks,
    delegasiFor: delegasiFor, isDelegasi: isDelegasi, sapParse: sapParse, sapMatch: sapMatch, normNpp: normNpp, nppList: nppList, nppListTxt: nppListTxt, proyekAR: proyekAR, fpFmt: fpFmt, MONTHS_SHORT: BLN3,
    DEC: DEC, DEC_ORDER: DEC_ORDER, VGRP: VGRP, usulCode: usulCode, valSnapshot: valSnapshot, valCheckUsulan: valCheckUsulan, valApplyDecisions: valApplyDecisions, valSidik: valSidik, valDecHash: valDecHash, kinerjaData: kinerjaData
  };
})();
if (typeof module !== 'undefined') module.exports = ENGINE;
