/**
 * SISTEM AR CSU — Concrete Services Unit, PT Wijaya Karya Beton Tbk
 * Dirancang dan dikembangkan oleh Fikri Zikra Akrama (AR Officer, Concrete Services Unit), 2026.
 * Penggandaan sistem, kode, maupun format laporan tanpa izin pengembang tidak diperkenankan.
 */
/**
 * SISTEM AR v2 — CONCRETE SERVICES UNIT, PT WIJAYA KARYA BETON TBK
 * Kode utama Apps Script: menu, pembaruan otomatis, penulis laporan, persetujuan digital, arsip PDF, dan API aplikasi.
 * File lain: Engine.gs (perhitungan), EngineClient.html (salinan Engine untuk aplikasi), Index.html (aplikasi), appsscript.json.
 */

var REPORT_SHEETS = ['Dashboard Summary', 'Kartu Piutang', 'Pareto', 'RA-RI Termin', 'INKASO', 'Aging Piutang', 'Monitoring UM & Retensi', 'Monitoring Tagbrut',
  'Rekap Invoice Keluar', 'Usulan Jurnal SAP', 'CKPN', 'Rekon SAP', 'Cek Data', 'Kinerja & Peringatan Dini'];
var SELECTOR_SHEETS = ['Pareto', 'Rekap Invoice Keluar', 'Kartu Piutang', 'INKASO', 'RA-RI Termin', 'Monitoring Tagbrut', 'CKPN', 'Usulan Jurnal SAP'];
var INPUT_SHEETS = ['Monitoring Invoice', 'Log Penerimaan', 'Master Pelanggan', 'Master Proyek', 'Riwayat Adendum', 'Parameter', 'Input Tagbrut', 'Data LK (Tempel)'];
var NUMF = { money: '#,##0;(#,##0);"-"', money0: '#,##0;(#,##0);0', date: 'dd-mmm-yyyy', pct: '0.00%', int: '0', num2: '#,##0.00' };
var C = { navy: '#1F3864', tot: '#D9E1F2', kpi: '#FFF2CC', input: '#FFF2CC', auto: '#E2EFDA', band: '#EEF3FA', app: '#2F5597', manual: '#9C6500', autoHdr: '#1D6B73' };
var TZ = 'Asia/Jakarta';
var PENGEMBANG = 'Fikri Zikra Akrama';

// =====================================================================  MENU & PEMICU
function onOpen() {
  // v2.8.5: menu harian di atas, sisanya di sub-menu. Alat sekali jalan (v2.6) & Ekspor Buku Penjualan (diganti file Pajak) dihapus dari menu —
  // fungsinya tetap ada dan bisa dijalankan dari editor Apps Script bila diperlukan
  var ui = SpreadsheetApp.getUi(), v28 = false; try { v28 = arProps_().getProperty('V28_DONE') === '1'; } catch (e) { }
  var m = ui.createMenu('AR CSU')
    .addItem('Perbarui semua laporan', 'rebuildAll')
    .addItem('Buka aplikasi (jendela)', 'openAppDialog')
    .addItem('Link aplikasi (lihat / simpan)', 'showAppLink');
  if (!v28) m = m.addItem('Pasang pembaruan v2.8 (backup dahulu)', 'applyV28');
  m.addSeparator()
    .addSubMenu(ui.createMenu('Analisis bulanan')
      .addItem('Buat / lihat analisis bulan lalu', 'arMonthlyMenu')
      .addItem('Revisi analisis bulanan…', 'arMonthlyRevisionMenu'))
    .addSubMenu(ui.createMenu('Akuntansi, P&E & Pajak')
      .addItem('Perbarui file Akuntansi (memorial), P&E & Pajak sekarang', 'syncTurunan')
      .addItem('Isi No. Dok SAP dari Tempel SAP', 'fillSapFromTempel')
      .addItem('Isi No. Dok SAP dari Data LK', 'fillSapFromLK')
      .addItem('Bangun ulang blok memorial (baris terpilih)', 'memoRebuildSelected'))
    .addSubMenu(ui.createMenu('Data proyek')
      .addItem('Ganti NPP proyek…', 'renameNpp')
      .addItem('Pindahkan proyek ARSIP ke sheet arsip…', 'arArsipPindahMenu')
      .addItem('Kembalikan proyek dari arsip…', 'arArsipKembalikanMenu'))
    .addSubMenu(ui.createMenu('Pengaturan (pemilik)')
      .addItem('Aktifkan otomatisasi (sekali saja)', 'setupTriggers')
      .addItem('Kunci struktur & periksa hak akses', 'lockSystem')
      .addItem('Delegasi persetujuan (sementara)…', 'aturDelegasi')
      .addItem('Sahkan pengajuan menunggu (TTD basah)…', 'arSahkanTtdBasah')
      .addItem('Tutup periode (kunci transaksi s.d. tanggal)…', 'arTutupPeriodeMenu')
      .addItem('Siapkan folder arsip PDF', 'setupArsipFolder')
      .addItem('Aktifkan pilihan tanggal P&E', 'arEnablePE')
      .addItem('Aktifkan file Pajak (bagikan ke peran PAJAK)', 'arEnablePajak')
      .addItem('Format tanggal Indonesia (dd/mm/yyyy)…', 'aturFormatTanggal'))
    .addSeparator()
    .addItem('Tentang sistem ini', 'aboutSystem')
    .addToUi();
}

function setupTriggers() {
  var ss = arSpreadsheet_();
  arRequireOwner_();
  arInstallTriggers_(ss);
  secret_();
  rebuildAll();
  SpreadsheetApp.getUi().alert('Otomatisasi aktif.\n\n• Setiap perubahan di sheet input memperbarui laporan (langsung bila sistem tidak sibuk, paling lambat 5 menit).\n• Setiap pagi pukul 06.00 status jatuh tempo diperbarui.\n• Setiap jam file Akuntansi (memorial) & P&E diperbarui, dan centang verifikasi Akuntansi ditarik ke sistem.');
}
// Fungsi pemicu juga bisa dipanggil lewat google.script.run, jadi tanpa event trigger yang sah tetap diperlakukan sebagai panggilan pengguna.
function dailyRebuild(e) { if (!arFromTrigger_(e)) arRequireMaker_();
  // v2.8.7: tanggal posisi pilihan manual hanya sementara — setiap pagi kembali ke HARI INI (sumber & file P&E)
  try { var kp0 = sheet_('Kartu Piutang'); if (kp0 && String(kp0.getRange('E5').getValue()).trim()) kp0.getRange('E5').clearContent(); } catch (x0) { }
  try { arPEResetCutoff_(); } catch (x1) { }
  rebuildAll_({ format: true }); arMonthlyAuto_(); try { arPropsCleanup_(); } catch (x) { } }
/** jadwal 5 menit: menyelesaikan pembaruan yang ditandai saat sistem sibuk */
function rebuildIfDirty(e) { if (!arFromTrigger_(e)) arRequireMaker_(); arRebuildIfDirty_(25000); }

function onEditInstalled(e) {
  // installable trigger berjalan sebagai pemilik trigger; Session.getActiveUser() pengedit sering kosong, jadi tidak memakai cek akses per pengguna
  if (!e || !e.range || !arFromTrigger_(e)) return;
  var sh = e.range.getSheet(), name = sh.getName(), row = e.range.getRow(), col = e.range.getColumn();
  if (name === 'Analisis Bulanan AR' && row === 5 && col === 3) { var snaps = arMonthlyRows_(), id = String(sh.getRange('C5').getValue()); var hit = snaps.filter(function(r) { return r[0] === id; })[0]; if (hit) arRenderMonthly_(hit); return; }
  if (name === 'Pareto') {
    if (row === 5 && col === 3) return arEditRebuild_();
    var map = JSON.parse(arProps_().getProperty('EDIT_' + name) || 'null');
    if (!map) return;
    var nr = e.range.getNumRows(), nc = e.range.getNumColumns(), changed = false;
    // kolom terkunci (keputusan Section Head): ketikan manual dikembalikan
    if ((map.lock || []).some(function (c) { return c >= col && c <= col + nc - 1; }) && row <= map.last && row + nr - 1 >= map.first) changed = true;
    for (var r = row; r < row + nr; r++) {
      if (r < map.first || r > map.last) continue;
      for (var c = col; c < col + nc; c++) {
        var key = map.cols[String(c)];
        if (!key) continue;
        var nomor = String(sh.getRange(r, map.keyCol).getValue()).trim();
        if (nomor) { setRegisterValue_(nomor, key, sh.getRange(r, c).getValue()); changed = true; }
      }
    }
    if (changed) arEditRebuild_();
    return;
  }
  if (name === ENGINE.MEMO.sheet) { if (row >= ENGINE.MEMO.firstRow) { memoMarkManual_(sh, row, e.range.getNumRows(), col, e.range.getNumColumns()); arEditRebuild_(); } return; }
  if (name === ENGINE.PEREF.sheet && row >= ENGINE.PEREF.firstRow) return arEditRebuild_();
  if (SELECTOR_SHEETS.indexOf(name) >= 0 && row === 5 && col === 3) return arEditRebuild_();
  if (name === 'Kartu Piutang' && row === 5 && col === 5) return arEditRebuild_();
  if (INPUT_SHEETS.indexOf(name) >= 0) {
    if (name === 'Monitoring Invoice' && row < ENGINE.REG.firstRow) return;
    if (name === 'Log Penerimaan' && row < ENGINE.LOG.firstRow) return;
    if (name === 'Monitoring Invoice' || name === 'Log Penerimaan') { try { arAuditEdit_(e, sh, name); } catch (err) { } }
    if (name === 'Riwayat Adendum' && row < ENGINE.ADD.firstRow) return;
    if (name === 'Input Tagbrut' && row < ENGINE.TAG.firstRow) return;
    if (name === 'Data LK (Tempel)' && row < 8 && !(row === 4 && col === 3)) return;
    arEditRebuild_();
  }
}

/** penghapusan baris tidak memicu onEdit: perubahan struktur → tandai perlu diperbarui (rebuild membandingkan ID baris) */
function onChangeInstalled(e) {
  if (!e || !arFromTrigger_(e)) return;
  if (['REMOVE_ROW', 'INSERT_ROW', 'OTHER'].indexOf(String(e.changeType)) < 0) return; // EDIT sudah ditangani onEditInstalled
  arEditRebuild_();
}

// =====================================================================  BACA BERDASARKAN JUDUL KOLOM
function norm_(x) { return String(x === null || x === undefined ? '' : x).replace(/\s+/g, ' ').trim().toUpperCase(); }
function colType_(c) { return typeof c[2] === 'string' ? c[2] : 'manual'; }
function colWidth_(c) { return typeof c[2] === 'string' ? c[3] : c[2]; }
function colFmt_(c) { return typeof c[2] === 'string' ? c[4] : c[3]; }
/** peta kunci → nomor kolom (1-based); 0 bila judul tidak ditemukan */
function colMap_(sh, schema, aliases) {
  var lc = Math.max(sh.getLastColumn(), 1), hdr = sh.getRange(schema.headerRow, 1, 1, lc).getValues()[0].map(norm_), used = {}, map = {};
  schema.cols.concat((aliases && aliases.addendum_lama) ? [['addendum_lama', '']] : []).forEach(function (c) {
    var names = (c[1] ? [norm_(c[1])] : []).concat(((aliases || {})[c[0]] || []).map(norm_)).concat(((ENGINE.HDR_LAMA || {})[schema.sheet + '|' + c[0]] || []).map(norm_));
    map[c[0]] = 0;
    for (var j = 0; j < hdr.length; j++) if (!used[j] && hdr[j] && names.indexOf(hdr[j]) >= 0) { used[j] = 1; map[c[0]] = j + 1; break; }
  });
  return map;
}
/** baris data dalam urutan skema (kolom yang tidak ada → '') */
function readSchema_(sh, schema, aliases, extraKeys) {
  if (!sh) return [];
  var last = sh.getLastRow(), lc = sh.getLastColumn();
  if (last < schema.firstRow || lc < 1) return [];
  var map = colMap_(sh, schema, aliases), keys = schema.cols.map(function (c) { return c[0]; }).concat(extraKeys || []);
  var data = sh.getRange(schema.firstRow, 1, last - schema.firstRow + 1, lc).getValues();
  return data.map(function (r) { return keys.map(function (k) { var j = map[k]; return j ? nd_(r[j - 1]) : ''; }); });
}
/** menambahkan kolom skema yang belum ada (di ujung kanan) — aman dijalankan berulang */
function ensureSchema_(sh, schema, aliases, typeRow) {
  var map = colMap_(sh, schema, aliases), add = schema.cols.filter(function (c) { return !map[c[0]]; });
  if (!add.length) return 0;
  var lc = sh.getLastColumn(), maxR = Math.max(sh.getMaxRows(), schema.firstRow + 50);
  add.forEach(function (c, i) {
    var col = lc + 1 + i, t = colType_(c), bg = t === 'auto' ? C.autoHdr : t === 'app' ? C.app : t === 'manual' ? C.manual : C.navy;
    if (col > sh.getMaxColumns()) sh.insertColumnsAfter(sh.getMaxColumns(), col - sh.getMaxColumns());
    sh.getRange(schema.headerRow, col).setValue(c[1]).setBackground(bg).setFontColor('#FFFFFF').setFontWeight('bold').setWrap(true).setHorizontalAlignment('center').setVerticalAlignment('middle').setFontSize(9);
    if (typeRow) sh.getRange(typeRow, col).setValue(t === 'auto' ? 'OTOMATIS' : t === 'app' ? 'DARI APLIKASI' : 'DIISI MANUAL').setBackground(bg).setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(7).setHorizontalAlignment('center');
    var body = sh.getRange(schema.firstRow, col, maxR - schema.firstRow + 1, 1).setFontSize(9).setFontFamily('Arial');
    if (t === 'manual') body.setBackground(C.input); else if (t === 'auto') body.setBackground(C.auto);
    var f = colFmt_(c); if (f && NUMF[f]) body.setNumberFormat(NUMF[f]);
    sh.setColumnWidth(col, (colWidth_(c) || 14) * 7);
  });
  return add.length;
}
function sheet_(name) { return arSpreadsheet_().getSheetByName(name); }
function rangeVals_(sh, a1) { try { return nd_(sh.getRange(a1).getValues()); } catch (e) { return []; } }

// =====================================================================  TANGGAL & ZONA WAKTU
// Semua tanggal dibaca & ditulis memakai zona waktu FILE spreadsheet (bukan zona waktu Apps Script),
// sehingga tanggal tidak bergeser walaupun setelan keduanya berbeda.
var SS_TZ_ = null;
function ssTz_() { if (!SS_TZ_) { try { SS_TZ_ = arSpreadsheet_().getSpreadsheetTimeZone() || TZ; } catch (e) { SS_TZ_ = TZ; } } return SS_TZ_; }
function isDate_(v) { return Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime()); }
/** Date dari sel → 'yyyy-mm-dd' sesuai tanggal yang TAMPIL di sel */
function dstr_(v, tz) { return Utilities.formatDate(v, tz || ssTz_(), 'yyyy-MM-dd'); }
/** konversi semua Date di dalam array/objek menjadi 'yyyy-mm-dd' */
function nd_(x) {
  if (isDate_(x)) { var t = Utilities.formatDate(x, ssTz_(), 'HH:mm'); return t === '00:00' ? dstr_(x) : Utilities.formatDate(x, ssTz_(), 'yyyy-MM-dd HH:mm'); }
  if (Array.isArray(x)) return x.map(nd_);
  if (x && typeof x === 'object' && Object.prototype.toString.call(x) === '[object Object]') { var o = {}; Object.keys(x).forEach(function (k) { o[k] = nd_(x[k]); }); return o; }
  return x;
}
/** pukul 00.00 tanggal y-m-d menurut zona waktu tz (default: zona waktu file) */
function dateAt_(y, m, d, tz) {
  tz = tz || ssTz_();
  var base = Date.UTC(y, m - 1, d), off = 0;
  for (var k = 0; k < 2; k++) { // dua kali agar tepat pada pergantian DST
    var z = Utilities.formatDate(new Date(base - off), tz, 'Z'), sg = z.charAt(0) === '-' ? -1 : 1;
    off = sg * (Number(z.substr(1, 2)) * 60 + Number(z.substr(3, 2))) * 60000;
  }
  return new Date(base - off);
}
function readLK_() {
  var sh = sheet_('Data LK (Tempel)'); if (!sh) return {};
  var last = sh.getLastRow(), out = { periode: sh.getRange('C4').getValue(), pu: [], ppn: [], ret: [] };
  if (last >= 8) { var n = last - 7; out.pu = sh.getRange(8, 1, n, 14).getValues(); out.ppn = sh.getRange(8, 17, n, 15).getValues(); out.ret = sh.getRange(8, 33, n, 13).getValues(); }
  return nd_(out);
}
function readAll_() {
  var ss = arSpreadsheet_(), par = ss.getSheetByName('Parameter'), it = ss.getSheetByName('Input Tagbrut');
  var sel = {};
  SELECTOR_SHEETS.forEach(function (n) { var s = ss.getSheetByName(n); if (s) sel[n] = nd_(s.getRange('C5').getValue()); });
  var kps = ss.getSheetByName('Kartu Piutang'); if (kps) sel['Kartu Piutang|asof'] = nd_(kps.getRange('E5').getValue());
  return {
    today: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'), ssTz: ssTz_(),
    paramKV: rangeVals_(par, 'A4:C120'), paramSigners: rangeVals_(par, 'E4:G15'), paramTargets: rangeVals_(par, 'I4:K100'),
    paramAkun: rangeVals_(par, 'O4:T80'), paramUsers: rangeVals_(par, 'U4:Y40'), paramCkpn: rangeVals_(par, 'AA4:AD30'), paramRek: rangeVals_(par, 'AF4:AH20'),
    pelanggan: readSchema_(ss.getSheetByName('Master Pelanggan'), ENGINE.MPEL),
    proyek: readSchema_(ss.getSheetByName('Master Proyek'), ENGINE.MPRO, ENGINE.MPRO_ALIAS, ['addendum_lama']),
    adendum: readSchema_(ss.getSheetByName('Riwayat Adendum'), ENGINE.ADD),
    register: readSchema_(ss.getSheetByName('Monitoring Invoice'), ENGINE.REG, ENGINE.REG_ALIAS),
    log: readSchema_(ss.getSheetByName('Log Penerimaan'), ENGINE.LOG),
    tagbrut: readSchema_(it, ENGINE.TAG), notes: readSchema_(it, ENGINE.NOTE),
    approvals: readSchema_(ss.getSheetByName('Log Persetujuan'), ENGINE.APPR),
    surat: readSchema_(ss.getSheetByName(ENGINE.SURAT.sheet), ENGINE.SURAT),
    handover: readSchema_(ss.getSheetByName(ENGINE.HO.sheet), ENGINE.HO),
    lk: readLK_(), selectors: sel,
    memo: ss.getSheetByName(ENGINE.MEMO.sheet) ? readSchema_(ss.getSheetByName(ENGINE.MEMO.sheet), ENGINE.MEMO) : undefined,
    peref: readSchema_(ss.getSheetByName(ENGINE.PEREF.sheet), ENGINE.PEREF),
    arsipKeys: readArsipKeys_()
  };
}
function toCell_(v, type) {
  if (type === 'date' && typeof v === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(v)) { var p = v.split('-'); return dateAt_(+p[0], +p[1], +p[2]); }
  return v === null || v === undefined ? '' : v;
}

// =====================================================================  HITUNG & TULIS
function rebuildAll() { arRequireAccess_(); return rebuildAll_({ format: true }); }
/** inti pembaruan tanpa cek identitas — hanya dipanggil fungsi yang sudah memeriksa akses atau oleh trigger terverifikasi.
 *  opt.format: rapikan format sheet sumber (berat; dilewati untuk pembaruan karena edit sel) */
function rebuildAll_(opt) {
  opt = opt || {};
  var lock = arLock_();
  lock.waitLock(28000);
  try {
    var ss = arSpreadsheet_();
    if (!ss.getSheetByName('Kinerja & Peringatan Dini')) { try { var rr = ss.getSheetByName('RA-RI Termin'); setupKinerja_(ss, rr ? rr.getIndex() : ss.getSheets().length); } catch (err) { } }
    var input = readAll_();
    try { arAuditCounts_(input); } catch (err) { }
    var res = ENGINE.compute(input);
    arEnrichReports_(res);
    try { var cj = JSON.stringify(arDashboardCharts_(res)); if (cj.length < 8500) arProps_().setProperty('DASH_CHARTS', cj); } catch (err) { }
    try { arProps_().setProperty('REPORT_TIME', arNow_()); } catch (err) { }
    if (res.ctx.memo && res.ctx.memo.changed) writeMemo_(ss.getSheetByName(ENGINE.MEMO.sheet), res.ctx.memo.rows);
    writeAutoColumns_(ss.getSheetByName('Monitoring Invoice'), ENGINE.REG, ENGINE.REG_ALIAS, res.regOut);
    writeAutoColumns_(ss.getSheetByName('Log Penerimaan'), ENGINE.LOG, null, res.logOut);
    try { arHideArsip_(ss, {}); } catch (err) { } // versi sebelumnya menyembunyikan baris ARSIP: tampilkan kembali (kini dipindah lewat menu)
    REPORT_SHEETS.forEach(function (n) { var sh = ss.getSheetByName(n); if (sh && res.reports[n]) writeReport_(sh, res.reports[n], n); });
    REPORT_SHEETS.forEach(function (n) { var sh = ss.getSheetByName(n), rp = res.reports[n]; if (sh && rp && rp.charts) { try { ensureCharts_(sh, rp.charts); } catch (err) { } } });
    ensureSelectors_(ss, res);
    arWriteReconciliation_(res);
    try { arExtendFilters_(ss); } catch (err) { }
    if (opt.format) arFormatSourceSheets_(ss);
    else { try { if (arProps_().getProperty('REG_GROUP_V285') !== '1') { arGroupRegister_(ss); arProps_().setProperty('REG_GROUP_V285', '1'); } } catch (err) { } }
    // v2.8.7: kolom ADJUSTMENT DPP (DASAR PPN SAJA) di Monitoring Invoice — ditambahkan sekali (di ujung kanan)
    try { if (arProps_().getProperty('REG_HO_V287') !== '1') { var mi7 = ss.getSheetByName('Monitoring Invoice'), lp7 = ss.getSheetByName('Log Penerimaan'); if (mi7) ensureSchema_(mi7, ENGINE.REG, ENGINE.REG_ALIAS, 6); if (lp7) ensureSchema_(lp7, ENGINE.LOG, null, 4); arProps_().setProperty('REG_HO_V287', '1'); } } catch (err) { }
    var aktif = false; try { aktif = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'onEditInstalled'; }); } catch (err) { aktif = true; }
    var stamp = 'Diperbarui otomatis: ' + Utilities.formatDate(new Date(), TZ, 'dd MMM yyyy HH:mm') + ' · Sistem AR CSU v' + ENGINE.VERSION + ' · dirancang & dikembangkan oleh ' + PENGEMBANG +
      (aktif ? '' : '   ⚠ OTOMATISASI BELUM AKTIF — jalankan menu AR CSU ▸ Aktifkan otomatisasi agar pilihan bulan/NPP langsung memperbarui laporan');
    REPORT_SHEETS.forEach(function (n) { var sh = ss.getSheetByName(n); if (sh) sh.getRange('A3').setValue(stamp); });
    // v2.7.1: keterangan baris 2 laporan yang berubah istilah — ditulis ulang bila berbeda (aman dijalankan berulang)
    Object.keys(REPORT_SUB_).forEach(function (n) { var sh = ss.getSheetByName(n); if (sh && String(sh.getRange('A2').getValue()) !== REPORT_SUB_[n]) sh.getRange('A2').setValue(REPORT_SUB_[n]); });
    ss.getSheetByName('Monitoring Invoice').getRange('A4').setValue(stamp);
    ss.getSheetByName('Log Penerimaan').getRange('A3').setValue(stamp);
    SpreadsheetApp.flush();
    LAST_RES_ = res;
    return { ok: true, issues: res.issues.length };
  } finally { lock.releaseLock(); }
}
var LAST_RES_ = null;
var REPORT_SUB_ = {
  'Pareto': 'Daftar tagihan yang harus cair. Usulan AR, rencana tgl cair & catatan diisi di aplikasi ▸ Otorisasi Dokumen AR ▸ Lembar Validasi Rencana Cair. Kolom 🔒 tidak diketik manual; kolom VALIDASI SECTION HEAD F&HR 🔒 terisi otomatis saat Section Head F&HR menyetujui lembar.',
  'RA-RI Termin': 'RA = invoice yang divalidasi ✓ CAIR oleh Section Head F&HR di Lembar Validasi (RA awal = lembar pertama yang disetujui, RA final = revisi terakhir); RI = penerimaan di bulan itu. Satu tabel: ✓ RA · ➖ DIKELUARKAN · ◇ CADANGAN · ✚ DI LUAR RA, dengan Pencapaian RA dan Total RI ÷ RA.'
};
function compute_() { return ENGINE.compute(readAll_()); }
/** hitungan tanpa laporan (cepat) — untuk aplikasi; mode 'memo' ikut menghitung memorial */
function computeLight_(mode) { var i = readAll_(); i.light = mode || true; return ENGINE.compute(i); }

/** dropdown pemilih (bulan / NPP) di sel C5 */
function ensureSelectors_(ss, res) {
  var par = ss.getSheetByName('Parameter'), mp = ss.getSheetByName('Master Proyek');
  // Kartu Piutang: hanya proyek AR yang belum SELESAI (proyek selesai tetap bisa diketik manual)
  if (res && res.ctx) {
    var act = ENGINE.proyekAR(res.ctx).filter(function (n) { var p = res.ctx.M.pro[n]; return !p.selesai && p.aktif !== 'TIDAK'; }).sort();
    var key = act.join('|'), dp = arProps_();
    var kps = ss.getSheetByName('Kartu Piutang');
    if (kps && act.length && dp.getProperty('KP_LIST') !== key) {
      kps.getRange('C5').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(act, true).setAllowInvalid(true).setHelpText('Proyek aktif. Proyek SELESAI disembunyikan dari daftar (tetap bisa diketik NPP-nya).').build());
      dp.setProperty('KP_LIST', key);
    }
  }
  SELECTOR_SHEETS.forEach(function (n) {
    var sh = ss.getSheetByName(n); if (!sh) return;
    var cell = sh.getRange('C5'), cur = cell.getDataValidation();
    var ok = cur && cur.getAllowInvalid() && (n === 'Kartu Piutang' ? !!(res && res.ctx) : cur.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE);
    if (!ok) {
      var src = n === 'Kartu Piutang' ? mp.getRange('A3:A400') : par.getRange('M4:M60');
      cell.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(src, true).setAllowInvalid(true)
        .setHelpText(n === 'Kartu Piutang' ? 'Pilih NPP dari daftar.' : 'Pilih bulan dari daftar (klik ▼). Bila diketik, tulis seperti: Desember 2026').build());
    }
    if (cell.getNumberFormat() !== '@') { var v = cell.getValue(); cell.setNumberFormat('@'); if (isDate_(v)) { var ds = dstr_(v).split('-'); cell.setValue(ENGINE.MONTHS[+ds[1] - 1] + ' ' + ds[0]); } }
  });
}

function writeAutoColumns_(sh, schema, aliases, out) {
  if (!sh || !out.length) return;
  var map = colMap_(sh, schema, aliases), cols = schema.cols, list = [];
  cols.forEach(function (c) { if (colType_(c) === 'auto' && map[c[0]]) list.push({ key: c[0], col: map[c[0]], fmt: colFmt_(c) }); });
  list.sort(function (a, b) { return a.col - b.col; });
  var runs = [], cur = null;
  list.forEach(function (x) { if (cur && x.col === cur.end + 1) { cur.items.push(x); cur.end = x.col; } else { cur = { start: x.col, end: x.col, items: [x] }; runs.push(cur); } });
  runs.forEach(function (r) {
    var vals = out.map(function (o) { return r.items.map(function (x) { return o ? toCell_(o[x.key], x.fmt) : ''; }); });
    sh.getRange(schema.firstRow, r.start, vals.length, r.items.length).setValues(vals);
  });
}

/** menyembunyikan baris (rows per sheet); baris yang disembunyikan pada pembaruan sebelumnya tetapi tidak ada di rows ditampilkan kembali.
 *  Dipanggil dengan {} agar baris ARSIP yang dulu disembunyikan muncul lagi (baris ARSIP kini dipindah ke sheet Arsip Data Lama). */
function arHideArsip_(ss, rows) {
  if (!rows) return;
  var pr = arProps_(), prev = {};
  try { prev = JSON.parse(pr.getProperty('ARSIP_ROWS') || '{}'); } catch (e) { prev = {}; }
  var sheets = { reg: 'Monitoring Invoice', log: 'Log Penerimaan', tag: ENGINE.TAG.sheet }, next = {};
  Object.keys(sheets).forEach(function (k) {
    var sh = ss.getSheetByName(sheets[k]); if (!sh) return;
    var now = (rows[k] || []).slice().sort(function (a, b) { return a - b; }), was = prev[k] || [], on = {}, max = sh.getMaxRows();
    now.forEach(function (r) { on[r] = 1; });
    was.forEach(function (r) { if (!on[r] && r <= max) sh.showRows(r); });
    for (var i = 0; i < now.length;) { var j = i; while (j + 1 < now.length && now[j + 1] === now[j] + 1) j++; sh.hideRows(now[i], now[j] - now[i] + 1); i = j + 1; }
    next[k] = now;
  });
  var js = JSON.stringify(next);
  if (js !== JSON.stringify(prev) && js.length < 8500) pr.setProperty('ARSIP_ROWS', js);
}

/** tampilan teks sel (perkiraan jumlah karakter) untuk menghitung lebar kolom */
function cellLen_(v, f) {
  if (v === '' || v === null || v === undefined) return 0;
  if (typeof v === 'number') { if (f === 'money' || f === 'money0') return ENGINE.fmtRp(v).length + (v < 0 ? 2 : 0); if (f === 'pct') return 7; if (f === 'num2') return ENGINE.fmtRp(v).length + 3; if (f === 'date') return 11; return String(Math.round(v)).length; }
  if (f === 'date' && /^\d{4}-\d{2}-\d{2}$/.test(v)) return 11;
  return String(v).length;
}
/** posisi kolom fisik untuk kolom logis j (kolom pertama blok ber-"span" menempati beberapa kolom yang digabung) */
function physCol_(b, j) { return b.c + (j === 0 || !(b.span > 1) ? j : j + b.span - 1); }
function writeReport_(sh, rep, name, opt) {
  opt = opt || {};
  var from = rep.from, maxR = sh.getMaxRows(), maxC = sh.getMaxColumns(), need = from;
  var lw = function (b) { var w = 0; b.rows.forEach(function (r) { w = Math.max(w, r.length); }); return w; };
  rep.blocks.forEach(function (b) { need = Math.max(need, b.r + b.rows.length + 1); var w = lw(b), pw = w + (b.span > 1 && w ? b.span - 1 : 0); if (b.c + pw - 1 > maxC) { sh.insertColumnsAfter(maxC, b.c + pw - 1 - maxC); maxC = sh.getMaxColumns(); } });
  if (need > maxR) { sh.insertRowsAfter(maxR, need - maxR + 20); maxR = sh.getMaxRows(); }
  var clr = sh.getRange(from, 1, maxR - from + 1, maxC);
  try { clr.getMergedRanges().forEach(function (m) { m.breakApart(); }); } catch (e) { try { clr.breakApart(); } catch (e2) { } }
  try { clr.shiftRowGroupDepth(-1); } catch (e) { }
  clr.clear(); clr.clearDataValidations();
  clr.setFontFamily('Arial').setFontSize(9).setFontWeight('normal').setFontColor('#000000').setWrap(false).setVerticalAlignment('middle').setHorizontalAlignment('center');
  var meas = {}, lastRow = from;
  rep.blocks.forEach(function (b) {
    var w = lw(b), span = b.span > 1 ? b.span : 1, pw = w + (span > 1 && w ? span - 1 : 0);
    if (!w) return;
    var vals = [], fmts = [], skip = {};
    b.styles.forEach(function (st) { if (st.s === 'title' || st.s === 'note') for (var q = 0; q < (st.n || 1); q++) skip[st.r + q] = 1; });
    var hdrRows = {}; b.styles.forEach(function (st) { if (st.s === 'hdr') for (var q = 0; q < (st.n || 1); q++) hdrRows[st.r + q] = 1; });
    var cfm = {}; (b.cellFmt || []).forEach(function (x) { cfm[x[0] + '|' + x[1]] = x[2]; });
    b.rows.forEach(function (r, ri) {
      var vr = [], fr = [];
      for (var j = 0; j < w; j++) {
        var f = cfm[ri + '|' + j] || b.fmt[j] || '', v = j < r.length ? r[j] : '', cell = toCell_(v, f), nf = NUMF[f] && (typeof v === 'number' || f === 'date') ? NUMF[f] : 'General';
        vr.push(cell); fr.push(nf);
        if (j === 0 && span > 1) { for (var k = 1; k < span; k++) { vr.push(''); fr.push('General'); } }
      }
      vals.push(vr); fmts.push(fr);
      // ukur lebar (judul, catatan, dan baris satu-sel panjang dilewati)
      var filled = r.filter(function (x) { return x !== '' && x !== null && x !== undefined; }).length;
      if (b.noWidth || skip[ri] || (filled === 1 && cellLen_(r[0], '') > 30)) return;
      for (var j2 = 0; j2 < r.length; j2++) {
        if (j2 === 0 && span > 1) continue;
        var pc = physCol_(b, j2), len = cellLen_(r[j2], b.fmt[j2] || ''), nd;
        if (!len) continue;
        if (hdrRows[ri]) { var words = String(r[j2]).split(/\s+/), mx = 0; words.forEach(function (x) { mx = Math.max(mx, x.length); }); nd = Math.min(20, Math.max(mx + 2, Math.ceil(len / 3) + 2)); }
        else nd = len + 2;
        meas[pc] = Math.max(meas[pc] || 0, nd);
      }
    });
    var rg = sh.getRange(b.r, b.c, vals.length, pw);
    rg.setNumberFormats(fmts); rg.setValues(vals);
    arAlignBlock_(sh, b, vals, pw, fmts);
    lastRow = Math.max(lastRow, b.r + vals.length - 1);
    if (span > 1) {
      // gabungkan kolom pertama per baris (judul/catatan tidak digabung agar teksnya tetap mengalir)
      var run0 = -1;
      for (var ri2 = 0; ri2 <= vals.length; ri2++) {
        var ok = ri2 < vals.length && !skip[ri2];
        if (ok && run0 < 0) run0 = ri2;
        if (!ok && run0 >= 0) { sh.getRange(b.r + run0, b.c, ri2 - run0, span).mergeAcross(); run0 = -1; }
      }
    }
    (b.merge || []).forEach(function (m) { sh.getRange(b.r + m[0], b.c + m[1], m[2], m[3]).merge(); });
    b.styles.forEach(function (s) {
      var x = sh.getRange(b.r + s.r, b.c, s.n || 1, pw);
      if (s.s === 'hdr') x.setBackground(C.navy).setFontColor('#FFFFFF').setFontWeight('bold').setWrap(true).setHorizontalAlignment('center').setVerticalAlignment('middle').setBorder(true, true, true, true, true, true, '#A6A6A6', SpreadsheetApp.BorderStyle.SOLID);
      else if (s.s === 'total') x.setBackground(C.tot).setFontWeight('bold').setBorder(true, null, null, null, null, null, C.navy, SpreadsheetApp.BorderStyle.SOLID);
      else if (s.s === 'title') sh.getRange(b.r + s.r, b.c).setFontWeight('bold').setFontSize(10).setFontColor(C.navy);
      else if (s.s === 'bold') x.setFontWeight('bold');
      else if (s.s === 'band') x.setBackground(C.band).setFontWeight('bold').setBorder(true, null, null, null, null, null, '#BFC8D4', SpreadsheetApp.BorderStyle.SOLID);
      else if (s.s === 'lbl') { sh.getRange(b.r + s.r, b.c, s.n || 1, 1).setFontColor('#4A5A72'); sh.getRange(b.r + s.r, b.c + 3, s.n || 1, 1).setFontColor('#4A5A72'); sh.getRange(b.r + s.r, b.c + 7, s.n || 1, 1).setFontColor('#4A5A72'); sh.getRange(b.r + s.r, b.c + 1, s.n || 1, 1).setFontWeight('bold'); sh.getRange(b.r + s.r, b.c + 5, s.n || 1, 1).setFontWeight('bold'); sh.getRange(b.r + s.r, b.c + 8, s.n || 1, 1).setFontWeight('bold'); }
      else if (s.s === 'note') sh.getRange(b.r + s.r, b.c).setFontStyle('italic').setFontColor('#4A5A72');
      else if (s.s === 'warn') x.setFontColor('#9C6500').setFontWeight('bold');
      else if (s.s === 'alert') sh.getRange(b.r + s.r, b.c, 1, Math.max(pw, 8)).setBackground('#FDECEA').setFontColor('#B42318').setFontWeight('bold');
      else if (s.s === 'sisa') x.setFontStyle('italic').setFontWeight('bold').setBorder(null, null, true, null, null, null, '#8EA0BC', SpreadsheetApp.BorderStyle.SOLID);
      else if (s.s === 'stamp') x.setFontColor('#2F7D4F').setFontWeight('bold').setFontSize(8);
      else if (s.s === 'signname') x.setFontWeight('bold').setFontLine('underline');
      else if (s.s === 'kpi') x.setBackground(C.kpi).setFontWeight('bold').setFontColor(C.navy).setFontSize(10).setBorder(true, true, true, true, true, true, '#A6A6A6', SpreadsheetApp.BorderStyle.SOLID);
      else if (s.s === 'tileL') x.setFontSize(8).setFontWeight('bold').setFontColor('#52514E').setVerticalAlignment('middle');
      else if (s.s === 'tileV') x.setFontSize(18).setFontWeight('bold').setFontColor(C.navy).setHorizontalAlignment('center');
      else if (s.s === 'tileS') x.setFontSize(8).setFontColor('#52514E');
      else if (s.s === 'tileF') { x.setFontSize(9).setFontWeight('bold'); (b.flags || []).forEach(function (fl, ix) { var col = b.c + [3, 9, 12][ix]; sh.getRange(b.r + s.r, col, 1, 3).setFontColor(fl === 'R' ? '#B42318' : fl === 'G' ? '#0A5A0A' : '#7A5200'); }); }
    });
    if (b.merge && b.merge.length && /tile/.test(JSON.stringify(b.styles))) {
      for (var t = 0; t < 5; t++) sh.getRange(b.r, b.c + t * 3, b.rows.length, 3).setBackground('#FBFCFE').setBorder(true, true, true, true, null, null, '#D6DBE3', SpreadsheetApp.BorderStyle.SOLID);
    }
    (b.levels || []).forEach(function (lv) { sh.getRange(b.r + lv[0], b.c).setFontColor(lv[1] === 'MERAH' ? '#B42318' : '#9C6500').setFontWeight('bold'); });
    if (b.group && vals.length > 2) { try { var gr = sh.getRange(b.r + 1, 1, vals.length - 1, 1); gr.shiftRowGroupDepth(1); var g = sh.getRowGroup(b.r + 1, 1); if (g) g.collapse(); } catch (e) { } }
    if (b.edit && b.edit.lastDataOff >= b.edit.firstDataOff) {
      var r1 = b.r + b.edit.firstDataOff, n = b.edit.lastDataOff - b.edit.firstDataOff + 1, cols = {};
      Object.keys(b.edit.cols).forEach(function (k) {
        var abs = physCol_(b, +k); cols[String(abs)] = b.edit.cols[k];
        var er = sh.getRange(r1, abs, n, 1).setBackground(C.input);
        if (b.edit.cols[k] === 'validasi') er.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(ENGINE.OPT.validasi, true).setAllowInvalid(true).build());
        if (b.edit.cols[k] === 'rencana') er.setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).build());
      });
      var locks = []; (b.edit.lock || []).forEach(function (k) { var abs = physCol_(b, +k); sh.getRange(r1, abs, n, 1).setBackground('#EEF0F3'); locks.push(abs); });
      arProps_().setProperty('EDIT_' + name, JSON.stringify({ first: r1, last: r1 + n - 1, keyCol: physCol_(b, b.edit.keyCol), cols: cols, lock: locks }));
    }
  });
  // lebar kolom: ditentukan isi tabel (bukan judul), disimpan agar tidak diatur ulang bila tidak berubah
  var W = {};
  if (rep.widths) Object.keys(rep.widths).forEach(function (c) { W[c] = rep.widths[c]; });
  else Object.keys(meas).forEach(function (c) { W[c] = Math.max(4, Math.min(50, meas[c])); });
  Object.keys(rep.minW || {}).forEach(function (c) { W[c] = Math.max(W[c] || 0, rep.minW[c]); });
  var key = 'W_' + (opt.key ? opt.key + '_' + sh.getSheetId() : name), dp = arProps_(), wj = JSON.stringify(W);
  if (dp.getProperty(key) !== wj) { Object.keys(W).forEach(function (c) { sh.setColumnWidth(+c, Math.round(W[c] * 7.3 + 10)); }); dp.setProperty(key, wj); }
  var wrap = (rep.wrapCols || []).slice(); Object.keys(meas).forEach(function (c) { if (meas[c] > 50 && +c > 1 && wrap.indexOf(+c) < 0) wrap.push(+c); });
  wrap.forEach(function (c) { if (c >= 1 && lastRow >= from) sh.getRange(from, c, lastRow - from + 1, 1).setWrap(true); });
}

/** grafik bawaan Google Sheets (dibuat sekali, dicari dari judulnya; data selalu di sel yang sama sehingga grafik ikut berubah). Posisi disesuaikan bila tata letak berubah. */
function ensureCharts_(sh, specs) {
  var have = {};
  sh.getCharts().forEach(function (c) { try { have[String(c.getOptions().get('title'))] = c; } catch (e) { } });
  specs.forEach(function (sp) {
    var ex = have[sp.title];
    var T = sp.type === 'BAR' ? Charts.ChartType.BAR : sp.type === 'LINE' ? Charts.ChartType.LINE : Charts.ChartType.COLUMN;
    var b = (ex ? ex.modify().clearRanges() : sh.newChart()).setChartType(T)
      .addRange(sh.getRange(sp.range[0], sp.range[1], sp.range[2], sp.range[3])).setNumHeaders(1)
      .setPosition(sp.row, sp.col, 4, 4)
      .setOption('title', sp.title).setOption('width', sp.w).setOption('height', sp.h).setOption('colors', sp.colors)
      .setOption('legend', { position: sp.legend || 'none', textStyle: { color: '#52514E', fontSize: 10 } })
      .setOption('titleTextStyle', { color: '#1F3864', fontSize: 12, bold: true }).setOption('backgroundColor', '#FFFFFF')
      .setOption(sp.type === 'BAR' ? 'hAxis' : 'vAxis', { format: sp.type === 'LINE' ? '#,##0' : '#,##0.0', gridlines: { color: '#E1E0D9' }, textStyle: { color: '#898781', fontSize: 9 }, minValue: 0 })
      .setOption(sp.type === 'BAR' ? 'vAxis' : 'hAxis', { textStyle: { color: '#52514E', fontSize: 9 } });
    if (sp.type === 'LINE') b = b.setOption('lineWidth', 2.5).setOption('pointSize', 5);
    else b = b.setOption('isStacked', !!sp.stacked).setOption('bar', { groupWidth: '62%' });
    if (sp.labels) b = b.setOption('series', { 0: { dataLabel: 'value', dataLabelPlacement: 'outsideEnd' } });
    if (ex) sh.updateChart(b.build()); else sh.insertChart(b.build());
  });
}

function setRegisterValue_(nomor, key, value) {
  var sh = sheet_('Monitoring Invoice'), map = colMap_(sh, ENGINE.REG, ENGINE.REG_ALIAS);
  var row = findRegisterRow_(sh, nomor, map);
  if (row && map[key]) sh.getRange(row, map[key]).setValue(value);
}
function findRegisterRow_(sh, nomor, map) {
  map = map || colMap_(sh, ENGINE.REG, ENGINE.REG_ALIAS);
  var ni = map.nomor, last = sh.getLastRow(), first = ENGINE.REG.firstRow;
  if (!ni || last < first) return 0;
  var vals = sh.getRange(first, ni, last - first + 1, 1).getValues(), key = ENGINE.normInv(nomor);
  for (var i = 0; i < vals.length; i++) if (String(vals[i][0]).trim() && ENGINE.normInv(vals[i][0]) === key) return first + i;
  return 0;
}
/** pastikan baris tujuan ada: bila sheet sudah habis barisnya, tambah 50 baris kosong (format mengikuti baris terakhir) */
function roomFor_(sh, row) { var mx = sh.getMaxRows(); if (row > mx) sh.insertRowsAfter(mx, row - mx + 50); }
function lastDataRow_(sh, firstRow, cols) {
  var last = sh.getLastRow(); if (last < firstRow) return firstRow - 1;
  var vals = sh.getRange(firstRow, 1, last - firstRow + 1, sh.getLastColumn()).getValues();
  for (var k = vals.length - 1; k >= 0; k--) if (cols.some(function (c) { return c && String(vals[k][c - 1]).trim(); })) return firstRow + k;
  return firstRow - 1;
}

// =====================================================================  MENU: ZONA WAKTU & PEMBARUAN v2.1
/** kolom bertipe tanggal per sheet input (berdasarkan skema) */
function dateTargets_(ss) {
  var out = [];
  [['Monitoring Invoice', ENGINE.REG, ENGINE.REG_ALIAS], ['Log Penerimaan', ENGINE.LOG, null], ['Master Proyek', ENGINE.MPRO, ENGINE.MPRO_ALIAS],
    ['Riwayat Adendum', ENGINE.ADD, null], ['Input Tagbrut', ENGINE.TAG, null], ['Input Tagbrut', ENGINE.NOTE, null]].forEach(function (x) {
    var sh = ss.getSheetByName(x[0]); if (!sh) return;
    var map = colMap_(sh, x[1], x[2]), last = sh.getLastRow(); if (last < x[1].firstRow) return;
    x[1].cols.forEach(function (c) { if (colFmt_(c) === 'date' && colType_(c) !== 'auto' && map[c[0]]) out.push({ sh: sh, r: x[1].firstRow, c: map[c[0]], n: last - x[1].firstRow + 1, w: 1 }); });
  });
  ['Parameter', 'Data LK (Tempel)'].forEach(function (n) {
    var sh = ss.getSheetByName(n); if (!sh || sh.getLastRow() < 1 || sh.getLastColumn() < 1) return;
    out.push({ sh: sh, r: 1, c: 1, n: sh.getLastRow(), w: sh.getLastColumn() });
  });
  return out;
}
/**
 * Memindahkan file ke zona waktu Asia/Jakarta dan membetulkan tanggal yang bergeser.
 * Tanggal yang diketik manual tersimpan pukul 00.00 → tidak diubah.
 * Tanggal yang ditulis aplikasi saat zona waktu file berbeda tersimpan dengan jam (mis. 10.00 hari sebelumnya) → dikembalikan ke tanggal WIB yang benar.
 */
function fixTimezone_() {
  var ss = arSpreadsheet_(), tz0 = ss.getSpreadsheetTimeZone(), J = 'Asia/Jakarta', T = dateTargets_(ss), fix = [];
  T.forEach(function (t) {
    var rg = t.sh.getRange(t.r, t.c, t.n, t.w), v = rg.getValues(), f = rg.getFormulas(); t.cells = [];
    for (var i = 0; i < t.n; i++) for (var j = 0; j < t.w; j++) {
      var x = v[i][j]; if (!isDate_(x) || f[i][j]) continue;
      var hm = Utilities.formatDate(x, tz0, 'HH:mm'), want;
      if (hm === '00:00') want = dstr_(x, tz0);                       // diketik manual → tanggal yang tampil
      else if (tz0 !== J) want = dstr_(x, J);                          // ditulis aplikasi → tanggal WIB
      else { var h = +hm.substr(0, 2), d = dstr_(x, tz0).split('-');   // sudah WIB tapi masih membawa jam sisa pergeseran
        var dt = new Date(Date.UTC(+d[0], +d[1] - 1, +d[2] + (h >= 5 ? 1 : 0))); want = Utilities.formatDate(dt, 'UTC', 'yyyy-MM-dd'); }
      t.cells.push({ i: i, j: j, want: want, was: Utilities.formatDate(x, tz0, 'dd-MM-yyyy HH:mm') });
    }
  });
  if (tz0 !== J) { ss.setSpreadsheetTimeZone(J); SpreadsheetApp.flush(); }
  SS_TZ_ = J;
  // tulis ulang hanya sel yang tampilannya belum sesuai (tanggal beda atau masih membawa jam)
  T.forEach(function (t) {
    if (!t.cells.length) return;
    var v = t.sh.getRange(t.r, t.c, t.n, t.w).getValues();
    t.cells.forEach(function (x) {
      var cur = v[x.i][x.j];
      if (isDate_(cur) && Utilities.formatDate(cur, J, 'yyyy-MM-dd HH:mm') === x.want + ' 00:00') return;
      var p = x.want.split('-'), cell = t.sh.getRange(t.r + x.i, t.c + x.j);
      cell.setValue(dateAt_(+p[0], +p[1], +p[2], J));
      fix.push(t.sh.getName() + ' ' + cell.getA1Notation() + ': ' + x.was + ' → ' + p[2] + '-' + p[1] + '-' + p[0]);
    });
  });
  return { tz0: tz0, fixed: fix };
}
function fixTimezoneMenu() { arRequireMaker_();
  var ui = SpreadsheetApp.getUi();
  if (ui.alert('Perbaiki zona waktu & tanggal', 'Zona waktu file akan diubah ke (GMT+07:00) Jakarta dan tanggal yang bergeser karena zona waktu akan dibetulkan. Tanggal yang Anda ketik manual tidak diubah. Lanjutkan?', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  var r = fixTimezone_(); rebuildAll();
  ui.alert(tzReport_(r));
}
function tzReport_(r) {
  return 'Zona waktu file: ' + r.tz0 + (r.tz0 !== 'Asia/Jakarta' ? ' → Asia/Jakarta' : ' (sudah benar)') + '\n' +
    (r.fixed.length ? r.fixed.length + ' tanggal dibetulkan:\n' + r.fixed.slice(0, 25).join('\n') + (r.fixed.length > 25 ? '\n… dan ' + (r.fixed.length - 25) + ' lainnya' : '') : 'Tidak ada tanggal yang perlu dibetulkan.');
}
/** langkah v2.1: kolom TOP RETENSI (Fajarputera 100 hari) + TOP invoice retensi Fajarputera yang sudah ada */
function v21Steps_(ss, log) {
  var mp = ss.getSheetByName('Master Pelanggan');
  if (mp) {
    var n = ensureSchema_(mp, ENGINE.MPEL, null, null); if (n) log.push('Master Pelanggan: +' + n + ' kolom');
    var map = colMap_(mp, ENGINE.MPEL), last = mp.getLastRow();
    if (map.top_ret && last >= ENGINE.MPEL.firstRow) {
      var nm = mp.getRange(ENGINE.MPEL.firstRow, map.nama, last - ENGINE.MPEL.firstRow + 1, 1).getValues();
      nm.forEach(function (r, i) { if (/FAJAR\s*PUTERA/.test(norm_(r[0]))) { var c = mp.getRange(ENGINE.MPEL.firstRow + i, map.top_ret); if (c.getValue() === '') { c.setValue(100); log.push('TOP retensi ' + r[0] + ' = 100 hari'); } } });
    }
  }
  var mi = ss.getSheetByName('Monitoring Invoice'), rm = colMap_(mi, ENGINE.REG, ENGINE.REG_ALIAS), lr = mi.getLastRow(), f0 = ENGINE.REG.firstRow;
  if (rm.top && rm.pelanggan && lr >= f0) {
    var V = mi.getRange(f0, 1, lr - f0 + 1, mi.getLastColumn()).getValues(), ch = [];
    V.forEach(function (r, i) {
      var isRet = norm_(rm.jenis ? r[rm.jenis - 1] : '') === 'RETENSI' || /RETENSI/.test(norm_(rm.invke ? r[rm.invke - 1] : ''));
      if (isRet && /FAJAR\s*PUTERA/.test(norm_(r[rm.pelanggan - 1])) && Number(r[rm.top - 1]) !== 100) { mi.getRange(f0 + i, rm.top).setValue(100); ch.push(String(r[rm.nomor - 1]).replace('KU.02.02/WB-4C.', '')); }
    });
    if (ch.length) log.push('TOP invoice retensi Fajarputera → 100 hari: ' + ch.join(', '));
  }
}
function applyV21() { arRequireMaker_(); // dipertahankan untuk kompatibilitas menu lama
  var ui = SpreadsheetApp.getUi(), log = []; v21Steps_(arSpreadsheet_(), log); var r = fixTimezone_(); rebuildAll();
  ui.alert('Pembaruan v2.1 selesai', (log.length ? log.join('\n') + '\n\n' : '') + tzReport_(r), ui.ButtonSet.OK);
}

// =====================================================================  PEMBARUAN v2.2
var V22_AKUN = [
  ['BANK_BNI', 'Rekening penerimaan BNI 149394879', '', '', 'MOHON KONFIRMASI Akuntansi: GL rekening penerimaan termin'],
  ['BANK_MANDIRI', 'Rekening penerimaan Mandiri 0060097045854', '', '', 'MOHON KONFIRMASI Akuntansi: GL rekening penerimaan termin'],
  ['BANK_BSI', 'Rekening penerimaan BSI 7148973088', '', '', 'MOHON KONFIRMASI Akuntansi: GL rekening penerimaan termin'],
  ['BANK_BTN', 'Rekening penerimaan BTN 0024101300010876', '', '', 'MOHON KONFIRMASI Akuntansi: GL rekening penerimaan termin']];
var V22_REK = [['BNI', '149394879', 'BANK_BNI'], ['MANDIRI', '0060097045854', 'BANK_MANDIRI'], ['BSI', '7148973088', 'BANK_BSI'], ['BTN', '0024101300010876', 'BANK_BTN']];
function invRow_(sh, map, nomor) { return findRegisterRow_(sh, nomor, map); }
function logRows_(sh, map) {
  var f = ENGINE.LOG.firstRow, last = sh.getLastRow(); if (last < f) return [];
  var v = sh.getRange(f, 1, last - f + 1, sh.getLastColumn()).getValues();
  return v.map(function (r, i) { var o = { _row: f + i }; ENGINE.LOG.cols.forEach(function (c) { o[c[0]] = map[c[0]] ? r[map[c[0]] - 1] : ''; }); return o; });
}
function appendLog_(sh, map, rows) {
  var t = lastDataRow_(sh, ENGINE.LOG.firstRow, [map.tgl, map.inv, map.nominal]) + 1;
  if (t + rows.length > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), rows.length + 20);
  rows.forEach(function (p, i) { ENGINE.LOG.cols.forEach(function (c) { var j = map[c[0]]; if (j && colType_(c) === 'manual' && p[c[0]] !== undefined && p[c[0]] !== '') sh.getRange(t + i, j).setValue(toCell_(p[c[0]], colFmt_(c))); }); });
}
/** daftar perbaikan data sekali jalan (dibaca dari kondisi file saat ini; setiap langkah dilewati bila kondisinya tidak cocok) */
function v22Fixes_(ss) {
  var res = compute_(), ctx = res.ctx, F = [], TOL = 5, K = function (n) { return ENGINE.normInv('KU.02.02/WB-4C.' + n); };
  var byNo = {}; ctx.invs.forEach(function (i) { byNo[ENGINE.normInv(i.nomor)] = i; });
  var lp = ss.getSheetByName('Log Penerimaan'), lm = colMap_(lp, ENGINE.LOG), mi = ss.getSheetByName('Monitoring Invoice'), mm = colMap_(mi, ENGINE.REG, ENGINE.REG_ALIAS);
  var pbN = 0; ctx.pays.forEach(function (p) { var m = String(p.ref).toUpperCase().match(/^PB-(\d+)$/); if (m) pbN = Math.max(pbN, +m[1]); });
  var hasPB = function (nomor) { return ctx.pays.some(function (p) { return p._pb && p._inv && ENGINE.normInv(p._inv.nomor) === nomor; }); };
  var fmt = ENGINE.fmtRp;
  // (a) Barry Callebaut: retensi tercatat ganda + kompensasi PPh Termin II–IV
  var r107 = byNo[K('0107/2026')], d2 = byNo[K('0176/2025#2')];
  if (r107 && r107.npp === '254C 0569 JT') {
    var L = logRows_(lp, lm), asum = L.filter(function (r) { return ENGINE.normInv(r.inv) === K('0107/2026') && /DIASUMSIKAN/i.test(String(r.ket)) && Math.round(Number(r.nominal)) === r107._a.kas; });
    var pind = L.filter(function (r) { return ENGINE.normInv(r.inv) === K('0176/2025#2'); });
    var other = L.filter(function (r) { return ENGINE.normInv(r.inv) === K('0107/2026') && asum.indexOf(r) < 0; });
    var cash = [].concat(pind, other).reduce(function (t, r) { return t + (Number(r.nominal) || 0); }, 0);
    var terms = ['0092/2025', '0122/2025', '0176/2025'].map(function (n) { return byNo[K(n)]; }).filter(function (i) { return i && i._res.closed && (i._res.pot_pph || 0) <= TOL && i._a.pph > 0 && !hasPB(ENGINE.normInv(i.nomor)); });
    var kompen = terms.reduce(function (t, i) { return t + i._a.pph; }, 0);
    if ((asum.length || pind.length) && terms.length === 3 && Math.abs(r107._a.kas - cash - kompen) <= TOL) {
      F.push({ d: 'Barry Callebaut 0107/2026: hapus pembayaran asumsi migrasi Rp' + fmt(r107._a.kas) + (pind.length ? ', pindahkan ' + pind.length + ' baris pembayaran 0176/2025#2 (Rp' + fmt(pind.reduce(function (t, r) { return t + Number(r.nominal); }, 0)) + ') ke 0107/2026, 0176/2025#2 → BATAL' : '') +
          ', lalu kompensasi PPh tidak dipotong Termin II–IV Rp' + fmt(kompen) + ' (PB) ke 0107/2026',
        run: function () {
          var L2 = logRows_(lp, lm);
          L2.filter(function (r) { return ENGINE.normInv(r.inv) === K('0176/2025#2'); }).forEach(function (r) { lp.getRange(r._row, lm.inv).setValue(K('0107/2026')); if (lm.ket) lp.getRange(r._row, lm.ket).setValue(String(r.ket) + ' · dipindah dari 0176/2025#2 (v2.2)'); });
          L2.filter(function (r) { return ENGINE.normInv(r.inv) === K('0107/2026') && /DIASUMSIKAN/i.test(String(r.ket)) && Math.round(Number(r.nominal)) === r107._a.kas; }).reverse().forEach(function (r) { lp.deleteRow(r._row); });
          if (d2) { var rr = invRow_(mi, mm, d2.nomor); if (rr) { mi.getRange(rr, mm.status_inv).setValue('BATAL'); if (mm.catatan) mi.getRange(rr, mm.catatan).setValue('Retensi tercatat ganda saat migrasi — tagihan retensi yang sah: 0107/2026 (v2.2)'); } }
          var tgl = ''; L2.forEach(function (r) { var k = ENGINE.normInv(r.inv); if (k === K('0176/2025#2') || (k === K('0107/2026') && !/DIASUMSIKAN/i.test(String(r.ket)))) { var d = isDate_(r.tgl) ? dstr_(r.tgl) : String(r.tgl).substring(0, 10); if (d > tgl) tgl = d; } });
          tgl = tgl || r107.tgl;
          var rows = [];
          terms.forEach(function (i) { pbN++; var ref = 'PB-' + ('000' + pbN).slice(-4);
            rows.push({ tgl: tgl, ref: ref, inv: i.nomor, nominal: -i._a.pph, jenis_trx: 'PINDAH BUKU', ket: 'Kompensasi PPh tidak dipotong ' + i.invke + ' → retensi 0107/2026' });
            rows.push({ tgl: tgl, ref: ref, inv: r107.nomor, nominal: i._a.pph, jenis_trx: 'PINDAH BUKU', ket: 'Kompensasi PPh tidak dipotong ' + i.invke + ' (dari ' + i.nomor.replace('KU.02.02/WB-4C.', '') + ')' }); });
          appendLog_(lp, lm, rows);
        } });
    }
  }
  // (b) WRK 254C 0823 JT: PPh 0109/2025 tidak dipotong → kompensasi ke 0091/2025
  var w109 = byNo[K('0109/2025')], w91 = byNo[K('0091/2025')];
  if (w109 && w91 && w109._res.closed && (w109._res.pot_pph || 0) <= TOL && w109._a.pph > 0 && !hasPB(ENGINE.normInv(w109.nomor)) && /^KURANG BAYAR/.test(w91._res.status)) {
    F.push({ d: 'WRK: PPh 0109/2025 tidak dipotong (Rp' + fmt(w109._a.pph) + ') dikompensasi ke 0091/2025 (PB) — sisa 0091 menjadi Rp' + fmt(w91._res.sisa - w109._a.pph),
      run: function () { pbN++; var ref = 'PB-' + ('000' + pbN).slice(-4), t = w109._res.tgl_bayar || w109.tgl;
        appendLog_(lp, lm, [{ tgl: t, ref: ref, inv: w109.nomor, nominal: -w109._a.pph, jenis_trx: 'PINDAH BUKU', ket: 'PPh tidak dipotong pelanggan → kompensasi ke 0091/2025' },
          { tgl: t, ref: ref, inv: w91.nomor, nominal: w109._a.pph, jenis_trx: 'PINDAH BUKU', ket: 'Kompensasi PPh tidak dipotong 0109/2025' }]); } });
  }
  // (c) WRK: kekurangan karena pembayaran SCF ditagih via BA / memo
  var scf = ['0091/2025', '0131/2025', '0193/2025', '0014/2026'].map(function (n) { return byNo[K(n)]; }).filter(function (i) { return i && /^KURANG BAYAR/.test(i._res.status) && !String(i.ba_no || '').trim(); });
  if (scf.length && mm.ba_no) F.push({ d: 'WRK (SCF): ' + scf.map(function (i) { return i.nomor.replace('KU.02.02/WB-4C.', ''); }).join(', ') + ' → status KURANG BAYAR – DITAGIH VIA BA/MEMO (nomor BA/memo diisi kemudian)',
    run: function () { scf.forEach(function (i) { var rr = invRow_(mi, mm, i.nomor); if (rr) mi.getRange(rr, mm.ba_no).setValue('SCF – NO. BA/MEMO BELUM DIISI'); }); } });
  // (d) PT Mulia Boga Raya selesai
  var mp = ss.getSheetByName('Master Proyek'), pm = colMap_(mp, ENGINE.MPRO, ENGINE.MPRO_ALIAS), mb = ctx.M.pro['244C 1975 JT'];
  if (mb && mb.aktif !== 'SELESAI' && pm.aktif && !ctx.invs.some(function (i) { return i.npp === '244C 1975 JT' && !i._batal && !i._res.closed; }))
    F.push({ d: 'Master Proyek 244C 1975 JT (PT Mulia Boga Raya): AKTIF = SELESAI', run: function () { mp.getRange(mb._row, pm.aktif).setValue('SELESAI'); } });
  // (e) Input Tagbrut: hanya jasa — baris NPP di luar Master Proyek dihapus
  var it = ss.getSheetByName('Input Tagbrut');
  if (it && it.getLastRow() >= ENGINE.TAG.firstRow) {
    var tm = colMap_(it, ENGINE.TAG), n0 = it.getLastRow() - ENGINE.TAG.firstRow + 1, tv = it.getRange(ENGINE.TAG.firstRow, tm.npp, n0, 1).getValues(), drop = [];
    tv.forEach(function (r, i) { var n = String(r[0]).trim(); if (n && !ctx.M.pro[n]) drop.push({ row: ENGINE.TAG.firstRow + i, npp: n }); });
    if (drop.length) F.push({ d: 'Input Tagbrut: hapus ' + drop.length + ' baris NPP material / di luar Master Proyek (' + drop.map(function (x) { return x.npp; }).filter(function (v, i, a) { return a.indexOf(v) === i; }).join(', ') + ')',
      run: function () { var cols = ENGINE.TAG.cols.map(function (c) { return tm[c[0]]; }).filter(Boolean); drop.forEach(function (x) { cols.forEach(function (c) { it.getRange(x.row, c).setValue(''); }); }); } });
  }
  return F;
}
function applyV22() { arRequireMaker_();
  var ui = SpreadsheetApp.getUi(), ss = arSpreadsheet_(), log = [];
  if (ui.alert('Pembaruan v2.2', 'Langkah 1: struktur (kolom rekening, pindah buku, BA/memo, alasan tutup, saldo awal UM, sheet Monitoring UM & Retensi, posisi kartu piutang) + perbaikan zona waktu.\nLangkah 2: daftar perbaikan data ditampilkan dulu sebelum diterapkan.\n\nLanjutkan?', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  // --- struktur
  [['Monitoring Invoice', ENGINE.REG, ENGINE.REG_ALIAS, 6], ['Log Penerimaan', ENGINE.LOG, null, 4], ['Master Proyek', ENGINE.MPRO, ENGINE.MPRO_ALIAS, null]].forEach(function (x) {
    var sh = ss.getSheetByName(x[0]); if (!sh) return; var n = ensureSchema_(sh, x[1], x[2], x[3]); if (n) log.push(x[0] + ': +' + n + ' kolom');
  });
  v21Steps_(ss, log);
  var par = ss.getSheetByName('Parameter');
  if (par.getMaxColumns() < 34) par.insertColumnsAfter(par.getMaxColumns(), 34 - par.getMaxColumns());
  if (norm_(par.getRange('AF3').getValue()) !== 'BANK') {
    par.getRange('AF2').setValue('REKENING PENERIMAAN TERMIN').setFontWeight('bold');
    par.getRange('AF3:AH3').setValues([['BANK', 'NO. REKENING', 'KODE AKUN (PEMETAAN AKUN SAP)']]).setBackground(C.navy).setFontColor('#FFFFFF').setFontWeight('bold');
    par.getRange(4, 32, V22_REK.length, 3).setNumberFormat('@').setValues(V22_REK); log.push('Parameter: tabel Rekening Penerimaan (BNI, Mandiri, BSI, BTN)');
  }
  var ak = par.getRange('O4:O80').getValues().map(function (r) { return norm_(r[0]); }), lastAk = 3; ak.forEach(function (k, i) { if (k) lastAk = 4 + i; });
  V22_AKUN.forEach(function (r) { if (ak.indexOf(r[0]) < 0) { lastAk++; par.getRange(lastAk, 15, 1, 5).setNumberFormat('@').setValues([r]); log.push('Pemetaan akun: ' + r[0] + ' (GL menunggu konfirmasi Akuntansi)'); } });
  var kv = par.getRange('A4:A120').getValues().map(function (r) { return norm_(r[0]); }), lastKv = 3; kv.forEach(function (k, i) { if (k) lastKv = 4 + i; });
  if (kv.indexOf('REKENING_WAJIB_MULAI') < 0) { lastKv++; par.getRange(lastKv, 1, 1, 3).setNumberFormat('@').setValues([['REKENING_WAJIB_MULAI', Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'), 'Pembayaran sejak tanggal ini wajib mencantumkan REKENING PENERIMAAN']]); }
  // sheet Monitoring UM & Retensi (menggantikan Piutang Retensi — jumlah sheet tetap)
  var um = ss.getSheetByName('Monitoring UM & Retensi');
  if (!um) { var old = ss.getSheetByName('Piutang Retensi'); if (old) { old.setName('Monitoring UM & Retensi'); um = old; log.push('Sheet Piutang Retensi → Monitoring UM & Retensi'); } else { um = setupReport_(ss, 'Monitoring UM & Retensi', '', '', null, 9); log.push('Sheet Monitoring UM & Retensi dibuat'); } }
  um.getRange('A1').setValue('MONITORING UANG MUKA & RETENSI').setFontSize(14).setFontWeight('bold').setFontColor(C.navy);
  um.getRange('A2').setValue('Sisa uang muka yang masih boleh dipotong per proyek (aplikasi membatasi potongan UM sampai sisa ini) dan retensi yang ditahan vs sudah ditagih.').setFontSize(10);
  var kp = ss.getSheetByName('Kartu Piutang');
  if (kp) { kp.getRange('D5').setValue('POSISI PER ▸').setFontWeight('bold').setHorizontalAlignment('right'); kp.getRange('E5').setBackground('#FFE699').setFontWeight('bold').setNumberFormat(NUMF.date)
    .setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).setHelpText('Kosong = hari ini (bawaan). Tanggal lain hanya sementara: setiap pagi pukul 06.00 dikosongkan lagi.').build()); }
  // pilihan (dropdown)
  var lp = ss.getSheetByName('Log Penerimaan'), lm = colMap_(lp, ENGINE.LOG), nR = Math.max(lp.getMaxRows() - ENGINE.LOG.firstRow + 1, 1);
  if (lm.rekening) lp.getRange(ENGINE.LOG.firstRow, lm.rekening, nR, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(V22_REK.map(function (r) { return r[0] + ' ' + r[1]; }), true).setAllowInvalid(true).build());
  if (lm.jenis_trx) lp.getRange(ENGINE.LOG.firstRow, lm.jenis_trx, nR, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(ENGINE.OPT.jenis_trx, true).setAllowInvalid(true).setHelpText('Kosong = transfer biasa. PINDAH BUKU = antar invoice (baris − di invoice asal, + di invoice tujuan, REF sama). KOMPENSASI = piutang dihapus dengan utang ke pelanggan (REF = no. BA/memo; tidak masuk inkaso).').build());
  var mp = ss.getSheetByName('Master Proyek'), pm = colMap_(mp, ENGINE.MPRO, ENGINE.MPRO_ALIAS);
  if (pm.aktif) mp.getRange(ENGINE.MPRO.firstRow, pm.aktif, Math.max(mp.getMaxRows() - 2, 1), 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(ENGINE.OPT.aktif_pro, true).setAllowInvalid(true).setHelpText('SELESAI = proyek selesai & lunas (tidak muncul di aplikasi, tidak dicek kelengkapannya). ARSIP = selesai & lunas sebelum sistem dipakai: semua invoice, penerimaan & tagbrut-nya diabaikan sistem dan barisnya disembunyikan (tidak dihapus).').build());
  var it = ss.getSheetByName('Input Tagbrut');
  if (it) it.getRange(ENGINE.TAG.firstRow, 1, 500, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(mp.getRange('A3:A400'), true).setAllowInvalid(false).setHelpText('Hanya NPP proyek jasa dari Master Proyek.').build());
  // zona waktu
  var tz = fixTimezone_();
  rebuildAll();
  // --- perbaikan data (pratinjau dulu)
  var F = v22Fixes_(ss), done = [];
  if (F.length) {
    var msg = 'Perbaikan data berikut akan diterapkan:\n\n' + F.map(function (f, i) { return (i + 1) + '. ' + f.d; }).join('\n\n') + '\n\nTerapkan sekarang? (Batal = struktur tetap diperbarui, data tidak diubah)';
    if (ui.alert('Pembaruan v2.2 — perbaikan data', msg, ui.ButtonSet.OK_CANCEL) === ui.Button.OK) { F.forEach(function (f) { f.run(); done.push(f.d); }); rebuildAll(); }
  }
  var res = compute_(), st = function (n) { var k = ENGINE.normInv('KU.02.02/WB-4C.' + n), i = res.ctx.invs.filter(function (x) { return ENGINE.normInv(x.nomor) === k; })[0]; return i ? n + ': ' + i._res.status + (i._res.sisa ? ' (sisa ' + ENGINE.fmtRp(i._res.sisa) + ')' : '') : ''; };
  var chk = ['0107/2026', '0092/2025', '0122/2025', '0176/2025', '0109/2025', '0091/2025', '0131/2025', '0003/2026'].map(st).filter(String);
  ui.alert('Pembaruan v2.2 selesai', (log.length ? '• ' + log.join('\n• ') + '\n\n' : '') + tzReport_(tz) + '\n\n' + (done.length ? 'Perbaikan data diterapkan:\n• ' + done.join('\n• ') : 'Tidak ada perbaikan data yang diterapkan.') +
    '\n\nStatus setelah pembaruan:\n' + chk.join('\n') + '\n\nLangkah berikut: terbitkan ulang aplikasi web (Kelola deployment ▸ Versi baru), lalu cek sheet Cek Data.', ui.ButtonSet.OK);
}

// =====================================================================  PENGAMANAN & KEPEMILIKAN
var LOCK_TAG = 'AR CSU · dikunci';
function aboutSystem() {
  SpreadsheetApp.getUi().alert('Sistem AR CSU v' + ENGINE.VERSION, 'Sistem Account Receivable Concrete Services Unit — invoice, kwitansi, BAP, penerimaan, kartu piutang, persetujuan digital, usulan jurnal SAP, CKPN, dan rekon LK.\n\nDirancang dan dikembangkan oleh ' + PENGEMBANG + ' (AR Officer, Concrete Services Unit), 2026.\nPenggandaan sistem, kode, maupun format laporan tanpa izin pengembang tidak diperkenankan.', SpreadsheetApp.getUi().ButtonSet.OK);
}
/** pengguna yang boleh mengubah struktur: pemilik file + PEMBUAT & PENYETUJU di Parameter (perlu agar aplikasi bisa menulis) */
function lockEditors_(ss) {
  var res = compute_(), out = [];
  try { var o = ss.getOwner(); if (o && o.getEmail()) out.push(o.getEmail().toLowerCase()); } catch (e) { }
  try { var me = Session.getEffectiveUser().getEmail(); if (me) out.push(me.toLowerCase()); } catch (e) { }
  res.params.users.forEach(function (u) { if (u.email && /^(PEMBUAT|PENYETUJU)/.test(u.peran)) out.push(u.email); });
  var dl = String(res.params.DELEGASI_EMAIL || '').toLowerCase(); if (dl && ENGINE.delegasiFor(res.params, dl, null, res.ctx.today)) out.push(dl); // delegasi aktif perlu menulis Log Persetujuan
  return out.filter(function (v, i, a) { return v && a.indexOf(v) === i; });
}
function protectRange_(target, editors, desc) {
  var p = target.protect().setDescription(LOCK_TAG + ' — ' + desc);
  try { p.removeEditors(p.getEditors()); } catch (e) { }
  try { p.addEditors(editors); } catch (e) { }
  try { if (p.canDomainEdit()) p.setDomainEdit(false); } catch (e) { }
  return p;
}
function lockSystem() { arRequireMaker_();
  var ui = SpreadsheetApp.getUi(), ss = arSpreadsheet_(), ed = lockEditors_(ss);
  if (ui.alert('Kunci struktur & periksa hak akses', 'Sheet laporan, Parameter, Log Persetujuan, dan judul kolom sheet input akan dikunci. Yang tetap boleh mengubah hanya:\n• ' + ed.join('\n• ') + '\n\n(Pemilik file serta PEMBUAT & PENYETUJU di Parameter — diperlukan agar aplikasi bisa menyimpan.) Lanjutkan?', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  // hapus kuncian lama buatan menu ini (aman dijalankan ulang)
  [SpreadsheetApp.ProtectionType.SHEET, SpreadsheetApp.ProtectionType.RANGE].forEach(function (t) { ss.getProtections(t).forEach(function (p) { if (String(p.getDescription()).indexOf(LOCK_TAG) === 0) p.remove(); }); });
  var n = 0;
  REPORT_SHEETS.concat(['Parameter', 'Log Persetujuan', 'Riwayat Revisi', 'Memorial AR', 'Referensi P&E', 'Rekonsiliasi Saldo', 'Analisis Bulanan AR', 'Jejak Edit Langsung', 'Log Surat Penagihan']).forEach(function (name) { var sh = ss.getSheetByName(name); if (sh) { protectRange_(sh, ed, name); n++; } });
  [['Monitoring Invoice', ENGINE.REG.headerRow], ['Log Penerimaan', ENGINE.LOG.headerRow], ['Master Pelanggan', ENGINE.MPEL.headerRow], ['Master Proyek', ENGINE.MPRO.headerRow],
    ['Riwayat Adendum', ENGINE.ADD.headerRow], ['Input Tagbrut', ENGINE.TAG.headerRow], ['Data LK (Tempel)', 7]].forEach(function (x) {
    var sh = ss.getSheetByName(x[0]); if (!sh) return; protectRange_(sh.getRange(1, 1, x[1], sh.getMaxColumns()), ed, 'judul kolom ' + x[0]); n++;
  });
  // hak berbagi: editor tidak boleh membagikan ulang; daftar akses ditampilkan
  var info = [], warn = [];
  try {
    var f = DriveApp.getFileById(ss.getId());
    try { f.setShareableByEditors(false); info.push('Editor tidak bisa lagi membagikan / mengubah izin file.'); } catch (e) { }
    var E = f.getEditors().map(function (u) { return u.getEmail().toLowerCase(); }), V = f.getViewers().map(function (u) { return u.getEmail().toLowerCase(); });
    E.forEach(function (e) { if (ed.indexOf(e) < 0) warn.push(e); });
    info.push('Editor: ' + (E.join(', ') || '-')); info.push('Pelihat / pengomentar: ' + (V.join(', ') || '-'));
  } catch (e) { info.push('Daftar akses tidak bisa dibaca: ' + e.message); }
  ui.alert('Selesai', n + ' bagian dikunci.\n\n' + info.join('\n') + (warn.length ? '\n\n⚠ Editor di luar daftar yang perlu diubah menjadi PELIHAT: ' + warn.join(', ') : '') +
    '\n\nLangkah manual (sekali): Bagikan ▸ ikon ⚙ ▸ hilangkan centang "Pelihat dan pengomentar dapat melihat opsi untuk mendownload, mencetak, dan menyalin".\n\nCatatan: pengguna ber-akses EDITOR (termasuk Section Head untuk persetujuan) secara teknis tetap bisa membuat salinan file dan melihat kode di Ekstensi ▸ Apps Script.', ui.ButtonSet.OK);
}

// =====================================================================  MENU: SAP & BUKU PENJUALAN
function fillSapFromLK() { arRequireMaker_();
  var ui = SpreadsheetApp.getUi(), res = compute_(), R = ENGINE.rekon(res.ctx);
  if (!R) { ui.alert('Sheet "Data LK (Tempel)" masih kosong. Tempel lampiran LK dan isi PERIODE LK di C4 terlebih dahulu.'); return; }
  var sh = sheet_('Monitoring Invoice'), map = colMap_(sh, ENGINE.REG, ENGINE.REG_ALIAS), n = 0, beda = [];
  R.suggest.forEach(function (s) {
    if (s.lama) { beda.push(s.nomor + ': sistem ' + s.lama + ' / LK ' + s.doc); return; }
    var row = findRegisterRow_(sh, s.nomor, map); if (row && map.sap_doc) { sh.getRange(row, map.sap_doc).setNumberFormat('@').setValue(s.doc); n++; }
  });
  rebuildAll();
  ui.alert(n + ' nomor dokumen SAP diisi dari lampiran LK per ' + R.C + '.' + (beda.length ? '\n\nTidak diubah (nomor di sistem berbeda dengan LK — mohon dicek):\n' + beda.slice(0, 15).join('\n') : ''));
}
function exportBukuPenjualan() { arRequireMaker_();
  var ui = SpreadsheetApp.getUi(), res = compute_(), ctx = res.ctx;
  var sel = ENGINE.parseMonthSel(sheet_('Rekap Invoice Keluar').getRange('C5').getValue(), ctx.today, true);
  var rows = ENGINE.bukuPenjualan(ctx, sel.start, sel.end);
  var ss = SpreadsheetApp.create('Buku Penjualan CSU — ' + sel.label), sh = ss.getSheets()[0];
  sh.setName('BUK PEN');
  sh.getRange('A1:A4').setValues([['PT WIJAYA KARYA BETON Tbk'], ['CONCRETE SERVICES UNIT'], ['BUKU PENJUALAN'], ['PERIODE ' + sel.label.toUpperCase()]]).setFontWeight('bold');
  var hdr = ['No', 'KODE FP', 'KD STATUS', 'NOMOR SERI FP', 'NOMOR FAKTUR', 'TANGGAL FP', 'NPWP', 'NAMA PELANGGAN', 'Uraian Faktur Pajak', 'HARGA JUAL', 'DPP NILAI LAIN', 'PPN', 'NO SAP', 'NOMOR INVOICE', 'CEK'];
  sh.getRange(6, 1, 1, hdr.length).setValues([hdr]).setBackground(C.navy).setFontColor('#FFFFFF').setFontWeight('bold').setWrap(true);
  sh.getRange(7, 1).setValue('I. PENYERAHAN PPN YANG HARUS DIPUNGUT SENDIRI').setFontWeight('bold');
  var data = rows.map(function (x, i) { return [i + 1, x.kode, x.status, x.seri, x.nomorFP, toCell_(x.tglFP, 'date'), x.npwp, x.nama, x.uraian, x.harga, x.dppnl, x.ppn, x.sap, x.nomor, x.cek]; });
  if (data.length) {
    sh.getRange(8, 1, data.length, hdr.length).setValues(data);
    sh.getRange(8, 2, data.length, 4).setNumberFormat('@'); sh.getRange(8, 7, data.length, 1).setNumberFormat('@'); sh.getRange(8, 13, data.length, 1).setNumberFormat('@');
    sh.getRange(8, 6, data.length, 1).setNumberFormat('dd/mm/yyyy'); sh.getRange(8, 10, data.length, 3).setNumberFormat(NUMF.money);
    var t = 8 + data.length;
    sh.getRange(t, 8, 1, 5).setValues([['TOTAL', '', '=SUM(J8:J' + (t - 1) + ')', '=SUM(K8:K' + (t - 1) + ')', '=SUM(L8:L' + (t - 1) + ')']]).setFontWeight('bold');
    sh.getRange(t, 10, 1, 3).setNumberFormat(NUMF.money);
  }
  arFormatGeneric_(sh);
  sh.setFrozenRows(6); sh.autoResizeColumns(1, hdr.length);
  ui.showModalDialog(HtmlService.createHtmlOutput('<p style="font:14px Arial">Buku Penjualan ' + sel.label + ' (' + data.length + ' faktur) sudah dibuat:</p><p><a target="_blank" href="' + ss.getUrl() + '">Buka file</a></p><p style="font:12px Arial;color:#555">Bagikan ke bagian Pajak. Harga jual memakai DPP pasti dari invoice (tidak dihitung mundur dari PPN).</p>').setWidth(420).setHeight(170), 'Ekspor Buku Penjualan');
}

// =====================================================================  GANTI NPP / NOMOR / ADENDUM (dari versi 1)
function renameNpp() { arRequireMaker_();
  var ui = SpreadsheetApp.getUi();
  var a = ui.prompt('Ganti NPP', 'NPP lama (mis. 264C 0000 JT):', ui.ButtonSet.OK_CANCEL); if (a.getSelectedButton() !== ui.Button.OK) return;
  var b = ui.prompt('Ganti NPP', 'NPP baru (format 264C 1234 JT):', ui.ButtonSet.OK_CANCEL); if (b.getSelectedButton() !== ui.Button.OK) return;
  var oldN = a.getResponseText().trim(), newN = b.getResponseText().trim().toUpperCase().replace(/\s+/g, ' ');
  if (!ENGINE.nppValid(newN)) { ui.alert('Format NPP baru tidak sesuai. Contoh: 264C 1234 JT'); return; }
  var ss = arSpreadsheet_(), n = 0;
  [['Master Proyek', ENGINE.MPRO, ENGINE.MPRO_ALIAS], ['Monitoring Invoice', ENGINE.REG, ENGINE.REG_ALIAS], ['Riwayat Adendum', ENGINE.ADD, null], ['Input Tagbrut', ENGINE.TAG, null], ['Input Tagbrut', ENGINE.NOTE, null]].forEach(function (x) {
    var sh = ss.getSheetByName(x[0]); if (!sh) return;
    var col = colMap_(sh, x[1], x[2]).npp, last = sh.getLastRow(); if (!col || last < x[1].firstRow) return;
    var r = sh.getRange(x[1].firstRow, col, last - x[1].firstRow + 1, 1), v = r.getValues();
    v.forEach(function (row) { if (String(row[0]).trim() === oldN) { row[0] = newN; n++; } });
    r.setValues(v);
  });
  var kp = ss.getSheetByName('Kartu Piutang'); if (kp && String(kp.getRange('C5').getValue()).trim() === oldN) kp.getRange('C5').setValue(newN);
  rebuildAll();
  ui.alert(n + ' sel NPP diganti dari ' + oldN + ' menjadi ' + newN + '.');
}
function normalizeNomor() { arRequireMaker_();
  var ss = arSpreadsheet_(), n = 0;
  function fix(sh, col, firstRow) {
    if (!sh || !col) return; var last = sh.getLastRow(); if (last < firstRow) return;
    var r = sh.getRange(firstRow, col, last - firstRow + 1, 1), v = r.getValues();
    v.forEach(function (row) { var s = String(row[0] || '').trim(); if (!s) return; var t = ENGINE.normInv(s); if (t !== s && /^KU\.02\.02\/WB-4C\./.test(t)) { row[0] = t; n++; } });
    r.setValues(v);
  }
  var mi = ss.getSheetByName('Monitoring Invoice'), lp = ss.getSheetByName('Log Penerimaan');
  fix(mi, colMap_(mi, ENGINE.REG, ENGINE.REG_ALIAS).nomor, ENGINE.REG.firstRow);
  fix(lp, colMap_(lp, ENGINE.LOG).inv, ENGINE.LOG.firstRow);
  fix(ss.getSheetByName('Riwayat Revisi'), 2, 6);
  var par = ss.getSheetByName('Parameter'), kv = par.getRange('A4:B120').getValues();
  kv.forEach(function (row, i) { if (String(row[0]).trim().toUpperCase() === 'PREFIX_INVOICE') par.getRange(4 + i, 2).setValue('KU.02.02/WB-4C.'); });
  rebuildAll();
  SpreadsheetApp.getUi().alert(n + ' nomor invoice diseragamkan ke format KU.02.02/WB-4C.xxxx/tahun.');
}
/** dari versi 1: membuat Riwayat Adendum & mengubah kolom Master Proyek lama (aman diulang) */
function setupAdendum() { arRequireMaker_();
  var ss = arSpreadsheet_(), mp = ss.getSheetByName('Master Proyek'), A = ENGINE.ADD, ui = SpreadsheetApp.getUi();
  var ra = ss.getSheetByName(A.sheet), dibuat = false, pindah = 0;
  if (!ra) {
    ra = ss.insertSheet(A.sheet, ss.getSheets().indexOf(mp) + 1); dibuat = true;
    ra.getRange('A1').setValue('RIWAYAT ADENDUM KONTRAK').setFontSize(14).setFontWeight('bold').setFontColor(C.navy);
    ra.getRange('A2').setValue('Satu baris = satu adendum. Nilai diisi angka positif; JENIS menentukan menambah atau mengurangi nilai kontrak. TANGGAL BERLAKU menentukan invoice mana yang memakai nilai baru.').setFontSize(9);
    ra.getRange(A.headerRow, 1, 1, A.cols.length).setValues([A.cols.map(function (c) { return c[1]; })])
      .setBackground(C.navy).setFontColor('#FFFFFF').setFontWeight('bold').setWrap(true).setHorizontalAlignment('center').setVerticalAlignment('middle');
    A.cols.forEach(function (c, i) { ra.setColumnWidth(i + 1, c[2] * 7); });
    ra.getRange(A.firstRow, 1, 400, A.cols.length).setBackground(C.input).setFontFamily('Arial').setFontSize(9);
    ra.getRange(A.firstRow, 4, 400, 1).setNumberFormat(NUMF.date); ra.getRange(A.firstRow, 6, 400, 1).setNumberFormat(NUMF.money);
    ra.getRange(A.firstRow, 1, 400, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(mp.getRange('A3:A400'), true).setAllowInvalid(true).build());
    ra.getRange(A.firstRow, 5, 400, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(ENGINE.OPT.jenis_add, true).setAllowInvalid(false).build());
    ra.setFrozenRows(A.headerRow); ra.setTabColor('#9C6500'); ra.setHiddenGridlines(true);
  }
  var lc = mp.getLastColumn(), hdr = mp.getRange(2, 1, 1, lc).getValues()[0].map(norm_);
  var kLama = hdr.indexOf('NILAI ADDENDUM TERAKHIR') + 1, kNilai = Math.max(hdr.indexOf('NILAI KONTRAK (EXCL. PPN)'), hdr.indexOf('NILAI KONTRAK AWAL (EXCL. PPN)')) + 1;
  if (kLama > 0 && kNilai > 0) {
    var last = Math.max(mp.getLastRow(), 3), n = last - 2;
    var npp = mp.getRange(3, 1, n, 1).getValues(), nil = mp.getRange(3, kNilai, n, 1).getValues(), lama = mp.getRange(3, kLama, n, 1).getValues(), rows = [];
    for (var i = 0; i < n; i++) {
      var p = String(npp[i][0] || '').trim(), v0 = Number(nil[i][0]) || 0, v1 = Number(lama[i][0]) || 0;
      if (!p || !v1) continue;
      if (!v0) { mp.getRange(3 + i, kNilai).setValue(v1); continue; }
      var d = v1 - v0; if (Math.abs(d) < 1) continue;
      rows.push([p, 1, '', '', d > 0 ? 'TAMBAH' : 'KURANG', Math.abs(Math.round(d)), 'Dari kolom lama "Nilai Addendum Terakhir" (Rp ' + v1 + '). Lengkapi NO. DOKUMEN & TANGGAL BERLAKU.']);
    }
    if (rows.length) { var lr = Math.max(ra.getLastRow(), A.headerRow); roomFor_(ra, lr + rows.length); ra.getRange(lr + 1, 1, rows.length, rows[0].length).setValues(rows); pindah = rows.length; }
    mp.insertColumnAfter(kLama);
    var L = function (c) { return mp.getRange(1, c).getA1Notation().replace(/\d+/g, ''); };
    var cJ = L(kNilai), cT = L(kLama), cK = L(kLama + 1);
    mp.getRange(2, kNilai).setValue('NILAI KONTRAK AWAL (EXCL. PPN)');
    mp.getRange(2, kLama, 1, 3).setValues([['TOTAL ADENDUM TAMBAH (OTOMATIS)', 'TOTAL ADENDUM KURANG (OTOMATIS)', 'NILAI OK / SETELAH ADENDUM (OTOMATIS)']])
      .setBackground(C.autoHdr).setFontColor('#FFFFFF').setFontWeight('bold').setWrap(true).setHorizontalAlignment('center');
    var maxR = Math.max(400, mp.getLastRow()), f = [];
    for (var r = 3; r <= maxR; r++) f.push([
      '=IF($A' + r + '="","",SUMIFS(\'Riwayat Adendum\'!$F:$F,\'Riwayat Adendum\'!$A:$A,$A' + r + ',\'Riwayat Adendum\'!$E:$E,"TAMBAH"))',
      '=IF($A' + r + '="","",SUMIFS(\'Riwayat Adendum\'!$F:$F,\'Riwayat Adendum\'!$A:$A,$A' + r + ',\'Riwayat Adendum\'!$E:$E,"KURANG"))',
      '=IF($A' + r + '="","",N(' + cJ + r + ')+N(' + cT + r + ')-N(' + cK + r + '))']);
    mp.getRange(3, kLama, f.length, 3).setFormulas(f).setNumberFormat(NUMF.money).setBackground(C.auto).clearDataValidations();
  }
  rebuildAll();
  ui.alert('Riwayat Adendum siap.' + (dibuat ? '\n• Sheet "Riwayat Adendum" dibuat.' : '') + (pindah ? '\n• ' + pindah + ' adendum dari kolom lama dipindahkan.' : ''));
}

// =====================================================================  APLIKASI (WEB APP)
function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('AR CSU — Receivables Management')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function include(name) { if (['Index','EngineClient','PdfLib'].indexOf(name) < 0) throw new Error('Template not allowed.'); return HtmlService.createHtmlOutputFromFile(name).getContent(); }
/** pustaka pembuat PDF (html2canvas + jsPDF) untuk arsip — dimuat aplikasi hanya saat dibutuhkan */
function apiPdfLib() { arRequireAccess_(); return HtmlService.createHtmlOutputFromFile('PdfLib').getContent(); }
function openAppDialog() {
  var html = HtmlService.createTemplateFromFile('Index').evaluate().setWidth(1400).setHeight(900);
  SpreadsheetApp.getUi().showModalDialog(html, 'AR CSU — Penagihan, Penerimaan & Otorisasi Dokumen');
}
/** link resmi aplikasi web: disimpan lewat menu (Script Properties) karena ScriptApp.getService().getUrl() bisa mengembalikan ID deployment lain */
var AR_URL_RE_ = /^https:\/\/script\.google\.com\/(a\/[^\/]+\/)?macros\/s\/[\w-]+\/exec$/;
function arAppUrl_() {
  var u = ''; try { u = PropertiesService.getScriptProperties().getProperty('APP_URL') || ''; } catch (e) { }
  return { url: u || ScriptApp.getService().getUrl() || '', saved: !!u };
}
function setAppUrl(url) {
  arRequireMaker_();
  url = String(url || '').trim().replace(/\?.*$/, '');
  if (!AR_URL_RE_.test(url)) throw new Error('Link harus berbentuk https://script.google.com/macros/s/…/exec (salin dari Terapkan ▸ Kelola deployment ▸ URL aplikasi web).');
  PropertiesService.getScriptProperties().setProperty('APP_URL', url); arBootBust_();
  return url;
}
function showAppLink() {
  var a = arAppUrl_(), e = function (x) { return String(x).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
  var h = '<div style="font:13px/1.5 Arial,sans-serif;color:#14213D">' +
    (a.url ? '<p style="margin:0 0 6px">Link aplikasi' + (a.saved ? '' : ' <b style="color:#B42318">(belum disimpan — bisa salah deployment)</b>') + ':</p>' +
      '<p style="margin:0 0 10px;word-break:break-all"><a id="u" href="' + e(a.url) + '" target="_blank">' + e(a.url) + '</a></p>' +
      '<input id="c" readonly value="' + e(a.url) + '" style="position:absolute;left:-9999px">' +
      '<button onclick="var c=document.getElementById(\'c\');c.select();var ok=false;try{ok=document.execCommand(\'copy\');}catch(x){}document.getElementById(\'m\').textContent=ok?\'✓ Link disalin\':\'Tekan Ctrl/Cmd+C untuk menyalin\';" style="padding:7px 14px;border:0;border-radius:7px;background:#0F6E78;color:#fff;font-weight:700;cursor:pointer">Salin link</button> <span id="m" style="color:#1E7A45"></span>' +
      '<p style="margin:12px 0 4px">Bagikan ke Fikri dan Pak Anas (keduanya perlu akses edit ke spreadsheet ini). Bila terbuka halaman "Maaf, saat ini tidak dapat membuka file", buka di jendela Incognito / profil Chrome dengan satu akun saja.</p>' :
      '<p>Aplikasi belum di-deploy sebagai Web App.</p>') +
    '<hr style="border:0;border-top:1px solid #DCE3EB;margin:12px 0"><p style="margin:0 0 4px"><b>Ganti / simpan link resmi</b> — tempel URL dari <i>Terapkan ▸ Kelola deployment ▸ URL aplikasi web</i>:</p>' +
    '<input id="n" style="width:100%;box-sizing:border-box;padding:7px;border:1px solid #BFC8D4;border-radius:6px" placeholder="https://script.google.com/macros/s/…/exec">' +
    '<p><button onclick="var v=document.getElementById(\'n\').value;google.script.run.withSuccessHandler(function(u){document.getElementById(\'m2\').textContent=\'✓ Tersimpan. Tutup lalu buka menu ini lagi.\';}).withFailureHandler(function(x){document.getElementById(\'m2\').textContent=x.message;}).setAppUrl(v)" style="padding:7px 14px;border:1px solid #0F6E78;border-radius:7px;background:#fff;color:#0F6E78;font-weight:700;cursor:pointer">Simpan link</button> <span id="m2"></span></p></div>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(h).setWidth(560).setHeight(a.saved ? 330 : 360), 'Link aplikasi AR CSU');
}
function plain_(o) { return JSON.parse(JSON.stringify(nd_(o), function (k, v) { return v === Infinity ? null : v; })); }
function email_() { var e = ''; try { e = Session.getActiveUser().getEmail(); } catch (x) { } return String(e || '').toLowerCase(); }

/** data aplikasi: hitungan ringan (tanpa laporan) + cache 5 menit per pengguna; cache otomatis gugur setiap ada penyimpanan (arBootBust_) */
function apiBootstrap() {
  var access = arRequireAccess_(), em = email_(), key = arBootKey_(em), hit = arCacheGet_(key);
  if (hit) { var b0 = JSON.parse(hit); b0.access = access; b0.cached = true; return b0; }
  var input = readAll_(); input.light = true;
  var res = ENGINE.compute(input), b = ENGINE.bootstrap(res, em);
  b.appUrl = arAppUrl_().url;
  b.charts = access.maker ? arChartsCached_() : null;
  b.access = access; b.summary = arDashboard_(res.ctx); b.sheetUrl = arSpreadsheet_().getUrl();
  var out = plain_(b); arCachePut_(key, JSON.stringify(out), 300);
  return out;
}
function apiGetInvoice(nomor) { arRequireAccess_();
  var sh = sheet_('Monitoring Invoice'), map = colMap_(sh, ENGINE.REG, ENGINE.REG_ALIAS), row = findRegisterRow_(sh, nomor, map);
  if (!row) return null;
  var vals = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0], o = {};
  ENGINE.REG.cols.forEach(function (c) { var j = map[c[0]]; var v = j ? vals[j - 1] : ''; o[c[0]] = isDate_(v) ? dstr_(v) : v; });
  try { o.form = o.json ? JSON.parse(o.json) : null; } catch (e) { o.form = null; }
  delete o.json;
  return plain_(o);
}
/** p = data register dari aplikasi; opt = {revisi: bool, alasan: string} */
function apiSaveInvoice(p, opt) { arRequireMaker_();
  opt = opt || {};
  arValidateInvoice_(p);
  var lock = arLock_(); lock.waitLock(28000);
  var result;
  try {
    var ss = arSpreadsheet_(), sh = ss.getSheetByName('Monitoring Invoice'), cols = ENGINE.REG.cols, map = colMap_(sh, ENGINE.REG, ENGINE.REG_ALIAS);
    var nomor = ENGINE.normInv(p.nomor);
    if (!/^KU\.02\.02\/WB-4C\.\d{4}[A-Z]?\/\d{4}$/.test(nomor)) throw new Error('Invalid invoice number format: ' + p.nomor);
    var row = findRegisterRow_(sh, nomor, map);
    if (row && !opt.revisi) { var ex = apiGetInvoice(nomor); return { needConfirm: true, existing: { nomor: nomor, pelanggan: ex.pelanggan, tgl: ex.tgl, dpp: ex.dpp, ppn: ex.ppn, rev: ex.rev || 0 } }; }
    if (row && opt.revisi && !String(opt.alasan || '').trim()) throw new Error('A revision reason is required.');
    if (row && map.status_inv && /^BATAL/i.test(String(sh.getRange(row, map.status_inv).getValue()))) throw new Error('Invoice ' + nomor + ' has been cancelled — it cannot be revised. Create a new invoice with a different number.');
    checkNppSave_(p.npp, p.pelanggan);
    arCheckInvoiceNumbers_(p, nomor);
    var lc = sh.getLastColumn(), old = row ? sh.getRange(row, 1, 1, lc).getValues()[0] : null, oo = {};
    if (row) cols.forEach(function (c) { var j = map[c[0]]; oo[c[0]] = j ? (isDate_(old[j - 1]) ? dstr_(old[j - 1]) : old[j - 1]) : ''; });
    // jenis revisi ditentukan dari perubahan sebenarnya: NILAI / INVOICE (teks invoice-kwitansi) → Rev naik & ajukan ulang; BAP saja → cukup tercatat
    var rj = row ? revisiJenis_(oo, p) : null;
    if (rj && !rj.jenis) throw new Error('No changes to save.');
    // tutup periode hanya untuk invoice baru / revisi yang mengubah nilai atau tanggal (revisi teks tidak berdampak akuntansi)
    var koreksi = !row || rj.jenis === 'NILAI' ? arCheckPeriod_([p.tgl, row ? String(oo.tgl || '') : ''], opt.periodeAlasan, 'Invoice') : '';
    var now = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'), user = email_() || 'pengguna';
    p.nomor = nomor; p.json = JSON.stringify(p.form || {}); p.dibuat = now + ' · ' + user + (koreksi ? ' · ' + koreksi : ''); p.sumber = 'APLIKASI';
    if (row) {
      if (opt.expectedRev !== undefined && Number(opt.expectedRev) !== (Number(oo.rev) || 0)) throw new Error('The invoice was revised by another user. Reopen it before saving.');
      // revisi BAP saja: nomor revisi tetap → sidik & persetujuan paket invoice tetap sah
      var rev = rj.jenis === 'BAP' ? (Number(oo.rev) || 0) : (Number(oo.rev) || 0) + 1; p.rev = rev;
      var perluFP = (Math.round(Number(oo.dpp) || 0) !== Math.round(p.dpp) || Math.round(Number(oo.ppn) || 0) !== Math.round(p.ppn)) ? 'YA — DPP/PPN berubah' : 'TIDAK';
      var jenisTxt = { NILAI: 'REVISI NILAI — ajukan ulang persetujuan', INVOICE: 'REVISI INVOICE/KWITANSI (nilai tetap) — ajukan ulang persetujuan', BAP: 'REVISI DOKUMEN BAP (nilai & persetujuan tetap)' }[rj.jenis];
      var rr = ss.getSheetByName('Riwayat Revisi');
      rr.appendRow([now, nomor, rev, (koreksi ? koreksi + ' ' : '') + opt.alasan, user, perluFP, Number(oo.dpp) || 0, Number(oo.ppn) || 0, (Number(oo.dpp) || 0) + (Number(oo.ppn) || 0),
        p.dpp, p.ppn, p.dpp + p.ppn, JSON.stringify(oo).substring(0, 49000), jenisTxt, rj.diff.join('; ').substring(0, 49000)]);
      revisiHeader_(rr);
      var nv = old.slice();
      cols.forEach(function (c) { var j = map[c[0]]; if (j && colType_(c) === 'app' && p[c[0]] !== undefined) nv[j - 1] = toCell_(p[c[0]], colFmt_(c)); });
      sh.getRange(row, 1, 1, lc).setValues([nv]);
      result = { ok: true, mode: 'revisi', nomor: nomor, rev: rev, perluFP: perluFP, jenis: rj.jenis, perubahan: rj.diff };
    } else {
      p.rev = 0; p.status_inv = 'AKTIF';
      var target = lastDataRow_(sh, ENGINE.REG.firstRow, [map.nomor]) + 1, lc2 = sh.getLastColumn(), arr = []; roomFor_(sh, target);
      for (var k = 0; k < lc2; k++) arr.push('');
      cols.forEach(function (c) { var j = map[c[0]]; if (j && (colType_(c) === 'app' || c[0] === 'status_inv' || c[0] === 'top') && p[c[0]] !== undefined) arr[j - 1] = toCell_(p[c[0]], colFmt_(c)); });
      var keep = sh.getRange(target, 1, 1, lc2).getFormulas()[0];
      arr = arr.map(function (v, i) { return keep[i] ? keep[i] : v; });
      sh.getRange(target, 1, 1, lc2).setValues([arr]);
      result = { ok: true, mode: 'baru', nomor: nomor, rev: 0, row: target };
    }
  } finally { lock.releaseLock(); }
  result.saved = true;
  var post = arQueueAfterSave_([p.npp], [String(p.tgl || '').substring(0, 7)]);
  result.reportPending = post.reportPending; result.background = post.background; result.warnings = post.warnings;
  if (koreksi) { var rf = arRefreshClosedMonths_([p.tgl, isDate_(tglLama) ? dstr_(tglLama) : String(tglLama || '')], koreksi); result.warnings = result.warnings.concat(rf.warnings); result.snapshots = rf.ids;
    var nw = arNotify_('Koreksi periode tertutup — invoice ' + nomor, ['Pembuat: ' + user, 'Mode: ' + result.mode + (result.mode === 'revisi' ? ' (rev ' + result.rev + ', alasan: ' + opt.alasan + ')' : ''), 'Tanggal invoice: ' + p.tgl, 'DPP Rp ' + ENGINE.fmtRp(p.dpp) + ' · PPN Rp ' + ENGINE.fmtRp(p.ppn), 'Tanda: ' + koreksi]); if (nw) result.warnings.push(nw); }
  return result;
}
// ---------------------------------------------------------------------  JENIS REVISI INVOICE
var REV_NILAI_ = { tgl: 'Tanggal invoice', jenis: 'Jenis tagihan', pelanggan: 'Pelanggan', npp: 'NPP', bruto: 'Bruto', umk: 'Potongan UM', retensi: 'Retensi', retpos: 'Posisi retensi', dpp: 'DPP', ppn: 'PPN', tarif: 'Tarif PPh', dasar: 'Dasar PPh', tagih: 'Nilai di invoice' };
var REV_DOK_ = { invke: 'Invoice ke-', ttd: 'Penandatangan', uraian: 'Uraian tagihan' };
var REV_BAP_ = { bapTgl: 'Tanggal BAP', bapProg: 'Progress penagihan (BAP)', bapSpk: 'Nomor SPK (BAP)', bapSpkTgl: 'Tanggal SPK (BAP)' };
var REV_ABAI_ = ['alasan', '_loadedRev', 'judulAuto', 'invkeAuto', 'bapProgAuto', 'ppnRetensiAlasan'];
/** bandingkan baris lama (oo) dengan simpanan baru (p) → { jenis: NILAI | INVOICE | BAP | '', diff: ['DPP: a → b', …] } */
function revisiJenis_(oo, p) {
  var txt = function (v) { return v === null || v === undefined ? '' : String(v).trim(); }, diff = [], has = { NILAI: 0, INVOICE: 0, BAP: 0 };
  var add = function (g, lbl, a, b) { has[g] = 1; diff.push(lbl + ': "' + (txt(a) || '-') + '" → "' + (txt(b) || '-') + '"'); };
  Object.keys(REV_NILAI_).forEach(function (k) {
    if (p[k] === undefined) return;
    var num = ['bruto', 'umk', 'retensi', 'dpp', 'ppn', 'tagih'].indexOf(k) >= 0, a = oo[k], b = p[k];
    var same = num ? Math.round(Number(a) || 0) === Math.round(Number(b) || 0) : k === 'tarif' ? Math.abs((Number(a) || 0) - (Number(b) || 0)) < 1e-9 : txt(a).toUpperCase() === txt(b).toUpperCase();
    if (!same) add('NILAI', REV_NILAI_[k], a, b);
  });
  Object.keys(REV_DOK_).forEach(function (k) { if (p[k] !== undefined && txt(oo[k]) !== txt(p[k])) add('INVOICE', REV_DOK_[k], oo[k], p[k]); });
  var of = {}; try { of = JSON.parse(String(oo.json || '{}')) || {}; } catch (e) { of = {}; }
  var nf = p.form || {}, legacy = !String(oo.json || '').trim(), keys = {};
  Object.keys(of).concat(Object.keys(nf)).forEach(function (k) { keys[k] = 1; });
  Object.keys(keys).forEach(function (k) {
    if (REV_ABAI_.indexOf(k) >= 0) return;
    var a = of[k], b = nf[k], same = JSON.stringify(a === undefined ? '' : a) === JSON.stringify(b === undefined ? '' : b);
    if (same) return;
    if (REV_BAP_[k]) add('BAP', REV_BAP_[k] + (/^bapSpk/.test(k) && !txt(a) ? ' (sebelumnya dari Master Proyek)' : ''), a, b);
    else if (!legacy && !/^(judul|invke)$/.test(k)) add('INVOICE', 'Isian form ' + k, typeof a === 'object' ? JSON.stringify(a) : a, typeof b === 'object' ? JSON.stringify(b) : b);
  });
  return { jenis: has.NILAI ? 'NILAI' : has.INVOICE ? 'INVOICE' : has.BAP ? 'BAP' : '', diff: diff };
}
/** judul kolom tambahan Riwayat Revisi (ditambahkan sekali) */
function revisiHeader_(rr) {
  var top = rr.getRange(1, 1, Math.min(6, rr.getLastRow()), 2).getValues(), hr = 0;
  top.forEach(function (r, i) { if (!hr && /NOMOR|INVOICE/i.test(String(r[1]))) hr = i + 1; });
  if (hr && !String(rr.getRange(hr, 14).getValue()).trim()) rr.getRange(hr, 14, 1, 2).setValues([['JENIS REVISI', 'PERUBAHAN']]).setFontWeight('bold');
}
/** baris Log Penerimaan dari aplikasi → array sesuai urutan skema LOG */
function payRow_(p) { return ENGINE.LOG.cols.map(function (c) { var v = p[c[0]]; return v === undefined || v === null ? '' : v; }); }
function checkPayRows_(rows) {
  if (!rows || !rows.length) throw new Error('No payment rows yet.');
  rows.forEach(function (p, i) {
    var jt = String(p.jenis_trx || '').toUpperCase(), pb = jt === 'PINDAH BUKU', kom = jt === 'KOMPENSASI', n = Number(p.nominal) || 0;
    if (!p.inv || !p.tgl || !n || (!pb && n < 0)) throw new Error('Row ' + (i + 1) + ': date, invoice number and amount are required (negative amounts only for offset / reclass).');
    if (!pb && !kom && !String(p.rekening || '').trim()) throw new Error('Row ' + (i + 1) + ': select the receiving bank account.');
    p.inv = ENGINE.normInv(p.inv);
  });
  var pbs = rows.filter(function (p) { return String(p.jenis_trx || '').toUpperCase() === 'PINDAH BUKU'; });
  if (pbs.length && Math.abs(pbs.reduce(function (t, p) { return t + Number(p.nominal); }, 0)) > 1) throw new Error('Offset / reclass does not balance: rows must sum to 0 (source negative, target positive).');
}
/** simpan beberapa baris sekaligus (satu transfer untuk beberapa invoice, atau pindah buku) */
function apiSavePayments(rows, requestId, opt) { return arSavePayments_(rows, requestId, opt); }
/** pembayaran tunggal (kompatibel versi 2.0) */
function apiSavePayment(p) { var r = apiSavePayments([p]); return plain_({ ok: true, saved: true, reportPending: r.reportPending, warnings:r.warnings, duplicate:r.duplicate, invoice: (r.invoices || [])[0] || null, payment: (r.payments || []).slice(-1)[0] || null }); }
/** pratinjau tanpa menyimpan: kartu piutang & alokasi bila baris-baris ini dicatat */
function apiPayPreview(rows) { arRequireMaker_();
  rows = (rows || []).filter(function (p) { return p && p.inv && Number(p.nominal); });
  rows.forEach(function (p) { p.inv = ENGINE.normInv(p.inv); });
  var input = readAll_(), n0 = input.log.length;
  rows.forEach(function (p) { input.log.push(payRow_(p)); });
  input.newFrom = n0; input.light = true;
  var res = ENGINE.compute(input), ctx = res.ctx, keys = rows.map(function (p) { return p.inv; }), npps = {};
  var out = ctx.pays.filter(function (p) { return p._new; }).map(function (p) {
    var i = p._inv; if (i) npps[i.npp] = 1;
    return { inv: i ? i.nomor : p.inv, keN: p._keN || 0, keP: p._keP || 0, pph: (p._pph || 0) + (p._sel || 0), adm: p._adm || 0, aturan: p._rule || '', cek: p._cek.join('; '), status: i ? i._res.status : '', sisa: i ? i._res.sisa : 0 };
  });
  return plain_({ rows: out, kartu: Object.keys(npps).map(function (n) { return ENGINE.kartuPlain(ctx, n, ctx.today); }) });
}
function apiKartu(npp, asOf) { arRequireAccess_(); var res = computeLight_(); return plain_(ENGINE.kartuPlain(res.ctx, npp, asOf || res.ctx.today)); }

// =====================================================================  PERSETUJUAN DIGITAL
function secret_() {
  var sp = PropertiesService.getScriptProperties(), s = sp.getProperty('APPROVAL_SECRET');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); sp.setProperty('APPROVAL_SECRET', s); }
  return s;
}
function sign_(r) {
  var msg = [r.id, r.jenis, r.ref, r.sidik, String(r.putus_oleh).toLowerCase(), r.putus_waktu, r.status].join('|');
  var b = Utilities.computeHmacSha256Signature(msg, secret_()), hex = b.map(function (x) { return ('0' + (x & 255).toString(16)).slice(-2); }).join('').toUpperCase();
  return hex.substring(0, 5) + '-' + hex.substring(5, 10);
}
function apprRows_() {
  var sh = sheet_('Log Persetujuan'); if (!sh) throw new Error('Sheet "Log Persetujuan" does not exist yet. Run menu AR CSU ▸ Upgrade struktur file ke versi 2.');
  var map = colMap_(sh, ENGINE.APPR), last = sh.getLastRow(), rows = [];
  if (last >= ENGINE.APPR.firstRow) {
    var v = sh.getRange(ENGINE.APPR.firstRow, 1, last - ENGINE.APPR.firstRow + 1, sh.getLastColumn()).getValues();
    v.forEach(function (r, i) { var o = { _row: ENGINE.APPR.firstRow + i }; ENGINE.APPR.cols.forEach(function (c) { var j = map[c[0]]; o[c[0]] = j ? (r[j - 1] instanceof Date ? Utilities.formatDate(r[j - 1], TZ, 'yyyy-MM-dd HH:mm') : r[j - 1]) : ''; }); if (String(o.id).trim()) rows.push(o); });
  }
  return { sh: sh, map: map, rows: rows };
}
function verifyAppr_(a) {
  if (!a || !a.kode) return a;
  try { var A = apprRows_().rows.filter(function (r) { return String(r.id) === String(a.id); })[0]; a.valid = !!A && sign_(A) === String(A.kode); } catch (e) { a.valid = false; }
  return a;
}
function apiPackage(kind, ref) { arRequireAccess_();
  var res = computeLight_('memo'), d = ENGINE.packageFor(res.ctx, kind, ref); // lembar pengesahan tidak memerlukan laporan
  if (!d) throw new Error('Data not found: ' + kind + ' ' + ref);
  if (kind === 'VAL') d.draft = valDraft_(d.ref);
  if (d.appr) verifyAppr_(d.appr);
  d.user = ENGINE.userBy(res.params, email_()) || { email: email_() };
  return plain_(d);
}
function apiSubmit(kind, ref, extra) { arRequireMaker_();
  var lock = arLock_(); lock.waitLock(28000);
  var out;
  try {
    var res = computeLight_('memo'), P = res.params, me = email_(), u = ENGINE.userBy(P, me);
    if (P.users.length && (!u || (u.peran.indexOf('PEMBUAT') !== 0 && u.peran.indexOf('PENYETUJU') !== 0))) throw new Error('Account ' + (me || '(unknown)') + ' is not registered as PEMBUAT in Parameter ▸ Pengguna & Peran.');
    var d = ENGINE.packageFor(res.ctx, kind, ref);
    if (!d) throw new Error('Data not found: ' + kind + ' ' + ref);
    if (kind === 'TUTUP' && !String(d.alasan || '').trim()) throw new Error('Fill in ALASAN TUTUP MANUAL in the Monitoring Invoice sheet first.');
    if ((kind === 'VAL') && !d.items.filter(function (x) { return !x.gone; }).length) throw new Error('No open invoices for ' + d.label + '.');
    if (kind === 'VAL' && extra && extra.items) {
      // v2.7: usulan AR dari aplikasi disimpan dulu, lalu diperiksa (tanggal rencana cair wajib untuk ✓ / ◇)
      writeUsulan_(extra);
      res = computeLight_('memo'); d = ENGINE.packageFor(res.ctx, kind, ref, { fresh: true });
      var ue = ENGINE.valCheckUsulan(d, extra); if (ue) throw new Error(ue);
    }
    // tutup buku: pembatalan invoice & lembar validasi periode tertutup hanya dengan izin koreksi; alasan wajib & ditandai
    var kTgl = arGuardClosed_(kind, d, 'Request ' + kind + ' ' + d.ref);
    if (kTgl && extra && extra.alasan !== undefined) extra.alasan = '[KOREKSI PERIODE TERTUTUP] ' + String(extra.alasan).trim();
    if (kind === 'VALX') {
      if (!d.canSubmit) throw new Error(d.blok);
      if (String((extra && extra.alasan) || '').trim().length < 10) throw new Error('A cancellation reason is required (at least 10 characters).');
    }
    if (kind === 'BATAL') {
      if (!d.canSubmit) throw new Error(d.blok);
      var als = String((extra && extra.alasan) || '').trim(); if (als.length < 10) throw new Error('A cancellation reason is required (at least 10 characters).');
    }
    var snap = { label: d.label, nilai: d.nilai || d.total || 0 };
    if (kind === 'VAL') { snap = ENGINE.valSnapshot(d, extra || {}); try { arProps_().deleteProperty('VALDRAFT_' + d.ref); } catch (e) { } }
    if (kind === 'VALX') snap = { label: d.label, nilai: 0, alasan: String(extra.alasan).trim(), list: d.list };
    if (kind === 'BATAL') snap = { label: d.label, nilai: d.nilai, nomor: d.inv.nomor, pelanggan: d.inv.pelanggan, alasan: String(extra.alasan).trim(), pengganti: String(extra.pengganti || '').trim() };
    if (kind === 'REK') { snap.invoices = d.invoices.length; snap.pays = d.pays.length; snap.totDebit = d.totDebit; snap.ckpn = d.ckpn; }
    var A = apprRows_(), now = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'), jenis = ENGINE.DOC[kind];
    A.rows.filter(function (r) { return String(r.jenis).toUpperCase() === jenis && String(r.ref) === String(d.ref) && String(r.status).toUpperCase() === 'MENUNGGU'; })
      .forEach(function (r) { A.sh.getRange(r._row, A.map.status).setValue('DIBATALKAN'); A.sh.getRange(r._row, A.map.catatan).setValue('Diganti pengajuan baru ' + now); });
    var maxId = 0; A.rows.forEach(function (r) { var m = String(r.id).match(/(\d+)$/); if (m) maxId = Math.max(maxId, +m[1]); });
    var rec = { id: 'APV-' + ('0000' + (maxId + 1)).slice(-5), waktu: now, jenis: jenis, ref: String(d.ref), label: d.label + (kind === 'VAL' ? ' · Rev.' + ('0' + snap.rev).slice(-2) + ' (' + snap.items.length + ' invoice)' : ''),
      nilai: Math.round(snap.nilai || 0), sidik: d.fp, oleh: me, status: 'MENUNGGU', putus_oleh: '', putus_waktu: '', catatan: '', kode: '', data: JSON.stringify(snap) };
    // v2.8.7: jenis yang disahkan lewat tanda tangan basah Section Head (Parameter PERSETUJUAN_TTD_BASAH) langsung tercatat DISETUJUI;
    // tidak atas nama Section Head secara elektronik (PUTUS OLEH kosong, tanpa cap elektronik) — bukti persetujuan = tanda tangan basah di lembar cetak
    if (ENGINE.basahKinds(P).indexOf(kind) >= 0) arBasah_(rec, now);
    var lc = Math.max(A.sh.getLastColumn(), ENGINE.APPR.cols.length), arr = []; for (var k = 0; k < lc; k++) arr.push('');
    ENGINE.APPR.cols.forEach(function (c) { var j = A.map[c[0]]; if (j) arr[j - 1] = rec[c[0]]; });
    var row = lastDataRow_(A.sh, ENGINE.APPR.firstRow, [A.map.id]) + 1; roomFor_(A.sh, row);
    A.sh.getRange(row, 1, 1, lc).setNumberFormat('@').setValues([arr]);
    if (A.map.nilai) A.sh.getRange(row, A.map.nilai).setNumberFormat(NUMF.money).setValue(rec.nilai);
    out = { ok: true, id: rec.id, label: rec.label, status: rec.status, basah: rec.status === 'DISETUJUI' };
  } finally { lock.releaseLock(); }
  // v2.8.6: laporan diperbarui di latar belakang (aplikasi memanggil apiAfterSave tanpa menunggu) — pengajuan langsung selesai
  out.background = arQueueAfterSave_().background;
  return out;
}
function arBasah_(rec, now) {
  rec.status = 'DISETUJUI'; rec.putus_oleh = ''; rec.putus_waktu = now;
  rec.catatan = ENGINE.TTD_BASAH_TAG + ' Persetujuan Section Head melalui tanda tangan basah pada lembar cetak — dicatat otomatis saat diajukan';
  rec.kode = sign_({ id: rec.id, jenis: rec.jenis, ref: rec.ref, sidik: rec.sidik, putus_oleh: '', putus_waktu: now, status: 'DISETUJUI' });
  return rec;
}
/** pengajuan lama yang masih MENUNGGU (jenis TTD basah) disahkan sekaligus — hanya bila datanya tidak berubah sejak diajukan */
function arSahkanCore_(hitungSaja) {
  var lock = arLock_(), n = 0, ubah = 0, K; lock.waitLock(28000);
  try {
    var res = computeLight_('memo'), A = apprRows_(), now = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'); K = ENGINE.basahKinds(res.params);
    var L = A.rows.filter(function (r) { var k = Object.keys(ENGINE.DOC).filter(function (x) { return ENGINE.DOC[x] === String(r.jenis).toUpperCase(); })[0]; r._k = k; return String(r.status).toUpperCase() === 'MENUNGGU' && K.indexOf(k) >= 0; });
    if (hitungSaja) return { n: L.length, kinds: K };
    L.forEach(function (r) {
      var d = ENGINE.packageFor(res.ctx, r._k, r.ref); if (!d || d.fp !== String(r.sidik)) { ubah++; return; }
      r.jenis = String(r.jenis).toUpperCase(); arBasah_(r, now);
      ['status', 'putus_oleh', 'putus_waktu', 'catatan', 'kode'].forEach(function (k) { if (A.map[k]) A.sh.getRange(r._row, A.map[k]).setNumberFormat('@').setValue(r[k]); });
      n++;
    });
  } finally { lock.releaseLock(); }
  if (n) arQueueAfterSave_();
  return { n: n, ubah: ubah, kinds: K };
}
function arSahkanTtdBasah() { arRequireMaker_();
  var ui = SpreadsheetApp.getUi(), c = arSahkanCore_(true);
  if (!c.n) { ui.alert('Tidak ada pengajuan menunggu untuk jenis: ' + (c.kinds.join(', ') || '-') + '.'); return; }
  if (ui.alert('Sahkan dengan TTD basah', c.n + ' pengajuan menunggu (' + c.kinds.join(', ') + ') akan dicatat DISETUJUI melalui tanda tangan basah Section Head. Pengajuan yang datanya berubah sejak diajukan dilewati. Lanjutkan?', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  var r = arSahkanCore_(false);
  ui.alert('Selesai', r.n + ' pengajuan disahkan (TTD basah).' + (r.ubah ? '\n' + r.ubah + ' dilewati karena datanya berubah — ajukan ulang dari aplikasi.' : ''), ui.ButtonSet.OK);
}
function apiSahkanTtdBasah() { arRequireMaker_(); var r = arSahkanCore_(false); r.background = true; return r; }
function apiDecide(id, keputusan, catatan, keputusanItem) { arRequireAccess_();
  if (['SETUJU','TOLAK'].indexOf(String(keputusan).toUpperCase()) < 0) throw new Error('Invalid decision.');
  var lock = arLock_(); lock.waitLock(28000);
  var out;
  try {
    var res = computeLight_('memo'), P = res.params, me = email_(), u = ENGINE.userBy(P, me);
    var A = apprRows_(), r = A.rows.filter(function (x) { return String(x.id) === String(id); })[0];
    if (!r) throw new Error('Request ' + id + ' not found.');
    var kind = Object.keys(ENGINE.DOC).filter(function (k) { return ENGINE.DOC[k] === String(r.jenis).toUpperCase(); })[0];
    // delegasi sementara (Parameter DELEGASI_*): diputuskan a.n. Section Head, tercatat jelas sebagai delegasi
    var dl = (!u || u.peran.indexOf('PENYETUJU') !== 0) ? ENGINE.delegasiFor(P, me, kind, res.ctx.today) : null;
    if ((!u || u.peran.indexOf('PENYETUJU') !== 0) && !dl) throw new Error('Only users with the PENYETUJU role (Section Head) can decide' + (String(P.DELEGASI_EMAIL || '').toLowerCase() === me ? ' (your delegation does not cover this type or expired on ' + P.DELEGASI_SAMPAI + ')' : '') + '. Your account: ' + (me || '(unknown)'));
    if (String(r.status).toUpperCase() !== 'MENUNGGU') throw new Error('Request ' + id + ' already has status ' + r.status + '.');
    // pemisahan tugas: berlaku juga untuk delegasi — pengaju tidak pernah memutuskan pengajuannya sendiri
    if (String(r.oleh).toLowerCase() === me) throw new Error('A request cannot be approved by its own preparer.');
    var kTglD = arGuardClosed_(kind, kind === 'BATAL' ? ENGINE.packageFor(res.ctx, kind, r.ref) : { ref: r.ref }, 'Decision ' + r.id);
    if (dl) catatan = '[DELEGASI a.n. ' + dl.atas + ' s.d. ' + dl.sampai + (dl.dasar ? ' · dasar: ' + dl.dasar : '') + ']' + (String(catatan || '').trim() ? ' ' + String(catatan).trim() : '');
    if (kind === 'INV' || kind === 'PAY' || kind === 'KP' || kind === 'TUTUP' || kind === 'BATAL' || kind === 'VALX') {
      var d = ENGINE.packageFor(res.ctx, kind, r.ref);
      if (!d || d.fp !== String(r.sidik)) {
        A.sh.getRange(r._row, A.map.status).setValue('DIBATALKAN'); A.sh.getRange(r._row, A.map.catatan).setValue('Data berubah sebelum diputuskan — perlu diajukan ulang');
        throw new Error('Data for ' + r.ref + ' has changed since submission. The request was cancelled; ask the preparer to resubmit.');
      }
    }
    var now = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'), st = String(keputusan).toUpperCase() === 'SETUJU' ? 'DISETUJUI' : 'DITOLAK';
    if (st === 'DITOLAK' && !String(catatan || '').replace(/^\[DELEGASI[^\]]*\]/, '').trim()) throw new Error('Enter a rejection reason.');
    // v2.5: validasi rencana cair — keputusan per invoice ikut disimpan & ditandatangani (kolom SIDIK = sidik data / sidik keputusan)
    var data = null; try { data = JSON.parse(String(r.data || '')); } catch (e) { data = null; }
    if (kind === 'VAL' && st === 'DISETUJUI' && data && data.v === 2) {
      var err = ENGINE.valApplyDecisions(data, keputusanItem || []);
      if (err) throw new Error(err);
      r.sidik = ENGINE.valSidik(data.fp0 || String(r.sidik).split('/')[0], data.items);
      A.sh.getRange(r._row, A.map.data).setNumberFormat('@').setValue(JSON.stringify(data));
      A.sh.getRange(r._row, A.map.sidik).setNumberFormat('@').setValue(r.sidik);
      if (A.map.nilai) A.sh.getRange(r._row, A.map.nilai).setNumberFormat(NUMF.money).setValue(Math.round(data.nilai || 0));
      // revisi sebelumnya untuk bulan yang sama: tetap sah sebagai riwayat, ditandai digantikan
      A.rows.filter(function (x) { return x.id !== r.id && String(x.jenis).toUpperCase() === ENGINE.DOC.VAL && String(x.ref) === String(r.ref) && String(x.status).toUpperCase() === 'DISETUJUI'; })
        .forEach(function (x) { if (!/DIGANTI oleh/.test(String(x.catatan))) A.sh.getRange(x._row, A.map.catatan).setNumberFormat('@').setValue((String(x.catatan || '') + ' · DIGANTI oleh ' + r.id).replace(/^ · /, '')); });
    }
    // v2.7: pembatalan lembar validasi — lembar lama bulan itu DIBATALKAN (tidak dihapus), PDF dipindah ke subfolder "Dibatalkan"
    if (kind === 'VALX' && st === 'DISETUJUI' && data) {
      var ids = {}; (data.list || []).forEach(function (x) { ids[x.id] = 1; });
      A.rows.filter(function (x) { return ids[String(x.id)] && String(x.jenis).toUpperCase() === ENGINE.DOC.VAL; }).forEach(function (x) {
        A.sh.getRange(x._row, A.map.status).setValue('DIBATALKAN');
        A.sh.getRange(x._row, A.map.catatan).setNumberFormat('@').setValue((String(x.catatan || '') + ' · DIBATALKAN oleh ' + r.id + ': ' + data.alasan).replace(/^ · /, ''));
      });
      try { moveValPdfs_(P, r.ref); } catch (e) { }
    }
    r.status = st; r.putus_oleh = me; r.putus_waktu = now; r.catatan = String(catatan || '');
    r.kode = sign_({ id: r.id, jenis: String(r.jenis).toUpperCase(), ref: r.ref, sidik: r.sidik, putus_oleh: me, putus_waktu: now, status: st });
    ['status', 'putus_oleh', 'putus_waktu', 'catatan', 'kode'].forEach(function (k) { if (A.map[k]) A.sh.getRange(r._row, A.map[k]).setNumberFormat('@').setValue(r[k]); });
    out = { ok: true, id: r.id, status: st, kode: r.kode };
    var notif = { r: r, me: me, nama: (u && u.nama) || '', dl: dl, koreksi: kTglD };
  } finally { lock.releaseLock(); }
  // bukti di luar sistem: kode HMAC bisa dihitung ulang oleh editor proyek, email di kotak masuk Section Head tidak
  out.warnings = [];
  var w = arNotifyDecision_(notif); if (w) out.warnings.push(w);
  // v2.8.6: laporan diperbarui di latar belakang; koreksi periode tertutup tetap langsung (snapshot analisis bulanan ikut direvisi)
  if (kTglD && out.status === 'DISETUJUI') rebuildAll(); else out.background = arQueueAfterSave_().background;
  if (kTglD && out.status === 'DISETUJUI') { var rf = arRefreshClosedMonths_([kTglD], out.id + ' ' + (notif.r.catatan || notif.r.label || '')); out.warnings = out.warnings.concat(rf.warnings); out.snapshots = rf.ids; }
  return out;
}
/** menu: lokal file Indonesia (ketikan 05/10/2026 = 5 Oktober) + validasi "harus tanggal" di kolom tanggal yang diisi manual */
function aturFormatTanggal() {
  arRequireMaker_();
  var ui = SpreadsheetApp.getUi(), ss = arSpreadsheet_(), lama = ss.getSpreadsheetLocale();
  if (ui.alert('Format tanggal Indonesia', 'Lokal file sekarang: ' + lama + '.\n\nSetelah diubah ke Indonesia:\n• Ketikan 05/10/2026 dibaca 5 Oktober 2026 (tanggal/bulan/tahun) dan tampil 05-Okt-2026.\n• Angka desimal diketik dengan KOMA (mis. 0,0265 atau 2,65%), ribuan dengan titik. Nilai yang sudah ada tidak berubah.\n• Kolom tanggal yang diisi manual hanya menerima tanggal yang sah.\n\nLanjutkan?', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  var ok = '';
  ['id_ID', 'in_ID'].some(function (lc) { try { ss.setSpreadsheetLocale(lc); if (/^(id|in)/i.test(ss.getSpreadsheetLocale())) { ok = ss.getSpreadsheetLocale(); return true; } } catch (e) { } return false; });
  var help = 'Ketik tanggal dengan format tanggal/bulan/tahun, mis. 05/10/2026', n = 0;
  var rule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).setHelpText(help).build();
  [[ENGINE.REG, ['submit', 'tt', 'tgl_fp', 'rencana', 'ba_tgl'], ENGINE.REG_ALIAS], [ENGINE.LOG, ['tgl']], [ENGINE.ADD, ['tgl']], [ENGINE.TAG, ['tgl']], [ENGINE.MPRO, ['tgl_kontrak'], ENGINE.MPRO_ALIAS]].forEach(function (x) {
    var sh = ss.getSheetByName(x[0].sheet); if (!sh) return;
    var map = colMap_(sh, x[0], x[2] || null), rows = Math.max(sh.getMaxRows() - x[0].firstRow + 1, 1);
    x[1].forEach(function (k) { if (!map[k]) return; sh.getRange(x[0].firstRow, map[k], rows, 1).setDataValidation(rule).setNumberFormat(NUMF.date); n++; });
  });
  ui.alert('Selesai', (ok ? 'Lokal file: ' + ok + ' (Indonesia).' : 'Lokal tidak bisa diubah otomatis — ubah manual: File ▸ Setelan ▸ Lokal: Indonesia.') + '\n' + n + ' kolom tanggal kini hanya menerima tanggal sah (tampil dd-mmm-yyyy).\n\nTanggal yang terlanjur tertukar ditandai di sheet Cek Data (kategori TANGGAL) setelah laporan diperbarui.', ui.ButtonSet.OK);
  rebuildAll();
}
/** menu: buat / temukan folder arsip, simpan ID di Parameter, bagikan ke PEMBUAT & PENYETUJU (sekali, oleh pemilik) */
function setupArsipFolder() {
  arRequireMaker_();
  var P = arParameters_(), ui = SpreadsheetApp.getUi(), x = paramKV_('ARSIP_FOLDER_ID'), root;
  if (x && x.val) { try { root = DriveApp.getFolderById(x.val); } catch (e) { root = null; } }
  if (!root) { var it = DriveApp.getFoldersByName(String(P.ARSIP_FOLDER || 'Arsip Paket AR CSU')); root = it.hasNext() ? it.next() : DriveApp.createFolder(String(P.ARSIP_FOLDER || 'Arsip Paket AR CSU')); }
  setParamKV_('ARSIP_FOLDER_ID', root.getId(), 'ID folder arsip PDF (dibuat otomatis, dibagikan ke PEMBUAT & PENYETUJU)');
  var have = {}, add = []; root.getEditors().forEach(function (u) { have[String(u.getEmail()).toLowerCase()] = 1; });
  P.users.forEach(function (u) { if (u.email && /^(PEMBUAT|PENYETUJU)/.test(u.peran) && !have[u.email]) { try { root.addEditor(u.email); add.push(u.email); have[u.email] = 1; } catch (e) { } } });
  ui.alert('Folder arsip siap', 'Folder: ' + root.getName() + '\n' + root.getUrl() + '\n\nID tersimpan di Parameter ▸ ARSIP_FOLDER_ID.' + (add.length ? '\nDibagikan ke: ' + add.join(', ') : ''), ui.ButtonSet.OK);
}
/** folder arsip: <ARSIP_FOLDER>/<jenis paket>/<tahun>/<MM Bulan tahun> */
function archiveFolder_(P, kind, month) {
  var folderName = String(P.ARSIP_FOLDER || 'Arsip Paket AR CSU'), sub = function (parent, name) { var it = parent.getFoldersByName(name); return it.hasNext() ? it.next() : parent.createFolder(name); };
  var root = null, x = paramKV_('ARSIP_FOLDER_ID');
  if (x && x.val) { try { root = DriveApp.getFolderById(x.val); } catch (e) { throw new Error('This account cannot open the archive folder — ask the folder owner to share it, or clear ARSIP_FOLDER_ID in Parameter.'); } }
  if (!root) {
    var it = DriveApp.getFoldersByName(folderName); root = it.hasNext() ? it.next() : DriveApp.createFolder(folderName);
    // ID wajib tersimpan: tanpa itu setiap akun membuat folder arsip sendiri di Drive-nya dan arsip tercecer
    try { setParamKV_('ARSIP_FOLDER_ID', root.getId(), 'ID folder arsip PDF (dibuat otomatis, dibagikan ke PEMBUAT & PENYETUJU)'); }
    catch (e) { throw new Error('The archive folder is not set up. Spreadsheet owner: run menu AR CSU ▸ Siapkan folder arsip PDF, then try again.'); }
    // folder dibagikan ke pembuat & penyetuju agar arsip otomatis dari siapa pun masuk ke folder yang sama
    try { var have = {}; root.getEditors().forEach(function (u) { have[String(u.getEmail()).toLowerCase()] = 1; }); P.users.forEach(function (u) { if (u.email && /^(PEMBUAT|PENYETUJU)/.test(u.peran) && !have[u.email]) { root.addEditor(u.email); have[u.email] = 1; } }); } catch (e) { }
  }
  var subName = { INV: 'Paket Invoice', PAY: 'Paket Penerimaan', VAL: 'Validasi Rencana Cair', REK: 'Rekap Bulanan Akuntansi', KP: 'Kartu Piutang', TUTUP: 'Penutupan Manual', BATAL: 'Pembatalan Invoice', VALX: 'Validasi Rencana Cair', SURAT: 'Surat Penagihan', DOK: 'Berkas Tagihan & Faktur Pajak' }[kind] || 'Lainnya';
  var f = sub(root, subName), m = /^\d{4}-\d{2}/.test(String(month || '')) ? String(month) : Utilities.formatDate(new Date(), TZ, 'yyyy-MM'), y = m.substring(0, 4), mm = +m.substring(5, 7);
  return sub(sub(f, y), ('0' + mm).slice(-2) + ' ' + ENGINE.MONTHS[mm - 1] + ' ' + y);
}
/**
 * simpan arsip PDF ke Google Drive.
 * meta = { b64 (PDF dari aplikasi) | html (cadangan), name, month 'yyyy-mm', prefix (awal nama file untuk memindahkan versi lama ke "Revisi lama") }
 */
function apiArchive(kind, ref, meta, fileName) { arRequireAccess_();
  if (typeof meta === 'string') meta = { html: meta, name: fileName };
  var res = computeLight_(), P = res.params, folder = archiveFolder_(P, kind, meta.month);
  var name = String(meta.name || (kind + ' ' + ref)).replace(/[\\\/:*?"<>|]/g, '-').replace(/\s+/g, ' ').trim();
  if (!/\.pdf$/i.test(name)) name += '.pdf';
  var blob = meta.b64 ? Utilities.newBlob(Utilities.base64Decode(meta.b64), 'application/pdf', name) : HtmlService.createHtmlOutput(meta.html || '').getAs('application/pdf').setName(name);
  // versi lama (awalan sama) dipindah ke subfolder "Revisi lama" agar folder bulan hanya berisi versi terbaru
  var moved = 0;
  if (meta.prefix) {
    var files = folder.getFiles(), old = null;
    while (files.hasNext()) { var f0 = files.next(); if (String(f0.getName()).indexOf(meta.prefix) === 0) { if (!old) { var it2 = folder.getFoldersByName('Revisi lama'); old = it2.hasNext() ? it2.next() : folder.createFolder('Revisi lama'); } f0.moveTo(old); moved++; } }
  }
  var f = folder.createFile(blob);
  if (kind === 'INV') setRegisterValue_(ref, 'link_paket', f.getUrl());
  return { ok: true, url: f.getUrl(), name: f.getName(), folder: folder.getName(), moved: moved };
}

// =====================================================================  DOCUMENT FILING (v2.8.5)
/** nama file mengikuti kebiasaan AR: "4C.0183.26 - INV Prog 4 - Metland Perbaikan Struktur Irigasi Sign.Rev.00.pdf" / "4C.0183.26 - FP Prog 4 - … .pdf" */
function docFileName_(inv, kind, ext) {
  var m = String(inv.nomor).match(/4C\.(\d{4}[A-Z]?)\/(\d{4})/), id = m ? '4C.' + m[1] + '.' + m[2].substring(2) : String(inv.nomor).replace(/[^\w.]/g, '');
  var ke = String(inv.invke || '').replace(/Progres\s+/gi, 'Prog ').replace(/^Down Payment\/?\s*UM$/i, 'UM').replace(/^Termin\s+/i, 'Termin ').split('(')[0].trim() || 'Tagihan';
  var pro = String((inv._pro && (inv._pro.singkat || inv._pro.nama_pe)) || inv.npp).replace(/\s*-\s*/g, ' ').toLowerCase().replace(/\b\w/g, function (c) { return c.toUpperCase(); });
  var base = id + ' - ' + (kind === 'FP' ? 'FP' : 'INV') + ' ' + ke + ' - ' + pro + (kind === 'FP' ? '' : ' Sign.Rev.' + ('0' + (Number(inv.rev) || 0)).slice(-2));
  return { prefix: id + ' - ' + (kind === 'FP' ? 'FP' : 'INV') + ' ', name: base.replace(/[\\\/:*?"<>|]/g, '-').replace(/\s+/g, ' ').trim() + '.' + ext };
}
/** unggah berkas tagihan ber-TTD (PACK) atau PDF faktur pajak (FP) untuk satu invoice.
 *  m = { b64, mime, size, fp: { no, tgl, ref, harga, ppn } (hasil baca PDF Coretax di aplikasi), force } */
function apiFileDoc(nomor, kind, m) {
  arRequireMaker_(); m = m || {};
  if (['PACK', 'FP'].indexOf(kind) < 0) throw new Error('Unknown document type.');
  var mime = String(m.mime || ''), ext = mime === 'application/pdf' ? 'pdf' : mime === 'image/jpeg' ? 'jpg' : mime === 'image/png' ? 'png' : '';
  if (!ext || (kind === 'FP' && ext !== 'pdf')) throw new Error(kind === 'FP' ? 'The tax invoice must be the PDF from Coretax.' : 'Only PDF, JPG or PNG files can be filed.');
  var bytes = Utilities.base64Decode(String(m.b64 || ''));
  if (!bytes.length) throw new Error('The file is empty.');
  if (bytes.length > 10 * 1024 * 1024) throw new Error('The file is larger than 10 MB. Scan at 150–200 dpi or compress the PDF, then try again.');
  var res = computeLight_(), ctx = res.ctx, key = ENGINE.normInv(nomor), inv = ctx.invs.filter(function (i) { return ENGINE.normInv(i.nomor) === key && !i._batal; })[0];
  if (!inv) throw new Error('Invoice ' + nomor + ' not found (or cancelled).');
  var warn = [], fp = m.fp || null;
  if (kind === 'FP' && fp) {
    if (fp.ref && ENGINE.normInv(String(fp.ref).split(/\s+-\s+/)[0]) !== key && !m.force) throw new Error('REF_MISMATCH: This tax invoice refers to ' + fp.ref + ', not ' + inv.nomor + '.');
    if (fp.harga && Math.abs(Number(fp.harga) - inv._a.dpp) > 1) warn.push('Selling price on the tax invoice (' + ENGINE.fmtRp(fp.harga) + ') differs from the invoice tax base (' + ENGINE.fmtRp(inv._a.dpp) + ').');
    if (fp.ppn && Math.abs(Number(fp.ppn) - inv._a.ppn) > 1) warn.push('VAT on the tax invoice (' + ENGINE.fmtRp(fp.ppn) + ') differs from the invoice VAT (' + ENGINE.fmtRp(inv._a.ppn) + ').');
  }
  var lock = arLock_(); lock.waitLock(28000);
  try {
    var folder = archiveFolder_(res.params, 'DOK', String(inv.tgl).substring(0, 7)), fn = docFileName_(inv, kind, ext), moved = 0, old = null, files = folder.getFiles();
    while (files.hasNext()) { var f0 = files.next(); if (String(f0.getName()).indexOf(fn.prefix) === 0) { if (!old) { var it = folder.getFoldersByName('Revisi lama'); old = it.hasNext() ? it.next() : folder.createFolder('Revisi lama'); } f0.moveTo(old); moved++; } }
    var f = folder.createFile(Utilities.newBlob(bytes, mime, fn.name)), sh = sheet_('Monitoring Invoice'), map = colMap_(sh, ENGINE.REG, ENGINE.REG_ALIAS), row = findRegisterRow_(sh, inv.nomor, map), was = {};
    if (!row) throw new Error('Invoice row not found in Monitoring Invoice.');
    var set = function (k, v, fmt) { if (!map[k]) return; var c = sh.getRange(row, map[k]); was[k] = c.getDisplayValue(); if (fmt) c.setNumberFormat(fmt); c.setValue(v); };
    set(kind === 'FP' ? 'link_fp' : 'link_inv', f.getUrl());
    if (kind === 'FP' && fp && /^\d{16,17}$/.test(String(fp.no || ''))) {
      var nf = String(fp.no); set('no_fp', nf, '@'); if (was.no_fp && was.no_fp.replace(/\D/g, '') !== nf.replace(/\D/g, '')) warn.push('Tax invoice no. in the register (' + was.no_fp + ') was replaced by ' + nf + ' from the PDF.');
      if (/^\d{4}-\d{2}-\d{2}$/.test(String(fp.tgl || ''))) { var p = fp.tgl.split('-'); set('tgl_fp', dateAt_(+p[0], +p[1], +p[2]), 'dd/mm/yyyy'); }
    }
    // tercatat lewat aplikasi (bukan edit langsung), jadi tidak masuk Jejak Edit Langsung
    arMarkDirty_();
    return plain_({ ok: true, url: f.getUrl(), name: fn.name, folder: folder.getName(), moved: moved, warnings: warn });
  } finally { lock.releaseLock(); }
}

// =====================================================================  SURAT PERMOHONAN PERCEPATAN PEMBAYARAN
/** invoice lewat jatuh tempo milik satu pelanggan, dikelompokkan per proyek (untuk menyusun surat di aplikasi) */
function apiSuratData(pelanggan) { arRequireMaker_(); return plain_(ENGINE.suratData(computeLight_().ctx, pelanggan)); }
/** arsipkan PDF surat ke Drive lalu catat di Log Surat Penagihan (nomor surat unik; invoice diperiksa ulang di server) */
function apiSuratSimpan(m) {
  var a = arRequireMaker_(); m = m || {};
  // mode: 'baru' · 'revisi' (belum dikirim: nomor sama, Rev.N) · 'ganti' (sudah sampai pelanggan: nomor + huruf, mis. 212A)
  var mode = ['revisi', 'ganti'].indexOf(m.mode) >= 0 ? m.mode : 'baru', alasan = String(m.alasan || '').trim();
  var nomor = String(m.nomor || '').trim(), tgl = String(m.tgl || ''), tenggat = String(m.tenggat || ''), list = (m.invoices || []).map(function (x) { return ENGINE.normInv(x); });
  if (nomor.length < 5) throw new Error('Letter number is required.');
  if (!arValidDate_(tgl)) throw new Error('Invalid letter date.');
  if (tenggat && (!arValidDate_(tenggat) || tenggat < tgl)) throw new Error('The payment deadline cannot be before the letter date.');
  if (!list.length) throw new Error('Select at least one invoice.');
  if (!m.b64) throw new Error('The letter PDF has not been generated.');
  if (mode !== 'baru' && alasan.length < 10) throw new Error('Enter a ' + (mode === 'revisi' ? 'revision' : 'replacement') + ' reason of at least 10 characters.');
  var allNo = function () { return readSchema_(sheet_(ENGINE.SURAT.sheet), ENGINE.SURAT).map(function (r) { return String(r[0]).trim(); }).filter(String); };
  var used = function (n) { return allNo().some(function (x) { return x.toUpperCase() === n.toUpperCase(); }); };
  var d = ENGINE.suratData(computeLight_().ctx, m.pelanggan), ok = {}, total = 0, ke = 1, ruj = {}, asal = null;
  if (mode !== 'baru') {
    asal = d.letters.filter(function (l) { return l.nomor.toUpperCase() === String(m.asal || '').trim().toUpperCase(); })[0];
    if (!asal) throw new Error('Original letter ' + m.asal + ' not found for ' + d.nama + '.');
    if (asal.tutup || asal.ganti) throw new Error('Letter ' + asal.nomor + ' has already been ' + (asal.ganti ? 'replaced by letter ' + asal.ganti : 'closed') + '.');
    if (mode === 'revisi' && nomor.toUpperCase() !== asal.nomor.toUpperCase()) throw new Error('A revision of an unsent letter must keep the same number (' + asal.nomor + ').');
    if (mode === 'ganti') { var exp = ENGINE.suratNoGanti(asal.nomor, arParameters_().SURAT_NOMOR_FORMAT, allNo()); if (!exp || nomor.toUpperCase() !== exp.toUpperCase()) throw new Error('The replacement letter number should be ' + (exp || '(original number format not recognized)') + '. Reload the app.'); }
  } else if (sheet_(ENGINE.SURAT.sheet) && used(nomor)) throw new Error('Letter number ' + nomor + ' is already recorded in the Log Surat Penagihan sheet.');
  d.projects.forEach(function (p) { p.invoices.forEach(function (i) { ok[ENGINE.normInv(i.nomor)] = i; }); });
  list.forEach(function (n) {
    if (!ok[n]) throw new Error('Invoice ' + n + ' is not overdue or does not belong to ' + d.nama + '. Reload the app.');
    total += ok[n].sisa; ke = Math.max(ke, Math.min(3, ok[n].stage.ke || 1));
    (ok[n].stage.chain || []).forEach(function (c) { ruj[c.nomor] = c; });
  });
  var rujukan = Object.keys(ruj).map(function (k) { return ruj[k]; }).sort(function (a, b) { return a.tgl < b.tgl ? -1 : 1; }).map(function (c) { return c.nomor + ' (' + c.tgl + ')'; }).join('; ');
  // revisi / pengganti tetap di tingkat surat asal (tidak naik ke surat berikutnya) dan merujuk surat yang sama
  if (asal) { ke = asal.ke; rujukan = asal.rujukan; }
  if (m.ke && Number(m.ke) !== ke) throw new Error('Previous letter data has changed (this should be letter no. ' + ke + '). Reload the app and check again.');
  var base = 'Surat Penagihan_' + d.nama + '_' + nomor, rev = mode === 'revisi' ? (asal.rev || 0) + 1 : 0;
  // revisi: PDF lama (awalan nama sama) dipindah ke subfolder "Revisi lama"
  var ar = apiArchive('SURAT', nomor, { b64: m.b64, name: base + (rev ? ' Rev.' + rev : ''), month: tgl.slice(0, 7), prefix: mode === 'revisi' ? base.replace(/[\\\/:*?"<>|]/g, '-').replace(/\s+/g, ' ').trim() : '' });
  var rec = { nomor: nomor, tgl: tgl, pelanggan: d.pelanggan, invoices: list.join('; '), sisa: Math.round(total), tenggat: tenggat, kepada: String(m.kepada || '').slice(0, 500),
    rekening: (m.rekening || []).join('; '), kontak: String(m.kontak || '').slice(0, 300), oleh: a.email, waktu: arNow_(), link: ar.url, ke: ke, rujukan: rujukan };
  if (mode === 'revisi') {
    rec.rev = rev; rec.rev_ket = ((asal.rev_ket ? asal.rev_ket + ' · ' : '') + 'Rev.' + rev + ' ' + arNow_() + ' ' + a.email + ': ' + alasan).slice(0, 2000);
    delete rec.nomor; suratUpdate_(asal.nomor, rec);
  } else {
    if (asal) rec.menggantikan = asal.nomor + ' (' + asal.tgl + ') — ' + alasan;
    var lock = arLock_(); lock.waitLock(28000);
    try {
      var ss = arSpreadsheet_(), sh = setupDataSheet_(ss, ENGINE.SURAT, 'LOG SURAT PERMOHONAN PERCEPATAN PEMBAYARAN', 'Dicatat otomatis saat surat dibuat di aplikasi ▸ Surat Penagihan. Dasar alert pelanggan & peringatan TENGGAT SURAT TERLEWAT. Jangan diubah manual.', 'Log Persetujuan', '#C55A11').sh;
      if (used(nomor)) throw new Error('Letter number ' + nomor + ' is already recorded in the Log Surat Penagihan sheet.');
      var map = colMap_(sh, ENGINE.SURAT), lc = Math.max(sh.getLastColumn(), ENGINE.SURAT.cols.length), row = Array(lc).fill('');
      ENGINE.SURAT.cols.forEach(function (c) { var j = map[c[0]]; if (j) row[j - 1] = arSafeCell_(toCell_(rec[c[0]], colFmt_(c))); });
      var at = lastDataRow_(sh, ENGINE.SURAT.firstRow, [map.nomor]) + 1; roomFor_(sh, at);
      sh.getRange(at, 1, 1, lc).setValues([row]);
    } finally { lock.releaseLock(); }
    if (asal) suratUpdate_(asal.nomor, { ganti: nomor }); else { try { rebuildAll_(); } catch (e) { } }
  }
  var warn = mode === 'baru' ? '' : arNotify_((mode === 'revisi' ? 'Revisi surat ' + nomor + ' Rev.' + rev : 'Surat ' + asal.nomor + ' diganti ' + nomor) + ' — ' + d.nama,
    ['Pelanggan: ' + d.nama, mode === 'revisi' ? 'Surat ' + nomor + ' direvisi sebelum dikirim (Rev.' + rev + ').' : 'Surat ' + asal.nomor + ' tanggal ' + asal.tgl + ' (sudah diterima pelanggan) diganti surat ' + nomor + ' tanggal ' + tgl + '.',
      'Alasan: ' + alasan, 'Total sisa tagihan: Rp ' + ENGINE.fmtRp(total), 'Dibuat oleh: ' + a.email, 'PDF: ' + ar.url]);
  return plain_({ ok: true, url: ar.url, name: ar.name, folder: ar.folder, total: total, ke: ke, mode: mode, rev: rev, warnings: warn ? [warn] : [] });
}
/** ubah kolom tindak lanjut satu surat di Log Surat Penagihan (dicari dari nomor surat) */
function suratUpdate_(nomor, fields) {
  var lock = arLock_(); lock.waitLock(28000);
  try {
    var sh = sheet_(ENGINE.SURAT.sheet); if (!sh) throw new Error('The Log Surat Penagihan sheet does not exist yet.');
    setupDataSheet_(arSpreadsheet_(), ENGINE.SURAT, 'LOG SURAT PERMOHONAN PERCEPATAN PEMBAYARAN', 'Dicatat otomatis saat surat dibuat di aplikasi ▸ Surat Penagihan. Dasar alert pelanggan & peringatan TENGGAT SURAT TERLEWAT. Jangan diubah manual.', 'Log Persetujuan', '#C55A11');
    var map = colMap_(sh, ENGINE.SURAT), last = sh.getLastRow(), f0 = ENGINE.SURAT.firstRow;
    var nos = last >= f0 ? sh.getRange(f0, map.nomor, last - f0 + 1, 1).getValues() : [], at = -1;
    nos.forEach(function (r, i) { if (String(r[0]).trim().toUpperCase() === String(nomor).trim().toUpperCase()) at = f0 + i; });
    if (at < 0) throw new Error('Letter ' + nomor + ' not found in the Log Surat Penagihan sheet.');
    if (String(sh.getRange(at, map.tutup).getValue()).trim()) throw new Error('Letter ' + nomor + ' is already closed.');
    if (map.ganti && String(sh.getRange(at, map.ganti).getValue()).trim()) throw new Error('Letter ' + nomor + ' has already been replaced by letter ' + sh.getRange(at, map.ganti).getValue() + '.');
    ENGINE.SURAT.cols.forEach(function (c) { if (fields[c[0]] !== undefined && map[c[0]]) sh.getRange(at, map[c[0]]).setNumberFormat(colFmt_(c) === 'date' ? NUMF.date : '@').setValue(arSafeCell_(toCell_(fields[c[0]], colFmt_(c)))); });
  } finally { lock.releaseLock(); }
  try { rebuildAll_(); } catch (e) { }
  return plain_({ ok: true });
}
/** tanggapan pelanggan (surat balasan / email / telepon / rapat), opsional janji bayar → surat berikutnya ditahan sampai janji */
function apiSuratTanggapan(nomor, t) {
  var a = arRequireMaker_(); t = t || {};
  var tgl = String(t.tgl || ''), janji = String(t.janji || ''), isi = String(t.isi || '').trim(), jenis = String(t.jenis || '').trim().toUpperCase();
  if (!arValidDate_(tgl) || tgl > arToday_()) throw new Error('Invalid or future response date.');
  if (['SURAT BALASAN', 'EMAIL', 'TELEPON', 'RAPAT', 'LAINNYA'].indexOf(jenis) < 0) throw new Error('Select the response type.');
  if (isi.length < 5) throw new Error('Enter a response summary.');
  if (janji && (!arValidDate_(janji) || janji < tgl)) throw new Error('The promised payment date cannot be before the response date.');
  return suratUpdate_(nomor, { tanggap_tgl: tgl, tanggap_jenis: jenis, tanggap_isi: (isi + ' · dicatat ' + a.email).slice(0, 500), janji: janji });
}
/** hentikan penagihan lewat surat untuk invoice di surat ini (sengketa, rekonsiliasi, invoice dibatalkan/diganti) */
function apiSuratTutup(nomor, alasan) {
  var a = arRequireMaker_(); alasan = String(alasan || '').trim();
  if (alasan.length < 10) throw new Error('The reason for closing the letter must be at least 10 characters.');
  return suratUpdate_(nomor, { tutup: alasan.slice(0, 500), tutup_oleh: a.email + ' · ' + arNow_() });
}

// =====================================================================  UPGRADE FILE VERSI 1 → VERSI 2
var V2_AKUN = [
  ['PIUTANG_USAHA', 'Piutang Usaha', '1104111000', '1104211000', 'Debit saat invoice (Netto), kredit saat penerimaan'],
  ['PIUTANG_PPN', 'Piutang PPN K & PPnBM', '1104112000', '1104212000', 'Debit saat invoice (PPN), kredit saat PPN diterima'],
  ['PIUTANG_RETENSI', 'Piutang Retensi', '1105111000', '1105211000', 'Retensi yang ditahan pelanggan'],
  ['TAGBRUT', 'Pendapatan Yang Diperhitungkan (Tagihan Bruto)', '1109111000', '1109121000', 'Dikredit saat invoice progres terbit (nilai bruto)'],
  ['HAD', 'Hasil Akan Diterima (Piutang Prestasi)', '1114111000', '1114121000', 'Dipakai bila AKUN PENGAKUAN proyek = HAD'],
  ['PPN_KELUARAN', 'PPN K / Kurang Bayar', '2105231000', '2105231000', 'Kredit PPN saat invoice'],
  ['UANG_MUKA_PROYEK', 'Uang Muka Proyek', '2106111000', '2106111000', 'Pengembalian uang muka di invoice progres'],
  ['PPH_FINAL', 'PPh Final Dibayar Di Muka', '1110181000', '1110181000', 'PPh final 4(2) (JT/JF) yang dipotong pelanggan'],
  ['PPH_23', 'PPh Pasal 23 Wajib Bayar', '1110141000', '1110141000', 'PPh 23 sewa alat (JL)'],
  ['BANK', 'Bank Mandiri 166-0004-4934-17', '1101211187', '1101211187', 'Rekening penerimaan unit'],
  ['BEBAN_ADM_BANK', 'Proses Penagihan Piutang Usaha', '6401140001', '6401140001', 'MOHON KONFIRMASI Akuntansi (alternatif 6311071003)'],
  ['UTANG_KOMPENSASI', 'Utang Usaha (kompensasi dengan piutang)', '', '', 'MOHON KONFIRMASI Akuntansi — akun utang ke pelanggan yang dikompensasi (mis. klaim solar)'],
  ['CKPN_PIUTANG_USAHA', 'Akumulasi Penyisihan Piutang Usaha', '1104119000', '1104219000', ''],
  ['CKPN_RETENSI', 'Cad. Kerugian Penurunan Nilai Piutang Retensi', '1105112000', '1105212000', ''],
  ['CKPN_TAGBRUT', 'Cad. KPN Pendapatan Yang Diperhitungkan', '1109112000', '1109122000', ''],
  ['CKPN_HAD', 'Cad. Kerugian Penurunan Nilai HAD', '1114112000', '1114122000', ''],
  ['BEBAN_CKPN_PIUTANG_USAHA', 'Beban CKPN Piutang Usaha', '8106111000', '8106111001', ''],
  ['BEBAN_CKPN_RETENSI', 'Beban CKPN Piutang Retensi', '8106121000', '8106121001', ''],
  ['BEBAN_CKPN_TAGBRUT', 'Beban CKPN Tagihan Bruto', '8106131000', '8106131001', ''],
  ['BEBAN_CKPN_HAD', 'Beban CKPN Piutang Prestasi', '8106131010', '8106131011', ''],
  ['PEMULIHAN_CKPN_PIUTANG_USAHA', 'Pemulihan Impairment Piutang Usaha', '8199221000', '8199221001', ''],
  ['PEMULIHAN_CKPN_RETENSI', 'Pemulihan Impairment Piutang Retensi', '8199222000', '8199222001', ''],
  ['PEMULIHAN_CKPN_TAGBRUT', 'Pemulihan Impairment Tagihan Bruto', '8199223000', '8199223010', ''],
  ['PEMULIHAN_CKPN_HAD', 'Pemulihan Impairment Piutang Prestasi', '8199223001', '8199223011', '']];
// templat awal tabel Pengguna & Peran: email diisi di Parameter (baris PEMBUAT = akun yang memasang), bukan di kode
var V2_USERS = [
  ['', 'Fikri Zikra Akrama', 'AR Officer', 'TX220018', 'PEMBUAT'],
  ['', 'Anas Afifudin', 'Section Head of Finance & HR', 'WB170219', 'PENYETUJU'],
  ['', 'Yusqi Audah Firdaus', 'Manager of Concrete Services Unit', 'WB090278', 'MENGETAHUI (TTD BASAH)'],
  ['', 'Elmarta Hexa Panggabean', 'Akuntansi', '', 'PENERIMA']];
var V2_KV = [
  ['JT_AKUNTANSI_HARI', 30, 'Jatuh tempo versi akuntansi = tanggal invoice + n hari (CKPN, sama dengan lampiran LK N/30)'],
  ['CKPN_DISKONTO', 0.0833, 'Suku diskonto NPV untuk CKPN retensi, tagbrut & HAD (per tahun)'],
  ['CKPN_PIHAK_BERELASI', 'TIDAK', 'YA = pihak berelasi ikut dicadangkan; TIDAK = 0'],
  ['TAGBRUT_SALDO_AWAL_PER', '2026-09-30', 'Tanggal saldo awal tagbrut/HAD di Input Tagbrut'],
  ['TAGBRUT_BATAS_HARI', 30, 'Pengakuan belum ditagih > n hari ditandai SEGERA TERBITKAN TAGIHAN'],
  ['FAKTUR_KODE_TRANSAKSI', '04', 'Kode transaksi faktur Coretax (04 = DPP Nilai Lain)'],
  ['FAKTUR_KODE_JASA', '', 'Kode barang/jasa Coretax bawaan (bila Master Proyek kosong)'],
  ['FAKTUR_SATUAN', 'Lainnya', 'Satuan barang/jasa di Coretax'],
  ['FAKTUR_CEK_MULAI', '2026-09-01', 'Pengingat faktur belum dicatat hanya untuk invoice sejak tanggal ini'],
  ['FAKTUR_PENGINGAT_HARI', 10, 'Pengingat muncul bila faktur belum dicatat n hari setelah invoice'],
  ['SAP_COMPANY_CODE', 'B000', 'Company code SAP'], ['SAP_KODE_UNIT', 'BE00100000', 'Kode SAP unit Concrete Services Unit'],
  ['TTD_MODE', 'KEDUANYA', 'DIGITAL = cap persetujuan elektronik saja · BASAH = kotak tanda tangan kosong · KEDUANYA = keduanya'],
  ['ARSIP_FOLDER', 'Arsip Paket AR CSU', 'Folder Google Drive untuk arsip PDF paket']];

function setupHeader_(sh, title, desc, hdr, cols) {
  sh.getRange('A1').setValue(title).setFontSize(14).setFontWeight('bold').setFontColor(C.navy);
  sh.getRange('A2').setValue(desc).setFontSize(9);
  if (cols) cols.forEach(function (c, i) { sh.getRange(hdr, i + 1).setValue(c[1]).setBackground(C.navy).setFontColor('#FFFFFF').setFontWeight('bold').setWrap(true).setHorizontalAlignment('center').setVerticalAlignment('middle').setFontSize(9); sh.setColumnWidth(i + 1, (colWidth_(c) || 14) * 7); });
}
function setupReport_(ss, name, title, desc, selector, idx) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name, Math.min(idx, ss.getSheets().length));
  sh.getRange('A1').setValue(title).setFontSize(14).setFontWeight('bold').setFontColor(C.navy);
  sh.getRange('A2').setValue(desc).setFontSize(10);
  if (selector) { sh.getRange('B5').setValue(selector).setFontWeight('bold').setHorizontalAlignment('right'); sh.getRange('C5').setValue(name === 'CKPN' ? 'OTOMATIS (AKHIR BULAN LALU)' : 'OTOMATIS (BULAN INI)').setBackground('#FFE699').setFontWeight('bold'); }
  sh.setFrozenRows(5); sh.setHiddenGridlines(true); sh.setTabColor('#1D6B73');
  return sh;
}
function upgradeV2() { arRequireMaker_();
  var ui = SpreadsheetApp.getUi(), ss = arSpreadsheet_(), log = [];
  if (ui.alert('Upgrade ke versi 2', 'Menambah kolom & sheet baru, memindahkan sheet yang digabung ke arsip (disembunyikan), dan menambah tabel di Parameter. Data yang sudah ada tidak dihapus. Lanjutkan?', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  // kolom baru
  [['Monitoring Invoice', ENGINE.REG, ENGINE.REG_ALIAS, 6], ['Log Penerimaan', ENGINE.LOG, null, 4], ['Master Pelanggan', ENGINE.MPEL, null, null], ['Master Proyek', ENGINE.MPRO, ENGINE.MPRO_ALIAS, null]].forEach(function (x) {
    var sh = ss.getSheetByName(x[0]); if (!sh) return; var n = ensureSchema_(sh, x[1], x[2], x[3]); if (n) log.push(x[0] + ': +' + n + ' kolom');
  });
  var mi = ss.getSheetByName('Monitoring Invoice'), vm = colMap_(mi, ENGINE.REG, ENGINE.REG_ALIAS);
  if (vm.validasi) mi.getRange(ENGINE.REG.headerRow, vm.validasi).setValue('USULAN RENCANA CAIR');
  // Parameter (pastikan kolom sampai AD tersedia)
  var par = ss.getSheetByName('Parameter');
  if (par.getMaxColumns() < 31) par.insertColumnsAfter(par.getMaxColumns(), 31 - par.getMaxColumns());
  if (norm_(par.getRange('O3').getValue()) !== 'KODE') {
    par.getRange('O3:S3').setValues([['KODE', 'NAMA AKUN', 'GL PIHAK BERELASI', 'GL PIHAK KETIGA', 'KETERANGAN']]).setBackground(C.navy).setFontColor('#FFFFFF').setFontWeight('bold');
    par.getRange(4, 15, V2_AKUN.length, 5).setValues(V2_AKUN).setNumberFormat('@'); par.getRange('O2').setValue('PEMETAAN AKUN SAP').setFontWeight('bold'); log.push('Parameter: tabel Pemetaan Akun SAP');
  }
  if (norm_(par.getRange('U3').getValue()) !== 'EMAIL') {
    par.getRange('U3:Y3').setValues([['EMAIL', 'NAMA', 'JABATAN', 'NIP', 'PERAN']]).setBackground(C.navy).setFontColor('#FFFFFF').setFontWeight('bold');
    var seed = V2_USERS.map(function (r) { return r.slice(); }); seed[0][0] = email_();
    par.getRange(4, 21, seed.length, 5).setValues(seed); par.getRange('U2').setValue('PENGGUNA & PERAN (PERSETUJUAN)').setFontWeight('bold');
    par.getRange('Y4:Y40').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(ENGINE.OPT.peran, true).setAllowInvalid(true).build()); log.push('Parameter: tabel Pengguna & Peran');
  }
  if (norm_(par.getRange('AA3').getValue()) !== 'KELAS') {
    par.getRange('AA3:AD3').setValues([['KELAS', 'DARI (HARI LEWAT JT)', 'SAMPAI (HARI)', 'TARIF CKPN']]).setBackground(C.navy).setFontColor('#FFFFFF').setFontWeight('bold');
    par.getRange(4, 27, ENGINE.CKPN_DEFAULT.length, 4).setValues(ENGINE.CKPN_DEFAULT.map(function (r) { return [r[0], r[1], r[2] === Infinity ? '' : r[2], r[3]]; }));
    par.getRange(4, 30, ENGINE.CKPN_DEFAULT.length, 1).setNumberFormat('0.00%'); par.getRange('AA2').setValue('TARIF CKPN PIUTANG USAHA (MATRIKS UMUR)').setFontWeight('bold'); log.push('Parameter: tabel Tarif CKPN');
  }
  var kv = par.getRange('A4:A120').getValues().map(function (r) { return norm_(r[0]); }), lastKv = 3; kv.forEach(function (k, i) { if (k) lastKv = 4 + i; });
  V2_KV.forEach(function (r) { if (kv.indexOf(r[0]) < 0) { lastKv++; par.getRange(lastKv, 1, 1, 3).setValues([r]); } });
  // sheet baru
  var it = ss.getSheetByName('Input Tagbrut');
  if (!it) {
    it = ss.insertSheet('Input Tagbrut'); setupHeader_(it, 'INPUT TAGBRUT / HAD — PENGAKUAN PENDAPATAN & CATATAN TINDAK LANJUT', 'Kiri: satu baris = satu pengakuan pendapatan (mingguan dari P&E; saldo awal & rekon akhir bulan dari lampiran LK 1109/1114). Kanan: catatan tindak lanjut per NPP/pelanggan.', 6, ENGINE.TAG.cols);
    ENGINE.NOTE.cols.forEach(function (c, i) { it.getRange(6, 9 + i).setValue(c[1]).setBackground(C.manual).setFontColor('#FFFFFF').setFontWeight('bold').setWrap(true).setHorizontalAlignment('center').setFontSize(9); it.setColumnWidth(9 + i, c[2] * 7); });
    it.getRange(7, 1, 500, 6).setBackground(C.input); it.getRange(7, 9, 500, 4).setBackground(C.input);
    it.getRange(7, 2, 500, 1).setNumberFormat(NUMF.date); it.getRange(7, 4, 500, 1).setNumberFormat(NUMF.money); it.getRange(7, 10, 500, 1).setNumberFormat(NUMF.date);
    it.setFrozenRows(6); it.setTabColor('#9C6500'); log.push('Sheet Input Tagbrut dibuat');
  }
  var dl = ss.getSheetByName('Data LK (Tempel)');
  if (!dl) {
    dl = ss.insertSheet('Data LK (Tempel)');
    if (dl.getMaxColumns() < 46) dl.insertColumnsAfter(dl.getMaxColumns(), 46 - dl.getMaxColumns());
    dl.getRange('A1').setValue('DATA LK (TEMPEL) — UNTUK REKON SAP').setFontSize(14).setFontWeight('bold').setFontColor(C.navy);
    dl.getRange('A2').setValue('Tempel isi lampiran LK unit setiap tutup buku (nilai saja: Tempel khusus ▸ Hanya nilai). Baris 8 ke bawah. Kosongkan dulu isi lama sebelum menempel.');
    dl.getRange('B4').setValue('PERIODE LK ▸').setFontWeight('bold'); dl.getRange('C4').setBackground('#FFE699').setNumberFormat(NUMF.date);
    dl.getRange('A6').setValue('A. PIUTANG USAHA — lampiran 1104111000 & 1104211000, kolom B s.d. O (mulai baris data pertama)').setFontWeight('bold');
    dl.getRange('Q6').setValue('B. PIUTANG PPN K — lampiran 1104112000 & 1104212000, kolom B s.d. P').setFontWeight('bold');
    dl.getRange('AG6').setValue('C. PIUTANG RETENSI — lampiran 1105111000 & 1105211000, kolom B s.d. N').setFontWeight('bold');
    dl.getRange(7, 1, 1, 14).setValues([['No', 'Pelanggan', 'Kode BP', 'NPWP', 'NPP', 'Proyek', 'Tgl Terbit', 'No. Bukti', 'Jumlah', 'Jatuh Tempo', 'Tgl Cair', 'No. Bukti Cair', 'Jumlah Cair', 'Saldo']]).setBackground(C.navy).setFontColor('#FFFFFF').setFontWeight('bold');
    dl.getRange(7, 17, 1, 15).setValues([['No', 'Pelanggan', 'Kode BP', 'NPWP', 'NPP', 'No. Faktur', 'Proyek', 'Tgl Terbit', 'No. Bukti', 'Jumlah', 'Jatuh Tempo', 'Tgl Cair', 'No. Bukti Cair', 'Jumlah Cair', 'Saldo']]).setBackground(C.navy).setFontColor('#FFFFFF').setFontWeight('bold');
    dl.getRange(7, 33, 1, 13).setValues([['No', 'Pelanggan', 'Kode BP', 'NPP', 'Proyek', 'Tgl Terbit', 'No. Bukti', 'Jumlah', 'Tgl Cair', 'No. Bukti Cair', 'Jumlah Cair', 'Saldo', 'Saldo Bersih']]).setBackground(C.navy).setFontColor('#FFFFFF').setFontWeight('bold');
    dl.setFrozenRows(7); dl.setTabColor('#9C6500'); log.push('Sheet Data LK (Tempel) dibuat');
  }
  var ap = ss.getSheetByName('Log Persetujuan');
  if (!ap) {
    ap = ss.insertSheet('Log Persetujuan'); setupHeader_(ap, 'LOG PERSETUJUAN DIGITAL', 'Diisi otomatis oleh aplikasi. JANGAN diubah manual — setiap keputusan memiliki kode verifikasi; perubahan manual membuat kode TIDAK VALID.', 5, ENGINE.APPR.cols);
    ap.setFrozenRows(5); ap.setTabColor('#7F7F7F'); ap.getRange(6, 1, 1000, ENGINE.APPR.cols.length).setNumberFormat('@');
    try { ap.protect().setDescription('Log persetujuan — diisi aplikasi').setWarningOnly(true); } catch (e) { }
    log.push('Sheet Log Persetujuan dibuat');
  }
  setupReport_(ss, 'Monitoring Tagbrut', 'MONITORING TAGBRUT / HAD', 'Pengakuan pendapatan (Input Tagbrut) dikurangi tagihan yang terbit. Pilih bulan di C5 untuk posisi akhir bulan; OTOMATIS = hari ini.', 'POSISI ▸', 99);
  setupReport_(ss, 'CKPN', 'CKPN / IMPAIRMENT PIUTANG', 'Matriks umur (piutang usaha) dan NPV (retensi, tagbrut, HAD), mengikuti lampiran LK unit. Pilih bulan di C5 (posisi akhir bulan).', 'PER AKHIR BULAN ▸', 99);
  setupReport_(ss, 'Usulan Jurnal SAP', 'USULAN JURNAL SAP', 'Jurnal invoice keluar, penerimaan, dan CKPN per bulan untuk diserahkan ke Akuntansi.', 'PILIH BULAN ▸', 99);
  setupReport_(ss, 'Rekon SAP', 'REKON SISTEM AR vs LAMPIRAN LK (SAP)', 'Membandingkan saldo per dokumen dengan lampiran LK yang ditempel di sheet Data LK (Tempel).', null, 99);
  // sheet yang digabung → arsip
  [['Validasi Termin', null], ['Rekap Invoice Cair', null], ['Kertas Kerja AR', 'Arsip Kertas Kerja AR'], ['Catatan Migrasi', 'Arsip Catatan Migrasi']].forEach(function (x) {
    var sh = ss.getSheetByName(x[0]); if (!sh) return;
    if (x[1]) { sh.setName(x[1]); sh.hideSheet(); log.push(x[0] + ' → ' + x[1] + ' (disembunyikan)'); }
    else { ss.deleteSheet(sh); log.push(x[0] + ' dihapus (digabung ke ' + (x[0] === 'Validasi Termin' ? 'Pareto' : 'INKASO') + ')'); }
  });
  ss.getSheets().forEach(function (s) { if (/^KP /.test(s.getName())) ss.deleteSheet(s); });
  var rr = ss.getSheetByName('Riwayat Revisi'); if (rr) rr.hideSheet();
  var order = ['Panduan', 'Dashboard Summary', 'Monitoring Invoice', 'Log Penerimaan', 'Kartu Piutang', 'Pareto', 'RA-RI Termin', 'INKASO', 'Aging Piutang', 'Piutang Retensi',
    'Monitoring Tagbrut', 'Rekap Invoice Keluar', 'Usulan Jurnal SAP', 'CKPN', 'Rekon SAP', 'Cek Data', 'Master Pelanggan', 'Master Proyek', 'Riwayat Adendum', 'Input Tagbrut', 'Data LK (Tempel)', 'Parameter', 'Log Persetujuan', 'Riwayat Revisi'];
  order.forEach(function (n, i) { var s = ss.getSheetByName(n); if (s) { ss.setActiveSheet(s); ss.moveActiveSheet(i + 1); } });
  ['Usulan Jurnal SAP', 'CKPN', 'Rekon SAP', 'Rekap Invoice Keluar'].forEach(function (n) { var s = ss.getSheetByName(n); if (s) s.setTabColor('#7030A0'); });
  secret_();
  rebuildAll();
  ui.alert('Upgrade selesai.\n\n' + (log.length ? '• ' + log.join('\n• ') : 'Struktur sudah versi 2 — tidak ada perubahan.') + '\n\nLangkah berikut: isi Kode BP & NITKU di Master Pelanggan, AKUN PENGAKUAN & KODE JASA CORETAX di Master Proyek, tempel saldo awal tagbrut di Input Tagbrut, lalu deploy ulang aplikasi web (Jalankan sebagai: pengguna yang mengakses).');
}

// =====================================================================  v2.4 — MEMORIAL AKUNTANSI, FILE AKUNTANSI & P&E, TEMPEL SAP
var MEMO_CF_TAG = 'memo';
/** tulis ulang sheet Memorial AR dari hasil mesin (blok baru / revisi / koreksi; isian manual dipertahankan) */
function writeMemo_(sh, rows) {
  if (!sh) return;
  var M = ENGINE.MEMO, nc = M.cols.length, f = M.firstRow, last = Math.max(sh.getLastRow(), f), need = f + rows.length + 5;
  if (need > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows() + 50);
  if (sh.getMaxColumns() < nc) sh.insertColumnsAfter(sh.getMaxColumns(), nc - sh.getMaxColumns());
  var clr = sh.getRange(f, 1, Math.max(last - f + 1, 1), nc);
  clr.clearContent(); clr.setBackground(null).setFontWeight('normal').setBorder(false, false, false, false, false, false);
  if (!rows.length) return;
  sh.getRange(f, 2, rows.length, 7).setNumberFormat('@'); sh.getRange(f, 14, rows.length, nc - 13).setNumberFormat('@');
  sh.getRange(f, 1, rows.length, nc).setValues(rows.map(function (r) { return r.map(function (v, j) { return j >= 1 && j <= 7 || j >= 13 ? (v === null || v === undefined ? '' : String(v)) : v; }); }));
  sh.getRange(f, 9, rows.length, 5).setNumberFormat(NUMF.money);
  // pita pada baris pertama tiap blok
  var firsts = [];
  rows.forEach(function (r, i) { if (i === 0 || rows[i - 1][15] !== r[15]) firsts.push('A' + (f + i) + ':' + 'W' + (f + i)); });
  for (var k = 0; k < firsts.length; k += 200) sh.getRangeList(firsts.slice(k, k + 200)).setBackground(C.band).setFontWeight('bold');
}
/** pengeditan manual di Memorial AR: seluruh blok (kode sama) ditandai MANUAL agar tidak ditimpa sistem */
function memoMarkManual_(sh, row, nr, col, ncols) {
  var M = ENGINE.MEMO, cK = ENGINE.colIndex(M, 'kode') + 1, cS = ENGINE.colIndex(M, 'sumber') + 1;
  if (col > 13) return; // hanya kolom isi jurnal (A–M)
  var last = sh.getLastRow(); if (last < M.firstRow) return;
  var kv = sh.getRange(M.firstRow, cK, last - M.firstRow + 1, 1).getValues().map(function (r) { return String(r[0]).trim(); }), kodes = {};
  for (var r = row; r < row + nr; r++) { var k = kv[r - M.firstRow]; if (k) kodes[k] = 1; }
  var sv = sh.getRange(M.firstRow, cS, kv.length, 1).getValues(), ch = false;
  kv.forEach(function (k, i) { if (kodes[k] && String(sv[i][0]).toUpperCase() !== 'MANUAL') { sv[i][0] = 'MANUAL'; ch = true; } });
  if (ch) sh.getRange(M.firstRow, cS, kv.length, 1).setValues(sv);
}
/** menu: kembalikan blok memorial pada baris terpilih ke hitungan sistem (sumber OTOMATIS) */
function memoRebuildSelected() { arRequireMaker_();
  var ui = SpreadsheetApp.getUi(), sh = SpreadsheetApp.getActiveSheet ? SpreadsheetApp.getActiveSheet() : null, M = ENGINE.MEMO;
  if (!sh || sh.getName() !== M.sheet) { ui.alert('Buka sheet "' + M.sheet + '", klik salah satu baris blok yang ingin dibangun ulang, lalu jalankan menu ini lagi.'); return; }
  var row = sh.getActiveRange().getRow(), cK = ENGINE.colIndex(M, 'kode') + 1, kode = String(sh.getRange(row, cK).getValue()).trim();
  if (!kode || row < M.firstRow) { ui.alert('Baris terpilih bukan baris memorial.'); return; }
  if (ui.alert('Bangun ulang ' + kode, 'Isi blok ' + kode + ' akan diganti dengan hitungan sistem terbaru (perubahan manual hilang). Blok yang sudah diverifikasi Akuntansi tidak diubah — sistem membuat blok KOREKSI. Lanjutkan?', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  var last = sh.getLastRow(), cS = ENGINE.colIndex(M, 'sumber') + 1, cF = ENGINE.colIndex(M, 'sidik') + 1;
  var kv = sh.getRange(M.firstRow, cK, last - M.firstRow + 1, 1).getValues();
  kv.forEach(function (r, i) { if (String(r[0]).trim() === kode) { sh.getRange(M.firstRow + i, cS).setValue('OTOMATIS'); sh.getRange(M.firstRow + i, cF).setValue(''); } });
  rebuildAll();
  ui.alert('Blok ' + kode + ' dibangun ulang.');
}

// ---------------------------------------------------------------- PARAMETER (kunci-nilai) & file turunan
function paramKV_(key) {
  var par = sheet_('Parameter'), v = par.getRange('A4:B120').getValues();
  for (var i = 0; i < v.length; i++) if (norm_(v[i][0]) === key) return { row: 4 + i, val: String(v[i][1] || '').trim() };
  return null;
}
function setParamKV_(key, val, desc) {
  var par = sheet_('Parameter'), x = paramKV_(key);
  if (x) { par.getRange(x.row, 2).setNumberFormat('@').setValue(val); return; }
  var v = par.getRange('A4:A120').getValues(), lastKv = 3; v.forEach(function (r, i) { if (String(r[0]).trim()) lastKv = 4 + i; });
  par.getRange(lastKv + 1, 1, 1, 3).setNumberFormat('@').setValues([[key, val, desc || '']]);
}
/** buka file turunan (Akuntansi / P&E); dibuat baru bila belum ada atau tidak bisa dibuka */
function derivedFile_(key, title) {
  var x = paramKV_(key), ss = null;
  if (x && x.val) {
    // file sudah pernah dibuat: jangan pernah membuat file pengganti (bisa jadi akun ini tidak punya akses)
    try { return SpreadsheetApp.openById(x.val); } catch (e) { throw new Error(title + ' cannot be opened by this account (' + e.message + '). Run it from the file owner account, or clear ' + key + ' in Parameter to create a new file.'); }
  }
  {
    ss = SpreadsheetApp.create(title);
    setParamKV_(key, ss.getId(), 'ID file ' + title + ' (dibuat otomatis — kosongkan untuk membuat file baru)');
    try { var f = DriveApp.getFileById(ss.getId()); f.setShareableByEditors(false); f.setDescription('Dibuat otomatis oleh Sistem AR CSU (dirancang & dikembangkan oleh ' + PENGEMBANG + '). Isi file ini ditimpa sistem — jangan diubah manual.'); } catch (e) { }
  }
  return ss;
}
function shareTo_(ss, emails, asEditor) {
  try {
    var f = DriveApp.getFileById(ss.getId()), have = {};
    f.getEditors().concat(f.getViewers()).forEach(function (u) { have[String(u.getEmail()).toLowerCase()] = 1; });
    emails.forEach(function (e) { if (!e || have[e]) return; if (asEditor) f.addEditor(e); else f.addViewer(e); });
  } catch (e) { }
}
function sheetIn_(ss, name, idx) { var sh = ss.getSheetByName(name); if (!sh) sh = ss.insertSheet(name, idx === undefined ? ss.getSheets().length : idx); return sh; }
function dropDefaultSheet_(ss) { var s1 = ss.getSheetByName('Sheet1') || ss.getSheetByName('Lembar1'); if (s1 && ss.getSheets().length > 1 && s1.getLastRow() === 0) ss.deleteSheet(s1); }
function memoTabName_(bulan, jenis) { var p = String(bulan).split('-'); return (jenis === 'TRM' ? 'TRM ' : 'INV ') + ENGINE.MONTHS[+p[1] - 1].toUpperCase() + ' ' + p[0].substring(2); }

/** tarik centang verifikasi dari file Akuntansi → kolom VERIFIKASI di sheet Memorial AR */
function pullVerif_() {
  var x = paramKV_('FILE_AKUNTANSI_ID'); if (!x || !x.val) return 0;
  var ss; try { ss = SpreadsheetApp.openById(x.val); } catch (e) { return 0; }
  var got = {};
  ss.getSheets().forEach(function (sh) {
    if (!/^(INV|TRM) /.test(sh.getName())) return;
    var last = sh.getLastRow(); if (last < 4) return;
    sh.getRange(4, 15, last - 3, 2).getValues().forEach(function (r) { var k = String(r[1]).trim(); if (k && (r[0] === true || r[0] === false)) got[k] = got[k] || r[0] === true; });
  });
  return applyVerif_(got);
}
function applyVerif_(got) {
  var M = ENGINE.MEMO, sh = sheet_(M.sheet); if (!sh) return 0;
  var last = sh.getLastRow(); if (last < M.firstRow) return 0;
  var n = last - M.firstRow + 1, cK = ENGINE.colIndex(M, 'kode') + 1, cV = ENGINE.colIndex(M, 'verif') + 1, cA = ENGINE.colIndex(M, 'verif_at') + 1;
  var K = sh.getRange(M.firstRow, cK, n, 1).getValues(), V = sh.getRange(M.firstRow, cV, n, 1).getValues(), A = sh.getRange(M.firstRow, cA, n, 1).getValues();
  var now = Utilities.formatDate(new Date(), TZ, 'dd-MM-yyyy HH:mm'), seen = {}, ch = 0;
  K.forEach(function (r, i) {
    var k = String(r[0]).trim(); if (!k || seen[k] || !(k in got)) return; seen[k] = 1;
    var on = !!got[k], was = /^✓/.test(String(V[i][0]));
    if (on && !was) { V[i][0] = '✓'; if (!String(A[i][0]).trim()) A[i][0] = now; ch++; }
    if (!on && was) { V[i][0] = ''; A[i][0] = ''; ch++; }
  });
  if (ch) { sh.getRange(M.firstRow, cV, n, 1).setValues(V); sh.getRange(M.firstRow, cA, n, 1).setNumberFormat('@').setValues(A); }
  return ch;
}
/** kirim memorial ke file Akuntansi: satu tab per bulan & jenis, centang VERIF dipertahankan */
function pushMemorial_(res, months) {
  var memo = res.ctx.memo; if (!memo) return 0;
  var ss = derivedFile_('FILE_AKUNTANSI_ID', 'MEMORIAL AR CSU — AKUNTANSI'), tabs = {};
  memo.blocks.forEach(function (b) { var t = memoTabName_(b.bulan, b.jenis); (tabs[t] = tabs[t] || { bulan: b.bulan, jenis: b.jenis, blocks: [] }).blocks.push(b); });
  var names = Object.keys(tabs).sort(function (a, b) { var x = tabs[a], y = tabs[b]; return x.bulan > y.bulan ? -1 : x.bulan < y.bulan ? 1 : x.jenis < y.jenis ? -1 : 1; });
  var stamp = Utilities.formatDate(new Date(), TZ, 'dd MMM yyyy HH:mm'), extra = {}, n = 0;
  names.forEach(function (t, ix) {
    var T = tabs[t]; if (months && months.indexOf(T.bulan) < 0 && ss.getSheetByName(t)) return;
    var sh = sheetIn_(ss, t, ix);
    // centang terbaru di tab ini (bisa saja baru dicentang sejak penarikan terakhir)
    var last0 = sh.getLastRow(), cur = {};
    if (last0 >= 4) sh.getRange(4, 15, last0 - 3, 2).getValues().forEach(function (r) { var k = String(r[1]).trim(); if (k && r[0] === true) cur[k] = true; });
    var rows = [], firsts = [], lasts = [], no = 0;
    T.blocks.forEach(function (b) {
      no++; var td = 0, tk = 0; b.rows.forEach(function (r) { td += Number(r.debet) || 0; tk += Number(r.kredit) || 0; });
      var v = b.verif || !!cur[b.kode]; if (v && !b.verif) extra[b.kode] = true;
      var base = String(b.status || '').replace(/^✓ DIVERIFIKASI( · )?/, '');
      var st = v ? '✓ DIVERIFIKASI' + (/DIKOREKSI/.test(base) ? ' · ' + base : '') : (base || (/-K\d+$/.test(b.kode) ? 'KOREKSI' : 'BARU'));
      b.rows.forEach(function (r, i) {
        if (i === 0) firsts.push({ r: rows.length, v: v });
        var lastR = i === b.rows.length - 1; if (lastR) lasts.push(rows.length);
        rows.push([i === 0 ? no : '', r.akun, r.uraian, r.npp, r.pelanggan, r.produk, r.pajak, r.inv, r.debet || '', r.kredit || '', lastR ? td : '', lastR ? tk : '', lastR ? td - tk : '', i === 0 ? st : '', '', b.kode]);
      });
      rows.push(['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
    });
    var maxR = Math.max(sh.getMaxRows(), rows.length + 10); if (sh.getMaxRows() < maxR) sh.insertRowsAfter(sh.getMaxRows(), maxR - sh.getMaxRows());
    if (sh.getMaxColumns() < 16) sh.insertColumnsAfter(sh.getMaxColumns(), 16 - sh.getMaxColumns());
    sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function (p) { try { p.remove(); } catch (e) { } });
    var all = sh.getRange(1, 1, sh.getMaxRows(), 16); all.clear(); try { all.clearDataValidations(); } catch (e) { } try { all.breakApart(); } catch (e) { }
    var p = T.bulan.split('-');
    sh.getRange('A1').setValue('DOKUMEN MEMORIAL ' + (T.jenis === 'TRM' ? 'PENERIMAAN' : 'INVOICE KELUAR') + ' — ' + ENGINE.MONTHS[+p[1] - 1].toUpperCase() + ' ' + p[0] + ' · CONCRETE SERVICES UNIT').setFontWeight('bold').setFontSize(12);
    sh.getRange('A2').setValue('Diisi otomatis dari Sistem AR CSU · diperbarui ' + stamp + ' WIB · centang kolom VERIF bila blok sudah diperiksa / diinput ke SAP (hanya kolom VERIF yang bisa diubah).').setFontSize(9).setFontColor('#52514E');
    var hdr = ['NO', 'AKUN', 'URAIAN TAMBAHAN', 'NO. NPP', 'PELANGGAN', 'PRODUK', 'PAJAK', 'NO INVOICE', 'DEBET', 'KREDIT', 'JUMLAH (DEBET)', 'JUMLAH (KREDIT)', 'SELISIH', 'STATUS', 'VERIF', 'KODE MEMORIAL'];
    sh.getRange(3, 1, 1, 16).setValues([hdr]).setBackground('#00B0F0').setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
    if (rows.length) {
      sh.getRange(4, 1, rows.length, 16).setValues(rows).setVerticalAlignment('middle').setHorizontalAlignment('center');
      sh.getRange(4, 3, rows.length, 1).setHorizontalAlignment('left').setWrap(true);
      sh.getRange(4, 5, rows.length, 2).setHorizontalAlignment('left').setWrap(true);
      sh.getRange(4, 8, rows.length, 1).setHorizontalAlignment('left').setWrap(true);
      sh.getRange(4, 9, rows.length, 5).setHorizontalAlignment('right');
      sh.getRange(4, 9, rows.length, 5).setNumberFormat('#,##0');
      var cbA = firsts.map(function (x) { return 'O' + (4 + x.r); }), onA = firsts.filter(function (x) { return x.v; }).map(function (x) { return 'O' + (4 + x.r); });
      for (var k0 = 0; k0 < cbA.length; k0 += 200) sh.getRangeList(cbA.slice(k0, k0 + 200)).insertCheckboxes();
      for (var k1 = 0; k1 < onA.length; k1 += 200) sh.getRangeList(onA.slice(k1, k1 + 200)).check();
      var fr = firsts.map(function (x) { return 'A' + (4 + x.r) + ':P' + (4 + x.r); });
      for (var k = 0; k < fr.length; k += 200) sh.getRangeList(fr.slice(k, k + 200)).setFontWeight('bold');
    }
    sh.getRange(1, 1, Math.max(rows.length + 3, 3), 16).setFontFamily('Play').setFontSize(10);
    sh.getRange('A1').setFontSize(12);
    [5, 9, 76, 12, 29, 55, 18, 24, 14, 15, 16, 16, 10, 20, 7, 15].forEach(function (w, i) { sh.setColumnWidth(i + 1, Math.round(w * 7.3 + 6)); });
    sh.setFrozenRows(3);
    // kunci: hanya kolom VERIF yang boleh diubah
    try {
      var pr = sh.protect().setDescription('AR CSU — hanya kolom VERIF yang diisi Akuntansi');
      if (rows.length) pr.setUnprotectedRanges([sh.getRange(4, 15, rows.length, 1)]);
      pr.removeEditors(pr.getEditors()); try { if (pr.canDomainEdit()) pr.setDomainEdit(false); } catch (e) { }
    } catch (e) { }
    n++;
  });
  ss.getSheets().forEach(function (sh) { if (/^(INV|TRM) /.test(sh.getName()) && !tabs[sh.getName()]) { try { ss.deleteSheet(sh); } catch (e) { } } });
  arOrderNamed_(ss, names);
  dropDefaultSheet_(ss);
  if (Object.keys(extra).length) applyVerif_(extra);
  var P = res.params, mail = P.users.filter(function (u) { return /^PENERIMA/.test(u.peran) && u.email; }).map(function (u) { return u.email; });
  shareTo_(ss, mail, true);
  return n;
}
/** kirim ringkasan & kartu piutang (bagian A + B) ke file P&E — hanya nilai, tanpa rumus/script */
function pushPE_(res, npps) { return arPushPE_(res, npps); }
/** sinkron file turunan: tarik centang Akuntansi → hitung ulang → kirim memorial & kartu P&E */
function syncTurunan_(opt) {
  opt = opt || {};
  var out = { verif: 0, memo: 0, pe: 0, pajak: 0, err: [] };
  // v2.8.7: kunci utama hanya dipegang saat menarik centang Akuntansi & membangun ulang laporan; pengiriman ke file Akuntansi/P&E/Pajak
  // (file lain, bisa beberapa menit) memakai penanda sinkron sendiri — simpan/ajukan di aplikasi tidak lagi antre di belakangnya
  if (!arSyncClaim_()) { out.err.push('another sync is running — it will retry on the next schedule'); return out; }
  try { return syncRun_(opt, out); } finally { arSyncRelease_(); }
}
function arSyncClaim_() {
  var l = arLock_(); if (!l.tryLock(10000)) return false;
  try { var p = PropertiesService.getScriptProperties(), t = +(p.getProperty('SYNC_BUSY') || 0); if (t && Date.now() - t < 7 * 60000) return false; p.setProperty('SYNC_BUSY', String(Date.now())); return true; } finally { l.releaseLock(); }
}
function arSyncRelease_() { try { PropertiesService.getScriptProperties().deleteProperty('SYNC_BUSY'); } catch (e) { } }
function syncRun_(opt, out) {
  var res;
  if (opt.noPull) res = LAST_RES_ || compute_(); // setelah simpan: hasil pembaruan laporan barusan (eksekusi yang sama) dipakai ulang
  else {
    var sl = arLock_();
    if (!sl.tryLock(30000)) { out.err.push('system busy — it will retry on the next schedule'); return out; }
    try {
      try { out.verif = pullVerif_(); } catch (e) { out.err.push('verification pull: ' + e.message); }
      rebuildAll_(); res = LAST_RES_ || compute_();
    } finally { sl.releaseLock(); }
  }
  try { out.memo = pushMemorial_(res, opt.months); } catch (e) { out.err.push('Accounting file: ' + e.message); }
  try { out.pe = pushPE_(res, opt.npps); } catch (e) { out.err.push('P&E file: ' + e.message); }
  try { out.pajak = arPushPajak_(res); } catch (e) { out.err.push('Tax file: ' + e.message); }
  arProps_().setProperty('SYNC_LAST', Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'));
  return out;
}
function syncTurunanAuto(e) { if (!arFromTrigger_(e)) arRequireMaker_(); var a = paramKV_('FILE_AKUNTANSI_ID'), b = paramKV_('FILE_PE_ID'), c = paramKV_('FILE_PAJAK_ID'); if (!(a && a.val) && !(b && b.val) && !(c && c.val)) return; syncTurunan_(); }
function syncTurunan() { arRequireMaker_();
  var r = syncTurunan_(), A = paramKV_('FILE_AKUNTANSI_ID'), P = paramKV_('FILE_PE_ID'), J = paramKV_('FILE_PAJAK_ID');
  var url = function (x) { return x && x.val ? 'https://docs.google.com/spreadsheets/d/' + x.val : '-'; };
  SpreadsheetApp.getUi().alert('File turunan diperbarui', 'Centang verifikasi baru dari Akuntansi: ' + r.verif + '\nTab memorial dikirim: ' + r.memo + '\nKartu proyek dikirim ke P&E: ' + r.pe +
    '\nTab Pajak diperbarui: ' + (r.pajak || 0) + '\n\nFile Akuntansi: ' + url(A) + '\nFile P&E: ' + url(P) + '\nFile Pajak: ' + url(J) + (r.err.length ? '\n\n⚠ ' + r.err.join('\n⚠ ') : '') +
    '\n\nAkses: email PENERIMA (Akuntansi) di Parameter ▸ Pengguna & Peran diberi akses edit kolom VERIF; email peran P&E (LIHAT) diberi akses lihat.', SpreadsheetApp.getUi().ButtonSet.OK);
}
/** setelah simpan invoice / penerimaan: kirim hanya bulan & proyek yang terkait (gagal pun tidak membatalkan simpan) */
function syncAfterSave_(npps, months) {
  try { var a = paramKV_('FILE_AKUNTANSI_ID'), b = paramKV_('FILE_PE_ID'), c = paramKV_('FILE_PAJAK_ID'); if (!(a && a.val) && !(b && b.val) && !(c && c.val)) return; syncTurunan_({ npps: npps, months: months, noPull: true }); } catch (e) { }
}

// ---------------------------------------------------------------- TEMPEL SAP → No. Dok SAP
function fillSapFromTempel() { arRequireMaker_();
  var ui = SpreadsheetApp.getUi(), sh = sheet_('Tempel SAP');
  if (!sh || sh.getLastRow() < 6) { ui.alert('Tempel export SAP (FBL5N / daftar dokumen, beserta baris judul kolomnya) di sheet "Tempel SAP" mulai baris 6, lalu jalankan menu ini lagi.'); return; }
  var vals = nd_(sh.getRange(6, 1, sh.getLastRow() - 5, Math.max(sh.getLastColumn(), 1)).getValues()), sap = ENGINE.sapParse(vals);
  if (!sap.ok) { ui.alert('Judul kolom tidak dikenali. Pastikan baris judul ikut ditempel (mis. Document Number, Reference/Assignment, Amount in LC, Posting Date, Document Type).'); return; }
  var res = compute_(), m = ENGINE.sapMatch(res.ctx, sap), mi = sheet_('Monitoring Invoice'), mm = colMap_(mi, ENGINE.REG, ENGINE.REG_ALIAS), lp = sheet_('Log Penerimaan'), lm = colMap_(lp, ENGINE.LOG);
  var nI = 0, nP = 0, beda = [];
  m.inv.forEach(function (x) { if (x.sama) return; if (x.lama) { beda.push(x.nomor + ': sistem ' + x.lama + ' / SAP ' + x.doc); return; } var r = findRegisterRow_(mi, x.nomor, mm); if (r && mm.sap_doc) { mi.getRange(r, mm.sap_doc).setNumberFormat('@').setValue(x.doc); nI++; } });
  m.pay.forEach(function (x) { if (x.sama) return; if (x.lama) { beda.push(x.idp + ' (' + x.nomor + '): sistem ' + x.lama + ' / SAP ' + x.doc); return; } if (lm.sap_doc) { lp.getRange(x.row, lm.sap_doc).setNumberFormat('@').setValue(x.doc); nP++; } });
  rebuildAll();
  ui.alert('No. Dok SAP diisi dari Tempel SAP', nI + ' invoice dan ' + nP + ' penerimaan diisi otomatis (' + sap.rows.length + ' dokumen SAP terbaca, ' + m.sisa.length + ' tidak berpasangan).' +
    (beda.length ? '\n\nTidak diubah (nomor di sistem berbeda — mohon dicek):\n' + beda.slice(0, 15).join('\n') : ''), ui.ButtonSet.OK);
}

// ---------------------------------------------------------------- PENGAMAN NPP SAAT SIMPAN INVOICE
function checkNppSave_(npp, pelanggan) {
  var rows = readSchema_(sheet_('Master Proyek'), ENGINE.MPRO, ENGINE.MPRO_ALIAS), iN = ENGINE.colIndex(ENGINE.MPRO, 'npp'), iP = ENGINE.colIndex(ENGINE.MPRO, 'pelanggan');
  var hit = rows.filter(function (r) { return String(r[iN]).trim() === String(npp).trim(); });
  if (!hit.length) throw new Error('NPP not found in Master Proyek: ' + npp);
  if (hit.length > 1) throw new Error('NPP ' + npp + ' is used by ' + hit.length + ' projects in Master Proyek (' + hit.map(function (r) { return r[iP]; }).join(' / ') + '). Give each project a unique NPP (for a temporary NPP use e.g. ' + String(npp).replace(/ \d{4} /, ' 9001 ') + '), then save again.');
  if (hit.length === 1 && String(hit[0][iP]).trim().toUpperCase() !== String(pelanggan).trim().toUpperCase()) throw new Error('Invoice customer (' + pelanggan + ') differs from the customer of project ' + npp + ' in Master Proyek (' + hit[0][iP] + '). Check the selected project / NPP.');
}

// =====================================================================  PEMBARUAN v2.4
/** daftar NPP dari RESUME OP P&E 2026 R01: [NPP, JENIS, NPP INDUK, NAMA DI P&E (salah ketik dirapikan), NAMA RINGKAS PRODUK, NILAI] */
var V24_PE = [
  ['244C 0660 JT', 'JASA', '', 'PEKERJAAN PONDASI WAREHOUSE 3 (WH3) DAN COLD STORAGE 4 (CS4) NAROGONG', '', 7618185000],
  ['244C 0781 BL', 'PRODUK', '244C 0660 JT', 'TOTAL NAROGONG (PRODUK)', 'Spun pile', 3181815000],
  ['244C 0500 JT', 'JASA', '', 'PEKERJAAN SHEETPILE AREA C PROYEK PARKLAND PODOMORO', '', 3028722196],
  ['244C 0508 EF', 'PRODUK', '244C 0500 JT', 'TOTAL ASTAKONA CCSP (PRODUK)', 'CCSP', 7925957097],
  ['244C 0818 JT', 'JASA', '', 'PEKERJAAN PEMBANGUNAN JEMBATAN ECO PARK PIK', '', 17227035000],
  ['244C 0817 BF', 'PRODUK', '244C 0818 JT', 'TOTAL ECO PARK PIK 3 (PRODUK SPUNPILE)', 'Spun pile', 6482170000],
  ['244C 0816 DF', 'PRODUK', '244C 0818 JT', 'TOTAL ECO PARK PIK 3 (PRODUK HALF SLAB)', 'Half slab', 7090795000],
  ['244C 1876 JT', 'JASA', '', 'PEKERJAAN PEMBANGUNAN JEMBATAN PENGHUBUNG DESA WANAJAYA DENGAN DESA SUKAJAYA CIBITUNG', '', 45337225183],
  ['244C 1875 EF', 'PRODUK', '244C 1876 JT', 'TOTAL JEMBATAN METLAND (PRODUK CCSP)', 'CCSP', 6767442000],
  ['244C 1874 DF', 'PRODUK', '244C 1876 JT', 'TOTAL JEMBATAN METLAND (PRODUK BULB TEE)', 'Bulb tee', 5000000000],
  ['244C 1872 JT', 'JASA', '', 'PEKERJAAN PERKUATAN STRUKTUR LONG POND SEKTOR 3 CIKARANG', '', 5714563200],
  ['244C 1873 EF', 'PRODUK', '244C 1872 JT', 'TOTAL LONG POND (PRODUK)', 'CCSP', 6675436800],
  ['244C 2045 JT', 'JASA', '', 'PEKERJAAN ERECTION BOX GIRDER PEMBANGUNAN JALAN TOL (ELEVATED) ANCOL TIMUR - PLUIT', '', 36558569500],
  ['254C 0166 JT', 'JASA', '', 'PEKERJAAN INFRASTRUKTUR MAINROAD KAWASAN PROYEK PARKLAND PODOMORO', '', 24567568153],
  ['254C 0569 JT', 'JASA', '', 'PEKERJAAN PEMANCANGAN PABRIK COKLAT KARAWANG BARRY CALLEBAUT INDONESIA', '', 14503981853],
  ['254C 0611 JT', 'JASA', '', 'PEKERJAAN JEMBATAN IRIGASI 2 METLAND CIKARANG', '', 10550000000],
  ['254C 0823 JT', 'JASA', '', 'PROYEK PEKERJAAN SHORING PROTECTION TANAH GROGOT', '', 39933668000],
  ['264C 0286 JT', 'ADENDUM', '254C 0823 JT', 'PROYEK PEMANCANGAN TANKI AREA TANAH GROGOT', 'Pemancangan tanki area Tanah Grogot', 5707147350],
  ['254C 1142 JT', 'JASA', '', 'PEKERJAAN TAMBAH METODE KERJA JEMBATAN PENGHUBUNG DESA WANAJAYA', '', 6849683246],
  ['254C 1141 JT', 'JASA', '', 'PEKERJAAN TAMBAH PERBAIKAN JALAN CBL', '', 857499080],
  ['254C 1140 JT', 'JASA', '', 'PEKERJAAN PENINGGIAN DAN PELEBARAN JALAN INSPEKSI KALI CBL', '', 8957567142],
  ['244C 0407 JF', 'JASA', '', 'PEMBANGUNAN JAKARTA SEWERAGE SYSTEM', '', 14000000000],
  ['254C 0824 JT', 'JASA', '', 'PEMBANGUNAN JEMBATAN AKSES UTAMA IKK', '', 158492250000],
  ['254C 1092 JT', 'JASA', '', 'PEMANCANGAN CCSP BSD', '', 1373630000],
  ['254C 1335 JT', 'JASA', '', 'PAKET PEKERJAAN AREA JETTY TANAH GROGOT', '', 125000000],
  ['254C 1584 JL', 'JASA', '', 'PENGADAAN PERALATAN LAUNCHER GANTRY UNTUK PEMBANGUNAN EMPLASEMEN DAN BANGUNAN STASIUN MEDAN TAHAP II (PAKET B)', '', 2430000000],
  ['254C 1643 JT', 'JASA', '', 'PROYEK PEMANCANGAN CCSP DI BSD', '', 1892999995],
  ['254C 1644 JL', 'JASA', '', 'SEWA BEKISTING PARAPET HARBOUR', '', 1968028650],
  ['254C 1645 JT', 'JASA', '', 'PAKET PEKERJAAN ERECTION BOX GIRDER SISI UTARA', '', 24579605500],
  ['254C 1777 JT', 'JASA', '', 'PEMBANGUNAN JEMBATAN PENGHUBUNG DESA WANAJAYA METLAND CIBITUNG', '', 8352086335],
  ['264C 0089 JT', 'JASA', '', 'ANALISIS STABILITAS LERENG DAN PENANGANAN KELONGSORAN SSP A2 METLAND', '', 14095397531],
  ['264C 0088 JT', 'JASA', '', 'SALURAN PEMBUANGAN SERPONG GARDEN 3', '', 2067975000],
  ['264C 0087 JT', 'JASA', '', 'BANGUNAN OUTLET SERPONG GARDEN 3', '', 510000000],
  ['264C 0397 JT', 'JASA', '', 'PEKERJAAN PASANGAN BATU KALI, TUTUP SALURAN UDITCH JEMBATAN IRIGASI TAHAP II', '', 883612415],
  ['264C 0285 JL', 'JASA', '', 'PENYEDIAAN ALAT LG JLKA PAKET 6 MEDAN', '', 7290000000],
  ['264C 0404 JT', 'JASA', '', 'PROYEK PEMBANGUNAN UNDERPASS DI JALAN LINTAS SUMATERA DESA MUARA MAUNG', '', 64873403490],
  ['264C 0718 JT', 'JASA', '', 'PEKERJAAN PERBAIKAN PROTEKSI ABUTMEN A2 PROYEK METLAND', '', 7280343044],
  ['264C 1236 JT', 'JASA', '', 'PEKERJAAN PEMBUATAN PLATFORM SECANT PILE SISI UTARA DAN SISI SELATAN METLAND CIKARANG', '', 5334705100],
  ['264C 1235 JT', 'JASA', '', 'PERBAIKAN STRUKTUR JEMBATAN IRIGASI 1 PROYEK METLAND', '', 1494263863],
  ['264C 1284 JL', 'JASA', '', 'SEWA STEEL SHEET PILE TIPE III', '', 143640000]];
var V24_SELESAI = ['254C 0569 JT', '244C 0660 JT', '244C 1975 JT'];

function setupDataSheet_(ss, schema, title, desc, after, tab) {
  var sh = ss.getSheetByName(schema.sheet), dibuat = false;
  if (!sh) { var af = ss.getSheetByName(after), ix = af ? af.getIndex() : ss.getSheets().length; sh = ss.insertSheet(schema.sheet, ix); dibuat = true; }
  if (sh.getMaxColumns() < schema.cols.length) sh.insertColumnsAfter(sh.getMaxColumns(), schema.cols.length - sh.getMaxColumns());
  setupHeader_(sh, title, desc, schema.headerRow, schema.cols);
  sh.setFrozenRows(schema.headerRow); sh.setTabColor(tab || '#7F7F7F');
  schema.cols.forEach(function (c, i) { var f = colFmt_(c); if (f && NUMF[f]) sh.getRange(schema.firstRow, i + 1, Math.max(sh.getMaxRows() - schema.firstRow, 1), 1).setNumberFormat(NUMF[f]); });
  return { sh: sh, dibuat: dibuat };
}
/** daftar perbaikan data v2.4 (pratinjau dulu sebelum diterapkan) */
function v24Fixes_(ss) {
  var res = compute_(), ctx = res.ctx, F = [], mp = ss.getSheetByName('Master Proyek'), pm = colMap_(mp, ENGINE.MPRO, ENGINE.MPRO_ALIAS), pro = ctx.M.pro;
  var set = function (npp, key, val) { var p = pro[npp]; if (p && pm[key]) mp.getRange(p._row, pm[key]).setValue(val); };
  var prod = {}, adds = {};
  V24_PE.forEach(function (r) { if (r[1] === 'PRODUK') (prod[r[2]] = prod[r[2]] || []).push(r[0] + ' (' + r[4] + ')'); if (r[1] === 'ADENDUM') (adds[r[2]] = adds[r[2]] || []).push(r[0] + ' (' + r[4] + ')'); });
  Object.keys(prod).forEach(function (n) { if (pro[n] && !pro[n].produk.length) F.push({ d: n + ': NPP produk = ' + prod[n].join('; '), run: function () { set(n, 'npp_produk', prod[n].join('; ')); } }); });
  Object.keys(adds).forEach(function (n) { if (pro[n] && !pro[n].adendumNpp.length) F.push({ d: n + ': NPP adendum terpisah = ' + adds[n].join('; ') + ' (tagihan tetap lewat ' + n + ')', run: function () { set(n, 'npp_adendum', adds[n].join('; ')); } }); });
  var nm = V24_PE.filter(function (r) { return r[1] === 'JASA' && pro[r[0]] && !pro[r[0]].nama_pe; });
  if (nm.length) F.push({ d: 'Nama proyek (P&E) diisi untuk ' + nm.length + ' proyek sesuai RESUME OP P&E (salah ketik dirapikan; 254C 0823 JT = SHORING PROTECTION). Nama singkat & deskripsi invoice tidak diubah.', run: function () { nm.forEach(function (r) { set(r[0], 'nama_pe', r[3]); }); } });
  V24_PE.filter(function (r) { return r[1] !== 'JASA' && pro[r[0]] && pro[r[0]].aktif !== 'TIDAK'; }).forEach(function (r) {
    F.push({ d: 'Master Proyek ' + r[0] + ': NPP ' + r[1].toLowerCase() + ' milik ' + r[2] + ' → AKTIF = TIDAK (bukan proyek AR; tidak muncul di daftar & kartu, data tidak dihapus)', run: function () { set(r[0], 'aktif', 'TIDAK'); if (pm.catatan) set(r[0], 'catatan', 'NPP ' + r[1].toLowerCase() + ' milik ' + r[2] + ' (v2.4) — tagihan lewat ' + r[2]); } });
  });
  V24_SELESAI.forEach(function (n) {
    var p = pro[n]; if (!p || p.aktif === 'SELESAI') return;
    if (ctx.invs.some(function (i) { return i.npp === n && !i._batal && !i._res.closed; })) return;
    F.push({ d: 'Master Proyek ' + n + ' (' + p.pelanggan + '): AKTIF = SELESAI (semua invoice lunas) — disembunyikan dari daftar, tidak dihapus', run: function () { set(n, 'aktif', 'SELESAI'); } });
  });
  // tandai sel invoice lama yang perlu dilengkapi manual
  var mi = ss.getSheetByName('Monitoring Invoice'), mm = colMap_(mi, ENGINE.REG, ENGINE.REG_ALIAS), lama = ctx.invs.filter(function (i) { var a = i._a, pr = i._pro || {}; return !i._batal && (a.jenis === 'PROGRES' || a.jenis === 'LAINNYA') && !(Number(i.bruto) > 0) && (pr.um > 0 || pr.um_awal > 0 || pr.ret_pct > 0); });
  if (lama.length) F.push({ d: lama.length + ' invoice lama ditandai kuning di kolom NILAI PROGRES / PENGEMBALIAN UM / RETENSI / POSISI RETENSI untuk diisi manual', run: function () { lama.forEach(function (i) { ['bruto', 'umk', 'retensi', 'retpos'].forEach(function (k) { if (mm[k]) mi.getRange(i._row, mm[k]).setBackground(C.input); }); }); } });
  return F;
}
function applyV24() { arRequireMaker_();
  var ui = SpreadsheetApp.getUi(), ss = arSpreadsheet_(), log = [];
  if (ui.alert('Pembaruan v2.4', 'Langkah 1: struktur — kolom NAMA PROYEK (P&E), NPP PRODUK, NPP ADENDUM di Master Proyek; kolom MEMORIAL AKUNTANSI; kode akun memorial di Parameter; sheet Memorial AR, Referensi P&E, Tempel SAP; dashboard bergrafik.\nLangkah 2: daftar perbaikan data ditampilkan dulu sebelum diterapkan.\nLangkah 3: (pilihan) membuat file Akuntansi & P&E.\n\nData yang sudah ada tidak dihapus. Lanjutkan?', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  [['Monitoring Invoice', ENGINE.REG, ENGINE.REG_ALIAS, 6], ['Log Penerimaan', ENGINE.LOG, null, 4], ['Master Proyek', ENGINE.MPRO, ENGINE.MPRO_ALIAS, null]].forEach(function (x) {
    var sh = ss.getSheetByName(x[0]); if (!sh) return; var n = ensureSchema_(sh, x[1], x[2], x[3]); if (n) log.push(x[0] + ': +' + n + ' kolom');
  });
  // Parameter: kode memorial (kolom T), kunci baru, peran P&E
  var par = ss.getSheetByName('Parameter');
  if (norm_(par.getRange('T3').getValue()) !== 'KODE MEMORIAL') {
    par.getRange('T3').setValue('KODE MEMORIAL').setBackground(C.navy).setFontColor('#FFFFFF').setFontWeight('bold');
    var ak = par.getRange('O4:O80').getValues(), tv = par.getRange('T4:T80').getValues(), nk = 0;
    ak.forEach(function (r, i) { var k = norm_(r[0]); if (k && ENGINE.MEMO_KODE[k] && !String(tv[i][0]).trim()) { tv[i][0] = ENGINE.MEMO_KODE[k]; nk++; } });
    par.getRange('T4:T80').setNumberFormat('@').setValues(tv); log.push('Parameter: kode akun memorial (' + nk + ' akun; bank, PPh & adm menunggu konfirmasi Akuntansi)');
  }
  var mulai = Utilities.formatDate(new Date(), TZ, 'yyyy-MM') + '-01';
  if (!paramKV_('MEMORIAL_MULAI')) { setParamKV_('MEMORIAL_MULAI', mulai, 'Memorial otomatis untuk invoice & penerimaan sejak tanggal ini (sebelumnya sudah dibuat manual)'); log.push('Parameter: MEMORIAL_MULAI = ' + mulai); }
  if (!paramKV_('PE_ABAIKAN')) setParamKV_('PE_ABAIKAN', '254C 1584 JL', 'NPP di daftar P&E yang tidak perlu dicocokkan (pisahkan dengan titik koma)');
  if (!paramKV_('FILE_AKUNTANSI_ID')) setParamKV_('FILE_AKUNTANSI_ID', '', 'ID file MEMORIAL AR CSU — AKUNTANSI (diisi otomatis)');
  if (!paramKV_('FILE_PE_ID')) setParamKV_('FILE_PE_ID', '', 'ID file KARTU PIUTANG CSU — P&E (diisi otomatis)');
  par.getRange('Y4:Y40').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(ENGINE.OPT.peran, true).setAllowInvalid(true).build());
  // sheet baru
  var m = setupDataSheet_(ss, ENGINE.MEMO, 'MEMORIAL AKUNTANSI — INVOICE KELUAR & PENERIMAAN', 'Dibuat otomatis saat invoice / penerimaan disimpan (format DOKUMEN MEMORIAL Akuntansi). Boleh diedit manual: blok yang diedit ditandai MANUAL dan tidak ditimpa. Blok yang sudah diverifikasi Akuntansi tidak pernah diubah — perubahan invoice dibuat sebagai blok KOREKSI.', 'Log Penerimaan', '#7030A0');
  if (m.dibuat) {
    log.push('Sheet Memorial AR dibuat');
    try {
      var st = m.sh.getRange(ENGINE.MEMO.firstRow, 14, 2000, 1), rules = m.sh.getConditionalFormatRules();
      [['✓', '#0A5A0A'], ['KOREKSI', '#7A5200'], ['BATAL', '#B42318'], ['⚠', '#B42318'], ['DIREVISI', '#7A5200']].forEach(function (x) { rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains(x[0]).setFontColor(x[1]).setRanges([st]).build()); });
      m.sh.setConditionalFormatRules(rules);
    } catch (e) { }
  }
  var r = setupDataSheet_(ss, ENGINE.PEREF, 'REFERENSI NPP P&E (RESUME OP)', 'Daftar NPP dari P&E untuk pemeriksaan "Selaras P&E" di Cek Data. Perbarui bila P&E menerbitkan NPP baru (salin dari RESUME OP). JENIS: JASA = proyek AR · PRODUK / ADENDUM = NPP pendamping milik NPP INDUK.', 'Master Proyek', '#9C6500');
  if (r.dibuat || r.sh.getLastRow() < ENGINE.PEREF.firstRow) { r.sh.getRange(ENGINE.PEREF.firstRow, 1, V24_PE.length, 6).setValues(V24_PE); log.push('Sheet Referensi P&E dibuat (' + V24_PE.length + ' NPP dari RESUME OP P&E 2026 R01)'); }
  var ts = ss.getSheetByName('Tempel SAP');
  if (!ts) {
    var dl = ss.getSheetByName('Data LK (Tempel)'); ts = ss.insertSheet('Tempel SAP', dl ? dl.getIndex() : ss.getSheets().length);
    ts.getRange('A1').setValue('TEMPEL SAP — UNTUK MENGISI NO. DOK SAP OTOMATIS').setFontSize(14).setFontWeight('bold').setFontColor(C.navy);
    ts.getRange('A2').setValue('Sebulan sekali: export FBL5N / daftar dokumen dari SAP, tempel (Hanya nilai) mulai baris 6 BERSAMA baris judul kolomnya, lalu menu AR CSU ▸ Isi No. Dok SAP dari Tempel SAP.').setFontSize(9);
    ts.getRange('A3').setValue('Dicocokkan: invoice = No. Invoice di Reference / Assignment / Text + nilai; penerimaan = nominal + tanggal (±10 hari). Nomor yang sudah terisi tidak ditimpa — perbedaannya dilaporkan.').setFontSize(9);
    ts.getRange('A5').setValue('▼ tempel di bawah ini').setFontWeight('bold').setFontColor('#9C6500'); ts.setTabColor('#9C6500'); log.push('Sheet Tempel SAP dibuat');
  }
  var kp = ss.getSheetByName('Kartu Piutang'); if (kp) kp.getRange('A2').setValue('Bagian A: rincian tagihan per invoice (nilai progres, pot. UM, retensi, DPP, PPN). Bagian B: mutasi debit/kredit, saldo di baris "Sisa tagihan" tiap invoice. Pilih NPP di C5, posisi per di E5.');
  // pemicu per jam (file Akuntansi & P&E)
  try { var tr = ScriptApp.getProjectTriggers(); if (tr.some(function (t) { return t.getHandlerFunction() === 'onEditInstalled'; }) && !tr.some(function (t) { return t.getHandlerFunction() === 'syncTurunanAuto'; })) { ScriptApp.newTrigger('syncTurunanAuto').timeBased().everyHours(1).create(); log.push('Pemicu per jam: sinkron file Akuntansi & P&E'); } } catch (e) { }
  rebuildAll();
  // perbaikan data (pratinjau dulu)
  var F = v24Fixes_(ss), done = [];
  if (F.length && ui.alert('Pembaruan v2.4 — perbaikan data', 'Perubahan berikut akan diterapkan:\n\n' + F.map(function (f, i) { return (i + 1) + '. ' + f.d; }).join('\n\n') + '\n\nTerapkan sekarang? (Batal = struktur tetap diperbarui, data tidak diubah)', ui.ButtonSet.OK_CANCEL) === ui.Button.OK) {
    F.forEach(function (f) { f.run(); done.push(f.d); }); rebuildAll();
  }
  var mk = '';
  if (ui.alert('Pembaruan v2.4 — file turunan', 'Buat / perbarui sekarang:\n• MEMORIAL AR CSU — AKUNTANSI (untuk Mbak Elmarta; hanya kolom VERIF yang bisa dicentang)\n• KARTU PIUTANG CSU — P&E (ringkasan + kartu per proyek; hanya lihat, boleh unduh)\n\nAkses diberikan ke email peran PENERIMA dan P&E (LIHAT) di Parameter ▸ Pengguna & Peran. SISTEM AR tetap hanya untuk Anda dan Pak Anas.', ui.ButtonSet.OK_CANCEL) === ui.Button.OK) {
    var s1 = syncTurunan_(), A = paramKV_('FILE_AKUNTANSI_ID'), B = paramKV_('FILE_PE_ID');
    mk = '\n\nFile Akuntansi: ' + (A && A.val ? 'https://docs.google.com/spreadsheets/d/' + A.val : '-') + '\nFile P&E: ' + (B && B.val ? 'https://docs.google.com/spreadsheets/d/' + B.val : '-') + (s1.err.length ? '\n⚠ ' + s1.err.join('\n⚠ ') : '');
  }
  ui.alert('Pembaruan v2.4 selesai', (log.length ? '• ' + log.join('\n• ') : 'Struktur sudah v2.4.') + '\n\n' + (done.length ? 'Perbaikan data diterapkan:\n• ' + done.join('\n• ') : 'Tidak ada perbaikan data yang diterapkan.') + mk +
    '\n\nLangkah berikut:\n1. Terbitkan ulang aplikasi web (Kelola deployment ▸ Versi baru).\n2. Isi email Mbak Elmarta (PENERIMA) dan tim P&E (peran P&E (LIHAT)) di Parameter ▸ Pengguna & Peran, lalu menu AR CSU ▸ Perbarui file Akuntansi & P&E.\n3. Lengkapi kode memorial bank, PPh & adm bank di Parameter kolom T (konfirmasi Akuntansi).\n4. Cek sheet Cek Data (DATA LAMA, SELARAS P&E).', ui.ButtonSet.OK);
}

// =====================================================================  DELEGASI PERSETUJUAN SEMENTARA
/** menu: akun pengganti boleh menyetujui a.n. Section Head sampai tanggal tertentu (tercatat sebagai DELEGASI di cap & Log Persetujuan) */
function aturDelegasi() { arRequireMaker_();
  var ui = SpreadsheetApp.getUi(), cur = paramKV_('DELEGASI_EMAIL'), sd = paramKV_('DELEGASI_SAMPAI'), P = arParameters_(), me = email_();
  var aktif = cur && cur.val ? '\n\nSaat ini: ' + cur.val + ' s.d. ' + (sd ? sd.val : '-') + ' (' + ((paramKV_('DELEGASI_JENIS') || {}).val || 'PAY,VAL') + ')' : '';
  var q = ui.prompt('Delegasi persetujuan Oktober 2026', 'Email akun pengganti yang menyetujui a.n. Section Head sampai 31 Oktober 2026 untuk:\n• Penerimaan termin (PAY)\n• Validasi rencana cair (VAL)\n\nSyarat pemisahan tugas: akun ini harus terdaftar di Parameter ▸ Pengguna & Peran dan BUKAN berperan PEMBUAT (AR Officer tidak boleh menyetujui pekerjaannya sendiri). Akun pengganti juga perlu akses edit ke spreadsheet ini.' + aktif + '\n\nKosongkan lalu OK untuk mencabut delegasi.', ui.ButtonSet.OK_CANCEL);
  if (q.getSelectedButton() !== ui.Button.OK) return;
  var mail = String(q.getResponseText() || '').trim().toLowerCase();
  if (!mail) { setParamKV_('DELEGASI_EMAIL', '', 'Akun yang boleh menyetujui a.n. Section Head (kosongkan untuk mencabut)'); rebuildAll(); ui.alert('Delegasi dicabut.'); return; }
  var du = ENGINE.userBy(P, mail);
  if (!du) throw new Error(mail + ' belum terdaftar di Parameter ▸ Pengguna & Peran. Daftarkan dulu (peran selain PEMBUAT).');
  if (/^PEMBUAT/i.test(du.peran) || mail === me) throw new Error('Delegasi tidak boleh jatuh ke akun PEMBUAT (' + mail + '). Pilih pejabat lain yang tidak menyusun paket penerimaan / validasi.');
  if (/^PENYETUJU/i.test(du.peran)) throw new Error(mail + ' sudah berperan PENYETUJU; delegasi tidak diperlukan.');
  var nama = du.nama || mail;
  if (ui.alert('Konfirmasi delegasi', 'Akun ' + mail + ' (' + nama + ') diberi wewenang menyetujui a.n. Section Head sampai 31 Oktober 2026 untuk PAY dan VAL.\n\nCap persetujuan tertulis "DISETUJUI a.n. SECTION HEAD oleh ' + nama + ' (DELEGASI)". Pengajuan yang dibuat akun ini sendiri tetap tidak dapat ia setujui. Setelah 31-10-2026 delegasi berhenti sendiri.\n\nOK = aktifkan · Batal = tidak diubah.', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  setParamKV_('DELEGASI_EMAIL', mail, 'Akun yang boleh menyetujui a.n. Section Head (kosongkan untuk mencabut; tidak boleh akun PEMBUAT)');
  setParamKV_('DELEGASI_SAMPAI', '2026-10-31', 'Delegasi berlaku sampai tanggal ini (yyyy-mm-dd)');
  setParamKV_('DELEGASI_JENIS', 'PAY,VAL', 'Jenis yang boleh didelegasikan: PAY, VAL, INV, KP, REK, TUTUP atau SEMUA');
  setParamKV_('DELEGASI_DASAR', 'Persetujuan Pak Anas Afifudin (Section Head) untuk laporan ke pusat Oktober 2026', 'Dasar delegasi (tampil di catatan persetujuan)');
  rebuildAll();
  ui.alert('Delegasi aktif sampai 31 Oktober 2026. Terbitkan ulang aplikasi web (Kelola deployment ▸ Versi baru) bila belum, lalu buka aplikasi dengan akun ' + mail + ': tombol Setujui / Tolak muncul pada paket penerimaan dan validasi rencana cair yang tidak diajukan oleh akun tersebut.');
}

// =====================================================================  v2.5: VALIDASI PER INVOICE, RA/RI GABUNGAN, KINERJA & PERINGATAN DINI
var V25_KV = [
  ['KINERJA_LANCAR_HARI', 15, 'Status kinerja bayar pelanggan: LANCAR bila rata-rata terlambat ≤ n hari'],
  ['KINERJA_PERHATIAN_HARI', 60, 'PERHATIAN bila rata-rata terlambat ≤ n hari'],
  ['KINERJA_KURANG_HARI', 90, 'KURANG LANCAR bila ≤ n hari; di atasnya BERMASALAH'],
  ['EWS_LEWAT_JT_HARI', 60, 'Peringatan MERAH: invoice lewat jatuh tempo lebih dari n hari'],
  ['EWS_CICILAN_HENTI_HARI', 30, 'Peringatan KUNING: cicilan berjalan tanpa pembayaran lebih dari n hari'],
  ['EWS_OVERDUE_PER_PELANGGAN', 2, 'Peringatan MERAH: pelanggan dengan ≥ n invoice overdue'],
  ['EWS_WAITING_SUBMIT_HARI', 14, 'Peringatan KUNING: invoice belum disubmit lebih dari n hari sejak tanggal invoice'],
  ['EWS_KONSENTRASI_PCT', 30, 'Peringatan MERAH: piutang satu pelanggan lebih dari n% seluruh piutang terbuka']];
function setupKinerja_(ss, idx) {
  var sh = setupReport_(ss, 'Kinerja & Peringatan Dini', 'KINERJA PIUTANG & PERINGATAN DINI',
    'DSO, receivables turnover, status kinerja pembayaran per pelanggan, dan peringatan dini (EWS). Dihitung otomatis; ambang diatur di Parameter (KINERJA_… dan EWS_…).', null, idx);
  sh.setTabColor('#B42318');
  return sh;
}
function applyV25() { arRequireMaker_();
  var ui = SpreadsheetApp.getUi(), ss = arSpreadsheet_(), log = [];
  if (ui.alert('Pembaruan v2.5', 'Yang diperbarui:\n• Lembar Validasi Rencana Cair: Section Head memutuskan per invoice (✓ CAIR / ◇ CADANGAN / → BULAN DEPAN / ⏸ HOLD).\n• Pareto: kolom USULAN AR (4 pilihan) + KEPUTUSAN SECTION HEAD (terkunci, otomatis).\n• RA-RI Termin: satu tabel dengan kategori ✓ RA / ◇ CADANGAN / ✚ DI LUAR RA dan dua persentase.\n• Sheet baru Kinerja & Peringatan Dini + 5 tile baru di Dashboard.\n• Urutan tanda tangan dokumen internal: Disetujui (kiri) · Diperiksa (tengah) · Dibuat (kanan).\n\nData yang sudah ada tidak dihapus. Lanjutkan?', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  // usulan lama → 4 pilihan baru
  var mi = ss.getSheetByName('Monitoring Invoice'), mm = colMap_(mi, ENGINE.REG, ENGINE.REG_ALIAS), nU = 0;
  if (mm.validasi && mi.getLastRow() >= ENGINE.REG.firstRow) {
    var rg = mi.getRange(ENGINE.REG.firstRow, mm.validasi, mi.getLastRow() - ENGINE.REG.firstRow + 1, 1), v = rg.getValues();
    v.forEach(function (r) { var u = String(r[0] || '').toUpperCase(); if (!u) return; var c = ENGINE.usulCode(u), nv = c ? ENGINE.DEC[c] : r[0]; if (nv !== r[0]) { r[0] = nv; nU++; } });
    if (nU) rg.setValues(v);
    rg.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(ENGINE.OPT.validasi, true).setAllowInvalid(true).setHelpText('Usulan AR untuk Lembar Validasi. Keputusan akhir oleh Section Head di aplikasi.').build());
    log.push('USULAN RENCANA CAIR: ' + nU + ' isian lama diseragamkan (✓ MASUK RENCANA → ✓ CAIR, ✗ TUNDA → → BULAN DEPAN) + daftar pilihan baru');
  }
  var nk = 0; V25_KV.forEach(function (r) { if (!paramKV_(r[0])) { setParamKV_(r[0], String(r[1]), r[2]); nk++; } });
  if (nk) log.push('Parameter: ' + nk + ' ambang Kinerja & Peringatan Dini');
  var pa = ss.getSheetByName('Pareto'); if (pa) pa.getRange('A2').setValue(REPORT_SUB_['Pareto']);
  var ra = ss.getSheetByName('RA-RI Termin'); if (ra) ra.getRange('A2').setValue(REPORT_SUB_['RA-RI Termin']);
  if (!ss.getSheetByName('Kinerja & Peringatan Dini')) { var rr = ss.getSheetByName('RA-RI Termin'); setupKinerja_(ss, rr ? rr.getIndex() : ss.getSheets().length); log.push('Sheet Kinerja & Peringatan Dini dibuat'); }
  rebuildAll();
  ui.alert('Pembaruan v2.5 selesai', (log.length ? '• ' + log.join('\n• ') : 'Struktur sudah v2.5.') +
    '\n\nLangkah berikut:\n1. Terbitkan ulang aplikasi web (Kelola deployment ▸ ikon pensil ▸ Versi baru).\n2. Aplikasi ▸ Persetujuan ▸ Oktober 2026 ▸ Lembar Validasi Rencana Cair ▸ Ajukan revisi (Rev.01), lalu putuskan per invoice.\n3. Jalankan AR CSU ▸ Kunci struktur & periksa hak akses agar sheet baru ikut terkunci.', ui.ButtonSet.OK);
}

// =====================================================================  v2.6: STATUS DOKUMEN, PEMBATALAN LEWAT APLIKASI, MONITORING INVOICE RINGKAS
function applyV26() { arRequireMaker_();
  var ui = SpreadsheetApp.getUi(), ss = arSpreadsheet_(), log = [];
  if (ui.alert('Pembaruan v2.6', 'Yang diperbarui di Monitoring Invoice:\n• STATUS INVOICE → STATUS DOKUMEN (AKTIF / BATAL), otomatis; pembatalan lewat aplikasi dengan persetujuan Section Head.\n• STATUS PEMBAYARAN dipindah ke sebelah STATUS DOKUMEN.\n• Kolom usulan & validasi rencana cair dan kolom penutupan manual masuk grup yang bisa dibuka-tutup.\n• CATATAN / PENJELASAN → CATATAN AR (OPSIONAL).\n• Menu AR CSU dirapikan (alat sekali jalan dihapus dari menu).\n\nData tidak dihapus. Lanjutkan?', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  var mi = ss.getSheetByName('Monitoring Invoice'), R = ENGINE.REG, map = colMap_(mi, R, ENGINE.REG_ALIAS), maxR = mi.getMaxRows(), n = maxR - R.firstRow + 1;
  var typeRow = R.headerRow - 1;
  var hdr = function (key, text, auto) {
    var c = map[key]; if (!c) return;
    var h = mi.getRange(R.headerRow, c).setValue(text);
    if (auto) { h.setBackground(C.autoHdr); mi.getRange(typeRow, c).setValue('OTOMATIS').setBackground(C.autoHdr).setFontColor('#FFFFFF'); }
  };
  // invoice yang sudah BATAL sebelum v2.6 dicatat di Parameter; sesudah ini BATAL hanya lewat aplikasi (ketikan di sheet tidak berlaku)
  if (map.status_inv && !paramKV_('BATAL_LAMA')) {
    var sv = mi.getRange(R.firstRow, map.status_inv, n, 1).getValues(), nv = map.nomor ? mi.getRange(R.firstRow, map.nomor, n, 1).getValues() : [], lama = [];
    sv.forEach(function (r, i) { if (/^BATAL/i.test(String(r[0])) && nv[i] && String(nv[i][0]).trim()) lama.push(ENGINE.normInv(nv[i][0])); });
    setParamKV_('BATAL_LAMA', lama.length ? lama.join('; ') : '-', 'Invoice yang sudah BATAL sebelum v2.6. Pembatalan berikutnya lewat aplikasi (persetujuan Section Head)');
    log.push(lama.length + ' invoice BATAL lama dicatat di Parameter (BATAL_LAMA)');
  }
  hdr('status_inv', 'STATUS DOKUMEN (AKTIF / BATAL)', true);
  if (map.status_inv) { mi.getRange(R.firstRow, map.status_inv, n, 1).clearDataValidations().setBackground(C.auto); log.push('STATUS INVOICE → STATUS DOKUMEN (otomatis; batal lewat aplikasi)'); }
  hdr('catatan', 'CATATAN AR (OPSIONAL)', false); log.push('CATATAN / PENJELASAN → CATATAN AR (OPSIONAL)');
  // STATUS PEMBAYARAN di sebelah STATUS DOKUMEN
  if (map.status && map.status_inv && map.status !== map.status_inv + 1) { try { mi.moveColumns(mi.getRange(1, map.status, 1, 1), map.status_inv + 1); log.push('STATUS PEMBAYARAN dipindah ke sebelah STATUS DOKUMEN'); } catch (e) { log.push('STATUS PEMBAYARAN belum bisa dipindah (' + e.message + ') — matikan filter di Monitoring Invoice lalu jalankan lagi'); } }
  map = colMap_(mi, R, ENGINE.REG_ALIAS);
  // ALASAN TUTUP MANUAL di sebelah TUTUP MANUAL (dulu ditambahkan di ujung kanan)
  if (map.tutup && map.alasan_tutup && map.alasan_tutup !== map.tutup + 1) { try { mi.moveColumns(mi.getRange(1, map.alasan_tutup, 1, 1), map.tutup + 1); log.push('ALASAN TUTUP MANUAL dipindah ke sebelah TUTUP MANUAL'); } catch (e) { log.push('ALASAN TUTUP MANUAL belum bisa dipindah (' + e.message + ')'); } map = colMap_(mi, R, ENGINE.REG_ALIAS); }
  // grup kolom (tertutup)
  var grp = function (keys, label) {
    var cs = keys.map(function (k) { return map[k]; }).filter(Boolean).sort(function (a, b) { return a - b; });
    if (cs.length !== keys.length || cs[cs.length - 1] - cs[0] + 1 !== cs.length) { log.push('Grup "' + label + '" dilewati (kolom tidak berurutan)'); return; }
    try { if (mi.getColumnGroupDepth(cs[0]) > 0) return; mi.getRange(1, cs[0], 1, cs.length).shiftColumnGroupDepth(1); var g = mi.getColumnGroup(cs[0], 1); if (g) g.collapse(); log.push('Grup tertutup: ' + label); } catch (e) { log.push('Grup "' + label + '" gagal: ' + e.message); }
  };
  grp(['validasi', 'rencana', 'validasi_ok'], 'Rencana cair (diisi lewat Pareto)');
  grp(['tutup', 'alasan_tutup'], 'Penutupan manual');
  arProps_().setProperty('V26_DONE', '1');
  rebuildAll();
  try { onOpen(); } catch (e) { }
  ui.alert('Pembaruan v2.6 selesai', '• ' + log.join('\n• ') + '\n\nLangkah berikut:\n1. Terbitkan ulang aplikasi web (Kelola deployment ▸ ikon pensil ▸ Versi baru).\n2. Jalankan AR CSU ▸ Kunci struktur & periksa hak akses agar STATUS DOKUMEN ikut terkunci.\n3. Muat ulang spreadsheet (F5): menu "Terapkan pembaruan v2.6" sudah hilang.\n\nMembatalkan invoice: aplikasi ▸ Buka / revisi ▸ tombol Batalkan pada baris invoice.', ui.ButtonSet.OK);
}

// =====================================================================  v2.7: USULAN AR DI APLIKASI, PEMBATALAN LEMBAR VALIDASI
function valDraft_(ref) { try { return JSON.parse(arProps_().getProperty('VALDRAFT_' + ref) || 'null'); } catch (e) { return null; } }
/** tulis usulan AR (USULAN, RENCANA TGL CAIR, CATATAN AR) ke Monitoring Invoice sekaligus; tanda rapat pusat & dasar revisi disimpan sebagai draf */
function writeUsulan_(pl) {
  var sh = sheet_('Monitoring Invoice'), R = ENGINE.REG, map = colMap_(sh, R, ENGINE.REG_ALIAS), last = lastDataRow_(sh, R.firstRow, [map.nomor]);
  if (last < R.firstRow) return 0;
  var n = last - R.firstRow + 1, nos = sh.getRange(R.firstRow, map.nomor, n, 1).getValues().map(function (r) { return ENGINE.normInv(r[0]); }), idx = {};
  nos.forEach(function (k, i) { if (k) idx[k] = i; });
  var cols = { validasi: map.validasi, rencana: map.rencana, catatan: map.catatan }, data = {};
  Object.keys(cols).forEach(function (k) { if (cols[k]) data[k] = sh.getRange(R.firstRow, cols[k], n, 1).getValues(); });
  var cnt = 0, pf = {};
  (pl.items || []).forEach(function (x) {
    var i = idx[ENGINE.normInv(x.nomor)]; if (i === undefined) return;
    if (data.validasi) data.validasi[i][0] = x.u ? ENGINE.DEC[x.u] : '';
    if (data.rencana) data.rencana[i][0] = /^\d{4}-\d{2}-\d{2}$/.test(String(x.d || '')) ? toCell_(x.d, 'date') : '';
    if (data.catatan) data.catatan[i][0] = String(x.c || '').trim();
    if (x.p) pf[ENGINE.normInv(x.nomor)] = x.p;
    cnt++;
  });
  Object.keys(cols).forEach(function (k) { if (cols[k] && data[k]) sh.getRange(R.firstRow, cols[k], n, 1).setValues(data[k]); });
  if (pl.ref) arProps_().setProperty('VALDRAFT_' + pl.ref, JSON.stringify({ at: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'), dasar: pl.dasar || '', tglRapat: pl.tglRapat || '', p: pf }));
  return cnt;
}
function apiSaveUsulan(ref, pl) { arRequireMaker_();
  var lock = arLock_(); lock.waitLock(28000); var out;
  try {
    var P = computeLight_().params, me = email_(), u = ENGINE.userBy(P, me);
    if (P.users.length && (!u || (u.peran.indexOf('PEMBUAT') !== 0 && u.peran.indexOf('PENYETUJU') !== 0))) throw new Error('Account ' + (me || '(unknown)') + ' is not registered as PEMBUAT in Parameter ▸ Pengguna & Peran.');
    pl = pl || {}; pl.ref = ref;
    var n = writeUsulan_(pl);
    out = { ok: true, n: n, at: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm') };
  } finally { lock.releaseLock(); }
  out.background = arQueueAfterSave_().background;
  return out;
}
/** PDF lembar validasi bulan itu (termasuk di "Revisi lama") dipindah ke subfolder "Dibatalkan" */
function moveValPdfs_(P, ref) {
  var folder = archiveFolder_(P, 'VAL', ref), pre = 'Validasi rencana cair_' + ref, dest = null, n = 0;
  var take = function (fd) { var it = fd.getFiles(); while (it.hasNext()) { var f = it.next(); if (String(f.getName()).indexOf(pre) === 0) { if (!dest) { var d = folder.getFoldersByName('Dibatalkan'); dest = d.hasNext() ? d.next() : folder.createFolder('Dibatalkan'); } f.moveTo(dest); n++; } } };
  take(folder);
  var rl = folder.getFoldersByName('Revisi lama'); if (rl.hasNext()) take(rl.next());
  return n;
}

// =====================================================================  v2.8.7: SERAH TERIMA DOKUMEN KE AKUNTANSI
function apiHandoverData() { arRequireAccess_();
  var res = computeLight_('memo'), d = ENGINE.handoverData(res.ctx), R = ENGINE.roles(res.params);
  d.penerima = R.PENERIMA ? R.PENERIMA.nama : ''; d.nextNo = ENGINE.hoNextNo(res.ctx);
  return plain_(d);
}
/** catat satu lembar tanda terima untuk dokumen terpilih (hanya yang sudah disahkan & belum diserahkan) */
function apiHandoverCreate(items, opt) { var acc = arRequireMaker_(); opt = opt || {};
  if (!items || !items.length) throw new Error('Select at least one document.');
  var lock = arLock_(), out; lock.waitLock(28000);
  try {
    var ss = arSpreadsheet_(), res = computeLight_('memo'), d = ENGINE.handoverData(res.ctx), want = {};
    items.forEach(function (x) { want[String(x.jenis) + '|' + String(x.ref)] = 1; });
    var L = d.pending.filter(function (x) { return want[x.jenis + '|' + x.ref]; });
    if (!L.length) throw new Error('The selected documents were already handed over or are no longer approved. Reload the list.');
    var sh = ss.getSheetByName(ENGINE.HO.sheet);
    if (!sh) sh = setupDataSheet_(ss, ENGINE.HO, 'SERAH TERIMA DOKUMEN KE AKUNTANSI', 'Dibuat aplikasi (Month-End Close ▸ Handover to Accounting). Satu baris per dokumen; satu NO. SERAH TERIMA per lembar tanda terima yang ditandatangani Akuntansi. Jangan diubah manual — batalkan lewat aplikasi.', 'Log Persetujuan', '#137C8B').sh || ss.getSheetByName(ENGINE.HO.sheet);
    var map = colMap_(sh, ENGINE.HO), no = ENGINE.hoNextNo(res.ctx), tgl = res.ctx.today, R = ENGINE.roles(res.params), pen = String(opt.penerima || (R.PENERIMA ? R.PENERIMA.nama : '')).trim().slice(0, 80);
    var lc = Math.max(sh.getLastColumn(), ENGINE.HO.cols.length), row = lastDataRow_(sh, ENGINE.HO.firstRow, [map.no, map.ref]) + 1; roomFor_(sh, row + L.length - 1);
    var vals = L.map(function (x) { var o = { no: no, tgl: tgl, jenis: x.jenis, ref: x.ref, label: x.label, pelanggan: x.pelanggan, nilai: x.nilai, apv: x.apv, kode: x.kode, oleh: acc.email, penerima: pen, status: 'DISERAHKAN', catatan: x.ulang ? 'Disahkan ulang — menggantikan ' + x.ulang : '' };
      var a = []; for (var k = 0; k < lc; k++) a.push(''); ENGINE.HO.cols.forEach(function (c) { var j = map[c[0]]; if (j) a[j - 1] = arSafeCell_(toCell_(o[c[0]], colFmt_(c))); }); return a; });
    sh.getRange(row, 1, vals.length, lc).setValues(vals);
    out = { ok: true, no: no, tgl: tgl, penerima: pen, oleh: acc.email, rows: L };
  } finally { lock.releaseLock(); }
  out.background = arQueueAfterSave_().background; // kolom SERAH TERIMA AKUNTANSI di Monitoring Invoice & Log Penerimaan
  return plain_(out);
}
/** batalkan satu lembar serah terima (salah cetak / belum jadi diserahkan) — baris tidak dihapus, STATUS = BATAL */
function apiHandoverCancel(no, alasan) { arRequireMaker_();
  alasan = String(alasan || '').trim(); if (alasan.length < 5) throw new Error('Enter a reason (at least 5 characters).');
  var lock = arLock_(), n = 0; lock.waitLock(28000);
  try {
    var sh = sheet_(ENGINE.HO.sheet); if (!sh) throw new Error('No handover records yet.');
    var map = colMap_(sh, ENGINE.HO), last = sh.getLastRow(); if (last < ENGINE.HO.firstRow) throw new Error('Handover ' + no + ' not found.');
    var nos = sh.getRange(ENGINE.HO.firstRow, map.no, last - ENGINE.HO.firstRow + 1, 1).getValues();
    nos.forEach(function (r, i) { if (String(r[0]).trim() === String(no)) { var rr = ENGINE.HO.firstRow + i; sh.getRange(rr, map.status).setValue('BATAL'); sh.getRange(rr, map.catatan).setNumberFormat('@').setValue(('BATAL ' + arNow_() + ': ' + alasan).slice(0, 300)); n++; } });
    if (!n) throw new Error('Handover ' + no + ' not found.');
  } finally { lock.releaseLock(); }
  var q = arQueueAfterSave_(); return { ok: true, n: n, background: q.background };
}
