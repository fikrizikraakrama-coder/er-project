// =====================================================================  v2.9.6: BANK GARANSI (BG) · v2.9.8: penerbit bank / asuransi, nama file scan otomatis, input jaminan terbit tanpa memo
// Permohonan jaminan ke pusat (memo PDF siap tanda tangan Manager) → BG terbit → dipantau sampai dikembalikan / roya / diperpanjang.
// Data: sheet "Monitoring BG" (satu baris per permohonan). Kolom DATA MEMO menyimpan isian memo agar PDF bisa dibuat ulang persis.
var BG_SCHEMA_ = {
  sheet: 'Monitoring BG', headerRow: 5, firstRow: 6,
  cols: [['id', 'ID BG', 'auto', 13], ['tgl', 'TGL MEMO', 'auto', 12, 'date'], ['jenis', 'JENIS JAMINAN', 'auto', 18], ['npp', 'NO. NPP', 'auto', 13],
    ['kepada', 'DITUJUKAN KEPADA (BOUWHEER)', 'auto', 30], ['pekerjaan', 'JENIS PEKERJAAN / PROYEK', 'auto', 36], ['kontrak', 'NO. KONTRAK / SPK', 'auto', 26],
    ['omzet', 'OMZET / NILAI DASAR (Rp)', 'auto', 16, 'money'], ['pct', 'PERSEN JAMINAN', 'auto', 9, 'pct'], ['nilai', 'NILAI JAMINAN (Rp)', 'auto', 16, 'money'],
    ['mulai', 'BERLAKU DARI', 'auto', 12, 'date'], ['sampai', 'BERLAKU S/D', 'auto', 12, 'date'], ['perlu', 'DIPERLUKAN TANGGAL', 'auto', 12, 'date'],
    ['status', 'STATUS', 'auto', 16], ['sisa', 'SISA HARI', 'auto', 9, 'int'], ['no_bg', 'NO. BG', 'auto', 22], ['bank', 'BANK PENERBIT', 'auto', 16], ['penerbit', 'JENIS PENERBIT (BANK / ASURANSI)', 'auto', 12], ['tgl_terbit', 'TGL TERBIT', 'auto', 12, 'date'],
    ['provisi', 'PERKIRAAN PROVISI (Rp)', 'auto', 14, 'money'], ['bg_lama', 'NO. BG LAMA (PERUBAHAN / PERPANJANGAN)', 'auto', 20], ['tgl_kirim', 'TGL KIRIM KE PUSAT', 'auto', 12, 'date'],
    ['tgl_selesai', 'TGL KEMBALI / ROYA', 'auto', 12, 'date'], ['memo', 'LINK MEMO PDF', 'auto', 22], ['scan', 'LINK SCAN BG', 'auto', 22], ['ket', 'KETERANGAN', 'auto', 34],
    ['oleh', 'DIBUAT OLEH', 'auto', 22], ['ubah', 'DIUBAH', 'auto', 15], ['data', 'DATA MEMO (JANGAN DIUBAH)', 'auto', 12]]
};
var BG_ST_ = ['DRAF', 'DIKIRIM KE PUSAT', 'AKTIF', 'DIKEMBALIKAN', 'ROYA', 'DIPERPANJANG', 'DICAIRKAN', 'BATAL'];
function bgSheet_(create) {
  var ss = arSpreadsheet_(), sh = ss.getSheetByName(BG_SCHEMA_.sheet);
  if (!sh && create) {
    sh = setupDataSheet_(ss, BG_SCHEMA_, 'MONITORING BANK GARANSI (BG)', 'Dibuat aplikasi (menu Bank Garansi). Satu baris per permohonan jaminan ke pusat: draf memo → dikirim ke pusat → BG terbit (AKTIF) → dikembalikan / roya / diperpanjang. SISA HARI dihitung otomatis untuk BG AKTIF. Ubah lewat aplikasi agar DATA MEMO tetap sesuai.', 'Monitoring Invoice', '#7C4DC0').sh;
    try { sh.getRange(BG_SCHEMA_.firstRow, 1, Math.max(sh.getMaxRows() - BG_SCHEMA_.firstRow, 1), BG_SCHEMA_.cols.length).setFontFamily('Arial').setFontSize(9).setVerticalAlignment('top'); } catch (e) { }
    try { bgRules_(sh); } catch (e) { }
  }
  if (sh) ensureSchema_(sh, BG_SCHEMA_, null, null);
  if (sh) try { var pr = arProps_(); if (!pr.getProperty('BG_RULE_CAIR')) { var m = colMap_(sh, BG_SCHEMA_), rg = sh.getRange(BG_SCHEMA_.firstRow, 1, sh.getMaxRows() - BG_SCHEMA_.firstRow + 1, BG_SCHEMA_.cols.length), rules = sh.getConditionalFormatRules();
    rules.unshift(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$' + xColA_(m.status) + BG_SCHEMA_.firstRow + '="DICAIRKAN"').setBackground('#F8C9C4').setFontColor('#8A1C12').setBold(true).setRanges([rg]).build()); sh.setConditionalFormatRules(rules); pr.setProperty('BG_RULE_CAIR', '1'); } } catch (e) { }
  return sh;
}
/** warna baris: AKTIF berakhir ≤ 14 hari merah, ≤ 30 hari kuning; selesai abu-abu */
function bgRules_(sh) {
  var m = colMap_(sh, BG_SCHEMA_), L = sh.getMaxRows() - BG_SCHEMA_.firstRow + 1, rg = sh.getRange(BG_SCHEMA_.firstRow, 1, L, BG_SCHEMA_.cols.length);
  var cS = xColA_(m.sisa), cT = xColA_(m.status), r = BG_SCHEMA_.firstRow;
  var rules = sh.getConditionalFormatRules();
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($' + cT + r + '="AKTIF",$' + cS + r + '<>"",$' + cS + r + '<=14)').setBackground('#FDE4E2').setRanges([rg]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($' + cT + r + '="AKTIF",$' + cS + r + '<>"",$' + cS + r + '<=30)').setBackground('#FFF1D6').setRanges([rg]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=OR($' + cT + r + '="DIKEMBALIKAN",$' + cT + r + '="ROYA",$' + cT + r + '="DIPERPANJANG",$' + cT + r + '="BATAL")').setFontColor('#8A97A8').setRanges([rg]).build());
  sh.setConditionalFormatRules(rules);
}
function xColA_(n) { var s = ''; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
function bgRows_() {
  var sh = bgSheet_(false); if (!sh) return [];
  var last = lastDataRow_(sh, BG_SCHEMA_.firstRow, [colMap_(sh, BG_SCHEMA_).id || 1]); if (last < BG_SCHEMA_.firstRow) return [];
  var m = colMap_(sh, BG_SCHEMA_), v = sh.getRange(BG_SCHEMA_.firstRow, 1, last - BG_SCHEMA_.firstRow + 1, sh.getLastColumn()).getValues(), tz = ssTz_();
  var out = [];
  v.forEach(function (r, i) {
    var o = { _row: BG_SCHEMA_.firstRow + i };
    BG_SCHEMA_.cols.forEach(function (c) { var j = m[c[0]], x = j ? r[j - 1] : ''; o[c[0]] = isDate_(x) ? dstr_(x, tz) : x; });
    if (!String(o.id).trim()) return;
    try { o.d = JSON.parse(String(o.data || '{}')); } catch (e) { o.d = {}; } delete o.data;
    out.push(o);
  });
  return out;
}
function bgDefaults_() {
  var P = arParameters_(), R = ENGINE.roles(P), mk = R.MENGETAHUI || {};
  var kv = function (k, d) { var x = paramKV_(k); return x && String(x.val || '').trim() ? String(x.val).trim() : d; };
  return { dari: kv('BG_DARI', 'Manager Concrete Services Unit'), untuk: kv('BG_UNTUK', 'Manager Treasury Sub Division'), up: kv('BG_UP', 'up. Bpk Agung A / Citra'),
    unit: kv('BG_UNIT', 'CONCRETE SERVICE UNIT'), kota: kv('BG_KOTA', 'Jakarta'),
    pemohon: [kv('BG_PEMOHON_1', 'PT. WIJAYA KARYA BETON Tbk.'), kv('BG_PEMOHON_2', 'Tamansari Hive Office Lt 3-5'), kv('BG_PEMOHON_3', 'Jalan D. I Panjaitan Kav. 2 Jakarta 13340')],
    ttd: mk.nama || 'Yusqi Audah Firdaus', jab: mk.jabatan || 'Manager of Concrete Service Unit', today: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd') };
}
function apiBgList() { arRequireAccess_(true);
  return plain_({ rows: bgRows_(), def: bgDefaults_(), sheetUrl: arSpreadsheet_().getUrl() + '#gid=' + (bgSheet_(false) ? bgSheet_(false).getSheetId() : 0) });
}
function bgLock_() { var l = LockService.getUserLock(); if (!l.tryLock(20000)) throw new Error('Penyimpanan lain sedang berjalan. Coba lagi sebentar lagi.'); return l; }
function bgWrite_(sh, row, o) {
  var m = colMap_(sh, BG_SCHEMA_), now = arNow_().slice(0, 16);
  o.ubah = now;
  BG_SCHEMA_.cols.forEach(function (c) {
    var j = m[c[0]]; if (!j || !(c[0] in o) || c[0] === 'sisa') return;
    var v = o[c[0]], f = colFmt_(c), cell = sh.getRange(row, j);
    if (f === 'date') { if (/^\d{4}-\d{2}-\d{2}$/.test(String(v || ''))) cell.setNumberFormat(NUMF.date).setValue(toCell_(v, 'date')); else cell.setValue(''); }
    else if (f === 'money' || f === 'pct') cell.setNumberFormat(NUMF[f]).setValue(v === '' || v === null || v === undefined ? '' : Number(v) || 0);
    else cell.setNumberFormat('@').setValue(c[0] === 'data' ? JSON.stringify(v || {}) : String(v === undefined || v === null ? '' : v));
  });
  if (m.sisa && m.sampai && m.status) sh.getRange(row, m.sisa).setNumberFormat('0').setFormula('=IF(AND(' + xColA_(m.status) + row + '="AKTIF",' + xColA_(m.sampai) + row + '<>""),' + xColA_(m.sampai) + row + '-TODAY(),"")');
}
function bgNextId_(rows, tgl, extra) { var y = String(tgl).slice(0, 4), mx = 0; rows.concat((extra || []).map(function (k) { return { id: k }; })).forEach(function (r) { var k = String(r.id).match(/^BG-(\d{4})-(\d+)$/); if (k && k[1] === y) mx = Math.max(mx, +k[2]); }); return 'BG-' + y + '-' + ('00' + (mx + 1)).slice(-3); }
/* ---------- v2.9.8 nama file scan jaminan terbit: NomorBG_NAMAPELANGGAN_JENIS_PENERBIT (sama dengan bgFileName di aplikasi) */
function bgKode_(jenis) { var j = String(jenis || '').toUpperCase(); return /UANG MUKA/.test(j) ? 'JAMUM' : /PELAKSANA/.test(j) ? 'JAMPEL' : /PEMELIHARAAN/.test(j) ? 'JAMPEM' : /PENAWARAN|TENDER/.test(j) ? 'JAMPEN' : 'JAMINAN'; }
function bgClean_(s) { return String(s || '').replace(/[\\\/:*?"<>|]/g, '-').replace(/\s+/g, ' ').trim(); }
function bgPelName_(s) { return bgClean_(String(s || '').toUpperCase().replace(/[.,]/g, ' ').replace(/\(\s*PERSERO\s*\)|\bPERSERO\b|\bTBK\b/g, ' ').replace(/^\s*(PT|CV|UD|PD)\s+/, '')); }
function bgPenName_(s) { var t = String(s || '').toUpperCase().replace(/[.,]/g, ' ').replace(/\(\s*PERSERO\s*\)|\bPERSERO\b|\bTBK\b/g, ' '); for (var i = 0; i < 3; i++) t = t.replace(/^\s*(PT|BANK|ASURANSI)\s+/, ''); return bgClean_(t); }
function bgFileBase_(no, kepada, jenis, penerbit) { return [bgClean_(no), bgPelName_(kepada), bgKode_(jenis), bgPenName_(penerbit)].filter(String).join('_').slice(0, 180); }
/** folder scan: <arsip>/Bank Garansi/<tahun terbit>/Scan Jaminan Terbit — file lama dengan nama sama dipindah ke "Revisi lama" */
function bgSaveScan_(P, tgl, base, scan) {
  var mf = archiveFolder_(P, 'BG', String(tgl).slice(0, 7)), yf = mf.getParents().next(), it = yf.getFoldersByName('Scan Jaminan Terbit'), f = it.hasNext() ? it.next() : yf.createFolder('Scan Jaminan Terbit');
  var ext = (/\.(\w{2,5})$/.exec(scan.name || '') || [0, 'pdf'])[1].toLowerCase(), files = f.getFiles(), old = null;
  while (files.hasNext()) { var f0 = files.next(); if (String(f0.getName()).replace(/\.\w{2,5}$/, '') === base) { if (!old) { var i2 = f.getFoldersByName('Revisi lama'); old = i2.hasNext() ? i2.next() : f.createFolder('Revisi lama'); } f0.moveTo(old); } }
  return f.createFile(Utilities.newBlob(Utilities.base64Decode(scan.b64), scan.mime || 'application/pdf', base + '.' + ext)).getUrl();
}
function bgNum_(x) { return Math.round(Number(x) || 0); }
function bgProvisi_(nilai, mulai, sampai) { var a = Date.parse(mulai), b = Date.parse(sampai); if (!(nilai > 0) || isNaN(a) || isNaN(b) || b < a) return 0; return Math.round(nilai * 0.00125 * (Math.round((b - a) / 864e5) + 1) / 91); }
/** simpan permohonan (baru / ubah draf). D = isian memo dari aplikasi, atau daftar isian (beberapa jaminan sekaligus → satu baris BG per jaminan) */
function apiBgSave(D) { arRequireMaker_();
  var L = Array.isArray(D) ? D : [D || {}]; if (!L.length) throw new Error('Pilih jenis jaminan.');
  L.forEach(function (d) { if (!String(d.jenisTxt || '').trim()) throw new Error('Pilih jenis jaminan.'); if (!String(d.kepada || '').trim()) throw new Error('Isi "Ditujukan kepada".'); });
  var lock = bgLock_(), ids = [];
  try {
    var sh = bgSheet_(true), rows = bgRows_(), me = email_(), idc = colMap_(sh, BG_SCHEMA_).id || 1, nextRow = lastDataRow_(sh, BG_SCHEMA_.firstRow, [idc]) + 1;
    L.forEach(function (d) {
      var x = d.id ? rows.filter(function (r) { return r.id === d.id; })[0] : null;
      if (d.id && !x) throw new Error('BG ' + d.id + ' tidak ditemukan.');
      if (x && ['DRAF', 'DIKIRIM KE PUSAT'].indexOf(String(x.status)) < 0) throw new Error('BG ' + x.id + ' sudah berstatus ' + x.status + ' — memo tidak bisa diubah lagi.');
      var id = x ? x.id : bgNextId_(rows, d.tglMemo || bgDefaults_().today, ids);
      var row = x ? x._row : nextRow++; roomFor_(sh, row);
      var nilai = bgNum_(d.nilai), o = { id: id, tgl: d.tglMemo || '', jenis: String(d.jenisTxt).toUpperCase(), npp: d.npp || '', kepada: d.kepada, pekerjaan: [d.pekerjaan, d.proyek].filter(String).join(' — '),
        kontrak: d.kontrak || '', omzet: bgNum_(d.omzet), pct: d.pct ? Number(d.pct) / 100 : '', nilai: nilai, mulai: d.mulai || '', sampai: d.sampai || '', perlu: d.perlu || '',
        provisi: bgProvisi_(nilai, d.mulai, d.sampai), bg_lama: d.bgLama || '', ket: [d.jwTxt ? 'Masa berlaku: ' + d.jwTxt : '', d.ket || ''].filter(String).join(' · ').slice(0, 500), data: d };
      if (L.length > 1) o.ket = (o.ket + (o.ket ? ' · ' : '') + 'dibuat bersama ' + L.length + ' jaminan').slice(0, 500);
      if (!x) { o.status = 'DRAF'; o.oleh = me; }
      var now0 = arNow_().slice(0, 16); if (x) { var xd = x.d || {}; ['h', 'prev', 'memoRev', 'amd', 'amdLinks', 'klaim', 'klaimLink'].forEach(function (k) { if (xd[k] !== undefined) d[k] = xd[k]; }); }
      bgHist_(d, x ? 'EDIT' : 'BARU', x ? 'Memo diubah' + (d.memoRev ? ' (Rev.' + ('0' + d.memoRev).slice(-2) + ')' : '') : 'Draf memo dibuat' + (L.length > 1 ? ' (bersama ' + L.length + ' jaminan)' : ''), me, now0); o.data = d;
      bgWrite_(sh, row, o); ids.push(id);
    });
  } finally { lock.releaseLock(); }
  try { arBootBust_(); } catch (e) { }
  return { ok: true, id: ids[0], ids: ids };
}
/** memo PDF (dibuat aplikasi; boleh berisi beberapa memo) disimpan ke arsip Drive: <arsip>/Bank Garansi/<tahun>/<bulan>; link dicatat untuk setiap ID */
function apiBgMemoPdf(ids, b64, name) { arRequireMaker_();
  ids = [].concat(ids || []).map(String); if (!ids.length) throw new Error('ID BG kosong.');
  var res = { params: arParameters_(), ctx: { today: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd') } }, rows = bgRows_(), X = rows.filter(function (r) { return ids.indexOf(r.id) >= 0; });
  if (!X.length) throw new Error('BG ' + ids.join(', ') + ' tidak ditemukan.');
  var folder = archiveFolder_(res.params, 'BG', String(X[0].tgl || res.ctx.today).slice(0, 7)), nm = String(name || ('Memo Permohonan Jaminan ' + ids.join('+'))).replace(/[\\\/:*?"<>|]/g, '-') + (/\.pdf$/i.test(name || '') ? '' : '.pdf');
  var files = folder.getFiles(), old = null; while (files.hasNext()) { var f0 = files.next(), n0 = String(f0.getName()); if (ids.some(function (id) { return n0.indexOf(id + ' ') === 0 || n0.indexOf(id + '+') === 0 || n0.indexOf(id + '_') === 0; })) { if (!old) { var it2 = folder.getFoldersByName('Revisi lama'); old = it2.hasNext() ? it2.next() : folder.createFolder('Revisi lama'); } f0.moveTo(old); } }
  var f = folder.createFile(Utilities.newBlob(Utilities.base64Decode(b64), 'application/pdf', nm));
  var lock = bgLock_(); try { var sh = bgSheet_(true); X.forEach(function (x) { bgWrite_(sh, x._row, { memo: f.getUrl() }); }); } finally { lock.releaseLock(); }
  return { ok: true, url: f.getUrl(), name: f.getName(), folder: folder.getName() };
}
/** v2.9.10 riwayat & langkah terakhir disimpan di DATA MEMO: d.h = [{t, a, by, n}] · d.prev = salinan kolom sebelum perubahan status (untuk "Batalkan langkah terakhir") */
var BG_SNAP_ = ['status', 'tgl_kirim', 'tgl_selesai', 'no_bg', 'bank', 'penerbit', 'tgl_terbit', 'mulai', 'sampai', 'nilai', 'provisi', 'scan', 'ket'];
function bgHist_(d, a, n, me, now) { d.h = (d.h || []).concat([{ t: now, a: a, by: me, n: String(n || '').slice(0, 300) }]).slice(-60); return d; }
function bgIsD_(v) { return /^\d{4}-\d{2}-\d{2}$/.test(String(v || '')); }
/** perubahan status / kendali BG (semua oleh AR):
 *  KIRIM {tgl} · TERBIT {penerbit,no_bg,bank,tgl_terbit,mulai,sampai,nilai,biaya,scan} · SCAN {scan} · NPP {npp}
 *  SELESAI {hasil DIKEMBALIKAN|ROYA, tgl, catatan} · BATAL {alasan} (draf / di pusat / aktif) · PUSAT {alasan} (dikembalikan pusat → draf, memo Rev+1)
 *  AMEND {tgl, mulai, sampai, nilai, scan, ket} (amandemen: nomor sama) · KLAIM {tgl, nilai, alasan, scan} · KOREKSI {field..., alasan} · UNDO {alasan} */
function apiBgStatus(id, act, p) { arRequireMaker_();
  p = p || {}; act = String(act || '').toUpperCase();
  var res = { params: arParameters_(), ctx: { today: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd') } }, lock = bgLock_(), out = { ok: true };
  try {
    var sh = bgSheet_(true), all = bgRows_(), x = all.filter(function (r) { return r.id === id; })[0]; if (!x) throw new Error('BG ' + id + ' tidak ditemukan.');
    var st = String(x.status), o = {}, today = res.ctx.today, me = email_(), now = arNow_().slice(0, 16), d = x.d || {}, ket = function (t) { return (String(x.ket || '') + (x.ket ? ' · ' : '') + t).slice(0, 500); };
    var why = function (min) { var a = String(p.alasan || '').trim(); if (a.length < (min || 5)) throw new Error('Isi alasan (min. ' + (min || 5) + ' characters).'); return a.slice(0, 300); };
    var snap = function () { var s0 = { at: now, act: act }; BG_SNAP_.forEach(function (k) { s0[k] = x[k] === undefined ? '' : x[k]; }); d.prev = s0; };
    var base = function (no, pen, suf) { return bgFileBase_(no || x.no_bg || id, x.kepada, x.jenis, pen || x.bank || '') + (suf || ''); };
    var saveScan = function (suf) {
      if (!p.scan || !p.scan.b64) return '';
      return bgSaveScan_(res.params, p.tgl_terbit || x.tgl_terbit || x.tgl || today, base(p.no_bg, p.bank, suf), p.scan);
    };
    var hist = '';
    if (act === 'KIRIM') { if (st !== 'DRAF') throw new Error('Hanya draf yang bisa ditandai terkirim.'); snap(); o.status = 'DIKIRIM KE PUSAT'; o.tgl_kirim = bgIsD_(p.tgl) ? p.tgl : today; hist = 'Memo dikirim ke pusat (' + o.tgl_kirim + ')'; }
    else if (act === 'PUSAT') {
      if (st !== 'DIKIRIM KE PUSAT') throw new Error('Hanya memo yang sudah dikirim ke pusat yang bisa dikembalikan.');
      var a0 = why(5); snap(); o.status = 'DRAF'; d.memoRev = (+d.memoRev || 0) + 1; hist = 'Dikembalikan pusat untuk revisi memo (Rev.' + ('0' + d.memoRev).slice(-2) + '): ' + a0;
    }
    else if (act === 'TERBIT') {
      if (['DRAF', 'DIKIRIM KE PUSAT', 'AKTIF'].indexOf(st) < 0) throw new Error('BG ' + id + ' berstatus ' + st + '.');
      if (!String(p.no_bg || '').trim()) throw new Error('Isi nomor BG.');
      if (!bgIsD_(p.sampai)) throw new Error('Isi tanggal berakhir BG.');
      snap(); o.status = 'AKTIF'; o.no_bg = String(p.no_bg).trim(); o.bank = String(p.bank || '').trim(); o.penerbit = p.penerbit === 'ASURANSI' ? 'ASURANSI' : 'BANK'; o.tgl_terbit = p.tgl_terbit || ''; o.mulai = p.mulai || x.mulai; o.sampai = p.sampai;
      if (p.nilai) o.nilai = bgNum_(p.nilai); o.provisi = +p.biaya > 0 ? bgNum_(p.biaya) : bgProvisi_(o.nilai || +x.nilai, o.mulai, o.sampai); if (!x.tgl_kirim) o.tgl_kirim = o.tgl_terbit || today;
      var sc = saveScan(); if (sc) o.scan = sc;
      hist = (o.penerbit === 'ASURANSI' ? 'Jaminan' : 'BG') + ' terbit ' + o.no_bg + ' · ' + o.bank + (sc ? ' · scan diupload' : '');
      // BG perpanjangan / perubahan terbit → BG lama otomatis DIPERPANJANG
      if (x.bg_lama) all.forEach(function (r) { if (r.id !== id && String(r.no_bg).trim() && String(r.no_bg).trim() === String(x.bg_lama).trim() && r.status === 'AKTIF') { var dd = bgHist_(r.d || {}, 'DIPERPANJANG', 'Diganti ' + id + ' (' + o.no_bg + ')', me, now); bgWrite_(sh, r._row, { status: 'DIPERPANJANG', tgl_selesai: o.tgl_terbit || today, ket: (String(r.ket || '') + (r.ket ? ' · ' : '') + 'diganti ' + id + ' (' + o.no_bg + ')').slice(0, 500), data: dd }); } });
    }
    else if (act === 'NPP') { var np = String(p.npp || '').trim(); if (!np) throw new Error('Pilih NPP.'); o.npp = np; d.npp = np; hist = 'NPP dihubungkan: ' + np; }
    else if (act === 'SCAN') { o.scan = saveScan(); if (!o.scan) throw new Error('Pilih file scan BG.'); hist = x.scan ? 'Scan diganti' : 'Scan diupload'; }
    else if (act === 'SELESAI') { if (st !== 'AKTIF') throw new Error('Hanya BG aktif yang bisa dikembalikan / roya.'); snap(); o.status = p.hasil === 'ROYA' ? 'ROYA' : 'DIKEMBALIKAN'; o.tgl_selesai = bgIsD_(p.tgl) ? p.tgl : today; if (p.catatan) o.ket = ket(String(p.catatan).slice(0, 200)); hist = titleCase_(o.status) + ' (' + o.tgl_selesai + ')' + (p.catatan ? ': ' + p.catatan : ''); }
    else if (act === 'BATAL') {
      if (['DRAF', 'DIKIRIM KE PUSAT', 'AKTIF'].indexOf(st) < 0) throw new Error('BG ' + id + ' berstatus ' + st + ' dan tidak bisa dibatalkan.');
      var a1 = why(5); snap(); o.status = 'BATAL'; o.ket = ket('Batal: ' + a1); hist = 'Catatan dibatalkan: ' + a1;
    }
    else if (act === 'AMEND') {
      if (st !== 'AKTIF') throw new Error('Hanya BG aktif yang bisa diamandemen.');
      var ch = [];
      if (bgIsD_(p.mulai) && p.mulai !== x.mulai) { o.mulai = p.mulai; ch.push('mulai ' + (x.mulai || '-') + ' → ' + p.mulai); }
      if (bgIsD_(p.sampai) && p.sampai !== x.sampai) { o.sampai = p.sampai; ch.push('s/d ' + (x.sampai || '-') + ' → ' + p.sampai); }
      if (+p.nilai > 0 && bgNum_(p.nilai) !== bgNum_(x.nilai)) { o.nilai = bgNum_(p.nilai); ch.push('nilai ' + bgNum_(x.nilai) + ' → ' + o.nilai); }
      if (!ch.length) throw new Error('Belum ada perubahan — isi tanggal berakhir atau nilai yang baru.');
      if ((o.mulai || x.mulai) && (o.sampai || x.sampai) < (o.mulai || x.mulai)) throw new Error('Tanggal berakhir sebelum tanggal mulai.');
      snap(); d.amd = (+d.amd || 0) + 1; var tag = 'Amd.' + ('0' + d.amd).slice(-2);
      var url = saveScan('_' + tag.replace('.', '')); if (url) { d.amdLinks = (d.amdLinks || []).concat([url]); }
      o.provisi = bgProvisi_(o.nilai || +x.nilai, o.mulai || x.mulai, o.sampai || x.sampai);
      o.ket = ket(tag + (bgIsD_(p.tgl) ? ' ' + p.tgl : '') + ': ' + ch.join(', '));
      hist = tag + (bgIsD_(p.tgl) ? ' (' + p.tgl + ')' : '') + ': ' + ch.join(', ') + (p.ket ? ' · ' + p.ket : '') + (url ? ' · surat amandemen diupload' : '');
    }
    else if (act === 'KLAIM') {
      if (st !== 'AKTIF') throw new Error('Hanya BG aktif yang bisa dicatat klaim.');
      var a2 = why(5); snap(); o.status = 'DICAIRKAN'; o.tgl_selesai = bgIsD_(p.tgl) ? p.tgl : today;
      var url2 = saveScan('_KLAIM'); if (url2) d.klaimLink = url2; d.klaim = { tgl: o.tgl_selesai, nilai: bgNum_(p.nilai) || bgNum_(x.nilai), alasan: a2 };
      o.ket = ket('DICAIRKAN ' + o.tgl_selesai + ' Rp ' + d.klaim.nilai + ': ' + a2); hist = 'Diklaim / dicairkan bouwheer Rp ' + d.klaim.nilai + ' (' + o.tgl_selesai + '): ' + a2;
    }
    else if (act === 'KOREKSI') {
      if (st === 'BATAL') throw new Error('Catatan yang sudah dibatalkan tidak bisa dikoreksi.');
      var a3 = why(5), ch2 = [], F = { no_bg: 'no. BG', bank: 'penerbit', penerbit: 'jenis penerbit', kepada: 'ditujukan kepada', npp: 'NPP', tgl_terbit: 'tgl terbit', mulai: 'mulai', sampai: 's/d', nilai: 'nilai', provisi: 'provisi / premi' };
      Object.keys(F).forEach(function (k) {
        if (p[k] === undefined || p[k] === null) return; var v = ['nilai', 'provisi'].indexOf(k) >= 0 ? bgNum_(p[k]) : String(p[k]).trim(), old = ['nilai', 'provisi'].indexOf(k) >= 0 ? bgNum_(x[k]) : String(x[k] || '').trim();
        if (['tgl_terbit', 'mulai', 'sampai'].indexOf(k) >= 0 && v && !bgIsD_(v)) throw new Error('Tanggal "' + F[k] + '" tidak valid.');
        if (k === 'penerbit') v = v === 'ASURANSI' ? 'ASURANSI' : 'BANK';
        if (v !== old) { o[k] = v; ch2.push(F[k] + ' ' + (old || '-') + ' → ' + (v || '-')); if (k === 'npp' || k === 'kepada') d[k] = v; }
      });
      if (!ch2.length) throw new Error('Tidak ada perubahan.');
      if ((o.no_bg !== undefined || o.bank !== undefined || o.kepada !== undefined) && x.scan) {
        try { var fid = (/[-\w]{25,}/.exec(String(x.scan)) || [])[0], f = DriveApp.getFileById(fid), ext = (/\.(\w{2,5})$/.exec(f.getName()) || [0, 'pdf'])[1];
          f.setName(bgFileBase_(o.no_bg !== undefined ? o.no_bg : x.no_bg, o.kepada !== undefined ? o.kepada : x.kepada, x.jenis, o.bank !== undefined ? o.bank : x.bank) + '.' + ext); ch2.push('nama file scan disesuaikan'); } catch (e) { }
      }
      hist = 'Koreksi data: ' + ch2.join(', ') + ' · alasan: ' + a3;
    }
    else if (act === 'UNDO') {
      var pv = d.prev; if (!pv) throw new Error('Tidak ada langkah yang bisa dibatalkan.');
      var age = (Date.parse(now.replace(' ', 'T')) - Date.parse(String(pv.at).replace(' ', 'T'))) / 864e5; if (!(age <= 7)) throw new Error('Langkah terakhir sudah lebih dari 7 hari dan tidak bisa dibatalkan. Gunakan "Koreksi data".');
      var a4 = why(5); BG_SNAP_.forEach(function (k) { if (k in pv) o[k] = pv[k]; });
      if (pv.act === 'PUSAT') d.memoRev = Math.max(0, (+d.memoRev || 0) - 1);
      if (pv.act === 'AMEND') d.amd = Math.max(0, (+d.amd || 0) - 1);
      if (pv.act === 'KLAIM') delete d.klaim;
      // terbit perpanjangan dibatalkan → BG lama aktif kembali
      if (pv.act === 'TERBIT' && x.bg_lama) all.forEach(function (r) { if (r.status === 'DIPERPANJANG' && String(r.no_bg).trim() === String(x.bg_lama).trim() && String(r.ket || '').indexOf('diganti ' + id) >= 0) { var dd = bgHist_(r.d || {}, 'UNDO', 'Aktif kembali — terbit ' + id + ' dibatalkan', me, now); bgWrite_(sh, r._row, { status: 'AKTIF', tgl_selesai: '', data: dd }); } });
      delete d.prev; hist = 'Langkah terakhir dibatalkan (' + String(pv.act || '').toLowerCase() + ', status kembali ' + String(pv.status || '').toLowerCase() + '): ' + a4;
    }
    else throw new Error('Tindakan tidak dikenal.');
    o.data = bgHist_(d, act, hist, me, now);
    bgWrite_(sh, x._row, o); out.status = o.status || st; out.scan = o.scan || '';
  } finally { lock.releaseLock(); }
  try { arBootBust_(); } catch (e) { }
  return out;
}
function titleCase_(s) { return String(s || '').toLowerCase().replace(/\b\w/g, function (c) { return c.toUpperCase(); }); }
/** jaminan yang sudah terbit tanpa memo dari aplikasi (BG lama / terbit langsung dari pusat) → baris AKTIF. Satu jaminan per panggilan (scan bisa besar). */
function apiBgImport(it) { arRequireMaker_();
  it = it || {}; var s = function (k) { return String(it[k] === undefined || it[k] === null ? '' : it[k]).trim(); }, isD = function (v) { return /^\d{4}-\d{2}-\d{2}$/.test(v); };
  if (!s('no_bg')) throw new Error('Isi nomor BG / jaminan.'); if (!s('kepada')) throw new Error('Isi pelanggan (ditujukan kepada).');
  if (!s('jenis')) throw new Error('Pilih jenis jaminan.'); if (!(+it.nilai > 0)) throw new Error('Isi nilai jaminan.');
  if (!isD(s('sampai'))) throw new Error('Isi tanggal berlaku s/d.'); if (s('mulai') && !isD(s('mulai'))) throw new Error('Tanggal mulai tidak valid.');
  if (s('mulai') && s('sampai') < s('mulai')) throw new Error('Tanggal berakhir sebelum tanggal mulai.');
  var res = { params: arParameters_(), today: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd') }, lock = bgLock_(), id;
  try {
    var sh = bgSheet_(true), rows = bgRows_(), dup = rows.filter(function (r) { return r.status !== 'BATAL' && String(r.no_bg).trim().toUpperCase() === s('no_bg').toUpperCase(); })[0];
    if (dup) throw new Error('Nomor BG ' + s('no_bg') + ' sudah tercatat (' + dup.id + ').');
    var tt = isD(s('tgl_terbit')) ? s('tgl_terbit') : '', pen = s('penerbit').toUpperCase() === 'ASURANSI' ? 'ASURANSI' : 'BANK', nilai = bgNum_(it.nilai);
    id = bgNextId_(rows, tt || res.today); var row = lastDataRow_(sh, BG_SCHEMA_.firstRow, [colMap_(sh, BG_SCHEMA_).id || 1]) + 1; roomFor_(sh, row);
    var o = { id: id, tgl: '', jenis: s('jenis').toUpperCase(), npp: s('npp'), kepada: s('kepada'), pekerjaan: s('pekerjaan'), kontrak: s('kontrak'), omzet: '', pct: '', nilai: nilai,
      mulai: s('mulai'), sampai: s('sampai'), perlu: '', status: 'AKTIF', no_bg: s('no_bg'), bank: s('bank'), penerbit: pen, tgl_terbit: tt,
      provisi: +it.biaya > 0 ? bgNum_(it.biaya) : bgProvisi_(nilai, s('mulai') || tt, s('sampai')), ket: ['Diinput sebagai jaminan terbit (tanpa memo aplikasi)', s('ket')].filter(String).join(' · ').slice(0, 500), oleh: email_(),
      data: bgHist_({ imp: 1, jenis: s('jenis').toUpperCase(), kepada: s('kepada'), npp: s('npp'), pekerjaan: s('pekerjaan'), kontrak: s('kontrak') }, 'BARU', 'Dicatat sebagai jaminan terbit' + (it.scan && it.scan.b64 ? ' · scan diupload' : ''), email_(), arNow_().slice(0, 16)) };
    if (it.scan && it.scan.b64) o.scan = bgSaveScan_(res.params, tt || res.today, bgFileBase_(o.no_bg, o.kepada, o.jenis, o.bank), it.scan);
    bgWrite_(sh, row, o);
  } finally { lock.releaseLock(); }
  try { arBootBust_(); } catch (e) { }
  return { ok: true, id: id };
}
/** v2.9.9 baca scan BG / jaminan asuransi: OCR Google Drive (file diubah sementara menjadi Google Docs, teksnya diambil, lalu dibuang).
 *  scan = {b64, name, mime} → {text}. Pemetaan isian dilakukan di aplikasi (bgOcrParse) agar bisa diperiksa sebelum disimpan. */
function apiBgOcr(scan) { arRequireMaker_();
  if (!scan || !scan.b64) throw new Error('Pilih file scan.');
  var mime = String(scan.mime || 'application/pdf'); if (!/^(application\/pdf|image\/)/.test(mime)) throw new Error('Hanya scan PDF atau gambar yang bisa dibaca.');
  var bytes = Utilities.base64Decode(scan.b64); if (bytes.length > 20 * 1024 * 1024) throw new Error('File terlalu besar untuk dibaca (maks. 20 MB).');
  var bd = 'bgocr' + Date.now(), meta = JSON.stringify({ name: 'OCR sementara - ' + String(scan.name || 'scan').slice(0, 80), mimeType: 'application/vnd.google-apps.document' });
  var head = Utilities.newBlob('--' + bd + '\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n' + meta + '\r\n--' + bd + '\r\nContent-Type: ' + mime + '\r\n\r\n').getBytes(), tail = Utilities.newBlob('\r\n--' + bd + '--').getBytes();
  var tok = ScriptApp.getOAuthToken(), r = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&ocrLanguage=id&fields=id', {
    method: 'post', contentType: 'multipart/related; boundary=' + bd, payload: head.concat(bytes, tail), headers: { Authorization: 'Bearer ' + tok }, muteHttpExceptions: true });
  if (r.getResponseCode() >= 300) throw new Error('Google Drive tidak bisa membaca scan (' + r.getResponseCode() + '): ' + String(r.getContentText()).slice(0, 200));
  var id = JSON.parse(r.getContentText()).id, text = '';
  try {
    var e = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + id + '/export?mimeType=text%2Fplain', { headers: { Authorization: 'Bearer ' + tok }, muteHttpExceptions: true });
    text = e.getResponseCode() < 300 ? e.getContentText('UTF-8') : '';
  } finally { try { DriveApp.getFileById(id).setTrashed(true); } catch (x) { } }
  return { text: String(text || '').slice(0, 60000) };
}
/** ringkas untuk Overview / lonceng (dibaca saat bootstrap; ringan) */
function bgSummary_(today) {
  try {
    var rows = bgRows_(); if (!rows.length) return null;
    var t = Date.parse(today), A = rows.filter(function (r) { return r.status === 'AKTIF'; }), exp = [];
    A.forEach(function (r) { var s = Date.parse(r.sampai); if (!isNaN(s)) { var d = Math.round((s - t) / 864e5); if (d <= 30) exp.push({ id: r.id, jenis: r.jenis, kepada: r.kepada, nilai: +r.nilai || 0, sampai: r.sampai, sisa: d, no_bg: r.no_bg }); } });
    exp.sort(function (a, b) { return a.sisa - b.sisa; });
    var tanpaScan = A.filter(function (r) { return !String(r.scan || '').trim(); }).length;
    return { tanpaScan: tanpaScan, aktif: A.length, nilai: A.reduce(function (s, r) { return s + (+r.nilai || 0); }, 0), exp: exp, proses: rows.filter(function (r) { return r.status === 'DIKIRIM KE PUSAT'; }).length, draf: rows.filter(function (r) { return r.status === 'DRAF'; }).length };
  } catch (e) { return null; }
}
