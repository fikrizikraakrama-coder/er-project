/**
 * AR CSU v2.8.0 — pembaruan kompatibel v2.7.1.
 * Sistem dirancang & dikembangkan oleh Fikri Zikra Akrama.
 * Tidak mengubah saldo pelanggan berdasarkan nama / menghapus transaksi lama.
 */
var AR28_LOCK_DEPTH_ = 0, AR28_NATIVE_LOCK_ = null, AR28_SS_ = null;
// Akun pengguna tidak ditulis di kode: PEMBUAT / PENYETUJU / delegasi diatur di Parameter ▸ Pengguna & Peran.
function arMakerEmails_(p) { return (p || arParameters_()).users.filter(function(u){return u.email&&/^PEMBUAT/.test(u.peran);}).map(function(u){return u.email;}); }
function arSpreadsheet_() {
  if (AR28_SS_) return AR28_SS_;
  var id = PropertiesService.getScriptProperties().getProperty('AR_SOURCE_ID');
  if (id) return AR28_SS_ = SpreadsheetApp.openById(id);
  var ss = SpreadsheetApp.getActive();
  if (!ss || !ss.getSheetByName('Monitoring Invoice')) throw new Error('Run "Pasang pembaruan v2.8" from the main spreadsheet first.');
  return AR28_SS_ = ss;
}
// Script lock works for both web app and triggers. Depth handles internal sync -> rebuild.
function arLock_() {
  var held = false;
  return {
    // v2.8.7: tunggu sampai 90 detik — pembaruan laporan (pengguna lain / jadwal 5 menit) bisa memegang kunci lebih dari 28 detik
    waitLock: function(ms) { if (held) return; if (!AR28_LOCK_DEPTH_) { AR28_NATIVE_LOCK_ = LockService.getScriptLock(); try { AR28_NATIVE_LOCK_.waitLock(Math.max(ms || 0, 90000)); } catch (e) { AR28_NATIVE_LOCK_ = null; throw new Error('The system is busy updating reports. Nothing was saved by this step — try again in a minute.'); } } AR28_LOCK_DEPTH_++; held = true; },
    tryLock: function(ms) { if (held) return true; if (!AR28_LOCK_DEPTH_) { AR28_NATIVE_LOCK_ = LockService.getScriptLock(); if (!AR28_NATIVE_LOCK_.tryLock(ms)) return false; } AR28_LOCK_DEPTH_++; held = true; return true; },
    releaseLock: function() { if (!held) return; held = false; AR28_LOCK_DEPTH_--; if (!AR28_LOCK_DEPTH_ && AR28_NATIVE_LOCK_) { try { SpreadsheetApp.flush(); } finally { AR28_NATIVE_LOCK_.releaseLock(); AR28_NATIVE_LOCK_ = null; arBootBust_(); } } }
  };
}
// Keep legacy document properties readable; new prefixed properties also work in web apps.
function arProps_() {
  var s = PropertiesService.getScriptProperties(), d = null;
  try { d = PropertiesService.getDocumentProperties(); } catch (e) {}
  return {
    getProperty: function(k) { var v = s.getProperty('AR28_' + k); return v !== null ? v : d ? d.getProperty(k) : null; },
    // hanya ditulis ke Script Properties; Document Properties lama tetap terbaca sampai dibersihkan arPropsCleanup_
    setProperty: function(k,v) { s.setProperty('AR28_' + k, String(v)); return this; },
    deleteProperty: function(k) { s.deleteProperty('AR28_' + k); if (d) d.deleteProperty(k); return this; }
  };
}
function arParameters_() {
  var p = sheet_('Parameter');
  return ENGINE.readParams(rangeVals_(p,'A4:C120'), [], [], [], rangeVals_(p,'U4:Y40'), [], []);
}
/** view: true untuk API yang hanya membaca — peran MANAJEMEN (LIHAT) boleh, selain itu ditolak (tidak bisa mengubah apa pun) */
function arRequireAccess_(view) {
  var email = email_(), p = arParameters_(), u = ENGINE.userBy(p,email);
  if (email && u && /^MANAJEMEN/.test(u.peran)) { if (!view) throw new Error('This account is view-only (MANAJEMEN). Nothing can be entered or changed.'); return { email:email, maker:false, approver:false, viewer:true, role:u.peran }; }
  if (!email) throw new Error('Your Google identity could not be read. Sign in with a registered account; the deployment must run as the user accessing the web app.');
  // delegasi aktif (bukan PEMBUAT — dijaga ENGINE.delegasiFor) hanya membuka mode Otorisasi Dokumen
  if (u && !/^(PEMBUAT|PENYETUJU)/.test(u.peran) && ENGINE.delegasiFor(p,email,null,arToday_())) return { email:email, maker:false, approver:false, delegate:true, role:u.peran };
  if (!u || !/^(PEMBUAT|PENYETUJU)/.test(u.peran)) throw new Error('This account has no access to the main AR app. Contact the Preparer.');
  return { email:email, maker:/^PEMBUAT/.test(u.peran), approver:/^PENYETUJU/.test(u.peran), role:u.peran };
}
function arRequireMaker_() { var a = arRequireAccess_(); if (!a.maker) throw new Error('This action is for the Preparer only. The Section Head uses AR Document Approvals.'); return a; }
function arRequireOwner_() {
  var email = email_(), ss = arSpreadsheet_(), owner = ss.getOwner();
  if (!email || !(owner && owner.getEmail().toLowerCase() === email)) throw new Error('Installation and trigger setup must be done by the owner of the main spreadsheet.');
}
function arValidDate_(s) {
  if (typeof s !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(s)) return false;
  var d = new Date(s + 'T00:00:00Z'); return !isNaN(d.getTime()) && d.toISOString().slice(0,10) === s;
}
function arNow_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'); }
function arToday_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function arDigest_(s) { return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,s,Utilities.Charset.UTF_8).map(function(b) { return ('0'+(b&255).toString(16)).slice(-2); }).join(''); }
function arSafeCell_(v) { return typeof v === 'string' && /^[=+@]/.test(v) ? "'" + v : v; }
function arValidateInvoice_(p) {
  if (!p || !arValidDate_(p.tgl) || p.tgl > arToday_()) throw new Error('Invalid or future invoice date.');
  ['dpp','ppn','bruto','umk','retensi'].forEach(function(k) { if (p[k] !== undefined && (!isFinite(Number(p[k])) || Number(p[k]) < 0)) throw new Error('Invalid ' + k + ' value.'); });
  if (!(Number(p.dpp) > 0) || !String(p.uraian || '').trim()) throw new Error('DPP and billing description are required.');
}
function arValidatePayments_(rows, input) {
  checkPayRows_(rows);
  if (rows.length > 100) throw new Error('Maximum 100 allocations per save.');
  var computed = ENGINE.compute(Object.assign({},input,{light:true}));
  var today = input.today, by = {}, p = computed.params, groups = {}, ctx = computed.ctx;
  ctx.invs.forEach(function(i) { by[ENGINE.normInv(i.nomor)] = i; });
  rows.forEach(function(r,i) {
    var n = Number(r.nominal), inv = by[ENGINE.normInv(r.inv)], jt = String(r.jenis_trx || '').toUpperCase(), pb = jt === 'PINDAH BUKU', kom = jt === 'KOMPENSASI', px = Number(r.pph_bupot) || 0;
    if (!isFinite(n) || !Number.isSafeInteger(n) || Math.abs(n) > 1e15) throw new Error('Row '+(i+1)+': amount must be a valid whole rupiah value.');
    if (!arValidDate_(r.tgl) || r.tgl > today) throw new Error('Row '+(i+1)+': invalid or future date.');
    if (!inv || inv._batal) throw new Error('Row '+(i+1)+': invoice not found or cancelled.');
    if (inv.tgl && r.tgl < inv.tgl) throw new Error('Payment date is before the invoice date.');
    if (inv._closedManual) throw new Error('The invoice has been manually closed. Reconcile the manual close before recording a new receipt.');
    if (r.jenis_trx && !pb && !kom && jt !== 'TRANSFER') throw new Error('Unrecognized transaction type.');
    if (kom && !(n > 0)) throw new Error('Row '+(i+1)+': compensation amount must be positive.');
    if (kom && String(r.ref || '').trim().length < 3) throw new Error('Row '+(i+1)+': enter the compensation BA / memo number as the reference.');
    if (kom && String(r.rekening || '').trim()) throw new Error('Row '+(i+1)+': a compensation has no receiving bank account.');
    if (r.pph_bupot !== undefined && r.pph_bupot !== '' && (!isFinite(px) || px < 0 || !Number.isSafeInteger(px))) throw new Error('Row '+(i+1)+': WHT per withholding slip must be a whole rupiah amount ≥ 0.');
    if (px > 0 && pb) throw new Error('WHT per withholding slip cannot be entered on an offset / reclass row.');
    if (px > 0 && inv._a && px > inv._a.netto) throw new Error('Row '+(i+1)+': WHT per withholding slip exceeds the invoice net amount.');
    if (r.alokasi && ENGINE.OPT.alokasi.indexOf(String(r.alokasi).toUpperCase()) < 0) throw new Error('Allocation must be OTOMATIS, NETTO or PPN.');
    if (!pb && !kom && !ENGINE.rekOf(p,r.rekening)) throw new Error('Bank account not found in Parameter.');
    if (r.transfer !== undefined && r.transfer !== '' && (!isFinite(Number(r.transfer)) || Number(r.transfer) < 0)) throw new Error('Invalid transfer amount.');
    if (String(r.ref || '').toUpperCase() === 'MIGRASI') throw new Error('Reference MIGRASI is reserved for legacy data.');
    if (pb && !String(r.ref || '').trim()) throw new Error('An offset / reclass reference is required.');
    var key = (pb ? 'PB|' : kom ? 'KOM|' : 'TRF|') + String(r.ref || '_TANPA_REF_').trim(), g = groups[key] || (groups[key] = {pb:pb,kom:kom,rows:[],total:0});
    g.rows.push(r); g.total += n; r.nominal = n;
  });
  Object.keys(groups).forEach(function(k) {
    var g = groups[k], a = g.rows, first = a[0];
    if (a.some(function(r) { return r.tgl !== first.tgl || (!g.pb && !g.kom && ENGINE.rekOf(p,r.rekening).label !== ENGINE.rekOf(p,first.rekening).label); })) throw new Error('All rows under one reference must use the same date and bank account.');
    if (g.pb) {
      var invoices = {}; a.forEach(function(r) { invoices[r.inv]=1; });
      if (a.length < 2 || Object.keys(invoices).length < 2 || Math.abs(g.total) > 1 || !a.some(function(r){return r.nominal<0;}) || !a.some(function(r){return r.nominal>0;})) throw new Error('Each offset / reclass reference must balance and have different source and target invoices.');
    } else if (!g.kom) {
      var totals = a.map(function(r) { return Number(r.transfer) || 0; }).filter(function(v){return v>0;});
      if (a.length > 1 && (!String(first.ref || '').trim() || !totals.length)) throw new Error('A split transfer requires a bank reference and the total transfer amount.');
      if (totals.some(function(v) { return Math.abs(v-g.total)>1; })) throw new Error('Total allocations differ from the transfer amount.');
    }
  });
}
/** baris UTANG_KOMPENSASI di Parameter ▸ Pemetaan Akun SAP (GL menunggu konfirmasi Akuntansi) */
function arEnsureAkunKompensasi_() {
  var par=sheet_('Parameter'); if(!par||norm_(par.getRange('O3').getValue())!=='KODE') return;
  var ak=par.getRange('O4:O80').getValues().map(function(r){return norm_(r[0]);}),last=3; if(ak.indexOf('UTANG_KOMPENSASI')>=0) return;
  ak.forEach(function(k,i){if(k)last=4+i;}); if(last>=80) return;
  par.getRange(last+1,15,1,5).setNumberFormat('@').setValues([['UTANG_KOMPENSASI','Utang Usaha (kompensasi dengan piutang)','','','MOHON KONFIRMASI Akuntansi — akun utang ke pelanggan yang dikompensasi (mis. klaim solar)']]);
}
function arPaymentCanonical_(rows) {
  var keys = ENGINE.LOG.cols.filter(function(c) { return colType_(c)==='manual'; }).map(function(c){return c[0];});
  return JSON.stringify(rows.map(function(r) { return keys.map(function(k) { var v=r[k]; if (['nominal','transfer','pph_bupot'].indexOf(k)>=0) return Number(v)||0; return v === undefined || v === null ? '' : String(v).trim(); }); }));
}
function arSavePayments_(rows, requestId, opt) {
  var access = arRequireMaker_(), lock = arLock_(), saved, npps = [], months = [];
  rows = JSON.parse(JSON.stringify(rows || []));
  rows.forEach(function(r) { r.inv = ENGINE.normInv(r.inv); });
  // ditandai sebelum sidik dihitung, agar pengiriman ulang dengan alasan yang sama tetap dikenali sebagai permintaan yang sama
  var koreksi = arCheckPeriod_(rows.map(function(r){return r.tgl;}), (opt||{}).periodeAlasan, 'Cash receipt');
  if (koreksi) rows.forEach(function(r) { r.ket = (koreksi + ' ' + String(r.ket || '')).trim(); });
  var hash = arDigest_(arPaymentCanonical_(rows)), token = requestId || 'legacy-' + hash;
  if (!/^[A-Za-z0-9_-]{12,100}$/.test(token)) throw new Error('Invalid request ID. Reload the app.');
  lock.waitLock(28000);
  try {
    var sh = sheet_('Log Penerimaan');
    // v2.8.7: kolom PPh SESUAI BUPOT & akun UTANG_KOMPENSASI dibuat otomatis saat pertama dipakai
    if (!colMap_(sh,ENGINE.LOG).pph_bupot && rows.some(function(r){return Number(r.pph_bupot)>0;})) ensureSchema_(sh,ENGINE.LOG,null,4);
    if (rows.some(function(r){return String(r.jenis_trx||'').toUpperCase()==='KOMPENSASI';})) arEnsureAkunKompensasi_();
    var input = readAll_(), map = colMap_(sh,ENGINE.LOG), cols = ENGINE.LOG.cols;
    if (!map.request_id || !map.request_hash || !map.request_user) throw new Error('The v2.8 update has not been installed on the spreadsheet.');
    var ix = function(k) { return ENGINE.colIndex(ENGINE.LOG,k); };
    var existing = input.log.filter(function(r) { return String(r[ix('request_id')] || '') === token; });
    if (existing.length) {
      if (existing.length !== rows.length || existing.some(function(r) { return r[ix('request_hash')] !== hash || r[ix('request_user')] !== access.email; })) throw new Error('This save ID was already used for different or incomplete data. Check the Log Penerimaan sheet; do not re-enter it as a new transaction.');
      var stored = existing.map(function(r){var o={};cols.forEach(function(c,j){o[c[0]]=r[j];});return o;});
      if (arDigest_(arPaymentCanonical_(stored)) !== hash) throw new Error('The transaction was saved but later modified. Check the Log Penerimaan sheet.');
      saved = {ok:true,saved:true,duplicate:true,requestId:token,invoices:[],payments:[],key:ENGINE.payKey(stored[stored.length-1])};
    } else {
      if (arProps_().getProperty('VOID_'+token)) throw new Error('The earlier request was cancelled. Resend with a new request ID.');
      arValidatePayments_(rows,input);
      // Existing bank references identify a transfer, including all split allocations.
      var refs = {}; rows.forEach(function(r){if(r.ref) refs[String(r.ref).trim()]=true;});
      if (input.log.some(function(r) { return refs[String(r[ix('ref')] || '').trim()]; })) throw new Error('This bank reference is already recorded. Check the Log Penerimaan sheet before recording it again.');
      if (input.log.some(function(r) { return r[ix('request_hash')] === hash; })) throw new Error('Identical payment data is already recorded. Check the earlier transaction.');
      var maxId = 0;
      input.log.forEach(function(r) { var m=String(r[ix('idp')] || '').match(/^PNR-(\d+)$/); if(m) maxId=Math.max(maxId,+m[1]); });
      // Include legacy blank IDs deterministically, so new IDs do not collide after rebuild.
      input.log.forEach(function(r) { if (!r[ix('idp')] && r.some(function(v){return v!=='' && v!==null && v!==undefined;})) maxId++; });
      var target = lastDataRow_(sh,ENGINE.LOG.firstRow,[map.tgl,map.inv,map.nominal])+1, width=sh.getLastColumn();
      if(target+rows.length-1>sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(),target+rows.length-sh.getMaxRows()+10);
      var values=rows.map(function(r) {
        r.idp='PNR-'+('00000'+(++maxId)).slice(-Math.max(5,String(maxId).length)); r.request_id=token;r.request_hash=hash;r.request_user=access.email;
        var a=Array(width).fill(''); cols.forEach(function(c) { var j=map[c[0]]; if(j && (colType_(c)==='manual' || ['idp','request_id','request_hash','request_user'].indexOf(c[0])>=0)) a[j-1]=arSafeCell_(toCell_(r[c[0]],colFmt_(c))); }); return a;
      });
      if (map.ref) sh.getRange(target,map.ref,values.length,1).setNumberFormat('@'); // referensi bank tetap teks (nol di depan tidak hilang)
      sh.getRange(target,1,values.length,width).setValues(values);
      // v2.8.7: tanggal bayar dibaca ulang — bila kosong / bukan tanggal ditulis lagi per sel (kasus: tanggal tidak muncul di sheet),
      // dan dicadangkan agar pembaruan laporan bisa memulihkannya
      if (map.tgl) try { SpreadsheetApp.flush(); var tg0=sh.getRange(target,map.tgl,rows.length,1).getValues(), bad=[];
        rows.forEach(function(r,k){ if(!isDate_(tg0[k][0])) { bad.push((target+k)+':'+String(tg0[k][0])); sh.getRange(target+k,map.tgl).setNumberFormat(NUMF.date).setValue(toCell_(r.tgl,'date')); } });
        arTglBackup_(rows.map(function(r){return [r.idp,r.tgl];}));
        if (bad.length) { SpreadsheetApp.flush(); var tg1=sh.getRange(target,map.tgl,rows.length,1).getValues(); AR_TGL_NOTE_='Payment date was not written on the first attempt (row '+bad.join(', ')+'); it was rewritten — please check the TANGGAL BAYAR cell.'; arDiag_('TGL_BAYAR', bad.join(', ')+' → '+tg1.map(function(x){return isDate_(x[0])?'ok':'GAGAL:'+String(x[0]);}).join(',')); } } catch(e) { arDiag_('TGL_BAYAR', 'error '+e.message); }
      // Metadata and all allocations are in the same write; retries use those persisted rows.
      saved={ok:true,saved:true,duplicate:false,requestId:token,invoices:[],payments:[],key:ENGINE.payKey(rows[rows.length-1])};
    }
    months=rows.map(function(r){return r.tgl.slice(0,7);});
    var keys=rows.map(function(r){return r.inv;});
    var ctx=ENGINE.compute(Object.assign({},input,{light:true})).ctx;
    npps=ctx.invs.filter(function(i){return keys.indexOf(ENGINE.normInv(i.nomor))>=0;}).map(function(i){return i.npp;});
    // v2.8.7: NAMA PELANGGAN & NPP (kolom otomatis) langsung diisi — baris baru tidak tersembunyi oleh filter yang menyembunyikan sel kosong
    if (!saved.duplicate && typeof target==='number') try { var byN={}; ctx.invs.forEach(function(i){byN[ENGINE.normInv(i.nomor)]=i;});
      rows.forEach(function(r,k){var i=byN[ENGINE.normInv(r.inv)];if(!i)return;if(map.pelanggan)sh.getRange(target+k,map.pelanggan).setValue(i.pelanggan);if(map.npp)sh.getRange(target+k,map.npp).setNumberFormat('@').setValue(i.npp);}); } catch(e) {}
  } finally { lock.releaseLock(); }
  var post=arQueueAfterSave_(npps,months); saved.reportPending=post.reportPending;saved.background=post.background;saved.warnings=post.warnings;
  if (AR_TGL_NOTE_) { saved.warnings=saved.warnings.concat([AR_TGL_NOTE_]); AR_TGL_NOTE_=''; }
  if (koreksi && !saved.duplicate) { var rf=arRefreshClosedMonths_(rows.map(function(r){return r.tgl;}),koreksi); saved.warnings=saved.warnings.concat(rf.warnings); saved.snapshots=rf.ids;
    var nw=arNotify_('Koreksi periode tertutup — penerimaan '+rows[0].tgl,['Pembuat: '+access.email,'Tanda: '+koreksi,'ID permintaan: '+token,'Snapshot direvisi: '+(rf.ids.join(', ')||'-'),''].concat(rows.map(function(r){return '• '+r.tgl+' · '+r.inv+' · Rp '+ENGINE.fmtRp(r.nominal)+(r.ref?' · ref '+r.ref:'');}))); if(nw) saved.warnings.push(nw); }
  try { var b=ENGINE.bootstrap(computeLight_(),access.email); saved.invoices=b.invoices.filter(function(i){return keys.indexOf(ENGINE.normInv(i.nomor))>=0;}); saved.payments=b.payments.filter(function(p){return (rows[0].ref && p.ref===rows[0].ref) || (!rows[0].ref && p.key===saved.key);}); } catch(e) { saved.reportPending=true;saved.warnings.push('Summary not loaded yet: '+e.message); }
  return plain_(saved);
}
/** cadangan tanggal bayar per ID penerimaan (maks. 400 terakhir) — dipakai arRepairPayDates_ bila sel tanggal kosong */
var AR_TGL_NOTE_ = '';
function arTglBackup_(pairs) {
  var pr=arProps_(), m={}; try{m=JSON.parse(pr.getProperty('PAY_TGL')||'{}');}catch(e){m={};}
  pairs.forEach(function(x){ if(x[0]&&x[1]) m[x[0]]=x[1]; });
  var k=Object.keys(m).sort(); if(k.length>400) k.slice(0,k.length-400).forEach(function(x){delete m[x];});
  try{pr.setProperty('PAY_TGL',JSON.stringify(m));}catch(e){}
}
/** catatan diagnosa singkat (Script Properties DIAG_LOG, 30 terakhir) */
function arDiag_(tag,msg) { var pr=arProps_(),L=[]; try{L=JSON.parse(pr.getProperty('DIAG_LOG')||'[]');}catch(e){} L.push(arNow_()+' '+tag+' '+String(msg).slice(0,300)); try{pr.setProperty('DIAG_LOG',JSON.stringify(L.slice(-30)));}catch(e){} }
function arLihatDiagnosa() { arRequireMaker_(); var L=[]; try{L=JSON.parse(arProps_().getProperty('DIAG_LOG')||'[]');}catch(e){}
  SpreadsheetApp.getUi().alert('Catatan diagnosa', L.length ? L.slice().reverse().join('\n') : 'Belum ada catatan. Bila tanggal bayar sempat tidak tertulis, kejadiannya dicatat di sini.', SpreadsheetApp.getUi().ButtonSet.OK); }
/** TANGGAL BAYAR kosong pada baris dari aplikasi → diisi dari cadangan (dipanggil rebuildAll_) */
function arRepairPayDates_(ss) {
  var sh=ss.getSheetByName('Log Penerimaan'); if(!sh) return 0; var map=colMap_(sh,ENGINE.LOG), f0=ENGINE.LOG.firstRow, last=sh.getLastRow();
  if(!map.tgl||!map.idp||last<f0) return 0;
  var m={}; try{m=JSON.parse(arProps_().getProperty('PAY_TGL')||'{}');}catch(e){return 0;}
  var n=last-f0+1, tg=sh.getRange(f0,map.tgl,n,1).getValues(), id=sh.getRange(f0,map.idp,n,1).getValues(), fixed=0;
  for(var i=0;i<n;i++){ var k=String(id[i][0]||'').trim(); if(k && m[k] && (tg[i][0]===''||tg[i][0]===null)) { sh.getRange(f0+i,map.tgl).setNumberFormat(NUMF.date).setValue(toCell_(m[k],'date')); fixed++; } }
  if(fixed) arDiag_('TGL_BAYAR','dipulihkan saat pembaruan: '+fixed+' baris');
  return fixed;
}
function arAfterSave_(npps,months,noDerived) {
  var out={reportPending:false,warnings:[]};
  // v2.8.7: tanpa merapikan format sheet sumber (berat; dikerjakan jadwal harian 06.00 & menu) → kunci dilepas lebih cepat
  try { rebuildAll_({ format: false }); } catch(e) { out.reportPending=true;out.warnings.push('Main reports: '+e.message); }
  if (!noDerived) try {
    var a=paramKV_('FILE_AKUNTANSI_ID'),b=paramKV_('FILE_PE_ID'),c=paramKV_('FILE_PAJAK_ID');
    // sinkron lain sedang berjalan bukan kegagalan: file Akuntansi/P&E/Pajak ikut diperbarui oleh sinkron itu / jadwal berikutnya
    if ((a&&a.val)||(b&&b.val)||(c&&c.val)) { var r=syncTurunan_({npps:npps,months:months,noPull:true}),re=r.err.filter(function(x){return !/another sync is running|system busy/.test(x);}); if(re.length) {out.reportPending=true;out.warnings=out.warnings.concat(re);} }
  } catch(e) {out.reportPending=true;out.warnings.push('Derived files: '+e.message);}
  return out;
}
function apiRefreshReports() { arRequireMaker_(); return arAfterSave_(null,null); }
// ---------------- v2.8.4: simpan cepat — laporan & file turunan diperbarui di latar belakang (apiAfterSave dipanggil aplikasi tanpa menunggu)
function arQueueAfterSave_(npps,months) {
  var pr=PropertiesService.getScriptProperties(),q={npps:[],months:[]};
  try{q=JSON.parse(pr.getProperty('AFTER_SAVE')||'{"npps":[],"months":[]}');}catch(e){}
  (npps||[]).forEach(function(n){if(n&&q.npps.indexOf(n)<0)q.npps.push(n);});(months||[]).forEach(function(m){if(m&&q.months.indexOf(m)<0)q.months.push(m);});
  try{pr.setProperty('AFTER_SAVE',JSON.stringify(q).slice(0,8000));}catch(e){}
  arMarkDirty_(); // jadwal 5 menit tetap memperbarui bila aplikasi ditutup sebelum selesai
  return {reportPending:false,background:true,warnings:[]};
}
function apiAfterSave() {
  // v2.8.6: juga dipanggil aplikasi Section Head setelah Approve/Reject; file turunan (Akuntansi, P&E, Pajak) hanya dari akun Pembuat
  //   (Section Head belum tentu punya akses ke file itu — jadwal per jam tetap memperbarui)
  var acc=arRequireAccess_(),pr=PropertiesService.getScriptProperties(),q=null,t=pr.getProperty(AR_DIRTY_KEY_);
  try{q=JSON.parse(pr.getProperty('AFTER_SAVE')||'null');}catch(e){}
  if(!q&&!t)return {reportPending:false,warnings:[],skipped:true}; // sudah diperbarui oleh panggilan lain
  pr.deleteProperty('AFTER_SAVE');
  var out=arAfterSave_(q&&q.npps.length?q.npps:null,q&&q.months.length?q.months:null,!acc.maker);
  if(!out.reportPending&&t&&pr.getProperty(AR_DIRTY_KEY_)===t)pr.deleteProperty(AR_DIRTY_KEY_);
  return out;
}
// ---------------- v2.8.4: cache data aplikasi (CacheService, dipecah per 30.000 karakter)
function arBootBust_(){try{PropertiesService.getScriptProperties().setProperty('BOOT_VER',String(Date.now()));}catch(e){}}
function arBootKey_(email){var v='0';try{v=PropertiesService.getScriptProperties().getProperty('BOOT_VER')||'0';}catch(e){}return 'BOOT_'+Utilities.base64EncodeWebSafe(String(email||'-')).slice(0,60)+'_'+v;}
function arCacheGet_(key){try{var c=CacheService.getScriptCache(),n=c.get(key+'#n');if(!n)return null;var ks=[];for(var i=0;i<+n;i++)ks.push(key+'#'+i);var m=c.getAll(ks),s='';for(var j=0;j<ks.length;j++){if(m[ks[j]]==null)return null;s+=m[ks[j]];}return s;}catch(e){return null;}}
function arCachePut_(key,s,ttl){try{var c=CacheService.getScriptCache(),o={},n=Math.ceil(s.length/30000);if(!n||n>90)return;for(var i=0;i<n;i++)o[key+'#'+i]=s.substr(i*30000,30000);o[key+'#n']=String(n);c.putAll(o,ttl||300);}catch(e){}}
function arChartsCached_(){try{var j=arProps_().getProperty('DASH_CHARTS');return j?JSON.parse(j):null;}catch(e){return null;}}
function arDashboard_(ctx) {
  var open=0,overdue=0,received=0,month=ctx.today.slice(0,7),pending=0;
  ctx.invs.forEach(function(i){if(!i._batal&&!i._res.closed){open+=i._res.sisa;if(i._res.jt&&i._res.jt<ctx.today)overdue+=i._res.sisa;}});
  ctx.pays.forEach(function(p){if(p._inv&&!p._inv._batal&&!p._pb&&p.nominal>0&&p.tgl.slice(0,7)===month)received+=p.nominal;});
  ctx.appr.forEach(function(a){if(a.status==='MENUNGGU')pending++;});
  return {open:open,overdue:overdue,received:received,pending:pending};
}

// ---------------- Formatting: values/formulas are never rewritten here.
function arLeftLabel_(value) {
  var label=String(value||'').toUpperCase().replace(/\s+/g,' ').trim();
  return /^(NO\.?|NOMOR)\s+INVOICE(?:$|[\s.(·])/.test(label) || /^(REF\s*\/\s*BUKTI|BUKTI\s*\/\s*REF)$/.test(label) || /^(URAIAN|DESKRIPSI|PRODUK)(?:$|\s)/.test(label) || /^(NAMA )?PELANGGAN(?:$|\s*\()/.test(label) || /^(NAMA )?PROYEK(?:$|\s*\()/.test(label) || /^INVOICE\s+KE(?:[\s\-:]|$)/.test(label) || /^NAMA DI INVOICE(?:$|\s*\()/.test(label);
}
/** perataan laporan: NO di tengah · kolom berisi tulisan rata kiri · tanggal, jumlah (bilangan bulat) & persen di tengah · nilai uang rata kanan.
 *  Laporan ditulis ulang setiap pembaruan, jadi aturan ini (bukan perataan manual) yang berlaku. */
function arAlignBlock_(sh,b,vals,pw,fmts) {
  var headers={},special={},labels={},tile=/tile/.test(JSON.stringify(b.styles));
  b.styles.forEach(function(s){for(var k=0;k<(s.n||1);k++){if(s.s==='hdr')headers[s.r+k]=true;if(s.s==='lbl')labels[s.r+k]=true;if(['title','note','stamp','signname','lbl'].indexOf(s.s)>=0)special[s.r+k]=true;}});
  var hdrTxt={};Object.keys(headers).forEach(function(r){(b.rows[r]||[]).forEach(function(v,j){var c=physCol_(b,j)-b.c;if(String(v).trim())hdrTxt[c]=String(v).replace(/\s+/g,' ').trim().toUpperCase();});});
  var fmtOf=function(c){for(var j=0;j<vals[0].length;j++)if(physCol_(b,j)-b.c===c)return b.fmt[j]||'';return '';};
  var isDate=function(v){return v instanceof Date||/^\d{4}-\d{2}-\d{2}$/.test(String(v))||/^\d{1,2}[-\/ ](\d{1,2}|[A-Za-z]{3,9})[-\/ ]\d{4}$/.test(String(v));};
  var colAl={};
  for(var c=0;c<pw;c++){
    var num=0,txt=0,dt=0,ints=0,big=false,f=fmtOf(c),h=hdrTxt[c]||'';
    vals.forEach(function(row,r){if(headers[r]||special[r])return;var v=row[c];if(v===''||v===null||v===undefined)return;
      if(typeof v==='number'){num++;if(Math.round(v)===v)ints++;if(Math.abs(v)>=100000)big=true;}
      else if(isDate(v))dt++;else if(/^Rp\s?-?[\d.]+/.test(String(v)))num++,big=true;else if(!/^[-–—]$/.test(String(v).trim()))txt++;});
    if(/^NO\.?$/.test(h)||h==='#')colAl[c]='center';
    else if(f==='date'||(dt&&dt>=num&&dt>=txt))colAl[c]='center';
    else if(num>txt){colAl[c]=f==='money'||f==='money0'||f==='num2'||(!f&&(big||ints<num))?'right':'center';}
    else if(txt)colAl[c]='left';
    else colAl[c]='center';
  }
  var align=vals.map(function(row,r){return row.map(function(v,c){
    if(headers[r]||tile)return 'center';
    if(b.align&&b.align[c])return b.align[c];
    if(labels[r]&&[0,1,3,7,8].indexOf(c)>=0)return 'left';
    if(special[r]&&c===0)return 'left';
    if(typeof v==='string'&&/^TOTAL\b/i.test(v.trim()))return 'left';
    return colAl[c];});});
  sh.getRange(b.r,b.c,vals.length,pw).setHorizontalAlignments(align).setVerticalAlignment('middle');
  Object.keys(labels).forEach(function(r){[0,1,3,7,8].forEach(function(c){if(c<pw)sh.getRange(b.r+(+r),b.c+c,1,1).setWrap(true);});});
  Object.keys(colAl).forEach(function(c){if(colAl[c]==='left'&&!tile&&/CATATAN|KETERANGAN|TINDAK LANJUT|URAIAN|MASALAH|SARAN/.test(hdrTxt[c]||''))sh.getRange(b.r,b.c+(+c),vals.length,1).setWrap(true);});
}
function arFormatSourceSheets_(ss) {
  var schemas=[ENGINE.REG,ENGINE.LOG,ENGINE.MPEL,ENGINE.MPRO,ENGINE.ADD,ENGINE.TAG,ENGINE.NOTE,ENGINE.APPR,ENGINE.MEMO,ENGINE.PEREF];
  schemas.forEach(function(sc){var sh=ss.getSheetByName(sc.sheet);if(!sh||sh.getLastRow()<sc.firstRow)return;var n=sh.getLastRow()-sc.firstRow+1,map=colMap_(sh,sc,sc===ENGINE.REG?ENGINE.REG_ALIAS:sc===ENGINE.MPRO?ENGINE.MPRO_ALIAS:null);
    sh.getRange(sc.firstRow,1,n,sh.getLastColumn()).setVerticalAlignment('middle');
    sc.cols.forEach(function(c){var j=map[c[0]];if(!j)return;var rg=sh.getRange(sc.firstRow,j,n,1),f=colFmt_(c),numeric=/^(money|money0|pct|int|num2)$/.test(f||'')||['nilai','nilai_ok','add_tambah','add_kurang','um','um_awal','ret_pct','tarif','top','adm','top_ret','limit'].indexOf(c[0])>=0;
      // NO & tanggal di tengah · nilai uang rata kanan · jumlah/persen/hari di tengah · tulisan rata kiri
      var al=c[0]==='no'||f==='date'||['lunas','tutup','ke'].indexOf(c[0])>=0?'center':/^(money|money0|num2)$/.test(f||'')||['nilai','nilai_ok','add_tambah','add_kurang','um','um_awal','adm','limit'].indexOf(c[0])>=0?'right':numeric?'center':'left';
      rg.setHorizontalAlignment(al);
      if(al==='left'&&(arLeftLabel_(c[1])||/catatan|ket$|alasan|uraian/.test(c[0])))rg.setWrap(true);
    });
  });
  ['Parameter','Data LK (Tempel)','Tempel SAP','Riwayat Revisi'].forEach(function(n){var sh=ss.getSheetByName(n);if(sh&&sh.getLastRow())arFormatGeneric_(sh);});
  try{arGroupRegister_(ss);}catch(e){}
  // judul kolom lama → judul baru (kolom tetap dikenali lewat ENGINE.HDR_LAMA)
  [ENGINE.REG,ENGINE.LOG].forEach(function(sc){try{var sh=ss.getSheetByName(sc.sheet);if(!sh)return;var lc=sh.getLastColumn(),h=sh.getRange(sc.headerRow,1,1,lc).getValues()[0];
    sc.cols.forEach(function(c){var old=(ENGINE.HDR_LAMA[sc.sheet+'|'+c[0]]||[]);h.forEach(function(v,j){if(old.indexOf(String(v).replace(/\s+/g,' ').trim().toUpperCase())>=0)sh.getRange(sc.headerRow,j+1).setValue(c[1]);});});}catch(e){}});
  // pilihan AKTIF Master Proyek selalu mengikuti daftar terbaru (termasuk ARSIP)
  try{var mp=ss.getSheetByName(ENGINE.MPRO.sheet),pm=mp?colMap_(mp,ENGINE.MPRO,ENGINE.MPRO_ALIAS):{};
    if(pm.aktif)mp.getRange(ENGINE.MPRO.firstRow,pm.aktif,Math.max(mp.getMaxRows()-ENGINE.MPRO.firstRow+1,1),1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(ENGINE.OPT.aktif_pro,true).setAllowInvalid(true)
      .setHelpText('YA = aktif · TIDAK = bukan proyek AR · SELESAI = selesai & lunas (tidak muncul di aplikasi) · ARSIP = selesai & lunas sebelum sistem dipakai: invoice, penerimaan & tagbrut-nya diabaikan dan barisnya disembunyikan (tidak dihapus).').build());}catch(e){}
}
/** sheet isian bebas (Parameter, Data LK, Tempel SAP, Riwayat Revisi): judul kiri · judul kolom tengah · tulisan kiri ·
 *  kode / tahun / bulan / persen / jumlah di tengah · nilai uang (format ribuan) rata kanan · tanggal tengah */
function arFormatGeneric_(sh) {
  var rg=sh.getDataRange(),v=rg.getValues(),nf=rg.getNumberFormats(),R=v.length,W=R?v[0].length:0;if(!R)return;
  var hr=-1,best=1;
  for(var i=0;i<Math.min(8,R);i++){var n=v[i].filter(function(x){return typeof x==='string'&&x.trim()&&x===x.toUpperCase()&&x.length<60;}).length;if(n>best){best=n;hr=i;}}
  var money=function(i,j){return /#,##0|#.##0/.test(nf[i][j])&&!/%/.test(nf[i][j]);},col={};
  for(var c=0;c<W;c++){var num=0,txt=0,mon=0,dt=0;
    for(var r=hr+1;r<R;r++){var x=v[r][c];if(x===''||x===null)continue;if(x instanceof Date)dt++;else if(typeof x==='number'){num++;if(money(r,c)&&Math.abs(x)>=1000)mon++;}else if(!/^[-–—]$/.test(String(x).trim()))txt++;}
    col[c]=txt>num+dt?'left':dt>=num&&dt?'center':mon&&mon*2>=num?'right':'center';}
  var al=v.map(function(row,r){return row.map(function(x,c){
    if(r<hr)return 'left';if(r===hr)return 'center';
    if(col[c]==='left')return 'left';
    if(typeof x==='string'&&x.trim()&&!/^[-–—]$/.test(x.trim()))return 'left';
    return col[c];});});
  rg.setVerticalAlignment('middle').setHorizontalAlignments(al);
  Object.keys(col).forEach(function(c){if(col[c]==='left'&&hr>=0&&R>hr+1)sh.getRange(hr+2,+c+1,R-hr-1,1).setWrap(true);});
}
function arOrderNamed_(ss,names) {
  var active=ss.getActiveSheet(),at=1;
  names.forEach(function(n){var sh=ss.getSheetByName(n);if(sh){ss.setActiveSheet(sh);ss.moveActiveSheet(at++);}});
  if(active)ss.setActiveSheet(active);
}
function arOrderSheets_(ss) {
  arOrderNamed_(ss,['Dashboard Summary','Kinerja & Peringatan Dini','Analisis Bulanan AR','Kartu Piutang','Aging Piutang','INKASO','RA-RI Termin','Pareto','Monitoring UM & Retensi','Monitoring Tagbrut','Rekap Invoice Keluar','Monitoring Invoice','Log Penerimaan','Input Tagbrut','Log Persetujuan','Riwayat Revisi','Memorial AR','Usulan Jurnal SAP','CKPN','Rekon SAP','Rekonsiliasi Saldo','Data LK (Tempel)','Tempel SAP','Master Pelanggan','Master Proyek','Riwayat Adendum','Referensi P&E','Parameter','Cek Data','Arsip Analisis AR']);
}
function arBlock_(r,rows,fmt) {return {r:r,c:1,rows:rows,fmt:fmt||{},styles:[{r:0,s:'hdr'}]};}
function arDashboardCharts_(res) {
  var out={asOf:res.ctx.today,aging:null,rari:null,inkaso:null},ag=res.reports['Aging Piutang'],rr=res.reports['RA-RI Termin'],ik=res.reports.INKASO;
  if(ag){var t=ag.blocks[0].rows.slice(-1)[0];out.aging={period:'Posisi '+res.ctx.today,rows:['Tanpa jatuh tempo','Belum jatuh tempo','Lewat 0–30 hari','Lewat 31–90 hari','Lewat 91–180 hari','Lewat >180 hari'].map(function(label,i){return [label,Number(t[5+i])||0];})};}
  if(rr){var sm=rr.blocks[1].rows;out.rari={period:sm[0][0],note:rr.blocks[0].rows[1][0],rows:[['RA awal',sm[1][1]],['RA final',sm[4][1]],['RI atas RA final',sm[5][1]],['Total RI',sm[8][1]]]};}
  if(ik){var tr=ik.blocks[3].rows;out.inkaso={period:tr[0][0],rows:tr.slice(2,14).map(function(r){return[r[0],r[1],r[2]];})};}
  return out;
}
function arEnrichReports_(res) {
  var ctx=res.ctx;
  function add(name,rows,title,type){var rep=res.reports[name];if(!rep)return;var bottom=rep.from;rep.blocks.forEach(function(b){bottom=Math.max(bottom,b.r+b.rows.length+2);});var bk=arBlock_(bottom,rows,{1:'num2',2:'num2'});bk.c=2;rep.blocks.push(bk);rep.charts=(rep.charts||[]).concat([{title:title+' (Rp juta)',type:type,range:[bottom,2,rows.length,rows[0].length],row:bottom+rows.length+1,col:2,w:760,h:300,colors:['#137C8B','#1F3864','#D89532'],legend:rows[0].length>2?'bottom':'none'}]);}
  var data=arDashboardCharts_(res);
  function juta(rows){return rows.map(function(r){return r.map(function(v,i){return i?v/1e6:v;});});}
  if(data.aging)add('Aging Piutang',[['Kelompok umur','Saldo kas']].concat(juta(data.aging.rows)),'AR CSU · Aging piutang','BAR');
  if(data.inkaso)add('INKASO',[['Bulan','Target','Realisasi']].concat(juta(data.inkaso.rows)),'AR CSU · Target dan realisasi INKASO','COLUMN');
  if(data.rari)add('RA-RI Termin',[['Komponen','Nilai']].concat(juta(data.rari.rows)),'AR CSU · Rencana dan realisasi penerimaan','COLUMN');
}
function arWriteReconciliation_(res) {
  var sh=sheet_('Rekonsiliasi Saldo');if(!sh)return;
  sh.getRange('A1').setValue('REKONSILIASI SALDO — PEMERIKSAAN TANPA MENGUBAH TRANSAKSI');
  sh.getRange('A2').setValue('Cocokkan invoice, kas, PPh diakui, bukti potong dan LK/SAP. Selisih basis kas vs buku pembantu tidak otomatis berarti salah.');
  sh.getRange('A3').setValue('Posisi '+res.today+' · '+arNow_()+' · tidak menolkan saldo berdasarkan nama pelanggan.');
  var hdr=['PELANGGAN','NPP','INVOICE','STATUS SISTEM','TAGIHAN DICATAT','KAS DITERIMA','PPh PERHITUNGAN','PPh DIAKUI','SALDO BUKU PEMBANTU','SISA KAS','BEDA BASIS','TINDAK LANJUT','BUKTI POTONG','NO DOK SAP'];
  var rec=[hdr].concat(ENGINE.reconciliationRows(res.ctx)),bp=ENGINE.bupotPendingRows(res.ctx),r2=5+rec.length+3;
  var bpRows=[['PPh SUDAH DIAKUI — BUKTI POTONG BELUM DICATAT ('+bp.length+' penerimaan). Isi kolom BUKTI POTONG di Log Penerimaan setelah bupot diterima; tanpa bupot PPh tidak bisa dikreditkan.'],['PELANGGAN','NPP','INVOICE','TGL BAYAR','REF','JENIS PPh','PPh DIAKUI','UMUR (HARI)','ID PENERIMAAN']].concat(bp.length?bp:[['Tidak ada — semua PPh yang diakui sudah memiliki nomor bukti potong.']]);
  var b2=arBlock_(r2,bpRows,{6:'money0',7:'int'});b2.styles=[{r:0,s:'title'},{r:1,s:'hdr'}];
  writeReport_(sh,{from:5,blocks:[arBlock_(5,rec,{4:'money0',5:'money0',6:'money0',7:'money0',8:'money0',9:'money0',10:'money0'}),b2],wrapCols:[12,13],minW:{1:28,3:28,12:48}},'Rekonsiliasi Saldo');
  sh.setFrozenRows(5); sh.setHiddenGridlines(true);
}

// ---------------- P&E: existing per-project tabs retained, protected date selectors.
function arPECutoff_(sh,fallback,today) {
  var mode=String(sh.getRange('E4').getValue()||'IKUT RINGKASAN'),raw=sh.getRange('C4').getValue(),date=isDate_(raw)?dstr_(raw,sh.getParent().getSpreadsheetTimeZone()):String(raw||'');
  if(['HARI INI','IKUT RINGKASAN','PILIH TANGGAL','AKHIR BULAN'].indexOf(mode)<0)throw new Error('Pilihan cutoff tidak dikenali.');
  if(mode==='HARI INI')return today;
  if(mode==='IKUT RINGKASAN')return fallback||today;
  if(!arValidDate_(date))throw new Error('Tanggal cutoff '+sh.getName()+' tidak valid. Isi C4 sebagai tanggal.');
  if(mode==='AKHIR BULAN'){var p=date.split('-');date=new Date(Date.UTC(+p[0],+p[1],0)).toISOString().slice(0,10);}
  if(date>today)throw new Error('Cutoff belum terjadi. Pilih tanggal sampai '+today+' atau mode HARI INI.');
  return date;
}
/** baris status A3 file P&E: merah bila posisi bukan hari ini */
function arPEStatus_(sh,text,alert){var r=sh.getRange('A3:H3');r.setBackground(alert?'#FDECEA':null).setFontColor(alert?'#B42318':'#4A5A72').setFontWeight(alert?'bold':'normal');sh.getRange('A3').setValue(text);}
/** pagi hari: semua tab P&E kembali ke HARI INI / IKUT RINGKASAN (tanggal pilihan manual hanya sementara) */
function arPEResetCutoff_() {
  var id=paramKV_('FILE_PE_ID');if(!id||!id.val)return 0;
  var ss=SpreadsheetApp.openById(id.val),today=Utilities.formatDate(new Date(),TZ,'yyyy-MM-dd'),n=0;
  ss.getSheets().forEach(function(sh){var nm=sh.getName();if(nm!=='RINGKASAN'&&!/^\d{3}C \d{4} [A-Z]{2}/.test(nm))return;
    var m=String(sh.getRange('E4').getValue()||''),want=nm==='RINGKASAN'?'HARI INI':'IKUT RINGKASAN';
    if(m===want&&!/^⚠/.test(String(sh.getRange('A3').getValue())))return;
    sh.getRange('E4').setValue(want);sh.getRange('C4').setValue(toCell_(today,'date'));
    arPEStatus_(sh,'MENUNGGU PEMBARUAN — posisi dikembalikan ke HARI INI ('+today+'); isi kartu diperbarui pada sinkron berikutnya (paling lambat 1 jam).',false);n++;});
  return n;
}
function arPESelectors_(sh,isSummary,today) {
  if(sh.getMaxColumns()<24)sh.insertColumnsAfter(sh.getMaxColumns(),24-sh.getMaxColumns());
  var c=sh.getRange('C4'),mode=sh.getRange('E4');
  if(!c.getValue())c.setValue(toCell_(today,'date'));
  if(!mode.getValue())mode.setValue(isSummary?'HARI INI':'IKUT RINGKASAN');
  sh.getRange('B4').setValue('Tanggal posisi');
  c.setNumberFormat(NUMF.date).setBackground('#FFF2CC').setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build());
  sh.getRange('D4').setValue('Pilihan cutoff');
  var opts=['HARI INI','PILIH TANGGAL','AKHIR BULAN'];if(!isSummary)opts.unshift('IKUT RINGKASAN');
  mode.setBackground('#FFF2CC').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(opts,true).setAllowInvalid(false).build());
  sh.getRange('A4:G4').setVerticalAlignment('middle').setHorizontalAlignment('center');
  var owner=sh.getParent().getOwner(),editors=[];if(owner)editors.push(owner.getEmail());
  var effective=Session.getEffectiveUser().getEmail();if(effective)editors.push(effective);
  var desc='AR CSU v2.8 — pilihan cutoff P&E',old=sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).filter(function(p){return p.getDescription()===desc;});
  var pr=old[0]||sh.protect().setDescription(desc);
  pr.setUnprotectedRanges([c,mode]);pr.removeEditors(pr.getEditors());if(editors.length)pr.addEditors(editors);if(pr.canDomainEdit())pr.setDomainEdit(false);
}
function arPushPE_(res,npps,onlyName) {
  var ss=derivedFile_('FILE_PE_ID','KARTU PIUTANG CSU — P&E'),today=res.ctx.today,stamp=arNow_(),rs=sheetIn_(ss,'RINGKASAN',0),cache={};
  if(ss.getSpreadsheetTimeZone()!==TZ)ss.setSpreadsheetTimeZone(TZ);
  arPESelectors_(rs,true,today);
  var cutoff=arPECutoff_(rs,today,today);
  var at=function(d){return cache[d]||(cache[d]=d===today?res:ENGINE.computeAt(res.ctx._input||readAll_(),d));};
  if(!onlyName||onlyName==='RINGKASAN'){
    var sr=at(cutoff),sum=ENGINE.peSummaryBlocks(sr.ctx,6);
    rs.getRange('A1').setValue('KARTU PIUTANG PROYEK — CONCRETE SERVICES UNIT').setFontSize(14).setFontWeight('bold').setFontColor(C.navy);
    rs.getRange('A2').setValue('P&E · ubah sel kuning untuk memilih cutoff. Pilihan pada file ini digunakan bersama.');
    arPEStatus_(rs,cutoff<today?'⚠ POSISI PER '+cutoff+' — BUKAN HARI INI ('+today+'). Pilih HARI INI di E4 untuk posisi terbaru; setiap pagi pilihan tanggal kembali ke HARI INI.':'Posisi '+cutoff+' (hari ini) · diperbarui '+stamp+' WIB · berdasarkan data sumber terkini; bukan arsip master historis.',cutoff<today);
    writeReport_(rs,sum,'RINGKASAN',{key:'PE28_'+ss.getId()+'_RINGKASAN'});rs.setFrozenRows(7); rs.setHiddenGridlines(true);
  }
  var all=ss.getSheets(),list=ENGINE.proyekAR(res.ctx).sort(),names=['RINGKASAN'],count=0,start=Date.now();
  list.forEach(function(n,ix){
    var p=res.ctx.M.pro[n];if(!p)return;
    var name=(n+' '+p.singkat).replace(/[\[\]*?\/\\:']/g,'-').slice(0,90),sh=all.filter(function(s){return s.getName().indexOf(n)===0;})[0];
    if(!sh){sh=ss.insertSheet(name,Math.min(ix+1,ss.getSheets().length));all.push(sh);}names.push(sh.getName());
    if(onlyName&&onlyName!=='RINGKASAN'&&sh.getName()!==onlyName)return;
    // A global date change only refreshes tabs that follow the summary date.
    if(onlyName==='RINGKASAN'&&sh.getRange('E4').getValue()&&sh.getRange('E4').getValue()!=='IKUT RINGKASAN')return;
    if(npps&&npps.indexOf(n)<0&&sh.getLastRow()>6)return;
    arPESelectors_(sh,false,today);
    if(Date.now()-start>200000){sh.getRange('A3').setValue('MENUNGGU PEMBARUAN · belum selesai '+stamp+' · jadwal berikutnya akan melanjutkan.');return;}
    var date;
    try {date=arPECutoff_(sh,cutoff,today);}catch(e){sh.getRange('A3').setValue('PILIHAN TANGGAL TIDAK VALID — '+e.message);return;}
    var hctx=at(date).ctx,kb=ENGINE.kartuBlocks(hctx,n,6,date),key='PE28_'+ss.getId()+'_'+n,digest=ENGINE.hash(JSON.stringify(kb.blocks.map(function(b){return b.rows;}))+'|'+date+'|'+ENGINE.VERSION+'|FORMAT-20261013-04');
    var wr=ENGINE.posisiWarning(res.ctx,n,date,'pilih HARI INI atau IKUT RINGKASAN di E4 untuk posisi terbaru; setiap pagi pilihan tanggal kembali otomatis');
    if(arProps_().getProperty(key)===digest){arPEStatus_(sh,wr||'Posisi '+date+' (hari ini) · data diperiksa '+stamp+' WIB · berdasarkan data sumber terkini.',!!wr);return;}
    sh.getRange('A1').setValue('KARTU PIUTANG PROYEK · '+n).setFontSize(13).setFontWeight('bold').setFontColor(C.navy);
    sh.getRange('A2').setValue(p.pelanggan+' · '+(p.nama_pe||p.singkat)+(p.selesai?' · PROYEK SELESAI':''));
    arPEStatus_(sh,wr||'Posisi '+date+' (hari ini) · diperbarui '+stamp+' WIB · berdasarkan data sumber terkini.',!!wr);
    writeReport_(sh,kb,'PE_'+n,{key:key});sh.setHiddenGridlines(true); sh.setFrozenRows(4);
    arProps_().setProperty(key,digest);count++;
  });
  // Never delete old project tabs: closed projects remain available for historical review.
  ss.getSheets().forEach(function(sheet){if(names.indexOf(sheet.getName())>=0){sheet.setColumnWidth(5,Math.max(190,sheet.getColumnWidth(5)));sheet.getRange('E4').setWrap(true);}});
  // proyek ARSIP (selesai sebelum sistem): tab kartunya disembunyikan, tidak dihapus
  var arsipN=Object.keys(res.ctx.M.arsip||{});
  if(arsipN.length)ss.getSheets().forEach(function(sheet){var nm=sheet.getName();if(arsipN.some(function(n){return nm.indexOf(n)===0;})&&!sheet.isSheetHidden()){try{sheet.hideSheet();}catch(e){}}});
  arOrderNamed_(ss,names);dropDefaultSheet_(ss);
  return count;
}
function arEnablePE() {
  arRequireOwner_();var lock=arLock_();lock.waitLock(28000);try {
  var res=compute_(),ss=derivedFile_('FILE_PE_ID','KARTU PIUTANG CSU — P&E');
  arPushPE_(res);
  var emails=res.params.users.filter(function(u){return /^P&E/.test(u.peran)&&u.email;}).map(function(u){return u.email;}),file=DriveApp.getFileById(ss.getId());
  file.setShareableByEditors(false);
  // Deliberate promotion: only protected C4/E4 are editable in generated tabs.
  if(emails.length)file.addEditors(emails);
  ScriptApp.getProjectTriggers().forEach(function(t){if(t.getHandlerFunction()==='onPEEditV28')ScriptApp.deleteTrigger(t);});
  var trigger=ScriptApp.newTrigger('onPEEditV28').forSpreadsheet(ss).onEdit().create();
  PropertiesService.getScriptProperties().setProperty('AR_PE_TRIGGER_ID',trigger.getUniqueId());
  return {ok:true,url:ss.getUrl(),editors:emails};
  }finally{lock.releaseLock();}
}
function onPEEditV28(e) {
  if(!e||!e.range||!e.source||!e.triggerUid||String(e.triggerUid)!==PropertiesService.getScriptProperties().getProperty('AR_PE_TRIGGER_ID'))return;
  var id=paramKV_('FILE_PE_ID');if(!id||e.source.getId()!==id.val)return;
  var r=e.range,sh=r.getSheet(),name=sh.getName();
  if(name!=='RINGKASAN'&&!/^\d{3}C \d{4} [A-Z]{2}/.test(name))return;
  if(r.getRow()>4||r.getLastRow()<4||r.getColumn()>5||r.getLastColumn()<3)return;
  var lock=arLock_();if(!lock.tryLock(25000)){sh.getRange('A3').setValue('MENUNGGU PEMBARUAN — sistem sibuk; akan dicoba pada sinkron berikutnya.');return;}
  try {
    if(r.getColumn()<=3&&r.getLastColumn()>=3)sh.getRange('E4').setValue('PILIH TANGGAL');
    sh.getRange('A3').setValue('MEMPERBARUI — pilihan cutoff baru sedang dihitung… (dimulai '+arNow_()+'; bila belum berubah, sinkron otomatis per jam melanjutkan)');
    arPushPE_(compute_(),null,name);
  }catch(err){sh.getRange('A3').setValue('BELUM DIPERBARUI — '+err.message);}finally{lock.releaseLock();}
}

// ---------------- Monthly snapshots: immutable revisions; AI is opt-in and aggregate-only.
function arMonthEnd_(month) {if(!/^\d{4}-(0[1-9]|1[0-2])$/.test(month))throw new Error('Month must be yyyy-mm.');var p=month.split('-');return new Date(Date.UTC(+p[0],+p[1],0)).toISOString().slice(0,10);}
function arMonthlyArchive_() {
  var sh=sheet_('Arsip Analisis AR');if(!sh)throw new Error('Install the v2.8 update first.');return sh;
}
function arMonthlyRows_() {var sh=arMonthlyArchive_();return sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,12).getValues():[];}
function arMonthlyNarrative_(m) {
  var pct=m.raFinal?Math.round((m.totalRI||m.riAtasRA)/m.raFinal*10000)/100:null,lines=[];
  lines.push('Posisi '+m.cutoff+': saldo buku pembantu Rp '+ENGINE.fmtRp(m.saldo)+', berubah Rp '+ENGINE.fmtRp(m.perubahan)+' dari akhir bulan sebelumnya.');
  lines.push('Penerimaan kas bulan ini Rp '+ENGINE.fmtRp(m.penerimaan)+'. RA final (DPP − PPh) Rp '+ENGINE.fmtRp(m.raFinal)+'; total RI Rp '+ENGINE.fmtRp(m.totalRI||0)+' (RI atas RA Rp '+ENGINE.fmtRp(m.riAtasRA)+')'+(pct===null?'; pencapaian belum dapat dihitung karena RA nol.':'; pencapaian (total RI ÷ RA final) '+pct+'%.'));
  lines.push('Sisa kas lewat jatuh tempo Rp '+ENGINE.fmtRp(m.overdue)+'. Selisih basis saldo buku pembantu dan kas harus dibaca bersama potongan PPh/adm.');
  if(m.top.length)lines.push('Prioritas menurut saldo buku pembantu: '+m.top.map(function(x){return x.pelanggan+' (Rp '+ENGINE.fmtRp(x.saldo)+')';}).join('; ')+'.');
  if(m.bedaStatus)lines.push(m.bedaStatus+' invoice berstatus selesai namun kartu bersaldo: cocokkan rincian pembayaran dan potongan sebelum penagihan.');
  lines.push('Tindakan '+m.nextMonth+': Minggu 1 — AR Officer mengonfirmasi jadwal bayar dan bukti penerimaan/PPh; AR bersama Akuntansi mencocokkan selisih dengan LK/SAP. Minggu 2 — Section Head mengevaluasi rencana dan realisasi. Alasan keterlambatan perlu dikonfirmasi kepada PIC; tidak disimpulkan dari angka saja.');
  return lines.join('\n\n');
}
function arAINarrative_(metrics) {
  var p=PropertiesService.getScriptProperties();
  if(p.getProperty('AR_AI_ENABLED')!=='YA')return {status:'NONAKTIF — analisis berbasis aturan',text:''};
  var key=p.getProperty('AR_AI_API_KEY'),model=p.getProperty('AR_AI_MODEL');
  if(!key||!model)return {status:'BELUM DIKONFIGURASI — API key / model kosong',text:''};
  // Only aggregate metrics leave the spreadsheet; no emails, invoices, bank accounts, customer names or free text.
  var safe={cutoff:metrics.cutoff,nextMonth:metrics.nextMonth,opening:metrics.opening,saldo:metrics.saldo,perubahan:metrics.perubahan,kasTerbuka:metrics.kasTerbuka,overdue:metrics.overdue,penerimaan:metrics.penerimaan,targetInkaso:metrics.targetInkaso,raAwal:metrics.raAwal,raFinal:metrics.raFinal,riAtasRA:metrics.riAtasRA,totalRI:metrics.totalRI,aging:metrics.aging,bedaStatus:metrics.bedaStatus};
  try {
    var response=UrlFetchApp.fetch('https://api.openai.com/v1/responses',{method:'post',contentType:'application/json',headers:{Authorization:'Bearer '+key},muteHttpExceptions:true,payload:JSON.stringify({model:model,store:false,max_output_tokens:1600,instructions:'Anda membantu analisis AR. Tulis bahasa Indonesia maksimal 350 kata, hanya berdasarkan JSON. Jangan menebak penyebab, menyatakan telah melakukan tindakan, atau menyarankan penghapusan saldo. Bedakan saldo buku pembantu dan sisa kas. Nyatakan keterbatasan dan usulkan tindak lanjut bulan berikutnya. Jika pembagi nol jangan hitung persentase. Narasi perlu ditinjau pengguna.',input:JSON.stringify(safe)})});
    if(response.getResponseCode()!==200)return {status:'GAGAL HTTP '+response.getResponseCode()+' — analisis aturan tetap tersedia',text:''};
    var data=JSON.parse(response.getContentText()),text=[];
    (data.output||[]).forEach(function(x){(x.content||[]).forEach(function(c){if(c.type==='output_text'&&c.text)text.push(c.text);});});
    if(data.status!=='completed'||!text.length)return {status:'RESPONS BELUM LENGKAP — analisis aturan tetap tersedia',text:''};
    return {status:'AI — PERLU DITINJAU',text:text.join('\n').slice(0,20000)};
  }catch(e){return {status:'GAGAL KONEKSI AI — analisis aturan tetap tersedia',text:''};}
}
function arCreateMonthly_(month,revise,reason) {
  var cutoff=arMonthEnd_(month);if(cutoff>=arToday_())throw new Error('A snapshot can only be created after the month ends.');
  var lock=arLock_();lock.waitLock(28000);
  try {
    var rows=arMonthlyRows_(),old=rows.filter(function(r){return r[1]===month;});
    if(old.length&&!revise){arRenderMonthly_(old[old.length-1]);return {created:false,id:old[old.length-1][0]};}
    if(revise&&String(reason||'').trim().length<10)throw new Error('The revision reason must be at least 10 characters.');
    var input=readAll_(),m=ENGINE.monthlyMetrics(input,cutoff),n=arMonthlyNarrative_(m),ai=arAINarrative_(m),rev=old.length,id=month+'-Rev.'+('0'+rev).slice(-2),digest=arDigest_(JSON.stringify(m));
    var row=[id,month,rev,cutoff,arNow_(),email_()||Session.getEffectiveUser().getEmail(),digest,reason||'Snapshot otomatis akhir bulan',JSON.stringify(m),n,ai.text,ai.status],sh=arMonthlyArchive_();
    roomFor_(sh,sh.getLastRow()+1);sh.getRange(sh.getLastRow()+1,1,1,12).setNumberFormat('@').setValues([row.map(arSafeCell_)]);
    arRenderMonthly_(row);
    return {created:true,id:id,ai:ai.status};
  }finally{lock.releaseLock();}
}
function arRenderMonthly_(r) {
  var sh=sheet_('Analisis Bulanan AR');if(!sh||!r)return;var m=JSON.parse(r[8]),rows=[['UKURAN','NILAI'],['Saldo buku pembantu awal',m.opening],['Saldo buku pembantu akhir',m.saldo],['Perubahan saldo',m.perubahan],['Sisa kas terbuka',m.kasTerbuka],['Kas lewat jatuh tempo',m.overdue],['Penerimaan bulan ini',m.penerimaan],['Target INKASO',m.targetInkaso],['RA awal',m.raAwal],['RA final',m.raFinal],['RI atas RA final',m.riAtasRA],['Total RI',m.totalRI]];
  sh.getRange('A1').setValue('ANALISIS BULANAN AR — '+r[0]).setFontSize(14).setFontWeight('bold').setFontColor(C.navy);
  sh.getRange('A2').setValue('Cutoff '+m.cutoff+' · tindak lanjut '+m.nextMonth+' · dibuat '+r[4]);
  sh.getRange('A3').setValue('Snapshot tersimpan; revisi baru tidak menimpa arsip sebelumnya.');
  sh.getRange('B5').setValue('Pilih snapshot');sh.getRange('C5').setValue(r[0]).setBackground('#FFF2CC').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(arMonthlyRows_().map(function(x){return x[0];}),true).setAllowInvalid(false).build());
  var nar=String(r[9]).split('\n\n').map(function(s){return[s];}),blocks=[arBlock_(7,rows,{1:'money0'}),arBlock_(21,[['ANALISIS OTOMATIS BERBASIS ATURAN']].concat(nar),{})];
  var ar=24+nar.length;blocks.push(arBlock_(ar,[['NARASI AI — '+r[11]],[r[10]||'Belum diaktifkan. Analisis angka dan tindak lanjut di atas tetap dibuat otomatis.'],[m.basis]],{}));
  writeReport_(sh,{from:7,blocks:blocks,widths:{1:85,2:25,3:25},wrapCols:[1],minW:{}},'Analisis Bulanan AR');
  blocks.slice(1).forEach(function(b){sh.getRange(b.r,1,b.rows.length,3).mergeAcross().setWrap(true).setHorizontalAlignment('left').setVerticalAlignment('middle');sh.autoResizeRows(b.r,b.rows.length);});
  sh.setHiddenGridlines(true); sh.setFrozenRows(5);
}
// Snapshot otomatis dibuat setelah bulan itu tutup buku (tanggal 3 bulan berikutnya, jadwal harian 06.00).
function arMonthlyDue_(month) { var st=arPeriodState_(); return !!st.tutup && st.tutup>=arMonthEnd_(month); }
/** refresh setelah koreksi berizin di periode tertutup: snapshot bulan terdampak yang sudah ada direvisi (arsip lama tetap) */
function arRefreshClosedMonths_(dates, alasan) {
  var st=arPeriodState_(),done=[],warn=[],seen={};
  (dates||[]).forEach(function(d){
    d=String(d||'').slice(0,10); if(!arValidDate_(d)||!st.tutup||d>st.tutup)return; var m=d.slice(0,7); if(seen[m])return; seen[m]=1;
    if(!arMonthlyRows_().some(function(r){return r[1]===m;}))return;
    try{var r=arCreateMonthly_(m,true,'Koreksi periode tertutup — '+String(alasan||'').slice(0,300));done.push(r.id);}catch(e){warn.push('Snapshot '+m+' not revised: '+e.message);}
  });
  return {ids:done,warnings:warn};
}
function arMonthlyAuto_(){if(!sheet_('Arsip Analisis AR'))return;var month=ENGINE.prevMonthEnd(arToday_()).slice(0,7);if(!arMonthlyDue_(month))return;try{arCreateMonthly_(month,false,'');arProps_().setProperty('MONTHLY_ERROR','');}catch(e){arProps_().setProperty('MONTHLY_ERROR',arNow_()+' '+e.message);var sh=sheet_('Analisis Bulanan AR');if(sh)sh.getRange('A3').setValue('Analisis otomatis belum berhasil: '+e.message);}}
function arMonthlyMenu(){arRequireMaker_();var ui=SpreadsheetApp.getUi(),mo=ENGINE.prevMonthEnd(arToday_()).slice(0,7);if(!arPeriodState_().tutup||arPeriodState_().tutup<arMonthEnd_(mo)){if(ui.alert('Periode belum ditutup','Bulan '+mo+' belum ditutup (PERIODE_DITUTUP_SAMPAI). Snapshot sekarang bisa belum memuat penerimaan akhir bulan yang terlambat diinput. Tetap buat?',ui.ButtonSet.OK_CANCEL)!==ui.Button.OK)return;}var r=arCreateMonthly_(ENGINE.prevMonthEnd(arToday_()).slice(0,7),false,'');SpreadsheetApp.getUi().alert('Analisis tersedia: '+r.id);}
function arMonthlyRevisionMenu(){arRequireMaker_();var ui=SpreadsheetApp.getUi(),m=ui.prompt('Revisi snapshot','Bulan yyyy-mm (bulan yang sudah selesai)',ui.ButtonSet.OK_CANCEL);if(m.getSelectedButton()!==ui.Button.OK)return;var a=ui.prompt('Alasan revisi','Contoh: koreksi bukti potong sudah diverifikasi Akuntansi.',ui.ButtonSet.OK_CANCEL);if(a.getSelectedButton()!==ui.Button.OK)return;var r=arCreateMonthly_(m.getResponseText().trim(),true,a.getResponseText());ui.alert('Tersimpan '+r.id);}

// ---------------- Installation: backup succeeds before schema / data formatting changes.
function applyV28() {
  var active=SpreadsheetApp.getActive();if(!active||!active.getSheetByName('Monitoring Invoice'))throw new Error('Jalankan dari spreadsheet utama AR.');
  AR28_SS_=active;arRequireOwner_();
  var lock=arLock_();lock.waitLock(28000);
  try {
    var dp=PropertiesService.getScriptProperties(),old=dp.getProperty('AR_SOURCE_ID');
    if(old&&old!==active.getId())throw new Error('Proyek script terikat ke file sumber lain. Periksa AR_SOURCE_ID sebelum instalasi.');
    var stamp=arNow_().replace(/[: ]/g,'-'),backup=[];
    [active.getId()].concat(['FILE_PE_ID','FILE_AKUNTANSI_ID'].map(function(k){var x=paramKV_(k);return x&&x.val;}).filter(Boolean)).forEach(function(id){var f=DriveApp.getFileById(id),b=f.makeCopy(f.getName()+' — BACKUP sebelum v2.8 '+stamp);backup.push({source:id,backup:b.getId(),url:b.getUrl()});});
    dp.setProperty('AR28_BACKUP_'+stamp,JSON.stringify(backup));
    dp.setProperty('AR_SOURCE_ID',active.getId());
    var legacy=PropertiesService.getDocumentProperties();if(legacy){var lp=legacy.getProperties();Object.keys(lp).forEach(function(k){if(dp.getProperty('AR28_'+k)===null)dp.setProperty('AR28_'+k,lp[k]);});}
    ensureSchema_(sheet_('Log Penerimaan'),ENGINE.LOG,null,4);
    ['Rekonsiliasi Saldo','Analisis Bulanan AR'].forEach(function(n){var sh=sheetIn_(active,n);sh.setTabColor('#137C8B');});
    var arch=sheetIn_(active,'Arsip Analisis AR');
    if(!arch.getLastRow())arch.getRange(1,1,1,12).setValues([['ID SNAPSHOT','BULAN','REVISI','CUTOFF','DIBUAT','OLEH','SIDIK','ALASAN REVISI','METRIK JSON','ANALISIS ATURAN','NARASI AI','STATUS AI']]);
    var pr=arch.getProtections(SpreadsheetApp.ProtectionType.SHEET).filter(function(p){return p.getDescription()==='AR CSU v2.8 — arsip analisis';})[0]||arch.protect().setDescription('AR CSU v2.8 — arsip analisis');pr.removeEditors(pr.getEditors());pr.addEditors([email_()].concat(arMakerEmails_()).filter(Boolean));if(pr.canDomainEdit())pr.setDomainEdit(false);
    arOrderSheets_(active);arFormatSourceSheets_(active);
    arInstallTriggers_(active);
    secret_();rebuildAll();
    arProps_().setProperty('V28_DONE','1');
    // Derived files and the first monthly snapshot run separately to avoid the execution limit.
    SpreadsheetApp.getUi().alert('Pembaruan v2.8 terpasang.\nBackup:\n'+backup.map(function(b){return b.url;}).join('\n')+'\n\nLanjutkan menu Aktifkan pilihan tanggal P&E, lalu Perbarui file Akuntansi & P&E, dan Buat / lihat analisis bulan lalu (satu per satu). Deploy versi web baru. Saldo pelanggan tidak dinolkan otomatis; cocokkan rincian PPh pada Rekonsiliasi Saldo.');
  }finally{lock.releaseLock();}
}
function apiResolvePayment(requestId) {
  var a=arRequireMaker_();if(!/^[A-Za-z0-9_-]{12,100}$/.test(requestId||''))throw new Error('Invalid ID.');
  var l=arLock_();l.waitLock(28000);
  try {
    var rows=readSchema_(sheet_('Log Penerimaan'),ENGINE.LOG),ri=ENGINE.colIndex(ENGINE.LOG,'request_id'),ui=ENGINE.colIndex(ENGINE.LOG,'request_user'),found=rows.filter(function(r){return r[ri]===requestId;});
    if(found.some(function(r){return r[ui]!==a.email;}))throw new Error('The request belongs to another user.');
    if(found.length)return {saved:true,count:found.length};
    // Cancel outstanding writes with this token under the same lock; a delayed RPC cannot commit later.
    arProps_().setProperty('VOID_'+requestId,a.email+'|'+arToday_());
    return {saved:false,canceled:true};
  }finally{l.releaseLock();}
}

// ---------------- Trigger: tanpa identitas pengedit; pembaruan berat ditandai lalu dikerjakan satu per satu.
var AR_DIRTY_KEY_ = 'AR_REBUILD_DIRTY';
/** benar hanya bila dipanggil oleh trigger terpasang proyek ini (event google.script.run tidak memiliki triggerUid yang terdaftar) */
function arFromTrigger_(e) {
  var uid = e && e.triggerUid ? String(e.triggerUid) : '';
  if (!uid) return false;
  try { return ScriptApp.getProjectTriggers().some(function(t){return String(t.getUniqueId())===uid;}); } catch (x) { return false; }
}
function arInstallTriggers_(ss) {
  var managed=['onEditInstalled','dailyRebuild','syncTurunanAuto','rebuildIfDirty','onChangeInstalled'];
  ScriptApp.getProjectTriggers().forEach(function(t){if(managed.indexOf(t.getHandlerFunction())>=0)ScriptApp.deleteTrigger(t);});
  ScriptApp.newTrigger('onEditInstalled').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('dailyRebuild').timeBased().everyDays(1).atHour(6).inTimezone(TZ).create();
  ScriptApp.newTrigger('syncTurunanAuto').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('rebuildIfDirty').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('onChangeInstalled').forSpreadsheet(ss).onChange().create();
}
function arMarkDirty_() { PropertiesService.getScriptProperties().setProperty(AR_DIRTY_KEY_, Date.now() + '-' + Math.random().toString(36).slice(2, 8)); arBootBust_(); }
/** kerjakan bila ditandai; bila lock dipegang proses lain, tanda dibiarkan untuk jadwal 5 menit (tidak ada "Lock timeout" bagi pengedit) */
function arRebuildIfDirty_(waitMs) {
  var sp = PropertiesService.getScriptProperties();
  if (!sp.getProperty(AR_DIRTY_KEY_)) return false;
  var lock = arLock_(); if (!lock.tryLock(waitMs)) return false;
  try {
    var t = sp.getProperty(AR_DIRTY_KEY_); if (!t) return false;
    rebuildAll_({ format: false });
    if (sp.getProperty(AR_DIRTY_KEY_) === t) sp.deleteProperty(AR_DIRTY_KEY_); // edit selama pembaruan → tanda baru tetap ada
    return true;
  } finally { lock.releaseLock(); }
}
function arEditRebuild_() { arMarkDirty_(); arRebuildIfDirty_(2000); }

// ---------------- Notifikasi email: bukti keputusan di luar spreadsheet & proyek Apps Script.
function arNotify_(subject, lines) {
  var P = arParameters_(), to = P.users.filter(function(u){return /^PENYETUJU/.test(u.peran)&&u.email;}).map(function(u){return u.email;});
  String(P.NOTIF_PERSETUJUAN_CC || '').split(/[,;\s]+/).forEach(function(x){ x=x.trim().toLowerCase(); if(/^[^@\s]+@[^@\s]+$/.test(x)) to.push(x); });
  to = to.filter(function(v,i,a){return a.indexOf(v)===i;});
  if (!to.length) return 'Email notification not sent: no PENYETUJU email is set in Parameter.';
  try {
    MailApp.sendEmail({ to: to.join(','), subject: '[AR CSU] ' + subject,
      body: lines.join('\n') + '\n\n— Email otomatis Sistem AR CSU (' + arNow_() + ' WIB). Simpan email ini: ini bukti di luar sistem. Bila Log Persetujuan berbeda dengan email ini, yang berlaku adalah email ini dan perbedaan harus diperiksa.' });
    return '';
  } catch (e) { return 'Email notification failed: ' + e.message; }
}
function arNotifyDecision_(n) {
  if (!n || !n.r) return '';
  var r = n.r;
  return arNotify_(r.status + ' — ' + (ENGINE.labelTampil(r.label) || ENGINE.jenisTampil(r.jenis) + ' ' + r.ref), [
    'Keputusan  : ' + r.status + (n.dl ? ' (DELEGASI a.n. ' + n.dl.atas + ')' : ''),
    'ID         : ' + r.id, 'Jenis      : ' + ENGINE.jenisTampil(r.jenis), 'Referensi  : ' + r.ref, 'Label      : ' + (ENGINE.labelTampil(r.label) || '-'),
    'Nilai      : Rp ' + ENGINE.fmtRp(Number(r.nilai) || 0),
    'Diajukan   : ' + r.oleh + ' · ' + r.waktu,
    'Diputuskan : ' + (n.nama || n.me) + ' <' + n.me + '> · ' + r.putus_waktu + ' WIB',
    'Sidik data : ' + r.sidik, 'Kode verif.: ' + r.kode, 'Catatan    : ' + (r.catatan || '-')]);
}

// ---------------- Tutup periode: transaksi bertanggal ≤ PERIODE_DITUTUP_SAMPAI ditolak,
// kecuali Section Head membuka koreksi sementara (aplikasi) DAN pembuat menulis alasan. Setiap koreksi dikirim ke email Section Head.
function arPeriodState_(P) {
  P = P || arParameters_();
  var t = ENGINE.periodeTutup(P, arToday_()), k = String(P.PERIODE_KOREKSI_SAMPAI || '');
  return { tutup: arValidDate_(t) ? t : '', koreksi: arValidDate_(k) && k >= arToday_() ? k : '', oleh: String(P.PERIODE_KOREKSI_OLEH || ''), alasan: String(P.PERIODE_KOREKSI_ALASAN || '') };
}
/** '' bila semua tanggal di periode terbuka; tanda koreksi bila diizinkan; selain itu error berkode PERIODE_TERTUTUP / PERIODE_TERTUTUP_ALASAN */
function arCheckPeriod_(dates, alasan, what) {
  var st = arPeriodState_(), hit = (dates || []).filter(function(d){return st.tutup && d && String(d) <= st.tutup;}).sort();
  if (!hit.length) return '';
  if (!st.koreksi) throw new Error('PERIODE_TERTUTUP: ' + what + ' dated ' + hit[0] + ' falls in a closed period (through ' + st.tutup + '). Ask the Section Head to open a temporary correction window in the app (Approvals).');
  if (String(alasan || '').trim().length < 10) throw new Error('PERIODE_TERTUTUP_ALASAN: ' + what + ' dated ' + hit[0] + ' falls in a closed period (through ' + st.tutup + '). Enter a correction reason of at least 10 characters.');
  return '[KOREKSI PERIODE TERTUTUP s.d. ' + st.tutup + ' · izin ' + st.oleh + ' s.d. ' + st.koreksi + ' · ' + String(alasan).trim().replace(/[\[\]]/g, '') + ']';
}
/** aplikasi: tutup (PEMBUAT/PENYETUJU, hanya maju) · buka koreksi / cabut (hanya PENYETUJU, bukan delegasi) */
function apiPeriode(action, tanggal, alasan) {
  var a = arRequireAccess_(), lock = arLock_(), st, msg;
  lock.waitLock(28000);
  try {
    st = arPeriodState_();
    if (action === 'tutup') {
      if (!a.maker && !a.approver) throw new Error('Not authorized to close the period.');
      if (!arValidDate_(tanggal) || tanggal >= arToday_()) throw new Error('The close date must be a past date.');
      if (st.tutup && tanggal <= st.tutup) throw new Error('The period is already closed through ' + st.tutup + ' (closed automatically on day ' + ((paramKV_('TUTUP_BUKU_TANGGAL') || {}).val || 3) + ' of each month at 00:00). The period cannot be moved back; use a Section Head correction permit.');
      setParamKV_('PERIODE_DITUTUP_SAMPAI', tanggal, 'Transaksi bertanggal s.d. tanggal ini ditolak (tutup buku)');
      msg = ['Periode ditutup s.d. ' + tanggal + ' oleh ' + a.email + '.'];
    } else if (action === 'buka' || action === 'cabut') {
      if (!a.approver) throw new Error('Only the Section Head (PENYETUJU) can open corrections for a closed period.');
      if (action === 'buka') {
        var maxD = ENGINE.addDays(arToday_(), 7);
        if (!st.tutup) throw new Error('No period has been closed yet.');
        if (!arValidDate_(tanggal) || tanggal < arToday_() || tanggal > maxD) throw new Error('A correction permit can last at most 7 days (through ' + maxD + ').');
        if (String(alasan || '').trim().length < 10) throw new Error('Enter a reason for opening corrections of at least 10 characters.');
        setParamKV_('PERIODE_KOREKSI_SAMPAI', tanggal, 'Izin koreksi transaksi di periode tertutup (diatur Section Head di aplikasi)');
        setParamKV_('PERIODE_KOREKSI_OLEH', a.email, 'Pemberi izin koreksi periode');
        setParamKV_('PERIODE_KOREKSI_ALASAN', String(alasan).trim(), 'Alasan izin koreksi periode');
        msg = ['Koreksi periode tertutup (s.d. ' + st.tutup + ') DIBUKA oleh ' + a.email + ' sampai ' + tanggal + '.', 'Alasan: ' + String(alasan).trim()];
      } else { setParamKV_('PERIODE_KOREKSI_SAMPAI', '', 'Izin koreksi transaksi di periode tertutup (diatur Section Head di aplikasi)'); msg = ['Izin koreksi periode tertutup DICABUT oleh ' + a.email + '.']; }
    } else throw new Error('Unrecognized period action.');
  } finally { lock.releaseLock(); }
  var w = arNotify_('Periode — ' + msg[0], msg);
  return plain_({ ok: true, periode: arPeriodState_(), warnings: w ? [w] : [] });
}
function arTutupPeriodeMenu() {
  arRequireAccess_();
  var ui = SpreadsheetApp.getUi(), st = arPeriodState_(), def = ENGINE.prevMonthEnd(arToday_());
  var q = ui.prompt('Tutup periode', 'Sekarang ditutup s.d.: ' + (st.tutup || '(belum ada)') + '\n\nTanggal tutup (yyyy-mm-dd), biasanya akhir bulan yang sudah dilaporkan ke pusat / LK ditutup. Contoh: ' + def + '\nSetelah ditutup, invoice & penerimaan bertanggal s.d. tanggal itu ditolak kecuali Section Head membuka koreksi di aplikasi.', ui.ButtonSet.OK_CANCEL);
  if (q.getSelectedButton() !== ui.Button.OK) return;
  var r = apiPeriode('tutup', String(q.getResponseText() || '').trim());
  ui.alert('Periode ditutup s.d. ' + r.periode.tutup + (r.warnings.length ? '\n\n' + r.warnings.join('\n') : ''));
}

// ---------------- Angka invoice dihitung ulang di server dari isian form + Master Proyek; aplikasi tidak dipercaya begitu saja.
function arCheckInvoiceNumbers_(p, nomor) {
  if (!p || !p.form) throw new Error('Invoice form data is missing. Save from the app.');
  var b = ENGINE.bootstrap(computeLight_(), email_()), pr = b.proyek.filter(function(x){return x.npp===p.npp;})[0];
  if (!pr) throw new Error('NPP ' + p.npp + ' not found in Master Proyek.');
  if (String(p.form.jenis || '') !== String(p.jenis || '') || (p.form.npp && String(p.form.npp) !== String(p.npp))) throw new Error('Form entries do not match the invoice data. Reload the app.');
  var ex = b.invoices.filter(function(i){return ENGINE.normInv(i.nomor)===nomor;})[0];
  var c = ENGINE.calcInvoice(ENGINE.invoiceInput(pr, Object.assign({}, p.form, { tgl: p.tgl, jenis: p.jenis }), ex ? ex.umk : 0));
  var want = { bruto: c.bruto, umk: c.umk || 0, retensi: c.retensi || 0, dpp: c.dpp, ppn: c.ppn, tagih: c.bayar }, beda = [];
  Object.keys(want).forEach(function(k) { if (Math.abs((Number(p[k]) || 0) - (Number(want[k]) || 0)) > 1) beda.push(k.toUpperCase() + ' in app Rp ' + ENGINE.fmtRp(Number(p[k]) || 0) + ' ≠ server Rp ' + ENGINE.fmtRp(Number(want[k]) || 0)); });
  if (beda.length) throw new Error('Invoice amounts differ from the server calculation (' + beda.join('; ') + '). Project data or the remaining down payment may have changed — reload the app and check again.');
  if (p.jenis === 'RETENSI' && (pr.template === 'IKK' || pr.template === 'KIDOH') && Number(pr.ret_pct) > 0 && want.ppn > 0 && !String((p.form || {}).ppnRetensiAlasan || '').trim())
    throw new Error('Retention release for project ' + p.npp + ' (retention position AKHIR, template ' + pr.template + ') is charged VAT, but retention VAT was already billed on the progress invoices. Turn off the VAT option, or enter the written basis from Tax / Accounting in the field below it.');
  // atribut pajak & template selalu dari Master Proyek
  p.template = pr.template; p.tarif = pr.tarif; p.dasar = pr.dasar_pph; p.retpos = c.retPos || 'TIDAK ADA';
  ['bruto','umk','retensi','dpp','ppn'].forEach(function(k){ p[k] = Number(want[k]) || 0; }); p.tagih = Number(want.tagih) || 0;
}

// ---------------- Script Properties dibatasi ±500 KB: hapus data turunan yang sudah tidak dipakai (aman dijalankan berulang).
function arPropsCleanup_() {
  var sp = PropertiesService.getScriptProperties(), all = sp.getProperties(), today = arToday_(), out = { hapus: 0, bytes: 0 };
  var pe = (paramKV_('FILE_PE_ID') || {}).val || '', voidLimit = ENGINE.addDays(today, -7), draftLimit = ENGINE.addDays(today, -186).slice(0, 7);
  Object.keys(all).forEach(function(k) {
    var v = String(all[k]), m, del = false;
    if ((m = k.match(/^AR28_VOID_/))) { var d = v.split('|')[1]; del = !d || d < voidLimit; } // token batal lama: RPC tertunda tidak mungkin masih berjalan
    else if (/^AR28_(?:W_)?PE28_/.test(k)) del = !!pe && k.indexOf('AR28_PE28_' + pe + '_') !== 0 && k.indexOf('AR28_W_PE28_' + pe + '_') !== 0; // digest / lebar kolom file P&E lama (ID Drive bisa memuat "_")
    else if ((m = k.match(/^AR28_VALDRAFT_(\d{4}-\d{2})/))) del = m[1] < draftLimit;           // draf usulan validasi > 6 bulan
    if (del) { sp.deleteProperty(k); out.hapus++; } else out.bytes += k.length + v.length;
  });
  // salinan lama di Document Properties yang sudah ada di Script Properties
  try { var dp = PropertiesService.getDocumentProperties(); if (dp) { var dAll = dp.getProperties(); Object.keys(dAll).forEach(function(k) { if (sp.getProperty('AR28_' + k) !== null || /^(VOID_|PE28_|W_PE28_)/.test(k)) { dp.deleteProperty(k); out.hapus++; } }); } } catch (e) {}
  arProps_().setProperty('PROPS_BYTES', String(out.bytes));
  return out;
}

/** pengajuan / keputusan yang mengubah data periode tertutup (pembatalan invoice, lembar validasi) hanya boleh saat izin koreksi aktif */
function arPeriodDateOf_(kind, d) {
  if (kind === 'BATAL') return d && d.inv ? String(d.inv.tgl || '') : '';
  if (kind === 'REVP') return d && d.tgl ? String(d.tgl) : '';
  if (kind === 'VAL' || kind === 'VALX') return /^\d{4}-\d{2}$/.test(String(d && d.ref || '')) ? arMonthEnd_(String(d.ref)) : '';
  return '';
}
function arGuardClosed_(kind, d, what) {
  var t = arPeriodDateOf_(kind, d), st = arPeriodState_();
  if (!t || !st.tutup || t > st.tutup) return '';
  if (!st.koreksi) throw new Error(what + ' concerns a period already closed (through ' + st.tutup + '). After month-end close, data in that period cannot be changed or cancelled unless the Section Head opens a correction permit.');
  return t;
}

// ---------------- Jejak edit langsung: invoice & penerimaan seharusnya dicatat lewat aplikasi.
// Tulisan aplikasi (setValues) tidak memicu onEdit, jadi setiap event di kolom terkontrol = ketikan / tempel manual di sheet.
var AR_AUDIT_SHEET_ = 'Jejak Edit Langsung';
var AR_AUDIT_FREE_ = { 'Log Penerimaan': ['ket', 'sap_doc', 'bupot'] }; // kolom yang memang boleh diisi di sheet
function arAuditControlled_(name, c) {
  if (name === 'Log Penerimaan') return (AR_AUDIT_FREE_[name].indexOf(c[0]) < 0 && colType_(c) !== 'auto') || ['idp'].indexOf(c[0]) >= 0;
  return colType_(c) === 'app'; // Monitoring Invoice: angka & identitas invoice dari aplikasi; kolom manual (submit, faktur, BA, …) bebas
}
function arAuditSheet_() {
  var ss = arSpreadsheet_(), sh = ss.getSheetByName(AR_AUDIT_SHEET_);
  if (!sh) {
    sh = ss.insertSheet(AR_AUDIT_SHEET_);
    sh.getRange(1, 1, 1, 10).setValues([['WAKTU', 'PENGEDIT', 'SHEET', 'SEL', 'BARIS (ID / NOMOR)', 'KOLOM', 'NILAI LAMA', 'NILAI BARU', 'CATATAN', 'DITINJAU (ISI SECTION HEAD)']])
      .setFontWeight('bold').setBackground(C.navy).setFontColor('#FFFFFF');
    sh.setFrozenRows(1); sh.setTabColor('#C00000');
  }
  return sh;
}
function arAuditLog_(rows, subject) {
  if (!rows.length) return;
  var sh = arAuditSheet_();
  roomFor_(sh, sh.getLastRow() + rows.length);
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, 10).setNumberFormat('@').setValues(rows.map(function(r){ return r.map(function(v){ return arSafeCell_(String(v === null || v === undefined ? '' : v)); }); }));
  arNotify_(subject, ['Perubahan berikut dibuat langsung di spreadsheet, tidak melalui aplikasi (tanpa validasi, tanpa tutup periode):', ''].concat(rows.slice(0, 40).map(function(r) {
    return '• ' + r[2] + ' ' + r[3] + ' · ' + r[4] + ' · ' + r[5] + ': "' + r[6] + '" → "' + r[7] + '"' + (r[8] ? ' · ' + r[8] : '') + ' · oleh ' + r[1];
  })).concat(rows.length > 40 ? ['… dan ' + (rows.length - 40) + ' perubahan lain (lihat sheet ' + AR_AUDIT_SHEET_ + ')'] : []));
}
function arEditor_(e) { var m = ''; try { m = e && e.user && e.user.getEmail ? e.user.getEmail() : ''; } catch (x) {} if (!m) m = email_(); return m || '(tidak terbaca — akun Gmail pribadi)'; }
/** dipanggil onEditInstalled untuk Monitoring Invoice / Log Penerimaan */
function arAuditEdit_(e, sh, name) {
  var schema = name === 'Log Penerimaan' ? ENGINE.LOG : ENGINE.REG, map = colMap_(sh, schema, schema === ENGINE.REG ? ENGINE.REG_ALIAS : null), byCol = {};
  schema.cols.forEach(function(c) { if (map[c[0]] && arAuditControlled_(name, c)) byCol[map[c[0]]] = c; });
  var r0 = Math.max(e.range.getRow(), schema.firstRow), r1 = e.range.getLastRow(), c0 = e.range.getColumn(), c1 = e.range.getLastColumn();
  if (r1 < r0) return;
  var hit = []; for (var c = c0; c <= c1; c++) if (byCol[c]) hit.push(c);
  if (!hit.length) return;
  var single = e.range.getNumRows() === 1 && e.range.getNumColumns() === 1, keyCol = map[name === 'Log Penerimaan' ? 'idp' : 'nomor'], tglCol = map.tgl;
  var n = Math.min(r1 - r0 + 1, 200), keys = keyCol ? sh.getRange(r0, keyCol, n, 1).getDisplayValues() : [], tgls = tglCol ? sh.getRange(r0, tglCol, n, 1).getValues() : [];
  var vals = sh.getRange(r0, c0, n, c1 - c0 + 1).getDisplayValues(), st = arPeriodState_(), who = arEditor_(e), now = arNow_(), out = [];
  for (var i = 0; i < n; i++) hit.forEach(function(c) {
    var t = tgls[i] ? (isDate_(tgls[i][0]) ? dstr_(tgls[i][0]) : String(tgls[i][0] || '')) : '';
    var note = st.tutup && t && t <= st.tutup ? 'PERIODE TERTUTUP s.d. ' + st.tutup : '';
    out.push([now, who, name, sh.getRange(r0 + i, c).getA1Notation(), keys[i] ? keys[i][0] : '', byCol[c][1], single ? (e.oldValue === undefined ? '' : e.oldValue) : '(beberapa sel)', vals[i][c - c0], note, '']);
  });
  arAuditLog_(out, 'Edit langsung di sheet ' + name + ' (' + out.length + ' sel)');
}
/** baris yang hilang (dihapus / ID diubah) dibandingkan pembaruan sebelumnya — dipanggil rebuildAll_.
 *  Daftar ID disimpan di sheet tersembunyi (Script Properties dibatasi 9 KB per nilai); ditulis hanya bila berubah. */
function arAuditCounts_(input, opt) {
  var ixI = ENGINE.colIndex(ENGINE.LOG, 'idp'), ixN = ENGINE.colIndex(ENGINE.REG, 'nomor'), pr = arProps_(), ss = arSpreadsheet_();
  var now = { LOG: [], REG: [] }, seen = {};
  var man = ENGINE.LOG.cols.map(function(c, i) { return c[2] === 'manual' ? i : -1; }).filter(function(i) { return i >= 0; });
  (input.log || []).forEach(function(r) {
    if (!man.some(function(i) { return r[i] !== '' && r[i] !== null && r[i] !== undefined; })) return; // baris kosong (hanya sisa ID otomatis) bukan data
    var k = String(r[ixI] || '').trim(); if (k && !seen['L' + k]) { seen['L' + k] = 1; now.LOG.push(k); } });
  (input.register || []).forEach(function(r) { var k = ENGINE.normInv(String(r[ixN] || '').trim()); if (k && !seen['R' + k]) { seen['R' + k] = 1; now.REG.push(k); } });
  var dig = arDigest_(JSON.stringify(now));
  if (pr.getProperty('AUDIT_IDS_DIGEST') === dig) return;
  var sh = ss.getSheetByName('_Jejak ID');
  if (!sh) { sh = ss.insertSheet('_Jejak ID'); sh.getRange(1, 1, 1, 2).setValues([['ID PENERIMAAN', 'NOMOR INVOICE']]); sh.hideSheet(); }
  var prev = { LOG: [], REG: [] }, last = sh.getLastRow(), had = !!pr.getProperty('AUDIT_IDS_DIGEST');
  if (last > 1) sh.getRange(2, 1, last - 1, 2).getValues().forEach(function(r) { if (r[0]) prev.LOG.push(String(r[0])); if (r[1]) prev.REG.push(String(r[1])); });
  var out = [], ts = arNow_();
  if (had && !(opt && opt.silent)) [['LOG', 'Log Penerimaan'], ['REG', 'Monitoring Invoice']].forEach(function(x) {
    var cur = {}; now[x[0]].forEach(function(k) { cur[k] = 1; });
    prev[x[0]].forEach(function(k) { if (!cur[k]) out.push([ts, '(terdeteksi saat pembaruan)', x[1], '', k, 'BARIS', k, '(hilang / dihapus / ID diubah)', 'BARIS HILANG', '']); });
  });
  var n = Math.max(now.LOG.length, now.REG.length, 1), grid = [];
  for (var i = 0; i < n; i++) grid.push([now.LOG[i] || '', now.REG[i] || '']);
  if (last > 1) sh.getRange(2, 1, last - 1, 2).clearContent();
  sh.getRange(2, 1, n, 2).setNumberFormat('@').setValues(grid);
  pr.setProperty('AUDIT_IDS_DIGEST', dig);
  if (out.length) arAuditLog_(out, 'Baris hilang dari ' + out.map(function(r){return r[2];}).filter(function(v,i,a){return a.indexOf(v)===i;}).join(' & ') + ' (' + out.length + ')');
}


// =====================================================================  ARSIP PROYEK LAMA (AKTIF = ARSIP)
/* Baris invoice, penerimaan & tagbrut proyek ARSIP (selesai & lunas sebelum sistem dipakai) dipindah ke sheet "Arsip Data Lama"
 * (bukan disembunyikan: baris tersembunyi rawan tertimpa saat tempel / tarik isi). Kolom: ASAL SHEET · NPP · KUNCI · DIPINDAH · isi baris asal. */
var AR_ARSIP_SHEET_ = 'Arsip Data Lama', AR_ARSIP_META_ = 4, AR_ARSIP_FIRST_ = 5;
function arArsipSheet_(ss) {
  var sh = ss.getSheetByName(AR_ARSIP_SHEET_);
  if (sh) return sh;
  sh = ss.insertSheet(AR_ARSIP_SHEET_);
  sh.getRange('A1').setValue('ARSIP DATA LAMA — proyek AKTIF = ARSIP (selesai & lunas sebelum sistem dipakai)').setFontSize(13).setFontWeight('bold').setFontColor(C.navy);
  sh.getRange('A2').setValue('Hanya baca. Diabaikan seluruh perhitungan. Kembalikan lewat menu AR CSU ▸ Kembalikan proyek dari arsip…');
  sh.getRange(4, 1, 1, 5).setValues([['ASAL SHEET', 'NPP', 'KUNCI', 'DIPINDAH (WAKTU · OLEH)', 'ISI BARIS ASAL (kolom A dst.) →']]).setFontWeight('bold').setBackground(C.navy).setFontColor('#FFFFFF');
  sh.setFrozenRows(4); sh.setTabColor('#7F7F7F');
  try { sh.protect().setDescription('Arsip data lama — hanya baca').setWarningOnly(true); } catch (e) {}
  return sh;
}
/** kunci baris arsip (dipakai ENGINE agar memorial lama tidak dibatalkan) */
function readArsipKeys_() {
  var sh = arSpreadsheet_().getSheetByName(AR_ARSIP_SHEET_);
  if (!sh || sh.getLastRow() < AR_ARSIP_FIRST_) return [];
  return sh.getRange(AR_ARSIP_FIRST_, 3, sh.getLastRow() - AR_ARSIP_FIRST_ + 1, 1).getValues().map(function(r) { return String(r[0] || ''); }).filter(function(k) { return /^(INV|TRM)\|/.test(k); });
}
/** sisa tagihan proyek ARSIP bila proyek dihitung seperti biasa (pengaman sebelum dipindah) */
function arArsipSisa_(input, npps) {
  var x = JSON.parse(JSON.stringify(input)), ia = ENGINE.colIndex(ENGINE.MPRO, 'aktif'), ip = ENGINE.colIndex(ENGINE.MPRO, 'npp'), out = {};
  x.proyek.forEach(function(r) { if (npps.indexOf(String(r[ip]).trim()) >= 0) r[ia] = 'YA'; });
  x.light = true;
  ENGINE.compute(x).ctx.invs.forEach(function(i) {
    if (npps.indexOf(i.npp) < 0) return;
    var o = out[i.npp] || (out[i.npp] = { inv: 0, sisa: 0, terbuka: [] });
    o.inv++;
    if (!i._batal && !i._res.closed && i._res.sisa > 1) { o.sisa += i._res.sisa; o.terbuka.push(i.nomor.replace(/^KU\.02\.02\/WB-4C\./, '')); }
  });
  return out;
}
function arArsipPindahMenu() {
  var ui = SpreadsheetApp.getUi(), acc = arRequireAccess_();
  if (acc.delegate) throw new Error('Menu ini hanya untuk Pembuat / Section Head.');
  var ss = arSpreadsheet_(), input = readAll_(), res = ENGINE.compute(JSON.parse(JSON.stringify(Object.assign({}, input, { light: true })))), R = res.arsipRows, npps = R.npp;
  if (!npps.length) { ui.alert('Arsip data lama', 'Belum ada proyek dengan AKTIF = ARSIP di Master Proyek.', ui.ButtonSet.OK); return; }
  if (!R.reg.length && !R.log.length && !R.tag.length) { ui.alert('Arsip data lama', 'Semua baris proyek ARSIP (' + npps.join(', ') + ') sudah dipindah.', ui.ButtonSet.OK); return; }
  var sisa = arArsipSisa_(input, npps), warn = [];
  var lines = npps.map(function(n) { var p = res.ctx.M.arsip[n], o = sisa[n] || { inv: 0, sisa: 0, terbuka: [] };
    if (o.sisa > 1) warn.push('• ' + n + ': sisa Rp ' + ENGINE.fmtRp(o.sisa) + ' (' + o.terbuka.slice(0, 8).join(', ') + (o.terbuka.length > 8 ? ', …' : '') + ')');
    return '• ' + n + ' ' + p.pelanggan + ' — ' + o.inv + ' invoice'; });
  var msg = 'Baris berikut dipindah ke sheet "' + AR_ARSIP_SHEET_ + '" lalu dihapus dari sheet kerja:\n\n' + lines.join('\n') +
    '\n\nMonitoring Invoice: ' + R.reg.length + ' baris · Log Penerimaan: ' + R.log.length + ' baris · Input Tagbrut: ' + R.tag.length + ' baris.' +
    '\nData tidak hilang (bisa dikembalikan), tercatat di ' + AR_AUDIT_SHEET_ + ' dan dikirim ke email Section Head. Lanjutkan?';
  if (ui.alert('Pindahkan proyek ARSIP', msg, ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  if (warn.length) {
    var r2 = ui.prompt('⚠ Masih ada sisa menurut data', 'Bila proyek dihitung biasa, data menunjukkan sisa tagihan:\n\n' + warn.join('\n') +
      '\n\nBisa karena data lama tidak lengkap. Bila Anda yakin proyek sudah lunas, ketik PINDAHKAN untuk tetap memindahkan.', ui.ButtonSet.OK_CANCEL);
    if (r2.getSelectedButton() !== ui.Button.OK || String(r2.getResponseText()).trim().toUpperCase() !== 'PINDAHKAN') { ui.alert('Dibatalkan. Tidak ada data yang dipindah.'); return; }
  }
  var lock = arLock_(), moved = []; lock.waitLock(28000);
  try {
    var ar = arArsipSheet_(ss), stamp = arNow_() + ' · ' + acc.email, out = [], aud = [];
    var regSh = ss.getSheetByName('Monitoring Invoice'), regMap = colMap_(regSh, ENGINE.REG, ENGINE.REG_ALIAS);
    var nppOf = {}; input.register.forEach(function(r) { var no = ENGINE.normInv(String(r[ENGINE.colIndex(ENGINE.REG, 'nomor')] || '')); if (no) nppOf[no] = String(r[ENGINE.colIndex(ENGINE.REG, 'npp')] || '').trim(); });
    // Monitoring Invoice & Log Penerimaan: seluruh baris dipindah lalu dihapus
    [['Monitoring Invoice', ENGINE.REG, R.reg], ['Log Penerimaan', ENGINE.LOG, R.log]].forEach(function(x) {
      var sh = ss.getSheetByName(x[0]); if (!sh || !x[2].length) return;
      var lc = sh.getLastColumn(), map = colMap_(sh, x[1], x[1] === ENGINE.REG ? ENGINE.REG_ALIAS : null), data = sh.getRange(1, 1, sh.getLastRow(), lc).getValues();
      out.push(['#JUDUL ' + x[0], '', '', stamp].concat(data[x[1].headerRow - 1]));
      x[2].forEach(function(r) {
        var v = data[r - 1], key, npp;
        if (x[1] === ENGINE.REG) { var no = String(v[map.nomor - 1] || '').trim(); key = 'INV|' + ENGINE.normInv(no); npp = String(v[map.npp - 1] || '').trim(); aud.push([arNow_(), acc.email, x[0], 'baris ' + r, no, 'BARIS', '', 'dipindah ke ' + AR_ARSIP_SHEET_, 'DIPINDAH KE ARSIP (proyek ARSIP)', '']); }
        else { var idp = String(v[map.idp - 1] || '').trim(), inv = String(v[map.inv - 1] || '').trim(); key = 'TRM|' + idp; npp = nppOf[ENGINE.normInv(inv)] || ''; aud.push([arNow_(), acc.email, x[0], 'baris ' + r, idp + ' (' + inv + ')', 'BARIS', '', 'dipindah ke ' + AR_ARSIP_SHEET_, 'DIPINDAH KE ARSIP (proyek ARSIP)', '']); }
        out.push([x[0], npp, key, stamp].concat(v));
      });
      moved.push([sh, x[2]]);
    });
    // Input Tagbrut: hanya kolom tagbrut yang dipindah & dikosongkan (sheet yang sama juga memuat catatan)
    var tsh = ss.getSheetByName(ENGINE.TAG.sheet);
    if (tsh && R.tag.length) {
      var tm = colMap_(tsh, ENGINE.TAG), tcols = ENGINE.TAG.cols.map(function(c) { return tm[c[0]]; }).filter(Boolean);
      out.push(['#JUDUL ' + ENGINE.TAG.sheet, '', '', stamp].concat(tcols.map(function(j) { return tsh.getRange(ENGINE.TAG.headerRow, j).getValue(); })));
      R.tag.forEach(function(r) {
        var v = tcols.map(function(j) { return tsh.getRange(r, j).getValue(); });
        out.push([ENGINE.TAG.sheet, String(tsh.getRange(r, tm.npp).getValue()).trim(), 'TAG|' + tcols.join(','), stamp].concat(v));
        tcols.forEach(function(j) { tsh.getRange(r, j).clearContent(); });
      });
    }
    var w = out.reduce(function(m, r) { return Math.max(m, r.length); }, 0), row0 = Math.max(ar.getLastRow() + 1, AR_ARSIP_FIRST_);
    if (ar.getMaxColumns() < w) ar.insertColumnsAfter(ar.getMaxColumns(), w - ar.getMaxColumns());
    roomFor_(ar, row0 + out.length - 1);
    ar.getRange(row0, 1, out.length, w).setValues(out.map(function(r) { while (r.length < w) r.push(''); return r; }));
    out.forEach(function(r, i) { if (/^#JUDUL/.test(r[0])) ar.getRange(row0 + i, 1, 1, w).setFontWeight('bold').setBackground(C.tot); });
    // hapus dari bawah ke atas, per rentang berurutan
    moved.forEach(function(m) {
      var rows = m[1].slice().sort(function(a, b) { return b - a; });
      for (var i = 0; i < rows.length;) { var j = i; while (j + 1 < rows.length && rows[j + 1] === rows[j] - 1) j++; m[0].deleteRows(rows[j], rows[i] - rows[j] + 1); i = j + 1; }
    });
    SpreadsheetApp.flush();
    arAuditCounts_(readAll_(), { silent: true });
    arAuditLog_(aud, 'Proyek ARSIP dipindah ke ' + AR_ARSIP_SHEET_ + ' (' + npps.join(', ') + ')');
  } finally { lock.releaseLock(); }
  rebuildAll_();
  ui.alert('Selesai', 'Baris proyek ARSIP sudah dipindah ke sheet "' + AR_ARSIP_SHEET_ + '" dan laporan sudah diperbarui.', ui.ButtonSet.OK);
}
function arArsipKembalikanMenu() {
  var ui = SpreadsheetApp.getUi(), acc = arRequireAccess_();
  if (acc.delegate) throw new Error('Menu ini hanya untuk Pembuat / Section Head.');
  var ss = arSpreadsheet_(), ar = ss.getSheetByName(AR_ARSIP_SHEET_);
  if (!ar || ar.getLastRow() < AR_ARSIP_FIRST_) { ui.alert('Sheet "' + AR_ARSIP_SHEET_ + '" masih kosong.'); return; }
  var r = ui.prompt('Kembalikan proyek dari arsip', 'Ketik NPP proyek yang dikembalikan (mis. 244C 0660 JT):', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var npp = String(r.getResponseText()).trim().toUpperCase().replace(/\s+/g, ' ');
  var data = ar.getRange(AR_ARSIP_FIRST_, 1, ar.getLastRow() - AR_ARSIP_FIRST_ + 1, ar.getLastColumn()).getValues(), pick = [];
  data.forEach(function(v, i) { if (!/^#JUDUL/.test(String(v[0])) && String(v[1]).trim().toUpperCase() === npp) pick.push({ i: i, v: v }); });
  if (!pick.length) { ui.alert('Tidak ada baris arsip untuk NPP ' + npp + '.'); return; }
  if (ui.alert('Kembalikan ' + pick.length + ' baris', pick.length + ' baris NPP ' + npp + ' dikembalikan ke sheet asalnya (ditambahkan di bawah). Setelah itu ubah AKTIF proyek di Master Proyek menjadi YA atau SELESAI. Lanjutkan?', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  var lock = arLock_(); lock.waitLock(28000);
  try {
    var bySheet = {}; pick.forEach(function(p) { (bySheet[p.v[0]] = bySheet[p.v[0]] || []).push(p); });
    Object.keys(bySheet).forEach(function(name) {
      var sh = ss.getSheetByName(name); if (!sh) return;
      var L = bySheet[name];
      if (name === ENGINE.TAG.sheet) {
        L.forEach(function(p) {
          var cols = String(p.v[2]).replace(/^TAG\|/, '').split(',').map(Number), tm = colMap_(sh, ENGINE.TAG), npc = tm.npp, row = ENGINE.TAG.firstRow;
          var vals = sh.getRange(ENGINE.TAG.firstRow, npc, Math.max(sh.getLastRow() - ENGINE.TAG.firstRow + 1, 1), 1).getValues();
          while (row - ENGINE.TAG.firstRow < vals.length && String(vals[row - ENGINE.TAG.firstRow][0]).trim()) row++;
          roomFor_(sh, row); cols.forEach(function(j, k) { sh.getRange(row, j).setValue(p.v[AR_ARSIP_META_ + k]); });
        });
      } else {
        var lc = sh.getLastColumn(), rows = L.map(function(p) { return p.v.slice(AR_ARSIP_META_, AR_ARSIP_META_ + lc); });
        roomFor_(sh, sh.getLastRow() + rows.length); sh.getRange(sh.getLastRow() + 1, 1, rows.length, lc).setValues(rows);
      }
    });
    pick.slice().reverse().forEach(function(p) { ar.deleteRow(AR_ARSIP_FIRST_ + p.i); });
    arAuditLog_([[arNow_(), acc.email, AR_ARSIP_SHEET_, '', npp, 'BARIS', '', pick.length + ' baris dikembalikan', 'DIKEMBALIKAN DARI ARSIP', '']], 'Proyek ' + npp + ' dikembalikan dari ' + AR_ARSIP_SHEET_);
  } finally { lock.releaseLock(); }
  ui.alert('Selesai', 'Baris dikembalikan. Ubah AKTIF proyek ' + npp + ' di Master Proyek menjadi YA / SELESAI, lalu jalankan Perbarui semua laporan.', ui.ButtonSet.OK);
}


/** Analisis Bulanan AR untuk menu Tutup Buku Bulanan di aplikasi: baca snapshot terakhir bulan itu; buat = hanya Pembuat & bulan sudah tutup buku */
function apiAnalisis(month, buat) {
  var acc = arRequireAccess_(!buat);
  month = String(month || '');
  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(month)) throw new Error('Month must be yyyy-mm.');
  var st = arPeriodState_(), tutup = !!st.tutup && st.tutup >= arMonthEnd_(month), rows;
  try { rows = arMonthlyRows_(); } catch (e) { return { ada: false, tutup: tutup, err: 'Analysis archive not installed' }; }
  var L = rows.filter(function(r) { return String(r[1]) === month; });
  if (!L.length && buat) {
    if (!acc.maker) throw new Error('Monthly analysis is created by the Preparer.');
    if (!tutup) throw new Error('Month ' + month + ' is not closed yet. The analysis is created after month-end close (day 3 of the following month).');
    arCreateMonthly_(month, false, '');
    L = arMonthlyRows_().filter(function(r) { return String(r[1]) === month; });
  }
  if (!L.length) return { ada: false, tutup: tutup };
  var r = L[L.length - 1], m = {}; try { m = JSON.parse(r[8]); } catch (e) {}
  return { ada: true, tutup: tutup, id: String(r[0]), rev: Number(r[2]) || 0, waktu: String(r[4]), alasan: String(r[7] || ''), m: m, narasi: String(r[9] || ''), ai: String(r[10] || ''), aiStatus: String(r[11] || '') };
}

/** v2.8.8: checklist Month-End AR Closing (bulan yyyy-mm; kosong = bulan lalu bila sudah lewat tanggal tutup, selain itu bulan berjalan) */
function apiClosing(month) {
  arRequireAccess_(true);
  month = String(month || ''); if (month && !/^\d{4}-(0[1-9]|1[0-2])$/.test(month)) throw new Error('Month must be yyyy-mm.');
  var res = computeLight_('memo'); if (!month) month = res.ctx.today.substring(0, 7);
  return plain_(ENGINE.closingData(res.ctx, month));
}
/** v2.8.5: data layar Overview (bulan yyyy-mm; kosong = bulan berjalan). Hitungan ringan + cache 5 menit (gugur setiap ada penyimpanan). */
function apiOverview(month) {
  arRequireAccess_(true);
  month = String(month || '');
  if (month && !/^\d{4}-(0[1-9]|1[0-2])$/.test(month)) throw new Error('Month must be yyyy-mm.');
  var key = arBootKey_('OVERVIEW|' + month + '|' + arToday_()), hit = arCacheGet_(key);
  if (hit) return JSON.parse(hit);
  var res = computeLight_(true), o = ENGINE.overview(res.ctx, month);
  o.reportTime = ''; try { o.reportTime = arProps_().getProperty('REPORT_TIME') || ''; } catch (e) {}
  var out = plain_(o); arCachePut_(key, JSON.stringify(out), 300);
  return out;
}


/** v2.8.5: Monitoring Invoice dipendekkan — kolom rincian hitungan, dokumen, alokasi penerimaan & kolom sistem dikelompokkan (tombol +/− di atas kolom),
 *  tidak dihapus karena dipakai aplikasi untuk mencetak ulang & merevisi invoice. Kelompok hanya dibuat (dan ditutup) sekali; bila dibuka pengguna, dibiarkan. */
var AR_REG_GROUP_ = ['template','bruto','umk','retensi','retpos','netto','tarif','dasar','tagih','ttd','rev','uraian','jt_akt','link_inv','link_fp','ba_no','ba_tgl',
  'ke_ppn','ke_net','pot_pph','adm','sel_pph','link_paket','sumber','dibuat','json','memorial'];
function arGroupRegister_(ss) {
  var sh=ss.getSheetByName(ENGINE.REG.sheet);if(!sh)return 0;
  var map=colMap_(sh,ENGINE.REG,ENGINE.REG_ALIAS),cols=AR_REG_GROUP_.map(function(k){return map[k];}).filter(Boolean).sort(function(a,b){return a-b;}),runs=[],made=0;
  cols.forEach(function(c){var r=runs[runs.length-1];if(r&&c===r[1]+1)r[1]=c;else runs.push([c,c]);});
  runs.forEach(function(r){
    try{
      if(sh.getColumnGroupDepth(r[0])>0)return; // sudah dikelompokkan
      sh.getRange(1,r[0],1,r[1]-r[0]+1).shiftColumnGroupDepth(1);made++;
      try{sh.getColumnGroup(r[0],1).collapse();}catch(e){}
    }catch(e){}
  });
  return made;
}

/** v2.8.5: file turunan untuk bagian Pajak — REKAP INVOICE KELUAR (per tahun), BUK PEN per bulan (berdasarkan tanggal faktur) & KELENGKAPAN DOKUMEN.
 *  Isi ditimpa sistem; tab bulan lama tidak dihapus. Pengguna peran PAJAK (LIHAT) di Parameter ▸ Pengguna & Peran diberi akses lihat. */
var AR_PAJAK_TITLE_='REKAP INVOICE KELUAR & BUKU PENJUALAN — CSU (PAJAK)';
function arPushPajak_(res,opt) {
  opt=opt||{};
  var x=paramKV_('FILE_PAJAK_ID');if(!(x&&x.val)&&!opt.create)return 0;
  var ss=derivedFile_('FILE_PAJAK_ID',AR_PAJAK_TITLE_),ctx=res.ctx,today=ctx.today,stamp=arNow_(),y=+today.substring(0,4),m=+today.substring(5,7),count=0,names=[];
  if(ss.getSpreadsheetTimeZone()!==TZ)ss.setSpreadsheetTimeZone(TZ);
  var put=function(name,rep,t1,t2,idx){
    var sh=ss.getSheetByName(name);if(!sh)sh=ss.insertSheet(name,Math.min(idx,ss.getSheets().length));names.push(name);
    var key='PJ285_'+ss.getId()+'_'+name,dg=ENGINE.hash(JSON.stringify(rep.blocks.map(function(b){return b.rows;}))+'|'+ENGINE.VERSION);
    sh.getRange('A3').setValue('Diperbarui '+stamp+' WIB · Sistem AR CSU v'+ENGINE.VERSION+' · isi file ini ditimpa sistem, jangan diubah manual.');
    if(arProps_().getProperty(key)===dg&&sh.getLastRow()>6)return;
    sh.getRange('A1').setValue(t1).setFontSize(14).setFontWeight('bold').setFontColor(C.navy);
    sh.getRange('A2').setValue(t2);
    writeReport_(sh,rep,'PAJAK_'+name,{key:key});sh.setHiddenGridlines(true);sh.setFrozenRows(7);
    arProps_().setProperty(key,dg);count++;
  };
  put('KELENGKAPAN DOKUMEN',ENGINE.pajakReport(ctx,'DOK'),'KELENGKAPAN DOKUMEN TAGIHAN & FAKTUR PAJAK','PT Wijaya Karya Beton Tbk · Concrete Services Unit · berkas tagihan ber-TTD & PDF faktur diunggah AR lewat aplikasi (Document filing)',0);
  put('REKAP '+y,ENGINE.pajakReport(ctx,'REKAP',y),'REKAP INVOICE KELUAR TAHUN '+y,'PT Wijaya Karya Beton Tbk · Concrete Services Unit · invoice batal tidak dihitung',1);
  // BUK PEN: bulan berjalan & bulan-bulan sebelumnya pada tahun ini (Januari: termasuk Desember tahun lalu)
  var months=[];for(var k=m;k>=1;k--)months.push([y,k]);if(m===1)months.push([y-1,12]);
  months.forEach(function(ym,ix){var rep=ENGINE.pajakReport(ctx,'BUKPEN',ym[0],ym[1]);
    put('BUK PEN '+ENGINE.MONTHS[ym[1]-1].substring(0,3).toUpperCase()+' '+String(ym[0]).substring(2),rep,'BUKU PENJUALAN — '+ENGINE.monthLabel(ym[0],ym[1]).toUpperCase(),'PT Wijaya Karya Beton Tbk · Concrete Services Unit · format BUK PEN · '+rep.n+' faktur',2+ix);});
  arOrderNamed_(ss,names);dropDefaultSheet_(ss);
  var emails=(res.params.users||[]).filter(function(u){return /^PAJAK/.test(u.peran)&&u.email;}).map(function(u){return u.email;});
  if(emails.length)shareTo_(ss,emails,false);
  return count;
}
/** menu Pengaturan ▸ Aktifkan file Pajak: membuat file (sekali) di akun pemilik lalu mengisi & membagikannya */
function arEnablePajak() {
  arRequireMaker_();var ui=SpreadsheetApp.getUi(),lock=arLock_(),n;lock.waitLock(28000);
  try{n=arPushPajak_(compute_(),{create:true});}finally{lock.releaseLock();}
  var x=paramKV_('FILE_PAJAK_ID'),url=x&&x.val?'https://docs.google.com/spreadsheets/d/'+x.val:'-';
  ui.showModalDialog(HtmlService.createHtmlOutput('<div style="font:13px/1.5 Arial,sans-serif"><p>File Pajak siap ('+n+' tab diperbarui):</p><p><a target="_blank" href="'+url+'">'+AR_PAJAK_TITLE_+'</a></p><p>Isi email bagian Pajak di <b>Parameter ▸ Pengguna &amp; Peran</b> dengan peran <b>PAJAK (LIHAT)</b> — akses lihat diberikan otomatis setiap sinkron. File diperbarui otomatis tiap jam bersama file Akuntansi &amp; P&amp;E.</p></div>').setWidth(520).setHeight(230),'File Pajak');
}

/** v2.8.6: filter yang dipasang pengguna di sheet isian/log tidak ikut melebar saat sistem menambah baris di bawah (perilaku Google Sheets),
 *  sehingga baris baru tampil walau tidak cocok dengan filter. Setiap pembaruan laporan, rentang filter dilebarkan sampai baris terakhir sheet;
 *  pilihan filter per kolom dipertahankan. Tampilan filter (filter view) ikut dilebarkan bila layanan lanjutan "Google Sheets API" diaktifkan. */
var AR_FILTER_SHEETS_ = ['Monitoring Invoice','Log Penerimaan','Log Persetujuan','Log Surat Penagihan','Master Pelanggan','Master Proyek','Riwayat Adendum','Input Tagbrut','Referensi P&E','Riwayat Revisi','Jejak Edit Langsung','Arsip Data Lama'];
function arExtendFilters_(ss) {
  var n = 0;
  AR_FILTER_SHEETS_.forEach(function(name){
    try {
      var sh = ss.getSheetByName(name), f = sh && sh.getFilter(); if (!f) return;
      var r = f.getRange(), maxR = sh.getMaxRows(); if (r.getLastRow() >= maxR) return;
      var c0 = r.getColumn(), crit = [];
      for (var c = c0; c <= r.getLastColumn(); c++) { var k = f.getColumnFilterCriteria(c); if (k) crit.push([c, k]); }
      f.remove();
      var nf;
      try { nf = sh.getRange(r.getRow(), c0, maxR - r.getRow() + 1, r.getNumColumns()).createFilter(); n++; }
      catch (e) { nf = r.createFilter(); } // tidak berhak mengubah: pasang lagi filter semula, jangan sampai filter hilang
      crit.forEach(function(x){ try { nf.setColumnFilterCriteria(x[0], x[1]); } catch (e) {} });
    } catch (e) {}
  });
  // tampilan filter (Data ▸ Tampilan filter): hanya bila layanan lanjutan Sheets API aktif di proyek Apps Script
  try {
    if (typeof Sheets === 'undefined') return n;
    var meta = Sheets.Spreadsheets.get(ss.getId(), { fields: 'sheets(properties(sheetId,title,gridProperties.rowCount),filterViews(filterViewId,range))' }), req = [];
    (meta.sheets || []).forEach(function(s){
      if (AR_FILTER_SHEETS_.indexOf(s.properties.title) < 0) return;
      var rows = s.properties.gridProperties.rowCount;
      (s.filterViews || []).forEach(function(v){ var rg = v.range || {};
        if (rg.endRowIndex !== undefined && rg.endRowIndex < rows) { rg.endRowIndex = rows; req.push({ updateFilterView: { filter: { filterViewId: v.filterViewId, range: rg }, fields: 'range' } }); } });
    });
    if (req.length) { Sheets.Spreadsheets.batchUpdate({ requests: req }, ss.getId()); n += req.length; }
  } catch (e) {}
  return n;
}

// =====================================================================  v2.8.8: MASTER DATA (pelanggan, proyek, adendum) LEWAT APLIKASI
var AR_MD_HIST_ = 'Riwayat Master Data';
var AR_MD_TEXT_ = ['npwp', 'nitku', 'kode_bp', 'npp', 'kontrak', 'kode_jasa', 'npp_produk', 'npp_adendum', 'pola_uraian', 'pola_kw'];
function arMdRead_(ss) {
  var obj = function (sheetName, schema, alias) { var sh = ss.getSheetByName(sheetName); if (!sh) return [];
    return readSchema_(sh, schema, alias).map(function (r, i) { var o = { _row: schema.firstRow + i }; schema.cols.forEach(function (c, j) { o[c[0]] = r[j]; }); return o; })
      .filter(function (o) { return String(o[schema.cols[0][0]] || '').trim(); }); };
  return { pel: obj('Master Pelanggan', ENGINE.MPEL), pro: obj('Master Proyek', ENGINE.MPRO, ENGINE.MPRO_ALIAS), add: obj(ENGINE.ADD.sheet, ENGINE.ADD) };
}
function arMdHistSheet_(ss) {
  var sh = ss.getSheetByName(AR_MD_HIST_); if (sh) return sh;
  sh = ss.insertSheet(AR_MD_HIST_); var h = ['WAKTU', 'OLEH', 'JENIS', 'KUNCI', 'KOLOM', 'NILAI LAMA', 'NILAI BARU', 'SENSITIF', 'DASAR PERUBAHAN'];
  sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#14284B').setFontColor('#FFFFFF'); sh.setFrozenRows(1);
  sh.getRange(2, 1, 999, h.length).setNumberFormat('@'); [20, 26, 10, 26, 30, 30, 30, 9, 50].forEach(function (w, i) { sh.setColumnWidth(i + 1, w * 7); });
  return sh;
}
function arMdLog_(ss, rows) { if (!rows.length) return; var sh = arMdHistSheet_(ss), r = Math.max(sh.getLastRow() + 1, 2); roomFor_(sh, r + rows.length); sh.getRange(r, 1, rows.length, 9).setNumberFormat('@').setValues(rows.map(function (x) { return x.map(function (v) { return String(v === null || v === undefined ? '' : v).slice(0, 500); }); })); }
function apiMasterData() {
  arRequireMaker_();
  var ss = arSpreadsheet_(), d = arMdRead_(ss), hs = ss.getSheetByName(AR_MD_HIST_), hist = [];
  if (hs && hs.getLastRow() > 1) { var n = Math.min(hs.getLastRow() - 1, 400); hist = hs.getRange(hs.getLastRow() - n + 1, 1, n, 9).getDisplayValues().reverse(); }
  return plain_({ pel: d.pel, pro: d.pro, add: d.add, hist: hist, sens: ENGINE.MD_SENS });
}
/** tulis satu baris master (baru = baris setelah data terakhir) */
function arMdWrite_(sh, schema, alias, row, data) {
  var map = colMap_(sh, schema, alias), auto = ['add_tambah', 'add_kurang', 'nilai_ok'];
  if (schema.cols.some(function (c) { return !map[c[0]] && (c[0] in data) && auto.indexOf(c[0]) < 0; })) { ensureSchema_(sh, schema, alias, null); map = colMap_(sh, schema, alias); } // v2.9: kolom baru (mis. BATAS KREDIT) dibuat saat pertama dipakai
  if (!row) { row = lastDataRow_(sh, schema.firstRow, [map[schema.cols[0][0]]]) + 1; roomFor_(sh, row); }
  schema.cols.forEach(function (c) { var k = c[0], j = map[k]; if (!j || !(k in data) || auto.indexOf(k) >= 0) return; var v = data[k], rg = sh.getRange(row, j);
    if (AR_MD_TEXT_.indexOf(k) >= 0) { rg.setNumberFormat('@').setValue(arSafeCell_(String(v || ''))); }
    else if (k === 'tgl_kontrak' || k === 'tgl') { rg.setNumberFormat(NUMF.date).setValue(v ? toCell_(String(v), 'date') : ''); }
    else rg.setValue(typeof v === 'string' ? arSafeCell_(v) : v); });
  return row;
}
function arMdSaveCore_(ss, all, kind, orig, data, reason, who, now, log) {
  var schema = kind === 'PEL' ? ENGINE.MPEL : ENGINE.MPRO, alias = kind === 'PEL' ? null : ENGINE.MPRO_ALIAS, list = kind === 'PEL' ? all.pel : all.pro, key = kind === 'PEL' ? 'nama' : 'npp';
  data = JSON.parse(JSON.stringify(data || {})); if (data[key] !== undefined) data[key] = kind === 'PEL' ? String(data[key]).toUpperCase().trim() : String(data[key]).toUpperCase().replace(/\s+/g, ' ').trim();
  var ex = orig ? list.filter(function (r) { return String(r[key]).toUpperCase().trim() === String(orig).toUpperCase().trim(); })[0] : null;
  if (orig && !ex) throw new Error((kind === 'PEL' ? 'Customer ' : 'Project ') + orig + ' not found — reload the list.');
  if (ex) data[key] = ex[key]; // kunci tidak diganti lewat form (NPP: menu Ganti NPP proyek)
  var errs = ENGINE.masterChecks(kind, Object.assign({}, ex || {}, data), all, orig).filter(function (x) { return x.lvl === 'E'; });
  if (errs.length) throw new Error(errs.map(function (x) { return x.t + (x.m ? ' (' + x.m + ')' : ''); }).join('; '));
  var diff = ENGINE.masterDiff(kind, ex, data), sens = diff.filter(function (x) { return x.sens; });
  if (!diff.length) return { key: data[key], diff: [], sens: [] };
  if (sens.length && String(reason || '').trim().length < 5) throw new Error('Sensitive fields changed (' + sens.map(function (x) { return x.label; }).join(', ') + ') — enter the basis for the change (document / reason).');
  if (kind === 'PRO' && (data.pola_uraian !== undefined || data.pola_kw !== undefined)) { var mpx = ss.getSheetByName('Master Proyek'); if (!colMap_(mpx, ENGINE.MPRO, ENGINE.MPRO_ALIAS).pola_uraian) ensureSchema_(mpx, ENGINE.MPRO, ENGINE.MPRO_ALIAS, null); }
  arMdWrite_(ss.getSheetByName(kind === 'PEL' ? 'Master Pelanggan' : 'Master Proyek'), schema, alias, ex ? ex._row : 0, data);
  diff.forEach(function (x) { log.push([now, who, kind === 'PEL' ? 'PELANGGAN' : 'PROYEK', data[key], x.label, ex ? x.from : '(baru)', x.to, x.sens ? 'YA' : '', reason || '']); });
  return { key: data[key], diff: diff, sens: sens, baru: !ex };
}
function arMdFinish_(ss, log, sensRows, subject) {
  arMdLog_(ss, log); arBootBust_();
  if (sensRows.length) try { arNotify_(subject, ['Perubahan Master Data (kolom sensitif) melalui aplikasi:', ''].concat(sensRows.map(function (r) { return '• ' + r[2] + ' ' + r[3] + ' · ' + r[4] + ': "' + r[5] + '" → "' + r[6] + '" · dasar: ' + r[8] + ' · oleh ' + r[1]; }))); } catch (e) { }
  return arQueueAfterSave_().background;
}
/** simpan pelanggan / proyek. kind PEL | PRO · orig = kunci asal (kosong = baru) */
function apiMasterSave(kind, orig, data, reason) {
  var acc = arRequireMaker_(); if (['PEL', 'PRO'].indexOf(kind) < 0) throw new Error('Invalid master type.');
  var lock = arLock_(), out, log = []; lock.waitLock(28000);
  try { var ss = arSpreadsheet_(); out = arMdSaveCore_(ss, arMdRead_(ss), kind, orig, data, reason, acc.email, arNow_(), log); }
  finally { lock.releaseLock(); }
  out.background = arMdFinish_(arSpreadsheet_(), log, log.filter(function (r) { return r[7] === 'YA'; }), 'Perubahan Master Data — ' + (kind === 'PEL' ? 'pelanggan ' : 'proyek ') + out.key);
  return plain_(out);
}
/** wizard: pelanggan baru + proyek pertama sekaligus */
function apiMasterWizard(pel, pro) {
  var acc = arRequireMaker_(), lock = arLock_(), a, b, log = []; lock.waitLock(28000);
  try {
    var ss = arSpreadsheet_(), all = arMdRead_(ss), now = arNow_();
    // keduanya diperiksa dulu agar tidak tersimpan setengah
    var e1 = ENGINE.masterChecks('PEL', pel, all, '').filter(function (x) { return x.lvl === 'E'; });
    var all2 = { pel: all.pel.concat([{ nama: String(pel.nama || '').toUpperCase().trim(), npwp: pel.npwp }]), pro: all.pro };
    var e2 = ENGINE.masterChecks('PRO', Object.assign({}, pro, { pelanggan: String(pel.nama || '').toUpperCase().trim() }), all2, '').filter(function (x) { return x.lvl === 'E'; });
    if (e1.length || e2.length) throw new Error(e1.concat(e2).map(function (x) { return x.t + (x.m ? ' (' + x.m + ')' : ''); }).join('; '));
    a = arMdSaveCore_(ss, all, 'PEL', '', pel, '', acc.email, now, log);
    b = arMdSaveCore_(ss, all2, 'PRO', '', Object.assign({}, pro, { pelanggan: a.key }), '', acc.email, now, log);
  } finally { lock.releaseLock(); }
  var bg = arMdFinish_(arSpreadsheet_(), log, [], '');
  return plain_({ ok: true, pelanggan: a.key, npp: b.key, background: bg });
}
/** adendum kontrak baru (selalu sensitif: wajib dasar, dicatat & diemailkan) */
function apiMasterAddendum(npp, d, reason) {
  var acc = arRequireMaker_(); d = d || {};
  if (String(reason || '').trim().length < 5) throw new Error('Enter the basis for the addendum (document / reason).');
  if (!arValidDate_(String(d.tgl || ''))) throw new Error('Addendum effective date is invalid.');
  if (!(Number(d.nilai) > 0)) throw new Error('Addendum value must be greater than 0.');
  if (['TAMBAH', 'KURANG'].indexOf(String(d.jenis || '').toUpperCase()) < 0) throw new Error('Type must be TAMBAH or KURANG.');
  var lock = arLock_(), log = [], ke; lock.waitLock(28000);
  try {
    var ss = arSpreadsheet_(), all = arMdRead_(ss), np = String(npp || '').toUpperCase().trim();
    if (!all.pro.some(function (r) { return String(r.npp).toUpperCase().trim() === np; })) throw new Error('Project ' + npp + ' not found.');
    ke = all.add.filter(function (r) { return String(r.npp).toUpperCase().trim() === np; }).reduce(function (m, r) { return Math.max(m, Number(r.ke) || 0); }, 0) + 1;
    var sh = ss.getSheetByName(ENGINE.ADD.sheet); if (!sh) throw new Error('Sheet ' + ENGINE.ADD.sheet + ' not found.');
    arMdWrite_(sh, ENGINE.ADD, null, 0, { npp: np, ke: ke, nodok: String(d.nodok || ''), tgl: d.tgl, jenis: String(d.jenis).toUpperCase(), nilai: Math.round(Number(d.nilai)), catatan: String(d.catatan || '') });
    log.push([arNow_(), acc.email, 'ADENDUM', np, 'ADENDUM KE-' + ke + ' (' + String(d.jenis).toUpperCase() + ')', '', (String(d.jenis).toUpperCase() === 'KURANG' ? '−' : '+') + ENGINE.fmtRp(d.nilai) + ' · ' + (d.nodok || '') + ' · berlaku ' + d.tgl, 'YA', reason]);
  } finally { lock.releaseLock(); }
  var bg = arMdFinish_(arSpreadsheet_(), log, log, 'Adendum kontrak ' + npp + ' ke-' + ke);
  return plain_({ ok: true, ke: ke, background: bg });
}
