/**
 * AR CSU v2.8.0 — pembaruan kompatibel v2.7.1.
 * Sistem dirancang & dikembangkan oleh Fikri Zikra Akrama.
 * Tidak mengubah saldo pelanggan berdasarkan nama / menghapus transaksi lama.
 */
var AR28_LOCK_DEPTH_ = 0, AR28_NATIVE_LOCK_ = null, AR28_SS_ = null;
var AR28_MAKERS_ = ['fikrizikraakrama@gmail.com', 'zikraakramaa@gmail.com'];
function arSpreadsheet_() {
  if (AR28_SS_) return AR28_SS_;
  var id = PropertiesService.getScriptProperties().getProperty('AR_SOURCE_ID');
  if (id) return AR28_SS_ = SpreadsheetApp.openById(id);
  var ss = SpreadsheetApp.getActive();
  if (!ss || !ss.getSheetByName('Monitoring Invoice')) throw new Error('Jalankan Pasang pembaruan v2.8 dari spreadsheet utama terlebih dahulu.');
  return AR28_SS_ = ss;
}
// Script lock works for both web app and triggers. Depth handles internal sync -> rebuild.
function arLock_() {
  var held = false;
  return {
    waitLock: function(ms) { if (held) return; if (!AR28_LOCK_DEPTH_) { AR28_NATIVE_LOCK_ = LockService.getScriptLock(); AR28_NATIVE_LOCK_.waitLock(ms); } AR28_LOCK_DEPTH_++; held = true; },
    tryLock: function(ms) { if (held) return true; if (!AR28_LOCK_DEPTH_) { AR28_NATIVE_LOCK_ = LockService.getScriptLock(); if (!AR28_NATIVE_LOCK_.tryLock(ms)) return false; } AR28_LOCK_DEPTH_++; held = true; return true; },
    releaseLock: function() { if (!held) return; held = false; AR28_LOCK_DEPTH_--; if (!AR28_LOCK_DEPTH_ && AR28_NATIVE_LOCK_) { try { SpreadsheetApp.flush(); } finally { AR28_NATIVE_LOCK_.releaseLock(); AR28_NATIVE_LOCK_ = null; } } }
  };
}
// Keep legacy document properties readable; new prefixed properties also work in web apps.
function arProps_() {
  var s = PropertiesService.getScriptProperties(), d = null;
  try { d = PropertiesService.getDocumentProperties(); } catch (e) {}
  return {
    getProperty: function(k) { var v = s.getProperty('AR28_' + k); return v !== null ? v : d ? d.getProperty(k) : null; },
    setProperty: function(k,v) { s.setProperty('AR28_' + k, String(v)); if (d) d.setProperty(k,String(v)); return this; },
    deleteProperty: function(k) { s.deleteProperty('AR28_' + k); if (d) d.deleteProperty(k); return this; }
  };
}
function arParameters_() {
  var p = sheet_('Parameter');
  return ENGINE.readParams(rangeVals_(p,'A4:C120'), [], [], [], rangeVals_(p,'U4:Y40'), [], []);
}
function arRequireAccess_() {
  var email = email_(), p = arParameters_(), u = ENGINE.userBy(p,email);
  if (!email) throw new Error('Identitas Google tidak terbaca. Masuk dengan akun terdaftar; deployment harus dijalankan sebagai pengguna yang mengakses.');
  // delegasi aktif (bukan PEMBUAT — dijaga ENGINE.delegasiFor) hanya membuka mode Otorisasi Dokumen
  if (u && !/^(PEMBUAT|PENYETUJU)/.test(u.peran) && ENGINE.delegasiFor(p,email,null,arToday_())) return { email:email, maker:false, approver:false, delegate:true, role:u.peran };
  if (!u || !/^(PEMBUAT|PENYETUJU)/.test(u.peran)) throw new Error('Akun ini tidak memiliki akses ke aplikasi utama AR. Hubungi Pembuat.');
  return { email:email, maker:/^PEMBUAT/.test(u.peran), approver:/^PENYETUJU/.test(u.peran), role:u.peran };
}
function arRequireMaker_() { var a = arRequireAccess_(); if (!a.maker) throw new Error('Tindakan ini hanya untuk Pembuat. Section Head menggunakan Otorisasi Dokumen AR.'); return a; }
function arRequireOwner_() {
  var email = email_(), ss = arSpreadsheet_(), owner = ss.getOwner();
  if (!email || !(owner && owner.getEmail().toLowerCase() === email)) throw new Error('Pemasangan dan pengaturan trigger dilakukan oleh pemilik spreadsheet utama.');
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
  if (!p || !arValidDate_(p.tgl) || p.tgl > arToday_()) throw new Error('Tanggal invoice tidak valid / di masa depan.');
  ['dpp','ppn','bruto','umk','retensi'].forEach(function(k) { if (p[k] !== undefined && (!isFinite(Number(p[k])) || Number(p[k]) < 0)) throw new Error('Nilai ' + k + ' tidak valid.'); });
  if (!(Number(p.dpp) > 0) || !String(p.uraian || '').trim()) throw new Error('DPP dan uraian tagihan wajib diisi.');
}
function arValidatePayments_(rows, input) {
  checkPayRows_(rows);
  if (rows.length > 100) throw new Error('Maksimal 100 alokasi dalam satu penyimpanan.');
  var computed = ENGINE.compute(Object.assign({},input,{light:true}));
  var today = input.today, by = {}, p = computed.params, groups = {}, ctx = computed.ctx;
  ctx.invs.forEach(function(i) { by[ENGINE.normInv(i.nomor)] = i; });
  rows.forEach(function(r,i) {
    var n = Number(r.nominal), inv = by[ENGINE.normInv(r.inv)], pb = String(r.jenis_trx || '').toUpperCase() === 'PINDAH BUKU';
    if (!isFinite(n) || !Number.isSafeInteger(n) || Math.abs(n) > 1e15) throw new Error('Baris '+(i+1)+': nominal harus rupiah bulat yang valid.');
    if (!arValidDate_(r.tgl) || r.tgl > today) throw new Error('Baris '+(i+1)+': tanggal tidak valid / di masa depan.');
    if (!inv || inv._batal) throw new Error('Baris '+(i+1)+': invoice tidak ditemukan atau dibatalkan.');
    if (inv.tgl && r.tgl < inv.tgl) throw new Error('Tanggal pembayaran mendahului tanggal invoice.');
    if (inv._closedManual) throw new Error('Invoice sudah ditutup manual. Rekonsiliasikan penutupan sebelum mencatat penerimaan baru.');
    if (r.jenis_trx && !pb && String(r.jenis_trx).toUpperCase() !== 'TRANSFER') throw new Error('Jenis transaksi tidak dikenali.');
    if (r.alokasi && ENGINE.OPT.alokasi.indexOf(String(r.alokasi).toUpperCase()) < 0) throw new Error('Alokasi harus OTOMATIS, NETTO, atau PPN.');
    if (!pb && !ENGINE.rekOf(p,r.rekening)) throw new Error('Rekening tidak ada dalam Parameter.');
    if (r.transfer !== undefined && r.transfer !== '' && (!isFinite(Number(r.transfer)) || Number(r.transfer) < 0)) throw new Error('Nominal transfer tidak valid.');
    if (String(r.ref || '').toUpperCase() === 'MIGRASI') throw new Error('Referensi MIGRASI dicadangkan untuk data lama.');
    if (pb && !String(r.ref || '').trim()) throw new Error('Referensi pindah buku wajib diisi.');
    var key = (pb ? 'PB|' : 'TRF|') + String(r.ref || '_TANPA_REF_').trim(), g = groups[key] || (groups[key] = {pb:pb,rows:[],total:0});
    g.rows.push(r); g.total += n; r.nominal = n;
  });
  Object.keys(groups).forEach(function(k) {
    var g = groups[k], a = g.rows, first = a[0];
    if (a.some(function(r) { return r.tgl !== first.tgl || (!g.pb && ENGINE.rekOf(p,r.rekening).label !== ENGINE.rekOf(p,first.rekening).label); })) throw new Error('Satu referensi harus memakai tanggal dan rekening yang sama.');
    if (g.pb) {
      var invoices = {}; a.forEach(function(r) { invoices[r.inv]=1; });
      if (a.length < 2 || Object.keys(invoices).length < 2 || Math.abs(g.total) > 1 || !a.some(function(r){return r.nominal<0;}) || !a.some(function(r){return r.nominal>0;})) throw new Error('Setiap referensi pindah buku harus seimbang dan memiliki invoice asal serta tujuan berbeda.');
    } else {
      var totals = a.map(function(r) { return Number(r.transfer) || 0; }).filter(function(v){return v>0;});
      if (a.length > 1 && (!String(first.ref || '').trim() || !totals.length)) throw new Error('Transfer terpecah wajib memiliki referensi dan nominal total transfer.');
      if (totals.some(function(v) { return Math.abs(v-g.total)>1; })) throw new Error('Jumlah alokasi berbeda dari nominal transfer.');
    }
  });
}
function arPaymentCanonical_(rows) {
  var keys = ENGINE.LOG.cols.filter(function(c) { return colType_(c)==='manual'; }).map(function(c){return c[0];});
  return JSON.stringify(rows.map(function(r) { return keys.map(function(k) { var v=r[k]; if (['nominal','transfer'].indexOf(k)>=0) return Number(v)||0; return v === undefined || v === null ? '' : String(v).trim(); }); }));
}
function arSavePayments_(rows, requestId) {
  var access = arRequireMaker_(), lock = arLock_(), saved, npps = [], months = [];
  rows = JSON.parse(JSON.stringify(rows || []));
  rows.forEach(function(r) { r.inv = ENGINE.normInv(r.inv); });
  var hash = arDigest_(arPaymentCanonical_(rows)), token = requestId || 'legacy-' + hash;
  if (!/^[A-Za-z0-9_-]{12,100}$/.test(token)) throw new Error('ID permintaan tidak valid. Muat ulang aplikasi.');
  lock.waitLock(28000);
  try {
    var input = readAll_(), sh = sheet_('Log Penerimaan'), map = colMap_(sh,ENGINE.LOG), cols = ENGINE.LOG.cols;
    if (!map.request_id || !map.request_hash || !map.request_user) throw new Error('Pembaruan v2.8 belum dipasang pada spreadsheet.');
    var ix = function(k) { return ENGINE.colIndex(ENGINE.LOG,k); };
    var existing = input.log.filter(function(r) { return String(r[ix('request_id')] || '') === token; });
    if (existing.length) {
      if (existing.length !== rows.length || existing.some(function(r) { return r[ix('request_hash')] !== hash || r[ix('request_user')] !== access.email; })) throw new Error('ID penyimpanan sudah dipakai untuk data berbeda / tidak lengkap. Periksa Log Penerimaan; jangan ulangi sebagai transaksi baru.');
      var stored = existing.map(function(r){var o={};cols.forEach(function(c,j){o[c[0]]=r[j];});return o;});
      if (arDigest_(arPaymentCanonical_(stored)) !== hash) throw new Error('Transaksi pernah tersimpan tetapi kemudian diubah. Periksa Log Penerimaan.');
      saved = {ok:true,saved:true,duplicate:true,requestId:token,invoices:[],payments:[],key:ENGINE.payKey(stored[stored.length-1])};
    } else {
      if (arProps_().getProperty('VOID_'+token)) throw new Error('Permintaan lama sudah dibatalkan. Kirim dengan ID permintaan baru.');
      arValidatePayments_(rows,input);
      // Existing bank references identify a transfer, including all split allocations.
      var refs = {}; rows.forEach(function(r){if(r.ref) refs[String(r.ref).trim()]=true;});
      if (input.log.some(function(r) { return refs[String(r[ix('ref')] || '').trim()]; })) throw new Error('Referensi transaksi sudah tercatat. Periksa Log Penerimaan sebelum mencatat ulang.');
      if (input.log.some(function(r) { return r[ix('request_hash')] === hash; })) throw new Error('Data pembayaran yang sama sudah tercatat. Periksa transaksi sebelumnya.');
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
      sh.getRange(target,1,values.length,width).setValues(values);
      // Metadata and all allocations are in the same write; retries use those persisted rows.
      saved={ok:true,saved:true,duplicate:false,requestId:token,invoices:[],payments:[],key:ENGINE.payKey(rows[rows.length-1])};
    }
    months=rows.map(function(r){return r.tgl.slice(0,7);});
    var keys=rows.map(function(r){return r.inv;});
    var ctx=ENGINE.compute(Object.assign({},input,{light:true})).ctx;
    npps=ctx.invs.filter(function(i){return keys.indexOf(ENGINE.normInv(i.nomor))>=0;}).map(function(i){return i.npp;});
  } finally { lock.releaseLock(); }
  var post=arAfterSave_(npps,months); saved.reportPending=post.reportPending;saved.warnings=post.warnings;
  try { var b=ENGINE.bootstrap(compute_(),access.email); saved.invoices=b.invoices.filter(function(i){return keys.indexOf(ENGINE.normInv(i.nomor))>=0;}); saved.payments=b.payments.filter(function(p){return (rows[0].ref && p.ref===rows[0].ref) || (!rows[0].ref && p.key===saved.key);}); } catch(e) { saved.reportPending=true;saved.warnings.push('Ringkasan belum termuat: '+e.message); }
  return plain_(saved);
}
function arAfterSave_(npps,months) {
  var out={reportPending:false,warnings:[]};
  try { rebuildAll(); } catch(e) { out.reportPending=true;out.warnings.push('Laporan utama: '+e.message); }
  try {
    var a=paramKV_('FILE_AKUNTANSI_ID'),b=paramKV_('FILE_PE_ID');
    if ((a&&a.val)||(b&&b.val)) { var r=syncTurunan_({npps:npps,months:months,noPull:true}); if(r.err.length) {out.reportPending=true;out.warnings=out.warnings.concat(r.err);} }
  } catch(e) {out.reportPending=true;out.warnings.push('File turunan: '+e.message);}
  return out;
}
function apiRefreshReports() { arRequireMaker_(); return arAfterSave_(null,null); }
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
function arAlignBlock_(sh,b,vals,pw,fmts) {
  var headers={},special={},labels={};
  b.styles.forEach(function(s){for(var k=0;k<(s.n||1);k++){if(s.s==='hdr')headers[s.r+k]=true;if(s.s==='lbl')labels[s.r+k]=true;if(['title','note','stamp','signname','lbl'].indexOf(s.s)>=0)special[s.r+k]=true;}});
  var leftCols={},noteCols={};Object.keys(headers).forEach(function(r){(b.rows[r]||[]).forEach(function(v,j){if(arLeftLabel_(v))leftCols[physCol_(b,j)-b.c]=true;if(/CATATAN|KETERANGAN|TINDAK LANJUT/i.test(String(v)))noteCols[physCol_(b,j)-b.c]=true;});});
  var align=vals.map(function(row,r){return row.map(function(v,c){if(headers[r])return 'center';if(labels[r]&&[0,1,3,7,8].indexOf(c)>=0)return 'left';if(special[r]&&c===0)return 'left';if(noteCols[c])return 'right';if(leftCols[c])return 'left';if(c>0&&arLeftLabel_(row[c-1]))return 'left';return typeof v==='number' || /^Rp\s?[\d.]+/.test(String(v)) ? 'right':'center';});});
  sh.getRange(b.r,b.c,vals.length,pw).setHorizontalAlignments(align).setVerticalAlignment('middle');
  Object.keys(labels).forEach(function(r){[0,1,3,7,8].forEach(function(c){if(c<pw)sh.getRange(b.r+(+r),b.c+c,1,1).setWrap(true);});});
  Object.keys(Object.assign({},noteCols,leftCols)).forEach(function(c){sh.getRange(b.r,b.c+(+c),vals.length,1).setWrap(true);});
}
function arFormatSourceSheets_(ss) {
  var schemas=[ENGINE.REG,ENGINE.LOG,ENGINE.MPEL,ENGINE.MPRO,ENGINE.ADD,ENGINE.TAG,ENGINE.NOTE,ENGINE.APPR,ENGINE.MEMO,ENGINE.PEREF];
  schemas.forEach(function(sc){var sh=ss.getSheetByName(sc.sheet);if(!sh||sh.getLastRow()<sc.firstRow)return;var n=sh.getLastRow()-sc.firstRow+1,map=colMap_(sh,sc,sc===ENGINE.REG?ENGINE.REG_ALIAS:sc===ENGINE.MPRO?ENGINE.MPRO_ALIAS:null);
    sh.getRange(sc.firstRow,1,n,sh.getLastColumn()).setVerticalAlignment('middle');
    sc.cols.forEach(function(c){var j=map[c[0]];if(!j)return;var rg=sh.getRange(sc.firstRow,j,n,1),f=colFmt_(c),numeric=/^(money|money0|pct|int|num2)$/.test(f||'')||['nilai','nilai_ok','add_tambah','add_kurang','um','um_awal','ret_pct','tarif','top','adm','top_ret'].indexOf(c[0])>=0;
      rg.setHorizontalAlignment(arLeftLabel_(c[1])?'left':numeric||/catatan|ket$|alasan/.test(c[0])?'right':'center');
      if(arLeftLabel_(c[1])||/catatan|ket$|alasan/.test(c[0]))rg.setWrap(true);
    });
  });
  ['Parameter','Data LK (Tempel)','Tempel SAP','Riwayat Revisi'].forEach(function(n){var sh=ss.getSheetByName(n);if(sh&&sh.getLastRow())arFormatGeneric_(sh);});
}
function arFormatGeneric_(sh) {
  var rg=sh.getDataRange(),v=rg.getValues(),headers={},leftCols={};
  v.slice(0,8).forEach(function(row){row.forEach(function(x,j){if(arLeftLabel_(x))leftCols[j]=true;if(/^(CATATAN|KETERANGAN|ALASAN)/i.test(String(x)))headers[j]=true;});});
  rg.setVerticalAlignment('middle').setHorizontalAlignments(v.map(function(r,i){return r.map(function(x,j){return leftCols[j]?'left':typeof x==='number'||headers[j]?'right':i<2&&j===0?'left':'center';});}));
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
  function add(name,rows,title,type){var rep=res.reports[name];if(!rep)return;var bottom=rep.from;rep.blocks.forEach(function(b){bottom=Math.max(bottom,b.r+b.rows.length+2);});rep.blocks.push(arBlock_(bottom,rows,{1:'num2',2:'num2'}));rep.charts=(rep.charts||[]).concat([{title:title+' (Rp juta)',type:type,range:[bottom,1,rows.length,rows[0].length],row:bottom+rows.length+1,col:1,w:760,h:300,colors:['#137C8B','#1F3864','#D89532'],legend:rows[0].length>2?'bottom':'none'}]);}
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
  writeReport_(sh,{from:5,blocks:[arBlock_(5,[hdr].concat(ENGINE.reconciliationRows(res.ctx)),{4:'money0',5:'money0',6:'money0',7:'money0',8:'money0',9:'money0',10:'money0'})],wrapCols:[12,13],minW:{1:28,3:28,12:48}},'Rekonsiliasi Saldo');
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
    rs.getRange('A3').setValue('Posisi '+cutoff+' · diperbarui '+stamp+' WIB · berdasarkan data sumber terkini; bukan arsip master historis.');
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
    if(arProps_().getProperty(key)===digest){sh.getRange('A3').setValue('Posisi '+date+' · data diperiksa '+stamp+' WIB · berdasarkan data sumber terkini.');return;}
    sh.getRange('A1').setValue('KARTU PIUTANG PROYEK · '+n).setFontSize(13).setFontWeight('bold').setFontColor(C.navy);
    sh.getRange('A2').setValue(p.pelanggan+' · '+(p.nama_pe||p.singkat)+(p.selesai?' · PROYEK SELESAI':''));
    sh.getRange('A3').setValue('Posisi '+date+' · diperbarui '+stamp+' WIB · berdasarkan data sumber terkini.');
    writeReport_(sh,kb,'PE_'+n,{key:key});sh.setHiddenGridlines(true); sh.setFrozenRows(4);
    arProps_().setProperty(key,digest);count++;
  });
  // Never delete old project tabs: closed projects remain available for historical review.
  ss.getSheets().forEach(function(sheet){if(names.indexOf(sheet.getName())>=0){sheet.setColumnWidth(5,Math.max(190,sheet.getColumnWidth(5)));sheet.getRange('E4').setWrap(true);}});
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
    sh.getRange('A3').setValue('MEMPERBARUI — pilihan cutoff baru sedang dihitung…');
    arPushPE_(compute_(),null,name);
  }catch(err){sh.getRange('A3').setValue('BELUM DIPERBARUI — '+err.message);}finally{lock.releaseLock();}
}

// ---------------- Monthly snapshots: immutable revisions; AI is opt-in and aggregate-only.
function arMonthEnd_(month) {if(!/^\d{4}-(0[1-9]|1[0-2])$/.test(month))throw new Error('Bulan harus yyyy-mm.');var p=month.split('-');return new Date(Date.UTC(+p[0],+p[1],0)).toISOString().slice(0,10);}
function arMonthlyArchive_() {
  var sh=sheet_('Arsip Analisis AR');if(!sh)throw new Error('Pasang pembaruan v2.8 terlebih dahulu.');return sh;
}
function arMonthlyRows_() {var sh=arMonthlyArchive_();return sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,12).getValues():[];}
function arMonthlyNarrative_(m) {
  var pct=m.raFinal?Math.round(m.riAtasRA/m.raFinal*10000)/100:null,lines=[];
  lines.push('Posisi '+m.cutoff+': saldo buku pembantu Rp '+ENGINE.fmtRp(m.saldo)+', berubah Rp '+ENGINE.fmtRp(m.perubahan)+' dari akhir bulan sebelumnya.');
  lines.push('Penerimaan kas bulan ini Rp '+ENGINE.fmtRp(m.penerimaan)+'. RA final Rp '+ENGINE.fmtRp(m.raFinal)+'; RI atas RA Rp '+ENGINE.fmtRp(m.riAtasRA)+(pct===null?'; pencapaian belum dapat dihitung karena RA nol.':'; pencapaian '+pct+'%.'));
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
  var cutoff=arMonthEnd_(month);if(cutoff>=arToday_())throw new Error('Snapshot dibuat setelah bulan berakhir.');
  var lock=arLock_();lock.waitLock(28000);
  try {
    var rows=arMonthlyRows_(),old=rows.filter(function(r){return r[1]===month;});
    if(old.length&&!revise){arRenderMonthly_(old[old.length-1]);return {created:false,id:old[old.length-1][0]};}
    if(revise&&String(reason||'').trim().length<10)throw new Error('Alasan revisi minimal 10 karakter.');
    var input=readAll_(),m=ENGINE.monthlyMetrics(input,cutoff),n=arMonthlyNarrative_(m),ai=arAINarrative_(m),rev=old.length,id=month+'-Rev.'+('0'+rev).slice(-2),digest=arDigest_(JSON.stringify(m));
    var row=[id,month,rev,cutoff,arNow_(),email_()||Session.getEffectiveUser().getEmail(),digest,reason||'Snapshot otomatis akhir bulan',JSON.stringify(m),n,ai.text,ai.status],sh=arMonthlyArchive_();
    sh.getRange(sh.getLastRow()+1,1,1,12).setNumberFormat('@').setValues([row.map(arSafeCell_)]);
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
function arMonthlyAuto_(){if(!sheet_('Arsip Analisis AR'))return;var month=ENGINE.prevMonthEnd(arToday_()).slice(0,7);try{arCreateMonthly_(month,false,'');arProps_().setProperty('MONTHLY_ERROR','');}catch(e){arProps_().setProperty('MONTHLY_ERROR',arNow_()+' '+e.message);var sh=sheet_('Analisis Bulanan AR');if(sh)sh.getRange('A3').setValue('Analisis otomatis belum berhasil: '+e.message);}}
function arMonthlyMenu(){arRequireMaker_();var r=arCreateMonthly_(ENGINE.prevMonthEnd(arToday_()).slice(0,7),false,'');SpreadsheetApp.getUi().alert('Analisis tersedia: '+r.id);}
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
    var pr=arch.getProtections(SpreadsheetApp.ProtectionType.SHEET).filter(function(p){return p.getDescription()==='AR CSU v2.8 — arsip analisis';})[0]||arch.protect().setDescription('AR CSU v2.8 — arsip analisis');pr.removeEditors(pr.getEditors());pr.addEditors([email_()].concat(AR28_MAKERS_));if(pr.canDomainEdit())pr.setDomainEdit(false);
    var par=sheet_('Parameter'),users=par.getRange('U4:Y40').getValues();
    AR28_MAKERS_.forEach(function(mail){var found=-1,empty=-1;users.forEach(function(r,i){if(String(r[0]).toLowerCase()===mail)found=i;if(empty<0&&!r[0]&&!r[1])empty=i;});if(found>=0){users[found][4]='PEMBUAT';}else{if(empty<0)throw new Error('Tabel pengguna penuh. Sediakan baris kosong U4:Y40.');users[empty]=[mail,'Fikri Zikra Akrama','AR Officer','','PEMBUAT'];}});
    par.getRange('U4:Y40').setValues(users);
    arOrderSheets_(active);arFormatSourceSheets_(active);
    var managed=['onEditInstalled','dailyRebuild','syncTurunanAuto'];ScriptApp.getProjectTriggers().forEach(function(t){if(managed.indexOf(t.getHandlerFunction())>=0)ScriptApp.deleteTrigger(t);});
    ScriptApp.newTrigger('onEditInstalled').forSpreadsheet(active).onEdit().create();
    ScriptApp.newTrigger('dailyRebuild').timeBased().everyDays(1).atHour(6).inTimezone(TZ).create();
    ScriptApp.newTrigger('syncTurunanAuto').timeBased().everyHours(1).create();
    secret_();rebuildAll();
    arProps_().setProperty('V28_DONE','1');
    // Derived files and the first monthly snapshot run separately to avoid the execution limit.
    SpreadsheetApp.getUi().alert('Pembaruan v2.8 terpasang.\nBackup:\n'+backup.map(function(b){return b.url;}).join('\n')+'\n\nLanjutkan menu Aktifkan pilihan tanggal P&E, lalu Perbarui file Akuntansi & P&E, dan Buat / lihat analisis bulan lalu (satu per satu). Deploy versi web baru. Saldo pelanggan tidak dinolkan otomatis; cocokkan rincian PPh pada Rekonsiliasi Saldo.');
  }finally{lock.releaseLock();}
}
function apiResolvePayment(requestId) {
  var a=arRequireMaker_();if(!/^[A-Za-z0-9_-]{12,100}$/.test(requestId||''))throw new Error('ID tidak valid.');
  var l=arLock_();l.waitLock(28000);
  try {
    var rows=readSchema_(sheet_('Log Penerimaan'),ENGINE.LOG),ri=ENGINE.colIndex(ENGINE.LOG,'request_id'),ui=ENGINE.colIndex(ENGINE.LOG,'request_user'),found=rows.filter(function(r){return r[ri]===requestId;});
    if(found.some(function(r){return r[ui]!==a.email;}))throw new Error('Permintaan milik pengguna lain.');
    if(found.length)return {saved:true,count:found.length};
    // Cancel outstanding writes with this token under the same lock; a delayed RPC cannot commit later.
    arProps_().setProperty('VOID_'+requestId,a.email);
    return {saved:false,canceled:true};
  }finally{l.releaseLock();}
}
