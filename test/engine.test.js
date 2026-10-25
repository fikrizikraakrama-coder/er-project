// Uji ENGINE di luar Google Sheets: node --test test/
const test = require('node:test'), assert = require('node:assert'), fs = require('fs'), path = require('path');
const ENGINE = new Function('module', fs.readFileSync(path.join(__dirname, '..', 'src', 'Engine.gs'), 'utf8') + ';return ENGINE;')({});

function row(schema, o) { return schema.cols.map(c => (o[c[0]] === undefined ? '' : o[c[0]])); }
const NOMOR = 'KU.02.02/WB-4C.0101/2026', NPP = '264C 1236 JT';
function fixture(inv) {
  return {
    today: '2026-10-13', ssTz: 'Asia/Jakarta',
    paramKV: [], paramSigners: [], paramTargets: [], paramAkun: [], paramCkpn: [], paramRek: [],
    paramUsers: [['maker@kantor.co.id', 'Pembuat', 'AR Officer', '', 'PEMBUAT'], ['sh@kantor.co.id', 'Section Head', 'Section Head', '', 'PENYETUJU'], ['pengganti@kantor.co.id', 'Pengganti', 'Manager', '', 'MENGETAHUI']],
    pelanggan: [row(ENGINE.MPEL, { nama: 'PT CONTOH', inv: 'PT CONTOH', aktif: 'YA', dasar_jt: 'SUBMIT' })],
    proyek: [row(ENGINE.MPRO, { npp: NPP, pelanggan: 'PT CONTOH', singkat: 'CONTOH', nilai: 1e9, nilai_ok: 1e9, tarif: 0.0265, top: 30, aktif: 'YA' }).concat([''])],
    adendum: [], tagbrut: [], notes: [], approvals: [], lk: null, selectors: {}, peref: [], log: [],
    register: [row(ENGINE.REG, Object.assign({ tgl: '2026-09-20', nomor: NOMOR, jenis: 'PROGRES', pelanggan: 'PT CONTOH', npp: NPP, dpp: 100e6, ppn: 11e6, tarif: 0.0265, top: 30, status_inv: 'AKTIF' }, inv))]
  };
}
const invOf = (res) => res.ctx.invs.filter(i => ENGINE.normInv(i.nomor) === ENGINE.normInv(NOMOR))[0];

test('EngineClient.html identik dengan Engine.gs', () => {
  const dir = path.join(__dirname, '..', 'src');
  assert.strictEqual(fs.readFileSync(path.join(dir, 'EngineClient.html'), 'utf8'), '<script>\n' + fs.readFileSync(path.join(dir, 'Engine.gs'), 'utf8') + '\n</script>\n');
});

test('delegasi tidak berlaku untuk akun PEMBUAT', () => {
  const P = ENGINE.readParams([['DELEGASI_SAMPAI', '2026-10-31'], ['DELEGASI_JENIS', 'PAY,VAL']], [], [], [], fixture().paramUsers, [], []);
  const at = (email, kind) => !!ENGINE.delegasiFor(Object.assign({}, P, { DELEGASI_EMAIL: email }), email, kind, '2026-10-13');
  assert.strictEqual(at('maker@kantor.co.id', 'PAY'), false);
  assert.strictEqual(at('pengganti@kantor.co.id', 'PAY'), true);
  assert.strictEqual(at('pengganti@kantor.co.id', 'INV'), false);
  assert.strictEqual(!!ENGINE.delegasiFor(Object.assign({}, P, { DELEGASI_EMAIL: 'pengganti@kantor.co.id' }), 'pengganti@kantor.co.id', 'PAY', '2026-11-01'), false, 'kedaluwarsa');
});

test('posisi historis: submit setelah cutoff → WAITING SUBMIT pada cutoff', () => {
  const input = fixture({ submit: '2026-10-05' });
  const now = invOf(ENGINE.compute(JSON.parse(JSON.stringify(input))));
  assert.notStrictEqual(now._res.status, 'WAITING SUBMIT', 'per hari ini sudah disubmit');
  const past = invOf(ENGINE.computeAt(input, '2026-09-30'));
  assert.strictEqual(past._res.status, 'WAITING SUBMIT');
  assert.ok(!past._res.jt, 'belum ada jatuh tempo per 30-09');
});

test('posisi historis: faktur & BA setelah cutoff dikosongkan, sebelum cutoff tetap', () => {
  const input = fixture({ submit: '2026-09-22', no_fp: '04002600000001', tgl_fp: '2026-10-02', ba_no: 'BA-1', ba_tgl: '2026-10-03' });
  const past = invOf(ENGINE.computeAt(input, '2026-09-30'));
  assert.strictEqual(past.submit, '2026-09-22');
  assert.ok(!past.tgl_fp && !past.no_fp && !past.ba_no);
});

test('contextAt di-cache per tanggal', () => {
  const res = ENGINE.compute(fixture({ submit: '2026-09-22' }));
  const a = ENGINE.kartuPlain(res.ctx, NPP, '2026-09-30'), b = ENGINE.kartuPlain(res.ctx, NPP, '2026-09-30');
  assert.deepStrictEqual(a, b);
  assert.strictEqual(Object.keys(res.ctx._atCache || {}).length, 1);
});

test('numForm sama dengan num() di aplikasi', () => {
  const idx = fs.readFileSync(path.join(__dirname, '..', 'src', 'Index.html'), 'utf8');
  const num = new Function(idx.match(/function num\(v\) \{.*\}/)[0] + ';return num;')();
  ['1.234.567', '1.234,5', '1234.5', '1,234,567', '', '0', 'Rp 25.000', '-1.000', 12.5, '12'].forEach(v => assert.strictEqual(ENGINE.numForm(v), num(v), String(v)));
});

test('invoiceInput = susunan input lama di aplikasi (server & aplikasi menghitung sama)', () => {
  const p = { template: 'IKK', nilai: 1e9, adendum: [{ ke: 1, tgl: '2026-08-01', jenis: 'TAMBAH', nilai: 1e8 }], nilai_ok: 1e9, um: 1e8, ret_pct: 0.05, ret_pos: 'AKHIR', tarif: 0.0265, dasar_pph: '', tampil_pph: 'YA', umLengkap: true, umSisa: 4e7 };
  const S = { tgl: '2026-10-01', jenis: 'PROGRES', progres: '250.000.000', items: [{ uraian: 'a', nilai: '1.000', bulan: '2' }], dpp: '', ppnOn: true };
  const num = ENGINE.numForm, k = ENGINE.kontrakAt({ nilai_awal: p.nilai, adds: p.adendum, nilai_ok: p.nilai_ok }, S.tgl);
  const old = { umSisa: (+p.umSisa || 0) + 5e6, template: p.template, jenis: S.jenis, nilaiKontrak: k.net, progres: num(S.progres), umDiterima: p.um, retPct: p.ret_pct, retPos: p.ret_pos, tarif: p.tarif,
    dasar: p.dasar_pph, tampilPPh: p.tampil_pph, items: S.items.map(x => ({ uraian: x.uraian, nilai: num(x.nilai), bulan: num(x.bulan) })), dpp: num(S.dpp), ppnOn: S.ppnOn };
  assert.deepStrictEqual(ENGINE.invoiceInput(p, S, 5e6), old);
  const c = ENGINE.calcInvoice(old);
  assert.ok(c.retPos === 'AKHIR' && c.ppn > 0 && c.retensi > 0, 'IKK: PPN dihitung dari DPP sebelum retensi');
});

test('bupotPendingRows: daftar hanya memuat PPh diakui tanpa nomor bukti potong', () => {
  const ctx = { today: '2026-10-13', pays: [
    { _inv: { pelanggan: 'A', npp: '264C 1236 JT', nomor: 'X1', _a: { tarif: 0.0265 } }, _pph: 2650000, tgl: '2026-09-13', ref: 'R1', bupot: '', idp: 'PNR-00001' },
    { _inv: { pelanggan: 'B', npp: '264C 1236 JL', nomor: 'X2', _a: { tarif: 0.02 } }, _pph: 100, tgl: '2026-10-01', bupot: 'BP-9', idp: 'PNR-00002' },
    { _inv: { pelanggan: 'C', npp: '264C 1236 JT', nomor: 'X3', _a: { tarif: 0.0265 } }, _pph: 0, tgl: '2026-10-01', idp: 'PNR-00003' }] };
  const r = ENGINE.bupotPendingRows(ctx);
  assert.strictEqual(r.length, 1);
  assert.strictEqual(r[0][2], 'X1'); assert.strictEqual(r[0][7], 30);
});

test('tutup buku otomatis tanggal 3 pukul 00.00 untuk bulan sebelumnya', () => {
  const t = (today, p) => ENGINE.periodeTutup(Object.assign({}, p || {}), today);
  assert.strictEqual(t('2026-10-02'), '2026-08-31', '2 Okt: September masih terbuka');
  assert.strictEqual(t('2026-10-03'), '2026-09-30', '3 Okt 00.00: September tertutup');
  assert.strictEqual(t('2026-10-13'), '2026-09-30');
  assert.strictEqual(t('2027-01-03'), '2026-12-31', 'lintas tahun');
  assert.strictEqual(t('2026-10-13', { PERIODE_DITUTUP_SAMPAI: '2026-10-10' }), '2026-10-10', 'manual hanya memajukan');
  assert.strictEqual(t('2026-10-13', { PERIODE_DITUTUP_SAMPAI: '2026-07-31' }), '2026-09-30', 'manual tidak memundurkan');
  assert.strictEqual(t('2026-10-04', { TUTUP_BUKU_TANGGAL: 5 }), '2026-08-31');
});

test('surat penagihan: hanya invoice lewat jatuh tempo, dikelompokkan per proyek; alert per pelanggan', () => {
  const input = fixture({ submit: '2026-09-01' }); // JT 01-10 → lewat 12 hari per 13-10
  input.register.push(row(ENGINE.REG, { tgl: '2026-10-01', nomor: 'KU.02.02/WB-4C.0102/2026', jenis: 'PROGRES', pelanggan: 'PT CONTOH', npp: NPP, dpp: 50e6, ppn: 5.5e6, tarif: 0.0265, top: 30, status_inv: 'AKTIF' })); // belum submit → tidak ikut
  const res = ENGINE.compute(input), d = ENGINE.suratData(res.ctx, 'PT CONTOH');
  assert.strictEqual(d.projects.length, 1);
  assert.deepStrictEqual(d.projects[0].invoices.map(i => i.nomor), [NOMOR]);
  const i = d.projects[0].invoices[0];
  assert.strictEqual(i.lewat, 12); assert.strictEqual(i.sisa, 111e6, 'sisa sebelum PPh = DPP + PPN'); assert.strictEqual(i.kas, 111e6 - 2650000, 'kas = sisa − PPh');
  let a = ENGINE.suratAlerts(res.ctx);
  assert.strictEqual(a.length, 1); assert.strictEqual(a[0].perlu, true);
  assert.strictEqual(a[0].ke, 1);
  const L = (o) => row(ENGINE.SURAT, Object.assign({ pelanggan: 'PT CONTOH', invoices: NOMOR }, o));
  const at = (surat) => { input.surat = surat; const r = ENGINE.compute(JSON.parse(JSON.stringify(input))); return { a: ENGINE.suratAlerts(r.ctx)[0], ews: ENGINE.kinerjaData(r.ctx).ews, d: ENGINE.suratData(r.ctx, 'PT CONTOH') }; };
  // surat ke-1 dengan tenggat belum lewat → menunggu
  let x = at([L({ nomor: 'S-1', tgl: '2026-10-12', tenggat: '2026-10-15', ke: 1 })]);
  assert.strictEqual(x.a.perlu, false); assert.strictEqual(x.a.tunggu, '2026-10-15');
  assert.strictEqual(x.d.letters[0].status, 'MENUNGGU TANGGAPAN');
  // tenggat lewat tanpa tanggapan → surat ke-2 disarankan + peringatan dini
  x = at([L({ nomor: 'S-1', tgl: '2026-10-05', tenggat: '2026-10-08', ke: 1 })]);
  assert.strictEqual(x.a.perlu, true); assert.strictEqual(x.a.ke, 2);
  assert.ok(x.ews.some(e => e.kat === 'TENGGAT SURAT TERLEWAT'));
  // tanpa tenggat → jeda 7 hari dari tanggal surat
  x = at([L({ nomor: 'S-1', tgl: '2026-10-08', ke: 1 })]); assert.strictEqual(x.a.perlu, false);
  // ditanggapi dengan janji bayar → ditahan sampai janji; janji lewat → surat berikutnya + JANJI BAYAR TERLEWAT
  x = at([L({ nomor: 'S-1', tgl: '2026-10-01', tenggat: '2026-10-04', ke: 1, tanggap_tgl: '2026-10-05', tanggap_jenis: 'EMAIL', janji: '2026-10-20' })]);
  assert.strictEqual(x.a.perlu, false); assert.strictEqual(x.d.letters[0].status, 'DITANGGAPI');
  x = at([L({ nomor: 'S-1', tgl: '2026-10-01', ke: 1, tanggap_tgl: '2026-10-02', janji: '2026-10-10' })]);
  assert.strictEqual(x.a.perlu, true); assert.strictEqual(x.a.janjiLewat, 1); assert.ok(x.ews.some(e => e.kat === 'JANJI BAYAR TERLEWAT'));
  // ditutup manual → invoice tidak lagi ditagih lewat surat
  x = at([L({ nomor: 'S-1', tgl: '2026-10-01', ke: 1, tutup: 'Sengketa volume, rekonsiliasi dengan PIC' })]);
  assert.strictEqual(x.a, undefined); assert.strictEqual(x.d.letters[0].status, 'DITUTUP');
  // surat ke-3 lewat tenggat → eskalasi, bukan surat ke-4
  x = at([L({ nomor: 'S-1', tgl: '2026-09-01', ke: 1 }), L({ nomor: 'S-2', tgl: '2026-09-15', ke: 2 }), L({ nomor: 'S-3', tgl: '2026-10-01', tenggat: '2026-10-04', ke: 3 })]);
  assert.strictEqual(x.a.perlu, false); assert.strictEqual(x.a.eskalasi, 1); assert.ok(x.ews.some(e => e.kat === 'ESKALASI PENAGIHAN'));
});

test('surat penagihan: lunas lewat penerimaan → SELESAI – LUNAS otomatis', () => {
  const input = fixture({ submit: '2026-09-01' });
  input.surat = [row(ENGINE.SURAT, { nomor: 'S-1', tgl: '2026-10-05', pelanggan: 'PT CONTOH', invoices: NOMOR, sisa: 111e6, ke: 1, tenggat: '2026-10-08' })];
  input.paramRek = [['BNI', '149394879', 'BANK_BNI']];
  input.log = [row(ENGINE.LOG, { tgl: '2026-10-10', inv: NOMOR, nominal: 111e6 - 2650000, rekening: 'BNI 149394879' })];
  const r = ENGINE.compute(input), d = ENGINE.suratData(r.ctx, 'PT CONTOH');
  assert.strictEqual(d.letters[0].status, 'SELESAI – LUNAS');
  assert.strictEqual(ENGINE.suratAlerts(r.ctx).length, 0);
  assert.ok(!ENGINE.kinerjaData(r.ctx).ews.some(e => /SURAT|JANJI/.test(e.kat)));
});

test('revisi & penggantian surat: nomor 212 → 212A → 212B; surat DIGANTI tidak aktif, invoice tetap ditagih', () => {
  assert.strictEqual(ENGINE.suratNoGanti('KU.03.03/WB-4C.212/2026', '', []), 'KU.03.03/WB-4C.212A/2026');
  assert.strictEqual(ENGINE.suratNoGanti('KU.03.03/WB-4C.212A/2026', '', []), 'KU.03.03/WB-4C.212B/2026');
  assert.strictEqual(ENGINE.suratNoGanti('KU.03.03/WB-4C.212/2026', '', ['KU.03.03/WB-4C.212A/2026']), 'KU.03.03/WB-4C.212B/2026');
  assert.strictEqual(ENGINE.suratNoGanti('KU.03.03/WB-4C.2026/2026', '', []), 'KU.03.03/WB-4C.2026A/2026', 'nomor sama dengan tahun');
  const input = fixture({ submit: '2026-09-01' });
  input.surat = [row(ENGINE.SURAT, { nomor: 'S-212', tgl: '2026-10-05', pelanggan: 'PT CONTOH', invoices: NOMOR, ke: 1, tenggat: '2026-10-20', ganti: 'S-212A' }),
                 row(ENGINE.SURAT, { nomor: 'S-212A', tgl: '2026-10-07', pelanggan: 'PT CONTOH', invoices: NOMOR, ke: 1, tenggat: '2026-10-20', menggantikan: 'S-212' })];
  const r = ENGINE.compute(input), d = ENGINE.suratData(r.ctx, 'PT CONTOH'), a = ENGINE.suratAlerts(r.ctx)[0];
  assert.strictEqual(d.letters[0].status, 'DIGANTI'); assert.strictEqual(d.letters[0].aktif, false);
  assert.strictEqual(d.letters[1].status, 'MENUNGGU TANGGAPAN');
  const st = d.projects[0].invoices[0].stage;
  assert.deepStrictEqual(st.chain.map(c => c.nomor), ['S-212A'], 'surat yang diganti tidak dihitung sebagai riwayat');
  assert.strictEqual(st.ke, 2); assert.strictEqual(a.perlu, false); assert.strictEqual(a.terakhir.nomor, 'S-212A');
});

test('tanggal tertukar di sheet (05/10 terbaca 10 Mei) ditandai di Cek Data', () => {
  const swap = ENGINE.compute(fixture({ submit: '2026-05-10' })).ctx.issues.filter(x => x[0] === 'TANGGAL');
  assert.strictEqual(swap.length, 1); assert.match(swap[0][3], /TGL SUBMIT DOKUMEN .* sebelum tanggal invoice/);
  const fut = ENGINE.compute(fixture({ submit: '2026-09-22', tt: '2026-12-09' })).ctx.issues.filter(x => x[0] === 'TANGGAL');
  assert.strictEqual(fut.length, 1); assert.match(fut[0][3], /TGL TANDA TERIMA di masa depan/);
  assert.strictEqual(ENGINE.compute(fixture({ submit: '2026-09-22' })).ctx.issues.filter(x => x[0] === 'TANGGAL').length, 0);
});

test('baris Log Penerimaan tanpa isian (sisa NO / ID otomatis) diabaikan', () => {
  const input = fixture({ submit: '2026-09-01' });
  input.log = [row(ENGINE.LOG, { no: 295, idp: 'PNR-00295' })];
  const r = ENGINE.compute(input);
  assert.strictEqual(r.logOut[0], null, 'kolom otomatis dikosongkan');
  assert.ok(!r.ctx.issues.some(x => x[0] === 'PEMBAYARAN'), 'tidak muncul di Cek Data');
});

test('IKK dasar PPh SEBELUM RETENSI: pelepasan retensi tanpa PPh (dan tanpa PPN bawaan)', () => {
  const c = ENGINE.calcInvoice({ template: 'IKK', jenis: 'RETENSI', dpp: 277209737, tarif: 0.0265, dasar: 'SEBELUM RETENSI', ppnOn: false });
  assert.strictEqual(c.pph, 0); assert.strictEqual(c.ppn, 0); assert.strictEqual(c.bayar, 277209737);
  const c2 = ENGINE.calcInvoice({ template: 'IKK', jenis: 'RETENSI', dpp: 277209737, tarif: 0.0265, dasar: 'SETELAH RETENSI', ppnOn: false });
  assert.ok(c2.pph > 0, 'dasar SETELAH RETENSI: PPh retensi dipotong saat pelepasan');
  const input = fixture();
  input.proyek[0] = row(ENGINE.MPRO, { npp: NPP, pelanggan: 'PT CONTOH', singkat: 'CONTOH', nilai: 1e10, nilai_ok: 1e10, tarif: 0.0265, top: 30, aktif: 'YA', template: 'IKK', dasar_pph: 'SEBELUM RETENSI', ret_pct: 0.05, ret_pos: 'AKHIR' }).concat(['']);
  input.register = [row(ENGINE.REG, { tgl: '2026-09-20', nomor: NOMOR, jenis: 'RETENSI', pelanggan: 'PT CONTOH', npp: NPP, dpp: 277209737, ppn: 0, tarif: 0.0265, dasar: 'SEBELUM RETENSI', template: 'IKK', status_inv: 'AKTIF' })];
  const i = ENGINE.compute(input).ctx.invs[0];
  assert.strictEqual(i._a.pph, 0); assert.strictEqual(i._a.kas, 277209737);
});

test('proyek ARSIP: invoice, penerimaan & tagbrut diabaikan, tanpa temuan Cek Data, baris dilaporkan untuk disembunyikan', () => {
  const ARS = '254C 0001 JT', NO2 = 'KU.02.02/WB-4C.0007/2024';
  const input = fixture();
  input.pelanggan.push(row(ENGINE.MPEL, { nama: 'PT LAMA', inv: 'PT LAMA', aktif: 'YA' }));
  input.proyek.push(row(ENGINE.MPRO, { npp: ARS, pelanggan: 'PT LAMA', singkat: 'LAMA', nilai: 5e8, ret_pct: 0.05, aktif: 'ARSIP' }).concat(['']));
  input.register.push(row(ENGINE.REG, { tgl: '2024-03-01', nomor: NO2, jenis: 'PROGRES', pelanggan: 'PT LAMA', npp: ARS, dpp: 1e8, ppn: 11e6 }));
  input.log = [row(ENGINE.LOG, { tgl: '2024-04-01', ref: 'MIGRASI', inv: NO2, nominal: 1e8 })];
  input.tagbrut = [row(ENGINE.TAG, { npp: ARS, nilai: 1e6 })];
  const r = ENGINE.compute(input);
  assert.ok(!r.ctx.invs.some(i => i.npp === ARS), 'invoice arsip tidak ikut dihitung');
  assert.strictEqual(r.ctx.pays.length, 0);
  assert.ok(!r.ctx.M.pro[ARS] && r.ctx.M.arsip[ARS]);
  assert.ok(!r.issues.some(x => JSON.stringify(x).indexOf(ARS) >= 0 || JSON.stringify(x).indexOf('0007/2024') >= 0 || JSON.stringify(x).indexOf('PT LAMA') >= 0), 'tidak ada temuan Cek Data');
  assert.deepStrictEqual(r.arsipRows.reg, [ENGINE.REG.firstRow + 1]);
  assert.deepStrictEqual(r.arsipRows.log, [ENGINE.LOG.firstRow]);
  assert.deepStrictEqual(r.arsipRows.tag, [ENGINE.TAG.firstRow]);
  assert.strictEqual(r.regOut[1].status_inv, 'ARSIP');
  assert.strictEqual(r.logOut[0].cek, 'ARSIP');
  assert.ok(invOf(r), 'invoice aktif tetap dihitung');
});

test('Overview: bulan berjalan & bulan lalu, ganti tahun, tidak melewati bulan berjalan', () => {
  const input = fixture({ submit: '2026-09-21' });
  input.register.push(row(ENGINE.REG, { tgl: '2026-10-02', nomor: 'KU.02.02/WB-4C.0102/2026', jenis: 'PROGRES', pelanggan: 'PT CONTOH', npp: NPP, dpp: 50e6, ppn: 5.5e6, tarif: 0.0265, top: 30, status_inv: 'AKTIF' }));
  input.log = [row(ENGINE.LOG, { idp: 'PNR-00001', tgl: '2026-10-05', inv: NOMOR, nominal: 108350000, rekening: 'BNI' })];
  input.light = true;
  const r = ENGINE.compute(input), o = ENGINE.overview(r.ctx, '');
  assert.strictEqual(o.ref, '2026-10'); assert.ok(o.cur); assert.strictEqual(o.asOf, '2026-10-13');
  assert.strictEqual(o.issued.n, 1); assert.strictEqual(o.issued.prevN, 1); assert.strictEqual(o.issued.total, 55.5e6);
  assert.strictEqual(o.receipts, 108350000);
  assert.strictEqual(o.trend.length, 12); assert.strictEqual(o.trend[0].ref, '2025-11'); assert.strictEqual(o.trend[11].ref, '2026-10');
  assert.strictEqual(o.aging.reduce((a, b) => a + b, 0), o.open);
  const sep = ENGINE.overview(r.ctx, '2026-09');
  assert.ok(!sep.cur); assert.strictEqual(sep.asOf, '2026-09-30'); assert.strictEqual(sep.issued.n, 1); assert.ok(sep.open > 0, 'saldo akhir September sebelum pembayaran Oktober');
  assert.strictEqual(ENGINE.overview(r.ctx, '2026-12').ref, '2026-10', 'bulan depan dibatasi ke bulan berjalan');
  const jan = ENGINE.overview(r.ctx, '2026-01'); assert.strictEqual(jan.trend[0].ref, '2025-02'); assert.strictEqual(jan.trend[11].ref, '2026-01');
});

test('File Pajak: rekap tahunan, buku penjualan per bulan (tanggal faktur), kelengkapan dokumen', () => {
  const input = fixture({ no_fp: '04002600398053546', tgl_fp: '2026-10-02', link_inv: 'https://drive.google.com/file/d/a' });
  input.register.push(row(ENGINE.REG, { tgl: '2026-10-05', nomor: 'KU.02.02/WB-4C.0102/2026', jenis: 'PROGRES', pelanggan: 'PT CONTOH', npp: NPP, dpp: 50e6, ppn: 5.5e6, tarif: 0.0265, top: 30, status_inv: 'AKTIF' }));
  input.light = true;
  const ctx = ENGINE.compute(input).ctx;
  const rk = ENGINE.pajakReport(ctx, 'REKAP', 2026), sum = rk.blocks[0].rows;
  assert.strictEqual(sum[10][1], 1); assert.strictEqual(sum[11][1], 1); assert.strictEqual(sum[14][1], 2, 'total tahun');
  const okt = ENGINE.pajakReport(ctx, 'BUKPEN', 2026, 10).blocks[0].rows, sep = ENGINE.pajakReport(ctx, 'BUKPEN', 2026, 9).blocks[0].rows;
  assert.strictEqual(okt.length, 2 + 2 + 1, 'faktur Oktober: invoice September (FP 2 Okt) + invoice Oktober tanpa FP');
  assert.strictEqual(sep.length, 2 + 1, 'September kosong karena faktur bertanggal Oktober');
  assert.ok(/HYPERLINK/.test(okt[2][15]), 'link berkas TTD');
  const dok = ENGINE.pajakReport(ctx, 'DOK').blocks[0].rows;
  assert.strictEqual(dok.length, 3); assert.strictEqual(dok[2][2], 'KU.02.02/WB-4C.0102/2026'); assert.ok(/berkas TTD/.test(dok[2][11]));
});

test('KOMPENSASI + PPh sesuai bupot: piutang berkurang tanpa masuk kas/inkaso; PPh diakui sebesar bupot; sisa = kurang bayar', () => {
  // tagihan 111 jt (Netto 100 + PPN 11), PPh sistem 2,65 jt; pelanggan: kas 85,35 jt + kompensasi klaim 20 jt, PPh bupot 2,5 jt, sisa 3,15 jt
  const input = fixture({ submit: '2026-09-21' });
  input.log = [
    row(ENGINE.LOG, { tgl: '2026-10-05', inv: NOMOR, nominal: 85350000, ref: 'BVF-1', rekening: 'MANDIRI 0060097045854' }),
    row(ENGINE.LOG, { tgl: '2026-10-05', inv: NOMOR, nominal: 20000000, ref: 'MEMO-15/X', jenis_trx: 'KOMPENSASI', ket: 'Klaim solar', lunas: '✓', pph_bupot: 2500000, bupot: 'BP-1' })];
  const res = ENGINE.compute(input), i = invOf(res), kom = res.ctx.pays[1];
  assert.deepStrictEqual(kom._cek, []);
  assert.strictEqual(kom._pph, 2500000);
  assert.strictEqual(i._res.pot_pph, 2500000);
  assert.strictEqual(i._res.sisa, 3150000);
  assert.strictEqual(i._res.status, 'KURANG BAYAR 3.150.000');
  const k = ENGINE.kartuPlain(res.ctx, NPP, res.ctx.today);
  assert.strictEqual(k.T.kas, 85350000, 'kompensasi bukan penerimaan kas');
  assert.strictEqual(k.T.pot, 2500000);
  const rows = k.groups[0].rows.map(r => r.uraian).join(' | ');
  assert.match(rows, /Dikompensasi dengan utang ke pelanggan — Klaim solar/);
  assert.match(rows, /sesuai bukti potong/);
  // jurnal: utang (D) / piutang (K), tanpa bank
  const pk = ENGINE.payPackage(res.ctx, 'REF MEMO-15/X'), kode = pk.jurnal[0].lines.map(l => l.nama + ':' + l.debit + ':' + l.kredit);
  assert.strictEqual(pk.rows[0].rekening, 'KOMPENSASI');
  assert.ok(kode.includes('UTANG_KOMPENSASI:20000000:0'), kode.join(' '));
  assert.ok(!kode.some(x => /^BANK/.test(x))); assert.strictEqual(pk.jurnal[0].tcode, 'FB01');
  // akun lawan belum diisi → temuan Parameter
  assert.ok(res.issues.some(x => /UTANG_KOMPENSASI/.test(x[2])));
  // kompensasi wajib REF
  const in2 = fixture(); in2.log = [row(ENGINE.LOG, { tgl: '2026-10-05', inv: NOMOR, nominal: 1000, jenis_trx: 'KOMPENSASI' })];
  assert.match(ENGINE.compute(in2).ctx.pays[0]._cek.join(), /KOMPENSASI WAJIB PUNYA REF/);
});

test('PPh sesuai bupot: kas + PPh bupot = tagihan → LUNAS walau PPh beda dari hitungan sistem; tanpa isian tetap otomatis', () => {
  const input = fixture({ submit: '2026-09-21' });
  input.log = [row(ENGINE.LOG, { tgl: '2026-10-05', inv: NOMOR, nominal: 108500000, ref: 'BVF-2', rekening: 'MANDIRI 0060097045854', pph_bupot: 2500000 })];
  let i = invOf(ENGINE.compute(input));
  assert.strictEqual(i._res.status, 'LUNAS'); assert.strictEqual(i._res.pot_pph, 2500000); assert.strictEqual(i._res.adm, 0);
  input.log = [row(ENGINE.LOG, { tgl: '2026-10-05', inv: NOMOR, nominal: 108350000, ref: 'BVF-3', rekening: 'MANDIRI 0060097045854' })];
  i = invOf(ENGINE.compute(input));
  assert.strictEqual(i._res.status, 'LUNAS'); assert.strictEqual(i._res.pot_pph, 2650000);
});

test('ADJUSTMENT DPP (dasar PPN saja): DPP faktur tetap, piutang & PPh dari DPP − adjustment; kas + kompensasi + PPh = LUNAS (kasus WRK 0141)', () => {
  const input = fixture({ dpp: 8069185245, dpp_adj: 365569857, ppn: 887610377, retensi: 444661870, retpos: 'AKHIR', bruto: 8893237398, umk: 1189622010, submit: '2026-08-06' });
  input.log = [
    row(ENGINE.LOG, { tgl: '2026-10-10', inv: NOMOR, nominal: 5766264127, ref: '086730', rekening: 'MANDIRI 0060097045854' }),
    row(ENGINE.LOG, { tgl: '2026-10-10', inv: NOMOR, nominal: 2187937500, ref: 'MEMO DIR 15-10-2026', jenis_trx: 'KOMPENSASI', ket: 'Klaim solar' })];
  const res = ENGINE.compute(input), i = invOf(res), a = i._a;
  assert.strictEqual(a.dpp, 8069185245, 'DPP faktur tidak berubah');
  assert.strictEqual(a.netto, 7258953518, 'piutang Netto = LK Akuntansi');
  assert.strictEqual(a.pph, 192362268, 'PPh = BAPP');
  assert.strictEqual(i._res.status, 'LUNAS'); assert.strictEqual(i._res.pot_pph, 192362268);
  assert.ok(!res.issues.some(x => x[2] === NOMOR && /tidak sama dengan DPP/.test(x[3])), 'bruto − UM cocok dengan DPP − adjustment');
  const k = ENGINE.kartuPlain(res.ctx, NPP, res.ctx.today);
  assert.match(k.groups[0].rows[0].uraian, /adjustment DPP 365\.569\.857 hanya dasar PPN/);
  assert.strictEqual(k.T.kas, 5766264127);
});
