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
