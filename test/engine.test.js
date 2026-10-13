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
