// Uji klasifikasi revisi invoice (fungsi diambil langsung dari Code.gs)
const test = require('node:test'), assert = require('node:assert'), fs = require('fs'), path = require('path');
const src = fs.readFileSync(path.join(__dirname, '..', 'src', 'Code.gs'), 'utf8');
const a = src.indexOf('var REV_NILAI_'), b = src.indexOf('/** judul kolom tambahan Riwayat Revisi');
const revisiJenis_ = new Function(src.slice(a, b) + ';return revisiJenis_;')();
const form = { jenis: 'PROGRES', progres: '100.000.000', judul: 'Termin 2', bapTgl: '2026-09-20', bapProg: 'Termin II', alasan: 'x' };
const oo = { tgl: '2026-09-20', jenis: 'PROGRES', pelanggan: 'PT FAJARPUTERA DINASTI', npp: '264C 1301 JT', bruto: 1e8, umk: 0, retensi: 0, retpos: 'TIDAK ADA', dpp: 1e8, ppn: 11e6, tarif: 0.0265, dasar: 'SETELAH RETENSI', tagih: 111e6,
  invke: 'Termin 2', ttd: 'Yusqi Audah Firdaus', uraian: 'Termin 2', json: JSON.stringify(form) };
const p = (o, f) => Object.assign({}, oo, { json: undefined, form: Object.assign({}, form, f || {}) }, o || {});

test('typo nomor SPK di BAP → revisi BAP (Rev tetap, persetujuan tetap)', () => {
  const r = revisiJenis_(oo, p({}, { bapSpk: '012/FPD/II/2026', alasan: 'beda alasan' }));
  assert.strictEqual(r.jenis, 'BAP'); assert.match(r.diff[0], /Nomor SPK \(BAP\) \(sebelumnya dari Master Proyek\)/);
});
test('ubah uraian / penandatangan → revisi INVOICE (ajukan ulang)', () => {
  assert.strictEqual(revisiJenis_(oo, p({ uraian: 'Termin 2 (revisi)' }, { judul: 'Termin 2 (revisi)' })).jenis, 'INVOICE');
  assert.strictEqual(revisiJenis_(oo, p({ ttd: 'Anas Afifudin' })).jenis, 'INVOICE');
});
test('ubah DPP / tanggal → revisi NILAI', () => {
  assert.strictEqual(revisiJenis_(oo, p({ dpp: 1.1e8, ppn: 12.1e6 })).jenis, 'NILAI');
  assert.strictEqual(revisiJenis_(oo, p({ tgl: '2026-09-21' })).jenis, 'NILAI');
});
test('tanpa perubahan → kosong (ditolak)', () => { assert.strictEqual(revisiJenis_(oo, p()).jenis, ''); });
test('BAP + uraian sekaligus → INVOICE (yang lebih ketat menang)', () => {
  assert.strictEqual(revisiJenis_(oo, p({ uraian: 'X' }, { bapSpk: 'Y', judul: 'X' })).jenis, 'INVOICE');
});
