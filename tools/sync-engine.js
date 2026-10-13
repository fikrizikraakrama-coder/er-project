#!/usr/bin/env node
// Engine.gs adalah sumber tunggal. EngineClient.html dibuat ulang darinya: node tools/sync-engine.js
// Dengan --check: gagal (exit 1) bila keduanya berbeda — dipakai uji otomatis.
const fs = require('fs'), path = require('path');
const dir = path.join(__dirname, '..', 'src');
const engine = fs.readFileSync(path.join(dir, 'Engine.gs'), 'utf8');
const client = '<script>\n' + engine + '\n</script>\n';
const target = path.join(dir, 'EngineClient.html');
if (process.argv.includes('--check')) {
  if (fs.readFileSync(target, 'utf8') !== client) { console.error('EngineClient.html tidak sama dengan Engine.gs — jalankan: node tools/sync-engine.js'); process.exit(1); }
  console.log('EngineClient.html sinkron dengan Engine.gs');
} else { fs.writeFileSync(target, client); console.log('EngineClient.html diperbarui dari Engine.gs'); }
