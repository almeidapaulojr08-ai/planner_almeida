// Roda todos os testes do FinançasCasal. Uso: node tools/tests/run_all.js
const { spawnSync } = require('child_process');
const path = require('path');
const tests = ['check_syntax.js', 'check_modules.js', 'test_sync.js', 'test_md.js'];
let fail = 0;
for (const t of tests) {
  const r = spawnSync(process.execPath, [path.join(__dirname, t)], { encoding: 'utf8' });
  const last = (r.stdout || '').trim().split('\n').pop();
  console.log((r.status === 0 ? 'PASS ' : 'FAIL ') + t + ' — ' + last + (r.status !== 0 ? '\n' + (r.stderr || '') : ''));
  if (r.status !== 0) fail++;
}
console.log(fail ? `\n${fail} teste(s) falharam` : '\nTODOS OS TESTES PASSARAM');
process.exit(fail ? 1 : 0);
