// Valida a versão modular: cada js/*.js parseia sozinho; index.html referencia todos em ordem;
// nenhum módulo executa no topo algo definido só em módulo posterior (checagem estática simples).
const fs = require('fs'), path = require('path');
const ROOT = 'C:/Users/Dell/financas-casal';
const html = fs.readFileSync(path.join(ROOT, 'index.html'), 'utf8');
const refs = [...html.matchAll(/<script src="js\/([^"?]+)\?v=(\d+)"><\/script>/g)].map(m => m[1]);
const files = fs.readdirSync(path.join(ROOT, 'js')).filter(f => f.endsWith('.js')).sort();
let ok = true;
if (refs.join() !== files.join()) { console.log('REFS != FILES', refs, files); ok = false; }
const stamps = new Set([...html.matchAll(/\?v=(\d+)"/g)].map(m => m[1]));
if (stamps.size !== 1) { console.log('carimbos diferentes:', [...stamps]); ok = false; }
// parse individual
const defined = new Map(); // nome → índice do arquivo
files.forEach((f, i) => {
  const src = fs.readFileSync(path.join(ROOT, 'js', f), 'utf8');
  try { new Function(src); } catch (e) { console.log('SYNTAX', f, e.message); ok = false; }
  for (const m of src.matchAll(/^(?:async\s+)?function\s+([A-Za-z_$][\w$]*)/gm)) if (!defined.has(m[1])) defined.set(m[1], i);
  for (const m of src.matchAll(/^(?:const|let|var)\s+([A-Za-z_$][\w$]*)/gm)) if (!defined.has(m[1])) defined.set(m[1], i);
});
// top-level statements que chamam algo de módulo posterior (heurística: linhas no nível 0 fora de function/const)
files.forEach((f, i) => {
  const src = fs.readFileSync(path.join(ROOT, 'js', f), 'utf8');
  let depth = 0, inStr = null;
  const lines = src.split('\n');
  lines.forEach((line, ln) => {
    const trimmed = line.trim();
    const top = depth === 0 && trimmed && !/^(function|async function|const|let|var|\/\/|\/\*|\*|\})/.test(trimmed);
    if (top) {
      for (const m of trimmed.matchAll(/\b([A-Za-z_$][\w$]*)\s*\(/g)) {
        const name = m[1];
        if (defined.has(name) && defined.get(name) > i) { console.log(`ORDEM: ${f}:${ln + 1} chama ${name}() definida em ${files[defined.get(name)]}`); ok = false; }
      }
    }
    // contagem simples de chaves (ignora strings/regex de forma grosseira)
    for (const ch of line.replace(/'[^']*'|"[^"]*"|`[^`]*`/g, '')) { if (ch === '{') depth++; else if (ch === '}') depth--; }
  });
});
console.log(ok ? `MODULOS OK (${files.length} arquivos, ${defined.size} símbolos)` : 'MODULOS FAIL');
process.exit(ok ? 0 : 1);
