const fs = require('fs');
const src = fs.readFileSync(process.argv[2] || require('path').resolve(__dirname, '..', '..', 'index.html'), 'utf8');
const blocks = src.match(/<script(?![^>]*\ssrc=)[^>]*>([\s\S]*?)<\/script>/g) || [];
let ok = 0;
for (const b of blocks) {
  const code = b.replace(/^<script[^>]*>/, '').replace(/<\/script>$/, '');
  if (!code.trim()) { ok++; continue; }
  try { new Function(code); ok++; }
  catch (e) { console.log('SYNTAX ERROR:', e.message); }
}
console.log('scripts parsed ok:', ok, '/', blocks.length);
