const src = require('fs').readdirSync(require('path').resolve(__dirname, '..', '..', 'js')).sort().map(f => require('fs').readFileSync(require('path').resolve(__dirname, '..', '..', 'js', f), 'utf8')).join('\n');
const a = src.indexOf('function mdToHtml'), b = src.indexOf('function addBubble');
const mdToHtml = new Function(src.slice(a, b) + '; return mdToHtml;')();
const out = mdToHtml('**Pedro em 2026: R$ 8.162,18** (204)\n\nMês a mês:\n- Jan: R$ 766,10\n- **Ago: R$ 1.621,39** (maior)\n\nAté *setembro* <x>');
console.log(out);
const ok = out.includes('<b>Pedro em 2026: R$ 8.162,18</b>') && out.includes('<ul') && out.includes('<li') && out.includes('<i>setembro</i>') && out.includes('&lt;x&gt;') && !out.includes('**');
console.log(ok ? 'MD OK' : 'MD FAIL');
process.exit(ok ? 0 : 1);
