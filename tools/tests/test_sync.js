// Testa pushToFirebase/_snapshotToRemote/_isLegacyShape extraídos do index.html com um Firebase falso
const fs = require('fs');
const src = fs.readdirSync('C:/Users/Dell/financas-casal/js').sort().map(f => fs.readFileSync('C:/Users/Dell/financas-casal/js/' + f, 'utf8')).join('\n');
const start = src.indexOf('// ─── SYNC GRANULAR');
const end = src.indexOf('// Merge dois arrays por ID');
const syncCode = src.slice(start, end);
const fbToArraySrc = src.slice(src.indexOf('function fbToArray'), src.indexOf('let firebaseReady = false;'));

// Fake Firebase: guarda uma árvore e aplica update() multi-path como o real (null = apaga, obj vazio = apaga)
const store = { data: null };
const writes = [];
function applyUpdate(root, upd) {
  for (const [path, val] of Object.entries(upd)) {
    const parts = path.split('/');
    let node = root;
    for (let i = 0; i < parts.length - 1; i++) {
      if (!node[parts[i]] || typeof node[parts[i]] !== 'object') node[parts[i]] = {};
      node = node[parts[i]];
    }
    const last = parts[parts.length - 1];
    if (val === null || (typeof val === 'object' && !Object.keys(val).length)) delete node[last];
    else node[last] = JSON.parse(JSON.stringify(val));
  }
}
// Firebase devolve chaves ordenadas: simulamos
function sortKeys(v) {
  if (Array.isArray(v)) return v.map(sortKeys);
  if (v && typeof v === 'object') { const o = {}; Object.keys(v).sort().forEach(k => o[k] = sortKeys(v[k])); return o; }
  return v;
}
var db = { ref: (p) => ({ update: (upd) => { writes.push(upd); if (!store.data) store.data = {}; applyUpdate(store.data, upd); return Promise.resolve(); } }) };

let S, firebaseAuthReady = true, firebaseReady = true;
let statuses = [];
function updateSyncStatus(s) { statuses.push(s); }
let processed = [];
function _onFirebaseValue(snap) { processed.push(snap); }

eval((fbToArraySrc + syncCode).replace(/^(const|let) /gm, 'var '));

const tx = (id, desc, extra = {}) => ({ id, desc, amount: 10, date: '2026-09-01', parcela: null, ...extra });
let fails = 0;
function check(name, cond, info) { if (cond) console.log('  ok  ', name); else { fails++; console.log('  FAIL', name, info !== undefined ? JSON.stringify(info).slice(0, 300) : ''); } }

(async () => {
  // ── 1. Legacy: Firebase tem arrays posicionais ──
  console.log('1) migração do formato array → {id: item}');
  store.data = {
    transactions: [tx('t1', 'A'), tx('t2', 'B')],
    accounts: [{ id: 'card_nubank', name: 'Nubank' }],
    deletedIds: [{ id: 'old', collection: 'transactions', deletedAt: '2026-09-20T00:00:00.000Z' }],
    settings: { u1: 'Paulo', u2: 'Thayse' },
    budget: { 'Casa': 100 },
  };
  let fbData = sortKeys(store.data);
  check('detecta legacy transactions', _isLegacyShape(fbData.transactions));
  check('detecta legacy accounts', _isLegacyShape(fbData.accounts));
  check('não detecta legacy em ausente', !_isLegacyShape(fbData.debts));
  ID_COLLECTIONS.forEach(c => { if (_isLegacyShape(fbData[c])) _migrateColls.add(c); });
  _remote = _snapshotToRemote(fbData);
  S = { transactions: fbToArray(fbData.transactions), accounts: fbToArray(fbData.accounts), debts: [], investments: [],
        deletedIds: fbToArray(fbData.deletedIds), loveMessages: [], settings: { ...fbData.settings }, budget: fbData.budget,
        catOrcGroup: {}, chatHistory: [], customCats: null };
  writes.length = 0;
  pushToFirebase(); await new Promise(r => setTimeout(r, 0));
  check('1 escrita', writes.length === 1, writes);
  check('reescreveu só as 3 coleções legacy', Object.keys(writes[0]).sort().join() === 'accounts,deletedIds,transactions', Object.keys(writes[0]));
  check('transactions agora {id: item}', store.data.transactions.t1 && store.data.transactions.t2 && !store.data.transactions[0]);
  check('settings/budget intactos (não reescritos)', store.data.settings.u1 === 'Paulo' && store.data.budget.Casa === 100);
  check('_migrateColls limpo', _migrateColls.size === 0);

  // ── 2. Snapshot novo já no formato por id: nenhum legacy, save sem mudança = 0 escritas ──
  console.log('2) novo snapshot, nada mudou → 0 escritas (ordem de chaves diferente não engana)');
  fbData = sortKeys(store.data);
  check('não é mais legacy', !ID_COLLECTIONS.some(c => _isLegacyShape(fbData[c])));
  _remote = _snapshotToRemote(fbData);
  // S local tem chaves em outra ordem + parcela:null (Firebase descarta null)
  S.transactions = fbToArray(fbData.transactions).map(t => ({ parcela: null, desc: t.desc, id: t.id, amount: t.amount, date: t.date, tags: [] }));
  writes.length = 0;
  pushToFirebase(); await new Promise(r => setTimeout(r, 0));
  check('0 escritas', writes.length === 0, writes);

  // ── 3. Editar 1 transação → 1 path ──
  console.log('3) editar uma transação');
  S.transactions.find(t => t.id === 't2').amount = 99;
  writes.length = 0;
  pushToFirebase(); await new Promise(r => setTimeout(r, 0));
  check('1 escrita com 1 path', writes.length === 1 && Object.keys(writes[0]).join() === 'transactions/t2', writes);
  check('valor gravado', store.data.transactions.t2.amount === 99);
  check('t1 intacto', store.data.transactions.t1.desc === 'A');

  // ── 4. Adicionar + deletar com tombstone ──
  console.log('4) adicionar t3, deletar t1 (com tombstone)');
  S.transactions.push(tx('t3', 'C'));
  S.transactions = S.transactions.filter(t => t.id !== 't1');
  S.deletedIds.push({ id: 't1', collection: 'transactions', deletedAt: new Date().toISOString() });
  writes.length = 0;
  pushToFirebase(); await new Promise(r => setTimeout(r, 0));
  const paths = Object.keys(writes[0] || {}).sort();
  check('paths = t3 novo, t1 null, tombstone t1', paths.join() === 'deletedIds/t1,transactions/t1,transactions/t3', paths);
  check('t1 null', writes[0]['transactions/t1'] === null);
  check('store sem t1, com t3', !store.data.transactions.t1 && store.data.transactions.t3.desc === 'C');

  // ── 5. Mudança em campo-chave (settings) ──
  console.log('5) settings.usdRate muda → só settings');
  S.settings.usdRate = 5.4;
  writes.length = 0;
  pushToFirebase(); await new Promise(r => setTimeout(r, 0));
  check('só settings', Object.keys(writes[0] || {}).join() === 'settings', writes);
  check('settings completo gravado', store.data.settings.u1 === 'Paulo' && store.data.settings.usdRate === 5.4);

  // ── 6. Snapshot durante escrita é enfileirado e processado depois ──
  console.log('6) snapshot enfileirado durante escrita');
  S.transactions[0].desc = 'X';
  writes.length = 0; processed.length = 0;
  let resolveWrite;
  const slowDb = { ref: () => ({ update: (u) => new Promise(r => { resolveWrite = r; }) }) };
  const realDb = db;
  // troca db temporariamente
  db = slowDb;
  pushToFirebase();
  check('_inflight = 1', _inflight === 1);
  const fakeSnap = { val: () => ({}) };
  // simula o listener
  if (_inflight > 0) { _pendingSnapshot = fakeSnap; }
  check('snapshot ficou pendente', _pendingSnapshot === fakeSnap);
  resolveWrite(); await new Promise(r => setTimeout(r, 0)); await new Promise(r => setTimeout(r, 0));
  check('_inflight = 0 e snapshot processado', _inflight === 0 && processed.length === 1 && _pendingSnapshot === null);
  db = realDb;

  // ── 7. _remote null (falha anterior) → escreve tudo por id ──
  console.log('7) _remote null → escrita completa por id');
  _remote = null; writes.length = 0;
  pushToFirebase(); await new Promise(r => setTimeout(r, 0));
  check('todas as chaves', Object.keys(writes[0]).sort().join() === [...ID_COLLECTIONS, ...KEY_FIELDS].sort().join(), Object.keys(writes[0]));
  check('transactions por id', writes[0].transactions.t2 && writes[0].transactions.t3 && !Array.isArray(writes[0].transactions));

  // ── 8. Coleção nova (debts ausente no remoto) recebe item ──
  console.log('8) primeiro item numa coleção ausente');
  fbData = sortKeys(store.data); _remote = _snapshotToRemote(fbData);
  S.debts = [{ id: 'd1', name: 'Carro' }]; writes.length = 0;
  pushToFirebase(); await new Promise(r => setTimeout(r, 0));
  check('debts/d1', Object.keys(writes[0]).join() === 'debts/d1', writes);

  console.log(fails ? `\n${fails} FALHA(S)` : '\nTUDO OK');
  process.exit(fails ? 1 : 0);
})();
