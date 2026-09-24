// FinançasCasal — 01-firebase-sync.js
// Módulo carregado em ordem por index.html (escopo global compartilhado). Não use import/export.

// ─── FIREBASE INIT ────────────────────────────────────────────────────────────
const firebaseConfig = {
  apiKey: "AIzaSyDzAAKHfTpNq16nzWXdZdJ8c4nRYswOrEw",
  authDomain: "almeida-wosniak-dre.firebaseapp.com",
  databaseURL: "https://almeida-wosniak-dre-default-rtdb.firebaseio.com",
  projectId: "almeida-wosniak-dre",
  storageBucket: "almeida-wosniak-dre.firebasestorage.app",
  messagingSenderId: "486795966461",
  appId: "1:486795966461:web:c7b8363c998b772f738b94"
};
let db = null;
try {
  firebase.initializeApp(firebaseConfig);
  db = firebase.database();
} catch(e) {
  console.warn('Firebase init failed (offline mode):', e);
}

function fbToArray(obj) {
  if (!obj) return [];
  if (Array.isArray(obj)) return obj;
  return Object.values(obj);
}

let firebaseReady = false;
let firebaseAuthReady = false;

// ─── TEMA (claro / escuro) ────────────────────────────────────────────────────
function cssVar(name) { return getComputedStyle(document.documentElement).getPropertyValue(name).trim(); }
function currentTheme() { return document.documentElement.getAttribute('data-theme') === 'dark' ? 'dark' : 'light'; }
function applyTheme(mode) {                       // mode: 'light' | 'dark' | 'auto'
  try { localStorage.setItem('fincasal_theme', mode); } catch (e) {}
  const dark = mode === 'dark' || (mode === 'auto' && matchMedia('(prefers-color-scheme: dark)').matches);
  if (dark) document.documentElement.setAttribute('data-theme', 'dark');
  else document.documentElement.removeAttribute('data-theme');
  const meta = document.querySelector('meta[name="theme-color"]');
  if (meta) meta.content = dark ? '#121a2b' : '#6366f1';
  const lbl = document.getElementById('theme-label');
  const tg = document.getElementById('theme-toggle');
  if (lbl) lbl.textContent = dark ? 'Modo claro' : 'Modo escuro';
  if (tg) tg.querySelector('use').setAttribute('href', dark ? '#i-sun' : '#i-moon');
  if (window.Chart) {                             // canvas não lê CSS vars: ajusta os defaults
    Chart.defaults.color = cssVar('--text-3');
    Chart.defaults.borderColor = cssVar('--chart-grid');
  }
  // re-renderiza a página atual (gráficos precisam redesenhar com as cores novas)
  const current = document.querySelector('section[id^="page-"]:not([style*="display:none"])');
  if (current && typeof goto === 'function' && firebaseReady) goto(current.id.replace('page-', ''));
}
function toggleTheme() { applyTheme(currentTheme() === 'dark' ? 'light' : 'dark'); }

function updateSyncStatus(status) {
  const el = document.getElementById('sync-status');
  if (!el) return;
  const icon = (n, color) => `<svg class="ico-sm" style="stroke:${color};fill:none;stroke-width:2;stroke-linecap:round;stroke-linejoin:round;width:13px;height:13px;vertical-align:-2px;"><use href="#i-${n}"/></svg>`;
  if (status === 'ok') { el.innerHTML = icon('cloud', '#10b981') + ' sincronizado'; el.title = 'Firebase sincronizado'; el.dataset.state = 'ok'; }
  else if (status === 'error') { el.innerHTML = icon('cloud-off', '#e11d48') + ' erro de sync'; el.title = 'Erro de sincronização'; el.dataset.state = 'error'; console.error('SYNC ERROR from:', new Error().stack); }
  else if (status === 'syncing') { el.innerHTML = icon('refresh', '#f59e0b') + ' salvando'; el.title = 'Sincronizando...'; el.dataset.state = 'syncing'; }
}

// ─── SYNC GRANULAR ────────────────────────────────────────────────────────────
// Cada coleção vive no Firebase como objeto {id: item} (não mais array inteiro).
// save() calcula o que mudou desde o último snapshot recebido (_remote) e manda só
// isso num único update() multi-path. Dois dispositivos editando itens diferentes
// ao mesmo tempo não se sobrescrevem mais.
const ID_COLLECTIONS = ['transactions', 'accounts', 'debts', 'investments', 'deletedIds', 'loveMessages'];
const KEY_FIELDS     = ['settings', 'budget', 'catOrcGroup', 'chatHistory', 'customCats', 'customBanks', 'acertos'];
let _remote = null;             // último estado conhecido do Firebase (strings estáveis por id)
let _migrateColls = new Set();  // coleções ainda no formato antigo (array) → reescrever a chave inteira 1x
let _inflight = 0;              // escritas em andamento
let _pendingSnapshot = null;    // snapshot que chegou durante uma escrita (processa depois)
let _lastPushFail = 0;          // quando a última escrita falhou (limita retentativa automática)
let _lastAutoPush = 0;          // última escrita automática do listener (reparo/migração)

// Normaliza como o Firebase armazena: chaves ordenadas, sem null/undefined/vazios.
// Assim comparar stableStr(local) === stableStr(remoto) não dá falso positivo.
function fbNormalize(v) {
  if (v === null || v === undefined) return undefined;
  if (Array.isArray(v)) {
    const out = v.map(fbNormalize).map(x => x === undefined ? null : x);
    return out.length ? out : undefined;
  }
  if (typeof v === 'object') {
    const out = {};
    Object.keys(v).sort().forEach(k => { const x = fbNormalize(v[k]); if (x !== undefined) out[k] = x; });
    return Object.keys(out).length ? out : undefined;
  }
  return v;
}
function stableStr(v) { const n = fbNormalize(v); return n === undefined ? '' : JSON.stringify(n); }

// Formato antigo = chave do Firebase não bate com o id do item (array posicional 0,1,2...).
function _isLegacyShape(v) {
  if (!v || typeof v !== 'object') return false;
  if (Array.isArray(v)) return v.some((it, i) => it && String(it.id) !== String(i));
  return Object.keys(v).some(k => v[k] && String(v[k].id) !== k);
}

function _snapshotToRemote(fbData) {
  const r = {};
  ID_COLLECTIONS.forEach(c => {
    r[c] = {};
    fbToArray(fbData[c]).forEach(it => { if (it && it.id) r[c][String(it.id)] = stableStr(it); });
  });
  KEY_FIELDS.forEach(k => { r[k] = stableStr(fbData[k]); });
  return r;
}

// Envia pro Firebase SÓ o que mudou em relação a _remote. Se _remote é null (primeiro uso
// ou falha anterior), grava tudo por id de uma vez.
function pushToFirebase() {
  if (!db || !firebaseAuthReady || !firebaseReady) return;
  const clean = v => JSON.parse(JSON.stringify(v === undefined ? null : v));
  const full = !_remote;
  const next = full ? {} : _remote;
  const upd = {};

  ID_COLLECTIONS.forEach(c => {
    const local = (S[c] || []).filter(it => it && it.id);
    if (full || _migrateColls.has(c) || !next[c]) {
      // Reescreve a chave inteira no formato {id: item}
      const obj = {}; next[c] = {};
      local.forEach(it => { obj[String(it.id)] = clean(it); next[c][String(it.id)] = stableStr(it); });
      upd[c] = obj;
      return;
    }
    const seen = new Set();
    local.forEach(it => {
      const id = String(it.id); seen.add(id);
      const str = stableStr(it);
      if (next[c][id] !== str) { upd[c + '/' + id] = clean(it); next[c][id] = str; }
    });
    Object.keys(next[c]).forEach(id => {
      if (!seen.has(id)) { upd[c + '/' + id] = null; delete next[c][id]; }
    });
  });

  KEY_FIELDS.forEach(k => {
    const str = stableStr(S[k]);
    if (full || next[k] !== str) { upd[k] = clean(S[k]); next[k] = str; }
  });

  _remote = next;
  _migrateColls.clear();
  const nPaths = Object.keys(upd).length;
  if (!nPaths) { updateSyncStatus('ok'); return; }

  updateSyncStatus('syncing');
  _inflight++;
  db.ref('data').update(upd).then(() => {
    updateSyncStatus('ok');
  }).catch(e => {
    _remote = null; // próximo save regrava tudo
    _lastPushFail = Date.now();
    console.warn('Firebase save error:', e);
    updateSyncStatus('error');
  }).finally(() => {
    _inflight--;
    if (_inflight === 0 && _pendingSnapshot) {
      const snap = _pendingSnapshot; _pendingSnapshot = null;
      _onFirebaseValue(snap);
    }
  });
}

// Merge dois arrays por ID — mantém ambos os lados, prefere o mais recente por updatedAt
// Tombstones impedem que itens deletados ressuscitem
function mergeArrayById(local, remote, tombstones) {
  const tombSet = new Set((tombstones || []).map(t => t.id));
  const merged = new Map();
  for (const item of local) {
    if (item && item.id && !tombSet.has(item.id)) merged.set(item.id, item);
  }
  for (const item of remote) {
    if (!item || !item.id || tombSet.has(item.id)) continue;
    const existing = merged.get(item.id);
    if (!existing) { merged.set(item.id, item); continue; }
    const remoteTime = item.updatedAt || item.at || '';
    const localTime = existing.updatedAt || existing.at || '';
    if (remoteTime > localTime) merged.set(item.id, item);
  }
  return Array.from(merged.values());
}

// Merge dois arrays de tombstones (union por ID)
function mergeTombstones(local, remote) {
  const map = new Map();
  for (const t of (local || [])) { if (t && t.id) map.set(t.id, t); }
  for (const t of (remote || [])) { if (t && t.id && !map.has(t.id)) map.set(t.id, t); }
  return Array.from(map.values());
}

function initFirebaseSync() {
  if (!db) return;

  // Persistir sessão entre reloads (fica logado no dispositivo)
  firebase.auth().setPersistence(firebase.auth.Auth.Persistence.LOCAL).catch(()=>{});

  // Sem login anônimo: os dados só são baixados após login REAL (e-mail/senha).
  firebase.auth().onAuthStateChanged(user => {
    if (user && !user.isAnonymous) {
      if (!firebaseAuthReady) {
        firebaseAuthReady = true;
        _startFirebaseListener();
      }
      onAuthenticated(user);
    } else {
      firebaseAuthReady = false;
      firebaseReady = false;
      showLoginScreen();
    }
  });
}

function _startFirebaseListener() {
  db.ref('data').on('value', (snapshot) => {
    // Chegou snapshot no meio de uma escrita nossa: guarda o mais recente e processa
    // quando a escrita terminar (antes era descartado → perdia mudança do outro dispositivo)
    if (_inflight > 0) { _pendingSnapshot = snapshot; return; }
    _onFirebaseValue(snapshot);
  }, (error) => {
    console.error('Firebase listen error:', error.code, error.message);
    updateSyncStatus('error');
  });
}

function _onFirebaseValue(snapshot) {
    const fbData = snapshot.val();

    // Firebase vazio — primeiro uso. Marcar como pronto e enviar dados locais.
    if (!fbData) {
      _remote = null;
      firebaseReady = true;
      save();
      return;
    }

    // Coleções ainda no formato antigo (array) serão reescritas 1x como {id: item}
    ID_COLLECTIONS.forEach(c => { if (_isLegacyShape(fbData[c])) _migrateColls.add(c); });
    if (_migrateColls.size) console.log('[sync] migrando para formato por id:', [..._migrateColls].join(', '));
    _remote = _snapshotToRemote(fbData);

    // ── FIREBASE É A FONTE DA VERDADE ──
    // Substituir estado local com dados do Firebase (sem merge)
    const tombSet = new Set((fbToArray(fbData.deletedIds) || []).map(t => t.id));

    if (fbData.transactions) S.transactions = fbToArray(fbData.transactions).filter(t => t && t.id && !tombSet.has(t.id));
    if (fbData.accounts)     S.accounts     = fbToArray(fbData.accounts).filter(a => a && a.id && !tombSet.has(a.id));
    if (fbData.debts)        S.debts        = fbToArray(fbData.debts).filter(d => d && d.id && !tombSet.has(d.id));
    if (fbData.investments)  S.investments  = fbToArray(fbData.investments).filter(i => i && i.id && !tombSet.has(i.id));
    if (fbData.deletedIds)   S.deletedIds   = fbToArray(fbData.deletedIds);
    if (fbData.settings)     S.settings     = { ...S.settings, ...fbData.settings };
    if (fbData.budget)       S.budget       = fbData.budget;
    if (fbData.catOrcGroup)  S.catOrcGroup  = fbData.catOrcGroup;
    if (fbData.chatHistory)  S.chatHistory  = fbToArray(fbData.chatHistory);
    if (fbData.customCats)   S.customCats   = fbData.customCats;
    if (fbData.acertos)      S.acertos      = fbData.acertos;
    if (fbData.customBanks)  { S.customBanks = fbData.customBanks; Object.assign(BANKS, S.customBanks); }
    if (fbData.loveMessages) S.loveMessages = fbToArray(fbData.loveMessages).filter(m => m && (!m.id || !tombSet.has(m.id)));

    // ── REPAROS (rodam SEMPRE, sem flags one-time) ──
    let needsRepair = false;

    // 1. Contas/cartões sem titular
    if (S.settings.u1) {
      S.accounts.forEach(a => {
        if (!a.owner) {
          const thayseCards = ['bradesco', 'mercado_pago', 'itau'];
          a.owner = thayseCards.includes(a.bank) ? S.settings.u2 : S.settings.u1;
          a.updatedAt = new Date().toISOString();
          needsRepair = true;
        }
      });
    }

    // 2. Categorias customizadas — garantir que nunca fiquem faltando
    if (!S.customCats) S.customCats = JSON.parse(JSON.stringify(DEFAULT_CATS));
    const desp = S.customCats.despesa;
    if (desp) {
      const required = { 'Diversos': [], 'Servicos': [], 'Presentes': [], 'Pedro': ['Festa de Aniversário'], 'Taxas': ['Pgto Fatura'] };
      for (const [cat, subs] of Object.entries(required)) {
        // Firebase não armazena array vazio: categoria sem subcategoria SEMPRE chega ausente.
        // Só marca reparo quando há subcategorias a gravar (senão vira escrita infinita).
        if (!desp[cat]) { desp[cat] = subs.slice(); if (subs.length) needsRepair = true; }
      }
      if (desp['Saúde'] && !desp['Saúde'].includes('Cuidados Pessoais')) {
        desp['Saúde'].push('Cuidados Pessoais'); needsRepair = true;
      }
    }

    // 3. Parcelas do formulário com a data da compra repetida em todas (bug antigo:
    //    histórico/dashboard empilhavam as N parcelas no mês da compra).
    //    Assinatura exata do bug (as duas condições juntas):
    //      a) outra parcela da MESMA compra tem a MESMA data;
    //      b) faturaRef está (NN-1) meses à frente do que a data daria.
    //    Só mexe em `date` (faturaRef intacto → faturas não mudam). Idempotente: depois de
    //    corrigidas, as parcelas ficam em meses diferentes e nunca mais casam com (a).
    let txRepaired = 0;
    const parcNum = t => {
      if (!(t.parcelaTotal > 1) || !t.faturaRef || !t.date) return 0;
      const mp = /^(\d+)\/(\d+)$/.exec(t.parcela || '');
      return mp ? parseInt(mp[1]) : 0;
    };
    const grupoKey = t => [t.accountId, t.parcelaTotal, (t.desc || '').replace(/\s*\(\d+\/\d+\)\s*$/, ''), t.date].join('|');
    const numsPorGrupo = {};
    S.transactions.forEach(t => {
      const nn = parcNum(t);
      if (nn) (numsPorGrupo[grupoKey(t)] = numsPorGrupo[grupoKey(t)] || new Set()).add(nn);
    });
    const alvos = S.transactions.filter(t => {
      const nn = parcNum(t);
      if (nn < 2 || numsPorGrupo[grupoKey(t)].size < 2) return false;
      const card = S.accounts.find(a => a.id === t.accountId);
      const fechaDia = card ? parseInt(card.fecha) || 1 : 1;
      const [y, m, d] = t.date.split('-').map(Number);
      const ymData = y * 12 + (m - 1) + (d > fechaDia ? 1 : 0);
      const [fy, fm] = t.faturaRef.split('-').map(Number);
      return (fy * 12 + (fm - 1)) - ymData === nn - 1;
    });
    alvos.forEach(t => {
      t.date = addMesesData(t.date, parcNum(t) - 1);
      t.updatedAt = new Date().toISOString();
      txRepaired++;
    });
    if (txRepaired) console.log(`[reparo] ${txRepaired} parcela(s) com data corrigida`);

    // 4. Tombstones com mais de 30 dias saem do Firebase (antes só limpava o localStorage)
    const cutoff = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000).toISOString();
    const nTombs = S.deletedIds.length;
    S.deletedIds = S.deletedIds.filter(t => t && t.deletedAt > cutoff);
    const tombsRemoved = nTombs - S.deletedIds.length;

    // Marcar Firebase como carregado ANTES de cachear e renderizar
    firebaseReady = true;

    // Reparos / migração / limpeza vão pro Firebase numa única escrita (só o que mudou).
    // Limites: 1 min após falha (ex.: sem permissão) e no máximo 1 escrita automática a cada
    // 10 s — evita ping-pong com outro dispositivo rodando versão antiga do app (que regrava
    // arrays) e loop de escrita falha → snapshot → escrita falha. Saves do usuário não passam
    // por aqui e não têm limite.
    const now = Date.now();
    const autoPushOk = now - _lastPushFail > 60000 && now - _lastAutoPush > 10000;
    if ((needsRepair || txRepaired || tombsRemoved || _migrateColls.size) && autoPushOk) {
      _lastAutoPush = now;
      pushToFirebase();
    }

    // Cachear no localStorage (backup offline)
    localStorage.setItem('fincasal_v2', JSON.stringify(S));

    // Update header
    document.getElementById('header-users').textContent = `${S.settings.u1} & ${S.settings.u2}`;

    // Re-render só se logado
    if (sessionStorage.getItem('fincasal_auth')) {
      const current = document.querySelector('section[id^="page-"]:not([style*="display:none"])');
      if (current) {
        const page = current.id.replace('page-', '');
        if (page === 'dashboard') renderDashboard();
        else if (page === 'historico') renderHistorico();
        else if (page === 'dividas') renderDividas();
        else if (page === 'investimentos') renderInvestimentos();
        else if (page === 'leiaaqui') renderLeiaAqui();
        else if (page === 'mensagens') renderMensagensPanel();
        else if (page === 'orcamento') renderOrcamento();
        else if (page === 'categorias') renderCategorias();
      }
      updateLeiaAquiVisibility();
    }

    updateSyncStatus('ok');
}
