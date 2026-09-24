// ─────────────────────────────────────────────────────────────────────────────
// Salem no Telegram — bot bidirecional do FinançasCasal.
// Fica escutando o Telegram (long polling), responde perguntas sobre as finanças com o
// Claude usando os MESMOS dados do app (Firebase) e consegue lançar despesas.
//
// Roda no PC do Paulo (tarefa agendada "FinancasCasal - Salem Telegram", ao fazer logon).
// Configuração em reports/.env (NUNCA commitar). Ver reports/.env.example.
//
// Segurança: só responde aos chat ids de TELEGRAM_CHAT_IDS. Qualquer outra pessoa é ignorada.
// ─────────────────────────────────────────────────────────────────────────────
const fs = require('fs');
const path = require('path');

// ── .env simples (sem dependência) ──
(function loadEnv() {
  const p = path.join(__dirname, '.env');
  if (!fs.existsSync(p)) return;
  for (const line of fs.readFileSync(p, 'utf8').split('\n')) {
    const m = /^\s*([A-Z0-9_]+)\s*=\s*(.*?)\s*$/.exec(line);
    if (m && !process.env[m[1]]) process.env[m[1]] = m[2].replace(/^["']|["']$/g, '');
  }
})();

const TOKEN = process.env.TELEGRAM_BOT_TOKEN;
const KEY = process.env.ANTHROPIC_API_KEY;
const DB_URL = process.env.FIREBASE_DB_URL;
const SA_FILE = process.env.FIREBASE_SERVICE_ACCOUNT_FILE;
const MODEL = process.env.SALEM_MODEL || 'claude-opus-5';
const CHATS = {};   // chatId → Nome
(process.env.TELEGRAM_CHAT_IDS || '').split(',').map(s => s.trim()).filter(Boolean).forEach(e => { const i = e.indexOf(':'); CHATS[i > 0 ? e.slice(0, i).trim() : e] = i > 0 ? e.slice(i + 1).trim() : null; });

function faltando() {
  const f = [];
  if (!TOKEN) f.push('TELEGRAM_BOT_TOKEN');
  if (!KEY) f.push('ANTHROPIC_API_KEY');
  if (!DB_URL) f.push('FIREBASE_DB_URL');
  if (!SA_FILE || !fs.existsSync(path.resolve(__dirname, SA_FILE))) f.push('FIREBASE_SERVICE_ACCOUNT_FILE (arquivo não encontrado)');
  if (!Object.keys(CHATS).length) f.push('TELEGRAM_CHAT_IDS');
  return f;
}
const log = (...a) => console.log(new Date().toISOString().slice(0, 19).replace('T', ' '), ...a);

// ── Firebase ──
const admin = require('firebase-admin');
let rtdb = null;
function initFirebase() {
  const svc = JSON.parse(fs.readFileSync(path.resolve(__dirname, SA_FILE), 'utf8'));
  admin.initializeApp({ credential: admin.credential.cert(svc), databaseURL: DB_URL });
  rtdb = admin.database();
}
let _cache = { at: 0, data: null };
async function getData() {
  if (Date.now() - _cache.at < 30000 && _cache.data) return _cache.data;
  const raw = (await rtdb.ref('data').once('value')).val() || {};
  const arr = o => !o ? [] : (Array.isArray(o) ? o.filter(Boolean) : Object.values(o));
  _cache = { at: Date.now(), data: {
    transactions: arr(raw.transactions), accounts: arr(raw.accounts), debts: arr(raw.debts), investments: arr(raw.investments),
    budget: raw.budget || {}, catOrcGroup: raw.catOrcGroup || {}, customCats: raw.customCats || null, customBanks: raw.customBanks || {},
    settings: raw.settings || {}
  } };
  return _cache.data;
}

// ── helpers de dinheiro / dados (espelham o app) ──
const brl = v => 'R$ ' + (Math.round((v || 0) * 100) / 100).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const norm = d => String(d || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/\s*\(\d+\/\d+\)\s*$/, '').replace(/\s+/g, ' ').trim();
const isPgtoFatura = t => (t.desc && t.desc.startsWith('Pagamento Fatura')) || (t.category === 'Taxas' && t.subcategory === 'Pgto Fatura');
const amount = (t, usd) => { const base = t.currency === 'USD' ? (t.amount || 0) * (usd || 5) : (t.amount || 0); return t.isNegative ? -base : base; };
function faturaRefOf(t, card) {
  if (t.faturaRef) return t.faturaRef;
  if (!card || !card.fecha || !t.date) return (t.date || '').slice(0, 7);
  const [y, m, d] = t.date.split('-').map(Number);
  const ym = y * 12 + (m - 1) + (d > (parseInt(card.fecha) || 1) ? 1 : 0);
  return `${Math.floor(ym / 12)}-${String(ym % 12 + 1).padStart(2, '0')}`;
}

// ── ferramentas (mesmo contrato do app) ──
const TOOLS = [
  { name: 'summarize_transactions', description: 'Totais agregados com filtros e agrupamento. Use para QUALQUER pergunta de "quanto", média, ranking ou comparação. Exclui pagamentos de fatura e transferências.',
    input_schema: { type: 'object', properties: {
      start: { type: 'string', description: 'YYYY-MM-DD inclusive' }, end: { type: 'string', description: 'YYYY-MM-DD inclusive' },
      type: { type: 'string', enum: ['despesa', 'receita', 'investimento'] }, user: { type: 'string' }, category: { type: 'string' }, subcategory: { type: 'string' },
      accountId: { type: 'string' }, formaPgto: { type: 'string', enum: ['debito', 'credito'] }, custoTipo: { type: 'string', enum: ['fixo', 'variavel'] },
      query: { type: 'string' }, compartilhada: { type: 'boolean' },
      group_by: { type: 'string', enum: ['category', 'subcategory', 'month', 'user', 'account', 'desc', 'formaPgto', 'custoTipo', 'none'] }, limit: { type: 'number' } }, required: ['start', 'end'] } },
  { name: 'search_transactions', description: 'Lista lançamentos individuais (até 20) por texto, mês, categoria, conta ou tipo.',
    input_schema: { type: 'object', properties: { query: { type: 'string' }, month: { type: 'number' }, year: { type: 'number' }, category: { type: 'string' }, accountId: { type: 'string' }, type: { type: 'string' }, limit: { type: 'number' } } } },
  { name: 'get_fatura', description: 'Fatura de cartão num mês: total, pago, status e maiores lançamentos. Sem cardId, todas.',
    input_schema: { type: 'object', properties: { month: { type: 'number' }, year: { type: 'number' }, cardId: { type: 'string' } }, required: ['month', 'year'] } },
  { name: 'get_budget', description: 'Orçamento planejado vs realizado por categoria num mês.',
    input_schema: { type: 'object', properties: { month: { type: 'number' }, year: { type: 'number' } }, required: ['month', 'year'] } },
  { name: 'add_transaction', description: 'Lança uma transação nova no app. Use quando a pessoa disser que gastou/recebeu algo. Confirme os campos na resposta.',
    input_schema: { type: 'object', properties: {
      type: { type: 'string', enum: ['receita', 'despesa'] }, desc: { type: 'string' }, amount: { type: 'number' }, date: { type: 'string', description: 'YYYY-MM-DD' },
      accountId: { type: 'string', description: 'id de uma conta/cartão listado no contexto' }, category: { type: 'string' }, subcategory: { type: 'string' },
      formaPgto: { type: 'string', enum: ['debito', 'credito'] }, user: { type: 'string' }, compartilhada: { type: 'boolean' } }, required: ['type', 'desc', 'amount', 'date', 'accountId', 'category', 'user'] } }
];

async function runTool(name, a, ctx) {
  const D = await getData();
  const usd = (D.settings || {}).usdRate;
  const T = D.transactions, A = D.accounts;
  if (name === 'summarize_transactions') {
    const tipo = a.type || 'despesa', q = a.query ? norm(a.query) : '';
    let txs = T.filter(t => t && t.date && t.type === tipo && !t.isTransfer && !isPgtoFatura(t) && t.date >= a.start && t.date <= a.end);
    if (a.user) txs = txs.filter(t => t.user === a.user);
    if (a.category) txs = txs.filter(t => (t.category || '') === a.category);
    if (a.subcategory) txs = txs.filter(t => (t.subcategory || '') === a.subcategory);
    if (a.accountId) txs = txs.filter(t => t.accountId === a.accountId);
    if (a.formaPgto) txs = txs.filter(t => t.formaPgto === a.formaPgto);
    if (a.custoTipo) txs = txs.filter(t => t.custoTipo === a.custoTipo);
    if (a.compartilhada === true) txs = txs.filter(t => t.compartilhada === true);
    if (q) txs = txs.filter(t => norm(t.desc).includes(q));
    const total = txs.reduce((s, t) => s + amount(t, usd), 0);
    const gb = a.group_by || 'category';
    const keyOf = t => gb === 'none' ? 'total' : gb === 'month' ? t.date.slice(0, 7) : gb === 'account' ? ((A.find(x => x.id === t.accountId) || {}).label || t.accountId) : gb === 'desc' ? norm(t.desc) : (t[gb] || '(sem ' + gb + ')');
    const g = {}; txs.forEach(t => { const k = keyOf(t); g[k] = g[k] || { total: 0, n: 0 }; g[k].total += amount(t, usd); g[k].n++; });
    let rows = Object.entries(g).sort((x, y) => y[1].total - x[1].total);
    if (gb === 'month') rows.sort((x, y) => x[0].localeCompare(y[0])); else rows = rows.slice(0, a.limit || 15);
    const head = `Período ${a.start} a ${a.end} · ${tipo}${a.user ? ' · ' + a.user : ''}${a.category ? ' · ' + a.category : ''}\nTOTAL: ${brl(total)} em ${txs.length} lançamentos`;
    if (!txs.length) return head + '\n(nenhum lançamento)';
    return head + `\nPor ${gb}:\n` + rows.map(([k, v]) => `${k}: ${brl(v.total)} (${v.n}, ${total ? (100 * v.total / total).toFixed(1) : 0}%)`).join('\n');
  }
  if (name === 'search_transactions') {
    let txs = [...T];
    if (a.query) txs = txs.filter(t => norm(t.desc).includes(norm(a.query)) || norm(t.category).includes(norm(a.query)));
    if (a.month && a.year) txs = txs.filter(t => (t.date || '').startsWith(`${a.year}-${String(a.month).padStart(2, '0')}`));
    else if (a.year) txs = txs.filter(t => (t.date || '').startsWith(String(a.year)));
    if (a.category) txs = txs.filter(t => t.category === a.category);
    if (a.accountId) txs = txs.filter(t => t.accountId === a.accountId);
    if (a.type) txs = txs.filter(t => t.type === a.type);
    txs.sort((x, y) => (y.date || '').localeCompare(x.date || ''));
    txs = txs.slice(0, a.limit || 20);
    return txs.length ? txs.map(t => `${t.date} | ${t.desc} | ${t.type} | ${t.category} | ${brl(amount(t, usd))} | ${t.user}${t.pago === false ? ' | pendente' : ''}`).join('\n') : 'Nenhuma transação encontrada.';
  }
  if (name === 'get_fatura') {
    const ym = `${a.year}-${String(a.month).padStart(2, '0')}`;
    const cards = A.filter(c => c.accountType === 'cartao' && (!a.cardId || c.id === a.cardId));
    if (!cards.length) return 'Nenhum cartão.';
    return cards.map(c => {
      const txs = T.filter(t => t.type === 'despesa' && t.formaPgto === 'credito' && t.accountId === c.id && faturaRefOf(t, c) === ym);
      const total = txs.reduce((s, t) => s + amount(t, usd), 0);
      const pago = T.filter(t => isPgtoFatura(t) && ((t.faturaCartaoId === c.id && t.faturaYM === ym) || (!t.faturaCartaoId && (t.desc || '').includes(c.label || '§') && (t.date || '').startsWith(ym)))).reduce((s, t) => s + Math.abs(t.amount || 0), 0);
      const status = total <= 0 ? 'sem lançamentos' : pago >= total - 0.01 ? 'PAGA' : pago > 0 ? `parcial (pago ${brl(pago)})` : 'em aberto';
      const top = [...txs].sort((x, y) => amount(y, usd) - amount(x, usd)).slice(0, 5).map(t => `  - ${t.date} ${t.desc}: ${brl(amount(t, usd))}`).join('\n');
      return `${c.label} (id ${c.id}, ${c.owner || '?'}, fecha ${c.fecha || '?'}, vence ${c.vence || '?'}) — fatura ${ym}: ${brl(total)} em ${txs.length} lançamentos · ${status}${top ? '\n' + top : ''}`;
    }).join('\n\n');
  }
  if (name === 'get_budget') {
    const ym = `${a.year}-${String(a.month).padStart(2, '0')}`;
    const gasto = {};
    T.filter(t => t.type === 'despesa' && !t.isTransfer && !isPgtoFatura(t) && (t.date || '').startsWith(ym)).forEach(t => { const c = t.category || '(sem)'; gasto[c] = (gasto[c] || 0) + amount(t, usd); });
    const cats = new Set(Object.keys(gasto)); Object.keys(D.budget).forEach(k => { if (k.startsWith(ym + '-')) cats.add(k.slice(ym.length + 1)); });
    const rows = [...cats].map(c => ({ c, plan: D.budget[`${ym}-${c}`] || 0, real: gasto[c] || 0 })).sort((x, y) => (y.real - y.plan) - (x.real - x.plan));
    return rows.length ? `Orçamento ${ym}\n` + rows.map(r => `${r.c}: planejado ${brl(r.plan)} · realizado ${brl(r.real)} · ${r.plan ? (r.real > r.plan ? 'ESTOUROU ' + brl(r.real - r.plan) : 'sobra ' + brl(r.plan - r.real)) : 'sem meta'}`).join('\n') : `Sem orçamento nem gastos em ${ym}.`;
  }
  if (name === 'add_transaction') {
    const acc = A.find(x => x.id === a.accountId);
    if (!acc) return 'Conta inválida: ' + a.accountId + '. Use um id listado no contexto.';
    if (!(a.amount > 0)) return 'Valor inválido.';
    const isCred = acc.accountType === 'cartao';
    const id = 'tg_' + Date.now() + '_' + Math.random().toString(36).slice(2, 6);
    const tx = { id, type: a.type || 'despesa', desc: a.desc, amount: Math.round(a.amount * 100) / 100, currency: 'BRL', accountId: acc.id,
      category: a.category || 'Outros', subcategory: a.subcategory || '', date: a.date, user: a.user || ctx.nome || D.settings.u1, notes: 'via Telegram',
      pago: true, formaPgto: a.type === 'despesa' ? (isCred ? 'credito' : (a.formaPgto || 'debito')) : null, custoTipo: a.type === 'despesa' ? 'variavel' : null,
      compartilhada: a.compartilhada === true, at: new Date().toISOString(), updatedAt: new Date().toISOString() };
    if (isCred) tx.faturaRef = faturaRefOf(tx, acc);
    await rtdb.ref('data/transactions/' + id).set(tx);
    _cache.at = 0;
    return `Lançado: ${tx.desc} — ${brl(tx.amount)} em ${tx.date}, ${acc.label} (${tx.formaPgto || tx.type}), ${tx.category}${tx.subcategory ? ' › ' + tx.subcategory : ''}, titular ${tx.user}${isCred ? ', fatura ' + tx.faturaRef : ''}.`;
  }
  return 'Ferramenta desconhecida: ' + name;
}

async function systemPrompt(nome) {
  const D = await getData();
  const now = new Date();
  const accs = D.accounts.map(a => `- ${a.label} (id: ${a.id}, ${a.accountType === 'cartao' ? 'cartão' : 'conta'}, titular ${a.owner || '?'})`).join('\n');
  const cats = D.customCats && D.customCats.despesa ? Object.entries(D.customCats.despesa).map(([c, s]) => Array.isArray(s) && s.length ? `${c} (${s.join(', ')})` : c).join('; ') : '';
  return `Você é Salem, assistente financeiro do casal ${D.settings.u1} e ${D.settings.u2}, respondendo pelo Telegram. Responda SEMPRE em PT-BR, curto e direto (é chat de celular): no máximo uns 8 linhas, sem markdown pesado — use *negrito* só em números importantes e listas com "•".
Está falando com: ${nome || 'desconhecido'}. Data de hoje: ${now.toISOString().slice(0, 10)}.

CONTAS E CARTÕES:
${accs}

CATEGORIAS DE DESPESA: ${cats}

REGRAS:
- NUNCA estime números: use summarize_transactions (totais), get_fatura, get_budget ou search_transactions.
- "Este mês" = ${now.toISOString().slice(0, 7)}. Ano corrente = ${now.getFullYear()}.
- Quando a pessoa disser que gastou algo ("gastei 50 no mercado no débito"), lance com add_transaction: titular = quem está falando (${nome || 'quem está falando'}), data = hoje se não disser, conta = a conta/cartão de quem fala que combine com o que foi dito; categoria pela descrição. Se faltar algo essencial (valor ou conta), pergunte antes de lançar.
- Despesas "do casal" (compartilhada) são divididas 50/50.
- Quando usar uma ferramenta, pode dizer uma frase curta antes. Se nenhuma ferramenta expressar o pedido, diga isso em vez de chutar.`;
}

// ── Claude ──
const history = {};   // chatId → [{role, content}]
async function claude(messages, system) {
  const body = { model: MODEL, max_tokens: 4000, system, messages, tools: TOOLS };
  const headers = { 'Content-Type': 'application/json', 'x-api-key': KEY, 'anthropic-version': '2023-06-01' };
  if (!MODEL.startsWith('claude-haiku-4-5')) { body.thinking = { type: 'adaptive' }; body.output_config = { effort: 'medium' }; }
  if (MODEL === 'claude-opus-5') { body.fallbacks = 'default'; headers['anthropic-beta'] = 'server-side-fallback-2026-07-01'; }
  const res = await fetch('https://api.anthropic.com/v1/messages', { method: 'POST', headers, body: JSON.stringify(body) });
  if (!res.ok) { let d = ''; try { d = (await res.json()).error?.message || ''; } catch (_) {} throw new Error(`Anthropic ${res.status} ${d}`); }
  return res.json();
}
async function responder(chatId, texto) {
  const nome = CHATS[chatId];
  const h = history[chatId] = (history[chatId] || []).slice(-12);
  h.push({ role: 'user', content: texto });
  const system = await systemPrompt(nome);
  let convo = [...h];
  let data = await claude(convo, system);
  for (let round = 0; round < 6 && data.stop_reason === 'tool_use'; round++) {
    const uses = data.content.filter(b => b.type === 'tool_use');
    const results = [];
    for (const u of uses) {
      let out; try { out = await runTool(u.name, u.input || {}, { nome }); } catch (e) { out = 'Erro: ' + e.message; }
      log('  tool', u.name, JSON.stringify(u.input).slice(0, 120));
      results.push({ type: 'tool_result', tool_use_id: u.id, content: String(out) });
    }
    convo = [...convo, { role: 'assistant', content: data.content }, { role: 'user', content: results }];
    data = await claude(convo, system);
  }
  if (data.stop_reason === 'refusal') return 'Não consigo responder a isso.';
  const txt = data.content.filter(b => b.type === 'text').map(b => b.text).join('\n').trim() || 'Pronto!';
  h.push({ role: 'assistant', content: txt });
  return txt;
}

// ── Telegram ──
const tg = (m, body) => fetch(`https://api.telegram.org/bot${TOKEN}/${m}`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) }).then(r => r.json());
function toTelegramHtml(md) {
  const esc = s => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  return esc(md).replace(/\*\*([^*\n]+)\*\*/g, '<b>$1</b>').replace(/(^|[^*])\*([^*\n]+)\*(?!\*)/g, '$1<b>$2</b>').replace(/`([^`\n]+)`/g, '<code>$1</code>').replace(/^#{1,3}\s+(.*)$/gm, '<b>$1</b>').replace(/^\s*[-•]\s+/gm, '• ');
}
async function send(chatId, texto) {
  const r = await tg('sendMessage', { chat_id: chatId, text: toTelegramHtml(texto), parse_mode: 'HTML', disable_web_page_preview: true });
  if (!r.ok) await tg('sendMessage', { chat_id: chatId, text: texto });   // fallback sem HTML
}

async function main() {
  const f = faltando();
  if (f.length) { log('Configuração incompleta em reports/.env — faltam:', f.join(', '), '\nVer reports/.env.example. Tentando de novo em 10 min.'); await new Promise(r => setTimeout(r, 600000)); process.exit(2); }
  initFirebase();
  const me = await tg('getMe', {});
  log(`Salem online como @${me.result && me.result.username} · modelo ${MODEL} · chats: ${Object.entries(CHATS).map(([i, n]) => n || i).join(', ')}`);
  let offset = 0;
  for (;;) {
    try {
      const r = await fetch(`https://api.telegram.org/bot${TOKEN}/getUpdates?timeout=50&offset=${offset}`, { signal: AbortSignal.timeout(60000) }).then(x => x.json());
      if (!r.ok) { log('getUpdates falhou:', r.description); await new Promise(x => setTimeout(x, 5000)); continue; }
      for (const u of r.result) {
        offset = u.update_id + 1;
        const msg = u.message; if (!msg || !msg.text) continue;
        const chatId = String(msg.chat.id);
        if (!(chatId in CHATS)) { log('ignorado chat desconhecido', chatId, msg.from && msg.from.username); continue; }
        const texto = msg.text.trim();
        log(`${CHATS[chatId] || chatId}: ${texto.slice(0, 100)}`);
        if (/^\/(start|ajuda|help)$/i.test(texto)) { await send(chatId, `Oi${CHATS[chatId] ? ', ' + CHATS[chatId] : ''}! Sou o Salem. Pergunte coisas como:\n• quanto gastamos este mês?\n• fatura do Bradesco de outubro\n• quanto gastei com mercado em setembro?\n• quem deve pra quem este mês?\nOu me conte um gasto: "gastei 48,90 no mercado no débito do Sicredi".`); continue; }
        if (/^\/limpar$/i.test(texto)) { history[chatId] = []; await send(chatId, 'Conversa reiniciada.'); continue; }
        await tg('sendChatAction', { chat_id: chatId, action: 'typing' });
        try { await send(chatId, await responder(chatId, texto)); }
        catch (e) { log('erro:', e.message); await send(chatId, 'Deu erro aqui: ' + e.message.slice(0, 200)); }
      }
    } catch (e) { log('loop:', e.message); await new Promise(x => setTimeout(x, 5000)); }
  }
}
main().catch(e => { log('fatal:', e.message); process.exit(1); });
