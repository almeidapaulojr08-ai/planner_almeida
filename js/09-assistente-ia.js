// FinançasCasal — 09-assistente-ia.js
// Módulo carregado em ordem por index.html (escopo global compartilhado). Não use import/export.

// ─── CHAT IA ──────────────────────────────────────────────────────────────────
// ─── CHAT ASSISTENTE IA (keyword-based) ──────────────────────────────────────
let chatInitialized = false;
let chatImportMode = false;
let chatImportState = { step: null, card: null, month: null, year: null, transactions: [] };

function initChat() {
  if (chatInitialized) return;
  chatInitialized = true;
  const c = document.getElementById('chat-msgs');
  c.innerHTML = '';
  addBubble('ai', 'Olá! Sou seu assistente financeiro. Posso responder perguntas sobre seus gastos, saldos, faturas, investimentos e muito mais. Também posso <b>importar faturas</b> de cartão a partir de texto colado. Como posso ajudar?');
}

// ─── AI INTEGRATION ──────────────────────────────────────────────────────────

const AI_DEFAULT_MODEL = 'claude-opus-5';
const AI_MODELS = [
  { id: 'claude-opus-5',   label: 'Claude Opus 5 — mais inteligente (padrão)' },
  { id: 'claude-sonnet-5', label: 'Claude Sonnet 5 — equilíbrio' },
  { id: 'claude-haiku-4-5', label: 'Claude Haiku 4.5 — mais rápido e barato' }
];

async function callAI(messages, tools) {
  const key = S.settings.apiKey;
  const provider = S.settings.apiProvider || 'openai';
  if (!key) return null;

  try {
    if (provider === 'anthropic') {
      const systemMsg = messages.find(m => m.role === 'system');
      const chatMsgs = messages.filter(m => m.role !== 'system');
      const anthropicTools = tools ? tools.map(t => ({
        name: t.function.name,
        description: t.function.description,
        input_schema: t.function.parameters
      })) : undefined;
      const model = S.settings.aiModel || AI_DEFAULT_MODEL;
      const body = {
        model,
        max_tokens: 8000,
        system: systemMsg?.content || '',
        messages: chatMsgs,
        tools: anthropicTools
      };
      const headers = {
        'Content-Type': 'application/json',
        'x-api-key': key,
        'anthropic-version': '2023-06-01',
        'anthropic-dangerous-direct-browser-access': 'true'
      };
      if (model.startsWith('claude-haiku-4-5')) {
        // Haiku 4.5 não tem adaptive thinking nem effort
      } else {
        body.thinking = { type: 'adaptive' };
        body.output_config = { effort: 'medium' };
      }
      if (model === 'claude-opus-5') {
        // Se um classificador de segurança recusar, a API reexecuta noutro modelo (server-side)
        body.fallbacks = 'default';
        headers['anthropic-beta'] = 'server-side-fallback-2026-07-01';
      }
      const res = await fetch('https://api.anthropic.com/v1/messages', { method: 'POST', headers, body: JSON.stringify(body) });
      if (!res.ok) {
        let detail = '';
        try { detail = (await res.json()).error?.message || ''; } catch (_) {}
        throw new Error(`Anthropic API ${res.status}${detail ? ': ' + detail : ''}`);
      }
      const data = await res.json();
      if (data.stop_reason === 'refusal') {
        return { text: 'Não consegui responder a essa pergunta (a solicitação foi recusada pela API).', toolCalls: [], content: data.content };
      }
      const textBlocks = data.content.filter(b => b.type === 'text').map(b => b.text).join('\n');
      const toolUses = data.content.filter(b => b.type === 'tool_use');
      // content bruto (inclui blocos de thinking) tem que voltar inteiro na próxima rodada
      return { text: textBlocks, toolCalls: toolUses.map(t => ({ name: t.name, arguments: t.input, id: t.id })), content: data.content, stopReason: data.stop_reason };
    } else {
      const res = await fetch('https://api.openai.com/v1/chat/completions', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + key },
        body: JSON.stringify({
          model: 'gpt-4o-mini',
          messages,
          tools: tools || undefined,
          temperature: 0.3
        })
      });
      if (!res.ok) throw new Error(`OpenAI API error: ${res.status}`);
      const data = await res.json();
      const msg = data.choices[0].message;
      const toolCalls = (msg.tool_calls || []).map(tc => ({
        name: tc.function.name,
        arguments: JSON.parse(tc.function.arguments),
        id: tc.id
      }));
      return { text: msg.content || '', toolCalls };
    }
  } catch(e) {
    console.error('AI API error:', e);
    return { text: `Erro na API: ${e.message}. Verifique sua chave de API nas Configurações.`, toolCalls: [] };
  }
}

function buildSystemPrompt() {
  const now = new Date();
  const curMonth = now.getMonth();
  const curYear = now.getFullYear();

  // Account balances
  const accInfo = S.accounts.map(a => {
    const txs = S.transactions.filter(t => t.accountId === a.id);
    const rec = txs.filter(t => t.type === 'receita').reduce((s, t) => s + amountBrl(t), 0);
    const desp = txs.filter(t => t.type === 'despesa').reduce((s, t) => s + amountBrl(t), 0);
    const bal = rec - desp;
    return `- ${a.label} (id: ${a.id}, tipo: ${a.accountType}): saldo ${brl(bal)}${a.accountType === 'cartao' && a.limit ? ', limite ' + brl(a.limit) : ''}`;
  }).join('\n');

  // Month summary
  const mTx = S.transactions.filter(t => {
    const d = new Date(t.date + 'T12:00:00');
    return d.getMonth() === curMonth && d.getFullYear() === curYear;
  });
  const totalRec = mTx.filter(t => t.type === 'receita').reduce((s, t) => s + amountBrl(t), 0);
  const totalDesp = mTx.filter(t => t.type === 'despesa').reduce((s, t) => s + amountBrl(t), 0);

  // Debts
  const debtsInfo = S.debts.length ? S.debts.map(d => `- ${d.desc}: ${brl(d.remaining || d.total)}`).join('\n') : 'Nenhuma';

  // Investments
  const invInfo = S.investments.length ? S.investments.map(i => `- ${i.desc}: ${brl(i.currentValue || i.amount)}`).join('\n') : 'Nenhum';

  // Categories
  // Categorias podem vir como array (receita) ou objeto {categoria: [subcategorias]} (despesa)
  const catsToText = (x) => {
    if (!x) return '';
    if (Array.isArray(x)) return x.map(c => typeof c === 'string' ? c : (c && c.name) || '').filter(Boolean).join(', ');
    return Object.entries(x).map(([c, subs]) => Array.isArray(subs) && subs.length ? `${c} (${subs.join(', ')})` : c).join('; ');
  };
  const despCats = catsToText(getDespesaCats());
  const recCats = catsToText(getReceitaCats());

  return `Você é Salem, assistente financeiro de um casal brasileiro. Responda SEMPRE em PT-BR.
Data atual: ${now.toISOString().slice(0,10)}. Usuários: ${S.settings.u1} e ${S.settings.u2}.

CONTAS:
${accInfo}

MÊS ATUAL (${MESES_FULL[curMonth]} ${curYear}):
- Receitas: ${brl(totalRec)}
- Despesas: ${brl(totalDesp)}
- ${mTx.length} transações

DÍVIDAS:
${debtsInfo}

INVESTIMENTOS:
${invInfo}

CATEGORIAS DESPESA: ${despCats}
CATEGORIAS RECEITA: ${recCats}

REGRAS:
- Formate valores como R$ X.XXX,XX
- Seja conciso e direto; use listas curtas quando comparar coisas
- NUNCA estime números de cabeça: qualquer total, média, comparação ou ranking vem de summarize_transactions (agregado), get_fatura (fatura de cartão) ou get_budget (orçamento vs realizado). search_transactions serve pra listar lançamentos individuais.
- "Este mês" = ${now.toISOString().slice(0,7)}; "mês passado" = mês anterior. Ano corrente = ${curYear}.
- Quando o usuário perguntar "com quem/quem gastou", agrupe por usuário (group_by: user). Categorias como Pedro são category.
- Despesas de cartão de crédito contam na data da compra; pagamentos de fatura já são excluídos automaticamente dos totais.
- Use as ferramentas para adicionar, buscar, importar ou excluir transações
- Ao adicionar transações, use IDs de conta válidos listados acima
- Para datas sem ano, use ${curYear}
- Quando usar uma ferramenta, pode dizer uma frase curta antes. Se nenhuma ferramenta expressar o que foi pedido, diga isso em vez de chutar.`;
}

const AI_TOOLS = [
  {
    type: 'function',
    function: {
      name: 'add_transaction',
      description: 'Add a new financial transaction (income, expense, or investment)',
      parameters: {
        type: 'object',
        properties: {
          type: { type: 'string', enum: ['receita', 'despesa', 'investimento'], description: 'Transaction type' },
          desc: { type: 'string', description: 'Description' },
          amount: { type: 'number', description: 'Amount in BRL' },
          date: { type: 'string', description: 'Date in YYYY-MM-DD format' },
          category: { type: 'string', description: 'Category name' },
          subcategory: { type: 'string', description: 'Subcategory (optional)' },
          accountId: { type: 'string', description: 'Account ID' },
          user: { type: 'string', description: 'Owner name (Paulo or Thayse)' },
          formaPgto: { type: 'string', enum: ['debito', 'credito', 'pix'], description: 'Payment method (for expenses)' }
        },
        required: ['type', 'desc', 'amount', 'date', 'category', 'accountId', 'user']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'search_transactions',
      description: 'Search transactions by description, category, month, or account',
      parameters: {
        type: 'object',
        properties: {
          query: { type: 'string', description: 'Search text (optional)' },
          month: { type: 'number', description: '1-12 month filter (optional)' },
          year: { type: 'number', description: 'Year filter (optional)' },
          category: { type: 'string', description: 'Category filter (optional)' },
          accountId: { type: 'string', description: 'Account filter (optional)' },
          type: { type: 'string', enum: ['receita', 'despesa', 'investimento'], description: 'Type filter (optional)' },
          limit: { type: 'number', description: 'Max results (default 20)' }
        }
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'import_fatura_text',
      description: 'Import invoice/fatura text. Parses bank statement text and creates multiple transactions.',
      parameters: {
        type: 'object',
        properties: {
          text: { type: 'string', description: 'Raw fatura/statement text to parse' },
          accountId: { type: 'string', description: 'Account/card ID for the transactions' },
          month: { type: 'number', description: 'Reference month (1-12)' },
          year: { type: 'number', description: 'Reference year' },
          user: { type: 'string', description: 'Owner name' },
          type: { type: 'string', enum: ['receita', 'despesa'], description: 'Transaction type (default: despesa)' }
        },
        required: ['text', 'accountId', 'month', 'year', 'user']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'summarize_transactions',
      description: 'Totais agregados de transações com filtros e agrupamento. Use para QUALQUER pergunta de "quanto", média, ranking ou comparação. Retorna total, quantidade e os grupos ordenados por valor. Pagamentos de fatura e transferências ficam de fora automaticamente.',
      parameters: {
        type: 'object',
        properties: {
          start: { type: 'string', description: 'Data inicial YYYY-MM-DD (inclusive)' },
          end: { type: 'string', description: 'Data final YYYY-MM-DD (inclusive)' },
          type: { type: 'string', enum: ['despesa', 'receita', 'investimento'], description: 'Tipo (padrão: despesa)' },
          user: { type: 'string', description: 'Titular (nome exato)' },
          category: { type: 'string', description: 'Categoria exata' },
          subcategory: { type: 'string', description: 'Subcategoria exata' },
          accountId: { type: 'string', description: 'ID da conta/cartão' },
          formaPgto: { type: 'string', enum: ['debito', 'credito'], description: 'Forma de pagamento' },
          custoTipo: { type: 'string', enum: ['fixo', 'variavel'], description: 'Custo fixo ou variável' },
          query: { type: 'string', description: 'Texto contido na descrição' },
          group_by: { type: 'string', enum: ['category', 'subcategory', 'month', 'user', 'account', 'desc', 'formaPgto', 'custoTipo', 'none'], description: 'Como agrupar (padrão: category)' },
          limit: { type: 'number', description: 'Máximo de grupos retornados (padrão 15)' }
        },
        required: ['start', 'end']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'get_fatura',
      description: 'Fatura de um cartão de crédito num mês: total, quantidade de lançamentos, quanto já foi pago, vencimento, e os maiores lançamentos. Sem cardId, retorna todas as faturas do mês.',
      parameters: {
        type: 'object',
        properties: {
          month: { type: 'number', description: 'Mês da fatura 1-12' },
          year: { type: 'number', description: 'Ano da fatura' },
          cardId: { type: 'string', description: 'ID do cartão (opcional)' }
        },
        required: ['month', 'year']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'get_budget',
      description: 'Orçamento planejado vs realizado por categoria num mês (só categorias com orçamento definido ou com gasto).',
      parameters: {
        type: 'object',
        properties: {
          month: { type: 'number', description: 'Mês 1-12' },
          year: { type: 'number', description: 'Ano' }
        },
        required: ['month', 'year']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'delete_transaction',
      description: 'Delete a transaction by ID',
      parameters: {
        type: 'object',
        properties: {
          id: { type: 'string', description: 'Transaction ID to delete' }
        },
        required: ['id']
      }
    }
  }
];

function executeAITool(name, args) {
  if (name === 'add_transaction') {
    const tx = {
      id: 'ai_' + Date.now() + '_' + Math.random().toString(36).slice(2,6),
      type: args.type || 'despesa',
      desc: args.desc,
      amount: args.amount,
      currency: 'BRL',
      accountId: args.accountId,
      category: args.category || 'Outros',
      subcategory: args.subcategory || '',
      date: args.date,
      user: args.user || S.settings.u1,
      notes: '',
      pago: true,
      formaPgto: args.formaPgto || (args.type === 'despesa' ? 'debito' : null),
      custoTipo: null,
      at: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    };
    S.transactions.push(tx);
    save();
    return `Transação adicionada: ${args.desc} - R$ ${args.amount.toFixed(2)}`;
  }

  if (name === 'search_transactions') {
    let txs = [...S.transactions];
    if (args.query) txs = txs.filter(t => t.desc.toLowerCase().includes(args.query.toLowerCase()) || (t.category||'').toLowerCase().includes(args.query.toLowerCase()));
    if (args.month && args.year) { const ym = `${args.year}-${String(args.month).padStart(2,'0')}`; txs = txs.filter(t => t.date.startsWith(ym)); }
    else if (args.year) txs = txs.filter(t => t.date.startsWith(String(args.year)));
    if (args.category) txs = txs.filter(t => t.category === args.category);
    if (args.accountId) txs = txs.filter(t => t.accountId === args.accountId);
    if (args.type) txs = txs.filter(t => t.type === args.type);
    txs.sort((a,b) => b.date.localeCompare(a.date));
    txs = txs.slice(0, args.limit || 20);
    if (!txs.length) return 'Nenhuma transação encontrada.';
    return txs.map(t => `${t.date} | ${t.desc} | ${t.type} | ${t.category} | R$ ${t.amount.toFixed(2)} | ${t.user}`).join('\n');
  }

  if (name === 'import_fatura_text') {
    const parsed = parseFaturaText(args.text, args.month, args.year);
    if (!parsed || !parsed.length) return 'Não consegui extrair transações do texto fornecido.';
    const acc = S.accounts.find(a => a.id === args.accountId);
    const isCartao = acc && acc.accountType === 'cartao';
    const tipo = args.type || 'despesa';
    let count = 0;
    for (const p of parsed) {
      S.transactions.push({
        id: 'ai_imp_' + Date.now() + '_' + (count++),
        type: tipo,
        desc: p.desc,
        amount: p.amount,
        currency: 'BRL',
        accountId: args.accountId,
        category: p.category || 'Outros',
        subcategory: p.subcategory || '',
        date: p.date,
        user: args.user || S.settings.u1,
        notes: '',
        pago: true,
        formaPgto: isCartao ? 'credito' : 'debito',
        custoTipo: p.custoTipo || null,
        faturaRef: args.year + '-' + String(args.month).padStart(2,'0'),
        at: new Date().toISOString(),
        updatedAt: new Date().toISOString()
      });
    }
    save();
    return `Importadas ${count} transações da fatura para ${acc ? acc.label : args.accountId}.`;
  }

  if (name === 'summarize_transactions') {
    const tipo = args.type || 'despesa';
    const q = args.query ? normDesc(args.query) : '';
    let txs = S.transactions.filter(t => t && t.date && t.type === tipo && !t.isTransfer && !isPgtoFatura(t)
      && t.date >= args.start && t.date <= args.end);
    if (args.user)        txs = txs.filter(t => t.user === args.user);
    if (args.category)    txs = txs.filter(t => (t.category || '') === args.category);
    if (args.subcategory) txs = txs.filter(t => (t.subcategory || '') === args.subcategory);
    if (args.accountId)   txs = txs.filter(t => t.accountId === args.accountId);
    if (args.formaPgto)   txs = txs.filter(t => t.formaPgto === args.formaPgto);
    if (args.custoTipo)   txs = txs.filter(t => t.custoTipo === args.custoTipo);
    if (q)                txs = txs.filter(t => normDesc(t.desc).includes(q));
    const total = txs.reduce((s, t) => s + amountBrl(t), 0);
    const gb = args.group_by || 'category';
    const keyOf = t => {
      if (gb === 'none') return 'total';
      if (gb === 'month') return t.date.slice(0, 7);
      if (gb === 'account') { const a = S.accounts.find(x => x.id === t.accountId); return a ? a.label : (t.accountId || '(sem conta)'); }
      if (gb === 'desc') return (t.desc || '').replace(/\s*\(\d+\/\d+\)\s*$/, '');
      return t[gb] || '(sem ' + gb + ')';
    };
    const groups = {};
    txs.forEach(t => { const k = keyOf(t); (groups[k] = groups[k] || { total: 0, n: 0 }); groups[k].total += amountBrl(t); groups[k].n++; });
    const rows = Object.entries(groups).sort((a, b) => b[1].total - a[1].total);
    const lim = args.limit || 15;
    const shown = gb === 'month' ? rows.sort((a, b) => a[0].localeCompare(b[0])) : rows.slice(0, lim);
    const head = `Período ${args.start} a ${args.end} · tipo ${tipo}${args.user ? ' · titular ' + args.user : ''}${args.category ? ' · categoria ' + args.category : ''}\nTOTAL: ${brl(total)} em ${txs.length} lançamentos`;
    if (!txs.length) return head + '\n(nenhum lançamento com esses filtros)';
    const body = shown.map(([k, v]) => `${k}: ${brl(v.total)} (${v.n} lanç., ${total ? (100 * v.total / total).toFixed(1) : 0}%)`).join('\n');
    const extra = rows.length > shown.length ? `\n(+${rows.length - shown.length} grupos menores)` : '';
    return head + `\nAgrupado por ${gb}:\n` + body + extra;
  }

  if (name === 'get_fatura') {
    const ym = `${args.year}-${String(args.month).padStart(2, '0')}`;
    const cards = S.accounts.filter(a => a.accountType === 'cartao' && (!args.cardId || a.id === args.cardId));
    if (!cards.length) return 'Nenhum cartão encontrado.';
    return cards.map(c => {
      const txs = getCreditTxByFatura(ym, c.id);
      const total = txs.reduce((s, t) => s + amountBrl(t), 0);
      const pagamentos = getPagamentosFatura(c.id, args.month - 1, args.year);
      const pago = pagamentos.reduce((s, p) => s + Math.abs(p.amount || 0), 0);
      const status = total <= 0 ? 'sem lançamentos' : (pago >= total - 0.01 ? 'PAGA' : (pago > 0 ? `parcial (pago ${brl(pago)})` : 'em aberto'));
      const top = [...txs].sort((a, b) => amountBrl(b) - amountBrl(a)).slice(0, 5).map(t => `  - ${t.date} ${t.desc}: ${brl(amountBrl(t))}`).join('\n');
      const parc = txs.filter(t => t.parcela).length;
      return `${c.label} (id ${c.id}, titular ${c.owner || '?'}, fecha dia ${c.fecha || '?'}, vence dia ${c.vence || '?'}) — fatura ${ym}: ${brl(total)} em ${txs.length} lançamentos (${parc} parcelas) · status: ${status}${top ? '\n  Maiores:\n' + top : ''}`;
    }).join('\n\n');
  }

  if (name === 'get_budget') {
    const mes = args.month - 1, ano = args.year;
    const ym = `${ano}-${String(args.month).padStart(2, '0')}`;
    const gasto = {};
    S.transactions.filter(t => t.type === 'despesa' && !t.isTransfer && !isPgtoFatura(t) && (t.date || '').startsWith(ym))
      .forEach(t => { const c = t.category || '(sem)'; gasto[c] = (gasto[c] || 0) + amountBrl(t); });
    const cats = new Set(Object.keys(gasto));
    Object.keys(S.budget || {}).forEach(k => { if (k.startsWith(ym + '-')) cats.add(k.slice(ym.length + 1)); });
    const rows = [...cats].map(c => ({ c, plan: (S.budget || {})[getBudgetKey(ano, mes, c)] || 0, real: gasto[c] || 0 }))
      .sort((a, b) => (b.real - b.plan) - (a.real - a.plan));
    if (!rows.length) return `Sem orçamento nem gastos em ${ym}.`;
    const tp = rows.reduce((s, r) => s + r.plan, 0), tr = rows.reduce((s, r) => s + r.real, 0);
    return `Orçamento ${ym} — planejado ${brl(tp)} · realizado ${brl(tr)}\n` +
      rows.map(r => `${r.c}: planejado ${brl(r.plan)} · realizado ${brl(r.real)} · ${r.plan ? (r.real > r.plan ? 'ESTOUROU ' + brl(r.real - r.plan) : 'sobra ' + brl(r.plan - r.real)) : 'sem meta'}`).join('\n');
  }

  if (name === 'delete_transaction') {
    const idx = S.transactions.findIndex(t => t.id === args.id);
    if (idx < 0) return 'Transação não encontrada.';
    const tx = S.transactions[idx];
    S.deletedIds.push({ id: args.id, collection: 'transactions', deletedAt: new Date().toISOString() });
    if (tx.isTransfer && tx.transferId) {
      S.transactions.filter(t => t.transferId === tx.transferId).forEach(t => {
        S.deletedIds.push({ id: t.id, collection: 'transactions', deletedAt: new Date().toISOString() });
      });
      S.transactions = S.transactions.filter(t => t.transferId !== tx.transferId);
    } else {
      S.transactions.splice(idx, 1);
    }
    save();
    return `Transação excluída: ${tx.desc}`;
  }

  return 'Ferramenta desconhecida: ' + name;
}

function updateAIStatus() {
  const el = document.getElementById('ai-status');
  if (!el) return;
  if (S.settings.apiKey) {
    const provider = S.settings.apiProvider === 'anthropic' ? ((AI_MODELS.find(m => m.id === (S.settings.aiModel || AI_DEFAULT_MODEL)) || {}).label || 'Claude').split(' — ')[0] : 'GPT';
    el.innerHTML = `<span style="color:#22c55e;">●</span> IA ativa (${provider})`;
  } else {
    el.innerHTML = '<span style="color:var(--muted);">●</span> Modo offline';
  }
}

let _chatFileContent = null;
let _chatFileName = null;

function onChatFileSelect(input) {
  const file = input.files[0];
  if (!file) return;
  _chatFileName = file.name;
  const ext = file.name.split('.').pop().toLowerCase();

  if (ext === 'pdf') {
    // PDF: tentar extrair texto com FileReader
    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        // Extrair texto simples do PDF (sem lib externa — pega strings legíveis)
        const bytes = new Uint8Array(e.target.result);
        let text = '';
        // Decode as latin1 for basic text extraction
        for (let i = 0; i < bytes.length; i++) text += String.fromCharCode(bytes[i]);
        // Extract text between parentheses (PDF text objects) and BT/ET blocks
        const matches = text.match(/\(([^)]+)\)/g);
        if (matches) {
          _chatFileContent = matches.map(m => m.slice(1, -1)).join(' ').replace(/\\n/g, '\n');
        } else {
          _chatFileContent = '[PDF sem texto extraível — copie e cole o conteúdo manualmente]';
        }
        showChatFilePreview();
      } catch { _chatFileContent = '[Erro ao ler PDF]'; showChatFilePreview(); }
    };
    reader.readAsArrayBuffer(file);
  } else if (file.type.startsWith('image/')) {
    // Imagem: converter pra base64 pra enviar como vision (se API suportar)
    const reader = new FileReader();
    reader.onload = (e) => {
      _chatFileContent = { type: 'image', base64: e.target.result, mimeType: file.type };
      showChatFilePreview();
    };
    reader.readAsDataURL(file);
  } else {
    // TXT/CSV: ler como texto
    const reader = new FileReader();
    reader.onload = (e) => {
      _chatFileContent = e.target.result;
      showChatFilePreview();
    };
    reader.readAsText(file, 'UTF-8');
  }
  input.value = '';
}

function showChatFilePreview() {
  const preview = document.getElementById('chat-file-preview');
  const nameEl = document.getElementById('chat-file-name');
  if (_chatFileName) {
    const isImage = _chatFileContent && _chatFileContent.type === 'image';
    nameEl.textContent = `📎 ${_chatFileName}` + (isImage ? ' (imagem)' : '');
    preview.style.display = 'block';
  }
}

function clearChatFile() {
  _chatFileContent = null;
  _chatFileName = null;
  document.getElementById('chat-file-preview').style.display = 'none';
}

async function sendChat() {
  const input = document.getElementById('chat-input');
  const msg = input.value.trim();
  const hasFile = !!_chatFileContent;
  if (!msg && !hasFile) return;
  input.value = '';
  const fileContent = _chatFileContent;
  const fileName = _chatFileName;
  clearChatFile();

  const displayMsg = msg + (fileName ? ` 📎 ${fileName}` : '');
  addBubble('user', displayMsg);
  const sugg = document.getElementById('chat-suggestions');
  if (sugg) sugg.style.display = 'none';

  // AI path: use real API if key is configured
  if (S.settings.apiKey && !chatImportMode) {
    const sendBtn = document.getElementById('chat-send');
    sendBtn.disabled = true;
    sendBtn.textContent = '...';
    const typId = addTyping();

    try {
      // Build messages: system + recent history + current
      const systemPrompt = buildSystemPrompt();
      const history = (S.chatHistory || []).slice(-20).map(h => ({ role: h.role === 'ai' ? 'assistant' : h.role, content: h.text }));

      // Build user message content (text + optional file)
      let userContent = msg || '';
      if (fileContent) {
        if (typeof fileContent === 'string') {
          // Text file / PDF extracted text
          userContent = (msg ? msg + '\n\n' : 'Analise e importe este conteúdo:\n\n') + '--- CONTEÚDO DO ARQUIVO (' + fileName + ') ---\n' + fileContent;
        } else if (fileContent.type === 'image') {
          // Image: use vision API (multimodal)
          const provider = S.settings.apiProvider || 'openai';
          if (provider === 'anthropic') {
            const base64Data = fileContent.base64.split(',')[1];
            const mediaType = fileContent.mimeType;
            userContent = [
              { type: 'text', text: msg || 'Analise esta imagem de fatura/extrato e importe as transações.' },
              { type: 'image', source: { type: 'base64', media_type: mediaType, data: base64Data } }
            ];
          } else {
            userContent = [
              { type: 'text', text: msg || 'Analise esta imagem de fatura/extrato e importe as transações.' },
              { type: 'image_url', image_url: { url: fileContent.base64 } }
            ];
          }
        }
      }

      const messages = [
        { role: 'system', content: systemPrompt },
        ...history,
        { role: 'user', content: userContent }
      ];

      let result = await callAI(messages, AI_TOOLS);
      if (!result) {
        // Fallback if callAI returns null (no key)
        removeTyping(typId);
        addBubble('ai', processChat(msg));
        sendBtn.disabled = false;
        sendBtn.textContent = 'Enviar';
        return;
      }

      // Loop de ferramentas: o modelo pode encadear várias consultas antes de responder
      const provider = S.settings.apiProvider || 'openai';
      if (provider === 'anthropic') {
        let convo = messages.filter(m => m.role !== 'system');
        let rounds = 0;
        while (result.toolCalls && result.toolCalls.length > 0 && rounds < 6) {
          rounds++;
          const toolResults = result.toolCalls.map(tc => {
            let out;
            try { out = executeAITool(tc.name, tc.arguments); }
            catch (err) { out = 'Erro ao executar ferramenta: ' + err.message; }
            return { type: 'tool_result', tool_use_id: tc.id, content: String(out) };
          });
          convo = [
            ...convo,
            { role: 'assistant', content: result.content },   // bruto: thinking + text + tool_use
            { role: 'user', content: toolResults }
          ];
          const next = await callAI([{ role: 'system', content: systemPrompt }, ...convo], AI_TOOLS);
          if (!next) break;
          result = next;
        }
      } else if (result.toolCalls && result.toolCalls.length > 0) {
        const toolResults = [];
        for (const tc of result.toolCalls) {
          const toolResult = executeAITool(tc.name, tc.arguments);
          toolResults.push({ id: tc.id, name: tc.name, result: toolResult });
        }
        {
          // OpenAI: append assistant message with tool_calls, then tool results
          const assistantMsg = { role: 'assistant', content: result.text || null, tool_calls: result.toolCalls.map(tc => ({ id: tc.id, type: 'function', function: { name: tc.name, arguments: JSON.stringify(tc.arguments) } })) };
          const toolMsgs = toolResults.map(tr => ({ role: 'tool', tool_call_id: tr.id, content: tr.result }));
          const finalMessages = [...messages, assistantMsg, ...toolMsgs];
          const finalResult = await callAI(finalMessages, AI_TOOLS);
          if (finalResult) result = finalResult;
        }
      }

      removeTyping(typId);
      const reply = result.text || 'Pronto!';
      addBubble('ai', mdToHtml(reply));

      // Save to chat history
      S.chatHistory.push({ role: 'user', text: displayMsg, at: new Date().toISOString() });
      S.chatHistory.push({ role: 'ai', text: reply, at: new Date().toISOString() });
      if (S.chatHistory.length > 100) S.chatHistory = S.chatHistory.slice(-100);
      save();

    } catch(e) {
      removeTyping(typId);
      console.error('Chat AI error:', e);
      addBubble('ai', `Erro ao consultar IA: ${e.message}. Usando modo offline.`);
      addBubble('ai', processChat(msg));
    }

    sendBtn.disabled = false;
    sendBtn.textContent = 'Enviar';
    return;
  }

  // Offline fallback path
  const typId = addTyping();
  setTimeout(() => {
    removeTyping(typId);
    let reply;
    if (chatImportMode) {
      reply = handleImportFlow(msg);
    } else {
      reply = processChat(msg);
    }
    addBubble('ai', reply);
  }, 500);
}

function normalize(str) {
  return str.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();
}

function parseMonthFromMsg(txt) {
  const n = normalize(txt);
  const mesesNorm = MESES_FULL.map(m => normalize(m));
  const mesesAbr = MESES.map(m => normalize(m));
  for (let i = 0; i < 12; i++) {
    if (n.includes(mesesNorm[i]) || n.includes(mesesAbr[i])) return i;
  }
  const numMatch = n.match(/\b(0[1-9]|1[0-2])\b/);
  if (numMatch) return parseInt(numMatch[1]) - 1;
  return null;
}

function parseYearFromMsg(txt) {
  const m = txt.match(/\b(20\d{2})\b/);
  return m ? parseInt(m[1]) : new Date().getFullYear();
}

function getAccName(a) {
  return BANKS[a.bank] ? BANKS[a.bank].name : a.label;
}

function findAccountByName(txt) {
  const n = normalize(txt);
  return S.accounts.find(a => n.includes(normalize(getAccName(a))));
}

function txMonth(tx, month, year) {
  const d = new Date(tx.date + 'T12:00:00');
  return d.getMonth() === month && d.getFullYear() === year;
}

function processChat(message) {
  const n = normalize(message);
  const now = new Date();
  const curMonth = now.getMonth();
  const curYear = now.getFullYear();

  // ── AJUDA ──
  if (n === 'ajuda' || n === 'help' || n.includes('o que voce pode') || n.includes('o que voce faz')) {
    return `Posso te ajudar com várias consultas. Experimente perguntar:<br><br>
<b>Gastos:</b> "quanto gastei em março", "gastos do Nubank", "quanto gastei com alimentação", "maior gasto", "gastos hoje"<br>
<b>Saldo:</b> "saldo", "saldo do Sicredi", "quanto tenho"<br>
<b>Receitas:</b> "receitas de janeiro", "quanto recebi"<br>
<b>Faturas:</b> "fatura do Nubank", "fatura atual", "fatura de março"<br>
<b>Investimentos:</b> "investimentos", "quanto investi"<br>
<b>Dívidas:</b> "dívidas", "quanto devo"<br>
<b>Análise:</b> "categorias", "top gastos", "resumo do mês"<br>
<b>Importar:</b> "importar fatura" para importar extrato de cartão`;
  }

  // ── IMPORTAR ──
  if (n.includes('importar fatura') || n.includes('importar lancamento') || n === 'importar') {
    chatImportMode = true;
    chatImportState = { step: 'awaiting_tipo', card: null, month: null, year: null, user: null, tipo: null, transactions: [] };
    return `Vamos importar! O que deseja importar?<br><br><b>1.</b> Fatura de <b>Crédito</b> (cartão)<br><b>2.</b> Lançamentos de <b>Débito</b> (banco)<br><br>Digite <b>1</b> ou <b>2</b>.`;
  }

  // ── RESUMO DO MÊS ──
  if (n.includes('resumo')) {
    let month = parseMonthFromMsg(message);
    let year = parseYearFromMsg(message);
    if (month === null) month = curMonth;
    const mTx = S.transactions.filter(t => txMonth(t, month, year));
    const rec = mTx.filter(t => t.type === 'receita').reduce((s, t) => s + amountBrl(t), 0);
    const despDeb = mTx.filter(t => t.type === 'despesa' && t.formaPgto !== 'credito').reduce((s, t) => s + amountBrl(t), 0);
    const despCred = mTx.filter(t => t.type === 'despesa' && t.formaPgto === 'credito').reduce((s, t) => s + amountBrl(t), 0);
    const inv = mTx.filter(t => t.type === 'investimento').reduce((s, t) => s + amountBrl(t), 0);
    const saldo = rec - despDeb;
    return `<b>Resumo de ${MESES_FULL[month]} ${year}</b><br><br>
<table style="width:100%;border-collapse:collapse;font-size:13px;">
<tr><td style="padding:6px 0;">Receitas</td><td style="text-align:right;color:#059669;font-weight:700;">${brl(rec)}</td></tr>
<tr><td style="padding:6px 0;">Despesas (débito)</td><td style="text-align:right;color:#e11d48;font-weight:700;">${brl(despDeb)}</td></tr>
<tr><td style="padding:6px 0;">Despesas (crédito)</td><td style="text-align:right;color:#e11d48;font-weight:700;">${brl(despCred)}</td></tr>
<tr><td style="padding:6px 0;">Investimentos</td><td style="text-align:right;color:#2563eb;font-weight:700;">${brl(inv)}</td></tr>
<tr style="border-top:2px solid var(--border);"><td style="padding:8px 0;font-weight:700;">Saldo (rec - desp débito)</td><td style="text-align:right;font-weight:800;color:${saldo >= 0 ? '#059669' : '#e11d48'};">${brl(saldo)}</td></tr>
</table>
${mTx.length === 0 ? '<br><span style="color:var(--muted);">Nenhuma transação neste mês.</span>' : `<br><span style="color:var(--muted);">${mTx.length} transações no total.</span>`}`;
  }

  // ── GASTOS HOJE ──
  if (n.includes('gastei hoje') || n.includes('gastos hoje')) {
    const today = now.toISOString().slice(0, 10);
    const txs = S.transactions.filter(t => t.type === 'despesa' && t.date === today);
    if (txs.length === 0) return 'Nenhuma despesa registrada hoje.';
    const total = txs.reduce((s, t) => s + amountBrl(t), 0);
    let rows = txs.map(t => `<tr><td style="padding:4px 8px 4px 0;">${escapeHtml(t.desc)}</td><td style="padding:4px 0;color:#e11d48;text-align:right;font-weight:600;">${brl(amountBrl(t))}</td></tr>`).join('');
    return `<b>Gastos de hoje:</b><br><br>
<table style="width:100%;border-collapse:collapse;font-size:13px;">${rows}
<tr style="border-top:2px solid var(--border);"><td style="padding:6px 0;font-weight:700;">Total</td><td style="text-align:right;font-weight:800;color:#e11d48;">${brl(total)}</td></tr></table>`;
  }

  // ── MAIOR GASTO ──
  if (n.includes('maior gasto') || n.includes('maior despesa')) {
    let month = parseMonthFromMsg(message);
    let year = parseYearFromMsg(message);
    if (month === null) month = curMonth;
    const txs = S.transactions.filter(t => t.type === 'despesa' && txMonth(t, month, year));
    if (txs.length === 0) return `Nenhuma despesa encontrada em ${MESES_FULL[month]} ${year}.`;
    const sorted = [...txs].sort((a, b) => amountBrl(b) - amountBrl(a));
    const top = sorted[0];
    const acc = S.accounts.find(a => a.id === top.accountId);
    return `<b>Maior gasto de ${MESES_FULL[month]}:</b><br><br>
<span style="font-size:18px;font-weight:800;color:#e11d48;">${brl(amountBrl(top))}</span><br>
<b>${top.desc}</b> — ${top.category}<br>
${fmtDate(top.date)}${acc ? ' | ' + getAccName(acc) : ''}${top.formaPgto === 'credito' ? ' (crédito)' : ''}`;
  }

  // ── QUANTO GASTEI EM [MES] ──
  if ((n.includes('quanto gastei') || n.includes('gastos de') || n.includes('gastos em') || n.includes('gastos do mes')) && !n.includes('com ') && !n.includes('no ') && !n.includes('na ') && !n.includes('hoje')) {
    let month = parseMonthFromMsg(message);
    let year = parseYearFromMsg(message);
    if (month === null) month = curMonth;
    const txs = S.transactions.filter(t => t.type === 'despesa' && txMonth(t, month, year));
    const total = txs.reduce((s, t) => s + amountBrl(t), 0);
    const deb = txs.filter(t => t.formaPgto !== 'credito').reduce((s, t) => s + amountBrl(t), 0);
    const cred = txs.filter(t => t.formaPgto === 'credito').reduce((s, t) => s + amountBrl(t), 0);
    return `Em <b>${MESES_FULL[month]} ${year}</b>, você gastou:<br><br>
<table style="width:100%;border-collapse:collapse;font-size:13px;">
<tr><td style="padding:4px 0;">Débito</td><td style="text-align:right;color:#e11d48;font-weight:600;">${brl(deb)}</td></tr>
<tr><td style="padding:4px 0;">Crédito</td><td style="text-align:right;color:#e11d48;font-weight:600;">${brl(cred)}</td></tr>
<tr style="border-top:2px solid var(--border);"><td style="padding:6px 0;font-weight:700;">Total</td><td style="text-align:right;font-weight:800;color:#e11d48;">${brl(total)}</td></tr>
</table>
<br><span style="color:var(--muted);">${txs.length} despesas no total.</span>`;
  }

  // ── QUANTO GASTEI NO [BANCO/CARTAO] ──
  if ((n.includes('gastei no') || n.includes('gastei na') || n.includes('gastos do') || n.includes('gastos da')) && !n.includes('mes')) {
    const acc = findAccountByName(message);
    if (!acc) return 'Não encontrei essa conta/cartão. As contas cadastradas são: ' + S.accounts.map(a => '<b>' + a.name + '</b>').join(', ');
    let month = parseMonthFromMsg(message);
    let year = parseYearFromMsg(message);
    if (month === null) month = curMonth;
    const txs = S.transactions.filter(t => t.type === 'despesa' && t.accountId === acc.id && txMonth(t, month, year));
    const total = txs.reduce((s, t) => s + amountBrl(t), 0);
    return `Gastos no <b>${getAccName(acc)}</b> em ${MESES_FULL[month]} ${year}:<br><br>
<span style="font-size:18px;font-weight:800;color:#e11d48;">${brl(total)}</span><br>
<span style="color:var(--muted);">${txs.length} transações.</span>`;
  }

  // ── QUANTO GASTEI COM [CATEGORIA] ──
  if (n.includes('gastei com') || n.includes('gastos com')) {
    let month = parseMonthFromMsg(message);
    let year = parseYearFromMsg(message);
    if (month === null) month = curMonth;
    // Extract category name: after "com"
    const comIdx = n.indexOf('com ');
    const catQuery = comIdx >= 0 ? n.substring(comIdx + 4).replace(/[?.!]/g, '').trim() : '';
    const allCats = Object.keys(CATS_DESPESA).map(c => ({ name: c, norm: normalize(c) }));
    const found = allCats.find(c => c.norm.includes(catQuery) || catQuery.includes(c.norm));
    if (!found) return `Não encontrei a categoria "<b>${catQuery}</b>". Categorias disponíveis: ${Object.keys(CATS_DESPESA).join(', ')}`;
    const txs = S.transactions.filter(t => t.type === 'despesa' && normalize(t.category) === found.norm && txMonth(t, month, year));
    const total = txs.reduce((s, t) => s + amountBrl(t), 0);
    let rows = txs.slice(0, 10).map(t => `<tr><td style="padding:3px 8px 3px 0;">${fmtDate(t.date)}</td><td style="padding:3px 0;">${escapeHtml(t.desc)}</td><td style="text-align:right;color:#e11d48;font-weight:600;">${brl(amountBrl(t))}</td></tr>`).join('');
    return `Gastos com <b>${found.name}</b> em ${MESES_FULL[month]} ${year}:<br><br>
<span style="font-size:18px;font-weight:800;color:#e11d48;">${brl(total)}</span> (${txs.length} transações)<br><br>
${txs.length > 0 ? '<table style="width:100%;border-collapse:collapse;font-size:13px;">' + rows + '</table>' : ''}
${txs.length > 10 ? '<br><span style="color:var(--muted);">Mostrando 10 de ' + txs.length + '.</span>' : ''}`;
  }

  // ── SALDO ──
  if (n.includes('saldo') || n.includes('quanto tenho')) {
    const specAcc = findAccountByName(message);
    if (specAcc && !n.includes('quanto tenho')) {
      // Saldo de conta específica
      const accTx = S.transactions.filter(t => t.accountId === specAcc.id && t.formaPgto !== 'credito');
      const rec = accTx.filter(t => t.type === 'receita').reduce((s, t) => s + amountBrl(t), 0);
      const desp = accTx.filter(t => t.type === 'despesa').reduce((s, t) => s + amountBrl(t), 0);
      const saldo = rec - desp + (specAcc.initialBalance || 0);
      return `Saldo do <b>${getAccName(specAcc)}</b>:<br><br>
<span style="font-size:20px;font-weight:800;color:${saldo >= 0 ? '#059669' : '#e11d48'};">${brl(saldo)}</span>`;
    }
    // Saldo geral de todos os bancos
    const bancos = S.accounts.filter(a => a.accountType !== 'cartao');
    if (bancos.length === 0) return 'Nenhuma conta cadastrada.';
    let rows = '';
    let totalSaldo = 0;
    bancos.forEach(acc => {
      const accTx = S.transactions.filter(t => t.accountId === acc.id && t.formaPgto !== 'credito');
      const rec = accTx.filter(t => t.type === 'receita').reduce((s, t) => s + amountBrl(t), 0);
      const desp = accTx.filter(t => t.type === 'despesa').reduce((s, t) => s + amountBrl(t), 0);
      const saldo = rec - desp + (acc.initialBalance || 0);
      totalSaldo += saldo;
      rows += `<tr><td style="padding:5px 0;">${getAccName(acc)}</td><td style="text-align:right;font-weight:700;color:${saldo >= 0 ? '#059669' : '#e11d48'};">${brl(saldo)}</td></tr>`;
    });
    return `<b>Saldo dos bancos:</b><br><br>
<table style="width:100%;border-collapse:collapse;font-size:13px;">${rows}
<tr style="border-top:2px solid var(--border);"><td style="padding:8px 0;font-weight:700;">Total</td><td style="text-align:right;font-weight:800;color:${totalSaldo >= 0 ? '#059669' : '#e11d48'};">${brl(totalSaldo)}</td></tr></table>`;
  }

  // ── RECEITAS ──
  if (n.includes('receita') || n.includes('quanto recebi')) {
    let month = parseMonthFromMsg(message);
    let year = parseYearFromMsg(message);
    if (month === null) month = curMonth;
    const txs = S.transactions.filter(t => t.type === 'receita' && txMonth(t, month, year));
    const total = txs.reduce((s, t) => s + amountBrl(t), 0);
    let rows = txs.map(t => `<tr><td style="padding:3px 8px 3px 0;">${fmtDate(t.date)}</td><td style="padding:3px 0;">${escapeHtml(t.desc)}</td><td style="text-align:right;color:#059669;font-weight:600;">${brl(amountBrl(t))}</td></tr>`).join('');
    return `<b>Receitas de ${MESES_FULL[month]} ${year}:</b><br><br>
${txs.length > 0 ? '<table style="width:100%;border-collapse:collapse;font-size:13px;">' + rows + '<tr style="border-top:2px solid var(--border);"><td style="padding:6px 0;font-weight:700;" colspan="2">Total</td><td style="text-align:right;font-weight:800;color:#059669;">' + brl(total) + '</td></tr></table>' : '<span style="color:var(--muted);">Nenhuma receita encontrada neste mês.</span>'}`;
  }

  // ── FATURA ──
  if (n.includes('fatura')) {
    const specAcc = findAccountByName(message);
    let month = parseMonthFromMsg(message);
    let year = parseYearFromMsg(message);
    if (month === null) month = curMonth;
    const fatRef = `${year}-${String(month + 1).padStart(2, '0')}`;
    let txs = S.transactions.filter(t => t.type === 'despesa' && t.formaPgto === 'credito');
    if (specAcc) txs = txs.filter(t => t.accountId === specAcc.id);
    txs = txs.filter(t => getTxFaturaRef(t) === fatRef);
    const total = txs.reduce((s, t) => s + amountBrl(t), 0);
    const cardLabel = specAcc ? getAccName(specAcc) : 'todos os cartões';
    let rows = txs.slice(0, 15).map(t => {
      const ca = S.accounts.find(a => a.id === t.accountId);
      return `<tr><td style="padding:3px 8px 3px 0;font-size:12px;color:var(--muted);">${fmtDate(t.date)}</td><td style="padding:3px 0;">${escapeHtml(t.desc)}</td>${!specAcc && ca ? '<td style="padding:3px 4px;font-size:11px;color:#6366f1;">' + ca.name + '</td>' : ''}<td style="text-align:right;color:#e11d48;font-weight:600;">${brl(amountBrl(t))}</td></tr>`;
    }).join('');
    return `<b>Fatura de ${MESES_FULL[month]} ${year}</b> (${cardLabel}):<br><br>
<span style="font-size:20px;font-weight:800;color:#e11d48;">${brl(total)}</span> <span style="color:var(--muted);">(${txs.length} lançamentos)</span><br><br>
${txs.length > 0 ? '<table style="width:100%;border-collapse:collapse;font-size:13px;">' + rows + '</table>' : ''}
${txs.length > 15 ? '<br><span style="color:var(--muted);">Mostrando 15 de ' + txs.length + '.</span>' : ''}`;
  }

  // ── INVESTIMENTOS ──
  if (n.includes('investimento') || n.includes('quanto investi')) {
    const txs = S.transactions.filter(t => t.type === 'investimento');
    const total = txs.reduce((s, t) => s + amountBrl(t), 0);
    const byCat = {};
    txs.forEach(t => byCat[t.category] = (byCat[t.category] || 0) + amountBrl(t));
    let rows = Object.entries(byCat).sort((a, b) => b[1] - a[1]).map(([cat, val]) =>
      `<tr><td style="padding:4px 0;">${escapeHtml(cat)}</td><td style="text-align:right;color:#2563eb;font-weight:600;">${brl(val)}</td></tr>`
    ).join('');
    return `<b>Investimentos:</b><br><br>
<span style="font-size:20px;font-weight:800;color:#2563eb;">${brl(total)}</span> total investido<br><br>
${rows ? '<table style="width:100%;border-collapse:collapse;font-size:13px;">' + rows + '</table>' : '<span style="color:var(--muted);">Nenhum investimento registrado.</span>'}
<br><span style="color:var(--muted);">${txs.length} aportes no total.</span>`;
  }

  // ── DÍVIDAS ──
  if (n.includes('divida') || n.includes('quanto devo')) {
    if (!S.debts || S.debts.length === 0) return 'Nenhuma dívida cadastrada.';
    let totalDevido = 0;
    let rows = S.debts.map(d => {
      const remaining = (d.total || 0) - (d.paid || 0);
      totalDevido += remaining;
      return `<tr><td style="padding:5px 0;"><b>${d.name || d.desc || 'Dívida'}</b></td><td style="text-align:right;color:#e11d48;font-weight:600;">${brl(remaining)}</td></tr>`;
    }).join('');
    return `<b>Suas dívidas:</b><br><br>
<table style="width:100%;border-collapse:collapse;font-size:13px;">${rows}
<tr style="border-top:2px solid var(--border);"><td style="padding:8px 0;font-weight:700;">Total devendo</td><td style="text-align:right;font-weight:800;color:#e11d48;">${brl(totalDevido)}</td></tr></table>`;
  }

  // ── TOP CATEGORIAS / POR CATEGORIA ──
  if (n.includes('categoria') || n.includes('top gasto') || n.includes('mais gastei') || n.includes('top categorias')) {
    let month = parseMonthFromMsg(message);
    let year = parseYearFromMsg(message);
    if (month === null) month = curMonth;
    const txs = S.transactions.filter(t => t.type === 'despesa' && txMonth(t, month, year));
    const byCat = {};
    txs.forEach(t => byCat[t.category] = (byCat[t.category] || 0) + amountBrl(t));
    const sorted = Object.entries(byCat).sort((a, b) => b[1] - a[1]);
    const total = txs.reduce((s, t) => s + amountBrl(t), 0);
    const top = sorted.slice(0, (n.includes('top') || n.includes('mais gastei')) ? 5 : 999);
    if (top.length === 0) return `Nenhuma despesa em ${MESES_FULL[month]} ${year}.`;
    let rows = top.map(([cat, val]) => {
      const pct = total > 0 ? ((val / total) * 100).toFixed(1) : 0;
      return `<tr><td style="padding:5px 0;">${escapeHtml(cat)}</td><td style="text-align:right;color:#e11d48;font-weight:600;">${brl(val)}</td><td style="text-align:right;color:var(--muted);font-size:12px;padding-left:8px;">${pct}%</td></tr>`;
    }).join('');
    const title = (n.includes('top') || n.includes('mais gastei')) ? 'Top 5 categorias' : 'Despesas por categoria';
    return `<b>${title} — ${MESES_FULL[month]} ${year}:</b><br><br>
<table style="width:100%;border-collapse:collapse;font-size:13px;">${rows}
<tr style="border-top:2px solid var(--border);"><td style="padding:6px 0;font-weight:700;">Total</td><td style="text-align:right;font-weight:800;color:#e11d48;">${brl(total)}</td><td></td></tr></table>`;
  }

  // ── FALLBACK ──
  return 'Desculpe, não entendi sua pergunta. Digite <b>ajuda</b> para ver o que posso fazer!';
}

// ── AUTO CATEGORIZE ──
const AUTO_CAT_RULES = [
  // Alimentação
  {p: /ifood|rappi|uber\s*eat/i, cat:'Alimentação', sub:'Delivery'},
  {p: /lanchonete|sorvete|chiquinho|tapioca|bapka|eskimo/i, cat:'Alimentação', sub:'Lanchonete/Bar'},
  {p: /mercado|mercearia|atacado|condor|festval|portao|gigante|circuito|sam.s.club/i, cat:'Mercado', sub:''},
  {p: /restaurante|outback|spoleto|pizza|burger|acapulco|campodoro|empada|taisho|gelasko|saborelly|frango/i, cat:'Alimentação', sub:'Restaurante'},
  // Assinaturas
  {p: /netflix|spotify|youtube|disney|hbo|prime|crunchyroll/i, cat:'Assinaturas', sub:'Streaming'},
  {p: /microsoft|canva|icloud|anthropic|cursor|chatgpt|gemini|discord|duolingo|ai\s*mirror/i, cat:'Assinaturas', sub:'Software/App'},
  {p: /xbox|playstation|steam|game/i, cat:'Games', sub:''},
  {p: /nubank\+|serasa|claro|livelo|twitch/i, cat:'Assinaturas', sub:'Outros'},
  // Saúde
  {p: /farmacia|drogaria|nissei|droga|morifarma/i, cat:'Saúde', sub:'Farmácia'},
  {p: /academia|growth|fitness/i, cat:'Saúde', sub:'Academia/Fitness'},
  {p: /medico|cardio|clinica|hospital|unimed|inc\s/i, cat:'Saúde', sub:'Médico'},
  {p: /hormonio|manipulado/i, cat:'Saúde', sub:'Geral'},
  // Carro
  {p: /posto|combusti|alphaville|gasolina|etanol/i, cat:'Carro', sub:'Combustível'},
  {p: /nutag|estaciona|park|kummer/i, cat:'Carro', sub:'Estacionamento'},
  {p: /ipva|licencia|detran/i, cat:'Carro', sub:'IPVA/Licenciamento'},
  {p: /lavacar|auto\s*eletrica|bello\s*auto/i, cat:'Carro', sub:'Manutenção'},
  // Casa
  {p: /copel|energia|luz/i, cat:'Casa', sub:'Luz'},
  {p: /nio|internet/i, cat:'Casa', sub:'Internet'},
  {p: /sanepar|agua/i, cat:'Casa', sub:'Água'},
  {p: /aluguel/i, cat:'Casa', sub:'Aluguel'},
  // Outros
  {p: /shopee|shoppee|shoppe|mercado\s*livre|temu/i, cat:'Outros', sub:''},
  {p: /presente|loccitane/i, cat:'Presente', sub:''},
  {p: /vestuario|roupa|brutal|centauro|parana\s*clube/i, cat:'Vestuário', sub:''},
  {p: /barbeiro/i, cat:'Lazer', sub:'Hobbies'},
  {p: /hotel|pousada|decolar/i, cat:'Viagem', sub:'Hospedagem'},
  {p: /pedagio/i, cat:'Viagem', sub:'Transporte'},
  {p: /consorcio/i, cat:'Outros', sub:''},
  {p: /educacao|curso/i, cat:'Educação', sub:'Cursos'},
];

function autoCategorizeTx(desc) {
  const d = desc.toLowerCase();
  for (const rule of AUTO_CAT_RULES) {
    if (rule.p.test(d)) {
      const custoTipo = ['Assinaturas','Casa','Mercado','Saúde'].includes(rule.cat) ? 'fixo' : 'variavel';
      return { category: rule.cat, subcategory: rule.sub, custoTipo };
    }
  }
  return { category: 'Outros', subcategory: '', custoTipo: 'variavel' };
}

// ── IMPORT FLOW ──
function handleImportFlow(msg) {
  const n = normalize(msg);
  const st = chatImportState;

  if (n === 'cancelar' || n === 'cancel') {
    chatImportMode = false;
    chatImportState = { step: null, card: null, month: null, year: null, user: null, tipo: null, transactions: [] };
    return 'Importação cancelada.';
  }

  // Step 1: tipo (crédito ou débito)
  if (st.step === 'awaiting_tipo') {
    if (n.includes('credit') || n.includes('cartao') || n.includes('cartão') || n === '1') {
      st.tipo = 'credito';
      st.step = 'awaiting_card';
      const cards = S.accounts.filter(a => a.accountType === 'cartao');
      const list = cards.map((c, i) => `${i + 1}. ${getAccName(c)} — ${escapeHtml(c.label)}`).join('<br>');
      return `Qual <b>cartão</b>?<br><br>${list}<br><br>Digite o número ou nome.`;
    } else if (n.includes('debit') || n.includes('banco') || n.includes('conta') || n === '2') {
      st.tipo = 'debito';
      st.step = 'awaiting_card';
      const contas = S.accounts.filter(a => a.accountType !== 'cartao');
      const list = contas.map((c, i) => `${i + 1}. ${getAccName(c)} — ${escapeHtml(c.label)}`).join('<br>');
      return `Qual <b>conta bancária</b>?<br><br>${list}<br><br>Digite o número ou nome.`;
    }
    return 'Digite <b>1</b> para Crédito (cartão) ou <b>2</b> para Débito (banco).';
  }

  // Step 2: card/conta
  if (st.step === 'awaiting_card') {
    const accs = st.tipo === 'credito'
      ? S.accounts.filter(a => a.accountType === 'cartao')
      : S.accounts.filter(a => a.accountType !== 'cartao');
    let found = null;
    const num = parseInt(msg);
    if (!isNaN(num) && num >= 1 && num <= accs.length) {
      found = accs[num - 1];
    } else {
      found = accs.find(c => normalize(getAccName(c)).includes(n) || normalize(c.label).includes(n));
    }
    if (!found) return 'Não encontrei. Tente novamente com o número ou nome.';
    st.card = found;
    st.step = 'awaiting_month';
    const label = st.tipo === 'credito' ? 'fatura' : 'mês de referência';
    return `${getAccName(found)} — ${found.label}. Qual o <b>${label}</b>? (ex: "março 2026", "03/2026")`;
  }

  // Step 3: mês/ano
  if (st.step === 'awaiting_month') {
    let month = parseMonthFromMsg(msg);
    let year = parseYearFromMsg(msg);
    if (month === null) return 'Não entendi o mês. Tente algo como "março 2026" ou "03/2026".';
    st.month = month;
    st.year = year;
    st.step = 'awaiting_user';
    return `<b>${MESES_FULL[month]} ${year}</b>. Qual o <b>titular</b>?<br><br>1. ${S.settings.u1}<br>2. ${S.settings.u2}<br><br>Digite o número ou nome.`;
  }

  // Step 4: titular
  if (st.step === 'awaiting_user') {
    if (n.includes(normalize(S.settings.u1)) || n === '1') {
      st.user = S.settings.u1;
    } else if (n.includes(normalize(S.settings.u2)) || n === '2') {
      st.user = S.settings.u2;
    } else {
      return `Não entendi. Digite <b>1</b> para ${S.settings.u1} ou <b>2</b> para ${S.settings.u2}.`;
    }
    st.step = 'awaiting_text';
    return `Titular: <b>${st.user}</b>.<br><br>Agora <b>cole os lançamentos</b>. Formatos aceitos:<br>
<code>05 MAR SPOTIFY 31.90</code><br>
<code>05/03 SPOTIFY R$ 31,90</code><br>
<code>SPOTIFY 31,90</code><br>
<code>05/03;SPOTIFY;31,90</code> (CSV)<br><br>
Cole tudo de uma vez. Digite <b>cancelar</b> para abortar.`;
  }

  // Step 5: parse text
  if (st.step === 'awaiting_text') {
    const parsed = parseFaturaText(msg, st.month, st.year);
    if (parsed.length === 0) return 'Não consegui extrair nenhuma transação. Verifique o formato e tente novamente.';
    st.transactions = parsed;
    st.step = 'awaiting_confirm';
    const total = parsed.reduce((s, t) => s + t.amount, 0);
    let rows = parsed.map(t => {
      const catInfo = autoCategorizeTx(t.desc);
      const catLabel = catInfo.category + (catInfo.subcategory ? ' › ' + catInfo.subcategory : '');
      const parcelaTag = t.parcela ? ` <span style="color:#f59e0b;font-size:11px;">(${t.parcela})</span>` : '';
      return `<tr><td style="padding:3px 0;font-size:12px;color:var(--muted);">${fmtDate(t.date)}</td><td style="padding:3px 8px;">${escapeHtml(t.desc)}${parcelaTag}</td><td style="font-size:11px;color:#6366f1;">${catLabel}</td><td style="text-align:right;color:#e11d48;font-weight:600;">${brl(t.amount)}</td></tr>`;
    }).join('');
    return `Encontrei <b>${parsed.length} transações</b> (categorias auto-detectadas):<br><br>
<table style="width:100%;border-collapse:collapse;font-size:13px;">${rows}
<tr style="border-top:2px solid var(--border);"><td colspan="3" style="padding:6px 0;font-weight:700;">Total</td><td style="text-align:right;font-weight:800;color:#e11d48;">${brl(total)}</td></tr></table>
<br>${getAccName(st.card)} | ${MESES_FULL[st.month]} ${st.year} | Titular: ${st.user}<br><br>
Digite <b>confirmar</b> para importar ou <b>cancelar</b> para desistir.`;
  }

  // Step 6: confirm
  if (st.step === 'awaiting_confirm') {
    if (n.includes('confirmar') || n.includes('sim') || n === 'ok') {
      const fatRef = `${st.year}-${String(st.month + 1).padStart(2, '0')}`;
      const isCredito = st.tipo === 'credito';
      const count = st.transactions.length;
      st.transactions.forEach(t => {
        const catInfo = autoCategorizeTx(t.desc);
        S.transactions.push({
          id: Date.now().toString(36) + Math.random().toString(36).substr(2, 5),
          type: 'despesa',
          desc: t.desc + (t.parcela ? ` (${t.parcela})` : ''),
          amount: t.amount,
          date: t.date,
          category: catInfo.category,
          subcategory: catInfo.subcategory,
          accountId: st.card.id,
          user: st.user,
          formaPgto: isCredito ? 'credito' : 'debito',
          faturaRef: isCredito ? fatRef : undefined,
          pago: !isCredito,
          custoTipo: catInfo.custoTipo,
          parcela: t.parcela || null,
          currency: 'BRL',
          isNegative: false,
          at: new Date().toISOString(),
          updatedAt: new Date().toISOString()
        });
      });
      save();
      chatImportMode = false;
      chatImportState = { step: null, card: null, month: null, year: null, user: null, tipo: null, transactions: [] };
      const catCount = new Set(st.transactions.map(t => autoCategorizeTx(t.desc).category)).size;
      return `Importação concluída! <b>${count} transações</b> adicionadas em <b>${catCount} categorias</b>.<br><br>Você pode ajustar as categorias no <b>Histórico</b> se necessário.`;
    }
    return 'Digite <b>confirmar</b> para importar ou <b>cancelar</b> para desistir.';
  }

  chatImportMode = false;
  return 'Algo deu errado. Tente novamente digitando <b>importar fatura</b>.';
}

function parseFaturaText(text, refMonth, refYear) {
  const mesesAbr = ['JAN','FEV','MAR','ABR','MAI','JUN','JUL','AGO','SET','OUT','NOV','DEZ'];
  const mesesPat = mesesAbr.join('|');

  // Pre-process: split glued lines. When text has no newlines but contains
  // patterns like "30,68 05 FEV" (amount followed by date), split there.
  let processed = text;
  // Split before "DD MES" pattern (e.g. "30,68 05 FEV" → "30,68\n05 FEV")
  processed = processed.replace(/(\d[\d.,]+)\s+(\d{1,2}\s+(?:JAN|FEV|MAR|ABR|MAI|JUN|JUL|AGO|SET|OUT|NOV|DEZ)\s)/gi, '$1\n$2');
  // Split before "DD/MM" pattern (e.g. "30,68 05/03" → "30,68\n05/03")
  processed = processed.replace(/(\d[\d.,]+)\s+(\d{1,2}\/\d{1,2}\s)/gi, '$1\n$2');

  const lines = processed.split('\n').map(l => l.trim()).filter(l => l.length > 3);
  const results = [];

  for (const line of lines) {
    let date = null, desc = '', amount = null;

    // CSV format: "05/03;SPOTIFY;31,90"
    const csvParts = line.split(/[;\t]/).map(p => p.trim());
    if (csvParts.length >= 3) {
      const dateP = csvParts[0].match(/^(\d{1,2})[\/\-](\d{1,2})(?:[\/\-](\d{2,4}))?$/);
      if (dateP) {
        const day = parseInt(dateP[1]);
        const mon = parseInt(dateP[2]) - 1;
        const yr = dateP[3] ? (dateP[3].length === 2 ? 2000 + parseInt(dateP[3]) : parseInt(dateP[3])) : refYear;
        date = `${yr}-${String(mon + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
        desc = csvParts.slice(1, -1).join(' ').trim();
        const valStr = csvParts[csvParts.length - 1].replace(/R\$\s*/i, '').replace(/\./g, '').replace(',', '.');
        amount = parseFloat(valStr);
      }
    }

    // Format: "15 JAN EBN*SONYPLAYSTATN CU Parcela 04/04 30,68"
    // Bradesco format with "Parcela XX/XX" before amount
    if (!date) {
      let m = line.match(/^(\d{1,2})\s+(JAN|FEV|MAR|ABR|MAI|JUN|JUL|AGO|SET|OUT|NOV|DEZ)\s+(.+?)\s+(?:R\$\s*)?(\d[\d.,]*)\s*$/i);
      if (m) {
        const day = parseInt(m[1]);
        const mi = mesesAbr.indexOf(m[2].toUpperCase());
        date = `${refYear}-${String(mi + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
        desc = m[3].trim();
        amount = parseFloat(m[4].replace(/\./g, '').replace(',', '.'));
      }
    }

    // Format: "05/03 SPOTIFY R$ 31,90"
    if (!date) {
      let m = line.match(/^(\d{1,2})[\/\-](\d{1,2})(?:[\/\-](\d{2,4}))?\s+(.+?)\s+(?:R\$\s*)?(\d[\d.,]*)\s*$/i);
      if (m) {
        const day = parseInt(m[1]);
        const mon = parseInt(m[2]) - 1;
        const yr = m[3] ? (m[3].length === 2 ? 2000 + parseInt(m[3]) : parseInt(m[3])) : refYear;
        date = `${yr}-${String(mon + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
        desc = m[4].trim();
        amount = parseFloat(m[5].replace(/\./g, '').replace(',', '.'));
      }
    }

    // Format: "IFOOD MAR.2026 R$ 30,00" or "IFOOD MAR 2026 30,00" or "IFOOD MAR.2026 30,00"
    // Description + month abbreviation + optional year + amount (no specific day → use day 01)
    if (!date) {
      let m = line.match(/^(.+?)\s+(JAN|FEV|MAR|ABR|MAI|JUN|JUL|AGO|SET|OUT|NOV|DEZ)[.\s\/\-]*(\d{4})?\s+(?:R\$\s*)?(\d[\d.,]*[.,]\d{2})\s*$/i);
      if (m) {
        const mi = mesesAbr.indexOf(m[2].toUpperCase());
        const yr = m[3] ? parseInt(m[3]) : refYear;
        date = `${yr}-${String(mi + 1).padStart(2, '0')}-01`;
        desc = m[1].trim();
        amount = parseFloat(m[4].replace(/\./g, '').replace(',', '.'));
      }
    }

    // Format: "SPOTIFY 31,90" (no date)
    if (!date) {
      let m = line.match(/^(.+?)\s+(?:R\$\s*)?(\d[\d.,]*[.,]\d{2})\s*$/);
      if (m) {
        date = `${refYear}-${String(refMonth + 1).padStart(2, '0')}-01`;
        desc = m[1].trim();
        amount = parseFloat(m[2].replace(/\./g, '').replace(',', '.'));
      }
    }

    // Clean up desc: remove trailing "Parcela" info but keep it for reference
    let parcela = null;
    if (desc) {
      const pMatch = desc.match(/\s+(?:Parcela|Parc\.?)\s*(\d{1,2}\/\d{1,2})\s*$/i);
      if (pMatch) {
        parcela = pMatch[1];
        desc = desc.replace(pMatch[0], '').trim();
      }
    }

    if (date && desc && amount && amount > 0) {
      const entry = { date, desc, amount };
      if (parcela) entry.parcela = parcela;
      results.push(entry);
    }
  }
  return results;
}

// Markdown leve → HTML (negrito, itálico, código, listas, títulos). Escapa HTML antes.
function mdToHtml(md) {
  const esc = String(md || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const lines = esc.split('\n');
  let html = '', inList = false;
  const inline = t => t
    .replace(/`([^`]+)`/g, '<code style="background:var(--surface-2);padding:1px 5px;border-radius:4px;font-size:12px;">$1</code>')
    .replace(/\*\*([^*]+)\*\*/g, '<b>$1</b>')
    .replace(/(^|[^*])\*([^*\n]+)\*(?!\*)/g, '$1<i>$2</i>');
  for (const raw of lines) {
    const line = raw.trimEnd();
    const li = /^\s*[-•*]\s+(.*)$/.exec(line);
    if (li) {
      if (!inList) { html += '<ul style="margin:6px 0 6px 18px;padding:0;list-style:disc;">'; inList = true; }
      html += '<li style="margin:2px 0;">' + inline(li[1]) + '</li>';
      continue;
    }
    if (inList) { html += '</ul>'; inList = false; }
    const h = /^#{1,3}\s+(.*)$/.exec(line);
    if (h) { html += '<div style="font-weight:700;margin:8px 0 4px;">' + inline(h[1]) + '</div>'; continue; }
    if (line === '') { html += '<div style="height:6px;"></div>'; continue; }
    html += '<div>' + inline(line) + '</div>';
  }
  if (inList) html += '</ul>';
  return html;
}

function addBubble(role, text) {
  const c = document.getElementById('chat-msgs');
  const div = document.createElement('div');
  div.style.cssText = 'display:flex;gap:10px;align-items:flex-start;' + (role === 'user' ? 'justify-content:flex-end;' : '');
  if (role === 'user') {
    const txt = text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\n/g, '<br>');
    div.innerHTML = `<div class="bubble-user">${txt}</div>`;
  } else {
    div.innerHTML = `<div style="width:32px;height:32px;background:var(--tint-indigo);border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:16px;flex-shrink:0;">🤖</div><div class="bubble-ai">${text}</div>`;
  }
  c.appendChild(div);
  c.scrollTop = c.scrollHeight;
}

function addTyping() {
  const id = 'typ-' + Date.now();
  const c = document.getElementById('chat-msgs');
  const div = document.createElement('div');
  div.id = id;
  div.style.cssText = 'display:flex;gap:10px;align-items:flex-start;';
  div.innerHTML = `<div style="width:32px;height:32px;background:var(--tint-indigo);border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:16px;flex-shrink:0;">🤖</div>
    <div class="bubble-ai"><span class="dot"></span><span class="dot"></span><span class="dot"></span></div>`;
  c.appendChild(div);
  c.scrollTop = c.scrollHeight;
  return id;
}

function removeTyping(id) {
  const el = document.getElementById(id);
  if (el) el.remove();
}

function askSugg(q) {
  document.getElementById('chat-input').value = q;
  sendChat();
}
