// FinançasCasal — 12b-vencimentos.js
// Módulo carregado em ordem por index.html (escopo global compartilhado). Não use import/export.

// ─── VENCIMENTOS (grade "contas do mês") ──────────────────────────────────────
// Igual à planilha da Thayse: uma linha por conta fixa, uma coluna por mês, dia do vencimento
// ao lado, verde = pago. Nada é digitado na grade: os valores vêm dos lançamentos. Cada conta
// tem uma regra que diz quais lançamentos são dela:
//   tipo 'cartao' → a fatura inteira do cartão (mês = mês de vencimento da fatura);
//   tipo 'lanc'   → despesas no débito do titular que batem com texto (palavras separadas por |)
//                   e/ou categoria/subcategoria. Crédito fica de fora (já está na fatura).
// S.contasFixas = [{ id, nome, dia, user, tipo, cartaoId, texto, cat, sub }]

let vencUser = null;
let vencAno = new Date().getFullYear();
let vencSel = null;      // { id, ym } célula aberta no detalhe
let vencEditId = null;   // conta aberta no formulário
let vencFormAberto = false;

function getContasFixas() {
  if (!Array.isArray(S.contasFixas)) S.contasFixas = [];
  return S.contasFixas;
}

function vencNorm(s) {
  return (s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase().replace(/\s+/g, ' ');
}

function vencDiaVenc(conta, ym) {
  if (!conta.dia) return null;
  const [y, m] = ym.split('-').map(Number);
  const ult = new Date(y, m, 0).getDate();
  return `${ym}-${String(Math.min(parseInt(conta.dia), ult)).padStart(2, '0')}`;
}

function vencLancDaConta(conta, ym) {
  const palavras = (conta.texto || '').split('|').map(vencNorm).map(p => p.trim()).filter(Boolean);
  return S.transactions.filter(t => t && t.type === 'despesa' && !t.isTransfer && t.formaPgto !== 'credito'
    && t.user === conta.user && (t.date || '').startsWith(ym)
    && (!conta.cat || t.category === conta.cat)
    && (!conta.sub || t.subcategory === conta.sub)
    && (!palavras.length || palavras.some(p => vencNorm(t.desc).includes(p))));
}

// → { valor, total, pago, status: vazio|pago|parcial|aberto|atrasado, txs, pagamentos, venc }
function vencCelula(conta, ym) {
  const hoje = new Date().toISOString().slice(0, 10);
  const venc = vencDiaVenc(conta, ym);
  const vencido = venc && hoje > venc;
  if (conta.tipo === 'cartao') {
    const txs = getCreditTxByFatura(ym, conta.cartaoId);
    const total = txs.reduce((s, t) => s + amountBrl(t), 0);
    const [y, m] = ym.split('-').map(Number);
    const pagamentos = getPagamentosFatura(conta.cartaoId, m - 1, y);
    const pago = pagamentos.reduce((s, p) => s + Math.abs(p.amount || 0), 0);
    if (total < 0.005 && pago < 0.005) return { status: 'vazio', valor: 0, total, pago, txs, pagamentos, venc };
    if (pago > 0) return { status: pago >= total - 0.01 ? 'pago' : 'parcial', valor: pago, total, pago, txs, pagamentos, venc };
    const semRegistro = txs.length && txs.every(t => t.pago === true);   // fatura importada já quitada
    return { status: semRegistro ? 'pago' : (vencido ? 'atrasado' : 'aberto'), valor: total, total, pago, txs, pagamentos, venc };
  }
  const txs = vencLancDaConta(conta, ym);
  if (!txs.length) return { status: 'vazio', valor: 0, total: 0, pago: 0, txs, pagamentos: [], venc };
  const total = txs.reduce((s, t) => s + amountBrl(t), 0);
  const pagos = txs.every(t => t.pago !== false);
  return { status: pagos ? 'pago' : (vencido ? 'atrasado' : 'aberto'), valor: total, total, pago: pagos ? total : 0, txs, pagamentos: [], venc };
}

const VENC_COR = { pago: '#059669', parcial: '#d97706', aberto: 'var(--text)', atrasado: '#e11d48', vazio: 'var(--muted)' };

function renderVencimentos() {
  const { u1, u2 } = S.settings;
  if (!vencUser || ![u1, u2].includes(vencUser)) vencUser = u2 || u1;
  document.getElementById('venc-ano').textContent = vencAno;
  document.getElementById('venc-tabs').innerHTML = [u1, u2].filter(Boolean).map(u =>
    `<button class="tf-btn ${u === vencUser ? 'tf-active' : ''}" onclick="vencUser='${escapeHtml(u)}';vencSel=null;renderVencimentos()">${escapeHtml(u)}</button>`).join('');

  const contas = getContasFixas().filter(c => c.user === vencUser).sort((a, b) => (parseInt(a.dia) || 99) - (parseInt(b.dia) || 99) || a.nome.localeCompare(b.nome, 'pt-BR'));
  const meses = Array.from({ length: 12 }, (_, i) => `${vencAno}-${String(i + 1).padStart(2, '0')}`);
  const ymAtual = new Date().toISOString().slice(0, 7);

  renderVencProximos(contas);

  const grid = document.getElementById('venc-grid');
  if (!contas.length) {
    grid.innerHTML = `<div class="empty-state" style="padding:36px 16px;"><div class="icon">📅</div>
      <p style="font-weight:600;color:var(--text-3);">Nenhuma conta fixa de ${escapeHtml(vencUser)} ainda</p>
      <p style="font-size:13px;margin-top:4px;">Cadastre abaixo as contas que vencem todo mês (escola, luz, cartões…).</p></div>`;
  } else {
    const th = 'padding:8px 10px;font-size:12px;font-weight:700;color:var(--text-3);text-align:right;white-space:nowrap;';
    const td = 'padding:7px 10px;font-size:13px;text-align:right;white-space:nowrap;border-top:1px solid var(--border);cursor:pointer;';
    const sticky = 'position:sticky;left:0;background:var(--surface);z-index:1;';
    let totais = meses.map(() => 0);
    let html = `<table style="border-collapse:collapse;width:100%;"><thead><tr>
      <th style="${th}text-align:left;${sticky}">Conta</th><th style="${th}text-align:center;">Dia</th>
      ${meses.map(ym => `<th style="${th}${ym === ymAtual ? 'color:#4f46e5;' : ''}">${MESES[parseInt(ym.slice(5)) - 1]}</th>`).join('')}
    </tr></thead><tbody>`;
    contas.forEach(c => {
      html += `<tr><td style="${td}text-align:left;font-weight:600;color:var(--text);${sticky}" onclick="abrirContaFixa('${c.id}')" title="Editar conta">${escapeHtml(c.nome)}${c.tipo === 'cartao' ? ' <span style="font-size:11px;color:var(--muted);">💳</span>' : ''}</td>
        <td style="${td}text-align:center;cursor:default;"><span style="display:inline-block;min-width:26px;padding:2px 6px;border-radius:6px;background:var(--tint-green);color:#059669;font-weight:700;font-size:12px;">${c.dia || '—'}</span></td>`;
      meses.forEach((ym, i) => {
        const cel = vencCelula(c, ym);
        totais[i] += cel.valor;
        const sel = vencSel && vencSel.id === c.id && vencSel.ym === ym;
        const bg = sel ? 'background:var(--tint-indigo);' : (ym === ymAtual ? 'background:var(--surface-2);' : '');
        const tip = cel.status === 'parcial' ? `Pago ${brl(cel.pago)} de ${brl(cel.total)}` : ({ pago: 'Pago', aberto: 'Em aberto', atrasado: 'Vencida e não paga', vazio: '' })[cel.status];
        html += `<td style="${td}${bg}color:${VENC_COR[cel.status]};${cel.status === 'atrasado' ? 'font-weight:700;' : ''}" title="${tip}" onclick="vencSel={id:'${c.id}',ym:'${ym}'};renderVencimentos()">${cel.status === 'vazio' ? '' : brl(cel.valor).replace('R$', '').trim()}${cel.status === 'parcial' ? ' <span style="font-size:10px;">◐</span>' : ''}</td>`;
      });
      html += '</tr>';
    });
    // Entradas do mês (salário, pensão, 13º…). Venda de bem fica de fora: é pontual e distorce o saldo.
    const entradas = meses.map(ym => S.transactions.filter(t => t && t.type === 'receita' && !t.isTransfer && t.category !== 'Venda de bem' && t.user === vencUser && (t.date || '').startsWith(ym)).reduce((s, t) => s + amountBrl(t), 0));
    const linha = (rotulo, vals, cor) => `<tr><td style="${td}text-align:left;font-weight:700;${sticky}cursor:default;">${rotulo}</td><td style="${td}cursor:default;"></td>${vals.map((v, i) => `<td style="${td}cursor:default;font-weight:700;${meses[i] === ymAtual ? 'background:var(--surface-2);' : ''}color:${cor ? cor(v) : 'var(--text)'};">${v ? brl(v).replace('R$', '').trim() : ''}</td>`).join('')}</tr>`;
    html += linha('Total contas', totais);
    html += linha('Entradas', entradas, () => '#059669');
    html += linha('Saldo', entradas.map((e, i) => e - totais[i]), v => v < 0 ? '#e11d48' : '#059669');
    html += '</tbody></table>';
    grid.innerHTML = html;
  }
  renderVencDetalhe();
  renderContasFixasLista();
}

function renderVencProximos(contas) {
  const el = document.getElementById('venc-proximos');
  const hoje = new Date(); hoje.setHours(0, 0, 0, 0);
  const itens = [];
  const ym = new Date().toISOString().slice(0, 7);
  [ym, ymShift(ym, 1)].forEach(mes => contas.forEach(c => {
    const cel = vencCelula(c, mes);
    if (!cel.venc || !['aberto', 'atrasado', 'parcial'].includes(cel.status)) return;
    const dias = Math.round((new Date(cel.venc + 'T00:00:00') - hoje) / 864e5);
    if (dias > 15) return;
    itens.push({ c, cel, dias });
  }));
  itens.sort((a, b) => a.dias - b.dias);
  if (!itens.length) { el.innerHTML = ''; return; }
  el.innerHTML = `<div style="display:flex;gap:10px;flex-wrap:wrap;margin-bottom:16px;">${itens.map(({ c, cel, dias }) => {
    const cor = dias < 0 ? '#e11d48' : dias <= 3 ? '#d97706' : '#4f46e5';
    const quando = dias < 0 ? `atrasada ${-dias} dia${dias === -1 ? '' : 's'}` : dias === 0 ? 'vence hoje' : dias === 1 ? 'vence amanhã' : `vence em ${dias} dias`;
    const falta = cel.status === 'parcial' ? cel.total - cel.pago : cel.valor;
    return `<div class="card" style="padding:10px 14px;border-left:4px solid ${cor};min-width:170px;">
      <div style="font-size:13px;font-weight:700;color:var(--text);">${escapeHtml(c.nome)}</div>
      <div style="font-size:12px;color:${cor};font-weight:600;">${quando} · ${fmtDate(cel.venc).slice(0, 5)}</div>
      <div style="font-size:14px;font-weight:700;color:var(--text);margin-top:2px;">${brl(falta)}</div></div>`;
  }).join('')}</div>`;
}

function renderVencDetalhe() {
  const el = document.getElementById('venc-detalhe');
  const c = vencSel && getContasFixas().find(x => x.id === vencSel.id);
  if (!c) { el.innerHTML = ''; return; }
  const cel = vencCelula(c, vencSel.ym);
  const [y, m] = vencSel.ym.split('-').map(Number);
  const linhas = cel.txs.slice().sort((a, b) => (a.date || '').localeCompare(b.date || '')).map(t => `
    <div style="display:flex;justify-content:space-between;gap:10px;padding:6px 0;border-top:1px solid var(--border);font-size:13px;cursor:pointer;" onclick="abrirEditModal('${t.id}')" title="Abrir lançamento">
      <span style="color:var(--text-2);">${fmtDate(t.date)} · ${escapeHtml(t.desc || '')}</span>
      <span style="white-space:nowrap;font-weight:600;color:${t.pago === false ? 'var(--text)' : '#059669'};">${brl(amountBrl(t))}${t.pago === false ? (t.confirmado ? ' · valor confirmado, em aberto' : ' · em aberto') : ''}</span></div>`).join('');
  const pagos = cel.pagamentos.map(p => `<div style="font-size:12px;color:#059669;">✓ Pago ${brl(Math.abs(p.amount))} em ${fmtDate(p.date)}</div>`).join('');
  el.innerHTML = `<div class="card" style="margin-top:16px;padding:16px 20px;">
    <div style="display:flex;justify-content:space-between;align-items:center;gap:10px;margin-bottom:6px;">
      <div><b style="color:var(--text);">${escapeHtml(c.nome)} — ${MESES_FULL[m - 1]} ${y}</b>
      <span style="font-size:12px;color:var(--muted);margin-left:6px;">${cel.venc ? 'vence ' + fmtDate(cel.venc) : ''}</span></div>
      <button class="ico-btn" onclick="vencSel=null;renderVencimentos()" aria-label="Fechar" style="background:none;border:none;cursor:pointer;color:var(--muted);"><svg class="ico-sm"><use href="#i-x"/></svg></button>
    </div>
    ${c.tipo === 'cartao' ? `<div style="font-size:13px;color:var(--text-2);margin-bottom:4px;">Fatura: <b>${brl(cel.total)}</b> em ${cel.txs.length} lançamento(s)</div>${pagos}` : ''}
    ${cel.txs.length ? (c.tipo === 'cartao' ? `<details style="margin-top:6px;"><summary style="font-size:12px;color:var(--muted);cursor:pointer;">Ver compras da fatura</summary>${linhas}</details>` : linhas)
      : `<p style="font-size:13px;color:var(--muted);">Nenhum lançamento encontrado pra essa conta nesse mês. Se foi pago, lance em Nova Transação (no débito) com uma descrição que a regra reconheça.</p>`}
  </div>`;
}

// ─── CADASTRO DAS CONTAS FIXAS ────────────────────────────────────────────────
function renderContasFixasLista() {
  const el = document.getElementById('venc-gerenciar');
  const contas = getContasFixas().filter(c => c.user === vencUser);
  const regra = c => c.tipo === 'cartao'
    ? `Fatura do ${escapeHtml((S.accounts.find(a => a.id === c.cartaoId) || {}).label || 'cartão')}`
    : [c.texto ? `descrição tem "${escapeHtml(c.texto.split('|').join('" ou "'))}"` : '', c.cat ? `categoria ${escapeHtml(c.cat)}${c.sub ? ' › ' + escapeHtml(c.sub) : ''}` : ''].filter(Boolean).join(' e ') || '—';
  el.innerHTML = `<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;gap:10px;flex-wrap:wrap;">
      <p style="font-weight:700;font-size:14px;color:var(--text-2);">Contas fixas de ${escapeHtml(vencUser)}</p>
      <button onclick="abrirContaFixa(null)" style="padding:8px 14px;background:#4f46e5;color:white;border:none;border-radius:10px;font-size:13px;font-weight:700;cursor:pointer;">+ Nova conta</button>
    </div>
    ${contas.length ? contas.map(c => `<div style="display:flex;justify-content:space-between;align-items:center;gap:10px;padding:7px 0;border-top:1px solid var(--border);font-size:13px;">
      <span><b style="color:var(--text);">${escapeHtml(c.nome)}</b> <span style="color:var(--muted);">· dia ${c.dia || '—'} · ${regra(c)}</span></span>
      <button class="ico-btn" onclick="abrirContaFixa('${c.id}')" aria-label="Editar" style="background:none;border:none;cursor:pointer;color:var(--text-3);"><svg class="ico-sm"><use href="#i-edit"/></svg></button></div>`).join('')
      : '<p style="font-size:13px;color:var(--muted);">Nenhuma ainda.</p>'}
    <div id="venc-form" style="display:none;margin-top:14px;padding-top:14px;border-top:2px solid var(--border);"></div>`;
  if (vencFormAberto) desenharFormContaFixa();
}

function abrirContaFixa(id) {
  vencEditId = id;
  vencFormAberto = true;
  desenharFormContaFixa();
  document.getElementById('venc-form')?.scrollIntoView({ behavior: 'smooth', block: 'center' });
}

function desenharFormContaFixa() {
  const el = document.getElementById('venc-form');
  if (!el) return;
  const c = getContasFixas().find(x => x.id === vencEditId) || { nome: '', dia: '', tipo: 'lanc', texto: '', cat: '', sub: '', cartaoId: '' };
  const cartoes = S.accounts.filter(a => a.accountType === 'cartao' && a.owner === vencUser);
  const cats = getDespesaCats();
  el.style.display = 'block';
  el.innerHTML = `<p style="font-weight:700;font-size:13px;color:var(--text-2);margin-bottom:10px;">${vencEditId ? 'Editar conta' : 'Nova conta'}</p>
    <div style="display:grid;grid-template-columns:2fr 1fr 1.4fr;gap:10px;margin-bottom:10px;">
      <input class="finput" id="vf-nome" placeholder="Nome (ex.: Escola Pedro)" value="${escapeHtml(c.nome)}">
      <input class="finput" id="vf-dia" type="number" min="1" max="31" placeholder="Dia do vencimento" value="${c.dia || ''}">
      <select class="finput" id="vf-tipo" onchange="document.getElementById('vf-cartao-row').style.display=this.value==='cartao'?'grid':'none';document.getElementById('vf-lanc-row').style.display=this.value==='cartao'?'none':'grid';">
        <option value="lanc" ${c.tipo !== 'cartao' ? 'selected' : ''}>Conta (lançamentos no débito)</option>
        <option value="cartao" ${c.tipo === 'cartao' ? 'selected' : ''}>Fatura de cartão</option>
      </select>
    </div>
    <div id="vf-cartao-row" style="display:${c.tipo === 'cartao' ? 'grid' : 'none'};grid-template-columns:1fr;gap:10px;margin-bottom:10px;">
      <select class="finput" id="vf-cartao">${cartoes.map(a => `<option value="${a.id}" ${a.id === c.cartaoId ? 'selected' : ''}>${escapeHtml(a.label)}</option>`).join('') || '<option value="">— nenhum cartão —</option>'}</select>
    </div>
    <div id="vf-lanc-row" style="display:${c.tipo === 'cartao' ? 'none' : 'grid'};grid-template-columns:2fr 1fr 1fr;gap:10px;margin-bottom:10px;">
      <input class="finput" id="vf-texto" placeholder="Palavras da descrição (separe com |)" value="${escapeHtml(c.texto || '')}">
      <select class="finput" id="vf-cat" onchange="vencOnCatChange()">
        <option value="">Qualquer categoria</option>${Object.keys(cats).map(k => `<option ${k === c.cat ? 'selected' : ''}>${escapeHtml(k)}</option>`).join('')}</select>
      <select class="finput" id="vf-sub"><option value="">Qualquer sub</option>${(cats[c.cat] || []).map(s => `<option ${s === c.sub ? 'selected' : ''}>${escapeHtml(s)}</option>`).join('')}</select>
    </div>
    <p style="font-size:11px;color:var(--muted);margin-bottom:10px;">A grade soma os lançamentos no débito de ${escapeHtml(vencUser)} que batem com a regra. Clique numa célula pra ver quais entraram.</p>
    <div style="display:flex;gap:8px;justify-content:flex-end;">
      ${vencEditId ? `<button onclick="excluirContaFixa()" style="padding:8px 14px;background:none;color:#e11d48;border:1px solid #e11d48;border-radius:10px;font-size:13px;font-weight:600;cursor:pointer;margin-right:auto;">Excluir</button>` : ''}
      <button onclick="vencFormAberto=false;document.getElementById('venc-form').style.display='none'" style="padding:8px 14px;background:var(--surface);color:var(--text-2);border:1px solid var(--border);border-radius:10px;font-size:13px;font-weight:600;cursor:pointer;">Cancelar</button>
      <button onclick="salvarContaFixa()" style="padding:8px 16px;background:#4f46e5;color:white;border:none;border-radius:10px;font-size:13px;font-weight:700;cursor:pointer;">Salvar</button>
    </div>`;
}

function vencOnCatChange() {
  const subs = getDespesaCats()[document.getElementById('vf-cat').value] || [];
  document.getElementById('vf-sub').innerHTML = '<option value="">Qualquer sub</option>' + subs.map(x => `<option>${escapeHtml(x)}</option>`).join('');
}

function salvarContaFixa() {
  const nome = document.getElementById('vf-nome').value.trim();
  if (!nome) return toast('❌ Informe o nome da conta');
  const tipo = document.getElementById('vf-tipo').value;
  const dados = {
    nome,
    dia: parseInt(document.getElementById('vf-dia').value) || null,
    user: vencUser,
    tipo,
    cartaoId: tipo === 'cartao' ? document.getElementById('vf-cartao').value : '',
    texto: tipo === 'cartao' ? '' : document.getElementById('vf-texto').value.trim(),
    cat: tipo === 'cartao' ? '' : document.getElementById('vf-cat').value,
    sub: tipo === 'cartao' ? '' : document.getElementById('vf-sub').value,
  };
  if (tipo === 'cartao' && !dados.cartaoId) return toast('❌ Escolha o cartão');
  if (tipo !== 'cartao' && !dados.texto && !dados.cat) return toast('❌ Informe palavras da descrição ou uma categoria');
  const lista = getContasFixas();
  const atual = lista.find(x => x.id === vencEditId);
  if (atual) Object.assign(atual, dados, { updatedAt: new Date().toISOString() });
  else lista.push({ id: 'cf_' + Date.now(), ...dados, updatedAt: new Date().toISOString() });
  vencFormAberto = false;
  save(); toast('✅ Conta salva'); renderVencimentos();
}

function excluirContaFixa() {
  const c = getContasFixas().find(x => x.id === vencEditId);
  if (!c || !confirm(`Tirar "${c.nome}" da grade de vencimentos? (Os lançamentos não são apagados.)`)) return;
  S.contasFixas = getContasFixas().filter(x => x.id !== vencEditId);
  vencFormAberto = false; vencSel = null;
  save(); toast('🗑️ Conta removida da grade'); renderVencimentos();
}
