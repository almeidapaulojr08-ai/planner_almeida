// FinançasCasal — 03-contas.js
// Módulo carregado em ordem por index.html (escopo global compartilhado). Não use import/export.

// ─── CONTAS ────────────────────────────────────────────────────────────────────
function bankIcon(bankKey, size=40) {
  const b = BANKS[bankKey] || { color:'var(--text-3)', svg:'<text x="24" y="31" text-anchor="middle" fill="white" font-size="16">?</text>' };
  const r = Math.round(size * 0.28);
  return `<div style="width:${size}px;height:${size}px;background:${b.color};border-radius:${r}px;flex-shrink:0;overflow:hidden;">
    <svg width="${size}" height="${size}" viewBox="0 0 48 48" xmlns="http://www.w3.org/2000/svg">${b.svg}</svg>
  </div>`;
}

let contaTabAtual = 'conta';
let tipoContaAtual = 'corrente';

function setContaTab(tab) {
  contaTabAtual = tab;
  const isConta = tab === 'conta';
  document.getElementById('cm-tab-conta').style.cssText  = `padding:10px;border-radius:10px;border:2px solid ${isConta?'#4f46e5':'var(--border)'};background:${isConta?'var(--tint-indigo)':'var(--surface)'};color:${isConta?'#4f46e5':'var(--text-3)'};font-weight:700;font-size:14px;cursor:pointer;`;
  document.getElementById('cm-tab-cartao').style.cssText = `padding:10px;border-radius:10px;border:2px solid ${!isConta?'#4f46e5':'var(--border)'};background:${!isConta?'var(--tint-indigo)':'var(--surface)'};color:${!isConta?'#4f46e5':'var(--text-3)'};font-weight:700;font-size:14px;cursor:pointer;`;
  document.getElementById('cm-fields-conta').style.display  = isConta  ? 'block' : 'none';
  document.getElementById('cm-fields-cartao').style.display = !isConta ? 'block' : 'none';
}

function setTipoConta(tipo) {
  tipoContaAtual = tipo;
  const isC = tipo === 'corrente';
  document.getElementById('cm-tipo-corrente').style.cssText = `padding:10px;border-radius:10px;border:2px solid ${isC?'#4f46e5':'var(--border)'};background:${isC?'var(--tint-indigo)':'var(--surface)'};color:${isC?'#4f46e5':'var(--text-3)'};font-weight:600;font-size:13px;cursor:pointer;`;
  document.getElementById('cm-tipo-poupanca').style.cssText = `padding:10px;border-radius:10px;border:2px solid ${!isC?'#4f46e5':'var(--border)'};background:${!isC?'var(--tint-indigo)':'var(--surface)'};color:${!isC?'#4f46e5':'var(--text-3)'};font-weight:600;font-size:13px;cursor:pointer;`;
}

function renderContaCard(a) {
  const b = BANKS[a.bank];
  const isUSD = b.currency === 'USD';
  const isCartao = a.accountType === 'cartao';
  let sub = `Titular: <strong style="color:var(--text-2);">${a.owner}</strong>`;
  if (!isCartao) {
    const tipoLabel = a.tipoConta === 'poupanca' ? 'Poupança' : 'Conta Corrente';
    sub += ` &nbsp;·&nbsp; ${tipoLabel}`;
    if (isUSD && S.settings.usdRate) sub += ` &nbsp;·&nbsp; Cotação: R$ ${S.settings.usdRate.toFixed(2)}`;
  } else {
    if (a.limite) sub += ` &nbsp;·&nbsp; Limite: R$ ${Number(a.limite).toLocaleString('pt-BR')}`;
    if (a.fecha)  sub += ` &nbsp;·&nbsp; Fecha dia ${a.fecha} &nbsp;·&nbsp; Vence dia ${a.vence}`;
  }
  return `<div class="card" style="display:flex;align-items:center;gap:16px;">
    ${bankIcon(a.bank, 48)}
    <div style="flex:1;">
      <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;">
        <p style="font-weight:700;font-size:15px;color:var(--text);">${escapeHtml(b.name)}</p>
        <span style="font-size:12px;color:var(--muted);">•</span>
        <p style="font-size:13px;color:var(--text-3);">${escapeHtml(a.label)}</p>
        ${isUSD ? `<span class="badge" style="background:#FFFBEA;color:var(--on-amber);">USD</span>` : ''}
      </div>
      <p style="font-size:12px;color:var(--muted);margin-top:3px;">${escapeHtml(sub)}</p>
    </div>
    <div style="display:flex;gap:4px;">
      <button onclick="editarConta('${a.id}')" style="background:none;border:none;cursor:pointer;padding:8px;border-radius:8px;font-size:16px;" title="Editar" class="ico-btn" aria-label="Editar"><svg class="ico-sm"><use href="#i-edit"/></svg></button>
      <button onclick="deletarConta('${a.id}')" style="background:none;border:none;cursor:pointer;padding:8px;border-radius:8px;font-size:16px;" title="Excluir" class="ico-btn" aria-label="Excluir"><svg class="ico-sm"><use href="#i-trash"/></svg></button>
    </div>
  </div>`;
}

function renderContas() {
  const contas  = S.accounts.filter(a => a.accountType !== 'cartao');
  const cartoes = S.accounts.filter(a => a.accountType === 'cartao');

  const elC = document.getElementById('contas-list');
  elC.innerHTML = contas.length
    ? contas.map(renderContaCard).join('')
    : `<div class="empty-state" style="padding:28px 16px;"><div class="icon">🏦</div><p style="font-weight:600;color:var(--text-3);">Nenhuma conta ainda</p><p style="font-size:13px;margin-top:4px;">Cadastre a conta corrente ou carteira de cada um pra acompanhar os saldos.</p><button onclick="abrirModalConta()" style="margin-top:12px;padding:10px 18px;background:#4f46e5;color:white;border:none;border-radius:10px;font-size:13px;font-weight:600;cursor:pointer;">Adicionar conta</button></div>`;

  const elK = document.getElementById('cartoes-list');
  elK.innerHTML = cartoes.length
    ? cartoes.map(renderContaCard).join('')
    : `<div class="empty-state" style="padding:28px 16px;"><div class="icon">💳</div><p style="font-weight:600;color:var(--text-3);">Nenhum cartão ainda</p><p style="font-size:13px;margin-top:4px;">Com o cartão cadastrado, as faturas fecham e vencem nas datas certas.</p><button onclick="abrirModalConta()" style="margin-top:12px;padding:10px 18px;background:#4f46e5;color:white;border:none;border-radius:10px;font-size:13px;font-weight:600;cursor:pointer;">Adicionar cartão</button></div>`;
}

// Carrega bancos customizados salvos no state
function getAllBanks() {
  const custom = S.customBanks || {};
  return { ...BANKS, ...custom };
}

function refreshBancoSelect(selectedBank) {
  const sel = document.getElementById('cm-banco');
  const banks = getAllBanks();
  sel.innerHTML = Object.entries(banks).map(([key, b]) =>
    `<option value="${key}">${escapeHtml(b.name)}${b.currency === 'USD' ? ' (USD)' : ''}</option>`
  ).join('') + '<option value="__novo__">+ Novo Banco...</option>';
  if (selectedBank && banks[selectedBank]) sel.value = selectedBank;
  onCmBancoChange();
}

function onCmBancoChange() {
  const isNovo = document.getElementById('cm-banco').value === '__novo__';
  document.getElementById('cm-novo-banco-fields').style.display = isNovo ? 'block' : 'none';
}

function criarBancoCustom() {
  const nome = document.getElementById('cm-novo-banco-nome').value.trim();
  const sigla = (document.getElementById('cm-novo-banco-sigla').value.trim() || nome.substring(0, 2)).toUpperCase();
  const cor = document.getElementById('cm-novo-banco-cor').value;
  const moeda = document.getElementById('cm-novo-banco-moeda').value;
  if (!nome) { toast('Informe o nome do banco'); return null; }
  const key = nome.toLowerCase().replace(/[^a-z0-9]/g, '_');
  const lighten = (hex) => {
    const r = parseInt(hex.slice(1,3),16), g = parseInt(hex.slice(3,5),16), b = parseInt(hex.slice(5,7),16);
    return `rgb(${Math.min(255,r+200)},${Math.min(255,g+200)},${Math.min(255,b+200)})`;
  };
  const fontSize = sigla.length > 2 ? 14 : 18;
  const bank = {
    name: nome, color: cor, bg: lighten(cor), currency: moeda,
    svg: `<rect width="48" height="48" rx="0" fill="${cor}"/><text x="24" y="32" text-anchor="middle" fill="white" font-family="Arial,sans-serif" font-weight="900" font-size="${fontSize}">${sigla}</text>`
  };
  if (!S.customBanks) S.customBanks = {};
  S.customBanks[key] = bank;
  BANKS[key] = bank;
  save();
  return key;
}

function abrirModalConta() {
  editContaId = null;
  contaTabAtual = 'conta';
  tipoContaAtual = 'corrente';
  document.getElementById('conta-modal-title').textContent = 'Adicionar';
  refreshBancoSelect('nubank');
  document.getElementById('cm-label').value  = '';
  document.getElementById('cm-limite').value = '';
  document.getElementById('cm-fecha').value  = '';
  document.getElementById('cm-vence').value  = '';
  const { u1, u2 } = S.settings;
  document.getElementById('cm-owner').innerHTML = `<option value="${u1}">${u1}</option><option value="${u2}">${u2}</option>`;
  setContaTab('conta');
  setTipoConta('corrente');
  document.getElementById('conta-modal').style.display = 'block';
}

function editarConta(id) {
  const a = S.accounts.find(x => x.id === id);
  if (!a) return;
  editContaId = id;
  contaTabAtual = a.accountType === 'cartao' ? 'cartao' : 'conta';
  tipoContaAtual = a.tipoConta || 'corrente';
  document.getElementById('conta-modal-title').textContent = 'Editar';
  refreshBancoSelect(a.bank);
  document.getElementById('cm-label').value  = a.label;
  document.getElementById('cm-limite').value = a.limite || '';
  document.getElementById('cm-fecha').value  = a.fecha  || '';
  document.getElementById('cm-vence').value  = a.vence  || '';
  const { u1, u2 } = S.settings;
  document.getElementById('cm-owner').innerHTML = `<option value="${u1}"${a.owner===u1?' selected':''}>${u1}</option><option value="${u2}"${a.owner===u2?' selected':''}>${u2}</option>`;
  setContaTab(contaTabAtual);
  setTipoConta(tipoContaAtual);
  document.getElementById('conta-modal').style.display = 'block';
}

function fecharModalConta() {
  document.getElementById('conta-modal').style.display = 'none';
  editContaId = null;
}

function salvarConta() {
  let bank = document.getElementById('cm-banco').value;
  if (bank === '__novo__') {
    bank = criarBancoCustom();
    if (!bank) return;
  }
  const label = document.getElementById('cm-label').value.trim();
  const owner = document.getElementById('cm-owner').value;
  if (!label) { toast('❌ Informe um apelido'); return; }

  const isCartao = contaTabAtual === 'cartao';
  const obj = {
    id: editContaId || Date.now().toString(),
    bank, label, owner,
    accountType: isCartao ? 'cartao' : 'conta',
    tipoConta:   isCartao ? null : tipoContaAtual,
    limite: isCartao ? document.getElementById('cm-limite').value : null,
    fecha:  isCartao ? document.getElementById('cm-fecha').value  : null,
    vence:  isCartao ? document.getElementById('cm-vence').value  : null,
    updatedAt: new Date().toISOString(),
  };

  let fechaAntigo = null;
  if (editContaId) {
    const idx = S.accounts.findIndex(x => x.id === editContaId);
    if (idx >= 0) {
      fechaAntigo = S.accounts[idx].fecha || null;
      S.accounts[idx] = obj;
    }
  } else {
    S.accounts.push(obj);
  }

  // If card fecha changed, recalculate faturaRef for all credit transactions on this card
  if (isCartao && obj.fecha && obj.fecha !== fechaAntigo) {
    const fechaDia = parseInt(obj.fecha);
    let recalc = 0;
    S.transactions.forEach(tx => {
      if (tx.formaPgto === 'credito' && tx.accountId === obj.id) {
        const [ty, tm, day] = tx.date.split('-').map(Number);
        // aritmética de mês inteiro (setMonth estoura em datas 29/30/31)
        const ymRef = ty * 12 + (tm - 1) + (day > fechaDia ? 1 : 0);
        const newRef = `${Math.floor(ymRef / 12)}-${String(ymRef % 12 + 1).padStart(2, '0')}`;
        if (tx.faturaRef !== newRef) { tx.faturaRef = newRef; recalc++; }
      }
    });
    if (recalc > 0) toast(`🔄 ${recalc} transações recalculadas para nova data de fechamento`);
  }

  save();
  fecharModalConta();
  renderContas();
  toast('✅ Salvo!');
}

function deletarConta(id) {
  if (!confirm('Excluir? As transações vinculadas não serão apagadas.')) return;
  S.deletedIds.push({ id, collection: 'accounts', deletedAt: new Date().toISOString() });
  S.accounts = S.accounts.filter(a => a.id !== id);
  save();
  renderContas();
  toast('🗑️ Removido');
}

// ─── TITULAR FILTER ────────────────────────────────────────────────────────────
function setTitularFilter(f) {
  titularFilter = f;
  ['tf-ambos','tf-paulo','tf-esposa'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.classList.toggle('tf-active', id === 'tf-' + f);
  });
  renderDashboard();
}

function txByTitular(txs) {
  if (titularFilter === 'ambos') return txs;
  const name = titularFilter === 'paulo' ? S.settings.u1 : S.settings.u2;
  return txs.filter(t => t.user === name);
}
