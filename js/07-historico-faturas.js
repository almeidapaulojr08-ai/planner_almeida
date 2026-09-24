// FinançasCasal — 07-historico-faturas.js
// Módulo carregado em ordem por index.html (escopo global compartilhado). Não use import/export.

// ─── HISTÓRICO ────────────────────────────────────────────────────────────────
let histMesSel = null; // selected month (0-11) or null for all
let histAno = new Date().getFullYear();
let histReady = false;
let selectedTxIds = new Set();
let pendingImportTxs = [];

function initHistFiltros() {
  const now = new Date();
  if (!histReady) {
    histAno = now.getFullYear();
    histMesSel = now.getMonth();
    histReady = true;
  }
  // Year selector — ignora datas malformadas (fora da faixa 2000-2100)
  const anoSel = document.getElementById('fil-ano');
  const years = new Set();
  S.transactions.forEach(t => {
    const y = parseInt((t.date||'').substring(0,4));
    if (y >= 2000 && y <= 2100) years.add(y);
  });
  years.add(now.getFullYear());
  const sorted = [...years].sort((a,b) => b - a);
  anoSel.innerHTML = sorted.map(y => `<option value="${y}">${y}</option>`).join('');
  anoSel.value = histAno;

  // Month pills
  const pillsEl = document.getElementById('fil-meses-pills');
  pillsEl.innerHTML = MESES.map((m, i) =>
    `<button onclick="setHistMes(${i})" class="mes-pill ${histMesSel === i ? 'mp-active' : ''}">${m}</button>`
  ).join('');

  // Titular filter (preserve current selection)
  const titSel = document.getElementById('fil-titular');
  const prevTit = titSel.value;
  titSel.innerHTML = `<option value="">Todos Titulares</option>
    <option value="${S.settings.u1}">${S.settings.u1}</option>
    <option value="${S.settings.u2}">${S.settings.u2}</option>`;
  if (prevTit) titSel.value = prevTit;

  // Conta filter (filtrada pelo titular selecionado)
  refreshHistContaFilter();

  // Categoria filter — filtrado pelo tipo selecionado (receita/despesa/etc)
  refreshHistCatFilter();
}

function refreshHistContaFilter() {
  const titular = document.getElementById('fil-titular')?.value || '';
  const contaSel = document.getElementById('fil-conta');
  const prevConta = contaSel.value;
  const filtered = titular ? S.accounts.filter(a => a.owner === titular) : S.accounts;
  contaSel.innerHTML = '<option value="">Todas Contas</option>' +
    filtered.map(a => {
      const b = BANKS[a.bank];
      const icon = a.accountType === 'cartao' ? '💳' : '🏦';
      return `<option value="${a.id}">${icon} ${b ? b.name : a.bank} — ${escapeHtml(a.label)} [${escapeHtml(a.owner)}]</option>`;
    }).join('');
  if (prevConta && filtered.some(a => a.id === prevConta)) contaSel.value = prevConta;
}

function refreshHistSubcatFilter() {
  const catFil = document.getElementById('fil-categoria')?.value || '';
  const subSel = document.getElementById('fil-subcategoria');
  const prevSub = subSel.value;
  if (!catFil) {
    subSel.innerHTML = '<option value="">Todas Subcategorias</option>';
    subSel.disabled = true;
    subSel.style.opacity = '0.5';
    return;
  }
  subSel.disabled = false;
  subSel.style.opacity = '1';
  const allSubs = new Set();
  S.transactions.forEach(t => {
    if (t.category === catFil) allSubs.add(t.subcategory || 'Geral');
  });
  const sortedSubs = [...allSubs].sort((a, b) => a.localeCompare(b, 'pt-BR'));
  subSel.innerHTML = '<option value="">Todas Subcategorias</option>' +
    sortedSubs.map(s => `<option value="${s}">${s}</option>`).join('');
  if (prevSub && allSubs.has(prevSub)) subSel.value = prevSub;
}

function refreshHistCatFilter() {
  const tipo = document.getElementById('fil-tipo')?.value || '';
  const catSel = document.getElementById('fil-categoria');
  const prevCat = catSel.value;
  const allCats = new Set();
  S.transactions.forEach(t => {
    if (!t.category) return;
    if (tipo === 'transferencia') { if (t.isTransfer) allCats.add(t.category); }
    else if (tipo) { if (t.type === tipo && !t.isTransfer) allCats.add(t.category); }
    else allCats.add(t.category);
  });
  const sortedCats = [...allCats].sort((a, b) => a.localeCompare(b, 'pt-BR'));
  catSel.innerHTML = '<option value="">Todas Categorias</option>' +
    sortedCats.map(c => `<option value="${c}">${c}</option>`).join('');
  if (prevCat && allCats.has(prevCat)) catSel.value = prevCat;
  else catSel.value = '';
  refreshHistSubcatFilter();
}

function onFilCategoriaChange() {
  document.getElementById('fil-subcategoria').value = '';
  refreshHistSubcatFilter();
}

function onFilTipoChange() {
  document.getElementById('fil-categoria').value = '';
  document.getElementById('fil-subcategoria').value = '';
  refreshHistCatFilter();
  renderHistorico();
}

function onHistTitularChange() {
  refreshHistContaFilter();
  renderHistorico();
}

function onHistAnoChange() {
  histAno = parseInt(document.getElementById('fil-ano').value);
  renderHistorico();
}

function setHistMes(m) {
  histMesSel = (histMesSel === m) ? null : m;
  renderHistorico();
}

function setHistPeriodo() {
  histMesSel = null;
  renderHistorico();
}

function getFilteredTxs() {
  const search   = (document.getElementById('fil-search')?.value || '').toLowerCase();
  const tipo     = document.getElementById('fil-tipo')?.value || '';
  const pagamento = document.getElementById('fil-pagamento')?.value || '';
  const conta    = document.getElementById('fil-conta')?.value || '';
  const titular  = document.getElementById('fil-titular')?.value || '';
  const dataDe   = document.getElementById('fil-data-de')?.value || '';
  const dataAte  = document.getElementById('fil-data-ate')?.value || '';

  let txs = S.transactions.filter(t => t.type !== 'investimento');

  // Date range filter (takes priority over year/month pills)
  if (dataDe || dataAte) {
    if (dataDe)  txs = txs.filter(t => t.date >= dataDe);
    if (dataAte) txs = txs.filter(t => t.date <= dataAte);
  } else {
    // Year filter
    txs = txs.filter(t => t.date.startsWith(String(histAno)));
    // Month filter
    if (histMesSel !== null) {
      const ym = `${histAno}-${String(histMesSel+1).padStart(2,'0')}`;
      txs = txs.filter(t => t.date.startsWith(ym));
    }
  }

  const custo    = document.getElementById('fil-custo')?.value || '';

  if (search)    txs = txs.filter(t => t.desc.toLowerCase().includes(search) || t.category.toLowerCase().includes(search) || (t.subcategory||'').toLowerCase().includes(search));
  if (tipo === 'transferencia') {
    txs = txs.filter(t => t.isTransfer);
  } else if (tipo) {
    txs = txs.filter(t => t.type === tipo && !t.isTransfer);
  }
  if (pagamento) txs = txs.filter(t => t.formaPgto === pagamento);
  if (conta)     txs = txs.filter(t => t.accountId === conta);
  const catFil   = document.getElementById('fil-categoria')?.value || '';
  const subFil   = document.getElementById('fil-subcategoria')?.value || '';
  if (catFil)    txs = txs.filter(t => t.category === catFil);
  if (subFil)    txs = txs.filter(t => (t.subcategory || 'Geral') === subFil);
  if (custo)     txs = txs.filter(t => (t.custoTipo || autoCustoTipo(t.category)) === custo);
  const status   = document.getElementById('fil-status')?.value || '';
  if (status === 'pendente') txs = txs.filter(t => t.pago === false);
  if (status === 'pago')     txs = txs.filter(t => t.pago !== false);
  const recFil   = document.getElementById('fil-recorrente')?.value || '';
  if (recFil === 'sim') txs = txs.filter(t => t.recorrente === true);
  if (recFil === 'nao') txs = txs.filter(t => !t.recorrente);
  if (titular)   txs = txs.filter(t => t.user === titular);

  txs.sort((a,b) => new Date(b.date) - new Date(a.date));
  return txs;
}

function limparFiltroData() {
  document.getElementById('fil-data-de').value = '';
  document.getElementById('fil-data-ate').value = '';
  renderHistorico();
}

function renderHistorico() {
  initHistFiltros();
  selectedTxIds.clear();
  const chk = document.getElementById('hist-select-all');
  if (chk) chk.checked = false;
  document.getElementById('hist-bulk-delete').style.display = 'none';
  document.getElementById('hist-bulk-pago').style.display = 'none';

  // Update pills
  const pillsEl = document.getElementById('fil-meses-pills');
  pillsEl.innerHTML = MESES.map((m, i) =>
    `<button onclick="setHistMes(${i})" class="mes-pill ${histMesSel === i ? 'mp-active' : ''}">${m}</button>`
  ).join('');
  document.getElementById('fil-periodo-all').className = 'tf-btn' + (histMesSel === null ? ' tf-active' : '');

  const txs = getFilteredTxs();
  const empty = document.getElementById('hist-empty');
  const cards = document.getElementById('hist-cards');

  // Summary cards (exclui transferências — não são despesa/receita real)
  const rec  = txs.filter(t=>t.type==='receita' && !t.isTransfer).reduce((s,t)=>s+amountBrl(t),0);
  const desp = txs.filter(t=>t.type==='despesa' && !t.isTransfer && !isPgtoFatura(t)).reduce((s,t)=>s+amountBrl(t),0);
  const pend = txs.filter(t=>t.type==='despesa' && t.pago===false && !t.isTransfer && !isPgtoFatura(t)).reduce((s,t)=>s+amountBrl(t),0);
  const inv  = txs.filter(t=>t.type==='investimento').reduce((s,t)=>s+amountBrl(t),0);
  document.getElementById('hist-sum-rec').textContent  = brl(rec);
  document.getElementById('hist-sum-desp').textContent = brl(desp);
  document.getElementById('hist-sum-pend').textContent = brl(pend);
  document.getElementById('hist-sum-inv').textContent  = brl(inv);

  document.getElementById('hist-count').textContent = `${txs.length} transações`;

  if (!txs.length) {
    cards.innerHTML = '';
    empty.style.display = 'block';
    return;
  }
  empty.style.display = 'none';

  const TYPE_CFG = {
    receita:      { ic:'<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#059669" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="19" x2="12" y2="5"/><polyline points="5 12 12 5 19 12"/></svg>', clr:'#059669', sign:'+', badgeBg:'var(--tint-green)', badgeTxt:'var(--on-green)', label:'Receita' },
    despesa:      { ic:'<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#e11d48" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="5" x2="12" y2="19"/><polyline points="19 12 12 19 5 12"/></svg>', clr:'#e11d48', sign:'-', badgeBg:'var(--tint-rose)', badgeTxt:'var(--on-rose)', label:'Despesa' },
    investimento: { ic:'🐷', clr:'#2563eb', sign:'+', badgeBg:'var(--tint-blue)', badgeTxt:'var(--on-blue)', label:'Investimento' }
  };
  const TRANSFER_CFG = { ic:'🔄', clr:'#6366f1', sign:'', badgeBg:'var(--tint-indigo)', badgeTxt:'#4f46e5', label:'Transferência' };

  cards.innerHTML = txs.map(t => {
    const s = t.isTransfer ? TRANSFER_CFG : TYPE_CFG[t.type];
    const acc = S.accounts.find(a => a.id === t.accountId);
    const bankName = acc ? (BANKS[acc.bank]?.name || acc.bank) : '';
    const bankIcon = acc ? (acc.accountType === 'cartao' ? '💳' : '🏦') : '';
    const catLabel = t.category + (t.subcategory ? ' › ' + t.subcategory : '');
    const pagoTag = t.type === 'despesa'
      ? `<span style="font-size:11px;font-weight:600;padding:2px 8px;border-radius:20px;background:${t.pago===false?'var(--tint-amber)':'var(--tint-green)'};color:${t.pago===false?'var(--on-amber)':'var(--on-green)'};margin-left:6px;">${t.pago===false?'Pendente':'Paga'}</span>`
      : '';

    return `<div class="tx-card">
      <input type="checkbox" data-txid="${t.id}" onchange="onTxSelect(this)" style="width:15px;height:15px;accent-color:#4f46e5;flex-shrink:0;">
      <div style="flex:1;min-width:0;">
        <div style="display:flex;align-items:center;gap:8px;margin-bottom:4px;">
          <span style="font-size:14px;font-weight:700;color:var(--text);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">${escapeHtml(t.desc)}</span>
          ${pagoTag}
        </div>
        <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;">
          <span style="font-size:11px;color:var(--muted);">${bankIcon} ${bankName}</span>
          ${acc ? `<span style="font-size:11px;color:var(--muted);">· ${escapeHtml(acc.label)}</span>` : ''}
          <span style="font-size:11px;font-weight:600;padding:2px 8px;border-radius:20px;background:${s.badgeBg};color:${s.badgeTxt};">${s.label}</span>
          <span style="font-size:11px;font-weight:600;padding:2px 8px;border-radius:20px;background:var(--surface-2);color:var(--text-2);">${catLabel}</span>
          ${t.formaPgto ? `<span style="font-size:11px;padding:2px 8px;border-radius:20px;background:var(--bg);color:var(--muted);">${t.formaPgto==='credito'?'Crédito':'Débito'}</span>` : ''}
          ${t.parcela ? `<span style="font-size:11px;font-weight:600;padding:2px 8px;border-radius:20px;background:var(--tint-amber);color:var(--on-amber);">🔄 ${t.parcela}</span>` : ''}
          ${t.custoTipo ? `<span style="font-size:11px;padding:2px 8px;border-radius:20px;background:${t.custoTipo==='fixo'?'var(--tint-indigo)':'#fff7ed'};color:${t.custoTipo==='fixo'?'#4f46e5':'#c2410c'};">${t.custoTipo==='fixo'?'📌 Fixo':'📊 Variável'}</span>` : ''}
          ${t.recorrente ? `<span style="font-size:11px;font-weight:600;padding:2px 8px;border-radius:20px;background:var(--tint-green);color:#15803d;">🔁 Recorrente</span>` : ''}
        </div>
      </div>
      <div style="text-align:right;flex-shrink:0;">
        <p style="font-size:16px;font-weight:800;color:${t.isNegative ? '#059669' : s.clr};">${t.isNegative ? '+' : s.sign}${brl(amountBrl(t) < 0 ? -amountBrl(t) : amountBrl(t))}</p>
        <p style="font-size:11px;color:var(--muted);margin-top:2px;">${fmtDate(t.date)}</p>
      </div>
      <div style="display:flex;gap:2px;flex-shrink:0;">
        <button onclick="abrirEditModal('${t.id}')" style="background:none;border:none;cursor:pointer;padding:6px;border-radius:8px;transition:background 0.1s;" onmouseover="this.style.background='var(--tint-indigo)'" onmouseout="this.style.background='none'" title="Editar" class="ico-btn" aria-label="Editar"><svg class="ico-sm"><use href="#i-edit"/></svg></button>
        <button onclick="del('${t.id}')" style="background:none;border:none;cursor:pointer;padding:6px;border-radius:8px;transition:background 0.1s;" onmouseover="this.style.background='var(--tint-rose)'" onmouseout="this.style.background='none'" title="Excluir" class="ico-btn" aria-label="Excluir"><svg class="ico-sm"><use href="#i-trash"/></svg></button>
      </div>
    </div>`;
  }).join('');
}

function onTxSelect(chk) {
  const id = chk.dataset.txid;
  if (chk.checked) selectedTxIds.add(id); else selectedTxIds.delete(id);
  document.getElementById('hist-bulk-delete').style.display = selectedTxIds.size > 0 ? 'inline-block' : 'none';
  document.getElementById('hist-bulk-pago').style.display = selectedTxIds.size > 0 ? 'inline-block' : 'none';
}

function toggleSelectAll() {
  const checked = document.getElementById('hist-select-all').checked;
  document.querySelectorAll('#hist-cards input[type="checkbox"]').forEach(c => {
    c.checked = checked;
    const id = c.dataset.txid;
    if (checked) selectedTxIds.add(id); else selectedTxIds.delete(id);
  });
  document.getElementById('hist-bulk-delete').style.display = selectedTxIds.size > 0 ? 'inline-block' : 'none';
  document.getElementById('hist-bulk-pago').style.display = selectedTxIds.size > 0 ? 'inline-block' : 'none';
}

function bulkDelete() {
  if (!selectedTxIds.size) return;
  const parcelasSel = S.transactions.filter(t => selectedTxIds.has(t.id) && ((t.parcelaTotal && t.parcelaTotal > 1) || (t.parcela && /\d+\/\d+/.test(t.parcela)) || (t.desc && /\(\d+\/\d+\)/.test(t.desc)))).length;
  const avisoParc = parcelasSel > 0 ? `\n\n⚠️ Inclui ${parcelasSel} parcela(s) de parcelamentos.` : '';
  if (!confirm(`Excluir ${selectedTxIds.size} transações selecionadas?${avisoParc}`)) return;
  selectedTxIds.forEach(id => S.deletedIds.push({id, collection: 'transactions', deletedAt: new Date().toISOString()}));
  S.transactions = S.transactions.filter(t => !selectedTxIds.has(t.id));
  selectedTxIds.clear();
  save();
  renderHistorico();
  toast(`🗑️ Transações excluídas`);
}

function bulkMarcarPago() {
  if (!selectedTxIds.size) return;
  if (!confirm(`Marcar ${selectedTxIds.size} transações como pagas?`)) return;
  const now = new Date().toISOString();
  let count = 0;
  S.transactions.forEach(t => {
    if (selectedTxIds.has(t.id) && t.pago !== true) {
      t.pago = true;
      t.updatedAt = now;
      count++;
    }
  });
  selectedTxIds.clear();
  save();
  renderHistorico();
  toast(`${count} transações marcadas como pagas`);
}

function clearFiltros() {
  ['fil-search','fil-tipo','fil-pagamento','fil-conta','fil-custo','fil-status','fil-recorrente'].forEach(id => { const el=document.getElementById(id); if(el)el.value=''; });
  histMesSel = null;
  renderHistorico();
}

// ─── FATURA MODAL ─────────────────────────────────────────────────────────────
function abrirFaturaModal() {
  const now = new Date();

  // Titular filter
  const titSel = document.getElementById('fat-titular');
  titSel.innerHTML = `<option value="">Todos Titulares</option>
    <option value="${S.settings.u1}">${S.settings.u1}</option>
    <option value="${S.settings.u2}">${S.settings.u2}</option>`;

  // Cartões (todos inicialmente)
  filtrarCartoesFatura();

  const mesSel = document.getElementById('fat-mes');
  mesSel.innerHTML = MESES.map((m,i) => `<option value="${i}" ${i===now.getMonth()?'selected':''}>${m}</option>`).join('');

  const anoSel = document.getElementById('fat-ano');
  const years = new Set([now.getFullYear()]);
  S.transactions.forEach(t => {
    const y = parseInt((t.date||'').substring(0,4));
    if (y >= 2000 && y <= 2100) years.add(y);
  });
  anoSel.innerHTML = [...years].sort((a,b)=>b-a).map(y => `<option value="${y}">${y}</option>`).join('');
  anoSel.value = now.getFullYear(); // sempre abre no ano atual, não no maior (parcelas futuras puxavam pra 2027)

  // Inst. financeira (contas bancárias para débito)
  const instSel = document.getElementById('fat-inst');
  const contas = S.accounts.filter(a => a.accountType !== 'cartao');
  instSel.innerHTML = '<option value="">Selecione a instituição</option>' +
    contas.map(a => `<option value="${a.id}">${BANKS[a.bank].name} — ${escapeHtml(a.label)}</option>`).join('');

  document.getElementById('fat-data-pgto').value = now.toISOString().split('T')[0];

  renderFaturaDetail();
  document.getElementById('fatura-modal').style.display = 'block';
}

function filtrarCartoesFatura() {
  const titular = document.getElementById('fat-titular').value;
  const cartoes = S.accounts.filter(a => a.accountType === 'cartao');
  const filtered = titular ? cartoes.filter(c => c.owner === titular) : cartoes;
  const sel = document.getElementById('fat-cartao');
  const prev = sel.value;
  sel.innerHTML = filtered.map(c => `<option value="${c.id}">${BANKS[c.bank].name} Cartão — ${escapeHtml(c.label)}</option>`).join('');
  if (prev && filtered.some(c => c.id === prev)) sel.value = prev;

  // Filtrar inst. financeira pelo titular também
  const instSel = document.getElementById('fat-inst');
  const contas = S.accounts.filter(a => a.accountType !== 'cartao');
  const contasFiltradas = titular ? contas.filter(a => a.owner === titular) : contas;
  const prevInst = instSel.value;
  instSel.innerHTML = '<option value="">Selecione a instituição</option>' +
    contasFiltradas.map(a => `<option value="${a.id}">${BANKS[a.bank].name} — ${escapeHtml(a.label)}</option>`).join('');
  if (prevInst && contasFiltradas.some(a => a.id === prevInst)) instSel.value = prevInst;

  renderFaturaDetail();
}

function fecharFaturaModal() { document.getElementById('fatura-modal').style.display = 'none'; }

// Busca pagamentos já feitos dessa fatura (usa faturaCartaoId+faturaYM; fallback no formato antigo via desc)
function getPagamentosFatura(cartaoId, mes, ano) {
  const faturaYM = `${ano}-${String(mes+1).padStart(2,'0')}`;
  const cartao = S.accounts.find(a => a.id === cartaoId);
  const bankName = cartao && BANKS[cartao.bank] ? BANKS[cartao.bank].name : '';
  const descPrefix = `Pagamento Fatura ${bankName} — ${MESES[mes]}/${ano}`;
  return S.transactions.filter(p => {
    if (p.category !== 'Taxas' || p.subcategory !== 'Pgto Fatura') return false;
    if (p.faturaCartaoId && p.faturaYM) {
      return p.faturaCartaoId === cartaoId && p.faturaYM === faturaYM;
    }
    // Fallback pra registros antigos sem faturaCartaoId/faturaYM
    return p.desc && p.desc.startsWith(descPrefix);
  });
}

function renderFaturaDetail() {
  const cartaoId = document.getElementById('fat-cartao').value;
  const mes = parseInt(document.getElementById('fat-mes').value);
  const ano = parseInt(document.getElementById('fat-ano').value);
  const faturaYM = `${ano}-${String(mes+1).padStart(2,'0')}`;

  const txs = getCreditTxByFatura(faturaYM, cartaoId)
    .sort((a,b) => new Date(a.date) - new Date(b.date));

  const total = txs.reduce((s,t) => s + amountBrl(t), 0);
  document.getElementById('fat-total').textContent = brl(total);
  document.getElementById('fat-total').style.color = total > 0 ? '#059669' : 'var(--muted)';
  document.getElementById('fat-count').textContent = `${txs.length} lançamentos encontrados`;

  // ─── Info de pagamento ───────────────────────────────────────────────
  const infoEl = document.getElementById('fat-pagamento-info');
  const pagarBtn = document.getElementById('fat-pagar-btn');
  const pagamentos = cartaoId ? getPagamentosFatura(cartaoId, mes, ano) : [];
  const totalPago = pagamentos.reduce((s,p) => s + Math.abs(p.amount||0), 0);
  const allCreditPaid = txs.length > 0 && txs.every(t => t.pago === true);
  const faturaPagaIntegral = total > 0 && (totalPago >= total - 0.01);
  const faturaPagaSemRegistro = total > 0 && allCreditPaid && pagamentos.length === 0;

  if (pagamentos.length > 0) {
    const linhas = pagamentos.map(p => {
      const conta = S.accounts.find(a => a.id === p.accountId);
      const bancoNome = conta ? `${BANKS[conta.bank]?.name || conta.bank} — ${conta.label}` : '—';
      const parcialTag = (p.desc||'').includes('(Parcial)') ? ' <span style="color:#d97706;font-weight:700;">(Parcial)</span>' : '';
      return `<div style="display:flex;justify-content:space-between;align-items:center;padding:6px 0;border-top:1px solid #d1fae5;font-size:12px;">
        <span style="color:var(--text-2);">🏦 ${bancoNome}${parcialTag}<br><span style="color:var(--muted);font-size:11px;">📅 ${fmtDate(p.date)}</span></span>
        <span style="color:#059669;font-weight:700;">${brl(Math.abs(p.amount||0))}</span>
      </div>`;
    }).join('');
    const restante = Math.max(0, total - totalPago);
    const statusLabel = faturaPagaIntegral
      ? `<span style="color:#059669;">✅ Fatura paga</span>`
      : `<span style="color:#d97706;">⚠️ Parcialmente paga — restante ${brl(restante)}</span>`;
    infoEl.innerHTML = `<div style="border:2px solid #a7f3d0;background:var(--tint-green);border-radius:12px;padding:12px 16px;">
      <p style="font-size:13px;font-weight:700;margin-bottom:4px;">${statusLabel}</p>
      <p style="font-size:11px;color:var(--text-3);margin-bottom:4px;">Pago: <b>${brl(totalPago)}</b> de ${brl(total)}</p>
      ${linhas}
    </div>`;
    infoEl.style.display = 'block';
  } else if (faturaPagaSemRegistro) {
    infoEl.innerHTML = `<div style="border:2px solid #bfdbfe;background:var(--tint-blue);border-radius:12px;padding:12px 16px;">
      <p style="font-size:13px;font-weight:700;color:#1d4ed8;">ℹ️ Fatura marcada como paga (sem registro de pagamento)</p>
      <p style="font-size:11px;color:var(--text-3);margin-top:4px;">Todos os lançamentos estão marcados como pagos, mas não há transação de Pgto Fatura. Pode ter sido importada já quitada.</p>
    </div>`;
    infoEl.style.display = 'block';
  } else {
    infoEl.style.display = 'none';
    infoEl.innerHTML = '';
  }

  // Trava visual: desabilita botão se integralmente paga
  if (pagarBtn) {
    if (faturaPagaIntegral || faturaPagaSemRegistro) {
      pagarBtn.disabled = true;
      pagarBtn.style.background = 'var(--border-2)';
      pagarBtn.style.cursor = 'not-allowed';
      pagarBtn.textContent = 'Fatura já paga';
    } else {
      pagarBtn.disabled = false;
      pagarBtn.style.background = '#059669';
      pagarBtn.style.cursor = 'pointer';
      pagarBtn.textContent = 'Pagar Fatura';
    }
  }

  const itemsEl = document.getElementById('fat-items');
  if (!txs.length) {
    itemsEl.innerHTML = '<p style="font-size:13px;color:var(--muted);text-align:center;padding:20px;">Nenhum lançamento neste período</p>';
    return;
  }

  itemsEl.innerHTML = txs.map(t => {
    const val = amountBrl(t);
    const parcelaTag = t.parcela ? `<span style="color:#4f46e5;font-weight:600;"> · 🔄 ${t.parcela}</span>` : '';
    const custoTag = t.custoTipo ? ` · ${t.custoTipo==='fixo'?'📌 Fixo':'📊 Var'}` : '';
    return `<div style="display:flex;align-items:center;padding:10px 14px;border-bottom:1px solid var(--bg);gap:10px;">
      <div style="flex:1;min-width:0;">
        <p style="font-size:13px;font-weight:600;color:var(--text);">${escapeHtml(t.desc)}</p>
        <p style="font-size:11px;color:var(--muted);">${fmtDate(t.date)} · ${t.category}${t.subcategory?' › '+t.subcategory:''}${parcelaTag}${custoTag} · ${t.pago===false?'<span style="color:#f59e0b;">Pendente</span>':'<span style="color:#10b981;">Pago</span>'}</p>
      </div>
      <p style="font-size:14px;font-weight:700;color:${val < 0 ? '#059669' : '#e11d48'};flex-shrink:0;">${val < 0 ? '+' : ''}${brl(Math.abs(val))}</p>
      <div style="display:flex;gap:2px;flex-shrink:0;">
        <button onclick="fecharFaturaModal();abrirEditModal('${t.id}')" style="background:none;border:none;cursor:pointer;padding:4px;border-radius:6px;font-size:13px;" title="Editar" class="ico-btn" aria-label="Editar"><svg class="ico-sm"><use href="#i-edit"/></svg></button>
        <button onclick="del('${t.id}');renderFaturaDetail()" style="background:none;border:none;cursor:pointer;padding:4px;border-radius:6px;font-size:13px;" title="Excluir" class="ico-btn" aria-label="Excluir"><svg class="ico-sm"><use href="#i-trash"/></svg></button>
      </div>
    </div>`;
  }).join('');
}

let fatPgtoTipo = 'integral';

function setFatPgtoTipo(tipo) {
  fatPgtoTipo = tipo;
  const isIntegral = tipo === 'integral';
  document.getElementById('fat-tipo-integral').style.border = isIntegral ? '2px solid #059669' : '2px solid #e2e8f0';
  document.getElementById('fat-tipo-integral').style.background = isIntegral ? 'var(--tint-green)' : 'var(--surface)';
  document.getElementById('fat-tipo-integral').style.color = isIntegral ? '#059669' : 'var(--text-3)';
  document.getElementById('fat-tipo-parcial').style.border = !isIntegral ? '2px solid #f59e0b' : '2px solid #e2e8f0';
  document.getElementById('fat-tipo-parcial').style.background = !isIntegral ? '#fffbeb' : 'var(--surface)';
  document.getElementById('fat-tipo-parcial').style.color = !isIntegral ? '#d97706' : 'var(--text-3)';
  document.getElementById('fat-valor-parcial-row').style.display = isIntegral ? 'none' : 'block';
  if (!isIntegral) {
    const totalEl = document.getElementById('fat-total');
    const totalText = totalEl.textContent.replace(/[^\d,.-]/g, '').replace('.','').replace(',','.');
    document.getElementById('fat-valor-parcial').placeholder = totalText;
  }
}

function pagarFatura() {
  const cartaoId = document.getElementById('fat-cartao').value;
  const mes = parseInt(document.getElementById('fat-mes').value);
  const ano = parseInt(document.getElementById('fat-ano').value);
  const instId = document.getElementById('fat-inst').value;
  const dataPgto = document.getElementById('fat-data-pgto').value;
  const faturaYM = `${ano}-${String(mes+1).padStart(2,'0')}`;

  if (!instId) { toast('Selecione a instituição financeira'); return; }
  if (!dataPgto) { toast('Selecione a data de pagamento'); return; }

  const faturaTxs = getCreditTxByFatura(faturaYM, cartaoId);
  const total = faturaTxs.reduce((s,t) => s + amountBrl(t), 0);

  // ─── Trava anti-duplicidade ──────────────────────────────────────────
  const pagamentosExistentes = getPagamentosFatura(cartaoId, mes, ano);
  const jaPago = pagamentosExistentes.reduce((s,p) => s + Math.abs(p.amount||0), 0);
  const allCreditPaid = faturaTxs.length > 0 && faturaTxs.every(t => t.pago === true);
  const faturaJaPagaIntegral = total > 0 && jaPago >= total - 0.01;
  const faturaPagaSemRegistro = total > 0 && allCreditPaid && pagamentosExistentes.length === 0;

  if (faturaJaPagaIntegral || faturaPagaSemRegistro) {
    toast('⛔ Esta fatura já está paga. Exclua o pagamento no Histórico antes de pagar de novo.');
    return;
  }

  let valorPago = total;
  const isParcial = fatPgtoTipo === 'parcial';

  if (isParcial) {
    const input = parseFloat(document.getElementById('fat-valor-parcial').value);
    if (!input || input <= 0) { toast('Informe o valor pago'); return; }
    if (input > total) { toast('Valor pago não pode ser maior que o total da fatura'); return; }
    // Trava parcial: não pode exceder o saldo restante (total - já pago anteriormente)
    const restante = total - jaPago;
    if (input > restante + 0.01) {
      toast(`⛔ Valor excede o saldo restante. Já pago: ${brl(jaPago)}. Restante: ${brl(restante)}.`);
      return;
    }
    valorPago = input;
  } else {
    // Integral: se já tem pagamento parcial, cobra só o restante
    valorPago = total - jaPago;
    if (valorPago <= 0.01) {
      toast('⛔ Esta fatura já está paga.');
      return;
    }
  }

  // Mark all credit transactions as paid
  let count = 0;
  faturaTxs.forEach(t => {
    t.pago = true;
    count++;
  });

  // Create debit transaction with the actual amount paid
  if (valorPago > 0) {
    const cartao = S.accounts.find(a => a.id === cartaoId);
    const bankName = cartao ? BANKS[cartao.bank].name : '';
    const titular = cartao ? cartao.owner : S.settings.u1;
    const descSuffix = isParcial ? ' (Parcial)' : '';
    S.transactions.push({
      id: Date.now().toString(),
      type: 'despesa',
      desc: `Pagamento Fatura ${bankName} — ${MESES[mes]}/${ano}${descSuffix}`,
      amount: valorPago,
      category: 'Taxas',
      subcategory: 'Pgto Fatura',
      date: dataPgto,
      user: titular,
      accountId: instId,
      formaPgto: 'debito',
      pago: true,
      faturaCartaoId: cartaoId,
      faturaYM: faturaYM,
      at: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    });
  }

  save();
  renderFaturaDetail();
  const msg = isParcial
    ? `Fatura paga parcialmente! R$ ${valorPago.toFixed(2).replace('.',',')} de R$ ${total.toFixed(2).replace('.',',')} — ${count} lançamentos marcados`
    : `Fatura paga! ${count} lançamentos marcados como pagos`;
  toast(msg);
  // Reset to integral
  setFatPgtoTipo('integral');
}

// ─── IMPORT MODAL ─────────────────────────────────────────────────────────────
function filtrarContasImport() {
  const titular = document.getElementById('imp-titular').value;
  const sel = document.getElementById('imp-conta');
  const filtered = titular ? S.accounts.filter(a => a.owner === titular) : S.accounts;
  sel.innerHTML = filtered.map(a => {
    const b = BANKS[a.bank] || { name: a.bank };
    const icon = a.accountType === 'cartao' ? '💳' : '🏦';
    return `<option value="${a.id}">${icon} ${b.name} — ${escapeHtml(a.label)} [${escapeHtml(a.owner)}]</option>`;
  }).join('');
}

function abrirImportModal() {
  // Popular titular
  const titSel = document.getElementById('imp-titular');
  const { u1, u2 } = S.settings;
  titSel.innerHTML = `<option value="">Todos</option><option value="${u1}">${u1}</option><option value="${u2}">${u2}</option>`;
  filtrarContasImport();
  document.getElementById('imp-file').value = '';
  document.getElementById('imp-file-label').innerHTML = 'Clique para selecionar arquivo<br><span style="font-size:11px;">.txt, .csv, .xlsx</span>';
  document.getElementById('imp-preview').style.display = 'none';
  document.getElementById('imp-btn-confirm').disabled = true;
  document.getElementById('imp-btn-confirm').style.opacity = '0.5';
  pendingImportTxs = [];
  document.getElementById('import-modal').style.display = 'block';
}

function fecharImportModal() { document.getElementById('import-modal').style.display = 'none'; pendingImportTxs = []; }

function onImportFileSelect(input) {
  const file = input.files[0];
  if (!file) return;
  document.getElementById('imp-file-label').textContent = file.name;

  const ext = file.name.split('.').pop().toLowerCase();
  const reader = new FileReader();

  if (ext === 'csv' || ext === 'txt') {
    reader.onload = e => parseCSV(e.target.result);
    reader.readAsText(file, 'UTF-8');
  } else if (ext === 'xlsx' || ext === 'xls') {
    toast('Para importar Excel, use o formato CSV (salve como CSV no Excel)');
  } else {
    toast('Formato não suportado. Use .csv, .txt ou .xlsx');
  }
}

function parseCSV(text) {
  const lines = text.trim().split('\n').map(l => l.split(/[;\t,]/).map(c => c.trim().replace(/^"|"$/g, '')));
  if (lines.length < 2) { toast('Arquivo vazio ou sem dados'); return; }

  const contaId = document.getElementById('imp-conta').value;
  const tipoDefault = document.getElementById('imp-tipo').value;
  const acc = S.accounts.find(a => a.id === contaId);
  const isCartao = acc && acc.accountType === 'cartao';

  // Try to detect columns: date, description, amount
  const header = lines[0].map(h => h.toLowerCase());
  let dateCol = header.findIndex(h => h.includes('data'));
  let descCol = header.findIndex(h => h.includes('descri') || h.includes('histor') || h.includes('lanca'));
  let amountCol = header.findIndex(h => h.includes('valor') || h.includes('amount'));

  // If no header detected, assume: 0=date, 1=desc, 2=amount
  const hasHeader = dateCol >= 0 || descCol >= 0;
  if (dateCol < 0) dateCol = 0;
  if (descCol < 0) descCol = 1;
  if (amountCol < 0) amountCol = lines[0].length - 1;

  const startRow = hasHeader ? 1 : 0;
  pendingImportTxs = [];

  for (let i = startRow; i < lines.length; i++) {
    const row = lines[i];
    if (row.length < 3) continue;

    const rawDate = row[dateCol];
    const rawDesc = row[descCol];
    let rawAmount = row[amountCol];

    if (!rawDate || !rawDesc || !rawAmount) continue;

    // Parse date
    let date = '';
    const dmMatch = rawDate.match(/(\d{2})[\/\-](\d{2})[\/\-](\d{4})/);
    const isoMatch = rawDate.match(/(\d{4})[\/\-](\d{2})[\/\-](\d{2})/);
    if (isoMatch) date = `${isoMatch[1]}-${isoMatch[2]}-${isoMatch[3]}`;
    else if (dmMatch) date = `${dmMatch[3]}-${dmMatch[2]}-${dmMatch[1]}`;
    else continue;

    // Parse amount
    rawAmount = rawAmount.replace(/[R$\s]/g, '').replace('.', '').replace(',', '.');
    const amount = Math.abs(parseFloat(rawAmount));
    if (isNaN(amount) || amount === 0) continue;

    pendingImportTxs.push({
      id: `imp_file_${Date.now()}_${i}`,
      type: tipoDefault,
      desc: rawDesc,
      amount,
      category: 'Outros',
      subcategory: '',
      date,
      user: acc?.owner || document.getElementById('imp-titular').value || S.settings.u1,
      accountId: contaId,
      formaPgto: isCartao ? 'credito' : (tipoDefault === 'despesa' ? 'debito' : null),
      pago: true,
      updatedAt: new Date().toISOString()
    });
  }

  // Show preview
  if (pendingImportTxs.length) {
    document.getElementById('imp-preview').style.display = 'block';
    document.getElementById('imp-preview-count').textContent = pendingImportTxs.length;
    document.getElementById('imp-preview-list').innerHTML = pendingImportTxs.slice(0, 10).map(t =>
      `<div style="display:flex;justify-content:space-between;padding:8px 12px;border-bottom:1px solid var(--bg);font-size:12px;">
        <span style="color:var(--text-2);">${fmtDate(t.date)} — ${escapeHtml(t.desc)}</span>
        <span style="font-weight:700;color:#e11d48;">R$ ${t.amount.toFixed(2)}</span>
      </div>`
    ).join('') + (pendingImportTxs.length > 10 ? `<p style="padding:8px 12px;font-size:11px;color:var(--muted);">+ ${pendingImportTxs.length - 10} mais...</p>` : '');
    document.getElementById('imp-btn-confirm').disabled = false;
    document.getElementById('imp-btn-confirm').style.opacity = '1';
  } else {
    toast('Nenhuma transação encontrada no arquivo');
  }
}

function confirmarImport() {
  if (!pendingImportTxs.length) return;
  pendingImportTxs.forEach(t => S.transactions.push(t));
  save();
  toast(`✅ ${pendingImportTxs.length} transações importadas!`);
  fecharImportModal();
  renderHistorico();
}

// ─── EXPORT EXCEL ─────────────────────────────────────────────────────────────
function exportarExcel() {
  const txs = getFilteredTxs();
  if (!txs.length) { toast('Nenhuma transação para exportar'); return; }

  // Build CSV content
  const headers = ['Data','Descrição','Tipo','Categoria','Subcategoria','Forma Pgto','Conta','Titular','Valor','Status'];
  const rows = txs.map(t => {
    const acc = S.accounts.find(a => a.id === t.accountId);
    const accLabel = acc ? `${BANKS[acc.bank]?.name || ''} — ${acc.label}` : '';
    return [
      t.date,
      `"${(t.desc||'').replace(/"/g,'""')}"`,
      t.type,
      t.category,
      t.subcategory || '',
      t.formaPgto || '',
      `"${accLabel}"`,
      t.user,
      t.amount.toFixed(2).replace('.', ','),
      t.type === 'despesa' ? (t.pago === false ? 'Pendente' : 'Paga') : ''
    ].join(';');
  });

  const csv = '\uFEFF' + headers.join(';') + '\n' + rows.join('\n');
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  const periodo = histMesSel !== null ? `${MESES[histMesSel]}_${histAno}` : `${histAno}`;
  a.href = url;
  a.download = `transacoes_${periodo}.csv`;
  a.click();
  URL.revokeObjectURL(url);
  toast(`📥 Exportadas ${txs.length} transações`);
}
