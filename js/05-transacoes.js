// FinançasCasal — 05-transacoes.js
// Módulo carregado em ordem por index.html (escopo global compartilhado). Não use import/export.

// ─── TYPE SELECTOR ────────────────────────────────────────────────────────────
function setType(t) {
  currentType = t;
  ['receita','despesa','transferencia'].forEach(x => document.getElementById('btn-'+x).className = 'type-btn');
  document.getElementById('btn-'+t).classList.add('active-' + (t==='transferencia'?'invest':t));

  const isDespesa = t === 'despesa';
  const isTransf  = t === 'transferencia';
  document.getElementById('f-row-pago').style.display       = isDespesa ? 'block' : 'none';
  document.getElementById('f-row-pagamento').style.display  = isDespesa ? 'block' : 'none';
  document.getElementById('f-row-recorrente').style.display = isDespesa ? 'block' : 'none';
  document.getElementById('f-row-compartilhada').style.display = isDespesa ? 'block' : 'none';
  if (!isDespesa) { document.getElementById('f-recorrente').checked = false; document.getElementById('f-recorrente-meses-row').style.display = 'none'; }
  document.getElementById('f-row-custo-tipo').style.display = isDespesa ? 'block' : 'none';
  document.getElementById('f-row-parcela').style.display    = 'none';
  document.getElementById('f-row-conta-destino').style.display = isTransf ? 'block' : 'none';
  document.getElementById('f-label-conta').textContent      = isTransf ? 'Conta Origem' : (isDespesa ? 'Conta' : 'Conta / Carteira');

  // Hide categoria for transfers
  const catRow = document.getElementById('f-categoria').closest('div[style*="margin-bottom"]') || document.getElementById('f-categoria').parentElement.parentElement;
  if (catRow) catRow.style.display = isTransf ? 'none' : '';

  const sel = document.getElementById('f-categoria');
  if (t === 'despesa') {
    sel.innerHTML = Object.keys(getDespesaCats()).map(c => `<option value="${c}">${c}</option>`).join('');
  } else if (!isTransf) {
    const catList = t === 'investimento' ? getInvestCats() : (getReceitaCats() || []);
    sel.innerHTML = catList.map(c => `<option value="${c}">${c}</option>`).join('');
  }
  if (!isTransf) onCatChange();

  if (isDespesa) setFormaPgto('debito');
  else refreshContaSelect('conta');

  if (isTransf) {
    // Populate destination account select (only bank accounts, filtered by titular)
    const titular = document.getElementById('f-usuario')?.value;
    const destSel = document.getElementById('f-conta-destino');
    let destAccounts = S.accounts.filter(a => a.accountType === 'conta');
    if (titular) destAccounts = destAccounts.filter(a => a.owner === titular);
    destSel.innerHTML = destAccounts.map(a => `<option value="${a.id}">${BANKS[a.bank]?.name || a.bank} — ${escapeHtml(a.label)}</option>`).join('');
  }
}

let custoTipo = 'variavel';

function setFormaPgto(tipo) {
  formaPgto = tipo;
  const isD = tipo === 'debito';
  document.getElementById('btn-debito').style.cssText  = `padding:12px;border-radius:10px;border:2px solid ${isD?'#4f46e5':'var(--border)'};background:${isD?'var(--tint-indigo)':'var(--surface)'};color:${isD?'#4f46e5':'var(--text-3)'};font-weight:700;font-size:14px;cursor:pointer;`;
  document.getElementById('btn-credito').style.cssText = `padding:12px;border-radius:10px;border:2px solid ${!isD?'#4f46e5':'var(--border)'};background:${!isD?'var(--tint-indigo)':'var(--surface)'};color:${!isD?'#4f46e5':'var(--text-3)'};font-weight:700;font-size:14px;cursor:pointer;`;
  refreshContaSelect(isD ? 'conta' : 'cartao');
  // Show/hide parcelamento (only for crédito)
  document.getElementById('f-row-parcela').style.display = !isD ? 'block' : 'none';
  if (isD) {
    document.getElementById('f-parcelado').checked = false;
    document.getElementById('f-parcela-num-row').style.display = 'none';
  }
}

function toggleParcelas() {
  const checked = document.getElementById('f-parcelado').checked;
  document.getElementById('f-parcela-num-row').style.display = checked ? 'block' : 'none';
  if (!checked) document.getElementById('f-num-parcelas').value = '1';
}

function setCustoTipo(tipo) {
  custoTipo = tipo;
  const isF = tipo === 'fixo';
  document.getElementById('btn-custo-fixo').style.cssText = `padding:10px;border-radius:10px;border:2px solid ${isF?'#4f46e5':'var(--border)'};background:${isF?'var(--tint-indigo)':'var(--surface)'};color:${isF?'#4f46e5':'var(--text-3)'};font-weight:700;font-size:13px;cursor:pointer;`;
  document.getElementById('btn-custo-variavel').style.cssText = `padding:10px;border-radius:10px;border:2px solid ${!isF?'#4f46e5':'var(--border)'};background:${!isF?'var(--tint-indigo)':'var(--surface)'};color:${!isF?'#4f46e5':'var(--text-3)'};font-weight:700;font-size:13px;cursor:pointer;`;
}

// Auto-detect custoTipo based on category
const CATEGORIAS_FIXAS = ['Assinaturas', 'Casa', 'Mercado', 'Saúde'];
function autoCustoTipo(cat) {
  return CATEGORIAS_FIXAS.includes(cat) ? 'fixo' : 'variavel';
}

function onCatChange() {
  const cat = document.getElementById('f-categoria')?.value;
  const sub = document.getElementById('f-subcategoria');
  if (!sub) return;
  const subs = (currentType === 'despesa' ? getDespesaCats()[cat] : null) || [];
  sub.innerHTML = '<option value="">Sem subcategoria</option>' + subs.map(s => `<option value="${s}">${s}</option>`).join('');
  // Auto-detect custoTipo
  if (currentType === 'despesa' && cat) setCustoTipo(autoCustoTipo(cat));
}

// ─── LANÇAMENTO RÁPIDO: sugestões pelo histórico ──────────────────────────────
// Agrupa transações passadas por descrição normalizada e guarda o "perfil" mais recente
// de cada uma (tipo, titular, conta, forma, categoria, subcategoria, custo, valor).
let _suggestIdx = null, _suggestIdxKey = '';
let _suggestList = [], _suggestSel = -1;

function normDesc(d) {
  return String(d || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/\s*\(\d+\/\d+\)\s*$/, '')          // tira "(03/12)" das parcelas
    .replace(/\s+/g, ' ').trim();
}
function buildSuggestIndex() {
  const key = S.transactions.length + ':' + (S.transactions[S.transactions.length - 1] || {}).id;
  if (_suggestIdx && _suggestIdxKey === key) return _suggestIdx;
  const map = new Map();
  for (const t of S.transactions) {
    if (!t || t.isTransfer || t.type === 'transferencia' || !t.desc) continue;
    const k = normDesc(t.desc);
    if (k.length < 2) continue;
    const cur = map.get(k);
    const date = t.date || '';
    if (!cur) map.set(k, { key: k, count: 1, last: date, t });
    else { cur.count++; if (date > cur.last) { cur.last = date; cur.t = t; } }
  }
  _suggestIdx = Array.from(map.values());
  _suggestIdxKey = key;
  return _suggestIdx;
}
function findSuggestions(q) {
  const nq = normDesc(q);
  if (nq.length < 2) return [];
  const idx = buildSuggestIndex();
  const scored = [];
  for (const e of idx) {
    let score = 0;
    if (e.key === nq) score = 3;
    else if (e.key.startsWith(nq)) score = 2;
    else if (e.key.includes(nq)) score = 1;
    else continue;
    scored.push({ e, score });
  }
  scored.sort((a, b) => b.score - a.score || b.e.count - a.e.count || (b.e.last > a.e.last ? 1 : -1));
  return scored.slice(0, 6).map(x => x.e);
}
function suggestMeta(t) {
  const acc = S.accounts.find(a => a.id === t.accountId);
  const bankName = acc ? ((BANKS[acc.bank] || {}).name || acc.bank) : '';
  const accName = !acc ? '' : (String(acc.label || '').toLowerCase().includes(String(bankName).toLowerCase()) ? acc.label : `${bankName} ${acc.label || ''}`.trim());
  const cat = [t.category, t.subcategory].filter(Boolean).join(' › ');
  const forma = t.type === 'despesa' ? (t.formaPgto === 'credito' ? 'Crédito' : 'Débito') : (t.type === 'receita' ? 'Receita' : '');
  return [cat, accName, forma, t.user].filter(Boolean).join(' · ');
}
function showDescSuggestions() {
  const inp = document.getElementById('f-descricao');
  const box = document.getElementById('desc-suggest');
  if (!inp || !box) return;
  _suggestList = findSuggestions(inp.value);
  if (!_suggestList.length) { hideDescSuggestions(); return; }
  _suggestSel = 0;
  const nq = normDesc(inp.value);
  box.innerHTML = _suggestList.map((e, i) => {
    const t = e.t;
    const d = escapeHtml(t.desc.replace(/\s*\(\d+\/\d+\)\s*$/, ''));
    const nd = normDesc(t.desc);
    const pos = nd.indexOf(nq);
    const hl = pos >= 0 ? d.slice(0, pos) + '<mark>' + d.slice(pos, pos + nq.length) + '</mark>' + d.slice(pos + nq.length) : d;
    const amt = (t.currency === 'USD' ? 'US$ ' : 'R$ ') + Number(t.amount || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2 });
    return `<div class="suggest-item${i === 0 ? ' active' : ''}" data-i="${i}" onmousedown="event.preventDefault();applySuggestion(${i})">
      <div><div class="s-desc">${hl}</div><div class="s-meta">${escapeHtml(suggestMeta(t))}</div></div>
      <div class="s-right"><div class="s-amt">${amt}</div><div class="s-cnt">${e.count}× · último ${e.last ? e.last.split('-').reverse().join('/') : ''}</div></div>
    </div>`;
  }).join('');
  box.style.display = 'block';
}
function hideDescSuggestions() {
  const box = document.getElementById('desc-suggest');
  if (box) box.style.display = 'none';
  _suggestSel = -1;
}
function highlightSuggestion(i) {
  _suggestSel = i;
  document.querySelectorAll('#desc-suggest .suggest-item').forEach((el, j) => el.classList.toggle('active', j === i));
}
function descKeydown(ev) {
  const open = document.getElementById('desc-suggest').style.display !== 'none' && _suggestList.length;
  if (ev.key === 'ArrowDown' && open) { ev.preventDefault(); highlightSuggestion((_suggestSel + 1) % _suggestList.length); }
  else if (ev.key === 'ArrowUp' && open) { ev.preventDefault(); highlightSuggestion((_suggestSel - 1 + _suggestList.length) % _suggestList.length); }
  else if (ev.key === 'Escape' && open) { ev.preventDefault(); hideDescSuggestions(); }
  else if (ev.key === 'Enter' || ev.key === 'Tab') {
    if (open && _suggestSel >= 0) { ev.preventDefault(); applySuggestion(_suggestSel); }
    else if (ev.key === 'Enter') { ev.preventDefault(); document.getElementById('f-valor').focus(); }
  }
}
const setSel = (id, val) => { const el = document.getElementById(id); if (!el) return false;
  if ([...el.options].some(o => o.value === val)) { el.value = val; return true; } return false; };
function applySuggestion(i) {
  const e = _suggestList[i]; if (!e) return;
  const t = e.t;
  hideDescSuggestions();
  document.getElementById('f-descricao').value = t.desc.replace(/\s*\(\d+\/\d+\)\s*$/, '');
  const tipo = t.type === 'receita' ? 'receita' : 'despesa';
  if (currentType !== tipo) setType(tipo);
  if (t.user && setSel('f-usuario', t.user)) onTitularChange();
  if (tipo === 'despesa') setFormaPgto(t.formaPgto === 'credito' ? 'credito' : 'debito');
  if (t.accountId) setSel('f-conta', t.accountId);
  if (t.category && setSel('f-categoria', t.category)) { onCatChange(); if (t.subcategory) setSel('f-subcategoria', t.subcategory); }
  if (tipo === 'despesa' && t.custoTipo) setCustoTipo(t.custoTipo);
  const valor = document.getElementById('f-valor');
  if (!valor.value && t.amount) valor.value = Number(t.amount).toFixed(2);
  const hint = document.getElementById('f-desc-hint');
  hint.innerHTML = '✨ Preenchido pelo histórico: <b style="color:var(--text-2)">' + escapeHtml(suggestMeta(t)) + '</b>' +
    (e.count > 1 ? ` · ${e.count} lançamentos` : '') + ' — confira e ajuste se precisar.';
  hint.style.display = 'block';
  valor.focus(); valor.select();
}

// ─── SAVE TRANSACTION ─────────────────────────────────────────────────────────
function saveTransacao(e) {
  e.preventDefault();
  const contaId = document.getElementById('f-conta').value;
  const acc = S.accounts.find(a => a.id === contaId);
  const isUSD = acc && BANKS[acc.bank].currency === 'USD';
  const amountRaw = parseFloat(document.getElementById('f-valor').value);
  const isDespesa = currentType === 'despesa';
  const isTransf  = currentType === 'transferencia';
  const isCredito = isDespesa && formaPgto === 'credito';
  const isParcelado = isCredito && document.getElementById('f-parcelado').checked;
  const numParcelas = isParcelado ? parseInt(document.getElementById('f-num-parcelas').value) || 1 : 1;
  const desc = document.getElementById('f-descricao').value.trim() || (isTransf ? 'Transferência' : document.getElementById('f-categoria').value);
  const date = document.getElementById('f-data').value;

  // Handle transfer
  if (isTransf) {
    const destContaId = document.getElementById('f-conta-destino').value;
    if (!contaId || !destContaId) { toast('Selecione conta origem e destino'); return; }
    if (contaId === destContaId) { toast('Conta origem e destino devem ser diferentes'); return; }
    if (!amountRaw || amountRaw <= 0) { toast('Informe o valor da transferência'); return; }

    const origemAcc = S.accounts.find(a => a.id === contaId);
    const destAcc = S.accounts.find(a => a.id === destContaId);
    const origemNome = BANKS[origemAcc?.bank]?.name || 'Origem';
    const destNome = BANKS[destAcc?.bank]?.name || 'Destino';
    const transferId = 'tf_' + Date.now();
    const user = document.getElementById('f-usuario').value;
    const notas = document.getElementById('f-notas').value.trim();
    const descLabel = desc === 'Transferência' ? '' : ` - ${desc}`;

    // Saída da conta origem (despesa)
    S.transactions.push({
      id: transferId + '_out',
      type: 'despesa',
      desc: `Transferência → ${destNome}${descLabel}`,
      amount: amountRaw,
      currency: isUSD ? 'USD' : 'BRL',
      accountId: contaId,
      category: 'Transferência',
      subcategory: '',
      date, user, notes: notas,
      pago: true,
      formaPgto: 'debito',
      custoTipo: null,
      isTransfer: true,
      transferId,
      transferDir: 'saida',
      at: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    });

    // Entrada na conta destino (receita)
    S.transactions.push({
      id: transferId + '_in',
      type: 'receita',
      desc: `Transferência ← ${origemNome}${descLabel}`,
      amount: amountRaw,
      currency: BANKS[destAcc?.bank]?.currency === 'USD' ? 'USD' : 'BRL',
      accountId: destContaId,
      category: 'Transferência',
      subcategory: '',
      date, user, notes: notas,
      pago: true,
      isTransfer: true,
      transferId,
      transferDir: 'entrada',
      at: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    });

    save();
    document.getElementById('f-valor').value = '';
    document.getElementById('f-descricao').value = '';
    document.getElementById('f-notas').value = '';
    toast(`✅ Transferência de ${brl(amountRaw)} registrada! ${origemNome} → ${destNome}`);
    return;
  }

  const baseTx = {
    type:         currentType,
    currency:     isUSD ? 'USD' : 'BRL',
    accountId:    contaId || null,
    category:     document.getElementById('f-categoria').value,
    subcategory:  document.getElementById('f-subcategoria')?.value || '',
    date:         date,
    user:         document.getElementById('f-usuario').value,
    notes:        document.getElementById('f-notas').value.trim(),
    pago:         isDespesa ? document.getElementById('f-pago').checked : true,
    formaPgto:    isDespesa ? formaPgto : null,
    recorrente:   isDespesa ? document.getElementById('f-recorrente').checked : false,
    compartilhada: isDespesa ? document.getElementById('f-compartilhada').checked : false,
    custoTipo:    isDespesa ? custoTipo : null,
    at:           new Date().toISOString(),
    updatedAt:    new Date().toISOString()
  };

  if (isParcelado && numParcelas > 1) {
    // Create one transaction per parcela with correct faturaRef
    const card = S.accounts.find(a => a.id === contaId);
    const fechaDia = card ? parseInt(card.fecha) || 1 : 1;
    const parcelaAmount = Math.round((amountRaw / numParcelas) * 100) / 100;

    // Determine the first fatura month
    const d = new Date(date + 'T12:00:00');
    const day = d.getDate();
    let startMonth, startYear;
    if (day > fechaDia) {
      // After closing → goes to next month's fatura
      const tmp = new Date(d.getFullYear(), d.getMonth() + 1, 1);
      startMonth = tmp.getMonth();
      startYear = tmp.getFullYear();
    } else {
      // On or before closing → goes to current month's fatura
      startMonth = d.getMonth();
      startYear = d.getFullYear();
    }

    for (let i = 0; i < numParcelas; i++) {
      const fatMonth = new Date(startYear, startMonth + i, 1);
      const faturaRef = `${fatMonth.getFullYear()}-${String(fatMonth.getMonth() + 1).padStart(2, '0')}`;
      S.transactions.push({
        ...baseTx,
        // Cada parcela cai no seu mês (antes todas herdavam a data da compra e
        // o histórico/dashboard empilhavam as N parcelas no mês da compra)
        date: addMesesData(date, i),
        id: Date.now().toString() + '_' + i,
        desc: `${desc} (${String(i+1).padStart(2,'0')}/${String(numParcelas).padStart(2,'0')})`,
        amount: parcelaAmount,
        faturaRef,
        parcela: `${String(i+1).padStart(2,'0')}/${String(numParcelas).padStart(2,'0')}`,
        parcelaTotal: numParcelas,
      });
    }
    toast(`✅ ${numParcelas} parcelas salvas!`);
  } else if (baseTx.recorrente && document.getElementById('f-recorrente').checked) {
    // Recorrente: replicate same transaction for N months
    const meses = parseInt(document.getElementById('f-recorrente-meses').value) || 12;
    const startDate = new Date(date + 'T12:00:00');
    for (let i = 0; i < meses; i++) {
      const txDate = new Date(startDate.getFullYear(), startDate.getMonth() + i, startDate.getDate());
      const txDateStr = `${txDate.getFullYear()}-${String(txDate.getMonth()+1).padStart(2,'0')}-${String(txDate.getDate()).padStart(2,'0')}`;
      const tx = {
        ...baseTx,
        id: Date.now().toString() + '_rec_' + i,
        desc,
        amount: amountRaw,
        date: txDateStr,
        recorrente: true,
        pago: i === 0 ? baseTx.pago : false
      };
      if (isCredito) {
        tx.faturaRef = getTxFaturaRef(tx);
      }
      S.transactions.push(tx);
    }
    toast(`✅ ${meses} transações recorrentes criadas!`);
  } else {
    // Single transaction
    const tx = { ...baseTx, id: Date.now().toString(), desc, amount: amountRaw };
    if (isCredito) {
      tx.faturaRef = getTxFaturaRef(tx);
    }
    S.transactions.push(tx);
    toast('✅ Transação salva!');
  }

  save();
  document.getElementById('f-valor').value = '';
  document.getElementById('f-descricao').value = '';
  document.getElementById('f-notas').value = '';
  document.getElementById('f-pago').checked = false;
  document.getElementById('f-recorrente').checked = false;
  document.getElementById('f-compartilhada').checked = false;
  document.getElementById('f-recorrente-meses-row').style.display = 'none';
  document.getElementById('f-recorrente-meses').value = '12';
  document.getElementById('f-parcelado').checked = false;
  document.getElementById('f-parcela-num-row').style.display = 'none';
  setTodayDate();
  goto('dashboard');
}

// ─── DELETE ───────────────────────────────────────────────────────────────────
function del(id) {
  deleteId = id;
  // Aviso visível quando a transação é parcela de um parcelamento (não apaga outras parcelas)
  const aviso = document.getElementById('del-parcela-aviso');
  if (aviso) {
    const t = S.transactions.find(x => x.id === id);
    let xn = '';
    if (t) {
      if (t.parcela && /\d+\/\d+/.test(t.parcela)) xn = (t.parcela.match(/\d+\/\d+/) || [''])[0];
      else if (t.desc && /\(\d+\/\d+\)/.test(t.desc)) xn = (t.desc.match(/\((\d+\/\d+)\)/) || ['',''])[1];
    }
    const ehParcela = !!(t && ((t.parcelaTotal && t.parcelaTotal > 1) || xn));
    if (ehParcela) {
      const desc = (t.desc || '').replace(/\s*\(\d+\/\d+\)\s*$/, '').trim() || 'transação';
      aviso.innerHTML = `⚠️ Esta é a parcela <strong>${xn || '?'}</strong> de um parcelamento: «${desc}». Excluir esta parcela? As outras parcelas <strong>NÃO</strong> serão afetadas.`;
      aviso.style.display = 'block';
    } else {
      aviso.style.display = 'none';
      aviso.innerHTML = '';
    }
  }
  document.getElementById('del-modal').style.display = 'block';
}
function closeModal() { document.getElementById('del-modal').style.display = 'none'; deleteId = null; const a = document.getElementById('del-parcela-aviso'); if (a) { a.style.display = 'none'; a.innerHTML = ''; } }
function confirmarDelete() {
  // If deleting a transfer, also delete the linked transaction
  const tx = S.transactions.find(t => t.id === deleteId);
  if (tx && tx.isTransfer && tx.transferId) {
    S.transactions.filter(t => t.transferId === tx.transferId).forEach(t => S.deletedIds.push({id: t.id, collection: 'transactions', deletedAt: new Date().toISOString()}));
    S.transactions = S.transactions.filter(t => t.transferId !== tx.transferId);
  } else {
    S.deletedIds.push({ id: deleteId, collection: 'transactions', deletedAt: new Date().toISOString() });
    S.transactions = S.transactions.filter(t => t.id !== deleteId);
  }
  save(); closeModal(); renderHistorico();
  toast('🗑️ Transação excluída');
}

// ─── EDIT TRANSACTION ─────────────────────────────────────────────────────────
let editTxId = null;
let editFormaPgto = 'debito';

function abrirEditModal(id) {
  const t = S.transactions.find(x => x.id === id);
  if (!t) return;
  editTxId = id;
  editFormaPgto = t.formaPgto || 'debito';

  document.getElementById('ed-type').value = t.type;
  document.getElementById('ed-desc').value = t.desc || '';
  document.getElementById('ed-amount').value = t.amount;
  document.getElementById('ed-date').value = t.date;
  document.getElementById('ed-notes').value = t.notes || '';

  // Populate user select
  const userSel = document.getElementById('ed-user');
  userSel.innerHTML = `<option value="${S.settings.u1}">${S.settings.u1}</option><option value="${S.settings.u2}">${S.settings.u2}</option>`;
  userSel.value = t.user || S.settings.u1;

  // Set forma pgto first (this calls onEditTypeChange which rebuilds dropdowns)
  setEditFormaPgto(editFormaPgto);

  // Now set values AFTER dropdowns are built
  if (t.accountId) document.getElementById('ed-conta').value = t.accountId;
  if (t.category) {
    document.getElementById('ed-category').value = t.category;
    onEditCatChange();
    if (t.subcategory) document.getElementById('ed-subcategory').value = t.subcategory;
  }
  if (t.type === 'despesa') document.getElementById('ed-pago').checked = t.pago !== false;
  if (t.type === 'despesa') document.getElementById('ed-recorrente').checked = t.recorrente === true;
  if (t.type === 'despesa') document.getElementById('ed-compartilhada').checked = t.compartilhada === true;
  if (t.type === 'despesa') {
    setEditCustoTipo(t.custoTipo || autoCustoTipo(t.category));
  }

  document.getElementById('edit-modal').style.display = 'block';
}

function fecharEditModal() {
  document.getElementById('edit-modal').style.display = 'none';
  editTxId = null;
}

function onEditTypeChange() {
  const tipo = document.getElementById('ed-type').value;
  const isDespesa = tipo === 'despesa';

  document.getElementById('ed-row-pagamento').style.display = isDespesa ? 'block' : 'none';
  document.getElementById('ed-row-pago').style.display = isDespesa ? 'block' : 'none';
  document.getElementById('ed-row-recorrente').style.display = isDespesa ? 'block' : 'none';
  document.getElementById('ed-row-compartilhada').style.display = isDespesa ? 'block' : 'none';
  document.getElementById('ed-row-custo-tipo').style.display = isDespesa ? 'block' : 'none';

  // Populate conta select
  const contaSel = document.getElementById('ed-conta');
  let accounts = S.accounts;
  if (isDespesa) {
    accounts = editFormaPgto === 'credito'
      ? S.accounts.filter(a => a.accountType === 'cartao')
      : S.accounts.filter(a => a.accountType !== 'cartao');
  } else if (tipo === 'receita' || tipo === 'investimento') {
    accounts = S.accounts.filter(a => a.accountType !== 'cartao');
  }
  contaSel.innerHTML = accounts.map(a => {
    const b = BANKS[a.bank];
    return `<option value="${a.id}">${b ? b.name : a.bank} — ${escapeHtml(a.label)}</option>`;
  }).join('');

  // Populate category select
  const catSel = document.getElementById('ed-category');
  if (tipo === 'despesa') {
    catSel.innerHTML = Object.keys(getDespesaCats()).map(c => `<option value="${c}">${c}</option>`).join('');
  } else {
    const catList = tipo === 'investimento' ? getInvestCats() : (getReceitaCats() || []);
    catSel.innerHTML = catList.map(c => `<option value="${c}">${c}</option>`).join('');
  }
  onEditCatChange();
}

function onEditCatChange() {
  const tipo = document.getElementById('ed-type').value;
  const cat = document.getElementById('ed-category').value;
  const subSel = document.getElementById('ed-subcategory');
  const despCats = getDespesaCats();
  if (tipo === 'despesa' && despCats[cat] && despCats[cat].length) {
    subSel.innerHTML = '<option value="">Sem subcategoria</option>' +
      despCats[cat].map(s => `<option value="${s}">${s}</option>`).join('');
    subSel.style.display = 'block';
  } else {
    subSel.innerHTML = '<option value="">Sem subcategoria</option>';
    subSel.style.display = tipo === 'despesa' ? 'block' : 'none';
  }
}

function setEditFormaPgto(tipo) {
  editFormaPgto = tipo;
  const isD = tipo === 'debito';
  document.getElementById('ed-btn-debito').style.cssText  = `padding:10px;border-radius:10px;border:2px solid ${isD?'#4f46e5':'var(--border)'};background:${isD?'var(--tint-indigo)':'var(--surface)'};color:${isD?'#4f46e5':'var(--text-3)'};font-weight:700;font-size:13px;cursor:pointer;`;
  document.getElementById('ed-btn-credito').style.cssText = `padding:10px;border-radius:10px;border:2px solid ${!isD?'#4f46e5':'var(--border)'};background:${!isD?'var(--tint-indigo)':'var(--surface)'};color:${!isD?'#4f46e5':'var(--text-3)'};font-weight:700;font-size:13px;cursor:pointer;`;
  onEditTypeChange();
}

let editCustoTipo = 'variavel';
function setEditCustoTipo(tipo) {
  editCustoTipo = tipo;
  const isF = tipo === 'fixo';
  document.getElementById('ed-btn-custo-fixo').style.cssText = `padding:10px;border-radius:10px;border:2px solid ${isF?'#4f46e5':'var(--border)'};background:${isF?'var(--tint-indigo)':'var(--surface)'};color:${isF?'#4f46e5':'var(--text-3)'};font-weight:700;font-size:13px;cursor:pointer;`;
  document.getElementById('ed-btn-custo-variavel').style.cssText = `padding:10px;border-radius:10px;border:2px solid ${!isF?'#4f46e5':'var(--border)'};background:${!isF?'var(--tint-indigo)':'var(--surface)'};color:${!isF?'#4f46e5':'var(--text-3)'};font-weight:700;font-size:13px;cursor:pointer;`;
}

function salvarEdicao() {
  const idx = S.transactions.findIndex(t => t.id === editTxId);
  if (idx < 0) return;
  const tipo = document.getElementById('ed-type').value;
  const isDespesa = tipo === 'despesa';

  S.transactions[idx] = {
    ...S.transactions[idx],
    type:       tipo,
    desc:       document.getElementById('ed-desc').value.trim() || S.transactions[idx].desc,
    amount:     parseFloat(document.getElementById('ed-amount').value) || S.transactions[idx].amount,
    date:       document.getElementById('ed-date').value || S.transactions[idx].date,
    accountId:  document.getElementById('ed-conta').value || S.transactions[idx].accountId,
    user:       document.getElementById('ed-user').value,
    category:   document.getElementById('ed-category').value,
    subcategory: isDespesa ? (document.getElementById('ed-subcategory').value || '') : '',
    notes:      document.getElementById('ed-notes').value.trim(),
    pago:       isDespesa ? document.getElementById('ed-pago').checked : true,
    recorrente: isDespesa ? document.getElementById('ed-recorrente').checked : false,
    compartilhada: isDespesa ? document.getElementById('ed-compartilhada').checked : false,
    formaPgto:  isDespesa ? editFormaPgto : null,
    custoTipo:  isDespesa ? editCustoTipo : null,
    updatedAt:  new Date().toISOString(),
  };

  // Compra de crédito avulsa (não parcela): se a data/cartão mudou, recalcula a fatura.
  // Parcelas têm faturaRef manual por mês — não mexer.
  const et = S.transactions[idx];
  if (et.formaPgto === 'credito' && !et.parcela && !et.parcelaTotal) {
    const card = S.accounts.find(a => a.id === et.accountId && a.accountType === 'cartao');
    if (card && card.fecha && et.date) {
      const fechaDia = parseInt(card.fecha) || 1;
      const [ey, em, eday] = et.date.split('-').map(Number);
      const eym = ey*12 + (em-1) + (eday > fechaDia ? 1 : 0);
      et.faturaRef = `${Math.floor(eym/12)}-${String(eym%12+1).padStart(2,'0')}`;
    }
  }

  save();
  fecharEditModal();
  renderHistorico();
  toast('✅ Transação atualizada!');
}
