// FinançasCasal — 11-dividas.js
// Módulo carregado em ordem por index.html (escopo global compartilhado). Não use import/export.

// ─── DÍVIDAS ──────────────────────────────────────────────────────────────────
const STATUS_DIVIDA = {
  em_dia:      { label:'Em dia',      bg:'var(--tint-green)', color:'var(--on-green)' },
  atrasada:    { label:'Atrasada',    bg:'var(--tint-rose)', color:'var(--on-rose)' },
  negociando:  { label:'Negociando',  bg:'var(--tint-amber)', color:'var(--on-amber)' },
  paga:        { label:'Paga',        bg:'var(--tint-blue)', color:'var(--on-blue)' }
};

function renderHistPagamentos(d) {
  const pgtos = (d.pagamentos || []).slice();
  const amorts = (d.amortizacoesExtra || []).map(a => ({ ...a, tipo: 'amort' }));
  const todos = [...pgtos.map(p => ({ ...p, tipo: 'parcela' })), ...amorts]
    .sort((a, b) => (b.date || '').localeCompare(a.date || ''));
  if (!todos.length) return '';

  const totalPago = todos.reduce((s, p) => s + (p.amount || 0), 0);

  const rows = todos.map(p => {
    const dt = p.date ? p.date.split('-') : [];
    const dataFmt = dt.length === 3 ? `${dt[2]}/${dt[1]}/${dt[0]}` : '-';
    const isParcela = p.tipo === 'parcela';
    const badge = isParcela
      ? `<span style="font-size:10px;font-weight:700;padding:2px 8px;border-radius:10px;background:var(--tint-green);color:var(--on-green);">Parcela ${p.parcela || ''}</span>`
      : `<span style="font-size:10px;font-weight:700;padding:2px 8px;border-radius:10px;background:var(--tint-indigo);color:#4f46e5;">Amortização</span>`;
    const titular = p.titular || d.titular || '-';
    const banco = p.banco || '-';
    return `<div style="display:flex;align-items:center;gap:10px;padding:8px 0;border-bottom:1px solid var(--surface-2);font-size:12px;">
      <span style="color:var(--text-3);min-width:72px;">${dataFmt}</span>
      ${badge}
      <span style="flex:1;color:var(--text-2);font-weight:600;">${titular}</span>
      <span style="color:var(--muted);">${banco}</span>
      <span style="font-weight:700;color:#059669;min-width:90px;text-align:right;">${brl(p.amount)}</span>
    </div>`;
  }).join('');

  return `<div style="margin-bottom:14px;border:1px solid var(--surface-2);border-radius:10px;overflow:hidden;">
    <div onclick="this.nextElementSibling.style.display=this.nextElementSibling.style.display==='none'?'block':'none';this.querySelector('span:last-child').textContent=this.nextElementSibling.style.display==='none'?'▾':'▴'" style="display:flex;justify-content:space-between;align-items:center;padding:10px 14px;background:var(--bg);cursor:pointer;">
      <span style="font-size:12px;font-weight:700;color:var(--text-2);">Histórico de Pagamentos (${todos.length})</span>
      <div style="display:flex;align-items:center;gap:10px;">
        <span style="font-size:12px;font-weight:700;color:#059669;">Total: ${brl(totalPago)}</span>
        <span style="font-size:12px;color:var(--muted);">▾</span>
      </div>
    </div>
    <div style="display:none;padding:6px 14px;max-height:200px;overflow-y:auto;">
      ${rows}
    </div>
  </div>`;
}

function renderDividas() {
  S.debts = S.debts || [];
  const search = (document.getElementById('div-search')?.value || '').toLowerCase();
  const statusFilter = document.getElementById('div-status-filter')?.value || '';

  let debts = S.debts.filter(d => {
    if (search && !d.desc.toLowerCase().includes(search) && !(BANKS[d.bank]?.name||'').toLowerCase().includes(search)) return false;
    if (statusFilter && d.status !== statusFilter) return false;
    return true;
  });

  // Summary (usa total com juros — é o que realmente será pago)
  const allDebts = S.debts;
  const totalDividas = allDebts.reduce((s,d) => s + calcTotalComJuros(d), 0);
  const totalRestante = allDebts.reduce((s,d) => {
    const hasSist = d.sistema === 'price' || d.sistema === 'sac';
    return s + (hasSist ? calcSaldoDevedor(d) : Math.max(0, (d.totalAmount||0) - (d.parcelasPagas||0) * (d.parcelaAmount||0)));
  }, 0);
  const vencidas = allDebts.filter(d => d.status === 'atrasada').length;
  const pgtoMensal = allDebts.filter(d => d.status !== 'paga').reduce((s,d) => s + (d.parcelaAmount||0), 0);

  document.getElementById('dividas-summary').innerHTML = `
    <div class="card" style="padding:20px;">
      <div style="display:flex;align-items:center;gap:10px;margin-bottom:8px;">
        <div style="width:38px;height:38px;border-radius:10px;background:var(--tint-rose);display:flex;align-items:center;justify-content:center;">
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#e11d48" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><rect x="1" y="4" width="22" height="16" rx="2"/><line x1="1" y1="10" x2="23" y2="10"/></svg>
        </div>
      </div>
      <p style="font-size:12px;color:var(--muted);margin-bottom:4px;">Total em Dívidas</p>
      <p style="font-size:22px;font-weight:800;color:var(--text);">${brl(totalDividas)}</p>
    </div>
    <div class="card" style="padding:20px;">
      <div style="display:flex;align-items:center;gap:10px;margin-bottom:8px;">
        <div style="width:38px;height:38px;border-radius:10px;background:var(--tint-amber);display:flex;align-items:center;justify-content:center;">
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#ca8a04" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>
        </div>
      </div>
      <p style="font-size:12px;color:var(--muted);margin-bottom:4px;">Restante a Pagar</p>
      <p style="font-size:22px;font-weight:800;color:#ca8a04;">${brl(totalRestante)}</p>
    </div>
    <div class="card" style="padding:20px;">
      <div style="display:flex;align-items:center;gap:10px;margin-bottom:8px;">
        <div style="width:38px;height:38px;border-radius:10px;background:var(--tint-rose);display:flex;align-items:center;justify-content:center;">
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#e11d48" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>
        </div>
      </div>
      <p style="font-size:12px;color:var(--muted);margin-bottom:4px;">Dívidas Vencidas</p>
      <p style="font-size:22px;font-weight:800;color:#e11d48;">${vencidas}</p>
    </div>
    <div class="card" style="padding:20px;">
      <div style="display:flex;align-items:center;gap:10px;margin-bottom:8px;">
        <div style="width:38px;height:38px;border-radius:10px;background:var(--tint-green);display:flex;align-items:center;justify-content:center;">
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#059669" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="1" x2="12" y2="23"/><path d="M17 5H9.5a3.5 3.5 0 000 7h5a3.5 3.5 0 010 7H6"/></svg>
        </div>
      </div>
      <p style="font-size:12px;color:var(--muted);margin-bottom:4px;">Pagamentos Mensais</p>
      <p style="font-size:22px;font-weight:800;color:#059669;">${brl(pgtoMensal)}</p>
    </div>
  `;

  const list = document.getElementById('dividas-list');
  const empty = document.getElementById('dividas-empty');

  if (!debts.length) {
    list.innerHTML = '';
    empty.style.display = 'block';
    return;
  }
  empty.style.display = 'none';

  list.innerHTML = debts.map(d => {
    const bank = BANKS[d.bank] || { name: d.bank, color:'var(--text-3)', svg:'' };
    const st = STATUS_DIVIDA[d.status] || STATUS_DIVIDA.em_dia;
    const hasSistema = d.sistema === 'price' || d.sistema === 'sac';
    const totalPgtos = ((d.pagamentos || []).reduce((s, p) => s + (p.amount || 0), 0)) + ((d.amortizacoesExtra || []).reduce((s, a) => s + (a.amount || 0), 0));
    const totalRef = calcTotalComJuros(d) || d.totalAmount || 0;
    const pctPago = (d.status === 'paga') ? 100 : (totalRef > 0 ? Math.min(100, Math.round((totalPgtos / totalRef) * 100)) : 0);
    const saldoDev = hasSistema ? calcSaldoDevedor(d) : Math.max(0, (d.totalAmount||0) - (d.parcelasPagas||0) * (d.parcelaAmount||0));
    const restante = saldoDev;
    const hoje = new Date();
    const vencDate = new Date(hoje.getFullYear(), hoje.getMonth(), d.vencimentoDia || 1);
    const vencStr = `${String(d.vencimentoDia||1).padStart(2,'0')}/${String(hoje.getMonth()+1).padStart(2,'0')}/${hoje.getFullYear()}`;

    return `<div class="card" style="padding:20px;">
      <div style="display:flex;align-items:center;gap:14px;margin-bottom:14px;">
        <div style="width:44px;height:44px;border-radius:12px;overflow:hidden;flex-shrink:0;">
          <svg width="44" height="44" viewBox="0 0 48 48">${bank.svg}</svg>
        </div>
        <div style="flex:1;min-width:0;">
          <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;">
            <span style="font-size:15px;font-weight:700;color:var(--text);">${bank.name}</span>
            <span style="font-size:11px;font-weight:600;padding:2px 10px;border-radius:20px;background:${st.bg};color:${st.color};">${st.label}</span>
          </div>
          <p style="font-size:13px;color:var(--text-3);margin-top:2px;">${escapeHtml(d.desc)}</p>
        </div>
        <div style="text-align:right;">
          <span style="font-size:12px;color:var(--muted);">Parcela ${d.parcelasPagas||0}/${d.totalParcelas||0}</span>
        </div>
      </div>
      <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:14px;">
        <div>
          <p style="font-size:11px;color:var(--muted);">Total c/ Juros</p>
          <p style="font-size:14px;font-weight:700;color:var(--text);">${brl(calcTotalComJuros(d))}</p>
          ${hasSistema ? `<p style="font-size:10px;color:var(--muted);margin-top:2px;">Principal: ${brl(d.principal||d.totalAmount)}</p>` : ''}
        </div>
        <div>
          <p style="font-size:11px;color:var(--muted);">Saldo Devedor</p>
          <p style="font-size:14px;font-weight:700;color:#e11d48;">${brl(restante)}</p>
          ${hasSistema ? `<p style="font-size:10px;color:var(--muted);margin-top:2px;">Juros: ${brl(calcTotalJuros(d))}</p>` : ''}
        </div>
        <div>
          <p style="font-size:11px;color:var(--muted);">${hasSistema ? 'Taxa / Sistema' : 'Vencimento'}</p>
          <p style="font-size:14px;font-weight:700;color:var(--text-2);">${hasSistema ? (d.taxaMensal||0)+'% · '+(d.sistema||'').toUpperCase() : 'Dia '+(d.vencimentoDia||'-')}</p>
        </div>
        <div>
          <p style="font-size:11px;color:var(--muted);">Parcela</p>
          <p style="font-size:14px;font-weight:700;color:var(--text-2);">${brl(d.parcelaAmount)}${hasSistema ? ' (dia '+(d.vencimentoDia||'-')+')' : ''}</p>
        </div>
      </div>
      <div style="margin-bottom:14px;">
        <div style="display:flex;justify-content:space-between;margin-bottom:4px;">
          <span style="font-size:12px;font-weight:600;color:var(--text-3);">Progresso</span>
          <span style="font-size:12px;font-weight:700;color:#4f46e5;">${pctPago}%</span>
        </div>
        <div style="height:8px;background:var(--surface-2);border-radius:4px;overflow:hidden;">
          <div style="height:100%;background:${pctPago>=100?'#059669':'#4f46e5'};border-radius:4px;width:${Math.min(pctPago,100)}%;transition:width 0.3s;"></div>
        </div>
      </div>
      ${renderHistPagamentos(d)}
      <div style="display:flex;gap:8px;flex-wrap:wrap;">
        ${d.status !== 'paga' ? `<button onclick="abrirEscolhaPgto('${d.id}')" style="padding:8px 16px;background:#059669;color:white;border:none;border-radius:8px;font-size:12px;font-weight:600;cursor:pointer;">Pagar</button>` : ''}
        <button onclick="abrirTabelaAmort('${d.id}')" style="padding:8px 16px;background:var(--tint-blue);color:#2563eb;border:none;border-radius:8px;font-size:12px;font-weight:600;cursor:pointer;">Tabela</button>
        ${d.status !== 'paga' ? `<button onclick="abrirSimulAmort('${d.id}')" style="padding:8px 16px;background:var(--tint-amber);color:var(--on-amber);border:none;border-radius:8px;font-size:12px;font-weight:600;cursor:pointer;">Simular Amortização</button>` : ''}
        <button onclick="editarDivida('${d.id}')" style="padding:8px 16px;background:var(--tint-indigo);color:#4f46e5;border:none;border-radius:8px;font-size:12px;font-weight:600;cursor:pointer;">Editar</button>
        <button onclick="deletarDivida('${d.id}')" style="padding:8px 16px;background:var(--tint-rose);color:#e11d48;border:none;border-radius:8px;font-size:12px;font-weight:600;cursor:pointer;">Excluir</button>
      </div>
    </div>`;
  }).join('');
}

function abrirNovaDividaModal(editId) {
  const modal = document.getElementById('divida-modal');
  modal.style.display = 'block';
  document.getElementById('div-edit-id').value = editId || '';
  document.getElementById('divida-modal-title').textContent = editId ? 'Editar Dívida' : 'Nova Dívida';

  // Populate bank select
  const bankSel = document.getElementById('div-banco');
  bankSel.innerHTML = Object.entries(BANKS).map(([k,v]) => `<option value="${k}">${v.name}</option>`).join('');

  // Populate titular select
  const titSel = document.getElementById('div-titular');
  titSel.innerHTML = `<option value="${S.settings.u1}">${S.settings.u1}</option><option value="${S.settings.u2}">${S.settings.u2}</option>`;

  if (editId) {
    const d = S.debts.find(x => x.id === editId);
    if (d) {
      bankSel.value = d.bank;
      document.getElementById('div-desc').value = d.desc;
      document.getElementById('div-sistema').value = d.sistema || 'simples';
      document.getElementById('div-principal').value = d.principal || d.totalAmount || '';
      document.getElementById('div-total').value = d.totalAmount || '';
      document.getElementById('div-parcelas').value = d.totalParcelas;
      document.getElementById('div-parcela-val').value = d.parcelaAmount;
      document.getElementById('div-vencimento').value = d.vencimentoDia;
      document.getElementById('div-status').value = d.status;
      document.getElementById('div-titular').value = d.titular;
      document.getElementById('div-pagas').value = d.parcelasPagas || 0;
      document.getElementById('div-taxa-mensal').value = d.taxaMensal || '';
      document.getElementById('div-taxa-inadimpl').value = d.taxaInadimpl || '';
      document.getElementById('div-taxa-efetiva-anual').value = d.taxaEfetivaAnual || '';
      document.getElementById('div-taxa-anual').value = d.taxaAnual || '';
      document.getElementById('div-iof-basico').value = d.iofBasico || '';
      document.getElementById('div-iof-adc').value = d.iofAdc || '';
      document.getElementById('div-data-inicio').value = d.dataInicio || '';
      togglePriceFields();
    }
  } else {
    document.getElementById('div-desc').value = '';
    document.getElementById('div-sistema').value = 'simples';
    document.getElementById('div-principal').value = '';
    document.getElementById('div-total').value = '';
    document.getElementById('div-total').dataset.auto = '0';
    document.getElementById('div-parcelas').value = '';
    document.getElementById('div-parcela-val').value = '';
    document.getElementById('div-vencimento').value = '';
    document.getElementById('div-status').value = 'em_dia';
    document.getElementById('div-pagas').value = '0';
    document.getElementById('div-taxa-mensal').value = '';
    document.getElementById('div-taxa-inadimpl').value = '';
    document.getElementById('div-taxa-efetiva-anual').value = '';
    document.getElementById('div-taxa-anual').value = '';
    document.getElementById('div-iof-basico').value = '';
    document.getElementById('div-iof-adc').value = '';
    document.getElementById('div-data-inicio').value = '';
    togglePriceFields();
  }
}

function fecharDividaModal() {
  document.getElementById('divida-modal').style.display = 'none';
}

function togglePriceFields() {
  const sistema = document.getElementById('div-sistema').value;
  document.getElementById('div-price-fields').style.display = (sistema === 'price' || sistema === 'sac') ? 'block' : 'none';
  calcParcelaDivida();
}

function calcPMT(pv, i, n) {
  if (i === 0) return pv / n;
  return pv * (i * Math.pow(1 + i, n)) / (Math.pow(1 + i, n) - 1);
}

function gerarTabelaAmortizacao(principal, taxaMensal, nParcelas, sistema, dataInicio, parcelasPagas) {
  const taxa = taxaMensal / 100;
  const tabela = [];
  let saldo = principal;
  const dt = dataInicio ? new Date(dataInicio + 'T12:00:00') : new Date();
  let totalJuros = 0, totalAmort = 0, totalPago = 0;

  if (sistema === 'price') {
    const pmt = calcPMT(principal, taxa, nParcelas);
    for (let i = 1; i <= nParcelas; i++) {
      const juros = saldo * taxa;
      const amort = pmt - juros;
      saldo = Math.max(0, saldo - amort);
      const venc = addMesesSeguro(dt, i - 1);
      totalJuros += juros;
      totalAmort += amort;
      totalPago += pmt;
      tabela.push({
        n: i, vencimento: venc, parcela: pmt, juros: juros,
        amortizacao: amort, saldo: saldo,
        pago: i <= (parcelasPagas || 0)
      });
    }
  } else if (sistema === 'sac') {
    const amortFixa = principal / nParcelas;
    for (let i = 1; i <= nParcelas; i++) {
      const juros = saldo * taxa;
      const pmt = amortFixa + juros;
      saldo = Math.max(0, saldo - amortFixa);
      const venc = addMesesSeguro(dt, i - 1);
      totalJuros += juros;
      totalAmort += amortFixa;
      totalPago += pmt;
      tabela.push({
        n: i, vencimento: venc, parcela: pmt, juros: juros,
        amortizacao: amortFixa, saldo: saldo,
        pago: i <= (parcelasPagas || 0)
      });
    }
  } else {
    const pmt = principal / nParcelas;
    for (let i = 1; i <= nParcelas; i++) {
      saldo = Math.max(0, saldo - pmt);
      const venc = addMesesSeguro(dt, i - 1);
      totalPago += pmt;
      totalAmort += pmt;
      tabela.push({
        n: i, vencimento: venc, parcela: pmt, juros: 0,
        amortizacao: pmt, saldo: saldo,
        pago: i <= (parcelasPagas || 0)
      });
    }
  }

  return { tabela, totalJuros, totalAmort, totalPago };
}

function calcParcelaDivida() {
  const sistema = document.getElementById('div-sistema').value;
  const principal = parseFloat(document.getElementById('div-principal').value) || 0;
  const parcelas = parseInt(document.getElementById('div-parcelas').value) || 0;
  const taxaMensal = parseFloat(document.getElementById('div-taxa-mensal')?.value) || 0;

  if (principal > 0 && parcelas > 0) {
    if ((sistema === 'price' || sistema === 'sac') && taxaMensal > 0) {
      const taxa = taxaMensal / 100;
      const pmt = calcPMT(principal, taxa, parcelas);
      document.getElementById('div-parcela-val').value = pmt.toFixed(2);
      const totalComJuros = pmt * parcelas;
      if (!document.getElementById('div-total').value || document.getElementById('div-total').dataset.auto === '1') {
        document.getElementById('div-total').value = totalComJuros.toFixed(2);
        document.getElementById('div-total').dataset.auto = '1';
      }
    } else {
      document.getElementById('div-parcela-val').value = (principal / parcelas).toFixed(2);
    }
  }
}

function salvarDivida() {
  const banco = document.getElementById('div-banco').value;
  const desc = document.getElementById('div-desc').value.trim();
  const sistema = document.getElementById('div-sistema').value;
  const principal = parseFloat(document.getElementById('div-principal').value) || 0;
  const totalAmount = parseFloat(document.getElementById('div-total').value) || 0;
  const totalParcelas = parseInt(document.getElementById('div-parcelas').value) || 0;
  let parcelaAmount = parseFloat(document.getElementById('div-parcela-val').value) || 0;
  const vencimentoDia = parseInt(document.getElementById('div-vencimento').value) || 1;
  const status = document.getElementById('div-status').value;
  const titular = document.getElementById('div-titular').value;
  const parcelasPagas = parseInt(document.getElementById('div-pagas').value) || 0;
  const editId = document.getElementById('div-edit-id').value;
  const taxaMensal = parseFloat(document.getElementById('div-taxa-mensal')?.value) || 0;
  const taxaInadimpl = parseFloat(document.getElementById('div-taxa-inadimpl')?.value) || 0;
  const taxaEfetivaAnual = parseFloat(document.getElementById('div-taxa-efetiva-anual')?.value) || 0;
  const taxaAnual = parseFloat(document.getElementById('div-taxa-anual')?.value) || 0;
  const iofBasico = parseFloat(document.getElementById('div-iof-basico')?.value) || 0;
  const iofAdc = parseFloat(document.getElementById('div-iof-adc')?.value) || 0;
  const dataInicio = document.getElementById('div-data-inicio')?.value || '';

  if (!desc || principal <= 0 || totalParcelas <= 0) {
    toast('Preencha todos os campos obrigatórios');
    return;
  }
  if (parcelaAmount <= 0) {
    if (sistema === 'price' && taxaMensal > 0) {
      parcelaAmount = calcPMT(principal, taxaMensal / 100, totalParcelas);
    } else {
      parcelaAmount = principal / totalParcelas;
    }
  }

  const debtData = {
    bank: banco, desc, sistema, principal, totalAmount: totalAmount || (parcelaAmount * totalParcelas),
    parcelaAmount, totalParcelas, parcelasPagas, vencimentoDia, status, titular,
    taxaMensal, taxaInadimpl, taxaEfetivaAnual, taxaAnual,
    iofBasico, iofAdc, dataInicio,
    amortizacoesExtra: []
  };

  S.debts = S.debts || [];

  if (editId) {
    const d = S.debts.find(x => x.id === editId);
    if (d) {
      Object.assign(d, debtData);
      d.updatedAt = new Date().toISOString();
    }
  } else {
    S.debts.push({
      id: 'div_' + Date.now(),
      ...debtData,
      pagamentos: [],
      amortizacoesExtra: [],
      at: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    });
  }

  save();
  fecharDividaModal();
  renderDividas();
  toast(editId ? 'Dívida atualizada!' : 'Dívida adicionada!');
}

function editarDivida(id) {
  abrirNovaDividaModal(id);
}

function deletarDivida(id) {
  if (!confirm('Tem certeza que deseja excluir esta dívida?')) return;
  S.deletedIds.push({ id, collection: 'debts', deletedAt: new Date().toISOString() });
  S.debts = S.debts.filter(d => d.id !== id);
  save();
  renderDividas();
  toast('Dívida excluída!');
}

function pagarParcelaDivida(id) {
  const d = S.debts.find(x => x.id === id);
  if (!d) return;

  const modal = document.getElementById('pagar-parcela-modal');
  modal.style.display = 'block';
  document.getElementById('pp-divida-id').value = id;

  const proxParcela = (d.parcelasPagas || 0) + 1;
  document.getElementById('pp-parcela-info').textContent = `Parcela ${proxParcela}/${d.totalParcelas} - ${brl(d.parcelaAmount)}`;
  document.getElementById('pp-valor').value = d.parcelaAmount;
  document.getElementById('pp-data').value = new Date().toISOString().slice(0,10);

  // Populate titular select and default to debt's titular
  popularTitularPgto('pp', d.titular);
}

// Shared: populate titular select for payment modals
function popularTitularPgto(prefix, defaultTitular) {
  const sel = document.getElementById(prefix + '-titular');
  sel.innerHTML = `<option value="${S.settings.u1}">${S.settings.u1}</option><option value="${S.settings.u2}">${S.settings.u2}</option>`;
  if (defaultTitular) sel.value = defaultTitular;
  filtrarContasPgto(prefix);
}

// Shared: filter accounts by selected titular
function filtrarContasPgto(prefix) {
  const titular = document.getElementById(prefix + '-titular').value;
  const contaSel = document.getElementById(prefix + '-conta');
  const contas = S.accounts.filter(a => a.accountType === 'conta' && a.owner === titular);
  contaSel.innerHTML = contas.length
    ? contas.map(a => `<option value="${a.id}">${BANKS[a.bank]?.name || a.bank} — ${a.label}</option>`).join('')
    : `<option value="">Nenhuma conta para ${titular}</option>`;
}

function fecharPagarParcelaModal() {
  document.getElementById('pagar-parcela-modal').style.display = 'none';
}

// ─── ESCOLHA TIPO DE PAGAMENTO ─────────────────────────────────────────────
function abrirEscolhaPgto(id) {
  const d = S.debts.find(x => x.id === id);
  if (!d) return;
  document.getElementById('escolha-pgto-divida-id').value = id;
  const proxParcela = (d.parcelasPagas || 0) + 1;
  const saldo = (d.sistema === 'price' || d.sistema === 'sac') ? calcSaldoDevedor(d) : Math.max(0, (d.totalAmount||0) - (d.parcelasPagas||0) * (d.parcelaAmount||0));
  document.getElementById('escolha-pgto-info').innerHTML = `<b>${escapeHtml(d.desc)}</b> — Parcela ${proxParcela}/${d.totalParcelas} | Saldo: ${brl(saldo)}`;
  document.getElementById('escolha-pgto-modal').style.display = 'block';
}

function fecharEscolhaPgtoModal() {
  document.getElementById('escolha-pgto-modal').style.display = 'none';
}

function escolherPagamentoNormal() {
  const id = document.getElementById('escolha-pgto-divida-id').value;
  fecharEscolhaPgtoModal();
  pagarParcelaDivida(id);
}

function escolherPagamentoAmortizacao() {
  const id = document.getElementById('escolha-pgto-divida-id').value;
  fecharEscolhaPgtoModal();
  abrirPagarAmortModal(id);
}

// ─── PAGAR AMORTIZAÇÃO EXTRA ────────────────────────────────────────────────
function abrirPagarAmortModal(id) {
  const d = S.debts.find(x => x.id === id);
  if (!d) return;

  const modal = document.getElementById('pagar-amort-modal');
  modal.style.display = 'block';
  document.getElementById('pa-divida-id').value = id;

  const saldo = (d.sistema === 'price' || d.sistema === 'sac') ? calcSaldoDevedor(d) : Math.max(0, (d.totalAmount||0) - (d.parcelasPagas||0) * (d.parcelaAmount||0));
  document.getElementById('pa-saldo-info').textContent = brl(saldo);
  document.getElementById('pa-valor').value = '';
  document.getElementById('pa-data').value = new Date().toISOString().slice(0,10);
  document.getElementById('pa-modo').value = 'prazo';
  setPaAmortModo('prazo');

  // Populate titular select and filter accounts
  popularTitularPgto('pa', d.titular);
}

function fecharPagarAmortModal() {
  document.getElementById('pagar-amort-modal').style.display = 'none';
}

function setPaAmortModo(modo) {
  document.getElementById('pa-modo').value = modo;
  document.getElementById('pa-modo-prazo').className = 'type-btn' + (modo === 'prazo' ? ' active-invest' : '');
  document.getElementById('pa-modo-parcela').className = 'type-btn' + (modo === 'parcela' ? ' active-invest' : '');
}

function confirmarPagarAmortizacao() {
  const dividaId = document.getElementById('pa-divida-id').value;
  const valor = parseFloat(document.getElementById('pa-valor').value) || 0;
  const data = document.getElementById('pa-data').value;
  const contaId = document.getElementById('pa-conta').value;
  const titularPgto = document.getElementById('pa-titular').value;
  const modo = document.getElementById('pa-modo').value;

  if (valor <= 0 || !data || !contaId) {
    toast('Preencha todos os campos');
    return;
  }

  const d = S.debts.find(x => x.id === dividaId);
  if (!d) return;

  const saldoAtual = (d.sistema === 'price' || d.sistema === 'sac') ? calcSaldoDevedor(d) : Math.max(0, (d.totalAmount||0) - (d.parcelasPagas||0) * (d.parcelaAmount||0));

  if (valor > saldoAtual) {
    toast('Valor maior que o saldo devedor (' + brl(saldoAtual) + ')');
    return;
  }

  // Record amortization
  d.amortizacoesExtra = d.amortizacoesExtra || [];
  const contaAmort = S.accounts.find(a => a.id === contaId);
  d.amortizacoesExtra.push({ date: data, amount: valor, modo: modo, titular: titularPgto, contaId: contaId, banco: BANKS[contaAmort?.bank]?.name || contaAmort?.bank || '' });
  d.updatedAt = new Date().toISOString();

  const novoSaldo = saldoAtual - valor;

  if (novoSaldo <= 0.01) {
    // Quitação total
    d.status = 'paga';
    d.parcelasPagas = d.totalParcelas;
  } else {
    const sistema = d.sistema || 'simples';
    const taxa = (d.taxaMensal || 0) / 100;
    const parcelasRestantes = d.totalParcelas - (d.parcelasPagas || 0);

    if (modo === 'prazo') {
      // Reduce number of installments, keep same payment amount
      if (sistema === 'price' && taxa > 0) {
        const pmt = d.parcelaAmount;
        // n = -log(1 - saldo*i/pmt) / log(1+i)
        const novasParcelas = Math.ceil(-Math.log(1 - novoSaldo * taxa / pmt) / Math.log(1 + taxa));
        d.totalParcelas = (d.parcelasPagas || 0) + Math.max(1, novasParcelas);
      } else if (sistema === 'sac' && taxa > 0) {
        const amortFixa = d.parcelaAmount - (saldoAtual * taxa);
        const novasParcelas = Math.ceil(novoSaldo / amortFixa);
        d.totalParcelas = (d.parcelasPagas || 0) + Math.max(1, novasParcelas);
      } else {
        const pmt = d.parcelaAmount;
        const novasParcelas = Math.ceil(novoSaldo / pmt);
        d.totalParcelas = (d.parcelasPagas || 0) + Math.max(1, novasParcelas);
      }
      // Recalculate totalAmount
      d.principal = novoSaldo;
    } else {
      // Reduce installment amount, keep same number of installments
      d.principal = novoSaldo;
      if (sistema === 'price' && taxa > 0) {
        d.parcelaAmount = calcPMT(novoSaldo, taxa, parcelasRestantes);
      } else if (sistema === 'sac' && taxa > 0) {
        const amortFixa = novoSaldo / parcelasRestantes;
        d.parcelaAmount = amortFixa + (novoSaldo * taxa);
      } else {
        d.parcelaAmount = novoSaldo / parcelasRestantes;
      }
      d.parcelaAmount = Math.round(d.parcelaAmount * 100) / 100;
    }
  }

  // Create transaction (user = who is actually paying)
  S.transactions.push({
    id: 'tx_' + Date.now(),
    type: 'despesa',
    desc: `Amortização - ${d.desc}`,
    amount: valor,
    date: data,
    category: 'Outros',
    subcategory: '',
    accountId: contaId,
    user: titularPgto,
    formaPgto: 'debito',
    pago: true,
    custoTipo: 'fixo',
    notas: `Amortização (${modo === 'prazo' ? 'reduzir prazo' : 'reduzir parcela'}): ${d.desc} (${BANKS[d.bank]?.name || d.bank})`,
    at: new Date().toISOString(),
    updatedAt: new Date().toISOString()
  });

  save();
  fecharPagarAmortModal();
  renderDividas();

  if (novoSaldo <= 0.01) {
    toast('Dívida quitada com amortização!');
  } else {
    toast(`Amortização de ${brl(valor)} registrada! Novo saldo: ${brl(novoSaldo)}`);
  }
}

function confirmarPagarParcela() {
  const dividaId = document.getElementById('pp-divida-id').value;
  const valor = parseFloat(document.getElementById('pp-valor').value) || 0;
  const data = document.getElementById('pp-data').value;
  const contaId = document.getElementById('pp-conta').value;
  const titularPgto = document.getElementById('pp-titular').value;

  if (valor <= 0 || !data || !contaId) {
    toast('Preencha todos os campos');
    return;
  }

  const d = S.debts.find(x => x.id === dividaId);
  if (!d) return;

  d.parcelasPagas = (d.parcelasPagas || 0) + 1;
  d.pagamentos = d.pagamentos || [];
  const contaPgto = S.accounts.find(a => a.id === contaId);
  d.pagamentos.push({ date: data, amount: valor, parcela: d.parcelasPagas, titular: titularPgto, contaId: contaId, banco: BANKS[contaPgto?.bank]?.name || contaPgto?.bank || '' });
  d.updatedAt = new Date().toISOString();

  if (d.parcelasPagas >= d.totalParcelas) {
    d.status = 'paga';
  }

  // Create a despesa transaction (user = who is actually paying)
  S.transactions.push({
    id: 'tx_' + Date.now(),
    type: 'despesa',
    desc: `Parcela ${d.parcelasPagas}/${d.totalParcelas} - ${d.desc}`,
    amount: valor,
    date: data,
    category: 'Outros',
    subcategory: '',
    accountId: contaId,
    user: titularPgto,
    formaPgto: 'debito',
    pago: true,
    custoTipo: 'fixo',
    notas: `Pagamento de dívida: ${d.desc} (${BANKS[d.bank]?.name || d.bank})`,
    at: new Date().toISOString(),
    updatedAt: new Date().toISOString()
  });

  save();
  fecharPagarParcelaModal();
  renderDividas();
  toast(`Parcela ${d.parcelasPagas}/${d.totalParcelas} paga com sucesso!`);
}

// ─── TABELA DE AMORTIZAÇÃO ─────────────────────────────────────────────────
function formatDateBR(d) {
  return `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`;
}

function abrirTabelaAmort(id) {
  const d = S.debts.find(x => x.id === id);
  if (!d) return;
  const modal = document.getElementById('amort-table-modal');
  modal.style.display = 'block';

  const sistema = d.sistema || 'simples';
  const principal = d.principal || d.totalAmount || 0;
  const taxaMensal = d.taxaMensal || 0;
  const nParcelas = d.totalParcelas || 0;
  const dataInicio = d.dataInicio || '';
  const parcelasPagas = d.parcelasPagas || 0;
  const iofTotal = (d.iofBasico || 0) + (d.iofAdc || 0);

  document.getElementById('amort-modal-title').textContent = `Tabela de Amortização — ${d.desc}`;

  const { tabela, totalJuros, totalAmort, totalPago } = gerarTabelaAmortizacao(principal, taxaMensal, nParcelas, sistema, dataInicio, parcelasPagas);

  // Resumo
  document.getElementById('amort-resumo').innerHTML = `
    <div style="background:var(--bg);border-radius:10px;padding:12px 16px;">
      <p style="font-size:11px;color:var(--muted);">Principal</p>
      <p style="font-size:16px;font-weight:800;color:var(--text);">${brl(principal)}</p>
    </div>
    <div style="background:var(--bg);border-radius:10px;padding:12px 16px;">
      <p style="font-size:11px;color:var(--muted);">Total Juros</p>
      <p style="font-size:16px;font-weight:800;color:#e11d48;">${brl(totalJuros)}</p>
    </div>
    <div style="background:var(--bg);border-radius:10px;padding:12px 16px;">
      <p style="font-size:11px;color:var(--muted);">IOF Total</p>
      <p style="font-size:16px;font-weight:800;color:#ca8a04;">${brl(iofTotal)}</p>
    </div>
    <div style="background:var(--bg);border-radius:10px;padding:12px 16px;">
      <p style="font-size:11px;color:var(--muted);">Total a Pagar</p>
      <p style="font-size:16px;font-weight:800;color:#4f46e5;">${brl(totalPago + iofTotal)}</p>
    </div>
    <div style="background:var(--bg);border-radius:10px;padding:12px 16px;">
      <p style="font-size:11px;color:var(--muted);">Sistema</p>
      <p style="font-size:16px;font-weight:800;color:var(--text-2);">${sistema.toUpperCase()}</p>
    </div>
    <div style="background:var(--bg);border-radius:10px;padding:12px 16px;">
      <p style="font-size:11px;color:var(--muted);">Taxa Mensal</p>
      <p style="font-size:16px;font-weight:800;color:var(--text-2);">${taxaMensal}%</p>
    </div>
  `;

  // Tabela
  const tbody = document.getElementById('amort-tbody');
  tbody.innerHTML = tabela.map(r => {
    const isPago = r.pago;
    const rowBg = isPago ? 'background:#f0fdf4;' : '';
    return `<tr style="border-bottom:1px solid var(--surface-2);${rowBg}">
      <td style="padding:8px 12px;text-align:center;font-weight:600;">${r.n}</td>
      <td style="padding:8px 12px;text-align:center;">${formatDateBR(r.vencimento)}</td>
      <td style="padding:8px 12px;text-align:right;font-weight:600;">${brl(r.parcela)}</td>
      <td style="padding:8px 12px;text-align:right;color:#e11d48;">${brl(r.juros)}</td>
      <td style="padding:8px 12px;text-align:right;color:#059669;">${brl(r.amortizacao)}</td>
      <td style="padding:8px 12px;text-align:right;font-weight:700;">${brl(r.saldo)}</td>
      <td style="padding:8px 12px;text-align:center;">
        ${isPago ? '<span style="font-size:11px;font-weight:600;padding:2px 10px;border-radius:20px;background:var(--tint-green);color:var(--on-green);">Pago</span>' :
          '<span style="font-size:11px;font-weight:600;padding:2px 10px;border-radius:20px;background:var(--surface-2);color:var(--text-3);">Pendente</span>'}
      </td>
    </tr>`;
  }).join('');

  // Totais
  document.getElementById('amort-tfoot').innerHTML = `
    <tr style="border-top:2px solid var(--border);background:var(--bg);font-weight:800;">
      <td style="padding:10px 12px;text-align:center;" colspan="2">TOTAL</td>
      <td style="padding:10px 12px;text-align:right;">${brl(totalPago)}</td>
      <td style="padding:10px 12px;text-align:right;color:#e11d48;">${brl(totalJuros)}</td>
      <td style="padding:10px 12px;text-align:right;color:#059669;">${brl(totalAmort)}</td>
      <td style="padding:10px 12px;text-align:right;">—</td>
      <td style="padding:10px 12px;text-align:center;">${parcelasPagas}/${nParcelas}</td>
    </tr>
  `;
}

function fecharAmortModal() {
  document.getElementById('amort-table-modal').style.display = 'none';
}

// ─── SIMULADOR DE AMORTIZAÇÃO EXTRAORDINÁRIA ─────────────────────────────
let simulModo = 'prazo';

function setSimulModo(modo) {
  simulModo = modo;
  document.getElementById('simul-modo-prazo').className = 'type-btn' + (modo === 'prazo' ? ' active-invest' : '');
  document.getElementById('simul-modo-parcela').className = 'type-btn' + (modo === 'parcela' ? ' active-invest' : '');
  calcSimulacao();
}

function abrirSimulAmort(id) {
  const d = S.debts.find(x => x.id === id);
  if (!d) return;
  const modal = document.getElementById('simul-amort-modal');
  modal.style.display = 'block';
  document.getElementById('simul-divida-id').value = id;
  document.getElementById('simul-valor').value = '';
  document.getElementById('simul-data').value = new Date().toISOString().slice(0, 10);
  document.getElementById('simul-resultado').style.display = 'none';
  setSimulModo('prazo');

  const sistema = (d.sistema || 'price').toUpperCase();
  const saldoAtual = calcSaldoDevedor(d);
  document.getElementById('simul-divida-info').innerHTML = `
    <div style="display:flex;justify-content:space-between;flex-wrap:wrap;gap:8px;">
      <div>
        <p style="font-size:12px;color:var(--muted);">Dívida</p>
        <p style="font-size:15px;font-weight:700;color:var(--text);">${escapeHtml(d.desc)}</p>
      </div>
      <div style="text-align:right;">
        <p style="font-size:12px;color:var(--muted);">Saldo Devedor Atual</p>
        <p style="font-size:18px;font-weight:800;color:#e11d48;">${brl(saldoAtual)}</p>
      </div>
    </div>
    <div style="display:flex;gap:16px;margin-top:10px;flex-wrap:wrap;">
      <div><span style="font-size:12px;color:var(--muted);">Sistema:</span> <strong>${sistema}</strong></div>
      <div><span style="font-size:12px;color:var(--muted);">Taxa:</span> <strong>${d.taxaMensal || 0}% a.m.</strong></div>
      <div><span style="font-size:12px;color:var(--muted);">Parcelas restantes:</span> <strong>${(d.totalParcelas||0) - (d.parcelasPagas||0)}</strong></div>
      <div><span style="font-size:12px;color:var(--muted);">Parcela atual:</span> <strong>${brl(d.parcelaAmount)}</strong></div>
    </div>
  `;
}

// Calcula o total com juros de uma dívida (quanto será pago ao todo)
function calcTotalComJuros(d) {
  const sistema = d.sistema || 'simples';
  const principal = d.principal || d.totalAmount || 0;
  const taxa = (d.taxaMensal || 0) / 100;
  const n = d.totalParcelas || 0;
  if (!taxa || taxa === 0 || !n) return d.totalAmount || principal;
  if (sistema === 'price') {
    const pmt = calcPMT(principal, taxa, n);
    return pmt * n;
  } else if (sistema === 'sac') {
    const amortFixa = principal / n;
    let total = 0, saldo = principal;
    for (let i = 0; i < n; i++) {
      total += amortFixa + saldo * taxa;
      saldo -= amortFixa;
    }
    return total;
  }
  return d.totalAmount || principal;
}

// Calcula total de juros da dívida
function calcTotalJuros(d) {
  const principal = d.principal || d.totalAmount || 0;
  return calcTotalComJuros(d) - principal;
}

function calcSaldoDevedor(d) {
  const sistema = d.sistema || 'simples';
  const principal = d.principal || d.totalAmount || 0;
  const taxa = (d.taxaMensal || 0) / 100;
  const n = d.totalParcelas || 0;
  const pagas = d.parcelasPagas || 0;

  if (sistema === 'price' && taxa > 0) {
    const pmt = calcPMT(principal, taxa, n);
    let saldo = principal;
    for (let i = 0; i < pagas; i++) {
      const juros = saldo * taxa;
      saldo = saldo - (pmt - juros);
    }
    return Math.max(0, saldo);
  } else if (sistema === 'sac' && taxa > 0) {
    const amortFixa = principal / n;
    return Math.max(0, principal - amortFixa * pagas);
  } else {
    return Math.max(0, principal - (principal / n) * pagas);
  }
}

function calcSimulacao() {
  const id = document.getElementById('simul-divida-id').value;
  const d = S.debts.find(x => x.id === id);
  if (!d) return;

  const valorAmort = parseFloat(document.getElementById('simul-valor').value) || 0;
  if (valorAmort <= 0) {
    document.getElementById('simul-resultado').style.display = 'none';
    return;
  }

  const sistema = d.sistema || 'simples';
  const principal = d.principal || d.totalAmount || 0;
  const taxa = (d.taxaMensal || 0) / 100;
  const totalParcelas = d.totalParcelas || 0;
  const parcelasPagas = d.parcelasPagas || 0;
  const parcelasRestantes = totalParcelas - parcelasPagas;
  const saldoAtual = calcSaldoDevedor(d);
  const parcelaAtual = d.parcelaAmount || 0;

  if (valorAmort >= saldoAtual) {
    document.getElementById('simul-resultado').style.display = 'block';
    document.getElementById('simul-antes').innerHTML = `
      <p style="font-size:13px;"><strong>Saldo:</strong> ${brl(saldoAtual)}</p>
      <p style="font-size:13px;"><strong>Parcelas:</strong> ${parcelasRestantes}x ${brl(parcelaAtual)}</p>
    `;
    const totalOriginal = parcelaAtual * parcelasRestantes;
    document.getElementById('simul-depois').innerHTML = `
      <p style="font-size:18px;font-weight:800;color:#059669;">QUITAÇÃO TOTAL!</p>
      <p style="font-size:13px;margin-top:4px;">Valor para quitar: ${brl(saldoAtual)}</p>
    `;
    document.getElementById('simul-economia').innerHTML = `
      <p style="font-size:14px;font-weight:700;color:#059669;">Economia total de juros: ${brl(totalOriginal - saldoAtual)}</p>
      <p style="font-size:12px;color:var(--text-3);margin-top:4px;">Troco: ${brl(valorAmort - saldoAtual)}</p>
    `;
    document.getElementById('simul-nova-tbody').innerHTML = '';
    return;
  }

  const novoSaldo = saldoAtual - valorAmort;
  let novasParcelas, novaParcela, totalAntes, totalDepois;

  if (simulModo === 'prazo') {
    // Reduzir prazo: manter valor da parcela, recalcular nº de parcelas
    novaParcela = parcelaAtual;
    if (sistema === 'price' && taxa > 0) {
      // n = -log(1 - saldo*i/PMT) / log(1+i)
      const x = 1 - (novoSaldo * taxa) / novaParcela;
      if (x <= 0) {
        novasParcelas = parcelasRestantes;
      } else {
        novasParcelas = Math.ceil(-Math.log(x) / Math.log(1 + taxa));
      }
    } else if (sistema === 'sac' && taxa > 0) {
      novasParcelas = Math.ceil(novoSaldo / (parcelaAtual - novoSaldo * taxa));
      if (novasParcelas <= 0 || !isFinite(novasParcelas)) novasParcelas = Math.ceil(novoSaldo / novaParcela);
    } else {
      novasParcelas = Math.ceil(novoSaldo / novaParcela);
    }
  } else {
    // Reduzir parcela: manter prazo, recalcular valor da parcela
    novasParcelas = parcelasRestantes;
    if (sistema === 'price' && taxa > 0) {
      novaParcela = calcPMT(novoSaldo, taxa, novasParcelas);
    } else if (sistema === 'sac' && taxa > 0) {
      const amortFixa = novoSaldo / novasParcelas;
      novaParcela = amortFixa + novoSaldo * taxa; // primeira parcela SAC
    } else {
      novaParcela = novoSaldo / novasParcelas;
    }
  }

  // Calcular totais
  if (sistema === 'price' && taxa > 0) {
    totalAntes = parcelaAtual * parcelasRestantes;
    totalDepois = (simulModo === 'prazo' ? novaParcela * novasParcelas : novaParcela * novasParcelas) + valorAmort;
  } else {
    totalAntes = parcelaAtual * parcelasRestantes;
    totalDepois = novaParcela * novasParcelas + valorAmort;
  }
  // For PRICE, calculate actual total with interest
  let totalJurosAntes = totalAntes - saldoAtual;
  let totalJurosDepois = 0;

  // Gerar nova tabela
  const dataSimul = document.getElementById('simul-data').value;
  const { tabela: novaTabela, totalJuros: jurosNovos } = gerarTabelaAmortizacao(
    novoSaldo, d.taxaMensal || 0, novasParcelas, sistema, dataSimul, 0
  );
  totalJurosDepois = jurosNovos;
  totalDepois = novoSaldo + jurosNovos + valorAmort;

  const economia = totalAntes - (novoSaldo + jurosNovos + valorAmort);

  document.getElementById('simul-resultado').style.display = 'block';

  document.getElementById('simul-antes').innerHTML = `
    <p style="font-size:13px;margin-bottom:4px;"><strong>Saldo devedor:</strong> ${brl(saldoAtual)}</p>
    <p style="font-size:13px;margin-bottom:4px;"><strong>Parcelas restantes:</strong> ${parcelasRestantes}x</p>
    <p style="font-size:13px;margin-bottom:4px;"><strong>Valor parcela:</strong> ${brl(parcelaAtual)}</p>
    <p style="font-size:13px;margin-bottom:4px;"><strong>Total juros restante:</strong> <span style="color:#e11d48;">${brl(totalJurosAntes)}</span></p>
    <p style="font-size:13px;"><strong>Total a pagar:</strong> ${brl(totalAntes)}</p>
  `;

  document.getElementById('simul-depois').innerHTML = `
    <p style="font-size:13px;margin-bottom:4px;"><strong>Novo saldo:</strong> ${brl(novoSaldo)}</p>
    <p style="font-size:13px;margin-bottom:4px;"><strong>Parcelas restantes:</strong> ${novasParcelas}x</p>
    <p style="font-size:13px;margin-bottom:4px;"><strong>Valor parcela:</strong> ${brl(novaParcela)}</p>
    <p style="font-size:13px;margin-bottom:4px;"><strong>Total juros restante:</strong> <span style="color:#e11d48;">${brl(totalJurosDepois)}</span></p>
    <p style="font-size:13px;"><strong>Total a pagar:</strong> ${brl(novoSaldo + jurosNovos + valorAmort)}</p>
  `;

  const econJuros = totalJurosAntes - totalJurosDepois;
  document.getElementById('simul-economia').innerHTML = `
    <div style="display:flex;align-items:center;gap:10px;margin-bottom:8px;">
      <span style="font-size:24px;">💰</span>
      <div>
        <p style="font-size:16px;font-weight:800;color:#059669;">Economia de juros: ${brl(econJuros)}</p>
        <p style="font-size:13px;color:var(--text-3);">
          ${simulModo === 'prazo'
            ? `Redução de ${parcelasRestantes - novasParcelas} parcelas (${parcelasRestantes} → ${novasParcelas})`
            : `Redução de ${brl(parcelaAtual - novaParcela)} por parcela (${brl(parcelaAtual)} → ${brl(novaParcela)})`
          }
        </p>
      </div>
    </div>
  `;

  // Nova tabela
  document.getElementById('simul-nova-tbody').innerHTML = novaTabela.map(r => `
    <tr style="border-bottom:1px solid var(--surface-2);">
      <td style="padding:6px 8px;text-align:center;">${r.n}</td>
      <td style="padding:6px 8px;text-align:center;">${formatDateBR(r.vencimento)}</td>
      <td style="padding:6px 8px;text-align:right;font-weight:600;">${brl(r.parcela)}</td>
      <td style="padding:6px 8px;text-align:right;color:#e11d48;">${brl(r.juros)}</td>
      <td style="padding:6px 8px;text-align:right;color:#059669;">${brl(r.amortizacao)}</td>
      <td style="padding:6px 8px;text-align:right;font-weight:600;">${brl(r.saldo)}</td>
    </tr>
  `).join('');
}

function fecharSimulModal() {
  document.getElementById('simul-amort-modal').style.display = 'none';
}
