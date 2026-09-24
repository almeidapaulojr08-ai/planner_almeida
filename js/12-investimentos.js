// FinançasCasal — 12-investimentos.js
// Módulo carregado em ordem por index.html (escopo global compartilhado). Não use import/export.

// ─── INVESTIMENTOS ────────────────────────────────────────────────────────────
const TIPO_INVEST = {
  reserva:    { label:'Reserva',           icon:'🏦', bg:'var(--tint-green)', color:'var(--on-green)' },
  fgts:       { label:'FGTS',             icon:'🏛️', bg:'var(--tint-blue)', color:'var(--on-blue)' },
  negocio:    { label:'Negócio / Retorno', icon:'🏢', bg:'var(--tint-amber)', color:'var(--on-amber)' },
  renda_fixa: { label:'Renda Fixa',       icon:'📄', bg:'#f0fdf4', color:'#166534' },
  acoes:      { label:'Ações / FIIs',     icon:'📈', bg:'var(--tint-indigo)', color:'#4f46e5' },
  cripto:     { label:'Cripto',           icon:'₿',  bg:'#fefce8', color:'#a16207' },
  previdencia:{ label:'Previdência',     icon:'🛡️', bg:'#fdf2f8', color:'#9d174d' },
  outro:      { label:'Outro',            icon:'💼', bg:'var(--surface-2)', color:'var(--text-2)' }
};

function renderInvestimentos() {
  S.investments = S.investments || [];
  const search = (document.getElementById('inv-search')?.value || '').toLowerCase();
  const tipoFilter = document.getElementById('inv-tipo-filter')?.value || '';
  const titFilter = document.getElementById('inv-titular-filter')?.value || '';

  // Populate titular filter
  const titSel = document.getElementById('inv-titular-filter');
  if (titSel && titSel.options.length <= 1) {
    titSel.innerHTML = `<option value="">Todos Titulares</option>
      <option value="${S.settings.u1}">${S.settings.u1}</option>
      <option value="${S.settings.u2}">${S.settings.u2}</option>`;
  }

  let invs = S.investments.filter(inv => {
    if (search && !inv.desc.toLowerCase().includes(search) && !(BANKS[inv.bank]?.name||'').toLowerCase().includes(search)) return false;
    if (tipoFilter && inv.tipo !== tipoFilter) return false;
    if (titFilter && inv.titular !== titFilter) return false;
    return true;
  });

  // Summary
  const all = S.investments;
  const totalInvestido = all.reduce((s,i) => s + (i.valorInvestido||0), 0);
  const totalAtual = all.reduce((s,i) => s + calcValorAtual(i), 0);
  const totalRendimentos = all.reduce((s,i) => s + calcTotalRendimentos(i), 0);
  const lucro = totalAtual - totalInvestido + totalRendimentos;

  document.getElementById('invest-summary').innerHTML = `
    <div class="card" style="padding:20px;">
      <div style="display:flex;align-items:center;gap:10px;margin-bottom:8px;">
        <div style="width:38px;height:38px;border-radius:10px;background:var(--tint-blue);display:flex;align-items:center;justify-content:center;">
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#2563eb" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="1" x2="12" y2="23"/><path d="M17 5H9.5a3.5 3.5 0 000 7h5a3.5 3.5 0 010 7H6"/></svg>
        </div>
      </div>
      <p style="font-size:12px;color:var(--muted);margin-bottom:4px;">Total Investido</p>
      <p style="font-size:22px;font-weight:800;color:var(--text);">${brl(totalInvestido)}</p>
    </div>
    <div class="card" style="padding:20px;">
      <div style="display:flex;align-items:center;gap:10px;margin-bottom:8px;">
        <div style="width:38px;height:38px;border-radius:10px;background:var(--tint-green);display:flex;align-items:center;justify-content:center;">
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#059669" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="23 6 13.5 15.5 8.5 10.5 1 18"/><polyline points="17 6 23 6 23 12"/></svg>
        </div>
      </div>
      <p style="font-size:12px;color:var(--muted);margin-bottom:4px;">Valor Atual</p>
      <p style="font-size:22px;font-weight:800;color:#059669;">${brl(totalAtual)}</p>
    </div>
    <div class="card" style="padding:20px;">
      <div style="display:flex;align-items:center;gap:10px;margin-bottom:8px;">
        <div style="width:38px;height:38px;border-radius:10px;background:var(--tint-amber);display:flex;align-items:center;justify-content:center;">
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#ca8a04" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="16"/><line x1="8" y1="12" x2="16" y2="12"/></svg>
        </div>
      </div>
      <p style="font-size:12px;color:var(--muted);margin-bottom:4px;">Total Rendimentos</p>
      <p style="font-size:22px;font-weight:800;color:#ca8a04;">${brl(totalRendimentos)}</p>
    </div>
    <div class="card" style="padding:20px;">
      <div style="display:flex;align-items:center;gap:10px;margin-bottom:8px;">
        <div style="width:38px;height:38px;border-radius:10px;background:${lucro >= 0 ? 'var(--tint-green)' : 'var(--tint-rose)'};display:flex;align-items:center;justify-content:center;">
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="${lucro >= 0 ? '#059669' : '#e11d48'}" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>
        </div>
      </div>
      <p style="font-size:12px;color:var(--muted);margin-bottom:4px;">Lucro / Prejuízo</p>
      <p style="font-size:22px;font-weight:800;color:${lucro >= 0 ? '#059669' : '#e11d48'};">${brl(lucro)}</p>
    </div>
  `;

  const list = document.getElementById('invest-list');
  const empty = document.getElementById('invest-empty');

  if (!invs.length) {
    list.innerHTML = '';
    empty.style.display = 'block';
    return;
  }
  empty.style.display = 'none';

  list.innerHTML = invs.map(inv => {
    const tp = TIPO_INVEST[inv.tipo] || TIPO_INVEST.outro;
    const bank = BANKS[inv.bank] || { name: inv.bank || '-', color:'var(--text-3)', svg:'' };
    const valorAtual = calcValorAtual(inv);
    const totalRend = calcTotalRendimentos(inv);
    const lucroInv = totalRend + valorAtual - inv.valorInvestido;
    const pctRetorno = inv.valorInvestido > 0 ? ((lucroInv / inv.valorInvestido) * 100) : 0;

    // Break-even para negócio
    let breakEvenHtml = '';
    if (inv.tipo === 'negocio' && inv.retornoMensal > 0) {
      const isUSD = inv.moedaRetorno === 'USD';
      const taxa = S.settings.usdRate || 5.0;
      const retornoMensalBRL = isUSD ? inv.retornoMensal * taxa : inv.retornoMensal;
      const mesesParaPagar = Math.ceil(inv.valorInvestido / retornoMensalBRL);
      const mesesPassados = calcMesesPassados(inv.dataInicio);
      const rendAcum = totalRend;
      const rendOriginal = (inv.rendimentos||[]).reduce((s,r) => s + (r.valor||0), 0);
      const faltam = Math.max(0, inv.valorInvestido - rendAcum);
      const mesesFaltam = retornoMensalBRL > 0 ? Math.ceil(faltam / retornoMensalBRL) : '?';
      const pctPago = inv.valorInvestido > 0 ? Math.min(100, Math.round((rendAcum / inv.valorInvestido) * 100)) : 0;
      const jaLucro = rendAcum >= inv.valorInvestido;

      breakEvenHtml = `
        <div style="margin-top:14px;padding:14px;background:var(--bg);border-radius:10px;">
          <div style="display:flex;justify-content:space-between;margin-bottom:6px;">
            <span style="font-size:12px;font-weight:700;color:var(--text-2);">Retorno do investimento</span>
            <span style="font-size:12px;font-weight:800;color:${jaLucro ? '#059669' : '#4f46e5'};">${jaLucro ? '✅ DANDO LUCRO!' : mesesFaltam + ' meses restantes'}</span>
          </div>
          <div style="height:10px;background:var(--border);border-radius:5px;overflow:hidden;margin-bottom:6px;">
            <div style="height:100%;background:${jaLucro ? '#059669' : '#2563eb'};border-radius:5px;width:${pctPago}%;transition:width 0.3s;"></div>
          </div>
          <div style="display:flex;justify-content:space-between;">
            <span style="font-size:11px;color:var(--muted);">Recuperado: ${brl(rendAcum)}${isUSD ? ' (US$ '+rendOriginal.toFixed(2)+')' : ''} de ${brl(inv.valorInvestido)}</span>
            <span style="font-size:11px;font-weight:700;color:${jaLucro ? '#059669' : '#4f46e5'};">${pctPago}%</span>
          </div>
          <div style="display:flex;gap:12px;margin-top:8px;flex-wrap:wrap;">
            <span style="font-size:11px;color:var(--text-3);">Retorno/mês: <strong>${isUSD ? 'US$ '+inv.retornoMensal.toFixed(2)+' (≈'+brl(retornoMensalBRL)+')' : brl(inv.retornoMensal)}</strong></span>
            <span style="font-size:11px;color:var(--text-3);">Break-even: <strong>${mesesParaPagar} meses</strong></span>
            <span style="font-size:11px;color:var(--text-3);">Passados: <strong>${mesesPassados} meses</strong></span>
            ${isUSD ? '<span style="font-size:11px;color:var(--muted);">Câmbio: US$ 1 = '+brl(taxa)+'</span>' : ''}
          </div>
        </div>`;
    }

    // Renda Fixa - IR info
    if (inv.tipo === 'renda_fixa') {
      const aliq = calcIRAliquota(inv.dataInicio);
      const dias = inv.dataInicio ? Math.floor((new Date() - new Date(inv.dataInicio+'T12:00:00'))/(1000*60*60*24)) : 0;
      const rendBruto = totalRend;
      const irValor = rendBruto * aliq;
      const rendLiq = rendBruto - irValor;
      const faixaLabel = aliq === 0.225 ? 'até 180d' : aliq === 0.20 ? '181-360d' : aliq === 0.175 ? '361-720d' : 'acima 720d';

      breakEvenHtml = `
        <div style="margin-top:14px;padding:14px;background:var(--bg);border-radius:10px;">
          <div style="display:flex;justify-content:space-between;margin-bottom:8px;">
            <span style="font-size:12px;font-weight:700;color:var(--text-2);">Rendimento ${inv.indice ? inv.indice.toUpperCase() : 'CDI'}</span>
            <span style="font-size:12px;font-weight:700;color:#4f46e5;">${inv.rendimentoPct ? inv.rendimentoPct+'% último mês' : ''}</span>
          </div>
          <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px;">
            <div>
              <p style="font-size:11px;color:var(--muted);">Rend. Bruto</p>
              <p style="font-size:13px;font-weight:700;color:#059669;">${brl(rendBruto)}</p>
            </div>
            <div>
              <p style="font-size:11px;color:var(--muted);">IR (${Math.round(aliq*100)}% · ${faixaLabel})</p>
              <p style="font-size:13px;font-weight:700;color:#e11d48;">-${brl(irValor)}</p>
            </div>
            <div>
              <p style="font-size:11px;color:var(--muted);">Rend. Líquido</p>
              <p style="font-size:13px;font-weight:700;color:#2563eb;">${brl(rendLiq)}</p>
            </div>
          </div>
          <p style="font-size:11px;color:var(--muted);margin-top:6px;">${dias} dias investidos · Próxima faixa: ${aliq > 0.15 ? 'menos IR em breve' : 'alíquota mínima atingida!'}</p>
        </div>`;
    }

    // Previdência details
    if (inv.tipo === 'previdencia') {
      const aportes = inv.valorInvestido;
      const atual = calcValorAtual(inv);
      const rendBruto = atual - aportes;
      const pctRend = aportes > 0 ? ((rendBruto / aportes) * 100) : 0;
      const mesesInv = calcMesesPassados(inv.dataInicio);
      const aporteMensal = inv.prevAporteMensal || 0;

      breakEvenHtml = `
        <div style="margin-top:14px;padding:14px;background:var(--bg);border-radius:10px;">
          <div style="display:flex;justify-content:space-between;margin-bottom:8px;">
            <span style="font-size:12px;font-weight:700;color:var(--text-2);">${inv.prevTipo || 'VGBL'} · Taxa Admin: ${inv.prevTaxaAdmin || 0}%/ano</span>
            <span style="font-size:12px;font-weight:700;color:${rendBruto >= 0 ? '#059669' : '#e11d48'};">${rendBruto >= 0 ? '📈' : '📉'} ${pctRend.toFixed(2)}% acumulado</span>
          </div>
          <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px;margin-bottom:8px;">
            <div>
              <p style="font-size:11px;color:var(--muted);">Total Aportado</p>
              <p style="font-size:13px;font-weight:700;color:var(--text);">${brl(aportes)}</p>
            </div>
            <div>
              <p style="font-size:11px;color:var(--muted);">Valor Atual</p>
              <p style="font-size:13px;font-weight:700;color:#2563eb;">${brl(atual)}</p>
            </div>
            <div>
              <p style="font-size:11px;color:var(--muted);">Rendimento</p>
              <p style="font-size:13px;font-weight:700;color:${rendBruto >= 0 ? '#059669' : '#e11d48'};">${rendBruto >= 0 ? '+' : ''}${brl(rendBruto)}</p>
            </div>
          </div>
          <div style="display:flex;gap:12px;flex-wrap:wrap;">
            <span style="font-size:11px;color:var(--text-3);">Aporte mensal: <strong>${brl(aporteMensal)}</strong></span>
            <span style="font-size:11px;color:var(--text-3);">Meses investidos: <strong>${mesesInv}</strong></span>
            ${inv.prevRent12m ? '<span style="font-size:11px;color:var(--text-3);">Rent. 12m: <strong>' + inv.prevRent12m + '%</strong></span>' : ''}
          </div>
        </div>`;
    }

    // Cripto details
    if (inv.tipo === 'cripto' && inv.criptoTokens) {
      const dolarAtual = S.settings.usdRate || inv.criptoPtaxCompra || 5.0;
      const precoAtual = inv.criptoPrecoAtual || 0;
      const valorAtualUSD = inv.criptoTokens * precoAtual;
      const valorAtualBRL = valorAtualUSD * dolarAtual;
      const custoTotalBRL = inv.valorInvestido;
      const lucroBRL = valorAtualBRL - custoTotalBRL;
      const lucroPct = custoTotalBRL > 0 ? ((lucroBRL / custoTotalBRL) * 100) : 0;
      const precoMedio = inv.criptoTokens > 0 ? (inv.criptoUSD / inv.criptoTokens) : 0;
      const nCompras = (inv.compras || []).length;
      const comprasLabel = nCompras > 1 ? `${nCompras} compras` : (nCompras === 1 ? '1 compra' : '');

      breakEvenHtml = `
        <div style="margin-top:14px;padding:14px;background:var(--bg);border-radius:10px;">
          <div style="display:flex;justify-content:space-between;margin-bottom:10px;">
            <span style="font-size:12px;font-weight:700;color:var(--text-2);">Detalhes Cripto${comprasLabel ? ' · ' + comprasLabel : ''}</span>
            <span style="font-size:12px;font-weight:800;color:${lucroBRL >= 0 ? '#059669' : '#e11d48'};">${lucroBRL >= 0 ? '📈 Valorizando' : '📉 Desvalorizando'}</span>
          </div>
          <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px;margin-bottom:10px;">
            <div>
              <p style="font-size:11px;color:var(--muted);">Investido</p>
              <p style="font-size:13px;font-weight:700;color:var(--text);">US$ ${inv.criptoUSD.toFixed(2)}</p>
              <p style="font-size:11px;color:var(--text-3);">PTAX médio: R$ ${(inv.criptoPtaxCompra||0).toFixed(2)}</p>
            </div>
            <div>
              <p style="font-size:11px;color:var(--muted);">Tokens</p>
              <p style="font-size:13px;font-weight:700;color:var(--text);">${inv.criptoTokens.toLocaleString('pt-BR', {maximumFractionDigits:6})}</p>
              <p style="font-size:11px;color:var(--text-3);">PM: US$ ${precoMedio.toFixed(6)}</p>
            </div>
            <div>
              <p style="font-size:11px;color:var(--muted);">Preço Atual</p>
              <p style="font-size:13px;font-weight:700;color:#2563eb;">US$ ${precoAtual.toFixed(6)}</p>
              <p style="font-size:11px;color:var(--text-3);">
                <a href="#" onclick="editarPrecoAtualCripto('${inv.id}');return false;" style="color:#2563eb;text-decoration:underline;">Atualizar</a>
              </p>
            </div>
          </div>
          <div style="border-top:1px solid var(--border);padding-top:10px;display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px;">
            <div>
              <p style="font-size:11px;color:var(--muted);">Valor Atual (USD)</p>
              <p style="font-size:13px;font-weight:700;color:var(--text);">US$ ${valorAtualUSD.toFixed(2)}</p>
            </div>
            <div>
              <p style="font-size:11px;color:var(--muted);">Valor Atual (BRL)</p>
              <p style="font-size:13px;font-weight:700;color:#2563eb;">${brl(valorAtualBRL)}</p>
            </div>
            <div>
              <p style="font-size:11px;color:var(--muted);">Lucro / Prejuízo</p>
              <p style="font-size:13px;font-weight:700;color:${lucroBRL >= 0 ? '#059669' : '#e11d48'};">${lucroBRL >= 0 ? '+' : ''}${brl(lucroBRL)} (${lucroPct >= 0 ? '+' : ''}${lucroPct.toFixed(1)}%)</p>
            </div>
          </div>
          <p style="font-size:11px;color:var(--muted);margin-top:8px;">Dólar atual: R$ ${dolarAtual.toFixed(2)} · Investido em BRL: ${brl(custoTotalBRL)}</p>
        </div>`;
    }

    // Objetivo
    let objHtml = inv.objetivo ? `<span style="font-size:11px;font-weight:600;padding:2px 10px;border-radius:20px;background:var(--tint-amber);color:var(--on-amber);">🎯 ${inv.objetivo}</span>` : '';

    // Histórico de rendimentos
    const nRend = (inv.rendimentos || []).length;

    return `<div class="card" style="padding:20px;">
      <div style="display:flex;align-items:center;gap:14px;margin-bottom:14px;">
        <div style="width:44px;height:44px;border-radius:12px;background:${tp.bg};display:flex;align-items:center;justify-content:center;font-size:22px;flex-shrink:0;">
          ${tp.icon}
        </div>
        <div style="flex:1;min-width:0;">
          <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;">
            <span style="font-size:15px;font-weight:700;color:var(--text);">${escapeHtml(inv.desc)}</span>
            <span style="font-size:11px;font-weight:600;padding:2px 10px;border-radius:20px;background:${tp.bg};color:${tp.color};">${tp.label}</span>
            ${objHtml}
          </div>
          <p style="font-size:13px;color:var(--text-3);margin-top:2px;">${bank.name}${inv.titular ? ' · ' + inv.titular : ''}</p>
        </div>
        <div style="text-align:right;">
          <span style="font-size:12px;color:var(--muted);">${nRend} rendimento${nRend !== 1 ? 's' : ''}</span>
        </div>
      </div>
      <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:14px;">
        <div>
          <p style="font-size:11px;color:var(--muted);">Investido</p>
          <p style="font-size:14px;font-weight:700;color:var(--text);">${brl(inv.valorInvestido)}</p>
        </div>
        <div>
          <p style="font-size:11px;color:var(--muted);">Valor Atual</p>
          <p style="font-size:14px;font-weight:700;color:#2563eb;">${brl(valorAtual)}</p>
        </div>
        <div>
          <p style="font-size:11px;color:var(--muted);">Rendimentos</p>
          <p style="font-size:14px;font-weight:700;color:#059669;">${brl(totalRend)}</p>
        </div>
        <div>
          <p style="font-size:11px;color:var(--muted);">Retorno</p>
          <p style="font-size:14px;font-weight:700;color:${pctRetorno >= 0 ? '#059669' : '#e11d48'};">${pctRetorno >= 0 ? '+' : ''}${pctRetorno.toFixed(1)}%</p>
        </div>
      </div>
      ${breakEvenHtml}
      <div style="display:flex;gap:8px;flex-wrap:wrap;margin-top:14px;">
        ${inv.tipo === 'cripto' ? `
        <button onclick="abrirCompraCriptoModal('${inv.id}')" style="padding:8px 16px;background:#f59e0b;color:white;border:none;border-radius:8px;font-size:12px;font-weight:600;cursor:pointer;">+ Nova Compra</button>
        <button onclick="verComprasCripto('${inv.id}')" style="padding:8px 16px;background:var(--tint-amber);color:var(--on-amber);border:none;border-radius:8px;font-size:12px;font-weight:600;cursor:pointer;">Compras</button>
        ` : inv.tipo === 'previdencia' ? `
        <button onclick="abrirAporteModal('${inv.id}')" style="padding:8px 16px;background:var(--on-rose);color:white;border:none;border-radius:8px;font-size:12px;font-weight:600;cursor:pointer;">+ Aporte</button>
        <button onclick="abrirAtualizarSaldoModal('${inv.id}')" style="padding:8px 16px;background:#059669;color:white;border:none;border-radius:8px;font-size:12px;font-weight:600;cursor:pointer;">Atualizar Saldo</button>
        <button onclick="verRendimentos('${inv.id}')" style="padding:8px 16px;background:var(--tint-blue);color:#2563eb;border:none;border-radius:8px;font-size:12px;font-weight:600;cursor:pointer;">Histórico</button>
        ` : `
        <button onclick="abrirRendimentoModal('${inv.id}')" style="padding:8px 16px;background:#059669;color:white;border:none;border-radius:8px;font-size:12px;font-weight:600;cursor:pointer;">+ Rendimento</button>
        <button onclick="verRendimentos('${inv.id}')" style="padding:8px 16px;background:var(--tint-blue);color:#2563eb;border:none;border-radius:8px;font-size:12px;font-weight:600;cursor:pointer;">Histórico</button>
        `}
        <button onclick="editarInvest('${inv.id}')" style="padding:8px 16px;background:var(--tint-indigo);color:#4f46e5;border:none;border-radius:8px;font-size:12px;font-weight:600;cursor:pointer;">Editar</button>
        <button onclick="deletarInvest('${inv.id}')" style="padding:8px 16px;background:var(--tint-rose);color:#e11d48;border:none;border-radius:8px;font-size:12px;font-weight:600;cursor:pointer;">Excluir</button>
      </div>
    </div>`;
  }).join('');
}

function calcValorAtual(inv) {
  if (inv.tipo === 'cripto' && inv.criptoTokens && inv.criptoPrecoAtual) {
    const dolarAtual = S.settings.usdRate || inv.criptoPtaxCompra || 5.0;
    return inv.criptoTokens * inv.criptoPrecoAtual * dolarAtual;
  }
  return inv.valorAtual || inv.valorInvestido || 0;
}

function calcTotalRendimentos(inv) {
  return (inv.rendimentos || []).filter(r => r.tipo !== 'aporte').reduce((s, r) => s + calcRendBRL(r, inv), 0);
}

function calcMesesPassados(dataInicio) {
  if (!dataInicio) return 0;
  const inicio = new Date(dataInicio + 'T12:00:00');
  const agora = new Date();
  return Math.max(0, (agora.getFullYear() - inicio.getFullYear()) * 12 + (agora.getMonth() - inicio.getMonth()));
}

function toggleInvestFields() {
  const tipo = document.getElementById('inv-tipo').value;
  document.getElementById('inv-negocio-fields').style.display = (tipo === 'negocio') ? 'block' : 'none';
  document.getElementById('inv-renda-fixa-fields').style.display = (tipo === 'renda_fixa') ? 'block' : 'none';
  document.getElementById('inv-cripto-fields').style.display = (tipo === 'cripto') ? 'block' : 'none';
  document.getElementById('inv-previdencia-fields').style.display = (tipo === 'previdencia') ? 'block' : 'none';
  // Quando é cripto, esconder "Valor Investido" e "Valor Atual" normais (são calculados)
  const valorRow = document.getElementById('inv-valor').closest('div').parentElement;
  if (tipo === 'cripto') {
    document.getElementById('inv-valor').closest('div').style.display = 'none';
    document.getElementById('inv-valor-atual').closest('div').style.display = 'none';
  } else {
    document.getElementById('inv-valor').closest('div').style.display = '';
    document.getElementById('inv-valor-atual').closest('div').style.display = '';
  }
}

function calcCriptoBRL() {
  const usd = parseFloat(document.getElementById('inv-cripto-usd').value) || 0;
  const ptax = parseFloat(document.getElementById('inv-cripto-ptax').value) || 0;
  const resumo = document.getElementById('inv-cripto-resumo');
  if (usd > 0 && ptax > 0) {
    const brlVal = usd * ptax;
    resumo.innerHTML = `Valor investido em BRL: <strong>R$ ${brlVal.toFixed(2)}</strong> (US$ ${usd.toFixed(2)} × R$ ${ptax.toFixed(2)})`;
    resumo.style.display = 'block';
  } else {
    resumo.style.display = 'none';
  }
}

function calcIRAliquota(dataInicio) {
  if (!dataInicio) return 0.225;
  const dias = Math.floor((new Date() - new Date(dataInicio + 'T12:00:00')) / (1000*60*60*24));
  if (dias <= 180) return 0.225;
  if (dias <= 360) return 0.20;
  if (dias <= 720) return 0.175;
  return 0.15;
}

function calcRendBRL(r, inv) {
  const valor = r.valor || 0;
  if (r.moeda === 'USD') {
    const taxa = S.settings.usdRate || 5.0;
    return valor * taxa;
  }
  return valor;
}

function abrirNovoInvestModal(editId) {
  const modal = document.getElementById('invest-modal');
  modal.style.display = 'block';
  document.getElementById('inv-edit-id').value = editId || '';
  document.getElementById('invest-modal-title').textContent = editId ? 'Editar Investimento' : 'Novo Investimento';

  const bankSel = document.getElementById('inv-banco');
  bankSel.innerHTML = '<option value="">Nenhum</option>' + Object.entries(BANKS).map(([k,v]) => `<option value="${k}">${v.name}</option>`).join('');

  const titSel = document.getElementById('inv-titular');
  titSel.innerHTML = `<option value="${S.settings.u1}">${S.settings.u1}</option><option value="${S.settings.u2}">${S.settings.u2}</option>`;

  if (editId) {
    const inv = S.investments.find(x => x.id === editId);
    if (inv) {
      document.getElementById('inv-tipo').value = inv.tipo || 'outro';
      bankSel.value = inv.bank || '';
      document.getElementById('inv-desc').value = inv.desc;
      document.getElementById('inv-valor').value = inv.valorInvestido;
      document.getElementById('inv-valor-atual').value = inv.valorAtual || '';
      document.getElementById('inv-retorno-mensal').value = inv.retornoMensal || '';
      document.getElementById('inv-moeda-retorno').value = inv.moedaRetorno || 'BRL';
      document.getElementById('inv-rendimento-pct').value = inv.rendimentoPct || '';
      document.getElementById('inv-indice').value = inv.indice || 'cdi';
      document.getElementById('inv-objetivo').value = inv.objetivo || '';
      document.getElementById('inv-data-inicio').value = inv.dataInicio || '';
      document.getElementById('inv-titular').value = inv.titular || S.settings.u1;
      document.getElementById('inv-notas').value = inv.notas || '';
      // Cripto fields
      const temCompras = inv.compras && inv.compras.length > 0;
      document.getElementById('inv-cripto-usd').value = inv.criptoUSD || '';
      document.getElementById('inv-cripto-tokens').value = inv.criptoTokens || '';
      document.getElementById('inv-cripto-ptax').value = inv.criptoPtaxCompra || '';
      document.getElementById('inv-cripto-preco-atual').value = inv.criptoPrecoAtual || '';
      // Previdência fields
      document.getElementById('inv-prev-tipo').value = inv.prevTipo || 'VGBL';
      document.getElementById('inv-prev-taxa-admin').value = inv.prevTaxaAdmin || '';
      document.getElementById('inv-prev-aporte').value = inv.prevAporteMensal || '';
      document.getElementById('inv-prev-rent12m').value = inv.prevRent12m || '';
      // Se já tem compras, desabilitar campos que são controlados pelo histórico
      document.getElementById('inv-cripto-usd').disabled = temCompras;
      document.getElementById('inv-cripto-tokens').disabled = temCompras;
      document.getElementById('inv-cripto-ptax').disabled = temCompras;
      if (inv.tipo === 'cripto') {
        calcCriptoBRL();
        if (temCompras) {
          document.getElementById('inv-cripto-resumo').innerHTML = `Valores controlados pelo histórico de compras (${inv.compras.length} compra${inv.compras.length!==1?'s':''}). Use <strong>"+ Nova Compra"</strong> para adicionar.`;
          document.getElementById('inv-cripto-resumo').style.display = 'block';
          document.getElementById('inv-cripto-resumo').style.background = 'var(--tint-blue)';
          document.getElementById('inv-cripto-resumo').style.color = 'var(--on-blue)';
        }
      }
    }
  } else {
    document.getElementById('inv-tipo').value = 'reserva';
    document.getElementById('inv-desc').value = '';
    document.getElementById('inv-valor').value = '';
    document.getElementById('inv-valor-atual').value = '';
    document.getElementById('inv-retorno-mensal').value = '';
    document.getElementById('inv-moeda-retorno').value = 'BRL';
    document.getElementById('inv-rendimento-pct').value = '';
    document.getElementById('inv-indice').value = 'cdi';
    document.getElementById('inv-objetivo').value = '';
    document.getElementById('inv-data-inicio').value = new Date().toISOString().slice(0,10);
    document.getElementById('inv-notas').value = '';
    // Cripto fields
    document.getElementById('inv-cripto-usd').value = '';
    document.getElementById('inv-cripto-tokens').value = '';
    document.getElementById('inv-cripto-ptax').value = '';
    document.getElementById('inv-cripto-preco-atual').value = '';
    document.getElementById('inv-cripto-usd').disabled = false;
    document.getElementById('inv-cripto-tokens').disabled = false;
    document.getElementById('inv-cripto-ptax').disabled = false;
    document.getElementById('inv-cripto-resumo').style.display = 'none';
    document.getElementById('inv-cripto-resumo').style.background = 'var(--tint-green)';
    document.getElementById('inv-cripto-resumo').style.color = 'var(--on-green)';
    // Previdência fields reset
    document.getElementById('inv-prev-tipo').value = 'VGBL';
    document.getElementById('inv-prev-taxa-admin').value = '';
    document.getElementById('inv-prev-aporte').value = '';
    document.getElementById('inv-prev-rent12m').value = '';
  }
  toggleInvestFields();
}

function fecharInvestModal() {
  document.getElementById('invest-modal').style.display = 'none';
}

function salvarInvest() {
  const tipo = document.getElementById('inv-tipo').value;
  const banco = document.getElementById('inv-banco').value;
  const desc = document.getElementById('inv-desc').value.trim();
  const valorInvestido = parseFloat(document.getElementById('inv-valor').value) || 0;
  const valorAtual = parseFloat(document.getElementById('inv-valor-atual').value) || 0;
  const retornoMensal = parseFloat(document.getElementById('inv-retorno-mensal').value) || 0;
  const rendimentoPct = parseFloat(document.getElementById('inv-rendimento-pct').value) || 0;
  const objetivo = document.getElementById('inv-objetivo').value.trim();
  const dataInicio = document.getElementById('inv-data-inicio').value;
  const titular = document.getElementById('inv-titular').value;
  const notas = document.getElementById('inv-notas').value.trim();
  const editId = document.getElementById('inv-edit-id').value;

  // Cripto: calcular valorInvestido e valorAtual a partir dos campos específicos
  let finalValorInvestido = valorInvestido;
  let finalValorAtual = valorAtual;
  let criptoData = {};

  if (tipo === 'cripto') {
    const criptoUSD = parseFloat(document.getElementById('inv-cripto-usd').value) || 0;
    const criptoTokens = parseFloat(document.getElementById('inv-cripto-tokens').value) || 0;
    const criptoPtax = parseFloat(document.getElementById('inv-cripto-ptax').value) || 0;
    const criptoPrecoAtual = parseFloat(document.getElementById('inv-cripto-preco-atual').value) || 0;

    // Se editando e já tem compras, só atualiza preço atual e dados gerais
    const existingInv = editId ? S.investments.find(x => x.id === editId) : null;
    if (existingInv && existingInv.compras && existingInv.compras.length > 0) {
      if (!desc) { toast('Preencha a descrição'); return; }
      criptoData = { criptoPrecoAtual: criptoPrecoAtual };
      // Manter totais das compras existentes
      finalValorInvestido = existingInv.valorInvestido;
      const dolarAtual = S.settings.usdRate || existingInv.criptoPtaxCompra || 5.0;
      finalValorAtual = existingInv.criptoTokens * criptoPrecoAtual * dolarAtual;
    } else {
      if (!desc || criptoUSD <= 0 || criptoTokens <= 0) {
        toast('Preencha descrição, valor em USD e tokens');
        return;
      }
      finalValorInvestido = criptoUSD * (criptoPtax || 1);
      const dolarAtual = S.settings.usdRate || criptoPtax || 5.0;
      finalValorAtual = criptoTokens * criptoPrecoAtual * dolarAtual;

      criptoData = {
        criptoUSD: criptoUSD,
        criptoTokens: criptoTokens,
        criptoPtaxCompra: criptoPtax,
        criptoPrecoAtual: criptoPrecoAtual,
        compras: [{
          id: 'cc_' + Date.now(),
          data: dataInicio || new Date().toISOString().slice(0,10),
          usd: criptoUSD,
          tokens: criptoTokens,
          ptax: criptoPtax || 5.0,
          obs: 'Compra inicial'
        }]
      };
    }
  } else if (!desc || valorInvestido <= 0) {
    toast('Preencha descrição e valor investido');
    return;
  }

  // Previdência: dados específicos
  let prevData = {};
  if (tipo === 'previdencia') {
    prevData = {
      prevTipo: document.getElementById('inv-prev-tipo')?.value || 'VGBL',
      prevTaxaAdmin: parseFloat(document.getElementById('inv-prev-taxa-admin')?.value) || 0,
      prevAporteMensal: parseFloat(document.getElementById('inv-prev-aporte')?.value) || 0,
      prevRent12m: parseFloat(document.getElementById('inv-prev-rent12m')?.value) || 0
    };
  }

  S.investments = S.investments || [];

  const moedaRetorno = document.getElementById('inv-moeda-retorno')?.value || 'BRL';
  const indice = document.getElementById('inv-indice')?.value || 'cdi';

  const data = {
    tipo, bank: banco, desc, valorInvestido: finalValorInvestido,
    valorAtual: finalValorAtual || finalValorInvestido,
    retornoMensal, rendimentoPct, moedaRetorno, indice,
    objetivo, dataInicio, titular, notas,
    ...criptoData, ...prevData
  };

  if (editId) {
    const inv = S.investments.find(x => x.id === editId);
    if (inv) { Object.assign(inv, data); inv.updatedAt = new Date().toISOString(); }
  } else {
    S.investments.push({
      id: 'inv_' + Date.now(),
      ...data,
      rendimentos: [],
      at: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    });
  }

  save();
  fecharInvestModal();
  renderInvestimentos();
  toast(editId ? 'Investimento atualizado!' : 'Investimento adicionado!');
}

function editarInvest(id) {
  abrirNovoInvestModal(id);
}

function editarPrecoAtualCripto(id) {
  const inv = S.investments.find(x => x.id === id);
  if (!inv) return;
  const novo = prompt('Preço atual do token em USD:', inv.criptoPrecoAtual || '');
  if (novo === null) return;
  const val = parseFloat(novo);
  if (isNaN(val) || val < 0) { toast('Valor inválido'); return; }
  inv.criptoPrecoAtual = val;
  const dolarAtual = S.settings.usdRate || inv.criptoPtaxCompra || 5.0;
  inv.valorAtual = inv.criptoTokens * val * dolarAtual;
  inv.updatedAt = new Date().toISOString();
  save();
  renderInvestimentos();
  toast('Preço atualizado!');
}

// ─── COMPRAS CRIPTO (múltiplas compras) ──────────────────────────────────────

// Migra cripto antiga (campos avulsos) para o array compras[]
function migrarCriptoCompras(inv) {
  if (inv.tipo !== 'cripto' || inv.compras) return;
  if (inv.criptoUSD && inv.criptoTokens) {
    inv.compras = [{
      id: 'cc_' + Date.now(),
      data: inv.dataInicio || new Date().toISOString().slice(0,10),
      usd: inv.criptoUSD,
      tokens: inv.criptoTokens,
      ptax: inv.criptoPtaxCompra || 5.0,
      obs: 'Compra inicial'
    }];
  } else {
    inv.compras = [];
  }
}

// Recalcula totais do cripto a partir das compras
function recalcCriptoTotais(inv) {
  if (inv.tipo !== 'cripto' || !inv.compras) return;
  const compras = inv.compras;
  inv.criptoUSD = compras.reduce((s, c) => s + (c.usd || 0), 0);
  inv.criptoTokens = compras.reduce((s, c) => s + (c.tokens || 0), 0);
  // Média ponderada da PTAX
  const totalBRL = compras.reduce((s, c) => s + (c.usd || 0) * (c.ptax || 5), 0);
  inv.criptoPtaxCompra = inv.criptoUSD > 0 ? totalBRL / inv.criptoUSD : 5.0;
  inv.valorInvestido = totalBRL;
  const dolarAtual = S.settings.usdRate || inv.criptoPtaxCompra || 5.0;
  inv.valorAtual = inv.criptoTokens * (inv.criptoPrecoAtual || 0) * dolarAtual;
}

function abrirCompraCriptoModal(id) {
  const inv = S.investments.find(x => x.id === id);
  if (!inv) return;
  migrarCriptoCompras(inv);

  document.getElementById('compra-cripto-modal').style.display = 'block';
  document.getElementById('cc-invest-id').value = id;
  document.getElementById('cc-usd').value = '';
  document.getElementById('cc-tokens').value = '';
  document.getElementById('cc-ptax').value = S.settings.usdRate ? S.settings.usdRate.toFixed(2) : '';
  document.getElementById('cc-data').value = new Date().toISOString().slice(0,10);
  document.getElementById('cc-obs').value = '';

  const nCompras = (inv.compras || []).length;
  document.getElementById('compra-cripto-info').innerHTML = `
    <p style="font-size:14px;font-weight:700;color:var(--text);">${escapeHtml(inv.desc)}</p>
    <p style="font-size:12px;color:var(--text-3);margin-top:4px;">Total: US$ ${inv.criptoUSD.toFixed(2)} · ${inv.criptoTokens.toLocaleString('pt-BR',{maximumFractionDigits:6})} tokens · ${nCompras} compra${nCompras!==1?'s':''}</p>
  `;
}

function fecharCompraCriptoModal() {
  document.getElementById('compra-cripto-modal').style.display = 'none';
}

function confirmarCompraCripto() {
  const id = document.getElementById('cc-invest-id').value;
  const usd = parseFloat(document.getElementById('cc-usd').value) || 0;
  const tokens = parseFloat(document.getElementById('cc-tokens').value) || 0;
  const ptax = parseFloat(document.getElementById('cc-ptax').value) || 0;
  const data = document.getElementById('cc-data').value;
  const obs = document.getElementById('cc-obs').value.trim();

  if (usd <= 0 || tokens <= 0 || ptax <= 0 || !data) {
    toast('Preencha valor USD, tokens, PTAX e data');
    return;
  }

  const inv = S.investments.find(x => x.id === id);
  if (!inv) return;

  migrarCriptoCompras(inv);

  inv.compras.push({
    id: 'cc_' + Date.now(),
    data, usd, tokens, ptax, obs
  });

  recalcCriptoTotais(inv);
  inv.updatedAt = new Date().toISOString();
  save();
  fecharCompraCriptoModal();
  renderInvestimentos();
  toast('Compra registrada!');
}

function verComprasCripto(id) {
  const inv = S.investments.find(x => x.id === id);
  if (!inv) return;
  migrarCriptoCompras(inv);

  if (!(inv.compras || []).length) {
    toast('Nenhuma compra registrada');
    return;
  }

  const compras = [...inv.compras].sort((a,b) => new Date(a.data) - new Date(b.data));
  const rows = compras.map((c, i) => {
    const brlVal = c.usd * c.ptax;
    const precoToken = c.tokens > 0 ? (c.usd / c.tokens) : 0;
    return `<tr style="border-bottom:1px solid var(--surface-2);">
      <td style="padding:8px 12px;font-size:13px;">${fmtDate(c.data)}</td>
      <td style="padding:8px 12px;text-align:right;font-size:13px;font-weight:600;">US$ ${c.usd.toFixed(2)}</td>
      <td style="padding:8px 12px;text-align:right;font-size:13px;">${c.tokens.toLocaleString('pt-BR',{maximumFractionDigits:6})}</td>
      <td style="padding:8px 12px;text-align:right;font-size:13px;color:var(--text-3);">R$ ${c.ptax.toFixed(2)}</td>
      <td style="padding:8px 12px;text-align:right;font-size:13px;font-weight:600;color:#2563eb;">${brl(brlVal)}</td>
      <td style="padding:8px 12px;text-align:center;">
        <button onclick="deletarCompraCripto('${inv.id}','${c.id}')" style="background:none;border:none;color:#e11d48;cursor:pointer;font-size:14px;" title="Excluir" class="ico-btn" aria-label="Excluir"><svg class="ico-sm"><use href="#i-trash"/></svg></button>
      </td>
    </tr>`;
  }).join('');

  const modal = document.getElementById('amort-table-modal');
  modal.style.display = 'block';
  document.getElementById('amort-modal-title').textContent = `Histórico de Compras — ${inv.desc}`;

  const totalUSD = inv.criptoUSD;
  const totalTokens = inv.criptoTokens;
  const totalBRL = inv.valorInvestido;
  const pmToken = totalTokens > 0 ? (totalUSD / totalTokens) : 0;

  document.getElementById('amort-resumo').innerHTML = `
    <div style="background:var(--bg);border-radius:10px;padding:12px 16px;">
      <p style="font-size:11px;color:var(--muted);">Total Investido</p>
      <p style="font-size:16px;font-weight:800;color:var(--text);">US$ ${totalUSD.toFixed(2)}</p>
      <p style="font-size:11px;color:var(--text-3);">${brl(totalBRL)}</p>
    </div>
    <div style="background:var(--bg);border-radius:10px;padding:12px 16px;">
      <p style="font-size:11px;color:var(--muted);">Total Tokens</p>
      <p style="font-size:16px;font-weight:800;color:var(--text-2);">${totalTokens.toLocaleString('pt-BR',{maximumFractionDigits:6})}</p>
    </div>
    <div style="background:var(--bg);border-radius:10px;padding:12px 16px;">
      <p style="font-size:11px;color:var(--muted);">Preço Médio / Token</p>
      <p style="font-size:16px;font-weight:800;color:#4f46e5;">US$ ${pmToken.toFixed(6)}</p>
    </div>
  `;

  document.getElementById('amort-thead').innerHTML = `
    <tr style="background:var(--bg);">
      <th style="padding:10px 12px;text-align:left;font-size:12px;">Data</th>
      <th style="padding:10px 12px;text-align:right;font-size:12px;">USD</th>
      <th style="padding:10px 12px;text-align:right;font-size:12px;">Tokens</th>
      <th style="padding:10px 12px;text-align:right;font-size:12px;">PTAX</th>
      <th style="padding:10px 12px;text-align:right;font-size:12px;">BRL</th>
      <th style="padding:10px 12px;text-align:center;font-size:12px;"></th>
    </tr>`;

  document.getElementById('amort-tbody').innerHTML = rows;
  document.getElementById('amort-tfoot').innerHTML = `
    <tr style="border-top:2px solid var(--border);background:var(--bg);font-weight:800;">
      <td style="padding:10px 12px;">TOTAL</td>
      <td style="padding:10px 12px;text-align:right;">US$ ${totalUSD.toFixed(2)}</td>
      <td style="padding:10px 12px;text-align:right;">${totalTokens.toLocaleString('pt-BR',{maximumFractionDigits:6})}</td>
      <td style="padding:10px 12px;"></td>
      <td style="padding:10px 12px;text-align:right;color:#2563eb;">${brl(totalBRL)}</td>
      <td></td>
    </tr>`;
}

function deletarCompraCripto(invId, compraId) {
  if (!confirm('Excluir esta compra?')) return;
  const inv = S.investments.find(x => x.id === invId);
  if (!inv || !inv.compras) return;
  inv.compras = inv.compras.filter(c => c.id !== compraId);
  recalcCriptoTotais(inv);
  inv.updatedAt = new Date().toISOString();
  save();
  renderInvestimentos();
  // Reabrir o histórico atualizado
  if (inv.compras.length > 0) verComprasCripto(invId);
  else { document.getElementById('amort-table-modal').style.display = 'none'; }
  toast('Compra excluída!');
}

function deletarInvest(id) {
  if (!confirm('Tem certeza que deseja excluir este investimento?')) return;
  S.deletedIds.push({ id, collection: 'investments', deletedAt: new Date().toISOString() });
  S.investments = S.investments.filter(i => i.id !== id);
  save();
  renderInvestimentos();
  toast('Investimento excluído!');
}

// Rendimentos
function abrirRendimentoModal(id) {
  const inv = S.investments.find(x => x.id === id);
  if (!inv) return;
  document.getElementById('rendimento-modal').style.display = 'block';
  document.getElementById('rend-invest-id').value = id;
  document.getElementById('rend-valor').value = inv.retornoMensal || '';
  document.getElementById('rend-moeda').value = inv.moedaRetorno || 'BRL';
  document.getElementById('rend-data').value = new Date().toISOString().slice(0,10);
  document.getElementById('rend-pct').value = '';
  document.getElementById('rend-obs').value = '';

  const totalRend = calcTotalRendimentos(inv);
  const moedaLabel = inv.moedaRetorno === 'USD' ? 'USD' : 'BRL';
  const totalOriginal = (inv.rendimentos||[]).reduce((s,r) => s + (r.valor||0), 0);
  document.getElementById('rend-invest-info').innerHTML = `
    <p style="font-size:14px;font-weight:700;color:var(--text);">${escapeHtml(inv.desc)}</p>
    <p style="font-size:12px;color:var(--text-3);margin-top:4px;">Investido: ${brl(inv.valorInvestido)} · Rendimentos: ${brl(totalRend)}${inv.moedaRetorno === 'USD' ? ' ('+totalOriginal.toFixed(2)+' USD)' : ''}</p>
    ${inv.tipo === 'renda_fixa' ? '<p style="font-size:11px;color:var(--muted);margin-top:2px;">IR: '+Math.round(calcIRAliquota(inv.dataInicio)*100)+'% ('+Math.floor((new Date()-new Date(inv.dataInicio+"T12:00:00"))/(1000*60*60*24))+' dias)</p>' : ''}
  `;
}

function fecharRendimentoModal() {
  document.getElementById('rendimento-modal').style.display = 'none';
}

function confirmarRendimento() {
  const id = document.getElementById('rend-invest-id').value;
  const valor = parseFloat(document.getElementById('rend-valor').value) || 0;
  const moeda = document.getElementById('rend-moeda').value;
  const data = document.getElementById('rend-data').value;
  const pct = parseFloat(document.getElementById('rend-pct').value) || 0;
  const obs = document.getElementById('rend-obs').value.trim();

  if (valor <= 0 || !data) { toast('Preencha valor e data'); return; }

  const inv = S.investments.find(x => x.id === id);
  if (!inv) return;

  inv.rendimentos = inv.rendimentos || [];
  inv.rendimentos.push({ valor, moeda, data, pct, obs, at: new Date().toISOString(), updatedAt: new Date().toISOString() });
  inv.updatedAt = new Date().toISOString();

  // Atualizar valor atual para renda fixa
  if (inv.tipo === 'renda_fixa' && pct > 0) {
    inv.valorAtual = (inv.valorAtual || inv.valorInvestido) * (1 + pct / 100);
    inv.rendimentoPct = pct;
  }

  const valorBRL = moeda === 'USD' ? valor * (S.settings.usdRate || 5.0) : valor;
  const moedaStr = moeda === 'USD' ? ` (US$ ${valor.toFixed(2)})` : '';

  save();
  fecharRendimentoModal();
  renderInvestimentos();
  toast(`Rendimento de ${brl(valorBRL)}${moedaStr} registrado!`);
}

// ─── APORTE PREVIDÊNCIA ─────────────────────────────────────────────────────

function abrirAporteModal(id) {
  const inv = S.investments.find(x => x.id === id);
  if (!inv) return;

  document.getElementById('aporte-modal').style.display = 'block';
  document.getElementById('aporte-invest-id').value = id;
  document.getElementById('aporte-valor').value = inv.prevAporteMensal || '';
  document.getElementById('aporte-data').value = new Date().toISOString().slice(0,10);
  document.getElementById('aporte-valor-atual').value = '';
  document.getElementById('aporte-obs').value = '';

  // Populate conta select with user's bank accounts
  const titular = inv.titular;
  const contas = S.accounts.filter(a => a.accountType === 'conta' && (!a.owner || a.owner === titular));
  const sel = document.getElementById('aporte-conta');
  sel.innerHTML = '<option value="">Não descontar</option>' +
    contas.map(a => {
      const b = BANKS[a.bank] || { name: a.bank };
      return `<option value="${a.id}">${escapeHtml(b.name)} — ${escapeHtml(a.label)}</option>`;
    }).join('');

  const valorAtual = calcValorAtual(inv);
  document.getElementById('aporte-invest-info').innerHTML = `
    <p style="font-size:14px;font-weight:700;color:var(--text);">${escapeHtml(inv.desc)} (${inv.prevTipo || 'VGBL'})</p>
    <p style="font-size:12px;color:var(--text-3);margin-top:4px;">Investido: ${brl(inv.valorInvestido)} · Valor Atual: ${brl(valorAtual)}</p>
    <p style="font-size:11px;color:var(--muted);margin-top:2px;">Aporte mensal: ${brl(inv.prevAporteMensal || 0)} · ${calcMesesPassados(inv.dataInicio)} meses</p>
  `;
}

function fecharAporteModal() {
  document.getElementById('aporte-modal').style.display = 'none';
}

function confirmarAporte() {
  const id = document.getElementById('aporte-invest-id').value;
  const valor = parseFloat(document.getElementById('aporte-valor').value) || 0;
  const data = document.getElementById('aporte-data').value;
  const novoValorAtual = parseFloat(document.getElementById('aporte-valor-atual').value) || 0;
  const contaId = document.getElementById('aporte-conta').value;
  const obs = document.getElementById('aporte-obs').value.trim();

  if (valor <= 0 || !data) { toast('Preencha valor e data do aporte'); return; }
  if (novoValorAtual <= 0) { toast('Informe o novo valor atual da previdência'); return; }

  const inv = S.investments.find(x => x.id === id);
  if (!inv) return;

  // 1. Atualizar investimento
  inv.valorInvestido = (inv.valorInvestido || 0) + valor;
  inv.valorAtual = novoValorAtual;
  inv.updatedAt = new Date().toISOString();

  // Registrar aporte no histórico de rendimentos (com flag de aporte)
  inv.rendimentos = inv.rendimentos || [];
  inv.rendimentos.push({
    valor: valor,
    moeda: 'BRL',
    data: data,
    pct: 0,
    obs: obs || `Aporte ${inv.prevTipo || 'VGBL'}`,
    tipo: 'aporte',
    at: new Date().toISOString()
  });

  // 2. Criar transação de débito na conta bancária (se selecionada)
  if (contaId) {
    S.transactions.push({
      id: Date.now().toString(),
      type: 'despesa',
      desc: `Aporte ${inv.prevTipo || 'Previdência'} — ${inv.desc}`,
      amount: valor,
      category: 'Investimentos',
      subcategory: '',
      date: data,
      user: inv.titular,
      accountId: contaId,
      formaPgto: 'debito',
      pago: true,
      at: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    });
  }

  save();
  fecharAporteModal();
  renderInvestimentos();
  toast(`Aporte de ${brl(valor)} registrado!${contaId ? ' Débito criado na conta.' : ''}`);
}

// ─── ATUALIZAR SALDO PREVIDÊNCIA ────────────────────────────────────────────

function abrirAtualizarSaldoModal(id) {
  const inv = S.investments.find(x => x.id === id);
  if (!inv) return;

  document.getElementById('atualizar-saldo-modal').style.display = 'block';
  document.getElementById('atsaldo-invest-id').value = id;
  document.getElementById('atsaldo-valor').value = '';
  document.getElementById('atsaldo-data').value = new Date().toISOString().slice(0,10);
  document.getElementById('atsaldo-diff').style.display = 'none';

  const valorAtual = calcValorAtual(inv);
  document.getElementById('atsaldo-invest-info').innerHTML = `
    <p style="font-size:14px;font-weight:700;color:var(--text);">${escapeHtml(inv.desc)} (${inv.prevTipo || 'VGBL'})</p>
    <p style="font-size:12px;color:var(--text-3);margin-top:4px;">Investido: ${brl(inv.valorInvestido)} · Saldo atual: <strong>${brl(valorAtual)}</strong></p>
  `;
}

function fecharAtualizarSaldoModal() {
  document.getElementById('atualizar-saldo-modal').style.display = 'none';
}

function calcDiffSaldo() {
  const id = document.getElementById('atsaldo-invest-id').value;
  const inv = S.investments.find(x => x.id === id);
  if (!inv) return;

  const novoSaldo = parseFloat(document.getElementById('atsaldo-valor').value) || 0;
  const saldoAnterior = calcValorAtual(inv);
  const diff = novoSaldo - saldoAnterior;
  const el = document.getElementById('atsaldo-diff');

  if (novoSaldo > 0) {
    const isPos = diff >= 0;
    el.style.display = 'block';
    el.style.background = isPos ? 'var(--tint-green)' : 'var(--tint-rose)';
    el.style.color = isPos ? '#059669' : '#e11d48';
    el.innerHTML = `${isPos ? '📈 Rendimento' : '📉 Variação'}: <span style="font-size:16px;">${isPos ? '+' : ''}${brl(diff)}</span>`;
  } else {
    el.style.display = 'none';
  }
}

function confirmarAtualizarSaldo() {
  const id = document.getElementById('atsaldo-invest-id').value;
  const novoSaldo = parseFloat(document.getElementById('atsaldo-valor').value) || 0;
  const data = document.getElementById('atsaldo-data').value;

  if (novoSaldo <= 0 || !data) { toast('Preencha o novo saldo e a data'); return; }

  const inv = S.investments.find(x => x.id === id);
  if (!inv) return;

  const saldoAnterior = calcValorAtual(inv);
  const diff = novoSaldo - saldoAnterior;

  // Atualizar valor atual
  inv.valorAtual = novoSaldo;
  inv.updatedAt = new Date().toISOString();

  // Registrar a diferença como rendimento (se houve variação)
  if (Math.abs(diff) >= 0.01) {
    inv.rendimentos = inv.rendimentos || [];
    inv.rendimentos.push({
      valor: diff,
      moeda: 'BRL',
      data: data,
      pct: saldoAnterior > 0 ? parseFloat(((diff / saldoAnterior) * 100).toFixed(2)) : 0,
      obs: `Saldo atualizado: ${brl(saldoAnterior)} → ${brl(novoSaldo)}`,
      tipo: 'saldo',
      at: new Date().toISOString()
    });
  }

  save();
  fecharAtualizarSaldoModal();
  renderInvestimentos();

  if (Math.abs(diff) >= 0.01) {
    toast(`Saldo atualizado! ${diff >= 0 ? 'Rendimento' : 'Variação'}: ${diff >= 0 ? '+' : ''}${brl(diff)}`);
  } else {
    toast('Saldo atualizado!');
  }
}

function verRendimentos(id) {
  const inv = S.investments.find(x => x.id === id);
  if (!inv || !(inv.rendimentos || []).length) {
    toast('Nenhum rendimento registrado ainda');
    return;
  }

  const rends = [...inv.rendimentos].sort((a,b) => new Date(b.data) - new Date(a.data));
  const rows = rends.map(r => {
    const isAporte = r.tipo === 'aporte';
    const isSaldo = r.tipo === 'saldo';
    const valorBRL = calcRendBRL(r, inv);
    const moedaInfo = r.moeda === 'USD' ? ` <span style="color:var(--muted);">(US$ ${r.valor.toFixed(2)})</span>` : '';
    const pctInfo = r.pct ? ` <span style="color:#4f46e5;">${r.pct}%</span>` : '';
    const tagHtml = isAporte ? '<span style="background:var(--tint-pink);color:var(--on-rose);font-size:10px;font-weight:700;padding:2px 6px;border-radius:4px;margin-right:4px;">APORTE</span>'
      : isSaldo ? '<span style="background:var(--tint-green);color:var(--on-green);font-size:10px;font-weight:700;padding:2px 6px;border-radius:4px;margin-right:4px;">SALDO</span>' : '';
    const valColor = isAporte ? '#9d174d' : valorBRL >= 0 ? '#059669' : '#e11d48';
    return `<tr style="border-bottom:1px solid var(--surface-2);">
      <td style="padding:8px 12px;">${fmtDate(r.data)}</td>
      <td style="padding:8px 12px;text-align:right;font-weight:600;color:${valColor};">${brl(valorBRL)}${moedaInfo}</td>
      <td style="padding:8px 12px;color:var(--text-3);font-size:12px;">${tagHtml}${pctInfo} ${r.obs || '-'}</td>
    </tr>`;
  }).join('');

  // Reuse amort modal to show rendimentos
  const modal = document.getElementById('amort-table-modal');
  modal.style.display = 'block';
  document.getElementById('amort-modal-title').textContent = `Histórico de Rendimentos — ${inv.desc}`;

  const totalRend = calcTotalRendimentos(inv);
  const pctRecup = inv.valorInvestido > 0 ? ((totalRend / inv.valorInvestido) * 100).toFixed(1) : '0';

  document.getElementById('amort-resumo').innerHTML = `
    <div style="background:var(--bg);border-radius:10px;padding:12px 16px;">
      <p style="font-size:11px;color:var(--muted);">Total Rendimentos</p>
      <p style="font-size:16px;font-weight:800;color:#059669;">${brl(totalRend)}</p>
    </div>
    <div style="background:var(--bg);border-radius:10px;padding:12px 16px;">
      <p style="font-size:11px;color:var(--muted);">Registros</p>
      <p style="font-size:16px;font-weight:800;color:var(--text-2);">${rends.length}</p>
    </div>
    <div style="background:var(--bg);border-radius:10px;padding:12px 16px;">
      <p style="font-size:11px;color:var(--muted);">% Recuperado</p>
      <p style="font-size:16px;font-weight:800;color:#4f46e5;">${pctRecup}%</p>
    </div>
  `;

  document.getElementById('amort-tbody').innerHTML = rows;
  document.getElementById('amort-tfoot').innerHTML = `
    <tr style="border-top:2px solid var(--border);background:var(--bg);font-weight:800;">
      <td style="padding:10px 12px;">TOTAL</td>
      <td style="padding:10px 12px;text-align:right;color:#059669;">${brl(totalRend)}</td>
      <td style="padding:10px 12px;"></td>
    </tr>
  `;
}

// ─── ANÁLISE POR CATEGORIA (BARRAS EMPILHADAS) ──────────────────────────────
let catChartMode = 'despesa';

function setCatChartMode(mode) {
  catChartMode = mode;
  document.getElementById('cat-ch-despesa').style.cssText = `padding:4px 12px;border-radius:6px;border:none;font-size:12px;font-weight:600;cursor:pointer;${mode==='despesa'?'background:#4f46e5;color:white;':'background:transparent;color:#64748b;'}`;
  document.getElementById('cat-ch-receita').style.cssText = `padding:4px 12px;border-radius:6px;border:none;font-size:12px;font-weight:600;cursor:pointer;${mode==='receita'?'background:#059669;color:white;':'background:transparent;color:#64748b;'}`;
  renderCatStackedChart();
}

function renderCatStackedChart() {
  const yearSel = document.getElementById('cat-ch-year');
  if (!yearSel.options.length) {
    const years = new Set();
    S.transactions.forEach(t => years.add(parseInt(t.date.substring(0,4))));
    years.add(new Date().getFullYear());
    [...years].sort((a,b) => b-a).forEach(y => {
      const o = document.createElement('option');
      o.value = y; o.textContent = y;
      yearSel.appendChild(o);
    });
    yearSel.value = new Date().getFullYear();
  }

  const year = parseInt(yearSel.value);
  const txs = txByTitular(S.transactions).filter(t => {
    if (t.isTransfer) return false;
    if (t.type !== catChartMode) return false;
    if (catChartMode === 'despesa' && t.formaPgto === 'credito') return false;
    if (isExcludedFromChart(t)) return false;
    return t.date.startsWith(String(year));
  });

  // Group by category and month
  const catMonths = {};
  txs.forEach(t => {
    const cat = t.category || 'Outros';
    const month = parseInt(t.date.substring(5,7)) - 1;
    if (!catMonths[cat]) catMonths[cat] = new Array(12).fill(0);
    catMonths[cat][month] += amountBrl(t);
  });

  // Sort categories by total (descending), keep top 8, rest as "Outros"
  const sorted = Object.entries(catMonths).sort((a,b) => b[1].reduce((s,v)=>s+v,0) - a[1].reduce((s,v)=>s+v,0));
  const topCats = sorted.slice(0, 8);
  if (sorted.length > 8) {
    const outrosArr = new Array(12).fill(0);
    sorted.slice(8).forEach(([,arr]) => arr.forEach((v,i) => outrosArr[i] += v));
    topCats.push(['Outros', outrosArr]);
  }

  const grandTotal = txs.reduce((s,t) => s + amountBrl(t), 0);
  document.getElementById('cat-ch-total').textContent = `Total ${year}: ${brl(grandTotal)}`;

  // Legend
  document.getElementById('cat-ch-legend').innerHTML = topCats.map(([cat], i) =>
    `<span style="display:flex;align-items:center;gap:4px;"><span style="width:10px;height:10px;border-radius:3px;background:${COLORS[i % COLORS.length]};"></span>${escapeHtml(cat)}</span>`
  ).join('');

  // Chart
  destroyChart('ch-cat-stacked');
  const ctx = document.getElementById('ch-cat-stacked');
  if (!ctx || !topCats.length) return;

  const datasets = topCats.map(([cat, data], i) => ({
    label: cat,
    data: data,
    backgroundColor: COLORS[i % COLORS.length],
    borderWidth: 0,
    borderRadius: 3
  }));

  charts['ch-cat-stacked'] = new Chart(ctx, {
    type: 'bar',
    data: { labels: MESES, datasets },
    options: {
      responsive: true, maintainAspectRatio: false,
      scales: {
        x: { stacked: true, grid: { display: false } },
        y: {
          stacked: true,
          grid: { color: cssVar('--chart-grid') },
          ticks: { callback: v => 'R$' + (v>=1000 ? (v/1000).toFixed(1)+'k' : v.toFixed(0)) }
        }
      },
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: ctx => `${ctx.dataset.label}: ${brl(ctx.raw)}`
          }
        }
      }
    }
  });
}

let valuesHidden = false;
function toggleHideValues() {
  valuesHidden = !valuesHidden;
  document.getElementById('page-dashboard').classList.toggle('hide-values', valuesHidden);
  const btn = document.getElementById('btn-hide-values');
  btn.style.background = valuesHidden ? '#4f46e5' : 'var(--surface)';
  btn.style.borderColor = valuesHidden ? '#4f46e5' : 'var(--border)';
  document.getElementById('eye-icon').style.stroke = valuesHidden ? 'white' : 'var(--text-3)';
  // Update icon to eye-off when hidden
  document.getElementById('eye-icon').innerHTML = valuesHidden
    ? '<path d="M17.94 17.94A10.07 10.07 0 0 1 12 20c-7 0-11-8-11-8a18.45 18.45 0 0 1 5.06-5.94M9.9 4.24A9.12 9.12 0 0 1 12 4c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19m-6.72-1.07a3 3 0 1 1-4.24-4.24"/><line x1="1" y1="1" x2="23" y2="23"/>'
    : '<path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/>';
}

function initSalemBot() {
  const container = document.getElementById('salem-bot');
  if (container && typeof lottie !== 'undefined') {
    lottie.loadAnimation({
      container: container,
      renderer: 'svg',
      loop: true,
      autoplay: true,
      path: 'anima-bot.json'
    });
  }
}

function init() {
  load();
  // Cleanup tombstones older than 30 days
  const thirtyDaysAgo = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000).toISOString();
  S.deletedIds = (S.deletedIds || []).filter(t => t.deletedAt > thirtyDaysAgo);
  // Limpeza: remover conta Caixa automática que foi criada por erro
  const caixaIdx = S.accounts.findIndex(a => a.id === 'caixa_conta');
  if (caixaIdx >= 0) {
    S.accounts.splice(caixaIdx, 1);
    save(); // salva no localStorage E no Firebase pra não voltar
  }
  seedAccounts();
  // Migração _faturaRefFixedV3 desativada: sobrescrevia faturaRef manual de parcelas e quebrava faturas.
  // One-time: add Unimed Curitiba 3/3 (Paulo pediu em 03/04/2026)
  if (!S.settings._unimed3Added) {
    const exists = S.transactions.find(t => t.id === 'manual_unimed_3_3');
    if (!exists) {
      S.transactions.push({
        id: 'manual_unimed_3_3',
        type: 'despesa',
        desc: 'UNIMED CURITIBA 3/3',
        amount: 280,
        date: '2026-02-09',
        faturaRef: '2026-05',
        category: 'Saúde',
        subcategory: 'Médico',
        accountId: 'card_sicredi',
        user: S.settings.u1,
        formaPgto: 'credito',
        custoTipo: 'fixo',
        parcela: '3/3',
        pago: false,
        currency: 'BRL',
        isNegative: false,
        at: new Date().toISOString(),
        updatedAt: new Date().toISOString()
      });
    }
    S.settings._unimed3Added = true;
    save();
  }
  // One-time: remove R$28000 payment from debt history (Paulo will re-enter it)
  if (!S.settings._clean28k) {
    (S.debts || []).forEach(d => {
      const hadPgto = (d.pagamentos || []).some(p => p.amount === 28000);
      const hadAmort = (d.amortizacoesExtra || []).some(p => p.amount === 28000);
      if (hadPgto) {
        d.pagamentos = d.pagamentos.filter(p => p.amount !== 28000);
        d.parcelasPagas = Math.max(0, (d.parcelasPagas || 0) - 1);
      }
      if (hadAmort) {
        d.amortizacoesExtra = d.amortizacoesExtra.filter(p => p.amount !== 28000);
        // Restore principal that was reduced by the amortization
        d.principal = (d.principal || 0) + 28000;
      }
    });
    S.settings._clean28k = true;
    save();
  }
  importExcelData();
  initFirebaseSync();
  setType('receita');
  setTodayDate();
  refreshUserSelects();
  document.getElementById('header-users').textContent = `${S.settings.u1} & ${S.settings.u2}`;
  fetchUsdRate();
  initDashPeriod();
  initSalemBot();
  updateAIStatus();
  checkAuth();
  if (sessionStorage.getItem('fincasal_auth')) goto('dashboard');
}
