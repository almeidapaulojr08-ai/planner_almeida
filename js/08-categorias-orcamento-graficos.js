// FinançasCasal — 08-categorias-orcamento-graficos.js
// Módulo carregado em ordem por index.html (escopo global compartilhado). Não use import/export.

// ─── CATEGORIAS EDITÁVEIS ─────────────────────────────────────────────────────
const DEFAULT_CATS = {
  receita:      ['Salário','Projetos','Freelance','Bônus','Transferências','Renda Extra','Aluguel Recebido','Dividendos','Reembolso','Outros'],
  despesa:      {
    'Alimentação':  ['Restaurante','Delivery','Lanchonete/Bar','Mercado','Geral'],
    'Mercado':      [],
    'Saúde':        ['Farmácia','Academia/Fitness','Médico','Exames','Suplementos','Geral'],
    'Assinaturas':  ['Streaming','Software/App','Games','Outros'],
    'Lazer':        ['Passeio','Shows/Eventos','Bar/Balada','Hobbies','Geral'],
    'Viagem':       ['Hospedagem','Transporte','Alimentação','Passeio','Geral'],
    'Carro':        ['Combustível','IPVA/Licenciamento','Manutenção','Estacionamento','Pedágio'],
    'Vestuário':    [],
    'Educação':     ['Cursos','Livros','Material','Escola','Geral'],
    'Casa':         ['Internet','Luz','Água','Gás','Aluguel','Condomínio','Manutenção','Geral'],
    'Games':        [],
    'Presente':     [],
    'Pets':         ['Ração','Veterinário','Banho/Tosa','Geral'],
    'Outros':       [],
  },
  investimento: ['Renda Fixa','Ações','FIIs','Tesouro Direto','Criptomoedas','Poupança','Previdência','Outros']
};

function getCats() {
  if (!S.customCats) {
    if (!firebaseReady) {
      // Firebase ainda não carregou — retornar defaults SEM salvar em S.customCats
      // para não sobrescrever categorias reais do Firebase ao chamar save()
      return JSON.parse(JSON.stringify(DEFAULT_CATS));
    }
    S.customCats = JSON.parse(JSON.stringify(DEFAULT_CATS));
  }
  return S.customCats;
}

function getDespesaCats()   { return getCats().despesa; }
function getReceitaCats()   { return getCats().receita; }
function getInvestCats()    { return getCats().investimento; }

let catTab = 'despesa';

function setCatTab(tab) {
  catTab = tab;
  document.querySelectorAll('[id^="cattab-"]').forEach(b => b.classList.remove('tf-active'));
  document.getElementById('cattab-'+tab)?.classList.add('tf-active');
  renderCategorias();
}

function renderCategorias() {
  const cats = getCats();
  const list = document.getElementById('cat-list');
  const subAdd = document.getElementById('cat-sub-add');
  let html = '';

  if (catTab === 'despesa') {
    subAdd.style.display = 'block';
    const despCats = cats.despesa;
    const parentSel = document.getElementById('cat-sub-parent');
    parentSel.innerHTML = Object.keys(despCats).map(c => `<option value="${c}">${c}</option>`).join('');

    for (const [cat, subs] of Object.entries(despCats)) {
      const grupo = (S.catOrcGroup || {})[cat] || '';
      const grupoColor = grupo === 'essencial' ? '#10b981' : grupo === 'estilo' ? '#3b82f6' : 'var(--muted)';
      html += `<div class="card" style="padding:14px 18px;">
        <div style="display:flex;justify-content:space-between;align-items:center;">
          <div style="display:flex;align-items:center;gap:10px;">
            <span style="color:#e11d48;font-size:10px;">●</span>
            <input type="text" value="${escapeHtml(cat)}" onchange="renameCatDespesa('${cat}',this.value)"
              style="border:none;font-size:14px;font-weight:700;color:var(--text);background:transparent;padding:0;outline:none;width:160px;">
          </div>
          <div style="display:flex;align-items:center;gap:6px;">
            <select onchange="setCatOrcGroup('${cat}',this.value)" style="padding:4px 8px;border:1.5px solid ${grupoColor};border-radius:8px;font-size:11px;font-weight:600;color:${grupoColor};background:var(--surface);cursor:pointer;outline:none;">
              <option value="" ${!grupo ? 'selected' : ''} style="color:var(--muted);">Sem grupo</option>
              <option value="essencial" ${grupo === 'essencial' ? 'selected' : ''} style="color:#10b981;">Essencial</option>
              <option value="estilo" ${grupo === 'estilo' ? 'selected' : ''} style="color:#3b82f6;">Estilo de Vida</option>
            </select>
            <button onclick="deleteCatDespesa('${cat}')" style="background:none;border:none;cursor:pointer;padding:4px 8px;border-radius:6px;font-size:13px;color:#e11d48;" onmouseover="this.style.background='var(--tint-rose)'" onmouseout="this.style.background='none'" title="Excluir" class="ico-btn" aria-label="Fechar"><svg class="ico-sm"><use href="#i-x"/></svg></button>
          </div>
        </div>
        ${subs.length ? `<div style="margin-top:8px;padding-left:20px;display:flex;flex-wrap:wrap;gap:6px;">
          ${subs.map(s => `<span style="display:inline-flex;align-items:center;gap:4px;background:var(--bg);border:1px solid var(--border);border-radius:8px;padding:4px 10px;font-size:12px;color:var(--text-2);">
            ${s}
            <button onclick="deleteSubcat('${cat}','${s}')" style="background:none;border:none;cursor:pointer;color:var(--muted);font-size:11px;padding:0 0 0 4px;" class="ico-btn" aria-label="Fechar"><svg class="ico-sm"><use href="#i-x"/></svg></button>
          </span>`).join('')}
        </div>` : ''}
      </div>`;
    }
  } else {
    subAdd.style.display = 'none';
    const arr = catTab === 'receita' ? cats.receita : cats.investimento;
    arr.forEach((cat, i) => {
      const color = catTab === 'receita' ? '#10b981' : '#3b82f6';
      const icon = catTab === 'receita' ? '💚' : '📈';
      html += `<div class="card" style="padding:12px 18px;">
        <div style="display:flex;justify-content:space-between;align-items:center;">
          <div style="display:flex;align-items:center;gap:10px;">
            <span style="color:${color};font-size:10px;">●</span>
            <input type="text" value="${escapeHtml(cat)}" onchange="renameSimpleCat('${catTab}',${i},this.value)"
              style="border:none;font-size:14px;font-weight:600;color:var(--text);background:transparent;padding:0;outline:none;width:250px;">
          </div>
          <button onclick="deleteSimpleCat('${catTab}',${i})" style="background:none;border:none;cursor:pointer;padding:4px 8px;border-radius:6px;font-size:13px;color:#e11d48;" onmouseover="this.style.background='var(--tint-rose)'" onmouseout="this.style.background='none'" title="Excluir" class="ico-btn" aria-label="Fechar"><svg class="ico-sm"><use href="#i-x"/></svg></button>
        </div>
      </div>`;
    });
  }

  list.innerHTML = html;
}

function addCategory() {
  const name = document.getElementById('cat-new-name').value.trim();
  if (!name) return toast('❌ Informe o nome da categoria');
  const cats = getCats();
  if (catTab === 'despesa') {
    if (cats.despesa[name]) return toast('❌ Categoria já existe');
    cats.despesa[name] = [];
  } else {
    const arr = catTab === 'receita' ? cats.receita : cats.investimento;
    if (arr.includes(name)) return toast('❌ Categoria já existe');
    arr.push(name);
  }
  save();
  document.getElementById('cat-new-name').value = '';
  toast('✅ Categoria adicionada!');
  renderCategorias();
}

function addSubcategory() {
  const parent = document.getElementById('cat-sub-parent').value;
  const name = document.getElementById('cat-sub-name').value.trim();
  if (!name) return toast('❌ Informe o nome da subcategoria');
  const cats = getCats();
  if (!cats.despesa[parent]) return;
  if (cats.despesa[parent].includes(name)) return toast('❌ Subcategoria já existe');
  cats.despesa[parent].push(name);
  save();
  document.getElementById('cat-sub-name').value = '';
  toast('✅ Subcategoria adicionada!');
  renderCategorias();
}

function renameCatDespesa(oldName, newName) {
  newName = newName.trim();
  if (!newName || oldName === newName) return;
  const cats = getCats();
  const entries = Object.entries(cats.despesa);
  cats.despesa = {};
  for (const [k, v] of entries) {
    cats.despesa[k === oldName ? newName : k] = v;
  }
  S.transactions.forEach(t => { if (t.type === 'despesa' && t.category === oldName) t.category = newName; });
  save(); toast('✅ Categoria renomeada!'); renderCategorias();
}

function deleteCatDespesa(name) {
  const cats = getCats();
  delete cats.despesa[name];
  save(); toast('🗑️ Categoria removida'); renderCategorias();
}

function deleteSubcat(cat, sub) {
  const cats = getCats();
  cats.despesa[cat] = cats.despesa[cat].filter(s => s !== sub);
  save(); renderCategorias();
}

function renameSimpleCat(type, index, newName) {
  newName = newName.trim();
  const cats = getCats();
  const arr = type === 'receita' ? cats.receita : cats.investimento;
  const oldName = arr[index];
  if (!newName || oldName === newName) return;
  arr[index] = newName;
  S.transactions.forEach(t => { if (t.type === type && t.category === oldName) t.category = newName; });
  save(); toast('✅ Categoria renomeada!');
}

function deleteSimpleCat(type, index) {
  const cats = getCats();
  const arr = type === 'receita' ? cats.receita : cats.investimento;
  arr.splice(index, 1);
  save(); toast('🗑️ Categoria removida'); renderCategorias();
}

// ─── ORÇAMENTO (50/30/20) ─────────────────────────────────────────────────────
let orcAno = new Date().getFullYear();
let orcMes = new Date().getMonth();

const ORC_GROUPS = {
  'Essencial': {
    color: '#10b981', bg: 'var(--tint-green)', pct: 50,
    cats: ['Alimentação','Mercado','Saúde','Casa','Carro','Educação','Pets']
  },
  'Estilo de Vida': {
    color: '#3b82f6', bg: 'var(--tint-blue)', pct: 30,
    cats: ['Lazer','Vestuário','Assinaturas','Viagem','Games','Presente','Outros']
  },
  'Investimentos': {
    color: '#8b5cf6', bg: '#f5f3ff', pct: 20,
    cats: ['Reserva','FGTS','Negócio','Renda Fixa','Ações / FIIs','Cripto','Previdência','Outro']
  }
};

function getBudgetKey(ano, mes, cat) { return `${ano}-${String(mes+1).padStart(2,'0')}-${cat}`; }

function setCatOrcGroup(cat, grupo) {
  if (!S.catOrcGroup) S.catOrcGroup = {};
  if (grupo) {
    S.catOrcGroup[cat] = grupo;
  } else {
    delete S.catOrcGroup[cat];
  }
  save();
  renderCategorias();
  toast(`${cat} → ${grupo === 'essencial' ? 'Essencial' : grupo === 'estilo' ? 'Estilo de Vida' : 'Sem grupo'}`);
}

function setBudget(cat, val) {
  if (!S.budget) S.budget = {};
  const key = getBudgetKey(orcAno, orcMes, cat);
  S.budget[key] = parseFloat(val) || 0;
  save();
}

function renderOrcamento() {
  document.getElementById('orc-ano').textContent = orcAno;

  // highlight active month button
  document.querySelectorAll('[data-orcmes]').forEach(b => {
    b.classList.toggle('tf-active', parseInt(b.dataset.orcmes) === orcMes);
  });

  const mesLabel = MESES_FULL[orcMes];
  document.getElementById('orc-mes-label').textContent = mesLabel;

  // Receita sidebar
  const recMeses = document.getElementById('orc-receita-meses');
  let totalAnual = 0;
  recMeses.innerHTML = MESES.map((m, i) => {
    const ym = `${orcAno}-${String(i+1).padStart(2,'0')}`;
    const rec = S.transactions.filter(t => t.type === 'receita' && t.date.startsWith(ym)).reduce((s,t) => s + amountBrl(t), 0);
    totalAnual += rec;
    const isActive = i === orcMes;
    return `<div onclick="orcMes=${i};renderOrcamento()" style="display:flex;justify-content:space-between;padding:8px 10px;border-radius:8px;cursor:pointer;font-size:13px;${isActive ? 'background:var(--tint-indigo);color:#4f46e5;font-weight:700;' : 'color:var(--text-3);'}transition:background 0.15s;" onmouseover="this.style.background='${isActive?'var(--tint-indigo)':'var(--bg)'}'" onmouseout="this.style.background='${isActive?'var(--tint-indigo)':'transparent'}'">
      <span>${m}</span>
      <span style="font-weight:${isActive?'800':'600'};">${brl(rec)}</span>
    </div>`;
  }).join('');
  document.getElementById('orc-total-anual').textContent = brl(totalAnual);

  // Calculate month's receita for %
  const ym = `${orcAno}-${String(orcMes+1).padStart(2,'0')}`;
  const receitaMes = S.transactions.filter(t => t.type === 'receita' && t.date.startsWith(ym)).reduce((s,t) => s + amountBrl(t), 0);

  // Build table
  const tbody = document.getElementById('orc-tbody');
  let html = '';

  // Build dynamic groups from S.catOrcGroup
  const dynamicGroups = {
    'Essencial': {
      color: '#10b981', bg: 'var(--tint-green)', pct: 50,
      cats: Object.keys(S.catOrcGroup || {}).filter(c => S.catOrcGroup[c] === 'essencial').sort()
    },
    'Estilo de Vida': {
      color: '#3b82f6', bg: 'var(--tint-blue)', pct: 30,
      cats: Object.keys(S.catOrcGroup || {}).filter(c => S.catOrcGroup[c] === 'estilo').sort()
    },
    'Investimentos': ORC_GROUPS['Investimentos']
  };

  for (const [groupName, group] of Object.entries(dynamicGroups)) {
    // Group totals
    let groupPlanejado = 0, groupRealizado = 0;
    const catRows = [];

    const isInvest = groupName === 'Investimentos';
    const cats = group.cats;

    for (const cat of cats) {
      const budgetKey = getBudgetKey(orcAno, orcMes, cat);
      const planejado = S.budget?.[budgetKey] || 0;

      let realizado = 0;
      if (isInvest) {
        // Map ORC category names to S.investments tipo
        const CAT_TO_TIPO = {
          'Reserva': 'reserva', 'FGTS': 'fgts', 'Negócio': 'negocio',
          'Renda Fixa': 'renda_fixa', 'Ações / FIIs': 'acoes',
          'Cripto': 'cripto', 'Previdência': 'previdencia', 'Outro': 'outro'
        };
        const tipoAlvo = CAT_TO_TIPO[cat];
        const invs = (S.investments || []).filter(i => i && i.tipo === tipoAlvo && i.dataInicio && i.dataInicio <= ym + '-31');
        for (const inv of invs) {
          if (inv.tipo === 'negocio') {
            // Negócio: só rendimentos (dinheiro que já retornou)
            realizado += (inv.rendimentos || [])
              .filter(r => r.data && r.data <= ym + '-31')
              .reduce((s, r) => s + calcRendBRL(r, inv), 0);
          } else {
            // Renda fixa, cripto, etc: valor atual do investimento
            realizado += calcValorAtual(inv);
          }
        }
      } else {
        realizado = S.transactions.filter(t => t.type === 'despesa' && t.category === cat && t.date.startsWith(ym)).reduce((s,t) => s + amountBrl(t), 0);
      }

      groupPlanejado += planejado;
      groupRealizado += realizado;

      const ating = planejado > 0 ? ((realizado / planejado) * 100) : 0;
      const pctRec = receitaMes > 0 ? ((realizado / receitaMes) * 100) : 0;
      const atingColor = ating > 100 ? '#e11d48' : ating > 80 ? '#f59e0b' : '#10b981';

      catRows.push(`<tr class="trow" style="border-bottom:1px solid var(--bg);">
        <td style="padding:10px 16px 10px 40px;color:var(--text-2);font-weight:500;">
          <span style="color:${group.color};margin-right:6px;">●</span> ${escapeHtml(cat)}
        </td>
        <td style="padding:10px 16px;text-align:right;">
          <input type="number" value="${planejado || ''}" placeholder="R$ 0" min="0" step="0.01"
            onchange="setBudget('${cat}', this.value)"
            style="width:90px;padding:4px 8px;border:1px solid var(--border);border-radius:6px;font-size:13px;text-align:right;color:var(--text-2);">
        </td>
        <td style="padding:10px 16px;text-align:right;font-weight:600;color:${realizado > 0 ? 'var(--text)' : 'var(--border-2)'};">${brl(realizado)}</td>
        <td style="padding:10px 16px;text-align:right;font-weight:600;color:${atingColor};">${planejado > 0 ? ating.toFixed(2)+'%' : '0.00%'}</td>
        <td style="padding:10px 16px;text-align:right;font-weight:500;color:var(--text-3);">${pctRec.toFixed(2)}%</td>
      </tr>`);
    }

    const groupAting = groupPlanejado > 0 ? ((groupRealizado / groupPlanejado) * 100) : 0;
    const groupPctRec = receitaMes > 0 ? ((groupRealizado / receitaMes) * 100) : 0;

    html += `<tr style="background:${group.bg};border-bottom:1px solid ${group.color}20;">
      <td style="padding:14px 16px;font-weight:800;color:${group.color};font-size:14px;">
        <span style="margin-right:6px;">▼</span> ${groupName}
        <span style="font-size:12px;font-weight:500;color:var(--text-3);margin-left:8px;">(${group.pct}%)</span>
      </td>
      <td style="padding:14px 16px;text-align:right;font-weight:800;color:${group.color};">${brl(groupPlanejado)}</td>
      <td style="padding:14px 16px;text-align:right;font-weight:800;color:${group.color};">${brl(groupRealizado)}</td>
      <td style="padding:14px 16px;text-align:right;font-weight:800;color:${group.color};">${groupAting.toFixed(2)}%</td>
      <td style="padding:14px 16px;text-align:right;font-weight:800;color:${group.color};">${groupPctRec.toFixed(2)}%</td>
    </tr>`;
    html += catRows.join('');
  }

  tbody.innerHTML = html;
}

// ─── GRÁFICOS ─────────────────────────────────────────────────────────────────
function getDateRange(periodo) {
  const now = new Date();
  const ranges = {
    mes:  [new Date(now.getFullYear(), now.getMonth(), 1), null],
    '3m': [new Date(now.getFullYear(), now.getMonth()-2, 1), null],
    '6m': [new Date(now.getFullYear(), now.getMonth()-5, 1), null],
    ano:  [new Date(now.getFullYear(), 0, 1), null],
    tudo: [null, null]
  };
  return ranges[periodo] || [null, null];
}

function getNumMonths(periodo) {
  return { mes:1, '3m':3, '6m':6, ano:12, tudo:12 }[periodo] || 6;
}

function renderGraficos() {
  const periodo = document.getElementById('graf-periodo').value;
  const [start] = getDateRange(periodo);
  const startStr = start ? `${start.getFullYear()}-${String(start.getMonth()+1).padStart(2,'0')}-01` : null;
  const noPeriodo = t => !startStr || (t.date || '') >= startStr;
  // Despesa "real": exclui transferências, pagamento de fatura e itens que não contam
  const despReal = S.transactions.filter(t => noPeriodo(t) && t.type==='despesa' && !t.isTransfer && !isExcludedFromChart(t));

  // Estado vazio: sem despesas no período, esconde os gráficos e explica
  let emptyEl = document.getElementById('graf-empty');
  if (!emptyEl) {
    emptyEl = document.createElement('div');
    emptyEl.id = 'graf-empty'; emptyEl.className = 'empty-state';
    emptyEl.innerHTML = '<div class="icon">📈</div><p style="font-weight:600;color:var(--text-3);">Sem despesas neste período</p><p style="font-size:13px;margin-top:4px;">Troque o período acima ou lance a primeira despesa.</p><button onclick="goto(\'nova\')" style="margin-top:12px;padding:10px 18px;background:#4f46e5;color:white;border:none;border-radius:10px;font-size:13px;font-weight:600;cursor:pointer;">Nova transação</button>';
    document.getElementById('graf-periodo').closest('section').appendChild(emptyEl);
  }
  const chartCards = [...document.querySelectorAll('#page-graficos canvas')].map(c => c.closest('.card') || c.parentElement);
  const vazio = despReal.length === 0;
  emptyEl.style.display = vazio ? 'block' : 'none';
  chartCards.forEach(c => { if (c) c.style.display = vazio ? 'none' : ''; });
  if (vazio) return;

  renderDonutG(despReal);
  renderBarG(periodo);
  renderLineG(periodo);
  renderHorizG(despReal);
}

function renderDonutG(txs) {
  const by={};
  txs.forEach(t => by[t.category]=(by[t.category]||0)+amountBrl(t));
  const labels=Object.keys(by), data=Object.values(by);
  destroyChart('ch-donut-g');
  const ctx=document.getElementById('ch-donut-g');
  if (!ctx||!labels.length) return;
  charts['ch-donut-g'] = new Chart(ctx, {
    type:'doughnut',
    data:{labels,datasets:[{data,backgroundColor:COLORS,borderWidth:2,borderColor:'#fff'}]},
    options:{responsive:true,maintainAspectRatio:false,cutout:'65%',
      plugins:{legend:{position:'right',labels:{boxWidth:12,font:{size:11},padding:8}},
        tooltip:{callbacks:{label:c=>` ${c.label}: ${brl(c.raw)}`}}}}
  });
}

function renderBarG(periodo) {
  const now=new Date(), n=getNumMonths(periodo);
  const labels=[],recs=[],desps=[],invs=[];
  for(let i=n-1;i>=0;i--){
    const d=new Date(now.getFullYear(),now.getMonth()-i,1);
    const k=`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;
    labels.push(MESES[d.getMonth()]);
    const tx=S.transactions.filter(t=>t.date.startsWith(k));
    recs.push(tx.filter(t=>t.type==='receita' && !t.isTransfer).reduce((s,t)=>s+amountBrl(t),0));
    desps.push(tx.filter(t=>t.type==='despesa' && !t.isTransfer && !isExcludedFromChart(t)).reduce((s,t)=>s+amountBrl(t),0));
    invs.push(tx.filter(t=>t.type==='investimento').reduce((s,t)=>s+amountBrl(t),0));
  }
  destroyChart('ch-bar-g');
  const ctx=document.getElementById('ch-bar-g');
  if(!ctx)return;
  charts['ch-bar-g']=new Chart(ctx,{
    type:'bar',
    data:{labels,datasets:[
      {label:'Receitas',data:recs,backgroundColor:'#10b981',borderRadius:4},
      {label:'Despesas',data:desps,backgroundColor:'#f43f5e',borderRadius:4},
      {label:'Investimentos',data:invs,backgroundColor:'#6366f1',borderRadius:4}
    ]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{position:'top',labels:{boxWidth:12,font:{size:11}}},
        tooltip:{callbacks:{label:c=>` ${c.dataset.label}: ${brl(c.raw)}`}}},
      scales:{x:{grid:{display:false}},y:{grid:{color:cssVar('--chart-grid')}}}}
  });
}

function renderLineG(periodo) {
  const now=new Date(), n=getNumMonths(periodo);
  const labels=[],saldos=[];
  let acum=0;
  for(let i=n-1;i>=0;i--){
    const d=new Date(now.getFullYear(),now.getMonth()-i,1);
    const k=`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;
    labels.push(MESES[d.getMonth()]);
    const tx=S.transactions.filter(t=>t.date.startsWith(k));
    acum += tx.filter(t=>t.type==='receita' && !t.isTransfer).reduce((s,t)=>s+amountBrl(t),0)
           - tx.filter(t=>t.type==='despesa' && !t.isTransfer && !isExcludedFromChart(t)).reduce((s,t)=>s+amountBrl(t),0);
    saldos.push(acum);
  }
  destroyChart('ch-line-g');
  const ctx=document.getElementById('ch-line-g');
  if(!ctx)return;
  charts['ch-line-g']=new Chart(ctx,{
    type:'line',
    data:{labels,datasets:[{label:'Saldo Acumulado',data:saldos,
      borderColor:'#6366f1',backgroundColor:'rgba(99,102,241,0.08)',
      fill:true,tension:0.4,pointBackgroundColor:'#6366f1',pointRadius:4}]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>` ${brl(c.raw)}`}}},
      scales:{x:{grid:{display:false}},y:{grid:{color:cssVar('--chart-grid')},
        ticks:{callback:v=>'R$'+v.toLocaleString('pt-BR')}}}}
  });
}

function renderHorizG(txs) {
  const by={};
  txs.forEach(t=>by[t.category]=(by[t.category]||0)+amountBrl(t));
  const sorted=Object.entries(by).sort((a,b)=>b[1]-a[1]).slice(0,8);
  destroyChart('ch-horiz-g');
  const ctx=document.getElementById('ch-horiz-g');
  if(!ctx||!sorted.length)return;
  charts['ch-horiz-g']=new Chart(ctx,{
    type:'bar',
    data:{labels:sorted.map(([l])=>l),datasets:[{data:sorted.map(([,v])=>v),backgroundColor:COLORS,borderRadius:4}]},
    options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>` ${brl(c.raw)}`}}},
      scales:{x:{grid:{color:cssVar('--chart-grid')},ticks:{callback:v=>v===0?'R$0':'R$'+v.toLocaleString('pt-BR')}},y:{grid:{display:false}}}}
  });
}
