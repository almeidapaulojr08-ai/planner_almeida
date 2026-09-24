// FinançasCasal — 02-state.js
// Módulo carregado em ordem por index.html (escopo global compartilhado). Não use import/export.

// ─── STATE ────────────────────────────────────────────────────────────────────
let S = {
  transactions: [],
  accounts: [],
  chatHistory: [],
  budget: {},
  catOrcGroup: {},
  debts: [],
  investments: [],
  deletedIds: [],
  loveMessages: [],
  acertos: {},          // { 'YYYY-MM': { amount, from, to, at } } — acerto de contas do casal
  customCats: null,
  settings: {
    u1: '', u2: '',
    apiKey: '', apiProvider: 'openai',
    password: '',
    usdRate: null, usdRateAt: null
  }
};
let currentType = 'receita';
let charts = {};
let deleteId = null;
let editContaId = null;
let titularFilter = 'ambos';
let dashView = 'conta';
let barChartType = 'bar';
let ccCardFilter = '';
let ccCatMode = 'cat';       // 'cat' or 'sub'
let ccCatDrillCat = null;    // categoria clicada para drill-down no crédito

// Detecta transação de pagamento de fatura (excluir dos gráficos de despesas)
function isPgtoFatura(t) {
  return (t.desc && t.desc.startsWith('Pagamento Fatura')) ||
         (t.category === 'Taxas' && t.subcategory === 'Pgto Fatura');
}

const CATS = {
  receita:      ['Salário','Freelance','Bônus','Aluguel Recebido','Dividendos','Reembolso','Outros'],
  investimento: ['Renda Fixa','Ações','FIIs','Tesouro Direto','Criptomoedas','Poupança','Previdência','Outros']
};

// Despesa categories with subcategories
const CATS_DESPESA = {
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
};

const COLORS = ['#6366f1','#f59e0b','#10b981','#f43f5e','#3b82f6','#8b5cf6','#ec4899','#14b8a6','#f97316','#84cc16','#06b6d4','#a855f7'];
const MESES  = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
const MESES_FULL = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];

const BANKS = {
  bradesco: {
    name:'Bradesco', color:'#CC0000', bg:'#FFF0F0', currency:'BRL',
    svg:`<rect width="48" height="48" rx="0" fill="#CC0000"/><text x="24" y="32" text-anchor="middle" fill="white" font-family="Arial,sans-serif" font-weight="900" font-size="20">B</text>`
  },
  nubank: {
    name:'Nubank', color:'#820AD1', bg:'#F5E6FF', currency:'BRL',
    svg:`<rect width="48" height="48" rx="0" fill="#820AD1"/><text x="24" y="32" text-anchor="middle" fill="white" font-family="Arial,sans-serif" font-weight="900" font-size="16" letter-spacing="-1">Nu</text>`
  },
  sicredi: {
    name:'Sicredi', color:'#007A3D', bg:'#E6FFF2', currency:'BRL',
    svg:`<rect width="48" height="48" rx="0" fill="#007A3D"/><text x="24" y="32" text-anchor="middle" fill="white" font-family="Arial,sans-serif" font-weight="900" font-size="14">SC</text>`
  },
  binance: {
    name:'Binance', color:'#F0B90B', bg:'#FFFBEA', currency:'USD',
    svg:`<rect width="48" height="48" rx="0" fill="#F0B90B"/><polygon points="24,10 28,14 24,18 20,14" fill="white"/><polygon points="15,19 19,15 23,19 19,23" fill="white"/><polygon points="33,19 29,15 25,19 29,23" fill="white"/><polygon points="24,20 28,24 24,28 20,24" fill="white"/><polygon points="24,30 28,26 32,30 28,34" fill="white"/><polygon points="24,30 20,26 16,30 20,34" fill="white"/>`
  },
  caixa: {
    name:'Caixa', color:'#005CA9', bg:'#E6F0FF', currency:'BRL',
    svg:`<rect width="48" height="48" rx="0" fill="#005CA9"/><path d="M24 8 L38 20 L34 20 L34 38 L14 38 L14 20 L10 20 Z" fill="#F79520"/><rect x="20" y="26" width="8" height="12" fill="#005CA9"/>`
  },
  itau: {
    name:'Itau', color:'#003399', bg:'#E6ECFF', currency:'BRL',
    svg:`<rect width="48" height="48" rx="0" fill="#003399"/><text x="24" y="32" text-anchor="middle" fill="#FF6600" font-family="Arial,sans-serif" font-weight="900" font-size="16">IT</text>`
  },
  mercado_pago: {
    name:'Mercado Pago', color:'#009EE3', bg:'#E6F7FF', currency:'BRL',
    svg:`<rect width="48" height="48" rx="0" fill="#009EE3"/><text x="24" y="32" text-anchor="middle" fill="white" font-family="Arial,sans-serif" font-weight="900" font-size="14">MP</text>`
  },
};

// ─── PERSISTENCE ──────────────────────────────────────────────────────────────
function load() {
  try {
    const d = JSON.parse(localStorage.getItem('fincasal_v2') || '{}');
    if (d.transactions) S.transactions = d.transactions;
    if (d.accounts)     S.accounts     = d.accounts;
    if (d.chatHistory)  S.chatHistory  = d.chatHistory;
    if (d.budget)       S.budget       = d.budget;
    if (d.catOrcGroup)  S.catOrcGroup  = d.catOrcGroup;
    if (d.debts)        S.debts        = d.debts;
    if (d.investments)  S.investments  = d.investments;
    if (d.deletedIds)   S.deletedIds   = d.deletedIds;
    if (d.loveMessages) S.loveMessages = d.loveMessages;
    if (d.acertos)      S.acertos      = d.acertos;
    if (d.customCats)   S.customCats   = d.customCats;
    if (d.customBanks)  { S.customBanks = d.customBanks; Object.assign(BANKS, d.customBanks); }
    if (d.settings)     S.settings     = { ...S.settings, ...d.settings };
  } catch(e) {}
}

function save() {
  localStorage.setItem('fincasal_v2', JSON.stringify(S));
  // Só envia pro Firebase se ele já carregou pelo menos 1x (evita sobrescrever dados reais
  // com defaults). Manda só o que mudou, por item — ver pushToFirebase().
  pushToFirebase();
}

// ─── USD RATE ─────────────────────────────────────────────────────────────────
async function fetchUsdRate() {
  const ONE_HOUR = 60 * 60 * 1000;
  const now = Date.now();
  if (S.settings.usdRate && S.settings.usdRateAt && (now - S.settings.usdRateAt) < ONE_HOUR) return;
  try {
    const r = await fetch('https://open.er-api.com/v6/latest/USD');
    const d = await r.json();
    S.settings.usdRate   = d.rates.BRL;
    S.settings.usdRateAt = now;
    save();
  } catch(e) {
    if (!S.settings.usdRate) S.settings.usdRate = 5.7; // fallback
  }
}

function usdToBrl(amount) {
  return amount * (S.settings.usdRate || 5.7);
}
