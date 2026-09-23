// ─────────────────────────────────────────────────────────────────────────────
// Relatório financeiro do FinançasCasal via Telegram.
// Roda no GitHub Actions (cron). Lê o Firebase com service account (admin,
// ignora as regras), calcula um resumo e envia pro Telegram de cada um.
//
// Variáveis de ambiente (definidas como GitHub Secrets):
//   FIREBASE_SERVICE_ACCOUNT  → JSON da chave de service account (string)
//   FIREBASE_DB_URL           → https://almeida-wosniak-dre-default-rtdb.firebaseio.com
//   TELEGRAM_BOT_TOKEN        → token do @BotFather
//   TELEGRAM_CHAT_IDS         → chat ids separados por vírgula (Paulo,Thayse)
// ─────────────────────────────────────────────────────────────────────────────
const admin = require('firebase-admin');

const MESES = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];

function brl(v) {
  return 'R$ ' + (Math.round((v || 0) * 100) / 100).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
function fbToArray(obj) {
  if (!obj) return [];
  if (Array.isArray(obj)) return obj.filter(Boolean);
  return Object.values(obj);
}

async function main() {
  // ── conecta no Firebase ──
  const svc = JSON.parse(process.env.FIREBASE_SERVICE_ACCOUNT);
  admin.initializeApp({
    credential: admin.credential.cert(svc),
    databaseURL: process.env.FIREBASE_DB_URL
  });
  const snap = await admin.database().ref('data').once('value');
  const data = snap.val() || {};

  const txs = fbToArray(data.transactions);
  const accounts = fbToArray(data.accounts);
  const investments = fbToArray(data.investments);
  const debts = fbToArray(data.debts);
  const settings = data.settings || {};
  const budget = data.budget || {};
  const catOrcGroup = data.catOrcGroup || {};
  const usdRate = settings.usdRate || 5.7;

  const amountBrl = t => {
    const base = t.currency === 'USD' ? (t.amount || 0) * usdRate : (t.amount || 0);
    return t.isNegative ? -base : base;
  };
  const isPgtoFatura = t => (t.desc && t.desc.startsWith('Pagamento Fatura')) || (t.category === 'Taxas' && t.subcategory === 'Pgto Fatura');
  const isExcl = t => {
    if (isPgtoFatura(t)) return true;
    const sub = (t.subcategory || '').toLowerCase(), cat = (t.category || '').toLowerCase();
    if (['investimento','investimentos','emprestimo','empréstimo','emprestimos','empréstimos','ajuste saldo','conciliacao','conciliação'].includes(sub)) return true;
    if (cat === 'taxas' && sub === 'banco') return true;
    return false;
  };
  const gastoReal = t => t.type === 'despesa' && !t.isTransfer && !isExcl(t);
  const bankName = a => (a && a.bank ? a.bank.charAt(0).toUpperCase() + a.bank.slice(1).replace('_', ' ') : (a && a.label) || 'Conta');

  // faturaRef: usa o campo se existir, senão calcula pelo fechamento do cartão
  const cardById = {}; accounts.forEach(a => { if (a.accountType === 'cartao') cardById[a.id] = a; });
  const faturaRefDe = t => {
    if (t.faturaRef) return t.faturaRef;
    const c = cardById[t.accountId];
    if (!c || !c.fecha || !t.date) return (t.date || '').slice(0, 7);
    const fecha = parseInt(c.fecha) || 1;
    const [y, m, d] = t.date.split('-').map(Number);
    const ym = y * 12 + (m - 1) + (d > fecha ? 1 : 0);
    return `${Math.floor(ym / 12)}-${String(ym % 12 + 1).padStart(2, '0')}`;
  };

  const hoje = new Date();
  const ym = `${hoje.getFullYear()}-${String(hoje.getMonth() + 1).padStart(2, '0')}`;
  const dPrev = new Date(hoje.getFullYear(), hoje.getMonth() - 1, 1);
  const ymPrev = `${dPrev.getFullYear()}-${String(dPrev.getMonth() + 1).padStart(2, '0')}`;
  const diaDoMes = hoje.getDate();
  const fechamentoMes = diaDoMes <= 3; // dia 1-3 = relatório de fechamento do mês anterior

  const L = []; // linhas da mensagem
  L.push(`<b>💰 FinançasCasal — Resumo</b>`);
  L.push(`<i>${fechamentoMes ? 'Fechamento de ' + MESES[dPrev.getMonth()] : MESES[hoje.getMonth()] + ' (parcial, dia ' + diaDoMes + ')'}</i>`);
  L.push('');

  const alvoYM = fechamentoMes ? ymPrev : ym;
  const compYM = fechamentoMes ? `${new Date(dPrev.getFullYear(), dPrev.getMonth() - 1, 1).getFullYear()}-${String(new Date(dPrev.getFullYear(), dPrev.getMonth() - 1, 1).getMonth() + 1).padStart(2, '0')}` : ymPrev;

  // ── Saldo por conta (acumulado até hoje) ──
  L.push(`<b>🏦 Saldos por conta</b>`);
  let saldoTotal = 0;
  accounts.filter(a => a.accountType !== 'cartao').forEach(a => {
    const acc = txs.filter(t => t.accountId === a.id && (t.date || '') <= ym + '-31');
    const rec = acc.filter(t => t.type === 'receita').reduce((s, t) => s + amountBrl(t), 0);
    const desp = acc.filter(t => t.type === 'despesa' && t.formaPgto !== 'credito' && t.pago !== false).reduce((s, t) => s + amountBrl(t), 0);
    const bal = rec - desp;
    saldoTotal += bal;
    L.push(`• ${a.label || bankName(a)}: <b>${brl(bal)}</b>`);
  });
  L.push(`<b>Total: ${brl(saldoTotal)}</b>`);
  L.push('');

  // ── Gasto do mês vs anterior ──
  const gastoAlvo = txs.filter(t => (t.date || '').startsWith(alvoYM) && gastoReal(t)).reduce((s, t) => s + amountBrl(t), 0);
  const gastoComp = txs.filter(t => (t.date || '').startsWith(compYM) && gastoReal(t)).reduce((s, t) => s + amountBrl(t), 0);
  const recAlvo = txs.filter(t => (t.date || '').startsWith(alvoYM) && t.type === 'receita' && !t.isTransfer).reduce((s, t) => s + amountBrl(t), 0);
  L.push(`<b>📊 Movimento do mês</b>`);
  L.push(`• Receitas: <b>${brl(recAlvo)}</b>`);
  let difLine = `• Despesas: <b>${brl(gastoAlvo)}</b>`;
  if (gastoComp > 0) { const dif = (gastoAlvo - gastoComp) / gastoComp * 100; difLine += ` (${dif >= 0 ? '📈 +' : '📉 '}${dif.toFixed(0)}% vs mês anterior)`; }
  L.push(difLine);
  L.push(`• Saldo do mês: <b>${brl(recAlvo - gastoAlvo)}</b>`);
  L.push('');

  // ── Top 5 categorias do mês ──
  const porCat = {};
  txs.filter(t => (t.date || '').startsWith(alvoYM) && gastoReal(t)).forEach(t => porCat[t.category] = (porCat[t.category] || 0) + amountBrl(t));
  const top = Object.entries(porCat).sort((a, b) => b[1] - a[1]).slice(0, 5);
  if (top.length) {
    L.push(`<b>🏷️ Top categorias</b>`);
    top.forEach(([c, v]) => L.push(`• ${c}: ${brl(v)}`));
    L.push('');
  }

  // ── Próximas faturas a vencer ──
  const fatYM = (() => { let m = hoje.getMonth(), y = hoje.getFullYear(); const minFecha = Math.min(...accounts.filter(a => a.accountType === 'cartao').map(c => parseInt(c.fecha) || 1)); if (diaDoMes > minFecha) m += 1; if (m > 11) { m -= 12; y++; } return `${y}-${String(m + 1).padStart(2, '0')}`; })();
  const fats = [];
  accounts.filter(a => a.accountType === 'cartao').forEach(c => {
    const total = txs.filter(t => t.type === 'despesa' && t.formaPgto === 'credito' && faturaRefDe(t) === fatYM && t.accountId === c.id).reduce((s, t) => s + amountBrl(t), 0);
    if (total <= 0) return;
    const venceDia = parseInt(c.vence) || 10;
    let venc = new Date(hoje.getFullYear(), hoje.getMonth(), venceDia);
    if (venc < hoje) venc = new Date(hoje.getFullYear(), hoje.getMonth() + 1, venceDia);
    const dias = Math.ceil((venc - hoje) / 86400000);
    fats.push({ nome: c.label || bankName(c), total, dias });
  });
  if (fats.length) {
    fats.sort((a, b) => a.dias - b.dias);
    L.push(`<b>💳 Próximas faturas</b>`);
    fats.forEach(f => L.push(`• ${f.nome}: <b>${brl(f.total)}</b> — vence em ${f.dias} dia${f.dias !== 1 ? 's' : ''}`));
    L.push('');
  }

  // ── Parcelas terminando neste mês ──
  const term = txs.filter(t => {
    if (!(t.date || '').startsWith(alvoYM)) return false;
    const m = /^(\d+)\/(\d+)$/.exec(t.parcela || '');
    return m && m[1] === m[2] && parseInt(m[2]) > 1;
  });
  if (term.length) {
    const val = term.reduce((s, t) => s + amountBrl(t), 0);
    L.push(`🎉 <b>${term.length} parcela(s) terminando</b> — alívio de ${brl(val)}/mês`);
    L.push('');
  }

  // ── Orçamento estourado ──
  const est = [];
  Object.keys(catOrcGroup).forEach(cat => {
    const key = `${alvoYM}-${cat}`; // mesmo formato do app: YYYY-MM-categoria
    const plan = budget[key] || 0;
    if (plan <= 0) return;
    const real = txs.filter(t => t.type === 'despesa' && t.category === cat && (t.date || '').startsWith(alvoYM)).reduce((s, t) => s + amountBrl(t), 0);
    if (real > plan) est.push(`• ${cat}: ${brl(real)} (planejado ${brl(plan)})`);
  });
  if (est.length) { L.push(`⚠️ <b>Orçamento estourado</b>`); est.forEach(e => L.push(e)); L.push(''); }

  // ── Patrimônio (investimentos - dívidas) ──
  const invTotal = investments.filter(i => i.tipo !== 'negocio').reduce((s, i) => s + (i.valorAtual || i.valorInvestido || 0), 0);
  const dividaTotal = debts.filter(d => d.status !== 'paga').reduce((s, d) => {
    const pago = (d.pagamentos || []).reduce((x, p) => x + (p.amount || 0), 0) + (d.amortizacoesExtra || []).reduce((x, a) => x + (a.amount || 0), 0);
    return s + Math.max(0, (d.total || d.principal || 0) - pago);
  }, 0);
  L.push(`<b>📈 Patrimônio</b>`);
  L.push(`• Investido: ${brl(invTotal)}`);
  if (dividaTotal > 0) L.push(`• Dívidas em aberto: ${brl(dividaTotal)}`);
  L.push(`• Líquido (contas + invest − dívidas): <b>${brl(saldoTotal + invTotal - dividaTotal)}</b>`);

  const mensagem = L.join('\n');

  // ── envia pro Telegram ──
  const token = process.env.TELEGRAM_BOT_TOKEN;
  const chatIds = (process.env.TELEGRAM_CHAT_IDS || '').split(',').map(s => s.trim()).filter(Boolean);
  if (!token || !chatIds.length) throw new Error('TELEGRAM_BOT_TOKEN ou TELEGRAM_CHAT_IDS ausentes');

  for (const chatId of chatIds) {
    const res = await fetch(`https://api.telegram.org/bot${token}/sendMessage`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ chat_id: chatId, text: mensagem, parse_mode: 'HTML', disable_web_page_preview: true })
    });
    const body = await res.json();
    if (!body.ok) console.error(`Falha ao enviar para ${chatId}:`, body.description);
    else console.log(`Enviado para ${chatId}`);
  }
  console.log('Relatório concluído.');
  process.exit(0);
}

main().catch(e => { console.error('Erro no relatório:', e.message); process.exit(1); });
