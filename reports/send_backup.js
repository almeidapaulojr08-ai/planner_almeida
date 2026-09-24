// ─────────────────────────────────────────────────────────────────────────────
// Backup semanal do FinançasCasal: exporta o banco inteiro (mesmo formato do botão
// "Exportar JSON" do app, importável pelo "Importar JSON") e manda como arquivo no
// Telegram de cada pessoa. Roda no GitHub Actions logo após o relatório.
//
// Variáveis de ambiente: FIREBASE_SERVICE_ACCOUNT, FIREBASE_DB_URL,
//                        TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_IDS (pares id:Nome ou só id)
// ─────────────────────────────────────────────────────────────────────────────
const admin = require('firebase-admin');

function fbToArray(obj) {
  if (!obj) return [];
  if (Array.isArray(obj)) return obj.filter(Boolean);
  return Object.values(obj);
}

async function main() {
  const svc = JSON.parse(process.env.FIREBASE_SERVICE_ACCOUNT);
  admin.initializeApp({ credential: admin.credential.cert(svc), databaseURL: process.env.FIREBASE_DB_URL });
  const data = (await admin.database().ref('data').once('value')).val() || {};

  // Mesmo formato do exportarDados() do app (coleções como arrays; sem chave de API)
  const exportData = {
    version: 2,
    exportedAt: new Date().toISOString(),
    source: 'backup-automatico',
    transactions: fbToArray(data.transactions),
    accounts: fbToArray(data.accounts),
    debts: fbToArray(data.debts),
    investments: fbToArray(data.investments),
    budget: data.budget || {},
    catOrcGroup: data.catOrcGroup || {},
    chatHistory: fbToArray(data.chatHistory),
    customCats: data.customCats || null,
    customBanks: data.customBanks || null,
    deletedIds: fbToArray(data.deletedIds),
    loveMessages: fbToArray(data.loveMessages),
    settings: { u1: (data.settings || {}).u1, u2: (data.settings || {}).u2 }
  };
  const json = JSON.stringify(exportData);
  const dia = exportData.exportedAt.slice(0, 10);
  const nome = `financas_casal_backup_${dia}.json`;
  const kb = Math.round(json.length / 1024);
  const resumo = `🗄️ Backup semanal — ${dia.split('-').reverse().join('/')}\n` +
    `${exportData.transactions.length} transações · ${exportData.accounts.length} contas · ${exportData.debts.length} dívidas · ${exportData.investments.length} investimentos · ${exportData.loveMessages.length} mensagens (${kb} KB)\n` +
    `Pra restaurar: Configurações → Importar JSON. Guarde este arquivo.`;

  const token = process.env.TELEGRAM_BOT_TOKEN;
  const ids = (process.env.TELEGRAM_CHAT_IDS || '').split(',').map(s => s.trim()).filter(Boolean).map(e => e.split(':')[0].trim());
  if (!token || !ids.length) throw new Error('TELEGRAM_BOT_TOKEN ou TELEGRAM_CHAT_IDS ausentes');

  for (const chatId of ids) {
    const form = new FormData();
    form.append('chat_id', chatId);
    form.append('caption', resumo);
    form.append('document', new Blob([json], { type: 'application/json' }), nome);
    const res = await fetch(`https://api.telegram.org/bot${token}/sendDocument`, { method: 'POST', body: form });
    const body = await res.json();
    if (!body.ok) console.error(`Falha ao enviar backup para ${chatId}:`, body.description);
    else console.log(`Backup enviado para ${chatId} (${nome}, ${kb} KB)`);
  }
  process.exit(0);
}

main().catch(e => { console.error('Erro no backup:', e.message); process.exit(1); });
