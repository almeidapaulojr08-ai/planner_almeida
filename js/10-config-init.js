// FinançasCasal — 10-config-init.js
// Módulo carregado em ordem por index.html (escopo global compartilhado). Não use import/export.

// ─── CONFIG ───────────────────────────────────────────────────────────────────
function loadConfig() {
  document.getElementById('cfg-u1').value=S.settings.u1;
  document.getElementById('cfg-u2').value=S.settings.u2;
  document.getElementById('cfg-apikey').value=S.settings.apiKey;
  document.getElementById('cfg-password').value='';
  document.getElementById('cfg-password2').value='';
  setProvider(S.settings.apiProvider);
}

function setProvider(p) {
  S.settings.apiProvider=p;
  document.getElementById('pbtn-openai').className  = 'provider-btn'+(p==='openai'?' active':'');
  document.getElementById('pbtn-anthropic').className = 'provider-btn'+(p==='anthropic'?' active':'');
  const row = document.getElementById('cfg-model-row');
  const sel = document.getElementById('cfg-model');
  if (row && sel) {
    row.style.display = p === 'anthropic' ? 'block' : 'none';
    sel.innerHTML = AI_MODELS.map(m => `<option value="${m.id}">${m.label}</option>`).join('');
    sel.value = AI_MODELS.some(m => m.id === S.settings.aiModel) ? S.settings.aiModel : AI_DEFAULT_MODEL;
  }
}

function salvarConfig() {
  S.settings.u1=document.getElementById('cfg-u1').value.trim()||S.settings.u1;
  S.settings.u2=document.getElementById('cfg-u2').value.trim()||S.settings.u2;
  S.settings.apiKey=document.getElementById('cfg-apikey').value.trim();
  const mSel=document.getElementById('cfg-model'); if (mSel && mSel.value) S.settings.aiModel=mSel.value;
  const p1=document.getElementById('cfg-password').value;
  const p2=document.getElementById('cfg-password2').value;
  if (p1 || p2) {
    if (p1 !== p2) { toast('❌ As senhas não coincidem!'); return; }
    if (p1.length < 6) { toast('❌ Senha deve ter pelo menos 6 caracteres'); return; }
    const u = db ? firebase.auth().currentUser : null;
    if (!u) { toast('❌ Faça login novamente para trocar a senha'); return; }
    u.updatePassword(p1).then(() => {
      document.getElementById('cfg-password').value='';
      document.getElementById('cfg-password2').value='';
      toast('✅ Senha alterada!');
    }).catch(e => {
      if (e.code === 'auth/requires-recent-login') toast('⚠️ Por segurança, saia e entre de novo antes de trocar a senha.');
      else toast('❌ ' + (e.message || 'Erro ao trocar senha'));
    });
  }
  save();
  document.getElementById('header-users').textContent=`${S.settings.u1} & ${S.settings.u2}`;
  refreshUserSelects();
  updateAIStatus();
  toast('✅ Configurações salvas!');
}

function exportarDados() {
  const exportData = {
    version: 2,
    exportedAt: new Date().toISOString(),
    transactions: S.transactions,
    accounts: S.accounts,
    debts: S.debts,
    investments: S.investments,
    budget: S.budget,
    catOrcGroup: S.catOrcGroup,
    chatHistory: S.chatHistory,
    customCats: S.customCats,
    customBanks: S.customBanks,
    deletedIds: S.deletedIds,
    loveMessages: S.loveMessages || [],
    acertos: S.acertos || {},
    settings: { u1: S.settings.u1, u2: S.settings.u2 }
  };
  const blob = new Blob([JSON.stringify(exportData, null, 2)], {type:'application/json'});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `financas_casal_backup_${new Date().toISOString().slice(0,10)}.json`;
  a.click();
  URL.revokeObjectURL(a.href);
  toast('📤 Backup completo exportado!');
}

function importarDados(e) {
  const file = e.target.files[0];
  if (!file) return;
  const r = new FileReader();
  r.onload = ev => {
    try {
      const d = JSON.parse(ev.target.result);
      const tombstones = mergeTombstones(S.deletedIds, d.deletedIds || []);
      // Merge inteligente por ID (não substitui, combina)
      if (d.transactions) S.transactions = mergeArrayById(S.transactions, d.transactions, tombstones);
      if (d.accounts)     S.accounts     = mergeArrayById(S.accounts, d.accounts, tombstones);
      if (d.debts)        S.debts        = mergeArrayById(S.debts, d.debts, tombstones);
      if (d.investments)  S.investments  = mergeArrayById(S.investments, d.investments, tombstones);
      if (d.loveMessages) S.loveMessages = mergeArrayById(S.loveMessages || [], d.loveMessages, tombstones);
      if (d.acertos)      S.acertos      = { ...(S.acertos || {}), ...d.acertos };
      if (d.deletedIds)   S.deletedIds   = tombstones;
      if (d.budget)       S.budget       = { ...(S.budget || {}), ...d.budget };
      if (d.catOrcGroup)  S.catOrcGroup  = { ...(S.catOrcGroup || {}), ...d.catOrcGroup };
      if (d.customCats && !S.customCats) S.customCats = d.customCats;
      if (d.customBanks)  { S.customBanks = { ...(S.customBanks||{}), ...d.customBanks }; Object.assign(BANKS, d.customBanks); }
      if (d.settings)     S.settings     = { ...S.settings, ...d.settings };
      save();
      toast('📥 Backup importado com sucesso!');
      goto('dashboard');
    } catch { toast('❌ Arquivo inválido'); }
  };
  r.readAsText(file);
  e.target.value = '';
}

function apagarTudo() {
  if(confirm('⚠️ Tem certeza? Todos os dados serão apagados permanentemente. Esta ação não pode ser desfeita.')) {
    S.transactions=[];
    S.chatHistory=[];
    save();
    toast('🗑️ Todos os dados foram apagados');
    goto('dashboard');
  }
}

// ─── INIT ─────────────────────────────────────────────────────────────────────
function seedAccounts() {
  // Use old timestamp so Firebase data ALWAYS wins in merge
  const _seed = '2020-01-01T00:00:00.000Z';
  const defaults = [
    { id: 'nubank_conta',  bank:'nubank',  label:'Conta Nubank',   owner: S.settings.u1, accountType:'conta', tipoConta:'corrente', updatedAt: _seed },
    { id: 'sicredi_conta', bank:'sicredi', label:'Conta Sicredi',  owner: S.settings.u1, accountType:'conta', tipoConta:'corrente', updatedAt: _seed },
    { id: 'card_nubank',   bank:'nubank',  label:'Cartão Nubank',  owner: S.settings.u1, accountType:'cartao', tipoConta:null, limite:'17500', fecha:'2', vence:'10', updatedAt: _seed },
    { id: 'card_sicredi',  bank:'sicredi', label:'Cartão Sicredi', owner: S.settings.u1, accountType:'cartao', tipoConta:null, limite:'13000', fecha:'2', vence:'15', updatedAt: _seed },
  ];
  let changed = false;
  // Only add missing accounts, NEVER remove or deduplicate existing ones
  // Check by bank+accountType to avoid duplicates when user already created one manually
  for (const a of defaults) {
    const exists = S.accounts.find(x => x.id === a.id || (x.bank === a.bank && x.accountType === a.accountType));
    if (!exists) { S.accounts.push(a); changed = true; }
  }
  if (changed) save();
}

function importExcelData() {
  // Importação automática DESATIVADA permanentemente.
  // Os dados já foram importados. Qualquer nova importação deve ser feita
  // manualmente pelo Assistente IA ou pelo modal de Import.
  return;
}
