// FinançasCasal — 04-auth-nav.js
// Módulo carregado em ordem por index.html (escopo global compartilhado). Não use import/export.

// ─── LOGIN ────────────────────────────────────────────────────────────────────
// Mapa e-mail -> usuário do app (privado, casal)
const USER_EMAILS = {
  'almeida.paulojr08@gmail.com': 'u1',
  'thay_wosniak@yahoo.com.br': 'u2'
};
function emailToUserKey(email) {
  return USER_EMAILS[(email || '').trim().toLowerCase()] || 'u1';
}

function showLoginScreen() {
  const ls = document.getElementById('login-screen');
  if (ls) ls.style.display = 'flex';
}

function checkAuth() {
  // Visibilidade controlada por onAuthStateChanged; mostra login até o Firebase decidir.
  if (!db || !firebase.auth().currentUser) showLoginScreen();
}

function onAuthenticated(user) {
  sessionStorage.setItem('fincasal_auth', '1');
  sessionStorage.setItem('fincasal_logged_user', emailToUserKey(user.email));
  const ls = document.getElementById('login-screen');
  if (ls) ls.style.display = 'none';
  const err = document.getElementById('login-err');
  if (err) err.style.display = 'none';
  updateLeiaAquiVisibility();
  const anyVisible = document.querySelector('section[id^="page-"]:not([style*="display: none"]):not([style*="display:none"])');
  if (!anyVisible) goto('dashboard');
}

function traduzErroAuth(e) {
  const m = {
    'auth/invalid-email': 'E-mail inválido',
    'auth/weak-password': 'Senha muito curta (mínimo 6 caracteres)',
    'auth/email-already-in-use': 'Este e-mail já tem conta — use a senha correta',
    'auth/too-many-requests': 'Muitas tentativas. Aguarde alguns minutos.',
    'auth/network-request-failed': 'Sem conexão. Verifique a internet.',
    'auth/operation-not-allowed': 'Login por e-mail/senha não está ativado no Firebase.',
    'auth/invalid-login-credentials': 'E-mail ou senha incorretos'
  };
  return '❌ ' + (m[e.code] || e.message || 'Erro ao entrar');
}

async function doLogin() {
  const email = document.getElementById('login-user').value.trim();
  const pass = document.getElementById('login-pass').value;
  const err = document.getElementById('login-err');
  const showErr = (msg, color) => { err.textContent = msg; err.style.display = 'block'; err.style.color = color || '#e11d48'; };

  if (!email || !pass) { showErr('Informe e-mail e senha'); return; }
  if (!email.includes('@')) { showErr('Use seu e-mail completo (ex: nome@dominio.com)'); return; }
  if (!db) { showErr('Sem conexão com o servidor. Tente recarregar.'); return; }

  showErr('⏳ Entrando...', '#f59e0b');
  try {
    await firebase.auth().signInWithEmailAndPassword(email, pass);
    // onAuthStateChanged cuida de mostrar o app
  } catch (e) {
    // Firebase moderno (email enumeration protection) devolve o MESMO erro para
    // "senha errada" e "conta inexistente". Então: tenta criar; se já existir, foi senha.
    const ambiguo = ['auth/user-not-found','auth/invalid-credential','auth/invalid-login-credentials'].includes(e.code);
    if (ambiguo) {
      try {
        await firebase.auth().createUserWithEmailAndPassword(email, pass);
        toast('✅ Conta criada! Você já está dentro.');
      } catch (e2) {
        if (e2.code === 'auth/email-already-in-use') {
          showErr('❌ Senha incorreta');
          document.getElementById('login-pass').value = '';
        } else {
          showErr(traduzErroAuth(e2));
        }
      }
    } else if (e.code === 'auth/wrong-password') {
      showErr('❌ Senha incorreta');
      document.getElementById('login-pass').value = '';
    } else {
      showErr(traduzErroAuth(e));
    }
  }
}

async function esqueciSenha() {
  const email = document.getElementById('login-user').value.trim();
  const err = document.getElementById('login-err');
  const showErr = (msg, color) => { err.textContent = msg; err.style.display = 'block'; err.style.color = color; };
  if (!email || !email.includes('@')) { showErr('Digite seu e-mail acima primeiro', '#f59e0b'); return; }
  try {
    await firebase.auth().sendPasswordResetEmail(email);
    showErr('📧 Link de redefinição enviado para ' + email, '#059669');
  } catch (e) { showErr(traduzErroAuth(e), '#e11d48'); }
}

function logout() {
  const done = () => {
    sessionStorage.removeItem('fincasal_auth');
    sessionStorage.removeItem('fincasal_logged_user');
    document.getElementById('login-user').value = '';
    document.getElementById('login-pass').value = '';
    const err = document.getElementById('login-err'); if (err) err.style.display = 'none';
    showLoginScreen();
    updateLeiaAquiVisibility();
  };
  if (db) firebase.auth().signOut().finally(done); else done();
}

// ─── MOBILE SIDEBAR ──────────────────────────────────────────────────────────
function toggleSidebar() {
  const sidebar = document.getElementById('app-sidebar');
  const overlay = document.getElementById('sidebar-overlay');
  const isOpen = sidebar.classList.contains('sidebar-open');
  if (isOpen) {
    closeSidebar();
  } else {
    sidebar.classList.add('sidebar-open');
    overlay.style.display = 'block';
  }
}

function closeSidebar() {
  const sidebar = document.getElementById('app-sidebar');
  const overlay = document.getElementById('sidebar-overlay');
  sidebar.classList.remove('sidebar-open');
  overlay.style.display = 'none';
}

// ─── NAVIGATION ───────────────────────────────────────────────────────────────
function goto(page) {
  closeSidebar();
  document.querySelectorAll('[id^="page-"]').forEach(el => el.style.display = 'none');
  const el = document.getElementById('page-' + page);
  if (el) el.style.display = (page === 'chat') ? 'flex' : 'block';

  document.querySelectorAll('.nav-item').forEach(n => { n.classList.remove('nav-active'); n.removeAttribute('aria-current'); });
  const nav = document.querySelector(`.nav-item[data-page="${page}"]`);
  if (nav) nav.setAttribute('aria-current', 'page');
  if (nav) nav.classList.add('nav-active');

  if (page === 'dashboard') renderDashboard();
  if (page === 'historico') { histReady = false; renderHistorico(); }
  if (page === 'graficos')  renderGraficos();
  if (page === 'config')    loadConfig();
  if (page === 'chat')      initChat();
  if (page === 'contas')    renderContas();
  if (page === 'nova')      { setTodayDate(); refreshTitularSelect(); setType(currentType);
    const h = document.getElementById('f-desc-hint'); if (h) h.style.display = 'none';
    setTimeout(() => { const d = document.getElementById('f-descricao'); if (d) d.focus(); }, 50); }
  if (page === 'orcamento')  renderOrcamento();
  if (page === 'categorias') renderCategorias();
  if (page === 'dividas')    renderDividas();
  if (page === 'investimentos') renderInvestimentos();
  if (page === 'mensagens') renderMensagensPanel();
  if (page === 'leiaaqui') renderLeiaAqui();
}

// ─── HELPERS ──────────────────────────────────────────────────────────────────
function brl(v) {
  return new Intl.NumberFormat('pt-BR', { style:'currency', currency:'BRL' }).format(v||0);
}

function fmtDate(s) {
  if (!s) return '';
  const [y,m,d] = s.split('-');
  return `${d}/${m}/${y}`;
}

function toggleRecorrenteMeses() {
  const checked = document.getElementById('f-recorrente').checked;
  document.getElementById('f-recorrente-meses-row').style.display = checked ? 'block' : 'none';
}

function setTodayDate() {
  document.getElementById('f-data').value = new Date().toISOString().slice(0,10);
}

function toast(msg, dur=3000) {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.style.display = 'block';
  t.style.opacity = '1';
  clearTimeout(t._to);
  t._to = setTimeout(() => { t.style.opacity = '0'; setTimeout(()=>t.style.display='none', 300); }, dur);
}

function refreshUserSelects() {
  const { u1, u2 } = S.settings;
  const fu = document.getElementById('fil-usuario');
  if (fu) fu.innerHTML = `<option value="">Todos</option><option value="${u1}">${u1}</option><option value="${u2}">${u2}</option>`;
}

function refreshContaSelect(filterType = 'all') {
  const sel = document.getElementById('f-conta');
  if (!sel) return;
  let accounts = S.accounts;
  if (filterType === 'conta')  accounts = accounts.filter(a => a.accountType !== 'cartao');
  if (filterType === 'cartao') accounts = accounts.filter(a => a.accountType === 'cartao');

  // Filtrar por titular selecionado
  const titular = document.getElementById('f-usuario')?.value;
  if (titular) accounts = accounts.filter(a => a.owner === titular);

  if (!accounts.length) {
    const msg = filterType === 'cartao' ? '— Nenhum cartão cadastrado —' : '— Nenhuma conta cadastrada —';
    sel.innerHTML = `<option value="">${msg}</option>`;
    return;
  }
  sel.innerHTML = accounts.map(a => {
    const b = BANKS[a.bank];
    const cur = b.currency === 'USD' ? ' (USD)' : '';
    return `<option value="${a.id}">${escapeHtml(b.name)} — ${escapeHtml(a.label)}${cur}</option>`;
  }).join('');
  onContaChange();
}

function onTitularChange() {
  // Recarregar contas filtradas pelo titular selecionado
  if (currentType === 'despesa') {
    setFormaPgto(formaPgto);
  } else {
    refreshContaSelect('conta');
  }
  if (currentType === 'transferencia') {
    const titular = document.getElementById('f-usuario')?.value;
    const destSel = document.getElementById('f-conta-destino');
    let destAccounts = S.accounts.filter(a => a.accountType === 'conta');
    if (titular) destAccounts = destAccounts.filter(a => a.owner === titular);
    destSel.innerHTML = destAccounts.map(a => `<option value="${a.id}">${BANKS[a.bank]?.name || a.bank} — ${escapeHtml(a.label)}</option>`).join('');
  }
}

function onContaChange() {
  const sel = document.getElementById('f-conta');
  const acc = S.accounts.find(a => a.id === sel?.value);
  // Update currency prefix
  const isUSD = acc && BANKS[acc.bank]?.currency === 'USD';
  const prefix = document.getElementById('f-currency-prefix');
  if (prefix) prefix.textContent = isUSD ? '$' : 'R$';
}

function refreshTitularSelect() {
  const { u1, u2 } = S.settings;
  const sel = document.getElementById('f-usuario');
  if (sel) sel.innerHTML = `<option value="${u1}">${u1}</option><option value="${u2}">${u2}</option>`;
}

function destroyChart(id) {
  if (charts[id]) { charts[id].destroy(); delete charts[id]; }
}

let formaPgto = 'debito';
