// FinançasCasal — 06-dashboard.js
// Módulo carregado em ordem por index.html (escopo global compartilhado). Não use import/export.

// ─── DASHBOARD ────────────────────────────────────────────────────────────────
function amountBrl(tx) {
  const base = tx.currency === 'USD' ? usdToBrl(tx.amount) : tx.amount;
  return tx.isNegative ? -base : base;
}

// Returns the raw amount in original currency (respecting isNegative but NOT converting USD)
function amountRaw(tx) {
  return tx.isNegative ? -tx.amount : tx.amount;
}

// Soma meses a uma data sem estourar o dia (31/01 + 1 mês = 28/02, não 03/03).
function addMesesSeguro(base, n) {
  const y = base.getFullYear(), m = base.getMonth(), d = base.getDate();
  const ultimoDia = new Date(y, m + n + 1, 0).getDate();
  return new Date(y, m + n, Math.min(d, ultimoDia), 12, 0, 0);
}

// Returns the fatura reference month (YYYY-MM) for a credit transaction.
// If the transaction has faturaRef (from Excel import), use it directly.
// Otherwise, calculate based on the card's closing day (fecha).
// Soma N meses a 'YYYY-MM-DD' travando no último dia do mês (31/01 + 1 → 28/02, não 03/03)
function addMesesData(dateStr, n) {
  const [y, m, d] = dateStr.split('-').map(Number);
  const ym = y * 12 + (m - 1) + n;
  const ny = Math.floor(ym / 12), nm = ym % 12;
  const ultimoDia = new Date(ny, nm + 1, 0).getDate();
  return `${ny}-${String(nm + 1).padStart(2, '0')}-${String(Math.min(d, ultimoDia)).padStart(2, '0')}`;
}

function getTxFaturaRef(tx) {
  if (tx.faturaRef) return tx.faturaRef;
  // Fallback: calculate from date + card closing day
  const card = S.accounts.find(a => a.id === tx.accountId && a.accountType === 'cartao');
  if (!card || !card.fecha) return tx.date.substring(0, 7);
  const fechaDia = parseInt(card.fecha);
  const [y, m, day] = tx.date.split('-').map(Number);
  // If purchase is after closing day, it falls into next month's fatura
  // e.g. March 28, fecha=2 → cycle Mar 3–Apr 2 → fatura de Abril (month+1)
  // If purchase is on or before closing day, it falls into current month's fatura
  // e.g. April 1, fecha=2 → cycle Mar 3–Apr 2 → fatura de Abril (same month)
  // NÃO usar Date.setMonth aqui: 31/08 + 1 mês vira 01/10 (setembro não tem dia 31)
  // e a transação sumia da fatura correta. Aritmética de mês inteiro é segura.
  const ym = y * 12 + (m - 1) + (day > fechaDia ? 1 : 0);
  return `${Math.floor(ym / 12)}-${String(ym % 12 + 1).padStart(2, '0')}`;
}

// Filter credit transactions that belong to a specific fatura month (YYYY-MM)
function getCreditTxByFatura(faturaYM, cardId) {
  return S.transactions.filter(t =>
    t.type === 'despesa' && t.formaPgto === 'credito' &&
    getTxFaturaRef(t) === faturaYM &&
    (!cardId || t.accountId === cardId)
  );
}

// Get the current fatura month label: the next bill to be paid.
// Today is in the billing cycle that closes on fecha day of this/next month.
function getCurrentFaturaYM() {
  const now = new Date();
  const cartoes = S.accounts.filter(a => a.accountType === 'cartao');
  const minFecha = Math.min(...cartoes.map(c => parseInt(c.fecha) || 1));
  const day = now.getDate();
  let m = now.getMonth(); // 0-based
  let y = now.getFullYear();
  // e.g. fecha=2, today Apr 3 (after closing) → cycle Apr 3–May 2 → fatura de Maio (m+1)
  // e.g. fecha=2, today Apr 1 (before closing) → cycle Mar 3–Apr 2 → fatura de Abril (m)
  if (day > minFecha) {
    m += 1;
  }
  if (m > 11) { m -= 12; y++; }
  return `${y}-${String(m + 1).padStart(2, '0')}`;
}

// ─── DASHBOARD VIEW TOGGLE ────────────────────────────────────────────────────
function setDashView(view) {
  dashView = view;
  const isConta = view === 'conta';
  document.getElementById('dv-conta').className   = 'tf-btn' + (isConta  ? ' tf-active' : '');
  document.getElementById('dv-credito').className  = 'tf-btn' + (!isConta ? ' tf-active' : '');
  document.getElementById('dash-view-conta').style.display   = isConta ? 'block' : 'none';
  document.getElementById('dash-view-credito').style.display = isConta ? 'none'  : 'block';
  renderDashboard();
}

function setBarType(type) {
  barChartType = type;
  ['area','line','bar'].forEach(t => {
    const btn = document.getElementById('bar-type-' + t);
    if (btn) {
      btn.style.background = t === type ? 'var(--tint-indigo)' : 'var(--surface)';
      btn.style.color = t === type ? '#4f46e5' : 'var(--text-3)';
    }
  });
  renderBarDash();
}

let dashMonth = new Date().getMonth();   // 0-based
let dashYear  = new Date().getFullYear();
let dashPeriodMode = 'mes'; // 'mes' or 'ano'

function setDashPeriodMode(mode) {
  dashPeriodMode = mode;
  document.getElementById('dp-mes').className = 'tf-btn' + (mode === 'mes' ? ' tf-active' : '');
  document.getElementById('dp-ano').className = 'tf-btn' + (mode === 'ano' ? ' tf-active' : '');
  document.getElementById('dash-month').style.display = mode === 'mes' ? '' : 'none';
  renderDashboard();
}

function initDashPeriod() {
  const mSel = document.getElementById('dash-month');
  const ySel = document.getElementById('dash-year');
  mSel.innerHTML = MESES_FULL.map((m, i) => `<option value="${i}" ${i === dashMonth ? 'selected' : ''}>${m}</option>`).join('');
  // Collect years from transactions
  const years = new Set();
  S.transactions.forEach(t => { if (t.date) years.add(parseInt(t.date.substring(0, 4))); });
  years.add(new Date().getFullYear());
  ySel.innerHTML = [...years].sort().map(y => `<option value="${y}" ${y === dashYear ? 'selected' : ''}>${y}</option>`).join('');
}

function onDashPeriodChange() {
  dashMonth = parseInt(document.getElementById('dash-month').value);
  dashYear  = parseInt(document.getElementById('dash-year').value);
  renderDashboard();
}

function dashNavMonth(dir) {
  if (dashPeriodMode === 'ano') {
    dashYear += dir;
  } else {
    dashMonth += dir;
    if (dashMonth > 11) { dashMonth = 0; dashYear++; }
    if (dashMonth < 0)  { dashMonth = 11; dashYear--; }
  }
  document.getElementById('dash-month').value = dashMonth;
  document.getElementById('dash-year').value = dashYear;
  // Add year if not in select
  const ySel = document.getElementById('dash-year');
  if (!ySel.querySelector(`option[value="${dashYear}"]`)) {
    ySel.innerHTML += `<option value="${dashYear}">${dashYear}</option>`;
    ySel.value = dashYear;
  }
  renderDashboard();
}

// ─── TEMPO JUNTOS ────────────────────────────────────────────────────────────
function calcTempoDesde(inicio) {
  const hoje = new Date();
  let anos = hoje.getFullYear() - inicio.getFullYear();
  let meses = hoje.getMonth() - inicio.getMonth();
  let dias = hoje.getDate() - inicio.getDate();
  if (dias < 0) {
    meses--;
    const mesAnterior = new Date(hoje.getFullYear(), hoje.getMonth(), 0);
    dias += mesAnterior.getDate();
  }
  if (meses < 0) {
    anos--;
    meses += 12;
  }
  const partes = [];
  if (anos > 0) partes.push(anos === 1 ? '1 ano' : anos + ' anos');
  if (meses > 0) partes.push(meses === 1 ? '1 mês' : meses + ' meses');
  partes.push(dias === 1 ? '1 dia' : dias + ' dias');
  return partes.join(', ');
}

function calcTempoJuntos() {
  return calcTempoDesde(new Date(2025, 5, 18)); // 18/06/2025 (month is 0-indexed)
}

function calcTempoCasados() {
  return calcTempoDesde(new Date(2026, 5, 12)); // 12/06/2026 (month is 0-indexed)
}

function updateTempoJuntos() {
  const elJ = document.getElementById('tempo-juntos-valor');
  if (elJ) elJ.textContent = calcTempoJuntos();
  const elC = document.getElementById('tempo-casados-valor');
  if (elC) elC.textContent = calcTempoCasados();
}

// ─── MENSAGENS DE AMOR - SISTEMA BIDIRECIONAL ──────────────────────────────

const LOVE_EMOJIS = ['💕','❤️','💗','💖','🥰','😍','💌','🌹','✨','💍','🏠','💋','🫶','😘','💝','🦋','🌙','⭐','🔥','💐','🍫','☕','🎵','🌈','💎'];

function getLoggedUserKey() {
  let loggedUser = sessionStorage.getItem('fincasal_logged_user');
  if (!loggedUser && sessionStorage.getItem('fincasal_auth')) {
    const loginField = document.getElementById('login-user');
    if (loginField && loginField.value.trim()) {
      const u = loginField.value.trim().toLowerCase();
      loggedUser = (u === S.settings.u2.toLowerCase()) ? 'u2' : 'u1';
      sessionStorage.setItem('fincasal_logged_user', loggedUser);
    }
  }
  return loggedUser || 'u1';
}

function getReceivedMessages() {
  const me = getLoggedUserKey();
  const now = new Date();
  return (S.loveMessages || []).filter(m => {
    if (m.para !== me) return false;
    const inicio = new Date(m.inicio);
    const expira = new Date(m.expira);
    return now >= inicio && now <= expira;
  });
}

function updateLeiaAquiVisibility() {
  const navLeia = document.getElementById('nav-leiaaqui');
  const navMsg = document.getElementById('nav-mensagens');
  const badge = document.getElementById('badge-leiaaqui');
  const loggedUser = getLoggedUserKey();
  const isLoggedIn = sessionStorage.getItem('fincasal_auth');

  // "Enviar Mensagem" aparece para quem está logado (u1 ou u2)
  if (navMsg) {
    navMsg.style.display = (isLoggedIn && (loggedUser === 'u1' || loggedUser === 'u2')) ? 'flex' : 'none';
  }

  // "Leia aqui" aparece se tem mensagens ativas para o usuário logado
  const received = getReceivedMessages();
  if (navLeia) {
    navLeia.style.display = (isLoggedIn && received.length > 0) ? 'flex' : 'none';
  }
  if (badge) {
    if (received.length > 0) {
      badge.style.display = 'inline';
      badge.textContent = received.length;
    } else {
      badge.style.display = 'none';
    }
  }
}

function toggleEmojiPicker() {
  const picker = document.getElementById('emoji-picker');
  picker.style.display = picker.style.display === 'none' ? 'block' : 'none';
}

function renderEmojiPicker() {
  const container = document.getElementById('emoji-picker');
  if (!container) return;
  const inner = container.querySelector('div');
  if (!inner) return;
  inner.innerHTML = LOVE_EMOJIS.map(e =>
    `<span onclick="insertEmoji('${e}')" style="font-size:22px;cursor:pointer;padding:5px 6px;border-radius:8px;transition:background .15s;" onmouseover="this.style.background='var(--tint-rose)'" onmouseout="this.style.background='transparent'">${e}</span>`
  ).join('');
}

function insertEmoji(emoji) {
  const ta = document.getElementById('msg-texto');
  const start = ta.selectionStart;
  const end = ta.selectionEnd;
  const text = ta.value;
  ta.value = text.substring(0, start) + emoji + text.substring(end);
  // Posiciona cursor depois do emoji
  const newPos = start + emoji.length;
  ta.setSelectionRange(newPos, newPos);
  ta.focus();
  // Fecha o picker
  document.getElementById('emoji-picker').style.display = 'none';
}

function renderMensagensPanel() {
  const loggedUser = getLoggedUserKey();
  const senderName = loggedUser === 'u2' ? S.settings.u2 : S.settings.u1;
  const destName = loggedUser === 'u2' ? S.settings.u1 : S.settings.u2;
  const destKey = loggedUser === 'u2' ? 'u1' : 'u2';

  document.getElementById('msg-subtitle').textContent = `Envie mensagens de amor para ${destName}`;
  document.getElementById('msg-dest-nome').textContent = destName;

  // Set defaults para datas
  const now = new Date();
  const inicioEl = document.getElementById('msg-inicio');
  const expiraEl = document.getElementById('msg-expira');
  if (inicioEl && !inicioEl.value) {
    inicioEl.value = toLocalDatetime(now);
  }
  if (expiraEl && !expiraEl.value) {
    const exp = new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000); // 7 dias default
    expiraEl.value = toLocalDatetime(exp);
  }

  renderEmojiPicker();
  renderMensagensEnviadas(loggedUser, destKey);
}

function toLocalDatetime(d) {
  const pad = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

function enviarMensagem() {
  const texto = document.getElementById('msg-texto').value.trim();
  const inicio = document.getElementById('msg-inicio').value;
  const expira = document.getElementById('msg-expira').value;

  if (!texto) { toast('Escreva uma mensagem!'); return; }
  if (!inicio || !expira) { toast('Defina as datas de início e expiração!'); return; }
  if (new Date(expira) <= new Date(inicio)) { toast('A data de expiração deve ser depois do início!'); return; }

  const loggedUser = getLoggedUserKey();
  const destKey = loggedUser === 'u2' ? 'u1' : 'u2';

  const msg = {
    id: 'love_' + Date.now() + '_' + Math.random().toString(36).substr(2, 6),
    de: loggedUser,
    para: destKey,
    texto: texto,
    inicio: inicio,
    expira: expira,
    criadoEm: new Date().toISOString()
  };

  if (!S.loveMessages) S.loveMessages = [];
  S.loveMessages.push(msg);
  save();

  // Limpar form
  document.getElementById('msg-texto').value = '';
  const now = new Date();
  document.getElementById('msg-inicio').value = toLocalDatetime(now);
  document.getElementById('msg-expira').value = toLocalDatetime(new Date(now.getTime() + 7*24*60*60*1000));

  toast('Mensagem enviada com amor! 💌');
  renderMensagensEnviadas(loggedUser, destKey);
  updateLeiaAquiVisibility();
}

function renderMensagensEnviadas(senderKey, destKey) {
  const container = document.getElementById('msg-enviadas-list');
  const msgs = (S.loveMessages || []).filter(m => m.de === senderKey && m.para === destKey)
    .sort((a, b) => new Date(b.criadoEm) - new Date(a.criadoEm));

  if (msgs.length === 0) {
    container.innerHTML = '<p style="color:var(--muted);font-size:14px;text-align:center;padding:20px;">Nenhuma mensagem enviada ainda</p>';
    return;
  }

  const now = new Date();
  container.innerHTML = msgs.map(m => {
    const inicio = new Date(m.inicio);
    const expira = new Date(m.expira);
    let status, statusColor, statusBg;
    if (now < inicio) {
      status = 'Agendada';
      statusColor = '#f59e0b';
      statusBg = '#fffbeb';
    } else if (now > expira) {
      status = 'Expirada';
      statusColor = 'var(--muted)';
      statusBg = 'var(--bg)';
    } else {
      status = 'Ativa';
      statusColor = '#10b981';
      statusBg = 'var(--tint-green)';
    }

    return `<div class="card" style="padding:20px;border:1.5px solid #fecdd3;background:linear-gradient(135deg,var(--tint-rose),#ffffff);${now > expira ? 'opacity:0.6;' : ''}">
      <div style="display:flex;justify-content:flex-end;align-items:center;margin-bottom:10px;gap:8px;">
        <span style="font-size:11px;padding:3px 10px;border-radius:20px;background:${statusBg};color:${statusColor};font-weight:600;">${status}</span>
        <button onclick="editarMensagem('${m.id}')" style="background:none;border:none;cursor:pointer;font-size:16px;padding:2px;" title="Editar" class="ico-btn" aria-label="Editar"><svg class="ico-sm"><use href="#i-edit"/></svg></button>
        <button onclick="deletarMensagem('${m.id}')" style="background:none;border:none;cursor:pointer;font-size:16px;padding:2px;" title="Excluir" class="ico-btn" aria-label="Excluir"><svg class="ico-sm"><use href="#i-trash"/></svg></button>
      </div>
      <div id="msg-view-${m.id}">
        <p style="font-size:14px;line-height:1.7;color:#4a1d2e;white-space:pre-wrap;">${escapeHtml(m.texto)}</p>
        <div style="display:flex;justify-content:space-between;margin-top:12px;font-size:11px;color:var(--muted);">
          <span>De ${fmtDatetime(m.inicio)} até ${fmtDatetime(m.expira)}</span>
        </div>
      </div>
      <div id="msg-edit-${m.id}" style="display:none;">
        <textarea id="edit-texto-${m.id}" rows="4" style="width:100%;border:1.5px solid var(--border);border-radius:10px;padding:12px;font-size:14px;line-height:1.7;resize:vertical;font-family:inherit;box-sizing:border-box;">${escapeHtml(m.texto)}</textarea>
        <div style="margin-top:8px;position:relative;">
          <button type="button" onclick="toggleEditEmojiPicker('${m.id}')" style="background:var(--bg);border:1.5px solid var(--border);border-radius:10px;padding:6px 14px;font-size:14px;cursor:pointer;display:flex;align-items:center;gap:6px;">
            <span style="font-size:18px;">😊</span> <span style="font-size:12px;color:var(--text-3);">Inserir emoji</span>
          </button>
          <div id="edit-emoji-picker-${m.id}" style="display:none;position:absolute;left:0;top:40px;z-index:50;background:var(--surface);border:1.5px solid var(--border);border-radius:14px;padding:12px;box-shadow:0 8px 30px rgba(0,0,0,0.12);max-width:320px;">
            <div style="display:flex;flex-wrap:wrap;gap:4px;"></div>
          </div>
        </div>
        <div style="margin-top:12px;display:flex;gap:12px;flex-wrap:wrap;align-items:end;">
          <div style="flex:1;min-width:130px;">
            <label style="font-size:12px;color:var(--text-3);display:block;margin-bottom:4px;">Início:</label>
            <input type="datetime-local" id="edit-inicio-${m.id}" value="${m.inicio}" style="width:100%;border:1.5px solid var(--border);border-radius:10px;padding:8px 10px;font-size:13px;box-sizing:border-box;">
          </div>
          <div style="flex:1;min-width:130px;">
            <label style="font-size:12px;color:var(--text-3);display:block;margin-bottom:4px;">Expirar:</label>
            <input type="datetime-local" id="edit-expira-${m.id}" value="${m.expira}" style="width:100%;border:1.5px solid var(--border);border-radius:10px;padding:8px 10px;font-size:13px;box-sizing:border-box;">
          </div>
        </div>
        <div style="display:flex;gap:10px;margin-top:14px;">
          <button onclick="salvarEdicaoMsg('${m.id}')" style="flex:1;padding:10px;background:linear-gradient(135deg,#10b981,#059669);color:#fff;border:none;border-radius:10px;font-weight:700;font-size:14px;cursor:pointer;">Salvar</button>
          <button onclick="cancelarEdicao('${m.id}')" style="flex:1;padding:10px;background:var(--surface-2);color:var(--text-3);border:none;border-radius:10px;font-weight:700;font-size:14px;cursor:pointer;">Cancelar</button>
        </div>
      </div>
    </div>`;
  }).join('');
}

function editarMensagem(id) {
  // Esconder view, mostrar edit
  document.getElementById('msg-view-' + id).style.display = 'none';
  document.getElementById('msg-edit-' + id).style.display = 'block';
  // Renderizar emoji picker da edição
  const pickerInner = document.querySelector('#edit-emoji-picker-' + id + ' > div');
  if (pickerInner) {
    pickerInner.innerHTML = LOVE_EMOJIS.map(e =>
      `<span onclick="insertEditEmoji('${id}','${e}')" style="font-size:22px;cursor:pointer;padding:5px 6px;border-radius:8px;transition:background .15s;" onmouseover="this.style.background='var(--tint-rose)'" onmouseout="this.style.background='transparent'">${e}</span>`
    ).join('');
  }
}

function cancelarEdicao(id) {
  document.getElementById('msg-view-' + id).style.display = 'block';
  document.getElementById('msg-edit-' + id).style.display = 'none';
}

function salvarEdicaoMsg(id) {
  const texto = document.getElementById('edit-texto-' + id).value.trim();
  const inicio = document.getElementById('edit-inicio-' + id).value;
  const expira = document.getElementById('edit-expira-' + id).value;

  if (!texto) { toast('A mensagem não pode ficar vazia!'); return; }
  if (!inicio || !expira) { toast('Defina as datas!'); return; }
  if (new Date(expira) <= new Date(inicio)) { toast('Expiração deve ser depois do início!'); return; }

  const msg = (S.loveMessages || []).find(m => m.id === id);
  if (msg) {
    msg.texto = texto;
    msg.inicio = inicio;
    msg.expira = expira;
    save();
    const loggedUser = getLoggedUserKey();
    const destKey = loggedUser === 'u2' ? 'u1' : 'u2';
    renderMensagensEnviadas(loggedUser, destKey);
    updateLeiaAquiVisibility();
    toast('Mensagem atualizada! 💌');
  }
}

function toggleEditEmojiPicker(id) {
  const picker = document.getElementById('edit-emoji-picker-' + id);
  picker.style.display = picker.style.display === 'none' ? 'block' : 'none';
}

function insertEditEmoji(id, emoji) {
  const ta = document.getElementById('edit-texto-' + id);
  const start = ta.selectionStart;
  const end = ta.selectionEnd;
  const text = ta.value;
  ta.value = text.substring(0, start) + emoji + text.substring(end);
  const newPos = start + emoji.length;
  ta.setSelectionRange(newPos, newPos);
  ta.focus();
  document.getElementById('edit-emoji-picker-' + id).style.display = 'none';
}

function deletarMensagem(id) {
  if (!confirm('Deseja excluir esta mensagem?')) return;
  S.loveMessages = (S.loveMessages || []).filter(m => m.id !== id);
  save();
  const loggedUser = getLoggedUserKey();
  const destKey = loggedUser === 'u2' ? 'u1' : 'u2';
  renderMensagensEnviadas(loggedUser, destKey);
  updateLeiaAquiVisibility();
  toast('Mensagem excluída');
}

function renderLeiaAqui() {
  const container = document.getElementById('cartas-amor');
  const semMsg = document.getElementById('sem-mensagens');
  const received = getReceivedMessages();

  if (received.length === 0) {
    container.innerHTML = '<p id="sem-mensagens" style="color:var(--muted);font-size:14px;text-align:center;padding:30px;">Nenhuma mensagem no momento</p>';
    return;
  }

  const sorted = received.sort((a, b) => new Date(b.criadoEm) - new Date(a.criadoEm));
  const senderName = sorted[0].de === 'u1' ? S.settings.u1 : S.settings.u2;

  container.innerHTML = sorted.map((m, i) => {
    const bg = i % 2 === 0 ? '#fff0f3' : '#fdf2f8';
    return `<div class="card" style="padding:28px;border:1.5px solid #fecdd3;background:linear-gradient(135deg,${bg},#ffffff);animation:fadeInUp 0.5s ease ${i*0.1}s both;">
      <p style="font-size:15px;line-height:1.8;color:#4a1d2e;white-space:pre-wrap;">${escapeHtml(m.texto)}</p>
    </div>`;
  }).join('') + `
    <div style="text-align:center;padding:20px;">
      <p style="font-size:16px;font-weight:700;color:var(--on-rose);">Com todo meu amor,</p>
      <p style="font-size:18px;font-weight:800;color:var(--on-rose);margin-top:4px;">${escapeHtml(senderName)}</p>
      <p style="font-size:28px;margin-top:8px;">❤️</p>
    </div>`;
}

function escapeHtml(text) {
  return String(text === null || text === undefined ? '' : text)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function fmtDatetime(dt) {
  if (!dt) return '';
  const d = new Date(dt);
  const pad = n => String(n).padStart(2, '0');
  return `${pad(d.getDate())}/${pad(d.getMonth()+1)}/${d.getFullYear()} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

function getLoggedUserName() {
  const logged = sessionStorage.getItem('fincasal_logged_user');
  if (logged === 'u2') return S.settings.u2;
  if (logged === 'u1') return S.settings.u1;
  return S.settings.u1;
}

function renderDashboard() {
  // Update filter button labels
  document.getElementById('tf-paulo').textContent  = S.settings.u1;
  document.getElementById('tf-esposa').textContent = S.settings.u2;
  updateTempoJuntos();
  updateLeiaAquiVisibility();

  // Greeting with logged user name
  const userName = getLoggedUserName();
  document.getElementById('dash-greeting').innerHTML = `Olá, ${userName}! 👋`;

  // Init period selectors (preserve current selection)
  const mSel = document.getElementById('dash-month');
  if (!mSel.options.length) initDashPeriod();

  const isAno = dashPeriodMode === 'ano';
  const ym = isAno ? String(dashYear) : `${dashYear}-${String(dashMonth + 1).padStart(2, '0')}`;
  const prevYear = isAno ? dashYear - 1 : null;
  const prev = isAno ? null : new Date(dashYear, dashMonth - 1, 1);
  const ymPrev = isAno ? String(dashYear - 1) : `${prev.getFullYear()}-${String(prev.getMonth() + 1).padStart(2, '0')}`;

  const filterLabel = titularFilter === 'ambos' ? 'Geral' : (titularFilter === 'paulo' ? S.settings.u1 : S.settings.u2);
  const periodLabel = isAno ? `Ano ${dashYear}` : `${MESES_FULL[dashMonth]} ${dashYear}`;
  document.getElementById('dash-subtitle').textContent = `Bem-vindo(a) de volta! Aqui está o resumo das suas finanças. — ${filterLabel} · ${periodLabel}`;

  if (dashView === 'conta') {
    renderDashConta(ym, ymPrev);
  } else {
    renderCreditView();
  }
}

function renderDashConta(ym, ymPrev) {
  const allMTx  = S.transactions.filter(t => t.date.startsWith(ym));
  const mTx = txByTitular(allMTx);
  const pTx = txByTitular(S.transactions.filter(t => t.date.startsWith(ymPrev)));

  const allTitTx = txByTitular(S.transactions);

  // SALDO = receitas - débitos pagos (o que realmente entrou/saiu da conta bancária)
  const rec   = mTx.filter(t=>t.type==='receita' && !t.isTransfer).reduce((s,t)=>s+amountBrl(t),0);
  const despDebito = mTx.filter(t=>t.type==='despesa' && t.formaPgto !== 'credito' && !t.isTransfer && t.pago !== false).reduce((s,t)=>s+amountBrl(t),0);
  const saldo = rec - despDebito;

  // DESPESAS (visão Conta) = tudo que saiu da conta bancária:
  //   débitos pagos (inclui Pgto Fatura) — exclui transferências e crédito individual
  const debitoYM = mTx.filter(t =>
    t.type === 'despesa' && t.formaPgto !== 'credito' && !t.isTransfer && t.pago !== false
  );
  const desp = debitoYM.reduce((s,t)=>s+amountBrl(t),0);

  // Mês anterior
  const pRec   = pTx.filter(t=>t.type==='receita' && !t.isTransfer).reduce((s,t)=>s+amountBrl(t),0);
  const pDespDebito = pTx.filter(t=>t.type==='despesa' && t.formaPgto !== 'credito' && !t.isTransfer && t.pago !== false).reduce((s,t)=>s+amountBrl(t),0);
  const pSaldo = pRec - pDespDebito;
  const debitoPrev = pTx.filter(t =>
    t.type === 'despesa' && t.formaPgto !== 'credito' && !t.isTransfer && t.pago !== false
  );
  const pDesp = debitoPrev.reduce((s,t)=>s+amountBrl(t),0);

  // Saldo acumulado até o mês selecionado (saldo real nas contas)
  // Inclui transferências (se cancelam no total, mas mantém consistência com breakdown por conta)
  // Exclui crédito (não afeta conta bancária) e pendentes (ainda não saíram)
  const acumTx = allTitTx.filter(t => t.date <= ym + '-31');
  const acumRec  = acumTx.filter(t=>t.type==='receita').reduce((s,t)=>s+amountBrl(t),0);
  const acumDesp = acumTx.filter(t=>t.type==='despesa' && t.formaPgto !== 'credito' && t.pago !== false).reduce((s,t)=>s+amountBrl(t),0);
  const saldoAcumulado = acumRec - acumDesp;

  const cs = document.getElementById('c-saldo');
  cs.textContent = brl(saldoAcumulado);
  cs.style.color = saldoAcumulado >= 0 ? '#059669' : '#e11d48';

  // Saldo acumulado do mês anterior
  const acumTxPrev = allTitTx.filter(t => t.date <= ymPrev + '-31');
  const acumRecPrev  = acumTxPrev.filter(t=>t.type==='receita').reduce((s,t)=>s+amountBrl(t),0);
  const acumDespPrev = acumTxPrev.filter(t=>t.type==='despesa' && t.formaPgto !== 'credito' && t.pago !== false).reduce((s,t)=>s+amountBrl(t),0);
  const saldoAcumPrev = acumRecPrev - acumDespPrev;

  // Percentage change (compara saldo acumulado atual vs anterior)
  if (saldoAcumPrev !== 0) {
    const pct = ((saldoAcumulado - saldoAcumPrev) / Math.abs(saldoAcumPrev)) * 100;
    document.getElementById('c-saldo-pct').textContent = `${pct >= 0 ? '+' : ''}${pct.toFixed(1)}%`;
    document.getElementById('c-saldo-pct').style.background = pct >= 0 ? 'var(--tint-green)' : 'var(--tint-rose)';
    document.getElementById('c-saldo-pct').style.color = pct >= 0 ? 'var(--on-green)' : 'var(--on-rose)';
  } else {
    document.getElementById('c-saldo-pct').textContent = '';
  }

  // Subtexto: saldo do mês isolado
  const saldoMesLabel = saldo >= 0 ? `+${brl(saldo)}` : `-${brl(Math.abs(saldo))}`;
  document.getElementById('c-saldo-sub').textContent = `No mês: ${saldoMesLabel}`;
  document.getElementById('c-saldo-sub').style.color = saldo >= 0 ? '#059669' : '#e11d48';

  // Receitas with percentage change
  document.getElementById('c-receitas').textContent = brl(rec);
  const recN = mTx.filter(t=>t.type==='receita').length;
  document.getElementById('c-receitas-n').textContent = `${recN} lançamentos`;
  if (pRec !== 0) {
    const pctR = ((rec - pRec) / Math.abs(pRec)) * 100;
    document.getElementById('c-receitas-pct').textContent = `${pctR >= 0 ? '+' : ''}${pctR.toFixed(1)}%`;
    document.getElementById('c-receitas-pct').style.background = pctR >= 0 ? 'var(--tint-green)' : 'var(--tint-rose)';
    document.getElementById('c-receitas-pct').style.color = pctR >= 0 ? 'var(--on-green)' : 'var(--on-rose)';
  } else {
    document.getElementById('c-receitas-pct').textContent = '';
  }

  // Despesas with percentage change and pending info
  document.getElementById('c-despesas').textContent = brl(desp);
  const despN = debitoYM.length;
  // Pendente: débito pendente do mês (excl transferências, inclui Pgto Fatura)
  const debitoPendYM = mTx.filter(t =>
    t.type === 'despesa' && t.formaPgto !== 'credito' && !t.isTransfer && t.pago === false
  );
  const pendente = debitoPendYM.reduce((s,t)=>s+amountBrl(t),0);
  let despSub = `${despN} lançamentos`;
  if (pendente > 0) despSub += ` · Pendente: ${brl(pendente)}`;
  document.getElementById('c-despesas-n').textContent = despSub;
  if (pDesp !== 0) {
    const pctD = ((desp - pDesp) / Math.abs(pDesp)) * 100;
    document.getElementById('c-despesas-pct').textContent = `${pctD >= 0 ? '+' : ''}${pctD.toFixed(1)}%`;
    document.getElementById('c-despesas-pct').style.background = pctD >= 0 ? 'var(--tint-rose)' : 'var(--tint-green)';
    document.getElementById('c-despesas-pct').style.color = pctD >= 0 ? 'var(--on-rose)' : 'var(--on-green)';
  } else {
    document.getElementById('c-despesas-pct').textContent = '';
  }

  const allInvests = (S.investments || []).filter(i => {
    if (i.tipo === 'negocio') return false;
    if (titularFilter === 'ambos') return true;
    const name = titularFilter === 'paulo' ? S.settings.u1 : S.settings.u2;
    return i.titular === name;
  });
  const invTotal = allInvests.reduce((s,i) => s + (i.valorAtual || i.valorInvestido||0), 0);
  const invRend = allInvests.reduce((s,i) => s + (i.rendimentos||[]).reduce((r,x) => r + (x.valor||0), 0), 0);
  document.getElementById('c-invest').textContent = brl(invTotal);
  document.getElementById('c-invest-n').textContent = `${allInvests.length} investimento${allInvests.length !== 1 ? 's' : ''} · Rend: ${brl(invRend)}`;

  // Populate expandable breakdowns per bank
  renderCardBreakdowns(mTx);

  // Donut: débito do mês + crédito pago do mês (por faturaRef)
  renderDebitCatBreakdown(debitoYM);
  renderBarDash();
  renderCatStackedChart();
  renderRecent();
  renderInsights(ym);
  renderConfirmar(ym);
  renderProjecao(ym);
}

function toggleCardBreakdown(type) {
  const el = document.getElementById('breakdown-' + type);
  const chevron = document.getElementById('chevron-' + type);
  if (el.style.display === 'none') {
    el.style.display = 'block';
    chevron.textContent = '▴';
  } else {
    el.style.display = 'none';
    chevron.textContent = '▾';
  }
}

function renderCardBreakdowns(mTx) {
  const contas = S.accounts.filter(a => a.accountType !== 'cartao');
  const accToShow = titularFilter === 'ambos'
    ? contas
    : contas.filter(a => a.owner === (titularFilter === 'paulo' ? S.settings.u1 : S.settings.u2));

  function bankRow(bank, val, color) {
    const b = BANKS[bank];
    return `<div style="display:flex;align-items:center;gap:10px;padding:6px 0;">
      ${bankIcon(bank, 28)}
      <span style="font-size:13px;font-weight:600;color:var(--text-2);flex:1;">${escapeHtml(b.name)}</span>
      <span style="font-size:13px;font-weight:700;color:${color};">${brl(val)}</span>
    </div>`;
  }

  // Saldo breakdown — acumulado até o mês selecionado
  const isAno = dashPeriodMode === 'ano';
  const ymLimit = isAno ? `${dashYear}-12-31` : `${dashYear}-${String(dashMonth+1).padStart(2,'0')}-31`;
  let saldoHtml = '';
  accToShow.forEach(a => {
    const accTxAcum = txByTitular(S.transactions).filter(t => t.accountId === a.id && t.date <= ymLimit);
    const rec  = accTxAcum.filter(t=>t.type==='receita').reduce((s,t)=>s+amountRaw(t),0);
    const desp = accTxAcum.filter(t=>t.type==='despesa' && t.formaPgto !== 'credito' && t.pago !== false).reduce((s,t)=>s+amountRaw(t),0);
    const bal = rec - desp;
    const isUSD = BANKS[a.bank].currency === 'USD';
    const balBrl = isUSD ? usdToBrl(bal) : bal;
    saldoHtml += bankRow(a.bank, balBrl, balBrl >= 0 ? '#059669' : '#e11d48');
  });
  document.getElementById('breakdown-saldo').innerHTML = saldoHtml;

  // Receitas breakdown (inclui transferências recebidas por banco)
  let recHtml = '';
  accToShow.forEach(a => {
    const val = mTx.filter(t => t.accountId === a.id && t.type === 'receita').reduce((s,t)=>s+amountBrl(t),0);
    recHtml += bankRow(a.bank, val, '#059669');
  });
  document.getElementById('breakdown-receitas').innerHTML = recHtml;

  // Despesas breakdown (inclui transferências enviadas por banco)
  let despHtml = '';
  accToShow.forEach(a => {
    const val = mTx.filter(t => t.accountId === a.id && t.type === 'despesa' && t.formaPgto !== 'credito').reduce((s,t)=>s+amountBrl(t),0);
    despHtml += bankRow(a.bank, val, '#e11d48');
  });
  document.getElementById('breakdown-despesas').innerHTML = despHtml;

  // Investimentos breakdown (from S.investments)
  const allInvestsTit = (S.investments || []).filter(i => {
    if (i.tipo === 'negocio') return false;
    if (titularFilter === 'ambos') return true;
    const name = titularFilter === 'paulo' ? S.settings.u1 : S.settings.u2;
    return i.titular === name;
  });
  let invHtml = '';
  allInvestsTit.forEach(i => {
    const tp = typeof TIPO_INVEST !== 'undefined' ? (TIPO_INVEST[i.tipo] || {}) : {};
    const rend = (i.rendimentos||[]).reduce((s,r) => s + (r.valor||0), 0);
    invHtml += `<div style="display:flex;align-items:center;gap:10px;padding:6px 0;">
      <span style="font-size:18px;">${tp.icon || '💼'}</span>
      <span style="font-size:13px;font-weight:600;color:var(--text-2);flex:1;">${escapeHtml(i.desc)}</span>
      <span style="font-size:13px;font-weight:700;color:#2563eb;">${brl(i.valorInvestido)}</span>
    </div>`;
  });
  document.getElementById('breakdown-invest').innerHTML = invHtml || '<p style="font-size:12px;color:var(--muted);text-align:center;">Nenhum investimento</p>';
}

let debCatMode = 'cat';
let debCatDrillCat = null;
let debCatTxCache = [];

function setDebitCatMode(mode) {
  debCatMode = mode;
  debCatDrillCat = null;
  document.getElementById('deb-cat-mode-cat').style.cssText = `padding:3px 10px;border-radius:6px;border:none;font-size:11px;font-weight:600;cursor:pointer;background:${mode==='cat'?'#4f46e5':'transparent'};color:${mode==='cat'?'white':'var(--text-3)'};`;
  document.getElementById('deb-cat-mode-sub').style.cssText = `padding:3px 10px;border-radius:6px;border:none;font-size:11px;font-weight:600;cursor:pointer;background:${mode==='sub'?'#4f46e5':'transparent'};color:${mode==='sub'?'white':'var(--text-3)'};`;
  document.getElementById('deb-cat-back').style.display = 'none';
  renderDebitCatBreakdown(debCatTxCache);
}

function debCatBack() {
  debCatDrillCat = null;
  document.getElementById('deb-cat-back').style.display = 'none';
  renderDebitCatBreakdown(debCatTxCache);
}

function debCatDrill(cat) {
  debCatMode = 'sub';
  debCatDrillCat = cat;
  document.getElementById('deb-cat-mode-cat').style.cssText = `padding:3px 10px;border-radius:6px;border:none;font-size:11px;font-weight:600;cursor:pointer;background:transparent;color:var(--text-3);`;
  document.getElementById('deb-cat-mode-sub').style.cssText = `padding:3px 10px;border-radius:6px;border:none;font-size:11px;font-weight:600;cursor:pointer;background:#4f46e5;color:white;`;
  document.getElementById('deb-cat-back').style.display = 'block';
  document.getElementById('deb-cat-back').querySelector('button').textContent = `← ${cat}`;
  renderDebitCatBreakdown(debCatTxCache);
}

function isExcludedFromChart(t) {
  if (isPgtoFatura(t)) return true;
  const sub = (t.subcategory || '').toLowerCase();
  const cat = (t.category || '').toLowerCase();
  if (sub === 'investimento' || sub === 'investimentos') return true;
  if (sub === 'emprestimo' || sub === 'empréstimo' || sub === 'emprestimos' || sub === 'empréstimos') return true;
  if (sub === 'ajuste saldo' || sub === 'conciliacao' || sub === 'conciliação') return true;
  if (cat === 'taxas' && sub === 'banco') return true;
  return false;
}

function renderDebitCatBreakdown(txs) {
  debCatTxCache = txs;
  // Só débito, excluir transferências e itens que Paulo não quer ver
  const filtered = txs.filter(t => !t.isTransfer && !isExcludedFromChart(t));

  const by = {};
  if (debCatMode === 'sub' && debCatDrillCat) {
    const catTxs = filtered.filter(t => t.category === debCatDrillCat);
    catTxs.forEach(t => {
      const key = t.subcategory || 'Geral';
      by[key] = (by[key] || 0) + amountBrl(t);
    });
  } else if (debCatMode === 'sub') {
    filtered.forEach(t => {
      const key = t.subcategory ? `${t.category} > ${t.subcategory}` : t.category;
      by[key] = (by[key] || 0) + amountBrl(t);
    });
  } else {
    filtered.forEach(t => by[t.category] = (by[t.category] || 0) + amountBrl(t));
  }

  const sorted = Object.entries(by).sort((a, b) => b[1] - a[1]);
  const total = sorted.reduce((s, [, v]) => s + v, 0);

  document.getElementById('deb-cat-total').textContent = `Total ${brl(total)}`;

  const el = document.getElementById('deb-cat-list');
  if (!sorted.length) {
    el.innerHTML = '<p style="font-size:13px;color:var(--muted);text-align:center;padding:20px;">Nenhuma despesa no débito este mês</p>';
    return;
  }

  el.innerHTML = sorted.slice(0, 12).map(([cat, val], i) => {
    const pct = total > 0 ? (val / total) * 100 : 0;
    const clickable = debCatMode === 'cat';
    return `<div style="display:flex;align-items:center;gap:12px;${clickable?'cursor:pointer;':''}" ${clickable ? `onclick="debCatDrill('${cat.replace(/'/g,"\\'")}')"` : ''}>
      <div style="width:10px;height:10px;border-radius:50%;background:${COLORS[i % COLORS.length]};flex-shrink:0;"></div>
      <div style="flex:1;">
        <div style="display:flex;justify-content:space-between;margin-bottom:4px;">
          <span style="font-size:13px;font-weight:600;color:var(--text-2);">${escapeHtml(cat)}</span>
          <span style="font-size:13px;font-weight:700;color:var(--text-2);">${brl(val)}</span>
        </div>
        <div style="height:6px;background:var(--surface-2);border-radius:3px;overflow:hidden;">
          <div style="height:100%;width:${pct}%;background:${COLORS[i % COLORS.length]};border-radius:3px;"></div>
        </div>
      </div>
    </div>`;
  }).join('');
}

function renderBarDash() {
  const now = new Date();
  const period = document.getElementById('bar-period')?.value || 'month';
  const labels=[], recs=[], desps=[];
  const months = period === 'year' ? 12 : 6;
  for (let i=months-1;i>=0;i--) {
    const d = new Date(now.getFullYear(), now.getMonth()-i, 1);
    const k = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;
    labels.push(MESES[d.getMonth()]);
    const tx = txByTitular(S.transactions.filter(t=>t.date.startsWith(k)));
    recs.push(tx.filter(t=>t.type==='receita').reduce((s,t)=>s+amountBrl(t),0));
    desps.push(tx.filter(t=>t.type==='despesa' && t.formaPgto !== 'credito').reduce((s,t)=>s+amountBrl(t),0));
  }

  destroyChart('ch-bar-dash');
  const ctx = document.getElementById('ch-bar-dash');
  if (!ctx) return;

  const chartType = barChartType === 'area' ? 'line' : barChartType;
  const fill = barChartType === 'area';

  charts['ch-bar-dash'] = new Chart(ctx, {
    type: chartType,
    data: { labels, datasets:[
      { label:'Receitas',  data:recs,  backgroundColor: chartType === 'line' ? 'rgba(16,185,129,0.1)' : '#10b981', borderColor:'#10b981', borderRadius: chartType === 'bar' ? 6 : 0, borderWidth: chartType === 'line' ? 2 : 0, fill, tension: 0.3, pointRadius: chartType === 'line' ? 3 : 0 },
      { label:'Despesas',  data:desps, backgroundColor: chartType === 'line' ? 'rgba(244,63,94,0.1)' : '#f43f5e', borderColor:'#f43f5e', borderRadius: chartType === 'bar' ? 6 : 0, borderWidth: chartType === 'line' ? 2 : 0, fill, tension: 0.3, pointRadius: chartType === 'line' ? 3 : 0 }
    ]},
    options: {
      responsive:true, maintainAspectRatio:false,
      plugins: { legend: { position:'top', labels:{boxWidth:12,font:{size:11}} },
        tooltip: { callbacks: { label: c => ` ${c.dataset.label}: ${brl(c.raw)}` } } },
      scales: { x:{grid:{display:false}}, y:{grid:{color:cssVar('--chart-grid')}, ticks:{callback:v=>v===0?'R$0':'R$'+v.toLocaleString('pt-BR')}} }
    }
  });
}

function insightCard(emoji, titulo, valor, sub, cor) {
  return `<div class="card" style="min-width:200px;flex:1;padding:16px;border-left:4px solid ${cor};">
    <div style="font-size:12px;color:var(--muted);font-weight:600;display:flex;align-items:center;gap:6px;">${emoji} ${titulo}</div>
    <div class="hv" style="font-size:20px;font-weight:800;color:var(--text);margin:6px 0 2px;">${valor}</div>
    <div class="hv" style="font-size:12px;color:var(--text-3);line-height:1.4;">${sub}</div>
  </div>`;
}

// Painel de análises automáticas do dashboard (só leitura, calculado do estado)
function renderInsights(ym) {
  const el = document.getElementById('dash-insights');
  if (!el) return;
  const cards = [];
  const isCurMonth = ym === new Date().toISOString().slice(0, 7);
  const my = txByTitular(S.transactions);
  const gastoReal = t => t.type === 'despesa' && !t.isTransfer && !isExcludedFromChart(t);
  const somaMes = (arr, k) => arr.filter(t => (t.date || '').startsWith(k) && gastoReal(t)).reduce((s, t) => s + amountBrl(t), 0);
  const ymPrev = (n) => {
    const d = new Date(parseInt(ym.slice(0, 4)), parseInt(ym.slice(5, 7)) - 1 - n, 1);
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
  };

  // 1. Ritmo / gasto do mês
  try {
    const gasto = somaMes(my, ym);
    let soma3 = 0, n3 = 0;
    for (let i = 1; i <= 3; i++) { const g = somaMes(my, ymPrev(i)); if (g > 0) { soma3 += g; n3++; } }
    const media3 = n3 ? soma3 / n3 : 0;
    if (isCurMonth) {
      const hoje = new Date(), diaHoje = hoje.getDate();
      const diasMes = new Date(hoje.getFullYear(), hoje.getMonth() + 1, 0).getDate();
      const proj = diaHoje > 0 ? gasto / diaHoje * diasMes : gasto;
      let sub = `Projeção fim do mês: <b>${brl(proj)}</b>`;
      if (media3 > 0) { const dif = (proj - media3) / media3 * 100; sub += ` · ${dif >= 0 ? '+' : ''}${dif.toFixed(0)}% vs média`; }
      cards.push(insightCard('📈', 'Ritmo do mês', brl(gasto) + ' até agora', sub, (proj > media3 && media3 > 0) ? '#f43f5e' : '#059669'));
    } else {
      cards.push(insightCard('📊', 'Gasto do mês', brl(gasto), media3 > 0 ? `Média 3 meses: ${brl(media3)}` : 'Sem histórico p/ comparar', '#4f46e5'));
    }
  } catch (e) {}

  // 2. Categoria em alta vs média dos 3 meses anteriores
  try {
    const cur = {}, prev = {};
    my.filter(t => (t.date || '').startsWith(ym) && gastoReal(t)).forEach(t => cur[t.category] = (cur[t.category] || 0) + amountBrl(t));
    for (let i = 1; i <= 3; i++) {
      my.filter(t => (t.date || '').startsWith(ymPrev(i)) && gastoReal(t)).forEach(t => prev[t.category] = (prev[t.category] || 0) + amountBrl(t));
    }
    let best = null;
    for (const c in cur) {
      const media = prev[c] ? prev[c] / 3 : 0;
      if (media > 50 && cur[c] > media) { const dif = (cur[c] - media) / media * 100; if (!best || dif > best.dif) best = { c, dif, val: cur[c] }; }
    }
    if (best) cards.push(insightCard('🔥', 'Categoria em alta', escapeHtml(best.c), `<b>+${best.dif.toFixed(0)}%</b> vs média · ${brl(best.val)}`, '#f43f5e'));
  } catch (e) {}

  // 3. Próxima fatura a vencer
  try {
    const cartoes = S.accounts.filter(a => a.accountType === 'cartao' && (titularFilter === 'ambos' || a.owner === (titularFilter === 'paulo' ? S.settings.u1 : S.settings.u2)));
    const hoje = new Date(); let prox = null;
    const fatYM = getCurrentFaturaYM();
    cartoes.forEach(c => {
      const total = txByTitular(getCreditTxByFatura(fatYM, c.id)).reduce((s, t) => s + amountBrl(t), 0);
      if (total <= 0) return;
      const venceDia = parseInt(c.vence) || 10;
      let venc = new Date(hoje.getFullYear(), hoje.getMonth(), venceDia);
      if (venc < hoje) venc = new Date(hoje.getFullYear(), hoje.getMonth() + 1, venceDia);
      const dias = Math.ceil((venc - hoje) / 86400000);
      if (!prox || dias < prox.dias) prox = { nome: BANKS[c.bank] ? BANKS[c.bank].name : c.label, total, dias };
    });
    if (prox) cards.push(insightCard('💳', 'Próxima fatura', brl(prox.total), `${escapeHtml(prox.nome)} · vence em <b>${prox.dias} dia${prox.dias !== 1 ? 's' : ''}</b>`, prox.dias <= 5 ? '#f43f5e' : '#4f46e5'));
  } catch (e) {}

  // 4. Parcelas terminando neste mês
  try {
    const term = my.filter(t => {
      if (!(t.date || '').startsWith(ym)) return false;
      const m = /^(\d+)\/(\d+)$/.exec(t.parcela || '');
      return m && m[1] === m[2] && parseInt(m[2]) > 1;
    });
    if (term.length) {
      const val = term.reduce((s, t) => s + amountBrl(t), 0);
      cards.push(insightCard('🎉', 'Parcelas terminando', `${term.length} última${term.length !== 1 ? 's' : ''}`, `Alívio de <b>${brl(val)}</b>/mês a partir do próximo`, '#059669'));
    }
  } catch (e) {}

  // 5. Orçamento estourado
  try {
    const mes = parseInt(ym.slice(5, 7)) - 1, ano = parseInt(ym.slice(0, 4));
    const est = [];
    Object.keys(S.catOrcGroup || {}).forEach(cat => {
      const plan = S.budget ? (S.budget[getBudgetKey(ano, mes, cat)] || 0) : 0;
      if (plan <= 0) return;
      const real = my.filter(t => t.type === 'despesa' && t.category === cat && (t.date || '').startsWith(ym)).reduce((s, t) => s + amountBrl(t), 0);
      if (real > plan) est.push({ cat, over: real - plan });
    });
    if (est.length) {
      est.sort((a, b) => b.over - a.over);
      cards.push(insightCard('⚠️', 'Orçamento estourado', `${est.length} categoria${est.length !== 1 ? 's' : ''}`, `Pior: ${escapeHtml(est[0].cat)} (+${brl(est[0].over)})`, '#f59e0b'));
    }
  } catch (e) {}

  // 6. Acerto do casal (despesas marcadas como "do casal", divididas 50/50)
  try {
    const ac = calcAcerto(ym);
    if (ac.total > 0) {
      const feito = (S.acertos || {})[ym];
      let valor, sub;
      if (feito) {
        valor = '✅ Acertado';
        sub = `${escapeHtml(feito.from)} pagou <b>${brl(feito.amount)}</b> pra ${escapeHtml(feito.to)} em ${(feito.at || '').slice(0, 10).split('-').reverse().join('/')} · <a href="#" onclick="desfazerAcerto('${ym}');return false;" style="color:var(--text-3);">desfazer</a>`;
      } else if (Math.abs(ac.saldo) < 0.01) {
        valor = 'Empatados';
        sub = `Cada um pagou ${brl(ac.total / 2)} das despesas do casal`;
      } else {
        valor = `${escapeHtml(ac.devedor)} deve ${brl(Math.abs(ac.saldo))}`;
        sub = `${escapeHtml(S.settings.u1)} pagou ${brl(ac.pagou[S.settings.u1] || 0)} · ${escapeHtml(S.settings.u2)} pagou ${brl(ac.pagou[S.settings.u2] || 0)} · total ${brl(ac.total)}` +
          ` <button onclick="marcarAcerto('${ym}')" style="margin-left:6px;padding:3px 10px;border-radius:8px;border:1px solid var(--border);background:var(--surface);color:var(--text-2);font-size:11px;font-weight:600;cursor:pointer;">Marcar acertado</button>`;
      }
      cards.push(insightCard('👥', 'Acerto do casal', valor, sub, feito ? '#059669' : '#ec4899'));
    }
  } catch (e) {}

  if (!cards.length) { el.style.display = 'none'; return; }
  el.style.display = 'flex';
  el.innerHTML = cards.join('');
}

// ─── ACERTO DO CASAL ──────────────────────────────────────────────────────────
// Despesas marcadas como "do casal" são divididas 50/50. Quem pagou mais recebe a diferença.
function calcAcerto(ym) {
  const u1 = S.settings.u1, u2 = S.settings.u2;
  const txs = S.transactions.filter(t => t.type === 'despesa' && t.compartilhada && !t.isTransfer && !isPgtoFatura(t) && (t.date || '').startsWith(ym));
  const pagou = {};
  txs.forEach(t => { pagou[t.user] = (pagou[t.user] || 0) + amountBrl(t); });
  const total = txs.reduce((s, t) => s + amountBrl(t), 0);
  const saldo = ((pagou[u1] || 0) - (pagou[u2] || 0)) / 2;   // >0: u2 deve pra u1
  return { total, pagou, saldo, devedor: saldo > 0 ? u2 : u1, credor: saldo > 0 ? u1 : u2, n: txs.length };
}
function marcarAcerto(ym) {
  const ac = calcAcerto(ym);
  if (!S.acertos) S.acertos = {};
  S.acertos[ym] = { amount: Math.round(Math.abs(ac.saldo) * 100) / 100, from: ac.devedor, to: ac.credor, at: new Date().toISOString() };
  save(); renderInsights(ym);
  toast(`✅ Acerto de ${MESES_FULL[parseInt(ym.slice(5, 7)) - 1]} registrado`);
}
function desfazerAcerto(ym) {
  if (S.acertos) delete S.acertos[ym];
  save(); renderInsights(ym);
}

// ─── CONTAS DO MÊS PRA CONFIRMAR ──────────────────────────────────────────────
// Despesas de débito ainda pendentes no mês (recorrentes geradas + lançadas como pendentes).
// Confirmar = marcar como paga (com o valor ajustado); "Não veio" = remove só esta ocorrência.
function renderConfirmar(ym) {
  const el = document.getElementById('dash-confirmar');
  if (!el) return;
  const pend = txByTitular(S.transactions).filter(t => t.type === 'despesa' && t.pago === false && t.formaPgto !== 'credito'
    && !t.isTransfer && !isPgtoFatura(t) && (t.date || '').startsWith(ym)).sort((a, b) => (a.date || '').localeCompare(b.date || ''));
  if (!pend.length) { el.style.display = 'none'; return; }
  const total = pend.reduce((s, t) => s + amountBrl(t), 0);
  el.style.display = 'block';
  el.innerHTML = `<div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px;margin-bottom:10px;">
      <div><span style="font-size:14px;font-weight:700;color:var(--text);">📋 Contas do mês pra confirmar</span>
        <span style="font-size:12px;color:var(--muted);margin-left:8px;">${pend.length} pendente${pend.length !== 1 ? 's' : ''} · ${brl(total)}</span></div>
      <button onclick="confirmarTodas('${ym}')" style="padding:6px 12px;border-radius:8px;border:1px solid var(--border);background:var(--surface);color:var(--text-2);font-size:12px;font-weight:600;cursor:pointer;">Confirmar todas</button>
    </div>
    <div style="display:flex;flex-direction:column;gap:6px;">` +
    pend.map(t => {
      const acc = S.accounts.find(a => a.id === t.accountId);
      return `<div id="conf-${t.id}" style="display:flex;align-items:center;gap:10px;padding:8px 10px;border-radius:10px;background:var(--bg);flex-wrap:wrap;">
        <div style="flex:1;min-width:160px;"><div style="font-size:13px;font-weight:600;color:var(--text);">${escapeHtml(t.desc)}${t.recorrente ? ' <span style="font-size:10px;color:#15803d;">🔁</span>' : ''}</div>
          <div style="font-size:11px;color:var(--muted);">${(t.date || '').split('-').reverse().join('/')} · ${acc ? escapeHtml(acc.label) : ''} · ${escapeHtml(t.category || '')}</div></div>
        <div style="position:relative;"><span style="position:absolute;left:10px;top:50%;transform:translateY(-50%);font-size:12px;color:var(--text-3);">R$</span>
          <input type="number" step="0.01" id="conf-val-${t.id}" value="${Number(t.amount || 0).toFixed(2)}" style="width:120px;padding:6px 8px 6px 30px;border:1.5px solid var(--border);border-radius:8px;font-size:13px;background:var(--surface);color:var(--text);"></div>
        <button onclick="confirmarConta('${t.id}')" style="padding:7px 12px;border-radius:8px;border:none;background:#059669;color:white;font-size:12px;font-weight:700;cursor:pointer;">Confirmar</button>
        <button onclick="pularConta('${t.id}')" title="Remove só esta ocorrência" style="padding:7px 10px;border-radius:8px;border:1px solid var(--border);background:var(--surface);color:var(--text-3);font-size:12px;cursor:pointer;">Não veio</button>
      </div>`;
    }).join('') + '</div>';
}
function confirmarConta(id) {
  const t = S.transactions.find(x => x.id === id); if (!t) return;
  const v = parseFloat((document.getElementById('conf-val-' + id) || {}).value);
  if (v > 0) t.amount = v;
  t.pago = true; t.updatedAt = new Date().toISOString();
  save(); renderDashboard();
  toast(`✅ ${t.desc} confirmada: ${brl(t.amount)}`);
}
function confirmarTodas(ym) {
  const pend = txByTitular(S.transactions).filter(t => t.type === 'despesa' && t.pago === false && t.formaPgto !== 'credito' && !t.isTransfer && !isPgtoFatura(t) && (t.date || '').startsWith(ym));
  pend.forEach(t => { const v = parseFloat((document.getElementById('conf-val-' + t.id) || {}).value); if (v > 0) t.amount = v; t.pago = true; t.updatedAt = new Date().toISOString(); });
  save(); renderDashboard();
  toast(`✅ ${pend.length} conta${pend.length !== 1 ? 's' : ''} confirmada${pend.length !== 1 ? 's' : ''}`);
}
function pularConta(id) {
  const idx = S.transactions.findIndex(x => x.id === id); if (idx < 0) return;
  const t = S.transactions[idx];
  S.deletedIds.push({ id, collection: 'transactions', deletedAt: new Date().toISOString() });
  S.transactions.splice(idx, 1);
  save(); renderDashboard();
  toast(`🗑️ ${t.desc} removida deste mês`);
}

// ─── PROJEÇÃO DE CAIXA (3 meses) ──────────────────────────────────────────────
// Estimativa simples: entradas = média das receitas dos últimos 3 meses; saídas = o que já está
// lançado no mês (recorrentes, parcelas) + média do gasto "espontâneo" (não recorrente, não parcela)
// dos últimos 3 meses. Compras no cartão contam pela data da compra.
function renderProjecao(ym) {
  const el = document.getElementById('dash-projecao');
  if (!el) return;
  const hojeYM = new Date().toISOString().slice(0, 7);
  if (ym !== hojeYM) { el.style.display = 'none'; return; }
  const my = txByTitular(S.transactions);
  const real = t => !t.isTransfer && !isPgtoFatura(t);
  const addM = (k, n) => { const d = new Date(parseInt(k.slice(0, 4)), parseInt(k.slice(5, 7)) - 1 + n, 1); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`; };
  const soma = (fn, k) => my.filter(t => (t.date || '').startsWith(k) && real(t) && fn(t)).reduce((s, t) => s + amountBrl(t), 0);
  const isEspont = t => t.type === 'despesa' && !t.recorrente && !t.parcela && !t.parcelaTotal;
  // Base "típica" dos últimos 6 meses: mediana, descartando meses fora do padrão (venda de imóvel,
  // gasto grande de uma vez). Média simples inflava tudo.
  const tipico = arr => {
    const v = arr.filter(x => x > 0).sort((a, b) => a - b);
    if (!v.length) return 0;
    const med = v[Math.floor(v.length / 2)];
    const ok = v.filter(x => x <= med * 2.5 && x >= med * 0.3);
    return ok.reduce((a, b) => a + b, 0) / ok.length;
  };
  const recs = [], esps = [];
  for (let i = 1; i <= 6; i++) { const k = addM(ym, -i); recs.push(soma(t => t.type === 'receita', k)); esps.push(soma(isEspont, k)); }
  const n = recs.filter((r, i) => r > 0 || esps[i] > 0).length;
  if (!n) { el.style.display = 'none'; return; }
  const recAuto = tipico(recs);
  const rendaManual = parseFloat(S.settings.rendaMensal) || 0;   // ajuste manual (campo no card)
  const recM = rendaManual > 0 ? rendaManual : recAuto;
  const espM = tipico(esps);
  // saldo atual em contas (não cartão): receitas - despesas lançadas até hoje
  const contas = S.accounts.filter(a => a.accountType !== 'cartao' && (titularFilter === 'ambos' || a.owner === (titularFilter === 'paulo' ? S.settings.u1 : S.settings.u2)));
  const hoje = new Date().toISOString().slice(0, 10);
  let saldo = 0;
  contas.forEach(a => { S.transactions.filter(t => t.accountId === a.id && (t.date || '') <= hoje).forEach(t => { saldo += t.type === 'receita' ? amountBrl(t) : (t.type === 'despesa' ? -amountBrl(t) : 0); }); });
  const rows = [];
  let acum = saldo;
  for (let i = 1; i <= 3; i++) {
    const k = addM(ym, i);
    const lancado = soma(t => t.type === 'despesa', k);
    const recLanc = soma(t => t.type === 'receita', k);
    const entradas = Math.max(recLanc, recM);
    const saidas = lancado + espM;
    const res = entradas - saidas;
    acum += res;
    const tag = res < 0 ? ['Negativo', '#e11d48', 'var(--tint-rose)'] : (res < entradas * 0.1 ? ['Apertado', '#d97706', 'var(--tint-amber)'] : ['Folga', '#059669', 'var(--tint-green)']);
    rows.push({ k, entradas, lancado, esp: espM, saidas, res, acum, tag });
  }
  el.style.display = 'block';
  el.innerHTML = `<div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px;margin-bottom:10px;">
      <span style="font-size:14px;font-weight:700;color:var(--text);">🔮 Projeção de caixa · próximos 3 meses</span>
      <span style="font-size:11px;color:var(--muted);display:flex;align-items:center;gap:6px;flex-wrap:wrap;">saldo em conta hoje ${brl(saldo)} · entradas/mês:
        <span style="position:relative;display:inline-block;"><span style="position:absolute;left:8px;top:50%;transform:translateY(-50%);font-size:11px;">R$</span>
        <input type="number" step="100" value="${Math.round(recM)}" onchange="S.settings.rendaMensal = parseFloat(this.value) || 0; save(); renderProjecao('${ym}')" title="Entradas mensais esperadas. Vazio = mediana dos últimos 6 meses (${brl(recAuto)}), ignorando meses fora do padrão" style="width:110px;padding:3px 6px 3px 26px;border:1px solid var(--border);border-radius:6px;font-size:11px;background:var(--surface);color:var(--text);"></span>
        ${rendaManual > 0 ? `<a href="#" onclick="S.settings.rendaMensal = 0; save(); renderProjecao('${ym}'); return false;" style="color:var(--text-3);">usar automático (${brl(recAuto)})</a>` : '<span>(automático: mediana de 6 meses)</span>'}
        · variável estimado = mediana dos gastos não recorrentes</span></div>
    <div style="overflow-x:auto;"><table style="width:100%;border-collapse:collapse;font-size:13px;">
      <thead><tr style="color:var(--muted);font-size:11px;text-align:right;"><th style="text-align:left;padding:6px 8px;">Mês</th><th style="padding:6px 8px;">Entradas</th><th style="padding:6px 8px;">Já lançado</th><th style="padding:6px 8px;">Variável estimado</th><th style="padding:6px 8px;">Resultado</th><th style="padding:6px 8px;">Saldo projetado</th><th style="padding:6px 8px;text-align:center;">Cenário</th></tr></thead>
      <tbody>` + rows.map(r => `<tr style="border-top:1px solid var(--surface-2);text-align:right;">
        <td style="text-align:left;padding:8px;font-weight:600;color:var(--text);">${MESES_FULL[parseInt(r.k.slice(5, 7)) - 1]} ${r.k.slice(0, 4)}</td>
        <td style="padding:8px;color:#059669;">${brl(r.entradas)}</td><td style="padding:8px;color:var(--text-2);">${brl(r.lancado)}</td><td style="padding:8px;color:var(--text-3);">${brl(r.esp)}</td>
        <td style="padding:8px;font-weight:700;color:${r.res < 0 ? '#e11d48' : '#059669'};">${r.res < 0 ? '-' : '+'}${brl(Math.abs(r.res))}</td>
        <td style="padding:8px;font-weight:700;color:${r.acum < 0 ? '#e11d48' : 'var(--text)'};">${brl(r.acum)}</td>
        <td style="padding:8px;text-align:center;"><span style="font-size:11px;font-weight:700;padding:2px 10px;border-radius:20px;background:${r.tag[2]};color:${r.tag[1]};">${r.tag[0]}</span></td>
      </tr>`).join('') + `</tbody></table></div>`;
}

function renderRecent() {
  // Only show non-credit transactions in bank view, filtered by titular
  const recent = txByTitular([...S.transactions])
    .filter(t => !(t.type === 'despesa' && t.formaPgto === 'credito'))
    .sort((a,b)=>new Date(b.date)-new Date(a.date)).slice(0,7);
  const el = document.getElementById('recent-list');
  if (!recent.length) {
    el.innerHTML = `<div class="empty-state"><div class="icon">📭</div><p style="font-weight:600;color:var(--text-3);">Nenhuma transação ainda</p><button onclick="goto('nova')" style="margin-top:12px;padding:10px 20px;background:#4f46e5;color:white;border:none;border-radius:10px;font-size:13px;font-weight:600;cursor:pointer;">Adicionar primeira</button></div>`;
    return;
  }
  el.innerHTML = recent.map(t => {
    const cfg = { receita:{ic:'<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#059669" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="19" x2="12" y2="5"/><polyline points="5 12 12 5 19 12"/></svg>',clr:'#059669',sign:'+'}, despesa:{ic:'<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#e11d48" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="5" x2="12" y2="19"/><polyline points="19 12 12 19 5 12"/></svg>',clr:'#e11d48',sign:'-'}, investimento:{ic:'🐷',clr:'#2563eb',sign:'+'} }[t.type];
    return `<div class="trow" style="display:flex;align-items:center;padding:14px 24px;border-bottom:1px solid var(--bg);">
      <div style="width:38px;height:38px;background:var(--bg);border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:18px;margin-right:14px;">${cfg.ic}</div>
      <div style="flex:1;">
        <p style="font-size:14px;font-weight:600;color:var(--text);">${escapeHtml(t.desc)}</p>
        <p style="font-size:12px;color:var(--muted);margin-top:2px;">${t.category} · ${fmtDate(t.date)} · ${t.user}</p>
      </div>
      <p style="font-size:15px;font-weight:800;color:${cfg.clr};">${cfg.sign}${brl(t.currency==='USD'?usdToBrl(t.amount):t.amount)}</p>
    </div>`;
  }).join('');
}

// ─── CREDIT VIEW ──────────────────────────────────────────────────────────────
function renderCreditView() {
  const isAno = dashPeriodMode === 'ano';
  const faturaYM = `${dashYear}-${String(dashMonth + 1).padStart(2, '0')}`;

  // Get all credit cards
  const cartoes = S.accounts.filter(a => a.accountType === 'cartao');
  const cartoesFiltered = titularFilter === 'ambos'
    ? cartoes
    : cartoes.filter(a => a.owner === (titularFilter === 'paulo' ? S.settings.u1 : S.settings.u2));

  // Credit transactions: all faturas of the year or specific month
  let creditTx;
  if (isAno) {
    creditTx = txByTitular(S.transactions.filter(t =>
      t.type === 'despesa' && t.formaPgto === 'credito' &&
      getTxFaturaRef(t).startsWith(String(dashYear))
    ));
  } else {
    creditTx = txByTitular(getCreditTxByFatura(faturaYM));
  }

  // Calculate totals
  const limiteTotal = cartoesFiltered.reduce((s, a) => s + Number(a.limite || 0), 0);
  const limiteUsado = creditTx.reduce((s, t) => s + amountBrl(t), 0);
  const limiteDisp  = limiteTotal - limiteUsado;

  // Fatura atual = total of credit expenses for current billing cycle
  const faturaAtual = limiteUsado;

  document.getElementById('cc-limite-total').textContent = brl(limiteTotal);
  document.getElementById('cc-limite-usado').textContent = brl(limiteUsado);
  document.getElementById('cc-limite-disp').textContent  = brl(limiteDisp);
  document.getElementById('cc-limite-disp').style.color  = limiteDisp >= 0 ? '#059669' : '#e11d48';
  document.getElementById('cc-fatura-atual').textContent = brl(faturaAtual);

  // Year selector for faturas chart
  const yearSel = document.getElementById('cc-fatura-year');
  const curYear = dashYear;
  const years = new Set();
  S.transactions.forEach(t => { if (t.type === 'despesa' && t.formaPgto === 'credito') years.add(parseInt(t.date.substring(0,4))); });
  years.add(curYear);
  const sortedYears = [...years].sort((a,b) => b - a);
  const prevVal = yearSel.value;
  yearSel.innerHTML = sortedYears.map(y => `<option value="${y}">${y}</option>`).join('');
  yearSel.value = prevVal && sortedYears.includes(parseInt(prevVal)) ? prevVal : curYear;

  // Card filter tabs
  const tabsEl = document.getElementById('cc-card-tabs');
  tabsEl.innerHTML = `<button onclick="ccCardFilter='';renderCreditView()" class="tf-btn ${ccCardFilter==='' ? 'tf-active' : ''}" style="font-size:12px;">Todos</button>` +
    cartoesFiltered.map(c => {
      const b = BANKS[c.bank];
      const active = ccCardFilter === c.id;
      return `<button onclick="ccCardFilter='${c.id}';renderCreditView()" class="tf-btn ${active ? 'tf-active' : ''}" style="font-size:12px;">${escapeHtml(b.name)} Cartão</button>`;
    }).join('');

  // Render faturas bar chart
  renderFaturasChart(cartoesFiltered);

  // Populate card filter for category breakdown
  const catFilter = document.getElementById('cc-cat-card-filter');
  catFilter.innerHTML = `<option value="">Todos os cartões</option>` +
    cartoesFiltered.map(c => `<option value="${c.id}">${BANKS[c.bank].name} Cartão</option>`).join('');

  // Category breakdown
  renderCreditCatBreakdown();

  // Fixo vs Variável
  renderFixoVariavel(creditTx);

  // Parcelas na fatura
  renderParcelasFatura(creditTx);

  // Recent credit transactions
  renderCreditRecent();
}

function renderFixoVariavel(creditTx) {
  let fixo = 0, variavel = 0;
  creditTx.forEach(t => {
    const tipo = t.custoTipo || autoCustoTipo(t.category);
    const val = amountBrl(t);
    if (tipo === 'fixo') fixo += val;
    else variavel += val;
  });
  const total = fixo + variavel;
  const fixoPct = total > 0 ? (fixo / total * 100) : 0;
  const varPct = total > 0 ? (variavel / total * 100) : 0;
  document.getElementById('cc-fixo-val').textContent = brl(fixo);
  document.getElementById('cc-var-val').textContent = brl(variavel);
  document.getElementById('cc-fixo-bar').style.width = fixoPct + '%';
  document.getElementById('cc-var-bar').style.width = varPct + '%';
  document.getElementById('cc-fixo-pct').textContent = total > 0 ? `Fixo ${fixoPct.toFixed(0)}% · Variável ${varPct.toFixed(0)}%` : '';
}

function renderParcelasFatura(creditTx) {
  const parcelas = creditTx.filter(t => t.parcela || (t.desc && /\(\d+\/\d+\)/.test(t.desc)));
  const el = document.getElementById('cc-parcelas-list');
  const totalEl = document.getElementById('cc-parcelas-total');

  if (!parcelas.length) {
    el.innerHTML = '<p style="font-size:13px;color:var(--muted);text-align:center;padding:10px;">Nenhuma parcela nesta fatura</p>';
    totalEl.textContent = 'R$ 0,00';
    return;
  }

  const totalParcelas = parcelas.reduce((s, t) => s + amountBrl(t), 0);
  totalEl.textContent = brl(totalParcelas);

  el.innerHTML = parcelas.map(t => {
    const acc = S.accounts.find(a => a.id === t.accountId);
    const bankName = acc ? BANKS[acc.bank]?.name || '' : '';
    const parcelaInfo = t.parcela || (t.desc.match(/\((\d+\/\d+)\)/)?.[1] || '');
    const descClean = t.desc.replace(/\s*\(?\d+\/\d+\)?\s*$/, '').replace(/\s*\(\(\d+\/\d+\)\)\s*$/, '');
    return `<div style="display:flex;justify-content:space-between;align-items:center;padding:8px 12px;background:var(--bg);border-radius:8px;">
      <div>
        <p style="font-size:13px;font-weight:600;color:var(--text);">${escapeHtml(descClean)}</p>
        <p style="font-size:11px;color:var(--muted);">${bankName} · ${fmtDate(t.date)} · <span style="color:#4f46e5;font-weight:600;">${parcelaInfo}</span></p>
      </div>
      <p style="font-size:13px;font-weight:700;color:#e11d48;">${brl(amountBrl(t))}</p>
    </div>`;
  }).join('');
}

function renderFaturasChart(cartoes) {
  const selYear = parseInt(document.getElementById('cc-fatura-year')?.value || new Date().getFullYear());
  const labels = MESES.slice();
  const datasets = [];

  const cardsToShow = ccCardFilter ? cartoes.filter(c => c.id === ccCardFilter) : cartoes;

  const cardColors = {};
  cardsToShow.forEach((c, i) => {
    cardColors[c.id] = BANKS[c.bank]?.color || COLORS[i];
  });

  cardsToShow.forEach(card => {
    const data = [];
    for (let m = 0; m < 12; m++) {
      const k = `${selYear}-${String(m+1).padStart(2,'0')}`;
      const total = getCreditTxByFatura(k, card.id)
        .reduce((s, t) => s + amountBrl(t), 0);
      data.push(total);
    }
    datasets.push({
      label: BANKS[card.bank].name + ' Cartão',
      data,
      backgroundColor: cardColors[card.id],
      borderRadius: 4
    });
  });

  destroyChart('ch-cc-faturas');
  const ctx = document.getElementById('ch-cc-faturas');
  if (!ctx) return;

  charts['ch-cc-faturas'] = new Chart(ctx, {
    type: 'bar',
    data: { labels, datasets },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { position: 'top', labels: { boxWidth: 12, font: { size: 11 } } },
        tooltip: { callbacks: { label: c => ` ${c.dataset.label}: ${brl(c.raw)}` } }
      },
      scales: {
        x: { grid: { display: false } },
        y: { grid: { color: cssVar('--chart-grid') }, ticks: { callback: v => v === 0 ? 'R$0' : 'R$' + v.toLocaleString('pt-BR') } }
      }
    }
  });
}

function setCcCatMode(mode) {
  ccCatMode = mode;
  ccCatDrillCat = null;
  document.getElementById('cc-cat-mode-cat').style.cssText = `padding:3px 10px;border-radius:6px;border:none;font-size:11px;font-weight:600;cursor:pointer;background:${mode==='cat'?'#4f46e5':'transparent'};color:${mode==='cat'?'white':'var(--text-3)'};`;
  document.getElementById('cc-cat-mode-sub').style.cssText = `padding:3px 10px;border-radius:6px;border:none;font-size:11px;font-weight:600;cursor:pointer;background:${mode==='sub'?'#4f46e5':'transparent'};color:${mode==='sub'?'white':'var(--text-3)'};`;
  document.getElementById('cc-cat-back').style.display = 'none';
  renderCreditCatBreakdown();
}

function ccCatBack() {
  ccCatDrillCat = null;
  document.getElementById('cc-cat-back').style.display = 'none';
  renderCreditCatBreakdown();
}

function ccCatDrill(cat) {
  ccCatMode = 'sub';
  ccCatDrillCat = cat;
  document.getElementById('cc-cat-mode-cat').style.cssText = `padding:3px 10px;border-radius:6px;border:none;font-size:11px;font-weight:600;cursor:pointer;background:transparent;color:var(--text-3);`;
  document.getElementById('cc-cat-mode-sub').style.cssText = `padding:3px 10px;border-radius:6px;border:none;font-size:11px;font-weight:600;cursor:pointer;background:#4f46e5;color:white;`;
  document.getElementById('cc-cat-back').style.display = 'block';
  document.getElementById('cc-cat-back').querySelector('button').textContent = `← ${cat}`;
  renderCreditCatBreakdown();
}

function renderCreditCatBreakdown() {
  const isAno = dashPeriodMode === 'ano';
  const faturaYM = `${dashYear}-${String(dashMonth + 1).padStart(2, '0')}`;
  const cardFilter = document.getElementById('cc-cat-card-filter')?.value || '';

  let txs;
  if (isAno) {
    txs = txByTitular(S.transactions.filter(t =>
      t.type === 'despesa' && t.formaPgto === 'credito' &&
      getTxFaturaRef(t).startsWith(String(dashYear))
    ));
  } else {
    txs = txByTitular(getCreditTxByFatura(faturaYM));
  }
  if (cardFilter) txs = txs.filter(t => t.accountId === cardFilter);

  const by = {};
  if (ccCatMode === 'sub' && ccCatDrillCat) {
    // Drill-down: subcategorias de uma categoria
    const catTxs = txs.filter(t => t.category === ccCatDrillCat);
    catTxs.forEach(t => {
      const key = t.subcategory || 'Geral';
      by[key] = (by[key] || 0) + amountBrl(t);
    });
  } else if (ccCatMode === 'sub') {
    // Todas subcategorias
    txs.forEach(t => {
      const key = t.subcategory ? `${t.category} > ${t.subcategory}` : t.category;
      by[key] = (by[key] || 0) + amountBrl(t);
    });
  } else {
    txs.forEach(t => by[t.category] = (by[t.category] || 0) + amountBrl(t));
  }

  const sorted = Object.entries(by).sort((a, b) => b[1] - a[1]);
  const total = sorted.reduce((s, [, v]) => s + v, 0);

  document.getElementById('cc-cat-total').textContent = `Total ${brl(total)}`;

  const el = document.getElementById('cc-cat-list');
  if (!sorted.length) {
    el.innerHTML = '<p style="font-size:13px;color:var(--muted);text-align:center;padding:20px;">Nenhum gasto no crédito este mês</p>';
    return;
  }

  el.innerHTML = sorted.slice(0, 12).map(([cat, val], i) => {
    const pct = total > 0 ? (val / total) * 100 : 0;
    const clickable = ccCatMode === 'cat';
    return `<div style="display:flex;align-items:center;gap:12px;${clickable?'cursor:pointer;':''}" ${clickable ? `onclick="ccCatDrill('${cat.replace(/'/g,"\\'")}')"` : ''}>
      <div style="width:10px;height:10px;border-radius:50%;background:${COLORS[i % COLORS.length]};flex-shrink:0;"></div>
      <div style="flex:1;">
        <div style="display:flex;justify-content:space-between;margin-bottom:4px;">
          <span style="font-size:13px;font-weight:600;color:var(--text-2);">${escapeHtml(cat)}</span>
          <span style="font-size:13px;font-weight:700;color:var(--text-2);">${brl(val)}</span>
        </div>
        <div style="height:6px;background:var(--surface-2);border-radius:3px;overflow:hidden;">
          <div style="height:100%;width:${pct}%;background:${COLORS[i % COLORS.length]};border-radius:3px;"></div>
        </div>
      </div>
    </div>`;
  }).join('');
}

function renderCreditRecent() {
  const isAno = dashPeriodMode === 'ano';
  const faturaYM = `${dashYear}-${String(dashMonth + 1).padStart(2, '0')}`;
  let recent;
  if (isAno) {
    recent = txByTitular(S.transactions.filter(t =>
      t.type === 'despesa' && t.formaPgto === 'credito' &&
      getTxFaturaRef(t).startsWith(String(dashYear))
    )).sort((a, b) => new Date(b.date) - new Date(a.date)).slice(0, 7);
  } else {
    recent = txByTitular(getCreditTxByFatura(faturaYM))
      .sort((a, b) => new Date(b.date) - new Date(a.date)).slice(0, 7);
  }
  const el = document.getElementById('cc-recent-list');
  if (!recent.length) {
    el.innerHTML = `<div class="empty-state"><div class="icon">💳</div><p style="font-weight:600;color:var(--text-3);">Nenhuma compra no crédito</p></div>`;
    return;
  }
  el.innerHTML = recent.map(t => {
    const acc = S.accounts.find(a => a.id === t.accountId);
    const bankName = acc ? BANKS[acc.bank]?.name || '' : '';
    return `<div class="trow" style="display:flex;align-items:center;padding:14px 24px;border-bottom:1px solid var(--bg);">
      <div style="width:38px;height:38px;background:var(--bg);border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:18px;margin-right:14px;">💳</div>
      <div style="flex:1;">
        <p style="font-size:14px;font-weight:600;color:var(--text);">${escapeHtml(t.desc)}</p>
        <p style="font-size:12px;color:var(--muted);margin-top:2px;">${t.category} · ${fmtDate(t.date)} · ${bankName} · ${t.user}</p>
      </div>
      <p style="font-size:15px;font-weight:800;color:${t.isNegative ? '#059669' : '#e11d48'};">${t.isNegative ? '+' : '-'}${brl(t.amount)}</p>
    </div>`;
  }).join('');
}
