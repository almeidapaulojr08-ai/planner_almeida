// FinançasCasal — 13-boot.js
// Módulo carregado em ordem por index.html (escopo global compartilhado). Não use import/export.

// ─── ATALHOS DE TECLADO ───────────────────────────────────────────────────────
//  N          → Nova transação (fora de campos de texto)
//  Ctrl+Enter → salva o formulário de nova transação
//  Esc        → fecha o modal aberto
// Fecha o modal aberto (chama a função fecharXxxModal correspondente quando existe)
function closeAnyModal() {
  const special = { 'del-modal': 'closeModal', 'conta-modal': 'fecharModalConta', 'simul-amort-modal': 'fecharSimulModal', 'amort-table-modal': 'fecharAmortModal' };
  const open = [...document.querySelectorAll('div[id$="-modal"]')].filter(d => d.style.display !== 'none' && d.querySelector('#modal-backdrop'));
  if (!open.length) return false;
  const wrap = open[open.length - 1];
  const camel = wrap.id.replace(/-modal$/, '').replace(/-([a-z])/g, (_, c) => c.toUpperCase());
  const fnName = special[wrap.id] || ('fechar' + camel.charAt(0).toUpperCase() + camel.slice(1) + 'Modal');
  if (typeof window[fnName] === 'function') window[fnName]();
  else wrap.style.display = 'none';
  return true;
}
document.addEventListener('keydown', (ev) => {
  if (ev.key === 'Escape' && closeAnyModal()) { ev.preventDefault(); return; }
  if (!sessionStorage.getItem('fincasal_auth')) return;
  const tag = (ev.target.tagName || '').toLowerCase();
  const typing = tag === 'input' || tag === 'textarea' || tag === 'select' || ev.target.isContentEditable;
  // Enter/Espaço ativam itens da navegação focados por teclado
  if ((ev.key === 'Enter' || ev.key === ' ') && ev.target.classList && ev.target.classList.contains('nav-item')) { ev.preventDefault(); ev.target.click(); return; }
  if (ev.key === 'Enter' && (ev.ctrlKey || ev.metaKey)) {
    const form = document.querySelector('#page-nova form');
    const visible = document.getElementById('page-nova').style.display !== 'none';
    if (form && visible) { ev.preventDefault(); form.requestSubmit(); }
    return;
  }
  if (typing || ev.ctrlKey || ev.metaKey || ev.altKey) return;
  const pageVisible = id => document.getElementById(id) && document.getElementById(id).style.display !== 'none';
  if (ev.key === 'n' || ev.key === 'N') { ev.preventDefault(); goto('nova'); return; }
  if (ev.key === '/' ) { ev.preventDefault(); if (!pageVisible('page-historico')) goto('historico'); setTimeout(() => { const f = document.getElementById('fil-search'); if (f) { f.focus(); f.select(); } }, 60); return; }
  if ((ev.key === 'ArrowLeft' || ev.key === 'ArrowRight') && pageVisible('page-dashboard') && typeof dashNavMonth === 'function') {
    ev.preventDefault(); dashNavMonth(ev.key === 'ArrowLeft' ? -1 : 1);
  }
});

window.onload = function () {
  init();
  try { applyTheme(localStorage.getItem('fincasal_theme') || 'auto'); } catch (e) {}
};

// ─── PWA: service worker + atualização automática ─────────────────────────────
// sw.js busca o index.html sempre na rede (revalidando) e só usa cache offline.
// Quando sai versão nova, avisa e recarrega — ninguém fica preso na versão velha.
if ('serviceWorker' in navigator) {
  window.addEventListener('load', () => {
    const hadController = !!navigator.serviceWorker.controller;
    navigator.serviceWorker.register('./sw.js').then(reg => {
      const check = () => reg.update().catch(() => {});
      check();
      setInterval(check, 30 * 60 * 1000);                       // app fica aberto o dia todo
      document.addEventListener('visibilitychange', () => {     // voltou pro app no celular
        if (document.visibilityState === 'visible') check();
      });
      reg.addEventListener('updatefound', () => {
        const nw = reg.installing;
        if (!nw) return;
        nw.addEventListener('statechange', () => {
          if (nw.state === 'installed' && navigator.serviceWorker.controller) {
            if (typeof toast === 'function') toast('🔄 Nova versão disponível, atualizando...', 2500);
            nw.postMessage('SKIP_WAITING');
          }
        });
      });
    }).catch(e => console.warn('SW register failed:', e));

    // Deploy de js/css não muda o sw.js, então aba aberta o dia todo ficava no código velho
    // (ex.: Confirmar conta marcando pago depois da correção). Compara o carimbo ?v= da página
    // aberta com o do index.html publicado e avisa pra recarregar (não recarrega sozinho pra
    // não perder formulário preenchido).
    const meuCarimbo = ((document.querySelector('script[src*="13-boot.js"]') || {}).src || '').split('v=')[1];
    const checarVersao = () => {
      if (!meuCarimbo || document.getElementById('nova-versao-banner')) return;
      fetch('./?_=' + Date.now(), { cache: 'no-store' }).then(r => r.ok ? r.text() : '').then(html => {
        const pub = (html.match(/13-boot\.js\?v=(\d+)/) || [])[1];
        if (!pub || pub === meuCarimbo) return;
        const b = document.createElement('div');
        b.id = 'nova-versao-banner';
        b.style.cssText = 'position:fixed;top:12px;left:50%;transform:translateX(-50%);z-index:9999;background:#4f46e5;color:white;padding:10px 16px;border-radius:12px;font-size:13px;font-weight:600;box-shadow:0 6px 20px rgba(0,0,0,.25);cursor:pointer;';
        b.textContent = '🔄 O site foi atualizado. Clique aqui pra recarregar.';
        b.onclick = () => location.reload();
        document.body.appendChild(b);
      }).catch(() => {});
    };
    setInterval(checarVersao, 5 * 60 * 1000);
    document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible') checarVersao(); });
    window.addEventListener('focus', checarVersao);

    // Só recarrega quando TROCA de versão (não na primeira instalação)
    let reloaded = false;
    navigator.serviceWorker.addEventListener('controllerchange', () => {
      if (!hadController || reloaded) return;
      reloaded = true;
      setTimeout(() => location.reload(), 1500);
    });
  });
}
