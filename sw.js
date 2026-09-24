/* Service worker do FinançasCasal (PWA)
 *
 * Estratégia:
 *  - index.html (navegação): REDE PRIMEIRO, sempre revalidando no servidor (cache: 'no-cache').
 *    Cache só entra se estiver offline. Isso evita alguém ficar preso numa versão velha
 *    (o que causa ping-pong de sync com quem está na versão nova).
 *  - Bibliotecas de CDN (Tailwind, Chart.js, Lottie, Firebase SDK) e arquivos estáticos
 *    do próprio site (ícones, anima-bot.json): CACHE PRIMEIRO, atualizando em segundo plano.
 *  - Qualquer outra coisa (Firebase Realtime/Auth, cotação do dólar, APIs de IA): não passa
 *    pelo cache — vai direto pra rede.
 *
 * Trocar CACHE_VERSION força limpar o cache antigo no próximo deploy.
 */
const CACHE_VERSION = 'fincasal-v3';
const OFFLINE_CACHE = CACHE_VERSION + '-shell';
const ASSET_CACHE = CACHE_VERSION + '-assets';

const CDN_HOSTS = ['cdn.tailwindcss.com', 'cdn.jsdelivr.net', 'cdnjs.cloudflare.com', 'www.gstatic.com', 'fonts.googleapis.com', 'fonts.gstatic.com'];

self.addEventListener('install', (event) => {
  // Já pré-carrega o app pra funcionar offline desde a primeira visita
  event.waitUntil(
    caches.open(OFFLINE_CACHE)
      .then(c => c.add(new Request('./', { cache: 'no-cache' })))
      .catch(() => {})
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => !k.startsWith(CACHE_VERSION)).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  const req = event.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);

  // 1. Navegação / index.html e os módulos do próprio app (js/, css/) → rede primeiro (revalida),
  //    cache se offline. Os módulos têm ?v=carimbo, então cada index.html puxa exatamente a sua versão.
  const isOwnModule = url.origin === self.location.origin && /\/(js|css)\/[^/]+\.(js|css)$/.test(url.pathname);
  if (isOwnModule) {
    event.respondWith(
      fetch(new Request(req, { cache: 'no-cache' }))
        .then(res => { if (res && res.ok) caches.open(ASSET_CACHE).then(c => c.put(req, res.clone())); return res; })
        .catch(() => caches.match(req))
    );
    return;
  }
  if (req.mode === 'navigate' || url.pathname.endsWith('/index.html')) {
    event.respondWith(
      fetch(new Request(req, { cache: 'no-cache' }))
        .then(res => {
          if (res && res.ok) {
            const copy = res.clone();
            caches.open(OFFLINE_CACHE).then(c => c.put('./', copy));
          }
          return res;
        })
        .catch(() => caches.match('./'))
    );
    return;
  }

  // 2. Bibliotecas de CDN e estáticos do site → cache primeiro, atualiza em segundo plano
  const isCdn = CDN_HOSTS.includes(url.hostname);
  const isOwnStatic = url.origin === self.location.origin &&
    /\.(png|json|webmanifest|ico|svg|woff2?)$/.test(url.pathname);
  if (isCdn || isOwnStatic) {
    event.respondWith(
      caches.open(ASSET_CACHE).then(cache =>
        cache.match(req).then(cached => {
          const network = fetch(req).then(res => {
            if (res && (res.ok || res.type === 'opaque')) cache.put(req, res.clone());
            return res;
          }).catch(() => cached);
          return cached || network;
        })
      )
    );
    return;
  }

  // 3. Resto (Firebase, APIs) → rede direta, sem cache
});

// A página pede pra ativar a versão nova na hora
self.addEventListener('message', (event) => {
  if (event.data === 'SKIP_WAITING') self.skipWaiting();
});
