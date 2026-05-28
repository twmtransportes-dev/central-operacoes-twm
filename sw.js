// TWM DataLink — Service Worker v202605281928
// Estratégia: Network First para HTML, Stale-While-Revalidate para assets
const CACHE_NAME = 'twm-v202605281928';
const ASSETS_CACHE = 'twm-assets-v202605281928';

self.addEventListener('install', event => {
  // Ativa imediatamente, sem esperar o SW antigo fechar
  self.skipWaiting();
  event.waitUntil(
    caches.open(ASSETS_CACHE).then(cache => {
      // Pré-cache apenas fontes do Google (estáticas, não mudam)
      return cache.addAll([
        'https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap'
      ]).catch(() => {});
    })
  );
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys => Promise.all(
      // Apaga TODOS os caches de versões antigas
      keys.filter(k => k !== ASSETS_CACHE && k !== CACHE_NAME)
          .map(k => caches.delete(k))
    )).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', event => {
  const url = new URL(event.request.url);
  const isLocal = url.origin === location.origin;
  
  // Recursos externos (Google Fonts, CDN) — Cache First
  if (!isLocal) {
    event.respondWith(
      caches.match(event.request)
        .then(cached => cached || fetch(event.request)
          .then(resp => {
            if (resp.ok) {
              const clone = resp.clone();
              caches.open(ASSETS_CACHE).then(c => c.put(event.request, clone));
            }
            return resp;
          })
        )
    );
    return;
  }

  // Arquivos HTML locais — Network First (sempre tenta servidor)
  // Se offline, usa cache como fallback
  if (url.pathname.endsWith('.html') || url.pathname === '/' || url.pathname.endsWith('/')) {
    event.respondWith(
      fetch(event.request, { cache: 'no-store' })
        .then(resp => {
          // Salva no cache para fallback offline
          const clone = resp.clone();
          caches.open(CACHE_NAME).then(c => c.put(event.request, clone));
          return resp;
        })
        .catch(() => caches.match(event.request))
    );
    return;
  }

  // JS, CSS, imagens locais — Stale While Revalidate
  // Responde com cache imediatamente + atualiza em background
  event.respondWith(
    caches.open(CACHE_NAME).then(cache => {
      return cache.match(event.request).then(cached => {
        const networkFetch = fetch(event.request, { cache: 'no-store' })
          .then(resp => {
            if (resp.ok) cache.put(event.request, resp.clone());
            return resp;
          })
          .catch(() => cached);
        
        // Se tem cache: responde imediatamente + atualiza em background
        // Se não tem: espera a rede
        return cached || networkFetch;
      });
    })
  );
});

// Notifica todos os clientes quando há atualização
self.addEventListener('message', event => {
  if (event.data === 'SKIP_WAITING') self.skipWaiting();
});
