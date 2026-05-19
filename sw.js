// TWM DataLink — Service Worker
// Sempre busca do servidor, sem cache de HTML
// Atualiza automaticamente a cada deploy no GitHub

self.addEventListener('install', event => {
  self.skipWaiting();
});

self.addEventListener('activate', event => {
  // Apaga todos os caches antigos
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', event => {
  const url = new URL(event.request.url);

  // Arquivos HTML: SEMPRE busca do servidor, sem cache
  if (url.pathname.endsWith('.html') || url.pathname === '/' || url.pathname === '') {
    event.respondWith(
      fetch(event.request, { cache: 'no-store' })
        .catch(() => caches.match(event.request))
    );
    return;
  }

  // JS, CSS, fontes: Network First com cache de fallback
  event.respondWith(
    fetch(event.request, { cache: 'no-store' })
      .then(response => {
        const clone = response.clone();
        caches.open('twm-assets').then(cache => cache.put(event.request, clone));
        return response;
      })
      .catch(() => caches.match(event.request))
  );
});
