const CACHE_NAME = 'twm-datalink-v9';
const ASSETS = [
  './',
  './index.html',
  './solicitacao_compras.html',
  './check_list_motoristas.html',
  './checklist_lvt.html',
  './checklist_seguranca_mensal.html',
  './pneus.html',
  './estoque.html',
  './carros-apoio.html',
  './lavagem.html',
  './abastecimento.html',
  './manutencao.html',
  './relatorios.html',
  './manifest.json',
  './icone-192.png',
  './icone-512.png',
  './favicon.png'
];

// Instalação: faz cache dos arquivos principais
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => cache.addAll(ASSETS))
  );
  self.skipWaiting();
});

// Ativação: remove caches antigos
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

// Fetch: network first para HTML, cache first para o resto
self.addEventListener('fetch', event => {
  const url = event.request.url;

  // Apps Script sempre vai para a rede
  if (url.includes('script.google.com') || url.includes('imgbb.com')) {
    event.respondWith(
      fetch(event.request).catch(() =>
        new Response('{}', { headers: { 'Content-Type': 'application/json' } })
      )
    );
    return;
  }

  // HTMLs: network first (garante sempre versão atualizada)
  if (url.endsWith('.html') || url.endsWith('/')) {
    event.respondWith(
      fetch(event.request)
        .then(response => {
          if (response && response.status === 200) {
            const clone = response.clone();
            caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
          }
          return response;
        })
        .catch(() => caches.match(event.request).then(cached => cached || caches.match('./index.html')))
    );
    return;
  }

  // Demais assets (imagens, ícones): cache first
  event.respondWith(
    caches.match(event.request).then(cached => {
      if (cached) return cached;
      return fetch(event.request).then(response => {
        if (response && response.status === 200 && response.type === 'basic') {
          const clone = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
        }
        return response;
      }).catch(() => caches.match('./index.html'));
    })
  );
});
