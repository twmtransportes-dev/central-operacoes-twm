// ══════════════════════════════════════════════════════════════════
// TWM DataLink — Service Worker (versão consolidada)
// HTML: Network-First com cache:no-store (SEMPRE versão nova do servidor)
// Assets: Stale-While-Revalidate (rápido + atualiza em background)
// Dados (Apps Script/Sheets): sempre rede, nunca cache
// ══════════════════════════════════════════════════════════════════

// ⚠️ IMPORTANTE: mude este número a cada deploy para forçar atualização
const VERSION = 'v20260711-01';
const CACHE_HTML   = 'twm-html-'   + VERSION;
const CACHE_ASSETS = 'twm-assets-' + VERSION;

// ── INSTALL ──────────────────────────────────────────────────────
self.addEventListener('install', function(event) {
  self.skipWaiting(); // ativa a nova versão imediatamente
});

// ── ACTIVATE: apaga TODOS os caches antigos ──────────────────────
self.addEventListener('activate', function(event) {
  event.waitUntil(
    caches.keys().then(function(keys) {
      return Promise.all(
        keys.filter(function(k){ return k !== CACHE_HTML && k !== CACHE_ASSETS; })
            .map(function(k){ return caches.delete(k); })
      );
    }).then(function(){ return self.clients.claim(); })
  );
});

// ── FETCH ────────────────────────────────────────────────────────
self.addEventListener('fetch', function(event) {
  if (event.request.method !== 'GET') return;
  const url = new URL(event.request.url);
  const isLocal = url.origin === location.origin;

  // 1. DADOS DINÂMICOS (Apps Script, Google Sheets, ImgBB) → SEMPRE rede
  if (url.hostname.includes('script.google.com') ||
      url.hostname.includes('googleusercontent.com') ||
      url.hostname.includes('docs.google.com') ||
      url.hostname.includes('sheets.google.com') ||
      url.hostname.includes('imgbb.com')) {
    event.respondWith(
      fetch(event.request).catch(function() {
        return new Response('{}', { headers: { 'Content-Type': 'application/json' } });
      })
    );
    return;
  }

  // 2. RECURSOS EXTERNOS (Google Fonts, CDN) → Cache-First
  if (!isLocal) {
    event.respondWith(
      caches.match(event.request).then(function(cached) {
        return cached || fetch(event.request).then(function(resp) {
          if (resp.ok) {
            const clone = resp.clone();
            caches.open(CACHE_ASSETS).then(function(c){ c.put(event.request, clone); });
          }
          return resp;
        });
      })
    );
    return;
  }

  // 3. HTML LOCAL → Network-First com no-store (SEMPRE versão nova)
  if (url.pathname.endsWith('.html') || url.pathname === '/' || url.pathname.endsWith('/')) {
    event.respondWith(
      fetch(event.request, { cache: 'no-store' })
        .then(function(resp) {
          const clone = resp.clone();
          caches.open(CACHE_HTML).then(function(c){ c.put(event.request, clone); });
          return resp;
        })
        .catch(function() {
          // Offline → usa cache como fallback
          return caches.match(event.request).then(function(cached) {
            return cached || caches.match('./index.html');
          });
        })
    );
    return;
  }

  // 4. ASSETS LOCAIS (js, css, imagens) → Stale-While-Revalidate
  event.respondWith(
    caches.open(CACHE_ASSETS).then(function(cache) {
      return cache.match(event.request).then(function(cached) {
        const networkFetch = fetch(event.request, { cache: 'no-store' })
          .then(function(resp) {
            if (resp.ok) cache.put(event.request, resp.clone());
            return resp;
          })
          .catch(function(){ return cached; });
        return cached || networkFetch;
      });
    })
  );
});

// ── MENSAGEM: força ativação ─────────────────────────────────────
self.addEventListener('message', function(event) {
  if (event.data === 'SKIP_WAITING' || event.data === 'skipWaiting') {
    self.skipWaiting();
  }
});
