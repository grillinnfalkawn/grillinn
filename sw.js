// Grill Inn Falkawn — service worker.
// Strategy: NETWORK FIRST for everything, falling back to the last cached
// copy only when offline. This deliberately avoids the classic PWA trap of
// customers being stuck on a stale menu page: a live visit always shows the
// freshest deployed site, and the cache only matters when there's no
// connection at all.
// The ordering API (orders.grillinnfalkawn.in) is never intercepted or
// cached — placing an order must always hit the real server.

const CACHE_NAME = 'grillinn-shell-v1';

self.addEventListener('install', (event) => {
  self.skipWaiting();
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) =>
      cache.addAll(['./', './index.html', './manifest.json']).catch(() => {})
    )
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  const req = event.request;
  if (req.method !== 'GET') return; // never touch order submissions
  const url = new URL(req.url);
  if (url.origin !== self.location.origin) return; // never intercept the ordering API or other sites

  event.respondWith(
    fetch(req)
      .then((res) => {
        if (res && res.ok) {
          const copy = res.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(req, copy)).catch(() => {});
        }
        return res;
      })
      .catch(() =>
        caches.match(req).then((cached) => {
          if (cached) return cached;
          if (req.mode === 'navigate') return caches.match('./index.html');
          return Response.error();
        })
      )
  );
});
