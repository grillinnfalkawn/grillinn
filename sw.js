// Grill Inn Falkawn — service worker
// Purpose: make the site installable ("Add to Home Screen") and keep it
// loading fast / working offline for the app shell. It deliberately does
// NOT cache calls to the order/settings API, so menu, hours, and orders
// are always fresh from your laptop dashboard.

const CACHE_NAME = 'grillinn-shell-v1';
const SCOPE = self.registration.scope; // e.g. https://grillinnfalkawn.github.io/grillinn/

const SHELL_FILES = [
  SCOPE,                       // index.html (start_url)
  SCOPE + 'manifest.json',
  SCOPE + 'icons/icon-192.png',
  SCOPE + 'icons/icon-512.png'
];

// Never touch these — always go straight to the network.
const NEVER_CACHE_HOST = 'orders.grillinnfalkawn.in';

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(SHELL_FILES))
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((names) =>
      Promise.all(
        names.filter((n) => n !== CACHE_NAME).map((n) => caches.delete(n))
      )
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  const req = event.request;

  // Only handle GET requests.
  if (req.method !== 'GET') return;

  const url = new URL(req.url);

  // Always bypass the cache for the live backend (settings, orders, etc.)
  if (url.hostname === NEVER_CACHE_HOST) return;

  // Network-first for the page itself, so customers always see the latest
  // menu/prices when online, with a cached fallback when offline.
  if (req.mode === 'navigate') {
    event.respondWith(
      fetch(req)
        .then((res) => {
          const copy = res.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(SCOPE, copy));
          return res;
        })
        .catch(() => caches.match(SCOPE))
    );
    return;
  }

  // Cache-first for same-origin static assets (icons, manifest, fonts, etc.)
  if (url.origin === self.location.origin) {
    event.respondWith(
      caches.match(req).then((cached) => {
        if (cached) return cached;
        return fetch(req).then((res) => {
          const copy = res.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(req, copy));
          return res;
        });
      })
    );
  }
});
