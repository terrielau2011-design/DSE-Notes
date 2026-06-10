const CACHE_NAME = 'dse-notes-v29';
const CACHE_URLS = [
  '/DSE-Notes/',
  '/DSE-Notes/index.html',
  '/DSE-Notes/manifest.json'
];

// Install - cache core files
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => {
      return cache.addAll(CACHE_URLS);
    })
  );
  self.skipWaiting();
});

// Activate - clean old caches, take control immediately
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys => {
      return Promise.all(
        keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k))
      );
    })
  );
  self.clients.claim();
});

// Handle skip_waiting message from page
self.addEventListener('message', event => {
  if (event.data && event.data.type === 'SKIP_WAITING') {
    self.skipWaiting();
  }
});

// Fetch - network first for HTML (ensures updates), cache first for static assets
self.addEventListener('fetch', event => {
  // Skip non-GET requests
  if (event.request.method !== 'GET') return;

  const url = new URL(event.request.url);
  if (url.hostname !== 'terrielau2011-design.github.io') return;

  // For HTML pages: network first (always check for updates)
  const isHTML = event.request.headers.get('accept')?.includes('text/html') ||
                 url.pathname.endsWith('/') ||
                 url.pathname.endsWith('.html');

  if (isHTML) {
    event.respondWith(
      fetch(event.request, { cache: 'no-cache' })
        .then(response => {
          if (response.ok) {
            const clone = response.clone();
            caches.open(CACHE_NAME).then(cache => {
              cache.put(event.request, clone);
            });
          }
          return response;
        })
        .catch(() => {
          return caches.match(event.request).then(cached => {
            return cached || caches.match('/DSE-Notes/index.html');
          });
        })
    );
  } else {
    // For static assets (CSS, JS, images): cache first, then network
    event.respondWith(
      caches.match(event.request).then(cached => {
        if (cached) return cached;
        return fetch(event.request).then(response => {
          if (response.ok) {
            const clone = response.clone();
            caches.open(CACHE_NAME).then(cache => {
              cache.put(event.request, clone);
            });
          }
          return response;
        });
      })
    );
  }
});
