const CACHE_NAME = 'yss-anantapur-v1';
const ASSETS_TO_CACHE = [
  '/',
  '/static/css/style.css?v=1.0.3',
  '/static/js/main.js',
  '/static/images/logo.png',
  '/static/images/hero.png',
  'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css',
  '/offline'
];

// Install Event
self.addEventListener('install', (e) => {
  e.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      console.log('[Service Worker] Pre-caching offline page and assets');
      return cache.addAll(ASSETS_TO_CACHE);
    }).then(() => self.skipWaiting())
  );
});

// Activate Event
self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys().then((keys) => {
      return Promise.all(
        keys.map((key) => {
          if (key !== CACHE_NAME) {
            console.log('[Service Worker] Removing old cache', key);
            return caches.delete(key);
          }
        })
      );
    }).then(() => self.clients.claim())
  );
});

// Fetch Event
self.addEventListener('fetch', (e) => {
  // Only handle GET requests
  if (e.request.method !== 'GET') return;

  const url = new URL(e.request.url);

  // Skip browser extension requests or non-http protocols
  if (!url.protocol.startsWith('http')) return;

  // 1. Navigation requests: Network-First, fallback to Cache, fallback to Offline
  if (e.request.mode === 'navigate') {
    e.respondWith(
      fetch(e.request)
        .then((networkResponse) => {
          // Cache the successful home page load for offline use
          if (networkResponse.status === 200 && url.pathname === '/') {
            caches.open(CACHE_NAME).then((cache) => cache.put('/', networkResponse.clone()));
          }
          return networkResponse;
        })
        .catch(() => {
          // Network failed, try cache
          return caches.match(e.request).then((cachedResponse) => {
            if (cachedResponse) {
              return cachedResponse;
            }
            // Fallback to pre-cached offline page
            return caches.match('/offline');
          });
        })
    );
    return;
  }

  // 2. Non-navigation requests: Stale-While-Revalidate
  e.respondWith(
    caches.match(e.request).then((cachedResponse) => {
      if (cachedResponse) {
        // Fetch new version in the background
        fetch(e.request).then((networkResponse) => {
          if (networkResponse.status === 200) {
            caches.open(CACHE_NAME).then((cache) => cache.put(e.request, networkResponse));
          }
        }).catch(() => {/* Ignore network failures in background */});
        
        return cachedResponse;
      }

      // Not in cache, fetch from network
      return fetch(e.request);
    })
  );
});
