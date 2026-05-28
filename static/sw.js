const CACHE_NAME = 'yss-anantapur-v1';
const RAILWAY_HOST = 'yssatp2026.up.railway.app';
const PROXY_HOST = 'holy-glitter-a694.224g1a3254.workers.dev';

const ASSETS_TO_CACHE = [
  '/',
  '/static/css/style.css?v=1.0.3',
  '/static/js/main.js',
  '/static/images/logo.png',
  '/static/images/hero.png',
  'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css',
  '/offline'
];

// Helper function to fetch with a timeout
function fetchWithTimeout(request, timeoutMs = 10000) {
  return new Promise((resolve, reject) => {
    const timeoutId = setTimeout(() => {
      reject(new Error('Network request timed out'));
    }, timeoutMs);

    fetch(request).then(
      (response) => {
        clearTimeout(timeoutId);
        resolve(response);
      },
      (err) => {
        clearTimeout(timeoutId);
        reject(err);
      }
    );
  });
}

// Install Event
self.addEventListener('install', (e) => {
  e.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      console.log('[Service Worker] Pre-caching offline page and assets');
      const cachePromises = ASSETS_TO_CACHE.map((asset) => {
        return cache.add(asset).catch((err) => {
          console.warn(`[Service Worker] Failed to pre-cache: ${asset}`, err);
          return null;
        });
      });
      return Promise.all(cachePromises);
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
  const url = new URL(e.request.url);

  // Skip browser extension requests or non-http protocols
  if (!url.protocol.startsWith('http')) return;

  // Intercept POST requests and proxy if direct connection fails
  if (e.request.method === 'POST') {
    if (url.hostname === RAILWAY_HOST) {
      e.respondWith(
        fetch(e.request.clone()).catch(async (err) => {
          console.warn('[Service Worker] POST failed. Trying proxy fallback...', err);
          const proxyUrl = new URL(e.request.url);
          proxyUrl.hostname = PROXY_HOST;
          
          try {
            const headers = new Headers(e.request.headers);
            const bodyBlob = await e.request.clone().blob();
            
            return fetch(new Request(proxyUrl.toString(), {
              method: 'POST',
              headers: headers,
              body: bodyBlob,
              redirect: 'manual'
            }));
          } catch (proxyErr) {
            console.error('[Service Worker] POST proxy fallback failed:', proxyErr);
            throw err;
          }
        })
      );
    }
    return;
  }

  // Only handle GET requests below
  if (e.request.method !== 'GET') return;

  // 1. Navigation requests: Network-First (with timeout), proxy fallback, cache fallback, offline fallback
  if (e.request.mode === 'navigate') {
    e.respondWith(
      fetchWithTimeout(e.request, 10000)
        .then((networkResponse) => {
          // Cache the successful home page load for offline use
          if (networkResponse.status === 200 && url.pathname === '/') {
            caches.open(CACHE_NAME).then((cache) => cache.put('/', networkResponse.clone()));
          }
          return networkResponse;
        })
        .catch(async (err) => {
          console.warn('[Service Worker] Navigation fetch failed. Trying proxy fallback...', err);
          
          const proxyUrl = new URL(e.request.url);
          proxyUrl.hostname = PROXY_HOST;
          
          try {
            const proxyResponse = await fetch(new Request(proxyUrl.toString(), {
              method: 'GET',
              headers: e.request.headers,
              redirect: 'manual'
            }));
            return proxyResponse;
          } catch (proxyErr) {
            console.warn('[Service Worker] Proxy navigation failed. Serving from cache/offline.', proxyErr);
            return caches.match(e.request).then((cachedResponse) => {
              if (cachedResponse) return cachedResponse;
              return caches.match('/offline');
            });
          }
        })
    );
    return;
  }

  // 2. Non-navigation requests: Stale-While-Revalidate with proxy fallback
  e.respondWith(
    caches.match(e.request).then((cachedResponse) => {
      if (cachedResponse) {
        // Fetch new version in the background
        fetch(e.request).then((networkResponse) => {
          if (networkResponse.status === 200) {
            caches.open(CACHE_NAME).then((cache) => cache.put(e.request, networkResponse));
          }
        }).catch(() => {
          // Ignore background failures
        });
        
        return cachedResponse;
      }

      // Not in cache, fetch from network
      return fetch(e.request).catch(async (networkErr) => {
        // Network failed (DNS block). Try proxy fallback!
        console.warn(`[Service Worker] Asset fetch failed for ${url.pathname}. Trying proxy fallback...`);
        const proxyUrl = new URL(e.request.url);
        proxyUrl.hostname = PROXY_HOST;
        
        try {
          return await fetch(proxyUrl);
        } catch (proxyErr) {
          console.error(`[Service Worker] Asset proxy fallback also failed for ${url.pathname}:`, proxyErr);
          throw networkErr;
        }
      });
    })
  );
});
