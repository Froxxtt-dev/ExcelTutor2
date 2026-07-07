const CACHE_NAME = 'excel-tutor-shell-v4';
const SHELL_FILES = [
  './',
  './index.html',
  './styles.css',
  './app.js',
  './config.js',
  './formula-engine.js',
  './manifest.json',
  './icon-192.png',
  './icon-512.png'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(SHELL_FILES)).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((names) =>
      Promise.all(names.filter((n) => n !== CACHE_NAME).map((n) => caches.delete(n)))
    ).then(() => self.clients.claim())
  );
});

// Network-first for the Groq API (live data), cache-first for the app shell.
self.addEventListener('fetch', (event) => {
  const req = event.request;
  const url = new URL(req.url);

  // Only ever handle plain http(s) GET requests. Browser extensions (chrome-extension://),
  // POST/streaming calls, and other schemes should pass straight through untouched —
  // the Cache API can't store them anyway, and trying to throws.
  if (req.method !== 'GET' || !url.protocol.startsWith('http')) {
    return; // let the browser handle it natively, no event.respondWith()
  }

  if (url.origin.includes('api.groq.com') || url.origin !== self.location.origin) {
    // Never cache live tutoring calls or third-party CDNs (Chart.js, SheetJS, Google Fonts) —
    // just pass them straight to the network.
    event.respondWith(fetch(req));
    return;
  }

  event.respondWith(
    caches.match(req).then((cached) => {
      return (
        cached ||
        fetch(req)
          .then((response) => {
            const clone = response.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(req, clone)).catch(() => {});
            return response;
          })
          .catch(() => cached)
      );
    })
  );
});
