const CACHE_NAME = 'excel-tutor-shell-v3';
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
  const url = new URL(event.request.url);

  if (url.origin.includes('api.groq.com')) {
    // Never cache live tutoring calls — always hit the network.
    event.respondWith(fetch(event.request));
    return;
  }

  event.respondWith(
    caches.match(event.request).then((cached) => {
      return (
        cached ||
        fetch(event.request)
          .then((response) => {
            const clone = response.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(event.request, clone));
            return response;
          })
          .catch(() => cached)
      );
    })
  );
});
