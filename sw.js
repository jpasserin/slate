/* sw.js — cache-first, so the installed app opens without the network.
   CACHE is bumped by deploy.js on every ship; a stale worker would keep
   serving the old app otherwise. */
const CACHE = 'slate-v7';
const FILES = ['./', './index.html', './backend.js', './manifest.webmanifest',
  './icon.svg', './icon-192.png', './icon-512.png', './icon-maskable-512.png'];

/* Every file is fetched past the browser's own cache: GitHub Pages lets a
   file live there for ten minutes, and a new worker precaching a stale
   backend.js gave a new index.html an old backend - "BK.rates is not a
   function". */
self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE)
    .then((c) => Promise.all(FILES.map((f) => fetch(f, { cache: 'reload' }).then((r) => {
      if (!r.ok) throw new Error(f + ' ' + r.status);
      return c.put(f, r);
    }))))
    .then(() => self.skipWaiting()));
});
self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys()
    .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
    .then(() => self.clients.claim()));
});
self.addEventListener('fetch', (e) => {
  if (e.request.method !== 'GET') return;
  e.respondWith(caches.match(e.request, { ignoreSearch: true }).then((hit) => {
    if (hit) {
      e.waitUntil(fetch(e.request)
        .then((r) => (r && r.ok ? caches.open(CACHE).then((c) => c.put(e.request, r)) : null))
        .catch(() => {}));
      return hit;
    }
    return fetch(e.request).catch(() => caches.match('./index.html'));
  }));
});
