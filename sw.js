/* Service worker: guarda la app, los datos y los retratos para usarla sin conexión.
   Sube VERSION cada vez que cambies datos o código para que los móviles se actualicen. */
const VERSION = 'mcoc-v27';
const CORE = [
  './', 'index.html', 'css/style.css', 'js/app.js', 'data/data.js', 'data/images.js', 'data/kit-effects.js',
  'manifest.webmanifest', 'icons/icon-192.png', 'icons/icon-512.png', 'icons/icon-maskable-512.png', 'icons/apple-touch-icon.png'
];

self.addEventListener('install', event => {
  event.waitUntil((async () => {
    const cache = await caches.open(VERSION);
    await cache.addAll(CORE);
    // Retratos: se descargan todos en segundo plano (si alguno falla, no pasa nada)
    let imgs = [];
    try { importScripts('data/images.js'); imgs = Object.values(self.MCOC_IMG || {}).map(f => 'img/' + f); } catch (e) { /* sin imágenes */ }
    await Promise.all(imgs.map(u => cache.add(u).catch(() => null)));
    await self.skipWaiting();
  })());
});

self.addEventListener('activate', event => {
  event.waitUntil((async () => {
    // solo las cachés propias (la v2, en /v2/, usa «mcoc2-…»)
    for (const k of await caches.keys()) if (k.startsWith('mcoc-v') && k !== VERSION) await caches.delete(k);
    await self.clients.claim();
  })());
});

self.addEventListener('fetch', event => {
  const req = event.request;
  if (req.method !== 'GET' || new URL(req.url).origin !== location.origin) return;
  const isImg = /\/img\//.test(req.url);
  event.respondWith((async () => {
    const cache = await caches.open(VERSION);
    if (isImg) {
      // Retratos: primero la caché (no cambian)
      const hit = await cache.match(req, { ignoreSearch: true });
      if (hit) return hit;
    }
    // Código y datos: primero la red (así llegan las actualizaciones); sin conexión, la caché
    try {
      // cache: 'no-cache' evita que la caché HTTP del navegador (10 min en GitHub Pages) sirva una versión vieja
      const res = await fetch(req.url, { cache: 'no-cache', credentials: 'same-origin' });
      if (res.ok) cache.put(req, res.clone());
      return res;
    } catch (e) {
      const hit = await cache.match(req, { ignoreSearch: true });
      if (hit) return hit;
      if (req.mode === 'navigate') return cache.match('index.html');
      throw e;
    }
  })());
});
