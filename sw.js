// Guarda la app en el teléfono para que abra aunque no haya señal.
const CACHE = 'rinti-v11';
const SHELL = ['./', 'index.html', 'app.html', 'config.js', 'manifest.webmanifest', 'icon.svg', 'img/free-blanco.png', 'img/rinti-r.png', 'img/ricocan.png', 'img/ricocat.png'];
self.addEventListener('install', e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL))); self.skipWaiting(); });
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))));
  self.clients.claim();
});
self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  if (e.request.method !== 'GET' || url.origin !== location.origin) return; // Apps Script y CDN van directo
  // Red primero (para recibir actualizaciones), caché si no hay señal.
  e.respondWith(fetch(e.request).then(r => {
    const copy = r.clone(); caches.open(CACHE).then(c => c.put(e.request, copy)); return r;
  }).catch(() => caches.match(e.request).then(r => r || caches.match('app.html'))));
});
