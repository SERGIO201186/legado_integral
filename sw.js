const CACHE_NAME = 'legado-integral-v2';
const PRECACHE_URLS = [
  './',
  './index.html',
  './manifest.json',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './icons/icon-512-maskable.png',
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then((cache) => cache.addAll(PRECACHE_URLS))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

// Red primero, caché solo como respaldo sin conexión — antes era al revés
// (caché primero) y como el nombre de la caché no cambiaba en cada
// despliegue, el teléfono se quedaba sirviendo el index.html viejo para
// siempre, sin importar cuántas actualizaciones se publicaran en el
// repositorio: nunca volvía a consultar la red mientras existiera algo
// guardado con esa misma URL. Con la app en línea (que además necesita
// internet para hablar con el backend de Apps Script) siempre se busca la
// versión más reciente primero.
self.addEventListener('fetch', (event) => {
  const req = event.request;
  if (req.method !== 'GET' || !req.url.startsWith(self.location.origin)) return;

  event.respondWith(
    fetch(req).then((res) => {
      if (res.ok) {
        const copy = res.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(req, copy));
      }
      return res;
    }).catch(() => caches.match(req))
  );
});
