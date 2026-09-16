/* Service worker mínimo: hace la página instalable en Android (Chrome) sin guardar nada en caché.
   Todo se pide siempre a la red, así los datos nuevos aparecen apenas se publican. */
self.addEventListener('install', function () { self.skipWaiting(); });
self.addEventListener('activate', function (e) { e.waitUntil(self.clients.claim()); });
self.addEventListener('fetch', function (e) { e.respondWith(fetch(e.request)); });
