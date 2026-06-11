const CACHE_NAME = 'botvip-cache-v1';
const ASSETS_TO_CACHE = [
  '/',
  '/static/icon.png',
  '/static/manifest.json'
];

// Instalar Service Worker
self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return cache.addAll(ASSETS_TO_CACHE);
    }).then(() => self.skipWaiting())
  );
});

// Activar y limpiar cachés antiguas
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((cacheNames) => {
      return Promise.all(
        cacheNames.map((cache) => {
          if (cache !== CACHE_NAME) {
            return caches.delete(cache);
          }
        })
      );
    }).then(() => self.clients.claim())
  );
});

// Interceptar solicitudes - Estrategia Network First
self.addEventListener('fetch', (event) => {
  // Solo interceptar peticiones GET normales
  if (event.request.method !== 'GET') return;
  
  // No interceptar llamadas a API, webhook ni sincronización
  if (
    event.request.url.includes('/api/') || 
    event.request.url.includes('/webhook') || 
    event.request.url.includes('/sync')
  ) {
    return;
  }

  event.respondWith(
    fetch(event.request)
      .then((response) => {
        // Clonar y guardar en caché si es una respuesta exitosa
        if (response && response.status === 200 && response.type === 'basic') {
          const responseToCache = response.clone();
          caches.open(CACHE_NAME).then((cache) => {
            cache.put(event.request, responseToCache);
          });
        }
        return response;
      })
      .catch(() => {
        // Retornar del caché si falla la red
        return caches.match(event.request);
      })
  );
});
