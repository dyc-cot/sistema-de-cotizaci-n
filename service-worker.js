// Service Worker — Alexander D&C
// Objetivo: (1) que el navegador pueda "instalar" el sistema como app, y
// (2) un respaldo mínimo por si se abre sin conexión (verás la última versión
// que se cargó con éxito, no la más reciente — para eso siempre necesitas internet).
//
// IMPORTANTE: usa "network-first" para el HTML principal. Esto es intencional:
// el sistema se actualiza seguido (nuevas funciones, correcciones), y no queremos
// que alguien quede atascado viendo una versión vieja en caché sin darse cuenta.
// Solo si NO hay conexión, se usa la copia guardada como respaldo.

const CACHE_NAME = 'alexander-dc-v2'; // súbele el número cuando quieras forzar limpieza de caché
const ARCHIVOS_APP_SHELL = [
  './index.html',
  './manifest.json',
  './icon-192.png',
  './icon-512.png',
  './icon-maskable-512.png'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(ARCHIVOS_APP_SHELL))
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((nombres) =>
      Promise.all(
        nombres.filter((n) => n !== CACHE_NAME).map((n) => caches.delete(n))
      )
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  const { request } = event;

  // Solo interceptamos GET del propio sitio (el HTML, el manifest, los íconos).
  // Peticiones a otros dominios (Google Apps Script, Gemini, Drive) NUNCA pasan por
  // este caché: no queremos servir una respuesta vieja de tus datos, y menos aún
  // una que dependa de la clave de acceso usada en ese momento.
  if (request.method !== 'GET') return;
  if (new URL(request.url).origin !== self.location.origin) return;

  event.respondWith(
    fetch(request)
      .then((respuesta) => {
        // Con internet: siempre la versión más reciente, y de paso actualizamos el respaldo.
        const copia = respuesta.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(request, copia));
        return respuesta;
      })
      .catch(() => {
        // Sin internet: usamos el respaldo guardado (si existe).
        return caches.match(request).then((coincidencia) => {
          return coincidencia || caches.match('./index.html');
        });
      })
  );
});
