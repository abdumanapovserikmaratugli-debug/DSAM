// DSAM — минимальный service worker.
// Задача: сделать приложение устанавливаемым (PWA) и отдавать
// оболочку из кэша, если сеть недоступна. Данные (клубы, брони)
// всегда идут напрямую в сеть — их не кэшируем, чтобы не показывать
// устаревшие цены/свободные места.
const CACHE_NAME = 'dsam-shell-v1';
const SHELL_FILES = ['./index.html', './manifest.json', './icon-192.png', './icon-512.png'];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(SHELL_FILES))
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  const url = new URL(event.request.url);

  // Запросы к Apps Script API — никогда не кэшируем, всегда свежие данные
  if (url.hostname.includes('script.google.com')) return;

  event.respondWith(
    caches.match(event.request).then((cached) => {
      return cached || fetch(event.request).catch(() => caches.match('./index.html'));
    })
  );
});
