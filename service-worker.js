const APP_CACHE = 'storecheck-app-v2';
const APP_SHELL = [
  './',
  './index.html',
  './manifest.webmanifest',
  './assets/favicon.svg',
  './assets/icon-192.png',
  './assets/icon-512.png'
];

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(APP_CACHE)
      .then(cache => cache.addAll(APP_SHELL))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(
        keys.filter(key => key !== APP_CACHE).map(key => caches.delete(key))
      ))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', event => {
  const request = event.request;
  if(request.method !== 'GET') return;

  const url = new URL(request.url);
  if(url.hostname === 'script.google.com' || url.hostname.endsWith('.googleusercontent.com')) return;

  if(request.mode === 'navigate'){
    event.respondWith(
      fetch(request).catch(async () =>
        (await caches.match('./index.html')) || caches.match('./')
      )
    );
    return;
  }

  event.respondWith(
    caches.match(request).then(cached => {
      if(cached) return cached;
      return fetch(request).then(response => {
        if(!response || (!response.ok && response.type !== 'opaque')) return response;
        const copy = response.clone();
        return caches.open(APP_CACHE)
          .then(cache => cache.put(request, copy))
          .then(() => response);
      });
    })
  );
});
