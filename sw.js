// Service Worker for Credit Card Tracker PWA v4 (Network-First Strategy)
const CACHE_NAME = 'cc-tracker-v4';
const ASSETS = [
  './',
  './index.html',
  './manifest.json',
  './icon-v3.png'
];

self.addEventListener('install', (e) => {
  self.skipWaiting();
  e.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(ASSETS))
  );
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys().then((keys) => {
      return Promise.all(
        keys.map((k) => {
          if (k !== CACHE_NAME) return caches.delete(k);
        })
      );
    }).then(() => self.clients.claim())
  );
});

// Network-First: Always try to get the newest version from the server first!
self.addEventListener('fetch', (e) => {
  if (e.request.url.includes('script.google.com') || e.request.url.includes('googleusercontent.com')) {
    return;
  }

  e.respondWith(
    fetch(e.request)
      .then((networkRes) => {
        if (networkRes && networkRes.status === 200) {
          const resClone = networkRes.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(e.request, resClone));
        }
        return networkRes;
      })
      .catch(() => {
        return caches.match(e.request);
      })
  );
});
