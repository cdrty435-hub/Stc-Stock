const CACHE_NAME = 'stc-stock-v1';
const urlsToCache = [
  '/Stc-Stock/stock_app.html',
  '/Stc-Stock/manifest.json'
];

self.addEventListener('install', e=>{
  e.waitUntil(
    caches.open(CACHE_NAME).then(cache=>cache.addAll(urlsToCache))
  );
});

self.addEventListener('fetch', e=>{
  e.respondWith(
    caches.match(e.request).then(response=>response||fetch(e.request))
  );
});
