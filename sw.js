// v4 — 2026-05-19 — wipe caches, never serve stale navigation
const CACHE = 'taskflow-v4';

self.addEventListener('install', e => {
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil((async () => {
    var keys = await caches.keys();
    await Promise.all(keys.map(k => caches.delete(k)));
    await self.clients.claim();
  })());
});

self.addEventListener('fetch', e => {
  if (e.request.mode !== 'navigate') return;
  e.respondWith((async () => {
    try {
      return await fetch(e.request, { cache: 'no-store' });
    } catch (_) {
      var c = await caches.open(CACHE);
      var hit = await c.match(e.request);
      if (hit) return hit;
      return new Response('Offline', { status: 503, statusText: 'Offline' });
    }
  })());
});
