/* Aquatic Paradise — offline service worker
 * Caches the apps on-device so they open on any network (or none).
 * Bump CACHE when you change files to force an update.
 */
const CACHE = 'apr-v1';
const CORE = [
  './',
  'index.html',
  'field-app.html',
  'admin.html',
  'manifest.webmanifest',
  'assets/logo.jpeg',
  'assets/whatsapp-qr.jpeg',
  'assets/vendor/react.production.min.js',
  'assets/vendor/react-dom.production.min.js',
  'assets/gear/paddleboard-turtle.jpg',
  'assets/gear/paddleboard-tiki.jpg',
  'assets/gear/paddleboard-sage.jpg',
  'assets/gear/paddleboard-blue.jpg',
  'assets/gear/kayak.jpg',
  'assets/gear/snorkel.jpg',
  'assets/gear/float-lounge.jpg',
  'assets/gear/float-tube.jpg'
];

self.addEventListener('install', (e) => {
  self.skipWaiting();
  e.waitUntil(
    caches.open(CACHE).then((c) => Promise.allSettled(CORE.map((u) => c.add(u))))
  );
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET') return;  // never cache POST (status updates, form posts)

  e.respondWith(
    caches.match(req).then((hit) => {
      if (hit) return hit;   // cache-first: instant, offline-proof
      return fetch(req).then((res) => {
        // Cache good responses (same-origin + CDNs/fonts) for next time.
        try {
          if (res && (res.ok || res.type === 'opaque')) {
            const copy = res.clone();
            caches.open(CACHE).then((c) => c.put(req, copy));
          }
        } catch (err) {}
        return res;
      }).catch(() => {
        // Offline and not cached: fall back to the app shell for navigations.
        if (req.mode === 'navigate') return caches.match('index.html');
        return new Response('', { status: 503, statusText: 'offline' });
      });
    })
  );
});
