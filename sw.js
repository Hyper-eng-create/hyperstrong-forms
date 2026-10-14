// HyperStrong Field Forms – offline support
const CACHE = 'hs-forms-1.0.0-202610141152';
const SHELL = ["./", "index.html", "lift.html", "pre.html", "cold.html", "hot.html", "job.html", "manifest.webmanifest", "lib/html2canvas.min.js", "lib/jspdf.umd.min.js", "icons/icon-192.png", "icons/icon-512.png", "icons/maskable-512.png", "icons/apple-touch-icon.png", "icons/favicon-32.png"];
self.addEventListener('install', e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting())); });
self.addEventListener('activate', e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k.startsWith('hs-forms-') && k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener('fetch', e => {
  const req = e.request; if (req.method !== 'GET') return;
  const url = new URL(req.url);
  // Fonts: use the cached copy, refresh in the background
  if (/fonts\.(googleapis|gstatic)\.com$/.test(url.hostname)) {
    e.respondWith(caches.open(CACHE).then(async c => { const hit = await c.match(req); const net = fetch(req).then(r => { c.put(req, r.clone()); return r; }).catch(() => hit); return hit || net; }));
    return;
  }
  if (url.origin !== location.origin) return;
  // Pages: network first so updates arrive; cached copy when offline
  if (req.mode === 'navigate' || req.destination === 'document') {
    e.respondWith(fetch(req).then(r => { const copy = r.clone(); caches.open(CACHE).then(c => c.put(req, copy)); return r; })
      .catch(() => caches.match(req, {ignoreSearch:true}).then(r => r || caches.match('index.html'))));
    return;
  }
  e.respondWith(caches.match(req, {ignoreSearch:true}).then(r => r || fetch(req)));
});
