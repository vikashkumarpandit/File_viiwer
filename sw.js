/* Service worker: offline support + receiving files shared from other apps */
const VERSION = 'doc-viewer-v2';
const SHARED = 'ppt-shared-v1';
const CORE = [
  './', './index.html', './style.css', './app.js', './pptx.js', './ppt-legacy.js', './readers.js', './jszip.min.js',
  './manifest.webmanifest', './icon-192.png', './icon-512.png', './icon-maskable-512.png'
];

self.addEventListener('install', (e) => {
  e.waitUntil(
    caches.open(VERSION)
      .then((c) => Promise.allSettled(CORE.map((u) => c.add(u))))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys()
      .then((ks) => Promise.all(ks.filter((k) => k !== VERSION && k !== SHARED).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

async function handleShare(request) {
  try {
    const fd = await request.formData();
    const files = fd.getAll('files').filter((f) => f && f.name);
    const cache = await caches.open(SHARED);
    for (const k of await cache.keys()) await cache.delete(k);
    let i = 0;
    for (const f of files) {
      const url = new URL('shared/' + (i++) + '/' + encodeURIComponent(f.name), self.registration.scope).href;
      await cache.put(new Request(url), new Response(f, {
        headers: {
          'Content-Type': f.type || 'application/octet-stream',
          'X-Name': encodeURIComponent(f.name),
          'X-Mod': String(f.lastModified || 0)
        }
      }));
    }
  } catch (err) { /* ignore */ }
  return Response.redirect(new URL('./?shared=1', self.registration.scope).href, 303);
}

self.addEventListener('fetch', (e) => {
  const req = e.request;
  const url = new URL(req.url);
  if (req.method === 'POST' && url.pathname.endsWith('/share-target')) {
    e.respondWith(handleShare(req));
    return;
  }
  if (req.method !== 'GET' || url.origin !== location.origin) return;
  // cached copy first, refresh in the background (also caches pdf.js files on first use)
  e.respondWith(
    caches.open(VERSION).then(async (c) => {
      const hit = await c.match(req, { ignoreSearch: true });
      const net = fetch(req).then((r) => { if (r && r.ok && r.status === 200) c.put(req, r.clone()); return r; }).catch(() => null);
      if (hit) { e.waitUntil(net); return hit; }
      const r = await net;
      return r || (req.mode === 'navigate' ? c.match('./index.html') : Response.error());
    })
  );
});
