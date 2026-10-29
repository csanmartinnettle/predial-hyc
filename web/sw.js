/* Service worker: permite abrir la app sin señal y la mantiene actualizada.
   - Archivos propios (HTML, JS, CSS, íconos): primero internet, si no hay señal usa la copia guardada.
     Así, cada vez que se publica una versión nueva, los teléfonos la toman solos al abrir la app.
   - Librerías externas (fuentes, Firebase, PDF): copia guardada primero. */
const CACHE = "hyc-predial-v1";
const SHELL = [
  "./", "index.html", "styles.css", "app.css", "app.js", "manifest.webmanifest",
  "icons/logo.png", "icons/icon-192.png", "icons/icon-512.png", "icons/apple-touch-icon.png", "icons/favicon-48.png"
];
const CDN = [];

self.addEventListener("install", e => {
  e.waitUntil(caches.open(CACHE).then(async c => {
    await c.addAll(SHELL);
    await Promise.all(CDN.map(u => c.add(new Request(u, { mode: "cors" })).catch(() => {})));
  }).then(() => self.skipWaiting()));
});

self.addEventListener("activate", e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});

self.addEventListener("fetch", e => {
  const req = e.request;
  if (req.method !== "GET") return;
  const url = new URL(req.url);
  // Firestore y la autenticación manejan su propia conexión: no se interceptan.
  if (url.hostname.endsWith("googleapis.com") && !url.hostname.startsWith("fonts.")) return;

  if (url.origin === location.origin) {
    e.respondWith(fetch(req).then(res => {
      if (res.ok) { const copy = res.clone(); caches.open(CACHE).then(c => c.put(req, copy)); }
      return res;
    }).catch(() => caches.match(req, { ignoreSearch: true }).then(r => r || caches.match("index.html"))));
    return;
  }

  const cacheable = url.hostname === "cdnjs.cloudflare.com" || url.hostname === "www.gstatic.com" ||
    url.hostname === "fonts.googleapis.com" || url.hostname === "fonts.gstatic.com";
  if (cacheable) {
    e.respondWith(caches.match(req).then(hit => hit || fetch(req).then(res => {
      if (res.ok || res.type === "opaque") { const copy = res.clone(); caches.open(CACHE).then(c => c.put(req, copy)); }
      return res;
    })));
  }
});
