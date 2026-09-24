/* ALEL OPEX Hub — service worker
   Strategy:
     • HTML documents (navigate)  → STALE-WHILE-REVALIDATE: serve instantly
       from cache, then refresh in the background. Repeat visits are instant.
     • scripts / styles / json     → NETWORK FIRST (small, must stay fresh)
     • icons / images              → CACHE FIRST (fast, rarely change)     */
const CACHE = "alel-opex-v3";
const SHELL = [
  "./",
  "./index.html",
  "./manifest.webmanifest",
  "./icons/favicon-48.png",
  "./icons/icon-192.png",
  "./icons/icon-512.png",
  "./icons/icon-maskable-512.png",
  "./icons/apple-touch-icon.png",
];

self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(SHELL)).catch(() => {}).then(() => self.skipWaiting()));
});

self.addEventListener("activate", (e) => {
  e.waitUntil(caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k)))).then(() => self.clients.claim()));
});

self.addEventListener("fetch", (e) => {
  const req = e.request;
  if (req.method !== "GET") return;
  const url = new URL(req.url);
  if (url.origin !== self.location.origin) return;            // cross-origin: Apps Script, Google Sheets, fonts
  if (url.pathname.indexOf("/dashboard/") !== -1) return;     // dashboard has its own SW

  // 1) HTML — stale-while-revalidate (instant repeat loads)
  if (req.mode === "navigate") {
    e.respondWith(
      caches.match(req).then((cached) => {
        const fetched = fetch(req).then((res) => {
          if (res && res.status === 200 && res.type === "basic") {
            const copy = res.clone();
            caches.open(CACHE).then((c) => c.put(req, copy));
          }
          return res;
        }).catch(() => cached || caches.match("./index.html"));
        return cached || fetched;
      })
    );
    return;
  }

  // 2) scripts / styles / json — network-first (small, always fresh)
  if (/\.(js|css|json|webmanifest)$/i.test(url.pathname)) {
    e.respondWith(
      fetch(req).then((res) => {
        if (res && res.status === 200 && res.type === "basic") {
          const copy = res.clone();
          caches.open(CACHE).then((c) => c.put(req, copy));
        }
        return res;
      }).catch(() => caches.match(req))
    );
    return;
  }

  // 3) icons / images — cache-first
  e.respondWith(
    caches.match(req).then((cached) => cached || fetch(req).then((res) => {
      if (res && res.status === 200 && res.type === "basic") {
        const copy = res.clone();
        caches.open(CACHE).then((c) => c.put(req, copy));
      }
      return res;
    }).catch(() => cached))
  );
});
