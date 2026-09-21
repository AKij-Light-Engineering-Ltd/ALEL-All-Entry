/* ALEL OPEX Hub — service worker
   Strategy:
     • documents + scripts/styles/json  → NETWORK FIRST (updates land immediately, cache is the offline fallback)
     • icons / images                    → CACHE FIRST (fast, rarely change)
   This prevents the hub from serving stale copies of the apps. */
const CACHE = "alel-opex-v2";
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
  e.waitUntil(
    caches.open(CACHE).then((c) => c.addAll(SHELL)).catch(() => {}).then(() => self.skipWaiting())
  );
});

self.addEventListener("activate", (e) => {
  e.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", (e) => {
  const req = e.request;
  if (req.method !== "GET") return;
  const url = new URL(req.url);
  // Never intercept cross-origin (Apps Script backend, Google Sheets, fonts)
  if (url.origin !== self.location.origin) return;
  // The production dashboard ships its own service worker
  if (url.pathname.indexOf("/dashboard/") !== -1) return;

  const fresh = req.mode === "navigate" || /\.(html|js|css|json|webmanifest)$/i.test(url.pathname);

  if (fresh) {
    // network-first: always try the live file, fall back to cache when offline
    e.respondWith(
      fetch(req)
        .then((res) => {
          if (res && res.status === 200 && res.type === "basic") {
            const copy = res.clone();
            caches.open(CACHE).then((c) => c.put(req, copy));
          }
          return res;
        })
        .catch(() => caches.match(req).then((r) => r || (req.mode === "navigate" ? caches.match("./index.html") : undefined)))
    );
    return;
  }

  // cache-first for static assets
  e.respondWith(
    caches.match(req).then(
      (cached) =>
        cached ||
        fetch(req)
          .then((res) => {
            if (res && res.status === 200 && res.type === "basic") {
              const copy = res.clone();
              caches.open(CACHE).then((c) => c.put(req, copy));
            }
            return res;
          })
          .catch(() => cached)
    )
  );
});
