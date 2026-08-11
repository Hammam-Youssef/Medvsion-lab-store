const CACHE = "labstore-v2";
const SHELL = ["./index.html", "./app.js", "./manifest.json", "./icon-192.png", "./icon-512.png"];

self.addEventListener("install", e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)));
  self.skipWaiting();
});

self.addEventListener("activate", e => {
  e.waitUntil(
    caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
  );
  self.clients.claim();
});

self.addEventListener("fetch", e => {
  // Only handle plain http(s) requests — skip chrome-extension:// and other schemes.
  if (!e.request.url.startsWith("http")) return;
  // Never cache API calls to Apps Script — always go to network for live data.
  if (e.request.url.includes("script.google.com")) return;
  // Network-first for our own files, so updates show up immediately.
  // Falls back to the cached copy only if the network request fails (offline).
  e.respondWith(
    fetch(e.request)
      .then(res => {
        const copy = res.clone();
        caches.open(CACHE).then(c => c.put(e.request, copy));
        return res;
      })
      .catch(() => caches.match(e.request))
  );
});
