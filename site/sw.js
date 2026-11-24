// CarWatch service worker. Network first, so the list is always fresh when online; the last copy of each page and
// data file answers when offline. It also lets Chrome offer to install the app, which puts CarWatch in Android's
// Share menu (see share_target in manifest.webmanifest).
const CACHE = "carwatch-v1";
self.addEventListener("install", () => self.skipWaiting());
self.addEventListener("activate", e => e.waitUntil(self.clients.claim()));
const keyOf = req => { const u = new URL(req.url); u.search = ""; return u.href; };     // data?123 and data are one file
self.addEventListener("fetch", e => {
  const req = e.request;
  if (req.method !== "GET" || new URL(req.url).origin !== self.location.origin) return;   // photos, Google, the group: untouched
  e.respondWith(fetch(req).then(res => {
    if (res.ok) { const copy = res.clone(); caches.open(CACHE).then(c => c.put(keyOf(req), copy)).catch(() => {}); }
    return res;
  }).catch(() => caches.match(keyOf(req)).then(r => r || (req.mode === "navigate" ? caches.match(new URL("./", self.location).href) : undefined))
    .then(r => r || Response.error())));
});
