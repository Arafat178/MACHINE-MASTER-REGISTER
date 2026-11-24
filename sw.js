const CACHE_NAME = "machine-master-v3";
const STATIC_ASSETS = [
  "./",
  "./index.html",
  "./style.css",
  "./script.js",
  "./manifest.json",
  "https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js"
];

const EXCEL_URLS = [
  encodeURI("./MACHINE MASTER REGISTER EXCEL.xlsx"),
  encodeURI("./MASTER INSTRUMENT LIST.xlsx")
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll([...STATIC_ASSETS, ...EXCEL_URLS]))
  );
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

// সব .xlsx ফাইলের জন্য Network-First: Excel আপডেট করলেই নতুন ডাটা অটোমেটিক লোড হবে
self.addEventListener("fetch", (event) => {
  if (event.request.url.includes(".xlsx")) {
    event.respondWith(
      fetch(event.request, { cache: "no-store" })
        .then((res) => {
          const clone = res.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(event.request, clone));
          return res;
        })
        .catch(() => caches.match(event.request))
    );
    return;
  }

  event.respondWith(
    caches.match(event.request).then((cached) => cached || fetch(event.request))
  );
});
