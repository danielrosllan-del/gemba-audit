
const CACHE = "gemba-v1";
const ASSETS = [
  "./",
  "./index.html",
  "./app.js",
  "./manifest.webmanifest",
  "./assets/icon-192.png",
  "./assets/icon-256.png",
  "./assets/icon-384.png",
  "./assets/icon-512.png"
];

self.addEventListener("install", e=>{
  e.waitUntil(caches.open(CACHE).then(c=>c.addAll(ASSETS)));
});

self.addEventListener("activate", e=>{
  e.waitUntil(self.clients.claim());
});

self.addEventListener("fetch", e=>{
  e.respondWith(
    caches.match(e.request).then(r=> r || fetch(e.request).then(res=>{
      const copy = res.clone();
      if(e.request.url.startsWith(self.location.origin)){
        caches.open(CACHE).then(c=>c.put(e.request, copy));
      }
      return res;
    }).catch(()=> caches.match("./index.html")))
  );
});
