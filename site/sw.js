// Минимальный service worker для PWA-установки
const CACHE='bntu-schedule-v1';
self.addEventListener('install',e=>{
  self.skipWaiting();
});
self.addEventListener('activate',e=>{
  e.waitUntil(self.clients.claim());
});
self.addEventListener('fetch',e=>{
  const url=new URL(e.request.url);
  // Не кэшируем data.json (он большой и должен быть свежим)
  if(url.pathname.endsWith('/data.json')) return;
  e.respondWith(
    caches.open(CACHE).then(cache=>
      cache.match(e.request).then(cached=>{
        const fetched=fetch(e.request).then(res=>{
          if(res.ok) cache.put(e.request,res.clone());
          return res;
        }).catch(()=>cached);
        return cached||fetched;
      })
    )
  );
});
