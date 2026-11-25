const PREFIX='gigo-shell-v';
const CACHE=PREFIX+'2';
const SHELL=['./index.html','./app.js?v=20261125-1','./sync-config.js?v=20261026-2','./pwa.js?v=20261125-2','./work-hub-link.js?v=20261029','./manifest.webmanifest?v=20261026-1','./icon.svg'];
self.addEventListener('install',event=>{
  event.waitUntil(caches.open(CACHE).then(cache=>cache.addAll(SHELL)).then(()=>self.skipWaiting()));
});
self.addEventListener('activate',event=>{
  event.waitUntil(caches.keys().then(keys=>Promise.all(keys.filter(k=>k.startsWith(PREFIX)&&k!==CACHE).map(k=>caches.delete(k)))).then(()=>self.clients.claim()));
});
self.addEventListener('fetch',event=>{
  const req=event.request;
  if(req.method!=='GET'||!req.url.startsWith(self.registration.scope))return;
  if(req.mode==='navigate'){
    event.respondWith(fetch(req).then(async response=>{
      if(response.ok){
        try{
          const cache=await caches.open(CACHE);
          await cache.put('./index.html',response.clone());
        }catch{/* A full cache must still serve a successful network response. */}
      }
      return response;
    }).catch(()=>caches.open(CACHE).then(cache=>cache.match('./index.html'))));
    return;
  }
  event.respondWith((async()=>{
    let cache;
    try{
      cache=await caches.open(CACHE);
      const cached=await cache.match(req);
      if(cached)return cached;
    }catch{/* Shell files can also be loaded directly. */}
    const response=await fetch(req);
    if(cache&&response.ok)await cache.put(req,response.clone()).catch(()=>{});
    return response;
  })());
});
self.addEventListener('message',event=>{if(event.data==='SKIP_WAITING')self.skipWaiting();});
