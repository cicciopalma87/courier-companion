const V='sda-v4', TILES='tiles-v1';
const SHELL=['./','./index.html','./app.js','./manifest.webmanifest','./icon-192.png','./icon-512.png','./lib/leaflet.js','./lib/leaflet.css','./lib/xlsx.full.min.js','./lib/zxing-wasm-reader.js','./lib/zxing_reader.wasm','./lib/images/marker-icon.png','./lib/images/layers.png'];
self.addEventListener('install',e=>{e.waitUntil(caches.open(V).then(c=>c.addAll(SHELL)).then(()=>self.skipWaiting()));});
self.addEventListener('activate',e=>{e.waitUntil(caches.keys().then(ks=>Promise.all(ks.filter(k=>k!==V&&k!==TILES).map(k=>caches.delete(k)))).then(()=>self.clients.claim()));});
self.addEventListener('fetch',e=>{
  const u=new URL(e.request.url);
  if(e.request.method!=='GET')return;
  if(u.hostname.endsWith('tile.openstreetmap.org')){
    e.respondWith(caches.open(TILES).then(async c=>{
      const hit=await c.match(e.request.url); if(hit)return hit;
      try{const r=await fetch(e.request.url,{mode:'cors'}); if(r&&r.ok)c.put(e.request.url,r.clone()); return r;}catch(err){return new Response('',{status:504});}
    }));
    return;
  }
  if(u.origin===location.origin){
    e.respondWith(caches.match(e.request,{ignoreSearch:true}).then(hit=>{
      const net=fetch(e.request).then(r=>{if(r&&r.ok){const cp=r.clone();caches.open(V).then(c=>c.put(e.request,cp));}return r;}).catch(()=>hit);
      return hit||net;
    }));
  }
});
