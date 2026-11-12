const V='sda-v15', TILES='tiles-v1';
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
    e.respondWith((async()=>{
      const c=await caches.open(V);
      try{
        const r=await Promise.race([fetch(e.request),new Promise((_,rej)=>setTimeout(()=>rej(new Error('t')),3500))]);
        if(r&&r.ok)c.put(e.request,r.clone());
        return r;
      }catch(err){
        const hit=await c.match(e.request,{ignoreSearch:true})||await c.match('./index.html');
        return hit||new Response('Offline',{status:504});
      }
    })());
  }
});
