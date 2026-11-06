// ใช้งานแบบไม่มีอินเทอร์เน็ต: เก็บหน้าแอปและไลบรารีที่ใช้ไว้ในเครื่อง
const C='dental-v3';
const CORE=['./','./index.html','./entry.html','./manifest.webmanifest','./entry.webmanifest','./icon-192.png','./icon-512.png','./entry-icon-192.png','./entry-icon-512.png'];
self.addEventListener('install',e=>{
  e.waitUntil(caches.open(C).then(c=>Promise.all(CORE.map(u=>c.add(u).catch(()=>{})))).then(()=>self.skipWaiting()));
});
self.addEventListener('activate',e=>{
  e.waitUntil(caches.keys().then(ks=>Promise.all(ks.filter(k=>k!==C).map(k=>caches.delete(k)))).then(()=>self.clients.claim()));
});
self.addEventListener('fetch',e=>{
  const r=e.request;if(r.method!=='GET')return;
  const u=new URL(r.url);
  if(u.origin===location.origin){
    // ไฟล์ของแอป: ใช้เน็ตก่อนเพื่อได้รุ่นล่าสุด ถ้าไม่มีเน็ตใช้ที่เก็บไว้
    e.respondWith(fetch(r).then(res=>{if(res.ok){const cp=res.clone();caches.open(C).then(c=>c.put(r,cp))}return res})
      .catch(()=>caches.match(r,{ignoreSearch:true}).then(m=>m||caches.match('./index.html'))));
  }else{
    // ไลบรารีจาก CDN (ล็อกรุ่นไว้แล้ว): ใช้ที่เก็บไว้ก่อน
    e.respondWith(caches.match(r).then(m=>m||fetch(r).then(res=>{if(res.ok||res.type==='opaque'){const cp=res.clone();caches.open(C).then(c=>c.put(r,cp))}return res})));
  }
});
