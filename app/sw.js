/* Team Biscuit · service worker: avisos al móvil y copia de la app para abrirla sin cobertura.
   Abre al instante con la última copia guardada y pide la nueva por detrás (la próxima vez ya sale la nueva);
   la app mira version.json (nunca guardado) y, si hay versión nueva, la trae fresca antes de recargar (real.js, verFresh).
   Sin copia guardada, espera a la red. Solo guarda la página de la app, la librería de Supabase y las imágenes propias; los datos van aparte (real.js). */
const CACHE = 'tb-app-v1';
const keep = (u) => (u.origin === self.location.origin && !/version\.json$/.test(u.pathname) && !/\/demo\.html$/.test(u.pathname))
  || (u.hostname === 'cdn.jsdelivr.net' && /supabase-js/.test(u.pathname));
self.addEventListener('fetch', (e) => {
  const r = e.request;
  if (r.method !== 'GET') return;
  const u = new URL(r.url);
  if (!keep(u)) return;
  if (navigator.onLine !== false) e.waitUntil(warm());
  e.respondWith((async () => {
    const c = await caches.open(CACHE);
    /* 9 nov (2): la página se guarda siempre con su dirección limpia (sin ?go=…). Antes cada apertura desde un aviso o tras
       actualizar dejaba otra copia con ?go=…; al buscar sin mirar el ?…, salía una copia vieja y la app se recargaba en bucle */
    const nav = r.mode === 'navigate';
    const key = nav ? u.origin + u.pathname : r;
    if (r.cache === 'no-store') return fetch(r).catch(async () => (await c.match(u.origin + u.pathname)) || Response.error()); // la app lo guarda ella misma (sin ?t=)
    /* 9 nov (3): la página se pide con ?t= para saltarse la copia del servidor de la web (GitHub Pages la guarda hasta 10 min):
       si no, justo después de publicar podía bajar la vieja y pisar la nueva que acababa de guardar la app */
    /* 13 nov: el código y los estilos van aparte (app-<huella>.js/.css). La página nueva solo se guarda cuando ya están
       guardados los suyos (si no, sin cobertura abriría en blanco); después se borran los de versiones anteriores */
    const net = nav ? fetch(u.origin + u.pathname + '?t=' + Date.now(), { cache: 'no-store', credentials: 'same-origin' }).then(async (res) => {
      if (res && res.ok) { try { const t = await res.clone().text(); if (await assets(c, t, u)) { await c.put(key, res.clone()); prune(c, t); } } catch (_) {} }
      return res; })
      : fetch(r).then((res) => { if (res && (res.ok || res.type === 'opaque')) c.put(key, res.clone()).catch(() => {}); return res; });
    /* 9 nov: la app pide la versión nueva con cache:'no-store' (verFresh): esa petición va siempre a la red.
       Antes se le daba la copia vieja guardada y el botón «Actualizar» recargaba una y otra vez sin cambiar nada */
    if (nav) for (const k of await c.keys()) { const ku = new URL(k.url); if (ku.search && ku.pathname === u.pathname) c.delete(k).catch(() => {}); }
    const old = await c.match(key);
    if (!old) return net;
    e.waitUntil(net.catch(() => {}));
    return old;
  })());
});
const APP = /app-[0-9a-f]{10}\.(?:js|css)/g;
async function assets(c, t, u) {
  for (const n of new Set(t.match(APP) || [])) {
    const k = new URL(n, u).href; if (await c.match(k)) continue;
    const r = await fetch(k); if (!r.ok) return false; await c.put(k, r);
  }
  return true;
}
async function prune(c, t) {
  const L = new Set(t.match(APP) || []); if (!L.size) return;
  for (const k of await c.keys()) { const m = new URL(k.url).pathname.match(/app-[0-9a-f]{10}\.(?:js|css)$/); if (m && !L.has(m[0])) c.delete(k).catch(() => {}); }
}
self.addEventListener('install', () => self.skipWaiting());
/* 13 nov: al pasar del service worker anterior (página todo en uno) a este, o si la app guardó ella sola la página nueva,
   se asegura de que el código y los estilos de la página guardada también lo estén (para abrir sin cobertura). Como mucho 1 vez/min */
let warmT = 0;
async function warm() {
  if (Date.now() - warmT < 60e3) return; warmT = Date.now();
  try { const c = await caches.open(CACHE); for (const k of await c.keys()) { const u = new URL(k.url); if (u.search || !/\/$|\/index\.html$/.test(u.pathname)) continue;
    const r = await c.match(k); if (r) await assets(c, await r.text(), u); } } catch (_) {}
}
self.addEventListener('activate', (e) => e.waitUntil(Promise.all([self.clients.claim(), warm()])));
self.addEventListener('push', (e) => {
  let d = {};
  try { d = e.data ? e.data.json() : {}; } catch (_) { d = { body: e.data ? e.data.text() : '' }; }
  e.waitUntil(self.registration.showNotification(d.title || 'Team Biscuit', {
    body: d.body || '', tag: d.tag || 'tb', renotify: true,
    icon: 'icons/icon-192.png', badge: 'icons/icon-192.png', lang: 'es', data: { go: d.go || null },
  }));
});
/* al tocar el aviso: abre la app en el chat o la ficha de ese cliente (go), y si ya estaba abierta, la lleva allí */
self.addEventListener('notificationclick', (e) => {
  e.notification.close();
  const go = e.notification.data && e.notification.data.go;
  e.waitUntil(self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then((ws) => {
    const w = ws.find((x) => x.url.startsWith(self.registration.scope));
    if (w) { if (go) w.postMessage({ tb: 'go', go }); return w.focus(); }
    return self.clients.openWindow(self.registration.scope + (go ? '?go=' + encodeURIComponent(JSON.stringify(go)) : ''));
  }));
});
