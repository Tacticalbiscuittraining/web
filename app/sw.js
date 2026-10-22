/* Team Biscuit · service worker: avisos al móvil y copia de la app para abrirla sin cobertura.
   Siempre pide primero la versión nueva; si no hay red (o tarda más de 4 s), usa la última guardada.
   Solo guarda la página de la app, la librería de Supabase y las imágenes propias; los datos van aparte (real.js). */
const CACHE = 'tb-app-v1';
const keep = (u) => (u.origin === self.location.origin && !/version\.json$/.test(u.pathname) && !/\/demo\.html$/.test(u.pathname))
  || (u.hostname === 'cdn.jsdelivr.net' && /supabase-js/.test(u.pathname));
self.addEventListener('fetch', (e) => {
  const r = e.request;
  if (r.method !== 'GET') return;
  const u = new URL(r.url);
  if (!keep(u)) return;
  e.respondWith((async () => {
    const c = await caches.open(CACHE);
    const net = fetch(r).then((res) => { if (res && (res.ok || res.type === 'opaque')) c.put(r, res.clone()).catch(() => {}); return res; });
    const old = await c.match(r, { ignoreSearch: r.mode === 'navigate' });
    if (!old) return net;
    return Promise.race([net.catch(() => old), new Promise((ok) => setTimeout(() => ok(old), 4000))]);
  })());
});
self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', (e) => e.waitUntil(self.clients.claim()));
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
