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
  e.respondWith((async () => {
    const c = await caches.open(CACHE);
    /* 9 nov (2): la página se guarda siempre con su dirección limpia (sin ?go=…). Antes cada apertura desde un aviso o tras
       actualizar dejaba otra copia con ?go=…; al buscar sin mirar el ?…, salía una copia vieja y la app se recargaba en bucle */
    const nav = r.mode === 'navigate';
    const key = nav ? u.origin + u.pathname : r;
    const net = fetch(r).then((res) => { if (res && (res.ok || res.type === 'opaque')) c.put(key, res.clone()).catch(() => {}); return res; });
    /* 9 nov: la app pide la versión nueva con cache:'no-store' (verFresh): esa petición va siempre a la red.
       Antes se le daba la copia vieja guardada y el botón «Actualizar» recargaba una y otra vez sin cambiar nada */
    if (r.cache === 'no-store') return net.catch(async () => (await c.match(r)) || Response.error());
    if (nav) for (const k of await c.keys()) { const ku = new URL(k.url); if (ku.search && ku.pathname === u.pathname) c.delete(k).catch(() => {}); }
    const old = await c.match(key);
    if (!old) return net;
    e.waitUntil(net.catch(() => {}));
    return old;
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
