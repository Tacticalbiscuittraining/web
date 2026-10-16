/* Team Biscuit · service worker: SOLO muestra los avisos al móvil (no guarda nada en caché). */
self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', (e) => e.waitUntil(self.clients.claim()));
self.addEventListener('push', (e) => {
  let d = {};
  try { d = e.data ? e.data.json() : {}; } catch (_) { d = { body: e.data ? e.data.text() : '' }; }
  e.waitUntil(self.registration.showNotification(d.title || 'Team Biscuit', {
    body: d.body || '', tag: d.tag || 'tb', renotify: true,
    icon: 'icons/icon-192.png', badge: 'icons/icon-192.png', lang: 'es',
  }));
});
self.addEventListener('notificationclick', (e) => {
  e.notification.close();
  e.waitUntil(self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then((ws) => {
    const w = ws.find((x) => x.url.startsWith(self.registration.scope));
    return w ? w.focus() : self.clients.openWindow(self.registration.scope);
  }));
});
