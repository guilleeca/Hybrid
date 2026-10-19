/* Hybrid · notificaciones push (cuando alguien entrena, da me gusta o comenta).
   Se carga en el service worker (vite.config.ts → workbox.importScripts). */

self.addEventListener('push', (event) => {
  let data = {}
  try {
    data = event.data ? event.data.json() : {}
  } catch {
    data = { title: 'Hybrid', body: event.data ? event.data.text() : '' }
  }
  const scope = self.registration.scope
  if (data.type === 'rest') {
    event.waitUntil(restDone(data, scope))
    return
  }
  event.waitUntil(
    self.registration.showNotification(data.title || 'Hybrid', {
      body: data.body || '',
      icon: scope + 'pwa-192.png',
      badge: scope + 'pwa-192.png',
      tag: data.tag,
      renotify: true,
      data: { url: data.url || '' },
    }),
  )
})

// Al tocarla se abre Hybrid en esa pantalla (las del entreno en curso las gestiona notifications-sw.js)
self.addEventListener('notificationclick', (event) => {
  const url = event.notification.data && event.notification.data.url
  if (typeof url !== 'string') return
  event.notification.close()
  event.waitUntil(
    (async () => {
      const scope = self.registration.scope
      const target = scope + url
      const windows = await self.clients.matchAll({ type: 'window', includeUncontrolled: true })
      const win = windows.find((w) => w.url.startsWith(scope))
      if (win) {
        await win.focus()
        win.postMessage({ type: 'hybrid-open', url })
      } else {
        await self.clients.openWindow(target)
      }
    })(),
  )
})

// Fin del descanso avisado por el servidor (rest-alarm). Sustituye a la notificación
// del entreno en curso, conservando su texto ("Siguiente: …") y el botón de completar serie.
async function restDone(data, scope) {
  const windows = await self.clients.matchAll({ type: 'window', includeUncontrolled: true })
  // Con la app en pantalla ya avisa ella misma
  if (windows.some((w) => w.visibilityState === 'visible')) return
  const current = (await self.registration.getNotifications({ tag: 'hybrid-workout' }))[0]
  // La app ya avisó (estaba despierta): no se repite el sonido
  if (current && current.title === data.title) return
  await self.registration.showNotification(data.title, {
    tag: 'hybrid-workout',
    body: current && current.body ? current.body : data.body,
    icon: scope + 'pwa-192.png',
    badge: scope + 'pwa-192.png',
    renotify: true,
    silent: false,
    requireInteraction: true,
    vibrate: [250, 120, 250],
    actions: [{ action: 'complete', title: '✓ Completar serie' }],
    timestamp: Date.now(),
  })
}
