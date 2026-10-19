/* Hybrid · botones de la notificación del entreno en curso.
   Se carga dentro del service worker (vite.config.ts → workbox.importScripts). */

self.addEventListener('notificationclick', (event) => {
  // Solo la notificación del entreno en curso (las demás las gestiona push-sw.js)
  if (event.notification.tag !== 'hybrid-workout') return
  const action = event.action // '', 'rest-minus', 'rest-plus', 'rest-skip', 'complete'
  // Los botones no cierran la notificación: la app la actualiza en el momento
  // (así se pueden pulsar varias veces seguidas). Tocarla en sí sí la cierra.
  if (!action) event.notification.close()
  event.waitUntil(
    (async () => {
      const windows = await self.clients.matchAll({ type: 'window', includeUncontrolled: true })
      if (action && windows.length) {
        // La app sigue abierta en segundo plano: aplica la acción sin abrirla
        windows.forEach((w) => w.postMessage({ type: 'hybrid-action', action }))
        return
      }
      const scope = self.registration.scope
      const target = windows.find((w) => w.url.startsWith(scope))
      if (target) {
        await target.focus()
        target.postMessage({ type: 'hybrid-action', action: action || 'open' })
      } else {
        // La app estaba cerrada: se abre en el entreno (que está guardado en el móvil)
        await self.clients.openWindow(scope + '#/entreno')
      }
    })(),
  )
})
