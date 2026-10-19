/* Hybrid · "Compartir con Hybrid" desde otras apps de Android (Web Share Target).
   Recibe archivos GPX/TCX, fotos y enlaces (p. ej. de Strava), los guarda un momento
   y abre la pantalla adecuada. Se carga en el service worker
   (vite.config.ts → workbox.importScripts). */

const SHARE_CACHE = 'hybrid-share'

const isRoute = (f) => /\.(gpx|tcx)$/i.test(f.name || '') || /gpx|tcx|xml/i.test(f.type || '')
const isImage = (f) => /^image\//i.test(f.type || '') || /\.(jpe?g|png|webp|heic|heif)$/i.test(f.name || '')

self.addEventListener('fetch', (event) => {
  const url = new URL(event.request.url)
  if (event.request.method !== 'POST' || !url.pathname.endsWith('/compartir')) return

  event.respondWith(
    (async () => {
      const scope = self.registration.scope
      try {
        const form = await event.request.formData()
        const files = form.getAll('file').filter((f) => f && typeof f !== 'string' && f.size > 0)
        const text = [form.get('title'), form.get('text'), form.get('url')].filter(Boolean).join(' ').slice(0, 1000)

        // Se guarda todo en la caché; la app lo recoge al abrirse (src/lib/shared.ts)
        const cache = await caches.open(SHARE_CACHE)
        for (const req of await cache.keys()) await cache.delete(req)
        const list = []
        for (let i = 0; i < files.length && i < 12; i++) {
          const f = files[i]
          const key = scope + 'shared/' + i
          await cache.put(key, new Response(f, { headers: { 'content-type': f.type || 'application/octet-stream' } }))
          list.push({ key, name: f.name || 'archivo', type: f.type || '' })
        }
        await cache.put(scope + 'shared/index', new Response(JSON.stringify({ files: list, text, at: Date.now() })))

        const route = files.some(isRoute)
        const images = files.some(isImage)
        const target = route || (!images && text) ? '#/actividad/nueva?compartido=1' : '#/compartir'
        return Response.redirect(scope + target, 303)
      } catch {
        return Response.redirect(scope + '#/actividad/nueva', 303)
      }
    })(),
  )
})
