/* Hybrid · fotos de entrenos guardadas en el móvil.
   Los enlaces de las fotos llevan un "token" que cambia cada pocos días; aquí se
   guardan por su ruta (sin el token), así una foto ya vista no se vuelve a descargar.
   Se carga en el service worker (vite.config.ts → workbox.importScripts). */

const PHOTO_CACHE = 'workout-photos-v1'
const PHOTO_MAX = 400

self.addEventListener('fetch', (event) => {
  const req = event.request
  if (req.method !== 'GET') return
  const url = new URL(req.url)
  if (!url.pathname.includes('/storage/v1/object/sign/workout-photos/')) return
  const key = url.origin + url.pathname
  event.respondWith(
    (async () => {
      const cache = await caches.open(PHOTO_CACHE)
      const hit = await cache.match(key)
      if (hit) return hit
      const res = await fetch(req)
      if (res.ok) {
        await cache.put(key, res.clone())
        trim(cache)
      }
      return res
    })(),
  )
})

// Como mucho PHOTO_MAX fotos: se borran las más antiguas
async function trim(cache) {
  const keys = await cache.keys()
  for (const k of keys.slice(0, Math.max(0, keys.length - PHOTO_MAX))) await cache.delete(k)
}
