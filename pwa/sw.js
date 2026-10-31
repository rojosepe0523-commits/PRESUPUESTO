// Service worker: guarda la app en el teléfono para que abra sin conexión.
// Al publicar cambios en index.html, sube el número de VERSION.
const VERSION = 'presupuesto-v4';
const ARCHIVOS = [
  './',
  './index.html',
  './manifest.webmanifest',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './icons/apple-touch-icon.png'
];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(VERSION).then(c => c.addAll(ARCHIVOS)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys()
    .then(ks => Promise.all(ks.filter(k => k !== VERSION).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});

function conTiempo(promesa, ms) {
  return new Promise((res, rej) => {
    const t = setTimeout(() => rej(new Error('timeout')), ms);
    promesa.then(v => { clearTimeout(t); res(v); }, err => { clearTimeout(t); rej(err); });
  });
}

self.addEventListener('fetch', e => {
  const req = e.request;
  // Solo archivos propios de la app; las llamadas a Apps Script (POST) pasan directo.
  if (req.method !== 'GET' || new URL(req.url).origin !== self.location.origin) return;

  if (req.mode === 'navigate') {
    // Red primero (para recibir actualizaciones), con 3 s de espera máximo; si no, la copia guardada.
    const red = fetch(req).then(r => {
      if (r.ok) { const copia = r.clone(); caches.open(VERSION).then(c => c.put('./index.html', copia)); }
      return r;
    });
    e.respondWith(conTiempo(red, 3000).catch(() => caches.match('./index.html').then(r => r || red)));
    return;
  }

  e.respondWith(caches.match(req).then(r => r || fetch(req)));
});
