/* Guarda o app no celular para abrir sem sinal */
const VERSAO = 'osc-v1';
const ARQUIVOS = ['./', 'index.html', 'comum.js', 'comum.css', 'config.js', 'manifest.webmanifest', 'icone-180.png', 'icone-192.png', 'icone-512.png'];
self.addEventListener('install', e => { e.waitUntil(caches.open(VERSAO).then(c => c.addAll(ARQUIVOS)).then(() => self.skipWaiting())); });
self.addEventListener('activate', e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== VERSAO).map(k => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET' || new URL(req.url).origin !== location.origin) return;
  e.respondWith(caches.open(VERSAO).then(async c => {
    const salvo = await c.match(req, { ignoreSearch: true });
    const rede = fetch(req).then(r => { if (r.ok) c.put(req, r.clone()); return r; }).catch(() => salvo);
    return salvo || rede;
  }));
});
