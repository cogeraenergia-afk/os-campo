/* Guarda o app no celular para abrir sem sinal.
   Com internet: busca sempre a versão nova (e atualiza a cópia guardada).
   Sem internet: usa a cópia guardada. */
const VERSAO = 'osc-v2';
const ARQUIVOS = ['./', 'index.html', 'comum.js', 'comum.css', 'config.js', 'manifest.webmanifest', 'icone-180.png', 'icone-192.png', 'icone-512.png'];
self.addEventListener('install', e => { e.waitUntil(caches.open(VERSAO).then(c => c.addAll(ARQUIVOS.map(u => new Request(u, { cache: 'reload' })))).then(() => self.skipWaiting())); });
self.addEventListener('activate', e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== VERSAO).map(k => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET' || new URL(req.url).origin !== location.origin) return;
  e.respondWith((async () => {
    const c = await caches.open(VERSAO);
    try {
      const r = await Promise.race([fetch(req, { cache: 'no-cache' }), new Promise((_, n) => setTimeout(() => n(new Error('lento')), 6000))]);
      if (r.ok) c.put(req, r.clone());
      return r;
    } catch (err) {
      const salvo = await c.match(req, { ignoreSearch: true });
      if (salvo) return salvo;
      throw err;
    }
  })());
});
