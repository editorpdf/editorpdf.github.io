// Guarda os arquivos do editor no aparelho: depois do primeiro uso, o site abre e funciona sem internet.
// Os PDFs nunca passam por aqui (são processados na própria página e não são enviados).
const VERSAO = '20261018-154525';
self.addEventListener('install', e => { self.skipWaiting(); e.waitUntil(caches.open(VERSAO).then(c => c.addAll(['./', './index.html', './motor.js']))); });
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== VERSAO).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET' || u.origin !== location.origin) return;
  // usa o que está guardado (rápido e offline) e atualiza em segundo plano
  e.respondWith(caches.open(VERSAO).then(async c => {
    const salvo = await c.match(e.request, { ignoreSearch: true });
    const rede = fetch(e.request).then(r => { if (r.ok) c.put(e.request, r.clone()); return r; }).catch(() => salvo);
    return salvo || rede;
  }));
});
