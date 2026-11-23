// Guarda os arquivos do editor no aparelho: depois do primeiro uso, o site abre e funciona sem internet.
// Os PDFs nunca passam por aqui (são processados na própria página e não são enviados).
const VERSAO = '20261123-111627';
self.addEventListener('message', e => { if (e.data === 'pular-espera') self.skipWaiting(); });
self.addEventListener('install', e => { self.skipWaiting(); e.waitUntil(caches.open(VERSAO).then(c => c.addAll(['./', './index.html', './motor.js', './manifest.webmanifest', './novidades.json'].map(u => new Request(u, { cache: 'reload' }))).catch(() => c.add('./index.html')))); });
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== VERSAO).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET' || u.origin !== location.origin) return;
  const guardar = (c, r) => { if (r.ok && r.status !== 206) e.waitUntil(c.put(e.request, r.clone()).catch(() => {})); };
  // /lib/: o cache já é por VERSAO, então usa o guardado sem revalidar; só vai à rede se faltar
  if (u.pathname.includes('/lib/')) {
    e.respondWith(caches.open(VERSAO).then(async c => {
      const salvo = await c.match(e.request, { ignoreSearch: true });
      if (salvo) return salvo;
      try { const r = await fetch(e.request); guardar(c, r); return r; } catch (err) { return Response.error(); }
    }));
    return;
  }
  // index.html, motor.js, novidades.json e o resto: usa o guardado (rápido e offline) e atualiza em segundo plano
  e.respondWith(caches.open(VERSAO).then(async c => {
    const salvo = await c.match(e.request, { ignoreSearch: true });
    const rede = fetch(e.request).then(r => { guardar(c, r); return r; }).catch(() => salvo || Response.error());
    if (salvo) { e.waitUntil(rede.catch(() => {})); return salvo; }
    return rede;
  }));
});
