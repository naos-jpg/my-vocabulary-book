const CACHE_NAME = 'my-tango-v12';
// 辞書などの大きいファイル（キャッシュ優先）
const STATIC = [
  './', './index.html', './manifest.json',
  './ejdict.js', './ejdict_supp.js', './jaen.js', './ejdict_pos.js', './cefr_j.js', './exdict.js',
  './icons/icon-192.png', './icons/icon-512.png',
];
self.addEventListener('install', e => {
  // cache:'reload' … ブラウザに残った古いファイルではなく、必ずサーバーの最新版を保存する
  e.waitUntil(caches.open(CACHE_NAME).then(c =>
    Promise.all(STATIC.map(u => fetch(new Request(u, { cache: 'reload' }))
      .then(r => r.ok ? c.put(u, r) : null).catch(() => null)))
  ));
  self.skipWaiting();
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys =>
    Promise.all(keys.filter(k=>k!==CACHE_NAME).map(k=>caches.delete(k)))
  ));
  self.clients.claim();
});
function isPage(req) {
  if (req.mode === 'navigate') return true;
  const p = new URL(req.url).pathname;
  return p.endsWith('/') || p.endsWith('/index.html') || p.endsWith('/manifest.json');
}
self.addEventListener('fetch', e => {
  const req = e.request; const url = req.url;
  if (req.method !== 'GET') return;
  if (url.includes('googleapis.com')||url.includes('accounts.google.com')||
      url.includes('dictionaryapi.dev')||url.includes('mymemory.translated.net')||
      url.includes('wiktionary.org')||url.includes('jisho.org')||
      url.includes('script.google.com')) return;
  if (isPage(req)) {
    // アプリ本体はネット優先（更新がすぐ届く）。4秒で返らない・オフラインの時は保存済みを使う
    e.respondWith((async () => {
      const cache = await caches.open(CACHE_NAME);
      try {
        const net = fetch(req, { cache: 'no-cache' });
        const res = await Promise.race([net, new Promise((_, rej) => setTimeout(() => rej(0), 4000))]);
        if (res && res.ok) { cache.put(req, res.clone()); return res; }
        throw 0;
      } catch {
        const hit = await cache.match(req, { ignoreSearch: true }) || await cache.match('./index.html');
        return hit || fetch(req);
      }
    })());
    return;
  }
  e.respondWith(
    caches.match(req).then(cached => cached || fetch(req).then(res => {
      if (res.ok) caches.open(CACHE_NAME).then(c => c.put(req, res.clone()));
      return res;
    }))
  );
});
