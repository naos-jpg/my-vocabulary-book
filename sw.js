const CACHE_NAME = 'my-tango-v11';
const STATIC = [
  './', './index.html', './manifest.json',
  './ejdict.js', './ejdict_supp.js', './jaen.js', './ejdict_pos.js', './cefr_j.js', './exdict.js',
  './icons/icon-192.png', './icons/icon-512.png',
];
self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE_NAME).then(c => c.addAll(STATIC).catch(()=>{})));
  self.skipWaiting();
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys =>
    Promise.all(keys.filter(k=>k!==CACHE_NAME).map(k=>caches.delete(k)))
  ));
  self.clients.claim();
});
self.addEventListener('fetch', e => {
  const url = e.request.url;
  if (url.includes('googleapis.com')||url.includes('accounts.google.com')||
      url.includes('dictionaryapi.dev')||url.includes('mymemory.translated.net')||
      url.includes('wiktionary.org')||url.includes('jisho.org')||
      url.includes('script.google.com')) return;
  e.respondWith(
    caches.match(e.request).then(cached => cached || fetch(e.request).then(res => {
      if (res.ok && e.request.method==='GET') {
        caches.open(CACHE_NAME).then(c => c.put(e.request, res.clone()));
      }
      return res;
    }))
  );
});
