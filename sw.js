// アプリ本体と辞書を別々に保存する
//  ・APP_CACHE … index.html など（アプリを直した時はこの番号を上げる）
//  ・DICT_CACHE … 辞書（約46MB）。辞書ファイルを直した時だけこの番号を上げる
//    → アプリを更新しても、辞書はダウンロードし直さない
const APP_CACHE  = 'my-tango-app-v14';
const DICT_CACHE = 'my-tango-dict-v1';
const APP_FILES  = ['./', './index.html', './manifest.json', './icons/icon-192.png', './icons/icon-512.png'];
const DICT_FILES = ['./ejdict.js', './ejdict_supp.js', './ejdict_pos.js', './cefr_j.js', './exdict.js', './ejdict_pos_add.js'];
const KEEP = [APP_CACHE, DICT_CACHE];

self.addEventListener('install', e => {
  e.waitUntil((async () => {
    // アプリ本体: 必ずサーバーの最新版（ブラウザに残った古いファイルは使わない）
    const app = await caches.open(APP_CACHE);
    await Promise.all(APP_FILES.map(u => fetch(new Request(u, { cache: 'reload' }))
      .then(r => r.ok ? app.put(u, r) : null).catch(() => null)));
    // 辞書: 前の保存場所にあれば移すだけ（ダウンロードしない）。無い時だけ取りに行く
    const dict = await caches.open(DICT_CACHE);
    const oldNames = (await caches.keys()).filter(k => k !== DICT_CACHE && k !== APP_CACHE);
    const sameDictVer = n => n.startsWith('my-tango-dict-') ? false : true;   // 古い辞書版の保存場所は使わない
    await Promise.all(DICT_FILES.map(async u => {
      if (await dict.match(u)) return;
      for (const n of oldNames.filter(sameDictVer)) {
        const hit = await (await caches.open(n)).match(u);
        if (hit) { await dict.put(u, hit); return; }
      }
      try { const r = await fetch(new Request(u, { cache: 'reload' })); if (r.ok) await dict.put(u, r); } catch {}
    }));
  })());
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys =>
    Promise.all(keys.filter(k => !KEEP.includes(k)).map(k => caches.delete(k)))
  ));
  self.clients.claim();
});

function isPage(req) {
  if (req.mode === 'navigate') return true;
  const p = new URL(req.url).pathname;
  return p.endsWith('/') || p.endsWith('/index.html') || p.endsWith('/manifest.json');
}
function isDict(req) {
  const name = './' + new URL(req.url).pathname.split('/').pop();
  return DICT_FILES.includes(name);
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
      const cache = await caches.open(APP_CACHE);
      try {
        const res = await Promise.race([fetch(req, { cache: 'no-cache' }), new Promise((_, rej) => setTimeout(() => rej(0), 4000))]);
        if (res && res.ok) { cache.put(req, res.clone()); return res; }
        throw 0;
      } catch {
        const hit = await cache.match(req, { ignoreSearch: true }) || await cache.match('./index.html');
        return hit || fetch(req);
      }
    })());
    return;
  }
  // 辞書・アイコンなど: 保存済み優先
  e.respondWith((async () => {
    const cacheName = isDict(req) ? DICT_CACHE : APP_CACHE;
    const cache = await caches.open(cacheName);
    const hit = await cache.match(req, { ignoreSearch: true });
    if (hit) return hit;
    const res = await fetch(req);
    if (res.ok) cache.put(req, res.clone());
    return res;
  })());
});
