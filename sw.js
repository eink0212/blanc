// Blanc - Service Worker
// キャッシュ方針:
//   - index.html … network-first（更新をすぐ反映。オフライン時のみキャッシュ）
//   - アイコン/manifest … cache-first（滅多に変わらない）
//   - GAS API（script.google.com / googleusercontent.com）… 一切キャッシュしない
// 中身を変えたら CACHE_VERSION を上げること。

var CACHE_VERSION = 'blanc-v3';
var SHELL = [
  './',
  './index.html',
  './manifest.json',
  './icon-192.png',
  './icon-512.png',
  './icon-maskable-512.png',
  './apple-touch-icon.png'
];

self.addEventListener('install', function (event) {
  event.waitUntil(
    caches.open(CACHE_VERSION)
      .then(function (cache) { return cache.addAll(SHELL); })
      .then(function () { return self.skipWaiting(); })
  );
});

self.addEventListener('activate', function (event) {
  event.waitUntil(
    caches.keys()
      .then(function (keys) {
        return Promise.all(keys.map(function (k) {
          return k === CACHE_VERSION ? null : caches.delete(k);
        }));
      })
      .then(function () { return self.clients.claim(); })
  );
});

// ページから postMessage({type:'SKIP_WAITING'}) で即時更新できるように
self.addEventListener('message', function (event) {
  if (event.data && event.data.type === 'SKIP_WAITING') self.skipWaiting();
});

self.addEventListener('fetch', function (event) {
  var req = event.request;
  if (req.method !== 'GET') return;                     // GAS への POST は素通し

  var url = new URL(req.url);
  if (url.origin !== self.location.origin) return;      // 外部ドメインは素通し

  var isDoc = req.mode === 'navigate' || url.pathname.endsWith('/') ||
              url.pathname.endsWith('index.html');

  if (isDoc) {
    // network-first
    event.respondWith(
      fetch(req)
        .then(function (res) {
          // ステータスを確認してからキャッシュする。
          // 確認せずに入れると、デプロイ中の一瞬に 404 を掴んだとき
          // その 404 が index.html として焼き付き、CACHE_VERSION を
          // 上げるまでオフライン時に 404 が出続ける。
          // redirected なレスポンスはナビゲーションに返すと画面が真っ白になるため除外。
          if (res && res.ok && res.status === 200 && !res.redirected) {
            var copy = res.clone();
            caches.open(CACHE_VERSION).then(function (c) { c.put('./index.html', copy); });
          }
          return res;
        })
        .catch(function () {
          return caches.match('./index.html').then(function (hit) {
            return hit || caches.match('./');
          });
        })
    );
    return;
  }

  // cache-first
  event.respondWith(
    caches.match(req).then(function (hit) {
      return hit || fetch(req).then(function (res) {
        if (res && res.status === 200 && res.type === 'basic') {
          var copy = res.clone();
          caches.open(CACHE_VERSION).then(function (c) { c.put(req, copy); });
        }
        return res;
      });
    })
  );
});
