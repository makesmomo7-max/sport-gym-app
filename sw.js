/* momo fit — 最小PWA用（オフラインキャッシュなし。後から拡張可） */
self.addEventListener("install", function (e) {
  self.skipWaiting();
});
self.addEventListener("activate", function (e) {
  e.waitUntil(self.clients.claim());
});
