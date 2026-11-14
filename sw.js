// B4 레슨 노트 — 서비스 워커
// 앱 화면 파일만 저장해 두고(빠른 실행), 레슨 데이터는 항상 새로 받아옵니다.
const VERSION = 'b4-v2'; // 화면 파일을 바꿀 때 숫자를 올리면 폰에서도 새로 받습니다
const SHELL = [
  './',
  './index.html',
  './style.css',
  './app.js',
  './config.js',
  './manifest.webmanifest',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './icons/apple-touch-icon.png'
];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(VERSION).then((c) => c.addAll(SHELL.map((u) => new Request(u, { cache: 'reload' })))).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== VERSION).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (e) => {
  const url = new URL(e.request.url);
  // 다른 주소(Apps Script 데이터 요청 등)는 관여하지 않음
  if (url.origin !== self.location.origin || e.request.method !== 'GET') return;

  // 네트워크 우선, 실패하면 저장본
  e.respondWith(
    // cache: 'no-cache' → 휴대폰에 저장된 파일이 있어도 항상 서버에 최신인지 확인
    fetch(url.href, { cache: 'no-cache', credentials: 'same-origin' })
      .then((res) => {
        if (res.ok) { const copy = res.clone(); caches.open(VERSION).then((c) => c.put(e.request, copy)); }
        return res;
      })
      .catch(() => caches.match(e.request, { ignoreSearch: true }))
  );
});
