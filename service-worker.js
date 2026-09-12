// Service worker minimal - syarat wajib browser buat nganggep situs ini
// "installable" sebagai PWA. Fokusnya cuma bikin app-shell (index.html)
// kebuka cepat & tetap bisa dibuka walau sinyal lagi jelek, BUKAN buat
// bikin form ini bisa dipakai penuh secara offline (submit izin tetap
// butuh koneksi internet karena nyimpen ke Google Sheets).
//
// PENTING: service worker ini scope-nya seluruh domain (bukan cuma
// index.html yang manggil register()-nya) -- jadi admin.html juga ikut
// "dijagain" sama dia begitu ke-install. Makanya admin.html WAJIB ikut
// di-precache & punya fallback sendiri, kalau enggak refresh admin.html
// pas koneksi sempet putus bisa nyasar/gagal kebuka.

const CACHE_NAME = 'sejasa-home-team-v2';
const APP_SHELL = [
  './',
  './index.html',
  './admin.html',
  './manifest.json',
  './icon-192.png',
  './icon-512.png'
];

self.addEventListener('install', function(event) {
  event.waitUntil(
    caches.open(CACHE_NAME).then(function(cache) {
      return cache.addAll(APP_SHELL);
    })
  );
  self.skipWaiting();
});

self.addEventListener('activate', function(event) {
  event.waitUntil(
    caches.keys().then(function(keys) {
      return Promise.all(
        keys.filter(function(key) { return key !== CACHE_NAME; })
            .map(function(key) { return caches.delete(key); })
      );
    })
  );
  self.clients.claim();
});

// Strategi: network-first buat HTML (biar selalu dapet versi terbaru kalau
// online), fallback ke cache kalau offline. Buat asset statis (icon dll),
// cache-first biar cepat.
self.addEventListener('fetch', function(event) {
  const req = event.request;

  // Jangan cache/intercept request ke Apps Script (API call) sama sekali -
  // itu harus selalu langsung ke jaringan, gak boleh ke-cache.
  if (req.url.indexOf('script.google.com') !== -1) {
    return;
  }

  if (req.mode === 'navigate') {
    // Fallback-nya harus ke HALAMAN YANG SAMA yang lagi diminta (admin.html
    // tetap admin.html, index.html tetap index.html) -- bukan selalu
    // dilempar ke index.html kayak sebelumnya, biar gak nyasar pas refresh.
    const isAdmin = req.url.indexOf('admin.html') !== -1;
    event.respondWith(
      fetch(req).catch(function() {
        return caches.match(isAdmin ? './admin.html' : './index.html');
      })
    );
    return;
  }

  event.respondWith(
    caches.match(req).then(function(cached) {
      return cached || fetch(req);
    })
  );
});
