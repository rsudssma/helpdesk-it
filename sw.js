/**
 * Service worker Helpdesk IT.
 *
 * Strategi: hanya cangkang aplikasi (HTML, manifest, ikon) yang disimpan
 * offline. Data tiket TIDAK PERNAH di-cache, supaya petugas tidak pernah
 * melihat status tiket yang sudah basi.
 *
 * Naikkan nomor VERSI setiap kali index.html diubah, supaya perangkat
 * pegawai mengambil versi terbaru.
 */

const VERSI = 'helpdesk-v2';

const CANGKANG = [
  './',
  './index.html',
  './manifest.json',
  './icon-192.png',
  './icon-512.png'
];

self.addEventListener('install', function (e) {
  e.waitUntil(
    caches.open(VERSI)
      .then(function (c) { return c.addAll(CANGKANG); })
      .then(function () { return self.skipWaiting(); })
  );
});

self.addEventListener('activate', function (e) {
  e.waitUntil(
    caches.keys()
      .then(function (kunci) {
        return Promise.all(kunci.map(function (k) {
          return k === VERSI ? null : caches.delete(k);
        }));
      })
      .then(function () { return self.clients.claim(); })
  );
});

self.addEventListener('fetch', function (e) {
  const req = e.request;

  // Panggilan API selalu ke jaringan, tidak pernah dari cache.
  if (req.method !== 'GET' || req.url.indexOf('script.google.com') !== -1) return;

  // Halaman: jaringan dulu, cache sebagai cadangan saat offline.
  if (req.mode === 'navigate') {
    e.respondWith(
      fetch(req).catch(function () { return caches.match('./index.html'); })
    );
    return;
  }

  // Aset statis: cache dulu, ambil dari jaringan kalau belum ada.
  e.respondWith(
    caches.match(req).then(function (tersimpan) {
      return tersimpan || fetch(req);
    })
  );
});
