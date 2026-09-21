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

const VERSI = 'helpdesk-v9';

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

  // Panggilan API, Firebase, dan situs lain selalu langsung ke jaringan.
  if (req.method !== 'GET' || new URL(req.url).origin !== self.location.origin) return;

  // Halaman: selalu tanyakan versi terbaru ke server (GitHub Pages menyuruh
  // HP menyimpan halaman 10 menit; tanpa 'no-cache', perbaikan di GitHub
  // bisa tidak terlihat di APK). Cache hanya dipakai saat offline.
  if (req.mode === 'navigate') {
    e.respondWith(
      fetch(req.url, { cache: 'no-cache', credentials: 'same-origin' })
        .catch(function () { return caches.match('./index.html'); })
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

// ------------------------------------------------------------
//  NOTIFIKASI
//  Pesan dikirim Code.gs lewat Firebase dalam bentuk data:
//  { data: { judul, isi, id } }. Ditampilkan di sini, jadi tetap
//  muncul walaupun aplikasi sedang tertutup.
// ------------------------------------------------------------

self.addEventListener('push', function (e) {
  let p = {};
  try { p = e.data ? e.data.json() : {}; }
  catch (x) { p = { data: { isi: e.data ? e.data.text() : '' } }; }
  const d = p.data || p.notification || p;
  const id = d.id || '';

  const opsi = {
    body: d.isi || d.body || '',
    icon: 'icon-192.png',
    data: { id: id }
  };
  // Kabar baru untuk tiket yang sama menggantikan kabar lamanya
  if (id) { opsi.tag = id; opsi.renotify = true; }

  e.waitUntil(Promise.all([
    self.registration.showNotification(d.judul || d.title || 'Helpdesk IT', opsi),
    // Kalau aplikasi sedang terbuka, segarkan daftarnya
    self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then(function (cs) {
      cs.forEach(function (c) { c.postMessage({ tipe: 'segarkan', id: id }); });
    })
  ]));
});

self.addEventListener('notificationclick', function (e) {
  e.notification.close();
  const id = (e.notification.data || {}).id || '';
  const tujuan = new URL(id ? './?tiket=' + encodeURIComponent(id) : './', self.registration.scope).href;

  e.waitUntil(
    self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then(function (cs) {
      for (let i = 0; i < cs.length; i++) {
        if (cs[i].url.indexOf(self.registration.scope) === 0 && 'focus' in cs[i]) {
          if (id) cs[i].postMessage({ tipe: 'buka', id: id });
          return cs[i].focus();
        }
      }
      return self.clients.openWindow(tujuan);
    })
  );
});
