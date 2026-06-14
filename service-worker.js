const CACHE_NAME = 'royyek-vault-v1.3';
const assets = [
  './',
  './index.html',
  './manifest.json',
  'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css'
];

// 1. Install Service Worker & Simpan Asset Utama ke Cache
self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE_NAME).then(cache => {
      console.log('🤖 PWA Cache dialokasikan dengan aman!');
      return cache.addAll(assets);
    })
  );
  self.skipWaiting();
});

// 2. Bersihkan Cache Lama jika ada Update Versi
self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys => {
      return Promise.all(
        keys.map(key => {
          if (key !== CACHE_NAME) {
            console.log('🤖 Membersihkan cache versi lama...');
            return caches.delete(key);
          }
        })
      );
    })
  );
  self.clients.claim();
});

// 3. Strategi Fetch: Ambil dari Internet dulu, kalau offline/gagal baru ambil dari Cache
// Ini sangat penting agar Web App Bos bisa selalu menembak API Google Sheets secara real-time saat online!
self.addEventListener('fetch', e => {
  // Biarkan request ke Google Apps Script (apiUrl) langsung lewat tanpa dihadang cache lokal
  if (e.request.url.includes('script.google.com')) {
    return;
  }

  e.respondWith(
    fetch(e.request)
      .then(response => {
        // Jika berhasil mengambil data terbaru dari internet, jalankan normal
        return response;
      })
      .catch(() => {
        // Jika internet putus (offline), ambil cadangan tampilan dari cache agar aplikasi tidak blank
        return caches.match(e.request).then(cachedResponse => {
          return cachedResponse || caches.match('./index.html');
        });
      })
  );
});
