// Service Worker بسيط — مطلوب من المتصفح عشان يعتبر الموقع "تطبيق قابل للتثبيت"
self.addEventListener('install', (e) => {
    self.skipWaiting();
});

self.addEventListener('activate', (e) => {
    self.clients.claim();
});

self.addEventListener('fetch', (e) => {
    // مفيش تخزين مؤقت (cache) حاليًا — بس وجود هذا الحدث كافٍ لتفعيل خاصية التثبيت
});
