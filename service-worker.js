// اسم الكاش — تم رفع الإصدار بعد تحديث الأصول
const CACHE_NAME = "musaid-cache-v3";

// الملفات الأساسية التي نخزّنها للعمل دون اتصال
const urlsToCache = [
  "/",
  "/static/css/app.css",
  "/static/js/app.js",
  "/static/images/logo.webp",
  "/static/images/student_icon.webp",
  "/static/images/teacher_icon.webp"
];

// التثبيت — نخزّن كل ملف على حدة حتى لا يفشل التثبيت كاملاً عند تعذّر أحدها
self.addEventListener("install", event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache =>
      Promise.allSettled(urlsToCache.map(url => cache.add(url)))
    ).then(() => self.skipWaiting())
  );
});

// التفعيل — حذف الكاش القديم
self.addEventListener("activate", event => {
  event.waitUntil(
    caches.keys().then(names =>
      Promise.all(names.filter(n => n !== CACHE_NAME).map(n => caches.delete(n)))
    ).then(() => self.clients.claim())
  );
});

// الجلب — شبكة أولاً لصفحات HTML مع رجوع للكاش، وكاش أولاً لبقية الأصول
self.addEventListener("fetch", event => {
  const req = event.request;
  if (req.method !== "GET") return;

  const accept = req.headers.get("accept") || "";
  const isHTML = accept.includes("text/html");
  const url = new URL(req.url);
  // كود الواجهة (CSS/JS) ليس مُبصَّماً بالمحتوى، لذا نستخدم الشبكة أولاً حتى يصل التحديث فوراً
  const isAppShell = url.origin === self.location.origin &&
    (url.pathname.startsWith("/static/css/") || url.pathname.startsWith("/static/js/"));

  if (isHTML || isAppShell) {
    // شبكة أولاً مع رجوع للكاش عند انقطاع الاتصال
    event.respondWith(
      fetch(req)
        .then(res => {
          const copy = res.clone();
          caches.open(CACHE_NAME).then(c => c.put(req, copy));
          return res;
        })
        .catch(() => caches.match(req).then(r => r || (isHTML ? caches.match("/") : r)))
    );
  } else {
    // كاش أولاً لبقية الأصول (صور، خطوط…)
    event.respondWith(
      caches.match(req).then(cached =>
        cached || fetch(req).then(res => {
          const copy = res.clone();
          caches.open(CACHE_NAME).then(c => c.put(req, copy));
          return res;
        }).catch(() => cached)
      )
    );
  }
});
