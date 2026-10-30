// Service worker: เก็บไฟล์หน้าแอปไว้ให้เปิดเร็ว (ข้อมูลประเมินดึงจากเซิร์ฟเวอร์ทุกครั้ง)
// เมื่อแก้ไฟล์แอป ให้เปลี่ยนเลขเวอร์ชันด้านล่าง ผู้ใช้จะได้ไฟล์ใหม่อัตโนมัติ
var CACHE = 'ris360-v1';
var SHELL = ['./', 'index.html', 'config.js', 'manifest.webmanifest', 'icons/icon-192.png', 'icons/icon-512.png'];

self.addEventListener('install', function (e) {
  e.waitUntil(caches.open(CACHE).then(function (c) { return c.addAll(SHELL); }).then(function () { return self.skipWaiting(); }));
});
self.addEventListener('activate', function (e) {
  e.waitUntil(caches.keys().then(function (keys) {
    return Promise.all(keys.filter(function (k) { return k !== CACHE; }).map(function (k) { return caches.delete(k); }));
  }).then(function () { return self.clients.claim(); }));
});
// ไฟล์ของแอป: ดึงใหม่ก่อน ถ้าออฟไลน์ค่อยใช้ของในเครื่อง · คำขอไปเซิร์ฟเวอร์อื่น (API, ฟอนต์) ไม่ยุ่ง
self.addEventListener('fetch', function (e) {
  var url = new URL(e.request.url);
  if (e.request.method !== 'GET' || url.origin !== location.origin) return;
  e.respondWith(fetch(e.request).then(function (res) {
    var copy = res.clone();
    caches.open(CACHE).then(function (c) { c.put(e.request, copy); });
    return res;
  }).catch(function () { return caches.match(e.request).then(function (r) { return r || caches.match('index.html'); }); }));
});
