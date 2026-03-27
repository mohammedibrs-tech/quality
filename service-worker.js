const CACHE = ‘noura-quality-v1’;

self.addEventListener(‘install’, e => { self.skipWaiting(); });
self.addEventListener(‘activate’, e => { e.waitUntil(clients.claim()); });

self.addEventListener(‘push’, e => {
if (!e.data) return;
let data;
try { data = e.data.json(); } catch { data = { title: ‘نورة للجودة’, body: e.data.text() }; }

const title   = data.title || ‘نورة للجودة’;
const options = {
body:    data.body    || ‘’,
icon:    data.icon    || ‘/icon-192.png’,
badge:   data.badge   || ‘/icon-192.png’,
tag:     data.tag     || ‘noura-notif’,
data:    data.url     || ‘/’,
dir:     ‘rtl’,
lang:    ‘ar’,
vibrate: [200, 100, 200],
requireInteraction: data.urgent || false,
actions: data.actions || [],
};

e.waitUntil(self.registration.showNotification(title, options));
});

self.addEventListener(‘notificationclick’, e => {
e.notification.close();
const url = e.notification.data || ‘/’;
e.waitUntil(
clients.matchAll({ type: ‘window’, includeUncontrolled: true }).then(list => {
for (const c of list) {
if (c.url.includes(self.location.origin) && ‘focus’ in c) return c.focus();
}
return clients.openWindow(url);
})
);
});