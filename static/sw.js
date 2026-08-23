// ==========================================
// YURT SOSYAL PWA SERVICE WORKER
// ==========================================

const CACHE_NAME = 'yurt-sosyal-v2.3-messaging-align-fix';
const STATIC_ASSETS = [
    '/',
    '/static/css/style.css',
    '/static/manifest.json',
    '/static/img/icons/icon-192.png',
    '/static/img/icons/icon-512.png',
    '/static/img/icons/badge-72.png',
    'https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css',
    'https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css',
    'https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js'
];

// Install: Cache essential assets
self.addEventListener('install', (event) => {
    event.waitUntil(
        caches.open(CACHE_NAME).then((cache) => {
            return cache.addAll(STATIC_ASSETS).catch((err) => {
                console.warn('PWA: Some static assets failed to cache:', err);
            });
        }).then(() => self.skipWaiting())
    );
});

// Activate: Clean up old caches
self.addEventListener('activate', (event) => {
    event.waitUntil(
        caches.keys().then((keys) => {
            return Promise.all(
                keys.map((key) => {
                    if (key !== CACHE_NAME) {
                        return caches.delete(key);
                    }
                })
            );
        }).then(() => self.clients.claim())
    );
});

// Fetch: Network First with Cache Fallback for dynamic pages; Cache First for static
self.addEventListener('fetch', (event) => {
    const request = event.request;
    const url = new URL(request.url);

    // Don't cache POST, PUT, DELETE, socket.io or API requests
    if (request.method !== 'GET' || url.pathname.startsWith('/socket.io') || url.pathname.startsWith('/api/')) {
        return;
    }

    // Static assets (CSS, JS, Fonts, Images) -> Cache First / Stale While Revalidate
    if (url.pathname.startsWith('/static/') || url.hostname.includes('cdn.jsdelivr.net')) {
        event.respondWith(
            caches.match(request).then((cachedResponse) => {
                if (cachedResponse) {
                    // Update cache in background
                    fetch(request).then((networkResponse) => {
                        if (networkResponse && networkResponse.status === 200) {
                            caches.open(CACHE_NAME).then((cache) => cache.put(request, networkResponse));
                        }
                    }).catch(() => {});
                    return cachedResponse;
                }
                return fetch(request).then((networkResponse) => {
                    if (networkResponse && networkResponse.status === 200) {
                        const responseClone = networkResponse.clone();
                        caches.open(CACHE_NAME).then((cache) => cache.put(request, responseClone));
                    }
                    return networkResponse;
                });
            })
        );
        return;
    }

    // HTML / Page Navigation -> Network First
    event.respondWith(
        fetch(request)
            .then((networkResponse) => {
                if (networkResponse && networkResponse.status === 200) {
                    const responseClone = networkResponse.clone();
                    caches.open(CACHE_NAME).then((cache) => cache.put(request, responseClone));
                }
                return networkResponse;
            })
            .catch(() => {
                return caches.match(request).then((cachedResponse) => {
                    if (cachedResponse) return cachedResponse;
                    if (request.headers.get('accept').includes('text/html')) {
                        return caches.match('/');
                    }
                });
            })
    );
});

// ==========================================
// WEB PUSH & BİLDİRİM İŞLEYİCİSİ
// ==========================================

self.addEventListener('push', (event) => {
    let data = {};
    try {
        data = event.data ? event.data.json() : {};
    } catch (e) {
        data = { title: 'Yurt Sosyal', body: event.data ? event.data.text() : 'Yeni bildiriminiz var.' };
    }

    const title = data.title || 'YURT. Bildirimi 🔔';
    const options = {
        body: data.body || data.message || 'Yeni bir duyuru veya yoklama bildirimi var.',
        icon: data.icon || '/static/img/icons/icon-192.png',
        badge: '/static/img/icons/badge-72.png',
        data: {
            url: data.url || data.link || '/'
        },
        tag: data.tag || 'yurt-notification',
        renotify: true
    };

    event.waitUntil(
        self.registration.showNotification(title, options)
    );
});

self.addEventListener('notificationclick', (event) => {
    event.notification.close();

    if (event.action === 'close') {
        return;
    }

    const targetUrl = event.notification.data && event.notification.data.url ? event.notification.data.url : '/';

    event.waitUntil(
        clients.matchAll({ type: 'window', includeUncontrolled: true }).then((windowClients) => {
            // If already open, focus and navigate
            for (let client of windowClients) {
                if ('focus' in client) {
                    client.navigate(targetUrl);
                    return client.focus();
                }
            }
            // Otherwise open a new window
            if (clients.openWindow) {
                return clients.openWindow(targetUrl);
            }
        })
    );
});
