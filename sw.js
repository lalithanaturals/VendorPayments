/* Lalitha Naturals Payables – offline support.
   Change CACHE_VERSION on every release so phones pick up the new app. */
const CACHE_VERSION = 'ln-payables-v8.2';
const APP_SHELL = ['./', './index.html', './manifest.json', './icon-192.png', './icon-512.png', './icon-180.png', './icon-maskable-512.png', './logo.png'];
const CDN_LIBS = [
    'https://cdn.tailwindcss.com',
    'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js',
    'https://unpkg.com/lucide@0.460.0/dist/umd/lucide.min.js'
];

self.addEventListener('install', event => {
    event.waitUntil((async () => {
        const cache = await caches.open(CACHE_VERSION);
        await cache.addAll(APP_SHELL);
        // CDN files are cached one by one so one slow CDN can't block install
        await Promise.all(CDN_LIBS.map(async url => {
            try { await cache.add(new Request(url, { mode: 'no-cors' })); } catch (e) { /* cached later on first use */ }
        }));
    })());
});

self.addEventListener('activate', event => {
    event.waitUntil((async () => {
        const keys = await caches.keys();
        await Promise.all(keys.filter(k => k.startsWith('ln-payables-') && k !== CACHE_VERSION).map(k => caches.delete(k)));
        await self.clients.claim();
    })());
});

// The page asks the waiting worker to take over when the user taps "Reload"
self.addEventListener('message', event => {
    if (event.data === 'SKIP_WAITING') self.skipWaiting();
});

self.addEventListener('fetch', event => {
    const req = event.request;
    if (req.method !== 'GET') return;                       // backups (POST) go straight to Google
    const url = new URL(req.url);
    if (url.hostname.endsWith('script.google.com') || url.hostname.endsWith('googleusercontent.com')) return;

    // App page: network first (so updates arrive when the app is reopened), cache when offline
    if (req.mode === 'navigate') {
        event.respondWith((async () => {
            try {
                const fresh = await fetch(req);
                const cache = await caches.open(CACHE_VERSION);
                cache.put('./index.html', fresh.clone());
                return fresh;
            } catch (e) {
                const cache = await caches.open(CACHE_VERSION);
                return (await cache.match('./index.html')) || (await cache.match('./')) || Response.error();
            }
        })());
        return;
    }

    // Everything else (icons, libraries, fonts): cache first, then network and remember it
    event.respondWith((async () => {
        const cache = await caches.open(CACHE_VERSION);
        const hit = await cache.match(req, { ignoreVary: true }) || await cache.match(req.url, { ignoreVary: true });
        if (hit) return hit;
        try {
            const res = await fetch(req);
            if (res && (res.ok || res.type === 'opaque')) cache.put(req, res.clone());
            return res;
        } catch (e) {
            return hit || Response.error();
        }
    })());
});
