const CACHE_NAME = 'clealco-voos-v3';
const ASSETS_ESTATICOS = [
    './',
    './index.html',
    './manifest.json',
    './dados_injetados.js',
    './logo_clealco.png',
    './icon-192.png',
    './icon-512.png',
    './apple-touch-icon.png',
    './favicon.png'
];

self.addEventListener('install', (event) => {
    event.waitUntil(
        caches.open(CACHE_NAME).then((cache) => {
            return cache.addAll(ASSETS_ESTATICOS);
        })
    );
    self.skipWaiting();
});

self.addEventListener('activate', (event) => {
    event.waitUntil(
        caches.keys().then((keys) => {
            return Promise.all(
                keys.map((k) => {
                    if (k !== CACHE_NAME) return caches.delete(k);
                })
            );
        })
    );
    self.clients.claim();
});

self.addEventListener('fetch', (event) => {
    // Para dados.xlsx e API, sempre buscar na rede primeiro
    if (event.request.url.includes('dados.xlsx')) {
        event.respondWith(
            fetch(event.request).catch(() => caches.match(event.request))
        );
        return;
    }

    event.respondWith(
        caches.match(event.request).then((response) => {
            return response || fetch(event.request);
        })
    );
});