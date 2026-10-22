const CACHE_NAME = 'clealco-voos-v9';
const ASSETS_ESTATICOS = [
    './',
    './index.html',
    './manifest.json',
    './dados_injetados.js',
    './dados_mapa_voos.js',
    './leaflet.css',
    './leaflet.js',
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
    const url = event.request.url;

    // Para dados, html e scripts: NETWORK FIRST (busca sempre a versão mais recente na rede)
    if (url.includes('dados.xlsx') || url.includes('dados_injetados.js') || url.includes('index.html') || event.request.mode === 'navigate') {
        event.respondWith(
            fetch(event.request).then((response) => {
                if (response && response.status === 200) {
                    const responseClone = response.clone();
                    caches.open(CACHE_NAME).then((cache) => cache.put(event.request, responseClone));
                }
                return response;
            }).catch(() => caches.match(event.request))
        );
        return;
    }

    // Para outros arquivos estáticos (ícones, logos): Cache com fallback para rede
    event.respondWith(
        caches.match(event.request).then((response) => {
            return response || fetch(event.request);
        })
    );
});