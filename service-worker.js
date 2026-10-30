const CACHE_NAME = 'clealco-voos-v226';
const ASSETS_ESTATICOS = [
    './',
    './index.html',
    './manifest.json',
    './dados_injetados.js',
    './dados_mapa_voos.js',
    './dados_mapas_safra.js',
    './dados_historico_falhas.js',
    './leaflet.css',
    './leaflet.js',
    './logo_clealco.png',
    './logo_clealco_oficial.png',
    './logo_login_clealco.jpg',
    './icon-192.png',
    './icon-512.png',
    './icon-maskable-512.png',
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

    // Estratégia Network-First estrita com { cache: 'no-store' } para arquivos de dados, HTML e navegação
    const isDynamicData = url.includes('dados.xlsx') ||
                          url.includes('dados_injetados.js') ||
                          url.includes('dados_mapa_voos.js') ||
                          url.includes('dados_historico_falhas.js') ||
                          url.includes('dados_mapas_safra.js') ||
                          url.includes('index.html') ||
                          event.request.mode === 'navigate';

    if (isDynamicData) {
        event.respondWith(
            fetch(event.request, { cache: 'no-store' })
                .then((response) => {
                    if (response && response.status === 200) {
                        const responseClone = response.clone();
                        caches.open(CACHE_NAME).then((cache) => cache.put(event.request, responseClone));
                    }
                    return response;
                })
                .catch(() => caches.match(event.request))
        );
        return;
    }

    // Estratégia Cache-First para os restantes ficheiros estáticos (ícones, logos, CSS, bibliotecas)
    event.respondWith(
        caches.match(event.request).then((cachedResponse) => {
            if (cachedResponse) {
                return cachedResponse;
            }
            return fetch(event.request).then((networkResponse) => {
                if (networkResponse && networkResponse.status === 200) {
                    const responseClone = networkResponse.clone();
                    caches.open(CACHE_NAME).then((cache) => cache.put(event.request, responseClone));
                }
                return networkResponse;
            });
        })
    );
});