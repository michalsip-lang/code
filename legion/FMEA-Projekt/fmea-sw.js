/**
 * Service Worker pro FMEA aplikaci
 * Umožňuje offline funkčnost a caching
 * 
 * Registrace v HTML:
 * if ('serviceWorker' in navigator) {
 *     navigator.serviceWorker.register('fmea-sw.js');
 * }
 */

const CACHE_NAME = 'fmea-v4-detail-escape-fix';
const URLS_TO_CACHE = [
    './',
    './FRM-FMEA-Formular.html',
    './FRM-FMEA-Formular.js',
    './FRM-FMEA-Config.js',
    './FRM-FMEA-Formular-TEST.html'
];

// Instalace Service Workera
self.addEventListener('install', event => {
    console.log('Service Worker instalace...');
    
    event.waitUntil(
        caches.open(CACHE_NAME).then(cache => {
            console.log('Cache otevřena');
            return cache.addAll(URLS_TO_CACHE).catch(err => {
                console.log('Některé soubory se nepodařilo cachovat:', err);
            });
        })
    );
    
    self.skipWaiting(); // Aktivuj ihned
});

// Aktivace Service Workera
self.addEventListener('activate', event => {
    console.log('Service Worker aktivován');
    
    event.waitUntil(
        caches.keys().then(cacheNames => {
            return Promise.all(
                cacheNames.map(cacheName => {
                    if (cacheName !== CACHE_NAME) {
                        console.log('Mažu starou cache:', cacheName);
                        return caches.delete(cacheName);
                    }
                })
            );
        })
    );
    
    self.clients.claim(); // Vezmi kontrolu hned
});

// Fetch handling - offline/online strategie
self.addEventListener('fetch', event => {
    const { request } = event;
    const url = new URL(request.url);
    
    // Ignoruj SharePoint API volání - vždy síť
    if (url.pathname.includes('/_api/')) {
        return event.respondWith(
            fetch(request)
                .catch(err => {
                    // Pokud API selže, vrať info o offline
                    return new Response(JSON.stringify({
                        error: 'Offline - API není dostupné',
                        cached: true
                    }), {
                        status: 503,
                        statusText: 'Service Unavailable',
                        headers: new Headers({
                            'Content-Type': 'application/json'
                        })
                    });
                })
        );
    }
    
    // Pro ostatní soubory: cache-first strategie
    event.respondWith(
        caches.match(request).then(response => {
            if (response) {
                console.log('Z cache:', request.url);
                return response;
            }
            
            // Pokud není v cache, stáhni ze sítě
            return fetch(request)
                .then(response => {
                    // Pouze úspěšné responsy cachuj
                    if (!response || response.status !== 200) {
                        return response;
                    }
                    
                    // Cachuj kopii
                    const responseToCache = response.clone();
                    caches.open(CACHE_NAME).then(cache => {
                        cache.put(request, responseToCache);
                    });
                    
                    return response;
                })
                .catch(err => {
                    console.log('Fetch chyba:', err);
                    
                    // Offline fallback - vrať offline stránku
                    return caches.match('./FRM-FMEA-Formular-TEST.html').then(response => {
                        return response || new Response('Offline - soubor není cachován', {
                            status: 503,
                            statusText: 'Service Unavailable'
                        });
                    });
                });
        })
    );
});

// Komunikace s aplikací - offline notifikace
self.addEventListener('message', event => {
    if (event.data && event.data.type === 'SKIP_WAITING') {
        self.skipWaiting();
    }
});
