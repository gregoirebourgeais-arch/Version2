const CACHE_NAME = "atelier-ppnc-v2";

const ASSETS = [
  "./",
  "./index.html",
  "./style.css",
  "./app.js",
  "./planning.js",
  "./manifest.json",
  "./icon-192.png",
  "./icon-512.png",
  "https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js",
  "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"
];

// Installation du service worker et mise en cache des ressources
self.addEventListener("install", (event) => {
  console.log("[Service Worker] Installation...");
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      console.log("[Service Worker] Mise en cache des ressources");
      return cache.addAll(ASSETS);
    })
  );
  // Force le nouveau service worker à devenir actif immédiatement
  self.skipWaiting();
});

// Activation et nettoyage des anciens caches
self.addEventListener("activate", (event) => {
  console.log("[Service Worker] Activation...");
  event.waitUntil(
    caches.keys().then((keys) => {
      return Promise.all(
        keys
          .filter((key) => key !== CACHE_NAME)
          .map((key) => {
            console.log("[Service Worker] Suppression ancien cache:", key);
            return caches.delete(key);
          })
      );
    })
  );
  // Prend immédiatement le contrôle de toutes les pages
  return self.clients.claim();
});

// Stratégie de cache: Cache First, puis Network
self.addEventListener("fetch", (event) => {
  // 🔹 AJOUT IMPORTANT : gérer les navigations (lancement du PWA, changements d'URL)
  if (event.request.mode === "navigate") {
    event.respondWith(
      caches.match("./index.html").then((cachedPage) => {
        if (cachedPage) {
          return cachedPage;
        }
        // si pas en cache (premier lancement en ligne) → réseau
        return fetch(event.request);
      })
    );
    return;
  }

  // 🔹 Le reste de ton code inchangé
  event.respondWith(
    caches.match(event.request).then((cachedResponse) => {
      if (cachedResponse) {
        // Ressource trouvée dans le cache
        return cachedResponse;
      }

      // Pas en cache, on fetch depuis le réseau
      return fetch(event.request)
        .then((networkResponse) => {
          // Optionnel: mettre en cache les nouvelles ressources
          if (
            event.request.method === "GET" &&
            !event.request.url.includes("chrome-extension")
          ) {
            const responseToCache = networkResponse.clone();
            caches.open(CACHE_NAME).then((cache) => {
              cache.put(event.request, responseToCache);
            });
          }
          return networkResponse;
        })
        .catch((error) => {
          console.error("[Service Worker] Erreur fetch:", error);
          // Optionnel: retourner une page d'erreur offline
          return new Response(
            "Application hors ligne. Veuillez vérifier votre connexion.",
            {
              status: 503,
              statusText: "Service Unavailable",
              headers: new Headers({
                "Content-Type": "text/plain",
              }),
            }
          );
        });
    })
  );
});
