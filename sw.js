// ============================================================
// sw.js — DACUM Live Pro Service Worker  v13
// Path-agnostic: BASE is derived dynamically from scope.
// Works regardless of repository name (V3.0, V3.1, etc.)
// ============================================================

const CACHE_VERSION = 'v33';
const CACHE_NAME    = 'dacum-live-pro-' + CACHE_VERSION;
// Derive BASE from the SW scope so this file works in any repo path
const BASE          = self.registration ? self.registration.scope : '/';
const OFFLINE_URL   = BASE + 'index.html';

// These resource types are served network-first (fresh code always wins).
// Images/icons are still cache-first since they rarely change.
const NETWORK_FIRST_EXT = /\.(html|js|css|json)(\?.*)?$/i;

// Keep this list in sync with:
//   • every <script src> and <link rel="stylesheet"> in index.html
//   • every ES module reachable from app.js's import graph
// A module missing here still works online (network-first), but the
// app breaks offline at the point it is first imported.
const PRECACHE_URLS = [
  BASE + 'index.html',

  // ── Stylesheets ──────────────────────────────────────────
  BASE + 'dacum-styles.css',
  BASE + 'dacum-responsive.css',
  BASE + 'dacum-typography.css',
  BASE + 'dacum-components.css',

  // ── ES modules (app.js import graph) ─────────────────────
  BASE + 'app.js',
  BASE + 'state.js',
  BASE + 'renderer.js',
  BASE + 'duties.js',
  BASE + 'events.js',
  BASE + 'history.js',
  BASE + 'storage.js',
  BASE + 'tabs.js',
  BASE + 'tasks.js',
  BASE + 'codes.js',
  BASE + 'modules.js',
  BASE + 'projects.js',
  BASE + 'exports.js',
  BASE + 'snapshots.js',
  BASE + 'workshop.js',
  BASE + 'workshop_snapshots.js',
  BASE + 'dacum_projects.js',
  BASE + 'refine.js',
  BASE + 'drag_drop.js',
  BASE + 'additional_info_ai.js',
  BASE + 'module_mapping_ai.js',
  BASE + 'clustering_ai.js',
  BASE + 'learning_outcomes_ai.js',
  BASE + 'image_store.js',
  BASE + 'autosave.js',
  BASE + 'verification_charts.js',
  BASE + 'error-handler.js',

  // ── Classic scripts loaded directly by index.html ────────
  BASE + 'dacum-ui.js',
  BASE + 'dacum-mobile.js',
  BASE + 'dacum-fixes.js',
  BASE + 'qrcode.min.js',

  // ── PWA assets ───────────────────────────────────────────
  BASE + 'manifest.json',
  BASE + 'icon-192.png',
  BASE + 'icon-512.png',
];

// ── Install: precache all assets, activate immediately ────────
self.addEventListener('install', function (event) {
  console.log('[SW ' + CACHE_VERSION + '] Installing...');
  self.skipWaiting();
  event.waitUntil(
    caches.open(CACHE_NAME).then(function (cache) {
      return Promise.allSettled(
        PRECACHE_URLS.map(function (url) {
          // Fetch with no-cache to bypass any HTTP cache — gets fresh files
          return fetch(url, { cache: 'no-store' })
            .then(function (res) {
              if (res && res.status === 200) return cache.put(url, res);
            })
            .catch(function (err) {
              console.warn('[SW] Precache skipped:', url, err.message);
            });
        })
      );
    }).then(function () {
      console.log('[SW ' + CACHE_VERSION + '] Precache complete');
    })
  );
});

// ── Activate: claim clients, purge old caches, warm cache, notify to reload
self.addEventListener('activate', function (event) {
  console.log('[SW ' + CACHE_VERSION + '] Activating...');
  event.waitUntil(
    self.clients.claim()
      .then(function () {
        // Delete ALL old version caches
        return caches.keys().then(function (keys) {
          return Promise.all(
            keys.filter(function (k) { return k !== CACHE_NAME; })
                .map(function (k) {
                  console.log('[SW] Deleting old cache:', k);
                  return caches.delete(k);
                })
          );
        });
      })
      .then(function () {
        // Pre-fetch the critical files fresh from network RIGHT NOW,
        // before notifying clients to reload. This ensures that when
        // old-HTML pages reload, they immediately get the new HTML/JS/CSS
        // instead of waiting for the next fetch cycle.
        var criticalUrls = [
          BASE + 'index.html',
          BASE + 'dacum_projects.js',
          BASE + 'dacum-mobile.js',
          BASE + 'dacum-styles.css',
          BASE + 'dacum-responsive.css',
          // Added in Phase 1: index.html no longer carries its own
          // <style> block, so it is useless without this file. It has to
          // be warmed alongside the HTML or an updated page would paint
          // unstyled until the next fetch cycle.
          BASE + 'dacum-components.css',
          BASE + 'app.js',
        ];
        return caches.open(CACHE_NAME).then(function (cache) {
          return Promise.allSettled(
            criticalUrls.map(function (url) {
              return fetch(url, { cache: 'no-store' }).then(function (res) {
                if (res && res.status === 200) {
                  cache.put(url, res.clone());
                  console.log('[SW] Warmed cache:', url);
                }
                return res;
              }).catch(function (err) {
                console.warn('[SW] Warm failed:', url, err.message);
              });
            })
          );
        });
      })
      .then(function () {
        // NOW notify all clients to reload — fresh files are ready
        console.log('[SW ' + CACHE_VERSION + '] Cache warm complete — notifying clients');
        return self.clients.matchAll({
          type: 'window',
          includeUncontrolled: true
        }).then(function (clients) {
          clients.forEach(function (client) {
            client.postMessage({ type: 'SW_UPDATED', version: CACHE_VERSION });
          });
        });
      })
  );
});

// ── Message: handle SKIP_WAITING + GET_VERSION from page ────────
self.addEventListener('message', function (event) {
  if (!event.data) return;
  if (event.data.type === 'SKIP_WAITING') {
    console.log('[SW] SKIP_WAITING received');
    self.skipWaiting();
  }
  // Version query — page checks if controller is up to date
  if (event.data.type === 'GET_VERSION') {
    var port = event.ports && event.ports[0];
    var reply = { type: 'VERSION_REPLY', version: CACHE_VERSION };
    if (port) {
      port.postMessage(reply);
    } else if (event.source) {
      event.source.postMessage(reply);
    }
  }
});

// ── Fetch: route each request to the right strategy ──────────
self.addEventListener('fetch', function (event) {
  var req = event.request;
  if (req.method !== 'GET') return;

  var url = new URL(req.url);

  // External API → network only
  if (url.hostname.includes('railway.app') || url.pathname.includes('/api/')) {
    event.respondWith(fetch(req));
    return;
  }

  // Cross-origin CDN → network-first with cache fallback
  if (url.hostname !== self.location.hostname) {
    event.respondWith(networkFirst(req));
    return;
  }

  // Same-origin HTML, JS, CSS, JSON → network-first
  // This ensures new sidebar code always loads, even after install
  if (NETWORK_FIRST_EXT.test(url.pathname)) {
    event.respondWith(networkFirst(req));
    return;
  }

  // Images and other static assets → cache-first
  event.respondWith(cacheFirst(req));
});

// ── Network-first: try network, cache on success, fallback to cache
async function networkFirst(request) {
  var canonical = request.url.split('?')[0]; // strip query for consistent cache key
  try {
    var res = await fetch(canonical, { cache: 'no-cache' });
    if (res && res.status === 200 && res.type !== 'opaque') {
      var cache = await caches.open(CACHE_NAME);
      cache.put(canonical, res.clone());
    }
    return res;
  } catch (_) {
    // Offline: serve from cache (any query variant accepted)
    var cached = await caches.match(canonical) ||
                 await caches.match(request, { ignoreSearch: true });
    if (cached) return cached;
    if (request.mode === 'navigate') {
      return (await caches.match(OFFLINE_URL)) ||
             new Response(
               '<!doctype html><html><body style="font-family:sans-serif;text-align:center;padding:60px">' +
               '<h2>\uD83D\uDCF6 You are offline</h2>' +
               '<p>DACUM Live Pro will reload when reconnected.</p>' +
               '</body></html>',
               { status: 503, headers: { 'Content-Type': 'text/html' } }
             );
    }
    return new Response('Offline', { status: 503 });
  }
}

// ── Cache-first: serve cache, refresh in background
async function cacheFirst(request) {
  var cached = await caches.match(request);
  if (cached) {
    // Background refresh
    fetch(request).then(function (res) {
      if (res && res.status === 200 && res.type !== 'opaque') {
        caches.open(CACHE_NAME).then(function (c) { c.put(request, res); });
      }
    }).catch(function () {});
    return cached;
  }
  try {
    var res = await fetch(request);
    if (res && res.status === 200 && res.type !== 'opaque') {
      var cache = await caches.open(CACHE_NAME);
      cache.put(request, res.clone());
    }
    return res;
  } catch (_) {
    return new Response('Offline', { status: 503 });
  }
}
