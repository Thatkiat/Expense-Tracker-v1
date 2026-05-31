/**
 * ═══════════════════════════════════════════════
 *  Field Expense Tracker — Service Worker
 *  Strategies:
 *    App Shell      → Cache First (offline-first)
 *    Google Fonts   → Stale-While-Revalidate
 *    Apps Script    → Network First (never cache)
 *    Drive Images   → Network First + cache fallback
 * ═══════════════════════════════════════════════
 */

const APP_CACHE   = 'field-expense-v1';
const FONT_CACHE  = 'field-expense-fonts-v1';
const IMG_CACHE   = 'field-expense-images-v1';
const CACHE_VER   = 1; // bump this number to force refresh on all clients

// Files to pre-cache on install (app shell)
const PRECACHE = [
  './',
  './index.html',
  './manifest.json',
  './icons/icon.svg',
  './icons/icon-maskable.svg',
];

// ──────────────────────────────
//  INSTALL — pre-cache app shell
// ──────────────────────────────
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(APP_CACHE)
      .then(cache => {
        // addAll with individual error handling (icons might not exist yet)
        return Promise.allSettled(
          PRECACHE.map(url =>
            cache.add(new Request(url, { cache: 'reload' })).catch(() => null)
          )
        );
      })
      .then(() => self.skipWaiting())
  );
});

// ──────────────────────────────
//  ACTIVATE — clean old caches
// ──────────────────────────────
self.addEventListener('activate', event => {
  const VALID = [APP_CACHE, FONT_CACHE, IMG_CACHE];
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(
        keys.filter(k => !VALID.includes(k)).map(k => caches.delete(k))
      ))
      .then(() => self.clients.claim())
  );
});

// ──────────────────────────────
//  FETCH — routing by URL type
// ──────────────────────────────
self.addEventListener('fetch', event => {
  // Ignore non-GET, chrome-extension, POST to Apps Script
  if (event.request.method !== 'GET') return;
  if (!event.request.url.startsWith('http')) return;

  const url = new URL(event.request.url);

  // 1. Google Apps Script — network only (never cache, contains sensitive tokens)
  if (url.hostname.includes('script.google.com')) {
    event.respondWith(networkOnly(event.request));
    return;
  }

  // 2. Google Fonts — stale-while-revalidate (perf + offline)
  if (
    url.hostname === 'fonts.googleapis.com' ||
    url.hostname === 'fonts.gstatic.com'
  ) {
    event.respondWith(staleWhileRevalidate(event.request, FONT_CACHE));
    return;
  }

  // 3. Google Drive images (receipt/slip photos) — network first, cache fallback
  if (
    url.hostname.includes('drive.google.com') ||
    url.hostname.includes('googleusercontent.com') ||
    url.hostname.includes('lh3.googleusercontent.com')
  ) {
    event.respondWith(networkFirst(event.request, IMG_CACHE));
    return;
  }

  // 4. Same-origin app shell (HTML, JSON, SVG, PNG) — cache first
  if (url.origin === self.location.origin) {
    event.respondWith(cacheFirst(event.request, APP_CACHE));
    return;
  }

  // 5. Everything else — network first
  event.respondWith(networkFirst(event.request, APP_CACHE));
});

// ──────────────────────────────
//  STRATEGIES
// ──────────────────────────────

/** Cache First: serve from cache, fetch & update in background */
async function cacheFirst(request, cacheName) {
  const cached = await caches.match(request);
  if (cached) return cached;

  try {
    const response = await fetch(request);
    if (response.ok) {
      const cache = await caches.open(cacheName);
      cache.put(request, response.clone());
    }
    return response;
  } catch {
    // Navigation fallback → serve app shell
    if (request.mode === 'navigate') {
      const fallback = await caches.match('./') || await caches.match('./index.html');
      if (fallback) return fallback;
    }
    return offlinePage();
  }
}

/** Network First: try network, fall back to cache */
async function networkFirst(request, cacheName) {
  try {
    const response = await fetch(request);
    if (response.ok && cacheName) {
      const cache = await caches.open(cacheName);
      cache.put(request, response.clone());
    }
    return response;
  } catch {
    const cached = await caches.match(request);
    if (cached) return cached;
    if (request.mode === 'navigate') {
      const fallback = await caches.match('./') || await caches.match('./index.html');
      if (fallback) return fallback;
    }
    // Return offline JSON for API-like requests
    return new Response(
      JSON.stringify({ success: false, error: 'offline', offline: true }),
      { headers: { 'Content-Type': 'application/json' } }
    );
  }
}

/** Stale-While-Revalidate: serve cache instantly, refresh in background */
async function staleWhileRevalidate(request, cacheName) {
  const cache  = await caches.open(cacheName);
  const cached = await cache.match(request);

  const fetchPromise = fetch(request)
    .then(response => {
      if (response.ok) cache.put(request, response.clone());
      return response;
    })
    .catch(() => null);

  return cached || await fetchPromise || offlinePage();
}

/** Network Only — no caching at all */
async function networkOnly(request) {
  try {
    return await fetch(request);
  } catch {
    return new Response(
      JSON.stringify({ success: false, error: 'offline', offline: true }),
      { headers: { 'Content-Type': 'application/json' } }
    );
  }
}

function offlinePage() {
  return new Response(
    `<!DOCTYPE html><html lang="th"><head><meta charset="UTF-8">
     <meta name="viewport" content="width=device-width,initial-scale=1">
     <title>Offline</title>
     <style>body{font-family:sans-serif;display:flex;flex-direction:column;align-items:center;
     justify-content:center;min-height:100vh;background:#F1F5F9;color:#1E293B;gap:12px;padding:20px;text-align:center}
     .ico{font-size:56px}.ttl{font-size:20px;font-weight:700}.sub{color:#64748B;font-size:14px}
     button{padding:12px 28px;background:#2563EB;color:#fff;border:none;border-radius:10px;
     font-size:15px;cursor:pointer;margin-top:8px}</style></head>
     <body><div class="ico">📶</div><div class="ttl">ไม่มีอินเทอร์เน็ต</div>
     <div class="sub">ข้อมูลที่บันทึกไว้จะยังคงอยู่<br>และจะซิงค์อัตโนมัติเมื่อกลับมาออนไลน์</div>
     <button onclick="location.reload()">ลองใหม่อีกครั้ง</button></body></html>`,
    { headers: { 'Content-Type': 'text/html; charset=utf-8' } }
  );
}

// ──────────────────────────────
//  BACKGROUND SYNC
// ──────────────────────────────
self.addEventListener('sync', event => {
  if (event.tag === 'sync-expenses') {
    event.waitUntil(triggerClientSync());
  }
});

async function triggerClientSync() {
  const clients = await self.clients.matchAll({ type: 'window', includeUncontrolled: true });
  clients.forEach(client => client.postMessage({ type: 'TRIGGER_SYNC' }));
}

// ──────────────────────────────
//  MESSAGE HANDLER
// ──────────────────────────────
self.addEventListener('message', event => {
  // Force update request from app
  if (event.data?.type === 'SKIP_WAITING') {
    self.skipWaiting();
  }
  // Cache a specific URL on demand
  if (event.data?.type === 'CACHE_URL') {
    const url = event.data.url;
    caches.open(IMG_CACHE).then(c => c.add(url)).catch(() => {});
  }
});
