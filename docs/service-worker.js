// ============================================
// Service Worker - 延平國小社團課表整理工具
// 提供離線快取功能，確保斷網後仍可正常使用
// ============================================

const CACHE_VERSION = 'v3.9';
const CACHE_NAME = `excel-web-tool-${CACHE_VERSION}`;

// 需要快取的核心靜態資源清單
// (og-image.jpg 僅供社群分享爬蟲使用，不需離線快取)
const STATIC_ASSETS = [
    './',
    './index.html',
    './css/style.css',
    './images/logo.png',
    './images/logo-192.png',
    './images/logo-512.png',
    './images/logo-maskable-192.png',
    './images/logo-maskable-512.png',
    './manifest.json',
    './fonts/fonts.css',
    './fonts/NotoSansTC-Regular.woff2',
    './fonts/NotoSansTC-Bold.woff2',
    './fonts/Outfit-Regular.woff2',
    './fonts/Outfit-SemiBold.woff2',
    './js/libs/exceljs.min.js',
    './js/utils.js',
    './js/exporter.js',
    './js/parser.js',
    './js/processor.js',
    './js/mapping-ui.js',
    './js/app.js',
    './js/register-sw.js',
    './css/404.css',
    './templates/115學期名冊_範本.xlsx'
];

// ==========================================
// 📦 安裝事件：預快取所有靜態資源
// ==========================================
self.addEventListener('install', (event) => {
    event.waitUntil(
        caches.open(CACHE_NAME)
            .then((cache) => {
                console.log('[SW] 正在預快取靜態資源...');
                // cache: 'reload' 略過瀏覽器的 HTTP 快取（GitHub Pages 預設快取 10 分鐘），
                // 避免把上一版的檔案存進新版本的快取
                return cache.addAll(STATIC_ASSETS.map((url) => new Request(url, { cache: 'reload' })));
            })
            .then(() => {
                console.log('[SW] 所有靜態資源已快取完成');
                return self.skipWaiting(); // 立即啟用新版 SW
            })
    );
});

// ==========================================
// 🧹 啟用事件：清除舊版快取
// ==========================================
self.addEventListener('activate', (event) => {
    event.waitUntil(
        caches.keys()
            .then((cacheNames) => {
                return Promise.all(
                    cacheNames
                        .filter((name) => name !== CACHE_NAME)
                        .map((name) => {
                            console.log(`[SW] 清除舊版快取: ${name}`);
                            return caches.delete(name);
                        })
                );
            })
            .then(() => {
                console.log('[SW] 新版 Service Worker 已啟用');
                return self.clients.claim(); // 立即控制所有頁面
            })
    );
});

// ==========================================
// 🌐 攔截請求策略：
// 1. HTML 導航頁面：Network-First (確保獲得最新版本，離線時 fallback 快取)
// 2. 靜態資源 (JS/CSS/字型/圖片)：Cache-First (急速載入，節省頻寬)
// ==========================================
self.addEventListener('fetch', (event) => {
    const request = event.request;

    // 僅處理 GET 請求
    if (request.method !== 'GET') return;

    // 忽略 chrome-extension 等非 http(s) 請求
    if (!request.url.startsWith('http')) return;

    // 策略 1：HTML 導航請求採用 Network-First
    if (request.mode === 'navigate' || request.headers.get('accept')?.includes('text/html')) {
        event.respondWith(
            fetch(request)
                .then((networkResponse) => {
                    if (networkResponse && networkResponse.status === 200) {
                        const responseClone = networkResponse.clone();
                        caches.open(CACHE_NAME).then((cache) => {
                            cache.put(request, responseClone);
                        });
                    }
                    return networkResponse;
                })
                .catch(() => {
                    // 斷網或離線時，從快取提供頁面
                    return caches.match('./index.html');
                })
        );
        return;
    }

    // 策略 2：靜態資源採用 Cache-First
    // 先以完整網址（含 ?v=）比對：新版 index.html 引用的 ?v= 在舊版 SW 的快取中一定找不到，
    // 會改從網路取得新檔，不會出現「新 HTML 配舊 JS」。
    // 只有在離線時才忽略 ?v= 改比對預快取的檔案（預快取網址不含查詢參數）。
    event.respondWith(
        caches.match(request)
            .then((cachedResponse) => {
                if (cachedResponse) {
                    return cachedResponse; // 快取命中，直接回傳
                }

                // 快取未命中，嘗試從網路取得
                return fetch(request)
                    .then((networkResponse) => {
                        if (!networkResponse || networkResponse.status !== 200) {
                            return networkResponse;
                        }

                        const responseClone = networkResponse.clone();
                        caches.open(CACHE_NAME).then((cache) => {
                            cache.put(request, responseClone);
                        });

                        return networkResponse;
                    })
                    .catch(() => {
                        // 離線：退回比對不含 ?v= 的預快取檔案
                        return caches.match(request, { ignoreSearch: true })
                            .then((fallback) => fallback || new Response('', { status: 503, statusText: 'Offline' }));
                    });
            })
    );
});
