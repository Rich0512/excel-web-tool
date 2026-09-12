// ============================================
// Service Worker - 延平國小社團課表整理工具
// 提供離線快取功能，確保斷網後仍可正常使用
// ============================================

const CACHE_VERSION = 'v3.1';
const CACHE_NAME = `excel-web-tool-${CACHE_VERSION}`;

// 需要快取的核心靜態資源清單
const STATIC_ASSETS = [
    './',
    './index.html',
    './style.css',
    './logo.png',
    './manifest.json',
    './fonts/fonts.css',
    './fonts/NotoSansTC-Light.ttf',
    './fonts/NotoSansTC-Regular.ttf',
    './fonts/NotoSansTC-Bold.ttf',
    './fonts/Outfit-Light.ttf',
    './fonts/Outfit-Regular.ttf',
    './fonts/Outfit-SemiBold.ttf',
    './fonts/Outfit-ExtraBold.ttf',
    './js/libs/exceljs.min.js',
    './js/utils.js',
    './js/exporter.js',
    './js/parser.js',
    './js/processor.js',
    './js/mapping-ui.js',
    './js/app.js',
    './115學期名冊_範本.xlsx'
];

// ==========================================
// 📦 安裝事件：預快取所有靜態資源
// ==========================================
self.addEventListener('install', (event) => {
    event.waitUntil(
        caches.open(CACHE_NAME)
            .then((cache) => {
                console.log('[SW] 正在預快取靜態資源...');
                return cache.addAll(STATIC_ASSETS);
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
// 🌐 攔截請求：Cache-First 策略
// ==========================================
self.addEventListener('fetch', (event) => {
    const request = event.request;

    // 僅處理 GET 請求
    if (request.method !== 'GET') return;

    // 忽略 chrome-extension 等非 http(s) 請求
    if (!request.url.startsWith('http')) return;

    event.respondWith(
        caches.match(request)
            .then((cachedResponse) => {
                if (cachedResponse) {
                    return cachedResponse; // 快取命中，直接回傳
                }

                // 快取未命中，嘗試從網路取得
                return fetch(request)
                    .then((networkResponse) => {
                        // 不快取非成功回應或不透明回應
                        if (!networkResponse || networkResponse.status !== 200) {
                            return networkResponse;
                        }

                        // 動態快取成功的網路回應（例如未來新增的資源）
                        const responseClone = networkResponse.clone();
                        caches.open(CACHE_NAME).then((cache) => {
                            cache.put(request, responseClone);
                        });

                        return networkResponse;
                    })
                    .catch(() => {
                        // 網路也失敗時，如果是導航請求，回傳離線首頁
                        if (request.mode === 'navigate') {
                            return caches.match('./index.html');
                        }
                        // 其他資源無法取得時不做特殊處理
                        return new Response('', { status: 503, statusText: 'Offline' });
                    });
            })
    );
});
