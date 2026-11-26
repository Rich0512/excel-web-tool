// ============================================
// Service Worker - 延平國小社團課表整理工具
// 提供離線快取功能，確保斷網後仍可正常使用
// ============================================

const CACHE_VERSION = 'v3.19';
const CACHE_NAME = `excel-web-tool-${CACHE_VERSION}`;

// 網路連得上但很慢時（學校 Wi-Fi 常見），等這麼久就先用快取的頁面
const NAVIGATION_TIMEOUT_MS = 4000;

// 需要快取的核心靜態資源清單
// - 頁面（index.html、guide.html、404.html）以 ?v=<版本> 引用的檔案，這裡也用完全相同的網址預快取，
//   頁面的請求第一次比對就會命中，不必再下載一次、存兩份。
//   ?v= 由 tools/bump-version.mjs 同步，tools/check-site.mjs 會檢查與各頁面一致。
// - docs/ 根目錄的頁面（404.html 除外）都要列在這裡，否則離線打不開；check-site.mjs 會檢查。
// - og-image.jpg 僅供社群分享爬蟲使用，不需離線快取
const STATIC_ASSETS = [
    './index.html',
    // 使用說明頁（沒有 JS）：離線時用這份預快取的版本回應，見下方 isGuideUrl()
    './guide.html',
    './css/style.css?v=3.19',
    './css/guide.css?v=3.19',
    './images/logo.png',
    './images/logo-192.png',
    './images/logo-512.png',
    './images/logo-maskable-192.png',
    './images/logo-maskable-512.png',
    './manifest.json',
    './fonts/fonts.css?v=3.19',
    // Noto Sans TC 依 unicode-range 拆成介面用字（ui）與其餘常用字（ext），見 fonts/fonts.css。
    // 只預快取 ui；ext（每個字重約 0.8 MB）在畫面第一次出現學生姓名等字時才下載，
    // 由下方 Cache-First 存進快取。從沒在線上顯示過名單就離線時，姓名改用系統字型。
    './fonts/NotoSansTC-Regular-ui.woff2',
    './fonts/NotoSansTC-Bold-ui.woff2',
    './fonts/Outfit-Regular.woff2',
    './fonts/Outfit-SemiBold.woff2',
    // ExcelJS 由 utils.js 的 loadExcelJS() 以 ?v=<版本> 載入；帶版本才不會在升級時拿到舊版 SW 快取的舊檔
    './js/libs/exceljs.min.js?v=3.19',
    './js/utils.js?v=3.19',
    './js/exporter.js?v=3.19',
    './js/parser.js?v=3.19',
    './js/processor.js?v=3.19',
    './js/mapping-ui.js?v=3.19',
    './js/app.js?v=3.19',
    './js/register-sw.js?v=3.19',
    './css/404.css?v=3.19',
    './templates/115學期名冊_範本.xlsx'
];

// 首頁（離線或網路太慢時的備用頁面）
const SHELL_URL = './index.html';
// 使用說明頁（離線時的備用頁面就是預快取的版本）
const GUIDE_URL = './guide.html';
// 備用頁面必須引用這一版的檔案（?v=），否則離線時會變成新 HTML 配舊 JS
// 版本號後面不能再接數字或小數點，否則 v3.1 會把引用 ?v=3.19 的 HTML 誤認為同一版
const ASSET_QUERY_RE = new RegExp(`\\?v=${CACHE_VERSION.slice(1).replace(/\./g, '\\.')}(?![\\d.])`);

// ==========================================
// 安裝事件：預快取所有靜態資源
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
// 啟用事件：清除舊版快取
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

// 去掉查詢參數與 # 的網址，用來比對導航的是哪一頁
function withoutSearch(url) {
    const target = new URL(url);
    target.search = '';
    target.hash = '';
    return target.href;
}

// 網址是否為首頁本身（忽略查詢參數與 #）：只有首頁會使用與更新備用頁面
function isShellUrl(url) {
    const href = withoutSearch(url);
    const scope = self.registration.scope;
    return href === scope || href === new URL(SHELL_URL, scope).href;
}

// 網址是否為使用說明頁（忽略查詢參數與 #）
function isGuideUrl(url) {
    return withoutSearch(url) === new URL(GUIDE_URL, self.registration.scope).href;
}

// 使用說明頁的導航：網路優先，網路失敗（離線）時回應預快取的說明頁。
// - 不把網路取得的說明頁寫回快取：預快取的版本是 install 時以 cache: 'reload' 取得、
//   與這一版 SW 的 ?v= 相符；寫回網站下一版的 HTML 會變成新 HTML 配這一版快取的 CSS。
// - 不像首頁那樣加逾時：說明頁從首頁以新分頁開啟，慢一點只是晚點顯示、不影響首頁的資料；
//   逾時改用快取還要處理「網路晚到的回應」，為了少用的頁面不值得多這段邏輯。
function respondToGuide(request) {
    return fetch(request).catch(() =>
        caches.match(GUIDE_URL).then((cached) => cached || Response.error())
    );
}

// 首頁的導航：網路先回應就用網路；超過 NAVIGATION_TIMEOUT_MS 且有快取時先用快取；
// 網路失敗（離線）時用快取。沒有快取時一律等網路的結果。
function respondToNavigation(networkPromise) {
    return new Promise((resolve) => {
        let settled = false;
        const settle = (response) => {
            if (settled || !response) return;
            settled = true;
            clearTimeout(timer);
            resolve(response);
        };
        const timer = setTimeout(() => {
            caches.match(SHELL_URL).then(settle);
        }, NAVIGATION_TIMEOUT_MS);
        networkPromise.then(settle, () => {
            caches.match(SHELL_URL).then((cached) => settle(cached || Response.error()));
        });
    });
}

// ==========================================
// 攔截請求策略：
// 1. HTML 導航頁面：
//    - 首頁：Network-First，加逾時（確保獲得最新版本，離線或太慢時用快取；同版回應存成備用頁面）
//    - 使用說明頁：Network-First，只在離線時用預快取的版本，不加逾時、不寫回快取
//    - 其他路徑：不攔截，交給瀏覽器取得 404 頁
// 2. 靜態資源 (JS/CSS/字型/圖片)：Cache-First (急速載入，節省頻寬)
// ==========================================
self.addEventListener('fetch', (event) => {
    const request = event.request;

    // 僅處理 GET 請求
    if (request.method !== 'GET') return;

    // 忽略 chrome-extension 等非 http(s) 請求
    if (!request.url.startsWith('http')) return;

    // 策略 1：HTML 導航請求採用 Network-First（首頁含逾時，說明頁只在離線時用快取）
    if (request.mode === 'navigate' || request.headers.get('accept')?.includes('text/html')) {
        if (isGuideUrl(request.url)) {
            event.respondWith(respondToGuide(request));
            return;
        }
        // 首頁與說明頁以外的路徑（例如打錯的網址）交給瀏覽器直接上網：回應首頁的話，
        // 頁面的相對路徑會以錯誤的資料夾為基準，整頁破版
        if (!isShellUrl(request.url)) return;

        const networkPromise = fetch(request);
        // 首頁的新版本存成備用頁面，但只存引用這一版 ?v= 的 HTML：
        // 網站已發佈下一版、這個 SW 還沒更新時，不能把新版 HTML 配上這一版快取的檔案
        event.waitUntil(
            networkPromise
                .then((networkResponse) => {
                    if (!networkResponse || networkResponse.status !== 200) return undefined;
                    const toCache = networkResponse.clone();
                    return networkResponse.clone().text().then((html) => {
                        if (!ASSET_QUERY_RE.test(html)) return undefined;
                        return caches.open(CACHE_NAME).then((cache) => cache.put(SHELL_URL, toCache));
                    });
                })
                .catch(() => {})
        );
        event.respondWith(respondToNavigation(networkPromise));
        return;
    }

    // 策略 2：靜態資源採用 Cache-First
    // 先以完整網址（含 ?v=）比對：新版 index.html 引用的 ?v= 在舊版 SW 的快取中一定找不到，
    // 會改從網路取得新檔，不會出現「新 HTML 配舊 JS」。
    // 只有在離線時才忽略 ?v= 改比對預快取的檔案。
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
                        // 離線：退回忽略 ?v= 比對預快取的檔案
                        return caches.match(request, { ignoreSearch: true })
                            .then((fallback) => fallback || new Response('', { status: 503, statusText: 'Offline' }));
                    });
            })
    );
});
