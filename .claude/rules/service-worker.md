---
paths:
  - "docs/service-worker.js"
  - "docs/index.html"
  - "docs/404.html"
  - "docs/guide.html"
  - "docs/manifest.json"
  - "docs/sitemap.xml"
---

# 快取與版本規則

使用者多半是安裝成 PWA 或重複開啟同一個分頁。Service Worker 一旦快取舊檔案，就只能靠版本號讓瀏覽器更新，所以版本管理錯了，使用者會一直拿到舊版而且很難察覺。

- **版本號要一致**：`service-worker.js` 的 `CACHE_VERSION` 與 `STATIC_ASSETS` 的 `?v=`、`index.html`、`guide.html`、`404.html` 所有 `?v=`、`sitemap.xml` 每個 `lastmod`。不要手動改，用 `node tools/bump-version.mjs`（或 `/release`）。
- **一次發佈只調升一次**：工具會比對 `origin/main`，已調升過就只同步日期。
- **`STATIC_ASSETS`**：只放執行時需要的檔案。新增、改名、刪除 `docs/` 內的 JS/CSS/字型/圖片時要同步更新。清單內任何一個檔案不存在，`cache.addAll` 就會失敗，整個 Service Worker 無法安裝。不要放 `og-image.jpg`（只給社群爬蟲）、`.ttf`，或 `NotoSansTC-*-ext.woff2`（常用字檔，畫面用到時才下載並由 Cache-First 存起來）。
- **`index.html` 以 `?v=` 引用的檔案，`STATIC_ASSETS` 要寫完全相同的網址**（例如 `./js/app.js?v=<版本>`；`404.html` 的 `/excel-web-tool/css/404.css?v=<版本>` 也一樣，ExcelJS 則由 `utils.js` 以 `js/libs/exceljs.min.js?v=<版本>` 動態載入），頁面的請求才會直接命中預快取；不帶 `?v=` 會比對不到，又下載並存一份。同一個檔案不要帶與不帶 `?v=` 各列一次。`check-site.mjs` 會檢查。
- **快取策略**：HTML 走 Network-First，`NAVIGATION_TIMEOUT_MS`（4 秒）內沒有回應且有快取時先用快取，離線時退回快取。只有首頁會這樣處理，而且只有引用這一版 `?v=` 的首頁回應才會存成備用頁面（`./index.html`），避免新 HTML 配舊檔。使用說明頁（`./guide.html`）網路優先、只在離線時回應預快取的版本，網路回應不寫回快取（預快取版本才保證與這一版的 `?v=` 相符）。首頁與說明頁以外的導航（例如打錯的網址）不攔截，交給瀏覽器取得 404 頁；若回應首頁，相對路徑會以錯誤的資料夾為基準而整頁破版。其他靜態資源走 Cache-First，先用含 `?v=` 的完整網址比對，查不到就上網取得；只有離線時才用 `ignoreSearch: true` 退回預快取。不要把 `ignoreSearch` 放在第一次比對：舊版 Service Worker 會拿舊檔回應新版的 `?v=`，造成新 HTML 配舊 JS。
- **預快取**用 `new Request(url, { cache: 'reload' })`，略過 GitHub Pages 的 10 分鐘 HTTP 快取，避免把上一版的檔案存進新版快取。
- **路徑**：網站在子路徑 `/excel-web-tool/`。`index.html`、`manifest.json`、`service-worker.js` 一律用相對路徑；只有 `404.html` 用 `/excel-web-tool/…` 絕對路徑，因為 GitHub Pages 會在任意深度的網址回傳它。
