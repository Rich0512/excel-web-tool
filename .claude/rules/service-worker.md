---
paths:
  - "docs/service-worker.js"
  - "docs/index.html"
  - "docs/404.html"
  - "docs/manifest.json"
  - "docs/sitemap.xml"
---

# 快取與版本規則

使用者多半是安裝成 PWA 或重複開啟同一個分頁。Service Worker 一旦快取舊檔案，就只能靠版本號讓瀏覽器更新，所以版本管理錯了，使用者會一直拿到舊版而且很難察覺。

- **版本號三處要一致**：`service-worker.js` 的 `CACHE_VERSION`、`index.html` 所有 `?v=`、`sitemap.xml` 的 `lastmod`。不要手動改，用 `node tools/bump-version.mjs`（或 `/release`）。
- **一次發佈只調升一次**：工具會比對 `origin/main`，已調升過就只同步日期。
- **`STATIC_ASSETS`**：只放執行時需要的檔案。新增、改名、刪除 `docs/` 內的 JS/CSS/字型/圖片時要同步更新。清單內任何一個檔案不存在，`cache.addAll` 就會失敗，整個 Service Worker 無法安裝。不要放 `og-image.jpg`（只給社群爬蟲）或 `.ttf`。
- **快取策略**：HTML 走 Network-First（離線時退回快取）。其他靜態資源走 Cache-First，先用含 `?v=` 的完整網址比對，查不到就上網取得；只有離線時才用 `ignoreSearch: true` 退回預快取。不要把 `ignoreSearch` 放在第一次比對：舊版 Service Worker 會拿舊檔回應新版的 `?v=`，造成新 HTML 配舊 JS。
- **預快取**用 `new Request(url, { cache: 'reload' })`，略過 GitHub Pages 的 10 分鐘 HTTP 快取，避免把上一版的檔案存進新版快取。
- **路徑**：網站在子路徑 `/excel-web-tool/`。`index.html`、`manifest.json`、`service-worker.js` 一律用相對路徑；只有 `404.html` 用 `/excel-web-tool/…` 絕對路徑，因為 GitHub Pages 會在任意深度的網址回傳它。
