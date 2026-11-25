---
paths:
  - "docs/**/*.html"
  - "docs/css/**/*.css"
  - "docs/fonts/*.css"
  - "docs/js/*.js"
---

# 前端規則

使用者是學校行政人員，在學校電腦或手機上操作，目標是「把名冊變成課表」，不是欣賞介面。設計以清楚、穩定、可信任為優先。

## 程式結構

- 原生 JavaScript，沒有 bundler 也沒有模組系統。`docs/js/*.js` 以 `defer` 依序載入並共用全域函式：utils → exporter → parser → processor → mapping-ui → app → register-sw。新增檔案要放進這個順序，並同步加入 `service-worker.js` 的 `STATIC_ASSETS` 與 `index.html` 的 `?v=`。
- `docs/js/libs/` 是第三方函式庫（ExcelJS 4.4.0，npm 官方 tarball 的 `dist/exceljs.min.js`），不要修改內容；升級時整檔替換，並確認 npm 公布的 integrity 雜湊一致。
- ExcelJS 不用 `<script>` 標籤載入，一律透過 `utils.js` 的 `loadExcelJS()`（頁面閒置時預先載入，使用前 `await`）。不要在 `index.html` 加回標籤：它約 0.9 MB，排在前面會讓按鈕要等它下載完才有反應。

## CSP 相容（`tools/check-site.mjs` 會檢查）

CSP 沒有開放 `'unsafe-inline'`，也沒有 `'unsafe-eval'`（不可用 `eval`、`new Function`、字串形式的 `setTimeout`）。ExcelJS 內附的 regenerator-runtime 會在全域變數未宣告時退回 `Function()`，`loadExcelJS()` 載入前會先宣告 `regeneratorRuntime`，不要移除。另外：
- HTML 不可有 `style="…"`、`onclick="…"` 等屬性、inline `<script>`、`<style>` 區塊。
- JS 產生的 HTML 字串同樣不可含 `style=` 或 `on*=`。改用 class；事件用 `addEventListener`。
- 用 JS 設定 `element.style.display` 或 `element.hidden` 是允許的（CSSOM 不受 CSP 限制），適合單純的顯示／隱藏。

## 防 XSS

- 名冊內容、Excel 工作表與欄位名稱、使用者輸入的社團名稱都是不可信任的資料。
- 只是顯示文字時用 `textContent`；需要組 HTML 時，每個插入值都要先經過 `utils.js` 的 `escapeHtml()`。

## 設計語言

完整規範在 repo 根目錄的 `DESIGN.md`（課表格線，2026-11-15 定案），改外觀前先讀。重點：

- 顏色只用 `style.css` `:root` 的 token：`--primary` #0f3b6d 校徽藍（主要按鈕、標題、目前步驟）、`--ink`、`--ink-muted`、`--line`（格線）、`--control`（輸入框邊框）、`--on-primary`（校徽藍底上的白字）、`--mark` 螢光筆黃（只用在課表範例）。不要直接寫色碼。語意狀態色見 `DESIGN.md`。不要新增裝飾色、漸層、毛玻璃或大陰影。
- 字型只有 Outfit（400/600）與 Noto Sans TC（400/700，已裁切）。不要新增字型或字重，也不要引用 CDN。
- **不使用 emoji**（HTML、JS 產生的訊息、`alert()` 都是）；圖示用 inline SVG 線條圖示。
- 內容靠左對齊、最寬 1180px；沿用 `DESIGN.md` 的元件表，不要另造相似元件。
- 動態只回應使用者動作；`prefers-reduced-motion` 已全域處理，新動畫不可繞過。

## 無障礙與行動裝置

- 每個表單元素都要有 `<label for>` 或 `aria-label`。
- 可點擊的非按鈕元素要加 `role="button"`、`tabindex="0"`，並處理 Enter／空白鍵。
- `role="button"` 或 `<button>` 裡面不可再放連結、按鈕等可互動元素（鍵盤與螢幕閱讀器無法分別操作）。需要在可點擊區塊裡放連結時，區塊本身不當按鈕，改在裡面放真正的 `<button>`（例如拖曳區）。
- 鍵盤焦點用 `:focus-visible` 顯示，不可只寫 `outline: none`。
- 斷點：768px（平板）、640px（手機）。觸控目標在所有寬度都至少 44px（桌機也可能是觸控螢幕）；輸入框字級至少 16px，避免 iOS 自動縮放。

## 文字

- 繁體中文，口吻直接、說明清楚。按鈕寫出會發生的事（「開始轉換並下載課表」），而不是「送出」。
- 錯誤訊息要說明發生什麼事、使用者該怎麼做，不要只寫「發生錯誤」。
- 新增罕見中文字後執行 `node tools/check-site.mjs`；回報缺字時執行 `/subset-fonts`。
