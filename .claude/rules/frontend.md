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
- `docs/js/libs/` 是第三方函式庫（ExcelJS），不要修改。

## CSP 相容（`tools/check-site.mjs` 會檢查）

CSP 沒有開放 `'unsafe-inline'`，所以：
- HTML 不可有 `style="…"`、`onclick="…"` 等屬性、inline `<script>`、`<style>` 區塊。
- JS 產生的 HTML 字串同樣不可含 `style=` 或 `on*=`。改用 class；事件用 `addEventListener`。
- 用 JS 設定 `element.style.display` 或 `element.hidden` 是允許的（CSSOM 不受 CSP 限制），適合單純的顯示／隱藏。

## 防 XSS

- 名冊內容、Excel 工作表與欄位名稱、使用者輸入的社團名稱都是不可信任的資料。
- 只是顯示文字時用 `textContent`；需要組 HTML 時，每個插入值都要先經過 `utils.js` 的 `escapeHtml()`。

## 設計語言（沿用，不另起爐灶）

- 顏色使用 `style.css` 的 `:root` 變數：`--primary` #0f3b6d（校徽藍，標題、焦點）、`--accent` #2ecc71（主要動作按鈕）、`--text-main`、`--text-muted`、`--border`。新顏色只用於語意狀態（警告 #d97706 系、錯誤 #c53030 系），不要新增裝飾色。
- 字型只有 Outfit（400/600/800）與 Noto Sans TC（400/700，已裁切）。不要新增字型或字重，也不要引用 CDN。
- 版面是置中的單張卡片（最寬 1050px）。沿用既有元件：`.btn` + `.btn-primary`／`.btn-secondary`／`.btn-outline`、`.tab-intro`（情境說明）、`.warning-callout`（提醒）、`.config-select`、`.club-input`。
- 動態效果只用在回應使用者動作（切換步驟、展開名單）。不要加入裝飾性的進場動畫；`prefers-reduced-motion` 已全域處理，新動畫不可繞過。
- 新增介面時避免常見的「AI 生成感」：不要每段都加全大寫小標、不要只為裝飾而用卡片陰影與漸層、不要在按鈕文字後面加箭頭。

## 無障礙與行動裝置

- 每個表單元素都要有 `<label for>` 或 `aria-label`。
- 可點擊的非按鈕元素要加 `role="button"`、`tabindex="0"`，並處理 Enter／空白鍵。
- 鍵盤焦點用 `:focus-visible` 顯示，不可只寫 `outline: none`。
- 斷點：768px（平板）、640px（手機）。手機上的觸控目標至少 44px；輸入框字級至少 16px，避免 iOS 自動縮放。

## 文字

- 繁體中文，口吻直接、說明清楚。按鈕寫出會發生的事（「開始轉換並下載課表」），而不是「送出」。
- 錯誤訊息要說明發生什麼事、使用者該怎麼做，不要只寫「發生錯誤」。
- 新增罕見中文字後執行 `node tools/check-site.mjs`；回報缺字時執行 `/subset-fonts`。
