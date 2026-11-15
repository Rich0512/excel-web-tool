---
name: preview
description: 在本機啟動與 GitHub Pages 行為相同的預覽伺服器（子路徑 /excel-web-tool/、自訂 404），並實際驗證頁面。修改 docs/ 底下的 HTML、CSS、JS 後、push 上線前，或使用者說「預覽」「本機看看」「測試網頁」「跑起來看」時使用。
argument-hint: "[port，預設 8000]"
allowed-tools: Bash(node tools/serve.mjs *) Bash(curl -s *)
---

# 本機預覽與驗證

部署檢查只能看靜態檔案，看不到「頁面實際能不能用」。在本機用相同的子路徑跑一次，才抓得到 CSP 擋掉的腳本、路徑錯誤和 404 頁面的問題。

## 步驟

1. **啟動伺服器**（背景執行）：`node tools/serve.mjs $ARGUMENTS`，網址是 `http://localhost:8000/excel-web-tool/`。
2. **基本驗證**（用 curl）：
   - 首頁回傳 200：`curl -s -o /dev/null -w "%{http_code}" http://localhost:8000/excel-web-tool/`
   - 多層的不存在路徑回傳 404 與 404 頁內容：`curl -s -w "%{http_code}" http://localhost:8000/excel-web-tool/a/b/c | tail -c 300`
   - `index.html` 引用的每個 JS／CSS 都回傳 200。
3. **瀏覽器驗證**（有瀏覽器工具時）：開啟首頁並檢查：
   - Console 沒有 CSP 違規（`Refused to execute`、`Refused to apply inline style`）或 JS 錯誤。
   - 切換兩個頁籤、點「下載測試範本」、用 `docs/templates/` 的範本上傳、貼上一小段假名單並載入。
   - 縮到手機寬度（約 375px）確認版面。
   沒有瀏覽器工具時，請使用者依上述清單手動確認，並告訴他們網址。
4. **注意快取**：Service Worker 會快取檔案。本機重複測試時，請在 DevTools → Application → Service Workers 勾選「Update on reload」，或用無痕視窗。
5. **結束**：驗證完成後停止伺服器。
