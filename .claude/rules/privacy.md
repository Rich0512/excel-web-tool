# 學生個資規則（每次都載入）

這個工具處理的是國小學生的姓名、班級、座號。外洩的後果是真實的，也無法撤回，所以以下規則沒有例外。

## 真實名冊在哪裡

- 上層資料夾 `../`（Excel_Integration）的 `*.xlsx` 是**真實學生名冊**。不要讀取、複製、解析，也不要在任何輸出中顯示其中的姓名。`guard-roster` hook 會擋下 Read 試算表、在 repo 外用 Grep 掃描試算表，以及指令中出現 repo 外的試算表路徑；被擋時不要嘗試繞過，hook 只是最後一道防線，不是可以試探的邊界。
- 需要測試資料時，只用 `docs/templates/` 內的假資料範本。範本一律存成 `.xlsx`，「姓名」欄的值必須是 `測試生<班級>_<座號>`（例如 `測試生201_23`）；`check-site.mjs` 會檢查，push 前也會再檢查一次。
- 試算表只能放在 `docs/templates/`。`.gitignore` 排除其他所有試算表，`guard-git` hook 會擋下 `git add -f` 與提交名冊。

## 程式碼裡的個資

學生資料只能存在記憶體中，關閉分頁即消失：
- 不可寫入 localStorage、sessionStorage、IndexedDB、Cache Storage。localStorage 只存社團對照設定（key `yenping_club_config`：社團名稱 → 星期）。
- 不可 `console.log` 學生資料；錯誤訊息只記錄例外本身。
- 不可發出任何網路請求（CSP 的 `connect-src 'self'` 是最後一道防線，不是唯一一道）。
- 匯出只透過使用者主動下載的 Excel 檔。

## 回報與稽核

需要描述名冊內容時（例如稽核範本是否為假資料），只回報筆數與判斷結果，不列出姓名。
