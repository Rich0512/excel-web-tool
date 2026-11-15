---
paths:
  - "tools/**"
  - ".claude/hooks/**"
---

# 工具腳本規則

- 開發環境是 Windows（Git Bash / PowerShell 都可能執行）。腳本用 Node.js ESM（`.mjs`），只用內建模組，不引入 npm 套件，不需要 `package.json`。字型裁切是唯一例外，使用 Python + fonttools（見 `/subset-fonts`）。
- **共用設定集中在 `tools/site-config.mjs`**：repo 路徑、`docs/`、子路徑、試算表副檔名、範本資料夾、版本號格式、`git()`、`isForbiddenSpreadsheet()`、`readHookInput()`、`block()`。hooks 以 `../../tools/site-config.mjs` 匯入。不要在個別腳本重新定義這些常數。
- 路徑一律由 `site-config.mjs` 依自身位置推算，不依賴目前工作目錄或 `CLAUDE_PROJECT_DIR`，因為使用者可能從上層資料夾啟動。
- 呼叫 git 用 `git()`（內部是 `execFileSync` 加參數陣列，不經過 shell），避免檔名中的空白與中文被誤解析。
- Hook 的結束碼：`0` 放行、`2` 阻擋（stderr 會回饋給 Claude）。其他結束碼不會阻擋。
- 權限規則要 Bash 與 PowerShell 各寫一份：`Bash(...)` 不會套用到 PowerShell 工具。hook 的 `matcher` 同理，要寫成 `Bash|PowerShell`。
- **失敗時的處理原則**：保護用的 guard（`guard-roster`、`guard-git`、子代理用的 `agent-readonly`）無法判斷時就阻擋（fail closed）；品質提醒類（`stop-check`）無法判斷時放行，避免卡住對話。
- Stop hook 要檢查 `stop_hook_active`，已經擋過一次就只提醒，避免無限迴圈。
- 修改 hook 或 `site-config.mjs` 後執行 `node tools/test-hooks.mjs`。新增攔截規則時，同時在裡面加上「應阻擋」與「應放行」兩種案例，避免誤擋正常操作。測試資料只能用虛構的檔名。
