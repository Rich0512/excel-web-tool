---
name: release
description: 準備發佈網站新版本：同步調升 Service Worker 快取版本、index.html 的 ?v= 與 sitemap 日期，跑部署檢查，並引導 commit 與 push 上線。只要 docs/ 底下有任何變更（HTML、CSS、JS、字型、圖片、範本、manifest），或使用者提到發佈、上線、部署、push、更新網站、改版本號，或 Stop hook 提示「請執行 /release」時都要使用，即使使用者沒有明說「release」。
argument-hint: "[版本號 X.Y，省略則自動判斷]"
allowed-tools: Bash(node tools/bump-version.mjs *) Bash(node tools/check-site.mjs *) Bash(git status *) Bash(git diff *)
---

# 發佈新版本

網站從 `main` 分支的 `docs/` 直接部署到 GitHub Pages，**push 即上線**，沒有 CI 可以攔截錯誤。已開過網站的使用者靠 Service Worker 的 `CACHE_VERSION` 判斷是否更新，版本沒調升，他們會一直拿到舊檔案，而且很難察覺。

## 步驟

1. **調升版本**：執行 `node tools/bump-version.mjs $ARGUMENTS`。
   - 它會同步 `docs/service-worker.js`、`docs/index.html` 的所有 `?v=`、`docs/sitemap.xml` 的日期。
   - 不帶參數時，它會比對 `origin/main`：尚未調升就次版號 +1，已調升就只同步，所以重複執行是安全的。
   - 使用者指定了版本號才帶參數。
2. **同步預快取清單**：若這次新增、改名或刪除了 `docs/` 內的檔案，更新 `docs/service-worker.js` 的 `STATIC_ASSETS`（規則見 `.claude/rules/service-worker.md`）。
3. **檢查**：執行 `node tools/check-site.mjs --require-bump=origin/main`，必須全部通過。失敗時修正後重跑，不要跳過。
4. **摘要**：用 `git status --short` 與 `git diff --stat` 列出變更，並寫出新版本號。
5. **commit 與 push 要等使用者同意**。這次對話中使用者明確要求才執行。push 前提醒：push 到 main 後約 1～2 分鐘就上線，已安裝 PWA 的使用者下次開啟會自動更新。

## 範例

使用者：「我改了按鈕文字，幫我上線」
→ 執行 `node tools/bump-version.mjs`（例如 v3.8 → v3.9），確認 `STATIC_ASSETS` 不需變動，跑檢查，列出變更，並問「要我 commit 並 push 嗎？」
