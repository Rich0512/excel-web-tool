# 延平國小社團課表整理工具（網頁版）

純前端靜態網站，把社團錄取名冊（Excel 或從網頁貼上的名單）轉成學生個人週課表，並支援線上抽籤。所有資料只在瀏覽器內處理。

- 網址：https://rich0512.github.io/excel-web-tool/
- 部署：GitHub Pages 從 `main` 分支的 `docs/` 發佈。**push 到 main 就是上線**，沒有建置步驟與 CI。

## 指令

```bash
node tools/serve.mjs              # 本機預覽 http://localhost:8000/excel-web-tool/（模擬 Pages 子路徑與 404）
node tools/check-site.mjs         # 部署檢查；加 --require-bump=origin/main 會一併確認版本已調升
node tools/bump-version.mjs [X.Y] # 同步調升快取版本、?v=、sitemap 日期
node tools/test-hooks.mjs         # hook 回歸測試（修改 .claude/hooks/ 或 tools/site-config.mjs 後）
```

不需要安裝套件，也沒有建置步驟；網站驗證靠 `check-site.mjs` 加上 `/preview` 實際操作。

## 結構

```
docs/                  對外發佈的網站（只有這裡會出現在網站上）
  index.html 404.html  主頁與 404（404 用 /excel-web-tool/ 絕對路徑）
  service-worker.js    離線快取；STATIC_ASSETS 要列出所有執行時需要的檔案
  js/                  原生 JS，依序 defer 載入並共用全域函式（順序見 rules/frontend.md）
  js/libs/             第三方 ExcelJS，不要修改
  css/ fonts/ images/  fonts 中的 Noto Sans TC 已裁切
  templates/           唯一允許的試算表：假資料範本
tools/                 檢查、版本、預覽、字型裁切腳本；共用常數在 site-config.mjs
.claude/               rules、skills、agents、hooks
```

## 容易出錯的地方

細節都在 `.claude/rules/`，這裡只列重點：

- **上層資料夾 `../` 有真實學生名冊**，學生資料不可離開記憶體 → `rules/privacy.md`
- **改了 `docs/` 就要調升版本**，否則使用者會一直拿到舊快取 → `/release`、`rules/service-worker.md`
- **CSP 沒有 `'unsafe-inline'`**，不能寫 `style=`、`onclick=` → `rules/frontend.md`
- GitHub Pages 無法設定 HTTP header，安全設定只能寫在 `<meta>` CSP。
- 介面新增罕見中文字時字型可能缺字 → `/subset-fonts`

## Claude Code 設定

| 類型 | 名稱 | 用途 |
|---|---|---|
| rules | `privacy`、`memory`（常駐）；`frontend`、`service-worker`、`tooling`（依檔案載入） | 碰到對應檔案時自動載入 |
| skills | `/release`、`/preview`、`/subset-fonts` | 發佈、本機驗證、字型裁切 |
| agents | `pages-auditor`、`frontend-reviewer` | 上線前稽核、前端審查（唯讀） |
| hooks | `guard-roster`、`guard-git`（PreToolUse）、`stop-check`（Stop） | 擋讀取名冊、擋提交名冊、push 前與回覆結束時檢查 |
| memory | `~/.claude/projects/excel-web-tool/memory/` | 跨對話記憶（本機，不進 repo） |

這些設定只在**以 `web_demo/` 為工作目錄啟動 Claude Code** 時生效。
