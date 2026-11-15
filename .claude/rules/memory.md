# 記憶規則（本專案補充）

記憶的格式與存放時機依 Claude Code 內建的 auto memory。這裡只列本專案額外的規定：

- 從 `web_demo/` 啟動時，記憶目錄是 `~/.claude/projects/excel-web-tool/memory/`（`.claude/settings.json` 的 `autoMemoryDirectory`），只存在這台電腦，不進 repo。從上層資料夾啟動時這個設定不生效，記憶會寫到上層資料夾的預設目錄，而且只有 `guard-roster` 會執行；這時要先提醒使用者改從 `web_demo/` 啟動。
- **絕不存學生個資**（姓名、班級、座號，或名冊內容的任何片段），也不存密碼、token。
- 已寫在 `CLAUDE.md`、`.claude/rules/`、skills 的內容不重複存。
- 只跟這台電腦有關、不限於本專案的事（例如 Python 安裝狀況）放在 `~/.claude/CLAUDE.md`，不放在專案記憶。
