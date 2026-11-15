---
name: subset-fonts
description: 重新裁切網站使用的 Noto Sans TC 中文字型（docs/fonts/*.woff2）。當部署檢查回報「字型未涵蓋 N 個介面用字」、介面新增了罕見中文字，或使用者提到字型缺字、字顯示成不同字體、字型檔太大時使用。
allowed-tools: Bash(git show *) Bash(node tools/check-site.mjs)
---

# 重新裁切中文字型

完整的 Noto Sans TC 每個字重約 7 MB，對手機使用者太重，所以網站使用裁切版（約 0.9 MB），只含 ASCII、Big5 符號、Big5 常用字 5,401 字，以及介面（HTML、JS、manifest）實際用到的字。裁切時用到的字元清單存在 `tools/subset-chars.txt`，`tools/check-site.mjs` 會拿它比對介面用字。

名冊裡的罕用字（例如學生姓名）會由系統字型補字，這是預期行為，不需要為此重新裁切。

## 步驟

1. **取出完整字型**：在 scratchpad（不是專案資料夾）建立暫存資料夾，從 git 歷史取出原始 TTF（`619c22a` 是網站還在 repo 根目錄、TTF 尚未刪除時的 commit，所以路徑是 `fonts/` 而不是 `docs/fonts/`）：
   ```bash
   git show 619c22a:fonts/NotoSansTC-Regular.ttf > <暫存>/NotoSansTC-Regular.ttf
   git show 619c22a:fonts/NotoSansTC-Bold.ttf    > <暫存>/NotoSansTC-Bold.ttf
   ```
   必須用 `>` 寫入二進位檔；PowerShell 的 `>` 會改變編碼，請在 Git Bash 執行。
2. **準備 Python 環境**：在 scratchpad 建立 venv，安裝 `fonttools brotli`。不要裝進專案或上層資料夾的 venv。
3. **裁切**：`<venv python> tools/subset-fonts.py <暫存>`。它會覆寫 `docs/fonts/NotoSansTC-{Regular,Bold}.woff2` 與 `tools/subset-chars.txt`，並印出檔案大小。大小明顯超過 1.2 MB 時，回報給使用者再繼續。
4. **檢查**：`node tools/check-site.mjs` 必須通過。
5. **發佈**：字型屬於網站檔案，接著執行 `/release`。
