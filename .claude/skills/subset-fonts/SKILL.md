---
name: subset-fonts
description: 重新裁切網站使用的 Noto Sans TC 中文字型（docs/fonts/*.woff2）。當部署檢查回報「字型未涵蓋 N 個介面用字」「介面用字不在字型介面檔」或 unicode-range 與清單不一致、介面新增了罕見中文字，或使用者提到字型缺字、字顯示成不同字體、字型檔太大時使用。
allowed-tools: Bash(git show *) Bash(node tools/check-site.mjs)
---

# 重新裁切中文字型

完整的 Noto Sans TC 每個字重約 7 MB，對手機使用者太重，所以網站使用裁切版。每個字重拆成兩個檔，由 `docs/fonts/fonts.css` 的 `unicode-range` 決定何時下載：

| 檔案 | 內容 | 大小（每個字重） | 何時下載 |
|---|---|---|---|
| `NotoSansTC-{Regular,Bold}-ui.woff2` | ASCII、Big5 符號、介面（HTML、自家 JS、manifest）實際用到的字 | 約 130 KB | 首屏（`index.html` 只 preload Regular 的這個檔） |
| `NotoSansTC-{Regular,Bold}-ext.woff2` | Big5 常用字（5,401 字）中不在介面檔的字 | 約 0.8 MB | 畫面出現這些字時（例如學生姓名） |

兩個檔的字元互斥，各自實際收錄的字存在 `tools/subset-chars-ui.txt`、`tools/subset-chars-ext.txt`。`tools/check-site.mjs` 會檢查介面用字都在 ui 清單內（落在 ext 會讓首屏多下載大檔），以及 `fonts.css` 的 `unicode-range`：ui 精確等於 ui 清單；ext 是粗範圍（CJK 符號、假名、CJK 統一表意文字中所有不在 ui 的碼位），必須涵蓋 ext 清單且不含任何 ui 字。粗範圍讓 `fonts.css` 從約 78 KB 降到約 30 KB（gzip 後約 7 KB）（它會擋住首次繪製），代價是名單只有罕用字時也會下載 ext 檔。ext 不在 Service Worker 預快取，第一次用到時才下載並快取。`fonts.css` 中 BEGIN／END 標記之間由腳本產生，不要手動修改。

名冊裡不在 Big5 常用字的罕用字（例如部分學生姓名）會由系統字型補字，這是預期行為，不需要為此重新裁切。

## 步驟

1. **取出完整字型**：在 scratchpad（不是專案資料夾）建立暫存資料夾，從 git 歷史取出原始 TTF（`619c22a` 是網站還在 repo 根目錄、TTF 尚未刪除時的 commit，所以路徑是 `fonts/` 而不是 `docs/fonts/`）：
   ```bash
   git show 619c22a:fonts/NotoSansTC-Regular.ttf > <暫存>/NotoSansTC-Regular.ttf
   git show 619c22a:fonts/NotoSansTC-Bold.ttf    > <暫存>/NotoSansTC-Bold.ttf
   ```
   必須用 `>` 寫入二進位檔；PowerShell 的 `>` 會改變編碼，請在 Git Bash 執行。
2. **準備 Python 環境**：在 scratchpad 建立 venv，安裝 `fonttools brotli`。不要裝進專案或上層資料夾的 venv。
   ```bash
   py -V:3.12 -m venv <scratchpad>/venv
   <scratchpad>/venv/Scripts/python.exe -m pip install fonttools brotli
   ```
   一定要用 `py -V:3.12`：這台電腦的 `python`／`python3` 會先指到 Microsoft Store 捷徑或 msys64 的 Python，`py -3.12` 也可能選到 32 位元版本。之後一律用 venv 裡的 `Scripts/python.exe`，不要用 `python`。
3. **裁切**：`<scratchpad>/venv/Scripts/python.exe tools/subset-fonts.py <暫存>`（路徑用反斜線或相對路徑；`py` 遇到 `C:/…` 會誤報 exit 103）。它會覆寫 4 個 woff2、兩份字元清單與 `fonts.css` 的 Noto Sans TC 區塊，並印出檔案大小；可重複執行。ui 檔明顯超過 200 KB（介面用字變多，首屏變慢）或 ext 檔超過 1 MB 時，回報給使用者再繼續。若印出「原始字型沒有 N 個介面用字」，那些字無法靠重新裁切補上，請改用其他字。
4. **檢查**：`node tools/check-site.mjs` 必須通過。檔名有變動時（例如新增字重），要同步改 `service-worker.js` 的 `STATIC_ASSETS` 與 `index.html` 的 preload。
5. **發佈**：字型屬於網站檔案，接著執行 `/release`。
