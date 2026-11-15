# 記憶規則

記憶存放在 `~/.claude/projects/excel-web-tool/memory/`（由 `.claude/settings.json` 的 `autoMemoryDirectory` 指定），只存在這台電腦，不會進 repo。`MEMORY.md` 是索引，每次對話開頭自動載入。

## 何時讀取

訊息涉及過去的決定、使用者偏好、上次做到哪裡（例如「上次那個」「我們決定的」）時，先看 `MEMORY.md` 的描述，只讀相關的檔案。內容已經在目前對話中就不用再讀。

## 何時儲存

每次實質回覆後判斷：

1. 這段對話有沒有使用者的指示或偏好、經過推理得到的結論、或之後還會用到的事實？都沒有就不存。
2. 有沒有重疊的記憶？有就更新；內容矛盾就取代；沒有才新增。
3. 值得存嗎？從程式碼或 git 查得到的就不存；需要綜合多個來源、或只有使用者知道的才存。

## 格式

每則記憶一個檔案，檔名用 kebab-case：

```markdown
---
name: descriptive-name
description: 一行摘要，用來判斷相關性
type: user | feedback | project | reference
---

內容。feedback 與 project 類型要加上：
**Why:** 原因
**How to apply:** 何時適用
```

新增或更新後，在 `MEMORY.md` 加一行索引：`- [標題](檔名.md) — 一句話說明`。日期一律寫成絕對日期（例如 2026-11-15），不要寫「下週」。

## 不要存

- 從程式碼、`CLAUDE.md`、`.claude/rules/` 讀得到的規則與架構
- git 歷史、已修好的 bug（修正就在程式碼裡）
- 只在這次對話有用的暫時細節
- 學生個資、密碼、token
