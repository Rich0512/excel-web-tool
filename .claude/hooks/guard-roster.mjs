#!/usr/bin/env node
// ============================================
// PreToolUse hook (Read / Grep / Bash / PowerShell)：防止讀取真實學生名冊
// 上層資料夾與其他位置的試算表可能含真實學生個資，只允許碰 docs/templates/ 內的假資料範本
// - Read：docs/templates/ 以外的試算表 → 阻擋
// - Grep：搜尋 repo 以外的位置且可能掃到試算表（未限定檔案類型，或限定為試算表）→ 阻擋
// - Bash / PowerShell：指令任何位置出現 repo 以外的試算表路徑（含萬用字元）→ 阻擋
//   repo 內的試算表由 .gitignore 與 guard-git 把關，這裡不擋，以免影響範本操作
// ============================================

import { resolve } from 'node:path';
import { ROOT, SPREADSHEET_RE, isInside, isForbiddenSpreadsheet, readHookInput, block } from '../../tools/site-config.mjs';

const HOOK = 'guard-roster';
// 路徑樣式的片段：允許夾在引號、括號、等號後面，例如 read_excel('../x.xlsx')、--file=../x.csv
const SPREADSHEET_PATH_RE = /[^\s"'`=<>|;(),]+\.(?:xlsx|xlsm|xls|csv|ods)\b/gi;

const input = readHookInput();
if (!input) block(HOOK, '無法解析 hook 輸入，為保護學生資料先行阻擋；請重試');
const cwd = input.cwd || ROOT;
const toolInput = input.tool_input || {};

switch (input.tool_name) {
    case 'Read': {
        const path = resolve(cwd, toolInput.file_path || '');
        if (isForbiddenSpreadsheet(path, cwd)) {
            block(HOOK, `拒絕讀取 docs/templates/ 以外的試算表（可能含真實學生個資）：${path}。需要測試資料請使用 docs/templates/ 內的範本。`);
        }
        break;
    }
    case 'Grep': {
        const searchPath = resolve(cwd, toolInput.path || '.');
        if (isInside(ROOT, searchPath)) break;
        const filter = `${toolInput.glob || ''} ${toolInput.type || ''}`;
        const mayHitSpreadsheets = !filter.trim() || /xlsx|xlsm|xls|csv|ods/i.test(filter);
        if (mayHitSpreadsheets) {
            block(HOOK, `在專案外部（${searchPath}）搜尋時必須用 glob 限定為非試算表的檔案類型，例如 "*.py"，以免讀到真實名冊。`);
        }
        break;
    }
    case 'Bash':
    case 'PowerShell': {
        const command = toolInput.command || '';
        for (const [match] of command.matchAll(SPREADSHEET_PATH_RE)) {
            const path = resolve(cwd, match.replace(/\\\\/g, '\\'));
            if (!isInside(ROOT, path) && SPREADSHEET_RE.test(path)) {
                block(HOOK, `指令涉及專案外部的試算表（可能是真實學生名冊）：${match}。請勿讀取、複製或解析真實名冊。`);
            }
        }
        break;
    }
}

process.exit(0);
