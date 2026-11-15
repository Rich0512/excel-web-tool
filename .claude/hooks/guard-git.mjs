#!/usr/bin/env node
// ============================================
// PreToolUse hook (Bash / PowerShell)：攔截可能把學生名冊送上 GitHub 的 git 指令
// - git add -f / 指名加入 docs/templates/ 以外的試算表 → 阻擋
// - git commit → 暫存區有 docs/templates/ 以外的試算表就阻擋
// - git push   → 執行完整部署檢查（含相對 origin/main 是否調升版本），失敗就阻擋（push 到 main 即上線）
// 註：`git add -A && git commit` 這類串接，commit 檢查時 add 尚未執行；
//     此時由 .gitignore 排除試算表，push 前的部署檢查會再把關一次。
// ============================================

import { execFileSync } from 'node:child_process';
import { join } from 'node:path';
import { ROOT, TEMPLATE_DIR, gitList, isForbiddenSpreadsheet, readHookInput, block } from '../../tools/site-config.mjs';

const HOOK = 'guard-git';
// 這些 git 全域選項會吃掉下一個參數，判斷子指令時要一併略過
const GIT_OPTIONS_WITH_VALUE = new Set(['-C', '-c', '--git-dir', '--work-tree', '--namespace', '--exec-path', '--config-env']);

const input = readHookInput();
if (!input) block(HOOK, '無法解析 hook 輸入，為保護學生資料先行阻擋；請重試指令');
const command = input.tool_input?.command || '';
const cwd = input.cwd || ROOT;

const unquote = (s) => s.replace(/^["']|["']$/g, '');

// 依空白切出參數，保留引號內的空白（例如 -C "path with spaces"）
const splitWords = (segment) => [...segment.matchAll(/"[^"]*"|'[^']*'|\S+/g)].map(m => unquote(m[0]));

// 找出 git 子指令與其後的參數，例如 `git -C repo commit -m x` → { sub: 'commit', args: ['-m', 'x'] }
function parseGit(words) {
    const gitIdx = words.findIndex(w => /(^|[\\/])git(\.exe)?$/i.test(w));
    if (gitIdx === -1) return null;
    for (let i = gitIdx + 1; i < words.length; i++) {
        const w = words[i];
        if (GIT_OPTIONS_WITH_VALUE.has(w)) {
            i++;
            continue;
        }
        if (w.startsWith('-')) continue;
        return { sub: w, args: words.slice(i + 1) };
    }
    return null;
}

// 依 shell 分隔符號切開，逐段判斷
for (const segment of command.split(/&&|\|\||;|\n|\|/)) {
    const parsed = parseGit(splitWords(segment.trim()));
    if (!parsed) continue;
    const { sub, args } = parsed;

    if (sub === 'add') {
        const named = args.filter(a => !a.startsWith('-') && isForbiddenSpreadsheet(a, cwd));
        if (named.length) block(HOOK, `不可加入 ${TEMPLATE_DIR} 以外的試算表（可能是真實學生名冊）：${named.join(', ')}`);
        // 含合併的短選項，例如 -fA
        if (args.some(a => a === '--force' || /^-[a-zA-Z]*f[a-zA-Z]*$/.test(a))) {
            block(HOOK, 'git add -f 會繞過 .gitignore 的名冊防護，請改為指名加入非試算表檔案');
        }
    }

    if (sub === 'commit') {
        const bad = gitList(['diff', '--cached', '--name-only']).filter(f => isForbiddenSpreadsheet(f));
        if (bad.length) block(HOOK, `暫存區含有 ${TEMPLATE_DIR} 以外的試算表，請先 git restore --staged：${bad.join(', ')}`);
    }

    if (sub === 'push') {
        try {
            // 與已上線的 origin/main 比較，確認有調升快取版本
            execFileSync('node', [join(ROOT, 'tools', 'check-site.mjs'), '--require-bump=origin/main'], { cwd: ROOT, encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'] });
        } catch (e) {
            block(HOOK, `push 會直接部署到 GitHub Pages，但部署檢查未通過：\n${e.stdout || e.message}`);
        }
    }
}

process.exit(0);
