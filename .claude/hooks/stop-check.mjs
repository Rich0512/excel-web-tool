#!/usr/bin/env node
// ============================================
// Stop hook：Claude 結束回覆前，若 docs/ 有未提交的變更就跑部署檢查
// 檢查失敗（含「改了網站卻沒調升快取版本」）→ exit 2，讓 Claude 繼續修正
// stop_hook_active 為 true 代表已被擋過一次，只提醒不再阻擋，避免無限迴圈
// 這是品質提醒而非安全防線：輸入格式錯誤或 git 無法執行時放行（exit 0）
// ============================================

import { execFileSync } from 'node:child_process';
import { join } from 'node:path';
import { ROOT, SITE, git, readHookInput, block } from '../../tools/site-config.mjs';

const HOOK = 'stop-check';
const input = readHookInput() || {};

const status = git(['status', '--porcelain', '--', SITE]);
if (!status || !status.trim()) process.exit(0);

try {
    execFileSync('node', [join(ROOT, 'tools', 'check-site.mjs'), '--require-bump'], { cwd: ROOT, encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'] });
} catch (e) {
    const report = e.stdout || e.message;
    if (input.stop_hook_active) {
        console.log(JSON.stringify({ systemMessage: `[${HOOK}] 部署檢查仍未通過：\n${report}` }));
        process.exit(0);
    }
    block(HOOK, `部署檢查未通過，請修正後再結束：\n${report}`);
}
process.exit(0);
