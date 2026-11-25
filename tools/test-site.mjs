#!/usr/bin/env node
// ============================================
// 網站單元測試：依序執行 tools/tests/*.test.mjs，最後印出總結
// 用法：node tools/test-site.mjs [--verbose]
//   --verbose  印出每個檢查（預設只印每個檔案的結果與 FAIL 的項目）
// 每個測試檔在獨立的 Node 子行程執行，以 node:vm 載入 docs/js 並搭配假的 DOM（見 tools/tests/harness.mjs）。
// 測試資料全部是虛構的；不讀取任何名冊，也不發出網路請求。
// 結束碼：0 全部通過、1 有失敗或測試檔中斷
// ============================================

import { spawnSync } from 'node:child_process';
import { readdirSync } from 'node:fs';
import { join } from 'node:path';
import { ROOT } from './site-config.mjs';

const TEST_DIR = join(ROOT, 'tools', 'tests');
const VERBOSE = process.argv.includes('--verbose');
const files = readdirSync(TEST_DIR).filter(f => f.endsWith('.test.mjs')).sort();

if (!files.length) {
    console.error('tools/tests/ 裡沒有 *.test.mjs');
    process.exit(1);
}

let totalPass = 0, totalFail = 0;
const broken = [];
for (const f of files) {
    const r = spawnSync(process.execPath, [join(TEST_DIR, f)], { cwd: ROOT, encoding: 'utf8', timeout: 120000 });
    const lines = (r.stdout || '').split(/\r?\n/);
    const pass = lines.filter(l => l.startsWith('PASS')).length;
    const fails = lines.filter(l => l.startsWith('FAIL'));
    const title = (lines.find(l => l.startsWith('== ')) || '').replace(/^==\s*|\s*==$/g, '');
    totalPass += pass;
    totalFail += fails.length;
    // 測試檔沒跑完（例外中斷、逾時）或結束碼與結果不符時，視為中斷
    const finished = lines.some(l => l.startsWith('結果：'));
    const crashed = !finished || r.status === null || (r.status !== 0 && fails.length === 0);
    const mark = crashed ? '中斷' : fails.length ? 'FAIL' : 'PASS';
    console.log(`${mark.padEnd(4)} ${f}${title ? `（${title}）` : ''}：${pass}/${pass + fails.length} 通過`);
    if (VERBOSE) for (const l of lines.filter(l => l.startsWith('PASS') || l.startsWith('FAIL'))) console.log(`       ${l}`);
    else for (const l of fails) console.log(`       ${l}`);
    if (crashed) {
        broken.push(f);
        const why = r.error ? r.error.message : (r.stderr || '').trim().split(/\r?\n/).slice(0, 6).join('\n         ');
        console.log(`       測試檔沒有正常結束（結束碼 ${r.status ?? r.signal}）${why ? `：\n         ${why}` : ''}`);
    }
}

console.log(`\n合計：${files.length} 個測試檔，${totalPass + totalFail} 項檢查，PASS ${totalPass}、FAIL ${totalFail}${broken.length ? `，中斷 ${broken.length} 個檔案（${broken.join('、')}）` : ''}`);
process.exitCode = totalFail || broken.length ? 1 : 0;
