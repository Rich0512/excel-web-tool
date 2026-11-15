#!/usr/bin/env node
// ============================================
// Hook 回歸測試：以模擬的 stdin JSON 執行 .claude/hooks/ 的每個 hook，確認放行／阻擋結果
// 用法：node tools/test-hooks.mjs（修改 hook 或 tools/site-config.mjs 後執行）
// 註：push 與 stop-check 的案例會實際跑部署檢查，需在 check-site.mjs 通過的狀態下執行
// ============================================

import { execFileSync } from 'node:child_process';
import { join } from 'node:path';
import { ROOT } from './site-config.mjs';

const cwd = ROOT;
let failures = 0;

function run(hook, input, label, expect) {
    let code = 0;
    let err = '';
    try {
        execFileSync('node', [join(ROOT, '.claude', 'hooks', `${hook}.mjs`)], {
            cwd: ROOT,
            input: typeof input === 'string' ? input : JSON.stringify(input),
            stdio: ['pipe', 'pipe', 'pipe'],
        });
    } catch (e) {
        code = e.status;
        err = String(e.stderr).replace(/\n/g, ' ').slice(0, 70);
    }
    if (code !== expect) failures++;
    console.log(`${code === expect ? 'PASS' : 'FAIL'} [${code}] ${label}${err ? ` :: ${err}` : ''}`);
}
const shell = (command, tool = 'Bash') => ({ tool_name: tool, cwd, tool_input: { command } });

console.log('--- guard-roster: bypasses from review');
run('guard-roster', shell(`python -c "pd.read_excel('../roster.xlsx')"`), 'embedded in python -c', 2);
run('guard-roster', shell(`python -c "load_workbook('..\\\\x.xlsx')"`), 'escaped backslashes', 2);
run('guard-roster', shell('tool --file=../x.csv'), '--file=../x.csv', 2);
run('guard-roster', shell('cmd >..\\x.csv', 'PowerShell'), 'redirect >..\\x.csv', 2);
run('guard-roster', shell('Get-Content ..\\*.xlsx', 'PowerShell'), 'glob ..\\*.xlsx', 2);
run('guard-roster', shell('cp ../*.csv .'), 'glob ../*.csv', 2);
run('guard-roster', { tool_name: 'Grep', cwd, tool_input: { pattern: 'x', path: '..' } }, 'Grep parent, no filter', 2);
run('guard-roster', { tool_name: 'Grep', cwd, tool_input: { pattern: 'x', path: '..', glob: '*.csv' } }, 'Grep parent *.csv', 2);
run('guard-roster', { tool_name: 'Grep', cwd, tool_input: { pattern: 'x', path: '..', glob: '*.py' } }, 'Grep parent *.py (allowed)', 0);
run('guard-roster', { tool_name: 'Grep', cwd, tool_input: { pattern: 'x', path: 'docs' } }, 'Grep inside repo (allowed)', 0);

console.log('--- guard-roster: legitimate use still allowed');
run('guard-roster', shell('ls docs/templates/*.xlsx'), 'ls templates glob', 0);
run('guard-roster', shell('git add docs/templates/new.xlsx'), 'git add template', 0);
run('guard-roster', { tool_name: 'Read', cwd, tool_input: { file_path: 'docs/templates/115學期名冊_範本.xlsx' } }, 'Read template', 0);
run('guard-roster', { tool_name: 'Read', cwd, tool_input: { file_path: '../generate_report.py' } }, 'Read parent .py', 0);
run('guard-roster', { tool_name: 'Read', cwd, tool_input: { file_path: '../roster.xlsx' } }, 'Read parent roster', 2);
run('guard-roster', 'not json', 'malformed input -> fail closed', 2);

console.log('--- guard-git');
run('guard-git', shell('git -C "my repo" add -fA'), 'combined -fA with quoted -C', 2);
run('guard-git', shell('git add ../docs/templates/x.xlsx'), '../docs/templates path (was loose)', 2);
run('guard-git', shell('git add other/docs/templates/x.xlsx'), 'other/docs/templates path (was loose)', 2);
run('guard-git', shell('git add docs/templates/x.xlsx'), 'real templates path (allowed)', 0);
run('guard-git', shell('git add -A'), 'git add -A (allowed, .gitignore covers)', 0);
run('guard-git', shell('git status | head'), 'git status (allowed)', 0);
run('guard-git', shell('git push origin main'), 'push with checks passing', 0);
run('guard-git', 'not json', 'malformed input -> fail closed', 2);

console.log('--- stop-check');
run('stop-check', '{"stop_hook_active":false}', 'current state passes', 0);
run('stop-check', 'not json', 'malformed input -> fail open', 0);

console.log(failures ? `\n✗ ${failures} 個測試失敗` : '\n✓ 全部通過');
process.exit(failures ? 1 : 0);
