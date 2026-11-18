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

console.log('--- guard-roster: 從上層資料夾啟動（上層 .claude/settings.json 掛同一支 hook）');
const parent = join(ROOT, '..');
const fromParent = (tool_name, tool_input) => ({ tool_name, cwd: parent, tool_input });
run('guard-roster', fromParent('Read', { file_path: 'roster.xlsx' }), 'Read roster.xlsx', 2);
run('guard-roster', fromParent('Bash', { command: 'python generate_report.py roster.xlsx' }), 'python with roster.xlsx', 2);
run('guard-roster', fromParent('PowerShell', { command: 'Get-ChildItem *.csv' }), 'glob *.csv', 2);
run('guard-roster', fromParent('Grep', { pattern: 'x' }), 'Grep without filter', 2);
run('guard-roster', fromParent('Read', { file_path: 'generate_report.py' }), 'Read generate_report.py (allowed)', 0);
run('guard-roster', fromParent('Grep', { pattern: 'x', glob: '*.py' }), 'Grep *.py (allowed)', 0);
run('guard-roster', fromParent('Read', { file_path: 'web_demo/docs/templates/115學期名冊_範本.xlsx' }), 'Read template via web_demo/ (allowed)', 0);

console.log('--- guard-git');
run('guard-git', shell('git -C "my repo" add -fA'), 'combined -fA with quoted -C', 2);
run('guard-git', shell('git add ../docs/templates/x.xlsx'), '../docs/templates path (was loose)', 2);
run('guard-git', shell('git add other/docs/templates/x.xlsx'), 'other/docs/templates path (was loose)', 2);
run('guard-git', shell('git add docs/templates/x.xlsx'), 'real templates path (allowed)', 0);
run('guard-git', shell('git add -A'), 'git add -A (allowed, .gitignore covers)', 0);
run('guard-git', shell('git status | head'), 'git status (allowed)', 0);
run('guard-git', shell('git push origin main'), 'push with checks passing', 0);
run('guard-git', 'not json', 'malformed input -> fail closed', 2);

console.log('--- agent-readonly: 應阻擋');
run('agent-readonly', shell('echo x > docs/index.html'), 'redirect write', 2);
run('agent-readonly', shell('git status && rm -rf docs'), 'chained rm', 2);
run('agent-readonly', shell('git checkout -- docs'), 'git checkout', 2);
run('agent-readonly', shell('git diff --output=x.txt'), 'git diff --output', 2);
run('agent-readonly', shell('node -e "require(\'fs\').writeFileSync(\'x\',\'\')"'), 'node -e', 2);
run('agent-readonly', shell('cat $(ls)'), 'command substitution', 2);
run('agent-readonly', shell('unzip docs/templates/fake.xlsx'), 'unzip extract', 2);
run('agent-readonly', shell('git -c core.pager=touch log'), 'git -c core.pager', 2);
run('agent-readonly', shell('git diff --ext-diff'), 'git diff --ext-diff', 2);
run('agent-readonly', shell('sort -o out.txt docs/js/app.js'), 'sort -o (not allowed)', 2);
run('agent-readonly', shell('echo \\"a; rm -rf x; \\"'), 'backslash-escaped quotes', 2);
run('agent-readonly', shell('ls & rm x'), 'lone & background', 2);
run('agent-readonly', shell('cat <> out.txt'), '<> read-write redirect', 2);
run('agent-readonly', shell('cat <(ls)'), 'process substitution', 2);
run('agent-readonly', shell('echo "${HOME}"'), '$ inside double quotes', 2);
run('agent-readonly', shell('unzip -p -d out docs/templates/fake.xlsx'), 'unzip -p with extra flag', 2);
run('agent-readonly', shell(`echo "'" $(rm x) "'"`), 'interleaved quotes hide $()', 2);
run('agent-readonly', shell(`echo "'" ; rm x ; "'"`), 'interleaved quotes hide ;', 2);
run('agent-readonly', shell(`grep "a" 'b`), 'unterminated quote', 2);
run('agent-readonly', shell('git diff --outp=x.txt'), 'abbreviated --output', 2);
run('agent-readonly', shell('git log --ext'), 'abbreviated --ext-diff', 2);
run('agent-readonly', shell('git --exec-path=. status'), 'global --exec-path', 2);
run('agent-readonly', shell('git log --form=%H'), 'abbreviated --format (not exact)', 2);
run('agent-readonly', shell('git log --oneline -5 --format=%h --stat -- docs/-x'), 'allowlisted git options (allowed)', 0);
run('agent-readonly', shell('git --no-pager diff --name-only HEAD~1 -- docs'), '--no-pager diff (allowed)', 0);
run('agent-readonly', shell('ls >&1file'), '>&1file writes file', 2);
run('agent-readonly', shell('ls >/dev/nullX'), '>/dev/nullX writes file', 2);
run('agent-readonly', shell('ls 2>&10'), '2>&10 not a safe fd', 2);
run('agent-readonly', shell('git status 2>&1 | head -3'), '2>&1 | head (allowed)', 0);
run('agent-readonly', shell('git log >/dev/null; git status'), '>/dev/null; (allowed)', 0);
run('agent-readonly', 'not json', 'malformed input -> fail closed', 2);

// 子代理的工作目錄可能是 docs/（沿用主對話當下的位置），check-site 的路徑依 cwd 解析後比對
const docsDir = join(ROOT, 'docs');
const atDocs = (command) => ({ tool_name: 'Bash', cwd: docsDir, tool_input: { command } });
const rootSlash = ROOT.replace(/\\/g, '/');
const rootMsys = rootSlash.replace(/^([a-zA-Z]):/, (_, d) => `/${d.toLowerCase()}`);
run('agent-readonly', atDocs('node tools/check-site.mjs'), 'cwd=docs: tools/check-site.mjs 指向 docs/tools（不存在）', 2);
run('agent-readonly', atDocs('node ../tools/bump-version.mjs'), 'cwd=docs: 其他腳本', 2);
run('agent-readonly', atDocs('node ../../tools/check-site.mjs'), 'cwd=docs: repo 外同名檔案', 2);
run('agent-readonly', atDocs('node --require ./x.js ../tools/check-site.mjs'), 'cwd=docs: node 選項', 2);
run('agent-readonly', atDocs(`node ${rootSlash}/tools/check-site.mjs.bak`), '相似檔名', 2);

console.log('--- agent-readonly: 應放行');
run('agent-readonly', atDocs('node ../tools/check-site.mjs --require-bump=origin/main'), 'cwd=docs: ../tools/check-site.mjs', 0);
run('agent-readonly', atDocs(`node ${rootSlash}/tools/check-site.mjs`), 'cwd=docs: 絕對路徑（C:/…）', 0);
run('agent-readonly', atDocs(`node ${rootMsys}/tools/check-site.mjs`), 'cwd=docs: Git Bash 路徑（/c/…）', 0);
run('agent-readonly', atDocs(`node ${rootSlash.toUpperCase()}/tools/check-site.mjs`), 'cwd=docs: 大小寫不同（Windows）', process.platform === 'win32' ? 0 : 2);
run('agent-readonly', shell("grep -n 'a$|b\\\\' docs/js/app.js"), '$ and \\ inside single quotes', 0);
run('agent-readonly', shell('grep -c "escapeHtml(" docs/js/app.js'), '( inside double quotes', 0);
run('agent-readonly', shell('node tools/check-site.mjs --require-bump=origin/main'), 'check-site', 0);
run('agent-readonly', shell('git status --short -- docs'), 'git status', 0);
run('agent-readonly', shell('git -C "my repo" diff HEAD -- docs 2>/dev/null'), 'git -C diff + 2>/dev/null', 0);
run('agent-readonly', shell('git log --oneline -5 | head -3'), 'git log | head', 0);
run('agent-readonly', shell('unzip -p docs/templates/fake.xlsx xl/sharedStrings.xml | wc -c'), 'unzip -p | wc', 0);
run('agent-readonly', shell('grep -n "=> |a;b" docs/js/app.js'), '> | ; inside quotes', 0);
run('agent-readonly', shell('grep -n "x" docs/js/app.js > out.txt'), 'redirect after quoted arg', 2);

console.log('--- stop-check');
run('stop-check', '{"stop_hook_active":false}', 'current state passes', 0);
run('stop-check', 'not json', 'malformed input -> fail open', 0);

console.log(failures ? `\n✗ ${failures} 個測試失敗` : '\n✓ 全部通過');
process.exit(failures ? 1 : 0);
