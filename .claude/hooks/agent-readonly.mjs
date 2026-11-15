#!/usr/bin/env node
// ============================================
// PreToolUse hook（Bash，掛在唯讀子代理 pages-auditor、frontend-reviewer 的 frontmatter）
// 子代理的 tools 欄位無法限制到個別指令，這裡用白名單確保 Bash 只做唯讀操作：
// - 只支援簡單的指令串接（&& || ; | 換行）；其他 shell 語法（$、\、`、&、<、( )、{ }）一律阻擋，
//   因為這裡不是完整的 shell 解析器，解析結果與 bash 不一致時就可能被繞過
// - 每一段指令都必須在白名單內，且不帶會寫檔或執行外部程式的選項
// - 禁止 > 重導（>/dev/null、2>/dev/null、2>&1 除外）
// 無法判斷時阻擋（fail closed）；被擋時請改用 Read / Grep / Glob
// ============================================

import { readHookInput, block } from '../../tools/site-config.mjs';

const HOOK = 'agent-readonly';
const GIT_READONLY = new Set(['status', 'diff', 'log', 'show', 'ls-files', 'rev-parse', 'cat-file', 'blame']);
// git 選項用白名單而非黑名單：git 接受長選項縮寫（--outp=x 等同 --output=x），黑名單擋不完
// 只比對完整名稱，所以任何縮寫都不會被放行；-- 之後是路徑，不再檢查
const GIT_FLAGS = new Set([
    '--short', '-s', '--porcelain', '--branch', '-b', '--untracked-files=no',
    '--stat', '--shortstat', '--numstat', '--name-only', '--name-status', '--cached', '--staged',
    '-p', '--patch', '-w', '--ignore-all-space', '--no-color', '--word-diff', '--no-ext-diff', '--no-textconv',
    '--oneline', '--graph', '--all', '--reverse', '--follow', '--decorate', '--abbrev-commit', '--no-merges', '--first-parent',
    '--others', '--exclude-standard', '--show-toplevel', '--abbrev-ref', '--verify', '-t', '-e', '-L', '-n',
]);
const GIT_VALUE_OPTS = new Set(['--format', '--pretty', '--date', '--since', '--until', '--author', '--grep', '--diff-filter', '--max-count', '--unified', '--abbrev']);
const gitOptionAllowed = (a) => GIT_FLAGS.has(a) || /^-(\d+|n\d+|U\d+)$/.test(a) || GIT_VALUE_OPTS.has(a.split('=')[0]) && a.includes('=');
// 這些工具沒有寫檔的選項（sort -o、uniq 第二個參數、file -C 會寫檔，所以不列入）
const TEXT_TOOLS = new Set(['ls', 'cat', 'head', 'tail', 'wc', 'grep', 'cut', 'tr', 'stat', 'echo', 'true']);

const input = readHookInput();
if (!input) block(HOOK, '無法解析 hook 輸入，唯讀子代理先行阻擋；請改用 Read / Grep / Glob');
const command = String(input.tool_input?.command || '');

// 逐字元掃描並追蹤引號狀態（兩段式的正規式比對會被 "'" 這類交錯引號騙過）：
// - 單引號內是純文字
// - 雙引號內與引號外的 $、`、\ 會被 bash 解讀（變數展開、指令替換、跳脫），一律阻擋
// - 引號內的 > | ; ( 只是文字（例如 grep "a|b"），換成佔位符後再判斷重導與切段
const quoted = [];
let masked = '';
let quote = null;
let buf = '';
for (const ch of command) {
    if (quote) {
        if (ch === quote) {
            masked += `\u0000${quoted.push(buf) - 1}\u0000`;
            quote = null;
            buf = '';
            continue;
        }
        if (quote === '"' && /[$`\\]/.test(ch)) block(HOOK, '唯讀子代理不可在雙引號內使用 $、` 或 \\，請改用單引號或 Read / Grep / Glob');
        buf += ch;
        continue;
    }
    if (ch === '"' || ch === "'") {
        quote = ch;
        continue;
    }
    if (/[$`\\\u0000]/.test(ch)) block(HOOK, '唯讀子代理不可使用 $、` 或 \\（變數展開、指令替換、跳脫），請改用 Read / Grep / Glob');
    masked += ch;
}
if (quote) block(HOOK, '引號不成對，無法判斷指令內容');
const restore = (s) => s.replace(/\u0000(\d+)\u0000/g, (_, i) => quoted[i]);

// 去掉無害的重導後，任何 > < 都視為讀寫檔案；剩下的 & 是背景執行或 |&
// 重導目標後面必須緊接空白、行尾或分隔符號，否則 >&1file、>/dev/nullX 會寫入名為 1file、/dev/nullX 的檔案
const rest = masked.replace(/\d?>&[0-2](?=\s|$|[|;&])|\d?>\s*\/dev\/null(?=\s|$|[|;&])/g, ' ').replace(/&&/g, '');
if (/>/.test(rest)) block(HOOK, '唯讀子代理不可用 > 寫入檔案');
if (/[<&(){}]/.test(rest)) block(HOOK, '唯讀子代理不可使用 <、&、( )、{ } 等 shell 語法');

const words = (segment) => segment.split(/\s+/).filter(Boolean).map(restore);

function allowed(w) {
    if (w.length === 0) return true;
    const [cmd, ...args] = w;
    if (cmd === 'git') {
        // 子指令前只允許 -C <dir> 與 --no-pager（-c、--exec-path 等全域選項都不允許）
        let i = 0;
        while (i < args.length && args[i].startsWith('-')) {
            if (args[i] === '-C') i += 2;
            else if (args[i] === '--no-pager') i += 1;
            else return false;
        }
        if (!GIT_READONLY.has(args[i])) return false;
        const rest = args.slice(i + 1);
        const end = rest.indexOf('--');
        return (end === -1 ? rest : rest.slice(0, end)).every(a => !a.startsWith('-') || gitOptionAllowed(a));
    }
    if (cmd === 'node') return args[0] === 'tools/check-site.mjs' || args[0] === './tools/check-site.mjs';
    // 只允許 unzip -p（輸出到 stdout）與 unzip -l（列出內容），且不帶其他選項
    if (cmd === 'unzip') return (args[0] === '-p' || args[0] === '-l') && !args.slice(1).some(a => a.startsWith('-'));
    return TEXT_TOOLS.has(cmd);
}

for (const segment of masked.split(/&&|\|\||;|\n|\|/)) {
    const w = words(segment.trim());
    if (!allowed(w)) block(HOOK, `唯讀子代理只能執行白名單內的指令（check-site、唯讀 git、ls/cat/grep 等），已阻擋：${restore(segment.trim()).slice(0, 80)}`);
}

process.exit(0);
