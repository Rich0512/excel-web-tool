// ============================================
// 共用設定：tools/ 腳本與 .claude/hooks/ 共用的常數與小工具
// 只用 Node 內建模組；修改這裡會同時影響部署檢查與所有 hook
// ============================================

import { readFileSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import { join, dirname, resolve, relative, isAbsolute } from 'node:path';
import { fileURLToPath } from 'node:url';

export const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
export const SITE = 'docs';
export const SITE_DIR = join(ROOT, SITE);
export const BASE_PATH = '/excel-web-tool/';
export const TEMPLATE_DIR = `${SITE}/templates/`;
export const SPREADSHEET_RE = /\.(xlsx|xlsm|xls|csv|ods)$/i;
export const VERSION_RE = /const CACHE_VERSION = 'v([\d.]+)';/;

// 在 repo 根目錄執行 git，失敗時回傳 null（例如 ref 不存在）
export function git(args) {
    try {
        return execFileSync('git', ['-c', 'core.quotepath=false', ...args], { cwd: ROOT, encoding: 'utf8', stdio: ['ignore', 'pipe', 'ignore'] });
    } catch {
        return null;
    }
}

export const gitList = (args) => (git(args) || '').split('\n').map(s => s.trim()).filter(Boolean);

// 從 service-worker.js 原始碼取出版本號（不含 v），找不到回傳 undefined
export const versionOf = (swSource) => (swSource.match(VERSION_RE) || [])[1];

// 讀取某個 git ref 的快取版本；舊版網站在 repo 根目錄，docs/ 讀不到時改讀根目錄
export function versionAt(ref) {
    return versionOf(git(['show', `${ref}:${SITE}/service-worker.js`]) || git(['show', `${ref}:service-worker.js`]) || '');
}

export const isInside = (dir, path) => {
    const rel = relative(dir, path);
    return rel === '' || (!rel.startsWith('..') && !isAbsolute(rel));
};

// 路徑（相對於 base 或絕對路徑）是否為 docs/templates/ 內的檔案
export const isTemplatePath = (path, base = ROOT) => isInside(join(ROOT, TEMPLATE_DIR), resolve(base, path));

// 是否為不允許的試算表：副檔名是試算表，且不在 docs/templates/ 內
export const isForbiddenSpreadsheet = (path, base = ROOT) => SPREADSHEET_RE.test(path) && !isTemplatePath(path, base);

// 讀取 hook 的 stdin JSON；格式錯誤時回傳 null，由呼叫端決定如何處理
export function readHookInput() {
    try {
        return JSON.parse(readFileSync(0, 'utf8') || '{}');
    } catch {
        return null;
    }
}

// hook 阻擋：stderr 會回饋給 Claude，exit 2 才會真正阻擋
export function block(hookName, reason) {
    process.stderr.write(`[${hookName}] ${reason}\n`);
    process.exit(2);
}
