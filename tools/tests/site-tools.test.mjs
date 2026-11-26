// tools/check-site.mjs 與 tools/bump-version.mjs：涵蓋 docs/ 根目錄的每個頁面（含使用說明頁 guide.html）
// - check-site：每頁的 ?v= 一致、自家 CSS 帶 ?v=、CSP 與預期相同、每頁的 preload（含 404 的絕對路徑）、
//   根目錄頁面（404 除外）都要預快取、sitemap 列出每個頁面
// - bump-version：每頁的 ?v= 都同步、sitemap 的每個 <lastmod> 都改成今天
// 在系統暫存資料夾建立 docs/ 與工具的副本再執行，不會改到 repo；副本不含任何試算表（templates/ 留空），
// 頁面內容是虛構的最小頁面。
//
// 環境變數 TOOLS_ROOT：只給「對照組驗證」用，改從另一個副本的 tools/ 取得 check-site.mjs、bump-version.mjs
// （例如舊版），確認測試抓得到舊版的問題。預設是本 repo。docs/ 的來源依 harness 的 SITE_ROOT。
import { spawnSync } from 'node:child_process';
import { cpSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync, existsSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { makeReporter, SITE_ROOT } from './harness.mjs';
import { ROOT, SITE, SPREADSHEET_RE, versionOf } from '../site-config.mjs';
const R = makeReporter('部署工具：check-site 與 bump-version 涵蓋每個頁面');

const TOOLS_SRC = join(process.env.TOOLS_ROOT ? resolve(process.env.TOOLS_ROOT) : ROOT, 'tools');
const TOOL_FILES = ['check-site.mjs', 'bump-version.mjs', 'site-config.mjs', 'template-check.mjs', 'subset-chars-ui.txt', 'subset-chars-ext.txt'];
const SRC_DOCS = join(SITE_ROOT, SITE);
const VERSION = versionOf(readFileSync(join(SRC_DOCS, 'service-worker.js'), 'utf8'));
const GUIDE_CSP = "default-src 'self'; script-src 'none'; style-src 'self'; img-src 'self'; font-src 'self'; object-src 'none'; base-uri 'self'; form-action 'none'";

// 虛構的最小說明頁（只用 ASCII，避免字型檢查的訊息干擾）
const guideHtml = (v = VERSION, { csp = GUIDE_CSP, extraHead = '', body = '<h1>Guide</h1>' } = {}) => `<!DOCTYPE html>
<html lang="zh-Hant">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="Content-Security-Policy" content="${csp}">
    <title>Guide</title>
    <link rel="preload" href="fonts/NotoSansTC-Regular-ui.woff2" as="font" type="font/woff2" crossorigin>
    <link rel="stylesheet" href="fonts/fonts.css?v=${v}">
    <link rel="stylesheet" href="css/style.css?v=${v}">
    <link rel="stylesheet" href="css/guide.css?v=${v}">
${extraHead}</head>
<body><main>${body}</main></body>
</html>
`;

const tmpRoots = [];
// 建立副本：<tmp>/docs（不含試算表）與 <tmp>/tools；說明頁一律換成上面的虛構頁面
function makeCopy() {
    const root = mkdtempSync(join(tmpdir(), 'site-tools-test-'));
    tmpRoots.push(root);
    const docs = join(root, SITE);
    cpSync(SRC_DOCS, docs, { recursive: true, filter: (src) => !SPREADSHEET_RE.test(src) });
    mkdirSync(join(docs, 'templates'), { recursive: true });
    mkdirSync(join(root, 'tools'));
    for (const f of TOOL_FILES) cpSync(join(TOOLS_SRC, f), join(root, 'tools', f));
    writeFileSync(join(docs, 'guide.html'), guideHtml());
    writeFileSync(join(docs, 'css', 'guide.css'), '.guide-section { margin: 0; }\n');
    return root;
}

const readIn = (root, rel) => readFileSync(join(root, SITE, rel), 'utf8');
const writeIn = (root, rel, text) => writeFileSync(join(root, SITE, rel), text);
const editIn = (root, rel, fn) => {
    const before = readIn(root, rel);
    const after = fn(before);
    if (after === before) throw new Error(`${rel} 的測試改動沒有作用（比對不到要改的文字）`);
    writeIn(root, rel, after);
};
const runTool = (root, script, args = []) => {
    const r = spawnSync(process.execPath, [join(root, 'tools', script), ...args], { cwd: root, encoding: 'utf8', timeout: 60000 });
    return { status: r.status, out: `${r.stdout || ''}${r.stderr || ''}` };
};
// check-site 要真的跑完（印出總結）；中途當掉時「不應報錯」的情境會誤判為通過
const CHECK_DONE_RE = /檢查失敗：\d+ 個錯誤|✓ 檢查通過/;
const runCheck = (root, args = []) => {
    const r = runTool(root, 'check-site.mjs', args);
    if (!CHECK_DONE_RE.test(r.out)) throw new Error(`check-site 沒有跑完（結束碼 ${r.status}）：${r.out.trim().split(/\r?\n/).slice(0, 3).join(' / ').slice(0, 160)}`);
    return r;
};

// 每個情境：在新副本上做改動，執行 check-site，檢查輸出有（或沒有）指定的錯誤
function checkCase(title, mutate, pattern, shouldReport) {
    R.test(title, () => {
        const root = makeCopy();
        mutate(root);
        const { out } = runCheck(root);
        const lines = out.split(/\r?\n/).filter(l => l.startsWith('✗'));
        const hit = lines.filter(l => pattern.test(l));
        return [shouldReport ? hit.length > 0 : hit.length === 0, hit.length ? hit.join(' / ').slice(0, 200) : `沒有符合 ${pattern} 的錯誤`];
    });
}

try {
    R.check('docs/service-worker.js 有 CACHE_VERSION', !!VERSION, VERSION ? `v${VERSION}` : '找不到');

    // ---------- 基準：正確的說明頁不應出現下列任何錯誤 ----------
    const PAGE_ERRORS = /guide\.html|404\.html 預載|sitemap\.xml/;
    checkCase('基準副本（正確的說明頁）：沒有說明頁、preload、sitemap 相關的錯誤', () => {}, PAGE_ERRORS, false);
    checkCase('404.html 以 /excel-web-tool/ 絕對路徑 preload 已預快取的字型：不報錯', (root) =>
        editIn(root, '404.html', t => t.replace('</head>', '<link rel="preload" href="/excel-web-tool/fonts/NotoSansTC-Regular-ui.woff2" as="font" type="font/woff2" crossorigin>\n</head>')),
    /404\.html 預載/, false);

    // ---------- 第 1 節：?v= ----------
    checkCase('guide.html 的 ?v= 與 CACHE_VERSION 不同：報錯', (root) =>
        writeIn(root, 'guide.html', guideHtml('9.9')), /guide\.html 的 \?v= \(9\.9\)/, true);
    checkCase('guide.html 引用自家 CSS 沒帶 ?v=：報錯', (root) =>
        editIn(root, 'guide.html', t => t.replace(`css/guide.css?v=${VERSION}`, 'css/guide.css')), /guide\.html 的 css\/guide\.css 缺少 \?v=/, true);
    // ---------- 第 2 節：預快取 ----------
    checkCase('guide.html 引用的 ?v= 網址不在 STATIC_ASSETS：報錯', (root) =>
        editIn(root, 'service-worker.js', t => t.replace(`'./css/guide.css?v=${VERSION}'`, "'./css/guide.css'")),
    /guide\.html 引用 css\/guide\.css\?v=.*STATIC_ASSETS 沒有完全相同/, true);
    checkCase('STATIC_ASSETS 沒有 ./guide.html：報錯（離線打不開）', (root) =>
        editIn(root, 'service-worker.js', t => t.replace(/^[ \t]*'\.\/guide\.html',\r?\n/m, '')), /docs\/guide\.html 不在 service-worker\.js 的 STATIC_ASSETS/, true);
    checkCase('根目錄新增其他頁面但沒有預快取：報錯', (root) =>
        writeIn(root, 'extra.html', guideHtml()), /docs\/extra\.html 不在 service-worker\.js 的 STATIC_ASSETS/, true);
    checkCase('guide.html preload 的檔案不存在：報錯', (root) =>
        writeIn(root, 'guide.html', guideHtml(VERSION, { extraHead: '    <link rel="preload" href="fonts/Missing-ui.woff2" as="font" type="font/woff2" crossorigin>\n' })),
    /guide\.html 預載的檔案不存在：fonts\/Missing-ui\.woff2/, true);
    checkCase('guide.html preload 沒有預快取的檔案：報錯', (root) =>
        writeIn(root, 'guide.html', guideHtml(VERSION, { extraHead: '    <link rel="preload" href="images/og-image.jpg" as="image">\n' })),
    /guide\.html 預載的 images\/og-image\.jpg 不在 STATIC_ASSETS/, true);
    checkCase('404.html preload 的絕對路徑檔案不存在：報錯', (root) =>
        editIn(root, '404.html', t => t.replace('</head>', '<link rel="preload" href="/excel-web-tool/fonts/Missing-ui.woff2" as="font" type="font/woff2" crossorigin>\n</head>')),
    /404\.html 預載的檔案不存在：\/excel-web-tool\/fonts\/Missing-ui\.woff2/, true);
    checkCase('404.html preload 用子路徑以外的絕對路徑：報錯', (root) =>
        editIn(root, '404.html', t => t.replace('</head>', '<link rel="preload" href="/fonts/NotoSansTC-Regular-ui.woff2" as="font" type="font/woff2" crossorigin>\n</head>')),
    /404\.html 預載的 \/fonts\/NotoSansTC-Regular-ui\.woff2 是 \/excel-web-tool\/ 以外的絕對路徑/, true);
    // ---------- 第 3、4 節：既有檢查自動涵蓋新頁面 ----------
    checkCase('guide.html 引用外部資源：報錯', (root) =>
        writeIn(root, 'guide.html', guideHtml(VERSION, { extraHead: '    <link rel="stylesheet" href="https://example.com/x.css">\n' })), /guide\.html 引用了外部資源/, true);
    checkCase('guide.html 含 inline style 屬性：報錯', (root) =>
        writeIn(root, 'guide.html', guideHtml(VERSION, { body: '<p style="color:red">x</p>' })), /guide\.html 含 inline style 屬性/, true);
    checkCase("guide.html 的 CSP 放寬成 'unsafe-inline'：報錯", (root) =>
        writeIn(root, 'guide.html', guideHtml(VERSION, { csp: GUIDE_CSP.replace("style-src 'self'", "style-src 'self' 'unsafe-inline'") })),
    /guide\.html 的 CSP 與預期不同/, true);
    checkCase('guide.html 的 CSP 開放 script-src：報錯', (root) =>
        writeIn(root, 'guide.html', guideHtml(VERSION, { csp: GUIDE_CSP.replace("script-src 'none'", "script-src 'self'") })),
    /guide\.html 的 CSP 與預期不同/, true);
    // ---------- 第 7 節：字型用字 ----------
    checkCase('guide.html 出現字型沒有的字：報錯（字型檢查涵蓋新頁面）', (root) =>
        writeIn(root, 'guide.html', guideHtml(VERSION, { body: '<p>\u{2A6D6}</p>' })), /字型未涵蓋 \d+ 個介面用字：.*\u{2A6D6}/u, true);
    // ---------- 第 10 節：sitemap ----------
    checkCase('sitemap.xml 沒有列出 guide.html：報錯', (root) =>
        editIn(root, 'sitemap.xml', t => t.replace(/\s*<url>\s*<loc>[^<]*guide\.html<\/loc>[\s\S]*?<\/url>/, '')),
    /sitemap\.xml 沒有列出 guide\.html/, true);
    checkCase('sitemap.xml 列出不存在的頁面：報錯', (root) =>
        editIn(root, 'sitemap.xml', t => t.replace('</urlset>', '  <url><loc>https://rich0512.github.io/excel-web-tool/missing.html</loc></url>\n</urlset>')),
    /sitemap\.xml 列出的 \/excel-web-tool\/missing\.html 不是網站上存在的頁面/, true);

    // ---------- bump-version：每頁的 ?v= 與每個 lastmod ----------
    R.test('bump-version：service-worker.js、index.html、guide.html、404.html 的 ?v= 全部同步', () => {
        const root = makeCopy();
        const { status, out } = runTool(root, 'bump-version.mjs', ['77.7']);
        if (status !== 0) return [false, `結束碼 ${status}：${out.trim().slice(0, 160)}`];
        const stale = ['service-worker.js', 'index.html', 'guide.html', '404.html'].filter(f => {
            const vs = [...readIn(root, f).matchAll(/\?v=([\d.]+)/g)].map(m => m[1]);
            return !vs.length || vs.some(v => v !== '77.7');
        });
        const swOk = versionOf(readIn(root, 'service-worker.js')) === '77.7';
        return [!stale.length && swOk, stale.length ? `沒有同步：${stale.join(', ')}` : swOk ? '全部是 77.7' : 'CACHE_VERSION 沒有改'];
    });
    R.test('bump-version：sitemap.xml 的每個 <lastmod> 都改成今天', () => {
        const root = makeCopy();
        editIn(root, 'sitemap.xml', t => t.replace(/<lastmod>[^<]*<\/lastmod>/g, '<lastmod>2000-01-01</lastmod>'));
        const { status, out } = runTool(root, 'bump-version.mjs', ['77.7']);
        if (status !== 0) return [false, `結束碼 ${status}：${out.trim().slice(0, 160)}`];
        const today = new Date().toLocaleDateString('sv-SE');
        const dates = [...readIn(root, 'sitemap.xml').matchAll(/<lastmod>([^<]*)<\/lastmod>/g)].map(m => m[1]);
        return [dates.length >= 2 && dates.every(d => d === today), `lastmod：${dates.join(', ')}（今天 ${today}）`];
    });
    // ---------- 第 11 節：--require-bump 檢查每個 <lastmod> ----------
    // 副本裡 git init、git add 後用 git write-tree 當「上一版」（tree 物件就能給 git show／git diff 用，不必 commit）
    const gitIn = (root, args) => {
        const r = spawnSync('git', ['-c', 'core.autocrlf=false', ...args], { cwd: root, encoding: 'utf8' });
        if (r.status !== 0) throw new Error(`git ${args[0]} 失敗：${(r.stderr || '').trim().slice(0, 120)}`);
        return r.stdout.trim();
    };
    // 依序設定每個 lastmod（不足的沿用最後一個日期）；結果可能與原本相同，所以不用 editIn
    const setLastmods = (root, dates) => {
        let i = 0;
        writeIn(root, 'sitemap.xml', readIn(root, 'sitemap.xml').replace(/<lastmod>[^<]*<\/lastmod>/g, () => `<lastmod>${dates[Math.min(i++, dates.length - 1)]}</lastmod>`));
    };
    // 上一版：每個 lastmod 都是 2000-01-02；這一版：先 bump-version（改版本、全部改成今天），再依 dates 改 sitemap
    function requireBumpCase(title, dates, pattern, shouldReport) {
        R.test(title, () => {
            const root = makeCopy();
            setLastmods(root, ['2000-01-02']);
            gitIn(root, ['init', '-q']);
            gitIn(root, ['add', '-A']);
            const base = gitIn(root, ['write-tree']);
            const bump = runTool(root, 'bump-version.mjs', ['77.7']);
            if (bump.status !== 0) return [false, `bump-version 結束碼 ${bump.status}`];
            if (dates) setLastmods(root, dates);
            const { out } = runCheck(root, [`--require-bump=${base}`]);
            const hit = out.split(/\r?\n/).filter(l => l.startsWith('✗') && pattern.test(l));
            return [shouldReport ? hit.length > 0 : hit.length === 0, hit.length ? hit.join(' / ').slice(0, 200) : `沒有符合 ${pattern} 的錯誤`];
        });
    }
    const today = new Date().toLocaleDateString('sv-SE');
    requireBumpCase('--require-bump：每個 lastmod 都改成今天，不報 sitemap 錯誤', null, /sitemap|CACHE_VERSION/, false);
    requireBumpCase('--require-bump：只改了第一個 lastmod、說明頁仍是上一版日期：報錯', [today, '2000-01-02'], /sitemap\.xml 有 lastmod 仍為 2000-01-02/, true);
    requireBumpCase('--require-bump：說明頁的 lastmod 比上一版舊：報錯', [today, '2000-01-01'], /sitemap\.xml 的 lastmod（2000-01-01）比上一版（2000-01-02）還舊/, true);

    R.test('bump-version 之後 check-site 沒有版本或 sitemap 不一致的錯誤', () => {
        const root = makeCopy();
        runTool(root, 'bump-version.mjs', ['77.7']);
        const { out } = runCheck(root);
        const bad = out.split(/\r?\n/).filter(l => l.startsWith('✗') && /\?v=|CACHE_VERSION|sitemap/.test(l));
        return [bad.length === 0, bad.length ? bad.join(' / ').slice(0, 200) : '沒有'];
    });
} finally {
    for (const root of tmpRoots) {
        try { rmSync(root, { recursive: true, force: true }); } catch { /* 暫存資料夾刪不掉不影響結果 */ }
    }
    R.check('暫存副本都已刪除', tmpRoots.every(r => !existsSync(r)), `${tmpRoots.length} 個副本`);
}

R.done();
