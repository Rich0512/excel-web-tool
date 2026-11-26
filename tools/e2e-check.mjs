#!/usr/bin/env node
// ============================================
// 瀏覽器端對端驗證（選用）：以 Chrome DevTools Protocol 控制無頭 Chrome 或 Edge，實際操作本機預覽的網站
// 需要：Node.js 22 以上（內建 WebSocket）、已安裝 Chrome 或 Edge
// 用法：
//   1. 另開一個終端機啟動預覽：node tools/serve.mjs 8124
//   2. node tools/e2e-check.mjs http://localhost:8124/excel-web-tool/ [情境,情境|all] [--browser=chrome|edge] [--repo=路徑] [--out=結果.json]
//   例：node tools/e2e-check.mjs http://localhost:8124/excel-web-tool/
//       node tools/e2e-check.mjs http://localhost:8124/excel-web-tool/ load,upload --browser=edge
// 情境：load csp-noeval upload paste quota offline mobile slow slow404（對話框一律自動按確定並記錄）
//   slow、slow404 各要等約 10 秒的模擬延遲；全部情境約 1～2 分鐘
// 參數：--repo 是取得假資料範本的 repo 根目錄（預設為本 repo）；--out 另存 JSON 結果，路徑不可在 repo 內
// 隱私：
//   - 網址只接受 http://localhost（或 127.0.0.1），不會連到正式網站或任何外部主機
//   - 上傳的檔案只用 <repo>/docs/templates/ 內唯一的 .xlsx 假資料範本；貼上流程只用「測試生」假姓名
//   - 瀏覽器設定檔與下載檔放在系統暫存資料夾的短路徑 <tmp>/ewt-e2e/<短id>/，結束時刪除，不在 repo 內留下任何檔案
//     （路徑要短：Chrome 的 Cache Storage 在設定檔深處，總長超過 Windows 260 字元時 Service Worker 會安裝失敗）
//   - 下載的 Excel 只檢查大小與前 2 個位元組（PK），不讀內容；輸出不含姓名
// 結束碼：0 全部通過、1 有 FAIL、2 參數或環境錯誤
// ============================================

import { spawn, execFileSync } from 'node:child_process';
import { mkdirSync, readdirSync, readFileSync, existsSync, rmSync, writeFileSync, rmdirSync, openSync, readSync, closeSync, statSync } from 'node:fs';
import { join, resolve, dirname } from 'node:path';
import { tmpdir } from 'node:os';
import { ROOT, TEMPLATE_DIR as TEMPLATE_REL, isInside } from './site-config.mjs';

const PROGRAM_FILES = [process.env.ProgramFiles, process.env['ProgramFiles(x86)'], process.env.LOCALAPPDATA, 'C:\\Program Files', 'C:\\Program Files (x86)'].filter(Boolean);
const BROWSERS = {
    chrome: PROGRAM_FILES.map(p => join(p, 'Google', 'Chrome', 'Application', 'chrome.exe')),
    edge: PROGRAM_FILES.map(p => join(p, 'Microsoft', 'Edge', 'Application', 'msedge.exe')),
};
const ALL_SCENARIOS = ['load', 'csp-noeval', 'upload', 'paste', 'quota', 'offline', 'mobile', 'slow', 'slow404'];
const PASTE_TEXT = '班級\t座號\t姓名\n301\t05\t測試生301_05\n302\t11\t測試生302_11\n405\t23\t測試生405_23';
const USER_CLUB = '使用者社團';

// ---------- 參數 ----------
function parseArgs(argv) {
    const opts = { browser: 'chrome', repo: ROOT, out: null, scenarios: ALL_SCENARIOS };
    const pos = [];
    for (const a of argv) {
        if (a.startsWith('--browser=')) opts.browser = a.slice(10);
        else if (a.startsWith('--repo=')) opts.repo = a.slice(7);
        else if (a.startsWith('--out=')) opts.out = a.slice(6);
        else if (a === '-h' || a === '--help') opts.help = true;
        else pos.push(a);
    }
    opts.baseUrl = pos[0];
    if (pos[1] && pos[1] !== 'all') opts.scenarios = pos[1].split(',').map(s => s.trim()).filter(Boolean);
    return opts;
}

function die(msg) {
    console.error(msg);
    process.exit(2);
}

const opts = parseArgs(process.argv.slice(2));
if (opts.help || !opts.baseUrl) {
    die('用法：先執行 node tools/serve.mjs 8124，再執行\n  node tools/e2e-check.mjs http://localhost:8124/excel-web-tool/ [情境,…|all] [--browser=chrome|edge] [--repo=路徑] [--out=結果.json]\n情境：' + ALL_SCENARIOS.join(' '));
}
let BASE;
try {
    BASE = new URL(opts.baseUrl);
} catch {
    die('網址格式錯誤：' + opts.baseUrl);
}
if (!['localhost', '127.0.0.1'].includes(BASE.hostname) || BASE.protocol !== 'http:') die('只允許 http://localhost（或 127.0.0.1）的網址，請先執行 node tools/serve.mjs 8124');
if (!BASE.pathname.endsWith('/')) BASE.pathname += '/';
const BASE_URL = BASE.href;
for (const s of opts.scenarios) if (!ALL_SCENARIOS.includes(s)) die('不認得的情境：' + s);
const BROWSER_EXE = (BROWSERS[opts.browser] || []).find(p => existsSync(p));
if (!BROWSER_EXE) die(`找不到瀏覽器：${opts.browser}（這個檢查需要 Chrome 或 Edge；可用 --browser=edge 改用 Edge）`);
if (opts.out && isInside(ROOT, resolve(opts.out))) die('--out 的路徑不可在 repo 內（結果檔不可留在 repo）：' + opts.out);

// 範本：<repo>/docs/templates/ 下唯一的 .xlsx（不寫死檔名）
const TEMPLATE_DIR = resolve(opts.repo, TEMPLATE_REL);
const templates = existsSync(TEMPLATE_DIR) ? readdirSync(TEMPLATE_DIR).filter(f => /\.xlsx$/i.test(f)) : [];
if (templates.length !== 1) die(`${TEMPLATE_REL} 內應該剛好有 1 個 .xlsx，實際 ${templates.length} 個`);
const TEMPLATE_PATH = join(TEMPLATE_DIR, templates[0]);
if (!isInside(TEMPLATE_DIR, TEMPLATE_PATH)) die(`範本路徑不在 ${TEMPLATE_REL} 內`);

const RUN_ID = new Date().toISOString().replace(/[-:T]/g, '').slice(0, 14);
// 暫存資料夾：<tmp>/ewt-e2e/<短id>，每個情境再用 <短id>/<序號>/p（設定檔）與 /d（下載）
const TMP_PARENT = join(tmpdir(), 'ewt-e2e');
const RUN_DIR = join(TMP_PARENT, (process.pid % 1296).toString(36) + (Date.now() % 46656).toString(36));
if (isInside(ROOT, RUN_DIR)) die('系統暫存資料夾在 repo 內，無法執行：' + RUN_DIR);
mkdirSync(RUN_DIR, { recursive: true });

const sleep = (ms) => new Promise(r => setTimeout(r, ms));
const trunc = (s, n = 200) => { s = String(s ?? ''); return s.length > n ? s.slice(0, n) + '…' : s; };
const shortUrl = (u) => (u.startsWith(BASE_URL) ? u.slice(BASE_URL.length) || '(index)' : u);

// ---------- CDP 用戶端 ----------
class CDP {
    constructor(ws) {
        this.ws = ws;
        this.id = 0;
        this.pending = new Map();
        this.handlers = new Set();
        ws.addEventListener('message', (ev) => this._onMessage(ev));
        ws.addEventListener('close', () => {
            for (const p of this.pending.values()) p.reject(new Error('CDP 連線已關閉'));
            this.pending.clear();
        });
    }
    static connect(url) {
        return new Promise((res, rej) => {
            const ws = new WebSocket(url);
            ws.addEventListener('open', () => res(new CDP(ws)), { once: true });
            ws.addEventListener('error', () => rej(new Error('無法連到 DevTools WebSocket')), { once: true });
        });
    }
    _onMessage(ev) {
        const msg = JSON.parse(typeof ev.data === 'string' ? ev.data : Buffer.from(ev.data).toString('utf8'));
        if (msg.id !== undefined) {
            const p = this.pending.get(msg.id);
            if (!p) return;
            this.pending.delete(msg.id);
            if (msg.error) p.reject(new Error(`${p.method}: ${msg.error.message}`));
            else p.resolve(msg.result);
            return;
        }
        for (const h of [...this.handlers]) {
            try { h(msg); } catch (e) { console.error('事件處理錯誤:', e.message); }
        }
    }
    send(method, params = {}, sessionId) {
        const id = ++this.id;
        const msg = { id, method, params };
        if (sessionId) msg.sessionId = sessionId;
        return new Promise((resolve, reject) => {
            this.pending.set(id, { resolve, reject, method });
            this.ws.send(JSON.stringify(msg));
        });
    }
    on(fn) {
        this.handlers.add(fn);
        return () => this.handlers.delete(fn);
    }
    close() {
        try { this.ws.close(); } catch { /* 忽略 */ }
    }
}

// ---------- 瀏覽器 ----------
async function launchBrowser(workDir) {
    const profile = join(workDir, 'p');
    const downloads = join(workDir, 'd');
    mkdirSync(profile, { recursive: true });
    mkdirSync(downloads, { recursive: true });
    const args = [
        '--headless=new', '--remote-debugging-port=0', `--user-data-dir=${profile}`,
        '--no-first-run', '--no-default-browser-check', '--disable-extensions', '--disable-sync',
        '--disable-background-networking', '--disable-component-update', '--disable-default-apps',
        '--disable-domain-reliability', '--disable-client-side-phishing-detection', '--no-pings',
        '--metrics-recording-only', '--safebrowsing-disable-auto-update', '--disable-breakpad',
        '--disable-features=Translate,OptimizationHints,MediaRouter,AutofillServerCommunication,msSmartScreenProtection',
        '--window-size=1280,900', 'about:blank',
    ];
    const proc = spawn(BROWSER_EXE, args, { stdio: 'ignore' });
    const portFile = join(profile, 'DevToolsActivePort');
    const deadline = Date.now() + 20000;
    let wsPath = null;
    while (Date.now() < deadline) {
        if (existsSync(portFile)) {
            const lines = readFileSync(portFile, 'utf8').split(/\r?\n/).filter(Boolean);
            if (lines.length >= 2) { wsPath = `ws://127.0.0.1:${lines[0]}${lines[1]}`; break; }
        }
        await sleep(100);
    }
    if (!wsPath) { killProc(proc); throw new Error('瀏覽器沒有在 20 秒內開啟 DevTools 連接埠'); }
    const cdp = await CDP.connect(wsPath);
    const version = await cdp.send('Browser.getVersion');
    const ctx = { proc, cdp, profile, downloadsDir: downloads, downloads: [], version: version.product };
    cdp.on((m) => {
        if (m.sessionId) return;
        if (m.method === 'Browser.downloadWillBegin') {
            ctx.downloads.push({ guid: m.params.guid, suggestedFilename: m.params.suggestedFilename, state: 'begin' });
        } else if (m.method === 'Browser.downloadProgress') {
            const d = ctx.downloads.find(x => x.guid === m.params.guid);
            if (d) { d.state = m.params.state; d.totalBytes = m.params.totalBytes; d.receivedBytes = m.params.receivedBytes; }
        }
    });
    await cdp.send('Browser.setDownloadBehavior', { behavior: 'allowAndName', downloadPath: downloads, eventsEnabled: true });
    return ctx;
}

function killProc(proc) {
    try { execFileSync('taskkill', ['/PID', String(proc.pid), '/T', '/F'], { stdio: 'ignore' }); } catch { /* 已結束 */ }
}

async function closeBrowser(ctx) {
    if (!ctx) return;
    const exited = new Promise(r => { if (ctx.proc.exitCode !== null) r(); else ctx.proc.once('exit', r); });
    try { await Promise.race([ctx.cdp.send('Browser.close'), sleep(3000)]); } catch { /* 忽略 */ }
    ctx.cdp.close();
    await Promise.race([exited, sleep(5000)]);
    if (ctx.proc.exitCode === null) killProc(ctx.proc);
    await sleep(300);
}

async function removeDir(dir) {
    for (let i = 0; i < 10; i++) {
        try { rmSync(dir, { recursive: true, force: true }); return true; } catch { await sleep(500); }
    }
    return !existsSync(dir);
}

// ---------- 頁面 ----------
// 注入的腳本不受頁面 CSP 限制：記錄 securitypolicyviolation
const INJECT = `(() => {
  const list = [];
  try { Object.defineProperty(window, '__e2eCsp', { value: list }); } catch (e) {}
  document.addEventListener('securitypolicyviolation', (e) => {
    list.push({ directive: e.effectiveDirective || e.violatedDirective, blocked: String(e.blockedURI || '').slice(0, 120),
      source: String(e.sourceFile || '').split('/').pop().slice(0, 80), line: e.lineNumber });
  }, true);
})();`;

async function openPage(ctx, { mobile = false, patchCsp = false } = {}) {
    const { cdp } = ctx;
    const { targetId } = await cdp.send('Target.createTarget', { url: 'about:blank' });
    const { sessionId } = await cdp.send('Target.attachToTarget', { targetId, flatten: true });
    const s = (method, params) => cdp.send(method, params, sessionId);
    const page = {
        ctx, s, sessionId, targetId,
        exceptions: [], consoleErrors: [], consoleWarnings: [], logErrors: [], cspViolations: [], dialogs: [],
        requests: new Map(), inflight: new Set(), loadEvents: 0, dclEvents: 0, dclAt: null, cspPatch: null, navStart: 0, loadAt: null,
    };
    cdp.on((m) => {
        if (m.sessionId !== sessionId) return;
        const p = m.params;
        switch (m.method) {
            case 'Runtime.exceptionThrown': {
                const d = p.exceptionDetails;
                page.exceptions.push({ text: trunc((d.exception && d.exception.description ? d.exception.description.split('\n')[0] : d.text), 200), url: shortUrl(d.url || ''), line: d.lineNumber });
                break;
            }
            case 'Runtime.consoleAPICalled':
                if (process.env.E2E_DEBUG) console.log('[console]', p.type, p.args.map(a => a.value ?? a.description).join(' ').slice(0, 200));
                if (p.type === 'error' || p.type === 'assert') {
                    page.consoleErrors.push(trunc(p.args.map(a => a.value !== undefined ? String(a.value) : (a.description || a.type)).join(' '), 200));
                }
                else if (p.type === 'warning') {
                    page.consoleWarnings.push(trunc(p.args.map(a => a.value !== undefined ? String(a.value) : (a.description || a.type)).join(' '), 200));
                }
                break;
            case 'Log.entryAdded':
                if (p.entry.level === 'error' || /Content Security Policy/i.test(p.entry.text)) {
                    const e = { source: p.entry.source, text: trunc(p.entry.text, 240), url: shortUrl(p.entry.url || '') };
                    if (/Content Security Policy/i.test(p.entry.text)) page.cspViolations.push({ via: 'log', ...e });
                    else page.logErrors.push(e);
                }
                break;
            case 'Page.javascriptDialogOpening':
                page.dialogs.push({ type: p.type, message: trunc(p.message, 300), at: Date.now() });
                s('Page.handleJavaScriptDialog', { accept: true, promptText: '' }).catch(() => {});
                break;
            case 'Page.domContentEventFired':
                page.dclEvents++;
                page.dclAt = Date.now();
                break;
            case 'Page.loadEventFired':
                page.loadEvents++;
                page.loadAt = Date.now();
                break;
            case 'Network.requestWillBeSent': {
                const prev = page.requests.get(p.requestId);
                page.requests.set(p.requestId, { url: p.request.url, type: p.type, method: p.request.method, t: Date.now() - page.navStart, beforeLoad: page.loadAt === null, ...(prev ? { redirected: true } : {}) });
                page.inflight.add(p.requestId);
                break;
            }
            case 'Network.responseReceived': {
                const r = page.requests.get(p.requestId);
                if (r) Object.assign(r, { status: p.response.status, fromServiceWorker: !!p.response.fromServiceWorker, fromDiskCache: !!p.response.fromDiskCache, mime: p.response.mimeType });
                break;
            }
            case 'Network.loadingFinished': {
                const r = page.requests.get(p.requestId);
                if (r) r.encodedDataLength = p.encodedDataLength;
                page.inflight.delete(p.requestId);
                break;
            }
            case 'Network.loadingFailed': {
                const r = page.requests.get(p.requestId);
                if (r) { r.failed = p.errorText; r.canceled = !!p.canceled; r.blockedReason = p.blockedReason; }
                page.inflight.delete(p.requestId);
                break;
            }
            case 'Fetch.requestPaused':
                handleCspPatch(page, p).catch((e) => { page.cspPatch = { error: e.message }; s('Fetch.continueRequest', { requestId: p.requestId }).catch(() => {}); });
                break;
        }
    });
    await s('Page.enable');
    await s('Runtime.enable');
    await s('Log.enable');
    await s('Network.enable');
    await s('DOM.enable');
    await s('Page.addScriptToEvaluateOnNewDocument', { source: INJECT });
    if (mobile) {
        await s('Emulation.setDeviceMetricsOverride', { width: 375, height: 812, deviceScaleFactor: 2, mobile: true });
        await s('Emulation.setTouchEmulationEnabled', { enabled: true, maxTouchPoints: 5 }).catch(() => {});
    }
    if (patchCsp) {
        await s('Fetch.enable', { patterns: [{ urlPattern: '*', resourceType: 'Document', requestStage: 'Response' }] });
    }
    return page;
}

// 攔截 HTML：把 CSP meta 的 'unsafe-eval' 拿掉後回應
async function handleCspPatch(page, p) {
    const { s } = page;
    if (p.responseStatusCode !== 200 || !p.request.url.startsWith(BASE_URL)) {
        await s('Fetch.continueRequest', { requestId: p.requestId });
        return;
    }
    const body = await s('Fetch.getResponseBody', { requestId: p.requestId });
    const html = body.base64Encoded ? Buffer.from(body.body, 'base64').toString('utf8') : body.body;
    let found = false, had = false, newCsp = null;
    const patched = html.replace(/<meta[^>]*Content-Security-Policy[^>]*>/gi, (tag) => {
        found = true;
        if (/'unsafe-eval'/.test(tag)) had = true;
        const t = tag.replace(/\s*'unsafe-eval'/g, '');
        newCsp = (t.match(/content="([^"]*)"/i) || [])[1] || null;
        return t;
    });
    page.cspPatch = { url: shortUrl(p.request.url), metaFound: found, hadUnsafeEval: had, patched: had, newCsp };
    const headers = (p.responseHeaders || []).filter(h => !/^content-length$/i.test(h.name));
    await s('Fetch.fulfillRequest', { requestId: p.requestId, responseCode: 200, responseHeaders: headers, body: Buffer.from(patched, 'utf8').toString('base64') });
}

async function evaluate(page, expression) {
    const r = await page.s('Runtime.evaluate', { expression, awaitPromise: true, returnByValue: true });
    if (r.exceptionDetails) {
        throw new Error('頁面內執行失敗：' + trunc(r.exceptionDetails.exception?.description || r.exceptionDetails.text, 200));
    }
    return r.result.value;
}

async function harvestCsp(page) {
    try {
        const list = await evaluate(page, 'Array.isArray(window.__e2eCsp) ? window.__e2eCsp.slice() : []');
        for (const v of list) page.cspViolations.push({ via: 'event', ...v });
        await evaluate(page, 'Array.isArray(window.__e2eCsp) && (window.__e2eCsp.length = 0)');
    } catch { /* 頁面可能正在導覽 */ }
}

async function waitLoad(page, action, timeout = 30000) {
    const before = page.loadEvents;
    await harvestCsp(page);
    page.navStart = Date.now();
    page.loadAt = null;
    await action();
    const deadline = Date.now() + timeout;
    while (page.loadEvents === before) {
        if (Date.now() > deadline) throw new Error('等待頁面 load 事件逾時');
        await sleep(25);
    }
}

async function navigate(page, url) {
    await waitLoad(page, async () => {
        const r = await page.s('Page.navigate', { url });
        if (r.errorText) throw new Error('導覽失敗：' + r.errorText);
    });
}

const reload = (page) => waitLoad(page, () => page.s('Page.reload', { ignoreCache: false }));

async function waitIdle(page, quietMs = 1000, maxMs = 12000) {
    const end = Date.now() + maxMs;
    let quietSince = Date.now();
    while (Date.now() < end) {
        if (page.inflight.size > 0) quietSince = Date.now();
        else if (Date.now() - quietSince >= quietMs) return true;
        await sleep(50);
    }
    return false;
}

async function waitUntil(page, expression, timeout = 20000, abort = null) {
    const end = Date.now() + timeout;
    while (Date.now() < end) {
        if (await evaluate(page, `!!(${expression})`)) return true;
        if (abort && abort()) return false;
        await sleep(25);
    }
    return false;
}

const js = (v) => JSON.stringify(v);

async function realClick(page, selector) {
    const info = await evaluate(page, `(() => {
        const el = document.querySelector(${js(selector)});
        if (!el) return null;
        el.scrollIntoView({ block: 'center', inline: 'center', behavior: 'instant' });
        const r = el.getBoundingClientRect();
        const x = r.left + r.width / 2, y = r.top + r.height / 2;
        const h = document.elementFromPoint(x, y);
        return { x, y, w: r.width, h: r.height, hit: !!h && (h === el || el.contains(h)) };
    })()`);
    if (!info) throw new Error(`找不到元素 ${selector}`);
    if (!info.w || !info.h || !info.hit) {
        // 被遮住或不可見時退回程式點擊，並記錄下來
        await evaluate(page, `document.querySelector(${js(selector)}).click()`);
        return { fallback: true, visible: !!(info.w && info.h) };
    }
    const base = { x: info.x, y: info.y, button: 'left', clickCount: 1 };
    await page.s('Input.dispatchMouseEvent', { type: 'mouseMoved', x: info.x, y: info.y });
    await page.s('Input.dispatchMouseEvent', { type: 'mousePressed', ...base });
    await page.s('Input.dispatchMouseEvent', { type: 'mouseReleased', ...base });
    return { fallback: false };
}

function setValue(page, selector, value, events = ['input', 'change']) {
    return evaluate(page, `(() => {
        const el = document.querySelector(${js(selector)});
        if (!el) throw new Error('找不到 ' + ${js(selector)});
        el.value = ${js(value)};
        for (const t of ${js(events)}) el.dispatchEvent(new Event(t, { bubbles: true }));
        return el.value;
    })()`);
}

const getValue = (page, selector) => evaluate(page, `(document.querySelector(${js(selector)}) || {}).value`);

async function waitForDownload(ctx, since, timeout = 30000) {
    const end = Date.now() + timeout;
    while (Date.now() < end) {
        const d = ctx.downloads.slice(since).find(x => x.state === 'completed' || x.state === 'canceled');
        if (d) {
            if (d.state === 'canceled') return { ok: false, reason: '下載被取消' };
            const file = join(ctx.downloadsDir, d.guid);
            for (let i = 0; i < 20 && !existsSync(file); i++) await sleep(100);
            if (!existsSync(file)) return { ok: false, reason: '下載完成但找不到檔案' };
            // 只讀前 2 個位元組確認是 xlsx（zip 的 PK 檔頭），不讀內容
            const head = Buffer.alloc(2);
            const fd = openSync(file, 'r');
            try { readSync(fd, head, 0, 2, 0); } finally { closeSync(fd); }
            return { ok: true, suggestedFilename: d.suggestedFilename, size: statSync(file).size, magic: head.toString('latin1') };
        }
        await sleep(50);
    }
    return { ok: false, reason: `${timeout / 1000} 秒內沒有完成下載` };
}

// ---------- 結果 ----------
function newResult(name) {
    return { name, checks: [], metrics: {}, notes: [] };
}
const check = (res, name, pass, detail = '') => res.checks.push({ name, status: pass === null ? 'INFO' : (pass ? 'PASS' : 'FAIL'), detail: String(detail) });

function attachDiagnostics(res, page) {
    res.exceptions = page.exceptions;
    res.consoleErrors = page.consoleErrors;
    res.logErrors = page.logErrors;
    res.consoleWarnings = page.consoleWarnings;
    res.cspViolations = page.cspViolations;
    res.dialogs = page.dialogs.map(d => ({ type: d.type, message: d.message }));
}

function commonChecks(res, page, label = '') {
    check(res, `${label}沒有未捕捉例外`, page.exceptions.length === 0, page.exceptions.length ? page.exceptions.map(e => `${e.text} @${e.url}:${e.line}`).join(' | ') : '0 筆');
    check(res, `${label}沒有 console.error`, page.consoleErrors.length === 0, page.consoleErrors.length ? page.consoleErrors.join(' | ') : '0 筆');
    check(res, `${label}沒有 CSP 違規`, page.cspViolations.length === 0, page.cspViolations.length ? page.cspViolations.map(v => v.via === 'log' ? v.text : `${v.directive} ${v.blocked} @${v.source}:${v.line}`).join(' | ') : '0 筆');
}

// ---------- 共用流程：上傳範本 ----------
async function uploadFlow(page, res, { download = true, label = '' } = {}) {
    const { ctx } = page;
    const dialogsBefore = page.dialogs.length;
    const { root } = await page.s('DOM.getDocument', { depth: 0 });
    const { nodeId } = await page.s('DOM.querySelector', { nodeId: root.nodeId, selector: '#file-input' });
    if (!nodeId) throw new Error('找不到 #file-input');
    const t0 = performance.now();
    await page.s('DOM.setFileInputFiles', { files: [TEMPLATE_PATH], nodeId });
    const ok = await waitUntil(page, `document.getElementById('step-mapping')?.classList.contains('active')`, 30000, () => page.dialogs.length > dialogsBefore);
    const ms = Math.round(performance.now() - t0);
    res.metrics[`${label}uploadToMappingMs`] = ok ? ms : null;
    const dlgs = page.dialogs.slice(dialogsBefore);
    check(res, `${label}上傳範本後出現對照表（#step-mapping active）`, ok, ok ? `${ms} ms` : `失敗（${ms} ms）${dlgs.length ? '，對話框：' + dlgs.map(d => d.message).join(' / ') : ''}`);
    if (!ok) return false;

    const rowInfo = await evaluate(page, `(() => {
        const rows = document.querySelectorAll('#mapping-table-body tr').length;
        let fixed = 0;
        document.querySelectorAll('#mapping-table-body .day-select').forEach(sel => {
            if (sel.value === '請選擇') { sel.value = '週一'; sel.dispatchEvent(new Event('change', { bubbles: true })); fixed++; }
        });
        return { rows, fixed, mode: document.getElementById('mode-select').value };
    })()`);
    res.metrics[`${label}mappingRows`] = rowInfo.rows;
    res.metrics[`${label}daySelectSetToMonday`] = rowInfo.fixed;
    res.metrics[`${label}mode`] = rowInfo.mode;
    check(res, `${label}對照表有資料列`, rowInfo.rows > 0, `${rowInfo.rows} 列（模式 ${rowInfo.mode}，補選週一 ${rowInfo.fixed} 列）`);
    if (!download) return true;

    const since = ctx.downloads.length;
    const d0 = page.dialogs.length;
    const t1 = performance.now();
    await realClick(page, '#btn-process-mapping');
    const dl = await waitForDownload(ctx, since);
    res.metrics[`${label}processToDownloadMs`] = dl.ok ? Math.round(performance.now() - t1) : null;
    const dlDialogs = page.dialogs.slice(d0);
    check(res, `${label}按「開始轉換」後下載 .xlsx`, dl.ok && /\.xlsx$/i.test(dl.suggestedFilename || ''), dl.ok ? `${dl.size} bytes，${res.metrics[`${label}processToDownloadMs`]} ms` : dl.reason + (dlDialogs.length ? '，對話框：' + dlDialogs.map(d => d.message).join(' / ') : ''));
    if (dl.ok) check(res, `${label}下載檔檔頭為 PK`, dl.magic === 'PK', `檔頭 ${js(dl.magic)}`);
    const success = await waitUntil(page, `document.getElementById('step-success')?.classList.contains('active')`, 5000);
    check(res, `${label}#step-success 為 active`, success);
    return dl.ok;
}

function requestSummary(page) {
    const list = [...page.requests.values()].filter(r => !r.url.startsWith('data:'));
    return list.map(r => ({
        url: shortUrl(r.url), type: r.type, status: r.status ?? null, bytes: r.encodedDataLength ?? null,
        tMs: r.t, beforeLoad: r.beforeLoad, fromServiceWorker: r.fromServiceWorker || false, failed: r.failed || null,
    }));
}

// 載入首頁 → 等 Service Worker activated → 重新整理讓它控制頁面；失敗時回傳 null
async function prepareControlledPage(ctx, res) {
    const page = await openPage(ctx);
    await navigate(page, BASE_URL);
    await waitIdle(page, 800, 10000);
    const onlineTitle = await evaluate(page, 'document.title');
    const swState = await evaluate(page, `(async () => {
        if (!('serviceWorker' in navigator)) return 'unsupported';
        const reg = await Promise.race([navigator.serviceWorker.ready, new Promise(r => setTimeout(() => r(null), 30000))]);
        if (!reg) return 'timeout';
        const end = Date.now() + 30000;
        while (reg.active && reg.active.state !== 'activated' && Date.now() < end) await new Promise(r => setTimeout(r, 100));
        return reg.active ? reg.active.state : 'no-active';
    })()`);
    check(res, 'Service Worker 已 activated', swState === 'activated', swState);
    if (swState !== 'activated') { attachDiagnostics(res, page); return null; }
    await reload(page);
    await waitIdle(page, 800, 10000);
    const controlled = await evaluate(page, '!!navigator.serviceWorker.controller');
    check(res, '重新整理後頁面受 Service Worker 控制', controlled);
    return { page, onlineTitle };
}

// 高延遲（仍在線）：頁面與 Service Worker 都套用。SW 用 setAutoAttach 接上（重啟後的新 SW 也會自動套用），
// 已存在但沒接上的 SW 再用 attachToTarget 補上。回傳收集 SW 例外的陣列
async function applyLatency(ctx, page, res, latency) {
    const params = { offline: false, latency, downloadThroughput: -1, uploadThroughput: -1 };
    const applied = [];
    const swExceptions = [];
    const done = new Set();
    const emulate = async (sid, label) => {
        try { await ctx.cdp.send('Network.enable', {}, sid); } catch { /* 忽略 */ }
        try { await ctx.cdp.send('Network.emulateNetworkConditions', params, sid); applied.push(`${label}:emulateNetworkConditions`); } catch (e) { applied.push(`${label}:失敗 ${e.message}`); }
    };
    const onSw = async (sessionId, info) => {
        if (done.has(info.targetId) || info.type !== 'service_worker' || !info.url.startsWith(BASE_URL)) return;
        done.add(info.targetId);
        ctx.cdp.on((m) => {
            if (m.sessionId === sessionId && m.method === 'Runtime.exceptionThrown') {
                const d = m.params.exceptionDetails;
                swExceptions.push(trunc(d.exception?.description?.split('\n')[0] || d.text, 200));
            }
        });
        await ctx.cdp.send('Runtime.enable', {}, sessionId).catch(() => {});
        await emulate(sessionId, 'sw');
        await ctx.cdp.send('Runtime.runIfWaitingForDebugger', {}, sessionId).catch(() => {});
    };
    ctx.cdp.on((m) => {
        if (!m.sessionId && m.method === 'Target.attachedToTarget') onSw(m.params.sessionId, m.params.targetInfo).catch(() => {});
    });
    await emulate(page.sessionId, 'page');
    try {
        await ctx.cdp.send('Target.setAutoAttach', { autoAttach: true, waitForDebuggerOnStart: false, flatten: true, filter: [{ type: 'service_worker', exclude: false }, { exclude: true }] });
    } catch (e) { applied.push('setAutoAttach 失敗 ' + e.message); }
    await sleep(300);
    const { targetInfos } = await ctx.cdp.send('Target.getTargets');
    for (const t of targetInfos.filter(t => t.type === 'service_worker' && t.url.startsWith(BASE_URL) && !done.has(t.targetId))) {
        const { sessionId } = await ctx.cdp.send('Target.attachToTarget', { targetId: t.targetId, flatten: true });
        await onSw(sessionId, t);
    }
    res.metrics.latencyEmulation = { latencyMs: latency, serviceWorkerTargets: done.size, applied };
    check(res, '高延遲已套用到頁面與 Service Worker', done.size > 0 && applied.some(a => a.startsWith('sw:emulate')) && applied.some(a => a.startsWith('page:emulate')), applied.join('、'));
    return { swExceptions };
}

// ---------- 情境 ----------
const scenarios = {
    // 1. 首頁載入
    async load(ctx, res) {
        const page = await openPage(ctx);
        await navigate(page, BASE_URL);
        const loadMs = Date.now() - page.navStart;
        const idle = await waitIdle(page, 1500, 15000);
        await harvestCsp(page);
        const timing = await evaluate(page, `(() => {
            const n = performance.getEntriesByType('navigation')[0] || {};
            const fcp = performance.getEntriesByName('first-contentful-paint')[0];
            return { domContentLoaded: Math.round(n.domContentLoadedEventEnd || 0), load: Math.round(n.loadEventEnd || 0), fcp: fcp ? Math.round(fcp.startTime) : null,
                     title: document.title, exceljsGlobal: typeof window.ExcelJS };
        })()`);
        const reqs = requestSummary(page);
        res.metrics.timing = timing;
        res.metrics.loadEventMsFromNavigate = loadMs;
        res.metrics.networkIdle = idle;
        res.metrics.requests = reqs;
        const total = reqs.reduce((a, r) => a + (r.bytes || 0), 0);
        const beforeLoadBytes = reqs.filter(r => r.beforeLoad).reduce((a, r) => a + (r.bytes || 0), 0);
        res.metrics.totalBytes = total;
        res.metrics.bytesBeforeLoadEvent = beforeLoadBytes;
        check(res, '頁面標題正常', /社團課表/.test(timing.title), timing.title);
        commonChecks(res, page);
        const bad = reqs.filter(r => r.failed || (r.status && r.status >= 400));
        check(res, '首次載入沒有失敗的請求', bad.length === 0, bad.length ? bad.map(r => `${r.url} ${r.status ?? r.failed}`).join(', ') : `${reqs.length} 個請求`);
        const xl = reqs.filter(r => /exceljs/i.test(r.url));
        check(res, 'exceljs.min.js 在首次載入期間被請求（資訊）', null, xl.length ? xl.map(r => `${r.url}：${r.bytes} bytes，${r.tMs} ms，${r.beforeLoad ? 'load 事件前' : 'load 事件後'}`).join('; ') : '沒有請求');
        const fonts = reqs.filter(r => /\.woff2/i.test(r.url));
        check(res, '字型檔在首次載入期間被請求（資訊）', null, fonts.length ? fonts.map(r => `${r.url}：${r.bytes} bytes，${r.beforeLoad ? 'load 前' : 'load 後'}`).join('; ') : '沒有請求');
        check(res, '首次載入傳輸量（資訊）', null, `全部 ${total} bytes（load 事件前 ${beforeLoadBytes} bytes），${reqs.length} 個請求；DCL ${timing.domContentLoaded} ms、load ${timing.load} ms、FCP ${timing.fcp} ms`);
        attachDiagnostics(res, page);
    },

    // 2. CSP 對照：移除 'unsafe-eval' 後跑上傳流程
    async 'csp-noeval'(ctx, res) {
        const page = await openPage(ctx, { patchCsp: true });
        await navigate(page, BASE_URL);
        await waitIdle(page, 800, 10000);
        res.metrics.cspPatch = page.cspPatch;
        check(res, '攔截 index.html 並改寫 CSP（資訊）', null, page.cspPatch ? (page.cspPatch.error ? '錯誤：' + page.cspPatch.error : (page.cspPatch.hadUnsafeEval ? `已移除 'unsafe-eval'；新 CSP：${page.cspPatch.newCsp}` : `原本就沒有 'unsafe-eval'，變體等同原版；CSP：${page.cspPatch.newCsp}`)) : '沒有攔截到 HTML');
        await uploadFlow(page, res);
        await harvestCsp(page);
        commonChecks(res, page);
        attachDiagnostics(res, page);
    },

    // 3. 上傳範本並下載
    async upload(ctx, res) {
        const page = await openPage(ctx);
        await navigate(page, BASE_URL);
        await waitIdle(page, 800, 10000);
        await uploadFlow(page, res);
        await harvestCsp(page);
        commonChecks(res, page);
        attachDiagnostics(res, page);
    },

    // 4. 貼上流程
    async paste(ctx, res) {
        const page = await openPage(ctx);
        await navigate(page, BASE_URL);
        await waitIdle(page, 800, 10000);
        await realClick(page, '#tab-btn-paste');
        await sleep(150);
        await setValue(page, '#paste-club-name', USER_CLUB);
        await setValue(page, '#paste-text-area', PASTE_TEXT, ['input']);
        await sleep(600);
        const kept = await getValue(page, '#paste-club-name');
        check(res, '貼上名單後不覆蓋使用者已填的社團名稱', kept === USER_CLUB, kept === USER_CLUB ? '維持原值' : `被改成其他值（長度 ${String(kept).length}，含數字：${/\d/.test(kept)}）`);

        await setValue(page, '#paste-club-name', '', ['input']);
        await evaluate(page, `document.getElementById('paste-text-area').dispatchEvent(new Event('input', { bubbles: true }))`);
        await sleep(600);
        const auto = String(await getValue(page, '#paste-club-name') ?? '');
        const hasDigit = /\d/.test(auto), hasFake = auto.includes('測試生');
        res.metrics.autoFill = { empty: auto === '', length: auto.length, hasDigit, hasFakeName: hasFake };
        check(res, '社團名稱空白時，自動填入值不含資料列內容（數字或姓名）', !hasDigit && !hasFake, auto === '' ? '沒有自動填入' : `自動填入長度 ${auto.length}，含數字：${hasDigit}，含「測試生」：${hasFake}`);
        const preview = await evaluate(page, `(document.getElementById('paste-parse-preview') || {}).textContent || ''`);
        const m = preview.match(/讀到\s*(\d+)\s*位/);
        res.metrics.previewCount = m ? Number(m[1]) : null;
        check(res, '解析預覽讀到 3 位學生', m && Number(m[1]) === 3, m ? `${m[1]} 位` : '沒有預覽筆數');

        // 填回社團名稱，年級放寬到 1–6（假資料是 3、4 年級，預設 1–2 年級會全部不符資格）
        await setValue(page, '#paste-club-name', USER_CLUB);
        await setValue(page, '#paste-grade-start', '1');
        await setValue(page, '#paste-grade-end', '6');
        const d0 = page.dialogs.length;
        await realClick(page, '#btn-load-club');
        await sleep(400);
        const loaded = await evaluate(page, `(() => {
            const items = document.querySelectorAll('#loaded-clubs-list .loaded-club-item');
            const last = items[items.length - 1];
            return { count: items.length, name: last ? last.querySelector('.loaded-club-name')?.textContent : null,
                     total: document.getElementById('total-pasted-students')?.textContent ?? null };
        })()`);
        res.metrics.loadedClub = { count: loaded.count, nameIsUserClub: loaded.name === USER_CLUB, totalSelected: loaded.total };
        check(res, '按「加入這個社團」後列出 1 個社團，名稱為使用者輸入值', loaded.count === 1 && loaded.name === USER_CLUB, `社團 ${loaded.count} 個，名稱相符：${loaded.name === USER_CLUB}，正取合計 ${loaded.total}${page.dialogs.length > d0 ? '，對話框：' + page.dialogs.slice(d0).map(d => d.message).join(' / ') : ''}`);

        const since = ctx.downloads.length;
        const d1 = page.dialogs.length;
        const t1 = performance.now();
        await realClick(page, '#btn-process-paste');
        const dl = await waitForDownload(ctx, since);
        res.metrics.pasteProcessToDownloadMs = dl.ok ? Math.round(performance.now() - t1) : null;
        check(res, '按「下載課表與抽籤結果」後下載 .xlsx', dl.ok && /\.xlsx$/i.test(dl.suggestedFilename || ''), dl.ok ? `${dl.size} bytes，${res.metrics.pasteProcessToDownloadMs} ms` : dl.reason + (page.dialogs.length > d1 ? '，對話框：' + page.dialogs.slice(d1).map(d => d.message).join(' / ') : ''));
        if (dl.ok) check(res, '下載檔檔頭為 PK', dl.magic === 'PK', `檔頭 ${js(dl.magic)}`);
        await harvestCsp(page);
        commonChecks(res, page);
        attachDiagnostics(res, page);
    },

    // 5. 正取人數為負數
    async quota(ctx, res) {
        const page = await openPage(ctx);
        await navigate(page, BASE_URL);
        await waitIdle(page, 800, 10000);
        await realClick(page, '#tab-btn-paste');
        await sleep(150);
        await setValue(page, '#paste-text-area', PASTE_TEXT, ['input']); // 先貼名單，再填名稱（確認名稱不會被資料列覆蓋）
        await sleep(200);
        await setValue(page, '#paste-club-name', '正取測試社');
        await setValue(page, '#paste-grade-start', '1');
        await setValue(page, '#paste-grade-end', '6');
        await setValue(page, '#paste-quota', '-3');
        const before = await evaluate(page, `document.querySelectorAll('#loaded-clubs-list .loaded-club-item').length`);
        const d0 = page.dialogs.length;
        await realClick(page, '#btn-load-club');
        await sleep(500);
        const after = await evaluate(page, `(() => {
            const items = document.querySelectorAll('#loaded-clubs-list .loaded-club-item');
            const last = items[items.length - 1];
            return { count: items.length, meta: last ? (last.querySelector('.loaded-club-meta')?.textContent || '') : '' };
        })()`);
        const added = after.count > before;
        const qm = after.meta.match(/正取\s*(-?\d+)\s*／\s*(-?\d+)/);
        const shownQuota = qm ? Number(qm[2]) : null;
        const dlgs = page.dialogs.slice(d0);
        res.metrics.quota = { added, shownQuota, dialogShown: dlgs.length > 0 };
        check(res, '正取人數 -3 時沒有以負數加入社團', !added || (shownQuota !== null && shownQuota >= 1), added ? `已加入，顯示正取上限 ${shownQuota}` : '未加入');
        check(res, '正取人數 -3 時有提示訊息（alert）', dlgs.length > 0, dlgs.length ? dlgs.map(d => `${d.type}：${d.message}`).join(' / ') : '沒有對話框');
        await harvestCsp(page);
        commonChecks(res, page);
        attachDiagnostics(res, page);
    },

    // 6. 離線
    async offline(ctx, res) {
        const page = await openPage(ctx);
        await navigate(page, BASE_URL);
        await waitIdle(page, 800, 10000);
        const onlineTitle = await evaluate(page, 'document.title');
        const swState = await evaluate(page, `(async () => {
            if (!('serviceWorker' in navigator)) return 'unsupported';
            const reg = await Promise.race([navigator.serviceWorker.ready, new Promise(r => setTimeout(() => r(null), 30000))]);
            if (!reg) return 'timeout';
            const end = Date.now() + 30000;
            while (reg.active && reg.active.state !== 'activated' && Date.now() < end) await new Promise(r => setTimeout(r, 100));
            return reg.active ? reg.active.state : 'no-active';
        })()`);
        let swDiag = '';
        if (swState !== 'activated') {
            swDiag = await evaluate(page, `(async () => JSON.stringify((await navigator.serviceWorker.getRegistrations()).map(r => ({ scope: r.scope, installing: r.installing && r.installing.state, waiting: r.waiting && r.waiting.state, active: r.active && r.active.state }))))()`);
            swDiag = '；registrations=' + swDiag + '；console.warn=' + page.consoleWarnings.join(' | ');
        }
        check(res, 'Service Worker 已 activated', swState === 'activated', swState + swDiag);
        if (swState !== 'activated') { attachDiagnostics(res, page); return; }
        await reload(page);
        await waitIdle(page, 800, 10000);
        const controlled = await evaluate(page, '!!navigator.serviceWorker.controller');
        check(res, '重新整理後頁面受 Service Worker 控制', controlled);
        const cacheInfo = await evaluate(page, `(async () => { const names = await caches.keys(); let n = 0; for (const k of names) n += (await (await caches.open(k)).keys()).length; return { names, entries: n }; })()`);
        res.metrics.caches = cacheInfo;

        // 開啟離線：頁面與 Service Worker 都要設定（SW 的請求不經過頁面的網路模擬）
        const offlineParams = { offline: true, latency: 0, downloadThroughput: -1, uploadThroughput: -1 };
        const applied = [];
        const emulate = async (sid, label) => {
            try { await ctx.cdp.send('Network.enable', {}, sid); } catch { /* 有些目標已啟用 */ }
            try { await ctx.cdp.send('Network.emulateNetworkConditions', offlineParams, sid); applied.push(`${label}:emulateNetworkConditions`); return; } catch (e) { applied.push(`${label}:emulateNetworkConditions 失敗 ${e.message}`); }
            try { await ctx.cdp.send('Network.overrideNetworkState', offlineParams, sid); applied.push(`${label}:overrideNetworkState`); } catch (e) { applied.push(`${label}:overrideNetworkState 失敗 ${e.message}`); }
        };
        await emulate(page.sessionId, 'page');
        const { targetInfos } = await ctx.cdp.send('Target.getTargets');
        const sws = targetInfos.filter(t => t.type === 'service_worker' && t.url.startsWith(BASE_URL));
        for (const t of sws) {
            const { sessionId } = await ctx.cdp.send('Target.attachToTarget', { targetId: t.targetId, flatten: true });
            await emulate(sessionId, 'sw');
        }
        res.metrics.offlineEmulation = { serviceWorkerTargets: sws.length, applied };

        const exBefore = page.exceptions.length;
        await reload(page);
        await sleep(800);
        const off = await evaluate(page, `(async () => {
            let probe;
            try { probe = (await fetch('./__e2e_offline_probe_' + Date.now() + '.txt')).status; } catch (e) { probe = 'rejected:' + e.name; }
            return { title: document.title, onLine: navigator.onLine, controlled: !!navigator.serviceWorker.controller, hasFileInput: !!document.getElementById('file-input'), probe };
        })()`);
        res.metrics.offlinePage = off;
        const offlineEffective = off.probe === 503 || String(off.probe).startsWith('rejected');
        check(res, '離線模擬確實生效（探測請求不是伺服器的 404）', offlineEffective, `探測結果 ${off.probe}，navigator.onLine=${off.onLine}`);
        check(res, '離線重新整理後標題正常', off.title === onlineTitle && off.hasFileInput, off.title);
        const newEx = page.exceptions.slice(exBefore);
        check(res, '離線重新整理後沒有未捕捉例外', newEx.length === 0, newEx.length ? newEx.map(e => e.text).join(' | ') : '0 筆');
        await uploadFlow(page, res, { label: '離線：' });
        await harvestCsp(page);
        commonChecks(res, page, '整個情境：');
        attachDiagnostics(res, page);
    },

    // 9. 高延遲但在線：驗證 Service Worker 的導航逾時（service-worker.js 的 NAVIGATION_TIMEOUT_MS，目前 4000）
    async slow(ctx, res) {
        const LATENCY = 9000, PASS_LIMIT_MS = 5500;
        const prep = await prepareControlledPage(ctx, res);
        if (!prep) return;
        const { page, onlineTitle } = prep;
        const { swExceptions } = await applyLatency(ctx, page, res, LATENCY);

        // 量 reload → DOMContentLoaded / load / 標題出現
        const exBefore = page.exceptions.length;
        const dcl0 = page.dclEvents, load0 = page.loadEvents;
        await harvestCsp(page);
        page.navStart = Date.now();
        page.loadAt = null;
        const t0 = Date.now();
        await page.s('Page.reload', { ignoreCache: false });
        let dclMs = null, loadMs = null;
        const end = Date.now() + 40000;
        while (Date.now() < end && (dclMs === null || loadMs === null)) {
            if (dclMs === null && page.dclEvents > dcl0) dclMs = page.dclAt - t0;
            if (loadMs === null && page.loadEvents > load0) loadMs = page.loadAt - t0;
            await sleep(20);
        }
        let titleMs = null, title = '';
        if (dclMs !== null) {
            title = await evaluate(page, 'document.title');
            titleMs = title === onlineTitle ? dclMs : null; // 標題在 DOMContentLoaded 時已確定
        }
        const nav = dclMs !== null ? await evaluate(page, `(() => { const n = performance.getEntriesByType('navigation')[0] || {}; return { responseStart: Math.round(n.responseStart || 0), workerStart: Math.round(n.workerStart || 0), transferSize: n.transferSize }; })()`) : null;
        res.metrics.slow = { reloadToDomContentLoadedMs: dclMs, reloadToLoadMs: loadMs, titleOk: title === onlineTitle, navigationTiming: nav };
        check(res, `高延遲（${LATENCY} ms）下 ${PASS_LIMIT_MS / 1000} 秒內出現頁面（DOMContentLoaded）`, dclMs !== null && dclMs <= PASS_LIMIT_MS && title === onlineTitle,
            dclMs === null ? '40 秒內沒有 DOMContentLoaded' : `reload→DCL ${dclMs} ms、→load ${loadMs ?? '逾時'} ms，HTML responseStart ${nav?.responseStart} ms，標題${title === onlineTitle ? '正常' : '不符：' + title}`);

        // 確認 SW 的網路請求真的變慢：快取沒有的檔案會經過 SW 上網（結果為 404）
        const probe = await evaluate(page, `(async () => { const t = performance.now(); let s; try { s = (await fetch('./__e2e_slow_probe_' + Date.now() + '.txt')).status; } catch (e) { s = 'rejected:' + e.name; } return { status: s, ms: Math.round(performance.now() - t) }; })()`);
        res.metrics.slowProbe = probe;
        check(res, '延遲模擬確實作用在 Service Worker 的網路請求', probe.ms >= LATENCY * 0.8, `探測 ${probe.status}，${probe.ms} ms`);

        await sleep(1000);
        const newEx = page.exceptions.slice(exBefore);
        check(res, '高延遲重新整理後頁面沒有未捕捉例外', newEx.length === 0, newEx.length ? newEx.map(e => e.text).join(' | ') : '0 筆');
        res.swExceptions = swExceptions;
        check(res, 'Service Worker 沒有未捕捉例外', swExceptions.length === 0, swExceptions.length ? swExceptions.join(' | ') : '0 筆');
        await harvestCsp(page);
        commonChecks(res, page, '整個情境：');
        attachDiagnostics(res, page);
    },

    // 10. 高延遲時導航到不存在的多層路徑：Service Worker 不可拿首頁回應（相對路徑會錯、整頁破版）
    async slow404(ctx, res) {
        const LATENCY = 9000;
        const prep = await prepareControlledPage(ctx, res);
        if (!prep) return;
        const { page, onlineTitle } = prep;
        const { swExceptions } = await applyLatency(ctx, page, res, LATENCY);

        const exBefore = page.exceptions.length;
        const dcl0 = page.dclEvents;
        await harvestCsp(page);
        page.navStart = Date.now();
        page.loadAt = null;
        const t0 = Date.now();
        const nav = await page.s('Page.navigate', { url: BASE_URL + 'a/b/c' });
        if (nav.errorText) throw new Error('導覽失敗：' + nav.errorText);
        // 每 100 ms 記錄標題，看是否曾短暫出現首頁
        const titles = [];
        let dclMs = null, homeSeenMs = null, notFoundMs = null;
        const end = Date.now() + 25000;
        while (Date.now() < end) {
            if (dclMs === null && page.dclEvents > dcl0) dclMs = page.dclAt - t0;
            let title = null, path = null;
            try { [title, path] = await evaluate(page, '[document.title, location.pathname]'); } catch { /* 導覽中 */ }
            const ms = Date.now() - t0;
            if (title !== null && path && path.endsWith('/a/b/c')) {
                if (!titles.length || titles[titles.length - 1].title !== title) titles.push({ ms, title });
                if (title === onlineTitle && homeSeenMs === null) homeSeenMs = ms;
                if (/找不到頁面/.test(title) && notFoundMs === null) notFoundMs = ms;
            }
            if (notFoundMs !== null || (homeSeenMs !== null && dclMs !== null && ms > dclMs + 1500)) break;
            await sleep(100);
        }
        // 標題出現後再等 DOMContentLoaded 最多 5 秒（404 頁與破版的首頁都可能稍晚觸發）
        for (const stop = Date.now() + 5000; dclMs === null && Date.now() < stop;) {
            if (page.dclEvents > dcl0) dclMs = page.dclAt - t0;
            else await sleep(50);
        }
        await sleep(500);
        let finalTitle = '';
        try { finalTitle = await evaluate(page, 'document.title'); } catch { /* 忽略 */ }
        const status = [...page.requests.values()].filter(r => r.type === 'Document' && r.url.endsWith('/a/b/c')).map(r => r.status).pop() ?? null;
        res.metrics.slow404 = { reloadToDomContentLoadedMs: dclMs, homeTitleSeenMs: homeSeenMs, notFoundTitleMs: notFoundMs, finalTitleIsNotFound: /找不到頁面/.test(finalTitle), documentStatus: status, titleTimeline: titles };
        check(res, '高延遲下 a/b/c 最後顯示 404 頁（「找不到頁面」）', /找不到頁面/.test(finalTitle), `最後標題：${finalTitle || '(空)'}，HTTP ${status}`);
        check(res, '高延遲下 a/b/c 不會出現首頁', homeSeenMs === null, homeSeenMs === null ? '沒有出現首頁標題' : `${homeSeenMs} ms 出現首頁標題`);
        check(res, 'a/b/c 頁面出現時間（資訊）', null, `DOMContentLoaded ${dclMs === null ? '未觸發（標題出現後再等 5 秒）' : dclMs + ' ms'}；404 標題 ${notFoundMs === null ? '未出現' : notFoundMs + ' ms'}；首頁標題 ${homeSeenMs === null ? '未出現' : homeSeenMs + ' ms'}`);

        const newEx = page.exceptions.slice(exBefore);
        check(res, '導航後頁面沒有未捕捉例外', newEx.length === 0, newEx.length ? newEx.map(e => e.text).join(' | ') : '0 筆');
        res.swExceptions = swExceptions;
        check(res, 'Service Worker 沒有未捕捉例外', swExceptions.length === 0, swExceptions.length ? swExceptions.join(' | ') : '0 筆');
        await harvestCsp(page);
        commonChecks(res, page, '整個情境：');
        attachDiagnostics(res, page);
    },

    // 7. 手機寬度
    async mobile(ctx, res) {
        const page = await openPage(ctx, { mobile: true });
        await navigate(page, BASE_URL);
        await waitIdle(page, 800, 10000);
        const measure = () => evaluate(page, `({ scrollWidth: document.documentElement.scrollWidth, innerWidth: window.innerWidth,
            wide: [...document.querySelectorAll('body *')].filter(el => { const r = el.getBoundingClientRect(); return r.width > 0 && r.right > 375.5 && getComputedStyle(el).position !== 'fixed' && !el.closest('.table-responsive, .sample-scroll'); }).slice(0, 5).map(el => el.tagName.toLowerCase() + (el.id ? '#' + el.id : '') + (el.className && typeof el.className === 'string' ? '.' + el.className.trim().split(/\\s+/).join('.') : '')) })`);
        const home = await measure();
        res.metrics.viewportInnerWidth = home.innerWidth;
        check(res, '手機寬度：首頁 scrollWidth <= 375', home.scrollWidth <= 375, `scrollWidth ${home.scrollWidth}，innerWidth ${home.innerWidth}${home.wide.length ? '，超出元素：' + home.wide.join(', ') : ''}`);
        await realClick(page, '#tab-btn-paste');
        await sleep(200);
        const paste = await measure();
        check(res, '手機寬度：貼上頁籤 scrollWidth <= 375', paste.scrollWidth <= 375, `scrollWidth ${paste.scrollWidth}${paste.wide.length ? '，超出元素：' + paste.wide.join(', ') : ''}`);
        await realClick(page, '#tab-btn-file');
        await sleep(200);
        if (await uploadFlow(page, res, { download: false, label: '手機：' })) {
            const mapping = await measure();
            check(res, '手機寬度：對照表步驟 scrollWidth <= 375', mapping.scrollWidth <= 375, `scrollWidth ${mapping.scrollWidth}${mapping.wide.length ? '，超出元素：' + mapping.wide.join(', ') : ''}`);
        }
        await harvestCsp(page);
        commonChecks(res, page);
        attachDiagnostics(res, page);
    },
};

// ---------- 主程式 ----------
const results = { baseUrl: BASE_URL, browser: opts.browser, browserVersion: null, template: templates[0], startedAt: new Date().toISOString(), runId: RUN_ID, scenarios: {} };

// 中途按 Ctrl+C 時也關掉瀏覽器並刪除暫存資料夾
let currentCtx = null;
process.once('SIGINT', () => {
    if (currentCtx) killProc(currentCtx.proc);
    setTimeout(() => {
        try { rmSync(RUN_DIR, { recursive: true, force: true }); } catch { /* 可能仍被占用 */ }
        console.error(`\n已中斷。${existsSync(RUN_DIR) ? '暫存資料夾未刪除，請手動刪除：' + RUN_DIR : '已刪除暫存資料夾'}`);
        process.exit(2);
    }, 1000);
});

for (const name of opts.scenarios) {
    const res = newResult(name);
    results.scenarios[name] = res;
    const workDir = join(RUN_DIR, String(opts.scenarios.indexOf(name)));
    let ctx = null;
    const t0 = Date.now();
    process.stdout.write(`執行情境 ${name} …\n`);
    try {
        ctx = currentCtx = await launchBrowser(workDir);
        results.browserVersion = ctx.version;
        await scenarios[name](ctx, res);
    } catch (e) {
        check(res, '情境執行完成', false, e.message);
    } finally {
        res.durationMs = Date.now() - t0;
        if (ctx) res.downloadsSeen = ctx.downloads.length;
        await closeBrowser(ctx);
        currentCtx = null;
        // 刪除下載的 Excel 與瀏覽器設定檔
        const removed = await removeDir(workDir);
        if (!removed) res.notes.push('暫存資料夾刪除失敗：' + workDir);
    }
}
const cleaned = await removeDir(RUN_DIR);
try { rmdirSync(TMP_PARENT); } catch { /* 還有其他執行中的資料夾 */ }

results.finishedAt = new Date().toISOString();
const counts = { PASS: 0, FAIL: 0, INFO: 0 };
for (const r of Object.values(results.scenarios)) for (const c of r.checks) counts[c.status]++;
if (!cleaned) counts.FAIL++;
results.counts = counts;

// 只有指定 --out 才寫 JSON（路徑已確認不在 repo 內）
if (opts.out) {
    mkdirSync(dirname(resolve(opts.out)), { recursive: true });
    writeFileSync(resolve(opts.out), JSON.stringify(results, null, 2), 'utf8');
}

console.log(`\n=== E2E 結果：${BASE_URL}（${results.browserVersion}）===`);
for (const r of Object.values(results.scenarios)) {
    console.log(`\n[${r.name}]（${r.durationMs} ms）`);
    for (const c of r.checks) console.log(`  ${c.status.padEnd(4)} ${c.name}${c.detail ? '：' + c.detail : ''}`);
    if (r.dialogs && r.dialogs.length) console.log(`  對話框 ${r.dialogs.length} 個：` + r.dialogs.map(d => `${d.type}「${trunc(d.message, 80)}」`).join('；'));
    for (const n of r.notes) console.log('  注意：' + n);
}
if (!cleaned) console.log(`\nFAIL 暫存資料夾刪除失敗，請手動刪除：${RUN_DIR}`);
console.log(`\n合計 PASS ${counts.PASS}、FAIL ${counts.FAIL}、INFO ${counts.INFO}`);
if (opts.out) console.log('JSON：' + resolve(opts.out));
process.exit(counts.FAIL > 0 ? 1 : 0);
