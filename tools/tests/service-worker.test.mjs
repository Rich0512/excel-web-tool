// docs/service-worker.js：導航取得的首頁只在「和 SW 同一版」時才存進快取當離線備用頁，
// 避免新版 HTML 配上舊版快取的 JS／CSS；版本比對要完整比對，v3.1 不可把 ?v=3.14 當成同版。
// 用 node:vm 模擬 Service Worker 環境（self、caches、fetch），不需要瀏覽器。
import vm from 'node:vm';
import { readFileSync } from 'node:fs';
import { makeReporter, SW_PATH } from './harness.mjs';
import { versionOf } from '../site-config.mjs';
const R = makeReporter('Service Worker：離線備用頁只存同版首頁');

const SCOPE = 'http://localhost/excel-web-tool/';
const SOURCE = readFileSync(SW_PATH, 'utf8');
const VERSION = versionOf(SOURCE);

// 用指定的 SW 原始碼處理一次首頁導航，網路回應 htmlBody；回傳存進快取的網址
async function navigateOnce(code, htmlBody) {
    const store = new Map();
    const handlers = {};
    const cache = {
        put: async (req, res) => { store.set(new URL(typeof req === 'string' ? req : req.url, SCOPE + 'service-worker.js').href, await res.text()); },
        addAll: async () => {},
    };
    const ctx = {
        self: { addEventListener: (t, h) => { handlers[t] = h; }, registration: { scope: SCOPE }, skipWaiting() {}, clients: { claim() {} } },
        caches: { open: async () => cache, match: async () => undefined, keys: async () => [], delete: async () => true },
        fetch: async () => new Response(htmlBody, { status: 200, headers: { 'content-type': 'text/html' } }),
        console: { log() {}, warn() {}, error() {} }, URL, Response, Request, Promise, setTimeout, clearTimeout, RegExp,
    };
    vm.createContext(ctx);
    vm.runInContext(code, ctx, { filename: 'service-worker.js' });
    if (typeof handlers.fetch !== 'function') throw new Error('service-worker.js 沒有註冊 fetch 事件');
    const waits = [];
    let responded;
    handlers.fetch({
        request: new Request(SCOPE, { mode: 'same-origin', headers: { accept: 'text/html' } }),
        respondWith: p => { responded = p; },
        waitUntil: p => waits.push(p),
    });
    await Promise.all(waits);
    await responded;
    // 沒有放進 waitUntil 的 cache.put 也要等它做完，才看得到實際存了什麼
    for (let i = 0; i < 20; i++) await new Promise(r => setImmediate(r));
    return [...store.keys()];
}

R.check('service-worker.js 有 CACHE_VERSION', !!VERSION, VERSION ? `v${VERSION}` : '找不到');

const cases = [
    [`同版首頁（?v=${VERSION}）存進快取當備用頁`, SOURCE, `<script src="js/app.js?v=${VERSION}"></script>`, true],
    ['下一版首頁（?v=99.0）不存進快取', SOURCE, '<script src="js/app.js?v=99.0"></script>', false],
];
// 把版本改成 v3.1，確認不是用字首比對（3.1 是 3.14 的字首）
const v31 = SOURCE.replace(/CACHE_VERSION = 'v[\d.]+'/, "CACHE_VERSION = 'v3.1'");
cases.push(
    ['版本 v3.1 的 SW：同版首頁（?v=3.1）存進快取', v31, '<script src="js/app.js?v=3.1"></script>', true],
    ['版本 v3.1 的 SW：?v=3.14 的首頁不當成同版、不存進快取', v31, '<script src="js/app.js?v=3.14"></script>', false],
);
// 同版：要存成 index.html（離線時用它回應）；不同版：什麼都不存
const SHELL = SCOPE + 'index.html';
for (const [title, code, html, shouldStore] of cases) {
    try {
        const keys = await navigateOnce(code, html);
        const ok = shouldStore ? keys.includes(SHELL) : keys.length === 0;
        R.check(title, ok, keys.length ? `存入 ${keys.map(k => k.replace(SCOPE, '') || '(根目錄)').join(', ')}` : '沒有存入');
    } catch (e) {
        R.check(title, false, `例外：${e.message}`);
    }
}
R.done();
