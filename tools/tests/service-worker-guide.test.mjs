// docs/service-worker.js：使用說明頁（guide.html）的導航
// - 在線時用網路的回應；離線時用預快取的 ./guide.html 回應
// - 網路取得的說明頁不寫回快取（預快取版本才保證與這一版 SW 的 ?v= 相符）
// - 比對時忽略查詢參數與 #；首頁與說明頁以外的導航仍不攔截（交給瀏覽器取 404 頁）
// - 預快取清單要有 ./guide.html 與 ./css/guide.css?v=<版本>
// 用 node:vm 模擬 Service Worker 環境（self、caches、fetch），不需要瀏覽器，也不發出網路請求。
import vm from 'node:vm';
import { readFileSync } from 'node:fs';
import { makeReporter, SW_PATH } from './harness.mjs';
import { versionOf } from '../site-config.mjs';
const R = makeReporter('Service Worker：使用說明頁離線可用、不寫回快取');

const SCOPE = 'http://localhost/excel-web-tool/';
const SW_URL = SCOPE + 'service-worker.js';
const SOURCE = readFileSync(SW_PATH, 'utf8');
const VERSION = versionOf(SOURCE);

const PRECACHED_GUIDE = '<p>預快取的說明頁</p>';
const NETWORK_GUIDE = '<p>網路上的說明頁</p>';
const keyOf = (req) => new URL(typeof req === 'string' ? req : req.url, SW_URL).href;

/**
 * 載入 SW 原始碼並建立假的環境。
 * - online：false 時 fetch 一律失敗（模擬離線）
 * - networkStatus：在線時網路回應的狀態碼
 * - precached：預先放進快取的 { 網址: 內容 }（網址相對 SCOPE）
 */
function loadWorker({ online = true, networkStatus = 200, precached = { 'guide.html': PRECACHED_GUIDE } } = {}) {
    const store = new Map(Object.entries(precached).map(([k, v]) => [SCOPE + k, v]));
    const handlers = {};
    const puts = [];
    const fetched = [];
    const added = [];
    const cache = {
        put: async (req, res) => { puts.push(keyOf(req)); store.set(keyOf(req), await res.text()); },
        addAll: async (reqs) => { added.push(...reqs); },
    };
    const ctx = {
        self: { addEventListener: (t, h) => { handlers[t] = h; }, registration: { scope: SCOPE }, skipWaiting() {}, clients: { claim() {} } },
        caches: {
            open: async () => cache,
            match: async (req, opts = {}) => {
                const want = new URL(keyOf(req));
                if (opts.ignoreSearch) want.search = '';
                const hit = [...store.keys()].find(k => {
                    const u = new URL(k);
                    if (opts.ignoreSearch) u.search = '';
                    return u.href === want.href;
                });
                return hit === undefined ? undefined : new Response(store.get(hit), { status: 200, headers: { 'content-type': 'text/html' } });
            },
            keys: async () => [], delete: async () => true,
        },
        fetch: async (req) => {
            fetched.push(req);
            if (!online) throw new TypeError('Failed to fetch');
            return new Response(NETWORK_GUIDE, { status: networkStatus, headers: { 'content-type': 'text/html' } });
        },
        console: { log() {}, warn() {}, error() {} }, URL, Response, Promise, setTimeout, clearTimeout, RegExp,
        // 瀏覽器的 SW 裡，new Request('./x') 以 SW 的網址為基準；Node 的 Request 不接受相對網址
        Request: class extends Request {
            constructor(input, init) { super(typeof input === 'string' ? new URL(input, SW_URL).href : input, init); }
        },
    };
    vm.createContext(ctx);
    vm.runInContext(SOURCE, ctx, { filename: 'service-worker.js' });
    if (typeof handlers.fetch !== 'function') throw new Error('service-worker.js 沒有註冊 fetch 事件');
    return { handlers, puts, fetched, added, store };
}

// 送出一次導航；回傳 { intercepted, response, body, puts, fetched }
async function navigate(path, opts) {
    const w = loadWorker(opts);
    const waits = [];
    let responded;
    const request = new Request(SCOPE + path, { headers: { accept: 'text/html' } });
    w.handlers.fetch({
        request,
        respondWith: (p) => { responded = p; },
        waitUntil: (p) => waits.push(p),
    });
    await Promise.all(waits);
    const response = responded === undefined ? undefined : await responded;
    // 沒有放進 waitUntil 的 cache.put 也要等它做完，才看得到實際存了什麼
    for (let i = 0; i < 20; i++) await new Promise(r => setImmediate(r));
    const body = response && response.type !== 'error' ? await response.text() : undefined;
    return { intercepted: responded !== undefined, response, body, request, puts: w.puts, fetched: w.fetched };
}

const describe = (r) => r.intercepted
    ? `回應 ${r.response?.type === 'error' ? 'Response.error()' : JSON.stringify(r.body)}；寫入快取 ${r.puts.length} 筆`
    : '沒有攔截';

async function test(title, fn) {
    try {
        const r = await fn();
        if (Array.isArray(r)) R.check(title, r[0], r[1]);
        else R.check(title, r);
    } catch (e) {
        R.check(title, false, `例外：${e.message}`);
    }
}

R.check('service-worker.js 有 CACHE_VERSION', !!VERSION, VERSION ? `v${VERSION}` : '找不到');

// ---------- 預快取清單 ----------
await test('install 預快取 ./guide.html 與 ./css/guide.css?v=<版本>，且都以 cache: reload 取得', async () => {
    const w = loadWorker();
    const waits = [];
    w.handlers.install({ waitUntil: (p) => waits.push(p) });
    await Promise.all(waits);
    const urls = w.added.map(keyOf);
    const want = [SCOPE + 'guide.html', `${SCOPE}css/guide.css?v=${VERSION}`];
    const missing = want.filter(u => !urls.includes(u));
    const notReload = w.added.filter(r => want.includes(keyOf(r)) && r.cache !== 'reload').map(keyOf);
    return [!missing.length && !notReload.length, missing.length ? `缺少 ${missing.join(', ')}` : notReload.length ? `沒有 cache: reload：${notReload.join(', ')}` : `共 ${urls.length} 項`];
});

// ---------- 在線 ----------
await test('在線：說明頁導航用網路的回應', async () => {
    const r = await navigate('guide.html');
    return [r.intercepted && r.body === NETWORK_GUIDE, describe(r)];
});
await test('在線：網路取得的說明頁不寫入快取（快取的仍是預快取版本）', async () => {
    const r = await navigate('guide.html');
    return [r.intercepted && r.puts.length === 0, describe(r)];
});
await test('在線：網路用的是原本的導航請求（不改寫網址）', async () => {
    const r = await navigate('guide.html?from=home');
    const sent = r.fetched[0];
    return [r.fetched.length === 1 && (sent === r.request || keyOf(sent) === keyOf(r.request)), `fetch ${r.fetched.length} 次：${r.fetched.map(keyOf).join(', ')}`];
});
await test('在線但伺服器回 404：照網路的回應，不改用快取', async () => {
    const r = await navigate('guide.html', { networkStatus: 404 });
    return [r.intercepted && r.response?.status === 404 && r.body === NETWORK_GUIDE && r.puts.length === 0, `狀態 ${r.response?.status}；${describe(r)}`];
});

// ---------- 離線 ----------
await test('離線：說明頁導航回應預快取的 ./guide.html', async () => {
    const r = await navigate('guide.html', { online: false });
    return [r.intercepted && r.body === PRECACHED_GUIDE, describe(r)];
});
for (const path of ['guide.html?x=1#a', 'guide.html?utm_source=line', 'guide.html#step-2']) {
    await test(`離線：${path} 也比對成說明頁，回應預快取版本`, async () => {
        const r = await navigate(path, { online: false });
        return [r.intercepted && r.body === PRECACHED_GUIDE, describe(r)];
    });
}
await test('離線且快取沒有說明頁：回應 Response.error()（不改用首頁）', async () => {
    const r = await navigate('guide.html', { online: false, precached: { 'index.html': '<p>首頁</p>' } });
    return [r.intercepted && r.response?.type === 'error', describe(r)];
});

// ---------- 不攔截的導航 ----------
for (const path of ['guide.htm', 'guides.html', 'sub/guide.html', 'guide.html/x', 'not-found', 'foo/index.html']) {
    await test(`其他非首頁導航不攔截：${path}`, async () => {
        const r = await navigate(path, { online: false, precached: { 'guide.html': PRECACHED_GUIDE, 'index.html': '<p>首頁</p>' } });
        return [!r.intercepted, describe(r)];
    });
}
await test('子路徑以外的 /guide.html 不攔截', async () => {
    const w = loadWorker({ online: false });
    let responded = false;
    w.handlers.fetch({
        request: new Request('http://localhost/guide.html', { headers: { accept: 'text/html' } }),
        respondWith: () => { responded = true; },
        waitUntil: () => {},
    });
    return [!responded, responded ? '被攔截' : '沒有攔截'];
});

// 首頁的行為不變（離線時仍回應預快取的首頁）
await test('首頁導航不受影響：離線時回應預快取的 ./index.html', async () => {
    const r = await navigate('', { online: false, precached: { 'guide.html': PRECACHED_GUIDE, 'index.html': '<p>首頁</p>' } });
    return [r.intercepted && r.body === '<p>首頁</p>', describe(r)];
});

R.done();
