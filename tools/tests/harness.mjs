// ============================================
// 單元測試輔助程式：用 node:vm 載入 docs/js 的檔案，搭配假的 DOM、計時器與 ExcelJS
// 用法：import { loadSite, makeReporter } from './harness.mjs'; const site = loadSite();
// 測試資料一律是虛構的（測試生301_05、王小明…），不可讀取任何真實名冊。
//
// 環境變數 SITE_ROOT：只給「對照組驗證」用，改從另一個 repo 副本的 docs/ 載入
// （例如放了舊版 parser.js 的暫存副本，確認測試抓得到舊版的問題）。預設是本 repo。
// ============================================

import vm from 'node:vm';
import { readFileSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { ROOT, SITE } from '../site-config.mjs';

export const SITE_ROOT = process.env.SITE_ROOT ? resolve(process.env.SITE_ROOT) : ROOT;
export const JS_DIR = join(SITE_ROOT, SITE, 'js');
export const CSS_PATH = join(SITE_ROOT, SITE, 'css', 'style.css');
export const SW_PATH = join(SITE_ROOT, SITE, 'service-worker.js');
export const EXCELJS_PATH = join(JS_DIR, 'libs', 'exceljs.min.js');

// 讀 docs/ 底下的檔案（相對於 docs/，例如 'js/exporter.js'）
export const readSiteFile = (rel) => readFileSync(join(SITE_ROOT, SITE, rel), 'utf8');

class FakeClassList {
    constructor() { this.set = new Set(); }
    add(...c) { c.forEach(x => this.set.add(x)); }
    remove(...c) { c.forEach(x => this.set.delete(x)); }
    contains(c) { return this.set.has(c); }
    toggle(c, force) {
        const on = force === undefined ? !this.set.has(c) : !!force;
        if (on) this.set.add(c); else this.set.delete(c);
        return on;
    }
}

// 只支援 tag、#id、.class 的組合（例如 tr、.day-select、button.btn）；後代選擇器只看最後一段
function matchesSimple(el, sel) {
    sel = sel.trim();
    const m = sel.match(/^([a-zA-Z]*)((?:[#.][\w-]+)*)$/);
    if (!m) return false;
    if (m[1] && el.tagName !== m[1].toUpperCase()) return false;
    const parts = m[2].match(/[#.][\w-]+/g) || [];
    return parts.every(p => p[0] === '#' ? el.id === p.slice(1) : el.classList.contains(p.slice(1)));
}

export class FakeElement {
    constructor(tag, id = '') {
        this.tagName = String(tag).toUpperCase();
        this.id = id;
        this.value = '';
        this.dataset = {};
        this.style = {};
        this.classList = new FakeClassList();
        this.attributes = {};
        this.children = [];
        this.parentNode = null;
        this.listeners = {};
        this.hidden = false;
        this.disabled = false;
        this.checked = false;
        this.files = [];
        this._innerHTML = '';
        this._text = '';
    }
    set className(v) { this.classList = new FakeClassList(); String(v).split(/\s+/).filter(Boolean).forEach(c => this.classList.add(c)); }
    get className() { return [...this.classList.set].join(' '); }
    get innerHTML() { return this._innerHTML; }
    set innerHTML(v) { this._innerHTML = String(v); this._text = ''; this.children = []; }
    get textContent() { return this._text || this.children.map(c => c.textContent).join(''); }
    set textContent(v) { this._text = String(v); this._innerHTML = ''; this.children = []; }
    appendChild(c) { c.parentNode = this; this.children.push(c); this.ownerDocument?.__onAppend?.(c, this); return c; }
    removeChild(c) { this.children = this.children.filter(x => x !== c); c.parentNode = null; return c; }
    remove() { if (this.parentNode) this.parentNode.removeChild(this); }
    setAttribute(k, v) { this.attributes[k] = String(v); }
    getAttribute(k) { return k in this.attributes ? this.attributes[k] : null; }
    hasAttribute(k) { return k in this.attributes; }
    removeAttribute(k) { delete this.attributes[k]; }
    addEventListener(t, fn) { (this.listeners[t] ||= []).push(fn); }
    removeEventListener(t, fn) { this.listeners[t] = (this.listeners[t] || []).filter(f => f !== fn); }
    dispatch(type, props = {}) {
        const ev = { type, target: this, currentTarget: this, preventDefault() {}, stopPropagation() {}, ...props };
        (this.listeners[type] || []).slice().forEach(fn => fn.call(this, ev));
        return ev;
    }
    dispatchEvent(ev) { return this.dispatch(ev.type, ev); }
    // 記錄目前焦點（document.activeElement），讓測試檢查操作後焦點移到哪裡
    focus() { if (this.ownerDocument) this.ownerDocument.activeElement = this; }
    click() { if (!this.disabled) this.dispatch('click'); }
    closest() { return null; }
    descendants() { return this.children.flatMap(c => [c, ...c.descendants()]); }
    querySelectorAll(sel) {
        const alts = sel.split(',').map(s => s.trim().split(/\s+/).pop());
        return this.descendants().filter(el => alts.some(a => matchesSimple(el, a)));
    }
    querySelector(sel) { return this.querySelectorAll(sel)[0] || null; }
}

// 依 index.html 的預設值初始化的欄位
const INITIAL_VALUES = {
    'mode-select': 'single',
    'slot-mode-select': 'semester',
    'paste-slot-mode-select': 'semester',
    'paste-day-select': '週一',
    'paste-quota': '12',
    'paste-grade-start': '1',
    'paste-grade-end': '2',
    'paste-priority': 'desc',
};

export const DEFAULT_FILES = ['utils', 'exporter', 'parser', 'processor', 'mapping-ui', 'app'];

/**
 * 依 index.html 的順序載入 docs/js 的檔案，回傳操作假頁面用的物件。
 * - files：要載入的檔案（不含 .js）
 * - withExcelJS：是否預先放好假的 window.ExcelJS（false 時模擬延遲載入）
 * - version：模擬 index.html 以 js/<檔名>.js?v=<version> 載入；null 時不帶 ?v=
 * - crypto：頁面的 window.crypto。預設是 Node 內建的 Web Crypto（globalThis.crypto）；
 *   可傳入假的物件（例如自訂 getRandomValues）；傳 null 時頁面沒有 crypto（模擬不支援的瀏覽器）
 */
export function loadSite({ files = DEFAULT_FILES, withExcelJS = true, codeGeneration, version = '3.14', crypto = globalThis.crypto } = {}) {
    const registry = new Map();
    const alerts = [];
    const logs = [];
    const timers = [];
    let timerSeq = 0;
    let now = 0;
    const appended = [];

    const document = {
        activeElement: null,
        __onAppend(child, parent) { appended.push({ child, parent }); },
        getElementById(id) {
            if (!registry.has(id)) {
                const el = new FakeElement(id.startsWith('btn') ? 'button' : 'div', id);
                el.ownerDocument = document;
                if (id in INITIAL_VALUES) el.value = INITIAL_VALUES[id];
                if (id === 'include-freshmen' || id === 'paste-include-freshmen') el.checked = true;
                if (id === 'paste-text-area') el.tagName = 'TEXTAREA';
                registry.set(id, el);
            }
            return registry.get(id);
        },
        createElement(tag) { const el = new FakeElement(tag); el.ownerDocument = document; return el; },
        querySelectorAll(sel) { return [...registry.values()].flatMap(el => [el, ...el.descendants()]).filter(el => sel.split(',').some(s => matchesSimple(el, s.trim().split(/\s+/).pop()))); },
        querySelector(sel) { return document.querySelectorAll(sel)[0] || null; },
    };
    document.head = document.createElement('head');
    document.body = document.createElement('body');

    const storage = new Map();
    const localStorage = {
        getItem: k => (storage.has(k) ? storage.get(k) : null),
        setItem: (k, v) => storage.set(k, String(v)),
        removeItem: k => storage.delete(k),
        clear: () => storage.clear(),
    };

    const windowListeners = {};
    const ctx = {
        document, localStorage,
        // 頁面的 console 輸出只收集起來，不印到終端機
        console: {
            log: (...a) => logs.push(['log', a]), warn: (...a) => logs.push(['warn', a]),
            error: (...a) => logs.push(['error', a]), info: (...a) => logs.push(['info', a]),
        },
        alert: (m) => alerts.push(String(m)),
        confirm: () => true,
        setTimeout: (fn, ms = 0) => { const id = ++timerSeq; timers.push({ id, fn, at: now + ms }); return id; },
        clearTimeout: (id) => { const i = timers.findIndex(t => t.id === id); if (i !== -1) timers.splice(i, 1); },
        addEventListener: (t, fn) => (windowListeners[t] ||= []).push(fn),
        removeEventListener: () => {},
        Promise, Map, Set, Uint8Array, ArrayBuffer, Blob: class { constructor(p, o) { this.parts = p; this.type = o?.type; } },
        URL: Object.assign(class extends URL {}, { createObjectURL: () => 'blob:x', revokeObjectURL() {} }),
        FileReader: class {
            readAsArrayBuffer(file) {
                // 瀏覽器裡事件處理器丟出的例外只會進到 window.onerror，不會中斷程式；這裡記下來
                file.arrayBuffer().then(buf => {
                    try {
                        const r = this.onload && this.onload({ target: { result: buf } });
                        if (r && typeof r.catch === 'function') r.catch(e => logs.push(['uncaught', [e.message]]));
                    } catch (e) { logs.push(['uncaught', [e.message]]); }
                }, () => this.onerror && this.onerror());
            }
        },
    };
    ctx.window = ctx;
    ctx.self = ctx;
    ctx.globalThis = ctx;
    if (crypto) ctx.crypto = crypto;
    if (withExcelJS) ctx.ExcelJS = makeFakeExcelJS(ctx);

    const context = vm.createContext(ctx, codeGeneration ? { codeGeneration } : undefined);
    for (const f of files) {
        const code = readFileSync(join(JS_DIR, f + '.js'), 'utf8');
        document.currentScript = { src: `http://localhost:8000/excel-web-tool/js/${f}.js${version ? `?v=${version}` : ''}` };
        vm.runInContext(code, context, { filename: f + '.js' });
        document.currentScript = null;
    }

    return {
        ctx: context, document, alerts, logs, storage, appended, windowListeners,
        el: (id) => document.getElementById(id),
        run: (code) => vm.runInContext(code, context),
        // 取得頁面的全域函式；不存在時回傳會丟出例外的函式，讓個別檢查記成 FAIL 而不是整個檔案中斷
        fn(name) {
            const f = vm.runInContext(`typeof ${name} === 'function' ? ${name} : undefined`, context);
            return f || (() => { throw new Error(`頁面沒有 ${name}()`); });
        },
        // 推進假時鐘並執行到期的計時器
        advance(ms) {
            now += ms;
            for (;;) {
                timers.sort((a, b) => a.at - b.at || a.id - b.id);
                const t = timers[0];
                if (!t || t.at > now) break;
                timers.shift();
                t.fn();
            }
        },
        pendingTimers: () => timers.length,
        fireWindow(type, props = {}) { (windowListeners[type] || []).forEach(fn => fn({ type, preventDefault() {}, ...props })); },
        scripts: () => appended.map(a => a.child).filter(c => c.tagName === 'SCRIPT'),
        flush: () => new Promise(r => setImmediate(r)),
        async settle() { for (let i = 0; i < 20; i++) await new Promise(r => setImmediate(r)); },
    };
}

// 假 ExcelJS：Workbook.xlsx.load() 會載入 ctx.__nextSheets（[{ name, rows }]）
// ctx.__loadGate（Promise）可讓讀取停在中途，ctx.__loadError 可模擬讀取失敗
export function makeFakeExcelJS(ctx) {
    class Workbook {
        constructor() {
            this.sheets = [];
            this.xlsx = {
                load: async () => {
                    if (ctx.__loadGate) await ctx.__loadGate;
                    if (ctx.__loadError) throw ctx.__loadError;
                    this.sheets = (ctx.__nextSheets || []).map(makeSheet);
                },
                writeBuffer: async () => new ArrayBuffer(8),
            };
        }
        eachSheet(fn) { this.sheets.forEach((s, i) => fn(s, i + 1)); }
        addWorksheet(name) { const s = makeSheet({ name, rows: [] }); this.sheets.push(s); return s; }
    }
    return { Workbook };
}

export function makeSheet({ name, rows }) {
    const makeRow = (r = []) => ({
        cellCount: r.length,
        height: 0,
        getCell: (c) => ({ value: r[c - 1] ?? null }),
        eachCell(opts, fn) {
            if (typeof opts === 'function') fn = opts;
            for (let c = 1; c <= r.length; c++) fn({ value: r[c - 1] ?? null }, c);
        },
    });
    return {
        name, rows,
        get rowCount() { return rows.length; },
        getRow: (n) => makeRow(rows[n - 1]),
        eachRow(fn) { rows.forEach((r, i) => { if (r && r.some(v => v !== '' && v != null)) fn(makeRow(r), i + 1); }); },
        // 匯出用
        addRow(arr) { rows.push(arr); },
        getCell: (r, c) => ({ value: rows[r - 1]?.[c - 1] ?? null }),
        mergeCells() {},
        get columns() {
            const n = Math.max(0, ...rows.map(r => r.length));
            return Array.from({ length: n }, (_, c) => ({
                eachCell(opts, fn) { rows.forEach(r => fn({ value: r[c] ?? null })); },
            }));
        },
    };
}

export function makeWorkbook(sheets) {
    const ws = sheets.map(makeSheet);
    return { eachSheet: (fn) => ws.forEach((s, i) => fn(s, i + 1)) };
}

// 假 File：內容是 PK 檔頭（通過 checkExcelFileFormat）
export function makeFile(name, size = 1024) {
    const bytes = new Uint8Array(16); bytes.set([0x50, 0x4B, 0x03, 0x04]);
    return {
        name, size,
        slice: () => ({ arrayBuffer: async () => bytes.buffer.slice(0, 8) }),
        arrayBuffer: async () => bytes.buffer,
    };
}

// 上傳假名冊：sheets 是 [{ name, rows }]
export async function uploadSheets(site, sheets, fileName = '名冊.xlsx') {
    site.ctx.__nextSheets = sheets;
    site.el('file-input').value = 'selected-file'; // 模擬已選擇檔案
    await site.run('handleFileSelect')(makeFile(fileName));
    await site.settle();
}

// 在貼上框輸入文字並等防抖跑完
export function paste(site, text) {
    const area = site.el('paste-text-area');
    area.value = text;
    area.dispatch('input');
    site.advance(1000);
}

/**
 * 換掉頁面的 confirm()：記下每則訊息，回傳值可以隨時改。loadSite() 預設的 confirm 一律回傳 true 且不記錄；
 * 需要檢查訊息或模擬「取消」時才呼叫。
 * 回傳 { messages, answer }：messages 是依序跳出的訊息，answer 設成 false 時模擬按「取消」
 */
export function stubConfirm(site, answer = true) {
    const stub = { messages: [], answer };
    site.ctx.confirm = (m) => { stub.messages.push(String(m)); return stub.answer; };
    return stub;
}

/**
 * 檢查 guessClubAndDay(text) 的結果。club、day 傳 undefined 表示不檢查該項，null 表示必須推不出來。
 * 回傳 (說明, 貼上內容, 社團, 星期) => void
 */
export function guessChecker(R, site) {
    const guess = site.fn('guessClubAndDay');
    return (title, text, club, day) => R.test(title, () => {
        const r = guess(text);
        const ok = (club === undefined || r.club === club) && (day === undefined || r.day === day);
        return [ok, `社團=${JSON.stringify(r.club)}，星期=${JSON.stringify(r.day)}`];
    });
}

/**
 * 測試記錄。每個檢查印一行「PASS／FAIL  說明 — 細節」，最後印「結果：n/m 通過」。
 * tools/test-site.mjs 依這個格式統計，不要改動行首。
 * - check(name, ok, detail)：記錄一個檢查
 * - test(name, fn)：fn 回傳布林值或 [布林值, 細節]；丟出例外時記成 FAIL
 */
export function makeReporter(title) {
    const results = [];
    const check = (name, ok, detail = '') => results.push({ name, ok: !!ok, detail: String(detail ?? '') });
    return {
        check,
        test(name, fn) {
            try {
                const r = fn();
                if (Array.isArray(r)) check(name, r[0], r[1]);
                else check(name, r);
            } catch (e) {
                check(name, false, `例外：${e.message}`);
            }
        },
        done() {
            console.log(`== ${title} ==`);
            for (const r of results) console.log(`${r.ok ? 'PASS' : 'FAIL'}  ${r.name}${r.detail ? '  — ' + r.detail : ''}`);
            const failed = results.filter(r => !r.ok).length;
            console.log(`結果：${results.length - failed}/${results.length} 通過`);
            process.exitCode = failed || results.length === 0 ? 1 : 0;
        },
    };
}
