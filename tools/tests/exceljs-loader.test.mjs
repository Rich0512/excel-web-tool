// ExcelJS 延遲載入（utils.js 的 loadExcelJS）：只建立一個 <script>、網址帶 utils.js 的 ?v=、失敗可重試、
// 閒置時預先載入、選檔與匯出前會等它載好；以及真正的 exceljs.min.js 在 CSP 沒有 'unsafe-eval' 時能運作。
import vm from 'node:vm';
import { readFileSync } from 'node:fs';
import { loadSite, makeReporter, makeFile, makeFakeExcelJS, readSiteFile, EXCELJS_PATH } from './harness.mjs';
const R = makeReporter('ExcelJS 延遲載入');

// ---- loadExcelJS 本身
{
    const s = loadSite({ withExcelJS: false });
    const has = s.run('typeof loadExcelJS') === 'function';
    R.check('utils.js 有 loadExcelJS()', has);
    if (has) {
        const p1 = s.run('loadExcelJS()');
        const p2 = s.run('loadExcelJS()');
        p1.catch(() => {});
        R.check('重複呼叫回傳同一個快取的 Promise', p1 === p2);
        const sc = s.scripts();
        R.check('只建立一個 <script>', sc.length === 1, `建立 ${sc.length} 個`);
        R.check('<script> 加在 document.head', sc[0]?.parentNode === s.document.head);
        R.check('載入前已宣告 regeneratorRuntime（舊瀏覽器免 eval）', s.run("'regeneratorRuntime' in window"));
        sc[0].dispatch('error');
        let err = null;
        try { await p1; } catch (e) { err = e; }
        R.check('載入失敗時 reject', !!err);
        R.check('失敗訊息說明下一步（檢查網路、再試一次）', err && /網路/.test(err.message) && /再試一次/.test(err.message), err?.message);
        const p3 = s.run('loadExcelJS()');
        R.check('失敗後再次呼叫會重新建立 <script>', p3 !== p1 && s.scripts().length === 2);
        s.ctx.ExcelJS = makeFakeExcelJS(s.ctx);
        s.scripts()[1].dispatch('load');
        const lib = await p3;
        R.check('載入成功時 resolve 為 window.ExcelJS', lib === s.ctx.ExcelJS);
        R.check('成功後再呼叫不再建立 <script>', (await s.run('loadExcelJS()')) === s.ctx.ExcelJS && s.scripts().length === 2);
    }
}
// ---- ExcelJS 網址跟著 utils.js 的 ?v=（快取版本一致）
for (const [version, expect] of [['3.14', 'js/libs/exceljs.min.js?v=3.14'], ['3.15', 'js/libs/exceljs.min.js?v=3.15'], [null, 'js/libs/exceljs.min.js']]) {
    const s = loadSite({ withExcelJS: false, version });
    R.test(`utils.js 以 ${version ? `?v=${version}` : '不帶 ?v='} 載入 → ExcelJS 網址是 ${expect}`, () => {
        s.run('loadExcelJS()').catch(() => {});
        const src = s.scripts().at(-1)?.src;
        return [src === expect, src];
    });
}
// ---- 已有 window.ExcelJS 時直接 resolve
{
    const s = loadSite({ withExcelJS: true });
    let ok = false, detail = 'loadExcelJS 不存在';
    if (s.run('typeof loadExcelJS') === 'function') {
        const lib = await s.run('loadExcelJS()');
        ok = lib === s.ctx.ExcelJS && s.scripts().length === 0;
        detail = `建立 ${s.scripts().length} 個 <script>`;
    }
    R.check('已有 window.ExcelJS 時直接 resolve，不建立 <script>', ok, detail);
}
// ---- 頁面 load 後閒置時預先載入；失敗不打擾使用者
{
    const s = loadSite({ withExcelJS: false });
    s.fireWindow('load');
    s.advance(10000);
    R.check('頁面 load 後閒置時預先載入 ExcelJS', s.scripts().length === 1, `建立 ${s.scripts().length} 個`);
    s.scripts()[0]?.dispatch('error');
    await s.settle();
    R.check('預先載入失敗不跳出 alert（使用者沒有操作）', s.alerts.length === 0, s.alerts.join(' / '));
}
// ---- 選擇檔案時等 ExcelJS 載好；失敗時顯示可行動的訊息，重試可成功
{
    const s = loadSite({ withExcelJS: false });
    s.ctx.__nextSheets = [{ name: '總表', rows: [['班級', '座號', '姓名', '社團'], ['301', '05', '測試生301_05', '桌遊社']] }];
    const p = s.run('handleFileSelect')(makeFile('115學期名冊.xlsx'));
    await s.settle();
    s.scripts().at(-1)?.dispatch('error');
    await p.catch(() => {});
    await s.settle();
    const msg = s.alerts.join(' / ');
    R.check('選檔時 ExcelJS 載入失敗：提示檢查網路', /網路/.test(msg), msg.slice(0, 120));
    R.check('選檔時 ExcelJS 載入失敗：不保留名冊狀態', s.run('uploadedWorkbook') === null);
    s.alerts.length = 0;
    const p2 = s.run('handleFileSelect')(makeFile('115學期名冊.xlsx'));
    await s.settle();
    s.ctx.ExcelJS = makeFakeExcelJS(s.ctx);
    s.scripts().at(-1)?.dispatch('load');
    await p2;
    await s.settle();
    R.check('重試且 ExcelJS 載入成功後可以讀取名冊', s.run('uploadedWorkbook') !== null, s.alerts.join(' / ').slice(0, 120));
}
// ---- 匯出前 await loadExcelJS
R.check('exporter.js 的匯出在使用 ExcelJS 前 await loadExcelJS()', /await\s+loadExcelJS\(\)/.test(readSiteFile('js/exporter.js')));

// ---- 真正的 exceljs.min.js：在禁止字串產生程式碼（等同 CSP 沒有 'unsafe-eval'）的環境
{
    const code = readFileSync(EXCELJS_PATH, 'utf8');
    const run = (predeclare) => {
        const ctx = { setTimeout, clearTimeout, console, TextDecoder, TextEncoder };
        ctx.window = ctx; ctx.self = ctx;
        const c = vm.createContext(ctx, { codeGeneration: { strings: false } });
        vm.runInContext('delete globalThis.globalThis', c); // 模擬沒有 globalThis 的舊瀏覽器
        if (predeclare) ctx.regeneratorRuntime = undefined;
        try { vm.runInContext(code, c); return { ok: typeof ctx.ExcelJS === 'object', lib: ctx.ExcelJS }; } catch (e) { return { ok: false, err: e.name }; }
    };
    const without = run(false);
    const withPre = run(true);
    R.check('（對照）沒有預先宣告 regeneratorRuntime 時，舊瀏覽器會需要 eval（EvalError）', !without.ok && without.err === 'EvalError', without.err || 'ok');
    R.check('預先宣告 regeneratorRuntime 後，不需要 eval 就能載入 ExcelJS', withPre.ok);
    if (withPre.ok) {
        try {
            const wb = new withPre.lib.Workbook();
            const ws = wb.addWorksheet('課程與學生清單');
            ws.addRow(['班級', '座號', '姓名']); ws.addRow(['301', '05', '測試生301_05']);
            const buf = await wb.xlsx.writeBuffer();
            const wb2 = new withPre.lib.Workbook();
            await wb2.xlsx.load(buf);
            R.check('不需要 eval 也能寫出並讀回 xlsx（比對工作表名稱）', wb2.worksheets.map(w => w.name).join() === '課程與學生清單');
        } catch (e) { R.check('不需要 eval 也能寫出並讀回 xlsx', false, e.name + ': ' + e.message.slice(0, 80)); }
    }
}
R.done();
