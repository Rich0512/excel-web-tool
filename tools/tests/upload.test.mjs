// 上傳名冊：大小上限、讀取中狀態（aria-busy、aria-disabled 保留焦點）、讀取途中清除資料、所有工作表都缺姓名欄。
// 名冊內容都是假的工作表（測試生301_05…），不讀取任何真實檔案。
import { readFileSync } from 'node:fs';
import { loadSite, makeReporter, makeFile, makeFakeExcelJS, uploadSheets, CSS_PATH } from './harness.mjs';
const R = makeReporter('上傳名冊：大小上限、讀取中狀態與競態');

const SHEETS = [{ name: '總表', rows: [['班級', '座號', '姓名', '社團'], ['301', '05', '測試生301_05', '桌遊社']] }];
const onMapping = (s) => s.el('step-mapping').classList.contains('active');

// ---- 大小上限
{
    const s = loadSite();
    s.ctx.__nextSheets = SHEETS;
    s.el('file-input').value = 'selected-file';
    await s.run('handleFileSelect')(makeFile('大檔.xlsx', 25 * 1024 * 1024));
    await s.settle();
    const msg = s.alerts.join(' / ');
    R.check('超過 20 MB 的檔案被拒絕', s.run('uploadedWorkbook') === null);
    R.check('拒絕訊息說明上限（20 MB）與下一步（重新選擇）', /20 MB/.test(msg) && /重新選擇/.test(msg), msg.slice(0, 140));
    R.check('拒絕後清空檔案選擇', s.el('file-input').value === '');
}
{
    const s = loadSite();
    s.ctx.__nextSheets = SHEETS;
    await s.run('handleFileSelect')(makeFile('名冊.xlsx', 20 * 1024 * 1024));
    await s.settle();
    R.check('剛好 20 MB 的檔案可以讀取', s.run('uploadedWorkbook') !== null, s.alerts.join(' / ').slice(0, 100));
}

// ---- 讀取中狀態：成功
{
    const s = loadSite();
    s.ctx.__nextSheets = SHEETS;
    let open; s.ctx.__loadGate = new Promise(r => { open = r; });
    const p = s.run('handleFileSelect')(makeFile('名冊.xlsx'));
    await s.settle();
    const zone = s.el('drop-zone'), btn = s.el('btn-pick-file'), status = s.el('upload-status');
    R.check('讀取中：上傳區 aria-busy="true"', zone.getAttribute('aria-busy') === 'true', String(zone.getAttribute('aria-busy')));
    R.check('讀取中：「選擇檔案」按鈕 aria-disabled="true"', btn.getAttribute('aria-disabled') === 'true', String(btn.getAttribute('aria-disabled')));
    R.check('讀取中：按鈕不用 disabled（保留鍵盤焦點）', btn.disabled !== true);
    R.check('讀取中：狀態文字顯示「正在讀取名冊」', /正在讀取名冊/.test(status.textContent), status.textContent);
    let opened = 0;
    s.el('file-input').addEventListener('click', () => opened++);
    s.el('drop-zone').dispatch('click', { target: btn });
    R.check('讀取中按「選擇檔案」不開啟選檔視窗', opened === 0);
    const before = s.alerts.length;
    s.el('drop-zone').dispatch('drop', { dataTransfer: { files: [makeFile('另一個.xlsx')] } });
    await s.settle();
    open(); await p; await s.settle();
    R.check('讀取中再拖入檔案會被忽略', s.alerts.length === before && s.run('originalFileName') === '名冊.xlsx', s.run('originalFileName'));
    R.check('成功後：aria-busy 還原', zone.getAttribute('aria-busy') !== 'true');
    R.check('成功後：移除 aria-disabled，按鈕可按', btn.getAttribute('aria-disabled') === null && btn.disabled === false);
    R.check('成功後：狀態文字清空', status.textContent === '', status.textContent);
}
// ---- 讀取中狀態：讀取後失敗（找不到工作表）也要還原，之後可以再選檔
{
    const s = loadSite();
    s.ctx.__nextSheets = [{ name: '說明', rows: [['x']] }];
    let open; s.ctx.__loadGate = new Promise(r => { open = r; });
    const p = s.run('handleFileSelect')(makeFile('名冊.xlsx'));
    await s.settle();
    const btn = s.el('btn-pick-file');
    let opened = 0;
    s.el('file-input').addEventListener('click', () => opened++);
    open(); await p; await s.settle();
    R.check('讀取後失敗：移除 aria-disabled', btn.getAttribute('aria-disabled') === null);
    s.el('drop-zone').dispatch('click', { target: btn });
    R.check('讀取後失敗：按「選擇檔案」可以開啟選檔視窗', opened === 1);
}
// ---- 讀取失敗（檔案損毀）也要還原
{
    const s = loadSite();
    s.ctx.__loadError = new Error('corrupt');
    await s.run('handleFileSelect')(makeFile('壞檔.xlsx'));
    await s.settle();
    const zone = s.el('drop-zone'), btn = s.el('btn-pick-file'), status = s.el('upload-status');
    R.check('檔案損毀：有 1 則錯誤說明', s.alerts.length === 1, s.alerts.join(' / ').slice(0, 80));
    R.check('檔案損毀：aria-busy 還原', zone.getAttribute('aria-busy') !== 'true');
    R.check('檔案損毀：按鈕恢復可按', btn.disabled === false && btn.getAttribute('aria-disabled') === null);
    R.check('檔案損毀：狀態文字清空', status.textContent === '');
}

// ---- 讀取途中按「清除這次的資料」：在兩個 await 點各測一次
for (const where of ['xlsx.load', 'loadExcelJS']) {
    const s = loadSite({ withExcelJS: where !== 'loadExcelJS' });
    s.ctx.__nextSheets = SHEETS;
    let open;
    if (where === 'xlsx.load') s.ctx.__loadGate = new Promise(r => { open = r; });
    const p = s.run('handleFileSelect')(makeFile('名冊.xlsx'));
    await s.settle();
    s.el('btn-hard-reset').dispatch('click'); // confirm() 在 harness 一律回傳 true
    const alertsAfterReset = s.alerts.length;
    R.check(`[等待 ${where} 時清除] 立刻解除讀取中狀態`, s.el('upload-status').textContent === '' && s.el('drop-zone').getAttribute('aria-busy') !== 'true', s.el('upload-status').textContent);
    if (where === 'xlsx.load') open();
    else {
        s.ctx.ExcelJS = makeFakeExcelJS(s.ctx);
        s.scripts().at(-1)?.dispatch('load');
    }
    await p; await s.settle();
    R.check(`[等待 ${where} 時清除] 讀完後不會把名冊設回去`, s.run('uploadedWorkbook') === null);
    R.check(`[等待 ${where} 時清除] 讀完後不會切到第 2 步`, !onMapping(s));
    R.check(`[等待 ${where} 時清除] 讀完後不再跳 alert`, s.alerts.length === alertsAfterReset, s.alerts.slice(alertsAfterReset).join(' / ').slice(0, 80));
    s.ctx.__loadGate = null;
    await s.run('handleFileSelect')(makeFile('名冊.xlsx'));
    await s.settle();
    R.check(`[等待 ${where} 時清除] 之後可以重新讀取名冊`, s.run('uploadedWorkbook') !== null && onMapping(s));
}

// ---- 所有工作表都缺姓名欄
{
    const s = loadSite();
    await uploadSheets(s, [
        { name: '社團-桌遊社', rows: [['班級', '座號'], ['301', '05']] },
        { name: '社團-直笛社', rows: [['班級', '座號'], ['302', '11']] },
    ]);
    const msg = s.alerts.join(' / ');
    R.check('所有工作表都缺姓名欄：不進入第 2 步', !onMapping(s));
    R.check('所有工作表都缺姓名欄：清掉名冊', s.run('uploadedWorkbook') === null);
    R.check('所有工作表都缺姓名欄：alert 列出工作表名稱與「姓名」', /社團-桌遊社/.test(msg) && /社團-直笛社/.test(msg) && /姓名/.test(msg), msg.slice(0, 160));
}
{
    const s = loadSite();
    await uploadSheets(s, [
        { name: '社團-桌遊社', rows: [['班級', '座號', '姓名'], ['301', '05', '測試生301_05']] },
        { name: '社團-直笛社', rows: [['班級', '座號'], ['302', '11']] },
    ]);
    R.check('只有部分工作表缺姓名欄：照常進入第 2 步、不跳 alert', onMapping(s) && s.alerts.length === 0, s.alerts.join(' / ').slice(0, 80));
}

// ---- style.css：aria-disabled 的按鈕外觀與 disabled 相同
{
    const css = readFileSync(CSS_PATH, 'utf8');
    const disabledRule = css.match(/([^{}]*)\{[^}]*\}/g).find(r => /\.btn:disabled/.test(r.split('{')[0])) || '';
    R.check('style.css：.btn[aria-disabled="true"] 與 .btn:disabled 寫在同一條規則', /\.btn\[aria-disabled="true"\]/.test(disabledRule.split('{')[0]), disabledRule.split('{')[0].trim());
    const hovers = css.match(/[^{}]*:hover:not\([^)]*\)[^{]*\{/g) || [];
    const ok = hovers.length > 0 && hovers.every(h => h.includes('[aria-disabled="true"]'));
    R.check('style.css：hover 規則同時排除 aria-disabled', ok, hovers.map(h => h.trim().replace(/\s+/g, ' ')).join(' | '));
}
R.done();
