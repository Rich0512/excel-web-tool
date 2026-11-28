// ============================================
// 完成畫面的「再下載一次」：
// - index.html 在 #step-success 有 #btn-redownload（次要按鈕、type="button"、預設隱藏），說明文字提到「再下載一次」
// - 貼上名單模式下載後，記住這次匯出的參數（lastExport），按鈕出現；再下載一次用同一組參數匯出，
//   不重新抽籤、抽籤時間與次數和第一次相同
// - 上傳 Excel 模式也會記住並可以再下載一次
// - 處理中停用按鈕並顯示處理中文字，結束後還原；失敗時的訊息不含名冊內容，結果仍保留
// - 「整理另一份名冊」「清除這次的資料」會清掉 lastExport 並隱藏按鈕；沒有 lastExport 時按鈕隱藏
// - lastExport 只在記憶體：不寫進 localStorage、不印到 console
// 測試資料全部是虛構的（測試生301_05 等）。
// ============================================

import { loadSite, makeReporter, readSiteFile, stubConfirm, uploadSheets } from './harness.mjs';

const R = makeReporter('完成畫面的再下載一次');

const leaks = (msgs) => msgs.some(m => /測試生/.test(m));
const safe = (s, code) => { try { return s.run(code); } catch (e) { return undefined; } };

// .btn 在 style.css 設了 display: inline-flex，會蓋過瀏覽器內建的 [hidden]；
// style.css 要有讓 .btn[hidden] 隱藏的規則，setLastExport() 只設 hidden 屬性、不設 style.display
const CSS = readSiteFile('css/style.css');
const cssHidesHiddenBtn = /(^|[\s,}])\.btn\[hidden\]\s*\{[^}]*display:\s*none/m.test(CSS);
const isHidden = (el) => el.hidden === true;
const isShown = (el) => el.hidden === false && el.style.display !== 'none';
const btnState = (el) => `hidden=${el.hidden}，display=${JSON.stringify(el.style.display)}`;

function addClub(s, text, { quota = '1', start = '1', end = '6', priority = 'desc', name = '桌遊社', day = '週二' } = {}) {
    s.el('paste-club-name').value = name;
    s.el('paste-day-select').value = day;
    s.el('paste-text-area').value = text;
    s.el('paste-quota').value = quota;
    s.el('paste-grade-start').value = start;
    s.el('paste-grade-end').value = end;
    s.el('paste-priority').value = priority;
    s.el('btn-load-club').dispatch('click');
    s.advance(1000);
}

// 換掉匯出與抽籤：記下每次匯出的參數（序列化）與抽籤次數。
// __exportGate 可讓匯出停在中途，__exportError 可模擬匯出失敗
function instrument(s) {
    s.ctx.__exports = [];
    s.ctx.__lotteryCalls = 0;
    s.ctx.__exportGate = null;
    s.ctx.__exportError = null;
    s.run(`exportWeeklySchedule = async (resultData, activeDays, slotMode, name, clubs = null, { showSuccess = true } = {}) => {
        if (__exportGate) await __exportGate;
        if (__exportError) throw __exportError;
        __exports.push(JSON.stringify([resultData, [...activeDays], slotMode, name, clubs]));
        if (showSuccess) switchStep('step-success');
    }`);
    s.run('const __origLottery = runClubLottery; runClubLottery = function (...a) { __lotteryCalls++; return __origLottery.apply(this, a); }');
    return { exports: () => s.ctx.__exports, lottery: () => s.ctx.__lotteryCalls };
}

const hasLastExport = (s) => safe(s, 'typeof lastExport !== "undefined" && lastExport !== null') === true;
const lastExportDefined = (s) => safe(s, 'typeof lastExport !== "undefined"') === true;

const ROSTER = '班級\t座號\t姓名\n301\t05\t測試生301_05\n302\t06\t測試生302_06\n401\t07\t測試生401_07';

// ---- 1. index.html 的按鈕與說明文字
{
    const html = readSiteFile('index.html');
    const panel = (html.match(/<div class="step-panel" id="step-success">[\s\S]*?<\/main>/) || [''])[0];
    const btn = (panel.match(/<button[^>]*id="btn-redownload"[^>]*>([^<]*)<\/button>/) || ['', '']);
    R.check('#step-success 有 #btn-redownload「再下載一次」', btn[1] === '再下載一次', btn[0] || '（找不到）');
    R.check('#btn-redownload 是 type="button" 的次要按鈕（btn btn-secondary）',
        /type="button"/.test(btn[0]) && /class="btn btn-secondary"/.test(btn[0]), btn[0]);
    R.check('#btn-redownload 預設有 hidden 屬性、沒有 style= 或 on*=',
        /\shidden(\s|>|=)/.test(btn[0]) && !/\sstyle=|\son\w+=/.test(btn[0]), btn[0]);
    R.check('「整理另一份名冊」仍在完成畫面', /id="btn-restart"[^>]*>整理另一份名冊</.test(panel));
    const desc = (panel.match(/<p class="success-desc">([\s\S]*?)<\/p>/) || ['', ''])[1].replace(/<br>|\s+/g, '');
    R.check('完成畫面說明：沒看到下載先檢查下載提示或資料夾，還是沒有就按「再下載一次」，用完刪除',
        desc.includes('請檢查瀏覽器的下載提示或下載資料夾') && desc.includes('「再下載一次」') && desc.includes('用完請刪除下載資料夾裡的課表Excel'), desc);
}

// ---- 2. 剛開啟網頁：沒有 lastExport，按鈕隱藏；按了也不會匯出
{
    const s = loadSite();
    const t = instrument(s);
    R.check('剛開啟網頁：lastExport 存在且是 null', lastExportDefined(s) && !hasLastExport(s), `typeof=${safe(s, 'typeof lastExport')}`);
    R.check('剛開啟網頁：「再下載一次」隱藏', isHidden(s.el('btn-redownload')), btnState(s.el('btn-redownload')));
    s.el('btn-redownload').click();
    await s.settle();
    R.check('沒有 lastExport 時按「再下載一次」：不匯出、不跳訊息', t.exports().length === 0 && s.alerts.length === 0, s.alerts.join(' / '));
}

// ---- 3. 貼上名單模式：下載後記住，再下載一次不重抽、內容相同
{
    const s = loadSite();
    addClub(s, ROSTER, { quota: '1', name: '桌遊社' });
    const t = instrument(s);
    const drawBefore = safe(s, 'JSON.stringify([pastedClubs[0].drawnAt, pastedClubs[0].drawRound, pastedClubs[0].students.map(x => [x.name, x.selected, x.drawSequence])])');
    s.el('paste-excel-filename').value = '115上學期';
    s.alerts.length = 0;
    await s.run('processPastedData()');
    R.check('貼上模式下載：匯出一次、已加入的社團清空（原本的行為）', t.exports().length === 1 && s.run('pastedClubs.length') === 0, `匯出 ${t.exports().length} 次`);
    R.check('貼上模式下載後：lastExport 保留這次的社團', hasLastExport(s) && safe(s, 'Array.isArray(lastExport.clubs) && lastExport.clubs.length') === 1,
        `clubs=${safe(s, 'lastExport && lastExport.clubs && lastExport.clubs.length')}`);
    R.check('貼上模式下載後：「再下載一次」出現', isShown(s.el('btn-redownload')), btnState(s.el('btn-redownload')));

    s.alerts.length = 0;
    s.el('btn-redownload').click();
    await s.settle();
    const ex = t.exports();
    R.check('按「再下載一次」：再匯出一次、沒有跳出訊息', ex.length === 2 && s.alerts.length === 0, `匯出 ${ex.length} 次 ${s.alerts.join(' / ')}`);
    R.check('按「再下載一次」：匯出的課表、星期、時段模式、檔名、社團資料與第一次完全相同', ex.length === 2 && ex[1] === ex[0]);
    R.check('按「再下載一次」：檔名仍是第一次的自訂檔名', ex.length === 2 && JSON.parse(ex[1])[3] === '115上學期', ex[1] ? JSON.parse(ex[1])[3] : '');
    const club = ex.length === 2 ? JSON.parse(ex[1])[4]?.[0] : null;
    R.check('按「再下載一次」：沒有重新抽籤，抽籤時間、次數、籤序與第一次相同',
        t.lottery() === 0 && club && JSON.stringify([club.drawnAt, club.drawRound, club.students.map(x => [x.name, x.selected, x.drawSequence])]) === drawBefore,
        `抽籤 ${t.lottery()} 次`);
    R.check('按「再下載一次」：抽籤次數 clubDrawCounts 不變', s.run('clubDrawCounts.get(clubDrawKey("桌遊社"))') === 1);
    R.check('按「再下載一次」：按鈕還原、仍可再按', !s.el('btn-redownload').disabled && isShown(s.el('btn-redownload')) && hasLastExport(s));
    s.el('btn-redownload').click();
    await s.settle();
    R.check('再按一次「再下載一次」：仍匯出同樣的內容', t.exports().length === 3 && t.exports()[2] === ex[0] && t.lottery() === 0);

    R.check('lastExport 不寫入 localStorage', ![...s.storage.values()].some(v => /測試生|桌遊社/.test(v)), [...s.storage.keys()].join(','));
    R.check('lastExport 的內容不印到 console', !s.logs.some(([, a]) => a.some(x => /測試生/.test(typeof x === 'string' ? x : JSON.stringify(x) || ''))));
}

// ---- 4. 處理中停用按鈕並顯示處理中文字，結束後還原
{
    const s = loadSite();
    addClub(s, ROSTER);
    const t = instrument(s);
    await s.run('processPastedData()');
    const btn = s.el('btn-redownload');
    const label = btn.textContent;
    let release;
    s.ctx.__exportGate = new Promise(r => { release = r; });
    btn.click();
    await s.settle();
    const midText = btn.textContent;
    R.check('再下載一次處理中：按鈕停用、顯示處理中文字', btn.disabled === true && midText !== label && /正在/.test(midText), `disabled=${btn.disabled}，文字=${midText}`);
    btn.click(); // 停用時再按不會重複匯出（FakeElement.click 會略過 disabled）
    s.ctx.__exportGate = null;
    release();
    await s.settle();
    R.check('再下載一次完成：按鈕還原文字與可按狀態', btn.disabled === false && btn.textContent === label && t.exports().length === 2,
        `disabled=${btn.disabled}，文字=${btn.textContent}，匯出 ${t.exports().length} 次`);
}

// ---- 5. 再下載一次失敗：訊息說明、不含名冊內容；結果仍保留，可以再按
{
    const s = loadSite();
    addClub(s, ROSTER);
    const t = instrument(s);
    await s.run('processPastedData()');
    s.ctx.__exportError = s.run('new TypeError("測試生301_05 儲存格錯誤")');
    s.alerts.length = 0;
    s.el('btn-redownload').click();
    await s.settle();
    R.check('再下載一次失敗：跳出一則訊息，說明沒有下載、可以再按「再下載一次」', s.alerts.length === 1 && /沒有下載/.test(s.alerts[0]) && s.alerts[0].includes('「再下載一次」'), s.alerts.join(' / ').replace(/測試生\S*/g, '[名冊]'));
    R.check('再下載一次失敗：訊息附錯誤類型、不含名冊內容', s.alerts.length === 1 && s.alerts[0].includes('TypeError') && !leaks(s.alerts));
    R.check('再下載一次失敗：lastExport 保留、按鈕還原', hasLastExport(s) && !s.el('btn-redownload').disabled && isShown(s.el('btn-redownload')));

    const loadErr = s.run('(() => { const e = new Error("x"); e.name = "ExcelJSLoadError"; return e; })()');
    s.ctx.__exportError = loadErr;
    s.alerts.length = 0;
    s.el('btn-redownload').click();
    await s.settle();
    const msg = s.run('EXCELJS_LOAD_ERROR_MESSAGE');
    R.check('再下載一次時 ExcelJS 載入失敗：先說明無法載入，再說連上網路後按「再下載一次」', s.alerts.length === 1 && s.alerts[0].startsWith(msg) && s.alerts[0].includes('「再下載一次」'), s.alerts.join(' / '));

    s.ctx.__exportError = null;
    s.alerts.length = 0;
    s.el('btn-redownload').click();
    await s.settle();
    R.check('失敗後再按一次：照常匯出同樣的內容、沒有重抽', s.alerts.length === 0 && t.exports().length === 2 && t.exports()[1] === t.exports()[0] && t.lottery() === 0);
}

// ---- 6. 上傳 Excel 模式：也會記住並可以再下載一次
{
    const s = loadSite();
    await uploadSheets(s, [{ name: '社團-週一桌遊社', rows: [['班級', '座號', '姓名'], ['301', '05', '測試生301_05'], ['302', '06', '測試生302_06']] }]);
    s.el('mapping-table-body').querySelectorAll('tr').forEach(tr => { tr.querySelector('.day-select').value = '週一'; });
    const t = instrument(s);
    s.alerts.length = 0;
    await s.run('processAndDownload()');
    R.check('Excel 模式轉換：匯出一次、lastExport 記住（沒有社團抽籤結果）', t.exports().length === 1 && hasLastExport(s) && safe(s, 'lastExport.clubs') === null,
        `匯出 ${t.exports().length} 次 ${s.alerts.join(' / ')}`);
    R.check('Excel 模式轉換後：「再下載一次」出現', isShown(s.el('btn-redownload')), btnState(s.el('btn-redownload')));
    s.el('btn-redownload').click();
    await s.settle();
    R.check('Excel 模式按「再下載一次」：匯出同樣的內容', t.exports().length === 2 && t.exports()[1] === t.exports()[0] && s.alerts.length === 0, s.alerts.join(' / '));

    // 新的一次匯出成功時覆蓋：改檔名再轉換一次
    s.el('excel-filename').value = '第二次';
    await s.run('processAndDownload()');
    s.el('btn-redownload').click();
    await s.settle();
    const ex = t.exports();
    R.check('新的一次匯出成功後：「再下載一次」匯出的是最新這次', ex.length === 4 && JSON.parse(ex[3])[3] === '第二次' && ex[3] === ex[2], ex.length === 4 ? JSON.parse(ex[3])[3] : `匯出 ${ex.length} 次`);

    // 匯出失敗不覆蓋上一次成功的結果
    s.el('excel-filename').value = '第三次';
    s.ctx.__exportError = s.run('new TypeError("x")');
    await s.run('processAndDownload()');
    s.ctx.__exportError = null;
    s.alerts.length = 0;
    s.el('btn-redownload').click();
    await s.settle();
    const ex2 = t.exports();
    R.check('匯出失敗時不覆蓋：「再下載一次」仍是上一次成功的內容', ex2.length === 5 && JSON.parse(ex2[4])[3] === '第二次', ex2.length === 5 ? JSON.parse(ex2[4])[3] : `匯出 ${ex2.length} 次`);
}

// ---- 7. 「整理另一份名冊」：清掉 lastExport 並隱藏按鈕
{
    const s = loadSite();
    addClub(s, ROSTER);
    const t = instrument(s);
    await s.run('processPastedData()');
    s.el('btn-restart').click();
    R.check('整理另一份名冊：清掉 lastExport', lastExportDefined(s) && !hasLastExport(s));
    R.check('整理另一份名冊：「再下載一次」隱藏', isHidden(s.el('btn-redownload')), btnState(s.el('btn-redownload')));
    s.el('btn-redownload').click();
    await s.settle();
    R.check('整理另一份名冊後按「再下載一次」：不匯出', t.exports().length === 1);

    // 之後再下載一次新的：按鈕重新出現、內容是新的社團
    addClub(s, '班級\t座號\t姓名\n501\t08\t測試生501_08', { name: '圍棋社' });
    await s.run('processPastedData()');
    R.check('整理另一份名冊後再下載新的：按鈕重新出現、記住的是新的社團',
        isShown(s.el('btn-redownload')) && safe(s, 'lastExport.clubs[0].clubName') === '圍棋社', `社團=${safe(s, 'lastExport && lastExport.clubs && lastExport.clubs[0].clubName')}`);
}

// ---- 8. 「清除這次的資料」：清掉 lastExport 並隱藏按鈕；按取消時保留
{
    const s = loadSite();
    addClub(s, ROSTER);
    instrument(s);
    await s.run('processPastedData()');
    const c = stubConfirm(s, false);
    s.run('manualHardResetAll()');
    R.check('清除這次的資料按取消：lastExport 與按鈕都保留', hasLastExport(s) && isShown(s.el('btn-redownload')));
    c.answer = true;
    s.run('manualHardResetAll()');
    R.check('清除這次的資料：清掉 lastExport', lastExportDefined(s) && !hasLastExport(s));
    R.check('清除這次的資料：「再下載一次」隱藏', isHidden(s.el('btn-redownload')), btnState(s.el('btn-redownload')));
}

// ---- 9. 隱藏只靠 hidden 屬性：style.css 有 .btn[hidden] 的隱藏規則，JS 不再設定 style.display
{
    R.check('style.css 有 .btn[hidden] { display: none } 規則（.btn 的 inline-flex 會蓋過瀏覽器內建的 [hidden]）', cssHidesHiddenBtn);
    const s = loadSite();
    addClub(s, ROSTER);
    instrument(s);
    await s.run('processPastedData()');
    const btn = s.el('btn-redownload');
    const noInlineDisplay = (el) => el.style.display === undefined || el.style.display === '';
    R.check('下載後按鈕顯示：hidden 為 false、沒有設定 style.display', btn.hidden === false && noInlineDisplay(btn), btnState(btn));
    s.run('setLastExport(null)');
    R.check('setLastExport(null)：hidden 為 true、沒有設定 style.display（不依賴行內樣式）', btn.hidden === true && noInlineDisplay(btn), btnState(btn));
}

// ---- 10. 再下載一次處理中按「整理另一份名冊」或「清除這次的資料」：完成後不跳回完成畫面
// 用真正的 exportWeeklySchedule（假 ExcelJS），讓 writeBuffer 停在中途模擬處理時間
function realExportSite() {
    const s = loadSite();
    ['step-upload', 'step-mapping', 'step-success'].forEach(id => { s.el(id).className = 'step-panel'; });
    s.el('step-upload').classList.add('active');
    const Base = s.ctx.ExcelJS.Workbook;
    s.ctx.__writeGate = null;
    s.ctx.ExcelJS = {
        Workbook: class extends Base {
            constructor() {
                super();
                const write = this.xlsx.writeBuffer;
                this.xlsx.writeBuffer = async () => { if (s.ctx.__writeGate) await s.ctx.__writeGate; return write(); };
            }
        },
    };
    const downloads = [];
    s.ctx.triggerFileDownload = (blob, name) => { downloads.push(name); };
    return { s, downloads };
}
const onStep = (s, id) => s.el(id).classList.contains('active');
const gate = (s) => { let release; s.ctx.__writeGate = new Promise(r => { release = r; }); return () => { s.ctx.__writeGate = null; release(); }; };

for (const [label, reset] of [
    ['整理另一份名冊（resetToUpload）', (s) => s.run('resetToUpload()')],
    ['清除這次的資料（確認）', (s) => { stubConfirm(s, true); s.run('manualHardResetAll()'); }],
]) {
    const { s, downloads } = realExportSite();
    addClub(s, ROSTER);
    s.el('paste-excel-filename').value = '競態測試';
    s.alerts.length = 0;
    await s.run('processPastedData()');
    await s.settle();
    R.check(`[${label}] 一般下載（貼上模式）：下載一次並切到完成畫面`, downloads.length === 1 && onStep(s, 'step-success') && s.alerts.length === 0,
        `下載 ${downloads.length} 次 ${s.alerts.join(' / ')}`);

    const release = gate(s);
    s.el('btn-redownload').click();
    await s.settle();
    const restart = s.el('btn-restart');
    R.check(`[${label}] 再下載一次處理中：「整理另一份名冊」停用`, restart.disabled === true, `disabled=${restart.disabled}`);
    reset(s);
    R.check(`[${label}] 處理中清除：回到第一步、lastExport 清掉`, onStep(s, 'step-upload') && !hasLastExport(s));
    s.alerts.length = 0; // 「清除這次的資料」本身會跳出已清除的說明，這裡只檢查匯出完成後有沒有再跳訊息
    release();
    await s.settle();
    R.check(`[${label}] 清除後匯出才完成：畫面不會跳回完成畫面`, !onStep(s, 'step-success') && onStep(s, 'step-upload'),
        `step-success=${onStep(s, 'step-success')} step-upload=${onStep(s, 'step-upload')}`);
    R.check(`[${label}] 匯出完成後：「整理另一份名冊」還原可按`, restart.disabled === false);
    R.check(`[${label}] 匯出完成後：仍下載按下時的那份（共 2 次、檔名相同）、沒有跳出訊息`,
        downloads.length === 2 && downloads[1] === downloads[0] && s.alerts.length === 0, `${downloads.length} 次 ${s.alerts.join(' / ')}`);
    R.check(`[${label}] 匯出完成後：「再下載一次」維持隱藏`, isHidden(s.el('btn-redownload')), btnState(s.el('btn-redownload')));
}

// 一般情況：在完成畫面按「再下載一次」，完成後仍停在完成畫面、「整理另一份名冊」可按
{
    const { s, downloads } = realExportSite();
    await uploadSheets(s, [{ name: '社團-週一桌遊社', rows: [['班級', '座號', '姓名'], ['301', '05', '測試生301_05']] }]);
    s.el('mapping-table-body').querySelectorAll('tr').forEach(tr => { tr.querySelector('.day-select').value = '週一'; });
    s.alerts.length = 0;
    await s.run('processAndDownload()');
    await s.settle();
    R.check('一般下載（上傳 Excel 模式）：下載一次並切到完成畫面', downloads.length === 1 && onStep(s, 'step-success'), `下載 ${downloads.length} 次 ${s.alerts.join(' / ')}`);
    s.el('btn-redownload').click();
    await s.settle();
    R.check('再下載一次（沒有清除）：下載第二次、仍在完成畫面、「整理另一份名冊」可按',
        downloads.length === 2 && onStep(s, 'step-success') && !s.el('btn-restart').disabled, `下載 ${downloads.length} 次`);
}


// ---- 11. 第一次下載處理中按「清除這次的資料」：完成後不保存結果、不跳到完成畫面
// 清除時學生資料就要從記憶體消失；匯出完成時不可把剛清掉的資料寫回 lastExport
for (const mode of ['貼上模式', '上傳 Excel 模式']) {
    const { s, downloads } = realExportSite();
    if (mode === '貼上模式') {
        addClub(s, ROSTER);
    } else {
        await uploadSheets(s, [{ name: '社團-週一桌遊社', rows: [['班級', '座號', '姓名'], ['301', '05', '測試生301_05']] }]);
        s.el('mapping-table-body').querySelectorAll('tr').forEach(tr => { tr.querySelector('.day-select').value = '週一'; });
    }
    const release = gate(s);
    const pending = s.run(mode === '貼上模式' ? 'processPastedData()' : 'processAndDownload()');
    await s.settle();
    stubConfirm(s, true);
    s.run('manualHardResetAll()');
    s.alerts.length = 0;
    release();
    await pending;
    await s.settle();
    R.check(`[${mode}] 第一次下載途中清除：完成後停在第一步，不跳到完成畫面`, onStep(s, 'step-upload') && !onStep(s, 'step-success'),
        `step-success=${onStep(s, 'step-success')} step-upload=${onStep(s, 'step-upload')}`);
    R.check(`[${mode}] 第一次下載途中清除：lastExport 維持清除、「再下載一次」隱藏`, !hasLastExport(s) && isHidden(s.el('btn-redownload')),
        btnState(s.el('btn-redownload')));
    R.check(`[${mode}] 第一次下載途中清除：沒有跳出錯誤訊息`, s.alerts.length === 0, s.alerts.join(' / '));
}

R.done();
