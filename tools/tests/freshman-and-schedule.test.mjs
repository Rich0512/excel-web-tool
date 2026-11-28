// 新生判斷與課程表略過：
// - 多工作表模式：學校系統匯出的課程時間表（前 5 列有「課程名稱」「上課時間」、沒有姓名欄）不論工作表叫什麼名稱都略過，
//   不列入對照表、不跳「找不到姓名欄」的提醒，但它的上課時間仍用來自動選星期
// - 只有課程表時的提示列出所有會略過的名稱關鍵字（含 README、1152）與課程表
// - 班級寫「新生」的學生：不包含新生時，單一清單、多工作表、貼上三條路徑都不排進課表（與抽籤的 parseGrade 一致）
// - OLE2 檔頭（.xls 或設了開啟密碼的 .xlsx）的訊息提到開啟密碼
// 名冊都是虛構的假工作表（測試生301_05…），不讀取任何真實檔案。
import { loadSite, makeReporter, makeSheet, uploadSheets } from './harness.mjs';
const R = makeReporter('新生判斷與課程表略過');

const H3 = ['班級', '座號', '姓名'];
const rows = (s) => s.el('mapping-table-body').querySelectorAll('tr');
const onMapping = (s) => s.el('step-mapping').classList.contains('active');
const names = (res) => Array.from(res.resultData, r => r.name).sort().join(',');

// ============================================================
// 1. findScheduleHeader：課程表標頭偵測
// ============================================================
{
    const s = loadSite();
    const find = s.fn('findScheduleHeader');
    R.test('findScheduleHeader：第 2 列是「課程名稱」「上課時間」→ 回傳列號與欄號', () => {
        const r = find(makeSheet({ name: '1161', rows: [['115 學年度第 1 學期課程時間表'], ['編號', '課程名稱', '上課時間'], ['1', '桌遊社', '週三']] }));
        return [r && r.headerRow === 2 && r.courseCol === 2 && r.dayCol === 3, JSON.stringify(r)];
    });
    R.test('findScheduleHeader：一般的社團名單 → null', () => {
        const r = find(makeSheet({ name: '社團-桌遊社', rows: [H3, ['301', '5', '測試生301_05']] }));
        return [r === null, JSON.stringify(r)];
    });
    R.test('findScheduleHeader：標頭在第 6 列 → null（只看前 5 列）', () => {
        const r = find(makeSheet({ name: 'x', rows: [[], [], [], [], [], ['課程名稱', '上課時間']] }));
        return [r === null, JSON.stringify(r)];
    });
}

// ============================================================
// 2. 多工作表模式：名稱沒有關鍵字的課程表被略過，上課時間仍用來選星期
// ============================================================
for (const scheduleName of ['1161', '課程時間表']) {
    const s = loadSite();
    await uploadSheets(s, [
        { name: scheduleName, rows: [['編號', '課程名稱', '上課時間', '授課教師'], ['1', '桌遊社', '週三', '王老師'], ['2', '直笛社', '週五', '李老師']] },
        { name: '社團-桌遊社', rows: [H3, ['301', '5', '測試生301_05']] },
        { name: '社團-直笛社', rows: [H3, ['302', '11', '測試生302_11']] },
    ]);
    const trs = rows(s);
    const keys = trs.map(tr => tr.dataset.original);
    R.check(`課程表「${scheduleName}」：進入第 2 步、沒有提示`, onMapping(s) && s.alerts.length === 0, s.alerts.join(' / '));
    R.check(`課程表「${scheduleName}」：不出現在對照表`, keys.join(',') === '社團-桌遊社,社團-直笛社', keys.join(','));
    const skip = s.el('mapping-skip-alert');
    R.check(`課程表「${scheduleName}」：不列入「找不到姓名欄」的提醒`, skip.hidden !== false && !skip.innerHTML.includes(scheduleName), skip.innerHTML.slice(0, 80));
    // 假 DOM 的 <select> 不會依 option.selected 更新 value，改看被選取的選項
    const days = trs.map(tr => tr.querySelector('.day-select').children.find(o => o.selected)?.value);
    R.check(`課程表「${scheduleName}」：上課時間仍用來自動選星期`, days.join(',') === '週三,週五', days.join(','));
}

// ---- 有課程表標頭、也有姓名欄的工作表仍是社團名單（不是只看標頭就略過）
{
    const s = loadSite();
    await uploadSheets(s, [
        { name: '社團-桌遊社', rows: [['班級', '座號', '姓名', '課程名稱', '上課時間'], ['301', '5', '測試生301_05', '桌遊社', '週三']] },
    ]);
    R.check('有姓名欄的工作表即使有「課程名稱」「上課時間」也列入對照表', rows(s).length === 1 && s.alerts.length === 0, `${rows(s).length} 列；${s.alerts.join(' / ')}`);
}

// ---- 沒有姓名欄、也不是課程表的工作表仍列入「找不到姓名欄」的提醒
{
    const s = loadSite();
    await uploadSheets(s, [
        { name: '1161', rows: [['課程名稱', '上課時間'], ['桌遊社', '週三']] },
        { name: '社團-桌遊社', rows: [H3, ['301', '5', '測試生301_05']] },
        { name: '社團-直笛社', rows: [['學生', '班'], ['測試生302_11', '302']] },
    ]);
    const skip = s.el('mapping-skip-alert');
    R.check('沒有姓名欄的社團工作表仍列入提醒，課程表不列入',
        skip.hidden === false && skip.innerHTML.includes('社團-直笛社') && !skip.innerHTML.includes('1161') && skip.innerHTML.includes('有 1 張'),
        skip.innerHTML.slice(0, 120));
}

// ============================================================
// 3. 只有課程表、沒有社團工作表：提示列出所有略過的關鍵字與課程表
// ============================================================
{
    const s = loadSite();
    await uploadSheets(s, [
        { name: '1161', rows: [['課程名稱', '上課時間'], ['桌遊社', '週三']] },
        { name: 'README', rows: [['這是說明']] },
    ]);
    const msg = s.alerts.join('\n');
    R.check('只有課程表：有一則提示、不進入第 2 步', s.alerts.length === 1 && !onMapping(s), `${s.alerts.length} 則`);
    R.check('只有課程表：提示是「找不到可以彙整的社團名單工作表」，不是「找不到姓名欄」',
        msg.includes('找不到可以') && !msg.includes('找不到「姓名」欄'), msg.slice(0, 80));
    for (const kw of ['README', '1152', '說明', '統計', '清單', '報表', '對照', '人數', '小計', '總表']) {
        R.check(`只有課程表：提示列出關鍵字「${kw}」`, msg.includes(`「${kw}」`), '');
    }
    R.check('只有課程表：提示說明課程表也會略過', msg.includes('課程表') && msg.includes('「課程名稱」') && msg.includes('「上課時間」'), msg);
    R.check('只有課程表：略過清單列出兩張工作表', msg.includes('- 1161') && msg.includes('- README'), msg);
}

// ============================================================
// 4. 班級寫「新生」的學生
// ============================================================
{
    const s = loadSite();
    const isFreshman = s.fn('isFreshmanClass');
    R.test('isFreshmanClass：空白、「新生」、「新生班」是新生；「301」「資源班」不是', () => {
        const got = ['', '  ', '新生', '新生班', '301', '資源班'].map(c => isFreshman(c));
        return [JSON.stringify(got) === JSON.stringify([true, true, true, true, false, false]), JSON.stringify(got)];
    });

    // 單一清單
    const key = s.fn('mappingKey')('社團', '桌遊社');
    const single = (includeFreshmen) => s.fn('processExcelData')({
        mode: 'single', includeFreshmen, slotMode: 'semester',
        finalMapping: { [key]: { day: '週一', editedName: '桌遊社' } },
        sheetData: [
            { 班級: '301', 座號: '5', 姓名: '測試生301_05', 社團: '桌遊社' },
            { 班級: '新生', 座號: '', 姓名: '測試生新生_01', 社團: '桌遊社' },
            { 班級: '', 座號: '', 姓名: '測試生新生_02', 社團: '桌遊社' },
        ],
        classCol: '班級', seatCol: '座號', nameCol: '姓名', slotCols: [{ name: '社團', index: 3 }],
    });
    R.check('單一清單：不含新生時，班級「新生」與空白都不列入', names(single(false)) === '測試生301_05', names(single(false)));
    R.check('單一清單：含新生時三位都列入', names(single(true)) === '測試生301_05,測試生新生_01,測試生新生_02', names(single(true)));

    // 多工作表
    const multi = (includeFreshmen) => s.fn('processExcelData')({
        mode: 'multi', includeFreshmen, slotMode: 'semester',
        finalMapping: { '社團-桌遊社': { day: '週一', editedName: '桌遊社' } },
        sheetStudents: new Map([['社團-桌遊社', [
            { class: '301', seat: '5', name: '測試生301_05' },
            { class: '新生', seat: '', name: '測試生新生_01' },
            { class: '', seat: '', name: '測試生新生_02' },
        ]]]),
    });
    R.check('多工作表：不含新生時，班級「新生」與空白都不列入', names(multi(false)) === '測試生301_05', names(multi(false)));
    R.check('多工作表：含新生時三位都列入', names(multi(true)) === '測試生301_05,測試生新生_01,測試生新生_02', names(multi(true)));

    // 貼上
    const pasted = (includeFreshmen) => s.fn('processPastedClubsData')([{
        clubName: '桌遊社', day: '週一', students: [
            { class: '301', seat: '5', name: '測試生301_05', selected: true },
            { class: '新生', seat: '', name: '測試生新生_01', selected: true },
            { class: '', seat: '', name: '測試生新生_02', selected: true },
        ],
    }], 'semester', includeFreshmen);
    R.check('貼上：不含新生時，班級「新生」與空白都不列入', names(pasted(false)) === '測試生301_05', names(pasted(false)));
    R.check('貼上：含新生時三位都列入', names(pasted(true)) === '測試生301_05,測試生新生_01,測試生新生_02', names(pasted(true)));
}

// ---- 從上傳走到轉換（多工作表，班級欄寫「新生」）
{
    const s = loadSite();
    await uploadSheets(s, [
        { name: '社團-週一桌遊社', rows: [H3, ['301', '5', '測試生301_05'], ['新生', '', '測試生新生_01']] },
    ]);
    s.el('include-freshmen').checked = false;
    const off = s.run('runExcelProcessing(collectMapping().finalMapping)');
    R.check('上傳多工作表、不含新生：班級寫「新生」的學生不列入', names(off) === '測試生301_05', names(off));
    s.el('include-freshmen').checked = true;
    const on = s.run('runExcelProcessing(collectMapping().finalMapping)');
    R.check('上傳多工作表、含新生：列入', names(on) === '測試生301_05,測試生新生_01', names(on));
}

// ============================================================
// 5. OLE2 檔頭：舊版 .xls 或設了開啟密碼的 .xlsx
// ============================================================
{
    const s = loadSite();
    const bytes = new Uint8Array(16);
    bytes.set([0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]);
    const file = { name: '名冊.xlsx', size: 4096, slice: () => ({ arrayBuffer: async () => bytes.buffer.slice(0, 8) }) };
    const r = await s.fn('checkExcelFileFormat')(file);
    R.check('OLE2 檔頭：不可讀取', r && r.valid === false, JSON.stringify(r));
    R.check('OLE2 檔頭：訊息保留「舊版 Excel 97-2003（.xls）格式」', !!r?.reason?.includes('舊版 Excel 97-2003（.xls）格式'), r?.reason);
    R.check('OLE2 檔頭：訊息提到開啟密碼與取消密碼', !!r?.reason?.includes('開啟密碼') && r.reason.includes('取消密碼') && r.reason.includes('另存新檔'), r?.reason);
}

R.done();
