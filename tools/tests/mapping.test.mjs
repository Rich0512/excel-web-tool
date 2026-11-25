// 第 2 步對照表：analyzeAndRenderMapping() 失敗時不切到空白對照表、切換彙整模式失敗時退回並保留修改、
// 自動改用多工作表時保留修改、時段欄位排除老師與電話等欄位、以名冊字串當鍵時防 __proto__。
// 名冊內容都是假的工作表（測試生301_05…）。
import { loadSite, makeReporter, makeWorkbook, uploadSheets } from './harness.mjs';
const R = makeReporter('對照表：失敗處理、模式切換與原型防護');

const rows = (s) => s.el('mapping-table-body').querySelectorAll('tr');
const onMapping = (s) => s.el('step-mapping').classList.contains('active');
const STUDENTS = [['班級', '座號', '姓名'], ['301', '05', '測試生301_05']];
const SINGLE = [{ name: '總表', rows: [['班級', '座號', '姓名', '社團'], ['301', '05', '測試生301_05', '桌遊社'], ['302', '11', '測試生302_11', '直笛社']] }];

// ---- 失敗時不切換到空白的對照表
for (const [title, sheets] of [
    ['單一清單缺班級欄', [{ name: '總表', rows: [['座號', '姓名', '社團'], ['05', '測試生301_05', '桌遊社']] }]],
    ['單一清單找不到社團欄', [{ name: '總表', rows: STUDENTS }]],
    ['多工作表找不到工作表', [{ name: '說明', rows: [['這是說明']] }]],
]) {
    const s = loadSite();
    await uploadSheets(s, sheets);
    R.check(`${title}：有提示`, s.alerts.length >= 1, `${s.alerts.length} 則`);
    R.check(`${title}：不切換到第 2 步`, !onMapping(s));
    R.check(`${title}：清掉 uploadedWorkbook（離開頁面不誤警告）`, s.run('uploadedWorkbook') === null);
    R.check(`${title}：清空檔案選擇`, s.el('file-input').value === '');
}

// ---- 回傳值（含自動改用多工作表的遞迴）
{
    const s = loadSite();
    await uploadSheets(s, [{ name: '社團-週一桌遊社', rows: STUDENTS }]);
    R.check('多工作表檔案：進入第 2 步', onMapping(s) && rows(s).length === 1, `列數 ${rows(s).length}`);
    s.el('mode-select').value = 'single';
    const ret = s.run('analyzeAndRenderMapping()');
    R.check('單一清單找不到總表、自動改用多工作表成功時回傳 true', ret === true, String(ret));
    R.check('自動改用多工作表後模式是 multi 且有對照列', s.el('mode-select').value === 'multi' && rows(s).length === 1);
    s.el('mode-select').value = 'single';
    s.run('uploadedWorkbook = { eachSheet() {} }');
    const ret2 = s.run('analyzeAndRenderMapping()');
    R.check('自動改用多工作表也失敗時回傳 false', ret2 === false, String(ret2));
}

// ---- 在第 2 步切到多工作表失敗：退回單一清單並保留修改
{
    const s = loadSite();
    await uploadSheets(s, SINGLE);
    R.check('（前提）單一清單檔案進入第 2 步', onMapping(s) && rows(s).length === 2, `列數 ${rows(s).length}`);
    s.alerts.length = 0;
    const [r1, r2] = rows(s);
    r1.querySelector('.club-input').value = '桌上遊戲社';
    r1.querySelector('.day-select').value = '週三';
    r2.querySelector('.day-select').value = '週五';
    const sel = s.el('mode-select');
    sel.value = 'multi';
    sel.dispatch('change');
    const [n1, n2] = rows(s);
    R.check('切到多工作表失敗：有提示', s.alerts.length >= 1);
    R.check('切到多工作表失敗：模式退回單一清單、對照表恢復原本的列', sel.value === 'single' && rows(s).length === 2, `${sel.value}，列數 ${rows(s).length}`);
    R.check('切到多工作表失敗：仍在第 2 步且保留名冊', onMapping(s) && s.run('uploadedWorkbook') !== null);
    R.check('退回後保留修改過的社團名稱', n1?.querySelector('.club-input').value === '桌上遊戲社', n1?.querySelector('.club-input').value);
    R.check('退回後保留選好的星期', n1?.querySelector('.day-select').value === '週三' && n2?.querySelector('.day-select').value === '週五',
        `${n1?.querySelector('.day-select').value}/${n2?.querySelector('.day-select').value}`);
}

// ---- 在第 2 步切到單一清單、因為沒有總表自動改回多工作表：保留修改
{
    const s = loadSite();
    await uploadSheets(s, [
        { name: '社團-週一桌遊社', rows: STUDENTS },
        { name: '社團-週二直笛社', rows: [['班級', '座號', '姓名'], ['302', '11', '測試生302_11']] },
    ]);
    const [r1, r2] = rows(s);
    r1.querySelector('.club-input').value = '桌上遊戲社';
    r2.querySelector('.day-select').value = '週四';
    const sel = s.el('mode-select');
    sel.value = 'single';
    sel.dispatch('change');
    const [n1, n2] = rows(s);
    R.check('沒有總表時自動改回多工作表，並提示 1 次', sel.value === 'multi' && rows(s).length === 2 && s.alerts.length === 1, s.alerts.join(' / ').slice(0, 60));
    R.check('自動改回後保留修改過的社團名稱', n1?.querySelector('.club-input').value === '桌上遊戲社', n1?.querySelector('.club-input').value);
    R.check('自動改回後保留選好的星期', n2?.querySelector('.day-select').value === '週四', n2?.querySelector('.day-select').value);
}

// ---- 老師、電話等欄位不當成時段欄
{
    const s = loadSite();
    s.ctx.__wb = makeWorkbook([{ name: '總表', rows: [
        ['班級', '座號', '姓名', '社團', '指導老師', '社團教師', '社團備註', '家長姓名', '家長電話'],
        ['301', '05', '測試生301_05', '桌遊社', '測試師甲', '測試師乙', '自備材料', '測試家長', '0900000000'],
    ] }]);
    s.run("uploadedWorkbook = __wb; document.getElementById('mode-select').value = 'single'; analyzeAndRenderMapping();");
    const slotNames = s.run('slotCols.map(c => c.name)');
    R.check('時段欄只有「社團」（老師、教師、備註、姓名、電話、家長欄被排除）', slotNames.length === 1 && slotNames[0] === '社團', slotNames.join(','));
}

// ---- 以名冊字串當鍵的物件要防 __proto__
const KEY = 'yenping_club_config';
const setAllDays = (s, day) => rows(s).forEach(tr => { tr.querySelector('.day-select').value = day; });
const own = (o, k) => !!o && Object.prototype.hasOwnProperty.call(o, k);
{
    const s = loadSite();
    s.ctx.__wb = makeWorkbook([
        { name: '__proto__', rows: STUDENTS },
        { name: 'constructor', rows: [['班級', '座號', '姓名'], ['302', '11', '測試生302_11']] },
    ]);
    s.run("uploadedWorkbook = __wb; document.getElementById('mode-select').value = 'multi'; analyzeAndRenderMapping();");
    R.check('多工作表：__proto__、constructor 工作表都出現在對照表', rows(s).length === 2, `${rows(s).length} 列`);
    setAllDays(s, '週一');
    const { finalMapping } = s.run('collectMapping()');
    R.check('collectMapping 保留 __proto__ 工作表', Object.keys(finalMapping).includes('__proto__'), Object.keys(finalMapping).join(','));
    const { resultData } = s.run('runExcelProcessing(collectMapping().finalMapping)');
    R.check('__proto__ 工作表的學生有排進課表', resultData.length === 2, `課表 ${resultData.length} 位`);
    const { mode, rows: mrows } = s.run('collectMapping()');
    s.run('saveMappingConfig')(mode, mrows);
    const raw = JSON.parse(s.storage.get(KEY));
    R.check('__proto__ 工作表的星期有存進瀏覽器設定', own(raw, '__proto__') && raw['__proto__'] === '週一', s.storage.get(KEY));
    const cfg = s.run('loadLocalStorageConfig()');
    R.check('讀回設定時 __proto__ 是一般的鍵', own(cfg, '__proto__') && cfg['__proto__'] === '週一');
    s.run('analyzeAndRenderMapping()');
    const tr = rows(s).find(t => t.dataset.key === '__proto__');
    const opt = tr?.querySelector('.day-select').children.find(o => o.selected);
    R.check('重建對照表時套用 __proto__ 工作表的記憶星期', opt?.value === '週一', opt?.value);
    R.check('Object.prototype 沒有被污染', s.run('Object.keys(Object.prototype).length') === 0 && Object.keys(Object.prototype).length === 0);
}
{
    const s = loadSite();
    s.ctx.__wb = makeWorkbook([{ name: '總表', rows: [['班級', '座號', '姓名', '社團'], ['301', '05', '測試生301_05', '__proto__'], ['302', '11', '測試生302_11', 'toString']] }]);
    s.run("uploadedWorkbook = __wb; document.getElementById('mode-select').value = 'single'; analyzeAndRenderMapping();");
    setAllDays(s, '週二');
    const { mode, rows: mrows } = s.run('collectMapping()');
    s.run('saveMappingConfig')(mode, mrows);
    const raw = JSON.parse(s.storage.get(KEY));
    R.check('單一清單：__proto__ 社團的星期有存進設定', own(raw['社團'], '__proto__'), s.storage.get(KEY));
    const { resultData } = s.run('runExcelProcessing(collectMapping().finalMapping)');
    R.check('單一清單：__proto__、toString 社團的學生有排進課表', resultData.length === 2 && resultData.every(r => r.schedule['週二']), `課表 ${resultData.length} 位`);
}
{
    const s = loadSite();
    s.ctx.__wb = makeWorkbook([
        { name: '社團-__proto__', rows: [['上課時間', '', '週三']] },
        { name: '1152', rows: [['課程名稱', '上課時間'], ['constructor', '週四']] },
    ]);
    const m = s.run('loadScheduleFromExcel(__wb)');
    R.check('loadScheduleFromExcel 保留名為 __proto__ 的社團', own(m, '__proto__') && m['__proto__'] === '週三', m ? Object.keys(m).join(',') : 'null');
    R.check('loadScheduleFromExcel 保留名為 constructor 的課程', own(m, 'constructor') && m['constructor'] === '週四');
    s.ctx.__wb2 = makeWorkbook([{ name: '1152', rows: [['課程名稱', '上課時間'], ['桌遊社', '週一']] }]);
    const m2 = s.run('loadScheduleFromExcel(__wb2)');
    R.check('查不存在的社團（toString）得到 undefined，不會查到 Object.prototype', m2 && m2['toString'] === undefined);
}
R.done();
