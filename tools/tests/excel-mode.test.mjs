// Excel 模式轉換後的週課表內容：單一清單（時段欄位、衝突備註、「無」與空白、同名同班、合併跨列、新生、寒暑假標籤）、
// 多工作表（工作表對到星期、衝突、新生、寒暑假標籤）、晨間時段不算衝突，以及匯出檔「課程與學生清單」工作表
// （用真的 ExcelJS 讀回，核對表頭、每一列與排序）。至少一個案例從第 1 步上傳、第 2 步對照表、按「開始轉換」一路走到下載。
// 名冊都是虛構的假工作表（測試生301_05…），不讀取任何真實檔案。
import vm from 'node:vm';
import { readFileSync } from 'node:fs';
import { loadSite, makeReporter, uploadSheets, EXCELJS_PATH } from './harness.mjs';
const R = makeReporter('Excel 模式：轉換後的課表內容與匯出檔');

const ZWSP = String.fromCharCode(0x200B);
const WEEKDAYS = ['週一', '週二', '週三', '週四', '週五'];
const conflict = (day, clubs) => `「${day}」時段衝突：同時錄取「${clubs}」`;
// 期望的學生：days 依星期順序寫（finalizeStudents 依 DAY_ORDER 排序 schedule 的鍵）
const st = (cls, seat, name, days = {}, remarks = []) => ({ class: cls, seat, name, schedule: days, remarks });
const canon = (r) => JSON.stringify({ class: r.class, seat: r.seat, name: r.name, schedule: r.schedule, remarks: r.remarks });
const who = (r) => `${r.class} 班 ${r.seat || '（無座號）'} 號 ${r.name}`;
const onStep = (s, id) => s.el(id).classList.contains('active');

// ---- 真的 ExcelJS（寫出與讀回匯出檔）
// 要載入頁面同一個 vm context（ExcelJS 用 instanceof Array 判斷 addRow 的參數）。上傳用的是 harness 的假 ExcelJS，
// 所以先上傳、再換成真的 ExcelJS 匯出。頁面的 setTimeout 是假時鐘，ExcelJS 的非同步工作要靠 advance() 推進。
function loadRealExcelJS(s) {
    Object.assign(s.ctx, { TextDecoder, TextEncoder, regeneratorRuntime: undefined });
    vm.runInContext(readFileSync(EXCELJS_PATH, 'utf8'), s.ctx);
    return s.ctx.ExcelJS;
}
async function waitFor(s, cond, max = 20000) {
    for (let i = 0; i < max && !cond(); i++) { s.advance(5); await new Promise(r => setImmediate(r)); }
    return cond();
}
async function pump(s, promise) {
    let done = false;
    const p = Promise.resolve(promise).finally(() => { done = true; });
    await waitFor(s, () => done);
    return p;
}
// 讀回匯出檔：回傳工作表名稱與「課程與學生清單」的所有列（空儲存格當成 ''）
async function readExport(s, lib, blob) {
    const wb = new lib.Workbook();
    await pump(s, wb.xlsx.load(blob.parts[0]));
    const names = wb.worksheets.map(w => w.name);
    const ws = wb.getWorksheet('課程與學生清單');
    if (!ws) return { names, header: [], rows: [] };
    const width = ws.getRow(1).cellCount;
    const rowAt = (r) => Array.from({ length: width }, (_, c) => ws.getCell(r, c + 1).value ?? '');
    return { names, header: rowAt(1), rows: Array.from({ length: ws.rowCount - 1 }, (_, i) => rowAt(i + 2)) };
}
// 按「開始轉換並下載課表」，等到下載（或跳出提示）
async function clickConvert(s, lib) {
    let dl = null;
    s.ctx.triggerFileDownload = (blob, name) => { dl = { blob, name }; };
    const before = s.alerts.length;
    s.el('btn-process-mapping').click();
    await waitFor(s, () => dl || s.alerts.length > before);
    const alerts = s.alerts.slice(before);
    if (!dl) return { dl: null, alerts };
    return { dl, alerts, file: await readExport(s, lib, dl.blob) };
}
// 期望的學生 → 匯出檔的一列
const toRow = (r, days) => [r.class, r.seat, r.name, ...days.map(d => r.schedule[d] ?? ''), r.remarks.join('；')];

// 在對照表設定星期（與改名）。single：鍵是「時段欄位|社團」；multi：鍵是工作表名稱
function setMapping(s, mode, spec) {
    const trs = s.el('mapping-table-body').querySelectorAll('tr');
    const seen = [];
    trs.forEach(tr => {
        const key = mode === 'single' ? `${tr.dataset.slot}|${tr.dataset.original}` : tr.dataset.original;
        seen.push(key);
        const want = spec[key];
        if (!want) return;
        tr.querySelector('.day-select').value = want.day;
        if (want.name) tr.querySelector('.club-input').value = want.name;
    });
    return seen;
}
const runMapping = (s) => s.run('runExcelProcessing(collectMapping().finalMapping)');

// 比對 processExcelData 的結果與期望（不看順序，逐位學生各記一個檢查）
function checkStudents(title, resultData, expected) {
    const got = new Map(Array.from(resultData, r => [`${r.class}|${r.seat}|${r.name}`, r]));
    R.check(`${title}：共 ${expected.length} 位學生，名單相同`,
        got.size === resultData.length && resultData.length === expected.length && expected.every(e => got.has(`${e.class}|${e.seat}|${e.name}`)),
        Array.from(resultData, who).join('、'));
    for (const e of expected) {
        const r = got.get(`${e.class}|${e.seat}|${e.name}`);
        R.check(`${title}：${who(e)} 的課表與備註`, !!r && canon(r) === canon(e), r ? canon(r) : '（沒有這位學生）');
    }
}
// 比對匯出檔：表頭、每一列（依期望順序）、與 processExcelData 的結果一致、依班級座號排序
function checkExport(title, file, expected, days, resultData) {
    const header = ['班級', '座號', '姓名', ...days, '備註'];
    R.check(`${title}：只有「課程與學生清單」工作表`, file.names.length === 1 && file.names[0] === '課程與學生清單', file.names.join(','));
    R.check(`${title}：表頭是 ${header.join('、')}`, JSON.stringify(file.header) === JSON.stringify(header), file.header.join('、'));
    R.check(`${title}：資料列數是 ${expected.length}`, file.rows.length === expected.length, `${file.rows.length} 列`);
    expected.forEach((e, i) => {
        const want = toRow(e, days);
        const got = file.rows[i];
        R.check(`${title}：第 ${i + 2} 列是 ${who(e)}，各星期社團與備註正確`, !!got && JSON.stringify(got) === JSON.stringify(want), got ? got.join(' | ') : '（沒有這一列）');
    });
    if (resultData) {
        const fromResult = Array.from(resultData, r => JSON.stringify(toRow({ ...r, remarks: Array.from(r.remarks) }, days))).sort();
        const fromFile = file.rows.map(r => JSON.stringify(r)).sort();
        R.check(`${title}：匯出的每一列與 processExcelData 的結果一致`, JSON.stringify(fromResult) === JSON.stringify(fromFile),
            `結果 ${fromResult.length} 位，檔案 ${fromFile.length} 列`);
    }
    // 排序：班級數字由小到大（新生最後），同班座號數字由小到大
    const num = (v) => (v === '' || v === '新生' ? Infinity : Number(v));
    const sorted = file.rows.every((r, i) => {
        if (i === 0) return true;
        const p = file.rows[i - 1];
        return num(p[0]) < num(r[0]) || (num(p[0]) === num(r[0]) && num(p[1]) <= num(r[1]));
    });
    R.check(`${title}：依班級、座號（數字大小）排序，新生在最後`, sorted, file.rows.map(r => `${r[0]}/${r[1]}`).join(' '));
}

// ============================================================
// 1. 單一清單模式：總表有兩個時段欄位
// ============================================================
const SINGLE_SHEETS = [{
    name: '總表', rows: [
        ['班級', '座號', '姓名', '時段A', '時段B'],
        ['302', '11', '測試生302_11', '桌遊社', '直笛社'],   // 同一天兩個社團 → 衝突
        ['301', '5', '測試生301_05', '桌遊社', '無'],         // 「無」不列入
        ['301', '12', '測試生301_12', '直笛社', ''],          // 空白不列入
        ['301', '5', '測試生301_05', '', '圍棋社'],           // 同一位學生跨列 → 合併
        ['301', '7', '測試生301_05', '籃球社', ''],           // 同名同班、座號不同 → 分開列
        ['', '', '測試生新生_01', '圍棋社', ''],              // 新生（班級空白）
        [ZWSP, '', '測試生新生_02', '', '直笛社'],            // 新生（班級只有不可見字元）
        ['201', '3', '測試生201_03', '無', '無'],             // 沒有任何社團
        ['302', '2', '測試生302_02', '圍棋社', '圍棋社'],     // A、B 同名社團同一天
    ],
}];
const SINGLE_MAPPING = {
    '時段A|桌遊社': { day: '週三' },
    '時段A|直笛社': { day: '週一' },
    '時段A|籃球社': { day: '週二', name: '籃球隊' },   // 在對照表改名
    '時段A|圍棋社': { day: '週四' },
    '時段B|直笛社': { day: '週三' },
    '時段B|圍棋社': { day: '週一、四' },               // 一個社團上兩天
};
// 期望的匯出順序（班級、座號數字排序；新生最後、照名冊順序）
const SEM_FRESH = [
    st('新生', '', '測試生新生_01', { 週四: '圍棋社' }),
    st('新生', '', '測試生新生_02', { 週三: '直笛社' }),
];
const SEM = [
    st('201', '3', '測試生201_03'),
    st('301', '5', '測試生301_05', { 週一: '圍棋社', 週三: '桌遊社', 週四: '圍棋社' }),
    st('301', '7', '測試生301_05', { 週二: '籃球隊' }),
    st('301', '12', '測試生301_12', { 週一: '直笛社' }),
    st('302', '2', '測試生302_02', { 週一: '圍棋社', 週四: '圍棋社' }),
    st('302', '11', '測試生302_11', { 週三: '桌遊社, 直笛社' }, [conflict('週三', '桌遊社, 直笛社')]),
];
const VAC_FRESH = [
    st('新生', '', '測試生新生_01', { 週四: '圍棋社(A)' }),
    st('新生', '', '測試生新生_02', { 週三: '直笛社(B)' }),
];
const VAC = [
    st('201', '3', '測試生201_03'),
    st('301', '5', '測試生301_05', { 週一: '圍棋社(B)', 週三: '桌遊社(A)', 週四: '圍棋社(B)' }),
    st('301', '7', '測試生301_05', { 週二: '籃球隊(A)' }),
    st('301', '12', '測試生301_12', { 週一: '直笛社(A)' }),
    st('302', '2', '測試生302_02', { 週一: '圍棋社(B)', 週四: '圍棋社(A), 圍棋社(B)' }, [conflict('週四', '圍棋社(A), 圍棋社(B)')]),
    st('302', '11', '測試生302_11', { 週三: '桌遊社(A), 直笛社(B)' }, [conflict('週三', '桌遊社(A), 直笛社(B)')]),
];
{
    const s = loadSite();
    await uploadSheets(s, SINGLE_SHEETS);
    R.check('單一清單：上傳後進入第 2 步、模式是單一清單', onStep(s, 'step-mapping') && s.el('mode-select').value === 'single', s.alerts.join(' / '));
    const seen = setMapping(s, 'single', SINGLE_MAPPING);
    R.check('單一清單：對照表是每個「時段欄位 × 社團」一列，沒有「無」與空白',
        JSON.stringify(seen) === JSON.stringify(Object.keys(SINGLE_MAPPING)), seen.join('、'));
    const lib = loadRealExcelJS(s);

    // ---- 學期、包含新生：從按鈕一路走到下載
    s.el('include-freshmen').checked = true;
    s.el('slot-mode-select').value = 'semester';
    const expected = [...SEM, ...SEM_FRESH];
    const resultData = runMapping(s);
    checkStudents('單一清單（學期、含新生）', resultData.resultData, expected);
    R.check('單一清單（學期、含新生）：activeDays 是週一到週四',
        JSON.stringify([...resultData.activeDays].sort()) === JSON.stringify(['週一', '週二', '週三', '週四'].sort()), [...resultData.activeDays].join(','));
    R.check('單一清單：沒有任何學生的課表出現「無」或空白的社團',
        Array.from(resultData.resultData).every(r => Object.values(r.schedule).every(v => v && !v.split(', ').some(c => c === '無' || c === ''))));
    R.check('單一清單：同名同班、座號不同的學生分開列（301 班 5 號與 7 號）',
        Array.from(resultData.resultData).filter(r => r.class === '301' && r.name === '測試生301_05').map(r => r.seat).sort().join(',') === '5,7');

    const r1 = await clickConvert(s, lib);
    R.check('按「開始轉換」（學期）：下載檔名是「名冊_學期社團彙整.xlsx」，沒有提示', r1.dl?.name === '名冊_學期社團彙整.xlsx' && r1.alerts.length === 0,
        `${r1.dl?.name}；${r1.alerts.join(' / ')}`);
    R.check('按「開始轉換」（學期）：切到完成步驟', onStep(s, 'step-success'));
    const cfg = s.storage.get('yenping_club_config') || '';
    R.check('按「開始轉換」：記住的星期設定沒有學生資料', cfg.includes('桌遊社') && !/測試生|301|302|201/.test(cfg), cfg);
    if (r1.file) checkExport('匯出檔（單一清單、學期、含新生）', r1.file, expected, WEEKDAYS, resultData.resultData);

    // ---- 學期、不含新生
    s.el('include-freshmen').checked = false;
    const noFresh = runMapping(s);
    checkStudents('單一清單（學期、不含新生）', noFresh.resultData, SEM);
    R.check('單一清單（學期、不含新生）：班級空白或只有不可見字元的學生都不出現',
        Array.from(noFresh.resultData).every(r => r.class !== '新生' && !r.name.startsWith('測試生新生')));

    // ---- 寒暑假、包含新生：再按一次按鈕
    s.el('include-freshmen').checked = true;
    s.el('slot-mode-select').value = 'vacation';
    const vac = runMapping(s);
    checkStudents('單一清單（寒暑假、含新生）', vac.resultData, [...VAC, ...VAC_FRESH]);
    const r2 = await clickConvert(s, lib);
    R.check('按「開始轉換」（寒暑假）：下載檔名是「名冊_寒暑假社團彙整.xlsx」', r2.dl?.name === '名冊_寒暑假社團彙整.xlsx' && r2.alerts.length === 0,
        `${r2.dl?.name}；${r2.alerts.join(' / ')}`);
    if (r2.file) checkExport('匯出檔（單一清單、寒暑假、含新生）', r2.file, [...VAC, ...VAC_FRESH], WEEKDAYS, vac.resultData);

    // ---- 寒暑假、不含新生
    s.el('include-freshmen').checked = false;
    checkStudents('單一清單（寒暑假、不含新生）', runMapping(s).resultData, VAC);
}

// ---- 單一清單：時段欄位名稱沒有 A/B 時，寒暑假模式用清掉通用字詞後的欄位名稱；學期模式不加標籤
{
    const s = loadSite();
    await uploadSheets(s, [{ name: '總表', rows: [
        ['班級', '座號', '姓名', '上午社團', '下午社團'],
        ['401', '1', '測試生401_01', '桌遊社', '直笛社'],
    ] }]);
    setMapping(s, 'single', { '上午社團|桌遊社': { day: '週一' }, '下午社團|直笛社': { day: '週二' } });
    s.el('slot-mode-select').value = 'vacation';
    checkStudents('單一清單（寒暑假、欄位名稱沒有 A/B）', runMapping(s).resultData,
        [st('401', '1', '測試生401_01', { 週一: '桌遊社(上午)', 週二: '直笛社(下午)' })]);
    s.el('slot-mode-select').value = 'semester';
    checkStudents('單一清單（學期、欄位名稱沒有 A/B）', runMapping(s).resultData,
        [st('401', '1', '測試生401_01', { 週一: '桌遊社', 週二: '直笛社' })]);
}

// ============================================================
// 2. 多工作表模式：每個社團一張工作表
// ============================================================
const H3 = ['班級', '座號', '姓名'];
const MULTI_SHEETS = [
    { name: '說明', rows: [['這是說明']] },                                   // 不是社團名單，略過
    { name: '社團-足球社', rows: [H3, ['301', '5', '測試生301_05']] },
    { name: '社團-籃球社', rows: [H3, ['301', '5', '測試生301_05']] },
    { name: '社團-桌遊社', rows: [H3, ['301', '5', '測試生301_05'], ['302', '11', '測試生302_11'], ['', '', '測試生新生_01']] },
    { name: '社團-直笛社', rows: [H3, ['303', '8', '測試生303_08'], ['301', '5', '測試生301_05']] },
    { name: '社團-圍棋社', rows: [H3, ['302', '11', '測試生302_11'], ['301', '7', '測試生301_05']] },
    { name: '社團-晨間扯鈴', rows: [H3, ['301', '5', '測試生301_05']] },
    { name: '社團-晨間跳繩', rows: [H3, ['301', '5', '測試生301_05']] },
];
const MULTI_MAPPING = {
    '社團-足球社': { day: '週四' },
    '社團-籃球社': { day: '週一、四' },
    '社團-桌遊社': { day: '週三' },
    '社團-直笛社': { day: '週三' },
    '社團-圍棋社': { day: '週一、四' },
    '社團-晨間扯鈴': { day: '晨間(週二、週五)' },
    '社團-晨間跳繩': { day: '晨間(週二、週五)' },
};
const MORNING_TF = '晨間(週二、週五)';
const MULTI = [
    st('301', '5', '測試生301_05',
        { 週一: '籃球社', 週三: '桌遊社, 直笛社', 週四: '足球社, 籃球社', [MORNING_TF]: '晨間扯鈴, 晨間跳繩' },
        [conflict('週三', '桌遊社, 直笛社'), conflict('週四', '足球社, 籃球社')]),
    st('301', '7', '測試生301_05', { 週一: '圍棋社', 週四: '圍棋社' }),
    st('302', '11', '測試生302_11', { 週一: '圍棋社', 週三: '桌遊社', 週四: '圍棋社' }),
    st('303', '8', '測試生303_08', { 週三: '直笛社' }),
];
const MULTI_FRESH = [st('新生', '', '測試生新生_01', { 週三: '桌遊社' })];
{
    const s = loadSite();
    await uploadSheets(s, MULTI_SHEETS);
    R.check('多工作表：上傳後進入第 2 步、模式是多工作表', onStep(s, 'step-mapping') && s.el('mode-select').value === 'multi', s.alerts.join(' / '));
    const seen = setMapping(s, 'multi', MULTI_MAPPING);
    R.check('多工作表：對照表每個社團工作表一列，「說明」工作表略過', JSON.stringify(seen) === JSON.stringify(Object.keys(MULTI_MAPPING)), seen.join('、'));
    const names = s.el('mapping-table-body').querySelectorAll('tr').map(tr => tr.querySelector('.club-input').value);
    R.check('多工作表：社團名稱去掉「社團-」前綴', names.join(',') === '足球社,籃球社,桌遊社,直笛社,圍棋社,晨間扯鈴,晨間跳繩', names.join(','));
    const lib = loadRealExcelJS(s);

    s.el('include-freshmen').checked = true;
    s.el('slot-mode-select').value = 'semester';
    const res = runMapping(s);
    checkStudents('多工作表（學期、含新生）', res.resultData, [...MULTI, ...MULTI_FRESH]);
    const r = await clickConvert(s, lib);
    R.check('多工作表按「開始轉換」：下載檔名是「名冊_學期社團彙整.xlsx」', r.dl?.name === '名冊_學期社團彙整.xlsx' && r.alerts.length === 0, `${r.dl?.name}；${r.alerts.join(' / ')}`);
    if (r.file) checkExport('匯出檔（多工作表、含晨間欄）', r.file, [...MULTI, ...MULTI_FRESH], [...WEEKDAYS, MORNING_TF], res.resultData);

    s.el('include-freshmen').checked = false;
    checkStudents('多工作表（學期、不含新生）', runMapping(s).resultData, MULTI);
}

// ---- 多工作表：A、B 時段的同名社團（寒暑假加標籤並算衝突；學期時是同一個社團）
{
    const s = loadSite();
    await uploadSheets(s, [
        { name: '社團-桌遊社A', rows: [H3, ['501', '4', '測試生501_04'], ['', '', '測試生新生_05']] },
        { name: '社團-桌遊社B', rows: [H3, ['501', '4', '測試生501_04'], ['502', '6', '測試生502_06']] },
    ]);
    setMapping(s, 'multi', { '社團-桌遊社A': { day: '週二' }, '社團-桌遊社B': { day: '週二' } });
    s.el('include-freshmen').checked = true;
    s.el('slot-mode-select').value = 'vacation';
    checkStudents('多工作表（寒暑假、含新生）', runMapping(s).resultData, [
        st('501', '4', '測試生501_04', { 週二: '桌遊社(A), 桌遊社(B)' }, [conflict('週二', '桌遊社(A), 桌遊社(B)')]),
        st('新生', '', '測試生新生_05', { 週二: '桌遊社(A)' }),
        st('502', '6', '測試生502_06', { 週二: '桌遊社(B)' }),
    ]);
    s.el('include-freshmen').checked = false;
    s.el('slot-mode-select').value = 'semester';
    checkStudents('多工作表（學期、不含新生）', runMapping(s).resultData, [
        st('501', '4', '測試生501_04', { 週二: '桌遊社' }),
        st('502', '6', '測試生502_06', { 週二: '桌遊社' }),
    ]);
}

// ============================================================
// 4. 晨間時段：同一個晨間欄位有多個社團不算衝突
// ============================================================
{
    const s = loadSite();
    await uploadSheets(s, [{ name: '總表', rows: [
        ['班級', '座號', '姓名', '社團', '晨間社團'],
        ['401', '1', '測試生401_01', '直笛社', '晨間扯鈴'],
        ['401', '1', '測試生401_01', '桌遊社', '晨間跳繩'],
        ['401', '2', '測試生401_02', '直笛社', '晨間扯鈴'],
    ] }]);
    const MT = '晨間(週一、週四)';
    setMapping(s, 'single', {
        '社團|直笛社': { day: '週一' }, '社團|桌遊社': { day: '週一' },
        '晨間社團|晨間扯鈴': { day: MT }, '晨間社團|晨間跳繩': { day: MT },
    });
    const lib = loadRealExcelJS(s);
    const res = runMapping(s);
    const expected = [
        st('401', '1', '測試生401_01', { 週一: '直笛社, 桌遊社', [MT]: '晨間扯鈴, 晨間跳繩' }, [conflict('週一', '直笛社, 桌遊社')]),
        st('401', '2', '測試生401_02', { 週一: '直笛社', [MT]: '晨間扯鈴' }),
    ];
    checkStudents('晨間（單一清單）', res.resultData, expected);
    R.check('晨間（單一清單）：晨間兩個社團不算衝突，只有週一的衝突一則', res.resultData[0]?.remarks.length === 1, JSON.stringify(res.resultData[0]?.remarks));
    const r = await clickConvert(s, lib);
    if (r.file) checkExport('匯出檔（晨間欄位排在週五之後）', r.file, expected, [...WEEKDAYS, MT], res.resultData);
    else R.check('匯出檔（晨間欄位）：匯出成功', false, r.alerts.join(' / '));
}
{
    const s = loadSite();
    const finalize = s.fn('finalizeStudents');
    const run = (days) => finalize(new Map([['k', { class: '401', seat: '1', name: '測試生401_01', days }]])).resultData[0];
    const ALL = '晨間(週一、週二、週四、週五)';
    R.test('finalizeStudents：晨間(週二、週五) 兩個社團、週五一個社團 → 沒有衝突', () => {
        const r = run({ '晨間(週二、週五)': ['晨間扯鈴', '晨間跳繩'], 週五: ['足球社'] });
        return [r.remarks.length === 0 && r.schedule['晨間(週二、週五)'] === '晨間扯鈴, 晨間跳繩' && r.schedule['週五'] === '足球社', canon(r)];
    });
    R.test(`finalizeStudents：${ALL} 三個社團 → 沒有衝突`, () => {
        const r = run({ [ALL]: ['晨間扯鈴', '晨間跳繩', '晨間游泳'] });
        return [r.remarks.length === 0 && r.schedule[ALL] === '晨間扯鈴, 晨間跳繩, 晨間游泳', canon(r)];
    });
    R.test('finalizeStudents：週五兩個社團 → 一則衝突（晨間兩個社團不算）', () => {
        const r = run({ 週五: ['足球社', '籃球社'], '晨間(週二、週五)': ['晨間扯鈴', '晨間跳繩'] });
        return [JSON.stringify(r.remarks) === JSON.stringify([conflict('週五', '足球社, 籃球社')]), JSON.stringify(r.remarks)];
    });
    R.test('finalizeStudents：週六、週日也算衝突，衝突備註依星期順序（週一 → 週日）', () => {
        const r = run({ 週日: ['甲社', '乙社'], [ALL]: ['晨間扯鈴'], 週一: ['丙社', '丁社'], 週六: ['戊社', '己社'] });
        const keys = Object.keys(r.schedule).join(',');
        const remarks = [conflict('週一', '丙社, 丁社'), conflict('週六', '戊社, 己社'), conflict('週日', '甲社, 乙社')];
        return [keys === `週一,${ALL},週六,週日` && JSON.stringify(r.remarks) === JSON.stringify(remarks), `${keys}；${JSON.stringify(r.remarks)}`];
    });
}

// ============================================================
// 3. 直接呼叫 processExcelData 與 exportWeeklySchedule：週六、週日、多個晨間欄位的表頭順序
// ============================================================
{
    const s = loadSite({ withExcelJS: false });
    const lib = loadRealExcelJS(s);
    const key = s.fn('mappingKey');
    const map = (slot, club, day) => [key(slot, club), { editedName: club, day, slot: '' }];
    const out = s.fn('processExcelData')({
        mode: 'single', includeFreshmen: false, slotMode: 'semester',
        finalMapping: Object.fromEntries([
            map('社團', '書法社', '週六'), map('社團', '圍棋社', '週日'),
            map('晨間社團', '晨間扯鈴', '晨間(週一、週二、週四、週五)'), map('晨間社團', '晨間跳繩', '晨間(週一、週四)'),
        ]),
        sheetData: [
            { 班級: '601', 座號: '1', 姓名: '測試生601_01', 社團: '書法社', 晨間社團: '晨間扯鈴' },
            { 班級: '102', 座號: '3', 姓名: '測試生102_03', 社團: '圍棋社', 晨間社團: '晨間跳繩' },
            { 班級: '', 座號: '', 姓名: '測試生新生_03', 社團: '書法社', 晨間社團: '' },
        ],
        classCol: '班級', seatCol: '座號', nameCol: '姓名',
        slotCols: [{ name: '社團', index: 3 }, { name: '晨間社團', index: 4 }],
        sheetStudents: new Map(),
    });
    const expected = [
        st('601', '1', '測試生601_01', { '晨間(週一、週二、週四、週五)': '晨間扯鈴', 週六: '書法社' }),
        st('102', '3', '測試生102_03', { '晨間(週一、週四)': '晨間跳繩', 週日: '圍棋社' }),
    ];
    R.check('processExcelData（直接呼叫、不含新生）：結果依名冊順序、內容正確',
        JSON.stringify(Array.from(out.resultData, canon)) === JSON.stringify(expected.map(canon)), Array.from(out.resultData, canon).join(' / '));
    let dl = null;
    s.ctx.triggerFileDownload = (blob, name) => { dl = { blob, name }; };
    await pump(s, s.fn('exportWeeklySchedule')(out.resultData, out.activeDays, 'semester', '115.2學期'));
    R.check('exportWeeklySchedule：檔名保留「115.2學期」並加上「_學期社團彙整.xlsx」', dl?.name === '115.2學期_學期社團彙整.xlsx', dl?.name);
    if (dl) {
        const file = await readExport(s, lib, dl.blob);
        const days = [...WEEKDAYS, '晨間(週一、週四)', '晨間(週一、週二、週四、週五)', '週六', '週日'];
        // exportWeeklySchedule 不排序，照傳入的順序寫出
        R.check('exportWeeklySchedule：表頭是週一到週五、晨間欄位（依選項順序）、週六、週日、備註',
            JSON.stringify(file.header) === JSON.stringify(['班級', '座號', '姓名', ...days, '備註']), file.header.join('、'));
        R.check('exportWeeklySchedule：每一列照傳入順序，與 processExcelData 的結果相同',
            JSON.stringify(file.rows) === JSON.stringify(expected.map(e => toRow(e, days))), file.rows.map(r => r.join('|')).join(' / '));
    } else {
        R.check('exportWeeklySchedule：匯出成功', false, s.alerts.join(' / '));
    }
}

R.done();
