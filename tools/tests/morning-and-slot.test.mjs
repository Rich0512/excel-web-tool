// 晨間社團的衝堂判斷（finalizeStudents），與寒暑假模式總表時段欄位的標籤（slotLabelFromColumn／slotLabelsFromColumns）
// - 晨間：同一個晨間欄位兩個社團、不同晨間欄位的上課日重疊都算衝堂；不重疊的晨間欄位、晨間與下午同一天不算
//   Excel 單一清單、多工作表、貼上名單（只比對正取）、匯出檔備註欄與衝堂提醒都要一致
// - 標籤：欄位名稱只剩通用字詞（「課程名稱」）時不加「(名稱)」；兩欄清成相同標籤時改用完整欄位名稱，不可合併成一筆
// 測試資料的姓名全部是虛構的（測試生301_05…）。
import vm from 'node:vm';
import { readFileSync } from 'node:fs';
import { loadSite, makeReporter, uploadSheets, EXCELJS_PATH } from './harness.mjs';
const R = makeReporter('晨間社團衝堂與寒暑假時段標籤');

const MT = '晨間(週一、週四)';
const TF = '晨間(週二、週五)';
const ALL = '晨間(週一、週二、週四、週五)';
const same = (day, clubs) => `「${day}」時段衝突：同時錄取「${clubs}」`;
const overlap = (a, b, days, clubs) => `「${a}」與「${b}」在${days}晨間衝突：同時錄取「${clubs}」`;
const J = (v) => JSON.stringify(v);
const remarksOf = (r) => Array.from(r?.remarks || []);

// ============================================================
// 1. finalizeStudents：直接給學生的 days
// ============================================================
{
    const s = loadSite();
    const finalize = s.fn('finalizeStudents');
    const run = (days) => finalize(new Map([['k', { class: '301', seat: '05', name: '測試生301_05', days }]])).resultData[0];
    const expect = (title, days, remarks) => R.test(title, () => {
        const r = run(days);
        return [J(remarksOf(r)) === J(remarks), J(remarksOf(r))];
    });

    expect('同一個晨間欄位兩個社團 → 一則衝突', { [MT]: ['晨間扯鈴', '晨間跳繩'] }, [same(MT, '晨間扯鈴, 晨間跳繩')]);
    expect('「晨間(週一、週四)」與「晨間(週一、週二、週四、週五)」重疊 → 週一、週四衝突',
        { [MT]: ['晨間扯鈴'], [ALL]: ['晨間游泳'] }, [overlap(MT, ALL, '週一、週四', '晨間扯鈴, 晨間游泳')]);
    expect('「晨間(週二、週五)」與「晨間(週一、週二、週四、週五)」重疊 → 週二、週五衝突',
        { [TF]: ['晨間跳繩'], [ALL]: ['晨間游泳'] }, [overlap(TF, ALL, '週二、週五', '晨間跳繩, 晨間游泳')]);
    expect('「晨間(週一、週四)」與「晨間(週二、週五)」不重疊 → 沒有衝突', { [MT]: ['晨間扯鈴'], [TF]: ['晨間跳繩'] }, []);
    expect('晨間與下午同一天（晨間(週一、週四) 與週一、週四）→ 沒有衝突', { 週一: ['桌遊社'], 週四: ['直笛社'], [MT]: ['晨間扯鈴'] }, []);
    expect('晨間與下午同一天，下午兩個社團 → 只有下午一則', { 週一: ['桌遊社', '直笛社'], [MT]: ['晨間扯鈴'] }, [same('週一', '桌遊社, 直笛社')]);
    expect('三個晨間欄位各一個社團 → 兩則重疊衝突（依晨間欄位順序）',
        { [ALL]: ['晨間游泳'], [TF]: ['晨間跳繩'], [MT]: ['晨間扯鈴'] },
        [overlap(MT, ALL, '週一、週四', '晨間扯鈴, 晨間游泳'), overlap(TF, ALL, '週二、週五', '晨間跳繩, 晨間游泳')]);
    expect('同名社團排在兩個重疊的晨間欄位 → 視為同一個社團，不算衝突', { [MT]: ['晨間扯鈴'], [ALL]: ['晨間扯鈴'] }, []);
    expect('晨間欄位自己有兩個社團、重疊欄位只有其中一個 → 只寫那一欄的衝突，不重複',
        { [MT]: ['晨間扯鈴', '晨間跳繩'], [ALL]: ['晨間扯鈴'] }, [same(MT, '晨間扯鈴, 晨間跳繩')]);
    expect('晨間欄位自己有兩個社團、重疊欄位另有社團 → 兩則（欄位內、欄位之間）',
        { [MT]: ['晨間扯鈴', '晨間跳繩'], [ALL]: ['晨間游泳'] },
        [same(MT, '晨間扯鈴, 晨間跳繩'), overlap(MT, ALL, '週一、週四', '晨間扯鈴, 晨間跳繩, 晨間游泳')]);
    expect('備註依星期順序：下午 → 晨間 → 週六',
        { 週六: ['甲社', '乙社'], [ALL]: ['晨間游泳'], 週三: ['丙社', '丁社'], [MT]: ['晨間扯鈴'] },
        [same('週三', '丙社, 丁社'), overlap(MT, ALL, '週一、週四', '晨間扯鈴, 晨間游泳'), same('週六', '甲社, 乙社')]);
    R.test('晨間欄位的課表內容不變（各欄位照列社團）', () => {
        const r = run({ [MT]: ['晨間扯鈴'], [ALL]: ['晨間游泳', '晨間跳繩'] });
        return [J(r.schedule) === J({ [MT]: '晨間扯鈴', [ALL]: '晨間游泳, 晨間跳繩' }), J(r.schedule)];
    });
}

// ============================================================
// 2. 下午：「週一、四」這類多日選項展開到各天，既有的衝堂判斷不變
// ============================================================
{
    const s = loadSite();
    const processPasted = s.fn('processPastedClubsData');
    const club = (clubName, day, students) => ({ id: clubName, day, clubName, quota: 12, gradeStart: 1, gradeEnd: 6, priority: 'desc', selectedCount: students.length, students });
    const A = { class: '301', seat: '05', name: '測試生301_05', selected: true };
    R.test('下午：「週一、四」與「週四」→ 只有週四衝突', () => {
        const r = processPasted([club('桌遊社', '週一、四', [{ ...A }]), club('直笛社', '週四', [{ ...A }])], 'semester', true).resultData[0];
        return [J(r.schedule) === J({ 週一: '桌遊社', 週四: '桌遊社, 直笛社' }) && J(remarksOf(r)) === J([same('週四', '桌遊社, 直笛社')]), J(r)];
    });
    R.test('下午：「週一、二、四、五」與「週二、五」→ 週二、週五各一則', () => {
        const r = processPasted([club('足球社', '週一、二、四、五', [{ ...A }]), club('圍棋社', '週二、五', [{ ...A }])], 'semester', true).resultData[0];
        return [J(remarksOf(r)) === J([same('週二', '足球社, 圍棋社'), same('週五', '足球社, 圍棋社')]), J(remarksOf(r))];
    });

    // ---- 貼上名單：晨間只比對正取
    R.test('貼上名單：兩個重疊的晨間社團都正取 → 衝突', () => {
        const r = processPasted([club('晨間扯鈴', MT, [{ ...A }]), club('晨間游泳', ALL, [{ ...A }])], 'semester', true).resultData[0];
        return [J(remarksOf(r)) === J([overlap(MT, ALL, '週一、週四', '晨間扯鈴, 晨間游泳')]), J(remarksOf(r))];
    });
    R.test('貼上名單：同一個晨間欄位兩個社團都正取 → 衝突', () => {
        const r = processPasted([club('晨間扯鈴', TF, [{ ...A }]), club('晨間跳繩', TF, [{ ...A }])], 'semester', true).resultData[0];
        return [J(remarksOf(r)) === J([same(TF, '晨間扯鈴, 晨間跳繩')]), J(remarksOf(r))];
    });
    R.test('貼上名單：其中一個晨間社團是備取 → 不算衝突', () => {
        const r = processPasted([club('晨間扯鈴', MT, [{ ...A }]), club('晨間游泳', ALL, [{ ...A, selected: false }])], 'semester', true).resultData[0];
        return [J(remarksOf(r)) === J([]) && J(r.schedule) === J({ [MT]: '晨間扯鈴' }), J(r)];
    });
    R.test('貼上名單：晨間社團與同一天的下午社團 → 不算衝突', () => {
        const r = processPasted([club('晨間扯鈴', MT, [{ ...A }]), club('桌遊社', '週一', [{ ...A }])], 'semester', true).resultData[0];
        return [J(remarksOf(r)) === J([]), J(remarksOf(r))];
    });

    // ---- 衝堂提醒：標題要涵蓋晨間
    const { resultData } = processPasted([club('晨間扯鈴', MT, [{ ...A }]), club('晨間游泳', ALL, [{ ...A }])], 'semester', true);
    s.fn('showConflictAlert')(resultData, 'paste-conflict-alert');
    const html = s.el('paste-conflict-alert').innerHTML;
    R.check('衝堂提醒：只有晨間衝堂時也會顯示', s.el('paste-conflict-alert').style.display === 'flex', s.el('paste-conflict-alert').style.display);
    R.check('衝堂提醒：標題寫「同一時段（下午或晨間）」', html.includes('有 1 位學生同一天的同一時段（下午或晨間）有兩個以上社團'), html.slice(0, 400));
    R.check('衝堂提醒：列出晨間衝突的備註', html.includes('「晨間(週一、週四)」與「晨間(週一、週二、週四、週五)」在週一、週四晨間衝突'), html.slice(-300));
}

// ============================================================
// 3. Excel 模式：單一清單、多工作表，以及匯出檔的備註欄
// ============================================================
function loadRealExcelJS(s) {
    Object.assign(s.ctx, { TextDecoder, TextEncoder, regeneratorRuntime: undefined });
    vm.runInContext(readFileSync(EXCELJS_PATH, 'utf8'), s.ctx);
    return s.ctx.ExcelJS;
}
async function pump(s, promise) {
    let done = false;
    const p = Promise.resolve(promise).finally(() => { done = true; });
    for (let i = 0; i < 20000 && !done; i++) { s.advance(5); await new Promise(r => setImmediate(r)); }
    return p;
}
{
    const s = loadSite({ withExcelJS: false });
    const lib = loadRealExcelJS(s);
    const key = s.fn('mappingKey');
    const map = (slot, club, day) => [key(slot, club), { editedName: club, day, slot: '' }];
    const single = s.fn('processExcelData')({
        mode: 'single', includeFreshmen: true, slotMode: 'semester',
        finalMapping: Object.fromEntries([
            map('社團', '桌遊社', '週一'),
            map('晨間社團1', '晨間扯鈴', MT), map('晨間社團1', '晨間跳繩', TF),
            map('晨間社團2', '晨間游泳', ALL), map('晨間社團2', '晨間籃球', MT),
        ]),
        sheetData: [
            { 班級: '301', 座號: '5', 姓名: '測試生301_05', 社團: '桌遊社', 晨間社團1: '晨間扯鈴', 晨間社團2: '晨間游泳' },   // 重疊
            { 班級: '302', 座號: '11', 姓名: '測試生302_11', 社團: '桌遊社', 晨間社團1: '晨間扯鈴', 晨間社團2: '晨間籃球' },  // 同一晨間欄
            { 班級: '303', 座號: '8', 姓名: '測試生303_08', 社團: '桌遊社', 晨間社團1: '晨間跳繩', 晨間社團2: '晨間籃球' },   // 不重疊
        ],
        classCol: '班級', seatCol: '座號', nameCol: '姓名',
        slotCols: [{ name: '社團', index: 3 }, { name: '晨間社團1', index: 4 }, { name: '晨間社團2', index: 5 }],
        sheetStudents: new Map(),
    });
    const want = {
        '測試生301_05': [overlap(MT, ALL, '週一、週四', '晨間扯鈴, 晨間游泳')],
        '測試生302_11': [same(MT, '晨間扯鈴, 晨間籃球')],
        '測試生303_08': [],
    };
    for (const r of Array.from(single.resultData)) {
        R.check(`Excel 單一清單：${r.name} 的晨間衝突備註`, J(remarksOf(r)) === J(want[r.name]), J(remarksOf(r)));
    }

    const sheetStudents = new Map([
        ['社團-晨間扯鈴', [{ class: '401', seat: '1', name: '測試生401_01' }]],
        ['社團-晨間游泳', [{ class: '401', seat: '1', name: '測試生401_01' }, { class: '401', seat: '2', name: '測試生401_02' }]],
        ['社團-晨間跳繩', [{ class: '401', seat: '2', name: '測試生401_02' }]],
        ['社團-晨間籃球', [{ class: '401', seat: '2', name: '測試生401_02' }]],
    ]);
    const multi = s.fn('processExcelData')({
        mode: 'multi', includeFreshmen: true, slotMode: 'semester',
        finalMapping: {
            '社團-晨間扯鈴': { editedName: '晨間扯鈴', day: MT, slot: '' },
            '社團-晨間游泳': { editedName: '晨間游泳', day: ALL, slot: '' },
            '社團-晨間跳繩': { editedName: '晨間跳繩', day: TF, slot: '' },
            '社團-晨間籃球': { editedName: '晨間籃球', day: TF, slot: '' },
        },
        sheetData: [], classCol: '', seatCol: '', nameCol: '', slotCols: [], sheetStudents,
    });
    const byName = new Map(Array.from(multi.resultData, r => [r.name, r]));
    R.check('Excel 多工作表：重疊的晨間欄位 → 衝突', J(remarksOf(byName.get('測試生401_01'))) === J([overlap(MT, ALL, '週一、週四', '晨間扯鈴, 晨間游泳')]),
        J(remarksOf(byName.get('測試生401_01'))));
    R.check('Excel 多工作表：同一晨間欄兩個社團、又與另一晨間欄重疊 → 兩則',
        J(remarksOf(byName.get('測試生401_02'))) === J([same(TF, '晨間跳繩, 晨間籃球'), overlap(TF, ALL, '週二、週五', '晨間跳繩, 晨間籃球, 晨間游泳')]),
        J(remarksOf(byName.get('測試生401_02'))));

    // 匯出檔的備註欄就是 remarks 以「；」串起來
    let dl = null;
    s.ctx.triggerFileDownload = (blob, name) => { dl = { blob, name }; };
    await pump(s, s.fn('exportWeeklySchedule')(multi.resultData, multi.activeDays, 'semester', '測試'));
    if (dl) {
        const wb = new lib.Workbook();
        await pump(s, wb.xlsx.load(dl.blob.parts[0]));
        const ws = wb.getWorksheet('課程與學生清單');
        const width = ws.getRow(1).cellCount;
        const rows = Array.from({ length: ws.rowCount }, (_, i) => Array.from({ length: width }, (_, c) => ws.getCell(i + 1, c + 1).value ?? ''));
        const header = rows[0];
        R.check('匯出檔：表頭有三個晨間欄位與備註', J(header) === J(['班級', '座號', '姓名', '週一', '週二', '週三', '週四', '週五', MT, TF, ALL, '備註']), header.join('、'));
        const remarkCol = header.indexOf('備註');
        const row = (name) => rows.find(r => r[2] === name) || [];
        R.check('匯出檔：測試生401_02 的備註欄寫出兩則晨間衝突',
            row('測試生401_02')[remarkCol] === [same(TF, '晨間跳繩, 晨間籃球'), overlap(TF, ALL, '週二、週五', '晨間跳繩, 晨間籃球, 晨間游泳')].join('；'),
            String(row('測試生401_02')[remarkCol]));
        R.check('匯出檔：測試生401_01 的備註欄寫出重疊衝突', row('測試生401_01')[remarkCol] === overlap(MT, ALL, '週一、週四', '晨間扯鈴, 晨間游泳'),
            String(row('測試生401_01')[remarkCol]));
    } else {
        R.check('匯出檔：匯出成功', false, s.alerts.join(' / '));
    }

    // 衝堂提醒（對照表下方）：三位學生中兩位有晨間衝突
    s.fn('showConflictAlert')(single.resultData, 'mapping-conflict-alert');
    R.check('衝堂提醒（Excel）：晨間衝突的學生列入人數', s.el('mapping-conflict-alert').innerHTML.includes('有 2 位學生同一天的同一時段'),
        s.el('mapping-conflict-alert').innerHTML.slice(0, 300));
}

// ============================================================
// 4. 寒暑假模式：時段欄位名稱的標籤
// ============================================================
{
    const s = loadSite();
    const label = s.fn('slotLabelFromColumn');
    const LABELS = [
        ['課程名稱', ''], ['社團名稱', ''], ['社團名單', ''], ['社團', ''], ['課程', ''], ['項目內容', ''], ['時段', ''],
        ['上午社團', '上午'], ['下午社團', '下午'], ['課程名稱(第一梯)', '第一梯'], ['第一梯社團', '第一梯'], ['社團-上午', '上午'],
        ['晨間社團', '晨間'], ['時段A', 'A'], ['第二時段', 'B'], ['社團(B)', 'B'], ['第一節', 'A'],
    ];
    for (const [col, want] of LABELS) {
        R.test(`slotLabelFromColumn：「${col}」→「${want}」`, () => [label(col) === want, J(label(col))]);
    }

    const labels = (cols) => Object.fromEntries(s.fn('slotLabelsFromColumns')(cols));
    const expectLabels = (title, cols, want) => R.test(title, () => { const got = labels(cols); return [J(got) === J(want), J(got)]; });
    expectLabels('slotLabelsFromColumns：只有「課程名稱」一欄 → 不加標籤', ['課程名稱'], { 課程名稱: '' });
    expectLabels('slotLabelsFromColumns：「社團」「課程名稱」都清成空白 → 改用完整欄位名稱', ['社團', '課程名稱'], { 社團: '社團', 課程名稱: '課程名稱' });
    expectLabels('slotLabelsFromColumns：「上午社團」「上午課程」都是「上午」→ 改用完整欄位名稱', ['上午社團', '上午課程'], { 上午社團: '上午社團', 上午課程: '上午課程' });
    expectLabels('slotLabelsFromColumns：「上午社團」「下午社團」不撞名 → 上午、下午', ['上午社團', '下午社團'], { 上午社團: '上午', 下午社團: '下午' });
    expectLabels('slotLabelsFromColumns：A/B 欄位照舊', ['時段A', '時段B'], { 時段A: 'A', 時段B: 'B' });
    expectLabels('slotLabelsFromColumns：A/B 欄位加一欄沒有標籤的欄位 → 那一欄不加', ['時段A', '時段B', '課程名稱'], { 時段A: 'A', 時段B: 'B', 課程名稱: '' });

    // processExcelData：課表文字
    const key = s.fn('mappingKey');
    const runVac = (cols, row, slotMode = 'vacation') => s.fn('processExcelData')({
        mode: 'single', includeFreshmen: true, slotMode,
        finalMapping: Object.fromEntries(cols.map(c => [key(c, row[c]), { editedName: row[c], day: '週二', slot: '' }])),
        sheetData: [{ 班級: '501', 座號: '4', 姓名: '測試生501_04', ...row }],
        classCol: '班級', seatCol: '座號', nameCol: '姓名',
        slotCols: cols.map((name, i) => ({ name, index: 3 + i })), sheetStudents: new Map(),
    }).resultData[0];
    R.test('寒暑假：只有「課程名稱」欄 → 「桌遊社」，不出現「(名稱)」', () => {
        const r = runVac(['課程名稱'], { 課程名稱: '桌遊社' });
        return [r.schedule['週二'] === '桌遊社', J(r.schedule)];
    });
    R.test('寒暑假：「社團」「課程名稱」兩欄同一天都是桌遊社 → 分開列出並算衝突', () => {
        const r = runVac(['社團', '課程名稱'], { 社團: '桌遊社', 課程名稱: '桌遊社' });
        return [r.schedule['週二'] === '桌遊社(社團), 桌遊社(課程名稱)' && J(remarksOf(r)) === J([same('週二', '桌遊社(社團), 桌遊社(課程名稱)')]), J(r)];
    });
    R.test('寒暑假：「社團名稱」「課程名稱」兩欄同一天都是桌遊社 → 不會合併成一筆', () => {
        const r = runVac(['社團名稱', '課程名稱'], { 社團名稱: '桌遊社', 課程名稱: '桌遊社' });
        return [r.schedule['週二'] === '桌遊社(社團名稱), 桌遊社(課程名稱)' && remarksOf(r).length === 1, J(r)];
    });
    R.test('寒暑假：「課程名稱(第一梯)」→「桌遊社(第一梯)」', () => {
        const r = runVac(['課程名稱(第一梯)'], { '課程名稱(第一梯)': '桌遊社' });
        return [r.schedule['週二'] === '桌遊社(第一梯)', J(r.schedule)];
    });
    R.test('學期模式：「社團」「課程名稱」兩欄同一天都是桌遊社 → 同一個社團只列一次', () => {
        const r = runVac(['社團', '課程名稱'], { 社團: '桌遊社', 課程名稱: '桌遊社' }, 'semester');
        return [r.schedule['週二'] === '桌遊社' && remarksOf(r).length === 0, J(r)];
    });
}

// ---- 從上傳走到轉換：總表只有「課程名稱」欄，寒暑假模式不出現「(名稱)」
{
    const s = loadSite();
    await uploadSheets(s, [{ name: '總表', rows: [
        ['班級', '座號', '姓名', '課程名稱'],
        ['501', '4', '測試生501_04', '桌遊社'],
    ] }]);
    s.el('mapping-table-body').querySelectorAll('tr').forEach(tr => { tr.querySelector('.day-select').value = '週三'; });
    s.el('slot-mode-select').value = 'vacation';
    const res = s.run('runExcelProcessing(collectMapping().finalMapping)');
    const r = res?.resultData?.[0];
    R.check('上傳總表（只有「課程名稱」欄）→ 寒暑假課表是「桌遊社」', r?.schedule?.['週三'] === '桌遊社', J(r?.schedule) + ' / ' + s.alerts.join(' / '));
}

R.done();
