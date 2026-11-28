// 姓名欄的模糊比對：只含單字「名」的「名額」「報名人數」「名次」「排名」「班級名稱」等欄不當成學生姓名欄
// - findNameColumn：這類欄位不選；「中文姓名」「名字」「學生名」「Name」「報名學生姓名」仍選得到
// - 多工作表模式：課程表只有「課程名稱、上課時間、名額」時被略過，名額的數字不會被當成學生
// - 貼上名單：「名額」欄不會被當成姓名欄
// 名冊都是虛構的假工作表（測試生301_05…），不讀取任何真實檔案。
import { loadSite, makeReporter, uploadSheets } from './harness.mjs';
const R = makeReporter('姓名欄模糊比對');

const H3 = ['班級', '座號', '姓名'];

// ============================================================
// 1. findNameColumn
// ============================================================
{
    const s = loadSite();
    const find = (headers) => s.fn('findNameColumn')(headers);
    const cases = [
        [['名額'], -1, '只有「名額」→ -1'],
        [['報名人數'], -1, '只有「報名人數」→ -1'],
        [['錄取名額'], -1, '只有「錄取名額」→ -1'],
        [['名次'], -1, '只有「名次」→ -1'],
        [['排名'], -1, '只有「排名」→ -1'],
        [['報名日期', '報名序號'], -1, '只有「報名日期」「報名序號」→ -1'],
        [['班級名稱'], -1, '只有「班級名稱」→ -1'],
        [['課程名稱', '上課時間', '名額'], -1, '課程表「課程名稱、上課時間、名額」→ -1'],
        [['姓名', '名額'], 0, '「姓名」+「名額」→ 姓名欄'],
        [['名額', '姓名'], 1, '「名額」在「姓名」之前 → 姓名欄'],
        [['錄取名額', '學生 姓名'], 1, '「錄取名額」在前 +「學生 姓名」（模糊比對）→ 學生欄'],
        [['名次', '中文姓名'], 1, '「名次」在前 +「中文姓名」→ 中文姓名欄'],
        [['中文姓名'], 0, '「中文姓名」仍選得到'],
        [['姓名(中文)'], 0, '「姓名(中文)」仍選得到'],
        [['名字'], 0, '「名字」仍選得到'],
        [['學生名'], 0, '「學生名」仍選得到'],
        [['Name'], 0, '「Name」仍選得到'],
        [['Student Name'], 0, '「Student Name」仍選得到'],
        [['報名人數', '報名學生姓名'], 1, '「報名學生姓名」仍選得到（含姓名）'],
        [['錄取名單'], 0, '「錄取名單」仍選得到（欄內是學生名字）'],
    ];
    for (const [headers, want, title] of cases) {
        R.test(`findNameColumn：${title}`, () => {
            const got = find(headers);
            return [got === want, `${JSON.stringify(headers)} → ${got}`];
        });
    }
}

// ============================================================
// 2. 多工作表模式：只有「名額」欄、沒有學生姓名欄的課程表被略過
// ============================================================
{
    const s = loadSite();
    await uploadSheets(s, [
        { name: '1161', rows: [['課程名稱', '上課時間', '名額'], ['桌遊社', '週三', '20'], ['直笛社', '週五', '15']] },
        { name: '社團-桌遊社', rows: [H3, ['301', '5', '測試生301_05']] },
        { name: '社團-直笛社', rows: [H3, ['302', '11', '測試生302_11']] },
    ]);
    const trs = s.el('mapping-table-body').querySelectorAll('tr');
    const keys = trs.map(tr => tr.dataset.original);
    R.check('名額課程表：進入第 2 步、沒有提示',
        s.el('step-mapping').classList.contains('active') && s.alerts.length === 0, s.alerts.join(' / '));
    R.check('名額課程表：不出現在對照表', keys.join(',') === '社團-桌遊社,社團-直笛社', keys.join(','));
    const skip = s.el('mapping-skip-alert');
    R.check('名額課程表：不列入「找不到姓名欄」的提醒', skip.hidden !== false && !skip.innerHTML.includes('1161'), skip.innerHTML.slice(0, 80));
    const days = trs.map(tr => tr.querySelector('.day-select').children.find(o => o.selected)?.value);
    R.check('名額課程表：上課時間仍用來自動選星期', days.join(',') === '週三,週五', days.join(','));

    let res = null;
    try { res = s.run('runExcelProcessing(collectMapping().finalMapping)'); } catch (e) { res = { error: e.message }; }
    const names = res && res.resultData ? Array.from(res.resultData, r => r.name).sort() : [];
    R.check('名額課程表：名額的數字不出現在課表',
        names.length === 2 && names.join(',') === '測試生301_05,測試生302_11' && !names.some(n => /^\d+$/.test(n)),
        res && res.error ? `例外：${res.error}` : names.join(','));
}

// ---- 單一清單總表：「錄取名額」在「學生 姓名」之前時，學生仍讀自姓名欄
{
    const s = loadSite();
    await uploadSheets(s, [
        { name: '總表', rows: [['班級', '座號', '錄取名額', '學生 姓名', '社團'], ['301', '5', '20', '測試生301_05', '桌遊社']] },
    ]);
    let res = null;
    try { res = s.run('runExcelProcessing(collectMapping().finalMapping)'); } catch (e) { res = { error: e.message }; }
    const names = res && res.resultData ? Array.from(res.resultData, r => r.name) : [];
    R.check('單一清單總表：有「錄取名額」時仍讀「學生 姓名」欄', names.join(',') === '測試生301_05',
        res && res.error ? `例外：${res.error}` : names.join(',') + ' / ' + s.alerts.join(' / '));
}

// ============================================================
// 3. 貼上名單：「名額」欄不當成姓名欄
// ============================================================
{
    const s = loadSite();
    const parse = s.fn('parsePastedText');
    const TAB = String.fromCharCode(9);
    const NL = String.fromCharCode(10);
    const table = (rows) => rows.map(r => r.join(TAB)).join(NL);
    const STUDENTS = [['301', '05', '20', '測試生301_05'], ['302', '11', '20', '測試生302_11']];
    for (const header of [['班級', '座號', '名額', '姓名'], ['班級', '座號', '報名人數', '學生姓名'], ['班級', '座號', '名次', '中文姓名']]) {
        R.test(`貼上名單「${header.join('、')}」：讀到的是姓名欄`, () => {
            const got = parse(table([header, ...STUDENTS])).map(x => x.name).join(',');
            return [got === '測試生301_05,測試生302_11', got];
        });
    }
    R.test('貼上名單：「名額」「名次」「排名」「報名人數」不是姓名欄名稱', () => {
        const scores = s.fn('pasteHeaderCellScores');
        const bad = ['名額', '錄取名額', '名次', '排名', '報名人數'].filter(w => scores(w).name > 0);
        return [bad.length === 0, bad.join(',')];
    });
}

R.done();
