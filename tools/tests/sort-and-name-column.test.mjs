// 新生排序與教師姓名欄：
// - getNumericSortKey：以「新生」開頭或空白的班級（isFreshmanClass 為 true）一律排到最後，
//   排序、抽籤、課表篩選的新生定義才一致；一般班級與座號的排序鍵不受影響
// - findNameColumn：「教師姓名」「指導老師」等教師欄不當成學生姓名欄
// - 多工作表模式：只有教師姓名欄、沒有學生姓名欄的課程表被略過，它的上課時間仍用來自動選星期
// 名冊都是虛構的假工作表（測試生301_05…），不讀取任何真實檔案。
import { loadSite, makeReporter, uploadSheets } from './harness.mjs';
const R = makeReporter('新生排序與教師姓名欄');

const H3 = ['班級', '座號', '姓名'];

// ============================================================
// 1. getNumericSortKey
// ============================================================
{
    const s = loadSite();
    const key = s.fn('getNumericSortKey');
    for (const v of ['新生', '新生班', '新生 ', ' 新生', '新生2班', '​新生', '', '  ', null, undefined]) {
        R.test(`getNumericSortKey(${JSON.stringify(v)}) → Infinity（新生或空白排最後）`, () => {
            const got = key(v);
            return [got === Infinity, String(got)];
        });
    }
    // 一般班級：數值與修改前相同
    const classes = [['301', 301], ['3年1班', 301], ['四年四班', 404], ['102', 102], ['5班', 5], ['資源班', 0]];
    for (const [v, want] of classes) {
        R.test(`getNumericSortKey("${v}") → ${want}（班級不受影響）`, () => {
            const got = key(v);
            return [got === want, String(got)];
        });
    }
    // 座號：數字照數值，空白排最後
    const seats = [['5', 5], ['05', 5], ['23', 23], [7, 7], ['', Infinity]];
    for (const [v, want] of seats) {
        R.test(`getNumericSortKey(${JSON.stringify(v)}) → ${want}（座號不受影響）`, () => {
            const got = key(v);
            return [got === want, String(got)];
        });
    }
    // 與 isFreshmanClass 一致
    const isFreshman = s.fn('isFreshmanClass');
    R.test('isFreshmanClass 為 true 的班級，排序鍵都是 Infinity', () => {
        const vals = ['新生', '新生班', '新生 ', '新生（甲）', '', '　'];
        const bad = vals.filter(v => isFreshman(v) && key(v) !== Infinity);
        return [bad.length === 0, JSON.stringify(bad)];
    });

    // 實際排序：「新生班」排在一般班級之後
    R.test('compareClassSeatName：「新生班」排在 301、302 之後', () => {
        const list = [
            { class: '新生班', seat: '1', name: '測試生新生_01' },
            { class: '302', seat: '2', name: '測試生302_02' },
            { class: '301', seat: '5', name: '測試生301_05' },
        ].sort(s.fn('compareClassSeatName'));
        const got = list.map(r => r.class).join(',');
        return [got === '301,302,新生班', got];
    });
}

// ============================================================
// 2. findNameColumn
// ============================================================
{
    const s = loadSite();
    const find = (headers) => s.fn('findNameColumn')(headers);
    const cases = [
        [['課程名稱', '上課時間', '教師姓名'], -1, '只有「教師姓名」→ -1'],
        [['班級', '座號', '導師姓名'], -1, '只有「導師姓名」→ -1'],
        [['指導老師姓名', '上課時間'], -1, '只有「指導老師姓名」→ -1'],
        [['講師名', '師資名單'], -1, '只有「講師名」「師資名單」→ -1'],
        [['教師姓名', '班級', '姓名'], 2, '「姓名」+「教師姓名」→ 姓名欄'],
        [['班級', '導師姓名', '座號', '姓名'], 3, '總表「導師姓名」在「姓名」之前 → 姓名欄'],
        [['學生姓名', '指導老師'], 0, '「學生姓名」+「指導老師」→ 學生姓名欄'],
        [['指導老師', '班級', '學生姓名'], 2, '「指導老師」在前 →「學生姓名」欄'],
        [['授課老師', '學生 姓名'], 1, '「授課老師」+「學生 姓名」（模糊比對）→ 學生欄'],
        [['班級', '座號', 'Name'], 2, '英文「Name」仍可用'],
        [['Teacher Name', 'Class'], -1, '只有「Teacher Name」→ -1'],
        [['講師姓名', '學生姓名'], 1, '「講師姓名」在前 →「學生姓名」欄'],
    ];
    for (const [headers, want, title] of cases) {
        R.test(`findNameColumn：${title}`, () => {
            const got = find(headers);
            return [got === want, `${JSON.stringify(headers)} → ${got}`];
        });
    }
}

// ============================================================
// 3. 多工作表模式：只有教師姓名欄的課程表被略過
// ============================================================
{
    const s = loadSite();
    await uploadSheets(s, [
        { name: '1161', rows: [['編號', '課程名稱', '上課時間', '教師姓名'], ['1', '桌遊社', '週三', '假老師甲'], ['2', '直笛社', '週五', '假老師乙']] },
        { name: '社團-桌遊社', rows: [H3, ['301', '5', '測試生301_05']] },
        { name: '社團-直笛社', rows: [H3, ['302', '11', '測試生302_11']] },
    ]);
    const trs = s.el('mapping-table-body').querySelectorAll('tr');
    const keys = trs.map(tr => tr.dataset.original);
    R.check('教師姓名課程表：進入第 2 步、沒有提示',
        s.el('step-mapping').classList.contains('active') && s.alerts.length === 0, s.alerts.join(' / '));
    R.check('教師姓名課程表：不出現在對照表', keys.join(',') === '社團-桌遊社,社團-直笛社', keys.join(','));
    const skip = s.el('mapping-skip-alert');
    R.check('教師姓名課程表：不列入「找不到姓名欄」的提醒', skip.hidden !== false && !skip.innerHTML.includes('1161'), skip.innerHTML.slice(0, 80));
    const days = trs.map(tr => tr.querySelector('.day-select').children.find(o => o.selected)?.value);
    R.check('教師姓名課程表：上課時間仍用來自動選星期', days.join(',') === '週三,週五', days.join(','));

    let res = null;
    try { res = s.run('runExcelProcessing(collectMapping().finalMapping)'); } catch (e) { res = { error: e.message }; }
    const names = res && res.resultData ? Array.from(res.resultData, r => r.name).sort() : [];
    R.check('教師姓名課程表：老師名字不出現在課表',
        names.length === 2 && names.join(',') === '測試生301_05,測試生302_11' && !names.some(n => n.startsWith('假老師')),
        res && res.error ? `例外：${res.error}` : names.join(','));
}

// ---- 單一清單總表：「導師姓名」在「姓名」之前時，學生仍讀自「姓名」欄
{
    const s = loadSite();
    await uploadSheets(s, [
        { name: '總表', rows: [['班級', '座號', '導師姓名', '姓名', '社團'], ['301', '5', '假老師甲', '測試生301_05', '桌遊社']] },
    ]);
    let res = null;
    try { res = s.run('runExcelProcessing(collectMapping().finalMapping)'); } catch (e) { res = { error: e.message }; }
    const names = res && res.resultData ? Array.from(res.resultData, r => r.name) : [];
    R.check('單一清單總表：有「導師姓名」時仍讀「姓名」欄', names.join(',') === '測試生301_05',
        res && res.error ? `例外：${res.error}` : names.join(',') + ' / ' + s.alerts.join(' / '));
}

R.done();
