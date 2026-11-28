// 社團名單表頭的嚴格判斷（多工作表模式），原則是「表頭判斷要嚴格、寧可提醒不可靜默」：
// - 找不到班級、座號、姓名兩個字詞的表頭時，只接受「像欄位標題」的列：整格是姓名欄名稱（姓名、學生姓名…），
//   或含「姓名」且其他格都是常見欄位名稱；而且沒有長文字、數字或星期值。
//   資料列裡的「Bring name tag」「Name Art」「Mr. Name」「衣服繡姓名」「rename」「English name Amy」都不算表頭
// - 課程表的人員欄（教練、助教、志工姓名）不是學生姓名欄，Excel 與貼上路徑都一樣
// - 資訊列「社團｜桌遊社｜上課時間｜週三」下方是「班級｜座號｜錄取名單」或「班級｜座號｜學生」時：
//   名單照常讀到，或列入「找不到姓名欄」提醒，不可靜默消失；也不可被當成課程表讀進上課時間
// 名冊都是虛構的假工作表（測試生301_05、王小明、假老師甲…），不讀取任何真實檔案。
import { loadSite, makeReporter, makeWorkbook, uploadSheets } from './harness.mjs';
const R = makeReporter('社團名單表頭的嚴格判斷');

const H3 = ['班級', '座號', '姓名'];
const runResult = (s) => {
    try { return s.run('runExcelProcessing(collectMapping().finalMapping)'); } catch (e) { return { error: e.message }; }
};
const resultNames = (res) => (res && res.resultData ? Array.from(res.resultData, r => r.name).sort().join(',') : (res && res.error ? `例外：${res.error}` : ''));
const mappingKeys = (s) => s.el('mapping-table-body').querySelectorAll('tr').map(tr => tr.dataset.original).join(',');
const selectedDays = (s) => s.el('mapping-table-body').querySelectorAll('tr')
    .map(tr => tr.querySelector('.day-select').children.find(o => o.selected)?.value).join(',');
const skipBox = (s) => s.el('mapping-skip-alert');
const skipped = (s, name) => skipBox(s).hidden === false && skipBox(s).innerHTML.includes(name);

// ============================================================
// H1. 課程表的資料列含 name 或「姓名」：不可被當成名單表頭
// ============================================================
const SCHED_HEADER = ['課程名稱', '上課時間', '備註'];
const H1_SCHEDULES = [
    ['備註「Bring name tag」', [SCHED_HEADER, ['桌遊社', '週三', 'Bring name tag'], ['直笛社', '週五', '']], ['桌遊社', '直笛社']],
    ['課程名「Name Art」', [SCHED_HEADER, ['Name Art', '週三', ''], ['直笛社', '週五', '']], ['Name Art', '直笛社']],
    ['備註「Mr. Name」', [SCHED_HEADER, ['桌遊社', '週三', 'Mr. Name'], ['直笛社', '週五', '']], ['桌遊社', '直笛社']],
    ['備註「衣服繡姓名」', [SCHED_HEADER, ['桌遊社', '週三', '衣服繡姓名'], ['直笛社', '週五', '']], ['桌遊社', '直笛社']],
    ['備註「rename」', [SCHED_HEADER, ['桌遊社', '週三', 'rename'], ['直笛社', '週五', '']], ['桌遊社', '直笛社']],
    ['備註「衣服繡姓名」在每一列', [SCHED_HEADER, ['桌遊社', '週三', '衣服繡姓名'], ['直笛社', '週五', '衣服繡姓名'], ['書法社', '週四', '']], ['桌遊社', '直笛社']],
];
for (const [label, rows, [clubA, clubB]] of H1_SCHEDULES) {
    const s = loadSite();
    await uploadSheets(s, [
        { name: '1161', rows },
        { name: clubA, rows: [H3, ['301', '5', '測試生301_05']] },
        { name: clubB, rows: [H3, ['302', '11', '測試生302_11']] },
    ]);
    R.check(`課程表${label}：不出現在對照表、沒有提示`,
        mappingKeys(s) === `${clubA},${clubB}` && s.alerts.length === 0, `${mappingKeys(s)} / ${s.alerts.join(' / ')}`);
    R.check(`課程表${label}：不列入「找不到姓名欄」的提醒`, !skipped(s, '1161'), skipBox(s).innerHTML.slice(0, 80));
    R.check(`課程表${label}：上課時間仍用來自動選星期`, selectedDays(s) === '週三,週五', selectedDays(s));
    const res = runResult(s);
    R.check(`課程表${label}：課表只有學生`, resultNames(res) === '測試生301_05,測試生302_11', resultNames(res));

    const s2 = loadSite();
    s2.ctx.__wb = makeWorkbook([{ name: '1161', rows }]);
    const m = s2.run('loadScheduleFromExcel(__wb)') || {};
    R.check(`課程表${label}：loadScheduleFromExcel 仍讀到上課時間`, m[clubA] === '週三' && m[clubB] === '週五',
        `${String(m[clubA])} / ${String(m[clubB])}`);
}

// ---- 名單「錄取名單｜備註」、某位學生的備註含 name 或「姓名」：那一列不可被當成表頭，學生全部讀到
const H1_ROSTERS = [
    ['備註「English name Amy」', 'English name Amy'],
    ['備註「衣服繡姓名」', '衣服繡姓名'],
    ['備註「Name tag」', 'Name tag'],
];
for (const [label, note] of H1_ROSTERS) {
    const rows = [['錄取名單', '備註'], ['王小明', ''], ['李小華', note], ['陳小美', '']];
    const want = '李小華,王小明,陳小美'.split(',').sort().join(',');
    const s = loadSite();
    await uploadSheets(s, [
        { name: '社團-桌遊社', rows },
        { name: '社團-直笛社', rows: [H3, ['302', '11', '測試生302_11']] },
    ]);
    const res = runResult(s);
    R.check(`名單「錄取名單｜備註」＋${label}：學生全部讀到`,
        mappingKeys(s) === '社團-桌遊社,社團-直笛社' && resultNames(res) === [want, '測試生302_11'].join(',').split(',').sort().join(','),
        `${mappingKeys(s)} / ${resultNames(res)} / ${s.alerts.join(' / ')}`);

    const s1 = loadSite();
    await uploadSheets(s1, [{ name: '社團-桌遊社', rows }]);
    const res1 = runResult(s1);
    R.check(`名單「錄取名單｜備註」＋${label}（只有這一張）：學生全部讀到、沒有提示`,
        s1.alerts.length === 0 && resultNames(res1) === want, `${resultNames(res1)} / ${s1.alerts.join(' / ')}`);
}

// ---- 名單「學生｜備註」（沒有姓名欄）、某位學生的備註「衣服繡姓名」：那一列不可被當成表頭，列入提醒
{
    const s = loadSite();
    await uploadSheets(s, [
        { name: '社團-桌遊社', rows: [['學生', '備註'], ['王小明', ''], ['李小華', '衣服繡姓名'], ['陳小美', '']] },
        { name: '社團-直笛社', rows: [H3, ['302', '11', '測試生302_11']] },
    ]);
    const res = runResult(s);
    R.check('名單「學生｜備註」＋備註「衣服繡姓名」：列入「找不到姓名欄」的提醒、備註不當成學生',
        skipped(s, '社團-桌遊社') && resultNames(res) === '測試生302_11', `${mappingKeys(s)} / ${resultNames(res)}`);
}

// ---- 資訊列下方「中文姓名｜備註」（含「姓名」、另一格是常見欄位名稱）：仍是表頭，學生讀到
{
    const s = loadSite();
    await uploadSheets(s, [
        { name: '社團-桌遊社', rows: [['社團', '桌遊社', '上課時間', '週三'], ['中文姓名', '備註'], ['王小明', ''], ['李小華', '']] },
    ]);
    const res = runResult(s);
    R.check('資訊列＋「中文姓名｜備註」：學生讀到、沒有提示',
        s.alerts.length === 0 && resultNames(res) === '李小華,王小明', `${resultNames(res)} / ${s.alerts.join(' / ')}`);
}

// ============================================================
// M1. 課程表的教練、助教、志工姓名欄不是學生姓名欄
// ============================================================
const STAFF = ['教練姓名', '外聘教練姓名', '助教姓名', '志工姓名'];
for (const staff of STAFF) {
    const s = loadSite();
    await uploadSheets(s, [
        { name: '1161', rows: [['課程名稱', '上課時間', staff], ['桌遊社', '週三', '假老師甲'], ['直笛社', '週五', '假老師乙']] },
        { name: '社團-桌遊社', rows: [H3, ['301', '5', '測試生301_05']] },
        { name: '社團-直笛社', rows: [H3, ['302', '11', '測試生302_11']] },
    ]);
    R.check(`課程表有「${staff}」欄：不出現在對照表、沒有提示`,
        mappingKeys(s) === '社團-桌遊社,社團-直笛社' && s.alerts.length === 0, `${mappingKeys(s)} / ${s.alerts.join(' / ')}`);
    R.check(`課程表有「${staff}」欄：上課時間仍用來自動選星期`, selectedDays(s) === '週三,週五', selectedDays(s));
    const res = runResult(s);
    R.check(`課程表有「${staff}」欄：課表只有學生`, resultNames(res) === '測試生301_05,測試生302_11', resultNames(res));
}
{
    const s = loadSite();
    R.test('findNameColumn：只有教練、助教、志工姓名欄 → -1', () => {
        const got = s.fn('findNameColumn')(STAFF);
        return [got === -1, String(got)];
    });
    R.test('findNameColumn：教練姓名在學生欄前面 → 選學生欄', () => {
        const got = s.fn('findNameColumn')(['教練姓名', '中文姓名']);
        return [got === 1, String(got)];
    });
    R.test('貼上：教練、助教、志工姓名欄不是學生姓名欄', () => {
        const got = STAFF.map(c => s.fn('pasteHeaderCellKinds')(c).length);
        return [got.every(n => n === 0), got.join(',')];
    });
    R.test('貼上：標題列「班級｜座號｜教練姓名｜姓名」的姓名欄是「姓名」', () => {
        const got = s.run(`matchPasteHeaderRow(${JSON.stringify(['班級', '座號', '教練姓名', '姓名'])})`);
        return [got && got.colName === 3, JSON.stringify(got)];
    });
}

// ============================================================
// M2. 資訊列下方是「班級｜座號｜錄取名單」或「班級｜座號｜學生」：不可靜默消失、不可當成課程表
// ============================================================
const INFO = ['社團', '桌遊社', '上課時間', '週三'];
const M2_CASES = [
    // 「錄取名單」只是 1 分，但下方是班級、座號表頭的名單：照常讀學生
    ['「班級｜座號｜錄取名單」', [INFO, ['班級', '座號', '錄取名單'], ['301', '5', '測試生301_05']], 'read'],
    // 沒有姓名欄：列入「找不到姓名欄」的提醒
    ['「班級｜座號｜學生」', [INFO, ['班級', '座號', '學生'], ['301', '5', '測試生301_05']], 'skip'],
];
for (const [label, rows, expect] of M2_CASES) {
    const s = loadSite();
    await uploadSheets(s, [
        { name: '社團-桌遊社', rows },
        { name: '社團-直笛社', rows: [H3, ['302', '11', '測試生302_11']] },
    ]);
    const res = runResult(s);
    if (expect === 'read') {
        R.check(`資訊列＋${label}：列入對照表、學生讀到`,
            mappingKeys(s) === '社團-桌遊社,社團-直笛社' && resultNames(res) === '測試生301_05,測試生302_11',
            `${mappingKeys(s)} / ${resultNames(res)} / ${s.alerts.join(' / ')}`);
        R.check(`資訊列＋${label}：星期取自資訊列`, selectedDays(s).startsWith('週三,'), selectedDays(s));
    } else {
        R.check(`資訊列＋${label}：列入「找不到姓名欄」的提醒`, skipped(s, '社團-桌遊社'),
            `${skipBox(s).hidden} / ${skipBox(s).innerHTML.slice(0, 80)} / ${mappingKeys(s)}`);
    }

    // 只有這一張工作表
    const s1 = loadSite();
    await uploadSheets(s1, [{ name: '社團-桌遊社', rows }]);
    const msg = s1.alerts.join('\n');
    if (expect === 'read') {
        R.check(`資訊列＋${label}（只有這一張）：學生讀到、沒有提示`,
            s1.alerts.length === 0 && resultNames(runResult(s1)) === '測試生301_05', `${mappingKeys(s1)} / ${msg}`);
    } else {
        R.check(`資訊列＋${label}（只有這一張）：訊息是「找不到姓名欄」，不說是課程表`,
            msg.includes('找不到「姓名」欄') && !msg.includes('課程表'), msg);
    }

    const s2 = loadSite();
    s2.ctx.__wb = makeWorkbook([{ name: '社團-桌遊社', rows }]);
    const m = s2.run('loadScheduleFromExcel(__wb)') || {};
    const keys = Object.keys(m);
    const values = Object.values(m);
    R.check(`資訊列＋${label}：loadScheduleFromExcel 不寫入表頭或學生`,
        !keys.some(k => k === '301' || k === '班級') && !values.some(v => String(v).startsWith('測試生') || v === '錄取名單' || v === '學生'),
        `鍵 ${keys.length} 個，表頭或班級 ${keys.filter(k => k === '301' || k === '班級').length} 個`);
    R.check(`資訊列＋${label}：loadScheduleFromExcel 仍讀到資訊列的上課時間`, m['桌遊社'] === '週三', String(m['桌遊社']));
}

R.done();
