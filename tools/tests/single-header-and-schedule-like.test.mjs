// 姓名欄與表頭列的幾個邊界情況（名冊都是虛構的假工作表：測試生301_05、假老師甲…，不讀取任何真實檔案）：
// - 多工作表：「姓名｜社團｜星期」而且社團欄有兩種以上的值、沒有班級和座號欄的工作表，和課程表分不出來，不讀，
//   但要列在第 2 步上方「看起來是課程表」的提醒（只有這一張時是 alert），不可靜默消失；
//   沒有姓名欄的課程表照舊靜默略過，「Name」人員欄的課程表不可讓人員變成學生
// - 姓名欄名稱後面括號裡的說明（「姓名(導師確認)」）先去掉再判斷，括號裡的「師」「家長」不會讓姓名欄被排除
// - 單一清單：總表上方的標題或說明含「各班級」時，表頭要取下方真正的「班級｜座號｜姓名」列
import { loadSite, makeReporter, uploadSheets } from './harness.mjs';
const R = makeReporter('姓名欄括號說明、單一清單表頭與看起來是課程表的工作表');

const runResult = (s) => {
    try { return s.run('runExcelProcessing(collectMapping().finalMapping)'); } catch (e) { return { error: e.message }; }
};
const resultNames = (res) => (res && res.resultData ? Array.from(res.resultData, r => r.name).sort().join(',') : (res && res.error ? `例外：${res.error}` : ''));
const mappingKeys = (s) => s.el('mapping-table-body').querySelectorAll('tr').map(tr => tr.dataset.original).join(',');
const skipHtml = (s) => { const b = s.el('mapping-skip-alert'); return b.hidden === false ? b.innerHTML : ''; };
const STUDENTS = ['測試生301_05', '測試生302_11'];
const noStudentName = (text) => !STUDENTS.some(n => text.includes(n));
const CLUB = { name: '社團-直笛社', rows: [['班級', '座號', '姓名'], ['304', '9', '測試生304_09']] };

// ============================================================
// M1. 「姓名｜社團｜星期」、社團欄有兩種以上的值
// ============================================================
const BOARD = { name: '社團-桌遊社', rows: [['姓名', '社團', '星期'], ['測試生301_05', '桌遊社A組', '週三'], ['測試生302_11', '桌遊社B組', '週三']] };
{
    const s = loadSite();
    await uploadSheets(s, [BOARD, CLUB]);
    const html = skipHtml(s);
    R.check('M1 搭配正常社團：工作表列在「看起來是課程表」的提醒', html.includes('看起來是課程表') && html.includes('社團-桌遊社'), html.slice(0, 160));
    R.check('M1 搭配正常社團：不列在「找不到姓名欄」', !html.includes('找不到「姓名」欄'), html.slice(0, 120));
    R.check('M1 搭配正常社團：提醒說明加上班級、座號欄後重新選擇檔案', html.includes('「班級」「座號」') && html.includes('重新選擇檔案'), html.slice(-200));
    R.check('M1 搭配正常社團：提醒不含學生姓名', noStudentName(html));
    R.check('M1 搭配正常社團：正常社團照常讀', resultNames(runResult(s)) === '測試生304_09', resultNames(runResult(s)));
}
{
    const s = loadSite();
    await uploadSheets(s, [BOARD]);
    const msg = s.alerts.join('\n');
    const i = msg.indexOf('看起來是課程表');
    R.check('M1 只有這一張：alert 說明這張工作表看起來是課程表', i !== -1 && msg.indexOf('社團-桌遊社', i) !== -1, msg.slice(0, 200));
    R.check('M1 只有這一張：alert 說明加上班級、座號欄後重新選擇檔案', msg.slice(i).includes('「班級」「座號」') && msg.includes('重新選擇檔案'), msg.slice(i, i + 160));
    R.check('M1 只有這一張：alert 不說找不到姓名欄、不含學生姓名', !msg.includes('找不到「姓名」欄') && noStudentName(msg), msg.slice(0, 80));
}
{
    // 社團欄只有一種值：單一社團的名單，照常讀、沒有提醒
    const s = loadSite();
    await uploadSheets(s, [{ name: '社團-桌遊社', rows: [['姓名', '社團', '星期'], ['測試生301_05', '桌遊社', '週三'], ['測試生302_11', '桌遊社', '週三']] }]);
    R.check('社團欄只有一種值：照名單讀、沒有提醒', resultNames(runResult(s)) === STUDENTS.join(',') && !skipHtml(s) && s.alerts.length === 0,
        `${resultNames(runResult(s))} / ${s.alerts.join(' / ')}`);
}
{
    // 沒有姓名欄的課程表：照舊靜默略過
    const s = loadSite();
    await uploadSheets(s, [{ name: '1161', rows: [['課程名稱', '上課時間'], ['桌遊社', '週三'], ['直笛社', '週五']] }, CLUB]);
    R.check('沒有姓名欄的課程表：不出現在對照表與提醒', mappingKeys(s) === '社團-直笛社' && !skipHtml(s) && s.alerts.length === 0,
        `${mappingKeys(s)} / ${skipHtml(s).slice(0, 60)}`);
}
{
    // 「Name」人員欄、兩門課的課程表：列在「看起來是課程表」，人員不可變成學生
    const s = loadSite();
    await uploadSheets(s, [{ name: '1161', rows: [['課程名稱', '上課時間', 'Name'], ['桌遊社', '週三', '假老師甲'], ['直笛社', '週五', '假老師乙']] }, CLUB]);
    const html = skipHtml(s);
    R.check('「Name」人員欄的課程表：人員不是學生', resultNames(runResult(s)) === '測試生304_09', resultNames(runResult(s)));
    R.check('「Name」人員欄的課程表：不在「找不到姓名欄」，只在「看起來是課程表」', !html.includes('找不到「姓名」欄') && html.includes('看起來是課程表') && html.includes('1161'), html.slice(0, 120));
}
{
    // 兩段同時出現：找不到姓名欄的工作表在前，看起來是課程表的在後
    const s = loadSite();
    await uploadSheets(s, [BOARD, { name: '名單整理', rows: [['班級', '座號'], ['301', '05']] }, CLUB]);
    const html = skipHtml(s);
    const a = html.indexOf('找不到「姓名」欄');
    const b = html.indexOf('看起來是課程表');
    R.check('兩種提醒同時出現：各自列出工作表', a !== -1 && b > a && html.slice(a, b).includes('名單整理') && !html.slice(a, b).includes('社團-桌遊社') && html.slice(b).includes('社團-桌遊社'), html.slice(0, 200));
}

// ============================================================
// L1. 姓名欄名稱後面括號裡的說明
// ============================================================
{
    const s = loadSite();
    const score = s.fn('nameColumnScore');
    for (const h of ['姓名(導師確認)', '學生姓名（非家長）', '姓名(英文字母)']) {
        R.check(`nameColumnScore「${h}」是 3 分（先去掉括號裡的說明）`, score(h) === 3, String(score(h)));
    }
    R.check('nameColumnScore「家長姓名(備註)」仍是 0 分', score('家長姓名(備註)') === 0, String(score('家長姓名(備註)')));
    R.check('nameColumnScore「(請填寫)」只有括號時不當成姓名欄', score('(請填寫)') === 0, String(score('(請填寫)')));
}
for (const h of ['姓名(導師確認)', '學生姓名（非家長）', '姓名(英文字母)']) {
    const s = loadSite();
    await uploadSheets(s, [{ name: '總表', rows: [['班級', '座號', h, '社團'], ['301', '05', '測試生301_05', '桌遊社'], ['302', '11', '測試生302_11', '直笛社']] }]);
    R.check(`單一清單姓名欄「${h}」：讀得到學生、沒有「缺少必要欄位」`,
        resultNames(runResult(s)) === STUDENTS.join(',') && s.alerts.length === 0, `${resultNames(runResult(s))} / ${s.alerts.join(' / ').slice(0, 60)}`);
}
{
    const s = loadSite();
    await uploadSheets(s, [{ name: '社團-桌遊社', rows: [['班級', '座號', '姓名(導師確認)'], ['301', '05', '測試生301_05'], ['302', '11', '測試生302_11']] }]);
    R.check('多工作表姓名欄「姓名(導師確認)」：讀得到學生', resultNames(runResult(s)) === STUDENTS.join(','), resultNames(runResult(s)));
}

// ============================================================
// L2. 單一清單：表頭上方的標題或說明含「各班級」
// ============================================================
const SINGLE_TITLES = [
    ['標題「延平國小各班級社團錄取名單」', [['延平國小各班級社團錄取名單']]],
    ['說明「說明：各班級請於週五前繳回」', [['說明：各班級請於週五前繳回']]],
    ['標題＋說明兩列', [['各班級社團錄取名單'], ['說明：各班級請於週五前繳回']]],
    ['沒有標題', []],
];
for (const [label, above] of SINGLE_TITLES) {
    const s = loadSite();
    await uploadSheets(s, [{ name: '總表', rows: [...above, ['班級', '座號', '姓名', '社團'], ['301', '05', '測試生301_05', '桌遊社'], ['302', '11', '測試生302_11', '直笛社']] }]);
    R.check(`單一清單${label}：取下方的表頭，讀得到學生、沒有訊息`,
        resultNames(runResult(s)) === STUDENTS.join(',') && s.alerts.length === 0 && mappingKeys(s).split(',').sort().join(',') === '桌遊社,直笛社',
        `${resultNames(runResult(s))} / ${mappingKeys(s)} / ${s.alerts.join(' / ').slice(0, 60)}`);
}
{
    // 表頭的姓名欄只有 1 分（「錄取名單」），上方標題也是：仍取有整格班級、座號欄的那一列
    const s = loadSite();
    await uploadSheets(s, [{ name: '總表', rows: [['各班級錄取名單'], ['班級', '座號', '錄取名單', '社團'], ['301', '05', '測試生301_05', '桌遊社'], ['302', '11', '測試生302_11', '直笛社']] }]);
    R.check('單一清單標題「各班級錄取名單」＋表頭「班級｜座號｜錄取名單｜社團」：讀得到學生',
        resultNames(runResult(s)) === STUDENTS.join(',') && s.alerts.length === 0, `${resultNames(runResult(s))} / ${s.alerts.join(' / ').slice(0, 60)}`);
}
{
    // findHeaderRow 不再回傳沒有人讀的 found
    const s = loadSite();
    const h = s.run('findHeaderRow({ rowCount: 1, getRow: () => ({ eachCell(o, fn) { ["班級", "座號", "姓名"].forEach((v, i) => fn({ value: v }, i + 1)); } }) })');
    R.check('findHeaderRow 回傳 { headers, rowIndex }', h.rowIndex === 1 && Array.from(h.headers).join('|') === '班級|座號|姓名' && !('found' in h), JSON.stringify(h));
}

R.done();
