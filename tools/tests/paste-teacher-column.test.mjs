// 貼上名單的教師欄：「導師姓名」「指導老師」「Teacher」等欄不當成學生姓名欄
// （與 Excel 姓名欄 findNameColumn 的 NAME_COLUMN_EXCLUDE_RE 一致）。名單都是虛構的（測試生…）。
import { loadSite, makeReporter } from './harness.mjs';
const R = makeReporter('貼上名單的教師欄');

const s = loadSite();
const parse = s.fn('parsePastedText');
const names = (text) => parse(text).map(x => x.name).join(',');
const TAB = String.fromCharCode(9);
const NL = String.fromCharCode(10);
const table = (rows) => rows.map(r => r.join(TAB)).join(NL);
const STUDENTS = [['301', '05', '假老師甲', '測試生301_05'], ['302', '11', '假老師乙', '測試生302_11']];

const cases = [
    ['導師姓名在學生姓名前面', ['班級', '座號', '導師姓名', '姓名']],
    ['導師姓名與中文姓名（同為 1 分）', ['班級', '座號', '導師姓名', '中文姓名']],
    ['指導老師在前', ['班級', '座號', '指導老師', '學生姓名']],
    ['Teacher Name 在前', ['班級', '座號', 'Teacher Name', 'Name']],
];
for (const [label, header] of cases) {
    R.test(`${label}：讀到的是學生不是老師`, () => {
        const got = names(table([header, ...STUDENTS]));
        return [got === '測試生301_05,測試生302_11', got];
    });
}

// 只有導師姓名欄時，這一列不算學生名單的欄位標題列（沒有學生姓名欄）。
// 之後會退回「沒有標題列」的推測（只看資料像不像姓名），那是貼上沒有標題列時的既有行為，不在這裡檢查
R.test('只有導師姓名欄：不當成欄位標題列（不會選導師姓名當姓名欄）', () => {
    const got = s.fn('matchPasteHeaderRow')(['班級', '座號', '導師姓名']);
    return [got === null || got.colName === -1, JSON.stringify(got)];
});

R.test('一般名單不受影響', () => {
    const got = names(table([['班級', '座號', '姓名'], ['301', '05', '測試生301_05'], ['302', '11', '測試生302_11']]));
    return [got === '測試生301_05,測試生302_11', got];
});

R.done();
