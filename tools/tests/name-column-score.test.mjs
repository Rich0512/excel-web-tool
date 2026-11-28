// 學生姓名欄的評分選欄（findNameColumn）與多工作表的課程表略過：
// - 比對前去掉所有空白（含全形空白）並轉小寫；家長、監護人、聯絡人、社團、課程、教師等欄不選；
//   整格是姓名欄名稱（3 分）> 含「姓名」或 name（2 分）> 去掉「名額」「社名」「簽名」等詞後仍含「名」（1 分），同分取最前面
// - 學生欄刻意用不完全吻合的寫法（「姓　名」「學生 姓名」「中文姓名」），舊版「第一個含關鍵字的欄先贏」會選錯
// - 多工作表：姓名欄誤中課程表（「隊名」「Course Name」、標題列、全形空白的「名　額」）時略過課程表；
//   社團工作表上方的「社團｜桌遊社｜上課時間｜週三」資訊列不影響學生名單，也不會讓學生資料進入檔案內的上課時間
// 名冊都是虛構的假工作表（測試生301_05、假老師甲、假家長甲…），不讀取任何真實檔案。
import { loadSite, makeReporter, makeWorkbook, uploadSheets } from './harness.mjs';
const R = makeReporter('姓名欄評分與課程表略過');

const H3 = ['班級', '座號', '姓名'];
const FW = String.fromCharCode(0x3000); // 全形空白
const runResult = (s) => {
    try { return s.run('runExcelProcessing(collectMapping().finalMapping)'); } catch (e) { return { error: e.message }; }
};
const resultNames = (res) => (res && res.resultData ? Array.from(res.resultData, r => r.name).sort().join(',') : (res && res.error ? `例外：${res.error}` : ''));
const mappingKeys = (s) => s.el('mapping-table-body').querySelectorAll('tr').map(tr => tr.dataset.original).join(',');
const selectedDays = (s) => s.el('mapping-table-body').querySelectorAll('tr')
    .map(tr => tr.querySelector('.day-select').children.find(o => o.selected)?.value).join(',');

// ============================================================
// 1. findNameColumn：學生欄不完全吻合時仍選對
// ============================================================
{
    const s = loadSite();
    const find = (headers) => s.fn('findNameColumn')(headers);
    const cases = [
        [['導師姓名', '中文姓名'], 1, '「導師姓名」在前 →「中文姓名」'],
        [['家長姓名', `姓${FW}名`], 1, '「家長姓名」在前 →「姓　名」（全形空白）'],
        [['監護人姓名', '學生 姓名'], 1, '「監護人姓名」在前 →「學生 姓名」'],
        [['父親姓名', '母親姓名', '學生 姓名'], 2, '父母姓名在前 →「學生 姓名」'],
        [['緊急聯絡人姓名', '中文姓名'], 1, '「緊急聯絡人姓名」在前 →「中文姓名」'],
        [['家長簽名', '中文姓名'], 1, '「家長簽名」在前 →「中文姓名」'],
        [['家長簽名', '姓名'], 1, '「家長簽名」在前 →「姓名」'],
        [['社名', '中文姓名'], 1, '「社名」在前 →「中文姓名」'],
        [['課名', '學生 姓名'], 1, '「課名」在前 →「學生 姓名」'],
        [['班名', '中文姓名'], 1, '「班名」在前 →「中文姓名」'],
        [['校名', '學生 姓名'], 1, '「校名」在前 →「學生 姓名」'],
        [['隊名', '名字'], 1, '「隊名」在前 →「名字」'],
        [['活動名', '錄取名單'], 1, '「活動名」在前 →「錄取名單」'],
        [['錄取名單', '中文姓名'], 1, '含「姓名」（2 分）優先於只含「名」（1 分）'],
        [['中文姓名', '姓名'], 1, '整格「姓名」（3 分）優先於「中文姓名」（2 分）'],
        [['中文姓名', '學生姓名'], 1, '整格「學生姓名」（3 分）優先'],
        [['中文姓名', '英文姓名'], 0, '同分取最前面'],
        [[`名${FW}額`], -1, '只有「名　額」（全形空白）→ -1'],
        [['名 次'], -1, '只有「名 次」（半形空白）→ -1'],
        [['家長簽名'], -1, '只有「家長簽名」→ -1'],
        [['簽名'], -1, '只有「簽名」→ -1'],
        [['班名', '校名', '隊名', '社名', '課名', '活動名'], -1, '只有班名、校名、隊名、社名、課名、活動名 → -1'],
        [['家長姓名', '監護人', '聯絡人姓名'], -1, '只有家長、監護人、聯絡人欄 → -1'],
        [['姓名稱謂'], 0, '「姓名稱謂」的「名稱」不吃掉「姓名」的「名」'],
        [['學生名稱'], 0, '「學生名稱」是姓名欄'],
        [['班級名稱', '學生名稱'], 1, '「班級名稱」不是、「學生名稱」是'],
        [['Club Name', 'Student Name'], 1, '「Club Name」在前 →「Student Name」'],
        [['School Name', 'Course Name', 'Coach Name', 'Parent Name', 'Nickname', 'Username', 'Name'], 6, '英文的社團、學校、課程、教練、家長、暱稱、帳號欄都不選'],
        [['Instructor Name', 'Guardian Name', 'Class', 'name'], 3, '英文的講師、監護人欄不選'],
        [['NAME'], 0, '大寫「NAME」'],
        [['STUDENT NAME'], 0, '大寫「STUDENT NAME」'],
        [['Seat', 'Student Name'], 1, '英文「Student Name」'],
        [['姓名(中文)'], 0, '「姓名(中文)」'],
        [['報名人數', '報名學生姓名'], 1, '「報名學生姓名」仍選得到'],
        [['Class Name', 'Club Name', 'Classmate Name'], 2, '「Classmate Name」選得到，「Class Name」「Club Name」仍不選'],
        [['Class Name', 'Club Name'], -1, '只有「Class Name」「Club Name」→ -1'],
    ];
    for (const [headers, want, title] of cases) {
        R.test(`findNameColumn：${title}`, () => {
            const got = find(headers);
            return [got === want, `${JSON.stringify(headers)} → ${got}`];
        });
    }

    // 貼上路徑排除的字詞，Excel 姓名欄也都排除（兩條路徑一致）
    R.test('PASTE_NOT_STUDENT_RE 的每個字詞都在 NAME_COLUMN_EXCLUDE_RE 裡', () => {
        const words = s.run('PASTE_NOT_STUDENT_RE.source').split('|').map(w => JSON.parse(`"${w}"`));
        const re = s.run('NAME_COLUMN_EXCLUDE_RE');
        const missing = words.filter(w => !re.test(w));
        return [words.length >= 9 && missing.length === 0, `${words.length} 個字詞，缺 ${missing.length} 個`];
    });

    // 班級、座號欄的選欄不受影響
    const col = (headers, kw, isSeat) => s.run(`findColumnByKeywords(${JSON.stringify(headers)}, ${kw}, ${isSeat})`);
    R.test('findColumnByKeywords：班級、座號欄不受影響', () => {
        const got = [
            col(['姓名', '班級', '座號'], 'CLASS_KEYWORDS', false),
            col(['姓名', '班級', '座號'], 'SEAT_KEYWORDS', true),
            col(['學號', '座號'], 'SEAT_KEYWORDS', true),
            col(['所在班', '班級'], 'CLASS_KEYWORDS', false),
        ];
        return [got.join(',') === '1,2,1,1', got.join(',')];
    });
}

// ============================================================
// 2. 多工作表：姓名欄誤中課程表時，課程表仍略過，上課時間仍用來選星期
// ============================================================
const SCHEDULES = [
    ['「社團、星期、隊名」', [['社團', '星期', '隊名'], ['桌遊社', '週三', '桌遊A隊'], ['直笛社', '週五', '直笛B隊']]],
    ['「課程名稱、上課時間、Course Name」', [['課程名稱', '上課時間', 'Course Name'], ['桌遊社', '週三', 'Board Games'], ['直笛社', '週五', 'Recorder']]],
    ['第 1 列是標題、第 2 列才是表頭', [['115 學年度課後活動錄取名單總覽'], ['課程名稱', '上課時間', '名額'], ['桌遊社', '週三', '20'], ['直笛社', '週五', '15']]],
    ['「課程名稱、上課時間、名　額」', [['課程名稱', '上課時間', `名${FW}額`], ['桌遊社', '週三', '20'], ['直笛社', '週五', '15']]],
];
for (const [label, rows] of SCHEDULES) {
    const s = loadSite();
    await uploadSheets(s, [
        { name: '1161', rows },
        { name: '社團-桌遊社', rows: [H3, ['301', '5', '測試生301_05']] },
        { name: '社團-直笛社', rows: [H3, ['302', '11', '測試生302_11']] },
    ]);
    R.check(`課程表${label}：不出現在對照表、沒有提示`,
        mappingKeys(s) === '社團-桌遊社,社團-直笛社' && s.alerts.length === 0, `${mappingKeys(s)} / ${s.alerts.join(' / ')}`);
    const skip = s.el('mapping-skip-alert');
    R.check(`課程表${label}：不列入「找不到姓名欄」的提醒`, skip.hidden !== false && !skip.innerHTML.includes('1161'), skip.innerHTML.slice(0, 80));
    R.check(`課程表${label}：上課時間仍用來自動選星期`, selectedDays(s) === '週三,週五', selectedDays(s));
    const res = runResult(s);
    R.check(`課程表${label}：課表只有學生`, resultNames(res) === '測試生301_05,測試生302_11', resultNames(res));
}

// ---- 社團工作表上方有「社團｜桌遊社｜上課時間｜週三」資訊列：學生仍正確讀到
{
    const s = loadSite();
    await uploadSheets(s, [
        { name: '社團-桌遊社', rows: [['社團', '桌遊社', '上課時間', '週三'], H3, ['301', '5', '測試生301_05'], ['302', '11', '測試生302_11']] },
    ]);
    R.check('資訊列：社團工作表列入對照表、沒有提示', mappingKeys(s) === '社團-桌遊社' && s.alerts.length === 0, `${mappingKeys(s)} / ${s.alerts.join(' / ')}`);
    R.check('資訊列：星期取自資訊列的上課時間', selectedDays(s) === '週三', selectedDays(s));
    const res = runResult(s);
    R.check('資訊列：學生正確讀到', resultNames(res) === '測試生301_05,測試生302_11', resultNames(res));
}

// ---- 只有姓名、社團、上課時間欄（沒有班級、座號）的社團名單：不因像課程表而略過
{
    const s = loadSite();
    await uploadSheets(s, [
        { name: '社團-桌遊社', rows: [['社團', '姓名', '上課時間'], ['桌遊社', '測試生301_05', '週三']] },
    ]);
    const res = runResult(s);
    R.check('只有社團、姓名、上課時間欄：仍是社團名單', mappingKeys(s) === '社團-桌遊社' && resultNames(res) === '測試生301_05',
        `${mappingKeys(s)} / ${resultNames(res)} / ${s.alerts.join(' / ')}`);
}

// ---- 有班級、座號、姓名欄，也有課程名稱、上課時間欄的名單：上課時間仍讀得到
{
    const s = loadSite();
    await uploadSheets(s, [
        { name: '社團-桌遊社', rows: [['班級', '座號', '姓名', '課程名稱', '上課時間'], ['301', '5', '測試生301_05', '桌遊社', '週三']] },
    ]);
    R.check('名單內有課程名稱、上課時間欄：仍自動選星期', mappingKeys(s) === '社團-桌遊社' && selectedDays(s) === '週三', `${mappingKeys(s)} / ${selectedDays(s)}`);
}

// ---- 社團名單只有「中文姓名｜課程名稱｜上課時間」（表頭沒找到、姓名欄 2 分、同一列又像課程表）：仍讀學生
{
    const s = loadSite();
    await uploadSheets(s, [
        { name: '社團-桌遊社', rows: [['中文姓名', '課程名稱', '上課時間'], ['測試生301_05', '桌遊社', '週三'], ['測試生302_11', '桌遊社', '週三']] },
    ]);
    R.check('中文姓名、課程名稱、上課時間欄：列入對照表', mappingKeys(s) === '社團-桌遊社', `${mappingKeys(s)} / ${s.alerts.join(' / ')}`);
    const res = runResult(s);
    R.check('中文姓名、課程名稱、上課時間欄：學生進課表', resultNames(res) === '測試生301_05,測試生302_11', resultNames(res));
}

// ============================================================
// 3. loadScheduleFromExcel：資訊列下方的學生名單不當成課程表
// ============================================================
{
    const s = loadSite();
    s.ctx.__wb = makeWorkbook([
        { name: '社團-桌遊社', rows: [['社團', '桌遊社', '上課時間', '週三'], H3, ['301', '5', '測試生301_05']] },
        { name: '直笛社', rows: [['社團', '直笛社', '上課時間', '週五'], H3, ['302', '11', '測試生302_11']] },
        { name: '1161', rows: [['課程名稱', '上課時間'], ['書法社', '週四']] },
    ]);
    const m = s.run('loadScheduleFromExcel(__wb)') || {};
    const keys = Object.keys(m);
    const values = Object.values(m);
    R.check('loadScheduleFromExcel：資訊列下方的表頭與學生列不寫入',
        !keys.includes('班級') && !keys.includes('301') && !keys.includes('302') && !values.some(v => String(v).startsWith('測試生')),
        `鍵 ${keys.length} 個：${keys.filter(k => /^\d+$|班級/.test(k)).length} 個是表頭或班級`);
    R.check('loadScheduleFromExcel：社團工作表的上課時間仍讀到', m['桌遊社'] === '週三', String(m['桌遊社']));
    R.check('loadScheduleFromExcel：一般課程表仍讀到', m['書法社'] === '週四', String(m['書法社']));
}

// ---- 課程表的資料列「書法社｜週三｜各班級｜限20名」不是名單表頭（「各班級」算到班級與班、「限20名」只是 1 分）
{
    const SCHED = [['課程名稱', '上課時間', '招收對象', '名額'], ['書法社', '週三', '各班級', '限20名'], ['直笛社', '週五', '三年級', '限15名']];
    const s = loadSite();
    s.ctx.__wb = makeWorkbook([{ name: '1161', rows: SCHED }]);
    const m = s.run('loadScheduleFromExcel(__wb)') || {};
    R.check('loadScheduleFromExcel：資料列有「各班級」「限20名」的課程表仍讀到', m['書法社'] === '週三' && m['直笛社'] === '週五',
        `${String(m['書法社'])} / ${String(m['直笛社'])}`);

    const s2 = loadSite();
    await uploadSheets(s2, [
        { name: '1161', rows: SCHED },
        { name: '社團-書法社', rows: [H3, ['301', '5', '測試生301_05']] },
    ]);
    R.check('資料列有「各班級」的課程表：自動選星期生效、課程表不列入對照表',
        mappingKeys(s2) === '社團-書法社' && selectedDays(s2) === '週三', `${mappingKeys(s2)} / ${selectedDays(s2)}`);
}

// ============================================================
// 4. 單一清單總表：「家長姓名」在「姓　名」之前
// ============================================================
{
    const s = loadSite();
    await uploadSheets(s, [
        { name: '總表', rows: [['班級', '座號', '家長姓名', `姓${FW}名`, '社團'], ['301', '5', '假家長甲', '測試生301_05', '桌遊社']] },
    ]);
    const res = runResult(s);
    R.check('單一清單總表：「家長姓名」在前時讀「姓　名」欄', resultNames(res) === '測試生301_05', `${resultNames(res)} / ${s.alerts.join(' / ')}`);
}

// ============================================================
// 5. 只有課程表時的提示：說明「社團」「時間」「星期」也算課程表的欄
// ============================================================
{
    const s = loadSite();
    await uploadSheets(s, [{ name: '1161', rows: [['社團', '星期'], ['桌遊社', '週三']] }]);
    const msg = s.alerts.join('\n');
    R.check('只有課程表：提示列出「課程名稱」「社團」「上課時間」「時間」「星期」',
        ['「課程名稱」', '「社團」', '「上課時間」', '「時間」', '「星期」', '姓名欄'].every(w => msg.includes(w)), msg);
}

// ============================================================
// 6. 資訊列下方的表頭只有姓名欄（班級、座號、姓名字詞不到兩個，findHeaderRow 找不到表頭）：
//    改取前 10 列中姓名欄分數最高（至少 2 分）的那一列當表頭，學生不可消失
// ============================================================
const INFO_ONLY_NAME = [
    ['第 2 列只有「姓名」', [['社團', '桌遊社', '上課時間', '週三'], ['姓名'], ['測試生301_05']], '週三'],
    ['第 2 列是「學生姓名｜備註」', [['社團', '桌遊社', '上課時間', '週三'], ['學生姓名', '備註'], ['測試生301_05', '']], '週三'],
    ['資訊列是「社團｜桌遊社｜星期｜週三」', [['社團', '桌遊社', '星期', '週三'], ['姓名'], ['測試生301_05']], null],
];
for (const [label, rows, wantDay] of INFO_ONLY_NAME) {
    const s = loadSite();
    await uploadSheets(s, [
        { name: '社團-桌遊社', rows },
        { name: '社團-直笛社', rows: [H3, ['302', '11', '測試生302_11']] },
    ]);
    R.check(`資訊列＋${label}：列入對照表、沒有提示`,
        mappingKeys(s) === '社團-桌遊社,社團-直笛社' && s.alerts.length === 0, `${mappingKeys(s)} / ${s.alerts.join(' / ')}`);
    const skip = s.el('mapping-skip-alert');
    R.check(`資訊列＋${label}：不列入「找不到姓名欄」的提醒`, skip.hidden !== false, skip.innerHTML.slice(0, 80));
    if (wantDay) R.check(`資訊列＋${label}：星期取自資訊列`, selectedDays(s).startsWith(`${wantDay},`), selectedDays(s));
    const res = runResult(s);
    R.check(`資訊列＋${label}：學生進課表`, resultNames(res) === '測試生301_05,測試生302_11', resultNames(res));

    // 只有這一張工作表：不可當成課程表略過而跳出「找不到社團名單」
    const s1 = loadSite();
    await uploadSheets(s1, [{ name: '社團-桌遊社', rows }]);
    R.check(`資訊列＋${label}（只有這一張）：列入對照表、沒有提示`,
        mappingKeys(s1) === '社團-桌遊社' && s1.alerts.length === 0, `${mappingKeys(s1)} / ${s1.alerts.join(' / ')}`);
}

// ---- 資訊列下方「姓名｜備註｜組別」：學生列不可被當成課程表的「社團→上課時間」寫進檔案內的上課時間
{
    const s = loadSite();
    s.ctx.__wb = makeWorkbook([
        { name: '社團-桌遊社', rows: [['社團', '桌遊社', '上課時間', '週三'], ['姓名', '備註', '組別'], ['測試生301_05', '', 'A組']] },
    ]);
    const m = s.run('loadScheduleFromExcel(__wb)') || {};
    const keys = Object.keys(m);
    R.check('loadScheduleFromExcel：資訊列下方只有姓名欄的名單不寫入學生',
        !keys.some(k => k.startsWith('測試生') || k === '姓名') && m['桌遊社'] === '週三',
        `鍵 ${keys.length} 個，學生或表頭 ${keys.filter(k => k.startsWith('測試生') || k === '姓名').length} 個，桌遊社 ${String(m['桌遊社'])}`);
}

// ---- 標題「學生姓名總覽」（2 分）、課程表標頭、下方又有「中文姓名」欄：不可靜默略過。
// 「中文姓名」那一列每格都是名單欄位名稱，課程表標頭被當成名單上方的資訊列，從下一列讀到學生（舊版讀不到、列入提醒；兩者都不會讓學生消失）
{
    const s = loadSite();
    await uploadSheets(s, [
        { name: '社團-桌遊社', rows: [['學生姓名總覽'], ['課程名稱', '上課時間'], ['中文姓名'], ['測試生301_05']] },
        { name: '社團-直笛社', rows: [H3, ['302', '11', '測試生302_11']] },
    ]);
    const res = runResult(s);
    R.check('課程表標頭下方有姓名欄的名單：學生讀到，「中文姓名」不是學生、沒有提醒',
        resultNames(res) === '測試生301_05,測試生302_11' && s.el('mapping-skip-alert').hidden !== false, resultNames(res));
}

// ---- 第 1 列是標題「桌遊社錄取名單」（含「名」1 分）、第 2 列是「姓名」：表頭是第 2 列，「姓名」二字不是學生
{
    const s = loadSite();
    await uploadSheets(s, [{ name: '社團-桌遊社', rows: [['桌遊社錄取名單'], ['姓名'], ['測試生301_05'], ['測試生302_11']] }]);
    const res = runResult(s);
    R.check('標題列「桌遊社錄取名單」＋「姓名」表頭：「姓名」不是學生', resultNames(res) === '測試生301_05,測試生302_11', resultNames(res));
}

// ---- 第 1 列是標題「桌遊社」、第 2 列是「姓名｜社團｜星期」：第 2 列的姓名欄 3 分，直接當表頭讀學生
{
    const s = loadSite();
    await uploadSheets(s, [
        { name: '社團-桌遊社', rows: [['桌遊社'], ['姓名', '社團', '星期'], ['測試生301_05', '桌遊社', '週三']] },
        { name: '社團-直笛社', rows: [H3, ['302', '11', '測試生302_11']] },
    ]);
    const res = runResult(s);
    R.check('標題列「桌遊社」＋「姓名｜社團｜星期」：學生進課表',
        mappingKeys(s) === '社團-桌遊社,社團-直笛社' && resultNames(res) === '測試生301_05,測試生302_11',
        `${mappingKeys(s)} / ${resultNames(res)}`);
}

// ---- 課程表有「\u8ca0\u8cac人姓名」「\u627f\u8fa6人姓名」欄：教職員不是學生，課程表略過，上課時間仍用來選星期
// （字詞用 \u 寫法：\u8ca0\u8cac、\u627f\u8fa6）
for (const staff of ['\u8ca0\u8cac人姓名', '\u627f\u8fa6人姓名']) {
    const s = loadSite();
    await uploadSheets(s, [
        { name: '1161', rows: [['課程名稱', '上課時間', staff], ['桌遊社', '週三', '假老師甲'], ['直笛社', '週五', '假老師乙']] },
        { name: '社團-桌遊社', rows: [H3, ['301', '5', '測試生301_05']] },
        { name: '社團-直笛社', rows: [H3, ['302', '11', '測試生302_11']] },
    ]);
    const tag = staff === '\u8ca0\u8cac人姓名' ? '\u8ca0\u8cac人欄' : '\u627f\u8fa6人欄';
    R.check(`課程表有${tag}：不出現在對照表、沒有提示`,
        mappingKeys(s) === '社團-桌遊社,社團-直笛社' && s.alerts.length === 0, `${mappingKeys(s)} / ${s.alerts.join(' / ')}`);
    R.check(`課程表有${tag}：上課時間仍用來自動選星期`, selectedDays(s) === '週三,週五', selectedDays(s));
    const res = runResult(s);
    R.check(`課程表有${tag}：課表只有學生`, resultNames(res) === '測試生301_05,測試生302_11', resultNames(res));
}
{
    const s = loadSite();
    R.test('findNameColumn：只有「\u8ca0\u8cac人姓名」「\u627f\u8fa6人姓名」→ -1', () => {
        const got = s.fn('findNameColumn')(['\u8ca0\u8cac人姓名', '\u627f\u8fa6人姓名']);
        return [got === -1, String(got)];
    });
    R.test('貼上：「\u8ca0\u8cac人姓名」「\u627f\u8fa6人姓名」不是學生姓名欄', () => {
        const got = ['\u8ca0\u8cac人姓名', '\u627f\u8fa6人姓名'].map(c => s.fn('pasteHeaderCellKinds')(c).length);
        return [got.every(n => n === 0), got.join(',')];
    });
}

R.done();
