// 多工作表模式的不變量矩陣：用程式組合產生大量假工作表（社團名單與課程表），
// 每個組合各跑「搭配正常的社團工作表」與「只有這一張」兩種，逐一檢查下列不變量：
//   1. 名單的每位學生，要嘛出現在課表輸出，要嘛該工作表列入「找不到姓名欄」或「看起來是課程表」的提醒
//      （只有這一張時是對應的 alert），不可靜默消失
//   2. 課表輸出的學生姓名只能是假學生姓名，不可出現表頭文字、標題、社團名稱、老師或人員名字、數字、星期
//   3. 課程表不出現在對照表；它的上課時間用來自動選星期；只有課程表時的訊息不把課程表說成名單，
//      只有名單時的訊息也不把名單說成課程表
//   4. loadScheduleFromExcel 的鍵不是學生姓名或表頭文字，值不是學生姓名或表頭文字
//   5. 提示訊息（alert、提醒區塊）與 console 輸出不含學生姓名
// 名單與課程表的特徵重疊、設計上無法判斷的組合不放進矩陣，另在檔案最後以「刻意的例外」明確寫出預期。
// 名冊都是虛構的假工作表（測試生小明、假老師甲…），不讀取任何真實檔案。
// 環境變數 MATRIX_DUMP=1 時另外印出每個違反的組合（對照組驗證用；行首不是 PASS／FAIL，不影響統計）。
import { loadSite, makeReporter, makeWorkbook, uploadSheets } from './harness.mjs';
const R = makeReporter('多工作表名單與課程表的不變量矩陣');
const DUMP = process.env.MATRIX_DUMP === '1';

// ---------- 觀察一次上傳的結果 ----------
const pure = loadSite();
function loadSchedule(sheets) {
    pure.ctx.__wb = makeWorkbook(sheets);
    try { return pure.run('loadScheduleFromExcel(__wb)') || {}; } catch (e) { return { '例外': e.message }; }
}

async function observe(sheets) {
    const s = loadSite();
    await uploadSheets(s, sheets);
    const rows = s.el('mapping-table-body').querySelectorAll('tr');
    const days = {};
    rows.forEach(tr => { days[tr.dataset.original] = tr.querySelector('.day-select').children.find(o => o.selected)?.value; });
    let names = [];
    if (rows.length) {
        try {
            const res = s.run('runExcelProcessing(collectMapping().finalMapping)');
            names = Array.from(res.resultData, r => r.name);
        } catch (e) { names = [`例外：${e.message}`]; }
    }
    const box = s.el('mapping-skip-alert');
    return {
        keys: rows.map(tr => tr.dataset.original),
        days,
        names,
        skipHtml: box.hidden === false ? box.innerHTML : '',
        alerts: s.alerts.join('\n'),
        logs: s.logs.map(l => JSON.stringify(l)).join('\n'),
        sched: loadSchedule(sheets),
    };
}

// ---------- 違反記錄 ----------
const INVARIANTS = {
    1: '學生不可靜默消失',
    2: '課表只有假學生姓名',
    3: '課程表不在對照表、上課時間有用到、訊息不混淆名單與課程表',
    4: 'loadScheduleFromExcel 不寫入學生或表頭文字',
    5: '訊息不含學生姓名',
};
const violations = { roster: {}, schedule: {} };
const totals = { roster: 0, schedule: 0 };
function violate(cat, inv, label, detail) {
    (violations[cat][inv] ||= []).push(`${label}：${detail}`);
    if (DUMP) console.log(`VIOL ${cat} ${inv} ${label}：${detail}`);
}

// 提醒區塊與 alert 分兩段：「找不到「姓名」欄」的工作表，以及「有姓名欄、但看起來是課程表」的工作表（後者排在後面）
const NOT_FOUND_MARK = '找不到「姓名」欄';
const SCHED_LIKE_MARK = '看起來是課程表';
function notFoundSection(text) {
    const i = text.indexOf(NOT_FOUND_MARK);
    if (i === -1) return '';
    const j = text.indexOf(SCHED_LIKE_MARK, i);
    return text.slice(i, j === -1 ? undefined : j);
}
const schedLikeSection = (text) => { const i = text.indexOf(SCHED_LIKE_MARK); return i === -1 ? '' : text.slice(i); };

// 表頭文字：不可變成學生，也不可成為 loadScheduleFromExcel 的鍵或值
const HEADER_TEXTS = new Set(['班級', '座號', '姓名', '學生姓名', '中文姓名', '姓名（中文）', '學生 姓名', '備註', 'Name', 'STUDENT NAME',
    '錄取名單', '學生', '生日', '手機', '家長姓名', '課程名稱', '上課時間', '社團', '星期', '時間', 'Time']);

// ============================================================
// A. 社團名單工作表
// ============================================================
// cells：表頭；student：學生姓名那一欄（「班級｜座號｜學生」沒有姓名欄，讀不到時要列入提醒）
const ROSTER_HEADERS = [
    { cells: ['班級', '座號', '姓名'], student: 2 },
    { cells: ['姓名'], student: 0 },
    { cells: ['學生姓名'], student: 0 },
    { cells: ['中文姓名'], student: 0 },
    { cells: ['姓名（中文）'], student: 0 },
    { cells: ['學生 姓名', '備註'], student: 0 },
    { cells: ['Name'], student: 0 },
    { cells: ['STUDENT NAME'], student: 0 },
    { cells: ['班級', '座號', '錄取名單'], student: 2 },
    { cells: ['班級', '座號', '學生'], student: 2 },
    { cells: ['中文姓名', '生日', '手機'], student: 0 },
    { cells: ['家長姓名', '姓名'], student: 1 },
    // 表頭不是欄位標題（「錄取名單」只有 1 分）、沒有姓名欄（「學生」）的名單
    { cells: ['錄取名單', '備註'], student: 0 },
    { cells: ['學生', '備註'], student: 0 },
];
const INFO_ROWS = {
    '資訊列(上課時間)': ['社團', '桌遊社', '上課時間', '週三'],
    '資訊列(星期)': ['社團', '桌遊社', '星期', '週三'],
};
const NOTE_LINES = ['請於上課前十分鐘到教室集合', '上課請攜帶水壺與文具', '如需請假請事先告知', '雨天備案另行通知', '請準時接送孩子'];
const ABOVE = [['無', []]];
for (const t of ['桌遊社錄取名單', '各班級學生名單', '115學年度學生姓名總覽', '學生姓名總覽']) ABOVE.push([`標題「${t}」`, [[t]]]);
ABOVE.push(['標題＋12列說明', [['桌遊社錄取名單'], ...Array.from({ length: 12 }, (_, i) => [NOTE_LINES[i % NOTE_LINES.length]])]]);
for (const [label, info] of Object.entries(INFO_ROWS)) {
    for (const n of [0, 5, 8, 9, 12]) {
        ABOVE.push([`${label}+${n}列說明`, [info, ...Array.from({ length: n }, (_, i) => [NOTE_LINES[i % NOTE_LINES.length]])]]);
    }
}
const DATA = [
    ['一般學生', null, false],
    ['備註「English name Amy」', 'English name Amy', false],
    ['備註「請寫姓名」', '請寫姓名', false],
    ['備註「Name tag」', 'Name tag', false],
    ['備註整格「Name」', 'Name', false],
    ['班級寫「新生」', null, true],
    ['學生 15 位、第 12 位備註整格「姓名」', '姓名', 'many'],
];
// 很多位學生時的姓名（沒有數字，資料列才像欄位標題，比較容易誤判）
const MANY_NAMES = Array.from('甲乙丙丁戊己庚辛壬癸子丑寅卯辰', c => `測試生${c}`);
const CLUB_SHEET = { name: '社團-直笛社', rows: [['班級', '座號', '姓名'], ['304', '9', '測試生直笛']] };

function rosterRows(header, above, note, freshman) {
    const cells = [...header.cells];
    if (note && !cells.includes('備註')) cells.push('備註');
    const students = freshman === 'many' ? MANY_NAMES.map((name, i) => ({ cls: '301', seat: String(i + 1), name, note: i === 11 ? note : '' })) : [
        { cls: '301', seat: '5', name: '測試生小明', note: '' },
        freshman === true ? { cls: '新生', seat: '', name: '測試生新生', note: '' } : { cls: '302', seat: '11', name: '測試生小華', note: note || '' },
        { cls: '303', seat: '7', name: '測試生小美', note: '' },
    ];
    const value = (h, i, st) => {
        if (i === header.student) return st.name;
        return { '班級': st.cls, '座號': st.seat, '備註': st.note, '家長姓名': '假家長甲', '生日': '2015/01/01', '手機': '0900-000-000' }[h] ?? '';
    };
    return {
        rows: [...above, cells, ...students.map(st => cells.map((h, i) => value(h, i, st)))],
        names: students.map(st => st.name),
    };
}

const FAKE_STUDENTS = new Set(['測試生小明', '測試生小華', '測試生小美', '測試生新生', '測試生直笛', ...MANY_NAMES]);

function checkCommon(cat, label, o, expectNames) {
    const bad = o.names.filter(n => !expectNames.includes(n));
    if (bad.length) violate(cat, 2, label, `課表多出 ${bad.join('、')}`);
    const msg = `${o.alerts}\n${o.skipHtml}\n${o.logs}`;
    const leaked = [...FAKE_STUDENTS].filter(n => msg.includes(n));
    if (leaked.length) violate(cat, 5, label, `訊息含學生姓名 ${leaked.length} 個`);
    const keys = Object.keys(o.sched);
    const vals = Object.values(o.sched).map(String);
    const badKey = keys.filter(k => FAKE_STUDENTS.has(k) || HEADER_TEXTS.has(k));
    const badVal = vals.filter(v => FAKE_STUDENTS.has(v) || HEADER_TEXTS.has(v));
    if (badKey.length || badVal.length) violate(cat, 4, label, `鍵 ${badKey.join('、') || '-'}／值 ${badVal.join('、') || '-'}`);
}

for (const header of ROSTER_HEADERS) {
    for (const [aboveLabel, above] of ABOVE) {
        for (const [dataLabel, note, freshman] of DATA) {
            const { rows, names } = rosterRows(header, above, note, freshman);
            const sheet = { name: '社團-桌遊社', rows };
            for (const withClub of [true, false]) {
                const label = `名單[${header.cells.join('|')}]／上方:${aboveLabel}／資料:${dataLabel}／${withClub ? '搭配正常社團' : '只有這一張'}`;
                totals.roster++;
                const o = await observe(withClub ? [sheet, CLUB_SHEET] : [sheet]);
                const expected = withClub ? [...names, '測試生直笛'] : names;
                // 1. 不可靜默消失
                const reminded = withClub
                    ? o.skipHtml.includes(sheet.name)
                    : o.alerts.includes('找不到「姓名」欄') && o.alerts.includes(sheet.name);
                const missing = names.filter(n => !o.names.includes(n));
                if (missing.length && !reminded) violate('roster', 1, label, `少了 ${missing.length} 位，對照表 ${o.keys.join(',') || '(空)'}，訊息「${o.alerts.split('\n')[0].slice(0, 40)}」`);
                if (missing.length && missing.length < names.length) violate('roster', 1, label, `只讀到部分學生（少 ${missing.length} 位）`);
                if (withClub && !o.names.includes('測試生直笛')) violate('roster', 1, label, '正常社團的學生消失');
                // 3. 名單不可被說成課程表或被當成課程表略過
                if (!withClub && o.alerts.includes('找不到可以')) violate('roster', 3, label, '只有名單時訊息說成沒有社團名單（當成課程表或略過）');
                checkCommon('roster', label, o, expected);
            }
        }
    }
}

// ============================================================
// B. 課程時間表工作表
// ============================================================
const SCHED_BASES = [['課程名稱', '上課時間'], ['社團', '星期'], ['社團', '時間']];
const STAFF = ['教師姓名', '指導老師', '教練姓名', '助教姓名', '志工姓名', '負責人姓名', '承辦人', '名額',
    '講者姓名', '外聘人員姓名', '帶隊老師', 'Coach Name', 'Name'];
const EXTRAS = [['無', null, null]];
for (const st of STAFF) EXTRAS.push([`人員欄「${st}」`, st, st === '名額' ? ['20', '15'] : ['假老師甲', '假老師乙']]);
for (const v of ['Bring name tag', '請寫姓名', '姓名', 'Name']) EXTRAS.push([`備註「${v}」`, '備註', [v, '']]);
const DAYS = [['星期正常', ['週三', '週五']], ['桌遊社星期空白', ['', '週五']], ['桌遊社寫下午', ['下午', '週五']]];
const ROSTER_FOR_SCHED = [
    { name: '桌遊社', rows: [['班級', '座號', '姓名'], ['301', '5', '測試生小明']] },
    { name: '直笛社', rows: [['班級', '座號', '姓名'], ['302', '11', '測試生小華']] },
];

for (const base of SCHED_BASES) {
    for (const [extraLabel, extraHeader, extraVals] of EXTRAS) {
        for (const [dayLabel, days] of DAYS) {
            for (const headerAt of [1, 3]) {
                for (const sub of ['', '英文副標頭', '分頁重複標頭']) {
                    const header = extraHeader ? [...base, extraHeader] : [...base];
                    const data = ['桌遊社', '直笛社'].map((club, i) => extraHeader ? [club, days[i], extraVals[i]] : [club, days[i]]);
                    const rows = [
                        ...(headerAt === 3 ? [['115學年度第一學期社團課程表'], ['依上課星期排列']] : []),
                        header,
                        ...(sub === '英文副標頭' ? [['Name', 'Time']] : []),
                        ...(sub === '分頁重複標頭' ? [data[0], header, data[1]] : data),
                    ];
                    const sheet = { name: '1161', rows };
                    for (const withRoster of [true, false]) {
                        const label = `課程表[${header.join('|')}]／${dayLabel}／標頭在第${headerAt}列${sub ? `＋${sub}` : ''}／${withRoster ? '搭配名單' : '只有這一張'}`;
                        totals.schedule++;
                        const o = await observe(withRoster ? [sheet, ...ROSTER_FOR_SCHED] : [sheet]);
                        const expected = withRoster ? ['測試生小明', '測試生小華'] : [];
                        const missing = expected.filter(n => !o.names.includes(n));
                        if (missing.length) violate('schedule', 1, label, `名單學生少了 ${missing.length} 位`);
                        if (o.keys.includes('1161') || notFoundSection(o.skipHtml).includes('1161')) violate('schedule', 3, label, '課程表出現在對照表或「找不到姓名欄」的提醒');
                        // 有姓名欄（「Name」）的課程表可以列在「看起來是課程表」的提醒；沒有姓名欄的課程表一律靜默略過
                        if (extraHeader !== 'Name' && (schedLikeSection(o.skipHtml).includes('1161') || schedLikeSection(o.alerts).includes('1161'))) {
                            violate('schedule', 3, label, '沒有姓名欄的課程表出現在「看起來是課程表」的提醒');
                        }
                        if (withRoster) {
                            const wantA = days[0] === '週三' ? '週三' : null;
                            if (wantA && o.days['桌遊社'] !== wantA) violate('schedule', 3, label, `桌遊社的星期是 ${o.days['桌遊社']}`);
                            if (o.days['直笛社'] !== '週五') violate('schedule', 3, label, `直笛社的星期是 ${o.days['直笛社']}`);
                        } else if (!o.alerts.includes('找不到可以') || o.alerts.includes('找不到「姓名」欄')) {
                            violate('schedule', 3, label, `只有課程表時的訊息「${o.alerts.split('\n')[0].slice(0, 40)}」`);
                        }
                        // 4. 只看課程表本身讀到的上課時間：鍵只能是社團名稱
                        const own = loadSchedule([sheet]);
                        const badKeys = Object.keys(own).filter(k => k !== '桌遊社' && k !== '直笛社');
                        if (badKeys.length) violate('schedule', 4, label, `鍵 ${badKeys.join('、')}`);
                        if (days[0] === '週三' && own['桌遊社'] !== '週三') violate('schedule', 3, label, `loadScheduleFromExcel 沒讀到桌遊社的上課時間（${own['桌遊社']}）`);
                        checkCommon('schedule', label, o, expected);
                    }
                }
            }
        }
    }
}

// ============================================================
// D. 名單表頭本身就有課程欄與星期欄（「姓名｜社團｜星期」），沒有班級、座號欄
// ============================================================
// 課程欄只有一種值時是單一社團的名單，照常讀；有兩種以上的值（A 組／B 組）時和課程表分不出來，
// 不讀，但要列在「看起來是課程表」的提醒（只有這一張時是對應的 alert），不可靜默消失（不變量 1）。
// 不變量 3（訊息不把名單說成課程表）在兩種以上的值時不適用：訊息寫「看起來是」並說明是名單時該怎麼做
const D_HEADERS = [['姓名', '社團', '星期'], ['學生姓名', '社團', '上課時間'], ['中文姓名', '課程名稱', '星期']];
const D_COURSES = [['兩組', ['桌遊社A組', '桌遊社B組', '桌遊社A組']], ['同一組', ['桌遊社', '桌遊社', '桌遊社']]];
const D_ABOVE = [['無', []], ['標題「桌遊社錄取名單」', [['桌遊社錄取名單']]]];
const D_NOTES = [['一般學生', null], ['備註「Name tag」', 'Name tag'], ['備註整格「姓名」', '姓名']];
const D_STUDENTS = ['測試生小明', '測試生小華', '測試生小美'];
for (const header of D_HEADERS) {
    for (const [courseLabel, courses] of D_COURSES) {
        for (const [aboveLabel, above] of D_ABOVE) {
            for (const [noteLabel, note] of D_NOTES) {
                const cells = note ? [...header, '備註'] : [...header];
                const data = D_STUDENTS.map((n, i) => {
                    const r = [n, courses[i], '週三'];
                    return note ? [...r, i === 1 ? note : ''] : r;
                });
                const sheet = { name: '社團-桌遊社', rows: [...above, cells, ...data] };
                for (const withClub of [true, false]) {
                    const label = `名單[${cells.join('|')}]／課程欄:${courseLabel}／上方:${aboveLabel}／資料:${noteLabel}／${withClub ? '搭配正常社團' : '只有這一張'}`;
                    totals.roster++;
                    const o = await observe(withClub ? [sheet, CLUB_SHEET] : [sheet]);
                    const missing = D_STUDENTS.filter(n => !o.names.includes(n));
                    const reminded = withClub
                        ? o.skipHtml.includes(sheet.name)
                        : notFoundSection(o.alerts).includes(sheet.name) || schedLikeSection(o.alerts).includes(sheet.name);
                    if (missing.length && !reminded) violate('roster', 1, label, `少了 ${missing.length} 位，對照表 ${o.keys.join(',') || '(空)'}，訊息「${o.alerts.split('\n')[0].slice(0, 40)}」`);
                    if (missing.length && missing.length < D_STUDENTS.length) violate('roster', 1, label, `只讀到部分學生（少 ${missing.length} 位）`);
                    if (courseLabel === '同一組' && missing.length) violate('roster', 1, label, '課程欄只有一種值的名單沒有讀到');
                    if (withClub && !o.names.includes('測試生直笛')) violate('roster', 1, label, '正常社團的學生消失');
                    checkCommon('roster', label, o, withClub ? [...D_STUDENTS, '測試生直笛'] : D_STUDENTS);
                }
            }
        }
    }
}

// ---------- 彙整成檢查項目 ----------
R.check(`名單組合數 ${totals.roster}`, totals.roster >= 1000, String(totals.roster));
R.check(`課程表組合數 ${totals.schedule}`, totals.schedule >= 1000, String(totals.schedule));
for (const cat of ['roster', 'schedule']) {
    for (const [inv, title] of Object.entries(INVARIANTS)) {
        const list = violations[cat][inv] || [];
        R.check(`${cat === 'roster' ? '名單' : '課程表'}：不變量 ${inv}（${title}）`, list.length === 0,
            list.length ? `違反 ${list.length} 個，例如 ${list.slice(0, 3).join(' ／ ')}` : '');
    }
}

// ============================================================
// C. 刻意的例外：名單與課程表的特徵重疊、無法同時滿足時，以「寧可提醒、不可靜默遺失或讀錯」為準
// ============================================================
{
    // 只有一門課的課程表，人員欄的標題是整格的「Name」：和「Name｜課程名稱｜上課時間」的單一社團名單分不出來，
    // 依名單讀（人員會變成學生）。兩門課以上時，課程欄有不同的社團，才判斷是課程表
    const o = await observe([{ name: '1161', rows: [['課程名稱', '上課時間', 'Name'], ['桌遊社', '週三', '假老師甲']] }]);
    R.check('刻意的例外：只有一門課的課程表＋「Name」欄 → 依名單讀（無法和單一社團名單區分）',
        o.keys.includes('1161') && o.names.join(',') === '假老師甲', `${o.keys.join(',')} / ${o.names.join(',')}`);
}
{
    // 同一列就有班級、座號、姓名與「社團」「星期」欄、課程欄有不同社團的名單：有班級或座號欄，照名單讀（不當成課程表略過）
    const o = await observe([{ name: '各社團', rows: [['班級', '座號', '姓名', '社團', '星期'], ['301', '5', '測試生小明', '桌遊社', '週三'], ['302', '11', '測試生小華', '直笛社', '週五']] }]);
    R.check('同一列有班級、座號、姓名與社團、星期欄的名單 → 照名單讀',
        o.names.sort().join(',') === '測試生小明,測試生小華', `${o.keys.join(',')} / ${o.names.join(',')} / ${o.alerts}`);
}
{
    // 只有表頭、沒有學生：表頭像名單欄位標題（「姓名」「中文姓名」）時是 0 人的社團，不列入「找不到姓名欄」的提醒
    for (const h of [['姓名'], ['中文姓名'], ['班級', '座號', '姓名']]) {
        const o = await observe([{ name: '社團-桌遊社', rows: [h] }, CLUB_SHEET]);
        R.check(`只有表頭「${h.join('|')}」沒有學生 → 0 人的社團（不列入提醒）`,
            o.keys.includes('社團-桌遊社') && !o.skipHtml.includes('社團-桌遊社'), `${o.keys.join(',')} / ${o.skipHtml.slice(0, 40)}`);
    }
}

R.done();
