// 貼上名單時，學生姓名絕不可被推成社團名稱（guessClubAndDay），以及姓名裡的不可見字元。
// 這裡所有姓名都是虛構的，只為了測試特殊的姓名形狀：
//   含「年」「班」「校」「組」的名字（林年豪、王班明…）、遮蔽姓名（王○明）、5 個字與 1 個字的名字、
//   帶間隔號的名字（測試·小明）、以「團」「課」結尾的名字（林志團、歐陽志團、林志課）、英文名（Amy Lin）。
// 不可放入任何取自真實名冊的姓名。
import { loadSite, makeReporter, guessChecker } from './harness.mjs';
const R = makeReporter('貼上名單：姓名不可推成社團名稱');
const s = loadSite();
const expect = guessChecker(R, s);
const noClub = (title, text) => expect(`不推測：${title}`, text, null, undefined);

const C = String.fromCharCode;
const DATA = '301\t05\t測試生301_05\n302\t11\t測試生302_11';
const H = '班級 座號 姓名\n301 05 測試生301_05';
const DATA_AFTER = '\n301\t05\t測試生301_05\n302\t11\t測試生302_11';
const HEADER_AFTER = '\n班級\t座號\t姓名\n301\t05\t測試生301_05';

// ---- 只有姓名、沒有標題列
noClub('只有姓名的清單（每行一個名字）', '測試甲\n測試乙\n測試丙');
noClub('只貼一個名字', '測試甲');
noClub('第一行的名字也出現在資料列中', '測試甲\n301\t05\t測試甲');
noClub('姓名＋星期（後面沒有資料列）', '測試甲 週三');
noClub('「班級 姓名」兩欄資料（沒有座號）', '三年二班 測試甲\n三年二班 測試乙');
noClub('前 3 行有 2 行是單一短詞', 'Amy\nBen\n' + DATA);

// ---- 名字含嚴格姓名規則排除的字
noClub('每行一個名字，名字含「年」', '林年豪\n陳小年\n王小年');
noClub('名字含「年」，後面接資料列', '林年豪\n' + DATA);
noClub('名字含「班」「校」「組」', '王班明\n陳校華\n林組明');
noClub('「三年二班 林年豪」兩行', '三年二班 林年豪\n三年二班 陳小年');
noClub('「三年二班 林年豪」只有一行', '三年二班 林年豪');
noClub('「三年二班」Tab「林年豪」', '三年二班\t林年豪');

// ---- 遮蔽姓名
for (const [label, mask] of [['○', C(0x25cb)], ['〇', C(0x3007)], ['◯', C(0x25ef)], ['Ｏ', C(0xff2f)], ['＊', C(0xff0a)], ['*', '*']]) {
    noClub(`遮蔽姓名（${label}）每行一個`, `王${mask}明\n陳${mask}華`);
    noClub(`遮蔽姓名（${label}）後面接資料列`, `王${mask}明\n` + DATA);
}

// ---- 5 個字以上、1 個字、帶間隔號的姓名
noClub('5 個字的姓名', '歐陽小明美\n司馬小華甲\n諸葛小明乙');
noClub('5 個字的姓名，後面接資料列', '歐陽小明美\n' + DATA);
noClub('1 個字的姓名', '林\n陳');
noClub('1 個字的姓名，後面接資料列', '林\n' + DATA);
for (const [label, sep] of [['·', C(0x00b7)], ['‧', C(0x2027)], ['・', C(0x30fb)], ['．', C(0xff0e)]]) {
    noClub(`帶「${label}」的姓名`, `測試${sep}小明\n` + DATA);
    noClub(`帶「${label}」的長姓名`, `測試${sep}法拉${sep}卡里布安\n` + DATA);
}

// ---- 各種像姓名的寫法：每個都試 1 行、接資料列、加星期、連寫星期、加「錄取名單」
const NAME_SHAPES = [
    'Amy', '林志團', '王小明5', '王小明(5)', '王小明（5）',
    'Amy Lin', 'Ben Wu', 'Cindy Chen',
    '30105王小明', '301005王小明', '10.歐陽小明', '10、歐陽小明',
];
for (const n of NAME_SHAPES) {
    noClub(`「${n}」只有 1 行`, n);
    noClub(`「${n}」後面接資料列`, `${n}\n${DATA}`);
    noClub(`「${n}」加上星期`, `${n} 週三\n${DATA}`);
    noClub(`「${n}」連寫星期`, `${n}週三`);
    noClub(`「${n}」加上「錄取名單」`, `${n} 錄取名單`);
}
noClub('多行英文全名', 'Amy Lin\nBen Wu\nCindy Chen');
noClub('多行班級座號連寫姓名', '30105王小明\n30211陳小華\n30323林小美');
noClub('混合：英文名、連寫、括號', 'Amy Lin\n30105王小明\n王小明(5)');
noClub('名字以「團」結尾＋資料列', '林志團\n' + DATA);
noClub('名字＋「錄取名單」連寫', '王小明錄取名單');

// ---- 標題列之前的行是姓名、資料列或老師
noClub('標題列前兩行都是名字', '王小明\n陳小華\n' + H);
noClub('標題列前是 Tab 分隔的資料列', '201\t5\t王小明\n' + H);
noClub('標題列前是連續空白分隔的資料列', '201  5  王小明\n' + H);
noClub('標題列前是空白分隔、含座號的行', '201 5 王小明\n' + H);
noClub('標題列前「指導老師：王小明」', '指導老師：王小明\n' + H);
noClub('標題列前「教師 王小明」', '教師 王小明\n' + H);
noClub('標題列前「導師王小明」', '導師王小明\n' + H);
noClub('標題列前遮蔽姓名', `王${C(0x25cb)}明\n` + H);
noClub('標題列前帶間隔號的姓名', `測試${C(0x00b7)}小明\n` + H);
noClub('標題列前名稱含數字又沒有標題字詞', '王小明5\n' + H);
noClub('標題列前「王小明(5)」', '王小明(5)\n' + H);

// ---- 英文字母混在姓名旁
noClub('「Amy王小明 週三」', 'Amy王小明 週三');
noClub('「王Amy 錄取名單」', '王Amy 錄取名單');
noClub('「Amy林志團」', 'Amy林志團');
noClub('「Amy林志團」＋資料列', 'Amy林志團\n' + DATA);
noClub('標題列前「Amy王小明」', 'Amy王小明\n' + H);
noClub('標題列前「王Amy」', '王Amy\n' + H);
const FW_AMY = C(0xff21) + C(0xff4d) + C(0xff59); // 全形 Amy
noClub('全形字母＋4 字以團結尾「Ａｍｙ歐陽志團」', FW_AMY + '歐陽志團\n' + DATA);
noClub('全形字母＋3 字「Ａｍｙ林志團」', FW_AMY + '林志團\n' + DATA);
noClub('3 字以「課」結尾的名字「林志課 週三」', '林志課 週三\n' + DATA);
noClub('3 字以「課」結尾的名字「林志課 錄取名單」', '林志課 錄取名單');

// ---- 姓名加上時間、全形數字
noClub('姓名＋時間「測試生 15:30」', '測試生 15:30' + DATA_AFTER);
noClub('姓名＋全形數字星期「王小明 週３」', '王小明 週３' + DATA_AFTER);
noClub('姓名＋全形數字班級「王小明 ３０１」', '王小明 ３０１' + DATA_AFTER);
noClub('標題列前的資料列＋時間', '201\t5\t王小明 15:30' + HEADER_AFTER);
noClub('標題列前「王小明 15:30」', '王小明 15:30' + HEADER_AFTER);

// ---- 不可見字元
noClub('不可見字元夾在名字中間（每行一個）', `王${C(0x200e)}小明\n陳${C(0x2060)}小華\n林${C(0xad)}小美`);
const clean = s.fn('cleanInvisibleChars');
const INVIS = [0x2060, 0x200e, 0x200f, 0x00ad, 0x202a, 0x202b, 0x202c, 0x202d, 0x202e, 0x2066, 0x2067, 0x2068, 0x2069, 0x200b, 0xfeff];
for (const cp of INVIS) {
    R.test(`cleanInvisibleChars 移除 U+${cp.toString(16).toUpperCase().padStart(4, '0')}`, () => {
        const out = clean(`王${C(cp)}小明`);
        return [out === '王小明', JSON.stringify(out)];
    });
}
R.test('cleanInvisibleChars 把全形空白、不換行空白轉成一般空白', () => clean(`王${C(0x3000)}小${C(0xa0)}明`) === '王 小 明');
R.test('cleanInvisibleChars 保留一般文字與換行', () => clean('桌遊社\n301') === '桌遊社\n301');
R.test('parsePastedText 清掉姓名裡的不可見字元', () => {
    const parsed = s.fn('parsePastedText')(`班級\t座號\t姓名\n301\t05\t測試${C(0x2066)}生${C(0xad)}`);
    return [parsed.length === 1 && parsed[0].name === '測試生', `筆數 ${parsed.length}，姓名長度 ${parsed.map(p => p.name.length).join(',')}`];
});
R.done();
