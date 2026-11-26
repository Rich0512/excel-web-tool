// 貼上名單時連同表格上方的社團標題一起選取：標題行不可被當成欄位標題列，切欄方式依真正的資料行決定。
// 班級、座號、姓名都要讀對，不可多出一位叫「姓名」的學生；社團名稱與星期照常從標題行推測。
// 測試資料的姓名全部是虛構的（王小明、測試甲、測試生301_05…）。
import { loadSite, makeReporter, guessChecker } from './harness.mjs';
const R = makeReporter('貼上名單：標題行與欄位標題列');
const s = loadSite();
const parse = s.fn('parsePastedText');
const expectGuess = guessChecker(R, s);

const HEADER = '班級\t座號\t姓名';
const DATA = '301\t05\t王小明\n302\t11\t測試甲';
const WANT = [['301', '05', '王小明'], ['302', '11', '測試甲']];
const show = (list) => list.map(x => `${x.class}/${x.seat}/${x.name}`).join('、') || '（沒有學生）';

// 解析結果要和 want（[班級, 座號, 姓名] 的陣列）完全相同
function expectStudents(title, text, want = WANT) {
    R.test(title, () => {
        const got = parse(text);
        const ok = got.length === want.length && want.every(([c, seat, n], i) => got[i].class === c && got[i].seat === seat && got[i].name === n);
        return [ok, show(got)];
    });
}
// 沒有假學生「姓名」（也沒有姓名是其他欄位名稱的學生）
function expectNoFakeHeader(title, text) {
    R.test(title, () => {
        const got = parse(text);
        const fake = got.filter(x => ['姓名', '班級', '座號'].includes(x.name));
        return [got.length > 0 && fake.length === 0, show(got)];
    });
}

// ---- 1. 回報的三種情境
const T1 = `桌遊社 A班 錄取名單\n${HEADER}\n${DATA}`;
expectStudents('標題「桌遊社 A班 錄取名單」＋Tab 標題列：班級、座號、姓名都對', T1);
expectStudents('標題「直笛班 週三 錄取名單」：座號沒有遺失、沒有假學生', `直笛班 週三 錄取名單\n${HEADER}\n${DATA}`);
const SIX = ['延平國小 115學年度第二學期', '課後社團 錄取名單', '桌遊社 週三', '請各班導師轉知學生', '上課地點：活動中心', '如有疑問請洽教務處'].join('\n');
expectStudents('標題列上方有 6 行說明：仍找得到標題列', `${SIX}\n${HEADER}\n${DATA}`);
const MANY = Array.from({ length: 25 }, (_, i) => `說明第${i + 1}行請家長留意`).join('\n');
expectStudents('標題列上方有 25 行說明（超過掃描範圍）：退回推測欄位也讀對', `${MANY}\n${HEADER}\n${DATA}`);
expectNoFakeHeader('標題列上方有 25 行說明：沒有假學生「姓名」', `${MANY}\n${HEADER}\n${DATA}`);

// ---- 2. 切欄方式依資料行決定（第一行沒有 Tab 不代表整份沒有 Tab）
expectStudents('標題行沒有 Tab、名單有 Tab：用 Tab 切欄', `桌遊社 週三 錄取名單\n${HEADER}\n${DATA}`);
expectStudents('標題行含 Tab（「桌遊社\\t週三」）不是欄位標題列', `桌遊社\t週三\n${HEADER}\n${DATA}`);
expectStudents('整份用一個空白分隔（標題也有空白）', '桌遊社 A班 錄取名單\n班級 座號 姓名\n301 05 王小明\n302 11 測試甲');
expectStudents('整份用全形空白分隔', '桌遊社　週三　錄取名單\n班級　座號　姓名\n301　05　王小明\n302　11　測試甲');
expectStudents('整份用兩個以上的空白分隔', '直笛班  週三\n班級  座號  姓名\n301  05  王小明\n302  11  測試甲');
expectStudents('空白分隔、行尾多一個 Tab：仍用空白切欄', '桌遊社 錄取名單\n班級 座號 姓名\t\n301 05 王小明\t\n302 11 測試甲\t');
expectStudents('逗號分隔', '桌遊社 錄取名單\n班級,座號,姓名\n301,05,王小明\n302,11,測試甲');

// ---- 3. 標題含「班」「名」等字，不可被當成欄位標題列
expectStudents('標題「名偵探社 A班 週三 名單」（含班與名）', `名偵探社 A班 週三 名單\n${HEADER}\n${DATA}`);
expectStudents('空白分隔的標題「名偵探社 A班 名單」', '名偵探社 A班 名單\n班級 座號 姓名\n301 05 王小明\n302 11 測試甲');
expectStudents('標題「三年級 姓名遊戲社 班際名單」', `三年級 姓名遊戲社 班際名單\n${HEADER}\n${DATA}`);
// 沒有標題列的資料：班級含「班」、姓名含「名」也是學生，不是欄位標題列
expectStudents('沒有標題列、班級「三年二班」姓名「陳名揚」：第一列是學生', '三年二班\t05\t陳名揚\n三年三班\t07\t王小明', [['三年二班', '05', '陳名揚'], ['三年三班', '07', '王小明']]);

// ---- 4. 沒有標題的名單（既有行為不變）與欄位標題列的其他寫法
expectStudents('只有標題列與名單', `${HEADER}\n${DATA}`);
expectStudents('沒有標題列也沒有社團標題：301 不會被 05 取代成班級', DATA);
expectStudents('沒有標題列、上方有社團標題', `桌遊社 A班 錄取名單\n${DATA}`);
expectStudents('沒有標題列、空白分隔、上方有社團標題', '桌遊社 週三 錄取名單\n301 05 王小明\n302 11 測試甲');
expectStudents('「序號」不當成座號', `序號\t班級\t座號\t姓名\t抽籤順序\n1\t301\t05\t王小明\t2\n2\t302\t11\t測試甲\t1`);
expectStudents('沒有標題列、最前面是序號：座號取班級後面的欄位', '1\t301\t05\t王小明\n2\t302\t11\t測試甲');
expectStudents('「學生姓名」「班級」「座號」順序不同', `學生姓名\t班級\t座號\n王小明\t301\t05\n測試甲\t302\t11`);
expectStudents('只有班級與姓名兩欄', '桌遊社 錄取名單\n班級\t姓名\n301\t王小明\n302\t測試甲', [['301', '', '王小明'], ['302', '', '測試甲']]);
expectStudents('分頁複製時重複出現的標題列不是學生', `${HEADER}\n301\t05\t王小明\n${HEADER}\n302\t11\t測試甲`);
expectStudents('合計列照常略過（「合計」在姓名欄）', `桌遊社 錄取名單\n${HEADER}\n${DATA}\n\t\t合計`);
expectStudents('合計列照常略過（「合計」在班級欄）', `${HEADER}\n${DATA}\n合計\t\t2`);
expectStudents('姓名欄是測試生代號', `桌遊社 錄取名單\n${HEADER}\n301\t05\t測試生301_05`, [['301', '05', '測試生301_05']]);

// ---- 5. 班級、座號空白的新生（開頭是空格子）不可被略過；只有 Tab 的空行仍略過
const FRESH = '\t\t測試生新生_01';
expectStudents('第一列是班級、座號空白的新生（整份的開頭就是 Tab）', `${FRESH}\n${DATA}`, [['', '', '測試生新生_01'], ...WANT]);
expectStudents('標題列後的第一列是新生', `${HEADER}\n${FRESH}\n${DATA}`, [['', '', '測試生新生_01'], ...WANT]);
expectStudents('名單中間與最後是新生', `${HEADER}\n301\t05\t王小明\n${FRESH}\n\t\t測試生新生_02`, [['301', '05', '王小明'], ['', '', '測試生新生_01'], ['', '', '測試生新生_02']]);
expectStudents('社團標題＋標題列＋新生', `桌遊社 週三 錄取名單\n${HEADER}\n${FRESH}\n${DATA}`, [['', '', '測試生新生_01'], ...WANT]);
expectStudents('只有座號空白（「301\\t\\t王小明」）', `${HEADER}\n301\t\t王小明`, [['301', '', '王小明']]);
expectStudents('只有 Tab 或空白的行是空行，不是學生', `${HEADER}\n\t\t\n301\t05\t王小明\n \t \t \n302\t11\t測試甲\n\t`);
expectStudents('沒有標題列：社團標題下第一列是新生、推測欄位用後面的資料列', `桌遊社 錄取名單\n${FRESH}\n${DATA}`, [['', '', '測試生新生_01'], ...WANT]);

// ---- 6. 社團名稱與星期仍從標題行推測（只看標題列之前的行）
expectGuess('推測：「桌遊社 A班 錄取名單」', T1, '桌遊社', null);
expectGuess('推測：「直笛班 週三 錄取名單」', `直笛班 週三 錄取名單\n${HEADER}\n${DATA}`, '直笛班', '週三');
expectGuess('推測：「名偵探社 A班 週三 名單」', `名偵探社 A班 週三 名單\n${HEADER}\n${DATA}`, '名偵探社', '週三');
expectGuess('推測：全形空白的標題', '桌遊社　週三　錄取名單\n班級　座號　姓名\n301　05　王小明', '桌遊社', '週三');

// ---- 7. 審查回報的問題
// (1) 欄位名稱後面括號裡的說明含有數字
expectStudents('欄位名稱「班級(例:301)」', '班級(例:301)\t座號\t姓名\n' + DATA);
expectStudents('欄位名稱「班級（三位數）」＋社團標題', '桌遊社 錄取名單\n班級（三位數）\t座號\t學生姓名\n' + DATA);
// (2) 沒有 Tab、沒有標題列，社團標題含有 1 到 3 位數字
expectStudents('空白分隔「桌遊社 限額 30 人」：標題不是學生', '桌遊社 限額 30 人\n301 05 王小明\n302 11 測試甲');
expectStudents('空白分隔「桌遊社 第 2 期 錄取」：標題不是學生', '桌遊社 第 2 期 錄取\n301 05 王小明\n302 11 測試甲');
expectStudents('空白分隔「桌遊社 限額 30 人」下面只有一位學生', '桌遊社 限額 30 人\n301 05 王小明', [['301', '05', '王小明']]);
expectStudents('「桌遊社 限額 30 人」下面是 Tab 分隔的名單', '桌遊社 限額 30 人\n' + DATA);
expectStudents('「桌遊社 第 2 期 錄取」下面是 Tab 分隔的名單', '桌遊社 第 2 期 錄取\n' + DATA);
// (3) 沒有標題列的兩欄名單（班級、姓名）；每行一個姓名的清單維持讀不到（沒有班級時全部會變成新生，寧可請使用者重新選取）
expectStudents('沒有標題列的兩欄名單（Tab）', '301\t王小明\n302\t測試甲', [['301', '', '王小明'], ['302', '', '測試甲']]);
expectStudents('沒有標題列的兩欄名單（空白）', '桌遊社 錄取名單\n301 王小明\n302 測試甲', [['301', '', '王小明'], ['302', '', '測試甲']]);
expectStudents('每行一個姓名、沒有班級座號：讀不到（和舊版相同）', '王小明\n測試甲\n測試乙', []);
// (4) 合在一格的欄位名稱
expectStudents('欄位名稱「座號姓名」合在一格：拆成座號與姓名', '班級\t座號姓名\n301\t05王小明\n302\t11測試甲');
expectStudents('欄位名稱「班級座號」合在一格：「30105」拆成班級與座號', '班級座號\t姓名\n30105\t王小明\n30211\t測試甲');
expectStudents('欄位名稱「班級座號」：拆不開的值兩欄都用原本的值（和舊版相同）', '班級座號\t姓名\n301-05\t王小明', [['301-05', '301-05', '王小明']]);
// (5) 標題列上方有很多行說明時，解析與推測找到的是同一列標題列，社團名稱取最靠近標題列的那一行
const DESC19 = Array.from({ length: 19 }, (_, i) => `說明第${i + 1}行請家長留意`).join('\n');
const WORDS = ['請家長留意上課時間', '上課地點在活動中心', '請自備文具', '如有疑問請洽教務處', '家長簽名後交回'];
const DESC19B = Array.from({ length: 19 }, (_, i) => WORDS[i % WORDS.length]).join('\n');
expectStudents('19 行說明＋社團標題＋標題列：名單讀對', `${DESC19}\n桌遊社 週三 錄取名單\n${HEADER}\n${DATA}`);
expectGuess('推測：19 行說明＋社團標題＋標題列', `${DESC19}\n桌遊社 週三 錄取名單\n${HEADER}\n${DATA}`, '桌遊社', '週三');
expectGuess('推測：19 行沒有數字的說明＋社團標題（不取說明當社團名稱）', `${DESC19B}\n桌遊社 週三 錄取名單\n${HEADER}\n${DATA}`, '桌遊社', '週三');
expectGuess('推測：社團名稱與星期分兩行寫在說明下面', `${DESC19B}\n直笛社\n週三\n${HEADER}\n${DATA}`, '直笛社', '週三');
expectGuess('推測：標題列上方一行說明、一行社團標題（取靠近標題列的那一行）', `請自備文具\n桌遊社 週三\n${HEADER}\n${DATA}`, '桌遊社', '週三');

// ---- 8. 家長、監護人等欄位不可當成學生的姓名、班級、座號欄
const ONE = [['301', '05', '王小明']];
expectStudents('「家長姓名」在「姓名」前面：用學生的姓名', '班級\t座號\t家長姓名\t姓名\n301\t05\t王大明\t王小明', ONE);
expectStudents('「家長姓名」在「姓名」後面：用學生的姓名', '班級\t座號\t姓名\t家長姓名\n301\t05\t王小明\t王大明', ONE);
expectStudents('「監護人姓名」在「學生姓名」前面', '班級\t監護人姓名\t座號\t學生姓名\n301\t王大明\t05\t王小明', ONE);
expectStudents('「家長名字」在「名字」前面', '班級\t家長名字\t座號\t名字\n301\t王大明\t05\t王小明', ONE);
expectStudents('「父親姓名」「母親姓名」都在前面', '父親姓名\t母親姓名\t班級\t座號\t姓名\n王大明\t陳美華\t301\t05\t王小明', ONE);
expectStudents('「緊急聯絡人姓名」在前面', '班級\t座號\t緊急聯絡人姓名\t姓名\n301\t05\t王大明\t王小明', ONE);
expectStudents('只有「中文姓名」時用「中文姓名」', '班級\t座號\t中文姓名\n301\t05\t王小明', ONE);
expectStudents('「中文姓名」在「姓名」前面時用「姓名」', '班級\t座號\t中文姓名\t姓名\n301\t05\t王小明甲\t王小明', ONE);
expectStudents('「家長班級」「家長座號」不當成班級、座號欄', '家長班級\t家長座號\t班級\t座號\t姓名\n甲\t乙\t301\t05\t王小明', ONE);

// ---- 9. 推測社團名稱與星期時，找標題列用的行與切欄方式和解析名單相同
{
    const s2 = loadSite();
    const seen = [];
    const orig = s2.ctx.findPasteHeaderRowIndex;
    if (typeof orig === 'function') s2.ctx.findPasteHeaderRowIndex = (rows) => { seen.push(JSON.stringify(rows)); return orig(rows); };
    // 只有行首有 Tab（班級、座號空白的新生）：整行 trim 會讓推測時改用空白切欄
    const text = '桌遊社 週三\n班級 座號 姓名\n\t\t測試生新生_01\n\t\t測試生新生_02';
    const parsed = s2.fn('parsePastedText')(text);
    const guessed = s2.fn('guessClubAndDay')(text);
    R.check('只有行首有 Tab：解析與推測切出的儲存格相同', seen.length === 2 && seen[0] === seen[1], `呼叫 ${seen.length} 次`);
    R.check('只有行首有 Tab：讀到 2 位新生', parsed.length === 2 && parsed.every(x => !x.class && !x.seat), show(parsed));
    R.check('只有行首有 Tab：仍推得出社團名稱與星期', guessed.club === '桌遊社' && guessed.day === '週三', JSON.stringify(guessed));
}

// ---- 10. 個資：解析時不寫 console
R.check('解析時沒有任何 console 輸出', s.logs.length === 0, `${s.logs.length} 筆`);

R.done();
