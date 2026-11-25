// 欄位推斷規則定義 (防範無表頭貼上時的智慧識別)
const FIELD_DETECTION_RULES = {
    // 1. 姓名規則：2~4個純中文，且不含有班級或年級的特定詞彙
    name: (val) => {
        const clean = cleanInvisibleChars(val);
        return /^[\u4e00-\u9fa5]{2,4}$/.test(clean) && 
               !['班', '年', '組', '級', '社', '校'].some(kw => clean.includes(kw));
    },
    // 2. 班級規則：字串包含「班」、「年」或符合3位數代碼 (如 102) 或特殊專班
    class: (val) => {
        const clean = cleanInvisibleChars(val);
        return clean.includes('班') || clean.includes('年') || /^\d{2,3}$/.test(clean) || 
               /^(幼|附幼|特教|資優|潛能|新生)/.test(clean);
    },
    // 3. 座號規則：限制在 1~60 之間的短數值字串 (避免誤判長位數的學號)
    seat: (val, idx, guessedClass) => {
        const clean = cleanInvisibleChars(val);
        const num = parseInt(clean, 10);
        const isValidNumber = !isNaN(num) && num > 0 && num <= 60;
        const isShortString = clean.length <= 2;
        // 優先考慮在班級欄位後面的數值欄位為座號
        const isAfterClass = guessedClass === -1 || idx > guessedClass;
        return isValidNumber && isShortString && isAfterClass;
    }
};

/**
 * 貼上文字預處理：清除不可見字元，並將被雙引號包覆的儲存格內換行 (\r\n / \n) 轉換為空格
 * 防止 Excel/Google Sheets 儲存格內換行 (Alt+Enter) 破壞列結構導致欄位錯位
 */
function normalizePastedText(text) {
    if (!text) return "";
    let clean = cleanInvisibleChars(text);
    
    let inQuotes = false;
    let result = '';
    for (let i = 0; i < clean.length; i++) {
        const char = clean[i];
        if (char === '"') {
            inQuotes = !inQuotes;
            result += char;
        } else if ((char === '\n' || char === '\r') && inQuotes) {
            result += ' '; // 替換儲存格內部的換行為空格
        } else {
            result += char;
        }
    }
    return result;
}

// 欄位索引匹配 (優先完全吻合，其次模糊包含；排除衝突欄位)
function findColumnByKeywords(headers, keywords, isSeat = false, isName = false) {
    // 1. 先尋找完全吻合欄位
    for (let idx = 0; idx < headers.length; idx++) {
        const val = String(headers[idx] || "").trim();
        if (!val) continue;
        if (isSeat && (val.includes('學號') || val.includes('學生編號'))) continue;
        if (isName && (val.includes('課程') || val.includes('社團') || val.includes('項目') || val.includes('學校'))) continue;
        if (keywords.includes(val)) {
            return idx;
        }
    }
    // 2. 其次尋找模糊包含欄位
    for (let idx = 0; idx < headers.length; idx++) {
        const val = String(headers[idx] || "").trim();
        if (!val) continue;
        if (isSeat && (val.includes('學號') || val.includes('學生編號'))) continue;
        if (isName && (val.includes('課程') || val.includes('社團') || val.includes('項目') || val.includes('學校'))) continue;
        if (keywords.some(kw => val.includes(kw))) {
            return idx;
        }
    }
    return -1;
}

// 班級、座號、姓名欄位的關鍵字（總表與各社團工作表共用）
const CLASS_KEYWORDS = ['班級', '班', 'Class', 'class'];
const SEAT_KEYWORDS = ['座號', '座', 'Seat', 'seat'];
const NAME_KEYWORDS = ['學生姓名', '姓名', 'Name', 'name', '名'];
// 名冊底部的合計列，不是學生
const TOTAL_ROW_KEYWORDS = ['合計', '總計', '統計', '人數', '小計'];

function isTotalRow(...values) {
    return TOTAL_ROW_KEYWORDS.some(k => values.some(v => v && String(v).includes(k)));
}

function readRowValues(ws, rowNum) {
    const values = [];
    ws.getRow(rowNum).eachCell({ includeEmpty: true }, (cell) => {
        values.push(getCellValueAsString(cell));
    });
    return values;
}

/**
 * 在工作表前幾列找表頭（同一列出現班級、座號、姓名等字詞至少兩個）；找不到時用第 1 列
 */
function findHeaderRow(ws, maxRows = 10) {
    for (let r = 1; r <= maxRows; r++) {
        const values = readRowValues(ws, r);
        const matchCount = ['班級', '班', '座號', '姓名'].filter(kw => values.some(v => v.includes(kw))).length;
        if (matchCount >= 2) return { headers: values, rowIndex: r };
    }
    return { headers: readRowValues(ws, 1), rowIndex: 1 };
}

/**
 * 讀取一張社團工作表的學生（多工作表模式）
 * 只有姓名欄是必要的，沒有班級或座號欄時留空；找不到姓名欄回傳 null
 */
function readClubSheetStudents(ws) {
    const { headers, rowIndex } = findHeaderRow(ws);
    const nameIdx = findColumnByKeywords(headers, NAME_KEYWORDS, false, true);
    if (nameIdx === -1) return null;
    const classIdx = findColumnByKeywords(headers, CLASS_KEYWORDS);
    const seatIdx = findColumnByKeywords(headers, SEAT_KEYWORDS, true, false);
    const cellText = (row, idx) => (idx === -1 ? '' : cleanInvisibleChars(getCellValueAsString(row.getCell(idx + 1))));

    const students = [];
    ws.eachRow((row, rowNum) => {
        if (rowNum <= rowIndex) return;
        const name = cellText(row, nameIdx);
        if (!name) return;
        const cls = cellText(row, classIdx);
        if (isTotalRow(name, cls)) return;
        students.push({ class: cls, seat: cellText(row, seatIdx), name });
    });
    return students;
}

// 讀取檔案內的內置對照工作表或課程總表 (支援 1152、對照表、個別社團 Sheet)
function loadScheduleFromExcel(workbook) {
    // 鍵是社團或工作表名稱（名冊字串），用沒有原型的物件，「__proto__」等名稱才不會被吞掉
    const mapping = createDict();
    let foundAny = false;

    workbook.eachSheet((ws) => {
        const sheetName = ws.name.trim();

        // 1. 檢查前 5 列看是否有「課程名稱」與「上課時間」標頭 (如 1152 工作表)
        let courseColIdx = -1;
        let dayColIdx = -1;
        let headerRow = -1;

        for (let r = 1; r <= Math.min(5, ws.rowCount); r++) {
            const row = ws.getRow(r);
            const vals = [];
            row.eachCell({ includeEmpty: true }, (c) => {
                vals.push(getCellValueAsString(c));
            });

            const cIdx = vals.findIndex(v => v.includes('課程名稱') || v === '社團');
            const dIdx = vals.findIndex(v => v.includes('上課時間') || v === '時間' || v.includes('星期'));

            if (cIdx !== -1 && dIdx !== -1) {
                courseColIdx = cIdx + 1;
                dayColIdx = dIdx + 1;
                headerRow = r;
                break;
            }
        }

        if (headerRow !== -1 && courseColIdx !== -1 && dayColIdx !== -1) {
            ws.eachRow((row, rowNum) => {
                if (rowNum <= headerRow) return;
                const club = getCellValueAsString(row.getCell(courseColIdx));
                const day = getCellValueAsString(row.getCell(dayColIdx));
                if (club && day) {
                    const cleanDay = day.replace(/\s+/g, '');
                    mapping[club] = cleanDay;
                    foundAny = true;
                }
            });
        }

        // 2. 檢查個別社團工作表 (如 社團-晨間棒球A班)
        if (sheetName.startsWith('社團-') || sheetName.startsWith('社團')) {
            const clubFromSheet = sheetName.replace(/^社團-?/, '').trim();
            for (let r = 1; r <= Math.min(5, ws.rowCount); r++) {
                const row = ws.getRow(r);
                for (let c = 1; c <= Math.min(15, row.cellCount); c++) {
                    const val = getCellValueAsString(row.getCell(c));
                    if (val === '上課時間') {
                        const dayVal = getCellValueAsString(row.getCell(c + 2)) || getCellValueAsString(row.getCell(c + 1));
                        if (dayVal) {
                            const cleanDay = dayVal.replace(/\s+/g, '');
                            mapping[clubFromSheet] = cleanDay;
                            foundAny = true;
                        }
                    }
                }
            }
        }
    });

    return foundAny ? mapping : null;
}

// 智慧解析直接貼上之文字/表格 (Tab 分隔)
function parsePastedText(text) {
    if (!text) return [];
    const normalized = normalizePastedText(text);
    const lines = normalized.split(/\r?\n/)
        .map(l => cleanInvisibleChars(l))
        .filter(l => l.length > 0);
    if (lines.length === 0) return [];
    
    // 輔助清洗儲存格兩側引號與空白
    const cleanCell = (c) => {
        const str = cleanInvisibleChars(c);
        return str.replace(/^["']+|["']+$/g, '').trim();
    };

    // 預設以 Tab 分割（網頁表格複製之標準格式）
    let rows = lines.map(line => line.split('\t').map(cleanCell));
    
    // 若無 Tab，嘗試多空格或逗號分割
    if (rows[0].length <= 1) {
        rows = lines.map(line => line.split(/\s{2,}|\t/).map(cleanCell));
    }
    if (rows[0].length <= 1) {
        rows = lines.map(line => line.split(/[,\s]+/).map(cleanCell));
    }
    
    let headerIdx = -1;
    let colClass = -1, colSeat = -1, colName = -1;
    
    // 1. 先嘗試掃描前 5 列尋找標準表頭
    for (let i = 0; i < Math.min(5, rows.length); i++) {
        const row = rows[i];
        const cIdx = row.findIndex(c => c.includes('班'));
        const sIdx = row.findIndex(c => (c.includes('座') || c.includes('號')) && !c.includes('學'));
        const nIdx = row.findIndex(c => c.includes('姓') || c.includes('名'));
        
        if (nIdx !== -1 && (cIdx !== -1 || sIdx !== -1)) {
            headerIdx = i;
            colClass = cIdx;
            colSeat = sIdx;
            colName = nIdx;
            break;
        }
    }
    
    // 2. 若無表頭 (最糟情況下)，啟用智慧推測 (根據規則進行欄位分類)
    if (headerIdx === -1) {
        let sampleRow = null;
        for (let i = 0; i < rows.length; i++) {
            if (rows[i].length >= 3 && rows[i].some(c => c.length > 0)) {
                sampleRow = rows[i];
                break;
            }
        }

        if (sampleRow) {
            let guessedClass = -1;
            let guessedSeat = -1;
            let guessedName = -1;

            for (let idx = 0; idx < sampleRow.length; idx++) {
                const val = sampleRow[idx];
                if (!val) continue;

                if (FIELD_DETECTION_RULES.name(val)) {
                    guessedName = idx;
                }
                if (FIELD_DETECTION_RULES.class(val)) {
                    guessedClass = idx;
                }
                if (FIELD_DETECTION_RULES.seat(val, idx, guessedClass)) {
                    guessedSeat = idx;
                }
            }

            if (guessedClass !== -1) colClass = guessedClass;
            if (guessedSeat !== -1) colSeat = guessedSeat;
            if (guessedName !== -1) colName = guessedName;
        }

        // 如果智慧規則推測皆未成功，才使用最後的硬性預設值 (0, 1, 2)
        if (colName === -1) {
            colClass = 0;
            colSeat = 1;
            colName = 2;
        }
    }
    
    const students = [];
    const startRow = headerIdx + 1;
    
    for (let i = startRow; i < rows.length; i++) {
        const row = rows[i];
        const rawName = row[colName];
        if (!rawName) continue;
        const nameVal = cleanCell(rawName);
        if (!nameVal) continue;
        
        if (isTotalRow(nameVal)) {
            continue;
        }
        
        const classVal = (colClass !== -1 && colClass < row.length) ? cleanCell(row[colClass]) : "";
        const seatVal = (colSeat !== -1 && colSeat < row.length) ? cleanCell(row[colSeat]) : "";
        
        students.push({
            class: classVal,
            seat: seatVal,
            name: nameVal
        });
    }
    
    return students;
}

// ---------- 貼上名單時推測社團名稱與星期 ----------

// 名單的欄位標題：同一行出現兩個以上就是標題列
const PASTE_HEADER_KEYWORDS = ['班級', '座號', '姓名', '學號', '序號', '抽籤順序', '選課階段', '備註'];
// 含有這些字詞的行是標題列或欄位說明，不拿來推測社團名稱
const PASTE_FIELD_KEYWORDS = [...PASTE_HEADER_KEYWORDS, '階段'];
// 標題行中不屬於社團名稱的字詞
const PASTE_TITLE_NOISE_RE = /延平|國小|\d*\s*學年度|第二學期|第一學期|\d*\s*學期|課後|社團|A班|B班|一時段|二時段|錄取|名冊|名單|學生|清單|班級|座號|姓名/g;
// 「週三」「星期三」，包含連寫的「桌遊社週三」。
// 後面是「次」時是頻率（「每週一次」），前面是數字時也是（「兩週一次」「2週」），都不是星期；
// 「每週三」則是星期三，所以「每」不排除。
// 不用 lookbehind：舊版 Safari 不支援，整個檔案會無法執行
const PASTE_DAY_RE = /(^|[^一二兩三四五六七八九十0-9])(?:星期|週|周)([一二三四五六日1-7])(?!次)/;
// 從文字中移除星期字樣用（連同前面的「每」：「每週三」整個去掉）。規則與 PASTE_DAY_RE 相同，修改時兩個一起改
const PASTE_DAY_WORD_RE = /(^|[^一二兩三四五六七八九十0-9])每?(?:星期|週|周)([一二三四五六日1-7])(?!次)/g;
// 整個詞等於這些欄位標題時，這一行是標題或欄位說明，不推測社團名稱（「名偵探社」只是含有「名」，照常推測）
const PASTE_HEADER_TOKENS = ['班', '座', '名', '號', '班級', '座號', '姓名', '學號', '序號', '名字'];
// 星期選單的預設值（index.html 的第一個選項）
const PASTE_DEFAULT_DAY = '週一';

function findPastedHeaderLine(lines) {
    for (let i = 0; i < Math.min(10, lines.length); i++) {
        const hits = PASTE_HEADER_KEYWORDS.filter(kw => lines[i].includes(kw)).length;
        if (hits >= 2) return i;
    }
    return -1;
}

function stripPastedWeekdays(str) {
    return str.replace(PASTE_DAY_WORD_RE, '$1');
}

function splitPastedTokens(line) {
    return line.split(/[\s,，、:：;；]+/).filter(Boolean);
}

// 遮蔽姓名用的字元（用圓圈或星號遮住中間的字）與原住民、外籍姓名中間的間隔號
const PASTE_NAME_MASK_RE = /[\u25cb\u3007\u25ef\uff2f\uff0a\*\u00b7\u2027\u30fb\uff0e\u2022]/;
const PASTE_NAME_SEPARATOR_RE = /[\u00b7\u2027\u30fb\uff0e\u2022]/;
// 寬鬆版「可能是姓名」：1 到 6 個中文字（可含遮蔽字元與間隔號）；有間隔號時放寬到 15 字。
// 只用來判斷「這一行是不是學生資料」，寧可多判：不套用姓名規則對班、年、組、級、社、校的排除
const PASTE_LOOSE_NAME_RE = /^[\u4e00-\u9fff\u25cb\u3007\u25ef\uff2f\uff0a\*\u00b7\u2027\u30fb\uff0e\u2022]{1,6}$/;
const PASTE_LOOSE_SEPARATED_NAME_RE = /^[\u4e00-\u9fff\u25cb\u3007\u25ef\uff2f\uff0a\*\u00b7\u2027\u30fb\uff0e\u2022]{2,15}$/;
// 社團名稱常見的結尾。以這些字結尾的詞當成社團，不當成姓名（「足球隊」「陶藝課」）。
// 不含「班」：「三年二班」是班級，「直笛班」這類社團在沒有標題列時就推不出來
const PASTE_CLUB_SUFFIX_RE = /(社|隊|團|營|課|坊)$/;

function isClubLikeToken(token) {
    return PASTE_CLUB_SUFFIX_RE.test(token) && !PASTE_NAME_MASK_RE.test(token);
}

function isPossibleNameToken(token) {
    if (isClubLikeToken(token)) return false;
    if (PASTE_LOOSE_NAME_RE.test(token)) return true;
    return PASTE_NAME_SEPARATOR_RE.test(token) && PASTE_LOOSE_SEPARATED_NAME_RE.test(token);
}

// 去掉純星期（「週三」）與純標題常用詞（「延平國小」「錄取名單」「115學年度」）的詞，剩下的才拿來判斷
function pastedCoreTokens(line) {
    // 回傳清掉星期與標題字詞後的詞：「桌遊社週三」→「桌遊社」、「直笛社團」→「直笛社」、「王小明錄取名單」→「王小明」
    return splitPastedTokens(line).map(t => normalizeClubWord(stripPastedWeekdays(t)).replace(PASTE_TITLE_NOISE_RE, '')).filter(Boolean);
}

/**
 * 沒有標題列時，判斷一行是不是學生資料（是的話不拿來推測社團名稱與星期）。
 * 原則：寧可推不出來，也不能把姓名填進社團名稱。下列情況都當成資料：
 * - 沒有星期字樣、而且拆得出多欄（Tab、連續空白、逗號）
 * - 含有遮蔽字元或間隔號的詞（遮蔽的姓名、原住民或外籍姓名）
 * - 空白分隔、含有班級或座號的數字
 * - 去掉純星期與標題常用詞後還有兩個以上的詞，其中有可能是姓名或像班級的詞（「三年二班 王小明」）
 * - 只剩一個詞，而且它可能是姓名或班級：「圍棋」和真實姓名分不出來；從網頁選取時也常從某一列的
 *   姓名格開始，第一行就只剩一個姓名，所以單獨一個像姓名的詞一律不推測。
 *   以社團常見結尾（社、隊、團、營、課、坊）結尾的詞才當成社團名稱。
 */
function looksLikePastedDataLine(rawLine) {
    // 上課時間（15:30）不是班級或座號，先去掉；去掉後只剩姓名的行仍會被下面的規則判成資料
    // 全形英數字先轉成一般英數字，「３０１」「週３」才會被當成數字與星期
    const line = stripClockTimeIfDay(rawLine);
    const hasWeekday = PASTE_DAY_RE.test(line);
    if (!hasWeekday && line.split(/\t|\s{2,}|[,，]/).filter(c => c.trim()).length >= 2) return true;

    const tokens = splitPastedTokens(line);
    if (tokens.some(t => PASTE_NAME_MASK_RE.test(t))) return true;
    if (tokens.length >= 2 && tokens.some(t => /^\d{1,3}$/.test(t))) return true;

    const core = pastedCoreTokens(line);
    if (core.length >= 2) return core.some(t => isPossibleNameToken(t) || FIELD_DETECTION_RULES.class(t));
    if (core.length === 0) return false; // 只有星期或標題常用詞
    const word = core[0];
    if (isClubLikeToken(word)) return false;
    return isPossibleNameToken(word) || FIELD_DETECTION_RULES.class(word);
}

/**
 * 沒有標題列時再保守一層：前 3 行有 2 行以上只有一個 6 字以內的詞，看起來是每行一個名字的清單，
 * 不推測社團名稱（例如英文名字「Amy」「Ben」不符合中文姓名的判斷）
 */
function looksLikePastedNameList(lines) {
    const single = lines.slice(0, 3).filter(l => {
        const core = pastedCoreTokens(l);
        return core.length === 1 && splitPastedTokens(l).length === 1 && core[0].length <= 6;
    });
    return single.length >= 2;
}

// 沒有標題列時，社團標題才會有的字詞（白名單條件之一）
const PASTE_TITLE_WORD_RE = /錄取|名單|名冊|社團|課程|學期|學年度/;
// 數字（含全形數字）與各種括號。沒有標題列時，候選名稱含有這些就不推測（「王小明(5)」「30105王小明」）
const PASTE_DIGIT_OR_BRACKET_RE = /[0-9\uff10-\uff19()\uff08\uff09\u3010\u3011\u300c\u300d\u300a\u300b\uff3b\uff3d\[\]{}<>]/;

// 「直笛社團」→「直笛社」：「社團」是標題常用詞會被清掉，接在社團名稱後面時保留「社」
function normalizeClubWord(str) {
    return str.replace(/([^\s,，、])社團/g, '$1社');
}

function guessClubNameFromLine(line) {
    if (splitPastedTokens(line).some(t => PASTE_HEADER_TOKENS.includes(t))) return null;
    let cleaned = normalizeClubWord(stripPastedWeekdays(cleanClubDisplayName(line))).replace(PASTE_TITLE_NOISE_RE, '');
    cleaned = cleaned.replace(/[^a-zA-Z0-9\u4e00-\u9fff]/g, '');
    if (cleaned.length < 2 || cleaned.length > 15) return null;
    if (!/[a-zA-Z\u4e00-\u9fff]/.test(cleaned)) return null; // 只剩數字（例如學年度）
    return cleaned;
}

// 全形英數字轉成一般英數字，讓全形字母也能被當成英文判斷（全形的 Amy 與 Amy 相同）
function toHalfwidthAlnum(str) {
    return str.replace(/[\uff10-\uff19\uff21-\uff3a\uff41-\uff5a]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xfee0));
}

/**
 * 名稱同時有英文字母與中文時，去掉英文後的中文部分要是社團（以社、隊、團、營、課、坊結尾），
 * 否則可能是英文名加上姓名（「Amy王小明」「王Amy」）。「Scratch程式社」「LEGO社」照常推測
 */
function isMixedScriptClubName(name) {
    if (!/[A-Za-z]/.test(name)) return true;
    const cjk = name.replace(/[^\u4e00-\u9fff]/g, '');
    return !cjk || PASTE_CLUB_SUFFIX_RE.test(cjk);
}

/**
 * 沒有標題列時改用白名單：一行要「看起來就是社團標題」才推測社團名稱，否則寧可推不出來。
 * 1. 至少符合一項：有星期字樣；有標題字詞（錄取、名單、名冊、社團、課程、學期、學年度）；
 *    清理後的名稱以社、隊、團、營、課、坊結尾，而且至少 4 個字
 * 2. 去掉星期與標題字詞後不含數字或括號，清理後的名稱不是純英文字母
 * 3. 清理後的名稱不能可能是姓名（「圍棋」「王小明」分不出來）；3 個字以下只接受以「社」結尾
 *    （以團、隊、營、課、坊結尾的 3 字姓名擋下來，缺點是「足球隊 週三」「陶藝課 週三」推不出來）
 * 4. 名稱含有英文字母時，一定要有星期或標題字詞，而且中文部分要是社團（英文名加上以團結尾的姓名推不出來）
 * 資料列判斷（looksLikePastedDataLine）與名單判斷（looksLikePastedNameList）在這之前已經先過濾一次
 */
// 標題行裡的上課時間（「15:30」「15:30-17:00」），判斷「含數字」之前先去掉，否則整行會被當成資料而推不出來
const PASTE_CLOCK_TIME_RE = /\d{1,2}[:：]\d{2}(?:\s*[-~\uff5e至到]\s*\d{1,2}[:：]\d{2})?/g;

// 全形英數字轉成一般英數字；同一行有星期時才去掉上課時間。
// 沒有星期的「王小明 15:30」去掉時間後，會和「圍棋」這種單行標題分不出來，所以保留數字讓它被擋下
function stripClockTimeIfDay(str) {
    const half = toHalfwidthAlnum(str);
    return PASTE_DAY_RE.test(half) ? half.replace(PASTE_CLOCK_TIME_RE, ' ') : half;
}

function guessClubNameWithoutHeader(line) {
    const half = stripClockTimeIfDay(line);
    const rest = normalizeClubWord(stripPastedWeekdays(half)).replace(PASTE_TITLE_NOISE_RE, '');
    if (PASTE_DIGIT_OR_BRACKET_RE.test(rest)) return null;

    const name = guessClubNameFromLine(half);
    if (!name) return null;
    const hasDay = PASTE_DAY_RE.test(half);
    const hasTitleWord = PASTE_TITLE_WORD_RE.test(line);
    const hasLatin = /[A-Za-z]/.test(name);
    const isLongClub = PASTE_CLUB_SUFFIX_RE.test(name) && name.length >= 4;
    if (!hasDay && !hasTitleWord && (hasLatin || !isLongClub)) return null;

    if (/^[A-Za-z]+$/.test(name) || /[0-9]/.test(name)) return null;
    if (!isMixedScriptClubName(name)) return null;
    if (isPossibleNameToken(name)) return null;
    if (name.length <= 3 && !/社$/.test(name)) return null;
    return name;
}

// 標題列之前的行：結尾的上課時間「(一 10,11)」不算數字
const PASTE_SCHEDULE_SUFFIX_RE = /[\s_]*[(（][一二三四五六日][^()（）]*[)）]\s*$/;
// 指導老師等欄位：後面接的是老師姓名，不是社團名稱
const PASTE_TEACHER_RE = /老師|教師|指導|導師/;

/**
 * 有標題列時，標題列之前的行只檢查結構（不套用沒有標題列時的整套規則，「圍棋」這類標題照常推測）。
 * 下列情況這一行不拿來推測：沒有星期卻拆得出多欄；含遮蔽字元或間隔號；有多個詞而且含 1 到 3 位的純數字；
 * 含老師、教師、指導、導師；去掉結尾的上課時間與標題常用詞後還有數字，又沒有標題字詞（「王小明(5)」）
 */
function isTitleLineBeforeHeader(line) {
    const half = stripClockTimeIfDay(line);
    const base = half.replace(PASTE_SCHEDULE_SUFFIX_RE, '');
    if (!PASTE_DAY_RE.test(base) && base.split(/\t|\s{2,}|[,，]/).filter(c => c.trim()).length >= 2) return false;
    if (PASTE_NAME_MASK_RE.test(base)) return false;
    const tokens = splitPastedTokens(base);
    if (tokens.length >= 2 && tokens.some(t => /^\d{1,3}$/.test(t))) return false;
    if (PASTE_TEACHER_RE.test(base)) return false;
    // 星期本身的數字（「週3」）不算
    if (/[0-9]/.test(stripPastedWeekdays(base).replace(PASTE_TITLE_NOISE_RE, '')) && !PASTE_TITLE_WORD_RE.test(base)) return false;
    return true;
}

function guessClubNameBeforeHeader(line) {
    const name = guessClubNameFromLine(stripClockTimeIfDay(line));
    if (!name || !isMixedScriptClubName(name)) return null;
    return name;
}

/**
 * 從貼上的內容推測社團名稱與星期，回傳 { club, day }（推測不出來是 null）
 * 只看標題列之前的行；找不到標題列時看前 3 行，並跳過看起來像學生資料的行
 */
function guessClubAndDay(text) {
    const result = { club: null, day: null };
    if (!text) return result;
    const lines = text.split(/\r?\n/).map(l => cleanInvisibleChars(l)).filter(l => l.length > 0);
    const headerIdx = findPastedHeaderLine(lines);
    const candidates = (headerIdx === -1
        ? lines.slice(0, 3).filter(l => !looksLikePastedDataLine(l))
        : lines.slice(0, Math.min(headerIdx, 5)).filter(isTitleLineBeforeHeader)
    ).filter(l => !PASTE_FIELD_KEYWORDS.some(kw => l.includes(kw)));
    // 每行一個名字的清單不推測社團名稱；有標題列時只看標題列之前的行
    const allowClub = !looksLikePastedNameList(headerIdx === -1 ? lines : lines.slice(0, headerIdx));

    for (const line of candidates) {
        if (!result.day) {
            // 全形數字的星期（「週３」）先轉成一般數字再判斷
            const m = PASTE_DAY_RE.exec(toHalfwidthAlnum(line));
            if (m && DAY_CHAR_MAP[m[2]]) result.day = DAY_CHAR_MAP[m[2]];
        }
        if (allowClub && !result.club) {
            result.club = headerIdx === -1 ? guessClubNameWithoutHeader(line) : guessClubNameBeforeHeader(line);
        }
        if (result.day && (result.club || !allowClub)) break;
    }
    return result;
}

/**
 * 依推測結果自動填入社團名稱與星期。
 * 只在欄位是空的（星期：使用者沒選過），或仍是上一次自動填入的值時才填，
 * 自動填入的值記在 dataset.autoFilled；新內容推測不出來時，清掉上一次自動填入的值。
 */
function guessClubAndDayFromPastedText(text) {
    const guess = guessClubAndDay(text);

    const clubEl = document.getElementById('paste-club-name');
    if (clubEl) {
        const auto = clubEl.dataset.autoFilled;
        const isAuto = auto !== undefined && clubEl.value === auto;
        if (clubEl.value.trim() === '' || isAuto) {
            if (guess.club) {
                clubEl.value = guess.club;
                clubEl.dataset.autoFilled = guess.club;
            } else if (isAuto) {
                clubEl.value = '';
                delete clubEl.dataset.autoFilled;
            }
        }
    }

    // 使用者在星期選單選過（change 事件，見 app.js）就不再自動改
    const dayEl = document.getElementById('paste-day-select');
    if (dayEl && dayEl.dataset.userChosen !== 'true') {
        if (guess.day) {
            dayEl.value = guess.day;
            dayEl.dataset.autoFilled = guess.day;
        } else if (dayEl.dataset.autoFilled !== undefined && dayEl.value === dayEl.dataset.autoFilled) {
            dayEl.value = PASTE_DEFAULT_DAY;
            delete dayEl.dataset.autoFilled;
        }
    }
}

/**
 * 加入社團或清除資料後，重設自動填入的記號，下一份名單重新推測
 */
function resetPasteAutoFill() {
    const clubEl = document.getElementById('paste-club-name');
    if (clubEl) delete clubEl.dataset.autoFilled;
    const dayEl = document.getElementById('paste-day-select');
    if (dayEl) {
        delete dayEl.dataset.autoFilled;
        delete dayEl.dataset.userChosen;
    }
}
