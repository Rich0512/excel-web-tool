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
    // 不去掉頭尾的 Tab：第一列的班級、座號空白時（「\t\t測試生新生_01」），去掉會讓姓名跑到班級欄。
    // 前後加上 \u0001 再切掉，cleanInvisibleChars 的 trim 就不會去掉原本的頭尾
    let clean = cleanInvisibleChars('\u0001' + text + '\u0001').slice(1, -1);
    
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

// ---------- 貼上名單的切欄與欄位標題列 ----------

// 找欄位標題列時看的行數。標題列上方可能有社團標題與好幾行說明，所以不只看前 5 行。
// 推測社團名稱與星期（findPastedHeaderLine）也用這個行數，解析與推測找到的標題列才會一樣
const PASTE_HEADER_SCAN_ROWS = 30;
// 欄位標題儲存格的最大字數（去掉空白與括號裡的說明後）。超過的是一段文字或整行標題，不是欄位名稱
const PASTE_HEADER_CELL_MAX = 12;
// 不是座號的「號」：序號、學號、編號（舊版把「序號」當成座號欄）
const PASTE_NOT_SEAT_RE = /學號|序號|編號|學生編號/;
// 欄位名稱後面括號裡的說明（「班級(例:301)」）：判斷前先去掉，裡面的數字不算
const PASTE_HEADER_NOTE_RE = /[(（\[【][^)）\]】]*[)）\]】]/g;

// 清理一格：不可見字元、兩側的引號與空白
function cleanPastedCell(c) {
    return cleanInvisibleChars(c).replace(/^["']+|["']+$/g, '').trim();
}

/**
 * 判斷一格是哪些欄位的名稱，回傳 'class'、'seat'、'name' 的陣列（不是欄位名稱時是空陣列）。
 * 只接受短的欄位名稱：「班級」「學生姓名」「座號」「班級(例:301)」才算，「桌遊社 A班 錄取名單」「三年二班」與含有「名」的姓名都不算
 * （舊版只要含有「班」與「姓／名」就當成標題，社團標題與資料列都會被誤認）。
 * 合在一格的「班級座號」「座號姓名」同時算兩種（和舊版一樣），值的拆法見 parsePastedText
 */
function pasteHeaderCellKinds(cell) {
    const scores = pasteHeaderCellScores(cell);
    return ['name', 'seat', 'class'].filter(k => scores[k] > 0);
}

// 不是學生本人的欄位（家長、監護人的姓名等）：不能當成學生的姓名、班級、座號欄。
// 其他三個字詞（不是學生本人的字樣）用 \u 寫法：這幾個字不在介面字型檔裡
const PASTE_NOT_STUDENT_RE = /家長|監護|\u7236|母|\u806f\u7d61|\u7dca\u6025/;
// 整格就是欄位名稱（2 分）優先於只是含有欄位字詞（1 分），例如「姓名」優先於「中文姓名」
const PASTE_EXACT_HEADERS = {
    name: ['姓名', '學生姓名', '名字', '學生名', '姓', '名', 'name', 'studentname'],
    seat: ['座號', '座', '號', '座次', 'seat', 'seatno'],
    class: ['班級', '班別', '班', '年班', 'class'],
};
const PASTE_HEADER_WORDS = { name: ['姓名', '名字'], seat: ['座號'], class: ['班級', '班別'] };

/**
 * 一格當成各欄位名稱的分數：{ name, seat, class }，2 是整格就是欄位名稱，1 是含有欄位字詞，0 是不是
 */
function pasteHeaderCellScores(cell) {
    const none = { name: 0, seat: 0, class: 0 };
    const v = toHalfwidthAlnum(String(cell || '')).replace(PASTE_HEADER_NOTE_RE, '').replace(/[\s:：*＊]/g, '').toLowerCase();
    if (!v || v.length > PASTE_HEADER_CELL_MAX || /[0-9]/.test(v)) return none;
    if (PASTE_NOT_SEAT_RE.test(v) || PASTE_NOT_STUDENT_RE.test(v)) return none;
    const score = (kind) => PASTE_EXACT_HEADERS[kind].includes(v) ? 2 : (PASTE_HEADER_WORDS[kind].some(w => v.includes(w)) ? 1 : 0);
    return { name: score('name'), seat: score('seat'), class: score('class') };
}

/**
 * 一列是不是欄位標題列：至少兩格是欄位名稱，有姓名欄，另外有班級或座號欄，而且沒有純數字的儲存格。
 * 是的話回傳 { colClass, colSeat, colName }（沒有的欄位是 -1），否則回傳 null
 */
function matchPasteHeaderRow(row) {
    if (row.filter(c => c).length < 2) return null;
    if (row.some(c => /^\d+$/.test(toHalfwidthAlnum(c)))) return null;
    const cols = { colClass: -1, colSeat: -1, colName: -1 };
    const key = { class: 'colClass', seat: 'colSeat', name: 'colName' };
    const best = { class: 0, seat: 0, name: 0 };
    let headerCells = 0;
    // 每種欄位取分數最高的那一格，同分取最前面的（「中文姓名」和「姓名」都有時用「姓名」）
    row.forEach((c, idx) => {
        const scores = pasteHeaderCellScores(c);
        if (scores.name || scores.seat || scores.class) headerCells++;
        Object.keys(key).forEach(kind => {
            if (scores[kind] > best[kind]) {
                best[kind] = scores[kind];
                cols[key[kind]] = idx;
            }
        });
    });
    if (headerCells < 2 || cols.colName === -1 || (cols.colClass === -1 && cols.colSeat === -1)) return null;
    return cols;
}

/**
 * 欄位標題列是第幾行（找不到是 -1），看前 PASTE_HEADER_SCAN_ROWS 行。lines 是已經去掉空行的各行文字。
 * 解析名單（parsePastedText）與推測社團名稱、星期（findPastedHeaderLine）共用，解析與推測找到的標題列才會一樣
 */
/**
 * 貼上的內容拆成行（parsePastedText 與 guessClubAndDay 共用，行號才會一樣）：
 * 儲存格內換行先變成空白，每一格各自清理，保留行首與行尾的 Tab（班級、座號空白的新生「\t\t測試生新生_01」，
 * 整行 trim 會把開頭的空儲存格去掉，姓名跑到班級欄）。只有 Tab 與空白的行是空行，不列入
 */
function splitPastedLines(text) {
    return normalizePastedText(text).split(/\r?\n/)
        .map(l => l.split('\t').map(c => cleanInvisibleChars(c)).join('\t'))
        .filter(l => l.replace(/\t/g, '').length > 0);
}

// 各行切成儲存格，切欄方式依整份內容決定（見 choosePasteSplitter）
function splitPastedRows(lines) {
    const split = choosePasteSplitter(lines);
    return lines.map(line => split(line).map(cleanPastedCell));
}

function findPasteHeaderRowIndex(rows) {
    for (let i = 0; i < Math.min(PASTE_HEADER_SCAN_ROWS, rows.length); i++) {
        if (matchPasteHeaderRow(rows[i])) return i;
    }
    return -1;
}

/**
 * 決定切欄方式，依整份內容而不是第一行（第一行常是沒有 Tab 的社團標題）：
 * - 任何一行有後面接著內容的 Tab：Tab（從網頁或 Excel 複製的表格）。沒有 Tab 的行（標題、說明）整行是一格。
 *   只有行尾的 Tab 不算（空白分隔的名單行尾多一個 Tab 時，仍用空白切欄）
 * - 沒有 Tab：能拆成多欄的行裡，一半以上用兩個以上的空白隔開時用連續空白，否則用空白（含全形空白）或逗號
 */
function choosePasteSplitter(lines) {
    if (lines.some(l => /\t[^\t]/.test(l))) return (l) => l.split('\t');
    const multiCell = (split) => lines.filter(l => split(l).filter(c => c.trim()).length >= 2).length;
    const byMultiSpace = (l) => l.split(/\s{2,}|\t/);
    const bySpace = (l) => l.split(/[,\s]+/);
    const m = multiCell(byMultiSpace);
    const s = multiCell(bySpace);
    return m > 0 && m * 2 >= s ? byMultiSpace : bySpace;
}

/**
 * 沒有欄位標題列時，名單上方的社團標題或說明列（不是學生）：沒有 1 到 3 位的純數字（班級、座號），而且
 * 只有一格（其他列有多欄時），或含有星期、標題字詞（錄取、名單、名冊…），或本身就是欄位標題列
 */
function isPasteTitleRow(row, multiColumn) {
    if (row.some(c => /^\d{1,3}$/.test(toHalfwidthAlnum(c)))) return false;
    // 看切出幾格而不是有內容的格數：班級、座號空白的新生（「\t\t測試生新生_01」）仍有 3 格，不是標題
    if (multiColumn && row.length < 2) return true;
    const filled = row.filter(c => c);
    const joined = toHalfwidthAlnum(filled.join(' '));
    return PASTE_DAY_RE.test(joined) || PASTE_TITLE_WORD_RE.test(joined) || matchPasteHeaderRow(row) !== null;
}

/**
 * 沒有欄位標題列時，拿來推測欄位的資料列：有像姓名的儲存格，另一格像班級或座號，含有數字，而且沒有標題字詞
 * （「桌遊社 A班 錄取名單」用空白切開後有「A班」「錄取名單」，不能當成資料列）
 */
function isPasteSampleDataRow(row) {
    // 兩欄（「301\t王小明」）也可以；舊版要 3 格以上，兩欄的名單一位學生都讀不到
    if (row.length < 2 || !row.some(c => /[0-9]/.test(toHalfwidthAlnum(c)))) return false;
    if (PASTE_TITLE_WORD_RE.test(row.join(' '))) return false;
    let n = -1, c = -1, s = -1;
    row.forEach((val, idx) => {
        if (!val) return;
        if (FIELD_DETECTION_RULES.name(val)) n = idx;
        if (FIELD_DETECTION_RULES.class(val)) c = idx;
        if (FIELD_DETECTION_RULES.seat(val, idx, c)) s = idx;
    });
    return n !== -1 && ((c !== -1 && c !== n) || (s !== -1 && s !== n));
}

/**
 * 沒有欄位標題列時，找出推測欄位用的那一列（找不到是 -1）。
 * 只看像學生資料的列（isPasteSampleDataRow），取最多列共有的格數，再取這種格數的第一列；
 * 格數一樣多時以最後一列的格數為準（標題在上方，名單在下方）。
 * 「桌遊社 限額 30 人」用空白切開也像資料（「限額」像姓名、「30」像班級），但格數和下面的名單不同
 */
function pickPasteSampleRow(rows) {
    const candidates = [];
    rows.forEach((r, i) => { if (isPasteSampleDataRow(r)) candidates.push(i); });
    if (candidates.length === 0) return -1;
    const counts = new Map();
    candidates.forEach(i => counts.set(rows[i].length, (counts.get(rows[i].length) || 0) + 1));
    let best = rows[candidates[candidates.length - 1]].length;
    counts.forEach((n, len) => { if (n > counts.get(best)) best = len; });
    return candidates.find(i => rows[i].length === best);
}

// 智慧解析直接貼上之文字/表格 (Tab 分隔)
function parsePastedText(text) {
    if (!text) return [];
    // 保留行首與行尾的 Tab（見 splitPastedLines）
    const lines = splitPastedLines(text);
    if (lines.length === 0) return [];

    const cleanCell = cleanPastedCell;

    // 切欄方式依整份內容決定（見 choosePasteSplitter），不只看第一行
    const rows = splitPastedRows(lines);
    const multiColumn = rows.some(r => r.filter(c => c).length >= 2);

    let colClass = -1, colSeat = -1, colName = -1;
    // 沒有欄位標題列時，從這一列開始讀（跳過上方的社團標題與說明）
    let dataStart = 0;

    // 1. 找欄位標題列（標題列上方可能有社團標題與說明，見 matchPasteHeaderRow）
    const headerIdx = findPasteHeaderRowIndex(rows);
    if (headerIdx !== -1) ({ colClass, colSeat, colName } = matchPasteHeaderRow(rows[headerIdx]));

    // 2. 若無表頭 (最糟情況下)，啟用智慧推測 (根據規則進行欄位分類)
    if (headerIdx === -1) {
        let sampleRow = null;
        // 先找看起來就是學生資料的列（見 pickPasteSampleRow）；找不到時退回第一個有 3 格以上的列（和舊版相同）
        let sampleIdx = pickPasteSampleRow(rows);
        if (sampleIdx === -1) {
            sampleIdx = rows.findIndex(r => r.length >= 3 && r.some(c => c.length > 0));
        }
        if (sampleIdx !== -1) sampleRow = rows[sampleIdx];

        if (sampleRow) {
            let guessedClass = -1;
            let guessedSeat = -1;
            let guessedName = -1;

            // 班級欄：「三年二班」「301」優先於兩位數（「05」也符合班級規則，舊版取最後一個，
            // 「301 05 王小明」會把座號讀成班級，班級無法辨識就不能抽籤）；同分取最前面的
            let classScore = 0;
            for (let idx = 0; idx < sampleRow.length; idx++) {
                const val = sampleRow[idx];
                if (!val) continue;

                if (FIELD_DETECTION_RULES.name(val)) {
                    guessedName = idx;
                }
                if (FIELD_DETECTION_RULES.class(val)) {
                    const score = /^\d{2}$/.test(cleanInvisibleChars(val)) ? 1 : 2;
                    if (score > classScore) {
                        classScore = score;
                        guessedClass = idx;
                    }
                }
            }
            // 座號欄：班級欄之後第一個符合座號規則的儲存格
            for (let idx = 0; idx < sampleRow.length; idx++) {
                const val = sampleRow[idx];
                if (!val || idx === guessedClass || idx === guessedName) continue;
                if (FIELD_DETECTION_RULES.seat(val, idx, guessedClass)) {
                    guessedSeat = idx;
                    break;
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

        // 名單上方的社團標題與說明不是學生：標題字詞、星期、只有一格，或班級、座號欄的值不像班級、座號
        // （「桌遊社 限額 30 人」的班級欄是「桌遊社」）
        const fitsColumns = (row) => {
            const cls = colClass < row.length ? row[colClass] : '';
            const seat = colSeat < row.length ? row[colSeat] : '';
            return (!cls || FIELD_DETECTION_RULES.class(cls)) && (!seat || /^\d{1,2}$/.test(toHalfwidthAlnum(seat)));
        };
        while (dataStart < sampleIdx && (isPasteTitleRow(rows[dataStart], multiColumn) || !fitsColumns(rows[dataStart]))) dataStart++;
    }

    const students = [];
    const startRow = headerIdx === -1 ? dataStart : headerIdx + 1;

    for (let i = startRow; i < rows.length; i++) {
        const row = rows[i];
        const rawName = row[colName];
        if (!rawName) continue;
        const nameVal = cleanCell(rawName);
        if (!nameVal) continue;

        // 合計列：「合計」可能在姓名欄，也可能在班級欄（「合計\t\t2」），和 Excel 模式一樣兩欄都看
        let classVal = (colClass !== -1 && colClass < row.length) ? cleanCell(row[colClass]) : "";
        if (isTotalRow(nameVal, classVal)) {
            continue;
        }
        // 欄位標題列（標題列沒被認出，或分頁複製時重複出現）不是學生：不可多出一位叫「姓名」的學生
        if (pasteHeaderCellKinds(nameVal).includes('name') || matchPasteHeaderRow(row)) {
            continue;
        }

        let seatVal = (colSeat !== -1 && colSeat < row.length) ? cleanCell(row[colSeat]) : "";
        let studentName = nameVal;
        // 欄位標題是「班級座號」（一格）：「30105」拆成班級 301、座號 05；拆不開時兩欄都用原本的值（和舊版一樣）
        if (colSeat !== -1 && colSeat === colClass) {
            const m = /^(\d{3})(\d{2})$/.exec(toHalfwidthAlnum(classVal));
            if (m) [classVal, seatVal] = [m[1], m[2]];
        }
        // 欄位標題是「座號姓名」（一格）：「05王小明」拆成座號 05、姓名王小明；拆不開時座號留空
        if (colSeat !== -1 && colSeat === colName) {
            const m = /^(\d{1,2})\s*(\D.*)$/.exec(toHalfwidthAlnum(nameVal));
            seatVal = m ? m[1] : '';
            if (m) studentName = nameVal.slice(nameVal.length - m[2].length);
        }

        students.push({
            class: classVal,
            seat: seatVal,
            name: studentName
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

/**
 * 欄位標題列是第幾行。先用解析名單時的判斷（findPasteHeaderRowIndex：同樣的切欄方式、同樣看前 30 行），
 * 解析與推測找到的標題列才會一樣；找不到時才用舊的寬鬆規則（前 10 行內有兩個以上的欄位字詞），維持原本的推測結果
 */
function findPastedHeaderLine(lines, rows) {
    const strict = findPasteHeaderRowIndex(rows);
    if (strict !== -1) return strict;
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
    // 和 parsePastedText 用同樣的行與切欄方式找標題列（splitPastedLines 保留行首的 Tab）；
    // 推測規則看的是去掉頭尾 Tab 的各行，行號和解析時相同
    const rawLines = splitPastedLines(text);
    const lines = rawLines.map(l => cleanInvisibleChars(l));
    const headerIdx = findPastedHeaderLine(lines, splitPastedRows(rawLines));
    // 有標題列時只看標題列上方的 5 行（社團標題通常就在標題列上方；更上面是說明文字）
    const aboveHeader = headerIdx === -1 ? [] : lines.slice(Math.max(0, headerIdx - 5), headerIdx);
    // 有標題列時從標題列上方第一行開始，一行一行看到上面：上方的說明文字（「請自備文具」）不會比社團標題先被拿來推測
    const candidates = (headerIdx === -1
        ? lines.slice(0, 3).filter(l => !looksLikePastedDataLine(l))
        : aboveHeader.filter(isTitleLineBeforeHeader).reverse()
    ).filter(l => !PASTE_FIELD_KEYWORDS.some(kw => l.includes(kw)));
    // 每行一個名字的清單不推測社團名稱；有標題列時只看上面那幾行
    const allowClub = !looksLikePastedNameList(headerIdx === -1 ? lines : aboveHeader);

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
