// 全域狀態（uploadedWorkbook、sheetData 等）宣告在 app.js

/**
 * ExcelJS 的網址（loadExcelJS() 使用），帶上與這個檔案相同的 ?v=。
 * Service Worker 的靜態資源走 Cache-First 並以完整網址比對：不帶版本時，升級後第一次載入
 * 仍會被舊版 Service Worker 用快取裡的舊版 ExcelJS 回應。從 document.currentScript 取出
 * utils.js 自己的 ?v=（index.html 以 js/utils.js?v=X.Y 載入），bump-version 就不必改這裡。
 * 必須在最上層執行：currentScript 只在這個檔案執行期間有值。
 * 取不到版本（例如直接開啟檔案或測試環境）時退回不帶參數的網址。
 */
const EXCELJS_SRC = (() => {
    const base = 'js/libs/exceljs.min.js';
    try {
        const src = document.currentScript && document.currentScript.src;
        const version = src ? new URL(src, document.baseURI).searchParams.get('v') : null;
        return version ? `${base}?v=${encodeURIComponent(version)}` : base;
    } catch (e) {
        return base;
    }
})();

const WEEKDAY_REGEX = /(星期|週|周)([一二三四五六日1-7])/;

// 星期字元 → 標準名稱（一 / 1 → 週一）
const DAY_CHAR_MAP = {
    '一': '週一', '1': '週一',
    '二': '週二', '2': '週二',
    '三': '週三', '3': '週三',
    '四': '週四', '4': '週四',
    '五': '週五', '5': '週五',
    '六': '週六', '6': '週六',
    '日': '週日', '7': '週日'
};

// 晨間社團的專屬欄位（依序排在週五之後）
const MORNING_OPTIONS = ['晨間(週一、週四)', '晨間(週二、週五)', '晨間(週一、週二、週四、週五)'];

const weekdaysList = [
    "請選擇",
    "週一",
    "週二",
    "週三",
    "週四",
    "週五",
    "晨間(週一、週四)",
    "晨間(週二、週五)",
    "晨間(週一、週二、週四、週五)",
    "週六",
    "週日",
    "週一、四",
    "週二、五",
    "週一、二、四、五"
];

// 提醒框（.warning-callout）左側的警告線條圖示，DESIGN.md 元件表「提醒」
const WARNING_ICON_SVG = '<svg class="icon" viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 3l10 18H2z"/><path d="M12 10v5M12 18h.01"/></svg>';

/**
 * XSS HTML 特殊字元安全跳脫防護
 * 只有 null / undefined 視為空字串；數字 0、false 照常轉成文字
 */
function escapeHtml(str) {
    if (str === null || str === undefined) return "";
    return String(str)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}

/**
 * 解析星期字串，支援單日或多日（如 "週一"、"一、四"、"週一、四"、"一、二、四、五"）或專屬晨間時段
 * 回傳標準星期陣列，例如 ['週一', '週四'] 或 ['晨間(週一、週四)']
 */
function parseDays(dayStr) {
    if (!dayStr) return [];
    const str = String(dayStr).trim();
    if (str.startsWith('晨間')) return [str];
    const found = [];
    const matches = str.match(/(?:星期|週|周)?([一二三四五六日1-7])/g);
    if (matches) {
        matches.forEach(m => {
            const d = DAY_CHAR_MAP[m.replace(/星期|週|周/, '')];
            if (d && !found.includes(d)) {
                found.push(d);
            }
        });
    }
    return found;
}

/**
 * 把各種星期寫法轉成下拉選單的選項，例如「一、二、四、五」→「週一、二、四、五」、「星期三」→「週三」
 * 對不上任何選項時回傳 null
 */
function normalizeDayOption(dayStr) {
    if (!dayStr) return null;
    const str = String(dayStr).replace(/\s+/g, '');
    if (weekdaysList.includes(str) && str !== '請選擇') return str;
    const days = parseDays(str);
    if (days.length === 0 || days[0].startsWith('晨間')) return null;
    const option = days.length === 1 ? days[0] : '週' + days.map(d => d.replace('週', '')).join('、');
    return weekdaysList.includes(option) ? option : null;
}

/**
 * 從上課時間判斷晨間社團屬於哪個專屬欄位；上課時間沒有星期資訊時回傳 null
 */
function morningDayFromSchedule(sched) {
    const s = String(sched || '');
    if (['一', '二', '四', '五'].every(c => s.includes(c))) return MORNING_OPTIONS[2];
    if (s.includes('二') && s.includes('五')) return MORNING_OPTIONS[1];
    if (s.includes('一') && s.includes('四')) return MORNING_OPTIONS[0];
    return null;
}

/**
 * 檔案中沒有上課時間時，依社團名稱推測晨間社團的欄位（預設週一、週四）
 */
function morningDayFromName(name) {
    const s = String(name || '');
    if (s.includes('游泳')) return MORNING_OPTIONS[2];
    if (/棒球\s*B/i.test(s)) return MORNING_OPTIONS[1];
    return MORNING_OPTIONS[0];
}

/**
 * 從欄位或工作表名稱辨識 A/B 時段，找不到回傳 ""
 * 單獨的 A/B 字母才算時段：「晨間棒球A班」的 A 是班別，「MV」等英文字中的字母也不算
 */
const SLOT_WORD_RE = /(第一時段|第二時段|一時段|二時段|第一節|第二節|Slot\s*[AB])/i;
const SLOT_LETTER_RE = /(^|[^A-Za-z])([AB])(?![A-Za-z班組])/i;

function detectSlot(text) {
    const str = String(text || '');
    const word = str.match(SLOT_WORD_RE);
    if (word) return /^(第一|一)|A$/i.test(word[1]) ? 'A' : 'B';
    const letter = str.match(SLOT_LETTER_RE);
    return letter ? letter[2].toUpperCase() : '';
}

/**
 * 拆出社團名稱「結尾」的時段後綴（貼上模式用），回傳 { name, slot }；認不出時 name 是原名、slot 是 ""
 * 時段字詞與 detectSlot() 相同（「第一時段」「一時段」「第一節」「Slot A」…），或單獨的 A/B 字母：
 * 沒有括號時，字母前面是英文字母就不算（「Yoga」「Club」的結尾不是時段）；結尾是「A班」時也不算。
 * 可以用括號包住（含全形括號）：「桌遊(A)」「桌遊（B）」。
 * 清掉的就是辨識出來的那一段，連同前面的空白與括號；清完沒有名稱時（例如「A」）維持原名、不算時段。
 * 單獨的「時段」二字沒有 A/B 資訊，不清掉。
 */
const SLOT_SUFFIX_RE = new RegExp(`\\s*([(\\uFF08]\\s*)?(${SLOT_WORD_RE.source}|[AB])(\\s*[)\\uFF09])?$`, 'i');

function splitSlotSuffix(text) {
    const str = String(text || '').trim();
    const none = { name: str, slot: '' };
    // 全形英數字先轉成一般英數字再比對（全形的 A 也算時段）；轉換不改變長度，名稱仍從原本的字串取出
    const half = toHalfwidthAlnum(str);
    const m = half.match(SLOT_SUFFIX_RE);
    // m[1]、m[4] 是左右括號，只有其中一個時不算
    if (!m || !m[1] !== !m[4]) return none;
    const token = m[2];
    if (!m[1] && token.length === 1 && /[A-Za-z]/.test(half.charAt(half.length - 2))) return none;
    // 名稱和時段之間的「-」「_」也一起清掉（「桌遊-A」→「桌遊」）
    const name = str.slice(0, m.index).replace(/[\s_-]+$/, '').trim();
    return name ? { name, slot: detectSlot(token) } : none;
}

/**
 * 寒暑假模式下，總表的時段欄位名稱要加在社團名稱後面的標籤
 * 能辨識出 A/B 時就只用字母（「第一時段(社團A)」→「A」），否則清掉通用字詞、括號與頭尾的分隔符號後使用剩下的字
 * （「上午社團」→「上午」、「課程名稱(上午)」→「上午」）；只剩通用字詞時回傳 ""（「課程名稱」「社團名單」不加標籤）
 * 欄位之間的標籤可能相同，處理總表要用 slotLabelsFromColumns()
 */
const SLOT_LABEL_GENERIC_RE = /時段|slot|社團|課程|項目|節|名稱|名單|內容|選項|類別|錄取|報名|欄位/gi;

function slotLabelFromColumn(colName) {
    const slot = detectSlot(colName);
    if (slot) return slot;
    return String(colName || '')
        .replace(SLOT_LABEL_GENERIC_RE, '')
        .replace(/[()（）]/g, '')
        .replace(/^[\s\-_:/、，：－]+|[\s\-_:/、，：－]+$/g, '')
        .trim();
}

/**
 * 寒暑假模式下，總表每個時段欄位的標籤，回傳 Map（欄位名稱 → 標籤，"" 表示不加標籤）
 * 辨識不出 A/B 的欄位，清完通用字詞後如果和另一個欄位相同（「社團」「課程名稱」都清成 ""、
 * 「上午社團」「上午課程」都是「上午」），這幾欄改用完整的欄位名稱當標籤：
 * 否則同一個社團在這兩欄排到同一天時，課表文字相同會被合併成一筆，看不出學生錄取了兩個時段，也不會提醒衝堂。
 * 欄位名稱是總表的表頭，不會重複。A/B 欄位維持字母（兩欄都是 A 就是同一個時段）。
 */
function slotLabelsFromColumns(colNames) {
    const base = colNames.map(col => [col, slotLabelFromColumn(col), !!detectSlot(col)]);
    const count = new Map();
    base.forEach(([, label]) => count.set(label, (count.get(label) || 0) + 1));
    return new Map(base.map(([col, label, isAB]) => [
        col,
        !isAB && count.get(label) > 1 ? (String(col).trim() || String(col)) : label,
    ]));
}

/**
 * 單一清單模式的對照鍵：同一個社團可能同時出現在 A、B 兩個時段欄位，星期可能不同，所以要連同欄位一起區分
 */
function mappingKey(slotCol, club) {
    return `${slotCol}\u0001${club}`;
}

// 全形英數字轉成一般英數字，讓全形字母也能被當成英文判斷（全形的 Amy 與 Amy 相同）；parser.js、parseGrade、splitSlotSuffix 共用
function toHalfwidthAlnum(str) {
    return str.replace(/[\uff10-\uff19\uff21-\uff3a\uff41-\uff5a]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xfee0));
}

const GRADE_CHAR_MAP = {
    '一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6,
    '1': 1, '2': 2, '3': 3, '4': 4, '5': 5, '6': 6
};

/**
 * 解析年級：空白或「新生」回傳 0；能辨識的回傳 1–6；不是空白卻讀不出年級時回傳 null
 * 只看第一個有意義的字元（跳過空白與標點，全形數字當成一般數字）：「301」「3年1班」「三年一班」與全形的 301 都是 3 年級；
 * 「701」「01」「資源班」回傳 null，不會在字串中間找到「1」就當成一年級
 */
function parseGrade(classStr) {
    const str = toHalfwidthAlnum(cleanInvisibleChars(classStr));
    if (!str || str.startsWith("新生")) return 0;
    const first = str.match(/[\p{L}\p{N}]/u);
    if (!first) return null;
    return GRADE_CHAR_MAP[first[0]] || null;
}

/**
 * 是否為新生：班級空白（含只有不可見字元）或以「新生」開頭。
 * 抽籤（parseGrade 的 0 年級）與課表的「包含新生」篩選都用這個判斷，兩者才會一致
 */
function isFreshmanClass(classStr) {
    return parseGrade(classStr) === 0;
}

/**
 * 瀏覽器是否提供密碼學亂數（crypto.getRandomValues）。沒有時不可抽籤，也不可退回 Math.random
 */
function hasSecureRandom() {
    return typeof crypto !== 'undefined' && !!crypto && typeof crypto.getRandomValues === 'function';
}

const UINT32_RANGE = 0x100000000;

/**
 * 回傳 0 到 n - 1 之間的隨機整數，每個結果的機率完全相同（n 是 1 到 2^32 的整數）
 * 亂數來自 crypto.getRandomValues；超出可以被 n 整除的範圍的值直接丟掉重抽（rejection sampling），否則數字小的結果會多出一點機率
 */
function randomInt(n) {
    if (!Number.isInteger(n) || n < 1 || n > UINT32_RANGE) {
        throw new RangeError('randomInt 的範圍要是 1 到 2^32 之間的整數');
    }
    const limit = UINT32_RANGE - (UINT32_RANGE % n);
    const buf = new Uint32Array(1);
    for (;;) {
        crypto.getRandomValues(buf);
        if (buf[0] < limit) return buf[0] % n;
    }
}

// 抽籤優先順序的顯示文字，與 index.html #paste-priority 的選項文字相同
const PRIORITY_LABELS = { desc: '高年級優先', asc: '低年級優先' };

// 年級選單的顯示文字，與 index.html #paste-grade-start／#paste-grade-end 的選項文字相同
function gradeOptionLabel(value) {
    return value === '新生' ? '新生' : `${value} 年級`;
}

/**
 * 抽籤時間的顯示格式（本地時間）：withDate 為 true 時是 YYYY/MM/DD HH:mm:ss，否則是 HH:mm
 * 不是有效的 Date 時回傳 ""
 */
function formatDrawTime(date, withDate = false) {
    if (!date || typeof date.getTime !== 'function' || isNaN(date.getTime())) return "";
    const pad = (n) => String(n).padStart(2, '0');
    const hm = `${pad(date.getHours())}:${pad(date.getMinutes())}`;
    if (!withDate) return hm;
    return `${date.getFullYear()}/${pad(date.getMonth() + 1)}/${pad(date.getDate())} ${hm}:${pad(date.getSeconds())}`;
}

/**
 * 智慧解析 Excel 儲存格內容，支援純文字、數值、公式結果及富文本
 */
function getCellValueAsString(cell) {
    if (!cell || cell.value === null || cell.value === undefined) {
        return "";
    }
    
    const val = cell.value;
    
    // 1. 處理公式儲存格 (Formula)
    if (typeof val === 'object' && val !== null && 'formula' in val) {
        if (val.result !== null && val.result !== undefined) {
            return String(val.result).trim();
        }
        return "";
    }
    
    // 2. 處理富文本儲存格 (Rich Text)
    if (typeof val === 'object' && val !== null && 'richText' in val) {
        return val.richText.map(rt => rt.text || '').join('').trim();
    }
    
    // 3. 處理包含 text 屬性的物件 (例如超連結)
    if (typeof val === 'object' && val !== null && 'text' in val) {
        return String(val.text).trim();
    }
    
    // 4. 一般字串與數值直接轉字串
    return String(val).trim();
}

/**
 * 轉換欄位索引為 Excel 字母標記 (如 1 -> A, 27 -> AA)
 */
function getColLetter(colIdx) {
    let temp = colIdx;
    let letter = "";
    while (temp > 0) {
        let modulo = (temp - 1) % 26;
        letter = String.fromCharCode(65 + modulo) + letter;
        temp = Math.floor((temp - modulo) / 26);
    }
    return letter;
}

/**
 * 提取數值作為排序鍵 (比照 Python get_numeric_sort_key)
 * 空白與新生（isFreshmanClass，例如「新生」「新生班」）回傳 Infinity 排到最後，與抽籤、課表篩選的新生判斷一致。
 * 座號也用這個函式：座號不會以「新生」開頭，空白座號同樣排最後
 */
function getNumericSortKey(val) {
    if (isFreshmanClass(val)) return Infinity;
    const strVal = String(val).trim();

    // 1. 如果是純阿拉伯數字，如 "404"、"102" 或座號 "5"，直接轉數值返回
    const pureDigits = strVal.replace(/[^\d]/g, '');
    if (pureDigits && pureDigits === strVal) {
        return parseInt(pureDigits, 10);
    }

    // 中文數字對照表，支援至二十幾班
    const chMap = {
        '一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6,
        '七': 7, '八': 8, '九': 9, '十': 10,
        '十一': 11, '十二': 12, '十三': 13, '十四': 14, '十五': 15,
        '十六': 16, '十七': 17, '十八': 18, '十九': 19, '二十': 20
    };

    // 輔助解析單個中文或阿拉伯數字
    function parseNum(part) {
        if (!part) return 0;
        part = part.trim();
        if (/^\d+$/.test(part)) {
            return parseInt(part, 10);
        }
        if (chMap[part] !== undefined) {
            return chMap[part];
        }
        // 處理如 "十五"、"二十三" 的複合中文數值
        let num = 0;
        if (part.includes('十')) {
            const parts = part.split('十');
            const tenPrev = parts[0];
            const tenVal = tenPrev ? (chMap[tenPrev] || 1) : 1;
            num += tenVal * 10;
            const tenNext = parts[1];
            if (tenNext) {
                num += chMap[tenNext] || 0;
            }
        } else {
            num = chMap[part[0]] || 0;
        }
        return num || parseInt(part, 10) || 0;
    }

    // 2. 嘗試匹配 "X年Y班" 格式 (如 "四年四班" -> 404, "4年2班" -> 402)
    const matchYearClass = strVal.match(/([^年]+)年([^班]+)班?/);
    if (matchYearClass) {
        const yearPart = matchYearClass[1];
        const classPart = matchYearClass[2];
        const yearNum = parseNum(yearPart);
        const classNum = parseNum(classPart);
        if (yearNum && classNum) {
            return yearNum * 100 + classNum; // 四年四班 -> 404
        }
    }

    // 3. 嘗試匹配 "X班" 格式
    const matchClassOnly = strVal.match(/([^班]+)班/);
    if (matchClassOnly) {
        const classPart = matchClassOnly[1];
        return parseNum(classPart);
    }

    // 4. 退回基本正則過濾，抽取所有數字
    const cleanVal = strVal.replace(/[^\d.]/g, '');
    if (cleanVal) {
        const num = parseFloat(cleanVal);
        return isNaN(num) ? Infinity : num;
    }

    return Infinity;
}

/**
 * 正規化星期名稱為 "週一" ~ "週日"，支援括號如 (一 10,11)
 */
function normalizeWeekdayName(dayStr) {
    if (!dayStr) return null;
    const str = String(dayStr).trim();
    const match = str.match(WEEKDAY_REGEX);
    if (match) return DAY_CHAR_MAP[match[2]] || null;
    const paren = str.match(/[(（]([一二三四五六日1-7])/);
    return paren ? (DAY_CHAR_MAP[paren[1]] || null) : null;
}

/**
 * 清理社團名稱，移除後綴的時間標註以及前綴的星期標註，只保留純社團名稱
 * 例如：「周一兒童繪畫社 (一 10,11)」->「兒童繪畫社」、「周四足球社」->「足球社」
 */
function cleanClubDisplayName(name) {
    if (!name) return "";
    let cleaned = String(name).replace(/[\s_]*[(（][^()（）]+[)）]$/, '').trim();
    cleaned = cleaned.replace(/^(星期|週|周)[一二三四五六日1-7][\s_-]*/, '').trim();
    return cleaned || String(name).trim();
}

/**
 * 從社團名稱或工作表字串自動提取星期
 */
function extractWeekdayFromName(clubName) {
    return normalizeWeekdayName(clubName);
}

/**
 * 智慧解析工作表名稱與時段資訊
 */
function parseSheetName(name) {
    const raw = String(name || '').trim();
    // 學校系統匯出的工作表名稱是「社團-周一桌遊社」，「社團-」是前綴，不是社團名稱的一部分
    const base = raw.replace(/^社團\s*[-－_]?\s*/, '');
    const day = normalizeWeekdayName(base);
    const slot = detectSlot(base);

    let cleanName = base.replace(/(星期|週|周)[一二三四五六日1-7]/g, '');
    cleanName = cleanName.replace(SLOT_WORD_RE, '');
    if (slot) cleanName = cleanName.replace(SLOT_LETTER_RE, '$1');
    cleanName = cleanName.replace(/時段|課程|項目/g, '');
    // 清掉刪字後殘留在頭尾的分隔符號與空括號
    cleanName = cleanName.replace(/[(（]\s*[)）]/g, '').replace(/^[\s\-－_]+|[\s\-－_]+$/g, '');

    return {
        day: day || "請選擇",
        slot: slot,
        baseName: base || raw,
        clubName: cleanName || base || raw
    };
}

/**
 * 以名冊字串（工作表、社團、欄位名稱）當鍵的物件一律用沒有原型的物件：
 * 一般物件遇到「__proto__」會改到原型而不是新增鍵（社團默默消失），
 * 查詢「constructor」「toString」等名稱也會拿到 Object.prototype 的成員
 */
function createDict() {
    return Object.create(null);
}

// LocalStorage 管理對照設定
// (無痕模式、封鎖網站資料或內容損毀時存取會失敗，一律視為沒有記憶設定)
// 格式（隱私規則：只存欄位標題、社團或工作表名稱、星期）：
//   單一清單模式：時段欄位名稱 → { 社團名稱 → 星期 }
//   多工作表模式：工作表名稱 → 星期（舊版的「社團名稱 → 星期」也是這個形狀）
const CONFIG_STORAGE_KEY = 'yenping_club_config';
// 每學期的社團都會累積，設上限避免無限成長；以物件鍵的插入順序當作新舊順序，
// 更新時把鍵移到最後，超過上限從最前面（最久沒更新）刪除
const CONFIG_MAX_TOP_KEYS = 100;
const CONFIG_MAX_CLUBS_PER_SLOT = 200;
// 鍵（欄位標題、社團或工作表名稱）的長度上限。正常的名稱遠短於此；
// 太長的通常是把整段說明或資料誤當成標題，不值得記，也避免單筆設定佔用大量空間
const CONFIG_MAX_KEY_LENGTH = 60;

function isStorableConfigKey(key) {
    return typeof key === 'string' && key.length > 0 && key.length <= CONFIG_MAX_KEY_LENGTH;
}

/**
 * 把讀到的設定轉成沒有原型的物件，並略過格式不符的值（只接受星期字串或一層「名稱 → 星期字串」）
 * 與太長的鍵
 */
function toSafeConfig(raw) {
    const config = createDict();
    if (!raw || typeof raw !== 'object' || Array.isArray(raw)) return config;
    Object.keys(raw).forEach(key => {
        if (!isStorableConfigKey(key)) return;
        const value = raw[key];
        if (typeof value === 'string') {
            config[key] = value;
        } else if (value && typeof value === 'object' && !Array.isArray(value)) {
            const slotMap = createDict();
            Object.keys(value).forEach(club => {
                if (isStorableConfigKey(club) && typeof value[club] === 'string') slotMap[club] = value[club];
            });
            config[key] = slotMap;
        }
    });
    return config;
}

/**
 * 設定一個鍵並移到最後（最新），超過 maxKeys 時從最舊的刪除；鍵太長時不存，回傳 false
 * 純數字的鍵（例如「101」）在 JavaScript 物件中一律排在最前面，看不出新舊，淘汰時會先被刪；
 * 不改儲存格式就無法記住它們的先後，所以只保證剛設定的鍵本身不會被淘汰
 */
function setConfigEntry(obj, key, value, maxKeys) {
    if (!isStorableConfigKey(key)) return false;
    delete obj[key];
    obj[key] = value;
    const keys = Object.keys(obj).filter(k => k !== key);
    for (let i = 0; i < keys.length + 1 - maxKeys; i++) {
        delete obj[keys[i]];
    }
    return true;
}

function loadLocalStorageConfig() {
    try {
        const data = localStorage.getItem(CONFIG_STORAGE_KEY);
        return toSafeConfig(data ? JSON.parse(data) : null);
    } catch (e) {
        console.warn('[設定] 無法讀取瀏覽器記憶設定，改用預設值:', e);
        return createDict();
    }
}

function saveLocalStorageConfig(config) {
    try {
        localStorage.setItem(CONFIG_STORAGE_KEY, JSON.stringify(config));
    } catch (e) {
        console.warn('[設定] 無法儲存瀏覽器記憶設定:', e);
    }
}

function clearLocalStorageConfig() {
    try {
        localStorage.removeItem(CONFIG_STORAGE_KEY);
    } catch (e) {
        console.warn('[設定] 無法清除瀏覽器記憶設定:', e);
    }
}

/**
 * 延遲載入 ExcelJS（約 0.9 MB）：不放在 index.html 的 <script>，避免按鈕要等它下載完才有反應。
 * 頁面閒置時由 app.js 預先載入；讀取或匯出 Excel 前一律 await 這個函式。
 * 回傳快取的 Promise；載入失敗時 reject，並清掉快取讓下次可以重試。
 */
const EXCELJS_LOAD_ERROR_MESSAGE = '無法載入讀寫 Excel 所需的程式，可能是網路中斷。請確認網路連線後再試一次；第一次開啟這個工具時需要連上網路，之後才能離線使用。';
let excelJSPromise = null;

function loadExcelJS() {
    if (window.ExcelJS) return Promise.resolve(window.ExcelJS);
    if (excelJSPromise) return excelJSPromise;

    excelJSPromise = new Promise((resolve, reject) => {
        // ExcelJS 內附的 regenerator-runtime 在 strict mode 下執行 `regeneratorRuntime = runtime`，
        // 全域變數沒有宣告時會丟例外，接著改寫 globalThis，沒有 globalThis 的舊瀏覽器則退回
        // Function("r", "regeneratorRuntime = r")，那需要 CSP 'unsafe-eval'。
        // 先在 window 上宣告這個變數，賦值就會直接成功，不會走到 Function()。
        if (!('regeneratorRuntime' in window)) {
            window.regeneratorRuntime = undefined;
        }

        const script = document.createElement('script');
        const fail = () => {
            script.remove();
            excelJSPromise = null; // 下次呼叫重新載入
            const err = new Error(EXCELJS_LOAD_ERROR_MESSAGE);
            err.name = 'ExcelJSLoadError';
            reject(err);
        };
        script.src = EXCELJS_SRC;
        script.async = true;
        script.addEventListener('load', () => {
            if (window.ExcelJS) resolve(window.ExcelJS);
            else fail();
        });
        script.addEventListener('error', fail);
        document.head.appendChild(script);
    });
    return excelJSPromise;
}

/**
 * 檔案格式檢查：透過二進制檔頭 (Magic Bytes) 精準識別檔案真實結構
 * 避免舊版 .xls (BIFF8) 或 CSV 偽裝副檔名造成 ExcelJS 崩潰
 */
async function checkExcelFileFormat(file) {
    if (!file) return { valid: false, reason: '沒有選到檔案。請重新選擇名冊檔案。' };
    
    try {
        const buffer = await file.slice(0, 8).arrayBuffer();
        const bytes = new Uint8Array(buffer);
        
        // 1. 檢查 PK\x03\x04 (50 4B 03 04) - 標準 OpenXML (.xlsx / Zip)
        if (bytes[0] === 0x50 && bytes[1] === 0x4B && bytes[2] === 0x03 && bytes[3] === 0x04) {
            return { valid: true, type: 'xlsx' };
        }
        
        // 2. 檢查 D0 CF 11 E0 (OLE2 Compound Document) - 舊版 Excel 97-2003 (.xls)，
        //    設了開啟密碼的 .xlsx 也會加密包成 OLE2，檔頭相同
        if (bytes[0] === 0xD0 && bytes[1] === 0xCF && bytes[2] === 0x11 && bytes[3] === 0xE0) {
            return {
                valid: false,
                type: 'xls_legacy',
                reason: '這個檔案是舊版 Excel 97-2003（.xls）格式，或設了開啟密碼，無法讀取。\n請用 Excel 開啟（設了開啟密碼的先取消密碼），選擇「另存新檔」存成「.xlsx」後，再重新選擇。'
            };
        }

        // 3. 檢查副檔名是否被強制改名
        const fileName = file.name.toLowerCase();
        if (fileName.endsWith('.xls') && !fileName.endsWith('.xlsx')) {
            return {
                valid: false,
                type: 'xls_legacy',
                reason: '這個工具只能讀取「.xlsx」格式。\n請用 Excel 開啟這個「.xls」檔案，另存成「.xlsx」後，再重新選擇。'
            };
        }

        // 4. 其他未知格式
        return {
            valid: false,
            type: 'unknown',
            reason: '這個檔案無法當成 Excel 讀取，可能不是 Excel 檔或已損毀。請用 Excel 開啟確認內容，另存成 .xlsx 後再重新選擇。'
        };
    } catch (err) {
        console.warn('檢查檔案格式時發生錯誤:', err);
        // 若無法讀取 slice，退回副檔名檢查
        if (file.name.toLowerCase().endsWith('.xlsx')) {
            return { valid: true, type: 'xlsx' };
        }
        return { valid: false, type: 'unknown', reason: '無法讀取這個檔案，可能正被其他程式使用。請關閉正在開啟這個檔案的程式（例如 Excel），再重新選擇。' };
    }
}

/**
 * Excel 工作表名稱消毒器 (Worksheet Name Sanitizer)
 * Excel 規範：禁止出現 \ / ? * [ ] : ，不可前後帶單引號，最長不可超過 31 個字元
 */
function sanitizeSheetName(name, fallback = '工作表') {
    if (!name) return fallback;
    let clean = cleanInvisibleChars(name);
    // 移除非法字元
    clean = clean.replace(/[\\/?*\[\]:]/g, '_').trim();
    // 移除開頭或結尾的單引號
    clean = clean.replace(/^'+|'+$/g, '').trim();
    if (!clean) clean = fallback;
    // 截斷至 31 字元上限
    if (clean.length > 31) {
        clean = clean.substring(0, 31).trim();
    }
    return clean || fallback;
}

/**
 * 不可見字元淨化器
 * 清理剪貼簿常夾帶的看不見的字元（零寬空格、零寬連接字元、文字左右排列的控制字元、
 * 斷行用的連字號，以及 BOM），這些字元看不見，卻會讓同一個姓名或社團比對不到；
 * 另外把全形空白 (\u3000) 與 &nbsp; (\u00A0) 轉成一般空白
 */
function cleanInvisibleChars(str) {
    if (str === null || str === undefined) return "";
    return String(str)
        .replace(/[\u200B-\u200F\u2060\uFEFF\u00AD\u202A-\u202E\u2066-\u2069]/g, '')
        .replace(/\u00A0/g, ' ')
        .replace(/\u3000/g, ' ')
        .trim();
}

/**
 * 跨平台安全的檔案下載觸發器 (強化 iOS Safari / Mobile WebKit)
 */
function triggerFileDownload(blob, filename) {
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.style.display = 'none';
    a.href = url;
    a.download = filename;
    
    // 安全掛載至 DOM 觸發點擊
    document.body.appendChild(a);
    a.click();
    
    // 延遲回收 URL，確保行動端非同步完成下載請求
    setTimeout(() => {
        try {
            document.body.removeChild(a);
            window.URL.revokeObjectURL(url);
        } catch (e) {
            // 忽略回收異常
        }
    }, 2000);
}
