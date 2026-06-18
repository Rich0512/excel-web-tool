const WEEKDAY_REGEX = /(星期|週|周)([一二三四五六日1-7])/;
const weekdaysList = ["請選擇", "週一", "週二", "週三", "週四", "週五", "週六", "週日"];

// 解析年級數值 (新生為 0, 1-6 年級為 1-6)
function parseGrade(classStr) {
    if (!classStr || classStr === "新生") return 0;
    const match = classStr.match(/([一二三四五六1-6])/);
    if (match) {
        const char = match[1];
        const map = {
            '一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6,
            '1': 1, '2': 2, '3': 3, '4': 4, '5': 5, '6': 6
        };
        return map[char] || 0;
    }
    return 0;
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
 */
function getNumericSortKey(val) {
    if (val === null || val === undefined) return Infinity;
    const strVal = String(val).trim();
    if (strVal === "新生" || strVal === "") return Infinity;

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
 * 正規化星期名稱為 "週一" ~ "週日"
 */
function normalizeWeekdayName(dayStr) {
    if (!dayStr) return null;
    const match = dayStr.match(WEEKDAY_REGEX);
    if (match) {
        const dayChar = match[2];
        const dayMap = {
            '一': '週一', '1': '週一',
            '二': '週二', '2': '週二',
            '三': '週三', '3': '週三',
            '四': '週四', '4': '週四',
            '五': '週五', '5': '週五',
            '六': '週六', '6': '週六',
            '日': '週日', '7': '週日'
        };
        return dayMap[dayChar] || null;
    }
    return null;
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
    const day = normalizeWeekdayName(name);
    let slot = "";
    
    const slotMatch = name.match(/([A-Ba-b]|一時段|二時段|第一節|第二節|SlotA|SlotB)/i);
    if (slotMatch) {
        const rawSlot = slotMatch[1].toUpperCase();
        if (rawSlot === "A" || rawSlot === "一時段" || rawSlot === "第一節" || rawSlot === "SLOTA") {
            slot = "A";
        } else if (rawSlot === "B" || rawSlot === "二時段" || rawSlot === "第二節" || rawSlot === "SLOTB") {
            slot = "B";
        }
    }
    
    let cleanName = name;
    cleanName = cleanName.replace(/(星期|週|周)[一二三四五六日1-7]/g, "");
    cleanName = cleanName.replace(/([A-Ba-b]|一時段|二時段|第一節|第二節|時段|Slot|slot|社團|課程|項目)/gi, "");
    cleanName = cleanName.trim();
    
    if (!cleanName) cleanName = name.trim();
    
    return {
        day: day || "請選擇",
        slot: slot,
        clubName: cleanName
    };
}

// LocalStorage 管理對照設定
function loadLocalStorageConfig() {
    const data = localStorage.getItem('yenping_club_config');
    return data ? JSON.parse(data) : {};
}

function saveLocalStorageConfig(config) {
    localStorage.setItem('yenping_club_config', JSON.stringify(config));
}
