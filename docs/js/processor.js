/**
 * 學生社團課表彙整核心處理器 (Business Logic / Model)
 */

// 匯出與衝堂備註使用的星期順序
const DAY_ORDER = ['週一', '週二', '週三', '週四', '週五', ...MORNING_OPTIONS, '週六', '週日'];

function dayOrderIndex(day) {
    const idx = DAY_ORDER.indexOf(day);
    return idx === -1 ? DAY_ORDER.length : idx;
}

/**
 * 取得（或建立）一位學生。以「班級＋姓名」合併同一位學生；同班同名但座號不同時視為不同人
 * 沒有姓名回傳 null；沒有班級時視為新生
 */
function upsertStudent(studentMap, rawClass, rawSeat, rawName) {
    const name = cleanInvisibleChars(rawName);
    if (!name) return null;
    const cls = cleanInvisibleChars(rawClass) || '新生';
    const seat = cleanInvisibleChars(rawSeat);

    let key = `${cls}_${name}`;
    const sameName = studentMap.get(key);
    if (seat && sameName && sameName.seat && sameName.seat !== seat) {
        key = `${cls}_${seat}_${name}`;
    }

    let student = studentMap.get(key);
    if (!student) {
        student = { class: cls, seat, name, days: {} };
        studentMap.set(key, student);
    } else if (!student.seat && seat) {
        student.seat = seat;
    }
    return student;
}

/**
 * 把社團排進學生的上課日。每天的社團以陣列保存，社團名稱含逗號也不會被誤判成衝堂
 */
function addClubToStudent(student, dayStr, entryText) {
    parseDays(dayStr).forEach(d => {
        const clubs = student.days[d] || (student.days[d] = []);
        if (!clubs.includes(entryText)) clubs.push(entryText);
    });
}

/**
 * 轉成匯出格式：schedule 是「星期 → 社團（逗號分隔）」，remarks 是衝堂說明
 * 晨間社團和下午的社團不同時段，不算衝堂
 */
function finalizeStudents(studentMap) {
    const resultData = [];
    const activeDays = new Set();

    studentMap.forEach(s => {
        const schedule = {};
        const remarks = [];
        Object.keys(s.days)
            .sort((a, b) => dayOrderIndex(a) - dayOrderIndex(b))
            .forEach(day => {
                const clubs = s.days[day];
                activeDays.add(day);
                schedule[day] = clubs.join(', ');
                if (!day.startsWith('晨間') && clubs.length > 1) {
                    remarks.push(`「${day}」時段衝突：同時錄取「${schedule[day]}」`);
                }
            });
        resultData.push({ class: s.class, seat: s.seat, name: s.name, schedule, remarks });
    });

    return { resultData, activeDays };
}

/**
 * 彙整上傳的 Excel 資料
 * - single：單一總表，每列是「一位學生 × 各時段欄位的社團」，對照鍵是 mappingKey(時段欄位, 社團)
 * - multi：每個社團一張工作表，學生已在載入時由 readClubSheetStudents() 讀進 sheetStudents，對照鍵是工作表名稱
 */
function processExcelData({ mode, includeFreshmen, slotMode, finalMapping, sheetData, classCol, seatCol, nameCol, slotCols, sheetStudents }) {
    const studentMap = new Map();

    if (mode === 'single') {
        sheetData.forEach(row => {
            if (!includeFreshmen && !cleanInvisibleChars(row[classCol])) return;
            const student = upsertStudent(studentMap, row[classCol], row[seatCol], row[nameCol]);
            if (!student) return;

            slotCols.forEach(slot => {
                const club = (row[slot.name] || "").trim();
                if (!club || club === "無") return;
                const mapData = finalMapping[mappingKey(slot.name, club)];
                if (!mapData) return;
                const label = slotMode === 'vacation' ? slotLabelFromColumn(slot.name) : "";
                addClubToStudent(student, mapData.day, label ? `${mapData.editedName}(${label})` : mapData.editedName);
            });
        });
    } else {
        Object.entries(finalMapping).forEach(([sheetName, mapData]) => {
            const students = sheetStudents.get(sheetName);
            if (!students) return;
            const label = slotMode === 'vacation' ? mapData.slot : "";
            const entryText = label ? `${mapData.editedName}(${label})` : mapData.editedName;

            students.forEach(s => {
                if (!includeFreshmen && !s.class) return;
                const student = upsertStudent(studentMap, s.class, s.seat, s.name);
                if (student) addClubToStudent(student, mapData.day, entryText);
            });
        });
    }

    return finalizeStudents(studentMap);
}

/**
 * 貼上名單的重複報名去重（抽籤前執行）
 * 判斷規則與 upsertStudent() 一致，因為彙整課表時它會把這些人併成同一位：
 * 班級（清除不可見字元，空白當「新生」）與姓名相同，而且座號相同或其中一方沒有座號，就算重複。
 * 保留第一筆；第一筆沒有座號、後面那筆有時，改用後面那筆的座號。沒有姓名的資料列不參與比對，原樣保留。
 * 回傳 { students: 去重後的名單, duplicateCount: 略過的筆數 }
 */
function dedupeApplicants(students) {
    const kept = [];
    const byKey = new Map(); // 班級 + 姓名 → 已保留的同名學生（同班同名但座號不同時可能有多位）
    let duplicateCount = 0;

    students.forEach(s => {
        const name = cleanInvisibleChars(s.name);
        if (!name) {
            kept.push(s);
            return;
        }
        const key = `${cleanInvisibleChars(s.class) || '新生'}\u0001${name}`;
        const seat = cleanInvisibleChars(s.seat);
        const sameName = byKey.get(key) || [];
        const match = sameName.find(k => {
            const keptSeat = cleanInvisibleChars(k.seat);
            return !seat || !keptSeat || keptSeat === seat;
        });
        if (match) {
            duplicateCount++;
            if (!cleanInvisibleChars(match.seat) && seat) match.seat = seat;
            return;
        }
        sameName.push(s);
        byKey.set(key, sameName);
        kept.push(s);
    });

    return { students: kept, duplicateCount };
}

/**
 * 隨機打亂陣列 (Fisher-Yates Shuffle)，亂數來自 randomInt()（瀏覽器內建密碼學亂數，每種排列的機率相同）
 */
function shuffleArray(array) {
    for (let i = array.length - 1; i > 0; i--) {
        const j = randomInt(i + 1);
        [array[i], array[j]] = [array[j], array[i]];
    }
}

// 班級讀不出年級的學生不參加抽籤，籤序欄顯示這段文字
const UNRECOGNIZED_CLASS_LABEL = "班級無法辨識";
const DISQUALIFIED_LABEL = "不符限制";

// 依班級、座號、姓名排序
function compareClassSeatName(a, b) {
    const classA = getNumericSortKey(a.class);
    const classB = getNumericSortKey(b.class);
    if (classA !== classB) return classA - classB;

    const seatA = getNumericSortKey(a.seat);
    const seatB = getNumericSortKey(b.seat);
    if (seatA !== seatB) return seatA - seatB;

    return String(a.name).localeCompare(b.name, 'zh-hant');
}

/**
 * 抽籤核心邏輯 (隨機抽籤與籤序分配)
 * 回傳順序：正取（先依年級，新生最後；同年級再依班級、座號、姓名）→ 備取（依備取 1、2、3…）→ 班級無法辨識 → 不符限制（後兩組依班級、座號、姓名）
 */
function runClubLottery(candidates, quota, gradeStart, gradeEnd, priority) {
    const minG = gradeStart === "新生" ? 0 : parseInt(gradeStart, 10);
    const maxG = gradeEnd === "新生" ? 0 : parseInt(gradeEnd, 10);

    const qualified = [];
    const unrecognized = [];
    const disqualified = [];

    // 1. 年級範圍篩選；班級讀不出年級（parseGrade 回傳 null）的不參加抽籤
    candidates.forEach(s => {
        const g = parseGrade(s.class);
        if (g === null) {
            s.selected = false;
            s.drawSequence = UNRECOGNIZED_CLASS_LABEL;
            unrecognized.push(s);
        } else if (g >= minG && g <= maxG) {
            qualified.push(s);
        } else {
            s.selected = false;
            s.drawSequence = DISQUALIFIED_LABEL;
            disqualified.push(s);
        }
    });

    // 2. 按年級分組
    const groups = {};
    qualified.forEach(s => {
        const g = parseGrade(s.class);
        if (!groups[g]) groups[g] = [];
        groups[g].push(s);
    });

    // 3. 排序年級優先順序 (以高年級優先或低年級優先)
    const grades = Object.keys(groups).map(Number);
    grades.sort((a, b) => {
        if (a === 0) return 1; // 新生在最後面
        if (b === 0) return -1;
        if (priority === 'desc') {
            return b - a; // 高年級優先 (6 > 5 > ...)
        } else {
            return a - b; // 低年級優先 (1 > 2 > ...)
        }
    });

    // 4. 依序對各年級隨機抽籤
    const selected = [];
    const backup = [];

    grades.forEach(g => {
        const list = groups[g];
        shuffleArray(list);
        list.forEach(s => {
            if (selected.length < quota) {
                s.selected = true;
                // 籤序即為隨機抽中的順序 (1, 2, 3...)
                s.drawSequence = selected.length + 1;
                selected.push(s);
            } else {
                s.selected = false;
                backup.push(s);
            }
        });
    });

    // 為備取名單依序發放「備取 1, 2, 3...」；backup 的順序就是備取順序，之後不再排序
    backup.forEach((s, idx) => {
        s.drawSequence = `備取 ${idx + 1}`;
    });

    // 5. 排序錄取名單，方便瀏覽 (不影響已隨機分配好的籤序)
    // 正取先依年級排在一起（1 到 6 年級，新生最後），同年級再依班級、座號、姓名。
    // 匯出檔的「錄取人數」欄會把連續的同年級合併，所以同年級要排在一起（班級寫法不同時也一樣，例如「301」「3年1班」「三年一班」）
    const gradeRank = s => { const g = parseGrade(s.class); return g === 0 ? 7 : g; };
    selected.sort((a, b) => gradeRank(a) - gradeRank(b) || compareClassSeatName(a, b));
    unrecognized.sort(compareClassSeatName);
    disqualified.sort(compareClassSeatName);

    return [...selected, ...backup, ...unrecognized, ...disqualified];
}

/**
 * 彙整直接貼上的名冊資料 (僅包含被抽中的錄取學生)
 */
function processPastedClubsData(pastedClubs, slotMode, includeFreshmen) {
    const studentMap = new Map();

    pastedClubs.forEach(club => {
        // 名稱結尾的時段後綴（「桌遊A」「桌遊（B）」「足球第一時段」）同時決定時段與清掉的文字
        const { name: cleanName, slot } = splitSlotSuffix(club.clubName);

        const label = slotMode === 'vacation' ? slot : "";
        const entryText = label ? `${cleanName}(${label})` : cleanName;

        club.students.forEach(s => {
            // 僅彙整正式被抽中錄取的學生！
            if (s.selected === false) return;
            if (!includeFreshmen && !cleanInvisibleChars(s.class)) return;
            const student = upsertStudent(studentMap, s.class, s.seat, s.name);
            if (student) addClubToStudent(student, club.day, entryText);
        });
    });

    return finalizeStudents(studentMap);
}
