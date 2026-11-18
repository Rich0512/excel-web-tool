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
 * 隨機打亂陣列 (Fisher-Yates Shuffle)
 */
function shuffleArray(array) {
    for (let i = array.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [array[i], array[j]] = [array[j], array[i]];
    }
}

/**
 * 抽籤核心邏輯 (隨機抽籤與籤序分配)
 */
function runClubLottery(candidates, quota, gradeStart, gradeEnd, priority) {
    const minG = gradeStart === "新生" ? 0 : parseInt(gradeStart, 10);
    const maxG = gradeEnd === "新生" ? 0 : parseInt(gradeEnd, 10);

    const qualified = [];
    const disqualified = [];

    // 1. 年級範圍篩選
    candidates.forEach(s => {
        const g = parseGrade(s.class);
        if (g >= minG && g <= maxG) {
            qualified.push(s);
        } else {
            s.selected = false;
            s.drawSequence = "不符限制";
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

    // 為備取名單依序發放「備取 1, 2, 3...」
    backup.forEach((s, idx) => {
        s.drawSequence = `備取 ${idx + 1}`;
    });

    // 5. 排序錄取名單，方便瀏覽 (不影響已隨機分配好的籤序)
    selected.sort((a, b) => {
        const classA = getNumericSortKey(a.class);
        const classB = getNumericSortKey(b.class);
        if (classA !== classB) return classA - classB;

        const seatA = getNumericSortKey(a.seat);
        const seatB = getNumericSortKey(b.seat);
        if (seatA !== seatB) return seatA - seatB;

        return String(a.name).localeCompare(b.name, 'zh-hant');
    });

    // 6. 合併並排序整份備取名單 (回傳以便個別 Sheet 輸出)
    const backupList = [...backup, ...disqualified];
    backupList.sort((a, b) => {
        // 先按是否有備取順序（備取在前，不符資格在後）
        const aIsBackup = String(a.drawSequence).startsWith("備取");
        const bIsBackup = String(b.drawSequence).startsWith("備取");
        if (aIsBackup && !bIsBackup) return -1;
        if (!aIsBackup && bIsBackup) return 1;

        const classA = getNumericSortKey(a.class);
        const classB = getNumericSortKey(b.class);
        if (classA !== classB) return classA - classB;

        const seatA = getNumericSortKey(a.seat);
        const seatB = getNumericSortKey(b.seat);
        if (seatA !== seatB) return seatA - seatB;

        return String(a.name).localeCompare(b.name, 'zh-hant');
    });

    return [...selected, ...backupList];
}

/**
 * 彙整直接貼上的名冊資料 (僅包含被抽中的錄取學生)
 */
function processPastedClubsData(pastedClubs, slotMode, includeFreshmen) {
    const studentMap = new Map();

    pastedClubs.forEach(club => {
        const rawClubName = club.clubName;

        let slot = "";
        const slotMatch = rawClubName.match(/([A-Ba-b]|一時段|二時段|第一節|第二節|SlotA|SlotB)$/);
        if (slotMatch) {
            const rawSlot = slotMatch[1].toUpperCase();
            if (["A", "一時段", "第一節", "SLOTA"].includes(rawSlot)) slot = "A";
            else if (["B", "二時段", "第二節", "SLOTB"].includes(rawSlot)) slot = "B";
        }

        let cleanName = rawClubName.replace(/([A-Ba-b]|一時段|二時段|第一節|第二節|時段|Slot|slot)$/i, '').trim();
        if (!cleanName) cleanName = rawClubName;

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
