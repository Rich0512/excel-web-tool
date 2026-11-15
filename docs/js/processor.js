/**
 * ⚙️ 學生社團課表彙整核心處理器 (Business Logic / Model)
 */

/**
 * 彙整上傳的 Excel 資料 (模式 A & 模式 B)
 */
function processExcelData(mode, includeFreshmen, slotMode, finalMapping, sheetData, detectedHeaders, colClassIdx, colSeatIdx, colNameIdx, slotCols) {
    const resultData = [];
    const activeDays = new Set();

    if (mode === 'single') {
        // ==========================================
        // 🔹 模式 A：單一總表分欄彙整 (依學生合併，同一位學生排好一到五課表，僅佔一列)
        // ==========================================
        const studentMap = {};

        sheetData.forEach(row => {
            const rawClass = row[detectedHeaders[colClassIdx]];
            const rawSeat = row[detectedHeaders[colSeatIdx]];
            const rawName = row[detectedHeaders[colNameIdx]];
            
            const cleanFn = typeof cleanInvisibleChars === 'function' ? cleanInvisibleChars : (s) => String(s || "").trim();
            const nameVal = cleanFn(rawName);
            if (!nameVal) return;
            
            const rawClassClean = cleanFn(rawClass);
            const studentClass = rawClassClean || "新生";
            const studentSeat = cleanFn(rawSeat);
            
            let studentKey = `${studentClass}_${nameVal}`;
            if (studentSeat && studentMap[studentKey] && studentMap[studentKey].seat && studentMap[studentKey].seat !== studentSeat) {
                studentKey = `${studentClass}_${studentSeat}_${nameVal}`;
            }

            let student = studentMap[studentKey];
            if (!student) {
                student = {
                    class: studentClass,
                    seat: studentSeat,
                    name: nameVal,
                    schedule: { '週一': '', '週二': '', '週三': '', '週四': '', '週五': '', '晨間社團': '', '週六': '', '週日': '' },
                    remarks: []
                };
                studentMap[studentKey] = student;
            } else if (!student.seat && studentSeat) {
                student.seat = studentSeat;
            }

            slotCols.forEach(slot => {
                const origClub = (row[slot.name] || "").trim();
                if (origClub && origClub !== "無") {
                    const mapData = finalMapping[origClub];
                    if (mapData) {
                        const day = mapData.day;
                        const displayName = mapData.editedName;
                        
                        let label = "";
                        if (slotMode === 'vacation') {
                            label = slot.name.replace("時段", "").replace("Slot", "").replace("slot", "").replace("社團", "").replace("課程", "").trim();
                        }
                        const entryText = label ? `${displayName}(${label})` : displayName;
                        
                        const targetDays = parseDays(day);
                        targetDays.forEach(d => {
                            activeDays.add(d);
                            if (student.schedule[d]) {
                                const existing = student.schedule[d].split(',').map(s => s.trim());
                                if (!existing.includes(entryText)) {
                                    student.schedule[d] += `, ${entryText}`;
                                }
                            } else {
                                student.schedule[d] = entryText;
                            }
                        });
                    }
                }
            });
        });

        // 統一生產衝突備註 (避免重複堆疊)
        Object.values(studentMap).forEach(student => {
            student.remarks = [];
            Object.keys(student.schedule).forEach(day => {
                if (day.startsWith('晨間')) return;
                if (student.schedule[day] && student.schedule[day].includes(',')) {
                    student.remarks.push(`「${day}」時段衝突：同時錄取「${student.schedule[day]}」`);
                }
            });
            resultData.push(student);
        });

    } else {
        // ==========================================
        // 🔸 模式 B：多工作表分社團彙整
        // ==========================================
        const studentMap = {};

        Object.entries(finalMapping).forEach(([sheetName, mapData]) => {
            const sheet = uploadedWorkbook.getWorksheet(sheetName);
            if (!sheet) return;

            const day = mapData.day;
            const displayName = mapData.editedName;
            const slot = mapData.slot;

            activeDays.add(day);

            // 在每個子 Sheet 智慧定位表頭
            let localHeaders = [];
            let localHeaderRowIndex = 1;
            let foundHeader = false;

            for (let r = 1; r <= 10; r++) {
                const row = sheet.getRow(r);
                const rowValues = [];
                row.eachCell({ includeEmpty: true }, (cell) => {
                    rowValues.push(getCellValueAsString(cell));
                });
                
                const matchCount = ['班級', '班', '座號', '姓名'].filter(kw => 
                    rowValues.some(val => val.includes(kw))
                ).length;
                
                if (matchCount >= 2) {
                    localHeaders = rowValues;
                    localHeaderRowIndex = r;
                    foundHeader = true;
                    break;
                }
            }

            if (!foundHeader) {
                const row = sheet.getRow(1);
                row.eachCell({ includeEmpty: true }, (cell) => {
                    localHeaders.push(getCellValueAsString(cell));
                });
                localHeaderRowIndex = 1;
            }

            const lClassIdx = findColumnByKeywords(localHeaders, ['班級', '班', 'Class', 'class']);
            const lSeatIdx = findColumnByKeywords(localHeaders, ['座號', '座', 'Seat', 'seat'], true, false);
            const lNameIdx = findColumnByKeywords(localHeaders, ['學生姓名', '姓名', 'Name', 'name', '名'], false, true);

            if (lClassIdx === -1 || lSeatIdx === -1 || lNameIdx === -1) {
                console.warn(`工作表「${sheetName}」缺少必要欄位，已跳過。`);
                return;
            }

            sheet.eachRow((row, rowNum) => {
                if (rowNum <= localHeaderRowIndex) return;

                const nameVal = getCellValueAsString(row.getCell(lNameIdx + 1));
                const classVal = getCellValueAsString(row.getCell(lClassIdx + 1));
                const seatVal = getCellValueAsString(row.getCell(lSeatIdx + 1));

                if (!nameVal) return;
                if (!includeFreshmen && !classVal) return;

                if (['合計', '總計', '統計', '人數', '小計'].some(k => nameVal.includes(k) || (classVal && classVal.includes(k)))) {
                    return;
                }

                const cleanFn = typeof cleanInvisibleChars === 'function' ? cleanInvisibleChars : (s) => String(s || "").trim();
                const nameClean = cleanFn(nameVal);
                if (!nameClean) return;

                const rawClassClean = cleanFn(classVal);
                const studentClass = rawClassClean || "新生";
                const studentSeat = cleanFn(seatVal);

                let studentKey = `${studentClass}_${nameClean}`;
                if (studentSeat && studentMap[studentKey] && studentMap[studentKey].seat && studentMap[studentKey].seat !== studentSeat) {
                    studentKey = `${studentClass}_${studentSeat}_${nameClean}`;
                }

                let student = studentMap[studentKey];
                if (!student) {
                    student = {
                        class: studentClass,
                        seat: studentSeat,
                        name: nameClean,
                        schedule: { '週一': '', '週二': '', '週三': '', '週四': '', '週五': '', '晨間社團': '', '週六': '', '週日': '' },
                        remarks: []
                    };
                    studentMap[studentKey] = student;
                } else if (!student.seat && studentSeat) {
                    student.seat = studentSeat;
                }

                const label = (slotMode === 'vacation') ? slot : "";
                const entryText = label ? `${displayName}(${label})` : displayName;

                const targetDays = parseDays(day);
                targetDays.forEach(d => {
                    activeDays.add(d);
                    if (student.schedule[d]) {
                        const existing = student.schedule[d].split(',').map(s => s.trim());
                        if (!existing.includes(entryText)) {
                            student.schedule[d] += `, ${entryText}`;
                        }
                    } else {
                        student.schedule[d] = entryText;
                    }
                });
            });
        });

        // 統一生產衝突備註
        Object.values(studentMap).forEach(student => {
            student.remarks = [];
            Object.keys(student.schedule).forEach(day => {
                if (day.startsWith('晨間')) return;
                if (student.schedule[day] && student.schedule[day].includes(',')) {
                    student.remarks.push(`「${day}」時段衝突：同時錄取「${student.schedule[day]}」`);
                }
            });
            resultData.push(student);
        });
    }

    return { resultData, activeDays };
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
    const studentMap = {};
    const activeDays = new Set(['週一', '週二', '週三', '週四', '週五']);

    pastedClubs.forEach(club => {
        const day = club.day;
        const rawClubName = club.clubName;
        activeDays.add(day);
        
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
        
        club.students.forEach(student => {
            // 🚨 僅彙整正式被抽中錄取的學生！
            if (student.selected === false) return;
            
            const cleanFn = typeof cleanInvisibleChars === 'function' ? cleanInvisibleChars : (s) => String(s || "").trim();
            let sClass = cleanFn(student.class);
            let sSeat = cleanFn(student.seat);
            let sName = cleanFn(student.name);
            
            if (!sName) return;
            if (!includeFreshmen && !sClass) return;
            
            if (!sClass) sClass = "新生";
            
            let studentKey = `${sClass}_${sName}`;
            if (sSeat && studentMap[studentKey] && studentMap[studentKey].seat && studentMap[studentKey].seat !== sSeat) {
                studentKey = `${sClass}_${sSeat}_${sName}`;
            }

            let studentObj = studentMap[studentKey];
            if (!studentObj) {
                studentObj = {
                    class: sClass,
                    seat: sSeat,
                    name: sName,
                    schedule: { '週一': '', '週二': '', '週三': '', '週四': '', '週五': '', '晨間社團': '', '週六': '', '週日': '' },
                    remarks: []
                };
                studentMap[studentKey] = studentObj;
            } else if (!studentObj.seat && sSeat) {
                studentObj.seat = sSeat;
            }
            
            const targetDays = parseDays(day);
            targetDays.forEach(d => {
                activeDays.add(d);
                if (studentObj.schedule[d]) {
                    const existing = studentObj.schedule[d].split(',').map(s => s.trim());
                    if (!existing.includes(entryText)) {
                        studentObj.schedule[d] += `, ${entryText}`;
                    }
                } else {
                    studentObj.schedule[d] = entryText;
                }
            });
        });
    });

    Object.values(studentMap).forEach(student => {
        student.remarks = [];
        Object.keys(student.schedule).forEach(day => {
            if (day.startsWith('晨間')) return;
            if (student.schedule[day] && student.schedule[day].includes(',')) {
                student.remarks.push(`「${day}」時段衝突：同時錄取「${student.schedule[day]}」`);
            }
        });
    });

    const resultData = Object.values(studentMap);
    return { resultData, activeDays };
}
