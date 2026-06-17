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
        // 🔹 模式 A：單一總表分欄彙整
        // ==========================================
        sheetData.forEach(row => {
            const classVal = row[detectedHeaders[colClassIdx]];
            const seatVal = row[detectedHeaders[colSeatIdx]];
            const nameVal = row[detectedHeaders[colNameIdx]];
            
            const studentClass = classVal || "新生";
            const studentSeat = seatVal || "";
            
            const student = {
                class: studentClass,
                seat: studentSeat,
                name: nameVal,
                schedule: { '週一': '', '週二': '', '週三': '', '週四': '', '週五': '', '週六': '', '週日': '' },
                remarks: []
            };

            slotCols.forEach(slot => {
                const origClub = row[slot.name];
                if (origClub && origClub !== "無") {
                    const mapData = finalMapping[origClub];
                    if (mapData) {
                        const day = mapData.day;
                        const displayName = mapData.editedName;
                        
                        activeDays.add(day);
                        
                        let label = "";
                        if (slotMode === 'vacation') {
                            label = slot.name.replace("時段", "").replace("Slot", "").replace("slot", "").replace("社團", "").replace("課程", "").trim();
                        }
                        const entryText = label ? `${displayName}(${label})` : displayName;
                        
                        if (student.schedule[day]) {
                            student.schedule[day] += `, ${entryText}`;
                            student.remarks.push(`「${day}」時段衝突：同時錄取「${student.schedule[day]}」`);
                        } else {
                            student.schedule[day] = entryText;
                        }
                    }
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
            const lSeatIdx = findColumnByKeywords(localHeaders, ['座號', '座', '號', 'Seat', 'seat']);
            const lNameIdx = findColumnByKeywords(localHeaders, ['姓名', '名', 'Name', 'name']);

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

                const studentClass = classVal || "新生";
                const studentSeat = seatVal || "";

                const studentKey = `${studentClass}_${nameVal}`;

                let student = studentMap[studentKey];
                if (!student) {
                    student = {
                        class: studentClass,
                        seat: studentSeat,
                        name: nameVal,
                        schedule: { '週一': '', '週二': '', '週三': '', '週四': '', '週五': '', '週六': '', '週日': '' },
                        remarks: []
                    };
                    studentMap[studentKey] = student;
                }

                const label = (slotMode === 'vacation') ? slot : "";
                const entryText = label ? `${displayName}(${label})` : displayName;

                if (student.schedule[day]) {
                    student.schedule[day] += `, ${entryText}`;
                    student.remarks.push(`「${day}」時段衝突：同時錄取「${student.schedule[day]}」`);
                } else {
                    student.schedule[day] = entryText;
                }
            });
        });

        Object.values(studentMap).forEach(s => resultData.push(s));
    }

    return { resultData, activeDays };
}

/**
 * 彙整直接貼上的名冊資料
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
            let sClass = student.class ? student.class.trim() : "";
            let sSeat = student.seat ? student.seat.trim() : "";
            let sName = student.name ? student.name.trim() : "";
            
            if (!sName) return;
            if (!includeFreshmen && !sClass) return;
            
            if (!sClass) sClass = "新生";
            
            const studentKey = `${sClass}_${sName}`;
            if (!studentMap[studentKey]) {
                studentMap[studentKey] = {
                    class: sClass,
                    seat: sSeat,
                    name: sName,
                    schedule: { '週一': '', '週二': '', '週三': '', '週四': '', '週五': '', '週六': '', '週日': '' },
                    remarks: []
                };
            }
            
            if (!studentMap[studentKey].seat && sSeat) {
                studentMap[studentKey].seat = sSeat;
            }
            
            if (studentMap[studentKey].schedule[day]) {
                studentMap[studentKey].schedule[day] += `, ${entryText}`;
                studentMap[studentKey].remarks.push(`「${day}」時段衝突：同時錄取「${studentMap[studentKey].schedule[day]}」`);
            } else {
                studentMap[studentKey].schedule[day] = entryText;
            }
        });
    });

    const resultData = Object.values(studentMap);
    return { resultData, activeDays };
}
