// 全域變數
let uploadedWorkbook = null;
let originalFileName = "";
let sheetData = [];
let detectedHeaders = [];
let headerRowIndex = 1;

let colClassIdx = -1;
let colSeatIdx = -1;
let colNameIdx = -1;
let slotCols = [];

const dropZone = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-input');

// ==========================================
// 📥 檔案拖曳與讀取邏輯
// ==========================================

// 拖曳視覺效果
['dragenter', 'dragover'].forEach(eventName => {
    dropZone.addEventListener(eventName, (e) => {
        e.preventDefault();
        dropZone.classList.add('dragover');
    }, false);
});

['dragleave', 'drop'].forEach(eventName => {
    dropZone.addEventListener(eventName, (e) => {
        e.preventDefault();
        dropZone.classList.remove('dragover');
    }, false);
});

// 檔案上傳觸發
dropZone.addEventListener('drop', (e) => {
    const files = e.dataTransfer.files;
    if (files.length > 0) handleFileSelect(files[0]);
});

fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) handleFileSelect(e.target.files[0]);
});

// 讀取 Excel 檔案
async function handleFileSelect(file) {
    originalFileName = file.name;
    const reader = new FileReader();
    
    reader.onload = async function(e) {
        const arrayBuffer = e.target.result;
        const workbook = new ExcelJS.Workbook();
        
        try {
            await workbook.xlsx.load(arrayBuffer);
            uploadedWorkbook = workbook;
            
            // 偵測是否存在「總表」工作表
            let hasTotalSheet = false;
            workbook.eachSheet((ws) => {
                if (ws.name.trim() === '總表') {
                    hasTotalSheet = true;
                }
            });
            
            // 依有無「總表」自動切換預設模式
            document.getElementById('mode-select').value = hasTotalSheet ? 'single' : 'multi';
            
            // 分析並渲染對照設定
            analyzeAndRenderMapping();
            switchStep('step-mapping');

        } catch (err) {
            alert("載入 Excel 失敗，可能檔案格式毀損：" + err.message);
        }
    };
    
    reader.readAsArrayBuffer(file);
}

// ==========================================
// ⚙️ 核心解析與渲染設定表
// ==========================================

function analyzeAndRenderMapping() {
    if (!uploadedWorkbook) return;
    const mode = document.getElementById('mode-select').value;
    const tbody = document.getElementById('mapping-table-body');
    tbody.innerHTML = "";

    const savedConfig = loadLocalStorageConfig();

    if (mode === 'single') {
        // --- 單一總表分欄模式 ---
        let sheet = null;
        uploadedWorkbook.eachSheet((ws) => {
            if (ws.name.trim() === '總表') {
                sheet = ws;
            }
        });

        if (!sheet) {
            alert("在 Excel 中找不到「總表」工作表！請改用「多工作表模式」。");
            document.getElementById('mode-select').value = 'multi';
            analyzeAndRenderMapping();
            return;
        }

        // 智慧偵測表頭
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
                detectedHeaders = rowValues;
                headerRowIndex = r;
                foundHeader = true;
                break;
            }
        }
        
        if (!foundHeader) {
            const row = sheet.getRow(1);
            detectedHeaders = [];
            row.eachCell({ includeEmpty: true }, (cell) => {
                detectedHeaders.push(getCellValueAsString(cell));
            });
            headerRowIndex = 1;
        }
        
        // 去重
        const headers = [];
        const seenCols = {};
        detectedHeaders.forEach(col => {
            if (!col) {
                headers.push("");
                return;
            }
            if (seenCols[col]) {
                seenCols[col]++;
                headers.push(`${col}_${seenCols[col]}`);
            } else {
                seenCols[col] = 1;
                headers.push(col);
            }
        });
        detectedHeaders = headers;

        colClassIdx = findColumnByKeywords(headers, ['班級', '班', 'Class', 'class']);
        colSeatIdx = findColumnByKeywords(headers, ['座號', '座', '號', 'Seat', 'seat']);
        colNameIdx = findColumnByKeywords(headers, ['姓名', '名', 'Name', 'name']);

        if (colClassIdx === -1 || colSeatIdx === -1 || colNameIdx === -1) {
            alert("「總表」中缺少必要欄位（班級、座號、姓名），無法進行整理。");
            return;
        }

        // 尋找時段欄位
        slotCols = [];
        const slotKeywords = ['時段', 'Slot', 'slot', '課程', '項目', '節', '社團'];
        headers.forEach((col, idx) => {
            if (idx === colClassIdx || idx === colSeatIdx || idx === colNameIdx) return;
            if (slotKeywords.some(kw => col.includes(kw))) {
                slotCols.push({ name: col, index: idx });
            }
        });

        if (slotCols.length === 0) {
            alert("找不到任何包含「時段」、「課程」或「節」的社團時段欄位。");
            return;
        }

        const excelSchedule = loadScheduleFromExcel(uploadedWorkbook);

        // 讀取學生行
        sheetData = [];
        sheet.eachRow((row, rowNum) => {
            if (rowNum <= headerRowIndex) return;
            
            const nameVal = getCellValueAsString(row.getCell(colNameIdx + 1));
            const classVal = getCellValueAsString(row.getCell(colClassIdx + 1));
            
            if (!nameVal) return;
            const includeFreshmen = document.getElementById('include-freshmen').checked;
            if (!includeFreshmen && !classVal) return;
            
            // 略過合計行
            if (['合計', '總計', '統計', '人數', '小計'].some(k => nameVal.includes(k) || (classVal && classVal.includes(k)))) {
                return;
            }
            
            const rowData = {};
            headers.forEach((colName, idx) => {
                if (!colName) return;
                const cell = row.getCell(idx + 1);
                rowData[colName] = getCellValueAsString(cell);
            });
            sheetData.push(rowData);
        });

        // 蒐集時段社團
        const activeClubs = {};
        slotCols.forEach(slot => {
            activeClubs[slot.name] = new Set();
            sheetData.forEach(row => {
                const val = row[slot.name];
                if (val && val !== "無") {
                    activeClubs[slot.name].add(val);
                }
            });
        });

        // 渲染對照表 (總表分欄模式)
        Object.entries(activeClubs).forEach(([slotCol, clubs]) => {
            const isASlot = slotCol.includes('A') || slotCol.includes('一');
            const badgeClass = isASlot ? "slot-badge slot-a" : "slot-badge slot-other";
            const savedSlotMap = savedConfig[slotCol] || {};

            clubs.forEach(club => {
                const tr = document.createElement('tr');
                tr.className = "mapping-row-single";
                tr.dataset.slot = slotCol;
                
                let day = extractWeekdayFromName(club);
                if (!day && excelSchedule) day = excelSchedule[club];
                if (!day) day = savedConfig[club];
                if (!day) day = savedSlotMap[club];

                const finalDay = weekdaysList.includes(day) ? day : "請選擇";

                // 時段
                const tdSlot = document.createElement('td');
                tdSlot.innerHTML = `<span class="${badgeClass}">${slotCol}</span>`;
                tr.appendChild(tdSlot);

                // 社團名輸入框
                const tdName = document.createElement('td');
                const input = document.createElement('input');
                input.type = "text";
                input.className = "club-input";
                input.value = club;
                input.dataset.slot = slotCol;
                input.dataset.original = club;
                tdName.appendChild(input);
                tr.appendChild(tdName);

                // 星期選擇
                const tdDay = document.createElement('td');
                const select = document.createElement('select');
                select.className = "day-select";
                
                weekdaysList.forEach(w => {
                    const opt = document.createElement('option');
                    opt.value = w;
                    opt.textContent = w;
                    if (w === finalDay) opt.selected = true;
                    select.appendChild(opt);
                });
                
                tdDay.appendChild(select);
                tr.appendChild(tdDay);

                tbody.appendChild(tr);
            });
        });

    } else {
        // --- 多工作表分社團模式 ---
        const sheetsToProcess = [];
        const ignoreKeywords = ['對照', '說明', 'README', '總表', '統計', '人數', '小計'];

        uploadedWorkbook.eachSheet((ws) => {
            const name = ws.name.trim();
            if (ignoreKeywords.some(kw => name.includes(kw))) {
                return;
            }
            sheetsToProcess.push(ws);
        });

        if (sheetsToProcess.length === 0) {
            alert("在 Excel 中找不到任何可以彙整的社團名單工作表！");
            return;
        }

        sheetsToProcess.forEach((ws) => {
            const name = ws.name.trim();
            const parsed = parseSheetName(name);

            // 讀取歷史星期設定
            let day = savedConfig[name];
            if (!day) day = parsed.day;

            const finalDay = weekdaysList.includes(day) ? day : "請選擇";

            const tr = document.createElement('tr');
            tr.className = "mapping-row-multi";
            tr.dataset.sheetName = name;
            tr.dataset.detectedSlot = parsed.slot;

            // 時段 (多 Sheet 模式下顯示工作表特徵)
            const tdSlot = document.createElement('td');
            const displaySlot = parsed.slot ? `時段 ${parsed.slot}` : "多工作表";
            const badgeClass = parsed.slot ? (parsed.slot === 'A' ? "slot-badge slot-a" : "slot-badge slot-other") : "slot-badge slot-other";
            tdSlot.innerHTML = `<span class="${badgeClass}">${displaySlot}</span>`;
            tr.appendChild(tdSlot);

            // 社團名輸入框
            const tdName = document.createElement('td');
            const input = document.createElement('input');
            input.type = "text";
            input.className = "club-input";
            input.value = parsed.clubName;
            input.dataset.original = name;
            tdName.appendChild(input);
            tr.appendChild(tdName);

            // 星期選擇
            const tdDay = document.createElement('td');
            const select = document.createElement('select');
            select.className = "day-select";
            
            weekdaysList.forEach(w => {
                const opt = document.createElement('option');
                opt.value = w;
                opt.textContent = w;
                if (w === finalDay) opt.selected = true;
                select.appendChild(opt);
            });
            
            tdDay.appendChild(select);
            tr.appendChild(tdDay);

            tbody.appendChild(tr);
        });
    }

    // 重新註冊鍵盤監聽
    window.removeEventListener('keydown', handleGlobalKeydown);
    window.addEventListener('keydown', handleGlobalKeydown);
}

function handleGlobalKeydown(e) {
    if (e.key === "Enter") {
        processAndDownload();
    } else if (e.key === "Escape") {
        resetToUpload();
    }
}

// 切換步驟顯示
function switchStep(stepId) {
    document.querySelectorAll('.step-panel').forEach(panel => {
        panel.classList.remove('active');
    });
    document.getElementById(stepId).classList.add('active');
    
    if (stepId !== 'step-mapping') {
        window.removeEventListener('keydown', handleGlobalKeydown);
    }
}

function resetToUpload() {
    fileInput.value = "";
    switchStep('step-upload');
}

// ==========================================
// ⚡ 資料轉換與下載邏輯 (檔案上傳模式)
// ==========================================

async function processAndDownload() {
    const mode = document.getElementById('mode-select').value;
    const includeFreshmen = document.getElementById('include-freshmen').checked;
    const slotMode = document.getElementById('slot-mode-select').value;

    const rows = document.getElementById('mapping-table-body').querySelectorAll('tr');
    
    const finalMapping = {};
    const unselectedClubs = [];

    rows.forEach(tr => {
        const input = tr.querySelector('.club-input');
        const select = tr.querySelector('.day-select');
        
        const origKey = input.dataset.original;
        const editedClub = input.value.trim();
        const selectedDay = select.value;

        if (!editedClub) {
            alert("社團名稱不能為空！");
            return;
        }

        if (selectedDay === "請選擇") {
            if (mode === 'single') {
                const slotCol = input.dataset.slot;
                unselectedClubs.push(`${slotCol} 的「${origKey}」`);
            } else {
                unselectedClubs.push(`工作表「${origKey}」`);
            }
        }

        finalMapping[origKey] = {
            editedName: editedClub,
            day: selectedDay,
            slot: tr.dataset.detectedSlot || ""
        };
    });

    if (unselectedClubs.length > 0) {
        alert(`請為以下社團選擇上課星期後再行轉換：\n\n${unselectedClubs.join('\n')}`);
        return;
    }

    // 自動更新並儲存 LocalStorage 設定
    const savedConfig = loadLocalStorageConfig();
    Object.entries(finalMapping).forEach(([origKey, data]) => {
        savedConfig[origKey] = data.day;
    });
    saveLocalStorageConfig(savedConfig);

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

        rows.forEach(tr => {
            const sheetName = tr.dataset.sheetName;
            const mapData = finalMapping[sheetName];
            if (!mapData) return;

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

    if (resultData.length === 0) {
        alert("無任何有效學生資料可供匯出，請確認上傳檔案。");
        return;
    }

    // 排序
    resultData.sort((a, b) => {
        const numClassA = getNumericSortKey(a.class);
        const numClassB = getNumericSortKey(b.class);
        if (numClassA !== numClassB) {
            return numClassA - numClassB;
        }
        
        const numSeatA = getNumericSortKey(a.seat);
        const numSeatB = getNumericSortKey(b.seat);
        if (numSeatA !== numSeatB) {
            return numSeatA - numSeatB;
        }

        const strClassA = String(a.class);
        const strClassB = String(b.class);
        if (strClassA !== strClassB) {
            return strClassA.localeCompare(strClassB, 'zh-hant');
        }

        const strSeatA = String(a.seat);
        const strSeatB = String(b.seat);
        return strSeatA.localeCompare(strSeatB, 'zh-hant');
    });

    await exportWeeklySchedule(resultData, activeDays, slotMode, originalFileName);
}

// 綁定彙整模式切換監聽
document.getElementById('mode-select').addEventListener('change', analyzeAndRenderMapping);

// ==========================================
// 📋 直接貼上名單模式邏輯
// ==========================================

let pastedClubs = [];

function switchInputTab(tab) {
    document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));
    
    if (tab === 'file') {
        document.getElementById('tab-btn-file').classList.add('active');
        document.getElementById('tab-file-content').classList.add('active');
    } else {
        document.getElementById('tab-btn-paste').classList.add('active');
        document.getElementById('tab-paste-content').classList.add('active');
    }
}

// 監聽貼上事件
document.getElementById('paste-text-area').addEventListener('paste', (e) => {
    setTimeout(() => {
        const text = document.getElementById('paste-text-area').value;
        guessClubAndDayFromPastedText(text);
    }, 50);
});

function loadPastedClub() {
    const day = document.getElementById('paste-day-select').value;
    const clubNameInput = document.getElementById('paste-club-name').value.trim();
    const pastedText = document.getElementById('paste-text-area').value.trim();
    
    if (!clubNameInput) {
        alert("請先輸入社團名稱！");
        return;
    }
    
    if (!pastedText) {
        alert("請貼上網頁複製的名單資料！");
        return;
    }
    
    const students = parsePastedText(pastedText);
    if (students.length === 0) {
        alert("無法解析名單資料。請確認資料中包含「姓名」等標題列與學生資料列。");
        return;
    }
    
    const newClub = {
        id: Date.now(),
        day: day,
        clubName: clubNameInput,
        students: students
    };
    
    pastedClubs.push(newClub);
    
    document.getElementById('paste-club-name').value = "";
    document.getElementById('paste-text-area').value = "";
    
    renderLoadedClubs();
}

function renderLoadedClubs() {
    const listDiv = document.getElementById('loaded-clubs-list');
    listDiv.innerHTML = "";
    
    let totalStudents = 0;
    
    pastedClubs.forEach((club, index) => {
        totalStudents += club.students.length;
        
        const item = document.createElement('div');
        item.className = "loaded-club-item";
        
        const info = document.createElement('span');
        info.className = "loaded-club-info";
        info.textContent = `📅 ${club.day} ── 🏫 ${club.clubName} (共 ${club.students.length} 人)`;
        
        const delBtn = document.createElement('button');
        delBtn.className = "loaded-club-delete";
        delBtn.innerHTML = "🗑️ 刪除";
        delBtn.onclick = () => {
            pastedClubs.splice(index, 1);
            renderLoadedClubs();
        };
        
        item.appendChild(info);
        item.appendChild(delBtn);
        listDiv.appendChild(item);
    });
    
    const section = document.getElementById('loaded-clubs-section');
    const totalSpan = document.getElementById('total-pasted-students');
    
    if (pastedClubs.length > 0) {
        section.style.display = "block";
        totalSpan.textContent = totalStudents;
    } else {
        section.style.display = "none";
        totalSpan.textContent = "0";
    }
}

async function processPastedData() {
    if (pastedClubs.length === 0) {
        alert("請先載入至少一個社團名單！");
        return;
    }
    
    const slotMode = document.getElementById('paste-slot-mode-select').value;
    const includeFreshmen = document.getElementById('paste-include-freshmen').checked;
    
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
    if (resultData.length === 0) {
        alert("無任何有效學生名單資料可供彙整！");
        return;
    }
    
    resultData.sort((a, b) => {
        const classA = getNumericSortKey(a.class);
        const classB = getNumericSortKey(b.class);
        if (classA !== classB) return classA - classB;
        
        const seatA = getNumericSortKey(a.seat);
        const seatB = getNumericSortKey(b.seat);
        if (seatA !== seatB) return seatA - seatB;
        
        if (a.class !== b.class) return String(a.class).localeCompare(b.class);
        return String(a.seat).localeCompare(b.seat);
    });
    
    await exportWeeklySchedule(resultData, activeDays, slotMode, "直接貼上名單彙整");
    
    // 重設狀態
    pastedClubs = [];
    renderLoadedClubs();
    document.getElementById('paste-text-area').value = "";
}
