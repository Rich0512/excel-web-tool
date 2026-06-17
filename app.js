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

// 正則表達式匹配星期
const WEEKDAY_REGEX = /(星期|週|周)([一二三四五六日1-7])/;
const weekdaysList = ["請選擇", "週一", "週二", "週三", "週四", "週五", "週六", "週日"];

// ==========================================
// 🛠️ 輔助工具函數 (防止極端邊緣錯誤)
// ==========================================

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
    // 僅保留數字與小數點
    const cleanVal = strVal.replace(/[^\d.]/g, '');
    if (cleanVal) {
        const num = parseFloat(cleanVal);
        return isNaN(num) ? Infinity : num;
    }
    return Infinity;
}

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

// 模糊匹配欄位索引
function findColumnByKeywords(headers, keywords) {
    for (let idx = 0; idx < headers.length; idx++) {
        const val = headers[idx];
        if (!val) continue;
        if (keywords.some(kw => val.includes(kw))) {
            return idx;
        }
    }
    return -1;
}

// 讀取檔案內的內置對照工作表
function loadScheduleFromExcel(workbook) {
    let scheduleSheet = null;
    workbook.eachSheet((ws) => {
        const name = ws.name.trim();
        if (name === '社團上課時間對照表' || name === '對照表' || name === '對照') {
            scheduleSheet = ws;
        }
    });
    if (!scheduleSheet) return null;

    const mapping = {};
    scheduleSheet.eachRow((row, rowNum) => {
        if (rowNum === 1) return; // 跳過表頭
        const club = getCellValueAsString(row.getCell(1));
        const day = getCellValueAsString(row.getCell(2));
        if (club && day) {
            const normDay = normalizeWeekdayName(day);
            if (normDay) {
                mapping[club] = normDay;
            }
        }
    });
    return mapping;
}

// 正規化星期文字
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

// 從社團字串自動提取星期
function extractWeekdayFromName(clubName) {
    return normalizeWeekdayName(clubName);
}

// 智慧解析工作表名稱與時段資訊
function parseSheetName(name) {
    const day = normalizeWeekdayName(name);
    let slot = "";
    
    // 偵測 A/B 時段後綴 (例: "週一舞蹈A", "週二羽球B時段", "第一節課桌球")
    const slotMatch = name.match(/([A-Ba-b]|一時段|二時段|第一節|第二節|SlotA|SlotB)/i);
    if (slotMatch) {
        const rawSlot = slotMatch[1].toUpperCase();
        if (rawSlot === "A" || rawSlot === "一時段" || rawSlot === "第一節" || rawSlot === "SLOTA") {
            slot = "A";
        } else if (rawSlot === "B" || rawSlot === "二時段" || rawSlot === "第二節" || rawSlot === "SLOTB") {
            slot = "B";
        }
    }
    
    // 智慧去蕪存菁得到乾淨社團名 (過濾星期、時段字眼)
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

// LocalStorage 管理設定
function loadLocalStorageConfig() {
    const data = localStorage.getItem('yenping_club_config');
    return data ? JSON.parse(data) : {};
}

function saveLocalStorageConfig(config) {
    localStorage.setItem('yenping_club_config', JSON.stringify(config));
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
                if (!day) day = savedConfig[club]; // 優先尋找扁平記錄
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

            // 時段 (多 Sheet 模式下預設改顯示工作表特徵)
            const tdSlot = document.createElement('td');
            const displaySlot = parsed.slot ? `時段 ${parsed.slot}` : "多工作表";
            const badgeClass = parsed.slot ? (parsed.slot === 'A' ? "slot-badge slot-a" : "slot-badge slot-other") : "slot-badge slot-other";
            tdSlot.innerHTML = `<span class="${badgeClass}">${displaySlot}</span>`;
            tr.appendChild(tdSlot);

            // 社團名輸入框 (預設過濾星期後的社團名)
            const tdName = document.createElement('td');
            const input = document.createElement('input');
            input.type = "text";
            input.className = "club-input";
            input.value = parsed.clubName;
            input.dataset.original = name; // 保留原始工作表名稱作為 Key
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
// ⚡ 資料轉換與下載邏輯
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
        
        const origKey = input.dataset.original; // 單一總表下為原始社團名；多 Sheet 下為工作表名
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

    // 自動更新並儲存 LocalStorage 設定 (平鋪存儲方便自動匹配)
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
                        
                        // 學期模式與寒暑假模式判斷
                        let label = "";
                        if (slotMode === 'vacation') {
                            label = slot.name.replace("時段", "").replace("Slot", "").replace("slot", "").replace("社團", "").replace("課程", "").trim();
                        }
                        const entryText = label ? `${displayName}(${label})` : displayName;
                        
                        // 若此星期該學生已有課，進行合併並標示衝突
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
            const sheetName = tr.dataset.detectedSlot ? tr.dataset.original : tr.dataset.sheetName;
            const mapData = finalMapping[sheetName];
            if (!mapData) return;

            const sheet = uploadedWorkbook.getWorksheet(sheetName);
            if (!sheet) return;

            const day = mapData.day;
            const displayName = mapData.editedName;
            const slot = mapData.slot;

            activeDays.add(day);

            // 在每個子 Sheet 智慧定位關鍵欄位列
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
                console.warn(`工作表「${sheetName}」缺少班級、座號或姓名，已跳過。`);
                return;
            }

            sheet.eachRow((row, rowNum) => {
                if (rowNum <= localHeaderRowIndex) return;

                const nameVal = getCellValueAsString(row.getCell(lNameIdx + 1));
                const classVal = getCellValueAsString(row.getCell(lClassIdx + 1));
                const seatVal = getCellValueAsString(row.getCell(lSeatIdx + 1));

                if (!nameVal) return;
                if (!includeFreshmen && !classVal) return;

                // 略過合計行
                if (['合計', '總計', '統計', '人數', '小計'].some(k => nameVal.includes(k) || (classVal && classVal.includes(k)))) {
                    return;
                }

                const studentClass = classVal || "新生";
                const studentSeat = seatVal || "";

                // 以 班級+姓名 組合為唯一 Key
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

                // 組合星期格式
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

        // 轉換為陣列以便排序
        Object.values(studentMap).forEach(s => resultData.push(s));
    }

    if (resultData.length === 0) {
        alert("無任何有效學生資料可供匯出，請確認上傳檔案。");
        return;
    }

    // 嚴格比照 Python 排序演算法：班級數值升冪、座號數值升冪、最後以字串兜底
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

        // 當數值完全一致時，使用字串比對做穩定排序 (zh-hant)
        const strClassA = String(a.class);
        const strClassB = String(b.class);
        if (strClassA !== strClassB) {
            return strClassA.localeCompare(strClassB, 'zh-hant');
        }

        const strSeatA = String(a.seat);
        const strSeatB = String(b.seat);
        return strSeatA.localeCompare(strSeatB, 'zh-hant');
    });

    // 動態裁剪星期直欄
    const targetWeekdays = ['週一', '週二', '週三', '週四', '週五'];
    if (activeDays.has('週六')) targetWeekdays.push('週六');
    if (activeDays.has('週日')) targetWeekdays.push('週日');

    // 建立 Excel 工作簿
    const outWorkbook = new ExcelJS.Workbook();
    const ws = outWorkbook.addWorksheet('學生週課表');

    // 寫入表頭
    const headersOut = ['班級', '座號', '姓名', ...targetWeekdays, '備註'];
    ws.addRow(headersOut);

    // 寫入資料
    resultData.forEach(student => {
        const rowData = [
            student.class,
            student.seat,
            student.name,
            ...targetWeekdays.map(day => student.schedule[day]),
            student.remarks.join('；')
        ];
        ws.addRow(rowData);
    });

    // 💅 延平深藍表頭渲染
    const headerRow = ws.getRow(1);
    headerRow.height = 28;
    
    headerRow.eachCell((cell) => {
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF0F3B6D' } // 經典延平藍
        };
        cell.font = {
            name: 'Microsoft JhengHei',
            size: 11,
            bold: true,
            color: { argb: 'FFFFFFFF' }
        };
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        cell.border = {
            top: { style: 'thin', color: { argb: 'FFD1D5DB' } },
            bottom: { style: 'medium', color: { argb: 'FFFFFFFF' } },
            left: { style: 'thin', color: { argb: 'FFD1D5DB' } },
            right: { style: 'thin', color: { argb: 'FFD1D5DB' } }
        };
    });

    // 資料行樣式與斑馬紋
    ws.eachRow((row, rowNum) => {
        if (rowNum === 1) return;
        
        row.height = 22;
        
        const isEven = rowNum % 2 === 0;
        const bgColor = isEven ? 'FFF8FAFC' : 'FFFFFFFF'; // 交替色

        row.eachCell((cell, colNum) => {
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: bgColor }
            };
            cell.font = {
                name: 'Microsoft JhengHei',
                size: 11
            };
            cell.border = {
                top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
                bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
                left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
                right: { style: 'thin', color: { argb: 'FFE2E8F0' } }
            };

            const colName = headersOut[colNum - 1];
            if (['姓名', '備註'].includes(colName)) {
                cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
            } else {
                cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
            }
        });
    });

    // 🏎️ 智慧自適應欄寬
    ws.columns.forEach(column => {
        let maxLen = 0;
        column.eachCell({ includeEmpty: true }, (cell) => {
            const val = getCellValueAsString(cell);
            let len = 0;
            for (let i = 0; i < val.length; i++) {
                const code = val.charCodeAt(i);
                if (code >= 0x4e00 && code <= 0x9fff) {
                    len += 2; // 中文字元
                } else {
                    len += 1;
                }
            }
            if (len > maxLen) maxLen = len;
        });
        column.width = Math.max(maxLen + 4, 10);
    });

    // 🔍 自動篩選器 (AutoFilter)
    const lastColLetter = getColLetter(headersOut.length);
    ws.autoFilter = `A1:${lastColLetter}${ws.rowCount}`;

    // 5. 檔案下載
    const buffer = await outWorkbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    
    const downloadLink = document.createElement('a');
    const outputSuffix = slotMode === 'semester' ? "_學期社團彙整.xlsx" : "_寒暑假社團彙整.xlsx";
    const outputName = originalFileName.replace(/\.[^/.]+$/, "") + outputSuffix;
    downloadLink.href = URL.createObjectURL(blob);
    downloadLink.download = outputName;
    
    document.body.appendChild(downloadLink);
    downloadLink.click();
    document.body.removeChild(downloadLink);

    // 切換至成功頁面
    switchStep('step-success');
}

// 綁定彙整模式切換監聽
document.getElementById('mode-select').addEventListener('change', analyzeAndRenderMapping);
