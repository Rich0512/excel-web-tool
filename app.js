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
            
            // 尋找工作表「總表」
            let sheet = null;
            workbook.eachSheet((ws) => {
                if (ws.name.trim() === '總表') {
                    sheet = ws;
                }
            });
            
            if (!sheet) {
                alert("Excel 檔案中找不到名為「總表」的工作表！請檢查您的檔案名稱。");
                return;
            }
            
            // 掃描前 10 行以智慧偵測表頭列位置
            let foundHeader = false;
            for (let r = 1; r <= 10; r++) {
                const row = sheet.getRow(r);
                const rowValues = [];
                row.eachCell({ includeEmpty: true }, (cell) => {
                    rowValues.push(getCellValueAsString(cell));
                });
                
                // 檢查是否包含關鍵欄位
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
            
            // 去重欄位名防止對照混亂
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

            // 模糊定位關鍵欄位索引
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

            // 嘗試載入 Excel 內置對照工作表
            const excelSchedule = loadScheduleFromExcel(workbook);

            // 讀取所有學生行
            sheetData = [];
            sheet.eachRow((row, rowNum) => {
                if (rowNum <= headerRowIndex) return; // 跳過表頭及其上方行
                
                const nameVal = getCellValueAsString(row.getCell(colNameIdx + 1));
                const classVal = getCellValueAsString(row.getCell(colClassIdx + 1));
                
                // 略過空行與合計行
                if (!nameVal || !classVal) return;
                if (['合計', '總計', '統計', '人數', '小計'].some(k => nameVal.includes(k) || classVal.includes(k))) {
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

            if (sheetData.length === 0) {
                alert("「總表」中無任何有效學生資料，請檢查檔案內容。");
                return;
            }

            // 蒐集各時段出現的社團
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

            // 從 LocalStorage 載入歷史設定
            const savedConfig = loadLocalStorageConfig();

            // 準備預填設定
            const slotsPredefined = {};
            slotCols.forEach(slot => {
                const predefined = {};
                const savedSlotMap = savedConfig[slot.name] || {};
                
                activeClubs[slot.name].forEach(club => {
                    let day = extractWeekdayFromName(club);
                    if (!day && excelSchedule) day = excelSchedule[club];
                    if (!day) day = savedSlotMap[club];
                    
                    predefined[club] = weekdaysList.includes(day) ? day : "請選擇";
                });
                slotsPredefined[slot.name] = predefined;
            });

            // 渲染 HTML 核對表
            renderMappingTable(slotsPredefined);
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
    const match = clubName.match(WEEKDAY_REGEX);
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

// LocalStorage 管理設定
function loadLocalStorageConfig() {
    const data = localStorage.getItem('yenping_club_config');
    return data ? JSON.parse(data) : {};
}

function saveLocalStorageConfig(config) {
    localStorage.setItem('yenping_club_config', JSON.stringify(config));
}

// 渲染對照設定表格
function renderMappingTable(slotsPredefined) {
    const tbody = document.getElementById('mapping-table-body');
    tbody.innerHTML = "";

    Object.entries(slotsPredefined).forEach(([slotCol, predefined]) => {
        const isASlot = slotCol.includes('A') || slotCol.includes('一');
        const badgeClass = isASlot ? "slot-badge slot-a" : "slot-badge slot-other";

        Object.entries(predefined).forEach(([club, day]) => {
            const tr = document.createElement('tr');
            
            // 時段
            const tdSlot = document.createElement('td');
            tdSlot.innerHTML = `<span class="${badgeClass}">${slotCol}</span>`;
            tr.appendChild(tdSlot);

            // 社團名稱輸入框
            const tdName = document.createElement('td');
            const input = document.createElement('input');
            input.type = "text";
            input.className = "club-input";
            input.value = club;
            input.dataset.slot = slotCol;
            input.dataset.original = club;
            tdName.appendChild(input);
            tr.appendChild(tdName);

            // 星期下拉選單
            const tdDay = document.createElement('td');
            const select = document.createElement('select');
            select.className = "day-select";
            
            weekdaysList.forEach(w => {
                const opt = document.createElement('option');
                opt.value = w;
                opt.textContent = w;
                if (w === day) opt.selected = true;
                select.appendChild(opt);
            });
            
            tdDay.appendChild(select);
            tr.appendChild(tdDay);

            tbody.appendChild(tr);
        });
    });

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
    const rows = document.getElementById('mapping-table-body').querySelectorAll('tr');
    
    const finalMapping = {};
    const unselectedClubs = [];

    rows.forEach(tr => {
        const input = tr.querySelector('.club-input');
        const select = tr.querySelector('.day-select');
        
        const slotCol = input.dataset.slot;
        const origClub = input.dataset.original;
        const editedClub = input.value.trim();
        const selectedDay = select.value;

        if (!editedClub) {
            alert("社團名稱不能為空！");
            return;
        }

        if (selectedDay === "請選擇") {
            unselectedClubs.push(`${slotCol} 的「${origClub}」`);
        }

        if (!finalMapping[slotCol]) finalMapping[slotCol] = {};
        finalMapping[slotCol][origClub] = {
            editedName: editedClub,
            day: selectedDay
        };
    });

    if (unselectedClubs.length > 0) {
        alert(`請為以下社團選擇上課星期後再行轉換：\n\n${unselectedClubs.join('\n')}`);
        return;
    }

    // 自動更新並儲存 LocalStorage 設定
    const savedConfig = loadLocalStorageConfig();
    Object.entries(finalMapping).forEach(([slotCol, mapping]) => {
        if (!savedConfig[slotCol]) savedConfig[slotCol] = {};
        Object.entries(mapping).forEach(([origClub, data]) => {
            savedConfig[slotCol][origClub] = data.day;
        });
    });
    saveLocalStorageConfig(savedConfig);

    // 處理學生資料重組與星期對應
    const resultData = [];
    const activeDays = new Set();

    sheetData.forEach(row => {
        const student = {
            class: row[detectedHeaders[colClassIdx]],
            seat: row[detectedHeaders[colSeatIdx]],
            name: row[detectedHeaders[colNameIdx]],
            schedule: { '週一': '', '週二': '', '週三': '', '週四': '', '週五': '', '週六': '', '週日': '' },
            remarks: []
        };

        slotCols.forEach(slot => {
            const origClub = row[slot.name];
            if (origClub && origClub !== "無") {
                const mapData = finalMapping[slot.name][origClub];
                if (mapData) {
                    const day = mapData.day;
                    const displayName = mapData.editedName;
                    
                    activeDays.add(day);
                    
                    // 取得時段代碼簡稱 (例 "A時段" -> "A")
                    const label = slot.name.replace("時段", "").replace("Slot", "").replace("slot", "").replace("社團", "").replace("課程", "").trim();
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
    const outputName = originalFileName.replace(/\.[^/.]+$/, "") + "_週一至週五彙整.xlsx";
    downloadLink.href = URL.createObjectURL(blob);
    downloadLink.download = outputName;
    
    document.body.appendChild(downloadLink);
    downloadLink.click();
    document.body.removeChild(downloadLink);

    // 切換至成功頁面
    switchStep('step-success');
}
