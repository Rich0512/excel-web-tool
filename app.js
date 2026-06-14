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

// 1. 讀取 Excel 檔案
async function handleFileSelect(file) {
    originalFileName = file.name;
    const reader = new FileReader();
    
    reader.onload = async function(e) {
        const arrayBuffer = e.target.result;
        const workbook = new ExcelJS.Workbook();
        
        try {
            await workbook.xlsx.load(arrayBuffer);
            
            // 尋找名為「總表」的工作表 (不區分前後空白)
            let sheet = null;
            workbook.eachSheet((ws) => {
                if (ws.name.trim() === '總表') {
                    sheet = ws;
                }
            });
            
            if (!sheet) {
                alert("Excel 檔案中找不到名為「總表」的工作表！請檢查您的檔案結構。");
                return;
            }
            
            // 掃描前 10 行以智慧偵測表頭列位置
            let foundHeader = false;
            for (let r = 1; r <= 10; r++) {
                const row = sheet.getRow(r);
                const rowValues = [];
                row.eachCell({ includeEmpty: true }, (cell) => {
                    rowValues.push(String(cell.value || '').trim());
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
                // 退回第一行
                const row = sheet.getRow(1);
                detectedHeaders = [];
                row.eachCell({ includeEmpty: true }, (cell) => {
                    detectedHeaders.push(String(cell.value || '').trim());
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

            // 嘗試載入 Excel 內置對照表 (如果有的話)
            const excelSchedule = loadScheduleFromExcel(workbook);

            // 讀取所有學生行
            sheetData = [];
            sheet.eachRow((row, rowNum) => {
                if (rowNum <= headerRowIndex) return; // 跳過表頭及其上方行
                
                const nameVal = row.getCell(colNameIdx + 1).value;
                const classVal = row.getCell(colClassIdx + 1).value;
                
                // 略過空行
                if (nameVal === null || classVal === null) return;
                
                const rowData = {};
                headers.forEach((colName, idx) => {
                    if (!colName) return;
                    const cell = row.getCell(idx + 1);
                    rowData[colName] = cell.value !== null ? String(cell.value).trim() : "";
                });
                sheetData.push(rowData);
            });

            if (sheetData.length === 0) {
                alert("「總表」中無任何有效學生資料，請檢查檔案內容。");
                return;
            }

            // 蒐集當前 Excel 各時段出現的社團
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
                    // 優先度：1. 社團名稱內置星期 2. 檔案內置對照表 3. 本地記憶歷史設定
                    let day = extractWeekdayFromName(club);
                    if (!day && excelSchedule) day = excelSchedule[club];
                    if (!day) day = savedSlotMap[club];
                    
                    predefined[club] = weekdaysList.includes(day) ? day : "請選擇";
                });
                slotsPredefined[slot.name] = predefined;
            });

            // 渲染 HTML 核對表
            renderMappingTable(slotsPredefined);
            
            // 切換至對照設定步驟
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
        const club = row.getCell(1).value;
        const day = row.getCell(2).value;
        if (club && day) {
            const normDay = normalizeWeekdayName(String(day));
            if (normDay) {
                mapping[String(club).trim()] = normDay;
            }
        }
    });
    return mapping;
}

// 正規化星期文字
function normalizeWeekdayName(dayStr) {
    if (!dayStr) return null;
    dayStr = dayStr.trim();
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

// 儲存設定
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

            // 社團名稱輸入框 (可點選修改)
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

    // 綁定鍵盤 Enter 與 Esc 事件
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

// 2. 執行彙整與 Excel 生成下載
async function processAndDownload() {
    const rows = document.getElementById('mapping-table-body').querySelectorAll('tr');
    
    const finalMapping = {}; // 格式: { slotCol: { origClub: [editedClub, selectedDay] } }
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

    // 3. 處理學生資料重組與星期對應
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
                    
                    // 若此時段該星期已有課，進行合併處理並加上備註
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

    // 排序資料：優先班級（升冪），次優先座號（數值升冪）
    resultData.sort((a, b) => {
        // 班級比較
        const classA = String(a.class);
        const classB = String(b.class);
        if (classA !== classB) {
            return classA.localeCompare(classB, 'zh-hant');
        }
        // 座號比較 (提取數字)
        const getNum = (val) => {
            const match = String(val).match(/\d+/);
            return match ? parseInt(match[0], 10) : Infinity;
        };
        return getNum(a.seat) - getNum(b.seat);
    });

    // 動態裁剪星期欄位：若全校無人使用週六或週日，則自動剔除該列，保持排版精美
    const targetWeekdays = ['週一', '週二', '週三', '週四', '週五'];
    if (activeDays.has('週六')) targetWeekdays.push('週六');
    if (activeDays.has('週日')) targetWeekdays.push('週日');

    // 4. 使用 ExcelJS 建立精美的延平藍週課表
    const outWorkbook = new ExcelJS.Workbook();
    const ws = outWorkbook.addWorksheet('學生週課表');

    // 表頭結構
    const headersOut = ['班級', '座號', '姓名', ...targetWeekdays, '備註'];
    ws.addRow(headersOut);

    // 寫入學生資料行
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

    // 💅 套用延平藍美學樣式
    const headerRow = ws.getRow(1);
    headerRow.height = 28;
    
    headerRow.eachCell((cell) => {
        // 經典延平深藍色背景 (#0f3b6d)
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF0F3B6D' }
        };
        // 白色加粗字體微軟正黑體
        cell.font = {
            name: 'Microsoft JhengHei',
            size: 11,
            bold: true,
            color: { argb: 'FFFFFFFF' }
        };
        // 文字置中
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        // 框線
        cell.border = {
            top: { style: 'thin', color: { argb: 'FFD1D5DB' } },
            bottom: { style: 'medium', color: { argb: 'FFFFFFFF' } },
            left: { style: 'thin', color: { argb: 'FFD1D5DB' } },
            right: { style: 'thin', color: { argb: 'FFD1D5DB' } }
        };
    });

    // 資料列樣式與斑馬紋
    ws.eachRow((row, rowNum) => {
        if (rowNum === 1) return; // 跳過表頭
        
        row.height = 22;
        
        // 斑馬紋交替色：白 / 淺灰藍 (#F8FAFC)
        const isEven = rowNum % 2 === 0;
        const bgColor = isEven ? 'FFF8FAFC' : 'FFFFFFFF';

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

            // 對齊設定：前三列與星期置中，姓名與備註靠左
            const colName = headersOut[colNum - 1];
            if (['姓名', '備註'].includes(colName)) {
                cell.alignment = { horizontal: 'left', vertical: 'middle' };
            } else {
                cell.alignment = { horizontal: 'center', vertical: 'middle' };
            }
        });
    });

    // 🏎️ 自動調整欄寬
    ws.columns.forEach(column => {
        let maxLen = 0;
        column.eachCell({ includeEmpty: true }, (cell) => {
            const val = cell.value ? String(cell.value) : '';
            // 智慧字元長度計算（繁體中文字元計為 2，英文數字計為 1）
            let len = 0;
            for (let i = 0; i < val.length; i++) {
                const code = val.charCodeAt(i);
                if (code >= 0x4e00 && code <= 0x9fff) {
                    len += 2;
                } else {
                    len += 1;
                }
            }
            if (len > maxLen) maxLen = len;
        });
        column.width = Math.max(maxLen + 4, 10);
    });

    // 🔍 套用 Excel 自動篩選漏斗 (AutoFilter) 到第一列
    ws.autoFilter = {
        from: { row: 1, column: 1 },
        to: { row: 1, column: headersOut.length }
    };

    // 5. 匯出下載
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
