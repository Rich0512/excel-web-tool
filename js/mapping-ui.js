/**
 * 🎨 課表對照設定 UI 渲染器 (View/UI Component)
 */

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

    // 綁定動態事件監聽以更新衝突 (若 app.js 已載入對應函數)
    if (typeof updateMappingConflicts === 'function') {
        tbody.querySelectorAll('.day-select, .club-input').forEach(el => {
            el.addEventListener('change', updateMappingConflicts);
            el.addEventListener('input', updateMappingConflicts);
        });
        updateMappingConflicts();
    }
}
