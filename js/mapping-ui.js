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
        // --- 單一總表/清單分欄模式 ---
        let sheet = null;
        uploadedWorkbook.eachSheet((ws) => {
            const n = ws.name.trim();
            if (n === '課程與學生清單' || n.includes('課程與學生清單')) {
                sheet = ws;
            } else if (!sheet && (n === '總表' || n.includes('總表'))) {
                sheet = ws;
            }
        });

        if (!sheet) {
            alert("在 Excel 中找不到「課程與學生清單」或「總表」工作表！請改用「多工作表模式」。");
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
        colSeatIdx = findColumnByKeywords(headers, ['座號', '座', 'Seat', 'seat'], true, false);
        colNameIdx = findColumnByKeywords(headers, ['學生姓名', '姓名', 'Name', 'name', '名'], false, true);

        if (colClassIdx === -1 || colSeatIdx === -1 || colNameIdx === -1) {
            alert(`工作表「${sheet.name}」中缺少必要欄位（班級、座號、學生姓名/姓名），無法進行整理。`);
            return;
        }

        // 尋找時段/課程欄位 (排除類型、狀態、金額等中繼欄位)
        slotCols = [];
        const slotKeywords = ['時段', 'Slot', 'slot', '社團', '課程名稱', '課程', '項目', '節'];
        const excludeKeywords = ['類型', '類別', '狀態', '金額', '繳費', '費用', '學校', '學號', '單號', '日期'];
        headers.forEach((col, idx) => {
            if (idx === colClassIdx || idx === colSeatIdx || idx === colNameIdx) return;
            if (excludeKeywords.some(ek => col.includes(ek))) return;
            if (slotKeywords.some(kw => col.includes(kw))) {
                slotCols.push({ name: col, index: idx });
            }
        });

        if (slotCols.length === 0) {
            alert(`工作表「${sheet.name}」中找不到任何包含「課程名稱」、「社團」或「時段」的欄位。`);
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
                input.value = cleanClubDisplayName(club) || club;
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
        const ignoreKeywords = ['對照', '說明', 'README', '總表', '統計', '人數', '小計', '1152', '清單', '報表'];

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

    // 綁定動態事件監聽以更新衝突 (若 app.js 已載入對應函數)
    if (typeof updateMappingConflicts === 'function') {
        tbody.querySelectorAll('.day-select, .club-input').forEach(el => {
            el.addEventListener('change', updateMappingConflicts);
        });
        updateMappingConflicts();
    }
}
