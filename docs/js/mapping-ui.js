/**
 * 課表對照設定 UI 渲染器 (View/UI Component)
 */

// 記憶設定裡是否為有效的星期選項（舊版設定或格式損毀時略過）
function isValidDayOption(day) {
    return typeof day === 'string' && day !== '請選擇' && weekdaysList.includes(day);
}

function isValidMorningOption(day) {
    return MORNING_OPTIONS.includes(day);
}

/**
 * 單一清單模式的記憶設定：依時段欄位分開存（同一社團在 A、B 時段可能不同天）
 * 舊版設定只存「社團 → 星期」，讀取時當作備援
 */
function getSavedSlotMap(savedConfig, slotCol) {
    // savedConfig 由 loadLocalStorageConfig() 轉成沒有原型的物件，「__proto__」等名稱也是一般的鍵
    const map = savedConfig[slotCol];
    return map && typeof map === 'object' ? map : createDict();
}

/**
 * 新增一列對照設定（時段標籤、可修改的社團名稱、星期選單）
 * 社團與工作表名稱來自名冊，一律用 textContent / setAttribute 放入，不經過 innerHTML
 */
function appendMappingRow(tbody, row) {
    const tr = document.createElement('tr');
    tr.className = row.rowClass;
    tr.dataset.key = row.key;
    tr.dataset.original = row.original;
    tr.dataset.slot = row.slotCol || '';
    tr.dataset.detectedSlot = row.detectedSlot || '';

    const tdSlot = document.createElement('td');
    const badge = document.createElement('span');
    badge.className = row.badgeClass;
    badge.textContent = row.badgeText;
    tdSlot.appendChild(badge);
    tr.appendChild(tdSlot);

    const tdName = document.createElement('td');
    const input = document.createElement('input');
    input.type = "text";
    input.className = "club-input";
    input.value = row.displayName;
    // 每列沒有可見的 <label>，用 aria-label 讓螢幕閱讀器唸出是哪個社團
    input.setAttribute('aria-label', `社團名稱：${row.ariaName}`);
    tdName.appendChild(input);
    tr.appendChild(tdName);

    const tdDay = document.createElement('td');
    const select = document.createElement('select');
    select.className = "day-select";
    select.setAttribute('aria-label', `上課星期：${row.ariaName}`);
    const finalDay = weekdaysList.includes(row.day) ? row.day : "請選擇";
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
}

/**
 * 多工作表模式中找不到姓名欄而略過的工作表，顯示在對照表上方，避免社團默默從課表消失
 */
function showSkippedSheets(names) {
    const box = document.getElementById('mapping-skip-alert');
    if (!box) return;
    if (!names || names.length === 0) {
        box.hidden = true;
        box.innerHTML = '';
        return;
    }
    box.innerHTML = WARNING_ICON_SVG +
        `<div><div class="conflict-title">有 ${escapeHtml(String(names.length))} 張工作表找不到「姓名」欄，沒有列入課表：</div>` +
        `<ul class="conflict-list">${names.map(n => `<li>${escapeHtml(n)}</li>`).join('')}</ul>` +
        `<p class="skip-alert-hint">請確認這些工作表的學生名單上方有「姓名」欄位標題，再重新選擇檔案。</p></div>`;
    box.hidden = false;
}

/**
 * 依彙整模式分析名冊並重建對照表
 * 成功回傳 true；名冊缺欄位、找不到社團或工作表時已用 alert 說明，回傳 false（呼叫端決定要退回哪一步）
 */
function analyzeAndRenderMapping() {
    if (!uploadedWorkbook) return false;
    const mode = document.getElementById('mode-select').value;
    const tbody = document.getElementById('mapping-table-body');
    tbody.innerHTML = "";

    sheetData = [];
    sheetStudents = new Map();
    showSkippedSheets([]);

    const savedConfig = loadLocalStorageConfig();
    // 檔案內的上課時間（1152 課程表、各社團工作表的「上課時間」），找不到時是空物件
    const excelSchedule = loadScheduleFromExcel(uploadedWorkbook) || createDict();

    if (mode === 'single') {
        // --- 單一總表/清單分欄模式 ---
        let sheet = null;
        uploadedWorkbook.eachSheet((ws) => {
            const n = ws.name.trim();
            if (n.includes('課程與學生清單')) {
                sheet = ws;
            } else if (!sheet && n.includes('總表')) {
                sheet = ws;
            }
        });

        if (!sheet) {
            alert("這個 Excel 沒有「課程與學生清單」或「總表」工作表，已自動改用「多工作表」模式。請確認下方每個社團的星期。");
            document.getElementById('mode-select').value = 'multi';
            return analyzeAndRenderMapping();
        }

        const header = findHeaderRow(sheet);
        headerRowIndex = header.rowIndex;

        // 欄位名稱去重（重複的欄位加上 _2、_3）
        const seenCols = createDict();
        detectedHeaders = header.headers.map(col => {
            if (!col) return "";
            seenCols[col] = (seenCols[col] || 0) + 1;
            return seenCols[col] > 1 ? `${col}_${seenCols[col]}` : col;
        });
        const headers = detectedHeaders;

        colClassIdx = findColumnByKeywords(headers, CLASS_KEYWORDS);
        colSeatIdx = findColumnByKeywords(headers, SEAT_KEYWORDS, true, false);
        colNameIdx = findColumnByKeywords(headers, NAME_KEYWORDS, false, true);

        if (colClassIdx === -1 || colSeatIdx === -1 || colNameIdx === -1) {
            alert(`工作表「${sheet.name}」缺少必要欄位（班級、座號、學生姓名或姓名），無法整理。請確認選擇的是學校系統匯出的錄取名冊，或把「彙整模式」改成「多工作表」。`);
            return false;
        }

        // 尋找時段/課程欄位 (排除類型、狀態、金額等中繼欄位)
        slotCols = [];
        const slotKeywords = ['時段', 'Slot', 'slot', '社團', '課程名稱', '課程', '項目', '節'];
        // 老師、電話、家長等欄位的值不是社團名稱，當成時段欄位會把它們存進瀏覽器設定
        const excludeKeywords = ['類型', '類別', '狀態', '金額', '繳費', '費用', '學校', '學號', '單號', '日期',
                                 '老師', '教師', '備註', '姓名', '電話', '家長'];
        headers.forEach((col, idx) => {
            if (idx === colClassIdx || idx === colSeatIdx || idx === colNameIdx) return;
            if (excludeKeywords.some(ek => col.includes(ek))) return;
            if (slotKeywords.some(kw => col.includes(kw))) {
                slotCols.push({ name: col, index: idx });
            }
        });

        if (slotCols.length === 0) {
            alert(`工作表「${sheet.name}」找不到「課程名稱」、「社團」或「時段」欄位，無法分出社團。請確認選擇的是學校系統匯出的錄取名冊，或把「彙整模式」改成「多工作表」。`);
            return false;
        }

        // 讀取學生列。新生（沒有班級）也先讀進來，是否納入由轉換時的勾選決定，
        // 這樣切換「包含新生」時不必重建對照表，已修改的名稱與星期才不會被重設
        sheet.eachRow((row, rowNum) => {
            if (rowNum <= headerRowIndex) return;
            const nameVal = getCellValueAsString(row.getCell(colNameIdx + 1));
            const classVal = getCellValueAsString(row.getCell(colClassIdx + 1));
            if (!nameVal || isTotalRow(nameVal, classVal)) return;

            const rowData = createDict(); // 鍵是欄位標題（名冊字串）
            headers.forEach((colName, idx) => {
                if (!colName) return;
                rowData[colName] = getCellValueAsString(row.getCell(idx + 1));
            });
            sheetData.push(rowData);
        });

        // 渲染對照表：每個「時段欄位 × 社團」一列
        slotCols.forEach(slot => {
            const slotCol = slot.name;
            const clubs = new Set();
            sheetData.forEach(row => {
                const val = row[slotCol];
                if (val && val !== "無") clubs.add(val);
            });

            const isASlot = detectSlot(slotCol) === 'A' || slotCol.includes('一');
            const savedSlotMap = getSavedSlotMap(savedConfig, slotCol);

            clubs.forEach(club => {
                const cleanName = cleanClubDisplayName(club);
                const sched = excelSchedule[club] || excelSchedule[cleanName];
                let day;
                if (club.includes('晨間')) {
                    // 晨間社團排在專屬欄位：檔案的上課時間 > 上次的設定 > 依名稱推測
                    day = morningDayFromSchedule(sched) ||
                          [savedSlotMap[club], savedConfig[club]].find(isValidMorningOption) ||
                          morningDayFromName(club);
                } else {
                    // 名稱本身的星期最可靠（「周一桌遊社 (一 10,11)」必為週一），其次是檔案的上課時間，最後是上次的設定
                    day = extractWeekdayFromName(club) ||
                          normalizeDayOption(sched) ||
                          [savedSlotMap[club], savedConfig[club]].find(isValidDayOption);
                }

                appendMappingRow(tbody, {
                    rowClass: "mapping-row-single",
                    key: mappingKey(slotCol, club),
                    original: club,
                    slotCol,
                    badgeClass: isASlot ? "slot-badge slot-a" : "slot-badge slot-other",
                    badgeText: slotCol,
                    displayName: cleanName || club,
                    ariaName: club,
                    day
                });
            });
        });

    } else {
        // --- 多工作表分社團模式 ---
        const ignoreKeywords = ['對照', '說明', 'README', '總表', '統計', '人數', '小計', '1152', '清單', '報表'];
        const skipped = [];
        let candidateCount = 0;

        uploadedWorkbook.eachSheet((ws) => {
            const name = ws.name.trim();
            if (ignoreKeywords.some(kw => name.includes(kw))) return;
            candidateCount++;

            // 學生名單在這裡讀一次，衝堂即時提醒與轉換都用同一份資料
            const students = readClubSheetStudents(ws);
            if (!students) {
                skipped.push(name);
                return;
            }
            sheetStudents.set(name, students);

            const parsed = parseSheetName(name);
            const sched = excelSchedule[parsed.baseName] || excelSchedule[name];
            const saved = savedConfig[name];
            let day;
            if (name.includes('晨間')) {
                day = (isValidMorningOption(saved) && saved) ||
                      morningDayFromSchedule(sched) ||
                      morningDayFromName(parsed.baseName);
            } else {
                // 上次的設定 > 工作表名稱的星期 > 檔案內的上課時間
                day = (isValidDayOption(saved) && saved) ||
                      (parsed.day !== '請選擇' ? parsed.day : null) ||
                      normalizeDayOption(sched);
            }

            appendMappingRow(tbody, {
                rowClass: "mapping-row-multi",
                key: name,
                original: name,
                detectedSlot: parsed.slot,
                badgeClass: parsed.slot === 'A' ? "slot-badge slot-a" : "slot-badge slot-other",
                badgeText: parsed.slot ? `時段 ${parsed.slot}` : "多工作表",
                displayName: parsed.clubName,
                ariaName: `工作表「${name}」`,
                day
            });
        });

        if (candidateCount === 0) {
            alert("這個 Excel 找不到可以彙整的社團名單工作表。請確認選擇的是錄取名冊，或把「彙整模式」改成「單一清單」。");
            return false;
        }
        if (sheetStudents.size === 0) {
            // 每張工作表都找不到姓名欄：對照表會是空的，不進入第 2 步。工作表名稱不是學生資料，可以列出
            const MAX_LISTED = 10;
            const listed = skipped.slice(0, MAX_LISTED).map(n => `- ${n}`).join('\n');
            const more = skipped.length > MAX_LISTED ? `\n…等 ${skipped.length} 張` : '';
            alert(`這個 Excel 的工作表都找不到「姓名」欄，無法整理：\n${listed}${more}\n\n請確認每張工作表的學生名單上方有「姓名」欄位標題，或把「彙整模式」改成「單一清單」，再重新選擇檔案。`);
            return false;
        }
        showSkippedSheets(skipped);
    }

    // 修改名稱或星期時即時更新衝堂提醒
    tbody.querySelectorAll('.day-select, .club-input').forEach(el => {
        el.addEventListener('change', updateMappingConflicts);
    });
    updateMappingConflicts();
    return true;
}
