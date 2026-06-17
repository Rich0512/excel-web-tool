// 模糊匹配欄位索引
function findColumnByKeywords(headers, keywords) {
    for (let idx = 0; idx < headers.length; idx++) {
        const val = headers[idx];
        if (!val) continue;
        if (keywords.some(kw => val.includes(kw))) {
            if (val.includes('學') && !val.includes('座')) {
                continue;
            }
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

// 智慧解析直接貼上之文字/表格 (Tab 分隔)
function parsePastedText(text) {
    const lines = text.split(/\r?\n/).map(l => l.trim()).filter(l => l.length > 0);
    if (lines.length === 0) return [];
    
    // 預設以 Tab 分割（網頁表格複製之標準格式）
    let rows = lines.map(line => line.split('\t').map(c => c.trim()));
    
    // 若無 Tab，嘗試多空格或逗號分割
    if (rows[0].length <= 1) {
        rows = lines.map(line => line.split(/\s{2,}|\t/).map(c => c.trim()));
    }
    if (rows[0].length <= 1) {
        rows = lines.map(line => line.split(/[,\s]+/).map(c => c.trim()));
    }
    
    let headerIdx = -1;
    let colClass = -1, colSeat = -1, colName = -1;
    
    // 掃描前 5 列尋找表頭
    for (let i = 0; i < Math.min(5, rows.length); i++) {
        const row = rows[i];
        const cIdx = row.findIndex(c => c.includes('班'));
        const sIdx = row.findIndex(c => (c.includes('座') || c.includes('號')) && !c.includes('學'));
        const nIdx = row.findIndex(c => c.includes('姓') || c.includes('名'));
        
        if (nIdx !== -1 && (cIdx !== -1 || sIdx !== -1)) {
            headerIdx = i;
            colClass = cIdx;
            colSeat = sIdx;
            colName = nIdx;
            break;
        }
    }
    
    // 若無表頭，預設第 0 欄為班級、第 1 欄為座號、第 2 欄為姓名
    if (headerIdx === -1) {
        colClass = 0;
        colSeat = 1;
        colName = 2;
    }
    
    const students = [];
    const startRow = headerIdx + 1;
    
    for (let i = startRow; i < rows.length; i++) {
        const row = rows[i];
        const nameVal = row[colName];
        if (!nameVal) continue;
        
        if (['合計', '總計', '統計', '人數', '小計'].some(k => nameVal.includes(k))) {
            continue;
        }
        
        const classVal = (colClass !== -1 && colClass < row.length) ? row[colClass] : "";
        const seatVal = (colSeat !== -1 && colSeat < row.length) ? row[colSeat] : "";
        
        students.push({
            class: classVal,
            seat: seatVal,
            name: nameVal
        });
    }
    
    return students;
}

// 監聽貼上事件，進行智慧預估
function guessClubAndDayFromPastedText(text) {
    if (!text) return;
    const lines = text.split(/\r?\n/).map(l => l.trim()).filter(l => l.length > 0);
    if (lines.length === 0) return;
    
    // 1. 嘗試預估星期
    for (let i = 0; i < Math.min(3, lines.length); i++) {
        const line = lines[i];
        const dayMatch = WEEKDAY_REGEX.exec(line);
        if (dayMatch) {
            const day_char = dayMatch[2];
            const dayMap = {
                '一': '週一', '1': '週一',
                '二': '週二', '2': '週二',
                '三': '週三', '3': '週三',
                '四': '週四', '4': '週四',
                '五': '週五', '5': '週五',
                '六': '週六', '6': '週六',
                '日': '週日', '7': '週日'
            };
            const matchedDay = dayMap[day_char];
            if (matchedDay) {
                document.getElementById('paste-day-select').value = matchedDay;
            }
        }
        
        // 2. 嘗試預估社團名稱
        let cleaned = line.replace(/延平|國小|學年度|第二學期|第一學期|學期|課後|社團|A班|B班|一時段|二時段|錄取|名冊|名單|學生|清單|班級|座號|姓名/g, '');
        cleaned = cleaned.replace(/[^a-zA-Z0-9\u4e00-\u9fff]/g, '').trim();
        if (cleaned.length >= 2 && cleaned.length <= 15) {
            if (!cleaned.includes("班") && !cleaned.includes("座") && !cleaned.includes("名")) {
                document.getElementById('paste-club-name').value = cleaned;
                break;
            }
        }
    }
}
