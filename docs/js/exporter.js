/**
 * 核心彙整與導出 Excel 樣式渲染輔助函數
 */

const GRADE_COLORS = {
    1: 'FFFFD2', // Light Yellow
    2: 'D6EAF8', // Light Blue
    3: 'D4EFDF', // Light Green
    4: 'E8DAEF', // Light Purple
    5: 'F5CBA7', // Light Orange
    6: 'FADBD8', // Light Red
    0: 'EAEDED'  // Light Grey (新生)
};

// 輔助函數：合併並套用同一格式到錄取人數儲存格
function mergeAndStyleCountCell(ws, startRow, endRow, count, grade) {
    const colNum = 5; // Column E
    const color = GRADE_COLORS[grade] || 'FFFFFF';

    if (startRow < endRow) {
        ws.mergeCells(startRow, colNum, endRow, colNum);
    }
    
    const cell = ws.getCell(startRow, colNum);
    cell.value = count;
    cell.font = { name: 'Microsoft JhengHei', size: 11, bold: true };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    
    for (let r = startRow; r <= endRow; r++) {
        const c = ws.getCell(r, colNum);
        c.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF' + color }
        };
        c.border = {
            top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
            bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
            left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
            right: { style: 'thin', color: { argb: 'FFE2E8F0' } }
        };
    }
}

/**
 * 抽籤資訊的內容（標題 → 值）。都是抽籤設定與筆數，不含學生資料。
 * 舊的社團物件沒有某一項時，筆數以名單推算，抽籤時間與抽籤次數寫「未記錄」。
 * 抽籤次數是本次開啟網頁後同名社團第幾次抽籤（app.js 的 clubDrawCounts），大於 1 表示移除後重新抽過籤
 */
function drawInfoRows(club) {
    const students = Array.isArray(club.students) ? club.students : [];
    const selectedCount = Number.isFinite(club.selectedCount) ? club.selectedCount : students.filter(s => s.selected).length;
    const duplicateCount = club.duplicateCount || 0;
    const unrecognizedCount = Number.isFinite(club.unrecognizedCount)
        ? club.unrecognizedCount
        : students.filter(s => s.drawSequence === UNRECOGNIZED_CLASS_LABEL).length;
    // 「報名人數」與畫面上的「報名 N 人」相同，是去除重複後的人數；重複的筆數另列一行
    return [
        ['抽籤時間', formatDrawTime(club.drawnAt, true) || '未記錄'],
        ['抽籤次數', Number.isInteger(club.drawRound) && club.drawRound >= 1
            ? `第 ${club.drawRound} 次（本次開啟網頁後，同名社團的抽籤次數）`
            : '未記錄'],
        ['正取人數', `${selectedCount} 人（名額 ${club.quota} 人）`],
        ['年級限制', `${gradeOptionLabel(club.gradeStart)} 至 ${gradeOptionLabel(club.gradeEnd)}`],
        ['優先順序', PRIORITY_LABELS[club.priority] || String(club.priority || '')],
        ['報名人數', `${students.length} 人（已去除重複）`],
        ['重複略過筆數', `${duplicateCount} 筆`],
        ['班級無法辨識筆數', `${unrecognizedCount} 筆`],
        ['亂數來源', '瀏覽器內建的安全亂數（crypto.getRandomValues）'],
    ];
}

/**
 * 在社團工作表的 startCol 欄（與下一欄）寫入「抽籤資訊」小表：第 1 列是標題（與資料表表頭同樣式、合併兩欄），
 * 之後每列是「項目、內容」。只寫這兩欄，不影響資料表與「錄取人數」欄的合併
 */
function writeDrawInfoBlock(ws, club, startCol) {
    const thinBorder = {
        top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        right: { style: 'thin', color: { argb: 'FFE2E8F0' } }
    };

    ws.mergeCells(1, startCol, 1, startCol + 1);
    const title = ws.getCell(1, startCol);
    title.value = '抽籤資訊';
    title.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F3B6D' } };
    title.font = { name: 'Microsoft JhengHei', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
    title.alignment = { horizontal: 'center', vertical: 'middle' };
    title.border = {
        top: { style: 'thin', color: { argb: 'FFD1D5DB' } },
        bottom: { style: 'medium', color: { argb: 'FFFFFFFF' } },
        left: { style: 'thin', color: { argb: 'FFD1D5DB' } },
        right: { style: 'thin', color: { argb: 'FFD1D5DB' } }
    };

    drawInfoRows(club).forEach(([label, value], i) => {
        const rowNum = i + 2;
        const labelCell = ws.getCell(rowNum, startCol);
        const valueCell = ws.getCell(rowNum, startCol + 1);
        labelCell.value = label;
        valueCell.value = value;
        labelCell.font = { name: 'Microsoft JhengHei', size: 11, bold: true };
        valueCell.font = { name: 'Microsoft JhengHei', size: 11 };
        labelCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
        [labelCell, valueCell].forEach(c => {
            c.border = thinBorder;
            c.alignment = { horizontal: 'left', vertical: 'middle' };
        });
    });
}

// options.showSuccess：下載後是否切到完成畫面（預設是）。完成畫面的「再下載一次」傳 false，
// 處理中使用者清除資料回到第一步時，匯出完成後才不會又跳回完成畫面
async function exportWeeklySchedule(resultData, activeDays, slotMode, outputBaseName, pastedClubsList = null, { showSuccess = true } = {}) {
    const targetWeekdays = ['週一', '週二', '週三', '週四', '週五'];
    // 晨間專屬欄位（依序排在週五之後）
    MORNING_OPTIONS.forEach(m => {
        if (activeDays.has(m)) targetWeekdays.push(m);
    });
    Array.from(activeDays).forEach(d => {
        if (d.startsWith('晨間') && !targetWeekdays.includes(d)) {
            targetWeekdays.push(d);
        }
    });
    if (activeDays.has('週六')) targetWeekdays.push('週六');
    if (activeDays.has('週日')) targetWeekdays.push('週日');

    // ExcelJS 延遲載入，通常在頁面閒置時已經載好；載入失敗時丟出 ExcelJSLoadError，由呼叫端顯示訊息
    const ExcelJSLib = await loadExcelJS();
    const outWorkbook = new ExcelJSLib.Workbook();
    
    // 1. 建立週課表總表（標籤名稱設定為「課程與學生清單」）
    const ws = outWorkbook.addWorksheet('課程與學生清單');

    const headersOut = ['班級', '座號', '姓名', ...targetWeekdays, '備註'];
    ws.addRow(headersOut);

    resultData.forEach(student => {
        const rowData = [
            student.class,
            student.seat,
            student.name,
            ...targetWeekdays.map(day => student.schedule[day]),
            Array.isArray(student.remarks) ? student.remarks.join('；') : (student.remarks || '')
        ];
        ws.addRow(rowData);
    });

    // 總表表頭樣式 (經典延平藍)
    const headerRow = ws.getRow(1);
    headerRow.height = 28;
    headerRow.eachCell((cell) => {
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF0F3B6D' }
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

    // 總表資料列樣式與斑馬紋
    ws.eachRow((row, rowNum) => {
        if (rowNum === 1) return;
        row.height = 22;
        const isEven = rowNum % 2 === 0;
        const bgColor = isEven ? 'FFF8FAFC' : 'FFFFFFFF';

        row.eachCell((cell, colNum) => {
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: bgColor }
            };
            cell.font = { name: 'Microsoft JhengHei', size: 11 };
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

    // 總表自適應欄寬
    ws.columns.forEach(column => {
        let maxLen = 0;
        column.eachCell({ includeEmpty: true }, (cell) => {
            const val = getCellValueAsString(cell);
            let len = 0;
            for (let i = 0; i < val.length; i++) {
                const code = val.charCodeAt(i);
                if (code >= 0x4e00 && code <= 0x9fff) len += 2;
                else len += 1;
            }
            if (len > maxLen) maxLen = len;
        });
        column.width = Math.max(maxLen + 4, 10);
    });

    const lastColLetter = getColLetter(headersOut.length);
    ws.autoFilter = `A1:${lastColLetter}${ws.rowCount}`;

    // 2. 建立個別社團抽籤結果工作表 (若有貼上資料)
    if (pastedClubsList && pastedClubsList.length > 0) {
        const usedSheetNames = new Set(['課程與學生清單']);

        pastedClubsList.forEach(club => {
            // 工作表名稱消毒與唯一性保護（防止特殊字元與重複命名導致 ExcelJS 拋錯）
            let baseWsName = sanitizeSheetName(club.clubName, '社團名冊');
            let finalWsName = baseWsName;
            let counter = 2;
            while (usedSheetNames.has(finalWsName)) {
                const suffix = ` (${counter})`;
                finalWsName = baseWsName.substring(0, 31 - suffix.length) + suffix;
                counter++;
            }
            usedSheetNames.add(finalWsName);

            const wsClub = outWorkbook.addWorksheet(finalWsName);

            // 欄位標題
            const headers = ['班級', '座號', '姓名', '籤序', '錄取人數'];
            wsClub.addRow(headers);

            // 填入學生資料 (包含已抽中及未抽中者)
            const students = club.students;
            students.forEach(s => {
                wsClub.addRow([
                    s.class,
                    s.seat,
                    s.name,
                    s.drawSequence || ""
                ]);
            });

            // 格式化社團表頭
            const hRow = wsClub.getRow(1);
            hRow.height = 28;
            hRow.eachCell((cell) => {
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FF0F3B6D' }
                };
                cell.font = {
                    name: 'Microsoft JhengHei',
                    size: 11,
                    bold: true,
                    color: { argb: 'FFFFFFFF' }
                };
                cell.alignment = { horizontal: 'center', vertical: 'middle' };
                cell.border = {
                    top: { style: 'thin', color: { argb: 'FFD1D5DB' } },
                    bottom: { style: 'medium', color: { argb: 'FFFFFFFF' } },
                    left: { style: 'thin', color: { argb: 'FFD1D5DB' } },
                    right: { style: 'thin', color: { argb: 'FFD1D5DB' } }
                };
            });

            // 格式化社團資料列與框線
            wsClub.eachRow((row, rowNum) => {
                if (rowNum === 1) return;
                row.height = 22;

                row.eachCell((cell, colNum) => {
                    cell.font = { name: 'Microsoft JhengHei', size: 11 };
                    cell.border = {
                        top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
                        bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
                        left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
                        right: { style: 'thin', color: { argb: 'FFE2E8F0' } }
                    };
                    if (colNum === 3) {
                        cell.alignment = { horizontal: 'left', vertical: 'middle' };
                    } else {
                        cell.alignment = { horizontal: 'center', vertical: 'middle' };
                    }
                });
            });

            // 年級區塊著色與錄取人數合併
            // gradeCount 為 0 表示目前沒有進行中的區間；正取的年級一定是 0–6，parseGrade 回傳 null 時也能分組（底色用白色）
            let currentGrade = null;
            let startRow = -1;
            let gradeCount = 0;

            for (let i = 0; i < students.length; i++) {
                const s = students[i];
                const rowNum = i + 2; // 資料列從第 2 行開始

                if (s.selected) {
                    const g = parseGrade(s.class);
                    const color = GRADE_COLORS[g] || 'FFFFFF';

                    // 著色第 D 欄 (籤序)
                    const cellD = wsClub.getCell(rowNum, 4);
                    cellD.fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: 'FF' + color }
                    };

                    if (gradeCount === 0 || g !== currentGrade) {
                        // 當前年級改變時，合併並著色前一個年級區間
                        if (gradeCount > 0) {
                            mergeAndStyleCountCell(wsClub, startRow, rowNum - 1, gradeCount, currentGrade);
                        }
                        currentGrade = g;
                        startRow = rowNum;
                        gradeCount = 1;
                    } else {
                        gradeCount++;
                    }
                } else {
                    // 若當前學生未被選中，則前一段的選中區間結束，進行合併
                    if (gradeCount > 0) {
                        mergeAndStyleCountCell(wsClub, startRow, rowNum - 1, gradeCount, currentGrade);
                        currentGrade = null;
                        startRow = -1;
                        gradeCount = 0;
                    }
                }
            }
            // 處理最後一組錄取學生的合併
            if (gradeCount > 0) {
                mergeAndStyleCountCell(wsClub, startRow, students.length + 1, gradeCount, currentGrade);
            }

            // 抽籤資訊放在資料表右側、中間空一欄；在 eachRow 格式化之後才寫入，不會被套成資料列樣式
            writeDrawInfoBlock(wsClub, club, headers.length + 2);

            // 自適應欄寬
            wsClub.columns.forEach(column => {
                let maxLen = 0;
                column.eachCell({ includeEmpty: true }, (cell) => {
                    const val = getCellValueAsString(cell);
                    let len = 0;
                    for (let i = 0; i < val.length; i++) {
                        const code = val.charCodeAt(i);
                        if (code >= 0x4e00 && code <= 0x9fff) len += 2;
                        else len += 1;
                    }
                    if (len > maxLen) maxLen = len;
                });
                column.width = Math.max(maxLen + 4, 10);
            });
        });
    }

    const buffer = await outWorkbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    
    const outputSuffix = slotMode === 'semester' ? "_學期社團彙整.xlsx" : "_寒暑假社團彙整.xlsx";
    // 只去掉 Excel 副檔名：自訂檔名「115.2學期」的「.2學期」不是副檔名
    const outputName = String(outputBaseName).replace(/\.xlsx?$/i, "") + outputSuffix;

    triggerFileDownload(blob, outputName);

    if (showSuccess) switchStep('step-success');
}
