/**
 * 💅 核心彙整與導出 Excel 樣式渲染輔助函數
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
    const color = GRADE_COLORS[grade] || 'FFFFFFFF';
    
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

async function exportWeeklySchedule(resultData, activeDays, slotMode, outputBaseName, pastedClubsList = null) {
    const targetWeekdays = ['週一', '週二', '週三', '週四', '週五'];
    if (activeDays.has('週六')) targetWeekdays.push('週六');
    if (activeDays.has('週日')) targetWeekdays.push('週日');

    const outWorkbook = new ExcelJS.Workbook();
    
    // 1. 建立週課表總表
    const ws = outWorkbook.addWorksheet('學生週課表');

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

    // 💅 總表表頭樣式 (經典延平藍)
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
        pastedClubsList.forEach(club => {
            // 清理工作表名稱長度及特殊字元
            let wsName = club.clubName.replace(/[:\\/?*\[\]]/g, "_").substring(0, 30);
            const wsClub = outWorkbook.addWorksheet(wsName);

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
            let currentGrade = null;
            let startRow = -1;
            let gradeCount = 0;

            for (let i = 0; i < students.length; i++) {
                const s = students[i];
                const rowNum = i + 2; // 資料列從第 2 行開始

                if (s.selected) {
                    const g = parseGrade(s.class);
                    const color = GRADE_COLORS[g] || 'FFFFFFFF';

                    // 著色第 D 欄 (籤序)
                    const cellD = wsClub.getCell(rowNum, 4);
                    cellD.fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: 'FF' + color }
                    };

                    if (g !== currentGrade) {
                        // 當前年級改變時，合併並著色前一個年級區間
                        if (currentGrade !== null && gradeCount > 0) {
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
                    if (currentGrade !== null && gradeCount > 0) {
                        mergeAndStyleCountCell(wsClub, startRow, rowNum - 1, gradeCount, currentGrade);
                        currentGrade = null;
                        startRow = -1;
                        gradeCount = 0;
                    }
                }
            }
            // 處理最後一組錄取學生的合併
            if (currentGrade !== null && gradeCount > 0) {
                mergeAndStyleCountCell(wsClub, startRow, students.length + 1, gradeCount, currentGrade);
            }

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
    
    const downloadLink = document.createElement('a');
    const outputSuffix = slotMode === 'semester' ? "_學期社團彙整.xlsx" : "_寒暑假社團彙整.xlsx";
    const outputName = outputBaseName.replace(/\.[^/.]+$/, "") + outputSuffix;
    downloadLink.href = URL.createObjectURL(blob);
    downloadLink.download = outputName;
    
    document.body.appendChild(downloadLink);
    downloadLink.click();
    document.body.removeChild(downloadLink);

    switchStep('step-success');
}
