/**
 * 💅 核心彙整與導出 Excel 樣式渲染輔助函數
 */
async function exportWeeklySchedule(resultData, activeDays, slotMode, outputBaseName) {
    const targetWeekdays = ['週一', '週二', '週三', '週四', '週五'];
    if (activeDays.has('週六')) targetWeekdays.push('週六');
    if (activeDays.has('週日')) targetWeekdays.push('週日');

    const outWorkbook = new ExcelJS.Workbook();
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

    // 💅 表頭樣式 (經典延平藍)
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

    // 資料列樣式與斑馬紋
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

    // 智慧自適應欄寬
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
