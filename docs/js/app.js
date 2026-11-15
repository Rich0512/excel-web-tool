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

// 點虛線框任何地方都開啟檔案選擇（滑鼠的便利操作）；
// 鍵盤與螢幕閱讀器使用框內真正的「選擇檔案」按鈕，按鈕的點擊會冒泡到這裡。
// 點到框內的範本連結時不開啟，讓連結照常下載。
// fileInput 也在框內，它被 click() 時的事件會冒泡回來，這裡略過以免重入。
dropZone.addEventListener('click', (e) => {
    if (e.target === fileInput || e.target.closest('a')) return;
    fileInput.click();
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
    if (!file) return;

    // 🛡️ 檔頭簽章檢查：精準攔截舊版 .xls (BIFF8) 或偽裝格式
    if (typeof checkExcelFileFormat === 'function') {
        const formatCheck = await checkExcelFileFormat(file);
        if (!formatCheck.valid) {
            alert(formatCheck.reason || "無效的 Excel 檔案格式。");
            if (fileInput) fileInput.value = "";
            return;
        }
    }

    originalFileName = file.name;
    const reader = new FileReader();
    
    reader.onload = async function(e) {
        const arrayBuffer = e.target.result;
        const workbook = new ExcelJS.Workbook();
        
        try {
            await workbook.xlsx.load(arrayBuffer);
            uploadedWorkbook = workbook;
            
            // 偵測是否存在「課程與學生清單」或「總表」工作表
            let hasSingleSheet = false;
            workbook.eachSheet((ws) => {
                const n = ws.name.trim();
                if (n === '課程與學生清單' || n.includes('課程與學生清單') || n === '總表' || n.includes('總表')) {
                    hasSingleSheet = true;
                }
            });
            
            // 依有無清單/總表自動切換預設模式
            const modeSelect = document.getElementById('mode-select');
            if (modeSelect) modeSelect.value = hasSingleSheet ? 'single' : 'multi';
            
            // 分析並渲染對照設定
            analyzeAndRenderMapping();
            switchStep('step-mapping');

        } catch (err) {
            console.error("載入 Excel 失敗:", err);
            alert("載入 Excel 活頁簿失敗，可能檔案損毀或受密碼保護：\n" + (err.message || err));
            if (fileInput) fileInput.value = "";
        }
    };

    reader.onerror = function() {
        alert("讀取本機檔案時發生錯誤，請確認檔案未被其他應用程式獨佔鎖定。");
        if (fileInput) fileInput.value = "";
    };
    
    reader.readAsArrayBuffer(file);
}

// ==========================================
// ⚙️ 步驟切換
// ==========================================

// 切換步驟顯示
function switchStep(stepId) {
    document.querySelectorAll('.step-panel').forEach(panel => {
        panel.classList.remove('active');
    });
    document.getElementById(stepId).classList.add('active');

    // 同步頁首步驟列 (目前步驟以 aria-current 標示，樣式由 CSS 處理)
    document.querySelectorAll('.stepper-item').forEach(item => {
        if (item.dataset.step === stepId) {
            item.setAttribute('aria-current', 'step');
        } else {
            item.removeAttribute('aria-current');
        }
    });
}

function resetToUpload() {
    fileInput.value = "";
    uploadedWorkbook = null;
    sheetData = [];
    detectedHeaders = [];
    colClassIdx = -1;
    colSeatIdx = -1;
    colNameIdx = -1;
    slotCols = [];
    document.getElementById('excel-filename').value = "";
    
    // 清空衝突警告
    const alertDiv = document.getElementById('mapping-conflict-alert');
    if (alertDiv) {
        alertDiv.style.display = 'none';
        alertDiv.innerHTML = '';
    }
    
    switchStep('step-upload');
}

// 🛡️ 一鍵安全清除所有暫存個資
function manualHardResetAll() {
    if (confirm("確定要安全清除本網頁載入的所有 Excel 檔案、對照表、貼上之社團暫存資料與瀏覽器記憶的社團星期設定嗎？這將會清除您目前在畫面上的所有操作與預覽。")) {
        // 清除檔案模式狀態
        fileInput.value = "";
        uploadedWorkbook = null;
        sheetData = [];
        detectedHeaders = [];
        colClassIdx = -1;
        colSeatIdx = -1;
        colNameIdx = -1;
        slotCols = [];
        document.getElementById('excel-filename').value = "";
        
        // 清除貼上模式狀態
        pastedClubs = [];
        document.getElementById('paste-text-area').value = "";
        document.getElementById('paste-club-name').value = "";
        document.getElementById('paste-excel-filename').value = "";

        // 清除瀏覽器記憶的社團對照設定
        clearLocalStorageConfig();

        // 重設 UI
        const mappingTableBody = document.getElementById('mapping-table-body');
        if (mappingTableBody) mappingTableBody.innerHTML = "";
        
        const pastePreview = document.getElementById('paste-parse-preview');
        if (pastePreview) pastePreview.style.display = 'none';
        
        const mappingConflictAlert = document.getElementById('mapping-conflict-alert');
        if (mappingConflictAlert) {
            mappingConflictAlert.style.display = 'none';
            mappingConflictAlert.innerHTML = '';
        }
        
        const pasteConflictAlert = document.getElementById('paste-conflict-alert');
        if (pasteConflictAlert) {
            pasteConflictAlert.style.display = 'none';
            pasteConflictAlert.innerHTML = '';
        }
        
        renderLoadedClubs();
        switchStep('step-upload');
        alert("這次讀取的名冊與設定都已從記憶體清除，可以重新開始。");
    }
}

// ==========================================
// ⚡ 資料轉換與下載邏輯 (檔案上傳模式)
// ==========================================

async function processAndDownload() {
    const actionBtn = document.querySelector('#step-mapping .btn-primary');
    const origBtnHtml = actionBtn ? actionBtn.innerHTML : "";
    
    try {
        if (actionBtn) {
            actionBtn.disabled = true;
            actionBtn.textContent = '正在轉換並下載課表…';
        }

        const mode = document.getElementById('mode-select').value;
        const includeFreshmen = document.getElementById('include-freshmen').checked;
        const slotMode = document.getElementById('slot-mode-select').value;

        const rows = document.getElementById('mapping-table-body').querySelectorAll('tr');
        
        const finalMapping = {};
        const unselectedClubs = [];

        for (const tr of rows) {
            const input = tr.querySelector('.club-input');
            const select = tr.querySelector('.day-select');
            
            const origKey = input.dataset.original;
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
        }

        if (unselectedClubs.length > 0) {
            alert(`請為以下社團選擇上課星期後再行轉換：\n\n${unselectedClubs.join('\n')}`);
            return;
        }

        // 自動更新並儲存 LocalStorage 設定
        const savedConfig = loadLocalStorageConfig();
        Object.entries(finalMapping).forEach(([origKey, data]) => {
            savedConfig[origKey] = data.day;
        });
        saveLocalStorageConfig(savedConfig);

        // 呼叫 Core 彙整處理器
        const { resultData, activeDays } = processExcelData(
            mode,
            includeFreshmen,
            slotMode,
            finalMapping,
            sheetData,
            detectedHeaders,
            colClassIdx,
            colSeatIdx,
            colNameIdx,
            slotCols
        );

        if (resultData.length === 0) {
            alert("無任何有效學生資料可供匯出，請確認上傳檔案。");
            return;
        }

        // 排序
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

            const strClassA = String(a.class);
            const strClassB = String(b.class);
            if (strClassA !== strClassB) {
                return strClassA.localeCompare(strClassB, 'zh-hant');
            }

            const strSeatA = String(a.seat);
            const strSeatB = String(b.seat);
            return strSeatA.localeCompare(strSeatB, 'zh-hant');
        });

        const customFilename = document.getElementById('excel-filename').value.trim();
        const finalFilename = customFilename || originalFileName;
        await exportWeeklySchedule(resultData, activeDays, slotMode, finalFilename);

    } catch (err) {
        console.error("處理或匯出失敗:", err);
        alert("匯出課表時發生錯誤：\n" + (err.message || err));
    } finally {
        if (actionBtn) {
            actionBtn.disabled = false;
            actionBtn.innerHTML = origBtnHtml;
        }
    }
}

// 綁定彙整模式切換監聽與衝突偵測監聽
document.getElementById('mode-select').addEventListener('change', analyzeAndRenderMapping);
document.getElementById('include-freshmen').addEventListener('change', analyzeAndRenderMapping);
document.getElementById('slot-mode-select').addEventListener('change', updateMappingConflicts);

// ==========================================
// 📋 直接貼上名單模式邏輯
// ==========================================

let pastedClubs = [];

function switchInputTab(tab) {
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.classList.remove('active');
        btn.setAttribute('aria-selected', 'false');
    });
    document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));

    const activeBtn = document.getElementById(tab === 'file' ? 'tab-btn-file' : 'tab-btn-paste');
    activeBtn.classList.add('active');
    activeBtn.setAttribute('aria-selected', 'true');

    if (tab === 'file') {
        document.getElementById('tab-file-content').classList.add('active');
    } else {
        document.getElementById('tab-paste-content').classList.add('active');
        // 自動聚焦於貼上文字區
        setTimeout(() => {
            const pasteArea = document.getElementById('paste-text-area');
            if (pasteArea) pasteArea.focus();
        }, 50);
    }
}

// 監聽輸入與貼上事件，進行即時解析預覽與特徵猜測
const pasteArea = document.getElementById('paste-text-area');
const parsePreview = document.getElementById('paste-parse-preview');

function updatePastePreview() {
    const text = pasteArea.value;
    if (text.length > 200000) {
        parsePreview.innerHTML = `<strong class="text-danger">無法讀取：</strong>貼上的內容超過 200,000 字元。請一次只貼上一個社團的名單。`;
        parsePreview.classList.add('is-warning');
        parsePreview.style.display = 'block';
        return;
    }
    if (!text.trim()) {
        parsePreview.style.display = 'none';
        return;
    }
    
    // 進行智慧欄位預估 (社團名稱與星期)
    guessClubAndDayFromPastedText(text);

    // 嘗試解析名單進行即時預覽
    const parsed = parsePastedText(text);
    if (parsed && parsed.length > 0) {
        const hasClass = parsed.filter(s => s.class && s.class !== "新生").length;
        const hasSeat = parsed.filter(s => s.seat).length;
        const hasName = parsed.filter(s => s.name).length;
        
        parsePreview.innerHTML = `讀到 <strong>${parsed.length}</strong> 位學生` +
                                 `<span class="parse-preview-detail">其中 ${hasClass} 位有班級、${hasSeat} 位有座號、${hasName} 位有姓名</span>`;
        parsePreview.classList.remove('is-warning');
        parsePreview.style.display = 'block';
    } else {
        parsePreview.innerHTML = `<strong>讀不到學生資料。</strong>請確認貼上的內容包含班級、座號、姓名這幾欄。`;
        parsePreview.classList.add('is-warning');
        parsePreview.style.display = 'block';
    }
}

// 同時監聽 input 與 paste 事件，確保任何編輯都能觸發預覽
pasteArea.addEventListener('input', updatePastePreview);
pasteArea.addEventListener('paste', () => {
    setTimeout(updatePastePreview, 50);
});

// 監聽貼上模式的設定變更以即時更新衝突狀態
document.getElementById('paste-slot-mode-select').addEventListener('change', renderLoadedClubs);
document.getElementById('paste-include-freshmen').addEventListener('change', renderLoadedClubs);

// 清空所有已載入的社團名單
function clearAllPastedClubs() {
    if (confirm("確定要清空所有已載入的社團名單嗎？這將會清除您之前載入的所有社團資料。")) {
        pastedClubs = [];
        renderLoadedClubs();
        pasteArea.value = "";
        parsePreview.style.display = 'none';
    }
}

function loadPastedClub() {
    const day = document.getElementById('paste-day-select').value;
    const clubNameInput = document.getElementById('paste-club-name').value.trim();
    const pastedText = pasteArea.value.trim();
    
    const quota = parseInt(document.getElementById('paste-quota').value, 10) || 12;
    const gradeStart = document.getElementById('paste-grade-start').value;
    const gradeEnd = document.getElementById('paste-grade-end').value;
    const priority = document.getElementById('paste-priority').value;

    const minG = gradeStart === "新生" ? 0 : parseInt(gradeStart, 10);
    const maxG = gradeEnd === "新生" ? 0 : parseInt(gradeEnd, 10);
    if (minG > maxG) {
        alert(`年級限制錯誤：起始年級（${gradeStart}）不能大於結束年級（${gradeEnd}）！`);
        return;
    }

    if (!clubNameInput) {
        alert("請先輸入社團名稱！");
        return;
    }
    
    if (pastedText.length > 200000) {
        alert("貼上的內容超過 200,000 字元，無法處理。請一次只貼上一個社團的名單。");
        return;
    }
    
    if (!pastedText) {
        alert("請貼上網頁複製的名單資料！");
        return;
    }
    
    let rawStudents = parsePastedText(pastedText);
    if (rawStudents.length === 0) {
        alert("無法解析名單資料。請確認資料中包含「姓名」等標題列與學生資料列。");
        return;
    }
    
    // 執行抽籤演算法
    const finalStudentsList = runClubLottery(rawStudents, quota, gradeStart, gradeEnd, priority);
    const selectedCount = finalStudentsList.filter(s => s.selected).length;

    const newClub = {
        id: Date.now(),
        day: day,
        clubName: clubNameInput,
        quota: quota,
        gradeStart: gradeStart,
        gradeEnd: gradeEnd,
        priority: priority,
        students: finalStudentsList,
        selectedCount: selectedCount
    };
    
    pastedClubs.push(newClub);
    
    document.getElementById('paste-club-name').value = "";
    pasteArea.value = "";
    parsePreview.style.display = 'none'; // 隱藏解析預覽
    
    renderLoadedClubs();
    
    // 自動聚焦回社團名稱輸入框，方便打字下一個社團
    setTimeout(() => {
        const clubNameInput = document.getElementById('paste-club-name');
        if (clubNameInput) clubNameInput.focus();
    }, 50);
}

function renderLoadedClubs() {
    const listDiv = document.getElementById('loaded-clubs-list');
    listDiv.innerHTML = "";
    
    let totalSelectedStudents = 0;
    
    pastedClubs.forEach((club, index) => {
        totalSelectedStudents += club.selectedCount;
        
        const item = document.createElement('div');
        item.className = "loaded-club-item";
        
        // 包裝資訊區以容納可折疊的學生清單
        const infoContainer = document.createElement('div');
        infoContainer.className = "loaded-club-body";

        // 🛡️ 社團名稱與學生姓名來自使用者貼上的內容，插入 HTML 前一律跳脫
        const infoHeader = document.createElement('div');
        infoHeader.className = "loaded-club-info";
        infoHeader.setAttribute('role', 'button');
        infoHeader.setAttribute('tabindex', '0');
        infoHeader.setAttribute('aria-expanded', 'false');
        infoHeader.innerHTML = `<span class="loaded-club-name">${escapeHtml(club.clubName)}</span>` +
                               `<span class="loaded-club-meta">${escapeHtml(club.day)}，正取 ${club.selectedCount}／${club.quota} 人，報名 ${club.students.length} 人，限 ${escapeHtml(club.gradeStart)}–${escapeHtml(club.gradeEnd)} 年級</span>` +
                               `<span class="loaded-club-badge">查看名單</span>`;

        const infoDetails = document.createElement('div');
        infoDetails.className = "loaded-club-details";
        infoDetails.hidden = true;

        const selectedList = club.students.filter(s => s.selected).map(s => escapeHtml(`${s.class || '新生'}${s.seat ? `座號${s.seat}` : ''} ${s.name}`)).join('、');
        const backupList = club.students.filter(s => !s.selected).map(s => escapeHtml(`${s.name}(${s.drawSequence})`)).join('、');

        infoDetails.innerHTML = `<div class="roster-selected"><strong>正取：</strong>${selectedList || '無'}</div>` +
                                 `<div class="roster-backup"><strong>備取或不符資格：</strong>${backupList || '無'}</div>`;

        const toggleDetails = () => {
            const willShow = infoDetails.hidden;
            infoDetails.hidden = !willShow;
            infoHeader.setAttribute('aria-expanded', String(willShow));
            infoHeader.querySelector('.loaded-club-badge').textContent = willShow ? "收合名單" : "查看名單";
        };
        infoHeader.addEventListener('click', toggleDetails);
        infoHeader.addEventListener('keydown', (e) => {
            if (e.key === 'Enter' || e.key === ' ') {
                e.preventDefault();
                toggleDetails();
            }
        });
        
        infoContainer.appendChild(infoHeader);
        infoContainer.appendChild(infoDetails);
        
        const delBtn = document.createElement('button');
        delBtn.type = "button";
        delBtn.className = "loaded-club-delete";
        delBtn.textContent = "移除";
        delBtn.setAttribute('aria-label', `移除 ${club.clubName}`);
        delBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            pastedClubs.splice(index, 1);
            renderLoadedClubs();
        });
        
        item.appendChild(infoContainer);
        item.appendChild(delBtn);
        listDiv.appendChild(item);
    });
    
    const section = document.getElementById('loaded-clubs-section');
    const totalSpan = document.getElementById('total-pasted-students');
    
    if (pastedClubs.length > 0) {
        section.style.display = "block";
        totalSpan.textContent = totalSelectedStudents;
        
        // 即時衝突檢查
        const slotMode = document.getElementById('paste-slot-mode-select').value;
        const includeFreshmen = document.getElementById('paste-include-freshmen').checked;
        const { resultData } = processPastedClubsData(pastedClubs, slotMode, includeFreshmen);
        showConflictAlert(resultData, "paste-conflict-alert");
    } else {
        section.style.display = "none";
        totalSpan.textContent = "0";
        showConflictAlert([], "paste-conflict-alert");
    }

    // 動態更新 Tab 按鈕上的社團計數器
    const pasteCount = document.getElementById('paste-club-count');
    if (pasteCount) {
        pasteCount.textContent = String(pastedClubs.length);
        pasteCount.hidden = pastedClubs.length === 0;
        pasteCount.setAttribute('aria-label', `已加入 ${pastedClubs.length} 個社團`);
    }
}

async function processPastedData() {
    const actionBtn = document.querySelector('#loaded-clubs-section .btn-primary');
    const origBtnHtml = actionBtn ? actionBtn.innerHTML : "";

    try {
        if (pastedClubs.length === 0) {
            alert("請先載入至少一個社團名單！");
            return;
        }
        
        if (actionBtn) {
            actionBtn.disabled = true;
            actionBtn.textContent = '正在抽籤並下載課表…';
        }

        const slotMode = document.getElementById('paste-slot-mode-select').value;
        const includeFreshmen = document.getElementById('paste-include-freshmen').checked;
        
        // 呼叫 Core 彙整處理器
        const { resultData, activeDays } = processPastedClubsData(pastedClubs, slotMode, includeFreshmen);
        
        if (resultData.length === 0) {
            alert("無任何有效學生名單資料可供彙整！");
            return;
        }
        
        resultData.sort((a, b) => {
            const classA = getNumericSortKey(a.class);
            const classB = getNumericSortKey(b.class);
            if (classA !== classB) return classA - classB;
            
            const seatA = getNumericSortKey(a.seat);
            const seatB = getNumericSortKey(b.seat);
            if (seatA !== seatB) return seatA - seatB;
            
            if (a.class !== b.class) return String(a.class).localeCompare(b.class, 'zh-hant');
            return String(a.seat).localeCompare(b.seat, 'zh-hant');
        });
        
        const customFilename = document.getElementById('paste-excel-filename').value.trim();
        const finalFilename = customFilename || "直接貼上名單彙整";
        await exportWeeklySchedule(resultData, activeDays, slotMode, finalFilename, pastedClubs);
        
        // 重設狀態
        pastedClubs = [];
        renderLoadedClubs();
        document.getElementById('paste-text-area').value = "";
        document.getElementById('paste-excel-filename').value = "";

    } catch (err) {
        console.error("貼上名單彙整失敗:", err);
        alert("彙整名冊時發生錯誤：\n" + (err.message || err));
    } finally {
        if (actionBtn) {
            actionBtn.disabled = false;
            actionBtn.innerHTML = origBtnHtml;
        }
    }
}

// ==========================================
// 🚨 衝突與衝堂即時偵測邏輯 (UI/UX 優化)
// ==========================================

function showConflictAlert(resultData, elementId) {
    const alertDiv = document.getElementById(elementId);
    if (!alertDiv) return;

    const conflictingStudents = resultData.filter(s => s.remarks && s.remarks.length > 0);
    if (conflictingStudents.length > 0) {
        let html = `<div class="conflict-title">有 ${conflictingStudents.length} 位學生同一天上兩個社團：</div>`;
        html += `<ul class="conflict-list">`;
        conflictingStudents.forEach(s => {
            const conflictDetails = escapeHtml(s.remarks.join('；'));
            const safeClass = escapeHtml(s.class);
            const safeSeat = escapeHtml(s.seat);
            const safeName = escapeHtml(s.name);
            html += `<li><strong>${safeClass} 班 ${safeSeat} 號 ${safeName}</strong>：${conflictDetails}</li>`;
        });
        html += `</ul>`;
        alertDiv.innerHTML = html;
        alertDiv.style.display = 'block';
    } else {
        alertDiv.style.display = 'none';
        alertDiv.innerHTML = '';
    }
}

function updateMappingConflicts() {
    if (!uploadedWorkbook || sheetData.length === 0) return;

    const mode = document.getElementById('mode-select').value;
    const includeFreshmen = document.getElementById('include-freshmen').checked;
    const slotMode = document.getElementById('slot-mode-select').value;
    const rows = document.getElementById('mapping-table-body').querySelectorAll('tr');
    
    const finalMapping = {};
    rows.forEach(tr => {
        const input = tr.querySelector('.club-input');
        const select = tr.querySelector('.day-select');
        if (!input || !select) return;
        
        const origKey = input.dataset.original;
        const editedClub = input.value.trim();
        const selectedDay = select.value;

        finalMapping[origKey] = {
            editedName: editedClub,
            day: selectedDay,
            slot: tr.dataset.detectedSlot || ""
        };
    });

    const { resultData } = processExcelData(
        mode,
        includeFreshmen,
        slotMode,
        finalMapping,
        sheetData,
        detectedHeaders,
        colClassIdx,
        colSeatIdx,
        colNameIdx,
        slotCols
    );

    showConflictAlert(resultData, 'mapping-conflict-alert');
}

// ==========================================
// 🔘 按鈕事件綁定 (CSP 禁止 HTML 內的 onclick，一律在此綁定)
// ==========================================
document.getElementById('btn-hard-reset').addEventListener('click', manualHardResetAll);
document.getElementById('tab-btn-file').addEventListener('click', () => switchInputTab('file'));
document.getElementById('tab-btn-paste').addEventListener('click', () => switchInputTab('paste'));
document.getElementById('btn-load-club').addEventListener('click', loadPastedClub);
document.getElementById('btn-clear-clubs').addEventListener('click', clearAllPastedClubs);
document.getElementById('btn-process-paste').addEventListener('click', processPastedData);
document.getElementById('btn-back-to-upload').addEventListener('click', resetToUpload);
document.getElementById('btn-process-mapping').addEventListener('click', processAndDownload);
document.getElementById('btn-restart').addEventListener('click', resetToUpload);

// ==========================================
// 🛡️ 網頁防誤關閉與重新整理機制
// ==========================================
window.addEventListener('beforeunload', (e) => {
    // 只要有讀取 Excel 檔案或貼上載入了社團，即進行警告
    if (uploadedWorkbook !== null || (pastedClubs && pastedClubs.length > 0)) {
        const message = '您填寫的對照設定或載入的社團名單尚未儲存/下載，確定要離開嗎？';
        e.preventDefault();
        e.returnValue = message;
        return message;
    }
});

// ==========================================
// 🛡️ 全域未捕獲異常 (Unhandled Errors) 兜底防護
// ==========================================
window.addEventListener('unhandledrejection', (event) => {
    console.error('全域非同步例外未捕獲:', event.reason);
    const msg = event.reason?.message || event.reason || '系統發生未預期的非同步異常';
    // 若為非取消操作則跳出提示
    if (!String(msg).includes('abort')) {
        alert('發生非預期的錯誤，請重新整理網頁後再試一次。\n錯誤內容：' + msg);
    }
});

window.onerror = function(message, source, lineno, colno, error) {
    console.error('全域例外捕捉:', { message, source, lineno, colno, error });
    return false;
};
