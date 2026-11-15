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
// 檔案拖曳與讀取邏輯
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

    // 檔頭簽章檢查：精準攔截舊版 .xls (BIFF8) 或偽裝格式
    if (typeof checkExcelFileFormat === 'function') {
        const formatCheck = await checkExcelFileFormat(file);
        if (!formatCheck.valid) {
            alert(formatCheck.reason || "這不是可以讀取的 Excel 檔案。請改選學校系統匯出的 .xlsx 名冊。");
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
            alert("無法讀取這個 Excel 檔案，可能已損毀或設有密碼。請用 Excel 開啟後另存成 .xlsx（不設密碼），再重新選擇。\n錯誤內容：" + (err.message || err));
            if (fileInput) fileInput.value = "";
        }
    };

    reader.onerror = function() {
        alert("無法讀取這個檔案，可能正被其他程式使用。請關閉正在開啟這個檔案的程式（例如 Excel），再重新選擇。");
        if (fileInput) fileInput.value = "";
    };
    
    reader.readAsArrayBuffer(file);
}

// ==========================================
// 步驟切換
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

// 一鍵安全清除所有暫存個資
function manualHardResetAll() {
    if (confirm("確定要清除這次的資料嗎？會清除選擇的 Excel 檔案、星期設定、已加入的社團，以及瀏覽器記住的社團星期，畫面會回到第一步。")) {
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
// 資料轉換與下載邏輯 (檔案上傳模式)
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
                alert("有社團的名稱是空白的。請在每一列的「社團名稱」填入名稱，再開始轉換。");
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
            alert(`以下社團還沒選擇上課星期。請選好星期後，再按「開始轉換並下載課表」：\n\n${unselectedClubs.join('\n')}`);
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
            alert("名冊中沒有可以匯出的學生資料。請確認選擇的是錄取名冊，並檢查每個社團都選了上課星期。");
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
        alert("匯出課表時發生錯誤，課表沒有下載。請重新整理網頁後再試一次；若仍失敗，請確認名冊是學校系統匯出的原始檔。\n錯誤內容：" + (err.message || err));
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
// 直接貼上名單模式邏輯
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
        // 無法處理是錯誤狀態，用錯誤色系（DESIGN.md「錯誤／刪除」）
        parsePreview.innerHTML = `<strong>無法讀取：</strong>貼上的內容超過 200,000 字元。請一次只貼上一個社團的名單。`;
        parsePreview.classList.remove('is-warning');
        parsePreview.classList.add('is-error');
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
        
        // 插入值雖然都是數字，仍依規則一律經過 escapeHtml()
        parsePreview.innerHTML = `讀到 <strong>${escapeHtml(String(parsed.length))}</strong> 位學生` +
                                 `<span class="parse-preview-detail">其中 ${escapeHtml(String(hasClass))} 位有班級、${escapeHtml(String(hasSeat))} 位有座號、${escapeHtml(String(hasName))} 位有姓名</span>`;
        parsePreview.classList.remove('is-warning', 'is-error');
        parsePreview.style.display = 'block';
    } else {
        // 讀不到資料時使用者可以修正貼上內容，屬於警告而非錯誤
        parsePreview.innerHTML = `<strong>讀不到學生資料。</strong>請確認貼上的內容包含班級、座號、姓名這幾欄。`;
        parsePreview.classList.remove('is-error');
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

// 移除所有已加入的社團
function clearAllPastedClubs() {
    if (confirm("確定要移除所有已加入的社團嗎？移除後要重新貼上名單並加入。")) {
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
        alert(`起始年級（${gradeStart}）大於結束年級（${gradeEnd}）。請重新選擇年級範圍，再按「加入這個社團」。`);
        return;
    }

    if (!clubNameInput) {
        alert("還沒有填社團名稱。請先填入社團名稱，再按「加入這個社團」。");
        return;
    }
    
    if (pastedText.length > 200000) {
        alert("貼上的內容超過 200,000 字元，無法處理。請一次只貼上一個社團的名單。");
        return;
    }
    
    if (!pastedText) {
        alert("還沒有貼上名單。請先把選課網頁上的名單複製並貼到文字框，再按「加入這個社團」。");
        return;
    }
    
    let rawStudents = parsePastedText(pastedText);
    if (rawStudents.length === 0) {
        alert("讀不到學生資料。請確認貼上的內容包含標題列（班級、座號、姓名）與學生資料列，再按「加入這個社團」。");
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

        // 社團名稱與學生姓名來自使用者貼上的內容，插入 HTML 前一律跳脫
        const infoHeader = document.createElement('div');
        infoHeader.className = "loaded-club-info";
        infoHeader.setAttribute('role', 'button');
        infoHeader.setAttribute('tabindex', '0');
        infoHeader.setAttribute('aria-expanded', 'false');
        infoHeader.innerHTML = `<span class="loaded-club-name">${escapeHtml(club.clubName)}</span>` +
                               `<span class="loaded-club-meta">${escapeHtml(club.day)}，正取 ${escapeHtml(String(club.selectedCount))}／${escapeHtml(String(club.quota))} 人，報名 ${escapeHtml(String(club.students.length))} 人，限 ${escapeHtml(club.gradeStart)}–${escapeHtml(club.gradeEnd)} 年級</span>` +
                               `<span class="loaded-club-badge">查看名單</span>`;

        const infoDetails = document.createElement('div');
        infoDetails.className = "loaded-club-details";
        infoDetails.hidden = true;

        // 先組成純文字，插入 HTML 時才統一跳脫，審查時一眼就能確認每個插入值都經過 escapeHtml()
        const selectedText = club.students.filter(s => s.selected).map(s => `${s.class || '新生'}${s.seat ? `座號${s.seat}` : ''} ${s.name}`).join('、');
        const backupText = club.students.filter(s => !s.selected).map(s => `${s.name}(${s.drawSequence})`).join('、');

        infoDetails.innerHTML = `<div class="roster-selected"><strong>正取：</strong>${escapeHtml(selectedText || '無')}</div>` +
                                 `<div class="roster-backup"><strong>備取或不符資格：</strong>${escapeHtml(backupText || '無')}</div>`;

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
    
    // 下載處理中按鈕文字會被暫時替換，筆數的 <span> 可能不存在，所以要先判斷
    if (pastedClubs.length > 0) {
        section.style.display = "block";
        if (totalSpan) totalSpan.textContent = String(totalSelectedStudents);
        
        // 即時衝突檢查
        const slotMode = document.getElementById('paste-slot-mode-select').value;
        const includeFreshmen = document.getElementById('paste-include-freshmen').checked;
        const { resultData } = processPastedClubsData(pastedClubs, slotMode, includeFreshmen);
        showConflictAlert(resultData, "paste-conflict-alert");
    } else {
        section.style.display = "none";
        if (totalSpan) totalSpan.textContent = "0";
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
            alert("還沒有加入任何社團。請先貼上名單並按「加入這個社團」，再下載課表。");
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
            alert("已加入的社團中沒有可以彙整的學生。請確認名單有學生，或調整年級範圍與「包含沒有班級座號的新生」設定後重新加入。");
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
        
        // 重設狀態（畫面在 finally 還原按鈕後才重繪）
        pastedClubs = [];
        document.getElementById('paste-text-area').value = "";
        document.getElementById('paste-excel-filename').value = "";
        // 改 value 不會觸發 input 事件，預覽要手動隱藏，否則會留著舊的筆數
        parsePreview.style.display = 'none';

    } catch (err) {
        console.error("貼上名單彙整失敗:", err);
        alert("抽籤或匯出課表時發生錯誤，課表沒有下載。請重新整理網頁後再試一次；若仍失敗，請確認貼上的名單包含班級、座號、姓名。\n錯誤內容：" + (err.message || err));
    } finally {
        if (actionBtn) {
            actionBtn.disabled = false;
            actionBtn.innerHTML = origBtnHtml;
        }
        // 必須在還原按鈕之後重繪：處理中按鈕的文字被替換，裡面的筆數 <span> 暫時不存在，
        // 之前在 try 裡重繪會丟出例外，下載成功卻跳出「課表沒有下載」
        renderLoadedClubs();
    }
}

// ==========================================
// 衝突與衝堂即時偵測邏輯 (UI/UX 優化)
// ==========================================

function showConflictAlert(resultData, elementId) {
    const alertDiv = document.getElementById(elementId);
    if (!alertDiv) return;

    const conflictingStudents = resultData.filter(s => s.remarks && s.remarks.length > 0);
    if (conflictingStudents.length > 0) {
        // 與其他 .warning-callout 相同的結構：左側線條圖示＋內容（DESIGN.md 元件表「提醒」）
        let html = `<svg class="icon" viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 3l10 18H2z"/><path d="M12 10v5M12 18h.01"/></svg>`;
        html += `<div><div class="conflict-title">有 ${escapeHtml(String(conflictingStudents.length))} 位學生同一天上兩個社團：</div>`;
        html += `<ul class="conflict-list">`;
        conflictingStudents.forEach(s => {
            // 姓名、班級、座號、社團都來自名冊，插入點一律 escapeHtml()；座號可能是數字，先轉成字串
            html += `<li><strong>${escapeHtml(String(s.class ?? ''))} 班 ${escapeHtml(String(s.seat ?? ''))} 號 ${escapeHtml(String(s.name ?? ''))}</strong>：${escapeHtml(s.remarks.join('；'))}</li>`;
        });
        html += `</ul></div>`;
        alertDiv.innerHTML = html;
        alertDiv.style.display = 'flex';
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
// 按鈕事件綁定 (CSP 禁止 HTML 內的 onclick，一律在此綁定)
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
// 網頁防誤關閉與重新整理機制
// ==========================================
window.addEventListener('beforeunload', (e) => {
    // 只要有讀取 Excel 檔案或加入了社團，即進行警告
    if (uploadedWorkbook !== null || (pastedClubs && pastedClubs.length > 0)) {
        const message = '星期設定或已加入的社團還沒下載成課表，離開後要重新開始。確定要離開嗎？';
        e.preventDefault();
        e.returnValue = message;
        return message;
    }
});

// ==========================================
// 全域未捕獲異常 (Unhandled Errors) 兜底防護
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
