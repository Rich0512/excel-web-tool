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
    originalFileName = file.name;
    const reader = new FileReader();
    
    reader.onload = async function(e) {
        const arrayBuffer = e.target.result;
        const workbook = new ExcelJS.Workbook();
        
        try {
            await workbook.xlsx.load(arrayBuffer);
            uploadedWorkbook = workbook;
            
            // 偵測是否存在「總表」工作表
            let hasTotalSheet = false;
            workbook.eachSheet((ws) => {
                if (ws.name.trim() === '總表') {
                    hasTotalSheet = true;
                }
            });
            
            // 依有無「總表」自動切換預設模式
            document.getElementById('mode-select').value = hasTotalSheet ? 'single' : 'multi';
            
            // 分析並渲染對照設定
            analyzeAndRenderMapping();
            switchStep('step-mapping');

        } catch (err) {
            alert("載入 Excel 失敗，可能檔案格式毀損：" + err.message);
        }
    };
    
    reader.readAsArrayBuffer(file);
}

// ==========================================
// ⚙️ 視窗按鍵與步驟切換
// ==========================================

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
    document.getElementById('excel-filename').value = "";
    switchStep('step-upload');
}

// ==========================================
// ⚡ 資料轉換與下載邏輯 (檔案上傳模式)
// ==========================================

async function processAndDownload() {
    const mode = document.getElementById('mode-select').value;
    const includeFreshmen = document.getElementById('include-freshmen').checked;
    const slotMode = document.getElementById('slot-mode-select').value;

    const rows = document.getElementById('mapping-table-body').querySelectorAll('tr');
    
    const finalMapping = {};
    const unselectedClubs = [];

    rows.forEach(tr => {
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
    });

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
}

// 綁定彙整模式切換監聽
document.getElementById('mode-select').addEventListener('change', analyzeAndRenderMapping);

// ==========================================
// 📋 直接貼上名單模式邏輯
// ==========================================

let pastedClubs = [];

function switchInputTab(tab) {
    document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));
    
    if (tab === 'file') {
        document.getElementById('tab-btn-file').classList.add('active');
        document.getElementById('tab-file-content').classList.add('active');
    } else {
        document.getElementById('tab-btn-paste').classList.add('active');
        document.getElementById('tab-paste-content').classList.add('active');
    }
}

// 監聽輸入與貼上事件，進行即時解析預覽與特徵猜測
const pasteArea = document.getElementById('paste-text-area');
const parsePreview = document.getElementById('paste-parse-preview');

function updatePastePreview() {
    const text = pasteArea.value;
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
        
        parsePreview.innerHTML = `🟢 <strong>系統已即時識別：</strong>偵測到 <strong>${parsed.length}</strong> 筆學生名冊資料！<br>` +
                                 `<span style="opacity: 0.95; font-size: 11px;">(包含：${hasClass} 個有效班級、${hasSeat} 個座號、${hasName} 個姓名，名單格式解析正常)</span>`;
        parsePreview.style.display = 'block';
    } else {
        parsePreview.innerHTML = `⚠️ <strong>解析提示：</strong>目前貼上的文字無法識別出學生資料，請確認是否包含「姓名、班級、座號」資料列。`;
        parsePreview.style.display = 'block';
    }
}

// 同時監聽 input 與 paste 事件，確保任何編輯都能觸發預覽
pasteArea.addEventListener('input', updatePastePreview);
pasteArea.addEventListener('paste', () => {
    setTimeout(updatePastePreview, 50);
});

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
}

function renderLoadedClubs() {
    const listDiv = document.getElementById('loaded-clubs-list');
    listDiv.innerHTML = "";
    
    let totalSelectedStudents = 0;
    
    pastedClubs.forEach((club, index) => {
        totalSelectedStudents += club.selectedCount;
        
        const item = document.createElement('div');
        item.className = "loaded-club-item";
        
        const info = document.createElement('span');
        info.className = "loaded-club-info";
        info.textContent = `📅 ${club.day} ── 🏫 ${club.clubName} (正取 ${club.selectedCount}/${club.quota} 人，報名 ${club.students.length} 人，限制 ${club.gradeStart}-${club.gradeEnd}年級)`;
        
        const delBtn = document.createElement('button');
        delBtn.className = "loaded-club-delete";
        delBtn.innerHTML = "🗑️ 刪除";
        delBtn.onclick = () => {
            pastedClubs.splice(index, 1);
            renderLoadedClubs();
        };
        
        item.appendChild(info);
        item.appendChild(delBtn);
        listDiv.appendChild(item);
    });
    
    const section = document.getElementById('loaded-clubs-section');
    const totalSpan = document.getElementById('total-pasted-students');
    
    if (pastedClubs.length > 0) {
        section.style.display = "block";
        totalSpan.textContent = totalSelectedStudents;
    } else {
        section.style.display = "none";
        totalSpan.textContent = "0";
    }
}

async function processPastedData() {
    if (pastedClubs.length === 0) {
        alert("請先載入至少一個社團名單！");
        return;
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
}
