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
// 多工作表模式：工作表名稱 → 學生名單（載入對照表時讀一次）
let sheetStudents = new Map();
// 上一次成功建立對照表的彙整模式：在第 2 步切換模式失敗時退回這個模式
let lastMappingMode = null;
// 正在讀取名冊（讀取中忽略新的檔案）
let uploadBusy = false;
// 每次開始讀取名冊與每次清除名冊狀態都加 1。handleFileSelect 在每個 await 之後比對，
// 不一致代表讀取途中使用者按了「清除這次的資料」等操作，這次讀取的結果直接丟掉
let fileLoadToken = 0;

// 上傳檔案大小上限。學校系統匯出的錄取名冊通常不到 1 MB；
// 太大的檔案會讓舊電腦的瀏覽器分頁記憶體不足而當掉（貼上名單也有 200,000 字元的上限）
const MAX_UPLOAD_MB = 20;
const MAX_UPLOAD_BYTES = MAX_UPLOAD_MB * 1024 * 1024;

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
    if (uploadBusy || e.target === fileInput || e.target.closest('a')) return;
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

// 以 Promise 讀取檔案內容（FileReader 支援的瀏覽器比 File.arrayBuffer() 多）
function readFileAsArrayBuffer(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => resolve(e.target.result);
        reader.onerror = () => {
            const err = new Error('FileReader 無法讀取檔案');
            err.name = 'FileReadError';
            reject(err);
        };
        reader.readAsArrayBuffer(file);
    });
}

// 讀取名冊期間：上傳區標示忙碌、停用選擇檔案，並顯示狀態文字（role="status" 會被螢幕閱讀器唸出）
// 「選擇檔案」按鈕用 aria-disabled 而不是 disabled：disabled 的按鈕會失去焦點（掉到 body），
// 鍵盤使用者讀取失敗留在第 1 步時就找不到原本的位置。按下去的 click 會冒泡到 dropZone，由 uploadBusy 擋下
function setUploadBusy(busy) {
    uploadBusy = busy;
    if (busy) dropZone.setAttribute('aria-busy', 'true');
    else dropZone.removeAttribute('aria-busy');
    dropZone.classList.toggle('is-busy', busy);
    const pickBtn = document.getElementById('btn-pick-file');
    if (pickBtn) {
        if (busy) pickBtn.setAttribute('aria-disabled', 'true');
        else pickBtn.removeAttribute('aria-disabled');
    }
    fileInput.disabled = busy;
    const status = document.getElementById('upload-status');
    if (status) status.textContent = busy ? '正在讀取名冊，請稍候…' : '';
}

// 讀取 Excel 檔案
async function handleFileSelect(file) {
    if (!file) return;
    if (uploadBusy) {
        // 讀取中又拖放或選了檔案：不中斷目前的讀取，但要讓使用者知道新檔案沒有被讀取
        const status = document.getElementById('upload-status');
        if (status) status.textContent = '仍在讀取上一個名冊，請稍候；讀完後再重新選擇或拖放新的檔案。';
        return;
    }

    if (file.size > MAX_UPLOAD_BYTES) {
        const sizeMb = (file.size / 1024 / 1024).toFixed(1);
        alert(`這個檔案有 ${sizeMb} MB，超過 ${MAX_UPLOAD_MB} MB 的上限，無法讀取。學校系統匯出的錄取名冊通常不到 1 MB，請確認選對了檔案；如果確定是名冊，請用 Excel 刪除不需要的工作表或圖片後另存，再重新選擇。`);
        fileInput.value = "";
        return;
    }

    const token = ++fileLoadToken;
    const isStale = () => token !== fileLoadToken;
    setUploadBusy(true);
    try {
        // 檔頭簽章檢查：精準攔截舊版 .xls (BIFF8) 或偽裝格式
        const formatCheck = await checkExcelFileFormat(file);
        if (isStale()) return;
        if (!formatCheck.valid) {
            alert(formatCheck.reason || "這不是可以讀取的 Excel 檔案。請改選學校系統匯出的 .xlsx 名冊。");
            fileInput.value = "";
            return;
        }

        const ExcelJSLib = await loadExcelJS();
        if (isStale()) return;
        const arrayBuffer = await readFileAsArrayBuffer(file);
        if (isStale()) return;
        const workbook = new ExcelJSLib.Workbook();
        await workbook.xlsx.load(arrayBuffer);
        if (isStale()) return;

        originalFileName = file.name;
        uploadedWorkbook = workbook;

        // 偵測是否存在「課程與學生清單」或「總表」工作表
        let hasSingleSheet = false;
        workbook.eachSheet((ws) => {
            const n = ws.name.trim();
            if (n.includes('課程與學生清單') || n.includes('總表')) {
                hasSingleSheet = true;
            }
        });

        // 依有無清單/總表自動切換預設模式
        const modeSelect = document.getElementById('mode-select');
        modeSelect.value = hasSingleSheet ? 'single' : 'multi';

        // 分析並渲染對照設定；名冊無法整理時已用 alert 說明，留在第 1 步並清掉讀進來的名冊
        if (analyzeAndRenderMapping()) {
            lastMappingMode = modeSelect.value;
            switchStep('step-mapping');
        } else {
            clearFileModeState();
        }
    } catch (err) {
        // 例外訊息可能含有儲存格內容，只記錄在主控台；alert 只顯示固定說明與例外類型
        console.error("載入 Excel 失敗:", err);
        if (isStale()) return; // 使用者已清除，不再提示
        clearFileModeState();
        if (err && err.name === 'ExcelJSLoadError') {
            alert(EXCELJS_LOAD_ERROR_MESSAGE);
        } else if (err && err.name === 'FileReadError') {
            alert("無法讀取這個檔案，可能正被其他程式使用。請關閉正在開啟這個檔案的程式（例如 Excel），再重新選擇。");
        } else {
            alert("無法讀取這個 Excel 檔案，可能已損毀或設有密碼。請用 Excel 開啟後另存成 .xlsx（不設密碼），再重新選擇。" + errorTypeNote(err));
        }
    } finally {
        // 已被清除時，clearFileModeState() 已經解除讀取中狀態，也可能已有新的讀取在進行
        if (!isStale()) setUploadBusy(false);
    }
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

// alert 附上的例外類型（例如「TypeError」）。不放例外訊息：ExcelJS 的例外訊息可能含有儲存格內容
function errorTypeNote(err) {
    const name = err && typeof err.name === 'string' && /^[A-Za-z]{1,40}$/.test(err.name) ? err.name : '';
    return name ? `\n（錯誤類型：${name}）` : '';
}

// 清除檔案模式讀進記憶體的名冊與對照表狀態
// 讀取中的名冊也一併作廢：遞增 fileLoadToken 讓進行中的 handleFileSelect 放棄結果，並解除讀取中狀態
function clearFileModeState() {
    fileLoadToken++;
    setUploadBusy(false);
    fileInput.value = "";
    uploadedWorkbook = null;
    sheetData = [];
    sheetStudents = new Map();
    detectedHeaders = [];
    colClassIdx = -1;
    colSeatIdx = -1;
    colNameIdx = -1;
    slotCols = [];
    lastMappingMode = null;
    document.getElementById('excel-filename').value = "";

    const mappingTableBody = document.getElementById('mapping-table-body');
    if (mappingTableBody) mappingTableBody.innerHTML = "";
    showSkippedSheets([]);
    showMemoryHint(0);
    showConflictAlert([], 'mapping-conflict-alert');
}

function resetToUpload() {
    clearFileModeState();
    switchStep('step-upload');
}

// 一鍵安全清除所有暫存個資
function manualHardResetAll() {
    // 已加入社團時提醒抽籤結果會一起清除：之後重新加入就是重新抽籤，抽籤次數會累加
    const lotteryNote = pastedClubs.length > 0 ? "已加入社團的抽籤結果會一起清除，之後重新加入會重新抽籤，抽籤次數會累加。" : "";
    if (confirm(`確定要清除這次的資料嗎？會清除選擇的 Excel 檔案、星期設定、已加入的社團，以及瀏覽器記住的社團星期，畫面會回到第一步。${lotteryNote}`)) {
        clearFileModeState();

        // 清除貼上模式狀態（文字框、待執行的預覽、自動填入的記號）；抽籤次數 clubDrawCounts 刻意不重設
        pastedClubs = [];
        resetPasteInput();
        document.getElementById('paste-club-name').value = "";
        document.getElementById('paste-day-select').value = PASTE_DEFAULT_DAY;
        document.getElementById('paste-excel-filename').value = "";

        // 清除瀏覽器記憶的社團對照設定
        clearLocalStorageConfig();

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

        const slotMode = document.getElementById('slot-mode-select').value;
        const { mode, finalMapping, rows, hasEmptyName, unselectedClubs } = collectMapping();

        if (hasEmptyName) {
            alert("有社團的名稱是空白的。請在每一列的「社團名稱」填入名稱，再開始轉換。");
            return;
        }

        if (unselectedClubs.length > 0) {
            alert(`以下社團還沒選擇上課星期。請選好星期後，再按「開始轉換並下載課表」：\n\n${unselectedClubs.join('\n')}`);
            return;
        }

        saveMappingConfig(mode, rows);

        const { resultData, activeDays } = runExcelProcessing(finalMapping);

        if (resultData.length === 0) {
            alert(emptyExcelResultMessage(finalMapping));
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
        if (err && err.name === 'ExcelJSLoadError') {
            alert(EXCELJS_LOAD_ERROR_MESSAGE + "\n星期設定都還在，連上網路後再按一次「開始轉換並下載課表」即可，不需要重新整理網頁。");
            return;
        }
        alert("匯出課表時發生錯誤，課表沒有下載。請重新整理網頁後再試一次；若仍失敗，請確認名冊是學校系統匯出的原始檔。" + errorTypeNote(err));
    } finally {
        if (actionBtn) {
            actionBtn.disabled = false;
            actionBtn.innerHTML = origBtnHtml;
        }
    }
}

/**
 * 在第 2 步切換彙整模式：重建對照表。
 * 新模式無法整理這份名冊時（analyzeAndRenderMapping 已用 alert 說明原因），退回上一個可用的模式並重建，
 * 使用者留在第 2 步、名冊不必重選；連原本的模式也失敗（不應發生）才回到第 1 步。
 */
function handleModeChange() {
    if (!uploadedWorkbook) return;
    const modeSelect = document.getElementById('mode-select');
    // 重建前記下目前每列修改過的名稱與星期，退回原模式時套回去
    const previousEdits = snapshotMappingEdits();
    if (analyzeAndRenderMapping()) {
        // 切到單一清單但沒有總表時，analyzeAndRenderMapping 會自己改回多工作表並重建，
        // 結果仍是原本的模式，修改要套回去
        if (modeSelect.value === lastMappingMode) restoreMappingEdits(previousEdits);
        lastMappingMode = modeSelect.value;
        return;
    }
    if (lastMappingMode && lastMappingMode !== modeSelect.value) {
        modeSelect.value = lastMappingMode;
        if (analyzeAndRenderMapping()) {
            restoreMappingEdits(previousEdits);
            return;
        }
    }
    resetToUpload();
}

// 對照表每列的鍵 → { 社團名稱, 星期 }（只在記憶體中，不寫入瀏覽器儲存空間）
function snapshotMappingEdits() {
    const edits = new Map();
    document.getElementById('mapping-table-body').querySelectorAll('tr').forEach(tr => {
        const input = tr.querySelector('.club-input');
        const select = tr.querySelector('.day-select');
        if (input && select) edits.set(tr.dataset.key, { editedName: input.value, day: select.value });
    });
    return edits;
}

function restoreMappingEdits(edits) {
    document.getElementById('mapping-table-body').querySelectorAll('tr').forEach(tr => {
        const saved = edits.get(tr.dataset.key);
        if (!saved) return;
        tr.querySelector('.club-input').value = saved.editedName;
        // 只套用選單裡有的星期
        if (weekdaysList.includes(saved.day)) tr.querySelector('.day-select').value = saved.day;
    });
    updateMappingConflicts();
}

// 綁定彙整模式切換監聽與衝突偵測監聽
document.getElementById('mode-select').addEventListener('change', handleModeChange);
// 新生是否納入只影響彙整結果，不重建對照表（否則已修改的名稱與星期會被重設）
document.getElementById('include-freshmen').addEventListener('change', updateMappingConflicts);
document.getElementById('slot-mode-select').addEventListener('change', updateMappingConflicts);

// ==========================================
// 直接貼上名單模式邏輯
// ==========================================

let pastedClubs = [];

// 本次開啟網頁後，每個社團名稱已經抽過幾次籤（社團名稱 → 次數），寫進匯出檔的「抽籤資訊」，看得出是否重抽過。
// 只放在記憶體，重新載入頁面才重設：「全部清除」、清除這次的資料、下載後清空社團都不重設，
// 否則移除後重新加入就能不留記錄地一直重抽。不可寫進 localStorage 等任何 storage。
const clubDrawCounts = new Map();

// 計數用的鍵：只要看起來是同一個名稱就算同一個社團，避免改一點寫法就能重抽而不被記下次數。
// 去掉不可見字元與所有空白（「桌遊 社」）、全形英數當成一般英數（NFKC，「ＡＩ社」）、英文不分大小寫（「LEGO社」）
function clubDrawKey(clubName) {
    return cleanInvisibleChars(clubName).normalize('NFKC').toLowerCase().replace(/\s+/g, '');
}

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
        // 清空時，仍是自動填入的社團名稱與星期一併還原（推測不出結果時的處理，見 parser.js）
        guessClubAndDayFromPastedText('');
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

// 預覽要解析全文（最多 200,000 字元），連續輸入時等停下來才解析一次。
// 貼上也會觸發 input 事件，不需要另外監聽 paste。
const PASTE_PREVIEW_DELAY_MS = 250;
let pastePreviewTimer = null;

function cancelPastePreview() {
    if (pastePreviewTimer !== null) {
        clearTimeout(pastePreviewTimer);
        pastePreviewTimer = null;
    }
}

function schedulePastePreview() {
    cancelPastePreview();
    // 清空時立刻隱藏預覽，不必等待
    if (!pasteArea.value.trim()) {
        updatePastePreview();
        return;
    }
    pastePreviewTimer = setTimeout(() => {
        pastePreviewTimer = null;
        updatePastePreview();
    }, PASTE_PREVIEW_DELAY_MS);
}

// 還有待執行的預覽時立刻執行（按「加入這個社團」前，讓社團名稱與星期的自動推測用到最新內容）
function flushPastePreview() {
    if (pastePreviewTimer !== null) {
        cancelPastePreview();
        updatePastePreview();
    }
}

// 清空文字框與預覽，並重設自動填入的記號（加入社團、移除全部、下載後、清除資料時使用）
function resetPasteInput() {
    cancelPastePreview();
    pasteArea.value = "";
    parsePreview.style.display = 'none';
    resetPasteAutoFill();
}

pasteArea.addEventListener('input', schedulePastePreview);

// 使用者自己選過星期後，貼上新名單不再自動改星期（見 parser.js 的 guessClubAndDayFromPastedText）
const pasteDaySelect = document.getElementById('paste-day-select');
pasteDaySelect.addEventListener('change', () => {
    pasteDaySelect.dataset.userChosen = 'true';
    delete pasteDaySelect.dataset.autoFilled;
});

// 監聽貼上模式的設定變更以即時更新衝突狀態
document.getElementById('paste-slot-mode-select').addEventListener('change', renderLoadedClubs);
document.getElementById('paste-include-freshmen').addEventListener('change', renderLoadedClubs);

// 移除所有已加入的社團
function clearAllPastedClubs() {
    if (confirm("確定要移除所有已加入的社團嗎？抽籤結果會一起清除。移除後要重新貼上名單並加入，加入時會重新抽籤，抽籤次數會累加。")) {
        // 抽籤次數 clubDrawCounts 刻意不重設，重新加入同名社團仍會累加
        pastedClubs = [];
        renderLoadedClubs();
        resetPasteInput();
    }
}

// 正取人數上限（一個社團的正取名額不會超過這個數字，打錯多一個 0 時可以攔下）
const MAX_PASTE_QUOTA = 500;

function loadPastedClub() {
    // 抽籤只用瀏覽器內建的密碼學亂數；沒有時不抽籤，也不退回 Math.random（結果無法保證公平）
    if (!hasSecureRandom()) {
        alert("這個瀏覽器版本太舊，不支援安全的亂數，無法公平抽籤。請改用最新版 Chrome、Edge 或 Safari 開啟這個網頁，再按「加入這個社團」。");
        return;
    }

    // 防抖中的預覽先跑完，社團名稱與星期的自動推測才會是最新內容
    flushPastePreview();

    const day = document.getElementById('paste-day-select').value;
    // trim() 不會去掉零寬空格等不可見字元，只貼了這些字元時也要當成「還沒有填社團名稱」
    const clubNameInput = cleanInvisibleChars(document.getElementById('paste-club-name').value);
    const pastedText = pasteArea.value.trim();

    const quotaText = String(document.getElementById('paste-quota').value).trim();
    const quota = /^\d+$/.test(quotaText) ? Number(quotaText) : NaN;
    if (!(quota >= 1 && quota <= MAX_PASTE_QUOTA)) {
        alert(`正取人數要填 1 到 ${MAX_PASTE_QUOTA} 之間的整數，目前是「${quotaText || '空白'}」。請修改正取人數，再按「加入這個社團」。`);
        return;
    }
    const gradeStart = document.getElementById('paste-grade-start').value;
    const gradeEnd = document.getElementById('paste-grade-end').value;
    const priority = document.getElementById('paste-priority').value;

    // 新生算 0 年級（與 processor.js 的 runClubLottery 相同），選單裡排在 1 年級前面
    const minG = gradeStart === "新生" ? 0 : parseInt(gradeStart, 10);
    const maxG = gradeEnd === "新生" ? 0 : parseInt(gradeEnd, 10);
    if (minG > maxG) {
        alert(`起始年級（${gradeOptionLabel(gradeStart)}）大於結束年級（${gradeOptionLabel(gradeEnd)}）。「新生」是最低的年級，請重新選擇招收年級，讓起始年級低於或等於結束年級，再按「加入這個社團」。`);
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
    
    // 同名社團這次開啟網頁後已經抽過籤：先確認是否重抽。放在所有輸入檢查之後，取消時不加入也不計次。
    // 訊息只寫次數，不寫社團名稱（使用者輸入的內容）
    const drawKey = clubDrawKey(clubNameInput);
    const previousDraws = clubDrawCounts.get(drawKey) || 0;
    // 同名社團還在清單裡時一併提醒：兩份都會匯出（工作表名稱會加上編號），課表也會兩份都算
    const stillListed = pastedClubs.some(c => clubDrawKey(c.clubName) === drawKey);
    const listedNote = stillListed ? "清單裡已經有同名的社團，兩份抽籤結果都會匯出；只要保留新的結果，請先移除舊的那一個。" : "";
    if (previousDraws > 0 && !confirm(`這個社團在這次開啟網頁後已經抽過 ${previousDraws} 次籤。再加入會重新抽籤，匯出檔的抽籤資訊會記錄為第 ${previousDraws + 1} 次抽籤。${listedNote}確定要重新抽籤嗎？`)) {
        return;
    }

    // 重複報名先去重（彙整課表時也會被併成同一位），再執行抽籤演算法
    const { students: applicants, duplicateCount } = dedupeApplicants(rawStudents);
    const finalStudentsList = runClubLottery(applicants, quota, gradeStart, gradeEnd, priority);
    // 抽籤完成才計次；抽籤丟出例外時不會走到這裡
    const drawRound = previousDraws + 1;
    clubDrawCounts.set(drawKey, drawRound);
    const drawnAt = new Date();
    const selectedCount = finalStudentsList.filter(s => s.selected).length;
    const unrecognizedCount = finalStudentsList.filter(s => s.drawSequence === UNRECOGNIZED_CLASS_LABEL).length;

    // 抽籤時間（drawnAt）、第幾次抽籤（drawRound）與各項筆數不是學生資料，只放在記憶體，下載時寫進匯出檔的「抽籤資訊」
    const newClub = {
        id: Date.now(),
        day: day,
        clubName: clubNameInput,
        quota: quota,
        gradeStart: gradeStart,
        gradeEnd: gradeEnd,
        priority: priority,
        students: finalStudentsList,
        selectedCount: selectedCount,
        applicantCount: rawStudents.length,
        duplicateCount: duplicateCount,
        unrecognizedCount: unrecognizedCount,
        drawnAt: drawnAt,
        drawRound: drawRound
    };

    pastedClubs.push(newClub);

    // 清空社團名稱與名單，並重設自動填入的記號，下一個社團重新推測名稱與星期
    document.getElementById('paste-club-name').value = "";
    resetPasteInput();

    renderLoadedClubs();

    // 有重複報名或班級無法辨識時提醒；訊息只有筆數，不含名冊內容
    const notice = lotteryNoticeMessage(duplicateCount, unrecognizedCount);
    if (notice) alert(notice);

    // 自動聚焦回社團名稱輸入框，方便打字下一個社團
    setTimeout(() => {
        const clubNameInput = document.getElementById('paste-club-name');
        if (clubNameInput) clubNameInput.focus();
    }, 50);
}

/**
 * 加入社團後的提醒：重複報名與班級無法辨識的筆數，以及下一步；兩者都沒有時回傳 ""
 * 只寫筆數，不可串入姓名、班級等名冊內容
 */
function lotteryNoticeMessage(duplicateCount, unrecognizedCount) {
    const parts = [];
    if (duplicateCount > 0) parts.push(`${duplicateCount} 筆重複報名（班級、姓名相同），重複的只保留一筆，已用去除重複後的名單抽籤`);
    if (unrecognizedCount > 0) parts.push(`${parts.length ? '另有 ' : ''}${unrecognizedCount} 筆班級無法辨識，沒有參加抽籤，名單中標示為「${UNRECOGNIZED_CLASS_LABEL}」`);
    if (parts.length === 0) return "";
    return `名單裡有 ${parts.join('；')}。請確認貼上的內容，需要時移除這個社團後重新貼上並加入（會重新抽籤）。`;
}

/**
 * 已加入社團那一行的招收年級，例如「限 3–5 年級」「限 4 年級」「限新生與 1–2 年級」「限新生」
 * 新生算 0 年級：結束年級是新生時，起始年級一定也是新生（起始大於結束時不會加入社團）
 */
function gradeRangeText(gradeStart, gradeEnd) {
    const start = String(gradeStart);
    const end = String(gradeEnd);
    if (start === "新生") {
        if (end === "新生") return "限新生";
        return end === "1" ? "限新生與 1 年級" : `限新生與 1–${end} 年級`;
    }
    return start === end ? `限 ${start} 年級` : `限 ${start}–${end} 年級`;
}

// 已加入社團那一行後面的說明：略過重複、班級無法辨識的筆數、抽籤時間，以及重抽時的第幾次抽籤
function loadedClubExtraMeta(club) {
    const extra = [];
    if (club.duplicateCount > 0) extra.push(`略過重複 ${club.duplicateCount} 筆`);
    if (club.unrecognizedCount > 0) extra.push(`班級無法辨識 ${club.unrecognizedCount} 筆`);
    const time = formatDrawTime(club.drawnAt);
    if (time) extra.push(`${time} 抽籤`);
    if (club.drawRound > 1) extra.push(`第 ${club.drawRound} 次抽籤`);
    return extra.map(t => `，${t}`).join('');
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
                               `<span class="loaded-club-meta">${escapeHtml(club.day)}，正取 ${escapeHtml(String(club.selectedCount))}／${escapeHtml(String(club.quota))} 人，報名 ${escapeHtml(String(club.students.length))} 人，${escapeHtml(gradeRangeText(club.gradeStart, club.gradeEnd))}${escapeHtml(loadedClubExtraMeta(club))}</span>` +
                               `<span class="loaded-club-badge">查看名單</span>`;

        const infoDetails = document.createElement('div');
        infoDetails.className = "loaded-club-details";
        infoDetails.hidden = true;

        // 先組成純文字，插入 HTML 時才統一跳脫，審查時一眼就能確認每個插入值都經過 escapeHtml()
        const selectedText = club.students.filter(s => s.selected).map(s => `${s.class || '新生'}${s.seat ? `座號${s.seat}` : ''} ${s.name}`).join('、');
        // 未錄取的順序與 runClubLottery() 回傳的相同：備取 1、2、3…，接著是班級無法辨識、不符限制
        const backupText = club.students.filter(s => !s.selected).map(s => `${s.name}(${s.drawSequence})`).join('、');

        infoDetails.innerHTML = `<div class="roster-selected"><strong>正取：</strong>${escapeHtml(selectedText || '無')}</div>` +
                                 `<div class="roster-backup"><strong>備取與未參加抽籤：</strong>${escapeHtml(backupText || '無')}</div>`;

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
        // 誤按會失去這個社團的抽籤結果，重新加入就是重新抽籤，所以移除前先確認；抽籤次數不會因移除而改變
        delBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            if (!confirm("確定要移除這個社團嗎？抽籤結果會一起清除。重新加入時會重新抽籤，抽籤次數會累加。")) return;
            pastedClubs.splice(index, 1);
            renderLoadedClubs();
            // 移除的按鈕已不存在，焦點移到下一個「移除」按鈕或貼上區，鍵盤使用者才不會失去焦點
            if (pastedClubs.length > 0) {
                const list = document.getElementById('loaded-clubs-list');
                const buttons = list ? list.querySelectorAll('.loaded-club-delete') : [];
                const next = buttons[Math.min(index, pastedClubs.length - 1)];
                if (next) next.focus();
            } else {
                const textArea = document.getElementById('paste-text-area');
                if (textArea) textArea.focus();
            }
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

/**
 * 下載時沒有可以彙整的學生：說明原因與該怎麼做。訊息只寫設定與按鈕名稱，不含名冊內容
 * 「包含沒有班級座號的新生」只影響彙整，勾回來再下載就好，不必重新加入（重新加入會重新抽籤、抽籤次數會累加），
 * 所以先實際判斷勾選時有沒有學生；只有其他原因時才說明要移除社團後重新加入
 */
function emptyPastedResultMessage(clubs, slotMode, includeFreshmen) {
    if (!includeFreshmen && processPastedClubsData(clubs, slotMode, true).resultData.length > 0) {
        return "已加入的社團中，正取的學生都沒有班級座號（新生），目前沒有勾選「包含沒有班級座號的新生」，所以課表沒有學生。" +
            "請勾選「包含沒有班級座號的新生」，再按一次「下載課表與抽籤結果」。已加入的社團與抽籤結果都還在，不需要重新加入，也不會重新抽籤。";
    }
    return "已加入的社團都沒有正取的學生，沒有課表可以下載。通常是報名的學生都不在招收年級內（名單中標示為「不符限制」），或班級無法辨識。" +
        "請點社團那一列查看名單確認。如果是招收年級選錯了，要先移除那個社團，改好招收年級後重新貼上名單並加入；重新加入會重新抽籤，抽籤次數會累加。";
}

async function processPastedData() {
    const actionBtn = document.querySelector('#loaded-clubs-section .btn-primary');
    const origBtnHtml = actionBtn ? actionBtn.innerHTML : "";

    try {
        if (pastedClubs.length === 0) {
            alert("還沒有加入任何社團。請先貼上名單並按「加入這個社團」，再按「下載課表與抽籤結果」。");
            return;
        }
        
        if (actionBtn) {
            actionBtn.disabled = true;
            actionBtn.textContent = '正在下載課表與抽籤結果…';
        }

        const slotMode = document.getElementById('paste-slot-mode-select').value;
        const includeFreshmen = document.getElementById('paste-include-freshmen').checked;
        
        // 呼叫 Core 彙整處理器
        const { resultData, activeDays } = processPastedClubsData(pastedClubs, slotMode, includeFreshmen);
        
        if (resultData.length === 0) {
            alert(emptyPastedResultMessage(pastedClubs, slotMode, includeFreshmen));
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
        
        // 重設狀態（畫面在 finally 還原按鈕後才重繪）；抽籤次數 clubDrawCounts 刻意不重設
        pastedClubs = [];
        document.getElementById('paste-excel-filename').value = "";
        // 改 value 不會觸發 input 事件，預覽要手動隱藏，否則會留著舊的筆數
        resetPasteInput();

    } catch (err) {
        console.error("貼上名單彙整失敗:", err);
        if (err && err.name === 'ExcelJSLoadError') {
            alert(EXCELJS_LOAD_ERROR_MESSAGE + "\n已加入的社團與抽籤結果都還在，連上網路後再按一次「下載課表與抽籤結果」即可，不需要重新整理網頁。");
            return;
        }
        // 下載不會抽籤（抽籤在 loadPastedClub），失敗時已加入的社團與抽籤結果都還在；重新整理網頁會清掉結果，所以先請使用者再按一次
        alert("匯出課表時發生錯誤，課表沒有下載。已加入的社團與抽籤結果都還在，請再按一次「下載課表與抽籤結果」。若仍失敗，請確認貼上的名單包含班級、座號、姓名；重新整理網頁會清除抽籤結果，要重新加入社團並重新抽籤。" + errorTypeNote(err));
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
        // 衝堂的判斷見 finalizeStudents()：下午同一天，或晨間同一天（含不同晨間欄位的上課日重疊）；晨間與下午不算
        let html = WARNING_ICON_SVG;
        html += `<div><div class="conflict-title">有 ${escapeHtml(String(conflictingStudents.length))} 位學生同一天的同一時段（下午或晨間）有兩個以上社團：</div>`;
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

/**
 * 讀取對照表畫面上的設定
 * 對照鍵在單一清單模式是「時段欄位＋社團」（同一社團可能在 A、B 時段不同天），多工作表模式是工作表名稱
 */
function collectMapping() {
    const mode = document.getElementById('mode-select').value;
    // 多工作表模式的鍵是工作表名稱，用沒有原型的物件，名為「__proto__」的工作表才不會消失
    const finalMapping = createDict();
    const rows = [];
    const unselectedClubs = [];
    let hasEmptyName = false;

    document.getElementById('mapping-table-body').querySelectorAll('tr').forEach(tr => {
        const input = tr.querySelector('.club-input');
        const select = tr.querySelector('.day-select');
        if (!input || !select) return;

        const row = {
            key: tr.dataset.key,
            original: tr.dataset.original,
            slotCol: tr.dataset.slot || "",
            editedName: input.value.trim(),
            day: select.value
        };
        rows.push(row);

        if (!row.editedName) hasEmptyName = true;
        if (row.day === "請選擇") {
            unselectedClubs.push(mode === 'single' ? `${row.slotCol} 的「${row.original}」` : `工作表「${row.original}」`);
        }

        finalMapping[row.key] = {
            editedName: row.editedName,
            day: row.day,
            slot: tr.dataset.detectedSlot || ""
        };
    });

    return { mode, finalMapping, rows, hasEmptyName, unselectedClubs };
}

/**
 * 記住這次的星期設定（只存社團／工作表名稱 → 星期，不含學生資料）
 * 單一清單模式依時段欄位分開存，避免 A、B 時段的同名社團互相覆蓋
 * 每學期的社團都會累積：這次用到的鍵移到最後，超過上限時刪除最久沒用到的（見 utils.js 的 setConfigEntry）
 */
function saveMappingConfig(mode, rows) {
    const savedConfig = loadLocalStorageConfig();
    rows.forEach(row => {
        // 太長的名稱（超過 CONFIG_MAX_KEY_LENGTH）不存，見 utils.js
        if (!isStorableConfigKey(row.original)) return;
        if (mode === 'single') {
            if (!isStorableConfigKey(row.slotCol)) return;
            // 最上層的鍵可能原本是字串（多工作表模式的工作表、或舊版格式的社團），這裡會改成時段欄位的物件；
            // 少記一個星期的影響，比無法記住整個時段欄位的社團小
            const slotMap = getSavedSlotMap(savedConfig, row.slotCol);
            setConfigEntry(slotMap, row.original, row.day, CONFIG_MAX_CLUBS_PER_SLOT);
            setConfigEntry(savedConfig, row.slotCol, slotMap, CONFIG_MAX_TOP_KEYS);
        } else {
            // 單一清單模式的時段欄位與多工作表模式的工作表共用最上層的鍵。
            // 不改儲存格式就無法同時保存兩者，所以同名時保留時段欄位的物件（一整組社團的星期），
            // 不用一個工作表的星期把它蓋掉；這個工作表這次的星期就不記
            const existing = savedConfig[row.original];
            if (existing && typeof existing === 'object') return;
            setConfigEntry(savedConfig, row.original, row.day, CONFIG_MAX_TOP_KEYS);
        }
    });
    saveLocalStorageConfig(savedConfig);
}

// includeFreshmen 預設依畫面上的勾選；emptyExcelResultMessage() 會傳 true 判斷勾選後有沒有學生
function runExcelProcessing(finalMapping, includeFreshmen = document.getElementById('include-freshmen').checked) {
    return processExcelData({
        mode: document.getElementById('mode-select').value,
        includeFreshmen,
        slotMode: document.getElementById('slot-mode-select').value,
        finalMapping,
        sheetData,
        classCol: detectedHeaders[colClassIdx],
        seatCol: detectedHeaders[colSeatIdx],
        nameCol: detectedHeaders[colNameIdx],
        slotCols,
        sheetStudents
    });
}

/**
 * 上傳 Excel 轉換時沒有可以匯出的學生：說明原因與該怎麼做。訊息不含名冊內容
 * 星期沒選好的社團在這之前就會擋下，所以不再請使用者檢查星期；
 * 取消勾選「包含沒有班級座號的新生名單」時，先實際判斷勾選後有沒有學生
 */
function emptyExcelResultMessage(finalMapping) {
    if (!document.getElementById('include-freshmen').checked) {
        if (runExcelProcessing(finalMapping, true).resultData.length > 0) {
            return "名冊中的學生都沒有班級座號（新生），目前沒有勾選「包含沒有班級座號的新生名單」，所以沒有學生可以匯出。" +
                "請勾選「包含沒有班級座號的新生名單」，再按一次「開始轉換並下載課表」，星期設定都還在。";
        }
    }
    return "名冊中讀不到有姓名的學生，沒有資料可以匯出。請確認選擇的是學校系統匯出的錄取名冊，而且「姓名」欄有填入學生姓名，再重新選擇檔案。";
}

function updateMappingConflicts() {
    if (!uploadedWorkbook) return;
    const { resultData } = runExcelProcessing(collectMapping().finalMapping);
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
// 預先載入 ExcelJS
// ==========================================
// 頁面載入完成後趁閒置時載入，使用者選檔或下載時通常已經載好。
// 這裡失敗（例如第一次開啟就離線）不打擾使用者，真正要用時 loadExcelJS() 會再試一次並顯示訊息。
function preloadExcelJS() {
    loadExcelJS().catch(err => console.warn('[ExcelJS] 預先載入失敗，使用時會再試一次:', err.message));
}
window.addEventListener('load', () => {
    if ('requestIdleCallback' in window) {
        window.requestIdleCallback(preloadExcelJS, { timeout: 3000 });
    } else {
        setTimeout(preloadExcelJS, 1000);
    }
});

// ==========================================
// 網頁防誤關閉與重新整理機制
// ==========================================
window.addEventListener('beforeunload', (e) => {
    // 只要有讀取 Excel 檔案或加入了社團，即進行警告
    if (uploadedWorkbook !== null || (pastedClubs && pastedClubs.length > 0)) {
        // 新版瀏覽器大多只顯示內建文字，這段訊息只在支援自訂文字的瀏覽器出現
        const message = (pastedClubs && pastedClubs.length > 0)
            ? '已加入的社團還沒下載成課表，離開後抽籤結果會消失，要重新加入並重新抽籤。確定要離開嗎？'
            : '星期設定還沒下載成課表，離開後要重新開始。確定要離開嗎？';
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
    const reason = event.reason;
    // 使用者取消的操作（AbortError）不提示
    const isAbort = (reason && reason.name === 'AbortError') || String((reason && reason.message) || reason || '').includes('abort');
    if (!isAbort) {
        // 例外內容可能含有名冊資料，alert 只顯示固定說明與例外類型，完整內容在上面的 console.error
        // 已加入社團時不要先叫使用者重新整理：重新整理會清除抽籤結果，等於要重新抽籤
        const hint = (pastedClubs && pastedClubs.length > 0)
            ? '已加入的社團與抽籤結果應該還在，請先再試一次剛才的操作；仍然失敗才重新整理網頁，但重新整理會清除抽籤結果，要重新加入社團並重新抽籤。'
            : '請重新整理網頁後再試一次。';
        alert('發生非預期的錯誤。' + hint + errorTypeNote(reason));
    }
});

window.onerror = function(message, source, lineno, colno, error) {
    console.error('全域例外捕捉:', { message, source, lineno, colno, error });
    return false;
};
