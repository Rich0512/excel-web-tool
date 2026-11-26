// 第 2 步對照表的提示：
// 1. analyzeAndRenderMapping() 的 alert 只給使用者做得到的建議。第一次讀檔失敗時畫面留在第 1 步（沒有彙整模式選單，
//    模式由工作表名稱決定），不可叫使用者「把彙整模式改成…」；在第 2 步切換模式失敗時會自動切回原模式，訊息要說明。
// 2. 「已套用這台電腦上次記住的星期」只在至少一個社團的星期真的取自記住的設定時顯示。
// 名冊內容都是虛構的（測試生301_05…）。
import { loadSite, makeReporter, makeWorkbook, uploadSheets, readSiteFile } from './harness.mjs';
const R = makeReporter('對照表提示：錯誤訊息的建議與記住設定的提示');

const KEY = 'yenping_club_config';
const SECRET = '測試生301_05';
const IMPOSSIBLE = /改成「多工作表」|改成「單一清單」|把「彙整模式」改/;
const onMapping = (s) => s.el('step-mapping').classList.contains('active');
const rows = (s) => s.el('mapping-table-body').querySelectorAll('tr');
const hint = (s) => s.el('memory-hint');
const hintText = (s) => s.el('memory-hint-text').textContent;
const STUDENTS = [['班級', '座號', '姓名'], ['301', '05', SECRET]];
const SINGLE = [{ name: '總表', rows: [['班級', '座號', '姓名', '社團'], ['301', '05', SECRET, '桌遊社'], ['302', '11', '測試生302_11', '直笛社']] }];

const checkMessage = (title, s, { must = [], mustNot = [] }) => {
    const msg = s.alerts.join('\n');
    R.check(`${title}：有 1 則提示`, s.alerts.length === 1, `${s.alerts.length} 則`);
    R.check(`${title}：沒有做不到的「把彙整模式改成…」建議`, !IMPOSSIBLE.test(msg), msg.replace(/\n/g, ' ').slice(0, 120));
    R.check(`${title}：不含名冊內容`, !msg.includes(SECRET));
    for (const m of must) R.check(`${title}：說明「${m}」`, msg.includes(m), msg.replace(/\n/g, ' ').slice(0, 160));
    for (const m of mustNot) R.check(`${title}：不說「${m}」`, !msg.includes(m), msg.replace(/\n/g, ' ').slice(0, 160));
};

// ---- 第一次讀檔失敗：留在第 1 步，建議改名冊本身
for (const [title, sheets, must] of [
    ['第一次讀檔、總表缺班級欄', [{ name: '總表', rows: [['座號', '姓名', '社團'], ['05', SECRET, '桌遊社']] }],
        ['工作表「總表」', '班級、座號、姓名', '把「總表」工作表刪除', '會被略過', '會被當成一個社團', '重新選擇檔案']],
    ['第一次讀檔、總表找不到社團欄', [{ name: '社團總表', rows: STUDENTS }],
        ['工作表「社團總表」', '「課程名稱」「社團」或「時段」', '把「社團總表」工作表刪除', '會被當成一個社團', '重新選擇檔案']],
    ['第一次讀檔、工作表都被略過', [{ name: '說明', rows: [['這是說明']] }, { name: '報名統計', rows: STUDENTS }],
        ['會略過', '- 說明', '- 報名統計', '改名為「總表」', '重新選擇檔案']],
    ['第一次讀檔、工作表都沒有姓名欄', [{ name: '社團-週一桌遊社', rows: [['班級', '座號'], ['301', '05']] }],
        ['「姓名」欄', '- 社團-週一桌遊社', '加上「姓名」欄位標題', '重新選擇檔案']],
]) {
    const s = loadSite();
    await uploadSheets(s, sheets);
    R.check(`${title}：（前提）留在第 1 步`, !onMapping(s));
    checkMessage(title, s, { must, mustNot: ['已切回原本的彙整模式'] });
}

// ---- 第 2 步切到多工作表失敗：說明已切回原本的模式，不叫使用者重選檔案
{
    const s = loadSite();
    await uploadSheets(s, SINGLE);
    R.check('（前提）單一清單檔案進入第 2 步', onMapping(s) && rows(s).length === 2);
    s.alerts.length = 0;
    const sel = s.el('mode-select');
    sel.value = 'multi';
    sel.dispatch('change');
    R.check('第 2 步切到多工作表失敗：（前提）退回單一清單並留在第 2 步', sel.value === 'single' && onMapping(s));
    checkMessage('第 2 步切到多工作表失敗', s, { must: ['「多工作表」', '已切回原本的彙整模式', '都會保留'], mustNot: ['重新選擇檔案'] });
}
{
    // 多工作表找到工作表但都沒有姓名欄
    const s = loadSite();
    await uploadSheets(s, [...SINGLE, { name: '社團-週一桌遊社', rows: [['班級', '座號'], ['301', '05']] }]);
    s.alerts.length = 0;
    const sel = s.el('mode-select');
    sel.value = 'multi';
    sel.dispatch('change');
    R.check('第 2 步切到多工作表、都沒有姓名欄：（前提）退回單一清單', sel.value === 'single' && onMapping(s));
    checkMessage('第 2 步切到多工作表、都沒有姓名欄', s, { must: ['「姓名」欄', '已切回原本的彙整模式'], mustNot: ['重新選擇檔案'] });
}
{
    // 直接在第 2 步以單一清單重建失敗（總表缺欄位）：訊息說明切回，不叫使用者刪工作表
    const s = loadSite();
    await uploadSheets(s, [{ name: '社團-週一桌遊社', rows: STUDENTS }]);
    s.alerts.length = 0;
    s.ctx.__wb = makeWorkbook([{ name: '總表', rows: [['座號', '姓名', '社團'], ['05', SECRET, '桌遊社']] }]);
    s.run('uploadedWorkbook = __wb');
    s.el('mode-select').value = 'single';
    const ok = s.run('analyzeAndRenderMapping()');
    R.check('第 2 步以單一清單重建失敗：回傳 false', ok === false, String(ok));
    checkMessage('第 2 步以單一清單重建失敗', s, { must: ['工作表「總表」', '已切回原本的彙整模式'], mustNot: ['工作表刪除', '重新選擇檔案'] });
}
{
    // 第 2 步切到單一清單、沒有總表而自動改用多工作表：仍只有一則提醒，且在對照表重建後才跳出
    const s = loadSite();
    await uploadSheets(s, [{ name: '社團-週一桌遊社', rows: STUDENTS }]);
    s.alerts.length = 0;
    const sel = s.el('mode-select');
    sel.value = 'single';
    sel.dispatch('change');
    R.check('第 2 步切到單一清單（沒有總表）：自動改用多工作表並提醒 1 次', sel.value === 'multi' && s.alerts.length === 1 && s.alerts[0].includes('已自動改用「多工作表」'), s.alerts.join(' / '));
    s.el('mode-select').value = 'single';
    s.run('uploadedWorkbook = { eachSheet() {} }');
    s.alerts.length = 0;
    s.run('analyzeAndRenderMapping()');
    R.check('自動改用多工作表也失敗時：只跳失敗原因，不先說「已自動改用」', s.alerts.length === 1 && !s.alerts[0].includes('已自動改用'), s.alerts.join(' / '));
}
{
    // 部分工作表沒有姓名欄的提醒（對照表上方）
    const s = loadSite();
    await uploadSheets(s, [{ name: '社團-週一桌遊社', rows: STUDENTS }, { name: '名單整理', rows: [['班級', '座號'], ['301', '05']] }]);
    const box = s.el('mapping-skip-alert');
    R.check('部分工作表沒有姓名欄：對照表上方有提醒', !box.hidden && box.innerHTML.includes('名單整理'));
    R.check('部分工作表沒有姓名欄：提醒說明不是社團名單時可不理會、改好後回到第 1 步', box.innerHTML.includes('可以不用理會') && box.innerHTML.includes('回到第 1 步'), box.innerHTML.slice(-120));
}

// ---- 「已套用這台電腦上次記住的星期」
{
    const html = readSiteFile('index.html');
    const tag = html.match(/<div[^>]*class="memory-hint"[^>]*>/)?.[0] || '';
    R.check('index.html：提示預設隱藏（JS 未執行時不誤顯示）', /\bid="memory-hint"/.test(tag) && /\shidden(\s|>|=)/.test(tag), tag);
    R.check('index.html：提示文字放在 #memory-hint-text', /id="memory-hint-text"/.test(html));
    const css = readSiteFile('css/style.css');
    R.check('style.css：.memory-hint[hidden] 會隱藏（display: inline-flex 不會蓋過 hidden）', /\.memory-hint\[hidden\]\s*\{[^}]*display:\s*none/.test(css));
}
{
    const s = loadSite();
    await uploadSheets(s, SINGLE);
    R.check('沒有記住任何設定：提示隱藏', onMapping(s) && hint(s).hidden === true, `hidden=${hint(s).hidden}`);
}
{
    const s = loadSite();
    s.storage.set(KEY, JSON.stringify({ '社團': { '桌遊社': '週三' } }));
    await uploadSheets(s, SINGLE);
    R.check('單一清單：1 個社團套用記住的星期時顯示提示', hint(s).hidden === false, `hidden=${hint(s).hidden}`);
    R.check('單一清單：提示寫出套用了 1 個社團', hintText(s).includes('1 個社團'), hintText(s));
    // 清除後重新讀取沒有記住設定的名冊：提示隱藏
    s.storage.clear();
    s.run("document.getElementById('mode-select').value = 'single'; analyzeAndRenderMapping();");
    R.check('重新渲染、記住的設定已清除：提示隱藏', hint(s).hidden === true);
}
{
    const s = loadSite();
    // 記住的星期被名稱裡的星期覆蓋（單一清單名稱優先）：不算套用
    s.storage.set(KEY, JSON.stringify({ '社團': { '周一桌遊社 (一 10,11)': '週三' } }));
    await uploadSheets(s, [{ name: '總表', rows: [['班級', '座號', '姓名', '社團'], ['301', '05', SECRET, '周一桌遊社 (一 10,11)']] }]);
    const day = rows(s)[0]?.querySelector('.day-select').children.find(o => o.selected)?.value;
    R.check('（前提）名稱裡的星期優先：選的是週一', day === '週一', day);
    R.check('記住的星期被名稱覆蓋：不算套用，提示隱藏', hint(s).hidden === true, `hidden=${hint(s).hidden}，${hintText(s)}`);
}
{
    const s = loadSite();
    // 舊格式（最上層「社團 → 星期」）與時段欄位格式各一個社團
    s.storage.set(KEY, JSON.stringify({ '直笛社': '週五', '社團': { '桌遊社': '週三' } }));
    await uploadSheets(s, SINGLE);
    R.check('單一清單：2 個社團套用時提示寫 2 個社團', hint(s).hidden === false && hintText(s).includes('2 個社團'), hintText(s));
    // 切到多工作表失敗、退回單一清單：提示狀態照舊
    const sel = s.el('mode-select');
    sel.value = 'multi';
    sel.dispatch('change');
    R.check('切換模式失敗退回後：提示仍正確', sel.value === 'single' && hint(s).hidden === false && hintText(s).includes('2 個社團'), hintText(s));
    // 載入失敗的名冊：提示隱藏（不留上一份名冊的狀態）
    await uploadSheets(s, [{ name: '說明', rows: [['這是說明']] }]);
    R.check('換成無法整理的名冊：提示隱藏', hint(s).hidden === true);
}
{
    const s = loadSite();
    // 多工作表：記住的設定優先於工作表名稱
    s.storage.set(KEY, JSON.stringify({ '社團-週一桌遊社': '週四' }));
    await uploadSheets(s, [{ name: '社團-週一桌遊社', rows: STUDENTS }, { name: '社團-週二直笛社', rows: STUDENTS }]);
    const day = rows(s)[0]?.querySelector('.day-select').children.find(o => o.selected)?.value;
    R.check('（前提）多工作表：記住的星期優先，選的是週四', day === '週四', day);
    R.check('多工作表：1 張工作表套用記住的星期時顯示提示並寫出 1 個社團', hint(s).hidden === false && hintText(s).includes('1 個社團'), hintText(s));
    // 切到單一清單（沒有總表，自動改回多工作表）：提示仍顯示
    const sel = s.el('mode-select');
    sel.value = 'single';
    sel.dispatch('change');
    R.check('自動改回多工作表後：提示仍正確', sel.value === 'multi' && hint(s).hidden === false && hintText(s).includes('1 個社團'), hintText(s));
}
{
    const s = loadSite();
    // 單一清單：同一社團出現在 A、B 兩個時段欄位（對照表兩列），都套用記住的星期，只算 1 個社團
    s.storage.set(KEY, JSON.stringify({ 'A時段': { '桌遊社': '週三' }, 'B時段': { '桌遊社': '週四', '直笛社': '週五' } }));
    await uploadSheets(s, [{ name: '總表', rows: [
        ['班級', '座號', '姓名', 'A時段', 'B時段'],
        ['301', '05', SECRET, '桌遊社', '桌遊社'],
        ['302', '11', '測試生302_11', '無', '直笛社'],
    ] }]);
    R.check('（前提）A、B 時段：對照表 3 列（桌遊社 A、桌遊社 B、直笛社 B）', rows(s).length === 3, `${rows(s).length} 列`);
    R.check('A、B 時段的同一社團只算 1 個：提示寫 2 個社團', hint(s).hidden === false && hintText(s).includes('（2 個社團）'), hintText(s));
}
{
    const s = loadSite();
    // 多工作表：A、B 時段各一張工作表的同一社團，只算 1 個
    s.storage.set(KEY, JSON.stringify({ '社團-桌遊社A': '週二', '社團-桌遊社B': '週四' }));
    await uploadSheets(s, [{ name: '社團-桌遊社A', rows: STUDENTS }, { name: '社團-桌遊社B', rows: STUDENTS }]);
    const names = rows(s).map(tr => tr.querySelector('.club-input').value);
    R.check('（前提）多工作表 A、B：對照表 2 列、顯示名稱相同', rows(s).length === 2 && names[0] === names[1], names.join('、'));
    R.check('多工作表 A、B 時段的同一社團只算 1 個', hint(s).hidden === false && hintText(s).includes('（1 個社團）'), hintText(s));
}
{
    const s = loadSite();
    // 記住的星期不是有效選項（例如損毀成「請選擇」）：不算套用
    s.storage.set(KEY, JSON.stringify({ '社團-週一桌遊社': '請選擇' }));
    await uploadSheets(s, [{ name: '社團-週一桌遊社', rows: STUDENTS }]);
    R.check('多工作表：記住的值無效時不算套用，提示隱藏', hint(s).hidden === true);
}
R.done();
