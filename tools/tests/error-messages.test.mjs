// 錯誤訊息不可洩漏名冊內容：例外訊息可能含儲存格值，alert 只顯示固定說明，例外本身只記在 console.error。
// 這裡用虛構的「測試生301_05」當作會被串進例外訊息的名冊內容。
import { loadSite, makeReporter, makeFile, makeWorkbook } from './harness.mjs';
const R = makeReporter('錯誤訊息：alert 不含名冊內容');
const SECRET = '測試生301_05';
const leaked = (s) => s.alerts.some(a => a.includes(SECRET));

{
    const s = loadSite();
    s.ctx.__loadError = new Error(`bad cell ${SECRET}`);
    await s.run('handleFileSelect')(makeFile('名冊.xlsx'));
    await s.settle();
    R.check('讀取名冊失敗：只有 1 則 alert 且不含例外內容', s.alerts.length === 1 && !leaked(s), `${s.alerts.length} 則，含名冊內容：${leaked(s)}`);
    R.check('讀取名冊失敗：例外完整記錄在 console.error', s.logs.some(([lv, a]) => lv === 'error' && a.some(x => x instanceof Error || (x && x.message))));
}
{
    const s = loadSite();
    s.ctx.__wb = makeWorkbook([{ name: '總表', rows: [['班級', '座號', '姓名', '社團'], ['301', '05', SECRET, '桌遊社']] }]);
    s.run("uploadedWorkbook = __wb; document.getElementById('mode-select').value = 'single'; analyzeAndRenderMapping();");
    s.el('mapping-table-body').querySelectorAll('tr').forEach(tr => { tr.querySelector('.day-select').value = '週一'; });
    s.ctx.triggerFileDownload = () => { throw new Error(`write failed ${SECRET}`); };
    await s.run('processAndDownload()');
    R.check('上傳模式匯出失敗：只有 1 則 alert 且不含例外內容', s.alerts.length === 1 && !leaked(s), `${s.alerts.length} 則，含名冊內容：${leaked(s)}`);
}
{
    const s = loadSite();
    s.run("pastedClubs = [{ id: 1, day: '週一', clubName: '桌遊社', quota: 1, gradeStart: '1', gradeEnd: '6', priority: 'desc', students: [{ class: '301', seat: '05', name: '" + SECRET + "', selected: true, drawSequence: 1 }], selectedCount: 1 }]");
    s.ctx.triggerFileDownload = () => { throw new Error(`write failed ${SECRET}`); };
    await s.run('processPastedData()');
    R.check('貼上模式匯出失敗：只有 1 則 alert 且不含例外內容', s.alerts.length === 1 && !leaked(s), `${s.alerts.length} 則，含名冊內容：${leaked(s)}`);
}
{
    const s = loadSite();
    s.fireWindow('unhandledrejection', { reason: new Error(`oops ${SECRET}`) });
    R.check('未處理的非同步例外：只有 1 則 alert 且不含例外內容', s.alerts.length === 1 && !leaked(s), `${s.alerts.length} 則，含名冊內容：${leaked(s)}`);
}
R.done();
