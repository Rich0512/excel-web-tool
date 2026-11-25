// 其他小功能：escapeHtml 對 0／false 的處理、匯出檔名、貼上模式的正取人數驗證。
// 測試資料的姓名全部是虛構的。
import { loadSite, makeReporter } from './harness.mjs';
const R = makeReporter('其他：escapeHtml、匯出檔名、正取人數');

// ---- escapeHtml
{
    const esc = loadSite().fn('escapeHtml');
    R.test('escapeHtml(0) 是 "0"（不可當成空值）', () => [esc(0) === '0', JSON.stringify(esc(0))]);
    R.test('escapeHtml(false) 是 "false"', () => [esc(false) === 'false', JSON.stringify(esc(false))]);
    R.test('escapeHtml(null)、escapeHtml(undefined) 是空字串', () => esc(null) === '' && esc(undefined) === '');
    R.test('escapeHtml("") 是空字串', () => esc('') === '');
    R.test('escapeHtml 跳脫 < > " \' &', () => esc('<a href="x">\'&') === '&lt;a href=&quot;x&quot;&gt;&#039;&amp;');
}

// ---- 匯出檔名：去掉原檔的試算表副檔名，加上「_學期社團彙整.xlsx」
async function exportName(base, slotMode = 'semester') {
    const s = loadSite();
    const data = [{ class: '301', seat: '05', name: '測試生301_05', schedule: { '週一': '桌遊社' }, remarks: [] }];
    await s.run('exportWeeklySchedule')(data, new Set(['週一']), slotMode, base);
    const a = s.appended.map(x => x.child).find(el => el.tagName === 'A');
    return a?.download;
}
for (const [base, expect] of [
    ['115.2學期', '115.2學期_學期社團彙整.xlsx'],
    ['名冊.xlsx', '名冊_學期社團彙整.xlsx'],
    ['名冊.XLSX', '名冊_學期社團彙整.xlsx'],
    ['名冊.xls', '名冊_學期社團彙整.xlsx'],
    ['v1.0 名冊.xlsx', 'v1.0 名冊_學期社團彙整.xlsx'],
    ['直接貼上名單彙整', '直接貼上名單彙整_學期社團彙整.xlsx'],
]) {
    const got = await exportName(base).catch(e => `例外：${e.message}`);
    R.check(`匯出檔名「${base}」→「${expect}」`, got === expect, got);
}

// ---- 正取人數必須是 1～500 的整數
const TEXT = '班級\t座號\t姓名\n101\t05\t測試生101_05\n102\t11\t測試生102_11';
function tryAdd(quota) {
    const s = loadSite();
    s.el('paste-club-name').value = '桌遊社';
    s.el('paste-text-area').value = TEXT;
    s.el('paste-quota').value = quota;
    s.el('btn-load-club').dispatch('click');
    s.advance(1000);
    return { count: s.run('pastedClubs.length'), quota: s.run('pastedClubs[0] && pastedClubs[0].quota'), alerts: s.alerts };
}
for (const bad of ['-3', '0', '2.5', '', '1e1', '501']) {
    const r = tryAdd(bad);
    R.check(`正取人數「${bad || '空白'}」被拒絕並說明`, r.count === 0 && r.alerts.some(a => /正取人數/.test(a)), `加入 ${r.count} 個，quota=${r.quota}`);
}
for (const [ok, expect] of [['5', 5], [' 12 ', 12], ['500', 500]]) {
    const r = tryAdd(ok);
    R.check(`正取人數「${ok}」可以加入`, r.count === 1 && r.quota === expect, `加入 ${r.count} 個，quota=${r.quota}`);
}
R.done();
