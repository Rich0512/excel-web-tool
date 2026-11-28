// 貼上名單的解析預覽：防抖（連續輸入只解析一次）、按「加入這個社團」時用最新內容、清空時取消待執行的預覽。
// 測試資料的姓名全部是虛構的。
import { loadSite, makeReporter } from './harness.mjs';
const R = makeReporter('貼上名單：解析預覽防抖');
const TEXT = '週三桌遊社\n班級\t座號\t姓名\n101\t05\t測試生101_05\n102\t11\t測試生102_11';

// 計算 parsePastedText 被呼叫幾次
function withCounter(s) {
    const orig = s.ctx.parsePastedText;
    const counter = { n: 0 };
    s.ctx.parsePastedText = function (...a) { counter.n++; return orig.apply(this, a); };
    return counter;
}
const previewShown = (s) => s.el('paste-parse-preview').style.display === 'block';

// 1. 連續輸入只解析一次；paste 不重複解析
{
    const s = loadSite();
    const c = withCounter(s);
    const area = s.el('paste-text-area');
    for (let i = 1; i <= 5; i++) {
        area.value = TEXT.slice(0, TEXT.length - 5 + i);
        area.dispatch('input');
        s.advance(50);
    }
    area.value = TEXT;
    area.dispatch('paste');
    area.dispatch('input'); // 瀏覽器在 paste 之後一定會觸發 input
    s.advance(10);
    R.check('輸入停止前不顯示預覽（防抖中）', !previewShown(s));
    s.advance(1000);
    R.check('連續 6 次 input＋1 次 paste 只解析 1 次', c.n === 1, `解析 ${c.n} 次`);
    R.check('停止輸入後顯示預覽', previewShown(s));
}
// 2. 按「加入這個社團」時使用最新內容（防抖還沒跑完也一樣）
{
    const s = loadSite();
    const area = s.el('paste-text-area');
    area.value = TEXT;
    area.dispatch('input');
    s.el('btn-load-club').dispatch('click'); // 不推進時鐘
    const count = s.run('pastedClubs.length');
    R.check('立刻按加入：成功加入且社團名稱已自動帶入', count === 1 && s.run('pastedClubs[0].clubName') === '桌遊社', `加入 ${count} 個；${s.alerts.join('/')}`);
    R.check('立刻按加入：學生數依最新內容', count === 1 && s.run('pastedClubs[0].students.length') === 2);
    const c = withCounter(s);
    s.advance(1000);
    R.check('加入後不會再跑舊內容的預覽', c.n === 0 && !previewShown(s), `解析 ${c.n} 次`);
}
// 3. 清空文字框時取消待執行的預覽
{
    const s = loadSite();
    const c = withCounter(s);
    const area = s.el('paste-text-area');
    area.value = TEXT;
    area.dispatch('input');
    s.advance(100);
    area.value = '';
    area.dispatch('input');
    R.check('清空後立刻隱藏預覽', s.el('paste-parse-preview').style.display === 'none');
    s.advance(1000);
    R.check('清空後不會再解析舊內容', c.n === 0, `解析 ${c.n} 次`);
    R.check('清空後預覽保持隱藏', s.el('paste-parse-preview').style.display === 'none');
}
// 4. 只靠 input 事件，沒有額外的 paste 監聽（避免重複解析）
{
    const s = loadSite();
    const n = (s.el('paste-text-area').listeners.paste || []).length;
    R.check('貼上框沒有額外的 paste 監聽', n === 0, `${n} 個`);
}
// 5. 預覽的「有班級」人數和抽籤、課表的新生判斷（isFreshmanClass）一致：
//    班級寫「新生」「新生班」或空白都算新生，不計入有班級
{
    const s = loadSite();
    const area = s.el('paste-text-area');
    area.value = '班級\t座號\t姓名\n301\t05\t測試生301_05\n新生班\t01\t測試生新生_01\n新生\t02\t測試生新生_02\n\t\t測試生新生_03';
    area.dispatch('input');
    s.advance(1000);
    const html = s.el('paste-parse-preview').innerHTML;
    const total = (html.match(/讀到 <strong>(\d+)<\/strong> 位學生/) || [])[1];
    const withClass = (html.match(/其中 (\d+) 位有班級/) || [])[1];
    R.check('預覽讀到 4 位學生（含 3 位新生）', total === '4', `讀到 ${total} 位`);
    R.check('預覽的「有班級」不計入班級為「新生班」「新生」或空白的學生', withClass === '1', `有班級 ${withClass} 位`);
}
R.done();
