// 貼上名單時自動填入社團名稱與星期（介面行為）：
// 只在欄位空白或仍是自動填入的值時才填，絕不覆蓋使用者自己填的值；加入社團、清除資料、清空貼上框時重設。
// 測試資料的姓名全部是虛構的。
import { loadSite, makeReporter, paste } from './harness.mjs';
const R = makeReporter('貼上名單：自動填入社團名稱與星期');

const HEADER_FIRST = '班級\t座號\t姓名\n301\t05\t測試生301_05\n302\t11\t測試生302_11';
const clubOf = (s) => s.el('paste-club-name').value;

// 1. 照介面提示從標題列開始貼：不可以把資料列當成社團名稱
{
    const s = loadSite();
    paste(s, HEADER_FIRST);
    R.check('從標題列開始貼，社團名稱維持空白', clubOf(s) === '', `社團名稱欄長度=${clubOf(s).length}`);
}
// 2. 使用者已輸入的社團名稱不可被蓋掉（介面流程）
{
    const s = loadSite();
    s.el('paste-club-name').value = '舞蹈社';
    paste(s, '直笛社錄取名單\n' + HEADER_FIRST);
    R.check('使用者自己打的社團名稱，貼上有標題的名單後不被覆蓋', clubOf(s) === '舞蹈社', clubOf(s));
}
// 3. 直接呼叫 guessClubAndDayFromPastedText：使用者填的名稱（沒有自動填入記號）一律保留
{
    const s = loadSite();
    const guessFill = s.fn('guessClubAndDayFromPastedText');
    const USER = '使用者自己打的社團';
    for (const [title, text] of [
        ['從標題列開始貼', HEADER_FIRST],
        ['有社團標題「周三舞蹈社 錄取名單」', '周三舞蹈社 錄取名單\n' + HEADER_FIRST],
        ['沒有標題列、兩字姓名', '301\t05\t王明\n302\t11\t陳華'],
    ]) {
        const club = s.el('paste-club-name');
        club.value = USER;
        club.dataset = {};
        R.test(`使用者已填社團名稱時不覆蓋：${title}`, () => {
            guessFill(text);
            return [club.value === USER, `社團名稱欄長度=${club.value.length}`];
        });
    }
}
// 4. 標題列之前的行可以推測社團與星期（星期字樣不放進社團名稱）
{
    const s = loadSite();
    paste(s, '週三直笛社錄取名單\n' + HEADER_FIRST);
    R.check('從標題列前一行推測社團名稱', clubOf(s) === '直笛社', clubOf(s));
    R.check('從標題列前一行推測星期', s.el('paste-day-select').value === '週三', s.el('paste-day-select').value);
    // 仍是自動填入的值時，改貼另一份會跟著更新
    paste(s, '週四桌遊社\n' + HEADER_FIRST);
    R.check('自動填入的社團名稱會隨新內容更新', clubOf(s) === '桌遊社', clubOf(s));
    R.check('自動填入的星期會隨新內容更新', s.el('paste-day-select').value === '週四', s.el('paste-day-select').value);
    paste(s, HEADER_FIRST);
    R.check('新內容推測不出名稱時，清掉上一份自動填入的名稱', clubOf(s) === '', `長度=${clubOf(s).length}`);
}
// 5. 使用者已選的星期不可被蓋掉
{
    const s = loadSite();
    const day = s.el('paste-day-select');
    day.value = '週五';
    day.dispatch('change');
    paste(s, '週二桌遊社\n' + HEADER_FIRST);
    R.check('使用者選的星期不被覆蓋', day.value === '週五', day.value);
    R.check('社團名稱欄空白時仍會自動填入', clubOf(s) === '桌遊社', clubOf(s));
}
// 6. 沒有標題列：資料列一律跳過
{
    const s = loadSite();
    paste(s, '301\t05\t測試生301_05\n302\t11\t測試生302_11');
    R.check('沒有標題列時，Tab 分隔的資料列不當成社團名稱', clubOf(s) === '', `長度=${clubOf(s).length}`);
    paste(s, '301 05 測試生\n302 11 測試生');
    R.check('沒有標題列時，空白分隔的資料列不當成社團名稱', clubOf(s) === '', `長度=${clubOf(s).length}`);
}
// 7. 加入社團後重設記號：上一個社團手動選的星期，不會擋住下一個社團的自動推測
{
    const s = loadSite();
    const day = s.el('paste-day-select');
    day.value = '週五';
    day.dispatch('change');
    s.el('paste-club-name').value = '舞蹈社';
    paste(s, HEADER_FIRST);
    s.el('btn-load-club').dispatch('click');
    s.advance(1000);
    const added = s.run('pastedClubs.length');
    R.check('（前提）加入這個社團成功', added === 1, `已加入 ${added} 個`);
    paste(s, '週二桌遊社\n' + HEADER_FIRST);
    R.check('加入社團後，下一份名單會重新自動推測星期', day.value === '週二', day.value);
    R.check('加入社團後，下一份名單會重新自動推測社團', clubOf(s) === '桌遊社', clubOf(s));
}
// 8. 清除這次的資料後也重設
{
    const s = loadSite();
    const day = s.el('paste-day-select');
    day.value = '週五';
    day.dispatch('change');
    s.el('btn-hard-reset').dispatch('click');
    paste(s, '週二桌遊社\n' + HEADER_FIRST);
    R.check('清除資料後，星期重新自動推測', day.value === '週二', day.value);
}
// 9. 清空貼上框：清掉自動填入的值，但不動使用者自己填的值
{
    const s = loadSite();
    paste(s, '週三桌遊社\n班級\t座號\t姓名\n301\t05\t測試生301_05');
    R.check('（前提）自動填入社團與星期', clubOf(s) === '桌遊社' && s.el('paste-day-select').value === '週三');
    paste(s, '');
    R.check('清空貼上框：清掉自動填入的社團名稱', clubOf(s) === '', clubOf(s));
    R.check('清空貼上框：星期還原成預設週一', s.el('paste-day-select').value === '週一', s.el('paste-day-select').value);
    R.check('清空貼上框：刪除自動填入記號', s.el('paste-club-name').dataset.autoFilled === undefined && s.el('paste-day-select').dataset.autoFilled === undefined);
}
{
    const s = loadSite();
    s.el('paste-club-name').value = '舞蹈社';
    const day = s.el('paste-day-select');
    day.value = '週五'; day.dispatch('change');
    paste(s, '週三桌遊社\n班級\t座號\t姓名\n301\t05\t測試生301_05');
    paste(s, '');
    R.check('清空貼上框：使用者自己填的社團名稱與星期不動', clubOf(s) === '舞蹈社' && day.value === '週五', `${clubOf(s)}／${day.value}`);
}
R.done();
