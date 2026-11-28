// ============================================
// 貼上名單的招收年級選單與下載時「沒有可以彙整的學生」的訊息：
// - #paste-grade-start／#paste-grade-end 的選項順序（新生在最前面）與預設值（1 年級至 2 年級）
// - 「新生」至「2 年級」可以加入並抽籤；起始年級大於結束年級仍會擋下，訊息通順
// - 已加入社團那一行的招收年級文字（限新生與 1–2 年級…）
// - 下載時沒有學生：取消勾選新生造成的，說勾回來再下載；其他原因說明要移除社團重新加入（會重新抽籤）
// - 下載不重抽、不計抽籤次數；勾回新生後再按一次下載就會匯出
// - 上傳 Excel 模式沒有學生時的訊息（不再叫使用者檢查已經擋過的星期）
// 測試資料全部是虛構的（測試生301_05 等）。
// ============================================

import { loadSite, makeReporter, readSiteFile, stubConfirm, uploadSheets } from './harness.mjs';

const R = makeReporter('招收年級選單與下載時沒有學生的訊息');

const GRADE_VALUES = ['新生', '1', '2', '3', '4', '5', '6'];
const FRESHMEN_MSG = '已加入的社團中，正取的學生都沒有班級座號（新生），目前沒有勾選「包含沒有班級座號的新生」，所以課表沒有學生。' +
    '請勾選「包含沒有班級座號的新生」，再按一次「下載課表與抽籤結果」。已加入的社團與抽籤結果都還在，不需要重新加入，也不會重新抽籤。';
const NO_SELECTED_MSG = '已加入的社團都沒有正取的學生，沒有課表可以下載。通常是報名的學生都不在招收年級內（名單中標示為「不符限制」），或班級無法辨識。' +
    '請點社團那一列查看名單確認。如果是招收年級選錯了，要先移除那個社團，改好招收年級後重新貼上名單並加入；重新加入會重新抽籤，抽籤次數會累加。';
const RANGE_MSG = (a, b) => `起始年級（${a}）大於結束年級（${b}）。「新生」是最低的年級，請重新選擇招收年級，讓起始年級低於或等於結束年級，再按「加入這個社團」。`;
const EXCEL_FRESHMEN_MSG = '名冊中的學生都沒有班級座號（新生），目前沒有勾選「包含沒有班級座號的新生」，所以沒有學生可以匯出。' +
    '請勾選「包含沒有班級座號的新生」，再按一次「開始轉換並下載課表」，星期設定都還在。';
// 訊息不可含名冊內容
const leaks = (msgs) => msgs.some(m => /測試生|\b[1-6]0\d\b/.test(m));

const safe = (s, code) => { try { return s.run(code); } catch (e) { return undefined; } };
// 沒有班級座號的新生：姓名欄放最前面（parsePastedText 會略過行首的空白欄位）
const roster = (...rows) => ['姓名\t班級\t座號', ...rows.map(([cls, seat, name]) => `${name}\t${cls}\t${seat}`)].join('\n');

function addClub(s, text, { quota = '10', start = '1', end = '6', priority = 'desc', name = '桌遊社' } = {}) {
    s.el('paste-club-name').value = name;
    s.el('paste-text-area').value = text;
    s.el('paste-quota').value = quota;
    s.el('paste-grade-start').value = start;
    s.el('paste-grade-end').value = end;
    s.el('paste-priority').value = priority;
    s.el('btn-load-club').dispatch('click');
    s.advance(1000);
}
// 社團那一列的 HTML（.loaded-club-meta 是 .loaded-club-info 的 innerHTML 字串的一部分）
const meta = (s, i = 0) => (s.el('loaded-clubs-list').descendants().filter(el => el.classList.contains('loaded-club-info'))[i]?.innerHTML || '');

// 換掉匯出與抽籤：記下呼叫次數，確認下載不會重抽
function instrument(s) {
    s.ctx.__exports = [];
    s.ctx.__lotteryCalls = 0;
    s.run('exportWeeklySchedule = async (...a) => { __exports.push(a.length); }');
    s.run('const __origLottery = runClubLottery; runClubLottery = function (...a) { __lotteryCalls++; return __origLottery.apply(this, a); }');
    return { exports: () => s.ctx.__exports.length, lottery: () => s.ctx.__lotteryCalls };
}

// ---- 1. 選項順序與預設值（index.html）
{
    const html = readSiteFile('index.html');
    for (const id of ['paste-grade-start', 'paste-grade-end']) {
        const sel = (html.match(new RegExp(`<select id="${id}"[\\s\\S]*?</select>`)) || [''])[0];
        const opts = [...sel.matchAll(/<option value="([^"]+)"([^>]*)>([^<]+)<\/option>/g)];
        const values = opts.map(m => m[1]);
        R.check(`#${id} 的選項順序是 新生、1…6 年級`, JSON.stringify(values) === JSON.stringify(GRADE_VALUES), values.join(','));
        const selected = opts.filter(m => /\bselected\b/.test(m[2])).map(m => m[1]);
        const want = id === 'paste-grade-start' ? '1' : '2';
        R.check(`#${id} 的預設值仍是 ${want} 年級（只有一個 selected）`, selected.length === 1 && selected[0] === want, selected.join(',') || '（沒有 selected）');
        const s = loadSite();
        const f = s.fn('gradeOptionLabel');
        const bad = opts.filter(m => f(m[1]) !== m[3].trim());
        R.check(`#${id} 的選項文字與 gradeOptionLabel 相同`, opts.length === 7 && bad.length === 0, bad.map(m => m[1]).join(','));
    }
}

// ---- 1b. 兩種模式的新生勾選框用同一個名稱（DESIGN.md：同一個動作從頭到尾用同一個名稱）
{
    const html = readSiteFile('index.html');
    const label = (id) => (html.match(new RegExp(`id="${id}"[^>]*>\\s*<span class="checkbox-text">([^<]+)</span>`)) || ['', ''])[1];
    R.check('上傳 Excel 與貼上名單的新生勾選框都叫「包含沒有班級座號的新生」',
        label('include-freshmen') === '包含沒有班級座號的新生' && label('paste-include-freshmen') === '包含沒有班級座號的新生',
        `Excel=${label('include-freshmen')}，貼上=${label('paste-include-freshmen')}`);
}

// ---- 2. 新生至 2 年級：可以加入並抽籤，新生與 1、2 年級參加抽籤
{
    const s = loadSite();
    const ROSTER = roster(['', '', '測試生新生_01'], ['101', '02', '測試生101_02'], ['201', '03', '測試生201_03'], ['301', '04', '測試生301_04']);
    addClub(s, ROSTER, { quota: '10', start: '新生', end: '2' });
    const club = safe(s, 'pastedClubs[0]');
    R.check('新生至 2 年級：加入社團，沒有跳出錯誤', !!club && s.alerts.length === 0, s.alerts.join(' / '));
    const tags = club ? club.students.map(x => `${x.class || '空白'}:${x.selected ? '正取' : x.drawSequence}`) : [];
    R.check('新生至 2 年級：新生、101、201 正取，301 不符限制', club && club.selectedCount === 3 &&
        club.students.find(x => x.name === '測試生301_04')?.drawSequence === '不符限制' &&
        club.students.find(x => x.name === '測試生新生_01')?.selected === true, tags.join(','));
    R.check('新生至 2 年級：抽籤次數記為第 1 次', club && club.drawRound === 1 && s.run('clubDrawCounts.size') === 1, club ? `drawRound=${club.drawRound}` : '');
    R.check('新生至 2 年級：已加入社團那一行寫「限新生與 1–2 年級」', meta(s).includes('，限新生與 1–2 年級') && !meta(s).includes('新生–'), meta(s).replace(/測試生\S*/g, ''));
}

// ---- 3. 已加入社團那一行的招收年級文字
{
    const s = loadSite();
    const f = s.fn('gradeRangeText');
    const cases = [
        [['1', '2'], '限 1–2 年級'],
        [['3', '6'], '限 3–6 年級'],
        [['4', '4'], '限 4 年級'],
        [['新生', '1'], '限新生與 1 年級'],
        [['新生', '6'], '限新生與 1–6 年級'],
        [['新生', '新生'], '限新生'],
    ];
    for (const [[a, b], want] of cases) {
        R.test(`招收年級文字：${a} 至 ${b} → ${want}`, () => { const got = f(a, b); return [got === want, got]; });
    }
}

// ---- 4. 起始年級大於結束年級仍會擋下（1 年級至新生、3 年級至 2 年級）
for (const [start, end, a, b] of [['1', '新生', '1 年級', '新生'], ['3', '2', '3 年級', '2 年級']]) {
    const s = loadSite();
    addClub(s, '班級\t座號\t姓名\n301\t05\t測試生301_05', { start, end });
    R.check(`${a}至${b}：不加入社團、不計抽籤次數`, s.run('pastedClubs.length') === 0 && s.run('clubDrawCounts.size') === 0,
        `社團 ${s.run('pastedClubs.length')} 個`);
    R.check(`${a}至${b}：訊息寫出兩個年級的完整名稱並說明新生最低`, s.alerts.length === 1 && s.alerts[0] === RANGE_MSG(a, b), s.alerts.join(' / '));
}
{
    const s = loadSite();
    addClub(s, roster(['', '', '測試生新生_01']), { start: '新生', end: '新生' });
    R.check('新生至新生：可以加入（起始等於結束）', s.run('pastedClubs.length') === 1 && s.alerts.length === 0 && meta(s).includes('，限新生'), s.alerts.join(' / '));
}

// ---- 5. 下載：取消勾選新生導致沒有學生 → 勾回來再下載，不重抽、不計次
{
    const s = loadSite();
    const confirmStub = stubConfirm(s);
    addClub(s, roster(['', '', '測試生新生_01'], ['', '', '測試生新生_02']), { quota: '1', start: '新生', end: '6', name: '桌遊社' });
    const before = safe(s, 'JSON.stringify(pastedClubs[0].students.map(x => [x.name, x.selected, x.drawSequence]))');
    const t = instrument(s);
    s.el('paste-include-freshmen').checked = false;
    s.alerts.length = 0;
    await s.run('processPastedData()');
    R.check('只有新生正取、取消勾選新生：跳出「勾選後再按一次下載」的訊息', s.alerts.length === 1 && s.alerts[0] === FRESHMEN_MSG, s.alerts.join(' / '));
    R.check('只有新生正取、取消勾選新生：訊息不含名冊內容', !leaks(s.alerts));
    R.check('只有新生正取、取消勾選新生：沒有匯出，已加入的社團與抽籤結果都還在',
        t.exports() === 0 && s.run('pastedClubs.length') === 1 &&
        safe(s, 'JSON.stringify(pastedClubs[0].students.map(x => [x.name, x.selected, x.drawSequence]))') === before,
        `匯出 ${t.exports()} 次，社團 ${s.run('pastedClubs.length')} 個`);
    R.check('只有新生正取、取消勾選新生：下載沒有重新抽籤、抽籤次數不變',
        t.lottery() === 0 && s.run('clubDrawCounts.get(clubDrawKey("桌遊社"))') === 1 && s.run('pastedClubs[0].drawRound') === 1,
        `抽籤 ${t.lottery()} 次，次數 ${s.run('clubDrawCounts.get(clubDrawKey("桌遊社"))')}`);
    // 照訊息做：勾回來再按一次下載
    s.el('paste-include-freshmen').checked = true;
    s.alerts.length = 0;
    await s.run('processPastedData()');
    R.check('勾回新生後再按一次下載：直接匯出，沒有跳出訊息', t.exports() === 1 && s.alerts.length === 0, s.alerts.join(' / '));
    R.check('勾回新生後再按一次下載：沒有重新抽籤、沒有確認重抽、抽籤次數仍是 1',
        t.lottery() === 0 && confirmStub.messages.length === 0 && s.run('clubDrawCounts.get(clubDrawKey("桌遊社"))') === 1,
        `抽籤 ${t.lottery()} 次，confirm ${confirmStub.messages.length} 則`);
}

// ---- 6. 下載：沒有任何正取的學生（都不符限制）→ 說明原因，要改招收年級才移除重新加入
for (const includeFreshmen of [true, false]) {
    const s = loadSite();
    addClub(s, '班級\t座號\t姓名\n301\t05\t測試生301_05\n401\t06\t測試生401_06', { start: '1', end: '2' });
    const t = instrument(s);
    s.el('paste-include-freshmen').checked = includeFreshmen;
    s.alerts.length = 0;
    await s.run('processPastedData()');
    const label = includeFreshmen ? '勾選新生' : '取消勾選新生但名單沒有新生';
    R.check(`都不符限制（${label}）：訊息說明原因與移除重新加入會重新抽籤`, s.alerts.length === 1 && s.alerts[0] === NO_SELECTED_MSG, s.alerts.join(' / '));
    R.check(`都不符限制（${label}）：訊息不含名冊內容、沒有匯出、沒有重抽`, !leaks(s.alerts) && t.exports() === 0 && t.lottery() === 0 && s.run('pastedClubs.length') === 1);
}

// ---- 7. 下載：正取有新生也有一般學生，取消勾選新生 → 照常下載（只少了新生）
{
    const s = loadSite();
    addClub(s, roster(['', '', '測試生新生_01'], ['301', '05', '測試生301_05']), { start: '新生', end: '6' });
    const t = instrument(s);
    s.el('paste-include-freshmen').checked = false;
    s.alerts.length = 0;
    await s.run('processPastedData()');
    R.check('正取有一般學生、取消勾選新生：照常匯出、不跳訊息', t.exports() === 1 && s.alerts.length === 0, s.alerts.join(' / '));
}

// ---- 8. 清除這次的資料：已加入社團時提醒抽籤結果會清除、重新加入會重抽
{
    const s = loadSite();
    addClub(s, '班級\t座號\t姓名\n301\t05\t測試生301_05', { start: '1', end: '6' });
    const c = stubConfirm(s, false);
    s.run('manualHardResetAll()');
    R.check('清除這次的資料（已加入社團）：確認訊息提醒抽籤結果會清除、重新加入會重新抽籤',
        c.messages.length === 1 && /抽籤結果會一起清除/.test(c.messages[0]) && /重新抽籤，抽籤次數會累加/.test(c.messages[0]), c.messages.join(' / '));
    R.check('清除這次的資料：按取消時社團還在', s.run('pastedClubs.length') === 1);
    const s2 = loadSite();
    const c2 = stubConfirm(s2, false);
    s2.run('manualHardResetAll()');
    R.check('清除這次的資料（沒有社團）：不提抽籤', c2.messages.length === 1 && !/抽籤/.test(c2.messages[0]), c2.messages.join(' / '));
    R.check('清除這次的資料：不動招收年級選單（仍是預設 1 至 2 年級）', s.el('paste-grade-start').value === '1' && s.el('paste-grade-end').value === '6' &&
        s2.el('paste-grade-start').value === '1' && s2.el('paste-grade-end').value === '2');
}

// ---- 9. 上傳 Excel：取消勾選新生導致沒有學生 → 勾選後再按一次轉換
{
    const s = loadSite();
    await uploadSheets(s, [{ name: '社團-週一桌遊社', rows: [['班級', '座號', '姓名'], ['', '', '測試生新生_01'], ['', '', '測試生新生_02']] }]);
    s.el('mapping-table-body').querySelectorAll('tr').forEach(tr => { tr.querySelector('.day-select').value = '週一'; });
    const t = instrument(s);
    s.el('include-freshmen').checked = false;
    s.alerts.length = 0;
    await s.run('processAndDownload()');
    R.check('Excel 模式只有新生、取消勾選新生：訊息說勾選後再按一次轉換', s.alerts.length === 1 && s.alerts[0] === EXCEL_FRESHMEN_MSG, s.alerts.join(' / '));
    R.check('Excel 模式只有新生、取消勾選新生：訊息不含名冊內容、沒有匯出', !leaks(s.alerts) && t.exports() === 0);
    s.el('include-freshmen').checked = true;
    s.alerts.length = 0;
    await s.run('processAndDownload()');
    R.check('Excel 模式勾回新生後再按一次：照常匯出', t.exports() === 1 && s.alerts.length === 0, s.alerts.join(' / '));
}

R.done();
