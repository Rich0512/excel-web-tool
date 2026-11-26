// ============================================
// 已加入社團的「移除」按鈕（app.js renderLoadedClubs() 的 .loaded-club-delete）：
// 先跳 confirm 確認、取消時不動、確定時只移除那一個社團並保留其他社團的抽籤結果、
// 不重設抽籤次數（clubDrawCounts），以及移除後的焦點位置。
// 測試資料全部是虛構的（測試生301_05 等）。
// ============================================

import { loadSite, makeReporter, stubConfirm } from './harness.mjs';

const R = makeReporter('移除已加入的社團');

const REMOVE_MSG = '確定要移除這個社團嗎？抽籤結果會一起清除。重新加入時會重新抽籤，抽籤次數會累加。';
const REDRAW_MSG = (n) => `這個社團在這次開啟網頁後已經抽過 ${n} 次籤。再加入會重新抽籤，匯出檔的抽籤資訊會記錄為第 ${n + 1} 次抽籤。確定要重新抽籤嗎？`;
const NAMES = ['桌遊社', '圍棋社', '籃球社'];
// 每個社團用不同的學生，名額 1 人，抽籤結果才有隨機性，可以確認移除後其他社團的結果沒有被重抽
const ROSTERS = [
    '班級\t座號\t姓名\n301\t05\t測試生301_05\n302\t11\t測試生302_11\n303\t08\t測試生303_08',
    '401\t03\t測試生401_03\n402\t12\t測試生402_12\n403\t21\t測試生403_21',
    '501\t07\t測試生501_07\n502\t14\t測試生502_14\n503\t19\t測試生503_19',
].map((t, i) => (i === 0 ? t : `班級\t座號\t姓名\n${t}`));

function addClub(s, text, name) {
    s.el('paste-club-name').value = name;
    s.el('paste-text-area').value = text;
    s.el('paste-quota').value = '1';
    s.el('paste-grade-start').value = '1';
    s.el('paste-grade-end').value = '6';
    s.el('paste-priority').value = 'desc';
    s.el('btn-load-club').dispatch('click');
    s.advance(1000);
}

const safe = (s, code) => { try { return s.run(code); } catch (e) { return undefined; } };
const clubs = (s) => s.run('pastedClubs');
const clubNames = (s) => clubs(s).map(c => c.clubName).join(',');
const buttons = (s) => s.el('loaded-clubs-list').descendants().filter(el => el.classList.contains('loaded-club-delete'));
const items = (s) => s.el('loaded-clubs-list').descendants().filter(el => el.classList.contains('loaded-club-item'));
// 按第 i 個「移除」；按鈕不存在時回傳 false（讓檢查記成 FAIL，不中斷）
const removeAt = (s, i, props = {}) => {
    const b = buttons(s)[i];
    if (!b) return false;
    b.dispatch('click', props);
    return true;
};
// 抽籤結果的快照：學生、是否正取、備取序號、抽籤次數、抽籤時間
const snapshot = (club) => JSON.stringify({
    name: club.clubName,
    students: club.students.map(x => [x.name, x.class, x.seat, !!x.selected, x.drawSequence]),
    selectedCount: club.selectedCount,
    drawRound: club.drawRound,
    drawnAt: club.drawnAt && typeof club.drawnAt.getTime === 'function' ? club.drawnAt.getTime() : null,
});
const removeMsgs = (c) => c.messages.filter(m => m === REMOVE_MSG);
const otherMsgs = (c) => c.messages.filter(m => m !== REMOVE_MSG);
const focusLabel = (s) => {
    const a = s.document.activeElement;
    if (!a) return '（沒有焦點）';
    if (a.id) return `#${a.id}`;
    return `${a.tagName}.${a.className}「${a.getAttribute('aria-label') || a.textContent}」`;
};

// 建立一個已加入 3 個社團的頁面
function setup3() {
    const s = loadSite();
    const c = stubConfirm(s);
    NAMES.forEach((n, i) => addClub(s, ROSTERS[i], n));
    s.document.activeElement = null;
    return { s, c };
}

// ---- 1. 前置：3 個社團都加入，各有一個「移除」按鈕
{
    const { s, c } = setup3();
    R.check('前置：加入 3 個社團，清單有 3 個「移除」按鈕，加入時沒有跳 confirm',
        clubs(s).length === 3 && clubNames(s) === NAMES.join(',') && buttons(s).length === 3 && c.messages.length === 0 && s.alerts.length === 0,
        `社團：${clubNames(s)}，按鈕 ${buttons(s).length} 個，confirm ${c.messages.length} 則，提醒 ${s.alerts.length} 則`);
}

// ---- 2. 按「移除」先跳 confirm；按取消時什麼都不變
{
    const { s, c } = setup3();
    c.answer = false;
    const before = clubs(s).map(snapshot);
    const countsBefore = JSON.stringify(safe(s, '[...clubDrawCounts.entries()]'));
    const clicked = removeAt(s, 1);
    R.check('移除：按「移除」跳出 1 則 confirm，文字完全符合', clicked && c.messages.length === 1 && c.messages[0] === REMOVE_MSG,
        c.messages.join(' / ') || '（沒有跳 confirm）');
    R.check('移除：confirm 訊息不含社團名稱、測試姓名或班級',
        c.messages.length === 1 && c.messages.every(m => !/測試生|桌遊社|圍棋社|籃球社|\d{3}/.test(m)), c.messages.join(' / ') || '（沒有跳 confirm）');
    R.check('移除：按取消時社團仍在清單裡（3 個，順序不變）', clubs(s).length === 3 && clubNames(s) === NAMES.join(','), `社團：${clubNames(s)}`);
    R.test('移除：按取消時每個社團的抽籤結果都沒變', () => {
        const after = clubs(s).map(snapshot);
        return [after.length === 3 && after.every((x, i) => x === before[i]), `變動 ${after.filter((x, i) => x !== before[i]).length} 個`];
    });
    R.check('移除：按取消後清單仍顯示 3 個社團與 3 個「移除」按鈕', items(s).length === 3 && buttons(s).length === 3,
        `項目 ${items(s).length} 個，按鈕 ${buttons(s).length} 個`);
    R.check('移除：按取消不影響抽籤次數', JSON.stringify(safe(s, '[...clubDrawCounts.entries()]')) === countsBefore && countsBefore !== undefined,
        `${countsBefore} → ${JSON.stringify(safe(s, '[...clubDrawCounts.entries()]'))}`);
    R.check('移除：按取消時不跳其他提醒', s.alerts.length === 0, s.alerts.join(' / '));
}

// ---- 3. 按確定：移除中間的社團，其他社團的抽籤結果不變，抽籤次數不歸零
{
    const { s, c } = setup3();
    const before = clubs(s).map(snapshot);
    const countsBefore = JSON.stringify(safe(s, '[...clubDrawCounts.entries()]'));
    const clicked = removeAt(s, 1);
    R.check('移除：按確定時跳出 1 則 confirm，文字完全符合', clicked && removeMsgs(c).length === 1 && c.messages.length === 1,
        c.messages.join(' / ') || '（沒有跳 confirm）');
    R.check('移除：按確定後只移除中間的社團（剩桌遊社、籃球社）', clubs(s).length === 2 && clubNames(s) === '桌遊社,籃球社', `社團：${clubNames(s)}`);
    R.test('移除：其他社團的抽籤結果不變（正取的學生相同，沒有重抽）', () => {
        const after = clubs(s).map(snapshot);
        return [after.length === 2 && after[0] === before[0] && after[1] === before[2],
            `桌遊社${after[0] === before[0] ? '不變' : '有變'}，籃球社${after[1] === before[2] ? '不變' : '有變'}`];
    });
    R.check('移除：清單重新顯示 2 個社團與 2 個「移除」按鈕', items(s).length === 2 && buttons(s).length === 2,
        `項目 ${items(s).length} 個，按鈕 ${buttons(s).length} 個`);
    R.check('移除：移除不影響抽籤次數（計數表不變）',
        countsBefore !== undefined && JSON.stringify(safe(s, '[...clubDrawCounts.entries()]')) === countsBefore,
        `${countsBefore} → ${JSON.stringify(safe(s, '[...clubDrawCounts.entries()]'))}`);
    R.check('移除：移除中間的社團後，焦點移到第 2 個「移除」按鈕（籃球社）',
        buttons(s).length === 2 && s.document.activeElement === buttons(s)[1], `焦點：${focusLabel(s)}`);
    R.check('移除：按確定時不跳其他提醒', s.alerts.length === 0, s.alerts.join(' / '));

    // 移除後再加入同名社團：沿用既有的重抽 confirm，記為第 2 次
    addClub(s, ROSTERS[1], '圍棋社');
    const re = otherMsgs(c);
    const readded = clubs(s).find(x => x.clubName === '圍棋社');
    R.check('移除：再加入同名社團時跳出重抽 confirm（已經抽過 1 次）', re.length === 1 && re[0] === REDRAW_MSG(1), re.join(' / ') || '（沒有跳重抽 confirm）');
    R.check('移除：再加入同名社團記為第 2 次抽籤（drawRound = 2）', clubs(s).length === 3 && readded?.drawRound === 2,
        `社團：${clubNames(s)}，drawRound=${readded?.drawRound}`);
}

// ---- 4. 移除後的焦點
{
    // 移除最後一個：焦點移到新的最後一個（index 1）
    const { s } = setup3();
    removeAt(s, 2);
    R.check('焦點：移除最後一個社團後，焦點移到第 2 個「移除」按鈕（圍棋社）',
        clubNames(s) === '桌遊社,圍棋社' && buttons(s).length === 2 && s.document.activeElement === buttons(s)[1], `社團：${clubNames(s)}，焦點：${focusLabel(s)}`);
}
{
    // 移除第一個：焦點移到新的第一個（index 0）
    const { s } = setup3();
    removeAt(s, 0);
    R.check('焦點：移除第一個社團後，焦點移到第 1 個「移除」按鈕（圍棋社）',
        clubNames(s) === '圍棋社,籃球社' && buttons(s).length === 2 && s.document.activeElement === buttons(s)[0], `社團：${clubNames(s)}，焦點：${focusLabel(s)}`);
}
{
    // 焦點必須落在重新產生的按鈕上，不是已經從清單拿掉的舊按鈕
    const { s } = setup3();
    removeAt(s, 1);
    const a = s.document.activeElement;
    R.check('焦點：焦點所在的按鈕仍在清單裡（不是已移除的舊按鈕）', !!a && buttons(s).includes(a), `焦點：${focusLabel(s)}`);
}
{
    // 移除唯一的社團：清單變空，焦點移到貼上框
    const s = loadSite();
    const c = stubConfirm(s);
    addClub(s, ROSTERS[0], '桌遊社');
    s.document.activeElement = null;
    removeAt(s, 0);
    R.check('焦點：移除唯一的社團後清單變空，焦點移到貼上框（#paste-text-area）',
        removeMsgs(c).length === 1 && clubs(s).length === 0 && buttons(s).length === 0 && s.document.activeElement === s.el('paste-text-area'),
        `confirm ${removeMsgs(c).length} 則，社團 ${clubs(s).length} 個，焦點：${focusLabel(s)}`);
    R.check('焦點：清單變空時隱藏「已加入的社團」區塊', s.el('loaded-clubs-section').style.display === 'none', `display=${s.el('loaded-clubs-section').style.display}`);
}

// ---- 5. 按「移除」不會展開或收合名單（stopPropagation）
{
    const { s, c } = setup3();
    c.answer = false;
    let stopped = 0;
    const clicked = removeAt(s, 0, { stopPropagation() { stopped++; } });
    const header = s.el('loaded-clubs-list').descendants().find(el => el.classList.contains('loaded-club-info'));
    const details = s.el('loaded-clubs-list').descendants().find(el => el.classList.contains('loaded-club-details'));
    R.check('移除：按「移除」時呼叫 stopPropagation，名單維持收合',
        clicked && stopped >= 1 && details?.hidden === true && header?.getAttribute('aria-expanded') === 'false' && removeMsgs(c).length === 1,
        `stopPropagation ${stopped} 次，名單${details?.hidden ? '收合' : '展開'}，confirm ${removeMsgs(c).length} 則`);
}

// ---- 6. 「全部移除」的 confirm 與單一移除用語一致：先問、再說後果，寫明抽籤次數會累加（不是歸零）
{
    const { s, c } = setup3();
    c.answer = false;
    s.run('clearAllPastedClubs()');
    const msg = c.messages.at(-1) || '';
    R.check('全部移除：confirm 是「確定要移除所有已加入的社團嗎？…抽籤次數會累加。」',
        msg === '確定要移除所有已加入的社團嗎？抽籤結果會一起清除。移除後要重新貼上名單並加入，加入時會重新抽籤，抽籤次數會累加。', msg);
    R.check('全部移除：按取消時 3 個社團都還在', clubs(s).length === 3, `剩 ${clubs(s).length} 個`);
}

R.done();
