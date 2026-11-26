// 貼上模式的抽籤：正取名額與籤序、年級範圍與優先順序、新生排最後、備取順序、班級無法辨識、parseGrade、
// 重複報名去重、只用密碼學亂數（拒絕取樣、分布、沒有 crypto 時不抽籤）、加入社團後的提醒、匯出檔的「抽籤資訊」，
// 以及同名社團重抽的計數與確認（第 16 節）。
// 測試資料的姓名全部是虛構的（測試生<班級>_<座號>）。
// 舊版沒有 dedupeApplicants() 等函式時，個別檢查記成 FAIL，不會讓整個檔案中斷。
import vm from 'node:vm';
import { readFileSync } from 'node:fs';
import { loadSite, makeReporter, readSiteFile, stubConfirm, EXCELJS_PATH } from './harness.mjs';
const R = makeReporter('抽籤：名額、順序、去重與亂數');

const mk = (cls, seat, name) => ({ class: cls, seat, name: name ?? `測試生${cls || '新生'}_${seat || '00'}` });
const seqs = (list) => list.filter(s => s.selected).map(s => s.drawSequence).sort((a, b) => a - b);
const consecutive = (arr) => arr.every((v, i) => v === i + 1);
const gradeOf = (cls) => Number(String(cls)[0]);
const backupNums = (list) => list.filter(s => /^備取 \d+$/.test(String(s.drawSequence))).map(s => Number(String(s.drawSequence).slice(3)));
const ZWSP = String.fromCharCode(0x200B);

// 一份跨 1–6 年級的假名單：每個年級 count 人，班級 g01，座號 01…
function roster(grades = [1, 2, 3, 4, 5, 6], count = 4) {
    const list = [];
    for (const g of grades) for (let i = 1; i <= count; i++) list.push(mk(`${g}01`, String(i).padStart(2, '0')));
    return list;
}

// ---- 1. 正取不超過名額、籤序 1..N 連續
{
    const s = loadSite();
    const lottery = s.fn('runClubLottery');
    for (const quota of [1, 7, 24, 30]) {
        R.test(`名額 ${quota}：正取人數不超過名額，籤序 1..N 連續`, () => {
            const out = lottery(roster(), quota, '1', '6', 'desc');
            const sq = seqs(out);
            const expect = Math.min(quota, 24);
            return [out.length === 24 && sq.length === expect && consecutive(sq), `共 ${out.length} 筆，正取 ${sq.length}，籤序 ${sq.join(',')}`];
        });
    }
}

// ---- 2. 年級範圍篩選
{
    const s = loadSite();
    const lottery = s.fn('runClubLottery');
    R.test('年級範圍 3–4：只有 3、4 年級參加抽籤，其他標示「不符限制」', () => {
        const out = lottery(roster(), 100, '3', '4', 'desc');
        const sel = out.filter(s => s.selected);
        const okSel = sel.length === 8 && sel.every(s => [3, 4].includes(gradeOf(s.class)));
        const rest = out.filter(s => !s.selected);
        const okRest = rest.length === 16 && rest.every(s => s.drawSequence === '不符限制' && ![3, 4].includes(gradeOf(s.class)));
        return [okSel && okRest, `正取 ${sel.length}，不符限制 ${rest.filter(s => s.drawSequence === '不符限制').length}`];
    });
}

// ---- 3. 高年級優先、低年級優先
{
    const s = loadSite();
    const lottery = s.fn('runClubLottery');
    R.test('高年級優先：名額 6 時 6 年級全上，再抽 5 年級 2 人；6 年級拿到籤序 1–4', () => {
        const out = lottery(roster([1, 5, 6], 4), 6, '1', '6', 'desc');
        const sel = out.filter(s => s.selected);
        const g6 = sel.filter(s => gradeOf(s.class) === 6), g5 = sel.filter(s => gradeOf(s.class) === 5);
        const ok = g6.length === 4 && g5.length === 2 && g6.every(s => s.drawSequence <= 4) && !sel.some(s => gradeOf(s.class) === 1);
        return [ok, `6 年級 ${g6.length}、5 年級 ${g5.length}`];
    });
    R.test('低年級優先：名額 6 時 1 年級全上，再抽 2 年級 2 人', () => {
        const out = lottery(roster([1, 2, 6], 4), 6, '1', '6', 'asc');
        const sel = out.filter(s => s.selected);
        const g1 = sel.filter(s => gradeOf(s.class) === 1), g2 = sel.filter(s => gradeOf(s.class) === 2);
        const ok = g1.length === 4 && g2.length === 2 && g1.every(s => s.drawSequence <= 4) && !sel.some(s => gradeOf(s.class) === 6);
        return [ok, `1 年級 ${g1.length}、2 年級 ${g2.length}`];
    });
}

// ---- 4. 新生排最後（高、低年級優先都一樣）
{
    const s = loadSite();
    const lottery = s.fn('runClubLottery');
    for (const priority of ['desc', 'asc']) {
        R.test(`新生排最後（${priority === 'desc' ? '高' : '低'}年級優先）：名額 3 時 3 位 1 年級全上，新生全是備取`, () => {
            const list = [mk('', '', '測試生新生_01'), mk('新生', '', '測試生新生_02'), mk('', '', '測試生新生_03'), ...roster([1], 3)];
            const out = lottery(list, 3, '新生', '6', priority);
            const sel = out.filter(s => s.selected);
            const fresh = out.filter(s => !s.class || s.class === '新生');
            const ok = sel.length === 3 && sel.every(s => gradeOf(s.class) === 1) && fresh.every(s => /^備取 /.test(String(s.drawSequence)));
            return [ok, `正取 ${sel.map(s => s.class || '新生').join(',')}`];
        });
    }
}

// ---- 5. 備取 1..M 連續且照號碼排列；後面依序是「班級無法辨識」「不符限制」，各自依班級、座號、姓名排序
{
    const s = loadSite();
    const lottery = s.fn('runClubLottery');
    R.test('備取 1..M 連續，且未錄取部分照備取號碼排列（連續 5 次）', () => {
        for (let round = 0; round < 5; round++) {
            const out = lottery(roster([3, 4], 8), 4, '3', '4', 'desc');
            const unselected = out.filter(s => !s.selected);
            const nums = unselected.map(s => Number(String(s.drawSequence).replace('備取 ', '')));
            if (unselected.length !== 12 || !consecutive(nums)) return [false, `第 ${round + 1} 次：${unselected.map(s => s.drawSequence).join(',')}`];
        }
        return true;
    });
    R.test('回傳順序：正取（依班級座號）→ 備取 1..M → 班級無法辨識 → 不符限制', () => {
        const list = [
            mk('101', '02'), mk('甲班', '09'), mk('資源班', '01'), mk('602', '05'), mk('601', '03'), mk('701', '04'),
            mk('102', '01'), mk('甲班', '02'), ...roster([3], 6),
        ];
        const out = lottery(list, 2, '3', '6', 'desc');
        const kinds = out.map(s => s.selected ? 'S' : /^備取 /.test(String(s.drawSequence)) ? 'B' : s.drawSequence === '班級無法辨識' ? 'U' : s.drawSequence === '不符限制' ? 'D' : '?');
        const order = kinds.join('');
        const grouped = /^S*B*U*D*$/.test(order);
        const bn = backupNums(out);
        const sel = out.filter(s => s.selected);
        const selSorted = sel.every((s, i) => i === 0 || Number(sel[i - 1].class + sel[i - 1].seat) <= Number(s.class + s.seat));
        const unrec = out.filter(s => s.drawSequence === '班級無法辨識').map(s => `${s.class}${s.seat}`);
        const disq = out.filter(s => s.drawSequence === '不符限制').map(s => `${s.class}${s.seat}`);
        // 班級無法辨識依班級（getNumericSortKey）、座號、姓名排序：與頁面自己的排序鍵比對，不依賴輸入順序
        const key = s.fn('getNumericSortKey');
        const unrecList = out.filter(x => x.drawSequence === '班級無法辨識');
        const unrecSorted = unrecList.every((x, i) => {
            if (i === 0) return true;
            const p = unrecList[i - 1];
            const dc = key(p.class) - key(x.class);
            return dc < 0 || (dc === 0 && key(p.seat) <= key(x.seat));
        });
        const unrecOk = unrec.length === 4 && unrecSorted && unrec.indexOf('甲班02') < unrec.indexOf('甲班09')
            && ['甲班02', '甲班09', '資源班01', '70104'].every(x => unrec.includes(x));
        const disqOk = disq.join() === '10102,10201';
        return [grouped && consecutive(bn) && selSorted && unrecOk && disqOk, `順序 ${order}；備取 ${bn.join(',')}；無法辨識 ${unrec.join(',')}；不符 ${disq.join(',')}`];
    });
}

// ---- 6. 班級無法辨識不參加抽籤
{
    const s = loadSite();
    const lottery = s.fn('runClubLottery');
    for (const cls of ['甲班', '701', '資源班', '01']) {
        R.test(`班級「${cls}」：不參加抽籤，標示「班級無法辨識」（名額足夠也一樣）`, () => {
            const out = lottery([mk(cls, '05'), mk('301', '05')], 10, '1', '6', 'desc');
            const t = out.find(x => x.class === cls);
            return [t.selected === false && t.drawSequence === '班級無法辨識' && seqs(out).join() === '1', `selected=${t.selected}，籤序=${t.drawSequence}`];
        });
    }
}

// ---- 7. parseGrade
{
    const s = loadSite();
    const parseGrade = s.fn('parseGrade');
    const cases = [
        ['', 0], [null, 0], [undefined, 0], ['新生', 0], ['  ', 0], ['新生1班', 0], ['新生A', 0],
        ['301', 3], ['3年1班', 3], ['三年甲班', 3], ['三甲', 3], ['３０１', 3], [' 601 ', 6], ['六年二班', 6], [401, 4],
        [ZWSP + '201', 2], ['(502)', 5], ['1', 1],
        ['701', null], ['甲班', null], ['資源班', null], ['01', null], ['0', null], ['十一班', null], ['七年一班', null],
    ];
    for (const [input, expect] of cases) {
        const label = input === ZWSP + '201' ? '零寬空白+201' : JSON.stringify(input);
        R.test(`parseGrade(${label}) 是 ${expect}`, () => {
            const got = parseGrade(input);
            return [got === expect, `得到 ${got}`];
        });
    }
}

// ---- 8. 去重：規則與 upsertStudent() 一致
{
    const s = loadSite();
    const dedupe = s.fn('dedupeApplicants');
    const upsert = s.fn('upsertStudent');
    const run = (list) => dedupe(list.map(x => ({ ...x })));
    R.test('同班同名同座號：算重複，只保留第一筆', () => {
        const r = run([mk('301', '05'), mk('302', '01'), mk('301', '05')]);
        return [r.duplicateCount === 1 && r.students.length === 2 && r.students[0].seat === '05', `重複 ${r.duplicateCount}，剩 ${r.students.length}`];
    });
    R.test('一方沒有座號：算重複；第一筆沒有座號時改用後面那筆的座號', () => {
        const a = run([mk('301', '', '測試生301_05'), mk('301', '05', '測試生301_05')]);
        const b = run([mk('301', '05', '測試生301_05'), mk('301', '', '測試生301_05')]);
        return [a.duplicateCount === 1 && a.students.length === 1 && a.students[0].seat === '05'
            && b.duplicateCount === 1 && b.students.length === 1 && b.students[0].seat === '05', `a 剩 ${a.students.length}（座號 ${a.students[0]?.seat}），b 剩 ${b.students.length}`];
    });
    R.test('同班同名但座號不同：不算重複', () => {
        const r = run([mk('301', '05', '測試生301_05'), mk('301', '06', '測試生301_05')]);
        return [r.duplicateCount === 0 && r.students.length === 2, `重複 ${r.duplicateCount}`];
    });
    R.test('不同班同名：不算重複', () => {
        const r = run([mk('301', '05', '測試生301_05'), mk('302', '05', '測試生301_05')]);
        return [r.duplicateCount === 0 && r.students.length === 2, `重複 ${r.duplicateCount}`];
    });
    R.test('班級含不可見字元、空白班級與「新生」：清除後相同就算重複', () => {
        const r = run([mk('301', '05'), mk('301' + ZWSP, '05', '測試生301_05'), mk('', '', '測試生新生_00'), mk('新生', '', '測試生新生_00')]);
        return [r.duplicateCount === 2 && r.students.length === 2, `重複 ${r.duplicateCount}`];
    });
    R.test('去重後的人數與 upsertStudent() 合併後的人數相同（多種組合）', () => {
        const N = '測試生301_05';
        const combos = [
            [['301', ''], ['301', '05'], ['301', '06']],
            [['301', '05'], ['301', '06'], ['301', '']],
            [['301', '05'], ['301', '06'], ['301', '06'], ['301', '05']],
            [['', ''], ['新生', '01'], ['', '02']],
            [['301', ''], ['301', ''], ['302', '']],
        ];
        const bad = [];
        combos.forEach((c, i) => {
            const list = c.map(([cls, seat]) => mk(cls, seat, N));
            const map = new Map();
            list.forEach(x => upsert(map, x.class, x.seat, x.name));
            const r = run(list);
            if (r.students.length !== map.size) bad.push(`組合 ${i + 1}：去重 ${r.students.length}、合併 ${map.size}`);
        });
        return [bad.length === 0, bad.join('；')];
    });
}

// ---- 9. Math.random 不可使用：換成會丟例外的函式後，抽籤仍然正常
{
    const s = loadSite();
    s.run('Math.random = () => { throw new Error("抽籤不可使用 Math.random"); }');
    R.test('Math.random 會丟例外時，runClubLottery 仍正常（正取 5、籤序連續）', () => {
        const out = s.fn('runClubLottery')(roster(), 5, '1', '6', 'desc');
        return [seqs(out).length === 5 && consecutive(seqs(out)), `正取 ${seqs(out).length}`];
    });
    R.test('Math.random 會丟例外時，從介面加入社團仍正常', () => {
        addClub(s, '班級\t座號\t姓名\n301\t05\t測試生301_05\n302\t11\t測試生302_11\n303\t07\t測試生303_07', { quota: '2', start: '1', end: '6' });
        return [s.run('pastedClubs.length') === 1 && s.run('pastedClubs[0].selectedCount') === 2, `社團 ${s.run('pastedClubs.length')} 個；alert：${s.alerts.join(' / ')}`];
    });
}

// ---- 10. 假 crypto：拒絕取樣會丟掉超出範圍的值
{
    const queue = [];
    let calls = 0, lastType = '';
    const fake = { getRandomValues(arr) { calls++; lastType = Object.prototype.toString.call(arr); for (let i = 0; i < arr.length; i++) arr[i] = queue.length ? queue.shift() : 0; return arr; } };
    const s = loadSite({ crypto: fake });
    const randomInt = s.fn('randomInt');
    const feed = (vals) => { queue.length = 0; queue.push(...vals); calls = 0; };
    R.test('randomInt(3)：0xFFFFFFFF（超出可整除範圍）被丟掉，改用下一個值 7 → 1', () => {
        feed([0xFFFFFFFF, 7]);
        const v = randomInt(3);
        return [v === 1 && calls === 2, `結果 ${v}，取了 ${calls} 次`];
    });
    R.test('randomInt(6)：丟掉 4294967292、4294967295，接受 4294967291', () => {
        feed([4294967292, 4294967295, 4294967291]);
        const v = randomInt(6);
        return [v === 4294967291 % 6 && calls === 3, `結果 ${v}，取了 ${calls} 次`];
    });
    R.test('randomInt(3)：範圍內的最大值 4294967294 直接接受', () => {
        feed([4294967294]);
        const v = randomInt(3);
        return [v === 4294967294 % 3 && calls === 1, `結果 ${v}，取了 ${calls} 次`];
    });
    R.test('randomInt 使用 Uint32Array 取亂數', () => { feed([5]); randomInt(10); return [lastType === '[object Uint32Array]', lastType]; });
    R.test('randomInt(0)、randomInt(1.5)、randomInt(-1) 丟出 RangeError', () => {
        // s.fn() 在函式不存在時也會丟例外，所以要先確認函式存在、而且丟的是 RangeError
        if (s.run('typeof randomInt') !== 'function') return [false, '頁面沒有 randomInt()'];
        let n = 0;
        for (const bad of [0, 1.5, -1]) { try { randomInt(bad); } catch (e) { if (e.name === 'RangeError') n++; } }
        return [n === 3, `${n}/3 丟出 RangeError`];
    });
    R.test('shuffleArray 用假 crypto 時結果可預期（全取 0 → [b, c, a]）', () => {
        feed([0, 0]);
        const arr = ['a', 'b', 'c'];
        s.fn('shuffleArray')(arr);
        return [arr.join() === 'b,c,a' && calls === 2, `${arr.join()}，取了 ${calls} 次`];
    });
}

// ---- 11. 真的 crypto：洗牌 3 個元素 12000 次，每種排列頻率合理，而且真的用了 crypto
{
    const real = globalThis.crypto;
    let calls = 0;
    const s = loadSite({ crypto: { getRandomValues: (a) => { calls++; return real.getRandomValues(a); } } });
    const shuffle = s.fn('shuffleArray');
    const N = 12000;
    const counts = {};
    let err = null;
    try {
        for (let i = 0; i < N; i++) { const a = ['a', 'b', 'c']; shuffle(a); const k = a.join(''); counts[k] = (counts[k] || 0) + 1; }
    } catch (e) { err = e; }
    const perms = Object.keys(counts);
    const exp = N / 6;
    const chi = perms.reduce((sum, k) => sum + (counts[k] - exp) ** 2 / exp, 0);
    R.check('洗牌 3 個元素 12000 次：6 種排列都出現，各自在期望值 ±10% 內（期望 2000）',
        !err && perms.length === 6 && perms.every(k => Math.abs(counts[k] - exp) <= exp * 0.1),
        err ? `例外：${err.message}` : `${perms.map(k => `${k}=${counts[k]}`).join(' ')}，卡方 ${chi.toFixed(2)}`);
    R.check('洗牌 12000 次都向 crypto.getRandomValues 取亂數（每次至少 2 個）', calls >= N * 2, `呼叫 ${calls} 次`);
}

// ---- 12. 沒有 crypto 時不抽籤、說明原因、不加入社團
{
    const s = loadSite({ crypto: null });
    let mathCalls = 0;
    s.ctx.__countMath = () => { mathCalls++; return 0.5; };
    s.run('Math.random = () => __countMath()');
    addClub(s, '班級\t座號\t姓名\n301\t05\t測試生301_05\n302\t11\t測試生302_11', { quota: '1', start: '1', end: '6' });
    R.check('沒有 crypto：不加入社團', s.run('pastedClubs.length') === 0, `社團 ${s.run('pastedClubs.length')} 個`);
    R.test('重抽：沒有 crypto 被擋下時不計抽籤次數', () => [s.run('clubDrawCounts.size') === 0, `計數表 ${s.run('clubDrawCounts.size')} 筆`]);
    R.check('沒有 crypto：跳出 1 則說明（不支援安全的亂數、改用最新版 Chrome、Edge 或 Safari）',
        s.alerts.length === 1 && /不支援安全的亂數/.test(s.alerts[0]) && /Chrome、Edge 或 Safari/.test(s.alerts[0]), s.alerts.join(' / '));
    R.test('沒有 crypto：runClubLottery 丟出例外，不退回 Math.random', () => {
        let threw = false;
        try { s.fn('runClubLottery')(roster([3], 4), 2, '1', '6', 'desc'); } catch (e) { threw = true; }
        return [threw && mathCalls === 0, `丟出例外=${threw}，Math.random 呼叫 ${mathCalls} 次`];
    });
}

// ---- 13. 加入社團後的提醒與 meta 行
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
const MIXED = [
    '班級\t座號\t姓名',
    '301\t05\t測試生301_05',
    '302\t\t測試生302_07',
    '301\t05\t測試生301_05',
    '302\t07\t測試生302_07',
    '甲班\t03\t測試生甲班_03',
    '101\t02\t測試生101_02',
    '303\t08\t測試生303_08',
].join('\n');
const EXPECT_BOTH = '名單裡有 2 筆重複報名（班級、姓名相同），重複的只保留一筆，已用去除重複後的名單抽籤；另有 1 筆班級無法辨識，沒有參加抽籤，名單中標示為「班級無法辨識」。請確認貼上的內容，需要時移除這個社團後重新貼上。';
{
    const s = loadSite();
    addClub(s, MIXED, { quota: '2', start: '3', end: '6' });
    const club = s.run('pastedClubs[0]');
    R.check('有重複與班級無法辨識：加入社團（去重後 5 人）', club && club.students.length === 5, club ? `名單 ${club.students.length} 人` : '沒有加入');
    R.check('有重複與班級無法辨識：只跳出 1 則提醒', s.alerts.length === 1, `${s.alerts.length} 則`);
    R.check('提醒文字：寫出筆數與下一步', s.alerts[0] === EXPECT_BOTH, s.alerts[0]);
    R.check('提醒文字不含任何測試姓名或班級', s.alerts.length > 0 && !s.alerts.some(a => /測試生|甲班|301|302/.test(a)), s.alerts.join(' / '));
    R.check('社團記下 duplicateCount=2、unrecognizedCount=1、applicantCount=7', club && club.duplicateCount === 2 && club.unrecognizedCount === 1 && club.applicantCount === 7,
        club ? `${club.duplicateCount}/${club.unrecognizedCount}/${club.applicantCount}` : '');
    R.check('社團記下抽籤時間 drawnAt（Date）', club && club.drawnAt && typeof club.drawnAt.getTime === 'function' && !isNaN(club.drawnAt.getTime()));
    const meta = s.el('loaded-clubs-list').descendants().find(el => el.classList.contains('loaded-club-info'))?.innerHTML || '';
    R.test('meta 行顯示「略過重複 2 筆」「班級無法辨識 1 筆」與 HH:mm 抽籤時間', () => {
        const hm = s.fn('formatDrawTime')(club.drawnAt);
        return [/略過重複 2 筆/.test(meta) && /班級無法辨識 1 筆/.test(meta) && /^\d{2}:\d{2}$/.test(hm) && meta.includes(`${hm} 抽籤`), meta.replace(/<[^>]+>/g, ' ').slice(0, 160)];
    });
    R.test('名單明細的未錄取順序：備取 1..M → 班級無法辨識 → 不符限制', () => {
        const details = s.el('loaded-clubs-list').descendants().find(el => el.classList.contains('loaded-club-details'))?.innerHTML || '';
        const backupDiv = (details.match(/roster-backup">([\s\S]*)<\/div>$/) || [])[1] || '';
        const tags = [...backupDiv.matchAll(/\(([^()]+)\)/g)].map(m => m[1]);
        const kinds = tags.map(t => /^備取 /.test(t) ? 'B' : t === '班級無法辨識' ? 'U' : t === '不符限制' ? 'D' : '?').join('');
        const nums = tags.filter(t => /^備取 /.test(t)).map(t => Number(t.slice(3)));
        // 去重後 5 人：3 年級以上 3 人（正取 2、備取 1）、甲班（無法辨識）、101（不符限制）
        return [kinds === 'BUD' && consecutive(nums), `${kinds}：${tags.join(',')}`];
    });
    R.check('抽籤資訊與名單沒有寫進 localStorage', [...s.storage.values()].every(v => !/測試生|drawnAt|抽籤/.test(v)), [...s.storage.keys()].join(','));
}
{
    const s = loadSite();
    addClub(s, '班級\t座號\t姓名\n301\t05\t測試生301_05\n302\t11\t測試生302_11', { quota: '1', start: '1', end: '6' });
    R.check('沒有重複也沒有班級無法辨識：不跳提醒', s.run('pastedClubs.length') === 1 && s.alerts.length === 0, s.alerts.join(' / '));
}
{
    const s = loadSite();
    addClub(s, '班級\t座號\t姓名\n301\t05\t測試生301_05\n301\t05\t測試生301_05\n302\t11\t測試生302_11', { quota: '1', start: '1', end: '6' });
    R.check('只有重複：提醒只寫重複筆數', s.alerts[0] === '名單裡有 1 筆重複報名（班級、姓名相同），重複的只保留一筆，已用去除重複後的名單抽籤。請確認貼上的內容，需要時移除這個社團後重新貼上。', s.alerts[0]);
}
{
    const s = loadSite();
    addClub(s, '班級\t座號\t姓名\n資源班\t05\t測試生資源班_05\n302\t11\t測試生302_11', { quota: '1', start: '1', end: '6' });
    R.check('只有班級無法辨識：提醒只寫無法辨識筆數', s.alerts[0] === '名單裡有 1 筆班級無法辨識，沒有參加抽籤，名單中標示為「班級無法辨識」。請確認貼上的內容，需要時移除這個社團後重新貼上。', s.alerts[0]);
}

// ---- 14. 匯出檔的「抽籤資訊」：用真的 ExcelJS 寫出再讀回
// ExcelJS 要載入頁面同一個 vm context：它用 instanceof Array 判斷 addRow 的參數，跨 context 的陣列會被當成物件而寫不出儲存格。
// 頁面的 setTimeout 是假時鐘，ExcelJS 的非同步工作要靠 advance() 推進。
function loadRealExcelJS(s) {
    Object.assign(s.ctx, { TextDecoder, TextEncoder, regeneratorRuntime: undefined });
    vm.runInContext(readFileSync(EXCELJS_PATH, 'utf8'), s.ctx);
    return s.ctx.ExcelJS;
}
async function pump(s, promise) {
    let done = false;
    const p = Promise.resolve(promise).finally(() => { done = true; });
    for (let i = 0; i < 20000 && !done; i++) { s.advance(5); await new Promise(r => setImmediate(r)); }
    return p;
}
{
    const s = loadSite({ withExcelJS: false });
    const lib = loadRealExcelJS(s);
    let blob = null;
    s.ctx.triggerFileDownload = (b) => { blob = b; };
    addClub(s, MIXED, { quota: '2', start: '3', end: '6', priority: 'desc', name: '桌遊社' });
    const club = s.run('pastedClubs[0]');
    // 匯出時下載按鈕會重新渲染清單，所以先記下畫面上的 meta 行
    const metaHtml = s.el('loaded-clubs-list').descendants().find(el => el.classList.contains('loaded-club-info'))?.innerHTML || '';
    const alertsBefore = s.alerts.length;
    await pump(s, s.run('processPastedData()'));
    let ws = null;
    if (blob) {
        const wb = new lib.Workbook();
        await pump(s, wb.xlsx.load(blob.parts[0]));
        ws = wb.getWorksheet('桌遊社');
    }
    R.check('匯出成功並有社團工作表', !!ws && s.alerts.length === alertsBefore, s.alerts.slice(alertsBefore).join(' / '));
    if (ws) {
        const v = (r, c) => ws.getCell(r, c).value;
        const fmt = s.fn('formatDrawTime');
        let expectTime = '';
        try { expectTime = fmt(club.drawnAt, true); } catch (e) { /* 舊版沒有 formatDrawTime */ }
        R.check('G1 是「抽籤資訊」標題，合併 G1:H1，表頭樣式（粗體白字、校徽藍底）',
            v(1, 7) === '抽籤資訊' && ws.model.merges.includes('G1:H1') && ws.getCell(1, 7).font?.bold && ws.getCell(1, 7).fill?.fgColor?.argb === 'FF0F3B6D',
            `G1=${v(1, 7)}，merges=${ws.model.merges.join(',')}`);
        // 抽籤資訊從 G2 往下讀到空白為止；各項依標題找列號，新增項目時不必改每個斷言
        const info = [];
        for (let r = 2; r <= ws.rowCount && v(r, 7); r++) info.push({ row: r, label: v(r, 7), value: v(r, 8) });
        const item = (label) => info.find(x => x.label === label) || { row: 0, label: '', value: undefined };
        const expect = [
            ['抽籤時間', expectTime],
            ['正取人數', '2 人（名額 2 人）'],
            ['年級限制', '3 年級 至 6 年級'],
            ['優先順序', '高年級優先'],
            ['報名人數', `${club.students.length} 人（已去除重複）`],
            ['重複略過筆數', '2 筆'],
            ['班級無法辨識筆數', '1 筆'],
        ];
        expect.forEach(([label, value]) => {
            const it = item(label);
            R.check(`抽籤資訊：${label} = ${value || '（本地時間）'}`, it.row > 0 && it.value === value && value !== '', `第 ${it.row} 列：${it.label} = ${it.value}`);
        });
        const ORDER = [...expect.map(([l]) => l), '亂數來源'];
        R.check('抽籤資訊的既有項目依序排列：抽籤時間在最前、亂數來源在最後', info.map(x => x.label).filter(l => ORDER.includes(l)).join(',') === ORDER.join(',') && info[0]?.label === '抽籤時間' && info.at(-1)?.label === '亂數來源',
            info.map(x => x.label).join(','));
        R.check('抽籤時間格式是 YYYY/MM/DD HH:mm:ss', /^\d{4}\/\d{2}\/\d{2} \d{2}:\d{2}:\d{2}$/.test(String(item('抽籤時間').value)), String(item('抽籤時間').value));
        R.check('抽籤資訊：亂數來源是瀏覽器內建的安全亂數', /^瀏覽器內建的安全亂數/.test(String(item('亂數來源').value)), `第 ${item('亂數來源').row} 列：${item('亂數來源').value}`);
        R.check('抽籤資訊的「報名人數」與畫面上的「報名 N 人」相同', String(item('報名人數').value).startsWith(`${club.students.length} 人`) && metaHtml.includes(`報名 ${club.students.length} 人`), `${item('報名人數').value}；畫面：${metaHtml.replace(/<[^>]+>/g, ' ').replace(/測試生\S*/g, '').slice(0, 80)}`);
        R.check('F 欄（資料表與抽籤資訊之間）是空的', Array.from({ length: ws.rowCount }, (_, i) => v(i + 1, 6)).every(x => x === null || x === undefined || x === ''));
        R.check('抽籤資訊每一列都沒有被套成資料列樣式（項目粗體、靠左、微軟正黑體、細框線；內容微軟正黑體、細框線）', info.length >= 8 && info.every(({ row }) => {
            const c = ws.getCell(row, 7);
            const d = ws.getCell(row, 8);
            return c.font?.bold && c.font?.name === 'Microsoft JhengHei' && c.alignment?.horizontal === 'left' && c.border?.top?.style === 'thin' &&
                !d.font?.bold && d.font?.name === 'Microsoft JhengHei' && d.alignment?.horizontal === 'left' && d.border?.top?.style === 'thin';
        }), `共 ${info.length} 列`);
        // 重抽紀錄：抽籤次數緊接在抽籤時間之後（第 3 列），格式與其他項目相同
        R.check('抽籤資訊第 3 列是「抽籤次數 = 第 1 次（本次開啟網頁後，同名社團的抽籤次數）」，其餘項目依序往下移一列',
            v(3, 7) === '抽籤次數' && v(3, 8) === '第 1 次（本次開啟網頁後，同名社團的抽籤次數）' &&
            info.map(x => x.label).join(',') === ['抽籤時間', '抽籤次數', ...ORDER.slice(1)].join(','),
            `第 3 列：${v(3, 7)} = ${v(3, 8)}；${info.map(x => x.label).join(',')}`);
        // 資料表：D 欄的順序與錄取人數合併
        const rows = Array.from({ length: club.students.length }, (_, i) => String(v(i + 2, 4)));
        const kinds = rows.map(t => /^\d+$/.test(t) ? 'S' : /^備取 /.test(t) ? 'B' : t === '班級無法辨識' ? 'U' : t === '不符限制' ? 'D' : '?').join('');
        R.check('匯出的籤序欄順序：正取 → 備取 1..M → 班級無法辨識 → 不符限制', kinds === 'SSBUD', `${kinds}：${rows.join(',')}`);
        const eMerges = ws.model.merges.filter(m => m.startsWith('E'));
        // 合併儲存格的其他格讀回時也會得到主格的值，只加總主格
        const eSum = Array.from({ length: club.students.length }, (_, i) => ws.getCell(i + 2, 5))
            .filter(c => !(c.isMerged && c.master.address !== c.address))
            .map(c => c.value).filter(x => typeof x === 'number').reduce((a, b) => a + b, 0);
        R.check('「錄取人數」欄照年級合併，數字合計等於正取人數', eSum === 2 && ws.model.merges.every(m => m === 'G1:H1' || m.startsWith('E')),
            `E 欄合計 ${eSum}，merges=${eMerges.join(',') || '（無）'}`);
        R.check('匯出檔沒有測試姓名以外的名冊欄位跑進抽籤資訊（G/H 欄不含姓名）',
            Array.from({ length: ws.rowCount }, (_, i) => `${v(i + 1, 7) ?? ''}${v(i + 1, 8) ?? ''}`).every(t => !/測試生/.test(t)));
    }
}

// ---- 14b. 班級寫法混用時，正取仍依年級排在一起，匯出檔「錄取人數」每個年級只有一段
{
    const MIXED_CLASS = [
        '班級\t座號\t姓名',
        '三甲\t03\t測試生三甲_03',
        '新生\t01\t測試生新生_01',
        '301\t02\t測試生301_02',
        '三甲\t05\t測試生三甲_05',
        '3年甲班\t04\t測試生3年甲班_04',
        '101\t06\t測試生101_06',
        '4年2班\t07\t測試生4年2班_07',
    ].join('\n');
    const s = loadSite({ withExcelJS: false });
    const lib = loadRealExcelJS(s);
    let blob = null;
    s.ctx.triggerFileDownload = (b) => { blob = b; };
    addClub(s, MIXED_CLASS, { quota: '10', start: '新生', end: '6', priority: 'desc', name: '混合班級' });
    const club = s.run('pastedClubs[0]');
    const parseGrade = s.fn('parseGrade');
    const grades = club ? club.students.filter(x => x.selected).map(x => parseGrade(x.class)) : [];
    R.check('混合班級寫法：7 人全部正取', grades.length === 7, `正取 ${grades.length} 人`);
    R.check('混合班級寫法：正取依年級排列（1、3、3、3、3、4，新生最後）', grades.join(',') === '1,3,3,3,3,4,0', grades.join(','));
    await pump(s, s.run('processPastedData()'));
    if (blob) {
        const wb = new lib.Workbook();
        await pump(s, wb.xlsx.load(blob.parts[0]));
        const ws = wb.getWorksheet('混合班級');
        const masters = Array.from({ length: grades.length }, (_, i) => ws.getCell(i + 2, 5))
            .filter(c => !(c.isMerged && c.master.address !== c.address))
            .map(c => c.value).filter(x => typeof x === 'number');
        R.check('混合班級寫法：「錄取人數」欄每個年級只有一段（1、4、1、1）', masters.join(',') === '1,4,1,1', masters.join(','));
    } else {
        R.check('混合班級寫法：匯出成功', false, s.alerts.join(' / '));
    }
}

// ---- 15. 抽籤資訊的顯示文字與 index.html 的選項一致
{
    const html = readSiteFile('index.html');
    const s = loadSite();
    R.test('PRIORITY_LABELS 與 #paste-priority 的選項文字相同', () => {
        const sel = (html.match(/<select id="paste-priority"[\s\S]*?<\/select>/) || [''])[0];
        const opts = Object.fromEntries([...sel.matchAll(/<option value="([^"]+)"[^>]*>([^<]+)<\/option>/g)].map(m => [m[1], m[2].trim()]));
        const labels = s.run('typeof PRIORITY_LABELS === "object" ? PRIORITY_LABELS : null');
        return [!!labels && Object.keys(opts).length === 2 && Object.entries(opts).every(([k, t]) => labels[k] === t), JSON.stringify(opts)];
    });
    R.test('gradeOptionLabel 與 #paste-grade-start 的選項文字相同', () => {
        const sel = (html.match(/<select id="paste-grade-start"[\s\S]*?<\/select>/) || [''])[0];
        const opts = [...sel.matchAll(/<option value="([^"]+)"[^>]*>([^<]+)<\/option>/g)].map(m => [m[1], m[2].trim()]);
        const f = s.fn('gradeOptionLabel');
        const bad = opts.filter(([k, t]) => f(k) !== t);
        return [opts.length === 7 && bad.length === 0, bad.map(([k, t]) => `${k}：${f(k)} ≠ ${t}`).join('；')];
    });
}

// ---- 16. 重抽紀錄：同名社團在本次開啟網頁後的抽籤次數（clubDrawCounts → club.drawRound）
// 只有重新載入頁面才歸零；「全部清除」、清除這次的資料、下載後清空都不歸零，否則等於提供了不留紀錄的重抽方法。
{
    const TWO = '班級\t座號\t姓名\n301\t05\t測試生301_05\n302\t11\t測試生302_11';
    const LISTED_NOTE = '清單裡已經有同名的社團，兩份抽籤結果都會匯出；只要保留新的結果，請先移除舊的那一個。';
    const REDRAW_MSG = (n, listed = false) => `這個社團在這次開啟網頁後已經抽過 ${n} 次籤。再加入會重新抽籤，匯出檔的抽籤資訊會記錄為第 ${n + 1} 次抽籤。${listed ? LISTED_NOTE : ''}確定要重新抽籤嗎？`;
    const ROUND_VALUE = (n) => `第 ${n} 次（本次開啟網頁後，同名社團的抽籤次數）`;
    // 舊版沒有 clubDrawCounts 時回傳 undefined，讓個別檢查記成 FAIL
    const safe = (s, code) => { try { return s.run(code); } catch (e) { return undefined; } };
    const clubs = (s) => s.run('pastedClubs');
    const lastClub = (s) => clubs(s).at(-1);
    const rounds = (s) => clubs(s).map(c => c.drawRound).join(',');
    const metaAt = (s, i) => (s.el('loaded-clubs-list').descendants().filter(el => el.classList.contains('loaded-club-info'))[i]?.innerHTML || '');
    const removeAt = (s, i) => s.el('loaded-clubs-list').descendants().filter(el => el.classList.contains('loaded-club-delete'))[i].dispatch('click');
    const redraws = (c) => c.messages.filter(m => /抽過/.test(m));
    const noRoster = (m) => !/測試生|桌遊社|圍棋社|301|302/.test(m);

    // 第一次加入 → 移除 → 再加入（確定）→ 移除 → 再加入（取消）
    {
        const s = loadSite();
        const c = stubConfirm(s);
        addClub(s, TWO, { name: '桌遊社' });
        R.check('重抽：第一次加入時 drawRound 是 1、不跳 confirm', clubs(s).length === 1 && lastClub(s).drawRound === 1 && c.messages.length === 0,
            `社團 ${clubs(s).length} 個，drawRound=${lastClub(s)?.drawRound}，confirm ${c.messages.length} 則`);
        R.check('重抽：第一次抽籤時 meta 行不顯示「第 N 次抽籤」', clubs(s).length === 1 && !/次抽籤/.test(metaAt(s, 0)), metaAt(s, 0).replace(/<[^>]+>/g, ' ').replace(/測試生\S*/g, '').slice(0, 120));

        removeAt(s, 0);
        addClub(s, TWO, { name: '桌遊社' });
        R.check('重抽：移除後再加入同名社團，跳出 1 則 confirm，文字只寫次數', redraws(c).length === 1 && redraws(c)[0] === REDRAW_MSG(1), redraws(c).join(' / '));
        R.check('重抽：confirm 訊息不含社團名稱、測試姓名或班級', redraws(c).length === 1 && redraws(c).every(noRoster), redraws(c).join(' / '));
        R.check('重抽：按確定後加入社團，drawRound 是 2', clubs(s).length === 1 && lastClub(s).drawRound === 2, `社團 ${clubs(s).length} 個，drawRound=${rounds(s)}`);
        R.check('重抽：meta 行顯示「第 2 次抽籤」', /，第 2 次抽籤</.test(metaAt(s, 0)), metaAt(s, 0).replace(/<[^>]+>/g, ' ').replace(/測試生\S*/g, '').slice(0, 160));

        removeAt(s, 0);
        c.answer = false;
        let lotteryCalls = 0;
        const realLottery = s.ctx.runClubLottery;
        s.ctx.runClubLottery = (...a) => { lotteryCalls++; return realLottery(...a); };
        const alertsBefore = s.alerts.length;
        addClub(s, TWO, { name: '桌遊社' });
        R.check('重抽：confirm 按取消時不加入、不抽籤、不跳其他提醒', clubs(s).length === 0 && lotteryCalls === 0 && s.alerts.length === alertsBefore && redraws(c).length === 2 && redraws(c)[1] === REDRAW_MSG(2),
            `社團 ${clubs(s).length} 個，抽籤 ${lotteryCalls} 次，confirm：${redraws(c).at(-1)}`);
        R.check('重抽：按取消時保留輸入的社團名稱與名單，可以直接再按一次', s.el('paste-club-name').value === '桌遊社' && s.el('paste-text-area').value === TWO);
        R.check('重抽：按取消不計次（計數仍是 2）', safe(s, 'clubDrawCounts.get("桌遊社")') === 2, `計數=${safe(s, 'clubDrawCounts.get("桌遊社")')}`);
        c.answer = true;
        addClub(s, TWO, { name: '桌遊社' });
        R.check('重抽：取消後再按確定，記錄為第 3 次（取消的那次沒有算進去）', clubs(s).length === 1 && lastClub(s).drawRound === 3 && redraws(c).at(-1) === REDRAW_MSG(2) && lotteryCalls === 1,
            `drawRound=${rounds(s)}，抽籤 ${lotteryCalls} 次`);
    }

    // 「全部清除」、清除這次的資料、下載後清空社團都不歸零
    {
        const s = loadSite();
        const c = stubConfirm(s);
        let downloads = 0;
        s.ctx.triggerFileDownload = () => { downloads++; };
        addClub(s, TWO, { name: '桌遊社' });
        s.run('clearAllPastedClubs()');
        const clearedAll = clubs(s).length === 0;
        addClub(s, TWO, { name: '桌遊社' });
        R.check('重抽：「全部清除」之後再加入，次數仍然累加（第 2 次）', clearedAll && clubs(s).length === 1 && lastClub(s).drawRound === 2 && redraws(c).at(-1) === REDRAW_MSG(1),
            `清除後 ${clearedAll ? '0' : '仍有'} 個，drawRound=${rounds(s)}`);
        s.run('manualHardResetAll()');
        const reset = clubs(s).length === 0;
        addClub(s, TWO, { name: '桌遊社' });
        R.check('重抽：手動清除這次的資料（manualHardResetAll）之後再加入，次數仍然累加（第 3 次）', reset && clubs(s).length === 1 && lastClub(s).drawRound === 3 && redraws(c).at(-1) === REDRAW_MSG(2),
            `清除後 ${reset ? '0' : '仍有'} 個，drawRound=${rounds(s)}`);
        await s.run('processPastedData()');
        await s.settle();
        const exported = downloads === 1 && clubs(s).length === 0;
        addClub(s, TWO, { name: '桌遊社' });
        R.check('重抽：轉換下載後清空社團，再加入次數仍然累加（第 4 次）', exported && clubs(s).length === 1 && lastClub(s).drawRound === 4 && redraws(c).at(-1) === REDRAW_MSG(3),
            `下載 ${downloads} 次，drawRound=${rounds(s)}`);
    }

    // 驗證失敗不計次；不同社團各自計數；前後空白與不可見字元視為同一個社團
    {
        const s = loadSite();
        const c = stubConfirm(s);
        addClub(s, TWO, { name: '桌遊社', quota: '0' });
        addClub(s, '', { name: '桌遊社' });
        addClub(s, '這不是名單', { name: '桌遊社' });
        const blocked = s.alerts.length === 3 && clubs(s).length === 0;
        R.check('重抽：名額錯誤、沒有名單、讀不到學生時不計次', blocked && safe(s, 'clubDrawCounts.size') === 0, `提醒 ${s.alerts.length} 則，計數表 ${safe(s, 'clubDrawCounts.size')} 筆`);
        addClub(s, TWO, { name: '桌遊社' });
        R.check('重抽：驗證失敗之後第一次成功加入仍是第 1 次，不跳 confirm', clubs(s).length === 1 && lastClub(s).drawRound === 1 && c.messages.length === 0, `drawRound=${rounds(s)}，confirm ${c.messages.length} 則`);

        const alertsBefore = s.alerts.length;
        addClub(s, TWO, { name: '桌遊社', quota: '0' });
        R.check('重抽：已抽過的社團名額填錯時，先提醒名額、不跳重抽 confirm、不計次', s.alerts.length === alertsBefore + 1 && /正取人數/.test(s.alerts.at(-1)) && c.messages.length === 0 && safe(s, 'clubDrawCounts.get("桌遊社")') === 1,
            `confirm ${c.messages.length} 則，計數=${safe(s, 'clubDrawCounts.get("桌遊社")')}`);

        addClub(s, TWO, { name: '圍棋社' });
        R.check('重抽：不同社團名稱各自計數（圍棋社第 1 次、不跳 confirm）', clubs(s).length === 2 && lastClub(s).drawRound === 1 && c.messages.length === 0, `drawRound=${rounds(s)}`);

        const SP = String.fromCharCode(0x3000);
        const WJ = String.fromCharCode(0x2060);
        addClub(s, TWO, { name: `${ZWSP} 桌遊社${SP}${WJ}` });
        // 第 1 個桌遊社還在清單裡，所以 confirm 多一句同名提醒
        R.check('重抽：名稱前後有空白或不可見字元時視為同一個社團（跳 confirm，記為第 2 次）', clubs(s).length === 3 && lastClub(s).drawRound === 2 && redraws(c).length === 1 && redraws(c)[0] === REDRAW_MSG(1, true),
            `drawRound=${rounds(s)}，confirm：${redraws(c).join(' / ')}`);
        R.check('重抽：計數表的鍵是清除不可見字元與空白後的名稱', JSON.stringify(safe(s, '[...clubDrawCounts.entries()]')) === JSON.stringify([['桌遊社', 2], ['圍棋社', 1]]),
            JSON.stringify(safe(s, '[...clubDrawCounts.entries()]')));
    }

    // 寫法不同但看起來是同一個名稱：大小寫、全形英數、名稱中間的空白都算同一個社團，不可藉此不留紀錄地重抽
    for (const [first, second, label] of [
        ['Lego社', 'LEGO社', '英文大小寫不同'],
        ['AI社', 'ＡＩ社', '全形英數'],
        ['桌遊社', '桌遊 社', '名稱中間多一個空白'],
        ['圍棋社', `圍${String.fromCharCode(0x3000)}棋社`, '名稱中間多一個全形空白'],
    ]) {
        const s = loadSite();
        const c = stubConfirm(s);
        addClub(s, TWO, { name: first });
        removeAt(s, 0);
        addClub(s, TWO, { name: second });
        R.check(`重抽：${label}（${first} → ${second}）視為同一個社團，跳 confirm 並記為第 2 次`,
            redraws(c).length === 1 && redraws(c)[0] === REDRAW_MSG(1) && clubs(s).length === 1 && lastClub(s).drawRound === 2,
            `confirm ${redraws(c).length} 則，drawRound=${rounds(s)}`);
    }

    // 社團名稱只有不可見字元：要當成還沒有填社團名稱
    {
        const s = loadSite();
        const c = stubConfirm(s);
        addClub(s, TWO, { name: `${ZWSP}${String.fromCharCode(0x2060)} ` });
        R.check('重抽：社團名稱只有不可見字元時，提醒還沒有填社團名稱，不加入也不計次',
            clubs(s).length === 0 && s.alerts.length === 1 && /還沒有填社團名稱/.test(s.alerts[0]) && safe(s, 'clubDrawCounts.size') === 0 && c.messages.length === 0,
            `社團 ${clubs(s).length} 個，提醒：${s.alerts.join(' / ')}，計數表 ${safe(s, 'clubDrawCounts.size')} 筆`);
    }

    // 同名社團還在清單裡時，confirm 多一句提醒；已移除時不加
    {
        const s = loadSite();
        const c = stubConfirm(s);
        addClub(s, TWO, { name: '桌遊社' });
        addClub(s, TWO, { name: '桌遊社' });
        R.check('重抽：同名社團還在清單裡時，confirm 提醒兩份都會匯出', redraws(c).length === 1 && redraws(c)[0] === REDRAW_MSG(1, true) && clubs(s).length === 2,
            redraws(c).join(' / '));
    }

    // 匯出檔的「抽籤次數」列（drawInfoRows 的內容；實際寫入 Excel 的列位置在第 14 節檢查）
    {
        const s = loadSite();
        stubConfirm(s);
        addClub(s, TWO, { name: '桌遊社' });
        removeAt(s, 0);
        addClub(s, TWO, { name: '桌遊社' });
        const rowsOf = (club) => { try { return s.fn('drawInfoRows')(club); } catch (e) { return []; } };
        const roundRow = (club) => rowsOf(club).find(([l]) => l === '抽籤次數');
        const labels = rowsOf(lastClub(s)).map(([l]) => l);
        R.check('重抽：抽籤資訊有「抽籤次數」列，緊接在「抽籤時間」之後，值是「第 2 次（…）」',
            roundRow(lastClub(s))?.[1] === ROUND_VALUE(2) && labels.indexOf('抽籤次數') === labels.indexOf('抽籤時間') + 1,
            `${labels.join(',')}；${roundRow(lastClub(s))?.[1]}`);
        const legacy = { ...lastClub(s) };
        delete legacy.drawRound;
        R.check('重抽：舊的社團物件沒有 drawRound 時，抽籤次數寫「未記錄」', roundRow(legacy)?.[1] === '未記錄', String(roundRow(legacy)?.[1]));
        R.check('重抽：抽籤資訊的內容不含社團名稱或測試姓名', rowsOf(lastClub(s)).length > 0 && rowsOf(lastClub(s)).every(([l, v]) => noRoster(`${l}${v}`)));
    }

    // 計數只放在記憶體：localStorage、sessionStorage 都沒有寫入
    {
        const s = loadSite();
        stubConfirm(s);
        const writes = [];
        const session = new Map();
        s.ctx.sessionStorage = {
            getItem: k => (session.has(k) ? session.get(k) : null),
            setItem: (k, v) => { writes.push(['session', k, String(v)]); session.set(k, String(v)); },
            removeItem: k => session.delete(k),
            clear: () => session.clear(),
        };
        const realSet = s.ctx.localStorage.setItem;
        s.ctx.localStorage.setItem = (k, v) => { writes.push(['local', k, String(v)]); return realSet(k, v); };
        for (let i = 0; i < 3; i++) {
            addClub(s, TWO, { name: '桌遊社' });
            if (clubs(s).length) removeAt(s, 0);
        }
        addClub(s, TWO, { name: '桌遊社' });
        const counted = lastClub(s)?.drawRound === 4;
        const leaked = writes.filter(([, k, v]) => /drawRound|DrawCount|抽籤|次|桌遊社|測試生/.test(`${k}${v}`));
        R.check('重抽：抽了 4 次之後，計數沒有寫進 localStorage 或 sessionStorage', counted && leaked.length === 0 && session.size === 0 && [...s.storage.values()].every(v => !/drawRound|DrawCount|抽籤|測試生/.test(v)),
            `drawRound=${lastClub(s)?.drawRound}，可疑寫入 ${leaked.length} 筆，sessionStorage ${session.size} 筆，localStorage 鍵：${[...s.storage.keys()].join(',') || '（無）'}`);
    }
}

R.done();
