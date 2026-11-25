// 貼上模式：社團名稱結尾的時段後綴（processPastedClubsData → splitSlotSuffix）
// 只有名稱結尾的時段字詞或單獨的 A/B 才算時段；英文單字結尾（Yoga、Club）與「A班」不算。
// 一般模式課表用清掉後綴的名稱，寒暑假模式再加上「(A)」「(B)」。
// 測試資料的姓名全部是虛構的。
import { loadSite, makeReporter } from './harness.mjs';
const R = makeReporter('貼上名單：社團名稱的時段後綴');
const s = loadSite();
const processPasted = s.fn('processPastedClubsData');
const DAY = '週一';
const dayKey = s.fn('parseDays')(DAY)[0];

// [社團名稱, 一般模式的課表內容, 寒暑假模式的課表內容]
const CASES = [
    // 要辨識成時段
    ['桌遊A', '桌遊', '桌遊(A)'],
    ['桌遊 B', '桌遊', '桌遊(B)'],
    ['桌遊(A)', '桌遊', '桌遊(A)'],
    ['桌遊（B）', '桌遊', '桌遊(B)'],
    ['桌遊 (A)', '桌遊', '桌遊(A)'],
    ['桌遊b', '桌遊', '桌遊(B)'],
    ['足球一時段', '足球', '足球(A)'],
    ['足球二時段', '足球', '足球(B)'],
    ['足球第二時段', '足球', '足球(B)'],
    ['足球第一節', '足球', '足球(A)'],
    ['足球SlotA', '足球', '足球(A)'],
    ['足球 Slot B', '足球', '足球(B)'],
    ['Club(A)', 'Club', 'Club(A)'],
    ['桌遊-A', '桌遊', '桌遊(A)'],
    ['桌遊_B', '桌遊', '桌遊(B)'],
    ['桌遊 - A', '桌遊', '桌遊(A)'],
    ['桌遊Ａ', '桌遊', '桌遊(A)'],
    ['桌遊（Ｂ）', '桌遊', '桌遊(B)'],
    ['ＹｏｇＡ', 'ＹｏｇＡ', 'ＹｏｇＡ'],
    // 不可誤判
    ['瑜珈Yoga', '瑜珈Yoga', '瑜珈Yoga'],
    ['Zumba', 'Zumba', 'Zumba'],
    ['Club', 'Club', 'Club'],
    ['Samba', 'Samba', 'Samba'],
    ['LEGO Club', 'LEGO Club', 'LEGO Club'],
    ['棒球A班', '棒球A班', '棒球A班'],
    ['桌遊(A', '桌遊(A', '桌遊(A'],
    // 清完是空的就用原名，也不當成時段
    ['A', 'A', 'A'],
    ['-A', '-A', '-A'],
    // 單獨的「時段」沒有 A/B 資訊，維持原名
    ['足球時段', '足球時段', '足球時段'],
];

const makeClub = (clubName) => ({
    id: 1, day: DAY, clubName, quota: 12, gradeStart: 1, gradeEnd: 6, priority: 'desc', selectedCount: 1,
    students: [{ class: '301', seat: '05', name: '測試生301_05', selected: true }],
});

function scheduleOf(clubName, slotMode) {
    const { resultData } = processPasted([makeClub(clubName)], slotMode, true);
    if (resultData.length !== 1) throw new Error(`學生筆數 ${resultData.length}`);
    return resultData[0].schedule;
}

for (const [mode, col] of [['semester', 1], ['vacation', 2]]) {
    const modeName = mode === 'vacation' ? '寒暑假模式' : '一般模式';
    for (const c of CASES) {
        const expected = c[col];
        R.test(`${modeName}：「${c[0]}」→「${expected}」`, () => {
            const sch = scheduleOf(c[0], mode);
            const ok = Object.keys(sch).length === 1 && sch[dayKey] === expected;
            return [ok, `課表 ${JSON.stringify(sch)}`];
        });
    }
}

// 同一個社團的 A、B 時段：一般模式合併成同一個名稱，寒暑假模式分開標示
R.test('一般模式：「桌遊A」「桌遊B」同一天只列一次「桌遊」', () => {
    const { resultData } = processPasted([makeClub('桌遊A'), makeClub('桌遊B')], 'semester', true);
    const sch = resultData[0]?.schedule || {};
    return [sch[dayKey] === '桌遊', `課表 ${JSON.stringify(sch)}`];
});
R.test('寒暑假模式：「桌遊A」「桌遊B」分別列成「桌遊(A)」「桌遊(B)」', () => {
    const { resultData } = processPasted([makeClub('桌遊A'), makeClub('桌遊B')], 'vacation', true);
    const sch = resultData[0]?.schedule || {};
    return [sch[dayKey] === '桌遊(A), 桌遊(B)', `課表 ${JSON.stringify(sch)}`];
});

// 沒抽中的學生不列入
R.test('selected: false 的學生不列入課表', () => {
    const club = makeClub('桌遊A');
    club.students.push({ class: '302', seat: '11', name: '測試生302_11', selected: false });
    const { resultData } = processPasted([club], 'semester', true);
    return [resultData.length === 1 && resultData[0].name === '測試生301_05', `學生 ${resultData.length} 筆`];
});

R.done();
