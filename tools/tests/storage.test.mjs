// 瀏覽器設定（localStorage 的 yenping_club_config）：只存「欄位／工作表 → 社團 → 星期」，
// 數量與鍵長度有上限、超過時刪最舊的、相容舊格式、略過格式不符的值。
// 社團與工作表名稱都是虛構的。
import { loadSite, makeReporter, makeWorkbook } from './harness.mjs';
const R = makeReporter('瀏覽器設定：上限、格式與相容性');
const KEY = 'yenping_club_config';
const stored = (s) => JSON.parse(s.storage.get(KEY) || '{}');
const singleRows = (names, slotCol = '社團', day = '週二') => names.map(n => ({ key: `${slotCol}\u0001${n}`, original: n, slotCol, editedName: n, day }));
const multiRows = (names, day = '週三') => names.map(n => ({ key: n, original: n, slotCol: '', editedName: n, day }));
const range = (a, b, p) => Array.from({ length: b - a }, (_, i) => `${p}${a + i}`);

// ---- 單一時段欄位最多 200 個社團，超過從最舊的刪
{
    const s = loadSite();
    const save = s.run('saveMappingConfig');
    save('single', singleRows(range(0, 200, '社團甲')));
    save('single', singleRows(['社團甲0'])); // 更新最舊的一筆 → 移到最後
    save('single', singleRows(range(200, 250, '社團甲')));
    const slot = stored(s)['社團'] || {};
    const keys = Object.keys(slot);
    R.check('每個時段欄位最多 200 個社團', keys.length === 200, `${keys.length} 個`);
    R.check('最近更新過的舊社團保留', '社團甲0' in slot);
    R.check('最舊且沒有更新的社團被刪除', !('社團甲1' in slot) && !('社團甲50' in slot));
    R.check('最新的社團保留在最後', '社團甲249' in slot && keys.at(-1) === '社團甲249', keys.at(-1));
}
// ---- 最上層最多 100 個鍵
{
    const s = loadSite();
    const save = s.run('saveMappingConfig');
    save('multi', multiRows(range(0, 100, '社團-工作表')));
    save('multi', multiRows(['社團-工作表0']));
    save('multi', multiRows(range(100, 150, '社團-工作表')));
    const cfg = stored(s);
    R.check('最上層最多 100 個鍵', Object.keys(cfg).length === 100, `${Object.keys(cfg).length} 個`);
    R.check('最上層：最近更新過的鍵保留', '社團-工作表0' in cfg);
    R.check('最上層：最舊的鍵被刪除', !('社團-工作表1' in cfg));
}
// ---- 已滿時存入純數字鍵：剛存的鍵不可被立刻淘汰（物件的整數鍵會排到最前面）
{
    const s = loadSite();
    const save = s.run('saveMappingConfig');
    save('multi', multiRows(range(0, 100, '社團-工作表')));
    save('multi', multiRows(['101']));
    const cfg = stored(s);
    R.check('最上層已滿時存入純數字鍵「101」：剛存的鍵保留', cfg['101'] === '週三', JSON.stringify(cfg['101']));
    R.check('最上層已滿時存入純數字鍵：總數仍是 100', Object.keys(cfg).length === 100, `${Object.keys(cfg).length}`);
    R.check('最上層已滿時存入純數字鍵：改刪最舊的文字鍵', !('社團-工作表0' in cfg) && '社團-工作表1' in cfg);
}
{
    const s = loadSite();
    const save = s.run('saveMappingConfig');
    save('single', singleRows(range(0, 200, '社團甲')));
    save('single', singleRows(['301']));
    const slot = stored(s)['社團'] || {};
    R.check('時段欄位已滿時存入純數字社團名：剛存的保留、總數仍是 200', slot['301'] === '週二' && Object.keys(slot).length === 200, `${Object.keys(slot).length}`);
}
// ---- 只存名稱與星期
{
    const s = loadSite();
    s.run('saveMappingConfig')('single', singleRows(['桌遊社']));
    const raw = s.storage.get(KEY);
    R.check('存的是 { 時段欄位: { 社團: 星期 } }，沒有時間戳或其他欄位', raw === JSON.stringify({ '社團': { '桌遊社': '週二' } }), raw);
}
// ---- 時段欄位（物件）與工作表名稱（字串）同名
{
    const s = loadSite();
    const save = s.run('saveMappingConfig');
    save('single', singleRows(['桌遊社'], '社團'));
    save('multi', multiRows(['社團']));
    const cfg = stored(s);
    R.check('同名的工作表不會用字串蓋掉時段欄位的社團設定', cfg['社團'] && typeof cfg['社團'] === 'object' && cfg['社團']['桌遊社'] === '週二', JSON.stringify(cfg['社團']));
}
// ---- 鍵長度上限 60 字
{
    const s = loadSite();
    const save = s.run('saveMappingConfig');
    const long = '長'.repeat(61);
    const ok60 = '短'.repeat(60);
    save('multi', multiRows([long, ok60]));
    save('single', singleRows([long, '桌遊社'], '社團'));
    save('single', singleRows(['直笛社'], long));
    const cfg = stored(s);
    R.check('超過 60 字的工作表名稱不存', !(long in cfg));
    R.check('60 字的工作表名稱照存', cfg[ok60] === '週三');
    R.check('超過 60 字的社團名稱不存、同批的其他社團照存', cfg['社團'] && !(long in cfg['社團']) && cfg['社團']['桌遊社'] === '週二', JSON.stringify(cfg['社團']));
    R.check('超過 60 字的時段欄位不存', !Object.keys(cfg).some(k => k.length > 60));
    s.storage.set(KEY, JSON.stringify({ [long]: '週一', '社團': { [long]: '週一', '桌遊社': '週二' } }));
    const loaded = s.run('loadLocalStorageConfig()');
    R.check('讀取時略過超過 60 字的鍵', !(long in loaded) && !(long in loaded['社團']) && loaded['社團']['桌遊社'] === '週二');
}
// ---- 相容舊格式（最上層「社團 → 星期」）與損毀內容
{
    const s = loadSite();
    s.storage.set(KEY, JSON.stringify({ '桌遊社': '週四' }));
    s.ctx.__wb = makeWorkbook([{ name: '總表', rows: [['班級', '座號', '姓名', '社團'], ['301', '05', '測試生301_05', '桌遊社']] }]);
    s.run("uploadedWorkbook = __wb; document.getElementById('mode-select').value = 'single'; analyzeAndRenderMapping();");
    const select = s.el('mapping-table-body').querySelector('.day-select');
    const opt = select?.children.find(o => o.selected);
    R.check('舊格式的星期設定仍會套用到對照表', opt?.value === '週四', opt?.value);
    s.storage.set(KEY, '{"社團": 5, "x": [1,2], "y": {"a": 1, "b": "週一"}}');
    const cfg = s.run('loadLocalStorageConfig()');
    R.check('讀取時略過格式不符的值（數字、陣列、非字串的星期）', !('社團' in cfg) && !('x' in cfg) && cfg.y && cfg.y.b === '週一' && !('a' in cfg.y));
}
R.done();
