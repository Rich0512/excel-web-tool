// 貼上模式的介面用語要符合實際流程：按「加入這個社團」時就抽籤（loadPastedClub → runClubLottery），
// 下載按鈕（processPastedData）只匯出已抽好的結果，不會再抽一次。
// 舊版寫成「下載時一起抽籤」，會讓使用者以為移除重貼沒有關係、清單上的結果下載時還會變。
// 只讀 docs/ 的原始碼做文字比對，不含任何名冊資料。
import { readdirSync } from 'node:fs';
import { makeReporter, readSiteFile, JS_DIR } from './harness.mjs';
const R = makeReporter('介面用語：加入時抽籤、下載只匯出');

const DOWNLOAD_LABEL = '下載課表與抽籤結果';
const html = readSiteFile('index.html');
const jsFiles = readdirSync(JS_DIR).filter(f => f.endsWith('.js')).sort();
const sources = [
    ['index.html', html],
    ['404.html', readSiteFile('404.html')],
    ...jsFiles.map(f => [`js/${f}`, readSiteFile(`js/${f}`)]),
];
const app = readSiteFile('js/app.js');

// ---- 1. 不再出現「下載時才抽籤」的錯誤描述（index.html、404.html、docs/js/*.js，含註解）
const WRONG = ['一起抽籤', '抽籤並下載', '抽籤或匯出', '下載時抽籤', '下載時才抽籤', '下載後才抽籤', '按下載後，已加入的社團'];
for (const phrase of WRONG) {
    const hits = sources.filter(([, src]) => src.includes(phrase)).map(([name]) => name);
    R.check(`沒有「${phrase}」`, hits.length === 0, hits.length ? `出現在 ${hits.join('、')}` : '0 個檔案');
}
R.check(`檢查範圍包含 docs/js 的 ${jsFiles.length} 個檔案`, jsFiles.includes('app.js') && jsFiles.length >= 6, jsFiles.join('、'));

// ---- 2. 下載按鈕的文字，且仍包含筆數 #total-pasted-students
const btn = (html.match(/<button\b[^>]*\bid="btn-process-paste"[^>]*>([\s\S]*?)<\/button>/) || [])[1];
R.check('找得到 #btn-process-paste', btn !== undefined);
const btnText = (btn || '').replace(/\s+/g, ' ').trim();
R.check(`下載按鈕是「${DOWNLOAD_LABEL}（共 N 筆錄取資料）」`,
    btnText === `${DOWNLOAD_LABEL}（共 <span id="total-pasted-students">0</span> 筆錄取資料）`, btnText);
R.check('#total-pasted-students 只出現一次（在下載按鈕裡）', (html.match(/id="total-pasted-students"/g) || []).length === 1);

// 處理中暫時替換按鈕文字、結束後還原 innerHTML（筆數的 <span> 才會回來），再重繪清單
{
    // 檔案可能是 CRLF 換行
    const fn = (app.match(/async function processPastedData\(\)[\s\S]*?\r?\n}\r?\n/) || [''])[0];
    R.check('processPastedData 找的按鈕就是 #btn-process-paste（#loaded-clubs-section 裡唯一的 .btn-primary）', (() => {
        const section = (html.match(/<div id="loaded-clubs-section"[\s\S]*?id="btn-process-paste"[\s\S]*?<\/button>/) || [''])[0];
        return /querySelector\('#loaded-clubs-section \.btn-primary'\)/.test(fn) && (section.match(/btn-primary/g) || []).length === 1;
    })());
    R.check(`處理中的按鈕文字是「正在${DOWNLOAD_LABEL}…」`, fn.includes(`actionBtn.textContent = '正在${DOWNLOAD_LABEL}…'`), (fn.match(/actionBtn\.textContent = '[^']*'/) || ['找不到'])[0]);
    R.check('處理前先記下按鈕的 innerHTML', /const origBtnHtml = actionBtn \? actionBtn\.innerHTML : ""/.test(fn));
    R.check('finally 先還原按鈕 innerHTML 再 renderLoadedClubs()', /finally\s*{[\s\S]*actionBtn\.innerHTML = origBtnHtml;[\s\S]*renderLoadedClubs\(\);/.test(fn));
    R.check('processPastedData 不會抽籤（沒有呼叫 runClubLottery）', fn.length > 0 && !/runClubLottery\(/.test(fn));
}

// 提示文字中用「」引用下載按鈕時，名稱要和按鈕相同（同一個動作同一個名稱）
{
    const refs = [...app.matchAll(/「([^」]*下載[^」]*)」/g)].map(m => m[1]).filter(t => t.includes('抽籤'));
    const bad = refs.filter(t => t !== DOWNLOAD_LABEL);
    R.check(`app.js 引用下載按鈕時都寫「${DOWNLOAD_LABEL}」`, refs.length >= 1 && bad.length === 0, `引用 ${refs.length} 次${bad.length ? '，不一致：' + bad.join('、') : ''}`);
}

// ---- 3. 「加入這個社團」旁邊說明加入時就會抽籤，以及重新抽籤的方法
{
    const m = html.match(/<button\b[^>]*\bid="btn-load-club"[^>]*>加入這個社團<\/button>([\s\S]{0,400})/);
    R.check('找得到「加入這個社團」按鈕', !!m);
    const btnTag = m ? m[0].slice(0, m[0].indexOf('>') + 1) : '';
    const tipId = (btnTag.match(/aria-describedby="([^"]+)"/) || [])[1];
    const tip = tipId ? (html.match(new RegExp(`<p class="field-tip" id="${tipId}">([^<]*)</p>`)) || [])[1] : undefined;
    R.check('按鈕以 aria-describedby 連到說明（field-tip）', !!tip, tipId ? `#${tipId}` : '沒有 aria-describedby');
    R.check('說明緊接在按鈕後面', !!m && !!tipId && m[1].includes(`id="${tipId}"`));
    R.check('說明寫「加入時就會抽籤」', !!tip && tip.includes('加入時就會抽籤'), tip);
    R.check('說明寫出重新抽籤要先移除再加入，且會記錄抽籤次數', !!tip && /移除/.test(tip) && /再加入/.test(tip) && tip.includes('抽籤次數'), tip);
}
{
    const intro = (html.match(/<div id="tab-paste-content"[\s\S]*?<p class="tab-intro">([^<]*)<\/p>/) || [])[1] || '';
    R.check('貼上分頁的說明寫加入時就會抽籤', intro.includes('「加入這個社團」') && intro.includes('加入時就會抽籤'), intro);
}

// ---- 4. 下載前的提醒：下載只匯出已抽好的結果、不會重新抽籤
{
    const section = (html.match(/<div id="loaded-clubs-section"[\s\S]*?id="btn-process-paste"/) || [''])[0];
    const callout = (section.match(/<div class="warning-callout">[\s\S]*?<div>([\s\S]*?)<\/div>/) || [])[1] || '';
    R.check('下載前的提醒寫「不會重新抽籤」', callout.includes('已抽好的結果') && callout.includes('不會重新抽籤'), callout.replace(/<[^>]+>/g, ''));
}

// 課表範例的說明：晨間社團預設放在獨立的「晨間(…)」欄位（MORNING_OPTIONS），不是填入每一個上課日
{
    const note = (html.match(/<p class="sample-note">([\s\S]*?)<\/p>/) || [])[1] || '';
    R.check('課表範例說明沒有寫「晨間社團會填入每一個上課日」', note !== '' && !/晨間社團會填入每一個上課日/.test(note), note);
    R.check('課表範例說明寫出晨間社團另外列在晨間欄位', note.includes('晨間社團預設另外列在「晨間(週一、週四)」這類晨間欄位'), note);
}

R.done();
