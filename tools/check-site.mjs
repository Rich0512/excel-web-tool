#!/usr/bin/env node
// ============================================
// GitHub Pages 部署前檢查（網站根目錄為 docs/）
// 用法：node tools/check-site.mjs [--require-bump[=<git ref>]]
//   --require-bump：docs/ 相對 <git ref>（預設 HEAD）有變更時，要求 CACHE_VERSION 也必須已調升
// 有錯誤時 exit 1；只有警告時 exit 0
// ============================================

import { readFileSync, existsSync, statSync, readdirSync } from 'node:fs';
import { join } from 'node:path';
import { ROOT, SITE, BASE_PATH, TEMPLATE_DIR, SPREADSHEET_RE, git, gitList, versionOf, versionAt, isForbiddenSpreadsheet } from './site-config.mjs';
import { inspectTemplate } from './template-check.mjs';

const PRECACHE_WARN_BYTES = 5 * 1024 * 1024;

const errors = [];
const warnings = [];
const read = (rel) => readFileSync(join(ROOT, rel), 'utf8');
const site = (rel) => `${SITE}/${rel}`;
const listSite = (dir, ext) => readdirSync(join(ROOT, SITE, dir))
    .filter(f => f.endsWith(ext))
    .map(f => site(dir === '.' ? f : `${dir}/${f}`));

const htmlFiles = listSite('.', '.html');
const jsFiles = listSite('js', '.js');
const cssFiles = [...listSite('css', '.css'), ...listSite('fonts', '.css')];

// ==========================================
// 1. 版本號一致：service-worker.js 與 index.html、404.html 的 ?v=
// ==========================================
const sw = read(site('service-worker.js'));
const swVersion = versionOf(sw);
const indexHtml = read(site('index.html'));
// 頁面引用的 JS/CSS（相對 docs/ 的路徑，含查詢參數）；404.html 用 /excel-web-tool/ 開頭的絕對路徑
const PAGES = { 'index.html': indexHtml, '404.html': read(site('404.html')) };
// 單、雙引號都認；./js/… 與 /excel-web-tool/js/… 都換成 js/… 再比對
const pageRefs = (text) => [...text.matchAll(/\b(?:src|href)=(["'])([^"']+\.(?:js|css)(?:\?[^"']*)?)\1/g)]
    .map(m => m[2].startsWith(BASE_PATH) ? m[2].slice(BASE_PATH.length) : m[2].replace(/^\.\//, ''))
    .filter(ref => /^(?:js|css|fonts)\//.test(ref));
if (!swVersion) {
    errors.push('service-worker.js 找不到 CACHE_VERSION');
} else {
    for (const [name, text] of Object.entries(PAGES)) {
        const queryVersions = [...text.matchAll(/\?v=([\d.]+)/g)].map(m => m[1]);
        const mismatched = [...new Set(queryVersions.filter(v => v !== swVersion))];
        if (mismatched.length) {
            errors.push(`${name} 的 ?v= (${mismatched.join(', ')}) 與 CACHE_VERSION (v${swVersion}) 不一致`);
        }
        // 自己的 JS/CSS 都要帶版本參數，Service Worker 才能以相同網址預快取
        for (const ref of pageRefs(text)) {
            if (!/\?v=[\d.]+$/.test(ref)) errors.push(`${name} 的 ${ref} 缺少 ?v= 版本參數`);
        }
    }
}

// ==========================================
// 2. 預快取清單：檔案都要存在（cache.addAll 只要缺一個就整個安裝失敗）
// ==========================================
const assetBlock = (sw.match(/const STATIC_ASSETS = \[([\s\S]*?)\];/) || [])[1] || '';
const assets = [...assetBlock.matchAll(/'([^']+)'/g)].map(m => m[1]);
// 預快取網址可能帶 ?v=（與 index.html 相同），檢查檔案時去掉查詢參數
const assetPath = (asset) => asset.replace(/\?.*$/, '');
const assetPaths = assets.map(assetPath);
let precacheBytes = 0;
for (const asset of assets) {
    if (asset === './') continue;
    const rel = site(assetPath(asset).replace(/^\.\//, ''));
    if (!existsSync(join(ROOT, rel))) {
        errors.push(`service-worker.js 預快取的檔案不存在：${asset}`);
        continue;
    }
    precacheBytes += statSync(join(ROOT, rel)).size;
    if (/\.ttf$/i.test(rel)) errors.push(`預快取不應包含 TTF（瀏覽器只會用 woff2）：${asset}`);
    const v = (asset.match(/\?v=([\d.]+)$/) || [])[1];
    if (v && swVersion && v !== swVersion) errors.push(`service-worker.js 預快取的 ${asset} 與 CACHE_VERSION (v${swVersion}) 不一致`);
}
for (const dup of new Set(assetPaths.filter((p, i) => assetPaths.indexOf(p) !== i))) {
    errors.push(`service-worker.js 的 STATIC_ASSETS 重複列出 ${dup}（帶與不帶 ?v= 各一次會下載兩次）`);
}
for (const file of [...jsFiles, ...cssFiles]) {
    const asset = `./${file.slice(SITE.length + 1)}`;
    if (!assetPaths.includes(asset)) errors.push(`${file} 不在 service-worker.js 的 STATIC_ASSETS，離線時會載不到`);
}
// 頁面以 ?v= 引用的檔案，預快取網址要完全相同，否則第一次比對不到、會再下載一份
for (const [name, text] of Object.entries(PAGES)) {
    for (const ref of pageRefs(text).filter(r => /\?v=[\d.]+$/.test(r))) {
        if (!assets.includes(`./${ref}`)) {
            errors.push(`${name} 引用 ${ref}，但 STATIC_ASSETS 沒有完全相同的網址（請把 ?v= 也寫進預快取清單）`);
        }
    }
}
// ExcelJS 由 utils.js 以 ?v=<版本> 動態載入，預快取也要帶同一個版本，升級時才不會拿到舊版 SW 快取的舊檔
if (swVersion && !assets.includes(`./js/libs/exceljs.min.js?v=${swVersion}`)) {
    errors.push(`STATIC_ASSETS 缺少 ./js/libs/exceljs.min.js?v=${swVersion}（離線時無法讀取 Excel）`);
}
// 字型：介面檔與 Outfit 要預快取（首屏就用到）；常用字檔（-ext）不可預快取，畫面用到才下載
for (const font of readdirSync(join(ROOT, SITE, 'fonts')).filter(f => f.endsWith('.woff2'))) {
    const asset = `./fonts/${font}`;
    const isExt = /-ext\.woff2$/.test(font);
    if (isExt && assetPaths.includes(asset)) errors.push(`STATIC_ASSETS 不應包含常用字檔 ${asset}（約 0.8 MB，畫面用到才下載）`);
    if (!isExt && !assetPaths.includes(asset)) errors.push(`${asset} 不在 STATIC_ASSETS，離線時介面會改用系統字型`);
}
// preload 的檔案要存在且已預快取
for (const m of indexHtml.matchAll(/<link\b[^>]*\brel="preload"[^>]*>/g)) {
    const href = (m[0].match(/\bhref="([^"]+)"/) || [])[1];
    if (!href) continue;
    if (!existsSync(join(ROOT, site(href.replace(/\?.*$/, ''))))) errors.push(`index.html 預載的檔案不存在：${href}`);
    else if (!assets.includes(`./${href}`)) errors.push(`index.html 預載的 ${href} 不在 STATIC_ASSETS（網址要完全相同）`);
}
if (precacheBytes > PRECACHE_WARN_BYTES) {
    warnings.push(`預快取總量 ${(precacheBytes / 1048576).toFixed(1)} MB，超過 ${PRECACHE_WARN_BYTES / 1048576} MB，手機首次開啟會很慢`);
}

// ==========================================
// 3. 不可引用外部資源（CSP 為 'self'，且個資頁面不應連外）
// ==========================================
const ownJsFiles = jsFiles.filter(f => !f.includes('/libs/'));
for (const file of [...htmlFiles, ...cssFiles, ...ownJsFiles]) {
    const text = read(file);
    const external = [
        ...text.matchAll(/<(?:script|img|iframe|link)\b[^>]*\b(?:src|href)=["'](https?:)?\/\/[^"']+["'][^>]*>/gi),
        ...text.matchAll(/url\(\s*["']?(?:https?:)?\/\/[^)]+\)/gi),
        ...text.matchAll(/@import\s+["']?(?:url\()?["']?(?:https?:)?\/\//gi),
        ...text.matchAll(/\bfetch\(\s*["'`](?:https?:)?\/\//gi),
    ].map(m => m[0]).filter(tag => !/rel=["']canonical["']/i.test(tag));
    for (const tag of external) errors.push(`${file} 引用了外部資源：${tag.slice(0, 120)}`);
}

// ==========================================
// 4. 不可有 inline script / style / 事件屬性（CSP 未開放 'unsafe-inline'）
// ==========================================
for (const file of htmlFiles) {
    const text = read(file);
    if (/<script(?![^>]*\bsrc=)[^>]*>/i.test(text)) errors.push(`${file} 含 inline <script>，請移到 js/`);
    if (/<style\b/i.test(text)) errors.push(`${file} 含 <style> 區塊，請移到 css/`);
    for (const m of text.matchAll(/<[^>]*\s(style|on[a-z]+)\s*=[^>]*>/gi)) {
        errors.push(`${file} 含 inline ${m[1]} 屬性：${m[0].slice(0, 80)}`);
    }
}
for (const file of ownJsFiles) {
    read(file).split('\n').forEach((line, i) => {
        if (/<[a-z][^>]*\s(style|on[a-z]+)\s*=/i.test(line) || /setAttribute\(\s*['"](style|on[a-z]+)['"]/i.test(line)) {
            errors.push(`${file}:${i + 1} 產生 inline style／事件屬性，會被 CSP 擋下，請改用 class 或 addEventListener`);
        }
        // CSP 也沒有開放 'unsafe-eval'（註解行略過）
        if (!/^\s*(\/\/|\*|\/\*)/.test(line) && /\beval\s*\(|\bnew\s+Function\s*\(|\bFunction\s*\(\s*['"`]|\bset(Timeout|Interval)\s*\(\s*['"`]/.test(line)) {
            errors.push(`${file}:${i + 1} 從字串產生程式碼（eval／new Function／字串形式的 setTimeout），會被 CSP 擋下`);
        }
    });
}
// CSP 必須與預期的政策完全相同（任何放寬，例如 'unsafe-eval'、https:、*、data: 都會被抓到）；
// 要調整政策時，先確認理由，再同步修改這裡的預期值。
// ExcelJS 由 loadExcelJS() 延遲載入，不可用 <script> 標籤。
const EXPECTED_CSP = {
    'index.html': "default-src 'self'; script-src 'self'; style-src 'self'; img-src 'self' data: blob:; font-src 'self'; connect-src 'self'; worker-src 'self'; object-src 'none'; base-uri 'self'; form-action 'none'",
    '404.html': "default-src 'self'; script-src 'none'; style-src 'self'; img-src 'self'; font-src 'self'; object-src 'none'; base-uri 'self'; form-action 'none'",
};
const cspDirectives = (csp) => new Set(csp.split(';').map(d => d.trim().replace(/\s+/g, ' ')).filter(Boolean));
for (const file of htmlFiles) {
    const text = read(file);
    const name = file.slice(SITE.length + 1);
    const csp = (text.match(/http-equiv="Content-Security-Policy"\s+content="([^"]*)"/i) || [])[1];
    if (csp === undefined) {
        errors.push(`${file} 沒有 Content-Security-Policy <meta>`);
    } else if (!EXPECTED_CSP[name]) {
        errors.push(`${file} 沒有預期的 CSP，請在 check-site.mjs 的 EXPECTED_CSP 加入`);
    } else {
        const actual = cspDirectives(csp);
        const expected = cspDirectives(EXPECTED_CSP[name]);
        const added = [...actual].filter(d => !expected.has(d));
        const removed = [...expected].filter(d => !actual.has(d));
        if (added.length || removed.length) {
            errors.push(`${file} 的 CSP 與預期不同：${[...added.map(d => `多了「${d}」`), ...removed.map(d => `少了「${d}」`)].join('；')}`);
        }
    }
    if (/<script\b[^>]*\bsrc=["'][^"']*js\/libs\/exceljs/i.test(text)) {
        errors.push(`${file} 用 <script> 載入 ExcelJS，會擋住後面的程式；請改用 utils.js 的 loadExcelJS()`);
    }
}

// ==========================================
// 5. 404.html 必須用絕對路徑（GitHub Pages 會在任意深度的路徑回傳它）
// ==========================================
for (const m of read(site('404.html')).matchAll(/\b(?:src|href)=["']([^"']+)["']/g)) {
    const url = m[1];
    if (!url.startsWith(BASE_PATH) && !/^(https?:|#|mailto:|data:)/.test(url)) {
        errors.push(`404.html 使用相對路徑 "${url}"，請改為 ${BASE_PATH} 開頭`);
    }
}

// ==========================================
// 6. 試算表只能存在於 docs/templates/（防止真實學生名冊上傳）
// ==========================================
const candidates = new Set([
    ...gitList(['ls-files']),
    ...gitList(['diff', '--cached', '--name-only']),
    ...gitList(['ls-files', '--others', '--exclude-standard']),
]);
for (const file of candidates) {
    if (isForbiddenSpreadsheet(file)) errors.push(`試算表不在 ${TEMPLATE_DIR} 內，可能是真實名冊：${file}`);
}

// ==========================================
// 7. 介面用字都在 Noto Sans TC 的介面檔（-ui）內，且 fonts.css 的 unicode-range 與字元清單相符
//    每個字重拆成 -ui（首屏）與 -ext（其餘常用字，畫面出現才下載），清單由 tools/subset-fonts.py 產生
// ==========================================
const fontLists = { ui: 'tools/subset-chars-ui.txt', ext: 'tools/subset-chars-ext.txt' };
const FONT_WEIGHTS = { Regular: '400', Bold: '700' };
// 與 subset-fonts.py 收字的條件相同（Python 的 str.isprintable()）；ASCII 一律在介面檔內
const isUiChar = (ch) => ch > '~' && /^[^\p{C}\p{Z}]$/u.test(ch);
const parseUnicodeRange = (value) => {
    const chars = new Set();
    for (const token of value.split(',').map(s => s.trim())) {
        const m = token.match(/^U\+([0-9A-F]{1,6})(?:-([0-9A-F]{1,6}))?$/i);
        if (!m) return null;
        const end = parseInt(m[2] ?? m[1], 16);
        for (let cp = parseInt(m[1], 16); cp <= end; cp++) chars.add(String.fromCodePoint(cp));
    }
    return chars;
};
const missingLists = Object.values(fontLists).filter(f => !existsSync(join(ROOT, f)));
if (missingLists.length) {
    errors.push(`${missingLists.join('、')} 不存在，無法檢查字型涵蓋（請執行 /subset-fonts）`);
} else {
    const listed = { ui: new Set(read(fontLists.ui)), ext: new Set(read(fontLists.ext)) };
    const overlap = [...listed.ui].filter(ch => listed.ext.has(ch));
    if (overlap.length) errors.push(`字元清單 ui 與 ext 有 ${overlap.length} 個重複字，unicode-range 會互相衝突（請執行 /subset-fonts）`);

    // (1) 介面用字：在 ext 會讓首屏多下載常用字檔；兩邊都沒有則會用系統字型顯示
    const inExt = new Set();
    const missing = new Set();
    // CSS 只看 content: 字串（畫面上會顯示的字），註解不算；規則與 subset-fonts.py 相同
    const cssContentText = cssFiles.map(f => [...read(f).matchAll(/\bcontent:\s*(["'])(.*?)\1/g)].map(m => m[2]).join('')).join('');
    for (const text of [...[...htmlFiles, site('manifest.json'), ...ownJsFiles].map(read), cssContentText]) {
        for (const ch of text) {
            if (isUiChar(ch) && !listed.ui.has(ch)) (listed.ext.has(ch) ? inExt : missing).add(ch);
        }
    }
    if (inExt.size) {
        errors.push(`${inExt.size} 個介面用字不在字型介面檔、只在常用字檔，首屏會多下載常用字檔：${[...inExt].join('')}（請執行 /subset-fonts）`);
    }
    if (missing.size) {
        errors.push(`字型未涵蓋 ${missing.size} 個介面用字：${[...missing].join('')}（請執行 /subset-fonts；重新裁切後仍缺的字是原始字型沒有，請改用其他字）`);
    }

    // (2) fonts.css：Noto Sans TC 的每個 @font-face 都要是拆分後的檔案；
    //     ui 的 unicode-range 等於 ui 清單，ext 的涵蓋 ext 清單且不含 ui 字
    const fontsCss = site('fonts/fonts.css');
    const seen = new Set();
    for (const [, body] of read(fontsCss).matchAll(/@font-face\s*\{([^}]*)\}/g)) {
        if (!/font-family:\s*['"]Noto Sans TC['"]/.test(body)) continue;
        const url = (body.match(/url\(\s*['"]?([^'")]+)['"]?\s*\)/) || [])[1] || '';
        const m = url.match(/^NotoSansTC-(Regular|Bold)-(ui|ext)\.woff2$/);
        if (!m) {
            errors.push(`${fontsCss} 的 Noto Sans TC 使用未拆分的字型檔：${url || '（沒有 src）'}（請執行 /subset-fonts）`);
            continue;
        }
        const [, weight, part] = m;
        seen.add(`${weight}-${part}`);
        if (!existsSync(join(ROOT, site(`fonts/${url}`)))) errors.push(`${fontsCss} 引用的字型檔不存在：${url}`);
        const declaredWeight = (body.match(/font-weight:\s*(\d+)/) || [])[1];
        if (declaredWeight !== FONT_WEIGHTS[weight]) errors.push(`${fontsCss} 的 ${url} font-weight 應為 ${FONT_WEIGHTS[weight]}`);
        const rangeValue = (body.match(/unicode-range:\s*([^;]+)/) || [])[1];
        const range = rangeValue && parseUnicodeRange(rangeValue);
        if (!range) {
            errors.push(`${fontsCss} 的 ${url} 缺少 unicode-range 或格式無法解析（請執行 /subset-fonts）`);
            continue;
        }
        const lacking = [...listed[part]].filter(ch => !range.has(ch)).length;
        if (part === 'ui') {
            // ui 範圍要精確：多出的字會在介面檔裡找不到字形
            const extra = [...range].filter(ch => !listed.ui.has(ch)).length;
            if (extra || lacking) {
                errors.push(`${fontsCss} 的 ${url} unicode-range 與 ${fontLists.ui} 不一致（多 ${extra} 字、少 ${lacking} 字），請執行 /subset-fonts`);
            }
        } else {
            // ext 是粗範圍：可以多含碼位，但要涵蓋 ext 清單，且不可含 ui 字（否則介面字會觸發下載 ext）
            const uiInRange = [...listed.ui].filter(ch => range.has(ch)).length;
            if (lacking) errors.push(`${fontsCss} 的 ${url} unicode-range 少了 ${fontLists.ext} 的 ${lacking} 個字，這些字會改用系統字型（請執行 /subset-fonts）`);
            if (uiInRange) errors.push(`${fontsCss} 的 ${url} unicode-range 含 ${uiInRange} 個介面檔的字，首屏會多下載常用字檔（請執行 /subset-fonts）`);
        }
    }
    for (const weight of Object.keys(FONT_WEIGHTS)) {
        for (const part of Object.keys(fontLists)) {
            if (!seen.has(`${weight}-${part}`)) errors.push(`${fontsCss} 缺少 NotoSansTC-${weight}-${part}.woff2 的 @font-face（請執行 /subset-fonts）`);
        }
    }
}

// ==========================================
// 8. 範本只含虛構姓名（「姓名」欄下方都要符合 測試生<班級>_<座號>）
//    錯誤訊息只列儲存格位置，不列內容，以免真實姓名出現在輸出中
// ==========================================
for (const file of readdirSync(join(ROOT, TEMPLATE_DIR)).filter(f => SPREADSHEET_RE.test(f))) {
    const rel = `${TEMPLATE_DIR}${file}`;
    if (!/\.xlsx$/i.test(file)) {
        errors.push(`${rel} 不是 xlsx，無法確認是否為假資料；範本請存成 .xlsx`);
        continue;
    }
    try {
        const { checked, bad } = inspectTemplate(join(ROOT, rel));
        if (!checked) errors.push(`${rel} 找不到任何「姓名」欄的資料，無法確認是否為假資料`);
        if (bad.length) {
            errors.push(`${rel} 有 ${bad.length} 個姓名不是「測試生<班級>_<座號>」格式，可能是真實名冊：${bad.slice(0, 5).join(', ')}${bad.length > 5 ? '…' : ''}`);
        }
    } catch (e) {
        errors.push(`${rel} 無法解析（${e.message}），無法確認是否為假資料`);
    }
}

// ==========================================
// 9. manifest 圖示存在、是 PNG，且實際像素與宣告的 sizes 一致
// ==========================================
for (const icon of JSON.parse(read(site('manifest.json'))).icons || []) {
    const rel = site(icon.src.replace(/^\.\//, ''));
    if (!existsSync(join(ROOT, rel))) {
        errors.push(`manifest.json 的圖示不存在：${icon.src}`);
        continue;
    }
    const png = readFileSync(join(ROOT, rel));
    if (png.length < 24 || png.readUInt32BE(0) !== 0x89504e47) {
        errors.push(`manifest.json 的圖示不是 PNG：${icon.src}`);
        continue;
    }
    const actual = `${png.readUInt32BE(16)}x${png.readUInt32BE(20)}`;
    if (!String(icon.sizes || '').split(/\s+/).includes(actual)) {
        errors.push(`manifest.json 的圖示 ${icon.src} 宣告 ${icon.sizes}，實際是 ${actual}`);
    }
}

// ==========================================
// 10. （選用）docs/ 有變更時，快取版本必須調升
// ==========================================
const bumpArg = process.argv.find(a => a.startsWith('--require-bump'));
if (bumpArg && swVersion) {
    const baseRef = bumpArg.split('=')[1] || 'HEAD';
    const changed = [
        ...gitList(['diff', baseRef, '--name-only', '--', SITE]),
        ...gitList(['ls-files', '--others', '--exclude-standard', '--', SITE]),
    ];
    if (changed.length && versionAt(baseRef) === swVersion) {
        errors.push(`網站檔案已變更（${changed.slice(0, 5).join(', ')}${changed.length > 5 ? '…' : ''}）但 CACHE_VERSION 仍為 v${swVersion}，請執行 /release 調升版本`);
    }
    // sitemap 的 lastmod 也要一起更新（bump-version 會改成今天，算法要和 bump-version.mjs 相同）
    // 同一天第二次發佈時，上一版的 lastmod 已經是今天，日期相同是正確的，不算漏改
    const lastmodOf = (xml) => ((xml || '').match(/<lastmod>([^<]*)<\/lastmod>/) || [])[1];
    const baseLastmod = lastmodOf(git(['show', `${baseRef}:${SITE}/sitemap.xml`]));
    // 前一天調升、隔天才 push 時 lastmod 是前一天，比上一版新，可以接受；比上一版舊表示被改錯
    const today = new Date().toLocaleDateString('sv-SE'); // YYYY-MM-DD（本地時區）
    const curLastmod = lastmodOf(read(site('sitemap.xml')));
    if (changed.some(f => f !== site('sitemap.xml')) && baseLastmod && curLastmod) {
        if (curLastmod === baseLastmod && baseLastmod !== today) {
            errors.push(`網站檔案已變更但 sitemap.xml 的 lastmod 仍為 ${baseLastmod}，請執行 /release`);
        } else if (curLastmod < baseLastmod) {
            errors.push(`sitemap.xml 的 lastmod（${curLastmod}）比上一版（${baseLastmod}）還舊，請執行 /release 改成今天`);
        }
    }
}

// ==========================================
// 輸出結果
// ==========================================
for (const w of warnings) console.log(`⚠ ${w}`);
for (const e of errors) console.log(`✗ ${e}`);
console.log(errors.length
    ? `\n檢查失敗：${errors.length} 個錯誤、${warnings.length} 個警告`
    : `✓ 檢查通過（v${swVersion}，預快取 ${(precacheBytes / 1048576).toFixed(1)} MB${warnings.length ? `，${warnings.length} 個警告` : ''}）`);
process.exit(errors.length ? 1 : 0);
