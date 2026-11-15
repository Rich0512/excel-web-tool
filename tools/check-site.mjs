#!/usr/bin/env node
// ============================================
// GitHub Pages 部署前檢查（網站根目錄為 docs/）
// 用法：node tools/check-site.mjs [--require-bump[=<git ref>]]
//   --require-bump：docs/ 相對 <git ref>（預設 HEAD）有變更時，要求 CACHE_VERSION 也必須已調升
// 有錯誤時 exit 1；只有警告時 exit 0
// ============================================

import { readFileSync, existsSync, statSync, readdirSync } from 'node:fs';
import { join } from 'node:path';
import { ROOT, SITE, BASE_PATH, TEMPLATE_DIR, SPREADSHEET_RE, gitList, versionOf, versionAt, isForbiddenSpreadsheet } from './site-config.mjs';
import { inspectTemplate } from './template-check.mjs';

const PRECACHE_WARN_BYTES = 5 * 1024 * 1024;
const CJK_RE = /[　-〿㐀-鿿＀-￯]/;

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
// 1. 版本號一致：service-worker.js 與 index.html 的 ?v=
// ==========================================
const sw = read(site('service-worker.js'));
const swVersion = versionOf(sw);
const indexHtml = read(site('index.html'));
if (!swVersion) {
    errors.push('service-worker.js 找不到 CACHE_VERSION');
} else {
    const queryVersions = [...indexHtml.matchAll(/\?v=([\d.]+)/g)].map(m => m[1]);
    const mismatched = [...new Set(queryVersions.filter(v => v !== swVersion))];
    if (mismatched.length) {
        errors.push(`index.html 的 ?v= (${mismatched.join(', ')}) 與 CACHE_VERSION (v${swVersion}) 不一致`);
    }
    // 自己的 JS/CSS 都要帶版本參數（js/libs/ 的第三方檔案不會改動，不需要）
    for (const m of indexHtml.matchAll(/(?:src|href)="((?:js|css|fonts)\/[^"]+\.(?:js|css))"/g)) {
        if (!m[1].startsWith('js/libs/')) errors.push(`index.html 的 ${m[1]} 缺少 ?v= 版本參數`);
    }
}

// ==========================================
// 2. 預快取清單：檔案都要存在（cache.addAll 只要缺一個就整個安裝失敗）
// ==========================================
const assetBlock = (sw.match(/const STATIC_ASSETS = \[([\s\S]*?)\];/) || [])[1] || '';
const assets = [...assetBlock.matchAll(/'([^']+)'/g)].map(m => m[1]);
let precacheBytes = 0;
for (const asset of assets) {
    if (asset === './') continue;
    const rel = site(asset.replace(/^\.\//, ''));
    if (!existsSync(join(ROOT, rel))) {
        errors.push(`service-worker.js 預快取的檔案不存在：${asset}`);
        continue;
    }
    precacheBytes += statSync(join(ROOT, rel)).size;
    if (/\.ttf$/i.test(rel)) errors.push(`預快取不應包含 TTF（瀏覽器只會用 woff2）：${asset}`);
}
for (const file of [...jsFiles, ...cssFiles]) {
    const asset = `./${file.slice(SITE.length + 1)}`;
    if (!assets.includes(asset)) errors.push(`${file} 不在 service-worker.js 的 STATIC_ASSETS，離線時會載不到`);
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
    });
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
// 7. 介面用到的中文字都在裁切後的字型內
// ==========================================
const charsFile = 'tools/subset-chars.txt';
if (existsSync(join(ROOT, charsFile))) {
    const covered = new Set(read(charsFile));
    const missing = new Set();
    for (const file of [...htmlFiles, site('manifest.json'), ...ownJsFiles]) {
        for (const ch of read(file)) {
            if (CJK_RE.test(ch) && !covered.has(ch)) missing.add(ch);
        }
    }
    if (missing.size) {
        errors.push(`字型未涵蓋 ${missing.size} 個介面用字：${[...missing].join('')}（請執行 /subset-fonts）`);
    }
} else {
    warnings.push(`${charsFile} 不存在，略過字型涵蓋檢查`);
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
