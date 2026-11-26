#!/usr/bin/env node
// ============================================
// 同步調升網站快取版本
// 用法：node tools/bump-version.mjs [X.Y]
//   - 同步更新 docs/service-worker.js 的 CACHE_VERSION 與 STATIC_ASSETS 的 ?v=、
//     docs/ 根目錄每個 .html（index.html、guide.html、404.html…）的所有 ?v=、
//     docs/sitemap.xml 的每個 <lastmod>（今天）
//   - 省略版本號時：若相對 origin/main 尚未調升，次版號 +1；已調升則只同步各處並更新日期
// ============================================

import { readFileSync, writeFileSync, readdirSync } from 'node:fs';
import { join } from 'node:path';
import { SITE_DIR, VERSION_RE, versionOf, versionAt } from './site-config.mjs';

const FILES = {
    sw: join(SITE_DIR, 'service-worker.js'),
    sitemap: join(SITE_DIR, 'sitemap.xml'),
};
// 根目錄的頁面動態列出，新增頁面時不必改這裡（check-site.mjs 也是用同樣的方式列出頁面）
const PAGES = readdirSync(SITE_DIR).filter(f => f.endsWith('.html')).sort();

const sw = readFileSync(FILES.sw, 'utf8');
const current = versionOf(sw);
if (!current) {
    console.error('✗ docs/service-worker.js 找不到 CACHE_VERSION');
    process.exit(1);
}

let next = process.argv[2]?.replace(/^v/, '');
if (next && !/^\d+\.\d+$/.test(next)) {
    console.error(`✗ 版本號格式應為 X.Y，收到：${next}`);
    process.exit(1);
}
if (!next) {
    const deployed = versionAt('origin/main');
    if (deployed && deployed !== current) {
        next = current;
        console.log(`已相對線上版本 v${deployed} 調升為 v${current}，只同步各處版本號與日期`);
    } else {
        const [major, minor] = current.split('.').map(Number);
        next = `${major}.${minor + 1}`;
    }
}

const today = new Date().toLocaleDateString('sv-SE'); // YYYY-MM-DD（本地時區）
// STATIC_ASSETS 以與頁面相同的 ?v= 預快取，兩邊要一起改
writeFileSync(FILES.sw, sw.replace(VERSION_RE, `const CACHE_VERSION = 'v${next}';`).replace(/\?v=[\d.]+/g, `?v=${next}`));
for (const page of PAGES) {
    const file = join(SITE_DIR, page);
    writeFileSync(file, readFileSync(file, 'utf8').replace(/\?v=[\d.]+/g, `?v=${next}`));
}
// 每一頁的 lastmod 都換成今天（只換第一個的話，其他頁面的日期會停在舊版）
writeFileSync(FILES.sitemap, readFileSync(FILES.sitemap, 'utf8').replace(/<lastmod>[^<]*<\/lastmod>/g, `<lastmod>${today}</lastmod>`));

console.log(`✓ v${current} → v${next}（${PAGES.join('、')}；sitemap lastmod：${today}）`);
