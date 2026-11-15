#!/usr/bin/env node
// ============================================
// 本機預覽伺服器：模擬 GitHub Pages 的行為
// - 網站掛在子路徑 /excel-web-tool/（與正式網址相同）
// - 不存在的路徑回傳 docs/404.html 與 404 狀態碼
// - 只監聽 127.0.0.1；localhost 屬於安全來源，Service Worker 可正常註冊
// 用法：node tools/serve.mjs [port]（預設 8000）
// ============================================

import { createServer } from 'node:http';
import { readFile, stat } from 'node:fs/promises';
import { join, extname, resolve } from 'node:path';
import { SITE_DIR, BASE_PATH, isInside } from './site-config.mjs';

const PORT = Number(process.argv[2]) || 8000;

const MIME = {
    '.html': 'text/html; charset=utf-8',
    '.js': 'text/javascript; charset=utf-8',
    '.css': 'text/css; charset=utf-8',
    '.json': 'application/json; charset=utf-8',
    '.xml': 'application/xml; charset=utf-8',
    '.txt': 'text/plain; charset=utf-8',
    '.png': 'image/png',
    '.jpg': 'image/jpeg',
    '.ico': 'image/x-icon',
    '.woff2': 'font/woff2',
    '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
};

async function resolveFile(urlPath) {
    if (!urlPath.startsWith(BASE_PATH)) return null;
    let rel;
    try {
        rel = decodeURIComponent(urlPath.slice(BASE_PATH.length)) || 'index.html';
    } catch {
        return null; // 格式錯誤的網址一律視為不存在
    }
    const full = resolve(SITE_DIR, rel);
    // 防止 ../ 跳出網站資料夾
    if (!isInside(SITE_DIR, full)) return null;
    try {
        const info = await stat(full);
        return info.isDirectory() ? join(full, 'index.html') : full;
    } catch {
        return null;
    }
}

const server = createServer(async (req, res) => {
    let urlPath;
    try {
        // Node 以 latin1 解讀請求行；還原成 UTF-8，未編碼的中文檔名（例如 curl 送出的）才能正確對應
        urlPath = new URL(Buffer.from(req.url, 'latin1').toString('utf8'), 'http://localhost').pathname;
    } catch {
        res.writeHead(400).end();
        return;
    }
    if (urlPath === '/' || urlPath === BASE_PATH.slice(0, -1)) {
        res.writeHead(302, { Location: BASE_PATH }).end();
        return;
    }

    let file = await resolveFile(urlPath);
    let status = 200;
    if (!file) {
        file = join(SITE_DIR, '404.html');
        status = 404;
    }
    try {
        const body = await readFile(file);
        res.writeHead(status, { 'Content-Type': MIME[extname(file)] || 'application/octet-stream' }).end(body);
    } catch {
        res.writeHead(404).end();
    }
    console.log(`${status} ${urlPath}`);
});

server.listen(PORT, '127.0.0.1', () => {
    console.log(`預覽網址：http://localhost:${PORT}${BASE_PATH}（Ctrl+C 結束）`);
});
