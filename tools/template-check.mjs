// ============================================
// 範本假資料檢查：確認 docs/templates/ 的 xlsx 只含虛構姓名
// 只用 Node 內建模組（zlib 解壓 zip，正規式讀工作表 XML），由 check-site.mjs 呼叫
// 回報只含儲存格位置與筆數，絕不輸出儲存格內容（萬一是真實名冊也不會外洩）
// ============================================

import { readFileSync } from 'node:fs';
import { inflateRawSync } from 'node:zlib';

// 範本姓名格式：測試生<班級>_<座號>，例如 測試生201_23
export const FAKE_NAME_RE = /^測試生\d+_\d+$/;
const NAME_HEADER_RE = /姓名/;

// 讀取 zip 內所有檔案（只支援 stored 與 deflate，xlsx 只會用到這兩種）
function unzip(buf) {
    const eocd = buf.lastIndexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06]));
    if (eocd === -1) throw new Error('不是有效的 zip／xlsx 檔');
    const count = buf.readUInt16LE(eocd + 10);
    let p = buf.readUInt32LE(eocd + 16);
    const files = new Map();
    for (let i = 0; i < count; i++) {
        if (buf.readUInt32LE(p) !== 0x02014b50) throw new Error('zip 中央目錄格式錯誤');
        const method = buf.readUInt16LE(p + 10);
        const size = buf.readUInt32LE(p + 20);
        const nameLen = buf.readUInt16LE(p + 28);
        const extraLen = buf.readUInt16LE(p + 30);
        const commentLen = buf.readUInt16LE(p + 32);
        const local = buf.readUInt32LE(p + 42);
        const name = buf.toString('utf8', p + 46, p + 46 + nameLen);
        const dataStart = local + 30 + buf.readUInt16LE(local + 26) + buf.readUInt16LE(local + 28);
        const data = buf.subarray(dataStart, dataStart + size);
        if (method === 0) files.set(name, data);
        else if (method === 8) files.set(name, inflateRawSync(data));
        else throw new Error(`不支援的 zip 壓縮方式 ${method}`);
        p += 46 + nameLen + extraLen + commentLen;
    }
    return files;
}

const decode = (s) => s
    .replace(/&#x([0-9a-f]+);/gi, (_, h) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(Number(d)))
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');

const texts = (xml) => [...xml.matchAll(/<t\b[^>]*>([^<]*)<\/t>/g)].map(m => decode(m[1])).join('');

// 解析一張工作表的儲存格：{ ref, col, row, value }
function parseCells(xml, shared) {
    const cells = [];
    for (const m of xml.matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
        const ref = (m[1].match(/\br="([A-Z]+)(\d+)"/) || []);
        if (!ref[1]) continue;
        const type = (m[1].match(/\bt="([^"]+)"/) || [])[1];
        const body = m[2] || '';
        let value = '';
        if (type === 'inlineStr') value = texts(body);
        else if (type === 's') value = shared[Number((body.match(/<v>([^<]*)<\/v>/) || [])[1])] ?? '';
        else value = decode((body.match(/<v>([^<]*)<\/v>/) || [])[1] || '');
        cells.push({ ref: ref[1] + ref[2], col: ref[1], row: Number(ref[2]), value: value.trim() });
    }
    return cells;
}

// 回傳 { checked, bad: [儲存格位置...] }；格式錯誤時丟出例外
export function inspectTemplate(path) {
    const files = unzip(readFileSync(path));
    const sharedXml = files.get('xl/sharedStrings.xml')?.toString('utf8') || '';
    const shared = [...sharedXml.matchAll(/<si>([\s\S]*?)<\/si>/g)].map(m => texts(m[1]));
    let checked = 0;
    const bad = [];
    for (const [name, data] of files) {
        const sheet = name.match(/^xl\/worksheets\/(sheet\d+)\.xml$/);
        if (!sheet) continue;
        const cells = parseCells(data.toString('utf8'), shared);
        for (const header of cells.filter(c => NAME_HEADER_RE.test(c.value))) {
            for (const c of cells) {
                // 同一欄下方若重複出現表頭（分段列表），表頭本身不算姓名
                if (c.col !== header.col || c.row <= header.row || !c.value || NAME_HEADER_RE.test(c.value)) continue;
                checked++;
                if (!FAKE_NAME_RE.test(c.value)) bad.push(`${sheet[1]}!${c.ref}`);
            }
        }
    }
    return { checked, bad };
}
