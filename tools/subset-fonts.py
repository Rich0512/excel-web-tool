"""
Noto Sans TC 字型裁切（subset）
每個字重拆成兩個檔，由 fonts.css 的 unicode-range 決定瀏覽器何時下載：
- 介面檔 NotoSansTC-<字重>-ui.woff2：ASCII、Big5 符號區，以及網站介面（docs/ 的 HTML、自家 JS、manifest）
  實際用到的字。首屏只需要這個檔。
- 常用字檔 NotoSansTC-<字重>-ext.woff2：Big5 常用字中不在介面檔的字。只有畫面出現這些字（例如學生姓名）時
  瀏覽器才會下載。
兩個檔的字元互斥。ui 的 unicode-range 精確等於 ui 檔的字；ext 的 unicode-range 用粗範圍（EXT_RANGE_BLOCKS 中
所有不在 ui 的碼位，再加上 ext 檔的字），讓 fonts.css 短很多（它會擋住首次繪製）。代價是名單只有不在 Big5 常用字
的罕用字時也會下載 ext 檔；ext 範圍不含任何 ui 字，所以介面用字不會觸發下載。
執行後會覆寫：
- docs/fonts/NotoSansTC-{Regular,Bold}-{ui,ext}.woff2
- tools/subset-chars-ui.txt、tools/subset-chars-ext.txt（各檔實際收錄的字，tools/check-site.mjs 比對用）
- docs/fonts/fonts.css 中 BEGIN/END 標記之間的 Noto Sans TC @font-face
可重複執行：介面用字改變後重跑即可。

需求：pip install fonttools brotli
用法：python tools/subset-fonts.py <原始 TTF 所在資料夾>
完整流程（含如何取得原始 TTF）見 .claude/skills/subset-fonts/SKILL.md
"""

import re
import sys
from pathlib import Path

from fontTools import subset
from fontTools.ttLib import TTFont

ROOT = Path(__file__).resolve().parent.parent
SITE = ROOT / "docs"
FONT_DIR = SITE / "fonts"
FONTS_CSS = FONT_DIR / "fonts.css"
WEIGHTS = {"Regular": 400, "Bold": 700}
PARTS = ["ui", "ext"]
# ext 粗範圍涵蓋的區塊：CJK 符號與標點、平假名、片假名（Big5 C6 列有假名）、CJK 統一表意文字
EXT_RANGE_BLOCKS = [(0x3000, 0x30FF), (0x4E00, 0x9FFF)]
CSS_BEGIN = "/* BEGIN Noto Sans TC：由 tools/subset-fonts.py 產生，請勿手動修改 */"
CSS_END = "/* END Noto Sans TC */"
# CSS 的 content: 字串（例如 content: "提示"）
CSS_CONTENT_RE = re.compile(r"""\bcontent:\s*(["'])(.*?)\1""")
# 拆檔前的舊檔名，重跑時一併移除
LEGACY = [FONT_DIR / "NotoSansTC-Regular.woff2", FONT_DIR / "NotoSansTC-Bold.woff2", ROOT / "tools" / "subset-chars.txt"]


def ui_files():
    # 與 tools/check-site.mjs 第 7 段掃描的檔案一致（js/libs/ 是第三方函式庫，不掃描）
    return [*sorted(SITE.glob("*.html")), SITE / "manifest.json", *sorted((SITE / "js").glob("*.js"))]


def big5_chars(lead_start, lead_end):
    """解碼 Big5 指定 lead byte 範圍內的所有字元"""
    chars = set()
    for lead in range(lead_start, lead_end + 1):
        for trail in [*range(0x40, 0x7F), *range(0xA1, 0xFF)]:
            try:
                chars.add(bytes([lead, trail]).decode("big5"))
            except UnicodeDecodeError:
                pass
    return chars


def printable(chars):
    return {c for c in chars if c.isprintable() and c != "﻿"}


def requested_chars():
    """回傳 (介面字, 介面檔案中用到的字, 常用字)；常用字已扣除介面字"""
    used = set()
    for path in ui_files():
        used |= set(path.read_text(encoding="utf-8"))
    # CSS 只收 content: 字串（畫面上會顯示的字），註解不收；規則與 check-site.mjs 第 7 段相同
    for path in sorted((SITE / "css").glob("*.css")):
        for m in CSS_CONTENT_RE.finditer(path.read_text(encoding="utf-8")):
            used |= set(m.group(2))
    used = printable(used)
    ui = printable({chr(c) for c in range(0x20, 0x7F)} | big5_chars(0xA1, 0xA3)) | used  # ASCII、標點、符號、注音
    ext = printable(big5_chars(0xA4, 0xC6)) - ui  # 常用字
    return ui, used, ext


def subset_font(src, out, chars):
    options = subset.Options()
    options.flavor = "woff2"
    options.layout_features = ["*"]
    options.name_IDs = ["*"]
    options.notdef_outline = True
    font = subset.load_font(str(src), options)
    subsetter = subset.Subsetter(options)
    subsetter.populate(text="".join(sorted(chars)))
    subsetter.subset(font)
    subset.save_font(font, str(out), options)
    # 以輸出檔的 cmap 為準（fontTools 可能把原始字型沒有的碼位對應到相近字形）
    return {chr(cp) for cp in TTFont(str(out)).getBestCmap()}


def unicode_range(chars):
    """把字元集合轉成 unicode-range，連續碼位合併成區段"""
    spans = []
    for cp in sorted(ord(c) for c in chars):
        if spans and cp == spans[-1][1] + 1:
            spans[-1][1] = cp
        else:
            spans.append([cp, cp])
    return ", ".join(f"U+{a:X}" if a == b else f"U+{a:X}-{b:X}" for a, b in spans)


def ext_range_chars(ui, ext):
    """ext 的粗範圍：區塊內所有不在 ui 的碼位，加上 ext 檔的字（確保 ⊇ ext 且與 ui 不重疊）"""
    block = {chr(cp) for start, end in EXT_RANGE_BLOCKS for cp in range(start, end + 1)}
    return (block | ext) - ui


def font_face(weight, part, chars):
    return "\n".join([
        "@font-face {",
        "    font-family: 'Noto Sans TC';",
        "    font-style: normal;",
        f"    font-weight: {WEIGHTS[weight]};",
        "    font-display: swap;",
        f"    src: url('NotoSansTC-{weight}-{part}.woff2') format('woff2');",
        f"    unicode-range: {unicode_range(chars)};",
        "}",
    ])


def write_css(coverage):
    css = FONTS_CSS.read_text(encoding="utf-8")
    pattern = re.compile(re.escape(CSS_BEGIN) + r".*?" + re.escape(CSS_END), re.S)
    if not pattern.search(css):
        sys.exit(f"{FONTS_CSS.relative_to(ROOT).as_posix()} 找不到 BEGIN/END 標記，請先把 Noto Sans TC 的 @font-face 換成：\n{CSS_BEGIN}\n{CSS_END}")
    ranges = {w: {"ui": coverage[w]["ui"], "ext": ext_range_chars(coverage[w]["ui"], coverage[w]["ext"])} for w in WEIGHTS}
    faces = [font_face(w, p, ranges[w][p]) for w in WEIGHTS for p in PARTS]
    block = "\n".join([CSS_BEGIN, "\n\n".join(faces), CSS_END])
    FONTS_CSS.write_text(pattern.sub(lambda _: block, css), encoding="utf-8", newline="\n")


def kb(path):
    return path.stat().st_size / 1024


def main():
    for stream in (sys.stdout, sys.stderr):
        stream.reconfigure(encoding="utf-8")  # 輸出導向管線時 Windows 預設 cp950，中文訊息會變亂碼
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    src_dir = Path(sys.argv[1])
    ui_req, used, ext_req = requested_chars()

    coverage = {}
    for weight in WEIGHTS:
        src = src_dir / f"NotoSansTC-{weight}.ttf"
        coverage[weight] = {
            part: subset_font(src, FONT_DIR / f"NotoSansTC-{weight}-{part}.woff2", chars)
            for part, chars in (("ui", ui_req), ("ext", ext_req))
        }
        overlap = coverage[weight]["ui"] & coverage[weight]["ext"]
        if overlap:
            sys.exit(f"{weight} 的 ui 與 ext 有 {len(overlap)} 個重複字，unicode-range 會互相衝突")

    # 兩個字重共用同一份清單，所以涵蓋範圍必須相同
    if coverage["Regular"] != coverage["Bold"]:
        sys.exit("Regular 與 Bold 裁切後的字元不同，請確認原始 TTF 是同一版本")
    ui, ext = coverage["Regular"]["ui"], coverage["Regular"]["ext"]

    for path in LEGACY:
        if path.exists():
            path.unlink()
            print(f"已移除舊檔 {path.relative_to(ROOT).as_posix()}")
    for part, chars in (("ui", ui), ("ext", ext)):
        (ROOT / "tools" / f"subset-chars-{part}.txt").write_text("".join(sorted(chars)), encoding="utf-8", newline="\n")
    write_css(coverage)

    for weight in WEIGHTS:
        for part in PARTS:
            out = FONT_DIR / f"NotoSansTC-{weight}-{part}.woff2"
            print(f"{out.relative_to(ROOT).as_posix()}: {kb(out):.0f} KB")
    print(f"tools/subset-chars-ui.txt: {len(ui)} 字；tools/subset-chars-ext.txt: {len(ext)} 字")
    print(f"首屏需要的中文字型（Regular ui）：{kb(FONT_DIR / 'NotoSansTC-Regular-ui.woff2'):.0f} KB")

    absent = sorted(used - ui)
    if absent:
        print(f"注意：原始字型沒有 {len(absent)} 個介面用字，會由系統字型顯示，請改用其他字："
              + "".join(absent) + "（" + " ".join(f"U+{ord(c):04X}" for c in absent) + "）")


if __name__ == "__main__":
    main()
