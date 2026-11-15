"""
Noto Sans TC 字型裁切（subset）
只保留 ASCII、Big5 符號區、Big5 常用字，以及網站介面（docs/ 的 HTML、JS、manifest）實際用到的字元。

需求：pip install fonttools brotli
用法：python tools/subset-fonts.py <原始 TTF 所在資料夾>
完整流程（含如何取得原始 TTF）見 .claude/skills/subset-fonts/SKILL.md
"""

import sys
from pathlib import Path

from fontTools import subset

ROOT = Path(__file__).resolve().parent.parent
WEIGHTS = ["Regular", "Bold"]
SITE = ROOT / "docs"
UI_FILES = [SITE / "index.html", SITE / "404.html", SITE / "manifest.json", *(SITE / "js").glob("*.js")]


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


def collect_chars():
    chars = {chr(c) for c in range(0x20, 0x7F)}
    chars |= big5_chars(0xA1, 0xA3)  # 標點、符號、注音
    chars |= big5_chars(0xA4, 0xC6)  # 常用字
    for path in UI_FILES:
        chars |= set(path.read_text(encoding="utf-8"))
    return sorted(c for c in chars if c.isprintable() and c != "\ufeff")


def main():
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    src_dir = Path(sys.argv[1])
    chars = collect_chars()
    text = "".join(chars)

    for weight in WEIGHTS:
        src = src_dir / f"NotoSansTC-{weight}.ttf"
        out = SITE / "fonts" / f"NotoSansTC-{weight}.woff2"
        options = subset.Options()
        options.flavor = "woff2"
        options.layout_features = ["*"]
        options.name_IDs = ["*"]
        options.notdef_outline = True
        font = subset.load_font(str(src), options)
        subsetter = subset.Subsetter(options)
        subsetter.populate(text=text)
        subsetter.subset(font)
        subset.save_font(font, str(out), options)
        print(f"{out.relative_to(ROOT)}: {out.stat().st_size / 1024:.0f} KB")

    (ROOT / "tools" / "subset-chars.txt").write_text(text, encoding="utf-8")
    print(f"tools/subset-chars.txt: {len(chars)} 字")


if __name__ == "__main__":
    main()
