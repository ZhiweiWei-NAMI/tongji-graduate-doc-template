from __future__ import annotations

import re
import sys
import zipfile
from pathlib import Path


SONGTI = "\u5b8b\u4f53"
RFONTS = (
    '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" '
    f'w:eastAsia="{SONGTI}" w:cs="Times New Roman"/>'
)


def patch_reference_region(xml: str) -> tuple[str, int]:
    start = xml.find("<w:t>参考文献</w:t>")
    if start < 0:
        raise RuntimeError("未找到参考文献标题。")
    end = xml.find("<w:t>致谢</w:t>", start)
    if end < 0:
        end = xml.find("<w:t>个人简历", start)
    if end < 0:
        raise RuntimeError("未找到参考文献结束边界。")
    before, region, after = xml[:start], xml[start:end], xml[end:]
    count = 0

    def replace_rfonts(match: re.Match[str]) -> str:
        nonlocal count
        count += 1
        return RFONTS

    region = re.sub(r"<w:rFonts\b[^>]*/>", replace_rfonts, region)
    return before + region + after, count


def patch_docx(path: Path) -> int:
    with zipfile.ZipFile(path, "r") as zin:
        entries = [(item, zin.read(item.filename)) for item in zin.infolist()]
    patched_entries = []
    count = 0
    for item, data in entries:
        if item.filename == "word/document.xml":
            xml = data.decode("utf-8")
            xml, count = patch_reference_region(xml)
            patched_entries.append((item, xml.encode("utf-8")))
        else:
            patched_entries.append((item, data))

    tmp_path = path.with_name(path.stem + ".__fontfix_tmp__.docx")
    if tmp_path.exists():
        tmp_path.unlink()
    try:
        with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item, data in patched_entries:
                zout.writestr(item, data)
        tmp_path.replace(path)
    finally:
        if tmp_path.exists():
            tmp_path.unlink()
    return count


def main() -> int:
    if len(sys.argv) != 2:
        print("Usage: fix_reference_fonts.py <docx>", file=sys.stderr)
        return 2
    count = patch_docx(Path(sys.argv[1]))
    print(f"patched rFonts tags in reference region: {count}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
