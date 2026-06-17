from __future__ import annotations

import argparse
import re
import zipfile
from pathlib import Path

from lxml import etree


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}

EMU_PER_INCH = 914400


def qn(tag: str) -> str:
    prefix, local = tag.split(":")
    return f"{{{NS[prefix]}}}{local}"


def load_package(path: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(path, "r") as zf:
        return {name: zf.read(name) for name in zf.namelist()}


def save_package(package: dict[str, bytes], path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name, data in package.items():
            zf.writestr(name, data)


def read_xml(package: dict[str, bytes], name: str) -> etree._Element:
    return etree.fromstring(package[name])


def write_xml(package: dict[str, bytes], name: str, root: etree._Element) -> None:
    package[name] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")


def visible_text(node: etree._Element) -> str:
    parts: list[str] = []
    for t in node.xpath(".//w:t | .//m:t", namespaces=NS):
        if t.text:
            parts.append(t.text)
    return "".join(parts).strip()


def p_style(p: etree._Element) -> str:
    vals = p.xpath("./w:pPr/w:pStyle/@w:val", namespaces=NS)
    return vals[0] if vals else ""


def has_drawing(p: etree._Element) -> bool:
    return bool(p.xpath(".//w:drawing | .//w:pict", namespaces=NS))


def has_equation_field(p: etree._Element) -> bool:
    return bool(p.xpath(".//w:instrText[contains(., 'SEQ Equation')]", namespaces=NS))


def has_field_or_break(p: etree._Element) -> bool:
    return bool(p.xpath(".//w:instrText | .//w:fldChar | .//w:br", namespaces=NS))


def ensure_child(parent: etree._Element, tag: str) -> etree._Element:
    child = parent.find(qn(tag))
    if child is None:
        child = etree.SubElement(parent, qn(tag))
    return child


def ensure_ppr(p: etree._Element) -> etree._Element:
    ppr = p.find(qn("w:pPr"))
    if ppr is None:
        ppr = etree.Element(qn("w:pPr"))
        p.insert(0, ppr)
    return ppr


def remove_ppr_children(ppr: etree._Element, *tags: str) -> None:
    for tag in tags:
        for child in list(ppr.xpath(f"./{tag}", namespaces=NS)):
            ppr.remove(child)


def remove_keep_next(p: etree._Element) -> bool:
    ppr = p.find(qn("w:pPr"))
    if ppr is None:
        return False
    removed = False
    for child in list(ppr.xpath("./w:keepNext", namespaces=NS)):
        ppr.remove(child)
        removed = True
    return removed


def set_rpr_size(rpr: etree._Element, size: str) -> None:
    sz = rpr.find(qn("w:sz"))
    if sz is None:
        sz = etree.SubElement(rpr, qn("w:sz"))
    sz.set(qn("w:val"), size)
    szcs = rpr.find(qn("w:szCs"))
    if szcs is None:
        szcs = etree.SubElement(rpr, qn("w:szCs"))
    szcs.set(qn("w:val"), size)


def set_run_size(run: etree._Element, size: str) -> None:
    rpr = run.find(qn("w:rPr"))
    if rpr is None:
        rpr = etree.Element(qn("w:rPr"))
        run.insert(0, rpr)
    set_rpr_size(rpr, size)


def set_paragraph_default_size(ppr: etree._Element, size: str) -> None:
    rpr = ensure_child(ppr, "w:rPr")
    set_rpr_size(rpr, size)


def set_all_word_and_math_run_sizes(p: etree._Element, size: str) -> None:
    for run in p.xpath(".//w:r | .//m:r", namespaces=NS):
        set_run_size(run, size)
    for rpr in p.xpath(".//m:oMath//w:rPr | .//m:oMathPara//w:rPr", namespaces=NS):
        set_rpr_size(rpr, size)


def set_math_run_sizes(p: etree._Element, size: str) -> None:
    for run in p.xpath(".//m:r", namespaces=NS):
        set_run_size(run, size)
    for rpr in p.xpath(".//m:oMath//w:rPr | .//m:oMathPara//w:rPr", namespaces=NS):
        set_rpr_size(rpr, size)


def set_formula_tabs(ppr: etree._Element) -> None:
    remove_ppr_children(ppr, "w:tabs")
    tabs = etree.SubElement(ppr, qn("w:tabs"))
    tab1 = etree.SubElement(tabs, qn("w:tab"))
    tab1.set(qn("w:val"), "center")
    tab1.set(qn("w:pos"), "4153")
    tab2 = etree.SubElement(tabs, qn("w:tab"))
    tab2.set(qn("w:val"), "right")
    tab2.set(qn("w:pos"), "8306")


def set_keep_next(p: etree._Element) -> None:
    ppr = ensure_ppr(p)
    if ppr.find(qn("w:keepNext")) is None:
        etree.SubElement(ppr, qn("w:keepNext"))


def normalize_formula_paragraph(p: etree._Element) -> None:
    ppr = ensure_ppr(p)
    remove_ppr_children(ppr, "w:jc", "w:keepNext", "w:pageBreakBefore")
    set_formula_tabs(ppr)
    spacing = ensure_child(ppr, "w:spacing")
    spacing.set(qn("w:before"), "0")
    spacing.set(qn("w:after"), "0")
    spacing.set(qn("w:line"), "240")
    spacing.set(qn("w:lineRule"), "auto")
    ind = ensure_child(ppr, "w:ind")
    ind.attrib.pop(qn("w:hanging"), None)
    ind.attrib.pop(qn("w:hangingChars"), None)
    ind.attrib.pop(qn("w:firstLineChars"), None)
    ind.set(qn("w:firstLine"), "480")
    set_paragraph_default_size(ppr, "24")
    set_all_word_and_math_run_sizes(p, "24")


def normalize_body_paragraph(p: etree._Element) -> None:
    ppr = ensure_ppr(p)
    remove_ppr_children(ppr, "w:jc", "w:spacing", "w:keepNext", "w:pageBreakBefore")
    ind = ensure_child(ppr, "w:ind")
    ind.attrib.pop(qn("w:hanging"), None)
    ind.attrib.pop(qn("w:hangingChars"), None)
    ind.attrib.pop(qn("w:firstLineChars"), None)
    ind.set(qn("w:firstLine"), "480")
    set_paragraph_default_size(ppr, "24")
    set_math_run_sizes(p, "24")


def body_blocks(root: etree._Element) -> list[etree._Element]:
    body = root.find(qn("w:body"))
    if body is None:
        raise RuntimeError("Document body not found")
    return [child for child in body if child.tag in {qn("w:p"), qn("w:tbl")}]


def remove_empty_ch4_paragraphs(root: etree._Element) -> int:
    body = root.find(qn("w:body"))
    if body is None:
        raise RuntimeError("Document body not found")
    in_ch4 = False
    removed = 0
    for block in list(body):
        if block.tag != qn("w:p"):
            continue
        text = visible_text(block)
        if re.match(r"^4\.1\s+引言$", text):
            in_ch4 = True
            continue
        if in_ch4 and re.match(r"^5\.1\s+引言$", text):
            break
        if not in_ch4:
            continue
        if text or has_drawing(block) or has_field_or_break(block) or p_style(block):
            continue
        body.remove(block)
        removed += 1
    return removed


def set_drawing_height_before_caption(root: etree._Element, caption_fragment: str, max_height_inches: float) -> bool:
    blocks = body_blocks(root)
    max_cy = int(max_height_inches * EMU_PER_INCH)
    for idx, block in enumerate(blocks):
        if block.tag != qn("w:p") or caption_fragment not in visible_text(block):
            continue
        image = None
        for prev in reversed(blocks[:idx]):
            if prev.tag == qn("w:p") and has_drawing(prev):
                image = prev
                break
            if prev.tag == qn("w:p") and visible_text(prev):
                break
        if image is None:
            return False
        ext_nodes = image.xpath(".//wp:extent | .//a:xfrm/a:ext", namespaces=NS)
        wp_ext = image.xpath(".//wp:extent", namespaces=NS)
        if not wp_ext:
            return False
        old_cx = int(wp_ext[0].get("cx"))
        old_cy = int(wp_ext[0].get("cy"))
        if old_cy <= max_cy:
            return True
        new_cy = max_cy
        new_cx = int(old_cx * new_cy / old_cy)
        for ext in ext_nodes:
            ext.set("cx", str(new_cx))
            ext.set("cy", str(new_cy))
        return True
    return False


def compact_ch4_late_figures(root: etree._Element) -> dict[str, bool]:
    targets = {
        "图 4-20 观测退化下事件宏平均F1": 2.35,
        "图 4-21 观测退化下关键事件召回率": 2.35,
        "图 4-22 端侧语义监测器的时延—关键事件召回权衡": 2.35,
        "图 4-24 本地修复相对全局重规划的成本缺口": 2.15,
        "图 4-25 修复范围变化下的重规划时延": 2.15,
        "图 4-26 Jetson Orin NX端侧语义监测部署基准": 2.60,
        "图 4-27 Jetson Orin NX端侧语义监测功耗轨迹": 1.85,
    }
    return {caption: set_drawing_height_before_caption(root, caption, height) for caption, height in targets.items()}


def remove_all_caption_keep_next(root: etree._Element) -> int:
    removed = 0
    for p in root.xpath(".//w:body/w:p", namespaces=NS):
        text = visible_text(p)
        if p_style(p) == "afd" or re.match(r"^(图|表|算法)\s*\d+-\d+", text):
            if remove_keep_next(p):
                removed += 1
    return removed


def remove_caption_style_keep_next(package: dict[str, bytes]) -> int:
    if "word/styles.xml" not in package:
        return 0
    styles = read_xml(package, "word/styles.xml")
    removed = 0
    for style in styles.xpath(".//w:style", namespaces=NS):
        style_id = style.get(qn("w:styleId"), "")
        names = style.xpath("./w:name/@w:val", namespaces=NS)
        style_name = names[0] if names else ""
        if style_id != "afd" and style_name.lower() != "caption":
            continue
        ppr = style.find(qn("w:pPr"))
        if ppr is None:
            continue
        for keep in list(ppr.xpath("./w:keepNext", namespaces=NS)):
            ppr.remove(keep)
            removed += 1
    if removed:
        write_xml(package, "word/styles.xml", styles)
    return removed


def repair_ch4(root: etree._Element) -> dict[str, int]:
    in_ch4 = False
    formula_count = 0
    body_count = 0
    for block in body_blocks(root):
        if block.tag != qn("w:p"):
            continue
        text = visible_text(block)
        if re.match(r"^4\.1\s+引言$", text):
            in_ch4 = True
            continue
        if in_ch4 and re.match(r"^5\.1\s+引言$", text):
            break
        if not in_ch4 or has_drawing(block):
            continue
        is_formula = has_equation_field(block) or (bool(re.search(r"（4-\d+）", text)) and bool(block.xpath(".//m:oMath", namespaces=NS)))
        if is_formula:
            normalize_formula_paragraph(block)
            formula_count += 1
        elif p_style(block) == "" and text:
            normalize_body_paragraph(block)
            body_count += 1
    removed_empty_paragraphs = remove_empty_ch4_paragraphs(root)
    figure_compaction = compact_ch4_late_figures(root)
    removed_caption_keep_next = remove_all_caption_keep_next(root)
    missing_compaction = sum(1 for ok in figure_compaction.values() if not ok)
    return {
        "formulas": formula_count,
        "body_paragraphs": body_count,
        "removed_empty_paragraphs": removed_empty_paragraphs,
        "removed_caption_keep_next": removed_caption_keep_next,
        "figure_compaction_missing": missing_compaction,
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()
    package = load_package(Path(args.input))
    root = read_xml(package, "word/document.xml")
    counts = repair_ch4(root)
    counts["removed_caption_style_keep_next"] = remove_caption_style_keep_next(package)
    write_xml(package, "word/document.xml", root)
    save_package(package, Path(args.output))
    print(counts)


if __name__ == "__main__":
    main()
