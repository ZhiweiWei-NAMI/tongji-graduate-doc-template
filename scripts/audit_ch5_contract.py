from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from pathlib import Path

from lxml import etree


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
}


def qn(tag: str) -> str:
    prefix, local = tag.split(":")
    return f"{{{NS[prefix]}}}{local}"


def node_text(node: etree._Element) -> str:
    return "".join(node.xpath(".//w:t/text() | .//m:t/text()", namespaces=NS)).strip()


def style_id(p: etree._Element) -> str:
    vals = p.xpath("./w:pPr/w:pStyle/@w:val", namespaces=NS)
    return vals[0] if vals else ""


def instr_text(node: etree._Element) -> str:
    return " ".join(
        node.xpath(".//w:instrText/text() | .//w:fldSimple/@w:instr", namespaces=NS)
    )


def has_omath(node: etree._Element) -> bool:
    return bool(node.xpath(".//m:oMath | .//m:oMathPara", namespaces=NS))


def has_drawing(node: etree._Element) -> bool:
    return bool(node.xpath(".//w:drawing | .//w:pict", namespaces=NS))


def run_sizes(node: etree._Element) -> set[str]:
    return set(node.xpath(".//w:rPr/w:sz/@w:val | .//m:rPr/w:sz/@w:val", namespaces=NS))


def body_blocks(root: etree._Element) -> list[etree._Element]:
    body = root.find(qn("w:body"))
    if body is None:
        raise RuntimeError("word/document.xml has no body")
    return [n for n in body if n.tag in {qn("w:p"), qn("w:tbl")}]


def is_heading5_start(text: str, style: str) -> bool:
    if style not in {"af", "af1", "1", "2", "3"}:
        return False
    return bool(re.match(r"^5\.1\s+", text))


def is_heading6_start(text: str, style: str) -> bool:
    if style not in {"af", "af1", "1", "2", "3"}:
        return False
    return bool(re.match(r"^6\.1\s+", text))


def chapter5_slice(blocks: list[etree._Element]) -> list[tuple[int, etree._Element]]:
    start = None
    end = None
    for i, node in enumerate(blocks):
        if node.tag != qn("w:p"):
            continue
        text = node_text(node)
        st = style_id(node)
        if start is None and is_heading5_start(text, st):
            start = i
            continue
        if start is not None and is_heading6_start(text, st):
            end = i
            break
    if start is None:
        raise RuntimeError("real Chapter 5 body start was not found")
    if end is None:
        end = len(blocks)
    return list(enumerate(blocks[start:end], start=start))


def border_val(el: etree._Element | None) -> str:
    if el is None:
        return ""
    return el.get(qn("w:val"), "")


def border_size(el: etree._Element | None) -> str:
    if el is None:
        return ""
    return el.get(qn("w:sz"), "")


def table_report(idx: int, tbl: etree._Element) -> dict:
    tbl_pr = tbl.find(qn("w:tblPr"))
    borders = tbl_pr.find(qn("w:tblBorders")) if tbl_pr is not None else None
    table_borders = {}
    for name in ["top", "bottom", "left", "right", "insideH", "insideV"]:
        el = borders.find(qn(f"w:{name}")) if borders is not None else None
        table_borders[name] = {"val": border_val(el), "sz": border_size(el)}
    rows = tbl.xpath("./w:tr", namespaces=NS)
    double_spaces = []
    vertical_cells = []
    for r_idx, tr in enumerate(rows):
        for c_idx, tc in enumerate(tr.xpath("./w:tc", namespaces=NS)):
            text = node_text(tc)
            if "  " in text:
                double_spaces.append({"row": r_idx, "col": c_idx, "text": text[:120]})
            tc_borders = tc.xpath("./w:tcPr/w:tcBorders", namespaces=NS)
            if tc_borders:
                left = border_val(tc_borders[0].find(qn("w:left")))
                right = border_val(tc_borders[0].find(qn("w:right")))
                if left not in {"", "nil", "none"} or right not in {"", "nil", "none"}:
                    vertical_cells.append({"row": r_idx, "col": c_idx, "left": left, "right": right})
    return {
        "index": idx,
        "rows": len(rows),
        "cols": max((len(r.xpath("./w:tc", namespaces=NS)) for r in rows), default=0),
        "text": node_text(tbl)[:200],
        "font_sizes": sorted(run_sizes(tbl)),
        "borders": table_borders,
        "double_space_cells": double_spaces,
        "vertical_border_cells": vertical_cells[:40],
    }


def formula_report(idx: int, p: etree._Element) -> dict:
    ppr = p.find(qn("w:pPr"))
    spacing = ppr.find(qn("w:spacing")) if ppr is not None else None
    tabs = []
    if ppr is not None:
        for tab in ppr.xpath("./w:tabs/w:tab", namespaces=NS):
            tabs.append({"val": tab.get(qn("w:val"), ""), "pos": tab.get(qn("w:pos"), "")})
    return {
        "index": idx,
        "text": node_text(p),
        "instr": instr_text(p),
        "has_omath": has_omath(p),
        "has_seq_equation": "SEQ Equation" in instr_text(p),
        "line": spacing.get(qn("w:line"), "") if spacing is not None else "",
        "lineRule": spacing.get(qn("w:lineRule"), "") if spacing is not None else "",
        "sizes": sorted(run_sizes(p)),
        "tabs": tabs,
    }


def is_standalone_formula(p: etree._Element, text: str) -> bool:
    if not has_omath(p):
        return False
    if not re.search(r"（5-\d+）\s*$", text):
        return False
    if text.startswith("式（"):
        return False
    return len(text) <= 260


def has_header_separator(tbl: etree._Element) -> bool:
    rows = tbl.xpath("./w:tr", namespaces=NS)
    if not rows:
        return False
    for bottom in rows[0].xpath(".//w:tcPr/w:tcBorders/w:bottom", namespaces=NS):
        if bottom.get(qn("w:val"), "") not in {"", "nil", "none"}:
            return True
    return False


def has_ref_field(node: etree._Element) -> bool:
    instr = instr_text(node)
    return " REF " in f" {instr} " or "PAGEREF" in instr


def audit_docx(path: Path) -> dict:
    with zipfile.ZipFile(path) as zf:
        document = etree.fromstring(zf.read("word/document.xml"))
        styles = etree.fromstring(zf.read("word/styles.xml"))
    blocks = body_blocks(document)
    ch5 = chapter5_slice(blocks)

    failures = []
    formulas = []
    tables = []
    captions = []
    plain_refs = []
    raw_inline = []
    centered_body = []
    page_breaks = []
    empty_paras = []
    forbidden = []

    caption_keep_styles = []
    for style in styles.xpath(".//w:style", namespaces=NS):
        sid = style.get(qn("w:styleId"), "")
        names = style.xpath("./w:name/@w:val", namespaces=NS)
        name = names[0] if names else ""
        if (sid in {"afd", "aff"} or "caption" in name.lower()) and style.xpath("./w:pPr/w:keepNext", namespaces=NS):
            caption_keep_styles.append({"styleId": sid, "name": name})

    formula_number_rx = re.compile(r"\(?5-\d+\)?")
    figure_table_ref_rx = re.compile(r"[\u56fe\u8868]\s*5-\d+")
    raw_inline_rx = re.compile(
        r"("
        r"[A-Za-z\u03b1-\u03c9\u0391-\u03a9][A-Za-z0-9]*_\{?[A-Za-z0-9,\-]+\}?"
        r"|[A-Za-z\u03b1-\u03c9\u0391-\u03a9][A-Za-z0-9]*\^\{?[A-Za-z0-9,\-]+\}?"
        r"|[A-Za-z\u03b1-\u03c9\u0391-\u03a9][A-Za-z0-9]*\^\{?[A-Za-z0-9,\-]+\}?_\{?[A-Za-z0-9,\-]+\}?"
        r")"
    )

    for idx, node in ch5:
        if node.tag == qn("w:tbl"):
            tr = table_report(idx, node)
            tables.append(tr)
            left = tr["borders"]["left"]["val"]
            right = tr["borders"]["right"]["val"]
            inside_v = tr["borders"]["insideV"]["val"]
            if left not in {"", "nil", "none"} or right not in {"", "nil", "none"} or inside_v not in {"", "nil", "none"}:
                failures.append({"type": "table_has_vertical_or_outer_borders", "index": idx, "text": tr["text"], "borders": tr["borders"]})
            if tr["vertical_border_cells"]:
                failures.append({"type": "table_cell_vertical_borders", "index": idx, "items": tr["vertical_border_cells"][:10]})
            if tr["double_space_cells"]:
                failures.append({"type": "table_double_spaces", "index": idx, "items": tr["double_space_cells"][:10]})
            if set(tr["font_sizes"]) - {"18"}:
                failures.append({"type": "table_font_size_not_xiaowu", "index": idx, "sizes": tr["font_sizes"], "text": tr["text"]})
            top = tr["borders"]["top"]["val"]
            bottom = tr["borders"]["bottom"]["val"]
            if top in {"", "nil", "none"} or bottom in {"", "nil", "none"} or not has_header_separator(node):
                failures.append({"type": "table_missing_three_line_rules", "index": idx, "borders": tr["borders"], "text": tr["text"]})
            continue

        text = node_text(node)
        st = style_id(node)
        instr = instr_text(node)

        if not text and not has_drawing(node) and not instr and not node.xpath(".//w:br", namespaces=NS):
            empty_paras.append(idx)
        has_forced_break = node.xpath("./w:pPr/w:pageBreakBefore | .//w:br[@w:type='page']", namespaces=NS)
        if has_forced_break and st not in {"af", "af1", "afd", "aff", "4"}:
            page_breaks.append({"index": idx, "text": text[:160]})
        jc = node.xpath("./w:pPr/w:jc/@w:val", namespaces=NS)
        if text and not has_drawing(node) and st not in {"af", "af1", "afd", "aff"} and jc and jc[0] == "center":
            centered_body.append({"index": idx, "style": st, "text": text[:160]})
        if "\u672c\u6b21\u4fee\u6539" in text:
            forbidden.append({"index": idx, "text": text[:180]})
        if "Error! Reference source not found" in text or "\u9519\u8bef" in text and "\u5f15\u7528" in text:
            failures.append({"type": "broken_reference_visible_text", "index": idx, "text": text[:180]})

        is_caption = st in {"afd", "aff"} or re.match(r"^[\u56fe\u8868]\s*5-\d+", text)
        if is_caption:
            captions.append({"index": idx, "style": st, "text": text, "instr": instr})
            if node.xpath("./w:pPr/w:keepNext", namespaces=NS):
                failures.append({"type": "caption_direct_keepNext", "index": idx, "text": text[:160]})

        if is_standalone_formula(node, text):
            fr = formula_report(idx, node)
            formulas.append(fr)
            if not fr["has_omath"]:
                failures.append({"type": "formula_not_omml", "index": idx, "text": text})
            if fr["line"] != "240" or fr["lineRule"] != "auto":
                failures.append({"type": "formula_not_single_line_spacing", "index": idx, "line": fr["line"], "lineRule": fr["lineRule"], "text": text})
            if set(fr["sizes"]) - {"24"}:
                failures.append({"type": "formula_size_not_xiaosi", "index": idx, "sizes": fr["sizes"], "text": text})

        if not is_caption and text and not has_drawing(node):
            if raw_inline_rx.search(text):
                raw_inline.append({"index": idx, "style": st, "text": text[:220]})
            if figure_table_ref_rx.search(text) and not has_ref_field(node):
                plain_refs.append({"index": idx, "text": text[:220]})

    if caption_keep_styles:
        failures.append({"type": "caption_style_keepNext", "items": caption_keep_styles})
    if empty_paras:
        failures.append({"type": "empty_paragraphs", "count": len(empty_paras), "indices": empty_paras[:80]})
    if page_breaks:
        failures.append({"type": "manual_or_forced_page_breaks", "items": page_breaks})
    if centered_body:
        failures.append({"type": "centered_body_paragraphs", "items": centered_body})
    if forbidden:
        failures.append({"type": "forbidden_or_defensive_phrases", "items": forbidden})
    if plain_refs:
        failures.append({"type": "plain_figure_table_references", "count": len(plain_refs), "items": plain_refs})
    if raw_inline:
        failures.append({"type": "raw_inline_math_text", "count": len(raw_inline), "items": raw_inline[:120]})

    return {
        "docx": str(path),
        "chapter_start_index": ch5[0][0],
        "chapter_end_index": ch5[-1][0],
        "block_count": len(ch5),
        "formula_count": len(formulas),
        "table_count": len(tables),
        "caption_count": len(captions),
        "formulas": formulas,
        "tables": tables,
        "captions": captions,
        "failures": failures,
    }


def main() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    parser = argparse.ArgumentParser()
    parser.add_argument("docx")
    parser.add_argument("--out")
    args = parser.parse_args()
    report = audit_docx(Path(args.docx))
    text = json.dumps(report, ensure_ascii=False, indent=2)
    if args.out:
        Path(args.out).write_text(text, encoding="utf-8")
    print(text)


if __name__ == "__main__":
    main()
