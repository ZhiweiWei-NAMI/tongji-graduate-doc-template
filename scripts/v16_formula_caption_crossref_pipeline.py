from __future__ import annotations

import argparse
import csv
import json
import re
import shutil
import tempfile
import zipfile
from dataclasses import dataclass
from pathlib import Path

from lxml import etree


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

W = NS["w"]
M = NS["m"]


def qn(tag: str) -> str:
    prefix, local = tag.split(":", 1)
    return f"{{{NS[prefix]}}}{local}"


def wval(value: str) -> str:
    return f"{{{W}}}{value}"


def xml_text(element: etree._Element) -> str:
    parts: list[str] = []
    for node in element.xpath(".//w:t | .//m:t", namespaces=NS):
        if node.text:
            parts.append(node.text)
    return "".join(parts)


def paragraph_text(paragraph: etree._Element) -> str:
    parts: list[str] = []
    for node in paragraph.xpath(".//w:t", namespaces=NS):
        if node.text:
            parts.append(node.text)
    return "".join(parts)


def paragraph_math_text(paragraph: etree._Element) -> str:
    parts: list[str] = []
    for node in paragraph.xpath(".//m:t", namespaces=NS):
        if node.text:
            parts.append(node.text)
    return "".join(parts)


def paragraph_style(paragraph: etree._Element) -> str:
    node = paragraph.find("./w:pPr/w:pStyle", namespaces=NS)
    return node.get(wval("val"), "") if node is not None else ""


def paragraph_outline_level(paragraph: etree._Element) -> int | None:
    node = paragraph.find("./w:pPr/w:outlineLvl", namespaces=NS)
    if node is None:
        return None
    raw = node.get(wval("val"))
    if raw is None:
        return None
    try:
        return int(raw)
    except ValueError:
        return None


def field_instrs(element: etree._Element) -> list[str]:
    instrs: list[str] = []
    for node in element.xpath(".//w:fldSimple", namespaces=NS):
        instr = node.get(wval("instr"))
        if instr:
            instrs.append(instr)
    current: list[str] | None = None
    for node in element.iter():
        if node.tag == qn("w:fldChar"):
            fld_type = node.get(wval("fldCharType"))
            if fld_type == "begin":
                current = []
            elif fld_type == "separate":
                if current:
                    instrs.append("".join(current))
                current = None
            elif fld_type == "end":
                current = None
        elif node.tag == qn("w:instrText") and current is not None and node.text:
            current.append(node.text)
    return instrs


def bookmarks(element: etree._Element) -> list[str]:
    names: list[str] = []
    for node in element.xpath(".//w:bookmarkStart", namespaces=NS):
        name = node.get(wval("name"))
        if name:
            names.append(name)
    return names


def load_xml_from_docx(docx_path: Path) -> tuple[etree._Element, etree._Element | None]:
    with zipfile.ZipFile(docx_path, "r") as zf:
        document = etree.fromstring(zf.read("word/document.xml"))
        styles = None
        if "word/styles.xml" in zf.namelist():
            styles = etree.fromstring(zf.read("word/styles.xml"))
    return document, styles


def style_outline_levels(styles_root: etree._Element | None) -> dict[str, int]:
    result: dict[str, int] = {}
    if styles_root is None:
        return result
    for style in styles_root.xpath(".//w:style[@w:type='paragraph']", namespaces=NS):
        sid = style.get(wval("styleId"), "")
        outline = style.find("./w:pPr/w:outlineLvl", namespaces=NS)
        if sid and outline is not None:
            raw = outline.get(wval("val"))
            if raw is not None:
                try:
                    result[sid] = int(raw)
                except ValueError:
                    pass
    return result


def chinese_num_to_int(text: str) -> int | None:
    text = text.strip()
    if text.isdigit():
        return int(text)
    digits = {"零": 0, "一": 1, "二": 2, "两": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7, "八": 8, "九": 9}
    if text in digits:
        return digits[text]
    if text == "十":
        return 10
    if "十" in text:
        left, right = text.split("十", 1)
        tens = digits.get(left, 1) if left else 1
        ones = digits.get(right, 0) if right else 0
        return tens * 10 + ones
    return None


def heading_chapter(text: str) -> int | None:
    match = re.search(r"第\s*([0-9一二三四五六七八九十两]+)\s*章", text)
    if match:
        return chinese_num_to_int(match.group(1))
    return None


def is_heading1(paragraph: etree._Element, style_levels: dict[str, int]) -> bool:
    direct = paragraph_outline_level(paragraph)
    if direct == 0:
        return True
    style = paragraph_style(paragraph)
    if style_levels.get(style) == 0:
        return True
    return False


def has_table_ancestor(paragraph: etree._Element) -> bool:
    parent = paragraph.getparent()
    while parent is not None:
        if parent.tag == qn("w:tbl"):
            return True
        parent = parent.getparent()
    return False


CAPTION_RE = re.compile(r"^\s*(图|表)\s*([0-9]+)\s*(?:[-－–—.．]\s*)?([0-9]+)\s*(.*)$")
TEXT_REF_RE = re.compile(r"(?<![第])([图表])\s*([0-9]+)\s*[-－–—.．]\s*([0-9]+)")
EQ_REF_RE = re.compile(r"(公式|式)\s*[\(（]?\s*([0-9]+)\s*[-－–—.．]\s*([0-9]+)\s*[\)）]?")
EQ_NUMBER_RE = re.compile(r"[\(（]\s*([0-9]+)\s*[-－–—.．]\s*([0-9]+)\s*[\)）]\s*$")


def is_true_caption(paragraph: etree._Element, text: str) -> bool:
    if not CAPTION_RE.match(text.strip()):
        return False
    style = paragraph_style(paragraph)
    return style in {"afd", "Caption", "caption"}


def is_standalone_math(paragraph: etree._Element) -> bool:
    math = paragraph_math_text(paragraph).strip()
    if not math:
        return False
    text = paragraph_text(paragraph).strip()
    if paragraph.xpath(".//w:drawing | .//w:pict", namespaces=NS):
        return False
    if text and len(text) > 16:
        return False
    if paragraph.xpath(".//m:oMathPara", namespaces=NS):
        return True
    if len(math) >= 2 and (not text or EQ_NUMBER_RE.search(text)):
        return True
    return False


@dataclass
class ParagraphRecord:
    index: int
    chapter: int | None
    style: str
    text: str
    math_text: str
    fields: list[str]
    bookmark_names: list[str]
    in_table: bool
    drawing_count: int
    kind: str


def is_toc_style(style: str) -> bool:
    return style.upper().startswith("TOC")


def is_frontmatter_heading(text: str) -> bool:
    compact = re.sub(r"\s+", "", text).upper()
    return compact in {"摘要", "ABSTRACT", "目录", "CONTENTS"}


def is_non_chapter_heading(text: str) -> bool:
    compact = re.sub(r"\s+", "", text)
    return compact in {"参考文献", "致谢", "攻读博士学位期间发表的论文及其他成果"}


def build_records(document: etree._Element, styles_root: etree._Element | None) -> list[ParagraphRecord]:
    style_levels = style_outline_levels(styles_root)
    chapter: int | None = None
    chapter_counter = 0
    body_started = False
    records: list[ParagraphRecord] = []
    paragraphs = document.xpath(".//w:body//w:p", namespaces=NS)
    for idx, paragraph in enumerate(paragraphs):
        text = paragraph_text(paragraph)
        style = paragraph_style(paragraph)
        if is_heading1(paragraph, style_levels) and not is_toc_style(style):
            detected = heading_chapter(text)
            if detected is not None:
                chapter = detected
                chapter_counter = max(chapter_counter, detected)
                body_started = True
            elif not is_frontmatter_heading(text) and not is_non_chapter_heading(text):
                if body_started or idx > 200:
                    chapter_counter += 1
                    chapter = chapter_counter
                    body_started = True
        math = paragraph_math_text(paragraph)
        fields = field_instrs(paragraph)
        if is_true_caption(paragraph, text):
            kind = "caption"
        elif is_standalone_math(paragraph):
            kind = "equation"
        else:
            kind = "body"
        records.append(
            ParagraphRecord(
                index=idx,
                chapter=chapter,
                style=style,
                text=text,
                math_text=math,
                fields=fields,
                bookmark_names=bookmarks(paragraph),
                in_table=has_table_ancestor(paragraph),
                drawing_count=len(paragraph.xpath(".//w:drawing | .//w:pict", namespaces=NS)),
                kind=kind,
            )
        )
    return records


def summarize_fields(records: list[ParagraphRecord]) -> dict[str, int]:
    result = {"SEQ": 0, "REF": 0, "PAGEREF": 0, "STYLEREF": 0, "other": 0}
    for record in records:
        for instr in record.fields:
            normalized = instr.strip().upper()
            if normalized.startswith("SEQ "):
                result["SEQ"] += 1
            elif normalized.startswith("REF "):
                result["REF"] += 1
            elif normalized.startswith("PAGEREF "):
                result["PAGEREF"] += 1
            elif normalized.startswith("STYLEREF "):
                result["STYLEREF"] += 1
            elif normalized:
                result["other"] += 1
    return result


def audit_docx(docx_path: Path, out_dir: Path) -> dict[str, object]:
    document, styles = load_xml_from_docx(docx_path)
    records = build_records(document, styles)
    style_levels = style_outline_levels(styles)
    heading_candidates: list[dict[str, object]] = []
    captions: list[dict[str, object]] = []
    equations: list[dict[str, object]] = []
    references: list[dict[str, object]] = []
    paragraphs = document.xpath(".//w:body//w:p", namespaces=NS)
    for paragraph, record in zip(paragraphs, records):
        text = record.text.strip()
        direct_outline = paragraph_outline_level(paragraph)
        style_outline = style_levels.get(record.style)
        if direct_outline is not None or style_outline is not None or heading_chapter(text) is not None:
            heading_candidates.append(
                {
                    "paragraph": record.index,
                    "chapter_context": record.chapter,
                    "style": record.style,
                    "direct_outline": "" if direct_outline is None else direct_outline,
                    "style_outline": "" if style_outline is None else style_outline,
                    "detected_chapter": "" if heading_chapter(text) is None else heading_chapter(text),
                    "text": text[:240],
                }
            )
        if record.kind == "caption":
            match = CAPTION_RE.match(text)
            assert match is not None
            captions.append(
                {
                    "paragraph": record.index,
                    "chapter_context": record.chapter,
                    "label": match.group(1),
                    "caption_chapter": int(match.group(2)),
                    "caption_number": int(match.group(3)),
                    "title": match.group(4).strip(),
                    "style": record.style,
                    "has_seq": any(instr.strip().upper().startswith("SEQ ") for instr in record.fields),
                    "bookmarks": ";".join(record.bookmark_names),
                }
            )
        if record.kind == "equation":
            joined = (record.text + record.math_text).strip()
            num_match = EQ_NUMBER_RE.search(joined)
            equations.append(
                {
                    "paragraph": record.index,
                    "chapter_context": record.chapter,
                    "style": record.style,
                    "has_seq": any(instr.strip().upper().startswith("SEQ ") for instr in record.fields),
                    "has_ref_bookmark": any(name.startswith("eq_") for name in record.bookmark_names),
                    "manual_number": num_match.group(0) if num_match else "",
                    "math_text": record.math_text[:220],
                    "text": record.text[:120],
                }
            )
        if record.kind != "caption":
            has_ref_field = any(instr.strip().upper().startswith("REF ") for instr in record.fields)
            for match in TEXT_REF_RE.finditer(text):
                references.append(
                    {
                        "paragraph": record.index,
                        "chapter_context": record.chapter,
                        "kind": "figure" if match.group(1) == "图" else "table",
                        "visible": match.group(0),
                        "target_chapter": int(match.group(2)),
                        "target_number": int(match.group(3)),
                        "has_ref_field_in_paragraph": has_ref_field,
                        "text": text[:240],
                    }
                )
            for match in EQ_REF_RE.finditer(text):
                references.append(
                    {
                        "paragraph": record.index,
                        "chapter_context": record.chapter,
                        "kind": "equation",
                        "visible": match.group(0),
                        "target_chapter": int(match.group(2)),
                        "target_number": int(match.group(3)),
                        "has_ref_field_in_paragraph": has_ref_field,
                        "text": text[:240],
                    }
                )
    out_dir.mkdir(parents=True, exist_ok=True)
    write_csv(out_dir / "04_heading_candidates.csv", heading_candidates)
    write_csv(out_dir / "01_captions.csv", captions)
    write_csv(out_dir / "02_equations.csv", equations)
    write_csv(out_dir / "03_references.csv", references)
    summary = {
        "docx": str(docx_path),
        "paragraphs": len(records),
        "field_summary": summarize_fields(records),
        "caption_count": len(captions),
        "figure_caption_count": sum(1 for row in captions if row["label"] == "图"),
        "table_caption_count": sum(1 for row in captions if row["label"] == "表"),
        "caption_without_seq": sum(1 for row in captions if not row["has_seq"]),
        "equation_count": len(equations),
        "equation_without_seq": sum(1 for row in equations if not row["has_seq"]),
        "equation_without_manual_number": sum(1 for row in equations if not row["manual_number"]),
        "detected_reference_count": len(references),
        "detected_reference_without_ref_field": sum(1 for row in references if not row["has_ref_field_in_paragraph"]),
    }
    (out_dir / "00_audit_summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    write_markdown_summary(out_dir / "00_audit_summary.md", summary, captions, equations, references)
    return summary


def write_csv(path: Path, rows: list[dict[str, object]]) -> None:
    if not rows:
        path.write_text("", encoding="utf-8-sig")
        return
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)


def write_markdown_summary(
    path: Path,
    summary: dict[str, object],
    captions: list[dict[str, object]],
    equations: list[dict[str, object]],
    references: list[dict[str, object]],
) -> None:
    lines = [
        "# Formula, Caption, and Cross-Reference Audit",
        "",
        f"- 文档：`{summary['docx']}`",
        f"- 段落数：{summary['paragraphs']}",
        f"- 字段统计：{summary['field_summary']}",
        f"- 题注：{summary['caption_count']}（图 {summary['figure_caption_count']}，表 {summary['table_caption_count']}）",
        f"- 无 SEQ 字段题注：{summary['caption_without_seq']}",
        f"- 独立公式段：{summary['equation_count']}",
        f"- 无 SEQ 字段公式：{summary['equation_without_seq']}",
        f"- 未检测到人工编号的公式：{summary['equation_without_manual_number']}",
        f"- 静态图/表/式引用：{summary['detected_reference_count']}",
        f"- 所在段落无 REF 字段的静态引用：{summary['detected_reference_without_ref_field']}",
        "",
        "## 样例题注",
        "",
    ]
    for row in captions[:20]:
        lines.append(f"- P{row['paragraph']}: {row['label']}{row['caption_chapter']}-{row['caption_number']} {row['title']}")
    lines.extend(["", "## 样例公式", ""])
    for row in equations[:20]:
        marker = row["manual_number"] or "未编号"
        formula = str(row["math_text"]).replace("\n", " ")[:120]
        lines.append(f"- P{row['paragraph']}: {marker} `{formula}`")
    lines.extend(["", "## 样例引用", ""])
    for row in references[:30]:
        lines.append(f"- P{row['paragraph']}: {row['visible']} / {row['kind']} / REF={row['has_ref_field_in_paragraph']}")
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def clone_rpr(run: etree._Element | None) -> etree._Element | None:
    if run is None:
        return None
    rpr = run.find("./w:rPr", namespaces=NS)
    return etree.fromstring(etree.tostring(rpr)) if rpr is not None else None


def make_run(text: str = "", rpr: etree._Element | None = None) -> etree._Element:
    run = etree.Element(qn("w:r"))
    if rpr is not None:
        run.append(etree.fromstring(etree.tostring(rpr)))
    if text:
        t = etree.SubElement(run, qn("w:t"))
        if text[:1].isspace() or text[-1:].isspace():
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t.text = text
    return run


def make_tab_run(rpr: etree._Element | None = None) -> etree._Element:
    run = etree.Element(qn("w:r"))
    if rpr is not None:
        run.append(etree.fromstring(etree.tostring(rpr)))
    etree.SubElement(run, qn("w:tab"))
    return run


def make_fld_simple(instr: str, result: str, rpr: etree._Element | None = None) -> etree._Element:
    field = etree.Element(qn("w:fldSimple"))
    field.set(wval("instr"), instr)
    field.append(make_run(result, rpr))
    return field


def make_bookmark_start(bookmark_id: int, name: str) -> etree._Element:
    node = etree.Element(qn("w:bookmarkStart"))
    node.set(wval("id"), str(bookmark_id))
    node.set(wval("name"), name)
    return node


def make_bookmark_end(bookmark_id: int) -> etree._Element:
    node = etree.Element(qn("w:bookmarkEnd"))
    node.set(wval("id"), str(bookmark_id))
    return node


def get_or_add(parent: etree._Element, tag: str) -> etree._Element:
    child = parent.find(f"./{tag}", namespaces=NS)
    if child is None:
        child = etree.Element(qn(tag))
        parent.insert(0, child)
    return child


def ensure_right_tab(paragraph: etree._Element, pos: int = 8312) -> None:
    ppr = paragraph.find("./w:pPr", namespaces=NS)
    if ppr is None:
        ppr = etree.Element(qn("w:pPr"))
        paragraph.insert(0, ppr)
    tabs = ppr.find("./w:tabs", namespaces=NS)
    if tabs is None:
        tabs = etree.Element(qn("w:tabs"))
        insert_at = 0
        paragraph_style_node = ppr.find("./w:pStyle", namespaces=NS)
        if paragraph_style_node is not None:
            insert_at = list(ppr).index(paragraph_style_node) + 1
        ppr.insert(insert_at, tabs)
    for tab in list(tabs):
        if tab.get(wval("val")) == "right":
            tabs.remove(tab)
    tab = etree.SubElement(tabs, qn("w:tab"))
    tab.set(wval("val"), "right")
    tab.set(wval("pos"), str(pos))


def clear_paragraph_content(paragraph: etree._Element) -> None:
    for child in list(paragraph):
        if child.tag != qn("w:pPr"):
            paragraph.remove(child)


def next_bookmark_id(document: etree._Element) -> int:
    max_id = 0
    for node in document.xpath(".//w:bookmarkStart | .//w:bookmarkEnd", namespaces=NS):
        raw = node.get(wval("id"))
        if raw and raw.isdigit():
            max_id = max(max_id, int(raw))
    return max_id + 1


def strip_trailing_manual_equation_number(paragraph: etree._Element) -> bool:
    nodes = list(paragraph.xpath(".//m:t", namespaces=NS))
    if not nodes:
        return False
    pieces = [node.text or "" for node in nodes]
    full = "".join(pieces)
    match = re.search(r"[\s\u2000-\u200A\u3000]*[\(（]\s*[0-9]+\s*[-－–—.．]\s*[0-9]+\s*[\)）]\s*$", full)
    if not match:
        return False
    remove_start = match.start()
    offset = 0
    changed = False
    for node, piece in zip(nodes, pieces):
        start = offset
        end = offset + len(piece)
        offset = end
        if end <= remove_start:
            continue
        keep_len = max(0, remove_start - start)
        new_text = piece[:keep_len]
        if node.text != new_text:
            node.text = new_text
            changed = True
    return changed


def append_equation_number(
    paragraph: etree._Element,
    chapter: int,
    seq: int,
    bookmark_name: str,
    bookmark_id: int,
) -> None:
    ensure_right_tab(paragraph)
    first_run = paragraph.find(".//w:r", namespaces=NS)
    rpr = clone_rpr(first_run)
    paragraph.append(make_tab_run(rpr))
    paragraph.append(make_bookmark_start(bookmark_id, bookmark_name))
    paragraph.append(make_run("（", rpr))
    paragraph.append(make_run(f"{chapter}-", rpr))
    paragraph.append(make_fld_simple(" SEQ Equation \\* ARABIC \\s 1 ", str(seq), rpr))
    paragraph.append(make_run("）", rpr))
    paragraph.append(make_bookmark_end(bookmark_id))


def rewrite_caption(
    paragraph: etree._Element,
    label: str,
    chapter: int,
    seq: int,
    title: str,
    bookmark_name: str,
    bookmark_id: int,
) -> None:
    first_run = paragraph.find(".//w:r", namespaces=NS)
    rpr = clone_rpr(first_run)
    clear_paragraph_content(paragraph)
    seq_name = "Figure" if label == "图" else "Table"
    paragraph.append(make_bookmark_start(bookmark_id, bookmark_name))
    paragraph.append(make_run(f"{label} {chapter}-", rpr))
    paragraph.append(make_fld_simple(f" SEQ {seq_name} \\* ARABIC \\s 1 ", str(seq), rpr))
    paragraph.append(make_bookmark_end(bookmark_id))
    if title:
        paragraph.append(make_run(f" {title}", rpr))


def has_ancestor(node: etree._Element, tag: str) -> bool:
    parent = node.getparent()
    target = qn(tag)
    while parent is not None:
        if parent.tag == target:
            return True
        parent = parent.getparent()
    return False


def replace_run_with_nodes(run: etree._Element, new_nodes: list[etree._Element]) -> None:
    parent = run.getparent()
    if parent is None:
        return
    index = parent.index(run)
    parent.remove(run)
    for offset, node in enumerate(new_nodes):
        parent.insert(index + offset, node)


def find_next_match(text: str, start: int = 0):
    candidates = []
    text_match = TEXT_REF_RE.search(text, start)
    if text_match is not None:
        candidates.append(("object", text_match))
    eq_match = EQ_REF_RE.search(text, start)
    if eq_match is not None:
        candidates.append(("equation", eq_match))
    if not candidates:
        return None
    return min(candidates, key=lambda item: item[1].start())


def choose_nearest(candidates: list[dict[str, object]], paragraph_index: int) -> dict[str, object] | None:
    if not candidates:
        return None
    return min(candidates, key=lambda item: (abs(int(item["paragraph"]) - paragraph_index), int(item["paragraph"]) > paragraph_index))


def is_text_only_run(run: etree._Element) -> bool:
    if run.tag != qn("w:r"):
        return False
    blocked = run.xpath(
        ".//w:fldChar | .//w:instrText | .//w:drawing | .//w:pict | .//w:tab | .//w:br | .//w:object | .//w:footnoteReference | .//w:endnoteReference",
        namespaces=NS,
    )
    if blocked:
        return False
    return bool(run.xpath(".//w:t", namespaces=NS))


def run_visible_text(run: etree._Element) -> str:
    return "".join(node.text or "" for node in run.xpath(".//w:t", namespaces=NS))


def reference_nodes_from_text(
    original: str,
    paragraph_index: int,
    rpr: etree._Element | None,
    caption_lookup: dict[tuple[str, int, int], list[dict[str, object]]],
    equation_lookup: dict[tuple[int, int], dict[str, object]],
    report: dict[str, object],
) -> tuple[list[etree._Element], bool]:
    pos = 0
    pieces: list[etree._Element] = []
    changed = False
    while True:
        found = find_next_match(original, pos)
        if found is None:
            break
        kind, match = found
        if match.start() > pos:
            pieces.append(make_run(original[pos : match.start()], rpr))
        if kind == "object":
            label = match.group(1)
            chapter = int(match.group(2))
            number = int(match.group(3))
            object_kind = "figure" if label == "图" else "table"
            target = choose_nearest(caption_lookup.get((object_kind, chapter, number), []), paragraph_index)
            if target is None:
                pieces.append(make_run(match.group(0), rpr))
                report["unresolved_references"].append(
                    {
                        "paragraph": paragraph_index,
                        "kind": object_kind,
                        "visible": match.group(0),
                        "reason": "no matching caption",
                    }
                )
            else:
                pieces.append(make_fld_simple(f" REF {target['bookmark']} \\h ", str(target["visible"]), rpr))
                report["reference_fields"] += 1
                changed = True
        else:
            prefix = match.group(1)
            chapter = int(match.group(2))
            number = int(match.group(3))
            target = equation_lookup.get((chapter, number))
            if target is None:
                pieces.append(make_run(match.group(0), rpr))
                report["unresolved_references"].append(
                    {
                        "paragraph": paragraph_index,
                        "kind": "equation",
                        "visible": match.group(0),
                        "reason": "no matching original equation number",
                    }
                )
            else:
                pieces.append(make_run(prefix, rpr))
                pieces.append(make_fld_simple(f" REF {target['bookmark']} \\h ", str(target["visible"]), rpr))
                report["reference_fields"] += 1
                changed = True
        pos = match.end()
    if pos < len(original):
        pieces.append(make_run(original[pos:], rpr))
    return pieces, changed


def replace_references_across_text_runs(
    paragraph: etree._Element,
    paragraph_index: int,
    caption_lookup: dict[tuple[str, int, int], list[dict[str, object]]],
    equation_lookup: dict[tuple[int, int], dict[str, object]],
    report: dict[str, object],
) -> None:
    new_children: list[etree._Element] = []
    buffer = ""
    buffer_rpr: etree._Element | None = None
    changed = False

    def flush_buffer() -> None:
        nonlocal buffer, buffer_rpr, changed
        if not buffer:
            return
        pieces, buffer_changed = reference_nodes_from_text(
            buffer, paragraph_index, buffer_rpr, caption_lookup, equation_lookup, report
        )
        new_children.extend(pieces)
        changed = changed or buffer_changed
        buffer = ""
        buffer_rpr = None

    for child in list(paragraph):
        if child.tag == qn("w:pPr"):
            continue
        if is_text_only_run(child):
            if buffer_rpr is None:
                buffer_rpr = clone_rpr(child)
            buffer += run_visible_text(child)
        else:
            flush_buffer()
            new_children.append(child)
    flush_buffer()
    if not changed:
        return
    clear_paragraph_content(paragraph)
    for node in new_children:
        paragraph.append(node)


def replace_reference_text_nodes(
    paragraph: etree._Element,
    paragraph_index: int,
    caption_lookup: dict[tuple[str, int, int], list[dict[str, object]]],
    equation_lookup: dict[tuple[int, int], dict[str, object]],
    report: dict[str, object],
) -> None:
    if is_true_caption(paragraph, paragraph_text(paragraph)) or is_standalone_math(paragraph):
        return
    changed_any = False
    text_nodes = list(paragraph.xpath(".//w:t", namespaces=NS))
    for text_node in text_nodes:
        if has_ancestor(text_node, "w:fldSimple"):
            continue
        run = text_node.getparent()
        if run is None or run.tag != qn("w:r"):
            continue
        original = text_node.text or ""
        if not original:
            continue
        pos = 0
        pieces: list[etree._Element] = []
        changed = False
        rpr = clone_rpr(run)
        while True:
            found = find_next_match(original, pos)
            if found is None:
                break
            kind, match = found
            if match.start() > pos:
                pieces.append(make_run(original[pos : match.start()], rpr))
            if kind == "object":
                label = match.group(1)
                chapter = int(match.group(2))
                number = int(match.group(3))
                object_kind = "figure" if label == "图" else "table"
                target = choose_nearest(caption_lookup.get((object_kind, chapter, number), []), paragraph_index)
                if target is None:
                    pieces.append(make_run(match.group(0), rpr))
                    report["unresolved_references"].append(
                        {
                            "paragraph": paragraph_index,
                            "kind": object_kind,
                            "visible": match.group(0),
                            "reason": "no matching caption",
                        }
                    )
                else:
                    pieces.append(make_fld_simple(f" REF {target['bookmark']} \\h ", str(target["visible"]), rpr))
                    report["reference_fields"] += 1
                    changed = True
            else:
                prefix = match.group(1)
                chapter = int(match.group(2))
                number = int(match.group(3))
                target = equation_lookup.get((chapter, number))
                if target is None:
                    pieces.append(make_run(match.group(0), rpr))
                    report["unresolved_references"].append(
                        {
                            "paragraph": paragraph_index,
                            "kind": "equation",
                            "visible": match.group(0),
                            "reason": "no matching original equation number",
                        }
                    )
                else:
                    pieces.append(make_run(prefix, rpr))
                    pieces.append(make_fld_simple(f" REF {target['bookmark']} \\h ", str(target["visible"]), rpr))
                    report["reference_fields"] += 1
                    changed = True
            pos = match.end()
        if not changed:
            continue
        if pos < len(original):
            pieces.append(make_run(original[pos:], rpr))
        replace_run_with_nodes(run, pieces)
        changed_any = True
    if not changed_any:
        replace_references_across_text_runs(paragraph, paragraph_index, caption_lookup, equation_lookup, report)


def is_page_break_only_paragraph(paragraph: etree._Element) -> bool:
    if paragraph_text(paragraph).strip() or paragraph_math_text(paragraph).strip():
        return False
    breaks = paragraph.xpath("./w:r/w:br", namespaces=NS)
    if len(breaks) != 1:
        return False
    break_type = breaks[0].get(wval("type"))
    if break_type not in {None, "page"}:
        return False
    other_content = [
        child
        for child in paragraph
        if child.tag != qn("w:pPr") and not (child.tag == qn("w:r") and child.xpath("./w:br", namespaces=NS))
    ]
    return not other_content


def remove_redundant_page_breaks_before_chapters(
    paragraphs: list[etree._Element],
    records: list[ParagraphRecord],
) -> int:
    removed = 0
    for idx, (paragraph, record) in enumerate(zip(paragraphs, records)):
        if record.style != "1" or record.text.strip() != "总结与展望" or idx == 0:
            continue
        previous = paragraphs[idx - 1]
        if is_page_break_only_paragraph(previous):
            parent = previous.getparent()
            if parent is not None:
                parent.remove(previous)
                removed += 1
    return removed


def ensure_update_fields_setting(tmp_path: Path) -> None:
    settings_path = tmp_path / "word" / "settings.xml"
    if not settings_path.exists():
        return
    root = etree.fromstring(settings_path.read_bytes())
    existing = root.find("./w:updateFields", namespaces=NS)
    if existing is None:
        existing = etree.Element(qn("w:updateFields"))
        root.insert(0, existing)
    existing.set(wval("val"), "true")
    settings_path.write_bytes(etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes"))


def repack_docx(tmp_path: Path, output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for item in sorted(tmp_path.rglob("*")):
            if item.is_file():
                zf.write(item, item.relative_to(tmp_path).as_posix())


def repair_docx(input_path: Path, output_path: Path, out_dir: Path) -> dict[str, object]:
    out_dir.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="v16_formula_caption_crossref_") as tmp_name:
        tmp_path = Path(tmp_name)
        with zipfile.ZipFile(input_path, "r") as zf:
            zf.extractall(tmp_path)
        document_path = tmp_path / "word" / "document.xml"
        styles_path = tmp_path / "word" / "styles.xml"
        document = etree.fromstring(document_path.read_bytes())
        styles = etree.fromstring(styles_path.read_bytes()) if styles_path.exists() else None
        records = build_records(document, styles)
        paragraphs = document.xpath(".//w:body//w:p", namespaces=NS)
        bookmark_id = next_bookmark_id(document)
        caption_sequences: dict[tuple[str, int], int] = {}
        equation_sequences: dict[int, int] = {}
        caption_lookup: dict[tuple[str, int, int], list[dict[str, object]]] = {}
        equation_lookup: dict[tuple[int, int], dict[str, object]] = {}
        report: dict[str, object] = {
            "input": str(input_path),
            "output": str(output_path),
            "caption_fields": 0,
            "equation_fields": 0,
            "manual_equation_numbers_removed": 0,
            "redundant_page_breaks_removed": 0,
            "reference_fields": 0,
            "unresolved_references": [],
            "caption_targets": [],
            "equation_targets": [],
        }

        for paragraph, record in zip(paragraphs, records):
            if record.kind != "caption":
                continue
            match = CAPTION_RE.match(record.text.strip())
            if match is None:
                continue
            label = match.group(1)
            original_chapter = int(match.group(2))
            original_number = int(match.group(3))
            chapter = original_chapter
            sequence_key = (label, chapter)
            caption_sequences[sequence_key] = caption_sequences.get(sequence_key, 0) + 1
            seq = caption_sequences[sequence_key]
            object_kind = "figure" if label == "图" else "table"
            prefix = "fig" if object_kind == "figure" else "tbl"
            bookmark = f"{prefix}_{chapter}_{seq}"
            visible = f"{label} {chapter}-{seq}"
            rewrite_caption(paragraph, label, chapter, seq, match.group(4).strip(), bookmark, bookmark_id)
            bookmark_id += 1
            item = {
                "paragraph": record.index,
                "kind": object_kind,
                "original_chapter": original_chapter,
                "original_number": original_number,
                "chapter": chapter,
                "number": seq,
                "bookmark": bookmark,
                "visible": visible,
                "title": match.group(4).strip(),
            }
            caption_lookup.setdefault((object_kind, original_chapter, original_number), []).append(item)
            report["caption_targets"].append(item)
            report["caption_fields"] += 1

        for paragraph, record in zip(paragraphs, records):
            if record.kind != "equation" or record.chapter is None:
                continue
            chapter = int(record.chapter)
            equation_sequences[chapter] = equation_sequences.get(chapter, 0) + 1
            seq = equation_sequences[chapter]
            joined = (record.text + record.math_text).strip()
            manual_match = EQ_NUMBER_RE.search(joined)
            original_key = None
            if manual_match is not None:
                original_key = (int(manual_match.group(1)), int(manual_match.group(2)))
                if strip_trailing_manual_equation_number(paragraph):
                    report["manual_equation_numbers_removed"] += 1
            bookmark = f"eq_{chapter}_{seq}"
            visible = f"（{chapter}-{seq}）"
            append_equation_number(paragraph, chapter, seq, bookmark, bookmark_id)
            bookmark_id += 1
            item = {
                "paragraph": record.index,
                "chapter": chapter,
                "number": seq,
                "bookmark": bookmark,
                "visible": visible,
                "original_number": "" if original_key is None else f"{original_key[0]}-{original_key[1]}",
                "math_text": record.math_text[:180],
            }
            if original_key is not None:
                equation_lookup[original_key] = item
            report["equation_targets"].append(item)
            report["equation_fields"] += 1

        for paragraph, record in zip(paragraphs, records):
            replace_reference_text_nodes(paragraph, record.index, caption_lookup, equation_lookup, report)

        report["redundant_page_breaks_removed"] = remove_redundant_page_breaks_before_chapters(paragraphs, records)

        document_path.write_bytes(etree.tostring(document, xml_declaration=True, encoding="UTF-8", standalone="yes"))
        ensure_update_fields_setting(tmp_path)
        repack_docx(tmp_path, output_path)

    report_path = out_dir / "06_repair_report.json"
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    write_repair_summary(out_dir / "06_repair_report.md", report)
    return report


def write_repair_summary(path: Path, report: dict[str, object]) -> None:
    unresolved = report.get("unresolved_references", [])
    lines = [
        "# Fieldization Repair Report",
        "",
        f"- 输入：`{report['input']}`",
        f"- 输出：`{report['output']}`",
        f"- 题注字段化：{report['caption_fields']}",
        f"- 公式字段化：{report['equation_fields']}",
        f"- 移除公式内手工尾号：{report['manual_equation_numbers_removed']}",
        f"- 移除冗余分页符：{report.get('redundant_page_breaks_removed', 0)}",
        f"- 正文 REF 引用替换：{report['reference_fields']}",
        f"- 未解析引用：{len(unresolved)}",
        "",
    ]
    if unresolved:
        lines.append("## 未解析引用")
        lines.append("")
        for item in unresolved[:80]:
            lines.append(f"- P{item['paragraph']}: {item['visible']}（{item['reason']}）")
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True)
    parser.add_argument("--out-dir", required=True)
    parser.add_argument("--output")
    parser.add_argument("--mode", choices=["audit", "repair"], default="audit")
    args = parser.parse_args()
    if args.mode == "audit":
        result = audit_docx(Path(args.input), Path(args.out_dir))
    else:
        if not args.output:
            raise SystemExit("--output is required in repair mode")
        result = repair_docx(Path(args.input), Path(args.output), Path(args.out_dir))
    print(json.dumps(result, ensure_ascii=True, indent=2))


if __name__ == "__main__":
    main()
