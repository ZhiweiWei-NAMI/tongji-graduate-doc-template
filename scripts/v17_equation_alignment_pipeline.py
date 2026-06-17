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
}

W = NS["w"]
M = NS["m"]


def qn(tag: str) -> str:
    prefix, local = tag.split(":", 1)
    return f"{{{NS[prefix]}}}{local}"


def wval(value: str) -> str:
    return f"{{{W}}}{value}"


def mval(value: str) -> str:
    return f"{{{M}}}{value}"


def xml_text(element: etree._Element, xpath: str) -> str:
    parts: list[str] = []
    for node in element.xpath(xpath, namespaces=NS):
        if node.text:
            parts.append(node.text)
    return "".join(parts)


def paragraph_text(paragraph: etree._Element) -> str:
    return xml_text(paragraph, ".//w:t")


def paragraph_math_text(paragraph: etree._Element) -> str:
    return xml_text(paragraph, ".//m:t")


def paragraph_style(paragraph: etree._Element) -> str:
    node = paragraph.find("./w:pPr/w:pStyle", namespaces=NS)
    return node.get(wval("val"), "") if node is not None else ""


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


def bookmark_names(element: etree._Element) -> list[str]:
    names: list[str] = []
    for node in element.xpath(".//w:bookmarkStart", namespaces=NS):
        name = node.get(wval("name"))
        if name:
            names.append(name)
    return names


def has_table_ancestor(paragraph: etree._Element) -> bool:
    parent = paragraph.getparent()
    while parent is not None:
        if parent.tag == qn("w:tbl"):
            return True
        parent = parent.getparent()
    return False


def has_blocked_non_formula_content(paragraph: etree._Element) -> bool:
    return bool(paragraph.xpath(".//w:drawing | .//w:pict | .//w:object", namespaces=NS))


EQ_NUMBER_RE = re.compile(r"^[\s\t]*[（(]\s*\d+\s*[-－–—.．]\s*\d+\s*[）)]\s*$")


@dataclass
class LayoutSpec:
    source_paragraph: int
    center_pos: int
    right_pos: int
    jc: str
    raw_tabs: list[dict[str, str]]
    math_wrapper: str
    starts_with_tab: bool
    tab_before_number: bool


@dataclass
class EquationRecord:
    index: int
    style: str
    text: str
    math_text: str
    bookmarks: list[str]
    has_seq: bool
    has_ref_bookmark: bool
    in_table: bool
    wrapper: str
    starts_with_tab: bool
    tab_before_number: bool
    break_before_number: bool
    tabs_before_number: int
    center_tab: int | None
    right_tab: int | None
    jc: str


def load_document_xml(docx_path: Path) -> etree._Element:
    with zipfile.ZipFile(docx_path, "r") as zf:
        return etree.fromstring(zf.read("word/document.xml"))


def ppr(paragraph: etree._Element) -> etree._Element | None:
    return paragraph.find("./w:pPr", namespaces=NS)


def paragraph_jc(paragraph: etree._Element) -> str:
    node = paragraph.find("./w:pPr/w:jc", namespaces=NS)
    return node.get(wval("val"), "") if node is not None else ""


def tab_stops(paragraph: etree._Element) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    for tab in paragraph.xpath("./w:pPr/w:tabs/w:tab", namespaces=NS):
        rows.append(
            {
                "val": tab.get(wval("val"), ""),
                "pos": tab.get(wval("pos"), ""),
                "leader": tab.get(wval("leader"), ""),
            }
        )
    return rows


def tab_position(paragraph: etree._Element, kind: str) -> int | None:
    positions: list[int] = []
    for tab in paragraph.xpath("./w:pPr/w:tabs/w:tab", namespaces=NS):
        if tab.get(wval("val")) != kind:
            continue
        raw = tab.get(wval("pos"))
        if raw and raw.lstrip("-").isdigit():
            positions.append(int(raw))
    return positions[-1] if positions else None


def child_tokens(paragraph: etree._Element) -> list[str]:
    tokens: list[str] = []
    for child in paragraph:
        if child.tag == qn("w:pPr"):
            continue
        if child.tag == qn("m:oMathPara"):
            tokens.append("m:oMathPara")
        elif child.tag == qn("m:oMath"):
            tokens.append("m:oMath")
        elif child.tag == qn("w:r"):
            if child.xpath("./w:tab", namespaces=NS):
                tokens.append("w:tab")
            elif child.xpath("./w:t", namespaces=NS):
                tokens.append("w:t")
            else:
                tokens.append("w:r")
        elif child.tag == qn("w:fldSimple"):
            tokens.append("w:fldSimple")
        elif child.tag == qn("w:bookmarkStart"):
            tokens.append("w:bookmarkStart")
        elif child.tag == qn("w:bookmarkEnd"):
            tokens.append("w:bookmarkEnd")
        else:
            tokens.append(etree.QName(child).localname)
    return tokens


def math_wrapper(paragraph: etree._Element) -> str:
    if paragraph.xpath("./m:oMathPara", namespaces=NS):
        return "oMathPara"
    if paragraph.xpath("./m:oMath", namespaces=NS):
        return "oMath"
    if paragraph.xpath(".//m:oMathPara", namespaces=NS):
        return "nested_oMathPara"
    if paragraph.xpath(".//m:oMath", namespaces=NS):
        return "nested_oMath"
    return ""


def first_content_child(paragraph: etree._Element) -> etree._Element | None:
    for child in paragraph:
        if child.tag != qn("w:pPr"):
            return child
    return None


def starts_with_tab(paragraph: etree._Element) -> bool:
    first = first_content_child(paragraph)
    return first is not None and first.tag == qn("w:r") and bool(first.xpath("./w:tab", namespaces=NS))


def is_tab_run(node: etree._Element) -> bool:
    return node.tag == qn("w:r") and bool(node.xpath("./w:tab", namespaces=NS))


def is_break_run(node: etree._Element) -> bool:
    return node.tag == qn("w:r") and bool(node.xpath("./w:br", namespaces=NS))


def is_empty_spacing_run(node: etree._Element) -> bool:
    return (
        node.tag == qn("w:r")
        and not node.xpath("./w:tab | ./w:br", namespaces=NS)
        and not paragraph_text(node).strip()
    )


def is_math_node(node: etree._Element) -> bool:
    return node.tag in {qn("m:oMathPara"), qn("m:oMath")}


def is_equation_number_prefix_run(node: etree._Element) -> bool:
    if node.tag != qn("w:r"):
        return False
    text = paragraph_text(node)
    return bool(text) and bool(re.fullmatch(r"[\s（(）)\d\-－–—.．]+", text))


def equation_number_block_start(children: list[etree._Element]) -> int | None:
    seq_idx: int | None = None
    for idx, child in enumerate(children):
        if child.tag == qn("w:fldSimple") and "SEQ Equation" in child.get(wval("instr"), ""):
            seq_idx = idx
            break
        if child.tag == qn("w:r") and child.xpath("./w:fldChar[@w:fldCharType='begin']", namespaces=NS):
            instr_parts: list[str] = []
            for later in children[idx + 1 :]:
                if later.tag == qn("w:r"):
                    for instr in later.xpath("./w:instrText", namespaces=NS):
                        if instr.text:
                            instr_parts.append(instr.text)
                    if later.xpath("./w:fldChar[@w:fldCharType='separate' or @w:fldCharType='end']", namespaces=NS):
                        break
                elif later.tag == qn("w:fldSimple"):
                    break
            if "SEQ Equation" in "".join(instr_parts):
                seq_idx = idx
                break
    if seq_idx is None:
        return None
    start_idx = seq_idx
    while start_idx > 0:
        previous = children[start_idx - 1]
        if previous.tag == qn("w:bookmarkStart"):
            start_idx -= 1
            continue
        if is_equation_number_prefix_run(previous):
            start_idx -= 1
            continue
        break
    return start_idx


def tab_before_number(paragraph: etree._Element) -> bool:
    children = [child for child in paragraph if child.tag != qn("w:pPr")]
    start_idx = equation_number_block_start(children)
    if start_idx is not None:
        previous = children[start_idx - 1] if start_idx > 0 else None
        return previous is not None and is_tab_run(previous)
    for idx, child in enumerate(children):
        if child.tag == qn("w:r") and re.search(r"[（(]\s*\d", paragraph_text(child)):
            previous = children[idx - 1] if idx > 0 else None
            return previous is not None and is_tab_run(previous)
    return False


def tabs_before_number_count(paragraph: etree._Element) -> int:
    children = [child for child in paragraph if child.tag != qn("w:pPr")]
    start_idx = equation_number_block_start(children)
    if start_idx is None:
        return 0
    count = 0
    previous_idx = start_idx - 1
    while previous_idx >= 0 and is_tab_run(children[previous_idx]):
        count += 1
        previous_idx -= 1
    return count


def explicit_break_before_number(paragraph: etree._Element) -> bool:
    children = [child for child in paragraph if child.tag != qn("w:pPr")]
    start_idx = equation_number_block_start(children)
    if start_idx is None:
        return False
    previous_idx = start_idx - 1
    while previous_idx >= 0 and is_tab_run(children[previous_idx]):
        previous_idx -= 1
    return previous_idx >= 0 and is_break_run(children[previous_idx])


def is_equation_paragraph(paragraph: etree._Element) -> bool:
    if not paragraph_math_text(paragraph).strip():
        return False
    if has_table_ancestor(paragraph) or has_blocked_non_formula_content(paragraph):
        return False
    text = paragraph_text(paragraph).strip()
    has_seq = any(instr.strip().upper().startswith("SEQ EQUATION") for instr in field_instrs(paragraph))
    has_eq_bookmark = any(name.startswith("eq_") for name in bookmark_names(paragraph))
    if has_seq or has_eq_bookmark:
        return True
    return not text or bool(EQ_NUMBER_RE.match(text))


def equation_records(document: etree._Element) -> list[EquationRecord]:
    rows: list[EquationRecord] = []
    for idx, paragraph in enumerate(document.xpath(".//w:body//w:p", namespaces=NS)):
        if not is_equation_paragraph(paragraph):
            continue
        names = bookmark_names(paragraph)
        rows.append(
            EquationRecord(
                index=idx,
                style=paragraph_style(paragraph),
                text=paragraph_text(paragraph).strip(),
                math_text=paragraph_math_text(paragraph).strip(),
                bookmarks=names,
                has_seq=any(instr.strip().upper().startswith("SEQ EQUATION") for instr in field_instrs(paragraph)),
                has_ref_bookmark=any(name.startswith("eq_") for name in names),
                in_table=has_table_ancestor(paragraph),
                wrapper=math_wrapper(paragraph),
                starts_with_tab=starts_with_tab(paragraph),
                tab_before_number=tab_before_number(paragraph),
                break_before_number=explicit_break_before_number(paragraph),
                tabs_before_number=tabs_before_number_count(paragraph),
                center_tab=tab_position(paragraph, "center"),
                right_tab=tab_position(paragraph, "right"),
                jc=paragraph_jc(paragraph),
            )
        )
    return rows


def candidate_template_rows(template_path: Path) -> list[dict[str, object]]:
    document = load_document_xml(template_path)
    rows: list[dict[str, object]] = []
    for idx, paragraph in enumerate(document.xpath(".//w:body//w:p", namespaces=NS)):
        if not paragraph_math_text(paragraph).strip():
            continue
        rows.append(
            {
                "paragraph": idx,
                "style": paragraph_style(paragraph),
                "text": paragraph_text(paragraph).strip(),
                "math_text": paragraph_math_text(paragraph).strip()[:220],
                "jc": paragraph_jc(paragraph),
                "tabs": tab_stops(paragraph),
                "wrapper": math_wrapper(paragraph),
                "starts_with_tab": starts_with_tab(paragraph),
                "tab_before_number": tab_before_number(paragraph),
                "children": child_tokens(paragraph),
            }
        )
    return rows


def choose_layout_spec(template_path: Path) -> LayoutSpec:
    rows = candidate_template_rows(template_path)
    if not rows:
        raise RuntimeError(f"No math paragraph found in template: {template_path}")
    scored: list[tuple[int, dict[str, object]]] = []
    for row in rows:
        tabs = row["tabs"]
        tab_vals = [item.get("val", "") for item in tabs] if isinstance(tabs, list) else []
        score = 0
        if "center" in tab_vals:
            score += 8
        if "right" in tab_vals:
            score += 8
        if row["starts_with_tab"]:
            score += 4
        if row["tab_before_number"]:
            score += 4
        if row["text"]:
            score += 1
        scored.append((score, row))
    selected = sorted(scored, key=lambda item: item[0], reverse=True)[0][1]
    tabs = selected["tabs"]
    center_pos: int | None = None
    right_pos: int | None = None
    if isinstance(tabs, list):
        for item in tabs:
            if not isinstance(item, dict):
                continue
            raw = item.get("pos", "")
            if not raw.lstrip("-").isdigit():
                continue
            if item.get("val") == "center":
                center_pos = int(raw)
            elif item.get("val") == "right":
                right_pos = int(raw)
    if right_pos is None and center_pos is not None:
        right_pos = center_pos * 2
    if center_pos is None and right_pos is not None:
        center_pos = right_pos // 2
    if center_pos is None or right_pos is None:
        # A4 Tongji-like content width fallback. The audit records this explicitly.
        center_pos = 4156
        right_pos = 8312
    return LayoutSpec(
        source_paragraph=int(selected["paragraph"]),
        center_pos=center_pos,
        right_pos=right_pos,
        jc=str(selected["jc"]),
        raw_tabs=tabs if isinstance(tabs, list) else [],
        math_wrapper=str(selected["wrapper"]),
        starts_with_tab=bool(selected["starts_with_tab"]),
        tab_before_number=bool(selected["tab_before_number"]),
    )


PPR_TAB_FOLLOWERS = {
    "suppressAutoHyphens",
    "kinsoku",
    "wordWrap",
    "overflowPunct",
    "topLinePunct",
    "autoSpaceDE",
    "autoSpaceDN",
    "bidi",
    "adjustRightInd",
    "snapToGrid",
    "spacing",
    "ind",
    "contextualSpacing",
    "mirrorIndents",
    "suppressOverlap",
    "jc",
    "textDirection",
    "textAlignment",
    "textboxTightWrap",
    "outlineLvl",
    "divId",
    "cnfStyle",
    "rPr",
    "sectPr",
    "pPrChange",
}


def ensure_ppr(paragraph: etree._Element) -> etree._Element:
    node = ppr(paragraph)
    if node is not None:
        return node
    node = etree.Element(qn("w:pPr"))
    paragraph.insert(0, node)
    return node


def remove_existing_tabs(ppr_node: etree._Element) -> None:
    for node in list(ppr_node.xpath("./w:tabs", namespaces=NS)):
        ppr_node.remove(node)


def insert_tabs(ppr_node: etree._Element, center_pos: int, right_pos: int) -> None:
    tabs = etree.Element(qn("w:tabs"))
    center = etree.SubElement(tabs, qn("w:tab"))
    center.set(wval("val"), "center")
    center.set(wval("pos"), str(center_pos))
    right = etree.SubElement(tabs, qn("w:tab"))
    right.set(wval("val"), "right")
    right.set(wval("pos"), str(right_pos))
    insert_at = len(ppr_node)
    for idx, child in enumerate(list(ppr_node)):
        if etree.QName(child).localname in PPR_TAB_FOLLOWERS:
            insert_at = idx
            break
    ppr_node.insert(insert_at, tabs)


def set_template_jc(paragraph: etree._Element, jc: str) -> None:
    if not jc:
        ppr_node = ppr(paragraph)
        if ppr_node is not None:
            jc_node = ppr_node.find("./w:jc", namespaces=NS)
            if jc_node is not None:
                ppr_node.remove(jc_node)
        return
    if not jc:
        return
    ppr_node = ensure_ppr(paragraph)
    jc_node = ppr_node.find("./w:jc", namespaces=NS)
    if jc_node is None:
        jc_node = etree.Element(qn("w:jc"))
        insert_at = len(ppr_node)
        for idx, child in enumerate(list(ppr_node)):
            if etree.QName(child).localname in {"textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"}:
                insert_at = idx
                break
        ppr_node.insert(insert_at, jc_node)
    jc_node.set(wval("val"), jc)


def clone_rpr_from_first_run(paragraph: etree._Element) -> etree._Element | None:
    first_run = paragraph.find(".//w:r", namespaces=NS)
    if first_run is None:
        return None
    rpr = first_run.find("./w:rPr", namespaces=NS)
    return etree.fromstring(etree.tostring(rpr)) if rpr is not None else None


def make_tab_run(rpr: etree._Element | None = None) -> etree._Element:
    run = etree.Element(qn("w:r"))
    if rpr is not None:
        run.append(etree.fromstring(etree.tostring(rpr)))
    etree.SubElement(run, qn("w:tab"))
    return run


def make_break_run(rpr: etree._Element | None = None) -> etree._Element:
    run = etree.Element(qn("w:r"))
    if rpr is not None:
        run.append(etree.fromstring(etree.tostring(rpr)))
    etree.SubElement(run, qn("w:br"))
    return run


def remove_tab_runs_before_first_math(paragraph: etree._Element) -> int:
    removed = 0
    for child in list(paragraph):
        if child.tag == qn("w:pPr"):
            continue
        if is_math_node(child):
            break
        if is_tab_run(child):
            paragraph.remove(child)
            removed += 1
            continue
        if child.tag == qn("w:r") and not paragraph_text(child).strip():
            paragraph.remove(child)
            removed += 1
            continue
        break
    return removed


def ensure_leading_center_tab(paragraph: etree._Element) -> bool:
    remove_tab_runs_before_first_math(paragraph)
    children = list(paragraph)
    insert_at: int | None = None
    for idx, child in enumerate(children):
        if child.tag == qn("w:pPr"):
            continue
        insert_at = idx
        break
    if insert_at is None:
        return False
    rpr = clone_rpr_from_first_run(paragraph)
    paragraph.insert(insert_at, make_tab_run(rpr))
    return True


def ensure_single_tab_before_number(paragraph: etree._Element, number_on_own_line: bool = False) -> bool:
    children = [child for child in paragraph if child.tag != qn("w:pPr")]
    start_idx = equation_number_block_start(children)
    if start_idx is None:
        return False
    actual_children = list(paragraph)
    target_node = children[start_idx]
    target_pos = actual_children.index(target_node)
    while target_pos > 0 and (
        is_tab_run(actual_children[target_pos - 1])
        or is_break_run(actual_children[target_pos - 1])
        or is_empty_spacing_run(actual_children[target_pos - 1])
    ):
        paragraph.remove(actual_children[target_pos - 1])
        actual_children = list(paragraph)
        target_pos = actual_children.index(target_node)
    rpr = clone_rpr_from_first_run(paragraph)
    if number_on_own_line:
        paragraph.insert(target_pos, make_break_run(rpr))
        target_pos += 1
        paragraph.insert(target_pos, make_tab_run(rpr))
        target_pos += 1
    paragraph.insert(target_pos, make_tab_run(rpr))
    return True


def convert_omathpara_to_omath(paragraph: etree._Element) -> int:
    converted = 0
    for node in list(paragraph.xpath("./m:oMathPara", namespaces=NS)):
        parent = node.getparent()
        if parent is None:
            continue
        idx = parent.index(node)
        new_nodes: list[etree._Element] = []
        for child in list(node):
            if child.tag == qn("m:oMathParaPr"):
                continue
            if child.tag == qn("m:oMath"):
                new_nodes.append(child)
        if not new_nodes:
            continue
        parent.remove(node)
        for offset, child in enumerate(new_nodes):
            parent.insert(idx + offset, child)
        converted += 1
    return converted


def apply_layout_to_paragraph(
    paragraph: etree._Element,
    spec: LayoutSpec,
    convert_to_inline: bool,
    number_on_own_line: bool = False,
) -> dict[str, object]:
    before = {
        "wrapper": math_wrapper(paragraph),
        "starts_with_tab": starts_with_tab(paragraph),
        "tab_before_number": tab_before_number(paragraph),
        "break_before_number": explicit_break_before_number(paragraph),
        "tabs_before_number": tabs_before_number_count(paragraph),
        "center_tab": tab_position(paragraph, "center"),
        "right_tab": tab_position(paragraph, "right"),
        "jc": paragraph_jc(paragraph),
    }
    ppr_node = ensure_ppr(paragraph)
    remove_existing_tabs(ppr_node)
    insert_tabs(ppr_node, spec.center_pos, spec.right_pos)
    set_template_jc(paragraph, spec.jc)
    converted = convert_omathpara_to_omath(paragraph) if convert_to_inline else 0
    inserted_leading = ensure_leading_center_tab(paragraph)
    inserted_number = ensure_single_tab_before_number(paragraph, number_on_own_line=number_on_own_line)
    after = {
        "wrapper": math_wrapper(paragraph),
        "starts_with_tab": starts_with_tab(paragraph),
        "tab_before_number": tab_before_number(paragraph),
        "break_before_number": explicit_break_before_number(paragraph),
        "tabs_before_number": tabs_before_number_count(paragraph),
        "center_tab": tab_position(paragraph, "center"),
        "right_tab": tab_position(paragraph, "right"),
        "jc": paragraph_jc(paragraph),
    }
    return {
        "before": before,
        "after": after,
        "converted_omathpara": converted,
        "inserted_leading_tab": inserted_leading,
        "inserted_number_tab": inserted_number,
        "number_on_own_line": number_on_own_line,
    }


def repack_docx(tmp_path: Path, output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for item in sorted(tmp_path.rglob("*")):
            if item.is_file():
                zf.write(item, item.relative_to(tmp_path).as_posix())


def copy_unmodified_docx(input_path: Path, output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(input_path, output_path)


def write_csv(path: Path, rows: list[dict[str, object]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if not rows:
        path.write_text("", encoding="utf-8-sig")
        return
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)


def read_bookmark_set(path: Path | None) -> set[str]:
    if path is None:
        return set()
    if not path.exists():
        raise FileNotFoundError(path)
    text = path.read_text(encoding="utf-8-sig").splitlines()
    if not text:
        return set()
    rows: set[str] = set()
    if "," in text[0] and "bookmark" in text[0].lower():
        with path.open("r", newline="", encoding="utf-8-sig") as f:
            for row in csv.DictReader(f):
                value = (row.get("bookmark") or "").strip()
                if value:
                    rows.add(value)
        return rows
    for line in text:
        value = line.strip()
        if not value or value.startswith("#"):
            continue
        rows.add(value.split(",", 1)[0].strip())
    return rows


def audit_docx(
    docx_path: Path,
    out_dir: Path,
    spec: LayoutSpec | None = None,
    number_break_bookmarks: set[str] | None = None,
) -> dict[str, object]:
    document = load_document_xml(docx_path)
    records = equation_records(document)
    break_targets = number_break_bookmarks or set()
    rows: list[dict[str, object]] = []
    for record in records:
        matches = True
        if spec is not None:
            matches = (
                record.starts_with_tab
                and record.tab_before_number
                and record.center_tab == spec.center_pos
                and record.right_tab == spec.right_pos
                and record.jc == spec.jc
            )
        target_break = bool(set(record.bookmarks) & break_targets)
        target_tab_count = 2 if target_break else 1
        matches_target_layout = (
            matches
            and record.break_before_number == target_break
            and record.tabs_before_number == target_tab_count
        )
        rows.append(
            {
                "paragraph": record.index,
                "style": record.style,
                "bookmarks": ";".join(record.bookmarks),
                "has_seq": record.has_seq,
                "has_ref_bookmark": record.has_ref_bookmark,
                "wrapper": record.wrapper,
                "starts_with_tab": record.starts_with_tab,
                "tab_before_number": record.tab_before_number,
                "break_before_number": record.break_before_number,
                "tabs_before_number": record.tabs_before_number,
                "number_break_target": target_break,
                "center_tab": "" if record.center_tab is None else record.center_tab,
                "right_tab": "" if record.right_tab is None else record.right_tab,
                "jc": record.jc,
                "matches_template": matches,
                "matches_target_layout": matches_target_layout,
                "children": " | ".join(child_tokens(document.xpath(".//w:body//w:p", namespaces=NS)[record.index])),
                "text": record.text,
                "math_text": record.math_text[:220],
            }
        )
    out_dir.mkdir(parents=True, exist_ok=True)
    write_csv(out_dir / "equation_alignment_audit.csv", rows)
    summary = {
        "docx": str(docx_path),
        "equation_count": len(records),
        "with_seq": sum(1 for row in rows if row["has_seq"]),
        "with_eq_bookmark": sum(1 for row in rows if row["has_ref_bookmark"]),
        "starts_with_tab": sum(1 for row in rows if row["starts_with_tab"]),
        "tab_before_number": sum(1 for row in rows if row["tab_before_number"]),
        "center_tab_matches": None if spec is None else sum(1 for row in rows if row["center_tab"] == spec.center_pos),
        "right_tab_matches": None if spec is None else sum(1 for row in rows if row["right_tab"] == spec.right_pos),
        "jc_matches": None if spec is None else sum(1 for row in rows if row["jc"] == spec.jc),
        "matches_template": None if spec is None else sum(1 for row in rows if row["matches_template"]),
        "mismatches": None if spec is None else sum(1 for row in rows if not row["matches_template"]),
        "number_break_targets": sum(1 for row in rows if row["number_break_target"]),
        "number_break_actual": sum(1 for row in rows if row["break_before_number"]),
        "two_tabs_before_number": sum(1 for row in rows if row["tabs_before_number"] == 2),
        "matches_target_layout": None if spec is None else sum(1 for row in rows if row["matches_target_layout"]),
        "target_mismatches": None if spec is None else sum(1 for row in rows if not row["matches_target_layout"]),
        "wrappers": {},
    }
    for row in rows:
        wrapper = str(row["wrapper"])
        summary["wrappers"][wrapper] = summary["wrappers"].get(wrapper, 0) + 1
    (out_dir / "equation_alignment_summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    return summary


def apply_layout(
    input_path: Path,
    output_path: Path,
    out_dir: Path,
    spec: LayoutSpec,
    limit: int | None = None,
    convert_to_inline: bool = False,
    number_break_bookmarks: set[str] | None = None,
) -> dict[str, object]:
    out_dir.mkdir(parents=True, exist_ok=True)
    break_targets = number_break_bookmarks or set()
    with tempfile.TemporaryDirectory(prefix="v17_equation_alignment_") as tmp_name:
        tmp_path = Path(tmp_name)
        with zipfile.ZipFile(input_path, "r") as zf:
            zf.extractall(tmp_path)
        document_path = tmp_path / "word" / "document.xml"
        document = etree.fromstring(document_path.read_bytes())
        paragraphs = document.xpath(".//w:body//w:p", namespaces=NS)
        changes: list[dict[str, object]] = []
        touched = 0
        for idx, paragraph in enumerate(paragraphs):
            if not is_equation_paragraph(paragraph):
                continue
            if limit is not None and touched >= limit:
                continue
            names = bookmark_names(paragraph)
            number_on_own_line = bool(set(names) & break_targets)
            detail = apply_layout_to_paragraph(
                paragraph,
                spec,
                convert_to_inline,
                number_on_own_line=number_on_own_line,
            )
            changes.append(
                {
                    "paragraph": idx,
                    "bookmarks": names,
                    "number_on_own_line": number_on_own_line,
                    "text": paragraph_text(paragraph).strip(),
                    "math_text": paragraph_math_text(paragraph).strip()[:220],
                    **detail,
                }
            )
            touched += 1
        document_path.write_bytes(etree.tostring(document, xml_declaration=True, encoding="UTF-8", standalone="yes"))
        repack_docx(tmp_path, output_path)
    report = {
        "input": str(input_path),
        "output": str(output_path),
        "template_source_paragraph": spec.source_paragraph,
        "center_pos": spec.center_pos,
        "right_pos": spec.right_pos,
        "template_jc": spec.jc,
        "limit": limit,
        "convert_to_inline": convert_to_inline,
        "number_break_bookmarks": sorted(break_targets),
        "equations_touched": touched,
        "number_on_own_line": sum(1 for row in changes if row["number_on_own_line"]),
        "changes": changes,
    }
    (out_dir / "alignment_apply_report.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    write_csv(
        out_dir / "alignment_apply_report.csv",
        [
            {
                "paragraph": row["paragraph"],
                "bookmarks": ";".join(str(x) for x in row["bookmarks"]),
                "number_on_own_line": row["number_on_own_line"],
                "converted_omathpara": row["converted_omathpara"],
                "inserted_leading_tab": row["inserted_leading_tab"],
                "inserted_number_tab": row["inserted_number_tab"],
                "before_break_before_number": row["before"]["break_before_number"],
                "after_break_before_number": row["after"]["break_before_number"],
                "before_tabs_before_number": row["before"]["tabs_before_number"],
                "after_tabs_before_number": row["after"]["tabs_before_number"],
                "before_wrapper": row["before"]["wrapper"],
                "after_wrapper": row["after"]["wrapper"],
                "before_center_tab": row["before"]["center_tab"],
                "after_center_tab": row["after"]["center_tab"],
                "before_right_tab": row["before"]["right_tab"],
                "after_right_tab": row["after"]["right_tab"],
                "text": row["text"],
                "math_text": row["math_text"],
            }
            for row in changes
        ],
    )
    return report


def inspect_template(template_path: Path, out_dir: Path) -> dict[str, object]:
    out_dir.mkdir(parents=True, exist_ok=True)
    rows = candidate_template_rows(template_path)
    spec = choose_layout_spec(template_path)
    (out_dir / "template_equation_candidates.json").write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")
    write_csv(
        out_dir / "template_equation_candidates.csv",
        [
            {
                "paragraph": row["paragraph"],
                "style": row["style"],
                "jc": row["jc"],
                "wrapper": row["wrapper"],
                "starts_with_tab": row["starts_with_tab"],
                "tab_before_number": row["tab_before_number"],
                "tabs": json.dumps(row["tabs"], ensure_ascii=False),
                "children": " | ".join(str(x) for x in row["children"]),
                "text": row["text"],
                "math_text": row["math_text"],
            }
            for row in rows
        ],
    )
    result = {
        "template": str(template_path),
        "candidate_count": len(rows),
        "selected": spec.__dict__,
    }
    (out_dir / "template_layout_spec.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    return result


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--mode", choices=["inspect-template", "audit", "apply"], required=True)
    parser.add_argument("--input")
    parser.add_argument("--template", required=True)
    parser.add_argument("--output")
    parser.add_argument("--out-dir", required=True)
    parser.add_argument("--limit", type=int)
    parser.add_argument("--convert-to-inline", action="store_true")
    parser.add_argument("--number-break-bookmarks-file")
    args = parser.parse_args()

    out_dir = Path(args.out_dir)
    number_break_bookmarks = read_bookmark_set(Path(args.number_break_bookmarks_file)) if args.number_break_bookmarks_file else set()
    if args.mode == "inspect-template":
        result = inspect_template(Path(args.template), out_dir)
    elif args.mode == "audit":
        if not args.input:
            raise SystemExit("--input is required for audit mode")
        spec = choose_layout_spec(Path(args.template))
        result = audit_docx(Path(args.input), out_dir, spec, number_break_bookmarks=number_break_bookmarks)
    else:
        if not args.input or not args.output:
            raise SystemExit("--input and --output are required for apply mode")
        spec = choose_layout_spec(Path(args.template))
        result = apply_layout(
            Path(args.input),
            Path(args.output),
            out_dir,
            spec,
            limit=args.limit,
            convert_to_inline=args.convert_to_inline,
            number_break_bookmarks=number_break_bookmarks,
        )
    print(json.dumps(result, ensure_ascii=True, indent=2))


if __name__ == "__main__":
    main()
