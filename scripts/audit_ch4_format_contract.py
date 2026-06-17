from __future__ import annotations

import argparse
import json
import re
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


def read_document_xml(path: Path) -> etree._Element:
    with zipfile.ZipFile(path, "r") as zf:
        return etree.fromstring(zf.read("word/document.xml"))


def visible_text(node: etree._Element) -> str:
    parts: list[str] = []
    for t in node.xpath(".//w:t | .//m:t", namespaces=NS):
        if t.text:
            parts.append(t.text)
    return "".join(parts).strip()


def p_style(p: etree._Element) -> str:
    vals = p.xpath("./w:pPr/w:pStyle/@w:val", namespaces=NS)
    return vals[0] if vals else ""


def jc_value(p: etree._Element) -> str:
    vals = p.xpath("./w:pPr/w:jc/@w:val", namespaces=NS)
    return vals[0] if vals else ""


def first_line(p: etree._Element) -> str:
    vals = p.xpath("./w:pPr/w:ind/@w:firstLine", namespaces=NS)
    return vals[0] if vals else ""


def has_drawing(p: etree._Element) -> bool:
    return bool(p.xpath(".//w:drawing | .//w:pict", namespaces=NS))


def has_equation_field(p: etree._Element) -> bool:
    return bool(p.xpath(".//w:instrText[contains(., 'SEQ Equation')]", namespaces=NS))


def run_size(run: etree._Element) -> tuple[str, str]:
    vals = run.xpath("./w:rPr/w:sz/@w:val", namespaces=NS)
    vals_cs = run.xpath("./w:rPr/w:szCs/@w:val", namespaces=NS)
    return (vals[0] if vals else "", vals_cs[0] if vals_cs else "")


def paragraph_default_size(p: etree._Element) -> tuple[str, str]:
    vals = p.xpath("./w:pPr/w:rPr/w:sz/@w:val", namespaces=NS)
    vals_cs = p.xpath("./w:pPr/w:rPr/w:szCs/@w:val", namespaces=NS)
    return (vals[0] if vals else "", vals_cs[0] if vals_cs else "")


def effective_run_size(run: etree._Element, p: etree._Element) -> tuple[str, str]:
    sz, szcs = run_size(run)
    if sz and szcs:
        return sz, szcs
    psz, pszcs = paragraph_default_size(p)
    return (sz or psz, szcs or pszcs)


def body_blocks(root: etree._Element) -> list[etree._Element]:
    body = root.find(qn("w:body"))
    if body is None:
        raise RuntimeError("word/document.xml has no w:body")
    return [child for child in body if child.tag in {qn("w:p"), qn("w:tbl")}]


def ch4_blocks(root: etree._Element) -> list[tuple[int, etree._Element]]:
    in_ch4 = False
    selected: list[tuple[int, etree._Element]] = []
    for idx, block in enumerate(body_blocks(root)):
        text = visible_text(block) if block.tag == qn("w:p") else ""
        if block.tag == qn("w:p") and re.match(r"^4\.1\s+引言$", text):
            in_ch4 = True
        if in_ch4:
            selected.append((idx, block))
        if in_ch4 and block.tag == qn("w:p") and re.match(r"^5\.1\s+引言$", text):
            selected.pop()
            break
    return selected


def audit_formula_paragraph(idx: int, p: etree._Element, failures: list[dict]) -> dict:
    text = visible_text(p)
    tabs = p.xpath("./w:pPr/w:tabs/w:tab", namespaces=NS)
    tab_pairs = [(tab.get(qn("w:val")), tab.get(qn("w:pos"))) for tab in tabs]
    spacing = p.find("./w:pPr/w:spacing", namespaces=NS)
    line = spacing.get(qn("w:line")) if spacing is not None else ""
    line_rule = spacing.get(qn("w:lineRule")) if spacing is not None else ""
    before = spacing.get(qn("w:before")) if spacing is not None else ""
    after = spacing.get(qn("w:after")) if spacing is not None else ""
    if ("center", "4153") not in tab_pairs or ("right", "8306") not in tab_pairs:
        failures.append({"type": "formula_tabs", "index": idx, "text": text, "tabs": tab_pairs})
    if line != "240" or line_rule != "auto":
        failures.append({"type": "formula_line_spacing", "index": idx, "text": text, "line": line, "lineRule": line_rule})
    if (before or "") not in {"", "0"} or (after or "") not in {"", "0"}:
        failures.append({"type": "formula_before_after", "index": idx, "text": text, "before": before, "after": after})
    if jc_value(p) == "center":
        failures.append({"type": "formula_centered", "index": idx, "text": text})
    bad_runs = []
    for run_idx, run in enumerate(p.xpath(".//w:r | .//m:r", namespaces=NS)):
        sz, szcs = effective_run_size(run, p)
        if sz != "24" or szcs != "24":
            bad_runs.append({"run": run_idx, "sz": sz, "szCs": szcs})
    if bad_runs:
        failures.append({"type": "formula_size", "index": idx, "text": text, "bad_runs": bad_runs[:20]})
    return {"index": idx, "text": text, "line": line, "lineRule": line_rule, "run_count": len(p.xpath(".//w:r | .//m:r", namespaces=NS))}


def audit_body_paragraph(idx: int, p: etree._Element, failures: list[dict]) -> None:
    text = visible_text(p)
    if p_style(p) != "" or not text or has_drawing(p) or has_equation_field(p):
        return
    if jc_value(p) == "center":
        failures.append({"type": "body_centered", "index": idx, "text": text[:120]})
    if first_line(p) != "480":
        failures.append({"type": "body_first_line", "index": idx, "text": text[:120], "firstLine": first_line(p)})
    bad_math = []
    for run_idx, run in enumerate(p.xpath(".//m:r", namespaces=NS)):
        sz, szcs = effective_run_size(run, p)
        if sz != "24" or szcs != "24":
            bad_math.append({"run": run_idx, "sz": sz, "szCs": szcs})
    if bad_math:
        failures.append({"type": "inline_math_size", "index": idx, "text": text[:120], "bad_runs": bad_math[:20]})


def border_val(tbl: etree._Element, name: str) -> str:
    vals = tbl.xpath(f"./w:tblPr/w:tblBorders/w:{name}/@w:val", namespaces=NS)
    return vals[0] if vals else ""


def cell_border_val(tc: etree._Element, name: str) -> str:
    vals = tc.xpath(f"./w:tcPr/w:tcBorders/w:{name}/@w:val", namespaces=NS)
    return vals[0] if vals else ""


def audit_table(idx: int, tbl: etree._Element, failures: list[dict]) -> None:
    top = border_val(tbl, "top")
    bottom = border_val(tbl, "bottom")
    inside_h = border_val(tbl, "insideH")
    vertical = {name: border_val(tbl, name) for name in ["left", "right", "insideV"]}
    if top != "single" or bottom != "single" or inside_h not in {"", "nil", "none"}:
        failures.append({"type": "table_horizontal_borders", "index": idx, "top": top, "bottom": bottom, "insideH": inside_h})
    for name, val in vertical.items():
        if val not in {"", "nil", "none"}:
            failures.append({"type": "table_vertical_border", "index": idx, "border": name, "val": val})
    rows = tbl.xpath("./w:tr", namespaces=NS)
    if rows:
        for cell_idx, tc in enumerate(rows[0].xpath("./w:tc", namespaces=NS)):
            if cell_border_val(tc, "top") != "single" or cell_border_val(tc, "bottom") != "single":
                failures.append({"type": "table_header_rules", "index": idx, "cell": cell_idx, "top": cell_border_val(tc, "top"), "bottom": cell_border_val(tc, "bottom")})
        for cell_idx, tc in enumerate(rows[-1].xpath("./w:tc", namespaces=NS)):
            if cell_border_val(tc, "bottom") != "single":
                failures.append({"type": "table_bottom_rule", "index": idx, "cell": cell_idx, "bottom": cell_border_val(tc, "bottom")})
    for cell_idx, tc in enumerate(tbl.xpath(".//w:tc", namespaces=NS)):
        text = visible_text(tc)
        if re.search(r" {2,}", text):
            failures.append({"type": "table_double_space", "index": idx, "cell": cell_idx, "text": text})
        bad_runs = []
        for run_idx, run in enumerate(tc.xpath(".//w:r | .//m:r", namespaces=NS)):
            if not visible_text(run):
                continue
            sz, szcs = run_size(run)
            if sz != "18" or szcs != "18":
                bad_runs.append({"run": run_idx, "sz": sz, "szCs": szcs})
        if bad_runs:
            failures.append({"type": "table_font_size", "index": idx, "cell": cell_idx, "text": text[:80], "bad_runs": bad_runs[:20]})


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("docx")
    parser.add_argument("--out", required=True)
    args = parser.parse_args()

    root = read_document_xml(Path(args.docx))
    failures: list[dict] = []
    formulas = []
    table_count = 0
    body_paragraph_count = 0
    for idx, block in ch4_blocks(root):
        if block.tag == qn("w:tbl"):
            table_count += 1
            audit_table(idx, block, failures)
            continue
        is_formula = has_equation_field(block) or (bool(re.search(r"（4-\d+）", visible_text(block))) and bool(block.xpath(".//m:oMath", namespaces=NS)))
        if is_formula:
            formulas.append(audit_formula_paragraph(idx, block, failures))
        elif p_style(block) == "" and visible_text(block) and not has_drawing(block):
            body_paragraph_count += 1
            audit_body_paragraph(idx, block, failures)

    report = {
        "docx": args.docx,
        "formula_count": len(formulas),
        "body_paragraph_count": body_paragraph_count,
        "table_count": table_count,
        "formulas": formulas,
        "failures": failures,
    }
    Path(args.out).write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({k: report[k] for k in ["formula_count", "body_paragraph_count", "table_count", "failures"]}, ensure_ascii=True, indent=2))
    if failures:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
