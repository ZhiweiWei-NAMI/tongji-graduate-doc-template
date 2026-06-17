from __future__ import annotations

import argparse
import json
import re
import zipfile
from dataclasses import dataclass
from pathlib import Path

from lxml import etree


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def qn(tag: str) -> str:
    prefix, local = tag.split(":")
    return f"{{{NS[prefix]}}}{local}"


def text_of(node: etree._Element) -> str:
    texts: list[str] = []
    for t in node.xpath(".//w:t | .//w:instrText | .//m:t", namespaces=NS):
        if t.text:
            texts.append(t.text)
    return "".join(texts).strip()


def visible_text_of(node: etree._Element) -> str:
    texts: list[str] = []
    for t in node.xpath(".//w:t | .//m:t", namespaces=NS):
        if t.text:
            texts.append(t.text)
    return "".join(texts).strip()


def blocks(root: etree._Element) -> list[etree._Element]:
    body = root.find(qn("w:body"))
    if body is None:
        return []
    return [child for child in body if child.tag in {qn("w:p"), qn("w:tbl")}]


def style_id(p: etree._Element) -> str:
    vals = p.xpath("./w:pPr/w:pStyle/@w:val", namespaces=NS)
    return vals[0] if vals else ""


def field_codes(node: etree._Element) -> list[str]:
    return [t.text or "" for t in node.xpath(".//w:instrText", namespaces=NS)]


def p_tabs(p: etree._Element) -> list[dict]:
    tabs = []
    for tab in p.xpath("./w:pPr/w:tabs/w:tab", namespaces=NS):
        tabs.append(
            {
                "val": tab.get(qn("w:val"), ""),
                "pos": tab.get(qn("w:pos"), ""),
                "leader": tab.get(qn("w:leader"), ""),
            }
        )
    return tabs


def paragraph_summary(p: etree._Element, index: int) -> dict:
    text = visible_text_of(p)
    return {
        "index": index,
        "kind": "p",
        "style": style_id(p),
        "text": text[:260],
        "has_drawing": bool(p.xpath(".//w:drawing | .//w:pict", namespaces=NS)),
        "math_count": len(p.xpath(".//m:oMath | .//m:oMathPara", namespaces=NS)),
        "field_codes": field_codes(p),
        "tabs": p_tabs(p),
        "jc": (p.xpath("./w:pPr/w:jc/@w:val", namespaces=NS) or [""])[0],
        "page_break_before": bool(p.xpath("./w:pPr/w:pageBreakBefore", namespaces=NS)),
        "page_break_runs": len(p.xpath(".//w:br[@w:type='page']", namespaces=NS)),
    }


def table_borders(tbl: etree._Element) -> dict:
    out = {}
    for side in ["top", "bottom", "left", "right", "insideH", "insideV"]:
        node = tbl.find(f"./w:tblPr/w:tblBorders/w:{side}", namespaces=NS)
        if node is None:
            out[side] = None
        else:
            out[side] = {
                "val": node.get(qn("w:val"), ""),
                "sz": node.get(qn("w:sz"), ""),
                "color": node.get(qn("w:color"), ""),
            }
    return out


def border_dict(node: etree._Element | None, side: str) -> dict | None:
    if node is None:
        return None
    border = node.find(f"./w:{side}", namespaces=NS)
    if border is None:
        return None
    return {
        "val": border.get(qn("w:val"), ""),
        "sz": border.get(qn("w:sz"), ""),
        "color": border.get(qn("w:color"), ""),
    }


def cell_border_summary(tbl: etree._Element) -> list[dict]:
    result = []
    for r_idx, row in enumerate(tbl.xpath("./w:tr", namespaces=NS)):
        row_result = {"row": r_idx, "cells": []}
        for c_idx, cell in enumerate(row.xpath("./w:tc", namespaces=NS)):
            borders = cell.find("./w:tcPr/w:tcBorders", namespaces=NS)
            row_result["cells"].append(
                {
                    "col": c_idx,
                    "top": border_dict(borders, "top"),
                    "bottom": border_dict(borders, "bottom"),
                    "left": border_dict(borders, "left"),
                    "right": border_dict(borders, "right"),
                }
            )
        result.append(row_result)
    return result


def row_flags(row: etree._Element, index: int) -> dict:
    return {
        "index": index,
        "header": bool(row.xpath("./w:trPr/w:tblHeader", namespaces=NS)),
        "cant_split": bool(row.xpath("./w:trPr/w:cantSplit", namespaces=NS)),
    }


def cell_texts(tbl: etree._Element) -> list[list[str]]:
    result = []
    for row in tbl.xpath("./w:tr", namespaces=NS):
        result.append([visible_text_of(cell) for cell in row.xpath("./w:tc", namespaces=NS)])
    return result


def cell_font_sizes(tbl: etree._Element) -> list[str]:
    sizes = []
    for sz in tbl.xpath(".//w:rPr/w:sz/@w:val", namespaces=NS):
        if sz:
            sizes.append(sz)
    return sorted(set(sizes), key=lambda x: int(x) if x.isdigit() else 9999)


def table_summary(tbl: etree._Element, index: int) -> dict:
    rows = tbl.xpath("./w:tr", namespaces=NS)
    cols = max((len(row.xpath("./w:tc", namespaces=NS)) for row in rows), default=0)
    texts = cell_texts(tbl)
    double_space_cells = []
    for r_idx, row in enumerate(texts):
        for c_idx, cell_text in enumerate(row):
            if "  " in cell_text:
                double_space_cells.append({"row": r_idx, "col": c_idx, "text": cell_text[:120]})
    return {
        "index": index,
        "kind": "tbl",
        "rows": len(rows),
        "cols": cols,
        "text": visible_text_of(tbl)[:300],
        "borders": table_borders(tbl),
        "cell_borders": cell_border_summary(tbl),
        "row_flags": [row_flags(row, idx) for idx, row in enumerate(rows)],
        "font_sizes": cell_font_sizes(tbl),
        "double_space_cells": double_space_cells,
    }


def all_summaries(root: etree._Element) -> list[dict]:
    result = []
    for idx, block in enumerate(blocks(root)):
        if block.tag == qn("w:p"):
            result.append(paragraph_summary(block, idx))
        elif block.tag == qn("w:tbl"):
            result.append(table_summary(block, idx))
    return result


def find_heading_indices(items: list[dict], start: str, end: str) -> tuple[int, int]:
    start_idx = None
    for item in items:
        if item["kind"] == "p" and item["text"].startswith(start):
            start_idx = item["index"]
            break
    if start_idx is None:
        raise RuntimeError(f"Cannot find heading {start!r}")
    end_idx = len(items)
    for item in items:
        if item["index"] > start_idx and item["kind"] == "p" and item["text"].startswith(end):
            end_idx = item["index"]
            break
    return start_idx, end_idx


def formula_items(items: list[dict], start: int, end: int) -> list[dict]:
    rx = re.compile(r"（\d+-\d+）")
    return [
        item
        for item in items
        if start <= item["index"] < end
        and item["kind"] == "p"
        and (rx.search(item["text"]) or item["math_count"] > 0)
    ]


def caption_items(items: list[dict], start: int, end: int) -> list[dict]:
    return [
        item
        for item in items
        if start <= item["index"] < end
        and item["kind"] == "p"
        and re.match(r"^(图|表)\s*\d+-\d+", item["text"])
    ]


def reference_mentions(items: list[dict], start: int, end: int) -> list[dict]:
    out = []
    rx = re.compile(r"(图|表)\s*\d+-\d+")
    for item in items:
        if not (start <= item["index"] < end and item["kind"] == "p"):
            continue
        text = item["text"]
        if not rx.search(text):
            continue
        is_caption = bool(re.match(r"^(图|表)\s*\d+-\d+", text))
        if is_caption:
            continue
        codes = " ".join(item["field_codes"])
        out.append({"index": item["index"], "text": text[:220], "has_ref_field": "REF" in codes, "field_codes": item["field_codes"]})
    return out


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("docx")
    parser.add_argument("--start", default="4.1")
    parser.add_argument("--end", default="5.1")
    parser.add_argument("--out", required=True)
    args = parser.parse_args()

    with zipfile.ZipFile(Path(args.docx)) as zf:
        root = etree.fromstring(zf.read("word/document.xml"))
    items = all_summaries(root)
    start, end = find_heading_indices(items, args.start, args.end)
    section = [item for item in items if start <= item["index"] < end]
    result = {
        "docx": str(Path(args.docx)),
        "range": {"start": start, "end": end, "start_heading": args.start, "end_heading": args.end},
        "headings": [
            item for item in section
            if item["kind"] == "p" and re.match(r"^\d+(?:\.\d+)+\s+", item["text"])
        ],
        "formulas": formula_items(items, start, end),
        "captions": caption_items(items, start, end),
        "reference_mentions": reference_mentions(items, start, end),
        "tables": [item for item in section if item["kind"] == "tbl"],
        "blocks": section,
    }
    Path(args.out).write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    main()
