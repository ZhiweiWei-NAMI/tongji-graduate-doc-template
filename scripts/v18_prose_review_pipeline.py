from __future__ import annotations

import argparse
import json
import re
import zipfile
from dataclasses import dataclass
from pathlib import Path

import fitz
from lxml import etree


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}


def qn(prefix: str, name: str) -> str:
    return f"{{{NS[prefix]}}}{name}"


def normalize_text(text: str) -> str:
    text = text.replace("\u3000", "")
    text = re.sub(r"\s+", "", text)
    return text


def block_text(element: etree._Element) -> str:
    return "".join(t.text or "" for t in element.xpath(".//w:t", namespaces=NS))


def paragraph_style(paragraph: etree._Element) -> str:
    style = paragraph.find("./w:pPr/w:pStyle", namespaces=NS)
    if style is None:
        return ""
    return style.get(qn("w", "val"), "")


@dataclass
class Unit:
    id: int
    kind: str
    style: str
    heading1: str
    heading2: str
    heading3: str
    page: int
    text: str


def load_pdf_pages(pdf_path: Path) -> list[str]:
    doc = fitz.open(pdf_path)
    pages: list[str] = []
    for page in doc:
        pages.append(normalize_text(page.get_text("text")))
    doc.close()
    return pages


def locate_page(text: str, pages: list[str], fallback_page: int) -> int:
    normalized = normalize_text(text)
    if not normalized:
        return fallback_page
    if len(normalized) <= 12:
        return fallback_page

    candidates: list[str] = []
    if len(normalized) <= 80:
        candidates.append(normalized)
    else:
        candidates.extend(
            [
                normalized[:70],
                normalized[max(0, len(normalized) // 2 - 35) : len(normalized) // 2 + 35],
                normalized[-70:],
            ]
        )
    for needle in candidates:
        if len(needle) < 12:
            continue
        for index, page_text in enumerate(pages, start=1):
            if needle in page_text:
                return index

    # Use shorter prose-rich windows for paragraphs containing formulas or unusual glyphs.
    windows = [normalized[i : i + 36] for i in range(0, max(0, len(normalized) - 35), 18)]
    for needle in windows[:20]:
        if len(needle) < 12:
            continue
        for index, page_text in enumerate(pages, start=1):
            if needle in page_text:
                return index
    return fallback_page


def extract_units(docx_path: Path, pdf_path: Path) -> list[Unit]:
    pages = load_pdf_pages(pdf_path)
    with zipfile.ZipFile(docx_path) as zf:
        document = etree.fromstring(zf.read("word/document.xml"))
    body = document.find("./w:body", namespaces=NS)
    if body is None:
        raise RuntimeError("DOCX body not found.")

    units: list[Unit] = []
    heading1 = ""
    heading2 = ""
    heading3 = ""
    last_page = 1
    uid = 0

    for child in body:
        tag = etree.QName(child).localname
        if tag == "p":
            text = block_text(child).strip()
            if not text:
                continue
            style = paragraph_style(child)
            if style == "1":
                heading1, heading2, heading3 = text, "", ""
            elif style == "af":
                heading2, heading3 = text, ""
            elif style == "af1":
                heading3 = text
            page = locate_page(text, pages, last_page)
            if page:
                last_page = page
            uid += 1
            units.append(Unit(uid, "paragraph", style, heading1, heading2, heading3, page, text))
        elif tag == "tbl":
            text = re.sub(r"\s+", " ", block_text(child)).strip()
            if not text:
                continue
            page = locate_page(text, pages, last_page)
            if page:
                last_page = page
            uid += 1
            units.append(Unit(uid, "table", "table", heading1, heading2, heading3, page, text))
    return units


def unit_to_dict(unit: Unit) -> dict[str, object]:
    return {
        "id": unit.id,
        "kind": unit.kind,
        "style": unit.style,
        "page": unit.page,
        "heading1": unit.heading1,
        "heading2": unit.heading2,
        "heading3": unit.heading3,
        "text": unit.text,
    }


def write_extract(args: argparse.Namespace) -> int:
    docx_path = Path(args.docx)
    pdf_path = Path(args.pdf)
    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    units = extract_units(docx_path, pdf_path)

    jsonl_path = out_dir / "v18_paragraph_units.jsonl"
    with jsonl_path.open("w", encoding="utf-8") as fh:
        for unit in units:
            fh.write(json.dumps(unit_to_dict(unit), ensure_ascii=False) + "\n")

    md_path = out_dir / "v18_paragraph_units.md"
    with md_path.open("w", encoding="utf-8") as fh:
        fh.write(f"# Page-Batched Prose Review Draft\n\nSource DOCX: {docx_path}\n\n")
        for unit in units:
            path = " / ".join(part for part in [unit.heading1, unit.heading2, unit.heading3] if part)
            fh.write(f"## U{unit.id:04d} | p.{unit.page} | {unit.kind} | style={unit.style}\n")
            fh.write(f"Path: {path}\n\n")
            fh.write(unit.text + "\n\n")

    chunk_specs = [
        ("front_ch1_ch2", lambda u: u.page <= 37),
        ("ch3", lambda u: 38 <= u.page <= 66),
        ("ch4", lambda u: 67 <= u.page <= 102),
        ("ch5_ch6", lambda u: 103 <= u.page <= 146),
        ("ch7_back", lambda u: u.page >= 147),
    ]
    chunks: list[dict[str, object]] = []
    for name, predicate in chunk_specs:
        selected = [u for u in units if predicate(u)]
        path = out_dir / f"chunk_{name}.md"
        with path.open("w", encoding="utf-8") as fh:
            fh.write(f"# Chunk {name}\n\n")
            fh.write("审校目标：识别正式博士论文中不应出现的修订痕迹、版本说明、过程日志、防御性表述、口语化解释、过度自辩、与证据不匹配的承诺式表达。\n\n")
            for unit in selected:
                unit_path = " / ".join(part for part in [unit.heading1, unit.heading2, unit.heading3] if part)
                fh.write(f"## U{unit.id:04d} | p.{unit.page} | {unit.kind} | style={unit.style}\n")
                fh.write(f"Path: {unit_path}\n\n")
                fh.write(unit.text + "\n\n")
        chunks.append({"name": name, "path": str(path), "units": len(selected)})

    summary = {
        "docx": str(docx_path),
        "pdf": str(pdf_path),
        "unit_count": len(units),
        "jsonl": str(jsonl_path),
        "markdown": str(md_path),
        "chunks": chunks,
    }
    (out_dir / "v18_paragraph_extract_summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(summary, ensure_ascii=False))
    return 0


def main() -> int:
    parser = argparse.ArgumentParser()
    sub = parser.add_subparsers(dest="cmd", required=True)
    p_extract = sub.add_parser("extract")
    p_extract.add_argument("--docx", required=True)
    p_extract.add_argument("--pdf", required=True)
    p_extract.add_argument("--out-dir", required=True)
    p_extract.set_defaults(func=write_extract)
    args = parser.parse_args()
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())
