from __future__ import annotations

import argparse
import json
import re
import shutil
from pathlib import Path

import fitz


def clean_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)
    for item in path.glob("page_*.png"):
        item.unlink()


def render_pages(pdf: Path, out_dir: Path, dpi: int) -> dict:
    clean_dir(out_dir)
    doc = fitz.open(pdf)
    zoom = dpi / 72
    matrix = fitz.Matrix(zoom, zoom)
    page_texts = []
    for i, page in enumerate(doc):
        pix = page.get_pixmap(matrix=matrix, alpha=False)
        out = out_dir / f"page_{i + 1:03d}.png"
        pix.save(out)
        text = page.get_text("text")
        page_texts.append({"page": i + 1, "text": text})
    return {"page_count": len(doc), "pages": page_texts}


def detect_pages(page_texts: list[dict]) -> dict:
    formula_pages = []
    table_pages = []
    figure_pages = []
    heading_pages = []
    for item in page_texts:
        page = item["page"]
        text = item["text"]
        if re.search(r"^\s*\d+(?:\.\d+)+\s+", text, flags=re.M):
            heading_pages.append(page)
        if re.search(r"（\d+-\d+）", text):
            formula_pages.append(page)
        if re.search(r"表\s*\d+-\d+", text):
            table_pages.append(page)
        if re.search(r"图\s*\d+-\d+", text):
            figure_pages.append(page)
    return {
        "heading_pages": sorted(set(heading_pages)),
        "formula_pages": sorted(set(formula_pages)),
        "table_pages": sorted(set(table_pages)),
        "figure_pages": sorted(set(figure_pages)),
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--pdf", required=True)
    parser.add_argument("--out-dir", required=True)
    parser.add_argument("--dpi", type=int, default=160)
    parser.add_argument("--report", required=True)
    args = parser.parse_args()

    pdf = Path(args.pdf)
    out_dir = Path(args.out_dir)
    report_path = Path(args.report)
    result = render_pages(pdf, out_dir, args.dpi)
    detected = detect_pages(result["pages"])
    report = {"pdf": str(pdf), "out_dir": str(out_dir), "dpi": args.dpi, **result, **detected}
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({k: v for k, v in report.items() if k != "pages"}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
