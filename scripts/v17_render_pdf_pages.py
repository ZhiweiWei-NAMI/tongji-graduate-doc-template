from __future__ import annotations

import argparse
import json
from pathlib import Path

import fitz


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--pdf", required=True)
    parser.add_argument("--out-dir", required=True)
    parser.add_argument("--zoom", type=float, default=2.0)
    args = parser.parse_args()

    pdf_path = Path(args.pdf)
    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    for stale in out_dir.glob("page-*.png"):
        stale.unlink()
    doc = fitz.open(pdf_path)
    paths: list[str] = []
    for page_index in range(doc.page_count):
        pix = doc[page_index].get_pixmap(matrix=fitz.Matrix(args.zoom, args.zoom), alpha=False)
        out_path = out_dir / f"page-{page_index + 1:03d}.png"
        pix.save(out_path)
        paths.append(str(out_path))
    result = {"pdf": str(pdf_path), "page_count": doc.page_count, "zoom": args.zoom, "pages": paths}
    print(json.dumps(result, ensure_ascii=True, indent=2))


if __name__ == "__main__":
    main()
