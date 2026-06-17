from __future__ import annotations

import json
import sys
from pathlib import Path

from PIL import Image


def page_metrics(path: Path) -> dict:
    img = Image.open(path).convert("L")
    w, h = img.size
    pix = img.load()
    rows = []
    for y in range(h):
        dark = 0
        for x in range(w):
            if pix[x, y] < 245:
                dark += 1
        rows.append(dark)
    threshold = max(8, int(w * 0.002))
    occupied = [i for i, v in enumerate(rows) if v >= threshold]
    if not occupied:
        return {"page": path.stem, "blank": True, "bottom_blank_ratio": 1.0, "large_gaps": []}
    top = occupied[0]
    bottom = occupied[-1]
    bottom_blank_ratio = (h - bottom - 1) / h
    gaps = []
    start = None
    for y, v in enumerate(rows):
        if top <= y <= bottom and v < threshold:
            if start is None:
                start = y
        elif start is not None:
            if y - start > h * 0.08:
                gaps.append({"start_px": start, "end_px": y - 1, "ratio": round((y - start) / h, 4)})
            start = None
    if start is not None and bottom - start > h * 0.08:
        gaps.append({"start_px": start, "end_px": bottom, "ratio": round((bottom - start + 1) / h, 4)})
    return {
        "page": path.stem,
        "blank": False,
        "content_top_ratio": round(top / h, 4),
        "content_bottom_ratio": round(bottom / h, 4),
        "bottom_blank_ratio": round(bottom_blank_ratio, 4),
        "large_gaps": gaps,
    }


def main() -> int:
    if len(sys.argv) != 4:
        print("usage: audit_ch5_page_whitespace.py pages_dir start_page end_page", file=sys.stderr)
        return 2
    pages_dir = Path(sys.argv[1])
    start = int(sys.argv[2])
    end = int(sys.argv[3])
    report = []
    for page in range(start, end + 1):
        path = pages_dir / f"page_{page:03d}.png"
        if not path.exists():
            raise FileNotFoundError(path)
        report.append(page_metrics(path))
    flagged = [
        item
        for item in report
        if item.get("blank") or item.get("bottom_blank_ratio", 0) > 0.32 or item.get("large_gaps")
    ]
    result = {"pages_dir": str(pages_dir), "start_page": start, "end_page": end, "flagged": flagged, "pages": report}
    out = pages_dir.parent / "r17_ch5_whitespace_audit.json"
    out.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"flagged_count": len(flagged), "flagged": flagged}, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
