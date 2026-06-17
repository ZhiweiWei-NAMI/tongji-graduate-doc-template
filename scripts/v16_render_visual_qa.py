from __future__ import annotations

import argparse
import csv
import json
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


def image_metrics(path: Path) -> dict[str, object]:
    img = Image.open(path).convert("L")
    width, height = img.size
    pixels = img.load()
    min_x, min_y = width, height
    max_x, max_y = -1, -1
    ink = 0
    for y in range(height):
        for x in range(width):
            if pixels[x, y] < 245:
                ink += 1
                if x < min_x:
                    min_x = x
                if x > max_x:
                    max_x = x
                if y < min_y:
                    min_y = y
                if y > max_y:
                    max_y = y
    if max_x < 0:
        bbox = None
        margins = None
    else:
        bbox = [min_x, min_y, max_x, max_y]
        margins = {
            "left": min_x,
            "top": min_y,
            "right": width - max_x - 1,
            "bottom": height - max_y - 1,
        }
    return {
        "file": path.name,
        "page": int(path.stem.split("-")[-1]),
        "width": width,
        "height": height,
        "ink_ratio": ink / (width * height),
        "bbox": "" if bbox is None else json.dumps(bbox),
        "left_margin_px": "" if margins is None else margins["left"],
        "top_margin_px": "" if margins is None else margins["top"],
        "right_margin_px": "" if margins is None else margins["right"],
        "bottom_margin_px": "" if margins is None else margins["bottom"],
    }


def write_csv(path: Path, rows: list[dict[str, object]]) -> None:
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)


def make_contact_sheets(paths: list[Path], out_dir: Path, *, cols: int = 4, rows: int = 3) -> list[Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    out_paths: list[Path] = []
    thumb_w = 340
    label_h = 28
    pad = 14
    font = ImageFont.load_default()
    per_sheet = cols * rows
    for offset in range(0, len(paths), per_sheet):
        chunk = paths[offset : offset + per_sheet]
        thumbs: list[tuple[Path, Image.Image]] = []
        max_h = 0
        for path in chunk:
            img = Image.open(path).convert("RGB")
            ratio = thumb_w / img.width
            thumb_h = int(img.height * ratio)
            img = img.resize((thumb_w, thumb_h), Image.Resampling.LANCZOS)
            max_h = max(max_h, thumb_h)
            thumbs.append((path, img))
        sheet_w = cols * thumb_w + (cols + 1) * pad
        sheet_h = rows * (max_h + label_h) + (rows + 1) * pad
        sheet = Image.new("RGB", (sheet_w, sheet_h), "white")
        draw = ImageDraw.Draw(sheet)
        for idx, (path, img) in enumerate(thumbs):
            col = idx % cols
            row = idx // cols
            x = pad + col * (thumb_w + pad)
            y = pad + row * (max_h + label_h + pad)
            draw.text((x, y), path.stem, fill=(0, 0, 0), font=font)
            sheet.paste(img, (x, y + label_h))
        first = int(chunk[0].stem.split("-")[-1])
        last = int(chunk[-1].stem.split("-")[-1])
        out_path = out_dir / f"contact_sheet_{first:03d}_{last:03d}.png"
        sheet.save(out_path)
        out_paths.append(out_path)
    return out_paths


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--render-dir", required=True)
    parser.add_argument("--out-dir", required=True)
    args = parser.parse_args()

    render_dir = Path(args.render_dir)
    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    paths = sorted(render_dir.glob("page-*.png"))
    metrics = [image_metrics(path) for path in paths]
    write_csv(out_dir / "page_image_metrics.csv", metrics)
    low_ink = [row for row in metrics if row["ink_ratio"] < 0.015]
    narrow = [
        row
        for row in metrics
        if row["left_margin_px"] != "" and (int(row["left_margin_px"]) < 40 or int(row["right_margin_px"]) < 40)
    ]
    sheets = make_contact_sheets(paths, out_dir / "contact_sheets")
    summary = {
        "page_count": len(paths),
        "low_ink_pages": [row["page"] for row in low_ink],
        "narrow_margin_pages": [row["page"] for row in narrow],
        "contact_sheets": [str(path) for path in sheets],
    }
    (out_dir / "visual_qa_summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(summary, ensure_ascii=True, indent=2))


if __name__ == "__main__":
    main()
