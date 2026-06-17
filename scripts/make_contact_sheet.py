from __future__ import annotations

import argparse
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--png-dir", required=True)
    parser.add_argument("--pages", required=True)
    parser.add_argument("--out", required=True)
    parser.add_argument("--cols", type=int, default=2)
    parser.add_argument("--scale", type=float, default=0.65)
    args = parser.parse_args()

    png_dir = Path(args.png_dir)
    pages = [int(x.strip()) for x in args.pages.split(",") if x.strip()]
    images = []
    for page in pages:
        path = png_dir / f"page_{page:03d}.png"
        img = Image.open(path).convert("RGB")
        if args.scale != 1:
            img = img.resize((int(img.width * args.scale), int(img.height * args.scale)), Image.LANCZOS)
        images.append((page, img))
    if not images:
        raise SystemExit("No images")
    label_h = 44
    pad = 18
    cell_w = max(img.width for _, img in images)
    cell_h = max(img.height for _, img in images) + label_h
    cols = max(1, args.cols)
    rows = (len(images) + cols - 1) // cols
    sheet = Image.new("RGB", (cols * cell_w + (cols + 1) * pad, rows * cell_h + (rows + 1) * pad), "white")
    draw = ImageDraw.Draw(sheet)
    try:
        font = ImageFont.truetype("arial.ttf", 26)
    except Exception:
        font = ImageFont.load_default()
    for idx, (page, img) in enumerate(images):
        row = idx // cols
        col = idx % cols
        x = pad + col * (cell_w + pad)
        y = pad + row * (cell_h + pad)
        draw.text((x, y), f"PDF page {page}", fill=(0, 0, 0), font=font)
        sheet.paste(img, (x, y + label_h))
    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    sheet.save(out)
    print(out)


if __name__ == "__main__":
    main()
