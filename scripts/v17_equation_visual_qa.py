from __future__ import annotations

import argparse
import csv
import json
import math
import re
import zipfile
from pathlib import Path

import fitz
from lxml import etree
from PIL import Image, ImageDraw, ImageFont


NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = NS["w"]


def wval(value: str) -> str:
    return f"{{{W}}}{value}"


def first_left_margin_pt(docx_path: Path) -> float:
    with zipfile.ZipFile(docx_path, "r") as zf:
        root = etree.fromstring(zf.read("word/document.xml"))
    pg_mar = root.find(".//w:sectPr/w:pgMar", namespaces=NS)
    if pg_mar is None:
        return 90.0
    raw = pg_mar.get(wval("left"))
    if raw is None:
        return 90.0
    try:
        return int(raw) / 20.0
    except ValueError:
        return 90.0


def is_true(value: str | None) -> bool:
    return str(value or "").strip().lower() in {"true", "1", "yes"}


def read_equations(audit_csv: Path, limit: int | None, only_matches: bool) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    with audit_csv.open("r", newline="", encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            if only_matches:
                target_value = row.get("matches_target_layout")
                template_value = row.get("matches_template")
                if target_value is not None:
                    if not is_true(target_value):
                        continue
                elif not is_true(template_value):
                    continue
            if not row.get("text", "").strip():
                continue
            rows.append(row)
            if limit is not None and len(rows) >= limit:
                break
    return rows


def read_word_page_hints(path: Path | None) -> dict[str, int]:
    if path is None:
        return {}
    if not path.exists():
        raise FileNotFoundError(path)
    hints: dict[str, int] = {}
    with path.open("r", newline="", encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            label = (row.get("label") or "").strip()
            raw_page = (row.get("page") or "").strip()
            if not label or not raw_page.isdigit():
                continue
            hints[label] = int(raw_page) - 1
    return hints


def label_variants(label: str, expanded: bool = False) -> list[str]:
    compact = re.sub(r"\s+", "", label)
    variants = [compact]
    variants.append(compact.replace("（", "(").replace("）", ")"))
    variants.append(compact.replace("(", "（").replace(")", "）"))
    if expanded:
        match = re.search(r"(\d+)\s*[-－–—.．]\s*(\d+)", compact)
        if match:
            chapter, number = match.group(1), match.group(2)
            for dash in ["-", "－", "–", "—", "．", "."]:
                variants.append(f"{chapter}{dash}{number}")
                variants.append(f"{chapter} {dash} {number}")
                variants.append(f"({chapter}{dash}{number})")
                variants.append(f"( {chapter}{dash}{number} )")
                variants.append(f"（{chapter}{dash}{number}）")
                variants.append(f"（ {chapter}{dash}{number} ）")
    # Some PDF extractors split or normalize dashes.
    variants.extend(
        [
            item.replace("-", "－")
            for item in list(variants)
        ]
    )
    seen: set[str] = set()
    result: list[str] = []
    for item in variants:
        if item and item not in seen:
            seen.add(item)
            result.append(item)
    return result


def find_number_rect(
    pdf: fitz.Document,
    label: str,
    expected_right_pt: float,
    page_hint: int | None = None,
) -> tuple[int, fitz.Rect] | None:
    if page_hint is not None and 0 <= page_hint < pdf.page_count:
        page_indexes = [page_hint]
    else:
        page_indexes = list(range(pdf.page_count))
    for expanded in [False, True]:
        candidates: list[tuple[float, int, fitz.Rect]] = []
        for page_index in page_indexes:
            page = pdf[page_index]
            for variant in label_variants(label, expanded=expanded):
                for rect in page.search_for(variant):
                    if rect.x0 < page.rect.width * 0.55:
                        continue
                    score = abs(rect.x1 - expected_right_pt)
                    candidates.append((score, page_index, rect))
        if candidates:
            _, page_index, rect = sorted(candidates, key=lambda item: item[0])[0]
            return page_index, rect
    return None


def render_page(pdf: fitz.Document, page_index: int, zoom: float) -> Image.Image:
    pix = pdf[page_index].get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
    return Image.frombytes("RGB", (pix.width, pix.height), pix.samples)


def dark_bbox(
    image: Image.Image,
    x0: int,
    y0: int,
    x1: int,
    y1: int,
    threshold: int = 210,
    ignore_long_rows: bool = True,
) -> tuple[int, int, int, int] | None:
    gray = image.convert("L")
    width, height = gray.size
    x0 = max(0, min(width - 1, x0))
    x1 = max(0, min(width, x1))
    y0 = max(0, min(height - 1, y0))
    y1 = max(0, min(height, y1))
    if x1 <= x0 or y1 <= y0:
        return None
    pixels = gray.load()
    min_x, min_y = width, height
    max_x, max_y = -1, -1
    count = 0
    long_row_limit = max(20, int((x1 - x0) * 0.62))
    for y in range(y0, y1):
        row_dark: list[int] = []
        for x in range(x0, x1):
            if pixels[x, y] < threshold:
                row_dark.append(x)
        if ignore_long_rows and len(row_dark) > long_row_limit:
            continue
        for x in row_dark:
            count += 1
            if x < min_x:
                min_x = x
            if x > max_x:
                max_x = x
            if y < min_y:
                min_y = y
            if y > max_y:
                max_y = y
    if count < 8:
        return None
    return min_x, min_y, max_x, max_y


def crop_line(
    page_image: Image.Image,
    page_rect: fitz.Rect,
    number_rect: fitz.Rect,
    left_margin_pt: float,
    zoom: float,
    above_pt: float = 24.0,
    below_pt: float = 24.0,
) -> Image.Image:
    x0 = int(max(0, (left_margin_pt - 18) * zoom))
    x1 = int(min(page_image.width, (page_rect.width - left_margin_pt + 28) * zoom))
    y_center = (number_rect.y0 + number_rect.y1) / 2.0
    y0 = int(max(0, (y_center - above_pt) * zoom))
    y1 = int(min(page_image.height, (y_center + below_pt) * zoom))
    return page_image.crop((x0, y0, x1, y1))


def make_contact_sheet(crop_paths: list[Path], output: Path, cols: int = 2) -> None:
    if not crop_paths:
        return
    output.parent.mkdir(parents=True, exist_ok=True)
    font = ImageFont.load_default()
    pad = 12
    label_h = 20
    target_w = 760
    thumbs: list[tuple[Path, Image.Image]] = []
    for path in crop_paths:
        img = Image.open(path).convert("RGB")
        scale = min(1.0, target_w / img.width)
        if scale != 1.0:
            img = img.resize((int(img.width * scale), int(img.height * scale)), Image.Resampling.LANCZOS)
        thumbs.append((path, img))
    rows = math.ceil(len(thumbs) / cols)
    cell_w = max(img.width for _, img in thumbs)
    cell_h = max(img.height for _, img in thumbs) + label_h
    sheet = Image.new("RGB", (cols * cell_w + (cols + 1) * pad, rows * cell_h + (rows + 1) * pad), "white")
    draw = ImageDraw.Draw(sheet)
    for idx, (path, img) in enumerate(thumbs):
        col = idx % cols
        row = idx // cols
        x = pad + col * (cell_w + pad)
        y = pad + row * (cell_h + pad)
        draw.text((x, y), path.stem, fill=(0, 0, 0), font=font)
        sheet.paste(img, (x, y + label_h))
    sheet.save(output)


def write_csv(path: Path, rows: list[dict[str, object]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if not rows:
        path.write_text("", encoding="utf-8-sig")
        return
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--docx", required=True)
    parser.add_argument("--pdf", required=True)
    parser.add_argument("--audit-csv", required=True)
    parser.add_argument("--out-dir", required=True)
    parser.add_argument("--center-twips", type=int, required=True)
    parser.add_argument("--right-twips", type=int, required=True)
    parser.add_argument("--word-bookmark-csv")
    parser.add_argument("--limit", type=int)
    parser.add_argument("--only-template-matches", action="store_true")
    parser.add_argument("--zoom", type=float, default=3.0)
    parser.add_argument("--number-tolerance-pt", type=float, default=10.0)
    parser.add_argument("--formula-tolerance-pt", type=float, default=18.0)
    parser.add_argument("--break-formula-tolerance-pt", type=float, default=45.0)
    parser.add_argument("--band-pad-pt", type=float, default=2.0)
    parser.add_argument("--render-pages", action="store_true")
    args = parser.parse_args()

    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    crops_dir = out_dir / "formula_line_crops"
    pages_dir = out_dir / "pages"
    crops_dir.mkdir(parents=True, exist_ok=True)
    if args.render_pages:
        pages_dir.mkdir(parents=True, exist_ok=True)

    docx_path = Path(args.docx)
    left_margin_pt = first_left_margin_pt(docx_path)
    expected_center_pt = left_margin_pt + args.center_twips / 20.0
    expected_right_pt = left_margin_pt + args.right_twips / 20.0
    equations = read_equations(Path(args.audit_csv), args.limit, args.only_template_matches)
    word_page_hints = read_word_page_hints(Path(args.word_bookmark_csv)) if args.word_bookmark_csv else {}
    pdf = fitz.open(args.pdf)
    page_cache: dict[int, Image.Image] = {}
    rendered_page_paths: list[str] = []
    rows: list[dict[str, object]] = []
    crop_paths: list[Path] = []

    for order, equation in enumerate(equations, start=1):
        label = equation["text"].strip()
        number_on_own_line = is_true(equation.get("break_before_number"))
        found = find_number_rect(pdf, label, expected_right_pt, page_hint=word_page_hints.get(label))
        if found is None:
            rows.append(
                {
                    "order": order,
                    "paragraph": equation["paragraph"],
                    "label": label,
                    "page": "",
                    "status": "missing_number",
                    "number_right_delta_pt": "",
                    "formula_center_delta_pt": "",
                    "formula_bbox_pt": "",
                    "number_on_own_line": number_on_own_line,
                    "crop": "",
                }
            )
            continue
        page_index, number_rect = found
        if page_index not in page_cache:
            page_cache[page_index] = render_page(pdf, page_index, args.zoom)
            if args.render_pages:
                page_path = pages_dir / f"page-{page_index + 1:03d}.png"
                page_cache[page_index].save(page_path)
                rendered_page_paths.append(str(page_path))
        page = pdf[page_index]
        image = page_cache[page_index]
        if number_on_own_line:
            band_y0 = int(max(0, (number_rect.y0 - 56.0) * args.zoom))
            band_y1 = int(min(image.height, (number_rect.y0 - args.band_pad_pt) * args.zoom))
            formula_x1_pt = page.rect.width - left_margin_pt
        else:
            band_y0 = int(max(0, (number_rect.y0 - args.band_pad_pt) * args.zoom))
            band_y1 = int(min(image.height, (number_rect.y1 + args.band_pad_pt) * args.zoom))
            formula_x1_pt = number_rect.x0 - 8
        formula_x0 = int(max(0, (left_margin_pt - 8) * args.zoom))
        formula_x1 = int(max(0, formula_x1_pt * args.zoom))
        bbox = dark_bbox(image, formula_x0, band_y0, formula_x1, band_y1)
        formula_center_delta: float | str = ""
        bbox_pt: str = ""
        if bbox is not None:
            center_px = (bbox[0] + bbox[2]) / 2.0
            center_pt = center_px / args.zoom
            formula_center_delta = abs(center_pt - expected_center_pt)
            bbox_pt = json.dumps([round(value / args.zoom, 2) for value in bbox], ensure_ascii=False)
        number_right_delta = abs(number_rect.x1 - expected_right_pt)
        crop = crop_line(
            image,
            page.rect,
            number_rect,
            left_margin_pt,
            args.zoom,
            above_pt=68.0 if number_on_own_line else 24.0,
            below_pt=24.0,
        )
        crop_path = crops_dir / f"eq_{order:03d}_p{page_index + 1:03d}_{label.replace('（', '').replace('）', '').replace('(', '').replace(')', '')}.png"
        crop.save(crop_path)
        crop_paths.append(crop_path)
        status = "pass"
        if bbox is None:
            status = "missing_formula_ink"
        elif number_right_delta > args.number_tolerance_pt:
            status = "number_right_mismatch"
        elif (
            isinstance(formula_center_delta, float)
            and formula_center_delta > (args.break_formula_tolerance_pt if number_on_own_line else args.formula_tolerance_pt)
        ):
            status = "formula_center_mismatch"
        rows.append(
            {
                "order": order,
                "paragraph": equation["paragraph"],
                "label": label,
                "page": page_index + 1,
                "status": status,
                "number_on_own_line": number_on_own_line,
                "number_right_delta_pt": round(number_right_delta, 2),
                "formula_center_delta_pt": "" if not isinstance(formula_center_delta, float) else round(formula_center_delta, 2),
                "formula_bbox_pt": bbox_pt,
                "crop": str(crop_path),
            }
        )

    write_csv(out_dir / "formula_visual_alignment.csv", rows)
    sheet_paths: list[str] = []
    for start in range(0, len(crop_paths), 16):
        chunk = crop_paths[start : start + 16]
        if not chunk:
            continue
        sheet_path = out_dir / f"formula_crop_sheet_{start + 1:03d}_{start + len(chunk):03d}.png"
        make_contact_sheet(chunk, sheet_path)
        sheet_paths.append(str(sheet_path))
    status_counts: dict[str, int] = {}
    for row in rows:
        status = str(row["status"])
        status_counts[status] = status_counts.get(status, 0) + 1
    summary = {
        "docx": str(docx_path),
        "pdf": str(args.pdf),
        "pdf_pages": pdf.page_count,
        "equations_checked": len(rows),
        "status_counts": status_counts,
        "left_margin_pt": left_margin_pt,
        "expected_center_pt": expected_center_pt,
        "expected_right_pt": expected_right_pt,
        "word_page_hints": len(word_page_hints),
        "rendered_pages": rendered_page_paths,
        "contact_sheets": sheet_paths,
    }
    (out_dir / "formula_visual_alignment_summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(summary, ensure_ascii=True, indent=2))


if __name__ == "__main__":
    main()
