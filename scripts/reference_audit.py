from __future__ import annotations

import csv
import json
import re
import sys
import time
import urllib.error
import urllib.request
import zipfile
from collections import Counter, defaultdict
from pathlib import Path
from xml.etree import ElementTree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"
NS = {"w": W_NS}


def paragraph_visible_text(p: ET.Element) -> str:
    return "".join((el.text or "") for el in p.iter() if el.tag == W + "t")


def paragraph_instr_text(p: ET.Element) -> str:
    return "".join((el.text or "") for el in p.iter() if el.tag == W + "instrText")


def paragraph_bib_bookmark(p: ET.Element) -> str:
    for el in p.iter():
        if el.tag == W + "bookmarkStart":
            name = el.attrib.get(W + "name", "")
            if name.startswith("bib_"):
                return name
    return ""


def extract_main_references(docx_path: Path):
    with zipfile.ZipFile(docx_path) as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
    paras = root.findall(".//w:body/w:p", NS)
    ref_start = None
    for idx, p in enumerate(paras):
        if paragraph_visible_text(p).strip() == "参考文献":
            ref_start = idx
            break
    if ref_start is None:
        raise RuntimeError("未找到主参考文献标题。")

    refs = []
    for idx, p in enumerate(paras[ref_start + 1 :], start=ref_start + 1):
        instr = paragraph_instr_text(p)
        if "SEQ RefSeq" not in instr:
            continue
        visible = paragraph_visible_text(p).strip()
        bookmark = paragraph_bib_bookmark(p)
        num_match = re.match(r"^\[(\d+)\]", visible)
        number = int(num_match.group(1)) if num_match else None
        doi_match = re.search(r"\bDOI:\s*([^\s。；;]+)", visible, flags=re.I)
        type_match = re.search(r"\[([A-Z]+(?:/[A-Z]+)?)\]", visible)
        refs.append(
            {
                "paragraph_index": idx,
                "number": number,
                "bookmark": bookmark,
                "type": type_match.group(1) if type_match else "",
                "doi": doi_match.group(1).rstrip(".") if doi_match else "",
                "has_chinese_deng": "等" in visible,
                "visible": visible,
            }
        )
    return root, paras, ref_start, refs


def expand_citation_content(content: str):
    result = []
    for part in re.split(r"[，,]", content):
        part = part.strip()
        if not part:
            continue
        range_match = re.match(r"^(\d+)\s*[-–]\s*(\d+)$", part)
        if range_match:
            a, b = map(int, range_match.groups())
            step = 1 if b >= a else -1
            result.extend(range(a, b + step, step))
        else:
            result.extend(int(x) for x in re.findall(r"\d+", part))
    return result


def citation_first_order(paras, ref_start: int, refs):
    number_to_bookmark = {r["number"]: r["bookmark"] for r in refs if r["number"]}
    seen = set()
    order = []
    locations = defaultdict(list)

    def add_bookmark(bookmark: str, para_idx: int, text: str):
        if not bookmark:
            return
        locations[bookmark].append({"paragraph_index": para_idx, "text": text[:180]})
        if bookmark not in seen:
            seen.add(bookmark)
            order.append(bookmark)

    for idx, p in enumerate(paras[:ref_start]):
        visible = paragraph_visible_text(p)
        instr = paragraph_instr_text(p)
        bracket_seen = False
        for match in re.finditer(r"\[\s*([0-9][0-9,，\s\-–]*)\s*\]", visible):
            nums = expand_citation_content(match.group(1))
            if nums and min(nums) >= 1 and max(nums) <= 500:
                for num in nums:
                    if num in number_to_bookmark:
                        add_bookmark(number_to_bookmark[num], idx, visible)
                bracket_seen = True
        if not bracket_seen:
            for match in re.finditer(r"REF\s+(bib_\d+)\b", instr):
                add_bookmark(match.group(1), idx, visible)
    return order, locations


def fetch_doi_metadata(doi: str, timeout: int = 8):
    url = f"https://doi.org/{doi}"
    req = urllib.request.Request(
        url,
        headers={
            "Accept": "application/vnd.citationstyles.csl+json",
            "User-Agent": "CodexReferenceAudit/1.0",
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            raw = resp.read()
        text = raw.decode("utf-8")
        return {"ok": True, "metadata": json.loads(text)}
    except Exception as exc:  # noqa: BLE001
        return {"ok": False, "error": f"{type(exc).__name__}: {exc}"}


def author_gbt_english(authors, limit: int = 3) -> str:
    names = []
    for a in authors[:limit]:
        family = (a.get("family") or a.get("literal") or "").strip()
        given = (a.get("given") or "").strip()
        if family and given:
            initials = " ".join(part[0].upper() for part in re.split(r"[\s\-]+", given) if part)
            names.append(f"{family} {initials}".strip())
        elif family:
            names.append(family)
    if len(authors) > limit:
        names.append("et al")
    return ", ".join(names)


def normalized_zenodo_entry(meta: dict) -> str:
    authors = author_gbt_english(meta.get("author", []))
    title = meta.get("title", "").strip()
    year = ""
    issued = meta.get("issued", {}).get("date-parts", [])
    if issued and issued[0]:
        year = str(issued[0][0])
    publisher = meta.get("publisher", "Zenodo").strip() or "Zenodo"
    doi = (meta.get("DOI") or "").lower()
    return f"{authors}. {title}[DS/OL]. {publisher}, {year}. DOI: {doi}."


def first_value(value):
    if isinstance(value, list):
        return value[0] if value else ""
    return value or ""


def issued_year(meta: dict) -> str:
    issued = meta.get("issued", {}).get("date-parts", [])
    if issued and issued[0]:
        return str(issued[0][0])
    return ""


def pages_or_article(meta: dict) -> str:
    page = (meta.get("page") or "").strip()
    if page:
        return page
    article = (meta.get("article-number") or "").strip()
    if article:
        return article
    return ""


def normalized_from_csl(meta: dict, fallback_type: str = "") -> str:
    authors = author_gbt_english(meta.get("author", []))
    title = (meta.get("title") or "").strip()
    year = issued_year(meta)
    doi = (meta.get("DOI") or "").lower()
    csl_type = meta.get("type", "")
    container = first_value(meta.get("container-title")).strip()
    publisher = (meta.get("publisher") or "").strip()
    volume = str(meta.get("volume") or "").strip()
    issue = str(meta.get("issue") or "").strip()
    pages = pages_or_article(meta)

    if csl_type == "dataset":
        marker = "DS/OL"
        pub = publisher or "Zenodo"
        return f"{authors}. {title}[{marker}]. {pub}, {year}. DOI: {doi}."

    if fallback_type:
        marker = fallback_type
    elif csl_type in {"article-journal", "article"}:
        marker = "J"
    elif csl_type in {"paper-conference", "proceedings-article"}:
        marker = "C"
    elif csl_type in {"book"}:
        marker = "M"
    elif csl_type in {"chapter"}:
        marker = "M"
    else:
        marker = "EB/OL"

    if marker == "J":
        tail = container
        if year:
            tail += f", {year}"
        if volume:
            tail += f", {volume}"
            if issue:
                tail += f"({issue})"
        if pages:
            tail += f": {pages}"
        if doi:
            tail += f". DOI: {doi}"
        return f"{authors}. {title}[J]. {tail}."

    if marker == "C":
        tail = container or publisher
        if year:
            tail += f". {year}" if tail else year
        if pages:
            tail += f": {pages}"
        if doi:
            tail += f". DOI: {doi}"
        return f"{authors}. {title}[C]//{tail}."

    if marker == "M":
        tail = publisher
        if year:
            tail += f", {year}" if tail else year
        if doi:
            tail += f". DOI: {doi}"
        return f"{authors}. {title}[M]. {tail}."

    tail = year
    if doi:
        tail += f". DOI: {doi}" if tail else f"DOI: {doi}"
    return f"{authors}. {title}[EB/OL]. {tail}."


def main() -> int:
    if len(sys.argv) < 3:
        print("Usage: reference_audit.py <docx> <out_dir> [--verify-doi]", file=sys.stderr)
        return 2
    docx_path = Path(sys.argv[1])
    out_dir = Path(sys.argv[2])
    verify = "--verify-doi" in sys.argv[3:]
    out_dir.mkdir(parents=True, exist_ok=True)

    _, paras, ref_start, refs = extract_main_references(docx_path)
    order, locations = citation_first_order(paras, ref_start, refs)
    order_index = {bookmark: idx + 1 for idx, bookmark in enumerate(order)}

    doi_counts = Counter(r["doi"].lower() for r in refs if r["doi"])
    visible_counts = Counter(r["visible"] for r in refs)

    audit_rows = []
    for r in refs:
        audit_rows.append(
            {
                "number": r["number"],
                "bookmark": r["bookmark"],
                "first_order": order_index.get(r["bookmark"], ""),
                "used_in_body": r["bookmark"] in order_index,
                "type": r["type"],
                "doi": r["doi"],
                "duplicate_doi": bool(r["doi"] and doi_counts[r["doi"].lower()] > 1),
                "duplicate_visible": visible_counts[r["visible"]] > 1,
                "has_chinese_deng": r["has_chinese_deng"],
                "visible": r["visible"],
            }
        )

    csv_path = out_dir / "reference_audit.csv"
    with csv_path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=list(audit_rows[0].keys()))
        writer.writeheader()
        writer.writerows(audit_rows)

    metadata_path = out_dir / "doi_metadata.json"
    if metadata_path.exists():
        try:
            doi_meta = json.loads(metadata_path.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            doi_meta = {}
    else:
        doi_meta = {}
    if verify:
        unique_dois = sorted({r["doi"].lower() for r in refs if r["doi"]})
        for idx, doi in enumerate(unique_dois, start=1):
            if doi in doi_meta:
                continue
            doi_meta[doi] = fetch_doi_metadata(doi)
            metadata_path.write_text(json.dumps(doi_meta, ensure_ascii=False, indent=2), encoding="utf-8")
            time.sleep(0.15)
            if idx % 20 == 0:
                print(f"verified {idx}/{len(unique_dois)} DOI records")
        metadata_path.write_text(json.dumps(doi_meta, ensure_ascii=False, indent=2), encoding="utf-8")

    zenodo_meta = doi_meta.get("10.5281/zenodo.17198632")
    zenodo_suggestion = ""
    if zenodo_meta and zenodo_meta.get("ok"):
        zenodo_suggestion = normalized_zenodo_entry(zenodo_meta["metadata"])

    suggestion_rows = []
    for r in refs:
        doi = r["doi"].lower()
        meta = doi_meta.get(doi, {}) if doi else {}
        suggestion = ""
        if meta.get("ok"):
            suggestion = normalized_from_csl(meta["metadata"], r["type"])
        elif r["has_chinese_deng"]:
            suggestion = r["visible"].replace("，等.", ", et al.").replace(", 等.", ", et al.")
        suggestion_rows.append(
            {
                "number": r["number"],
                "bookmark": r["bookmark"],
                "doi": r["doi"],
                "current_type": r["type"],
                "current": r["visible"],
                "suggested": suggestion,
            }
        )
    suggestions_path = out_dir / "normalized_reference_suggestions.csv"
    with suggestions_path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=list(suggestion_rows[0].keys()))
        writer.writeheader()
        writer.writerows(suggestion_rows)

    duplicate_dois = [doi for doi, count in doi_counts.items() if count > 1]
    unused = [r for r in refs if r["bookmark"] not in order_index]
    chinese_deng = [r for r in refs if r["has_chinese_deng"]]

    report = [
        "# 参考文献格式与引用顺序审计",
        "",
        "## 本地模板规则",
        "",
        "- 同济写作指南要求参考文献采用顺序编码制，按正文引用出现顺序列于文末。",
        "- 同济写作指南要求正文引用位置用上标标注参考文献序号。",
        "- `template.docx` 的 Zotero 参考文献示例采用 GB/T 7714 顺序编码制风格：英文多作者为 `et al.`，文献类型置于题名后，如 `[J]`、`[C]`、`[M]`。",
        "- 英文文献不应使用中文 `等`；中文文献可使用中文作者和中文标点。",
        "- 数据集 Zenodo 记录应按数据集/在线资源处理，建议类型标识为 `[DS/OL]`。",
        "",
        "## 当前文档审计",
        "",
        f"- 主参考文献条数：{len(refs)}",
        f"- 正文可解析引用顺序覆盖条数：{len(order)}",
        f"- 未发现正文引用的主参考文献条数：{len(unused)}",
        f"- 含中文 `等` 的主参考文献条数：{len(chinese_deng)}",
        f"- 重复 DOI：{len(duplicate_dois)}",
        "",
        "## 重复 DOI",
        "",
    ]
    if duplicate_dois:
        for doi in duplicate_dois:
            rows = [r for r in refs if r["doi"].lower() == doi]
            report.append(f"- `{doi}`: " + ", ".join(f"{r['bookmark']} / [{r['number']}]" for r in rows))
    else:
        report.append("- 未发现。")
    report.extend(["", "## Zenodo 条目建议", ""])
    report.append(zenodo_suggestion or "- DOI 元数据尚未成功解析。")
    report.extend(["", "## 未发现正文引用的条目", ""])
    if unused:
        for r in unused:
            report.append(f"- {r['bookmark']} / [{r['number']}]: {r['visible'][:220]}")
    else:
        report.append("- 无。")
    report.extend(["", "## 输出文件", ""])
    report.append(f"- `{csv_path}`")
    report.append(f"- `{suggestions_path}`")
    if verify:
        report.append(f"- `{metadata_path}`")
    (out_dir / "reference_audit_report.md").write_text("\n".join(report), encoding="utf-8")
    print(f"wrote {csv_path}")
    print(f"wrote {out_dir / 'reference_audit_report.md'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
