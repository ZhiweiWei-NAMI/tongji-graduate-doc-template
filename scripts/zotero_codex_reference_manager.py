# -*- coding: utf-8 -*-
"""Zotero/Better BibTeX reference workflow for the thesis workspace.

The script keeps Codex-side reference work explicit and auditable:

1. Load local Better BibTeX ``.bib`` exports and optional CSL JSON files.
2. Optionally fetch a Zotero Web API library snapshot when credentials are set.
3. Inspect DOCX/Markdown documents for citation keys, numeric citations, and
   embedded Zotero field citations.
4. Write a Markdown audit report plus CSV action list under ``03_参考文献``.

It intentionally does not rewrite the dissertation DOCX. Zotero's Word plugin
should remain the final authority for Word-native citation fields and refreshes.
"""

from __future__ import annotations

import argparse
import csv
import html
import json
import os
import re
import sys
import urllib.error
import urllib.parse
import urllib.request
import zipfile
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET


SCRIPT_DIR = Path(__file__).resolve().parent
THESIS_DIR = SCRIPT_DIR.parent
WORKSPACE_DIR = THESIS_DIR.parent
REF_DIR = THESIS_DIR / "03_参考文献"
DEFAULT_BIB_DIR = REF_DIR / "小论文BibTeX"
DEFAULT_REPORT = REF_DIR / "10_Zotero_Codex_文献库审计.md"
DEFAULT_ACTIONS = REF_DIR / "11_Zotero_Codex_待处理清单.csv"
DEFAULT_ZOTERO_SNAPSHOT = REF_DIR / "12_Zotero_API_快照.json"

XML_NAMESPACES = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}

LATEX_COMMAND_RE = re.compile(r"\\[a-zA-Z]+\*?(?:\[[^\]]*\])?(?:\{([^{}]*)\})?")
WHITESPACE_RE = re.compile(r"\s+")
DOI_RE = re.compile(r"\b10\.\d{4,9}/[^\s\"<>]+", re.IGNORECASE)
YEAR_RE = re.compile(r"\b(18|19|20|21)\d{2}[a-z]?\b")
MAX_NUMERIC_CITATION = 500


@dataclass
class Reference:
    key: str
    entry_type: str = ""
    title: str = ""
    year: str = ""
    authors: list[str] = field(default_factory=list)
    doi: str = ""
    url: str = ""
    source: str = ""
    source_kind: str = ""
    raw: dict[str, object] = field(default_factory=dict)

    @property
    def normalized_title(self) -> str:
        return normalize_title(self.title)

    @property
    def normalized_doi(self) -> str:
        return normalize_doi(self.doi)

    def author_summary(self) -> str:
        if not self.authors:
            return ""
        if len(self.authors) == 1:
            return self.authors[0]
        return f"{self.authors[0]} et al."


@dataclass
class DocumentAudit:
    path: Path
    citation_keys: set[str] = field(default_factory=set)
    numeric_citations: set[int] = field(default_factory=set)
    zotero_item_uris: set[str] = field(default_factory=set)
    zotero_item_titles: set[str] = field(default_factory=set)
    zotero_item_dois: set[str] = field(default_factory=set)
    reference_entries: list[str] = field(default_factory=list)
    extraction_notes: list[str] = field(default_factory=list)


def rel(path: Path) -> str:
    try:
        return str(path.resolve().relative_to(WORKSPACE_DIR.resolve()))
    except ValueError:
        return str(path)


def read_text(path: Path) -> str:
    for encoding in ("utf-8-sig", "utf-8", "gb18030"):
        try:
            return path.read_text(encoding=encoding)
        except UnicodeDecodeError:
            continue
    return path.read_text(encoding="utf-8", errors="replace")


def normalize_space(value: str) -> str:
    return WHITESPACE_RE.sub(" ", value).strip()


def strip_outer_braces(value: str) -> str:
    value = value.strip()
    while len(value) >= 2 and value[0] == "{" and value[-1] == "}":
        depth = 0
        balanced = True
        for idx, char in enumerate(value):
            if char == "{":
                depth += 1
            elif char == "}":
                depth -= 1
                if depth == 0 and idx != len(value) - 1:
                    balanced = False
                    break
        if not balanced:
            break
        value = value[1:-1].strip()
    return value


def clean_latex(value: str) -> str:
    value = strip_outer_braces(value)
    replacements = {
        "\\&": "&",
        "\\_": "_",
        "\\%": "%",
        "\\#": "#",
        "\\$": "$",
        "\\textendash": "-",
        "\\textemdash": "-",
        "\\textbackslash": "\\",
        "{\\i}": "i",
    }
    for src, dst in replacements.items():
        value = value.replace(src, dst)
    value = re.sub(r"\\url\{([^{}]+)\}", r"\1", value)
    value = re.sub(r"\\doi\{([^{}]+)\}", r"\1", value)
    value = LATEX_COMMAND_RE.sub(lambda match: match.group(1) or "", value)
    value = value.replace("{", "").replace("}", "")
    return normalize_space(value)


def normalize_title(value: str) -> str:
    value = clean_latex(value).casefold()
    value = re.sub(r"[^0-9a-z\u4e00-\u9fff]+", " ", value)
    return normalize_space(value)


def normalize_doi(value: str) -> str:
    value = clean_latex(value).strip()
    value = re.sub(r"^https?://(dx\.)?doi\.org/", "", value, flags=re.IGNORECASE)
    value = re.sub(r"^doi:\s*", "", value, flags=re.IGNORECASE)
    value = value.rstrip(".,;")
    return value.casefold()


def split_top_level(value: str, delimiter: str = ",") -> list[str]:
    parts: list[str] = []
    start = 0
    brace_depth = 0
    quote_open = False
    escaped = False
    for idx, char in enumerate(value):
        if escaped:
            escaped = False
            continue
        if char == "\\":
            escaped = True
            continue
        if char == '"' and brace_depth == 0:
            quote_open = not quote_open
            continue
        if not quote_open:
            if char == "{":
                brace_depth += 1
            elif char == "}":
                brace_depth = max(0, brace_depth - 1)
            elif char == delimiter and brace_depth == 0:
                parts.append(value[start:idx].strip())
                start = idx + 1
    tail = value[start:].strip()
    if tail:
        parts.append(tail)
    return parts


def parse_bib_value(value: str) -> str:
    chunks = split_top_level(value, delimiter="#")
    cleaned: list[str] = []
    for chunk in chunks:
        chunk = chunk.strip().rstrip(",")
        if len(chunk) >= 2 and chunk[0] == '"' and chunk[-1] == '"':
            chunk = chunk[1:-1]
        cleaned.append(clean_latex(chunk))
    return normalize_space("".join(cleaned))


def find_bib_entries(text: str) -> Iterable[tuple[str, str, str]]:
    idx = 0
    while idx < len(text):
        at = text.find("@", idx)
        if at < 0:
            break
        type_match = re.match(r"@([A-Za-z]+)\s*([({])", text[at:])
        if not type_match:
            idx = at + 1
            continue
        entry_type = type_match.group(1)
        opener = type_match.group(2)
        closer = "}" if opener == "{" else ")"
        body_start = at + type_match.end()
        depth = 1
        quote_open = False
        escaped = False
        pos = body_start
        while pos < len(text):
            char = text[pos]
            if escaped:
                escaped = False
            elif char == "\\":
                escaped = True
            elif char == '"' and depth == 1:
                quote_open = not quote_open
            elif not quote_open:
                if char == opener:
                    depth += 1
                elif char == closer:
                    depth -= 1
                    if depth == 0:
                        break
            pos += 1
        if depth != 0:
            break
        body = text[body_start:pos]
        comma = body.find(",")
        if comma >= 0:
            key = body[:comma].strip()
            fields_text = body[comma + 1 :]
            yield entry_type, key, fields_text
        idx = pos + 1


def parse_bibtex(path: Path) -> list[Reference]:
    text = read_text(path)
    references: list[Reference] = []
    for entry_type, key, fields_text in find_bib_entries(text):
        fields: dict[str, str] = {}
        for field_text in split_top_level(fields_text):
            if "=" not in field_text:
                continue
            name, value = field_text.split("=", 1)
            fields[name.strip().casefold()] = parse_bib_value(value.strip())
        authors = []
        if fields.get("author"):
            authors = [clean_latex(item) for item in re.split(r"\s+and\s+", fields["author"], flags=re.IGNORECASE)]
        year = fields.get("year") or ""
        if not year and fields.get("date"):
            match = YEAR_RE.search(fields["date"])
            year = match.group(0) if match else ""
        doi = fields.get("doi") or ""
        if not doi:
            match = DOI_RE.search(" ".join(fields.values()))
            doi = match.group(0) if match else ""
        references.append(
            Reference(
                key=clean_latex(key),
                entry_type=entry_type,
                title=fields.get("title", ""),
                year=year,
                authors=authors,
                doi=doi,
                url=fields.get("url", ""),
                source=rel(path),
                source_kind="bibtex",
                raw=dict(fields),
            )
        )
    return references


def csl_creators_to_authors(item: dict[str, object]) -> list[str]:
    authors: list[str] = []
    creators = item.get("author") or item.get("creators") or []
    if not isinstance(creators, list):
        return authors
    for creator in creators:
        if not isinstance(creator, dict):
            continue
        literal = creator.get("literal") or creator.get("name")
        if literal:
            authors.append(str(literal))
            continue
        family = creator.get("family") or creator.get("lastName")
        given = creator.get("given") or creator.get("firstName")
        name = " ".join(str(part) for part in (given, family) if part)
        if name:
            authors.append(name)
    return authors


def parse_csl_json(path: Path) -> list[Reference]:
    data = json.loads(read_text(path))
    if isinstance(data, dict):
        data = data.get("items", [])
    if not isinstance(data, list):
        raise ValueError(f"CSL JSON must contain a list of items: {path}")
    references: list[Reference] = []
    for idx, item in enumerate(data, 1):
        if not isinstance(item, dict):
            continue
        key = str(item.get("citation-key") or item.get("id") or item.get("key") or f"csl_item_{idx}")
        title = item.get("title") or item.get("container-title") or ""
        issued = item.get("issued") or {}
        year = ""
        if isinstance(issued, dict):
            date_parts = issued.get("date-parts")
            if isinstance(date_parts, list) and date_parts and isinstance(date_parts[0], list) and date_parts[0]:
                year = str(date_parts[0][0])
        if not year:
            date_text = str(item.get("date") or item.get("issued") or "")
            match = YEAR_RE.search(date_text)
            year = match.group(0) if match else ""
        references.append(
            Reference(
                key=key,
                entry_type=str(item.get("type") or ""),
                title=str(title),
                year=year,
                authors=csl_creators_to_authors(item),
                doi=str(item.get("DOI") or item.get("doi") or ""),
                url=str(item.get("URL") or item.get("url") or ""),
                source=rel(path),
                source_kind="csl-json",
                raw=item,
            )
        )
    return references


def zotero_data_to_reference(item: dict[str, object], source: str) -> Reference | None:
    data = item.get("data", item)
    if not isinstance(data, dict):
        return None
    item_type = str(data.get("itemType") or data.get("type") or "")
    if item_type in {"attachment", "note", "annotation"}:
        return None
    key = str(data.get("citationKey") or item.get("key") or data.get("key") or "")
    if not key:
        return None
    creators = data.get("creators") or []
    authors: list[str] = []
    if isinstance(creators, list):
        for creator in creators:
            if not isinstance(creator, dict):
                continue
            if creator.get("creatorType") not in {None, "author", "editor"}:
                continue
            name = creator.get("name")
            if not name:
                name = " ".join(str(part) for part in (creator.get("firstName"), creator.get("lastName")) if part)
            if name:
                authors.append(str(name))
    date_text = str(data.get("date") or "")
    year_match = YEAR_RE.search(date_text)
    return Reference(
        key=key,
        entry_type=item_type,
        title=str(data.get("title") or ""),
        year=year_match.group(0) if year_match else "",
        authors=authors,
        doi=str(data.get("DOI") or ""),
        url=str(data.get("url") or ""),
        source=source,
        source_kind="zotero-api",
        raw=data,
    )


def fetch_zotero_items(library: str, api_key: str, collection: str | None = None, limit: int = 100) -> list[dict[str, object]]:
    if not re.fullmatch(r"(users|groups)/\d+", library):
        raise ValueError("Zotero library must look like users/123456 or groups/123456")
    endpoint = f"https://api.zotero.org/{library}"
    if collection:
        endpoint += f"/collections/{urllib.parse.quote(collection)}/items"
    else:
        endpoint += "/items/top"

    headers = {
        "Zotero-API-Version": "3",
        "Zotero-API-Key": api_key,
        "User-Agent": "Codex-Thesis-Zotero-Audit/1.0",
    }
    items: list[dict[str, object]] = []
    start = 0
    while True:
        params = urllib.parse.urlencode(
            {
                "format": "json",
                "include": "data",
                "limit": str(limit),
                "start": str(start),
            }
        )
        request = urllib.request.Request(f"{endpoint}?{params}", headers=headers)
        try:
            with urllib.request.urlopen(request, timeout=30) as response:
                payload = response.read().decode("utf-8")
        except urllib.error.HTTPError as exc:
            detail = exc.read().decode("utf-8", errors="replace")
            raise RuntimeError(f"Zotero API HTTP {exc.code}: {detail}") from exc
        batch = json.loads(payload)
        if not isinstance(batch, list):
            raise RuntimeError("Unexpected Zotero API response; expected a JSON list")
        items.extend(batch)
        if len(batch) < limit:
            break
        start += limit
    return items


def extract_docx_xml_text(path: Path) -> tuple[str, list[str]]:
    parts = [
        "word/document.xml",
        "word/footnotes.xml",
        "word/endnotes.xml",
    ]
    notes: list[str] = []
    texts: list[str] = []
    with zipfile.ZipFile(path) as docx:
        parts.extend(name for name in docx.namelist() if re.fullmatch(r"word/(header|footer)\d+\.xml", name))
        for part in parts:
            if part not in docx.namelist():
                continue
            try:
                root = ET.fromstring(docx.read(part))
            except ET.ParseError as exc:
                notes.append(f"{part}: XML parse failed: {exc}")
                continue
            part_text: list[str] = []
            for elem in root.iter():
                tag = elem.tag.rsplit("}", 1)[-1]
                if tag in {"t", "instrText", "delText"} and elem.text:
                    part_text.append(elem.text)
                elif tag in {"tab"}:
                    part_text.append("\t")
                elif tag in {"br", "cr", "p"}:
                    part_text.append("\n")
            texts.append("\n".join(part_text))
    return "\n".join(texts), notes


def extract_zotero_field_payloads(text: str) -> list[dict[str, object]]:
    payloads: list[dict[str, object]] = []
    marker = "CSL_CITATION"
    start = 0
    while True:
        idx = text.find(marker, start)
        if idx < 0:
            break
        brace = text.find("{", idx)
        if brace < 0:
            break
        depth = 0
        in_string = False
        escaped = False
        pos = brace
        while pos < len(text):
            char = text[pos]
            if escaped:
                escaped = False
            elif char == "\\":
                escaped = True
            elif char == '"':
                in_string = not in_string
            elif not in_string:
                if char == "{":
                    depth += 1
                elif char == "}":
                    depth -= 1
                    if depth == 0:
                        raw = text[brace : pos + 1]
                        try:
                            payloads.append(json.loads(raw))
                        except json.JSONDecodeError:
                            pass
                        start = pos + 1
                        break
            pos += 1
        else:
            break
    return payloads


def extract_citation_keys(text: str) -> set[str]:
    keys: set[str] = set()
    for match in re.finditer(r"\\(?:cite|parencite|textcite|autocite|citep|citet)\w*\*?(?:\[[^\]]*\]){0,2}\{([^{}]+)\}", text):
        for key in match.group(1).split(","):
            key = key.strip()
            if key:
                keys.add(key)
    for match in re.finditer(r"(?<![\w.-])@([A-Za-z0-9_:+./-]+)", text):
        keys.add(match.group(1).rstrip(".,;:"))
    return keys


def extract_numeric_citations(text: str) -> set[int]:
    numbers: set[int] = set()
    for match in re.finditer(r"(?<!\w)[\[\(（【]([0-9,\-–—，、\s]+)[\]\)）】]", text):
        content = match.group(1)
        if not re.search(r"\d", content):
            continue
        for part in re.split(r"[,，、]\s*", content):
            part = part.strip()
            if not part:
                continue
            range_match = re.fullmatch(r"(\d+)\s*[-–—]\s*(\d+)", part)
            if range_match:
                start, end = int(range_match.group(1)), int(range_match.group(2))
                if 1 <= start <= end <= MAX_NUMERIC_CITATION:
                    numbers.update(range(start, end + 1))
                continue
            if part.isdigit():
                number = int(part)
                if 1 <= number <= MAX_NUMERIC_CITATION:
                    numbers.add(number)
    return numbers


def extract_reference_entries(text: str) -> list[str]:
    lines = [normalize_space(line) for line in text.splitlines()]
    entries: list[str] = []
    in_refs = False
    current = ""
    for line in lines:
        if not line:
            continue
        if re.fullmatch(r"(参考文献|References|REFERENCES)", line):
            in_refs = True
            current = ""
            continue
        if not in_refs:
            continue
        if re.fullmatch(r"(附录.*|Appendix.*|致谢|攻读.*成果|声明)", line):
            break
        entry_match = re.match(r"^\[?(\d{1,4})\]?[\.\]、\s]+(.+)", line)
        if entry_match:
            if current:
                entries.append(current)
            current = line
        elif current:
            current = normalize_space(current + " " + line)
    if current:
        entries.append(current)
    return entries


def audit_document(path: Path) -> DocumentAudit:
    audit = DocumentAudit(path=path)
    suffix = path.suffix.casefold()
    if suffix == ".docx":
        text, notes = extract_docx_xml_text(path)
        audit.extraction_notes.extend(notes)
        xml_text = html.unescape(text)
        audit.citation_keys.update(extract_citation_keys(xml_text))
        audit.numeric_citations.update(extract_numeric_citations(xml_text))
        audit.reference_entries.extend(extract_reference_entries(xml_text))
        for payload in extract_zotero_field_payloads(xml_text):
            items = payload.get("citationItems", [])
            if not isinstance(items, list):
                continue
            for item in items:
                if not isinstance(item, dict):
                    continue
                uris = item.get("uris", [])
                if isinstance(uris, list):
                    audit.zotero_item_uris.update(str(uri) for uri in uris)
                item_data = item.get("itemData", {})
                if isinstance(item_data, dict):
                    title = item_data.get("title")
                    doi = item_data.get("DOI") or item_data.get("doi")
                    if title:
                        audit.zotero_item_titles.add(normalize_title(str(title)))
                    if doi:
                        audit.zotero_item_dois.add(normalize_doi(str(doi)))
    else:
        text = read_text(path)
        audit.citation_keys.update(extract_citation_keys(text))
        audit.numeric_citations.update(extract_numeric_citations(text))
        audit.reference_entries.extend(extract_reference_entries(text))
    return audit


def load_references(paths: list[Path]) -> list[Reference]:
    references: list[Reference] = []
    for path in paths:
        suffix = path.suffix.casefold()
        if suffix in {".bib", ".bibtex"}:
            references.extend(parse_bibtex(path))
        elif suffix in {".json", ".csljson"}:
            references.extend(parse_csl_json(path))
        else:
            raise ValueError(f"Unsupported reference file: {path}")
    return references


def default_bib_paths() -> list[Path]:
    if not DEFAULT_BIB_DIR.exists():
        return []
    return sorted(DEFAULT_BIB_DIR.glob("*.bib"))


def group_duplicates(references: list[Reference], attr: str) -> dict[str, list[Reference]]:
    grouped: dict[str, list[Reference]] = {}
    for ref in references:
        value = getattr(ref, attr)
        if callable(value):
            value = value()
        if not value:
            continue
        grouped.setdefault(str(value), []).append(ref)
    return {key: refs for key, refs in grouped.items() if len(refs) > 1}


def audit_reference_quality(ref: Reference) -> list[str]:
    issues: list[str] = []
    if not ref.key:
        issues.append("missing citation key")
    if not ref.title:
        issues.append("missing title")
    if not ref.year:
        issues.append("missing year")
    if not ref.authors:
        issues.append("missing author/editor")
    if ref.doi and not DOI_RE.search(normalize_doi(ref.doi)):
        issues.append("suspicious DOI")
    if any(token in ref.title for token in ("\\", "{", "}")):
        issues.append("LaTeX residue in title")
    return issues


def write_actions(path: Path, rows: list[dict[str, str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    headers = ["priority", "category", "key", "title", "detail", "source"]
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=headers)
        writer.writeheader()
        for row in rows:
            writer.writerow({header: row.get(header, "") for header in headers})


def render_report(
    references: list[Reference],
    documents: list[DocumentAudit],
    actions: list[dict[str, str]],
    zotero_snapshot: Path | None,
) -> str:
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    by_source: dict[str, int] = {}
    for ref in references:
        by_source[ref.source] = by_source.get(ref.source, 0) + 1
    duplicate_keys: dict[str, list[Reference]] = {}
    for ref in references:
        if ref.key:
            duplicate_keys.setdefault(ref.key, []).append(ref)
    duplicate_keys = {key: refs for key, refs in duplicate_keys.items() if len(refs) > 1}
    duplicate_dois = group_duplicates(references, "normalized_doi")
    duplicate_titles = group_duplicates(references, "normalized_title")

    cited_keys = set().union(*(doc.citation_keys for doc in documents)) if documents else set()
    library_keys = {ref.key for ref in references if ref.key}
    uncited_keys = library_keys - cited_keys if cited_keys else set()
    missing_keys = cited_keys - library_keys if cited_keys else set()

    lines: list[str] = [
        "# Zotero/Codex 文献库审计",
        "",
        f"- 更新时间：{now}",
        f"- 文献条目总数：{len(references)}",
        f"- 数据源数量：{len(by_source)}",
        f"- 审计文档数量：{len(documents)}",
        f"- 待处理事项：{len(actions)}",
    ]
    if zotero_snapshot:
        lines.append(f"- Zotero API 快照：`{rel(zotero_snapshot)}`")

    lines.extend(["", "## 数据源统计", ""])
    if by_source:
        for source, count in sorted(by_source.items()):
            lines.append(f"- `{source}`：{count} 条")
    else:
        lines.append("- 未读取到文献库。")

    lines.extend(["", "## 文档引用概览", ""])
    if documents:
        for doc in documents:
            lines.append(f"### `{rel(doc.path)}`")
            lines.append(f"- citation key：{len(doc.citation_keys)}")
            lines.append(f"- 数字型引用编号：{len(doc.numeric_citations)}")
            lines.append(f"- Zotero 字段引用 URI：{len(doc.zotero_item_uris)}")
            lines.append(f"- 参考文献条目候选：{len(doc.reference_entries)}")
            if doc.numeric_citations:
                min_num, max_num = min(doc.numeric_citations), max(doc.numeric_citations)
                lines.append(f"- 数字型引用范围：[{min_num}] - [{max_num}]")
            if doc.extraction_notes:
                for note in doc.extraction_notes:
                    lines.append(f"- 提取提示：{note}")
            lines.append("")
    else:
        lines.append("- 未指定 DOCX 或 Markdown 文档。")

    lines.extend(["", "## 一致性检查", ""])
    lines.append(f"- 重复 citation key：{len(duplicate_keys)}")
    lines.append(f"- 重复 DOI：{len(duplicate_dois)}")
    lines.append(f"- 重复标题：{len(duplicate_titles)}")
    if cited_keys:
        lines.append(f"- 文档引用 key 数：{len(cited_keys)}")
        lines.append(f"- 文献库缺失 key：{len(missing_keys)}")
        lines.append(f"- 文献库未被 key 直接引用：{len(uncited_keys)}")
    else:
        lines.append("- 文档未发现 Pandoc/LaTeX citation key；若正文使用 Word 数字域或 Zotero 字段，这是正常现象。")

    if duplicate_keys:
        lines.extend(["", "### 重复 citation key", ""])
        for key, refs in sorted(duplicate_keys.items()):
            lines.append(f"- `{key}`：{len(refs)} 次")
            for ref in refs:
                lines.append(f"  - `{ref.source}` | {ref.year} | {ref.title[:120]}")

    if missing_keys:
        lines.extend(["", "### 文档引用但文献库缺失的 key", ""])
        for key in sorted(missing_keys):
            lines.append(f"- `{key}`")

    quality_rows = [(ref, audit_reference_quality(ref)) for ref in references]
    quality_rows = [(ref, issues) for ref, issues in quality_rows if issues]
    lines.extend(["", "## 条目质量问题", ""])
    if not quality_rows:
        lines.append("- 未发现基础字段缺失或明显格式问题。")
    else:
        for ref, issues in quality_rows[:200]:
            lines.append(f"- `{ref.key}` | {ref.year} | {ref.title[:120]}")
            lines.append(f"  - 问题：{'; '.join(issues)}")
            lines.append(f"  - 来源：`{ref.source}`")
        if len(quality_rows) > 200:
            lines.append(f"- 另有 {len(quality_rows) - 200} 条未在报告中展开，详见 CSV。")

    lines.extend(["", "## 建议工作流", ""])
    lines.append("1. 在 Zotero 中维护真实文献条目，并用 Better BibTeX 自动导出 `.bib` 到 `03_参考文献`。")
    lines.append("2. 用本脚本检查重复 key、缺失 DOI/年份/作者，以及正文引用和文献库的一致性。")
    lines.append("3. Word 正文中的正式 citation/bibliography 仍用 Zotero Word 插件插入和刷新。")
    lines.append("4. 每次大改正文前后运行一次审计，保留 `10_...md` 与 `11_...csv` 作为可追溯证据。")
    return "\n".join(lines) + "\n"


def build_actions(references: list[Reference], documents: list[DocumentAudit]) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    by_key: dict[str, list[Reference]] = {}
    for ref in references:
        if ref.key:
            by_key.setdefault(ref.key, []).append(ref)
    for key, refs in sorted(by_key.items()):
        if len(refs) > 1:
            rows.append(
                {
                    "priority": "P1",
                    "category": "duplicate-key",
                    "key": key,
                    "title": refs[0].title,
                    "detail": f"{len(refs)} entries share the same citation key",
                    "source": "; ".join(ref.source for ref in refs),
                }
            )

    for doi, refs in sorted(group_duplicates(references, "normalized_doi").items()):
        rows.append(
            {
                "priority": "P2",
                "category": "duplicate-doi",
                "key": "; ".join(ref.key for ref in refs if ref.key),
                "title": refs[0].title,
                "detail": f"{len(refs)} entries share DOI {doi}",
                "source": "; ".join(ref.source for ref in refs),
            }
        )

    for title, refs in sorted(group_duplicates(references, "normalized_title").items()):
        rows.append(
            {
                "priority": "P2",
                "category": "duplicate-title",
                "key": "; ".join(ref.key for ref in refs if ref.key),
                "title": refs[0].title,
                "detail": f"{len(refs)} entries share a normalized title",
                "source": "; ".join(ref.source for ref in refs),
            }
        )

    for ref in references:
        for issue in audit_reference_quality(ref):
            rows.append(
                {
                    "priority": "P2",
                    "category": "metadata-quality",
                    "key": ref.key,
                    "title": ref.title,
                    "detail": issue,
                    "source": ref.source,
                }
            )

    cited_keys = set().union(*(doc.citation_keys for doc in documents)) if documents else set()
    library_keys = {ref.key for ref in references if ref.key}
    for key in sorted(cited_keys - library_keys):
        rows.append(
            {
                "priority": "P1",
                "category": "missing-library-entry",
                "key": key,
                "title": "",
                "detail": "Document cites this key, but no loaded BibTeX/CSL/Zotero entry provides it",
                "source": "; ".join(rel(doc.path) for doc in documents if key in doc.citation_keys),
            }
        )

    for doc in documents:
        if doc.numeric_citations and doc.reference_entries:
            max_cited = max(doc.numeric_citations)
            if max_cited > len(doc.reference_entries):
                rows.append(
                    {
                        "priority": "P1",
                        "category": "numeric-citation-range",
                        "key": "",
                        "title": rel(doc.path),
                        "detail": f"max cited number [{max_cited}] exceeds extracted reference entry count {len(doc.reference_entries)}",
                        "source": rel(doc.path),
                    }
                )
    return rows


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Audit Zotero/Better BibTeX references against thesis documents.")
    parser.add_argument("--bib", action="append", type=Path, help="Better BibTeX .bib file. Can be repeated.")
    parser.add_argument("--csl-json", action="append", type=Path, help="CSL JSON export. Can be repeated.")
    parser.add_argument("--docx", action="append", type=Path, help="DOCX document to inspect. Can be repeated.")
    parser.add_argument("--markdown", action="append", type=Path, help="Markdown document to inspect. Can be repeated.")
    parser.add_argument("--report", type=Path, default=DEFAULT_REPORT, help="Markdown audit report path.")
    parser.add_argument("--actions", type=Path, default=DEFAULT_ACTIONS, help="CSV action list path.")
    parser.add_argument("--zotero-library", default=os.environ.get("ZOTERO_LIBRARY", ""), help="Zotero library, e.g. users/123456 or groups/123456.")
    parser.add_argument("--zotero-api-key", default=os.environ.get("ZOTERO_API_KEY", ""), help="Zotero API key. Prefer environment variable.")
    parser.add_argument("--zotero-collection", default=os.environ.get("ZOTERO_COLLECTION", ""), help="Optional Zotero collection key.")
    parser.add_argument("--fetch-zotero", action="store_true", help="Fetch Zotero Web API items and include them in the audit.")
    parser.add_argument("--zotero-snapshot", type=Path, default=DEFAULT_ZOTERO_SNAPSHOT, help="Where to write the Zotero API snapshot JSON.")
    return parser.parse_args(argv)


def main(argv: list[str]) -> int:
    args = parse_args(argv)
    reference_paths: list[Path] = []
    if args.bib:
        reference_paths.extend(args.bib)
    else:
        reference_paths.extend(default_bib_paths())
    if args.csl_json:
        reference_paths.extend(args.csl_json)

    references = load_references(reference_paths)
    zotero_snapshot_path: Path | None = None
    if args.fetch_zotero:
        if not args.zotero_library or not args.zotero_api_key:
            raise SystemExit("--fetch-zotero requires --zotero-library and --zotero-api-key, or ZOTERO_LIBRARY/ZOTERO_API_KEY env vars")
        items = fetch_zotero_items(args.zotero_library, args.zotero_api_key, args.zotero_collection or None)
        args.zotero_snapshot.parent.mkdir(parents=True, exist_ok=True)
        args.zotero_snapshot.write_text(json.dumps(items, ensure_ascii=False, indent=2), encoding="utf-8")
        zotero_snapshot_path = args.zotero_snapshot
        references.extend(
            ref
            for item in items
            if (ref := zotero_data_to_reference(item, rel(args.zotero_snapshot))) is not None
        )

    document_paths: list[Path] = []
    if args.docx:
        document_paths.extend(args.docx)
    if args.markdown:
        document_paths.extend(args.markdown)
    documents = [audit_document(path) for path in document_paths]
    actions = build_actions(references, documents)

    args.report.parent.mkdir(parents=True, exist_ok=True)
    args.report.write_text(render_report(references, documents, actions, zotero_snapshot_path), encoding="utf-8")
    write_actions(args.actions, actions)

    print(f"references={len(references)}")
    print(f"documents={len(documents)}")
    print(f"actions={len(actions)}")
    print(f"report={args.report}")
    print(f"actions_csv={args.actions}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
