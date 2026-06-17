---
name: tongji-graduate-doc-template
description: Review, repair, and visually validate Tongji University graduate thesis Word documents. Use for .docx dissertation formatting, headings, body text, formulas and equation numbers, figure/table layout, captions and cross-references, page headers, page numbers, section breaks, references including Zenodo dataset entries, Word COM field updates, PDF rendering, and page-batched subagent visual review against local Tongji graduate document templates.
---

# Tongji Graduate DOC Template

Use this skill for Tongji University graduate thesis `.docx` work where formatting, fields, formulas, references, pagination, and visual acceptance matter. Treat it as a hard-gate workflow: understand the local official template, modify Word structures conservatively, update fields with Word COM, export PDF, render pages, and require page-batched visual review before delivery.

## Local Contract

- Work in PowerShell, not bash. Do not use `rg`, heredoc, or bash-style pipes when the repository instructions forbid them.
- Use the Python interpreter specified by the active thesis workspace. When none is specified, use the current environment's `python`.
- Preserve the user worktree. Do not revert unrelated changes.
- Use Word COM for final field update and PDF export whenever possible; use OpenXML only for controlled structural edits.
- Use PDF page renders for visual QA. Machine checks are necessary but never sufficient.
- For long documents, assign subagents page ranges in 5-page batches for visual/content review. The main agent waits, aggregates, fixes, and re-renders.

## Official Sources

Check these local files before deciding format rules:

- The user's local Tongji graduate thesis writing guide.
- The user's local Tongji graduate thesis example/template document.
- The user's local GB/T 7714 reference-format standard.

This public skill does not bundle official DOC/PDF template files. Use the local official files supplied by the user or institution. A safe summary is bundled here:

- `references/tongji-format-summary.md`

Observed reference-entry template from the Tongji example:

- Heading `参考文献`: 黑体, 16 pt, centered.
- Reference entries: 宋体/Times New Roman, 五号 10.5 pt, hanging indent 21 pt, exact line spacing 16 pt, style `正文`.
- References use sequential numeric coding by first appearance in the body.
- Body citations use superscript `[参考文献序号]`.

## Session-Derived Categories

Audit and repair by category, not by isolated keywords:

- 正文格式: normal body font, size, alignment, indentation, line spacing, paragraph spacing, accidental centering, defensive or process prose, grammar, and inline mathematical symbols embedded in prose.
- 标题格式: chapter/section/subsection font, size, bold, alignment, outline level, paragraph spacing, and accidental page breaks after headings.
- 页眉页码: chapter headers, reference/acknowledgment headers, section boundaries, page-number continuity, and `STYLEREF` behavior.
- 公式+编号: Word-native OMML, inline math, display math, sub/superscripts, fraction bars, font size, single spacing, tab stops, `SEQ Equation`, bookmarks, and `REF` citations.
- 参考文献+Zenodo: order by first appearance, body superscript fields, reference-list fields, DOI deduplication, GB/T 7714 style, Zenodo `[DS/OL]`, and unused-reference removal or semantic citation.
- 图像: figure size, caption style, caption proximity, cross-reference fields, readable evidence-chain figures, and no stranded captions.
- 表格: three-line scientific table style, no vertical borders, 小五 cell text, no double spaces, no awkward split, repeated header rows only when intentional.
- 分页、分节: section breaks, `keepNext`, `pageBreakBefore`, blank regions, heading-following flow, page header inheritance, and adjacent-page rerendering.
- 目录和交叉引用: TOC, `REF`, `PAGEREF`, `SEQ`, bookmarks, broken field text, and Word COM field refresh.
- 视觉审查和内容审查: page-rendered visual review and independent content review in 5-page batches, with the main agent aggregating and fixing.

## Workflow

1. Intake the authoritative source `.docx`, latest version number, local template files, and prior audit reports.
2. Build a change contract before editing: targeted chapters/pages, forbidden regressions, required visual gates, and output version name.
3. Inspect Word structure: `word/document.xml`, styles, bookmarks, `REF`, `SEQ`, captions, section breaks, headers, tables, equations.
4. Make scoped edits through OpenXML or Word COM. Avoid static numbers when Word fields are required.
5. Update all fields with Word COM, save as a new `.docx`, export `.pdf`, and close/clean Word processes.
6. Run machine audits: fields, bookmarks, formulas, references, table formatting, section breaks, page headers, static citations, bad text.
7. Render PDF pages and visually inspect every affected page. Use subagents in 5-page batches for thesis-level review.
8. Iterate until all hard gates pass. Do not claim completion if Word cannot open normally, PDF cannot export, or visual pages reveal layout defects.

## Hard Gates By Category

### Body Text

- Normal body text should be Chinese 宋体 小四 and English Times New Roman 小四, usually `w:sz="24"` / `w:szCs="24"`.
- Body paragraphs are normally justified with first-line indent of 2 Chinese characters; strict local audits used `w:firstLine="480"`.
- Body paragraph spacing should be before/after 0. The local skeleton used 20 pt line spacing for normal body text unless a stricter chapter contract overrides it.
- Body paragraphs may not have a stable `pStyle`. Detect by paragraph role after excluding tables, captions, formulas, TOC, references, figures, and headings.
- Body paragraphs must use the thesis normal body style, not accidental caption/heading/centered style.
- Fix centered body text after figures or captions; body following `图4-4`-type captions must resume normal alignment.
- Body text must not contain low-level grammar errors, defensive modification prose, or meta-edit phrases such as `本次修改`.
- Remove final-thesis prose defects discovered in prior sessions: `新版`, `原文所称`, `本节不重复列出`, `Phase A/B`, `timeout/OOM/crash` as engineering-log wording, `claim audit`, source file names, internal pipeline names, and manuscript/revision-status language.
- Do not insert unexplained concepts in defensive form such as `并非xxx` when the concept was not introduced.
- Inline technical symbols that are mathematical variables must be Word math or correctly styled formula runs, not plain text like `hgeo`, `scurr`, `pgoal`.
- For prose cleanup, prefer Word Find/Replace or tightly scoped paragraph/run replacement. Whole-XML reserialization or cross-node rewriting can make Word hang even when the DOCX ZIP is valid.

### Headings

- Chapter headings: 黑体, 三号 16 pt, centered, bold, single line spacing, before 24 pt, after 18 pt.
- Second-level headings: 黑体, 15 pt, left aligned, no first-line indent, single line spacing, before 24 pt, after 6 pt.
- Third-level headings: 黑体, 14 pt, left aligned, no first-line indent, single line spacing, before 12 pt, after 6 pt.
- Style IDs drift across versions, e.g. `1`, `af`, `af1`, `Heading 1`, `标题二`, or custom IDs. Inspect `styles.xml`, outline level, and template comparison instead of trusting style names alone.
- Chapter and section headings must follow the existing thesis style and numbering.
- Section headings must not force unnecessary page breaks after them unless required by the template.
- After a heading, the first body paragraph, figure, or table should flow naturally; do not leave a blank page or large empty block.
- Recheck page headers after moving headings or section breaks.

### Headers, Page Numbers, Sections

- Page headers must match the current chapter/section region. Reference pages must show `参考文献`; acknowledgment pages must show `致谢`.
- Known local header evidence from the reference repair: reference pages used `word/header32.xml`, `word/header33.xml`, and `word/header34.xml`; acknowledgment pages used `word/header35.xml`, `word/header36.xml`, and `word/header37.xml`. Treat these as examples, not constants.
- Some sections use `STYLEREF "标题 1"` for dynamic chapter headers. If headings move or section breaks change, update fields and visually verify header text.
- When references or chapters are reordered, move or recreate `sectPr` at the correct boundary. A common failure is the last reference page inheriting the `致谢` header.
- If adding `pageBreakBefore`, verify it does not move previous content into the next section header.
- Page numbers must remain continuous and visually correct after Word COM save/export.

### Formulas And Equation Numbers

- All display formulas must be Word-native OMML/OMath, not plain text, images, or malformed Unicode.
- Display formulas should contain `m:oMath` or `m:oMathPara`; plain Unicode, LaTeX residue, screenshots, or hand-styled subscript text are failures.
- If the user asks for MathType, first probe MathType/Word add-in automation in a temporary document. Prior local testing found MathType 6 and the Word add-in installed, but the conversion/numbering macros were not safe for unattended whole-thesis automation. Do not claim MathType OLE conversion unless the plugin conversion is proven on the actual machine and passes visual QA.
- Use Word OMML preservation as the default thesis-safe route. Rebuild OMML only when the source formula is actually broken.
- Inline formula symbols must preserve subscripts/superscripts/underscores visually. Known failure examples include `pgoal`, `scurr`, and `hgeo` being left as plain text or losing subscript.
- Inline `m:r` formula runs should use 小四, with effective `w:sz="24"` and `w:szCs="24"`.
- Formula font size should be 小四 where required by the thesis contract; formula paragraphs use single line spacing.
- Strict local formula audits require formula paragraphs to use `w:spacing line="240" lineRule="auto"` and no nonzero before/after spacing.
- Equation numbers must be right-aligned and use the same format as the correct earlier chapters.
- Do not accept paragraph centering as equation alignment. The strict chapter contract treats `w:jc="center"` on equation paragraphs as a failure.
- Preferred display-equation layout is: center tab before formula, formula body, one tab before the right-aligned equation number. In the chapter-4 contract the known good tab stops were approximately center `4153` DXA and right `8306` DXA, but always extract the real values from the template or a correct earlier chapter first.
- If a formula cannot fit on one visual line, do not squeeze the equation number into a broken line. Use the approved long-formula layout: formula body on the upper line, explicit Word line break before the number, then two tab characters so the equation number reaches the right tab stop.
- Equation numbers must be live fields, commonly `SEQ Equation \* ARABIC \s 1`, with equation bookmarks such as `eq_4_1`; body equation references use `REF ... \h`.
- Rebuild equation rows with a stable two/three-cell or tab-stop layout only if it matches the existing thesis pattern.
- Visually inspect formula pages after every repair. Machine OMML checks do not catch all baseline, underline, or spacing failures.
- For formula visual QA, use Word bookmark page numbers as the anchor whenever possible. PDF text search can misidentify body references such as `见式（4-22）` as the equation number and create false pass/fail results.

### Tables

- Use the thesis/third-chapter simulation-table style: scientific three-line table, top/bottom rules and header rule, no vertical borders.
- Table text should be 小五 unless the template requires otherwise.
- Table cells must not contain two-character spaces or layout-padding text.
- Tables should not split awkwardly across pages. Header rows repeat only when intended.
- Captions should not use `keep-next` in a way that creates large blank space.
- Check known failure patterns: table near page 93/95/97 leaving large blank areas, table captions stranded from tables, and tables pushed after section headings.
- In OpenXML audits, check `w:tblBorders`, `w:tcBorders`, `w:tblHeader`, `w:cantSplit`, `w:sz`, and `w:szCs`. For 小五 table text, expected direct size is usually `w:sz="18"` / `w:szCs="18"`.
- Do not accept vertical borders, left/right outer borders, or cell-level side borders in scientific tables unless the official template explicitly uses them.

### Figures And Captions

- Figures and captions must be Word cross-referenced, not static text only.
- Figure captions should use the document caption style and normal placement. In the local thesis chapters, figure captions are below figures and table captions are above tables unless the official template says otherwise.
- Do not over-integrate many results into one huge unreadable figure when the content requires a layered evidence chain.
- Avoid figure/caption keep-next behavior that creates blank space before a figure, such as the previous `图4-26` blank-page defect.
- Visually inspect before/after pages for every moved or inserted figure.
- Check whether captions contain `w:keepNext` directly or through their style. Removing caption `keepNext` may be required when it creates large blanks; validate by render, not by assumption.
- AI-generated or imagegen figures are layout/style guides unless the user explicitly asks for illustrative assets. Dissertation result figures must be redrawn from real data; do not let mockup values, failure dashboards, or generated images become quantitative evidence.

### Pagination And Blank Space

- Audit blank regions page by page, especially before/after figures, tables, section headings, and chapter boundaries.
- Delete or relax bad `keepNext`, `pageBreakBefore`, and floating-object anchoring only after checking the surrounding layout.
- A heading followed by immediate page break or a mostly blank page is a failure unless mandated by the template.
- If fixing one blank page changes page headers or references, rerender all adjacent pages.
- Before rendering, clear old `page_*.png` files from the evidence directory; stale page images can make a failed run look complete.
- Known local defect patterns from prior sessions: `图4-26` before-page blank, `表4-9` around page 93, and related large blanks around pages 95 and 97. Treat analogous pages as fixed hotspots after any figure/table repair.
- Use pixel/row whitespace scans as triage only. A page with acceptable whitespace metrics can still fail visually if a caption, table, or heading is stranded.

### References And Zenodo

- Follow the Tongji guide: list references at the end, not by chapter; order by first appearance; cite in body with superscript numeric brackets.
- Every main reference must be actually cited in the body. If a source is important and the real citation count must stay high, add it to a semantically appropriate body sentence; do not leave uncited tail references.
- Keep actual cited references over the requested lower bound when specified, e.g. more than 130.
- Merge bulk citations into ranges when consecutive after renumbering, e.g. `[22-24]`, while preserving nonconsecutive groups.
- Use Word `REF bib_x` fields and reference-list `SEQ RefSeq` fields/bookmarks. Do not make body citations static numbers.
- Detect and remove duplicate DOI entries. Redirect duplicate bookmarks to the first semantically correct source before deleting duplicates.
- English references use `et al.`, not Chinese `等`. Chinese `等` is allowed only for Chinese-language references.
- Normalize English quotes to ASCII quotes where needed; avoid Chinese punctuation in English titles unless it is truly part of the title.
- Zenodo dataset entries should be `[DS/OL]`, with `Zenodo` and DOI, e.g. `Author A, Author B, Author C, et al. Dataset title[DS/OL]. Zenodo, 2025. DOI: 10.xxxx/zenodo.xxxxx.`
- Verify DOI/source availability where possible. Cache DOI metadata for reproducibility.
- Always update fields with Word COM before a final reference audit; OpenXML field rewrites can look correct before Word refresh and fail after pagination/field update.
- After Word COM save, audit: total references, used references, unused references, duplicate DOI, bad `等`, bad quotes/punctuation, missing bookmarks, and `Error! Reference source not found`.
- Prior successful local reference cleanup evidence: an initial audit found 180 references, 26 duplicate DOI groups, 113 English entries with Chinese `等`, and 15 uncited entries. The cleaned V22 state reduced the list to 153 references, had 153 used references and 0 unused references, deleted 27 duplicate bookmarks, and rebuilt 98 body/static citations. Use these numbers only as historical sanity checks for that document, not as universal targets.
- Historical Zenodo deduplication example: keep the entry/bookmark for DOI `10.5281/zenodo.17198632` at the first semantically correct location, e.g. `bib_2`, and redirect/delete duplicate entries such as `bib_16`.

### Cross-References And TOC

- Rebuild figure, table, equation, and reference fields with Word fields rather than static text.
- Update all fields with Word COM before final PDF export.
- Scan `document.xml` for `REF`, `PAGEREF`, `SEQ`, bookmark starts/ends, and broken field result text.
- Refresh TOC/list fields if chapter titles, captions, references, or pagination changed.
- Prior successful local field inventory after repair: `TOC=1`, `REF=327`, `PAGEREF=99`, `SEQ=417`; bookmarks total 537, with `bib_*` 153, `eq_*` 169, `fig_*` 58, and `tbl_*` 28. Use as an example of the expected audit granularity.

### Visual And Content Review

- Use subagents for independent review when the user asks for agent review or when changes span many pages.
- Split visual review by 5-page blocks. Each subagent reports page number, defect type, evidence, and pass/fail.
- Use separate content-review agents for formula-symbol semantics, reference relevance, and paragraph logic. Visual review alone is not content review.
- The main agent must wait for subagent reports, summarize issues, fix, rerender, and rerun targeted review.
- Do not accept a page because a script passed. Final pass requires rendered-page inspection.
- Subagent review reports must include page number, original text or visual evidence, issue type, reason, suggested fix, severity, and PASS/FAIL. The subagent is read-only unless explicitly assigned a repair task.
- For content review, require semantic checks such as whether variables mentioned after a formula actually appear in the formula, whether inline symbols are defined, whether claims match evidence, and whether citations support the sentence.
- For visual review, require checks for formula baselines, fraction bars, sub/superscripts, missing underlines, right-aligned numbers, table borders, table text size, caption proximity, page headers, page numbers, and abnormal blank regions.

## Useful Scripts

Bundled reusable scripts:

- `scripts/reference_audit.py`: extracts main references, body citation order, duplicate DOI, uncited entries, bad `等`, and DOI metadata.
- `scripts/fix_reference_fonts.py`: post-processes reference-region `rFonts` after Word COM save so entries retain Times New Roman and 宋体 XML attributes.
- `scripts/audit_ch4_format_contract.py`: chapter-format contract audit for formula/table/reference issues from the fourth-chapter repair session.
- `scripts/docx_v19_audit.py`: OpenXML summary audit for document structure, captions, references, tables, and field state.
- `scripts/repair_ch4_after_word_update.py`: reference repair script for caption keep-next removal, figure sizing, and paragraph pagination cleanup after Word update.
- `scripts/audit_ch5_contract.py`: fifth-chapter contract audit covering caption keep-next, three-line tables, plain references, and style drift.
- `scripts/render_pdf_pages.py`: PDF-to-PNG rendering helper for visual review.
- `scripts/v17_render_pdf_pages.py`: compact alternate renderer from the thesis shared script set.
- `scripts/audit_page_whitespace.py`: PyMuPDF/PIL-style whitespace audit helper for large blank-page triage.
- `scripts/make_contact_sheet.py`: builds contact sheets from rendered page PNGs for quick page-batch visual review.
- `scripts/v16_render_visual_qa.py`: page-image visual QA helper from the cross-reference fieldization session.
- `scripts/update_fields_export_pdf.ps1`: PowerShell Word COM field update and PDF export reference.
- `scripts/update_word_fields_and_export_full.ps1`: full Word field update/export script from the main thesis pipeline.
- `scripts/export_pdf_no_save.ps1`: read-only Word COM PDF export reference when the source document must not be saved.
- `scripts/probe_mathtype_word_macros.ps1`: safe MathType/Word macro probe for temporary documents; use only to establish automation boundaries, not to run whole-thesis conversion blindly.
- `scripts/v17_equation_alignment_pipeline.py`: extracts/audits/applies equation center/right tab-stop alignment; use it to avoid paragraph-centered equation lines.
- `scripts/v17_equation_visual_qa.py`: locates formula-number evidence and produces visual crops for equation QA.
- `scripts/v16_formula_caption_crossref_pipeline.py`: reference OpenXML implementation for formula, caption, bookmark, `SEQ`, `REF`, and `PAGEREF` field handling.
- `scripts/v17_word_bookmark_pages.ps1`: uses Word COM to locate bookmark page numbers and help map `eq_*`, `fig_*`, `tbl_*`, and `bib_*` references to visual pages.
- `scripts/v18_prose_review_pipeline.py`: extracts page/paragraph review units for page-batched content review.
- `scripts/v18_word_find_replace_export.ps1`: Word COM Find/Replace and export path used after OpenXML rewriting made Word hang.
- `scripts/extract_docx_paragraph_pages.ps1`: Word COM paragraph/page extraction helper.
- `scripts/zotero_codex_reference_manager.py`: Zotero/Better BibTeX audit helper for reference-library consistency; use as an auxiliary tool, not as a replacement for Word field checks.

Bundled reference:

- `references/v19_ch4_contract.md`: fourth-chapter hard contract from the local repair session. Read it when repairing formula, table, figure, and visual QA regressions similar to chapters 4/5.
- `references/tongji-format-summary.md`: public-safe summary of the key Tongji/GB/T 7714 checks used by this skill.

Run scripts with:

```powershell
$python = '<path-to-python.exe>'
& $python '<script.py>' <args>
```

## Word COM Patterns

Use Word COM for field update and PDF export. Prefer SaveAs to a temporary docx then replace, because direct `doc.Save()` can fail after large field updates.

```powershell
$docx = '<path-to-thesis.docx>'
$tmpDocx = '<path-to-thesis.__word_tmp__.docx>'
$pdf = '<path-to-thesis.pdf>'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$doc = $word.Documents.Open($docx)
[void]$doc.Fields.Update()
foreach ($section in $doc.Sections) {
  foreach ($header in $section.Headers) { [void]$header.Range.Fields.Update() }
  foreach ($footer in $section.Footers) { [void]$footer.Range.Fields.Update() }
}
$doc.SaveAs([ref]$tmpDocx, [ref]16)
$doc.SaveAs([ref]$pdf, [ref]17)
$doc.Close([ref]$false)
$word.Quit()
Move-Item -LiteralPath $tmpDocx -Destination $docx -Force
```

Always release COM objects and kill orphaned `WINWORD` only when it is a stale automation process from the current task.

## PDF Rendering Pattern

Use PyMuPDF to render pages for visual QA:

```powershell
$python = '<path-to-python.exe>'
& $python -c "from pathlib import Path; import fitz; doc=fitz.open(Path('<path-to-thesis.pdf>')); out=Path('<path-to-renders>'); out.mkdir(exist_ok=True); [doc[i].get_pixmap(matrix=fitz.Matrix(2,2), alpha=False).save(out/f'page_{i+1:03d}.png') for i in range(0, min(5, doc.page_count))]"
```

Open rendered pages with the local image viewer or `view_image` tool. Inspect page headers, blank space, formula baselines, table borders, caption proximity, and reference formatting.

## Final Acceptance Checklist

Before final response, verify and report:

- Output `.docx` and `.pdf` paths.
- Word opens or was successfully saved/exported by Word COM.
- Field update and PDF export succeeded.
- No broken references or missing bookmarks.
- Formula pages visually pass.
- Figure/table pages visually pass.
- No large unexpected blanks around headings, captions, tables, or figures.
- Reference count, used count, unused count, duplicate DOI count, bad `等` count, and Zenodo format.
- Reference pages render with `参考文献` header; following sections start on correct pages.
- No temporary `__word_tmp__`, `__fontfix_tmp__`, or `tmp*.docx` files remain.
