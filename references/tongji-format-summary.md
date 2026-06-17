# Tongji Graduate Thesis Format Summary

This public-safe reference summarizes the local checks used by the skill. It is not a replacement for the official Tongji University graduate thesis guide or the GB/T 7714 standard. Always prefer the official files supplied by the user or institution.

## Core Word Checks

- Use the official thesis template as the source of truth for margins, headers, page numbers, heading styles, captions, body text, formulas, and reference entries.
- Update all Word fields with Word COM before final review.
- Export to PDF and render pages to images before accepting layout.
- Review visual pages in small batches, especially formula, table, figure, reference, and section-break pages.

## Body And Headings

- Body text should follow the thesis normal style, with Chinese and English fonts handled separately.
- Heading levels should be identified by style and outline level, not by text alone.
- Body text after captions and figures should not inherit centered or caption formatting.
- Remove draft, revision, process-log, and defensive wording from final thesis prose.

## Formulas

- Display formulas should be Word-native OMML/OMath whenever possible.
- Formula numbers should use Word fields and right-aligned tab-stop layout.
- Inline mathematical symbols with subscripts, superscripts, fractions, or special notation should preserve formula semantics.
- Long formulas may use a two-line layout: formula body above and equation number on the next line at the right tab stop.

## Tables And Figures

- Tables should use the scientific three-line style unless the official template requires otherwise.
- Table text should use the template's table font size and must not contain padding spaces used as alignment.
- Table captions normally appear above tables; figure captions normally appear below figures.
- Captions and objects must remain visually paired after Word pagination.

## References

- References should follow sequential numeric coding by first appearance in the body.
- Body citations should use superscript numeric brackets.
- Reference-list entries should use Word fields/bookmarks when automated references are required.
- English multi-author entries use `et al.`; Chinese `等` is reserved for Chinese-language references.
- Zenodo datasets should use a dataset/online-resource marker such as `[DS/OL]`, with repository name and DOI.

## Privacy Boundary

Do not commit user thesis files, local absolute paths, private datasets, rendered page images, temporary Word files, or institution-supplied template binaries into this public skill repository.
