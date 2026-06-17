# Tongji Graduate DOC Template Skill

Codex skill for reviewing, repairing, and visually validating Tongji University graduate thesis Word documents.

It focuses on DOCX formatting, Word fields, equations, figure/table captions, cross-references, pagination, references, Word COM export, PDF rendering, and page-batched visual review.

## Contents

- `SKILL.md`: main Codex skill instructions.
- `agents/openai.yaml`: UI metadata for Codex skill discovery.
- `scripts/`: reusable PowerShell and Python helpers for DOCX audits, field updates, PDF rendering, formula checks, reference checks, and visual QA.
- `references/`: public-safe notes and contracts. Official DOC/PDF template files are not bundled.

## Privacy

This repository is intended to contain only the public skill package. It should not include local thesis drafts, personal data, private paths, generated page images, Word lock files, or institution-supplied template binaries.

Users should provide their own local official Tongji template/reference files when running the skill on a real thesis.

## Validation

Validate the skill folder with:

```powershell
$env:PYTHONUTF8 = '1'
python '<path-to-skill-creator>\scripts\quick_validate.py' '<path-to-this-skill>'
```

For thesis work, final acceptance should also include Word COM field refresh, PDF export, rendered-page visual review, and targeted content review.
