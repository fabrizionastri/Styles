**Developer Summary**
Goal: support two-way authoring of contracts in both MS Word and Markdown, with style fidelity preserved and a Markdown-first workflow for scale and AI-assisted drafting/review.

**What We Agreed**

- Markdown should be the primary editable format for contract templates.
- Word remains a first-class source/target format (draft in Word or MD, convert both ways).
- Use Pandoc + Lua filters as the conversion backbone.
- Prefer concise attribute-style authoring in Markdown for non-standard styles (not verbose `:::` blocks).
- Use a shared CSS layer for MD -> HTML/PDF output and VS Code WYSIWYG-like preview.

**Style Strategy**

- Standard document structure uses native Markdown:
    - `#`..`######` for heading hierarchy
    - `-` for bullet lists
    - `1.` / `a)` / `i.` for ordered lists
    - bold/italic/etc. for inline emphasis
- Non-standard Word styles use Markdown attributes with class aliases:
    - Example: `Comments {.Comments}`
- Important syntax constraint:
    - Pandoc class names cannot contain spaces.
    - So Word styles with spaces need alias mapping (e.g. `.Heading-2` -> `"Heading 2"` in Lua).
- Direct form `{custom-style="Heading 2"}` is still valid, but less human-friendly.

**Canonical Markdown Serialization Rules**

- Bullet list indentation is normalized as:
    - `- <text>` for `List`
    - `  - <text>` for `List 2`
    - `    - <text>` for `List 3`
- `Article 2` must be serialized as explicit `1.1.` style text in Markdown (e.g. `1.1. <text>`), not as a plain `1.` list marker.
- Avoid loose list formatting:
    - no extra blank lines before list items
    - no extra blank lines before `a)`/`i.` subclauses in Articles/Appendices
    - no extra blank lines before `Article 2` lines

**Implemented Components**

- `filters/docx_to_compact.lua`
    - DOCX -> MD cleanup/compaction (removes TOC/anchor noise, simplifies style representation).
- `filters/compact_to_docx.lua`
    - MD -> DOCX mapping from compact class attributes to Word `custom-style`.
    - Supports paragraph suffix form like `Text {.Comments}`.
- `scripts/clean_docx_template.ps1`
    - Optional DOCX cleanup step (prunes unused styles/metadata/custom XML).

**Core Commands**

- DOCX -> compact MD:

```powershell
pandoc -f docx+styles -t markdown --lua-filter=filters/docx_to_compact.lua contract.docx -o contract.compact.md
```

- MD -> DOCX (with style restoration):

```powershell
pandoc -f markdown+fancy_lists+lists_without_preceding_blankline -t docx --no-highlight --reference-doc=contract.slim.docx --lua-filter=filters/compact_to_docx.lua contract.compact.md -o contract.out.docx
```

- Optional template cleanup:

```powershell
powershell -ExecutionPolicy Bypass -File scripts\\clean_docx_template.ps1 -InputDocx contract.docx -OutputDocx contract.slim.docx
```

- MD -> HTML/PDF with CSS:

```powershell
pandoc contract.compact.md -c contract.css -o contract.html
pandoc contract.compact.md -c contract.css -o contract.pdf
```

**Practical Outcome**

- This gives a scalable MD-first contract workflow with round-trip style preservation for legal templates.
- It reduces AI processing overhead versus XML/Word-only editing.
- It centralizes styling and numbering behavior in filters/CSS instead of manual Word editing across hundreds of templates.

**Known Limits / Expectations**

- “Lossless” means style/format intent is preserved; exact Word internal metadata/IDs/TOC field structures may differ after round-trips.
- Any new custom Word style must be added once to the Lua style map (alias -> Word style name).
