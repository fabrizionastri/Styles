# FlexUp Contract Style Guide

This document defines every paragraph style, character format, numbering scheme, and
Markdown mapping used in FlexUp contract templates. It is the single source of truth
for the Pandoc Lua filters that convert between Markdown and Word (docx).

---

## 1. Document Defaults

| Property         | Value                                              |
|------------------|----------------------------------------------------|
| Default font     | Open Sans                                          |
| Default size     | 10 pt                                              |
| Language         | en-GB                                              |
| Page size        | A4 (210 × 297 mm)                                  |
| Left margin      | 1.5 cm                                             |
| Right margin     | 1.5 cm                                             |
| Top margin       | 1.25 cm                                            |
| Bottom margin    | 1.6 cm                                             |

---

## 2. Paragraph Styles – Plain Headings (unnumbered)

These are simple structural headings with no auto-numbering. They map directly
to Markdown heading levels H1–H4.

### 2.1 Heading 1

The top-level document title (e.g. the contract name). Not numbered.

| Property       | Value                                          |
|----------------|------------------------------------------------|
| Font           | Open Sans, 14 pt, bold                         |
| Alignment      | Centred                                        |
| Bottom border  | Single line                                    |
| Spacing before | 0                                              |
| Spacing after  | 12 pt                                          |
| Outline level  | 0                                              |
| Markdown       | `# <text>`                                     |

### 2.2 Heading 2

A major sub-heading within the contract body (unnumbered).

| Property       | Value                                          |
|----------------|------------------------------------------------|
| Font           | Open Sans, 12 pt, bold, underline              |
| Alignment      | Left                                           |
| Spacing before | 24 pt                                          |
| Spacing after  | 6 pt                                           |
| TOC level      | 1                                               |
| Outline level  | 1                                              |
| Markdown       | `## <text>`                                    |

### 2.3 Heading 3

A secondary sub-heading (unnumbered).

| Property       | Value                                          |
|----------------|------------------------------------------------|
| Font           | Open Sans, 11 pt, bold                         |
| Alignment      | Left                                           |
| TOC level      | 2                                               |
| Outline level  | 2                                              |
| Markdown       | `### <text>`                                   |

### 2.4 Heading 4

A minor sub-heading (unnumbered).

| Property       | Value                                          |
|----------------|------------------------------------------------|
| Font           | Open Sans, 10 pt, underline                    |
| Alignment      | Left                                           |
| Markdown       | `#### <text>`                                  |

---

## 3. Paragraph Styles – Section and Articles (numbered, main body)

These styles form the legal hierarchy of a contract. They share a single
multi-level numbering list so that sub-article numbers reset correctly.

### 3.1 Section (custom style)

The main structural division of a contract (e.g. "Section I. Definitions").
Auto-numbered with upper-Roman numerals.

| Property       | Value                                          |
|----------------|------------------------------------------------|
| Font           | Open Sans, 12 pt, bold, underline              |
| Alignment      | Left                                           |
| Spacing before | 24 pt                                          |
| Spacing after  | 6 pt                                           |
| TOC level      | 1                                               |
| Numbering      | "Section I." / "Section II." (upper-Roman)     |
| Follow number  | Space                                          |
| Markdown       | `## Section I. <text>`                         |

Visual rule: Section looks identical to Heading 2 — same font, size, bold,
underline, alignment, and spacing. The only difference is the auto-numbering
prefix "Section I."

### 3.2 Article 1 (custom style)

Top-level numbered article within a Section.

| Property       | Value                                          |
|----------------|------------------------------------------------|
| Font           | Open Sans, 11 pt, bold                         |
| Alignment      | Left                                           |
| Spacing before | 18 pt                                          |
| Spacing after  | 6 pt                                           |
| TOC level      | 2                                               |
| Numbering      | "Article 1." / "Article 2." (decimal)          |
| Follow number  | Space                                          |
| Indent left    | 0                                              |
| Text indent    | 0                                              |
| Markdown       | `### Article 1. <text>`                        |

Visual rule: Article 1 looks identical to Heading 3 — same font, size, bold,
alignment. The only difference is the auto-numbering prefix "Article N."

### 3.3 Article 2 (custom style)

Second-level numbered clause (e.g. "1.1", "1.2").

| Property       | Value                                          |
|----------------|------------------------------------------------|
| Font           | Open Sans, 10 pt, regular                      |
| Alignment      | Justified                                      |
| Spacing before | 6 pt                                           |
| Spacing after  | 6 pt                                           |
| Numbering      | "1.1" / "1.2" (parent.child decimal)           |
| Number style   | 1, 2, 3                                        |
| Follow number  | Tab                                            |
| Restart after  | Article 1                                      |
| Indent left    | 0                                              |
| Text indent    | 0.75 cm                                        |
| Markdown       | `1.1. <text>` (numbered list)                  |

### 3.4 Article 3 (custom style)

Third-level lettered clause (e.g. "a)", "b)").

| Property       | Value                                          |
|----------------|------------------------------------------------|
| Font           | Open Sans, 10 pt, regular                      |
| Alignment      | Justified                                      |
| Spacing before | 3 pt                                           |
| Spacing after  | 3 pt                                           |
| Numbering      | "a)" / "b)" (lower-letter)                    |
| Number style   | a, b, c                                        |
| Follow number  | Tab                                            |
| Restart after  | Article 2                                      |
| Aligned at     | 1.00 cm                                        |
| Text indent    | 1.50 cm                                        |
| Markdown       | `a) <text>` (numbered list)                    |

### 3.5 Article 4 (custom style)

Fourth-level Roman-numeral clause (e.g. "i.", "ii.").

| Property       | Value                                          |
|----------------|------------------------------------------------|
| Font           | Open Sans, 10 pt, regular                      |
| Alignment      | Justified                                      |
| Spacing before | 3 pt                                           |
| Spacing after  | 3 pt                                           |
| Numbering      | "i." / "ii." (lower-Roman)                    |
| Number style   | i, ii, iii                                     |
| Follow number  | Tab                                            |
| Restart after  | Article 3                                      |
| Aligned at     | 1.75 cm                                        |
| Text indent    | 2.25 cm                                        |
| Markdown       | `i. <text>` (numbered list)                    |

---

## 4. Paragraph Styles – Appendices (numbered, separate list)

Appendices use a separate numbering list, independent of the main body.
The numbering hierarchy mirrors Articles but uses different prefixes to
avoid confusion.

### 4.1 Appendix 1 (custom style)

Top-level appendix title. Always starts on a new page.

| Property       | Value                                          |
|----------------|------------------------------------------------|
| Font           | Open Sans, 12 pt, bold                         |
| Alignment      | Centred                                        |
| Bottom border  | Single line                                    |
| Page break     | Before (always starts a new page)              |
| Spacing before | 0                                              |
| Spacing after  | 12 pt                                          |
| TOC level      | 1                                               |
| Numbering      | "Appendix 1." / "Appendix 2." (decimal)       |
| Follow number  | Space                                          |
| Markdown       | `# Appendix 1. <text>`                         |

### 4.2 Appendix 2 (custom style)

Section-level heading within an appendix, numbered with upper-case letters.

| Property       | Value                                          |
|----------------|------------------------------------------------|
| Font           | Open Sans, 12 pt, bold, underline              |
| Alignment      | Left                                           |
| Spacing before | 24 pt                                          |
| Spacing after  | 6 pt                                           |
| TOC level      | 2                                               |
| Numbering      | "Section A." / "Section B." (upper-letter)     |
| Follow number  | Space                                          |
| Markdown       | `## Section A. <text>`                         |

Visual rule: Appendix 2 looks identical to Heading 2 and Section — same font,
size, bold, underline, alignment. Only the numbering prefix differs.

### 4.3 Appendix 3 (custom style)

Numbered sub-section within an appendix.

| Property       | Value                                          |
|----------------|------------------------------------------------|
| Font           | Open Sans, 11 pt, bold                         |
| Alignment      | Left                                           |
| Spacing before | 18 pt                                          |
| Spacing after  | 6 pt                                           |
| Numbering      | "1." / "2." (decimal with period)              |
| Number style   | 1, 2, 3                                        |
| Follow number  | Tab                                            |
| Indent left    | 0                                              |
| Text indent    | 0.75 cm                                        |
| Markdown       | `### 1. <text>` (H3 level)                     |

Visual rule: Appendix 3 looks identical to Article 1 — same font, size, bold,
alignment. The numbering uses plain "1." instead of "Article 1."

### 4.4 Appendix 4 (custom style)

Lettered sub-clauses within an appendix section.

| Property       | Value                                          |
|----------------|------------------------------------------------|
| Font           | Open Sans, 10 pt, regular                      |
| Alignment      | Justified                                      |
| Spacing before | 6 pt                                           |
| Spacing after  | 6 pt                                           |
| Numbering      | "a)" / "b)" (lower-letter)                    |
| Number style   | a, b, c                                        |
| Follow number  | Tab                                            |
| Restart after  | Appendix 3                                     |
| Aligned at     | 1.00 cm                                        |
| Text indent    | 1.50 cm                                        |
| Markdown       | `b) <text>` (numbered list)                    |

### 4.5 Appendix 5 (custom style)

Roman-numeral sub-clauses within an appendix.

| Property       | Value                                          |
|----------------|------------------------------------------------|
| Font           | Open Sans, 10 pt, regular                      |
| Alignment      | Justified                                      |
| Spacing before | 3 pt                                           |
| Spacing after  | 3 pt                                           |
| Numbering      | "i." / "ii." (lower-Roman)                    |
| Number style   | i, ii, iii                                     |
| Follow number  | Tab                                            |
| Restart after  | Appendix 4                                     |
| Aligned at     | 1.75 cm                                        |
| Text indent    | 2.25 cm                                        |
| Markdown       | `i. <text>` (numbered list)                    |

---

## 5. Special Paragraph Styles

### 5.1 Normal

The base style for body text. All other styles inherit from Normal.

| Property       | Value                                          |
|----------------|------------------------------------------------|
| Font           | Open Sans, 10 pt, regular                      |
| Alignment      | Justified                                      |
| Spacing before | 6 pt                                           |
| Spacing after  | 3 pt                                           |
| Line spacing   | ~1.05 × multiple                               |
| Markdown       | Plain paragraph (no attribute)                 |

### 5.2 Comments (custom style)

Internal drafting notes. Green italic text, not printed in final versions.

| Property       | Value                                          |
|----------------|------------------------------------------------|
| Font           | Open Sans, 10 pt, italic                       |
| Colour         | Green (#00B050)                                |
| Alignment      | Left                                           |
| Spacing before | 6 pt                                           |
| Spacing after  | 6 pt                                           |
| Markdown       | `<text> {.Comments}`                           |

### 5.3 Published (custom style)

Publication or version metadata line.

| Property       | Value                                          |
|----------------|------------------------------------------------|
| Font           | Open Sans, 10 pt, italic                       |
| Alignment      | Right                                          |
| Spacing before | 6 pt                                           |
| Spacing after  | 6 pt                                           |
| Markdown       | `<text> {.Published}`                          |

---

## 6. List Styles (bullet lists)

| Style   | Bullet | Indent left | Hanging  | Markdown            |
|---------|--------|-------------|----------|---------------------|
| List    | −      | 0.5 cm      | 0.38 cm  | `- <text>`          |
| List 2  | −      | 1.0 cm      | 0.5 cm   | `  - <text>`        |
| List 3  | −      | 1.5 cm      | 0.5 cm   | `    - <text>`      |

---

## 7. Numbering Schemes

### 7.1 Articles (main body)

| ilvl | Word Style  | Format       | Pattern        | Restarts After | Aligned at | Text indent |
|------|-------------|--------------|----------------|----------------|------------|-------------|
| 0    | Section     | Upper Roman  | "Section I."   | —              | 0          | 0           |
| 1    | Article 1   | Decimal      | "Article 1."   | —              | 0          | 0           |
| 2    | Article 2   | Decimal      | "1.1"          | Article 1      | 0          | 0.75 cm     |
| 3    | Article 3   | Lower letter | "a)"           | Article 2      | 1.00 cm    | 1.50 cm     |
| 4    | Article 4   | Lower Roman  | "i."           | Article 3      | 1.75 cm    | 2.25 cm     |

### 7.2 Annexes (appendices)

| ilvl | Word Style  | Format       | Pattern          | Restarts After | Aligned at | Text indent |
|------|-------------|--------------|------------------|----------------|------------|-------------|
| 0    | Appendix 1  | Decimal      | "Appendix 1."   | —              | 0          | 0           |
| 1    | Appendix 2  | Upper letter | "Section A."     | —              | 0          | 0           |
| 2    | Appendix 3  | Decimal      | "1."             | —              | 0          | 0.75 cm     |
| 3    | Appendix 4  | Lower letter | "a)"             | Appendix 3     | 1.00 cm    | 1.50 cm     |
| 4    | Appendix 5  | Lower Roman  | "i."             | Appendix 4     | 1.75 cm    | 2.25 cm     |

---

## 8. Visual Consistency Rules

Styles at the same Markdown heading level should look identical in both the
Markdown source and the Word output, differing only in their numbering prefix.

| MD Level | Styles that share this level          | Common format                  |
|----------|---------------------------------------|--------------------------------|
| H1       | Heading 1, Appendix 1                | Bold, centred, bottom border   |
| H2       | Heading 2, Section, Appendix 2       | Bold 12 pt, left, underline   |
| H3       | Heading 3, Article 1, Appendix 3     | Bold 11 pt, left              |
| H4       | Heading 4                            | Normal 10 pt, underline, left |

Note: Heading 1 is 14 pt; Appendix 1 is 12 pt. This is intentional — the
contract title is visually larger than appendix titles.

---

## 9. Markdown ↔ Word Mapping Summary

### 9.1 Full mapping table

| Markdown Syntax                | Word Style     | MD Level | Notes                              |
|--------------------------------|----------------|----------|------------------------------------|
| `# <text>`                     | Heading 1      | H1       | Contract title, no numbering       |
| `# Appendix 1. <text>`        | Appendix 1     | H1       | Page break before, numbered        |
| `## <text>`                    | Heading 2      | H2       | Plain sub-heading, unnumbered      |
| `## Section I. <text>`         | Section        | H2       | Auto-numbered, Roman numerals      |
| `## Section A. <text>`         | Appendix 2     | H2       | Auto-numbered, upper letters       |
| `### <text>`                   | Heading 3      | H3       | Plain sub-heading, unnumbered      |
| `### Article 1. <text>`        | Article 1      | H3       | Auto-numbered, decimal             |
| `### 1. <text>`                | Appendix 3     | H3       | Auto-numbered (appendix context)   |
| `#### <text>`                  | Heading 4      | H4       | Underlined sub-heading             |
| `1.1. <text>`                  | Article 2      | list     | Numbered sub-clause                |
| `a) <text>`                    | Article 3      | list     | Lettered sub-clause                |
| `i. <text>`                    | Article 4      | list     | Roman-numeral sub-clause           |
| `- <text>`                     | List           | bullet   | Dash bullet                        |
| `<text> {.Comments}`           | Comments       | para     | Green italic drafting note         |
| `<text> {.Published}`          | Published      | para     | Italic right-aligned metadata      |

### 9.2 Lua Filter Disambiguation Rules

Because several Word styles share the same Markdown heading level, the Lua
filters must use pattern matching on the heading text to decide which Word
style to apply.

**H1 disambiguation:**
If text starts with `Appendix` followed by a number and period → Appendix 1;
otherwise → Heading 1.

**H2 disambiguation:**
If text starts with `Section` followed by a Roman numeral and period → Section;
if text starts with `Section` followed by an upper-case letter and period → Appendix 2;
otherwise → Heading 2 (plain, unnumbered).

**H3 disambiguation:**
If text starts with `Article` followed by a number and period → Article 1;
if text starts with a number followed by a period → Appendix 3 (but only when
the heading appears after an Appendix 1 heading in the document; if it appears
before any Appendix 1, it is ambiguous — see §9.3 below);
otherwise → Heading 3 (plain, unnumbered).

**H4 disambiguation:**
No disambiguation needed — always maps to Heading 4.

**Numbered list disambiguation (for `1.1.`, `a)`, `i.`):**
When a numbered list item appears before the first Appendix 1 heading in the
document, it maps to Article 2/3/4. When it appears after an Appendix 1
heading, it maps to Appendix 3/4/5. The filter must track document position
to determine context.

### 9.3 Context-Based Style Resolution

The Lua filter maintains a state flag that tracks whether the current position
in the document is in the "main body" or in the "appendix" section. The flag
flips to "appendix" when an H1 heading matching the Appendix pattern is
encountered.

| Pattern              | Before first Appendix 1 | After first Appendix 1 |
|----------------------|-------------------------|------------------------|
| `### 1. <text>`      | *(ambiguous — see note)* | Appendix 3             |
| `1.1. <text>`        | Article 2               | *(not used)*           |
| `a) <text>`          | Article 3               | Appendix 4             |
| `i. <text>`          | Article 4               | Appendix 5             |

Note: `### 1. <text>` in the main body is ambiguous because Heading 3 uses
plain `### <text>` and Article 1 uses `### Article 1. <text>`. A bare `### 1.`
in the main body would be unusual. If encountered, the filter should treat it
as a plain Heading 3 with text starting with "1.".

---

## 10. Table Formatting

Tables use the "Table Grid" style with single-line borders on all sides. The
first row typically has light grey shading (#F2F2F2) and bold text. Table cell
text inherits Normal formatting (Open Sans 10 pt, justified). Tables use fixed
column widths.

---

## 11. Footer

| Property       | Value                                          |
|----------------|------------------------------------------------|
| Font           | Open Sans, 8 pt                                |
| Tabs           | Right-aligned at page width                    |
| Content        | `{FILENAME}` [tab] `page {PAGE} on {NUMPAGES}`|

---

## 12. TOC (Table of Contents)

The TOC should include:
- Section headings (TOC level 1)
- Article 1 headings (TOC level 2)
- Appendix 1 headings (TOC level 1)
- Appendix 2 headings (TOC level 2)
- Heading 2 (TOC level 1)
- Heading 3 (TOC level 2)

---

## 13. Inline Formatting

| Markdown            | Word Rendering                          |
|----------------------|-----------------------------------------|
| `**text**`          | Bold                                    |
| `*text*`            | Italic                                  |
| `***text***`        | Bold + Italic                           |
| `[text](url)`       | Hyperlink (blue #0563C1, underlined)    |
| `~~text~~`          | Strikethrough                           |
| `^super^`           | Superscript (Pandoc syntax)             |
| `~sub~`             | Subscript (Pandoc syntax)               |

---

## 14. Character Styles

These are automatically maintained by Word and do not need explicit handling in
the Lua filters:

- Linked character styles for each heading (Heading1Char, etc.)
- Hyperlink: Blue (#0563C1), underline
- FootnoteReference / EndnoteReference: Superscript
- FootnoteText: 8 pt with tight spacing

---

*This style guide is based on the style_matrix.md as updated on 2026-02-17,
with all inconsistencies resolved per Fabrizio's decisions.*
